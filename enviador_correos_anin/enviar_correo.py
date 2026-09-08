"""
IN Piura - Enviador de correos institucionales ANIN
Envio programado de correos desde la cuenta institucional (@dime.gob.pe),
con plantilla HTML de identidad ANIN - DIME - SESDI.

Disenado para correr en Windows bajo el Programador de tareas, sin
dependencias externas: solo biblioteca estandar de Python 3.8+.

La clave NUNCA se escribe en este archivo ni en config.ini. Se lee de una
variable de entorno o de un archivo cifrado con DPAPI (ver LEEME.md).

Uso rapido:
    python enviar_correo.py --detectar dime.gob.pe
    python enviar_correo.py --probar-conexion
    python enviar_correo.py correos/reporte_semanal.ini --prueba
    python enviar_correo.py correos/reporte_semanal.ini

Autor: Ing. Hector Salomon Cahuas Miller - ANIN / DIME / SESDI
"""

import argparse
import configparser
import html as _html
import mimetypes
import os
import re
import smtplib
import ssl
import subprocess
import sys
import unicodedata
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from pathlib import Path

# --- Identidad institucional ANIN -----------------------------------------
COLOR_VERDE_ANIN = "#1B4D2E"
COLOR_AZUL_ANIN = "#1B4F72"
COLOR_DORADO = "#B8860B"

BASE = Path(__file__).resolve().parent
RUTA_CONFIG_POR_DEFECTO = BASE / "config.ini"
RUTA_PLANTILLA_POR_DEFECTO = BASE / "plantilla_anin.html"
DIR_REGISTRO = BASE / "registro"
DIR_MARCADORES = BASE / "marcadores"
DIR_PRUEBAS = BASE / "pruebas"

LIMITE_ADJUNTOS_BYTES = 25 * 1024 * 1024  # Gmail y Microsoft 365: 25 MB

# --- Perfiles SMTP --------------------------------------------------------
# El correo institucional puede estar en Google Workspace o en Microsoft 365.
# El perfil se declara en config.ini; 'manual' permite fijar host y puerto.
PERFILES_SMTP = {
    "google": {
        "host": "smtp.gmail.com",
        "puerto": 587,
        "cifrado": "starttls",
        "ayuda": (
            "Google Workspace exige una CONTRASENA DE APLICACION de 16 caracteres\n"
            "  (no la contrasena normal). Requiere verificacion en 2 pasos activa.\n"
            "  Generela en: https://myaccount.google.com/apppasswords"
        ),
    },
    "microsoft": {
        "host": "smtp.office365.com",
        "puerto": 587,
        "cifrado": "starttls",
        "ayuda": (
            "Microsoft 365 desactiva por defecto la autenticacion SMTP basica.\n"
            "  El area de TI debe habilitar 'Authenticated SMTP' para este buzon\n"
            "  (Centro de administracion > Usuarios > Correo > Aplicaciones de correo)."
        ),
    },
}

MESES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "setiembre", "octubre", "noviembre", "diciembre",
]
DIAS = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]


# ==========================================================================
# Utilidades
# ==========================================================================

def morir(mensaje, ayuda=None):
    """Termina el programa con un mensaje de error legible en espanol."""
    print("\n[ERROR] " + mensaje, file=sys.stderr)
    if ayuda:
        print("  " + ayuda.replace("\n", "\n  "), file=sys.stderr)
    print("", file=sys.stderr)
    sys.exit(1)


def avisar(mensaje):
    print("[AVISO] " + mensaje)


def informar(mensaje):
    print("[OK] " + mensaje)


def leer_ini(ruta, obligatorio=True):
    """Lee un .ini en UTF-8 sin interpolacion.

    La interpolacion se desactiva a proposito: los cuerpos de correo
    contienen '%' (porcentajes de avance, cobertura) y '{{VAR}}', que
    configparser interpretaria como sintaxis propia y romperia la lectura.
    """
    ruta = Path(ruta)
    if not ruta.exists():
        if obligatorio:
            morir(
                "No se encuentra el archivo: %s" % ruta,
                "Copie el archivo de ejemplo correspondiente y editelo.",
            )
        return None
    # comment_prefixes=(';',) es deliberado: por defecto configparser borra
    # toda linea que empiece con '#', lo que se comeria los subtitulos '## '
    # escritos dentro del cuerpo del correo. Los comentarios del .ini van con ';'.
    cfg = configparser.ConfigParser(
        interpolation=None,
        comment_prefixes=(";",),
        inline_comment_prefixes=None,
        empty_lines_in_values=True,
    )
    cfg.optionxform = str  # conserva mayusculas en los nombres de variables
    try:
        with open(ruta, "r", encoding="utf-8-sig") as fh:
            cfg.read_file(fh)
    except UnicodeDecodeError:
        morir(
            "El archivo %s no esta guardado en UTF-8." % ruta.name,
            "Abralo en el Bloc de notas y use Archivo > Guardar como > "
            "Codificacion: UTF-8.",
        )
    except configparser.Error as exc:
        morir("El archivo %s tiene un error de formato:\n%s" % (ruta.name, exc))
    return cfg


def sin_acentos(texto):
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


def lista_correos(valor, campo):
    """Convierte 'a@x.pe, b@y.pe' o una lista multilinea en lista validada."""
    if not valor:
        return []
    crudos = re.split(r"[,;\n]+", valor)
    salida = []
    for item in crudos:
        item = item.strip()
        if not item:
            continue
        if not re.match(r"^[^@\s<>]+@[^@\s<>]+\.[A-Za-z]{2,}$", item):
            morir(
                "Direccion de correo no valida en '%s': %r" % (campo, item),
                "Escriba solo la direccion, sin nombre ni <>. "
                "Separe varias con comas.",
            )
        salida.append(item)
    return salida


# ==========================================================================
# Variables de plantilla  {{FECHA}}, {{MES}}, ...
# ==========================================================================

def variables_de_fecha(ahora=None):
    """Variables automaticas de fecha, en hora local del equipo (Peru UTC-5)."""
    ahora = ahora or datetime.now()
    iso = ahora.isocalendar()
    return {
        "FECHA": ahora.strftime("%d/%m/%Y"),
        "FECHA_ISO": ahora.strftime("%Y-%m-%d"),
        "FECHA_LARGA": "%d de %s de %d" % (ahora.day, MESES[ahora.month - 1], ahora.year),
        "HORA": ahora.strftime("%H:%M"),
        "DIA": "%02d" % ahora.day,
        "DIA_SEMANA": DIAS[ahora.weekday()],
        "MES": MESES[ahora.month - 1],
        "MES_NUM": "%02d" % ahora.month,
        "ANIO": str(ahora.year),
        "SEMANA": str(iso[1]),
        "TRIMESTRE": str((ahora.month - 1) // 3 + 1),
    }


def resolver_variables(texto, variables, donde=""):
    """Reemplaza {{VAR}} por su valor. Avisa si queda alguna sin resolver."""
    if not texto:
        return texto

    def _sub(m):
        nombre = m.group(1).strip().upper()
        if nombre in variables:
            return str(variables[nombre])
        avisar(
            "La variable {{%s}} no tiene valor%s; se deja vacia. "
            "Definala en la seccion [variables]." % (nombre, (" en " + donde) if donde else "")
        )
        return ""

    return re.sub(r"\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\}\}", _sub, texto)


# ==========================================================================
# Conversion del cuerpo a HTML y a texto plano
# ==========================================================================

def _formato_en_linea(texto):
    """Aplica **negrita** y *cursiva* sobre texto ya escapado para HTML."""
    texto = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", texto)
    texto = re.sub(r"(?<![\*\w])\*(?!\s)(.+?)(?<!\s)\*(?![\*\w])", r"<em>\1</em>", texto)
    return texto


def _es_subtitulo(linea):
    return linea.startswith("## ") or linea.startswith(">> ")


def _marcador_item(linea):
    """Devuelve ('ul'|'ol', contenido) si la linea inicia un item de lista."""
    m = re.match(r"^[-*]\s+(.*)$", linea)
    if m:
        return "ul", m.group(1)
    m = re.match(r"^\d+[.)]\s+(.*)$", linea)
    if m:
        return "ol", m.group(1)
    return None, None


def _recorrer_cuerpo(cuerpo):
    """Recorre el cuerpo y devuelve una lista de bloques normalizados.

    Cada bloque es una tupla:
        ("parrafo", texto)
        ("subtitulo", texto)
        ("lista", "ul"|"ol", [texto_item, ...])

    Las lineas que continuan un item de lista se unen a ese item. Esto es
    necesario porque al redactar en el archivo .ini un item largo se parte en
    varias lineas, y configparser entrega cada linea por separado: si no se
    unieran, la continuacion saldria como un parrafo suelto fuera de la lista.
    """
    bloques = []
    parrafo = []
    lista_tipo = None
    items = []
    item = []

    def cerrar_parrafo():
        if parrafo:
            bloques.append(("parrafo", " ".join(parrafo)))
            parrafo.clear()

    def cerrar_item():
        if item:
            items.append(" ".join(item))
            item.clear()

    def cerrar_lista():
        nonlocal lista_tipo
        cerrar_item()
        if items:
            bloques.append(("lista", lista_tipo, list(items)))
            items.clear()
        lista_tipo = None

    for linea_cruda in cuerpo.splitlines():
        linea = linea_cruda.strip()

        if not linea:
            cerrar_parrafo()
            cerrar_lista()
            continue

        if _es_subtitulo(linea):
            cerrar_parrafo()
            cerrar_lista()
            bloques.append(("subtitulo", linea[3:].strip()))
            continue

        tipo, contenido = _marcador_item(linea)
        if tipo:
            cerrar_parrafo()
            if lista_tipo and lista_tipo != tipo:
                cerrar_lista()
            cerrar_item()
            lista_tipo = tipo
            item.append(contenido)
            continue

        # Linea sin marcador: continua el item de lista abierto, si lo hay;
        # en caso contrario, continua el parrafo.
        if lista_tipo:
            item.append(linea)
        else:
            parrafo.append(linea)

    cerrar_parrafo()
    cerrar_lista()
    return bloques


def cuerpo_a_html(cuerpo):
    """Convierte el cuerpo (texto sencillo) en HTML de correo.

    Sintaxis aceptada, pensada para redactar sin saber HTML:
      ## Subtitulo  o  >> Subtitulo   -> subtitulo en verde ANIN
      - item  /  * item     -> lista con vinetas
      1. item               -> lista numerada
      **negrita**  *cursiva*
      linea en blanco       -> separacion de parrafos

    Un item o parrafo puede ocupar varias lineas: se unen automaticamente.
    """
    estilo_p = ("margin:0 0 12px 0;font-family:Arial,Helvetica,sans-serif;"
                "font-size:14px;line-height:1.6;color:#333333;")
    estilo_h = ("margin:22px 0 10px 0;font-family:Arial,Helvetica,sans-serif;"
                "font-size:15px;font-weight:bold;color:%s;"
                "border-bottom:1px solid #D5DBDB;padding-bottom:4px;" % COLOR_VERDE_ANIN)
    estilo_li = ("margin:0 0 6px 0;font-family:Arial,Helvetica,sans-serif;"
                 "font-size:14px;line-height:1.6;color:#333333;")
    estilo_lista = "margin:0 0 14px 0;padding-left:22px;"

    salida = []
    for bloque in _recorrer_cuerpo(cuerpo):
        if bloque[0] == "parrafo":
            texto = _formato_en_linea(_html.escape(bloque[1]))
            salida.append('<p style="%s">%s</p>' % (estilo_p, texto))
        elif bloque[0] == "subtitulo":
            texto = _formato_en_linea(_html.escape(bloque[1]))
            salida.append('<h2 style="%s">%s</h2>' % (estilo_h, texto))
        else:
            _, tipo, items = bloque
            salida.append('<%s style="%s">' % (tipo, estilo_lista))
            for texto_item in items:
                texto = _formato_en_linea(_html.escape(texto_item))
                salida.append('<li style="%s">%s</li>' % (estilo_li, texto))
            salida.append("</%s>" % tipo)
    return "\n".join(salida)


def _quitar_formato(texto):
    texto = re.sub(r"\*\*(.+?)\*\*", r"\1", texto)
    texto = re.sub(r"(?<![\*\w])\*(?!\s)(.+?)(?<!\s)\*(?![\*\w])", r"\1", texto)
    return texto


def cuerpo_a_texto(cuerpo):
    """Version en texto plano (alternativa para clientes sin HTML)."""
    lineas = []
    for bloque in _recorrer_cuerpo(cuerpo):
        if bloque[0] == "parrafo":
            lineas.extend([_quitar_formato(bloque[1]), ""])
        elif bloque[0] == "subtitulo":
            titulo = _quitar_formato(bloque[1])
            lineas.extend([titulo.upper(), "-" * len(titulo), ""])
        else:
            _, tipo, items = bloque
            for indice, texto_item in enumerate(items, start=1):
                vineta = "  - " if tipo == "ul" else "  %d. " % indice
                lineas.append(vineta + _quitar_formato(texto_item))
            lineas.append("")
    return "\n".join(lineas).strip() + "\n"


# ==========================================================================
# Armado del mensaje
# ==========================================================================

def cargar_plantilla(ruta):
    ruta = Path(ruta)
    if not ruta.exists():
        morir("No se encuentra la plantilla HTML: %s" % ruta)
    return ruta.read_text(encoding="utf-8")


def construir_html(plantilla, campos):
    """Rellena la plantilla ANIN. Usa reemplazo literal de {{CAMPO}}.

    No se usa str.format() ni % porque el CSS de la plantilla lleva llaves
    y porcentajes que romperian el formateo.
    """
    salida = plantilla
    for clave, valor in campos.items():
        salida = salida.replace("{{%s}}" % clave, valor if valor is not None else "")
    # Limpia cualquier marcador opcional que no se haya rellenado.
    salida = re.sub(r"\{\{[A-Za-z_][A-Za-z0-9_]*\}\}", "", salida)
    return salida


def adjuntar_archivos(msg, rutas, base_correo):
    """Adjunta archivos validando existencia y limite de 25 MB en conjunto."""
    total = 0
    adjuntados = []
    for cruda in rutas:
        ruta = Path(os.path.expandvars(os.path.expanduser(cruda.strip())))
        if not ruta.is_absolute():
            ruta = (base_correo / ruta).resolve()
        if not ruta.exists():
            morir(
                "No se encuentra el adjunto: %s" % ruta,
                "Revise la ruta en la seccion [adjuntos] del archivo del correo.\n"
                "Puede usar rutas absolutas (C:\\Users\\...) o relativas al .ini.",
            )
        if not ruta.is_file():
            morir("El adjunto no es un archivo: %s" % ruta)

        datos = ruta.read_bytes()
        total += len(datos)
        if total > LIMITE_ADJUNTOS_BYTES:
            morir(
                "Los adjuntos superan el limite de 25 MB (van %.1f MB)."
                % (total / 1024 / 1024),
                "Suba los archivos pesados a Google Drive o SharePoint y pegue\n"
                "el enlace en el cuerpo del correo.",
            )

        tipo, _ = mimetypes.guess_type(ruta.name)
        mayor, menor = (tipo or "application/octet-stream").split("/", 1)
        msg.add_attachment(datos, maintype=mayor, subtype=menor, filename=ruta.name)
        adjuntados.append((ruta.name, len(datos)))
    return adjuntados


def construir_mensaje(datos_correo, cfg_smtp, plantilla, variables, base_correo):
    """Devuelve (EmailMessage, resumen dict) listo para enviar."""
    seccion = datos_correo["mensaje"]

    asunto = resolver_variables(seccion.get("asunto", "").strip(), variables, "asunto")
    if not asunto:
        morir("Falta 'asunto' en la seccion [mensaje] del archivo del correo.")

    cuerpo = resolver_variables(seccion.get("cuerpo", "").strip(), variables, "cuerpo")
    if not cuerpo:
        morir("Falta 'cuerpo' en la seccion [mensaje] del archivo del correo.")

    para = lista_correos(seccion.get("para", ""), "para")
    copia = lista_correos(seccion.get("copia", ""), "copia")
    copia_oculta = lista_correos(seccion.get("copia_oculta", ""), "copia_oculta")
    if not (para or copia or copia_oculta):
        morir("El correo no tiene destinatarios. Complete 'para' en [mensaje].")

    msg = EmailMessage()
    msg["Subject"] = asunto
    msg["From"] = formataddr((cfg_smtp["nombre_remitente"], cfg_smtp["remitente"]))
    if para:
        msg["To"] = ", ".join(para)
    if copia:
        msg["Cc"] = ", ".join(copia)
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=cfg_smtp["remitente"].split("@")[-1])
    if cfg_smtp.get("responder_a"):
        msg["Reply-To"] = cfg_smtp["responder_a"]

    cuerpo_html = cuerpo_a_html(cuerpo)
    campos = {
        "ASUNTO": _html.escape(asunto),
        "CUERPO": cuerpo_html,
        "COLOR_PRIMARIO": COLOR_VERDE_ANIN,
        "COLOR_SECUNDARIO": COLOR_AZUL_ANIN,
        "COLOR_ACENTO": COLOR_DORADO,
        "FIRMA_NOMBRE": _html.escape(cfg_smtp.get("firma_nombre", "")),
        "FIRMA_CARGO": _html.escape(cfg_smtp.get("firma_cargo", "")),
        "FIRMA_AREA": _html.escape(cfg_smtp.get("firma_area", "")),
        "FIRMA_CORREO": _html.escape(cfg_smtp["remitente"]),
        "FIRMA_TELEFONO": _html.escape(cfg_smtp.get("firma_telefono", "")),
        "PROYECTO": _html.escape(cfg_smtp.get("proyecto", "")),
        "FECHA_LARGA": variables["FECHA_LARGA"],
        "ANIO": variables["ANIO"],
    }
    html_final = construir_html(plantilla, campos)

    firma_txt = "\n".join(
        x for x in [
            "", "--",
            cfg_smtp.get("firma_nombre", ""),
            cfg_smtp.get("firma_cargo", ""),
            cfg_smtp.get("firma_area", ""),
            cfg_smtp["remitente"],
            cfg_smtp.get("firma_telefono", ""),
            cfg_smtp.get("proyecto", ""),
        ] if x
    )
    msg.set_content(cuerpo_a_texto(cuerpo) + firma_txt + "\n")
    msg.add_alternative(html_final, subtype="html")

    adjuntos = []
    if datos_correo.get("adjuntos"):
        rutas = [v for _, v in datos_correo["adjuntos"].items() if v.strip()]
        if rutas:
            adjuntos = adjuntar_archivos(msg, rutas, base_correo)

    resumen = {
        "asunto": asunto,
        "para": para,
        "copia": copia,
        "copia_oculta": copia_oculta,
        "adjuntos": adjuntos,
        "html": html_final,
    }
    return msg, resumen


# ==========================================================================
# Credenciales y envio
# ==========================================================================

def obtener_clave(cfg_smtp):
    """Obtiene la clave SMTP sin escribirla nunca en disco en texto plano.

    Orden de busqueda:
      1. Archivo cifrado con DPAPI (clave_archivo)  - recomendado
      2. Variable de entorno (clave_variable_entorno)
    """
    ruta_cifrada = cfg_smtp.get("clave_archivo", "").strip()
    if ruta_cifrada:
        ruta = Path(os.path.expandvars(os.path.expanduser(ruta_cifrada)))
        if not ruta.is_absolute():
            ruta = (BASE / ruta).resolve()
        if not ruta.exists():
            morir(
                "No se encuentra el archivo de clave cifrada: %s" % ruta,
                "Generelo ejecutando en PowerShell:  .\\guardar_clave.ps1",
            )
        if os.name != "nt":
            morir("El archivo de clave cifrada (DPAPI) solo funciona en Windows.")
        comando = (
            "$ErrorActionPreference='Stop';"
            "$s = Get-Content -Raw -LiteralPath '%s' | ConvertTo-SecureString;"
            "$b = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($s);"
            "[Runtime.InteropServices.Marshal]::PtrToStringAuto($b)" % ruta
        )
        try:
            res = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", comando],
                capture_output=True, text=True, timeout=30,
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            morir("No se pudo descifrar la clave con PowerShell: %s" % exc)
        if res.returncode != 0 or not res.stdout.strip():
            morir(
                "No se pudo descifrar el archivo de clave.",
                "La proteccion DPAPI ata el archivo al usuario y equipo que lo creo.\n"
                "Si cambio de usuario o de PC, vuelva a ejecutar guardar_clave.ps1.\n"
                "Detalle: " + (res.stderr or "").strip(),
            )
        return res.stdout.strip()

    variable = cfg_smtp.get("clave_variable_entorno", "ANIN_SMTP_PASSWORD").strip()
    clave = os.environ.get(variable, "")
    if not clave:
        morir(
            "No hay clave disponible: la variable de entorno %s esta vacia." % variable,
            "Opcion A (recomendada): ejecute  .\\guardar_clave.ps1  y ponga en\n"
            "  config.ini la linea  clave_archivo = clave_smtp.txt\n"
            "Opcion B: abra CMD y ejecute   setx %s \"su-clave\"\n"
            "  Cierre y vuelva a abrir la consola para que tome efecto." % variable,
        )
    return clave


def enviar(msg, cfg_smtp, destinatarios):
    """Envia el mensaje por SMTP con TLS. Traduce los errores comunes."""
    clave = obtener_clave(cfg_smtp)
    contexto = ssl.create_default_context()
    host = cfg_smtp["host"]
    puerto = int(cfg_smtp["puerto"])
    ayuda_perfil = PERFILES_SMTP.get(cfg_smtp["perfil"], {}).get("ayuda", "")

    try:
        if cfg_smtp["cifrado"] == "ssl":
            servidor = smtplib.SMTP_SSL(host, puerto, context=contexto, timeout=60)
        else:
            servidor = smtplib.SMTP(host, puerto, timeout=60)
            servidor.ehlo()
            servidor.starttls(context=contexto)
            servidor.ehlo()
        with servidor:
            servidor.login(cfg_smtp["usuario"], clave)
            servidor.send_message(msg, to_addrs=destinatarios)
    except smtplib.SMTPAuthenticationError as exc:
        detalle = str(exc)
        extra = ayuda_perfil
        if "5.7.139" in detalle or "basic authentication is disabled" in detalle.lower():
            extra = PERFILES_SMTP["microsoft"]["ayuda"]
        elif "5.7.8" in detalle or "not accepted" in detalle.lower():
            extra = PERFILES_SMTP["google"]["ayuda"]
        morir("El servidor rechazo el usuario o la clave.\n%s" % detalle, extra)
    except smtplib.SMTPRecipientsRefused as exc:
        morir("El servidor rechazo destinatarios: %s" % exc.recipients)
    except smtplib.SMTPSenderRefused as exc:
        morir(
            "El servidor rechazo el remitente %s." % cfg_smtp["remitente"],
            "El campo 'remitente' debe coincidir con la cuenta 'usuario', o ser\n"
            "un alias autorizado para ella. Detalle: %s" % exc,
        )
    except (smtplib.SMTPException, OSError) as exc:
        morir(
            "Fallo la conexion con %s:%s\n%s" % (host, puerto, exc),
            "Verifique la conexion a internet y que el cortafuegos o la red de\n"
            "la entidad no bloquee el puerto %s de salida." % puerto,
        )


def registrar(linea):
    """Anexa una linea al registro mensual de envios."""
    DIR_REGISTRO.mkdir(exist_ok=True)
    archivo = DIR_REGISTRO / ("envios_%s.log" % datetime.now().strftime("%Y-%m"))
    with open(archivo, "a", encoding="utf-8") as fh:
        fh.write("%s\t%s\n" % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), linea))


# ==========================================================================
# Configuracion
# ==========================================================================

def leer_configuracion(ruta):
    cfg = leer_ini(ruta)
    if not cfg.has_section("smtp"):
        morir("Falta la seccion [smtp] en %s" % ruta)
    s = cfg["smtp"]
    r = cfg["remitente"] if cfg.has_section("remitente") else {}

    perfil = s.get("perfil", "google").strip().lower()
    if perfil not in PERFILES_SMTP and perfil != "manual":
        morir(
            "Perfil SMTP desconocido: %r" % perfil,
            "Valores validos: google, microsoft, manual",
        )

    if perfil == "manual":
        host = s.get("host", "").strip()
        puerto = s.get("puerto", "").strip()
        if not host or not puerto:
            morir("Con perfil = manual debe indicar 'host' y 'puerto' en [smtp].")
        cifrado = s.get("cifrado", "starttls").strip().lower()
    else:
        base = PERFILES_SMTP[perfil]
        host = s.get("host", "").strip() or base["host"]
        puerto = s.get("puerto", "").strip() or str(base["puerto"])
        cifrado = s.get("cifrado", "").strip().lower() or base["cifrado"]

    if cifrado not in ("starttls", "ssl"):
        morir("'cifrado' debe ser starttls o ssl (esta: %r)" % cifrado)

    usuario = s.get("usuario", "").strip()
    if not usuario:
        morir("Falta 'usuario' en [smtp]: la cuenta institucional completa.")
    remitente = s.get("remitente", "").strip() or usuario
    for etiqueta, valor in (("usuario", usuario), ("remitente", remitente)):
        if not re.match(r"^[^@\s<>]+@[^@\s<>]+\.[A-Za-z]{2,}$", valor):
            morir("'%s' no es una direccion valida: %r" % (etiqueta, valor))

    return {
        "perfil": perfil,
        "host": host,
        "puerto": puerto,
        "cifrado": cifrado,
        "usuario": usuario,
        "remitente": remitente,
        "responder_a": s.get("responder_a", "").strip(),
        "clave_variable_entorno": s.get("clave_variable_entorno", "ANIN_SMTP_PASSWORD"),
        "clave_archivo": s.get("clave_archivo", ""),
        "nombre_remitente": r.get("nombre_visible", "").strip() or remitente,
        "firma_nombre": r.get("nombre", "").strip(),
        "firma_cargo": r.get("cargo", "").strip(),
        "firma_area": r.get("area", "").strip(),
        "firma_telefono": r.get("telefono", "").strip(),
        "proyecto": r.get("proyecto", "").strip(),
    }


def detectar_plataforma(dominio):
    """Consulta los registros MX del dominio para saber Google o Microsoft."""
    print("\nConsultando registros MX de %s ...\n" % dominio)
    salida = ""
    for comando in (["nslookup", "-type=mx", dominio],
                    ["dig", "+short", "MX", dominio]):
        try:
            res = subprocess.run(comando, capture_output=True, text=True, timeout=30)
            if res.returncode == 0 and res.stdout.strip():
                salida = res.stdout
                break
        except (OSError, subprocess.TimeoutExpired):
            continue

    if not salida.strip():
        print("No se pudo consultar el DNS desde este equipo.\n")
        print("Metodo alternativo, igual de fiable:")
        print("  Abra su correo institucional en el navegador y mire la direccion:")
        print("    mail.google.com    -> perfil = google")
        print("    outlook.office.com -> perfil = microsoft\n")
        return 2

    print(salida.strip() + "\n")
    minus = salida.lower()
    if "google" in minus or "googlemail" in minus:
        print("=> Plataforma detectada: GOOGLE WORKSPACE")
        print("   En config.ini escriba:  perfil = google")
        print("   " + PERFILES_SMTP["google"]["ayuda"].replace("\n", "\n   "))
        return 0
    if "outlook.com" in minus or "protection.outlook" in minus or "microsoft" in minus:
        print("=> Plataforma detectada: MICROSOFT 365")
        print("   En config.ini escriba:  perfil = microsoft")
        print("   " + PERFILES_SMTP["microsoft"]["ayuda"].replace("\n", "\n   "))
        return 0
    print("=> No es Google ni Microsoft: servidor de correo propio de la entidad.")
    print("   Pida a TI el host y puerto SMTP y use  perfil = manual  en config.ini")
    return 0


def probar_conexion(cfg_smtp):
    """Verifica credenciales sin enviar ningun correo."""
    clave = obtener_clave(cfg_smtp)
    contexto = ssl.create_default_context()
    print("\nProbando %s:%s como %s ..."
          % (cfg_smtp["host"], cfg_smtp["puerto"], cfg_smtp["usuario"]))
    try:
        if cfg_smtp["cifrado"] == "ssl":
            servidor = smtplib.SMTP_SSL(cfg_smtp["host"], int(cfg_smtp["puerto"]),
                                        context=contexto, timeout=60)
        else:
            servidor = smtplib.SMTP(cfg_smtp["host"], int(cfg_smtp["puerto"]), timeout=60)
            servidor.ehlo()
            servidor.starttls(context=contexto)
            servidor.ehlo()
        with servidor:
            servidor.login(cfg_smtp["usuario"], clave)
    except smtplib.SMTPAuthenticationError as exc:
        morir("Usuario o clave rechazados.\n%s" % exc,
              PERFILES_SMTP.get(cfg_smtp["perfil"], {}).get("ayuda", ""))
    except (smtplib.SMTPException, OSError) as exc:
        morir("No se pudo conectar: %s" % exc)
    informar("Conexion y credenciales correctas. Ya puede programar envios.")
    return 0


# ==========================================================================
# Programa principal
# ==========================================================================

def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Enviador de correos institucionales ANIN - Proyecto IN Piura",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Ejemplos:\n"
            "  python enviar_correo.py --detectar dime.gob.pe\n"
            "  python enviar_correo.py --probar-conexion\n"
            "  python enviar_correo.py correos/reporte_semanal.ini --prueba\n"
            "  python enviar_correo.py correos/reporte_semanal.ini\n"
            "  python enviar_correo.py correos/aviso.ini --variable BLOQUE=M5-B1\n"
        ),
    )
    ap.add_argument("correo", nargs="?",
                    help="Archivo .ini con el correo a enviar (carpeta correos/)")
    ap.add_argument("--config", default=str(RUTA_CONFIG_POR_DEFECTO),
                    help="Ruta de config.ini (por defecto: junto al script)")
    ap.add_argument("--plantilla", default=str(RUTA_PLANTILLA_POR_DEFECTO),
                    help="Plantilla HTML a usar")
    ap.add_argument("--prueba", action="store_true",
                    help="No envia: genera una vista previa HTML y la abre")
    ap.add_argument("--probar-conexion", action="store_true",
                    help="Verifica usuario y clave contra el servidor, sin enviar")
    ap.add_argument("--detectar", metavar="DOMINIO",
                    help="Detecta si el dominio usa Google o Microsoft")
    ap.add_argument("--variable", action="append", default=[], metavar="NOMBRE=VALOR",
                    help="Define o sobrescribe una variable de plantilla")
    ap.add_argument("--no-repetir", action="store_true",
                    help="Evita reenviar el mismo correo el mismo dia (util en "
                         "tareas programadas que pueden dispararse dos veces)")
    args = ap.parse_args(argv)

    if args.detectar:
        return detectar_plataforma(args.detectar)

    cfg_smtp = leer_configuracion(args.config)

    if args.probar_conexion:
        return probar_conexion(cfg_smtp)

    if not args.correo:
        ap.print_help()
        print("\n[ERROR] Indique el archivo .ini del correo a enviar.\n", file=sys.stderr)
        return 1

    ruta_correo = Path(args.correo)
    if not ruta_correo.is_absolute():
        candidata = (BASE / ruta_correo)
        ruta_correo = candidata if candidata.exists() else ruta_correo.resolve()
    datos_correo = leer_ini(ruta_correo)
    if not datos_correo.has_section("mensaje"):
        morir("Falta la seccion [mensaje] en %s" % ruta_correo.name)

    # Variables: automaticas de fecha + [variables] del .ini + las de la linea
    # de comandos (estas ultimas tienen prioridad).
    variables = variables_de_fecha()
    if datos_correo.has_section("variables"):
        for clave, valor in datos_correo["variables"].items():
            variables[clave.strip().upper()] = valor.strip()
    for par in args.variable:
        if "=" not in par:
            morir("--variable espera el formato NOMBRE=VALOR (recibido: %r)" % par)
        clave, valor = par.split("=", 1)
        variables[clave.strip().upper()] = valor.strip()

    # Marcador anti-duplicado del dia, por si la tarea programada se dispara
    # dos veces (equipo suspendido, reinicio, "ejecutar si se perdio el inicio").
    marcador = None
    if args.no_repetir:
        DIR_MARCADORES.mkdir(exist_ok=True)
        etiqueta = re.sub(r"[^A-Za-z0-9_.-]", "_", ruta_correo.stem)
        marcador = DIR_MARCADORES / ("%s_%s.enviado" % (etiqueta, variables["FECHA_ISO"]))
        if marcador.exists():
            informar("Ya se envio hoy (%s). No se reenvia." % variables["FECHA_ISO"])
            return 0

    seccion_correo = {"mensaje": datos_correo["mensaje"]}
    if datos_correo.has_section("adjuntos"):
        seccion_correo["adjuntos"] = datos_correo["adjuntos"]

    plantilla = cargar_plantilla(args.plantilla)
    msg, resumen = construir_mensaje(
        seccion_correo, cfg_smtp, plantilla, variables, ruta_correo.parent
    )

    print("")
    print("  De        : %s <%s>" % (cfg_smtp["nombre_remitente"], cfg_smtp["remitente"]))
    print("  Para      : %s" % (", ".join(resumen["para"]) or "-"))
    if resumen["copia"]:
        print("  Copia     : %s" % ", ".join(resumen["copia"]))
    if resumen["copia_oculta"]:
        print("  C. oculta : %s" % ", ".join(resumen["copia_oculta"]))
    print("  Asunto    : %s" % resumen["asunto"])
    for nombre, tam in resumen["adjuntos"]:
        print("  Adjunto   : %s (%.1f KB)" % (nombre, tam / 1024))
    print("")

    if args.prueba:
        DIR_PRUEBAS.mkdir(exist_ok=True)
        salida = DIR_PRUEBAS / ("vista_previa_%s.html"
                                % datetime.now().strftime("%Y%m%d_%H%M%S"))
        salida.write_text(resumen["html"], encoding="utf-8")
        informar("MODO PRUEBA: no se envio nada.")
        print("  Vista previa: %s" % salida)
        if os.name == "nt":
            try:
                os.startfile(str(salida))  # noqa: S606  (solo Windows)
            except OSError:
                pass
        return 0

    destinatarios = resumen["para"] + resumen["copia"] + resumen["copia_oculta"]
    enviar(msg, cfg_smtp, destinatarios)

    informar("Correo enviado a %d destinatario(s)." % len(destinatarios))
    registrar("ENVIADO\t%s\t%s\t%s"
              % (ruta_correo.name, resumen["asunto"], ";".join(destinatarios)))
    if marcador is not None:
        marcador.write_text(datetime.now().isoformat(), encoding="utf-8")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nCancelado por el usuario.")
        sys.exit(130)
