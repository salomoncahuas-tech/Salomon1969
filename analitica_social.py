"""
Analitica del Diagnostico Social (F-DS-01..F-DS-07) - Proyecto IN Piura.

Genera, a partir de las fichas sociales registradas en el aplicativo, los
mismos tres productos que ya existen para el Diagnostico Territorial:

  1. Series listas para graficos interactivos (Altair) en el aplicativo.
  2. Un libro Excel con hojas de datos y graficos nativos de Excel,
     editables por el usuario, con encabezado institucional ANIN.
  3. Un anexo grafico en PDF que acompana a la ficha PDF del bloque.

El modulo NO depende de Streamlit: recibe los registros tal como los
devuelve `database.obtener_diagnosticos_sociales_por_bloque()` y devuelve
estructuras de datos puras, de modo que puede probarse sin levantar la
aplicacion.

Fuente de los datos: columna `dsNN_data_v3` de cada registro (JSON del
formulario V4) mas las columnas de cabecera (bloque, centro poblado,
distrito, responsable, fecha).

ANIN - DIME - SESDI | CUI 2669244 | UTM WGS 84 Zona 17S (EPSG:32717).
"""

import io
import json
import math
import re
from datetime import datetime

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import fds_actores as FA
import fds_listas as FL


# ══════════════════════════════════════════════════════════════════════════
# PALETA
# ══════════════════════════════════════════════════════════════════════════
# Los colores no se eligieron a ojo: se validaron con el verificador de
# paletas (banda de luminosidad, piso de croma, separacion bajo simulacion
# de daltonismo y contraste contra el fondo real de cada tema). Cada juego
# declara al lado el peor par adyacente medido, para que quien lo modifique
# vuelva a correr la validacion y no rompa la legibilidad.
#
# Regla de uso (una sola por grafico):
#   - magnitud sin signo  -> rampa secuencial de un solo tono (NEUTRA)
#   - magnitud adversa    -> rampa CRITICA (a mas oscuro, mas grave)
#   - magnitud favorable  -> rampa FAVORABLE
#   - identidad           -> paleta CATEGORICA, en orden fijo, nunca ciclada
#   - escala ordenada     -> paleta DIVERGENTE, centrada en el neutro
#   - estado si/no        -> par SI/NO (siempre con leyenda y tabla al lado)

ANIN_VERDE = "1B4D2E"
ANIN_AZUL = "1B4F72"
ANIN_VERDE_CLARO = "E8F0EA"
ANIN_GRIS = "F2F2F2"

ENCABEZADOS_ANIN = [
    "AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN",
    "DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME",
    "SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN - SESDI",
]

SUBTITULO_PROYECTO = (
    "PROYECTO IN PIURA | CUI 2669244 | Recuperación del servicio de regulación "
    "de riesgos naturales y recuperación de ecosistemas degradados en la "
    "Cuenca Alta del Río Piura | UTM WGS 84 Zona 17S (EPSG:32717)"
)

# Identidad (8 ranuras, orden fijo). Claro: peor par adyacente DeltaE CVD 14.0,
# vision normal 17.3 sobre #FFFFFF. Oscuro: 8.6 / 19.3 sobre #0E1117.
CATEGORICA_CLARA = ["#2E7D4F", "#1F6FB2", "#C4A03C", "#C0392B",
                    "#6A4C93", "#1BA39C", "#D2792E", "#C2477F"]
CATEGORICA_OSCURA = ["#199E70", "#3987E5", "#D95926", "#9085E9",
                     "#E66767", "#008300", "#D55181", "#C98500"]

# Rampas ordinales de un solo tono (claro -> oscuro = menos -> mas).
RAMPA_NEUTRA = ["#93BCDD", "#6BA1CD", "#4785B8", "#28679F", "#164C77"]
RAMPA_CRITICA = ["#E8A08E", "#D97B62", "#C55437", "#A63A22", "#8A2F18"]
RAMPA_FAVORABLE = ["#8FC2A0", "#66AC7C", "#40945C", "#2A7A48", "#14592F"]

# Divergente para escalas ordenadas (favorable <-> desfavorable) con neutro
# gris al centro. Peor par adyacente DeltaE CVD 10.1, vision normal 15.1.
DIVERGENTE = ["#14592F", "#66AC7C", "#C2BFB4", "#D97B62", "#8A2F18"]

SIN_DATO = "#8A8F98"

TINTAS = {
    "claro": {
        "categorica": CATEGORICA_CLARA,
        "si": "#2E7D4F", "no": "#C0392B", "sin_dato": SIN_DATO,
        "fondo": "#FFFFFF", "tinta": "#1A1A18", "tinta_2": "#52514E",
        "apagado": "#7A7975", "rejilla": "#E4E3DC",
    },
    "oscuro": {
        "categorica": CATEGORICA_OSCURA,
        "si": "#199E70", "no": "#E66767", "sin_dato": SIN_DATO,
        "fondo": "#0E1117", "tinta": "#F5F5F2", "tinta_2": "#C3C2B7",
        "apagado": "#898781", "rejilla": "#2C2C2A",
    },
}

FUENTE = "Arial, Helvetica, sans-serif"

# Maximo de clases con color propio: pasado ese punto las clases adyacentes
# se confunden, de modo que el resto se agrupa en "Otros".
MAX_CLASES = 8


def _rampa(valores, rampa):
    """Reparte una rampa de 5 pasos entre `n` clases ordenadas."""
    n = len(valores)
    if n == 0:
        return []
    if n == 1:
        return [rampa[len(rampa) // 2]]
    return [rampa[round(i * (len(rampa) - 1) / (n - 1))] for i in range(n)]


def _divergente(valores):
    """Reparte la paleta divergente entre `n` clases ordenadas."""
    return _rampa(valores, DIVERGENTE)


def _categorica(valores, tema="claro"):
    paleta = TINTAS[tema]["categorica"]
    return [paleta[i % len(paleta)] for i in range(len(valores))]


# ══════════════════════════════════════════════════════════════════════════
# LECTURA DE LAS FICHAS
# ══════════════════════════════════════════════════════════════════════════

FICHAS_DS = ["F-DS-01", "F-DS-02", "F-DS-03", "F-DS-04",
             "F-DS-05", "F-DS-06", "F-DS-07"]

FICHAS_DS_TITULOS = {
    "F-DS-01": "Diagnostico Socioeconomico del CP / Comunidad",
    "F-DS-02": "Mapeo y Analisis de Actores Clave",
    "F-DS-03": "Entrevista Semiestructurada a Autoridades / Lideres",
    "F-DS-04": "Acta de Taller Participativo",
    "F-DS-05": "Conflictos Socioambientales y Oportunidades",
    "F-DS-06": "Percepcion Local de Peligros y Cambio Climatico",
    "F-DS-07": "Disposicion a Participar y Consentimiento Previo Informado",
}

_SI = {"si", "sí", "s", "yes", "1", "true", "x"}
_NO = {"no", "n", "0", "false"}


def _txt(valor):
    """Texto limpio; los vacios y marcadores de ausencia devuelven ''."""
    if valor is None:
        return ""
    texto = str(valor).strip()
    if texto.lower() in ("", "nan", "none", "-", "s/d", "sin dato"):
        return ""
    return texto


def _num(valor):
    """Numero a partir del texto libre de campo (tolera %, unidades y miles).

    Los campos numericos de las fichas son cajas de texto: el tecnico escribe
    "1,250", "35 %", "12 ha" o "aprox. 40". Se recupera la primera cifra y se
    devuelve None cuando no hay ninguna, para no confundir "sin dato" con 0.
    """
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return None if isinstance(valor, float) and math.isnan(valor) else float(valor)
    texto = str(valor).strip()
    if not texto:
        return None
    texto = texto.replace(" ", "")
    # Separador de miles con coma y decimal con punto (1,250.5) o al reves.
    if re.search(r",\d{3}(\D|$)", texto):
        texto = texto.replace(",", "")
    else:
        texto = texto.replace(",", ".")
    m = re.search(r"-?\d+(?:\.\d+)?", texto)
    if not m:
        return None
    try:
        return float(m.group(0))
    except ValueError:
        return None


def _entero(valor):
    """Conteo (habitantes, familias) a partir del texto de campo.

    Igual que `_num`, pero un punto seguido de tres cifras es separador de
    miles ("1.500" son mil quinientos habitantes, no uno y medio).
    """
    texto = _txt(valor).replace(" ", "")
    if re.match(r"^\d{1,3}(\.\d{3})+$", texto):
        return float(texto.replace(".", ""))
    return _num(valor)


def _pct(valor):
    """Porcentaje a partir del texto de campo.

    La cobertura suele escribirse con aclaraciones ("cada 15 días 50%",
    "80% (de 8am a 6pm)"): manda la cifra que lleva el signo %, y solo a
    falta de ella la primera cifra del texto. No recorta: un valor fuera de
    0-100 se devuelve tal cual para que el control de calidad lo observe.
    """
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return _num(valor)
    texto = _txt(valor)
    if not texto:
        return None
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*%", texto)
    if m:
        return _num(m.group(1))
    # Sin signo %, solo se acepta una cifra sola: "80'/día" o "30 min" no
    # son un porcentaje de viviendas sino otra medida anotada en la casilla.
    if re.fullmatch(r"\d+(?:[.,]\d+)?", texto.replace(" ", "")):
        return _num(texto)
    return None


def _pct_valido(valor):
    """Porcentaje dentro de 0-100, o None (sin dato o fuera de rango)."""
    pct = _pct(valor)
    return pct if pct is not None and 0 <= pct <= 100 else None


def _sino(valor):
    """Normaliza a 'Si', 'No', 'No aplica' o '' (sin responder)."""
    texto = _txt(valor).lower()
    if not texto:
        return ""
    if texto in _SI:
        return "Sí"
    if texto in _NO:
        return "No"
    if texto.startswith("no aplica") or texto in ("na", "n/a"):
        return "No aplica"
    if texto.startswith("parcial"):
        return "Parcial"
    return _txt(valor)


def _lista(valor, opciones=None):
    """Valor de un multiselect: lista de opciones marcadas (sin repetir).

    Las fichas antiguas o importadas guardan el marcado como texto unido
    ("Letrina seca / Pozo séptico"). Como algunas opciones llevan su propia
    barra ("JASS / Sistema local"), el texto se separa reconociendo las
    opciones de la lista de la ficha y no partiendo en cada barra.
    """
    if isinstance(valor, (list, tuple, set)):
        return list(dict.fromkeys(_txt(v) for v in valor if _txt(v)))
    texto = _txt(valor)
    if not texto:
        return []
    if opciones and texto not in opciones:
        clave = _clave(texto)
        halladas = [o for o in opciones if _clave(o) and _clave(o) in clave]
        # Una opcion contenida en otra mas larga ya hallada no cuenta dos veces.
        halladas = [o for o in halladas
                    if not any(o != h and _clave(o) in _clave(h) for h in halladas)]
        if halladas:
            return halladas
    if ";" in texto:
        return list(dict.fromkeys(p.strip() for p in texto.split(";") if p.strip()))
    return [texto]


def formulario(registro):
    """Dict del formulario V4 guardado en `dsNN_data_v3`."""
    ficha = registro.get("ficha", "") or ""
    num = ficha.split("-")[-1] if ficha else ""
    crudo = registro.get(f"ds{num}_data_v3", "") or ""
    if isinstance(crudo, dict):
        return crudo
    if not crudo:
        return {}
    try:
        valor = json.loads(crudo)
    except (json.JSONDecodeError, TypeError):
        return {}
    return valor if isinstance(valor, dict) else {}


def _tabla(form, slot):
    """Filas de un slot de tabla (st.data_editor) del formulario."""
    filas = form.get(slot, [])
    if not isinstance(filas, list):
        return []
    return [f for f in filas if isinstance(f, dict)]


def _col(fila, *nombres):
    """Valor de la primera columna presente (los encabezados cambian de
    plantilla a plantilla; se aceptan alias)."""
    for nombre in nombres:
        if nombre in fila:
            return _txt(fila[nombre])
    # Busqueda laxa: sin tildes, sin puntuacion y en minusculas.
    claves = {_clave(k): v for k, v in fila.items()}
    for nombre in nombres:
        valor = claves.get(_clave(nombre))
        if valor is not None:
            return _txt(valor)
    return ""


_TILDES = str.maketrans("áéíóúüñÁÉÍÓÚÜÑ", "aeiouunAEIOUUN")


def _clave(texto):
    return re.sub(r"[^a-z0-9]", "", str(texto).translate(_TILDES).lower())


def _clave_dedup(reg):
    """Clave de una ficha para descartar sus copias.

    El nombre del responsable no forma parte de la clave: el control de
    guardado obligaba a variarlo ("Stefany Campos..") para registrar una
    segunda ficha, de modo que no sirve para reconocer copias.
      - F-DS-03: un entrevistado; F-DS-04: un taller (ambos por CP, aunque
        se hayan cargado en dos bloques que lo comparten); F-DS-07: un
        titular del bloque. La misma persona o el mismo taller registrados
        dos veces son una copia aunque cambie el texto.
      - Demas fichas: copia identica (mismo bloque, ambito, entrevistado y
        contenido del formulario). Las fichas distintas de un mismo CP se
        conservan y se consolidan despues, por CP.
    """
    ficha = reg.get("ficha", "") or ""
    form = formulario(reg)
    bloque = _txt(reg.get("bloque_codigo"))
    ambito = _clave(_txt(reg.get("centro_poblado")) or _comunidad(reg))
    distrito = _clave(_distrito(reg))
    # La entrevista y el taller no dependen del bloque: cargados en dos
    # bloques que comparten el CP (Miguel Pampa en 83 y 84) son uno solo.
    if ficha == "F-DS-03" and _clave(form.get("f3_nombre")):
        return (ficha, distrito, ambito, _clave(form.get("f3_nombre")))
    if ficha == "F-DS-04" and (_txt(form.get("f4_fecha")) or _txt(form.get("f4_lugar"))):
        return (ficha, distrito, ambito, _txt(form.get("f4_fecha")),
                _clave(form.get("f4_lugar")))
    if ficha == "F-DS-07" and (_txt(form.get("f7_dni")) or _clave(form.get("f7_nombre"))):
        # DNI y nombre juntos: un DNI mal digitado no debe fundir a dos
        # titulares distintos (perder un titular es peor que contar dos veces
        # al mismo, que el control de calidad deja a la vista).
        return (ficha, bloque, re.sub(r"\D", "", _txt(form.get("f7_dni"))),
                _clave(form.get("f7_nombre")))
    contenido = json.dumps(form, ensure_ascii=False, sort_keys=True, default=str)
    return (ficha, bloque, ambito, _clave(reg.get("nombre_entrevistado")),
            contenido)


def deduplicar(registros):
    """Descarta las copias de una misma ficha y conserva la mas reciente.

    Ver _clave_dedup. Devuelve las fichas de la mas reciente a la mas antigua.
    """
    vistos, salida = set(), []
    for reg in sorted(registros or [], key=_orden_reciente, reverse=True):
        clave = _clave_dedup(reg)
        if clave not in vistos:
            vistos.add(clave)
            salida.append(reg)
    return salida


def _minutos(valor):
    """Duracion en minutos a partir del texto de campo ("1 hora", "1h 30m",
    "45 min", "una hora y media"). None si no se reconoce."""
    texto = _txt(valor).lower()
    if not texto:
        return None
    m = re.fullmatch(r"(\d{1,2}):(\d{2})(?::\d{2})?\s*(?:h|hrs?|horas?)?", texto)
    if m:                                       # "1:30" = una hora y media
        return 60 * int(m.group(1)) + int(m.group(2)) or None
    texto = (texto.replace("una hora y media", "1 hora 30 min")
             .replace("hora y media", "1 hora 30 min").replace("una hora", "1 hora")
             .replace("media hora", "30 min").replace("y media", " 30 min"))
    horas = re.search(r"(\d+(?:[.,]\d+)?)\s*(?:horas?|hrs?|h)(?![a-z])", texto)
    if horas:
        total = 60 * float(horas.group(1).replace(",", "."))
        # Los minutos van despues de las horas, con o sin unidad ("1h30").
        resto = re.match(r"\s*(?:y\s*)?(\d+)", texto[horas.end():])
        if resto:
            total += float(resto.group(1))
        return total or None
    mins = re.search(r"(\d+(?:[.,]\d+)?)\s*(?:m\b|min|')", texto)
    if mins:
        return float(mins.group(1).replace(",", ".")) or None
    return _num(texto)


def _por_ficha(registros, ficha):
    return [r for r in registros if (r.get("ficha", "") or "") == ficha]


def _comunidad(registro):
    """Comunidad campesina declarada; "Ninguna" y sus variantes no lo son."""
    texto = _txt(registro.get("comunidad_campesina"))
    if re.match(r"^(ningun|no pertenec|sin comunidad|no tiene|n/?a$)",
                texto.strip(" .").lower()):
        return ""
    return texto


def _ambito(registro):
    """Etiqueta del ambito de un registro: centro poblado o, a falta de el,
    la comunidad campesina o el bloque."""
    return (_txt(registro.get("centro_poblado"))
            or _comunidad(registro)
            or _txt(registro.get("bloque_codigo"))
            or "Sin ambito consignado")


# ══════════════════════════════════════════════════════════════════════════
# UNIDAD DE ANALISIS: EL CENTRO POBLADO
# ══════════════════════════════════════════════════════════════════════════
# La F-DS-01 describe al centro poblado, no al informante. En campo se aplica
# mas de una ficha a un mismo CP (varios informantes, reediciones, o el mismo
# CP asociado a dos o tres bloques), y el control de duplicados del registro
# obligaba a variar el nombre del responsable ("Stefany Campos..",
# "Stefany Campos A") para guardar la segunda. Contar o sumar fichas inflaba
# la poblacion, llevaba las coberturas por encima del 100 % y contaba un mismo
# caserio tantas veces como fichas tuviera. Aqui cada centro poblado se
# identifica una sola vez y sus fichas se consolidan en una sola fila.
#
# Identidad: nombre normalizado (sin tildes, mayusculas ni puntuacion) y
# distrito de la ficha. Cuando el nombre figura en el catalogo INEI del bloque
# se usa ademas su ubicacion, que une al CP compartido por bloques de distritos
# distintos y separa a los homonimos (p. ej. Coyona de Canchaque y Coyona de
# San Miguel de El Faique son dos centros poblados distintos).

def _catalogo_bloque(codigo):
    try:
        import centros_poblados as CPB
    except ImportError:          # el catalogo es opcional para el modulo
        return {}
    return CPB.datos_bloque(_txt(codigo)) or {}


def _partes_cp(nombre):
    """'A / B / C' -> ['A', 'B', 'C']: ficha registrada para varios CP."""
    return [p.strip() for p in re.split(r"\s+/\s+|;", _txt(nombre)) if p.strip()]


def _ubicacion_catalogo(bloque, nombre):
    """(este, norte) del CP en el catalogo INEI del bloque, o None."""
    clave = _clave(nombre)
    for fila in _catalogo_bloque(bloque).get("demografia") or []:
        if _clave(fila.get("centro_poblado")) == clave:
            este, norte = fila.get("utm_este") or 0, fila.get("utm_norte") or 0
            if este and norte:
                return (round(float(este)), round(float(norte)))
    return None


def _poblacion_catalogo(bloque, nombre):
    """Poblacion INEI del CP en el catalogo del bloque (None si no figura o
    si el catalogo no la trae)."""
    clave = _clave(nombre)
    for fila in _catalogo_bloque(bloque).get("demografia") or []:
        if _clave(fila.get("centro_poblado")) == clave:
            return fila.get("poblacion_total") or None
    return None


def _distrito(registro):
    """Distrito de la ficha; si falta, el del bloque."""
    return (_txt(registro.get("distrito"))
            or _txt(registro.get("bloque_distrito")))


def _identidad(registro):
    """Identidad del ambito de una ficha: tipo, nombre, distrito, bloque y,
    si el CP figura en el catalogo INEI de su bloque, su ubicacion (xy)."""
    nombre = _txt(registro.get("centro_poblado"))
    bloque = _txt(registro.get("bloque_codigo"))
    distrito = _distrito(registro)
    if not nombre:
        # Sin CP consignado: el ambito es la comunidad o, a falta de ella,
        # el propio bloque. No se mezcla con ningun otro ambito.
        comunidad = _comunidad(registro)
        if comunidad:
            return {"nombre": comunidad, "distrito": distrito, "xy": None,
                    "clave": ("cc", _clave(comunidad), _clave(distrito)),
                    "tipo": "comunidad", "bloque": bloque}
        etiqueta = f"Bloque {bloque} (sin CP consignado)" if bloque else \
            "Sin ámbito consignado"
        return {"nombre": etiqueta, "distrito": distrito, "xy": None,
                "clave": ("bloque", bloque), "tipo": "sin_cp", "bloque": bloque}
    partes = _partes_cp(nombre)
    if len(partes) > 1:
        # Una sola ficha para varios CP: sus cifras no se pueden repartir,
        # de modo que el conjunto es su propio ambito.
        clave = "+".join(sorted(_clave(p) for p in partes))
        return {"nombre": " / ".join(partes), "distrito": distrito, "xy": None,
                "clave": ("nd", clave, _clave(distrito)),
                "tipo": "compuesto", "bloque": bloque}
    return {"nombre": nombre, "distrito": distrito,
            "xy": _ubicacion_catalogo(bloque, nombre),
            "clave": ("nd", _clave(nombre), _clave(distrito)),
            "tipo": "cp", "bloque": bloque}


def _agrupar_ambitos(registros):
    """Asigna a cada registro la clave de su centro poblado.

    La ubicacion del catalogo INEI manda: dos fichas con la misma ubicacion
    son el mismo CP aunque sus bloques esten en distritos distintos
    (Chililique Alto, Maray), y dos con ubicaciones distintas son CP
    distintos aunque se llamen igual (las dos Coyona). Una ficha cuyo CP no
    figura en el catalogo de su bloque se une a la unica ubicacion conocida
    de ese nombre en su distrito; si no la hay, a las demas fichas del mismo
    nombre y distrito. Devuelve (clave_por_registro, etiquetas), donde la
    etiqueta es el nombre a mostrar, con el distrito entre parentesis solo
    cuando hay homonimos.
    """
    identidades = [_identidad(r) for r in registros]
    ubicaciones = {}
    for ident in identidades:
        if ident["tipo"] == "cp" and ident["xy"]:
            ubicaciones.setdefault(ident["clave"], set()).add(ident["xy"])
    grupos = {}
    for i, ident in enumerate(identidades):
        k = ident["clave"]
        if ident["tipo"] == "cp":
            xy = ident["xy"]
            if not xy and len(ubicaciones.get(k, ())) == 1:
                xy = next(iter(ubicaciones[k]))
            if xy:
                k = ("xy",) + tuple(xy)
        grupos.setdefault(k, []).append(i)

    claves, nombres, distritos, bloques = {}, {}, {}, {}
    for k, indices in grupos.items():
        # Nombre a mostrar: la grafia mas frecuente entre sus fichas.
        conteo = {}
        for i in indices:
            n = identidades[i]["nombre"]
            conteo[n] = conteo.get(n, 0) + 1
        nombre = sorted(conteo, key=lambda n: (-conteo[n], n))[0]
        clave = "|".join(str(p) for p in k)
        nombres[clave] = nombre
        distritos[clave] = next((identidades[i]["distrito"] for i in indices
                                 if identidades[i]["distrito"]), "")
        bloques[clave] = sorted({identidades[i]["bloque"] for i in indices
                                 if identidades[i]["bloque"]})
        for i in indices:
            claves[i] = clave
    # Homonimos (mismo nombre, distinto CP): se distinguen por el distrito y,
    # si comparten distrito, tambien por el bloque.
    por_nombre = {}
    for clave, nombre in nombres.items():
        por_nombre.setdefault(_clave(nombre), []).append(clave)
    etiquetas = {}
    for lista in por_nombre.values():
        mismo_distrito = {}
        for clave in lista:
            mismo_distrito.setdefault(_clave(distritos[clave]), []).append(clave)
        for clave in lista:
            etiqueta = nombres[clave]
            if len(lista) > 1:
                detalle = [distritos[clave]] if distritos[clave] else []
                if len(mismo_distrito[_clave(distritos[clave])]) > 1 and bloques[clave]:
                    detalle.append("bloque " + ", ".join(bloques[clave]))
                if detalle:
                    etiqueta = f"{etiqueta} ({', '.join(detalle)})"
            etiquetas[clave] = etiqueta
    # Ultimo recurso: si aun se repite, se numera.
    vistas = {}
    for clave in sorted(etiquetas):
        etiqueta = etiquetas[clave]
        if etiqueta in vistas:
            vistas[etiqueta] += 1
            etiquetas[clave] = f"{etiqueta} [{vistas[etiqueta]}]"
        else:
            vistas[etiqueta] = 1
    return claves, etiquetas


def _etiquetas_por_registro(registros):
    """{id(registro): etiqueta del CP}, calculada sobre TODAS las fichas del
    informe: asi un homonimo se rotula igual ("Coyona (Canchaque)") en todas
    las secciones, aunque en alguna ficha solo figure uno de los dos."""
    claves, etiquetas = _agrupar_ambitos(registros)
    return {id(r): etiquetas[claves[i]] for i, r in enumerate(registros)}


def _tiene(valor):
    if isinstance(valor, (list, tuple)):
        return any(_tiene(v) for v in valor)
    if isinstance(valor, dict):
        return any(_tiene(v) for v in valor.values())
    return bool(_txt(valor))


def _es_tabla(valor):
    return isinstance(valor, list) and any(isinstance(v, dict) for v in valor)


# Campos que se leen juntos: la poblacion total, su desagregacion por sexo y
# por edad salen siempre de una misma ficha, para que no se mezclen las
# cifras de dos informantes (hombres de uno, total de otro).
_GRUPO_DEMOGRAFIA = ("f1_pob_t", "f1_pob_h", "f1_pob_m", "f1_pob_men18",
                     "f1_pob_may65")

# La priorizacion local de peligros (1.°, 2.° y 3.°) es una sola respuesta:
# se toma completa de una ficha para no repetir un peligro en dos puestos.
_GRUPO_PRIORIDAD = ("f6_p1", "f6_p2", "f6_p3")

_CAMPOS_PCT = {"f1_agua_cob", "f1_energia_cob", "f1_pct_tituladas"}

# Campos de marcado multiple y sus opciones (para separar el texto heredado).
_OPCIONES_MULTIPLES = {"f1_agua": FL.L_AGUA, "f1_sanea": FL.L_SANEA,
                       "f1_energia": FL.L_ENERG}

# Tablas que se combinan por su columna clave: cada fila (actividad, peligro,
# cambio, oportunidad) se toma de la ficha de mayor prioridad que la
# registre, de modo que una misma actividad no se suma dos veces.
_CLAVE_TABLA = {
    "f1_activ": ("Actividad / Rubro", "Actividad"),
    "f5_conflictos": ("Tipo",),
    "f5_oportunidades": ("Oportunidad identificada",),
    "f6_peligros": ("Peligro observado",),
    "f6_cambios": ("Cambio observado",),
}

# Campos de la F-DS-01 cuyo desacuerdo entre fichas de un mismo CP se reporta
# en el control de calidad. (clave, etiqueta)
_CONTROL_FDS01 = [
    ("f1_pob_t", "Población total (hab.)"),
    ("f1_pob_h", "Hombres"),
    ("f1_pob_m", "Mujeres"),
    ("f1_nfam", "Familias / viviendas"),
    ("f1_agua_cob", "Cobertura de agua (%)"),
    ("f1_energia_cob", "Cobertura de energía (%)"),
    ("f1_juntos", "JUNTOS (familias)"),
    ("f1_pension65", "Pensión 65 (personas)"),
    ("f1_migracion", "Tasa de migración juvenil"),
    ("f1_presencia_estatal", "Percepción de presencia estatal"),
]


def _valor_norm(campo, valor):
    """Forma comparable de un valor: '100 %' y '100' son el mismo dato."""
    if isinstance(valor, (list, tuple)):
        return ("l",) + tuple(sorted(_clave(v) for v in valor if _txt(v)))
    texto = _txt(valor)
    if campo in _CAMPOS_PCT:
        pct = _pct(texto)
        if pct is not None:
            return ("n", pct)
    elif re.fullmatch(r"[\d\s.,]+", texto):
        numero = _entero(texto)
        if numero is not None:
            return ("n", numero)
    return ("t", _clave(texto))


def _mostrar(campo, valor):
    """Valor legible para el control de calidad."""
    norm = _valor_norm(campo, valor)
    if norm[0] == "n":
        return _fmt_valor(norm[1])
    return _txt(valor)


def _fmt_valor(valor):
    if isinstance(valor, float) and valor.is_integer():
        return f"{int(valor):,}".replace(",", " ")
    if isinstance(valor, float):
        return f"{valor:,.1f}".replace(",", " ")
    return str(valor)


def _moda(campo, formularios):
    """(valor, posicion, n_fichas, n_con_dato) del valor mas frecuente de un
    campo entre las fichas (ordenadas de la mas reciente a la mas antigua);
    en caso de empate manda la ficha mas reciente. None si ninguna lo trae."""
    conteo, primero = {}, {}
    validos = (campo in _CAMPOS_PCT and any(
        _pct_valido(f.get(campo)) is not None for f in formularios))
    for pos, form in enumerate(formularios):
        valor = form.get(campo)
        if not _tiene(valor):
            continue
        # Un porcentaje ilegible ("80'/dia") no le gana a uno valido de otra
        # ficha del mismo CP: se lo sigue informando en el control.
        if validos and _pct_valido(valor) is None:
            continue
        k = _valor_norm(campo, valor)
        conteo[k] = conteo.get(k, 0) + 1
        primero.setdefault(k, (pos, valor))
    if not conteo:
        return None
    mejor = max(conteo, key=lambda k: (conteo[k], -primero[k][0]))
    pos, valor = primero[mejor]
    return valor, pos, conteo[mejor], sum(conteo.values())


def _completitud(registro):
    """N.° de campos con dato en el formulario (desempata fichas del mismo dia)."""
    return sum(1 for v in formulario(registro).values() if _tiene(v))


def _orden_reciente(registro):
    """Clave de orden: mas reciente y mas completa primero."""
    return (_txt(registro.get("fecha_evaluacion")), _completitud(registro),
            _txt(registro.get("fecha_registro")), registro.get("id") or 0)


def _combinar_tabla(campo, formularios, orden):
    """Filas de una tabla del formulario combinadas entre las fichas de un CP.

    `orden` da la prioridad de las fichas (la de referencia primero). Cada
    clave de fila (p. ej. la actividad) se toma completa de la primera ficha
    que la registra; sin columna clave conocida, manda la tabla de la ficha
    de mayor prioridad que tenga filas."""
    columnas = _CLAVE_TABLA.get(campo)
    if not columnas:
        for i in orden:
            filas = _tabla(formularios[i], campo)
            if filas:
                return filas
        return []
    def con_datos(fila):
        # La app precarga las filas de peligros y cambios con solo su nombre:
        # una fila sin ningun otro dato no "registra" la clave y no debe
        # tapar la respuesta de una ficha mas antigua del mismo CP.
        return any(_txt(v) for k, v in fila.items()
                   if _clave(k) not in {_clave(c) for c in columnas})

    salida, tomadas, vacias = [], set(), {}
    for i in orden:
        propias = {}
        for fila in _tabla(formularios[i], campo):
            clave = _clave(_col(fila, *columnas))
            if not clave or clave in tomadas:
                continue
            if not con_datos(fila):
                vacias.setdefault(clave, fila)
                continue
            propias.setdefault(clave, []).append(fila)
        for clave, filas in propias.items():
            tomadas.add(clave)
            salida += filas
    salida += [fila for clave, fila in vacias.items() if clave not in tomadas]
    return salida


def _consolidar_formularios(fichas, referencia=0):
    """Formulario unico de un CP a partir de sus fichas.

    `fichas` va de la mas reciente a la mas antigua; `referencia` es la
    posicion de la ficha de referencia. Reglas (nunca se suman fichas):
      - cada dato es el valor mas frecuente entre las fichas que lo consignan
        (moda); en caso de empate, el de la ficha mas reciente;
      - poblacion total, sexo y edad salen juntos de la ficha de referencia
        (o, si no los trae, de la mas reciente que los consigne);
      - las opciones de marcado multiple se reunen: el CP reporta una opcion
        si alguna de sus fichas la marca;
      - las tablas se combinan por su columna clave (ver _combinar_tabla).
    """
    formularios = [formulario(r) for r in fichas]
    if not formularios:
        return {}
    orden = [referencia] + [i for i in range(len(formularios)) if i != referencia]
    base = dict(formularios[referencia])
    campos = list(dict.fromkeys(k for f in formularios for k in f))
    for campo in campos:
        if campo in _GRUPO_DEMOGRAFIA or campo in _GRUPO_PRIORIDAD:
            continue
        valores = [f.get(campo) for f in formularios if _tiene(f.get(campo))]
        if not valores:
            continue
        if any(_es_tabla(v) for v in valores):
            base[campo] = _combinar_tabla(campo, formularios, orden)
        elif campo in _OPCIONES_MULTIPLES or any(isinstance(v, (list, tuple))
                                                 for v in valores):
            # Se decide por el campo, no por el tipo guardado: las fichas
            # antiguas traen el marcado como texto ("Letrina seca / Pozo
            # septico") y deben separarse con las opciones de la ficha.
            reunidas = []
            for i in orden:
                reunidas += _lista(formularios[i].get(campo),
                                   opciones=_OPCIONES_MULTIPLES.get(campo))
            base[campo] = list(dict.fromkeys(reunidas))
        else:
            base[campo] = _moda(campo, formularios)[0]
    for grupo in (_GRUPO_DEMOGRAFIA, _GRUPO_PRIORIDAD):
        fuente = next((i for i in orden
                       if any(_tiene(formularios[i].get(k)) for k in grupo)), None)
        for k in grupo:
            base[k] = formularios[fuente].get(k, "") if fuente is not None else ""
    return base


def _referencia(fichas, ficha):
    """Posicion de la ficha de referencia de un CP.

    En la F-DS-01 es la que declara la poblacion total mas frecuente entre
    las fichas del CP (asi los datos demograficos son los que mas informantes
    respaldan); en empate, la mas reciente. En las demas fichas, la mas
    reciente."""
    if ficha != "F-DS-01":
        return 0
    formularios = [formulario(r) for r in fichas]
    moda = _moda("f1_pob_t", formularios)
    if not moda:
        return 0
    modal = _valor_norm("f1_pob_t", moda[0])
    # Entre las fichas que declaran esa poblacion, la que mas la desagrega
    # (sexo y edad); en empate, la mas reciente.
    candidatas = [i for i, f in enumerate(formularios)
                  if _tiene(f.get("f1_pob_t"))
                  and _valor_norm("f1_pob_t", f.get("f1_pob_t")) == modal]
    return max(candidatas, key=lambda i: (
        sum(1 for k in _GRUPO_DEMOGRAFIA if _tiene(formularios[i].get(k))), -i))


def unidades_por_cp(registros, ficha="F-DS-01"):
    """Una entrada por centro poblado con sus fichas de un tipo consolidadas.

    Cada unidad trae: clave, etiqueta (nombre a mostrar), distrito, bloques,
    fichas (de la mas reciente a la mas antigua), referencia (la ficha que
    manda), form (formulario consolidado), discrepancias (campos en que las
    fichas del CP no coinciden) y tipo ("cp", "compuesto", "comunidad" o
    "sin_cp").
    """
    registros = list(registros or [])
    claves, etiquetas = _agrupar_ambitos(registros)
    grupos = {}
    for i, reg in enumerate(registros):
        if (reg.get("ficha", "") or "") == ficha:
            grupos.setdefault(claves[i], []).append(reg)
    unidades = []
    for clave, fichas in grupos.items():
        fichas = sorted(fichas, key=_orden_reciente, reverse=True)
        ref = _referencia(fichas, ficha)
        form = _consolidar_formularios(fichas, ref)
        discrepancias = []
        if ficha == "F-DS-01" and len(fichas) > 1:
            formularios = [formulario(r) for r in fichas]
            for campo, etiqueta in _CONTROL_FDS01:
                distintos = {}
                for f in formularios:
                    if _tiene(f.get(campo)):
                        k = _valor_norm(campo, f.get(campo))
                        distintos.setdefault(k, [f.get(campo), 0])[1] += 1
                if len(distintos) > 1:
                    usado = _valor_norm(campo, form.get(campo))
                    discrepancias.append({
                        "campo": etiqueta,
                        "usado": _mostrar(campo, form.get(campo))
                        or "(la ficha de referencia no lo consigna)",
                        "n_usado": distintos.get(usado, [None, 0])[1],
                        "n_con_dato": sum(n for _v, n in distintos.values()),
                        "otros": [f"{_mostrar(campo, v)} ({n})"
                                  for k, (v, n) in distintos.items() if k != usado]})
        bloques = sorted({_txt(r.get("bloque_codigo")) for r in fichas
                          if _txt(r.get("bloque_codigo"))})
        unidades.append({
            "clave": clave, "etiqueta": etiquetas[clave],
            "distrito": next((_distrito(r) for r in fichas if _distrito(r)), ""),
            "bloques": bloques, "fichas": fichas, "referencia": fichas[ref],
            "form": form, "discrepancias": discrepancias,
            "tipo": _identidad(fichas[0])["tipo"],
        })
    unidades.sort(key=lambda u: _clave(u["etiqueta"]))
    _marcar_repetidos(unidades)
    return unidades


def _marcar_repetidos(unidades):
    """Marca con "suma"=False los ambitos que repetirian a un CP con ficha
    propia: una ficha "A / B" cuando A o B tienen su propia ficha, y una
    ficha sin CP (solo comunidad, o ni eso) en un bloque que ya tiene fichas
    de sus CP. Sumarlos contaria dos veces la misma poblacion (en M10B4,
    Rio Seco Alto y la ficha de la comunidad Carlos Augusto Rivera declaran
    los mismos 233 hab.). No se grafican ni se suman; quedan en el
    control de calidad."""
    propios, bloques_con_cp = set(), set()
    for u in unidades:
        if u["tipo"] == "cp":
            for r in u["fichas"]:
                propios.add((_clave(r.get("centro_poblado")), _clave(_distrito(r))))
            bloques_con_cp.update(u["bloques"])
    for u in unidades:
        u["suma"], u["motivo_excluido"] = True, ""
        if u["tipo"] == "compuesto":
            repetidos = sorted({p for r in u["fichas"]
                                for p in _partes_cp(r.get("centro_poblado"))
                                if (_clave(p), _clave(_distrito(r))) in propios})
            if repetidos:
                u["suma"] = False
                u["motivo_excluido"] = ("sus centros poblados " + ", ".join(repetidos)
                                        + " ya tienen ficha propia")
        elif u["tipo"] in ("comunidad", "sin_cp"):
            comunes = sorted(set(u["bloques"]) & bloques_con_cp)
            if comunes:
                u["suma"] = False
                u["motivo_excluido"] = ("la ficha no consigna centro poblado y el bloque "
                                        + ", ".join(comunes)
                                        + " ya tiene fichas de sus centros poblados")


# Seccion 4 de la F-DS-01 ("Tenencia de la tierra relacionada al bloque").
_CAMPOS_TENENCIA = [
    ("f1_tenencia", "Régimen predominante de tenencia"),
    ("f1_n_predios", "N.° aprox. de predios individuales"),
    ("f1_pct_tituladas", "Tierras tituladas (%)"),
    ("f1_conf_linderos", "Conflictos de linderos registrados"),
    ("f1_superpone", "El bloque se superpone a tierras comunales"),
    ("f1_reg_titulacion", "Organismo responsable del registro"),
]


def tenencia_por_bloque(registros):
    """Un registro por bloque con la seccion 4 de la F-DS-01 consolidada.

    La tenencia de la tierra se responde para el bloque, no para cada centro
    poblado: todas las F-DS-01 del bloque (de cualquiera de sus CP) son
    fuentes del mismo dato. Cada campo es el valor mas frecuente entre ellas;
    en empate, el de la ficha mas reciente. Las fichas que no coinciden se
    informan en "discrepancias"."""
    por_bloque = {}
    for r in _por_ficha(registros or [], "F-DS-01"):
        por_bloque.setdefault(_txt(r.get("bloque_codigo")) or "(sin bloque)", []).append(r)
    salida = []
    for bloque, fichas in sorted(por_bloque.items()):
        fichas = sorted(fichas, key=_orden_reciente, reverse=True)
        formularios = [formulario(r) for r in fichas]
        form, discrepancias = {}, []
        for campo, etiqueta in _CAMPOS_TENENCIA:
            moda = _moda(campo, formularios)
            form[campo] = moda[0] if moda else ""
            distintos = {}
            for f in formularios:
                if _tiene(f.get(campo)):
                    k = _valor_norm(campo, f.get(campo))
                    distintos.setdefault(k, [f.get(campo), 0])[1] += 1
            if len(distintos) > 1:
                usado = _valor_norm(campo, form[campo])
                discrepancias.append({
                    "campo": etiqueta, "usado": _mostrar(campo, form[campo]),
                    "n_usado": distintos.get(usado, [None, 0])[1],
                    "n_con_dato": sum(n for _v, n in distintos.values()),
                    "otros": [f"{_mostrar(campo, v)} ({n})"
                              for k, (v, n) in distintos.items() if k != usado]})
        salida.append({
            "bloque": bloque, "etiqueta": f"Bloque {bloque}",
            "distrito": next((_txt(r.get("bloque_distrito")) or _distrito(r)
                              for r in fichas if _txt(r.get("bloque_distrito"))
                              or _distrito(r)), ""),
            "bloques": [bloque], "fichas": fichas, "referencia": fichas[0],
            "form": form, "discrepancias": discrepancias})
    return salida


def _unidades_fds01(registros):
    """(incluidas, excluidas): unidades de la F-DS-01 que entran al analisis
    y las que se dejan fuera por repetir a un CP con ficha propia."""
    unidades = unidades_por_cp(registros, "F-DS-01")
    return ([u for u in unidades if u["suma"]],
            [u for u in unidades if not u["suma"]])



# ══════════════════════════════════════════════════════════════════════════
# CONSTRUCTORES DE SERIES
# ══════════════════════════════════════════════════════════════════════════
# Cada serie es un dict autodescriptivo que sirve por igual al grafico
# interactivo (Altair), al grafico nativo de Excel y al anexo PDF.

def _serie(id_, titulo, forma, filas, cat, val, **extra):
    serie = {
        "id": id_, "titulo": titulo, "forma": forma, "filas": filas,
        "cat": cat, "val": val, "sub": None, "orden_cat": None,
        "orden_sub": None, "escala": "neutra", "colores": None,
        "unidad": "", "nota": "", "descripcion": "", "eje_x": "", "eje_y": "",
        "decimales": 0,
        # Base de los porcentajes del libro Excel (ver _base_porcentaje):
        # "auto", "fila", "columna", "total", ("ref", subclase), un entero
        # (N.° de fichas o de CP, para marcado multiple) o None (no aplica).
        "pct_base": "auto",
        # Rotulo de la base entera (marcado multiple): "ficha(s)" o
        # "centro(s) poblado(s) con dato".
        "base_texto": "ficha(s)",
        # Totales de la tabla (ver _modo_totales): "auto", "columnas",
        # "filas", "ambos", "promedio", "base" o "ninguno".
        "totales": "auto",
        "etiqueta_total": "Total",
    }
    serie.update(extra)
    # Una serie con color declarado por clase viene de una escala ordenada
    # (disposicion, posicion, nivel de conflictividad): el eje debe respetar
    # ese orden, no reordenarse por frecuencia.
    if serie.get("colores") and not serie.get("orden_cat") and not serie.get("sub"):
        serie["orden_cat"] = list(serie["colores"])
    return serie


def _conteo(valores, orden=None, agrupar_otros=True):
    """Frecuencia de cada valor, respetando un orden dado si se entrega."""
    conteo = {}
    for v in valores:
        texto = _txt(v)
        if not texto:
            continue
        conteo[texto] = conteo.get(texto, 0) + 1
    if orden:
        claves = [o for o in orden if o in conteo]
        claves += sorted(k for k in conteo if k not in orden)
    else:
        claves = sorted(conteo, key=lambda k: (-conteo[k], k))
        if agrupar_otros and len(claves) > MAX_CLASES:
            # Nunca se generan colores nuevos: lo que excede la paleta se
            # agrupa en "Otros" y el detalle queda en la tabla.
            cabeza, cola = claves[:MAX_CLASES - 1], claves[MAX_CLASES - 1:]
            resto = sum(conteo[k] for k in cola)
            conteo = {k: conteo[k] for k in cabeza}
            conteo[f"Otros ({len(cola)} clases)"] = resto
            claves = cabeza + [f"Otros ({len(cola)} clases)"]
    return [{"clase": k, "valor": conteo[k]} for k in claves]


def _apiladas(filas_grupo, orden_sub, rampa=None, categorica=False,
              tema="claro", descendente=False):
    """Normaliza filas {cat, sub, valor} y fija el color de cada subclase.

    El tono se reparte sobre el dominio COMPLETO declarado en la ficha, no
    solo sobre las clases presentes: asi una misma clase conserva su color
    entre un bloque y otro y los graficos se pueden comparar entre si.
    """
    dominio = list(dict.fromkeys(orden_sub))
    if categorica:
        tonos = dict(zip(dominio, _categorica(dominio, tema)))
    else:
        pasos = _rampa(dominio, rampa or RAMPA_NEUTRA)
        # Las listas de la ficha van de mayor a menor ("Alto, Medio, Bajo") y
        # la rampa de claro a oscuro: sin invertir, el nivel alto saldria en
        # el tono mas palido y se leeria al reves.
        tonos = dict(zip(dominio, reversed(pasos) if descendente else pasos))
    vistos = {f["sub"] for f in filas_grupo}
    presentes = [s for s in dominio if s in vistos]
    presentes += sorted(vistos - set(dominio))
    return presentes, {s: tonos.get(s, SIN_DATO) for s in presentes}


def _divergente_centrada(filas_grupo, orden_sub, indice_neutro):
    """Convierte una escala ordenada en barras apiladas divergentes.

    Las clases favorables se apilan hacia la derecha, las desfavorables hacia
    la izquierda y la clase neutra queda a caballo del cero, que es lo que
    permite comparar de un vistazo el saldo de cada ambito. Devuelve las filas
    con los extremos `ini`/`fin` ya calculados en porcentaje.
    """
    salida = []
    grupos = {}
    for fila in filas_grupo:
        grupos.setdefault(fila["cat"], []).append(fila)
    for cat, filas in grupos.items():
        total = sum(f["valor"] for f in filas) or 1
        pct = {f["sub"]: 100.0 * f["valor"] / total for f in filas}
        crudo = {f["sub"]: f["valor"] for f in filas}
        neutro = pct.get(orden_sub[indice_neutro], 0.0)
        # Tramo desfavorable: hacia la izquierda desde el borde del neutro.
        cursor = -neutro / 2
        for sub in reversed(orden_sub[:indice_neutro]):
            ancho = pct.get(sub, 0.0)
            if ancho:
                salida.append({"cat": cat, "sub": sub, "valor": crudo.get(sub, 0),
                               "pct": ancho, "ini": cursor - ancho, "fin": cursor})
            cursor -= ancho
        if neutro:
            salida.append({"cat": cat, "sub": orden_sub[indice_neutro],
                           "valor": crudo.get(orden_sub[indice_neutro], 0),
                           "pct": neutro, "ini": -neutro / 2, "fin": neutro / 2})
        cursor = neutro / 2
        for sub in orden_sub[indice_neutro + 1:]:
            ancho = pct.get(sub, 0.0)
            if ancho:
                salida.append({"cat": cat, "sub": sub, "valor": crudo.get(sub, 0),
                               "pct": ancho, "ini": cursor, "fin": cursor + ancho})
            cursor += ancho
    return salida


def _bateria_sino(registros, campos, orden=("Sí", "No", "No aplica", "Parcial"),
                 form=formulario):
    """Barras apiladas de una bateria de preguntas Si/No.

    `campos` es una lista de (clave, etiqueta). Devuelve filas {cat, sub, valor}
    donde `cat` es la pregunta y `sub` la respuesta. `form` extrae el
    formulario de cada elemento (registro o unidad consolidada por CP).
    """
    filas = []
    for clave, etiqueta in campos:
        conteo = {}
        for reg in registros:
            respuesta = _sino(form(reg).get(clave))
            if respuesta:
                conteo[respuesta] = conteo.get(respuesta, 0) + 1
        for sub in list(orden) + sorted(k for k in conteo if k not in orden):
            if conteo.get(sub):
                filas.append({"cat": etiqueta, "sub": sub, "valor": conteo[sub]})
    return filas


def _colores_sino(tema="claro"):
    t = TINTAS[tema]
    return {"Sí": t["si"], "No": t["no"], "No aplica": t["sin_dato"],
            "Parcial": "#C4A03C"}


# ── F-DS-01: socioeconomico ───────────────────────────────────────────────

_ORDEN_ABC = ["Alto", "Medio", "Bajo"]
_ORDEN_BRM = ["Bueno", "Regular", "Malo"]


def _observacion(unidad, tema, detalle, ficha="F-DS-01"):
    """Fila del control de calidad de los datos."""
    ref = unidad.get("referencia") or {}
    return {"Ficha": ficha,
            "Centro poblado / ámbito": unidad.get("etiqueta", ""),
            "Distrito": unidad.get("distrito", ""),
            "Bloque(s)": ", ".join(unidad.get("bloques") or []),
            "Tema": tema, "Detalle": detalle,
            "Ficha de referencia": _txt(ref.get("fecha_evaluacion")) + (
                f" · {_txt(ref.get('evaluador'))}" if _txt(ref.get("evaluador")) else ""),
            "N.° de fichas del CP": len(unidad.get("fichas") or [])}


def _seccion_socioeconomica(registros, tema="claro"):
    unidades, excluidas = _unidades_fds01(registros)
    if not unidades and not excluidas:
        return None
    n_fichas = sum(len(u["fichas"]) for u in unidades + excluidas)
    series, tablas, control, formato_tablas = [], [], [], {}
    for u in excluidas:
        control.append(_observacion(
            u, "Ámbito repetido (no se suma)",
            f"No se grafica ni se suma: {u['motivo_excluido']}. Si describe a "
            "otro centro poblado, corrija el nombre en la ficha."))

    for u in unidades:
        if u["tipo"] == "compuesto":
            control.append(_observacion(
                u, "Ámbito agrupado",
                "La ficha se registró para varios centros poblados a la vez; "
                "sus cifras no se pueden atribuir a cada uno por separado."))
        elif u["tipo"] == "sin_cp":
            control.append(_observacion(
                u, "Sin centro poblado",
                "La ficha no consigna centro poblado ni comunidad; se muestra "
                "con el código del bloque."))
        for d in u["discrepancias"]:
            control.append(_observacion(
                u, "Fichas que no coinciden",
                f"{d['campo']}: se usa {d['usado']} ({d['n_usado']} de "
                f"{d['n_con_dato']} fichas con dato); las demás fichas del "
                f"mismo CP declaran {', '.join(d['otros'])} (entre paréntesis, "
                "N.° de fichas)."))

    # 1. Poblacion por centro poblado, desagregada por sexo.
    filas_pob, detalle_pob = [], []
    for u in unidades:
        f, ambito = u["form"], u["etiqueta"]
        hombres, mujeres = _entero(f.get("f1_pob_h")), _entero(f.get("f1_pob_m"))
        total = _entero(f.get("f1_pob_t"))
        desagregada = (hombres or 0) + (mujeres or 0)
        # La poblacion total es el dato obligatorio de la ficha y el largo de
        # la barra: lo que no esta desagregado por sexo se muestra como tal.
        # Si hombres + mujeres supera el total, la desagregacion no es
        # confiable y la barra queda entera como "sin desagregar".
        if total is not None and desagregada > total:
            filas_pob.append({"cat": ambito, "sub": "Sin desagregar por sexo",
                              "valor": total})
        else:
            if hombres is not None:
                filas_pob.append({"cat": ambito, "sub": "Hombres", "valor": hombres})
            if mujeres is not None:
                filas_pob.append({"cat": ambito, "sub": "Mujeres", "valor": mujeres})
            if total is not None and total > desagregada:
                filas_pob.append({"cat": ambito, "sub": "Sin desagregar por sexo",
                                  "valor": total - desagregada})
        if total is not None and (hombres is not None or mujeres is not None) \
                and (desagregada > total + 0.5 or (
                    hombres is not None and mujeres is not None
                    and abs(desagregada - total) > 0.5)):
            control.append(_observacion(
                u, "Población",
                f"Hombres + mujeres = {_fmt_valor(desagregada)} no coincide "
                f"con la población total declarada ({_fmt_valor(total)})"
                + ("; el gráfico no la desagrega por sexo."
                   if desagregada > total else ".")))
        inei = None
        if u["tipo"] == "cp":
            inei = next((p for p in (_poblacion_catalogo(b, u["fichas"][0].get(
                "centro_poblado")) for b in u["bloques"]) if p), None)
        if inei and total and (total > 2 * inei or total < inei / 2):
            control.append(_observacion(
                u, "Población",
                f"La población declarada ({_fmt_valor(total)}) difiere en más "
                f"del doble de la registrada por el INEI ({_fmt_valor(float(inei))})."))
        ref = u["referencia"]
        detalle_pob.append({
            "Centro poblado / ámbito": ambito,
            "Distrito": u["distrito"],
            "Bloque(s)": ", ".join(u["bloques"]),
            "Fichas F-DS-01 del CP": len(u["fichas"]),
            "Ficha de referencia": _txt(ref.get("fecha_evaluacion")) + (
                f" · {_txt(ref.get('evaluador'))}" if _txt(ref.get("evaluador")) else ""),
            "Familias / viviendas": _entero(f.get("f1_nfam")),
            "Población total (hab.)": total,
            "Hombres": hombres, "Mujeres": mujeres,
            "Menores de 18 años": _entero(f.get("f1_pob_men18")),
            "Mayores de 65 años": _entero(f.get("f1_pob_may65")),
            "Población originaria": _entero(f.get("f1_pob_orig")),
            "Mano de obra disponible (pers.)": _entero(f.get("f1_mano_obra")),
            "Población INEI (catálogo)": float(inei) if inei else None,
            "Idioma predominante": _txt(f.get("f1_idioma")),
            "Nivel educativo predominante": _txt(f.get("f1_nivel_edu")),
        })
    if filas_pob:
        orden_sub = ["Hombres", "Mujeres", "Sin desagregar por sexo"]
        presentes, colores = _apiladas(filas_pob, orden_sub, categorica=True, tema=tema)
        series.append(_serie(
            "f1_poblacion", "Población por centro poblado, según sexo",
            "apiladas", filas_pob, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="hab.",
            eje_x="Habitantes", eje_y="Centro poblado", escala="categorica",
            totales="ambos", etiqueta_total="Población total",
            descripcion="Población declarada en la ficha F-DS-01, una barra "
                        "por centro poblado: el largo es la población total y "
                        "los tramos, su desagregación por sexo.",
            nota="F-DS-01, numeral 2 (Datos demográficos). Un CP con varias "
                 "fichas se cuenta una sola vez, con los datos de su ficha de "
                 "referencia (la población total más declarada entre sus "
                 "fichas; en empate, la más reciente). «Sin desagregar por "
                 "sexo» es la parte de la población total que la ficha no "
                 "reparte entre hombres y mujeres; si hombres + mujeres supera "
                 "el total, la barra no se desagrega (ver Control de calidad)."))
    if detalle_pob:
        tablas.append(("Demografía por centro poblado", detalle_pob))
        formato_tablas["Demografía por centro poblado"] = {
            "totales": ["Fichas F-DS-01 del CP", "Familias / viviendas",
                        "Población total (hab.)", "Hombres", "Mujeres",
                        "Menores de 18 años", "Mayores de 65 años",
                        "Población originaria", "Mano de obra disponible (pers.)",
                        "Población INEI (catálogo)"],
            "calculadas": [
                ("Hombres + mujeres", "suma", ["Hombres", "Mujeres"]),
                ("Diferencia: total − (H + M)", "resta",
                 ["Población total (hab.)", "Hombres", "Mujeres"]),
                ("Diferencia: declarada − INEI", "resta",
                 ["Población total (hab.)", "Población INEI (catálogo)"]),
            ]}

    # 2. Estructura etaria declarada.
    filas_edad, sin_etaria = [], []
    for u in unidades:
        f = u["form"]
        total = _entero(f.get("f1_pob_t"))
        men, may = _entero(f.get("f1_pob_men18")), _entero(f.get("f1_pob_may65"))
        if total is None or men is None or may is None:
            # Sin los tres datos el tramo intermedio no se puede derivar: se
            # omite el CP en vez de suponer cero menores o cero mayores.
            sin_etaria.append(u["etiqueta"])
            continue
        if men + may > total:
            control.append(_observacion(
                u, "Estructura etaria",
                f"Menores de 18 ({_fmt_valor(men)}) + mayores de 65 "
                f"({_fmt_valor(may)}) superan la población total "
                f"({_fmt_valor(total)})."))
            continue
        ambito = u["etiqueta"]
        filas_edad += [
            {"cat": ambito, "sub": "Menores de 18 años", "valor": men},
            {"cat": ambito, "sub": "De 18 a 65 años", "valor": total - men - may},
            {"cat": ambito, "sub": "Mayores de 65 años", "valor": may},
        ]
    if filas_edad:
        orden_sub = ["Menores de 18 años", "De 18 a 65 años", "Mayores de 65 años"]
        presentes, colores = _apiladas(filas_edad, orden_sub, rampa=RAMPA_NEUTRA)
        nota = ("El tramo de 18 a 65 años se obtiene por diferencia con la "
                "población total; solo se grafican los CP que consignan total, "
                "menores de 18 y mayores de 65.")
        if sin_etaria:
            nota += (" Sin estos tres datos (no se grafican): "
                     + ", ".join(sin_etaria) + ".")
        series.append(_serie(
            "f1_etaria", "Estructura etaria de la población",
            "apiladas", filas_edad, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="hab.",
            eje_x="Habitantes", eje_y="Centro poblado", totales="ambos",
            etiqueta_total="Población total",
            descripcion="Población dependiente (menores de 18 y mayores de 65) "
                        "frente a la población en edad de trabajar.",
            nota=nota))

    # 3. Cobertura de servicios basicos (un valor por CP; nunca se suman %).
    filas_cob = []
    for u in unidades:
        f, ambito = u["form"], u["etiqueta"]
        for clave, etiqueta in (("f1_agua_cob", "Agua para consumo"),
                                ("f1_energia_cob", "Energía eléctrica")):
            texto = _txt(f.get(clave))
            valor = _pct_valido(f.get(clave))
            if valor is not None:
                filas_cob.append({"cat": ambito, "sub": etiqueta, "valor": valor})
            elif texto:
                control.append(_observacion(
                    u, "Cobertura de servicios",
                    f"{etiqueta}: «{texto}» no es un porcentaje entre 0 y 100; "
                    "no se grafica."))
    if filas_cob:
        presentes, colores = _apiladas(
            filas_cob, ["Agua para consumo", "Energía eléctrica"],
            categorica=True, tema=tema)
        series.append(_serie(
            "f1_cobertura", "Cobertura de agua y energía eléctrica (%)",
            "agrupadas", filas_cob, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="%",
            eje_x="Cobertura (%)", eje_y="Centro poblado", maximo=100,
            decimales=0, totales="promedio",
            descripcion="Porcentaje de viviendas con acceso declarado en la "
                        "ficha, un valor por centro poblado (0 a 100 %).",
            nota="F-DS-01, numeral 5 (Servicios básicos e infraestructura "
                 "social). Un CP con varias fichas toma el valor más frecuente "
                 "entre ellas (en empate, el de la más reciente); los "
                 "porcentajes nunca se suman. En el libro Excel, la fila "
                 "PROMEDIO es la media simple de los centros poblados con dato."))

    # 4. Fuentes de agua, saneamiento y energia (marcado multiple), contadas
    #    por centro poblado: cada CP suma una sola vez en cada opcion.
    for clave, titulo, lista, id_ in (
            ("f1_agua", "Fuentes de agua para consumo", FL.L_AGUA, "f1_fuentes_agua"),
            ("f1_sanea", "Tipos de saneamiento", FL.L_SANEA, "f1_saneamiento"),
            ("f1_energia", "Fuentes de energía", FL.L_ENERG, "f1_fuentes_energia")):
        valores, con_dato = [], 0
        for u in unidades:
            opciones = _lista(u["form"].get(clave), opciones=lista)
            if opciones:
                con_dato += 1
                valores += opciones
        filas = _conteo(valores, orden=lista)
        if filas:
            series.append(_serie(
                id_, f"{titulo} (centros poblados que la reportan)",
                "barras_h", filas, "clase", "valor", escala="neutra",
                unidad="CP", eje_x="Centros poblados",
                descripcion=f"Base: {con_dato} centro(s) poblado(s) con dato. "
                            "Cada CP se cuenta una vez por opción; como puede "
                            "reportar más de una, la suma de las barras puede "
                            "superar ese número.",
                nota="F-DS-01, numeral 5. Marcado múltiple en la ficha.",
                pct_base=con_dato, base_texto="centro(s) poblado(s) con dato"))

    # 5. Actividades economicas (familias por CP, sin repetir fichas).
    filas_act, detalle_act = [], []
    for u in unidades:
        f, ambito = u["form"], u["etiqueta"]
        nfam = _entero(f.get("f1_nfam"))
        por_clave, vistas = {}, set()
        for fila in _tabla(f, "f1_activ"):
            actividad = _col(fila, "Actividad / Rubro", "Actividad")
            if not actividad:
                continue
            familias = _entero(_col(fila, "N fam.", "N familias", "Nfam"))
            destino = _col(fila, "Destino") or "Sin destino consignado"
            clave = (actividad, destino)
            if clave in por_clave and familias is not None:
                control.append(_observacion(
                    u, "Actividades económicas",
                    f"«{actividad}» con destino «{destino}» figura más de una "
                    "vez en la ficha; se toma el mayor N.° de familias, no la "
                    "suma."))
            if familias is not None:
                por_clave[clave] = max(por_clave.get(clave, 0), familias)
                if nfam and familias > nfam:
                    control.append(_observacion(
                        u, "Actividades económicas",
                        f"«{actividad}»: {_fmt_valor(familias)} familias supera "
                        f"el total de familias del CP ({_fmt_valor(nfam)})."))
            fila_det = (actividad, familias, _col(fila, "Productos principales"),
                        _col(fila, "Destino"),
                        _num(_col(fila, "Ingreso (S/./mes)", "Ingreso")))
            if fila_det in vistas:
                continue
            vistas.add(fila_det)
            detalle_act.append({
                "Centro poblado / ámbito": ambito,
                "Actividad / Rubro": actividad,
                "N.° de familias": familias,
                "Productos principales": fila_det[2],
                "Destino de la producción": fila_det[3],
                "Ingreso (S/ / mes)": fila_det[4],
            })
        for (actividad, destino), familias in por_clave.items():
            if familias > 0:
                filas_act.append({"cat": actividad, "sub": destino,
                                  "valor": familias})
    if filas_act:
        presentes, colores = _apiladas(filas_act, FL.L_DESTINO, rampa=RAMPA_NEUTRA)
        series.append(_serie(
            "f1_actividades", "Familias por actividad económica y destino de la producción",
            "apiladas", filas_act, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="familias",
            eje_x="Familias dedicadas", totales="ambos",
            eje_y="Actividad / rubro",
            descripcion="El color ordena el destino de la producción, del "
                        "autoconsumo (claro) al mercado (oscuro): es el "
                        "indicador de articulación de los medios de vida.",
            nota="F-DS-01, numeral 6 (Actividades económicas y medios de "
                 "vida). Se suman los centros poblados, cada uno una sola vez: "
                 "cada actividad se toma de una sola ficha del CP (la de "
                 "referencia o, si no la registra, la más reciente que lo "
                 "haga). Una familia puede dedicarse a varias "
                 "actividades: la suma de la tabla no es el N.° de familias del "
                 "ámbito."))
    if detalle_act:
        # Sin total de familias: la misma familia figura en varias actividades.
        tablas.append(("Actividades económicas", detalle_act))

    # 6. Programas sociales (suma de los CP, cada uno una sola vez).
    programas = [("f1_juntos", "JUNTOS (familias)"),
                 ("f1_pension65", "Pensión 65 (personas)"),
                 ("f1_beca18", "Beca 18 (personas)"),
                 ("f1_qaliwarma", "Qali Warma (IIEE)")]
    filas_prog = []
    for clave, etiqueta in programas:
        valores = [_entero(u["form"].get(clave)) for u in unidades]
        valores = [v for v in valores if v is not None]
        total = sum(valores)
        if total:
            filas_prog.append({"clase": f"{etiqueta} · {len(valores)} CP",
                               "valor": total})
    for u in unidades:
        nfam = _entero(u["form"].get("f1_nfam"))
        juntos = _entero(u["form"].get("f1_juntos"))
        if nfam and juntos and juntos > nfam:
            control.append(_observacion(
                u, "Programas sociales",
                f"JUNTOS: {_fmt_valor(juntos)} familias supera el total de "
                f"familias del CP ({_fmt_valor(nfam)})."))
    if filas_prog:
        series.append(_serie(
            "f1_programas", "Cobertura de programas sociales en el ámbito",
            "barras_h", filas_prog, "clase", "valor", escala="neutra",
            eje_x="Beneficiarios declarados", totales="ninguno",
            unidad="beneficiarios", eje_y="Programa social",
            descripcion="Suma de los beneficiarios declarados por los centros "
                        "poblados, cada uno una sola vez; el rótulo indica "
                        "cuántos CP consignan el dato.",
            nota="F-DS-01, numeral 7. Las unidades difieren por programa "
                 "(familias, personas o instituciones educativas): las barras "
                 "no se suman entre sí.",
            pct_base=None))

    # 7. Gobernanza comunal y presencia institucional (un voto por CP).
    campos_gob = [
        ("f1_junta_vig", "Junta Directiva vigente"),
        ("f1_ronda", "Ronda Campesina activa"),
        ("f1_reglamento", "Reglamento interno"),
        ("f1_comite_rrnn", "Comité de Recursos Naturales"),
        ("f1_pdc", "Plan de Desarrollo Concertado vigente"),
        ("f1_ongs", "ONG operando en la zona"),
        ("f1_agrorural", "Proyectos AGRORURAL"),
        ("f1_prodern", "PRODERN / FONCODES"),
    ]
    filas_gob = _bateria_sino(unidades, campos_gob, form=lambda u: u["form"])
    if filas_gob:
        series.append(_serie(
            "f1_gobernanza", "Capacidades de gobernanza comunal e institucional",
            "apiladas", filas_gob, "cat", "valor", sub="sub",
            orden_sub=["Sí", "No", "No aplica"], colores=_colores_sino(tema),
            unidad="CP", eje_x="Centros poblados", escala="sino",
            totales="filas", etiqueta_total="CP que responden",
            eje_y="Capacidad evaluada",
            descripcion="Presencia efectiva de cada capacidad, contada sobre "
                        "los centros poblados que respondieron (cada CP una "
                        "sola vez).",
            nota="F-DS-01, numerales 3 y 7. Insumo directo de la línea de "
                 "gobernanza y gestión comunitaria del proyecto."))

    # 8. Tenencia de la tierra: la seccion 4 de la ficha se responde para el
    #    BLOQUE ("Tenencia de la tierra relacionada al bloque"), de modo que
    #    se resume con un valor por bloque y nunca por CP ni por ficha.
    bloques_ten = tenencia_por_bloque(registros)
    for b in bloques_ten:
        for d in b["discrepancias"]:
            control.append(_observacion(
                b, "Tenencia del bloque: fichas que no coinciden",
                f"{d['campo']}: se usa {d['usado'] or '(vacío)'} ({d['n_usado']} de "
                f"{d['n_con_dato']} fichas F-DS-01 del bloque con dato); las demás "
                f"declaran {', '.join(d['otros'])} (entre paréntesis, N.° de fichas)."))
    filas_ten = _conteo([_txt(b["form"].get("f1_tenencia")) for b in bloques_ten],
                        orden=FL.L_TENENCIA)
    if filas_ten:
        series.append(_serie(
            "f1_tenencia", "Régimen predominante de tenencia de la tierra en los bloques",
            "barras_h", filas_ten, "clase", "valor", escala="neutra",
            unidad="bloques", eje_x="Bloques",
            descripcion="Determina con quién se suscriben los acuerdos de "
                        "intervención. Se responde para el bloque: cada bloque "
                        "se cuenta una sola vez.",
            nota="F-DS-01, numeral 4 (Tenencia de la tierra relacionada al "
                 "bloque). Si el bloque tiene varias fichas F-DS-01, cuenta el "
                 "régimen más frecuente entre ellas (en empate, el de la más "
                 "reciente)."))

    filas_tit = []
    for b in bloques_ten:
        texto = _txt(b["form"].get("f1_pct_tituladas"))
        pct = _pct_valido(b["form"].get("f1_pct_tituladas"))
        if pct is not None:
            filas_tit.append({"clase": b["etiqueta"], "valor": pct})
        elif texto:
            control.append(_observacion(
                b, "Tierras tituladas",
                f"«{texto}» no es un porcentaje entre 0 y 100; no se grafica."))
    if filas_tit:
        series.append(_serie(
            "f1_tituladas", "Tierras tituladas por bloque (%)",
            "barras_h", filas_tit, "clase", "valor", escala="favorable",
            unidad="%", eje_x="Tierras tituladas (%)", eje_y="Bloque",
            maximo=100, decimales=1, totales="promedio",
            descripcion="A mayor porcentaje titulado, menor riesgo de "
                        "observaciones prediales en el tamizaje del bloque.",
            nota="F-DS-01, numeral 4. Un valor por bloque (el más frecuente "
                 "entre sus fichas; en empate, el de la más reciente); los "
                 "porcentajes nunca se suman."))

    filas_lind = _bateria_sino(bloques_ten, [
        ("f1_superpone", "El bloque se superpone a tierras comunales"),
        ("f1_conf_linderos", "Conflictos de linderos registrados")],
        form=lambda b: b["form"])
    if filas_lind:
        series.append(_serie(
            "f1_tenencia_bloque", "Superposición con tierras comunales y conflictos de linderos",
            "apiladas", filas_lind, "cat", "valor", sub="sub",
            orden_sub=["Sí", "No", "No aplica"], colores=_colores_sino(tema),
            unidad="bloques", eje_x="Bloques", escala="sino", totales="filas",
            etiqueta_total="Bloques que responden", eje_y="Condición del bloque",
            descripcion="Condiciones prediales del bloque que anticipan "
                        "observaciones en el tamizaje predial (paso 5).",
            nota="F-DS-01, numeral 4. Un valor por bloque."))
    if bloques_ten:
        tablas.append(("Tenencia por bloque", [{
            "Bloque": b["bloque"], "Distrito": b["distrito"],
            "Fichas F-DS-01 del bloque": len(b["fichas"]),
            "Régimen predominante de tenencia": _txt(b["form"].get("f1_tenencia")),
            "N.° aprox. de predios individuales": _entero(b["form"].get("f1_n_predios")),
            "Tierras tituladas (%)": _pct_valido(b["form"].get("f1_pct_tituladas")),
            "Conflictos de linderos": _sino(b["form"].get("f1_conf_linderos")),
            "Se superpone a tierras comunales": _sino(b["form"].get("f1_superpone")),
            "Organismo responsable del registro": _txt(b["form"].get("f1_reg_titulacion")),
        } for b in bloques_ten]))

    # 9. Percepciones ordinales del ambito (un voto por CP).
    filas_perc = []
    etiquetas_perc = [("f1_presencia_estatal", "Percepción de presencia estatal"),
                      ("f1_migracion", "Tasa de migración juvenil")]
    for clave, etiqueta in etiquetas_perc:
        for fila in _conteo([_txt(u["form"].get(clave)) for u in unidades],
                            orden=_ORDEN_ABC):
            filas_perc.append({"cat": etiqueta, "sub": fila["clase"],
                               "valor": fila["valor"]})
    if filas_perc:
        presentes, colores = _apiladas(filas_perc, _ORDEN_ABC,
                                       rampa=RAMPA_NEUTRA, descendente=True)
        series.append(_serie(
            "f1_percepciones", "Percepciones declaradas sobre el ámbito",
            "apiladas", filas_perc, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="CP",
            eje_x="Centros poblados", totales="filas",
            etiqueta_total="CP que responden",
            descripcion="Escala Alto / Medio / Bajo de la ficha; cada CP se "
                        "cuenta una sola vez.",
            nota="F-DS-01, numerales 2.5 y 7.1."))

    n_cp = sum(1 for u in unidades if u["tipo"] in ("cp", "compuesto"))
    return {"id": "F-DS-01", "titulo": "F-DS-01 · Diagnóstico socioeconómico",
            "descripcion": (f"{n_fichas} ficha(s) registrada(s) de {len(unidades)} "
                            f"ámbito(s) ({n_cp} centro(s) poblado(s)); cada CP "
                            "se cuenta una sola vez."
                            + (f" {len(excluidas)} ámbito(s) que repiten a un CP "
                               "con ficha propia no se suman (ver Control de "
                               "calidad)." if excluidas else "")),
            "series": series, "tablas": tablas, "control": control,
            "formato_tablas": formato_tablas, "unidades": unidades}


# ── F-DS-02: actores clave ────────────────────────────────────────────────

# Escala ordenada de posicion frente al proyecto: se colorea con la paleta
# divergente, con "Neutral" como punto muerto.
_ORDEN_POSICION = ["A favor del proyecto", "Neutral / Sin posición definida",
                   "Reticente (requiere persuasión)", "En contra del proyecto",
                   "No determinada"]
_COLOR_POSICION = {
    "A favor del proyecto": DIVERGENTE[0],
    "Neutral / Sin posición definida": DIVERGENTE[2],
    "Reticente (requiere persuasión)": DIVERGENTE[3],
    "En contra del proyecto": DIVERGENTE[4],
    "No determinada": SIN_DATO,
}


def _actores_unicos(registros):
    """Actores de las F-DS-02, cada uno una sola vez.

    Un mismo actor figura en varias fichas cuando el CP se registro dos veces
    o cuando su alcance es distrital (la municipalidad, la agencia agraria).
    Se identifica por su nombre y distrito o, si no tiene nombre, por su
    cargo y tipo dentro del mismo CP; se conservan los datos de la ficha mas
    reciente. Devuelve (actores, n_filas_leidas)."""
    regs = sorted(_por_ficha(registros, "F-DS-02"), key=_orden_reciente,
                  reverse=True)
    etiqueta_de = _etiquetas_por_registro(registros)
    actores, vistos, leidas = [], {}, 0
    for reg in regs:
        for fila in _tabla(formulario(reg), "f2_actores"):
            fila = FA.migrar_fila(fila)
            nombre = _col(fila, FA.COL_NOMBRE)
            cargo = _col(fila, FA.COL_CARGO)
            tipo = _col(fila, "Tipo")
            if not (nombre or cargo or tipo):
                continue
            leidas += 1
            if nombre:
                clave = ("n", _clave(nombre), _clave(_distrito(reg)))
            else:
                clave = ("c", _clave(cargo), _clave(tipo), etiqueta_de[id(reg)])
            if clave in vistos:
                vistos[clave]["fichas"] += 1
                continue
            actor = {
                "ambito": etiqueta_de[id(reg)], "nombre": nombre, "cargo": cargo,
                "tipo": tipo, "bloque": _txt(reg.get("bloque_codigo")),
                "distrito": _distrito(reg),
                "influencia": _col(fila, "Influencia"),
                "interes": _col(fila, "Interes", "Interés"),
                "posicion": _col(fila, "Posicion", "Posición"),
                "nivel": _col(fila, "Nivel territorial"),
                "rol": _col(fila, "Rol / Funcion frente al proyecto",
                            "Rol / Función frente al proyecto", "Rol"),
                "fichas": 1,
            }
            vistos[clave] = actor
            actores.append(actor)
    return actores, leidas


def _seccion_actores(registros, tema="claro"):
    regs = _por_ficha(registros, "F-DS-02")
    if not regs:
        return None
    actores, leidas = _actores_unicos(registros)
    if not actores:
        return None
    detalle = [{
        "Centro poblado / ámbito": a["ambito"], "Distrito": a["distrito"],
        "Bloque": a["bloque"],
        "Nombre del actor": a["nombre"], "Cargo": a["cargo"], "Tipo": a["tipo"],
        "Rol frente al proyecto": a["rol"],
        "Influencia": a["influencia"], "Interés": a["interes"],
        "Posición": a["posicion"], "Nivel territorial": a["nivel"],
        "Fichas en que figura": a["fichas"],
    } for a in actores]

    series = []

    # 1. Actores por tipo.
    filas_tipo = _conteo([a["tipo"] for a in actores], orden=FL.L_TIPO_ACTOR)
    if filas_tipo:
        series.append(_serie(
            "f2_tipos", "Actores identificados por tipo de organización",
            "barras_h", filas_tipo, "clase", "valor", escala="neutra",
            unidad="actores", eje_x="N.° de actores",
            descripcion="Composición del mapa de actores del ámbito del bloque.",
            nota="F-DS-02, numeral 3 (Registro de actores identificados)."))

    # 2. Posicion frente al proyecto (escala ordenada, barra divergente).
    filas_pos = _conteo([a["posicion"] for a in actores], orden=_ORDEN_POSICION)
    if filas_pos:
        series.append(_serie(
            "f2_posicion", "Posición de los actores frente al proyecto",
            "barras_h", filas_pos, "clase", "valor",
            colores=_COLOR_POSICION, escala="posicion",
            unidad="actores", eje_x="N.° de actores",
            descripcion="Del verde (a favor) al rojo (en contra); el gris "
                        "reúne a quienes no declararon posición.",
            nota="F-DS-02. Define la estrategia de relacionamiento y el "
                 "orden de acercamiento a cada bloque."))

    # 3. Matriz influencia x interes.
    celdas = {}
    nombres = {}
    for a in actores:
        inf, ints = a["influencia"], a["interes"]
        if not (inf and ints):
            continue
        clave = (inf, ints)
        celdas[clave] = celdas.get(clave, 0) + 1
        etiqueta = a["nombre"] or a["cargo"] or "(sin nombre)"
        if a["nombre"] and a["cargo"]:
            etiqueta = f'{a["nombre"]} ({a["cargo"]})'
        nombres.setdefault(clave, []).append(etiqueta)
    if celdas:
        filas_matriz = [{"cat": inf, "sub": ints, "valor": n,
                         "detalle": ", ".join(nombres[(inf, ints)][:8])}
                        for (inf, ints), n in celdas.items()]
        series.append(_serie(
            "f2_matriz", "Matriz de influencia × interés de los actores",
            "mapa_calor", filas_matriz, "cat", "valor", sub="sub",
            orden_cat=_ORDEN_ABC, orden_sub=_ORDEN_ABC, escala="neutra",
            unidad="actores", eje_x="Interés en el proyecto",
            eje_y="Influencia sobre el territorio",
            descripcion="El cuadrante influencia alta / interés alto concentra "
                        "a los actores con los que debe negociarse primero; "
                        "influencia alta / interés bajo, a quienes hay que "
                        "informar para que no se conviertan en oposición.",
            nota="F-DS-02. Clasificación clásica de mapeo de actores."))

    # 4. Nivel territorial.
    filas_niv = _conteo([a["nivel"] for a in actores], orden=FL.L_NIV_TERR)
    if filas_niv:
        series.append(_serie(
            "f2_nivel", "Actores por nivel territorial de actuación",
            "barras_h", filas_niv, "clase", "valor", escala="neutra",
            unidad="actores", eje_x="N.° de actores",
            descripcion="Permite ver si la gobernanza del bloque descansa en "
                        "actores comunales o requiere articulación con niveles "
                        "distrital, provincial o regional.",
            nota="F-DS-02, numeral 3."))

    repetidos = leidas - len(actores)
    return {"id": "F-DS-02", "titulo": "F-DS-02 · Mapeo de actores clave",
            "descripcion": f"{len(actores)} actor(es) distinto(s) registrado(s) "
                           f"en {len(regs)} ficha(s)"
                           + (f"; {repetidos} repetición(es) del mismo actor "
                              "se cuentan una sola vez." if repetidos else "."),
            "series": series,
            "tablas": [("Actores clave", detalle)]}


# ── F-DS-03: entrevistas a autoridades ────────────────────────────────────

def _rango_edad(valor):
    edad = _num(valor)
    if edad is None:
        return ""
    if edad < 30:
        return "Menos de 30 años"
    if edad < 45:
        return "De 30 a 44 años"
    if edad < 60:
        return "De 45 a 59 años"
    return "60 años o más"


_ORDEN_EDAD = ["Menos de 30 años", "De 30 a 44 años", "De 45 a 59 años",
               "60 años o más"]


def _cargo_principal(texto):
    """Cargo normalizado: el primer rol antes de '/' o ',' ("Teniente
    Gobernador/Agricultor" -> "Teniente Gobernador"); el femenino se rotula
    con el masculino para no partir el mismo cargo en dos barras."""
    principal = re.split(r"\s*[/,]\s*|\s+-\s+", _txt(texto))[0].strip()
    principal = re.sub(r"(?i)\bgobernadora\b", "Gobernador", principal)
    principal = re.sub(r"(?i)\bpresidenta\b", "Presidente", principal)
    return principal


def _conteo_normalizado(valores):
    """_conteo que reune las grafias de un mismo texto (mayusculas, tildes,
    espacios) bajo la grafia mas frecuente."""
    grupos = {}
    for v in valores:
        texto = _txt(v)
        if texto:
            grupos.setdefault(_clave(texto), []).append(texto)
    unificados = []
    for textos in grupos.values():
        conteo = {}
        for t in textos:
            conteo[t] = conteo.get(t, 0) + 1
        rotulo = sorted(conteo, key=lambda t: (-conteo[t], t))[0]
        unificados += [rotulo] * len(textos)
    return _conteo(unificados, agrupar_otros=False)


def _seccion_entrevistas(registros, tema="claro"):
    regs = _por_ficha(registros, "F-DS-03")
    if not regs:
        return None
    etiqueta_de = _etiquetas_por_registro(registros)
    series, detalle = [], []
    filas_edad, generos, duraciones = [], [], []
    for reg in regs:
        f = formulario(reg)
        genero = {"M": "Hombre", "F": "Mujer"}.get(_txt(f.get("f3_genero")),
                                                   _txt(f.get("f3_genero")))
        rango = _rango_edad(f.get("f3_edad"))
        if genero:
            generos.append(genero)
        if rango:
            filas_edad.append({"cat": rango, "sub": genero or "Sin consignar",
                               "valor": 1})
        duracion = _minutos(f.get("f3_dur"))
        if duracion:
            duraciones.append(duracion)
        detalle.append({
            "Centro poblado / ámbito": etiqueta_de[id(reg)],
            "Distrito": _distrito(reg),
            "Bloque": _txt(reg.get("bloque_codigo")),
            "Fecha": _txt(reg.get("fecha_evaluacion")),
            "Entrevistado/a": _txt(f.get("f3_nombre")),
            "Cargo / Rol": _txt(f.get("f3_cargo")),
            "Institución / Organización": _txt(f.get("f3_inst")),
            "Género": genero, "Edad": _num(f.get("f3_edad")),
            "Años en el cargo": _num(f.get("f3_anios")),
            "Duración (min)": duracion,
            "Consiente uso de nombre": _sino(f.get("f3_c_nom")),
            "Consiente fotografías": _sino(f.get("f3_c_foto")),
        })

    if filas_edad:
        # Se agregan las filas unitarias antes de graficar.
        acumulado = {}
        for fila in filas_edad:
            clave = (fila["cat"], fila["sub"])
            acumulado[clave] = acumulado.get(clave, 0) + 1
        filas = [{"cat": c, "sub": s, "valor": v}
                 for (c, s), v in acumulado.items()]
        presentes, colores = _apiladas(filas, ["Hombre", "Mujer", "Sin consignar"],
                                       categorica=True, tema=tema)
        series.append(_serie(
            "f3_perfil", "Perfil de las autoridades y líderes entrevistados",
            "apiladas", filas, "cat", "valor", sub="sub",
            orden_cat=_ORDEN_EDAD, orden_sub=presentes, colores=colores,
            unidad="entrevistas", eje_x="N.° de entrevistas", escala="categorica",
            descripcion="Composición por rango de edad y género: evidencia la "
                        "representatividad del recojo de información.",
            nota="F-DS-03, numeral 2 (Datos del entrevistado/a)."))

    cargos = _conteo_normalizado(
        [_cargo_principal(formulario(r).get("f3_cargo")) for r in regs])
    if cargos:
        series.append(_serie(
            "f3_cargos", "Entrevistas por cargo o rol del informante",
            "barras_h", cargos, "clase", "valor", escala="neutra",
            unidad="entrevistas", eje_x="N.° de entrevistas",
            descripcion="Cargos efectivamente cubiertos por el equipo social.",
            nota="F-DS-03, numeral 2. Se cuenta el cargo principal (el "
                 "primero cuando se declaran varios, p. ej. «Teniente "
                 "Gobernador/Agricultor»), sin distinguir mayúsculas, tildes "
                 "ni género gramatical; el texto completo queda en la tabla "
                 "de entrevistas."))

    filas_cons = _bateria_sino(regs, [
        ("f3_c_nom", "Consiente el uso de su nombre en el informe"),
        ("f3_c_foto", "Consiente la toma de fotografías")])
    if filas_cons:
        series.append(_serie(
            "f3_consentimiento", "Consentimientos otorgados por los entrevistados",
            "apiladas", filas_cons, "cat", "valor", sub="sub",
            orden_sub=["Sí", "No", "No aplica"], colores=_colores_sino(tema),
            unidad="entrevistas", eje_x="N.° de entrevistas", escala="sino",
            descripcion="Condiciona el tratamiento de los datos personales en "
                        "los entregables del estudio.",
            nota="F-DS-03, numeral 2. Ley 29733 de Protección de Datos "
                 "Personales."))

    return {"id": "F-DS-03", "titulo": "F-DS-03 · Entrevistas a autoridades y líderes",
            "descripcion": f"{len(regs)} entrevista(s) registrada(s) (una por "
                           "entrevistado)"
                           + (f"; duración media {sum(duraciones)/len(duraciones):.0f} min."
                              if duraciones else "."),
            "series": series, "tablas": [("Entrevistas", detalle)]}


# ── F-DS-04: talleres participativos ──────────────────────────────────────

def _seccion_talleres(registros, tema="claro"):
    regs = _por_ficha(registros, "F-DS-04")
    if not regs:
        return None
    series, detalle = [], []
    filas_asist, filas_sexo, filas_etaria, metodologias = [], [], [], []
    acuerdos_total = participantes_total = 0

    etiqueta_de = _etiquetas_por_registro(registros)
    vistas = {}
    for reg in regs:
        f = formulario(reg)
        etiqueta = etiqueta_de[id(reg)]
        fecha = _txt(f.get("f4_fecha")) or _txt(reg.get("fecha_evaluacion"))
        if fecha:
            etiqueta = f"{etiqueta} ({fecha})"
        # Dos talleres distintos del mismo CP y dia (otro lugar) no se funden
        # en una barra: se numeran.
        vistas[etiqueta] = vistas.get(etiqueta, 0) + 1
        if vistas[etiqueta] > 1:
            etiqueta = f"{etiqueta} #{vistas[etiqueta]}"
        convocados = _num(f.get("f4_conv_n"))
        hombres, mujeres = _num(f.get("f4_h")), _num(f.get("f4_m"))
        asistentes = _num(f.get("f4_tot"))
        if asistentes is None and (hombres is not None or mujeres is not None):
            asistentes = (hombres or 0) + (mujeres or 0)
        if convocados is not None:
            filas_asist.append({"cat": etiqueta, "sub": "Convocados",
                                "valor": convocados})
        if asistentes is not None:
            filas_asist.append({"cat": etiqueta, "sub": "Asistentes",
                                "valor": asistentes})
        if hombres is not None:
            filas_sexo.append({"cat": etiqueta, "sub": "Hombres", "valor": hombres})
        if mujeres is not None:
            filas_sexo.append({"cat": etiqueta, "sub": "Mujeres", "valor": mujeres})
        jovenes, mayores = _num(f.get("f4_jov")), _num(f.get("f4_am"))
        if asistentes and (jovenes is not None or mayores is not None):
            jovenes, mayores = jovenes or 0, mayores or 0
            filas_etaria += [
                {"cat": etiqueta, "sub": "Jóvenes (menores de 30)", "valor": jovenes},
                {"cat": etiqueta, "sub": "Población adulta", "valor":
                    max(asistentes - jovenes - mayores, 0)},
                {"cat": etiqueta, "sub": "Adultos mayores (más de 60)", "valor": mayores},
            ]
        metodologias += _lista(f.get("f4_metod"))
        n_part = len(_tabla(f, "f4_part"))
        n_acuerdos = len(_tabla(f, "f4_acuerdos"))
        participantes_total += n_part
        acuerdos_total += n_acuerdos
        tasa = (100.0 * asistentes / convocados
                if convocados and asistentes is not None else None)
        detalle.append({
            "Taller": etiqueta,
            "Centro poblado / ámbito": etiqueta_de[id(reg)],
            "Distrito": _distrito(reg), "Bloque": _txt(reg.get("bloque_codigo")),
            "Fecha del taller": fecha,
            "Lugar": _txt(f.get("f4_lugar")),
            "Entidad convocante": _txt(f.get("f4_conv")),
            "Convocados": convocados, "Asistentes": asistentes,
            "Tasa de asistencia (%)": round(tasa, 1) if tasa is not None else None,
            "Hombres": hombres, "Mujeres": mujeres,
            "Jóvenes (<30)": _num(f.get("f4_jov")),
            "Adultos mayores (>60)": _num(f.get("f4_am")),
            "Participantes en lista": n_part, "Acuerdos registrados": n_acuerdos,
            "Idioma de la facilitación": _txt(f.get("f4_idioma")),
        })

    if filas_asist:
        presentes, colores = _apiladas(filas_asist, ["Convocados", "Asistentes"],
                                       categorica=True, tema=tema)
        series.append(_serie(
            "f4_asistencia", "Convocatoria frente a asistencia efectiva por taller",
            "agrupadas", filas_asist, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="personas",
            eje_x="N.° de personas", escala="categorica",
            descripcion="La brecha entre ambas barras mide la capacidad real "
                        "de convocatoria en cada centro poblado.",
            nota="F-DS-04, numeral 2.1 (Asistencia).",
            pct_base=("ref", "Convocados")))

    if filas_sexo:
        presentes, colores = _apiladas(filas_sexo, ["Hombres", "Mujeres"],
                                       categorica=True, tema=tema)
        series.append(_serie(
            "f4_sexo", "Participación por sexo en los talleres",
            "apiladas", filas_sexo, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="personas",
            eje_x="N.° de asistentes", escala="categorica",
            descripcion="Indicador de participación equitativa exigido a los "
                        "procesos participativos del proyecto.",
            nota="F-DS-04, numeral 2.1."))

    if filas_etaria:
        presentes, colores = _apiladas(
            filas_etaria,
            ["Jóvenes (menores de 30)", "Población adulta",
             "Adultos mayores (más de 60)"], rampa=RAMPA_NEUTRA)
        series.append(_serie(
            "f4_etaria", "Composición etaria de los asistentes",
            "apiladas", filas_etaria, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="personas",
            eje_x="N.° de asistentes",
            descripcion="El tramo adulto se obtiene por diferencia con el "
                        "total de asistentes declarado.",
            nota="F-DS-04, numeral 2.1."))

    filas_metod = _conteo(metodologias, agrupar_otros=False)
    if filas_metod:
        series.append(_serie(
            "f4_metodologias", "Metodologías participativas empleadas",
            "barras_h", filas_metod, "clase", "valor", escala="neutra",
            unidad="talleres", eje_x="N.° de talleres",
            descripcion="Cada taller puede emplear más de una metodología.",
            nota="F-DS-04, numeral 2.2.", pct_base=len(regs)))

    filas_acu = [{"clase": d["Taller"], "valor": d["Acuerdos registrados"]}
                 for d in detalle if d["Acuerdos registrados"]]
    if filas_acu:
        series.append(_serie(
            "f4_acuerdos", "Acuerdos y compromisos suscritos por taller",
            "barras_h", filas_acu, "clase", "valor", escala="favorable",
            unidad="acuerdos", eje_x="N.° de acuerdos",
            descripcion="Los acuerdos son el medio de verificación de la "
                        "aceptación comunal del bloque.",
            nota="F-DS-04, numeral 5 (Acuerdos y compromisos)."))

    return {"id": "F-DS-04", "titulo": "F-DS-04 · Talleres participativos",
            "descripcion": f"{len(regs)} taller(es); {participantes_total} "
                           f"participante(s) en lista y {acuerdos_total} acuerdo(s).",
            "series": series, "tablas": [("Talleres participativos", detalle)]}


# ── F-DS-05: conflictos y oportunidades ───────────────────────────────────

_ORDEN_CONFLICTIVIDAD = ["Muy bajo", "Bajo", "Medio", "Alto", "Muy alto"]
_ORDEN_POLARIZACION = ["Muy bajo / Inexistente", "Bajo", "Medio", "Alto", "Muy alto"]
_ORDEN_VIABILIDAD = [
    "Alta — viable para intervencion inmediata",
    "Media — requiere acercamiento reforzado",
    "Baja — requiere mesa de dialogo previa",
    "Muy baja — reevaluar inclusion del bloque",
]
_ORDEN_PLAZO = ["Inmediato (<1 mes)", "Corto (1-3 meses)", "Medio (3-6 meses)",
                "Largo (6-12 meses)", "Requiere >12 meses"]


def _seccion_conflictos(registros, tema="claro"):
    regs = _por_ficha(registros, "F-DS-05")
    if not regs:
        return None
    # La F-DS-05 caracteriza al CP: varias fichas de un mismo CP se
    # consolidan en una (moda de cada escala; conflictos y oportunidades
    # combinados sin repetir el mismo tipo u oportunidad).
    unidades = unidades_por_cp(registros, "F-DS-05")
    series = []
    conflictos, oportunidades, detalle_sintesis = [], [], []

    for u in unidades:
        f = u["form"]
        ambito = u["etiqueta"]
        for fila in _tabla(f, "f5_conflictos"):
            tipo = _col(fila, "Tipo")
            if not tipo:
                continue
            conflictos.append({
                "Centro poblado / ámbito": ambito, "Tipo": tipo,
                "Actores involucrados": _col(fila, "Actores involucrados"),
                "Estado": _col(fila, "Estado"),
                "Antigüedad": _col(fila, "Antiguedad", "Antigüedad"),
                "Descripción / Causa raíz": _col(fila, "Descripcion / Causa raiz",
                                                 "Descripción / Causa raíz"),
                "Impacto potencial en el proyecto":
                    _col(fila, "Impacto potencial en el proyecto"),
            })
        for fila in _tabla(f, "f5_oportunidades"):
            oportunidad = _col(fila, "Oportunidad identificada")
            if not oportunidad:
                continue
            oportunidades.append({
                "Centro poblado / ámbito": ambito,
                "Oportunidad identificada": oportunidad,
                "Actores relacionados": _col(fila, "Actores relacionados"),
                "Tipo": _col(fila, "Tipo (alianza / plataforma / proy.)", "Tipo"),
                "Potencial": _col(fila, "Potencial"),
                "Cómo aprovecharla": _col(fila, "Como aprovecharla",
                                          "Cómo aprovecharla"),
            })
        detalle_sintesis.append({
            "Centro poblado / ámbito": ambito,
            "Distrito": u["distrito"], "Bloque(s)": ", ".join(u["bloques"]),
            "Fichas F-DS-05 del CP": len(u["fichas"]),
            "Nivel global de conflictividad": _txt(f.get("f5_confglob")),
            "Nivel de polarización actual": _txt(f.get("f5_polar")),
            "Viabilidad social preliminar": _txt(f.get("f5_viab")),
            "Plazo estimado de aceptación comunal": _txt(f.get("f5_plazo")),
            "¿Requiere mesa de diálogo?": _sino(f.get("f5_mesa")),
            "Vínculo con el conflicto Río Blanco": _sino(f.get("f5_rb1")),
            "Estrategia de acercamiento recomendada": _txt(f.get("f5_estrategia")),
        })

    # 1. Conflictos por tipo y estado.
    if conflictos:
        acumulado = {}
        for c in conflictos:
            clave = (c["Tipo"], c["Estado"] or "Sin estado consignado")
            acumulado[clave] = acumulado.get(clave, 0) + 1
        filas = [{"cat": t, "sub": e, "valor": n} for (t, e), n in acumulado.items()]
        presentes, colores = _apiladas(filas, FL.L_NIVEL_CONFL, rampa=RAMPA_CRITICA)
        # El estado "Resuelto / Inactivo" no es un agravante: se saca de la
        # rampa critica y se pinta en verde para no leerse como riesgo vivo.
        for clave in list(colores):
            if clave.lower().startswith("resuelto"):
                colores[clave] = RAMPA_FAVORABLE[3]
        series.append(_serie(
            "f5_tipos", "Conflictos identificados por tipo y estado",
            "apiladas", filas, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="conflictos",
            eje_x="N.° de conflictos",
            descripcion="El color va del latente (claro) al activo (oscuro); "
                        "los resueltos se muestran en verde.",
            nota="F-DS-05, numeral 2 (Identificación de conflictos)."))

        filas_ant = _conteo([c["Antigüedad"] for c in conflictos],
                            orden=FL.L_ANTIG_CONFL)
        if filas_ant:
            series.append(_serie(
                "f5_antiguedad", "Antigüedad de los conflictos registrados",
                "barras_h", filas_ant, "clase", "valor", escala="critica",
                orden_cat=FL.L_ANTIG_CONFL,
                unidad="conflictos", eje_x="N.° de conflictos",
                descripcion="Los conflictos históricos exigen mesas de diálogo "
                            "previas; los recientes admiten acercamiento directo.",
                nota="F-DS-05, numeral 2."))

    # 2. Oportunidades por potencial.
    if oportunidades:
        filas_op = _conteo([o["Potencial"] for o in oportunidades],
                           orden=_ORDEN_ABC)
        if filas_op:
            series.append(_serie(
                "f5_oportunidades", "Oportunidades identificadas según potencial",
                "barras_h", filas_op, "clase", "valor", escala="favorable",
                unidad="oportunidades", eje_x="N.° de oportunidades",
                descripcion="Alianzas, plataformas y proyectos con los que el "
                            "bloque puede apalancarse.",
                nota="F-DS-05, numeral 4 (Identificación de oportunidades)."))

    # 3. Sintesis estrategica por ambito (escalas ordenadas).
    escalas = [
        ("f5_confglob", "Nivel global de conflictividad", _ORDEN_CONFLICTIVIDAD, RAMPA_CRITICA),
        ("f5_polar", "Nivel de polarización actual", _ORDEN_POLARIZACION, RAMPA_CRITICA),
        ("f5_viab", "Viabilidad social preliminar", _ORDEN_VIABILIDAD, RAMPA_CRITICA),
        ("f5_plazo", "Plazo estimado de aceptación comunal", _ORDEN_PLAZO, RAMPA_CRITICA),
    ]
    for clave, etiqueta, orden, rampa in escalas:
        filas = _conteo([_txt(u["form"].get(clave)) for u in unidades], orden=orden)
        if not filas:
            continue
        clases = [f["clase"] for f in filas]
        colores = dict(zip(clases, _rampa(clases, rampa)))
        series.append(_serie(
            f"f5_{clave}", etiqueta, "barras_h", filas, "clase", "valor",
            colores=colores, escala="critica", unidad="CP",
            eje_x="Centros poblados",
            descripcion="Escala ordenada de la ficha; a mayor intensidad de "
                        "color, mayor exigencia de gestión social previa. "
                        "Cada CP se cuenta una sola vez.",
            nota="F-DS-05, numeral 5 (Síntesis estratégica). Si un CP tiene "
                 "varias fichas, cuenta el nivel más frecuente entre ellas "
                 "(en empate, el de la más reciente)."))

    filas_rb = _bateria_sino(unidades, [
        ("f5_rb1", "Vínculo histórico con el conflicto Río Blanco"),
        ("f5_rb2", "Participó en la consulta de 2007"),
        ("f5_rb3", "Persiste sentimiento anti-minero fuerte"),
        ("f5_rb4", "Liderazgos activos anti-mineros"),
        ("f5_rb5", "Se diferencia el proyecto IN del contexto minero")],
        form=lambda u: u["form"])
    if filas_rb:
        series.append(_serie(
            "f5_riobranco", "Contexto del conflicto minero Río Blanco",
            "apiladas", filas_rb, "cat", "valor", sub="sub",
            orden_sub=["Sí", "No", "No aplica"], colores=_colores_sino(tema),
            unidad="CP", eje_x="Centros poblados", escala="sino",
            descripcion="Antecedente determinante de la aceptación social en "
                        "Huancabamba y Ayabaca: el proyecto debe diferenciarse "
                        "explícitamente de toda actividad extractiva.",
            nota="F-DS-05, numeral 3 (Contexto específico)."))

    tablas = []
    if conflictos:
        tablas.append(("Conflictos", conflictos))
    if oportunidades:
        tablas.append(("Oportunidades", oportunidades))
    tablas.append(("Síntesis estratégica", detalle_sintesis))

    return {"id": "F-DS-05",
            "titulo": "F-DS-05 · Conflictos socioambientales y oportunidades",
            "descripcion": f"{len(conflictos)} conflicto(s) y "
                           f"{len(oportunidades)} oportunidad(es) en "
                           f"{len(unidades)} centro(s) poblado(s) "
                           f"({len(regs)} ficha(s)).",
            "series": series, "tablas": tablas}


# ── F-DS-06: peligros y cambio climatico ──────────────────────────────────

_ORDEN_FRECUENCIA = FL.FDS06_FRECUENCIA           # MF, F, O, R, SR
_ORDEN_MAGNITUD = FL.FDS06_MAGNITUD               # A, M, B
_ORDEN_INTENSIDAD = FL.FDS06_INTENSIDAD           # Alta, Media, Baja
_ORDEN_TENDENCIA = FL.FDS06_TENDENCIA

# La tendencia no es una magnitud sino un signo: aumentar es adverso,
# disminuir favorable y estable, neutro.
_COLOR_TENDENCIA = {
    "Sube (Aumentando)": DIVERGENTE[4],
    "I (Irregular)": DIVERGENTE[3],
    "E (Estable)": DIVERGENTE[2],
    "Baja (Disminuyendo)": DIVERGENTE[0],
    "NS (No sabe)": SIN_DATO,
}


def _seccion_peligros(registros, tema="claro"):
    regs = _por_ficha(registros, "F-DS-06")
    if not regs:
        return None
    # La percepcion se cuenta por CP ("centros poblados que lo reportan"):
    # varias fichas de un mismo CP se consolidan en una.
    unidades = unidades_por_cp(registros, "F-DS-06")
    series = []
    peligros, cambios = [], []
    prioridades = {}

    for u in unidades:
        f = u["form"]
        ambito = u["etiqueta"]
        for fila in _tabla(f, "f6_peligros"):
            peligro = _col(fila, "Peligro observado")
            if not peligro:
                continue
            peligros.append({
                "Centro poblado / ámbito": ambito, "Peligro observado": peligro,
                "¿Ocurre?": _sino(_col(fila, "¿Ocurre?", "Ocurre")),
                "Frecuencia": _col(fila, "Frecuencia"),
                "Magnitud": _col(fila, "Magnitud"),
                "Tendencia": _col(fila, "Tendencia"),
                "Último evento (año)": _col(fila, "Ultimo evento (año)",
                                            "Último evento (año)"),
                "Principales daños observados":
                    _col(fila, "Principales daños observados"),
            })
        for fila in _tabla(f, "f6_cambios"):
            cambio = _col(fila, "Cambio observado")
            if not cambio:
                continue
            cambios.append({
                "Centro poblado / ámbito": ambito, "Cambio observado": cambio,
                "¿Se percibe?": _sino(_col(fila, "¿Se percibe?", "Se percibe")),
                "Intensidad": _col(fila, "Intensidad"),
                "Año aproximado de inicio": _col(fila, "Año aprox. de inicio"),
                "Impacto en la comunidad / territorio":
                    _col(fila, "Impacto en la comunidad / territorio"),
            })
        # Priorizacion local: el primer peligro pesa 3, el segundo 2, el
        # tercero 1, una sola vez por CP.
        for clave, peso in (("f6_p1", 3), ("f6_p2", 2), ("f6_p3", 1)):
            nombre = _txt(f.get(clave))
            if nombre:
                prioridades[nombre] = prioridades.get(nombre, 0) + peso

    ocurren = [p for p in peligros if p["¿Ocurre?"] == "Sí"]

    # 1. Peligros que ocurren, por frecuencia.
    if ocurren:
        acumulado = {}
        for p in ocurren:
            clave = (p["Peligro observado"],
                     p["Frecuencia"] or "Sin frecuencia consignada")
            acumulado[clave] = acumulado.get(clave, 0) + 1
        filas = [{"cat": c, "sub": s, "valor": v} for (c, s), v in acumulado.items()]
        presentes, colores = _apiladas(filas, _ORDEN_FRECUENCIA,
                                       rampa=RAMPA_CRITICA, descendente=True)
        series.append(_serie(
            "f6_frecuencia", "Peligros que ocurren en el territorio, según frecuencia",
            "apiladas", filas, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="CP",
            eje_x="Centros poblados que lo reportan",
            descripcion="Solo se cuentan los peligros marcados como "
                        "ocurrentes. A mayor oscuridad, mayor frecuencia.",
            nota="F-DS-06, numeral 2. Corrobora en campo el peligro integrado "
                 "modelado por MCA-AHP en la mesolocalización."))

        acumulado = {}
        for p in ocurren:
            clave = (p["Peligro observado"], p["Magnitud"] or "Sin magnitud consignada")
            acumulado[clave] = acumulado.get(clave, 0) + 1
        filas = [{"cat": c, "sub": s, "valor": v} for (c, s), v in acumulado.items()]
        presentes, colores = _apiladas(filas, _ORDEN_MAGNITUD,
                                       rampa=RAMPA_CRITICA, descendente=True)
        series.append(_serie(
            "f6_magnitud", "Magnitud percibida de los peligros ocurrentes",
            "apiladas", filas, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="CP",
            eje_x="Centros poblados que lo reportan",
            descripcion="Magnitud declarada por la población para cada peligro.",
            nota="F-DS-06, numeral 2."))

        acumulado = {}
        for p in ocurren:
            clave = (p["Peligro observado"], p["Tendencia"] or "Sin tendencia consignada")
            acumulado[clave] = acumulado.get(clave, 0) + 1
        filas = [{"cat": c, "sub": s, "valor": v} for (c, s), v in acumulado.items()]
        presentes, _ = _apiladas(filas, _ORDEN_TENDENCIA)
        colores = {s: _COLOR_TENDENCIA.get(s, SIN_DATO) for s in presentes}
        series.append(_serie(
            "f6_tendencia", "Tendencia percibida de los peligros",
            "apiladas", filas, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="CP",
            eje_x="Centros poblados que lo reportan", escala="divergente",
            descripcion="Rojo: el peligro va en aumento. Verde: disminuye. "
                        "Gris: la población no sabe.",
            nota="F-DS-06, numeral 2. Evidencia local de la tendencia a la "
                 "degradación que el proyecto busca revertir."))

    # 2. Cambios climaticos percibidos.
    percibidos = [c for c in cambios if c["¿Se percibe?"] == "Sí"]
    if percibidos:
        acumulado = {}
        for c in percibidos:
            clave = (c["Cambio observado"], c["Intensidad"] or "Sin intensidad consignada")
            acumulado[clave] = acumulado.get(clave, 0) + 1
        filas = [{"cat": c, "sub": s, "valor": v} for (c, s), v in acumulado.items()]
        presentes, colores = _apiladas(filas, _ORDEN_INTENSIDAD,
                                       rampa=RAMPA_CRITICA, descendente=True)
        series.append(_serie(
            "f6_cambios", "Cambios climáticos percibidos en los últimos 10-15 años",
            "apiladas", filas, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="CP",
            eje_x="Centros poblados que lo perciben",
            descripcion="Percepción local contrastable con los escenarios "
                        "climáticos del estudio (MPI-ESM1-2HR, SSP 5-8.5).",
            nota="F-DS-06, numeral 3. Ley 30754, Ley Marco sobre Cambio Climático."))

    # 3. Priorizacion local de peligros.
    if prioridades:
        filas = sorted(({"clase": k, "valor": v} for k, v in prioridades.items()),
                       key=lambda r: -r["valor"])
        series.append(_serie(
            "f6_prioridad", "Priorización local de peligros (índice ponderado)",
            "barras_h", filas, "clase", "valor", escala="critica",
            unidad="pts", eje_x="Índice ponderado",
            descripcion="Índice construido por el aplicativo: el peligro más "
                        "grave señalado en cada centro poblado suma 3 puntos, "
                        "el segundo 2 y el tercero 1.", totales="ninguno",
            nota="F-DS-06, numeral 5. El índice ordena la percepción local; "
                 "no sustituye la evaluación técnica del peligro."))

    # 4. Capacidad de respuesta local.
    filas_cap = _bateria_sino(unidades, [
        ("f6_medidas", "La comunidad ha tomado medidas"),
        ("f6_alerta", "Sistemas de alerta temprana comunitarios"),
        ("f6_saberes", "Saberes tradicionales de predicción"),
        ("f6_apoyo", "Requiere apoyo externo para la adaptación")],
        form=lambda u: u["form"])
    if filas_cap:
        series.append(_serie(
            "f6_capacidad", "Capacidad local de respuesta y adaptación",
            "apiladas", filas_cap, "cat", "valor", sub="sub",
            orden_sub=["Sí", "No", "No aplica"], colores=_colores_sino(tema),
            unidad="CP", eje_x="Centros poblados", escala="sino",
            descripcion="Resiliencia declarada de la población: insumo del "
                        "análisis de vulnerabilidad (fragilidad y resiliencia) "
                        "del AdR-CCC.",
            nota="F-DS-06, numeral 4. Anexo 2 de la Guía General DGPMI-MEF."))

    tablas = []
    if peligros:
        tablas.append(("Peligros percibidos", peligros))
    if cambios:
        tablas.append(("Cambios climáticos", cambios))

    return {"id": "F-DS-06",
            "titulo": "F-DS-06 · Percepción de peligros y cambio climático",
            "descripcion": f"{len(ocurren)} registro(s) de peligros ocurrentes y "
                           f"{len(percibidos)} de cambios percibidos en "
                           f"{len(unidades)} centro(s) poblado(s) "
                           f"({len(regs)} ficha(s)).",
            "series": series, "tablas": tablas}


# ── F-DS-07: disposicion a participar y consentimiento ────────────────────

_PUNTOS_CPI = [
    ("f7_info_anin", "Qué es la ANIN"),
    ("f7_info_objetivo", "Objetivo del proyecto IN Piura"),
    ("f7_info_no_minero", "No es actividad minera ni extractiva"),
    ("f7_info_medidas", "Medidas que podrían implementarse en el predio"),
    ("f7_info_temporalidad", "Temporalidad del proyecto"),
    ("f7_info_voluntaria", "La participación es voluntaria"),
    ("f7_info_actualizada", "Derecho a información actualizada"),
    ("f7_info_confidencialidad", "Tratamiento confidencial de los datos"),
    ("f7_info_preguntas", "Se absolvieron todas las preguntas"),
    ("f7_info_material", "Se entregó material informativo"),
]

_COLOR_DISPOSICION = {
    "Totalmente dispuesto/a": DIVERGENTE[0],
    "Dispuesto/a con condiciones": DIVERGENTE[1],
    "Requiere más información": DIVERGENTE[2],
    "No dispuesto/a por el momento": DIVERGENTE[3],
    "No dispuesto/a rotundamente": DIVERGENTE[4],
}


def _seccion_consentimiento(registros, tema="claro"):
    regs = _por_ficha(registros, "F-DS-07")
    if not regs:
        return None
    series, detalle = [], []
    superficies, documentos = [], []
    etiqueta_de = _etiquetas_por_registro(registros)
    rotulos = {}

    for reg in regs:
        f = formulario(reg)
        superficie = _num(f.get("f7_superficie"))
        if superficie:
            # Una barra por titular: los homonimos o los titulares sin nombre
            # no se suman en una sola barra.
            rotulo = _txt(f.get("f7_nombre")) or "Titular sin nombre"
            rotulos[rotulo] = rotulos.get(rotulo, 0) + 1
            if rotulos[rotulo] > 1:
                dni = re.sub(r"\D", "", _txt(f.get("f7_dni")))
                rotulo = f"{rotulo} ({'DNI ' + dni if dni else '#' + str(rotulos[rotulo])})"
            superficies.append({"clase": rotulo, "valor": superficie})
        documentos += _lista(f.get("f7_docs"))
        informados = sum(1 for clave, _ in _PUNTOS_CPI
                         if _sino(f.get(clave)) == "Sí")
        detalle.append({
            "Centro poblado / ámbito": etiqueta_de[id(reg)],
            "Distrito": _distrito(reg), "Bloque": _txt(reg.get("bloque_codigo")),
            "Titular / Representante": _txt(f.get("f7_nombre")),
            "DNI": _txt(f.get("f7_dni")),
            "Tipo de propietario": _txt(f.get("f7_tipo_prop")),
            "Superficie en el bloque (ha)": superficie,
            "Puntos del CPI explicados (de 10)": informados,
            "Disposición a participar": _txt(f.get("f7_disp")),
            "Autoriza ingreso preliminar": _sino(f.get("f7_aut_ing")),
            "Autoriza fotografías": _sino(f.get("f7_aut_foto")),
            "Autoriza uso de su nombre": _sino(f.get("f7_aut_nom")),
            "Conflictos de linderos": _sino(f.get("f7_linderos")),
            "Residente permanente": _sino(f.get("f7_residente")),
        })

    # 1. Disposicion a participar (escala ordenada favorable <-> desfavorable).
    filas_disp = _conteo([_txt(formulario(r).get("f7_disp")) for r in regs],
                         orden=FL.L_DISPOSICION)
    if filas_disp:
        series.append(_serie(
            "f7_disposicion", "Disposición de los titulares a participar",
            "barras_h", filas_disp, "clase", "valor",
            colores=_COLOR_DISPOSICION, escala="divergente",
            unidad="titulares", eje_x="N.° de titulares",
            descripcion="Del verde (totalmente dispuesto) al rojo (no "
                        "dispuesto); el gris marca a quienes requieren más "
                        "información antes de decidir.",
            nota="F-DS-07, numeral 4. Es el indicador de aceptación predial "
                 "que condiciona la inclusión definitiva del bloque."))

    # 2. Consentimiento previo informado: cobertura de los 10 puntos.
    filas_cpi = _bateria_sino(regs, _PUNTOS_CPI)
    if filas_cpi:
        series.append(_serie(
            "f7_cpi", "Consentimiento previo informado: puntos efectivamente explicados",
            "apiladas", filas_cpi, "cat", "valor", sub="sub",
            orden_sub=["Sí", "No", "No aplica"], colores=_colores_sino(tema),
            unidad="titulares", eje_x="N.° de titulares", escala="sino",
            descripcion="Un punto en rojo es una brecha del proceso de "
                        "consentimiento que debe subsanarse antes de "
                        "formalizar el acuerdo.",
            nota="F-DS-07, numeral 3. Convenio 169 de la OIT y estándares "
                 "de salvaguardas sociales."))

    # 3. Autorizaciones otorgadas.
    filas_aut = _bateria_sino(regs, [
        ("f7_aut_ing", "Autoriza el ingreso preliminar al predio"),
        ("f7_aut_foto", "Autoriza fotografías del predio"),
        ("f7_aut_nom", "Autoriza el uso de su nombre"),
        ("f7_consultar", "Requiere consultar con familia o asamblea"),
        ("f7_linderos", "Declara conflictos de linderos"),
        ("f7_residente", "Es residente permanente")])
    if filas_aut:
        series.append(_serie(
            "f7_autorizaciones", "Autorizaciones y condiciones declaradas por los titulares",
            "apiladas", filas_aut, "cat", "valor", sub="sub",
            orden_sub=["Sí", "No", "No aplica"], colores=_colores_sino(tema),
            unidad="titulares", eje_x="N.° de titulares", escala="sino",
            descripcion="Las tres primeras habilitan el trabajo de campo; las "
                        "tres últimas describen condiciones del predio.",
            nota="F-DS-07, numerales 2 y 4."))

    # 4. Tipo de propietario y documentacion de tenencia.
    filas_tipo = _conteo([_txt(formulario(r).get("f7_tipo_prop")) for r in regs],
                         orden=FL.L_TIPO_PROPIETARIO)
    if filas_tipo:
        series.append(_serie(
            "f7_tipo_prop", "Titulares y representantes por tipo",
            "barras_h", filas_tipo, "clase", "valor", escala="neutra",
            unidad="titulares", eje_x="N.° de titulares",
            descripcion="Define el instrumento jurídico aplicable a cada "
                        "acuerdo de intervención.",
            nota="F-DS-07, numeral 2.1."))

    filas_doc = _conteo(documentos, agrupar_otros=False)
    if filas_doc:
        series.append(_serie(
            "f7_documentos", "Documentación de tenencia disponible",
            "barras_h", filas_doc, "clase", "valor", escala="neutra",
            unidad="menciones", eje_x="N.° de menciones",
            descripcion="Un mismo titular puede presentar más de un documento.",
            nota="F-DS-07, numeral 2.2. Insumo del tamizaje predial del bloque.",
            pct_base=len(regs)))

    if superficies:
        total_ha = sum(s["valor"] for s in superficies)
        series.append(_serie(
            "f7_superficies", "Superficie predial comprometida por titular (ha)",
            "barras_h", sorted(superficies, key=lambda r: -r["valor"]),
            "clase", "valor", escala="favorable", unidad="ha",
            eje_x="Superficie (ha)", decimales=2,
            descripcion=f"Suma declarada: {total_ha:,.2f} ha dentro del bloque."
                        .replace(",", " "),
            nota="F-DS-07, numeral 2. Superficie declarada por el titular; no "
                 "sustituye la medición predial del expediente."))

    return {"id": "F-DS-07",
            "titulo": "F-DS-07 · Disposición a participar y consentimiento previo",
            "descripcion": f"{len(regs)} titular(es) / representante(s) registrado(s).",
            "series": series, "tablas": [("Titulares y consentimiento", detalle)]}


# ══════════════════════════════════════════════════════════════════════════
# COBERTURA DEL DIAGNOSTICO Y ARMADO DEL INFORME
# ══════════════════════════════════════════════════════════════════════════

def catalogo_cp(bloques):
    """Centros poblados del catalogo INEI asociados a un conjunto de bloques,
    cada uno una sola vez aunque figure en varios bloques.

    Devuelve [{"nombre", "bloques", "poblacion"}]. Dos entradas son el mismo
    CP cuando comparten ubicacion; una entrada sin coordenadas se une a la
    unica entrada con coordenadas del mismo nombre, si la hay.
    """
    entradas = []
    for bloque in dict.fromkeys(_txt(b) for b in bloques or [] if _txt(b)):
        for fila in _catalogo_bloque(bloque).get("demografia") or []:
            nombre = _txt(fila.get("centro_poblado"))
            if not nombre:
                continue
            este, norte = fila.get("utm_este") or 0, fila.get("utm_norte") or 0
            xy = (round(float(este)), round(float(norte))) if este and norte else None
            entradas.append((bloque, nombre, xy, fila.get("poblacion_total") or 0))
    con_xy = {}
    for _b, nombre, xy, _p in entradas:
        if xy:
            con_xy.setdefault(_clave(nombre), set()).add(xy)
    unicos = {}
    for bloque, nombre, xy, pob in entradas:
        if xy is None and len(con_xy.get(_clave(nombre), ())) == 1:
            xy = next(iter(con_xy[_clave(nombre)]))
        clave = ("xy",) + xy if xy else ("n", _clave(nombre))
        cp = unicos.setdefault(clave, {"nombre": nombre, "bloques": [],
                                       "poblacion": 0})
        if bloque not in cp["bloques"]:
            cp["bloques"].append(bloque)
        cp["poblacion"] = max(cp["poblacion"], pob or 0)
    return list(unicos.values())


def _nombres_registrados(registros):
    """{bloque: {claves de los CP con alguna ficha}} (partes de 'A / B')."""
    por_bloque = {}
    for r in registros:
        bloque = _txt(r.get("bloque_codigo"))
        for parte in _partes_cp(r.get("centro_poblado")) or [_comunidad(r)]:
            if _clave(parte):
                por_bloque.setdefault(bloque, set()).add(_clave(parte))
    return por_bloque


def _cobertura_catalogo(registros, catalogo):
    """(cubiertos, pendientes): CP del catalogo con y sin ficha social.

    Un CP del catalogo esta cubierto si alguna ficha de uno de sus bloques lo
    nombra (sin distinguir tildes ni mayusculas)."""
    registrados = _nombres_registrados(registros)
    todos = set().union(*registrados.values()) if registrados else set()
    cubiertos, pendientes = [], []
    for cp in catalogo:
        clave = _clave(cp["nombre"])
        bloques = cp.get("bloques") or []
        ok = (any(clave in registrados.get(b, ()) for b in bloques)
              if bloques else clave in todos)
        (cubiertos if ok else pendientes).append(cp)
    return cubiertos, pendientes


def _seccion_cobertura(registros, catalogo=(), tema="claro"):
    """Avance del levantamiento: que fichas hay y en que centros poblados."""
    if not registros:
        return None
    series = []

    # 1. Fichas registradas por tipo.
    conteo_ficha = {f: len(_por_ficha(registros, f)) for f in FICHAS_DS}
    filas = [{"clase": f"{f} · {FICHAS_DS_TITULOS[f]}", "valor": conteo_ficha[f]}
             for f in FICHAS_DS]
    series.append(_serie(
        "cob_fichas", "Fichas sociales registradas por tipo",
        "barras_h", filas, "clase", "valor", escala="favorable",
        unidad="fichas", eje_x="N.° de fichas registradas", orden_cat=
        [f["clase"] for f in filas],
        descripcion="Las siete fichas de la Plantilla V4. Una barra vacía "
                    "señala un instrumento aún no aplicado en el bloque.",
        nota="Conteo tras descartar reediciones y copias idénticas de una "
             "misma ficha."))

    # 2. Fichas por centro poblado (un CP escrito de dos formas, o asociado a
    #    dos bloques, es una sola barra).
    claves, etiquetas = _agrupar_ambitos(registros)
    acumulado = {}
    for i, reg in enumerate(registros):
        clave = (etiquetas[claves[i]], reg.get("ficha", "") or "Sin ficha")
        acumulado[clave] = acumulado.get(clave, 0) + 1
    filas_cp = [{"cat": ambito, "sub": ficha, "valor": n}
                for (ambito, ficha), n in acumulado.items()]
    if filas_cp:
        presentes, colores = _apiladas(filas_cp, FICHAS_DS, categorica=True,
                                       tema=tema)
        series.append(_serie(
            "cob_cp", "Cobertura del diagnóstico social por centro poblado",
            "apiladas", filas_cp, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="fichas",
            eje_x="N.° de fichas", eje_y="Centro poblado", escala="categorica",
            totales="ambos", etiqueta_total="Total de fichas",
            descripcion="Cada color es una ficha distinta: permite ver de un "
                        "vistazo qué centro poblado quedó incompleto y cuál "
                        "acumula fichas repetidas.",
            nota="Ámbito declarado en cada ficha (centro poblado, comunidad "
                 "campesina o, a falta de ambos, el propio bloque)."))

    # 3. Centros poblados del catalogo INEI con y sin ficha.
    if catalogo:
        cubiertos, pendientes = _cobertura_catalogo(registros, catalogo)
        filas_pend = [
            {"clase": "Con al menos una ficha social", "valor": len(cubiertos)},
            {"clase": "Sin ficha social registrada", "valor": len(pendientes)},
        ]
        nota = ("El cotejo es por nombre dentro de los bloques del CP, sin "
                "distinguir tildes ni mayúsculas.")
        if pendientes:
            nota += " Sin ficha: " + ", ".join(
                f"{cp['nombre']} ({', '.join(cp['bloques'])})"
                if cp.get("bloques") else cp["nombre"]
                for cp in pendientes) + "."
        series.append(_serie(
            "cob_catalogo", "Centros poblados del catálogo oficial cubiertos",
            "barras_h", filas_pend, "clase", "valor",
            colores={"Con al menos una ficha social": TINTAS[tema]["si"],
                     "Sin ficha social registrada": TINTAS[tema]["no"]},
            escala="sino", unidad="CP", eje_x="Centros poblados",
            descripcion=f"{len(catalogo)} centro(s) poblado(s) en la relación "
                        "oficial (INEI) de los bloques del ámbito, cada uno "
                        "una sola vez aunque figure en varios bloques.",
            nota=nota))

    return {"id": "COBERTURA", "titulo": "Cobertura del diagnóstico social",
            "descripcion": f"{len(registros)} ficha(s) vigente(s).",
            "series": series, "tablas": []}


_CONSTRUCTORES = [
    _seccion_socioeconomica, _seccion_actores, _seccion_entrevistas,
    _seccion_talleres, _seccion_conflictos, _seccion_peligros,
    _seccion_consentimiento,
]


def _poblacion_unidad(form):
    """Poblacion de un CP: el total declarado o, si falta, hombres + mujeres."""
    total = _entero(form.get("f1_pob_t"))
    if total is not None:
        return total
    hombres, mujeres = _entero(form.get("f1_pob_h")), _entero(form.get("f1_pob_m"))
    if hombres is None and mujeres is None:
        return None
    return (hombres or 0) + (mujeres or 0)


def _cp_cubiertos(registros):
    """Centros poblados distintos con al menos una ficha: los ambitos 'A / B'
    cuentan cada CP y no se cuentan comunidades ni fichas sin CP."""
    claves, _etiquetas = _agrupar_ambitos(registros)
    cubiertos = set()
    for i, reg in enumerate(registros):
        ident = _identidad(reg)
        if ident["tipo"] == "cp":
            cubiertos.add(claves[i])
        elif ident["tipo"] == "compuesto":
            for parte in _partes_cp(reg.get("centro_poblado")):
                cubiertos.add(("parte", _clave(parte), _clave(ident["distrito"])))
    # Una parte de un ambito compuesto que tambien tiene ficha propia no
    # cuenta dos veces.
    propios = {(_clave(_txt(r.get("centro_poblado"))), _clave(_distrito(r)))
               for r in registros if _identidad(r)["tipo"] == "cp"}
    return {c for c in cubiertos
            if not (isinstance(c, tuple) and c[0] == "parte" and c[1:] in propios)}


def _metricas(registros, datos_cp, secciones, catalogo=()):
    """Cifras de cabecera del informe (fila de indicadores del aplicativo).

    Todas las cifras de centros poblados se calculan sobre CP unicos (ver
    unidades_por_cp): un CP con varias fichas, o asociado a varios bloques,
    suma una sola vez."""
    unidades, _excluidas = _unidades_fds01(registros)
    poblaciones = [_poblacion_unidad(u["form"]) for u in unidades]
    con_pob = [p for p in poblaciones if p is not None]
    poblacion = sum(con_pob)
    familias = sum(_entero(u["form"].get("f1_nfam")) or 0 for u in unidades)
    if con_pob:
        detalle_pob = (f"F-DS-01: {len(con_pob)} CP con dato"
                       + (f" · {int(familias):,} familias".replace(",", " ")
                          if familias else ""))
    else:
        poblacion = (datos_cp or {}).get("poblacion_total", 0) or \
            sum(cp.get("poblacion") or 0 for cp in catalogo or [])
        detalle_pob = "INEI (catálogo): sin dato de campo en F-DS-01"

    actores, _leidas = _actores_unicos(registros)
    a_favor = sum(1 for a in actores
                  if a["posicion"].lower().startswith("a favor"))

    conflictos = activos = 0
    for u in unidades_por_cp(registros, "F-DS-05"):
        for fila in _tabla(u["form"], "f5_conflictos"):
            if not _col(fila, "Tipo"):
                continue
            conflictos += 1
            estado = _col(fila, "Estado").lower()
            if estado.startswith("activo") or estado.startswith("en escalada"):
                activos += 1

    regs07 = _por_ficha(registros, "F-DS-07")
    dispuestos = sum(1 for r in regs07
                     if _txt(formulario(r).get("f7_disp")).lower()
                     .startswith(("totalmente dispuesto", "dispuesto")))
    superficie = sum(_num(formulario(r).get("f7_superficie")) or 0 for r in regs07)

    asistentes = 0
    for r in _por_ficha(registros, "F-DS-04"):
        f = formulario(r)
        total = _num(f.get("f4_tot"))
        if total is None:
            total = (_num(f.get("f4_h")) or 0) + (_num(f.get("f4_m")) or 0)
        asistentes += total or 0

    cubiertos = _cp_cubiertos(registros)
    fichas_con_datos = {r.get("ficha", "") for r in registros}
    detalle_cp = "con al menos una ficha (cada CP una vez)"
    if catalogo:
        en_catalogo, _pend = _cobertura_catalogo(registros, catalogo)
        detalle_cp = (f"{len(en_catalogo)} de {len(catalogo)} del catálogo INEI "
                      "con ficha")

    metricas = [
        {"etiqueta": "Fichas sociales vigentes", "valor": f"{len(registros)}",
         "detalle": f"{len(fichas_con_datos)} de 7 tipos aplicados"},
        {"etiqueta": "Centros poblados cubiertos", "valor": f"{len(cubiertos)}",
         "detalle": detalle_cp},
        {"etiqueta": "Población del ámbito",
         "valor": f"{int(poblacion):,}".replace(",", " "),
         "detalle": detalle_pob},
        {"etiqueta": "Asistentes a talleres", "valor": f"{int(asistentes):,}".replace(",", " "),
         "detalle": f"{len(_por_ficha(registros, 'F-DS-04'))} taller(es)"},
        {"etiqueta": "Actores mapeados", "valor": f"{len(actores)}",
         "detalle": f"{a_favor} declaradamente a favor",
         "tono": "favorable" if a_favor else "neutro"},
        {"etiqueta": "Conflictos identificados", "valor": f"{conflictos}",
         "detalle": f"{activos} activo(s) o en escalada",
         "tono": "critico" if activos else "favorable"},
        {"etiqueta": "Titulares dispuestos a participar",
         "valor": f"{dispuestos} / {len(regs07)}" if regs07 else "s/d",
         "detalle": "F-DS-07 (consentimiento previo informado)",
         "tono": "favorable" if regs07 and dispuestos >= len(regs07) / 2 else "critico"},
        {"etiqueta": "Superficie predial comprometida",
         "valor": f"{superficie:,.2f} ha".replace(",", " ") if superficie else "s/d",
         "detalle": "declarada por los titulares en F-DS-07"},
    ]
    return metricas


def _control_duplicados_probables(unidades):
    """Fichas de CP distintos con los mismos datos de identificacion: el mismo
    presidente de junta, o la misma poblacion y familias, en dos CP suele
    ser la ficha de un CP registrada con el nombre de otro."""
    observaciones = []
    por_presidente, por_cifras = {}, {}
    for u in unidades:
        f = u["form"]
        presidente = _clave(f.get("f1_pres_junta"))
        if len(presidente) > 6:
            por_presidente.setdefault(presidente, []).append(u)
        cifras = (_entero(f.get("f1_pob_t")), _entero(f.get("f1_nfam")))
        if all(c for c in cifras):
            por_cifras.setdefault(cifras, []).append(u)
    avisados = set()
    for motivo, grupos in (("el mismo presidente de junta", por_presidente),
                           ("la misma población total y N.° de familias", por_cifras)):
        for grupo in grupos.values():
            if len(grupo) < 2:
                continue
            nombres = ", ".join(sorted(g["etiqueta"] for g in grupo))
            for u in grupo:
                clave = (u["clave"], nombres)
                if clave in avisados:
                    continue
                avisados.add(clave)
                observaciones.append(_observacion(
                    u, "Posible ficha duplicada",
                    f"Comparte {motivo} con: {nombres}. Verifique que la "
                    "ficha corresponda a este centro poblado y no a otro."))
    return observaciones


def _control_calidad(registros, secciones):
    """Observaciones de calidad de datos del informe, para depurar la base."""
    control = []
    for seccion in secciones:
        control += seccion.get("control") or []
    unidades = unidades_por_cp(registros, "F-DS-01")
    control += _control_duplicados_probables(unidades)
    return control


def _fichas_fds01(registros):
    """Detalle ficha por ficha de la F-DS-01 con el CP al que se asigno y si
    es la ficha de referencia del CP (la que manda en los datos
    demograficos). Es la tabla que permite depurar los duplicados."""
    filas = []
    for u in unidades_por_cp(registros, "F-DS-01"):
        for reg in u["fichas"]:
            f = formulario(reg)
            filas.append({
                "Centro poblado (consolidado)": u["etiqueta"],
                "Centro poblado (ficha)": _txt(reg.get("centro_poblado")),
                "Comunidad campesina": _comunidad(reg),
                "Distrito": _distrito(reg),
                "Bloque": _txt(reg.get("bloque_codigo")),
                "Fecha": _txt(reg.get("fecha_evaluacion")),
                "Responsable": _txt(reg.get("evaluador")),
                "Entrevistado": _txt(reg.get("nombre_entrevistado")),
                "Ficha de referencia del CP": "Sí" if reg is u["referencia"] else "No",
                "Se suma en el análisis": "Sí" if u.get("suma", True) else
                f"No: {u.get('motivo_excluido', '')}",
                "Familias / viviendas": _entero(f.get("f1_nfam")),
                "Población total (hab.)": _entero(f.get("f1_pob_t")),
                "Hombres": _entero(f.get("f1_pob_h")),
                "Mujeres": _entero(f.get("f1_pob_m")),
                "Cobertura de agua (texto)": _txt(f.get("f1_agua_cob")),
                "Cobertura de energía (texto)": _txt(f.get("f1_energia_cob")),
                "ID": reg.get("id"),
            })
    return filas


def _seccion_control(registros, secciones):
    """Seccion sin graficos: control de calidad y detalle de las fichas."""
    control = _control_calidad(registros, secciones)
    fichas = _fichas_fds01(registros)
    tablas = []
    if control:
        tablas.append(("Control de calidad", control))
    if fichas:
        tablas.append(("Fichas F-DS-01 por CP", fichas))
    if not tablas:
        return None
    return {"id": "CONTROL", "titulo": "Control de calidad de los datos",
            "descripcion": f"{len(control)} observación(es) para revisar en "
                           "las fichas.",
            "series": [], "tablas": tablas, "control": control}


def _avisos_consolidacion(registros):
    unidades = unidades_por_cp(registros, "F-DS-01")
    repetidos = [u for u in unidades if len(u["fichas"]) > 1]
    if not repetidos:
        return []
    n = sum(len(u["fichas"]) for u in repetidos)
    return [f"{len(repetidos)} centro(s) poblado(s) tienen más de una ficha "
            f"F-DS-01 ({n} fichas): cada CP se cuenta una sola vez y sus "
            "datos son el valor más frecuente entre sus fichas (en empate, el "
            "de la más reciente). Las diferencias entre fichas se listan en "
            "«Control de calidad de los datos»."]


def indicadores_bloque(bloque, registros, datos_cp=None, tema="claro"):
    """Informe analitico del Diagnostico Social de UN bloque.

    `bloque`   dict del bloque (database.obtener_bloque_por_id).
    `registros` fichas sociales del bloque, tal como las devuelve la BD.
    `datos_cp` entrada del catalogo de centros poblados del bloque.

    Devuelve un dict con metricas de cabecera, secciones (cada una con sus
    series graficables y sus tablas de respaldo) y avisos.
    """
    bloque = bloque or {}
    datos_cp = datos_cp or {}
    registros = deduplicar(registros)
    codigo = _txt(bloque.get("codigo")) or _txt(datos_cp.get("codigo"))
    centros = [c for c in (datos_cp.get("centros_poblados") or []) if _txt(c)]
    poblaciones = {_clave(d.get("centro_poblado")): d.get("poblacion_total") or 0
                   for d in datos_cp.get("demografia") or []}
    catalogo = [{"nombre": c, "bloques": [codigo] if codigo else [],
                 "poblacion": poblaciones.get(_clave(c), 0)} for c in centros]

    secciones, sin_series = [], []
    cobertura = _seccion_cobertura(registros, catalogo, tema)
    if cobertura:
        secciones.append(cobertura)
    for constructor in _CONSTRUCTORES:
        seccion = constructor(registros, tema)
        if seccion and seccion["series"]:
            secciones.append(seccion)
        elif seccion:
            sin_series.append(seccion)

    avisos = []
    faltantes = [f for f in FICHAS_DS if not _por_ficha(registros, f)]
    if faltantes:
        avisos.append("Sin registros de " + ", ".join(faltantes) +
                      ": las secciones correspondientes no se grafican.")
    if not registros:
        avisos.append("El bloque no tiene fichas sociales registradas.")
    avisos += _avisos_consolidacion(registros)
    # El control recoge tambien las secciones que no llegaron a graficar
    # nada (p. ej. solo porcentajes ilegibles): sus observaciones cuentan.
    control = _seccion_control(registros, secciones + sin_series)

    return {
        "alcance": "bloque",
        "codigo": codigo,
        "bloque": bloque,
        "centros_poblados": centros,
        "n_registros": len(registros),
        "metricas": _metricas(registros, datos_cp, secciones, catalogo),
        "secciones": secciones,
        "control": control,
        "avisos": avisos,
        "registros": registros,
        "generado": datetime.now(),
        "tema": tema,
    }


def indicadores_consolidado(registros, etiqueta="", tema="claro",
                            bloques_ambito=None):
    """Informe analitico de un conjunto de bloques (todo el ambito o un filtro).

    Agrega ademas la distribucion de fichas por bloque, provincia y distrito,
    que es lo que distingue la mirada consolidada de la de un solo bloque.

    `bloques_ambito`: codigos de TODOS los bloques del filtro, tengan o no
    fichas. Con ellos se arma la relacion de CP del catalogo INEI del ambito
    (p. ej. los 11 caserios del distrito), de modo que se vea cuantos faltan.
    Si no se entrega, se usan los bloques que tienen fichas.
    """
    registros = deduplicar(registros)
    bloques = sorted({_txt(r.get("bloque_codigo")) for r in registros
                      if _txt(r.get("bloque_codigo"))})
    ambito = sorted({_txt(b) for b in (bloques_ambito or []) if _txt(b)}) or bloques
    catalogo = catalogo_cp(ambito)
    secciones = []
    cobertura = _seccion_cobertura(registros, catalogo, tema)
    if cobertura:
        # En la mirada consolidada interesa el reparto por bloque y por
        # distrito, no el centro poblado individual. El distrito y la
        # provincia son los del bloque, los mismos del filtro.
        for campo, titulo, id_ in (
                (("bloque_codigo",), "Fichas sociales por bloque", "cons_bloque"),
                (("bloque_distrito", "distrito"), "Fichas sociales por distrito del bloque",
                 "cons_distrito"),
                (("bloque_provincia", "provincia"), "Fichas sociales por provincia del bloque",
                 "cons_provincia")):
            filas = _conteo([next((_txt(r.get(c)) for c in campo if _txt(r.get(c))), "")
                             for r in registros], agrupar_otros=False)
            if filas:
                cobertura["series"].append(_serie(
                    id_, titulo, "barras_h", filas, "clase", "valor",
                    escala="neutra", unidad="fichas", eje_x="N.° de fichas",
                    descripcion="Distribución del levantamiento social en el "
                                "ámbito seleccionado.",
                    nota="Cabecera de cada ficha registrada y bloque al que "
                         "pertenece."))
        secciones.append(cobertura)
    sin_series = []
    for constructor in _CONSTRUCTORES:
        seccion = constructor(registros, tema)
        if seccion and seccion["series"]:
            secciones.append(seccion)
        elif seccion:
            sin_series.append(seccion)

    metricas = _metricas(registros, {}, secciones, catalogo)
    detalle = etiqueta or "ámbito seleccionado"
    if bloques_ambito:
        detalle = f"de {len(ambito)} bloque(s) del ámbito · {detalle}"
    metricas.insert(0, {"etiqueta": "Bloques con diagnóstico social",
                        "valor": f"{len(bloques)}", "detalle": detalle})
    avisos = ([] if registros else
              ["No hay fichas sociales en el ámbito seleccionado."])
    avisos += _avisos_consolidacion(registros)
    return {
        "alcance": "consolidado",
        "codigo": etiqueta or f"{len(bloques)} bloques",
        "bloque": {}, "centros_poblados": [cp["nombre"] for cp in catalogo],
        "n_registros": len(registros),
        "metricas": metricas, "secciones": secciones,
        "control": _seccion_control(registros, secciones + sin_series),
        "avisos": avisos,
        "registros": registros, "bloques": bloques,
        "bloques_ambito": ambito,
        "generado": datetime.now(), "tema": tema,
    }


# ══════════════════════════════════════════════════════════════════════════
# GRAFICOS INTERACTIVOS (Altair / Vega-Lite)
# ══════════════════════════════════════════════════════════════════════════
# Altair viaja con Streamlit; se importa de forma diferida para que el resto
# del modulo (lectura de fichas y Excel) siga funcionando sin el.

def _alt():
    import altair as alt
    return alt


def _df(filas):
    import pandas as pd
    return pd.DataFrame(filas)


def _escala_valor(serie):
    """Rampa secuencial que corresponde al signo de la magnitud graficada."""
    return {"critica": RAMPA_CRITICA, "favorable": RAMPA_FAVORABLE,
            "neutra": RAMPA_NEUTRA}.get(serie.get("escala"), RAMPA_NEUTRA)


def _formato(serie):
    return f",.{serie.get('decimales', 0)}f"


def _eje_y(alt, serie, campo, orden):
    """Eje de categorias: sin regla ni marcas, y con aire antes de la barra."""
    return alt.Y(f"{campo}:N", sort=orden, title=None,
                 axis=alt.Axis(labelLimit=320, labelFontSize=11,
                               labelPadding=10, domain=False, ticks=False))


def grafico_altair(serie, tema="claro", altura=None):
    """Construye el grafico interactivo de una serie.

    Devuelve un objeto Altair listo para `st.altair_chart(..., use_container_width=True)`.
    Todos los graficos llevan tooltip; las barras llevan ademas el valor
    rotulado al extremo, que es lo que sostiene la lectura cuando el color
    por si solo no alcanza el contraste minimo sobre el fondo.
    """
    alt = _alt()
    t = TINTAS.get(tema, TINTAS["claro"])
    filas = serie.get("filas") or []
    if not filas:
        return None
    df, categorias, subclases = _datos_grafico(serie)
    if df.empty:
        return None
    forma = serie.get("forma", "barras_h")
    base = alt.Chart(df)
    if forma == "mapa_calor":
        grafico = _mapa_calor(alt, base, serie, t)
        alto = altura or 62 * max(len(serie.get("orden_cat") or categorias), 3) + 40
    elif forma in ("apiladas", "agrupadas"):
        grafico = _barras_multiples(alt, base, serie, t, forma, categorias,
                                    subclases)
        alto = altura or _alto_barras(len(categorias),
                                      agrupadas=(forma == "agrupadas"),
                                      n_sub=len(subclases))
    else:
        grafico = _barras_simples(alt, base, df, serie, t, categorias)
        alto = altura or _alto_barras(len(categorias))

    return (grafico
            # "fit-x" ajusta solo el ancho al contenedor. Con el "fit" que
            # Streamlit aplica por defecto, la leyenda inferior se descuenta
            # tambien del alto y las barras apiladas quedan sin dibujar.
            .properties(autosize=alt.AutoSizeParams(type="fit-x",
                                                    contains="padding"),
                        height=alto, title=alt.TitleParams(
                text=serie.get("titulo", ""),
                subtitle=_envolver(serie.get("descripcion", "")),
                anchor="start", fontSize=14, font=FUENTE, color=t["tinta"],
                subtitleFontSize=11, subtitleFont=FUENTE,
                subtitleColor=t["tinta_2"], offset=10))
            .configure_view(stroke=None)
            .configure_axis(grid=True, gridColor=t["rejilla"], gridWidth=1,
                            domainColor=t["rejilla"], tickColor=t["rejilla"],
                            labelColor=t["tinta_2"], titleColor=t["tinta_2"],
                            labelFont=FUENTE, titleFont=FUENTE,
                            titleFontSize=11, titleFontWeight="normal")
            .configure_legend(labelColor=t["tinta_2"], titleColor=t["tinta_2"],
                              labelFont=FUENTE, titleFont=FUENTE,
                              labelFontSize=11, titleFontSize=11,
                              symbolType="square", symbolSize=110,
                              orient="bottom", direction="horizontal",
                              columns=3, labelLimit=280)
            .configure_title(font=FUENTE))


def _datos_grafico(serie):
    """DataFrame acumulado del grafico, con el orden de apilado resuelto.

    Devuelve tambien el orden de categorias y subclases para que el eje, la
    leyenda y el apilado usen exactamente el mismo que el Excel y el PDF.
    """
    import pandas as pd
    cat, val, sub = serie["cat"], serie["val"], serie.get("sub")
    acumulado, extras = _acumular(serie)
    categorias = _orden_categorias(serie, acumulado)
    subclases = []
    if sub:
        subclases = [x for x in (serie.get("orden_sub") or [])
                     if any(k[1] == x for k in acumulado)]
        subclases += sorted({k[1] for k in acumulado} - set(subclases))
    if serie.get("forma") == "mapa_calor":
        for c in (serie.get("orden_cat") or categorias):
            for sc in (serie.get("orden_sub") or subclases):
                acumulado.setdefault((c, sc), 0)
        categorias = serie.get("orden_cat") or categorias
    filas = []
    for (c, sc), valor in acumulado.items():
        fila = {cat: c, val: valor}
        if sub:
            # El indice explicito fija el orden de los tramos: sin el, Vega
            # los apila por orden alfabetico y el "No" puede quedar antes que
            # el "Si" en unas barras y despues en otras.
            fila[sub] = sc
            fila["_orden"] = subclases.index(sc) if sc in subclases else 99
        if extras.get((c, sc)):
            fila["detalle"] = extras[(c, sc)]
        filas.append(fila)
    df = pd.DataFrame(filas)
    if not df.empty:
        df = df.sort_values([cat] + (["_orden"] if sub else []))
    return df, categorias, subclases


def _alto_barras(n_categorias, agrupadas=False, n_sub=1):
    """Altura util: barras finas, con aire entre ellas."""
    paso = 30 if not agrupadas else max(30, 16 * max(n_sub, 1))
    return int(min(760, max(160, n_categorias * paso + 60)))


def _envolver(texto, ancho=118):
    """Parte la descripcion en lineas para el subtitulo del grafico."""
    if not texto:
        return ""
    palabras, lineas, actual = texto.split(), [], ""
    for palabra in palabras:
        if len(actual) + len(palabra) + 1 > ancho:
            lineas.append(actual)
            actual = palabra
        else:
            actual = f"{actual} {palabra}".strip()
    if actual:
        lineas.append(actual)
    return lineas[:3]


def _eje_x(alt, serie, escala, tope=0):
    """Eje de valores.

    Donde se cuentan fichas, actores o titulares las marcas tienen que caer en
    enteros: con el reparto automatico un eje que llega a 2 se rotula
    "0 1 1 2". Por debajo de una docena se fijan los enteros uno a uno y por
    encima basta con exigir el paso minimo de 1.
    """
    fmt = _formato(serie)
    ejes = {"format": fmt, "labelFontSize": 10, "labelFlush": False}
    if serie.get("decimales", 0) == 0:
        ejes["tickMinStep"] = 1
        if 0 < tope <= 12:
            ejes["values"] = list(range(0, int(math.ceil(tope)) + 1))
        else:
            # Un eje de poblacion llega a cuatro cifras: sin limitar el numero
            # de marcas, Vega reparte una cada 50 y los rotulos se amontonan.
            ejes["tickCount"] = 8
    return alt.X(f"{serie['val']}:Q", title=serie.get("eje_x") or None,
                 scale=escala, axis=alt.Axis(**ejes))


def _barras_simples(alt, base, df, serie, t, categorias):
    """Barras horizontales de una sola serie.

    El color codifica la propia magnitud (rampa de un solo tono) salvo que la
    serie traiga un color explicito por clase, que es el caso de las escalas
    ordenadas favorable <-> desfavorable.
    """
    cat, val = serie["cat"], serie["val"]
    unidad, fmt = serie.get("unidad", ""), _formato(serie)
    colores = serie.get("colores")
    if colores:
        dominio = [c for c in categorias if c in colores]
        color = alt.Color(f"{cat}:N", legend=None, scale=alt.Scale(
            domain=dominio, range=[colores[c] for c in dominio]))
    else:
        color = alt.Color(f"{val}:Q", legend=None, scale=alt.Scale(
            range=_escala_valor(serie), type="linear"))

    tope = float(df[val].max() or 0)
    escala_x = alt.Scale(domain=[0, serie.get("maximo") or max(tope * 1.16, 1)],
                         nice=False)
    eje_x = _eje_x(alt, serie, escala_x, tope)
    barras = base.mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4,
                           height=alt.RelativeBandSize(0.6)).encode(
        x=eje_x, y=_eje_y(alt, serie, cat, categorias), color=color,
        tooltip=[alt.Tooltip(f"{cat}:N", title="Clase"),
                 alt.Tooltip(f"{val}:Q", title=unidad or "Valor", format=fmt)])
    # El valor al extremo de la barra es lo que sostiene la lectura cuando el
    # tono no alcanza 3:1 de contraste contra el fondo.
    etiquetas = base.mark_text(align="left", dx=7, fontSize=11, font=FUENTE,
                               color=t["tinta_2"], fontWeight="bold").encode(
        x=alt.X(f"{val}:Q", scale=escala_x, title=None, axis=None),
        y=_eje_y(alt, serie, cat, categorias),
        text=alt.Text(f"{val}:Q", format=fmt))
    return (barras + etiquetas)


def _barras_multiples(alt, base, serie, t, forma, categorias, subclases):
    """Barras apiladas (parte-todo) o agrupadas (comparacion de dos medidas)."""
    cat, val, sub = serie["cat"], serie["val"], serie["sub"]
    unidad, fmt = serie.get("unidad", ""), _formato(serie)
    colores = serie.get("colores") or {}
    # La leyenda solo declara las clases que aparecen: una entrada sin tramo
    # en ninguna barra hace dudar de si el dato falta o vale cero.
    color = alt.Color(f"{sub}:N", title=None, sort=subclases,
                      scale=alt.Scale(domain=subclases,
                                      range=[colores.get(x, SIN_DATO)
                                             for x in subclases]))
    tooltip = [alt.Tooltip(f"{cat}:N", title=serie.get("eje_y") or "Ámbito"),
               alt.Tooltip(f"{sub}:N", title="Clase"),
               alt.Tooltip(f"{val}:Q", title=unidad or "Valor", format=fmt)]

    if forma == "apiladas":
        totales = base.data.groupby(cat)[val].sum()
        tope, escala_x = float(totales.max() or 0), alt.Scale(nice=False)
    elif serie.get("maximo"):
        tope = float(serie["maximo"])
        escala_x = alt.Scale(domain=[0, tope], nice=False)
    else:
        tope, escala_x = float(base.data[val].max() or 0), alt.Scale()
    codificacion = {
        "x": _eje_x(alt, serie, escala_x, tope).copy(),
        "y": _eje_y(alt, serie, cat, categorias),
        "color": color,
        "tooltip": tooltip,
        "order": alt.Order("_orden:Q", sort="ascending"),
    }
    codificacion["x"]["stack"] = True if forma == "apiladas" else None
    if forma == "agrupadas":
        codificacion["yOffset"] = alt.YOffset(f"{sub}:N", sort=subclases)
    # El filete del color del fondo separa los tramos contiguos: sin el, dos
    # tramos de tonos vecinos se leen como uno solo.
    return base.mark_bar(stroke=t["fondo"], strokeWidth=2,
                         cornerRadiusTopRight=3, cornerRadiusBottomRight=3,
                         height=alt.RelativeBandSize(0.68)).encode(**codificacion)


def _mapa_calor(alt, base, serie, t):
    """Rejilla categoria x subcategoria con el conteo rotulado en cada celda.

    El dominio de ambos ejes es el declarado completo, no el observado: una
    matriz de influencia x interes tiene que mostrar los nueve cuadrantes,
    incluidos los vacios, porque el cuadrante sin actores tambien informa.
    """
    cat, val, sub = serie["cat"], serie["val"], serie.get("sub")
    orden_cat = serie.get("orden_cat") or sorted(base.data[cat].unique())
    orden_sub = serie.get("orden_sub") or sorted(base.data[sub].unique())
    tope = float(base.data[val].max() or 1)
    escala_color = alt.Scale(range=_escala_valor(serie), domain=[0, tope])
    eje_comun = dict(labelFontSize=11, domain=False, ticks=False, grid=False,
                     labelLimit=200)

    codificacion = {
        "x": alt.X(f"{sub}:N", sort=orden_sub, scale=alt.Scale(domain=orden_sub),
                   title=serie.get("eje_x") or None,
                   axis=alt.Axis(labelAngle=0, orient="top", **eje_comun)),
        "y": alt.Y(f"{cat}:N", sort=orden_cat, scale=alt.Scale(domain=orden_cat),
                   title=serie.get("eje_y") or None, axis=alt.Axis(**eje_comun)),
    }
    tooltip = [alt.Tooltip(f"{cat}:N", title=serie.get("eje_y") or "Fila"),
               alt.Tooltip(f"{sub}:N", title=serie.get("eje_x") or "Columna"),
               alt.Tooltip(f"{val}:Q", title=serie.get("unidad") or "Valor")]
    if "detalle" in base.data.columns:
        tooltip.append(alt.Tooltip("detalle:N", title="Actores"))

    celdas = base.mark_rect(stroke=t["fondo"], strokeWidth=4,
                            cornerRadius=6).encode(
        color=alt.condition(
            alt.datum[val] > 0,
            alt.Color(f"{val}:Q", title=serie.get("unidad") or None,
                      scale=escala_color,
                      legend=alt.Legend(orient="right", direction="vertical",
                                        gradientLength=120, format=",.0f",
                                        # Con pocos actores la escala es
                                        # degenerada: se rotulan los enteros
                                        # en vez de repetir "1 / 1 / 0".
                                        values=sorted({0, *range(1, int(tope) + 1)})
                                        if tope <= 6 else alt.Undefined,
                                        tickMinStep=1)),
            alt.value(t["rejilla"])),
        tooltip=tooltip, **codificacion)
    # El numero va sobre la celda; el rotulo pasa a blanco en la mitad oscura
    # de la rampa, donde la tinta secundaria dejaria de leerse.
    rotulos = (base.transform_filter(alt.datum[val] > 0)
               .mark_text(fontSize=15, fontWeight="bold", font=FUENTE)
               .encode(text=alt.Text(f"{val}:Q", format=",.0f"),
                       color=alt.condition(alt.datum[val] > tope / 2,
                                           alt.value("#FFFFFF"),
                                           alt.value(t["tinta"])),
                       **codificacion))
    return (celdas + rotulos).properties(
        width=alt.Step(150), height=alt.Step(58))


def tabla_serie(serie):
    """DataFrame de respaldo de una serie, con las mismas cifras del grafico.

    Es la vista tabular que acompana a cada grafico: sostiene la lectura de
    los tonos que no alcanzan 3:1 de contraste sobre el fondo y permite
    copiar los valores sin exportar el libro. Sale de la misma matriz que el
    grafico y el Excel (_pivote), de modo que las tres vistas coinciden, y
    solo agrega un total o un porcentaje donde tiene sentido (ver
    _modo_totales): nunca suma porcentajes ni unidades distintas.
    """
    import pandas as pd
    filas = serie.get("filas") or []
    if not filas:
        return pd.DataFrame()
    cat, val, sub = serie["cat"], serie["val"], serie.get("sub")
    if serie.get("forma") == "mapa_calor":
        df = pd.DataFrame(filas).pivot_table(
            index=cat, columns=sub, values=val, aggfunc="sum", fill_value=0)
        orden_cat = [c for c in (serie.get("orden_cat") or []) if c in df.index]
        orden_sub = [c for c in (serie.get("orden_sub") or []) if c in df.columns]
        if orden_cat:
            df = df.reindex(orden_cat + [i for i in df.index if i not in orden_cat])
        if orden_sub:
            df = df[orden_sub + [c for c in df.columns if c not in orden_sub]]
        return df.reset_index()
    modo = _modo_totales(serie)
    categorias, subclases, matriz = _pivote(serie)
    cabecera = serie.get("eje_y") or "Clase"
    if sub:
        df = pd.DataFrame(matriz, columns=subclases, dtype="float")
        df.insert(0, cabecera, categorias)
        if modo in ("ambos", "filas") and len(subclases) > 1:
            df[serie.get("etiqueta_total") or "Total"] = \
                df[subclases].sum(axis=1, min_count=1)
        return df
    etiqueta = serie.get("unidad") or "Valor"
    df = pd.DataFrame({cabecera: categorias,
                       etiqueta: [fila[0] for fila in matriz]})
    total = df[etiqueta].sum()
    base = serie.get("pct_base")
    if modo == "columnas" and total:
        df["% del total"] = (100.0 * df[etiqueta] / total).round(1)
    elif modo == "base" and isinstance(base, int) and base:
        df["% de la base"] = (100.0 * df[etiqueta] / base).round(1)
    return df


# ══════════════════════════════════════════════════════════════════════════
# LIBRO EXCEL CON GRAFICOS NATIVOS
# ══════════════════════════════════════════════════════════════════════════
# Los graficos se generan como objetos nativos de Excel (no como imagenes):
# el usuario puede reordenarlos, cambiar la escala o copiarlos a Word sin
# perder calidad, y las series quedan a la vista en la misma hoja.

_BORDE = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
_ALTO_GRAFICO_FILAS = 16          # filas que ocupa un grafico de 8 cm


def _hex(color):
    """Normaliza '#1B4D2E' a '1B4D2E', que es lo que espera openpyxl."""
    return str(color or "").lstrip("#").upper() or "808080"


def _titulo_hoja(ws, titulo, ancho=10):
    """Encabezado institucional ANIN al inicio de la hoja."""
    fila = 1
    for texto in ENCABEZADOS_ANIN:
        celda = ws.cell(fila, 1, texto)
        celda.font = Font(name="Arial", size=9, bold=True, color=ANIN_VERDE)
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ancho)
        fila += 1
    celda = ws.cell(fila, 1, SUBTITULO_PROYECTO)
    celda.font = Font(name="Arial", size=8, italic=True)
    celda.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ancho)
    ws.row_dimensions[fila].height = 24
    fila += 1
    celda = ws.cell(fila, 1, titulo)
    celda.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    celda.fill = PatternFill("solid", fgColor=ANIN_VERDE)
    celda.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ancho)
    ws.row_dimensions[fila].height = 22
    return fila + 2


def _escribir_tabla(ws, fila, titulo, cabeceras, filas, subtitulo=""):
    """Tabla con estilo ANIN. Devuelve (fila_cabecera, fila_siguiente)."""
    if titulo:
        celda = ws.cell(fila, 1, titulo)
        celda.font = Font(name="Arial", size=10, bold=True, color=ANIN_AZUL)
        fila += 1
    if subtitulo:
        celda = ws.cell(fila, 1, subtitulo)
        celda.font = Font(name="Arial", size=8, italic=True, color="595959")
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila,
                       end_column=max(len(cabeceras), 4))
        fila += 1
    fila_cab = fila
    for i, texto in enumerate(cabeceras, start=1):
        celda = ws.cell(fila, i, texto)
        celda.font = Font(name="Arial", size=9, bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=ANIN_VERDE)
        celda.alignment = Alignment(horizontal="center", wrap_text=True,
                                    vertical="center")
        celda.border = _BORDE
    fila += 1
    for j, registro in enumerate(filas):
        for i, valor in enumerate(registro, start=1):
            celda = ws.cell(fila, i, valor)
            celda.font = Font(name="Arial", size=9)
            celda.border = _BORDE
            celda.alignment = Alignment(wrap_text=isinstance(valor, str)
                                        and len(str(valor)) > 40,
                                        vertical="top")
            if j % 2:
                celda.fill = PatternFill("solid", fgColor=ANIN_GRIS)
            if isinstance(valor, float):
                celda.number_format = "#,##0.00"
            elif isinstance(valor, int):
                celda.number_format = "#,##0"
        fila += 1
    return fila_cab, fila


def _acumular(serie):
    """Suma las filas repetidas de una misma (categoria, subclase).

    Varios centros poblados declaran la misma actividad o el mismo peligro: sin
    este paso el apilado dibujaria dos tramos contiguos del mismo color en vez
    de uno solo con el total.
    """
    cat, val, sub = serie["cat"], serie["val"], serie.get("sub")
    acumulado, extras, n = {}, {}, {}
    for f in serie.get("filas") or []:
        clave = (f[cat], f[sub]) if sub else (f[cat], None)
        acumulado[clave] = acumulado.get(clave, 0) + (f[val] or 0)
        n[clave] = n.get(clave, 0) + 1
        # El tooltip del mapa de calor lista los actores de cada celda.
        if f.get("detalle") and clave not in extras:
            extras[clave] = f["detalle"]
    if _modo_totales(serie) == "promedio":
        # Un porcentaje no se suma: si una categoria llegara repetida (dos
        # fichas de un mismo CP), se promedia y nunca supera el 100 %.
        acumulado = {k: v / n[k] for k, v in acumulado.items()}
    return acumulado, extras


def _orden_categorias(serie, acumulado):
    """Orden del eje de categorias, comun a los tres formatos de salida."""
    declarado = serie.get("orden_cat")
    presentes = list(dict.fromkeys(c for c, _ in acumulado))
    if declarado:
        orden = [c for c in declarado if c in set(presentes)]
        return orden + [c for c in presentes if c not in set(orden)]
    if serie.get("escala") == "sino" and serie.get("orden_sub"):
        # Bateria de preguntas: manda el numero de respuestas afirmativas, que
        # es la lectura util (que capacidades existen y cuales faltan).
        favorable = serie["orden_sub"][0]
        return sorted(presentes,
                      key=lambda c: (-acumulado.get((c, favorable), 0), c))
    totales = {}
    for (c, _s), v in acumulado.items():
        totales[c] = totales.get(c, 0) + v
    return sorted(presentes, key=lambda c: (-totales[c], str(c)))


def _pivote(serie):
    """Convierte una serie en (categorias, subclases, matriz de valores)."""
    sub = serie.get("sub")
    acumulado, _extras = _acumular(serie)
    categorias = _orden_categorias(serie, acumulado)
    if not sub:
        return categorias, [serie.get("unidad") or "Valor"], \
            [[acumulado.get((c, None), 0)] for c in categorias]
    subclases = [s for s in (serie.get("orden_sub") or [])
                 if any(k[1] == s for k in acumulado)]
    subclases += sorted({k[1] for k in acumulado} - set(subclases))
    # Una celda sin dato queda vacia en las series de valor (porcentajes,
    # habitantes): leer "0 %" o "0 hab." donde la ficha no trae el dato es
    # un error. En los conteos, la ausencia si es un cero.
    vacio = (None if _modo_totales(serie) == "promedio"
             or serie.get("unidad") == "hab."
             or isinstance(serie.get("pct_base"), tuple) else 0)
    matriz = [[acumulado.get((c, s), vacio) for s in subclases] for c in categorias]
    return categorias, subclases, matriz


def _colores_serie(serie, subclases, categorias):
    """Color de cada serie (o de cada punto) del grafico de Excel."""
    colores = serie.get("colores") or {}
    if serie.get("sub"):
        return [_hex(colores.get(s) or _rampa(subclases, _escala_valor(serie))[i])
                for i, s in enumerate(subclases)]
    if colores:
        return [_hex(colores.get(c, SIN_DATO)) for c in categorias]
    # Sin color por clase, la propia magnitud ordena la rampa: la barra mas
    # larga recibe el paso mas oscuro.
    rampa = _escala_valor(serie)
    acumulado, _extras = _acumular(serie)
    valores = {c: v or 0 for (c, _s), v in acumulado.items()}
    orden = sorted(categorias, key=lambda c: valores.get(c, 0))
    paso = {c: rampa[min(len(rampa) - 1,
                         int(i * len(rampa) / max(len(orden), 1)))]
            for i, c in enumerate(orden)}
    return [_hex(paso.get(c, rampa[2])) for c in categorias]


# ── Totales de cada tabla ────────────────────────────────────────────────
# Un total solo se escribe donde tiene sentido. Sumar porcentajes de
# distintos centros poblados (o los "Si" de preguntas distintas, o familias
# con personas) produce cifras sin significado que el lector toma por
# validas: cada serie declara que total le corresponde.

def _modo_totales(serie):
    """Totales de la tabla de una serie.

    - "columnas": fila TOTAL con =SUMA de cada columna (conteos simples).
    - "ambos": ademas, columna de total por fila (composiciones aditivas).
    - "filas": solo la columna de total por fila (baterias de preguntas:
      sumar los "Si" de preguntas distintas no significa nada).
    - "promedio": fila PROMEDIO con =PROMEDIO (series en %).
    - "base": marcado multiple; fila con la base y columna en % de la base.
    - "ninguno": sin totales (unidades mixtas o indices).
    """
    modo = serie.get("totales", "auto")
    if modo and modo != "auto":
        return modo
    if serie.get("unidad") == "%":
        return "promedio"
    base = serie.get("pct_base", "auto")
    if base is None:
        return "ninguno"
    if isinstance(base, int) and not isinstance(base, bool):
        return "base"
    forma = serie.get("forma")
    if forma == "mapa_calor":
        return "ambos"
    if forma == "apiladas":
        return "filas" if serie.get("escala") == "sino" else "ambos"
    return "columnas"


def _formato_numero(serie):
    decimales = serie.get("decimales", 0) or 0
    return "#,##0" if not decimales else "#,##0." + "0" * decimales


# ── Tablas en porcentaje ─────────────────────────────────────────────────
# Opcion del libro Excel: expresar las tablas de resultados en porcentaje.
# Solo cambia el libro descargado; los graficos y tablas del aplicativo se
# mantienen en valores absolutos.

MODOS_EXCEL = {
    "absoluto": "Valores absolutos",
    "porcentaje": "Porcentaje (%)",
    "ambos": "Valores absolutos y porcentaje (%)",
}

_FORMATO_PCT = "0.0%"


def _base_porcentaje(serie):
    """Denominador de los porcentajes de una serie, o None si no aplica.

    - "fila": barras apiladas; cada categoria suma 100 % (composicion).
    - "columna": barras simples o agrupadas; cada clase sobre el total.
    - "total": mapa de calor; cada celda sobre el total general.
    - ("ref", subclase): cada valor sobre el de una subclase de la misma
      fila (p. ej. asistentes sobre convocados).
    - entero: N.° de fichas, para preguntas de marcado multiple, donde la
      suma supera el numero de fichas.
    - None: la serie ya esta en porcentaje o mezcla unidades.
    """
    base = serie.get("pct_base", "auto")
    if base is None or serie.get("unidad") == "%":
        return None
    if isinstance(base, tuple) and not any(
            f.get(serie.get("sub")) == base[1] for f in serie.get("filas") or []):
        return None             # ningun ambito trae la referencia (convocados)
    if base != "auto":
        return base
    if serie.get("forma") == "mapa_calor":
        return "total"
    if serie.get("sub") and serie.get("forma") == "apiladas":
        return "fila"
    return "columna"


def _texto_base(base, subclases, base_texto="ficha(s)"):
    if base == "fila":
        return "Porcentaje sobre el total de cada fila (cada categoría suma 100 %)."
    if base == "columna":
        return ("Porcentaje sobre el total de cada columna (la columna suma "
                "100 %).")
    if base == "total":
        return "Porcentaje sobre el total general de la tabla."
    if isinstance(base, tuple):
        return f"Porcentaje respecto de «{base[1]}» en cada fila."
    return (f"Porcentaje sobre {base} {base_texto}; por el marcado "
            "múltiple, la suma puede superar el 100 %.")


def _div(a, b):
    return (a / b) if b else 0.0


def _matriz_porcentaje(base, subclases, matriz):
    """Matriz en fracciones (0-1, formato 0.0 %) y su fila de totales."""
    n_sub = len(subclases)
    col_tot = [sum(f[j] or 0 for f in matriz) for j in range(n_sub)]
    gran = sum(col_tot)
    if base == "fila":
        pct = [[_div(v or 0, sum(x or 0 for x in f)) for v in f] for f in matriz]
        tot = [_div(c, gran) for c in col_tot]
    elif base == "columna":
        pct = [[_div(v or 0, col_tot[j]) for j, v in enumerate(f)] for f in matriz]
        tot = [1.0 if c else 0.0 for c in col_tot]
    elif base == "total":
        pct = [[_div(v or 0, gran) for v in f] for f in matriz]
        tot = [_div(c, gran) for c in col_tot]
    elif isinstance(base, tuple):
        # Solo las filas con referencia (p. ej. talleres con convocados): una
        # fila sin denominador queda vacia y no entra al total, que asi no
        # puede superar el 100 % por asistentes sin convocados.
        k = subclases.index(base[1]) if base[1] in subclases else None
        con_ref = [f for f in matriz if k is not None and (f[k] or 0) > 0]
        pct = [[_div(v or 0, f[k]) if k is not None and (f[k] or 0) > 0 else None
                for v in f] for f in matriz]
        ref_tot = sum(f[k] for f in con_ref) if k is not None else 0
        tot = [_div(sum(f[j] or 0 for f in con_ref), ref_tot) for j in range(n_sub)]
    else:
        pct = [[_div(v or 0, base) for v in f] for f in matriz]
        tot = None          # la suma de un marcado multiple no es un total
    return pct, tot


def _formulas_porcentaje(base, subclases, fila_ini, n_cat, fila_tot, base_ref=None):
    """Formulas de la tabla % que apuntan a la tabla de valores absolutos
    (filas fila_ini..fila_ini+n_cat-1, columnas B..). `fila_tot` es la fila
    TOTAL de esa tabla, o None si no la tiene (entonces la formula suma el
    rango). Devuelve (celdas, total)."""
    n_sub = len(subclases)
    col = [get_column_letter(2 + j) for j in range(n_sub)]
    ult = col[-1]
    fin = fila_ini + n_cat - 1
    celdas, total = [], []
    if fila_tot is None:
        return _formulas_porcentaje_sin_total(base, subclases, col, fila_ini,
                                              fin, n_cat, base_ref)
    k = (subclases.index(base[1])
         if isinstance(base, tuple) and base[1] in subclases else None)
    for i in range(n_cat):
        r = fila_ini + i
        fila = []
        for c in col:
            if base == "fila":
                f = f"=IFERROR({c}{r}/SUM($B{r}:${ult}{r}),0)"
            elif base == "columna":
                f = f"=IFERROR({c}{r}/{c}${fila_tot},0)"
            elif base == "total":
                f = f"=IFERROR({c}{r}/SUM($B${fila_ini}:${ult}${fin}),0)"
            elif isinstance(base, tuple):
                f = (f'=IF(N(${col[k]}{r})>0,{c}{r}/${col[k]}{r},"")'
                     if k is not None else "")
            else:
                f = f"=IFERROR({c}{r}/{base_ref or base},0)"
            fila.append(f)
        celdas.append(fila)
    for c in col:
        if base in ("fila", "total"):
            total.append(f"=IFERROR({c}{fila_tot}/SUM($B${fila_tot}:${ult}${fila_tot}),0)")
        elif base == "columna":
            total.append(f"=IF({c}{fila_tot}>0,1,0)")
        elif isinstance(base, tuple) and k is not None:
            ref = f"${col[k]}${fila_ini}:${col[k]}${fin}"
            total.append(f'=IFERROR(SUMIFS({c}{fila_ini}:{c}{fin},{ref},">0")'
                         f'/SUMIFS({ref},{ref},">0"),0)')
        else:
            total = None
            break
    return celdas, total


def _formulas_porcentaje_sin_total(base, subclases, col, fila_ini, fin, n_cat,
                                   base_ref=None):
    """Variante de _formulas_porcentaje para tablas sin fila TOTAL: los
    denominadores se calculan sobre el rango de datos."""
    ult = col[-1]
    k = (subclases.index(base[1])
         if isinstance(base, tuple) and base[1] in subclases else None)
    celdas = []
    for i in range(n_cat):
        r = fila_ini + i
        fila = []
        for c in col:
            if base == "fila":
                f = f"=IFERROR({c}{r}/SUM($B{r}:${ult}{r}),0)"
            elif base == "columna":
                f = f"=IFERROR({c}{r}/SUM({c}${fila_ini}:{c}${fin}),0)"
            elif base == "total":
                f = f"=IFERROR({c}{r}/SUM($B${fila_ini}:${ult}${fin}),0)"
            elif isinstance(base, tuple):
                f = (f'=IF(N(${col[k]}{r})>0,{c}{r}/${col[k]}{r},"")'
                     if k is not None else "")
            elif isinstance(base, int):
                # La base es la celda de la tabla de valores, si la hay: al
                # corregirla se recalculan las dos columnas de %.
                f = f"=IFERROR({c}{r}/{base_ref or base},0)"
            else:
                f = 0
            fila.append(f)
        celdas.append(fila)
    if base in ("fila", "total"):
        total = [f"=IFERROR(SUM({c}{fila_ini}:{c}{fin})/SUM($B${fila_ini}:${ult}${fin}),0)"
                 for c in col]
    elif base == "columna":
        total = [f"=IF(SUM({c}{fila_ini}:{c}{fin})>0,1,0)" for c in col]
    elif isinstance(base, tuple) and k is not None:
        ref = f"${col[k]}${fila_ini}:${col[k]}${fin}"
        total = [f'=IFERROR(SUMIFS({c}{fila_ini}:{c}{fin},{ref},">0")'
                 f'/SUMIFS({ref},{ref},">0"),0)' for c in col]
    else:
        total = None
    return celdas, total


def _tabla_porcentaje(ws, fila, titulo, cabecera_cat, categorias, subclases,
                      celdas, total, base, simple=False, base_texto="ficha(s)"):
    """Escribe una tabla en porcentaje con estilo ANIN. Devuelve
    (fila_cabecera, fila_siguiente)."""
    if simple:
        cabeceras = [cabecera_cat, "Porcentaje (%)"]
    else:
        cabeceras = [cabecera_cat] + [f"{s} (%)" for s in subclases]
    filas = [[c] + list(v) for c, v in zip(categorias, celdas)]
    subt = _texto_base(base, subclases, base_texto)
    fila_cab, fila_fin = _escribir_tabla(ws, fila, titulo, cabeceras, filas,
                                         subtitulo=subt)
    for r in range(fila_cab + 1, fila_fin):
        for j in range(len(subclases)):
            ws.cell(r, 2 + j).number_format = _FORMATO_PCT
    if total is False:
        pass                            # sin fila de cierre
    elif total is not None:
        ws.cell(fila_fin, 1, "TOTAL").font = Font(name="Arial", size=9, bold=True)
        for j, v in enumerate(total):
            celda = ws.cell(fila_fin, 2 + j, v)
            celda.font = Font(name="Arial", size=9, bold=True)
            celda.number_format = _FORMATO_PCT
    else:
        celda = ws.cell(fila_fin, 1, f"Base: {base} {base_texto}")
        celda.font = Font(name="Arial", size=9, bold=True)
    return fila_cab, fila_fin


def _grafico_de_serie(ws, serie, fila_cab, categorias, subclases, colores,
                      ancla, porcentaje=False):
    """Agrega a la hoja el grafico nativo que corresponde a la serie."""
    n_cat, n_sub = len(categorias), len(subclases)
    if not n_cat or not n_sub:
        return
    graf = BarChart()
    graf.type = "bar"                      # barras horizontales, como en la app
    graf.style = 2
    graf.title = serie.get("titulo", "")
    graf.y_axis.title = ("Porcentaje (%)" if porcentaje
                         else serie.get("unidad") or serie.get("eje_x") or "")
    if porcentaje:
        graf.y_axis.number_format = "0%"
    graf.x_axis.title = ""
    graf.gapWidth = 60
    graf.height = 8.5
    graf.width = 17 if n_sub > 1 else 15
    graf.grouping = "stacked" if serie.get("forma") == "apiladas" else "clustered"
    if graf.grouping == "stacked":
        graf.overlap = 100

    datos = Reference(ws, min_col=2, max_col=1 + n_sub,
                      min_row=fila_cab, max_row=fila_cab + n_cat)
    cats = Reference(ws, min_col=1, min_row=fila_cab + 1,
                     max_row=fila_cab + n_cat)
    graf.add_data(datos, titles_from_data=True)
    graf.set_categories(cats)

    if n_sub > 1:
        for i, s in enumerate(graf.series):
            s.graphicalProperties = GraphicalProperties(solidFill=colores[i])
            s.graphicalProperties.line.solidFill = "FFFFFF"
            s.graphicalProperties.line.width = 12700        # 1 pt
        graf.legend.position = "b"
    else:
        graf.legend = None
        serie_excel = graf.series[0]
        serie_excel.graphicalProperties = GraphicalProperties(
            solidFill=colores[0] if colores else _hex(RAMPA_NEUTRA[3]))
        # Un color por barra: en Excel eso se expresa como puntos de datos.
        serie_excel.data_points = [
            DataPoint(idx=i, spPr=GraphicalProperties(solidFill=color))
            for i, color in enumerate(colores)]
        serie_excel.dLbls = DataLabelList()
        serie_excel.dLbls.showVal = True
        serie_excel.dLbls.showSerName = False
        serie_excel.dLbls.showCatName = False
        serie_excel.dLbls.showLegendKey = False
        if porcentaje:
            serie_excel.dLbls.numFmt = _FORMATO_PCT
    ws.add_chart(graf, ancla)


def _hoja_seccion(wb, seccion, usados, modo="absoluto"):
    """Una hoja por seccion: cada serie con su tabla y su grafico al lado.

    `modo` (ver MODOS_EXCEL): "absoluto" conserva el libro tal cual;
    "porcentaje" reemplaza cada tabla por su version en % (y el grafico la
    sigue); "ambos" agrega bajo la tabla de valores su tabla en %, con
    formulas que apuntan a ella.
    """
    ws = wb.create_sheet(_nombre_hoja(seccion["titulo"], usados))
    fila = _titulo_hoja(ws, seccion["titulo"].upper(), ancho=10)
    for col, ancho in zip("ABCDEFGHIJ", (46, 15, 15, 15, 15, 15, 15, 15, 15, 15)):
        ws.column_dimensions[col].width = ancho
    if seccion.get("descripcion"):
        celda = ws.cell(fila, 1, seccion["descripcion"])
        celda.font = Font(name="Arial", size=9, italic=True, color="595959")
        fila += 2

    for i, serie in enumerate(seccion.get("series", []), start=1):
        categorias, subclases, matriz = _pivote(serie)
        if not categorias:
            continue
        colores = _colores_serie(serie, subclases, categorias)
        cab_cat = serie.get("eje_y") or "Clase"
        titulo = f"{chr(64 + i)}. {serie['titulo']}"
        base = _base_porcentaje(serie) if modo != "absoluto" else None
        aviso_na = (modo != "absoluto" and base is None)

        base_texto = serie.get("base_texto") or "ficha(s)"
        n_cols = len(subclases)
        if base is not None and modo == "porcentaje":
            celdas, total = _matriz_porcentaje(base, subclases, matriz)
            if _modo_totales(serie) == "filas":
                total = False           # cada fila es una pregunta distinta
            fila_cab, fila_fin = _tabla_porcentaje(
                ws, fila, titulo, cab_cat, categorias, subclases, celdas,
                total, base, simple=not serie.get("sub"), base_texto=base_texto)
            graf_pct = True
        else:
            fila_cab, fila_fin, n_cols = _tabla_valores(
                ws, fila, titulo, cab_cat, categorias, subclases, matriz, serie,
                aviso_na)
            graf_pct = False
        ancla = f"{get_column_letter(3 + n_cols)}{fila_cab}"
        _grafico_de_serie(ws, serie, fila_cab, categorias, subclases, colores,
                          ancla, porcentaje=graf_pct)
        fila_sig = fila_fin + 1
        if serie.get("nota"):
            celda = ws.cell(fila_sig, 1, "Fuente: " + serie["nota"])
            celda.font = Font(name="Arial", size=8, italic=True, color="595959")
            ws.merge_cells(start_row=fila_sig, start_column=1,
                           end_row=fila_sig, end_column=8)
            fila_sig += 1
        if base is not None and modo == "ambos":
            # La fila TOTAL de la tabla de valores es la ultima con datos; si
            # la serie no lleva total por columnas se calcula en la formula.
            fila_tot = fila_fin if _modo_totales(serie) in ("columnas", "ambos") \
                else None
            base_ref = (f"$B${fila_fin}" if _modo_totales(serie) == "base"
                        and isinstance(base, int) else None)
            celdas, total = _formulas_porcentaje(
                base, subclases, fila_cab + 1, len(categorias), fila_tot, base_ref)
            if _modo_totales(serie) == "filas":
                total = False
            _cab, fila_fin = _tabla_porcentaje(
                ws, fila_sig + 1, f"{titulo} (%)", cab_cat, categorias,
                subclases, celdas, total, base, simple=not serie.get("sub"),
                base_texto=base_texto)
        fila = max(fila_fin + 3,
                   fila_cab + _ALTO_GRAFICO_FILAS + 2)
    ws.sheet_view.showGridLines = False
    return ws


def _tabla_valores(ws, fila, titulo, cab_cat, categorias, subclases, matriz,
                   serie, aviso_na=False):
    """Tabla de valores de una serie con sus totales como formula.

    Los totales van como formula para que el usuario pueda corregir una
    celda y ver el recalculo; cuales se escriben depende de la naturaleza de
    la serie (ver _modo_totales). Devuelve (fila_cabecera, fila_siguiente,
    n_columnas_de_datos).
    """
    modo_tot = _modo_totales(serie)
    n_sub = len(subclases)
    total_fila = modo_tot in ("ambos", "filas") and bool(serie.get("sub")) \
        and n_sub > 1
    base = serie.get("pct_base")
    con_base = modo_tot == "base" and isinstance(base, int) and base > 0
    cabeceras = [cab_cat] + list(subclases)
    if total_fila:
        cabeceras.append(serie.get("etiqueta_total") or "Total")
    if con_base:
        cabeceras.append("% de la base")
    filas = [[c] + [_celda_num(v) for v in valores]
             for c, valores in zip(categorias, matriz)]
    subt = serie.get("descripcion", "")
    if aviso_na:
        subt = (subt + " " if subt else "") + (
            "(Se presenta en valores: la serie ya está expresada en "
            "porcentaje o combina unidades distintas.)")
    fila_cab, fila_fin = _escribir_tabla(ws, fila, titulo, cabeceras, filas,
                                         subtitulo=subt)
    formato = _formato_numero(serie)
    negrita = Font(name="Arial", size=9, bold=True)
    ult = get_column_letter(1 + n_sub)
    for r in range(fila_cab + 1, fila_fin):
        for j in range(n_sub):
            ws.cell(r, 2 + j).number_format = formato
        if total_fila:
            celda = ws.cell(r, 2 + n_sub, f"=SUM(B{r}:{ult}{r})")
            celda.font, celda.number_format = negrita, formato
            celda.border = _BORDE
        if con_base:
            # La base va en la fila de cierre (columna B): la formula la
            # referencia, de modo que corregirla recalcula todos los %.
            celda = ws.cell(r, 2 + n_sub, f"=IFERROR(B{r}/$B${fila_fin},0)")
            celda.number_format, celda.border = _FORMATO_PCT, _BORDE
    n_cols = n_sub + (1 if total_fila or con_base else 0)
    ini, fin = fila_cab + 1, fila_fin - 1
    if not categorias:
        return fila_cab, fila_fin, n_cols
    if modo_tot in ("columnas", "ambos"):
        ws.cell(fila_fin, 1, "TOTAL").font = negrita
        for j in range(n_sub + (1 if total_fila else 0)):
            letra = get_column_letter(2 + j)
            celda = ws.cell(fila_fin, 2 + j, f"=SUM({letra}{ini}:{letra}{fin})")
            celda.font, celda.number_format = negrita, formato
    elif modo_tot == "promedio":
        ws.cell(fila_fin, 1, "PROMEDIO (CP con dato)").font = negrita
        for j in range(n_sub):
            letra = get_column_letter(2 + j)
            celda = ws.cell(fila_fin, 2 + j,
                            f'=IFERROR(AVERAGE({letra}{ini}:{letra}{fin}),"")')
            celda.font = negrita
            celda.number_format = "0.0"
    elif con_base:
        ws.cell(fila_fin, 1, f"Base: {base} {serie.get('base_texto') or 'ficha(s)'}"
                ).font = negrita
        celda = ws.cell(fila_fin, 2, base)
        celda.font, celda.number_format = negrita, "#,##0"
    # Como en _escribir_tabla, fila_fin es la fila siguiente a los datos: la
    # del total cuando lo hay; la nota de fuente va debajo.
    return fila_cab, fila_fin, n_cols


def _celda_num(valor):
    """Entero cuando el valor es entero: evita '12.00' donde se cuentan fichas."""
    if isinstance(valor, float) and valor.is_integer():
        return int(valor)
    return valor


def _nombre_hoja(nombre, usados):
    """Nombre de hoja valido y unico (Excel: 31 caracteres, sin : \\ / ? * [ ])."""
    limpio = re.sub(r"[\\/?*\[\]:]", " ", str(nombre)).strip() or "Hoja"
    limpio = limpio.replace("·", "-")[:31].strip()
    base, i = limpio, 2
    while limpio.lower() in usados:
        sufijo = f" {i}"
        limpio = base[:31 - len(sufijo)] + sufijo
        i += 1
    usados.add(limpio.lower())
    return limpio


def _valor_resumen(valor):
    """'11 225' -> 11225 (numero, no texto); lo demas ('1 / 2', 's/d',
    '12.50 ha') queda como esta."""
    texto = str(valor)
    if re.fullmatch(r"\d{1,3}( \d{3})*", texto):
        return int(texto.replace(" ", ""))
    return valor


def _hoja_resumen(wb, informe, modo="absoluto"):
    """Portada del libro: identificacion, cifras de cabecera y avisos."""
    ws = wb.active
    ws.title = "Resumen DS"
    titulo = (f"DIAGNOSTICO SOCIAL - BLOQUE {informe['codigo']}"
              if informe.get("alcance") == "bloque"
              else f"DIAGNOSTICO SOCIAL CONSOLIDADO - {informe['codigo']}")
    fila = _titulo_hoja(ws, titulo, ancho=6)
    for col, ancho in zip("ABCDEF", (42, 24, 34, 16, 16, 16)):
        ws.column_dimensions[col].width = ancho

    bloque = informe.get("bloque") or {}
    consolidado = informe.get("alcance") != "bloque"
    generales = [
        ("Ámbito" if consolidado else "Código del bloque",
         _txt(bloque.get("codigo")) or informe["codigo"]),
        ("Bloques del ámbito", ", ".join(informe.get("bloques_ambito") or [])
         if consolidado else ""),
        ("Microcuenca", _txt(bloque.get("microcuenca"))),
        ("Provincia", _txt(bloque.get("provincia"))),
        ("Distrito", _txt(bloque.get("distrito"))),
        (f"Centros poblados del catálogo INEI "
         f"({len(informe.get('centros_poblados') or [])})",
         ", ".join(informe.get("centros_poblados") or []) or "—"),
        ("Fichas sociales vigentes", informe.get("n_registros", 0)),
        ("Fecha de emisión", informe["generado"].strftime("%d/%m/%Y %H:%M")),
        ("Expresión de las tablas", MODOS_EXCEL.get(modo, MODOS_EXCEL["absoluto"])),
    ]
    fila_cab, fila = _escribir_tabla(
        ws, fila, "1. Identificación", ["Campo", "Valor"],
        [[e, v] for e, v in generales if v not in ("", None)])
    fila += 1

    fila_cab, fila = _escribir_tabla(
        ws, fila, "2. Cifras de cabecera",
        ["Indicador", "Valor", "Detalle"],
        [[m["etiqueta"], _valor_resumen(m["valor"]), m.get("detalle", "")]
         for m in informe.get("metricas", [])])
    fila += 1

    inventario = [[s["titulo"], len(s.get("series", [])),
                   s.get("descripcion", "")]
                  for s in informe.get("secciones", [])]
    if inventario:
        fila_cab, fila = _escribir_tabla(
            ws, fila, "3. Contenido del libro",
            ["Sección", "N.° de gráficos", "Alcance"], inventario,
            subtitulo="Cada sección ocupa una hoja con su tabla de datos y su "
                      "gráfico nativo de Excel, editable por el usuario.")
        fila += 1

    if informe.get("avisos"):
        for aviso in informe["avisos"]:
            celda = ws.cell(fila, 1, "Aviso: " + aviso)
            celda.font = Font(name="Arial", size=9, italic=True, color="9C5700")
            ws.merge_cells(start_row=fila, start_column=1, end_row=fila,
                           end_column=6)
            fila += 1
        fila += 1

    celda = ws.cell(fila, 1,
                    "Libro generado por el aplicativo IN Piura sobre las fichas "
                    "F-DS-01 a F-DS-07 registradas. Solo se grafica lo "
                    "declarado en campo: los campos sin respuesta no se "
                    "estiman ni se completan por analogía, conforme a la "
                    "declaración de integridad de datos del proyecto.")
    celda.font = Font(name="Arial", size=8, italic=True)
    celda.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila + 2, end_column=6)
    ws.sheet_view.showGridLines = False


def _hoja_tablas(wb, seccion, usados):
    """Hojas de respaldo con el detalle fila a fila de las fichas.

    `seccion["formato_tablas"][titulo]` puede pedir, para una tabla:
      - "totales": columnas numericas con fila TOTAL (=SUMA);
      - "calculadas": columnas derivadas como formula, cada una
        (nombre, "suma" | "resta", [columnas]); "resta" es la primera menos
        las demas. Quedan vacias si falta alguno de los datos, para no
        confundir "sin dato" con cero.
    """
    formatos = seccion.get("formato_tablas") or {}
    for titulo, filas in seccion.get("tablas", []):
        if not filas:
            continue
        opciones = formatos.get(titulo) or {}
        calculadas = opciones.get("calculadas") or []
        ws = wb.create_sheet(_nombre_hoja(f"T {titulo}", usados))
        cabeceras = list(filas[0].keys())
        todas = cabeceras + [nombre for nombre, _op, _cols in calculadas]
        fila = _titulo_hoja(ws, titulo.upper(), ancho=min(len(todas), 12) or 6)
        for i, cab in enumerate(todas, start=1):
            ws.column_dimensions[get_column_letter(i)].width = \
                min(max(len(str(cab)) + 4, 14), 46)
        fila_cab, fila_fin = _escribir_tabla(
            ws, fila, "", todas,
            [[_celda_num(f.get(c)) if isinstance(f.get(c), (int, float))
              else _txt(f.get(c)) for c in cabeceras] for f in filas])
        letra = {c: get_column_letter(i) for i, c in enumerate(todas, start=1)}
        negrita = Font(name="Arial", size=9, bold=True)
        for r in range(fila_cab + 1, fila_fin):
            for j, (nombre, operacion, columnas) in enumerate(calculadas):
                refs = [f"{letra[c]}{r}" for c in columnas if c in letra]
                if len(refs) != len(columnas):
                    continue
                expresion = ("+".join(refs) if operacion == "suma"
                             else refs[0] + "".join(f"-{x}" for x in refs[1:]))
                condicion = ",".join(f"ISNUMBER({x})" for x in refs)
                celda = ws.cell(r, len(cabeceras) + 1 + j,
                                f'=IF(AND({condicion}),{expresion},"")')
                celda.font = Font(name="Arial", size=9, italic=True)
                celda.border = _BORDE
                celda.number_format = "#,##0"
                if (r - fila_cab - 1) % 2:
                    celda.fill = PatternFill("solid", fgColor=ANIN_GRIS)
        sumar = [c for c in (opciones.get("totales") or []) if c in letra]
        sumar += [nombre for nombre, _op, _cols in calculadas]
        if sumar:
            ws.cell(fila_fin, 1, "TOTAL").font = negrita
            for c in sumar:
                col = letra[c]
                rango = f"{col}{fila_cab + 1}:{col}{fila_fin - 1}"
                # SUBTOTAL(109) suma solo las filas visibles: con el
                # autofiltro el total sigue a la seleccion del usuario.
                celda = ws.cell(fila_fin, todas.index(c) + 1,
                                f'=IF(COUNT({rango})=0,"",SUBTOTAL(109,{rango}))')
                celda.font, celda.number_format = negrita, "#,##0"
        ws.freeze_panes = ws.cell(fila_cab + 1, 2)
        ws.auto_filter.ref = (f"A{fila_cab}:"
                              f"{get_column_letter(len(todas))}"
                              f"{fila_cab + len(filas)}")
        ws.sheet_view.showGridLines = False


def generar_excel_social(informe, modo="absoluto"):
    """Libro Excel del Diagnostico Social con graficos nativos. Devuelve bytes.

    `modo`: "absoluto" (por defecto), "porcentaje" o "ambos"; ver MODOS_EXCEL.
    Solo afecta a las tablas de resultados de las hojas de cada ficha.
    """
    if modo not in MODOS_EXCEL:
        raise ValueError(f"Modo de Excel no reconocido: {modo!r}")
    wb = Workbook()
    usados = {"resumen ds"}
    _hoja_resumen(wb, informe, modo)
    for seccion in informe.get("secciones", []):
        if seccion.get("series"):
            _hoja_seccion(wb, seccion, usados, modo)
    for seccion in informe.get("secciones", []):
        _hoja_tablas(wb, seccion, usados)
    if informe.get("control"):
        _hoja_tablas(wb, informe["control"], usados)
    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue()


def nombre_excel(informe, modo="absoluto"):
    marca = datetime.now().strftime("%Y%m%d_%H%M%S")
    sufijo = {"porcentaje": "_PCT", "ambos": "_VAL_PCT"}.get(modo, "")
    if informe.get("alcance") == "bloque":
        return (f"Graficos_DS_Bloque_{informe['codigo']}_IN_Piura"
                f"{sufijo}_{marca}.xlsx")
    return f"Graficos_DS_Consolidado_IN_Piura{sufijo}_{marca}.xlsx"


# ══════════════════════════════════════════════════════════════════════════
# ANEXO GRAFICO EN PDF
# ══════════════════════════════════════════════════════════════════════════
# Reutiliza el PDF institucional de los resumenes territoriales (cabecera,
# pie, tablas y graficos vectoriales) para que la ficha social y la
# territorial salgan con la misma factura.

def _rgb(color):
    """'#2E7D4F' -> (46, 125, 79)."""
    texto = str(color or "").lstrip("#")
    if len(texto) != 6:
        return (120, 120, 120)
    return tuple(int(texto[i:i + 2], 16) for i in (0, 2, 4))


def _recortar(texto, largo):
    texto = str(texto or "")
    return texto if len(texto) <= largo else texto[:largo - 1] + "…"


def _pdf_barras_apiladas(pdf, serie, alto_fila=6.2):
    """Barras horizontales apiladas dibujadas con primitivas de fpdf2.

    El PDF institucional de los resumenes solo trae barras simples y tortas;
    las escalas ordenadas del diagnostico social necesitan el apilado para
    que se lea el reparto dentro de cada ambito.
    """
    from resumenes_bloques import _s
    categorias, subclases, matriz = _pivote(serie)
    if not categorias or not subclases:
        return
    colores = serie.get("colores") or {}
    rampa = _rampa(subclases, _escala_valor(serie))
    tintas = [_rgb(colores.get(s) or rampa[i]) for i, s in enumerate(subclases)]

    filas_visibles = categorias[:18]
    alto_total = len(filas_visibles) * alto_fila + 12
    pdf._salto_si_falta(alto_total + 12)
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(0, 5, _s(serie.get("titulo", "")), 0, 1, "L")
    if serie.get("descripcion"):
        pdf.set_font("Helvetica", "I", 6.5)
        pdf.set_text_color(80, 80, 80)
        pdf.multi_cell(0, 3.2, _s(serie["descripcion"]))
        pdf.set_text_color(0, 0, 0)

    x0, y0 = 12.0, pdf.get_y() + 1
    ancho = pdf.w - 24
    ancho_etq = ancho * 0.36
    ancho_barra = ancho - ancho_etq - 16
    agrupadas = serie.get("forma") == "agrupadas"
    if agrupadas:
        # Dos medidas distintas (agua y energia, convocados y asistentes):
        # cada una con su barra y su rotulo; sumarlas no tiene sentido. Una
        # celda sin dato (None) no se dibuja ni se rotula como 0.
        maximo = serie.get("maximo") or max(
            (v for f in matriz for v in f if v is not None), default=0) or 1.0
    else:
        matriz = [[v or 0 for v in fila] for fila in matriz]
        maximo = max((sum(f) for f in matriz), default=0) or 1.0

    for i, categoria in enumerate(filas_visibles):
        valores = matriz[categorias.index(categoria)]
        y = y0 + i * alto_fila
        pdf.set_xy(x0, y)
        pdf.set_font("Helvetica", "", 6.3)
        pdf.cell(ancho_etq, alto_fila, _s(_recortar(categoria, 46)), 0, 0, "L")
        x = x0 + ancho_etq
        if agrupadas:
            alto_sub = alto_fila * 0.8 / max(len(valores), 1)
            for j, (valor, tinta) in enumerate(zip(valores, tintas)):
                if valor is None:
                    continue
                ys = y + alto_fila * 0.1 + j * alto_sub
                largo = ancho_barra * valor / maximo if valor else 0
                if largo:
                    pdf.set_fill_color(*tinta)
                    pdf.rect(x, ys + alto_sub * 0.1, max(largo, 0.3),
                             alto_sub * 0.8, "F")
                pdf.set_xy(x + largo + 1.0, ys - 0.4)
                pdf.set_font("Helvetica", "B", 5)
                pdf.cell(12, alto_sub + 0.8, _s(_fmt_pdf(valor)), 0, 0, "L")
            continue
        for valor, tinta in zip(valores, tintas):
            if not valor:
                continue
            largo = ancho_barra * valor / maximo
            pdf.set_fill_color(*tinta)
            pdf.rect(x, y + alto_fila * 0.18, max(largo, 0.3),
                     alto_fila * 0.64, "F")
            x += largo
        pdf.set_xy(x + 1.2, y)
        pdf.set_font("Helvetica", "B", 6)
        pdf.cell(14, alto_fila, _s(_fmt_pdf(sum(valores))), 0, 0, "L")

    y = y0 + len(filas_visibles) * alto_fila + 1.5
    pdf.set_y(y)
    # Leyenda: sin ella el color no identifica nada, y los pares verde/rojo
    # del diagnostico necesitan ese segundo canal.
    pdf.set_font("Helvetica", "", 6.3)
    x = x0
    for subclase, tinta in zip(subclases, tintas):
        etiqueta = _recortar(subclase, 34)
        ancho_item = pdf.get_string_width(_s(etiqueta)) + 8
        if x + ancho_item > pdf.w - 12:
            x = x0
            y += 4.2
            pdf.set_y(y)
        pdf.set_fill_color(*tinta)
        pdf.rect(x, y + 1.0, 2.8, 2.8, "F")
        pdf.set_xy(x + 4, y)
        pdf.cell(ancho_item - 4, 4.6, _s(etiqueta), 0, 0, "L")
        x += ancho_item
    pdf.set_y(y + 6)
    if len(categorias) > len(filas_visibles):
        pdf.nota(f"Se grafican los primeros {len(filas_visibles)} de "
                 f"{len(categorias)} ámbitos, en el orden del gráfico; el "
                 f"detalle completo está en el libro Excel.")
    if serie.get("nota"):
        pdf.nota("Fuente: " + serie["nota"])


def _fmt_pdf(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and not valor.is_integer():
        return f"{valor:,.2f}".replace(",", " ")
    return f"{int(valor):,}".replace(",", " ")


def _pdf_barras_simples(pdf, serie):
    """Barras horizontales de una serie simple, con su color por clase."""
    categorias, _sub, matriz = _pivote(serie)
    if not categorias:
        return
    colores = _colores_serie(serie, [serie.get("unidad") or "Valor"], categorias)
    mapa = {str(c).strip().upper(): _rgb("#" + colores[i])
            for i, c in enumerate(categorias)}
    valores = [fila[0] or 0 for fila in matriz]
    pdf.set_font("Helvetica", "I", 6.5)
    pdf.grafico_barras(serie.get("titulo", ""), categorias, valores,
                       unidad=(" " + serie["unidad"]) if serie.get("unidad") else "",
                       alto=min(96, max(24, 6.4 * len(categorias))),
                       horizontal=True, colores=mapa)
    if serie.get("descripcion"):
        pdf.nota(serie["descripcion"])
    if serie.get("nota"):
        pdf.nota("Fuente: " + serie["nota"])


def generar_pdf_social(informe):
    """Anexo grafico en PDF del Diagnostico Social. Devuelve bytes."""
    from resumenes_bloques import _PDFResumen, _pdf_bytes

    subtitulo = ("ANEXO GRAFICO DEL DIAGNOSTICO SOCIAL - BLOQUE "
                 f"{informe['codigo']}" if informe.get("alcance") == "bloque"
                 else "ANEXO GRAFICO DEL DIAGNOSTICO SOCIAL CONSOLIDADO")
    pdf = _PDFResumen(subtitulo=subtitulo)
    pdf.alias_nb_pages()
    pdf.add_page()

    bloque = informe.get("bloque") or {}
    pdf.seccion("1. Identificacion")
    generales = [
        ("Codigo del bloque", _txt(bloque.get("codigo")) or informe["codigo"]),
        ("Microcuenca", _txt(bloque.get("microcuenca"))),
        ("Provincia", _txt(bloque.get("provincia"))),
        ("Distrito", _txt(bloque.get("distrito"))),
        ("Centros poblados", ", ".join(informe.get("centros_poblados") or []) or "-"),
        ("Fichas sociales vigentes", str(informe.get("n_registros", 0))),
        ("Fecha de emision", informe["generado"].strftime("%d/%m/%Y %H:%M")),
    ]
    pdf.tabla(["Campo", "Valor"], [[e, v] for e, v in generales if v],
              anchos_rel=[1, 2], alineaciones=["L", "L"])

    pdf.seccion("2. Cifras de cabecera")
    pdf.tabla(["Indicador", "Valor", "Detalle"],
              [[m["etiqueta"], m["valor"], m.get("detalle", "")]
               for m in informe.get("metricas", [])],
              anchos_rel=[3, 2, 4], alineaciones=["L", "R", "L"])

    for i, seccion in enumerate(informe.get("secciones", []), start=3):
        series = seccion.get("series") or []
        if not series:
            continue
        # El titulo de seccion no puede quedar solo al pie de la pagina: se
        # reserva sitio para el primer grafico antes de escribirlo.
        pdf._salto_si_falta(48)
        pdf.seccion(f"{i}. {seccion['titulo'].replace('·', '-')}")
        for serie in series:
            if serie.get("forma") in ("apiladas", "agrupadas", "mapa_calor"):
                _pdf_barras_apiladas(pdf, serie)
            else:
                _pdf_barras_simples(pdf, serie)

    if informe.get("avisos"):
        pdf.seccion("Avisos")
        for aviso in informe["avisos"]:
            pdf.nota(aviso)
    pdf.nota("Anexo generado por el aplicativo IN Piura sobre las fichas "
             "F-DS-01 a F-DS-07 registradas. Solo se grafica lo declarado en "
             "campo; los campos sin respuesta no se estiman.")
    return _pdf_bytes(pdf)


def nombre_pdf(informe):
    marca = datetime.now().strftime("%Y%m%d_%H%M%S")
    if informe.get("alcance") == "bloque":
        return f"Anexo_Graficos_DS_Bloque_{informe['codigo']}_IN_Piura_{marca}.pdf"
    return f"Anexo_Graficos_DS_Consolidado_IN_Piura_{marca}.pdf"
