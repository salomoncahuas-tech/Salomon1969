"""
IN Piura - Analitica de las fichas F-DT actualizadas (117 bloques).

Lee los libros de campo "Plantillas Excel FDT actualizadas" -un archivo por
bloque, con las cinco fichas oficiales F-DT-01 a F-DT-05 de la plantilla V5-
y construye sobre ellos las tablas y los graficos que el aplicativo ya
ofrece para los resumenes de bloque, desagregables por bloque, distrito,
provincia y consolidado total.

El parseo no se reimplementa: se usa `excel_diagnostico_territorial`, el
mismo lector con el que el aplicativo importa una ficha a mano, de modo que
lo que se ve en la analitica y lo que se registra en el formulario salen de
la misma lectura.

Criterio de integridad declarativa del proyecto: no se estima ni se infiere
ningun valor ausente. "Por determinar", "Por verificar" y las celdas vacias
no entran en ningun promedio; cada indicador viaja con el numero de bloques
que si lo declaran (`n`), que es lo que permite leer un promedio sin
confundirlo con un censo.

API publica:
    fichas_del_repositorio()            -> [(nombre, contenido)]
    parsear_ficha(contenido, nombre)    -> dict plano de un bloque
    cargar_fichas(catalogo=..., progreso=...) -> [dict]
    tabla_resumen(fichas, nivel)        -> filas de indicadores por grupo
    distribucion(fichas, campo, nivel)  -> reparto de una variable categorica
    inventario(fichas, tipo)            -> filas del inventario con contexto
    ranking(fichas, tipo, campo)        -> frecuencias de un inventario
    generar_excel_fdt(fichas, nivel)    -> bytes (.xlsx con graficos)
    generar_pdf_fdt(fichas, nivel)      -> bytes (.pdf institucional)
"""

import io
import os
import re
import zipfile
from collections import Counter, OrderedDict

from openpyxl import Workbook

import excel_diagnostico_territorial as edt

# Estilo institucional ANIN y utilidades de lectura: se toman del modulo de
# resumenes para que ambas salidas (resumenes y fichas) compartan una sola
# implementacion de la identidad grafica del proyecto.
from resumenes_bloques import (  # noqa: E402
    _agregar_grafico_barras,
    _agregar_grafico_torta,
    _escribir_bloque_datos,
    _fmt,
    _norm,
    _num,
    _pdf_bytes,
    _recortar,
    _titulo_hoja,
    _txt,
    _PDFResumen,
    expandir_zip,
)

# ── Origen de los libros de campo ─────────────────────────────────────────
# Viajan con el aplicativo, igual que los resumenes: la carpeta extraida si
# alguien la deja puesta y, si no, el .zip tal como se publica en GitHub.
CARPETA_FICHAS = "Plantillas Excel FDT actualizadas"
ZIP_FICHAS = "Plantillas Excel FDT actualizadas.zip"

NIVELES = ("bloque", "distrito", "provincia", "total")
ETIQUETA_NIVEL = {
    "bloque": "Bloque",
    "distrito": "Distrito",
    "provincia": "Provincia",
    "total": "Consolidado total",
}

SIN_DATO = "Sin registro en ficha"

# Textos que la plantilla usa para declarar un dato ausente. No son valores:
# no promedian, no cuentan y no se grafican.
_NO_DECLARADOS = {
    "", "-", "--", "n/a", "na", "no aplica", "none", "null", "s/r",
    "por determinar", "por verificar", "por confirmar", "sin registro",
    "sin dato", "sin datos", "sin inventario", "pendiente",
}


# ══════════════════════════════════════════════════════════════════════════
# Catalogo declarativo de indicadores
# ══════════════════════════════════════════════════════════════════════════

# (clave, etiqueta, unidad, agregacion). La agregacion es "promedio" cuando
# el indicador describe una condicion del bloque y "suma" cuando cuenta
# objetos inventariados.
INDICADORES_NUM = (
    ("dt03_cobertura_total", "Cobertura vegetal total", "%", "promedio"),
    ("dt03_cobertura_dosel", "Cobertura de dosel", "%", "promedio"),
    ("dt03_cobertura_arbustiva", "Cobertura arbustiva", "%", "promedio"),
    ("dt03_cobertura_herbacea", "Cobertura herbacea", "%", "promedio"),
    ("dt03_cobertura_hojarasca", "Cobertura de hojarasca", "%", "promedio"),
    ("dt03_suelo_desnudo", "Suelo desnudo", "%", "promedio"),
    ("dt03_altura_estrato_dom", "Altura del estrato dominante", "m", "promedio"),
    ("dt03_altura_max", "Altura maxima", "m", "promedio"),
    ("dt03_dap_promedio", "DAP promedio", "cm", "promedio"),
    ("dt03_superficie_ecosistema", "Superficie de ecosistema declarada", "ha", "suma"),
    ("dt03_pendiente_parcela", "Pendiente de la parcela", "%", "promedio"),
    ("altitud_gps", "Altitud del punto de muestreo", "msnm", "promedio"),
    ("dt02_num_carcavas", "Carcavas declaradas en la sintesis", "n", "suma"),
    ("dt02_longitud_total_carcavas", "Longitud total de carcavas", "m", "suma"),
    ("dt02_pct_bloque_carcavas", "Bloque afectado por carcavas", "%", "promedio"),
    ("dt02_erosion_laminar_pct", "Erosion laminar observada", "%", "promedio"),
    ("dt05_tiempo_dist_capital", "Tiempo a la capital distrital", "min", "promedio"),
    ("dt05_dist_captacion", "Distancia a la captacion mas cercana", "m", "promedio"),
)

# Conteos derivados de los inventarios tabulares. Siempre se suman.
INDICADORES_CONTEO = (
    ("n_carcavas_inventario", "Carcavas inventariadas (F-DT-02)"),
    ("n_taxones", "Taxones del elenco floristico (F-DT-03)"),
    ("n_especies_clave", "Especies clave registradas (F-DT-03)"),
    ("n_especies_amenazadas", "Especies clave con categoria de amenaza"),
    ("n_causas_presentes", "Causas de degradacion presentes (F-DT-04)"),
    ("n_indicadores", "Indicadores cuantitativos (F-DT-04)"),
    ("n_fuentes_agua", "Fuentes de agua inventariadas (F-DT-05)"),
    ("n_fuentes_permanentes", "Fuentes de agua de regimen permanente"),
)

# Observaciones de si/no. Se reportan como numero de bloques con "Si" y su
# porcentaje sobre los bloques que responden (no sobre el total del grupo).
INDICADORES_SINO = (
    ("dt01_afloramientos_rocosos", "Afloramientos rocosos"),
    ("dt01_escarpes_activos", "Escarpes activos"),
    ("dt01_reptacion_suelo", "Reptacion de suelo"),
    ("dt01_deslizamientos_antiguos", "Deslizamientos antiguos"),
    ("dt01_remociones_masa_activas", "Remociones en masa activas"),
    ("dt02_sellamiento_costra", "Sellamiento / costra superficial"),
    ("dt02_compactacion_pisoteo", "Compactacion por pisoteo de ganado"),
    ("dt02_raices_expuestas", "Raices expuestas en superficie"),
    ("dt02_socavamiento_cauce", "Socavamiento de cauce"),
    ("dt05_zona_recarga", "Zona de recarga hidrica"),
    ("dt05_escorrentia_concentrada", "Escorrentia concentrada"),
    ("dt05_interferencia_riego", "Interferencia con sistemas de riego"),
    ("dt05_requiere_ronda", "Requiere ronda o acompanamiento"),
)

# Variables categoricas con reparto de frecuencias.
CATEGORICOS = (
    ("dt02_nivel_erosion_sintesis", "Nivel general de erosion (sintesis F-DT-02)"),
    ("dt02_urgencia_control", "Urgencia de control de erosion"),
    ("dt04_urgencia_intervencion", "Urgencia de intervencion"),
    ("dt04_velocidad_degradacion", "Velocidad de degradacion"),
    ("dt04_reversibilidad", "Reversibilidad tecnica"),
    ("dt03_estado_conservacion_eco", "Estado de conservacion del ecosistema"),
    ("dt03_tipo_ecosistema", "Tipo de ecosistema (UP)"),
    ("dt03_uso_dominante", "Uso actual dominante del suelo"),
    ("dt03_tipo_cobertura_dom", "Tipo de cobertura dominante"),
    ("dt03_regeneracion_natural", "Regeneracion natural"),
    ("dt03_estado_sanitario", "Estado sanitario"),
    ("forma_terreno", "Forma predominante del terreno"),
    ("pendiente", "Rango de pendiente dominante"),
    ("posicion_fisiografica", "Posicion fisiografica"),
    ("exposicion_orientacion", "Exposicion / orientacion"),
    ("rango_altitudinal", "Rango altitudinal"),
    ("dt02_patron_carcavas", "Patron de carcavas dominante"),
    ("dt05_modalidad_acceso", "Modalidad de acceso"),
    ("dt05_tipo_via_final", "Tipo de via final"),
    ("dt05_transitabilidad_seca", "Transitabilidad en epoca seca"),
    ("dt05_transitabilidad_lluviosa", "Transitabilidad en epoca lluviosa"),
    ("dt05_senal_celular", "Senal celular"),
)

# Inventarios tabulares: (tipo, etiqueta, clave en la ficha, claves que
# acreditan que la fila tiene contenido, columnas a mostrar).
INVENTARIOS = (
    ("carcavas", "Carcavas inventariadas (F-DT-02)", "dt02_carcavas",
     ("codigo", "tipo", "longitud_m"),
     (("codigo", "Codigo"), ("tipo", "Tipo"), ("utm_e_ini", "UTM ESTE inicio"),
      ("utm_n_ini", "UTM NORTE inicio"), ("longitud_m", "Longitud (m)"),
      ("prof_m", "Profundidad (m)"), ("ancho_m", "Ancho (m)"),
      ("estado", "Estado"), ("causa", "Causa"))),
    ("floristica", "Elenco floristico (F-DT-03)", "dt03_floristica",
     ("nombre_comun", "nombre_cientifico", "familia"),
     (("nombre_comun", "Nombre comun"), ("nombre_cientifico", "Nombre cientifico"),
      ("familia", "Familia"), ("estrato", "Estrato"), ("origen", "Origen"),
      ("abundancia", "Abundancia"), ("dap_cm", "DAP (cm)"),
      ("altura_m", "Altura (m)"))),
    ("especies_clave", "Especies clave (F-DT-03)", "dt03_especies_clave",
     ("nombre", "categoria"),
     (("nombre", "Especie"), ("categoria", "Categoria"),
      ("estado_uicn", "Estado UICN"), ("utm_e", "UTM ESTE"),
      ("utm_n", "UTM NORTE"), ("n_indiv", "N.o de individuos"),
      ("observacion", "Observacion"))),
    ("causas", "Matriz de causas de degradacion (F-DT-04)", "dt04_causas",
     ("causa",),
     (("causa", "Causa directa"), ("presencia", "Presencia"),
      ("intensidad", "Intensidad"), ("extension", "Extension (%)"),
      ("antiguedad", "Antiguedad"), ("evidencia", "Evidencia"))),
    ("indicadores", "Indicadores cuantitativos (F-DT-04)", "dt04_indicadores",
     ("indicador",),
     (("indicador", "Indicador"), ("unidad", "Unidad"), ("valor", "Valor"),
      ("nivel", "Nivel"), ("umbral", "Umbral"), ("fuente", "Fuente"))),
    ("fuentes_agua", "Fuentes de agua (F-DT-05)", "dt05_fuentes_agua",
     ("tipo", "regimen", "calidad"),
     (("tipo", "Tipo"), ("regimen", "Regimen"), ("calidad", "Calidad"),
      ("utm_e", "UTM ESTE"), ("utm_n", "UTM NORTE"),
      ("distancia_m", "Distancia (m)"), ("uso_obs", "Uso / observacion"))),
)

_INVENTARIOS_POR_TIPO = {inv[0]: inv for inv in INVENTARIOS}


# ══════════════════════════════════════════════════════════════════════════
# Utilidades de normalizacion
# ══════════════════════════════════════════════════════════════════════════

def declarado(valor):
    """True si el campo trae un dato y no una declaracion de ausencia."""
    txt = _txt(valor)
    if not txt:
        return False
    return _norm(txt).strip(" .") not in _NO_DECLARADOS


def valor_num(ficha, clave):
    """Numero declarado en la ficha, o None si el campo no lo declara."""
    if not declarado(ficha.get(clave)):
        return None
    return _num(ficha.get(clave))


def es_si(valor):
    """Interpreta una respuesta de si/no. Devuelve True, False o None."""
    if not declarado(valor):
        return None
    txt = _norm(valor)
    if txt.startswith("si") or txt.startswith("s ") or txt == "s":
        return True
    if txt.startswith("no"):
        return False
    return None


def categoria(ficha, clave):
    """Valor categorico limpio, o el texto convenido de ausencia."""
    valor = _txt(ficha.get(clave))
    return valor if declarado(valor) else SIN_DATO


def _titulo_localidad(texto):
    """Normaliza un nombre de provincia o distrito para poder agrupar.

    Las fichas escriben "Morropon" y "Morropón", "Frias" y "Frías". Se
    conserva el nombre tal como lo escribe la ficha, pero la clave de
    agrupacion ignora acentos y mayusculas.
    """
    return _norm(texto)


def filas_con_contenido(lista, claves):
    """Filas de un inventario que declaran algo en alguna clave dada.

    La plantilla V5 trae filas en blanco para que la brigada escriba en
    campo; el lector las devuelve tal cual y aqui se descartan, porque una
    fila vacia no es un registro.
    """
    salida = []
    for fila in lista or []:
        if any(declarado(fila.get(k)) for k in claves):
            salida.append(fila)
    return salida


# ══════════════════════════════════════════════════════════════════════════
# Carga de los libros de campo del repositorio
# ══════════════════════════════════════════════════════════════════════════

_FICHAS_ZIP_CACHE = {}


def fichas_del_repositorio(carpeta=None):
    """[(nombre, contenido)] de los libros F-DT incluidos en el repositorio.

    Devuelve lista vacia si ni la carpeta ni el .zip viajan en el
    despliegue. Con `carpeta` explicita solo se mira esa ruta.
    """
    base = os.path.dirname(os.path.abspath(__file__))
    ruta = carpeta or os.path.join(base, CARPETA_FICHAS)
    if os.path.isdir(ruta):
        libros = []
        for nombre in sorted(os.listdir(ruta)):
            if not nombre.lower().endswith(".xlsx") or nombre.startswith("~$"):
                continue
            with open(os.path.join(ruta, nombre), "rb") as fh:
                libros.append((nombre, fh.read()))
        return libros
    if carpeta:
        return []
    return fichas_del_zip(os.path.join(base, ZIP_FICHAS))


def fichas_del_zip(ruta):
    """[(nombre, contenido)] de los libros F-DT contenidos en un .zip.

    Se memoriza por (ruta, fecha, tamano) para no descomprimir 7 MB en cada
    rerun de Streamlit. Un .zip ausente o ilegible devuelve lista vacia.
    """
    if not os.path.isfile(ruta):
        return []
    try:
        firma = (ruta, os.path.getmtime(ruta), os.path.getsize(ruta))
    except OSError:
        return []
    if _FICHAS_ZIP_CACHE.get("firma") != firma:
        try:
            with open(ruta, "rb") as fh:
                libros = expandir_zip(fh.read())
        except (OSError, ValueError, zipfile.BadZipFile):
            return []
        _FICHAS_ZIP_CACHE.clear()
        _FICHAS_ZIP_CACHE.update(firma=firma, libros=sorted(libros))
    return list(_FICHAS_ZIP_CACHE["libros"])


def codigo_desde_nombre(nombre_archivo):
    """Codigo de bloque deducido del nombre del archivo.

    Solo se usa cuando la propia ficha no declara el codigo. Los nombres no
    siguen una convencion unica: DT_B13_IN_Piura, F-DT_M6B10_..._rev1_dron,
    Plantilla_DT_Campo_Check_Validada_V5_Bloque23, DT_Bloque67_V5.
    """
    base = os.path.basename(nombre_archivo or "")
    base = re.sub(r"\.xlsx$", "", base, flags=re.IGNORECASE)
    # El guion bajo es caracter de palabra, de modo que "\bB13\b" no casa
    # en "DT_B13_IN_Piura": el limite se escribe a mano.
    for patron in (r"M\d+B\d+(?:-\d+)?", r"[Bb]loque[_\s]*(\d+)",
                   r"(?:^|[^A-Za-z0-9])B(\d+)(?![0-9])"):
        m = re.search(patron, base)
        if m:
            return (m.group(1) if m.groups() else m.group(0)).upper()
    return ""


# ══════════════════════════════════════════════════════════════════════════
# Parseo de una ficha
# ══════════════════════════════════════════════════════════════════════════

def parsear_ficha(contenido, nombre_archivo=""):
    """Lee un libro F-DT y devuelve un diccionario plano del bloque.

    Las cinco fichas se funden en un solo registro: los campos de cabecera
    (codigo, microcuenca, distrito) se repiten entre ellas y la ultima
    lectura no debe pisar a una anterior que si traia dato.
    """
    if isinstance(contenido, (bytes, bytearray)):
        contenido = io.BytesIO(contenido)
    leidas = edt.parsear_excel_dt(contenido)

    ficha = {
        "nombre_archivo": os.path.basename(nombre_archivo or ""),
        "fichas_leidas": [r.get("ficha") for r in leidas],
    }
    for registro in leidas:
        for clave, valor in (registro.get("datos") or {}).items():
            if isinstance(valor, list):
                if valor or clave not in ficha:
                    ficha[clave] = valor
            elif clave not in ficha or (declarado(valor)
                                        and not declarado(ficha[clave])):
                ficha[clave] = valor

    if not declarado(ficha.get("codigo_bloque")):
        ficha["codigo_bloque"] = codigo_desde_nombre(ficha["nombre_archivo"])
    ficha["codigo_bloque"] = _txt(ficha.get("codigo_bloque")).upper()

    _completar_conteos(ficha)
    _completar_numeros(ficha)
    return ficha


def _completar_conteos(ficha):
    """Conteos de los inventarios, descartando las filas en blanco."""
    for tipo, _etiqueta, clave, acreditan, _cols in INVENTARIOS:
        ficha[f"inv_{tipo}"] = filas_con_contenido(ficha.get(clave), acreditan)

    ficha["n_carcavas_inventario"] = len(ficha["inv_carcavas"])
    ficha["n_taxones"] = len(ficha["inv_floristica"])
    ficha["n_especies_clave"] = len(ficha["inv_especies_clave"])
    ficha["n_especies_amenazadas"] = sum(
        1 for f in ficha["inv_especies_clave"]
        if "amenaz" in _norm(f.get("categoria")) or
        _norm(f.get("estado_uicn")).startswith(("en", "vu", "cr")))
    ficha["n_causas_presentes"] = sum(
        1 for f in ficha["inv_causas"] if es_si(f.get("presencia")))
    ficha["n_indicadores"] = len(ficha["inv_indicadores"])
    ficha["n_fuentes_agua"] = len(ficha["inv_fuentes_agua"])
    ficha["n_fuentes_permanentes"] = sum(
        1 for f in ficha["inv_fuentes_agua"]
        if _norm(f.get("regimen")).startswith("permanente"))


def _completar_numeros(ficha):
    """Expone como numero cada indicador numerico declarado."""
    for clave, _etq, _unidad, _agg in INDICADORES_NUM:
        ficha[clave + "_num"] = valor_num(ficha, clave)
    for clave, _etq in INDICADORES_SINO:
        ficha[clave + "_si"] = es_si(ficha.get(clave))


def resolver_localidad(ficha, catalogo=None):
    """Asienta provincia, distrito, microcuenca y area del bloque.

    El catalogo de bloques del aplicativo manda sobre lo escrito en la
    ficha: hay libros que dejaron la localizacion en blanco y alguno que
    trae un valor que no es una provincia. Lo declarado en la ficha se
    conserva aparte para poder contrastarlo.
    """
    codigo = _txt(ficha.get("codigo_bloque")).upper()
    registro = (catalogo or {}).get(codigo, {})
    for campo in ("provincia", "distrito", "microcuenca"):
        declarado_en_ficha = _txt(ficha.get(campo))
        ficha[campo + "_ficha"] = declarado_en_ficha
        del_catalogo = _txt(registro.get(campo))
        if declarado(del_catalogo):
            ficha[campo] = del_catalogo
        elif declarado(declarado_en_ficha):
            ficha[campo] = declarado_en_ficha
        else:
            ficha[campo] = ""
    ficha["area_ha"] = _num(registro.get("area_ha"))
    return ficha


def cargar_fichas(libros=None, catalogo=None, progreso=None):
    """Parsea los libros F-DT y devuelve una ficha por bloque.

    `libros` por omision son los del repositorio. `catalogo` es el catalogo
    de bloques del aplicativo {codigo: {provincia, distrito, microcuenca,
    area_ha}}. `progreso` es un invocable f(i, total, nombre) para informar
    el avance: leer los 117 libros toma del orden de diez segundos.
    """
    libros = fichas_del_repositorio() if libros is None else libros
    fichas, errores = [], []
    total = len(libros)
    for i, (nombre, contenido) in enumerate(libros, start=1):
        try:
            ficha = parsear_ficha(contenido, nombre)
            resolver_localidad(ficha, catalogo)
            fichas.append(ficha)
        except Exception as exc:                       # libro ilegible
            errores.append((nombre, f"{type(exc).__name__}: {exc}"))
        if progreso:
            progreso(i, total, nombre)
    fichas.sort(key=lambda f: _clave_orden(f.get("codigo_bloque", "")))
    return fichas, errores


def _clave_orden(codigo):
    """Ordena 2, 10, M3B1, M12B1 como los lee una persona."""
    partes = re.findall(r"\d+|\D+", str(codigo))
    return tuple((0, int(p)) if p.isdigit() else (1, p) for p in partes)


# ══════════════════════════════════════════════════════════════════════════
# Agregacion por bloque, distrito, provincia y consolidado total
# ══════════════════════════════════════════════════════════════════════════

_CAMPO_NIVEL = {"bloque": "codigo_bloque", "distrito": "distrito",
                "provincia": "provincia"}


def agrupar(fichas, nivel):
    """[(etiqueta, [fichas])] al nivel pedido, en orden de lectura.

    Distrito y provincia se agrupan por nombre normalizado -las fichas
    escriben "Morropon" y "Morropón", "Frias" y "Frías"- y se rotulan con
    la grafia mas frecuente entre las fichas del grupo.
    """
    if nivel not in NIVELES:
        raise ValueError(f"Nivel no reconocido: {nivel}. Use {NIVELES}.")
    if nivel == "total":
        return [(ETIQUETA_NIVEL["total"], list(fichas))] if fichas else []

    campo = _CAMPO_NIVEL[nivel]
    grupos = OrderedDict()
    for ficha in fichas:
        valor = _txt(ficha.get(campo))
        clave = _titulo_localidad(valor) if nivel != "bloque" else valor.upper()
        grupos.setdefault(clave or "", []).append(ficha)

    salida = []
    for clave, grupo in grupos.items():
        grafias = Counter(_txt(f.get(campo)) for f in grupo if declarado(f.get(campo)))
        etiqueta = grafias.most_common(1)[0][0] if grafias else SIN_DATO
        salida.append((etiqueta, grupo))
    if nivel == "bloque":
        salida.sort(key=lambda par: _clave_orden(par[0]))
    else:
        salida.sort(key=lambda par: (par[0] == SIN_DATO, _norm(par[0])))
    return salida


def tabla_resumen(fichas, nivel):
    """Una fila por grupo con los indicadores de las cinco fichas.

    Cada indicador numerico viaja con `<clave>_n`: los bloques del grupo que
    lo declaran. Un promedio sobre 3 de 14 bloques no es el promedio del
    grupo, y la tabla lo dice en lugar de disimularlo.
    """
    filas = []
    for etiqueta, grupo in agrupar(fichas, nivel):
        areas = [f["area_ha"] for f in grupo if f.get("area_ha") is not None]
        fila = OrderedDict()
        fila["grupo"] = etiqueta
        fila["n_bloques"] = len(grupo)
        fila["area_ha"] = round(sum(areas), 2) if areas else None
        fila["area_ha_n"] = len(areas)

        for clave, _etq, _unidad, agg in INDICADORES_NUM:
            valores = [f.get(clave + "_num") for f in grupo
                       if f.get(clave + "_num") is not None]
            if not valores:
                fila[clave] = None
            elif agg == "suma":
                fila[clave] = round(sum(valores), 2)
            else:
                fila[clave] = round(sum(valores) / len(valores), 2)
            fila[clave + "_n"] = len(valores)

        for clave, _etq in INDICADORES_CONTEO:
            fila[clave] = sum(int(f.get(clave) or 0) for f in grupo)

        for clave, _etq in INDICADORES_SINO:
            respuestas = [f.get(clave + "_si") for f in grupo
                          if f.get(clave + "_si") is not None]
            con_si = sum(1 for r in respuestas if r)
            fila[clave + "_si"] = con_si
            fila[clave + "_n"] = len(respuestas)
            fila[clave + "_pct"] = (round(100 * con_si / len(respuestas), 1)
                                    if respuestas else None)
        filas.append(fila)
    return filas


def indicadores_largos(fila, incluir_conteos=True):
    """Convierte una fila de `tabla_resumen` en pares legibles.

    Es la forma en que el aplicativo y el PDF presentan un grupo: una linea
    por indicador, con su unidad y los bloques que lo sustentan.
    """
    salida = []
    for clave, etiqueta, unidad, agg in INDICADORES_NUM:
        salida.append({
            "Indicador": etiqueta,
            "Unidad": unidad,
            "Agregacion": "Suma" if agg == "suma" else "Promedio",
            "Valor": fila.get(clave),
            "Bloques con dato": fila.get(clave + "_n", 0),
        })
    if incluir_conteos:
        for clave, etiqueta in INDICADORES_CONTEO:
            salida.append({
                "Indicador": etiqueta,
                "Unidad": "n",
                "Agregacion": "Suma",
                "Valor": fila.get(clave),
                "Bloques con dato": fila.get("n_bloques", 0),
            })
    return salida


def tabla_sino(fichas, nivel):
    """Observaciones de si/no: bloques con "Si" y su porcentaje declarado."""
    filas = []
    for fila in tabla_resumen(fichas, nivel):
        for clave, etiqueta in INDICADORES_SINO:
            filas.append({
                "Grupo": fila["grupo"],
                "Observacion": etiqueta,
                "Bloques con Si": fila.get(clave + "_si", 0),
                "Bloques que responden": fila.get(clave + "_n", 0),
                "% de los que responden": fila.get(clave + "_pct"),
            })
    return filas


def distribucion(fichas, campo, nivel="total"):
    """Reparto de frecuencias de una variable categorica.

    Devuelve [{Grupo, Categoria, Bloques, % del grupo}] ordenado de mayor a
    menor dentro de cada grupo, con "Sin registro en ficha" al final.
    """
    filas = []
    for etiqueta, grupo in agrupar(fichas, nivel):
        conteo = Counter(categoria(f, campo) for f in grupo)
        total = sum(conteo.values()) or 1
        ordenado = sorted(conteo.items(),
                          key=lambda par: (par[0] == SIN_DATO, -par[1], par[0]))
        for valor, n in ordenado:
            filas.append({
                "Grupo": etiqueta,
                "Categoria": valor,
                "Bloques": n,
                "% del grupo": round(100 * n / total, 1),
            })
    return filas


def inventario(fichas, tipo, nivel="bloque"):
    """Filas de un inventario tabular con su contexto territorial.

    El contexto (bloque, distrito, provincia) viaja en cada fila para que la
    misma tabla sirva a cualquier nivel de agregacion.
    """
    if tipo not in _INVENTARIOS_POR_TIPO:
        raise ValueError(f"Inventario no reconocido: {tipo}.")
    _tipo, _etiqueta, _clave, _acreditan, columnas = _INVENTARIOS_POR_TIPO[tipo]
    filas = []
    for ficha in fichas:
        base = OrderedDict((
            ("Bloque", ficha.get("codigo_bloque", "")),
            ("Distrito", ficha.get("distrito") or SIN_DATO),
            ("Provincia", ficha.get("provincia") or SIN_DATO),
        ))
        for registro in ficha.get(f"inv_{tipo}", []):
            fila = OrderedDict(base)
            for clave, cabecera in columnas:
                fila[cabecera] = _txt(registro.get(clave))
            filas.append(fila)
    return filas


def ranking(fichas, tipo, campo, limite=15, filtro=None):
    """Frecuencia de un campo de un inventario en todos los bloques.

    `filtro` es un invocable f(registro) -> bool para restringir las filas
    (por ejemplo, las causas cuya presencia esta declarada como "Si").
    Devuelve [{Valor, Registros, Bloques}] de mayor a menor.
    """
    if tipo not in _INVENTARIOS_POR_TIPO:
        raise ValueError(f"Inventario no reconocido: {tipo}.")
    registros = Counter()
    bloques = {}
    for ficha in fichas:
        for fila in ficha.get(f"inv_{tipo}", []):
            if filtro and not filtro(fila):
                continue
            valor = _txt(fila.get(campo))
            if not declarado(valor):
                continue
            registros[valor] += 1
            bloques.setdefault(valor, set()).add(ficha.get("codigo_bloque"))
    orden = sorted(registros.items(), key=lambda par: (-par[1], par[0]))
    return [{"Valor": v, "Registros": n, "Bloques": len(bloques[v])}
            for v, n in orden[:limite]]


def causas_presentes(fichas, limite=20):
    """Causas directas de degradacion declaradas como presentes (F-DT-04)."""
    return ranking(fichas, "causas", "causa", limite=limite,
                   filtro=lambda f: es_si(f.get("presencia")) is True)


def cobertura_declarada(fichas):
    """Bloques que declaran cada indicador numerico, para leer la analitica.

    Sirve para advertir en pantalla que un promedio se sostiene sobre pocos
    bloques antes de que alguien lo cite en el estudio.
    """
    total = len(fichas) or 1
    salida = []
    for clave, etiqueta, unidad, _agg in INDICADORES_NUM:
        n = sum(1 for f in fichas if f.get(clave + "_num") is not None)
        salida.append({
            "Indicador": etiqueta,
            "Unidad": unidad,
            "Bloques con dato": n,
            "% de los bloques": round(100 * n / total, 1),
        })
    salida.sort(key=lambda d: -d["Bloques con dato"])
    return salida


def resumen_general(fichas):
    """Cifras de cabecera del conjunto cargado."""
    total = tabla_resumen(fichas, "total")
    fila = total[0] if total else {}
    return {
        "bloques": len(fichas),
        "distritos": len(agrupar(fichas, "distrito")),
        "provincias": len(agrupar(fichas, "provincia")),
        "area_ha": fila.get("area_ha"),
        "carcavas": fila.get("n_carcavas_inventario", 0),
        "taxones": fila.get("n_taxones", 0),
        "especies_clave": fila.get("n_especies_clave", 0),
        "especies_amenazadas": fila.get("n_especies_amenazadas", 0),
        "fuentes_agua": fila.get("n_fuentes_agua", 0),
        "causas_presentes": fila.get("n_causas_presentes", 0),
        "cobertura_media": fila.get("dt03_cobertura_total"),
        "cobertura_media_n": fila.get("dt03_cobertura_total_n", 0),
        "suelo_desnudo_medio": fila.get("dt03_suelo_desnudo"),
        "suelo_desnudo_medio_n": fila.get("dt03_suelo_desnudo_n", 0),
    }


# ══════════════════════════════════════════════════════════════════════════
# Presentacion: tabla compacta y textos de cabecera
# ══════════════════════════════════════════════════════════════════════════

# Indicadores de cabecera: los que abren el reporte y alimentan los
# graficos. (clave, etiqueta, etiqueta breve, clave del n que lo sustenta).
# La etiqueta breve es para el PDF y la tabla en pantalla, donde el ancho de
# columna no da para el nombre completo del indicador.
CABECERA_COMPACTA = (
    ("n_bloques", "N.o de bloques", "Bloques", None),
    ("area_ha", "Superficie de catalogo (ha)", "Area (ha)", "area_ha_n"),
    ("dt03_cobertura_total", "Cobertura vegetal total (%)", "Cobertura (%)",
     "dt03_cobertura_total_n"),
    ("dt03_suelo_desnudo", "Suelo desnudo (%)", "Suelo desn. (%)",
     "dt03_suelo_desnudo_n"),
    ("dt02_pct_bloque_carcavas", "Bloque afectado por carcavas (%)",
     "Carcavas (% bl.)", "dt02_pct_bloque_carcavas_n"),
    ("n_carcavas_inventario", "Carcavas inventariadas", "Carcavas (n)", None),
    ("n_taxones", "Taxones registrados", "Taxones", None),
    ("n_especies_clave", "Especies clave", "Esp. clave", None),
    ("n_fuentes_agua", "Fuentes de agua", "Fuentes", None),
    ("n_causas_presentes", "Causas presentes", "Causas", None),
)

# Categoricas que abren el reporte cuando no se pide una seleccion propia.
CATEGORICOS_DESTACADOS = (
    "dt02_nivel_erosion_sintesis",
    "dt04_urgencia_intervencion",
    "dt04_velocidad_degradacion",
    "dt03_estado_conservacion_eco",
    "dt03_regeneracion_natural",
    "dt05_transitabilidad_lluviosa",
)

NOTA_METODOLOGICA = (
    "Fuente: 117 fichas F-DT-01 a F-DT-05 (plantilla V5) del Diagnostico "
    "Territorial, leidas con el mismo lector con que el aplicativo importa "
    "una ficha. Ningun valor ausente se estima: los campos declarados como "
    "'Por determinar' o 'Por verificar' y las celdas vacias quedan fuera de "
    "promedios y conteos, y cada indicador informa cuantos bloques lo "
    "sustentan. Los promedios son simples entre los bloques que declaran el "
    "dato, no ponderados por superficie. La provincia y el distrito de cada "
    "bloque se toman del catalogo de bloques del aplicativo; cuando el "
    "catalogo no los tiene, de lo declarado en la ficha. Sistema de "
    "referencia UTM WGS 84 Zona 17S (EPSG:32717)."
)


def etiquetas_indicador():
    """{clave: (etiqueta, unidad)} de todo el catalogo de indicadores."""
    salida = {c: (e, u) for c, e, u, _a in INDICADORES_NUM}
    salida.update({c: (e, "n") for c, e in INDICADORES_CONTEO})
    salida.update({c: (e, "si/no") for c, e in INDICADORES_SINO})
    return salida


def tabla_compacta(fichas, nivel, breve=False):
    """(cabeceras, filas) de los indicadores de cabecera por grupo.

    `breve` devuelve las cabeceras cortas, para el PDF y la pantalla.
    """
    cabeceras = [ETIQUETA_NIVEL[nivel] if nivel != "total" else "Ambito"]
    cabeceras += [corta if breve else etq
                  for _c, etq, corta, _n in CABECERA_COMPACTA]
    filas = []
    for fila in tabla_resumen(fichas, nivel):
        registro = [fila["grupo"]]
        for clave, _etq, _corta, _clave_n in CABECERA_COMPACTA:
            registro.append(fila.get(clave))
        filas.append(registro)
    return cabeceras, filas


# ══════════════════════════════════════════════════════════════════════════
# Exportacion a Excel con graficos nativos
# ══════════════════════════════════════════════════════════════════════════

def _hoja(wb, nombre):
    """Crea una hoja con nombre valido y unico (Excel admite 31 caracteres)."""
    limpio = re.sub(r"[\\/*?:\[\]]", "-", nombre)[:31]
    base, i = limpio, 2
    while limpio in wb.sheetnames:
        sufijo = f" {i}"
        limpio = base[:31 - len(sufijo)] + sufijo
        i += 1
    return wb.create_sheet(limpio)


def _escribir_tabla_dicts(ws, fila, titulo, filas, columnas=None):
    """Escribe una lista de diccionarios como tabla ANIN."""
    if not filas:
        celda = ws.cell(fila, 1, f"{titulo}: sin registros.")
        celda.font = celda.font.copy(italic=True)
        return None, fila + 2
    columnas = columnas or list(filas[0].keys())
    datos = [[f.get(c) for c in columnas] for f in filas]
    return _escribir_bloque_datos(ws, fila, titulo, columnas, datos)


def generar_excel_fdt(fichas, nivel="distrito", categoricos=None,
                      incluir_inventarios=True):
    """Libro Excel con la analitica de las fichas F-DT al nivel pedido.

    Hojas: resumen del nivel con graficos, indicadores detallados,
    observaciones de si/no, distribuciones categoricas, rankings de los
    inventarios y, si se piden, los inventarios completos con su contexto
    territorial.
    """
    if nivel not in NIVELES:
        raise ValueError(f"Nivel no reconocido: {nivel}. Use {NIVELES}.")
    categoricos = categoricos or list(CATEGORICOS_DESTACADOS)
    wb = Workbook()
    wb.remove(wb.active)
    resumen = tabla_resumen(fichas, nivel)

    # ── Hoja 1: resumen del nivel ──
    ws = _hoja(wb, f"Resumen {ETIQUETA_NIVEL[nivel]}")
    fila = _titulo_hoja(ws, f"FICHAS F-DT - {ETIQUETA_NIVEL[nivel].upper()}",
                        ancho=11)
    general = resumen_general(fichas)
    fila_cab, fila = _escribir_bloque_datos(
        ws, fila, "A. Ambito de la lectura",
        ["Concepto", "Valor"],
        [["Bloques con ficha F-DT leida", general["bloques"]],
         ["Distritos representados", general["distritos"]],
         ["Provincias representadas", general["provincias"]],
         ["Carcavas inventariadas", general["carcavas"]],
         ["Taxones del elenco floristico", general["taxones"]],
         ["Especies clave", general["especies_clave"]],
         ["Especies clave con categoria de amenaza",
          general["especies_amenazadas"]],
         ["Fuentes de agua inventariadas", general["fuentes_agua"]],
         ["Causas de degradacion presentes", general["causas_presentes"]]])
    fila += 1

    cabeceras, filas_compactas = tabla_compacta(fichas, nivel)
    fila_cab, fila = _escribir_bloque_datos(
        ws, fila, f"B. Indicadores de cabecera por {ETIQUETA_NIVEL[nivel].lower()}",
        cabeceras, filas_compactas)
    n = len(filas_compactas)
    _agregar_grafico_barras(ws, "Cobertura vegetal total (%)", fila_cab, n,
                            1, 4, f"A{fila + 1}", eje_y="%", ancho=18, alto=8)
    _agregar_grafico_barras(ws, "Suelo desnudo (%)", fila_cab, n,
                            1, 5, f"A{fila + 18}", eje_y="%", ancho=18, alto=8)
    _agregar_grafico_barras(ws, "Carcavas inventariadas", fila_cab, n,
                            1, 7, f"A{fila + 35}", eje_y="n", ancho=18, alto=8)
    fila += 52
    ws.cell(fila, 1, NOTA_METODOLOGICA)
    ws.column_dimensions["A"].width = 34
    for letra in "BCDEFGHIJK":
        ws.column_dimensions[letra].width = 16

    # ── Hoja 2: indicadores detallados ──
    ws = _hoja(wb, "Indicadores detallados")
    fila = _titulo_hoja(ws, "INDICADORES POR FICHA F-DT", ancho=6)
    detalle = []
    for registro in resumen:
        for item in indicadores_largos(registro):
            detalle.append({ETIQUETA_NIVEL[nivel]: registro["grupo"], **item})
    _escribir_tabla_dicts(ws, fila, "Indicadores declarados", detalle)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 40
    for letra in "CDEF":
        ws.column_dimensions[letra].width = 16

    # ── Hoja 3: observaciones de si/no ──
    ws = _hoja(wb, "Observaciones Si-No")
    fila = _titulo_hoja(ws, "OBSERVACIONES DE CAMPO (SI / NO)", ancho=5)
    _escribir_tabla_dicts(ws, fila, "Bloques que declaran 'Si'",
                          tabla_sino(fichas, nivel))
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 40
    for letra in "CDE":
        ws.column_dimensions[letra].width = 18

    # ── Hoja 4: distribuciones categoricas ──
    ws = _hoja(wb, "Distribuciones")
    fila = _titulo_hoja(ws, "DISTRIBUCIONES POR CATEGORIA", ancho=5)
    etiquetas = dict(CATEGORICOS)
    for campo in categoricos:
        filas_dist = distribucion(fichas, campo, nivel)
        fila_cab, fila = _escribir_tabla_dicts(
            ws, fila, etiquetas.get(campo, campo), filas_dist)
        if fila_cab and nivel == "total" and len(filas_dist) <= 12:
            _agregar_grafico_torta(ws, etiquetas.get(campo, campo), fila_cab,
                                   len(filas_dist), 2, 3, f"G{fila_cab}",
                                   ancho=11, alto=7)
        fila += 2
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 46
    for letra in "CD":
        ws.column_dimensions[letra].width = 14

    # ── Hoja 5: rankings de los inventarios ──
    ws = _hoja(wb, "Rankings")
    fila = _titulo_hoja(ws, "RANKINGS DE LOS INVENTARIOS", ancho=4)
    for titulo, filas_rank in (
            ("Causas de degradacion presentes (F-DT-04)",
             causas_presentes(fichas, limite=20)),
            ("Especies mas frecuentes del elenco floristico (F-DT-03)",
             ranking(fichas, "floristica", "nombre_comun", limite=20)),
            ("Familias botanicas mas frecuentes (F-DT-03)",
             ranking(fichas, "floristica", "familia", limite=15)),
            ("Especies clave por categoria (F-DT-03)",
             ranking(fichas, "especies_clave", "categoria", limite=12)),
            ("Tipos de fuente de agua (F-DT-05)",
             ranking(fichas, "fuentes_agua", "tipo", limite=12)),
            ("Regimen de las fuentes de agua (F-DT-05)",
             ranking(fichas, "fuentes_agua", "regimen", limite=10)),
            ("Tipos de carcava inventariados (F-DT-02)",
             ranking(fichas, "carcavas", "tipo", limite=10)),
            ("Indicadores cuantitativos registrados (F-DT-04)",
             ranking(fichas, "indicadores", "indicador", limite=15))):
        _, fila = _escribir_tabla_dicts(ws, fila, titulo, filas_rank)
        fila += 2
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14

    # ── Hojas 6+: inventarios completos ──
    if incluir_inventarios:
        for tipo, etiqueta, _clave, _acreditan, _cols in INVENTARIOS:
            filas_inv = inventario(fichas, tipo)
            ws = _hoja(wb, f"Inv. {etiqueta.split(' (')[0]}")
            fila = _titulo_hoja(ws, etiqueta.upper(), ancho=9)
            _escribir_tabla_dicts(ws, fila, etiqueta, filas_inv)
            ws.column_dimensions["A"].width = 12
            ws.column_dimensions["B"].width = 20
            ws.column_dimensions["C"].width = 16
            for letra in "DEFGHIJK":
                ws.column_dimensions[letra].width = 20

    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue()


# ══════════════════════════════════════════════════════════════════════════
# Exportacion a PDF institucional
# ══════════════════════════════════════════════════════════════════════════

# Un grafico de barras con 117 categorias no se lee. Cuando el nivel tiene
# mas grupos que este limite, el grafico muestra los primeros por valor y la
# tabla completa queda en el propio PDF.
MAX_BARRAS_PDF = 20


def _top_para_grafico(filas_compactas, columna, limite=MAX_BARRAS_PDF):
    """(etiquetas, valores) de los grupos con mayor valor en una columna."""
    datos = [(f[0], f[columna]) for f in filas_compactas
             if isinstance(f[columna], (int, float))]
    datos.sort(key=lambda par: -par[1])
    datos = datos[:limite]
    return [d[0] for d in datos], [d[1] for d in datos]


def generar_pdf_fdt(fichas, nivel="distrito", categoricos=None,
                    incluir_graficos=True):
    """Reporte PDF institucional de las fichas F-DT al nivel pedido."""
    if nivel not in NIVELES:
        raise ValueError(f"Nivel no reconocido: {nivel}. Use {NIVELES}.")
    categoricos = categoricos or list(CATEGORICOS_DESTACADOS)
    etiquetas = dict(CATEGORICOS)
    general = resumen_general(fichas)

    pdf = _PDFResumen(
        subtitulo=f"Diagnostico Territorial - Fichas F-DT | "
                  f"{ETIQUETA_NIVEL[nivel]}",
        orientacion="L")
    pdf.alias_nb_pages()
    pdf.add_page()

    pdf.seccion("A. AMBITO DE LA LECTURA")
    pdf.campos([
        ("Bloques con ficha F-DT", general["bloques"]),
        ("Distritos representados", general["distritos"]),
        ("Provincias representadas", general["provincias"]),
        ("Cobertura vegetal total (promedio)",
         f"{_fmt(general['cobertura_media'])} % "
         f"({general['cobertura_media_n']} bloques declaran el dato)"),
        ("Suelo desnudo (promedio)",
         f"{_fmt(general['suelo_desnudo_medio'])} % "
         f"({general['suelo_desnudo_medio_n']} bloques declaran el dato)"),
        ("Carcavas inventariadas", general["carcavas"]),
        ("Taxones del elenco floristico", general["taxones"]),
        ("Especies clave (con amenaza)",
         f"{general['especies_clave']} ({general['especies_amenazadas']})"),
        ("Fuentes de agua inventariadas", general["fuentes_agua"]),
        ("Causas de degradacion presentes", general["causas_presentes"]),
    ], columnas=2)

    cabeceras, filas_compactas = tabla_compacta(fichas, nivel, breve=True)
    pdf.seccion(f"B. INDICADORES DE CABECERA POR "
                f"{ETIQUETA_NIVEL[nivel].upper()}")
    pdf.tabla(
        cabeceras,
        [[f[0]] + [_fmt(v) for v in f[1:]] for f in filas_compactas],
        anchos_rel=[2.2] + [1.1] * (len(cabeceras) - 1),
        alineaciones=["L"] + ["R"] * (len(cabeceras) - 1),
        tam=6.5)
    pdf.nota(
        "El promedio de cada indicador se calcula solo sobre los bloques que "
        "lo declaran; la hoja 'Indicadores detallados' del Excel informa "
        "cuantos son en cada caso.")

    if incluir_graficos and filas_compactas:
        pdf.add_page()
        pdf.seccion("C. GRAFICOS DEL NIVEL")
        for columna, titulo, unidad in (
                (3, "Cobertura vegetal total (%)", "%"),
                (4, "Suelo desnudo (%)", "%"),
                (6, "Carcavas inventariadas", "n"),
                (7, "Taxones registrados", "n")):
            etqs, vals = _top_para_grafico(filas_compactas, columna)
            if not vals:
                continue
            sufijo = (f" - {len(etqs)} primeros de {len(filas_compactas)}"
                      if len(filas_compactas) > len(etqs) else "")
            pdf.grafico_barras(f"{titulo}{sufijo}", etqs, vals, unidad=unidad,
                               horizontal=len(etqs) > 8)

    pdf.add_page()
    pdf.seccion("D. DISTRIBUCIONES POR CATEGORIA")
    for campo in categoricos:
        filas_dist = distribucion(fichas, campo, "total")
        if not filas_dist:
            continue
        titulo = etiquetas.get(campo, campo)
        if incluir_graficos and len(filas_dist) <= 10:
            pdf.grafico_torta(titulo, [f["Categoria"] for f in filas_dist],
                              [f["Bloques"] for f in filas_dist])
        else:
            pdf.tabla(["Categoria", "Bloques", "% del total"],
                      [[f["Categoria"], f["Bloques"], _fmt(f["% del grupo"])]
                       for f in filas_dist],
                      anchos_rel=[4, 1, 1], alineaciones=["L", "R", "R"])

    pdf.add_page()
    pdf.seccion("E. RANKINGS DE LOS INVENTARIOS")
    for titulo, filas_rank in (
            ("Causas de degradacion presentes (F-DT-04)",
             causas_presentes(fichas, limite=15)),
            ("Especies mas frecuentes del elenco floristico (F-DT-03)",
             ranking(fichas, "floristica", "nombre_comun", limite=15)),
            ("Especies clave por categoria (F-DT-03)",
             ranking(fichas, "especies_clave", "categoria", limite=10)),
            ("Tipos de fuente de agua (F-DT-05)",
             ranking(fichas, "fuentes_agua", "tipo", limite=10))):
        if not filas_rank:
            continue
        pdf.tabla(["Registro", "Filas", "Bloques"],
                  [[_recortar(f["Valor"], 78), f["Registros"], f["Bloques"]]
                   for f in filas_rank],
                  anchos_rel=[6, 1, 1], alineaciones=["L", "R", "R"])
        pdf.nota(titulo)

    pdf.seccion("F. NOTA METODOLOGICA")
    pdf.nota(NOTA_METODOLOGICA)
    return _pdf_bytes(pdf)
