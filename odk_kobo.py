"""
IN Piura - Modulo de Integracion ODK / KoBoToolbox
Verificacion de campo de bloques de intervencion (Paso 6 - AdR-CCC territorial).

- Genera el formulario XLSForm (ODK Collect / KoBoCollect / Enketo) con listas
  cerradas de distritos y de los bloques del catalogo.
- Importa envios desde la API v2 de KoBoToolbox (paginada) o desde un archivo
  CSV / Excel exportado.

Reglas de importacion (no destructivas):
  * NUNCA crea, modifica ni borra bloques: el maestro de bloques validado en
    gabinete queda intacto. Un codigo que no existe en el catalogo se rechaza.
  * Cada envio se registra una sola vez (clave = _uuid de KoBo o huella del
    contenido). Reimportar no duplica.
  * La coordenada se calcula del GPS del dispositivo (UTM WGS84 17S,
    EPSG:32717) y se valida contra el rango del ambito del proyecto. La
    coordenada digitada a mano queda solo como referencia.
  * Las fotos se descargan de `_attachments` con el token (version mediana).
  * La verificacion SSL esta siempre activa.

Cuenca Alta del Rio Piura, Peru.
"""

try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
    _HAS_TK = True
except ImportError:
    tk = None
    ttk = None
    _HAS_TK = False

from datetime import datetime
import csv
import hashlib
import json
import math
import os
import re
import ssl
import unicodedata
import urllib.error
import urllib.parse
import urllib.request

import database as db

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    OPENPYXL_DISPONIBLE = True
except ImportError:
    OPENPYXL_DISPONIBLE = False

try:
    from pyproj import Transformer
    _TRANSFORMER_17S = Transformer.from_crs("EPSG:4326", "EPSG:32717",
                                            always_xy=True)
except Exception:  # pragma: no cover - pyproj esta en requirements.txt
    _TRANSFORMER_17S = None


# ══════════════════════════════════════════════════════════════════════════
# Parametros del ambito
# ══════════════════════════════════════════════════════════════════════════

FORM_ID = "in_piura_verificacion_paso6"
FORM_ID_ANTERIOR = "in_piura_verificacion_campo"

# Rango valido UTM WGS84 17S para el ambito del proyecto (estandar ANIN).
UTM_ESTE_MIN, UTM_ESTE_MAX = 450_000.0, 750_000.0
UTM_NORTE_MIN, UTM_NORTE_MAX = 9_300_000.0, 9_600_000.0

# Rango equivalente aproximado en lat/lon, usado como filtro en el formulario
# (el control estricto se hace en UTM al importar).
LAT_MIN, LAT_MAX = -6.35, -3.60
LON_MIN, LON_MAX = -81.50, -78.70

MAX_PREDIOS_BLOQUE = 19            # Paso 5 - tamizaje predial
DIST_ALERTA_CENTROIDE_M = 2_500    # punto GPS lejos del centroide del bloque
DIF_ALERTA_GPS_MANUAL_M = 100      # GPS del celular vs GPS navegador
PRECISION_ALERTA_M = 20

PAGINA_API = 1000                  # maximo permitido por KoBo desde 03/2026

DISTRITOS_IN_PIURA = [
    "Frías", "Canchaque", "Huarmaca", "Huancabamba", "Lalaquiz",
    "San Miguel de El Faique", "Buenos Aires", "Chalaco", "Chulucanas",
    "Morropón", "Salitral", "San Juan de Bigote", "Santa Catalina de Mossa",
    "Santo Domingo", "Yamango",
]


def _slug(texto):
    """'San Miguel de El Faique' -> 'san_miguel_de_el_faique'."""
    t = unicodedata.normalize("NFKD", str(texto or ""))
    t = "".join(c for c in t if not unicodedata.combining(c)).lower()
    return re.sub(r"[^a-z0-9]+", "_", t).strip("_")


def nombre_opcion_bloque(codigo):
    """Nombre de opcion XLSForm para un codigo de bloque (sin espacios)."""
    c = str(codigo or "").strip()
    return c if re.fullmatch(r"[A-Za-z0-9_.-]+", c) else _slug(c)


# ══════════════════════════════════════════════════════════════════════════
# Listas de opciones (choices)
# ══════════════════════════════════════════════════════════════════════════

CHOICES = {
    "si_no": [("si", "Sí"), ("no", "No")],
    "clima": [
        ("despejado", "Despejado"), ("parcialmente_nublado", "Parcialmente nublado"),
        ("nublado", "Nublado"), ("lluvia_ligera", "Lluvia ligera"),
        ("lluvia_moderada", "Lluvia moderada"), ("lluvia_intensa", "Lluvia intensa"),
        ("neblina", "Neblina"),
    ],
    "accesibilidad": [
        ("accesible", "Accesible"),
        ("accesible_limitaciones", "Accesible con limitaciones"),
        ("no_accesible", "No accesible"),
    ],
    "tipo_acceso": [
        ("carretera_asfaltada", "Carretera asfaltada"),
        ("carretera_afirmada", "Carretera afirmada"),
        ("trocha_carrozable", "Trocha carrozable"),
        ("camino_herradura", "Camino de herradura"),
        ("sendero", "Sendero peatonal"),
        ("sin_acceso", "Sin acceso"),
    ],
    "estacionalidad": [
        ("todo_anio", "Todo el año"),
        ("solo_estiaje", "Solo en estiaje (época seca)"),
        ("intermitente", "Intermitente"),
    ],
    "nivel": [("bajo", "Bajo"), ("medio", "Medio"), ("alto", "Alto")],
    "factor_social": [
        ("conflicto_tierras", "Conflicto por tierras / linderos"),
        ("oposicion_comunal", "Oposición comunal o de ronda"),
        ("conflicto_agua", "Conflicto por uso de agua"),
        ("mineria_informal", "Minería informal"),
        ("inseguridad", "Inseguridad / delincuencia"),
        ("expectativas", "Expectativas no atendidas (empleo, pagos)"),
        ("otro", "Otro"),
    ],
    "uso_suelo": [
        ("bosque_seco", "Bosque seco"),
        ("bosque_montano", "Bosque montano"),
        ("matorral", "Matorral"),
        ("pastizal_natural", "Pastizal natural"),
        ("pastoreo", "Pastoreo"),
        ("agricola_activo", "Agrícola activo"),
        ("agricola_descanso", "Agrícola en descanso"),
        ("suelo_desnudo", "Suelo desnudo / erosionado"),
        ("plantacion_forestal", "Plantación forestal"),
        ("infraestructura", "Vivienda / infraestructura"),
    ],
    "tipo_mm": [
        ("deslizamiento", "Deslizamiento"),
        ("flujo_detritos", "Flujo de detritos / huaico"),
        ("caida_rocas", "Caída de rocas"),
        ("reptacion", "Reptación de suelos"),
        ("carcavas", "Cárcavas"),
        ("erosion_surcos", "Erosión en surcos"),
        ("socavamiento", "Socavamiento de márgenes"),
    ],
    "magnitud": [("leve", "Leve"), ("moderada", "Moderada"), ("severa", "Severa")],
    "actividad": [("activo", "Activo"), ("latente", "Latente"), ("inactivo", "Inactivo")],
    "tenencia": [
        ("comunidad_campesina", "Comunidad campesina"),
        ("propiedad_privada", "Propiedad privada"),
        ("posesionario", "Posesionario"),
        ("estado", "Estado"),
        ("no_determinado", "No determinado"),
    ],
    "aceptacion": [
        ("acepta", "Acepta"),
        ("acepta_condiciones", "Acepta con condiciones"),
        ("no_acepta", "No acepta"),
        ("no_consultado", "No consultado"),
    ],
    "dictamen": [
        ("apto", "Apto"),
        ("apto_observaciones", "Apto con observaciones"),
        ("no_apto", "No apto"),
    ],
}

# Campo del formulario -> lista de opciones (para traducir valores a etiquetas)
CAMPO_LISTA = {
    "condiciones_climaticas": "clima", "accesibilidad": "accesibilidad",
    "tipo_acceso": "tipo_acceso", "acceso_estacional": "estacionalidad",
    "riesgo_social": "nivel", "riesgo_social_factores": "factor_social",
    "uso_suelo_predominante": "uso_suelo", "uso_suelo_otros": "uso_suelo",
    "agricultura_activa": "si_no", "evidencia_mm": "si_no",
    "tipos_mm": "tipo_mm", "magnitud_mm": "magnitud", "actividad_mm": "actividad",
    "tenencia": "tenencia", "aceptacion_titular": "aceptacion",
    "acta_firmada": "si_no", "dictamen": "dictamen",
}

CAMPOS_FOTO = ("foto_panoramica", "foto_detalle", "foto_adicional",
               "foto_mm", "foto_acta")


def _etiqueta(lista, valor):
    """Traduce el/los valores de una opcion a su etiqueta legible."""
    if valor in (None, ""):
        return ""
    mapa = {n: l for n, l in CHOICES.get(lista, [])}
    partes = str(valor).split()
    return "; ".join(mapa.get(p, p) for p in partes)


# ══════════════════════════════════════════════════════════════════════════
# Formulario XLSForm - Verificacion de campo (Paso 6)
# ══════════════════════════════════════════════════════════════════════════

SURVEY_COLUMNAS = ["type", "name", "label", "hint", "required",
                   "required_message", "constraint", "constraint_message",
                   "relevant", "appearance", "choice_filter", "calculation",
                   "default", "parameters"]

_REQ = "Campo obligatorio"


def _fila(tipo, nombre="", etiqueta="", **kw):
    fila = {"type": tipo, "name": nombre, "label": etiqueta}
    fila.update(kw)
    if fila.get("required") == "yes" and "required_message" not in fila:
        fila["required_message"] = _REQ
    if tipo == "image":
        # Reduce las fotos a 1280 px en el celular: envio mas rapido con
        # poca senal y menos almacenamiento, con detalle suficiente.
        fila.setdefault("parameters", "max-pixels=1280")
    return fila


def _filas_survey(con_catalogo):
    """Filas de la hoja survey. `con_catalogo` indica si hay lista de bloques."""
    if con_catalogo:
        fila_bloque = _fila(
            "select_one bloque", "codigo_bloque", "Código del bloque",
            hint="Solo se listan los bloques del distrito elegido. Si no aparece, elija 'Otro'.",
            required="yes",
            choice_filter="distrito=${distrito} or name='otro'")
        fila_otro = _fila(
            "text", "codigo_bloque_otro", "Escriba el código del bloque",
            hint="Formato del catálogo, p. ej. M17B4 o 25",
            required="yes", relevant="${codigo_bloque}='otro'",
            constraint="regex(., '^[A-Za-z0-9-]+$')",
            constraint_message="Solo letras, números y guion, sin espacios")
    else:
        fila_bloque = _fila(
            "text", "codigo_bloque", "Código del bloque",
            hint="Formato del catálogo, p. ej. M17B4 o 25", required="yes",
            constraint="regex(., '^[A-Za-z0-9-]+$')",
            constraint_message="Solo letras, números y guion, sin espacios")
        fila_otro = None

    geo_constraint = (
        f"selected-at(., 0) >= {LAT_MIN} and selected-at(., 0) <= {LAT_MAX} and "
        f"selected-at(., 1) >= {LON_MIN} and selected-at(., 1) <= {LON_MAX}")

    filas = [
        _fila("start", "inicio"), _fila("end", "fin"),
        _fila("today", "fecha_hoy"), _fila("deviceid", "id_dispositivo"),
        _fila("username", "usuario"),

        _fila("begin_group", "g_identificacion", "1. Identificación",
              appearance="field-list"),
        _fila("date", "fecha_visita", "Fecha de la visita", required="yes",
              default="today()",
              constraint=". <= today()", constraint_message="La fecha no puede ser futura"),
        _fila("text", "inspector", "Nombre del verificador", required="yes"),
        _fila("text", "brigada", "Brigada / equipo"),
        _fila("select_one distrito", "distrito", "Distrito", required="yes"),
        fila_bloque,
        fila_otro,
        _fila("select_one clima", "condiciones_climaticas",
              "Condiciones climáticas", required="yes"),
        _fila("end_group"),

        _fila("begin_group", "g_ubicacion", "2. Ubicación del punto de verificación",
              appearance="field-list"),
        _fila("geopoint", "ubicacion_gps", "Capture el GPS en el bloque",
              hint="Funciona sin internet. Espere a que la precisión baje de 10 m.",
              required="yes", constraint=geo_constraint,
              constraint_message="El punto está fuera del ámbito del proyecto (Piura)"),
        _fila("calculate", "gps_precision_calc",
              calculation="selected-at(${ubicacion_gps}, 3)"),
        _fila("note", "nota_precision",
              f"Precisión baja (${{gps_precision_calc}} m). Si puede, repita la captura.",
              relevant=f"${{gps_precision_calc}} > {PRECISION_ALERTA_M}"),
        _fila("decimal", "utm_este_ref", "UTM Este leído en GPS navegador (opcional)",
              hint="Solo como referencia; el aplicativo calcula la UTM desde el GPS",
              constraint=f". >= {int(UTM_ESTE_MIN)} and . <= {int(UTM_ESTE_MAX)}",
              constraint_message="Este fuera de rango (450,000 – 750,000 m)"),
        _fila("decimal", "utm_norte_ref", "UTM Norte leído en GPS navegador (opcional)",
              constraint=f". >= {int(UTM_NORTE_MIN)} and . <= {int(UTM_NORTE_MAX)}",
              constraint_message="Norte fuera de rango (9,300,000 – 9,600,000 m)"),
        _fila("end_group"),

        _fila("begin_group", "g_accesibilidad", "3. Accesibilidad real",
              appearance="field-list"),
        _fila("select_one accesibilidad", "accesibilidad", "Accesibilidad al bloque",
              required="yes"),
        _fila("select_one tipo_acceso", "tipo_acceso", "Tipo de vía de acceso",
              required="yes"),
        _fila("integer", "tiempo_acceso_min",
              "Tiempo desde la vía carrozable más cercana (minutos)",
              constraint=". >= 0 and . <= 600",
              constraint_message="Entre 0 y 600 minutos"),
        _fila("select_one estacionalidad", "acceso_estacional",
              "¿En qué época es transitable?", required="yes"),
        _fila("text", "accesibilidad_obs", "Describa las limitaciones de acceso",
              relevant="${accesibilidad}!='accesible'"),
        _fila("end_group"),

        _fila("begin_group", "g_social", "4. Riesgo social", appearance="field-list"),
        _fila("select_one nivel", "riesgo_social", "Nivel de riesgo social",
              required="yes"),
        _fila("select_multiple factor_social", "riesgo_social_factores",
              "Factores de riesgo social", required="yes",
              relevant="${riesgo_social}!='bajo'"),
        _fila("text", "riesgo_social_detalle", "Detalle del riesgo social",
              relevant="${riesgo_social}!='bajo'", appearance="multiline"),
        _fila("end_group"),

        _fila("begin_group", "g_suelo", "5. Uso actual del suelo",
              appearance="field-list"),
        _fila("select_one uso_suelo", "uso_suelo_predominante",
              "Uso predominante del suelo", required="yes"),
        _fila("select_multiple uso_suelo", "uso_suelo_otros",
              "Otros usos presentes en el bloque"),
        _fila("select_one si_no", "agricultura_activa",
              "¿Hay agricultura activa dentro del bloque?",
              hint="Las áreas agrícolas se excluyen (Paso 5)", required="yes"),
        _fila("integer", "cobertura_vegetal_pct", "Cobertura vegetal estimada (%)",
              constraint=". >= 0 and . <= 100", constraint_message="Entre 0 y 100"),
        _fila("end_group"),

        _fila("begin_group", "g_mm", "6. Evidencia de movimientos en masa",
              appearance="field-list"),
        _fila("select_one si_no", "evidencia_mm",
              "¿Hay evidencia de movimientos en masa o erosión?", required="yes"),
        _fila("select_multiple tipo_mm", "tipos_mm", "Tipo de evidencia",
              required="yes", relevant="${evidencia_mm}='si'"),
        _fila("select_one magnitud", "magnitud_mm", "Magnitud",
              required="yes", relevant="${evidencia_mm}='si'"),
        _fila("select_one actividad", "actividad_mm", "Estado de actividad",
              relevant="${evidencia_mm}='si'"),
        _fila("image", "foto_mm", "Foto de la evidencia",
              relevant="${evidencia_mm}='si'"),
        _fila("end_group"),

        _fila("begin_group", "g_titulares", "7. Titulares de predios",
              appearance="field-list"),
        _fila("select_one tenencia", "tenencia", "Tenencia de la tierra",
              required="yes"),
        _fila("integer", "n_predios", "N.° de predios interceptados por el bloque",
              constraint=". >= 0 and . <= 200", constraint_message="Entre 0 y 200"),
        _fila("note", "nota_predios",
              f"Supera el máximo de {MAX_PREDIOS_BLOQUE} predios por bloque (Paso 5). Confirme el dato.",
              relevant=f"${{n_predios}} > {MAX_PREDIOS_BLOQUE}"),
        _fila("text", "titular", "Titular o representante consultado"),
        _fila("select_one aceptacion", "aceptacion_titular",
              "Aceptación del titular", required="yes"),
        _fila("text", "condiciones_aceptacion", "Condiciones planteadas",
              required="yes", relevant="${aceptacion_titular}='acepta_condiciones'",
              appearance="multiline"),
        _fila("select_one si_no", "acta_firmada", "¿Se firmó acta de aceptación?",
              relevant="${aceptacion_titular}='acepta' or ${aceptacion_titular}='acepta_condiciones'"),
        _fila("image", "foto_acta", "Foto del acta", relevant="${acta_firmada}='si'"),
        _fila("end_group"),

        _fila("begin_group", "g_resultado", "8. Resultado de la verificación",
              appearance="field-list"),
        _fila("select_one dictamen", "dictamen", "Dictamen del bloque", required="yes"),
        _fila("text", "motivo_no_apto", "Motivo / observaciones del dictamen",
              required="yes", relevant="${dictamen}!='apto'", appearance="multiline"),
        _fila("text", "observaciones", "Observaciones generales", appearance="multiline"),
        _fila("image", "foto_panoramica", "Foto panorámica del bloque", required="yes"),
        _fila("image", "foto_detalle", "Foto de detalle"),
        _fila("image", "foto_adicional", "Foto adicional"),
        _fila("end_group"),

        _fila("calculate", "codigo_verificacion",
              calculation="concat('VER-', format-date(${fecha_visita}, '%Y%m%d'), "
                          "'-', substr(${id_dispositivo}, string-length(${id_dispositivo}) - 7, "
                          "string-length(${id_dispositivo})))"),
        _fila("note", "nota_verificacion",
              "Código de verificación: ${codigo_verificacion}"),
    ]
    return [f for f in filas if f]


def _opciones_bloques(bloques):
    """(filas de choices para distrito y bloque) desde el catalogo."""
    distritos = {}
    for d in DISTRITOS_IN_PIURA:
        distritos[_slug(d)] = d
    for b in bloques:
        s = _slug(b.get("distrito"))
        if s and s not in distritos:
            distritos[s] = b.get("distrito")
    filas_d = [("distrito", s, l, "") for s, l in sorted(distritos.items(),
                                                        key=lambda x: x[1])]
    filas_d.append(("distrito", "otro", "Otro / no listado", ""))

    filas_b = []
    vistos = set()
    for b in sorted(bloques, key=lambda x: (str(x.get("distrito") or ""),
                                            str(x.get("codigo")))):
        nombre = nombre_opcion_bloque(b.get("codigo"))
        if not nombre or nombre in vistos:
            continue
        vistos.add(nombre)
        partes = [str(b.get("codigo"))]
        if b.get("microcuenca"):
            partes.append(str(b["microcuenca"]))
        try:
            if float(b.get("area_hectareas") or 0) > 0:
                partes.append(f"{float(b['area_hectareas']):,.1f} ha")
        except (TypeError, ValueError):
            pass
        filas_b.append(("bloque", nombre, " · ".join(partes),
                        _slug(b.get("distrito")) or "otro"))
    filas_b.append(("bloque", "otro", "Otro (no está en la lista)", ""))
    return filas_d, filas_b


def generar_xlsform(ruta_salida=None, bloques=None):
    """Genera el XLSForm de verificacion de campo (Paso 6).

    `bloques`: lista de dicts del catalogo (db.obtener_bloques()). Si se
    omite, el codigo del bloque se captura como texto validado.
    Devuelve la ruta del archivo generado.
    """
    if not OPENPYXL_DISPONIBLE:
        raise ImportError("Se requiere openpyxl para generar formularios XLSForm.")

    if ruta_salida is None:
        directorio = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                  "formularios")
        os.makedirs(directorio, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_salida = os.path.join(directorio, f"IN_Piura_Verificacion_Paso6_{ts}.xlsx")

    bloques = [b for b in (bloques or []) if b.get("codigo")]
    con_catalogo = bool(bloques)

    verde = PatternFill(start_color="1B4D2E", fill_type="solid")
    blanco = Font(bold=True, color="FFFFFF", name="Arial")

    wb = Workbook()
    ws = wb.active
    ws.title = "survey"
    for c, h in enumerate(SURVEY_COLUMNAS, 1):
        celda = ws.cell(row=1, column=c, value=h)
        celda.font, celda.fill = blanco, verde
        celda.alignment = Alignment(horizontal="center")
    for r, fila in enumerate(_filas_survey(con_catalogo), 2):
        for c, h in enumerate(SURVEY_COLUMNAS, 1):
            v = fila.get(h, "")
            if v != "":
                ws.cell(row=r, column=c, value=v)
    for letra, ancho in zip("ABCDEFGHIJKLMN", (26, 26, 48, 40, 9, 16, 40, 34, 34, 12, 30, 40, 10, 18)):
        ws.column_dimensions[letra].width = ancho
    ws.freeze_panes = "A2"

    wc = wb.create_sheet("choices")
    for c, h in enumerate(["list_name", "name", "label", "distrito"], 1):
        celda = wc.cell(row=1, column=c, value=h)
        celda.font, celda.fill = blanco, verde
    filas_c = [(l, n, e, "") for l, ops in CHOICES.items() for n, e in ops]
    filas_d, filas_b = _opciones_bloques(bloques)
    filas_c += filas_d
    if con_catalogo:
        filas_c += filas_b
    for r, (l, n, e, d) in enumerate(filas_c, 2):
        wc.cell(row=r, column=1, value=l)
        wc.cell(row=r, column=2, value=n)
        wc.cell(row=r, column=3, value=e)
        if d:
            wc.cell(row=r, column=4, value=d)
    for letra, ancho in zip("ABCD", (18, 26, 44, 22)):
        wc.column_dimensions[letra].width = ancho
    wc.freeze_panes = "A2"

    wset = wb.create_sheet("settings")
    ajustes = {
        "form_title": "IN Piura - Verificación de campo de bloques (Paso 6)",
        "form_id": FORM_ID,
        "version": datetime.now().strftime("%Y%m%d%H%M"),
        "instance_name": "concat(${codigo_bloque}, ' - ', ${fecha_visita})",
    }
    for c, (k, v) in enumerate(ajustes.items(), 1):
        celda = wset.cell(row=1, column=c, value=k)
        celda.font, celda.fill = blanco, verde
        wset.cell(row=2, column=c, value=v)
        wset.column_dimensions[chr(64 + c)].width = 44

    wb.save(ruta_salida)
    return ruta_salida


def encabezados_plantilla_csv():
    """Columnas del CSV equivalente al formulario (para carga manual)."""
    cols = []
    for f in _filas_survey(True):
        if f["type"] in ("begin_group", "end_group", "note", "start", "end",
                         "today", "deviceid", "username") or not f.get("name"):
            continue
        if f["name"] in ("gps_precision_calc",):
            continue
        cols.append(f["name"])
    return cols


# ══════════════════════════════════════════════════════════════════════════
# Coordenadas
# ══════════════════════════════════════════════════════════════════════════

def latlon_a_utm17s(lat, lon):
    """Lat/Lon WGS84 -> (Este, Norte) UTM 17S (EPSG:32717) con pyproj."""
    if _TRANSFORMER_17S is not None:
        e, n = _TRANSFORMER_17S.transform(lon, lat)
        return float(e), float(n)
    from georeferenciacion import latlon_a_utm  # respaldo (formula exacta)
    e, n, _ = latlon_a_utm(lat, lon, 17)
    return float(e), float(n)


def utm_en_rango(este, norte):
    try:
        return (UTM_ESTE_MIN <= float(este) <= UTM_ESTE_MAX and
                UTM_NORTE_MIN <= float(norte) <= UTM_NORTE_MAX)
    except (TypeError, ValueError):
        return False


def _num(valor):
    if valor in (None, ""):
        return None
    try:
        v = float(str(valor).replace(",", "."))
        return v if math.isfinite(v) else None
    except (TypeError, ValueError):
        return None


def _parsear_geopoint(datos):
    """(lat, lon, alt, precision) desde 'lat lon alt prec' o columnas separadas."""
    gp = datos.get("ubicacion_gps")
    if isinstance(gp, str) and gp.strip():
        p = gp.split()
        vals = [_num(x) for x in p[:4]] + [None] * (4 - len(p[:4]))
        if vals[0] is not None and vals[1] is not None:
            return tuple(vals)
    lat = _num(datos.get("_ubicacion_gps_latitude"))
    lon = _num(datos.get("_ubicacion_gps_longitude"))
    if lat is not None and lon is not None:
        return (lat, lon, _num(datos.get("_ubicacion_gps_altitude")),
                _num(datos.get("_ubicacion_gps_precision")))
    return (None, None, None, None)


# ══════════════════════════════════════════════════════════════════════════
# Preparacion de envios (sin escribir en la BD)
# ══════════════════════════════════════════════════════════════════════════

def _aplanar(envio):
    """Quita prefijos de grupo ('g_identificacion/inspector' -> 'inspector')."""
    datos = {}
    for k, v in (envio or {}).items():
        clave = str(k).split("/")[-1].strip()
        if isinstance(v, str):
            v = v.strip()
        if clave not in datos or datos[clave] in (None, ""):
            datos[clave] = v
    return datos


def _clave_envio(datos):
    """_uuid de KoBo; si no hay (CSV manual), huella estable del contenido."""
    for k in ("_uuid", "instanceID", "uuid"):
        v = str(datos.get(k) or "").strip()
        if v:
            return v.replace("uuid:", "")
    base = {k: str(v) for k, v in datos.items()
            if not str(k).startswith("_") and v not in (None, "")}
    huella = hashlib.sha1(json.dumps(base, sort_keys=True,
                                     ensure_ascii=False).encode()).hexdigest()
    return f"hash:{huella}"


class CatalogoBloques:
    """Busqueda de bloques por codigo. Solo lectura."""

    def __init__(self, bloques_activos=None):
        if bloques_activos is None:
            bloques_activos = db.obtener_bloques()
        self._por_clave = {}
        for b in bloques_activos:
            for k in (str(b["codigo"]).strip().lower(),
                      nombre_opcion_bloque(b["codigo"]).lower()):
                self._por_clave.setdefault(k, b)

    def buscar(self, codigo):
        """(bloque | None, mensaje_error)."""
        c = str(codigo or "").strip()
        if not c:
            return None, "código de bloque vacío"
        b = self._por_clave.get(c.lower()) or self._por_clave.get(_slug(c))
        if b:
            return b, ""
        existente = db.obtener_bloque_por_codigo(c)
        if existente and int(existente.get("activo", 1) or 0) == 0:
            return None, f"el bloque '{c}' está retirado del catálogo"
        if existente:
            return existente, ""
        return None, f"el bloque '{c}' no existe en el catálogo (no se crea)"


def preparar_envio(envio, catalogo, origen="api", formulario_uid=""):
    """Normaliza un envio. No escribe nada.

    Devuelve dict con: estado ('listo'|'error'), clave, errores, alertas,
    verificacion, inspeccion, indicadores, adjuntos.
    """
    datos = _aplanar(envio)
    res = {"estado": "listo", "clave": _clave_envio(datos), "errores": [],
           "alertas": [], "verificacion": None, "inspeccion": None,
           "indicadores": None, "adjuntos": envio.get("_attachments") or [],
           "codigo_bloque": ""}

    codigo = str(datos.get("codigo_bloque") or "").strip()
    if codigo == "otro":
        codigo = str(datos.get("codigo_bloque_otro") or "").strip()
    res["codigo_bloque"] = codigo
    bloque, err = catalogo.buscar(codigo)
    if err:
        res["errores"].append(err)

    fecha = str(datos.get("fecha_visita") or datos.get("fecha_hoy") or "").strip()[:10]
    inspector = str(datos.get("inspector") or "").strip()
    if not fecha:
        res["errores"].append("falta la fecha de visita")
    if not inspector:
        res["errores"].append("falta el nombre del verificador")

    # ── Coordenadas: el GPS del dispositivo manda ──
    lat, lon, alt, prec = _parsear_geopoint(datos)
    e_man = _num(datos.get("utm_este_ref")) or _num(datos.get("utm_este"))
    n_man = _num(datos.get("utm_norte_ref")) or _num(datos.get("utm_norte"))
    if e_man == 0:
        e_man = None
    if n_man == 0:
        n_man = None
    este = norte = None
    fuente = ""
    if lat is not None and lon is not None:
        este, norte = latlon_a_utm17s(lat, lon)
        fuente = "gps"
    elif e_man is not None and n_man is not None:
        este, norte, fuente = e_man, n_man, "manual"
        res["alertas"].append("sin GPS del dispositivo: se usó la UTM digitada")
    else:
        res["alertas"].append("sin coordenadas")

    valida = este is not None and utm_en_rango(este, norte)
    if este is not None and not valida:
        res["alertas"].append(
            f"UTM fuera de rango ({este:,.0f} E / {norte:,.0f} N)")
    if prec is not None and prec > PRECISION_ALERTA_M:
        res["alertas"].append(f"precisión GPS baja ({prec:.0f} m)")
    if fuente == "gps" and e_man is not None and n_man is not None:
        dif = math.hypot(este - e_man, norte - n_man)
        if dif > DIF_ALERTA_GPS_MANUAL_M:
            res["alertas"].append(
                f"GPS del celular y UTM digitada difieren {dif:,.0f} m")

    dist_c = None
    if bloque and este is not None:
        be, bn = _num(bloque.get("utm_este")), _num(bloque.get("utm_norte"))
        if be and bn:
            dist_c = math.hypot(este - be, norte - bn)
            if dist_c > DIST_ALERTA_CENTROIDE_M:
                res["alertas"].append(
                    f"punto a {dist_c:,.0f} m del centroide del bloque")

    # ── Alertas de criterios del proceso de seleccion ──
    n_predios = _num(datos.get("n_predios"))
    if n_predios is not None and n_predios > MAX_PREDIOS_BLOQUE:
        res["alertas"].append(
            f"{int(n_predios)} predios (> {MAX_PREDIOS_BLOQUE}, Paso 5)")
    if datos.get("agricultura_activa") == "si":
        res["alertas"].append("agricultura activa en el bloque (criterio de exclusión)")
    if datos.get("accesibilidad") == "no_accesible":
        res["alertas"].append("bloque no accesible")
    if datos.get("aceptacion_titular") == "no_acepta":
        res["alertas"].append("el titular no acepta la intervención")
    if datos.get("dictamen") == "no_apto":
        res["alertas"].append("dictamen: no apto")

    lbl = {campo: _etiqueta(lista, datos.get(campo))
           for campo, lista in CAMPO_LISTA.items()}
    # Clima del formulario anterior (valores libres ya traducidos)
    if not lbl["condiciones_climaticas"]:
        lbl["condiciones_climaticas"] = str(datos.get("condiciones_climaticas") or "")

    fotos = []
    for campo in CAMPOS_FOTO + ("foto_1", "foto_2", "foto_3"):
        v = str(datos.get(campo) or "").strip()
        if v:
            fotos.append({"campo": campo, "archivo": os.path.basename(v)})

    obs_partes = [str(datos.get(k) or "").strip() for k in
                  ("observaciones", "motivo_no_apto", "accesibilidad_obs")]
    observaciones = " | ".join(p for p in obs_partes if p)

    codigo_ver = str(datos.get("codigo_verificacion") or "").strip()
    if not codigo_ver and fecha:
        codigo_ver = f"VER-{fecha.replace('-', '')}-{res['clave'][-8:]}"

    datos_crudos = {k: v for k, v in envio.items() if k != "_attachments"}

    res["verificacion"] = {
        "clave_envio": res["clave"],
        "kobo_id": int(_num(datos.get("_id")) or 0) or None,
        "formulario_uid": formulario_uid or "",
        "formulario_version": str(datos.get("__version__") or datos.get("_version_") or ""),
        "origen": origen,
        "bloque_id": bloque["id"] if bloque else None,
        "codigo_bloque": bloque["codigo"] if bloque else codigo,
        "distrito": (bloque.get("distrito") if bloque else "") or str(datos.get("distrito") or ""),
        "fecha_visita": fecha, "inspector": inspector,
        "gps_lat": lat, "gps_lon": lon, "gps_altitud": alt, "gps_precision": prec,
        "utm_este": este, "utm_norte": norte, "utm_zona": "17S",
        "utm_fuente": fuente, "utm_este_manual": e_man, "utm_norte_manual": n_man,
        "coord_valida": 1 if valida else 0,
        "distancia_centroide_m": dist_c,
        "accesibilidad": lbl["accesibilidad"], "tipo_acceso": lbl["tipo_acceso"],
        "tiempo_acceso_min": int(_num(datos.get("tiempo_acceso_min"))) if _num(datos.get("tiempo_acceso_min")) is not None else None,
        "acceso_estacional": lbl["acceso_estacional"],
        "riesgo_social": lbl["riesgo_social"],
        "riesgo_social_factores": lbl["riesgo_social_factores"],
        "riesgo_social_detalle": str(datos.get("riesgo_social_detalle") or ""),
        "uso_suelo_predominante": lbl["uso_suelo_predominante"],
        "uso_suelo_otros": lbl["uso_suelo_otros"],
        "agricultura_activa": lbl["agricultura_activa"],
        "cobertura_vegetal_pct": _num(datos.get("cobertura_vegetal_pct")),
        "evidencia_mm": lbl["evidencia_mm"], "tipos_mm": lbl["tipos_mm"],
        "magnitud_mm": lbl["magnitud_mm"], "actividad_mm": lbl["actividad_mm"],
        "tenencia": lbl["tenencia"],
        "n_predios": int(n_predios) if n_predios is not None else None,
        "titular": str(datos.get("titular") or ""),
        "aceptacion_titular": lbl["aceptacion_titular"],
        "condiciones_aceptacion": str(datos.get("condiciones_aceptacion") or ""),
        "acta_firmada": lbl["acta_firmada"],
        "dictamen": lbl["dictamen"],
        "motivo_no_apto": str(datos.get("motivo_no_apto") or ""),
        "condiciones_climaticas": lbl["condiciones_climaticas"],
        "observaciones": str(datos.get("observaciones") or ""),
        "fotos": json.dumps(fotos, ensure_ascii=False),
        "alertas": "; ".join(res["alertas"]),
        "datos_json": json.dumps(datos_crudos, ensure_ascii=False, default=str),
        "fecha_envio": str(datos.get("_submission_time") or ""),
    }

    # Inspeccion ligada (para que el envio aparezca en las vistas actuales)
    resumen = [f"[ODK Paso 6] Dictamen: {lbl['dictamen'] or 's/d'}"]
    for titulo, campo in (("Accesibilidad", "accesibilidad"),
                          ("Riesgo social", "riesgo_social"),
                          ("Uso del suelo", "uso_suelo_predominante"),
                          ("Mov. en masa", "evidencia_mm"),
                          ("Aceptación", "aceptacion_titular")):
        if lbl[campo]:
            resumen.append(f"{titulo}: {lbl[campo]}")
    if not lbl["dictamen"]:
        resumen = ["[ODK]"]
    if observaciones:
        resumen.append(f"Obs.: {observaciones}")
    res["inspeccion"] = {
        "fecha_visita": fecha, "inspector": inspector,
        "condiciones_climaticas": lbl["condiciones_climaticas"],
        "avance_fisico": _num(datos.get("avance_fisico")) or 0,
        "observaciones": " | ".join(resumen),
        "desviaciones": "; ".join(filter(None, [str(datos.get("desviaciones") or ""),
                                                "; ".join(res["alertas"])])),
        "registro_fotografico": "; ".join(f["archivo"] for f in fotos),
        "codigo_verificacion": codigo_ver,
        "microcuenca": (bloque.get("microcuenca") if bloque else "") or "",
    }

    # Indicadores del formulario anterior (si el envio los trae)
    ind = {
        "cobertura_vegetal_planificada": _num(datos.get("cobertura_vegetal_planificada")) or 0,
        "cobertura_vegetal_lograda": _num(datos.get("cobertura_vegetal_lograda")) or 0,
        "sobrevivencia_especies": _num(datos.get("sobrevivencia_especies")) or 0,
        "longitud_zanjas_ejecutada": _num(datos.get("longitud_zanjas")) or 0,
        "volumen_retencion_sedimentos": _num(datos.get("volumen_retencion")) or 0,
    }
    if any(ind.values()):
        res["indicadores"] = ind

    if res["errores"]:
        res["estado"] = "error"
    return res


def preparar_envios(envios, origen="api", formulario_uid="", catalogo=None,
                    claves_existentes=None):
    """Prepara una lista de envios y marca los ya importados. No escribe."""
    catalogo = catalogo or CatalogoBloques()
    if claves_existentes is None:
        claves_existentes = db.obtener_claves_verificacion_odk()
    vistos = set()
    salida = []
    for envio in envios:
        p = preparar_envio(envio, catalogo, origen, formulario_uid)
        if p["clave"] in claves_existentes:
            p["estado"] = "duplicado"
        elif p["clave"] in vistos:
            p["estado"] = "duplicado"
            p["alertas"].append("repetido dentro del mismo archivo")
        vistos.add(p["clave"])
        salida.append(p)
    return salida


def resumen_preparacion(preparados):
    """DataFrame-friendly: una fila por envio para la vista previa."""
    filas = []
    for i, p in enumerate(preparados, 1):
        v = p["verificacion"] or {}
        filas.append({
            "N°": i, "Estado": {"listo": "Nuevo", "duplicado": "Ya importado",
                                "error": "Rechazado"}[p["estado"]],
            "Bloque": p["codigo_bloque"], "Fecha": v.get("fecha_visita", ""),
            "Verificador": v.get("inspector", ""),
            "Dictamen": v.get("dictamen", ""),
            "UTM Este": round(v["utm_este"], 1) if v.get("utm_este") is not None else None,
            "UTM Norte": round(v["utm_norte"], 1) if v.get("utm_norte") is not None else None,
            "Fuente UTM": v.get("utm_fuente", ""),
            "Fotos": len(json.loads(v.get("fotos") or "[]")),
            "Errores": "; ".join(p["errores"]),
            "Alertas": "; ".join(p["alertas"]),
        })
    return filas


# ══════════════════════════════════════════════════════════════════════════
# Registro en la BD (solo altas nuevas)
# ══════════════════════════════════════════════════════════════════════════

def _resultado_vacio(total):
    return {
        "total_filas": total, "verificaciones_nuevas": 0,
        "duplicados_omitidos": 0, "rechazados": 0,
        "inspecciones_creadas": 0, "inspecciones_vinculadas": 0,
        "indicadores_creados": 0,
        "fotos_descargadas": 0, "fotos_fallidas": 0,
        "con_alertas": 0, "errores": [],
        # Se mantienen por compatibilidad: la importacion ya no toca bloques.
        "bloques_nuevos": 0, "bloques_actualizados": 0,
    }


def registrar_envios(preparados, cliente=None, descargar_fotos=True,
                     progreso=None):
    """Escribe los envios en estado 'listo'. Nunca modifica bloques ni borra.

    `cliente`: KoBoClient para descargar fotos (opcional).
    `progreso`: callback(i, total) opcional.
    """
    r = _resultado_vacio(len(preparados))
    for i, p in enumerate(preparados, 1):
        if progreso:
            progreso(i, len(preparados))
        if p["estado"] == "duplicado":
            r["duplicados_omitidos"] += 1
            continue
        if p["estado"] == "error":
            r["rechazados"] += 1
            r["errores"].append(f"Envío {i} ({p['codigo_bloque'] or 's/código'}): "
                                + "; ".join(p["errores"]))
            continue
        try:
            out = db.registrar_verificacion_odk(p["verificacion"], p["inspeccion"],
                                                p["indicadores"])
        except Exception as e:
            r["errores"].append(f"Envío {i}: {e}")
            continue
        if out["estado"] == "duplicado":
            r["duplicados_omitidos"] += 1
            continue
        r["verificaciones_nuevas"] += 1
        if out.get("inspeccion_existente"):
            r["inspecciones_vinculadas"] += 1
        elif out["inspeccion_id"]:
            r["inspecciones_creadas"] += 1
            if p["indicadores"]:
                r["indicadores_creados"] += 1
        if p["alertas"]:
            r["con_alertas"] += 1
        if descargar_fotos and cliente and p["adjuntos"]:
            ok, fallo = _guardar_fotos(cliente, p, out["verificacion_id"])
            r["fotos_descargadas"] += ok
            r["fotos_fallidas"] += fallo
            if fallo:
                r["errores"].append(f"Envío {i}: {fallo} foto(s) no se pudieron descargar")
    return r


def _guardar_fotos(cliente, preparado, verificacion_id):
    campos = {f["archivo"]: f["campo"]
              for f in json.loads(preparado["verificacion"]["fotos"] or "[]")}
    ok = fallo = 0
    for adj in preparado["adjuntos"]:
        nombre = os.path.basename(str(adj.get("filename") or adj.get("media_file_basename") or ""))
        if not nombre:
            continue
        url = (adj.get("download_medium_url") or adj.get("download_large_url")
               or adj.get("download_url"))
        try:
            contenido = cliente.descargar(url)
            db.guardar_adjunto_odk(
                verificacion_id, preparado["clave"],
                campos.get(nombre, str(adj.get("question_xpath") or "").split("/")[-1]),
                nombre, adj.get("mimetype") or "",
                adj.get("download_url") or url, contenido)
            ok += 1
        except Exception:
            fallo += 1
    return ok, fallo


# ══════════════════════════════════════════════════════════════════════════
# Lectura de archivos CSV / Excel exportados
# ══════════════════════════════════════════════════════════════════════════

def leer_archivo_odk(ruta):
    """Lee un CSV (',' o ';') o XLSX exportado de KoBo/ODK -> lista de dicts."""
    if ruta.lower().endswith((".xlsx", ".xlsm")):
        if not OPENPYXL_DISPONIBLE:
            raise ImportError("Se requiere openpyxl para leer Excel.")
        wb = load_workbook(ruta, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        filas = list(ws.iter_rows(values_only=True))
        wb.close()
        if not filas:
            return []
        enc = [str(h or "").strip() for h in filas[0]]
        return [{enc[j]: ("" if v is None else (v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else str(v)))
                 for j, v in enumerate(f) if j < len(enc) and enc[j]}
                for f in filas[1:] if any(v not in (None, "") for v in f)]
    with open(ruta, "r", encoding="utf-8-sig", newline="") as f:
        muestra = f.read(4096)
        f.seek(0)
        try:
            dialecto = csv.Sniffer().sniff(muestra, delimiters=",;\t")
        except csv.Error:
            dialecto = csv.excel
        return [dict(r) for r in csv.DictReader(f, dialect=dialecto)]


def importar_csv_odk(ruta_csv):
    """Importa un CSV/XLSX exportado (compatibilidad con la version anterior)."""
    envios = leer_archivo_odk(ruta_csv)
    return registrar_envios(preparar_envios(envios, origen="csv"))


# ══════════════════════════════════════════════════════════════════════════
# Cliente API KoBoToolbox v2
# ══════════════════════════════════════════════════════════════════════════

def _dominio_base(host):
    partes = (host or "").lower().split(".")
    return ".".join(partes[-2:])


class KoBoClient:
    """Cliente minimo de la API v2 de KoBoToolbox (verificacion SSL activa)."""

    def __init__(self, url_servidor, token_api, timeout=60):
        self.url_servidor = url_servidor.rstrip("/")
        self.token_api = token_api
        self.timeout = timeout
        self._host = urllib.parse.urlparse(self.url_servidor).hostname or ""
        if urllib.parse.urlparse(self.url_servidor).scheme != "https":
            raise ValueError("El servidor debe usar https://")
        self.ssl_context = ssl.create_default_context()

    def _url_permitida(self, url):
        """El token solo se envia al mismo dominio del servidor configurado."""
        u = urllib.parse.urlparse(url)
        return (u.scheme == "https" and
                _dominio_base(u.hostname) == _dominio_base(self._host))

    def _abrir(self, url, aceptar="application/json"):
        if not self._url_permitida(url):
            raise ConnectionError(f"URL fuera del servidor configurado: {url}")
        req = urllib.request.Request(url, headers={
            "Authorization": f"Token {self.token_api}", "Accept": aceptar})
        try:
            with urllib.request.urlopen(req, context=self.ssl_context,
                                        timeout=self.timeout) as resp:
                return resp.read()
        except urllib.error.HTTPError as e:
            cuerpo = e.read().decode("utf-8", errors="replace")[:300]
            if e.code in (401, 403):
                raise ConnectionError("Token inválido o sin permiso sobre el formulario")
            raise ConnectionError(f"Error HTTP {e.code}: {cuerpo}")
        except urllib.error.URLError as e:
            raise ConnectionError(f"Error de conexión: {e.reason}")

    def _get_json(self, endpoint, params=None):
        url = f"{self.url_servidor}/api/v2/{endpoint}"
        q = dict(params or {})
        q.setdefault("format", "json")
        url += ("&" if "?" in url else "?") + urllib.parse.urlencode(q)
        return json.loads(self._abrir(url).decode("utf-8"))

    def _paginar(self, endpoint, params=None, limite=PAGINA_API):
        """Recorre todas las paginas con limit/start (no se trunca en 100)."""
        resultados, inicio = [], 0
        while True:
            q = dict(params or {})
            q.update({"limit": limite, "start": inicio})
            pagina = self._get_json(endpoint, q)
            lote = pagina.get("results", [])
            resultados.extend(lote)
            inicio += len(lote)
            total = pagina.get("count")
            if not lote or not pagina.get("next") or (total is not None and inicio >= total):
                break
        return resultados

    def listar_formularios(self):
        formularios = []
        for a in self._paginar("assets/", {"asset_type": "survey"}, limite=300):
            if a.get("asset_type") != "survey":
                continue
            formularios.append({
                "uid": a["uid"], "nombre": a.get("name", "Sin nombre"),
                "fecha_modificacion": a.get("date_modified", ""),
                "envios": a.get("deployment__submission_count", 0),
                "desplegado": a.get("has_deployment", False),
            })
        return formularios

    def obtener_envios(self, formulario_uid):
        """Todos los envios del formulario (todas las paginas)."""
        uid = urllib.parse.quote(formulario_uid, safe="")
        return self._paginar(f"assets/{uid}/data/")

    def descargar(self, url):
        """Descarga un adjunto (foto) con autenticacion."""
        return self._abrir(url, aceptar="*/*")

    def test_conexion(self):
        try:
            self._get_json("assets/", {"limit": 1})
            return True, "Conexión exitosa"
        except ConnectionError as e:
            return False, str(e)
        except Exception as e:
            return False, f"Error inesperado: {e}"


def importar_desde_kobo(url_servidor, token_api, formulario_uid,
                        descargar_fotos=True):
    """Descarga todos los envios del formulario y registra los nuevos."""
    cliente = KoBoClient(url_servidor, token_api)
    envios = cliente.obtener_envios(formulario_uid)
    preparados = preparar_envios(envios, origen="api", formulario_uid=formulario_uid)
    return registrar_envios(preparados, cliente=cliente,
                            descargar_fotos=descargar_fotos)


def texto_resultado(r):
    """Resumen legible de una importacion."""
    t = (f"{r['total_filas']} envío(s) procesados | nuevos: {r['verificaciones_nuevas']} | "
         f"ya importados (omitidos): {r['duplicados_omitidos']} | "
         f"rechazados: {r['rechazados']} | fotos: {r['fotos_descargadas']}")
    if r.get("con_alertas"):
        t += f" | con alertas: {r['con_alertas']}"
    return t


# ══════════════════════════════════════════════════════════════════════════
# Pestana ODK / KoBoToolbox (version de escritorio Tkinter)
# ══════════════════════════════════════════════════════════════════════════

_FrameBase = ttk.Frame if _HAS_TK else object


class TabODKKobo(_FrameBase):
    """Pestana de integracion con ODK/KoBoToolbox para colecta en campo."""

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.kobo_client = None
        self._crear_widgets()

    def _crear_widgets(self):
        canvas = tk.Canvas(self, bg="#ECF0F1", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas)
        self.scroll_frame.bind("<Configure>",
                               lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        frame = self.scroll_frame
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="ODK / KoBoToolbox - Verificación de campo (Paso 6)",
                  style="Header.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(12, 4))
        ttk.Label(frame,
                  text="La importación NO modifica ni crea bloques y no duplica envíos "
                       "ya importados.", wraplength=700).grid(
            row=1, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 10))

        sec1 = ttk.LabelFrame(frame, text=" Generar Formulario XLSForm ", padding=12)
        sec1.grid(row=2, column=0, columnspan=3, sticky="ew", padx=12, pady=6)
        ttk.Button(sec1, text="Generar Formulario XLSForm",
                   command=self._generar_formulario,
                   style="Accent.TButton").grid(row=0, column=0, sticky="w", pady=4)
        self.label_estado_form = ttk.Label(sec1, text="", wraplength=600)
        self.label_estado_form.grid(row=1, column=0, columnspan=3, sticky="w", pady=4)

        sec2 = ttk.LabelFrame(frame, text=" Importar CSV / Excel exportado ", padding=12)
        sec2.grid(row=3, column=0, columnspan=3, sticky="ew", padx=12, pady=6)
        btns = ttk.Frame(sec2)
        btns.grid(row=0, column=0, sticky="w", pady=4)
        ttk.Button(btns, text="Seleccionar archivo...", command=self._importar_csv,
                   style="Accent.TButton").pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Descargar plantilla CSV",
                   command=self._descargar_plantilla_csv).pack(side="left", padx=4)
        self.label_estado_csv = ttk.Label(sec2, text="", wraplength=600)
        self.label_estado_csv.grid(row=1, column=0, sticky="w", pady=4)
        self.frame_resultados = ttk.Frame(sec2)
        self.frame_resultados.grid(row=2, column=0, sticky="ew", pady=4)

        sec3 = ttk.LabelFrame(frame, text=" Sincronización con KoBoToolbox (API) ", padding=12)
        sec3.grid(row=4, column=0, columnspan=3, sticky="ew", padx=12, pady=6)
        ttk.Label(sec3, text="Servidor:").grid(row=0, column=0, sticky="w", pady=3)
        self.combo_servidor = ttk.Combobox(sec3, values=[
            "https://kf.kobotoolbox.org", "https://eu.kobotoolbox.org",
            "https://kobo.humanitarianresponse.info"], width=40)
        self.combo_servidor.grid(row=0, column=1, sticky="w", padx=6, pady=3)
        self.combo_servidor.set(os.environ.get("KOBO_SERVER", "https://kf.kobotoolbox.org"))
        ttk.Label(sec3, text="Token API:").grid(row=1, column=0, sticky="w", pady=3)
        self.entry_token = ttk.Entry(sec3, width=42, show="*")
        self.entry_token.grid(row=1, column=1, sticky="w", padx=6, pady=3)
        if os.environ.get("KOBO_TOKEN"):
            self.entry_token.insert(0, os.environ["KOBO_TOKEN"])
        bf = ttk.Frame(sec3)
        bf.grid(row=2, column=0, columnspan=3, sticky="w", pady=6)
        ttk.Button(bf, text="Probar Conexión", command=self._probar_conexion).pack(side="left", padx=(0, 8))
        ttk.Button(bf, text="Listar Formularios", command=self._listar_formularios).pack(side="left", padx=4)
        ttk.Button(bf, text="Importar Envíos", command=self._importar_envios_kobo,
                   style="Accent.TButton").pack(side="left", padx=4)
        self.label_estado_api = ttk.Label(sec3, text="", wraplength=600)
        self.label_estado_api.grid(row=3, column=0, columnspan=3, sticky="w", pady=4)
        cols = ("uid", "nombre", "envios", "estado")
        self.tree_formularios = ttk.Treeview(sec3, columns=cols, show="headings", height=5)
        for c, t, w in zip(cols, ("UID", "Nombre", "Envíos", "Estado"), (120, 280, 70, 100)):
            self.tree_formularios.heading(c, text=t)
            self.tree_formularios.column(c, width=w)
        self.tree_formularios.grid(row=4, column=0, columnspan=3, sticky="ew", pady=4)

    def _generar_formulario(self):
        try:
            try:
                bloques = db.obtener_bloques()
            except Exception:
                bloques = []
            ruta = generar_xlsform(bloques=bloques)
            self.label_estado_form.config(text=f"Formulario generado:\n{ruta}",
                                          foreground="#1B4D2E")
            messagebox.showinfo("XLSForm Generado", f"Formulario listo:\n\n{ruta}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar el formulario:\n{e}")

    def _importar_csv(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo exportado de KoBo/ODK",
            filetypes=[("CSV / Excel", "*.csv *.xlsx"), ("Todos", "*.*")])
        if not ruta:
            return
        try:
            self._mostrar_resultados_importacion(importar_csv_odk(ruta))
            self.app.refrescar_todo()
        except Exception as e:
            messagebox.showerror("Error de importación", f"No se pudo importar:\n{e}")

    def _descargar_plantilla_csv(self):
        ruta = filedialog.asksaveasfilename(
            title="Guardar plantilla CSV", defaultextension=".csv",
            filetypes=[("CSV", "*.csv")], initialfile="plantilla_in_piura_paso6.csv")
        if not ruta:
            return
        with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerow(encabezados_plantilla_csv())
        self.label_estado_csv.config(text=f"Plantilla guardada en: {ruta}",
                                     foreground="#1B4D2E")

    def _mostrar_resultados_importacion(self, r):
        for w in self.frame_resultados.winfo_children():
            w.destroy()
        color = "#1B4D2E" if not r["errores"] else "#F39C12"
        self.label_estado_csv.config(text=texto_resultado(r), foreground=color)
        for err in r["errores"][:10]:
            ttk.Label(self.frame_resultados, text=f"  - {err}",
                      foreground="#E74C3C", wraplength=580).pack(anchor="w")

    def _probar_conexion(self):
        url, token = self.combo_servidor.get().strip(), self.entry_token.get().strip()
        if not url or not token:
            messagebox.showwarning("Validación", "Ingrese la URL del servidor y el token API.")
            return
        try:
            cliente = KoBoClient(url, token)
        except ValueError as e:
            self.label_estado_api.config(text=str(e), foreground="#E74C3C")
            return
        ok, msg = cliente.test_conexion()
        self.kobo_client = cliente if ok else None
        self.label_estado_api.config(text=msg, foreground="#1B4D2E" if ok else "#E74C3C")

    def _listar_formularios(self):
        if not self.kobo_client:
            self._probar_conexion()
            if not self.kobo_client:
                return
        try:
            for item in self.tree_formularios.get_children():
                self.tree_formularios.delete(item)
            for f in self.kobo_client.listar_formularios():
                self.tree_formularios.insert("", "end", iid=f["uid"], values=(
                    f["uid"], f["nombre"], f["envios"],
                    "Desplegado" if f["desplegado"] else "Borrador"))
        except Exception as e:
            self.label_estado_api.config(text=f"Error: {e}", foreground="#E74C3C")

    def _importar_envios_kobo(self):
        if not self.kobo_client:
            self._probar_conexion()
            if not self.kobo_client:
                return
        sel = self.tree_formularios.selection()
        if not sel:
            messagebox.showwarning("Selección", "Seleccione un formulario de la lista.")
            return
        if not messagebox.askyesno("Confirmar Importación",
                                   f"¿Importar envíos del formulario {sel[0]}?\n"
                                   "Solo se agregan envíos nuevos; no se modifican bloques."):
            return
        try:
            envios = self.kobo_client.obtener_envios(sel[0])
            r = registrar_envios(preparar_envios(envios, "api", sel[0]),
                                 cliente=self.kobo_client)
            self._mostrar_resultados_importacion(r)
            self.app.refrescar_todo()
            self.label_estado_api.config(text=texto_resultado(r), foreground="#1B4D2E")
        except Exception as e:
            self.label_estado_api.config(text=f"Error: {e}", foreground="#E74C3C")
