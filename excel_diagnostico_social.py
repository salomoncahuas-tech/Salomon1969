"""
Modulo de importacion/exportacion Excel para Diagnostico Social V3.
Proyecto IN Piura CUI 2669244 | ANIN - DIME - SESDI

Reemplaza los 5 formatos anteriores (F-DS-01..05 generados por codigo) por la
plantilla oficial validada V3 con 7 fichas (F-DS-01..07), hojas auxiliares
(_Listas, _Codigos, _Datos) y celdas de validacion (desplegables) nativas.

Estrategia (igual que Diagnostico Territorial V5): se sirve el archivo .xlsx
oficial que vive en el repo y solo se inyecta la lista de bloques actual en la
hoja oculta `_Datos`, conservando intactas todas las listas de validacion.

API publica (estable):
    generar_plantilla_ds(fichas=None, bloques_data=None) -> bytes
    parsear_excel_ds(file_bytes, ficha=None) -> list[dict]
    mapear_a_session_state(resultado, bloques_map, fecha_min=None, fecha_max=None) -> dict
"""

import io
import os
import re
import unicodedata
import zipfile
import warnings
from datetime import date, datetime, time as dtime
from xml.sax.saxutils import escape

from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple

from bloque_lookup import buscar_label_bloque, normalizar_codigo_bloque
import fds_actores as FA
import fds_listas as FL


# ─── Plantilla oficial (vive junto al codigo en el repo) ───────────────────
# Version actual: V4. Se mantiene el alias PLANTILLA_V3_PATH por compatibilidad.
PLANTILLA_DS_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "Plantilla_Diagnostico_Social_IN_Piura_V4.xlsx",
)
PLANTILLA_V3_PATH = PLANTILLA_DS_PATH

FICHAS_HOJAS = ["F-DS-01", "F-DS-02", "F-DS-03", "F-DS-04",
                "F-DS-05", "F-DS-06", "F-DS-07"]


# ─── Generacion de la plantilla descargable ────────────────────────────────

_DATOS_HEADER = ["codigo", "microcuenca", "provincia", "distrito",
                 "zona", "este_utm", "norte_utm", "area_ha"]


def _col_letter(idx0):
    """Letra de columna 0-based (0->A)."""
    s = ""
    n = idx0
    while True:
        s = chr(ord("A") + n % 26) + s
        n = n // 26 - 1
        if n < 0:
            break
    return s


def _build_datos_sheetdata(bloques_data):
    """Construye el bloque <sheetData> de la hoja _Datos (cabecera + filas),
    usando cadenas en linea para no depender de sharedStrings."""
    def cell(col0, row, val):
        ref = f"{_col_letter(col0)}{row}"
        txt = escape("" if val is None else str(val))
        return (f'<c r="{ref}" t="inlineStr"><is>'
                f'<t xml:space="preserve">{txt}</t></is></c>')
    rows_xml = []
    # Fila 1: cabecera
    cells = "".join(cell(c, 1, h) for c, h in enumerate(_DATOS_HEADER))
    rows_xml.append(f'<row r="1">{cells}</row>')
    # Filas de datos
    for i, b in enumerate(bloques_data, start=2):
        vals = list(b) + [""] * (len(_DATOS_HEADER) - len(b))
        cells = "".join(cell(c, i, vals[c]) for c in range(len(_DATOS_HEADER)))
        rows_xml.append(f'<row r="{i}">{cells}</row>')
    return "<sheetData>" + "".join(rows_xml) + "</sheetData>"


def _resolver_hoja_datos(zf):
    """Devuelve la ruta del XML de la hoja `_Datos` resolviendo los rels."""
    wb_xml = zf.read("xl/workbook.xml").decode("utf-8", "ignore")
    m = re.search(r'<sheet[^>]*name="_Datos"[^>]*r:id="(rId\d+)"', wb_xml)
    if not m:
        return None
    rid = m.group(1)
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
    m2 = re.search(rf'<Relationship[^>]*Id="{rid}"[^>]*Target="([^"]+)"', rels)
    if not m2:
        return None
    target = m2.group(1).lstrip("/")
    if not target.startswith("xl/"):
        target = "xl/" + target
    return target


def generar_plantilla_ds(fichas=None, bloques_data=None):
    """Devuelve los bytes del .xlsx oficial V3.

    Para conservar intactas TODAS las celdas de validacion (desplegables x14
    que referencian la hoja `_Listas`), NO se reabre el libro con openpyxl
    (openpyxl elimina esas validaciones al guardar). En su lugar se reescribe
    a nivel ZIP unicamente el XML de la hoja oculta `_Datos` con la lista de
    bloques actual; el resto del archivo se conserva byte a byte.

    Args:
        fichas: ignorado (la plantilla trae siempre las 7 fichas).
        bloques_data: lista de tuplas
            (codigo, microcuenca, provincia, distrito[, zona, este, norte, area]).

    Returns:
        bytes con el contenido del .xlsx.
    """
    if not os.path.exists(PLANTILLA_V3_PATH):
        raise FileNotFoundError(
            f"No se encontro la plantilla V3 oficial: {PLANTILLA_V3_PATH}")

    with open(PLANTILLA_V3_PATH, "rb") as fh:
        raw = fh.read()

    if not bloques_data:
        return raw

    zin = zipfile.ZipFile(io.BytesIO(raw))
    datos_path = _resolver_hoja_datos(zin)
    if not datos_path or datos_path not in zin.namelist():
        return raw  # sin _Datos no hay nada que inyectar; servir tal cual

    sheet_xml = zin.read(datos_path).decode("utf-8", "ignore")
    nuevo_sheetdata = _build_datos_sheetdata(bloques_data)
    if "<sheetData" in sheet_xml:
        sheet_xml = re.sub(r"<sheetData[^>]*>.*?</sheetData>|<sheetData[^>]*/>",
                           nuevo_sheetdata, sheet_xml, count=1, flags=re.S)
    # Recalcular dimension para abarcar las filas escritas
    last_row = len(bloques_data) + 1
    sheet_xml = re.sub(r'<dimension ref="[^"]*"/>',
                       f'<dimension ref="A1:H{last_row}"/>', sheet_xml, count=1)

    buf_out = io.BytesIO()
    with zipfile.ZipFile(buf_out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == datos_path:
                data = sheet_xml.encode("utf-8")
            zout.writestr(item, data)
    buf_out.seek(0)
    return buf_out.getvalue()


# ─── Lectura de la plantilla V4 llenada ────────────────────────────────────
#
# La plantilla V4 es un formulario de posiciones FIJAS: cada dato tiene su
# celda. Por eso se lee por coordenada exacta (nunca "el primer texto a la
# derecha del rotulo", que devolvia el rotulo vecino cuando el campo estaba
# vacio). Tipos de dato que maneja el tecnico en campo:
#   * Texto / numero en la celda de valor (esquina superior izquierda de la
#     celda combinada).
#   * Casillas "[ ]": la casilla es la celda inmediatamente a la IZQUIERDA del
#     rotulo de la opcion y se marca con «X» (tambien x, ✓, ✔, 1, Si).
#   * Desplegables (validaciones) y codigos de la hoja «_Codigos» (GL, A, F,
#     SH, LT, 2...), que se traducen al valor completo del aplicativo.
# Las claves de los formularios son las mismas keys de los widgets de
# streamlit_app (convencion clave del dict == key del widget).

FECHA_MIN_DEFECTO = date(2024, 1, 1)

_MARCAS = {"x", "✓", "✔", "√", "●", "•", "1", "si", "sí", "true", "verdadero"}
_RE_MARCA = re.compile(r"(?:^|[\s\[\(«])[xX✓✔√](?:$|[\s\]\)»:.,;])")


def _norm(v):
    """Texto comparable: sin acentos, minusculas, espacios simples."""
    s = unicodedata.normalize("NFKD", str(v or "")).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s).strip().lower()


def _clean(v):
    """Valor de celda como texto limpio (formatea fechas, horas y enteros)."""
    if v is None:
        return ""
    if isinstance(v, datetime):
        if (v.hour, v.minute, v.second) == (0, 0, 0):
            return v.strftime("%d/%m/%Y")
        if v.date() == date(1899, 12, 30) or v.year < 1901:
            return v.strftime("%H:%M")
        return v.strftime("%d/%m/%Y %H:%M")
    if isinstance(v, date):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, dtime):
        return v.strftime("%H:%M")
    if isinstance(v, bool):
        return "Sí" if v else "No"
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    if s in ("▼", "-", "—"):
        return ""
    return s


def _rc(coord):
    return coordinate_to_tuple(coord)


def _val(ws, coord):
    r, c = _rc(coord)
    return ws.cell(row=r, column=c).value


def _txt(ws, coords, sep=" "):
    """Texto de una celda o de varias (lineas de un bloque de respuesta)."""
    if isinstance(coords, str):
        coords = [coords]
    partes = [_clean(_val(ws, c)) for c in coords]
    return sep.join(p for p in partes if p)


def _es_marca(v):
    s = _clean(v)
    if not s:
        return False
    return _norm(s) in _MARCAS or bool(_RE_MARCA.search(s))


def _marca(ws, coord):
    return _es_marca(_val(ws, coord))


def _rotulo_marcado(ws, coord):
    """Opciones sin casilla propia (p. ej. operador celular, genero F-DS-07):
    el tecnico escribe una «X» junto al texto del rotulo."""
    s = _clean(_val(ws, coord))
    return bool(s and _RE_MARCA.search(s))


def _a_opcion(valor, opciones, codigos=None):
    """Traduce un valor capturado (texto, desplegable o codigo) a la opcion
    exacta del aplicativo. Si no hay equivalencia se devuelve tal cual."""
    v = _clean(valor)
    if not v:
        return ""
    if v in opciones:
        return v
    n = _norm(v)
    for k, o in (codigos or {}).items():
        if _norm(k) == n:
            return o
    for o in opciones:
        if _norm(o) == n:
            return o
    for o in opciones:                       # "MF" -> "MF (Muy frecuente)"
        if "(" in o and _norm(o.split("(")[0]) == n:
            return o
    for o in opciones:                       # "Muy frecuente" -> "MF (...)"
        m = re.search(r"\(([^)]*)\)", o)
        if m and _norm(m.group(1)) == n:
            return o
    if "(" in v:                             # "Otro (especificar...)" -> "Otro"
        base = _norm(v.split("(")[0])
        for o in opciones:
            if _norm(o) == base:
                return o
    cands = [o for o in opciones if _norm(o).startswith(n)]
    if len(cands) == 1:                      # "Alta" -> "Alta — viable ..."
        return cands[0]
    return v


# Codigos de la hoja «_Codigos» (llenado manual en campo)
COD_TIPO_ACTOR = dict(zip(
    ["GN", "GR", "GL", "CC", "RC", "JU", "JS", "CG", "ONG", "AC", "EP", "AP", "IR", "LI", "MC"],
    FL.L_TIPO_ACTOR))
COD_NIV_TERR = dict(zip(["N", "R", "P", "D", "C", "F"], FL.L_NIV_TERR))
COD_ABC = dict(zip(["A", "M", "B"], FL.L_ABC))
COD_POSICION = dict(zip(["F", "N", "R", "C", "ND"], FL.L_POSICION))
COD_TIPO_CONFL = dict(zip(
    ["SM", "SH", "SF", "TL", "UR", "PE", "OP", "CU", "LP", "IF", "II", "OT"], FL.L_TIPO_CONFL))
COD_ESTADO_CONFL = dict(zip(["LT", "ES", "AC", "NG", "RS"], FL.L_NIVEL_CONFL))
COD_ANTIG_CONFL = dict(zip(["1", "2", "3", "4", "5", "6"], FL.L_ANTIG_CONFL))
COD_ACTIV = {f"A{i:02d}": a for i, a in enumerate(FL.L_ACTIV, start=1)}
COD_SEXO = {"H": "M", "HOMBRE": "M", "MASCULINO": "M", "V": "M",
            "MUJER": "F", "FEMENINO": "F"}
COD_SINO = {"S": "Sí", "SI": "Sí", "N": "No", "NA": "No aplica", "N/A": "No aplica"}

_SI, _NO, _NA = "Sí", "No", "No aplica"


# ─── Especificacion celda a celda de la plantilla V4 ───────────────────────

# Cabecera: fecha, responsable y codigo de bloque de cada ficha.
CABECERA = {
    "F-DS-01": {"fecha": "D7", "evaluador": "K7", "codigo_bloque": "D8"},
    "F-DS-02": {"fecha": "C7", "evaluador": "E7", "codigo_bloque": "J7"},
    "F-DS-03": {"fecha": "D7", "evaluador": "J7", "codigo_bloque": "D8"},
    "F-DS-04": {"fecha": "C7", "evaluador": "H7", "codigo_bloque": "C8"},
    "F-DS-05": {"fecha": "D7", "evaluador": "J7", "codigo_bloque": "D8"},
    "F-DS-06": {"fecha": "D7", "evaluador": "J7", "codigo_bloque": "D8"},
    "F-DS-07": {"fecha": "D7", "evaluador": "J7", "codigo_bloque": "D8"},
}

# Datos generales comunes (widgets ds_* del aplicativo).
GENERALES = {
    "F-DS-01": {"ds_fnum": "G7", "ds_cpob": "I9", "ds_ubigeo": "M9", "ds_ccam": "D10",
                "ds_este": "E11", "ds_norte": "I11", "ds_alt": "M11",
                "ds_entrev_nombre": "D12", "ds_entrev_dni": "G12",
                "ds_entrev_oficio": "K12"},
    "F-DS-02": {"ds_cpob": "L8"},
}

# Observaciones generales (ds_obs): lineas del bloque «OBSERVACIONES».
OBSERVACIONES = {
    "F-DS-01": [f"A{r}" for r in range(133, 138)],
    "F-DS-02": [f"A{r}" for r in range(51, 57)],
    "F-DS-05": [f"A{r}" for r in range(61, 65)],
    "F-DS-06": [f"A{r}" for r in range(61, 66)],
}


def _lineas(ini, fin):
    return [f"A{r}" for r in range(ini, fin + 1)]


# Campos de texto: clave -> celda (o lista de celdas que se concatenan).
TEXTOS = {
    "F-DS-01": {
        "f1_nombre_oficial": "F19", "f1_anio_fund": "N19",
        "f1_nfam": "E21", "f1_pob_t": "L21", "f1_pob_h": "E22", "f1_pob_m": "L22",
        "f1_pob_men18": "E23", "f1_pob_may65": "L23", "f1_pob_orig": "E24",
        "f1_mano_obra": "L24",
        "f1_junta_fin": "L40", "f1_pres_junta": "D41", "f1_pres_dni": "J41",
        "f1_pres_tel": "M41", "f1_ronda_tel": "L42", "f1_pres_ronda": "D43",
        "f1_aut_adic": "E44", "f1_nombre_comite": "D50",
        "f1_n_predios": "E58", "f1_pct_tituladas": "L58",
        "f1_agua_cob": "D71", "f1_agua_pago": "I71", "f1_agua_acarreo": "M71",
        "f1_energia_cob": "D80", "f1_energia_pago": "I80",
        "f1_ie_nombre": "D89", "f1_ie_dist": "M89",
        "f1_eess_nombre": "E93", "f1_eess_dist": "N93",
        "f1_juntos": "H118", "f1_pension65": "H119", "f1_qaliwarma": "H120",
        "f1_beca18": "H121", "f1_otros_prog": "H122", "f1_nombre_ongs": "H128",
    },
    "F-DS-02": {
        "f2_favor": "F44", "f2_contra": "F45", "f2_decision": "F46",
        "f2_ronda": "F47", "f2_plataforma": "F48",
    },
    "F-DS-03": {
        "f3_nombre": "D10", "f3_dni": "D11", "f3_edad": "I11", "f3_cargo": "D12",
        "f3_inst": "D13", "f3_anios": "D14", "f3_tel": ["G14", "H14"],
        "f3_correo": ["K14", "L14"], "f3_lugar": "D15", "f3_dur": "M15",
    },
    "F-DS-04": {
        "f4_lugar": "D11", "f4_conv": "D12", "f4_fecha": "D13", "f4_hi": "I13",
        "f4_hf": "M13", "f4_conv_n": "E15", "f4_h": "L15", "f4_m": "E16",
        "f4_jov": "L16", "f4_am": "E17", "f4_tot": "L17", "f4_obj": "D18",
    },
    "F-DS-05": {"f5_otros": "I29", "f5_rol_rondas": "E30", "f5_estrategia": "F54"},
    "F-DS-06": {
        "f6_fuente": "D9", "f6_desc": "H49", "f6_p1": "F52", "f6_p2": "F53",
        "f6_p3": "F54", "f6_genero_desc": ["F56", "A57", "A58"],
    },
    "F-DS-07": {
        "f7_nombre": "D14", "f7_dni": "J14", "f7_edad": "M14", "f7_residencia": "K15",
        "f7_contacto": "D16", "f7_superficie": "D17", "f7_cond": "F44", "f7_plazo": "H46",
    },
}

# Respuestas largas (text_area): clave -> lineas del bloque de respuesta.
TEXTOS_LARGOS = {
    "F-DS-03": {
        "f3_r1": _lineas(20, 23), "f3_r2": _lineas(26, 29), "f3_r3": _lineas(31, 34),
        "f3_r4": _lineas(36, 39), "f3_r5": _lineas(42, 45), "f3_r6": _lineas(47, 50),
        "f3_r7": _lineas(52, 55), "f3_r8": _lineas(57, 59), "f3_r9": _lineas(62, 65),
        "f3_r10": _lineas(67, 70), "f3_r11": _lineas(72, 75),
        "f3_r_acuerdo": _lineas(78, 81), "f3_r_horarios": _lineas(83, 86),
        "f3_r_ant": _lineas(88, 91), "f3_cierre": _lineas(93, 98),
    },
    "F-DS-07": {"f7_preg": _lineas(53, 55), "f7_comp": _lineas(57, 59)},
}


def _sino(si, no, na=None):
    out = [(si, _SI), (no, _NO)]
    if na:
        out.append((na, _NA))
    return out


def _grid(celdas, opciones):
    """Empareja casillas con opciones de una lista oficial (mismo orden)."""
    return list(zip(celdas, opciones))


# Casillas «marcar uno»: clave -> [(celda_casilla, valor)].
UNICA = {
    "F-DS-01": {
        "f1_org_terr": _grid(["A15", "H15", "A16", "H16", "A17", "H17", "A18", "H18"],
                             FL.L_ORG_COMUNAL),
        "f1_idioma": _grid(["A26", "F26", "K26", "A27", "F27"], FL.L_IDIOMA),
        "f1_nivel_edu": _grid(["A29", "F29", "K29", "A30", "F30", "K30", "A31"], FL.L_NIV_EDU),
        "f1_migracion": _grid(["A33", "F33", "K33"], FL.L_ABC),
        "f1_destino_mig": [("A35", FL.L_MIG_DEST[0]), ("F35", FL.L_MIG_DEST[1]),
                           ("K35", FL.L_MIG_DEST[2]), ("A36", FL.L_MIG_DEST[3]),
                           ("F36", FL.L_MIG_DEST[4]), ("K36", FL.L_MIG_DEST[5]),
                           ("A37", FL.L_MIG_DEST[6]), ("F37", FL.L_MIG_DEST[7]),
                           ("K37", FL.L_MIG_DEST[9])],
        "f1_junta_vig": _sino("D40", "F40"),
        "f1_ronda": _sino("D42", "F42"),
        "f1_periodicidad": _grid(["A46", "F46", "K46", "A47", "F47", "K47", "A48"],
                                 FL.L_PERIODIC),
        "f1_reglamento": _sino("D49", "F49", "H49"),
        "f1_comite_rrnn": _sino("K49", "M49"),
        "f1_tenencia": _grid(["A54", "H54", "A55", "H55", "A56", "H56", "A57"], FL.L_TENENCIA),
        "f1_conf_linderos": _sino("E59", "G59"),
        "f1_superpone": _sino("K59", "M59"),
        "f1_reg_titulacion": [("A62", "COFOPRI"), ("F62", "Dirección Regional de Agricultura"),
                              ("K62", "SUNARP"), ("A63", "Comunidad (autónomo)"),
                              ("F63", "Otro"), ("K63", "Sin registro")],
        "f1_telecom": _grid(["A82", "H82", "A83", "H83"], FL.L_TELECOM),
        "f1_ie_niveles": [("A86", "Solo inicial"), ("F86", "Inicial + Primaria"),
                          ("K86", "Primaria"), ("A87", "Primaria + Secundaria"),
                          ("F87", "Secundaria"), ("K87", "Superior"), ("A88", "No hay IE"),
                          ("F88", "Inicial + Primaria + Secundaria")],
        "f1_eess": _grid(["A91", "D91", "H91", "K91", "A92", "D92", "H92", "K92"],
                         FL.L_CAT_EESS),
        "f1_local_comunal": _sino("E94", "G94"),
        "f1_local_estado": _grid(["E95", "G95", "I95"], FL.L_BRM),
        "f1_agrorural": _sino("H123", "J123"),
        "f1_prodern": _sino("H124", "J124"),
        "f1_otros_proy": _sino("H125", "J125"),
        "f1_pdc": _sino("H126", "J126"),
        "f1_ongs": _sino("H127", "J127"),
        "f1_presencia_estatal": _grid(["A130", "F130", "K130"], FL.L_ABC),
    },
    "F-DS-03": {
        "f3_c_nom": _sino("H16", "J16"),
        "f3_c_foto": _sino("H17", "J17"),
    },
    "F-DS-04": {
        "f4_idioma": _grid(["A31", "D31", "H31", "K31"], FL.FDS04_IDIOMA),
    },
    "F-DS-05": {
        "f5_rb1": _sino("I24", "K24", "M24"), "f5_rb2": _sino("I25", "K25", "M25"),
        "f5_rb3": _sino("I26", "K26", "M26"), "f5_rb4": _sino("I27", "K27", "M27"),
        "f5_rb5": _sino("I28", "K28", "M28"),
        "f5_polar": _grid(["A32", "F32", "K32", "A33", "F33"], FL.FDS05_POLARIZACION),
        "f5_confglob": _grid(["A49", "F49", "K49", "A50", "F50"], FL.FDS05_CONFLICTIVIDAD),
        "f5_viab": _grid(["A52", "H52", "A53", "H53"], FL.FDS05_VIABILIDAD),
        "f5_mesa": _sino("I55", "K55"),
        "f5_plazo": _grid(["A57", "F57", "K57", "A58", "F58"], FL.FDS05_PLAZO),
    },
    "F-DS-06": {
        "f6_medidas": _sino("H45", "J45"), "f6_alerta": _sino("H46", "J46"),
        "f6_saberes": _sino("H47", "J47"), "f6_apoyo": _sino("H48", "J48"),
        "f6_genero": _sino("H55", "J55"),
    },
    "F-DS-07": {
        "f7_tipo_prop": _grid(["A11", "H11", "A12", "H12", "A13", "H13"],
                              FL.L_TIPO_PROPIETARIO),
        "f7_linderos": _sino("H22", "J22"),
        "f7_residente": _sino("H23", "J23"),
        "f7_disp": _grid(["A39", "A40", "A41", "A42", "A43"], FL.L_DISPOSICION),
        "f7_consultar": _sino("H45", "J45"),
        "f7_aut_ing": _sino("H47", "J47"),
        "f7_aut_foto": _sino("H48", "J48"),
        "f7_aut_nom": _sino("H49", "J49"),
    },
}

# Desplegable alternativo a las casillas (columna O de F-DS-06).
UNICA_DESPLEGABLE = {
    "F-DS-06": {"f6_medidas": "O45", "f6_alerta": "O46", "f6_saberes": "O47",
                "f6_apoyo": "O48", "f6_genero": "O55"},
}

# Casillas «marcar las que apliquen»: clave -> [(celda_casilla, valor)].
MULTIPLE = {
    "F-DS-01": {
        "f1_agua": _grid(["A67", "H67", "A68", "H68", "A69", "H69", "A70", "H70"], FL.L_AGUA),
        "f1_sanea": _grid(["A73", "H73", "A74", "H74", "A75", "H75"], FL.L_SANEA),
        "f1_energia": _grid(["A77", "H77", "A78", "H78", "A79", "H79"], FL.L_ENERG),
    },
    "F-DS-04": {
        "f4_metod": _grid(["A20", "H20", "A21", "H21", "A22", "H22", "A23", "H23"],
                          FL.FDS04_METODOS),
        "f4_mater": _grid(["A25", "H25", "A26", "H26", "A27", "H27", "A28", "H28", "A29", "H29"],
                          FL.FDS04_MATERIALES),
    },
    "F-DS-07": {
        "f7_docs": _grid(["A19", "H19", "A20", "H20", "A21", "H21"], FL.FDS07_DOCS),
    },
}


# ─── Lectores ──────────────────────────────────────────────────────────────

def _leer_unica(ws, opciones, alterna=None, codigos=None):
    marcadas = [v for celda, v in opciones if _marca(ws, celda)]
    if marcadas:
        return marcadas[0]
    if alterna:
        valores = [v for _, v in opciones]
        return _a_opcion(_val(ws, alterna), valores, codigos or COD_SINO)
    return ""


def _leer_multiple(ws, opciones):
    return [v for celda, v in opciones if _marca(ws, celda)]


def _read_table(ws, filas, colmap):
    """colmap: {columna_app: celda_col (letra) | callable(ws, fila)}.
    Devuelve solo las filas con algun valor."""
    out = []
    for r in filas:
        rec = {}
        for name, col in colmap.items():
            rec[name] = col(ws, r) if callable(col) else _clean(_val(ws, f"{col}{r}"))
        if any(v for v in rec.values()):
            out.append(rec)
    return out


def _col_opcion(col, opciones, codigos=None, alterna=None):
    def leer(ws, r):
        v = _val(ws, f"{col}{r}")
        if not _clean(v) and alterna:
            v = _val(ws, f"{alterna}{r}")
        return _a_opcion(v, opciones, codigos)
    return leer


def _col_marca_o_valor(col, opciones, si_marca, otras=(), codigos=None):
    """Columna con desplegable que tambien admite «X»; `otras` son columnas
    vecinas de casilla (p. ej. «No», «Rara», «Baja»)."""
    def leer(ws, r):
        v = _val(ws, f"{col}{r}")
        if _es_marca(v) and _norm(_clean(v)) not in ("si", "1"):
            return si_marca
        txt = _a_opcion(v, opciones, codigos)
        if txt:
            return txt
        for c, valor in otras:
            if _marca(ws, f"{c}{r}"):
                return valor
        return ""
    return leer


def _tabla_actividades(ws):
    def actividad(ws, r):
        v = _clean(_val(ws, f"B{r}")) or _clean(_val(ws, f"C{r}"))
        return _a_opcion(v, FL.L_ACTIV, COD_ACTIV)

    def destino(ws, r):
        for c, d in zip("IJKLM", FL.L_DESTINO):
            if _marca(ws, f"{c}{r}"):
                return d
        return ""

    return _read_table(ws, range(100, 108), {
        "Actividad / Rubro": actividad, "N fam.": "E",
        "Productos principales": "F", "Destino": destino, "Ingreso (S/./mes)": "N",
    })


def _tabla_actores(ws):
    filas = _read_table(ws, range(16, 36), {
        "Nombre del actor": "B",
        "Tipo": _col_opcion("C", FL.L_TIPO_ACTOR, COD_TIPO_ACTOR),
        "Rol / Funcion frente al proyecto": "D",
        "Influencia": _col_opcion("E", FL.L_ABC, COD_ABC),
        "Interes": _col_opcion("F", FL.L_ABC, COD_ABC),
        "Posicion": _col_opcion("G", FL.L_POSICION, COD_POSICION),
        "Nivel territorial": _col_opcion("H", FL.L_NIV_TERR, COD_NIV_TERR),
        "Telefono": "I", "Correo / Contacto": "J", "Observaciones / Historial": "K",
    })
    return FA.migrar_filas(filas)


def _tabla_participantes(ws):
    return _read_table(ws, range(35, 55), {
        "Nombres y Apellidos": "B", "DNI": "D", "Institucion / Comunidad": "E",
        "Cargo / Rol": "G", "Telefono": "I",
        "Sexo": _col_opcion("K", FL.FDS03_GENERO, COD_SEXO), "Edad": "L",
    })


def _tabla_agenda(ws):
    return _read_table(ws, range(58, 63), {
        "Hora": "B", "Agenda": "C", "Responsable": "H", "Resultado / Aporte": "K",
    })


def _tabla_acuerdos(ws):
    return _read_table(ws, range(66, 71), {
        "Acuerdo / Compromiso": "B", "Responsable": "G", "Plazo": "J",
        "Medio de verificacion": "L",
    })


def _tabla_conflictos(ws):
    # Estado / Antig.: en la plantilla original el desplegable quedo en D/E
    # (dentro de la celda combinada C:E); se lee F/G y, si estan vacias, D/E.
    return _read_table(ws, range(14, 22), {
        "Tipo": _col_opcion("B", FL.L_TIPO_CONFL, COD_TIPO_CONFL),
        "Actores involucrados": "C",
        "Estado": _col_opcion("F", FL.L_NIVEL_CONFL, COD_ESTADO_CONFL, alterna="D"),
        "Antiguedad": _col_opcion("G", FL.L_ANTIG_CONFL, COD_ANTIG_CONFL, alterna="E"),
        "Descripcion / Causa raiz": "H",
        "Impacto potencial en el proyecto": "K",
    })


def _tabla_oportunidades(ws):
    return _read_table(ws, range(36, 46), {
        "Oportunidad identificada": "B", "Actores relacionados": "E",
        "Tipo (alianza / plataforma / proy.)": "H",
        "Potencial": _col_opcion("J", FL.L_ABC, COD_ABC), "Como aprovecharla": "K",
    })


def _tabla_con_nombres(ws, filas, col_nombre, nombres_app, colmap):
    """Tablas con filas pre-rotuladas (peligros, cambios climaticos): se
    conservan las filas predefinidas del aplicativo y la fila «Otro» solo si
    se lleno."""
    out = []
    for r in filas:
        nombre = _a_opcion(_val(ws, f"B{r}"), nombres_app)
        rec = {col_nombre: nombre}
        for name, col in colmap.items():
            rec[name] = col(ws, r) if callable(col) else _clean(_val(ws, f"{col}{r}"))
        lleno = any(v for k, v in rec.items() if k != col_nombre)
        if lleno or (nombre and _norm(nombre) != "otro"):
            out.append(rec)
    return out


def _tabla_peligros(ws):
    F = FL.FDS06_FRECUENCIA
    M = FL.FDS06_MAGNITUD
    T = FL.FDS06_TENDENCIA
    return _tabla_con_nombres(ws, range(15, 27), "Peligro observado", FL.L_PELIGRO_OBS, {
        "¿Ocurre?": _col_marca_o_valor("C", FL.L_SINO, _SI, [("D", _NO)], COD_SINO),
        "Frecuencia": _col_marca_o_valor("E", F, "F (Frecuente)", [("F", "R (Rara)")]),
        "Magnitud": _col_marca_o_valor("G", M, "A (Alta)", [("H", "B (Baja)")]),
        "Tendencia": _col_marca_o_valor("I", T, "Sube (Aumentando)",
                                        [("J", "Baja (Disminuyendo)")]),
        "Ultimo evento (año)": "K",
        "Principales daños observados": "L",
    })


def _tabla_cambios(ws):
    I = FL.FDS06_INTENSIDAD
    return _tabla_con_nombres(ws, range(35, 43), "Cambio observado", FL.L_CAMBIO_CLIMA, {
        "¿Se percibe?": _col_marca_o_valor("E", FL.L_SINO, _SI, [("F", _NO)], COD_SINO),
        "Intensidad": _col_marca_o_valor("G", I, "Alta", [("H", "Media"), ("I", "Baja")],
                                         {"A": "Alta", "M": "Media", "B": "Baja"}),
        "Año aprox. de inicio": "J",
        "Impacto en la comunidad / territorio": "L",
    })


def _puntos_fds07(ws):
    claves = ["f7_info_anin", "f7_info_objetivo", "f7_info_no_minero",
              "f7_info_medidas", "f7_info_temporalidad", "f7_info_voluntaria",
              "f7_info_actualizada", "f7_info_confidencialidad",
              "f7_info_preguntas", "f7_info_material"]
    leer = _col_marca_o_valor("M", FL.L_SINO, _SI, [("N", _NO)], COD_SINO)
    return {k: leer(ws, 27 + i) for i, k in enumerate(claves)}


def _operador_celular(ws):
    ops = [n for c, n in (("D84", "Movistar"), ("F84", "Claro"), ("H84", "Entel"),
                          ("J84", "Bitel")) if _rotulo_marcado(ws, c)]
    otros = _clean(_val(ws, "L84"))
    m = re.match(r"(?i)\s*otros\s*:?\s*(.*)$", otros)
    resto = (m.group(1) if m else otros).strip()
    resto = _RE_MARCA.sub(" ", resto).strip(" :")
    if resto:
        ops.append(resto)
    return ", ".join(ops)


def _genero_fds07(ws):
    for c, v in zip(("D15", "F15", "H15"), FL.FDS07_GENERO):
        if _rotulo_marcado(ws, c):
            return v
    return ""


_TABLAS = {
    "F-DS-01": {"f1_activ": _tabla_actividades},
    "F-DS-02": {"f2_actores": _tabla_actores},
    "F-DS-04": {"f4_part": _tabla_participantes, "f4_agenda": _tabla_agenda,
                "f4_acuerdos": _tabla_acuerdos},
    "F-DS-05": {"f5_conflictos": _tabla_conflictos,
                "f5_oportunidades": _tabla_oportunidades},
    "F-DS-06": {"f6_peligros": _tabla_peligros, "f6_cambios": _tabla_cambios},
}

# Columnas que en tablas pre-rotuladas no cuentan como dato capturado.
_COLS_ROTULO = {"Peligro observado", "Cambio observado"}


def _parse_form(ws, hoja):
    form = {}
    for k, celdas in TEXTOS.get(hoja, {}).items():
        form[k] = _txt(ws, celdas)
    for k, celdas in TEXTOS_LARGOS.get(hoja, {}).items():
        form[k] = _txt(ws, celdas, sep="\n")
    alternas = UNICA_DESPLEGABLE.get(hoja, {})
    for k, ops in UNICA.get(hoja, {}).items():
        form[k] = _leer_unica(ws, ops, alternas.get(k))
    for k, ops in MULTIPLE.get(hoja, {}).items():
        form[k] = _leer_multiple(ws, ops)
    for k, fn in _TABLAS.get(hoja, {}).items():
        form[k] = fn(ws)
    if hoja == "F-DS-01":
        form["f1_telecom_op"] = _operador_celular(ws)
    elif hoja == "F-DS-03":
        form["f3_genero"] = _a_opcion(_val(ws, "M11"), FL.FDS03_GENERO, COD_SEXO)
    elif hoja == "F-DS-07":
        form["f7_genero"] = _genero_fds07(ws)
        form.update(_puntos_fds07(ws))
    return form


def _fecha_iso(valor):
    """Fecha de la cabecera como 'AAAA-MM-DD' (o el texto tal cual)."""
    if isinstance(valor, datetime):
        return valor.date().isoformat()
    if isinstance(valor, date):
        return valor.isoformat()
    s = _clean(valor).split(" ")[0]
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return s


def _tiene_contenido(v):
    if isinstance(v, list):
        for x in v:
            if isinstance(x, dict):
                if any(val for k, val in x.items() if k not in _COLS_ROTULO):
                    return True
            elif x:
                return True
        return False
    return bool(v)


def _leer_catalogo_datos(wb):
    """{codigo_normalizado: (microcuenca, provincia, distrito)} de `_Datos`."""
    if "_Datos" not in wb.sheetnames:
        return {}
    cat = {}
    for fila in wb["_Datos"].iter_rows(min_row=2, max_col=4, values_only=True):
        if not fila or fila[0] in (None, ""):
            continue
        cod = normalizar_codigo_bloque(fila[0])
        cat[cod] = tuple(_clean(x) for x in (list(fila[1:4]) + [None] * 3)[:3])
    return cat


def parsear_excel_ds(file_bytes, ficha=None):
    """Lee la plantilla V4 llenada. Devuelve lista de
    {"ficha", "datos": {fecha, evaluador, codigo_bloque, distrito,
    microcuenca, generales, form}, "avisos": [...]}.
    Solo incluye las fichas con algun dato capturado. Nunca lanza por una
    hoja individual (el error se informa en "avisos")."""
    from excel_diagnostico_territorial import _SheetGrid
    if hasattr(file_bytes, "getvalue"):
        raw = file_bytes.getvalue()
    elif hasattr(file_bytes, "read"):
        raw = file_bytes.read()
    else:
        raw = file_bytes
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        # read_only + _SheetGrid: una sola pasada acotada por hoja, para
        # que un rango usado inflado por Excel no agote CPU/memoria.
        wb = load_workbook(io.BytesIO(raw), data_only=True, read_only=True)

    try:
        catalogo = _leer_catalogo_datos(wb)
        objetivo = [ficha] if ficha else FICHAS_HOJAS
        resultados = []
        for hoja in objetivo:
            if hoja not in wb.sheetnames or hoja not in CABECERA:
                continue
            ws = _SheetGrid(wb[hoja])
            try:
                cab = CABECERA[hoja]
                codigo = _txt(ws, cab["codigo_bloque"])
                datos = {
                    "fecha": _fecha_iso(_val(ws, cab["fecha"])),
                    "evaluador": _txt(ws, cab["evaluador"]),
                    "codigo_bloque": codigo,
                }
                mc, _prov, dist = catalogo.get(normalizar_codigo_bloque(codigo), ("", "", ""))
                datos["microcuenca"], datos["distrito"] = mc, dist
                datos["generales"] = {k: _txt(ws, c)
                                      for k, c in GENERALES.get(hoja, {}).items()}
                if hoja in OBSERVACIONES:
                    datos["generales"]["ds_obs"] = _txt(ws, OBSERVACIONES[hoja], sep="\n")
                form = _parse_form(ws, hoja)
                datos["form"] = form
                n_campos = sum(1 for v in form.values() if _tiene_contenido(v))
                datos["n_campos"] = n_campos
                if n_campos or any(datos.get(k) for k in ("fecha", "evaluador", "codigo_bloque")) \
                        or any(datos["generales"].values()):
                    resultados.append({"ficha": hoja, "datos": datos, "avisos": []})
            except Exception as e:  # noqa: BLE001 - una hoja no debe tumbar el resto
                resultados.append({"ficha": hoja, "datos": {"form": {}, "n_campos": 0},
                                   "avisos": [f"No se pudo leer la hoja {hoja}: {e}"]})
    finally:
        wb.close()
    return resultados


# Slots de tablas (deben coincidir con streamlit_app._DS_TABLE_SLOTS)
_TABLE_SLOTS = {
    "f1_activ", "f2_actores", "f4_part", "f4_agenda", "f4_acuerdos",
    "f5_conflictos", "f5_oportunidades", "f6_peligros", "f6_cambios",
}

# Rango valido de coordenadas UTM WGS84 17S del ambito (Cuenca Alta Piura).
_ESTE_RANGO = (450_000, 750_000)
_NORTE_RANGO = (9_300_000, 9_600_000)


def _numero(txt):
    try:
        return float(str(txt).replace(",", ".").replace(" ", ""))
    except (TypeError, ValueError):
        return None


def mapear_a_session_state(resultado, bloques_map, fecha_min=None, fecha_max=None):
    """Convierte un resultado de parseo en el dict de precarga {key: valor}
    consumido por streamlit_app._ds_apply_pending(). Los avisos (bloque no
    encontrado, fecha fuera de rango, coordenadas invalidas) se agregan a
    resultado["avisos"]."""
    ficha = resultado.get("ficha", "")
    datos = resultado.get("datos", {}) or {}
    form = datos.get("form", {}) or {}
    avisos = resultado.setdefault("avisos", [])
    fecha_min = fecha_min or FECHA_MIN_DEFECTO
    fecha_max = fecha_max or date.today()

    pend = {"ds_ficha_sel": ficha}
    if datos.get("evaluador"):
        pend["ds_eval"] = datos["evaluador"]

    # Fecha (el widget solo admite fechas dentro del rango del proyecto)
    fecha_txt = str(datos.get("fecha", "") or "")
    if fecha_txt:
        try:
            f = datetime.strptime(fecha_txt, "%Y-%m-%d").date()
            if fecha_min <= f <= fecha_max:
                pend["ds_fecha"] = f
            else:
                avisos.append(f"{ficha}: la fecha {f:%d/%m/%Y} esta fuera del rango "
                              f"permitido ({fecha_min:%d/%m/%Y} - {fecha_max:%d/%m/%Y}); "
                              "corrijala en el formulario.")
        except ValueError:
            avisos.append(f"{ficha}: no se reconoce la fecha «{fecha_txt}» "
                          "(use DD-MM-AAAA).")

    # Bloque (igualdad exacta del codigo; nunca se adivina)
    codigo = str(datos.get("codigo_bloque", "") or "").strip()
    if codigo:
        label_bl = buscar_label_bloque(codigo, bloques_map) if bloques_map else None
        if label_bl:
            pend["ds_bl"] = label_bl
        else:
            avisos.append(f"{ficha}: el bloque «{codigo}» no existe en el aplicativo; "
                          "seleccionelo manualmente antes de guardar.")
    else:
        avisos.append(f"{ficha}: la plantilla no indica el codigo de bloque; "
                      "seleccionelo manualmente antes de guardar.")

    # Datos generales (centro poblado, coordenadas, entrevistado, observaciones)
    for k, v in (datos.get("generales") or {}).items():
        if not v:
            continue
        if k in ("ds_este", "ds_norte", "ds_alt"):
            num = _numero(v)
            if num is None:
                avisos.append(f"{ficha}: valor no numerico en {k.replace('ds_', '')}: «{v}».")
                continue
            rango = {"ds_este": _ESTE_RANGO, "ds_norte": _NORTE_RANGO}.get(k)
            if rango and not (rango[0] <= num <= rango[1]):
                avisos.append(f"{ficha}: la coordenada {k.replace('ds_', '').upper()} "
                              f"{num:,.0f} no corresponde a UTM 17S del ambito "
                              f"({rango[0]:,}–{rango[1]:,}); se mantiene el centroide del bloque.")
                continue
            pend[k] = float(num)
        else:
            pend[k] = v

    # Volcar el formulario completo (las claves vacias limpian valores previos)
    for k, v in form.items():
        if k in _TABLE_SLOTS:
            pend[f"_dsinit_{k}"] = v if isinstance(v, list) else []
        else:
            pend[k] = v
    return pend
