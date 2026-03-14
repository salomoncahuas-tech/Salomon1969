"""
Modulo de importacion/exportacion Excel para Diagnostico Territorial.
Proyecto IN Piura CUI 2669244 | ANIN - DIME - SESDI

Genera plantillas Excel estandarizadas para las fichas F-DT-01 a F-DT-06
y parsea archivos Excel llenados por tecnicos de campo para autocompletar
los formularios del aplicativo.
"""

import io
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


# ── Estilos (mismos que diagnostico social) ──────────────────────────────
HEADER_FILL = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
SUBHEADER_FILL = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
SUBHEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
LABEL_FILL = PatternFill(start_color="D5DBDB", end_color="D5DBDB", fill_type="solid")
LABEL_FONT = Font(name="Calibri", size=10, bold=True)
VALUE_FILL = PatternFill(start_color="FDEBD0", end_color="FDEBD0", fill_type="solid")
VALUE_FONT = Font(name="Calibri", size=10)
SECTION_FILL = PatternFill(start_color="1ABC9C", end_color="1ABC9C", fill_type="solid")
SECTION_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
INSTRUCCION_FONT = Font(name="Calibri", size=9, italic=True, color="7F8C8D")

# ── Opciones de listas (deben coincidir con streamlit_app.py) ──────────
FDT01_FORMA_TERRENO = [
    "Plano (0-2%)", "Ondulado (2-8%)", "Colinado (8-25%)",
    "Montanoso (25-50%)", "Escarpado (>50%)",
]
FDT01_PENDIENTE = [
    "0-5% (Plano a ligeramente inclinado)",
    "5-15% (Ligeramente inclinado a moderadamente empinado)",
    "15-25% (Moderadamente empinado)",
    "25-50% (Empinado)",
    ">50% (Muy empinado a extremadamente empinado)",
]
FDT01_POSICION_FISIOGRAFICA = [
    "Cima / Cresta", "Ladera superior (tercio superior)",
    "Ladera media (tercio medio)", "Ladera inferior (tercio inferior)",
    "Pie de ladera / Base", "Fondo de valle", "Terraza fluvial", "Llanura",
]
FDT01_EXPOSICION = [
    "Norte (N)", "Sur (S)", "Este (E)", "Oeste (O)",
    "Noreste (NE)", "Noroeste (NO)", "Sureste (SE)", "Suroeste (SO)",
    "Variable / Sin predominancia",
]
FDT01_PAISAJE = [
    "Montana", "Colina", "Piedemonte", "Planicie / Llanura",
    "Valle interandino", "Terraza aluvial",
]
FDT01_RANGO_ALTITUDINAL = [
    "<500 msnm", "500-1000 msnm", "1000-1500 msnm",
    "1500-2000 msnm", "2000-2500 msnm", "2500-3000 msnm",
    "3000-3500 msnm", "3500-4000 msnm", ">4000 msnm",
]

FDT02_PRECIPITACION = [
    "<250 mm/anio (Muy seco / Arido)",
    "250-500 mm/anio (Seco / Semiarido)",
    "500-1000 mm/anio (Sub-humedo)",
    "1000-2000 mm/anio (Humedo / Lluvioso)",
    ">2000 mm/anio (Muy humedo / Pluvial)",
]
FDT02_TEMPERATURA = [
    "<5 C (Muy frio / Gelido)", "5-12 C (Frio)", "12-18 C (Templado)",
    "18-24 C (Calido / Semicalido)", ">24 C (Muy calido / Tropical)",
]
FDT02_HUMEDAD = [
    "Muy baja (<30%)", "Baja (30-50%)", "Media (50-70%)",
    "Alta (70-85%)", "Muy alta (>85%)",
]
FDT02_ZONA_VIDA = [
    "Desierto superarido", "Desierto arido", "Matorral desertico",
    "Monte espinoso", "Bosque seco", "Bosque humedo premontano",
    "Bosque humedo montano bajo", "Bosque humedo montano",
    "Bosque muy humedo premontano", "Bosque muy humedo montano",
    "Paramo / Jalca", "Puna",
]
FDT02_HELADAS = [
    "Frecuente (>30 dias/anio)", "Ocasional (10-30 dias/anio)",
    "Rara (<10 dias/anio)", "Ausente",
]
FDT02_VIENTOS = [
    "Calmo (< 2 m/s)", "Suave (2-4 m/s)", "Moderado (4-8 m/s)",
    "Fuerte (8-14 m/s)", "Muy fuerte (>14 m/s)",
]

FDT03_TEXTURA = [
    "Arenoso", "Franco arenoso", "Franco", "Franco limoso",
    "Franco arcilloso", "Franco arcillo arenoso", "Arcilloso", "Limoso",
]
FDT03_COLOR = [
    "Negro / Muy oscuro", "Pardo oscuro", "Pardo / Marron",
    "Pardo claro / Amarillento", "Rojizo / Rojo amarillento",
    "Gris / Gris claro",
]
FDT03_PROFUNDIDAD = [
    "Muy superficial (<25 cm)", "Superficial (25-50 cm)",
    "Moderadamente profundo (50-100 cm)", "Profundo (>100 cm)",
]
FDT03_PEDREGOSIDAD = [
    "Sin piedras (0%)", "Pocas piedras (0-15%)",
    "Frecuentes (15-35%)", "Abundantes (35-60%)",
    "Muy pedregoso (>60%)",
]
FDT03_DRENAJE = [
    "Excesivo (suelo muy arenoso, seca rapido)",
    "Bueno (suelo drena adecuadamente)",
    "Moderado (drena con cierta lentitud)",
    "Imperfecto (retiene humedad excesiva)",
    "Pobre / Muy pobre (encharcamiento frecuente)",
]
FDT03_EROSION = [
    "Sin erosion aparente",
    "Erosion laminar leve",
    "Erosion laminar moderada a severa",
    "Erosion en surcos",
    "Erosion en carcavas",
    "Erosion mixta (laminar + surcos/carcavas)",
    "Movimientos en masa (deslizamientos)",
]
FDT03_MATERIA_ORGANICA = [
    "Muy baja (suelo claro, sin restos organicos)",
    "Baja (pocos restos, suelo claro)",
    "Media (presencia moderada de restos organicos)",
    "Alta (suelo oscuro, abundantes restos organicos)",
]

FDT04_TIPO_COBERTURA = [
    "Bosque denso (natural)", "Bosque ralo / Abierto",
    "Matorral / Arbustal", "Pastizal / Herbazal / Pajonal",
    "Cultivo agricola", "Plantacion forestal",
    "Suelo desnudo / Eriazo", "Vegetacion riberena",
    "Area urbana / Infraestructura",
]
FDT04_DENSIDAD = [
    "Muy rala (<10%)", "Rala (10-25%)", "Abierta (25-50%)",
    "Semicerrada (50-75%)", "Cerrada (>75%)",
]
FDT04_ESTADO_CONSERVACION = [
    "Bueno (sin intervencion significativa)",
    "Regular (intervencion parcial)",
    "Degradado (intervencion severa, con regeneracion)",
    "Muy degradado (sin regeneracion natural evidente)",
]
FDT04_USO_SUELO = [
    "Forestal / Proteccion", "Agricola (secano)",
    "Agricola (bajo riego)", "Pecuario / Pastoreo",
    "Agrosilvopastoril", "Minero",
    "Sin uso / En abandono", "Conservacion / Area protegida",
]
FDT04_CONFLICTO_USO = [
    "Sin conflicto (uso adecuado a capacidad)",
    "Sobreuso leve", "Sobreuso moderado", "Sobreuso severo",
    "Subuso (capacidad no aprovechada)",
]

FDT05_FUENTE_AGUA = [
    "Rio permanente", "Quebrada / Riachuelo", "Manantial / Puquio",
    "Laguna / Reservorio", "Canal de riego",
    "Agua subterranea (pozo)", "Ninguna visible en el area",
]
FDT05_REGIMEN = [
    "Permanente (flujo todo el anio)",
    "Estacional (flujo en temporada de lluvias)",
    "Temporal / Efimero (solo con eventos de lluvia)",
    "Sin escurrimiento superficial",
]
FDT05_CALIDAD_AGUA = [
    "Buena (clara, sin olor, sin sedimentos)",
    "Regular (ligeramente turbia o con sedimentos)",
    "Mala (turbia, con olor, con color)",
    "Muy mala (contaminacion evidente)",
    "No evaluable (sin fuente de agua accesible)",
]
FDT05_DISTANCIA_AGUA = [
    "<100 m", "100-500 m", "500-1000 m", "1-5 km", ">5 km",
]
FDT05_USO_HIDRICO = [
    "Consumo humano", "Riego agricola", "Pecuario / Abrevadero",
    "Piscicola", "Uso multiple", "Sin uso actual",
]

FDT06_TENENCIA = [
    "Comunal (comunidad campesina)",
    "Privada individual (titulo)",
    "Privada individual (sin titulo / posesionario)",
    "Estatal / Fiscal",
    "Mixta (comunal + privada)",
    "Sin informacion",
]
FDT06_ORGANIZACION = [
    "Bien organizada (junta directiva activa, asambleas regulares)",
    "Moderadamente organizada (funcional pero irregular)",
    "Debilmente organizada (directiva nominal, poca participacion)",
    "Sin organizacion comunitaria identificada",
]
FDT06_ACTIVIDAD_ECONOMICA = [
    "Agricultura de subsistencia", "Agricultura comercial",
    "Ganaderia extensiva", "Ganaderia intensiva",
    "Actividad forestal / Extraccion", "Mineria artesanal",
    "Comercio / Servicios", "Mixta (agropecuaria)",
]
FDT06_ACCESIBILIDAD = [
    "Carretera asfaltada (acceso permanente)",
    "Carretera afirmada (acceso con restricciones en lluvia)",
    "Trocha carrozable (acceso limitado)",
    "Camino de herradura (solo a pie o acemila)",
    "Sin acceso vehicular (zona remota)",
]
FDT06_DISTANCIA_CENTRO = [
    "<1 km", "1-5 km", "5-10 km", "10-20 km", ">20 km",
]
FDT06_SERVICIOS = [
    "Agua potable", "Electricidad", "Telecomunicaciones / Internet",
    "Puesto de salud", "Escuela / IE", "Ninguno de los anteriores",
]


# ══════════════════════════════════════════════════════════════════════════
# UTILIDADES
# ══════════════════════════════════════════════════════════════════════════

def _cell(ws, row, col, value, font=None, fill=None, alignment=None, border=None):
    """Escribe un valor en una celda con formato opcional."""
    cell = ws.cell(row=row, column=col, value=value)
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border
    return cell


def _add_header(ws, ficha_titulo, max_col=4):
    """Agrega cabecera institucional comun a todas las fichas."""
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    _cell(ws, 1, 1, "PROYECTO IN PIURA - CUI 2669244",
          HEADER_FONT, HEADER_FILL, Alignment(horizontal="center", vertical="center"))
    for c in range(1, max_col + 1):
        ws.cell(row=1, column=c).fill = HEADER_FILL

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
    _cell(ws, 2, 1, ficha_titulo,
          SUBHEADER_FONT, SUBHEADER_FILL, Alignment(horizontal="center", vertical="center"))
    for c in range(1, max_col + 1):
        ws.cell(row=2, column=c).fill = SUBHEADER_FILL

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=max_col)
    _cell(ws, 3, 1, "ANIN - DIME - SESDI | Complete las celdas naranja",
          INSTRUCCION_FONT, None, Alignment(horizontal="center"))


def _add_datos_generales(ws, start_row=5):
    """Agrega la seccion de datos generales. Retorna la fila siguiente disponible."""
    r = start_row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    _cell(ws, r, 1, "DATOS GENERALES", SECTION_FONT, SECTION_FILL,
          Alignment(horizontal="center"))
    for c in range(1, 5):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    campos = [
        ("Fecha (AAAA-MM-DD)", ""),
        ("Evaluador / Especialista", ""),
        ("Codigo de Bloque", ""),
        ("Microcuenca", ""),
    ]
    for label, default in campos:
        _cell(ws, r, 1, label, LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        _cell(ws, r, 2, default, VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
        r += 1

    return r + 1  # fila vacia de separacion


def _add_validation(ws, col_letter, min_row, max_row, options):
    """Agrega validacion de lista desplegable a un rango de celdas."""
    formula = '"' + ",".join(options) + '"'
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    dv.error = "Seleccione un valor de la lista"
    dv.errorTitle = "Valor no valido"
    dv.prompt = "Seleccione de la lista"
    dv.promptTitle = "Opciones"
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}{min_row}:{col_letter}{max_row}")


def _add_parametro(ws, row, label, opciones=None):
    """Agrega una fila de parametro con label y celda de valor con validacion opcional."""
    _cell(ws, row, 1, label, LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    _cell(ws, row, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
    if opciones:
        _add_validation(ws, "B", row, row, opciones)
    return row + 1


# ══════════════════════════════════════════════════════════════════════════
# GENERACION DE PLANTILLAS
# ══════════════════════════════════════════════════════════════════════════

def _generar_dt01(wb):
    """Genera la hoja F-DT-01: Caracteristicas Fisiograficas."""
    ws = wb.create_sheet("F-DT-01")
    _add_header(ws, "F-DT-01: CARACTERISTICAS FISIOGRAFICAS")
    r = _add_datos_generales(ws)

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    _cell(ws, r, 1, "PARAMETROS DE EVALUACION", SECTION_FONT, SECTION_FILL,
          Alignment(horizontal="center"))
    for c in range(1, 5):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    r = _add_parametro(ws, r, "Forma del terreno", FDT01_FORMA_TERRENO)
    r = _add_parametro(ws, r, "Pendiente del terreno", FDT01_PENDIENTE)
    r = _add_parametro(ws, r, "Posicion fisiografica", FDT01_POSICION_FISIOGRAFICA)
    r = _add_parametro(ws, r, "Exposicion / Orientacion", FDT01_EXPOSICION)
    r = _add_parametro(ws, r, "Paisaje dominante", FDT01_PAISAJE)
    r = _add_parametro(ws, r, "Rango altitudinal", FDT01_RANGO_ALTITUDINAL)

    r += 1
    _cell(ws, r, 1, "Observaciones", LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=4)
    _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)

    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15


def _generar_dt02(wb):
    """Genera la hoja F-DT-02: Condiciones Climaticas."""
    ws = wb.create_sheet("F-DT-02")
    _add_header(ws, "F-DT-02: CONDICIONES CLIMATICAS")
    r = _add_datos_generales(ws)

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    _cell(ws, r, 1, "PARAMETROS DE EVALUACION", SECTION_FONT, SECTION_FILL,
          Alignment(horizontal="center"))
    for c in range(1, 5):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    r = _add_parametro(ws, r, "Precipitacion anual estimada", FDT02_PRECIPITACION)
    r = _add_parametro(ws, r, "Temperatura media anual", FDT02_TEMPERATURA)
    r = _add_parametro(ws, r, "Humedad relativa", FDT02_HUMEDAD)
    r = _add_parametro(ws, r, "Zona de vida (Holdridge)", FDT02_ZONA_VIDA)
    r = _add_parametro(ws, r, "Presencia de heladas", FDT02_HELADAS)
    r = _add_parametro(ws, r, "Regimen de vientos", FDT02_VIENTOS)

    r += 1
    _cell(ws, r, 1, "Observaciones", LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=4)
    _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)

    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15


def _generar_dt03(wb):
    """Genera la hoja F-DT-03: Caracteristicas del Suelo."""
    ws = wb.create_sheet("F-DT-03")
    _add_header(ws, "F-DT-03: CARACTERISTICAS DEL SUELO (Observacion de campo)")
    r = _add_datos_generales(ws)

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    _cell(ws, r, 1, "PARAMETROS DE EVALUACION", SECTION_FONT, SECTION_FILL,
          Alignment(horizontal="center"))
    for c in range(1, 5):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    r = _add_parametro(ws, r, "Textura al tacto", FDT03_TEXTURA)
    r = _add_parametro(ws, r, "Color predominante del suelo", FDT03_COLOR)
    r = _add_parametro(ws, r, "Profundidad efectiva", FDT03_PROFUNDIDAD)
    r = _add_parametro(ws, r, "Pedregosidad superficial", FDT03_PEDREGOSIDAD)
    r = _add_parametro(ws, r, "Drenaje", FDT03_DRENAJE)
    r = _add_parametro(ws, r, "Presencia de erosion", FDT03_EROSION)
    r = _add_parametro(ws, r, "Materia organica (estimacion visual)", FDT03_MATERIA_ORGANICA)

    r += 1
    _cell(ws, r, 1, "Observaciones", LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=4)
    _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)

    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15


def _generar_dt04(wb):
    """Genera la hoja F-DT-04: Cobertura Vegetal y Uso del Suelo."""
    ws = wb.create_sheet("F-DT-04")
    _add_header(ws, "F-DT-04: COBERTURA VEGETAL Y USO DEL SUELO")
    r = _add_datos_generales(ws)

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    _cell(ws, r, 1, "PARAMETROS DE EVALUACION", SECTION_FONT, SECTION_FILL,
          Alignment(horizontal="center"))
    for c in range(1, 5):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    r = _add_parametro(ws, r, "Tipo de cobertura vegetal", FDT04_TIPO_COBERTURA)
    r = _add_parametro(ws, r, "Densidad de cobertura", FDT04_DENSIDAD)
    r = _add_parametro(ws, r, "Estado de conservacion", FDT04_ESTADO_CONSERVACION)
    r = _add_parametro(ws, r, "Uso actual del suelo", FDT04_USO_SUELO)
    r = _add_parametro(ws, r, "Estado de Uso del Suelo", FDT04_CONFLICTO_USO)

    r += 1
    _cell(ws, r, 1, "Observaciones", LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=4)
    _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)

    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15


def _generar_dt05(wb):
    """Genera la hoja F-DT-05: Recursos Hidricos."""
    ws = wb.create_sheet("F-DT-05")
    _add_header(ws, "F-DT-05: RECURSOS HIDRICOS")
    r = _add_datos_generales(ws)

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    _cell(ws, r, 1, "PARAMETROS DE EVALUACION", SECTION_FONT, SECTION_FILL,
          Alignment(horizontal="center"))
    for c in range(1, 5):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    r = _add_parametro(ws, r, "Fuente de agua mas cercana", FDT05_FUENTE_AGUA)
    r = _add_parametro(ws, r, "Regimen hidrico", FDT05_REGIMEN)
    r = _add_parametro(ws, r, "Calidad aparente del agua", FDT05_CALIDAD_AGUA)
    r = _add_parametro(ws, r, "Distancia a fuente de agua", FDT05_DISTANCIA_AGUA)
    r = _add_parametro(ws, r, "Uso del recurso hidrico", FDT05_USO_HIDRICO)

    r += 1
    _cell(ws, r, 1, "Observaciones", LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=4)
    _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)

    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15


def _generar_dt06(wb):
    """Genera la hoja F-DT-06: Aspectos Socioeconomicos y Accesibilidad."""
    ws = wb.create_sheet("F-DT-06")
    _add_header(ws, "F-DT-06: ASPECTOS SOCIOECONOMICOS Y ACCESIBILIDAD")
    r = _add_datos_generales(ws)

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    _cell(ws, r, 1, "PARAMETROS DE EVALUACION", SECTION_FONT, SECTION_FILL,
          Alignment(horizontal="center"))
    for c in range(1, 5):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    r = _add_parametro(ws, r, "Tenencia de la tierra", FDT06_TENENCIA)
    r = _add_parametro(ws, r, "Organizacion comunal", FDT06_ORGANIZACION)
    r = _add_parametro(ws, r, "Actividad economica principal", FDT06_ACTIVIDAD_ECONOMICA)
    r = _add_parametro(ws, r, "Accesibilidad (via principal)", FDT06_ACCESIBILIDAD)
    r = _add_parametro(ws, r, "Distancia al centro poblado", FDT06_DISTANCIA_CENTRO)

    # Servicios basicos - campo especial con instruccion para multiseleccion
    _cell(ws, r, 1, "Servicios basicos disponibles", LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
    # Agregar comentario/instruccion en la fila siguiente
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    _cell(ws, r, 1,
          "  (Separe con coma los servicios: Agua potable, Electricidad, "
          "Telecomunicaciones / Internet, Puesto de salud, Escuela / IE, "
          "Ninguno de los anteriores)",
          INSTRUCCION_FONT)
    r += 1

    r += 1
    _cell(ws, r, 1, "Observaciones", LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=4)
    _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)

    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15


# ── Funcion publica para generar plantilla combinada ─────────────────────

def generar_plantilla_dt(fichas=None):
    """
    Genera un archivo Excel con plantillas para las fichas especificadas.

    Args:
        fichas: Lista de fichas a generar (ej: ["F-DT-01", "F-DT-03"]).
                Si es None, genera todas (F-DT-01 a F-DT-06).

    Returns:
        bytes: Contenido del archivo Excel listo para descarga.
    """
    if fichas is None:
        fichas = ["F-DT-01", "F-DT-02", "F-DT-03", "F-DT-04", "F-DT-05", "F-DT-06"]

    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    generadores = {
        "F-DT-01": _generar_dt01,
        "F-DT-02": _generar_dt02,
        "F-DT-03": _generar_dt03,
        "F-DT-04": _generar_dt04,
        "F-DT-05": _generar_dt05,
        "F-DT-06": _generar_dt06,
    }

    for ficha in fichas:
        gen = generadores.get(ficha)
        if gen:
            gen(wb)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════
# PARSEO DE EXCEL LLENADO
# ══════════════════════════════════════════════════════════════════════════

def _val(ws, row, col):
    """Lee el valor de una celda, retorna cadena vacia si es None."""
    v = ws.cell(row=row, column=col).value
    if v is None:
        return ""
    return str(v).strip()


def _find_label_row(ws, label_text, col=1, start=1, end=None):
    """Busca la fila que contiene un texto de etiqueta en la columna dada."""
    if end is None:
        end = ws.max_row
    label_lower = label_text.lower()
    for r in range(start, end + 1):
        v = ws.cell(row=r, column=col).value
        if v and label_lower in str(v).lower():
            return r
    return None


def _read_label_value_pairs(ws, start_row, end_row, label_col=1, value_col=2):
    """Lee pares etiqueta-valor de filas consecutivas."""
    result = {}
    for r in range(start_row, end_row + 1):
        label = _val(ws, r, label_col)
        value = _val(ws, r, value_col)
        if label:
            result[label] = value
    return result


def _parse_datos_generales(ws):
    """Parsea los datos generales comunes a todas las fichas DT."""
    dg_start = _find_label_row(ws, "Fecha")
    if dg_start is None:
        dg_start = 6

    pairs = _read_label_value_pairs(ws, dg_start, dg_start + 6)

    def _get(pairs, *keywords):
        for k, v in pairs.items():
            kl = k.lower()
            for kw in keywords:
                if kw.lower() in kl:
                    return v
        return ""

    return {
        "fecha": _get(pairs, "fecha"),
        "evaluador": _get(pairs, "evaluador", "especialista"),
        "codigo_bloque": _get(pairs, "codigo de bloque", "bloque"),
        "microcuenca": _get(pairs, "microcuenca"),
    }


def _parse_dt01(ws):
    """Parsea la ficha F-DT-01 llenada."""
    datos = _parse_datos_generales(ws)

    all_pairs = _read_label_value_pairs(ws, 1, ws.max_row)

    def _g(*kws):
        for k, v in all_pairs.items():
            kl = k.lower()
            for kw in kws:
                if kw.lower() in kl:
                    return v
        return ""

    datos["forma_terreno"] = _g("forma del terreno")
    datos["pendiente"] = _g("pendiente del terreno", "pendiente")
    datos["posicion_fisiografica"] = _g("posicion fisiografica")
    datos["exposicion_orientacion"] = _g("exposicion", "orientacion")
    datos["paisaje_dominante"] = _g("paisaje dominante")
    datos["rango_altitudinal"] = _g("rango altitudinal")
    datos["observaciones"] = _g("observaciones")

    return datos


def _parse_dt02(ws):
    """Parsea la ficha F-DT-02 llenada."""
    datos = _parse_datos_generales(ws)

    all_pairs = _read_label_value_pairs(ws, 1, ws.max_row)

    def _g(*kws):
        for k, v in all_pairs.items():
            kl = k.lower()
            for kw in kws:
                if kw.lower() in kl:
                    return v
        return ""

    datos["precipitacion_anual"] = _g("precipitacion")
    datos["temperatura_media"] = _g("temperatura")
    datos["humedad_relativa"] = _g("humedad relativa")
    datos["zona_vida"] = _g("zona de vida")
    datos["presencia_heladas"] = _g("heladas")
    datos["regimen_vientos"] = _g("regimen de vientos", "vientos")
    datos["observaciones"] = _g("observaciones")

    return datos


def _parse_dt03(ws):
    """Parsea la ficha F-DT-03 llenada."""
    datos = _parse_datos_generales(ws)

    all_pairs = _read_label_value_pairs(ws, 1, ws.max_row)

    def _g(*kws):
        for k, v in all_pairs.items():
            kl = k.lower()
            for kw in kws:
                if kw.lower() in kl:
                    return v
        return ""

    datos["textura_suelo"] = _g("textura")
    datos["color_suelo"] = _g("color predominante", "color")
    datos["profundidad_efectiva"] = _g("profundidad")
    datos["pedregosidad"] = _g("pedregosidad")
    datos["drenaje"] = _g("drenaje")
    datos["presencia_erosion"] = _g("erosion")
    datos["materia_organica"] = _g("materia organica")
    datos["observaciones"] = _g("observaciones")

    return datos


def _parse_dt04(ws):
    """Parsea la ficha F-DT-04 llenada."""
    datos = _parse_datos_generales(ws)

    all_pairs = _read_label_value_pairs(ws, 1, ws.max_row)

    def _g(*kws):
        for k, v in all_pairs.items():
            kl = k.lower()
            for kw in kws:
                if kw.lower() in kl:
                    return v
        return ""

    datos["tipo_cobertura"] = _g("tipo de cobertura")
    datos["densidad_cobertura"] = _g("densidad de cobertura", "densidad")
    datos["estado_conservacion"] = _g("estado de conservacion")
    datos["uso_actual_suelo"] = _g("uso actual del suelo")
    datos["conflicto_uso"] = _g("estado de uso del suelo", "conflicto")
    datos["observaciones"] = _g("observaciones")

    return datos


def _parse_dt05(ws):
    """Parsea la ficha F-DT-05 llenada."""
    datos = _parse_datos_generales(ws)

    all_pairs = _read_label_value_pairs(ws, 1, ws.max_row)

    def _g(*kws):
        for k, v in all_pairs.items():
            kl = k.lower()
            for kw in kws:
                if kw.lower() in kl:
                    return v
        return ""

    datos["fuente_agua"] = _g("fuente de agua")
    datos["regimen_hidrico"] = _g("regimen hidrico")
    datos["calidad_agua"] = _g("calidad aparente", "calidad")
    datos["distancia_fuente_agua"] = _g("distancia a fuente", "distancia")
    datos["uso_recurso_hidrico"] = _g("uso del recurso hidrico", "uso hidrico")
    datos["observaciones"] = _g("observaciones")

    return datos


def _parse_dt06(ws):
    """Parsea la ficha F-DT-06 llenada."""
    datos = _parse_datos_generales(ws)

    all_pairs = _read_label_value_pairs(ws, 1, ws.max_row)

    def _g(*kws):
        for k, v in all_pairs.items():
            kl = k.lower()
            for kw in kws:
                if kw.lower() in kl:
                    return v
        return ""

    datos["tenencia_tierra"] = _g("tenencia")
    datos["organizacion_comunal"] = _g("organizacion comunal")
    datos["actividad_economica"] = _g("actividad economica")
    datos["accesibilidad_via"] = _g("accesibilidad")
    datos["distancia_centro_poblado"] = _g("distancia al centro")
    datos["servicios_basicos"] = _g("servicios basicos")
    datos["observaciones"] = _g("observaciones")

    return datos


def parsear_excel_dt(file_bytes, ficha=None):
    """
    Parsea un archivo Excel de diagnostico territorial llenado por un tecnico.

    Args:
        file_bytes: Contenido del archivo Excel (bytes o file-like).
        ficha: Ficha especifica a parsear (ej: "F-DT-01").
               Si es None, detecta automaticamente.

    Returns:
        list de dicts, cada uno con:
            - "ficha": codigo de la ficha detectada
            - "datos": dict con todos los datos parseados
    """
    wb = load_workbook(file_bytes, data_only=True)

    parsers = {
        "F-DT-01": _parse_dt01,
        "F-DT-02": _parse_dt02,
        "F-DT-03": _parse_dt03,
        "F-DT-04": _parse_dt04,
        "F-DT-05": _parse_dt05,
        "F-DT-06": _parse_dt06,
    }

    resultados = []

    if ficha and ficha in parsers:
        sheet_names = wb.sheetnames
        target_ws = None
        for sn in sheet_names:
            if ficha.lower().replace("-", "") in sn.lower().replace("-", ""):
                target_ws = wb[sn]
                break
        if target_ws is None and len(sheet_names) == 1:
            target_ws = wb[sheet_names[0]]
        if target_ws:
            datos = parsers[ficha](target_ws)
            resultados.append({"ficha": ficha, "datos": datos})
    else:
        for sn in wb.sheetnames:
            ws = wb[sn]
            ficha_detectada = None
            for fid in parsers:
                fid_clean = fid.lower().replace("-", "")
                sn_clean = sn.lower().replace("-", "").replace(" ", "")
                if fid_clean in sn_clean:
                    ficha_detectada = fid
                    break
            if not ficha_detectada:
                for r in range(1, 4):
                    cv = str(ws.cell(row=r, column=1).value or "")
                    for fid in parsers:
                        if fid in cv:
                            ficha_detectada = fid
                            break
                    if ficha_detectada:
                        break

            if ficha_detectada and ficha_detectada in parsers:
                datos = parsers[ficha_detectada](ws)
                resultados.append({"ficha": ficha_detectada, "datos": datos})

    return resultados


def mapear_dt_a_session_state(datos_parseados, bloques_map):
    """
    Convierte los datos parseados del Excel a un dict de session_state keys
    para autocompletar el formulario de diagnostico territorial.

    Args:
        datos_parseados: dict con "ficha" y "datos" (un elemento de parsear_excel_dt).
        bloques_map: dict {label: id} de bloques disponibles.

    Returns:
        dict con las claves de session_state y sus valores.
    """
    datos = datos_parseados["datos"]
    ss = {}

    # Datos generales
    ss["dt_eval"] = datos.get("evaluador", "")

    # Fecha
    fecha_str = str(datos.get("fecha", ""))
    if fecha_str:
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                ss["dt_fecha"] = datetime.strptime(fecha_str.split(" ")[0], fmt).date()
                break
            except (ValueError, TypeError):
                continue

    # Bloque - buscar en bloques_map
    codigo_bloque = datos.get("codigo_bloque", "")
    if codigo_bloque:
        for label in bloques_map:
            if codigo_bloque in label:
                ss["dt_bl"] = label
                break

    # Microcuenca
    mc = datos.get("microcuenca", "")
    if mc:
        ss["dt_mc"] = mc

    # Mapear parametros a session_state keys segun la ficha
    # F-DT-01
    _map_select(ss, "f01_ft", datos.get("forma_terreno", ""))
    _map_select(ss, "f01_pe", datos.get("pendiente", ""))
    _map_select(ss, "f01_pf", datos.get("posicion_fisiografica", ""))
    _map_select(ss, "f01_ex", datos.get("exposicion_orientacion", ""))
    _map_select(ss, "f01_pa", datos.get("paisaje_dominante", ""))
    _map_select(ss, "f01_ra", datos.get("rango_altitudinal", ""))

    # F-DT-02
    _map_select(ss, "f02_pr", datos.get("precipitacion_anual", ""))
    _map_select(ss, "f02_te", datos.get("temperatura_media", ""))
    _map_select(ss, "f02_hr", datos.get("humedad_relativa", ""))
    _map_select(ss, "f02_zv", datos.get("zona_vida", ""))
    _map_select(ss, "f02_he", datos.get("presencia_heladas", ""))
    _map_select(ss, "f02_vi", datos.get("regimen_vientos", ""))

    # F-DT-03
    _map_select(ss, "f03_tx", datos.get("textura_suelo", ""))
    _map_select(ss, "f03_co", datos.get("color_suelo", ""))
    _map_select(ss, "f03_pr", datos.get("profundidad_efectiva", ""))
    _map_select(ss, "f03_pd", datos.get("pedregosidad", ""))
    _map_select(ss, "f03_dr", datos.get("drenaje", ""))
    _map_select(ss, "f03_er", datos.get("presencia_erosion", ""))
    _map_select(ss, "f03_mo", datos.get("materia_organica", ""))

    # F-DT-04
    _map_select(ss, "f04_tc", datos.get("tipo_cobertura", ""))
    _map_select(ss, "f04_dc", datos.get("densidad_cobertura", ""))
    _map_select(ss, "f04_ec", datos.get("estado_conservacion", ""))
    _map_select(ss, "f04_us", datos.get("uso_actual_suelo", ""))
    _map_select(ss, "f04_cu", datos.get("conflicto_uso", ""))

    # F-DT-05
    _map_select(ss, "f05_fa", datos.get("fuente_agua", ""))
    _map_select(ss, "f05_rh", datos.get("regimen_hidrico", ""))
    _map_select(ss, "f05_ca", datos.get("calidad_agua", ""))
    _map_select(ss, "f05_da", datos.get("distancia_fuente_agua", ""))
    _map_select(ss, "f05_uh", datos.get("uso_recurso_hidrico", ""))

    # F-DT-06
    _map_select(ss, "f06_tt", datos.get("tenencia_tierra", ""))
    _map_select(ss, "f06_oc", datos.get("organizacion_comunal", ""))
    _map_select(ss, "f06_ae", datos.get("actividad_economica", ""))
    _map_select(ss, "f06_ac", datos.get("accesibilidad_via", ""))
    _map_select(ss, "f06_dp", datos.get("distancia_centro_poblado", ""))

    # Servicios basicos (multiselect) - parsear separados por coma
    serv_str = datos.get("servicios_basicos", "")
    if serv_str:
        servicios = [s.strip() for s in serv_str.split(",") if s.strip()]
        # Validar contra opciones conocidas
        servicios_validos = [s for s in servicios if s in FDT06_SERVICIOS]
        if servicios_validos:
            ss["f06_sb"] = servicios_validos

    # Observaciones
    obs = datos.get("observaciones", "")
    if obs:
        ss["dt_obs"] = obs

    return ss


def _map_select(ss, key, value):
    """Mapea un valor a una clave de session_state solo si tiene contenido."""
    if value:
        ss[key] = value
