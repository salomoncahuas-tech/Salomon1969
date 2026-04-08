"""
Modulo de importacion/exportacion Excel para Diagnostico Social.
Proyecto IN Piura CUI 2669244 | ANIN - DIME - SESDI

Genera plantillas Excel estandarizadas para las fichas F-DS-01 a F-DS-05
y parsea archivos Excel llenados por tecnicos de campo para autocompletar
los formularios del aplicativo.
"""

import io
import json
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


# ── Estilos ───────────────────────────────────────────────────────────────
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
TABLE_HEADER_FILL = PatternFill(start_color="5DADE2", end_color="5DADE2", fill_type="solid")
TABLE_HEADER_FONT = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
INSTRUCCION_FONT = Font(name="Calibri", size=9, italic=True, color="7F8C8D")
AUTOFILL_FILL = PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid")
AUTOFILL_FONT = Font(name="Calibri", size=10, italic=True)

# ── Opciones de listas (deben coincidir con streamlit_app.py) ──────────
FDS01_IDIOMA = ["Castellano", "Quechua", "Bilingue", "Otro"]
FDS01_NIVEL_EDUCATIVO = ["Sin instruccion", "Primaria", "Secundaria", "Superior"]
FDS01_TASA_MIGRACION = ["Alta", "Media", "Baja"]
FDS01_ORGANIZACION = ["Comunidad Campesina", "Caserio", "Centro poblado", "Anexo"]
FDS01_AGUA_POTABLE = ["Red publica", "Pileta", "Manantial", "Rio/acequia", "Otro"]
FDS01_SANEAMIENTO = ["Red de alcantarillado", "Letrina", "Pozo septico", "Campo abierto"]
FDS01_ENERGIA = ["Red publica", "Panel solar", "Sin servicio"]
FDS01_TELECOMUNICACIONES = ["Telefonia movil", "Internet", "Radio", "Sin servicio"]
FDS01_ACCESO_VIAL = ["Carretera asfaltada", "Afirmada", "Trocha", "Camino de herradura"]
FDS01_TRANSPORTE = ["Vehiculo diario", "Interdiario", "Semanal", "Solo particular"]
FDS01_SALUD = ["Hospital", "Centro de salud", "Puesto de salud", "Ninguno"]
FDS01_EDUCACION = ["Inicial", "Primaria", "Secundaria", "Ninguna"]
FDS01_ACTIVIDADES_ECON = ["Agricultura", "Ganaderia", "Foresteria/Lena",
                          "Comercio", "Jornales", "Artesania", "Otra"]
FDS01_DESTINO_PRODUCCION = ["Autoconsumo", "Venta", "Autoconsumo/Venta"]
FDS01_PROBLEMAS_AGUA = ["Escasez", "Contaminacion", "Conflictos de uso",
                        "Infraestructura deficiente", "Ninguno"]
FDS01_USO_RECURSOS_FOREST = ["Lena", "Madera", "Productos forestales no maderables", "No usa"]
FDS01_DISPOSICION = ["Alta", "Media", "Baja", "Condicionada"]
FDS02_TIPO_ACTOR = ["Publico", "Privado", "Soc. Civil"]
FDS02_NIVEL = ["A", "M", "B"]
FDS05_NIVEL = ["A", "M", "B"]
FDS05_ESTADO_CONFLICTO = ["Activo", "Latente"]
FDS05_TIPO_OPORTUNIDAD = ["Social", "Institucional", "Productivo"]


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


def _add_datos_sheet(wb, bloques_data):
    """Crea una hoja oculta '_Datos' con la tabla de referencia de bloques.

    Args:
        wb: Workbook
        bloques_data: lista de tuplas (codigo, microcuenca, provincia, distrito)

    Returns:
        int: numero de bloques agregados
    """
    ws = wb.create_sheet("_Datos")
    ws.cell(row=1, column=1, value="Codigo")
    ws.cell(row=1, column=2, value="Microcuenca")
    ws.cell(row=1, column=3, value="Provincia")
    ws.cell(row=1, column=4, value="Distrito")
    for i, (codigo, mc, prov, dist) in enumerate(bloques_data, start=2):
        ws.cell(row=i, column=1, value=codigo)
        ws.cell(row=i, column=2, value=mc)
        ws.cell(row=i, column=3, value=prov)
        ws.cell(row=i, column=4, value=dist)
    ws.sheet_state = "hidden"
    return len(bloques_data)


def _add_datos_generales(ws, start_row=5, num_bloques=0):
    """Agrega la seccion de datos generales con validacion y autocompletado.

    Si num_bloques > 0, agrega validacion de lista para Codigo de Bloque
    y formulas VLOOKUP para autocompletar Microcuenca, Provincia y Distrito.

    Returns:
        int: fila siguiente disponible
    """
    r = start_row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    _cell(ws, r, 1, "DATOS GENERALES", SECTION_FONT, SECTION_FILL,
          Alignment(horizontal="center"))
    for c in range(1, 5):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    # Campos editables simples
    for label in ("Fecha (AAAA-MM-DD)", "Ficha N\u00b0", "Responsable"):
        _cell(ws, r, 1, label, LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
        r += 1

    # Codigo de Bloque (con validacion dropdown si hay datos)
    bloque_row = r
    _cell(ws, r, 1, "Codigo de Bloque", LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
    if num_bloques > 0:
        last_data_row = num_bloques + 1
        dv = DataValidation(
            type="list",
            formula1=f"_Datos!$A$2:$A${last_data_row}",
            allow_blank=True,
        )
        dv.error = "Seleccione un bloque registrado en el aplicativo"
        dv.errorTitle = "Bloque no valido"
        dv.prompt = "Seleccione un bloque registrado"
        dv.promptTitle = "Bloques disponibles"
        ws.add_data_validation(dv)
        dv.add(f"B{r}")
    r += 1

    # Provincia (VLOOKUP auto-completado)
    _cell(ws, r, 1, "Provincia (auto)", LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    if num_bloques > 0:
        last_data_row = num_bloques + 1
        ws.cell(row=r, column=2).value = (
            f'=IFERROR(VLOOKUP(B{bloque_row},_Datos!$A$2:$D${last_data_row},3,FALSE),"")')
        ws.cell(row=r, column=2).font = AUTOFILL_FONT
        ws.cell(row=r, column=2).fill = AUTOFILL_FILL
        ws.cell(row=r, column=2).border = THIN_BORDER
    else:
        _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
    r += 1

    # Distrito (VLOOKUP auto-completado)
    _cell(ws, r, 1, "Distrito (auto)", LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    if num_bloques > 0:
        last_data_row = num_bloques + 1
        ws.cell(row=r, column=2).value = (
            f'=IFERROR(VLOOKUP(B{bloque_row},_Datos!$A$2:$D${last_data_row},4,FALSE),"")')
        ws.cell(row=r, column=2).font = AUTOFILL_FONT
        ws.cell(row=r, column=2).fill = AUTOFILL_FILL
        ws.cell(row=r, column=2).border = THIN_BORDER
    else:
        _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
    r += 1

    # Campos editables restantes
    for label in ("Centro Poblado / Localidad", "Comunidad Campesina",
                  "Coordenada Este (UTM)", "Coordenada Norte (UTM)",
                  "Altitud (msnm)", "Codigo UBIGEO"):
        _cell(ws, r, 1, label, LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
        r += 1

    # Microcuenca (VLOOKUP auto-completado)
    _cell(ws, r, 1, "Microcuenca (auto)", LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    if num_bloques > 0:
        last_data_row = num_bloques + 1
        ws.cell(row=r, column=2).value = (
            f'=IFERROR(VLOOKUP(B{bloque_row},_Datos!$A$2:$D${last_data_row},2,FALSE),"")')
        ws.cell(row=r, column=2).font = AUTOFILL_FONT
        ws.cell(row=r, column=2).fill = AUTOFILL_FILL
        ws.cell(row=r, column=2).border = THIN_BORDER
    else:
        _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
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


# ══════════════════════════════════════════════════════════════════════════
# GENERACION DE PLANTILLAS
# ══════════════════════════════════════════════════════════════════════════

def _generar_ds01(wb):
    """Genera hoja F-DS-01: Diagnostico Socioeconomico de Centro Poblado."""
    ws = wb.create_sheet("F-DS-01")
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22

    _add_header(ws, "F-DS-01: DIAGNOSTICO SOCIOECONOMICO DE CENTRO POBLADO")
    r = _add_datos_generales(ws, num_bloques=wb._num_bloques)

    # -- Seccion 1: Datos Demograficos --
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    _cell(ws, r, 1, "1. DATOS DEMOGRAFICOS", SECTION_FONT, SECTION_FILL)
    for c in range(1, 5):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    campos_demo = [
        ("N° de familias/viviendas", None),
        ("Nombre del centro poblado", None),
        ("Poblacion Hombres", None),
        ("Poblacion Mujeres", None),
        ("Poblacion Total", None),
        ("Idioma predominante", "Separar con coma si son varios: " + ", ".join(FDS01_IDIOMA)),
        ("Nivel educativo predominante", "Separar con coma: " + ", ".join(FDS01_NIVEL_EDUCATIVO)),
        ("Tasa de migracion (percepcion)", None),
        ("Destino principal migracion", None),
        ("Organizacion comunal", "Separar con coma: " + ", ".join(FDS01_ORGANIZACION)),
        ("Junta directiva vigente", None),
        ("Presidente/a de junta", None),
    ]
    for label, hint in campos_demo:
        _cell(ws, r, 1, label, LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
        if hint:
            ws.cell(row=r, column=2).comment = None  # openpyxl comment
            # Add hint in smaller text merged in column
        r += 1

    # Validaciones para campos con lista
    _add_validation(ws, "B", r - 5, r - 5, FDS01_TASA_MIGRACION)  # Tasa migracion
    _add_validation(ws, "B", r - 2, r - 2, ["Si", "No"])  # Junta directiva

    r += 1

    # -- Seccion 2: Servicios Basicos --
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    _cell(ws, r, 1, "2. SERVICIOS BASICOS E INFRAESTRUCTURA", SECTION_FONT, SECTION_FILL)
    for c in range(1, 5):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    campos_serv = [
        ("Agua potable (tipo)", "Separar con coma: " + ", ".join(FDS01_AGUA_POTABLE)),
        ("Cobertura agua (%)", None),
        ("Saneamiento", "Separar con coma: " + ", ".join(FDS01_SANEAMIENTO)),
        ("Energia electrica (tipo)", "Separar con coma: " + ", ".join(FDS01_ENERGIA)),
        ("Cobertura energia (%)", None),
        ("Telecomunicaciones", "Separar con coma: " + ", ".join(FDS01_TELECOMUNICACIONES)),
        ("Operador telecom", None),
        ("Acceso vial", "Separar con coma: " + ", ".join(FDS01_ACCESO_VIAL)),
        ("Distancia a capital distrital (km)", None),
        ("Transporte", None),
        ("Establecimiento de salud", "Separar con coma: " + ", ".join(FDS01_SALUD)),
        ("Distancia salud (km)", None),
        ("Institucion educativa", "Separar con coma: " + ", ".join(FDS01_EDUCACION)),
    ]
    for label, hint in campos_serv:
        _cell(ws, r, 1, label, LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
        r += 1

    # Validacion transporte
    _add_validation(ws, "B", r - 4, r - 4, FDS01_TRANSPORTE)

    r += 1

    # -- Seccion 3: Actividades Economicas (tabla) --
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    _cell(ws, r, 1, "3. ACTIVIDADES ECONOMICAS", SECTION_FONT, SECTION_FILL)
    for c in range(1, 5):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    # Encabezados de tabla - expandir a 5 columnas
    ws.column_dimensions["E"].width = 18
    act_headers = ["Actividad", "% Familias", "Productos", "Destino", "Ingreso Est."]
    for i, h in enumerate(act_headers):
        _cell(ws, r, i + 1, h, TABLE_HEADER_FONT, TABLE_HEADER_FILL,
              Alignment(horizontal="center"), THIN_BORDER)
    r += 1
    act_start = r
    for _ in range(7):  # 7 filas para actividades
        for c in range(1, 6):
            _cell(ws, r, c, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
        r += 1
    _add_validation(ws, "A", act_start, act_start + 6, FDS01_ACTIVIDADES_ECON)
    _add_validation(ws, "D", act_start, act_start + 6, FDS01_DESTINO_PRODUCCION)

    r += 1

    # -- Seccion 4: Recursos Naturales --
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    _cell(ws, r, 1, "4. RELACION CON RECURSOS NATURALES Y AGUA", SECTION_FONT, SECTION_FILL)
    for c in range(1, 5):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    campos_nat = [
        ("Fuente principal de agua", None),
        ("Problemas con el agua", "Separar con coma: " + ", ".join(FDS01_PROBLEMAS_AGUA)),
        ("Uso de recursos forestales", "Separar con coma: " + ", ".join(FDS01_USO_RECURSOS_FOREST)),
        ("Frecuencia uso forestal", None),
        ("Percepcion de cambios ambientales", None),
        ("Disposicion a participar en proyecto", None),
        ("Comentario disposicion", None),
    ]
    for label, hint in campos_nat:
        _cell(ws, r, 1, label, LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
        r += 1

    _add_validation(ws, "B", r - 2, r - 2, FDS01_DISPOSICION)  # Disposicion

    r += 1

    # -- Seccion 4b: Activos Asociados y Tenencia --
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    _cell(ws, r, 1, "4b. ACTIVOS ASOCIADOS Y PROPIEDADES DE AREAS A INTERVENIR",
          SECTION_FONT, SECTION_FILL)
    for c in range(1, 5):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    # Activos asociados (campo de texto amplio)
    _cell(ws, r, 1, "Activos asociados al bloque", LABEL_FONT, LABEL_FILL,
          Alignment(wrap_text=True), THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=4)
    _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
    r += 3

    # Instruccion para activos
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    _cell(ws, r, 1,
          "Indique los activos vinculados al bloque: cultivos, ganado, infraestructura de riego, "
          "vias, viviendas, infraestructura publica, etc. (critico para analisis de exposicion - "
          "Guia GRD-CC)",
          INSTRUCCION_FONT, None, Alignment(horizontal="left", wrap_text=True))
    r += 1

    # Propiedades de areas a intervenir por regimen de tenencia
    _cell(ws, r, 1, "Propiedades de areas a intervenir",
          LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    _cell(ws, r, 2, "(Superficie en hectareas por regimen de tenencia)",
          INSTRUCCION_FONT, LABEL_FILL, border=THIN_BORDER)
    r += 1

    for label_ten in ("Area comunal (ha)", "Area privada (ha)", "Area estatal (ha)"):
        _cell(ws, r, 1, label_ten, LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
        r += 1

    r += 1
    # Observaciones
    _cell(ws, r, 1, "Observaciones generales", LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=4)
    _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)


def _generar_ds02(wb):
    """Genera hoja F-DS-02: Identificacion de Actores Clave."""
    ws = wb.create_sheet("F-DS-02")
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 22

    _add_header(ws, "F-DS-02: IDENTIFICACION Y CARACTERIZACION DE ACTORES CLAVE", max_col=7)
    r = _add_datos_generales(ws, num_bloques=wb._num_bloques)

    # Tabla de actores
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    _cell(ws, r, 1, "REGISTRO DE ACTORES IDENTIFICADOS", SECTION_FONT, SECTION_FILL)
    for c in range(1, 8):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    act_headers = ["Nombre/Organizacion", "Tipo", "Rol/Funcion",
                   "Rel. Proyecto", "Influencia", "Interes", "Contacto"]
    for i, h in enumerate(act_headers):
        _cell(ws, r, i + 1, h, TABLE_HEADER_FONT, TABLE_HEADER_FILL,
              Alignment(horizontal="center", wrap_text=True), THIN_BORDER)
    r += 1
    act_start = r
    for _ in range(20):
        for c in range(1, 8):
            _cell(ws, r, c, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
        r += 1
    _add_validation(ws, "B", act_start, act_start + 19, FDS02_TIPO_ACTOR)
    _add_validation(ws, "E", act_start, act_start + 19, FDS02_NIVEL)
    _add_validation(ws, "F", act_start, act_start + 19, FDS02_NIVEL)

    r += 1

    # Clasificacion por tipo
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    _cell(ws, r, 1, "CLASIFICACION POR TIPO DE ACTOR", SECTION_FONT, SECTION_FILL)
    for c in range(1, 8):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    clasificaciones = [
        "Gobierno Local (Municipalidades)",
        "Gobierno Regional (Gerencias/Direcciones)",
        "Gobierno Nacional (ANA, SERNANP, MINAM, MIDAGRI)",
        "Comunidades Campesinas",
        "Juntas de Usuarios de Riego",
        "Comites de Gestion de Cuenca",
        "ONG / Cooperacion Internacional",
        "Empresa Privada (EPS, agroindustria)",
        "Instituciones Educativas / Universidades",
        "Organizaciones de Base (rondas campesinas, comites)",
    ]
    for cat in clasificaciones:
        _cell(ws, r, 1, cat, LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
        r += 1

    r += 1
    _cell(ws, r, 1, "Observaciones generales", LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=7)
    _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)


def _generar_ds03(wb):
    """Genera hoja F-DS-03: Guia de Entrevista Semiestructurada."""
    ws = wb.create_sheet("F-DS-03")
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 20

    _add_header(ws, "F-DS-03: GUIA DE ENTREVISTA SEMIESTRUCTURADA A ACTORES")
    r = _add_datos_generales(ws, num_bloques=wb._num_bloques)

    # Datos del entrevistado
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    _cell(ws, r, 1, "DATOS DEL ENTREVISTADO", SECTION_FONT, SECTION_FILL)
    for c in range(1, 5):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    campos_ent = [
        "Nombre del entrevistado/a",
        "Cargo / Funcion",
        "Institucion / Organizacion",
        "Telefono / Correo",
        "Duracion de la entrevista",
    ]
    for label in campos_ent:
        _cell(ws, r, 1, label, LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
        r += 1

    r += 1

    # Preguntas agrupadas por seccion
    secciones = [
        ("1. PERCEPCION DEL TERRITORIO Y RECURSOS NATURALES", [
            "1.1 Cuales son los principales recursos naturales de esta zona?",
            "1.2 Ha observado cambios en agua, bosques o suelos en los ultimos 10-20 anios?",
            "1.3 Cuales son los principales problemas ambientales de la comunidad?",
            "1.4 Que zonas son mas importantes para la conservacion del agua?",
        ]),
        ("2. ACTIVIDADES PRODUCTIVAS Y MEDIOS DE VIDA", [
            "2.1 Cuales son las principales actividades economicas de la zona?",
            "2.2 Como se abastecen de agua para riego y consumo? Es suficiente?",
            "2.3 Utilizan productos del bosque? Cuales y con que frecuencia?",
            "2.4 Existen cadenas productivas organizadas? Cuales?",
        ]),
        ("3. ORGANIZACION SOCIAL Y GOBERNANZA", [
            "3.1 Que organizaciones existen en la comunidad? Cuales son las mas activas?",
            "3.2 Como se toman las decisiones sobre el uso del territorio?",
            "3.3 Existen conflictos por el uso del agua o la tierra? Entre quienes?",
            "3.4 Han participado en proyectos similares antes? Cual fue la experiencia?",
            "3.5 Existe experiencia previa en reforestacion en la zona? Describa especies utilizadas, superficie intervenida y resultados obtenidos.",
        ]),
        ("4. CONOCIMIENTO Y EXPECTATIVAS SOBRE EL PROYECTO", [
            "4.1 Tiene conocimiento sobre infraestructura natural o restauracion?",
            "4.2 Que expectativas tiene respecto a un proyecto de esta naturaleza?",
            "4.3 Estaria dispuesto/a a participar o contribuir? De que manera?",
            "4.4 Que condiciones o preocupaciones tendria respecto al proyecto?",
        ]),
        ("5. MECANISMOS DE RETRIBUCION (MERESE)", [
            "5.1 Conoce los mecanismos de retribucion por servicios ecosistemicos?",
            "5.2 Quienes serian los principales beneficiarios?",
            "5.3 Que instituciones podrian contribuir economicamente a la conservacion?",
            "5.4 Existen experiencias previas de pago o compensacion ambiental?",
        ]),
    ]

    for sec_title, preguntas in secciones:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        _cell(ws, r, 1, sec_title, SECTION_FONT, SECTION_FILL)
        for c in range(1, 5):
            ws.cell(row=r, column=c).fill = SECTION_FILL
        r += 1
        for preg in preguntas:
            _cell(ws, r, 1, preg, LABEL_FONT, LABEL_FILL,
                  Alignment(wrap_text=True), THIN_BORDER)
            ws.merge_cells(start_row=r, start_column=2, end_row=r + 1, end_column=4)
            _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
            r += 2
        r += 1

    _cell(ws, r, 1, "Observaciones generales", LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=4)
    _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)


def _generar_ds04(wb):
    """Genera hoja F-DS-04: Acta de Taller Participativo."""
    ws = wb.create_sheet("F-DS-04")
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14

    _add_header(ws, "F-DS-04: FORMATO DE ACTA DE TALLER PARTICIPATIVO", max_col=5)
    r = _add_datos_generales(ws, num_bloques=wb._num_bloques)

    # Datos del taller
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    _cell(ws, r, 1, "DATOS DEL TALLER", SECTION_FONT, SECTION_FILL)
    for c in range(1, 6):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    campos_taller = [
        "Lugar del taller", "Convocante",
        "Hora de inicio", "Hora de finalizacion",
        "Objetivo del taller",
    ]
    for label in campos_taller:
        _cell(ws, r, 1, label, LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
        r += 1

    r += 1

    # Lista de participantes
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    _cell(ws, r, 1, "1. LISTA DE PARTICIPANTES", SECTION_FONT, SECTION_FILL)
    for c in range(1, 6):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    part_headers = ["Nombres y Apellidos", "DNI", "Institucion/Comunidad", "Cargo", "Telefono"]
    for i, h in enumerate(part_headers):
        _cell(ws, r, i + 1, h, TABLE_HEADER_FONT, TABLE_HEADER_FILL,
              Alignment(horizontal="center"), THIN_BORDER)
    r += 1
    for _ in range(30):
        for c in range(1, 6):
            _cell(ws, r, c, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
        r += 1

    r += 1

    # Desarrollo del taller
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    _cell(ws, r, 1, "2. DESARROLLO DEL TALLER", SECTION_FONT, SECTION_FILL)
    for c in range(1, 6):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    campos_des = [
        "Presentacion del proyecto y objetivos",
        "Principales intervenciones de los participantes",
        "Preguntas y respuestas",
        "Acuerdos y compromisos",
        "Observaciones",
    ]
    for label in campos_des:
        _cell(ws, r, 1, label, LABEL_FONT, LABEL_FILL,
              Alignment(wrap_text=True), THIN_BORDER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=5)
        _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
        r += 3

    r += 1
    _cell(ws, r, 1, "Observaciones generales", LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=5)
    _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)


def _generar_ds05(wb):
    """Genera hoja F-DS-05: Conflictos y Oportunidades."""
    ws = wb.create_sheet("F-DS-05")
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 22

    _add_header(ws, "F-DS-05: IDENTIFICACION DE CONFLICTOS Y OPORTUNIDADES", max_col=6)
    r = _add_datos_generales(ws, num_bloques=wb._num_bloques)

    # Tabla de conflictos
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    _cell(ws, r, 1, "1. IDENTIFICACION DE CONFLICTOS", SECTION_FONT, SECTION_FILL)
    for c in range(1, 7):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    conf_headers = ["Tipo de Conflicto", "Actores Involucrados", "Nivel (A/M/B)",
                    "Estado", "Descripcion / Causa", "Impacto en proyecto"]
    for i, h in enumerate(conf_headers):
        _cell(ws, r, i + 1, h, TABLE_HEADER_FONT, TABLE_HEADER_FILL,
              Alignment(horizontal="center", wrap_text=True), THIN_BORDER)
    r += 1
    conf_start = r
    for _ in range(10):
        for c in range(1, 7):
            _cell(ws, r, c, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
        r += 1
    _add_validation(ws, "C", conf_start, conf_start + 9, FDS05_NIVEL)
    _add_validation(ws, "D", conf_start, conf_start + 9, FDS05_ESTADO_CONFLICTO)

    r += 1

    # Tabla de oportunidades
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    _cell(ws, r, 1, "2. IDENTIFICACION DE OPORTUNIDADES", SECTION_FONT, SECTION_FILL)
    for c in range(1, 7):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1

    opor_headers = ["Oportunidad Identificada", "Actores Relacionados", "Tipo",
                    "Potencial (A/M/B)", "Como aprovecharla", ""]
    for i, h in enumerate(opor_headers[:5]):
        _cell(ws, r, i + 1, h, TABLE_HEADER_FONT, TABLE_HEADER_FILL,
              Alignment(horizontal="center", wrap_text=True), THIN_BORDER)
    r += 1
    opor_start = r
    for _ in range(10):
        for c in range(1, 6):
            _cell(ws, r, c, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)
        r += 1
    _add_validation(ws, "C", opor_start, opor_start + 9, FDS05_TIPO_OPORTUNIDAD)
    _add_validation(ws, "D", opor_start, opor_start + 9, FDS05_NIVEL)

    r += 1
    _cell(ws, r, 1, "Observaciones generales", LABEL_FONT, LABEL_FILL, border=THIN_BORDER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=5)
    _cell(ws, r, 2, "", VALUE_FONT, VALUE_FILL, border=THIN_BORDER)


def generar_plantilla_ds(fichas=None, bloques_data=None):
    """
    Genera un archivo Excel con plantillas para las fichas especificadas.

    Args:
        fichas: Lista de fichas a generar (ej: ["F-DS-01", "F-DS-03"]).
                Si es None, genera todas (F-DS-01 a F-DS-05).
        bloques_data: Lista de tuplas (codigo, microcuenca, provincia, distrito)
                      para validacion y autocompletado. Si es None, no agrega
                      validacion de bloque.

    Returns:
        bytes: Contenido del archivo Excel listo para descarga.
    """
    if fichas is None:
        fichas = ["F-DS-01", "F-DS-02", "F-DS-03", "F-DS-04", "F-DS-05"]

    wb = Workbook()
    # Eliminar hoja por defecto
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    # Crear hoja de datos de referencia y guardar cantidad en el workbook
    if bloques_data:
        wb._num_bloques = _add_datos_sheet(wb, bloques_data)
    else:
        wb._num_bloques = 0

    generadores = {
        "F-DS-01": _generar_ds01,
        "F-DS-02": _generar_ds02,
        "F-DS-03": _generar_ds03,
        "F-DS-04": _generar_ds04,
        "F-DS-05": _generar_ds05,
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
    """Parsea los datos generales comunes a todas las fichas."""
    # Los datos generales empiezan en fila 6 despues del header (filas 1-3)
    # y la seccion "DATOS GENERALES" (fila 5)
    dg_start = _find_label_row(ws, "Fecha")
    if dg_start is None:
        dg_start = 6

    pairs = _read_label_value_pairs(ws, dg_start, dg_start + 12)

    def _get(pairs, *keywords):
        for k, v in pairs.items():
            kl = k.lower()
            for kw in keywords:
                if kw.lower() in kl:
                    return v
        return ""

    return {
        "fecha": _get(pairs, "fecha"),
        "ficha_numero": _get(pairs, "ficha n"),
        "evaluador": _get(pairs, "responsable"),
        "codigo_bloque": _get(pairs, "codigo de bloque", "bloque"),
        "provincia": _get(pairs, "provincia"),
        "distrito": _get(pairs, "distrito"),
        "centro_poblado": _get(pairs, "centro poblado"),
        "comunidad_campesina": _get(pairs, "comunidad camp"),
        "coordenada_este": _get(pairs, "coordenada este", "este"),
        "coordenada_norte": _get(pairs, "coordenada norte", "norte"),
        "altitud": _get(pairs, "altitud"),
        "codigo_ubigeo": _get(pairs, "ubigeo"),
        "microcuenca": _get(pairs, "microcuenca"),
    }


def _parse_ds01(ws):
    """Parsea la ficha F-DS-01 llenada."""
    datos = _parse_datos_generales(ws)

    # Seccion 1: Datos demograficos
    r_demo = _find_label_row(ws, "DATOS DEMOGRAFICO") or _find_label_row(ws, "familias")
    if r_demo is None:
        r_demo = 20

    # Buscar por labels - leer desde datos generales hasta el final
    all_pairs = _read_label_value_pairs(ws, 6, ws.max_row)

    def _g(*kws):
        for k, v in all_pairs.items():
            kl = k.lower()
            for kw in kws:
                if kw.lower() in kl:
                    return v
        return ""

    datos["ds01_num_familias"] = _g("familias", "viviendas")
    datos["ds01_nombre_cp"] = _g("nombre del centro")
    datos["ds01_poblacion_hombres"] = _g("hombres")
    datos["ds01_poblacion_mujeres"] = _g("mujeres")
    datos["ds01_poblacion_total"] = _g("total")
    datos["ds01_idioma"] = _g("idioma")
    datos["ds01_nivel_educativo"] = _g("nivel educativo")
    datos["ds01_tasa_migracion"] = _g("tasa de migra", "migracion (per")
    datos["ds01_destino_migracion"] = _g("destino principal", "destino migra")
    datos["ds01_organizacion_comunal"] = _g("organizacion comunal")
    datos["ds01_junta_directiva"] = _g("junta directiva")
    datos["ds01_presidente_junta"] = _g("presidente")

    # Servicios
    datos["ds01_agua_potable_tipo"] = _g("agua potable")
    datos["ds01_agua_potable_cobertura"] = _g("cobertura agua")
    datos["ds01_saneamiento"] = _g("saneamiento")
    datos["ds01_energia_tipo"] = _g("energia elect")
    datos["ds01_energia_cobertura"] = _g("cobertura energ")
    datos["ds01_telecomunicaciones"] = _g("telecomunicac")
    datos["ds01_telecom_operador"] = _g("operador")
    datos["ds01_acceso_vial"] = _g("acceso vial")
    datos["ds01_distancia_capital"] = _g("distancia a capital", "distancia capital")
    datos["ds01_transporte"] = _g("transporte")
    datos["ds01_salud_tipo"] = _g("salud")
    datos["ds01_salud_distancia"] = _g("distancia salud")
    datos["ds01_educacion"] = _g("institucion educ", "educativa")

    # Actividades economicas (tabla)
    act_header_row = _find_label_row(ws, "Actividad", col=1,
                                     start=r_demo)
    actividades = []
    if act_header_row:
        for r in range(act_header_row + 1, act_header_row + 8):
            act = _val(ws, r, 1)
            # Omitir encabezados de tabla que se lean como datos
            if act and act.lower() not in ("actividad", "% familias", "productos",
                                           "destino", "ingreso est.", "ingreso est"):
                actividades.append({
                    "actividad": act,
                    "pct_familias": _val(ws, r, 2),
                    "productos": _val(ws, r, 3),
                    "destino": _val(ws, r, 4),
                    "ingreso": _val(ws, r, 5),
                })
    datos["ds01_actividades_economicas"] = actividades

    # Recursos naturales
    datos["ds01_fuente_agua"] = _g("fuente principal")
    datos["ds01_problemas_agua"] = _g("problemas con el agua")
    datos["ds01_uso_recursos_forestales"] = _g("uso de recursos forest")
    datos["ds01_frecuencia_uso_forestal"] = _g("frecuencia uso")
    datos["ds01_percepcion_cambios"] = _g("percepcion de cambio")
    datos["ds01_disposicion_participar"] = _g("disposicion a particip", "disposicion participar")
    datos["ds01_comentario_disposicion"] = _g("comentario disp")

    # Activos asociados y tenencia
    datos["ds01_activos_asociados"] = _g("activos asociados")
    datos["ds01_tenencia_comunal_ha"] = _g("area comunal")
    datos["ds01_tenencia_privada_ha"] = _g("area privada")
    datos["ds01_tenencia_estatal_ha"] = _g("area estatal")

    # Observaciones
    datos["observaciones"] = _g("observaciones general")

    return datos


def _parse_ds02(ws):
    """Parsea la ficha F-DS-02 llenada."""
    datos = _parse_datos_generales(ws)

    # Tabla de actores
    act_header = _find_label_row(ws, "Nombre/Organizacion") or \
                 _find_label_row(ws, "Nombre del Actor")
    actores = []
    if act_header:
        for r in range(act_header + 1, act_header + 21):
            nombre = _val(ws, r, 1)
            if nombre:
                actores.append({
                    "nombre": nombre,
                    "tipo": _val(ws, r, 2),
                    "rol": _val(ws, r, 3),
                    "relacion": _val(ws, r, 4),
                    "influencia": _val(ws, r, 5),
                    "interes": _val(ws, r, 6),
                    "contacto": _val(ws, r, 7),
                })
    datos["ds02_registro_actores"] = actores

    # Clasificacion por tipo
    clasif_start = _find_label_row(ws, "CLASIFICACION POR TIPO") or \
                   _find_label_row(ws, "Gobierno Local")
    if clasif_start:
        # Si encontramos el header de seccion, las categorias empiezan en la siguiente fila
        if "CLASIFICACION" in str(ws.cell(row=clasif_start, column=1).value or "").upper():
            clasif_start += 1

        mapping = [
            ("gob_local", "Gobierno Local"),
            ("gob_regional", "Gobierno Regional"),
            ("gob_nacional", "Gobierno Nacional"),
            ("comunidades", "Comunidades"),
            ("juntas_riego", "Juntas"),
            ("comites_cuenca", "Comites"),
            ("ong", "ONG"),
            ("empresa", "Empresa"),
            ("educacion", "Instituciones Educ"),
            ("org_base", "Organizaciones de Base"),
        ]
        for r in range(clasif_start, min(clasif_start + 12, ws.max_row + 1)):
            cell_val = _val(ws, r, 1)
            for key, prefix in mapping:
                if prefix.lower() in cell_val.lower():
                    datos[f"ds02_actores_{key}"] = _val(ws, r, 2)
                    break

    datos["observaciones"] = ""
    obs_row = _find_label_row(ws, "Observaciones general")
    if obs_row:
        datos["observaciones"] = _val(ws, obs_row, 2)

    return datos


def _parse_ds03(ws):
    """Parsea la ficha F-DS-03 llenada."""
    datos = _parse_datos_generales(ws)

    # Datos del entrevistado
    all_pairs = _read_label_value_pairs(ws, 1, ws.max_row)

    def _g(*kws):
        for k, v in all_pairs.items():
            kl = k.lower()
            for kw in kws:
                if kw.lower() in kl:
                    return v
        return ""

    datos["ds03_nombre_entrevistado"] = _g("nombre del entrevistado")
    datos["ds03_cargo_funcion"] = _g("cargo")
    datos["ds03_institucion"] = _g("institucion", "organizacion")
    datos["ds03_telefono_correo"] = _g("telefono", "correo")
    datos["ds03_duracion"] = _g("duracion")

    # Respuestas - buscar cada pregunta y leer la respuesta de la celda al lado
    preguntas_map = [
        ("ds03_resp_recursos_naturales", "1.1"),
        ("ds03_resp_cambios_ambiente", "1.2"),
        ("ds03_resp_problemas_ambientales", "1.3"),
        ("ds03_resp_zonas_conservacion", "1.4"),
        ("ds03_resp_actividades_economicas", "2.1"),
        ("ds03_resp_abastecimiento_agua", "2.2"),
        ("ds03_resp_productos_bosque", "2.3"),
        ("ds03_resp_cadenas_productivas", "2.4"),
        ("ds03_resp_organizaciones", "3.1"),
        ("ds03_resp_decisiones_territorio", "3.2"),
        ("ds03_resp_conflictos", "3.3"),
        ("ds03_resp_proyectos_anteriores", "3.4"),
        ("ds03_resp_experiencia_reforestacion", "3.5"),
        ("ds03_resp_conocimiento_restauracion", "4.1"),
        ("ds03_resp_expectativas", "4.2"),
        ("ds03_resp_disposicion_participar", "4.3"),
        ("ds03_resp_condiciones", "4.4"),
        ("ds03_resp_conocimiento_merese", "5.1"),
        ("ds03_resp_beneficiarios", "5.2"),
        ("ds03_resp_instituciones_contribuyentes", "5.3"),
        ("ds03_resp_experiencias_pago", "5.4"),
    ]

    for db_field, num_preg in preguntas_map:
        preg_row = _find_label_row(ws, num_preg)
        if preg_row:
            # La respuesta esta en col B de la misma fila
            resp = _val(ws, preg_row, 2)
            if not resp:
                # O puede estar en la fila siguiente si es el formato original
                resp = _val(ws, preg_row + 1, 2) or _val(ws, preg_row + 1, 1)
                # Si la fila siguiente dice "Respuesta:" buscar en col B
                next_val = _val(ws, preg_row + 1, 1)
                if next_val.lower().startswith("respuesta"):
                    resp = _val(ws, preg_row + 1, 2)
            datos[db_field] = resp

    datos["observaciones"] = _g("observaciones general")
    return datos


def _parse_ds04(ws):
    """Parsea la ficha F-DS-04 llenada."""
    datos = _parse_datos_generales(ws)

    all_pairs = _read_label_value_pairs(ws, 1, ws.max_row)

    def _g(*kws):
        for k, v in all_pairs.items():
            kl = k.lower()
            for kw in kws:
                if kw.lower() in kl:
                    return v
        return ""

    datos["ds04_lugar_taller"] = _g("lugar del taller")
    datos["ds04_convocante"] = _g("convocante")
    datos["ds04_hora_inicio"] = _g("hora de inicio")
    datos["ds04_hora_fin"] = _g("hora de finaliz")
    datos["ds04_objetivo"] = _g("objetivo del taller")

    # Participantes (tabla)
    part_header = _find_label_row(ws, "Nombres y Apellidos") or \
                  _find_label_row(ws, "Nombres")
    participantes = []
    if part_header:
        for r in range(part_header + 1, part_header + 31):
            nombre = _val(ws, r, 1)
            if nombre:
                participantes.append({
                    "nombre": nombre,
                    "dni": _val(ws, r, 2),
                    "institucion": _val(ws, r, 3),
                    "cargo": _val(ws, r, 4),
                    "telefono": _val(ws, r, 5),
                })
    datos["ds04_lista_participantes"] = participantes

    # Desarrollo del taller
    datos["ds04_presentacion"] = _g("presentacion del proyecto", "presentacion")
    datos["ds04_intervenciones"] = _g("principales intervenciones", "intervenciones")
    datos["ds04_preguntas_respuestas"] = _g("preguntas y respuestas")
    datos["ds04_acuerdos"] = _g("acuerdos y compromisos", "acuerdos")
    datos["ds04_observaciones_taller"] = _g("observaciones")

    datos["observaciones"] = _g("observaciones general")
    return datos


def _parse_ds05(ws):
    """Parsea la ficha F-DS-05 llenada."""
    datos = _parse_datos_generales(ws)

    # Tabla de conflictos
    conf_header = _find_label_row(ws, "Tipo de Conflicto")
    conflictos = []
    if conf_header:
        for r in range(conf_header + 1, conf_header + 11):
            tipo = _val(ws, r, 1)
            if tipo:
                conflictos.append({
                    "tipo": tipo,
                    "actores": _val(ws, r, 2),
                    "nivel": _val(ws, r, 3),
                    "estado": _val(ws, r, 4),
                    "descripcion": _val(ws, r, 5),
                    "impacto": _val(ws, r, 6),
                })
    datos["ds05_conflictos"] = conflictos

    # Tabla de oportunidades
    opor_header = _find_label_row(ws, "Oportunidad Identificada") or \
                  _find_label_row(ws, "Oportunidad")
    oportunidades = []
    if opor_header:
        for r in range(opor_header + 1, opor_header + 11):
            desc = _val(ws, r, 1)
            if desc:
                oportunidades.append({
                    "oportunidad": desc,
                    "actores": _val(ws, r, 2),
                    "tipo": _val(ws, r, 3),
                    "potencial": _val(ws, r, 4),
                    "como_aprovechar": _val(ws, r, 5),
                })
    datos["ds05_oportunidades"] = oportunidades

    datos["observaciones"] = ""
    obs_row = _find_label_row(ws, "Observaciones general")
    if obs_row:
        datos["observaciones"] = _val(ws, obs_row, 2)

    return datos


def parsear_excel_ds(file_bytes, ficha=None):
    """
    Parsea un archivo Excel de diagnostico social llenado por un tecnico.

    Args:
        file_bytes: Contenido del archivo Excel (bytes o file-like).
        ficha: Ficha especifica a parsear (ej: "F-DS-01").
               Si es None, detecta automaticamente.

    Returns:
        dict con:
            - "ficha": codigo de la ficha detectada
            - "datos_generales": dict con datos comunes
            - "datos_especificos": dict con datos de la ficha
            - "observaciones": texto de observaciones
        o None si no se puede parsear.
    """
    wb = load_workbook(file_bytes, data_only=True)

    parsers = {
        "F-DS-01": _parse_ds01,
        "F-DS-02": _parse_ds02,
        "F-DS-03": _parse_ds03,
        "F-DS-04": _parse_ds04,
        "F-DS-05": _parse_ds05,
    }

    resultados = []

    if ficha and ficha in parsers:
        # Buscar la hoja correspondiente
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
        # Intentar parsear todas las hojas
        for sn in wb.sheetnames:
            ws = wb[sn]
            # Detectar ficha por contenido de celda A2 o nombre de hoja
            ficha_detectada = None
            for fid in parsers:
                fid_clean = fid.lower().replace("-", "")
                sn_clean = sn.lower().replace("-", "").replace(" ", "")
                if fid_clean in sn_clean:
                    ficha_detectada = fid
                    break
            if not ficha_detectada:
                # Buscar en celdas del encabezado
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


def mapear_a_session_state(datos_parseados, bloques_map):
    """
    Convierte los datos parseados del Excel a un dict de session_state keys.

    Args:
        datos_parseados: dict retornado por parsear_excel_ds (un elemento).
        bloques_map: dict {label: id} de bloques disponibles.

    Returns:
        dict con las claves de session_state y sus valores.
    """
    ficha = datos_parseados["ficha"]
    datos = datos_parseados["datos"]
    ss = {}

    # Datos generales
    ss["ds_eval"] = datos.get("evaluador", "")
    ss["ds_fnum"] = datos.get("ficha_numero", "")
    ss["ds_prov"] = datos.get("provincia", "")
    ss["ds_dist"] = datos.get("distrito", "")
    ss["ds_cpob"] = datos.get("centro_poblado", "")
    ss["ds_ccam"] = datos.get("comunidad_campesina", "")
    ss["ds_ubigeo"] = datos.get("codigo_ubigeo", "")
    ss["ds_obs"] = datos.get("observaciones", "")

    try:
        ss["ds_este"] = float(datos.get("coordenada_este") or 0)
    except (ValueError, TypeError):
        ss["ds_este"] = 0.0
    try:
        ss["ds_norte"] = float(datos.get("coordenada_norte") or 0)
    except (ValueError, TypeError):
        ss["ds_norte"] = 0.0
    try:
        ss["ds_alt"] = float(datos.get("altitud") or 0)
    except (ValueError, TypeError):
        ss["ds_alt"] = 0.0

    # Fecha
    fecha_str = datos.get("fecha", "")
    if fecha_str:
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                ss["ds_fecha"] = datetime.strptime(str(fecha_str).split(" ")[0], fmt).date()
                break
            except (ValueError, TypeError):
                continue

    # Microcuenca
    mc = datos.get("microcuenca", "")
    if mc:
        ss["ds_mc"] = mc

    # Buscar bloque por codigo
    codigo_bloque = datos.get("codigo_bloque", "")
    if codigo_bloque:
        for label in bloques_map:
            if codigo_bloque in label:
                ss["ds_bl"] = label
                break

    # Ficha seleccionada
    ss["ds_ficha_sel"] = ficha

    # ── Datos especificos por ficha ──────────────────────────────────
    if ficha == "F-DS-01":
        ss["s01_nf"] = datos.get("ds01_num_familias", "")
        ss["s01_ncp"] = datos.get("ds01_nombre_cp", "")
        ss["s01_ph"] = datos.get("ds01_poblacion_hombres", "")
        ss["s01_pm"] = datos.get("ds01_poblacion_mujeres", "")
        ss["s01_pt"] = datos.get("ds01_poblacion_total", "")

        # Multiselects: convertir cadena "A, B" a lista ["A", "B"]
        def _to_list(val, valid_options=None):
            if isinstance(val, list):
                return val
            items = [x.strip() for x in str(val).split(",") if x.strip()]
            if valid_options:
                items = [x for x in items if x in valid_options]
            return items

        ss["s01_id"] = _to_list(datos.get("ds01_idioma", ""), FDS01_IDIOMA)
        ss["s01_ne"] = _to_list(datos.get("ds01_nivel_educativo", ""), FDS01_NIVEL_EDUCATIVO)
        ss["s01_mi"] = datos.get("ds01_tasa_migracion", "")
        ss["s01_dest"] = datos.get("ds01_destino_migracion", "")
        ss["s01_org"] = _to_list(datos.get("ds01_organizacion_comunal", ""), FDS01_ORGANIZACION)
        ss["s01_jd"] = datos.get("ds01_junta_directiva", "")
        ss["s01_pres"] = datos.get("ds01_presidente_junta", "")

        ss["s01_ag"] = _to_list(datos.get("ds01_agua_potable_tipo", ""), FDS01_AGUA_POTABLE)
        ss["s01_agcob"] = datos.get("ds01_agua_potable_cobertura", "")
        ss["s01_san"] = _to_list(datos.get("ds01_saneamiento", ""), FDS01_SANEAMIENTO)
        ss["s01_en"] = _to_list(datos.get("ds01_energia_tipo", ""), FDS01_ENERGIA)
        ss["s01_encob"] = datos.get("ds01_energia_cobertura", "")
        ss["s01_tel"] = _to_list(datos.get("ds01_telecomunicaciones", ""), FDS01_TELECOMUNICACIONES)
        ss["s01_telop"] = datos.get("ds01_telecom_operador", "")
        ss["s01_via"] = _to_list(datos.get("ds01_acceso_vial", ""), FDS01_ACCESO_VIAL)
        ss["s01_dcap"] = datos.get("ds01_distancia_capital", "")
        ss["s01_tr"] = datos.get("ds01_transporte", "")
        ss["s01_sal"] = _to_list(datos.get("ds01_salud_tipo", ""), FDS01_SALUD)
        ss["s01_sdist"] = datos.get("ds01_salud_distancia", "")
        ss["s01_educ"] = _to_list(datos.get("ds01_educacion", ""), FDS01_EDUCACION)

        # Actividades economicas
        acts = datos.get("ds01_actividades_economicas", [])
        if acts:
            ss["s01_nact"] = max(len(acts), 1)
            for i, a in enumerate(acts):
                ss[f"s01_act{i}"] = a.get("actividad", "")
                ss[f"s01_pct{i}"] = a.get("pct_familias", "")
                ss[f"s01_prod{i}"] = a.get("productos", "")
                ss[f"s01_dest{i}"] = a.get("destino", "")
                ss[f"s01_ing{i}"] = a.get("ingreso", "")

        ss["s01_fag"] = datos.get("ds01_fuente_agua", "")
        ss["s01_pag"] = _to_list(datos.get("ds01_problemas_agua", ""), FDS01_PROBLEMAS_AGUA)
        ss["s01_uf"] = _to_list(datos.get("ds01_uso_recursos_forestales", ""), FDS01_USO_RECURSOS_FOREST)
        ss["s01_ff"] = datos.get("ds01_frecuencia_uso_forestal", "")
        ss["s01_pcam"] = datos.get("ds01_percepcion_cambios", "")
        ss["s01_disp"] = datos.get("ds01_disposicion_participar", "")
        ss["s01_cdisp"] = datos.get("ds01_comentario_disposicion", "")
        ss["s01_activos"] = datos.get("ds01_activos_asociados", "")
        ss["s01_ten_com"] = datos.get("ds01_tenencia_comunal_ha", "")
        ss["s01_ten_pri"] = datos.get("ds01_tenencia_privada_ha", "")
        ss["s01_ten_est"] = datos.get("ds01_tenencia_estatal_ha", "")

    elif ficha == "F-DS-02":
        actores = datos.get("ds02_registro_actores", [])
        if actores:
            ss["s02_nact"] = max(len(actores), 1)
            for i, a in enumerate(actores):
                ss[f"s02_nom{i}"] = a.get("nombre", "")
                ss[f"s02_tip{i}"] = a.get("tipo", "")
                ss[f"s02_rol{i}"] = a.get("rol", "")
                ss[f"s02_rel{i}"] = a.get("relacion", "")
                ss[f"s02_inf{i}"] = a.get("influencia", "")
                ss[f"s02_int{i}"] = a.get("interes", "")
                ss[f"s02_con{i}"] = a.get("contacto", "")

        FDS02_CLASIFICACION = [
            "Gobierno Local (Municipalidades)",
            "Gobierno Regional (Gerencias/Direcciones)",
            "Gobierno Nacional (ANA, SERNANP, MINAM, MIDAGRI)",
            "Comunidades Campesinas",
            "Juntas de Usuarios de Riego",
            "Comites de Gestion de Cuenca",
            "ONG / Cooperacion Internacional",
            "Empresa Privada (EPS, agroindustria)",
            "Instituciones Educativas / Universidades",
            "Organizaciones de Base (rondas campesinas, comites)",
        ]
        for db_key in ["gob_local", "gob_regional", "gob_nacional", "comunidades",
                       "juntas_riego", "comites_cuenca", "ong", "empresa",
                       "educacion", "org_base"]:
            val = datos.get(f"ds02_actores_{db_key}", "")
            if val:
                # Mapear a session_state key usando FDS02_CLASIFICACION
                cat_map = {
                    "gob_local": "Gobierno Local (M",
                    "gob_regional": "Gobierno Regional",
                    "gob_nacional": "Gobierno Nacional",
                    "comunidades": "Comunidades Camp",
                    "juntas_riego": "Juntas de Usuar",
                    "comites_cuenca": "Comites de Gesti",
                    "ong": "ONG / Cooperacio",
                    "empresa": "Empresa Privada",
                    "educacion": "Instituciones Ed",
                    "org_base": "Organizaciones d",
                }
                prefix = cat_map.get(db_key, "")
                for cat in FDS02_CLASIFICACION:
                    if cat.startswith(prefix[:15]):
                        ss[f"s02_cl_{cat[:15]}"] = val
                        break

    elif ficha == "F-DS-03":
        ss["s03_nom"] = datos.get("ds03_nombre_entrevistado", "")
        ss["s03_car"] = datos.get("ds03_cargo_funcion", "")
        ss["s03_inst"] = datos.get("ds03_institucion", "")
        ss["s03_tel"] = datos.get("ds03_telefono_correo", "")
        ss["s03_dur"] = datos.get("ds03_duracion", "")

        resp_map = {
            "s03_r1": "ds03_resp_recursos_naturales",
            "s03_r2": "ds03_resp_cambios_ambiente",
            "s03_r3": "ds03_resp_problemas_ambientales",
            "s03_r4": "ds03_resp_zonas_conservacion",
            "s03_r5": "ds03_resp_actividades_economicas",
            "s03_r6": "ds03_resp_abastecimiento_agua",
            "s03_r7": "ds03_resp_productos_bosque",
            "s03_r8": "ds03_resp_cadenas_productivas",
            "s03_r9": "ds03_resp_organizaciones",
            "s03_r10": "ds03_resp_decisiones_territorio",
            "s03_r11": "ds03_resp_conflictos",
            "s03_r12": "ds03_resp_proyectos_anteriores",
            "s03_r12b": "ds03_resp_experiencia_reforestacion",
            "s03_r13": "ds03_resp_conocimiento_restauracion",
            "s03_r14": "ds03_resp_expectativas",
            "s03_r15": "ds03_resp_disposicion_participar",
            "s03_r16": "ds03_resp_condiciones",
            "s03_r17": "ds03_resp_conocimiento_merese",
            "s03_r18": "ds03_resp_beneficiarios",
            "s03_r19": "ds03_resp_instituciones_contribuyentes",
            "s03_r20": "ds03_resp_experiencias_pago",
        }
        for wk, dbk in resp_map.items():
            ss[wk] = datos.get(dbk, "")

    elif ficha == "F-DS-04":
        ss["s04_lug"] = datos.get("ds04_lugar_taller", "")
        ss["s04_conv"] = datos.get("ds04_convocante", "")
        ss["s04_hi"] = datos.get("ds04_hora_inicio", "")
        ss["s04_hf"] = datos.get("ds04_hora_fin", "")
        ss["s04_obj"] = datos.get("ds04_objetivo", "")
        ss["s04_pres"] = datos.get("ds04_presentacion", "")
        ss["s04_interv"] = datos.get("ds04_intervenciones", "")
        ss["s04_pregs"] = datos.get("ds04_preguntas_respuestas", "")
        ss["s04_acuerd"] = datos.get("ds04_acuerdos", "")
        ss["s04_obs"] = datos.get("ds04_observaciones_taller", "")

        participantes = datos.get("ds04_lista_participantes", [])
        if participantes:
            ss["s04_np"] = max(len(participantes), 1)
            for i, p in enumerate(participantes):
                ss[f"s04_pn{i}"] = p.get("nombre", "")
                ss[f"s04_pd{i}"] = p.get("dni", "")
                ss[f"s04_pi{i}"] = p.get("institucion", "")
                ss[f"s04_pc{i}"] = p.get("cargo", "")
                ss[f"s04_pt{i}"] = p.get("telefono", "")

    elif ficha == "F-DS-05":
        conflictos = datos.get("ds05_conflictos", [])
        if conflictos:
            ss["s05_nc"] = max(len(conflictos), 1)
            for i, c in enumerate(conflictos):
                ss[f"s05_ct{i}"] = c.get("tipo", "")
                ss[f"s05_ca{i}"] = c.get("actores", "")
                ss[f"s05_cn{i}"] = c.get("nivel", "")
                ss[f"s05_ce{i}"] = c.get("estado", "")
                ss[f"s05_cd{i}"] = c.get("descripcion", "")
                ss[f"s05_ci{i}"] = c.get("impacto", "")

        oportunidades = datos.get("ds05_oportunidades", [])
        if oportunidades:
            ss["s05_no"] = max(len(oportunidades), 1)
            for i, o in enumerate(oportunidades):
                ss[f"s05_od{i}"] = o.get("oportunidad", "")
                ss[f"s05_oa{i}"] = o.get("actores", "")
                ss[f"s05_ot{i}"] = o.get("tipo", "")
                ss[f"s05_op{i}"] = o.get("potencial", "")
                ss[f"s05_oc{i}"] = o.get("como_aprovechar", "")

    return ss
