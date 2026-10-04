"""Helpers openpyxl con el formato institucional ANIN (Proyecto IN Piura).

    import openpyxl
    from anin_excel import encabezado_institucional, escribir_tabla

    wb = openpyxl.Workbook(); ws = wb.active
    fila = encabezado_institucional(ws, "MATRIZ DE BLOQUES", ncols=5)
    escribir_tabla(ws, fila, ["Bloque", "Distrito", "Área (ha)"], filas,
                   totales={"Área (ha)": "SUM"}, formato_num={"Área (ha)": "#,##0.000"})
"""
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

VERDE = "1B4D2E"
AZUL = "1B4F72"
FILA_ALTERNA = "E8F0EA"

ENCABEZADOS = [
    "AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN",
    "DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME",
    "SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN - SESDI",
]
SUBTITULO = ("PROYECTO IN PIURA | CUI 2669244 | Cuenca Alta del Río Piura | "
             "UTM WGS 84 Zona 17S (EPSG:32717)")

_FINO = Side(style="thin", color="999999")
BORDE = Border(left=_FINO, right=_FINO, top=_FINO, bottom=_FINO)


def encabezado_institucional(ws, titulo, ncols, fuente=None, color=VERDE):
    """Escribe los encabezados ANIN, el título y (opcional) la fuente.

    Devuelve la siguiente fila libre (deja una fila en blanco)."""
    textos = ENCABEZADOS + [titulo, SUBTITULO] + ([f"Fuente: {fuente}"] if fuente else [])
    for i, texto in enumerate(textos, start=1):
        c = ws.cell(i, 1, texto)
        if i <= 3:
            c.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=color)
        else:
            c.font = Font(name="Arial", size=10, bold=(i == 4), color=color)
        c.alignment = Alignment(vertical="center", wrap_text=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=max(ncols, 1))
        ws.row_dimensions[i].height = 18
    return len(textos) + 2


def escribir_tabla(ws, fila_ini, columnas, filas, anchos=None, totales=None,
                   formato_num=None, color=VERDE):
    """Tabla con cabecera verde, filas alternas, bordes finos, paneles congelados y filtro.

    totales:     {nombre_columna: "SUM"|"AVERAGE"|"COUNT"|"MAX"|"MIN"} -> fila de totales con fórmula.
    formato_num: {nombre_columna: "#,##0.00"} formato numérico por columna.
    Devuelve la fila siguiente a la tabla."""
    formato_num = formato_num or {}
    for j, nombre in enumerate(columnas, start=1):
        c = ws.cell(fila_ini, j, nombre)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDE
    ws.row_dimensions[fila_ini].height = 30

    for i, fila in enumerate(filas, start=1):
        r = fila_ini + i
        for j, valor in enumerate(fila, start=1):
            c = ws.cell(r, j, valor)
            c.font = Font(name="Arial", size=10)
            c.border = BORDE
            if i % 2 == 0:
                c.fill = PatternFill("solid", fgColor=FILA_ALTERNA)
            fmt = formato_num.get(columnas[j - 1])
            if fmt:
                c.number_format = fmt

    ultima = fila_ini + len(filas)
    siguiente = ultima + 1
    if totales:
        ws.cell(siguiente, 1, "TOTAL").font = Font(name="Arial", size=10, bold=True)
        for j, nombre in enumerate(columnas, start=1):
            c = ws.cell(siguiente, j)
            c.border = BORDE
            c.fill = PatternFill("solid", fgColor="D9E6DD")
            c.font = Font(name="Arial", size=10, bold=True)
            if nombre in totales:
                letra = get_column_letter(j)
                c.value = f"={totales[nombre]}({letra}{fila_ini + 1}:{letra}{ultima})"
                if nombre in formato_num:
                    c.number_format = formato_num[nombre]
        siguiente += 1

    for j, nombre in enumerate(columnas, start=1):
        ancho = (anchos or {}).get(nombre) or min(max(len(str(nombre)) + 4, 14), 40)
        ws.column_dimensions[get_column_letter(j)].width = ancho
    ws.freeze_panes = ws.cell(fila_ini + 1, 1)
    ws.auto_filter.ref = f"A{fila_ini}:{get_column_letter(len(columnas))}{ultima}"
    return siguiente + 1
