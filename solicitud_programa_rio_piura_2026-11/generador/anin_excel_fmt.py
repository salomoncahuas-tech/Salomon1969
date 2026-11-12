"""Formato Excel ANIN (openpyxl): encabezado institucional, cabecera azul, filas alternas, bordes, paneles."""
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

AZUL, AZUL2, ALT, TOT = "1B4F72", "2E86AB", "EAF2F8", "D6EAF8"
ENC = ["AUTORIDAD NACIONAL DE INFRAESTRUCTURA – ANIN",
       "DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA – DIME",
       "SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN"]
SUB = "PROYECTO IN PIURA | CUI 2669244 | Cuenca Alta del Río Piura | UTM WGS 84 Zona 17S (EPSG:32717)"
fino = Side(style="thin", color="7F8C8D")
BORDE = Border(left=fino, right=fino, top=fino, bottom=fino)


def encabezado(ws, titulo, fuente, ncols):
    for i, t in enumerate(ENC + [SUB, titulo, fuente], start=1):
        ws.cell(row=i, column=1, value=t)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=max(ncols, 2))
        c = ws.cell(row=i, column=1)
        if i <= 3:
            c.font = Font(name="Arial", bold=True, size=10 if i == 1 else 9, color=AZUL)
        elif i == 4:
            c.font = Font(name="Arial", size=8, color=AZUL2, bold=True)
        elif i == 5:
            c.font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=AZUL)
        else:
            c.font = Font(name="Arial", italic=True, size=8, color="566573")
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[i].height = 30
    return 8  # fila de cabecera de tabla


def tabla(ws, fila, cabecera, filas, anchos=None, total_cols=None, total_label="TOTAL", fmt=None):
    """Escribe tabla; total_cols = índices (1-based) que se totalizan con =SUBTOTAL(9,...)/SUM."""
    for j, h in enumerate(cabecera, start=1):
        c = ws.cell(row=fila, column=j, value=h)
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=9)
        c.fill = PatternFill("solid", fgColor=AZUL)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDE
    ws.row_dimensions[fila].height = 32
    for i, r in enumerate(filas, start=1):
        for j, v in enumerate(r, start=1):
            c = ws.cell(row=fila + i, column=j, value=v)
            c.font = Font(name="Arial", size=9)
            c.border = BORDE
            c.alignment = Alignment(vertical="top", wrap_text=isinstance(v, str) and len(v) > 25)
            if i % 2 == 0:
                c.fill = PatternFill("solid", fgColor=ALT)
            if fmt and j in fmt:
                c.number_format = fmt[j]
    last = fila + len(filas)
    if total_cols:
        t = last + 1
        ws.cell(row=t, column=1, value=total_label)
        for j in range(1, len(cabecera) + 1):
            c = ws.cell(row=t, column=j)
            c.font = Font(name="Arial", bold=True, size=9)
            c.fill = PatternFill("solid", fgColor=TOT)
            c.border = BORDE
        for j in total_cols:
            L = get_column_letter(j)
            c = ws.cell(row=t, column=j, value=f"=SUM({L}{fila + 1}:{L}{last})")
            if fmt and j in fmt:
                c.number_format = fmt[j]
        last = t
    if anchos:
        for j, w in enumerate(anchos, start=1):
            ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = ws.cell(row=fila + 1, column=1)
    ws.auto_filter.ref = f"A{fila}:{get_column_letter(len(cabecera))}{fila + len(filas)}"
    ws.sheet_view.showGridLines = False
    return last


def notas(ws, fila, textos):
    for k, t in enumerate(textos):
        c = ws.cell(row=fila + k, column=1, value=t)
        c.font = Font(name="Arial", size=8, italic=True, color="566573")
    return fila + len(textos)
