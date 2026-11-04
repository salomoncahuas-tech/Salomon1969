# -*- coding: utf-8 -*-
"""Matrices Excel del Plan de Contingencia (formato ANIN, openpyxl)."""
import sys, os, datetime as dt
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", ".claude", "skills", "anin-in-piura", "scripts"))
sys.path.insert(0, os.path.dirname(__file__))
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
import anin_excel
from anin_excel import encabezado_institucional, VERDE, FILA_ALTERNA, BORDE
from datos_plan import DATOS as D, nivel
anin_excel.SUBTITULO = D["meta"]["subtitulo_excel"]

OUT = sys.argv[1] if len(sys.argv) > 1 else "Matrices_Plan_Contingencia_IN_Piura.xlsx"
AMAR = "FFF2CC"; TOTAL = "D9E6DD"; GRIS = "F2F2F2"
F = lambda **k: Font(name="Arial", size=k.pop("size", 10), **k)
WRAP = Alignment(vertical="top", wrap_text=True)
FUENTE = "Equipo formulador del Proyecto IN Piura (ANIN - DIME - SESDI). Avance del Perfil a la SESDI y acuerdos de la reunión del 03/11/2026; D.S. N° 149-2026-PCM (El Peruano, 03/11/2026, pp. 13-25); Comunicados ENFEN N° 18 y 19-2026. Versión 2.0 del plan (04/11/2026)."

def tabla(ws, fila, cols, filas, anchos, formatos=None, alt=True, altura=None, input_cols=()):
    """Cabecera verde, filas alternas, bordes, texto ajustado. Devuelve la fila siguiente."""
    formatos = formatos or {}
    for j, n in enumerate(cols, 1):
        c = ws.cell(fila, j, n); c.font = F(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=VERDE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BORDE
    ws.row_dimensions[fila].height = 32
    for i, r in enumerate(filas, 1):
        for j, v in enumerate(r, 1):
            c = ws.cell(fila + i, j, v); c.font = F(); c.border = BORDE; c.alignment = WRAP
            if alt and i % 2 == 0: c.fill = PatternFill("solid", fgColor=FILA_ALTERNA)
            if j in input_cols: c.fill = PatternFill("solid", fgColor=AMAR)
            if cols[j-1] in formatos: c.number_format = formatos[cols[j-1]]
        if altura: ws.row_dimensions[fila + i].height = altura
    for j, a in enumerate(anchos, 1): ws.column_dimensions[L(j)].width = a
    ws.freeze_panes = ws.cell(fila + 1, 1)
    ws.auto_filter.ref = f"A{fila}:{L(len(cols))}{fila + len(filas)}"
    return fila + len(filas) + 1

def nota(ws, fila, texto, ncols, bold=False, color=None):
    c = ws.cell(fila, 1, texto); c.font = F(size=9, bold=bold, italic=not bold, color=color or "595959"); c.alignment = WRAP
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
    ws.row_dimensions[fila].height = max(15, 13 * (1 + len(texto) // 140))
    return fila + 1

wb = openpyxl.Workbook()

# ---------------------------------------------------------------- 1. Resumen
ws = wb.active; ws.title = "1_Resumen"
f = encabezado_institucional(ws, "PLAN DE CONTINGENCIA DEL ESTUDIO DE PERFIL · RESUMEN Y AVANCE POR ENTREGABLE", 8, fuente=FUENTE)
datos = [("Proyecto", D["meta"]["proyecto"]), ("Código de Idea", "337830"), ("Fase", D["meta"]["fase"]),
         ("Contexto", f"{D['decreto']['norma']} ({D['decreto']['publicacion']}): {D['decreto']['plazo']}"),
         ("Ámbito", D["meta"]["ambito"]), ("Modalidad", D["meta"]["modalidad"]),
         ("Presupuesto total de la inversión", f"S/ {D['meta']['techo']:,}"),
         ("Fecha de corte", D["corte"]), ("Versión del plan", D["meta"]["version"]), ("Elaborado por", D["meta"]["elaborado"])]
for k, v in datos:
    a = ws.cell(f, 1, k); a.font = F(bold=True); a.fill = PatternFill("solid", fgColor=FILA_ALTERNA); a.border = BORDE; a.alignment = WRAP
    b = ws.cell(f, 2, v); b.font = F(); b.border = BORDE; b.alignment = WRAP
    ws.merge_cells(start_row=f, start_column=2, end_row=f, end_column=8); ws.row_dimensions[f].height = 30; f += 1
f += 1
ws.cell(f, 1, "Avance ponderado del Perfil al 03/11/2026 (15 entregables)").font = F(bold=True, color=VERDE); f += 1
cols = ["Entregable", "Capítulo", "Nombre", "Peso", "Avance (%)", "Aporte ponderado (puntos)", "Estado", "Sustento al corte"]
ini = f + 1
filas = [[e[0], D["capitulos"][e[1]], e[2], e[3], e[4], f"=D{ini+i}*E{ini+i}/100", e[5], e[6]] for i, e in enumerate(D["entregables"])]
f = tabla(ws, f, cols, filas, [11, 26, 44, 7, 10, 14, 12, 80], {"Aporte ponderado (puntos)": "0.00"}, altura=42)
fin = ini + len(filas) - 1
ws.cell(f, 1, "TOTAL").font = F(bold=True)
for j in range(1, 9):
    c = ws.cell(f, j); c.border = BORDE; c.fill = PatternFill("solid", fgColor=TOTAL); c.font = F(bold=True)
ws.cell(f, 4, f"=SUM(D{ini}:D{fin})"); ws.cell(f, 6, f"=SUM(F{ini}:F{fin})").number_format = "0.00"
ws.cell(f, 5, f"=F{f}/D{f}*100").number_format = "0.0"; ws.cell(f, 7, "Avance ponderado (%) en la columna E")
f += 2
ws.cell(f, 1, "Programación de referencia (reunión del 03/11/2026) frente a los dos escenarios del Plan de Contingencia").font = F(bold=True, color=VERDE); f += 1
cols = ["Hito", "Programación de referencia", "Escenario optimista (sin FEN o FEN retrasado)", "Escenario base (con FEN desde diciembre)"]
filas = [["Capítulo de Identificación (E1 a E9)", "Diciembre de 2026 (≥ 66.5 %)", "23/12/2026 versión para revisión; D4 y D5 al 15/12/2026; final en marzo de 2027 con el catálogo definitivo", "23/12/2026 versión para revisión; final en mayo de 2027 con el campo residual"],
         ["Tamizaje predial (ruta F-LA)", "Noviembre de 2026 en adelante", "Ventana continua del 09/11/2026 al 19/02/2027", "09/11 al 30/11/2026; ventanas puntuales de enero a marzo; residual del 05/04 al 14/05/2027"],
         ["Capítulos de Formulación y Evaluación (E10 a E14)", "Enero - febrero de 2027 (≥ 95 %)", "Borradores al 26/02/2027; final al 26/03/2027", "Borradores al 26/02 y 26/03/2027; final en mayo de 2027"],
         ["Perfil completo (E15)", "Marzo de 2027 (100 %)", "31/03/2027", "31/05/2027"],
         ["Registro en el Banco de Inversiones y viabilidad", "Abril de 2027", "30/04/2027", "30/06/2027"]]
f = tabla(ws, f, cols, filas, None or [11, 26, 44, 7, 10, 14, 12, 80], altura=48)
f += 1
ws.cell(f, 1, "Escenarios de culminación").font = F(bold=True, color=VERDE); f += 1
cols = ["Escenario", "Supuestos", "Estrategia", "Perfil completo", "Viabilidad", "Variación"]
f = tabla(ws, f, cols, [list(e) for e in D["escenarios"]], [11, 26, 44, 7, 10, 14, 12, 80], altura=70)
ws.column_dimensions["A"].width = 30; ws.column_dimensions["B"].width = 40

# ---------------------------------------------------------------- 2. Cronograma con Gantt semanal (dos escenarios)
colores = {"Fase 0": "95A5A6", "Fase 1": "C0392B", "Fase 2": "1B4D2E", "Fase 3": "1B4F72", "Fase 4": "B7950B", "Fase 5": "7D3C98"}
def hoja_cronograma(nombre, titulo, acts, nota_txt):
    ws = wb.create_sheet(nombre)
    f = encabezado_institucional(ws, titulo, 12, fuente=FUENTE)
    f = nota(ws, f, nota_txt, 12)
    f += 1
    cols = ["ID", "Fase", "Actividad", "Inicio", "Fin", "Días", "Entregables", "Campo", "Responsable", "Hito"]
    ini = f + 1
    filas = []
    for i, a in enumerate(acts):
        r = ini + i
        filas.append([a[0], D["fases"][a[1]][0], a[2], dt.date.fromisoformat(a[3]), dt.date.fromisoformat(a[4]), f"=E{r}-D{r}+1", a[5], a[6], a[7], a[8]])
    w0 = dt.date(2026, 11, 2); semanas = [w0 + dt.timedelta(days=7*k) for k in range(36)]
    f_tab = f
    f = tabla(ws, f, cols, filas, [6, 22, 60, 11, 11, 6, 14, 10, 24, 28], {"Inicio": "DD/MM/YYYY", "Fin": "DD/MM/YYYY"}, altura=45, input_cols=(4, 5))
    fin = ini + len(filas) - 1
    for k, sem in enumerate(semanas):
        col = len(cols) + 1 + k
        c = ws.cell(f_tab, col, sem); c.number_format = "DD/MM"; c.font = F(bold=True, color="FFFFFF", size=8); c.fill = PatternFill("solid", fgColor=VERDE)
        c.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90); c.border = BORDE
        m = ws.cell(f_tab - 1, col, sem); m.number_format = "MMM-YY"; m.font = F(bold=True, size=8, color=VERDE); m.alignment = Alignment(horizontal="center")
        ws.column_dimensions[L(col)].width = 3.2
        for r in range(ini, fin + 1):
            cc = ws.cell(r, col, f'=IF(AND({L(col)}${f_tab}<=$E{r},{L(col)}${f_tab}+6>=$D{r}),"■","")')
            cc.font = F(size=8, color=VERDE); cc.alignment = Alignment(horizontal="center", vertical="center"); cc.border = Border(left=Side(style="hair", color="BBBBBB"), right=Side(style="hair", color="BBBBBB"), top=Side(style="thin", color="999999"), bottom=Side(style="thin", color="999999"))
    ws.row_dimensions[f_tab].height = 40
    rng = f"{L(len(cols)+1)}{ini}:{L(len(cols)+len(semanas))}{fin}"
    for fase, col in colores.items():
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND({L(len(cols)+1)}{ini}="■",LEFT($B{ini},6)="{fase}")'], fill=PatternFill("solid", fgColor=col), font=Font(color=col)))
    ws.freeze_panes = ws.cell(ini, 4)
    f += 1
    ws.cell(f, 1, "Fases del plan").font = F(bold=True, color=VERDE); f += 1
    f_fases = f + 1
    f = tabla(ws, f, ["Fase", "Periodo", "Color"], [[v[0], v[1], ""] for v in D["fases"].values()], None or [6, 22, 60], altura=30)
    for i, (k, col) in enumerate(colores.items()):
        ws.cell(f_fases + i, 3).fill = PatternFill("solid", fgColor=col)
    f += 1
    ws.cell(f, 1, "Hitos").font = F(bold=True, color=VERDE); f += 1
    hitos = [[a[8].split("·")[0].strip(), a[8].split("·")[1].strip() if "·" in a[8] else a[8], a[2]] for a in acts if a[8] != "—"]
    f = tabla(ws, f, ["Hito", "Descripción y fecha", "Actividad asociada"], hitos, [6, 22, 60], altura=30)
    return ws

hoja_cronograma("2_Cronograma_Base", "CRONOGRAMA HASTA LA CULMINACIÓN DEL PERFIL · ESCENARIO BASE (CON FEN DESDE DICIEMBRE) · PERFIL 31/05/2027 · VIABILIDAD 30/06/2027", D["cronogramas"]["base"],
    "Escenario base: lluvias por encima de lo normal desde diciembre; campo para el tamizaje predial solo en noviembre y en ventanas puntuales; campo residual en abril y mayo de 2027. Celdas amarillas editables (fechas); el diagrama semanal se calcula con fórmulas. Fechas referenciales.")
hoja_cronograma("2b_Cronograma_Optimista", "CRONOGRAMA HASTA LA CULMINACIÓN DEL PERFIL · ESCENARIO OPTIMISTA (SIN FEN O FEN RETRASADO) · PERFIL 31/03/2027 · VIABILIDAD 30/04/2027", D["cronogramas"]["optimista"],
    "Escenario optimista: el FEN se retrasa o es de menor intensidad; la ventana de campo para el tamizaje predial se amplía del 09/11/2026 al 19/02/2027 y se mantiene la previsión de la reunión del 03/11/2026 (Perfil al 31/03/2027 y viabilidad al 30/04/2027). Celdas amarillas editables; diagrama con fórmulas.")

# ---------------------------------------------------------------- 2c. Decisiones y acuerdos
ws = wb.create_sheet("2c_Decisiones_Acuerdos")
f = encabezado_institucional(ws, "DECISIONES REQUERIDAS (D1 A D5) Y ACUERDOS DE LA REUNIÓN DEL 03/11/2026", 5, fuente=FUENTE)
f = tabla(ws, f, ["N.°", "Decisión solicitada", "Plazo", "Instancia", "Entregables"], [list(d) for d in D["decisiones"]], [6, 90, 14, 30, 14], altura=55)
f += 1
ws.cell(f, 1, "Acuerdos y notas de la reunión del 03/11/2026 incorporados al plan").font = F(bold=True, color=VERDE); f += 1
f = tabla(ws, f, ["N.°", "Acuerdo"], [[i + 1, a] for i, a in enumerate(D["acuerdos"])], [6, 90], altura=42)

# ---------------------------------------------------------------- 3. Presupuesto
ws = wb.create_sheet("3_Presupuesto")
f = encabezado_institucional(ws, "PRESUPUESTO Y RECURSOS NECESARIOS PARA LA CULMINACIÓN DEL PERFIL · ESTRUCTURA PARA VALORIZACIÓN (CAMPO EN PIURA Y GABINETE EN LIMA)", 12, fuente=FUENTE)
f = nota(ws, f, "Las cantidades son supuestos de planificación (columna 'Base del cálculo') y las celdas amarillas (cantidad y costo unitario) deben completarse o ajustarse por la Administración de la DIME. Solo el viático diario tiene un valor referencial normativo (D.S. N° 007-2013-EF, S/ 320 por día); los demás costos unitarios están [POR DEFINIR] y el total es parcial hasta completarlos. La columna 'Escenario' indica si la partida aplica a ambos escenarios, solo al optimista (ventana de campo ampliada) o solo al base (campo residual). No se han inventado precios.", 12)
cols = ["Código", "Componente", "Partida / recurso", "Unidad", "Cantidad", "Base del cálculo", "Costo unitario (S/)", "Parcial (S/)", "Fuente del costo unitario", "Fase", "Escenario", "Estado del dato"]
ini = f + 1
filas = []
for i, p in enumerate(D["presupuesto"]):
    r = ini + i
    filas.append([p[0], p[1], p[2], p[3], p[4], p[5], p[6] if p[6] is not None else None, f'=IF(AND(ISNUMBER(E{r}),ISNUMBER(G{r})),E{r}*G{r},"")', p[7], p[8], p[9], "Referencial" if p[6] is not None else "[POR DEFINIR]"])
f = tabla(ws, f, cols, filas, [8, 34, 52, 12, 10, 40, 14, 16, 40, 8, 11, 14], {"Cantidad": "#,##0", "Costo unitario (S/)": "#,##0.00", "Parcial (S/)": "#,##0.00"}, altura=45, input_cols=(5, 7))
fin = ini + len(filas) - 1
f += 1
ws.cell(f, 1, "Resumen por componente").font = F(bold=True, color=VERDE); f += 1
comps = sorted(set(p[1] for p in D["presupuesto"]))
cols2 = ["Componente", "Partidas", "Parcial valorizado (S/)", "Partidas sin costo unitario"]
ini2 = f + 1
filas2 = [[c, f'=COUNTIF($B${ini}:$B${fin},A{ini2+i})', f'=SUMIF($B${ini}:$B${fin},A{ini2+i},$H${ini}:$H${fin})', f'=COUNTIFS($B${ini}:$B${fin},A{ini2+i},$G${ini}:$G${fin},"")'] for i, c in enumerate(comps)]
f = tabla(ws, f, cols2, filas2, [8, 34, 52, 12], {"Parcial valorizado (S/)": "#,##0.00"}, altura=22)
fin2 = ini2 + len(filas2) - 1
C = D["cartera"]
res = [("Subtotal valorizado (ambos escenarios; descontar la partida del escenario que no aplique)", f"=SUM(C{ini2}:C{fin2})"), (f"Reserva de contingencia [SUPUESTO {int(D['contingencia_pct']*100)} %]", f"=C{f}*{D['contingencia_pct']}"), ("TOTAL ESTIMADO (parcial hasta completar los costos unitarios)", f"=C{f}+C{f+1}"),
       ("Presupuesto asignado al Estudio de Perfil en el POI/PIA de la ANIN", "[POR DEFINIR]"),
       ("Referencia · cartera de la Matriz v1.6 con diques en cárcavas: inversión + O&M 15 años (S/ M)", C["total"]), ("Referencia · % del presupuesto total de la inversión (S/ 372,442,809)", C["pct"]),
       ("Referencia · brecha a cubrir con la decisión D1 (S/ M)", C["brecha"]), ("Referencia · con la zanja de la ficha ANIN 0.40 × 0.30 m (D1) · % del techo", C["d1_zanja"]), ("Referencia · además sin obras en los 5 bloques con P < 400 mm (D1) · % del techo", C["d1_zanja_p400"])]
for i, (k, v) in enumerate(res):
    a = ws.cell(f + i, 1, k); a.font = F(bold=True); a.fill = PatternFill("solid", fgColor=TOTAL); a.border = BORDE; a.alignment = WRAP
    ws.merge_cells(start_row=f + i, start_column=1, end_row=f + i, end_column=2)
    c = ws.cell(f + i, 3, v); c.font = F(bold=True); c.fill = PatternFill("solid", fgColor=TOTAL); c.border = BORDE
    c.number_format = "#,##0.00" if i < 4 else "0.0"; ws.row_dimensions[f + i].height = 28
f += len(res) + 1
f = nota(ws, f, "Recursos humanos del equipo núcleo: la planilla y los contratos vigentes de la DIME no constan en el repositorio del proyecto; se registran como partida C con cantidad en meses y costo [POR DEFINIR]. Equipo de campo: 6 asistentes prediales (4 contratados y 2 por contratar, D3). Fuente del viático: Decreto Supremo N° 007-2013-EF; verificar la escala interna vigente de la ANIN. La cartera (Resumen Ejecutivo del 03/11/2026) supera el presupuesto total en 5.8 %: el plan no cuenta con margen para estudios hasta que se adopte D1.", 12)

# ---------------------------------------------------------------- 4. Matriz de riesgos
ws = wb.create_sheet("4_Riesgos")
f = encabezado_institucional(ws, "MATRIZ DE RIESGOS DEL PLAN DE CONTINGENCIA · FACTORES QUE PODRÍAN IMPEDIR O DIFICULTAR LA CULMINACIÓN DEL PERFIL", 14, fuente=FUENTE)
f = nota(ws, f, "Probabilidad e impacto en escala 1 (muy bajo) a 5 (muy alto); el nivel se calcula como P × |I| (Bajo < 6, Medio 6-11, Alto 12-19, Muy alto ≥ 20). El impacto negativo (-) identifica una oportunidad (R20). Celdas amarillas editables; el nivel y el semáforo se recalculan.", 14)
cols = ["ID", "Categoría", "Riesgo", "Causa", "Efecto en el Perfil", "Probabilidad (1-5)", "Impacto (1-5)", "Puntaje", "Nivel", "Medidas prospectivas y correctivas (mitigación)", "Medidas reactivas (respuesta)", "Responsable", "Indicador / disparador", "Entregables afectados"]
ini = f + 1
filas = []
for i, r in enumerate(D["riesgos"]):
    k = ini + i
    filas.append([r[0], r[1], r[2], r[3], r[4], r[5], r[6], f"=F{k}*ABS(G{k})", f'=IF(H{k}>=20,"Muy alto",IF(H{k}>=12,"Alto",IF(H{k}>=6,"Medio","Bajo")))', r[7], r[8], r[9], r[10], r[11]])
f = tabla(ws, f, cols, filas, [6, 16, 40, 36, 36, 11, 10, 9, 10, 52, 40, 22, 34, 14], altura=95, input_cols=(6, 7))
fin = ini + len(filas) - 1
dv = DataValidation(type="whole", operator="between", formula1="-5", formula2="5", allow_blank=False); ws.add_data_validation(dv); dv.add(f"F{ini}:G{fin}")
rng = f"I{ini}:I{fin}"
for txt, col, fc in [("Muy alto", "C0392B", "FFFFFF"), ("Alto", "E67E22", "000000"), ("Medio", "F1C40F", "000000"), ("Bajo", "27AE60", "FFFFFF")]:
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{txt}"'], fill=PatternFill("solid", fgColor=col), font=Font(color=fc, bold=True)))
f += 1
ws.cell(f, 1, "Resumen por nivel").font = F(bold=True, color=VERDE); f += 1
ini2 = f + 1
f = tabla(ws, f, ["Nivel", "Cantidad de riesgos"], [[n, f'=COUNTIF($I${ini}:$I${fin},A{ini2+i})'] for i, n in enumerate(["Muy alto", "Alto", "Medio", "Bajo"])], [6, 16], altura=18)
f += 1
ws.cell(f, 1, "Mapa de calor (probabilidad × impacto): cantidad de riesgos").font = F(bold=True, color=VERDE); f += 1
ws.cell(f, 1, "P \\ I").font = F(bold=True)
for j in range(1, 6):
    c = ws.cell(f, 1 + j, j); c.font = F(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=VERDE); c.border = BORDE; c.alignment = Alignment(horizontal="center")
for p in range(5, 0, -1):
    r = f + (6 - p)
    c = ws.cell(r, 1, p); c.font = F(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=VERDE); c.border = BORDE; c.alignment = Alignment(horizontal="center")
    for imp in range(1, 6):
        s = p * imp; col = "C0392B" if s >= 20 else "E67E22" if s >= 12 else "F1C40F" if s >= 6 else "27AE60"
        c = ws.cell(r, 1 + imp, f'=COUNTIFS($F${ini}:$F${fin},{p},$G${ini}:$G${fin},{imp})'); c.fill = PatternFill("solid", fgColor=col); c.border = BORDE; c.alignment = Alignment(horizontal="center"); c.font = F(bold=True)

# ---------------------------------------------------------------- 5. Pendientes por entregable
ws = wb.create_sheet("5_Pendientes_Entregable")
f = encabezado_institucional(ws, "ACTIVIDADES PENDIENTES PROGRAMADAS POR ENTREGABLE · ESTADO AL 03/11/2026 Y VENTANA PROPUESTA (CAMPO SOLO PARA EL TAMIZAJE PREDIAL)", 10, fuente=FUENTE)
ent = {e[0]: e for e in D["entregables"]}
cols = ["Entregable", "Nombre", "Peso", "Avance al 03/11 (%)", "Actividad pendiente", "Depende de campo", "Ventana propuesta", "Producto", "Prioridad", "Dependencia / responsable"]
filas = [[p[0], ent[p[0]][2], ent[p[0]][3], ent[p[0]][4], p[1], p[2], p[3], p[4], p[5], p[6]] for p in D["pendientes"]]
f = tabla(ws, f, cols, filas, [10, 30, 6, 10, 70, 10, 24, 30, 9, 24], altura=55)
rng = f"F{f-len(filas)}:F{f-1}"
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Sí"'], fill=PatternFill("solid", fgColor="F5B7B1")))
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Puntual"'], fill=PatternFill("solid", fgColor="FAD7A0")))
f += 1
ws.cell(f, 1, "Resumen: pendientes por entregable y dependencia de campo").font = F(bold=True, color=VERDE); f += 1
ini2 = f + 1; i0 = f - len(filas) - 2
filas2 = [[e[0], e[2], e[3], e[4], f'=COUNTIF($A${i0+1}:$A${i0+len(filas)},A{ini2+k})', f'=COUNTIFS($A${i0+1}:$A${i0+len(filas)},A{ini2+k},$F${i0+1}:$F${i0+len(filas)},"Sí")', f'=COUNTIFS($A${i0+1}:$A${i0+len(filas)},A{ini2+k},$F${i0+1}:$F${i0+len(filas)},"Puntual")'] for k, e in enumerate(D["entregables"])]
f = tabla(ws, f, ["Entregable", "Nombre", "Peso", "Avance (%)", "Pendientes", "Dependen de campo (tamizaje predial)", "Verificación puntual en zona accesible"], filas2, [10, 30, 6, 10, 70, 10, 24], altura=20)

# ---------------------------------------------------------------- 6. Cambios de alcance y contrataciones
ws = wb.create_sheet("6_Alcance_Contrataciones")
f = encabezado_institucional(ws, "POSIBLES CAMBIOS DE ALCANCE DEL PERFIL Y CONTRATACIONES ESTRICTAMENTE NECESARIAS", 6, fuente=FUENTE)
cols = ["Tema", "Situación actual", "Cambio propuesto", "Efecto en el Perfil", "Decisión requerida de", "Tipo"]
f = tabla(ws, f, cols, [list(c) for c in D["cambios"]], [30, 44, 60, 40, 22, 22], altura=80)

# ---------------------------------------------------------------- 7. Aspectos administrativos
ws = wb.create_sheet("7_Administrativo")
f = encabezado_institucional(ws, "ASPECTOS ADMINISTRATIVOS Y DE GESTIÓN VINCULADOS AL PERFIL · CAMBIOS DE AUTORIDADES Y CONTEXTO INSTITUCIONAL", 6, fuente=FUENTE)
cols = ["Tema", "Situación", "Efecto esperado sobre el Perfil", "Acciones del plan", "Plazo", "Responsable"]
f = tabla(ws, f, cols, [list(a) for a in D["admin"]], [30, 44, 44, 60, 16, 26], altura=85)
f += 1
ws.cell(f, 1, "Calendario institucional y electoral de referencia").font = F(bold=True, color=VERDE); f += 1
cal = [["02/11/2026", "D.S. N° 149-2026-PCM dado en la Casa de Gobierno; publicado el 03/11/2026 (60 días calendario, artículo 1)", "El Peruano, Normas Legales, pp. 13-25"],
       ["30/11/2026", "Plazo de las decisiones D1, D2 y D3 (reunión del 03/11/2026)", "Resumen Ejecutivo con apuntes"],
       ["15/12/2026", "Plazo de las decisiones D4 y D5", "Resumen Ejecutivo con apuntes"],
       ["Inicios de enero de 2027", "Término de la vigencia del Estado de Emergencia, salvo prórroga (D.S. N° 074-2014-PCM)", "Artículo 1"],
       ["06/12/2026", "Segunda Elección Regional (Piura incluida)", "[POR VERIFICAR] con el JNE (prensa: reprogramada desde el 29/11)"],
       ["31/12/2026", "Fin del mandato de autoridades regionales y municipales; cierre del año fiscal", "Ley N° 27867, Ley N° 27972, normativa presupuestal"],
       ["01/01/2027", "Asumen las autoridades regionales y municipales electas (periodo 2027-2030)", "Ley N° 27867, Ley N° 27972"],
       ["Dic 2026 - Ene 2027", "Asambleas eleccionarias de comunidades campesinas (según cada estatuto)", "[POR VERIFICAR] por comunidad"],
       ["Nov 2026 - Ene 2027", "Lluvias por encima de lo normal en la sierra norte occidental; tendencia que continúa en el verano de 2027", "Comunicados ENFEN N° 18-2026 (citado en el decreto) y N° 19-2026"],
       ["Feb 2027", "Magnitud extraordinaria de El Niño Costero (Niño 1+2) hasta febrero; descenso gradual hasta junio de 2027", "Comunicado ENFEN N° 19-2026"]]
f = tabla(ws, f, ["Fecha", "Evento", "Fuente / estado"], cal, [30, 44, 44], altura=30)

# ---------------------------------------------------------------- 8. Seguimiento
ws = wb.create_sheet("8_Seguimiento")
f = encabezado_institucional(ws, "SEGUIMIENTO Y ACTUALIZACIÓN DEL PLAN DE CONTINGENCIA", 4, fuente=FUENTE)
f = tabla(ws, f, ["Instrumento", "Contenido", "Frecuencia", "Responsable"], [list(s) for s in D["seguimiento"]], [30, 70, 34, 30], altura=35)
f += 1
ws.cell(f, 1, "Distritos del ámbito en el Anexo del D.S. N° 149-2026-PCM (verificado en el texto oficial)").font = F(bold=True, color=VERDE); f += 1
f = tabla(ws, f, ["Distrito", "Provincia", "Incluido en el Anexo", "N.° en el Anexo"], [[d, p, "Sí", n] for d, p, n in D["distritos"]], [30, 70, 34, 30], altura=18)
f = nota(ws, f, "Piura: 65 distritos (n.° 1193 a 1257 del Anexo): Ayabaca 10 (1193-1202), Huancabamba 8 (1203-1210), Morropón 10 (1211-1220), Paita, Piura, Sullana, Talara y Sechura.", 4)

# ---------------------------------------------------------------- 9. Notas y fuentes
ws = wb.create_sheet("9_Notas_Fuentes")
f = encabezado_institucional(ws, "NOTAS, SUPUESTOS Y FUENTES", 3, fuente=FUENTE)
notas = [
 "1. Estado del Perfil: avance ponderado del 51.4 % al 03/11/2026 con los pesos de los 15 entregables (HTML de avance a la SESDI y Resumen Ejecutivo con apuntes del 03/11/2026); Capítulo de Identificación al 76.4 %. E8 culminado con resultados preliminares, por actualizar con el AID (D4).",
 "2. Decreto Supremo N° 149-2026-PCM: texto oficial y Anexo verificados en El Peruano (Normas Legales del 03/11/2026, pp. 13-25). Dado el 02/11/2026; 60 días calendario (artículo 1); acciones (artículo 2), financiamiento (artículo 3) y refrendo (artículo 4). Los 15 distritos del proyecto figuran en el Anexo (hoja 8).",
 "3. Escenarios: optimista (sin FEN o FEN retrasado) con ventana de campo ampliada hasta el 19/02/2027, Perfil al 31/03/2027 y viabilidad al 30/04/2027 (previsión de la reunión del 03/11/2026); base (con FEN desde diciembre) con campo solo en noviembre y residual en abril y mayo, Perfil al 31/05/2027 y viabilidad al 30/06/2027.",
 "4. Campo: solo el tamizaje predial requiere salidas de campo, priorizadas por superficie y bloques comunales; los demás entregables se culminan en gabinete, con verificaciones puntuales en zonas accesibles (E4, cárcavas). El inventario de cárcavas se completa con salidas seguras o con imágenes satelitales de alta resolución.",
 "5. Liberación de áreas en el Perfil: predios del Estado (GORE Piura y Agricultura) mediante arreglos institucionales (50 %), comunidades campesinas mediante actas (25 %); predios privados al expediente técnico. Los lotes de uso sostenible se gestionan dentro de cada uno de los 117 bloques V8.",
 "6. Presupuesto: cantidades [SUPUESTO] y costos unitarios [POR DEFINIR]; viático S/ 320 por día (D.S. N° 007-2013-EF) como único valor referencial. La cartera v1.6 con diques supera el presupuesto total (105.8 %, S/ 394.12 M); D1 la devuelve al 97.4 % o 95.0 %.",
 "7. Calendario electoral: ERM del 04/10/2026; segunda elección regional en Piura (prensa: 06/12/2026, [POR VERIFICAR] con el JNE); autoridades regionales y municipales asumen el 01/01/2027; alcaldes electos [POR VERIFICAR] (ONPE/JNE).",
 "8. Riesgos: construidos con los vacíos del Diagnóstico Integral V3, el EVAR, el E5 final, el E2 B3, los acuerdos del 03/11/2026, los Comunicados ENFEN N° 18 y 19-2026 y el decreto.",
 "9. Coordenadas: UTM WGS 84 Zona 17S (EPSG:32717). Áreas del catálogo V8 (09/10/2026): 12,157.47 ha netas en 117 bloques.",
]
for n in notas: f = nota(ws, f, n, 3)
f += 1
ws.cell(f, 1, "Fuentes consultadas").font = F(bold=True, color=VERDE); f += 1
fuentes = [("Proyecto", "HTML de avance del Perfil ante la SESDI (03/11/2026); Resumen Ejecutivo con apuntes de la reunión del 03/11/2026; Diagnóstico Integral IN Piura V3; Matriz de Costos v1.6; E5 final; EVAR final; E2 B3; catálogo V8", "presentaciones/avance_perfil_sesdi_2026-11-03/")]
fuentes += [("Decreto", a, b) for a, b in D["decreto"]["fuentes"]] + [("Clima", a, b) for a, b in D["clima"]["fuentes"]] + [("Electoral", a, b) for a, b in D["electoral"]["fuentes"]]
fuentes += [("Normativa", "Guía General para la Identificación, Formulación y Evaluación de Proyectos de Inversión (DGPMI-MEF, 2022), Anexo 2 GdR-CCC", "https://www.mef.gob.pe"),
            ("Normativa", "Ley N° 29664 (SINAGERD) y D.S. N° 048-2011-PCM; Ley N° 30754; D.S. N° 017-2018-MINAM; R.M. N° 178-2019-MINAM; R.M. N° 00213-2024-MINAM", "—"),
            ("Normativa", "D.S. N° 007-2013-EF (viáticos nacionales); Ley N° 27867 (Gobiernos Regionales); Ley N° 27972 (Municipalidades); Ley N° 24656 (Comunidades Campesinas)", "—")]
f = tabla(ws, f, ["Tipo", "Fuente", "Enlace / ubicación"], [list(x) for x in fuentes], [14, 90, 80], altura=28)

for w in wb.worksheets:
    w.sheet_view.showGridLines = False
    w.page_setup.orientation = "landscape"; w.page_setup.fitToWidth = 1; w.sheet_properties.pageSetUpPr.fitToPage = True
wb.save(OUT); print("ok", OUT)
