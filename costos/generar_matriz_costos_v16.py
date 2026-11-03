#!/usr/bin/env python3
"""Matriz de Costos Paramétricos v1.6 — Proyecto IN Piura (CUI 2669244).

Parte de Costos_Parametros_IN_Piura_v1.5.xlsx (libro entregado por el Ing. H. S. Cahuas
Miller con las partidas a eliminar resaltadas en amarillo) y genera
Costos_Parametros_IN_Piura_v1.6.xlsx:

1. Elimina de PARTIDAS y APU_BASE las partidas resaltadas en amarillo (pircado de terrazas
   y riegos de la revegetación de terrazas) y las subpartidas que solo ellas usaban.
2. Toma las metas de infraestructura marrón del plano de obras de conservación de suelos
   (06/10/2026; Matrices_Obras_Conservacion_Suelos_IN_Piura.xlsx) en la hoja nueva
   METAS_MARRON: terrazas de formación lenta, zanjas de infiltración de 0.60 × 0.50 m y
   zanjas de derivación por bloque. Los bloques 74 y 75 (retirados en V6) se registran
   pero no se costean.
3. Rehace CU_COMP: terraza de formación lenta sin pirca (trazo y excavación del bordo),
   zanja ajustada a la sección del plano y barrera viva / revegetación de bordos y
   camellones (fichas ANIN 03.02 y 03.03 sin riego). Registra las partidas desestimadas
   y su sustento.
4. Desestima en O&M la fertilización y las podas de las plantaciones (años = 0).
5. Bloques con precipitación < 400 mm (8, 20, 71, M4B3, M4B4): parámetro P_ARIDO.

Uso:  python3 generar_matriz_costos_v16.py <v1.5.xlsx> <salida_v1.6.xlsx>
Luego recalcular con LibreOffice (ver README del directorio).
"""
import copy
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

sys.path.insert(0, str(Path(__file__).resolve().parent))
from obras_planos import AVISO, BARRERA, RETIRADOS, load as load_obras  # noqa: E402

AMARILLO = "FFFFFF00"
FECHA = "03-nov-2026"
AZUL_INPUT = Font(name="Arial", size=9, color="FF1F4E9E")
FILL_INPUT = PatternFill("solid", fgColor="FFFFF2CC")
FILL_DES = PatternFill("solid", fgColor="FFF7E9E6")


# ----------------------------------------------------------------------------- utilidades
def set_name(wb, name, ref):
    if name in wb.defined_names:
        del wb.defined_names[name]
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def copy_style(src, dst):
    if src.has_style:
        dst.font = copy.copy(src.font)
        dst.fill = copy.copy(src.fill)
        dst.border = copy.copy(src.border)
        dst.alignment = copy.copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy.copy(src.protection)


def is_yellow(cell):
    f = cell.fill
    return f is not None and f.fill_type == "solid" and f.fgColor is not None and f.fgColor.rgb == AMARILLO


def delete_rows(ws, rows):
    """Borra filas desplazando hacia arriba y trasladando las fórmulas relativas."""
    rows = sorted(set(rows))
    runs, start = [], None
    for r in rows:
        if start is None:
            start = prev = r
        elif r == prev + 1:
            prev = r
        else:
            runs.append((start, prev))
            start = prev = r
    if start is not None:
        runs.append((start, prev))
    maxc = ws.max_column
    for a, b in reversed(runs):
        last = ws.max_row
        n = b - a + 1
        if b < last:
            ws.move_range(f"A{b + 1}:{get_column_letter(maxc)}{last}", rows=-n, translate=True)
        for r in range(last - n + 1, last + 1):
            for c in range(1, maxc + 1):
                cell = ws.cell(r, c)
                cell.value = None
                cell.style = "Normal"


def replace_text_everywhere(wb, pairs):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and not c.value.startswith("="):
                    v = c.value
                    for a, b in pairs:
                        v = v.replace(a, b)
                    c.value = v


# ----------------------------------------------------------------------------- 1. partidas en amarillo
def eliminar_amarillas(wb, wbv):
    wp, wpv = wb["PARTIDAS"], wbv["PARTIDAS"]
    borr_part, log = [], []
    for r in range(7, wp.max_row + 1):
        if any(is_yellow(wp.cell(r, c)) for c in range(1, 13)):
            key = wp.cell(r, 1).value
            borr_part.append(r)
            log.append({
                "clave": key, "item": wp.cell(r, 3).value, "desc": wp.cell(r, 4).value,
                "und": wp.cell(r, 5).value, "cu2020": wp.cell(r, 6).value,
                "cu": wpv.cell(r, 8).value, "nc": wpv.cell(r, 9).value, "c": wpv.cell(r, 10).value,
            })
    keys = {x["clave"] for x in log}
    wa = wb["APU_BASE"]
    borr_apu, subs = [], set()
    for r in range(7, wa.max_row + 1):
        k = wa.cell(r, 1).value
        if k is None:
            continue
        if k in keys or any(is_yellow(wa.cell(r, c)) for c in range(1, 19)):
            borr_apu.append(r)
            if str(k).startswith("TERRAZAS_SUB|"):
                subs.add(k)
    end_apu_old, end_part_old = wa.max_row, wp.max_row
    delete_rows(wa, borr_apu)
    delete_rows(wp, borr_part)
    end_apu, end_part = end_apu_old - len(borr_apu), end_part_old - len(borr_part)

    rx_apu = re.compile(r"APU_BASE!\$([A-Z]+)\$7:\$([A-Z]+)\$" + str(end_apu_old) + r"\b")
    rx_part = re.compile(r"PARTIDAS!\$([A-Z]+)\$7:\$([A-Z]+)\$" + str(end_part_old) + r"\b")
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                if ws.title == "APU_BASE":
                    v = re.sub(r"\$" + str(end_apu_old) + r"\b", "$" + str(end_apu), v)
                v = rx_apu.sub(lambda m: f"APU_BASE!${m.group(1)}$7:${m.group(2)}${end_apu}", v)
                if ws.title != "CALCULO":
                    v = rx_part.sub(lambda m: f"PARTIDAS!${m.group(1)}$7:${m.group(2)}${end_part}", v)
                c.value = v
    return log, sorted(subs), len(borr_apu), end_apu, end_part


# ----------------------------------------------------------------------------- 2. parámetros nuevos
NUEVOS_PAR = [
    ("INFRAESTRUCTURA MARRÓN — PLANO DE OBRAS DE CONSERVACIÓN DE SUELOS (06/10/2026)", None, None, None, None),
    ("P_ZI_ANCHO", "Ancho de la zanja de infiltración", 0.6, "m", "Plano de obras, leyenda de cada bloque: «sección 0.6 × 0.5 m». La ficha ANIN es de 0.40 × 0.30 m."),
    ("P_ZI_PROF", "Profundidad de la zanja de infiltración", 0.5, "m", "Plano de obras, leyenda de cada bloque."),
    ("P_ZI_FV", "Factor de volumen de la zanja respecto de la ficha ANIN", "=P_ZI_ANCHO*P_ZI_PROF/(0.4*0.3)", "factor", "Cálculo: escala apertura y camellón de la ficha ANIN (0.40 × 0.30 m) a la sección del plano."),
    ("P_ZI_FT", "Factor de taludes de la zanja respecto de la ficha ANIN", "=P_ZI_PROF/0.3", "factor", "Cálculo: el perfilado ANIN (600 m²/km) corresponde a los dos taludes de 0.30 m."),
    ("P_DERIV", "Incluir zanjas de derivación del plano (1 = sí; 0 = no)", 1, "0/1", "Plano de obras: 1.81 km en 12 bloques; se costean con el CU de zanja."),
    ("P_BVIVA", "Incluir barrera viva y revegetación de bordos y camellones (1 = sí; 0 = no)", 1, "0/1", "Fichas ANIN 03.02 y 03.03 de terrazas y zanjas, sin las partidas de riego. Especie de referencia en METAS_MARRON."),
    ("P_ARIDO", "Bloques con precipitación < 400 mm: 0 = obras y barrera viva; 1 = obras sin barrera viva; 2 = sin obras marrón", 1, "0/1/2", "Plano de obras, nota 8: en 8, 20, 71, M4B3 y M4B4 la infiltración y el establecimiento de la barrera viva son inciertos."),
]


def parametros(wb):
    ws = wb["PARAMETROS"]
    r0 = 50
    hdr_src, par_src = ws.cell(45, 1), [ws.cell(46, c) for c in range(1, 6)]
    for i, (n, d, v, u, f) in enumerate(NUEVOS_PAR):
        r = r0 + i
        if d is None:
            ws.cell(r, 1, n)
            copy_style(hdr_src, ws.cell(r, 1))
            for c in range(2, 6):
                copy_style(ws.cell(45, c), ws.cell(r, c))
            continue
        for c, val in enumerate([n, d, v, u, f], start=1):
            ws.cell(r, c, val)
            copy_style(par_src[c - 1], ws.cell(r, c))
        if not (isinstance(v, str) and v.startswith("=")):
            ws.cell(r, 3).font = AZUL_INPUT
            ws.cell(r, 3).fill = FILL_INPUT
        set_name(wb, n, f"PARAMETROS!$C${r}")
    ws.cell(33, 5).value = ("POR DEFINIR (0 % mientras no se fije). En v1.6 el margen entre el techo y la inversión + O&M "
                            "(RESUMEN) es lo disponible para preinversión y expediente técnico.")


# ----------------------------------------------------------------------------- 3. METAS_MARRON
CLASE = {"a": "árido", "s": "semiárido", "ss": "subhúmedo seco", "sh": "subhúmedo húmedo"}


def metas_marron(wb):
    obras = load_obras()
    src = wb["BLOQUES"]
    ws = wb.create_sheet("METAS_MARRON", index=wb.sheetnames.index("BLOQUES") + 1)
    hdrs = ["Código de bloque", "Hoja del plano", "Provincia", "Distrito", "Microcuenca", "Estado en el catálogo V6",
            "Superficie del bloque (ha)", "Superficie con pendiente < 27° (ha)", "Superficie con obra proyectada (ha)",
            "% con obra", "Precipitación anual (mm)", "Temperatura media (°C)", "ETP (mm)", "Índice de aridez P/ETP",
            "Clase de aridez", "Aviso P < 400 mm", "TFL — zona (ha)", "TFL — longitud trazada (m)",
            "TFL — longitud estimada (m)", "TFL — longitud total (km)", "ZI — zona (ha)", "ZI — longitud trazada (m)",
            "ZI — longitud estimada (m)", "ZI — longitud total (km)", "ZI — volumen de excavación (m³)",
            "ZI — volumen según plano (m³)", "Zanja de derivación (km)", "Gradiente lateral de terrazas (%)",
            "Barrera viva (especie de referencia del plano)"]
    ncol = len(hdrs)
    titles = ["AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN",
              "DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME",
              "SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN - SESDI",
              "PROYECTO IN PIURA – CUI 2669244 · Metas de infraestructura marrón por bloque (plano de obras de conservación de suelos)",
              "Fuente: planos_obras.pdf — «Plano de distribución de obras de conservación de suelos» (120 hojas A3, 06/10/2026) y "
              "Matrices_Obras_Conservacion_Suelos_IN_Piura.xlsx (M1 y M2). TFL = terraza de formación lenta; ZI = zanja de "
              "infiltración (sección del plano: P_ZI_ANCHO × P_ZI_PROF). Trazo indicativo sobre curvas de nivel de 10 m: requiere "
              "replanteo. Los bloques 74 y 75 figuran en el plano pero fueron retirados en el catálogo V6: no se costean. "
              "Azul = dato transcrito del plano; negro = fórmula."]
    for i, t in enumerate(titles, start=1):
        ws.cell(i, 1, t)
        copy_style(src.cell(i, 1), ws.cell(i, 1))
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=ncol)
    ws.row_dimensions[5].height = 42
    for c, h in enumerate(hdrs, start=1):
        cell = ws.cell(6, c, h)
        copy_style(src.cell(6, 2), cell)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[6].height = 58
    data_src = [src.cell(7, c) for c in (1, 5)]
    r = 7
    for i, (code, b) in enumerate(obras.items()):
        estado = "Retirado en V6" if code in RETIRADOS else "Vigente V6"
        vals = [code, i + 2, b["prov"], b["dist"], b["mc"], estado, b["sup"], b["sup27"], b["obra"],
                f"=IF(G{r}=0,0,I{r}/G{r})", b["P"], b["T"], b["ETP"], b["IA"], CLASE[b["clase"]],
                "Sí" if code in AVISO else "No", b["tfl_zona"], b["tfl_traz"], b["tfl_est"], f"=(R{r}+S{r})/1000",
                b["zi_zona"], b["zi_traz"], b["zi_est"], f"=(V{r}+W{r})/1000", f"=(V{r}+W{r})*P_ZI_ANCHO*P_ZI_PROF",
                b["vol_plano"], b["deriv_m"] / 1000, b["grad"] if b["grad"] else None, BARRERA[b["barrera"]]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(r, c, v)
            copy_style(data_src[0] if c <= 6 or c in (15, 16, 29) else data_src[1], cell)
            if not (isinstance(v, str) and v.startswith("=")) and c >= 7 and c not in (15, 16, 29):
                cell.font = Font(name="Arial", size=9, color="FF1F4E9E")
            if c in (10,):
                cell.number_format = "0.0%"
            elif c in (7, 8, 9, 17, 21):
                cell.number_format = "#,##0.0"
            elif c in (18, 19, 22, 23, 25, 26):
                cell.number_format = "#,##0"
            elif c in (20, 24, 27):
                cell.number_format = "#,##0.000"
            elif c == 14:
                cell.number_format = "0.00"
        if code in RETIRADOS:
            for c in range(1, ncol + 1):
                ws.cell(r, c).fill = FILL_DES
        if code in AVISO:
            ws.cell(r, 16).font = Font(name="Arial", size=9, bold=True, color="FFA63A2B")
        r += 1
    last = r - 1
    tot = r
    ws.cell(tot, 1, "TOTAL 117 BLOQUES VIGENTES")
    ws.cell(tot + 1, 1, "Bloques retirados en V6 (74 y 75) — no se costean")
    for rr, crit in ((tot, '"Vigente V6"'), (tot + 1, '"Retirado en V6"')):
        for c in (7, 8, 9, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27):
            L = get_column_letter(c)
            ws.cell(rr, c, f"=SUMIFS({L}7:{L}{last},$F$7:$F${last},{crit})")
            ws.cell(rr, c).number_format = ws.cell(7, c).number_format
        ws.cell(rr, 16, f'=COUNTIFS(P7:P{last},"Sí",$F$7:$F${last},{crit})')
        for c in range(1, ncol + 1):
            copy_style(wb["CALCULO"].cell(124, 1), ws.cell(rr, c)) if c == 1 else None
            ws.cell(rr, c).font = Font(name="Arial", size=9, bold=True)
            ws.cell(rr, c).fill = PatternFill("solid", fgColor="FFDDEBF7")
    ws.cell(tot, 10, f"=IF(G{tot}=0,0,I{tot}/G{tot})")
    ws.cell(tot, 10).number_format = "0.0%"
    widths = [10, 6, 12, 18, 13, 13] + [10] * 22 + [44]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "B7"
    for n, col in (("MM_ID", "A"), ("MM_TFL_KM", "T"), ("MM_ZI_KM", "X"), ("MM_DER_KM", "AA"), ("MM_AVISO", "P"),
                   ("MM_ESTADO", "F"), ("MM_VOL", "Y")):
        set_name(wb, n, f"METAS_MARRON!${col}$7:${col}${last}")
    return obras, tot


# ----------------------------------------------------------------------------- 4. CU_COMP
def cu_comp(wb, end_part, log_amarillas):
    ws = wb["CU_COMP"]
    old = {}
    for r in range(7, 58):
        g, k = ws.cell(r, 1).value, ws.cell(r, 2).value
        if g:
            old.setdefault(g, []).append((k, ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 5).value))
    st_det = [copy.copy(ws.cell(7, c)._style) for c in range(1, 12)]
    st_det2 = [copy.copy(ws.cell(8, c)._style) for c in range(1, 12)]
    st_hdr = [copy.copy(ws.cell(6, c)._style) for c in range(1, 12)]
    st_sum = [copy.copy(ws.cell(60, c)._style) for c in range(1, 12)]
    st_note = copy.copy(ws.cell(68, 1)._style)
    for row in ws.iter_rows(min_row=6, max_row=ws.max_row):
        for c in row:
            c.value = None
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= 6:
            ws.unmerge_cells(str(mr))

    TER, ZAN = "TERRAZAS|", "ZANJAS|"
    det = []
    for g in ("PLANTON", "PLANTA", "HA", "CERCO"):
        det += [(g, k, d, m, b, None, "") for k, d, m, b in old[g]]
    det += [
        ("ZANJA", ZAN + "01.04", "TRAZO, NIVELACION Y REPLANTEO POR KILOMETRO", 360, 360, None, ""),
        ("ZANJA", ZAN + "03.01.01", "APERTURA MANUAL DE ZANJAS (CU ANIN de 0.40x0.30 M, escalado a la sección del plano)", 360, 360,
         "P_ZI_FV", "× P_ZI_FV: volumen de la sección del plano"),
        ("ZANJA", ZAN + "03.01.02", "PERFILADO MANUAL DE TALUDES", 216000, 360, "P_ZI_FT", "× P_ZI_FT: altura de talud del plano"),
        ("ZANJA", ZAN + "03.01.03", "CONFORMACION DEL CAMELLON", 36000, 360, "P_ZI_FV", "× P_ZI_FV"),
        ("TERRAZA", TER + "01.04", "TRAZO, NIVELACION Y REPLANTEO POR KILOMETRO", 120, 120, None, ""),
        ("TERRAZA", TER + "03.01.01", "EXCAVACION MANUAL PARA CONFORMACION DE TERRAZAS (bordo de tierra, sin pirca)", 138000, 120, None,
         "El material excavado forma el bordo"),
    ]
    reveg = [("03.02.01", "VOLTEO Y GRADEO MANUAL DE TERRENO", 25), ("03.02.02", "SIEMBRA AL VOLEO CON CHILLIWA", 25),
             ("03.03.01", "APERTURA DE HOYOS 0.30x0.30x0.30 M. PARA ESQUEJES", 6250),
             ("03.03.02", "SELECCION Y CORTE DE ESQUEJES", 7562), ("03.03.03", "TRANSPORTE DE ESQUEJES HASTA 15 KM", 7562),
             ("03.03.04", "TRATAMIENTO CON HORMONA ENRAIZADORA", 7562), ("03.03.06", "APLICACION Y SIEMBRA DE ESQUEJES", 7562)]
    det += [("BVIVA_T", TER + i, d, m, 120, None, "Ficha ANIN de terrazas, sin riego") for i, d, m in reveg]
    det += [("BVIVA_Z", ZAN + i, d, m, 360, None, "Ficha ANIN de zanjas, sin riego") for i, d, m in reveg]
    det += [("CARCAVA", k, d, m, b, None, "Meta por definir (BLOQUES!AD vacío)") for k, d, m, b in old["CARCAVA"]]

    hdr = ["Grupo", "Clave partida", "Descripción", "Metrado ANIN", "Base ANIN", "Coeficiente por unidad",
           "CU actualizado", "Costo por unidad", "MO no calificada", "MO calificada", "Bienes y servicios", "Nota v1.6"]
    for c, h in enumerate(hdr, start=1):
        ws.cell(6, c, h)._style = copy.copy(st_hdr[min(c, 11) - 1])
    P = f"PARTIDAS!$%s$7:$%s${end_part}"
    rngs = {}
    r = 7
    for i, (g, k, d, m, b, fac, nota) in enumerate(det):
        st = st_det if i % 2 == 0 else st_det2
        vals = [g, k, d, m, b, f"=D{r}/E{r}" + (f"*{fac}" if fac else ""),
                f"=INDEX({P % ('H', 'H')},MATCH(B{r},{P % ('A', 'A')},0))", f"=F{r}*G{r}",
                f"=F{r}*INDEX({P % ('I', 'I')},MATCH(B{r},{P % ('A', 'A')},0))",
                f"=F{r}*INDEX({P % ('J', 'J')},MATCH(B{r},{P % ('A', 'A')},0))", f"=H{r}-I{r}-J{r}", nota]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(r, c, v)
            cell._style = copy.copy(st[min(c, 11) - 1])
        ws.cell(r, 12).font = Font(name="Arial", size=8, italic=True, color="FF6C8494")
        rngs.setdefault(g, [r, r])[1] = r
        r += 1
    r += 1
    sh = r
    for c, h in enumerate(["Compuesto", "Nombre", "Descripción", "", "", "", "", "Costo unitario", "MO no calificada",
                           "MO calificada", "Bienes y servicios"], start=1):
        ws.cell(sh, c, h)._style = copy.copy(st_hdr[c - 1])
    desc = {
        "PLANTON": "S/ por plantón producido en vivero propio (ANIN: 650 000 plantones; infraestructura de vivero amortizada en una campaña)",
        "PLANTA": "S/ por planta instalada (sin plantón)",
        "HA": "S/ por ha plantada (R1 + R2)",
        "CERCO": "S/ por m de cerco",
        "ZANJA": "S/ por km de zanja de infiltración de 0.60 × 0.50 m (plano de obras; ficha ANIN escalada; sin eliminación de desmonte)",
        "TERRAZA": "S/ por km de terraza de formación lenta sin pirca: trazo y excavación para conformar el bordo (ANIN: 120 km)",
        "BVIVA_T": "S/ por km de terraza: barrera viva y revegetación del bordo (ANIN 03.02 y 03.03, sin riego)",
        "BVIVA_Z": "S/ por km de zanja: barrera viva y revegetación del camellón (ANIN 03.02 y 03.03, sin riego)",
        "CARCAVA": "S/ por km de cárcava con diques de piedra canteada (ANIN: 3.24 km) · meta por definir",
    }
    r = sh + 1
    for g in ("PLANTON", "PLANTA", "HA", "CERCO", "ZANJA", "TERRAZA", "BVIVA_T", "BVIVA_Z", "CARCAVA"):
        a, b = rngs[g]
        vals = [g, f"CU_{g}", desc[g], None, None, None, None, f"=SUM(H{a}:H{b})", f"=SUM(I{a}:I{b})",
                f"=SUM(J{a}:J{b})", f"=SUM(K{a}:K{b})"]
        for c, v in enumerate(vals, start=1):
            ws.cell(r, c, v)._style = copy.copy(st_sum[c - 1])
        set_name(wb, f"CU_{g}", f"CU_COMP!$H${r}")
        set_name(wb, f"CU_{g}_NC", f"CU_COMP!$I${r}")
        set_name(wb, f"CU_{g}_C", f"CU_COMP!$J${r}")
        r += 1
    r += 1
    ws.cell(r, 1, "Notas: (1) El plantón se costea con producción en vivero propio. (2) Coeficiente = metrado ANIN ÷ meta base ANIN; "
                  "en zanjas se escala a la sección del plano de obras con P_ZI_FV y P_ZI_FT. (3) Desde v1.6 la terraza de formación lenta "
                  "no lleva pirca ni relleno compactado: el bordo se levanta con el material excavado y la plataforma se forma con el "
                  "sedimento retenido por el bordo y la barrera viva.")._style = st_note
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
    ws.row_dimensions[r].height = 40
    ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")

    # ---- registro de partidas desestimadas
    r += 2
    ws.cell(r, 1, "PARTIDAS ELIMINADAS Y DESESTIMADAS EN v1.6")._style = copy.copy(st_hdr[0])
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
    r += 1
    for c, h in enumerate(["Tratamiento", "Clave partida", "Descripción", "Und", "Metrado por km de obra", "CU (S/)",
                           "Costo evitado por km (S/)", "Sustento"], start=1):
        ws.cell(r, c, h)._style = copy.copy(st_hdr[min(c, 11) - 1])
    ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=11)
    r += 1
    mpk = {"TERRAZAS|03.01.03": 28800 / 120, "TERRAZAS|03.02.03": 75 / 120, "TERRAZAS|03.03.05": 107.42 / 120,
           "TERRAZAS|03.03.07": 20625 / 120}
    sus_y = {"TERRAZAS|03.01.03": "Resaltada en v1.5. La terraza de formación lenta se construye sin pirca de piedra: bordo de tierra con barrera viva.",
             "TERRAZAS|03.02.03": "Resaltada en v1.5. La revegetación del bordo se siembra al inicio de lluvias, sin riego con cisterna.",
             "TERRAZAS|03.03.05": "Resaltada en v1.5. Sin riego de cama: los esquejes se instalan con las lluvias.",
             "TERRAZAS|03.03.07": "Resaltada en v1.5. Sin riego manual de esquejes."}
    filas = []
    for x in log_amarillas:
        filas.append(("Eliminada (amarillo v1.5)", x["clave"], x["desc"], x["und"], mpk.get(x["clave"]), x["cu"],
                      sus_y.get(x["clave"], "Resaltada en v1.5.")))
    PH = f"INDEX({P % ('H', 'H')},MATCH(\"%s\",{P % ('A', 'A')},0))"
    filas += [
        ("Desestimada (Matriz v1.6)", "TERRAZAS|03.01.02", "RELLENO COMPACTADO CON MATERIAL PROPIO", "m3", 132000 / 120,
         "=" + PH % "TERRAZAS|03.01.02",
         "En la terraza de formación lenta la plataforma no se construye por corte y relleno: se forma con el sedimento que retiene el bordo. "
         "El bordo se levanta con el material de la excavación 03.01.01. Validar la estabilidad del bordo en el expediente técnico."),
        ("Desestimada (Matriz v1.6)", "TERRAZAS|03.01.04", "ELIMINACION MANUAL DE DESMONTE DIST. PROM. 0.5 KM", "m3", 7500 / 120,
         "=" + PH % "TERRAZAS|03.01.04",
         "Sin pirca ni relleno no hay excedente que acarrear a 0.5 km: el material queda en el bordo o se esparce en la ladera."),
        ("Desestimada (Matriz v1.6)", "ZANJAS|03.01.04", "ELIMINACION MANUAL DE DESMONTE DIST. PROM. 0.5 KM", "m3", "=9000/360*P_ZI_FV",
         "=" + PH % "ZANJAS|03.01.04",
         "El excedente de la excavación se esparce aguas abajo del camellón; en ladera natural no se requiere acarreo a botadero."),
        ("Desestimada (Matriz v1.6)", "ZANJAS|03.02.03", "RIEGO POR GRAVEDAD CON CISTERNA", "ha", 75 / 360,
         "=" + PH % "ZANJAS|03.02.03", "Mismo criterio que las partidas de riego resaltadas en terrazas."),
        ("Desestimada (Matriz v1.6)", "ZANJAS|03.03.05", "RIEGO MANUAL EN CAMA DE ALMACIGO/REPIQUE", "m2", 107.42 / 360,
         "=" + PH % "ZANJAS|03.03.05", "Mismo criterio que las partidas de riego resaltadas en terrazas."),
        ("Desestimada (Matriz v1.6)", "ZANJAS|03.03.07", "RIEGO MANUAL CON BALDE", "und", 20625 / 360,
         "=" + PH % "ZANJAS|03.03.07", "Mismo criterio que las partidas de riego resaltadas en terrazas."),
    ]
    for f in filas:
        for c, v in enumerate(f[:6], start=1):
            cell = ws.cell(r, c, v)
            cell._style = copy.copy(st_det[min(c, 11) - 1])
            cell.fill = FILL_DES
        ws.cell(r, 7, f"=IFERROR(E{r}*F{r},0)")._style = copy.copy(st_det[7])
        ws.cell(r, 7).fill = FILL_DES
        ws.cell(r, 8, f[6])
        ws.cell(r, 8).font = Font(name="Arial", size=8)
        ws.cell(r, 8).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 8).fill = FILL_DES
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=11)
        ws.row_dimensions[r].height = 34
        r += 1
    ws.column_dimensions["L"].width = 30
    ws.freeze_panes = "D7"


# ----------------------------------------------------------------------------- 5. CALCULO
def calculo(wb):
    ws = wb["CALCULO"]
    for mr in list(ws.merged_cells.ranges):
        if mr.max_row <= 5:
            ws.unmerge_cells(str(mr))
    for i in range(1, 6):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=38)
    ws.cell(5, 1).value = (ws.cell(5, 1).value or "") + (
        " · v1.6: zanjas (K) y terrazas (L) se leen de METAS_MARRON (plano de obras); barrera viva en AJ–AL; "
        "AI marca los bloques con precipitación < 400 mm.")
    new_h = {35: "Aviso P < 400 mm (plano)", 36: "Barrera viva en terrazas (km)", 37: "Barrera viva en zanjas (km)",
             38: "Barrera viva y revegetación (S/)"}
    for c, h in new_h.items():
        ws.cell(6, c, h)
        copy_style(ws.cell(6, 34), ws.cell(6, c))
        ws.column_dimensions[get_column_letter(c)].width = 12
    look = lambda rng, r: f"INDEX({rng},MATCH($A{r},MM_ID,0))"  # noqa: E731
    for r in range(7, 124):
        ar = f'IF(AND({look("MM_AVISO", r)}="Sí",P_ARIDO=2),0,1)'
        ws[f"K{r}"] = f"=IFERROR(({look('MM_ZI_KM', r)}+P_DERIV*{look('MM_DER_KM', r)})*{ar},0)"
        ws[f"L{r}"] = f"=IFERROR({look('MM_TFL_KM', r)}*{ar},0)"
        ws[f"AI{r}"] = f'=IFERROR({look("MM_AVISO", r)},"—")'
        bv = f'P_BVIVA*IF(AND(AI{r}="Sí",P_ARIDO>=1),0,1)'
        ws[f"AJ{r}"] = f"=L{r}*{bv}"
        ws[f"AK{r}"] = f"=K{r}*{bv}"
        ws[f"AL{r}"] = f"=AJ{r}*CU_BVIVA_T+AK{r}*CU_BVIVA_Z"
        ws[f"Y{r}"] = f"=SUM(P{r}:X{r})+AL{r}"
        ws[f"AC{r}"] = ws[f"AC{r}"].value + f"+AJ{r}*CU_BVIVA_T_NC+AK{r}*CU_BVIVA_Z_NC"
        ws[f"AD{r}"] = ws[f"AD{r}"].value + f"+AJ{r}*CU_BVIVA_T_C+AK{r}*CU_BVIVA_Z_C"
        ws[f"AH{r}"] = f"=(U{r}+V{r}+X{r}+AL{r})*(1+P_GG+P_SUP+P_ET)"
        for c in range(35, 39):
            copy_style(ws.cell(r, 34), ws.cell(r, c))
        ws.cell(r, 35).number_format = "@"
        ws.cell(r, 35).alignment = Alignment(horizontal="center")
        ws.cell(r, 36).number_format = ws.cell(r, 37).number_format = "#,##0.00"
    ws["AI124"], ws["AI125"] = '=COUNTIF(AI7:AI123,"Sí")', '=COUNTIFS(AI7:AI123,"Sí",$C$7:$C$123,"Sí")'
    for L in ("AJ", "AK", "AL"):
        ws[f"{L}124"] = f"=SUM({L}7:{L}123)"
        ws[f"{L}125"] = f'=SUMIFS({L}7:{L}123,$C$7:$C$123,"Sí")'
    for rr in (124, 125):
        for c in range(35, 39):
            copy_style(ws.cell(rr, 34), ws.cell(rr, c))
    for n, ref in (("C_BVT_T", "AJ124"), ("C_BVZ_T", "AK124"), ("C_BV_T", "AL124"), ("C_BVT_P", "AJ125"),
                   ("C_BVZ_P", "AK125"), ("C_BV_P", "AL125"), ("C_ARIDO_T", "AI124")):
        set_name(wb, n, "CALCULO!$%s$%s" % (re.match(r"[A-Z]+", ref).group(), re.search(r"\d+", ref).group()))
    ws.freeze_panes = "D7"


# ----------------------------------------------------------------------------- 6. OYM
def oym(wb):
    ws = wb["OYM"]
    for r in range(7, 21):
        ws[f"J{r}"] = f"=IF(G{r}=0,0,F{r}*H{r})"
        ws[f"K{r}"] = f"=IF(G{r}=0,0,F{r}*I{r})"
        act = str(ws.cell(r, 2).value)
        if act.startswith("Podas") or act.startswith("Fertilización"):
            ws.cell(r, 7).value = 0
            for c in range(1, 13):
                ws.cell(r, c).fill = FILL_DES
            motivo = ("restauración con especies nativas, no plantación comercial: sin poda de formación"
                      if act.startswith("Podas") else
                      "especies nativas adaptadas; el abonamiento se hace al plantar (insumos de 04.01.04), sin fertilización de mantenimiento")
            ws.cell(r, 12).value = f"DESESTIMADA en v1.6 (años = 0): {motivo}. " + str(ws.cell(r, 12).value)
    ws.cell(22, 2).value = ("Años de aplicación = 0: actividad desestimada en la Matriz de Costos Paramétricos v1.6 "
                            "(se conserva la tarifa ANIN para el expediente técnico).")
    ws.cell(22, 2).font = Font(name="Arial", size=8, italic=True)


# ----------------------------------------------------------------------------- 7. RESUMEN, LEEME, PAQUETES
def textos(wb, n_apu, log_amarillas, subs):
    ws = wb["RESUMEN"]
    ws.cell(5, 1).value = (f"Generado el {FECHA} (v1.6). Todas las cifras son fórmulas vivas. Metas de infraestructura "
                           "marrón: plano de obras de conservación de suelos (06/10/2026).")
    ws.cell(13, 1).value = "Zanjas de infiltración y de derivación (plano de obras, 0.60 × 0.50 m)"
    ws.cell(14, 1).value = "Terrazas de formación lenta (plano de obras, sin pirca)"
    ws.cell(16, 1).value = "Cárcavas (diques) — meta por definir con levantamiento de campo"
    ws.cell(25, 1).value = "Referencia: techo del proyecto (incluye preinversión, expediente técnico, ejecución y O&M)"
    for c in range(1, 5):
        copy_style(ws.cell(26, c), ws.cell(30, c))
    ws.cell(30, 1).value = "Margen para preinversión y expediente técnico (techo − inversión − O&M)"
    ws.cell(30, 3).value = "=P_TECHO-C24"
    ws.cell(30, 3).number_format = ws.cell(24, 3).number_format
    ws.cell(30, 4).value = "S/"
    notas = [
        "1. Calcula los 117 bloques V6 (12 270.235 ha). Zanjas y terrazas vienen del plano de obras (METAS_MARRON); el piloto (21, 66, 73, 31, M6B10) se calibra primero.",
        "2. Terraza de formación lenta sin pirca (amarillo v1.5): bordo de tierra con barrera viva. Se desestiman además el relleno compactado y la eliminación de desmonte (ver CU_COMP).",
        "3. Cárcavas: meta por definir; ingresar metros lineales por bloque en BLOQUES!AD con el levantamiento de campo o el vuelo de dron recomendado por el E4.",
        "4. O&M: podas y fertilización desestimadas (años = 0); el resto de tarifas ANIN se mantiene. Los años de aplicación siguen siendo supuestos por calibrar.",
        "5. Bloques con P < 400 mm (8, 20, 71, M4B3, M4B4): P_ARIDO = 1 mantiene zanjas y terrazas sin barrera viva; validar infiltración y prendimiento en campo.",
        "6. GG (15.08 %) es la referencia ANIN; supervisión y expediente técnico en 0 %: el margen de la fila 30 es lo disponible para estudios.",
    ]
    for i, t in enumerate(notas):
        ws.cell(47 + i, 1).value = t

    ws = wb["LEEME"]
    ws.cell(5, 1).value = ("Matriz de Costos Paramétricos para el Estudio de Perfil (Invierte.pe) · v1.6 · " + FECHA)
    ws.cell(11, 2).value = ("R1 = plantación plena (MSAVI DN 1–2); R2 = enriquecimiento (DN 3); C = protección (DN 4–5, cerco). "
                            "Marrón (v1.6): terrazas de formación lenta, zanjas de infiltración de 0.60 × 0.50 m y zanjas de derivación "
                            "por bloque según el plano de obras de conservación de suelos (hoja METAS_MARRON), con barrera viva en bordos y camellones.")
    for r, (k, t) in enumerate([
        ("CAMBIOS v1.5", "Libro entregado por el Ing. H. S. Cahuas Miller (03/11/2026) con las partidas a eliminar resaltadas en amarillo: "
                         + "; ".join(f"{x['clave']} {x['desc']}" for x in log_amarillas) + "."),
        ("CAMBIOS v1.6", f"Se eliminaron de PARTIDAS ({len(log_amarillas)} partidas) y de APU_BASE ({n_apu} líneas, con las subpartidas "
                         f"{', '.join(s.split('|')[1][:12] for s in subs)}) las partidas en amarillo. Metas marrón del plano de obras "
                         "(METAS_MARRON, 117 bloques vigentes; 74 y 75 retirados). CU de zanja escalado a 0.60 × 0.50 m. Terraza sin pirca, "
                         "relleno compactado ni eliminación de desmonte; zanja sin eliminación de desmonte; barrera viva con las partidas "
                         "ANIN 03.02 y 03.03 sin riego. O&M sin podas ni fertilización. Bloques con P < 400 mm: parámetro P_ARIDO. "
                         "Registro y sustento en CU_COMP (tabla final)."),
    ], start=20):
        ws.cell(r, 1).value, ws.cell(r, 2).value = k, t
        for c in range(1, 7):
            copy_style(ws.cell(19, c), ws.cell(r, c))
        for mr in list(ws.merged_cells.ranges):
            if mr.min_row == 19 and mr.min_col == 2:
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=mr.max_col)
        ws.row_dimensions[r].height = ws.row_dimensions[19].height or 60

    ws = wb["PAQUETES"]
    ws.cell(5, 1).value = (str(ws.cell(5, 1).value or "") + " Desde v1.6 las metas de zanjas y terrazas se leen del plano de obras "
                           "(METAS_MARRON); los espaciamientos de las columnas G y H quedan como referencia de las fichas ANIN.")
    replace_text_everywhere(wb, [("Motor de costos paramétrico", "Matriz de Costos Paramétricos"),
                                 ("motor de costos paramétrico", "matriz de costos paramétricos"),
                                 ("Motor de costos", "Matriz de Costos Paramétricos"),
                                 ("motor de costos", "matriz de costos paramétricos")])


def main(src, dst):
    wb = openpyxl.load_workbook(src)
    wbv = openpyxl.load_workbook(src, data_only=True)
    log, subs, n_apu, end_apu, end_part = eliminar_amarillas(wb, wbv)
    parametros(wb)
    metas_marron(wb)
    cu_comp(wb, end_part, log)
    calculo(wb)
    oym(wb)
    textos(wb, n_apu, log, subs)
    wb.save(dst)
    print(f"Partidas eliminadas: {[x['clave'] for x in log]}")
    print(f"Subpartidas eliminadas: {subs}")
    print(f"Líneas APU eliminadas: {n_apu}; APU_BASE termina en {end_apu}; PARTIDAS en {end_part}")
    print(f"Guardado: {dst}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
