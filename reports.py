"""
IN Piura - Plan de Ingreso / Verificación de Campo
Módulo de generación de reportes PDF y Excel.
Incluye fichas de inspección, resumen de bloques, presupuesto y cronograma.
Cuenca alta del río Piura, Perú.
"""

import os
from datetime import datetime

from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

import database as db

REPORTES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes")


def _asegurar_directorio():
    os.makedirs(REPORTES_DIR, exist_ok=True)


# ── Estilos comunes Excel ────────────────────────────────────────────────

def _estilos_excel():
    encabezado_font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    encabezado_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    borde = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )
    centrado = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return encabezado_font, encabezado_fill, borde, centrado


# ── Reporte PDF - Ficha de Inspección por Bloque ──────────────────────────

class FichaInspeccionPDF(FPDF):
    """PDF personalizado para fichas de inspección IN Piura."""

    def header(self):
        self.set_font("Helvetica", "B", 14)
        self.cell(0, 8, "IN Piura - Plan de Ingreso - Verificacion de Campo", 0, 1, "C")
        self.set_font("Helvetica", "", 9)
        self.cell(0, 5, "Restauracion de Ecosistemas - Cuenca Alta del Rio Piura", 0, 1, "C")
        self.line(10, self.get_y() + 2, 200, self.get_y() + 2)
        self.ln(6)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.cell(0, 10,
                  f"Pagina {self.page_no()}/{{nb}}  |  Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                  0, 0, "C")

    def _seccion(self, titulo):
        self.set_font("Helvetica", "B", 11)
        self.set_fill_color(44, 62, 80)
        self.set_text_color(255, 255, 255)
        self.cell(0, 7, f"  {titulo}", 0, 1, fill=True)
        self.set_text_color(0, 0, 0)
        self.ln(2)

    def _campo(self, etiqueta, valor, ancho_etiqueta=55):
        x = self.get_x()
        y = self.get_y()
        self.set_font("Helvetica", "B", 9)
        self.set_xy(x, y)
        self.cell(ancho_etiqueta, 6, etiqueta + ":")
        self.set_font("Helvetica", "", 9)
        texto = str(valor) if valor else "-"
        self.cell(0, 6, texto, 0, 1)

    def _campo_largo(self, etiqueta, valor):
        self.set_font("Helvetica", "B", 9)
        self.cell(0, 6, etiqueta + ":", 0, 1)
        self.set_font("Helvetica", "", 9)
        texto = str(valor) if valor else "Sin observaciones."
        x_margin = self.l_margin
        self.set_x(x_margin)
        self.multi_cell(190, 5, texto)
        self.ln(1)


def generar_ficha_pdf(bloque_id, inspeccion_id=None):
    """Genera un PDF con la ficha de inspección de un bloque."""
    _asegurar_directorio()

    bloque = db.obtener_bloque_por_id(bloque_id)
    if not bloque:
        raise ValueError(f"Bloque con id {bloque_id} no encontrado.")

    inspecciones = db.obtener_inspecciones_por_bloque(bloque_id)
    if inspeccion_id:
        inspecciones = [i for i in inspecciones if i["id"] == inspeccion_id]

    pdf = FichaInspeccionPDF()
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()

    # Datos del bloque
    pdf._seccion("1. Datos del Bloque de Intervencion")
    pdf._campo("Codigo de bloque", bloque["codigo"])
    pdf._campo("Tipo de intervencion", bloque["tipo_intervencion"])
    pdf._campo("Cuenca hidrografica", bloque["cuenca"])
    microcuenca = bloque.get("microcuenca", "") or ""
    if microcuenca:
        pdf._campo("Codigo microcuenca", microcuenca)
    pdf._campo("Distrito", bloque["distrito"])
    pdf._campo("Coordenadas UTM Este", f"{bloque['utm_este']:.2f} m")
    pdf._campo("Coordenadas UTM Norte", f"{bloque['utm_norte']:.2f} m")
    pdf._campo("Zona UTM", bloque["utm_zona"])
    altitud = bloque.get("altitud", 0) or 0
    pdf._campo("Altitud", f"{altitud:.0f} m.s.n.m.")
    pdf._campo("Area", f"{bloque['area_hectareas']:.4f} ha")
    responsable = bloque.get("responsable", "") or ""
    if responsable:
        pdf._campo("Responsable", responsable)
    pdf._campo("Estado", bloque["estado"])
    pdf.ln(3)

    # Presupuesto del bloque
    presupuesto = db.obtener_presupuesto_por_bloque(bloque_id)
    if presupuesto:
        pdf._seccion("1.1 Resumen Presupuestal del Bloque")
        total_plan = sum(p["monto_planificado"] for p in presupuesto)
        total_ejec = sum(p["monto_ejecutado"] for p in presupuesto)
        pct = (total_ejec / total_plan * 100) if total_plan > 0 else 0

        for p in presupuesto:
            pct_p = (p["monto_ejecutado"] / p["monto_planificado"] * 100) if p["monto_planificado"] > 0 else 0
            pdf._campo(p["categoria"],
                       f"Plan: S/ {p['monto_planificado']:,.2f} | "
                       f"Ejec: S/ {p['monto_ejecutado']:,.2f} ({pct_p:.1f}%)")

        pdf.ln(1)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(0, 6,
                 f"TOTAL: Planificado S/ {total_plan:,.2f} | "
                 f"Ejecutado S/ {total_ejec:,.2f} ({pct:.1f}%)", 0, 1)
        pdf.ln(3)

    if not inspecciones:
        pdf._seccion("2. Inspecciones de Campo")
        pdf.set_font("Helvetica", "I", 9)
        pdf.cell(0, 6, "No se han registrado inspecciones para este bloque.", ln=True)
    else:
        for idx, insp in enumerate(inspecciones, 1):
            pdf._seccion(f"2.{idx} Inspeccion - {insp['fecha_visita']}")
            pdf._campo("Fecha de visita", insp["fecha_visita"])
            pdf._campo("Inspector", insp["inspector"])
            pdf._campo("Condiciones climaticas", insp["condiciones_climaticas"])
            pdf._campo("Avance fisico", f"{insp['avance_fisico']:.1f} %")
            pdf._campo("Codigo de verificacion", insp["codigo_verificacion"])
            pdf._campo_largo("Observaciones tecnicas", insp["observaciones"])
            pdf._campo_largo("Desviaciones al exp. tecnico", insp["desviaciones"])

            if insp["registro_fotografico"]:
                pdf.set_font("Helvetica", "B", 9)
                pdf.cell(0, 6, "Registro fotografico:", ln=True)
                pdf.set_font("Helvetica", "", 8)
                for ruta in insp["registro_fotografico"].split(";"):
                    ruta = ruta.strip()
                    if ruta:
                        pdf.cell(10)
                        pdf.cell(0, 5, f"- {ruta}", ln=True)
            pdf.ln(2)

            # Indicadores de calidad de esta inspección
            indicadores = db.obtener_indicadores_por_inspeccion(insp["id"])
            if indicadores:
                pdf._seccion(f"   Indicadores de Calidad - Inspeccion {insp['fecha_visita']}")
                mc_ind = indicadores.get("microcuenca", "") or ""
                if mc_ind:
                    pdf._campo("Codigo microcuenca", mc_ind)
                pct_cob = indicadores.get("porcentaje_cobertura_vegetal", 0) or 0
                pdf._campo("Cobertura vegetal", f"{pct_cob:.1f} %")
                tipo_cob = indicadores.get("tipo_cobertura_vegetal", "") or ""
                if tipo_cob:
                    pdf._campo("Tipo cobertura vegetal", tipo_cob)
                vigor_cob = indicadores.get("vigor_cobertura_vegetal", "") or ""
                if vigor_cob:
                    pdf._campo("Vigor cobertura vegetal", vigor_cob)
                pdf._campo("Sobrevivencia de especies", f"{indicadores['sobrevivencia_especies']:.1f} %")
                pdf._campo("Longitud zanjas ejecutada", f"{indicadores['longitud_zanjas_ejecutada']:.2f} ml")
                pdf._campo("Vol. retencion sedimentos", f"{indicadores['volumen_retencion_sedimentos']:.2f} m3")
                pdf.ln(3)

    # Cronograma del bloque
    actividades = db.obtener_actividades_por_bloque(bloque_id)
    if actividades:
        pdf._seccion("3. Cronograma de Actividades")
        for a in actividades:
            estado_act = a.get("estado", "Programado")
            avance_act = a.get("porcentaje_avance", 0)
            pdf._campo(a["actividad"],
                       f"{a['fecha_inicio_plan']} a {a['fecha_fin_plan']} | "
                       f"Estado: {estado_act} | Avance: {avance_act:.0f}%")
        pdf.ln(3)

    nombre_archivo = f"ficha_{bloque['codigo']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    ruta_pdf = os.path.join(REPORTES_DIR, nombre_archivo)
    pdf.output(ruta_pdf)
    return ruta_pdf


# ── Reporte Excel - Tabla Resumen ──────────────────────────────────────────

def generar_resumen_excel():
    """Genera un archivo Excel con la tabla resumen de todos los bloques inspeccionados.
    Incluye columnas UTM compatibles con importación a ArcGIS."""
    _asegurar_directorio()

    wb = Workbook()
    enc_font, enc_fill, borde, centrado = _estilos_excel()

    # ── Hoja 1: Resumen de Bloques ──
    ws = wb.active
    ws.title = "Resumen Bloques IN Piura"

    # Título
    ws.merge_cells("A1:V1")
    ws["A1"] = "IN Piura - Plan de Ingreso - Resumen de Bloques de Intervencion"
    ws["A1"].font = Font(name="Calibri", bold=True, size=13)
    ws["A1"].alignment = Alignment(horizontal="center")

    ws.merge_cells("A2:V2")
    ws["A2"] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = Font(name="Calibri", italic=True, size=9)
    ws["A2"].alignment = Alignment(horizontal="center")

    # Encabezados (fila 4)
    encabezados = [
        "Codigo Bloque", "Microcuenca", "Tipo Intervencion", "Cuenca Hidrografica",
        "Distrito", "UTM_Este", "UTM_Norte", "Zona_UTM", "Altitud (m.s.n.m.)",
        "Area (ha)", "Responsable", "Estado", "Total Inspecciones",
        "Ultima Visita", "Avance Fisico (%)",
        "Cobertura Vegetal (%)", "Tipo Cobertura", "Vigor Cobertura",
        "Sobrevivencia (%)", "Long. Zanjas (ml)",
        "Vol. Ret. Sedimentos (m3)", "Fecha Registro"
    ]

    for col_idx, enc in enumerate(encabezados, 1):
        celda = ws.cell(row=4, column=col_idx, value=enc)
        celda.font = enc_font
        celda.fill = enc_fill
        celda.border = borde
        celda.alignment = centrado

    # Datos
    resumen = db.obtener_resumen_bloques()
    for row_idx, bloque in enumerate(resumen, 5):
        indicadores = db.obtener_indicadores_por_bloque(bloque["id"])
        ultimo_ind = indicadores[0] if indicadores else {}

        datos = [
            bloque["codigo"],
            bloque.get("microcuenca", "") or "",
            bloque["tipo_intervencion"],
            bloque["cuenca"],
            bloque["distrito"],
            bloque["utm_este"],
            bloque["utm_norte"],
            bloque["utm_zona"],
            bloque.get("altitud", 0) or 0,
            bloque["area_hectareas"],
            bloque.get("responsable", "") or "",
            bloque["estado"],
            bloque.get("total_inspecciones", 0),
            bloque.get("ultima_visita", ""),
            bloque.get("ultimo_avance", 0),
            ultimo_ind.get("porcentaje_cobertura_vegetal", 0),
            ultimo_ind.get("tipo_cobertura_vegetal", "") or "",
            ultimo_ind.get("vigor_cobertura_vegetal", "") or "",
            ultimo_ind.get("sobrevivencia_especies", 0),
            ultimo_ind.get("longitud_zanjas_ejecutada", 0),
            ultimo_ind.get("volumen_retencion_sedimentos", 0),
            bloque["fecha_registro"],
        ]

        for col_idx, valor in enumerate(datos, 1):
            celda = ws.cell(row=row_idx, column=col_idx, value=valor)
            celda.border = borde
            celda.alignment = Alignment(horizontal="center", vertical="center")

    anchos = [14, 18, 24, 22, 16, 14, 14, 10, 14, 10, 16, 14, 14, 14, 14, 16, 16, 16, 14, 16, 18, 18]
    for i, ancho in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = ancho

    # ── Hoja 2: Resumen Presupuestal ──
    ws_pres = wb.create_sheet("Resumen Presupuestal")

    ws_pres.merge_cells("A1:G1")
    ws_pres["A1"] = "IN Piura - Resumen Presupuestal por Bloque"
    ws_pres["A1"].font = Font(name="Calibri", bold=True, size=13)
    ws_pres["A1"].alignment = Alignment(horizontal="center")

    enc_pres = ["Codigo Bloque", "Tipo Intervencion", "Distrito",
                "Planificado (S/)", "Ejecutado (S/)", "% Ejecucion", "N Partidas"]
    for col_idx, enc in enumerate(enc_pres, 1):
        celda = ws_pres.cell(row=3, column=col_idx, value=enc)
        celda.font = enc_font
        celda.fill = enc_fill
        celda.border = borde
        celda.alignment = centrado

    resumen_pres = db.obtener_resumen_presupuesto()
    total_plan_global = 0
    total_ejec_global = 0
    for row_idx, r in enumerate(resumen_pres, 4):
        pct = (r["total_ejecutado"] / r["total_planificado"] * 100) if r["total_planificado"] > 0 else 0
        total_plan_global += r["total_planificado"]
        total_ejec_global += r["total_ejecutado"]

        datos_p = [
            r["codigo"], r["tipo_intervencion"], r["distrito"],
            r["total_planificado"], r["total_ejecutado"],
            round(pct, 1), r["num_partidas"],
        ]
        for col_idx, valor in enumerate(datos_p, 1):
            celda = ws_pres.cell(row=row_idx, column=col_idx, value=valor)
            celda.border = borde
            celda.alignment = Alignment(horizontal="center", vertical="center")

    # Fila de totales
    fila_total = len(resumen_pres) + 4
    ws_pres.cell(row=fila_total, column=1, value="TOTAL PROYECTO").font = Font(bold=True)
    ws_pres.cell(row=fila_total, column=4, value=total_plan_global).font = Font(bold=True)
    ws_pres.cell(row=fila_total, column=5, value=total_ejec_global).font = Font(bold=True)
    pct_global = (total_ejec_global / total_plan_global * 100) if total_plan_global > 0 else 0
    ws_pres.cell(row=fila_total, column=6, value=round(pct_global, 1)).font = Font(bold=True)

    anchos_pres = [14, 24, 16, 16, 16, 12, 12]
    for i, ancho in enumerate(anchos_pres, 1):
        ws_pres.column_dimensions[get_column_letter(i)].width = ancho

    # ── Hoja 3: Cronograma ──
    ws_crono = wb.create_sheet("Cronograma")

    ws_crono.merge_cells("A1:H1")
    ws_crono["A1"] = "IN Piura - Cronograma de Actividades"
    ws_crono["A1"].font = Font(name="Calibri", bold=True, size=13)
    ws_crono["A1"].alignment = Alignment(horizontal="center")

    enc_crono = ["Bloque", "Actividad", "Inicio Plan.", "Fin Plan.",
                 "Inicio Real", "Fin Real", "Avance %", "Estado"]
    for col_idx, enc in enumerate(enc_crono, 1):
        celda = ws_crono.cell(row=3, column=col_idx, value=enc)
        celda.font = enc_font
        celda.fill = enc_fill
        celda.border = borde
        celda.alignment = centrado

    actividades = db.obtener_todas_actividades()
    for row_idx, a in enumerate(actividades, 4):
        datos_a = [
            a.get("bloque_codigo", ""),
            a["actividad"],
            a["fecha_inicio_plan"],
            a["fecha_fin_plan"],
            a["fecha_inicio_real"] or "-",
            a["fecha_fin_real"] or "-",
            a["porcentaje_avance"],
            a["estado"],
        ]
        for col_idx, valor in enumerate(datos_a, 1):
            celda = ws_crono.cell(row=row_idx, column=col_idx, value=valor)
            celda.border = borde
            celda.alignment = Alignment(horizontal="center", vertical="center")

    anchos_crono = [14, 28, 14, 14, 14, 14, 10, 14]
    for i, ancho in enumerate(anchos_crono, 1):
        ws_crono.column_dimensions[get_column_letter(i)].width = ancho

    nombre_archivo = f"resumen_bloques_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    ruta_excel = os.path.join(REPORTES_DIR, nombre_archivo)
    wb.save(ruta_excel)
    return ruta_excel


def generar_excel_arcgis():
    """Genera un Excel simplificado con columnas UTM listas para importar a ArcGIS."""
    _asegurar_directorio()

    wb = Workbook()
    ws = wb.active
    ws.title = "Bloques_UTM_ArcGIS"

    encabezados_arcgis = [
        "OBJECTID", "COD_BLOQUE", "MICROCUENCA", "TIPO_INTERV", "CUENCA",
        "DISTRITO", "POINT_X", "POINT_Y", "ZONA_UTM", "ALTITUD",
        "AREA_HA", "RESPONSABLE", "ESTADO", "AVANCE_PCT",
        "COB_VEG_PCT", "TIPO_COB_VEG", "VIGOR_COB_VEG",
        "SOBREV_PCT", "LONG_ZANJAS_ML", "VOL_RETEN_M3"
    ]

    encabezado_font = Font(name="Calibri", bold=True, size=10)
    for col_idx, enc in enumerate(encabezados_arcgis, 1):
        celda = ws.cell(row=1, column=col_idx, value=enc)
        celda.font = encabezado_font

    resumen = db.obtener_resumen_bloques()
    for row_idx, bloque in enumerate(resumen, 2):
        indicadores = db.obtener_indicadores_por_bloque(bloque["id"])
        ultimo_ind = indicadores[0] if indicadores else {}

        datos = [
            row_idx - 1,
            bloque["codigo"],
            bloque.get("microcuenca", "") or "",
            bloque["tipo_intervencion"],
            bloque["cuenca"],
            bloque["distrito"],
            bloque["utm_este"],
            bloque["utm_norte"],
            bloque["utm_zona"],
            bloque.get("altitud", 0) or 0,
            bloque["area_hectareas"],
            bloque.get("responsable", "") or "",
            bloque["estado"],
            bloque.get("ultimo_avance", 0) or 0,
            ultimo_ind.get("porcentaje_cobertura_vegetal", 0),
            ultimo_ind.get("tipo_cobertura_vegetal", "") or "",
            ultimo_ind.get("vigor_cobertura_vegetal", "") or "",
            ultimo_ind.get("sobrevivencia_especies", 0),
            ultimo_ind.get("longitud_zanjas_ejecutada", 0),
            ultimo_ind.get("volumen_retencion_sedimentos", 0),
        ]

        for col_idx, valor in enumerate(datos, 1):
            ws.cell(row=row_idx, column=col_idx, value=valor)

    anchos = [10, 14, 18, 24, 20, 16, 14, 14, 10, 10, 10, 16, 14, 12, 14, 16, 16, 12, 16, 14]
    for i, ancho in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = ancho

    nombre_archivo = f"bloques_arcgis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    ruta_excel = os.path.join(REPORTES_DIR, nombre_archivo)
    wb.save(ruta_excel)
    return ruta_excel
