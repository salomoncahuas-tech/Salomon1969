"""
IN Piura - Plan de Verificación de Campo
Módulo de generación de reportes PDF y Excel.
Cuenca alta del río Piura, Perú.
"""

import os
from datetime import datetime

from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

import database as db

REPORTES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes")


def _asegurar_directorio():
    os.makedirs(REPORTES_DIR, exist_ok=True)


# ── Reporte PDF - Ficha de Inspección por Bloque ──────────────────────────

class FichaInspeccionPDF(FPDF):
    """PDF personalizado para fichas de inspección IN Piura."""

    def header(self):
        self.set_font("Helvetica", "B", 14)
        self.cell(0, 8, "IN Piura - Plan de Verificacion de Campo", 0, 1, "C")
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
    pdf._seccion("1. Datos del Bloque de Intervención")
    pdf._campo("Código de bloque", bloque["codigo"])
    pdf._campo("Tipo de intervención", bloque["tipo_intervencion"])
    pdf._campo("Cuenca hidrográfica", bloque["cuenca"])
    pdf._campo("Distrito", bloque["distrito"])
    pdf._campo("Coordenadas UTM Este", f"{bloque['utm_este']:.2f} m")
    pdf._campo("Coordenadas UTM Norte", f"{bloque['utm_norte']:.2f} m")
    pdf._campo("Zona UTM", bloque["utm_zona"])
    pdf._campo("Área", f"{bloque['area_hectareas']:.4f} ha")
    pdf._campo("Estado", bloque["estado"])
    pdf.ln(3)

    if not inspecciones:
        pdf._seccion("2. Inspecciones de Campo")
        pdf.set_font("Helvetica", "I", 9)
        pdf.cell(0, 6, "No se han registrado inspecciones para este bloque.", ln=True)
    else:
        for idx, insp in enumerate(inspecciones, 1):
            pdf._seccion(f"2.{idx} Inspección - {insp['fecha_visita']}")
            pdf._campo("Fecha de visita", insp["fecha_visita"])
            pdf._campo("Inspector", insp["inspector"])
            pdf._campo("Condiciones climáticas", insp["condiciones_climaticas"])
            pdf._campo("Avance físico", f"{insp['avance_fisico']:.1f} %")
            pdf._campo("Código de verificación", insp["codigo_verificacion"])
            pdf._campo_largo("Observaciones técnicas", insp["observaciones"])
            pdf._campo_largo("Desviaciones al exp. técnico", insp["desviaciones"])

            if insp["registro_fotografico"]:
                pdf.set_font("Helvetica", "B", 9)
                pdf.cell(0, 6, "Registro fotográfico:", ln=True)
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
                pdf._seccion(f"   Indicadores de Calidad - Inspección {insp['fecha_visita']}")
                pdf._campo("Densidad planificada", f"{indicadores['densidad_planificada']:.0f} pl/ha")
                pdf._campo("Densidad lograda", f"{indicadores['densidad_lograda']:.0f} pl/ha")
                cumplimiento = 0
                if indicadores["densidad_planificada"] > 0:
                    cumplimiento = (indicadores["densidad_lograda"] / indicadores["densidad_planificada"]) * 100
                pdf._campo("Cumplimiento densidad", f"{cumplimiento:.1f} %")
                pdf._campo("Sobrevivencia de especies", f"{indicadores['sobrevivencia_especies']:.1f} %")
                pdf._campo("Longitud zanjas ejecutada", f"{indicadores['longitud_zanjas_ejecutada']:.2f} ml")
                pdf._campo("Vol. retención sedimentos", f"{indicadores['volumen_retencion_sedimentos']:.2f} m³")
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
    ws = wb.active
    ws.title = "Resumen Bloques IN Piura"

    # Estilos
    encabezado_font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    encabezado_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    borde = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )
    centrado = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Título
    ws.merge_cells("A1:R1")
    ws["A1"] = "IN Piura - Resumen de Bloques de Intervención - Cuenca Alta del Río Piura"
    ws["A1"].font = Font(name="Calibri", bold=True, size=13)
    ws["A1"].alignment = Alignment(horizontal="center")

    ws.merge_cells("A2:R2")
    ws["A2"] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = Font(name="Calibri", italic=True, size=9)
    ws["A2"].alignment = Alignment(horizontal="center")

    # Encabezados (fila 4)
    encabezados = [
        "Código Bloque", "Tipo Intervención", "Cuenca Hidrográfica",
        "Distrito", "UTM_Este", "UTM_Norte", "Zona_UTM",
        "Área (ha)", "Estado", "Total Inspecciones",
        "Última Visita", "Avance Físico (%)",
        "Densidad Plan. (pl/ha)", "Densidad Logr. (pl/ha)",
        "Sobrevivencia (%)", "Long. Zanjas (ml)",
        "Vol. Ret. Sedimentos (m³)", "Fecha Registro"
    ]

    for col_idx, enc in enumerate(encabezados, 1):
        celda = ws.cell(row=4, column=col_idx, value=enc)
        celda.font = encabezado_font
        celda.fill = encabezado_fill
        celda.border = borde
        celda.alignment = centrado

    # Datos
    resumen = db.obtener_resumen_bloques()
    for row_idx, bloque in enumerate(resumen, 5):
        indicadores = db.obtener_indicadores_por_bloque(bloque["id"])
        ultimo_ind = indicadores[0] if indicadores else {}

        datos = [
            bloque["codigo"],
            bloque["tipo_intervencion"],
            bloque["cuenca"],
            bloque["distrito"],
            bloque["utm_este"],
            bloque["utm_norte"],
            bloque["utm_zona"],
            bloque["area_hectareas"],
            bloque["estado"],
            bloque.get("total_inspecciones", 0),
            bloque.get("ultima_visita", ""),
            bloque.get("ultimo_avance", 0),
            ultimo_ind.get("densidad_planificada", 0),
            ultimo_ind.get("densidad_lograda", 0),
            ultimo_ind.get("sobrevivencia_especies", 0),
            ultimo_ind.get("longitud_zanjas_ejecutada", 0),
            ultimo_ind.get("volumen_retencion_sedimentos", 0),
            bloque["fecha_registro"],
        ]

        for col_idx, valor in enumerate(datos, 1):
            celda = ws.cell(row=row_idx, column=col_idx, value=valor)
            celda.border = borde
            celda.alignment = Alignment(horizontal="center", vertical="center")

    # Ajustar anchos de columna
    anchos = [14, 24, 22, 16, 14, 14, 10, 10, 14, 14, 14, 14, 16, 16, 14, 16, 18, 18]
    for i, ancho in enumerate(anchos, 1):
        ws.column_dimensions[chr(64 + i) if i <= 26 else ""].width = ancho

    # Ajustar anchos usando openpyxl column letters
    from openpyxl.utils import get_column_letter
    for i, ancho in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = ancho

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
        "OBJECTID", "COD_BLOQUE", "TIPO_INTERV", "CUENCA",
        "DISTRITO", "POINT_X", "POINT_Y", "ZONA_UTM",
        "AREA_HA", "ESTADO", "AVANCE_PCT", "SOBREV_PCT",
        "DENS_PLAN", "DENS_LOGR", "LONG_ZANJAS_ML", "VOL_RETEN_M3"
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
            bloque["tipo_intervencion"],
            bloque["cuenca"],
            bloque["distrito"],
            bloque["utm_este"],
            bloque["utm_norte"],
            bloque["utm_zona"],
            bloque["area_hectareas"],
            bloque["estado"],
            bloque.get("ultimo_avance", 0) or 0,
            ultimo_ind.get("sobrevivencia_especies", 0),
            ultimo_ind.get("densidad_planificada", 0),
            ultimo_ind.get("densidad_lograda", 0),
            ultimo_ind.get("longitud_zanjas_ejecutada", 0),
            ultimo_ind.get("volumen_retencion_sedimentos", 0),
        ]

        for col_idx, valor in enumerate(datos, 1):
            ws.cell(row=row_idx, column=col_idx, value=valor)

    from openpyxl.utils import get_column_letter
    anchos = [10, 14, 24, 20, 16, 14, 14, 10, 10, 14, 12, 12, 12, 12, 16, 14]
    for i, ancho in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = ancho

    nombre_archivo = f"bloques_arcgis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    ruta_excel = os.path.join(REPORTES_DIR, nombre_archivo)
    wb.save(ruta_excel)
    return ruta_excel
