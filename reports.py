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
    provincia = bloque.get("provincia", "") or ""
    if provincia:
        pdf._campo("Provincia", provincia)
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
            pdf._campo_largo("Desviaciones observadas al Plan de Trabajo", insp["desviaciones"])

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
                pdf.ln(3)

    # Diagnostico Territorial del bloque
    diagnosticos_dt = db.obtener_diagnosticos_por_bloque(bloque_id)
    if diagnosticos_dt:
        pdf._seccion("3. Diagnostico Territorial")
        for idx_dt, dt in enumerate(diagnosticos_dt, 1):
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(0, 6, f"  Ficha {dt.get('ficha', '')} - {dt.get('fecha_evaluacion', '')}", 0, 1)
            pdf.set_font("Helvetica", "", 9)
            if dt.get("microcuenca"):
                pdf._campo("Microcuenca", dt["microcuenca"])
            if dt.get("evaluador"):
                pdf._campo("Evaluador", dt["evaluador"])
            # F-DT-01: Caracteristicas Fisiograficas
            campos_dt01 = [
                ("Forma del terreno", "forma_terreno"), ("Pendiente", "pendiente"),
                ("Posicion fisiografica", "posicion_fisiografica"), ("Exposicion", "exposicion_orientacion"),
                ("Paisaje dominante", "paisaje_dominante"), ("Rango altitudinal", "rango_altitudinal")]
            tiene_dt01 = any(dt.get(c[1]) for c in campos_dt01)
            if tiene_dt01:
                pdf.set_font("Helvetica", "BI", 9)
                pdf.cell(0, 5, "  F-DT-01: Caracteristicas Fisiograficas", 0, 1)
                pdf.set_font("Helvetica", "", 9)
                for label, key in campos_dt01:
                    val = dt.get(key, "") or ""
                    if val:
                        pdf._campo(label, val)
            # F-DT-02: Condiciones Climaticas
            campos_dt02 = [
                ("Precipitacion anual", "precipitacion_anual"), ("Temperatura media", "temperatura_media"),
                ("Humedad relativa", "humedad_relativa"), ("Zona de vida", "zona_vida"),
                ("Presencia de heladas", "presencia_heladas"), ("Regimen de vientos", "regimen_vientos")]
            tiene_dt02 = any(dt.get(c[1]) for c in campos_dt02)
            if tiene_dt02:
                pdf.set_font("Helvetica", "BI", 9)
                pdf.cell(0, 5, "  F-DT-02: Condiciones Climaticas", 0, 1)
                pdf.set_font("Helvetica", "", 9)
                for label, key in campos_dt02:
                    val = dt.get(key, "") or ""
                    if val:
                        pdf._campo(label, val)
            # F-DT-03: Caracteristicas del Suelo
            campos_dt03 = [
                ("Textura", "textura_suelo"), ("Color del suelo", "color_suelo"),
                ("Profundidad efectiva", "profundidad_efectiva"), ("Pedregosidad", "pedregosidad"),
                ("Drenaje", "drenaje"), ("Presencia de erosion", "presencia_erosion"),
                ("Materia organica", "materia_organica")]
            tiene_dt03 = any(dt.get(c[1]) for c in campos_dt03)
            if tiene_dt03:
                pdf.set_font("Helvetica", "BI", 9)
                pdf.cell(0, 5, "  F-DT-03: Caracteristicas del Suelo", 0, 1)
                pdf.set_font("Helvetica", "", 9)
                for label, key in campos_dt03:
                    val = dt.get(key, "") or ""
                    if val:
                        pdf._campo(label, val)
            # F-DT-04: Cobertura Vegetal y Uso del Suelo
            campos_dt04 = [
                ("Tipo de cobertura", "tipo_cobertura"), ("Densidad de cobertura", "densidad_cobertura"),
                ("Estado de conservacion", "estado_conservacion"), ("Uso actual del suelo", "uso_actual_suelo"),
                ("Estado de uso del suelo", "conflicto_uso")]
            tiene_dt04 = any(dt.get(c[1]) for c in campos_dt04)
            if tiene_dt04:
                pdf.set_font("Helvetica", "BI", 9)
                pdf.cell(0, 5, "  F-DT-04: Cobertura Vegetal y Uso del Suelo", 0, 1)
                pdf.set_font("Helvetica", "", 9)
                for label, key in campos_dt04:
                    val = dt.get(key, "") or ""
                    if val:
                        pdf._campo(label, val)
            # F-DT-05: Recursos Hidricos
            campos_dt05 = [
                ("Fuente de agua", "fuente_agua"), ("Regimen hidrico", "regimen_hidrico"),
                ("Calidad del agua", "calidad_agua"), ("Distancia a fuente", "distancia_fuente_agua"),
                ("Uso recurso hidrico", "uso_recurso_hidrico")]
            tiene_dt05 = any(dt.get(c[1]) for c in campos_dt05)
            if tiene_dt05:
                pdf.set_font("Helvetica", "BI", 9)
                pdf.cell(0, 5, "  F-DT-05: Recursos Hidricos", 0, 1)
                pdf.set_font("Helvetica", "", 9)
                for label, key in campos_dt05:
                    val = dt.get(key, "") or ""
                    if val:
                        pdf._campo(label, val)
            # F-DT-06: Aspectos Socioeconomicos
            campos_dt06 = [
                ("Tenencia de la tierra", "tenencia_tierra"), ("Organizacion comunal", "organizacion_comunal"),
                ("Actividad economica", "actividad_economica"), ("Accesibilidad vial", "accesibilidad_via"),
                ("Distancia centro poblado", "distancia_centro_poblado"), ("Servicios basicos", "servicios_basicos")]
            tiene_dt06 = any(dt.get(c[1]) for c in campos_dt06)
            if tiene_dt06:
                pdf.set_font("Helvetica", "BI", 9)
                pdf.cell(0, 5, "  F-DT-06: Aspectos Socioeconomicos", 0, 1)
                pdf.set_font("Helvetica", "", 9)
                for label, key in campos_dt06:
                    val = dt.get(key, "") or ""
                    if val:
                        pdf._campo(label, val)
            # Observaciones generales
            obs_dt = dt.get("observaciones_generales", "") or ""
            if obs_dt:
                pdf._campo_largo("Observaciones generales", obs_dt)
            pdf.ln(2)

    # Diagnostico Social del bloque
    diagnosticos_ds = db.obtener_diagnosticos_sociales_por_bloque(bloque_id)
    if diagnosticos_ds:
        pdf._seccion("4. Diagnostico Social")
        for idx_ds, ds in enumerate(diagnosticos_ds, 1):
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(0, 6, f"  Ficha {ds.get('ficha', '')} - {ds.get('fecha_evaluacion', '')}", 0, 1)
            pdf.set_font("Helvetica", "", 9)
            if ds.get("microcuenca"):
                pdf._campo("Microcuenca", ds["microcuenca"])
            if ds.get("evaluador"):
                pdf._campo("Evaluador", ds["evaluador"])
            if ds.get("centro_poblado"):
                pdf._campo("Centro poblado", ds["centro_poblado"])
            if ds.get("comunidad_campesina"):
                pdf._campo("Comunidad campesina", ds["comunidad_campesina"])
            if ds.get("provincia"):
                pdf._campo("Provincia", ds["provincia"])
            if ds.get("distrito"):
                pdf._campo("Distrito", ds["distrito"])
            # F-DS-01: Datos socioeconomicos
            campos_ds01 = [
                ("N de familias", "ds01_num_familias"),
                ("Poblacion hombres", "ds01_poblacion_hombres"),
                ("Poblacion mujeres", "ds01_poblacion_mujeres"),
                ("Poblacion total", "ds01_poblacion_total"),
                ("Organizacion comunal", "ds01_organizacion_comunal"),
                ("Junta directiva", "ds01_junta_directiva"),
                ("Presidente junta", "ds01_presidente_junta"),
                ("Agua potable tipo", "ds01_agua_potable_tipo"),
                ("Agua potable cobertura", "ds01_agua_potable_cobertura"),
                ("Saneamiento", "ds01_saneamiento"),
                ("Energia tipo", "ds01_energia_tipo"),
                ("Energia cobertura", "ds01_energia_cobertura"),
                ("Actividades economicas", "ds01_actividades_economicas"),
                ("Fuente de agua", "ds01_fuente_agua"),
                ("Problemas de agua", "ds01_problemas_agua"),
                ("Percepcion de cambios", "ds01_percepcion_cambios"),
                ("Disposicion a participar", "ds01_disposicion_participar")]
            tiene_ds01 = any(ds.get(c[1]) for c in campos_ds01)
            if tiene_ds01:
                pdf.set_font("Helvetica", "BI", 9)
                pdf.cell(0, 5, "  F-DS-01: Diagnostico Socioeconomico", 0, 1)
                pdf.set_font("Helvetica", "", 9)
                for label, key in campos_ds01:
                    val = ds.get(key, "") or ""
                    if val:
                        pdf._campo(label, val)
            # F-DS-02: Actores clave
            campos_ds02 = [
                ("Registro de actores", "ds02_registro_actores"),
                ("Actores gobierno local", "ds02_actores_gob_local"),
                ("Actores comunidades", "ds02_actores_comunidades"),
                ("Actores juntas de riego", "ds02_actores_juntas_riego"),
                ("Actores ONG", "ds02_actores_ong")]
            tiene_ds02 = any(ds.get(c[1]) for c in campos_ds02)
            if tiene_ds02:
                pdf.set_font("Helvetica", "BI", 9)
                pdf.cell(0, 5, "  F-DS-02: Actores Clave", 0, 1)
                pdf.set_font("Helvetica", "", 9)
                for label, key in campos_ds02:
                    val = ds.get(key, "") or ""
                    if val:
                        pdf._campo(label, val)
            # F-DS-03: Entrevista
            campos_ds03 = [
                ("Nombre entrevistado", "ds03_nombre_entrevistado"),
                ("Cargo / Funcion", "ds03_cargo_funcion"),
                ("Institucion", "ds03_institucion")]
            tiene_ds03 = any(ds.get(c[1]) for c in campos_ds03)
            if tiene_ds03:
                pdf.set_font("Helvetica", "BI", 9)
                pdf.cell(0, 5, "  F-DS-03: Entrevista Semiestructurada", 0, 1)
                pdf.set_font("Helvetica", "", 9)
                for label, key in campos_ds03:
                    val = ds.get(key, "") or ""
                    if val:
                        pdf._campo(label, val)
                # Respuestas clave
                resp_keys = [
                    ("Recursos naturales", "ds03_resp_recursos_naturales"),
                    ("Cambios ambientales", "ds03_resp_cambios_ambiente"),
                    ("Problemas ambientales", "ds03_resp_problemas_ambientales"),
                    ("Actividades economicas", "ds03_resp_actividades_economicas"),
                    ("Expectativas del proyecto", "ds03_resp_expectativas"),
                    ("Disposicion a participar", "ds03_resp_disposicion_participar")]
                for label, key in resp_keys:
                    val = ds.get(key, "") or ""
                    if val:
                        pdf._campo(label, val)
            # F-DS-04: Taller participativo
            campos_ds04 = [
                ("Lugar del taller", "ds04_lugar_taller"),
                ("Objetivo", "ds04_objetivo"),
                ("Acuerdos", "ds04_acuerdos")]
            tiene_ds04 = any(ds.get(c[1]) for c in campos_ds04)
            if tiene_ds04:
                pdf.set_font("Helvetica", "BI", 9)
                pdf.cell(0, 5, "  F-DS-04: Taller Participativo", 0, 1)
                pdf.set_font("Helvetica", "", 9)
                for label, key in campos_ds04:
                    val = ds.get(key, "") or ""
                    if val:
                        pdf._campo(label, val)
            # F-DS-05: Conflictos y oportunidades
            campos_ds05 = [
                ("Conflictos identificados", "ds05_conflictos"),
                ("Oportunidades identificadas", "ds05_oportunidades")]
            tiene_ds05 = any(ds.get(c[1]) for c in campos_ds05)
            if tiene_ds05:
                pdf.set_font("Helvetica", "BI", 9)
                pdf.cell(0, 5, "  F-DS-05: Conflictos y Oportunidades", 0, 1)
                pdf.set_font("Helvetica", "", 9)
                for label, key in campos_ds05:
                    val = ds.get(key, "") or ""
                    if val:
                        pdf._campo(label, val)
            pdf.ln(2)

    # Cronograma del bloque
    actividades = db.obtener_actividades_por_bloque(bloque_id)
    if actividades:
        pdf._seccion("5. Cronograma de Actividades")
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
    ws.merge_cells("A1:T1")
    ws["A1"] = "IN Piura - Plan de Ingreso - Resumen de Bloques de Intervencion"
    ws["A1"].font = Font(name="Calibri", bold=True, size=13)
    ws["A1"].alignment = Alignment(horizontal="center")

    ws.merge_cells("A2:T2")
    ws["A2"] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = Font(name="Calibri", italic=True, size=9)
    ws["A2"].alignment = Alignment(horizontal="center")

    # Encabezados (fila 4)
    encabezados = [
        "Codigo Bloque", "Microcuenca", "Tipo Intervencion", "Cuenca Hidrografica",
        "Provincia", "Distrito", "UTM_Este", "UTM_Norte", "Zona_UTM", "Altitud (m.s.n.m.)",
        "Area (ha)", "Responsable", "Estado", "Total Inspecciones",
        "Ultima Visita", "Avance Fisico (%)",
        "Cobertura Vegetal (%)", "Tipo Cobertura", "Vigor Cobertura",
        "Fecha Registro"
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
            bloque.get("provincia", "") or "",
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
            bloque["fecha_registro"],
        ]

        for col_idx, valor in enumerate(datos, 1):
            celda = ws.cell(row=row_idx, column=col_idx, value=valor)
            celda.border = borde
            celda.alignment = Alignment(horizontal="center", vertical="center")

    anchos = [14, 18, 24, 22, 14, 16, 14, 14, 10, 14, 10, 16, 14, 14, 14, 14, 16, 16, 16, 18]
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
        "PROVINCIA", "DISTRITO", "POINT_X", "POINT_Y", "ZONA_UTM", "ALTITUD",
        "AREA_HA", "RESPONSABLE", "ESTADO", "AVANCE_PCT",
        "COB_VEG_PCT", "TIPO_COB_VEG", "VIGOR_COB_VEG"
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
            bloque.get("provincia", "") or "",
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
        ]

        for col_idx, valor in enumerate(datos, 1):
            ws.cell(row=row_idx, column=col_idx, value=valor)

    anchos = [10, 14, 18, 24, 20, 14, 16, 14, 14, 10, 10, 10, 16, 14, 12, 14, 16, 16]
    for i, ancho in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = ancho

    nombre_archivo = f"bloques_arcgis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    ruta_excel = os.path.join(REPORTES_DIR, nombre_archivo)
    wb.save(ruta_excel)
    return ruta_excel
