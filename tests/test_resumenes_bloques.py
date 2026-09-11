"""Pruebas del lector de resumenes Excel de Diagnostico Territorial.

El fixture reproduce la estructura de los libros
"Plantilla_Excel_Bloque_<codigo>_IN_Piura.xlsx": cinco hojas, encabezados
institucionales, pares etiqueta/valor a dos columnas y tablas de largo
variable. Las pruebas cubren el caso completo, el bloque sin distribucion
MSAVI (el caso mayoritario en los 117 libros) y los libros parciales.
"""

import io
import os
import sys
import unittest
import zipfile

from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import resumenes_bloques as rb  # noqa: E402


# ══════════════════════════════════════════════════════════════════════════
# Fixture: libro con la estructura real del generador
# ══════════════════════════════════════════════════════════════════════════

def _encabezados(ws, titulo, ancho):
    for texto in ("AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN",
                  "DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME",
                  "SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN - SESDI",
                  "PROYECTO IN PIURA | CUI 2669244 | Recuperación del servicio"):
        ws.append([texto] + [None] * (ancho - 1))
    ws.append([titulo] + [None] * (ancho - 1))
    ws.append(["Subtítulo metodológico"] + [None] * (ancho - 1))


def _pares(ws, filas):
    """Escribe pares etiqueta/valor en (A,B) y (C,D), como el generador."""
    for fila in filas:
        ws.append(list(fila) + [None] * (4 - len(fila)))


def construir_libro(codigo="M9B1", con_msavi_areal=False, hojas=None,
                    n_discrepancias=4):
    """Genera un libro de resumen equivalente al de produccion."""
    hojas = hojas or ["Resumen", "Cobertura MSAVI-NDVI", "Estaciones fotográficas",
                      "Microcuenca", "Control de consistencia"]
    wb = Workbook()
    wb.remove(wb.active)

    if "Resumen" in hojas:
        ws = wb.create_sheet("Resumen")
        _encabezados(ws, f"FICHA RESUMEN — BLOQUE PRELIMINAR DE INTERVENCIÓN {codigo}", 4)
        ws.append(["1. IDENTIFICACIÓN Y LOCALIZACIÓN"])
        _pares(ws, [
            ("Código del bloque", codigo, "Microcuenca (catálogo)", "C1096-Q9545"),
            ("Zona de planificación", "Z01",
             "Microcuenca declarada en ficha DT", "C1076-Q9584 (según ficha)"),
            ("Departamento", "Piura", "Provincia", "Morropon"),
            ("Distrito", "Chulucanas", "Capital distrital", "Chulucanas"),
            ("Centro poblado asociado", "La Peña",
             "Comunidad campesina", "María Ángela Alvarado de Zeta"),
            ("Superficie de catálogo (V5/V6), ha", 355.36,
             "Fuente de superficie", "Catálogo maestro Bloques V5/V6"),
            ("Centroide UTM ESTE (m)", 595552, "Centroide UTM NORTE (m)", 9451222),
            ("Sistema de coordenadas", "UTM WGS 84 Zona 17S (EPSG:32717)",
             "Tipo de intervención", "Restauración"),
        ])
        ws.append(["2. PARÁMETROS FÍSICOS (ESTADÍSTICA ZONAL SOBRE MDE)"])
        _pares(ws, [
            ("Altitud mínima (msnm)", 147, "Altitud máxima (msnm)", 265),
            ("Amplitud altitudinal (m)", 118,
             "Piso altitudinal dominante", "<1000 m (Yunga)"),
            ("Pendiente promedio (%)", 10.17, "Pendiente promedio (grados)", 5.8),
            ("Clase de pendiente equivalente", "8-15% (Mod. inclinado)",
             "Rango de pendiente declarado en campo", "15-25% (Fuert. inclinado)"),
            ("Forma predominante del terreno", "Colinoso",
             "Posición fisiográfica", "Ladera baja"),
            ("Exposición / orientación", "Noreste", "Afloramientos rocosos", "Sí"),
            ("Escarpes activos", "No", "Remociones en masa activas", "No"),
        ])
        ws.append(["3. ÍNDICES DE VEGETACIÓN (SENTINEL-2)"])
        _pares(ws, [
            ("MSAVI 2024 — media del bloque", 0.3143,
             "MSAVI 2024 — clase de la media", "0.2650 - 0.3813 (Vigor muy bajo)"),
            ("Condición frente al umbral 0.4976", "BAJO umbral 0.4976",
             "NDVI mediana 2025 — clase modal", "Vegetación ligera (78.10 %)"),
            ("Superficie clasificada NDVI 2025 (ha)", 355.568,
             "Desviación frente al catálogo (%)", 0.06),
        ])
        ws.append(["4. ECOSISTEMA Y ESTADO DE CONSERVACIÓN"])
        _pares(ws, [
            ("Tipo de ecosistema (UP)", "Bosque Estacionalmente Seco de Colina",
             "Superficie de ecosistema (ha)", "320 ha (de 355.36 ha del bloque)"),
            ("Estado de conservación", "Medianamente alterado",
             "Uso actual dominante del suelo", "Pastoreo extensivo"),
            ("Tipo de cobertura dominante", "Bosque ralo (30-70%)",
             "Cobertura vegetal total — campo (%)", 90),
            ("Suelo desnudo — campo (%)", 10,
             "Regeneración natural", "Regular (10-50 pl./100m²)"),
            ("Nivel general de erosión", "FICHA F-DT-02 SIN CONTENIDO",
             "N.° de cárcavas registradas", "No inventariadas"),
            ("Elenco florístico (n.° de taxones)", 9, "Estado sanitario", "Sano"),
        ])
        ws.append(["5. RESPONSABLE Y MODALIDAD DE LA EVALUACIÓN"])
        _pares(ws, [
            ("Responsable de la evaluación", "Juan Domínguez",
             "Fecha de evaluación", "2026-07-03"),
            ("Hora de registro", "09:55 h",
             "Correlativo de ficha", "Sin registro en ficha"),
            ("Entidad", "ANIN - DIME - SESDI",
             "Fase del estudio", "Preinversión (Perfil) · Invierte.pe"),
            ("Instrumento aplicado", "Fichas F-DT-01 a F-DT-05 (Plantilla V5)",
             "Parcela de muestreo", "Sin registro en ficha"),
            ("Estaciones fotográficas georreferenciadas", 0,
             "Modalidad de acceso", "Vehicular + caminata <30 min"),
        ])
        ws.append(["6. SÍNTESIS PARA LA GESTIÓN DEL RIESGO"])
        _pares(ws, [
            ("Causa subyacente principal", "No consignada en la ficha",
             "Velocidad de degradación", "Moderada"),
            ("Reversibilidad técnica", "Parcialmente reversible",
             "Urgencia de intervención", "Media"),
            ("Urgencia de control de erosión", "Por determinar",
             "Zona de recarga hídrica", "Sin registro en ficha"),
            ("Peligro integrado preliminar (MCA-AHP)", "No disponible",
             "Prioridad de intervención", "Por definir en mesolocalización"),
            ("Estado de verificación de campo", "VERIFICADO",
             "Marco del indicador de brecha", "R.M. N.° 00213-2024-MINAM"),
        ])
        ws.append(["DECLARACIÓN DE INTEGRIDAD DE DATOS. Ningún valor ausente "
                   "ha sido estimado o inferido sin declararlo."])

    if "Cobertura MSAVI-NDVI" in hojas:
        ws = wb.create_sheet("Cobertura MSAVI-NDVI")
        _encabezados(ws, f"DISTRIBUCIÓN AREAL DE CLASES ESPECTRALES — BLOQUE {codigo}", 5)
        ws.append(["A. MSAVI 2024 — CLASIFICACIÓN POR UMBRALES DEL PROYECTO"])
        ws.append(["Clase MSAVI", "Superficie (ha)", "% del área clasificada",
                   "Interpretación", "Condición frente al umbral 0.4976"])
        clases = [("> 0.6139", "Vigor alto", "Sobre umbral"),
                  ("0.4976 - 0.6139", "Vigor moderado", "Sobre umbral"),
                  ("0.3813 - 0.4976", "Vigor bajo", "BAJO umbral 0.4976"),
                  ("0.2650 - 0.3813", "Vigor muy bajo", "BAJO umbral 0.4976"),
                  ("<= 0.2650", "Suelo desnudo", "BAJO umbral 0.4976")]
        areas = [12.5, 44.0, 120.25, 150.0, 28.81]
        for i, (clase, interp, cond) in enumerate(clases):
            if con_msavi_areal:
                ws.append([clase, areas[i], round(areas[i] / 3.5556, 2), interp, cond])
            else:
                ws.append([clase, "Por determinar", "Por determinar", interp, cond])
        ws.append(["MSAVI 2024 — MEDIA DEL BLOQUE", 0.314312, "—",
                   "Vigor muy bajo", "BAJO umbral 0.4976"])
        ws.append(["NOTA METODOLÓGICA. La media del MSAVI 2024 procede del "
                   "catálogo maestro y es dato oficial. La distribución areal "
                   "no está disponible: se consigna «Por determinar»."])
        ws.append([])
        ws.append(["B. NDVI MEDIANA 2025 — DISTRIBUCIÓN AREAL (ESTADÍSTICA ZONAL)"])
        ws.append(["Clase NDVI", "Superficie (ha)", "% del área clasificada",
                   "Interpretación", "Observación"])
        for clase, sup, pct, interp in [
                ("Vegetación alta", 0.5237, 0.15, "Dosel continuo"),
                ("Vegetación mediana", 28.9887, 8.15, "Pastizal cultivado"),
                ("Vegetación ligera", 277.6993, 78.10, "Cobertura discontinua"),
                ("Tierra desnuda", 48.3560, 13.60, "Suelo desnudo")]:
            ws.append([clase, sup, pct, interp, None])
        ws.append(["TOTAL CLASIFICADO", 355.5677, 100.00, None, None])
        ws.append(["Superficie de catálogo (V5/V6)", 355.36, "—",
                   "Desviación planimétrica", "+0.06 %"])
        ws.append(["LECTURA CRÍTICA. Los índices MSAVI y NDVI miden vigor y "
                   "densidad de biomasa, no composición ni integridad."])

    if "Estaciones fotográficas" in hojas:
        ws = wb.create_sheet("Estaciones fotográficas")
        _encabezados(ws, f"PUNTOS GEORREFERENCIADOS DE VERIFICACIÓN — BLOQUE {codigo}", 7)
        ws.append(["INVENTARIO DE PUNTOS GEORREFERENCIADOS DE LA FICHA DT"])
        ws.append(["Código", "Naturaleza del punto", "UTM ESTE (m)",
                   "UTM NORTE (m)", "Dist. al centroide (m)", "Estado / régimen",
                   "Contenido registrado"])
        ws.append(["P-00", "Punto de muestreo declarado en ficha DT", "645,575",
                   "9,393,500", 76381.5, None, "Coordenada consignada en F-DT-01"])
        ws.append(["P-01", "Rasgo erosivo", 595600, 9451300, 95.2, "Activo",
                   "Erosión laminar"])
        ws.append(["TOTAL: 2 puntos"])
        ws.append(["Centroide del bloque — UTM ESTE (m)", 595552,
                   "Centroide del bloque — UTM NORTE (m)", 9451222])
        ws.append(["N.° de estaciones fotográficas declaradas", 0,
                   "Estaciones dentro del polígono", 0])
        ws.append(["CONTROL GEOMÉTRICO. Las coordenadas proceden de la ficha DT."])

    if "Microcuenca" in hojas:
        ws = wb.create_sheet("Microcuenca")
        _encabezados(ws, f"CONTEXTO INTRAMICROCUENCA C1096-Q9545 — BLOQUE {codigo}", 8)
        ws.append(["BLOQUES DE LA MICROCUENCA C1096-Q9545"])
        ws.append(["Bloque", "Área (ha)", "% microcuenca", "Rango altitudinal (msnm)",
                   "Amplitud (m)", "Pendiente prom. (%)", "MSAVI 2024",
                   "NDVI 2025 — Veg. alta (%)"])
        ws.append(["3", 459.290, 56.38, "951 – 2173", 1222, 29.52, 0.5677, 57.21])
        ws.append([f"► {codigo}", 355.360, 43.62, "147 – 265", 118, 10.17, 0.3143, 0.15])
        ws.append(["TOTAL / PROMEDIO", 814.650, 100.00, "147 – 2173", None,
                   19.85, 0.4410, 28.68])
        ws.append(["LECTURA INTRAMICROCUENCA. El bloque ocupa el puesto 2 de 2."])

    if "Control de consistencia" in hojas:
        ws = wb.create_sheet("Control de consistencia")
        _encabezados(ws, f"CONTROL DE CONSISTENCIA DE LA INFORMACIÓN — BLOQUE {codigo}", 5)
        ws.append(["DISCREPANCIAS Y VERIFICACIONES"])
        ws.append(["Cód.", "Campo afectado", "Discrepancia observada",
                   "Calificación", "Tratamiento adoptado"])
        plantilla = [
            ("Código de microcuenca", "La ficha declara C1076-Q9584", "SUSTANTIVA",
             "Se conserva el código del catálogo maestro"),
            ("Superficie del polígono", "Coincide dentro del ±2 %", "CONFORME",
             "Sin acción: la segmentación valida la geometría"),
            ("Rango de pendiente", "Campo 15-25% frente a MDE 10.17 %",
             "NO SUSTANTIVA", "Se adopta el valor del MDE"),
            ("Plan de Ingreso", "Altitud = 0 msnm por defecto", "CORREGIDO",
             "Se reemplaza por la altitud GPS de campo"),
        ]
        for i in range(n_discrepancias):
            campo, disc, cal, trat = plantilla[i % len(plantilla)]
            ws.append([f"D-{i + 1:02d}", campo, disc, cal, trat])
        conteo = {}
        for i in range(n_discrepancias):
            conteo[plantilla[i % len(plantilla)][2]] = \
                conteo.get(plantilla[i % len(plantilla)][2], 0) + 1
        ws.append(["RESUMEN", f"{n_discrepancias} verificaciones",
                   " · ".join(f"{k}: {v}" for k, v in sorted(conteo.items()))])
        ws.append(["NOTA METODOLÓGICA. Este control se genera de forma reproducible."])

    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue()


# ══════════════════════════════════════════════════════════════════════════
# Pruebas
# ══════════════════════════════════════════════════════════════════════════

class TestParseo(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.contenido = construir_libro("M9B1")
        cls.datos = rb.parsear_resumen_bloque(
            cls.contenido, "Plantilla_Excel_Bloque_M9B1_IN_Piura.xlsx")

    def test_lee_las_cinco_hojas(self):
        self.assertEqual(len(self.datos["hojas_leidas"]), 5)
        self.assertEqual(self.datos["advertencias"], [])

    def test_identificacion(self):
        d = self.datos
        self.assertEqual(d["codigo_bloque"], "M9B1")
        self.assertEqual(d["microcuenca"], "C1096-Q9545")
        self.assertEqual(d["provincia"], "Morropon")
        self.assertEqual(d["distrito"], "Chulucanas")
        self.assertEqual(d["departamento"], "Piura")
        self.assertEqual(d["centro_poblado"], "La Peña")
        self.assertEqual(d["tipo_intervencion"], "Restauración")

    def test_valores_numericos(self):
        d = self.datos
        self.assertAlmostEqual(d["area_ha_num"], 355.36, places=2)
        self.assertAlmostEqual(d["utm_este_num"], 595552)
        self.assertAlmostEqual(d["utm_norte_num"], 9451222)
        self.assertAlmostEqual(d["pendiente_pct_num"], 10.17, places=2)
        self.assertAlmostEqual(d["msavi_2024_num"], 0.3143, places=4)
        self.assertAlmostEqual(d["altitud_min_num"], 147)
        self.assertAlmostEqual(d["altitud_max_num"], 265)

    def test_etiquetas_de_ambas_columnas(self):
        """Las etiquetas de la columna C se leen igual que las de la A."""
        self.assertEqual(self.datos["estado_conservacion"], "Medianamente alterado")
        self.assertEqual(self.datos["capital_distrital"], "Chulucanas")
        self.assertEqual(self.datos["estado_sanitario"], "Sano")

    def test_validacion_utm_17s(self):
        self.assertEqual(self.datos["validacion_utm"], "Conforme")
        self.assertEqual(rb.validar_utm(100000, 9451222)[:4], "ESTE")
        self.assertEqual(rb.validar_utm(595552, 100)[:5], "NORTE")
        self.assertEqual(rb.validar_utm(None, None), "Sin coordenadas")

    def test_tabla_ndvi(self):
        ndvi = self.datos["ndvi_tabla"]
        self.assertEqual(len(ndvi), 4)
        self.assertEqual(ndvi[0]["clase"], "Vegetación alta")
        self.assertAlmostEqual(ndvi[2]["superficie_ha"], 277.6993, places=4)
        self.assertAlmostEqual(ndvi[2]["pct"], 78.10, places=2)
        # La fila TOTAL no debe entrar como una clase mas.
        self.assertNotIn("TOTAL CLASIFICADO", [r["clase"] for r in ndvi])
        self.assertAlmostEqual(self.datos["ndvi_total_ha"], 355.5677, places=4)

    def test_msavi_sin_distribucion_areal_no_se_estima(self):
        """El caso mayoritario: 'Por determinar' no debe volverse 0."""
        for fila in self.datos["msavi_tabla"]:
            self.assertIsNone(fila["superficie_ha"])
            self.assertEqual(fila["superficie_ha_txt"], "Por determinar")

    def test_msavi_con_distribucion_areal(self):
        datos = rb.parsear_resumen_bloque(construir_libro(con_msavi_areal=True))
        areas = [f["superficie_ha"] for f in datos["msavi_tabla"]]
        self.assertEqual(len(areas), 5)
        self.assertAlmostEqual(areas[3], 150.0)

    def test_estaciones(self):
        est = self.datos["estaciones"]
        self.assertEqual(len(est), 2)
        # Coordenada con separador de miles.
        self.assertAlmostEqual(est[0]["utm_este"], 645575)
        self.assertAlmostEqual(est[0]["utm_norte"], 9393500)
        self.assertAlmostEqual(est[1]["dist_centroide"], 95.2, places=1)
        self.assertEqual(self.datos["n_estaciones_declaradas"], 0)

    def test_microcuenca_marca_el_bloque_actual(self):
        micro = self.datos["microcuenca_tabla"]
        self.assertEqual(len(micro), 2)
        actual = [r for r in micro if r["es_actual"]]
        self.assertEqual(len(actual), 1)
        self.assertEqual(actual[0]["bloque"], "M9B1")
        self.assertAlmostEqual(actual[0]["area_ha"], 355.36, places=2)
        self.assertNotIn("TOTAL / PROMEDIO", [r["bloque"] for r in micro])

    def test_consistencia_y_conteo(self):
        cons = self.datos["consistencia"]
        self.assertEqual(len(cons), 4)
        self.assertEqual(cons[0]["codigo"], "D-01")
        resumen = self.datos["consistencia_resumen"]
        self.assertEqual(resumen["total"], 4)
        self.assertEqual(resumen["SUSTANTIVA"], 1)
        self.assertEqual(resumen["CONFORME"], 1)
        self.assertEqual(resumen["NO SUSTANTIVA"], 1)
        self.assertEqual(resumen["CORREGIDO"], 1)

    def test_numero_variable_de_discrepancias(self):
        for n in (1, 8, 19):
            datos = rb.parsear_resumen_bloque(construir_libro(n_discrepancias=n))
            self.assertEqual(len(datos["consistencia"]), n)
            self.assertEqual(datos["consistencia_resumen"]["total"], n)


class TestRobustez(unittest.TestCase):

    def test_libro_sin_algunas_hojas(self):
        """Un libro parcial se carga igual y declara lo que falta."""
        datos = rb.parsear_resumen_bloque(
            construir_libro(hojas=["Resumen"]),
            "Plantilla_Excel_Bloque_M9B1_IN_Piura.xlsx")
        self.assertEqual(datos["codigo_bloque"], "M9B1")
        self.assertEqual(datos["hojas_leidas"], ["Resumen"])
        self.assertEqual(len(datos["advertencias"]), 4)
        self.assertEqual(datos["ndvi_tabla"], [])
        self.assertEqual(datos["consistencia_resumen"]["total"], 0)

    def test_codigo_desde_nombre_de_archivo(self):
        self.assertEqual(
            rb.codigo_desde_nombre("Plantilla_Excel_Bloque_M6B2-1_IN_Piura.xlsx"),
            "M6B2-1")
        self.assertEqual(
            rb.codigo_desde_nombre("ruta/Plantilla_Excel_Bloque_28_IN_Piura.xlsx"),
            "28")
        self.assertEqual(rb.codigo_desde_nombre("otro.xlsx"), "")

    def test_codigo_de_respaldo_desde_nombre(self):
        """Sin hoja Resumen, el codigo se toma del nombre de archivo."""
        datos = rb.parsear_resumen_bloque(
            construir_libro(hojas=["Microcuenca"]),
            "Plantilla_Excel_Bloque_28_IN_Piura.xlsx")
        self.assertEqual(datos["codigo_bloque"], "28")

    def test_numero_no_inventa_valores(self):
        for entrada in (None, "", "Por determinar", "Sin registro en ficha", "—"):
            self.assertIsNone(rb._num(entrada))
        self.assertAlmostEqual(rb._num("645,575"), 645575)
        self.assertAlmostEqual(rb._num("+0.06 %"), 0.06)
        self.assertAlmostEqual(rb._num("9 451 222"), 9451222)

    def test_lote_aisla_archivos_ilegibles(self):
        buenos_y_malos = [
            ("Plantilla_Excel_Bloque_M9B1_IN_Piura.xlsx", construir_libro("M9B1")),
            ("roto.xlsx", b"esto no es un xlsx"),
            ("Plantilla_Excel_Bloque_28_IN_Piura.xlsx", construir_libro("28")),
        ]
        resultados, errores = rb.parsear_lote(buenos_y_malos)
        self.assertEqual(len(resultados), 2)
        self.assertEqual(len(errores), 1)
        self.assertEqual(errores[0][0], "roto.xlsx")

    def test_expandir_zip(self):
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as z:
            z.writestr("salida/Plantilla_Excel_Bloque_M9B1_IN_Piura.xlsx",
                       construir_libro("M9B1"))
            z.writestr("salida/Plantilla_Excel_Bloque_28_IN_Piura.xlsx",
                       construir_libro("28"))
            z.writestr("salida/README.md", "no es un libro")
            z.writestr("__MACOSX/._Plantilla.xlsx", b"basura")
        archivos = rb.expandir_zip(buffer.getvalue())
        self.assertEqual(len(archivos), 2)
        self.assertTrue(all(n.endswith(".xlsx") for n, _ in archivos))


class TestManifiesto(unittest.TestCase):

    def test_manifiesto_de_117_bloques(self):
        codigos = rb.codigos_esperados()
        self.assertEqual(len(codigos), 117)
        self.assertEqual(len(set(codigos)), 117)
        for esperado in ("M9B1", "28", "M6B2-1", "M22B1"):
            self.assertIn(esperado, codigos)

    def test_manifiesto_trae_enlace_a_drive(self):
        bloques = rb.cargar_manifiesto()["bloques"]
        self.assertTrue(all(b["drive_url"].startswith("https://drive.google.com")
                            for b in bloques))
        self.assertTrue(all(b["archivo"].endswith(".xlsx") for b in bloques))


class TestExportaciones(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.contenido = construir_libro("M9B1")
        cls.datos = rb.parsear_resumen_bloque(
            cls.contenido, "Plantilla_Excel_Bloque_M9B1_IN_Piura.xlsx")

    def test_excel_con_graficos_conserva_hojas_originales(self):
        salida = rb.generar_excel_con_graficos(self.contenido, self.datos)
        wb = load_workbook(io.BytesIO(salida))
        for hoja in ("Resumen", "Cobertura MSAVI-NDVI", "Estaciones fotográficas",
                     "Microcuenca", "Control de consistencia"):
            self.assertIn(hoja, wb.sheetnames)
        self.assertIn("Graficos", wb.sheetnames)
        self.assertGreaterEqual(len(wb["Graficos"]._charts), 3)
        # El valor original permanece intacto.
        self.assertEqual(wb["Resumen"]["B8"].value, "M9B1")

    def test_regenerar_graficos_no_duplica_la_hoja(self):
        una = rb.generar_excel_con_graficos(self.contenido, self.datos)
        dos = rb.generar_excel_con_graficos(una, self.datos)
        wb = load_workbook(io.BytesIO(dos))
        self.assertEqual(wb.sheetnames.count("Graficos"), 1)

    def test_excel_consolidado(self):
        datos_28 = rb.parsear_resumen_bloque(construir_libro("28"))
        salida = rb.generar_excel_consolidado([self.datos, datos_28])
        wb = load_workbook(io.BytesIO(salida))
        self.assertEqual(wb.sheetnames, ["Consolidado", "Graficos"])
        self.assertGreaterEqual(len(wb["Graficos"]._charts), 3)
        # Los totales van como formula, no como valor precalculado.
        textos = [c.value for fila in wb["Consolidado"].iter_rows()
                  for c in fila if isinstance(c.value, str)]
        self.assertTrue(any(t.startswith("=SUM(") for t in textos))

    def test_pdf_de_bloque(self):
        pdf = rb.generar_pdf_bloque(self.datos)
        self.assertTrue(pdf.startswith(b"%PDF"))
        self.assertGreater(len(pdf), 4000)

    def test_pdf_con_distribucion_msavi(self):
        datos = rb.parsear_resumen_bloque(construir_libro(con_msavi_areal=True))
        self.assertTrue(rb.generar_pdf_bloque(datos).startswith(b"%PDF"))

    def test_pdf_de_libro_parcial(self):
        datos = rb.parsear_resumen_bloque(construir_libro(hojas=["Resumen"]))
        self.assertTrue(rb.generar_pdf_bloque(datos).startswith(b"%PDF"))

    def test_pdf_consolidado(self):
        datos = [rb.parsear_resumen_bloque(construir_libro(f"M{i}B1"))
                 for i in range(1, 6)]
        pdf = rb.generar_pdf_consolidado(datos)
        self.assertTrue(pdf.startswith(b"%PDF"))
        # El reporte debe crecer con el numero de bloques listados.
        self.assertGreater(len(pdf), len(rb.generar_pdf_consolidado(datos[:1])))

    def test_pdf_maneja_texto_no_latin1(self):
        contenido = construir_libro("M9B1")
        datos = rb.parsear_resumen_bloque(contenido)
        datos["centro_poblado"] = "Peña — «Ñañañán» ► 30 m² ≈ 5"
        self.assertTrue(rb.generar_pdf_bloque(datos).startswith(b"%PDF"))


if __name__ == "__main__":
    unittest.main(verbosity=2)


# ══════════════════════════════════════════════════════════════════════════
# Maquetacion de los reportes
# ══════════════════════════════════════════════════════════════════════════

def cajas_de_graficos(ws):
    """Caja (x0, y0, x1, y1) en cm de cada grafico anclado en la hoja.

    Un grafico no ocupa celdas: flota sobre la rejilla desde su celda de
    anclaje con el tamano que guarda el anclaje (en EMU). Para saber si dos
    se encimarian al abrir el libro hay que reconstruir esas cajas.
    """
    emu_cm = 360000.0
    punto_cm = 2.54 / 72
    ancho_defecto = (8.43 * 7 + 5) / 96 * 2.54

    def alto_fila(fila):
        dim = ws.row_dimensions.get(fila)
        return (dim.height if dim is not None and dim.height else 15) * punto_cm

    def ancho_col(col):
        from openpyxl.utils import get_column_letter
        dim = ws.column_dimensions.get(get_column_letter(col))
        if dim is None or not dim.width:
            return ancho_defecto
        return (dim.width * 7 + 5) / 96 * 2.54

    cajas = []
    for grafico in ws._charts:
        origen = grafico.anchor._from
        x0 = sum(ancho_col(c) for c in range(1, origen.col + 1))
        y0 = sum(alto_fila(f) for f in range(1, origen.row + 1))
        ext = getattr(grafico.anchor, "ext", None)
        ancho = ext.cx / emu_cm if ext is not None else 12.0
        alto = ext.cy / emu_cm if ext is not None else 8.0
        cajas.append((x0, y0, x0 + ancho, y0 + alto))
    return cajas


def graficos_encimados(ws):
    """Pares de graficos cuyas cajas se solapan."""
    cajas = cajas_de_graficos(ws)
    encimados = []
    for i in range(len(cajas)):
        for j in range(i + 1, len(cajas)):
            ax0, ay0, ax1, ay1 = cajas[i]
            bx0, by0, bx1, by1 = cajas[j]
            if ax0 < bx1 and bx0 < ax1 and ay0 < by1 and by0 < ay1:
                encimados.append((i + 1, j + 1))
    return encimados


class TestHojaDeGraficos(unittest.TestCase):
    """Los graficos de la hoja se apilan; ninguno puede pisar al anterior."""

    def setUp(self):
        self.libro = construir_libro(codigo="M9B1", con_msavi_areal=True)
        self.datos = rb.completar_sintesis_msavi(
            rb.parsear_resumen_bloque(self.libro, "Plantilla_Excel_Bloque_M9B1_IN_Piura.xlsx"))

    def _hoja_graficos(self):
        contenido = rb.generar_excel_con_graficos(self.libro, self.datos)
        return load_workbook(io.BytesIO(contenido))["Graficos"]

    def test_la_hoja_trae_un_grafico_por_serie(self):
        # NDVI (torta + barras), MSAVI, microcuenca (dos) y consistencia.
        self.assertGreaterEqual(len(self._hoja_graficos()._charts), 4)

    def test_ningun_grafico_se_encima_con_otro(self):
        self.assertEqual(graficos_encimados(self._hoja_graficos()), [])

    def test_los_graficos_no_invaden_la_tabla_siguiente(self):
        """Cada seccion empieza debajo de los graficos de la anterior."""
        ws = self._hoja_graficos()
        cajas = cajas_de_graficos(ws)
        punto_cm = 2.54 / 72
        for titulo in ("B. MSAVI", "C. Contexto", "D. Control"):
            fila = next((c.row for col in ws.iter_cols(min_col=1, max_col=1)
                         for c in col
                         if str(c.value or "").startswith(titulo)), None)
            if fila is None:
                continue
            y_titulo = sum(
                (ws.row_dimensions[f].height
                 if f in ws.row_dimensions and ws.row_dimensions[f].height
                 else 15) * punto_cm for f in range(1, fila))
            previos = [caja for caja in cajas if caja[1] < y_titulo]
            for caja in previos:
                self.assertLessEqual(
                    caja[3], y_titulo + 0.01,
                    f"un grafico llega a {caja[3]:.1f} cm y la seccion "
                    f"«{titulo}» empieza en {y_titulo:.1f} cm")

    def test_el_consolidado_tampoco_encima_sus_graficos(self):
        contenido = rb.generar_excel_consolidado([self.datos, self.datos])
        wb = load_workbook(io.BytesIO(contenido))
        self.assertEqual(graficos_encimados(wb["Graficos"]), [])


class TestTablaDelPDF(unittest.TestCase):
    """Una celda de varias lineas no debe desparramar la tabla.

    `multi_cell` toma su segundo argumento como alto de CADA linea: si se le
    pasa el alto de la fila, una celda larga se estira hasta invadir las
    filas siguientes y la tabla ocupa varias paginas casi vacias.
    """

    TEXTO = ("Se adopta el valor de campo por ser la fuente que manda sobre "
             "el dato y la mas reciente. El valor anterior queda declarado "
             "para trazabilidad, conforme a la nota metodologica de la hoja "
             "de control de consistencia del bloque.")

    def _pdf(self, n_filas):
        pdf = rb.PDFResumen(subtitulo="PRUEBA DE MAQUETACION")
        pdf.alias_nb_pages()
        pdf.add_page()
        pdf.tabla(["Cod.", "Campo afectado", "Discrepancia observada",
                   "Calificacion", "Tratamiento adoptado"],
                  [[f"D-{i:02d}", "Campo afectado", self.TEXTO, "CONFORME",
                    self.TEXTO] for i in range(n_filas)],
                  anchos_rel=[0.5, 1.6, 3.4, 1.0, 3.2])
        return rb.pdf_bytes(pdf)

    def test_diez_filas_largas_caben_en_una_pagina(self):
        import pdfplumber
        with pdfplumber.open(io.BytesIO(self._pdf(10))) as pdf:
            self.assertEqual(len(pdf.pages), 1)

    def test_las_lineas_de_una_celda_van_seguidas(self):
        """Dentro de una celda las lineas se separan por el interlineado."""
        import pdfplumber
        with pdfplumber.open(io.BytesIO(self._pdf(4))) as pdf:
            pagina = pdf.pages[0]
            # Solo la caja de la tabla: fuera quedan la banda institucional
            # de arriba y el pie de pagina, separados por diseno.
            alturas = sorted({round(p["top"], 1) for p in pagina.extract_words()
                              if 95 <= p["top"] <= pagina.height - 60})
        saltos = [b - a for a, b in zip(alturas, alturas[1:])]
        # El interlineado es de ~3.85 mm (11 pt). Un salto de mas de 25 pt
        # significa que una celda se estiro y arrastro el resto de la fila.
        self.assertLess(max(saltos), 25, f"saltos verticales: {saltos}")

    def test_la_tabla_vacia_no_dibuja_nada(self):
        pdf = rb.PDFResumen()
        pdf.add_page()
        y = pdf.get_y()
        pdf.tabla(["A", "B"], [])
        self.assertEqual(pdf.get_y(), y)


class TestLibrosDelRepositorio(unittest.TestCase):
    """Los 117 libros de gabinete (V6), linea base de la integracion.

    El aplicativo carga los libros ya integrados con la ficha de campo (V7);
    esta clase comprueba la carpeta de origen, que es la que sostiene el
    traspaso del MSAVI por clase DN.
    """

    @classmethod
    def setUpClass(cls):
        raiz = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        cls.libros = rb.libros_del_repositorio(
            os.path.join(raiz, rb.CARPETA_LIBROS_V6))

    def test_el_repositorio_trae_los_117_libros_del_manifiesto(self):
        self.assertEqual(len(self.libros), 117)
        nombres = {n for n, _ in self.libros}
        esperados = {b["archivo"] for b in rb.cargar_manifiesto()["bloques"]}
        self.assertEqual(nombres, esperados)

    def test_carpeta_ausente_no_rompe_la_carga(self):
        self.assertEqual(rb.libros_del_repositorio("/no/existe"), [])

    def test_cada_libro_declara_su_codigo_de_bloque(self):
        codigos = set()
        for nombre, contenido in self.libros:
            datos = rb.parsear_resumen_bloque(contenido, nombre)
            self.assertTrue(datos.get("codigo_bloque"), nombre)
            codigos.add(datos["codigo_bloque"])
        self.assertEqual(codigos, set(rb.codigos_esperados()))

    def test_la_distribucion_msavi_trae_superficie_y_porcentaje(self):
        """Las cinco clases DN con valor: ninguna queda «Por determinar».

        El porcentaje es una fórmula en el libro; se comprueba que llegue con
        su valor en caché, porque el aplicativo lee valores, no fórmulas.
        """
        for nombre, contenido in self.libros:
            datos = rb.parsear_resumen_bloque(contenido, nombre)
            tabla = datos.get("msavi_tabla") or []
            self.assertEqual(len(tabla), 5, nombre)
            for fila in tabla:
                self.assertIsNotNone(fila.get("superficie_ha"), nombre)
                self.assertIsNotNone(fila.get("pct"), nombre)

    def test_la_sintesis_msavi_cuadra_con_las_clases(self):
        for nombre, contenido in self.libros:
            datos = rb.parsear_resumen_bloque(contenido, nombre)
            clases = sum(f["superficie_ha"] for f in datos["msavi_tabla"])
            self.assertAlmostEqual(datos["msavi_total_ha"], clases, places=3,
                                   msg=nombre)
            self.assertAlmostEqual(
                datos["msavi_sobre_umbral_ha"] + datos["msavi_bajo_umbral_ha"],
                datos["msavi_total_ha"], places=3, msg=nombre)

    def test_la_hoja_resumen_replica_la_sintesis_msavi(self):
        for nombre, contenido in self.libros:
            datos = rb.parsear_resumen_bloque(contenido, nombre)
            self.assertAlmostEqual(datos["superficie_msavi_ha_num"],
                                   datos["msavi_total_ha"], places=2, msg=nombre)
            self.assertAlmostEqual(datos["superficie_bajo_umbral_ha_num"],
                                   datos["msavi_bajo_umbral_ha"], places=2,
                                   msg=nombre)
            self.assertTrue(datos.get("msavi_clase_dominante"), nombre)
            self.assertGreater(datos["msavi_poligonos_num"], 0, nombre)

    def test_el_control_de_consistencia_registra_el_traspaso(self):
        for nombre, contenido in self.libros:
            datos = rb.parsear_resumen_bloque(contenido, nombre)
            campos = [f["campo"] for f in datos["consistencia"]]
            self.assertIn("Distribución areal MSAVI 2024", campos, nombre)
            # La serie D queda correlativa; otras series (C-01: hallazgos
            # sobre el archivo de la ficha DT) conservan su código.
            serie_d = [f["codigo"] for f in datos["consistencia"]
                       if f["codigo"].startswith("D-")]
            self.assertEqual(serie_d,
                             [f"D-{i:02d}" for i in range(1, len(serie_d) + 1)],
                             nombre)
            total = datos["consistencia_resumen"]["total"]
            self.assertEqual(total, len(datos["consistencia"]), nombre)


class TestSintesisMsaviCalculada(unittest.TestCase):
    """Reparto porcentual derivado cuando el libro no trae el resultado.

    Reproduce el caso de los resumenes cargados desde libros guardados sin
    recalcular: las superficies estan, el porcentaje y los totales no.
    """

    def _datos_sin_porcentaje(self):
        return {
            "msavi_tabla": [
                {"clase": "> 0.6139 · DN 5", "superficie_ha": 106.9938,
                 "pct": None, "pct_txt": "", "condicion": "Sobre umbral"},
                {"clase": "0.4976 - 0.6139 · DN 4", "superficie_ha": 12.4651,
                 "pct": None, "pct_txt": "", "condicion": "Sobre umbral"},
                {"clase": "0.3813 - 0.4976 · DN 3", "superficie_ha": 0.0792,
                 "pct": None, "pct_txt": "", "condicion": "BAJO umbral 0.4976"},
                {"clase": "0.2650 - 0.3813 · DN 2", "superficie_ha": 0.0,
                 "pct": None, "pct_txt": "", "condicion": "BAJO umbral 0.4976"},
                {"clase": "<= 0.2650 · DN 1", "superficie_ha": 0.0,
                 "pct": None, "pct_txt": "", "condicion": "BAJO umbral 0.4976"},
            ]
        }

    def test_deriva_el_porcentaje_de_cada_clase(self):
        datos = rb.completar_sintesis_msavi(self._datos_sin_porcentaje())
        self.assertEqual([f["pct"] for f in datos["msavi_tabla"]],
                         [89.51, 10.43, 0.07, 0.0, 0.0])
        self.assertTrue(all(f["pct_txt"] for f in datos["msavi_tabla"]))
        self.assertTrue(all(f["pct_calculado"] for f in datos["msavi_tabla"]))
        self.assertTrue(datos["msavi_sintesis_calculada"])

    def test_deriva_los_totales_y_el_reparto_por_umbral(self):
        datos = rb.completar_sintesis_msavi(self._datos_sin_porcentaje())
        self.assertAlmostEqual(datos["msavi_total_ha"], 119.5381, places=4)
        self.assertAlmostEqual(datos["msavi_sobre_umbral_ha"], 119.4589, places=4)
        self.assertAlmostEqual(datos["msavi_bajo_umbral_ha"], 0.0792, places=4)
        self.assertAlmostEqual(datos["msavi_bajo_umbral_pct"], 0.07, places=2)

    def test_lo_derivado_coincide_con_lo_que_declara_el_libro(self):
        for nombre, contenido in rb.libros_del_repositorio()[:12]:
            declarado = rb.parsear_resumen_bloque(contenido, nombre)
            desnudo = {"msavi_tabla": [
                {k: (None if k in ("pct", "pct_txt") else v)
                 for k, v in fila.items()}
                for fila in declarado["msavi_tabla"]]}
            derivado = rb.completar_sintesis_msavi(desnudo)
            self.assertEqual([f["pct"] for f in derivado["msavi_tabla"]],
                             [f["pct"] for f in declarado["msavi_tabla"]], nombre)
            self.assertAlmostEqual(derivado["msavi_total_ha"],
                                   declarado["msavi_total_ha"], places=3, msg=nombre)
            self.assertAlmostEqual(derivado["msavi_bajo_umbral_ha"],
                                   declarado["msavi_bajo_umbral_ha"], places=3,
                                   msg=nombre)

    def test_no_reescribe_lo_que_el_libro_declara(self):
        datos = {"msavi_tabla": [
            {"clase": "> 0.6139", "superficie_ha": 10.0, "pct": 99.0,
             "pct_txt": "99.00", "condicion": "Sobre umbral"}],
            "msavi_total_ha": 12.0}
        salida = rb.completar_sintesis_msavi(datos)
        self.assertEqual(salida["msavi_tabla"][0]["pct"], 99.0)
        self.assertEqual(salida["msavi_total_ha"], 12.0)
        self.assertNotIn("pct_calculado", salida["msavi_tabla"][0])

    def test_sin_superficies_no_inventa_nada(self):
        datos = {"msavi_tabla": [
            {"clase": "> 0.6139", "superficie_ha": None,
             "superficie_ha_txt": "Por determinar", "pct": None,
             "condicion": "Sobre umbral"}]}
        salida = rb.completar_sintesis_msavi(datos)
        self.assertIsNone(salida["msavi_tabla"][0]["pct"])
        self.assertIsNone(salida.get("msavi_total_ha"))
        self.assertNotIn("msavi_sintesis_calculada", salida)

    def test_tabla_vacia_no_rompe(self):
        self.assertEqual(rb.completar_sintesis_msavi({}), {})
