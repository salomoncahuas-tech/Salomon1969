"""
Paquetes 01-05 (+ nota del 06, informe de remision e indice) de la respuesta a la solicitud de informacion
del Programa de Inversion del Rio Piura sobre el Perfil «Infraestructura Natural» (CUI 2669244).

Uso: python build_docs.py --repo <raiz> --drive <descargas Drive> --paquetes <carpeta de paquetes>
Las cifras se leen de las fuentes (datos.py); los textos citan el documento de origen.
"""
import argparse
import json
import shutil
from pathlib import Path

import geopandas as gpd
import openpyxl
from openpyxl import Workbook

import anin_excel_fmt as X
import anin_word as W
import datos

FECHA = "12 de noviembre de 2026"
CORTE = "Corte de información: 12/11/2026 (E2 B4 al 07/11/2026; línea base E6 al 09/11/2026; EVAR al 07/11/2026)"

CARPETAS = {
    "01": "01_Delimitacion_Ambito_Estudio",
    "02": "02_Criterios_Seleccion",
    "03": "03_Diagnostico_Poblacion_Afectada",
    "04": "04_Diagnostico_Unidades_Productoras",
    "05": "05_Diagnostico_Involucrados",
    "06": "06_Informacion_Georreferenciada_Cartografica",
}

POR_ADJUNTAR = [
    ("02", "Nota Técnica de Meso y Microlocalización de las intervenciones con IN (ANIN-DIME, 2023), .docx de 18.3 MB",
     "Supera el límite de descarga del conector de Drive (10 MB); adjuntar desde el archivo del equipo"),
    ("04", "Capa «Peligro por movimientos en masa» del E8 (shapefile de 112 MB) y shapefiles/CSV de vulnerabilidad y riesgo del EVAR",
     "Supera el límite de descarga; remitir por enlace institucional o medio físico"),
    ("04", "Parcelas virtuales del E6 (Parcelas_Virtuales_E6_IN_Piura.gpkg, 17.9 MB) y mapas de ecosistemas de los 117 bloques V7 (PDF, 58.4 MB)",
     "Supera el límite de descarga; remitir por enlace institucional"),
    ("06", "Mapas SERRN por bloque V7 (Mapas_SERRN_AMA_Bloques_V7_IN_Piura.zip, 71.3 MB) y MSAVI vectorizado (13.9 MB)",
     "Supera el límite de descarga; remitir por enlace institucional"),
    ("06", "Proyectos QGIS (.qgz/.qgs), estilos (.qml) y rásteres del modelamiento de mesolocalización (PMM, PHI, EHS, PI, MSAVI, Theil-Sen, DEM)",
     "Por verificar su disponibilidad con el equipo SIG; no se ubicaron en el repositorio ni en Drive"),
    ("06", "Polígonos del catálogo V8 (recorte de los bloques 8, 71, 22, M17B1 y 5)",
     "Pendiente en el equipo SIG; se remite la geometría V6/V7 y el descuento V8 en atributos"),
]


def doc_base(titulo, subtitulo, asunto, horizontal=False):
    d = W.nuevo(horizontal)
    W.titulo_doc(d, titulo, subtitulo)
    W.datos(d, [("Proyecto", "Recuperación del servicio de regulación de riesgos naturales y recuperación de ecosistemas "
                              "degradados en la Cuenca Alta del Río Piura (Perfil «Infraestructura Natural»)"),
                ("CUI / sistema", "2669244 · Invierte.pe · Fase de preinversión (Estudio de Perfil)"),
                ("Asunto", asunto),
                ("Elaborado por", "Ing. Héctor Salomón Cahuas Miller – Especialista en Infraestructura Verde, DIME / SESDI"),
                ("Fecha", f"Lima, {FECHA}"),
                ("Corte", CORTE)])
    return d


def copiar(src, dst_dir, nombre=None):
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = dst_dir / (nombre or Path(src).name)
    shutil.copy2(src, dst)
    return dst


def lista_anexos(d, filas):
    W.h(d, "Archivos que integran el paquete", 1)
    zw = "\u200b"
    W.tabla(d, ["N.°", "Archivo", "Contenido", "Fecha / versión"],
            [[str(i + 1), f[0].replace("_", "_" + zw).replace("/", "/" + zw)] + list(f[1:]) for i, f in enumerate(filas)],
            anchos=[0.06, 0.36, 0.40, 0.18], size=8)


def pendientes(d, cod):
    filas = [[t, m] for c, t, m in POR_ADJUNTAR if c == cod]
    if filas:
        W.h(d, "Información por adjuntar", 2)
        W.tabla(d, ["Archivo o insumo", "Situación"], filas, anchos=[0.6, 0.4], size=8)


def prov_tit(p):
    return {"MORROPON": "Morropón", "HUANCABAMBA": "Huancabamba", "AYABACA": "Ayabaca"}.get(p, p)


# ======================================================================================== 01
def paquete_01(D, repo, drv, out):
    an = out / "Anexos"
    d = doc_base("NOTA TÉCNICA", "Delimitación del ámbito de estudio, área de intervención y área de influencia",
                 "Información de avance del Perfil para su articulación con el Programa de Inversión del Río Piura – "
                 "Componente 1: delimitación del ámbito de estudio")
    W.h(d, "1. Objeto")
    W.p(d, "Presentar la delimitación vigente del ámbito del Perfil «Infraestructura Natural» —área de estudio, área de "
           "intervención (bloques) y área de influencia— y los criterios técnicos con que se definió cada una, para que el "
           "equipo del Programa de Inversión del Río Piura incorpore el alcance territorial del estudio.")
    W.h(d, "2. Ámbitos del estudio")
    W.p(d, "El Perfil diferencia cuatro ámbitos, según la Nota Técnica N.° 001-2026-ANIN/DIME/SESDI (14/10/2026), que "
           "mantiene la delimitación de los bloques y del área de influencia directa y propone delimitar por separado la "
           "zona de beneficio aguas abajo:")
    W.tabla(d, ["Ámbito", "Definición", "Criterio de delimitación", "Situación al corte"], [
        ["Área de estudio", "Envolvente de los ámbitos del proyecto: 3 provincias y 15 distritos de la cuenca alta del río "
         f"Piura ({W.n(D['ambito_ha'], 0)} ha de superficie distrital referencial)",
         "Guía General de Invierte.pe; límites distritales INEI y microcuencas priorizadas",
         "Se actualiza al cerrar la zona de beneficio"],
        ["Área de intervención", f"{D['bloques_n']} bloques preliminares ({W.n(D['bloques_ha'], 3)} ha V6/V7; "
         f"{W.n(D['demanda_total']['Superficie neta V8 (ha)'], 2)} ha netas en el catálogo V8)",
         "Meso y microlocalización (peligro integrado MCA-AHP, exclusión, idoneidad, tamizaje predial, campo)",
         "Vigente; rango del proyecto 10,000 a 15,000 ha; la meta física se fija en Formulación"],
        ["Área de influencia directa", f"{D['ai_n']} polígonos asociados a los {D['ai_bloques']} bloques "
         f"({W.n(D['ai_ha'], 3)} ha, campo AREA_HA)", "Buffer de 500 m alrededor de cada bloque, recortado por el límite "
         "de la microcuenca (capa «AI aprobado 2»)", "Vigente; base del diagnóstico social (E2) y del riesgo (E8)"],
        ["Zona de beneficio (impacto aguas abajo)", "Elementos expuestos aguas abajo que reciben el servicio regulado",
         "Conectividad hidrológica y reducción modelada del caudal pico y de sedimentos", "Por determinar (requiere el "
         "modelamiento hidrológico del E4 y un umbral de significancia)"],
    ], anchos=[0.17, 0.33, 0.30, 0.20], size=8)
    W.fuente(d, "NT N.° 001-2026-ANIN/DIME/SESDI (14/10/2026); capas bloques_v6, area_influencia_v6 y ambito_distritos "
                "(IN_Piura_area_influencia_bloques_v6.gpkg); Matrices E2 B4, hoja 08 (catálogo V8).")
    W.h(d, "3. Área de intervención por provincia")
    filas = []
    dem = {}
    for r in D["demanda_distrito"]:
        k = r["Provincia"]
        x = dem.setdefault(k, [0, 0, 0, 0])
        x[0] += r["Bloques V8"]; x[1] += r["Superficie V7 (ha)"]; x[2] += r["Descuento V8 (ha)"]; x[3] += r["Superficie neta V8 (ha)"]
    for pv in D["por_provincia_v6"]:
        pt = prov_tit(pv["prov"])
        x = dem[pt]
        filas.append([pt, str(pv["dist"]), str(pv["micro"]), str(pv["bloques"]), W.n(pv["ha"], 3), W.n(x[2], 3),
                      W.n(x[3], 3), str(pv["ai_n"]), W.n(pv["ai_ha"], 3)])
    tot = ["Total", "15", str(D["microcuencas_bloques"]), str(D["bloques_n"]), W.n(D["bloques_ha"], 3),
           W.n(D["demanda_total"]["Descuento V8 (ha)"], 3), W.n(D["demanda_total"]["Superficie neta V8 (ha)"], 3),
           str(D["ai_n"]), W.n(D["ai_ha"], 3)]
    W.tabla(d, ["Provincia", "Distritos", "Microcuencas", "Bloques", "Bloques V6/V7 (ha)", "Descuento V8 (ha)",
                "Neta V8 (ha)", "Polígonos AI", "Área de influencia (ha)"], filas, total=tot,
            anchos=[0.13, 0.10, 0.12, 0.09, 0.12, 0.11, 0.11, 0.10, 0.12], size=7.5)
    W.fuente(d, "Geopaquete de bloques y AI V6 (equipo SIG); Matrices E2 B4, hoja 08. Microcuencas contadas en cada "
                "provincia: las compartidas se cuentan en ambas, por eso la suma por provincia supera el total de 26.")
    W.h(d, "4. Criterios técnicos de la delimitación")
    W.vinetas(d, [
        ("Enfoque de zona de origen. ", "El proyecto interviene en la cuenca alta, donde se originan los movimientos en masa, "
         "la erosión y la escorrentía que generan inundaciones aguas abajo. La unidad productora es el ecosistema; los bloques "
         "son las áreas donde se aplican las medidas."),
        ("Área de intervención. ", "Resulta del proceso de meso y microlocalización descrito en el paquete 02. El catálogo V8 "
         f"(09/10/2026) mantiene los {D['bloques_n']} bloques y descuenta {W.n(D['demanda_total']['Descuento V8 (ha)'], 3)} ha "
         "superpuestas con derechos mineros en trámite (bloques 8, 71, 22 y M17B1) y con un sitio arqueológico (bloque 5). "
         "Los polígonos V8 aún no se recortan en la capa SIG."),
        ("Área de influencia directa. ", "Buffer de 500 m alrededor de cada bloque, recortado por el límite de la "
         "microcuenca. Se excluyen AI_74 y AI_75 por corresponder a bloques retirados; por ello el AI vigente suma "
         f"{W.n(D['ai_ha'], 3)} ha y no 17,295.547 ha (total del shapefile original). El área geométrica recalculada es "
         f"{W.n(D['ai_geom_ha'], 3)} ha; los cuadros usan el campo AREA_HA."),
        ("Población del área de influencia. ", "Centros poblados del CPV 2017 (INEI) asociados a cada bloque y localidades "
         "sin código INEI identificadas en campo; proyección a 2026 con tasas intercensales (paquete 03)."),
        ("Ámbitos del análisis de riesgo (E8). ", "El EVAR evalúa el peligro por movimientos en masa en los bloques "
         "(12,270.24 ha) y en 85 áreas de aporte (7,538.37 ha), y la vulnerabilidad y el riesgo en las 117 áreas de "
         "influencia (17,132.45 ha analíticas). Las superficies de los tres ámbitos no se suman."),
        ("Zona de beneficio. ", "Se delimitará por conectividad hidrológica aguas abajo de cada bloque y reducción modelada "
         "del caudal pico y de sedimentos en puntos de control; solo se incluirán elementos que superen un umbral de "
         "significancia fijado previamente. Es el ámbito que conecta el Perfil con las intervenciones del Programa en la zona "
         "de impacto."),
    ])
    W.h(d, "5. Distribución por distrito")
    filas = []
    for r in D["demanda_distrito"]:
        aid = next((x for x in D["por_distrito_v6"] if x["dist"].replace("Í", "I") ==
                    r["Distrito"].upper().replace("Í", "I").replace("Ó", "O").replace("Á", "A")), None)
        filas.append([r["Provincia"], r["Distrito"], str(r["Bloques V8"]), W.n(r["Superficie V7 (ha)"], 3),
                      W.n(r["Superficie neta V8 (ha)"], 3), W.n(aid["ai_ha"], 3) if aid else "Por verificar"])
    W.tabla(d, ["Provincia", "Distrito", "Bloques", "V6/V7 (ha)", "Neta V8 (ha)", "AI (ha)"], filas,
            total=["Total", "15 distritos", str(D["bloques_n"]), W.n(D["bloques_ha"], 3),
                   W.n(D["demanda_total"]["Superficie neta V8 (ha)"], 3), W.n(D["ai_ha"], 3)],
            anchos=[0.17, 0.27, 0.10, 0.16, 0.16, 0.14], size=8)
    W.fuente(d, "Matrices E2 B4, hoja 08; geopaquete de AI V6.")
    W.h(d, "6. Pendientes")
    W.vinetas(d, ["Recortar los polígonos de los cinco bloques con descuento V8 y recalcular sus atributos.",
                  "Delimitar la zona de beneficio aguas abajo con el modelamiento hidrológico (E4) y fijar el umbral de "
                  "significancia antes del procesamiento.",
                  "Fijar la superficie final de intervención tras el tamizaje predial (actas de libre disponibilidad).",
                  "Actualizar el área de estudio con la envolvente de los ámbitos anteriores."])
    # anexos
    x = excel_01(D, repo, drv, an / "Matriz_Ambito_Intervencion_e_Influencia_IN_Piura.xlsx")
    copiar(drv / "NT_Redelimitacion_AreaInfluencia_IN_Piura_rev_hscm.docx", an,
           "NT_001-2026_Redelimitacion_Area_Influencia_IN_Piura.docx")
    copiar(repo / "mapas/salidas/IN_Piura_Bloques_AI_por_Distrito_V6.xlsx", an)
    copiar(repo / "mapas/salidas/IN_Piura_Mapa_Area_Influencia_Bloques_V6.pdf", an)
    kz = out.parent / CARPETAS["06"] / "02_KMZ"
    for k in ("01_Bloques_V6_117.kmz", "02_Area_Influencia_V6.kmz", "03_Ambito_15_Distritos.kmz"):
        copiar(kz / k, an / "KMZ")
    lista_anexos(d, [
        ("01_Nota_Tecnica_Delimitacion_Ambito.docx", "Esta nota", FECHA),
        ("Anexos/Matriz_Ambito_Intervencion_e_Influencia_IN_Piura.xlsx", "Bloques V8 y AI por bloque, distrito y provincia (fórmulas)", "Generado al corte"),
        ("Anexos/NT_001-2026_Redelimitacion_Area_Influencia_IN_Piura.docx", "Opinión técnica sobre el AI y propuesta de zona de beneficio", "14/10/2026"),
        ("Anexos/IN_Piura_Bloques_AI_por_Distrito_V6.xlsx", "Bloques V6 y AI por distrito (fórmulas COUNTIFS/SUMIFS)", "Base V6"),
        ("Anexos/IN_Piura_Mapa_Area_Influencia_Bloques_V6.pdf", "Mapa general y por provincia de bloques y AI (A3)", "Base V6"),
        ("Anexos/KMZ/*.kmz", "Bloques, AI y ámbito distrital para Google Earth", "Base V6"),
    ])
    W.p(d, "Las capas en formato nativo (SHP, GPKG) se remiten en el paquete 06.", italic=True, size=8.5)
    W.firma(d)
    d.save(out / "01_Nota_Tecnica_Delimitacion_Ambito.docx")


def excel_01(D, repo, drv, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    rows = datos._rows(drv / "Matrices_E2_IN_Piura_B4.xlsx", "07_Demanda_Bloques")
    i = next(k for k, r in enumerate(rows) if r and r[0] == "Provincia")
    hdr = rows[i]
    bl = [dict(zip(hdr, r)) for r in rows[i + 1:] if r and r[0] in datos.PROVS]
    ai = gpd.read_file(repo / "datos/gis/IN_Piura_area_influencia_bloques_v6.gpkg", layer="area_influencia_v6")
    # Hoja AI
    ws2 = wb.active; ws2.title = "AI_Poligonos"
    f = X.encabezado(ws2, "ÁREA DE INFLUENCIA DIRECTA — POLÍGONOS (AI aprobado 2, sin AI_74 ni AI_75)",
                     "Fuente: capa AI_aprobado_2 (buffer de 500 m recortado por microcuenca); geopaquete IN Piura V6.", 7)
    import unicodedata
    def _k(t):
        return unicodedata.normalize("NFKD", str(t)).encode("ascii", "ignore").decode().upper()
    nom_prov = {_k(r["Provincia"]): r["Provincia"] for r in bl}
    nom_dist = {_k(r["Distrito"]): r["Distrito"] for r in bl}
    filas = [[r.CODIGO, str(r.BLOQUE), r.MICROC, nom_prov.get(_k(r.NOMBPROV), r.NOMBPROV),
              nom_dist.get(_k(r.NOMBDIST), r.NOMBDIST), round(r.AREA_HA, 3), round(r.AREA_GEOM_HA, 3)]
             for r in ai.itertuples()]
    last_ai = X.tabla(ws2, f, ["Código AI", "Bloque", "Microcuenca", "Provincia", "Distrito", "Área AREA_HA (ha)",
                              "Área geométrica (ha)"], filas, anchos=[14, 10, 14, 14, 24, 16, 16], total_cols=[6, 7],
                     fmt={6: "#,##0.000", 7: "#,##0.000"})
    n_ai = f + len(filas)
    # Hoja bloques
    ws = wb.create_sheet("Bloques_V8", 0)
    f = X.encabezado(ws, "ÁREA DE INTERVENCIÓN — 117 BLOQUES (CATÁLOGO V8) Y SU ÁREA DE INFLUENCIA",
                     "Fuente: catálogo Bloques_V8_Normalizado (09/10/2026) vía Matrices E2 B4 (hoja 07); AI por SUMIF sobre la "
                     "hoja AI_Poligonos. Centroides en UTM 17S.", 12)
    filas = []
    for k, r in enumerate(bl):
        fila = f + 1 + k
        filas.append([r["Provincia"], r["Distrito"], str(r["Bloque (catálogo)"]), r["Microcuenca"], r["Superficie V7 (ha)"],
                      r["Descuento V8 (ha)"], f"=E{fila}-F{fila}",
                      f"=SUMIF(AI_Poligonos!$B${9}:$B${n_ai},C{fila},AI_Poligonos!$F${9}:$F${n_ai})",
                      f"=COUNTIF(AI_Poligonos!$B${9}:$B${n_ai},C{fila})",
                      r["Este (m) UTM 17S"], r["Norte (m) UTM 17S"], r["Observación V8"] or ""])
    X.tabla(ws, f, ["Provincia", "Distrito", "Bloque", "Microcuenca", "Superficie V6/V7 (ha)", "Descuento V8 (ha)",
                    "Superficie neta V8 (ha)", "Área de influencia (ha)", "Polígonos AI", "Este (m)", "Norte (m)",
                    "Observación V8"], filas, anchos=[13, 22, 10, 14, 14, 13, 14, 14, 10, 12, 12, 40],
            total_cols=[5, 6, 7, 8, 9], fmt={5: "#,##0.000", 6: "#,##0.000", 7: "#,##0.000", 8: "#,##0.000",
                                              10: "#,##0", 11: "#,##0"})
    n_bl = f + len(filas)
    # Resumen
    ws3 = wb.create_sheet("Resumen_Provincia", 0)
    f = X.encabezado(ws3, "RESUMEN POR PROVINCIA — ÁREA DE INTERVENCIÓN Y ÁREA DE INFLUENCIA",
                     "Fórmulas COUNTIFS/SUMIFS sobre las hojas Bloques_V8 y AI_Poligonos.", 7)
    filas = []
    for k, pv in enumerate(datos.PROVS):
        fila = f + 1 + k
        filas.append([pv, f'=COUNTIF(Bloques_V8!$A$9:$A${n_bl},A{fila})',
                      f'=SUMIF(Bloques_V8!$A$9:$A${n_bl},A{fila},Bloques_V8!$E$9:$E${n_bl})',
                      f'=SUMIF(Bloques_V8!$A$9:$A${n_bl},A{fila},Bloques_V8!$F$9:$F${n_bl})',
                      f'=SUMIF(Bloques_V8!$A$9:$A${n_bl},A{fila},Bloques_V8!$G$9:$G${n_bl})',
                      f'=COUNTIF(AI_Poligonos!$D$9:$D${n_ai},A{fila})',
                      f'=SUMIF(AI_Poligonos!$D$9:$D${n_ai},A{fila},AI_Poligonos!$F$9:$F${n_ai})'])
    last = X.tabla(ws3, f, ["Provincia", "Bloques", "Superficie V6/V7 (ha)", "Descuento V8 (ha)", "Neta V8 (ha)",
                           "Polígonos AI", "Área de influencia (ha)"], filas, anchos=[16, 10, 18, 16, 16, 12, 20],
                   total_cols=[2, 3, 4, 5, 6, 7], fmt={3: "#,##0.000", 4: "#,##0.000", 5: "#,##0.000", 7: "#,##0.000"})
    X.notas(ws3, last + 2, [
        f"Controles esperados: 117 bloques; {W.n(D['bloques_ha'], 3)} ha V6/V7; {W.n(D['demanda_total']['Superficie neta V8 (ha)'], 2)} ha netas V8; "
        f"{D['ai_n']} polígonos AI; {W.n(D['ai_ha'], 3)} ha de AI.",
        "Área de influencia directa: buffer de 500 m alrededor de cada bloque recortado por la microcuenca (NT N.° 001-2026-ANIN/DIME/SESDI).",
        "Bloques retirados del catálogo (no figuran): 1, 7, 25, 29, 32, 33, 46, 48, 68, 74, 75 y M18B5.",
        "Zona de beneficio aguas abajo: Por determinar (modelamiento hidrológico E4)."])
    wb.save(path)
    return path


# ======================================================================================== 02
CRITERIOS = [
    ("1. Mesolocalización", "Peligro por movimientos en masa (PMM)", "MCA con ponderación de Saaty: precipitación (10), pendiente (8), geología (7), cobertura vegetal (5), rugosidad (3)", "SENAMHI/PISCO, DEM, INGEMMET, Sentinel-2", "Mapa PMM"),
    ("1. Mesolocalización", "Potencial hídrico de inundación (PHI)", "Escorrentía (10), precipitación (9), pendiente (8), NDMI (6), cobertura vegetal (5)", "DEM, SENAMHI, Sentinel-2", "Mapa PHI"),
    ("1. Mesolocalización", "Erosión hídrica superficial (EHS)", "Ecuación RUSLE (R, K, L, S, C, P)", "SENAMHI, suelos, DEM, cobertura", "Mapa EHS"),
    ("1. Mesolocalización", "Peligro integrado (PI)", "PI = 0.4·PMM + 0.3·EPH + 0.3·PGI (MCA-AHP), 5 niveles; cruzado con cobertura (mediana MSAVI 2014-2024) y degradación (tendencia MSAVI, Theil-Sen)", "Sentinel-2 (nubosidad < 15 %)", "PI_MSAVI_ED"),
    ("2. Priorización de microcuencas", "Exclusión de zonas", "Áreas agrícolas, cuerpos de agua, superficies nivales, afloramientos rocosos, pendiente > 75 %, precipitación < 300 mm/año, altitud > 4,600 msnm", "MINAM, DEM, SENAMHI", "Superficie apta"),
    ("2. Priorización de microcuencas", "Clase de peligro", "Se retienen solo PI Medio, Alto y Muy alto", "Modelo PI", "Áreas elegibles"),
    ("2. Priorización de microcuencas", "Déficit hídrico futuro", "Descarte de microcuencas con déficit (salvo distritos priorizados)", "MPI-ESM1-2HR, SSP5-8.5, 2021-2040", "Microcuencas filtradas"),
    ("2. Priorización de microcuencas", "Ranking", "Suma zonal de peligro + proporción agrícola, normalizadas y promediadas", "Modelo PI; cobertura", "Orden de prioridad"),
    ("3. Microlocalización de bloques", "Digitalización", "Sobre píxeles PI Medio a Muy alto; prioridad a bloques con o colindantes a elementos expuestos (centros poblados, áreas agrícolas, vías)", "Modelo PI; INEI; MTC", "Bloques V#1"),
    ("3. Microlocalización de bloques", "Exclusión", "Restos arqueológicos (SIGDA); concesiones mineras en exploración o explotación (INGEMMET); franja de 10 m en líneas eléctricas (OSINERGMIN); áreas naturales protegidas (SERNANP)", "SIGDA-Min. Cultura, INGEMMET, OSINERGMIN, SERNANP", "Bloques V#2"),
    ("3. Microlocalización de bloques", "Idoneidad", "Accesibilidad (red vial MTC e imágenes); pendiente ≤ 75 %; distritos priorizados por pobreza; déficit hídrico futuro", "MTC, DEM, INEI, MPI-ESM1-2HR", "Bloques V#3"),
    ("3. Microlocalización de bloques", "Tamizaje predial", "Hasta 19 predios rurales por bloque; sin superposición con áreas agrícolas (ajuste de geometría o descarte)", "COFOPRI, GEORURAL-MIDAGRI", "Bloques filtrados"),
    ("4. Verificación de campo", "Social y legal", "Aceptación de titulares; condición jurídica del predio; titular identificado", "Fichas F-DT, F-DS, F-LA", "Bloques preliminares"),
    ("4. Verificación de campo", "Accesibilidad y riesgo social", "Acceso real; minería ilegal, pasivos, delincuencia u otros conflictos", "Fichas F-DT", "Bloques preliminares"),
    ("4. Verificación de campo", "Condición del terreno", "Evidencia de movimientos en masa; uso actual; afloramientos; suelos superficiales; cobertura forestal densa", "Fichas F-DT, fotografías, dron", "Codificación M#-B#"),
    ("5. Ajustes posteriores (V6 → V8)", "Retiro de bloques", "12 códigos retirados del catálogo del aplicativo (1, 7, 25, 29, 32, 33, 46, 48, 68, 74, 75, M18B5)", "Verificación de campo; tamizaje", "Catálogo V6 (117)"),
    ("5. Ajustes posteriores (V6 → V8)", "Descuento por superposición", "Derechos mineros en trámite (bloques 8, 71, 22, M17B1) y Sitio Arqueológico Juan Velasco Alvarado 4 (bloque 5)", "INGEMMET, SIGDA", "Catálogo V8 (09/10/2026)"),
]


def paquete_02(D, repo, drv, out):
    an = out / "Anexos"
    d = doc_base("NOTA TÉCNICA", "Criterios de selección e identificación de los ámbitos de intervención",
                 "Información de avance del Perfil para su articulación con el Programa de Inversión del Río Piura – "
                 "Componente 2: criterios de selección")
    W.h(d, "1. Enfoque")
    W.p(d, "El proyecto recupera ecosistemas degradados que brindan el servicio de regulación de riesgos naturales "
           "(indicador de brecha de la R.M. N.° 00213-2024-MINAM). Los ámbitos de intervención se ubican en la zona de origen "
           "de los peligros, donde la infraestructura natural atenúa los factores condicionantes y desencadenantes de "
           "movimientos en masa, erosión y escorrentía. La selección sigue un proceso secuencial de gabinete y campo "
           "(cuenca → microcuenca → bloque) establecido en la Nota Técnica de Meso y Microlocalización de las intervenciones "
           "con Infraestructura Natural (ANIN-DIME, 2023) y aplicado al ámbito en el Modelamiento Geoespacial de "
           "Mesolocalización – Perfil IN Piura (400012-ANIN001-000-XX-MD-ZZ-000001, ANIN-SESDI, junio 2025).")
    W.h(d, "2. Criterios aplicados por etapa")
    W.tabla(d, ["Etapa", "Criterio", "Parámetro / umbral", "Fuente de datos", "Producto"],
            [list(c) for c in CRITERIOS], anchos=[0.15, 0.17, 0.36, 0.18, 0.14], size=7.5)
    W.fuente(d, "Criterios_Seleccion_Bloques_IN_Piura_Rev_hscm.docx (DIME-SESDI, set. 2026); NT N.° 001-2026-ANIN/DIME/SESDI; "
                "Mapa conceptual de criterios; catálogo V8 (09/10/2026).")
    W.h(d, "3. Criterios territoriales")
    W.vinetas(d, [
        ("Ámbito político: ", "3 provincias (Morropón, Huancabamba y Ayabaca) y 15 distritos de la cuenca alta del río Piura."),
        ("Unidad hidrográfica: ", f"{D['microcuencas_bloques']} microcuencas (nivel 12) que contienen los bloques, priorizadas por "
         "peligro, proporción agrícola y déficit hídrico futuro."),
        ("Zona de origen del peligro: ", "laderas y cabeceras con PI medio a muy alto, donde se generan los procesos que "
         "afectan a la zona baja (valle del Alto Piura)."),
        ("Distritos priorizados: ", "presencia en distritos priorizados por nivel de pobreza como criterio de idoneidad."),
        ("Exposición: ", "prioridad a bloques con o colindantes a centros poblados, áreas agrícolas e infraestructura vial."),
    ])
    W.h(d, "4. Resultado y verificación posterior")
    dt = D["demanda_total"]
    W.p(d, f"El proceso definió {D['bloques_n']} bloques preliminares con {W.n(D['bloques_ha'], 3)} ha en 3 provincias y 15 "
           f"distritos. Con el catálogo V8 la superficie neta es {W.n(dt['Superficie neta V8 (ha)'], 2)} ha. Por clase de "
           f"peligro integrado (índice medio del bloque): {dt['Bloques PI muy alto']} bloques muy alto, {dt['Bloques PI alto']} "
           f"alto, {dt['Bloques PI medio']} medio y {dt['Bloques PI bajo']} bajo; {W.n(dt['ha en PI alto y muy alto'], 2)} ha "
           f"({W.pct(dt['% PI alto y muy alto (sobre V7)'])}) están en PI alto o muy alto.")
    s = D["serrn_total"]
    W.p(d, f"Como contraste con la capa oficial de áreas degradadas que prestan el servicio de regulación de riesgos "
           f"naturales (SERRN, niveles alto y muy alto), {W.n(s['area_serrn'], 2)} ha ({W.pct(s['pct'])}) de los bloques V7 se "
           f"superponen con dicha capa y {s['sin_superp']} bloques no presentan superposición. El indicador de brecha del "
           "proyecto se mide con el umbral MSAVI (paquete 04).")
    W.h(d, "5. Pertinencia para el Programa de Inversión del Río Piura")
    W.p(d, "Los criterios priorizan áreas con peligro medio a muy alto y con degradación en la cuenca alta, lo que asegura "
           "coherencia con los objetivos de reducción del riesgo de inundaciones del Programa. Las medidas de infraestructura "
           "natural actúan en la zona de origen y complementan las medidas estructurales en la zona de impacto; la "
           "articulación se concretará con la zona de beneficio aguas abajo (paquete 01, numeral 4).")
    W.h(d, "6. Carácter preliminar")
    W.vinetas(d, ["La delimitación final se ajusta con el tamizaje predial y la aceptación de los titulares (actas de libre "
                  "disponibilidad), en curso.",
                  "Cinco bloques de Huancabamba (4, 9, 35, 38 y 43) superan los 19 predios COFOPRI: se redelimitan o sectorizan.",
                  "60 lotes preliminares de sistemas de uso sostenible (256.97 ha) se suman al tamizaje."])
    pendientes(d, "02")
    xls = excel_02(D, drv, an / "Matriz_Criterios_Seleccion_IN_Piura.xlsx")
    for f, nm in [("Criterios_Seleccion_Bloques_IN_Piura_Rev_hscm.docx", None),
                  ("Resumen_Ejecutivo_Criterios_Seleccion_Bloques_IN_Piura.docx", None),
                  ("Mapa_Conceptual_Criterios_Seleccion_Bloques_IN_Piura.png", None),
                  ("400012-ANIN001-000-XX-MD-ZZ-000001_Modelamiento_Mesolocalizacion.pdf", None),
                  ("Interseccion_SERRN_AMA_Bloques_V7_IN_Piura.xlsx", None)]:
        copiar(drv / f, an, nm)
    lista_anexos(d, [
        ("02_Nota_Tecnica_Criterios_Seleccion.docx", "Esta nota", FECHA),
        ("Anexos/Matriz_Criterios_Seleccion_IN_Piura.xlsx", "Criterios por etapa, ajustes V8 y clasificación de los 117 bloques", "Generado al corte"),
        ("Anexos/Criterios_Seleccion_Bloques_IN_Piura_Rev_hscm.docx", "Sustento de criterios técnicos y metodológicos", "Set. 2026"),
        ("Anexos/Resumen_Ejecutivo_Criterios_Seleccion_Bloques_IN_Piura.docx", "Resumen ejecutivo", "Set. 2026"),
        ("Anexos/Mapa_Conceptual_Criterios_Seleccion_Bloques_IN_Piura.png", "Mapa conceptual del proceso de selección", "Set. 2026"),
        ("Anexos/400012-ANIN001-000-XX-MD-ZZ-000001_Modelamiento_Mesolocalizacion.pdf", "Modelamiento geoespacial de mesolocalización", "Jun. 2025"),
        ("Anexos/Interseccion_SERRN_AMA_Bloques_V7_IN_Piura.xlsx", "Superposición de bloques V7 con áreas SERRN alto/muy alto", "07/11/2026"),
    ])
    W.firma(d)
    d.save(out / "02_Nota_Tecnica_Criterios_Seleccion.docx")


def excel_02(D, drv, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook(); ws = wb.active; ws.title = "Criterios"
    f = X.encabezado(ws, "MATRIZ DE CRITERIOS DE SELECCIÓN DE BLOQUES DE INTERVENCIÓN",
                     "Fuente: Criterios_Seleccion_Bloques_IN_Piura_Rev_hscm.docx; NT de Meso y Microlocalización (ANIN-DIME, 2023); "
                     "Modelamiento 400012-ANIN001 (2025); catálogo V8.", 5)
    X.tabla(ws, f, ["Etapa", "Criterio", "Parámetro / umbral", "Fuente de datos", "Producto"], [list(c) for c in CRITERIOS],
            anchos=[26, 28, 70, 34, 22])
    rows = datos._rows(drv / "Matrices_E2_IN_Piura_B4.xlsx", "07_Demanda_Bloques")
    i = next(k for k, r in enumerate(rows) if r and r[0] == "Provincia")
    hdr = rows[i]
    bl = [dict(zip(hdr, r)) for r in rows[i + 1:] if r and r[0] in datos.PROVS]
    ws2 = wb.create_sheet("Bloques_Clasificacion")
    f = X.encabezado(ws2, "CLASIFICACIÓN DE LOS 117 BLOQUES POR CRITERIOS DE PELIGRO, DEGRADACIÓN Y CONDICIÓN",
                     "Fuente: Matrices E2 B4, hoja 07 (catálogo V8; MSAVI 2024 y PI MCA-AHP del Diagnóstico Territorial v4; "
                     "umbral MSAVI de brecha 0.4976).", 14)
    filas = []
    for k, r in enumerate(bl):
        fila = f + 1 + k
        filas.append([r["Provincia"], r["Distrito"], str(r["Bloque (catálogo)"]), r["Microcuenca"],
                      r["Superficie neta V8 (ha)"], r["MSAVI medio 2024"], r["Bajo el umbral"], r["PI: clase por índice"],
                      r["ha en PI alto y muy alto"], r["Exposición de la población"], r["Peligro declarado en campo (DT)"],
                      r["Condición socio-operativa"], r["Situación predial SIG"], r["Observación V8"] or ""])
    last = X.tabla(ws2, f, ["Provincia", "Distrito", "Bloque", "Microcuenca", "Neta V8 (ha)", "MSAVI medio 2024",
                            "Bajo umbral MSAVI", "Clase PI", "ha PI alto y muy alto", "Exposición", "Peligro declarado (campo)",
                            "Condición socio-operativa", "Situación predial SIG", "Observación V8"], filas,
                   anchos=[13, 20, 10, 14, 12, 11, 10, 10, 12, 12, 30, 12, 20, 36], total_cols=[5, 9],
                   fmt={5: "#,##0.000", 6: "0.0000", 9: "#,##0.00"})
    X.notas(ws2, last + 2, ["Conteos de control: use COUNTIF sobre las columnas G (Sí/No) y H (Muy alto, Alto, Medio, Bajo).",
                            "Condición socio-operativa: clasificación del Diagnóstico Territorial y Social (A, B, C)."])
    ws3 = wb.create_sheet("Retirados_y_V8")
    f = X.encabezado(ws3, "AJUSTES DEL CATÁLOGO: BLOQUES RETIRADOS (V6) Y DESCUENTOS V8",
                     "Fuente: README de liberación de áreas (28/09/2026); catálogo Bloques_V8_Normalizado (09/10/2026).", 4)
    filas = [[c, "Retirado del catálogo vigente V6", "Verificación de campo / tamizaje", "No se costea ni se diagnostica"]
             for c in ["1", "7", "25", "29", "32", "33", "46", "48", "68", "74", "75", "M18B5"]]
    filas += [[str(r["Bloque (catálogo)"]), f"Descuento V8: {r['Descuento V8 (ha)']:.3f} ha", "Superposición INGEMMET / SIGDA",
               r["Observación V8"] or ""] for r in bl if r["Descuento V8 (ha)"]]
    X.tabla(ws3, f, ["Bloque", "Ajuste", "Motivo", "Observación"], filas, anchos=[12, 30, 32, 70])
    wb.save(path)


# ======================================================================================== 03
def paquete_03(D, repo, drv, out):
    an = out / "Anexos"
    d = doc_base("NOTA TÉCNICA", "Avance del diagnóstico de la población afectada",
                 "Información de avance del Perfil para su articulación con el Programa de Inversión del Río Piura – "
                 "Componente 3: población afectada")
    pt = D["pob_total"]
    W.h(d, "1. Estado del entregable")
    W.p(d, "El diagnóstico de la población afectada forma parte del Entregable 2 (E2), que se presenta en tres volúmenes "
           "provinciales (I Morropón, II Huancabamba, III Ayabaca) en borrador B4 del 07/11/2026, con su libro de matrices. "
           "El avance del E2 reportado al 03/11/2026 es 75 %.")
    W.h(d, "2. Metodología aplicada")
    W.vinetas(d, [
        ("Delimitación: ", "población del área de influencia directa (buffer de 500 m recortado por microcuenca) asociada a "
         "cada bloque; la población de la zona de impacto aguas abajo se tratará como beneficiaria indirecta cuando se "
         "delimite la zona de beneficio."),
        ("Fuente secundaria: ", "Censos Nacionales 2017 (INEI) por centro poblado, más localidades sin código INEI "
         "identificadas en campo (población no censal, por validar)."),
        ("Proyección: ", "tasas intercensales 2007-2017 distritales (o provincial si la distrital no está disponible) para "
         "2026 (año base), 2027 (inicio de ejecución), 2032 (inicio de O&M) y horizonte a 2046."),
        ("Tipos de población (Guía General, Gráfico 2.5): ", "por la búsqueda implícita del servicio de regulación de "
         "riesgos, referencia = demandante potencial = efectiva = objetivo. La demanda del servicio se expresa además en "
         "hectáreas de ecosistema degradado (indicador de brecha)."),
        ("Fuente primaria: ", "fichas F-DS-01 (caracterización del centro poblado), F-DS-02 (actores), F-DS-03 (entrevista a "
         "autoridades) y F-DS-06 (percepción del peligro y alerta temprana), aplicadas con el aplicativo IN Piura; muestra de "
         "población finita (e = 5 %, p = q = 0.5, 95 % de confianza) con ajuste por bloque."),
        ("Protección de datos: ", "actores y entrevistados codificados o agregados (Ley N.° 29733)."),
    ])
    W.h(d, "3. Resultados preliminares: población afectada")
    filas = []
    for r in D["pob_distrito"]:
        filas.append([r["Provincia"], r["Distrito"], W.n(r["Pob. 2017 CP con código INEI"], 0),
                      W.n(r["Pob. 2017 localidades sin código (no censal)"], 0), W.n(r["Pob. 2017 total"], 0),
                      f"{r['Tasa r'] * 100:.2f} %", W.n(r["2026 (base)"], 0), W.n(r["2032 (inicio O&M)"], 0)])
    W.tabla(d, ["Provincia", "Distrito", "CPV 2017 (CP con código)", "Sin código INEI", "Total 2017", "Tasa r",
                "2026", "2032"], filas,
            total=["Total", "15 distritos", W.n(pt["Pob. 2017 CP con código INEI"], 0),
                   W.n(pt["Pob. 2017 localidades sin código (no censal)"], 0), W.n(pt["Pob. 2017 total"], 0), "",
                   W.n(pt["2026 (base)"], 0), W.n(pt["2032 (inicio O&M)"], 0)],
            anchos=[0.15, 0.21, 0.12, 0.10, 0.10, 0.10, 0.11, 0.11], size=8)
    W.fuente(d, "Matrices E2 B4, hoja 05 (INEI CPV 2017; catálogo CentrosPoblados-INEI-Bloques_v5 CON SIN 02.10.26). "
                "Tasas provinciales en Morropón: distrital por validar.")
    W.p(d, f"La población del área de influencia es {W.n(pt['Pob. 2017 total'], 0)} habitantes en 2017 y "
           f"{W.n(pt['2026 (base)'], 0)} proyectados a 2026. El análisis de riesgo del E8 usa 21,672 habitantes sin duplicación "
           "en 106 centros poblados de referencia (62 interiores y 55 exteriores a las áreas de influencia), con un criterio "
           "distinto; ambas cifras no son comparables y se conciliarán en la versión final.")
    W.h(d, "4. Cobertura del trabajo de campo")
    cob = [r for r in D["cobertura"] if r["Provincia"] in datos.PROVS and r.get("Distrito")]
    def s(c):
        return sum((r.get(c) or 0) for r in cob)
    filas = []
    for pv in datos.PROVS:
        rr = [r for r in cob if r["Provincia"] == pv]
        filas.append([pv, str(sum(r["Bloques V8"] for r in rr)), str(sum(r["Bloques con diagnóstico social"] for r in rr)),
                      str(sum(r["Localidades asociadas"] for r in rr)), str(sum(r["Localidades con ficha social"] for r in rr)),
                      str(sum(r["F-DS-01 vigentes"] for r in rr)), str(sum(r["F-DS-02 vigentes"] for r in rr)),
                      str(sum(r["F-DS-03 vigentes (entrevistados)"] for r in rr)), str(sum(r["F-DS-06 vigentes"] for r in rr)),
                      str(sum(r["Total fichas vigentes"] for r in rr))])
    W.tabla(d, ["Provincia", "Bloques", "Bloques con DS", "Localidades", "Localidades con ficha", "F-DS-01", "F-DS-02",
                "F-DS-03", "F-DS-06", "Total fichas"], filas,
            total=["Total", str(s("Bloques V8")), str(s("Bloques con diagnóstico social")), str(s("Localidades asociadas")),
                   str(s("Localidades con ficha social")), str(s("F-DS-01 vigentes")), str(s("F-DS-02 vigentes")),
                   str(s("F-DS-03 vigentes (entrevistados)")), str(s("F-DS-06 vigentes")), str(s("Total fichas vigentes"))],
            anchos=[0.15, 0.09, 0.10, 0.11, 0.11, 0.09, 0.09, 0.08, 0.08, 0.10], size=7.5)
    W.fuente(d, "Matrices E2 B4, hoja 03 (consolidado DS del aplicativo, corte 07/11/2026 18:25).")
    W.h(d, "5. Hallazgos por provincia")
    for r in D["conclusiones_e2"]:
        if r["Tipo"] == "Conclusión":
            W.vinetas(d, [(f"{r['Provincia']}. ", r["Texto"])])
    W.h(d, "6. Vacíos de información")
    W.vinetas(d, [(f"{r['Provincia']}. ", r["Texto"]) for r in D["conclusiones_e2"] if r["Tipo"] != "Conclusión"])
    W.p(d, "Las fichas F-DS-04 (taller participativo), F-DS-05 (conflictos) y F-DS-07 (consentimiento previo) no se han "
           "aplicado; las variables de PEA, pobreza, discapacidad y autoidentificación se completarán con REDATAM-INEI.",
        italic=True)
    for f in ["Informe_E2_Vol_I_Morropon_IN_Piura_B4.docx", "Informe_E2_Vol_II_Huancabamba_IN_Piura_B4.docx",
              "Informe_E2_Vol_III_Ayabaca_IN_Piura_B4.docx", "Matrices_E2_IN_Piura_B4.xlsx",
              "CONTENIDO MINIMO DIAGNOSTICO DE LA POBLACION AFECTADA Y OAI.pdf",
              "Diagnostico_Poblacion_Involucrados_IN_Piura_rev2_hscm.docx"]:
        copiar(drv / f, an)
    copiar(repo / "CentrosPoblados-INEI-Bloques_v5(CON SIN) - 24.09.26.xlsx", an)
    copiar(repo / "Plantilla_Diagnostico_Social_IN_Piura_V4.xlsx", an / "Instrumentos")
    lista_anexos(d, [
        ("03_Nota_Tecnica_Poblacion_Afectada.docx", "Esta nota", FECHA),
        ("Anexos/Informe_E2_Vol_I/II/III_..._B4.docx", "Informes provinciales del E2 (población afectada, actores, riesgos sociales, situación predial)", "B4, 07/11/2026"),
        ("Anexos/Matrices_E2_IN_Piura_B4.xlsx", "Libro de 33 matrices del E2 (población, demanda, CP, actores, riesgos)", "B4, 07/11/2026"),
        ("Anexos/CentrosPoblados-INEI-Bloques_v5(CON SIN) - 24.09.26.xlsx", "Centros poblados por bloque, CPV 2017 y cálculo de muestra", "24/09/2026"),
        ("Anexos/CONTENIDO MINIMO ... OAI.pdf", "Contenido mínimo ANIN del diagnóstico de población afectada", "Referencia"),
        ("Anexos/Diagnostico_Poblacion_Involucrados_IN_Piura_rev2_hscm.docx", "Propuesta metodológica inicial del diagnóstico", "Ago. 2026"),
        ("Anexos/Instrumentos/Plantilla_Diagnostico_Social_IN_Piura_V4.xlsx", "Fichas F-DS-01 a F-DS-07", "V4"),
    ])
    W.p(d, "La capa de localidades del ámbito (puntos UTM 17S) se remite en el paquete 06 (SHP 19 y KMZ).", italic=True, size=8.5)
    W.firma(d)
    d.save(out / "03_Nota_Tecnica_Poblacion_Afectada.docx")


# ======================================================================================== 04
def paquete_04(D, repo, drv, out):
    an = out / "Anexos"
    d = doc_base("NOTA TÉCNICA", "Avance del diagnóstico de la Unidad Productora (ecosistemas)",
                 "Información de avance del Perfil para su articulación con el Programa de Inversión del Río Piura – "
                 "Componente 4: Unidades Productoras")
    W.h(d, "1. Identificación de la Unidad Productora")
    W.p(d, "En la tipología de recuperación de ecosistemas la Unidad Productora (UP) no es una infraestructura convencional: "
           "es el ecosistema que brinda el servicio de regulación de riesgos naturales en la cuenca alta del río Piura. Comprende "
           "los bosques estacionalmente secos de colina y montaña, matorrales, pastizales y transiciones andino-costeras, con sus "
           "suelos, cobertura vegetal y red de drenaje. Los 117 bloques son las áreas donde se aplicarán las medidas dentro de la UP.")
    W.vinetas(d, [("Servicio ecosistémico: ", "regulación de riesgos naturales (movimientos en masa, erosión, escorrentía que "
                   "genera inundaciones); servicios asociados: regulación hídrica, control de erosión y captura de carbono."),
                  ("Indicador de brecha: ", "porcentaje de superficie de ecosistemas degradados que brindan servicios "
                   "ecosistémicos que requieren recuperación (R.M. N.° 00213-2024-MINAM).")])
    W.h(d, "2. Ecosistemas que conforman la UP")
    eco = D["ecosistemas_v7"]
    tot = sum(e["Superficie (ha)"] for e in eco)
    W.tabla(d, ["Símbolo", "Ecosistema (MINAM 2018)", "Bloques", "Superficie (ha)", "% del total"],
            [[e["Símbolo"], e["Ecosistema"], str(e["N° de bloques"]), W.n(e["Superficie (ha)"], 2), W.pct(e["% del total"])]
             for e in eco], total=["", "Total 117 bloques V7", "", W.n(tot, 2), "100.0 %"],
            anchos=[0.12, 0.46, 0.12, 0.16, 0.14], size=8.5)
    W.fuente(d, "E6_Matriz_Ecosistemas_Bloques_V7.xlsx (Mapa Nacional de Ecosistemas, MINAM 2018). Un bloque puede contener "
                "varios ecosistemas; el ecosistema de cada bloque se confirma en campo.")
    W.h(d, "3. Diagnóstico de la condición (degradación)")
    lb = {r["Provincia"]: r for r in D["linea_base"]}
    filas = [[pv, str(lb[pv]["Bloques"]), W.n(lb[pv]["Área V7 (ha)"], 3), W.n(lb[pv]["Bajo umbral MSAVI, V7 (ha)"], 3),
              W.pct(lb[pv]["% del área V7"])] for pv in datos.PROVS]
    t = lb["TOTAL · línea base"]
    W.tabla(d, ["Provincia", "Bloques", "Área V7 (ha)", "Bajo umbral MSAVI (ha)", "% del área"], filas,
            total=["Total (línea base)", str(t["Bloques"]), W.n(t["Área V7 (ha)"], 3), W.n(t["Bajo umbral MSAVI, V7 (ha)"], 3),
                   W.pct(t["% del área V7"])], anchos=[0.26, 0.12, 0.20, 0.24, 0.18], size=8.5)
    W.fuente(d, "Linea_Base_Brecha_E6_IN_Piura.xlsx (línea base adoptada el 09/11/2026: píxeles MSAVI 2024 bajo el umbral "
                f"0.4976). Estimado sobre V8: {W.n(t['Bajo umbral en V8, estimado (ha)'], 2)} ha ({W.pct(t['% de V8'])}).")
    di = D["delim_interna"]
    W.vinetas(d, [
        ("Restauración activa: ", "5,503.41 ha con MSAVI ≤ 0.3813 (clases DN1-DN2) y 2,838.41 ha de enriquecimiento (DN3) "
         "dentro de la línea base (E6)."),
        ("Delimitación interna (NDVI / coberturas S2DR4): ", f"restauración activa {W.n(di['Restauración activa'][0], 2)} ha, "
         f"enriquecimiento {W.n(di['Enriquecimiento'][0], 2)} ha, conservación {W.n(di['Conservación'][0], 2)} ha y "
         f"tamizaje agrícola {W.n(di['Tamizaje agrícola'][0], 2)} ha; no es la métrica del indicador de brecha."),
        ("Índice de condición (E6): ", "53 bloques no degradados, 44 leves, 14 moderados y 6 severos (20 bloques moderados o "
         "severos con 1,645.00 ha); califica la condición relativa y no reemplaza al MSAVI."),
        ("Concentración: ", "la degradación se concentra en el bosque estacionalmente seco del piso bajo de Morropón (Buenos "
         "Aires, Salitral, Chulucanas y San Juan de Bigote superan el 97 % bajo umbral)."),
        ("Composición florística: ", "69 especies nativas identificadas; especies clave del bosque seco: faique, pasallo, "
         "hualtaco, palo santo, charán, overo, ceibo y sapote. Oferta de plantones declarada: 21,000 (vivero Hualtacal); la "
         "brecha de viveros se cuantifica en la fase 6 del E6."),
    ])
    W.h(d, "4. Peligro y riesgo en la UP")
    dt = D["demanda_total"]
    W.vinetas(d, [
        ("Peligro integrado (MCA-AHP): ", f"{W.n(dt['ha en PI alto y muy alto'], 2)} ha ({W.pct(dt['% PI alto y muy alto (sobre V7)'])}) "
         "en PI alto o muy alto; ningún bloque en clase muy baja."),
        ("Peligro por movimientos en masa (EVAR, E8): ", "categorías alta y muy alta en el 93.88 % de los bloques y el 96.88 % "
         "de las áreas de aporte; 49.94 % en las áreas de influencia."),
        ("Vulnerabilidad y riesgo (EVAR): ", "102 áreas de influencia con vulnerabilidad media y 15 alta; riesgo medio "
         "9,943.14 ha, alto 7,132.39 ha y muy alto 56.92 ha (AI_73 Chococa y AI_41 Cambruran, Huarmaca). El EVAR no valora "
         "daños ni pérdidas: los beneficios por daños evitados están Por determinar."),
        ("SERRN: ", f"{W.n(D['serrn_total']['area_serrn'], 2)} ha de los bloques V7 se superponen con áreas degradadas de nivel alto "
         "y muy alto que prestan el servicio (capa SERRN)."),
        ("Erosión y cárcavas: ", "55 cárcavas registradas en gabinete (23.08 km); 85 bloques sin registro, lo que no prueba "
         "ausencia."),
    ])
    W.h(d, "5. Suelos y clima")
    W.vinetas(d, [
        ("Estudio de Suelos E5 (informe final, 01/11/2026): ", "74 calicatas, 152 muestras, 13 pruebas de infiltración y 7 "
         "unidades de suelo; 87.9 % del área estudiada es tierra de protección; la infiltración es lenta en cinco de las siete "
         "unidades, lo que restringe las zanjas de infiltración."),
        ("Clima: ", "semiárido, 544 mm/año e índice de aridez 0.34 (PISCOp 1981-2025); lluvias de El Niño de 2 a 6 veces la media."),
    ])
    W.h(d, "6. Pendientes")
    W.vinetas(d, ["Validación ciega de 567 puntos y tendencia TAU (fase 7 del E6); compuesto de época de lluvias 2027.",
                  "Recorte V8 de polígonos y recálculo del MSAVI en cinco bloques.",
                  "Modelamiento hidrológico (E4: lluvias de diseño, IDF, HEC-HMS) para el balance de oferta y demanda del "
                  "servicio y la zona de beneficio.",
                  "Escenarios de cambio climático en el análisis de riesgo (el EVAR usa un desencadenante fijo)."])
    pendientes(d, "04")
    for f in ["E6_Informe_Diagnostico_Ecosistema_IN_Piura_Rev_hscm.docx", "E6_Matriz_Ecosistemas_Bloques_V7.xlsx",
              "Linea_Base_Brecha_E6_IN_Piura.xlsx", "Delimitacion_Interna_Bloques_E6_IN_Piura.xlsx",
              "Informe_EVAR_Piura_07Nov.docx", "Informe_Final_Estudio_Suelos_IN_Piura.docx",
              "Interseccion_SERRN_AMA_Bloques_V7_IN_Piura.xlsx", "Arboles_Problemas_IN_Piura.pptx"]:
        copiar(drv / f, an)
    for f in ["mapas/salidas/IN_Piura_Ecosistemas_por_Bloque_V6.xlsx", "mapas/salidas/IN_Piura_Mapa_Ecosistemas_Bloques_V6.pdf",
              "AREAS_MSAVI_BLOQUES_V6_REV_HSCM.xlsx", "Caracterización de cárcavas.xlsx"]:
        copiar(repo / f, an)
    copiar(repo / "Plantilla_DT_Campo_Check_Validada_V5.xlsx", an / "Instrumentos")
    lista_anexos(d, [
        ("04_Nota_Tecnica_Unidad_Productora.docx", "Esta nota", FECHA),
        ("Anexos/E6_Informe_Diagnostico_Ecosistema_IN_Piura_Rev_hscm.docx", "E6: diagnóstico del ecosistema y oferta de material vegetal", "10/11/2026"),
        ("Anexos/E6_Matriz_Ecosistemas_Bloques_V7.xlsx", "Ecosistemas por bloque V7", "10/11/2026"),
        ("Anexos/Linea_Base_Brecha_E6_IN_Piura.xlsx", "Línea base del indicador de brecha (MSAVI)", "09/11/2026"),
        ("Anexos/Delimitacion_Interna_Bloques_E6_IN_Piura.xlsx", "Restauración activa, enriquecimiento y conservación por bloque", "08/11/2026"),
        ("Anexos/Informe_EVAR_Piura_07Nov.docx", "E8: evaluación del riesgo por movimientos en masa", "07/11/2026"),
        ("Anexos/Informe_Final_Estudio_Suelos_IN_Piura.docx", "E5: estudio de suelos, informe final", "01/11/2026"),
        ("Anexos/Interseccion_SERRN_AMA_Bloques_V7_IN_Piura.xlsx", "Bloques V7 × áreas SERRN alto/muy alto", "07/11/2026"),
        ("Anexos/Arboles_Problemas_IN_Piura.pptx", "Árbol de problemas unificado v1.0 (E7)", "30/10/2026"),
        ("Anexos/IN_Piura_Ecosistemas_por_Bloque_V6.xlsx y mapa PDF", "Ecosistemas por bloque V6 y mapa temático", "Base V6"),
        ("Anexos/AREAS_MSAVI_BLOQUES_V6_REV_HSCM.xlsx", "Áreas MSAVI por clase DN y bloque", "Base V6"),
        ("Anexos/Caracterización de cárcavas.xlsx", "Inventario de cárcavas de gabinete", "Set. 2026"),
        ("Anexos/Instrumentos/Plantilla_DT_Campo_Check_Validada_V5.xlsx", "Fichas de diagnóstico territorial F-DT", "V5"),
    ])
    W.firma(d)
    d.save(out / "04_Nota_Tecnica_Unidad_Productora.docx")


# ======================================================================================== 05
def paquete_05(D, repo, drv, out):
    an = out / "Anexos"
    d = doc_base("NOTA TÉCNICA", "Avance del diagnóstico de involucrados", "Información de avance del Perfil para su "
                 "articulación con el Programa de Inversión del Río Piura – Componente 5: involucrados", horizontal=True)
    W.h(d, "1. Metodología")
    W.vinetas(d, [
        "Mapeo de actores con la ficha F-DS-02 (tipo, rol frente al proyecto, influencia, interés, posición y nivel territorial) "
        "y entrevistas a autoridades F-DS-03; codificación sin nombres (Ley N.° 29733).",
        "Matriz de población afectada y otros involucrados según la Guía General (Tabla 1.7) y el Anexo N.° 07 de la Directiva "
        "N.° 001-2019-EF/63.01 (numeral 2.1.4).",
        "Matriz de riesgos sociales con probabilidad × impacto y medidas de prevención (hoja 15 de las Matrices E2).",
    ])
    W.h(d, "2. Actores mapeados por provincia y tipo")
    tot = [r for r in D["actores"] if r["Provincia"] == "Total ámbito"]
    W.tabla(d, ["Tipo de actor", "N.° de actores", "Cooperantes", "Neutrales", "Oponentes", "Sin posición",
                "Influencia alta", "Interés alto", "Gestionar de cerca"],
            [[r["Tipo de actor"]] + [str(r[c]) for c in ("N.° de actores", "Cooperantes", "Neutrales", "Oponentes",
                                                        "Sin posición", "Influencia alta", "Interés alto", "Gestionar de cerca")]
             for r in tot if r["Tipo de actor"] != "Todos los tipos"],
            total=["Todos los tipos"] + [str(next(r for r in tot if r["Tipo de actor"] == "Todos los tipos")[c]) for c in
                                        ("N.° de actores", "Cooperantes", "Neutrales", "Oponentes", "Sin posición",
                                         "Influencia alta", "Interés alto", "Gestionar de cerca")],
            anchos=[0.24, 0.10, 0.10, 0.09, 0.09, 0.09, 0.10, 0.09, 0.10], size=8)
    pv = {r["Provincia"]: r["N.° de actores"] for r in D["actores"] if r["Tipo de actor"] == "Todos los tipos"}
    W.fuente(d, f"Matrices E2 B4, hoja 13 (F-DS-02, corte 07/11/2026). Por provincia: Morropón {pv.get('Morropón')}, "
                f"Huancabamba {pv.get('Huancabamba')}, Ayabaca {pv.get('Ayabaca')}. La ausencia de oponentes es provisional "
                "hasta aplicar F-DS-05 (conflictos) y F-DS-07 (consentimiento previo).")
    W.h(d, "3. Matriz de involucrados (características, intereses y participación)")
    filas = []
    for r in D["involucrados"]:
        g = r["Grupo involucrado"] or ""
        if "ANIN" in g:
            g = "ANIN (unidad formuladora y ejecutora)"
        filas.append([r["Provincia"], g, r["Situación negativa percibida"], r["Intereses o expectativas"],
                      r["Estrategias del proyecto"], r["Acuerdos y compromisos"], r["Actitud (cooperante / oponente)"]])
    W.tabla(d, ["Provincia", "Grupo involucrado", "Problema percibido", "Intereses / necesidades", "Estrategia del proyecto",
                "Participación / compromisos", "Actitud"], filas,
            anchos=[0.08, 0.17, 0.17, 0.14, 0.17, 0.15, 0.12], size=7, alinear_num=False)
    W.fuente(d, "Matrices E2 B4, hoja 14 (Guía General, Tabla 1.7), elaborada con F-DS-02, F-DS-03 y F-DS-06. Al corte, los "
                "acuerdos y compromisos son propuestas: no hay actas suscritas con los grupos.")
    W.h(d, "4. Hallazgos para la articulación con el Programa")
    W.vinetas(d, [
        "No existe plataforma de concertación en ninguna provincia; el teniente gobernador, la ronda campesina, el agente "
        "municipal y, en algunos distritos, la JASS o la comunidad campesina concentran la decisión local.",
        "Aceptación alta pero condicionada: en Morropón 63 de 69 entrevistados F-DS-03 están de acuerdo y 5 condicionan; en "
        "Huancabamba 21 de 23 de acuerdo y 2 condicionan (consulta previa; zanjas y ganado). Ayabaca no tiene F-DS-03.",
        "Riesgos sociales principales: extracción comercial de palo santo y hualtaco, pastoreo en áreas a clausurar, "
        "oposición a la minería (Río Blanco) que exige deslindar el proyecto de actividades extractivas, conflictos de linderos "
        "(La Cría – C.C. San Andrés de Culcas) y desconfianza por promesas incumplidas.",
        "Actores institucionales comunes con el Programa: Gobierno Regional de Piura (reforestaciones y viveros Hualtacal y "
        "Maraypampa; ACR Bosques Secos de Salitral-Huarmaca), SERFOR/ATFFS, ANA y comisiones de regantes, municipalidades "
        "distritales y provinciales, DRA Piura y SUNARP (saneamiento de polígonos comunales).",
        "Cambio de autoridades regionales y locales el 01/01/2027: se requiere re-socializar y ratificar acuerdos en "
        "enero-febrero de 2027.",
    ])
    W.h(d, "5. Vacíos")
    W.vinetas(d, ["F-DS-04, F-DS-05 y F-DS-07 sin aplicar; Ayabaca sin F-DS-03 ni F-DS-06.",
                  "Tamizaje predial inicial: sin actas de libre disponibilidad conformes; titulares por identificar en 32 bloques "
                  "sin ficha social.",
                  "Posición de los presidentes de algunas comunidades campesinas por confirmar."])
    copiar(drv / "Matrices_E2_IN_Piura_B4.xlsx", an)
    copiar(drv / "Reporte_Liberacion_Areas_IN_Piura_20261109 AP3.xlsx", an)
    copiar(repo / "Plantilla_Diagnostico_Social_IN_Piura_V4.xlsx", an / "Instrumentos")
    excel_05(D, an / "Matriz_Involucrados_IN_Piura.xlsx")
    lista_anexos(d, [
        ("05_Nota_Tecnica_Involucrados.docx", "Esta nota", FECHA),
        ("Anexos/Matriz_Involucrados_IN_Piura.xlsx", "Matriz de involucrados y resumen de actores (extracto editable)", "Generado al corte"),
        ("Anexos/Matrices_E2_IN_Piura_B4.xlsx", "Hojas 12 a 20 y 28: actores, involucrados, riesgos sociales, conflictos, percepción y oportunidades institucionales", "B4, 07/11/2026"),
        ("Anexos/Reporte_Liberacion_Areas_IN_Piura_20261109 AP3.xlsx", "Tamizaje predial: titulares, predios inspeccionados y observaciones", "09/11/2026"),
        ("Anexos/Instrumentos/Plantilla_Diagnostico_Social_IN_Piura_V4.xlsx", "Fichas F-DS-02 (actores) y F-DS-03 (entrevistas)", "V4"),
    ])
    W.p(d, "El detalle por provincia (cap. 9.1.4 Mapa de actores y 9.1.5 Riesgos sociales) está en los volúmenes del E2 del paquete 03.",
        italic=True, size=8.5)
    W.firma(d)
    d.save(out / "05_Nota_Tecnica_Involucrados.docx")


def excel_05(D, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook(); ws = wb.active; ws.title = "Matriz_Involucrados"
    f = X.encabezado(ws, "MATRIZ DE POBLACIÓN AFECTADA Y OTROS INVOLUCRADOS (GUÍA GENERAL, TABLA 1.7)",
                     "Fuente: Matrices E2 B4, hoja 14 (F-DS-02, F-DS-03, F-DS-06; corte 07/11/2026).", 7)
    filas = []
    for r in D["involucrados"]:
        g = r["Grupo involucrado"] or ""
        if "ANIN" in g:
            g = "ANIN (unidad formuladora y ejecutora)"
        filas.append([r["Provincia"], g, r["Situación negativa percibida"], r["Intereses o expectativas"],
                      r["Estrategias del proyecto"], r["Acuerdos y compromisos"], r["Actitud (cooperante / oponente)"]])
    X.tabla(ws, f, ["Provincia", "Grupo involucrado", "Problema percibido", "Intereses / necesidades", "Estrategia",
                    "Acuerdos y compromisos", "Actitud"], filas, anchos=[13, 38, 40, 32, 40, 34, 26])
    ws2 = wb.create_sheet("Resumen_Actores")
    f = X.encabezado(ws2, "RESUMEN DE ACTORES MAPEADOS (F-DS-02)", "Fuente: Matrices E2 B4, hoja 13.", 10)
    filas = [[r["Provincia"], r["Tipo de actor"], r["N.° de actores"], r["Cooperantes"], r["Neutrales"], r["Oponentes"],
              r["Sin posición"], r["Influencia alta"], r["Interés alto"], r["Gestionar de cerca"]]
             for r in D["actores"] if r["Provincia"] in datos.PROVS and r["Tipo de actor"] != "Todos los tipos"]
    X.tabla(ws2, f, ["Provincia", "Tipo de actor", "N.° de actores", "Cooperantes", "Neutrales", "Oponentes", "Sin posición",
                     "Influencia alta", "Interés alto", "Gestionar de cerca"], filas,
            anchos=[13, 32, 12, 12, 10, 10, 12, 12, 12, 14], total_cols=[3, 4, 5, 6, 7, 8, 9, 10])
    wb.save(path)


# ======================================================================================== 06 (nota e inventario)
def paquete_06(D, repo, out):
    inv = json.loads((out / "_inventario_capas.json").read_text(encoding="utf8"))
    # el JSON se excluye al empaquetar
    wb = Workbook(); ws = wb.active; ws.title = "Inventario"
    f = X.encabezado(ws, "INVENTARIO DE INFORMACIÓN GEORREFERENCIADA Y CARTOGRÁFICA",
                     "SHP y GPKG en UTM WGS 84 Zona 17S (EPSG:32717), salvo indicación; KMZ en WGS 84 geográficas (EPSG:4326).", 8)
    filas = [[k + 1, r["Capa"], r["Formato"], r["Ruta"], r["Registros"], r["Geometria"], r["Fuente"], r["Nota"]]
             for k, r in enumerate(inv)]
    last = X.tabla(ws, f, ["N.°", "Capa / producto", "Formato", "Ruta en el paquete", "Registros", "Geometría", "Fuente",
                           "Nota"], filas, anchos=[6, 60, 11, 60, 10, 16, 45, 40])
    ws.cell(row=last + 2, column=1, value="Total de elementos inventariados:")
    ws.cell(row=last + 2, column=4, value=f"=COUNTA(B{f + 1}:B{f + len(filas)})")
    ws2 = wb.create_sheet("Por_adjuntar")
    f = X.encabezado(ws2, "INFORMACIÓN ESPACIAL POR ADJUNTAR", "Archivos no incluidos en este paquete.", 3)
    X.tabla(ws2, f, ["Paquete", "Archivo o insumo", "Situación"], [[c, t, m] for c, t, m in POR_ADJUNTAR],
            anchos=[10, 90, 70])
    wb.save(out / "06_Inventario_Capas_IN_Piura.xlsx")

    d = doc_base("NOTA TÉCNICA", "Información georreferenciada y cartográfica del Perfil",
                 "Información de avance del Perfil para su articulación con el Programa de Inversión del Río Piura – "
                 "Componente 6: información espacial")
    W.h(d, "1. Contenido")
    W.tabla(d, ["Carpeta", "Contenido"], [
        ["01_SHP", "Capas en shapefile con sus archivos complementarios (.shp, .shx, .dbf, .prj, .cpg y, en las nativas, .qmd/.sbn/.xml)"],
        ["02_KMZ", "Capas principales para Google Earth (WGS 84 geográficas)"],
        ["03_GPKG", "Geopaquetes nativos del proyecto (bloques V6, área de influencia, ecosistemas, lotes SUS, SERRN, medidas E10)"],
        ["04_Mapas", "Mapas temáticos en PDF/PNG (área de influencia, ecosistemas) y 15 mapas distritales de bloques (JPG)"],
        ["05_Procesos", "Scripts y leyendas usados para generar los mapas, cuadros y la aptitud de zanjas (reproducibles)"],
        ["06_Inventario_Capas_IN_Piura.xlsx", f"Inventario de {len(inv)} elementos con formato, registros, geometría y fuente"],
    ], anchos=[0.25, 0.75], size=8.5)
    W.h(d, "2. Sistema de referencia y controles")
    W.vinetas(d, [
        "Sistema de trabajo: UTM WGS 84 Zona 17S (EPSG:32717). La capa nativa de ecosistemas por microcuenca está en EPSG:32718 "
        "(se respeta su formato original); en el GPKG y en el SHP 08/09 se reproyectó a 17S.",
        f"Controles: {D['bloques_n']} bloques y {W.n(D['bloques_ha'], 3)} ha (V6/V7); {D['ai_n']} polígonos de AI con "
        f"{W.n(D['ai_ha'], 3)} ha; coordenadas de localidades validadas en Este 450,000-750,000 y Norte 9,300,000-9,600,000.",
        "Los bloques retirados (74 y 75) y sus AI no forman parte de las capas vigentes; se incluyen aparte como «bloques descartados».",
        "Catálogo V8: los polígonos de los bloques 8, 71, 22, M17B1 y 5 aún no se recortan; el descuento consta en el paquete 01.",
        "Las medidas E10 son una distribución preliminar de trabajo, no metas del Perfil.",
    ])
    pendientes(d, "06")
    W.h(d, "3. Partes del paquete")
    W.p(d, "Por su tamaño, el paquete 06 se entrega en dos archivos: «06A ... Capas» (SHP, KMZ, GPKG, procesos e inventario) y "
           "«06B ... Mapas» (PDF, PNG y JPG).")
    W.firma(d)
    d.save(out / "06_Nota_Tecnica_Informacion_Espacial.docx")


# ======================================================================================== memo e indice
def memo(D, base):
    d = doc_base("INFORME TÉCNICO N.° [POR DEFINIR]-2026-ANIN/DIME/SESDI",
                 "Remisión de información de avance del Perfil «Infraestructura Natural» (CUI 2669244)",
                 "Atención a la solicitud de información para incorporar el alcance del estudio en el Programa de Inversión "
                 "del Río Piura")
    W.h(d, "1. Antecedente")
    W.p(d, "Se solicitó información actualizada sobre los avances en la formulación del Perfil de Inversión «Infraestructura "
           "Natural», a fin de incorporar el alcance del estudio en el Programa de Inversión del Río Piura y realizar las "
           "acciones técnicas correspondientes para su formulación. El Perfil, a cargo de la ANIN por Administración Directa, "
           "registra un avance ponderado de 48.5 % al 03/11/2026.")
    W.h(d, "2. Información remitida")
    dt = D["demanda_total"]; pt = D["pob_total"]
    W.tabla(d, ["Paquete (ZIP)", "Contenido principal", "Cifras clave al corte"], [
        ["01 Delimitación del ámbito", "Área de estudio, área de intervención, área de influencia y zona de beneficio; criterios",
         f"{D['bloques_n']} bloques; {W.n(D['bloques_ha'], 3)} ha (V6/V7); {W.n(dt['Superficie neta V8 (ha)'], 2)} ha netas V8; "
         f"AI {W.n(D['ai_ha'], 3)} ha"],
        ["02 Criterios de selección", "Meso y microlocalización: PI MCA-AHP, exclusión, idoneidad, tamizaje, campo",
         f"{W.n(dt['ha en PI alto y muy alto'], 2)} ha en PI alto y muy alto ({W.pct(dt['% PI alto y muy alto (sobre V7)'])})"],
        ["03 Población afectada", "Metodología, población del AI 2017-2046, cobertura de fichas, hallazgos (E2 B4)",
         f"{W.n(pt['Pob. 2017 total'], 0)} hab. (2017); {W.n(pt['2026 (base)'], 0)} hab. (2026)"],
        ["04 Unidad Productora", "Ecosistemas, degradación (línea base), peligro y riesgo, suelos (E5, E6, E8)",
         "8,341.825 ha bajo el umbral MSAVI (68.0 %)"],
        ["05 Involucrados", "Mapa de actores, matriz de involucrados, riesgos sociales",
         f"{next(r['N.° de actores'] for r in D['actores'] if r['Provincia'] == 'Total ámbito' and r['Tipo de actor'] == 'Todos los tipos')} actores mapeados"],
        ["06A / 06B Información espacial", "SHP, KMZ, GPKG, mapas temáticos y distritales, procesos", "UTM WGS 84 Zona 17S"],
    ], anchos=[0.22, 0.43, 0.35], size=8)
    W.h(d, "3. Consideraciones")
    W.vinetas(d, [
        "Las cifras son preliminares y pueden ajustarse en la versión final del Perfil; la meta física se fija en Formulación "
        "(E10), después del tamizaje predial. El rango de intervención directa del proyecto es de 10,000 a 15,000 ha.",
        "El catálogo vigente es el V8 (09/10/2026): 117 bloques y 12,157.47 ha netas; las capas SIG mantienen la geometría V6/V7 "
        "hasta el recorte de cinco bloques.",
        "La zona de beneficio aguas abajo, que vincula el Perfil con la zona de impacto del Programa, está Por determinar y "
        "requiere el modelamiento hidrológico (E4).",
        "Los archivos que superan el límite de descarga o que no se ubicaron se listan como «por adjuntar» en cada nota y en el "
        "inventario del paquete 06.",
    ])
    W.h(d, "4. Recomendación")
    W.p(d, "Remitir los seis paquetes al equipo del Programa de Inversión del Río Piura y programar una reunión técnica para "
           "acordar la delimitación de la zona de beneficio, el intercambio de capas (zona de impacto, obras proyectadas) y "
           "la conciliación de la población beneficiaria entre ambos estudios.")
    W.firma(d)
    d.save(base / "00_Informe_Remision_Avance_Perfil_IN_Piura.docx")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo", required=True)
    ap.add_argument("--drive", required=True)
    ap.add_argument("--paquetes", required=True)
    a = ap.parse_args()
    repo, drv, base = Path(a.repo), Path(a.drive), Path(a.paquetes)
    D = datos.cargar(repo, drv)
    for k in ("01", "02", "03", "04", "05"):
        p = base / CARPETAS[k]
        if p.exists():
            shutil.rmtree(p)
        p.mkdir(parents=True)
    paquete_01(D, repo, drv, base / CARPETAS["01"])
    paquete_02(D, repo, drv, base / CARPETAS["02"])
    paquete_03(D, repo, drv, base / CARPETAS["03"])
    paquete_04(D, repo, drv, base / CARPETAS["04"])
    paquete_05(D, repo, drv, base / CARPETAS["05"])
    paquete_06(D, repo, base / CARPETAS["06"])
    memo(D, base)
    print("OK")


if __name__ == "__main__":
    main()
