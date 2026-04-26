"""
IN Piura - Plan de Ingreso / Verificacion de Campo
Aplicacion web con Streamlit.
Restauracion de ecosistemas - Cuenca alta del rio Piura, Peru.
"""

import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import os
import uuid
import io
import csv
import json
import tempfile

import database as db

# ── Constante de version de cache (incrementar tras escritura) ───────────
# Usada para invalidar @st.cache_data despues de inserciones/actualizaciones.
def _cache_version():
    """Retorna un contador que se incrementa con cada escritura a la BD."""
    if "db_cache_version" not in st.session_state:
        st.session_state["db_cache_version"] = 0
    return st.session_state["db_cache_version"]

def _invalidar_cache():
    """Incrementa el contador de cache para forzar recarga de datos."""
    st.session_state["db_cache_version"] = st.session_state.get("db_cache_version", 0) + 1

# ── Funciones cacheadas de lectura de BD ─────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)
def _cached_obtener_bloques(_version):
    return db.obtener_bloques()

@st.cache_data(ttl=300, show_spinner=False)
def _cached_obtener_todas_inspecciones(_version):
    return db.obtener_todas_inspecciones()

@st.cache_data(ttl=300, show_spinner=False)
def _cached_obtener_estadisticas(_version):
    return db.obtener_estadisticas_generales()

@st.cache_data(ttl=300, show_spinner=False)
def _cached_obtener_resumen_bloques(_version):
    return db.obtener_resumen_bloques()

@st.cache_data(ttl=300, show_spinner=False)
def _cached_obtener_todos_diagnosticos(_version):
    return db.obtener_todos_diagnosticos()

@st.cache_data(ttl=300, show_spinner=False)
def _cached_obtener_todos_diagnosticos_sociales(_version):
    return db.obtener_todos_diagnosticos_sociales()

# ── Constantes de fecha para validacion ──────────────────────────────────
FECHA_MIN_PROYECTO = date(2024, 1, 1)
FECHA_MAX_PROYECTO = date.today() + timedelta(days=365)

# ── Helper de paginacion ─────────────────────────────────────────────────
def _paginar(items, page_key, items_por_pagina=15):
    """Retorna (items_pagina, total_paginas, pagina_actual) para listas largas."""
    total = len(items)
    total_paginas = max(1, (total + items_por_pagina - 1) // items_por_pagina)
    if page_key not in st.session_state:
        st.session_state[page_key] = 1
    pagina = st.session_state[page_key]
    pagina = max(1, min(pagina, total_paginas))
    inicio = (pagina - 1) * items_por_pagina
    fin = inicio + items_por_pagina
    return items[inicio:fin], total_paginas, pagina

def _controles_paginacion(total_paginas, pagina_actual, page_key):
    """Muestra controles de paginacion si hay mas de 1 pagina."""
    if total_paginas <= 1:
        return
    cols = st.columns([1, 2, 1])
    with cols[0]:
        if st.button("Anterior", key=f"{page_key}_prev", disabled=pagina_actual <= 1):
            st.session_state[page_key] = pagina_actual - 1
            st.rerun()
    with cols[1]:
        st.markdown(f"<div style='text-align:center'>Pagina **{pagina_actual}** de **{total_paginas}**</div>",
                    unsafe_allow_html=True)
    with cols[2]:
        if st.button("Siguiente", key=f"{page_key}_next", disabled=pagina_actual >= total_paginas):
            st.session_state[page_key] = pagina_actual + 1
            st.rerun()
import reports
from georeferenciacion import utm_a_latlon, latlon_a_utm
from odk_kobo import generar_xlsform, importar_csv_odk, importar_desde_kobo, KoBoClient
from excel_diagnostico_social import generar_plantilla_ds, parsear_excel_ds, mapear_a_session_state
from excel_diagnostico_territorial import generar_plantilla_dt, parsear_excel_dt, mapear_dt_a_session_state
from excel_elementos_expuestos import (generar_plantilla_ee, parsear_excel_ee,
    mapear_a_session_state as mapear_ee_a_session_state,
    FEE01_TIPO_ELEMENTO, FEE01_UBICACION_PELIGRO, FEE_ESTADO, FEE_NIVEL_AMB,
    FEE_SI_NO, FEE02_MATERIAL_VIV, FEE03_SECTOR, FEE_TIPO_PELIGRO,
    FEE04_TIPO_ACTIVIDAD, FEE05_TIPO_DEGRADACION, FEE05_FUENTES_AGUA,
    FEE05_PROB_RECURRENCIA, FEE06_NIVEL_VULN, FEE06_ELEMENTOS, FEE06_FACTORES)

try:
    import pdf_converter as pdfconv
    PDF_CONV_OK = True
except ImportError:
    PDF_CONV_OK = False

# ── Configuracion ─────────────────────────────────────────────────────────
st.set_page_config(
    page_title="IN Piura - Plan de Ingreso",
    page_icon="\U0001F331",
    layout="wide",
    initial_sidebar_state="expanded",
)

db.inicializar_bd()

# ── Constantes ────────────────────────────────────────────────────────────
TIPOS_INTERVENCION = [
    "Revegetacion", "Zanjas de infiltracion",
    "Terrazas de formacion lenta", "Diques de mamposteria", "Lotes SUS", "Otras",
]
ESTADOS_BLOQUE = ["Pendiente", "En progreso", "Verificado"]
CONDICIONES_CLIMATICAS = [
    "Despejado", "Parcialmente nublado", "Nublado",
    "Lluvia ligera", "Lluvia moderada", "Lluvia intensa", "Neblina",
]
MICROCUENCAS = [
    "C1075-Q9580","C1076-Q9581","C1076-Q9584","C1076-Q9585","C1076-Q9586",
    "C1076-Q9587","C1076-Q9588","C1076-Q9589","C1076-Q9592","C1077-Q9566",
    "C1077-Q9579","C1078-Q9562","C1080-Q9560","C1081-Q9582","C1081-Q9583",
    "C1081-Q9591","C1086-Q9569","C1086-Q9570","C1086-Q9575","C1086-Q9576",
    "C1096-Q9545","C1096-Q9547","C1096-Q9556","C1096-Q9557","C1096-Q9558",
    "C1096-Q9564","C1107-Q9539","C1107-Q9541","C1108-Q9552",
]

# ── Datos de Origen: 124 Bloques de Intervencion V5 ──────────────────
# Fuente: Reporte_Bloques_V5_25abril.xlsx (124 bloques con centroides UTM Zona 17S WGS84)
# Cada entrada: (N, Bloque, Microcuenca, Area_ha, Provincia, Distrito, Accesibilidad, Dia, UTM_Este, UTM_Norte, MSAVI_2024)
BLOQUES_124 = [
    # (n, codigo, microcuenca, area_ha, provincia, distrito, accesibilidad, dia_eval, utm_este, utm_norte, msavi_2024)
    (1, "56", "C1096-Q9564", 77.666, "Ayabaca", "Frias", 0, 0, 621613, 9452884, 0.677658),
    (2, "57", "C1086-Q9576", 25.668, "Morropon", "Chalaco", 0, 0, 631723, 9439702, 0.603055),
    (3, "59", "C1086-Q9570", 13.554, "Morropon", "Chalaco", 0, 0, 631195, 9441777, 0.614739),
    (4, "58", "C1086-Q9569", 11.475, "Morropon", "Santa Catalina de Mossa", 0, 0, 619094, 9426406, 0.420325),
    (5, "61", "C1081-Q9590", 26.826, "Huancabamba", "Lalaquiz", 0, 0, 646236, 9425052, 0.611186),
    (6, "60", "C1081-Q9591", 40.132, "Huancabamba", "Canchaque", 0, 0, 653850, 9426236, 0.67447),
    (7, "64", "C1081-Q9591", 35.774, "Huancabamba", "Huancabamba", 0, 0, 657819, 9428541, 0.525448),
    (8, "66", "C1081-Q9591", 102.335, "Huancabamba", "Canchaque", 0, 0, 652566, 9417311, 0.640254),
    (9, "63", "C1081-Q9583", 50.459, "Huancabamba", "Canchaque", 0, 0, 655059, 9413845, 0.603582),
    (10, "75", "C1076-Q9587", 15.176, "Huancabamba", "San Miguel de El Faique", 0, 0, 657722, 9401669, 0.630888),
    (11, "67", "C1076-Q9586", 13.184, "Huancabamba", "San Miguel de El Faique", 0, 0, 654207, 9394210, 0.605293),
    (12, "70", "C1076-Q9592", 22.388, "Huancabamba", "Huarmaca", 0, 0, 666781, 9395573, 0.622295),
    (13, "68", "C1076-Q9592", 21.291, "Huancabamba", "Huarmaca", 0, 0, 664225, 9396339, 0.482981),
    (14, "73", "C1076-Q9592", 54.29, "Huancabamba", "Huarmaca", 0, 0, 658547, 9395594, 0.554974),
    (15, "69", "C1076-Q9592", 19.425, "Huancabamba", "Huarmaca", 0, 0, 663291, 9392332, 0.448926),
    (16, "72", "C1076-Q9592", 12.398, "Huancabamba", "Huarmaca", 0, 0, 662236, 9391211, 0.518477),
    (17, "71", "C1076-Q9593", 10.673, "Huancabamba", "Huarmaca", 0, 0, 652104, 9373045, 0.557066),
    (18, "80", "C1076-Q9593", 28.83, "Huancabamba", "Huarmaca", 0, 0, 657160, 9373050, 0.593069),
    (19, "74", "C1076-Q9588", 29.226, "Huancabamba", "Huarmaca", 0, 0, 652449, 9389179, 0.356463),
    (20, "76", "C1076-Q9593", 19.615, "Huancabamba", "Huarmaca", 0, 0, 657922, 9377477, 0.492462),
    (21, "79", "C1076-Q9587", 94.812, "Huancabamba", "Canchaque", 0, 0, 650305, 9402785, 0.547879),
    (22, "77", "C1081-Q9583", 17.821, "Morropon", "San Juan de Bigote", 0, 0, 638648, 9411370, 0.272488),
    (23, "82", "C1086-Q9570", 35.661, "Morropon", "Santo Domingo", 0, 0, 621452, 9443855, 0.62835),
    (24, "M17B1", "C1096-Q9556", 42.449, "Morropon", "Chulucanas", 0, 0, 600644, 9440436, 0.332263),
    (25, "32", "C1096-Q9556", 35.83, "Morropon", "Chulucanas", 0, 0, 599927, 9442604, 0.371619),
    (26, "M22B1", "C1076-Q9586", 55.17, "Huancabamba", "San Miguel de El Faique", 0, 0, 664082, 9404172, 0.594025),
    (27, "M6B2-1", "C1077-Q9566", 514.359, "Morropon", "Buenos Aires", 0, 0, 612864, 9413345, 0.3381),
    (28, "M6B2-2", "C1077-Q9566", 333.011, "Morropon", "Buenos Aires", 0, 0, 611196, 9416725, 0.302636),
    (29, "M6B10", "C1077-Q9566", 249.89, "Morropon", "Buenos Aires", 0, 0, 618299, 9423119, 0.30989),
    (30, "M17B4", "C1096-Q9556", 124.709, "Morropon", "Chulucanas", 0, 0, 604926, 9444327, 0.423543),
    (31, "M10B4", "C1096-Q9547", 68.887, "Morropon", "Chulucanas", 0, 0, 596773, 9448273, 0.313301),
    (32, "M12B1", "C1076-Q9588", 268.393, "Huancabamba", "Huarmaca", 0, 0, 644546, 9389106, 0.34753),
    (33, "M3B3", "C1081-Q9582", 84.825, "Morropon", "San Juan de Bigote", 0, 0, 633015, 9411483, 0.298483),
    (34, "M17B7", "C1096-Q9556", 258.837, "Ayabaca", "Frias", 0, 0, 608472, 9448056, 0.508715),
    (35, "M18B5", "C1077-Q9579", 197.352, "Morropon", "Salitral", 0, 0, 621551, 9412989, 0.356522),
    (36, "M28B4", "C1086-Q9575", 283.428, "Morropon", "Yamango", 0, 0, 626078, 9427110, 0.42751),
    (37, "M8B2", "C1096-Q9558", 35.347, "Morropon", "Santo Domingo", 0, 0, 617596, 9436227, 0.530453),
    (38, "M4B3", "C1076-Q9589", 234.069, "Huancabamba", "Huarmaca", 0, 0, 643474, 9385620, 0.322153),
    (39, "M4B4", "C1076-Q9589", 289.914, "Huancabamba", "Huarmaca", 0, 0, 642775, 9388574, 0.319788),
    (40, "M3B9", "C1081-Q9582", 294.204, "Morropon", "Salitral", 0, 0, 629338, 9413305, 0.321386),
    (41, "M3B8", "C1081-Q9582", 565.584, "Morropon", "San Juan de Bigote", 0, 0, 632119, 9415141, 0.293351),
    (42, "M3B7", "C1081-Q9582", 122.946, "Morropon", "San Juan de Bigote", 0, 0, 638309, 9415217, 0.370678),
    (43, "M3B5", "C1081-Q9582", 52.743, "Morropon", "San Juan de Bigote", 0, 0, 635887, 9411738, 0.283016),
    (44, "M3B6", "C1081-Q9582", 60.352, "Morropon", "San Juan de Bigote", 0, 0, 637918, 9412134, 0.342535),
    (45, "M11B3", "C1081-Q9583", 106.027, "Morropon", "San Juan de Bigote", 0, 0, 642079, 9413823, 0.395609),
    (46, "M18B3", "C1077-Q9579", 373.807, "Morropon", "Salitral", 0, 0, 625311, 9412195, 0.369424),
    (47, "M17B6", "C1096-Q9556", 106.538, "Ayabaca", "Frias", 0, 0, 607036, 9446211, 0.446441),
    (48, "M28B2", "C1086-Q9575", 90.197, "Morropon", "Yamango", 0, 0, 636249, 9425665, 0.475342),
    (49, "M2B5", "C1076-Q9584", 30.909, "Morropon", "Salitral", 0, 0, 639987, 9394392, 0.304733),
    (50, "M17B5", "C1096-Q9556", 74.028, "Ayabaca", "Frias", 0, 0, 608800, 9445585, 0.54817),
    (51, "M28B3", "C1086-Q9575", 58.292, "Morropon", "Yamango", 0, 0, 634061, 9426778, 0.507529),
    (52, "33", "C1081-Q9590", 29.681, "Huancabamba", "Lalaquiz", 0, 0, 646466, 9431126, 0.71923),
    (53, "50", "C1081-Q9590", 19.266, "Huancabamba", "Lalaquiz", 0, 0, 647350, 9427252, 0.68307),
    (54, "38", "C1081-Q9590", 44.201, "Huancabamba", "Lalaquiz", 0, 0, 647123, 9427602, 0.672487),
    (55, "34", "C1081-Q9590", 28.45, "Huancabamba", "Lalaquiz", 0, 0, 646483, 9426305, 0.648975),
    (56, "26", "C1081-Q9591", 46.962, "Huancabamba", "Lalaquiz", 0, 0, 649472, 9425705, 0.673303),
    (57, "37", "C1081-Q9591", 35.491, "Huancabamba", "Lalaquiz", 0, 0, 649081, 9424506, 0.658846),
    (58, "28", "C1081-Q9591", 52.966, "Huancabamba", "Canchaque", 0, 0, 654188, 9422849, 0.639425),
    (59, "40", "C1081-Q9591", 28.523, "Huancabamba", "Canchaque", 0, 0, 654137, 9424425, 0.640099),
    (60, "21", "C1081-Q9591", 84.227, "Huancabamba", "Canchaque", 0, 0, 654604, 9417736, 0.620078),
    (61, "24", "C1081-Q9583", 90.159, "Huancabamba", "Canchaque", 0, 0, 652191, 9414358, 0.674839),
    (62, "47", "C1076-Q9586", 13.551, "Huancabamba", "San Miguel de El Faique", 0, 0, 659649, 9401436, 0.593628),
    (63, "43", "C1076-Q9586", 23.171, "Huancabamba", "San Miguel de El Faique", 0, 0, 656489, 9395403, 0.551525),
    (64, "35", "C1076-Q9586", 30.804, "Huancabamba", "San Miguel de El Faique", 0, 0, 654607, 9395038, 0.624477),
    (65, "44", "C1076-Q9586", 28.234, "Huancabamba", "San Miguel de El Faique", 0, 0, 663637, 9400618, 0.511324),
    (66, "53", "C1076-Q9592", 13.065, "Huancabamba", "Huarmaca", 0, 0, 666537, 9395918, 0.596277),
    (67, "41", "C1076-Q9592", 64.923, "Huancabamba", "Huarmaca", 0, 0, 660525, 9394327, 0.592669),
    (68, "19", "C1076-Q9592", 82.088, "Huancabamba", "Huarmaca", 0, 0, 665942, 9391144, 0.537184),
    (69, "31", "C1076-Q9592", 47.581, "Huancabamba", "Huarmaca", 0, 0, 663334, 9392820, 0.435805),
    (70, "20", "C1076-Q9593", 81.334, "Huancabamba", "Huarmaca", 0, 0, 649847, 9376260, 0.48513),
    (71, "51", "C1076-Q9593", 22.927, "Huancabamba", "Huarmaca", 0, 0, 656291, 9379096, 0.523788),
    (72, "29", "C1076-Q9592", 40.104, "Huancabamba", "Huarmaca", 0, 0, 662123, 9397575, 0.548192),
    (73, "1", "C1076-Q9584", 1370.958, "Morropon", "Salitral", 0, 0, 645575, 9393500, 0.454493),
    (74, "7", "C1076-Q9584", 173.428, "Morropon", "Salitral", 0, 0, 639728, 9396836, 0.375288),
    (75, "5", "C1077-Q9566", 295.726, "Morropon", "Buenos Aires", 0, 0, 617860, 9413182, 0.3394),
    (76, "2", "C1096-Q9558", 330.599, "Morropon", "Chulucanas", 0, 0, 603653, 9439351, 0.358601),
    (77, "25", "C1096-Q9547", 53.484, "Morropon", "Chulucanas", 0, 0, 596578, 9442453, 0.256543),
    (78, "9", "C1076-Q9586", 109.861, "Huancabamba", "San Miguel de El Faique", 0, 0, 656578, 9397762, 0.574527),
    (79, "4", "C1076-Q9588", 230.098, "Huancabamba", "Huarmaca", 0, 0, 656549, 9382009, 0.660111),
    (80, "18", "C1076-Q9592", 103.617, "Huancabamba", "Huarmaca", 0, 0, 656445, 9393180, 0.640219),
    (81, "10", "C1076-Q9593", 150.214, "Huancabamba", "Huarmaca", 0, 0, 650830, 9382303, 0.546492),
    (82, "49", "C1086-Q9575", 15.839, "Morropon", "Yamango", 0, 0, 642334, 9438263, 0.686036),
    (83, "55", "C1086-Q9575", 0.859, "Morropon", "Yamango", 0, 0, 642690, 9438409, 0.630032),
    (84, "52", "C1086-Q9575", 11.059, "Morropon", "Yamango", 0, 0, 641522, 9438064, 0.66855),
    (85, "46", "C1086-Q9575", 18.866, "Morropon", "Yamango", 0, 0, 638751, 9433236, 0.724382),
    (86, "48", "C1086-Q9575", 5.754, "Morropon", "Yamango", 0, 0, 639496, 9433679, 0.697809),
    (87, "3", "C1096-Q9545", 459.295, "Ayabaca", "Frias", 0, 0, 606816, 9458469, 0.567707),
    (88, "6", "C1096-Q9547", 168.411, "Ayabaca", "Frias", 0, 0, 605636, 9455269, 0.538144),
    (89, "36", "C1096-Q9547", 24.843, "Ayabaca", "Frias", 0, 0, 606468, 9454689, 0.539458),
    (90, "27", "C1096-Q9556", 78.154, "Ayabaca", "Frias", 0, 0, 613302, 9450231, 0.53228),
    (91, "12", "C1086-Q9570", 119.538, "Morropon", "Santo Domingo", 0, 0, 625419, 9444932, 0.706689),
    (92, "13", "C1096-Q9564", 100.514, "Morropon", "Santo Domingo", 0, 0, 620263, 9442343, 0.671094),
    (93, "11", "C1086-Q9570", 103.948, "Morropon", "Santo Domingo", 0, 0, 620759, 9439988, 0.604743),
    (94, "17", "C1086-Q9570", 96.079, "Morropon", "Santo Domingo", 0, 0, 618661, 9438267, 0.647895),
    (95, "23", "C1086-Q9570", 81.776, "Morropon", "Santo Domingo", 0, 0, 619939, 9438544, 0.573141),
    (96, "15", "C1086-Q9570", 86.147, "Morropon", "Santo Domingo", 0, 0, 620200, 9437862, 0.560518),
    (97, "14", "C1086-Q9570", 92.905, "Morropon", "Santa Catalina de Mossa", 0, 0, 624397, 9437085, 0.51974),
    (98, "42", "C1086-Q9570", 80.616, "Morropon", "Santa Catalina de Mossa", 0, 0, 625152, 9435011, 0.634841),
    (99, "30", "C1081-Q9591", 34.874, "Huancabamba", "Canchaque", 0, 0, 652758, 9424380, 0.663804),
    (100, "8", "C1076-Q9593", 126.952, "Huancabamba", "Huarmaca", 0, 0, 651673, 9373383, 0.554802),
    (101, "M17B10", "C1096-Q9556", 75.363, "Ayabaca", "Frias", 0, 0, 606173, 9447361, 0.472466),
    (102, "39", "C1096-Q9564", 68.569, "Ayabaca", "Frias", 0, 0, 618523, 9448918, 0.701601),
    (103, "M19B7", "C1086-Q9570", 34.171, "Morropon", "Santo Domingo", 0, 0, 620770, 9436562, 0.591921),
    (104, "M36B2", "C1086-Q9576", 28.19, "Morropon", "Santa Catalina de Mossa", 0, 0, 626438, 9435060, 0.597433),
    (105, "M30B5", "C1081-Q9591", 90.903, "Huancabamba", "Huancabamba", 0, 0, 655263, 9427340, 0.5564),
    (106, "16", "C1076-Q9593", 83.327, "Huancabamba", "Huarmaca", 0, 0, 652939, 9382105, 0.628423),
    (107, "M18B1", "C1077-Q9579", 160.36, "Morropon", "Buenos Aires", 0, 0, 621058, 9410427, 0.334299),
    (108, "M1B1", "C1077-Q9580", 188.797, "Morropon", "Buenos Aires", 0, 0, 617695, 9409136, 0.319305),
    (109, "M6B2-3", "C1077-Q9566", 160.978, "Morropon", "Buenos Aires", 0, 0, 611305, 9419274, 0.406259),
    (110, "M32B3", "C1086-Q9569", 50.972, "Morropon", "Morropon", 0, 0, 618248, 9426493, 0.385213),
    (111, "M19B2", "C1086-Q9570", 74.372, "Morropon", "Morropon", 0, 0, 617715, 9427284, 0.401185),
    (112, "22", "C1096-Q9556", 79.33, "Morropon", "Chulucanas", 0, 0, 601122, 9443651, 0.42317),
    (113, "M2B8", "C1076-Q9584", 116.099, "Huancabamba", "San Miguel de El Faique", 0, 0, 639059, 9398295, 0.375171),
    (114, "M3B1", "C1081-Q9582", 80.983, "Morropon", "Salitral", 0, 0, 630685, 9410800, 0.323882),
    (115, "M20B1", "C1076-Q9585", 279.68, "Huancabamba", "San Miguel de El Faique", 0, 0, 641196, 9398710, 0.355329),
    (116, "M7B1", "C1076-Q9581", 130.425, "Morropon", "Salitral", 0, 0, 629006, 9407536, 0.313647),
    (117, "M7B6", "C1076-Q9581", 91.285, "Morropon", "Salitral", 0, 0, 635110, 9399092, 0.287833),
    (118, "M7B2", "C1076-Q9581", 115.379, "Morropon", "Salitral", 0, 0, 631438, 9403547, 0.287689),
    (119, "M2B1", "C1076-Q9584", 81.795, "Morropon", "Salitral", 0, 0, 635876, 9397941, 0.29836),
    (120, "M7B3", "C1076-Q9581", 65.745, "Morropon", "Salitral", 0, 0, 633232, 9401227, 0.298017),
    (121, "M27B1", "C1096-Q9557", 448.804, "Morropon", "Chulucanas", 0, 0, 593246, 9430139, 0.363228),
    (122, "M9B1", "C1096-Q9545", 355.364, "Morropon", "Chulucanas", 0, 0, 595552, 9451222, 0.314312),
    (123, "78", "C1086-Q9570", 35.1, "Morropon", "Santo Domingo", 0, 0, 619630, 9440043, 0.707402),
    (124, "81", "C1076-Q9585", 8.929, "Huancabamba", "Canchaque", 0, 0, 641538, 9401360, 0.274714),
]


# Diccionario para busqueda rapida por codigo de bloque
BLOQUES_124_MAP = {b[1]: {"n": b[0], "codigo": b[1], "microcuenca": b[2],
    "area_ha": b[3], "provincia": b[4], "distrito": b[5],
    "accesibilidad": b[6], "dia_evaluacion": b[7],
    "utm_este": b[8], "utm_norte": b[9], "msavi_2024": b[10]} for b in BLOQUES_124}

# Lista de codigos para el dropdown (solo codigo de bloque)
BLOQUES_124_OPCIONES = [b[1] for b in BLOQUES_124]

# ── Centros Poblados por Bloque de Intervencion V5 ───────────────────
# Fuente: CP´s y Bloques_V5.xls
CENTROS_POBLADOS_BLOQUE = {
    "1": {"centros_poblados": ["Gramadal", "Mamayaco"], "comunidades_campesinas": [], "poblacion_total": 0},
    "2": {"centros_poblados": ["Balcon de Talandracas"], "comunidades_campesinas": [], "poblacion_total": 0},
    "3": {"centros_poblados": ["Huayabal", "El mirador", "Los checches", "Huasipe de Geraldo", "Rincon de Geraldo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "4": {"centros_poblados": ["La rinconada", "San Antonio de Succhirca"], "comunidades_campesinas": [], "poblacion_total": 0},
    "5": {"centros_poblados": ["Juan Velasco", "El ala", "Linderos del ala"], "comunidades_campesinas": [], "poblacion_total": 0},
    "6": {"centros_poblados": ["Guanabano alto", "Nogal", "Guabal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "7": {"centros_poblados": ["Hornopampa"], "comunidades_campesinas": [], "poblacion_total": 0},
    "8": {"centros_poblados": ["Pirga", "Chonta de Platanal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "9": {"centros_poblados": ["La capilla", "El higueron"], "comunidades_campesinas": [], "poblacion_total": 0},
    "10": {"centros_poblados": ["La peña"], "comunidades_campesinas": [], "poblacion_total": 0},
    "11": {"centros_poblados": ["Taylin de Tuñali"], "comunidades_campesinas": [], "poblacion_total": 0},
    "12": {"centros_poblados": ["Santo Domingo", "Chacayo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "13": {"centros_poblados": ["Huacas"], "comunidades_campesinas": [], "poblacion_total": 0},
    "14": {"centros_poblados": ["La laja", "Overazal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "15": {"centros_poblados": ["Jacanacas"], "comunidades_campesinas": [], "poblacion_total": 0},
    "16": {"centros_poblados": ["Laguna de paltama"], "comunidades_campesinas": [], "poblacion_total": 0},
    "17": {"centros_poblados": ["El Checo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "18": {"centros_poblados": ["Piedra blanca"], "comunidades_campesinas": [], "poblacion_total": 0},
    "19": {"centros_poblados": ["Nuevo progreso"], "comunidades_campesinas": [], "poblacion_total": 0},
    "20": {"centros_poblados": ["Tupac Amaru"], "comunidades_campesinas": [], "poblacion_total": 0},
    "21": {"centros_poblados": ["Huamala alto"], "comunidades_campesinas": [], "poblacion_total": 0},
    "22": {"centros_poblados": ["Platanal bajo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "23": {"centros_poblados": ["Nueva esperanza", "Jacanacas"], "comunidades_campesinas": [], "poblacion_total": 0},
    "24": {"centros_poblados": ["Coyona"], "comunidades_campesinas": [], "poblacion_total": 0},
    "25": {"centros_poblados": ["Cruzpampa-Yapatera"], "comunidades_campesinas": [], "poblacion_total": 0},
    "26": {"centros_poblados": ["La laguna"], "comunidades_campesinas": [], "poblacion_total": 0},
    "27": {"centros_poblados": ["linderos de Misquis", "Nueva Esperanza de Misquis"], "comunidades_campesinas": [], "poblacion_total": 0},
    "28": {"centros_poblados": ["Abalque"], "comunidades_campesinas": [], "poblacion_total": 0},
    "29": {"centros_poblados": ["Santa Cruz"], "comunidades_campesinas": [], "poblacion_total": 0},
    "30": {"centros_poblados": ["Chamelico", "Sapce"], "comunidades_campesinas": [], "poblacion_total": 0},
    "31": {"centros_poblados": ["Molle", "Hualanga pampa"], "comunidades_campesinas": [], "poblacion_total": 0},
    "32": {"centros_poblados": ["Panecillo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "33": {"centros_poblados": ["Sambe"], "comunidades_campesinas": [], "poblacion_total": 0},
    "34": {"centros_poblados": ["Maray grande", "Maray chico"], "comunidades_campesinas": [], "poblacion_total": 0},
    "35": {"centros_poblados": ["Chamelico"], "comunidades_campesinas": [], "poblacion_total": 0},
    "36": {"centros_poblados": ["Guabal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "37": {"centros_poblados": ["El papayo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "38": {"centros_poblados": ["Ullma"], "comunidades_campesinas": [], "poblacion_total": 0},
    "39": {"centros_poblados": ["La cria"], "comunidades_campesinas": [], "poblacion_total": 0},
    "40": {"centros_poblados": ["Chacchacal", "Chorro blanco"], "comunidades_campesinas": [], "poblacion_total": 0},
    "41": {"centros_poblados": ["Cambruran"], "comunidades_campesinas": [], "poblacion_total": 0},
    "42": {"centros_poblados": ["Santa Rosa de Chirimoyos"], "comunidades_campesinas": [], "poblacion_total": 0},
    "43": {"centros_poblados": ["Coyona", "Pizarrume"], "comunidades_campesinas": [], "poblacion_total": 0},
    "44": {"centros_poblados": ["Cruz de piedra"], "comunidades_campesinas": [], "poblacion_total": 0},
    "46": {"centros_poblados": ["Nangay"], "comunidades_campesinas": [], "poblacion_total": 0},
    "47": {"centros_poblados": ["Gaspar"], "comunidades_campesinas": [], "poblacion_total": 0},
    "48": {"centros_poblados": ["Nangay"], "comunidades_campesinas": [], "poblacion_total": 0},
    "49": {"centros_poblados": ["Las huacas"], "comunidades_campesinas": [], "poblacion_total": 0},
    "50": {"centros_poblados": ["Ullma"], "comunidades_campesinas": [], "poblacion_total": 0},
    "51": {"centros_poblados": ["Talla", "Shain"], "comunidades_campesinas": [], "poblacion_total": 0},
    "52": {"centros_poblados": ["Miraflores"], "comunidades_campesinas": [], "poblacion_total": 0},
    "53": {"centros_poblados": ["Zururan"], "comunidades_campesinas": [], "poblacion_total": 0},
    "55": {"centros_poblados": ["Las huacas"], "comunidades_campesinas": [], "poblacion_total": 0},
    "56": {"centros_poblados": ["Putagas", "Banda de la cruz"], "comunidades_campesinas": [], "poblacion_total": 0},
    "57": {"centros_poblados": ["Taspa", "San Lorenzo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "58": {"centros_poblados": ["Maray"], "comunidades_campesinas": [], "poblacion_total": 0},
    "59": {"centros_poblados": ["Guabo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "60": {"centros_poblados": ["Flor de café"], "comunidades_campesinas": [], "poblacion_total": 0},
    "61": {"centros_poblados": ["Pedregal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "63": {"centros_poblados": ["Shuturumbe"], "comunidades_campesinas": [], "poblacion_total": 0},
    "64": {"centros_poblados": ["Pariamarca centro"], "comunidades_campesinas": [], "poblacion_total": 0},
    "66": {"centros_poblados": ["Huamala"], "comunidades_campesinas": [], "poblacion_total": 0},
    "67": {"centros_poblados": ["Piedra grande", "Quitahuajara"], "comunidades_campesinas": [], "poblacion_total": 0},
    "68": {"centros_poblados": ["Tablurán"], "comunidades_campesinas": [], "poblacion_total": 0},
    "69": {"centros_poblados": ["Molle", "Hualanga pampa"], "comunidades_campesinas": [], "poblacion_total": 0},
    "70": {"centros_poblados": ["Zururan"], "comunidades_campesinas": [], "poblacion_total": 0},
    "71": {"centros_poblados": ["Pirga", "Chonta de platanal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "72": {"centros_poblados": ["Ramon Castilla"], "comunidades_campesinas": [], "poblacion_total": 0},
    "73": {"centros_poblados": ["Chococa"], "comunidades_campesinas": [], "poblacion_total": 0},
    "74": {"centros_poblados": ["Chalpa"], "comunidades_campesinas": [], "poblacion_total": 0},
    "75": {"centros_poblados": ["San Cristobal", "Sanchez Cerro"], "comunidades_campesinas": [], "poblacion_total": 0},
    "76": {"centros_poblados": ["Sahuate Hualanga", "Hualanga"], "comunidades_campesinas": [], "poblacion_total": 0},
    "77": {"centros_poblados": ["Miguelpampa"], "comunidades_campesinas": [], "poblacion_total": 0},
    "78": {"centros_poblados": ["Tasajeras"], "comunidades_campesinas": [], "poblacion_total": 0},
    "79": {"centros_poblados": ["Almirante Miguel Grau", "Huabal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "80": {"centros_poblados": ["Yumbe", "Cruz roja"], "comunidades_campesinas": [], "poblacion_total": 0},
    "81": {"centros_poblados": ["Hualtacal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "82": {"centros_poblados": ["Santa fe de Portachuelo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M10B4": {"centros_poblados": ["Rio seco alto"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M11B3": {"centros_poblados": ["Quemazon", "La pareja"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M12B1": {"centros_poblados": ["Nueva esperanza", "Faicalito"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M17B1": {"centros_poblados": ["Papelillo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M17B10": {"centros_poblados": ["Platanal alto"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M17B4": {"centros_poblados": ["Chililique alto"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M17B5": {"centros_poblados": ["El guabo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M17B6": {"centros_poblados": ["Platanal alto", "El guabo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M17B7": {"centros_poblados": ["Pampa de ramada"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M18B1": {"centros_poblados": ["Hualas"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M18B3": {"centros_poblados": ["Selva andina", "Huaro quispampa", "Mangomanguia"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M18B5": {"centros_poblados": ["Morroponcito"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M19B2": {"centros_poblados": ["Boca negra"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M19B7": {"centros_poblados": ["El Faique"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M1B1": {"centros_poblados": ["Rio seco"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M20B1": {"centros_poblados": ["Las huacas"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M22B1": {"centros_poblados": ["Santa Rosa"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M27B1": {"centros_poblados": ["Vicus linderos", "Vicus la merced", "Vicus santa Rosa", "Huasimal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M28B2": {"centros_poblados": ["Flor de agua", "Victor Raul (El Checo)"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M28B3": {"centros_poblados": ["Ricardo Palma"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M28B4": {"centros_poblados": ["Mambluque", "Alto Mambluque"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M2B1": {"centros_poblados": ["Santa Rosa"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M2B5": {"centros_poblados": ["Hornopampa"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M2B8": {"centros_poblados": ["Huacas baja"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M30B5": {"centros_poblados": ["Pariamarca centro"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M32B3": {"centros_poblados": ["Maray"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M36B2": {"centros_poblados": ["Santa Rosa de Chirimoyos"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M3B1": {"centros_poblados": ["Piedra blanca"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M3B3": {"centros_poblados": ["Alan Garcia", "Bigote"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M3B5": {"centros_poblados": ["San Juan Bautista"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M3B6": {"centros_poblados": ["Bado de garzas", "Manzanares"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M3B7": {"centros_poblados": ["Sinai", "Polluco"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M3B8": {"centros_poblados": ["Santa Rosa"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M3B9": {"centros_poblados": ["San Pedro", "Tortola"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M4B3": {"centros_poblados": ["Chignia baja"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M4B4": {"centros_poblados": ["Hualcas I", "Hualcas II"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M6B10": {"centros_poblados": ["La Maravilla", "La Pilca", "Ingenio de Buenos Aires"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M6B2-1": {"centros_poblados": ["Buenos Aires"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M6B2-2": {"centros_poblados": ["Buenos Aires"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M6B2-3": {"centros_poblados": ["Buenos Aires"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M7B1": {"centros_poblados": ["Victor Raul", "Nuevo progreso", "La alberca"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M7B2": {"centros_poblados": ["Palo blanco-El cerezo", "La alberca"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M7B3": {"centros_poblados": ["Palo blanco-El cerezo", "La tranca"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M7B6": {"centros_poblados": ["Serran", "Nuevo san Juan"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M8B2": {"centros_poblados": ["Botijas"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M9B1": {"centros_poblados": ["La peña"], "comunidades_campesinas": [], "poblacion_total": 0},
}

PROVINCIAS_DISTRITOS = {
    "Ayabaca": ["Frias"],
    "Huancabamba": ["Canchaque","Huancabamba","Huarmaca","Lalaquiz","San Miguel de El Faique"],
    "Morropon": ["Buenos Aires","Chalaco","Chulucanas","La Matanza","Morropon",
                 "Salitral","San Juan de Bigote","Santa Catalina de Mossa",
                 "Santo Domingo","Yamango"],
    "Piura": ["Las Lomas","Tambo Grande"],
    "Sullana": ["Sullana"],
}
PROVINCIAS = list(PROVINCIAS_DISTRITOS.keys())
DISTRITOS_PIURA = [d for ds in PROVINCIAS_DISTRITOS.values() for d in ds]
TIPOS_COBERTURA = ["Arborea","Arbustiva","Herbacea","Mixta"]
VIGOR_COBERTURA = ["Excelente","Bueno","Regular","Deficiente","Muy deficiente"]
CATEGORIAS_PRESUPUESTO = [
    "Mano de obra","Materiales e insumos","Equipos y herramientas",
    "Transporte y logistica","Plantones y semillas","Asistencia tecnica",
    "Supervision y monitoreo","Capacitacion","Gastos administrativos","Otros",
]
FUENTES_FINANCIAMIENTO = [
    "Presupuesto publico","Cooperacion internacional","Canon y sobrecanon",
    "Recursos propios","Donaciones","Otro",
]
ESTADOS_ACTIVIDAD = ["Programado","En ejecucion","Completado","Retrasado","Suspendido"]
ACTIVIDADES_TIPO = [
    "Preparacion de terreno","Produccion de plantones","Plantacion / Revegetacion",
    "Excavacion de zanjas de infiltracion","Construccion de terrazas",
    "Construccion de diques","Mantenimiento y riego","Monitoreo y evaluacion",
    "Capacitacion a comunidades","Supervision tecnica","Elaboracion de informes",
    "Otra actividad",
]
COLORES_ESTADO = {"Pendiente":[231,76,60],"En progreso":[243,156,18],"Verificado":[39,174,96]}

# ── Fichas de Diagnostico Territorial (F-DT-01 a F-DT-06) ───────────────
FICHAS_DT = ["F-DT-01","F-DT-02","F-DT-03","F-DT-04","F-DT-05","F-DT-06"]

# F-DT-01: CARACTERISTICAS FISIOGRAFICAS
FDT01_FORMA_TERRENO = [
    "Plano (0-2%)", "Ondulado (2-8%)", "Colinado (8-25%)",
    "Montanoso (25-50%)", "Escarpado (>50%)",
]
FDT01_PENDIENTE = [
    "0-5% (Plano a ligeramente inclinado)",
    "5-15% (Ligeramente inclinado a moderadamente empinado)",
    "15-25% (Moderadamente empinado)",
    "25-50% (Empinado)",
    ">50% (Muy empinado a extremadamente empinado)",
]
FDT01_POSICION_FISIOGRAFICA = [
    "Cima / Cresta", "Ladera superior (tercio superior)",
    "Ladera media (tercio medio)", "Ladera inferior (tercio inferior)",
    "Pie de ladera / Base", "Fondo de valle", "Terraza fluvial", "Llanura",
]
FDT01_EXPOSICION = [
    "Norte (N)", "Sur (S)", "Este (E)", "Oeste (O)",
    "Noreste (NE)", "Noroeste (NO)", "Sureste (SE)", "Suroeste (SO)",
    "Variable / Sin predominancia",
]
FDT01_PAISAJE = [
    "Montana", "Colina", "Piedemonte", "Planicie / Llanura",
    "Valle interandino", "Terraza aluvial",
]
FDT01_RANGO_ALTITUDINAL = [
    "<500 msnm", "500-1000 msnm", "1000-1500 msnm",
    "1500-2000 msnm", "2000-2500 msnm", "2500-3000 msnm",
    "3000-3500 msnm", "3500-4000 msnm", ">4000 msnm",
]

# F-DT-02: CONDICIONES CLIMATICAS
FDT02_PRECIPITACION = [
    "<250 mm/anio (Muy seco / Arido)",
    "250-500 mm/anio (Seco / Semiarido)",
    "500-1000 mm/anio (Sub-humedo)",
    "1000-2000 mm/anio (Humedo / Lluvioso)",
    ">2000 mm/anio (Muy humedo / Pluvial)",
]
FDT02_TEMPERATURA = [
    "<5 C (Muy frio / Gelido)",
    "5-12 C (Frio)",
    "12-18 C (Templado)",
    "18-24 C (Calido / Semicalido)",
    ">24 C (Muy calido / Tropical)",
]
FDT02_HUMEDAD = [
    "Muy baja (<30%)", "Baja (30-50%)", "Media (50-70%)",
    "Alta (70-85%)", "Muy alta (>85%)",
]
FDT02_ZONA_VIDA = [
    "Desierto superarido", "Desierto arido", "Matorral desertico",
    "Monte espinoso", "Bosque seco", "Bosque humedo premontano",
    "Bosque humedo montano bajo", "Bosque humedo montano",
    "Bosque muy humedo premontano", "Bosque muy humedo montano",
    "Paramo / Jalca", "Puna",
]
FDT02_HELADAS = [
    "Frecuente (>30 dias/anio)", "Ocasional (10-30 dias/anio)",
    "Rara (<10 dias/anio)", "Ausente",
]
FDT02_VIENTOS = [
    "Calmo (< 2 m/s)", "Suave (2-4 m/s)", "Moderado (4-8 m/s)",
    "Fuerte (8-14 m/s)", "Muy fuerte (>14 m/s)",
]

# F-DT-03: CARACTERISTICAS DEL SUELO (Observacion de campo)
FDT03_TEXTURA = [
    "Arenoso", "Franco arenoso", "Franco", "Franco limoso",
    "Franco arcilloso", "Franco arcillo arenoso", "Arcilloso", "Limoso",
]
FDT03_COLOR = [
    "Negro / Muy oscuro", "Pardo oscuro", "Pardo / Marron",
    "Pardo claro / Amarillento", "Rojizo / Rojo amarillento",
    "Gris / Gris claro",
]
FDT03_PROFUNDIDAD = [
    "Muy superficial (<25 cm)", "Superficial (25-50 cm)",
    "Moderadamente profundo (50-100 cm)", "Profundo (>100 cm)",
]
FDT03_PEDREGOSIDAD = [
    "Sin piedras (0%)", "Pocas piedras (0-15%)",
    "Frecuentes (15-35%)", "Abundantes (35-60%)",
    "Muy pedregoso (>60%)",
]
FDT03_DRENAJE = [
    "Excesivo (suelo muy arenoso, seca rapido)",
    "Bueno (suelo drena adecuadamente)",
    "Moderado (drena con cierta lentitud)",
    "Imperfecto (retiene humedad excesiva)",
    "Pobre / Muy pobre (encharcamiento frecuente)",
]
FDT03_EROSION = [
    "Sin erosion aparente",
    "Erosion laminar leve",
    "Erosion laminar moderada a severa",
    "Erosion en surcos",
    "Erosion en carcavas",
    "Erosion mixta (laminar + surcos/carcavas)",
    "Movimientos en masa (deslizamientos)",
]
FDT03_MATERIA_ORGANICA = [
    "Muy baja (suelo claro, sin restos organicos)",
    "Baja (pocos restos, suelo claro)",
    "Media (presencia moderada de restos organicos)",
    "Alta (suelo oscuro, abundantes restos organicos)",
]

# F-DT-04: COBERTURA VEGETAL Y USO DEL SUELO
FDT04_TIPO_COBERTURA = [
    "Bosque denso (natural)", "Bosque ralo / Abierto",
    "Matorral / Arbustal", "Pastizal / Herbazal / Pajonal",
    "Cultivo agricola", "Plantacion forestal",
    "Suelo desnudo / Eriazo", "Vegetacion riberena",
    "Area urbana / Infraestructura",
]
FDT04_DENSIDAD = [
    "Muy rala (<10%)", "Rala (10-25%)", "Abierta (25-50%)",
    "Semicerrada (50-75%)", "Cerrada (>75%)",
]
FDT04_ESTADO_CONSERVACION = [
    "Bueno (sin intervencion significativa)",
    "Regular (intervencion parcial)",
    "Degradado (intervencion severa, con regeneracion)",
    "Muy degradado (sin regeneracion natural evidente)",
]
FDT04_USO_SUELO = [
    "Forestal / Proteccion", "Agricola (secano)",
    "Agricola (bajo riego)", "Pecuario / Pastoreo",
    "Agrosilvopastoril", "Minero",
    "Sin uso / En abandono", "Conservacion / Area protegida",
]
FDT04_CONFLICTO_USO = [
    "Sin conflicto (uso adecuado a capacidad)",
    "Sobreuso leve", "Sobreuso moderado", "Sobreuso severo",
    "Subuso (capacidad no aprovechada)",
]

# F-DT-05: RECURSOS HIDRICOS
FDT05_FUENTE_AGUA = [
    "Rio permanente", "Quebrada / Riachuelo", "Manantial / Puquio",
    "Laguna / Reservorio", "Canal de riego",
    "Agua subterranea (pozo)", "Ninguna visible en el area",
]
FDT05_REGIMEN = [
    "Permanente (flujo todo el anio)",
    "Estacional (flujo en temporada de lluvias)",
    "Temporal / Efimero (solo con eventos de lluvia)",
    "Sin escurrimiento superficial",
]
FDT05_CALIDAD_AGUA = [
    "Buena (clara, sin olor, sin sedimentos)",
    "Regular (ligeramente turbia o con sedimentos)",
    "Mala (turbia, con olor, con color)",
    "Muy mala (contaminacion evidente)",
    "No evaluable (sin fuente de agua accesible)",
]
FDT05_DISTANCIA_AGUA = [
    "<100 m", "100-500 m", "500-1000 m", "1-5 km", ">5 km",
]
FDT05_USO_HIDRICO = [
    "Consumo humano", "Riego agricola", "Pecuario / Abrevadero",
    "Piscicola", "Uso multiple", "Sin uso actual",
]

# F-DT-06: ASPECTOS SOCIOECONOMICOS Y ACCESIBILIDAD
FDT06_TENENCIA = [
    "Comunal (comunidad campesina)",
    "Privada individual (titulo)",
    "Privada individual (sin titulo / posesionario)",
    "Estatal / Fiscal",
    "Mixta (comunal + privada)",
    "Sin informacion",
]
FDT06_ORGANIZACION = [
    "Bien organizada (junta directiva activa, asambleas regulares)",
    "Moderadamente organizada (funcional pero irregular)",
    "Debilmente organizada (directiva nominal, poca participacion)",
    "Sin organizacion comunitaria identificada",
]
FDT06_ACTIVIDAD_ECONOMICA = [
    "Agricultura de subsistencia",
    "Agricultura comercial",
    "Ganaderia extensiva",
    "Ganaderia intensiva",
    "Actividad forestal / Extraccion",
    "Mineria artesanal",
    "Comercio / Servicios",
    "Mixta (agropecuaria)",
]
FDT06_ACCESIBILIDAD = [
    "Carretera asfaltada (acceso permanente)",
    "Carretera afirmada (acceso con restricciones en lluvia)",
    "Trocha carrozable (acceso limitado)",
    "Camino de herradura (solo a pie o acemila)",
    "Sin acceso vehicular (zona remota)",
]
FDT06_DISTANCIA_CENTRO = [
    "<1 km", "1-5 km", "5-10 km", "10-20 km", ">20 km",
]
FDT06_SERVICIOS = [
    "Agua potable", "Electricidad", "Telecomunicaciones / Internet",
    "Puesto de salud", "Escuela / IE", "Ninguno de los anteriores",
]

# ── Fichas de Diagnostico Social (F-DS-01 a F-DS-05) ─────────────────────
# Fuente: Plantilla_Diagnostico_Social_IN_Piura_rev_22abril.xlsx
FICHAS_DS = ["F-DS-01","F-DS-02","F-DS-03","F-DS-04","F-DS-05"]

# F-DS-01: DIAGNOSTICO SOCIOECONOMICO DE CENTRO POBLADO
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
FDS01_ACTIVIDADES_ECON = [
    "Agricultura", "Ganaderia", "Foresteria/Lena",
    "Comercio", "Jornales", "Artesania", "Otra",
]
FDS01_DESTINO_PRODUCCION = ["Autoconsumo", "Venta", "Autoconsumo/Venta"]
FDS01_PROBLEMAS_AGUA = [
    "Escasez", "Contaminacion", "Conflictos de uso",
    "Infraestructura deficiente", "Ninguno",
]
FDS01_USO_RECURSOS_FOREST = [
    "Lena", "Madera", "Productos forestales no maderables", "No usa",
]
FDS01_DISPOSICION = ["Alta", "Media", "Baja", "Condicionada"]

# F-DS-02: IDENTIFICACION Y CARACTERIZACION DE ACTORES CLAVE
FDS02_TIPO_ACTOR = ["Publico", "Privado", "Soc. Civil"]
FDS02_NIVEL = ["A", "M", "B"]  # Alto / Medio / Bajo
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

# F-DS-05: IDENTIFICACION DE CONFLICTOS Y OPORTUNIDADES
FDS05_NIVEL = ["A", "M", "B"]
FDS05_ESTADO_CONFLICTO = ["Activo", "Latente"]
FDS05_TIPO_OPORTUNIDAD = ["Social", "Institucional", "Productivo"]

# ── CSS ───────────────────────────────────────────────────────────────────
st.markdown("""<style>
.main-header{background:#2C3E50;padding:1rem 2rem;border-radius:.5rem;margin-bottom:1rem}
.main-header h1{color:#fff!important;margin:0!important;font-size:1.8rem!important}
.main-header p{color:#BDC3C7!important;margin:0!important}
.edit-mode-banner{background:linear-gradient(90deg,#f39c12,#e67e22);color:#fff;padding:.6rem 1.2rem;
    border-radius:.4rem;margin-bottom:.8rem;font-weight:600;display:flex;align-items:center;gap:.5rem}
.edit-mode-banner .icon{font-size:1.2rem}
.dup-warning{background:#fff3cd;border-left:4px solid #ffc107;padding:.8rem 1rem;border-radius:0 .4rem .4rem 0;margin:.5rem 0}
</style>""", unsafe_allow_html=True)

st.markdown("""<div class="main-header">
<h1>\U0001F331 IN Piura</h1>
<p>Plan de Ingreso | Verificacion de Campo | Cuenca Alta del Rio Piura</p>
</div>""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────
pagina = st.sidebar.selectbox("Navegacion", [
    "Panel de Control","Bloques de Intervencion","Inspeccion de Campo",
    "Indicadores de Calidad","Diagnostico Territorial","Diagnostico Social",
    "Elementos Expuestos (AdR)",
    "Presupuesto","Cronograma",
    "Georreferenciacion","ODK / KoBoToolbox","Reportes",
    "Conversor PDF -> Excel",
    "⚙️ Migracion V4 (temporal)",
])
st.sidebar.markdown("---")
st.sidebar.markdown("**IN Piura** v2.1 Web\n\nRestauracion de Ecosistemas\nCuenca Alta del Rio Piura")

# ── Helpers ───────────────────────────────────────────────────────────────
def _bloques_map():
    return {f"{b['codigo']} - {b['tipo_intervencion']}": b["id"]
            for b in _cached_obtener_bloques(_cache_version())}

def _distritos(prov):
    return PROVINCIAS_DISTRITOS.get(prov, DISTRITOS_PIURA) if prov else DISTRITOS_PIURA

def _resolver_microcuenca(bloque_label):
    """Resuelve la microcuenca para un bloque dado su label 'CODIGO - TIPO'.
    Busca primero en la BD (cacheada) y luego en BLOQUES_124_MAP como fallback."""
    codigo = bloque_label.split(" - ")[0].strip() if " - " in bloque_label else bloque_label.strip()
    # Buscar en BD (cacheada)
    for b in _cached_obtener_bloques(_cache_version()):
        if b["codigo"] == codigo:
            mc = b.get("microcuenca", "") or ""
            if mc and mc in MICROCUENCAS:
                return mc
            break
    # Fallback: buscar en los 128 bloques predefinidos
    datos = BLOQUES_124_MAP.get(codigo, {})
    mc = datos.get("microcuenca", "")
    if mc and mc in MICROCUENCAS:
        return mc
    return ""

# ══════════════════════════════════════════════════════════════════════════
# PANEL DE CONTROL
# ══════════════════════════════════════════════════════════════════════════
def pagina_dashboard():
    st.subheader("Panel de Control - Resumen Ejecutivo")
    stats = _cached_obtener_estadisticas(_cache_version())
    c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
    c1.metric("Total Bloques", stats["total_bloques"])
    c2.metric("Area Total", f"{stats['area_total_ha']:.2f} ha")
    c3.metric("Inspecciones", stats["total_inspecciones"])
    c4.metric("Avance Promedio", f"{stats['avance_promedio']:.1f}%")
    c5.metric("Diag. Territorial", stats.get("total_diagnosticos", 0))
    c6.metric("Diag. Social", stats.get("total_diagnosticos_sociales", 0))
    c7.metric("Personal Activo", stats["personal_activo"])
    st.markdown("---")
    ci,cd = st.columns(2)
    with ci:
        st.markdown("**Distribucion por Estado**")
        tb = max(stats["total_bloques"],1)
        for e in ["Pendiente","En progreso","Verificado"]:
            n = stats["bloques_por_estado"].get(e,0)
            st.progress(n/tb, text=f"{e}: {n} ({n/tb*100:.1f}%)")
    with cd:
        st.markdown("**Distribucion por Tipo**")
        if stats["bloques_por_tipo"]:
            df = pd.DataFrame(list(stats["bloques_por_tipo"].items()), columns=["Tipo","Cantidad"])
            st.bar_chart(df.set_index("Tipo"))
        else:
            st.info("Sin bloques.")
    st.markdown("---")
    cp,cc = st.columns(2)
    with cp:
        st.markdown("**Resumen Presupuestal**")
        pl,ej = stats["presupuesto_planificado"],stats["presupuesto_ejecutado"]
        pe = (ej/pl*100) if pl>0 else 0
        st.metric("Planificado",f"S/ {pl:,.2f}"); st.metric("Ejecutado",f"S/ {ej:,.2f}")
        st.progress(min(pe/100,1.0), text=f"Ejecucion: {pe:.1f}%")
        st.caption(f"Saldo: S/ {pl-ej:,.2f}")
    with cc:
        st.markdown("**Cronograma**")
        ae = stats["actividades_por_estado"]; ta = sum(ae.values()) if ae else 0
        st.caption(f"Total actividades: {ta}")
        for en in ["Programado","En ejecucion","Completado","Retrasado"]:
            cn = ae.get(en,0); st.progress((cn/ta) if ta>0 else 0, text=f"{en}: {cn}")
    st.markdown("---")
    st.markdown("**Resumen de Bloques**")
    res = _cached_obtener_resumen_bloques(_cache_version())
    if res:
        st.dataframe(pd.DataFrame([{"Codigo":b["codigo"],"Tipo":b["tipo_intervencion"],
            "Distrito":b["distrito"],"Area (ha)":f"{b['area_hectareas']:.4f}",
            "Estado":b["estado"],"Avance %":f"{(b.get('ultimo_avance') or 0):.1f}",
            "Inspecciones":b.get("total_inspecciones",0)} for b in res]),
            use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════
# BLOQUES DE INTERVENCION
# ══════════════════════════════════════════════════════════════════════════
def _extraer_codigo_bloque_128(opcion):
    """Retorna el codigo de bloque directamente (el dropdown solo muestra codigos)."""
    if not opcion:
        return None
    return opcion.strip()

def _bl_load_edit(bloque):
    """Carga un registro de bloque en session_state para edicion."""
    st.session_state["bl_edit_id"] = bloque["id"]
    st.session_state["bl_codigo"] = bloque["codigo"]
    st.session_state["bl_cuenca"] = bloque.get("cuenca", "Cuenca Alta del Rio Piura") or "Cuenca Alta del Rio Piura"
    st.session_state["bl_microcuenca"] = bloque.get("microcuenca", "") or ""
    st.session_state["bl_provincia"] = bloque.get("provincia", "") or ""
    st.session_state["bl_distrito"] = bloque.get("distrito", "") or ""
    st.session_state["bl_tipo"] = bloque.get("tipo_intervencion", "") or ""
    st.session_state["bl_utm_este"] = str(bloque.get("utm_este", 0) or 0)
    st.session_state["bl_utm_norte"] = str(bloque.get("utm_norte", 0) or 0)
    st.session_state["bl_utm_zona"] = bloque.get("utm_zona", "17S") or "17S"
    st.session_state["bl_altitud"] = str(bloque.get("altitud", 0) or 0)
    st.session_state["bl_area"] = str(bloque.get("area_hectareas", 0) or 0)
    st.session_state["bl_responsable"] = bloque.get("responsable", "") or ""
    st.session_state["bl_estado"] = bloque.get("estado", "Pendiente") or "Pendiente"

def pagina_bloques():
    st.subheader("Bloques de Intervencion")

    # Inicializar estado de edicion
    if "bl_edit_id" not in st.session_state:
        st.session_state["bl_edit_id"] = None

    edit_id = st.session_state.get("bl_edit_id")

    cf,ct = st.columns([1,2])
    with cf:
        if edit_id:
            st.markdown('<div class="edit-mode-banner"><span class="icon">&#9998;</span> '
                        f'Modo Edicion - Bloque ID {edit_id}</div>', unsafe_allow_html=True)
            if st.button("Cancelar edicion", key="bl_cancel_edit"):
                st.session_state["bl_edit_id"] = None
                for k in list(st.session_state.keys()):
                    if k.startswith("bl_") and k != "bl_edit_id":
                        del st.session_state[k]
                st.rerun()
        else:
            st.markdown("**Registro de Bloque**")

        # Selector rapido de los 128 bloques V4 (solo en modo nuevo)
        if not edit_id:
            st.markdown("##### Seleccion rapida - 128 Bloques de Intervencion V4")
            sel_79 = st.selectbox(
                "Seleccionar bloque predefinido",
                ["(Seleccionar bloque predefinido)"] + BLOQUES_124_OPCIONES,
                key="sel_bloque_128",
                help="Seleccione un bloque de la lista de 128 bloques V4 para autocompletar los campos"
            )

            # Determinar valores por defecto segun seleccion
            cod_sel = _extraer_codigo_bloque_128(sel_79 if sel_79 != "(Seleccionar bloque predefinido)" else "")
            datos_79 = BLOQUES_124_MAP.get(cod_sel, {}) if cod_sel else {}

            def_codigo = datos_79.get("codigo", "")
            def_microcuenca = datos_79.get("microcuenca", "")
            def_area = str(datos_79.get("area_ha", "0"))
            def_provincia = datos_79.get("provincia", "")
            def_distrito = datos_79.get("distrito", "")
            def_accesibilidad = datos_79.get("accesibilidad", 0)
            def_dia = datos_79.get("dia_evaluacion", 0)

            if datos_79:
                acc_txt = "Acceso limitado" if def_accesibilidad == 1 else "Acceso normal"
                cp_data = CENTROS_POBLADOS_BLOQUE.get(def_codigo, {})
                cp_txt = ", ".join(cp_data.get("centros_poblados", [])) if cp_data else "—"
                cc_txt = ", ".join(cp_data.get("comunidades_campesinas", [])) if cp_data else "—"
                st.info(f"Bloque **{def_codigo}** | Microcuenca: {def_microcuenca} | "
                        f"{def_distrito} ({def_provincia}) | {def_area} ha | "
                        f"{acc_txt} | Dia eval.: {def_dia}\n\n"
                        f"CC.PP.: {cp_txt} | Com. Campesinas: {cc_txt} | "
                        f"UTM: {datos_79.get('utm_este', 0)} E, {datos_79.get('utm_norte', 0)} N")
        else:
            datos_79 = {}
            def_codigo = st.session_state.get("bl_codigo", "")
            def_microcuenca = st.session_state.get("bl_microcuenca", "")
            def_area = st.session_state.get("bl_area", "0")
            def_provincia = st.session_state.get("bl_provincia", "")
            def_distrito = st.session_state.get("bl_distrito", "")

        st.markdown("---")
        with st.form("form_bloque", clear_on_submit=False):
            codigo = st.text_input("Codigo de bloque",
                value=st.session_state.get("bl_codigo", def_codigo) if edit_id else def_codigo)
            cuenca = st.text_input("Cuenca",
                value=st.session_state.get("bl_cuenca", "Cuenca Alta del Rio Piura") if edit_id else "Cuenca Alta del Rio Piura")
            # Microcuenca: preseleccionar
            mc_idx = 0
            mc_val = st.session_state.get("bl_microcuenca", def_microcuenca) if edit_id else def_microcuenca
            if mc_val and mc_val in MICROCUENCAS:
                mc_idx = MICROCUENCAS.index(mc_val) + 1
            microcuenca = st.selectbox("Microcuenca", [""]+MICROCUENCAS, index=mc_idx)
            # Provincia: preseleccionar
            prov_idx = 0
            prov_val = st.session_state.get("bl_provincia", def_provincia) if edit_id else def_provincia
            if prov_val and prov_val in PROVINCIAS:
                prov_idx = PROVINCIAS.index(prov_val) + 1
            provincia = st.selectbox("Provincia", [""]+PROVINCIAS, index=prov_idx)
            # Distrito: preseleccionar
            dist_list = _distritos(provincia)
            dist_idx = 0
            dist_val = st.session_state.get("bl_distrito", def_distrito) if edit_id else def_distrito
            if dist_val and dist_val in dist_list:
                dist_idx = dist_list.index(dist_val) + 1
            distrito = st.selectbox("Distrito", [""]+dist_list, index=dist_idx)
            # Tipo intervencion
            tipo_idx = 0
            if edit_id:
                tipo_val = st.session_state.get("bl_tipo", "")
                if tipo_val and tipo_val in TIPOS_INTERVENCION:
                    tipo_idx = TIPOS_INTERVENCION.index(tipo_val)
            tipo = st.selectbox("Tipo intervencion", TIPOS_INTERVENCION, index=tipo_idx)
            a1,a2 = st.columns(2)
            def_ue = str(datos_79.get("utm_este", 0)) if datos_79 else "0"
            def_un = str(datos_79.get("utm_norte", 0)) if datos_79 else "0"
            ue = a1.text_input("UTM Este", st.session_state.get("bl_utm_este", def_ue) if edit_id else def_ue)
            un = a2.text_input("UTM Norte", st.session_state.get("bl_utm_norte", def_un) if edit_id else def_un)
            b1,b2 = st.columns(2)
            uz = b1.text_input("Zona UTM", st.session_state.get("bl_utm_zona", "17S") if edit_id else "17S")
            alt = b2.text_input("Altitud", st.session_state.get("bl_altitud", "0") if edit_id else "0")
            area = st.text_input("Area (ha)",
                value=st.session_state.get("bl_area", def_area) if edit_id else (def_area if datos_79 else "0"))
            resp = st.text_input("Responsable",
                value=st.session_state.get("bl_responsable", "") if edit_id else "")
            # Estado: preseleccionar
            est_idx = 0
            if edit_id:
                est_val = st.session_state.get("bl_estado", "Pendiente")
                if est_val and est_val in ESTADOS_BLOQUE:
                    est_idx = ESTADOS_BLOQUE.index(est_val)
            estado = st.selectbox("Estado", ESTADOS_BLOQUE, index=est_idx)
            btn_label = "Actualizar Bloque" if edit_id else "Guardar"
            guardar = st.form_submit_button(btn_label, type="primary")
        if guardar:
            if not codigo: st.warning("Codigo obligatorio.")
            elif not distrito: st.warning("Seleccione distrito.")
            else:
                try:
                    if edit_id:
                        db.actualizar_bloque(bloque_id=edit_id,codigo=codigo,
                            tipo_intervencion=tipo,cuenca=cuenca,
                            distrito=distrito,utm_este=float(ue),utm_norte=float(un),
                            utm_zona=uz,area_hectareas=float(area),estado=estado,
                            altitud=float(alt or 0),responsable=resp,
                            microcuenca=microcuenca,provincia=provincia)
                        _invalidar_cache()
                        st.session_state["bl_edit_id"] = None
                        for k in list(st.session_state.keys()):
                            if k.startswith("bl_") and k != "bl_edit_id":
                                del st.session_state[k]
                        st.success(f"Bloque {codigo} actualizado correctamente.")
                    else:
                        db.insertar_bloque(codigo=codigo,tipo_intervencion=tipo,cuenca=cuenca,
                            distrito=distrito,utm_este=float(ue),utm_norte=float(un),
                            utm_zona=uz,area_hectareas=float(area),estado=estado,
                            altitud=float(alt or 0),responsable=resp,
                            microcuenca=microcuenca,provincia=provincia)
                        _invalidar_cache()
                        st.success(f"Bloque {codigo} registrado.")
                    st.rerun()
                except Exception as e: st.error(f"Error: {e}")
    with ct:
        st.markdown("**Bloques Registrados**")
        st.caption("Haga clic en **Editar** para modificar un bloque existente y evitar duplicidades.")
        busq = st.text_input("Buscar","",key="busq_bl")
        bloques = db.buscar_bloques(busq) if busq else _cached_obtener_bloques(_cache_version())
        if bloques:
            # Paginacion
            bloques_pag, total_pags, pag_actual = _paginar(bloques, "pag_bloques")
            _controles_paginacion(total_pags, pag_actual, "pag_bloques")
            # Tabla con botones de edicion por fila
            header_cols = st.columns([0.5, 1.2, 1.2, 1.2, 1, 1, 0.8, 0.8, 0.7])
            headers = ["ID", "Codigo", "Microcuenca", "Tipo", "Provincia", "Distrito", "Area", "Estado", ""]
            for col, h in zip(header_cols, headers):
                col.markdown(f"**{h}**")
            st.markdown("---")
            for b in bloques_pag:
                row_cols = st.columns([0.5, 1.2, 1.2, 1.2, 1, 1, 0.8, 0.8, 0.7])
                row_cols[0].write(b["id"])
                row_cols[1].write(b["codigo"])
                row_cols[2].write(b.get("microcuenca", "") or "")
                row_cols[3].write(b["tipo_intervencion"])
                row_cols[4].write(b.get("provincia", "") or "")
                row_cols[5].write(b["distrito"])
                row_cols[6].write(f"{b['area_hectareas']:.4f}")
                row_cols[7].write(b["estado"])
                if row_cols[8].button("Editar", key=f"edit_bl_{b['id']}", type="primary"):
                    _bl_load_edit(b)
                    st.rerun()
            st.markdown("---")
            bm = {f"{b['codigo']} - {b['tipo_intervencion']}":b["id"] for b in bloques}
            sel = st.selectbox("Seleccionar bloque para eliminar",[""]+list(bm.keys()),key="del_bl")
            if sel and sel in bm and st.button("Eliminar bloque", key="btn_del_bl"):
                try:
                    db.eliminar_bloque(bm[sel])
                    _invalidar_cache()
                    st.success(f"Bloque {sel} eliminado correctamente.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al eliminar bloque: {e}")
        else: st.info("Sin bloques.")

        st.markdown("---")
        with st.expander("Tabla de Referencia - 128 Bloques de Intervencion V4", expanded=False):
            st.caption("Fuente: Reporte_Bloques_V5_25abril.xlsx - Base de datos completa del proyecto IN Piura (124 bloques)")
            df_128 = pd.DataFrame([{
                "N":b[0], "Bloque":b[1], "Microcuenca":b[2],
                "Area (ha)":b[3], "Provincia":b[4], "Distrito":b[5],
                "MSAVI 2024":b[10],
            } for b in BLOQUES_124])
            st.dataframe(df_128, use_container_width=True, hide_index=True, height=400)

# ══════════════════════════════════════════════════════════════════════════
# INSPECCION DE CAMPO
# ══════════════════════════════════════════════════════════════════════════
def pagina_inspeccion():
    st.subheader("Inspeccion de Campo")
    bm = _bloques_map()
    if not bm: st.warning("Registre un bloque primero."); return

    # Inicializar estado de edicion
    if "insp_edit_id" not in st.session_state:
        st.session_state["insp_edit_id"] = None

    edit_id = st.session_state.get("insp_edit_id")

    if edit_id:
        st.markdown('<div class="edit-mode-banner"><span class="icon">&#9998;</span> '
                    f'Modo Edicion - Inspeccion ID {edit_id}</div>', unsafe_allow_html=True)
        if st.button("Cancelar edicion", key="insp_cancel_edit"):
            st.session_state["insp_edit_id"] = None
            for k in list(st.session_state.keys()):
                if k.startswith("insp_e_"):
                    del st.session_state[k]
            st.rerun()

    # Selector de bloque FUERA del form para auto-enlazar microcuenca
    opciones_bloque = list(bm.keys())
    # En modo edicion, preseleccionar el bloque
    bl_idx = 0
    if edit_id:
        bl_code = st.session_state.get("insp_e_bloque", "")
        for i, op in enumerate(opciones_bloque):
            if bl_code and bl_code in op:
                bl_idx = i
                break
    bl = st.selectbox("Bloque", opciones_bloque, index=bl_idx, key="insp_bloque")

    # Auto-resolver microcuenca del bloque seleccionado
    mc_auto = _resolver_microcuenca(bl)
    if mc_auto:
        mc_idx = MICROCUENCAS.index(mc_auto) + 1
        st.info(f"Microcuenca vinculada automaticamente: **{mc_auto}**")
    else:
        mc_idx = 0
    if edit_id:
        mc_val = st.session_state.get("insp_e_mc", "")
        if mc_val and mc_val in MICROCUENCAS:
            mc_idx = MICROCUENCAS.index(mc_val) + 1

    # Carga de archivos PDF (fuera del form por limitaciones de Streamlit)
    if not edit_id:
        pdf_files = st.file_uploader(
            "Adjuntar archivos PDF (max. 25 MB por archivo)",
            type=["pdf"], accept_multiple_files=True, key="insp_pdf_upload")
        if pdf_files:
            for f in pdf_files:
                if f.size > 25 * 1024 * 1024:
                    st.warning(f"El archivo '{f.name}' excede 25 MB y no sera adjuntado.")
            st.info(f"{len(pdf_files)} archivo(s) PDF seleccionado(s)")
    else:
        pdf_files = None

    # Valores de edicion
    def_fecha = datetime.now()
    def_inspector = ""
    def_clima_idx = 0
    def_avance = 0.0
    def_obs = ""
    def_desv = ""
    def_ver = f"VER-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}"
    if edit_id:
        def_inspector = st.session_state.get("insp_e_inspector", "")
        def_clima = st.session_state.get("insp_e_clima", "")
        if def_clima and def_clima in CONDICIONES_CLIMATICAS:
            def_clima_idx = CONDICIONES_CLIMATICAS.index(def_clima)
        def_avance = float(st.session_state.get("insp_e_avance", 0))
        def_obs = st.session_state.get("insp_e_obs", "")
        def_desv = st.session_state.get("insp_e_desv", "")
        def_ver = st.session_state.get("insp_e_ver", def_ver)
        try:
            def_fecha = datetime.strptime(st.session_state.get("insp_e_fecha", ""), "%Y-%m-%d")
        except (ValueError, TypeError):
            pass

    with st.form("form_insp", clear_on_submit=not edit_id):
        mc = st.selectbox("Microcuenca", [""] + MICROCUENCAS, index=mc_idx)
        fecha = st.date_input("Fecha de visita", value=def_fecha,
                              min_value=FECHA_MIN_PROYECTO, max_value=date.today(),
                              help="Seleccione la fecha de la visita de campo")
        inspector = st.text_input("Inspector", value=def_inspector)
        clima = st.selectbox("Condiciones climaticas", CONDICIONES_CLIMATICAS, index=def_clima_idx)
        avance = st.number_input("Avance fisico (%)", 0.0, 100.0, def_avance)
        obs = st.text_area("Observaciones tecnicas", value=def_obs)
        desv = st.text_area("Desviaciones observadas al Plan de Trabajo", value=def_desv)
        ver = st.text_input("Codigo de verificacion", value=def_ver)
        btn_label = "Actualizar Inspeccion" if edit_id else "Guardar Inspeccion"
        guardar = st.form_submit_button(btn_label, type="primary")
    if guardar:
        if not inspector: st.warning("Inspector obligatorio.")
        else:
            try:
                if edit_id:
                    db.actualizar_inspeccion(inspeccion_id=edit_id,
                        fecha_visita=fecha.strftime("%Y-%m-%d"),
                        inspector=inspector, condiciones_climaticas=clima,
                        avance_fisico=avance, observaciones=obs, desviaciones=desv,
                        codigo_verificacion=ver, microcuenca=mc)
                    _invalidar_cache()
                    st.session_state["insp_edit_id"] = None
                    for k in list(st.session_state.keys()):
                        if k.startswith("insp_e_"):
                            del st.session_state[k]
                    st.success("Inspeccion actualizada."); st.rerun()
                else:
                    # Verificar duplicados antes de insertar
                    existentes = db.obtener_inspecciones_por_bloque(bm[bl])
                    dup = [e for e in existentes if e["fecha_visita"] == fecha.strftime("%Y-%m-%d")
                           and e["inspector"] == inspector]
                    if dup:
                        st.warning(f"Ya existe una inspeccion para este bloque en {fecha.strftime('%Y-%m-%d')} "
                                   f"por {inspector}. Use 'Editar' en el historial para modificarla.")
                    else:
                        # Guardar archivos PDF adjuntos
                        rutas_pdf = []
                        if pdf_files:
                            pdf_dir = os.path.join(os.path.dirname(__file__), "adjuntos_pdf")
                            os.makedirs(pdf_dir, exist_ok=True)
                            for f in pdf_files:
                                if f.size <= 25 * 1024 * 1024:
                                    nombre_pdf = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{f.name}"
                                    ruta_pdf = os.path.join(pdf_dir, nombre_pdf)
                                    with open(ruta_pdf, "wb") as fp:
                                        fp.write(f.getbuffer())
                                    rutas_pdf.append(ruta_pdf)
                        archivos_pdf_str = ";".join(rutas_pdf) if rutas_pdf else ""

                        db.insertar_inspeccion(bloque_id=bm[bl],fecha_visita=fecha.strftime("%Y-%m-%d"),
                            inspector=inspector,condiciones_climaticas=clima,avance_fisico=avance,
                            observaciones=obs,desviaciones=desv,registro_fotografico="",
                            codigo_verificacion=ver,microcuenca=mc,archivos_pdf=archivos_pdf_str)
                        _invalidar_cache()
                        msg = "Inspeccion registrada."
                        if rutas_pdf:
                            msg += f" {len(rutas_pdf)} PDF(s) adjuntado(s)."
                        st.success(msg); st.rerun()
            except Exception as e: st.error(f"Error: {e}")
    st.markdown("---")
    st.markdown("**Historial de Inspecciones**")
    st.caption("Haga clic en **Editar** para modificar una inspeccion existente y evitar duplicidades.")
    insp = _cached_obtener_todas_inspecciones(_cache_version())
    if insp:
        # Paginacion
        insp_pag, total_pags_i, pag_actual_i = _paginar(insp, "pag_insp")
        _controles_paginacion(total_pags_i, pag_actual_i, "pag_insp")
        # Tabla con botones de edicion
        header_cols = st.columns([0.5, 1, 1, 0.9, 0.9, 0.7, 1, 0.5, 0.6])
        for col, h in zip(header_cols, ["ID", "Bloque", "Microcuenca", "Fecha", "Inspector", "Avance%", "Verificacion", "PDFs", ""]):
            col.markdown(f"**{h}**")
        st.markdown("---")
        for i in insp_pag:
            row = st.columns([0.5, 1, 1, 0.9, 0.9, 0.7, 1, 0.5, 0.6])
            row[0].write(i["id"])
            row[1].write(i["bloque_codigo"])
            row[2].write(i.get("microcuenca", "") or "")
            row[3].write(i["fecha_visita"])
            row[4].write(i["inspector"])
            row[5].write(f"{i['avance_fisico']:.1f}")
            row[6].write(i["codigo_verificacion"])
            row[7].write(len([p for p in (i.get("archivos_pdf", "") or "").split(";") if p.strip()]))
            if row[8].button("Editar", key=f"edit_insp_{i['id']}", type="primary"):
                st.session_state["insp_edit_id"] = i["id"]
                st.session_state["insp_e_bloque"] = i["bloque_codigo"]
                st.session_state["insp_e_mc"] = i.get("microcuenca", "") or ""
                st.session_state["insp_e_fecha"] = i["fecha_visita"]
                st.session_state["insp_e_inspector"] = i["inspector"]
                st.session_state["insp_e_clima"] = i["condiciones_climaticas"]
                st.session_state["insp_e_avance"] = i["avance_fisico"]
                st.session_state["insp_e_obs"] = i.get("observaciones", "") or ""
                st.session_state["insp_e_desv"] = i.get("desviaciones", "") or ""
                st.session_state["insp_e_ver"] = i["codigo_verificacion"]
                st.rerun()
        st.markdown("---")

        # Eliminar inspeccion
        insp_del_map = {f"ID {i['id']} - {i['fecha_visita']} - {i['inspector']}": i["id"] for i in insp}
        sel_del = st.selectbox("Seleccionar inspeccion a eliminar", [""] + list(insp_del_map.keys()), key="del_insp")
        if sel_del and sel_del in insp_del_map and st.button("Eliminar inspeccion", key="btn_del_insp"):
            db.eliminar_inspeccion(insp_del_map[sel_del]); _invalidar_cache(); st.success("Inspeccion eliminada."); st.rerun()

        # Seccion para descargar PDFs adjuntos de una inspeccion
        st.markdown("**Descargar PDFs adjuntos**")
        insp_map = {f"ID {i['id']} - {i['fecha_visita']} - {i['inspector']}": i for i in insp}
        sel_insp = st.selectbox("Seleccionar inspeccion", list(insp_map.keys()), key="pdf_download_sel")
        if sel_insp:
            insp_sel = insp_map[sel_insp]
            pdfs_str = insp_sel.get("archivos_pdf", "") or ""
            pdfs = [p.strip() for p in pdfs_str.split(";") if p.strip()]
            if pdfs:
                for ruta in pdfs:
                    nombre = os.path.basename(ruta)
                    if os.path.exists(ruta):
                        with open(ruta, "rb") as fp:
                            st.download_button(
                                label=f"Descargar: {nombre}",
                                data=fp.read(),
                                file_name=nombre,
                                mime="application/pdf",
                                key=f"dl_{nombre}")
                    else:
                        st.caption(f"Archivo no disponible: {nombre}")
            else:
                st.caption("Esta inspeccion no tiene archivos PDF adjuntos.")

        # Adjuntar PDFs a inspeccion existente
        st.markdown("**Adjuntar PDF a inspeccion existente**")
        pdf_adicional = st.file_uploader(
            "Seleccionar PDF (max. 25 MB)",
            type=["pdf"], accept_multiple_files=True, key="pdf_hist_upload")
        if pdf_adicional and st.button("Adjuntar a inspeccion seleccionada", key="btn_adjuntar_pdf"):
            insp_sel = insp_map[sel_insp]
            pdf_dir = os.path.join(os.path.dirname(__file__), "adjuntos_pdf")
            os.makedirs(pdf_dir, exist_ok=True)
            nuevas_rutas = []
            for f in pdf_adicional:
                if f.size <= 25 * 1024 * 1024:
                    nombre_pdf = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{f.name}"
                    ruta_pdf = os.path.join(pdf_dir, nombre_pdf)
                    with open(ruta_pdf, "wb") as fp:
                        fp.write(f.getbuffer())
                    nuevas_rutas.append(ruta_pdf)
                else:
                    st.warning(f"'{f.name}' excede 25 MB.")
            if nuevas_rutas:
                existentes = insp_sel.get("archivos_pdf", "") or ""
                todas = [p.strip() for p in existentes.split(";") if p.strip()] + nuevas_rutas
                db.actualizar_archivos_pdf_inspeccion(insp_sel["id"], ";".join(todas))
                _invalidar_cache()
                st.success(f"{len(nuevas_rutas)} PDF(s) adjuntado(s)."); st.rerun()

# ══════════════════════════════════════════════════════════════════════════
# INDICADORES DE CALIDAD
# ══════════════════════════════════════════════════════════════════════════
def pagina_indicadores():
    st.subheader("Indicadores de Calidad")
    bm = _bloques_map()
    if not bm: st.warning("Registre un bloque primero."); return

    # Inicializar estado de edicion
    if "ind_edit_id" not in st.session_state:
        st.session_state["ind_edit_id"] = None

    edit_id = st.session_state.get("ind_edit_id")

    if edit_id:
        st.markdown('<div class="edit-mode-banner"><span class="icon">&#9998;</span> '
                    f'Modo Edicion - Indicador ID {edit_id}</div>', unsafe_allow_html=True)
        if st.button("Cancelar edicion", key="ind_cancel_edit"):
            st.session_state["ind_edit_id"] = None
            for k in list(st.session_state.keys()):
                if k.startswith("ind_e_"):
                    del st.session_state[k]
            st.rerun()

    bl = st.selectbox("Bloque", list(bm.keys()), key="ind_bl")
    bid = bm[bl]

    # Auto-resolver microcuenca del bloque seleccionado
    mc_auto = _resolver_microcuenca(bl)
    if mc_auto:
        mc_idx = MICROCUENCAS.index(mc_auto) + 1
        st.info(f"Microcuenca vinculada automaticamente: **{mc_auto}**")
    else:
        mc_idx = 0
    if edit_id:
        mc_val = st.session_state.get("ind_e_mc", "")
        if mc_val and mc_val in MICROCUENCAS:
            mc_idx = MICROCUENCAS.index(mc_val) + 1

    ins = db.obtener_inspecciones_por_bloque(bid)
    if not ins: st.warning("Sin inspecciones para este bloque."); return
    im = {f"{i['fecha_visita']} - {i['inspector']}":i["id"] for i in ins}
    isel = st.selectbox("Inspeccion", list(im.keys()))

    def_pc = float(st.session_state.get("ind_e_pc", 0)) if edit_id else 0.0
    def_tc = st.session_state.get("ind_e_tc", "") if edit_id else ""
    def_vi = st.session_state.get("ind_e_vi", "") if edit_id else ""
    tc_idx = ([""]+TIPOS_COBERTURA).index(def_tc) if def_tc in ([""]+TIPOS_COBERTURA) else 0
    vi_idx = ([""]+VIGOR_COBERTURA).index(def_vi) if def_vi in ([""]+VIGOR_COBERTURA) else 0

    with st.form("form_ind", clear_on_submit=not edit_id):
        mc = st.selectbox("Microcuenca", [""] + MICROCUENCAS, index=mc_idx, key="ind_mc")
        pc = st.number_input("Cobertura vegetal (%)", 0.0, 100.0, def_pc)
        tc = st.selectbox("Tipo cobertura", [""]+TIPOS_COBERTURA, index=tc_idx)
        vi = st.selectbox("Vigor cobertura", [""]+VIGOR_COBERTURA, index=vi_idx)
        btn_label = "Actualizar Indicadores" if edit_id else "Guardar Indicadores"
        guardar = st.form_submit_button(btn_label, type="primary")
    if guardar:
        try:
            if edit_id:
                db.actualizar_indicadores(indicador_id=edit_id,
                    porcentaje_cobertura_vegetal=pc,
                    tipo_cobertura_vegetal=tc, vigor_cobertura_vegetal=vi,
                    microcuenca=mc)
                _invalidar_cache()
                st.session_state["ind_edit_id"] = None
                for k in list(st.session_state.keys()):
                    if k.startswith("ind_e_"):
                        del st.session_state[k]
                st.success("Indicadores actualizados."); st.rerun()
            else:
                # Verificar duplicado: un indicador por inspeccion
                existente = db.obtener_indicadores_por_inspeccion(im[isel])
                if existente:
                    st.warning("Ya existen indicadores para esta inspeccion. Use 'Editar' en la tabla para modificarlos.")
                else:
                    db.insertar_indicadores(bloque_id=bid,inspeccion_id=im[isel],
                        cobertura_vegetal_planificada=0,cobertura_vegetal_lograda=0,
                        sobrevivencia_especies=0,longitud_zanjas_ejecutada=0,
                        volumen_retencion_sedimentos=0,porcentaje_cobertura_vegetal=pc,
                        tipo_cobertura_vegetal=tc,vigor_cobertura_vegetal=vi,microcuenca=mc)
                    _invalidar_cache()
                    st.success("Indicadores guardados."); st.rerun()
        except Exception as e: st.error(f"Error: {e}")
    st.markdown("---")
    st.markdown("**Indicadores Registrados**")
    st.caption("Haga clic en **Editar** para modificar indicadores existentes.")
    ind = db.obtener_indicadores_por_bloque(bid)
    if ind:
        header_cols = st.columns([1, 1, 1, 1, 1, 0.6])
        for col, h in zip(header_cols, ["Fecha", "Microcuenca", "Cobert.%", "Tipo", "Vigor", ""]):
            col.markdown(f"**{h}**")
        st.markdown("---")
        for x in ind:
            row = st.columns([1, 1, 1, 1, 1, 0.6])
            row[0].write(x.get("fecha_visita", ""))
            row[1].write(x.get("microcuenca", "") or "")
            row[2].write(f"{x.get('porcentaje_cobertura_vegetal', 0):.1f}")
            row[3].write(x.get("tipo_cobertura_vegetal", "") or "")
            row[4].write(x.get("vigor_cobertura_vegetal", "") or "")
            if row[5].button("Editar", key=f"edit_ind_{x['id']}", type="primary"):
                st.session_state["ind_edit_id"] = x["id"]
                st.session_state["ind_e_mc"] = x.get("microcuenca", "") or ""
                st.session_state["ind_e_pc"] = x.get("porcentaje_cobertura_vegetal", 0)
                st.session_state["ind_e_tc"] = x.get("tipo_cobertura_vegetal", "") or ""
                st.session_state["ind_e_vi"] = x.get("vigor_cobertura_vegetal", "") or ""
                st.rerun()
        st.markdown("---")
        # Eliminar indicador
        ind_del = {f"ID {x['id']} - {x.get('fecha_visita', '')}": x["id"] for x in ind}
        sel_del = st.selectbox("Seleccionar indicador a eliminar", [""] + list(ind_del.keys()), key="del_ind")
        if sel_del and sel_del in ind_del and st.button("Eliminar indicador", key="btn_del_ind"):
            db.eliminar_indicadores(ind_del[sel_del]); _invalidar_cache(); st.success("Indicador eliminado."); st.rerun()

# ══════════════════════════════════════════════════════════════════════════
# DIAGNOSTICO TERRITORIAL (F-DT-01 a F-DT-06)
# ══════════════════════════════════════════════════════════════════════════
def pagina_diagnostico_territorial():
    st.subheader("Diagnostico Territorial - Fichas de Evaluacion")
    st.caption("Fichas F-DT-01 a F-DT-06: Parametros de evaluacion en campo para el diagnostico del territorio")
    bm = _bloques_map()
    if not bm:
        st.warning("Registre un bloque primero.")
        return

    # Inicializar estado de edicion DT
    if "dt_edit_id" not in st.session_state:
        st.session_state["dt_edit_id"] = None
    if "dt_edit_data" not in st.session_state:
        st.session_state["dt_edit_data"] = None

    dt_edit_id = st.session_state.get("dt_edit_id")
    dt_edit = st.session_state.get("dt_edit_data") or {}

    tab_reg, tab_hist, tab_excel = st.tabs(["Registro de Diagnostico", "Historial / Consulta", "Importar desde Excel"])

    with tab_reg:
        if dt_edit_id:
            st.markdown('<div class="edit-mode-banner"><span class="icon">&#9998;</span> '
                        f'Modo Edicion - Diagnostico Territorial ID {dt_edit_id}</div>', unsafe_allow_html=True)
            if st.button("Cancelar edicion", key="dt_cancel_edit"):
                st.session_state["dt_edit_id"] = None
                st.session_state["dt_edit_data"] = None
                st.rerun()

        bl = st.selectbox("Bloque de Intervencion", list(bm.keys()), key="dt_bl")
        bid = bm[bl]

        # Auto-resolver microcuenca del bloque seleccionado
        mc_auto_dt = _resolver_microcuenca(bl)
        if mc_auto_dt:
            mc_idx_dt = MICROCUENCAS.index(mc_auto_dt) + 1
            st.info(f"Microcuenca vinculada automaticamente: **{mc_auto_dt}**")
        else:
            mc_idx_dt = 0
        # En modo edicion, preseleccionar la microcuenca del registro
        if dt_edit_id:
            mc_val = dt_edit.get("microcuenca", "")
            if mc_val and mc_val in MICROCUENCAS:
                mc_idx_dt = MICROCUENCAS.index(mc_val) + 1

        def_fecha_dt = datetime.now()
        def_evaluador = ""
        if dt_edit_id:
            def_evaluador = dt_edit.get("evaluador", "") or ""
            try:
                def_fecha_dt = datetime.strptime(dt_edit.get("fecha_evaluacion", ""), "%Y-%m-%d")
            except (ValueError, TypeError):
                pass

        r1, r2, r3 = st.columns(3)
        mc = r1.selectbox("Microcuenca", [""] + MICROCUENCAS, index=mc_idx_dt, key="dt_mc")
        fecha_ev = r2.date_input("Fecha de evaluacion", value=def_fecha_dt, key="dt_fecha",
                                 min_value=FECHA_MIN_PROYECTO, max_value=date.today(),
                                 help="Seleccione la fecha de evaluacion en campo")
        evaluador = r3.text_input("Evaluador / Especialista", value=def_evaluador, key="dt_eval")

        st.markdown("---")
        st.markdown("### Seleccione los parametros de evaluacion por ficha")
        st.markdown("*Complete las fichas que correspondan a la visita de campo realizada.*")

        # ── F-DT-01 ──────────────────────────────────────────────────────
        with st.expander("F-DT-01: CARACTERISTICAS FISIOGRAFICAS", expanded=False):
            st.markdown("*Evaluacion de las condiciones fisiograficas del area de intervencion.*")
            c1, c2 = st.columns(2)
            forma_terreno = c1.selectbox("Forma del terreno", [""] + FDT01_FORMA_TERRENO, key="f01_ft")
            pendiente = c2.selectbox("Pendiente del terreno", [""] + FDT01_PENDIENTE, key="f01_pe")
            c3, c4 = st.columns(2)
            posicion_fisio = c3.selectbox("Posicion fisiografica", [""] + FDT01_POSICION_FISIOGRAFICA, key="f01_pf")
            exposicion = c4.selectbox("Exposicion / Orientacion", [""] + FDT01_EXPOSICION, key="f01_ex")
            c5, c6 = st.columns(2)
            paisaje = c5.selectbox("Paisaje dominante", [""] + FDT01_PAISAJE, key="f01_pa")
            rango_alt = c6.selectbox("Rango altitudinal", [""] + FDT01_RANGO_ALTITUDINAL, key="f01_ra")

        # ── F-DT-02 ──────────────────────────────────────────────────────
        with st.expander("F-DT-02: CONDICIONES CLIMATICAS", expanded=False):
            st.markdown("*Evaluacion de las condiciones climaticas predominantes en la zona.*")
            c1, c2 = st.columns(2)
            precipitacion = c1.selectbox("Precipitacion anual estimada", [""] + FDT02_PRECIPITACION, key="f02_pr")
            temperatura = c2.selectbox("Temperatura media anual", [""] + FDT02_TEMPERATURA, key="f02_te")
            c3, c4 = st.columns(2)
            humedad = c3.selectbox("Humedad relativa", [""] + FDT02_HUMEDAD, key="f02_hr")
            zona_vida = c4.selectbox("Zona de vida (Holdridge)", [""] + FDT02_ZONA_VIDA, key="f02_zv")
            c5, c6 = st.columns(2)
            heladas = c5.selectbox("Presencia de heladas", [""] + FDT02_HELADAS, key="f02_he")
            vientos = c6.selectbox("Regimen de vientos", [""] + FDT02_VIENTOS, key="f02_vi")

        # ── F-DT-03 ──────────────────────────────────────────────────────
        with st.expander("F-DT-03: CARACTERISTICAS DEL SUELO (Observacion de campo)", expanded=False):
            st.markdown("*Parametros del suelo evaluados mediante observacion directa en campo.*")
            c1, c2 = st.columns(2)
            textura = c1.selectbox("Textura al tacto", [""] + FDT03_TEXTURA, key="f03_tx")
            color_suelo = c2.selectbox("Color predominante del suelo", [""] + FDT03_COLOR, key="f03_co")
            c3, c4 = st.columns(2)
            profundidad = c3.selectbox("Profundidad efectiva", [""] + FDT03_PROFUNDIDAD, key="f03_pr")
            pedregosidad = c4.selectbox("Pedregosidad superficial", [""] + FDT03_PEDREGOSIDAD, key="f03_pd")
            c5, c6 = st.columns(2)
            drenaje = c5.selectbox("Drenaje", [""] + FDT03_DRENAJE, key="f03_dr")
            erosion = c6.selectbox("Presencia de erosion", [""] + FDT03_EROSION, key="f03_er")
            materia_org = st.selectbox("Materia organica (estimacion visual)", [""] + FDT03_MATERIA_ORGANICA, key="f03_mo")

        # ── F-DT-04 ──────────────────────────────────────────────────────
        with st.expander("F-DT-04: COBERTURA VEGETAL Y USO DEL SUELO", expanded=False):
            st.markdown("*Evaluacion de la cobertura vegetal existente y el uso actual del suelo.*")
            c1, c2 = st.columns(2)
            tipo_cob = c1.selectbox("Tipo de cobertura vegetal", [""] + FDT04_TIPO_COBERTURA, key="f04_tc")
            densidad_cob = c2.selectbox("Densidad de cobertura", [""] + FDT04_DENSIDAD, key="f04_dc")
            c3, c4 = st.columns(2)
            estado_cons = c3.selectbox("Estado de conservacion", [""] + FDT04_ESTADO_CONSERVACION, key="f04_ec")
            uso_suelo = c4.selectbox("Uso actual del suelo", [""] + FDT04_USO_SUELO, key="f04_us")
            conflicto = st.selectbox("Estado de Uso del Suelo", [""] + FDT04_CONFLICTO_USO, key="f04_cu")

        # ── F-DT-05 ──────────────────────────────────────────────────────
        with st.expander("F-DT-05: RECURSOS HIDRICOS", expanded=False):
            st.markdown("*Evaluacion de la disponibilidad y calidad de los recursos hidricos.*")
            c1, c2 = st.columns(2)
            fuente = c1.selectbox("Fuente de agua mas cercana", [""] + FDT05_FUENTE_AGUA, key="f05_fa")
            regimen_hid = c2.selectbox("Regimen hidrico", [""] + FDT05_REGIMEN, key="f05_rh")
            c3, c4 = st.columns(2)
            calidad_ag = c3.selectbox("Calidad aparente del agua", [""] + FDT05_CALIDAD_AGUA, key="f05_ca")
            dist_agua = c4.selectbox("Distancia a fuente de agua", [""] + FDT05_DISTANCIA_AGUA, key="f05_da")
            uso_hidrico = st.selectbox("Uso del recurso hidrico", [""] + FDT05_USO_HIDRICO, key="f05_uh")

        # ── F-DT-06 ──────────────────────────────────────────────────────
        with st.expander("F-DT-06: ASPECTOS SOCIOECONOMICOS Y ACCESIBILIDAD", expanded=False):
            st.markdown("*Evaluacion de factores socioeconomicos y accesibilidad del area.*")
            c1, c2 = st.columns(2)
            tenencia = c1.selectbox("Tenencia de la tierra", [""] + FDT06_TENENCIA, key="f06_tt")
            organizacion = c2.selectbox("Organizacion comunal", [""] + FDT06_ORGANIZACION, key="f06_oc")
            c3, c4 = st.columns(2)
            act_econ = c3.selectbox("Actividad economica principal", [""] + FDT06_ACTIVIDAD_ECONOMICA, key="f06_ae")
            accesib = c4.selectbox("Accesibilidad (via principal)", [""] + FDT06_ACCESIBILIDAD, key="f06_ac")
            c5, c6 = st.columns(2)
            dist_centro = c5.selectbox("Distancia al centro poblado", [""] + FDT06_DISTANCIA_CENTRO, key="f06_dp")
            servicios = c6.multiselect("Servicios basicos disponibles", FDT06_SERVICIOS, key="f06_sb")

        st.markdown("---")
        observ_gen = st.text_area("Observaciones generales del diagnostico", key="dt_obs")

        # Determinar fichas completadas
        fichas_sel = []
        if any([forma_terreno, pendiente, posicion_fisio, exposicion, paisaje, rango_alt]):
            fichas_sel.append("F-DT-01")
        if any([precipitacion, temperatura, humedad, zona_vida, heladas, vientos]):
            fichas_sel.append("F-DT-02")
        if any([textura, color_suelo, profundidad, pedregosidad, drenaje, erosion, materia_org]):
            fichas_sel.append("F-DT-03")
        if any([tipo_cob, densidad_cob, estado_cons, uso_suelo, conflicto]):
            fichas_sel.append("F-DT-04")
        if any([fuente, regimen_hid, calidad_ag, dist_agua, uso_hidrico]):
            fichas_sel.append("F-DT-05")
        if any([tenencia, organizacion, act_econ, accesib, dist_centro, servicios]):
            fichas_sel.append("F-DT-06")

        if fichas_sel:
            st.info(f"Fichas con datos: **{', '.join(fichas_sel)}** ({len(fichas_sel)}/6)")

        btn_label_dt = "Actualizar Diagnostico Territorial" if dt_edit_id else "Guardar Diagnostico Territorial"
        if st.button(btn_label_dt, type="primary", key="dt_guardar"):
            if not evaluador:
                st.warning("Ingrese el nombre del evaluador.")
            elif not fichas_sel:
                st.warning("Complete al menos una ficha de diagnostico.")
            else:
                try:
                    _dt_kwargs = dict(
                        ficha=", ".join(fichas_sel),
                        fecha_evaluacion=fecha_ev.strftime("%Y-%m-%d"),
                        evaluador=evaluador,
                        microcuenca=mc,
                        forma_terreno=forma_terreno,
                        pendiente=pendiente,
                        posicion_fisiografica=posicion_fisio,
                        exposicion_orientacion=exposicion,
                        paisaje_dominante=paisaje,
                        rango_altitudinal=rango_alt,
                        precipitacion_anual=precipitacion,
                        temperatura_media=temperatura,
                        humedad_relativa=humedad,
                        zona_vida=zona_vida,
                        presencia_heladas=heladas,
                        regimen_vientos=vientos,
                        textura_suelo=textura,
                        color_suelo=color_suelo,
                        profundidad_efectiva=profundidad,
                        pedregosidad=pedregosidad,
                        drenaje=drenaje,
                        presencia_erosion=erosion,
                        materia_organica=materia_org,
                        tipo_cobertura=tipo_cob,
                        densidad_cobertura=densidad_cob,
                        estado_conservacion=estado_cons,
                        uso_actual_suelo=uso_suelo,
                        conflicto_uso=conflicto,
                        fuente_agua=fuente,
                        regimen_hidrico=regimen_hid,
                        calidad_agua=calidad_ag,
                        distancia_fuente_agua=dist_agua,
                        uso_recurso_hidrico=uso_hidrico,
                        tenencia_tierra=tenencia,
                        organizacion_comunal=organizacion,
                        actividad_economica=act_econ,
                        accesibilidad_via=accesib,
                        distancia_centro_poblado=dist_centro,
                        servicios_basicos=", ".join(servicios) if servicios else "",
                        observaciones_generales=observ_gen,
                    )
                    if dt_edit_id:
                        db.actualizar_diagnostico_territorial(dt_edit_id, **_dt_kwargs)
                        _invalidar_cache()
                        st.session_state["dt_edit_id"] = None
                        st.session_state["dt_edit_data"] = None
                        st.success(f"Diagnostico territorial actualizado ({', '.join(fichas_sel)}).")
                    else:
                        # Verificar duplicados
                        existentes = db.obtener_diagnosticos_por_bloque(bid)
                        dup = [e for e in existentes
                               if e.get("fecha_evaluacion") == fecha_ev.strftime("%Y-%m-%d")
                               and e.get("evaluador") == evaluador]
                        if dup:
                            st.markdown('<div class="dup-warning">Ya existe un diagnostico para este bloque en '
                                       f'{fecha_ev.strftime("%Y-%m-%d")} por {evaluador}. '
                                       f'Use <b>Editar</b> en Historial para modificarlo.</div>',
                                       unsafe_allow_html=True)
                        else:
                            db.insertar_diagnostico_territorial(bloque_id=bid, **_dt_kwargs)
                            _invalidar_cache()
                            st.success(f"Diagnostico territorial guardado ({', '.join(fichas_sel)}).")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

    with tab_hist:
        st.markdown("### Historial de Diagnosticos Territoriales")
        st.caption("Haga clic en **Editar** para modificar un diagnostico existente y evitar duplicidades.")
        todos_dt = _cached_obtener_todos_diagnosticos(_cache_version())
        if not todos_dt:
            st.info("No hay diagnosticos registrados.")
        else:
            # Paginacion
            dt_pag, total_pags_dt, pag_actual_dt = _paginar(todos_dt, "pag_dt")
            _controles_paginacion(total_pags_dt, pag_actual_dt, "pag_dt")
            # Tabla con botones de edicion
            header_cols = st.columns([0.4, 1, 0.8, 0.8, 0.8, 0.8, 0.8, 0.5])
            for col, h in zip(header_cols, ["ID", "Bloque", "Fichas", "Fecha", "Evaluador", "Microcuenca", "Distrito", ""]):
                col.markdown(f"**{h}**")
            st.markdown("---")
            for d in dt_pag:
                row = st.columns([0.4, 1, 0.8, 0.8, 0.8, 0.8, 0.8, 0.5])
                row[0].write(d["id"])
                row[1].write(d.get("bloque_codigo", ""))
                row[2].write(d.get("ficha", ""))
                row[3].write(d.get("fecha_evaluacion", ""))
                row[4].write(d.get("evaluador", ""))
                row[5].write(d.get("microcuenca", "") or "")
                row[6].write(d.get("distrito", ""))
                if row[7].button("Editar", key=f"edit_dt_{d['id']}", type="primary"):
                    # Cargar el diagnostico completo para edicion en tab_reg
                    det = db.obtener_diagnostico_por_id(d["id"])
                    if det:
                        st.session_state["dt_edit_id"] = det["id"]
                        st.session_state["dt_edit_data"] = det
                    st.rerun()

            st.markdown("---")
            st.markdown("### Detalle de Diagnostico")
            dm = {f"ID {d['id']} - {d.get('bloque_codigo','')} ({d.get('ficha','')})": d["id"] for d in todos_dt}
            sel_dt = st.selectbox("Seleccionar diagnostico", [""] + list(dm.keys()), key="dt_det")
            if sel_dt and sel_dt in dm:
                det = db.obtener_diagnostico_por_id(dm[sel_dt])
                if det:
                    st.markdown(f"**Bloque:** {det.get('bloque_codigo','')} | "
                                f"**Fecha:** {det.get('fecha_evaluacion','')} | "
                                f"**Evaluador:** {det.get('evaluador','')}")

                    # Mostrar cada ficha completada
                    fichas_str = det.get("ficha", "")
                    if "F-DT-01" in fichas_str:
                        with st.expander("F-DT-01: CARACTERISTICAS FISIOGRAFICAS", expanded=True):
                            c1, c2, c3 = st.columns(3)
                            c1.markdown(f"**Forma terreno:** {det.get('forma_terreno','') or '-'}")
                            c2.markdown(f"**Pendiente:** {det.get('pendiente','') or '-'}")
                            c3.markdown(f"**Posicion:** {det.get('posicion_fisiografica','') or '-'}")
                            c1.markdown(f"**Exposicion:** {det.get('exposicion_orientacion','') or '-'}")
                            c2.markdown(f"**Paisaje:** {det.get('paisaje_dominante','') or '-'}")
                            c3.markdown(f"**Altitud:** {det.get('rango_altitudinal','') or '-'}")

                    if "F-DT-02" in fichas_str:
                        with st.expander("F-DT-02: CONDICIONES CLIMATICAS", expanded=True):
                            c1, c2, c3 = st.columns(3)
                            c1.markdown(f"**Precipitacion:** {det.get('precipitacion_anual','') or '-'}")
                            c2.markdown(f"**Temperatura:** {det.get('temperatura_media','') or '-'}")
                            c3.markdown(f"**Humedad:** {det.get('humedad_relativa','') or '-'}")
                            c1.markdown(f"**Zona de vida:** {det.get('zona_vida','') or '-'}")
                            c2.markdown(f"**Heladas:** {det.get('presencia_heladas','') or '-'}")
                            c3.markdown(f"**Vientos:** {det.get('regimen_vientos','') or '-'}")

                    if "F-DT-03" in fichas_str:
                        with st.expander("F-DT-03: CARACTERISTICAS DEL SUELO", expanded=True):
                            c1, c2, c3 = st.columns(3)
                            c1.markdown(f"**Textura:** {det.get('textura_suelo','') or '-'}")
                            c2.markdown(f"**Color:** {det.get('color_suelo','') or '-'}")
                            c3.markdown(f"**Profundidad:** {det.get('profundidad_efectiva','') or '-'}")
                            c1.markdown(f"**Pedregosidad:** {det.get('pedregosidad','') or '-'}")
                            c2.markdown(f"**Drenaje:** {det.get('drenaje','') or '-'}")
                            c3.markdown(f"**Erosion:** {det.get('presencia_erosion','') or '-'}")
                            st.markdown(f"**Materia organica:** {det.get('materia_organica','') or '-'}")

                    if "F-DT-04" in fichas_str:
                        with st.expander("F-DT-04: COBERTURA VEGETAL Y USO DEL SUELO", expanded=True):
                            c1, c2 = st.columns(2)
                            c1.markdown(f"**Tipo cobertura:** {det.get('tipo_cobertura','') or '-'}")
                            c2.markdown(f"**Densidad:** {det.get('densidad_cobertura','') or '-'}")
                            c1.markdown(f"**Estado conservacion:** {det.get('estado_conservacion','') or '-'}")
                            c2.markdown(f"**Uso actual:** {det.get('uso_actual_suelo','') or '-'}")
                            st.markdown(f"**Estado de Uso del Suelo:** {det.get('conflicto_uso','') or '-'}")

                    if "F-DT-05" in fichas_str:
                        with st.expander("F-DT-05: RECURSOS HIDRICOS", expanded=True):
                            c1, c2 = st.columns(2)
                            c1.markdown(f"**Fuente de agua:** {det.get('fuente_agua','') or '-'}")
                            c2.markdown(f"**Regimen hidrico:** {det.get('regimen_hidrico','') or '-'}")
                            c1.markdown(f"**Calidad agua:** {det.get('calidad_agua','') or '-'}")
                            c2.markdown(f"**Distancia:** {det.get('distancia_fuente_agua','') or '-'}")
                            st.markdown(f"**Uso recurso hidrico:** {det.get('uso_recurso_hidrico','') or '-'}")

                    if "F-DT-06" in fichas_str:
                        with st.expander("F-DT-06: ASPECTOS SOCIOECONOMICOS", expanded=True):
                            c1, c2 = st.columns(2)
                            c1.markdown(f"**Tenencia tierra:** {det.get('tenencia_tierra','') or '-'}")
                            c2.markdown(f"**Organizacion:** {det.get('organizacion_comunal','') or '-'}")
                            c1.markdown(f"**Act. economica:** {det.get('actividad_economica','') or '-'}")
                            c2.markdown(f"**Accesibilidad:** {det.get('accesibilidad_via','') or '-'}")
                            c1.markdown(f"**Dist. centro poblado:** {det.get('distancia_centro_poblado','') or '-'}")
                            c2.markdown(f"**Servicios basicos:** {det.get('servicios_basicos','') or '-'}")

                    if det.get("observaciones_generales"):
                        st.markdown(f"**Observaciones:** {det['observaciones_generales']}")

                    if st.button("Eliminar este diagnostico", key="dt_eliminar"):
                        db.eliminar_diagnostico(dm[sel_dt])
                        _invalidar_cache()
                        st.success("Diagnostico eliminado.")
                        st.rerun()

            st.markdown("---")
            st.markdown("### Resumen por Bloque")
            resumen = db.obtener_resumen_diagnosticos()
            if resumen:
                st.dataframe(pd.DataFrame([{
                    "Bloque": r["codigo"],
                    "Tipo": r["tipo_intervencion"],
                    "Distrito": r["distrito"],
                    "Total Fichas": r["total_fichas"],
                    "Fichas Completadas": r.get("fichas_completadas", "") or "Ninguna",
                } for r in resumen]), use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════════════
    # TAB IMPORTAR DESDE EXCEL
    # ══════════════════════════════════════════════════════════════════
    with tab_excel:
        st.markdown("### Importar Diagnostico Territorial desde Excel")
        st.caption("Suba un archivo Excel llenado por el tecnico de campo para autocompletar "
                   "los formularios. Puede descargar la plantilla estandarizada para su uso en campo.")

        # ── Descargar plantilla ──────────────────────────────────────
        st.markdown("---")
        st.markdown("**1. Descargar Plantilla para Tecnicos**")
        col_dl1, col_dl2 = st.columns(2)
        fichas_descarga_dt = col_dl1.multiselect(
            "Fichas a incluir en la plantilla",
            FICHAS_DT, default=FICHAS_DT, key="dt_excel_fichas_dl")
        if col_dl2.button("Generar Plantilla Excel", type="secondary", key="dt_gen_plantilla"):
            if fichas_descarga_dt:
                bloques_data_dt = [(b[1], b[2], b[4], b[5]) for b in BLOQUES_124]
                plantilla_bytes_dt = generar_plantilla_dt(fichas_descarga_dt, bloques_data_dt)
                st.session_state["dt_plantilla_bytes"] = plantilla_bytes_dt
                st.success("Plantilla generada correctamente.")

        if st.session_state.get("dt_plantilla_bytes"):
            st.download_button(
                "Descargar Plantilla Excel",
                st.session_state["dt_plantilla_bytes"],
                file_name="Plantilla_Diagnostico_Territorial_IN_Piura.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dt_dl_plantilla")

        # ── Subir Excel llenado ──────────────────────────────────────
        st.markdown("---")
        st.markdown("**2. Subir Excel Llenado por el Tecnico**")
        st.info("Al cargar el archivo, el sistema leera los datos y autocompletara "
                "el formulario en la pestana **Registro de Diagnostico**. Podra revisar "
                "y ajustar antes de guardar.")

        uploaded_excel_dt = st.file_uploader(
            "Seleccionar archivo Excel (.xlsx)",
            type=["xlsx"], key="dt_excel_upload")

        if uploaded_excel_dt is not None:
            try:
                resultados_dt = parsear_excel_dt(uploaded_excel_dt)
                if not resultados_dt:
                    st.error("No se pudieron detectar fichas en el archivo. "
                             "Verifique que el formato sea correcto.")
                else:
                    st.success(f"Se detectaron {len(resultados_dt)} ficha(s) en el archivo.")

                    # Consolidar datos de todas las fichas para vista previa
                    datos_consolidados = {}
                    fichas_detectadas = []
                    for res in resultados_dt:
                        ficha_det = res["ficha"]
                        datos_res = res["datos"]
                        fichas_detectadas.append(ficha_det)

                        with st.expander(f"Vista previa: {ficha_det}", expanded=True):
                            cols_prev = st.columns(4)
                            cols_prev[0].markdown(f"**Fecha:** {datos_res.get('fecha', '-')}")
                            cols_prev[1].markdown(f"**Evaluador:** {datos_res.get('evaluador', '-')}")
                            cols_prev[2].markdown(f"**Bloque:** {datos_res.get('codigo_bloque', '-')}")
                            cols_prev[3].markdown(f"**Microcuenca:** {datos_res.get('microcuenca', '-')}")

                            # Mostrar parametros segun ficha
                            campos_ficha = {k: v for k, v in datos_res.items()
                                            if k not in ("fecha", "evaluador", "codigo_bloque",
                                                         "microcuenca", "observaciones")
                                            and v}
                            if campos_ficha:
                                st.markdown(f"**Parametros completados:** {len(campos_ficha)}")
                                for k, v in campos_ficha.items():
                                    label = k.replace("_", " ").title()
                                    st.markdown(f"- **{label}:** {v}")

                        datos_consolidados.update(datos_res)

                    # Botones de accion
                    col_ac1, col_ac2 = st.columns(2)

                    if col_ac1.button(
                        "Autocompletar formulario",
                        type="primary", key="dt_autocompletar"):
                        # Consolidar todos los datos de todas las fichas
                        datos_todos = {}
                        for res in resultados_dt:
                            datos_todos.update(res["datos"])
                        # Crear un objeto consolidado para mapear
                        consolidado = {"ficha": ", ".join(fichas_detectadas), "datos": datos_todos}
                        ss_vals = mapear_dt_a_session_state(consolidado, bm)
                        for k, v in ss_vals.items():
                            st.session_state[k] = v
                        st.session_state["dt_edit_id"] = None
                        st.success(f"Formulario autocompletado con {len(fichas_detectadas)} ficha(s): "
                                   f"{', '.join(fichas_detectadas)}. "
                                   f"Cambie a la pestana **Registro de Diagnostico** para revisar y guardar.")
                        st.rerun()

                    if col_ac2.button(
                        "Guardar directamente",
                        type="secondary", key="dt_guardar_directo"):
                        try:
                            # Consolidar datos
                            datos_todos = {}
                            for res in resultados_dt:
                                datos_todos.update(res["datos"])

                            # Resolver bloque
                            codigo_bloque = datos_todos.get("codigo_bloque", "")
                            bid_excel = None
                            for label, id_val in bm.items():
                                if codigo_bloque and codigo_bloque in label:
                                    bid_excel = id_val
                                    break
                            if not bid_excel:
                                bid_excel = list(bm.values())[0]
                                st.warning(f"Bloque '{codigo_bloque}' no encontrado. "
                                           f"Se asigno al primer bloque disponible.")

                            # Normalizar fecha
                            fecha_str = str(datos_todos.get("fecha", ""))
                            if not fecha_str:
                                fecha_str = datetime.now().strftime("%Y-%m-%d")
                            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
                                try:
                                    fecha_str = datetime.strptime(
                                        fecha_str.split(" ")[0], fmt).strftime("%Y-%m-%d")
                                    break
                                except (ValueError, TypeError):
                                    continue

                            # Servicios basicos
                            serv_str = datos_todos.get("servicios_basicos", "")

                            _dt_kwargs = dict(
                                ficha=", ".join(fichas_detectadas),
                                fecha_evaluacion=fecha_str,
                                evaluador=datos_todos.get("evaluador", ""),
                                microcuenca=datos_todos.get("microcuenca", ""),
                                forma_terreno=datos_todos.get("forma_terreno", ""),
                                pendiente=datos_todos.get("pendiente", ""),
                                posicion_fisiografica=datos_todos.get("posicion_fisiografica", ""),
                                exposicion_orientacion=datos_todos.get("exposicion_orientacion", ""),
                                paisaje_dominante=datos_todos.get("paisaje_dominante", ""),
                                rango_altitudinal=datos_todos.get("rango_altitudinal", ""),
                                precipitacion_anual=datos_todos.get("precipitacion_anual", ""),
                                temperatura_media=datos_todos.get("temperatura_media", ""),
                                humedad_relativa=datos_todos.get("humedad_relativa", ""),
                                zona_vida=datos_todos.get("zona_vida", ""),
                                presencia_heladas=datos_todos.get("presencia_heladas", ""),
                                regimen_vientos=datos_todos.get("regimen_vientos", ""),
                                textura_suelo=datos_todos.get("textura_suelo", ""),
                                color_suelo=datos_todos.get("color_suelo", ""),
                                profundidad_efectiva=datos_todos.get("profundidad_efectiva", ""),
                                pedregosidad=datos_todos.get("pedregosidad", ""),
                                drenaje=datos_todos.get("drenaje", ""),
                                presencia_erosion=datos_todos.get("presencia_erosion", ""),
                                materia_organica=datos_todos.get("materia_organica", ""),
                                tipo_cobertura=datos_todos.get("tipo_cobertura", ""),
                                densidad_cobertura=datos_todos.get("densidad_cobertura", ""),
                                estado_conservacion=datos_todos.get("estado_conservacion", ""),
                                uso_actual_suelo=datos_todos.get("uso_actual_suelo", ""),
                                conflicto_uso=datos_todos.get("conflicto_uso", ""),
                                fuente_agua=datos_todos.get("fuente_agua", ""),
                                regimen_hidrico=datos_todos.get("regimen_hidrico", ""),
                                calidad_agua=datos_todos.get("calidad_agua", ""),
                                distancia_fuente_agua=datos_todos.get("distancia_fuente_agua", ""),
                                uso_recurso_hidrico=datos_todos.get("uso_recurso_hidrico", ""),
                                tenencia_tierra=datos_todos.get("tenencia_tierra", ""),
                                organizacion_comunal=datos_todos.get("organizacion_comunal", ""),
                                actividad_economica=datos_todos.get("actividad_economica", ""),
                                accesibilidad_via=datos_todos.get("accesibilidad_via", ""),
                                distancia_centro_poblado=datos_todos.get("distancia_centro_poblado", ""),
                                servicios_basicos=serv_str,
                                observaciones_generales=datos_todos.get("observaciones", ""),
                            )

                            db.insertar_diagnostico_territorial(bloque_id=bid_excel, **_dt_kwargs)
                            _invalidar_cache()
                            st.success(f"Diagnostico territorial guardado directamente "
                                       f"({', '.join(fichas_detectadas)}).")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error al guardar: {e}")

            except Exception as e:
                st.error(f"Error al leer el archivo Excel: {e}")

        # ── Importacion masiva ───────────────────────────────────────
        st.markdown("---")
        st.markdown("**3. Importacion Masiva (multiples fichas)**")
        st.caption("Si el archivo Excel contiene multiples hojas (una por ficha), "
                   "puede importar todas a la vez usando el boton de arriba. "
                   "Cada hoja sera detectada automaticamente.")
        st.caption("Tambien puede subir la plantilla original "
                   "(`Plantilla_DT_Campo_Check_24abr_hscm.xlsx`) "
                   "llenada por el tecnico.")


# ══════════════════════════════════════════════════════════════════════════
# DIAGNOSTICO SOCIAL
# ══════════════════════════════════════════════════════════════════════════
def _ds_datos_generales(bloque_label=""):
    """Campos de datos generales compartidos por todas las fichas DS.
    Auto-vincula centros poblados, comunidades campesinas, provincia,
    distrito y coordenadas aproximadas del bloque seleccionado."""
    codigo_bloque = bloque_label.split(" - ")[0].strip() if " - " in bloque_label else bloque_label.strip()
    bloque_info = BLOQUES_124_MAP.get(codigo_bloque, {})
    cp_info = CENTROS_POBLADOS_BLOQUE.get(codigo_bloque, {})
    lista_cp = cp_info.get("centros_poblados", [])
    lista_cc = cp_info.get("comunidades_campesinas", [])
    pob_total = cp_info.get("poblacion_total", 0)

    # ── Vinculacion automatica de centros poblados ──
    if lista_cp:
        n_cp = len(lista_cp)
        n_cc = len(lista_cc)
        resumen = f"**{n_cp}** centro(s) poblado(s)"
        if lista_cc:
            resumen += f" y **{n_cc}** comunidad(es) campesina(s)"
        resumen += f" vinculado(s) al bloque **{codigo_bloque}**"
        if pob_total:
            resumen += f" | Poblacion total: **{pob_total:,}** hab."
        st.info(resumen)

        # Mostrar tabla resumen si hay multiples centros poblados
        if n_cp > 1:
            with st.expander(f"Ver {n_cp} centros poblados asociados al bloque {codigo_bloque}", expanded=False):
                df_cp = pd.DataFrame({"N": range(1, n_cp + 1), "Centro Poblado": lista_cp})
                if lista_cc:
                    df_cp_cc = pd.DataFrame({
                        "Centros Poblados": ", ".join(lista_cp),
                        "Comunidades Campesinas": ", ".join(lista_cc),
                    }, index=[0])
                    st.dataframe(df_cp_cc, use_container_width=True, hide_index=True)
                else:
                    st.dataframe(df_cp, use_container_width=True, hide_index=True)

    # ── Auto-resolver provincia y distrito del bloque ──
    prov_auto = bloque_info.get("provincia", "")
    dist_auto = bloque_info.get("distrito", "")
    utm_este_auto = float(bloque_info.get("utm_este", 0))
    utm_norte_auto = float(bloque_info.get("utm_norte", 0))

    # Pre-poblar session_state con valores del bloque si estan vacios
    if "ds_prov" not in st.session_state and prov_auto:
        st.session_state["ds_prov"] = prov_auto
    if "ds_dist" not in st.session_state and dist_auto:
        st.session_state["ds_dist"] = dist_auto
    if "ds_cpob" not in st.session_state and lista_cp:
        st.session_state["ds_cpob"] = " / ".join(lista_cp)
    if "ds_ccam" not in st.session_state and lista_cc:
        st.session_state["ds_ccam"] = " / ".join(lista_cc)
    if "ds_este" not in st.session_state and utm_este_auto:
        st.session_state["ds_este"] = utm_este_auto
    if "ds_norte" not in st.session_state and utm_norte_auto:
        st.session_state["ds_norte"] = utm_norte_auto

    c1, c2, c3, c4 = st.columns(4)
    prov = c1.text_input("Provincia", key="ds_prov")
    dist = c2.text_input("Distrito", key="ds_dist")

    # Selector de centro poblado cuando hay multiples opciones
    if len(lista_cp) > 1:
        opciones_cp = [" / ".join(lista_cp)] + lista_cp
        # Asegurar que el valor en session_state sea valido para el selectbox
        val_cpob = st.session_state.get("ds_cpob", "")
        if val_cpob and val_cpob not in opciones_cp:
            st.session_state["ds_cpob"] = opciones_cp[0]
        cpob = c3.selectbox("Centro Poblado / Localidad", opciones_cp,
                            key="ds_cpob",
                            help="Seleccione un centro poblado especifico o todos los asociados al bloque")
    else:
        cpob = c3.text_input("Centro Poblado / Localidad", key="ds_cpob")

    if len(lista_cc) > 1:
        opciones_cc = [" / ".join(lista_cc)] + lista_cc
        val_ccam = st.session_state.get("ds_ccam", "")
        if val_ccam and val_ccam not in opciones_cc:
            st.session_state["ds_ccam"] = opciones_cc[0]
        ccam = c4.selectbox("Comunidad Campesina", opciones_cc,
                            key="ds_ccam",
                            help="Seleccione una comunidad campesina especifica o todas las asociadas")
    else:
        ccam = c4.text_input("Comunidad Campesina", key="ds_ccam")

    # ── Coordenadas: usar centroide del bloque como aproximacion ──
    c5, c6, c7, c8 = st.columns(4)
    este = c5.number_input("Coordenada Este (UTM)", format="%.1f", key="ds_este",
                           help="Coordenada UTM Este (Zona 17S). Auto-completada con centroide del bloque.")
    norte = c6.number_input("Coordenada Norte (UTM)", format="%.1f", key="ds_norte",
                            help="Coordenada UTM Norte (Zona 17S). Auto-completada con centroide del bloque.")
    alt_v = c7.number_input("Altitud (msnm)", value=0.0, format="%.0f", key="ds_alt")
    ubigeo = c8.text_input("Codigo UBIGEO", key="ds_ubigeo")

    if utm_este_auto and utm_norte_auto and (este == utm_este_auto and norte == utm_norte_auto):
        st.caption("📍 Coordenadas aproximadas (centroide del bloque). Ajuste manualmente si dispone de la ubicacion exacta del centro poblado.")

    return dict(provincia=prov, distrito=dist, centro_poblado=cpob,
                comunidad_campesina=ccam, coordenada_este=este,
                coordenada_norte=norte, altitud=alt_v, codigo_ubigeo=ubigeo)


def _ds_load_edit(det, bm):
    """Carga un registro de diagnostico social en session_state para edicion."""
    st.session_state["ds_edit_id"] = det["id"]
    # Encontrar label del bloque
    bm_rev = {v: k for k, v in bm.items()}
    bl_label = bm_rev.get(det.get("bloque_id"), "")
    if bl_label:
        st.session_state["ds_bl"] = bl_label
    st.session_state["ds_mc"] = det.get("microcuenca", "") or ""
    if det.get("fecha_evaluacion"):
        try:
            st.session_state["ds_fecha"] = datetime.strptime(det["fecha_evaluacion"], "%Y-%m-%d").date()
        except (ValueError, TypeError):
            pass
    st.session_state["ds_eval"] = det.get("evaluador", "") or ""
    st.session_state["ds_fnum"] = det.get("ficha_numero", "") or ""
    st.session_state["ds_prov"] = det.get("provincia", "") or ""
    st.session_state["ds_dist"] = det.get("distrito", "") or ""
    st.session_state["ds_cpob"] = det.get("centro_poblado", "") or ""
    st.session_state["ds_ccam"] = det.get("comunidad_campesina", "") or ""
    st.session_state["ds_este"] = float(det.get("coordenada_este") or 0)
    st.session_state["ds_norte"] = float(det.get("coordenada_norte") or 0)
    st.session_state["ds_alt"] = float(det.get("altitud") or 0)
    st.session_state["ds_ubigeo"] = det.get("codigo_ubigeo", "") or ""
    st.session_state["ds_obs"] = det.get("observaciones_generales", "") or ""
    ficha = det.get("ficha", "")
    st.session_state["ds_ficha_sel"] = ficha

    def _split(v):
        return [x.strip() for x in (v or "").split(",") if x.strip()]

    if ficha == "F-DS-01":
        st.session_state["s01_nf"] = det.get("ds01_num_familias", "") or ""
        st.session_state["s01_ph"] = det.get("ds01_poblacion_hombres", "") or ""
        st.session_state["s01_pm"] = det.get("ds01_poblacion_mujeres", "") or ""
        st.session_state["s01_pt"] = det.get("ds01_poblacion_total", "") or ""
        st.session_state["s01_id"] = _split(det.get("ds01_idioma"))
        st.session_state["s01_ne"] = _split(det.get("ds01_nivel_educativo"))
        st.session_state["s01_mi"] = det.get("ds01_tasa_migracion", "") or ""
        st.session_state["s01_dest"] = det.get("ds01_destino_migracion", "") or ""
        st.session_state["s01_org"] = _split(det.get("ds01_organizacion_comunal"))
        st.session_state["s01_jd"] = det.get("ds01_junta_directiva", "") or ""
        st.session_state["s01_pres"] = det.get("ds01_presidente_junta", "") or ""
        st.session_state["s01_ag"] = _split(det.get("ds01_agua_potable_tipo"))
        st.session_state["s01_agcob"] = det.get("ds01_agua_potable_cobertura", "") or ""
        st.session_state["s01_san"] = _split(det.get("ds01_saneamiento"))
        st.session_state["s01_en"] = _split(det.get("ds01_energia_tipo"))
        st.session_state["s01_encob"] = det.get("ds01_energia_cobertura", "") or ""
        st.session_state["s01_tel"] = _split(det.get("ds01_telecomunicaciones"))
        st.session_state["s01_telop"] = det.get("ds01_telecom_operador", "") or ""
        st.session_state["s01_via"] = _split(det.get("ds01_acceso_vial"))
        st.session_state["s01_dcap"] = det.get("ds01_distancia_capital", "") or ""
        st.session_state["s01_tr"] = det.get("ds01_transporte", "") or ""
        st.session_state["s01_sal"] = _split(det.get("ds01_salud_tipo"))
        st.session_state["s01_sdist"] = det.get("ds01_salud_distancia", "") or ""
        st.session_state["s01_educ"] = _split(det.get("ds01_educacion"))
        st.session_state["s01_fag"] = det.get("ds01_fuente_agua", "") or ""
        st.session_state["s01_pag"] = _split(det.get("ds01_problemas_agua"))
        st.session_state["s01_uf"] = _split(det.get("ds01_uso_recursos_forestales"))
        st.session_state["s01_ff"] = det.get("ds01_frecuencia_uso_forestal", "") or ""
        st.session_state["s01_pcam"] = det.get("ds01_percepcion_cambios", "") or ""
        st.session_state["s01_disp"] = det.get("ds01_disposicion_participar", "") or ""
        st.session_state["s01_cdisp"] = det.get("ds01_comentario_disposicion", "") or ""
        st.session_state["s01_activos"] = det.get("ds01_activos_asociados", "") or ""
        st.session_state["s01_ten_com"] = det.get("ds01_tenencia_comunal_ha", "") or ""
        st.session_state["s01_ten_pri"] = det.get("ds01_tenencia_privada_ha", "") or ""
        st.session_state["s01_ten_est"] = det.get("ds01_tenencia_estatal_ha", "") or ""
        # Actividades economicas (JSON)
        act_json = det.get("ds01_actividades_economicas", "") or ""
        if act_json:
            try:
                acts = json.loads(act_json)
                st.session_state["s01_nact"] = max(len(acts), 1)
                for i, a in enumerate(acts):
                    st.session_state[f"s01_act{i}"] = a.get("actividad", "")
                    st.session_state[f"s01_pct{i}"] = a.get("pct_familias", "")
                    st.session_state[f"s01_prod{i}"] = a.get("productos", "")
                    st.session_state[f"s01_dest{i}"] = a.get("destino", "")
                    st.session_state[f"s01_ing{i}"] = a.get("ingreso", "")
            except (json.JSONDecodeError, TypeError):
                pass

    elif ficha == "F-DS-02":
        act_json = det.get("ds02_registro_actores", "") or ""
        if act_json:
            try:
                acts = json.loads(act_json)
                st.session_state["s02_nact"] = max(len(acts), 1)
                for i, a in enumerate(acts):
                    st.session_state[f"s02_nom{i}"] = a.get("nombre", "")
                    st.session_state[f"s02_tip{i}"] = a.get("tipo", "")
                    st.session_state[f"s02_rol{i}"] = a.get("rol", "")
                    st.session_state[f"s02_rel{i}"] = a.get("relacion", "")
                    st.session_state[f"s02_inf{i}"] = a.get("influencia", "")
                    st.session_state[f"s02_int{i}"] = a.get("interes", "")
                    st.session_state[f"s02_con{i}"] = a.get("contacto", "")
            except (json.JSONDecodeError, TypeError):
                pass
        for campo, cat_prefix in [
            ("ds02_actores_gob_local", "Gobierno Local (M"),
            ("ds02_actores_gob_regional", "Gobierno Regional"),
            ("ds02_actores_gob_nacional", "Gobierno Nacional"),
            ("ds02_actores_comunidades", "Comunidades Camp"),
            ("ds02_actores_juntas_riego", "Juntas de Usuar"),
            ("ds02_actores_comites_cuenca", "Comites de Gesti"),
            ("ds02_actores_ong", "ONG / Cooperacio"),
            ("ds02_actores_empresa", "Empresa Privada"),
            ("ds02_actores_educacion", "Instituciones Ed"),
            ("ds02_actores_org_base", "Organizaciones d"),
        ]:
            v = det.get(campo, "") or ""
            for cat in FDS02_CLASIFICACION:
                if cat.startswith(cat_prefix[:15]):
                    st.session_state[f"s02_cl_{cat[:15]}"] = v
                    break

    elif ficha == "F-DS-03":
        st.session_state["s03_nom"] = det.get("ds03_nombre_entrevistado", "") or ""
        st.session_state["s03_car"] = det.get("ds03_cargo_funcion", "") or ""
        st.session_state["s03_inst"] = det.get("ds03_institucion", "") or ""
        st.session_state["s03_tel"] = det.get("ds03_telefono_correo", "") or ""
        st.session_state["s03_dur"] = det.get("ds03_duracion", "") or ""
        resp_map = {
            "s03_r1": "ds03_resp_recursos_naturales", "s03_r2": "ds03_resp_cambios_ambiente",
            "s03_r3": "ds03_resp_problemas_ambientales", "s03_r4": "ds03_resp_zonas_conservacion",
            "s03_r5": "ds03_resp_actividades_economicas", "s03_r6": "ds03_resp_abastecimiento_agua",
            "s03_r7": "ds03_resp_productos_bosque", "s03_r8": "ds03_resp_cadenas_productivas",
            "s03_r9": "ds03_resp_organizaciones", "s03_r10": "ds03_resp_decisiones_territorio",
            "s03_r11": "ds03_resp_conflictos", "s03_r12": "ds03_resp_proyectos_anteriores",
            "s03_r12b": "ds03_resp_experiencia_reforestacion",
            "s03_r13": "ds03_resp_conocimiento_restauracion", "s03_r14": "ds03_resp_expectativas",
            "s03_r15": "ds03_resp_disposicion_participar", "s03_r16": "ds03_resp_condiciones",
            "s03_r17": "ds03_resp_conocimiento_merese", "s03_r18": "ds03_resp_beneficiarios",
            "s03_r19": "ds03_resp_instituciones_contribuyentes", "s03_r20": "ds03_resp_experiencias_pago",
        }
        for wk, dbk in resp_map.items():
            st.session_state[wk] = det.get(dbk, "") or ""

    elif ficha == "F-DS-04":
        st.session_state["s04_lug"] = det.get("ds04_lugar_taller", "") or ""
        st.session_state["s04_conv"] = det.get("ds04_convocante", "") or ""
        st.session_state["s04_hi"] = det.get("ds04_hora_inicio", "") or ""
        st.session_state["s04_hf"] = det.get("ds04_hora_fin", "") or ""
        st.session_state["s04_obj"] = det.get("ds04_objetivo", "") or ""
        st.session_state["s04_pres"] = det.get("ds04_presentacion", "") or ""
        st.session_state["s04_interv"] = det.get("ds04_intervenciones", "") or ""
        st.session_state["s04_pregs"] = det.get("ds04_preguntas_respuestas", "") or ""
        st.session_state["s04_acuerd"] = det.get("ds04_acuerdos", "") or ""
        st.session_state["s04_obs"] = det.get("ds04_observaciones", "") or ""
        part_json = det.get("ds04_lista_participantes", "") or ""
        if part_json:
            try:
                parts = json.loads(part_json)
                st.session_state["s04_np"] = max(len(parts), 1)
                for i, p in enumerate(parts):
                    st.session_state[f"s04_pn{i}"] = p.get("nombre", "")
                    st.session_state[f"s04_pd{i}"] = p.get("dni", "")
                    st.session_state[f"s04_pi{i}"] = p.get("institucion", "")
                    st.session_state[f"s04_pc{i}"] = p.get("cargo", "")
                    st.session_state[f"s04_pt{i}"] = p.get("telefono", "")
            except (json.JSONDecodeError, TypeError):
                pass

    elif ficha == "F-DS-05":
        conf_json = det.get("ds05_conflictos", "") or ""
        if conf_json:
            try:
                confs = json.loads(conf_json)
                st.session_state["s05_nc"] = max(len(confs), 1)
                for i, c in enumerate(confs):
                    st.session_state[f"s05_ct{i}"] = c.get("tipo", "")
                    st.session_state[f"s05_ca{i}"] = c.get("actores", "")
                    st.session_state[f"s05_cn{i}"] = c.get("nivel", "")
                    st.session_state[f"s05_ce{i}"] = c.get("estado", "")
                    st.session_state[f"s05_cd{i}"] = c.get("descripcion", "")
                    st.session_state[f"s05_ci{i}"] = c.get("impacto", "")
            except (json.JSONDecodeError, TypeError):
                pass
        opor_json = det.get("ds05_oportunidades", "") or ""
        if opor_json:
            try:
                opors = json.loads(opor_json)
                st.session_state["s05_no"] = max(len(opors), 1)
                for i, o in enumerate(opors):
                    st.session_state[f"s05_od{i}"] = o.get("oportunidad", "")
                    st.session_state[f"s05_oa{i}"] = o.get("actores", "")
                    st.session_state[f"s05_ot{i}"] = o.get("tipo", "")
                    st.session_state[f"s05_op{i}"] = o.get("potencial", "")
                    st.session_state[f"s05_oc{i}"] = o.get("como_aprovechar", "")
            except (json.JSONDecodeError, TypeError):
                pass


def pagina_diagnostico_social():
    st.subheader("Diagnostico Social - Fichas de Campo")
    st.caption("Proyecto IN Piura CUI 2669244 | ANIN - DIME - SESDI | Fichas F-DS-01 a F-DS-05")
    bm = _bloques_map()
    if not bm:
        st.warning("Registre un bloque primero.")
        return

    # Inicializar estado de edicion
    if "ds_edit_id" not in st.session_state:
        st.session_state["ds_edit_id"] = None

    ficha_sel = st.radio("Seleccionar ficha", FICHAS_DS, horizontal=True, key="ds_ficha_sel")

    tab_reg, tab_hist, tab_excel = st.tabs(["Registro", "Historial / Consulta", "Importar desde Excel"])

    with tab_reg:
        edit_id = st.session_state.get("ds_edit_id")
        if edit_id:
            st.markdown('<div class="edit-mode-banner"><span class="icon">&#9998;</span> '
                        f'Modo Edicion - Diagnostico Social ID {edit_id}</div>', unsafe_allow_html=True)
            if st.button("Cancelar edicion (nuevo registro)", key="ds_cancel_edit"):
                st.session_state["ds_edit_id"] = None
                st.rerun()

        # ── Datos comunes ─────────────────────────────────────────────
        bl = st.selectbox("Bloque de Intervencion", list(bm.keys()), key="ds_bl")
        bid = bm[bl]

        # Detectar cambio de bloque para resetear campos de ubicacion auto-completados
        prev_bl = st.session_state.get("_ds_prev_bl", "")
        if prev_bl and prev_bl != bl and not edit_id:
            for k in ("ds_prov", "ds_dist", "ds_cpob", "ds_ccam", "ds_este", "ds_norte"):
                st.session_state.pop(k, None)
        st.session_state["_ds_prev_bl"] = bl

        # Auto-resolver microcuenca del bloque seleccionado
        mc_auto_ds = _resolver_microcuenca(bl)
        if mc_auto_ds:
            mc_idx_ds = MICROCUENCAS.index(mc_auto_ds) + 1  # +1 por el "" inicial
            st.info(f"Microcuenca vinculada automaticamente: **{mc_auto_ds}**")
        else:
            mc_idx_ds = 0

        r1, r2, r3, r4 = st.columns(4)
        mc = r1.selectbox("Microcuenca", [""] + MICROCUENCAS, index=mc_idx_ds, key="ds_mc")
        fecha_ev = r2.date_input("Fecha", value=datetime.now(), key="ds_fecha",
                                 min_value=FECHA_MIN_PROYECTO, max_value=date.today(),
                                 help="Seleccione la fecha de evaluacion")
        evaluador = r3.text_input("Responsable", key="ds_eval")
        ficha_num = r4.text_input("Ficha N", key="ds_fnum")
        dg = _ds_datos_generales(bloque_label=bl)
        st.markdown("---")

        datos = {}  # se llenara segun la ficha

        # ══════════════════════════════════════════════════════════════
        if ficha_sel == "F-DS-01":
            st.markdown("### F-DS-01: DIAGNOSTICO SOCIOECONOMICO DE CENTRO POBLADO")
            st.caption("Caracterizacion socioeconomica del centro poblado o localidad del area de intervencion.")
            st.markdown("**1. Datos Demograficos**")
            c1, c2 = st.columns(2)
            ds01_nfam = c1.text_input("N de familias/viviendas", key="s01_nf")
            ds01_nom_cp = c2.text_input("Nombre del centro poblado", key="s01_ncp")
            c3, c4, c5 = st.columns(3)
            ds01_pob_h = c3.text_input("Poblacion Hombres", key="s01_ph")
            ds01_pob_m = c4.text_input("Poblacion Mujeres", key="s01_pm")
            ds01_pob_t = c5.text_input("Poblacion Total", key="s01_pt")
            c6, c7 = st.columns(2)
            ds01_idioma = c6.multiselect("Idioma predominante", FDS01_IDIOMA, key="s01_id")
            ds01_edu = c7.multiselect("Nivel educativo predominante", FDS01_NIVEL_EDUCATIVO, key="s01_ne")
            c8, c9 = st.columns(2)
            ds01_migra = c8.selectbox("Tasa de migracion (percepcion)", [""] + FDS01_TASA_MIGRACION, key="s01_mi")
            ds01_destino = c9.text_input("Destino principal migracion", key="s01_dest")
            c10, c11 = st.columns(2)
            ds01_org = c10.multiselect("Organizacion comunal", FDS01_ORGANIZACION, key="s01_org")
            ds01_junta = c11.selectbox("Junta directiva vigente", ["", "Si", "No"], key="s01_jd")
            ds01_pres = st.text_input("Presidente/a de junta", key="s01_pres")

            st.markdown("**2. Servicios Basicos e Infraestructura**")
            c1, c2 = st.columns(2)
            ds01_agua = c1.multiselect("Agua potable", FDS01_AGUA_POTABLE, key="s01_ag")
            ds01_agua_cob = c2.text_input("Cobertura agua (%)", key="s01_agcob")
            c3, c4 = st.columns(2)
            ds01_sanea = c3.multiselect("Saneamiento", FDS01_SANEAMIENTO, key="s01_san")
            ds01_energ = c4.multiselect("Energia electrica", FDS01_ENERGIA, key="s01_en")
            ds01_energ_cob = st.text_input("Cobertura energia (%)", key="s01_encob")
            c5, c6 = st.columns(2)
            ds01_telec = c5.multiselect("Telecomunicaciones", FDS01_TELECOMUNICACIONES, key="s01_tel")
            ds01_telec_op = c6.text_input("Operador telecom", key="s01_telop")
            c7, c8 = st.columns(2)
            ds01_via = c7.multiselect("Acceso vial", FDS01_ACCESO_VIAL, key="s01_via")
            ds01_dist_cap = c8.text_input("Distancia a capital distrital (km)", key="s01_dcap")
            c9, c10 = st.columns(2)
            ds01_transp = c9.selectbox("Transporte", [""] + FDS01_TRANSPORTE, key="s01_tr")
            ds01_salud = c10.multiselect("Establecimiento de salud", FDS01_SALUD, key="s01_sal")
            c11, c12 = st.columns(2)
            ds01_sal_dist = c11.text_input("Distancia salud (km)", key="s01_sdist")
            ds01_educ = c12.multiselect("Institucion educativa", FDS01_EDUCACION, key="s01_educ")

            st.markdown("**3. Actividades Economicas**")
            st.caption("Registre actividades economicas: Actividad, % Familias, Productos, Destino, Ingreso estimado")
            n_act = st.number_input("N de actividades a registrar", 1, 7, 3, key="s01_nact")
            act_rows = []
            for i in range(int(n_act)):
                cx = st.columns(5)
                a_act = cx[0].selectbox(f"Actividad {i+1}", [""] + FDS01_ACTIVIDADES_ECON, key=f"s01_act{i}")
                a_pct = cx[1].text_input("% Familias", key=f"s01_pct{i}")
                a_prod = cx[2].text_input("Productos", key=f"s01_prod{i}")
                a_dest = cx[3].selectbox("Destino", [""] + FDS01_DESTINO_PRODUCCION, key=f"s01_dest{i}")
                a_ing = cx[4].text_input("Ingreso est.", key=f"s01_ing{i}")
                if a_act:
                    act_rows.append({"actividad": a_act, "pct_familias": a_pct,
                                     "productos": a_prod, "destino": a_dest, "ingreso": a_ing})

            st.markdown("**4. Relacion con Recursos Naturales y Agua**")
            ds01_fuente_agua = st.text_input("Fuente principal de agua", key="s01_fag")
            ds01_prob_agua = st.multiselect("Problemas con el agua", FDS01_PROBLEMAS_AGUA, key="s01_pag")
            c1, c2 = st.columns(2)
            ds01_uso_forest = c1.multiselect("Uso de recursos forestales", FDS01_USO_RECURSOS_FOREST, key="s01_uf")
            ds01_freq_forest = c2.text_input("Frecuencia uso forestal", key="s01_ff")
            ds01_percep_cambios = st.text_area("Percepcion de cambios ambientales", key="s01_pcam", height=80)
            c3, c4 = st.columns(2)
            ds01_disp = c3.selectbox("Disposicion a participar en el proyecto", [""] + FDS01_DISPOSICION, key="s01_disp")
            ds01_com_disp = c4.text_input("Comentario disposicion", key="s01_cdisp")

            st.markdown("**4b. Activos Asociados y Propiedades de Areas a Intervenir**")
            ds01_activos = st.text_area(
                "Activos asociados al bloque (cultivos, ganado, infraestructura de riego, "
                "vias, viviendas, infraestructura publica, etc.)",
                key="s01_activos", height=100,
                help="Dato critico para el analisis de exposicion conforme a la Guia GRD-CC. "
                     "Indique los activos vinculados a los bloques preliminares.")
            st.caption("Propiedades de areas a intervenir (superficie en hectareas por regimen de tenencia)")
            ct1, ct2, ct3 = st.columns(3)
            ds01_ten_com = ct1.text_input("Area comunal (ha)", key="s01_ten_com")
            ds01_ten_pri = ct2.text_input("Area privada (ha)", key="s01_ten_pri")
            ds01_ten_est = ct3.text_input("Area estatal (ha)", key="s01_ten_est")

            datos = {
                "ds01_num_familias": ds01_nfam,
                "ds01_poblacion_hombres": ds01_pob_h, "ds01_poblacion_mujeres": ds01_pob_m,
                "ds01_poblacion_total": ds01_pob_t,
                "ds01_idioma": ", ".join(ds01_idioma), "ds01_nivel_educativo": ", ".join(ds01_edu),
                "ds01_tasa_migracion": ds01_migra, "ds01_destino_migracion": ds01_destino,
                "ds01_organizacion_comunal": ", ".join(ds01_org),
                "ds01_junta_directiva": ds01_junta, "ds01_presidente_junta": ds01_pres,
                "ds01_agua_potable_tipo": ", ".join(ds01_agua), "ds01_agua_potable_cobertura": ds01_agua_cob,
                "ds01_saneamiento": ", ".join(ds01_sanea),
                "ds01_energia_tipo": ", ".join(ds01_energ), "ds01_energia_cobertura": ds01_energ_cob,
                "ds01_telecomunicaciones": ", ".join(ds01_telec), "ds01_telecom_operador": ds01_telec_op,
                "ds01_acceso_vial": ", ".join(ds01_via), "ds01_distancia_capital": ds01_dist_cap,
                "ds01_transporte": ds01_transp,
                "ds01_salud_tipo": ", ".join(ds01_salud), "ds01_salud_distancia": ds01_sal_dist,
                "ds01_educacion": ", ".join(ds01_educ),
                "ds01_actividades_economicas": json.dumps(act_rows, ensure_ascii=False) if act_rows else "",
                "ds01_fuente_agua": ds01_fuente_agua,
                "ds01_problemas_agua": ", ".join(ds01_prob_agua),
                "ds01_uso_recursos_forestales": ", ".join(ds01_uso_forest),
                "ds01_frecuencia_uso_forestal": ds01_freq_forest,
                "ds01_percepcion_cambios": ds01_percep_cambios,
                "ds01_disposicion_participar": ds01_disp,
                "ds01_comentario_disposicion": ds01_com_disp,
                "ds01_activos_asociados": ds01_activos,
                "ds01_tenencia_comunal_ha": ds01_ten_com,
                "ds01_tenencia_privada_ha": ds01_ten_pri,
                "ds01_tenencia_estatal_ha": ds01_ten_est,
            }

        # ══════════════════════════════════════════════════════════════
        elif ficha_sel == "F-DS-02":
            st.markdown("### F-DS-02: IDENTIFICACION Y CARACTERIZACION DE ACTORES CLAVE")
            st.caption("Mapeo de actores del territorio con nivel de influencia e interes.")
            st.markdown("**Registro de Actores Identificados**")
            n_actores = st.number_input("N de actores a registrar", 1, 20, 5, key="s02_nact")
            actores = []
            for i in range(int(n_actores)):
                with st.container():
                    cx = st.columns([2, 1, 2, 1, 1, 1, 2])
                    a_nom = cx[0].text_input(f"Actor {i+1} - Nombre/Organizacion", key=f"s02_nom{i}")
                    a_tipo = cx[1].selectbox("Tipo", [""] + FDS02_TIPO_ACTOR, key=f"s02_tip{i}")
                    a_rol = cx[2].text_input("Rol/Funcion", key=f"s02_rol{i}")
                    a_rel = cx[3].text_input("Rel. Proy.", key=f"s02_rel{i}")
                    a_inf = cx[4].selectbox("Influencia", [""] + FDS02_NIVEL, key=f"s02_inf{i}")
                    a_int = cx[5].selectbox("Interes", [""] + FDS02_NIVEL, key=f"s02_int{i}")
                    a_con = cx[6].text_input("Contacto", key=f"s02_con{i}")
                    if a_nom:
                        actores.append({"nombre": a_nom, "tipo": a_tipo, "rol": a_rol,
                                        "relacion": a_rel, "influencia": a_inf,
                                        "interes": a_int, "contacto": a_con})

            st.markdown("---")
            st.markdown("**Clasificacion por Tipo de Actor**")
            clasif = {}
            for cat in FDS02_CLASIFICACION:
                clasif[cat] = st.text_input(f"{cat} - Actores identificados:", key=f"s02_cl_{cat[:15]}")

            datos = {
                "ds02_registro_actores": json.dumps(actores, ensure_ascii=False) if actores else "",
            }
            for cat in FDS02_CLASIFICACION:
                campo = "ds02_actores_" + {
                    "Gobierno Local": "gob_local",
                    "Gobierno Regional": "gob_regional",
                    "Gobierno Nacional": "gob_nacional",
                    "Comunidades Camp": "comunidades",
                    "Juntas de Usuar": "juntas_riego",
                    "Comites de Gesti": "comites_cuenca",
                    "ONG / Cooperacio": "ong",
                    "Empresa Privada": "empresa",
                    "Instituciones Ed": "educacion",
                    "Organizaciones d": "org_base",
                }.get(cat[:16], "gob_local")
                datos[campo] = clasif[cat]

        # ══════════════════════════════════════════════════════════════
        elif ficha_sel == "F-DS-03":
            st.markdown("### F-DS-03: GUIA DE ENTREVISTA SEMIESTRUCTURADA A ACTORES")
            st.caption("Registro de entrevistas a actores clave del territorio.")
            st.markdown("**Datos del Entrevistado**")
            c1, c2 = st.columns(2)
            ds03_nombre = c1.text_input("Nombre del entrevistado/a", key="s03_nom")
            ds03_cargo = c2.text_input("Cargo / Funcion", key="s03_car")
            c3, c4 = st.columns(2)
            ds03_inst = c3.text_input("Institucion / Organizacion", key="s03_inst")
            ds03_tel = c4.text_input("Telefono / Correo", key="s03_tel")
            ds03_dur = st.text_input("Duracion de la entrevista", key="s03_dur")

            st.markdown("---")
            st.markdown("**1. Percepcion del Territorio y Recursos Naturales**")
            ds03_r1 = st.text_area("1.1 Cuales considera que son los principales recursos naturales de esta zona?", key="s03_r1", height=80)
            ds03_r2 = st.text_area("1.2 Ha observado cambios en el agua, bosques o suelos en los ultimos 10-20 anios?", key="s03_r2", height=80)
            ds03_r3 = st.text_area("1.3 Cuales son los principales problemas ambientales que enfrenta la comunidad?", key="s03_r3", height=80)
            ds03_r4 = st.text_area("1.4 Que zonas considera mas importantes para la conservacion del agua?", key="s03_r4", height=80)

            st.markdown("**2. Actividades Productivas y Medios de Vida**")
            ds03_r5 = st.text_area("2.1 Cuales son las principales actividades economicas de la zona?", key="s03_r5", height=80)
            ds03_r6 = st.text_area("2.2 Como se abastecen de agua para riego y consumo? Es suficiente?", key="s03_r6", height=80)
            ds03_r7 = st.text_area("2.3 Utilizan productos del bosque? Cuales y con que frecuencia?", key="s03_r7", height=80)
            ds03_r8 = st.text_area("2.4 Existen cadenas productivas organizadas? Cuales?", key="s03_r8", height=80)

            st.markdown("**3. Organizacion Social y Gobernanza**")
            ds03_r9 = st.text_area("3.1 Que organizaciones existen en la comunidad? Cuales son las mas activas?", key="s03_r9", height=80)
            ds03_r10 = st.text_area("3.2 Como se toman las decisiones sobre el uso del territorio y los recursos naturales?", key="s03_r10", height=80)
            ds03_r11 = st.text_area("3.3 Existen conflictos por el uso del agua o la tierra? Entre quienes?", key="s03_r11", height=80)
            ds03_r12 = st.text_area("3.4 Han participado en proyectos similares antes? Cual fue la experiencia?", key="s03_r12", height=80)
            ds03_r12b = st.text_area(
                "3.5 Existe experiencia previa en reforestacion en la zona? "
                "Describa especies utilizadas, superficie intervenida y resultados obtenidos.",
                key="s03_r12b", height=100,
                help="Campo con alta relevancia tecnica para el diseno de medidas de "
                     "infraestructura natural (linea 1 del proyecto).")

            st.markdown("**4. Conocimiento y Expectativas sobre el Proyecto**")
            ds03_r13 = st.text_area("4.1 Tiene conocimiento sobre infraestructura natural o proyectos de restauracion?", key="s03_r13", height=80)
            ds03_r14 = st.text_area("4.2 Que expectativas tiene respecto a un proyecto de esta naturaleza?", key="s03_r14", height=80)
            ds03_r15 = st.text_area("4.3 Estaria dispuesto/a a participar o contribuir? De que manera?", key="s03_r15", height=80)
            ds03_r16 = st.text_area("4.4 Que condiciones o preocupaciones tendria respecto al proyecto?", key="s03_r16", height=80)

            st.markdown("**5. Mecanismos de Retribucion (MERESE)**")
            ds03_r17 = st.text_area("5.1 Conoce los mecanismos de retribucion por servicios ecosistemicos?", key="s03_r17", height=80)
            ds03_r18 = st.text_area("5.2 Quienes serian los principales beneficiarios del servicio de regulacion de riesgos?", key="s03_r18", height=80)
            ds03_r19 = st.text_area("5.3 Que instituciones podrian contribuir economicamente a la conservacion?", key="s03_r19", height=80)
            ds03_r20 = st.text_area("5.4 Existen experiencias previas de pago o compensacion por servicios ambientales?", key="s03_r20", height=80)

            datos = {
                "ds03_nombre_entrevistado": ds03_nombre, "ds03_cargo_funcion": ds03_cargo,
                "ds03_institucion": ds03_inst, "ds03_telefono_correo": ds03_tel,
                "ds03_duracion": ds03_dur,
                "ds03_resp_recursos_naturales": ds03_r1, "ds03_resp_cambios_ambiente": ds03_r2,
                "ds03_resp_problemas_ambientales": ds03_r3, "ds03_resp_zonas_conservacion": ds03_r4,
                "ds03_resp_actividades_economicas": ds03_r5, "ds03_resp_abastecimiento_agua": ds03_r6,
                "ds03_resp_productos_bosque": ds03_r7, "ds03_resp_cadenas_productivas": ds03_r8,
                "ds03_resp_organizaciones": ds03_r9, "ds03_resp_decisiones_territorio": ds03_r10,
                "ds03_resp_conflictos": ds03_r11, "ds03_resp_proyectos_anteriores": ds03_r12,
                "ds03_resp_experiencia_reforestacion": ds03_r12b,
                "ds03_resp_conocimiento_restauracion": ds03_r13, "ds03_resp_expectativas": ds03_r14,
                "ds03_resp_disposicion_participar": ds03_r15, "ds03_resp_condiciones": ds03_r16,
                "ds03_resp_conocimiento_merese": ds03_r17, "ds03_resp_beneficiarios": ds03_r18,
                "ds03_resp_instituciones_contribuyentes": ds03_r19, "ds03_resp_experiencias_pago": ds03_r20,
            }

        # ══════════════════════════════════════════════════════════════
        elif ficha_sel == "F-DS-04":
            st.markdown("### F-DS-04: FORMATO DE ACTA DE TALLER PARTICIPATIVO")
            st.caption("Registro del desarrollo y acuerdos de talleres participativos con la comunidad.")
            c1, c2 = st.columns(2)
            ds04_lugar = c1.text_input("Lugar del taller", key="s04_lug")
            ds04_conv = c2.text_input("Convocante", key="s04_conv")
            c3, c4 = st.columns(2)
            ds04_h_ini = c3.text_input("Hora de inicio", key="s04_hi")
            ds04_h_fin = c4.text_input("Hora de finalizacion", key="s04_hf")
            ds04_obj = st.text_area("Objetivo del taller", key="s04_obj", height=60)

            st.markdown("**1. Lista de Participantes**")
            n_part = st.number_input("N de participantes", 1, 30, 10, key="s04_np")
            participantes = []
            for i in range(int(n_part)):
                cx = st.columns([3, 1, 2, 1, 1])
                p_nom = cx[0].text_input(f"Part. {i+1} - Nombres y Apellidos", key=f"s04_pn{i}")
                p_dni = cx[1].text_input("DNI", key=f"s04_pd{i}")
                p_inst = cx[2].text_input("Inst./Comunidad", key=f"s04_pi{i}")
                p_car = cx[3].text_input("Cargo", key=f"s04_pc{i}")
                p_tel = cx[4].text_input("Telefono", key=f"s04_pt{i}")
                if p_nom:
                    participantes.append({"nombre": p_nom, "dni": p_dni, "institucion": p_inst,
                                          "cargo": p_car, "telefono": p_tel})

            st.markdown("**2. Desarrollo del Taller**")
            ds04_pres = st.text_area("Presentacion del proyecto y objetivos del taller", key="s04_pres", height=100)
            ds04_interv = st.text_area("Principales intervenciones de los participantes", key="s04_interv", height=100)
            ds04_pregs = st.text_area("Preguntas y respuestas", key="s04_pregs", height=100)
            ds04_acuerd = st.text_area("Acuerdos y compromisos", key="s04_acuerd", height=100)
            ds04_obs = st.text_area("Observaciones", key="s04_obs", height=80)

            datos = {
                "ds04_lugar_taller": ds04_lugar, "ds04_hora_inicio": ds04_h_ini,
                "ds04_hora_fin": ds04_h_fin, "ds04_convocante": ds04_conv,
                "ds04_objetivo": ds04_obj,
                "ds04_lista_participantes": json.dumps(participantes, ensure_ascii=False) if participantes else "",
                "ds04_presentacion": ds04_pres, "ds04_intervenciones": ds04_interv,
                "ds04_preguntas_respuestas": ds04_pregs, "ds04_acuerdos": ds04_acuerd,
                "ds04_observaciones": ds04_obs,
            }

        # ══════════════════════════════════════════════════════════════
        elif ficha_sel == "F-DS-05":
            st.markdown("### F-DS-05: IDENTIFICACION DE CONFLICTOS Y OPORTUNIDADES")
            st.caption("Mapeo de conflictos existentes y oportunidades para el proyecto.")
            st.markdown("**1. Identificacion de Conflictos**")
            st.caption("Nivel: A=Alto, M=Medio, B=Bajo")
            n_conf = st.number_input("N de conflictos a registrar", 1, 10, 3, key="s05_nc")
            conflictos = []
            for i in range(int(n_conf)):
                cx = st.columns([2, 2, 1, 1, 3, 2])
                co_tipo = cx[0].text_input(f"Conflicto {i+1} - Tipo", key=f"s05_ct{i}")
                co_act = cx[1].text_input("Actores", key=f"s05_ca{i}")
                co_niv = cx[2].selectbox("Nivel", [""] + FDS05_NIVEL, key=f"s05_cn{i}")
                co_est = cx[3].selectbox("Estado", [""] + FDS05_ESTADO_CONFLICTO, key=f"s05_ce{i}")
                co_desc = cx[4].text_input("Descripcion/Causa", key=f"s05_cd{i}")
                co_imp = cx[5].text_input("Impacto en proyecto", key=f"s05_ci{i}")
                if co_tipo:
                    conflictos.append({"tipo": co_tipo, "actores": co_act, "nivel": co_niv,
                                       "estado": co_est, "descripcion": co_desc, "impacto": co_imp})

            st.markdown("---")
            st.markdown("**2. Identificacion de Oportunidades**")
            n_opor = st.number_input("N de oportunidades a registrar", 1, 10, 3, key="s05_no")
            oportunidades = []
            for i in range(int(n_opor)):
                cx = st.columns([3, 2, 1, 1, 3])
                op_desc = cx[0].text_input(f"Oportunidad {i+1}", key=f"s05_od{i}")
                op_act = cx[1].text_input("Actores relacionados", key=f"s05_oa{i}")
                op_tipo = cx[2].selectbox("Tipo", [""] + FDS05_TIPO_OPORTUNIDAD, key=f"s05_ot{i}")
                op_pot = cx[3].selectbox("Potencial", [""] + FDS05_NIVEL, key=f"s05_op{i}")
                op_como = cx[4].text_input("Como aprovecharla", key=f"s05_oc{i}")
                if op_desc:
                    oportunidades.append({"oportunidad": op_desc, "actores": op_act, "tipo": op_tipo,
                                          "potencial": op_pot, "como_aprovechar": op_como})

            datos = {
                "ds05_conflictos": json.dumps(conflictos, ensure_ascii=False) if conflictos else "",
                "ds05_oportunidades": json.dumps(oportunidades, ensure_ascii=False) if oportunidades else "",
            }

        # ── Observaciones y guardar ───────────────────────────────────
        st.markdown("---")
        observ_gen = st.text_area("Observaciones generales", key="ds_obs")
        adj_files = st.file_uploader(
            "Adjuntar archivos de soporte (PDF, max. 25 MB por archivo)",
            type=["pdf"], accept_multiple_files=True, key="ds_adj_upload")

        btn_label = "Actualizar Diagnostico Social" if edit_id else "Guardar Diagnostico Social"
        if st.button(btn_label, type="primary", key="ds_guardar"):
            if not evaluador:
                st.warning("Ingrese el nombre del responsable.")
            elif not datos:
                st.warning("Complete al menos un campo de la ficha.")
            else:
                try:
                    archivos_guardados = []
                    if adj_files:
                        carpeta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "adjuntos_ds")
                        os.makedirs(carpeta, exist_ok=True)
                        for f in adj_files:
                            if f.size > 25 * 1024 * 1024:
                                st.warning(f"Archivo {f.name} excede 25 MB, omitido.")
                                continue
                            nombre = f"{uuid.uuid4().hex[:8]}_{f.name}"
                            ruta = os.path.join(carpeta, nombre)
                            with open(ruta, "wb") as out:
                                out.write(f.getbuffer())
                            archivos_guardados.append(nombre)

                    reg = {
                        "bloque_id": bid, "ficha": ficha_sel,
                        "ficha_numero": ficha_num, "microcuenca": mc,
                        "fecha_evaluacion": fecha_ev.strftime("%Y-%m-%d"),
                        "evaluador": evaluador,
                        "observaciones_generales": observ_gen,
                    }
                    if archivos_guardados:
                        reg["archivos_adjuntos"] = "|".join(archivos_guardados)
                    reg.update(dg)
                    reg.update(datos)

                    if edit_id:
                        db.actualizar_diagnostico_social(edit_id, reg)
                        _invalidar_cache()
                        st.session_state["ds_edit_id"] = None
                        st.success(f"Ficha {ficha_sel} actualizada correctamente (ID {edit_id}).")
                    else:
                        # Verificar duplicados
                        existentes_ds = db.obtener_diagnosticos_sociales_por_bloque(bid)
                        dup_ds = [e for e in existentes_ds
                                  if e.get("ficha") == ficha_sel
                                  and e.get("fecha_evaluacion") == fecha_ev.strftime("%Y-%m-%d")
                                  and e.get("evaluador") == evaluador]
                        if dup_ds:
                            st.markdown('<div class="dup-warning">Ya existe una ficha '
                                       f'{ficha_sel} para este bloque en '
                                       f'{fecha_ev.strftime("%Y-%m-%d")} por {evaluador}. '
                                       f'Use <b>Editar</b> en Historial para modificarla.</div>',
                                       unsafe_allow_html=True)
                        else:
                            if "archivos_adjuntos" not in reg:
                                reg["archivos_adjuntos"] = ""
                            db.insertar_diagnostico_social(reg)
                            _invalidar_cache()
                            st.success(f"Ficha {ficha_sel} guardada correctamente.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

    # ══════════════════════════════════════════════════════════════════
    # HISTORIAL
    # ══════════════════════════════════════════════════════════════════
    with tab_hist:
        st.markdown("### Historial de Diagnosticos Sociales")
        st.caption("Haga clic en **Editar** para modificar un diagnostico existente y evitar duplicidades.")
        todos_ds = _cached_obtener_todos_diagnosticos_sociales(_cache_version())
        if not todos_ds:
            st.info("No hay diagnosticos sociales registrados.")
        else:
            # Paginacion
            ds_pag, total_pags_ds, pag_actual_ds = _paginar(todos_ds, "pag_ds")
            _controles_paginacion(total_pags_ds, pag_actual_ds, "pag_ds")
            # Tabla con botones de edicion
            header_cols = st.columns([0.4, 0.9, 0.7, 0.8, 0.8, 0.8, 0.8, 0.5])
            for col, h in zip(header_cols, ["ID", "Bloque", "Ficha", "Fecha", "Responsable", "C.Poblado", "Distrito", ""]):
                col.markdown(f"**{h}**")
            st.markdown("---")
            for d in ds_pag:
                row = st.columns([0.4, 0.9, 0.7, 0.8, 0.8, 0.8, 0.8, 0.5])
                row[0].write(d["id"])
                row[1].write(d.get("bloque_codigo", ""))
                row[2].write(d.get("ficha", ""))
                row[3].write(d.get("fecha_evaluacion", ""))
                row[4].write(d.get("evaluador", ""))
                row[5].write(d.get("centro_poblado", "") or "")
                row[6].write(d.get("distrito", "") or "")
                if row[7].button("Editar", key=f"edit_ds_{d['id']}", type="primary"):
                    det = db.obtener_diagnostico_social_por_id(d["id"])
                    if det:
                        st.session_state["ds_edit_id"] = det["id"]
                    st.rerun()
            st.markdown("---")

            st.markdown("---")
            st.markdown("### Detalle de Ficha")
            dm_ds = {f"ID {d['id']} - {d.get('bloque_codigo','')} ({d.get('ficha','')})": d["id"] for d in todos_ds}
            sel_ds = st.selectbox("Seleccionar registro", [""] + list(dm_ds.keys()), key="ds_det")
            if sel_ds and sel_ds in dm_ds:
                det = db.obtener_diagnostico_social_por_id(dm_ds[sel_ds])
                if det:
                    ficha_t = det.get("ficha", "")
                    st.markdown(f"**Ficha:** {ficha_t} | **Bloque:** {det.get('bloque_codigo','')} | "
                                f"**Fecha:** {det.get('fecha_evaluacion','')} | "
                                f"**Responsable:** {det.get('evaluador','')}")
                    if det.get("centro_poblado"):
                        st.markdown(f"**Centro Poblado:** {det['centro_poblado']} | "
                                    f"**Comunidad:** {det.get('comunidad_campesina','') or '-'} | "
                                    f"**Distrito:** {det.get('distrito','') or det.get('provincia','')}")

                    if ficha_t == "F-DS-01":
                        with st.expander("DATOS DEMOGRAFICOS", expanded=True):
                            c1, c2, c3 = st.columns(3)
                            c1.markdown(f"**N Familias:** {det.get('ds01_num_familias','') or '-'}")
                            c2.markdown(f"**Pob. H/M/T:** {det.get('ds01_poblacion_hombres','') or '?'} / {det.get('ds01_poblacion_mujeres','') or '?'} / {det.get('ds01_poblacion_total','') or '?'}")
                            c3.markdown(f"**Idioma:** {det.get('ds01_idioma','') or '-'}")
                            c1.markdown(f"**Educacion:** {det.get('ds01_nivel_educativo','') or '-'}")
                            c2.markdown(f"**Migracion:** {det.get('ds01_tasa_migracion','') or '-'}")
                            c3.markdown(f"**Organizacion:** {det.get('ds01_organizacion_comunal','') or '-'}")
                        with st.expander("SERVICIOS BASICOS", expanded=True):
                            c1, c2 = st.columns(2)
                            c1.markdown(f"**Agua:** {det.get('ds01_agua_potable_tipo','') or '-'} (Cob: {det.get('ds01_agua_potable_cobertura','') or '-'})")
                            c2.markdown(f"**Saneamiento:** {det.get('ds01_saneamiento','') or '-'}")
                            c1.markdown(f"**Energia:** {det.get('ds01_energia_tipo','') or '-'} (Cob: {det.get('ds01_energia_cobertura','') or '-'})")
                            c2.markdown(f"**Telecom:** {det.get('ds01_telecomunicaciones','') or '-'}")
                            c1.markdown(f"**Via:** {det.get('ds01_acceso_vial','') or '-'}")
                            c2.markdown(f"**Salud:** {det.get('ds01_salud_tipo','') or '-'}")
                        act_json = det.get("ds01_actividades_economicas", "")
                        if act_json:
                            try:
                                acts = json.loads(act_json)
                                if acts:
                                    with st.expander("ACTIVIDADES ECONOMICAS", expanded=True):
                                        st.dataframe(pd.DataFrame(acts), use_container_width=True, hide_index=True)
                            except (json.JSONDecodeError, TypeError):
                                pass
                        with st.expander("RECURSOS NATURALES Y AGUA", expanded=True):
                            st.markdown(f"**Fuente agua:** {det.get('ds01_fuente_agua','') or '-'}")
                            st.markdown(f"**Problemas agua:** {det.get('ds01_problemas_agua','') or '-'}")
                            st.markdown(f"**Uso forestal:** {det.get('ds01_uso_recursos_forestales','') or '-'}")
                            st.markdown(f"**Percepcion cambios:** {det.get('ds01_percepcion_cambios','') or '-'}")
                            st.markdown(f"**Disposicion participar:** {det.get('ds01_disposicion_participar','') or '-'}")
                        with st.expander("ACTIVOS ASOCIADOS Y PROPIEDADES DE AREAS", expanded=True):
                            st.markdown(f"**Activos asociados:** {det.get('ds01_activos_asociados','') or '-'}")
                            c1, c2, c3 = st.columns(3)
                            c1.markdown(f"**Area comunal (ha):** {det.get('ds01_tenencia_comunal_ha','') or '-'}")
                            c2.markdown(f"**Area privada (ha):** {det.get('ds01_tenencia_privada_ha','') or '-'}")
                            c3.markdown(f"**Area estatal (ha):** {det.get('ds01_tenencia_estatal_ha','') or '-'}")

                    elif ficha_t == "F-DS-02":
                        act_json = det.get("ds02_registro_actores", "")
                        if act_json:
                            try:
                                acts = json.loads(act_json)
                                if acts:
                                    with st.expander("REGISTRO DE ACTORES", expanded=True):
                                        st.dataframe(pd.DataFrame(acts), use_container_width=True, hide_index=True)
                            except (json.JSONDecodeError, TypeError):
                                pass
                        with st.expander("CLASIFICACION POR TIPO", expanded=True):
                            for campo, label in [
                                ("ds02_actores_gob_local", "Gobierno Local"),
                                ("ds02_actores_gob_regional", "Gobierno Regional"),
                                ("ds02_actores_gob_nacional", "Gobierno Nacional"),
                                ("ds02_actores_comunidades", "Comunidades Campesinas"),
                                ("ds02_actores_juntas_riego", "Juntas de Riego"),
                                ("ds02_actores_comites_cuenca", "Comites de Cuenca"),
                                ("ds02_actores_ong", "ONG / Cooperacion"),
                                ("ds02_actores_empresa", "Empresa Privada"),
                                ("ds02_actores_educacion", "Educacion"),
                                ("ds02_actores_org_base", "Org. de Base"),
                            ]:
                                v = det.get(campo, "") or ""
                                if v:
                                    st.markdown(f"**{label}:** {v}")

                    elif ficha_t == "F-DS-03":
                        st.markdown(f"**Entrevistado:** {det.get('ds03_nombre_entrevistado','') or '-'} | "
                                    f"**Cargo:** {det.get('ds03_cargo_funcion','') or '-'} | "
                                    f"**Institucion:** {det.get('ds03_institucion','') or '-'}")
                        secciones = [
                            ("1. Percepcion del Territorio", [
                                ("ds03_resp_recursos_naturales", "Recursos naturales"),
                                ("ds03_resp_cambios_ambiente", "Cambios ambientales"),
                                ("ds03_resp_problemas_ambientales", "Problemas ambientales"),
                                ("ds03_resp_zonas_conservacion", "Zonas de conservacion"),
                            ]),
                            ("2. Actividades Productivas", [
                                ("ds03_resp_actividades_economicas", "Actividades economicas"),
                                ("ds03_resp_abastecimiento_agua", "Abastecimiento agua"),
                                ("ds03_resp_productos_bosque", "Productos del bosque"),
                                ("ds03_resp_cadenas_productivas", "Cadenas productivas"),
                            ]),
                            ("3. Organizacion Social", [
                                ("ds03_resp_organizaciones", "Organizaciones"),
                                ("ds03_resp_decisiones_territorio", "Decisiones territorio"),
                                ("ds03_resp_conflictos", "Conflictos"),
                                ("ds03_resp_proyectos_anteriores", "Proyectos anteriores"),
                                ("ds03_resp_experiencia_reforestacion", "Experiencia en reforestacion"),
                            ]),
                            ("4. Conocimiento y Expectativas", [
                                ("ds03_resp_conocimiento_restauracion", "Conocimiento restauracion"),
                                ("ds03_resp_expectativas", "Expectativas"),
                                ("ds03_resp_disposicion_participar", "Disposicion participar"),
                                ("ds03_resp_condiciones", "Condiciones/Preocupaciones"),
                            ]),
                            ("5. MERESE", [
                                ("ds03_resp_conocimiento_merese", "Conocimiento MERESE"),
                                ("ds03_resp_beneficiarios", "Beneficiarios"),
                                ("ds03_resp_instituciones_contribuyentes", "Instituciones contribuyentes"),
                                ("ds03_resp_experiencias_pago", "Experiencias pago SA"),
                            ]),
                        ]
                        for titulo, campos in secciones:
                            with st.expander(titulo, expanded=True):
                                for campo, label in campos:
                                    v = det.get(campo, "") or ""
                                    if v:
                                        st.markdown(f"**{label}:** {v}")

                    elif ficha_t == "F-DS-04":
                        st.markdown(f"**Lugar:** {det.get('ds04_lugar_taller','') or '-'} | "
                                    f"**Hora:** {det.get('ds04_hora_inicio','') or '-'} - {det.get('ds04_hora_fin','') or '-'} | "
                                    f"**Convocante:** {det.get('ds04_convocante','') or '-'}")
                        if det.get("ds04_objetivo"):
                            st.markdown(f"**Objetivo:** {det['ds04_objetivo']}")
                        part_json = det.get("ds04_lista_participantes", "")
                        if part_json:
                            try:
                                parts = json.loads(part_json)
                                if parts:
                                    with st.expander(f"LISTA DE PARTICIPANTES ({len(parts)})", expanded=True):
                                        st.dataframe(pd.DataFrame(parts), use_container_width=True, hide_index=True)
                            except (json.JSONDecodeError, TypeError):
                                pass
                        with st.expander("DESARROLLO DEL TALLER", expanded=True):
                            for campo, label in [
                                ("ds04_presentacion", "Presentacion"),
                                ("ds04_intervenciones", "Intervenciones"),
                                ("ds04_preguntas_respuestas", "Preguntas y respuestas"),
                                ("ds04_acuerdos", "Acuerdos y compromisos"),
                                ("ds04_observaciones", "Observaciones"),
                            ]:
                                v = det.get(campo, "") or ""
                                if v:
                                    st.markdown(f"**{label}:** {v}")

                    elif ficha_t == "F-DS-05":
                        conf_json = det.get("ds05_conflictos", "")
                        if conf_json:
                            try:
                                confs = json.loads(conf_json)
                                if confs:
                                    with st.expander(f"CONFLICTOS IDENTIFICADOS ({len(confs)})", expanded=True):
                                        st.dataframe(pd.DataFrame(confs), use_container_width=True, hide_index=True)
                            except (json.JSONDecodeError, TypeError):
                                pass
                        opor_json = det.get("ds05_oportunidades", "")
                        if opor_json:
                            try:
                                opors = json.loads(opor_json)
                                if opors:
                                    with st.expander(f"OPORTUNIDADES IDENTIFICADAS ({len(opors)})", expanded=True):
                                        st.dataframe(pd.DataFrame(opors), use_container_width=True, hide_index=True)
                            except (json.JSONDecodeError, TypeError):
                                pass

                    if det.get("observaciones_generales"):
                        st.markdown(f"**Observaciones generales:** {det['observaciones_generales']}")

                    if det.get("archivos_adjuntos"):
                        st.markdown("**Archivos adjuntos:**")
                        carpeta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "adjuntos_ds")
                        for archivo in det["archivos_adjuntos"].split("|"):
                            if archivo.strip():
                                ruta = os.path.join(carpeta, archivo.strip())
                                if os.path.exists(ruta):
                                    with open(ruta, "rb") as fp:
                                        st.download_button(
                                            f"Descargar {archivo.strip().split('_', 1)[-1]}",
                                            fp.read(), archivo.strip(), mime="application/pdf",
                                            key=f"ds_dl_{archivo.strip()}")

                    if st.button("Eliminar este registro", key="ds_eliminar"):
                        db.eliminar_diagnostico_social(dm_ds[sel_ds])
                        _invalidar_cache()
                        st.success("Registro eliminado.")
                        st.rerun()

            st.markdown("---")
            st.markdown("### Resumen por Bloque")
            resumen_ds = db.obtener_resumen_diagnosticos_sociales()
            if resumen_ds:
                st.dataframe(pd.DataFrame([{
                    "Bloque": r["codigo"], "Tipo": r["tipo_intervencion"],
                    "Distrito": r["distrito"], "Total Fichas": r["total_fichas"],
                    "Fichas Completadas": r.get("fichas_completadas", "") or "Ninguna",
                } for r in resumen_ds]), use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════════════
    # TAB IMPORTAR DESDE EXCEL
    # ══════════════════════════════════════════════════════════════════
    with tab_excel:
        st.markdown("### Importar Diagnostico Social desde Excel")
        st.caption("Suba un archivo Excel llenado por el tecnico de campo para autocompletar "
                   "los formularios. Puede descargar la plantilla estandarizada para su uso en campo.")

        # ── Descargar plantilla ──────────────────────────────────────
        st.markdown("---")
        st.markdown("**1. Descargar Plantilla para Tecnicos**")
        col_dl1, col_dl2 = st.columns(2)
        fichas_descarga = col_dl1.multiselect(
            "Fichas a incluir en la plantilla",
            FICHAS_DS, default=FICHAS_DS, key="ds_excel_fichas_dl")
        if col_dl2.button("Generar Plantilla Excel", type="secondary", key="ds_gen_plantilla"):
            if fichas_descarga:
                bloques_data_ds = [(b[1], b[2], b[4], b[5]) for b in BLOQUES_124]
                plantilla_bytes = generar_plantilla_ds(fichas_descarga, bloques_data_ds)
                st.session_state["ds_plantilla_bytes"] = plantilla_bytes
                st.success("Plantilla generada correctamente.")

        if st.session_state.get("ds_plantilla_bytes"):
            st.download_button(
                "Descargar Plantilla Excel",
                st.session_state["ds_plantilla_bytes"],
                file_name="Plantilla_Diagnostico_Social_IN_Piura.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="ds_dl_plantilla")

        # ── Subir Excel llenado ──────────────────────────────────────
        st.markdown("---")
        st.markdown("**2. Subir Excel Llenado por el Tecnico**")
        st.info("Al cargar el archivo, el sistema leera los datos y autocompletara "
                "el formulario en la pestana **Registro**. Podra revisar y ajustar "
                "antes de guardar.")

        uploaded_excel = st.file_uploader(
            "Seleccionar archivo Excel (.xlsx)",
            type=["xlsx"], key="ds_excel_upload")

        if uploaded_excel is not None:
            try:
                resultados = parsear_excel_ds(uploaded_excel)
                if not resultados:
                    st.error("No se pudieron detectar fichas en el archivo. "
                             "Verifique que el formato sea correcto.")
                else:
                    st.success(f"Se detectaron {len(resultados)} ficha(s) en el archivo.")
                    for i, res in enumerate(resultados):
                        ficha_det = res["ficha"]
                        datos_res = res["datos"]

                        with st.expander(f"Vista previa: {ficha_det}", expanded=True):
                            # Mostrar datos generales
                            cols_prev = st.columns(4)
                            cols_prev[0].markdown(f"**Fecha:** {datos_res.get('fecha', '-')}")
                            cols_prev[1].markdown(f"**Responsable:** {datos_res.get('evaluador', '-')}")
                            cols_prev[2].markdown(f"**Bloque:** {datos_res.get('codigo_bloque', '-')}")
                            cols_prev[3].markdown(f"**Ficha N:** {datos_res.get('ficha_numero', '-')}")

                            cols_prev2 = st.columns(4)
                            cols_prev2[0].markdown(f"**Provincia:** {datos_res.get('provincia', '-')}")
                            cols_prev2[1].markdown(f"**Distrito:** {datos_res.get('distrito', '-')}")
                            cols_prev2[2].markdown(f"**C. Poblado:** {datos_res.get('centro_poblado', '-')}")
                            cols_prev2[3].markdown(f"**Microcuenca:** {datos_res.get('microcuenca', '-')}")

                            # Contar campos llenados
                            campos_llenos = sum(1 for k, v in datos_res.items()
                                                if v and str(v).strip() and
                                                k not in ("fecha", "evaluador", "codigo_bloque",
                                                          "ficha_numero", "provincia", "distrito",
                                                          "centro_poblado", "microcuenca",
                                                          "comunidad_campesina", "coordenada_este",
                                                          "coordenada_norte", "altitud",
                                                          "codigo_ubigeo", "observaciones"))
                            st.markdown(f"**Campos especificos llenados:** {campos_llenos}")

                            # Mostrar datos especificos segun ficha
                            if ficha_det == "F-DS-01":
                                acts = datos_res.get("ds01_actividades_economicas", [])
                                if acts:
                                    st.markdown(f"**Actividades economicas registradas:** {len(acts)}")
                                    st.dataframe(pd.DataFrame(acts), use_container_width=True, hide_index=True)
                            elif ficha_det == "F-DS-02":
                                actores = datos_res.get("ds02_registro_actores", [])
                                if actores:
                                    st.markdown(f"**Actores registrados:** {len(actores)}")
                                    st.dataframe(pd.DataFrame(actores), use_container_width=True, hide_index=True)
                            elif ficha_det == "F-DS-04":
                                parts = datos_res.get("ds04_lista_participantes", [])
                                if parts:
                                    st.markdown(f"**Participantes registrados:** {len(parts)}")
                                    st.dataframe(pd.DataFrame(parts), use_container_width=True, hide_index=True)
                            elif ficha_det == "F-DS-05":
                                confs = datos_res.get("ds05_conflictos", [])
                                opors = datos_res.get("ds05_oportunidades", [])
                                if confs:
                                    st.markdown(f"**Conflictos registrados:** {len(confs)}")
                                    st.dataframe(pd.DataFrame(confs), use_container_width=True, hide_index=True)
                                if opors:
                                    st.markdown(f"**Oportunidades registradas:** {len(opors)}")
                                    st.dataframe(pd.DataFrame(opors), use_container_width=True, hide_index=True)

                        # Boton para autocompletar
                        col_ac1, col_ac2 = st.columns(2)
                        if col_ac1.button(
                            f"Autocompletar formulario {ficha_det}",
                            type="primary", key=f"ds_autocompletar_{i}"):
                            ss_vals = mapear_a_session_state(res, bm)
                            for k, v in ss_vals.items():
                                st.session_state[k] = v
                            st.session_state["ds_edit_id"] = None
                            st.success(f"Formulario {ficha_det} autocompletado. "
                                       f"Cambie a la pestana **Registro** para revisar y guardar.")
                            st.rerun()

                        # Boton para guardar directo
                        if col_ac2.button(
                            f"Guardar directamente {ficha_det}",
                            type="secondary", key=f"ds_guardar_directo_{i}"):
                            try:
                                # Construir registro para BD
                                codigo_bloque = datos_res.get("codigo_bloque", "")
                                bid_excel = None
                                for label, id_val in bm.items():
                                    if codigo_bloque and codigo_bloque in label:
                                        bid_excel = id_val
                                        break
                                if not bid_excel:
                                    bid_excel = list(bm.values())[0]
                                    st.warning(f"Bloque '{codigo_bloque}' no encontrado. "
                                               f"Se asigno al primer bloque disponible.")

                                fecha_str = str(datos_res.get("fecha", ""))
                                if not fecha_str:
                                    fecha_str = datetime.now().strftime("%Y-%m-%d")
                                # Normalizar fecha
                                for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
                                    try:
                                        fecha_str = datetime.strptime(
                                            fecha_str.split(" ")[0], fmt).strftime("%Y-%m-%d")
                                        break
                                    except (ValueError, TypeError):
                                        continue

                                reg = {
                                    "bloque_id": bid_excel,
                                    "ficha": ficha_det,
                                    "ficha_numero": datos_res.get("ficha_numero", ""),
                                    "microcuenca": datos_res.get("microcuenca", ""),
                                    "fecha_evaluacion": fecha_str,
                                    "evaluador": datos_res.get("evaluador", ""),
                                    "provincia": datos_res.get("provincia", ""),
                                    "distrito": datos_res.get("distrito", ""),
                                    "centro_poblado": datos_res.get("centro_poblado", ""),
                                    "comunidad_campesina": datos_res.get("comunidad_campesina", ""),
                                    "coordenada_este": float(datos_res.get("coordenada_este") or 0),
                                    "coordenada_norte": float(datos_res.get("coordenada_norte") or 0),
                                    "altitud": float(datos_res.get("altitud") or 0),
                                    "codigo_ubigeo": datos_res.get("codigo_ubigeo", ""),
                                    "observaciones_generales": datos_res.get("observaciones", ""),
                                    "archivos_adjuntos": "",
                                }

                                # Agregar campos especificos por ficha
                                if ficha_det == "F-DS-01":
                                    acts = datos_res.get("ds01_actividades_economicas", [])
                                    reg.update({
                                        k: v for k, v in datos_res.items()
                                        if k.startswith("ds01_") and k != "ds01_actividades_economicas"
                                           and k != "ds01_nombre_cp"
                                    })
                                    if acts:
                                        reg["ds01_actividades_economicas"] = json.dumps(
                                            acts, ensure_ascii=False)
                                    else:
                                        reg["ds01_actividades_economicas"] = ""
                                elif ficha_det == "F-DS-02":
                                    actores = datos_res.get("ds02_registro_actores", [])
                                    reg.update({
                                        k: v for k, v in datos_res.items()
                                        if k.startswith("ds02_") and k != "ds02_registro_actores"
                                    })
                                    if actores:
                                        reg["ds02_registro_actores"] = json.dumps(
                                            actores, ensure_ascii=False)
                                    else:
                                        reg["ds02_registro_actores"] = ""
                                elif ficha_det == "F-DS-03":
                                    reg.update({
                                        k: v for k, v in datos_res.items()
                                        if k.startswith("ds03_")
                                    })
                                elif ficha_det == "F-DS-04":
                                    parts = datos_res.get("ds04_lista_participantes", [])
                                    reg.update({
                                        k: v for k, v in datos_res.items()
                                        if k.startswith("ds04_") and k != "ds04_lista_participantes"
                                           and k != "ds04_observaciones_taller"
                                    })
                                    reg["ds04_observaciones"] = datos_res.get(
                                        "ds04_observaciones_taller", "")
                                    if parts:
                                        reg["ds04_lista_participantes"] = json.dumps(
                                            parts, ensure_ascii=False)
                                    else:
                                        reg["ds04_lista_participantes"] = ""
                                elif ficha_det == "F-DS-05":
                                    confs = datos_res.get("ds05_conflictos", [])
                                    opors = datos_res.get("ds05_oportunidades", [])
                                    reg["ds05_conflictos"] = json.dumps(
                                        confs, ensure_ascii=False) if confs else ""
                                    reg["ds05_oportunidades"] = json.dumps(
                                        opors, ensure_ascii=False) if opors else ""

                                # Limpiar keys que no son columnas de BD
                                keys_no_db = {"fecha", "evaluador", "codigo_bloque",
                                              "observaciones", "ds01_nombre_cp",
                                              "ds04_observaciones_taller"}
                                for k in keys_no_db:
                                    reg.pop(k, None)

                                db.insertar_diagnostico_social(reg)
                                _invalidar_cache()
                                st.success(f"Ficha {ficha_det} guardada directamente en la BD.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error al guardar: {e}")

            except Exception as e:
                st.error(f"Error al leer el archivo Excel: {e}")

        # ── Importacion masiva ───────────────────────────────────────
        st.markdown("---")
        st.markdown("**3. Importacion Masiva (multiples fichas)**")
        st.caption("Si el archivo Excel contiene multiples hojas (una por ficha), "
                   "puede importar todas a la vez usando el boton de arriba. "
                   "Cada hoja sera detectada automaticamente.")
        st.caption("Tambien puede subir la plantilla original "
                   "(`Plantilla_Diagnostico_Social_IN_Piura_rev_22abril.xlsx`) "
                   "llenada por el tecnico.")



# ══════════════════════════════════════════════════════════════════════════
# ELEMENTOS EXPUESTOS (AdR / RIESGOS)
# ══════════════════════════════════════════════════════════════════════════
def pagina_elementos_expuestos():
    st.subheader("Elementos Expuestos - Analisis de Riesgos (AdR)")
    st.caption("Formatos F-EE-01 a F-EE-07 | Registro de activos expuestos, vulnerabilidad y peligros.")

    # Obtener bloques
    bloques = db.obtener_bloques()
    if not bloques:
        st.warning("No hay bloques registrados. Registre bloques primero.")
        return

    bloques_map = {f"{b['codigo']} - {b['distrito']}": b["id"] for b in bloques}
    bloques_labels = list(bloques_map.keys())
    bloques_data = [(b["codigo"], b.get("microcuenca", ""), b.get("provincia", ""),
                     b["distrito"]) for b in bloques]

    # ── Toolbar ──────────────────────────────────────────────────────────
    col_dl, col_up = st.columns(2)
    with col_dl:
        st.markdown("**Descargar plantilla Excel F-EE**")
        excel_bytes = generar_plantilla_ee(bloques_data)
        st.download_button("Descargar Plantilla F-EE (Excel)", excel_bytes,
                           "Plantilla_Elementos_Expuestos_IN_Piura.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_ee")

    with col_up:
        st.markdown("**Importar Excel llenado**")
        archivo = st.file_uploader("Cargar Excel F-EE llenado", type=["xlsx"], key="up_ee")
        if archivo and st.button("Procesar Excel F-EE", key="btn_proc_ee"):
            try:
                resultados = parsear_excel_ee(archivo.read())
                if resultados:
                    st.success(f"Se procesaron {len(resultados)} fichas: {', '.join(resultados.keys())}")
                    st.session_state["ee_import"] = resultados
                else:
                    st.warning("No se encontraron fichas validas en el archivo.")
            except Exception as e:
                st.error(f"Error al procesar: {e}")

    st.markdown("---")

    # ── Tabs: Registrar / Consultar ──────────────────────────────────────
    tab_reg, tab_con = st.tabs(["Registrar / Editar", "Consultar Registros"])

    with tab_reg:
        fichas_ee = ["F-EE-01", "F-EE-02", "F-EE-03", "F-EE-04", "F-EE-05", "F-EE-06", "F-EE-07"]
        ficha_sel = st.selectbox("Seleccionar Ficha", fichas_ee, key="ee_ficha_sel")

        # Bloque
        bloque_label = st.selectbox("Bloque de Intervencion", [""] + bloques_labels, key="ee_bl")
        if not bloque_label:
            st.info("Seleccione un bloque para continuar.")
            return

        bloque_id = bloques_map[bloque_label]
        bloque_info = next((b for b in bloques if b["id"] == bloque_id), {})

        # Datos generales
        st.markdown("**Datos de Identificacion**")
        c1, c2 = st.columns(2)
        ee_fecha = c1.text_input("Fecha de campo (DD/MM/AAAA)", key="ee_fecha")
        ee_resp = c2.text_input("Responsable / Brigada", key="ee_resp")
        ee_cp = st.text_input("Centro(s) Poblado(s)", key="ee_cp")
        c3, c4, c5 = st.columns(3)
        ee_este = c3.text_input("Coord. UTM Este (m)", value=str(bloque_info.get("utm_este", "")), key="ee_este")
        ee_norte = c4.text_input("Coord. UTM Norte (m)", value=str(bloque_info.get("utm_norte", "")), key="ee_norte")
        ee_alt = c5.text_input("Altitud (msnm)", value=str(bloque_info.get("altitud", "")), key="ee_alt")

        st.markdown("---")
        datos = {}

        # ── F-EE-01: Inventario ──────────────────────────────────────────
        if ficha_sel == "F-EE-01":
            st.markdown("### F-EE-01: INVENTARIO DE ELEMENTOS EXPUESTOS")
            st.caption("Registro general de activos dentro del area de impacto de cada bloque.")
            n_reg = st.number_input("N de elementos a registrar", 1, 20, 5, key="e01_n")
            registros = []
            for i in range(int(n_reg)):
                with st.expander(f"Elemento {i+1}", expanded=(i < 3)):
                    cx = st.columns([2, 2, 2])
                    tipo = cx[0].selectbox("Tipo", [""] + FEE01_TIPO_ELEMENTO, key=f"e01_t{i}")
                    subtipo = cx[1].text_input("Subtipo/Categoria", key=f"e01_st{i}")
                    nombre = cx[2].text_input("Nombre/Identificador", key=f"e01_nm{i}")
                    cx2 = st.columns([1, 1, 1, 1])
                    ub = cx2[0].selectbox("Ubicacion Peligro", [""] + FEE01_UBICACION_PELIGRO, key=f"e01_ub{i}")
                    estado = cx2[1].selectbox("Estado (B/R/M)", [""] + FEE_ESTADO, key=f"e01_es{i}")
                    dist = cx2[2].text_input("Distancia Bloque (m)", key=f"e01_di{i}")
                    benef = cx2[3].text_input("N Beneficiarios", key=f"e01_be{i}")
                    material = st.text_input("Material predominante", key=f"e01_ma{i}")
                    if tipo:
                        registros.append({"tipo_elemento": tipo, "subtipo": subtipo,
                            "nombre": nombre, "ubicacion_peligro": ub, "estado": estado,
                            "distancia_bloque": dist, "beneficiarios": benef, "material": material})
            datos["ee01_registros"] = json.dumps(registros, ensure_ascii=False) if registros else ""

        # ── F-EE-02: Poblacion y Viviendas ───────────────────────────────
        elif ficha_sel == "F-EE-02":
            st.markdown("### F-EE-02: POBLACION Y VIVIENDAS")
            st.caption("Detalle de centros poblados, poblacion y viviendas en area de peligro.")
            n_reg = st.number_input("N de centros poblados", 1, 15, 3, key="e02_n")
            registros = []
            for i in range(int(n_reg)):
                with st.expander(f"Centro Poblado {i+1}", expanded=(i < 2)):
                    cx = st.columns([2, 1, 1, 1])
                    cp_nom = cx[0].text_input("Nombre", key=f"e02_cp{i}")
                    viv_t = cx[1].text_input("N Viviendas Total", key=f"e02_vt{i}")
                    viv_p = cx[2].text_input("Viviendas en Peligro", key=f"e02_vp{i}")
                    pob_t = cx[3].text_input("Poblacion Total", key=f"e02_pt{i}")
                    cx2 = st.columns([1, 1, 1, 1])
                    mat = cx2[0].selectbox("Material Viviendas", [""] + FEE02_MATERIAL_VIV, key=f"e02_mat{i}")
                    agua = cx2[1].selectbox("Agua Potable", [""] + FEE_SI_NO, key=f"e02_ag{i}")
                    elec = cx2[2].selectbox("Electricidad", [""] + FEE_SI_NO, key=f"e02_el{i}")
                    antec = cx2[3].selectbox("Antecedente Evento", [""] + FEE_SI_NO, key=f"e02_ant{i}")
                    cx3 = st.columns([1, 1])
                    nd = cx3[0].selectbox("Nivel Danio Anterior", [""] + FEE_NIVEL_AMB, key=f"e02_nd{i}")
                    obs = cx3[1].text_input("Observaciones", key=f"e02_obs{i}")
                    if cp_nom:
                        registros.append({"centro_poblado": cp_nom, "viviendas_total": viv_t,
                            "viviendas_peligro": viv_p, "poblacion_total": pob_t,
                            "material_viviendas": mat, "agua_potable": agua,
                            "electricidad": elec, "antecedente_evento": antec,
                            "nivel_danio": nd, "observaciones": obs})
            datos["ee02_registros"] = json.dumps(registros, ensure_ascii=False) if registros else ""

        # ── F-EE-03: Infraestructura Publica ─────────────────────────────
        elif ficha_sel == "F-EE-03":
            st.markdown("### F-EE-03: INFRAESTRUCTURA PUBLICA EXPUESTA")
            st.caption("I.E., EESS, locales comunales, infraestructura vial, riego, saneamiento.")
            n_reg = st.number_input("N de infraestructuras", 1, 15, 3, key="e03_n")
            registros = []
            for i in range(int(n_reg)):
                with st.expander(f"Infraestructura {i+1}", expanded=(i < 2)):
                    cx = st.columns([1, 2, 2])
                    sector = cx[0].selectbox("Sector", [""] + FEE03_SECTOR, key=f"e03_sec{i}")
                    tipo_inf = cx[1].text_input("Tipo Infraestructura", key=f"e03_ti{i}")
                    nombre = cx[2].text_input("Nombre/Identificador", key=f"e03_nm{i}")
                    cx2 = st.columns([1, 1, 1, 1])
                    estado = cx2[0].selectbox("Estado", [""] + FEE_ESTADO, key=f"e03_es{i}")
                    nexp = cx2[1].selectbox("Nivel Exposicion", [""] + FEE_NIVEL_AMB, key=f"e03_ne{i}")
                    tpel = cx2[2].selectbox("Tipo Peligro", [""] + FEE_TIPO_PELIGRO, key=f"e03_tp{i}")
                    antd = cx2[3].selectbox("Antecedente Danio", [""] + FEE_SI_NO, key=f"e03_ad{i}")
                    cx3 = st.columns([1, 1])
                    costo = cx3[0].text_input("Costo Estimado Activo (S/)", key=f"e03_ca{i}")
                    crepo = cx3[1].text_input("Costo Reposicion (S/)", key=f"e03_cr{i}")
                    if sector or nombre:
                        registros.append({"sector": sector, "tipo_infraestructura": tipo_inf,
                            "nombre": nombre, "estado": estado, "nivel_exposicion": nexp,
                            "tipo_peligro": tpel, "antecedente_danio": antd,
                            "costo_activo": costo, "costo_reposicion": crepo})
            datos["ee03_registros"] = json.dumps(registros, ensure_ascii=False) if registros else ""

        # ── F-EE-04: Actividades Economicas ──────────────────────────────
        elif ficha_sel == "F-EE-04":
            st.markdown("### F-EE-04: ACTIVIDADES ECONOMICAS Y AGROPECUARIAS")
            st.caption("Parcelas agricolas, ganado, actividades productivas expuestas.")
            n_reg = st.number_input("N de actividades", 1, 14, 3, key="e04_n")
            registros = []
            for i in range(int(n_reg)):
                with st.expander(f"Actividad {i+1}", expanded=(i < 2)):
                    cx = st.columns([1, 2, 1])
                    tipo_act = cx[0].selectbox("Tipo", [""] + FEE04_TIPO_ACTIVIDAD, key=f"e04_ta{i}")
                    desc = cx[1].text_input("Descripcion/Cultivo Principal", key=f"e04_de{i}")
                    area = cx[2].text_input("Area (ha)", key=f"e04_ar{i}")
                    cx2 = st.columns([1, 1, 1, 1])
                    fam = cx2[0].text_input("N Familias Depend.", key=f"e04_fa{i}")
                    val_prod = cx2[1].text_input("Valor Prod. Anual (S/)", key=f"e04_vp{i}")
                    nexp = cx2[2].selectbox("Nivel Exposicion", [""] + FEE_NIVEL_AMB, key=f"e04_ne{i}")
                    tpel = cx2[3].selectbox("Tipo Peligro", [""] + FEE_TIPO_PELIGRO, key=f"e04_tp{i}")
                    cx3 = st.columns([1, 1])
                    perd = cx3[0].selectbox("Perdidas Anteriores", [""] + FEE_SI_NO, key=f"e04_pe{i}")
                    monto = cx3[1].text_input("Monto Perdida (S/)", key=f"e04_mp{i}")
                    if tipo_act:
                        registros.append({"tipo_actividad": tipo_act, "descripcion": desc,
                            "area_ha": area, "familias_dependientes": fam,
                            "valor_produccion": val_prod, "nivel_exposicion": nexp,
                            "tipo_peligro": tpel, "perdidas_anteriores": perd,
                            "monto_perdida": monto})
            datos["ee04_registros"] = json.dumps(registros, ensure_ascii=False) if registros else ""

        # ── F-EE-05: Ecosistema ──────────────────────────────────────────
        elif ficha_sel == "F-EE-05":
            st.markdown("### F-EE-05: ECOSISTEMA (UP) Y ACTIVOS AMBIENTALES")
            st.caption("Estado del ecosistema, cobertura vegetal, suelos, peligros observados.")

            st.markdown("**A. Caracterizacion del Ecosistema**")
            c1, c2 = st.columns(2)
            e05_eco = c1.text_input("Tipo de Ecosistema (MINAM)", key="e05_eco")
            e05_zv = c2.text_input("Zona de Vida (Holdridge)", key="e05_zv")
            e05_cv = st.text_input("Cobertura Vegetal Predominante", key="e05_cv")
            c3, c4 = st.columns(2)
            e05_pcv = c3.text_input("% Cobertura Vegetal Estimado", key="e05_pcv")
            e05_esp = c4.text_input("Especies Dominantes (listar)", key="e05_esp")
            c5, c6 = st.columns(2)
            e05_deg = c5.selectbox("Evidencia de Degradacion", [""] + FEE_SI_NO, key="e05_deg")
            e05_tdeg = c6.selectbox("Tipo de Degradacion", [""] + FEE05_TIPO_DEGRADACION, key="e05_tdeg")
            c7, c8 = st.columns(2)
            e05_ndeg = c7.selectbox("Nivel de Degradacion (1-5)", ["", "1", "2", "3", "4", "5"], key="e05_ndeg")
            e05_pend = c8.text_input("Pendiente Predominante (%)", key="e05_pend")
            c9, c10 = st.columns(2)
            e05_suelo = c9.text_input("Tipo de Suelo Observado", key="e05_suelo")
            e05_prof = c10.text_input("Profundidad Efectiva Estimada (cm)", key="e05_prof")
            c11, c12 = st.columns(2)
            e05_carc = c11.selectbox("Presencia de Carcavas/Surcos", [""] + FEE_SI_NO, key="e05_carc")
            e05_queb = c12.selectbox("Presencia de Quebrada/Cauce", [""] + FEE_SI_NO, key="e05_queb")
            c13, c14 = st.columns(2)
            e05_nqueb = c13.text_input("Nombre de Quebrada", key="e05_nqueb")
            e05_fag = c14.selectbox("Fuentes de Agua Identificadas", [""] + FEE05_FUENTES_AGUA, key="e05_fag")

            datos.update({
                "ee05_tipo_ecosistema": e05_eco, "ee05_zona_vida": e05_zv,
                "ee05_cobertura_vegetal": e05_cv, "ee05_pct_cobertura": e05_pcv,
                "ee05_especies_dominantes": e05_esp, "ee05_evidencia_degradacion": e05_deg,
                "ee05_tipo_degradacion": e05_tdeg, "ee05_nivel_degradacion": e05_ndeg,
                "ee05_pendiente": e05_pend, "ee05_tipo_suelo": e05_suelo,
                "ee05_profundidad_efectiva": e05_prof, "ee05_presencia_carcavas": e05_carc,
                "ee05_presencia_quebrada": e05_queb, "ee05_nombre_quebrada": e05_nqueb,
                "ee05_fuentes_agua": e05_fag,
            })

            st.markdown("**B. Evidencias de Peligros Observados en Campo**")
            n_pel = st.number_input("N de peligros observados", 1, 8, 2, key="e05_np")
            peligros = []
            for i in range(int(n_pel)):
                with st.expander(f"Peligro {i+1}", expanded=(i < 2)):
                    cx = st.columns([2, 3])
                    tp = cx[0].selectbox("Tipo Peligro", [""] + FEE_TIPO_PELIGRO, key=f"e05p_tp{i}")
                    desc = cx[1].text_input("Descripcion Indicio/Evidencia", key=f"e05p_de{i}")
                    cx2 = st.columns([1, 1, 1])
                    niv = cx2[0].selectbox("Nivel Estimado", [""] + FEE_NIVEL_AMB, key=f"e05p_ni{i}")
                    prob = cx2[1].selectbox("Prob. Recurrencia", [""] + FEE05_PROB_RECURRENCIA, key=f"e05p_pr{i}")
                    actam = cx2[2].text_input("Activos Amenazados", key=f"e05p_aa{i}")
                    if tp:
                        peligros.append({"tipo_peligro": tp, "descripcion": desc,
                            "nivel_estimado": niv, "probabilidad": prob,
                            "activos_amenazados": actam})
            datos["ee05_peligros_observados"] = json.dumps(peligros, ensure_ascii=False) if peligros else ""

        # ── F-EE-06: Resumen Vulnerabilidad ──────────────────────────────
        elif ficha_sel == "F-EE-06":
            st.markdown("### F-EE-06: RESUMEN DE VULNERABILIDAD DEL BLOQUE")
            st.caption("Sintesis de exposicion, fragilidad y resiliencia del bloque.")

            st.markdown("**A. Cuantificacion de Elementos Expuestos**")
            cuant = []
            for idx, elem in enumerate(FEE06_ELEMENTOS):
                cx = st.columns([3, 1, 1, 1, 2])
                cx[0].text_input("Elemento", value=elem, disabled=True, key=f"e06e_lb{idx}")
                gab = cx[1].text_input("Gabinete", key=f"e06e_gab{idx}")
                campo = cx[2].text_input("Campo", key=f"e06e_cam{idx}")
                coinc = cx[3].selectbox("Coincide", [""] + FEE_SI_NO, key=f"e06e_co{idx}")
                obs = cx[4].text_input("Obs.", key=f"e06e_obs{idx}")
                cuant.append({"elemento": elem, "cantidad_gabinete": gab,
                    "cantidad_campo": campo, "coincide": coinc, "observaciones": obs})
            datos["ee06_cuantificacion"] = json.dumps(cuant, ensure_ascii=False)

            st.markdown("**B. Valoracion Cualitativa de Vulnerabilidad**")
            valor = []
            for idx, (factor, descriptor) in enumerate(FEE06_FACTORES):
                cx = st.columns([2, 2, 1, 1, 3])
                cx[0].text_input("Factor", value=factor, disabled=True, key=f"e06v_fc{idx}")
                cx[1].text_input("Descriptor", value=descriptor, disabled=True, key=f"e06v_ds{idx}")
                nivel = cx[2].selectbox("Nivel", [""] + FEE06_NIVEL_VULN, key=f"e06v_nv{idx}")
                peso = cx[3].text_input("Peso", key=f"e06v_pe{idx}")
                just = cx[4].text_input("Justificacion", key=f"e06v_ju{idx}")
                valor.append({"factor": factor, "descriptor": descriptor,
                    "nivel": nivel, "peso": peso, "justificacion": just})
            datos["ee06_valoracion_vulnerabilidad"] = json.dumps(valor, ensure_ascii=False)

            st.markdown("---")
            c1, c2 = st.columns(2)
            e06_nvuln = c1.selectbox("NIVEL DE VULNERABILIDAD DEL BLOQUE",
                                      [""] + FEE06_NIVEL_VULN, key="e06_nvuln")
            e06_nriesgo = c2.selectbox("NIVEL DE RIESGO PRELIMINAR DEL BLOQUE",
                                        [""] + FEE06_NIVEL_VULN, key="e06_nriesgo")
            datos["ee06_nivel_vulnerabilidad"] = e06_nvuln
            datos["ee06_nivel_riesgo"] = e06_nriesgo

        # ── F-EE-07: Control Fotografico ─────────────────────────────────
        elif ficha_sel == "F-EE-07":
            st.markdown("### F-EE-07: CONTROL DE REGISTRO FOTOGRAFICO")
            st.caption("Inventario de fotografias georeferenciadas vinculadas a elementos expuestos.")
            n_reg = st.number_input("N de fotos a registrar", 1, 25, 5, key="e07_n")
            registros = []
            for i in range(int(n_reg)):
                with st.expander(f"Foto {i+1}", expanded=(i < 3)):
                    cx = st.columns([1, 1, 1, 2])
                    cod = cx[0].text_input("Codigo Foto", key=f"e07_cf{i}")
                    fecha = cx[1].text_input("Fecha", key=f"e07_fe{i}")
                    hora = cx[2].text_input("Hora", key=f"e07_ho{i}")
                    elem = cx[3].text_input("Elemento Fotografiado", key=f"e07_el{i}")
                    cx2 = st.columns([1, 2])
                    ref = cx2[0].text_input("Formato Ref. (F-EE-XX)", key=f"e07_rf{i}")
                    desc = cx2[1].text_input("Descripcion", key=f"e07_de{i}")
                    if cod:
                        registros.append({"codigo_foto": cod, "fecha": fecha, "hora": hora,
                            "elemento_fotografiado": elem, "formato_referencia": ref,
                            "descripcion": desc})
            datos["ee07_registros"] = json.dumps(registros, ensure_ascii=False) if registros else ""

        # ── Guardar ──────────────────────────────────────────────────────
        st.markdown("---")
        observaciones = st.text_area("Observaciones generales", key="ee_obs", height=80)

        if st.button("Guardar Ficha", type="primary", key="btn_guardar_ee"):
            if not ee_fecha or not ee_resp:
                st.error("Fecha de campo y Responsable son obligatorios.")
                return

            datos_guardar = {
                "bloque_id": bloque_id,
                "ficha": ficha_sel,
                "fecha_campo": ee_fecha,
                "responsable_brigada": ee_resp,
                "centro_poblado": ee_cp,
                "coordenada_este": ee_este,
                "coordenada_norte": ee_norte,
                "altitud": ee_alt,
                "observaciones_generales": observaciones,
            }
            datos_guardar.update(datos)

            try:
                edit_id = st.session_state.get("ee_edit_id")
                if edit_id:
                    db.actualizar_elementos_expuestos(edit_id, datos_guardar)
                    st.success(f"Ficha {ficha_sel} actualizada correctamente.")
                    st.session_state.pop("ee_edit_id", None)
                else:
                    db.insertar_elementos_expuestos(datos_guardar)
                    st.success(f"Ficha {ficha_sel} guardada correctamente.")
                _invalidar_cache()
            except Exception as e:
                if "uq_ee_bloque_ficha_fecha_responsable" in str(e):
                    st.error("Ya existe un registro con ese bloque/ficha/fecha/responsable.")
                else:
                    st.error(f"Error al guardar: {e}")

    # ── Tab Consultar ────────────────────────────────────────────────────
    with tab_con:
        registros_ee = db.obtener_todos_elementos_expuestos()
        if not registros_ee:
            st.info("No hay fichas de Elementos Expuestos registradas.")
            return

        df = pd.DataFrame(registros_ee)
        cols_show = ["id", "bloque_codigo", "ficha", "fecha_campo", "responsable_brigada"]
        cols_available = [c for c in cols_show if c in df.columns]
        st.dataframe(df[cols_available], use_container_width=True, hide_index=True)

        sel_id = st.selectbox("Seleccionar ficha para ver detalle",
                              df["id"].tolist() if "id" in df.columns else [],
                              format_func=lambda x: f"ID {x} - {df[df['id']==x].iloc[0].get('ficha','')} - {df[df['id']==x].iloc[0].get('bloque_codigo','')}",
                              key="ee_sel_det")

        if sel_id:
            det = db.obtener_elementos_expuestos_por_id(sel_id)
            if det:
                ficha_t = det.get("ficha", "")
                st.markdown(f"**{ficha_t}** | Bloque: {det.get('bloque_codigo','')} | "
                            f"Fecha: {det.get('fecha_campo','')} | Resp: {det.get('responsable_brigada','')}")

                # Mostrar datos de tabla segun ficha
                for json_field in ["ee01_registros", "ee02_registros", "ee03_registros",
                                   "ee04_registros", "ee05_peligros_observados",
                                   "ee06_cuantificacion", "ee06_valoracion_vulnerabilidad",
                                   "ee07_registros"]:
                    val = det.get(json_field, "")
                    if val:
                        try:
                            items = json.loads(val)
                            if items:
                                label = json_field.replace("ee0", "F-EE-0").replace("_", " ").upper()
                                with st.expander(label, expanded=True):
                                    st.dataframe(pd.DataFrame(items), use_container_width=True, hide_index=True)
                        except (json.JSONDecodeError, TypeError):
                            pass

                # F-EE-05 campos individuales
                if ficha_t == "F-EE-05":
                    with st.expander("ECOSISTEMA - CARACTERIZACION", expanded=True):
                        for campo, label in [
                            ("ee05_tipo_ecosistema", "Tipo Ecosistema"),
                            ("ee05_zona_vida", "Zona de Vida"),
                            ("ee05_cobertura_vegetal", "Cobertura Vegetal"),
                            ("ee05_pct_cobertura", "% Cobertura"),
                            ("ee05_evidencia_degradacion", "Evidencia Degradacion"),
                            ("ee05_tipo_degradacion", "Tipo Degradacion"),
                            ("ee05_nivel_degradacion", "Nivel Degradacion"),
                            ("ee05_presencia_carcavas", "Carcavas/Surcos"),
                            ("ee05_presencia_quebrada", "Quebrada/Cauce"),
                            ("ee05_fuentes_agua", "Fuentes Agua"),
                        ]:
                            v = det.get(campo, "") or ""
                            if v:
                                st.markdown(f"**{label}:** {v}")

                # F-EE-06 resumen
                if ficha_t == "F-EE-06":
                    c1, c2 = st.columns(2)
                    c1.metric("Nivel Vulnerabilidad", det.get("ee06_nivel_vulnerabilidad", "-"))
                    c2.metric("Nivel Riesgo Preliminar", det.get("ee06_nivel_riesgo", "-"))

                obs = det.get("observaciones_generales", "")
                if obs:
                    st.markdown(f"**Observaciones:** {obs}")

                # Botones de accion
                c1, c2 = st.columns(2)
                if c1.button("Editar", key="btn_edit_ee"):
                    st.session_state["ee_edit_id"] = sel_id
                    st.session_state["ee_ficha_sel"] = ficha_t
                    st.rerun()
                if c2.button("Eliminar", key="btn_del_ee"):
                    db.eliminar_elementos_expuestos(sel_id)
                    st.success("Ficha eliminada.")
                    _invalidar_cache()
                    st.rerun()


# ══════════════════════════════════════════════════════════════════════════
# PRESUPUESTO
# ══════════════════════════════════════════════════════════════════════════
def pagina_presupuesto():
    st.subheader("Presupuesto y Recursos")
    bm = _bloques_map()
    if not bm: st.warning("Registre un bloque primero."); return

    # Inicializar estado de edicion
    if "pres_edit_id" not in st.session_state:
        st.session_state["pres_edit_id"] = None

    edit_id = st.session_state.get("pres_edit_id")

    if edit_id:
        st.markdown('<div class="edit-mode-banner"><span class="icon">&#9998;</span> '
                    f'Modo Edicion - Partida ID {edit_id}</div>', unsafe_allow_html=True)
        if st.button("Cancelar edicion", key="pres_cancel_edit"):
            st.session_state["pres_edit_id"] = None
            for k in list(st.session_state.keys()):
                if k.startswith("pres_e_"):
                    del st.session_state[k]
            st.rerun()

    bl = st.selectbox("Bloque", list(bm.keys()), key="pres_bl")
    bid = bm[bl]

    def_cat_idx = 0
    def_desc = ""
    def_mp = 0.0
    def_me = 0.0
    def_fu_idx = 0
    if edit_id:
        def_cat = st.session_state.get("pres_e_cat", "")
        if def_cat and def_cat in CATEGORIAS_PRESUPUESTO:
            def_cat_idx = CATEGORIAS_PRESUPUESTO.index(def_cat)
        def_desc = st.session_state.get("pres_e_desc", "")
        def_mp = float(st.session_state.get("pres_e_mp", 0))
        def_me = float(st.session_state.get("pres_e_me", 0))
        def_fu = st.session_state.get("pres_e_fu", "")
        if def_fu and def_fu in FUENTES_FINANCIAMIENTO:
            def_fu_idx = FUENTES_FINANCIAMIENTO.index(def_fu)

    with st.form("form_pres", clear_on_submit=not edit_id):
        cat = st.selectbox("Categoria", CATEGORIAS_PRESUPUESTO, index=def_cat_idx)
        desc = st.text_input("Descripcion", value=def_desc)
        x1,x2 = st.columns(2)
        mp = x1.number_input("Monto planificado (S/)", 0.0, value=def_mp, format="%.2f")
        me = x2.number_input("Monto ejecutado (S/)", 0.0, value=def_me, format="%.2f")
        fu = st.selectbox("Fuente financiamiento", FUENTES_FINANCIAMIENTO, index=def_fu_idx)
        btn_label = "Actualizar Partida" if edit_id else "Guardar Partida"
        guardar = st.form_submit_button(btn_label, type="primary")
    if guardar:
        try:
            if edit_id:
                db.actualizar_presupuesto(edit_id, cat, desc, mp, me, fu)
                _invalidar_cache()
                st.session_state["pres_edit_id"] = None
                for k in list(st.session_state.keys()):
                    if k.startswith("pres_e_"):
                        del st.session_state[k]
                st.success("Partida actualizada."); st.rerun()
            else:
                # Verificar duplicados
                existentes = db.obtener_presupuesto_por_bloque(bid)
                dup = [e for e in existentes if e["categoria"] == cat and e["descripcion"] == desc]
                if dup:
                    st.warning(f"Ya existe una partida '{cat} - {desc}' para este bloque. "
                               f"Use 'Editar' en la tabla para modificarla.")
                else:
                    db.insertar_presupuesto(bid, cat, desc, mp, me, fu)
                    _invalidar_cache()
                    st.success("Partida registrada."); st.rerun()
        except Exception as e: st.error(f"Error: {e}")
    st.markdown("---")
    st.markdown("**Partidas del Bloque**")
    st.caption("Haga clic en **Editar** para modificar una partida existente.")
    pa = db.obtener_presupuesto_por_bloque(bid)
    if pa:
        tp = sum(p["monto_planificado"] for p in pa)
        te = sum(p["monto_ejecutado"] for p in pa)
        # Tabla con botones de edicion
        header_cols = st.columns([0.5, 1.2, 1.2, 1, 1, 0.7, 1, 0.5])
        for col, h in zip(header_cols, ["ID", "Categoria", "Descripcion", "Planificado", "Ejecutado", "%Ejec", "Fuente", ""]):
            col.markdown(f"**{h}**")
        st.markdown("---")
        for p in pa:
            row = st.columns([0.5, 1.2, 1.2, 1, 1, 0.7, 1, 0.5])
            row[0].write(p["id"])
            row[1].write(p["categoria"])
            row[2].write(p["descripcion"])
            row[3].write(f"S/ {p['monto_planificado']:,.2f}")
            row[4].write(f"S/ {p['monto_ejecutado']:,.2f}")
            row[5].write(f"{(p['monto_ejecutado']/p['monto_planificado']*100) if p['monto_planificado']>0 else 0:.1f}%")
            row[6].write(p["fuente_financiamiento"])
            if row[7].button("Editar", key=f"edit_pres_{p['id']}", type="primary"):
                st.session_state["pres_edit_id"] = p["id"]
                st.session_state["pres_e_cat"] = p["categoria"]
                st.session_state["pres_e_desc"] = p["descripcion"]
                st.session_state["pres_e_mp"] = p["monto_planificado"]
                st.session_state["pres_e_me"] = p["monto_ejecutado"]
                st.session_state["pres_e_fu"] = p["fuente_financiamiento"]
                st.rerun()
        st.markdown("---")
        st.info(f"**Subtotal:** Plan S/ {tp:,.2f} | Ejec S/ {te:,.2f} | {(te/tp*100) if tp>0 else 0:.1f}%")
        pm = {f"ID {p['id']} - {p['categoria']}":p["id"] for p in pa}
        sp = st.selectbox("Partida a eliminar",[""]+list(pm.keys()),key="del_pa")
        if sp and sp in pm and st.button("Eliminar partida"):
            db.eliminar_presupuesto(pm[sp]); _invalidar_cache(); st.success("Eliminada."); st.rerun()
    st.markdown("---")
    st.markdown("**Resumen General**")
    rp = db.obtener_resumen_presupuesto()
    if rp:
        st.dataframe(pd.DataFrame([{"Codigo":r["codigo"],"Tipo":r["tipo_intervencion"],
            "Distrito":r["distrito"],
            "Planificado":f"S/ {r['total_planificado']:,.2f}",
            "Ejecutado":f"S/ {r['total_ejecutado']:,.2f}",
            "%Ejec":f"{(r['total_ejecutado']/r['total_planificado']*100) if r['total_planificado']>0 else 0:.1f}%",
            "Partidas":r["num_partidas"]} for r in rp]),
            use_container_width=True, hide_index=True)
        t = db.obtener_presupuesto_total()
        pt,et = t["total_planificado"],t["total_ejecutado"]
        st.success(f"**TOTAL:** Plan S/ {pt:,.2f} | Ejec S/ {et:,.2f} ({(et/pt*100) if pt>0 else 0:.1f}%) | Saldo S/ {pt-et:,.2f}")

# ══════════════════════════════════════════════════════════════════════════
# CRONOGRAMA
# ══════════════════════════════════════════════════════════════════════════
def pagina_cronograma():
    st.subheader("Cronograma de Actividades")
    bm = _bloques_map()
    if not bm: st.warning("Registre un bloque primero."); return

    # Inicializar estado de edicion
    if "crono_edit_id" not in st.session_state:
        st.session_state["crono_edit_id"] = None

    edit_id = st.session_state.get("crono_edit_id")

    if edit_id:
        st.markdown('<div class="edit-mode-banner"><span class="icon">&#9998;</span> '
                    f'Modo Edicion - Actividad ID {edit_id}</div>', unsafe_allow_html=True)
        if st.button("Cancelar edicion", key="crono_cancel_edit"):
            st.session_state["crono_edit_id"] = None
            for k in list(st.session_state.keys()):
                if k.startswith("crono_e_"):
                    del st.session_state[k]
            st.rerun()

    bl = st.selectbox("Bloque", list(bm.keys()), key="crono_bl")
    bid = bm[bl]

    def_act_idx = 0
    def_ip = datetime.now()
    def_fp = datetime.now()
    def_ir = None
    def_fr = None
    def_av = 0.0
    def_ea_idx = 0
    def_re = ""
    def_ob = ""
    if edit_id:
        def_act = st.session_state.get("crono_e_act", "")
        if def_act and def_act in ACTIVIDADES_TIPO:
            def_act_idx = ACTIVIDADES_TIPO.index(def_act)
        try:
            def_ip = datetime.strptime(st.session_state.get("crono_e_ip", ""), "%Y-%m-%d")
        except (ValueError, TypeError):
            pass
        try:
            def_fp = datetime.strptime(st.session_state.get("crono_e_fp", ""), "%Y-%m-%d")
        except (ValueError, TypeError):
            pass
        ir_str = st.session_state.get("crono_e_ir", "") or ""
        fr_str = st.session_state.get("crono_e_fr", "") or ""
        try:
            def_ir = datetime.strptime(ir_str, "%Y-%m-%d").date() if ir_str else None
        except (ValueError, TypeError):
            def_ir = None
        try:
            def_fr = datetime.strptime(fr_str, "%Y-%m-%d").date() if fr_str else None
        except (ValueError, TypeError):
            def_fr = None
        def_av = float(st.session_state.get("crono_e_av", 0))
        def_ea = st.session_state.get("crono_e_ea", "")
        if def_ea and def_ea in ESTADOS_ACTIVIDAD:
            def_ea_idx = ESTADOS_ACTIVIDAD.index(def_ea)
        def_re = st.session_state.get("crono_e_re", "")
        def_ob = st.session_state.get("crono_e_ob", "")

    with st.form("form_crono", clear_on_submit=not edit_id):
        act = st.selectbox("Actividad", ACTIVIDADES_TIPO, index=def_act_idx)
        x1,x2 = st.columns(2)
        ip = x1.date_input("Inicio planificado", value=def_ip,
                           min_value=FECHA_MIN_PROYECTO, max_value=FECHA_MAX_PROYECTO,
                           help="Fecha de inicio planificada")
        fp = x2.date_input("Fin planificado", value=def_fp,
                           min_value=FECHA_MIN_PROYECTO, max_value=FECHA_MAX_PROYECTO,
                           help="Fecha de fin planificada")
        x3,x4 = st.columns(2)
        ir = x3.date_input("Inicio real", value=def_ir,
                           min_value=FECHA_MIN_PROYECTO, max_value=date.today(),
                           help="Fecha de inicio real (dejar vacio si no ha iniciado)")
        fr = x4.date_input("Fin real", value=def_fr,
                           min_value=FECHA_MIN_PROYECTO, max_value=date.today(),
                           help="Fecha de fin real (dejar vacio si no ha finalizado)")
        x5,x6 = st.columns(2)
        av = x5.number_input("Avance %", 0.0, 100.0, def_av)
        ea = x6.selectbox("Estado", ESTADOS_ACTIVIDAD, index=def_ea_idx)
        re = st.text_input("Responsable", value=def_re); ob = st.text_area("Observaciones", value=def_ob)
        btn_label = "Actualizar Actividad" if edit_id else "Guardar Actividad"
        guardar = st.form_submit_button(btn_label, type="primary")
    if guardar:
        ir_str = ir.strftime("%Y-%m-%d") if ir else ""
        fr_str = fr.strftime("%Y-%m-%d") if fr else ""
        try:
            if edit_id:
                db.actualizar_actividad(actividad_id=edit_id, actividad=act,
                    fecha_inicio_plan=ip.strftime("%Y-%m-%d"), fecha_fin_plan=fp.strftime("%Y-%m-%d"),
                    fecha_inicio_real=ir_str, fecha_fin_real=fr_str, porcentaje_avance=av,
                    responsable=re, observaciones=ob, estado=ea)
                _invalidar_cache()
                st.session_state["crono_edit_id"] = None
                for k in list(st.session_state.keys()):
                    if k.startswith("crono_e_"):
                        del st.session_state[k]
                st.success("Actividad actualizada."); st.rerun()
            else:
                db.insertar_actividad(bloque_id=bid, actividad=act,
                    fecha_inicio_plan=ip.strftime("%Y-%m-%d"), fecha_fin_plan=fp.strftime("%Y-%m-%d"),
                    fecha_inicio_real=ir_str, fecha_fin_real=fr_str, porcentaje_avance=av,
                    responsable=re, observaciones=ob, estado=ea)
                _invalidar_cache()
                st.success("Actividad registrada."); st.rerun()
        except Exception as e: st.error(f"Error: {e}")
    st.markdown("---")
    st.markdown("**Actividades del Bloque**")
    st.caption("Haga clic en **Editar** para modificar una actividad existente.")
    acs = db.obtener_actividades_por_bloque(bid)
    if acs:
        header_cols = st.columns([0.4, 1.2, 0.8, 0.8, 0.7, 0.7, 0.6, 0.7, 0.8, 0.5])
        for col, h in zip(header_cols, ["ID", "Actividad", "Inicio", "Fin", "Ini.Real", "Fin Real", "Avance", "Estado", "Resp.", ""]):
            col.markdown(f"**{h}**")
        st.markdown("---")
        for a in acs:
            row = st.columns([0.4, 1.2, 0.8, 0.8, 0.7, 0.7, 0.6, 0.7, 0.8, 0.5])
            row[0].write(a["id"])
            row[1].write(a["actividad"])
            row[2].write(a["fecha_inicio_plan"])
            row[3].write(a["fecha_fin_plan"])
            row[4].write(a["fecha_inicio_real"] or "-")
            row[5].write(a["fecha_fin_real"] or "-")
            row[6].write(f"{a['porcentaje_avance']:.0f}%")
            row[7].write(a["estado"])
            row[8].write(a["responsable"])
            if row[9].button("Editar", key=f"edit_crono_{a['id']}", type="primary"):
                st.session_state["crono_edit_id"] = a["id"]
                st.session_state["crono_e_act"] = a["actividad"]
                st.session_state["crono_e_ip"] = a["fecha_inicio_plan"]
                st.session_state["crono_e_fp"] = a["fecha_fin_plan"]
                st.session_state["crono_e_ir"] = a["fecha_inicio_real"] or ""
                st.session_state["crono_e_fr"] = a["fecha_fin_real"] or ""
                st.session_state["crono_e_av"] = a["porcentaje_avance"]
                st.session_state["crono_e_ea"] = a["estado"]
                st.session_state["crono_e_re"] = a["responsable"]
                st.session_state["crono_e_ob"] = a.get("observaciones", "") or ""
                st.rerun()
        st.markdown("---")
        am = {f"ID {a['id']} - {a['actividad']}":a["id"] for a in acs}
        sa = st.selectbox("Actividad a eliminar",[""]+list(am.keys()),key="del_ac")
        if sa and sa in am and st.button("Eliminar actividad"):
            db.eliminar_actividad(am[sa]); _invalidar_cache(); st.success("Eliminada."); st.rerun()
    st.markdown("---")
    st.markdown("**Cronograma General**")
    fe = st.selectbox("Filtrar estado",["Todos"]+ESTADOS_ACTIVIDAD,key="f_crono")
    ta = db.obtener_todas_actividades()
    if ta:
        if fe != "Todos": ta = [a for a in ta if a.get("estado")==fe]
        st.dataframe(pd.DataFrame([{"Bloque":a.get("bloque_codigo",""),
            "Actividad":a["actividad"],"Inicio":a["fecha_inicio_plan"],
            "Fin":a["fecha_fin_plan"],"Avance":f"{a['porcentaje_avance']:.0f}%",
            "Estado":a["estado"],"Responsable":a["responsable"]} for a in ta]),
            use_container_width=True, hide_index=True)
        rc = db.obtener_resumen_cronograma(); tt = sum(rc.values())
        co = rc.get("Completado",0)
        st.info(f"Total: {tt} | Programadas: {rc.get('Programado',0)} | En ejecucion: {rc.get('En ejecucion',0)} | Completadas: {co} ({(co/tt*100) if tt>0 else 0:.0f}%) | Retrasadas: {rc.get('Retrasado',0)}")

# ══════════════════════════════════════════════════════════════════════════
# GEORREFERENCIACION
# ══════════════════════════════════════════════════════════════════════════
def pagina_georreferenciacion():
    st.subheader("Georreferenciacion")
    bloques = db.obtener_bloques()
    f1,f2 = st.columns(2)
    fe = f1.selectbox("Filtrar estado",["Todos","Pendiente","En progreso","Verificado"],key="ge")
    ft = f2.selectbox("Filtrar tipo",["Todos"]+TIPOS_INTERVENCION,key="gt")
    bf = bloques
    if fe != "Todos": bf = [b for b in bf if b.get("estado")==fe]
    if ft != "Todos": bf = [b for b in bf if b.get("tipo_intervencion")==ft]
    cm,ci = st.columns([3,1])
    with cm:
        if bf:
            md = []
            for b in bf:
                try:
                    utm_e = b.get("utm_este", 0) or 0
                    utm_n = b.get("utm_norte", 0) or 0
                    # Fallback: si las coords en BD son 0, usar las de BLOQUES_124_MAP
                    if utm_e == 0 or utm_n == 0:
                        datos_b79 = BLOQUES_124_MAP.get(b.get("codigo", ""), {})
                        utm_e = datos_b79.get("utm_este", 0)
                        utm_n = datos_b79.get("utm_norte", 0)
                    # Filtrar bloques sin coordenadas validas
                    if utm_e == 0 or utm_n == 0:
                        continue
                    zona_str = b.get("utm_zona", "17S") or "17S"
                    zn = int(zona_str.replace("S","").replace("N",""))
                    he = "S" if "S" in zona_str else "N"
                    la,lo = utm_a_latlon(float(utm_e),float(utm_n),zn,he)
                    # Validar que las coordenadas estan en rango razonable para Peru
                    if not (-20 <= la <= 0 and -82 <= lo <= -68):
                        continue
                    co = COLORES_ESTADO.get(b.get("estado",""),[149,165,166])
                    # Obtener centros poblados y comunidades del bloque
                    cp_info = CENTROS_POBLADOS_BLOQUE.get(b.get("codigo", ""), {})
                    cp_lista = ", ".join(cp_info.get("centros_poblados", [])) if cp_info else "—"
                    cc_lista = ", ".join(cp_info.get("comunidades_campesinas", [])) if cp_info else "—"
                    pob = cp_info.get("poblacion_total", 0) if cp_info else 0
                    md.append({"lat":la,"lon":lo,"codigo":b["codigo"],
                        "tipo":b["tipo_intervencion"],"estado":b["estado"],
                        "area":b["area_hectareas"],"distrito":b.get("distrito",""),
                        "provincia":b.get("provincia",""),
                        "centros_poblados":cp_lista,"comunidades":cc_lista,
                        "poblacion":pob,
                        "r":co[0],"g":co[1],"b":co[2]})
                except (ValueError,KeyError,TypeError): pass
            if md:
                import pydeck as pdk
                df = pd.DataFrame(md)
                layer = pdk.Layer("ScatterplotLayer",data=df,
                    get_position=["lon","lat"],get_color=["r","g","b",200],
                    get_radius=300,pickable=True)
                vs = pdk.ViewState(latitude=df["lat"].mean(),longitude=df["lon"].mean(),zoom=10)
                st.pydeck_chart(pdk.Deck(layers=[layer],initial_view_state=vs,
                    tooltip={"text":"Codigo: {codigo}\nTipo: {tipo}\nEstado: {estado}\nArea: {area} ha\nDistrito: {distrito} ({provincia})\nCC.PP.: {centros_poblados}\nComunidades: {comunidades}\nPoblacion: {poblacion}"}))
                st.caption(":red_circle: Pendiente | :orange_circle: En progreso | :green_circle: Verificado")
                # Tabla de ubicacion con centros poblados y comunidades
                st.markdown("---")
                st.markdown("**Ubicacion, Centros Poblados y Comunidades Campesinas**")
                df_ubic = df[["codigo","distrito","provincia","area","centros_poblados","comunidades","poblacion"]].copy()
                df_ubic.columns = ["Bloque","Distrito","Provincia","Superficie (ha)","Centros Poblados","Comunidades Campesinas","Poblacion"]
                st.dataframe(df_ubic, use_container_width=True, hide_index=True)
            else: st.info("No se pudieron convertir coordenadas.")
        else: st.info("Sin bloques para los filtros.")
    with ci:
        st.markdown("**Resumen**")
        st.metric("Bloques",len(bf))
        st.metric("Area",f"{sum(b.get('area_hectareas',0) for b in bf):.4f} ha")
        st.markdown("---")
        st.markdown("**Conversor UTM <-> Lat/Lon**")
        t1,t2 = st.tabs(["UTM->LatLon","LatLon->UTM"])
        with t1:
            with st.form("conv_u"):
                ce = st.text_input("Este",key="ce"); cn = st.text_input("Norte",key="cn")
                cz = st.text_input("Zona","17S",key="cz")
                if st.form_submit_button("Convertir") and ce and cn:
                    try:
                        zn = int(cz.replace("S","").replace("N",""))
                        he = "S" if "S" in cz.upper() else "N"
                        la,lo = utm_a_latlon(float(ce),float(cn),zn,he)
                        st.success(f"Lat: {la:.8f}\nLon: {lo:.8f}")
                    except: st.error("Valores invalidos.")
        with t2:
            with st.form("conv_l"):
                cl = st.text_input("Latitud",key="cl"); co = st.text_input("Longitud",key="co")
                if st.form_submit_button("Convertir") and cl and co:
                    try:
                        e,n,z = latlon_a_utm(float(cl),float(co))
                        st.success(f"Este: {e:.2f}\nNorte: {n:.2f}\nZona: {z}")
                    except: st.error("Valores invalidos.")

# ══════════════════════════════════════════════════════════════════════════
# ODK / KoBoToolbox
# ══════════════════════════════════════════════════════════════════════════
def pagina_odk():
    st.subheader("ODK / KoBoToolbox")
    st.markdown("### Generar Formulario XLSForm")
    if st.button("Generar Formulario XLSForm", type="primary"):
        try:
            ruta = generar_xlsform()
            with open(ruta,"rb") as f: data = f.read()
            st.download_button("Descargar XLSForm",data,os.path.basename(ruta),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.success("Formulario generado.")
        except Exception as e: st.error(f"Error: {e}")
    st.markdown("---")
    st.markdown("### Importar CSV")
    uf = st.file_uploader("Archivo CSV",type=["csv"])
    if uf:
        with tempfile.NamedTemporaryFile(mode="w",suffix=".csv",delete=False,encoding="utf-8") as tmp:
            tmp.write(uf.read().decode("utf-8-sig")); tp = tmp.name
        try:
            r = importar_csv_odk(tp)
            st.success(f"**{r['total_filas']} registros** | Nuevos: {r['bloques_nuevos']} | Actualizados: {r['bloques_actualizados']} | Inspecciones: {r['inspecciones_creadas']} | Indicadores: {r['indicadores_creados']}")
            if r["errores"]:
                with st.expander(f"{len(r['errores'])} errores"):
                    for e in r["errores"]: st.warning(e)
        except Exception as e: st.error(f"Error: {e}")
        finally: os.unlink(tp)
    enc = ["codigo_bloque","tipo_intervencion","cuenca","distrito","utm_este","utm_norte",
        "utm_zona","area_hectareas","estado","ubicacion_gps","fecha_visita","inspector",
        "condiciones_climaticas","avance_fisico","observaciones","desviaciones",
        "foto_1","foto_2","foto_3","cobertura_vegetal_planificada",
        "cobertura_vegetal_lograda","sobrevivencia_especies","longitud_zanjas",
        "volumen_retencion","codigo_verificacion"]
    buf = io.StringIO(); w = csv.writer(buf); w.writerow(enc)
    w.writerow(["BLQ-001","revegetacion","Cuenca Alta del Rio Piura","Canchaque",
        "622150.50","9436720.30","17S","2.5","pendiente","","2026-01-15","Juan Perez",
        "despejado","45","","","","","","1100","850","78.5","120.5","35.2",""])
    st.download_button("Descargar plantilla CSV",buf.getvalue(),"plantilla_odk.csv","text/csv")
    st.markdown("---")
    st.markdown("### API KoBoToolbox")
    sv = st.selectbox("Servidor",["https://kf.kobotoolbox.org","https://kobo.humanitarianresponse.info"])
    tk = st.text_input("Token API",type="password")
    c1,c2,c3 = st.columns(3)
    if c1.button("Probar Conexion"):
        if not tk: st.warning("Ingrese token.")
        else:
            cl = KoBoClient(sv,tk); ok,msg = cl.test_conexion()
            if ok: st.success("Conexion exitosa"); st.session_state["kobo"]=cl
            else: st.error(msg)
    if c2.button("Listar Formularios"):
        if "kobo" not in st.session_state: st.warning("Pruebe la conexion primero.")
        else:
            try:
                fs = st.session_state["kobo"].listar_formularios()
                if fs: st.dataframe(pd.DataFrame([{"UID":f["uid"],"Nombre":f["nombre"],
                    "Envios":f["envios"],"Estado":"Desplegado" if f["desplegado"] else "Borrador"} for f in fs]),
                    use_container_width=True, hide_index=True)
            except Exception as e: st.error(f"Error: {e}")
    uid = st.text_input("UID del formulario")
    if c3.button("Importar Envios"):
        if not uid or not tk: st.warning("Complete los campos.")
        else:
            try:
                r = importar_desde_kobo(sv,tk,uid)
                st.success(f"Importado: {r['total_filas']} envios | Nuevos: {r['bloques_nuevos']} | Inspecciones: {r['inspecciones_creadas']}")
            except Exception as e: st.error(f"Error: {e}")
    with st.expander("Guia Rapida"):
        st.markdown("""
1. **GENERAR** el formulario XLSForm
2. **SUBIR** a KoBoToolbox o ODK Central
3. **DESPLEGAR** el formulario
4. **RECOLECTAR** datos con KoBoCollect/ODK Collect (sin internet)
5. **SINCRONIZAR** al tener conexion
6. **IMPORTAR** CSV o usar API directa""")

# ══════════════════════════════════════════════════════════════════════════
# REPORTES
# ══════════════════════════════════════════════════════════════════════════
def pagina_reportes():
    st.subheader("Generacion de Reportes")
    bm = _bloques_map()
    st.markdown("### Ficha de Inspeccion (PDF)")
    if bm:
        bl = st.selectbox("Bloque",list(bm.keys()),key="rep_bl")
        if st.button("Generar Ficha PDF", type="primary"):
            try:
                ruta = reports.generar_ficha_pdf(bm[bl])
                with open(ruta,"rb") as f: data = f.read()
                st.download_button("Descargar PDF",data,os.path.basename(ruta),"application/pdf")
                st.success("PDF generado.")
            except Exception as e: st.error(f"Error: {e}")
    else: st.info("Registre bloques primero.")
    st.markdown("---")
    st.markdown("### Tabla Resumen (Excel)")
    if st.button("Generar Resumen Excel"):
        try:
            ruta = reports.generar_resumen_excel()
            with open(ruta,"rb") as f: data = f.read()
            st.download_button("Descargar Excel",data,os.path.basename(ruta),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.success("Excel generado.")
        except Exception as e: st.error(f"Error: {e}")
    st.markdown("---")
    st.markdown("### Tabla UTM para ArcGIS (Excel)")
    if st.button("Generar Excel ArcGIS"):
        try:
            ruta = reports.generar_excel_arcgis()
            with open(ruta,"rb") as f: data = f.read()
            st.download_button("Descargar Excel ArcGIS",data,os.path.basename(ruta),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.success("Archivo generado.")
        except Exception as e: st.error(f"Error: {e}")

# ══════════════════════════════════════════════════════════════════════════
# ROUTER
# ══════════════════════════════════════════════════════════════════════════
# == CONVERSOR PDF -> EXCEL ================================================
# Pegar este bloque completo en streamlit_app.py, antes del bloque ROUTER
# =========================================================================

def pagina_conversor_pdf():
    """Pagina de conversion de reportes PDF a Excel con formato ANIN."""

    st.subheader("Conversor PDF → Excel")
    st.caption(
        "Carga cualquier reporte PDF del proyecto, extrae sus tablas automaticamente, "
        "agrega coordenadas UTM WGS84 Zona 17S y genera un Excel con formato institucional ANIN."
    )

    if not PDF_CONV_OK:
        st.error(
            "El modulo pdf_converter no esta disponible. "
            "Verifica que pdf_converter.py este en el repositorio "
            "y que 'pdfplumber' y 'pyproj' esten en requirements.txt."
        )
        return

    # ── Panel de carga ─────────────────────────────────────────────────
    st.markdown("### 1. Cargar archivo PDF")

    col_upload, col_opts = st.columns([2, 1])

    with col_upload:
        pdf_file = st.file_uploader(
            "Selecciona el PDF a convertir (max. 25 MB)",
            type=["pdf"],
            key="pdfconv_upload",
        )

    with col_opts:
        st.markdown("**Opciones de conversion**")
        opt_dedup = st.checkbox(
            "Eliminar duplicados",
            value=True,
            key="pdfconv_dedup",
            help="Elimina filas duplicadas del resultado final.",
        )
        opt_solo_utm = st.checkbox(
            "Solo filas con coordenadas UTM validas",
            value=False,
            key="pdfconv_solo_utm",
            help="Filtra filas que no tienen coordenadas UTM validas.",
        )
        tipos_disponibles = [
            "Auto-detectar",
            "centros_poblados",
            "bloques_intervencion",
            "areas_conservacion",
            "catastro_minero",
            "areas_degradadas",
            "meteorologico",
            "generico",
        ]
        tipo_forzado = st.selectbox(
            "Tipo de reporte",
            tipos_disponibles,
            key="pdfconv_tipo",
            help="Selecciona manualmente si la deteccion automatica no es correcta.",
        )

    # ── Procesamiento ──────────────────────────────────────────────────
    if pdf_file is None:
        st.info("Carga un archivo PDF para comenzar.")
        _mostrar_tipos_soportados()
        return

    if pdf_file.size > 25 * 1024 * 1024:
        st.error("El archivo excede el limite de 25 MB.")
        return

    pdf_bytes    = pdf_file.read()
    nombre_pdf   = pdf_file.name
    forzar_tipo  = "" if tipo_forzado == "Auto-detectar" else tipo_forzado

    # Ejecutar extraccion con spinner
    with st.spinner(f"Extrayendo tablas de '{nombre_pdf}'..."):
        try:
            resultado = pdfconv.procesar_pdf(
                pdf_bytes,
                nombre_archivo=nombre_pdf,
                forzar_tipo=forzar_tipo,
            )
        except ImportError as e:
            st.error(f"Dependencia faltante: {e}. Agrega 'pdfplumber' a requirements.txt.")
            return
        except Exception as e:
            st.error(f"Error al procesar el PDF: {e}")
            return

    df         = resultado["df"]
    tipo_det   = resultado["tipo"]
    n_tablas   = resultado["n_tablas"]
    texto_prev = resultado["texto_prev"]

    # ── Panel de resultados ────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 2. Resultado de la extraccion")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Tablas encontradas",  n_tablas)
    m2.metric("Registros extraidos", len(df))
    m3.metric("Columnas",            len(df.columns))
    m4.metric("Tipo detectado",      tipo_det.replace("_", " ").title())

    if df.empty:
        st.warning(
            "No se encontraron tablas estructuradas en el PDF. "
            "Verifica que el archivo no sea una imagen escaneada. "
            "Si es un PDF con texto seleccionable, prueba con tipo 'generico'."
        )
        if texto_prev:
            with st.expander("Texto extraido del PDF (primeros 600 caracteres)"):
                st.text(texto_prev)
        return

    # Validacion de coordenadas UTM
    validacion = pdfconv.validar_utm(df)
    if validacion["estado"] == "ok":
        cov = validacion["cobertura"]
        val = validacion["validos"]
        tot = validacion["total"]
        if val == tot:
            st.success(f"✅ Coordenadas UTM Zona 17S: {val}/{tot} registros validos ({cov})")
        elif val > 0:
            st.warning(
                f"⚠️ Coordenadas UTM Zona 17S: {val}/{tot} validas ({cov}). "
                f"{validacion['invalidos']} registros fuera del rango de la cuenca del Piura."
            )
        else:
            st.info("ℹ️ No se encontraron columnas de latitud/longitud para convertir a UTM.")
    else:
        st.info("ℹ️ El reporte no contiene columnas de coordenadas geograficas.")

    # Preview de la tabla
    st.markdown("**Vista previa (primeros 20 registros)**")
    st.dataframe(df.head(20), use_container_width=True, hide_index=True)

    # ── Generacion del Excel ───────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 3. Generar y descargar Excel")

    col_gen, col_info = st.columns([1, 2])

    with col_gen:
        nombre_salida = st.text_input(
            "Nombre del archivo de salida",
            value=nombre_pdf.replace(".pdf", "").replace(".PDF", "") + "_ANIN",
            key="pdfconv_nombre_salida",
        )
        hoja_nombre = st.text_input(
            "Nombre de la hoja de calculo",
            value=tipo_det.replace("_", " ").title()[:31],
            key="pdfconv_hoja",
        )
        btn_generar = st.button(
            "Generar Excel ANIN",
            type="primary",
            key="pdfconv_generar",
        )

    with col_info:
        st.markdown("**El Excel incluira:**")
        st.markdown(
            "- Encabezado institucional ANIN completo\n"
            "- Columnas UTM en verde (si corresponde)\n"
            "- Filas alternas para facilitar lectura\n"
            "- Fila de totales con suma de areas\n"
            "- Paneles congelados desde la fila de datos\n"
            "- Nota metodologica UTM WGS84 Zona 17S al pie"
        )

    if btn_generar:
        opciones = {
            "deduplicar":   opt_dedup,
            "solo_con_utm": opt_solo_utm,
            "hoja_nombre":  hoja_nombre,
        }
        with st.spinner("Generando Excel con formato ANIN..."):
            try:
                excel_bytes = pdfconv.generar_excel(
                    df,
                    tipo=tipo_det,
                    nombre_origen=nombre_pdf,
                    opciones=opciones,
                )
                nombre_xlsx = (nombre_salida.strip() or "reporte") + ".xlsx"

                st.download_button(
                    label=f"⬇️ Descargar {nombre_xlsx}",
                    data=excel_bytes,
                    file_name=nombre_xlsx,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="pdfconv_download",
                )
                st.success(
                    f"Excel generado: {len(df)} registros, "
                    f"{len(df.columns)} columnas, "
                    f"tipo '{tipo_det}'."
                )

                # Registrar conversion en historial (Supabase)
                _registrar_historial_conversion(
                    nombre_archivo=nombre_pdf,
                    tipo=tipo_det,
                    n_registros=len(df),
                    n_columnas=len(df.columns),
                    utm_ok=validacion["estado"] == "ok",
                )

            except Exception as e:
                st.error(f"Error al generar el Excel: {e}")

    # ── Historial de conversiones ──────────────────────────────────────
    st.markdown("---")
    st.markdown("### Historial de conversiones")
    _mostrar_historial()


def _mostrar_tipos_soportados():
    """Muestra un panel informativo de los tipos de reporte soportados."""
    with st.expander("Tipos de reporte soportados", expanded=False):
        st.markdown("""
| Tipo | Se detecta cuando el PDF contiene |
|---|---|
| **Centros Poblados** | columnas NOMCP, CPINEI, UBIGEO, coordenadas lat/lon |
| **Bloques de Intervencion** | codigos tipo M10B1, microcuenca, area_ha |
| **Areas de Conservacion** | ACR, ACP, denominacion, resolucion ministerial |
| **Catastro Minero** | concesion, titular, derecho minero |
| **Areas Degradadas** | categoria de degradacion, ecosistema, superficie |
| **Meteorologico** | estacion, precipitacion, temperatura, altitud |
| **Generico** | cualquier tabla estructurada (fallback) |

**Nota:** Si la deteccion automatica no es correcta, usa el selector de tipo en las opciones.
        """)


def _registrar_historial_conversion(nombre_archivo, tipo, n_registros,
                                    n_columnas, utm_ok):
    """Guarda un registro de la conversion en session_state (historial en memoria)."""
    if "pdfconv_historial" not in st.session_state:
        st.session_state["pdfconv_historial"] = []
    st.session_state["pdfconv_historial"].insert(0, {
        "Fecha/Hora":  datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Archivo PDF": nombre_archivo,
        "Tipo":        tipo.replace("_", " ").title(),
        "Registros":   n_registros,
        "Columnas":    n_columnas,
        "UTM 17S":     "Si" if utm_ok else "No",
    })
    # Mantener solo los ultimos 20
    st.session_state["pdfconv_historial"] = st.session_state["pdfconv_historial"][:20]


def _mostrar_historial():
    """Muestra el historial de conversiones de la sesion actual."""
    historial = st.session_state.get("pdfconv_historial", [])
    if not historial:
        st.caption("Aun no hay conversiones en esta sesion.")
        return
    st.caption(f"Conversiones en esta sesion: {len(historial)}")
    st.dataframe(
        pd.DataFrame(historial),
        use_container_width=True,
        hide_index=True,
    )

# == FIN BLOQUE CONVERSOR ==


# ══════════════════════════════════════════════════════════════════════════
# MIGRACION V4 — PAGINA TEMPORAL
# ══════════════════════════════════════════════════════════════════════════
def pagina_migracion_v4():
    st.subheader("⚙️ Migracion: Reemplazar Bloques por V4")
    st.warning(
        "**ATENCION:** Esta accion eliminara TODOS los bloques actuales de la base de datos "
        "(incluyendo inspecciones, diagnosticos y registros vinculados) "
        "e insertara los 128 bloques del Proyecto IN Piura V4.",
        icon="⚠️",
    )

    confirmar = st.checkbox("Entiendo que se borraran todos los datos existentes y quiero continuar")

    if not confirmar:
        st.info("Marca la casilla de confirmacion para habilitar la migracion.")
        return

    if st.button("🚀 Ejecutar Migracion V4", type="primary"):
        from datetime import datetime
        import database as db

        try:
            # Usar _ConnectionWrapper que si tiene execute/commit/close
            conn = db.get_connection()

            # 1. Contar bloques existentes
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM bloques")
            row = cursor.fetchone()
            eliminados = list(row.values())[0] if hasattr(row, 'values') else row[0]

            # 2. Eliminar todos los bloques usando conn.execute (no cursor)
            conn.execute("DELETE FROM bloques")

            # 3. Insertar 128 bloques V4
            bloques_v4 = [
                ("1",      "C1076-Q9584", 1371.335, "Morropon",    "Salitral",                 ),
                ("2",      "C1096-Q9558",  330.776, "Morropon",    "Chulucanas",               ),
                ("3",      "C1096-Q9545",  459.532, "Ayabaca",     "Frias",                    ),
                ("4",      "C1076-Q9588",  230.143, "Huancabamba", "Huarmaca",                 ),
                ("5",      "C1077-Q9566",  295.861, "Morropon",    "Buenos Aires",             ),
                ("6",      "C1096-Q9547",  168.499, "Ayabaca",     "Frias",                    ),
                ("7",      "C1076-Q9584",  173.482, "Morropon",    "Salitral",                 ),
                ("8",      "C1076-Q9593",  126.981, "Huancabamba", "Huarmaca",                 ),
                ("9",      "C1076-Q9586",  109.882, "Huancabamba", "San Miguel de El Faique",  ),
                ("10",     "C1076-Q9593",  150.25,  "Huancabamba", "Huarmaca",                 ),
                ("11",     "C1086-Q9570",  103.993, "Morropon",    "Santo Domingo",            ),
                ("12",     "C1086-Q9570",  119.587, "Morropon",    "Santo Domingo",            ),
                ("13",     "C1096-Q9564",  100.558, "Morropon",    "Santo Domingo",            ),
                ("14",     "C1086-Q9570",   92.944, "Morropon",    "Santa Catalina de Mossa",  ),
                ("15",     "C1086-Q9570",   86.185, "Morropon",    "Santo Domingo",            ),
                ("16",     "C1076-Q9593",   83.346, "Huancabamba", "Huarmaca",                 ),
                ("17",     "C1086-Q9570",   96.123, "Morropon",    "Santo Domingo",            ),
                ("18",     "C1076-Q9592",  103.637, "Huancabamba", "Huarmaca",                 ),
                ("19",     "C1076-Q9592",   82.098, "Huancabamba", "Huarmaca",                 ),
                ("20",     "C1076-Q9593",   81.354, "Huancabamba", "Huarmaca",                 ),
                ("21",     "C1081-Q9591",   84.244, "Huancabamba", "Canchaque",                ),
                ("22",     "C1096-Q9556",   79.374, "Morropon",    "Chulucanas",               ),
                ("23",     "C1086-Q9570",   81.812, "Morropon",    "Santo Domingo",            ),
                ("24",     "C1081-Q9583",   90.18,  "Huancabamba", "Canchaque",                ),
                ("25",     "C1096-Q9547",   53.515, "Morropon",    "Chulucanas",               ),
                ("26",     "C1081-Q9591",   46.973, "Huancabamba", "Lalaquiz",                 ),
                ("27",     "C1096-Q9556",   78.192, "Ayabaca",     "Frias",                    ),
                ("28",     "C1081-Q9591",   52.977, "Huancabamba", "Canchaque",                ),
                ("29",     "C1076-Q9592",   40.11,  "Huancabamba", "Huarmaca",                 ),
                ("30",     "C1081-Q9591",   34.882, "Huancabamba", "Canchaque",                ),
                ("31",     "C1076-Q9592",   47.588, "Huancabamba", "Huarmaca",                 ),
                ("32",     "C1096-Q9556",   35.85,  "Morropon",    "Chulucanas",               ),
                ("33",     "C1081-Q9590",   29.689, "Huancabamba", "Lalaquiz",                 ),
                ("34",     "C1081-Q9590",   28.458, "Huancabamba", "Lalaquiz",                 ),
                ("35",     "C1076-Q9586",   30.811, "Huancabamba", "San Miguel de El Faique",  ),
                ("36",     "C1096-Q9547",   24.856, "Ayabaca",     "Frias",                    ),
                ("37",     "C1081-Q9591",   35.5,   "Huancabamba", "Lalaquiz",                 ),
                ("38",     "C1081-Q9590",   44.212, "Huancabamba", "Lalaquiz",                 ),
                ("39",     "C1096-Q9564",   68.6,   "Ayabaca",     "Frias",                    ),
                ("40",     "C1081-Q9591",   28.529, "Huancabamba", "Canchaque",                ),
                ("41",     "C1076-Q9592",   64.933, "Huancabamba", "Huarmaca",                 ),
                ("42",     "C1086-Q9570",   80.649, "Morropon",    "Santa Catalina de Mossa",  ),
                ("43",     "C1076-Q9586",   23.176, "Huancabamba", "San Miguel de El Faique",  ),
                ("44",     "C1076-Q9586",   28.238, "Huancabamba", "San Miguel de El Faique",  ),
                ("45",     "C1076-Q9592",   17.626, "Huancabamba", "Huarmaca",                 ),
                ("46",     "C1086-Q9575",   18.872, "Morropon",    "Yamango",                  ),
                ("47",     "C1076-Q9586",   13.554, "Huancabamba", "San Miguel de El Faique",  ),
                ("48",     "C1086-Q9575",    5.755, "Morropon",    "Yamango",                  ),
                ("49",     "C1086-Q9575",   15.844, "Morropon",    "Yamango",                  ),
                ("50",     "C1081-Q9590",   19.272, "Huancabamba", "Lalaquiz",                 ),
                ("51",     "C1076-Q9593",   22.932, "Huancabamba", "Huarmaca",                 ),
                ("52",     "C1086-Q9575",   11.063, "Morropon",    "Yamango",                  ),
                ("53",     "C1076-Q9592",   13.066, "Huancabamba", "Huarmaca",                 ),
                ("54",     "C1081-Q9591",   40.27,  "Huancabamba", "Huancabamba",              ),
                ("55",     "C1086-Q9575",    0.859, "Morropon",    "Yamango",                  ),
                ("56",     "C1096-Q9564",   77.699, "Ayabaca",     "Frias",                    ),
                ("57",     "C1086-Q9576",   25.677, "Morropon",    "Chalaco",                  ),
                ("58",     "C1086-Q9569",   11.48,  "Morropon",    "Santa Catalina de Mossa",  ),
                ("59",     "C1086-Q9570",   13.559, "Morropon",    "Chalaco",                  ),
                ("60",     "C1081-Q9591",   40.14,  "Huancabamba", "Canchaque",                ),
                ("61",     "C1081-Q9590",   26.833, "Huancabamba", "Lalaquiz",                 ),
                ("62",     "C1081-Q9591",   73.837, "Huancabamba", "Canchaque",                ),
                ("63",     "C1081-Q9583",   50.469, "Huancabamba", "Canchaque",                ),
                ("64",     "C1081-Q9591",   35.78,  "Huancabamba", "Huancabamba",              ),
                ("65",     "C1081-Q9591",   53.944, "Huancabamba", "Canchaque",                ),
                ("66",     "C1081-Q9591",  102.358, "Huancabamba", "Canchaque",                ),
                ("67",     "C1076-Q9586",   13.187, "Huancabamba", "San Miguel de El Faique",  ),
                ("68",     "C1076-Q9592",   21.294, "Huancabamba", "Huarmaca",                 ),
                ("69",     "C1076-Q9592",   19.428, "Huancabamba", "Huarmaca",                 ),
                ("70",     "C1076-Q9592",   22.39,  "Huancabamba", "Huarmaca",                 ),
                ("71",     "C1076-Q9593",   10.676, "Huancabamba", "Huarmaca",                 ),
                ("72",     "C1076-Q9592",   12.4,   "Huancabamba", "Huarmaca",                 ),
                ("73",     "C1076-Q9592",   54.3,   "Huancabamba", "Huarmaca",                 ),
                ("74",     "C1076-Q9588",   29.233, "Huancabamba", "Huarmaca",                 ),
                ("75",     "C1076-Q9587",   15.179, "Huancabamba", "San Miguel de El Faique",  ),
                ("76",     "C1076-Q9593",   19.619, "Huancabamba", "Huarmaca",                 ),
                ("77",     "C1081-Q9583",   17.827, "Morropon",    "San Juan de Bigote",       ),
                ("78",     "C1086-Q9570",   35.116, "Morropon",    "Santo Domingo",            ),
                ("79",     "C1076-Q9587",   94.834, "Huancabamba", "Canchaque",                ),
                ("80",     "C1076-Q9593",   28.835, "Huancabamba", "Huarmaca",                 ),
                ("81",     "C1076-Q9585",    8.931, "Huancabamba", "Canchaque",                ),
                ("82",     "C1086-Q9570",   35.677, "Morropon",    "Santo Domingo",            ),
                ("M1B1",   "C1077-Q9580",  188.883, "Morropon",    "Buenos Aires",             ),
                ("M2B1",   "C1076-Q9584",   81.823, "Morropon",    "Salitral",                 ),
                ("M2B5",   "C1076-Q9584",   30.918, "Morropon",    "Salitral",                 ),
                ("M2B8",   "C1076-Q9584",  116.136, "Huancabamba", "San Miguel de El Faique",  ),
                ("M3B1",   "C1081-Q9582",   81.014, "Morropon",    "Salitral",                 ),
                ("M3B3",   "C1081-Q9582",   84.855, "Morropon",    "San Juan de Bigote",       ),
                ("M3B5",   "C1081-Q9582",   52.761, "Morropon",    "San Juan de Bigote",       ),
                ("M3B6",   "C1081-Q9582",   60.372, "Morropon",    "San Juan de Bigote",       ),
                ("M3B7",   "C1081-Q9582",  122.987, "Morropon",    "San Juan de Bigote",       ),
                ("M3B8",   "C1081-Q9582",  565.792, "Morropon",    "San Juan de Bigote",       ),
                ("M3B9",   "C1081-Q9582",  294.318, "Morropon",    "Salitral",                 ),
                ("M4B3",   "C1076-Q9589",  234.137, "Huancabamba", "Huarmaca",                 ),
                ("M4B4",   "C1076-Q9589",  290.0,   "Huancabamba", "Huarmaca",                 ),
                ("M6B2-1", "C1077-Q9566",  514.608, "Morropon",    "Buenos Aires",             ),
                ("M6B2-2", "C1077-Q9566",  333.176, "Morropon",    "Buenos Aires",             ),
                ("M6B2-3", "C1077-Q9566",  161.058, "Morropon",    "Buenos Aires",             ),
                ("M6B10",  "C1077-Q9566",  250.004, "Morropon",    "Buenos Aires",             ),
                ("M7B1",   "C1076-Q9581",  130.475, "Morropon",    "Salitral",                 ),
                ("M7B2",   "C1076-Q9581",  115.422, "Morropon",    "Salitral",                 ),
                ("M7B3",   "C1076-Q9581",   65.768, "Morropon",    "Salitral",                 ),
                ("M7B6",   "C1076-Q9581",   91.317, "Morropon",    "Salitral",                 ),
                ("M8B2",   "C1096-Q9558",   35.363, "Morropon",    "Santo Domingo",            ),
                ("M9B1",   "C1096-Q9545",  355.568, "Morropon",    "Chulucanas",               ),
                ("M10B4",  "C1096-Q9547",   68.926, "Morropon",    "Chulucanas",               ),
                ("M11B3",  "C1081-Q9583",  106.059, "Morropon",    "San Juan de Bigote",       ),
                ("M12B1",  "C1076-Q9588",  268.469, "Huancabamba", "Huarmaca",                 ),
                ("M17B1",  "C1096-Q9556",   42.472, "Morropon",    "Chulucanas",               ),
                ("M17B4",  "C1096-Q9556",  124.775, "Morropon",    "Chulucanas",               ),
                ("M17B5",  "C1096-Q9556",   74.065, "Ayabaca",     "Frias",                    ),
                ("M17B6",  "C1096-Q9556",  106.593, "Ayabaca",     "Frias",                    ),
                ("M17B7",  "C1096-Q9556",  258.969, "Ayabaca",     "Frias",                    ),
                ("M17B10", "C1096-Q9556",   75.403, "Ayabaca",     "Frias",                    ),
                ("M18B1",  "C1077-Q9579",  160.43,  "Morropon",    "Buenos Aires",             ),
                ("M18B3",  "C1077-Q9579",  373.96,  "Morropon",    "Salitral",                 ),
                ("M18B5",  "C1077-Q9579",  197.437, "Morropon",    "Salitral",                 ),
                ("M19B2",  "C1086-Q9570",   74.406, "Morropon",    "Morropon",                 ),
                ("M19B7",  "C1086-Q9570",   34.186, "Morropon",    "Santo Domingo",            ),
                ("M20B1",  "C1076-Q9585",  279.766, "Huancabamba", "San Miguel de El Faique",  ),
                ("M22B1",  "C1076-Q9586",   55.177, "Huancabamba", "San Miguel de El Faique",  ),
                ("M27B1",  "C1096-Q9557",  449.066, "Morropon",    "Chulucanas",               ),
                ("M28B2",  "C1086-Q9575",   90.228, "Morropon",    "Yamango",                  ),
                ("M28B3",  "C1086-Q9575",   58.312, "Morropon",    "Yamango",                  ),
                ("M28B4",  "C1086-Q9575",  283.543, "Morropon",    "Yamango",                  ),
                ("M30B5",  "C1081-Q9591",   90.922, "Huancabamba", "Huancabamba",              ),
                ("M32B3",  "C1086-Q9569",   50.996, "Morropon",    "Morropon",                 ),
                ("M36B2",  "C1086-Q9576",   28.202, "Morropon",    "Santa Catalina de Mossa",  ),
            ]

            conn.commit()
            conn.close()

            # 3. Insertar 128 bloques usando la funcion publica del modulo
            for (codigo, microcuenca, area_ha, provincia, distrito) in bloques_v4:
                db.insertar_bloque(
                    codigo=codigo,
                    tipo_intervencion="Restauracion",
                    cuenca=microcuenca,
                    distrito=distrito,
                    utm_este=0.0,
                    utm_norte=0.0,
                    utm_zona="17S",
                    area_hectareas=area_ha,
                    estado="Pendiente",
                    microcuenca=microcuenca,
                    provincia=provincia,
                )

            st.success(f"✅ Migracion completada: {eliminados} bloques eliminados, 128 bloques V4 insertados.")
            st.info("Recarga la pagina o navega al Panel de Control para ver los cambios.")
            st.cache_data.clear()

        except Exception as e:
            st.error(f"Error durante la migracion: {e}")


# ══════════════════════════════════════════════════════════════════════════
# ROUTER
# ══════════════════════════════════════════════════════════════════════════
if pagina == "Panel de Control": pagina_dashboard()
elif pagina == "Bloques de Intervencion": pagina_bloques()
elif pagina == "Inspeccion de Campo": pagina_inspeccion()
elif pagina == "Indicadores de Calidad": pagina_indicadores()
elif pagina == "Diagnostico Territorial": pagina_diagnostico_territorial()
elif pagina == "Diagnostico Social": pagina_diagnostico_social()
elif pagina == "Elementos Expuestos (AdR)": pagina_elementos_expuestos()
elif pagina == "Presupuesto": pagina_presupuesto()
elif pagina == "Cronograma": pagina_cronograma()
elif pagina == "Georreferenciacion": pagina_georreferenciacion()
elif pagina == "ODK / KoBoToolbox": pagina_odk()
elif pagina == "Reportes": pagina_reportes()
elif pagina == "Conversor PDF -> Excel": pagina_conversor_pdf()
elif pagina == "⚙️ Migracion V4 (temporal)": pagina_migracion_v4()
