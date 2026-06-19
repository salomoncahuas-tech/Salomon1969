"""
IN Piura - Plan de Ingreso / Verificacion de Campo
Aplicacion web con Streamlit.
Restauracion de ecosistemas - Cuenca alta del rio Piura, Peru.
"""

import streamlit as st
import pandas as pd
from datetime import datetime, date, time as dtime, timedelta
import os
import uuid
import io
import csv
import json
import tempfile

import database as db
import export_diagnosticos as exp_diag

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

# ── Mensajes flash (sobreviven a st.rerun) ───────────────────────────────
def _flash(msg, tipo="success"):
    """Guarda un mensaje para mostrarlo tras el siguiente st.rerun()."""
    st.session_state["_flash_msg"] = (tipo, msg)

def _mostrar_flash():
    """Muestra (y descarta) el mensaje flash pendiente, si existe."""
    item = st.session_state.pop("_flash_msg", None)
    if not item:
        return
    tipo, msg = item
    {"success": st.success, "info": st.info,
     "warning": st.warning, "error": st.error}.get(tipo, st.success)(msg)

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

@st.cache_resource
def _inicializar_bd_una_vez():
    db.inicializar_bd()

def _arrancar_con_reintento():
    """Intenta inicializar la BD; si falla muestra error y botón de reintento."""
    try:
        _inicializar_bd_una_vez()
    except Exception as e:
        st.error(
            "**No se pudo conectar a la base de datos.**\n\n"
            "La base de datos puede estar pausada (esto ocurre en el plan gratuito de Supabase "
            "cuando no hay actividad por varios días).\n\n"
            "**¿Qué hacer?**\n"
            "1. Entra a [supabase.com](https://supabase.com) y verifica que tu proyecto esté activo (no pausado).\n"
            "2. Si está pausado, haz clic en **Resume project** y espera 1-2 minutos.\n"
            "3. Luego presiona el botón de abajo para reintentar."
        )
        if st.button("🔄 Reintentar conexión"):
            st.cache_resource.clear()
            st.rerun()
        st.stop()

_arrancar_con_reintento()

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

# ── Datos de Origen: 127 Bloques de Intervencion V5 ──────────────────
# Fuente: Plantilla_DT_Campo_Check_Validada_V5.xlsx (hoja 'Bloques V4')
# Cada entrada: (N, Bloque, Microcuenca, Area_ha, Provincia, Distrito,
#                Accesibilidad, Dia, UTM_Este, UTM_Norte, MSAVI_2024, Zona)
BLOQUES_V5 = [
    # (n, codigo, microcuenca, area_ha, provincia, distrito, accesibilidad, dia_eval, utm_este, utm_norte, msavi_2024, zona)
    (1, "M9B1", "C1096-Q9545", 355.36, "Morropon", "Chulucanas", 0, 0, 595552, 9451222, 0.314312, "Z01"),
    (2, "M10B4", "C1096-Q9547", 68.89, "Morropon", "Chulucanas", 0, 0, 596773, 9448273, 0.313301, "Z01"),
    (3, "25", "C1096-Q9547", 53.48, "Morropon", "Chulucanas", 0, 0, 596578, 9442453, 0.256543, "Z01"),
    (4, "32", "C1096-Q9556", 35.83, "Morropon", "Chulucanas", 0, 0, 599927, 9442604, 0.371619, "Z01"),
    (5, "22", "C1096-Q9556", 79.33, "Morropon", "Chulucanas", 0, 0, 601122, 9443651, 0.42317, "Z01"),
    (6, "M17B1", "C1096-Q9556", 42.45, "Morropon", "Chulucanas", 0, 0, 600644, 9440436, 0.332263, "Z01"),
    (7, "2", "C1096-Q9558", 330.6, "Morropon", "Chulucanas", 0, 0, 603653, 9439351, 0.358601, "Z01"),
    (8, "M17B4", "C1096-Q9556", 124.71, "Morropon", "Chulucanas", 0, 0, 604926, 9444327, 0.423543, "Z01"),
    (9, "M27B1", "C1096-Q9557", 448.8, "Morropon", "Chulucanas", 0, 0, 593246, 9430139, 0.363228, "Z01"),
    (10, "3", "C1096-Q9545", 459.29, "Ayabaca", "Frias", 0, 0, 606816, 9458469, 0.567707, "Z02"),
    (11, "6", "C1096-Q9547", 168.41, "Ayabaca", "Frias", 0, 0, 605636, 9455269, 0.538144, "Z02"),
    (12, "36", "C1096-Q9547", 24.84, "Ayabaca", "Frias", 0, 0, 606468, 9454689, 0.539458, "Z02"),
    (13, "M17B7", "C1096-Q9556", 258.84, "Ayabaca", "Frias", 0, 0, 608472, 9448056, 0.508715, "Z02"),
    (14, "M17B6", "C1096-Q9556", 106.54, "Ayabaca", "Frias", 0, 0, 607036, 9446211, 0.446441, "Z02"),
    (15, "M17B10", "C1096-Q9556", 75.36, "Ayabaca", "Frias", 0, 0, 606173, 9447361, 0.472466, "Z02"),
    (16, "M17B5", "C1096-Q9556", 74.03, "Ayabaca", "Frias", 0, 0, 608800, 9445585, 0.54817, "Z02"),
    (17, "27", "C1096-Q9556", 78.15, "Ayabaca", "Frias", 0, 0, 613302, 9450231, 0.53228, "Z02"),
    (18, "39", "C1096-Q9564", 68.57, "Ayabaca", "Frias", 0, 0, 618523, 9448918, 0.701601, "Z02"),
    (19, "56", "C1096-Q9564", 77.67, "Ayabaca", "Frias", 0, 0, 631723, 9439702, 0.677658, "Z02"),
    (20, "12", "C1086-Q9570", 119.54, "Morropon", "Santo Domingo", 0, 0, 625419, 9444932, 0.706689, "Z03"),
    (21, "82", "C1086-Q9570", 35.66, "Morropon", "Santo Domingo", 0, 0, 621452, 9443855, 0.62835, "Z03"),
    (22, "13", "C1096-Q9564", 100.51, "Morropon", "Santo Domingo", 0, 0, 620263, 9442343, 0.671094, "Z03"),
    (23, "78", "C1086-Q9570", 35.1, "Morropon", "Santo Domingo", 0, 0, 619630, 9440043, 0.707402, "Z03"),
    (24, "11", "C1086-Q9570", 103.95, "Morropon", "Santo Domingo", 0, 0, 620759, 9439988, 0.604743, "Z03"),
    (25, "23", "C1086-Q9570", 81.78, "Morropon", "Santo Domingo", 0, 0, 619939, 9438544, 0.573141, "Z03"),
    (26, "15", "C1086-Q9570", 86.15, "Morropon", "Santo Domingo", 0, 0, 620200, 9437862, 0.560518, "Z03"),
    (27, "M19B7", "C1086-Q9570", 34.17, "Morropon", "Santo Domingo", 0, 0, 620770, 9436562, 0.591921, "Z03"),
    (28, "17", "C1086-Q9570", 96.08, "Morropon", "Santo Domingo", 0, 0, 618661, 9438267, 0.647895, "Z03"),
    (29, "M8B2", "C1096-Q9558", 35.35, "Morropon", "Santo Domingo", 0, 0, 617596, 9436227, 0.530453, "Z03"),
    (30, "57", "C1086-Q9576", 25.67, "Morropon", "Chalaco", 0, 0, 631195, 9441777, 0.603055, "Z04"),
    (31, "59", "C1086-Q9570", 13.55, "Morropon", "Chalaco", 0, 0, 631195, 9441777, 0.614739, "Z04"),
    (32, "M36B2", "C1086-Q9576", 28.19, "Morropon", "Santa Catalina de Mossa", 0, 0, 626438, 9435060, 0.597433, "Z04"),
    (33, "42", "C1086-Q9570", 80.62, "Morropon", "Santa Catalina de Mossa", 0, 0, 625152, 9435011, 0.634841, "Z04"),
    (34, "14", "C1086-Q9570", 92.9, "Morropon", "Santa Catalina de Mossa", 0, 0, 624397, 9437085, 0.51974, "Z04"),
    (35, "58", "C1086-Q9569", 11.47, "Morropon", "Santa Catalina de Mossa", 0, 0, 619094, 9426406, 0.420325, "Z04"),
    (36, "M19B2", "C1086-Q9570", 74.37, "Morropon", "Morropon", 0, 0, 617715, 9427284, 0.401185, "Z05"),
    (37, "M32B3", "C1086-Q9569", 50.97, "Morropon", "Morropon", 0, 0, 618248, 9426493, 0.385213, "Z05"),
    (38, "M6B10", "C1077-Q9566", 249.89, "Morropon", "Buenos Aires", 0, 0, 618299, 9423119, 0.30989, "Z06"),
    (39, "M6B2-3", "C1077-Q9566", 160.98, "Morropon", "Buenos Aires", 0, 0, 611305, 9419274, 0.406259, "Z06"),
    (40, "M6B2-2", "C1077-Q9566", 333.01, "Morropon", "Buenos Aires", 0, 0, 611196, 9416725, 0.302636, "Z06"),
    (41, "M6B2-1", "C1077-Q9566", 514.36, "Morropon", "Buenos Aires", 0, 0, 612864, 9413345, 0.3381, "Z06"),
    (42, "5", "C1077-Q9566", 295.73, "Morropon", "Buenos Aires", 0, 0, 617860, 9413182, 0.3394, "Z06"),
    (43, "M1B1", "C1077-Q9580", 188.8, "Morropon", "Buenos Aires", 0, 0, 617695, 9409136, 0.319305, "Z06"),
    (44, "M18B1", "C1077-Q9579", 160.36, "Morropon", "Buenos Aires", 0, 0, 621058, 9410427, 0.334299, "Z06"),
    (45, "M3B9", "C1081-Q9582", 294.2, "Morropon", "Salitral", 0, 0, 629338, 9413305, 0.321386, "Z07"),
    (46, "M3B1", "C1081-Q9582", 80.98, "Morropon", "Salitral", 0, 0, 630685, 9410800, 0.323882, "Z07"),
    (47, "M7B1", "C1076-Q9581", 130.43, "Morropon", "Salitral", 0, 0, 629006, 9407536, 0.313647, "Z07"),
    (48, "M7B2", "C1076-Q9581", 115.38, "Morropon", "Salitral", 0, 0, 631438, 9403547, 0.287689, "Z07"),
    (49, "M7B3", "C1076-Q9581", 65.74, "Morropon", "Salitral", 0, 0, 633232, 9401227, 0.298017, "Z07"),
    (50, "M7B6", "C1076-Q9581", 91.29, "Morropon", "Salitral", 0, 0, 635110, 9399092, 0.287833, "Z07"),
    (51, "M2B1", "C1076-Q9584", 81.79, "Morropon", "Salitral", 0, 0, 635876, 9397941, 0.29836, "Z07"),
    (52, "7", "C1076-Q9584", 173.43, "Morropon", "Salitral", 0, 0, 639728, 9396836, 0.375288, "Z07"),
    (53, "M2B5", "C1076-Q9584", 30.91, "Morropon", "Salitral", 0, 0, 639987, 9394392, 0.304733, "Z07"),
    (54, "1", "C1076-Q9584", 1370.96, "Morropon", "Salitral", 0, 0, 645575, 9393500, 0.454493, "Z07"),
    (55, "M18B3", "C1077-Q9579", 373.81, "Morropon", "Salitral", 0, 0, 625311, 9412195, 0.369424, "Z07"),
    (56, "M18B5", "C1077-Q9579", 197.35, "Morropon", "Salitral", 0, 0, 621551, 9412989, 0.356522, "Z07"),
    (57, "M3B7", "C1081-Q9582", 122.95, "Morropon", "San Juan de Bigote", 0, 0, 638309, 9415217, 0.370678, "Z08"),
    (58, "M3B6", "C1081-Q9582", 60.35, "Morropon", "San Juan de Bigote", 0, 0, 637918, 9412134, 0.342535, "Z08"),
    (59, "77", "C1081-Q9583", 17.82, "Morropon", "San Juan de Bigote", 0, 0, 638648, 9411370, 0.272488, "Z08"),
    (60, "M3B5", "C1081-Q9582", 52.74, "Morropon", "San Juan de Bigote", 0, 0, 635887, 9411738, 0.283016, "Z08"),
    (61, "M3B3", "C1081-Q9582", 84.82, "Morropon", "San Juan de Bigote", 0, 0, 633015, 9411483, 0.298483, "Z08"),
    (62, "M3B8", "C1081-Q9582", 565.58, "Morropon", "San Juan de Bigote", 0, 0, 632119, 9415141, 0.293351, "Z08"),
    (63, "M11B3", "C1081-Q9583", 106.03, "Morropon", "San Juan de Bigote", 0, 0, 642079, 9413823, 0.395609, "Z08"),
    (64, "55", "C1086-Q9575", 0.86, "Morropon", "Yamango", 0, 0, 642690, 9438409, 0.630032, "Z09"),
    (65, "49", "C1086-Q9575", 15.84, "Morropon", "Yamango", 0, 0, 642334, 9438263, 0.686036, "Z09"),
    (66, "52", "C1086-Q9575", 11.06, "Morropon", "Yamango", 0, 0, 641522, 9438064, 0.66855, "Z09"),
    (67, "48", "C1086-Q9575", 5.75, "Morropon", "Yamango", 0, 0, 639496, 9433679, 0.697809, "Z09"),
    (68, "46", "C1086-Q9575", 18.87, "Morropon", "Yamango", 0, 0, 638751, 9433236, 0.724382, "Z09"),
    (69, "M28B2", "C1086-Q9575", 90.2, "Morropon", "Yamango", 0, 0, 636249, 9425665, 0.475342, "Z09"),
    (70, "M28B3", "C1086-Q9575", 58.29, "Morropon", "Yamango", 0, 0, 634061, 9426778, 0.507529, "Z09"),
    (71, "M28B4", "C1086-Q9575", 283.43, "Morropon", "Yamango", 0, 0, 626078, 9427110, 0.42751, "Z09"),
    (72, "33", "C1081-Q9590", 29.68, "Huancabamba", "Lalaquiz", 0, 0, 646466, 9431126, 0.71923, "Z10"),
    (73, "38", "C1081-Q9590", 44.2, "Huancabamba", "Lalaquiz", 0, 0, 647123, 9427602, 0.672487, "Z10"),
    (74, "50", "C1081-Q9590", 19.27, "Huancabamba", "Lalaquiz", 0, 0, 647350, 9427252, 0.68307, "Z10"),
    (75, "34", "C1081-Q9590", 28.45, "Huancabamba", "Lalaquiz", 0, 0, 646483, 9426305, 0.648975, "Z10"),
    (76, "61", "C1081-Q9590", 26.83, "Huancabamba", "Lalaquiz", 0, 0, 646236, 9425052, 0.611186, "Z10"),
    (77, "37", "C1081-Q9591", 35.49, "Huancabamba", "Lalaquiz", 0, 0, 649081, 9424506, 0.658846, "Z10"),
    (78, "26", "C1081-Q9591", 46.96, "Huancabamba", "Lalaquiz", 0, 0, 649472, 9425705, 0.673303, "Z10"),
    (79, "64", "C1081-Q9591", 35.77, "Huancabamba", "Huancabamba", 0, 0, 657819, 9428541, 0.525448, "Z11"),
    (80, "54", "C1081-Q9591 (inferida)", 40.26, "Huancabamba", "Huancabamba", 0, 0, 656897, 9426033, 0.600447, "Z11"),
    (81, "M30B5", "C1081-Q9591", 90.9, "Huancabamba", "Huancabamba", 0, 0, 655263, 9427340, 0.5564, "Z11"),
    (82, "60", "C1081-Q9591", 40.13, "Huancabamba", "Canchaque", 0, 0, 653850, 9426236, 0.67447, "Z11"),
    (83, "40", "C1081-Q9591", 28.52, "Huancabamba", "Canchaque", 0, 0, 654137, 9424425, 0.640099, "Z11"),
    (84, "30", "C1081-Q9591", 34.87, "Huancabamba", "Canchaque", 0, 0, 652758, 9424380, 0.663804, "Z11"),
    (85, "28", "C1081-Q9591", 52.97, "Huancabamba", "Canchaque", 0, 0, 654188, 9422849, 0.639425, "Z11"),
    (86, "62", "C1081-Q9591", 73.82, "Huancabamba", "Canchaque", 0, 0, 655046, 9420450, 0.681648, "Z11"),
    (87, "65", "C1081-Q9591 (inferida)", 53.93, "Huancabamba", "Canchaque", 0, 0, 657541, 9421083, 0.649297, "Z11"),
    (88, "21", "C1081-Q9591", 84.23, "Huancabamba", "Canchaque", 0, 0, 654604, 9417736, 0.620078, "Z11"),
    (89, "66", "C1081-Q9591", 102.34, "Huancabamba", "Canchaque", 0, 0, 652566, 9417311, 0.640254, "Z11"),
    (90, "24", "C1081-Q9583", 90.16, "Huancabamba", "Canchaque", 0, 0, 652191, 9414358, 0.674839, "Z11"),
    (91, "63", "C1081-Q9583", 50.46, "Huancabamba", "Canchaque", 0, 0, 655059, 9413845, 0.603582, "Z11"),
    (92, "79", "C1076-Q9587", 94.81, "Huancabamba", "Canchaque", 0, 0, 650305, 9402785, 0.547879, "Z11"),
    (93, "81", "C1076-Q9585", 8.93, "Huancabamba", "Canchaque", 0, 0, 641538, 9401360, 0.274714, "Z11"),
    (94, "M22B1", "C1076-Q9586", 55.17, "Huancabamba", "San Miguel de El Faique", 0, 0, 664082, 9404172, 0.594025, "Z12"),
    (95, "44", "C1076-Q9586", 28.23, "Huancabamba", "San Miguel de El Faique", 0, 0, 663637, 9400618, 0.511324, "Z12"),
    (96, "47", "C1076-Q9586", 13.55, "Huancabamba", "San Miguel de El Faique", 0, 0, 659649, 9401436, 0.593628, "Z12"),
    (97, "75", "C1076-Q9587", 15.18, "Huancabamba", "San Miguel de El Faique", 0, 0, 657722, 9401669, 0.630888, "Z12"),
    (98, "9", "C1076-Q9586", 109.86, "Huancabamba", "San Miguel de El Faique", 0, 0, 656578, 9397762, 0.574527, "Z12"),
    (99, "43", "C1076-Q9586", 23.17, "Huancabamba", "San Miguel de El Faique", 0, 0, 656489, 9395403, 0.551525, "Z12"),
    (100, "35", "C1076-Q9586", 30.8, "Huancabamba", "San Miguel de El Faique", 0, 0, 654607, 9395038, 0.624477, "Z12"),
    (101, "67", "C1076-Q9586", 13.18, "Huancabamba", "San Miguel de El Faique", 0, 0, 654207, 9394210, 0.605293, "Z12"),
    (102, "M20B1", "C1076-Q9585", 279.68, "Huancabamba", "San Miguel de El Faique", 0, 0, 641196, 9398710, 0.355329, "Z12"),
    (103, "M2B8", "C1076-Q9584", 116.1, "Huancabamba", "San Miguel de El Faique", 0, 0, 639059, 9398295, 0.375171, "Z12"),
    (104, "29", "C1076-Q9592", 40.1, "Huancabamba", "Huarmaca", 0, 0, 662123, 9397575, 0.548192, "Z13"),
    (105, "68", "C1076-Q9592", 21.29, "Huancabamba", "Huarmaca", 0, 0, 664225, 9396339, 0.482981, "Z13"),
    (106, "53", "C1076-Q9592", 13.06, "Huancabamba", "Huarmaca", 0, 0, 666537, 9395918, 0.596277, "Z13"),
    (107, "70", "C1076-Q9592", 22.39, "Huancabamba", "Huarmaca", 0, 0, 666781, 9395573, 0.622295, "Z13"),
    (108, "31", "C1076-Q9592", 47.58, "Huancabamba", "Huarmaca", 0, 0, 663334, 9392820, 0.435805, "Z13"),
    (109, "69", "C1076-Q9592", 19.43, "Huancabamba", "Huarmaca", 0, 0, 663291, 9392332, 0.448926, "Z13"),
    (110, "72", "C1076-Q9592", 12.4, "Huancabamba", "Huarmaca", 0, 0, 662236, 9391211, 0.518477, "Z13"),
    (111, "41", "C1076-Q9592", 64.92, "Huancabamba", "Huarmaca", 0, 0, 660525, 9394327, 0.592669, "Z13"),
    (112, "73", "C1076-Q9592", 54.29, "Huancabamba", "Huarmaca", 0, 0, 658547, 9395594, 0.554974, "Z13"),
    (113, "18", "C1076-Q9592", 103.62, "Huancabamba", "Huarmaca", 0, 0, 656445, 9393180, 0.640219, "Z13"),
    (114, "74", "C1076-Q9588", 29.23, "Huancabamba", "Huarmaca", 0, 0, 652449, 9389179, 0.356463, "Z13"),
    (115, "M12B1", "C1076-Q9588", 268.39, "Huancabamba", "Huarmaca", 0, 0, 644546, 9389106, 0.34753, "Z13"),
    (116, "19", "C1076-Q9592", 82.09, "Huancabamba", "Huarmaca", 0, 0, 665942, 9391144, 0.537184, "Z13"),
    (117, "M4B4", "C1076-Q9589", 289.91, "Huancabamba", "Huarmaca", 0, 0, 642775, 9388574, 0.319788, "Z14"),
    (118, "M4B3", "C1076-Q9589", 234.07, "Huancabamba", "Huarmaca", 0, 0, 643474, 9385620, 0.322153, "Z14"),
    (119, "10", "C1076-Q9593", 150.21, "Huancabamba", "Huarmaca", 0, 0, 650830, 9382303, 0.546492, "Z14"),
    (120, "16", "C1076-Q9593", 83.33, "Huancabamba", "Huarmaca", 0, 0, 652939, 9382105, 0.628423, "Z14"),
    (121, "4", "C1076-Q9588", 230.1, "Huancabamba", "Huarmaca", 0, 0, 656549, 9382009, 0.660111, "Z14"),
    (122, "51", "C1076-Q9593", 22.93, "Huancabamba", "Huarmaca", 0, 0, 656291, 9379096, 0.523788, "Z14"),
    (123, "76", "C1076-Q9593", 19.62, "Huancabamba", "Huarmaca", 0, 0, 657922, 9377477, 0.492462, "Z14"),
    (124, "80", "C1076-Q9593", 28.83, "Huancabamba", "Huarmaca", 0, 0, 657160, 9373050, 0.593069, "Z14"),
    (125, "71", "C1076-Q9593", 10.67, "Huancabamba", "Huarmaca", 0, 0, 652104, 9373045, 0.557066, "Z14"),
    (126, "8", "C1076-Q9593", 126.95, "Huancabamba", "Huarmaca", 0, 0, 651673, 9373383, 0.554802, "Z14"),
    (127, "20", "C1076-Q9593", 81.33, "Huancabamba", "Huarmaca", 0, 0, 649847, 9376260, 0.48513, "Z14"),
]

# Alias para compatibilidad con codigo previo
BLOQUES_128 = BLOQUES_V5

BLOQUES_128_MAP = {b[1]: {"n": b[0], "codigo": b[1], "microcuenca": b[2],
    "area_ha": b[3], "provincia": b[4], "distrito": b[5],
    "accesibilidad": b[6], "dia_evaluacion": b[7],
    "utm_este": b[8], "utm_norte": b[9], "msavi_2024": b[10],
    "zona": b[11] if len(b) > 11 else ""} for b in BLOQUES_128}

# Aliases V5
BLOQUES_V5_MAP = BLOQUES_128_MAP

# Lista de codigos para el dropdown (solo codigo de bloque)
BLOQUES_128_OPCIONES = [b[1] for b in BLOQUES_128]
BLOQUES_V5_OPCIONES = BLOQUES_128_OPCIONES

# Lista de Zonas V5 (Z01..Z14)
ZONAS_V5 = sorted({b[11] for b in BLOQUES_V5 if len(b) > 11 and b[11]})

# ── Centros Poblados por Bloque de Intervencion ──────────────────────
# Fuente: REPORTE BLOQUES V4 PARA APLICATIVO.xlsx (sin duplicados, buffer 500m)
CENTROS_POBLADOS_BLOQUE = {
    "1": {"centros_poblados": ["Gramadal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "3": {"centros_poblados": ["Huayabal", "Rincon De Geraldo", "El Mirador", "Los Checches", "Huasipe de Geraldo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "5": {"centros_poblados": ["Juan Velasco", "El Ala", "Hacienda El Muerto", "Lindero del Ala"], "comunidades_campesinas": [], "poblacion_total": 0},
    "6": {"centros_poblados": ["Nogal", "Guanabano / Huanabano", "Huanabano Alto", "Loma Guabal", "Cascajal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "8": {"centros_poblados": ["Pirga"], "comunidades_campesinas": [], "poblacion_total": 0},
    "10": {"centros_poblados": ["La Pena", "Pena Grande"], "comunidades_campesinas": [], "poblacion_total": 0},
    "11": {"centros_poblados": ["Taylin de Tunali", "Tailin"], "comunidades_campesinas": [], "poblacion_total": 0},
    "13": {"centros_poblados": ["La Cruz", "Huacas", "Simiris"], "comunidades_campesinas": [], "poblacion_total": 0},
    "14": {"centros_poblados": ["Las Velas / Las Vegas", "Brasal", "El Murcielago"], "comunidades_campesinas": [], "poblacion_total": 0},
    "15": {"centros_poblados": ["Jacanacas"], "comunidades_campesinas": [], "poblacion_total": 0},
    "16": {"centros_poblados": ["Laguna de Paltama"], "comunidades_campesinas": [], "poblacion_total": 0},
    "17": {"centros_poblados": ["El Checo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "18": {"centros_poblados": ["Piedra Blanca"], "comunidades_campesinas": [], "poblacion_total": 0},
    "19": {"centros_poblados": ["Nuevo Progreso"], "comunidades_campesinas": [], "poblacion_total": 0},
    "20": {"centros_poblados": ["San Juan", "Lanchipuque", "Tupac Amaru"], "comunidades_campesinas": [], "poblacion_total": 0},
    "21": {"centros_poblados": ["Huamala Alto"], "comunidades_campesinas": [], "poblacion_total": 0},
    "22": {"centros_poblados": ["Platanal Bajo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "23": {"centros_poblados": ["Jacanacas", "Nueva Esperanza"], "comunidades_campesinas": [], "poblacion_total": 0},
    "24": {"centros_poblados": ["Villanueva"], "comunidades_campesinas": [], "poblacion_total": 0},
    "26": {"centros_poblados": ["La Laguna"], "comunidades_campesinas": [], "poblacion_total": 0},
    "27": {"centros_poblados": ["Nueva Esperanza de Misquis"], "comunidades_campesinas": [], "poblacion_total": 0},
    "28": {"centros_poblados": ["Abalque"], "comunidades_campesinas": [], "poblacion_total": 0},
    "30": {"centros_poblados": ["La Virgen"], "comunidades_campesinas": [], "poblacion_total": 0},
    "33": {"centros_poblados": ["Sambe"], "comunidades_campesinas": [], "poblacion_total": 0},
    "34": {"centros_poblados": ["Maray"], "comunidades_campesinas": [], "poblacion_total": 0},
    "35": {"centros_poblados": ["Quitahuajara", "Chamelico"], "comunidades_campesinas": [], "poblacion_total": 0},
    "36": {"centros_poblados": ["Cascajal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "37": {"centros_poblados": ["La Oficina", "El Papayo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "38": {"centros_poblados": ["Curana"], "comunidades_campesinas": [], "poblacion_total": 0},
    "39": {"centros_poblados": ["La Cria"], "comunidades_campesinas": [], "poblacion_total": 0},
    "40": {"centros_poblados": ["Chorro Blanco", "Chacchacal / Chorro Blanco"], "comunidades_campesinas": [], "poblacion_total": 0},
    "42": {"centros_poblados": ["Portachuelo", "Santa Rosa de Chirimoyos", "El Ceibo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "43": {"centros_poblados": ["Coyona", "Collona", "Pizarrume"], "comunidades_campesinas": [], "poblacion_total": 0},
    "49": {"centros_poblados": ["Las Huacas", "Pajal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "51": {"centros_poblados": ["Talla"], "comunidades_campesinas": [], "poblacion_total": 0},
    "52": {"centros_poblados": ["Miraflores", "Pajal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "53": {"centros_poblados": ["Zururan"], "comunidades_campesinas": [], "poblacion_total": 0},
    "55": {"centros_poblados": ["Pajal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "61": {"centros_poblados": ["Pedregal", "Alto Zonal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "63": {"centros_poblados": ["Shuturumbe"], "comunidades_campesinas": [], "poblacion_total": 0},
    "66": {"centros_poblados": ["Huamala", "Naranjo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "67": {"centros_poblados": ["El Palto"], "comunidades_campesinas": [], "poblacion_total": 0},
    "68": {"centros_poblados": ["Tabluran"], "comunidades_campesinas": [], "poblacion_total": 0},
    "70": {"centros_poblados": ["Zururan"], "comunidades_campesinas": [], "poblacion_total": 0},
    "71": {"centros_poblados": ["Pirga"], "comunidades_campesinas": [], "poblacion_total": 0},
    "72": {"centros_poblados": ["Ramon Castilla"], "comunidades_campesinas": [], "poblacion_total": 0},
    "74": {"centros_poblados": ["Chalpa"], "comunidades_campesinas": [], "poblacion_total": 0},
    "76": {"centros_poblados": ["Hualanga", "Sahuate"], "comunidades_campesinas": [], "poblacion_total": 0},
    "78": {"centros_poblados": ["Tasajeras"], "comunidades_campesinas": [], "poblacion_total": 0},
    "80": {"centros_poblados": ["Cruz Roja"], "comunidades_campesinas": [], "poblacion_total": 0},
    "81": {"centros_poblados": ["Hualtacal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "82": {"centros_poblados": ["Portachuelo/Portachuelo de San Francisco"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M10B4": {"centros_poblados": ["Rio Seco Alto"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M12B1": {"centros_poblados": ["Nueva Esperanza de Los Molinos", "Hualcas", "Gramadal"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M17B1": {"centros_poblados": ["Papelillo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M17B4": {"centros_poblados": ["Chililique Alto"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M17B5": {"centros_poblados": ["El Guabo"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M17B6": {"centros_poblados": ["Chililique Alto", "Platanal Alto"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M17B7": {"centros_poblados": ["Pampa de Ramada"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M18B3": {"centros_poblados": ["Huaroquispampa", "Selva Andina", "Mangamanguilla"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M19B2": {"centros_poblados": ["Boca Negra"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M1B1": {"centros_poblados": ["Rio Seco"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M20B1": {"centros_poblados": ["Las Huacas", "Mantequillera"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M22B1": {"centros_poblados": ["Santa Rosa La Antena"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M27B1": {"centros_poblados": ["Vicus la Merced", "Huasimal", "Linderos de Vicus"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M28B2": {"centros_poblados": ["Flor de Agua", "Victor Raul Haya de La Torre"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M28B3": {"centros_poblados": ["Ricardo Palma"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M28B4": {"centros_poblados": ["Alto Mambluque", "Mambluque", "Nueva Esperanza"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M2B1": {"centros_poblados": ["Santa Rosa"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M2B5": {"centros_poblados": ["Hornopampa"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M2B8": {"centros_poblados": ["Las Huacas"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M32B3": {"centros_poblados": ["Maray"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M36B2": {"centros_poblados": ["Portachuelo", "Santa Rosa de Chirimoyos", "Laguna Colorada"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M3B1": {"centros_poblados": ["Piedra Blanca"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M3B3": {"centros_poblados": ["Alan Garcia", "Bigote", "Bigote de Gato"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M3B5": {"centros_poblados": ["San Juan Bautista"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M3B6": {"centros_poblados": ["San Juan Bautista", "Manzanares", "Bado De Garzas"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M3B7": {"centros_poblados": ["Polluco", "Sinai"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M3B8": {"centros_poblados": ["San Pedro", "Santa Rosa"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M4B3": {"centros_poblados": ["La Huaca"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M4B4": {"centros_poblados": ["Hualcas Alto", "Hualcas"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M6B10": {"centros_poblados": ["La Maravilla", "La Pilca"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M7B1": {"centros_poblados": ["La Alberca", "Nuevo Progreso", "Victor Raul"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M7B2": {"centros_poblados": ["El Cerezo", "La Alberca"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M7B3": {"centros_poblados": ["La Tranca", "Palo Blanco"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M7B6": {"centros_poblados": ["Santa Rosa", "Serran"], "comunidades_campesinas": [], "poblacion_total": 0},
    "M8B2": {"centros_poblados": ["Hualtacal"], "comunidades_campesinas": [], "poblacion_total": 0},
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

# ── Fichas de Diagnostico Territorial V5 (F-DT-01 a F-DT-05) ────────────
# Fuente: Plantilla_DT_Campo_Check_Validada_V5.xlsx
FICHAS_DT = ["F-DT-01","F-DT-02","F-DT-03","F-DT-04","F-DT-05"]

# Listas de opciones (dropdowns) — extraidas de la hoja '_Listas' del V5
FDT_FORMA_TERRENO = ["Plano", "Ondulado", "Colinoso", "Montañoso", "Muy escarpado"]
FDT_RANGO_PENDIENTE = [
    "0-8% (Plano-lig. inclinado)", "8-15% (Mod. inclinado)",
    "15-25% (Fuert. inclinado)", "25-50% (Mod. escarpado)",
    "50-75% (Escarpado)", ">75% (Muy escarpado)",
]
FDT_POSICION_FISIO = [
    "Cresta / Divortium", "Ladera alta", "Ladera media", "Ladera baja",
    "Pie de ladera", "Terraza aluvial", "Cauce / Ribera",
    "Llanura / Pampa", "Cima (sin exp.)",
]
FDT_EXPOSICION = [
    "Norte", "Sur", "Este", "Oeste", "Noreste",
    "Noroeste", "Sureste", "Suroeste", "Cima (sin dom.)",
]
FDT_RANGO_ALTITUD = [
    "<1000 m (Yunga)", "1000-1500 m (Quechua baja)",
    "1500-2000 m (Quechua media)", "2000-2500 m (Quechua alta)",
    "2500-3000 m (Suni)", "3000-3500 m (Puna baja)",
    ">3500 m (Puna/Jalca)",
]
FDT_NIVEL_EROSION = [
    "Nula", "Ligera (laminar)", "Moderada (surcos)",
    "Fuerte (cárcavas incipientes)", "Severa (cárcavas activas)", "Extrema",
]
FDT_TIPO_CARCAVA = [
    "Surco (<30 cm)", "Cárcava incipiente (30-50 cm)",
    "Cárcava somera (0.5-2 m)", "Cárcava profunda (2-5 m)",
    "Cárcava muy profunda (>5 m)",
]
FDT_ESTADO_CARCAVA = [
    "Activa", "Semi-activa", "Estabilizada con vegetación",
    "Estabilizada con obras", "Inactiva",
]
FDT_CAUSA_CARCAVA = [
    "Sobrepastoreo", "Tala / pérdida de cobertura", "Quema",
    "Surcos de labranza en pendiente", "Caminos / huellas de ganado",
    "Concentración natural de escorrentía", "Socavamiento de cauce",
    "Cambio de uso del suelo", "Otro",
]
FDT_PATRON_CARCAVAS = ["Dendrítico", "Paralelo", "Lineal", "Aislado", "Mixto"]
FDT_URGENCIA = ["Crítica", "Alta", "Media", "Baja"]
FDT_TIPO_ECOSISTEMA = [
    "Páramo", "Bosque Relicto Montano V. Occidental (BMVO)",
    "Bosque Estacionalmente Seco Colina/Montaña (Bes-cm)",
    "Bosque Estacionalmente Seco Llanura (Bes-ll)",
    "Bosque Estacionalmente Seco Ribereño (Bes-rb)",
    "Matorral Andino árido", "Matorral Andino semiárido",
    "Matorral Andino subhúmedo", "Matorral Andino húmedo",
    "Pastizal andino / Jalca", "Agroecosistema / Mosaico",
    "Área urbana / rural", "Otro",
]
FDT_ESTADO_CONSERVACION = [
    "Conservado (sin intervención evidente)", "Levemente alterado",
    "Medianamente alterado", "Alterado (intervención marcada)",
    "Muy alterado / Degradado", "En restauración / Recuperación",
]
FDT_USO_SUELO_DOM = [
    "Forestal protección", "Forestal producción", "Pastoreo extensivo",
    "Pastoreo intensivo", "Agricultura secano", "Agricultura bajo riego",
    "Sin uso aparente / Abandonado", "Uso mixto", "Otro",
]
FDT_REGENERACION = [
    "Abundante (>50 pl./100m²)", "Regular (10-50 pl./100m²)",
    "Escasa (<10 pl./100m²)", "Ausente",
]
FDT_ESTADO_SANITARIO = [
    "Sano", "Enfermo (plaga)", "Enfermo (hongo)", "Daño mecánico",
    "Muerto en pie", "Rebrote vigoroso",
]
FDT_FENOLOGIA = [
    "Vegetativo", "Floración", "Fructificación",
    "Caducifolio (hojas caídas)", "Latencia / Estrés hídrico",
]
FDT_TIPO_COBERTURA = [
    "Bosque denso (>70%)", "Bosque ralo (30-70%)", "Matorral denso",
    "Matorral ralo", "Pastizal natural", "Pastizal cultivado",
    "Cultivo anual", "Cultivo permanente", "Suelo desnudo",
    "Afloramiento rocoso", "Mixto / Mosaico",
]
FDT_ESTRATO = [
    "Arbóreo superior (>15 m)", "Arbóreo medio (8-15 m)",
    "Arbóreo bajo (4-8 m)", "Arbustivo alto (1.5-4 m)",
    "Arbustivo bajo (0.5-1.5 m)", "Herbáceo (<0.5 m)",
    "Epífito", "Trepador",
]
FDT_ORIGEN = ["Nativa", "Endémica", "Introducida", "Invasora"]
FDT_ABUNDANCIA = [
    "Dominante (>40%)", "Abundante (20-40%)", "Frecuente (5-20%)",
    "Ocasional (1-5%)", "Rara (<1%)",
]
FDT_CATEGORIA_INDICADORA = [
    "Indicadora de buen estado", "Indicadora de perturbación",
    "Endémica del Perú", "Endémica norte / Piura",
    "Amenazada (D.S. 043-2006-AG)", "Clave cultural",
]
FDT_ESTADO_UICN = [
    "EX — Extinto", "EW — Extinto en Estado Silvestre",
    "CR — En Peligro Crítico", "EN — En Peligro", "VU — Vulnerable",
    "NT — Casi Amenazado", "LC — Preocupación Menor",
    "DD — Datos Insuficientes", "NE — No Evaluado", "No aplica",
]
FDT_INTENSIDAD = ["Nula", "Ligera", "Moderada", "Fuerte", "Muy fuerte"]
FDT_NIVEL_IND = ["Alto", "Medio", "Bajo"]
FDT_VELOCIDAD = ["Muy rápida", "Rápida", "Moderada", "Lenta", "Estable"]
FDT_REVERSIBILIDAD = [
    "Totalmente reversible", "Parcialmente reversible",
    "Difícilmente reversible", "En recuperación",
]
FDT_TIPO_FUENTE = [
    "Manantial / Puquio", "Quebrada permanente", "Quebrada estacional",
    "Río", "Canal de riego", "Reservorio / Represa", "Laguna / Humedal",
    "Pozo / Galería filtrante", "Bofedal", "Otro",
]
FDT_REGIMEN_HIDRICO = [
    "Permanente", "Estacional (lluvias)", "Intermitente",
    "Efímero", "Ausente",
]
FDT_CALIDAD_AGUA = [
    "Buena (clara, sin olor, sin sedimentos)",
    "Regular (ligeramente turbia o con sedimentos)",
    "Mala (turbia, con olor o color)",
    "Muy mala (contaminación evidente)", "No evaluable",
]
FDT_MODALIDAD_ACCESO = [
    "Vehicular 4x4 hasta el bloque", "Vehicular + caminata <30 min",
    "Vehicular + caminata 30-90 min", "Vehicular + caminata >90 min",
    "Requiere acémila (páramo/jalca)", "Inaccesible en lluvias",
]
FDT_TIPO_VIA = [
    "Nacional asfaltada", "Nacional afirmada", "Departamental asfaltada",
    "Departamental afirmada", "Vecinal afirmada",
    "Trocha carrozable", "Camino de herradura",
]
FDT_NIVEL_TRANSITAB = ["Alto", "Medio", "Bajo"]
FDT_SENAL_CELULAR = ["Completa", "Parcial", "Intermitente", "Sin señal"]
FDT_OPERADOR = ["Movistar", "Claro", "Entel", "Bitel", "Otro", "Sin señal"]
FDT_SI_NO = ["Sí", "No"]
FDT_SI_NO_NA = ["Sí", "No", "No aplica"]

# Causas predefinidas para la matriz de F-DT-04 (16 filas fijas)
FDT04_CAUSAS_LABELS = [
    "Sobrepastoreo (caprino/vacuno)", "Tala para leña", "Tala para carbón",
    "Tala para madera", "Quema (renovación de pastos)",
    "Quema accidental / Incendio", "Cambio de uso a agricultura",
    "Cambio de uso a pastoreo", "Expansión urbana / vial",
    "Minería informal", "Extracción de áridos (cauces)",
    "Especies invasoras", "Erosión natural severa",
    "Sequía prolongada / CC", "Plagas o enfermedades forestales", "Otro",
]
# Indicadores cuantitativos predefinidos para la matriz de F-DT-04 (8 filas)
FDT04_INDICADORES_LABELS = [
    ("Cobertura vegetal total", "%", ">60% bueno; 30-60% medio; <30% malo"),
    ("Porcentaje de suelo desnudo", "%", "<10% bueno; 10-30% medio; >30% malo"),
    ("Densidad de cárcavas", "N°/ha", "0 nulo; 1-3 medio; >3 alto"),
    ("Presencia de especies invasoras", "% área", "<5% bueno; 5-20% medio; >20% alto"),
    ("Carga ganadera estimada", "UA/ha", "Según tipo de ecosistema"),
    ("Frecuencia de quemas últimos 5 años", "N°", "0 nulo; 1-2 medio; >2 alto"),
    ("% área con pendiente >50%", "%", "Factor de susceptibilidad"),
    ("Distancia a vía vecinal más cercana", "m", "A menor distancia → mayor presión"),
]

# ── Fichas de Diagnostico Social V3 (F-DS-01 a F-DS-07) ──────────────────
# Fuente unica de verdad: fds_listas.py (generado desde la hoja `_Listas` de
# la Plantilla de Diagnostico Social validada V3). Las 7 fichas reemplazan a
# los 5 formatos anteriores y usan celdas de validacion (desplegables) que
# replican las listas oficiales del Excel.
import fds_listas as FL

FICHAS_DS = ["F-DS-01", "F-DS-02", "F-DS-03", "F-DS-04",
             "F-DS-05", "F-DS-06", "F-DS-07"]

FICHAS_DS_TITULOS = {
    "F-DS-01": "Diagnostico Socioeconomico del CP / Comunidad",
    "F-DS-02": "Mapeo y Analisis de Actores Clave",
    "F-DS-03": "Entrevista Semiestructurada a Autoridades / Lideres",
    "F-DS-04": "Acta de Taller Participativo",
    "F-DS-05": "Conflictos Socioambientales y Oportunidades",
    "F-DS-06": "Percepcion Local de Peligros y Cambio Climatico",
    "F-DS-07": "Disposicion a Participar y Consentimiento Previo Informado",
}

# Listas Si/No reutilizadas en multiples fichas
DS_SINO = FL.L_SINO                       # ["Si", "No"]
DS_SINONA = FL.L_SINONA                   # ["Si", "No", "No aplica"]

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
    return {b["codigo"]: b["id"]
            for b in _cached_obtener_bloques(_cache_version())}

def _distritos(prov):
    return PROVINCIAS_DISTRITOS.get(prov, DISTRITOS_PIURA) if prov else DISTRITOS_PIURA

def _resolver_microcuenca(bloque_label):
    """Resuelve la microcuenca para un bloque dado su label 'CODIGO - TIPO'.
    Busca primero en la BD (cacheada) y luego en BLOQUES_128_MAP como fallback."""
    codigo = bloque_label.split(" - ")[0].strip() if " - " in bloque_label else bloque_label.strip()
    # Buscar en BD (cacheada)
    for b in _cached_obtener_bloques(_cache_version()):
        if b["codigo"] == codigo:
            mc = b.get("microcuenca", "") or ""
            if mc and mc in MICROCUENCAS:
                return mc
            break
    # Fallback: buscar en los 128 bloques predefinidos
    datos = BLOQUES_128_MAP.get(codigo, {})
    mc = datos.get("microcuenca", "")
    if mc and mc in MICROCUENCAS:
        return mc
    return ""


def _resolver_datos_bloque(bloque_label):
    """Resuelve metadatos completos del bloque (microcuenca, provincia, distrito,
    superficie ha y coordenadas UTM del centroide).

    Combina datos de la BD (si existen) con BLOQUES_128_MAP como respaldo.
    Retorna dict con claves: microcuenca, provincia, distrito, area_ha,
    utm_este, utm_norte (cualquier valor faltante queda como cadena vacía o 0).
    """
    codigo = bloque_label.split(" - ")[0].strip() if " - " in bloque_label else bloque_label.strip()
    info = {"microcuenca": "", "provincia": "", "distrito": "",
            "area_ha": 0.0, "utm_este": 0.0, "utm_norte": 0.0}
    # 1) Buscar en BD
    for b in _cached_obtener_bloques(_cache_version()):
        if b["codigo"] == codigo:
            info["microcuenca"] = (b.get("microcuenca") or "")
            info["provincia"] = (b.get("provincia") or "")
            info["distrito"] = (b.get("distrito") or "")
            try: info["area_ha"] = float(b.get("area_ha") or 0)
            except (TypeError, ValueError): info["area_ha"] = 0.0
            try: info["utm_este"] = float(b.get("utm_este") or 0)
            except (TypeError, ValueError): info["utm_este"] = 0.0
            try: info["utm_norte"] = float(b.get("utm_norte") or 0)
            except (TypeError, ValueError): info["utm_norte"] = 0.0
            break
    # 2) Completar campos faltantes con catalogo predefinido
    fallback = BLOQUES_128_MAP.get(codigo, {})
    if fallback:
        if not info["microcuenca"]:
            info["microcuenca"] = fallback.get("microcuenca", "")
        if not info["provincia"]:
            info["provincia"] = fallback.get("provincia", "")
        if not info["distrito"]:
            info["distrito"] = fallback.get("distrito", "")
        if not info["area_ha"]:
            try: info["area_ha"] = float(fallback.get("area_ha") or 0)
            except (TypeError, ValueError): pass
        if not info["utm_este"]:
            try: info["utm_este"] = float(fallback.get("utm_este") or 0)
            except (TypeError, ValueError): pass
        if not info["utm_norte"]:
            try: info["utm_norte"] = float(fallback.get("utm_norte") or 0)
            except (TypeError, ValueError): pass
    return info

# ══════════════════════════════════════════════════════════════════════════
# PANEL DE CONTROL
# ══════════════════════════════════════════════════════════════════════════
def pagina_dashboard():
    st.subheader("Panel de Control - Resumen Ejecutivo")
    stats = _cached_obtener_estadisticas(_cache_version())
    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.metric("Total Bloques", stats["total_bloques"])
    c2.metric("Area Total", f"{stats['area_total_ha']:.2f} ha")
    c3.metric("Inspecciones", stats["total_inspecciones"])
    c4.metric("Avance Promedio", f"{stats['avance_promedio']:.1f}%")
    c5.metric("Diag. Territorial", stats.get("total_diagnosticos", 0))
    c6.metric("Diag. Social", stats.get("total_diagnosticos_sociales", 0))
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
        def _fmt_area(v):
            try:
                return f"{float(v):.4f}" if v not in (None, "", 0, 0.0) else ""
            except (TypeError, ValueError):
                return ""
        def _fmt_avance(b):
            v = b.get("ultimo_avance")
            if v in (None, "", 0, 0.0) and not b.get("total_inspecciones"):
                return ""
            try:
                return f"{float(v or 0):.1f}"
            except (TypeError, ValueError):
                return ""
        def _fmt_int(v):
            try:
                iv = int(v)
                return iv if iv else ""
            except (TypeError, ValueError):
                return ""
        st.dataframe(pd.DataFrame([{
            "Codigo": b["codigo"],
            "Tipo": b.get("tipo_intervencion", "") or "",
            "Distrito": b.get("distrito", "") or "",
            "Area (ha)": _fmt_area(b.get("area_hectareas")),
            "Estado": b.get("estado", "") or "",
            "Avance %": _fmt_avance(b),
            "Inspecciones": _fmt_int(b.get("total_inspecciones")),
        } for b in res]),
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

        # Selector rapido de los bloques V5 (solo en modo nuevo)
        if not edit_id:
            st.markdown(f"##### Seleccion rapida - {len(BLOQUES_V5)} Bloques de Intervencion V5")
            sel_79 = st.selectbox(
                "Seleccionar bloque predefinido",
                ["(Seleccionar bloque predefinido)"] + BLOQUES_V5_OPCIONES,
                key="sel_bloque_128",
                help=f"Seleccione un bloque de la lista de {len(BLOQUES_V5)} bloques V5 para autocompletar los campos"
            )

            # Determinar valores por defecto segun seleccion
            cod_sel = _extraer_codigo_bloque_128(sel_79 if sel_79 != "(Seleccionar bloque predefinido)" else "")
            datos_79 = BLOQUES_128_MAP.get(cod_sel, {}) if cod_sel else {}

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
                row_cols[1].write(b.get("codigo", "") or "")
                row_cols[2].write(b.get("microcuenca", "") or "")
                row_cols[3].write(b.get("tipo_intervencion", "") or "")
                row_cols[4].write(b.get("provincia", "") or "")
                row_cols[5].write(b.get("distrito", "") or "")
                area_val = b.get("area_hectareas")
                try:
                    area_txt = f"{float(area_val):.4f}" if area_val not in (None, "", 0, 0.0) else ""
                except (TypeError, ValueError):
                    area_txt = ""
                row_cols[6].write(area_txt)
                row_cols[7].write(b.get("estado", "Pendiente") or "Pendiente")
                if row_cols[8].button("Editar", key=f"edit_bl_{b['id']}", type="primary"):
                    _bl_load_edit(b)
                    st.rerun()
            st.markdown("---")
            bm = {b["codigo"]: b["id"] for b in bloques}
            sel = st.selectbox("Seleccionar bloque para eliminar",[""]+list(bm.keys()),key="del_bl")
            if sel and sel in bm and st.button("Eliminar bloque", key="btn_del_bl"):
                try:
                    db.eliminar_bloque(bm[sel])
                    _invalidar_cache()
                    st.success(f"Bloque {sel} eliminado correctamente.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al eliminar bloque: {e}")

        st.markdown("---")
        with st.expander(f"Tabla de Referencia - {len(BLOQUES_V5)} Bloques de Intervencion V5", expanded=False):
            st.caption("Fuente: Plantilla_DT_Campo_Check_Validada_V5.xlsx - Base de datos completa del proyecto IN Piura")
            df_v5 = pd.DataFrame([{
                "N":b[0], "Bloque":b[1], "Microcuenca":b[2],
                "Area (ha)":b[3], "Provincia":b[4], "Distrito":b[5],
                "Zona": b[11] if len(b) > 11 else "",
                "UTM Este": b[8], "UTM Norte": b[9],
                "MSAVI 2024":b[10],
            } for b in BLOQUES_V5])
            st.dataframe(df_v5, use_container_width=True, hide_index=True, height=400)

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
    _now = datetime.now()
    def_fecha = _now
    def_hora = _now.time().replace(second=0, microsecond=0)
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
            def_clima_idx = CONDICIONES_CLIMATICAS.index(def_clima) + 1
        def_avance = float(st.session_state.get("insp_e_avance", 0))
        def_obs = st.session_state.get("insp_e_obs", "")
        def_desv = st.session_state.get("insp_e_desv", "")
        def_ver = st.session_state.get("insp_e_ver", def_ver)
        _fecha_e = (st.session_state.get("insp_e_fecha", "") or "").strip()
        for _fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
            try:
                _parsed = datetime.strptime(_fecha_e, _fmt)
                def_fecha = _parsed
                if _fmt != "%Y-%m-%d":
                    def_hora = _parsed.time().replace(second=0, microsecond=0)
                break
            except (ValueError, TypeError):
                continue

    with st.form("form_insp", clear_on_submit=not edit_id):
        mc = st.selectbox("Microcuenca", [""] + MICROCUENCAS, index=mc_idx)
        c_fh1, c_fh2 = st.columns(2)
        fecha = c_fh1.date_input("Fecha de visita", value=def_fecha,
                              min_value=FECHA_MIN_PROYECTO, max_value=date.today(),
                              help="Seleccione la fecha de la visita de campo")
        hora = c_fh2.time_input("Hora de visita", value=def_hora,
                              help="Seleccione la hora de la visita de campo")
        inspector = st.text_input("Inspector", value=def_inspector)
        clima = st.selectbox("Condiciones climaticas", [""] + CONDICIONES_CLIMATICAS, index=def_clima_idx)
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
                fecha_hora_str = datetime.combine(fecha, hora).strftime("%Y-%m-%d %H:%M:%S")
                fecha_str = fecha.strftime("%Y-%m-%d")
                if edit_id:
                    db.actualizar_inspeccion(inspeccion_id=edit_id,
                        fecha_visita=fecha_hora_str,
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
                    # Verificar duplicados antes de insertar (compara solo la fecha)
                    existentes = db.obtener_inspecciones_por_bloque(bm[bl])
                    dup = [e for e in existentes
                           if (e["fecha_visita"] or "")[:10] == fecha_str
                           and e["inspector"] == inspector]
                    if dup:
                        st.warning(f"Ya existe una inspeccion para este bloque en {fecha_str} "
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

                        db.insertar_inspeccion(bloque_id=bm[bl],fecha_visita=fecha_hora_str,
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
    st.caption("Haga clic en **Editar** para modificar una inspeccion existente o en **Eliminar** para descartar una version.")
    insp = _cached_obtener_todas_inspecciones(_cache_version())
    if insp:
        # Paginacion
        insp_pag, total_pags_i, pag_actual_i = _paginar(insp, "pag_insp")
        _controles_paginacion(total_pags_i, pag_actual_i, "pag_insp")
        # Tabla con botones de edicion / eliminacion
        col_widths = [0.4, 0.9, 0.9, 0.9, 0.6, 0.85, 0.55, 0.9, 0.45, 0.6, 0.7]
        header_cols = st.columns(col_widths)
        for col, h in zip(header_cols, ["ID", "Bloque", "Microcuenca", "Fecha", "Hora", "Inspector", "Avance%", "Verificacion", "PDFs", "", ""]):
            col.markdown(f"**{h}**")
        st.markdown("---")
        confirm_del_id = st.session_state.get("insp_confirm_del_id")
        for i in insp_pag:
            row = st.columns(col_widths)
            row[0].write(i["id"])
            row[1].write(i["bloque_codigo"])
            row[2].write(i.get("microcuenca", "") or "")
            fv = (i.get("fecha_visita") or "").strip()
            fecha_part = fv[:10] if len(fv) >= 10 else fv
            hora_part = fv[11:16] if len(fv) >= 16 else "—"
            row[3].write(fecha_part)
            row[4].write(hora_part)
            row[5].write(i["inspector"])
            row[6].write(f"{i['avance_fisico']:.1f}")
            row[7].write(i["codigo_verificacion"])
            row[8].write(len([p for p in (i.get("archivos_pdf", "") or "").split(";") if p.strip()]))
            if row[9].button("Editar", key=f"edit_insp_{i['id']}", type="primary"):
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
                st.session_state.pop("insp_confirm_del_id", None)
                st.rerun()
            if confirm_del_id == i["id"]:
                if row[10].button("Confirmar", key=f"del_confirm_insp_{i['id']}", type="primary"):
                    db.eliminar_inspeccion(i["id"])
                    _invalidar_cache()
                    st.session_state.pop("insp_confirm_del_id", None)
                    st.success(f"Inspeccion ID {i['id']} eliminada.")
                    st.rerun()
            else:
                if row[10].button("Eliminar", key=f"del_insp_{i['id']}", type="secondary"):
                    st.session_state["insp_confirm_del_id"] = i["id"]
                    st.rerun()
        if confirm_del_id is not None:
            st.warning(f"Confirme la eliminacion de la inspeccion ID {confirm_del_id} pulsando 'Confirmar' en su fila. "
                       "Esta accion no puede deshacerse.")
            if st.button("Cancelar eliminacion", key="insp_cancel_del"):
                st.session_state.pop("insp_confirm_del_id", None)
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
# DIAGNOSTICO TERRITORIAL V5 (F-DT-01 a F-DT-05)
# ══════════════════════════════════════════════════════════════════════════
def _dt_load_json(raw, default):
    """Decodifica un JSON guardado o devuelve `default` si esta vacio o malformado."""
    import json as _json
    if not raw:
        return default
    try:
        v = _json.loads(raw)
        return v if isinstance(v, list) else default
    except Exception:
        return default


def _dt_dump_json(rows):
    """Serializa una lista de dicts (filas) a JSON, descartando filas totalmente vacias."""
    import json as _json
    limpios = []
    for r in rows:
        if any((str(v).strip() if v is not None else "") for v in r.values()):
            limpios.append({k: ("" if v is None else str(v)) for k, v in r.items()})
    return _json.dumps(limpios, ensure_ascii=False)


# Claves de los widgets escalares de la cabecera del formulario F-DT.
_DT_HEADER_INPUT_KEYS = (
    "dt_fecha", "dt_hora", "dt_corr", "dt_eval", "dt_brigada", "dt_mc",
    "dt_utm_e", "dt_utm_n", "dt_altitud", "dt_cp", "dt_cc", "dt_obs",
)


def _dt_limpiar_widgets_edicion():
    """Elimina del session_state los valores de los widgets del formulario
    F-DT para que, al entrar en modo edicion, los parametros value=/index=
    (que leen de dt_edit_data) vuelvan a tomar efecto.

    Streamlit ignora value=/index= cuando la key del widget ya existe en
    session_state; sin esta limpieza la precarga del registro a editar no se
    refleja en los campos.
    """
    for k in list(st.session_state.keys()):
        if k.startswith(("f01_", "f02_", "f03_", "f04_", "f05_")) or \
           k in _DT_HEADER_INPUT_KEYS:
            del st.session_state[k]


def _dt_dataeditor(label, columns, num_rows, key, options=None, edit_data=None):
    """Renderiza un st.data_editor con `num_rows` filas y columnas dadas."""
    import pandas as _pd
    options = options or {}
    if edit_data:
        df = _pd.DataFrame(edit_data)
        for k, _ in columns:
            if k not in df.columns:
                df[k] = ""
        while len(df) < num_rows:
            df = _pd.concat([df, _pd.DataFrame([{k: "" for k, _ in columns}])],
                            ignore_index=True)
        df = df[[k for k, _ in columns]].head(num_rows).fillna("")
    else:
        df = _pd.DataFrame([{k: "" for k, _ in columns} for _ in range(num_rows)])

    col_config = {}
    for k, lbl in columns:
        if k in options:
            col_config[k] = st.column_config.SelectboxColumn(
                lbl, options=[""] + list(options[k]))
        else:
            col_config[k] = st.column_config.TextColumn(lbl)
    st.markdown(f"**{label}**")
    edited = st.data_editor(
        df, key=key, hide_index=True, num_rows="fixed",
        use_container_width=True, column_config=col_config,
    )
    return edited.to_dict(orient="records")


def pagina_diagnostico_territorial():
    import json as _json
    st.subheader("Diagnostico Territorial - Fichas de Evaluacion")
    st.caption("Fichas F-DT-01 a F-DT-05 (Plantilla V5): Parametros de evaluacion en campo "
               "para el diagnostico del territorio del Proyecto IN Piura.")
    _mostrar_flash()
    bm = _bloques_map()
    if not bm:
        st.warning("Registre un bloque primero.")
        return

    if "dt_edit_id" not in st.session_state:
        st.session_state["dt_edit_id"] = None
    if "dt_edit_data" not in st.session_state:
        st.session_state["dt_edit_data"] = None

    _dt_pending = st.session_state.pop("_dt_autocompletar_pending", None)
    if _dt_pending:
        for _k, _v in _dt_pending.items():
            st.session_state[_k] = _v
        st.session_state["dt_edit_id"] = None
        st.session_state["dt_edit_data"] = None

    dt_edit_id = st.session_state.get("dt_edit_id")
    dt_edit = st.session_state.get("dt_edit_data") or {}

    tab_reg, tab_hist, tab_excel = st.tabs([
        "Registro de Diagnostico", "Historial / Consulta", "Importar desde Excel",
    ])

    with tab_reg:
        if dt_edit_id:
            st.markdown('<div class="edit-mode-banner"><span class="icon">&#9998;</span> '
                        f'Modo Edicion - Diagnostico Territorial ID {dt_edit_id}</div>',
                        unsafe_allow_html=True)
            if st.button("Cancelar edicion", key="dt_cancel_edit"):
                st.session_state["dt_edit_id"] = None
                st.session_state["dt_edit_data"] = None
                st.rerun()

        bl = st.selectbox("Bloque de Intervencion", list(bm.keys()), key="dt_bl")
        bid = bm[bl]

        # ── Auto-vinculacion: microcuenca, provincia, distrito, superficie, UTM ──
        info_bl = _resolver_datos_bloque(bl)
        mc_auto_dt = info_bl.get("microcuenca") or ""
        prov_auto_dt = info_bl.get("provincia") or ""
        dist_auto_dt = info_bl.get("distrito") or ""
        area_auto_dt = info_bl.get("area_ha") or 0.0
        utm_e_auto = info_bl.get("utm_este") or 0.0
        utm_n_auto = info_bl.get("utm_norte") or 0.0
        zona_auto_dt = info_bl.get("zona") or ""

        if mc_auto_dt and mc_auto_dt in MICROCUENCAS:
            mc_idx_dt = MICROCUENCAS.index(mc_auto_dt) + 1
        else:
            mc_idx_dt = 0
        if dt_edit_id:
            mc_val = dt_edit.get("microcuenca", "")
            if mc_val and mc_val in MICROCUENCAS:
                mc_idx_dt = MICROCUENCAS.index(mc_val) + 1

        if any([mc_auto_dt, prov_auto_dt, dist_auto_dt, area_auto_dt, utm_e_auto, utm_n_auto]):
            st.info(
                f"**Datos vinculados automáticamente al bloque {bl}:**  \n"
                f"• Microcuenca: **{mc_auto_dt or '-'}** | Zona: **{zona_auto_dt or '-'}**  \n"
                f"• Provincia: **{prov_auto_dt or '-'}** | Distrito: **{dist_auto_dt or '-'}**  \n"
                f"• Superficie: **{area_auto_dt:.3f} ha**  \n"
                f"• Centroide UTM (Zona 17S, WGS84): **{int(utm_e_auto):,} E / "
                f"{int(utm_n_auto):,} N**"
            )

        # ── Datos generales V5 (compartidos por todas las fichas) ────────
        st.markdown("### 1. Datos Generales (compartidos)")

        def_fecha_dt = datetime.now()
        def_evaluador = ""
        def_brigada = ""
        def_correlativo = ""
        def_hora = datetime.now().strftime("%H:%M")
        def_altitud = ""
        def_cp = ""
        def_cc = ""
        def_utm_e = str(int(utm_e_auto)) if utm_e_auto else ""
        def_utm_n = str(int(utm_n_auto)) if utm_n_auto else ""
        if dt_edit_id:
            def_evaluador = dt_edit.get("evaluador", "") or ""
            def_brigada = dt_edit.get("brigada", "") or ""
            def_correlativo = dt_edit.get("ficha_correlativo", "") or ""
            def_hora = dt_edit.get("hora_registro", "") or def_hora
            def_altitud = dt_edit.get("altitud_gps", "") or ""
            def_cp = dt_edit.get("centro_poblado_cercano", "") or ""
            def_cc = dt_edit.get("comunidad_campesina_dt", "") or ""
            def_utm_e = dt_edit.get("utm_este_dt", "") or def_utm_e
            def_utm_n = dt_edit.get("utm_norte_dt", "") or def_utm_n
            try:
                def_fecha_dt = datetime.strptime(dt_edit.get("fecha_evaluacion", ""), "%Y-%m-%d")
            except (ValueError, TypeError):
                pass

        rA, rB, rC = st.columns(3)
        fecha_ev = rA.date_input(
            "Fecha de evaluacion *", value=def_fecha_dt, key="dt_fecha",
            min_value=FECHA_MIN_PROYECTO, max_value=date.today())
        hora_reg = rB.text_input("Hora de registro", value=def_hora, key="dt_hora")
        correlativo = rC.text_input("Ficha N° (correlativo)", value=def_correlativo, key="dt_corr")

        rD, rE, rF = st.columns(3)
        evaluador = rD.text_input("Evaluador / Brigada / Responsable *",
                                  value=def_evaluador or def_brigada, key="dt_eval")
        brigada = rE.text_input("Brigada (si difiere del responsable)",
                                value=def_brigada, key="dt_brigada")
        mc = rF.selectbox("Microcuenca", [""] + MICROCUENCAS, index=mc_idx_dt, key="dt_mc")

        rG, rH, rI = st.columns(3)
        utm_e_in = rG.text_input("UTM Este (m) — punto de muestreo *",
                                 value=def_utm_e, key="dt_utm_e")
        utm_n_in = rH.text_input("UTM Norte (m) — punto de muestreo *",
                                 value=def_utm_n, key="dt_utm_n")
        altitud_gps = rI.text_input("Altitud GPS (msnm) *", value=def_altitud, key="dt_altitud")

        rJ, rK = st.columns(2)
        centro_poblado = rJ.text_input("Centro poblado más cercano",
                                       value=def_cp, key="dt_cp")
        comunidad_campesina = rK.text_input("Comunidad Campesina (si aplica)",
                                            value=def_cc, key="dt_cc")

        st.markdown("---")
        st.markdown("### 2. Fichas de Diagnostico V5")
        st.markdown("*Despliegue cada ficha y complete los parametros relevantes a la visita de campo.*")

        # ╔════════════════════════════════════════════════════════════════╗
        # ║ F-DT-01: DATOS GENERALES Y FISIOGRAFIA DEL BLOQUE              ║
        # ╚════════════════════════════════════════════════════════════════╝
        with st.expander("F-DT-01: DATOS GENERALES Y FISIOGRAFIA DEL BLOQUE", expanded=False):
            st.markdown("**Caracterizacion fisiografica del bloque**")
            c1, c2 = st.columns(2)
            forma_terreno = c1.selectbox(
                "Forma predominante del terreno *",
                [""] + FDT_FORMA_TERRENO,
                index=([""] + FDT_FORMA_TERRENO).index(dt_edit.get("forma_terreno", ""))
                    if dt_edit.get("forma_terreno", "") in FDT_FORMA_TERRENO else 0,
                key="f01_ft")
            pendiente = c2.selectbox(
                "Rango de pendiente dominante *",
                [""] + FDT_RANGO_PENDIENTE,
                index=([""] + FDT_RANGO_PENDIENTE).index(dt_edit.get("pendiente", ""))
                    if dt_edit.get("pendiente", "") in FDT_RANGO_PENDIENTE else 0,
                key="f01_pe")
            c3, c4 = st.columns(2)
            posicion_fisio = c3.selectbox(
                "Posición fisiográfica *",
                [""] + FDT_POSICION_FISIO,
                index=([""] + FDT_POSICION_FISIO).index(dt_edit.get("posicion_fisiografica", ""))
                    if dt_edit.get("posicion_fisiografica", "") in FDT_POSICION_FISIO else 0,
                key="f01_pf")
            exposicion = c4.selectbox(
                "Exposición / Orientación dominante *",
                [""] + FDT_EXPOSICION,
                index=([""] + FDT_EXPOSICION).index(dt_edit.get("exposicion_orientacion", ""))
                    if dt_edit.get("exposicion_orientacion", "") in FDT_EXPOSICION else 0,
                key="f01_ex")
            c5, c6 = st.columns(2)
            rango_alt = c5.selectbox(
                "Rango altitudinal del bloque *",
                [""] + FDT_RANGO_ALTITUD,
                index=([""] + FDT_RANGO_ALTITUD).index(dt_edit.get("rango_altitudinal", ""))
                    if dt_edit.get("rango_altitudinal", "") in FDT_RANGO_ALTITUD else 0,
                key="f01_ra")
            paisaje = c6.text_input("Paisaje dominante (descripción libre)",
                                    value=dt_edit.get("paisaje_dominante", ""), key="f01_pa")

            st.markdown("**Indicios de remoción en masa e inestabilidad**")
            d1, d2 = st.columns(2)
            f01_aflora = d1.selectbox(
                "Presencia de afloramientos rocosos", [""] + FDT_SI_NO,
                index=([""] + FDT_SI_NO).index(dt_edit.get("dt01_afloramientos_rocosos", ""))
                    if dt_edit.get("dt01_afloramientos_rocosos", "") in FDT_SI_NO else 0,
                key="f01_af")
            f01_escarpe = d2.selectbox(
                "Presencia de escarpes activos *", [""] + FDT_SI_NO,
                index=([""] + FDT_SI_NO).index(dt_edit.get("dt01_escarpes_activos", ""))
                    if dt_edit.get("dt01_escarpes_activos", "") in FDT_SI_NO else 0,
                key="f01_es")
            d3, d4 = st.columns(2)
            f01_reptacion = d3.selectbox(
                "Reptación de suelo observada *", [""] + FDT_SI_NO,
                index=([""] + FDT_SI_NO).index(dt_edit.get("dt01_reptacion_suelo", ""))
                    if dt_edit.get("dt01_reptacion_suelo", "") in FDT_SI_NO else 0,
                key="f01_rp")
            f01_desliz = d4.selectbox(
                "Deslizamientos antiguos observados", [""] + FDT_SI_NO,
                index=([""] + FDT_SI_NO).index(dt_edit.get("dt01_deslizamientos_antiguos", ""))
                    if dt_edit.get("dt01_deslizamientos_antiguos", "") in FDT_SI_NO else 0,
                key="f01_de")
            f01_remocion = st.selectbox(
                "Remociones en masa activas *", [""] + FDT_SI_NO,
                index=([""] + FDT_SI_NO).index(dt_edit.get("dt01_remociones_masa_activas", ""))
                    if dt_edit.get("dt01_remociones_masa_activas", "") in FDT_SI_NO else 0,
                key="f01_rm")

            f01_obs = st.text_area("Observaciones F-DT-01",
                                   value=dt_edit.get("dt01_observaciones", ""),
                                   key="f01_obs")

        # ╔════════════════════════════════════════════════════════════════╗
        # ║ F-DT-02: SUELO Y PROCESOS EROSIVOS (INCL. CARCAVAS)            ║
        # ╚════════════════════════════════════════════════════════════════╝
        with st.expander("F-DT-02: SUELO Y PROCESOS EROSIVOS", expanded=False):
            st.markdown("**Descripción de procesos erosivos del suelo**")
            e1, e2 = st.columns(2)
            f02_sell = e1.selectbox(
                "Sellamiento / Costra superficial", [""] + FDT_SI_NO,
                index=([""] + FDT_SI_NO).index(dt_edit.get("dt02_sellamiento_costra", ""))
                    if dt_edit.get("dt02_sellamiento_costra", "") in FDT_SI_NO else 0,
                key="f02_sell")
            f02_compa = e2.selectbox(
                "Compactación por pisoteo de ganado", [""] + FDT_SI_NO,
                index=([""] + FDT_SI_NO).index(dt_edit.get("dt02_compactacion_pisoteo", ""))
                    if dt_edit.get("dt02_compactacion_pisoteo", "") in FDT_SI_NO else 0,
                key="f02_compa")
            e3, e4 = st.columns(2)
            f02_raices = e3.selectbox(
                "Raíces expuestas en superficie", [""] + FDT_SI_NO,
                index=([""] + FDT_SI_NO).index(dt_edit.get("dt02_raices_expuestas", ""))
                    if dt_edit.get("dt02_raices_expuestas", "") in FDT_SI_NO else 0,
                key="f02_raices")
            f02_nivel = e4.selectbox(
                "Nivel general de erosión observado *", [""] + FDT_NIVEL_EROSION,
                index=([""] + FDT_NIVEL_EROSION).index(dt_edit.get("dt02_nivel_erosion_general", ""))
                    if dt_edit.get("dt02_nivel_erosion_general", "") in FDT_NIVEL_EROSION else 0,
                key="f02_nivel")

            st.markdown("**Inventario georreferenciado de cárcavas / surcos** "
                        "(registre una fila por cárcava o grupo homogéneo)")
            carcavas_cols = [
                ("codigo", "Código"), ("tipo", "Tipo"),
                ("utm_e_ini", "UTM E inicio"), ("utm_n_ini", "UTM N inicio"),
                ("utm_e_fin", "UTM E fin"), ("utm_n_fin", "UTM N fin"),
                ("longitud_m", "Long. (m)"), ("prof_m", "Prof. (m)"),
                ("ancho_m", "Ancho (m)"), ("estado", "Estado"),
                ("causa", "Causa principal"), ("foto", "Cód. foto"),
            ]
            carcavas_opts = {
                "tipo": FDT_TIPO_CARCAVA,
                "estado": FDT_ESTADO_CARCAVA,
                "causa": FDT_CAUSA_CARCAVA,
            }
            carcavas_data = _dt_load_json(dt_edit.get("dt02_carcavas_json", ""), [])
            carcavas_rows = _dt_dataeditor(
                "Cárcavas / surcos (10 filas)", carcavas_cols, 10,
                "f02_carcavas_ed", options=carcavas_opts, edit_data=carcavas_data)

            st.markdown("**Síntesis de procesos erosivos del bloque**")
            f1, f2 = st.columns(2)
            f02_sint = f1.selectbox(
                "Nivel general de erosión (síntesis) *", [""] + FDT_NIVEL_EROSION,
                index=([""] + FDT_NIVEL_EROSION).index(dt_edit.get("dt02_nivel_erosion_sintesis", ""))
                    if dt_edit.get("dt02_nivel_erosion_sintesis", "") in FDT_NIVEL_EROSION else 0,
                key="f02_sint")
            f02_num = f2.text_input("N° total de cárcavas registradas",
                                    value=dt_edit.get("dt02_num_carcavas", ""), key="f02_num")
            f3, f4 = st.columns(2)
            f02_long = f3.text_input("Longitud total de cárcavas (m)",
                                     value=dt_edit.get("dt02_longitud_total_carcavas", ""), key="f02_long")
            f02_pct = f4.text_input("% del bloque afectado por cárcavas",
                                    value=dt_edit.get("dt02_pct_bloque_carcavas", ""), key="f02_pct")
            f5, f6 = st.columns(2)
            f02_lam = f5.text_input("Erosión laminar observada (%)",
                                    value=dt_edit.get("dt02_erosion_laminar_pct", ""), key="f02_lam")
            f02_pat = f6.selectbox(
                "Patrón de cárcavas dominante", [""] + FDT_PATRON_CARCAVAS,
                index=([""] + FDT_PATRON_CARCAVAS).index(dt_edit.get("dt02_patron_carcavas", ""))
                    if dt_edit.get("dt02_patron_carcavas", "") in FDT_PATRON_CARCAVAS else 0,
                key="f02_pat")
            f7, f8 = st.columns(2)
            f02_soc = f7.selectbox(
                "Presencia de socavamiento de cauce", [""] + FDT_SI_NO,
                index=([""] + FDT_SI_NO).index(dt_edit.get("dt02_socavamiento_cauce", ""))
                    if dt_edit.get("dt02_socavamiento_cauce", "") in FDT_SI_NO else 0,
                key="f02_soc")
            f02_urg = f8.selectbox(
                "Urgencia de control", [""] + FDT_URGENCIA,
                index=([""] + FDT_URGENCIA).index(dt_edit.get("dt02_urgencia_control", ""))
                    if dt_edit.get("dt02_urgencia_control", "") in FDT_URGENCIA else 0,
                key="f02_urg")
            f02_obs = st.text_area("Observaciones F-DT-02",
                                   value=dt_edit.get("dt02_observaciones", ""), key="f02_obs")

        # ╔════════════════════════════════════════════════════════════════╗
        # ║ F-DT-03: ECOSISTEMA: COMPOSICION, ESTRUCTURA Y VALOR ECOLOGICO ║
        # ╚════════════════════════════════════════════════════════════════╝
        with st.expander("F-DT-03: ECOSISTEMA — COMPOSICION, ESTRUCTURA Y VALOR ECOLOGICO",
                         expanded=False):
            st.markdown("**Datos de la parcela de muestreo**")
            g1, g2, g3 = st.columns(3)
            f03_parcela = g1.text_input("Parcela de muestreo",
                                        value=dt_edit.get("dt03_parcela_muestreo", ""), key="f03_parcela")
            f03_dim = g2.text_input("Dimensiones (m × m)",
                                    value=dt_edit.get("dt03_dim_parcela", ""), key="f03_dim")
            f03_pend = g3.text_input("Pendiente promedio parcela (%)",
                                     value=dt_edit.get("dt03_pendiente_parcela", ""), key="f03_pend")
            g4, g5 = st.columns(2)
            f03_cobtot = g4.text_input("Cobertura vegetal total estimada (%) *",
                                       value=dt_edit.get("dt03_cobertura_total", ""), key="f03_cobtot")
            f03_uso = g5.selectbox(
                "Uso actual dominante del suelo *", [""] + FDT_USO_SUELO_DOM,
                index=([""] + FDT_USO_SUELO_DOM).index(dt_edit.get("dt03_uso_dominante", ""))
                    if dt_edit.get("dt03_uso_dominante", "") in FDT_USO_SUELO_DOM else 0,
                key="f03_uso")

            st.markdown("**Clasificación del ecosistema (UP)**")
            h1, h2 = st.columns(2)
            f03_eco = h1.selectbox(
                "Tipo de ecosistema MINAM (dominante) *", [""] + FDT_TIPO_ECOSISTEMA,
                index=([""] + FDT_TIPO_ECOSISTEMA).index(dt_edit.get("dt03_tipo_ecosistema", ""))
                    if dt_edit.get("dt03_tipo_ecosistema", "") in FDT_TIPO_ECOSISTEMA else 0,
                key="f03_eco")
            f03_supe = h2.text_input("Superficie del ecosistema en el bloque (ha)",
                                     value=dt_edit.get("dt03_superficie_ecosistema", ""), key="f03_supe")
            f03_cons = st.selectbox(
                "Estado de conservación general *", [""] + FDT_ESTADO_CONSERVACION,
                index=([""] + FDT_ESTADO_CONSERVACION).index(dt_edit.get("dt03_estado_conservacion_eco", ""))
                    if dt_edit.get("dt03_estado_conservacion_eco", "") in FDT_ESTADO_CONSERVACION else 0,
                key="f03_cons")

            st.markdown("**Estructura del ecosistema** (porcentajes y alturas)")
            i1, i2 = st.columns(2)
            f03_dosel = i1.text_input("Cobertura del dosel arbóreo (%)",
                                      value=dt_edit.get("dt03_cobertura_dosel", ""), key="f03_dosel")
            f03_arbus = i2.text_input("Cobertura arbustiva (%)",
                                      value=dt_edit.get("dt03_cobertura_arbustiva", ""), key="f03_arbus")
            i3, i4 = st.columns(2)
            f03_herb = i3.text_input("Cobertura herbácea-graminoide (%)",
                                     value=dt_edit.get("dt03_cobertura_herbacea", ""), key="f03_herb")
            f03_hoja = i4.text_input("Cobertura de hojarasca (%)",
                                     value=dt_edit.get("dt03_cobertura_hojarasca", ""), key="f03_hoja")
            i5, i6 = st.columns(2)
            f03_desn = i5.text_input("Cobertura de suelo desnudo (%)",
                                     value=dt_edit.get("dt03_suelo_desnudo", ""), key="f03_desn")
            f03_haltd = i6.text_input("Altura promedio estrato dominante (m)",
                                      value=dt_edit.get("dt03_altura_estrato_dom", ""), key="f03_haltd")
            i7, i8 = st.columns(2)
            f03_hmax = i7.text_input("Altura máxima observada (m)",
                                     value=dt_edit.get("dt03_altura_max", ""), key="f03_hmax")
            f03_dap = i8.text_input("DAP promedio árboles DAP≥10 cm (cm)",
                                    value=dt_edit.get("dt03_dap_promedio", ""), key="f03_dap")

            j1, j2 = st.columns(2)
            f03_regen = j1.selectbox(
                "Regeneración natural observada *", [""] + FDT_REGENERACION,
                index=([""] + FDT_REGENERACION).index(dt_edit.get("dt03_regeneracion_natural", ""))
                    if dt_edit.get("dt03_regeneracion_natural", "") in FDT_REGENERACION else 0,
                key="f03_regen")
            f03_san = j2.selectbox(
                "Estado sanitario general del dosel *", [""] + FDT_ESTADO_SANITARIO,
                index=([""] + FDT_ESTADO_SANITARIO).index(dt_edit.get("dt03_estado_sanitario", ""))
                    if dt_edit.get("dt03_estado_sanitario", "") in FDT_ESTADO_SANITARIO else 0,
                key="f03_san")
            j3, j4 = st.columns(2)
            f03_epif = j3.text_input("Presencia de epífitas (cualitativa)",
                                     value=dt_edit.get("dt03_presencia_epifitas", ""), key="f03_epif")
            f03_feno = j4.selectbox(
                "Fenología dominante (época de visita)", [""] + FDT_FENOLOGIA,
                index=([""] + FDT_FENOLOGIA).index(dt_edit.get("dt03_fenologia_dominante", ""))
                    if dt_edit.get("dt03_fenologia_dominante", "") in FDT_FENOLOGIA else 0,
                key="f03_feno")
            f03_tipocob = st.selectbox(
                "Tipo de cobertura vegetal dominante *", [""] + FDT_TIPO_COBERTURA,
                index=([""] + FDT_TIPO_COBERTURA).index(dt_edit.get("dt03_tipo_cobertura_dom", ""))
                    if dt_edit.get("dt03_tipo_cobertura_dom", "") in FDT_TIPO_COBERTURA else 0,
                key="f03_tipocob")

            st.markdown("**Composición florística (mín. 15 especies)**")
            flora_cols = [
                ("nombre_comun", "Nombre común"),
                ("nombre_cientifico", "Nombre científico"),
                ("familia", "Familia"),
                ("estrato", "Estrato"),
                ("origen", "Origen"),
                ("abundancia", "Abundancia"),
                ("dap_cm", "DAP (cm)"),
                ("altura_m", "Altura (m)"),
            ]
            flora_opts = {
                "estrato": FDT_ESTRATO,
                "origen": FDT_ORIGEN,
                "abundancia": FDT_ABUNDANCIA,
            }
            flora_data = _dt_load_json(dt_edit.get("dt03_floristica_json", ""), [])
            flora_rows = _dt_dataeditor(
                "Especies (15 filas)", flora_cols, 15,
                "f03_flora_ed", options=flora_opts, edit_data=flora_data)

            st.markdown("**Especies clave / indicadoras**")
            esp_cols = [
                ("nombre", "Nombre científico/común"),
                ("categoria", "Categoría"),
                ("estado_uicn", "Estado UICN / D.S. 043"),
                ("utm_e", "UTM Este"), ("utm_n", "UTM Norte"),
                ("n_indiv", "N° indiv."), ("foto", "Foto N°"),
                ("observacion", "Observación"),
            ]
            esp_opts = {
                "categoria": FDT_CATEGORIA_INDICADORA,
                "estado_uicn": FDT_ESTADO_UICN,
            }
            esp_data = _dt_load_json(dt_edit.get("dt03_especies_clave_json", ""), [])
            esp_rows = _dt_dataeditor(
                "Especies clave (10 filas)", esp_cols, 10,
                "f03_esp_ed", options=esp_opts, edit_data=esp_data)

            f03_obs = st.text_area("Observaciones F-DT-03",
                                   value=dt_edit.get("dt03_observaciones", ""), key="f03_obs")

        # ╔════════════════════════════════════════════════════════════════╗
        # ║ F-DT-04: CAUSAS E INDICADORES DE DEGRADACION                   ║
        # ╚════════════════════════════════════════════════════════════════╝
        with st.expander("F-DT-04: CAUSAS E INDICADORES DE DEGRADACION", expanded=False):
            st.markdown("**Matriz de causas de degradación** "
                        "(intensidad: Nula/Ligera/Moderada/Fuerte/Muy fuerte)")
            causas_cols = [
                ("n", "N°"), ("causa", "Causa / Factor"),
                ("presencia", "Presencia"), ("intensidad", "Intensidad"),
                ("extension", "Extensión (%)"),
                ("antiguedad", "Antigüedad (años)"),
                ("evidencia", "Evidencia / Descripción"),
            ]
            causas_opts = {
                "presencia": FDT_SI_NO,
                "intensidad": FDT_INTENSIDAD,
            }
            # Datos predefinidos: 16 causas fijas con presencia/intensidad vacios
            causas_existentes = _dt_load_json(dt_edit.get("dt04_causas_json", ""), [])
            causas_default = []
            for i, lbl in enumerate(FDT04_CAUSAS_LABELS, 1):
                row = {"n": str(i), "causa": lbl, "presencia": "", "intensidad": "",
                       "extension": "", "antiguedad": "", "evidencia": ""}
                if i - 1 < len(causas_existentes):
                    prev = causas_existentes[i - 1]
                    for k in row:
                        if prev.get(k):
                            row[k] = prev[k]
                causas_default.append(row)
            causas_rows = _dt_dataeditor(
                "Causas (16 fijas)", causas_cols, 16,
                "f04_causas_ed", options=causas_opts, edit_data=causas_default)

            st.markdown("**Indicadores cuantitativos de degradación**")
            ind_cols = [
                ("n", "N°"), ("indicador", "Indicador"), ("unidad", "Unidad"),
                ("valor", "Valor"), ("fuente", "Fuente"),
                ("umbral", "Umbral / Referencia"), ("nivel", "Nivel"),
            ]
            ind_opts = {"nivel": FDT_NIVEL_IND}
            ind_existentes = _dt_load_json(dt_edit.get("dt04_indicadores_json", ""), [])
            ind_default = []
            for i, (nombre, unidad, umbral) in enumerate(FDT04_INDICADORES_LABELS, 1):
                row = {"n": str(i), "indicador": nombre, "unidad": unidad,
                       "valor": "", "fuente": "", "umbral": umbral, "nivel": ""}
                if i - 1 < len(ind_existentes):
                    prev = ind_existentes[i - 1]
                    for k in row:
                        if prev.get(k):
                            row[k] = prev[k]
                ind_default.append(row)
            ind_rows = _dt_dataeditor(
                "Indicadores (8 fijos)", ind_cols, 8,
                "f04_ind_ed", options=ind_opts, edit_data=ind_default)

            st.markdown("**Síntesis diagnóstica de la degradación**")
            f04_dir = st.text_area(
                "Principales causas directas (texto libre)",
                value=dt_edit.get("dt04_causas_directas_texto", ""), key="f04_dir")
            k1, k2 = st.columns(2)
            f04_sub = k1.text_input("Principal causa subyacente (motor)",
                                    value=dt_edit.get("dt04_causa_subyacente", ""), key="f04_sub")
            f04_vel = k2.selectbox(
                "Velocidad de degradación percibida", [""] + FDT_VELOCIDAD,
                index=([""] + FDT_VELOCIDAD).index(dt_edit.get("dt04_velocidad_degradacion", ""))
                    if dt_edit.get("dt04_velocidad_degradacion", "") in FDT_VELOCIDAD else 0,
                key="f04_vel")
            k3, k4 = st.columns(2)
            f04_rev = k3.selectbox(
                "Reversibilidad técnica", [""] + FDT_REVERSIBILIDAD,
                index=([""] + FDT_REVERSIBILIDAD).index(dt_edit.get("dt04_reversibilidad", ""))
                    if dt_edit.get("dt04_reversibilidad", "") in FDT_REVERSIBILIDAD else 0,
                key="f04_rev")
            f04_urg = k4.selectbox(
                "Urgencia de intervención", [""] + FDT_URGENCIA,
                index=([""] + FDT_URGENCIA).index(dt_edit.get("dt04_urgencia_intervencion", ""))
                    if dt_edit.get("dt04_urgencia_intervencion", "") in FDT_URGENCIA else 0,
                key="f04_urg")
            f04_obs = st.text_area("Observaciones F-DT-04",
                                   value=dt_edit.get("dt04_observaciones", ""), key="f04_obs")

        # ╔════════════════════════════════════════════════════════════════╗
        # ║ F-DT-05: RECURSOS HIDRICOS Y ACCESIBILIDAD                     ║
        # ╚════════════════════════════════════════════════════════════════╝
        with st.expander("F-DT-05: RECURSOS HIDRICOS Y ACCESIBILIDAD", expanded=False):
            st.markdown("**Inventario de fuentes de agua** (dentro o hasta 500 m del bloque)")
            fuente_cols = [
                ("n", "N°"), ("tipo", "Tipo de fuente"),
                ("utm_e", "UTM Este"), ("utm_n", "UTM Norte"),
                ("regimen", "Régimen"), ("calidad", "Calidad aparente"),
                ("distancia_m", "Distancia (m)"),
                ("uso_obs", "Uso observado / Obs."),
            ]
            fuente_opts = {
                "tipo": FDT_TIPO_FUENTE,
                "regimen": FDT_REGIMEN_HIDRICO,
                "calidad": FDT_CALIDAD_AGUA,
            }
            fuentes_existentes = _dt_load_json(dt_edit.get("dt05_fuentes_agua_json", ""), [])
            fuente_default = []
            for i in range(1, 11):
                row = {"n": str(i), "tipo": "", "utm_e": "", "utm_n": "",
                       "regimen": "", "calidad": "", "distancia_m": "", "uso_obs": ""}
                if i - 1 < len(fuentes_existentes):
                    prev = fuentes_existentes[i - 1]
                    for k in row:
                        if prev.get(k):
                            row[k] = prev[k]
                fuente_default.append(row)
            fuente_rows = _dt_dataeditor(
                "Fuentes (10 filas)", fuente_cols, 10,
                "f05_fuentes_ed", options=fuente_opts, edit_data=fuente_default)

            st.markdown("**Análisis hídrico preliminar**")
            l1, l2 = st.columns(2)
            f05_recarga = l1.selectbox(
                "¿Bloque en zona de recarga hídrica?", [""] + FDT_SI_NO_NA,
                index=([""] + FDT_SI_NO_NA).index(dt_edit.get("dt05_zona_recarga", ""))
                    if dt_edit.get("dt05_zona_recarga", "") in FDT_SI_NO_NA else 0,
                key="f05_recarga")
            f05_humedad = l2.selectbox(
                "¿Zonas de humedad persistente observadas?", [""] + FDT_SI_NO,
                index=([""] + FDT_SI_NO).index(dt_edit.get("dt05_humedad_persistente", ""))
                    if dt_edit.get("dt05_humedad_persistente", "") in FDT_SI_NO else 0,
                key="f05_humedad")
            l3, l4 = st.columns(2)
            f05_escor = l3.selectbox(
                "¿Escorrentía concentrada observada?", [""] + FDT_SI_NO,
                index=([""] + FDT_SI_NO).index(dt_edit.get("dt05_escorrentia_concentrada", ""))
                    if dt_edit.get("dt05_escorrentia_concentrada", "") in FDT_SI_NO else 0,
                key="f05_escor")
            f05_distcap = l4.text_input("Distancia a captación poblacional más cercana (m)",
                                        value=dt_edit.get("dt05_dist_captacion", ""), key="f05_distcap")
            l5, l6 = st.columns(2)
            f05_jass = l5.text_input("Nombre JASS o captación asociada",
                                     value=dt_edit.get("dt05_jass_captacion", ""), key="f05_jass")
            f05_inter = l6.selectbox(
                "¿Interferencia con obras de riego?", [""] + FDT_SI_NO,
                index=([""] + FDT_SI_NO).index(dt_edit.get("dt05_interferencia_riego", ""))
                    if dt_edit.get("dt05_interferencia_riego", "") in FDT_SI_NO else 0,
                key="f05_inter")
            f05_riego = st.text_input("Nombre del sistema de riego (si aplica)",
                                      value=dt_edit.get("dt05_sistema_riego_nombre", ""), key="f05_riego")

            st.markdown("**Accesibilidad y logística**")
            m1, m2 = st.columns(2)
            f05_modo = m1.selectbox(
                "Modalidad de acceso al bloque *", [""] + FDT_MODALIDAD_ACCESO,
                index=([""] + FDT_MODALIDAD_ACCESO).index(dt_edit.get("dt05_modalidad_acceso", ""))
                    if dt_edit.get("dt05_modalidad_acceso", "") in FDT_MODALIDAD_ACCESO else 0,
                key="f05_modo")
            f05_via = m2.text_input("Vía principal de acceso (PE-/PI-/trocha) *",
                                    value=dt_edit.get("dt05_via_principal", ""), key="f05_via")
            m3, m4 = st.columns(2)
            f05_tvia = m3.selectbox(
                "Tipo de vía final", [""] + FDT_TIPO_VIA,
                index=([""] + FDT_TIPO_VIA).index(dt_edit.get("dt05_tipo_via_final", ""))
                    if dt_edit.get("dt05_tipo_via_final", "") in FDT_TIPO_VIA else 0,
                key="f05_tvia")
            f05_tseca = m4.selectbox(
                "Transitabilidad — época seca", [""] + FDT_NIVEL_TRANSITAB,
                index=([""] + FDT_NIVEL_TRANSITAB).index(dt_edit.get("dt05_transitabilidad_seca", ""))
                    if dt_edit.get("dt05_transitabilidad_seca", "") in FDT_NIVEL_TRANSITAB else 0,
                key="f05_tseca")
            m5, m6 = st.columns(2)
            f05_tllu = m5.selectbox(
                "Transitabilidad — época lluviosa", [""] + FDT_NIVEL_TRANSITAB,
                index=([""] + FDT_NIVEL_TRANSITAB).index(dt_edit.get("dt05_transitabilidad_lluviosa", ""))
                    if dt_edit.get("dt05_transitabilidad_lluviosa", "") in FDT_NIVEL_TRANSITAB else 0,
                key="f05_tllu")
            f05_tcap = m6.text_input("Tiempo desde capital distrital (min)",
                                     value=dt_edit.get("dt05_tiempo_dist_capital", ""), key="f05_tcap")
            m7, m8 = st.columns(2)
            f05_tprov = m7.text_input("Tiempo desde capital provincial (min)",
                                      value=dt_edit.get("dt05_tiempo_prov_capital", ""), key="f05_tprov")
            f05_senal = m8.selectbox(
                "Señal celular", [""] + FDT_SENAL_CELULAR,
                index=([""] + FDT_SENAL_CELULAR).index(dt_edit.get("dt05_senal_celular", ""))
                    if dt_edit.get("dt05_senal_celular", "") in FDT_SENAL_CELULAR else 0,
                key="f05_senal")
            m9, m10 = st.columns(2)
            f05_oper = m9.selectbox(
                "Operador celular dominante", [""] + FDT_OPERADOR,
                index=([""] + FDT_OPERADOR).index(dt_edit.get("dt05_operador_celular", ""))
                    if dt_edit.get("dt05_operador_celular", "") in FDT_OPERADOR else 0,
                key="f05_oper")
            f05_aloj = m10.selectbox(
                "Alojamiento rural disponible", [""] + FDT_SI_NO,
                index=([""] + FDT_SI_NO).index(dt_edit.get("dt05_alojamiento", ""))
                    if dt_edit.get("dt05_alojamiento", "") in FDT_SI_NO else 0,
                key="f05_aloj")
            m11, m12 = st.columns(2)
            f05_ronda = m11.selectbox(
                "¿Requiere autorización de Ronda Campesina? *", [""] + FDT_SI_NO,
                index=([""] + FDT_SI_NO).index(dt_edit.get("dt05_requiere_ronda", ""))
                    if dt_edit.get("dt05_requiere_ronda", "") in FDT_SI_NO else 0,
                key="f05_ronda")
            f05_contacto = m12.text_input("Nombre / Contacto responsable de Ronda",
                                          value=dt_edit.get("dt05_contacto_ronda", ""), key="f05_contacto")
            f05_obs = st.text_area("Observaciones F-DT-05",
                                   value=dt_edit.get("dt05_observaciones", ""), key="f05_obs")

        st.markdown("---")
        observ_gen = st.text_area("Observaciones generales del diagnóstico",
                                  value=dt_edit.get("observaciones_generales", ""),
                                  key="dt_obs")

        # ── Determinar fichas con datos ─────────────────────────────────
        fichas_sel = []
        if any([forma_terreno, pendiente, posicion_fisio, exposicion, rango_alt,
                paisaje, f01_aflora, f01_escarpe, f01_reptacion, f01_desliz,
                f01_remocion, f01_obs]):
            fichas_sel.append("F-DT-01")
        if any([f02_sell, f02_compa, f02_raices, f02_nivel, f02_sint,
                f02_num, f02_long, f02_pct, f02_lam, f02_pat, f02_soc,
                f02_urg, f02_obs]) or any(any(r.values()) for r in carcavas_rows):
            fichas_sel.append("F-DT-02")
        if any([f03_parcela, f03_dim, f03_pend, f03_cobtot, f03_uso,
                f03_eco, f03_supe, f03_cons, f03_dosel, f03_arbus, f03_herb,
                f03_hoja, f03_desn, f03_haltd, f03_hmax, f03_dap, f03_regen,
                f03_san, f03_epif, f03_feno, f03_tipocob, f03_obs]) or \
           any(any(r.values()) for r in flora_rows) or \
           any(any(r.values()) for r in esp_rows):
            fichas_sel.append("F-DT-03")
        if any([f04_dir, f04_sub, f04_vel, f04_rev, f04_urg, f04_obs]) or \
           any(r.get("presencia") or r.get("intensidad") or r.get("extension")
               or r.get("antiguedad") or r.get("evidencia") for r in causas_rows) or \
           any(r.get("valor") or r.get("fuente") or r.get("nivel") for r in ind_rows):
            fichas_sel.append("F-DT-04")
        if any([f05_recarga, f05_humedad, f05_escor, f05_distcap, f05_jass,
                f05_inter, f05_riego, f05_modo, f05_via, f05_tvia, f05_tseca,
                f05_tllu, f05_tcap, f05_tprov, f05_senal, f05_oper, f05_aloj,
                f05_ronda, f05_contacto, f05_obs]) or \
           any(any(r.values()) for r in fuente_rows):
            fichas_sel.append("F-DT-05")

        if fichas_sel:
            st.info(f"Fichas con datos: **{', '.join(fichas_sel)}** ({len(fichas_sel)}/5)")

        btn_label_dt = "Actualizar Diagnostico Territorial" if dt_edit_id else "Guardar Diagnostico Territorial"
        if st.button(btn_label_dt, type="primary", key="dt_guardar"):
            if not evaluador:
                st.warning("Ingrese el nombre del evaluador / responsable.")
            elif not fichas_sel:
                st.warning("Complete al menos una ficha de diagnostico.")
            else:
                try:
                    data_v5 = {
                        "ficha": ", ".join(fichas_sel),
                        "fecha_evaluacion": fecha_ev.strftime("%Y-%m-%d"),
                        "evaluador": evaluador,
                        "microcuenca": mc,
                        "brigada": brigada,
                        "ficha_correlativo": correlativo,
                        "altitud_gps": altitud_gps,
                        "centro_poblado_cercano": centro_poblado,
                        "comunidad_campesina_dt": comunidad_campesina,
                        "hora_registro": hora_reg,
                        "utm_este_dt": utm_e_in,
                        "utm_norte_dt": utm_n_in,
                        # F-DT-01
                        "forma_terreno": forma_terreno,
                        "pendiente": pendiente,
                        "posicion_fisiografica": posicion_fisio,
                        "exposicion_orientacion": exposicion,
                        "rango_altitudinal": rango_alt,
                        "paisaje_dominante": paisaje,
                        "dt01_afloramientos_rocosos": f01_aflora,
                        "dt01_escarpes_activos": f01_escarpe,
                        "dt01_reptacion_suelo": f01_reptacion,
                        "dt01_deslizamientos_antiguos": f01_desliz,
                        "dt01_remociones_masa_activas": f01_remocion,
                        "dt01_observaciones": f01_obs,
                        # F-DT-02
                        "dt02_sellamiento_costra": f02_sell,
                        "dt02_compactacion_pisoteo": f02_compa,
                        "dt02_raices_expuestas": f02_raices,
                        "dt02_nivel_erosion_general": f02_nivel,
                        "dt02_carcavas_json": _dt_dump_json(carcavas_rows),
                        "dt02_nivel_erosion_sintesis": f02_sint,
                        "dt02_num_carcavas": f02_num,
                        "dt02_longitud_total_carcavas": f02_long,
                        "dt02_pct_bloque_carcavas": f02_pct,
                        "dt02_erosion_laminar_pct": f02_lam,
                        "dt02_patron_carcavas": f02_pat,
                        "dt02_socavamiento_cauce": f02_soc,
                        "dt02_urgencia_control": f02_urg,
                        "dt02_observaciones": f02_obs,
                        # F-DT-03
                        "dt03_parcela_muestreo": f03_parcela,
                        "dt03_dim_parcela": f03_dim,
                        "dt03_pendiente_parcela": f03_pend,
                        "dt03_cobertura_total": f03_cobtot,
                        "dt03_tipo_ecosistema": f03_eco,
                        "dt03_superficie_ecosistema": f03_supe,
                        "dt03_estado_conservacion_eco": f03_cons,
                        "dt03_uso_dominante": f03_uso,
                        "dt03_cobertura_dosel": f03_dosel,
                        "dt03_cobertura_arbustiva": f03_arbus,
                        "dt03_cobertura_herbacea": f03_herb,
                        "dt03_cobertura_hojarasca": f03_hoja,
                        "dt03_suelo_desnudo": f03_desn,
                        "dt03_altura_estrato_dom": f03_haltd,
                        "dt03_altura_max": f03_hmax,
                        "dt03_dap_promedio": f03_dap,
                        "dt03_regeneracion_natural": f03_regen,
                        "dt03_estado_sanitario": f03_san,
                        "dt03_presencia_epifitas": f03_epif,
                        "dt03_fenologia_dominante": f03_feno,
                        "dt03_tipo_cobertura_dom": f03_tipocob,
                        "dt03_floristica_json": _dt_dump_json(flora_rows),
                        "dt03_especies_clave_json": _dt_dump_json(esp_rows),
                        "dt03_observaciones": f03_obs,
                        # F-DT-04
                        "dt04_causas_json": _dt_dump_json(causas_rows),
                        "dt04_indicadores_json": _dt_dump_json(ind_rows),
                        "dt04_causas_directas_texto": f04_dir,
                        "dt04_causa_subyacente": f04_sub,
                        "dt04_velocidad_degradacion": f04_vel,
                        "dt04_reversibilidad": f04_rev,
                        "dt04_urgencia_intervencion": f04_urg,
                        "dt04_observaciones": f04_obs,
                        # F-DT-05
                        "dt05_fuentes_agua_json": _dt_dump_json(fuente_rows),
                        "dt05_zona_recarga": f05_recarga,
                        "dt05_humedad_persistente": f05_humedad,
                        "dt05_escorrentia_concentrada": f05_escor,
                        "dt05_dist_captacion": f05_distcap,
                        "dt05_jass_captacion": f05_jass,
                        "dt05_interferencia_riego": f05_inter,
                        "dt05_sistema_riego_nombre": f05_riego,
                        "dt05_modalidad_acceso": f05_modo,
                        "dt05_via_principal": f05_via,
                        "dt05_tipo_via_final": f05_tvia,
                        "dt05_transitabilidad_seca": f05_tseca,
                        "dt05_transitabilidad_lluviosa": f05_tllu,
                        "dt05_tiempo_dist_capital": f05_tcap,
                        "dt05_tiempo_prov_capital": f05_tprov,
                        "dt05_senal_celular": f05_senal,
                        "dt05_operador_celular": f05_oper,
                        "dt05_alojamiento": f05_aloj,
                        "dt05_requiere_ronda": f05_ronda,
                        "dt05_contacto_ronda": f05_contacto,
                        "dt05_observaciones": f05_obs,
                        # Comun
                        "observaciones_generales": observ_gen,
                    }
                    if dt_edit_id:
                        db.actualizar_diagnostico_territorial_v5(dt_edit_id, data_v5)
                        _invalidar_cache()
                        st.session_state["dt_edit_id"] = None
                        st.session_state["dt_edit_data"] = None
                        _flash(f"Diagnostico territorial ID {dt_edit_id} actualizado "
                               f"correctamente ({', '.join(fichas_sel)}).")
                    else:
                        existentes = db.obtener_diagnosticos_por_bloque(bid)
                        dup = [e for e in existentes
                               if e.get("fecha_evaluacion") == fecha_ev.strftime("%Y-%m-%d")
                               and e.get("evaluador") == evaluador]
                        if dup:
                            st.markdown(
                                '<div class="dup-warning">Ya existe un diagnostico para este bloque en '
                                f'{fecha_ev.strftime("%Y-%m-%d")} por {evaluador}. '
                                f'Use <b>Editar</b> en Historial para modificarlo.</div>',
                                unsafe_allow_html=True)
                        else:
                            db.insertar_diagnostico_territorial_v5(bid, data_v5)
                            _invalidar_cache()
                            _flash(f"Diagnostico territorial guardado ({', '.join(fichas_sel)}).")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

    with tab_hist:
        st.markdown("### Historial de Diagnosticos Territoriales")
        st.caption("Haga clic en **Editar** para modificar un diagnostico existente o en **Eliminar** para descartarlo.")
        if st.session_state.get("dt_edit_id"):
            st.info("✏️ Registro cargado en **modo edicion**. Abra la pestaña "
                    "**'Registro de Diagnostico'** (arriba) para corregir los campos y "
                    "pulsar **Actualizar Diagnostico Territorial**.")
        todos_dt = _cached_obtener_todos_diagnosticos(_cache_version())
        if not todos_dt:
            st.info("No hay diagnosticos registrados.")
        else:
            st.download_button(
                "⬇️ Descargar todo (Excel)",
                data=exp_diag.exportar_fdt_consolidado(todos_dt),
                file_name=f"Diagnostico_Territorial_FDT_{datetime.now():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dt_export_excel", type="secondary",
                help="Genera un unico archivo Excel con una hoja por ficha (F-DT-01..05) "
                     "y hojas adicionales para las tablas (carcavas, floristica, causas, etc.).")
            dt_pag, total_pags_dt, pag_actual_dt = _paginar(todos_dt, "pag_dt")
            _controles_paginacion(total_pags_dt, pag_actual_dt, "pag_dt")
            col_widths_dt = [0.4, 1, 0.8, 0.8, 0.8, 0.8, 0.8, 0.5, 0.6]
            header_cols = st.columns(col_widths_dt)
            for col, h in zip(header_cols,
                              ["ID", "Bloque", "Fichas", "Fecha", "Evaluador",
                               "Microcuenca", "Distrito", "", ""]):
                col.markdown(f"**{h}**")
            st.markdown("---")
            confirm_del_dt_id = st.session_state.get("dt_confirm_del_id")
            for d in dt_pag:
                row = st.columns(col_widths_dt)
                row[0].write(d["id"])
                row[1].write(d.get("bloque_codigo", ""))
                row[2].write(d.get("ficha", ""))
                row[3].write(d.get("fecha_evaluacion", ""))
                row[4].write(d.get("evaluador", ""))
                row[5].write(d.get("microcuenca", "") or "")
                row[6].write(d.get("distrito", ""))
                if row[7].button("Editar", key=f"edit_dt_{d['id']}", type="primary"):
                    det = db.obtener_diagnostico_por_id(d["id"])
                    if det:
                        # Limpiar los widgets para que la precarga (value=/index=)
                        # del registro a editar surta efecto.
                        _dt_limpiar_widgets_edicion()
                        st.session_state["dt_edit_id"] = det["id"]
                        st.session_state["dt_edit_data"] = det
                        # Posicionar el selector de bloque en el bloque del registro.
                        cod = det.get("bloque_codigo", "")
                        for label_bl in bm:
                            if cod and cod in label_bl:
                                st.session_state["dt_bl"] = label_bl
                                break
                    st.session_state.pop("dt_confirm_del_id", None)
                    st.rerun()
                if confirm_del_dt_id == d["id"]:
                    if row[8].button("Confirmar", key=f"del_confirm_dt_{d['id']}", type="primary"):
                        db.eliminar_diagnostico(d["id"])
                        _invalidar_cache()
                        st.session_state.pop("dt_confirm_del_id", None)
                        st.success(f"Diagnostico territorial ID {d['id']} eliminado.")
                        st.rerun()
                else:
                    if row[8].button("Eliminar", key=f"del_dt_{d['id']}", type="secondary"):
                        st.session_state["dt_confirm_del_id"] = d["id"]
                        st.rerun()
            if confirm_del_dt_id is not None:
                st.warning(f"Confirme la eliminacion del diagnostico ID {confirm_del_dt_id} pulsando 'Confirmar' en su fila. "
                           "Esta accion no puede deshacerse.")
                if st.button("Cancelar eliminacion", key="dt_cancel_del"):
                    st.session_state.pop("dt_confirm_del_id", None)
                    st.rerun()

            st.markdown("---")
            st.markdown("### Detalle de Diagnostico")
            dm = {f"ID {d['id']} - {d.get('bloque_codigo','')} ({d.get('ficha','')})": d["id"]
                  for d in todos_dt}
            sel_dt = st.selectbox("Seleccionar diagnostico", [""] + list(dm.keys()), key="dt_det")
            if sel_dt and sel_dt in dm:
                det = db.obtener_diagnostico_por_id(dm[sel_dt])
                if det:
                    st.markdown(f"**Bloque:** {det.get('bloque_codigo','')} | "
                                f"**Fecha:** {det.get('fecha_evaluacion','')} | "
                                f"**Evaluador:** {det.get('evaluador','')}")
                    fichas_str = det.get("ficha", "")

                    if "F-DT-01" in fichas_str:
                        with st.expander("F-DT-01: Datos generales y fisiografía", expanded=True):
                            c1, c2 = st.columns(2)
                            c1.markdown(f"**Forma terreno:** {det.get('forma_terreno','') or '-'}")
                            c2.markdown(f"**Pendiente:** {det.get('pendiente','') or '-'}")
                            c1.markdown(f"**Posición fisio.:** {det.get('posicion_fisiografica','') or '-'}")
                            c2.markdown(f"**Exposición:** {det.get('exposicion_orientacion','') or '-'}")
                            c1.markdown(f"**Altitud:** {det.get('rango_altitudinal','') or '-'}")
                            c2.markdown(f"**Paisaje:** {det.get('paisaje_dominante','') or '-'}")
                            c1.markdown(f"**Afloramientos rocosos:** {det.get('dt01_afloramientos_rocosos','') or '-'}")
                            c2.markdown(f"**Escarpes activos:** {det.get('dt01_escarpes_activos','') or '-'}")
                            c1.markdown(f"**Reptación suelo:** {det.get('dt01_reptacion_suelo','') or '-'}")
                            c2.markdown(f"**Deslizamientos antiguos:** {det.get('dt01_deslizamientos_antiguos','') or '-'}")
                            st.markdown(f"**Remociones masa activas:** {det.get('dt01_remociones_masa_activas','') or '-'}")
                            if det.get("dt01_observaciones"):
                                st.markdown(f"**Obs:** {det['dt01_observaciones']}")

                    if "F-DT-02" in fichas_str:
                        with st.expander("F-DT-02: Suelo y procesos erosivos", expanded=True):
                            c1, c2 = st.columns(2)
                            c1.markdown(f"**Sellamiento/costra:** {det.get('dt02_sellamiento_costra','') or '-'}")
                            c2.markdown(f"**Compactación:** {det.get('dt02_compactacion_pisoteo','') or '-'}")
                            c1.markdown(f"**Raíces expuestas:** {det.get('dt02_raices_expuestas','') or '-'}")
                            c2.markdown(f"**Nivel erosión:** {det.get('dt02_nivel_erosion_general','') or '-'}")
                            cars = _dt_load_json(det.get("dt02_carcavas_json", ""), [])
                            if cars:
                                st.markdown(f"**Cárcavas registradas:** {len(cars)}")
                                st.dataframe(pd.DataFrame(cars), use_container_width=True, hide_index=True)
                            c1.markdown(f"**Nivel erosión (síntesis):** {det.get('dt02_nivel_erosion_sintesis','') or '-'}")
                            c2.markdown(f"**N° cárcavas:** {det.get('dt02_num_carcavas','') or '-'}")
                            c1.markdown(f"**Longitud total (m):** {det.get('dt02_longitud_total_carcavas','') or '-'}")
                            c2.markdown(f"**% bloque afectado:** {det.get('dt02_pct_bloque_carcavas','') or '-'}")
                            c1.markdown(f"**Patrón cárcavas:** {det.get('dt02_patron_carcavas','') or '-'}")
                            c2.markdown(f"**Urgencia control:** {det.get('dt02_urgencia_control','') or '-'}")
                            if det.get("dt02_observaciones"):
                                st.markdown(f"**Obs:** {det['dt02_observaciones']}")

                    if "F-DT-03" in fichas_str:
                        with st.expander("F-DT-03: Ecosistema", expanded=True):
                            c1, c2 = st.columns(2)
                            c1.markdown(f"**Tipo ecosistema:** {det.get('dt03_tipo_ecosistema','') or '-'}")
                            c2.markdown(f"**Estado conservación:** {det.get('dt03_estado_conservacion_eco','') or '-'}")
                            c1.markdown(f"**Uso dominante:** {det.get('dt03_uso_dominante','') or '-'}")
                            c2.markdown(f"**Cobertura total (%):** {det.get('dt03_cobertura_total','') or '-'}")
                            c1.markdown(f"**Tipo cobertura dom.:** {det.get('dt03_tipo_cobertura_dom','') or '-'}")
                            c2.markdown(f"**Regeneración:** {det.get('dt03_regeneracion_natural','') or '-'}")
                            flora = _dt_load_json(det.get("dt03_floristica_json", ""), [])
                            if flora:
                                st.markdown(f"**Especies registradas:** {len(flora)}")
                                st.dataframe(pd.DataFrame(flora), use_container_width=True, hide_index=True)
                            esp = _dt_load_json(det.get("dt03_especies_clave_json", ""), [])
                            if esp:
                                st.markdown(f"**Especies clave/indicadoras:** {len(esp)}")
                                st.dataframe(pd.DataFrame(esp), use_container_width=True, hide_index=True)
                            if det.get("dt03_observaciones"):
                                st.markdown(f"**Obs:** {det['dt03_observaciones']}")

                    if "F-DT-04" in fichas_str:
                        with st.expander("F-DT-04: Causas e indicadores de degradación", expanded=True):
                            causas = _dt_load_json(det.get("dt04_causas_json", ""), [])
                            if causas:
                                st.markdown("**Matriz de causas**")
                                st.dataframe(pd.DataFrame(causas), use_container_width=True, hide_index=True)
                            inds = _dt_load_json(det.get("dt04_indicadores_json", ""), [])
                            if inds:
                                st.markdown("**Indicadores cuantitativos**")
                                st.dataframe(pd.DataFrame(inds), use_container_width=True, hide_index=True)
                            c1, c2 = st.columns(2)
                            c1.markdown(f"**Causa subyacente:** {det.get('dt04_causa_subyacente','') or '-'}")
                            c2.markdown(f"**Velocidad degradación:** {det.get('dt04_velocidad_degradacion','') or '-'}")
                            c1.markdown(f"**Reversibilidad:** {det.get('dt04_reversibilidad','') or '-'}")
                            c2.markdown(f"**Urgencia intervención:** {det.get('dt04_urgencia_intervencion','') or '-'}")
                            if det.get("dt04_causas_directas_texto"):
                                st.markdown(f"**Causas directas:** {det['dt04_causas_directas_texto']}")
                            if det.get("dt04_observaciones"):
                                st.markdown(f"**Obs:** {det['dt04_observaciones']}")

                    if "F-DT-05" in fichas_str:
                        with st.expander("F-DT-05: Recursos hídricos y accesibilidad", expanded=True):
                            fts = _dt_load_json(det.get("dt05_fuentes_agua_json", ""), [])
                            if fts:
                                st.markdown(f"**Fuentes de agua registradas:** {len(fts)}")
                                st.dataframe(pd.DataFrame(fts), use_container_width=True, hide_index=True)
                            c1, c2 = st.columns(2)
                            c1.markdown(f"**Zona de recarga:** {det.get('dt05_zona_recarga','') or '-'}")
                            c2.markdown(f"**Humedad persistente:** {det.get('dt05_humedad_persistente','') or '-'}")
                            c1.markdown(f"**Modalidad acceso:** {det.get('dt05_modalidad_acceso','') or '-'}")
                            c2.markdown(f"**Vía principal:** {det.get('dt05_via_principal','') or '-'}")
                            c1.markdown(f"**Transit. seca:** {det.get('dt05_transitabilidad_seca','') or '-'}")
                            c2.markdown(f"**Transit. lluviosa:** {det.get('dt05_transitabilidad_lluviosa','') or '-'}")
                            c1.markdown(f"**Señal celular:** {det.get('dt05_senal_celular','') or '-'}")
                            c2.markdown(f"**Operador:** {det.get('dt05_operador_celular','') or '-'}")
                            c1.markdown(f"**Requiere autoriz. Ronda:** {det.get('dt05_requiere_ronda','') or '-'}")
                            c2.markdown(f"**Contacto Ronda:** {det.get('dt05_contacto_ronda','') or '-'}")
                            if det.get("dt05_observaciones"):
                                st.markdown(f"**Obs:** {det['dt05_observaciones']}")

                    if det.get("observaciones_generales"):
                        st.markdown(f"**Observaciones generales:** {det['observaciones_generales']}")

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
        st.caption("Genere la plantilla V5 y suba el archivo llenado por el técnico.")

        st.markdown("---")
        st.markdown("**1. Descargar Plantilla V5 para Tecnicos**")
        col_dl1, col_dl2 = st.columns(2)
        fichas_descarga_dt = col_dl1.multiselect(
            "Fichas a incluir en la plantilla",
            FICHAS_DT, default=FICHAS_DT, key="dt_excel_fichas_dl")
        if col_dl2.button("Generar Plantilla V5", type="secondary", key="dt_gen_plantilla"):
            if fichas_descarga_dt:
                bloques_data_dt = [
                    (b[1], b[2], b[4], b[5], b[3], b[8], b[9],
                     (b[11] if len(b) > 11 else ""))
                    for b in BLOQUES_V5
                ]
                plantilla_bytes_dt = generar_plantilla_dt(fichas_descarga_dt, bloques_data_dt)
                st.session_state["dt_plantilla_bytes"] = plantilla_bytes_dt
                st.success("Plantilla V5 generada correctamente.")

        if st.session_state.get("dt_plantilla_bytes"):
            st.download_button(
                "Descargar Plantilla V5 (.xlsx)",
                st.session_state["dt_plantilla_bytes"],
                file_name="Plantilla_DT_Campo_Check_Validada_V5.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dt_dl_plantilla")

        st.markdown("---")
        st.markdown("**2. Subir Excel V5 Llenado**")
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
                    st.error("No se pudieron detectar fichas V5 en el archivo.")
                else:
                    st.success(f"Se detectaron {len(resultados_dt)} ficha(s) en el archivo.")
                    fichas_detectadas = [r["ficha"] for r in resultados_dt]
                    datos_todos = {}
                    for r in resultados_dt:
                        datos_todos.update(r["datos"])

                    with st.expander("Vista previa de datos detectados", expanded=True):
                        c1, c2, c3 = st.columns(3)
                        c1.markdown(f"**Fecha:** {datos_todos.get('fecha_evaluacion', '-')}")
                        c2.markdown(f"**Evaluador:** {datos_todos.get('evaluador', '-')}")
                        c3.markdown(f"**Bloque:** {datos_todos.get('codigo_bloque', '-')}")
                        st.markdown(f"**Fichas detectadas:** {', '.join(fichas_detectadas)}")
                        campos_pob = sum(1 for v in datos_todos.values() if v)
                        st.markdown(f"**Campos con datos:** {campos_pob}")

                    if st.button("Autocompletar formulario", type="primary", key="dt_autocompletar"):
                        ss_vals = mapear_dt_a_session_state(
                            {"ficha": ", ".join(fichas_detectadas), "datos": datos_todos}, bm)
                        st.session_state["_dt_autocompletar_pending"] = ss_vals
                        st.success(
                            f"Formulario autocompletado con {len(fichas_detectadas)} ficha(s). "
                            "Cambie a la pestaña **Registro de Diagnostico** para revisar y guardar.")
                        st.rerun()
            except Exception as e:
                st.error(f"Error al leer el archivo Excel: {e}")

# ══════════════════════════════════════════════════════════════════════════
# DIAGNOSTICO SOCIAL
# ══════════════════════════════════════════════════════════════════════════
def _ds_datos_generales(bloque_label=""):
    """Campos de datos generales compartidos por todas las fichas DS.
    Auto-vincula centros poblados, comunidades campesinas, provincia,
    distrito y coordenadas aproximadas del bloque seleccionado."""
    codigo_bloque = bloque_label.split(" - ")[0].strip() if " - " in bloque_label else bloque_label.strip()
    bloque_info = BLOQUES_128_MAP.get(codigo_bloque, {})
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


# ══════════════════════════════════════════════════════════════════════════
# DIAGNOSTICO SOCIAL V3 — Plantilla validada (F-DS-01 a F-DS-07)
# Formularios con celdas de validacion (desplegables) que replican las listas
# oficiales del Excel `Plantilla_Diagnostico_Social_IN_Piura_V4`.
# Cada ficha guarda su formulario completo como JSON en la columna dsNN_data_v3
# y, donde aplica, alimenta las columnas legacy para compatibilidad de reportes.
# Convencion: la clave del dict del formulario == la key del widget Streamlit,
# por lo que la precarga (edicion / import) es generica y sin desfases.
# ══════════════════════════════════════════════════════════════════════════

# Slots que corresponden a tablas (st.data_editor) y no a widgets escalares.
_DS_TABLE_SLOTS = {
    "f1_activ", "f2_actores", "f4_part", "f4_agenda", "f4_acuerdos",
    "f5_conflictos", "f5_oportunidades", "f6_peligros", "f6_cambios",
}


def _ds_nonce():
    return st.session_state.get("_ds_nonce", 0)


def _T(f, container, label, key, **kw):
    f[key] = container.text_input(label, key=key, **kw)
    return f[key]


def _TA(f, label, key, height=80, **kw):
    f[key] = st.text_area(label, key=key, height=height, **kw)
    return f[key]


def _SB(f, container, label, options, key, help=None):
    """Selectbox 'marcar uno' (celda de validacion)."""
    opts = [""] + list(options)
    if key in st.session_state and st.session_state[key] not in opts:
        st.session_state[key] = ""
    f[key] = container.selectbox(label, opts, key=key, help=help)
    return f[key]


def _MS(f, container, label, options, key, help=None):
    """Multiselect 'marcar las que apliquen' (celda de validacion)."""
    opts = list(options)
    if key in st.session_state:
        st.session_state[key] = [v for v in st.session_state[key] if v in opts]
    f[key] = container.multiselect(label, opts, key=key, help=help)
    return f[key]


def _ds_tabla(label, slot, columns, default_rows=None, default_n=3, help=None):
    """Editor de tabla con celdas de validacion por columna (st.data_editor)."""
    if label:
        st.markdown(label)
    if help:
        st.caption(help)
    col_names = [c[0] for c in columns]
    init = st.session_state.get(f"_dsinit_{slot}")
    if init is None:
        if default_rows is not None:
            init = [dict(r) for r in default_rows]
        else:
            init = [{c: "" for c in col_names} for _ in range(default_n)]
    rows = [{c: r.get(c, "") for c in col_names} for r in init]
    df = pd.DataFrame(rows, columns=col_names)
    colcfg = {}
    for name, kind, opts in columns:
        if kind == "select":
            colcfg[name] = st.column_config.SelectboxColumn(name, options=list(opts), width="medium")
        else:
            colcfg[name] = st.column_config.TextColumn(name)
    edited = st.data_editor(
        df, key=f"_dstbl_{slot}_{_ds_nonce()}", column_config=colcfg,
        num_rows="dynamic", use_container_width=True, hide_index=True)
    out = []
    for r in edited.to_dict("records"):
        if any(str(v).strip() for v in r.values() if v is not None):
            out.append({k: ("" if v is None else str(v)) for k, v in r.items()})
    return out


def _TB(f, label, slot, columns, **kw):
    f[slot] = _ds_tabla(label, slot, columns, **kw)
    return f[slot]


# ─── Render de cada ficha: devuelven el dict completo del formulario V3 ──────

def _render_fds01():
    f = {}
    st.markdown("**1. Datos del Entrevistado/a**")
    c1, c2, c3 = st.columns(3)
    _T(f, c1, "Nombres y apellidos del entrevistado", "f1_entrevistado")
    _T(f, c2, "DNI", "f1_entrevistado_dni")
    _T(f, c3, "Oficio/ocupacion", "f1_entrevistado_oficio")

    st.markdown("**2. Datos Demograficos**")
    _SB(f, st, "2.1 Tipo de organizacion territorial", FL.L_ORG_COMUNAL, "f1_org_terr")
    c1, c2 = st.columns(2)
    _T(f, c1, "Nombre oficial completo (CP/CC/Anexo)", "f1_nombre_oficial")
    _T(f, c2, "Año de fundacion", "f1_anio_fund")
    c1, c2, c3 = st.columns(3)
    _T(f, c1, "N total de familias / viviendas *", "f1_nfam")
    _T(f, c2, "Poblacion total estimada (hab.) *", "f1_pob_t")
    _T(f, c3, "Pob. autoidentificada originaria", "f1_pob_orig")
    c1, c2, c3, c4 = st.columns(4)
    _T(f, c1, "Pob. hombres", "f1_pob_h")
    _T(f, c2, "Pob. mujeres", "f1_pob_m")
    _T(f, c3, "Pob. < 18 años", "f1_pob_men18")
    _T(f, c4, "Pob. > 65 años", "f1_pob_may65")
    c1, c2 = st.columns(2)
    _SB(f, c1, "2.3 Idioma predominante", FL.L_IDIOMA, "f1_idioma")
    _SB(f, c2, "2.4 Nivel educativo predominante", FL.L_NIV_EDU, "f1_nivel_edu")
    c1, c2 = st.columns(2)
    _SB(f, c1, "2.5 Tasa de migracion juvenil", FL.L_ABC, "f1_migracion")
    _SB(f, c2, "2.6 Destino principal de migracion", FL.L_MIG_DEST, "f1_destino_mig")

    st.markdown("**3. Organizacion Comunal y Gobernanza Local**")
    c1, c2, c3 = st.columns(3)
    _SB(f, c1, "¿Junta Directiva vigente?", DS_SINO, "f1_junta_vig")
    _T(f, c2, "Fin de mandato (AAAA)", "f1_junta_fin")
    _SB(f, c3, "3.1 Periodicidad de asambleas", FL.L_PERIODIC, "f1_periodicidad")
    c1, c2, c3 = st.columns(3)
    _T(f, c1, "Presidente/a de Junta", "f1_pres_junta")
    _T(f, c2, "DNI", "f1_pres_dni")
    _T(f, c3, "Telefono", "f1_pres_tel")
    c1, c2, c3 = st.columns(3)
    _SB(f, c1, "¿Ronda Campesina activa? *", DS_SINO, "f1_ronda")
    _T(f, c2, "Presidente/a de Ronda", "f1_pres_ronda")
    _T(f, c3, "Telefono Pres. Ronda", "f1_ronda_tel")
    _T(f, st, "Autoridad adicional (Tnte Gobernador, Agente Municipal)", "f1_aut_adic")
    c1, c2, c3 = st.columns(3)
    _SB(f, c1, "¿Existe reglamento interno?", DS_SINONA, "f1_reglamento")
    _SB(f, c2, "¿Comite de RRNN?", DS_SINO, "f1_comite_rrnn")
    _T(f, c3, "Nombre del comite (si existe)", "f1_nombre_comite")

    st.markdown("**4. Tenencia de la Tierra**")
    _SB(f, st, "4.1 Regimen predominante de tenencia", FL.L_TENENCIA, "f1_tenencia")
    c1, c2 = st.columns(2)
    _T(f, c1, "N aprox. de predios individuales", "f1_n_predios")
    _T(f, c2, "% tierras tituladas", "f1_pct_tituladas")
    c1, c2 = st.columns(2)
    _SB(f, c1, "¿Conflictos de linderos registrados?", DS_SINO, "f1_conf_linderos")
    _SB(f, c2, "¿Bloque se superpone a tierras comunales?", DS_SINO, "f1_superpone")
    _SB(f, st, "4.2 Organismo responsable del registro", FL.L_REG_TIT, "f1_reg_titulacion")

    st.markdown("**5. Servicios Basicos e Infraestructura Social**")
    c1, c2, c3, c4 = st.columns(4)
    _SB(f, c1, "5.1 Agua para consumo", FL.L_AGUA, "f1_agua")
    _T(f, c2, "Cobertura agua (%)", "f1_agua_cob")
    _T(f, c3, "Pago mensual S/", "f1_agua_pago")
    _T(f, c4, "Tiempo acarreo / viajes", "f1_agua_acarreo")
    c1, c2 = st.columns(2)
    _SB(f, c1, "5.2 Saneamiento", FL.L_SANEA, "f1_sanea")
    _SB(f, c2, "5.3 Energia electrica", FL.L_ENERG, "f1_energia")
    c1, c2 = st.columns(2)
    _T(f, c1, "Cobertura energia (%)", "f1_energia_cob")
    _T(f, c2, "Pago mensual energia S/", "f1_energia_pago")
    c1, c2 = st.columns(2)
    _SB(f, c1, "5.4 Telecomunicaciones", FL.L_TELECOM, "f1_telecom")
    _T(f, c2, "Operador celular dominante", "f1_telecom_op")
    c1, c2, c3 = st.columns(3)
    _SB(f, c1, "5.5 IE - niveles disponibles", FL.L_NIV_EDU_IE, "f1_ie_niveles")
    _T(f, c2, "Nombre IE principal", "f1_ie_nombre")
    _T(f, c3, "Distancia al CP (Km)", "f1_ie_dist")
    c1, c2, c3 = st.columns(3)
    _SB(f, c1, "5.6 EESS - categoria", FL.L_CAT_EESS, "f1_eess")
    _T(f, c2, "Nombre EESS mas cercano", "f1_eess_nombre")
    _T(f, c3, "Distancia EESS (km)", "f1_eess_dist")
    c1, c2 = st.columns(2)
    _SB(f, c1, "¿Existe local comunal?", DS_SINO, "f1_local_comunal")
    _SB(f, c2, "Estado del local comunal", FL.L_BRM, "f1_local_estado")

    _TB(f, "**6. Actividades Economicas y Medios de Vida**", "f1_activ",
        [("Actividad / Rubro", "select", FL.L_ACTIV),
         ("N fam.", "text", None),
         ("Productos principales", "text", None),
         ("Destino", "select", FL.L_DESTINO),
         ("Ingreso (S/./mes)", "text", None)],
        default_n=3,
        help="Registre hasta 8 actividades. Elija la actividad y el destino del desplegable.")

    st.markdown("**7. Programas Sociales y Presencia Institucional**")
    c1, c2, c3 = st.columns(3)
    _T(f, c1, "Beneficiarios JUNTOS (N fam.)", "f1_juntos")
    _T(f, c2, "Pension 65 (N pers.)", "f1_pension65")
    _T(f, c3, "Qali Warma (IIEE)", "f1_qaliwarma")
    c1, c2 = st.columns(2)
    _T(f, c1, "Beca 18 (N pers.)", "f1_beca18")
    _T(f, c2, "Otros programas (especificar)", "f1_otros_prog")
    c1, c2, c3 = st.columns(3)
    _SB(f, c1, "¿Proyectos AGRORURAL?", DS_SINO, "f1_agrorural")
    _SB(f, c2, "¿PRODERN / FONCODES?", DS_SINO, "f1_prodern")
    _SB(f, c3, "¿Otros proyectos?", DS_SINO, "f1_otros_proy")
    c1, c2, c3 = st.columns(3)
    _SB(f, c1, "¿PDC vigente?", DS_SINO, "f1_pdc")
    _SB(f, c2, "¿ONGs operando?", DS_SINO, "f1_ongs")
    _T(f, c3, "Nombre de ONGs", "f1_nombre_ongs")
    _SB(f, st, "7.1 Percepcion de presencia estatal", FL.L_ABC, "f1_presencia_estatal")
    return f


def _render_fds02():
    f = {}
    st.info("TIPO de actor, Influencia, Interes, Posicion y Nivel territorial se eligen "
            "de los desplegables (valores completos). Ver hoja _Codigos del Excel para la leyenda.")
    _TB(f, "**3. Registro de Actores Identificados**", "f2_actores",
        [("Nombre del actor / Organizacion", "text", None),
         ("Tipo", "select", FL.L_TIPO_ACTOR),
         ("Rol / Funcion frente al proyecto", "text", None),
         ("Influencia", "select", FL.L_ABC),
         ("Interes", "select", FL.L_ABC),
         ("Posicion", "select", FL.L_POSICION),
         ("Nivel territorial", "select", FL.L_NIV_TERR),
         ("Telefono", "text", None),
         ("Correo / Contacto", "text", None),
         ("Observaciones / Historial", "text", None)],
        default_n=5,
        help="Registre todos los actores relevantes para la zona del bloque.")
    st.markdown("**5. Actores Criticos Prioritarios (sintesis)**")
    _T(f, st, "Actor mas influyente a favor", "f2_favor")
    _T(f, st, "Actor mas influyente en contra / reticente", "f2_contra")
    _T(f, st, "Actor clave en decisiones comunales", "f2_decision")
    _T(f, st, "Actor clave para coordinar acceso (Ronda)", "f2_ronda")
    _T(f, st, "Plataforma / Mesa existente que podria servir de vehiculo", "f2_plataforma")
    return f


def _render_fds03():
    f = {}
    st.markdown("**2. Datos del Entrevistado/a**")
    c1, c2, c3 = st.columns(3)
    _T(f, c1, "Nombres y apellidos *", "f3_nombre")
    _T(f, c2, "DNI", "f3_dni")
    _T(f, c3, "Edad", "f3_edad")
    c1, c2, c3 = st.columns(3)
    _SB(f, c1, "Genero", ["M", "F"], "f3_genero")
    _T(f, c2, "Cargo / Rol *", "f3_cargo")
    _T(f, c3, "Institucion / Organizacion", "f3_inst")
    c1, c2, c3 = st.columns(3)
    _T(f, c1, "Años en el cargo", "f3_anios")
    _T(f, c2, "Telefono", "f3_tel")
    _T(f, c3, "Correo", "f3_correo")
    c1, c2 = st.columns(2)
    _T(f, c1, "Lugar de la entrevista", "f3_lugar")
    _T(f, c2, "Duracion (min)", "f3_dur")
    c1, c2 = st.columns(2)
    _SB(f, c1, "¿Consiente uso de nombre en el informe?", DS_SINO, "f3_c_nom")
    _SB(f, c2, "¿Consiente toma de fotografias?", DS_SINO, "f3_c_foto")

    st.markdown("**3. Percepcion del Territorio y de los Recursos Naturales**")
    _TA(f, "3.1 ¿Que especies de arboles y arbustos nativos/silvestres existen actualmente "
           "y cuales han disminuido o desaparecido?", "f3_r1")
    _TA(f, "3.2 En los ultimos 10-15 años, ¿ha observado cambios en la disponibilidad de "
           "agua, bosque, suelo? Describa.", "f3_r2")
    _TA(f, "3.3 ¿Cuales son los principales problemas ambientales actuales (deforestacion, "
           "erosion, sequias, contaminacion, carcavas, huaicos)? ¿con que frecuencia e intensidad?", "f3_r3")
    _TA(f, "3.4 ¿Que zonas del territorio (bosques, manantiales, etc) considera mas "
           "importantes para conservar y por que?", "f3_r4")

    st.markdown("**4. Actividades Productivas y Medios de Vida**")
    _TA(f, "4.1 ¿Cuales son las actividades economicas mas importantes de la comunidad y "
           "quienes se dedican a cada una (hombres/mujeres)?", "f3_r5")
    _TA(f, "4.2 ¿Como se abastecen de agua para riego y consumo? ¿Hay deficit en alguna "
           "epoca del año, cuando? ¿Hay conflictos por el agua?", "f3_r6")
    _TA(f, "4.3 ¿Utilizan productos del bosque? ¿Cuales?", "f3_r7")
    _TA(f, "4.4 ¿Existen cadenas productivas organizadas o asociaciones de productores en la zona?", "f3_r8")

    st.markdown("**5. Organizacion, Gobernanza y Relacion con el Estado**")
    _TA(f, "5.1 ¿Que organizaciones sociales de base existen en la comunidad y cual es su "
           "rol en la toma de decisiones sobre el territorio?", "f3_r9")
    _TA(f, "5.2 ¿Como se toman las decisiones sobre el uso de recursos comunales (agua, bosque, pastos)?", "f3_r10")
    _TA(f, "5.3 ¿Como ha sido la relacion historica con el Estado y con proyectos de inversion "
           "externos (incluidos los mineros)?", "f3_r11")

    st.markdown("**6. Expectativas frente al Proyecto IN Piura**")
    _TA(f, "6.1 ¿Usted se encuentra de acuerdo con el proyecto o tiene alguna duda?", "f3_r_acuerdo")
    _TA(f, "6.2 ¿En que dias de la semana y horarios puede o preferiria participar en "
           "actividades relacionadas al proyecto?", "f3_r_horarios")
    _TA(f, "6.3 ¿Hay antecedentes de conflictos por proyectos externos (mineros, hidrocarburos, "
           "infraestructura) que ANIN deba tener en cuenta?", "f3_r_ant")
    _TA(f, "7. Cierre y observaciones del entrevistador/a", "f3_cierre")
    return f


def _render_fds04():
    f = {}
    st.markdown("**2. Datos del Taller**")
    c1, c2 = st.columns(2)
    _T(f, c1, "Lugar del taller *", "f4_lugar")
    _T(f, c2, "Entidad convocante", "f4_conv")
    c1, c2, c3 = st.columns(3)
    _T(f, c1, "Fecha del taller *", "f4_fecha")
    _T(f, c2, "Hora inicio", "f4_hi")
    _T(f, c3, "Hora fin", "f4_hf")
    st.caption("2.1 Asistencia")
    c1, c2, c3 = st.columns(3)
    _T(f, c1, "N convocados", "f4_conv_n")
    _T(f, c2, "N hombres", "f4_h")
    _T(f, c3, "N mujeres", "f4_m")
    c1, c2, c3 = st.columns(3)
    _T(f, c1, "N jovenes (<30)", "f4_jov")
    _T(f, c2, "N adultos mayores (>60)", "f4_am")
    _T(f, c3, "N total asistentes", "f4_tot")
    _TA(f, "Objetivo general del taller *", "f4_obj", height=60)
    METODOS = ["Exposicion magistral", "Mesas de trabajo / Grupos focales", "Mapa parlante",
               "Matriz Foda / Vester", "Cartografia participativa", "Lluvia de ideas",
               "Entrevistas semiestructuradas", "Sociodrama / Dinamicas"]
    MATERIALES = ["Papelografos", "Plumones / Lapices", "Tarjetas de colores", "Cinta adhesiva",
                  "Mapas impresos / Ortofotos", "Proyector / Laptop", "Camara fotografica",
                  "Grabadora de audio", "Refrigerio", "Otros"]
    _MS(f, st, "2.2 Metodologia empleada", METODOS, "f4_metod")
    _MS(f, st, "2.3 Materiales utilizados", MATERIALES, "f4_mater")
    _SB(f, st, "2.4 Idioma de la facilitacion",
        ["Español", "Quechua", "Bilingue español-quechua", "Otro"], "f4_idioma")

    _TB(f, "**3. Lista de Participantes**", "f4_part",
        [("Nombres y Apellidos", "text", None), ("DNI", "text", None),
         ("Institucion / Comunidad", "text", None), ("Cargo / Rol", "text", None),
         ("Telefono", "text", None), ("Sexo", "select", ["M", "F"]),
         ("Edad", "text", None)],
        default_n=10)
    _TB(f, "**4. Agenda Desarrollada**", "f4_agenda",
        [("Hora", "text", None), ("Agenda", "text", None),
         ("Responsable", "text", None), ("Resultado / Aporte", "text", None)],
        default_n=3)
    _TB(f, "**5. Acuerdos y Compromisos**", "f4_acuerdos",
        [("Acuerdo / Compromiso", "text", None), ("Responsable", "text", None),
         ("Plazo", "text", None), ("Medio de verificacion", "text", None)],
        default_n=3)
    return f


def _render_fds05():
    f = {}
    st.info("TIPO: SM=Minero | SH=Hidrico | SF=Forestal ... ESTADO: LT=Latente | ES=Escalada | "
            "AC=Activo | NG=Negociacion | RS=Resuelto (ver hoja _Codigos).")
    _TB(f, "**2. Identificacion de Conflictos**", "f5_conflictos",
        [("Tipo", "select", FL.L_TIPO_CONFL),
         ("Actores involucrados", "text", None),
         ("Estado", "select", FL.L_NIVEL_CONFL),
         ("Antiguedad", "select", FL.L_ANTIG_CONFL),
         ("Descripcion / Causa raiz", "text", None),
         ("Impacto potencial en el proyecto", "text", None)],
        default_n=3,
        help="Registre conflictos activos, latentes y resueltos recientes.")

    st.markdown("**3. Contexto especifico - Conflicto minero Rio Blanco y otros**")
    c1, c2 = st.columns(2)
    _SB(f, c1, "¿Vinculo historico con conflicto Rio Blanco?", DS_SINONA, "f5_rb1")
    _SB(f, c2, "¿Fue parte de la consulta de 2007?", DS_SINONA, "f5_rb2")
    c1, c2 = st.columns(2)
    _SB(f, c1, "¿Persiste sentimiento anti-minero fuerte?", DS_SINONA, "f5_rb3")
    _SB(f, c2, "¿Liderazgos activos anti-mineros?", DS_SINONA, "f5_rb4")
    _SB(f, st, "¿Se diferencia el proyecto IN del contexto minero?", DS_SINONA, "f5_rb5")
    _T(f, st, "Otros conflictos relevantes (Tambogrande, Majaz, otros)", "f5_otros")
    _T(f, st, "Rol de las rondas en conflictos historicos", "f5_rol_rondas")
    _SB(f, st, "3.1 Nivel de polarizacion actual",
        ["Muy alto", "Alto", "Medio", "Bajo", "Muy bajo / Inexistente"], "f5_polar")

    _TB(f, "**4. Identificacion de Oportunidades**", "f5_oportunidades",
        [("Oportunidad identificada", "text", None),
         ("Actores relacionados", "text", None),
         ("Tipo (alianza / plataforma / proy.)", "text", None),
         ("Potencial", "select", FL.L_ABC),
         ("Como aprovecharla", "text", None)],
        default_n=3)

    st.markdown("**5. Sintesis Estrategica**")
    _SB(f, st, "5.1 Nivel global de conflictividad",
        ["Muy bajo", "Bajo", "Medio", "Alto", "Muy alto"], "f5_confglob")
    _SB(f, st, "5.2 Viabilidad social preliminar",
        ["Alta — viable para intervencion inmediata", "Media — requiere acercamiento reforzado",
         "Baja — requiere mesa de dialogo previa", "Muy baja — reevaluar inclusion del bloque"], "f5_viab")
    _T(f, st, "Estrategia de acercamiento recomendada", "f5_estrategia")
    c1, c2 = st.columns(2)
    _SB(f, c1, "¿Se requiere mesa de dialogo especifica?", DS_SINO, "f5_mesa")
    _SB(f, c2, "5.3 Plazo estimado para aceptacion comunal",
        ["Inmediato (<1 mes)", "Corto (1-3 meses)", "Medio (3-6 meses)",
         "Largo (6-12 meses)", "Requiere >12 meses"], "f5_plazo")
    return f


def _render_fds06():
    f = {}
    _T(f, st, "Fuente de informacion (N personas consultadas)", "f6_fuente")
    peligros_default = [{"Peligro observado": p} for p in FL.L_PELIGRO_OBS[:-1]]
    _TB(f, "**2. Percepcion de Peligros Naturales**", "f6_peligros",
        [("Peligro observado", "text", None),
         ("¿Ocurre?", "select", DS_SINO),
         ("Frecuencia", "select", FL.FDS06_FRECUENCIA),
         ("Magnitud", "select", FL.FDS06_MAGNITUD),
         ("Tendencia", "select", FL.FDS06_TENDENCIA),
         ("Ultimo evento (año)", "text", None),
         ("Principales daños observados", "text", None)],
        default_rows=peligros_default,
        help="Para cada peligro indique ocurrencia, frecuencia, magnitud y tendencia (desplegables).")

    cambios_default = [{"Cambio observado": c} for c in FL.L_CAMBIO_CLIMA[:-1]]
    _TB(f, "**3. Cambios Climaticos Observados (ultimos 10-15 años)**", "f6_cambios",
        [("Cambio observado", "text", None),
         ("¿Se percibe?", "select", DS_SINO),
         ("Intensidad", "select", FL.FDS06_INTENSIDAD),
         ("Año aprox. de inicio", "text", None),
         ("Impacto en la comunidad / territorio", "text", None)],
        default_rows=cambios_default)

    st.markdown("**4. Respuestas y Adaptaciones Locales**")
    c1, c2 = st.columns(2)
    _SB(f, c1, "¿La comunidad ha tomado medidas?", DS_SINO, "f6_medidas")
    _SB(f, c2, "¿Sistemas de alerta temprana comunitarios?", DS_SINO, "f6_alerta")
    c1, c2 = st.columns(2)
    _SB(f, c1, "¿Saberes tradicionales de prediccion?", DS_SINO, "f6_saberes")
    _SB(f, c2, "¿Se requiere apoyo externo para adaptacion?", DS_SINO, "f6_apoyo")
    _TA(f, "Describa medidas / sistemas / saberes / apoyo requerido", "f6_desc")

    st.markdown("**5. Priorizacion Local de Peligros**")
    c1, c2, c3 = st.columns(3)
    _T(f, c1, "Peligro mas grave (1)", "f6_p1")
    _T(f, c2, "Peligro mas grave (2)", "f6_p2")
    _T(f, c3, "Peligro mas grave (3)", "f6_p3")
    _SB(f, st, "¿Percepcion diferenciada por genero?", DS_SINO, "f6_genero")
    _T(f, st, "Percepcion diferenciada (describa)", "f6_genero_desc")
    return f


def _render_fds07():
    f = {}
    st.markdown("**2. Datos del Titular / Representante**")
    _SB(f, st, "2.1 Tipo de propietario/representante *", FL.L_TIPO_PROPIETARIO, "f7_tipo_prop")
    c1, c2, c3 = st.columns(3)
    _T(f, c1, "Nombres y apellidos *", "f7_nombre")
    _T(f, c2, "DNI *", "f7_dni")
    _T(f, c3, "Edad", "f7_edad")
    c1, c2 = st.columns(2)
    _SB(f, c1, "Genero", ["Hombre", "Mujer", "Otro / Pref. no decir"], "f7_genero")
    _T(f, c2, "Direccion / Residencia habitual", "f7_residencia")
    c1, c2 = st.columns(2)
    _T(f, c1, "Telefono / Correo de contacto", "f7_contacto")
    _T(f, c2, "Superficie predio en el bloque (ha)", "f7_superficie")
    DOCS = ["Titulo de propiedad SUNARP", "Constancia de posesion (municipal)", "Titulo COFOPRI",
            "Certificado catastral", "Resolucion de adjudicacion", "Sin documentacion disponible"]
    _MS(f, st, "2.2 Documentacion de tenencia disponible", DOCS, "f7_docs")
    c1, c2 = st.columns(2)
    _SB(f, c1, "¿Tiene conflictos de linderos?", DS_SINO, "f7_linderos")
    _SB(f, c2, "¿Es residente permanente?", DS_SINO, "f7_residente")

    st.markdown("**3. Informacion Provista al Titular / Representante**")
    st.caption("Marque Si cuando el punto fue efectivamente explicado y comprendido.")
    PUNTOS = [
        ("f7_info_anin", "3.1 Se explico que es la ANIN"),
        ("f7_info_objetivo", "3.2 Se explico el objetivo del proyecto IN Piura"),
        ("f7_info_no_minero", "3.3 Se explico que NO es actividad minera ni extractiva"),
        ("f7_info_medidas", "3.4 Se explico que medidas podrian implementarse en el predio"),
        ("f7_info_temporalidad", "3.5 Se explico la temporalidad del proyecto"),
        ("f7_info_voluntaria", "3.6 Se explico que la participacion es voluntaria"),
        ("f7_info_actualizada", "3.7 Se explico el derecho a informacion actualizada"),
        ("f7_info_confidencialidad", "3.8 Se explico el tratamiento confidencial de datos"),
        ("f7_info_preguntas", "3.9 Se absolvieron todas las preguntas"),
        ("f7_info_material", "3.10 Se entrego material informativo"),
    ]
    for i in range(0, len(PUNTOS), 2):
        cols = st.columns(2)
        for j, (k, lbl) in enumerate(PUNTOS[i:i + 2]):
            _SB(f, cols[j], lbl, DS_SINO, k)

    st.markdown("**4. Manifestacion de Disposicion a Participar**")
    _SB(f, st, "4.1 Disposicion general a participar *", FL.L_DISPOSICION, "f7_disp")
    _TA(f, "Condiciones o requisitos planteados", "f7_cond", height=70)
    c1, c2 = st.columns(2)
    _SB(f, c1, "¿Requiere consultar con familia/asamblea?", DS_SINO, "f7_consultar")
    _T(f, c2, "Plazo solicitado para respuesta (dias)", "f7_plazo")
    c1, c2, c3 = st.columns(3)
    _SB(f, c1, "¿Autoriza ingreso preliminar?", DS_SINO, "f7_aut_ing")
    _SB(f, c2, "¿Autoriza fotografias del predio?", DS_SINO, "f7_aut_foto")
    _SB(f, c3, "¿Autoriza uso de su nombre?", DS_SINO, "f7_aut_nom")

    st.markdown("**5. Preguntas, Preocupaciones y Compromisos**")
    _TA(f, "5.1 Preguntas / preocupaciones del titular", "f7_preg")
    _TA(f, "5.2 Compromisos adquiridos por ANIN/SESDI", "f7_comp")
    return f


_DS_RENDERERS = {
    "F-DS-01": _render_fds01, "F-DS-02": _render_fds02, "F-DS-03": _render_fds03,
    "F-DS-04": _render_fds04, "F-DS-05": _render_fds05, "F-DS-06": _render_fds06,
    "F-DS-07": _render_fds07,
}


def _ds_legacy_cols(ficha, f):
    """Mapea el formulario V3 a columnas legacy para compatibilidad de reportes."""
    L = {}
    if ficha == "F-DS-01":
        L.update({
            "ds01_entrevistado": f.get("f1_entrevistado", ""),
            "ds01_entrevistado_dni": f.get("f1_entrevistado_dni", ""),
            "ds01_entrevistado_oficio": f.get("f1_entrevistado_oficio", ""),
            "ds01_num_familias": f.get("f1_nfam", ""),
            "ds01_poblacion_hombres": f.get("f1_pob_h", ""),
            "ds01_poblacion_mujeres": f.get("f1_pob_m", ""),
            "ds01_poblacion_total": f.get("f1_pob_t", ""),
            "ds01_idioma": f.get("f1_idioma", ""),
            "ds01_nivel_educativo": f.get("f1_nivel_edu", ""),
            "ds01_tasa_migracion": f.get("f1_migracion", ""),
            "ds01_destino_migracion": f.get("f1_destino_mig", ""),
            "ds01_organizacion_comunal": f.get("f1_org_terr", ""),
            "ds01_junta_directiva": f.get("f1_junta_vig", ""),
            "ds01_presidente_junta": f.get("f1_pres_junta", ""),
            "ds01_agua_potable_tipo": f.get("f1_agua", ""),
            "ds01_agua_potable_cobertura": f.get("f1_agua_cob", ""),
            "ds01_saneamiento": f.get("f1_sanea", ""),
            "ds01_energia_tipo": f.get("f1_energia", ""),
            "ds01_energia_cobertura": f.get("f1_energia_cob", ""),
            "ds01_telecomunicaciones": f.get("f1_telecom", ""),
            "ds01_telecom_operador": f.get("f1_telecom_op", ""),
            "ds01_salud_tipo": f.get("f1_eess", ""),
            "ds01_salud_distancia": f.get("f1_eess_dist", ""),
            "ds01_educacion": f.get("f1_ie_niveles", ""),
            "ds01_actividades_economicas": json.dumps(f.get("f1_activ", []), ensure_ascii=False) if f.get("f1_activ") else "",
        })
    elif ficha == "F-DS-02":
        L["ds02_registro_actores"] = json.dumps(f.get("f2_actores", []), ensure_ascii=False) if f.get("f2_actores") else ""
    elif ficha == "F-DS-03":
        L.update({
            "ds03_nombre_entrevistado": f.get("f3_nombre", ""),
            "ds03_cargo_funcion": f.get("f3_cargo", ""),
            "ds03_institucion": f.get("f3_inst", ""),
            "ds03_telefono_correo": f.get("f3_tel", "") or f.get("f3_correo", ""),
            "ds03_duracion": f.get("f3_dur", ""),
            "ds03_resp_recursos_naturales": f.get("f3_r1", ""),
            "ds03_resp_cambios_ambiente": f.get("f3_r2", ""),
            "ds03_resp_problemas_ambientales": f.get("f3_r3", ""),
            "ds03_resp_zonas_conservacion": f.get("f3_r4", ""),
            "ds03_resp_actividades_economicas": f.get("f3_r5", ""),
            "ds03_resp_abastecimiento_agua": f.get("f3_r6", ""),
            "ds03_resp_productos_bosque": f.get("f3_r7", ""),
            "ds03_resp_cadenas_productivas": f.get("f3_r8", ""),
            "ds03_resp_organizaciones": f.get("f3_r9", ""),
            "ds03_resp_decisiones_territorio": f.get("f3_r10", ""),
            "ds03_resp_expectativas": f.get("f3_r_acuerdo", ""),
            "ds03_resp_condiciones": f.get("f3_r_horarios", ""),
        })
    elif ficha == "F-DS-04":
        L.update({
            "ds04_lugar_taller": f.get("f4_lugar", ""),
            "ds04_convocante": f.get("f4_conv", ""),
            "ds04_hora_inicio": f.get("f4_hi", ""),
            "ds04_hora_fin": f.get("f4_hf", ""),
            "ds04_objetivo": f.get("f4_obj", ""),
            "ds04_lista_participantes": json.dumps(f.get("f4_part", []), ensure_ascii=False) if f.get("f4_part") else "",
        })
    elif ficha == "F-DS-05":
        L["ds05_conflictos"] = json.dumps(f.get("f5_conflictos", []), ensure_ascii=False) if f.get("f5_conflictos") else ""
        L["ds05_oportunidades"] = json.dumps(f.get("f5_oportunidades", []), ensure_ascii=False) if f.get("f5_oportunidades") else ""
    return L


def _ds_build_edit_pending(det):
    """Construye {session_key: valor} para precargar un registro (modo edicion)."""
    ficha = det.get("ficha", "")
    pend = {
        "ds_mc": det.get("microcuenca", "") or "",
        "ds_eval": det.get("evaluador", "") or "",
        "ds_fnum": det.get("ficha_numero", "") or "",
        "ds_prov": det.get("provincia", "") or "",
        "ds_dist": det.get("distrito", "") or "",
        "ds_cpob": det.get("centro_poblado", "") or "",
        "ds_ccam": det.get("comunidad_campesina", "") or "",
        "ds_este": float(det.get("coordenada_este") or 0),
        "ds_norte": float(det.get("coordenada_norte") or 0),
        "ds_alt": float(det.get("altitud") or 0),
        "ds_ubigeo": det.get("codigo_ubigeo", "") or "",
        "ds_obs": det.get("observaciones_generales", "") or "",
        "ds_ficha_sel": ficha,
    }
    if det.get("fecha_evaluacion"):
        try:
            pend["ds_fecha"] = datetime.strptime(det["fecha_evaluacion"], "%Y-%m-%d").date()
        except (ValueError, TypeError):
            pass
    num = ficha.split("-")[-1] if ficha else ""
    form = {}
    raw = det.get(f"ds{num}_data_v3", "") or ""
    if raw:
        try:
            form = json.loads(raw)
        except (json.JSONDecodeError, TypeError):
            form = {}
    for k, v in form.items():
        if k in _DS_TABLE_SLOTS:
            pend[f"_dsinit_{k}"] = v if isinstance(v, list) else []
        else:
            pend[k] = v
    return pend


def _ds_apply_pending():
    """Aplica precarga pendiente (edicion / import) a session_state una vez."""
    pend = st.session_state.pop("_ds_pending_state", None)
    if not pend:
        return
    edit_id = st.session_state.pop("_ds_pending_edit_id", None)
    # Limpiar widgets de la ficha anterior para evitar valores residuales
    st.session_state["_ds_nonce"] = _ds_nonce() + 1
    for k, v in pend.items():
        st.session_state[k] = v
    st.session_state["ds_edit_id"] = edit_id


def _ds_humaniza(key):
    s = key
    for p in ("f1_", "f2_", "f3_", "f4_", "f5_", "f6_", "f7_"):
        if s.startswith(p):
            s = s[len(p):]
            break
    return s.replace("_", " ").capitalize()


def _ds_render_detalle(ficha, form):
    """Muestra el formulario V3 de un registro (vista de Historial)."""
    if not isinstance(form, dict) or not form:
        st.caption("Sin datos detallados V4 para esta ficha.")
        return
    # Tablas primero
    for slot in _DS_TABLE_SLOTS:
        if slot in form and isinstance(form[slot], list) and form[slot]:
            titulo = {
                "f1_activ": "Actividades economicas", "f2_actores": "Registro de actores",
                "f4_part": "Participantes", "f4_agenda": "Agenda", "f4_acuerdos": "Acuerdos",
                "f5_conflictos": "Conflictos", "f5_oportunidades": "Oportunidades",
                "f6_peligros": "Peligros naturales", "f6_cambios": "Cambios climaticos",
            }.get(slot, slot)
            with st.expander(f"{titulo} ({len(form[slot])})", expanded=True):
                st.dataframe(pd.DataFrame(form[slot]), use_container_width=True, hide_index=True)
    # Escalares
    escalares = {k: v for k, v in form.items()
                 if k not in _DS_TABLE_SLOTS and v not in ("", None, [])}
    if escalares:
        with st.expander("Campos de la ficha", expanded=True):
            items = list(escalares.items())
            for i in range(0, len(items), 2):
                cols = st.columns(2)
                for j, (k, v) in enumerate(items[i:i + 2]):
                    if isinstance(v, list):
                        v = ", ".join(str(x) for x in v)
                    cols[j].markdown(f"**{_ds_humaniza(k)}:** {v}")


def pagina_diagnostico_social():
    st.subheader("Diagnostico Social - Fichas de Campo (V4)")
    st.caption("Proyecto IN Piura CUI 2669244 | ANIN - DIME - SESDI | "
               "Plantilla validada F-DS-01 a F-DS-07 con celdas de validacion")
    _mostrar_flash()
    bm = _bloques_map()
    if not bm:
        st.warning("Registre un bloque primero.")
        return

    if "ds_edit_id" not in st.session_state:
        st.session_state["ds_edit_id"] = None

    # Aplicar precarga pendiente (import Excel o edicion) antes de los widgets
    _ds_apply_pending()

    ficha_sel = st.radio(
        "Seleccionar ficha", FICHAS_DS, horizontal=True, key="ds_ficha_sel",
        format_func=lambda x: f"{x} - {FICHAS_DS_TITULOS.get(x, '')}")

    tab_reg, tab_hist, tab_excel = st.tabs(
        ["Registro", "Historial / Consulta", "Importar desde Excel"])

    with tab_reg:
        edit_id = st.session_state.get("ds_edit_id")
        if edit_id:
            st.markdown('<div class="edit-mode-banner"><span class="icon">&#9998;</span> '
                        f'Modo Edicion - Diagnostico Social ID {edit_id}</div>', unsafe_allow_html=True)
            if st.button("Cancelar edicion (nuevo registro)", key="ds_cancel_edit"):
                st.session_state["ds_edit_id"] = None
                st.session_state["_ds_nonce"] = _ds_nonce() + 1
                st.rerun()

        bl = st.selectbox("Bloque de Intervencion", list(bm.keys()), key="ds_bl")
        bid = bm[bl]

        prev_bl = st.session_state.get("_ds_prev_bl", "")
        if prev_bl and prev_bl != bl and not edit_id:
            for k in ("ds_prov", "ds_dist", "ds_cpob", "ds_ccam", "ds_este", "ds_norte"):
                st.session_state.pop(k, None)
        st.session_state["_ds_prev_bl"] = bl

        mc_auto_ds = _resolver_microcuenca(bl)
        if mc_auto_ds:
            mc_idx_ds = MICROCUENCAS.index(mc_auto_ds) + 1
            st.info(f"Microcuenca vinculada automaticamente: **{mc_auto_ds}**")
        else:
            mc_idx_ds = 0

        r1, r2, r3, r4 = st.columns(4)
        mc = r1.selectbox("Microcuenca", [""] + MICROCUENCAS, index=mc_idx_ds, key="ds_mc")
        fecha_ev = r2.date_input("Fecha", value=datetime.now(), key="ds_fecha",
                                 min_value=FECHA_MIN_PROYECTO, max_value=date.today())
        evaluador = r3.text_input("Responsable", key="ds_eval")
        ficha_num = r4.text_input("Ficha N", key="ds_fnum")
        dg = _ds_datos_generales(bloque_label=bl)
        st.markdown("---")
        st.markdown(f"### {ficha_sel}: {FICHAS_DS_TITULOS.get(ficha_sel, '').upper()}")

        renderer = _DS_RENDERERS.get(ficha_sel)
        form = renderer() if renderer else {}

        st.markdown("---")
        observ_gen = st.text_area("Observaciones generales", key="ds_obs")
        adj_files = st.file_uploader(
            "Adjuntar archivos de soporte (PDF, max. 25 MB por archivo)",
            type=["pdf"], accept_multiple_files=True, key="ds_adj_upload")

        btn_label = "Actualizar Diagnostico Social" if edit_id else "Guardar Diagnostico Social"
        if st.button(btn_label, type="primary", key="ds_guardar"):
            if not evaluador:
                st.warning("Ingrese el nombre del responsable.")
            else:
                try:
                    archivos_guardados = []
                    if adj_files:
                        carpeta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "adjuntos_ds")
                        os.makedirs(carpeta, exist_ok=True)
                        for fobj in adj_files:
                            if fobj.size > 25 * 1024 * 1024:
                                st.warning(f"Archivo {fobj.name} excede 25 MB, omitido.")
                                continue
                            nombre = f"{uuid.uuid4().hex[:8]}_{fobj.name}"
                            ruta = os.path.join(carpeta, nombre)
                            with open(ruta, "wb") as out:
                                out.write(fobj.getbuffer())
                            archivos_guardados.append(nombre)

                    num = ficha_sel.split("-")[-1]
                    reg = {
                        "bloque_id": bid, "ficha": ficha_sel,
                        "ficha_numero": ficha_num, "microcuenca": mc,
                        "fecha_evaluacion": fecha_ev.strftime("%Y-%m-%d"),
                        "evaluador": evaluador,
                        "observaciones_generales": observ_gen,
                        f"ds{num}_data_v3": json.dumps(form, ensure_ascii=False),
                    }
                    reg.update(_ds_legacy_cols(ficha_sel, form))
                    if archivos_guardados:
                        reg["archivos_adjuntos"] = "|".join(archivos_guardados)
                    reg.update(dg)

                    if edit_id:
                        db.actualizar_diagnostico_social(edit_id, reg)
                        _invalidar_cache()
                        st.session_state["ds_edit_id"] = None
                        _flash(f"Ficha {ficha_sel} actualizada correctamente (ID {edit_id}).")
                    else:
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
                            _flash(f"Ficha {ficha_sel} guardada correctamente.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

    # ── HISTORIAL ──────────────────────────────────────────────────────
    with tab_hist:
        st.markdown("### Historial de Diagnosticos Sociales")
        st.caption("Pulse **Editar** para modificar un registro o **Eliminar** para descartarlo.")
        if st.session_state.get("ds_edit_id"):
            st.info("✏️ Registro cargado en **modo edicion**. Abra la pestaña "
                    "**'Registro'** (arriba) para corregir los campos y "
                    "pulsar **Actualizar Diagnostico Social**.")
        todos_ds = _cached_obtener_todos_diagnosticos_sociales(_cache_version())
        if not todos_ds:
            st.info("No hay diagnosticos sociales registrados.")
        else:
            st.download_button(
                "⬇️ Descargar todo (Excel)",
                data=exp_diag.exportar_fds_consolidado(todos_ds),
                file_name=f"Diagnostico_Social_FDS_{datetime.now():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="ds_export_excel", type="secondary",
                help="Genera un unico archivo Excel con una hoja por ficha (F-DS-01..07) "
                     "y hojas adicionales para las tablas (actores, participantes, conflictos, etc.).")
            ds_pag, total_pags_ds, pag_actual_ds = _paginar(todos_ds, "pag_ds")
            _controles_paginacion(total_pags_ds, pag_actual_ds, "pag_ds")
            col_widths_ds = [0.4, 0.9, 0.7, 0.8, 0.8, 0.8, 0.8, 0.5, 0.6]
            header_cols = st.columns(col_widths_ds)
            for col, h in zip(header_cols, ["ID", "Bloque", "Ficha", "Fecha", "Responsable", "C.Poblado", "Distrito", "", ""]):
                col.markdown(f"**{h}**")
            st.markdown("---")
            confirm_del_ds_id = st.session_state.get("ds_confirm_del_id")
            for d in ds_pag:
                row = st.columns(col_widths_ds)
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
                        st.session_state["_ds_pending_state"] = _ds_build_edit_pending(det)
                        st.session_state["_ds_pending_edit_id"] = det["id"]
                    st.session_state.pop("ds_confirm_del_id", None)
                    st.rerun()
                if confirm_del_ds_id == d["id"]:
                    if row[8].button("Confirmar", key=f"del_confirm_ds_{d['id']}", type="primary"):
                        db.eliminar_diagnostico_social(d["id"])
                        _invalidar_cache()
                        st.session_state.pop("ds_confirm_del_id", None)
                        st.success(f"Diagnostico social ID {d['id']} eliminado.")
                        st.rerun()
                else:
                    if row[8].button("Eliminar", key=f"del_ds_{d['id']}", type="secondary"):
                        st.session_state["ds_confirm_del_id"] = d["id"]
                        st.rerun()
            if confirm_del_ds_id is not None:
                st.warning(f"Confirme la eliminacion del diagnostico social ID {confirm_del_ds_id} "
                           "pulsando 'Confirmar' en su fila. Esta accion no puede deshacerse.")
                if st.button("Cancelar eliminacion", key="ds_cancel_del"):
                    st.session_state.pop("ds_confirm_del_id", None)
                    st.rerun()
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
                    num = ficha_t.split("-")[-1] if ficha_t else ""
                    raw = det.get(f"ds{num}_data_v3", "") or ""
                    form_det = {}
                    if raw:
                        try:
                            form_det = json.loads(raw)
                        except (json.JSONDecodeError, TypeError):
                            form_det = {}
                    _ds_render_detalle(ficha_t, form_det)

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
                    "Bloque": r["codigo"],
                    "Total Fichas": r.get("total_fichas", "") or "",
                    "Fichas Completadas": r.get("fichas_completadas", "") or "",
                } for r in resumen_ds]), use_container_width=True, hide_index=True)

    # ── IMPORTAR DESDE EXCEL ───────────────────────────────────────────
    with tab_excel:
        st.markdown("### Plantilla Excel de Diagnostico Social (V4 validada)")
        st.caption("Descargue la plantilla oficial con celdas de validacion (desplegables) "
                   "para el llenado en campo, o suba un archivo llenado para autocompletar.")
        st.markdown("---")
        st.markdown("**1. Descargar Plantilla Oficial para Tecnicos**")
        if st.button("Generar Plantilla Excel V4", type="secondary", key="ds_gen_plantilla"):
            try:
                bloques_data_ds = [(b[1], b[2], b[4], b[5]) for b in BLOQUES_128]
                st.session_state["ds_plantilla_bytes"] = generar_plantilla_ds(None, bloques_data_ds)
                st.success("Plantilla V4 generada correctamente.")
            except Exception as e:
                st.error(f"No se pudo generar la plantilla: {e}")
        if st.session_state.get("ds_plantilla_bytes"):
            st.download_button(
                "Descargar Plantilla Excel V4",
                st.session_state["ds_plantilla_bytes"],
                file_name="Plantilla_Diagnostico_Social_IN_Piura_V4.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="ds_dl_plantilla")

        st.markdown("---")
        st.markdown("**2. Subir Excel Llenado por el Tecnico**")
        st.info("El sistema leera los datos disponibles y autocompletara el formulario "
                "en la pestana **Registro** para su revision antes de guardar.")
        uploaded_excel = st.file_uploader("Seleccionar archivo Excel (.xlsx)", type=["xlsx"], key="ds_excel_upload")
        if uploaded_excel is not None:
            try:
                resultados = parsear_excel_ds(uploaded_excel)
                if not resultados:
                    st.error("No se detectaron fichas F-DS en el archivo. Verifique el formato.")
                else:
                    st.success(f"Se detectaron {len(resultados)} ficha(s) en el archivo.")
                    for i, res in enumerate(resultados):
                        ficha_det = res.get("ficha", "")
                        datos_res = res.get("datos", {})
                        with st.expander(f"Vista previa: {ficha_det}", expanded=True):
                            cols_prev = st.columns(4)
                            cols_prev[0].markdown(f"**Fecha:** {datos_res.get('fecha', '-')}")
                            cols_prev[1].markdown(f"**Responsable:** {datos_res.get('evaluador', '-')}")
                            cols_prev[2].markdown(f"**Bloque:** {datos_res.get('codigo_bloque', '-')}")
                            cols_prev[3].markdown(f"**Distrito:** {datos_res.get('distrito', '-')}")
                            form_prev = datos_res.get("form", {})
                            n_campos = sum(1 for v in form_prev.values() if v not in ("", None, []))
                            st.markdown(f"**Campos detectados:** {n_campos}")
                        if st.button(f"Autocompletar formulario {ficha_det}", type="primary", key=f"ds_ac_{i}"):
                            pend = mapear_a_session_state(res, bm)
                            st.session_state["_ds_pending_state"] = pend
                            st.session_state["_ds_pending_edit_id"] = None
                            st.success(f"Formulario {ficha_det} autocompletado. "
                                       "Vaya a la pestana **Registro** para revisar y guardar.")
                            st.rerun()
            except Exception as e:
                st.error(f"Error al leer el archivo Excel: {e}")




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

    bloques_map = {b["codigo"]: b["id"] for b in bloques}
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

        st.markdown("### Historial de Fichas de Elementos Expuestos")
        st.caption("Haga clic en **Editar** para modificar una ficha existente o en **Eliminar** para descartarla.")
        col_widths_ee = [0.4, 0.9, 0.7, 0.8, 1.0, 0.5, 0.6]
        header_cols = st.columns(col_widths_ee)
        for col, h in zip(header_cols, ["ID", "Bloque", "Ficha", "Fecha", "Responsable", "", ""]):
            col.markdown(f"**{h}**")
        st.markdown("---")
        confirm_del_ee_id = st.session_state.get("ee_confirm_del_id")
        for r in registros_ee:
            row = st.columns(col_widths_ee)
            row[0].write(r.get("id", ""))
            row[1].write(r.get("bloque_codigo", "") or "")
            row[2].write(r.get("ficha", "") or "")
            row[3].write(r.get("fecha_campo", "") or "")
            row[4].write(r.get("responsable_brigada", "") or "")
            if row[5].button("Editar", key=f"edit_ee_row_{r['id']}", type="primary"):
                st.session_state["ee_edit_id"] = r["id"]
                st.session_state["ee_ficha_sel"] = r.get("ficha", "")
                st.session_state.pop("ee_confirm_del_id", None)
                st.rerun()
            if confirm_del_ee_id == r["id"]:
                if row[6].button("Confirmar", key=f"del_confirm_ee_{r['id']}", type="primary"):
                    db.eliminar_elementos_expuestos(r["id"])
                    _invalidar_cache()
                    st.session_state.pop("ee_confirm_del_id", None)
                    st.success(f"Ficha de Elementos Expuestos ID {r['id']} eliminada.")
                    st.rerun()
            else:
                if row[6].button("Eliminar", key=f"del_ee_row_{r['id']}", type="secondary"):
                    st.session_state["ee_confirm_del_id"] = r["id"]
                    st.rerun()
        if confirm_del_ee_id is not None:
            st.warning(f"Confirme la eliminacion de la ficha ID {confirm_del_ee_id} pulsando 'Confirmar' en su fila. "
                       "Esta accion no puede deshacerse.")
            if st.button("Cancelar eliminacion", key="ee_cancel_del"):
                st.session_state.pop("ee_confirm_del_id", None)
                st.rerun()
        st.markdown("---")

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
        def _fmt_money(v):
            try:
                return f"S/ {float(v):,.2f}" if v not in (None, "", 0, 0.0) else ""
            except (TypeError, ValueError):
                return ""
        def _fmt_pct(num, den):
            try:
                num_f = float(num or 0); den_f = float(den or 0)
                if den_f <= 0: return ""
                return f"{(num_f/den_f*100):.1f}%"
            except (TypeError, ValueError):
                return ""
        def _fmt_count(v):
            try:
                iv = int(v); return iv if iv else ""
            except (TypeError, ValueError):
                return ""
        st.dataframe(pd.DataFrame([{
            "Codigo": r["codigo"],
            "Planificado": _fmt_money(r.get("total_planificado")),
            "Ejecutado": _fmt_money(r.get("total_ejecutado")),
            "%Ejec": _fmt_pct(r.get("total_ejecutado"), r.get("total_planificado")),
            "Partidas": _fmt_count(r.get("num_partidas")),
        } for r in rp]),
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
                    # Fallback: si las coords en BD son 0, usar las de BLOQUES_128_MAP
                    if utm_e == 0 or utm_n == 0:
                        datos_b79 = BLOQUES_128_MAP.get(b.get("codigo", ""), {})
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
# MIGRACION V5 — PAGINA TEMPORAL
# ══════════════════════════════════════════════════════════════════════════
def pagina_migracion_v4():
    st.subheader("⚙️ Migracion: Reemplazar Bloques por V5")
    st.warning(
        f"**ATENCION:** Esta accion eliminara TODOS los bloques actuales de la base de datos "
        f"(incluyendo inspecciones, diagnosticos y registros vinculados) "
        f"e insertara los {len(BLOQUES_V5)} bloques del Proyecto IN Piura V5 "
        f"(con zonas Z01..Z14 y coordenadas UTM).",
        icon="⚠️",
    )

    confirmar = st.checkbox("Entiendo que se borraran todos los datos existentes y quiero continuar")

    if not confirmar:
        st.info("Marca la casilla de confirmacion para habilitar la migracion.")
        return

    if st.button("🚀 Ejecutar Migracion V5", type="primary"):
        from datetime import datetime
        import database as db

        conn = None
        try:
            # Toda la migracion corre en UNA sola transaccion: si algun
            # INSERT falla, el TRUNCATE tambien hace rollback y la BD
            # queda intacta. Evita estados parciales y duplicados que
            # ocurrian al usar varias conexiones (DELETE + N inserts).
            conn = db.get_connection()
            cursor = conn.cursor()

            cursor.execute("SELECT COUNT(*) FROM bloques")
            row = cursor.fetchone()
            eliminados = list(row.values())[0] if hasattr(row, 'values') else row[0]

            # TRUNCATE ... CASCADE elimina bloques y todos los registros
            # dependientes (inspecciones, diagnosticos, etc.) en bloque,
            # mas rapido y atomico que DELETE.
            conn.execute("TRUNCATE TABLE bloques RESTART IDENTITY CASCADE")

            fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            insertados = 0
            for b in BLOQUES_V5:
                codigo      = b[1]
                microcuenca = b[2]
                area_ha     = b[3]
                provincia   = b[4]
                distrito    = b[5]
                utm_este    = b[8]
                utm_norte   = b[9]
                cursor.execute("""
                    INSERT INTO bloques (codigo, tipo_intervencion, cuenca, distrito,
                                         utm_este, utm_norte, utm_zona, altitud,
                                         area_hectareas, responsable, estado,
                                         microcuenca, provincia, fecha_registro)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    codigo, "Restauracion", microcuenca, distrito,
                    float(utm_este or 0.0), float(utm_norte or 0.0), "17S", 0.0,
                    float(area_ha or 0.0), "", "Pendiente",
                    microcuenca, provincia, fecha,
                ))
                insertados += 1

            conn.commit()

            st.success(
                f"✅ Migracion V5 completada: {eliminados} bloques eliminados, "
                f"{insertados} bloques V5 insertados."
            )
            st.info("Recarga la pagina o navega al Panel de Control para ver los cambios.")
            st.cache_data.clear()

        except Exception as e:
            if conn is not None:
                try:
                    conn._conn.rollback()
                except Exception:
                    pass
            st.error(f"Error durante la migracion: {e}")
        finally:
            if conn is not None:
                try:
                    conn.close()
                except Exception:
                    pass


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
