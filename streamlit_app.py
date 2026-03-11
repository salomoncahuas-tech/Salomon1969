"""
IN Piura - Plan de Ingreso / Verificacion de Campo
Aplicacion web con Streamlit.
Restauracion de ecosistemas - Cuenca alta del rio Piura, Peru.
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import os
import uuid
import io
import csv
import json
import tempfile

import database as db
import reports
from georeferenciacion import utm_a_latlon, latlon_a_utm
from odk_kobo import generar_xlsform, importar_csv_odk, importar_desde_kobo, KoBoClient
from excel_diagnostico_social import generar_plantilla_ds, parsear_excel_ds, mapear_a_session_state

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

# ── Datos de Origen: 79 Bloques Preliminares de Intervencion ─────────
# Fuente: Bloques_V3_Aplicativo.xlsx
# Cada entrada: (N, Bloque, Microcuenca, Area_ha, Provincia, Distrito, Accesibilidad, Dia)
BLOQUES_79 = [
    (1, "M10B1", "C1096-Q9547", 416.522, "Morropon", "Chulucanas", 0, 0),
    (2, "M10B4", "C1096-Q9547", 365.749, "Morropon", "Chulucanas", 0, 0),
    (3, "M11B3", "C1081-Q9583", 106.059, "Morropon", "San Juan de Bigote", 0, 0),
    (4, "M12B1", "C1076-Q9588", 590.622, "Huancabamba", "Huarmaca", 0, 0),
    (5, "M12B2", "C1076-Q9588", 55.463, "Huancabamba", "Huarmaca", 0, 0),
    (6, "M12B3", "C1076-Q9588", 346.064, "Huancabamba", "Huarmaca", 0, 0),
    (7, "M12B4", "C1076-Q9588", 207.944, "Huancabamba", "Huarmaca", 0, 0),
    (8, "M12B8", "C1076-Q9588", 473.542, "Morropon", "Salitral", 0, 0),
    (9, "M15B1", "C1076-Q9592", 135.22, "Huancabamba", "Huarmaca", 0, 0),
    (10, "M15B2", "C1076-Q9592", 101.409, "Huancabamba", "Huarmaca", 0, 0),
    (11, "M15B5", "C1076-Q9592", 135.664, "Huancabamba", "Huarmaca", 0, 0),
    (12, "M16B2", "C1080-Q9560", 277.582, "Morropon", "Morropon", 0, 0),
    (13, "M16B3", "C1080-Q9560", 106.81, "Morropon", "Morropon", 0, 0),
    (14, "M17B1", "C1096-Q9556", 11.448, "Morropon", "Chulucanas", 0, 0),
    (15, "M17B10", "C1096-Q9556", 924.365, "Ayabaca", "Frias", 0, 0),
    (16, "M17B11", "C1096-Q9556", 819.771, "Morropon", "Chulucanas", 0, 0),
    (17, "M17B4", "C1096-Q9556", 425.034, "Morropon", "Chulucanas", 0, 0),
    (18, "M17B5", "C1096-Q9556", 74.065, "Ayabaca", "Frias", 0, 0),
    (19, "M17B6", "C1096-Q9556", 106.593, "Ayabaca", "Frias", 0, 0),
    (20, "M17B7", "C1096-Q9556", 258.969, "Ayabaca", "Frias", 0, 0),
    (21, "M18B1", "C1077-Q9579", 160.43, "Morropon", "Buenos Aires", 0, 0),
    (22, "M18B3", "C1077-Q9579", 373.96, "Morropon", "Salitral", 0, 0),
    (23, "M18B5", "C1077-Q9579", 197.437, "Morropon", "Salitral", 0, 0),
    (24, "M19B2", "C1086-Q9570", 73.999, "Morropon", "Morropon", 0, 0),
    (25, "M19B5", "C1086-Q9570", 33.614, "Morropon", "Santo Domingo", 0, 0),
    (26, "M19B7", "C1086-Q9570", 68.867, "Morropon", "Santo Domingo", 0, 0),
    (27, "M1B1", "C1077-Q9580", 426.891, "Morropon", "Buenos Aires", 0, 0),
    (28, "M20B1", "C1076-Q9585", 279.766, "Huancabamba", "San Miguel de El Faique", 0, 0),
    (29, "M22B1", "C1076-Q9586", 61.447, "Huancabamba", "San Miguel de El Faique", 0, 0),
    (30, "M25B1", "C1107-Q9541", 313.007, "Piura", "Las Lomas", 0, 0),
    (31, "M26B4", "C1076-Q9587", 36.946, "Huancabamba", "Canchaque", 0, 0),
    (32, "M27B1", "C1096-Q9557", 449.066, "Morropon", "Chulucanas", 0, 0),
    (33, "M27B4", "C1096-Q9557", 588.839, "Morropon", "Morropon", 0, 0),
    (34, "M27B5", "C1096-Q9557", 97.792, "Morropon", "Morropon", 0, 0),
    (35, "M28B1", "C1086-Q9575", 245.908, "Morropon", "Yamango", 0, 0),
    (36, "M28B2", "C1086-Q9575", 90.228, "Morropon", "Yamango", 0, 0),
    (37, "M28B3", "C1086-Q9575", 178.072, "Morropon", "Yamango", 0, 0),
    (38, "M28B4", "C1086-Q9575", 283.543, "Morropon", "Yamango", 0, 0),
    (39, "M29B1", "C1096-Q9564", 27.184, "Ayabaca", "Frias", 0, 0),
    (40, "M2B1", "C1076-Q9584", 81.823, "Morropon", "Salitral", 0, 0),
    (41, "M2B5", "C1076-Q9584", 30.918, "Morropon", "Salitral", 0, 0),
    (42, "M2B8", "C1076-Q9584", 116.136, "Huancabamba", "San Miguel de El Faique", 0, 0),
    (43, "M30B1", "C1081-Q9591", 664.828, "Huancabamba", "Canchaque", 0, 0),
    (44, "M30B5", "C1081-Q9591", 107.389, "Huancabamba", "Huancabamba", 0, 0),
    (45, "M30B6", "C1081-Q9591", 80.218, "Huancabamba", "Huancabamba", 0, 0),
    (46, "M32B1", "C1086-Q9569", 434.48, "Morropon", "Buenos Aires", 0, 0),
    (47, "M32B2", "C1086-Q9569", 82.209, "Morropon", "Santa Catalina de Mossa", 0, 0),
    (48, "M32B3", "C1086-Q9569", 98.628, "Morropon", "Morropon", 0, 0),
    (49, "M34B1", "C1078-Q9562", 738.048, "Morropon", "La Matanza", 0, 0),
    (50, "M35B1", "C1107-Q9539", 244.529, "Sullana", "Sullana", 0, 0),
    (51, "M36B1", "C1086-Q9576", 137.132, "Morropon", "Chalaco", 0, 0),
    (52, "M36B2", "C1086-Q9576", 57.349, "Morropon", "Santa Catalina de Mossa", 0, 0),
    (53, "M3B1", "C1081-Q9582", 81.014, "Morropon", "Salitral", 0, 0),
    (54, "M3B3", "C1081-Q9582", 84.855, "Morropon", "San Juan de Bigote", 0, 0),
    (55, "M3B5", "C1081-Q9582", 52.761, "Morropon", "San Juan de Bigote", 0, 0),
    (56, "M3B6", "C1081-Q9582", 60.372, "Morropon", "San Juan de Bigote", 0, 0),
    (57, "M3B7", "C1081-Q9582", 122.987, "Morropon", "San Juan de Bigote", 0, 0),
    (58, "M3B8", "C1081-Q9582", 565.792, "Morropon", "San Juan de Bigote", 0, 0),
    (59, "M3B9", "C1081-Q9582", 294.318, "Morropon", "Salitral", 0, 0),
    (60, "M4B1", "C1076-Q9589", 37.034, "Huancabamba", "Huarmaca", 0, 0),
    (61, "M4B3", "C1076-Q9589", 234.137, "Huancabamba", "Huarmaca", 0, 0),
    (62, "M4B4", "C1076-Q9589", 290.0, "Huancabamba", "Huarmaca", 0, 0),
    (63, "M5B1", "C1108-Q9552", 44.363, "Piura", "Tambo Grande", 0, 0),
    (64, "M5B3", "C1108-Q9552", 333.968, "Piura", "Tambo Grande", 0, 0),
    (65, "M6B10", "C1077-Q9566", 409.666, "Morropon", "Buenos Aires", 0, 0),
    (66, "M6B2", "C1077-Q9566", 2232.934, "Morropon", "Buenos Aires", 0, 0),
    (67, "M6B5", "C1077-Q9566", 228.134, "Morropon", "Buenos Aires", 0, 0),
    (68, "M6B6", "C1077-Q9566", 91.615, "Morropon", "Buenos Aires", 0, 0),
    (69, "M6B7", "C1077-Q9566", 194.339, "Morropon", "Buenos Aires", 0, 0),
    (70, "M6B8", "C1077-Q9566", 62.066, "Morropon", "Buenos Aires", 0, 0),
    (71, "M7B1", "C1076-Q9581", 130.475, "Morropon", "Salitral", 0, 0),
    (72, "M7B2", "C1076-Q9581", 115.422, "Morropon", "Salitral", 0, 0),
    (73, "M7B3", "C1076-Q9581", 65.768, "Morropon", "Salitral", 0, 0),
    (74, "M7B6", "C1076-Q9581", 91.317, "Morropon", "Salitral", 0, 0),
    (75, "M8B1", "C1096-Q9558", 160.448, "Morropon", "Morropon", 0, 0),
    (76, "M8B2", "C1096-Q9558", 56.747, "Morropon", "Santo Domingo", 0, 0),
    (77, "M8B4", "C1096-Q9558", 58.489, "Morropon", "Chulucanas", 0, 0),
    (78, "M8B5", "C1096-Q9558", 93.213, "Morropon", "Chulucanas", 0, 0),
    (79, "M9B1", "C1096-Q9545", 1116.558, "Morropon", "Chulucanas", 0, 0),
]


# Diccionario para busqueda rapida por codigo de bloque
BLOQUES_79_MAP = {b[1]: {"n": b[0], "codigo": b[1], "microcuenca": b[2],
    "area_ha": b[3], "provincia": b[4], "distrito": b[5],
    "accesibilidad": b[6], "dia_evaluacion": b[7]} for b in BLOQUES_79}

# Lista de codigos para el dropdown (solo codigo de bloque)
BLOQUES_79_OPCIONES = [b[1] for b in BLOQUES_79]
PROVINCIAS_DISTRITOS = {
    "Ayabaca": ["Frias"],
    "Huancabamba": ["Canchaque","Huancabamba","Huarmaca","San Miguel de El Faique"],
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
# Fuente: Formatos_Sociales_Registros_de_Campo_IN_Piura_2026.xlsx
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
</style>""", unsafe_allow_html=True)

st.markdown("""<div class="main-header">
<h1>\U0001F331 IN Piura</h1>
<p>Plan de Ingreso | Verificacion de Campo | Cuenca Alta del Rio Piura</p>
</div>""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────
pagina = st.sidebar.selectbox("Navegacion", [
    "Panel de Control","Bloques de Intervencion","Inspeccion de Campo",
    "Indicadores de Calidad","Diagnostico Territorial","Diagnostico Social",
    "Presupuesto","Cronograma",
    "Georreferenciacion","ODK / KoBoToolbox","Reportes",
    "Conversor PDF -> Excel",
])
st.sidebar.markdown("---")
st.sidebar.markdown("**IN Piura** v2.0 Web\n\nRestauracion de Ecosistemas\nCuenca Alta del Rio Piura")

# ── Helpers ───────────────────────────────────────────────────────────────
def _bloques_map():
    return {f"{b['codigo']} - {b['tipo_intervencion']}": b["id"] for b in db.obtener_bloques()}

def _distritos(prov):
    return PROVINCIAS_DISTRITOS.get(prov, DISTRITOS_PIURA) if prov else DISTRITOS_PIURA

def _resolver_microcuenca(bloque_label):
    """Resuelve la microcuenca para un bloque dado su label 'CODIGO - TIPO'.
    Busca primero en la BD y luego en BLOQUES_79_MAP como fallback."""
    codigo = bloque_label.split(" - ")[0].strip() if " - " in bloque_label else bloque_label.strip()
    # Buscar en BD
    for b in db.obtener_bloques():
        if b["codigo"] == codigo:
            mc = b.get("microcuenca", "") or ""
            if mc and mc in MICROCUENCAS:
                return mc
            break
    # Fallback: buscar en los 79 bloques predefinidos
    datos = BLOQUES_79_MAP.get(codigo, {})
    mc = datos.get("microcuenca", "")
    if mc and mc in MICROCUENCAS:
        return mc
    return ""

# ══════════════════════════════════════════════════════════════════════════
# PANEL DE CONTROL
# ══════════════════════════════════════════════════════════════════════════
def pagina_dashboard():
    st.subheader("Panel de Control - Resumen Ejecutivo")
    stats = db.obtener_estadisticas_generales()
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
    res = db.obtener_resumen_bloques()
    if res:
        st.dataframe(pd.DataFrame([{"Codigo":b["codigo"],"Tipo":b["tipo_intervencion"],
            "Distrito":b["distrito"],"Area (ha)":f"{b['area_hectareas']:.4f}",
            "Estado":b["estado"],"Avance %":f"{(b.get('ultimo_avance') or 0):.1f}",
            "Inspecciones":b.get("total_inspecciones",0)} for b in res]),
            use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════
# BLOQUES DE INTERVENCION
# ══════════════════════════════════════════════════════════════════════════
def _extraer_codigo_bloque_79(opcion):
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
            st.markdown("**Editar Bloque**")
            st.info(f"Editando bloque ID {edit_id}. Modifique los campos y presione Actualizar.")
            if st.button("Cancelar edicion", key="bl_cancel_edit"):
                st.session_state["bl_edit_id"] = None
                for k in list(st.session_state.keys()):
                    if k.startswith("bl_") and k != "bl_edit_id":
                        del st.session_state[k]
                st.rerun()
        else:
            st.markdown("**Registro de Bloque**")

        # Selector rapido de los 79 bloques preliminares (solo en modo nuevo)
        if not edit_id:
            st.markdown("##### Seleccion rapida - 79 Bloques Preliminares")
            sel_79 = st.selectbox(
                "Seleccionar bloque predefinido",
                ["(Seleccionar bloque predefinido)"] + BLOQUES_79_OPCIONES,
                key="sel_bloque_79",
                help="Seleccione un bloque de la lista de 79 bloques preliminares para autocompletar los campos"
            )

            # Determinar valores por defecto segun seleccion
            cod_sel = _extraer_codigo_bloque_79(sel_79 if sel_79 != "(Seleccionar bloque predefinido)" else "")
            datos_79 = BLOQUES_79_MAP.get(cod_sel, {}) if cod_sel else {}

            def_codigo = datos_79.get("codigo", "")
            def_microcuenca = datos_79.get("microcuenca", "")
            def_area = str(datos_79.get("area_ha", "0"))
            def_provincia = datos_79.get("provincia", "")
            def_distrito = datos_79.get("distrito", "")
            def_accesibilidad = datos_79.get("accesibilidad", 0)
            def_dia = datos_79.get("dia_evaluacion", 0)

            if datos_79:
                acc_txt = "Acceso limitado" if def_accesibilidad == 1 else "Acceso normal"
                st.info(f"Bloque **{def_codigo}** | Microcuenca: {def_microcuenca} | "
                        f"{def_distrito} ({def_provincia}) | {def_area} ha | "
                        f"{acc_txt} | Dia eval.: {def_dia}")
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
            ue = a1.text_input("UTM Este", st.session_state.get("bl_utm_este", "0") if edit_id else "0")
            un = a2.text_input("UTM Norte", st.session_state.get("bl_utm_norte", "0") if edit_id else "0")
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
                        st.success(f"Bloque {codigo} registrado.")
                    st.rerun()
                except Exception as e: st.error(f"Error: {e}")
    with ct:
        st.markdown("**Bloques Registrados**")
        st.caption("Haga clic en **Editar** para modificar un bloque existente y evitar duplicidades.")
        busq = st.text_input("Buscar","",key="busq_bl")
        bloques = db.buscar_bloques(busq) if busq else db.obtener_bloques()
        if bloques:
            # Tabla con botones de edicion por fila
            header_cols = st.columns([0.5, 1.2, 1.2, 1.2, 1, 1, 0.8, 0.8, 0.7])
            headers = ["ID", "Codigo", "Microcuenca", "Tipo", "Provincia", "Distrito", "Area", "Estado", ""]
            for col, h in zip(header_cols, headers):
                col.markdown(f"**{h}**")
            st.markdown("---")
            for b in bloques:
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
            if sel and sel in bm and st.button("Eliminar bloque"):
                db.eliminar_bloque(bm[sel]); st.success("Eliminado."); st.rerun()
        else: st.info("Sin bloques.")

        st.markdown("---")
        with st.expander("Tabla de Referencia - 79 Bloques Preliminares de Intervencion", expanded=False):
            st.caption("Fuente: Bloques_Preliminares_V3.xlsx - Base de datos completa del proyecto IN Piura")
            df_79 = pd.DataFrame([{
                "N":b[0], "Bloque":b[1], "Microcuenca":b[2],
                "Area (ha)":b[3], "Provincia":b[4], "Distrito":b[5],
                "Accesibilidad":"Limitado" if b[6]==1 else "Normal",
                "Dia Eval.":b[7],
            } for b in BLOQUES_79])
            st.dataframe(df_79, use_container_width=True, hide_index=True, height=400)

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
        st.info(f"Editando inspeccion ID {edit_id}. Modifique los campos y presione Actualizar.")
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
        fecha = st.date_input("Fecha de visita", value=def_fecha)
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
                        msg = "Inspeccion registrada."
                        if rutas_pdf:
                            msg += f" {len(rutas_pdf)} PDF(s) adjuntado(s)."
                        st.success(msg); st.rerun()
            except Exception as e: st.error(f"Error: {e}")
    st.markdown("---")
    st.markdown("**Historial de Inspecciones**")
    st.caption("Haga clic en **Editar** para modificar una inspeccion existente y evitar duplicidades.")
    insp = db.obtener_todas_inspecciones()
    if insp:
        # Tabla con botones de edicion
        header_cols = st.columns([0.5, 1, 1, 0.9, 0.9, 0.7, 1, 0.5, 0.6])
        for col, h in zip(header_cols, ["ID", "Bloque", "Microcuenca", "Fecha", "Inspector", "Avance%", "Verificacion", "PDFs", ""]):
            col.markdown(f"**{h}**")
        st.markdown("---")
        for i in insp:
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
            db.eliminar_inspeccion(insp_del_map[sel_del]); st.success("Inspeccion eliminada."); st.rerun()

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
        st.info(f"Editando indicador ID {edit_id}. Modifique los campos y presione Actualizar.")
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
            db.eliminar_indicadores(ind_del[sel_del]); st.success("Indicador eliminado."); st.rerun()

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

    tab_reg, tab_hist = st.tabs(["Registro de Diagnostico", "Historial / Consulta"])

    with tab_reg:
        if dt_edit_id:
            st.info(f"Editando diagnostico ID {dt_edit_id}. Modifique los campos y presione Actualizar.")
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
        fecha_ev = r2.date_input("Fecha de evaluacion", value=def_fecha_dt, key="dt_fecha")
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
                            st.warning(f"Ya existe un diagnostico para este bloque en {fecha_ev.strftime('%Y-%m-%d')} "
                                       f"por {evaluador}. Use 'Editar' en Historial para modificarlo.")
                        else:
                            db.insertar_diagnostico_territorial(bloque_id=bid, **_dt_kwargs)
                            st.success(f"Diagnostico territorial guardado ({', '.join(fichas_sel)}).")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

    with tab_hist:
        st.markdown("### Historial de Diagnosticos Territoriales")
        st.caption("Haga clic en **Editar** para modificar un diagnostico existente y evitar duplicidades.")
        todos_dt = db.obtener_todos_diagnosticos()
        if not todos_dt:
            st.info("No hay diagnosticos registrados.")
        else:
            # Tabla con botones de edicion
            header_cols = st.columns([0.4, 1, 0.8, 0.8, 0.8, 0.8, 0.8, 0.5])
            for col, h in zip(header_cols, ["ID", "Bloque", "Fichas", "Fecha", "Evaluador", "Microcuenca", "Distrito", ""]):
                col.markdown(f"**{h}**")
            st.markdown("---")
            for d in todos_dt:
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


# ══════════════════════════════════════════════════════════════════════════
# DIAGNOSTICO SOCIAL
# ══════════════════════════════════════════════════════════════════════════
def _ds_datos_generales():
    """Campos de datos generales compartidos por todas las fichas DS."""
    c1, c2, c3, c4 = st.columns(4)
    prov = c1.text_input("Provincia", key="ds_prov")
    dist = c2.text_input("Distrito", key="ds_dist")
    cpob = c3.text_input("Centro Poblado / Localidad", key="ds_cpob")
    ccam = c4.text_input("Comunidad Campesina", key="ds_ccam")
    c5, c6, c7, c8 = st.columns(4)
    este = c5.number_input("Coordenada Este (UTM)", value=0.0, format="%.1f", key="ds_este")
    norte = c6.number_input("Coordenada Norte (UTM)", value=0.0, format="%.1f", key="ds_norte")
    alt_v = c7.number_input("Altitud (msnm)", value=0.0, format="%.0f", key="ds_alt")
    ubigeo = c8.text_input("Codigo UBIGEO", key="ds_ubigeo")
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
            st.info(f"Editando registro ID {edit_id}. Modifique los campos necesarios y presione Guardar.")
            if st.button("Cancelar edicion (nuevo registro)", key="ds_cancel_edit"):
                st.session_state["ds_edit_id"] = None
                st.rerun()

        # ── Datos comunes ─────────────────────────────────────────────
        bl = st.selectbox("Bloque de Intervencion", list(bm.keys()), key="ds_bl")
        bid = bm[bl]

        # Auto-resolver microcuenca del bloque seleccionado
        mc_auto_ds = _resolver_microcuenca(bl)
        if mc_auto_ds:
            mc_idx_ds = MICROCUENCAS.index(mc_auto_ds) + 1  # +1 por el "" inicial
            st.info(f"Microcuenca vinculada automaticamente: **{mc_auto_ds}**")
        else:
            mc_idx_ds = 0

        r1, r2, r3, r4 = st.columns(4)
        mc = r1.selectbox("Microcuenca", [""] + MICROCUENCAS, index=mc_idx_ds, key="ds_mc")
        fecha_ev = r2.date_input("Fecha", value=datetime.now(), key="ds_fecha")
        evaluador = r3.text_input("Responsable", key="ds_eval")
        ficha_num = r4.text_input("Ficha N", key="ds_fnum")
        dg = _ds_datos_generales()
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
                        st.session_state["ds_edit_id"] = None
                        st.success(f"Ficha {ficha_sel} actualizada correctamente (ID {edit_id}).")
                    else:
                        if "archivos_adjuntos" not in reg:
                            reg["archivos_adjuntos"] = ""
                        db.insertar_diagnostico_social(reg)
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
        todos_ds = db.obtener_todos_diagnosticos_sociales()
        if not todos_ds:
            st.info("No hay diagnosticos sociales registrados.")
        else:
            # Tabla con botones de edicion
            header_cols = st.columns([0.4, 0.9, 0.7, 0.8, 0.8, 0.8, 0.8, 0.5])
            for col, h in zip(header_cols, ["ID", "Bloque", "Ficha", "Fecha", "Responsable", "C.Poblado", "Distrito", ""]):
                col.markdown(f"**{h}**")
            st.markdown("---")
            for d in todos_ds:
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
                plantilla_bytes = generar_plantilla_ds(fichas_descarga)
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
                   "(`Formatos_Sociales_Registros_de_Campo_IN_Piura_2026.xlsx`) "
                   "llenada por el tecnico.")


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
        st.info(f"Editando partida ID {edit_id}. Modifique los campos y presione Actualizar.")
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
            db.eliminar_presupuesto(pm[sp]); st.success("Eliminada."); st.rerun()
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
        st.info(f"Editando actividad ID {edit_id}. Modifique los campos y presione Actualizar.")
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
    def_ir = ""
    def_fr = ""
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
        def_ir = st.session_state.get("crono_e_ir", "") or ""
        def_fr = st.session_state.get("crono_e_fr", "") or ""
        def_av = float(st.session_state.get("crono_e_av", 0))
        def_ea = st.session_state.get("crono_e_ea", "")
        if def_ea and def_ea in ESTADOS_ACTIVIDAD:
            def_ea_idx = ESTADOS_ACTIVIDAD.index(def_ea)
        def_re = st.session_state.get("crono_e_re", "")
        def_ob = st.session_state.get("crono_e_ob", "")

    with st.form("form_crono", clear_on_submit=not edit_id):
        act = st.selectbox("Actividad", ACTIVIDADES_TIPO, index=def_act_idx)
        x1,x2 = st.columns(2)
        ip = x1.date_input("Inicio plan.", value=def_ip)
        fp = x2.date_input("Fin plan.", value=def_fp)
        x3,x4 = st.columns(2)
        ir = x3.text_input("Inicio real", def_ir); fr = x4.text_input("Fin real", def_fr)
        x5,x6 = st.columns(2)
        av = x5.number_input("Avance %", 0.0, 100.0, def_av)
        ea = x6.selectbox("Estado", ESTADOS_ACTIVIDAD, index=def_ea_idx)
        re = st.text_input("Responsable", value=def_re); ob = st.text_area("Observaciones", value=def_ob)
        btn_label = "Actualizar Actividad" if edit_id else "Guardar Actividad"
        guardar = st.form_submit_button(btn_label, type="primary")
    if guardar:
        try:
            if edit_id:
                db.actualizar_actividad(actividad_id=edit_id, actividad=act,
                    fecha_inicio_plan=ip.strftime("%Y-%m-%d"), fecha_fin_plan=fp.strftime("%Y-%m-%d"),
                    fecha_inicio_real=ir, fecha_fin_real=fr, porcentaje_avance=av,
                    responsable=re, observaciones=ob, estado=ea)
                st.session_state["crono_edit_id"] = None
                for k in list(st.session_state.keys()):
                    if k.startswith("crono_e_"):
                        del st.session_state[k]
                st.success("Actividad actualizada."); st.rerun()
            else:
                db.insertar_actividad(bloque_id=bid, actividad=act,
                    fecha_inicio_plan=ip.strftime("%Y-%m-%d"), fecha_fin_plan=fp.strftime("%Y-%m-%d"),
                    fecha_inicio_real=ir, fecha_fin_real=fr, porcentaje_avance=av,
                    responsable=re, observaciones=ob, estado=ea)
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
            db.eliminar_actividad(am[sa]); st.success("Eliminada."); st.rerun()
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
                    zn = int(b["utm_zona"].replace("S","").replace("N",""))
                    he = "S" if "S" in b["utm_zona"] else "N"
                    la,lo = utm_a_latlon(b["utm_este"],b["utm_norte"],zn,he)
                    co = COLORES_ESTADO.get(b.get("estado",""),[149,165,166])
                    md.append({"lat":la,"lon":lo,"codigo":b["codigo"],
                        "tipo":b["tipo_intervencion"],"estado":b["estado"],
                        "area":b["area_hectareas"],"r":co[0],"g":co[1],"b":co[2]})
                except (ValueError,KeyError): pass
            if md:
                import pydeck as pdk
                df = pd.DataFrame(md)
                layer = pdk.Layer("ScatterplotLayer",data=df,
                    get_position=["lon","lat"],get_color=["r","g","b",200],
                    get_radius=300,pickable=True)
                vs = pdk.ViewState(latitude=df["lat"].mean(),longitude=df["lon"].mean(),zoom=10)
                st.pydeck_chart(pdk.Deck(layers=[layer],initial_view_state=vs,
                    tooltip={"text":"Codigo: {codigo}\nTipo: {tipo}\nEstado: {estado}\nArea: {area} ha"}))
                st.caption(":red_circle: Pendiente | :orange_circle: En progreso | :green_circle: Verificado")
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
# ROUTER
# ══════════════════════════════════════════════════════════════════════════
if pagina == "Panel de Control": pagina_dashboard()
elif pagina == "Bloques de Intervencion": pagina_bloques()
elif pagina == "Inspeccion de Campo": pagina_inspeccion()
elif pagina == "Indicadores de Calidad": pagina_indicadores()
elif pagina == "Diagnostico Territorial": pagina_diagnostico_territorial()
elif pagina == "Diagnostico Social": pagina_diagnostico_social()
elif pagina == "Presupuesto": pagina_presupuesto()
elif pagina == "Cronograma": pagina_cronograma()
elif pagina == "Georreferenciacion": pagina_georreferenciacion()
elif pagina == "ODK / KoBoToolbox": pagina_odk()
elif pagina == "Reportes": pagina_reportes()
elif pagina == "Conversor PDF -> Excel": pagina_conversor_pdf()
