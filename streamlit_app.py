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
import tempfile

import database as db
import reports
from georeferenciacion import utm_a_latlon, latlon_a_utm
from odk_kobo import generar_xlsform, importar_csv_odk, importar_desde_kobo, KoBoClient

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
    "Terrazas de formacion lenta", "Diques de mamposteria", "Otras",
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

# ── Datos de Origen: 83 Bloques Preliminares de Intervencion ─────────
# Fuente: DATOS DE ORIGEN 83 BLOQUES.xlsx
# Cada entrada: (N, Bloque, Microcuenca, Area_ha, Provincia, Distrito, Accesibilidad, Dia)
BLOQUES_83 = [
    (1, "M25B1", "C1107-Q9541", 313.131, "Piura", "Las Lomas", 0, 1),
    (2, "M35B1", "C1107-Q9539", 229.177, "Sullana", "Sullana", 0, 1),
    (3, "M10B4", "C1096-Q9547", 95.459, "Morropon", "Chulucanas", 0, 2),
    (4, "M9B1", "C1096-Q9545", 1115.665, "Morropon", "Chulucanas", 0, 2),
    (5, "M5B3", "C1108-Q9552", 561.842, "Piura", "Tambo Grande", 0, 2),
    (6, "M5B1", "C1108-Q9552", 58.27, "Piura", "Tambo Grande", 1, 2),
    (7, "M8B5", "C1096-Q9558", 94.918, "Morropon", "Chulucanas", 1, 3),
    (8, "M17B1", "C1096-Q9556", 30.348, "Morropon", "Chulucanas", 0, 3),
    (9, "M27B1", "C1096-Q9557", 289.02, "Morropon", "Chulucanas", 0, 3),
    (10, "M10B1", "C1096-Q9547", 596.766, "Morropon", "Chulucanas", 0, 3),
    (11, "M17B6", "C1096-Q9556", 118.175, "Ayabaca", "Frias", 1, 4),
    (12, "M17B11", "C1096-Q9556", 618.871, "Morropon", "Chulucanas", 1, 4),
    (13, "M17B4", "C1096-Q9556", 159.267, "Morropon", "Chulucanas", 1, 4),
    (14, "M17B10", "C1096-Q9556", 594.76, "Ayabaca", "Frias", 1, 4),
    (15, "M17B5", "C1096-Q9556", 139.603, "Ayabaca", "Frias", 1, 4),
    (16, "M8B4", "C1096-Q9558", 113.764, "Morropon", "Chulucanas", 1, 4),
    (17, "M17B7", "C1096-Q9556", 261.051, "Ayabaca", "Frias", 1, 4),
    (18, "M36B2", "C1086-Q9576", 64.769, "Morropon", "Santa Catalina de Mossa", 0, 5),
    (19, "M29B1", "C1096-Q9564", 28.289, "Ayabaca", "Frias", 1, 5),
    (20, "M19B5", "C1086-Q9570", 79.177, "Morropon", "Santo Domingo", 1, 5),
    (21, "M36B1", "C1086-Q9576", 131.383, "Morropon", "Chalaco", 1, 5),
    (22, "M19B7", "C1086-Q9570", 108.891, "Morropon", "Santo Domingo", 1, 6),
    (23, "M19B2", "C1086-Q9570", 54.187, "Morropon", "Morropon", 1, 6),
    (24, "M8B2", "C1096-Q9558", 67.561, "Morropon", "Santo Domingo", 1, 6),
    (25, "M32B2", "C1086-Q9569", 44.889, "Morropon", "Santa Catalina de Mossa", 1, 6),
    (26, "M32B3", "C1086-Q9569", 102.916, "Morropon", "Morropon", 1, 6),
    (27, "M34B1", "C1078-Q9562", 736.319, "Morropon", "La Matanza", 0, 7),
    (28, "M8B1", "C1096-Q9558", 148.266, "Morropon", "Morropon", 0, 7),
    (29, "M27B4", "C1096-Q9557", 511.53, "Morropon", "Morropon", 0, 7),
    (30, "M16B3", "C1080-Q9560", 110.809, "Morropon", "Morropon", 0, 7),
    (31, "M16B2", "C1080-Q9560", 462.688, "Morropon", "Morropon", 0, 7),
    (32, "M27B5", "C1096-Q9557", 88.152, "Morropon", "Morropon", 0, 7),
    (33, "M6B10", "C1077-Q9566", 376.183, "Morropon", "Buenos Aires", 1, 8),
    (34, "M28B3", "C1086-Q9575", 44.709, "Morropon", "Yamango", 1, 8),
    (35, "M28B1", "C1086-Q9575", 49.608, "Morropon", "Yamango", 0, 8),
    (36, "M28B4", "C1086-Q9575", 160.485, "Morropon", "Yamango", 0, 8),
    (37, "M28B2", "C1086-Q9575", 155.139, "Morropon", "Yamango", 1, 8),
    (38, "M32B1", "C1086-Q9569", 313.334, "Morropon", "Buenos Aires", 1, 8),
    (39, "M6B6", "C1077-Q9566", 80.898, "Morropon", "Buenos Aires", 0, 9),
    (40, "M6B8", "C1077-Q9566", 68.037, "Morropon", "Buenos Aires", 0, 9),
    (41, "M6B7", "C1077-Q9566", 185.483, "Morropon", "Buenos Aires", 0, 9),
    (42, "M6B2", "C1077-Q9566", 2141.763, "Morropon", "Buenos Aires", 0, 9),
    (43, "M1B1", "C1075-Q9580", 289.68, "Morropon", "Buenos Aires", 0, 10),
    (44, "M18B1", "C1077-Q9579", 75.61, "Morropon", "Buenos Aires", 1, 10),
    (45, "M6B5", "C1077-Q9566", 286.591, "Morropon", "Buenos Aires", 0, 10),
    (46, "M18B3", "C1077-Q9579", 236.273, "Morropon", "Salitral", 0, 10),
    (47, "M18B5", "C1077-Q9579", 157.818, "Morropon", "Salitral", 0, 10),
    (48, "M3B9", "C1081-Q9582", 328.824, "Morropon", "Salitral", 1, 11),
    (49, "M3B8", "C1081-Q9582", 567.472, "Morropon", "San Juan de Bigote", 1, 11),
    (50, "M3B1", "C1081-Q9582", 84.927, "Morropon", "Salitral", 0, 11),
    (51, "M7B1", "C1076-Q9581", 144.726, "Morropon", "Salitral", 1, 11),
    (52, "M3B3", "C1081-Q9582", 84.715, "Morropon", "San Juan de Bigote", 1, 11),
    (53, "M11B1", "C1081-Q9583", 48.951, "Morropon", "San Juan de Bigote", 1, 12),
    (54, "M3B5", "C1081-Q9582", 153.642, "Morropon", "San Juan de Bigote", 0, 12),
    (55, "M3B7", "C1081-Q9582", 125.348, "Morropon", "San Juan de Bigote", 1, 12),
    (56, "M11B3", "C1081-Q9583", 95.537, "Morropon", "San Juan de Bigote", 1, 12),
    (57, "M30B6", "C1081-Q9591", 70.778, "Huancabamba", "Huancabamba", 1, 13),
    (58, "M30B1", "C1081-Q9591", 535.022, "Huancabamba", "Canchaque", 0, 13),
    (59, "M30B5", "C1081-Q9591", 89.999, "Huancabamba", "Huancabamba", 0, 13),
    (60, "M7B2", "C1076-Q9581", 154.81, "Morropon", "Salitral", 1, 14),
    (61, "M7B6", "C1076-Q9581", 95.209, "Morropon", "Salitral", 0, 14),
    (62, "M7B3", "C1076-Q9581", 98.029, "Morropon", "Salitral", 1, 14),
    (63, "M20B1", "C1076-Q9585", 31.619, "Huancabamba", "San Miguel de El Faique", 0, 14),
    (64, "M2B8", "C1076-Q9584", 63.364, "Huancabamba", "San Miguel de El Faique", 0, 14),
    (65, "M2B1", "C1076-Q9584", 83.883, "Morropon", "Salitral", 0, 14),
    (66, "M26B4", "C1076-Q9587", 37.933, "Huancabamba", "Canchaque", 0, 15),
    (67, "M11B2", "C1081-Q9583", 48.988, "Huancabamba", "Canchaque", 0, 15),
    (68, "M22B1", "C1076-Q9586", 150.796, "Huancabamba", "San Miguel de El Faique", 1, 15),
    (69, "M12B1", "C1076-Q9588", 262.716, "Huancabamba", "Huarmaca", 0, 16),
    (70, "M4B1", "C1076-Q9589", 40.638, "Huancabamba", "Huarmaca", 0, 16),
    (71, "M12B8", "C1076-Q9588", 235.281, "Morropon", "Salitral", 1, 16),
    (72, "M4B4", "C1076-Q9589", 293.115, "Huancabamba", "Huarmaca", 0, 16),
    (73, "M4B3", "C1076-Q9589", 181.941, "Huancabamba", "Huarmaca", 0, 16),
    (74, "M2B5", "C1076-Q9584", 34.164, "Morropon", "Salitral", 0, 16),
    (75, "M12B2", "C1076-Q9588", 62.825, "Huancabamba", "Huarmaca", 1, 17),
    (76, "M12B6", "C1076-Q9588", 101.408, "Huancabamba", "Huarmaca", 1, 17),
    (77, "M12B7", "C1076-Q9588", 49.703, "Huancabamba", "Huarmaca", 1, 17),
    (78, "M12B3", "C1076-Q9588", 262.721, "Huancabamba", "Huarmaca", 1, 17),
    (79, "M12B4", "C1076-Q9588", 258.994, "Huancabamba", "Huarmaca", 1, 18),
    (80, "M15B2", "C1076-Q9592", 142.732, "Huancabamba", "Huarmaca", 1, 18),
    (81, "M15B1", "C1076-Q9592", 144.28, "Huancabamba", "Huarmaca", 1, 18),
    (82, "M12B5", "C1076-Q9588", 64.986, "Huancabamba", "Huarmaca", 1, 18),
    (83, "M15B5", "C1076-Q9592", 181.872, "Huancabamba", "Huarmaca", 1, 19),
]

# Diccionario para busqueda rapida por codigo de bloque
BLOQUES_83_MAP = {b[1]: {"n": b[0], "codigo": b[1], "microcuenca": b[2],
    "area_ha": b[3], "provincia": b[4], "distrito": b[5],
    "accesibilidad": b[6], "dia_evaluacion": b[7]} for b in BLOQUES_83}

# Lista de codigos para el dropdown (solo codigo de bloque)
BLOQUES_83_OPCIONES = [b[1] for b in BLOQUES_83]
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
    "Indicadores de Calidad","Diagnostico Territorial",
    "Presupuesto","Cronograma",
    "Georreferenciacion","ODK / KoBoToolbox","Reportes",
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
    Busca primero en la BD y luego en BLOQUES_83_MAP como fallback."""
    codigo = bloque_label.split(" - ")[0].strip() if " - " in bloque_label else bloque_label.strip()
    # Buscar en BD
    for b in db.obtener_bloques():
        if b["codigo"] == codigo:
            mc = b.get("microcuenca", "") or ""
            if mc and mc in MICROCUENCAS:
                return mc
            break
    # Fallback: buscar en los 83 bloques predefinidos
    datos = BLOQUES_83_MAP.get(codigo, {})
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
    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.metric("Total Bloques", stats["total_bloques"])
    c2.metric("Area Total", f"{stats['area_total_ha']:.2f} ha")
    c3.metric("Inspecciones", stats["total_inspecciones"])
    c4.metric("Avance Promedio", f"{stats['avance_promedio']:.1f}%")
    c5.metric("Diag. Territoriales", stats.get("total_diagnosticos", 0))
    c6.metric("Personal Activo", stats["personal_activo"])
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
def _extraer_codigo_bloque_83(opcion):
    """Retorna el codigo de bloque directamente (el dropdown solo muestra codigos)."""
    if not opcion:
        return None
    return opcion.strip()

def pagina_bloques():
    st.subheader("Bloques de Intervencion")
    cf,ct = st.columns([1,2])
    with cf:
        st.markdown("**Registro de Bloque**")

        # Selector rapido de los 83 bloques preliminares
        st.markdown("##### Seleccion rapida - 83 Bloques Preliminares")
        sel_83 = st.selectbox(
            "Seleccionar bloque predefinido",
            ["(Seleccionar bloque predefinido)"] + BLOQUES_83_OPCIONES,
            key="sel_bloque_83",
            help="Seleccione un bloque de la lista de 83 bloques preliminares para autocompletar los campos"
        )

        # Determinar valores por defecto segun seleccion
        cod_sel = _extraer_codigo_bloque_83(sel_83 if sel_83 != "(Seleccionar bloque predefinido)" else "")
        datos_83 = BLOQUES_83_MAP.get(cod_sel, {}) if cod_sel else {}

        def_codigo = datos_83.get("codigo", "")
        def_microcuenca = datos_83.get("microcuenca", "")
        def_area = str(datos_83.get("area_ha", "0"))
        def_provincia = datos_83.get("provincia", "")
        def_distrito = datos_83.get("distrito", "")
        def_accesibilidad = datos_83.get("accesibilidad", 0)
        def_dia = datos_83.get("dia_evaluacion", 0)

        if datos_83:
            acc_txt = "Acceso limitado" if def_accesibilidad == 1 else "Acceso normal"
            st.info(f"Bloque **{def_codigo}** | Microcuenca: {def_microcuenca} | "
                    f"{def_distrito} ({def_provincia}) | {def_area} ha | "
                    f"{acc_txt} | Dia eval.: {def_dia}")

        st.markdown("---")
        with st.form("form_bloque", clear_on_submit=False):
            codigo = st.text_input("Codigo de bloque", value=def_codigo)
            cuenca = st.text_input("Cuenca", value="Cuenca Alta del Rio Piura")
            # Microcuenca: preseleccionar si viene de los 83 bloques
            mc_idx = 0
            if def_microcuenca and def_microcuenca in MICROCUENCAS:
                mc_idx = MICROCUENCAS.index(def_microcuenca) + 1
            microcuenca = st.selectbox("Microcuenca", [""]+MICROCUENCAS, index=mc_idx)
            # Provincia: preseleccionar
            prov_idx = 0
            if def_provincia and def_provincia in PROVINCIAS:
                prov_idx = PROVINCIAS.index(def_provincia) + 1
            provincia = st.selectbox("Provincia", [""]+PROVINCIAS, index=prov_idx)
            # Distrito: preseleccionar
            dist_list = _distritos(provincia)
            dist_idx = 0
            if def_distrito and def_distrito in dist_list:
                dist_idx = dist_list.index(def_distrito) + 1
            distrito = st.selectbox("Distrito", [""]+dist_list, index=dist_idx)
            tipo = st.selectbox("Tipo intervencion", TIPOS_INTERVENCION)
            a1,a2 = st.columns(2)
            ue = a1.text_input("UTM Este","0"); un = a2.text_input("UTM Norte","0")
            b1,b2 = st.columns(2)
            uz = b1.text_input("Zona UTM","17S"); alt = b2.text_input("Altitud","0")
            area = st.text_input("Area (ha)", value=def_area if datos_83 else "0")
            resp = st.text_input("Responsable")
            estado = st.selectbox("Estado", ESTADOS_BLOQUE)
            guardar = st.form_submit_button("Guardar", type="primary")
        if guardar:
            if not codigo: st.warning("Codigo obligatorio.")
            elif not distrito: st.warning("Seleccione distrito.")
            else:
                try:
                    db.insertar_bloque(codigo=codigo,tipo_intervencion=tipo,cuenca=cuenca,
                        distrito=distrito,utm_este=float(ue),utm_norte=float(un),
                        utm_zona=uz,area_hectareas=float(area),estado=estado,
                        altitud=float(alt or 0),responsable=resp,
                        microcuenca=microcuenca,provincia=provincia)
                    st.success(f"Bloque {codigo} registrado."); st.rerun()
                except Exception as e: st.error(f"Error: {e}")
    with ct:
        st.markdown("**Bloques Registrados**")
        busq = st.text_input("Buscar","",key="busq_bl")
        bloques = db.buscar_bloques(busq) if busq else db.obtener_bloques()
        if bloques:
            st.dataframe(pd.DataFrame([{"ID":b["id"],"Codigo":b["codigo"],
                "Microcuenca":b.get("microcuenca","") or "","Tipo":b["tipo_intervencion"],
                "Provincia":b.get("provincia","") or "","Distrito":b["distrito"],
                "UTM Este":f"{b['utm_este']:.2f}","UTM Norte":f"{b['utm_norte']:.2f}",
                "Altitud":f"{(b.get('altitud',0) or 0):.0f}",
                "Area":f"{b['area_hectareas']:.4f}",
                "Responsable":b.get("responsable","") or "","Estado":b["estado"]
                } for b in bloques]), use_container_width=True, hide_index=True)
            st.markdown("---")
            bm = {f"{b['codigo']} - {b['tipo_intervencion']}":b["id"] for b in bloques}
            sel = st.selectbox("Seleccionar bloque para eliminar",[""]+list(bm.keys()),key="del_bl")
            if sel and sel in bm and st.button("Eliminar bloque"):
                db.eliminar_bloque(bm[sel]); st.success("Eliminado."); st.rerun()
        else: st.info("Sin bloques.")

        st.markdown("---")
        with st.expander("Tabla de Referencia - 83 Bloques Preliminares de Intervencion", expanded=False):
            st.caption("Fuente: DATOS DE ORIGEN 83 BLOQUES.xlsx - Base de datos completa del proyecto IN Piura")
            df_83 = pd.DataFrame([{
                "N":b[0], "Bloque":b[1], "Microcuenca":b[2],
                "Area (ha)":b[3], "Provincia":b[4], "Distrito":b[5],
                "Accesibilidad":"Limitado" if b[6]==1 else "Normal",
                "Dia Eval.":b[7],
            } for b in BLOQUES_83])
            st.dataframe(df_83, use_container_width=True, hide_index=True, height=400)

# ══════════════════════════════════════════════════════════════════════════
# INSPECCION DE CAMPO
# ══════════════════════════════════════════════════════════════════════════
def pagina_inspeccion():
    st.subheader("Inspeccion de Campo")
    bm = _bloques_map()
    if not bm: st.warning("Registre un bloque primero."); return

    # Selector de bloque FUERA del form para auto-enlazar microcuenca
    opciones_bloque = list(bm.keys())
    bl = st.selectbox("Bloque", opciones_bloque, key="insp_bloque")

    # Auto-resolver microcuenca del bloque seleccionado
    mc_auto = _resolver_microcuenca(bl)
    if mc_auto:
        mc_idx = MICROCUENCAS.index(mc_auto) + 1  # +1 por el "" inicial
        st.info(f"Microcuenca vinculada automaticamente: **{mc_auto}**")
    else:
        mc_idx = 0

    # Carga de archivos PDF (fuera del form por limitaciones de Streamlit)
    pdf_files = st.file_uploader(
        "Adjuntar archivos PDF (max. 25 MB por archivo)",
        type=["pdf"], accept_multiple_files=True, key="insp_pdf_upload")
    if pdf_files:
        total_size = sum(f.size for f in pdf_files)
        for f in pdf_files:
            if f.size > 25 * 1024 * 1024:
                st.warning(f"El archivo '{f.name}' excede 25 MB y no sera adjuntado.")
        st.info(f"{len(pdf_files)} archivo(s) PDF seleccionado(s)")

    with st.form("form_insp", clear_on_submit=True):
        mc = st.selectbox("Microcuenca", [""] + MICROCUENCAS, index=mc_idx)
        fecha = st.date_input("Fecha de visita",value=datetime.now())
        inspector = st.text_input("Inspector")
        clima = st.selectbox("Condiciones climaticas",CONDICIONES_CLIMATICAS)
        avance = st.number_input("Avance fisico (%)",0.0,100.0,0.0)
        obs = st.text_area("Observaciones tecnicas")
        desv = st.text_area("Desviaciones observadas al Plan de Trabajo")
        ver = st.text_input("Codigo de verificacion",
            value=f"VER-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}")
        guardar = st.form_submit_button("Guardar Inspeccion", type="primary")
    if guardar:
        if not inspector: st.warning("Inspector obligatorio.")
        else:
            try:
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
    insp = db.obtener_todas_inspecciones()
    if insp:
        st.dataframe(pd.DataFrame([{"ID":i["id"],"Bloque":i["bloque_codigo"],
            "Microcuenca":i.get("microcuenca","") or "","Fecha":i["fecha_visita"],
            "Inspector":i["inspector"],"Avance %":f"{i['avance_fisico']:.1f}",
            "Verificacion":i["codigo_verificacion"],
            "PDFs":len([p for p in (i.get("archivos_pdf","") or "").split(";") if p.strip()])} for i in insp]),
            use_container_width=True, hide_index=True)

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
    bl = st.selectbox("Bloque", list(bm.keys()), key="ind_bl")
    bid = bm[bl]

    # Auto-resolver microcuenca del bloque seleccionado
    mc_auto = _resolver_microcuenca(bl)
    if mc_auto:
        mc_idx = MICROCUENCAS.index(mc_auto) + 1
        st.info(f"Microcuenca vinculada automaticamente: **{mc_auto}**")
    else:
        mc_idx = 0

    ins = db.obtener_inspecciones_por_bloque(bid)
    if not ins: st.warning("Sin inspecciones para este bloque."); return
    im = {f"{i['fecha_visita']} - {i['inspector']}":i["id"] for i in ins}
    isel = st.selectbox("Inspeccion", list(im.keys()))
    with st.form("form_ind", clear_on_submit=True):
        mc = st.selectbox("Microcuenca", [""] + MICROCUENCAS, index=mc_idx, key="ind_mc")
        pc = st.number_input("Cobertura vegetal (%)",0.0,100.0,0.0)
        tc = st.selectbox("Tipo cobertura",[""]+TIPOS_COBERTURA)
        vi = st.selectbox("Vigor cobertura",[""]+VIGOR_COBERTURA)
        so = st.number_input("Sobrevivencia especies (%)",0.0,100.0,0.0)
        guardar = st.form_submit_button("Guardar Indicadores", type="primary")
    if guardar:
        try:
            db.insertar_indicadores(bloque_id=bid,inspeccion_id=im[isel],
                cobertura_vegetal_planificada=0,cobertura_vegetal_lograda=0,
                sobrevivencia_especies=so,longitud_zanjas_ejecutada=0,
                volumen_retencion_sedimentos=0,porcentaje_cobertura_vegetal=pc,
                tipo_cobertura_vegetal=tc,vigor_cobertura_vegetal=vi,microcuenca=mc)
            st.success("Indicadores guardados."); st.rerun()
        except Exception as e: st.error(f"Error: {e}")
    st.markdown("---")
    st.markdown("**Indicadores Registrados**")
    ind = db.obtener_indicadores_por_bloque(bid)
    if ind:
        st.dataframe(pd.DataFrame([{"Fecha":x.get("fecha_visita",""),
            "Microcuenca":x.get("microcuenca","") or "",
            "Cobert.%":f"{x.get('porcentaje_cobertura_vegetal',0):.1f}",
            "Tipo":x.get("tipo_cobertura_vegetal","") or "",
            "Vigor":x.get("vigor_cobertura_vegetal","") or "",
            "Sobrev.%":f"{x['sobrevivencia_especies']:.1f}"} for x in ind]),
            use_container_width=True, hide_index=True)

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

    tab_reg, tab_hist = st.tabs(["Registro de Diagnostico", "Historial / Consulta"])

    with tab_reg:
        bl = st.selectbox("Bloque de Intervencion", list(bm.keys()), key="dt_bl")
        bid = bm[bl]
        r1, r2, r3 = st.columns(3)
        mc = r1.selectbox("Microcuenca", [""] + MICROCUENCAS, key="dt_mc")
        fecha_ev = r2.date_input("Fecha de evaluacion", value=datetime.now(), key="dt_fecha")
        evaluador = r3.text_input("Evaluador / Especialista", key="dt_eval")

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

        if st.button("Guardar Diagnostico Territorial", type="primary", key="dt_guardar"):
            if not evaluador:
                st.warning("Ingrese el nombre del evaluador.")
            elif not fichas_sel:
                st.warning("Complete al menos una ficha de diagnostico.")
            else:
                try:
                    db.insertar_diagnostico_territorial(
                        bloque_id=bid,
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
                    st.success(f"Diagnostico territorial guardado ({', '.join(fichas_sel)}).")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

    with tab_hist:
        st.markdown("### Historial de Diagnosticos Territoriales")
        todos_dt = db.obtener_todos_diagnosticos()
        if not todos_dt:
            st.info("No hay diagnosticos registrados.")
        else:
            st.dataframe(pd.DataFrame([{
                "ID": d["id"],
                "Bloque": d.get("bloque_codigo", ""),
                "Tipo": d.get("tipo_intervencion", ""),
                "Distrito": d.get("distrito", ""),
                "Fichas": d.get("ficha", ""),
                "Fecha Eval.": d.get("fecha_evaluacion", ""),
                "Evaluador": d.get("evaluador", ""),
                "Microcuenca": d.get("microcuenca", "") or "",
            } for d in todos_dt]), use_container_width=True, hide_index=True)

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
# PRESUPUESTO
# ══════════════════════════════════════════════════════════════════════════
def pagina_presupuesto():
    st.subheader("Presupuesto y Recursos")
    bm = _bloques_map()
    if not bm: st.warning("Registre un bloque primero."); return
    bl = st.selectbox("Bloque", list(bm.keys()), key="pres_bl")
    bid = bm[bl]
    with st.form("form_pres", clear_on_submit=True):
        cat = st.selectbox("Categoria",CATEGORIAS_PRESUPUESTO)
        desc = st.text_input("Descripcion")
        x1,x2 = st.columns(2)
        mp = x1.number_input("Monto planificado (S/)",0.0,value=0.0,format="%.2f")
        me = x2.number_input("Monto ejecutado (S/)",0.0,value=0.0,format="%.2f")
        fu = st.selectbox("Fuente financiamiento",FUENTES_FINANCIAMIENTO)
        guardar = st.form_submit_button("Guardar Partida", type="primary")
    if guardar:
        try:
            db.insertar_presupuesto(bid,cat,desc,mp,me,fu)
            st.success("Partida registrada."); st.rerun()
        except Exception as e: st.error(f"Error: {e}")
    st.markdown("---")
    st.markdown("**Partidas del Bloque**")
    pa = db.obtener_presupuesto_por_bloque(bid)
    if pa:
        tp = sum(p["monto_planificado"] for p in pa)
        te = sum(p["monto_ejecutado"] for p in pa)
        st.dataframe(pd.DataFrame([{"ID":p["id"],"Categoria":p["categoria"],
            "Descripcion":p["descripcion"],
            "Planificado":f"S/ {p['monto_planificado']:,.2f}",
            "Ejecutado":f"S/ {p['monto_ejecutado']:,.2f}",
            "%Ejec":f"{(p['monto_ejecutado']/p['monto_planificado']*100) if p['monto_planificado']>0 else 0:.1f}%",
            "Fuente":p["fuente_financiamiento"]} for p in pa]),
            use_container_width=True, hide_index=True)
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
    bl = st.selectbox("Bloque", list(bm.keys()), key="crono_bl")
    bid = bm[bl]
    with st.form("form_crono", clear_on_submit=True):
        act = st.selectbox("Actividad",ACTIVIDADES_TIPO)
        x1,x2 = st.columns(2)
        ip = x1.date_input("Inicio plan.",value=datetime.now())
        fp = x2.date_input("Fin plan.",value=datetime.now())
        x3,x4 = st.columns(2)
        ir = x3.text_input("Inicio real",""); fr = x4.text_input("Fin real","")
        x5,x6 = st.columns(2)
        av = x5.number_input("Avance %",0.0,100.0,0.0)
        ea = x6.selectbox("Estado",ESTADOS_ACTIVIDAD)
        re = st.text_input("Responsable"); ob = st.text_area("Observaciones")
        guardar = st.form_submit_button("Guardar Actividad", type="primary")
    if guardar:
        try:
            db.insertar_actividad(bloque_id=bid,actividad=act,
                fecha_inicio_plan=ip.strftime("%Y-%m-%d"),fecha_fin_plan=fp.strftime("%Y-%m-%d"),
                fecha_inicio_real=ir,fecha_fin_real=fr,porcentaje_avance=av,
                responsable=re,observaciones=ob,estado=ea)
            st.success("Actividad registrada."); st.rerun()
        except Exception as e: st.error(f"Error: {e}")
    st.markdown("---")
    st.markdown("**Actividades del Bloque**")
    acs = db.obtener_actividades_por_bloque(bid)
    if acs:
        st.dataframe(pd.DataFrame([{"ID":a["id"],"Actividad":a["actividad"],
            "Inicio":a["fecha_inicio_plan"],"Fin":a["fecha_fin_plan"],
            "Inicio Real":a["fecha_inicio_real"] or "-","Fin Real":a["fecha_fin_real"] or "-",
            "Avance":f"{a['porcentaje_avance']:.0f}%","Estado":a["estado"],
            "Responsable":a["responsable"]} for a in acs]),
            use_container_width=True, hide_index=True)
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
if pagina == "Panel de Control": pagina_dashboard()
elif pagina == "Bloques de Intervencion": pagina_bloques()
elif pagina == "Inspeccion de Campo": pagina_inspeccion()
elif pagina == "Indicadores de Calidad": pagina_indicadores()
elif pagina == "Diagnostico Territorial": pagina_diagnostico_territorial()
elif pagina == "Presupuesto": pagina_presupuesto()
elif pagina == "Cronograma": pagina_cronograma()
elif pagina == "Georreferenciacion": pagina_georreferenciacion()
elif pagina == "ODK / KoBoToolbox": pagina_odk()
elif pagina == "Reportes": pagina_reportes()
