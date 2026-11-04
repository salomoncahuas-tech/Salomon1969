# -*- coding: utf-8 -*-
"""Datos maestros del Plan de Contingencia del Estudio de Perfil del Proyecto IN Piura
ante el D.S. N° 149-2026-PCM (Estado de Emergencia por peligro inminente, FEN 2026-2027).

Un solo origen de datos para el Word, el Excel y el HTML. Todo lo que no tiene fuente
verificada queda marcado [POR DEFINIR], [POR VERIFICAR] o [SUPUESTO]."""
import json, sys, datetime as dt

CORTE = "04/11/2026"

META = {
    "titulo": "PLAN DE CONTINGENCIA PARA LA CULMINACIÓN DEL ESTUDIO DE PREINVERSIÓN A NIVEL DE PERFIL",
    "subtitulo": "Proyecto IN Piura · CUI 2669244 · en el contexto del Decreto Supremo N° 149-2026-PCM (Estado de Emergencia por peligro inminente ante intensas precipitaciones pluviales asociadas al Fenómeno El Niño 2026-2027)",
    "version": "1.0 · propuesta para validación del equipo formulador y la SESDI",
    "fecha": CORTE,
    "elaborado": "Ing. Hector Salomon Cahuas Miller · Ingeniero Forestal · ANIN - DIME - SESDI",
    "proyecto": "Recuperación del servicio de regulación de riesgos naturales y recuperación de ecosistemas degradados en la Cuenca Alta del Río Piura",
    "cui": "2669244", "fase": "Preinversión · Estudio de Perfil (48.5 % de avance ponderado al 03/11/2026)",
    "modalidad": "Administración Directa", "techo": 372442809, "tc": 3.40,
    "ambito": "3 provincias (Morropón, Huancabamba, Ayabaca) · 15 distritos · 117 bloques · 12,157.47 ha netas (catálogo V8) · 26 microcuencas",
    "crs": "UTM WGS 84 Zona 17S (EPSG:32717)",
}

DISTRITOS = [  # (distrito, provincia)
    ("Buenos Aires", "Morropón"), ("Chalaco", "Morropón"), ("Chulucanas", "Morropón"), ("Morropón", "Morropón"),
    ("Salitral", "Morropón"), ("San Juan de Bigote", "Morropón"), ("Santa Catalina de Mossa", "Morropón"),
    ("Santo Domingo", "Morropón"), ("Yamango", "Morropón"),
    ("Canchaque", "Huancabamba"), ("Huancabamba", "Huancabamba"), ("Huarmaca", "Huancabamba"),
    ("Lalaquiz", "Huancabamba"), ("San Miguel de El Faique", "Huancabamba"),
    ("Frías", "Ayabaca"),
]

# ---------------------------------------------------------------- Decreto y contexto climático
DECRETO = {
    "norma": "Decreto Supremo N° 149-2026-PCM",
    "publicacion": "Diario Oficial El Peruano, 03/11/2026",
    "objeto": "Declara el Estado de Emergencia en varios distritos de algunas provincias de 24 departamentos (entre ellos Piura) y la Provincia Constitucional del Callao, por peligro inminente ante intensas precipitaciones pluviales asociadas al Fenómeno El Niño 2026-2027.",
    "plazo": "60 días calendario (vigencia aproximada hasta inicios de enero de 2027; fecha exacta de término y eventual prórroga [POR VERIFICAR] en el texto del artículo 1 y en normas posteriores).",
    "alcance_nacional": "1,380 distritos de 191 provincias según la prensa (La República, El Tiempo, Correo, 03/11/2026).",
    "alcance_piura": "65 distritos del departamento de Piura en las provincias de Piura, Ayabaca, Huancabamba, Morropón, Paita, Sullana, Talara y Sechura (El Tiempo, 03/11/2026). La relación distrito por distrito figura en el Anexo del decreto: la inclusión de cada uno de los 15 distritos del proyecto está [POR VERIFICAR] contra el Anexo oficial (ver Anexo A del plan).",
    "acciones": "Los gobiernos regionales y locales, con la coordinación técnica y seguimiento del INDECI y la participación de los ministerios sectoriales (Salud, Educación, Desarrollo Agrario y Riego, Vivienda, Defensa, Interior, Transportes y Comunicaciones, Energía y Minas, Mujer y Poblaciones Vulnerables, entre otros), ejecutan medidas y acciones de excepción, inmediatas y necesarias, de reducción del Muy Alto Riesgo existente, así como de respuesta y rehabilitación. El INDECI coordina, supervisa y reporta a la PCM (prensa; texto del artículo 2 [POR VERIFICAR]).",
    "financiamiento": "Con cargo al presupuesto institucional de las entidades involucradas, sin demandar recursos adicionales al Tesoro Público (prensa; artículo 3 [POR VERIFICAR]).",
    "sustento": "Informes técnicos de SENAMHI, CENEPRED, ENFEN e INDECI sobre la continuidad de El Niño Costero y el desarrollo de El Niño en el Pacífico central (prensa; considerandos [POR VERIFICAR]).",
    "nota_acceso": "El texto oficial (busquedas.elperuano.pe/dispositivo/NL/2571017-4) no pudo descargarse desde el entorno de trabajo por restricción de red; lo consignado proviene de la reseña de El Peruano y de la prensa del 03/11/2026 y debe cotejarse con la norma y su Anexo antes de la aprobación del plan.",
    "fuentes": [
        ("El Peruano · Normas Legales · D.S. N° 149-2026-PCM", "https://busquedas.elperuano.pe/dispositivo/NL/2571017-4"),
        ("El Peruano · Fenómeno El Niño 2026-2027: declaran estado de emergencia en distritos de 25 regiones", "https://elperuano.pe/noticia/308230-fenomeno-el-nino-2026-2027-declaran-estado-de-emergencia-en-distritos-de-25-regiones"),
        ("La República · Gobierno declara en emergencia 1380 distritos del Perú (03/11/2026)", "https://larepublica.pe/sociedad/2026/11/03/gobierno-declara-en-emergencia-1380-distritos-del-peru-por-lluvias-intensas-ante-el-nino-20262027-119328"),
        ("El Tiempo (Piura) · Declaran en emergencia 65 distritos por peligro de lluvias", "https://eltiempo.pe/local/declaran-en-emergencia-65-distritos-por-peligro-de-lluvias/"),
        ("Caretas · Lima, Callao y más regiones en estado de emergencia por 60 días", "https://caretas.pe/sociedad/lima-callao-y-mas-regiones-en-estado-de-emergencia-por-60-dias-gobierno-declara-medida-ante-crisis-por-el-fenomeno-el-nino"),
        ("Exitosa · El Niño: Gobierno declara estado de emergencia (financiamiento)", "https://www.exitosanoticias.pe/actualidad/el-nino-gobierno-declara-estado-emergencia-distritos-todas-regiones-peru-peligro-lluvias-n187431"),
        ("Andina · Normas legales más importantes del 03/11/2026", "https://andina.pe/agencia/noticia-estas-son-las-normas-legales-mas-importantes-del-martes-3-noviembre-2026-1098356.aspx"),
    ],
}

CLIMA = {
    "comunicado": "Comunicado Oficial ENFEN N° 19-2026 (30/10/2026): se mantiene la Alerta de El Niño Costero.",
    "magnitud": "Región Niño 1+2 con magnitud extraordinaria hasta febrero de 2027 y descenso gradual hasta junio de 2027; El Niño en el Pacífico central (Niño 3.4) de magnitud muy fuerte hasta febrero de 2027.",
    "lluvias": "Entre noviembre de 2026 y enero de 2027 las precipitaciones seguirán por encima de lo normal en la costa norte y la sierra norte occidental; la tendencia continúa durante el verano de 2027.",
    "implicancia": "El ámbito del proyecto (vertiente occidental andina de Piura, 498 a 1,420 m de altitud media por provincia, 69.0 % del área con peligro integrado Alto o Muy alto y 93.9 % de los bloques con peligro alto o muy alto por movimientos en masa según el EVAR) coincide con la zona de mayor incremento esperado de lluvias. La temporada crítica para el trabajo de campo es diciembre de 2026 a abril de 2027.",
    "fuentes": [
        ("ENFEN · Comunicados oficiales", "https://enfen.imarpe.gob.pe/comunicados/"),
        ("ENFEN · Comunicado Oficial N° 16-2026 (14/09/2026)", "https://enfen.imarpe.gob.pe/download/comunicado-oficial-enfen-n-16-2026/"),
        ("Radio RSD · ENFEN mantiene alerta de El Niño Costero hasta febrero de 2027", "https://radiorsd.pe/noticias/enfen-mantiene-alerta-de-el-nino-costero-hasta-febrero-de-2027"),
        ("CENEPRED · Escenarios de riesgo por inundaciones y movimientos en masa (Plan Multisectorial 2025-2027)", "https://sigrid.cenepred.gob.pe/sigridv3/storage/biblioteca//17791_escenarios-de-riesgo-por-inundaciones-y-movimientos-en-masa-en-el-marco-del-plan-multisectorial-2025-2027.pdf"),
    ],
}

ELECTORAL = {
    "erm": "Elecciones Regionales y Municipales del 04/10/2026 (autoridades para el periodo 2027-2030). Las autoridades regionales y municipales vigentes culminan su mandato el 31/12/2026 y las electas asumen funciones el 01/01/2027 (Ley N° 27867 y Ley N° 27972).",
    "segunda_vuelta": "Piura va a Segunda Elección Regional: ninguna fórmula alcanzó el 30 % de votos válidos (ONPE/JNE). La prensa reporta la fecha reprogramada por el JNE para el 06/12/2026 (antes 29/11/2026) [POR VERIFICAR con el JNE].",
    "municipal": "Se renuevan las 3 municipalidades provinciales (Morropón, Huancabamba, Ayabaca) y las 15 municipalidades distritales del ámbito; los alcaldes electos por distrito están [POR VERIFICAR] en los resultados oficiales de la ONPE/JNE.",
    "comunal": "Las juntas directivas de las comunidades campesinas se renuevan cada dos años (Ley N° 24656 y su reglamento). El calendario de cada comunidad del ámbito (3,929.61 ha comunales, 32.3 % de la superficie neta; 83.9 % en Ayabaca) está [POR VERIFICAR] con la DRA Piura y SUNARP; en la práctica varias asambleas eleccionarias se realizan entre diciembre y enero.",
    "nacional": "El Gobierno Nacional electo en 2026 asumió el 28/07/2026; los cambios de titulares en la PCM, el MEF y la ANIN posteriores a esa fecha [POR VERIFICAR] afectan las instancias de aprobación y priorización.",
    "fuentes": [
        ("Andina · ERM 2026: 15 de las 25 regiones irían a segunda vuelta", "https://andina.pe/agencia/noticia-erm-2026-15-las-25-regiones-irian-a-segunda-vuelta-para-elegir-a-su-gobernador-1094780.aspx"),
        ("El Tiempo · Segunda vuelta decidirá gobernador de Piura", "https://eltiempo.pe/local/segunda-vuelta-decidira-gobernador-de-piura/"),
        ("Panamericana · JNE cambia fecha de segunda vuelta regional al 6 de diciembre", "https://panamericana.pe/politica/475578-jne-cambia-fecha-segunda-vuelta-elecciones-regionales-2026-6-diciembre"),
        ("Forbes Perú · Gobierno convoca ERM para el 4 de octubre", "https://forbes.pe/actualidad/2026-01-07/gobierno-convoca-elecciones-regionales-y-municipales-para-el-4-de-octubre/"),
    ],
}

# ---------------------------------------------------------------- Estado del Perfil al 03/11/2026
ENTREGABLES = [
 # id, capítulo, nombre, peso, avance 03/11, estado, sustento al corte
 ("E1","ID","Microlocalización e identificación de las unidades de intervención",8,85,"En curso","117 bloques verificados en campo; catálogo V8 (09/10/2026) con 12,157.47 ha netas; 45 bloques con código M#-B# y 72 por codificar; polígonos de 5 bloques por recortar."),
 ("E2","ID","Diagnóstico de la población afectada y actores; expediente de liberación de áreas",10,75,"En curso","Volúmenes I-III en borrador B3 (03/11/2026): 422 fichas F-DS en 85 de 117 bloques. Tamizaje predial: 22 predios inspeccionados en 10 bloques de Morropón, 1 acta observada, 0 conformes."),
 ("E3","ID","Estudio de Geología",5,100,"Culminado","Culminado e integrado en los volúmenes de diagnóstico (ISL-MM medio 3.81; ISL-EH 3.68)."),
 ("E4","ID","Estudio de Hidrología",6,40,"En curso","Informe de avance: 60 microcuencas y 264 nanocuencas (QSWAT); 15 nanocuencas priorizadas. Faltan lluvias máximas de 24 h, curvas IDF y caudales HEC-HMS."),
 ("E5","ID","Estudio de Suelos",6,100,"Culminado","Informe final (GEOSIG, 01/11/2026): 74 calicatas, 152 muestras, 13 pruebas de infiltración, 7 unidades de suelo, Factor K, CUM y matriz suelo-intervención."),
 ("E6","ID","Diagnóstico del ecosistema y oferta de material vegetal",8,39,"En curso","Borrador del informe final (15/10/2026), fases 1-5 de 7: 8,341.82 ha bajo el umbral MSAVI y 69 especies nativas. Faltan el censo de viveros (fase 6) y la validación ciega (fase 7)."),
 ("E7","ID","Problema, causas y efectos; planteamiento del proyecto",5,80,"En curso","Árbol de problemas unificado v1.0 (30/10/2026); falta el Informe Final con medios y fines alineados."),
 ("E8","ID","Análisis de riesgos en contexto de cambio climático",8,100,"Culminado","Informe final EVAR (19/10/2026, Manual CENEPRED v02): peligro en 117 bloques y 85 áreas de aporte; riesgo en 117 áreas de influencia. Sin escenario de cambio climático ni valoración de daños (insumos del E13)."),
 ("E9","ID","Informe del Capítulo de Identificación",8,40,"En curso","Integración de diagnósticos y estudios básicos; incorpora E5 final, EVAR y árbol v1.0."),
 ("E10","FOR","Mercado, análisis técnico, metas, costos y presupuesto",10,25,"En curso","Matriz de Costos Paramétricos v1.6 (03/11/2026): inversión S/ 241.23 M y O&M S/ 115.40 M (95.8 % del techo). Coeficientes por calibrar con el piloto; cárcavas sin metrado; supervisión y expediente por definir."),
 ("E11","FOR","Plan de gestión de la implementación",4,0,"Programado","Programado desde diciembre de 2026."),
 ("E12","FOR","Plan de operación y mantenimiento",4,0,"Programado","Programado desde noviembre de 2026; tarifas de O&M (ANIN, Anexo 11) ya cargadas en la Matriz v1.6."),
 ("E13","EVA","Evaluación social, sostenibilidad y marco lógico",8,0,"Programado","Programado desde diciembre de 2026; requiere beneficios por daños evitados (EVAR no valora daños)."),
 ("E14","EVA","Informe de Formulación y Evaluación",5,0,"Programado","Programado desde diciembre de 2026."),
 ("E15","RES","Resumen ejecutivo y anexos del Perfil",5,0,"Programado","Programado desde febrero de 2027."),
]
CAPITULOS = {"ID":"Capítulo I · Identificación","FOR":"Capítulo II · Formulación","EVA":"Capítulo III · Evaluación","RES":"Resumen ejecutivo y anexos"}

PROGRAMACION_ORIGINAL = [  # presentación SESDI 03/11/2026
    ("Dic 2026", "≥ 66.5 %", "Capítulo de Identificación culminado (E1 a E9) con el E10 al 25 %"),
    ("Feb 2027", "≥ 95 %", "Capítulos de Formulación y Evaluación (E10 a E14)"),
    ("Mar 2027", "100 %", "Perfil completo con resumen ejecutivo y anexos"),
    ("Abr 2027", "—", "Revisión, registro en el Banco de Inversiones y declaración de viabilidad"),
]

# ---------------------------------------------------------------- Actividades pendientes por entregable
# campo: "Sí" (depende de campo), "Parcial", "No" (gabinete)
PENDIENTES = [
 # entregable, actividad, campo, ventana propuesta, producto, prioridad, dependencia
 ("E1","Recortar los polígonos V8 de los bloques 5, 22, M17B1, 8 y 71 y recalcular PI, MSAVI, EVAR y costos en UTM 17S","No","Nov-Dic 2026","Capa V8 definitiva (gpkg) y tablas recalculadas","Media","Equipo SIG"),
 ("E1","Decidir sobre el bloque 8 (conserva el 25.4 % de su área V7): mantener, redefinir o descartar","No","Nov 2026","Acta de decisión técnica","Media","E1, E10"),
 ("E1","Asignar código definitivo M#-B# a los 72 bloques numéricos","No","Dic 2026","Catálogo con codificación única","Baja","Equipo SIG"),
 ("E1","Pendiente de campo: reemplazar la medición con clinómetro por el análisis del DEM y validar solo en bloques cercanos al límite de 75 %","Parcial","Nov 2026 y Abr-May 2027","Tabla de pendientes por bloque","Baja","DEM, campo residual"),
 ("E1","Catálogo definitivo de bloques y lotes SUS tras la liberación de áreas","Sí","May 2027","Catálogo V9 definitivo","Alta","E2-SE2"),
 ("E2","SE1 · Levantar fichas F-DS-01/02/03 en los 32 bloques sin ficha (22 en Huancabamba, concentrados en Huarmaca; Ayabaca sin F-DS-03)","Sí","Nov 2026 (prioritario) y Abr-May 2027 (residual)","Fichas F-DS en 117 bloques","Alta","Ventana de campo"),
 ("E2","SE1 · Aplicar F-DS-04, F-DS-05 y F-DS-07 en el conglomerado 83-87 y en los bloques comunales; F-DS-06 fuera de Morropón","Sí","Nov 2026 y Abr 2027","Fichas aplicadas","Alta","Asambleas comunales"),
 ("E2","SE1 · Completar la digitación en el aplicativo (70.1 %) y el registro de saberes ancestrales (3 entrevistas sin ingresar)","No","Nov-Dic 2026","Base de datos consolidada","Media","Gabinete"),
 ("E2","SE1 · Volúmenes I-III versión final del diagnóstico social","No","Dic 2026 (versión A) y May 2027 (versión final)","Informe E2 final","Alta","Fichas de campo"),
 ("E2","SE2 · Ruta F-LA (F-LA-01 asamblea, F-LA-02 registro, F-LA-03 inspección, F-LA-04 acta) en los 5 bloques piloto (21, 31, 66, 73, M6B10) y en los de mayor riesgo (AI_73 Chococa y AI_41 Cambruran, Huarmaca)","Sí","Nov 2026","Actas conformes en bloques piloto y críticos","Alta","Ventana de campo; autoridades vigentes"),
 ("E2","SE2 · Resolver las 7 observaciones (6 inspecciones con interferencia de bosque natural y 1 con vivienda habitada) y el acta observada del bloque 85","Parcial","Nov 2026","Observaciones levantadas","Alta","Gabinete y campo"),
 ("E2","SE2 · Identificar el polígono comunal sin nombre (ID 2790, 10 bloques) con la DRA Piura y SUNARP","No","Nov-Dic 2026","Titularidad identificada","Alta","DRA Piura, SUNARP"),
 ("E2","SE2 · Tamizaje de los 5 bloques con más de 19 predios (4, 9, 35, 38, 43): redefinir o descartar","No","Dic 2026","Decisión por bloque","Media","Capa predial"),
 ("E2","SE2 · Ruta F-LA en los 107 bloques y 60 lotes SUS restantes (256.97 ha)","Sí","Abr-May 2027 (campo residual) y, si el clima lo permite, ventanas de enero a marzo","Actas de libre disponibilidad y expedientes AX-04","Alta","Clima; nuevas autoridades"),
 ("E2","SE2 · Verificación registral F-LA-05 (SUNARP, COFOPRI, DRA Piura) y armado de expedientes AX-04","No","Dic 2026 - May 2027","Expedientes por unidad","Alta","Entidades registrales"),
 ("E4","Lluvias máximas de 24 h (PISCO y 6 estaciones, 1964-2023, años El Niño), curvas IDF y caudales HEC-HMS (SCS-CN) para las 15 nanocuencas priorizadas","No","Nov-Dic 2026","Informe final del Estudio de Hidrología","Alta","Consultoría hidrológica"),
 ("E4","Incorporar los registros de precipitación y eventos del FEN 2026-2027 (SENAMHI, estaciones del ámbito) como escenario de contraste","No","Feb-Abr 2027","Anexo de evento 2026-2027","Media","SENAMHI"),
 ("E4","Vuelo de dron para validar el drenaje y las cárcavas en las nanocuencas priorizadas","Sí","Nov 2026 (si hay ventana) o Abr-May 2027","Ortomosaicos y modelo de drenaje","Media","Clima; servicio de dron"),
 ("E6","Fase 6: censo de viveros y brecha de plantones (oferta declarada 21,000 frente a 4,959,516 requeridos en la Matriz v1.6); propuesta de vivero central y temporales","Parcial","Nov 2026 - Ene 2027 (visitas a viveros en valle y gabinete)","Informe de oferta de material vegetal","Alta","Viveros del ámbito"),
 ("E6","Fase 7: validación ciega de 567 puntos; alternativa remota con imágenes de alta resolución si el campo no es posible","Sí","Abr-May 2027 (campo) o Ene-Mar 2027 (remota)","Informe final del E6 validado","Media","Clima; imágenes"),
 ("E6","Recalcular la línea base del indicador (8,341.83 ha bajo el umbral MSAVI) sobre los polígonos V8","No","Dic 2026","Cifra del indicador de brecha (R.M. N° 00213-2024-MINAM)","Alta","E1"),
 ("E7","Informe Final del E7 con árbol de medios y fines, objetivo central y alternativas alineadas al árbol v1.0","No","Nov 2026","Informe final E7","Alta","Gabinete"),
 ("E8","Escenario de cambio climático SSP5-8.5 para el desencadenante; registro de eventos del FEN 2026-2027 como evidencia","No","Dic 2026 - Mar 2027","Anexo del EVAR","Media","E4; SENAMHI; INDECI/COER"),
 ("E8","Inventario de elementos expuestos y costos de reposición (beneficios por daños evitados) para el E13","Parcial","Dic 2026 - Feb 2027","Base de elementos expuestos y daños","Alta","F-EE, SINPAD/INDECI"),
 ("E9","Integrar el Capítulo de Identificación con los vacíos de campo explícitos y su plan de cierre","No","Dic 2026","Informe E9 versión para revisión","Alta","E1-E8"),
 ("E9","Actualizar el Capítulo con los resultados del campo residual","No","May 2027","Informe E9 final","Alta","Campo residual"),
 ("E10","Análisis de mercado del servicio ecosistémico (oferta, demanda y brecha) y análisis técnico de alternativas","No","Nov 2026 - Ene 2027","Capítulo de mercado y análisis técnico","Alta","E6, E8"),
 ("E10","Calibrar metas y costos con el piloto (21, 31, 66, 73, M6B10); validar terraza sin pirca; metrar cárcavas; definir supervisión y expediente técnico","Parcial","Dic 2026 - Feb 2027","Matriz de costos v2.0","Alta","Piloto; E5"),
 ("E10","Compatibilizar las zanjas del plano (2,250.56 km) con la aptitud del E5 (7.44 km): pruebas de infiltración adicionales y reemplazo por terraza o barrera viva","Sí","Nov 2026 y Abr-May 2027","Plano de obras ajustado","Alta","Campo; E5"),
 ("E10","Decidir las obras en los bloques con P < 400 mm (8, 20, 71, M4B3, M4B4)","Parcial","Dic 2026","Decisión técnica","Media","Ensayos"),
 ("E10","Cronograma de ejecución (5 años) y presupuesto por componente a precios de mercado y sociales","No","Feb 2027","Presupuesto del Perfil","Alta","Metas"),
 ("E11","Plan de gestión de la implementación por Administración Directa: organización, UEI, cronograma, adquisiciones, riesgos de ejecución","No","Dic 2026 - Feb 2027","Informe E11","Media","E10"),
 ("E12","Plan de operación y mantenimiento (15 años): actividades, costos (tarifas ANIN Anexo 11), arreglos institucionales (comités, MERESE), monitoreo","No","Nov 2026 - Feb 2027","Informe E12","Media","E10, E2"),
 ("E13","Evaluación social (costo-efectividad y, de ser posible, VANS y TIR social con beneficios por daños evitados), análisis de sensibilidad, sostenibilidad y marco lógico","No","Ene - Mar 2027","Informe E13","Alta","E8, E10, E12"),
 ("E14","Informe de Formulación y Evaluación","No","Mar 2027 (borrador) y May 2027 (final)","Informe E14","Alta","E10-E13"),
 ("E15","Resumen ejecutivo, anexos y Formato N° 07-A actualizado para el registro en el Banco de Inversiones","No","May 2027","Perfil completo","Alta","Todos"),
]

# ---------------------------------------------------------------- Cronograma del Plan (escenario base)
# (id, fase, actividad, inicio, fin, entregables, campo, responsable, hito)
CRONOGRAMA = [
 ("A01","F0","Aprobación del Plan de Contingencia por la SESDI y comunicación al equipo y a los consultores","2026-11-04","2026-11-10","Todos","No","SESDI / Coordinación del Perfil","H0 · Plan aprobado (10/11/2026)"),
 ("A02","F0","Protocolo de seguridad de brigadas en temporada de lluvias (SENAMHI, COER Piura, umbrales de suspensión, comunicación, EPP, SCTR)","2026-11-04","2026-11-13","E2, E6, E10","No","Coordinación del Perfil / Seguridad y salud","—"),
 ("A03","F0","Coordinación con el COER Piura, INDECI y las municipalidades para el acceso a vías y la agenda de asambleas durante la emergencia","2026-11-04","2026-11-30","E2","No","Coordinación / Especialista social","—"),
 ("A10","F1","Ventana de campo intensiva: ruta F-LA en bloques piloto (21, 31, 66, 73, M6B10) y de mayor riesgo (Huarmaca: AI_73, AI_41)","2026-11-09","2026-11-30","E2-SE2","Sí","Brigadas de liberación de áreas","H1 · Fin de la ventana de campo (30/11/2026)"),
 ("A11","F1","Fichas F-DS en los 32 bloques sin ficha (Huarmaca y Frías prioritarios); F-DS-05 y F-DS-07 en bloques comunales y 83-87","2026-11-09","2026-11-30","E2-SE1","Sí","Brigadas sociales","—"),
 ("A12","F1","Pruebas de infiltración adicionales en bloques con zanjas proyectadas y registro F-DT-02 de cárcavas accesibles","2026-11-09","2026-11-30","E5, E10","Sí","Especialista en suelos / brigadas","—"),
 ("A13","F1","Levantamiento de las 7 observaciones de inspección y del acta del bloque 85; identificación del polígono comunal ID 2790 (DRA Piura, SUNARP)","2026-11-09","2026-12-18","E2-SE2","Parcial","Especialista legal-predial","—"),
 ("A14","F1","Informe Final del E7 (medios y fines, objetivo, alternativas)","2026-11-04","2026-11-28","E7","No","Coordinación del Perfil","—"),
 ("A15","F1","Estudio de Hidrología: lluvias de diseño, curvas IDF y caudales HEC-HMS","2026-11-04","2026-12-23","E4","No","Consultoría hidrológica","—"),
 ("A16","F1","E6 fase 6: censo de viveros y brecha de plantones; propuesta de viveros","2026-11-16","2027-01-29","E6","Parcial","Especialista en ecosistemas / viveros","—"),
 ("A17","F1","Inicio del análisis de mercado del servicio ecosistémico y del plan de O&M","2026-11-16","2026-12-23","E10, E12","No","Economista / Especialista O&M","—"),
 ("A20","F2","Recorte de polígonos V8, recálculo de PI, MSAVI, EVAR y costos; decisión del bloque 8; línea base del indicador sobre V8","2026-12-01","2026-12-23","E1, E6, E8, E10","No","Equipo SIG","—"),
 ("A21","F2","Digitación y consolidación del E2 (aplicativo, saberes ancestrales); Volúmenes I-III versión A con vacíos explícitos","2026-12-01","2026-12-23","E2","No","Especialista social","—"),
 ("A22","F2","Capítulo de Identificación (E9) versión para revisión de la SESDI, con el plan de cierre de vacíos de campo","2026-12-01","2026-12-23","E9","No","Coordinación del Perfil","H2 · Identificación en revisión (23/12/2026)"),
 ("A23","F2","Cierre administrativo del año fiscal 2026: certificaciones, conformidades, renovación o nueva contratación del personal y servicios para 2027","2026-12-01","2026-12-31","Todos","No","Coordinación / Administración DIME","—"),
 ("A24","F2","Monitoreo climático y de eventos (SENAMHI, COER, SINPAD) en los 15 distritos; registro de daños y activación de cárcavas como evidencia para E4, E8 y E13","2026-12-01","2027-04-30","E4, E8, E13","No","Especialista en riesgos / SIG","—"),
 ("A30","F3","Mapeo y re-socialización con las nuevas autoridades regionales, provinciales, distritales y comunales (asumen el 01/01/2027): reuniones, actas de compromiso, actualización del mapa de actores","2027-01-04","2027-02-26","E2, E11","Parcial","Especialista social / Coordinación","H3 · Actores reactualizados (26/02/2027)"),
 ("A31","F3","Calibración de metas y costos con el piloto; Matriz de costos v2.0; cronograma de ejecución y presupuesto","2027-01-04","2027-02-26","E10","No","Economista / Especialista técnico","—"),
 ("A32","F3","Plan de gestión de la implementación (E11) y Plan de O&M (E12)","2027-01-04","2027-02-26","E11, E12","No","Especialistas E11/E12","H4 · Formulación en borrador (26/02/2027)"),
 ("A33","F3","Inventario de elementos expuestos y daños (con registros del FEN 2026-2027); escenario SSP5-8.5 del EVAR","2027-01-04","2027-02-26","E8, E13","No","Especialista en riesgos","—"),
 ("A34","F3","Campo oportunista en ventanas sin lluvia: asambleas F-LA-01 y registro F-LA-02 en bloques accesibles (previa autorización del protocolo de seguridad)","2027-01-11","2027-03-26","E2-SE2","Sí","Brigadas","—"),
 ("A35","F3","Evaluación social, sostenibilidad y marco lógico (E13); Informe de Formulación y Evaluación (E14) borrador","2027-02-01","2027-03-31","E13, E14","No","Economista / Coordinación","H5 · Evaluación en borrador (31/03/2027)"),
 ("A40","F4","Campo residual post-lluvias: ruta F-LA en los bloques y lotes SUS restantes; validación ciega del E6 (567 puntos); infiltración complementaria; verificación de pendientes","2027-04-05","2027-05-14","E1, E2, E6, E10","Sí","Brigadas / especialistas","H6 · Campo residual culminado (14/05/2027)"),
 ("A41","F4","Verificación registral F-LA-05 y expedientes AX-04; catálogo definitivo de bloques y lotes SUS","2027-01-04","2027-05-21","E1, E2","No","Especialista legal-predial / SIG","—"),
 ("A42","F4","Actualización final de E2, E9, E10, E13 y E14 con los datos del campo residual","2027-05-03","2027-05-28","E2, E9, E10, E13, E14","No","Equipo formulador","—"),
 ("A43","F4","Resumen ejecutivo, anexos y Formato N° 07-A; Perfil completo","2027-05-10","2027-05-31","E15","No","Coordinación del Perfil","H7 · Perfil completo (31/05/2027)"),
 ("A50","F5","Revisión de la SESDI y la OPMI, levantamiento de observaciones, registro en el Banco de Inversiones y declaración de viabilidad","2027-06-01","2027-06-30","Todos","No","SESDI / OPMI ANIN","H8 · Viabilidad (junio de 2027)"),
]
FASES = {
 "F0": ("Fase 0 · Activación del plan", "04/11 - 13/11/2026"),
 "F1": ("Fase 1 · Ventana de campo intensiva antes de las lluvias", "09/11 - 30/11/2026"),
 "F2": ("Fase 2 · Gabinete y cierre de la Identificación", "01/12 - 31/12/2026"),
 "F3": ("Fase 3 · Formulación y Evaluación en gabinete durante la temporada de lluvias; re-socialización con nuevas autoridades", "04/01 - 31/03/2027"),
 "F4": ("Fase 4 · Campo residual post-lluvias y cierre del Perfil", "05/04 - 31/05/2027"),
 "F5": ("Fase 5 · Revisión, registro y viabilidad", "01/06 - 30/06/2027"),
}

ESCENARIOS = [
 ("Optimista","Lluvias de intensidad moderada en el ámbito; vías transitables la mayor parte de enero a marzo; autoridades nuevas receptivas; equipo completo.","Campo continuo con pausas cortas; se recupera la programación original.","31/03/2027","Abril de 2027","0 meses respecto de la programación original"),
 ("Base (este plan)","Lluvias por encima de lo normal (ENFEN N° 19-2026); campo suspendido de enero a marzo salvo ventanas; transición de autoridades consume enero y febrero; equipo estable.","Campo intensivo en noviembre, gabinete en la temporada de lluvias, campo residual en abril y mayo.","31/05/2027","Junio de 2027","+2 meses"),
 ("Pesimista","Estado de emergencia prorrogado; daños en vías hasta mayo; reasignación parcial del equipo DIME a la emergencia; demora en contratar servicios 2027.","Campo residual se desplaza a junio y julio; se evalúa Perfil con cartera condicionada.","31/07/2027","Agosto - setiembre de 2027","+4 meses"),
]

# ---------------------------------------------------------------- Presupuesto y recursos (estructura; CU por definir)
# (código, componente, partida, unidad, cantidad, base del cálculo [SUPUESTO], CU referencial o None, fuente del CU, fase)
VIATICO_REF = 320.00  # D.S. N° 007-2013-EF: viáticos por día para comisiones de servicio en el territorio nacional (funcionarios y servidores)
PRESUPUESTO = [
 ("A","A · Campo intensivo (noviembre de 2026)","Viáticos de brigadas de campo (liberación de áreas, social, suelos)","persona-día",3*2*18,"3 brigadas × 2 integrantes × 18 días hábiles de campo [SUPUESTO]",VIATICO_REF,"D.S. N° 007-2013-EF (S/ 320 por día); confirmar escala vigente de la ANIN","F1"),
 ("A","A · Campo intensivo (noviembre de 2026)","Alquiler de camioneta 4x4 con conductor","vehículo-día",3*18,"3 vehículos × 18 días [SUPUESTO]",None,"Cotización o contrato vigente de la DIME [POR DEFINIR]","F1"),
 ("A","A · Campo intensivo (noviembre de 2026)","Combustible","galón",3*18*8,"8 galones por vehículo y día en trocha [SUPUESTO]",None,"Precio local en Piura [POR DEFINIR]","F1"),
 ("A","A · Campo intensivo (noviembre de 2026)","Asambleas comunales y reuniones F-LA-01 (local, refrigerios, materiales, difusión)","evento",12,"Bloques piloto, críticos y comunales priorizados [SUPUESTO]",None,"[POR DEFINIR]","F1"),
 ("A","A · Campo intensivo (noviembre de 2026)","Pruebas de infiltración adicionales (doble anillo) en bloques con zanjas proyectadas","prueba",10,"Unidades de suelo sin prueba propia y bloques piloto [SUPUESTO]",None,"Tarifa de laboratorio o consultor E5 [POR DEFINIR]","F1"),
 ("B","B · Campo residual (abril - mayo de 2027)","Viáticos de brigadas de campo","persona-día",3*2*28,"3 brigadas × 2 integrantes × 28 días hábiles [SUPUESTO]",VIATICO_REF,"D.S. N° 007-2013-EF; confirmar escala vigente de la ANIN","F4"),
 ("B","B · Campo residual (abril - mayo de 2027)","Alquiler de camioneta 4x4 con conductor","vehículo-día",3*28,"3 vehículos × 28 días [SUPUESTO]",None,"[POR DEFINIR]","F4"),
 ("B","B · Campo residual (abril - mayo de 2027)","Combustible","galón",3*28*8,"8 galones por vehículo y día [SUPUESTO]",None,"[POR DEFINIR]","F4"),
 ("B","B · Campo residual (abril - mayo de 2027)","Asambleas y reuniones F-LA-01 / F-LA-04 (107 bloques y 60 lotes SUS restantes, agrupados por localidad)","evento",40,"Agrupación de unidades por centro poblado [SUPUESTO]",None,"[POR DEFINIR]","F4"),
 ("B","B · Campo residual (abril - mayo de 2027)","Validación ciega del E6 (567 puntos) y verificación de pendientes","punto",567,"Fase 7 del E6",None,"Incluido en honorarios o servicio específico [POR DEFINIR]","F4"),
 ("C","C · Especialistas de gabinete (noviembre de 2026 - junio de 2027)","Especialista en evaluación social de proyectos (E13) · locación de servicios","mes",4,"Enero a abril de 2027 [SUPUESTO]",None,"Honorario según TDR y escala de la ANIN [POR DEFINIR]","F3"),
 ("C","C · Especialistas de gabinete (noviembre de 2026 - junio de 2027)","Especialista en gestión de la implementación y O&M (E11, E12)","mes",4,"Noviembre de 2026 a febrero de 2027 [SUPUESTO]",None,"[POR DEFINIR]","F1-F3"),
 ("C","C · Especialistas de gabinete (noviembre de 2026 - junio de 2027)","Especialista legal-predial (F-LA-05, expedientes AX-04, polígono comunal ID 2790)","mes",7,"Noviembre de 2026 a mayo de 2027 [SUPUESTO]",None,"[POR DEFINIR]","F1-F4"),
 ("C","C · Especialistas de gabinete (noviembre de 2026 - junio de 2027)","Continuidad del equipo núcleo del Perfil (coordinación, SIG, social, riesgos, ecosistemas) durante 2027","mes",6,"Enero a junio de 2027; número de personas y escala [POR DEFINIR]",None,"Planilla / contratos vigentes de la DIME [POR DEFINIR]","F3-F5"),
 ("D","D · Estudios y servicios especializados","Culminación del Estudio de Hidrología (E4): saldo contractual o adenda","global",1,"Según contrato vigente [POR VERIFICAR]",None,"Contrato E4 [POR DEFINIR]","F1-F2"),
 ("D","D · Estudios y servicios especializados","Vuelo de dron (ortomosaico y MDE) en las 15 nanocuencas priorizadas y bloques con cárcavas","hectárea",5936,"Superficie de las 15 nanocuencas priorizadas (E4)",None,"Cotización de servicio [POR DEFINIR]","F1 o F4"),
 ("D","D · Estudios y servicios especializados","Imágenes satelitales de alta resolución post-evento (validación remota E6 y actualización de cárcavas)","km²",122,"12,157.47 ha netas V8 ≈ 122 km²",None,"Proveedor de imágenes [POR DEFINIR]","F3"),
 ("D","D · Estudios y servicios especializados","Ensayos de laboratorio complementarios (suelos, Factor K) si se requieren","muestra",0,"[POR DEFINIR] según resultados de infiltración",None,"Laboratorio [POR DEFINIR]","F4"),
 ("E","E · Socialización con nuevas autoridades (enero - marzo de 2027)","Talleres de presentación del proyecto a autoridades regionales, provinciales y distritales electas","taller",4,"1 regional + 3 provinciales [SUPUESTO]",None,"[POR DEFINIR]","F3"),
 ("E","E · Socialización con nuevas autoridades (enero - marzo de 2027)","Reuniones distritales y con juntas directivas comunales renovadas","reunión",15,"15 distritos [SUPUESTO: una por distrito]",None,"[POR DEFINIR]","F3"),
 ("E","E · Socialización con nuevas autoridades (enero - marzo de 2027)","Material de difusión (mapas por distrito, cartillas, actas modelo)","juego",20,"15 distritos + 3 provincias + GORE + reserva [SUPUESTO]",None,"[POR DEFINIR]","F3"),
 ("F","F · Seguridad, salud y comunicaciones","Equipos de protección personal para lluvias (botas, ponchos, cascos, chalecos, linternas)","juego",8,"6 brigadistas + 2 de reserva [SUPUESTO]",None,"[POR DEFINIR]","F1"),
 ("F","F · Seguridad, salud y comunicaciones","Seguro Complementario de Trabajo de Riesgo (SCTR) salud y pensión para el personal de campo","persona-mes",8*3,"8 personas × 3 meses de campo [SUPUESTO]",None,"Aseguradora [POR DEFINIR]","F1, F4"),
 ("F","F · Seguridad, salud y comunicaciones","Comunicación satelital o radio para brigadas en zonas sin cobertura","equipo-mes",3*3,"3 equipos × 3 meses [SUPUESTO]",None,"[POR DEFINIR]","F1, F4"),
 ("F","F · Seguridad, salud y comunicaciones","Botiquines, repelente, vacunación y kits de emergencia","juego",3,"1 por brigada",None,"[POR DEFINIR]","F1"),
 ("G","G · Logística, datos y tecnología","Servicio de base de datos y aplicativo (PostgreSQL, Render, KoboToolbox) durante 2027","mes",8,"Noviembre de 2026 a junio de 2027",None,"Planes vigentes [POR DEFINIR]","F0-F5"),
 ("G","G · Logística, datos y tecnología","Impresión de fichas, planos y expedientes AX-04; copias certificadas (SUNARP, COFOPRI)","global",1,"[POR DEFINIR]",None,"[POR DEFINIR]","F2-F4"),
 ("G","G · Logística, datos y tecnología","Pasajes Lima - Piura - Lima del equipo de gabinete para misiones de coordinación","pasaje",12,"6 misiones × 2 personas [SUPUESTO]",None,"Tarifa vigente [POR DEFINIR]","F0-F5"),
]
CONTINGENCIA_PCT = 0.10  # [SUPUESTO] reserva de contingencia sobre el subtotal; ajustar según la política de la ANIN
MARGEN_ESTUDIOS = 15810000  # Matriz v1.6: S/ 356.63 M de inversión + O&M frente al techo de S/ 372.44 M

# ---------------------------------------------------------------- Cambios de alcance y contrataciones
CAMBIOS = [
 # tema, situación actual, cambio propuesto, efecto en el Perfil, decisión requerida de, tipo
 ("Liberación de áreas por fases","Ruta F-LA prevista para 117 bloques y 60 lotes SUS antes de fijar metas; 0 actas conformes al 03/11/2026.","Fase 1 (nov 2026): bloques piloto y de mayor riesgo. Fase 2 (abr-may 2027): resto. Si en mayo de 2027 persisten bloques sin acta, el Perfil fija metas sobre los bloques con libre disponibilidad verificada y presenta el resto como cartera condicionada a saneamiento en el expediente técnico.","Metas físicas sustentadas en superficie disponible; posible reducción del área de intervención directa dentro del rango de 10,000 a 15,000 ha.","SESDI / OPMI ANIN","Alcance"),
 ("Validación del E6 (fase 7)","Validación ciega de 567 puntos prevista en campo.","Validación remota con imágenes de alta resolución post-evento durante la temporada de lluvias; verificación de campo de una submuestra en abril y mayo de 2027.","Mantiene el calendario del E6 sin exponer brigadas; menor precisión en bloques con nubosidad.","Coordinación del Perfil","Método"),
 ("Pendiente de campo (E1)","Medición con clinómetro en parcelas prevista para todos los bloques.","Análisis del DEM en UTM 17S y verificación de campo solo en bloques cercanos al límite de 75 %.","Sin efecto en el indicador; reduce días de campo.","Equipo SIG","Método"),
 ("Cárcavas y diques","85 bloques sin cárcava registrada; levantamiento F-DT-02 previsto en campo.","Inventario post-evento con dron o imágenes (las cárcavas activadas por el FEN 2026-2027 serán más visibles) y verificación en el campo residual; meta de diques en el Perfil como rango con metrado preliminar.","Metrado de cárcavas más completo; meta de diques sujeta al expediente técnico.","Especialista en suelos / SIG","Alcance y método"),
 ("Escenario climático del EVAR","Desencadenante fijo, sin escenarios de cambio climático.","Incorporar el escenario SSP5-8.5 y los registros del FEN 2026-2027 (SENAMHI, SINPAD) como evidencia del peligro y de los daños.","Fortalece la justificación del proyecto y los beneficios por daños evitados (E13).","Especialista en riesgos","Alcance"),
 ("Zanjas de infiltración frente al E5","2,250.56 km de zanjas en el plano frente a 7.44 km aptos según suelo, pendiente y cobertura.","Sustituir zanjas por terrazas de formación lenta o barreras vivas donde la infiltración no lo sustente; dejar las zanjas solo en Canchaque y El Ingenio con restricción.","Cambio en metas y costos de la línea marrón; Matriz v2.0 dentro del techo.","Equipo formulador","Alcance técnico"),
 ("Bloques con condiciones límite","Bloque 8 con el 25.4 % de su área V7; 5 bloques con más de 19 predios; 5 bloques con P < 400 mm.","Decidir en diciembre de 2026 mantener, redefinir o descartar cada caso con los criterios de idoneidad y tamizaje predial.","Catálogo definitivo con menos incertidumbre; posible reducción de superficie.","Coordinación del Perfil","Alcance"),
 ("Evaluación social","Beneficios no valorados; EVAR sin daños.","Evaluación costo-efectividad como método principal (tipología de servicios ecosistémicos, R.M. N° 178-2019-MINAM) y, si el inventario de daños lo permite, VANS y TIR social complementarios.","Cumple la Guía DGPMI con la información disponible.","Economista / SESDI","Método"),
 ("Contratación: culminación del E4","Informe de avance al 40 %; faltan lluvias de diseño, IDF y HEC-HMS.","Asegurar el saldo contractual o la adenda para culminar en diciembre de 2026, con entrega digital y sustento de campo ya existente.","Cierre del Capítulo de Identificación.","Administración DIME","Contratación estrictamente necesaria"),
 ("Contratación: especialista en evaluación social (E13)","Entregable al 0 %; requiere competencia específica en evaluación social de proyectos de servicios ecosistémicos.","Locación de servicios de enero a abril de 2027, con TDR alineados a la Guía DGPMI y al Anexo 2 (GdR-CCC).","E13 y E14 en plazo.","Administración DIME / SESDI","Contratación estrictamente necesaria"),
 ("Contratación: servicio de dron e imágenes","No contratado.","Servicio de vuelo de dron (15 nanocuencas priorizadas) e imágenes de alta resolución post-evento.","Cárcavas, drenaje y validación remota del E6.","Administración DIME","Contratación estrictamente necesaria"),
 ("Contratación: transporte 4x4, SCTR, EPP y comunicaciones","Según disponibilidad actual de la DIME [POR VERIFICAR].","Asegurar antes del 09/11/2026 para la ventana de campo y renovar para abril de 2027.","Seguridad de brigadas y ejecución del campo.","Administración DIME","Contratación estrictamente necesaria"),
 ("Continuidad del equipo en 2027","Contratos del personal y servicios con vencimiento al 31/12/2026 [POR VERIFICAR].","Programar en noviembre las renovaciones o nuevas contrataciones y las certificaciones presupuestales del PIA 2027 para evitar la interrupción de enero.","Sin pérdida de capacidad en la Fase 3.","Administración DIME / Jefatura","Contratación estrictamente necesaria"),
]

# ---------------------------------------------------------------- Matriz de riesgos
# (id, categoría, riesgo, causa, efecto en el Perfil, prob 1-5, imp 1-5, medidas prospectivas/correctivas, medidas reactivas, responsable, indicador/disparador, entregables)
RIESGOS = [
 ("R01","Climático","Lluvias intensas y crecidas de quebradas impiden el acceso a los bloques (vías y trochas cortadas)","FEN 2026-2027 con lluvias por encima de lo normal en la sierra norte occidental (ENFEN N° 19-2026)","Suspensión del tamizaje predial, de las fichas sociales y de las pruebas de infiltración; atraso de E2, E6 y E10",5,5,"Concentrar el campo en noviembre de 2026; programar el campo residual para abril y mayo de 2027; priorizar bloques por riesgo y accesibilidad; usar imágenes y dron como sustituto parcial","Activar ventanas de campo cortas con autorización del protocolo; reprogramar por bloque; actualizar el tablero semanal","Coordinación del Perfil","Avisos SENAMHI nivel 3 o 4; reportes COER de vías interrumpidas en los 15 distritos","E2, E6, E10"),
 ("R02","Seguridad","Exposición de las brigadas a movimientos en masa, huaicos y crecidas durante el campo","El 93.9 % de los bloques tiene peligro alto o muy alto por movimientos en masa (EVAR); quebradas activas en temporada de lluvias","Accidente, paralización del campo, responsabilidad institucional",3,5,"Protocolo de seguridad con umbrales de suspensión (aviso SENAMHI, lluvia acumulada), horarios de retorno, EPP, SCTR, comunicación satelital, vehículos 4x4, no pernoctar en cauces ni laderas inestables","Plan de evacuación por bloque; coordinación con COER y centros de salud; suspensión inmediata ante alerta","Responsable de seguridad / jefes de brigada","Aviso meteorológico vigente; lluvia acumulada en 24 h sobre el umbral; reporte de evento en el distrito","E2, E5, E6"),
 ("R03","Institucional","Autoridades locales y regionales concentradas en la emergencia (COER, COEL, plataformas de Defensa Civil)","D.S. N° 149-2026-PCM asigna a gobiernos regionales y locales las acciones de reducción del riesgo, respuesta y rehabilitación","Asambleas F-LA-01, actas y reuniones postergadas; menor disponibilidad de la DRA Piura y de las municipalidades para la información predial",5,4,"Agendar reuniones en noviembre; delegar en autoridades comunales y tenientes gobernadores; coordinar con el COER la agenda del proyecto como medida de reducción del riesgo","Reprogramar asambleas; reuniones virtuales con autoridades; actas con firmas en campo posteriores","Especialista social","Cancelación de dos o más reuniones programadas en una semana","E2"),
 ("R04","Político-administrativo","Cambio de autoridades regionales, provinciales y distritales el 01/01/2027 (segunda vuelta regional en diciembre)","ERM 2026; nuevos interlocutores en el GORE Piura, 3 provincias y 15 distritos","Pérdida de acuerdos informales, nueva socialización, demora en actas de compromiso y en el apoyo logístico local",5,4,"Documentar todos los acuerdos antes del 31/12/2026 (actas y oficios); preparar carpetas de transferencia por distrito; mapa de actores actualizado con los electos; agenda de re-socialización en enero y febrero","Reuniones de presentación con las autoridades electas en la primera quincena de enero; ratificación de actas","Coordinación / Especialista social","Resultados oficiales ONPE-JNE; juramentación del 01/01/2027","E2, E11"),
 ("R05","Social-comunal","Renovación de juntas directivas de comunidades campesinas en el periodo del plan","Mandatos de dos años (Ley N° 24656); asambleas eleccionarias frecuentes entre diciembre y enero [POR VERIFICAR por comunidad]","Actas de asamblea deben ratificarse; nuevos directivos desconocen el proyecto; 32.3 % de la superficie neta es comunal (83.9 % en Ayabaca)",4,4,"Levantar el calendario electoral de cada comunidad con la DRA Piura; cerrar actas con las juntas vigentes en noviembre; inscribir acuerdos en libros de actas comunales","Re-socialización con las nuevas juntas; ratificación de actas en la primera asamblea ordinaria","Especialista social / legal-predial","Convocatoria a asamblea eleccionaria en una comunidad del ámbito","E2"),
 ("R06","Institucional ANIN","Reasignación del equipo formulador de la DIME a intervenciones de emergencia","La DIME es la dirección de intervenciones multisectoriales y de emergencia; el decreto moviliza a las entidades del Gobierno Nacional","Reducción de la capacidad del equipo; atraso de gabinete y campo",4,5,"Formalizar ante la jefatura la dedicación del equipo núcleo al Perfil durante el plan; identificar reemplazos y prioridades; registrar el Perfil como intervención de reducción del riesgo","Redistribuir tareas; priorizar la ruta crítica; solicitar refuerzo o ampliación de plazos","Jefatura DIME / SESDI","Memorando de reasignación o encargo de emergencia a un integrante del equipo","Todos"),
 ("R07","Presupuestal-administrativo","Cierre del año fiscal 2026 y transición al PIA 2027; vencimiento de contratos de personal y servicios al 31/12/2026","Calendario presupuestal del sector público; certificaciones y conformidades de cierre","Interrupción del equipo y de servicios en enero de 2027; demoras en pagos a consultores",4,4,"Programar en noviembre las certificaciones y la contratación 2027; adelantar conformidades; priorizar los servicios de la ruta crítica (E4, E13, dron, transporte)","Contratación directa de servicios menores; redistribución de tareas en enero","Administración DIME","Certificaciones no emitidas al 15/12/2026; contratos sin renovación al 20/12/2026","Todos"),
 ("R08","Contractual","Retraso de consultorías en curso (hidrología E4) o dificultad para contratar nuevos servicios durante la emergencia","Dependencia de campo o de datos externos (SENAMHI); mercado de proveedores demandado por la emergencia","E4 no cierra en diciembre; falta de insumos para el diseño de obras y para el E13",3,4,"Seguimiento quincenal del E4; entregas parciales; TDR listos para los nuevos servicios antes del 15/11/2026","Adenda de plazo; uso de PISCO y estaciones ya procesadas; alternativas de proveedores","Coordinación / Administración DIME","Entrega parcial del E4 no recibida al 30/11/2026","E4, E10, E13"),
 ("R09","Técnico","Datos de campo incompletos impiden fijar las metas físicas (superficie con libre disponibilidad)","107 bloques y 60 lotes SUS sin tamizaje predial; 32 bloques sin ficha social; campo limitado por el clima","E10 sin metas firmes; Perfil con cartera condicionada; observaciones de la OPMI",5,5,"Priorizar bloques por riesgo, superficie y tenencia; ruta F-LA por fases; cartera condicionada explícita en el Perfil; criterio de corte acordado con la SESDI","Fijar metas sobre bloques con acta conforme y presentar el resto como condicionada al expediente técnico","Coordinación del Perfil / SESDI","Menos del 50 % de la superficie con acta conforme al 14/05/2027","E1, E2, E10"),
 ("R10","Técnico","Cambios físicos en los bloques por el FEN 2026-2027 (activación de cárcavas, deslizamientos, pérdida de cobertura, cambios de cauce)","Evento climático extraordinario sobre laderas con peligro alto y muy alto","Desactualización del catálogo V8, del MSAVI 2024 y del EVAR; metas y costos por revisar; también oportunidad de evidencia",4,3,"Monitoreo de eventos (SENAMHI, SINPAD, COER) por distrito; imágenes post-evento; protocolo de actualización de capas","Actualizar el catálogo, el MSAVI y el inventario de cárcavas tras el evento; ajustar metas en la Matriz v2.0","Equipo SIG / Especialista en riesgos","Registro SINPAD de emergencia en un distrito del ámbito","E1, E6, E8, E10"),
 ("R11","Social","Expectativas de ayuda inmediata frente a un proyecto de largo plazo; condicionamientos o rechazo","Población afectada por la emergencia; 8 de 76 autoridades condicionan su aceptación","Rechazo de asambleas, actas condicionadas, conflictos en el tamizaje predial",3,3,"Mensajes claros sobre fase y plazos; articular con las acciones de emergencia del COER sin asumir funciones ajenas; registrar compromisos realistas","Mediación con autoridades comunales; postergar asambleas en localidades afectadas","Especialista social","Acta condicionada o rechazo registrado en F-DS-03","E2"),
 ("R12","Seguridad ciudadana","Inseguridad en el corredor logístico Piura - Chulucanas y provincias costeras","Estado de emergencia por criminalidad en Piura, Sullana, Paita y Talara (D.S. N° 147-2026-PCM)","Restricciones de horario, riesgo para el personal y los equipos",2,3,"Desplazamientos diurnos; rutas y alojamientos seguros; coordinación con la PNP","Suspender desplazamientos ante incidentes","Coordinación","Reporte de incidente en el corredor","Todos"),
 ("R13","Logístico","Escasez o encarecimiento de transporte 4x4, combustible y alojamiento en la zona de emergencia","Demanda de las entidades de respuesta; vías dañadas","Mayor costo y menor disponibilidad para las brigadas",4,3,"Contratar o reservar antes del 09/11/2026; convenios con municipalidades para alojamiento","Reasignar vehículos; priorizar bloques por accesibilidad","Administración DIME","Cotización por encima del presupuesto; sin disponibilidad de vehículos","E2, E5, E6"),
 ("R14","Tecnológico","Pérdida de conectividad o de datos del aplicativo IN Piura y de KoboToolbox","Cortes de energía y telecomunicaciones por el evento; zonas sin cobertura","Fichas sin sincronizar; pérdida de registros",3,3,"Formularios Kobo sin conexión; copias de seguridad diarias de PostgreSQL; plantillas Excel ANIN como respaldo; bitácora","Restaurar desde copia; digitar desde plantillas","Equipo SIG / TI","Fallo de sincronización por más de 48 h","E2"),
 ("R15","Normativo","Cambios en la normativa Invierte.pe o en parámetros de evaluación social durante 2027; observaciones de la OPMI o la DGPMI","Actualizaciones del MEF; revisión del Perfil en junio de 2027","Reformulación de capítulos; atraso de la viabilidad",2,3,"Seguimiento de la normativa; reuniones de avance con la OPMI antes de la entrega; cumplimiento de la Guía DGPMI y del Anexo 2","Levantamiento de observaciones con plazo acordado","SESDI / OPMI","Publicación de norma modificatoria; informe de observaciones","E13, E14, E15"),
 ("R16","Presupuestal","Incremento de costos unitarios (fletes, jornales, insumos) por la emergencia y cartera que supera el techo","Matriz v1.6 al 95.8 % del techo; margen de S/ 15.81 M para estudios; presión de precios por el FEN","Cartera fuera del techo de S/ 372.44 M; recorte de metas",3,4,"Actualizar precios en la Matriz v2.0 con índices oficiales; escenarios de techo; priorización por tipología (T1 concentra el 76.6 % de la inversión)","Ajustar metas o alternativa; sustentar incremento ante la SESDI","Economista","Inversión + O&M por encima del 100 % del techo en la Matriz v2.0","E10, E13"),
 ("R17","Salud","Enfermedades asociadas a lluvias (dengue, leptospirosis, infecciones respiratorias) en brigadas y población","Temporada de lluvias en zonas endémicas de Piura","Bajas en el equipo; comunidades sin disposición para reuniones",4,3,"Vacunación y prevención; repelentes; botiquín; seguro; evitar zonas con brotes","Reemplazo de personal; reprogramación","Responsable de seguridad y salud","Alerta epidemiológica DIRESA Piura en un distrito del ámbito","E2, E6"),
 ("R18","Información","Dificultad para obtener información registral y catastral (SUNARP, COFOPRI, DRA Piura, Ministerio de Cultura, INGEMMET) durante la emergencia","Oficinas en modo de emergencia; trámites presenciales","F-LA-05 y expedientes AX-04 demorados; polígono comunal ID 2790 sin identificar",3,3,"Solicitudes formales en noviembre; uso de plataformas digitales (SUNARP en línea, GEOLLAQTA, GEOCATMIN); oficios de la ANIN","Visitas programadas en ventanas; acuerdos interinstitucionales","Especialista legal-predial","Respuesta no recibida en 30 días","E2"),
 ("R19","Gestión","Pérdida de continuidad por cambios de jefaturas en la ANIN, la DIME o la SESDI, o por rotación del equipo","Contexto político nacional 2026-2027 y presión de la emergencia","Decisiones pendientes sin dueño; cambio de criterios de revisión",2,3,"Actas de decisiones técnicas; repositorio documental versionado; este plan como referencia aprobada","Reunión de inducción a nuevas jefaturas; ratificación del plan","Coordinación del Perfil","Cambio de titular en la cadena de aprobación","Todos"),
 ("R20","Oportunidad","El FEN 2026-2027 genera evidencia directa del peligro y de los daños en la zona de origen","Evento climático sobre el ámbito del proyecto","Mejor sustento del problema (E7), del EVAR (E8), de la hidrología (E4) y de los beneficios por daños evitados (E13)",4,-3,"Protocolo de registro de eventos (fecha, lugar UTM 17S, tipo, daños, fuente) con el COER y el SINPAD; imágenes antes y después; coordinación con SENAMHI","Misión de evaluación post-evento en abril de 2027","Especialista en riesgos / SIG","Evento registrado en un distrito del ámbito","E4, E7, E8, E13"),
]

# ---------------------------------------------------------------- Aspectos administrativos y de gestión
ADMIN = [
 # tema, situación, efecto esperado sobre el Perfil, acciones del plan, plazo, responsable
 ("Autoridades regionales (GORE Piura)","Segunda Elección Regional en diciembre de 2026 [fecha POR VERIFICAR]; nueva gestión desde el 01/01/2027.","Nuevo gobernador y consejo regional; cambio en la Gerencia de Recursos Naturales, la DRA Piura y el COER; posible revisión de prioridades y convenios.","Oficio de presentación y reunión de trabajo en enero de 2027; carpeta de transferencia del proyecto (ámbito, bloques, avances, compromisos); articulación del Perfil con el Plan Regional de GRD y la emergencia.","Ene - Feb 2027","Jefatura SESDI / Coordinación"),
 ("Autoridades provinciales (Morropón, Huancabamba, Ayabaca)","Alcaldes provinciales electos el 04/10/2026 [POR VERIFICAR]; asumen el 01/01/2027.","Nuevos interlocutores para la coordinación territorial, el apoyo logístico y las plataformas de Defensa Civil.","Cerrar actas de compromiso con la gestión saliente antes del 31/12/2026; presentación a los electos en enero; reuniones de trabajo por provincia.","Nov 2026 - Feb 2027","Especialista social / Coordinación"),
 ("Autoridades distritales (15 distritos)","Alcaldes distritales electos el 04/10/2026 [POR VERIFICAR]; asumen el 01/01/2027.","Cambio de regidores, gerencias y agentes municipales que convocan asambleas y validan información local; posible pérdida de acuerdos informales.","Carpetas por distrito (mapa de bloques, centros poblados, fichas aplicadas, compromisos); actas de ratificación; agenda de asambleas F-LA-01 coordinada con las nuevas gestiones.","Dic 2026 - Mar 2027","Especialista social"),
 ("Autoridades comunales","Juntas directivas con mandatos de dos años; calendario eleccionario por comunidad [POR VERIFICAR] con la DRA Piura y SUNARP.","Las actas de asamblea y los acuerdos de libre disponibilidad requieren ratificación por la nueva junta; riesgo de reapertura de acuerdos.","Levantar el calendario de cada comunidad; cerrar actas con las juntas vigentes en noviembre; inscribir acuerdos en libros de actas; ratificación en la primera asamblea de 2027.","Nov 2026 - Mar 2027","Especialista social / legal-predial"),
 ("Gobierno Nacional y ANIN","Gestión nacional desde el 28/07/2026; Estado de Emergencia nacional por el FEN; la DIME es la dirección de intervenciones de emergencia de la ANIN.","Prioridad institucional en la respuesta; posible reasignación de personal; cambios de titulares en la cadena de aprobación [POR VERIFICAR].","Formalizar la dedicación del equipo núcleo; presentar el Perfil como medida de reducción del riesgo en la zona de origen; informes quincenales a la SESDI.","Nov 2026 - Jun 2027","Jefatura DIME / SESDI"),
 ("Presupuesto institucional y contratos","Cierre del ejercicio 2026; PIA 2027; contratos de personal y servicios con vencimiento al 31/12/2026 [POR VERIFICAR].","Interrupción de enero si no se programan las certificaciones y contrataciones.","Programar en noviembre las certificaciones 2026 y 2027, la renovación de contratos y los TDR de los servicios estrictamente necesarios (sección 5).","Nov - Dic 2026","Administración DIME"),
 ("Banco de Inversiones (Invierte.pe)","Idea registrada (CUI 2669244); Formato N° 07-A por actualizar con el avance del Perfil.","El cambio de fecha de culminación y de monto debe reflejarse en el registro; la viabilidad depende de la revisión de la OPMI.","Actualizar el Formato N° 07-A al cerrar la Identificación y al cerrar el Perfil; reunión previa con la OPMI; sustentar el ajuste de cronograma con este plan.","Dic 2026 y May 2027","SESDI / OPMI ANIN"),
 ("Coordinación con el SINAGERD (COER Piura, INDECI, CENEPRED)","Emergencia declarada; el INDECI coordina y supervisa las acciones del decreto.","Acceso a reportes de eventos, vías y daños; posibilidad de registrar el proyecto como intervención de reducción del riesgo; agenda compartida con las autoridades.","Oficio de coordinación al COER Piura y al INDECI; acceso al SINPAD; participación en las plataformas provinciales cuando corresponda.","Nov 2026 - Abr 2027","Especialista en riesgos / Coordinación"),
 ("Entidades de información predial y de exclusión","DRA Piura, SUNARP, COFOPRI, Ministerio de Cultura (SIGDA), INGEMMET (derechos mineros en trámite en los bloques 8, 71, 22 y M17B1).","Demoras en respuestas durante la emergencia y la transición de gestión.","Solicitudes formales en noviembre de 2026; seguimiento mensual; uso de plataformas digitales.","Nov 2026 - May 2027","Especialista legal-predial"),
 ("Comunicación y transparencia","Población y autoridades en situación de emergencia.","Confusión entre el proyecto y la ayuda humanitaria; expectativas.","Mensajes institucionales únicos (fase de estudio, plazos, qué sí y qué no hace el proyecto); voceros definidos; materiales por distrito.","Nov 2026 - Mar 2027","Coordinación / Especialista social"),
]

SEGUIMIENTO = [
 ("Tablero semanal del plan","Avance por actividad del cronograma, bloques con campo ejecutado, actas conformes, alertas climáticas","Semanal (lunes)","Coordinación del Perfil"),
 ("Informe quincenal a la SESDI","Avance ponderado por entregable, riesgos activados, decisiones requeridas","Quincenal","Coordinación del Perfil"),
 ("Revisión de la matriz de riesgos","Probabilidad, impacto, disparadores y medidas; nuevos riesgos","Mensual y ante cada evento","Especialista en riesgos"),
 ("Actualización del plan","Cronograma, presupuesto y alcance según el escenario vigente","Al cierre de cada fase (13/11, 30/11, 31/12/2026; 31/03, 31/05/2027)","SESDI / Coordinación"),
 ("Protocolo de seguridad","Autorización diaria de salidas de campo según avisos SENAMHI y reportes COER","Diaria en fases de campo","Responsable de seguridad"),
]

def avance_ponderado(ents):
    return sum(e[3]*e[4] for e in ents)/sum(e[3] for e in ents)

def nivel(p, i):
    s = p*abs(i)
    return "Muy alto" if s >= 20 else "Alto" if s >= 12 else "Medio" if s >= 6 else "Bajo"

DATOS = dict(meta=META, corte=CORTE, distritos=DISTRITOS, decreto=DECRETO, clima=CLIMA, electoral=ELECTORAL,
             entregables=ENTREGABLES, capitulos=CAPITULOS, programacion_original=PROGRAMACION_ORIGINAL,
             pendientes=PENDIENTES, cronograma=CRONOGRAMA, fases=FASES, escenarios=ESCENARIOS,
             presupuesto=PRESUPUESTO, contingencia_pct=CONTINGENCIA_PCT, margen_estudios=MARGEN_ESTUDIOS,
             viatico_ref=VIATICO_REF, cambios=CAMBIOS, riesgos=RIESGOS, admin=ADMIN, seguimiento=SEGUIMIENTO,
             avance=round(avance_ponderado(ENTREGABLES), 2))

if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "plan_contingencia.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(DATOS, f, ensure_ascii=False, indent=1)
    print("avance ponderado", DATOS["avance"], "| riesgos", len(RIESGOS), "| actividades", len(CRONOGRAMA), "| pendientes", len(PENDIENTES), "| partidas", len(PRESUPUESTO))
