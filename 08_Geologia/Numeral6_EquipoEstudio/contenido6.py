# -*- coding: utf-8 -*-
"""Contenido del numeral 6. EQUIPO DE ESTUDIO — Entregable 3, Estudio de Geologia.
Proyecto IN Piura (CUI 2669244) — ANIN / DIME / SESDI."""

PARRAFOS_61 = [
    "El Estudio de Geología del Proyecto IN Piura se elabora bajo la modalidad de "
    "administración directa, con personal profesional de la Subdirección de Estudios de "
    "Inversión (SESDI) de la DIME. En consecuencia, los roles previstos en el flujograma "
    "del Entregable 3 —Especialista en Geología, Especialista SIG-CAD, Especialista "
    "Marrón, Gerente de Diseño y Gerente de Proyecto— son asumidos por los profesionales "
    "del equipo de estudio de la ANIN, conforme a la equivalencia consignada en la tercera "
    "fila del cuadro.",
    "El entregable se organiza de manera modular en tres (03) volúmenes provinciales "
    "—Volumen I: Morropón; Volumen II: Huancabamba; Volumen III: Ayabaca—, aplicándose a "
    "cada uno la misma secuencia metodológica de tres actividades: (I) recopilación y "
    "revisión de información secundaria, (II) procesamiento y análisis de información y "
    "(III) elaboración del informe final del entregable y sus anexos.",
]

LEYENDA_TABLA = ("Tabla 6.1. Asignación de actividades y responsables para la elaboración "
                 "del Estudio de Geología (Entregable 3).")

ENCABEZADO_PROF = [
    "Ing. Albert Flores Mendoza\nIng. Geógrafo",
    "Ing. Christian Camacho Aponte\nIng. Agrícola",
    "Eco. Luis García Iguía\nEconomista",
    "Ing. Héctor S. Cahuas Miller\nIng. Forestal",
    "Ing. Pedro Talledo Hernández\nIng. Forestal",
    "Ing. Alejandro Falconi Valdivia\nIngeniero",
]

ENCABEZADO_ROL = [
    "Analista SIG (Especialista SIG-CAD)",
    "Asistente en SIG (Apoyo SIG-CAD)",
    "Formulador de Proyectos IN",
    "Responsable del Proyecto IN Piura (Gerente de Diseño)",
    "Especialista IM – asume la función de Esp. en Geología",
    "Coordinador de Proyectos IN (Gerente de Proyecto)",
]

# Cada entrada: (etiqueta, marcas) ; marcas = None indica banda de ACTIVIDAD.
# marcas = lista de 6 booleanos en el orden de ENCABEZADO_PROF.
FILAS = [
 ("ACTIVIDAD I. RECOPILACIÓN Y REVISIÓN DE INFORMACIÓN SECUNDARIA (Semanas 9 – 13)", None),
 ("1.1 Definición del ámbito de estudio y estructuración modular del entregable en tres (03) "
  "volúmenes provinciales —Morropón, Huancabamba y Ayabaca—, a partir de las unidades de "
  "intervención del Entregable 1.", "X..XX."),
 ("1.2 Compilación de estudios geológicos, geomorfológicos y de peligros existentes en la zona "
  "de estudio (antecedentes técnicos e institucionales).", "XX..X."),
 ("1.3 Compilación de cartografía temática publicada del ámbito de estudio: Carta Geológica "
  "Nacional 1:50 000 (GEOCATMIN – INGEMMET), geomorfología y peligros.", "XX..X."),
 ("1.4 Descarga y preparación del modelo digital de terreno, imágenes satelitales y base "
  "cartográfica; estructuración de la geodatabase del entregable en UTM WGS 84 – Zona 17S.", "XX...."),
 ("1.5 Revisión y jerarquización del marco normativo y de las guías técnicas aplicables al "
  "entregable (Invierte.pe, NTP CE.020, Guía Técnica de IN, ICS).", "..XXX."),
 ("1.6 Elaboración del flujo de proceso conceptual del entregable y del glosario de términos.", "X..XX."),
 ("1.7 Conformación del equipo de estudio, elaboración del cuadro de responsables por actividad "
  "y del organigrama del entregable.", "...X.X"),
 ("ACTIVIDAD II. PROCESAMIENTO Y ANÁLISIS DE INFORMACIÓN (Semanas 14 – 16)", None),
 ("2.1 Formulación del objetivo general y de los objetivos específicos de cada volumen provincial.", "...XX."),
 ("2.2 Delimitación de unidades geológicas: intersección geométrica de la capa de Bloques v5 con "
  "la cobertura de unidades geológicas del INGEMMET (QGIS 3.34).", "XX..X."),
 ("2.3 Control de calidad geoespacial: verificación del cierre de superficies por bloque "
  "(Σ unidades geológicas = 100 % del polígono) en los tres volúmenes.", "XX.X.."),
 ("2.4 Caracterización litoestratigráfica por bloque: número de unidades, unidad dominante y su "
  "proporción, clasificación por tipo de unidad y posición cronoestratigráfica.", "X...X."),
 ("2.5 Caracterización del contexto geológico regional de la cuenca alta del río Piura: "
  "formaciones aflorantes, marco geodinámico y sistemas de fallas.", "X...X."),
 ("2.6 Análisis geológico estructural y geomorfológico; descripción de las geoformas asociadas "
  "al sustrato y de los procesos geodinámicos activos.", "X...X."),
 ("2.7 Evaluación de peligros geológicos: identificación de peligros por bloque y análisis de "
  "las posibles causas condicionantes y desencadenantes.", "XX.XX."),
 ("2.8 Construcción del Índice de Susceptibilidad Litológica (ISL-MM e ISL-EH) y jerarquización "
  "de los bloques por provincia.", "X..XX."),
 ("2.9 Diagnóstico de amenazas y vulnerabilidad; identificación de las condicionantes geológicas "
  "para el diseño de la infraestructura natural y de la infraestructura gris complementaria.", "...XX."),
 ("2.10 Generación de productos cartográficos: mapas geológicos por bloque, por distrito y por "
  "provincia, para los tres volúmenes.", "XX..X."),
 ("2.11 Consolidación de las tablas de resultados (Hojas 1 a 6) y de la base de datos "
  "alfanumérica del entregable.", "XX.X.."),
 ("2.12 Articulación de los resultados geológicos como insumo de los Entregables 4 (Hidrología), "
  "5 (Suelos), 6 (Ecosistemas), 8 (Análisis de riesgos) y 10 (metas, costos y presupuesto).", "..XXX."),
 ("ACTIVIDAD III. ELABORACIÓN DEL INFORME FINAL DEL ENTREGABLE, ANEXOS Y PRESENTACIÓN (Semana 16)", None),
 ("3.1 Redacción del Informe Final — Volumen I: Provincia de MORROPÓN.", "X...X."),
 ("3.2 Redacción del Informe Final — Volumen II: Provincia de HUANCABAMBA.", "X...X."),
 ("3.3 Redacción del Informe Final — Volumen III: Provincia de AYABACA.", "X...X."),
 ("3.4 Redacción del resumen ejecutivo, conclusiones y recomendaciones de cada volumen provincial.", "...XX."),
 ("3.5 Compaginación de anexos: archivos digitales alfanuméricos, geoarchivos (mapas en PDF y "
  "JPG, nativos y geodatabase) y panel fotográfico georreferenciado.", "XX...."),
 ("3.6 Revisión técnica interna e integración de observaciones en los tres volúmenes (control de "
  "consistencia entre provincias).", "...XX."),
 ("3.7 Revisión y conformidad técnica del Informe Final del Entregable 3 y sus anexos.", "...X.."),
 ("3.8 Aprobación y presentación del Informe Final del Entregable 3 a través del Sistema de "
  "Gestión Documental (SGD) de la ANIN.", "...X.X"),
]

NOTAS = [
    "Nota 1: La marca (X) identifica al profesional responsable o participante directo en la "
    "ejecución de la actividad o tarea.",
    "Nota 2: Las actividades I, II y III se ejecutan de forma iterativa para cada uno de los tres "
    "volúmenes provinciales; las tareas 3.1, 3.2 y 3.3 corresponden a la redacción diferenciada "
    "de cada volumen.",
    "Nota 3: Las horas hombre efectivas por profesional y actividad se consignan en el registro "
    "de dedicación del equipo de estudio de la SESDI.",
]

TEXTO_62 = ("Organigrama del equipo que participa en la elaboración del Estudio de Geología, con "
            "la línea de dependencia respecto de la Gerencia de Diseño (Responsable del Proyecto "
            "IN Piura) y de la Gerencia de Proyecto (Coordinador de Proyectos IN).")

LEYENDA_FIGURA = ("Figura 6.1. Organigrama del equipo de estudio del Entregable 3 – Estudio de "
                  "Geología, Proyecto IN Piura.")
