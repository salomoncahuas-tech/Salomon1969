// Anexo N° 2 — Actividad c): seguimiento técnico del Estudio de Suelos a Nivel de Reconocimiento
const L = require("./lib");
const { para, b, i, run, h1, h2, h3, guion, bullet, num, fuente, tabla, vacio, centro, W, WH, C, R, fmt, figura } = L;
const { seccionCaratula } = require("./caratula");
const DEF = require("./anexos_def");
const D = require("./datos.json");
const OUT = process.argv[2];
const FIG = "/tmp/claude-0/-home-user-Salomon1969/03b82ed8-c0e9-54ec-93c9-bff5f93b575b/scratchpad/build/fig/";
const H = D.prov["Huancabamba"], M = D.prov["Morropón"];
const N = L.numerador("A2.");
const PIE = "Anexo N° 2 · Informe N° 008-2026-HSCM · Segundo Entregable · OS N° 4062-2026";

const x = [];
x.push(centro([b("ANEXO N° 2", { size: 26 })], { spacing: { after: 60 } }));
x.push(centro([b("SEGUIMIENTO TÉCNICO DEL ESTUDIO DE SUELOS A NIVEL DE RECONOCIMIENTO DEL PROYECTO IN PIURA", { size: 22 })], { spacing: { after: 60 } }));
x.push(centro([i("Revisión de entregables, observaciones, conformidades y aplicación al componente verde — periodo del 26/09 al 27/10/2026", { size: 19 })], { spacing: { after: 240 } }));

x.push(h1("1. OBJETIVO Y ALCANCE"));
x.push(para("El presente anexo desarrolla la actividad c) del numeral 5.1 de los Términos de Referencia durante el periodo del segundo entregable. Documenta el seguimiento técnico del Estudio de Suelos a Nivel de Reconocimiento —revisión técnica, formulación y levantamiento de observaciones y estado de las conformidades respecto del plan de trabajo, las actividades de campo y los informes presentados—, ejercido en coordinación con el Especialista en Infraestructura Marrón, y verifica que la información generada cumpla los requerimientos técnicos para su aplicación en el componente verde del proyecto IN Piura."));

x.push(h1("2. DATOS DEL SERVICIO SUPERVISADO"));
x.push(N.cuadro("Ficha del servicio supervisado"));
x.push(tabla([{ w: 3000, bold: true }, { w: 6638 }], ["CAMPO", "DETALLE"], [
  ["Servicio", "Estudio de Suelos a Nivel de Reconocimiento del Proyecto de Infraestructura Natural Piura (E5 del Perfil; código de entregable 999990-ANIN001-000-XX-SC-GE-000001-R00)"],
  ["Contratista", "GeoSIG Ingenieros E.I.R.L."],
  ["Procedimiento / contrato", "CP-ABR-6-2026-ANIN/DAI-1; contrato suscrito el 17/08/2026"],
  ["Plazo y estructura", "75 días calendario; cuatro fases (gabinete inicial, campo, laboratorio y gabinete final) y cuatro entregables (20 %, 30 %, 20 % y 30 %)"],
  ["Ámbito", "117 bloques de intervención vigentes, 12,270.24 ha; 15 distritos de Morropón, Ayabaca y Huancabamba; 500 a 3,500 msnm"],
  ["Marco técnico", "D.S. N° 013-2010-AG (levantamiento de suelos); D.S. N° 005-2022-MIDAGRI (CUM); Soil Taxonomy (USDA, 13.ª ed.) y WRB/FAO (2022); Guías del Proyecto Infraestructura Natural para la Seguridad Hídrica (Guías 5 y 8); USLE (Factor K de Wischmeier y Smith; Factor R con PISCO v3)"],
  ["Coordinación ANIN", "Especialista en Infraestructura Verde (suscrito) y Especialista en Infraestructura Marrón, Ing. Pedro Talledo Hernández; supervisión del servicio con el Especialista en Intervenciones Vegetativas de Uso Sostenible para los Lotes SUS"],
]));
x.push(fuente("Fuente: TdR del Estudio de Suelos; Plan de Trabajo aprobado; Plan de Trabajo Detallado del Perfil, Actualización N° 01, Versión 2 (16/10/2026)."));

x.push(h1("3. ESTADO DE LOS ENTREGABLES Y DE LAS CONFORMIDADES"));
x.push(N.cuadro("Hitos contractuales del Estudio de Suelos y su estado al corte del segundo entregable"));
x.push(tabla([{ w: 1300, align: C }, { w: 2300 }, { w: 1400, align: C }, { w: 2838 }, { w: 1800, align: C }], ["ENTREGABLE", "CONTENIDO", "PLAZO CONTRACTUAL", "SEGUIMIENTO EN EL PERIODO", "CONDICIÓN"], [
  ["N° 1 (20 %)", "Plan de Trabajo y Plan de Muestreo georreferenciado; identificación del laboratorio", "Día 10 – 27/08/2026", "Sin cambios; base de la revisión del Entregable N° 2 (Factor K, Factor R PISCO 2017/2020, esquema de laboratorio, red de puntos)", "CONFORME – Informe N° 000251-2026-ANIN/DIME-SESDI-ARV (04/09/2026)"],
  ["N° 2 (30 %)", "Informe de Avance – Fase de Campo: fichas de calicatas y chequeos, infiltración, registro fotográfico, plan de muestreo actualizado, cadena de custodia, GDB", "Día 40 – 26/09/2026", "Recibido (versión en Drive del 27/09/2026); revisado al 100 % con reporte del 03/10/2026; 14 observaciones formuladas (numeral 5)", "REVISADO – conformidad sujeta al levantamiento de observaciones"],
  ["N° 3 (20 %)", "Resultados de laboratorio e interpretación preliminar; Factor K", "Día 55 – 11/10/2026", "Muestras acopiadas, codificadas, enviadas e ingresadas al laboratorio; resultados no recibidos al corte", "EN PROCESO"],
  ["N° 4 (30 %)", "Informe final: CUM por unidad, Factor K, Matriz de Compatibilidad Suelo–Intervención, estudio agronómico de Lotes SUS y geodatabase integrada", "Día 75 – 31/10/2026", "Pendiente; el Plan de Trabajo Detallado del Perfil reprograma su recepción y revisión al 13/11/2026", "PENDIENTE"],
], { size: 16 }));
x.push(fuente("Fuente: Informe de Avance – Fase de Campo (Cuadros 1 y 22); tablero de avance del Perfil (03/10/2026); Plan de Trabajo Detallado, Actualización N° 01, Versión 2 (líneas E5.1 a E5.4 y hito H7)."));
x.push(para("El avance ponderado del E5 en el Plan de Trabajo Detallado del Perfil es de 63 % (Entregable N° 1: 100 %; N° 2: 100 %; N° 3: 40 %; N° 4: 0 %). La diferencia entre el término contractual (31/10/2026) y la fecha reprogramada (13/11/2026) debe formalizarse con la adenda o ampliación de plazo que corresponda, pues condiciona la integración del laboratorio al E9 y a las metas del E10."));

x.push(h1("4. SEGUIMIENTO DE LAS ACTIVIDADES DE CAMPO"));
x.push(h2("4.1 Cronograma real frente al programado"));
x.push(para("El Plan de Muestreo previó la movilización en los días 11 y 12 y la apertura de calicatas del día 13 al día 38 con dos brigadas. La fase de campo se inició el día 17 (03/09/2026), una vez otorgada la conformidad del Entregable N° 1, y se ejecutó con tres brigadas para recuperar el plazo: las 74 calicatas se cerraron el día 33 (19/09/2026), cinco días antes de lo previsto, en 17 jornadas efectivas, con una producción media de 4.4 calicatas y 8.7 chequeos por jornada. Los chequeos finales y la prueba INF-11 se completaron el día 34 (20/09/2026)."));
x.push(N.cuadro("Itinerario de campo verificado (síntesis por semana)"));
x.push(tabla([{ w: 1900 }, { w: 1300, align: C }, { w: 2200 }, { w: 1200, align: C }, { w: 1200, align: C }, { w: 1838 }], ["SEMANA", "DÍAS DEL SERVICIO", "ZONAS HOMOGÉNEAS", "CALICATAS", "CHEQUEOS", "PRUEBAS DE INFILTRACIÓN"], [
  ["03 al 06/09/2026", "17 a 20", "ZH-05, 23, 24, 30, 31, 33", "13", "25", "INF-02b, INF-04, INF-09"],
  ["07 al 13/09/2026", "21 a 27", "ZH-03, 06, 07, 09, 10, 11, 13, 17, 22, 25, 27, 29, 32", "32", "60", "INF-02, INF-03, INF-06, INF-08, INF-12"],
  ["14 al 20/09/2026", "28 a 34", "ZH-02, 10, 12, 14, 15, 18, 19, 20, 21, 23, 26, 28, 31, 33", "29", "63", "INF-01, INF-05, INF-07, INF-10, INF-11"],
], { total: ["TOTAL", "17 jornadas", "33 zonas homogéneas", "74", "148", "13 sitios"] }));
x.push(fuente("Fuente: Informe de Avance – Fase de Campo, Cuadro 11 (itinerario por fecha), agregado por semana por el suscrito."));

x.push(h2("4.2 Cumplimiento de la intensidad de muestreo"));
x.push(N.cuadro("Verificación de la intensidad de muestreo frente al TdR"));
x.push(tabla([{ w: 2300 }, { w: 2000, align: C }, { w: 1800, align: C }, { w: 2338, align: C }, { w: 1200, align: C }], ["EXIGENCIA", "MÍNIMO DEL TdR", "PLAN", "EJECUTADO", "CUMPLE"], [
  ["Calicatas (numeral 6.2.3)", "1 por 200 ha (62)", "74 (1 por 166 ha)", "74: 35 de bloque, 34 de referencia, 5 en Lotes SUS", "Sí"],
  ["Barrenaciones de comprobación", "2 por calicata (148)", "148", "148; distancia mediana a su calicata: 227 m", "Sí"],
  ["Pruebas de infiltración", "1 por suelo dominante, 5 repeticiones", "12 sitios", "13 sitios; 10 con 5 repeticiones, INF-08 con 4, INF-02b con 3 e INF-11 con 2", "Parcial (repeticiones)"],
  ["Muestreo de horizontes", "2 por calicata (148)", "≥ 148", "209 horizontes registrados", "Sí"],
  ["Georreferenciación", "GPS, UTM WGS 84 17S, altitud", "Sí", "74 + 148 + 13 puntos en la GDB", "Sí"],
  ["Registro fotográfico", "Perfil con escala, entorno, chequeos y paisaje", "Sí", "Fotografías renombradas con el código oficial del punto", "Sí"],
]));
x.push(fuente("Fuente: Informe de Avance – Fase de Campo, Cuadros 8, 9 y 10, verificados por el suscrito."));
x.push(para("La observación del primer entregable sobre la intensidad del plan (242 puntos en el Plan de Trabajo, 230 observaciones en el avance y 244 en la suma por categorías) quedó levantada: el Entregable N° 2 consolida 74 calicatas, 148 chequeos —dos por calicata— y 13 sitios de infiltración, es decir, 235 puntos georreferenciados en la geodatabase ENTREGABLE2_PUNTOS_MUESTREADOS."));

x.push(h2("4.3 Reubicaciones e incidencias verificadas"));
x.push(para("De las 74 calicatas, 41 se abrieron a 100 m o menos del punto de gabinete, 16 entre 100 y 300 m y 17 a más de 300 m. Todas permanecen en su unidad de suelo y zona homogénea, y los dos chequeos de cada calicata confirmaron la extensión lateral de la unidad. Las reubicaciones mayores se listan en el Cuadro N° " + N.sigC() + "."));
x.push(N.cuadro("Calicatas desplazadas más de 300 m respecto del Plan de Muestreo (UTM WGS 84 Zona 17S)"));
const reub = [
  ["CAL-06", "Ref.", "641,777", "9,401,148", "2,248", "Hualtacal; el sector de gabinete estaba en valle aluvial"],
  ["CAL-13", "Bloque", "606,267", "9,455,115", "425", "Ajuste dentro del bloque 36 por acceso"],
  ["CAL-20", "Ref.", "618,323", "9,428,263", "305", "Casa Blanca, ladera inferior conservada"],
  ["CAL-33", "Bloque", "605,935", "9,445,793", "530", "Ajuste dentro de M17B4 (Platanal Alto)"],
  ["CAL-34", "Ref.", "607,329", "9,445,666", "1,258", "El Guabo; INF-09 ejecutada en el nuevo punto"],
  ["CAL-36", "Ref.", "600,152", "9,442,787", "602", "Panecillo, mismo sector conservado"],
  ["CAL-41", "Bloque", "618,448", "9,413,273", "373", "Ajuste dentro del bloque 5 (El Ala)"],
  ["CAL-46", "Ref.", "648,423", "9,427,827", "818", "Naranjito, ladera con bosque conservado"],
  ["CAL-50", "Ref.", "654,262", "9,422,149", "517", "Nuevo Progreso Limón por acceso"],
  ["CAL-56", "Ref.", "662,233", "9,394,048", "321", "Ajuste en Ramón Castilla"],
  ["CAL-60", "Ref.", "619,701", "9,448,265", "510", "Condorhuachina, bosque relicto conservado"],
  ["CAL-61", "Bloque", "622,268", "9,453,891", "968", "Alto Poclus (Peña Negra) por acceso"],
  ["CAL-62", "Ref.", "624,660", "9,451,896", "500", "Huaylingas, sector Culcas Alto"],
  ["CAL-63", "Bloque", "664,205", "9,400,931", "813", "Cruz de Piedra; INF-08 en el nuevo punto"],
  ["CAL-65", "Bloque", "659,664", "9,401,481", "5,183", "Par trasladado a Gaspar, dentro de M22B1"],
  ["CAL-66", "Ref.", "658,772", "9,400,795", "6,438", "Par trasladado a Gaspar, dentro de M22B1"],
  ["CAL-68", "Lote SUS", "625,098", "9,444,680", "441", "Ubicado con el agricultor sobre la parcela"],
];
x.push(tabla([{ w: 1000, align: C }, { w: 1000, align: C }, { w: 1300, align: R }, { w: 1400, align: R }, { w: 1300, align: R }, { w: 3638 }], ["CALICATA", "TIPO", "ESTE CAMPO (m)", "NORTE CAMPO (m)", "DESPLAZ. (m)", "MOTIVO REGISTRADO"], reub, { size: 15 }));
x.push(fuente("Fuente: Informe de Avance – Fase de Campo, Cuadro 7. Coordenadas dentro del rango de validación del proyecto (Este 450,000–750,000 m; Norte 9,300,000–9,600,000 m)."));
x.push(para("Se verificaron además tres incidencias: (i) la logística de agua y cilindros limitó las pruebas INF-08 (páramo, 3,199 msnm, cuatro repeticiones) e INF-11 (bosque relicto, 2,412 msnm, dos repeticiones), y la prueba prevista en Maray se trasladó a Chililique Alto; (ii) 27 calicatas se abrieron en laderas de más de 50 % de pendiente y 26 entre 25 y 50 %, con altitudes de 148 a 3,201 msnm; y (iii) en el sector Hualcas–Chignia (Huarmaca), con actividad minera y rondas campesinas, el ingreso de CAL-17, CAL-18 e INF-12 se coordinó con autoridades y rondas, como medida de seguridad del personal conforme a la Ley N° 29783, sin reubicar puntos."));

x.push(h1("5. REVISIÓN TÉCNICA DEL ENTREGABLE N° 2 Y MATRIZ DE OBSERVACIONES"));
x.push(para("La revisión contrastó el Informe de Avance con el TdR del estudio, el Plan de Muestreo aprobado, el catálogo V6 y la información del componente verde. El entregable cumple el alcance de la Fase 2; las observaciones formuladas no invalidan la fase de campo y se clasifican en: consistencia (C), alcance para los Entregables N° 3 y N° 4 (A) y forma (F). Su levantamiento se exige en los entregables siguientes, como condición de la conformidad del informe final."));
x.push(N.cuadro("Matriz de observaciones formuladas al Entregable N° 2 del Estudio de Suelos"));
const OBS = [
  ["OS-01", "C", "2.1 y Cuadro 1", "Se consignan 118 bloques y 12,299 ha; el catálogo vigente V6 tiene 117 bloques y 12,270.235 ha.", "Homologar el universo con el catálogo V6 y declarar las calicatas en bloques retirados, si las hubiera.", "E4"],
  ["OS-02", "C", "Resumen ejecutivo", "Se reportan 65 repeticiones de infiltración medidas; los Cuadros 8 y 9 y el numeral 15.1 suman 59.", "Corregir la cifra en todo el documento.", "E4"],
  ["OS-03", "C", "Resumen ejecutivo y 8.3", "Se reportan 152 horizontes muestreados; los Cuadros 9 y 10 y el numeral 8.3 reportan 209.", "Precisar el número de muestras remitidas al laboratorio por tipo de análisis (caracterización, fertilidad, densidad aparente).", "E3"],
  ["OS-04", "C", "Cuadros 8 y 19; 7.1", "INF-11 figura como «Culcas Alto» (Cuadro 8) y como «La Unión» (Cuadro 19 y 7.1).", "Uniformizar la denominación del sitio con el centro poblado del catálogo INEI.", "E4"],
  ["OS-05", "A", "5.6, 7.1 y 12.4", "INF-11 (2 repeticiones, 73.48 cm/h), INF-08 (4) e INF-02b (3) no alcanzan las 5 repeticiones del TdR; el libro de cálculo estima repeticiones faltantes.", "Declarar estos valores como indicativos, separar las repeticiones estimadas de las medidas y no usarlos para dimensionar obras sin repetición.", "E3/E4"],
  ["OS-06", "A", "12.4", "En INF-02b (El Ala) la primera repetición (17.2 cm/h) se atribuye a flujo preferencial y eleva la velocidad representativa a 6.88 cm/h; las otras dos dan 2.17 y 1.55 cm/h.", "Reportar la velocidad representativa con y sin la repetición atípica y justificar la adoptada.", "E4"],
  ["OS-07", "C", "15.2 y Cuadro 22", "El laboratorio figura como UNALM (Entregable N° 2), LASPAF-UNALM (Plan de Trabajo) y «Valle Grande» (Plan de Trabajo Detallado del Perfil).", "Acreditar el laboratorio efectivo, su acreditación vigente, la fecha de envío y la cadena de custodia de cada lote.", "E3"],
  ["OS-08", "F", "Cuadro 11", "La fila del 20/09 y la fila de totales del itinerario están desalineadas.", "Corregir el formato del cuadro.", "E4"],
  ["OS-09", "A", "5.4 y Cuadro 7", "17 calicatas se reubicaron a más de 300 m; el par CAL-65/CAL-66 a 5.2 y 6.4 km del punto de gabinete.", "Adjuntar la superposición de cada punto reubicado con el polígono V6 y la unidad de suelo (GDB), como evidencia de pertenencia.", "E4"],
  ["OS-10", "A", "10 y Cuadro 16", "La CUM de campo difiere de la CUM de gabinete en varias unidades (p. ej., CAL-17/18: P3 y X frente a C2se(r); CAL-31/32/37/38: F2/P3 y X frente a A3s(r)).", "Explicar las diferencias y resolverlas con los resultados de laboratorio y el ajuste por zona de vida del D.S. N° 005-2022-MIDAGRI.", "E4"],
  ["OS-11", "A", "14 y Cuadro 21", "La Matriz de Compatibilidad Suelo–Intervención es preliminar por grupo de suelos y no por unidad de suelo ni por bloque.", "Entregarla por unidad de suelo y por bloque (unidades de manejo), con la máscara de profundidad efectiva < 50 cm y conciliando las especies con la lista priorizada del E6.", "E4"],
  ["OS-12", "A", "13 y Cuadro 20", "Los códigos de permeabilidad del Factor K se asignaron por textura al tacto donde no hay prueba de infiltración; el Factor R con PISCO (2017 y 2020) no se reporta.", "Recalcular el Factor K con la textura y la materia orgánica de laboratorio y reportar el Factor R por unidad.", "E3/E4"],
  ["OS-13", "A", "Cobertura del estudio", "Del total, 103 de 117 bloques tienen alguna unidad de suelo muestreada; 14 se caracterizarán por analogía.", "Listar los bloques caracterizados por analogía y la unidad muestreada que los representa.", "E4"],
  ["OS-14", "A", "Cuadro 22", "El término contractual del Entregable N° 4 es el 31/10/2026; el Plan de Trabajo Detallado reprograma su recepción al 13/11/2026.", "Formalizar la reprogramación mediante el instrumento contractual que corresponda.", "Gestión"],
];
x.push(tabla([{ w: 800, align: C }, { w: 450, align: C }, { w: 1150 }, { w: 3300 }, { w: 3138 }, { w: 800, align: C }], ["CÓD.", "TIPO", "SECCIÓN", "OBSERVACIÓN", "LEVANTAMIENTO REQUERIDO", "SE LEVANTA EN"], OBS, { size: 15 }));
x.push(fuente("Fuente: revisión técnica del suscrito sobre el Informe de Avance – Fase de Campo (GeoSIG Ingenieros, setiembre de 2026). Tipo: C consistencia; A alcance para los Entregables N° 3 y N° 4; F forma."));

x.push(h1("6. VERIFICACIÓN DE LA APLICACIÓN AL COMPONENTE VERDE"));
x.push(h2("6.1 Infiltración y dimensionamiento de obras"));
x.push(para("La velocidad de infiltración básica es el parámetro hidráulico con el que se dimensionan las zanjas de infiltración y las terrazas de formación lenta (Guía 5). Los resultados separan tres situaciones con consecuencias de diseño distintas, que se verificaron con el Especialista en Infraestructura Marrón:"));
x.push(bullet([b("Suelos lentos a moderadamente lentos (0.33 a 1.37 cm/h, 9 sitios): "), "bosque seco y matorral andino con contacto paralítico o dénsico entre 30 y 100 cm. Las zanjas requieren mayor capacidad de almacenamiento y menor espaciamiento; donde el contacto está a menos de 50 cm deben priorizarse medidas de cobertura y retención superficial."]));
x.push(bullet([b("Suelos moderados (2.20 y 6.88 cm/h, 2 sitios): "), "La Laguna (Lalaquiz) y El Ala (Buenos Aires); el valor de El Ala está sujeto a la observación OS-06."]));
x.push(bullet([b("Suelos rápidos de altura (13.25 y 73.48 cm/h, 2 sitios): "), "páramo y bosque relicto; las zanjas no son prioritarias porque el suelo ya infiltra, y la medida es la conservación del horizonte orgánico (valores indicativos, OS-05)."]));
x.push(figura(FIG + "fig_infiltracion.png", 15.5));
x.push(N.figura("Velocidad de infiltración básica representativa por sitio (escala logarítmica)"));
x.push(fuente("Fuente: Informe de Avance – Fase de Campo, Cuadro 19."));
x.push(N.cuadro("Pruebas de infiltración: sitio, unidad de suelo y clase"));
x.push(tabla([{ w: 900, align: C }, { w: 2100 }, { w: 900, align: C }, { w: 2300 }, { w: 700, align: C }, { w: 1100, align: R }, { w: 1638 }], ["PRUEBA", "SITIO", "CALICATA", "UNIDAD DE SUELO", "REP.", "VEL. BÁSICA (cm/h)", "CLASE"], [
  ["INF-01", "Ricardo Palma (Yamango)", "CAL-27", "M-ge / mte-PT / Bes-cm", "5", "0.59", "Moderadamente lenta"],
  ["INF-02", "Hualas (Buenos Aires)", "CAL-07", "MS-Cu / mte-PT / Bes-cm", "5", "0.79", "Moderadamente lenta"],
  ["INF-02b", "El Ala (Buenos Aires)", "CAL-41", "MS-Cu / md-T / Bes-cm", "3", "6.88", "Moderadamente rápida"],
  ["INF-03", "Sahuate Hualanga (Huarmaca)", "CAL-43", "V-ar / bs-PT / Ma", "5", "0.41", "Lenta"],
  ["INF-04", "Chililique Alto (Chulucanas)", "CAL-33", "P-gm / mte-PT / Bes-cm", "5", "0.78", "Moderadamente lenta"],
  ["INF-05", "Santa Rosa de Chirimoyos", "CAL-01", "VS-at / bs-PT / Bes-cm", "5", "0.33", "Lenta"],
  ["INF-06", "Quemazón (San Juan de Bigote)", "CAL-15", "M-ge / mte-PT / Bes-cm", "5", "1.37", "Moderadamente lenta"],
  ["INF-07", "La Peña (Chulucanas)", "CAL-22", "Qal-ga / md-T / Bes-cm", "5", "0.93", "Moderadamente lenta"],
  ["INF-08", "Cruz de Piedra (S. M. de El Faique)", "CAL-63", "VS-at / ph-SaT / Pa", "4", "13.25", "Rápida"],
  ["INF-09", "El Guabo (Frías)", "CAL-34", "P-gm / mte-PT / Bes-cm", "5", "0.35", "Lenta"],
  ["INF-10", "La Laguna (Lalaquiz)", "CAL-45", "M-ge / bs-PT / Ma", "5", "2.20", "Moderada"],
  ["INF-11", "La Unión / Culcas Alto", "CAL-62", "P-tg / bh-MBT / Br-mvoc", "2", "73.48", "Muy rápida"],
  ["INF-12", "Pampa Quemada (Huarmaca)", "CAL-17", "S-cm / mte-PT / Bes-cm", "5", "0.83", "Moderadamente lenta"],
], { size: 15 }));
x.push(fuente("Fuente: Informe de Avance – Fase de Campo, Cuadros 8 y 19. Método de doble anillo (Muntz) con ajuste de Kostiakov; clases USDA."));

x.push(h2("6.2 Capacidad de uso y grupos de suelos"));
x.push(para("La pre-clasificación de CUM por calicata, con los parámetros de campo dominantes (pendiente, profundidad efectiva y pedregosidad), es coherente con la CUM de gabinete que asigna protección (Xse) a la mayor parte de la superficie y con el objetivo del proyecto de intervenir laderas degradadas con medidas de recuperación y conservación."));
x.push(N.cuadro("Pre-clasificación de CUM de campo y grupos de suelos (74 calicatas)"));
x.push(tabla([{ w: 3638 }, { w: 1200, align: C }, { w: 4800 }], ["CATEGORÍA", "CALICATAS", "IMPLICANCIA PARA EL COMPONENTE VERDE"], [
  ["Protección (X)", "34", "Revegetación con especies nativas, exclusión de pastoreo y control de cárcavas; sin obras de remoción"],
  ["Forestal de calidad baja con alternativa de protección (F3 / X)", "19", "Reforestación y enriquecimiento con especies nativas del piso"],
  ["Forestal de calidad media con alternativa de pastoreo (F2 / P3)", "8", "Reforestación, sistemas silvopastoriles y terrazas de formación lenta"],
  ["Pastoreo (P3)", "8", "Manejo del pastoreo y clausura temporal; revegetación de amarre"],
  ["Cultivo permanente (C3 / P2)", "4", "Sistemas agroforestales; candidatos a Lotes SUS"],
  ["Cultivo en limpio de calidad baja (A3 / C2)", "1", "Lotes SUS con prácticas de conservación de suelos"],
  ["Grupo 1 – ócricos sin horizonte de diagnóstico (Entisoles; Regosoles, Leptosoles)", "50", "Zanjas solo con ≥ 50 cm de suelo y pendiente de 15 a 50 %; proteger el horizonte A"],
  ["Grupo 2 – ócricos con horizonte cámbico (Inceptisoles; Cambisoles)", "18", "Zanjas y terrazas en pendientes < 50 %; agroforestería en Lotes SUS"],
  ["Grupo 3 – epipedón mólico (Mollisoles; Phaeozems)", "6", "Conservación de la cobertura; evitar excavaciones; zanjas no prioritarias"],
]));
x.push(fuente("Fuente: Informe de Avance – Fase de Campo, numerales 9.4 y 10 (Cuadro 16). Clasificación tentativa, sin laboratorio."));
x.push(h2("6.3 Integración por bloque en el diagnóstico territorial"));
x.push(para(`En coordinación con el Especialista en Infraestructura Marrón, los resultados de campo se llevaron al bloque en el Diagnóstico Territorial y Social (Versión 4 + suelos del 03/10/2026 y Versión 5 del 17/10/2026): Huancabamba con ${H.cal_n} calicatas en ${H.cal_bloques} bloques y Morropón con ${M.cal_n} calicatas en ${M.cal_bloques} bloques. Cada ficha de bloque incorpora el perfil de la calicata de bloque y su referencia, la profundidad efectiva, la CUM de campo y la regla preliminar suelo–intervención: zanjas de infiltración en pendientes de 15 a 50 % con 50 cm o más de suelo; terrazas de formación lenta de 15 a 40 %; y, en pendientes mayores de 50 %, revegetación, exclusión de pastoreo y control de cárcavas.`));
x.push(para("El contraste de los 34 pares bloque–referencia, que comparten unidad de suelo, valida el diseño del muestreo: en 17 pares la referencia es más profunda o conserva un horizonte de diagnóstico que el bloque perdió (p. ej., CAL-63/CAL-64: 18 cm frente a 150 cm), lo que cuantificará con el laboratorio la pérdida de materia orgánica, densidad aparente y saturación de bases atribuible a la degradación y servirá como línea base de recuperación para el indicador de brecha."));

x.push(h1("7. COORDINACIÓN Y GESTIÓN DEL SEGUIMIENTO"));
[
  "Revisión conjunta del Informe de Avance con el Especialista en Infraestructura Marrón, para los productos compartidos: infiltración (dimensionamiento de zanjas y terrazas) y erodabilidad (selección de obras físicas que protegen a las plantaciones antes del prendimiento).",
  "Reporte del estado del E5 en el tablero de avance del Perfil (03/10/2026) y en el Plan de Trabajo Detallado, Actualización N° 01, versiones 1 (11/10/2026) y 2 (16/10/2026), con la fecha de recepción del laboratorio y del informe final.",
  "Integración de los datos de campo en los tres volúmenes de diagnóstico y en la pestaña «Suelos», como insumo del E9 (Capítulo de Identificación) y del E10 (análisis técnico).",
  "Seguimiento del envío de muestras y de la cadena de custodia para el Entregable N° 3, con verificación de la etiqueta física de cada bolsa contra el registro.",
].forEach((t) => x.push(bullet(t)));

x.push(h1("8. CONCLUSIONES"));
[
  "El Estudio de Suelos mantiene la conformidad de su Entregable N° 1; el Entregable N° 2 fue recibido y revisado al 100 %, el Entregable N° 3 está en laboratorio y el Entregable N° 4 está pendiente, con reprogramación al 13/11/2026 por formalizar.",
  "La fase de campo cumplió la intensidad exigida —74 calicatas, 148 chequeos, 13 sitios de infiltración y 209 horizontes— y las reubicaciones se mantuvieron dentro de la unidad de suelo; solo el número de repeticiones de tres pruebas de infiltración no alcanza el mínimo del TdR.",
  "Las 14 observaciones formuladas son de consistencia, alcance y forma, no invalidan la fase de campo y se levantan en los Entregables N° 3 y N° 4.",
  "La información edáfica es aplicable al componente verde: condiciona el dimensionamiento de zanjas en suelos lentos, ubica la mayor parte de las tierras en protección y forestal y ya está integrada en la ficha de cada bloque.",
].forEach((t, k) => x.push(num(`8.${k + 1}.`, t)));
x.push(h1("9. RECOMENDACIONES"));
[
  "Condicionar la conformidad del Entregable N° 4 al levantamiento de las observaciones OS-01 a OS-14 y a la homologación del universo de bloques con el catálogo V6.",
  "Entregar la Matriz de Compatibilidad Suelo–Intervención por unidad de suelo y por bloque, con la máscara de profundidad efectiva y la conciliación de especies con el E6, para su uso directo en las metas del E10.",
  "Interpretar los resultados de laboratorio en función del componente verde —aptitud por especie, dosis de enmiendas y Factor K por unidad— y no solo como reporte analítico.",
  "Someter al criterio de idoneidad del Paso 4 (pendiente máxima de 75 %) los sectores con pendiente en el umbral, como CAL-65 en M22B1, y formalizar la reprogramación del Entregable N° 4.",
].forEach((t, k) => x.push(num(`9.${k + 1}.`, t)));

x.push(h1("10. MEDIOS VERIFICADORES"));
x.push(N.cuadro("Medios verificadores de la actividad c)"));
x.push(tabla([{ w: 3500 }, { w: 4138 }, { w: 2000, align: C }], ["MEDIO VERIFICADOR", "UBICACIÓN / IDENTIFICADOR", "FECHA"], [
  ["Informe de Avance – Fase de Campo (Entregable N° 2), PDF y Word", "Google Drive ID 1iWNiXKXVTf5XYVXS0RHFaJBZYSO5iqnE (PDF) y 1B5Z7G-scRsjtE7hpchibTXC9_qWg2x9s (Word), carpeta «Levantamiento de suelos_ANIN»", "27/09/2026"],
  ["Carpeta compartida por la consultora", "Google Drive, carpeta «Levantamiento de suelos_ANIN» (ID 1COyqfGePNgQUa9ioCOFDW_sM5nNu7mRf)", "10/09/2026 en adelante"],
  ["TdR del Estudio de Suelos y contenido mínimo", "Google Drive ID 1VFZfzJdX0_Ow_gfsMoLvKfwIvkoUe8Hr; carpeta «CONTENIDO MINIMO ESTUDIO DE SUELOS»", "27/09/2026"],
  ["Tablero de avance del Perfil (E5 revisado al 100 %)", "Google Drive ID 1wRdIvYH8ni8kbj3XsDKa0cgrgSIdVWn6 («IN Piura Avance al 3 de octubre 2026»)", "03/10/2026"],
  ["Plan de Trabajo Detallado, Actualización N° 01, Versión 2 (estado del E5)", "Google Drive ID 1QkhMG4cWkPFAyhxOH8qtaWG2YVCCJsGy", "16–17/10/2026"],
  ["Integración por bloque (pestaña «Suelos»)", "Volúmenes I y II, Versión 5 (enlaces en el Anexo 1)", "17/10/2026"],
  ["Matriz de observaciones en formato nativo", "Anexo 4 (Excel), hoja «Obs_Suelos_E2»", "Octubre 2026"],
], { size: 15 }));
x.push(fuente("Fuente: elaboración propia."));

L.guardar([seccionCaratula(DEF.anexo2), L.seccion(x, { pie: PIE })], OUT, "Anexo N° 2 – Actividad c)");
module.exports = { OBS };
