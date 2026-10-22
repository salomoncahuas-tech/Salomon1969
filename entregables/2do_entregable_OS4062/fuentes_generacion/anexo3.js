// Anexo N° 3 — Actividad d): Diagnóstico del Estado de los Ecosistemas (3-A) y Estudio de Oferta de Material Vegetal (3-B)
// Alineado al Plan de Trabajo Detallado del Perfil – Actualización N° 01, Versión 2 (16/10/2026): metodología de firma espectral
const L = require("./lib");
const { para, b, i, run, h1, h2, h3, guion, bullet, num, fuente, tabla, vacio, centro, W, C, R, fmt, figura } = L;
const { seccionCaratula } = require("./caratula");
const DEF = require("./anexos_def");
const OUT = process.argv[2];
const FIG = "/tmp/claude-0/-home-user-Salomon1969/03b82ed8-c0e9-54ec-93c9-bff5f93b575b/scratchpad/build/fig/";
const PIE = "Anexo N° 3 · Informe N° 008-2026-HSCM · Segundo Entregable · OS N° 4062-2026";
const N = L.numerador("A3.");
const PTD = "Plan de Trabajo Detallado del Perfil – Actualización N° 01, Versión 2 (16/10/2026)";

// ---------------- Parte común ----------------
const x = [];
x.push(centro([b("ANEXO N° 3", { size: 26 })], { spacing: { after: 60 } }));
x.push(centro([b("SEGUIMIENTO DEL DIAGNÓSTICO DEL ESTADO DE LOS ECOSISTEMAS Y DEL ESTUDIO DE OFERTA DE MATERIAL VEGETAL DE PROPAGACIÓN", { size: 22 })], { spacing: { after: 60 } }));
x.push(centro([i("Entregable 6 del Perfil (E6) por metodología de firma espectral — periodo del 26/09 al 27/10/2026", { size: 19 })], { spacing: { after: 240 } }));
x.push(h1("1. OBJETIVO Y ALCANCE"));
x.push(para("El presente anexo desarrolla la actividad d) del numeral 5.1 de los Términos de Referencia durante el periodo del segundo entregable. Documenta el seguimiento, en coordinación con el Especialista en Intervenciones Vegetativas de Uso Sostenible, Blgo. Abel Aucasime Orihuela, del Diagnóstico del Estado de los Ecosistemas (parte 3-A) y del Estudio de Oferta de Material Vegetal de Propagación (parte 3-B): la revisión y actualización de sus planes de trabajo, el seguimiento de las actividades, la revisión técnica de los entregables, la formulación y el levantamiento de observaciones y el estado de las conformidades."));
x.push(para("Ambos estudios forman un solo producto del Perfil, el Entregable 6 «Diagnóstico del estado del ecosistema y estudio de oferta de material vegetal», regido por el Índice de Contenidos Mínimos 999990-ANIN001-000-XX-SC-EN-000002-R00 (SESDI, mayo de 2025), con tres componentes: estado del ecosistema, flora nativa para la gestión del riesgo de desastres (GRD) y oferta de material vegetal."));
x.push(para([b("Documento rector. "), `El seguimiento se alinea al ${PTD}, en adelante «Plan de Trabajo institucional v2», elaborado por el suscrito para revisión de la Coordinación de Proyectos IN y conformidad de SESDI. Este plan establece que el E6 se desarrolla con la `, b("metodología de firma espectral"), " por el equipo base, sin servicio externo ni contratación de personal, con redacción final al 18/11/2026, y reemplaza en todo su texto la modalidad anterior del E6 (Anexo 4 del plan, comentario N° 1)."]));

x.push(h1("2. CONTEXTO: DEL PROCESO DESIERTO A LA METODOLOGÍA DE FIRMA ESPECTRAL"));
x.push(para("Tres hechos del periodo definieron el seguimiento de la actividad d):"));
x.push(bullet([b("Proceso desierto y opción descartada. "), "El procedimiento CP-ABR-5-2026-ANIN/DAI-1 para el Diagnóstico del Ecosistema y Estudio de Oferta de Material Vegetal fue declarado desierto. La opción de reconvocarlo quedó descartada al haberse desarrollado con éxito, por Administración Directa, la metodología de firma espectral; el E6 se elabora con el equipo base y su redacción final culmina el 18/11/2026 (Plan de Trabajo institucional v2, numeral 1.1 y Tabla 6)."]));
x.push(bullet([b("Alerta de El Niño Costero. "), "Los Comunicados Oficiales ENFEN N° 16-2026 (14/09/2026: 48 % de probabilidad de magnitud extraordinaria y 45 % de fuerte para diciembre 2026–marzo 2027) y N° 17-2026 (28/09/2026: lluvias sobre lo normal en la costa norte desde octubre) obligan a cerrar todo trabajo de campo antes de la temporada de lluvias; el plan institucional registra el E6 «sin parcelas de campo» como medida del riesgo R01."]));
x.push(bullet([b("Plaza de viveros desierta. "), "La plaza de Especialista en Infraestructura Rural para la Producción (viveros) quedó desierta (octubre de 2026). El plan institucional dispone reasignar formalmente sus funciones —oferta y brecha de viveros (E6), E10.2 y E12.2— entre el Especialista en Intervenciones Vegetativas de Uso Sostenible, el Especialista en Infraestructura Verde y el Especialista en Infraestructura Marrón antes del 23/10/2026 (numeral 1.4 c) y riesgo R05)."]));
x.push(para(`En consecuencia, el seguimiento se ejerció sobre (i) la secuencia de planes del E6 hasta su adopción en el Plan de Trabajo institucional v2, (ii) los insumos producidos por órdenes de servicio (OS 1147, 1148, 2357, 2726 y 2797), (iii) la Nota Técnica N° 002-2026-ANIN/DIME/SESDI de resultados por firma espectral y (iv) el Informe E6 en versión para revisión de SESDI. El equipo del E6, según el plan institucional (Tabla 16 y hoja «E6_Firma_Espectral» de sus matrices), es el siguiente (Cuadro N° ${N.sigC()}).`));
x.push(N.cuadro("Equipo del E6 según el Plan de Trabajo institucional v2 (equipo base, sin contratación de personal)"));
x.push(tabla([{ w: 3200 }, { w: 3300 }, { w: 3138 }], ["ROL EN EL E6", "PROFESIONAL", "FUNCIÓN"], [
  ["Responsable técnico del E6", "Ing. Héctor Salomón Cahuas Miller – Esp. en Infraestructura Verde", "Metodología de firma espectral, redacción y matrices"],
  ["Teledetección, procesamiento espectral, SIG y RPAS", "Ing. Albert Flores Mendoza – Especialista SIG y operador RPAS", "Procesamiento S2DR4, cartografía del E6 y ortofotos de control"],
  ["Coautor del diagnóstico", "Ing. Pedro Talledo Hernández – Esp. en Infraestructura Marrón", "Análisis del estado del ecosistema y vínculo con las medidas de IN RRD"],
  ["Consultado (RACI) y coordinación según el TdR", "Blgo. Abel Aucasime Orihuela – Esp. en Intervenciones Vegetativas de Uso Sostenible", "Flora de los Lotes SUS, oferta de plantones y reasignación de funciones de viveros"],
  ["Oferta de material vegetal y brecha de viveros", "Equipo base del E6", "Plaza de viveros desierta, sin asignación presupuestal"],
  ["Especialista en flora (P3), fotointérpretes (P5) y asistente de BD (P6)", "Sin contratación", "No requeridos con la metodología de firma espectral"],
]));
x.push(fuente("Fuente: Plan de Trabajo institucional v2, Tablas 16 y 17 (RACI: E6 con Coordinación IN «A»; Esp. Infraestructura Verde, Esp. Infraestructura Marrón y SIG «R»; Esp. Uso Sostenible «C»); matrices del plan, hoja «E6_Firma_Espectral». El plan consigna al Especialista SIG como «Albert Mendoza Flores»; se mantiene el orden de apellidos del Informe E6."));

x.push(h1("3. REVISIÓN Y ACTUALIZACIÓN DE LOS PLANES DE TRABAJO"));
x.push(h2("3.1 Secuencia de planes del E6 hasta la metodología de firma espectral"));
x.push(para(`El plan del E6 pasó por tres estados durante el periodo. La versión 1.1 suponía un servicio externo con parcelas de campo; la versión 2.0 «In house» (08/10/2026) lo adecuó a inspección remota con parcelas virtuales en 10 semanas; y el Plan de Trabajo institucional v2 (16/10/2026), sobre los resultados de la Nota Técnica N° 002-2026, adoptó la metodología de firma espectral como forma de ejecución del E6. El Cuadro N° ${N.sigC()} compara los tres estados.`));
x.push(N.cuadro("Evolución del plan de trabajo del E6: versión 1.1, versión 2.0 y Plan de Trabajo institucional v2"));
x.push(tabla([{ w: 1750 }, { w: 2350 }, { w: 2650 }, { w: 2888 }], ["COMPONENTE", "VERSIÓN 1.1", "VERSIÓN 2.0 «IN HOUSE» (08/10)", "PLAN INSTITUCIONAL v2 – FIRMA ESPECTRAL (16/10)"], [
  ["Modalidad", "Servicio externo con especialistas locales por provincia", "Equipo ANIN, inspección remota", "Administración Directa por el equipo base; sin servicio externo ni contratación de personal"],
  ["Plazo", "16 semanas (Plan aprobado: servicio de 120 días)", "10 semanas (S1 = 12/10/2026; informe final 18/12/2026)", "12/10/2026 – 18/11/2026 (redacción final)"],
  ["Unidad de análisis", "157 parcelas de campo de 0.1 ha", "≈ 1,400–1,700 parcelas virtuales de 0.5 ha", "Píxel de 1 m (S2DR4) en 92 tomas; K-means en 9 clases de cobertura"],
  ["Imagen base", "Espectrorradiómetro 350–2,500 nm", "SPOT 0.6 m, PerúSAT-1 0.7 m y S2DR4 1 m", "Sentinel-2 Deep Resolution (S2DR4, 1 m): 92 tomas TCI, IRP y NDVI"],
  ["Validación", "Parcelas de campo", "≈ 600 puntos ciegos; exactitud global ≥ 85 %", "Control con evidencia de campo 2026 (F-DT 117/117, ortofotos RPAS, fotos, calicatas); r = 0.72 con 82 fichas. Matriz de confusión formal, TAU y estructura vertical al estudio definitivo"],
  ["Flora", "Medida en parcelas", "Registros secundarios y consulta Delphi", "Especies asignadas por cobertura y piso altitudinal, contrastadas con registros de campo 2026"],
  ["Equipo", "Especialistas por provincia", "P2–P9 (flora, viveros, fotointérpretes, BD, RPAS)", "Esp. Infraestructura Verde, Especialista SIG y Esp. Infraestructura Marrón"],
  ["Costo", "S/ 526,900 referenciales", "Personal P3, P5 y P6 por contratar", "Sin contratación; las partidas 1.19, 1.20 y 1.53–1.59 quedan como referencia de la línea base"],
  ["Conformidad", "De entregables del servicio", "Memorando de SESDI en la semana 1", "Conformidad de IN/SESDI sobre la metodología (hito H7b), sin número de documento al corte"],
], { size: 15 }));
x.push(fuente("Fuente: E6 v2.0, sección «Cambios respecto a la v1.1» (08/10/2026); Plan de Trabajo institucional v2, numerales 7.2 y 12.2, Tablas 6, 8 y 16; matrices del plan, hojas «E6_Firma_Espectral» y «Control_Cambios»."));
x.push(para("Con ello, la programación de 10 semanas del E6 v2.0 (hitos H1 a H5, cierre al 18/12/2026) y su dotación de personal quedan sin efecto; se mantienen de esa versión la adecuación de los ítems del índice R00 y el criterio del índice de condición. La revisión del suscrito al E6 v2.0 («rev_hscm», 08/10/2026) —denominación de la modalidad, horizonte de registro del Perfil a marzo de 2027, asignación del Especialista SIG y operador RPAS, perfil del especialista en viveros y solicitud de inventarios florísticos de las ACR— se trasladó al plan institucional en lo que sigue vigente."));

x.push(h2("3.2 Adecuación del índice de contenidos mínimos a la firma espectral"));
x.push(para(`El plan institucional precisa cinco ítems del índice R00 que requieren conformidad de SESDI por su adecuación a la metodología de firma espectral. El Cuadro N° ${N.sigC()} los contrasta con la adecuación que aplica el Informe E6 en su versión para revisión.`));
x.push(N.cuadro("Ítems adaptados del índice R00: Plan de Trabajo institucional v2 frente al Informe E6"));
x.push(tabla([{ w: 1000, align: C }, { w: 1900 }, { w: 3369 }, { w: 3369 }], ["ÍTEM", "EXIGENCIA", "PLAN INSTITUCIONAL v2", "INFORME E6 (VERSIÓN PARA REVISIÓN)"], [
  ["7.2", "Razón social de la consultora", "ANIN – DIME – SESDI, por Administración Directa (equipo base, sin contratación de personal)", "ANIN – DIME – SESDI, Administración Directa"],
  ["9.1", "GPS y formatos de campo", "Sin levantamiento nuevo; fichas F-DT y tracks GPS de la verificación 2026", "Ráster S2DR4, DEM 12.5 m, catálogo V6 y evidencia de campo 2026; sin salidas nuevas"],
  ["9.5 y 10.7", "Unidades de muestreo", "Píxel de 1 m (S2DR4) en 92 tomas; K-means en 9 clases, sin parcelas de campo", "1,374 parcelas virtuales de 0.5 ha en bloques y 102 de referencia, con cálculo de tamaño de muestra"],
  ["9.6 y 10.8", "Mapas con validación en campo", "Control con evidencia de campo 2026: F-DT (117/117), ortofotos RPAS, fotos georreferenciadas y calicatas", "Contraste con 82 fichas F-DT (r = 0.72); matriz de confusión como vacío"],
  ["10.3", "Levantamiento de campo", "Se reporta la evidencia ya levantada; no hay levantamiento nuevo", "Se sistematiza la evidencia de mayo–septiembre de 2026"],
], { size: 15 }));
x.push(fuente("Fuente: matrices del Plan de Trabajo institucional v2, hoja «E6_Firma_Espectral»; Informe E6, sección «Adecuación de los contenidos mínimos a la inspección remota»."));
x.push(para("Los ítems coinciden en el fondo. La diferencia está en el ítem 9.5 y 10.7: para el plan la unidad de análisis es el píxel clasificado, mientras que el Informe E6 presenta las parcelas virtuales como unidad de muestreo. Por ser el plan el documento rector, el informe final debe presentar las parcelas virtuales como resumen estadístico de la cobertura clasificada a 1 m y no como sustituto de la parcela de campo (observación OE-16)."));

x.push(h2("3.3 Programación, avance e hitos del E6 en el plan institucional"));
x.push(para(`El plan institucional reprograma el E6 al periodo 12/10/2026 – 18/11/2026 (línea base 0: 01/05–31/08/2026; desviación de +79 días) y lo saca de la ruta crítica: sus insumos para el E10 —especies priorizadas y brecha de viveros— se entregan con el informe final, con 35 días de holgura frente al cierre del E9 (23/12/2026). El Cuadro N° ${N.sigC()} presenta la estructura de desglose del E6 con sus pesos y avances.`));
x.push(N.cuadro("Estructura de desglose del E6 en el Plan de Trabajo institucional v2"));
const EDT = [
  ["E6.1", "Caracterización de ecosistemas por bloque (F-DT, MINAM V5, MSAVI/NDVI); insumo de control de la clasificación S2DR4", 15, 85, 100],
  ["E6.2", "Estado de conservación y degradación: atributos, índice de condición, TAU y % bajo el umbral MSAVI 0.4976", 10, 60, 100],
  ["E6.3", "Composición florística por ecosistema y piso (registros secundarios, especies 2026 y asignación por piso)", 20, 28, 90],
  ["E6.4", "Imágenes S2DR4 (1 m) y ortofotos de dron, base de la firma espectral", 10, 30, 100],
  ["E6.5", "Lista corta de especies nativas para GRD y fichas técnicas", 5, 25, 75],
  ["E6.6", "Demanda de plantones por especie y bloque (117 bloques)", 10, 10, 75],
  ["E6.7", "Oferta: directorio de viveros, fuentes semilleras y brecha de viveros", 14, 10, 50],
  ["E6.8", "Propuesta de viveros permanentes y temporales, propagación y costos", 6, 20, 25],
  ["E6.9", "Metodología de firma espectral: clasificación de 92 tomas S2DR4 y asignación de especies por piso", 8, 90, 100],
  ["E6.10", "Informe final del E6 (resultados 10.2 a 10.11, conclusiones y anexos)", 2, 0, 75],
];
const pond = (k) => EDT.reduce((s, r) => s + r[2] * r[k], 0) / 100;
x.push(tabla([{ w: 900, align: C }, { w: 5338 }, { w: 900, align: C }, { w: 1250, align: C }, { w: 1250, align: C }], ["CÓD.", "COMPONENTE", "PESO", "AVANCE PLAN (WORD)", "AVANCE MATRICES (EXCEL)"],
  EDT.map((r) => [r[0], r[1], `${r[2]}`, `${r[3]} %`, `${r[4]} %`]),
  { total: ["E6", "Avance ponderado del entregable (peso del E6 en el Perfil: 8)", "100", `${fmt(pond(3), 1)} %`, `${fmt(pond(4), 1)} %`], size: 15 }));
x.push(fuente("Fuente: Plan de Trabajo institucional v2, Tabla 12 (corte 08/10/2026); matrices del plan, hoja de avance por EDT (actualizada al 16/10/2026). La diferencia entre ambos avances se registra como observación OE-14."));
x.push(N.cuadro("Hitos, metas y riesgos del E6 en el Plan de Trabajo institucional v2"));
x.push(tabla([{ w: 1200, align: C }, { w: 4300 }, { w: 1700, align: C }, { w: 2438 }], ["CÓD.", "DESCRIPCIÓN", "FECHA", "ESTADO SEGÚN EL PLAN"], [
  ["H6g", "Diagnóstico del ecosistema por firma espectral: clasificación de 92 tomas S2DR4 concluida", "13/10/2026", "Cumplido (equipo base, sin parcelas ni contratación)"],
  ["H7b", "Conformidad de SESDI / Coordinación IN sobre la metodología de firma espectral", "Por registrar", "Pendiente: remisión de la nota técnica e ítems adaptados del R00"],
  ["H7c", "Entregas parciales al E10: especies priorizadas por piso y brecha de viveros por escenarios", "18/11/2026", "Programado"],
  ["H7", "Entrega de estudios especializados (E6: redacción final)", "18/11/2026", "Reprogramado – en riesgo"],
  ["Mes 8", "Meta de octubre: clasificación espectral concluida y conformidad de IN/SESDI sobre la metodología", "oct-2026", "Clasificación cumplida; conformidad pendiente"],
  ["Mes 9", "Meta de noviembre: redacción final del E6 con especies priorizadas y brecha de viveros para el E10", "nov-2026", "En curso"],
  ["Mes 10", "Meta de diciembre: actualización del E8 con los resultados del E6", "dic-2026", "Programado"],
  ["R04", "La metodología de firma espectral no obtiene la conformidad de SESDI o es observada por la UF (12 · Alto)", "—", "Mitigación: nota técnica, trazabilidad de cada dato y control con evidencia de campo 2026"],
  ["R05", "Funciones de oferta y brecha de viveros sin responsable formal (16 · Muy alto)", "23/10/2026", "Mitigación: reasignación formal entre ESUS, EIV y EIM"],
  ["R10", "Criterio de brecha no unificado entre volúmenes (12 · Alto)", "—", "Adoptar un criterio único para los 117 bloques"],
  ["R17", "Imágenes SPOT y PerúSAT-1 y equipo RPAS disponibles tarde (12 · Alto)", "30/10/2026", "Usar las tomas S2DR4 del E6 y las ortofotos como control"],
], { size: 15 }));
x.push(fuente("Fuente: Plan de Trabajo institucional v2, Tablas 5, 13, 15 y 20; numeral 9.3 (ruta crítica)."));

x.push(h2("3.4 Comentarios del suscrito absueltos en el plan institucional"));
x.push(para(`Como parte de la revisión y actualización del plan, el suscrito formuló comentarios a la Actualización N° 01; el plan institucional v2 deja constancia de ellos en su Anexo 4. El Cuadro N° ${N.sigC()} recoge los que atañen al E6.`));
x.push(N.cuadro("Comentarios del Especialista en Infraestructura Verde sobre el E6 y su tratamiento en el plan institucional"));
x.push(tabla([{ w: 600, align: C }, { w: 1900 }, { w: 3600 }, { w: 3538 }], ["N°", "UBICACIÓN", "COMENTARIO (TEXTUAL)", "TRATAMIENTO EN LA VERSIÓN 2"], [
  ["1", "1.1 (viñeta del E6)", "Opción descartada al haberse desarrollado con éxito metodología de firma espectral", "Viñeta reescrita; se retiran del documento todas las referencias a la modalidad anterior del E6"],
  ["2", "1.4 c)", "Se sugiere reasignarlas entre los Especialistas en Lotes SUS, el Espec. En IV y el Espec. En IM", "Incorporado en 1.4 c) y en las Tablas 6, 8, 16 y 20 (R05)"],
  ["9", "Tabla 5 – H7b", "Leer columna de Estado ubicada a la izquierda dice PROGRAMADO, no se puede consignar por tanto numero de memorando", "H7b reformulado como conformidad pendiente de la metodología de firma espectral, sin número de documento"],
  ["11", "Tabla 8 – criterio de brecha", "Se aclara la necesidad de asignar un umbral MSAVI por ecosistema que se definirá en consenso con el equipo de Especialistas", "Incorporado en la Tabla 8"],
  ["14", "Tabla 12 – E6.6", "Esta cifra será proporcionada al momento de la elaboración del Entregable o Estudio 10 del perfil como parte del diseño preliminar de intervenciones", "Se retira la cifra de plantones; se proporcionará en el E10"],
], { size: 15 }));
x.push(fuente("Fuente: Plan de Trabajo institucional v2, Anexo 4, Tabla 24 (comentarios de la revisión del 16/10/2026)."));

// ---------------- 3-A ----------------
const a = [];
a.push(h1("3-A. DIAGNÓSTICO DEL ESTADO DE LOS ECOSISTEMAS"));
a.push(h2("3-A.1 Seguimiento de las actividades programadas"));
a.push(para("Conforme al plan institucional no hay levantamiento nuevo de campo: la evidencia de 2026 actúa como control de la clasificación espectral —fichas F-DT de los 117 bloques (evidencia A en 101, B en 12 y C en 4), 222 registros de calicatas y chequeos del E5 (34 sitios de referencia), fichas sociales con 143 registros de especies y ortofotos de vuelos RPAS (bloques 21, 60, 66 y 73)—. El avance de las actividades del cronograma del E6 se presenta en el Cuadro N° " + N.sigC() + "."));
a.push(N.cuadro("Avance de las actividades del E6 según el cronograma del Plan de Trabajo institucional v2"));
a.push(tabla([{ w: 600, align: C }, { w: 4400 }, { w: 4638 }], ["N°", "ACTIVIDAD DEL CRONOGRAMA (12/10 – 18/11/2026)", "ESTADO AL CIERRE DEL PERIODO"], [
  ["1", "Imágenes Sentinel-2 Deep Resolution (S2DR4, 1 m): 92 tomas (TCI, IRP, NDVI)", "Concluida: 92 tomas recibidas y procesadas (compuesto del 02/09 al 09/10/2026)"],
  ["2", "Extracción de firmas espectrales y clasificación K-means (8 variables) en 9 clases de cobertura", "Concluida el 13/10/2026 (hito H6g); 30 ráster regenerados y controlados, con acuerdo de 100 % en la toma de control"],
  ["3", "Asignación de especies por piso altitudinal", "Concluida (Nota Técnica N° 002-2026)"],
  ["4", "Control con la evidencia de campo 2026 y umbral MSAVI 0.4976", "Ejecutado: r = 0.72 frente a la cobertura de 82 fichas F-DT; error absoluto medio 19.5 pp y sesgo de +15.9 pp"],
  ["5", "Nota técnica de resultados y conformidad de IN/SESDI sobre la metodología", "Nota Técnica N° 002-2026 remitida (14/10/2026); conformidad pendiente (hito H7b)"],
  ["6", "Recorte de superficies a los 117 bloques (12,270.235 ha) y reporte por bloque, microcuenca y piso", "Ejecutado en el Informe E6, versión para revisión (15/10/2026); por consolidar en el informe final"],
  ["7", "Especies priorizadas por piso altitudinal y brecha de viveros por escenarios para el E10", "En curso: matriz de 69 especies y 16 fichas técnicas; brecha a la espera de los parámetros del E10 (hito H7c, 18/11/2026)"],
  ["8", "Redacción final del informe (índice R00: resultados 10.2–10.11, conclusiones y anexos)", "En curso; culmina el 18/11/2026"],
]));
a.push(fuente("Fuente: matrices del Plan de Trabajo institucional v2, hoja «E6_Firma_Espectral»; Nota Técnica N° 002-2026; Informe E6, numerales 9.6 y 10.1."));
a.push(para("La validación ciega con ≈ 600 puntos, la tendencia de degradación (TAU) y la estructura vertical medida, previstas en el E6 v2.0, no forman parte del alcance del plan institucional: este las traslada expresamente al estudio definitivo (numeral 7.2). El seguimiento las registra como vacíos declarados y no como actividades atrasadas."));

a.push(h2("3-A.2 Revisión técnica de los insumos producidos por órdenes de servicio"));
a.push(para("El 03/10/2026 se revisaron las carpetas de entregables de las órdenes de servicio que produjeron insumos para el E6 (Plan de Trabajo institucional v2, Tabla 6: «Culminadas; entregables revisados el 03/10/2026; listas florísticas por depurar»), con el resultado del Cuadro N° " + N.sigC() + "."));
a.push(N.cuadro("Revisión técnica de los entregables de las órdenes de servicio insumo del E6"));
a.push(tabla([{ w: 1600 }, { w: 2500 }, { w: 3138 }, { w: 2400 }], ["ORDEN DE SERVICIO", "PRODUCTO", "HALLAZGO DE LA REVISIÓN", "TRATAMIENTO EN EL E6"], [
  ["OS 1148 y OS 2797", "Flora y ecosistemas", "≈ 90 taxones (OS 1148) en listas preliminares con errores taxonómicos y sin depurar", "Depuración contra F-DT, F-DS y chequeos de suelos; matriz de 69 especies nativas"],
  ["OS 1147", "Oferta de material vegetal", "≈ 77 nativas; priorización preliminar de 35 arbóreas y 34 arbustivas; demanda de 2.94 M plantones solo para Lotes SUS en 83 bloques", "Lista recalculada por piso y ecosistema; la demanda para los bloques se proporcionará en el E10"],
  ["OS 2726", "Fitosanidad y presiones", "Presiones en 58 bloques: sobrepastoreo 96.6 %, quemas 86.2 %, especies invasoras 67.2 %; propagación de 10 especies; protocolo de vivero limpio; costo paramétrico S/ 12 por plantón", "Atributo «ausencia de amenazas» del IC con F-DT-04 de los 117 bloques; costo a sustentar en E10/E12"],
  ["OS 2357", "Sistematización de bloques", "Sistematización de la evidencia de bloques", "Insumo de la clasificación de evidencia A/B/C"],
  ["Conjunto", "Oferta", "Cinco fuentes de gabinete, ninguna en los 15 distritos; 6,000 plantones confirmados (Suyo)", "Directorio ampliado a 12 iniciativas (parte 3-B)"],
], { size: 16 }));
a.push(fuente("Fuente: tablero de avance del Perfil al 03/10/2026 (componentes del E6 y vacíos de integración)."));

a.push(h2("3-A.3 Revisión de la Nota Técnica N° 002-2026-ANIN/DIME/SESDI"));
a.push(para("La Nota Técnica (14/10/2026), dirigida a la Coordinación de Proyectos IN, es el documento de resultados de la metodología de firma espectral cuya remisión condiciona el hito H7b. Informa la clasificación de la cobertura vegetal de 92 tomas S2DR4 de 1 m (≈ 1,681 ha cada una; compuesto del 02/09 al 09/10/2026) con un único modelo K-means de 16 clases sobre 8 variables, agrupadas en 9 coberturas, y la asignación de especies probables por piso altitudinal. Al final de la época seca, el 50.9 % de la superficie analizada conserva vegetación verde, el 27.5 % está seca o defoliada y el 13.9 % sin cobertura vegetal efectiva; el piso A (< 1,000 msnm) concentra el suelo expuesto y el piso B (1,000–2,000 msnm) es el mejor conservado."));
a.push(para("La revisión precisó el texto respecto de la versión del 13/10/2026: numeración de la nota, ampliación de los ecosistemas a bosque montano relicto y páramos, y vínculo de las coberturas con los metrados. Se mantienen como limitaciones declaradas que las coberturas 4 y 5 incluyen bosque caducifolio sano sin hojas y que las superficies por toma incluyen traslapes; dos expresiones de la versión revisada requieren corrección (observaciones OE-04 y OE-05)."));

a.push(h2("3-A.4 Revisión técnica del Informe E6 (versión para revisión de SESDI)"));
a.push(para("El Informe E6 cubre los 14 capítulos del índice con tres líneas de evidencia —firma espectral de 92 tomas S2DR4, evidencia de campo 2026 e información secundaria—. Sus resultados coinciden con los que el plan institucional consigna en su numeral 7.2 y se presentan por provincia en el Cuadro N° " + N.sigC() + "."));
a.push(N.cuadro("Síntesis del estado de los ecosistemas por provincia (E6)"));
a.push(tabla([{ w: 1500 }, { w: 800, align: C }, { w: 1200, align: R }, { w: 1000, align: R }, { w: 1300, align: R }, { w: 1000, align: R }, { w: 1638, align: R }, { w: 1200, align: R }], ["PROVINCIA", "BLOQ.", "ÁREA (ha)", "MSAVI POND.", "BAJO 0.4976 (ha)", "% BAJO UMBRAL", "RESTAURACIÓN ACTIVA ≤ 0.3813 (ha)", "IC POND."], [
  ["Ayabaca", "10", "1,391.70", "0.548", "525.21", "37.7 %", "42.28", "0.851"],
  ["Huancabamba", "48", "3,551.16", "0.509", "1,599.99", "45.1 %", "925.51", "0.733"],
  ["Morropón", "59", "7,327.37", "0.379", "6,216.63", "84.8 %", "4,535.62", "0.706"],
], { total: ["TOTAL", "117", "12,270.24", "0.436", "8,341.82", "68.0 %", "5,503.41", "0.730"] }));
a.push(fuente("Fuente: Informe E6, Tabla 1; Plan de Trabajo institucional v2, numeral 7.2."));
a.push(N.cuadro("Clases de estado del ecosistema e intersección con el umbral MSAVI (117 bloques)"));
a.push(tabla([{ w: 2638 }, { w: 1500, align: C }, { w: 1900, align: R }, { w: 1500, align: R }, { w: 2100, align: R }], ["CLASE DE ESTADO (IC)", "BLOQUES", "ÁREA (ha)", "% ÁREA", "BAJO UMBRAL (ha)"], [
  ["No degradado", "53", "3,949.76", "32.2 %", "1,070.34"], ["Leve", "44", "6,675.47", "54.4 %", "5,669.78"],
  ["Moderado", "14", "1,493.62", "12.2 %", "1,450.79"], ["Severo", "6", "151.39", "1.2 %", "150.91"],
], { total: ["TOTAL", "117", "12,270.24", "100.0 %", "8,341.82"] }));
a.push(fuente("Fuente: Informe E6, Tabla 12. Los seis bloques severos son 77 y 83 a 87 (San Juan de Bigote)."));
a.push(N.cuadro("Índice de condición por atributo y valores de referencia por piso"));
a.push(tabla([{ w: 2800 }, { w: 1500, align: R }, { w: 5338 }], ["ATRIBUTO", "IC PONDERADO", "LECTURA"], [
  ["Ausencia de amenazas", "0.51", "Presión más baja del conjunto: 103 registros de sobrepastoreo y 83 de cambio de uso a pastoreo en F-DT-04"],
  ["Condiciones físicas", "0.94", "El suelo desnudo se concentra en pocos bloques del piso A"],
  ["Diversidad estructural", "0.76", "Dosel leñoso abierto frente a la referencia"],
  ["Funciones (MSAVI)", "0.81", "Vigor relativo al P90 del piso"],
  ["Intercambios externos", "0.62", "Paisaje circundante fragmentado en el piso bajo"],
  ["Referencia piso A (46 parcelas)", "—", "55.6 % leñosa, 16.7 % suelo desnudo, MSAVI P90 0.446; confianza media"],
  ["Referencia piso B (25 parcelas)", "—", "75.8 % leñosa, 6.4 % suelo desnudo, MSAVI P90 0.674; confianza media-baja"],
  ["Referencia pisos C–D (16 parcelas)", "—", "68.8 % leñosa, 3.8 % suelo desnudo, MSAVI P90 0.679; confianza baja"],
]));
a.push(fuente("Fuente: Informe E6, Tablas 5 y 6."));
a.push(figura(FIG + "fig_coberturas.png", 15.5));
a.push(N.figura("Cobertura al final de la época seca en los 117 bloques (mosaico S2DR4 de 1 m)"));
a.push(fuente("Fuente: Informe E6, Tabla 7. Estratos aparentes: arbóreo 28.4 %, arbustivo 18.4 %, herbáceo 33.5 %, suelo desnudo 12.8 %, sin lectura 6.9 %."));
a.push(para("Para resumir la cobertura clasificada por bloque, el Informe E6 la lee en parcelas virtuales de 0.5 ha (49 puntos sobre el ráster de 1 m). Bajo la metodología del plan institucional, estas lecturas son un resumen estadístico del píxel clasificado y un control de su variabilidad por estrato, no unidades de muestreo de campo (Cuadro N° " + N.sigC() + ")."));
a.push(N.cuadro("Lectura de la cobertura clasificada en parcelas virtuales por estrato"));
a.push(tabla([{ w: 2700 }, { w: 1300 }, { w: 1000, align: R }, { w: 1300, align: R }, { w: 900, align: R }, { w: 1100, align: R }, { w: 1338, align: C }], ["ECOSISTEMA", "CONDICIÓN MSAVI", "PARCELAS", "% COB. VEG.", "CV (%)", "ERROR (%)", "¿CUMPLE 15 %?"], [
  ["Bosque estacionalmente seco (Bes-cm)", "Bajo umbral", "789", "79.4", "38.0", "2.7", "Sí"],
  ["Bosque estacionalmente seco (Bes-cm)", "Sobre umbral", "142", "96.9", "7.6", "1.2", "Sí"],
  ["Bosque relicto montano (Br-mvoc)", "Sobre umbral", "27", "96.4", "10.6", "4.0", "Sí"],
  ["Matorral andino (Ma)", "Bajo umbral", "10", "67.9", "34.0", "21.1", "No"],
  ["Matorral andino (Ma)", "Sobre umbral", "152", "96.2", "9.1", "1.4", "Sí"],
  ["Páramo (Pa)", "Sobre umbral", "12", "96.3", "8.2", "4.6", "Sí"],
  ["Zona agrícola / mosaico agroforestal", "Sobre umbral", "233", "97.7", "6.4", "0.8", "Sí"],
]));
a.push(fuente("Fuente: Informe E6, Tabla 11. Se leyeron 1,374 parcelas en bloques (1,365 con lectura válida) y 102 de referencia. Contraste con la cobertura de 82 fichas F-DT: r = 0.72; error absoluto medio 19.5 pp; sesgo +15.9 pp."));
a.push(N.cuadro("Quince bloques con menor índice de condición"));
a.push(tabla([{ w: 900, align: C }, { w: 1900 }, { w: 1100, align: R }, { w: 900, align: R }, { w: 1100, align: R }, { w: 1000, align: R }, { w: 800, align: R }, { w: 1100, align: C }, { w: 838, align: C }], ["BLOQUE", "DISTRITO", "ÁREA (ha)", "MSAVI", "% BAJO UMBRAL", "% SUELO", "IC", "CLASE", "TIPOL."], [
  ["85", "San Juan de Bigote", "11.91", "0.229", "100.0", "99.0", "0.22", "Severo", "A"],
  ["84", "San Juan de Bigote", "5.10", "0.252", "100.0", "96.6", "0.24", "Severo", "A"],
  ["87", "San Juan de Bigote", "62.22", "0.249", "99.7", "92.6", "0.26", "Severo", "A"],
  ["86", "San Juan de Bigote", "30.00", "0.273", "99.6", "90.9", "0.27", "Severo", "A"],
  ["77", "San Juan de Bigote", "17.82", "0.272", "99.8", "75.8", "0.30", "Severo", "A"],
  ["83", "San Juan de Bigote", "24.34", "0.305", "99.4", "66.3", "0.40", "Severo", "A"],
  ["81", "Canchaque", "8.93", "0.275", "100.0", "76.4", "0.41", "Moderado", "A"],
  ["M3B5", "San Juan de Bigote", "52.74", "0.283", "99.9", "48.1", "0.43", "Moderado", "A"],
  ["M2B5", "Salitral", "30.91", "0.305", "99.0", "56.5", "0.47", "Moderado", "A"],
  ["M4B4", "Huarmaca", "289.91", "0.320", "99.2", "56.5", "0.50", "Moderado", "A"],
  ["M17B1", "Chulucanas", "42.45", "0.332", "99.8", "17.5", "0.50", "Moderado", "A"],
  ["31", "Huarmaca", "47.58", "0.436", "78.4", "34.5", "0.52", "Moderado", "C"],
  ["M4B3", "Huarmaca", "234.07", "0.322", "98.6", "46.4", "0.52", "Moderado", "A"],
  ["M7B3", "Salitral", "65.75", "0.298", "99.8", "27.5", "0.53", "Moderado", "A"],
  ["M27B1", "Chulucanas", "448.80", "0.363", "96.9", "22.4", "0.57", "Moderado", "A"],
], { size: 15 }));
a.push(fuente("Fuente: Informe E6, Tabla 14. Tipología MSAVI–NDVI: A degradación convergente; C núcleos degradados."));
a.push(para("La revisión confirma que el E6 aporta por primera vez una medida de condición por bloque y una delimitación interna cuantificada (restauración activa 5,566.89 ha; enriquecimiento 5,032.34 ha; conservación 1,207.99 ha; tamizaje agrícola 463.01 ha en los 117 bloques), y que la regla «MSAVI para jerarquizar, NDVI para delimitar» que recoge el plan institucional se valida en el ámbito: los 16 bloques de Morropón que el NDVI 2025 mostraba enmascarados por verdor pasan a degradación convergente al final de la época seca. Las observaciones de consistencia, de alcance y de alineación con el plan institucional se registran en el numeral 3-A.5."));

a.push(h2("3-A.5 Matriz de observaciones del Diagnóstico del Estado de los Ecosistemas"));
a.push(N.cuadro("Observaciones formuladas al Diagnóstico del Estado de los Ecosistemas (E6) y a sus insumos"));
const OE = [
  ["OE-01", "Insumos OS 1148 / 2797", "Listas florísticas con errores taxonómicos y sin depurar.", "Depurar contra F-DT, F-DS y chequeos de suelos.", "Atendida parcialmente: matriz de 69 especies; 48 taxones confirmados en campo 2026"],
  ["OE-02", "Insumo OS 2726", "Presiones evaluadas solo en 58 bloques.", "Extender el atributo de amenazas a los 117 bloques.", "Atendida: IC con F-DT-04 de los 117 bloques"],
  ["OE-03", "Informe E6, ficha técnica y 9.2", "El informe cita como referencia metodológica el «E6 v2.0 sin parcelas» y toma de él las clases del IC; el plan institucional v2 dejó sin efecto esa modalidad y su programación al 18/12/2026.", "Citar como referencia el Plan de Trabajo institucional v2 y la Nota Técnica N° 002-2026; mantener el criterio del IC como propio del E6; fecha de redacción final 18/11/2026.", "Pendiente (informe final)"],
  ["OE-04", "Nota Técnica 002-2026, numeral 1", "La versión revisada cita el «catálogo V7»; el catálogo vigente es V6.", "Corregir la referencia al catálogo V6.", "Pendiente"],
  ["OE-05", "Nota Técnica 002-2026, Cuadro 2", "La cobertura 7 no registra superficie ni denominación.", "Confirmar si se fusionó con otra clase y declararlo.", "Pendiente"],
  ["OE-06", "Informe E6, 9.6, 10.8 y 11; plan v2, 7.2", "Matriz de confusión formal (≈ 600 puntos ciegos), TAU y estructura vertical no ejecutadas; el plan institucional las traslada al estudio definitivo.", "Declararlas en el informe final como vacíos trasladados al estudio definitivo y someterlas a la conformidad de IN/SESDI (H7b).", "Encauzada por el plan v2; pendiente de conformidad"],
  ["OE-07", "Informe E6, Tabla 11", "El estrato de matorral andino bajo umbral tiene 10 de 20 lecturas requeridas (error 21.1 %).", "Densificar la lectura del estrato sobre el ráster de 1 m hasta el error de 15 %, o declarar la limitación.", "Pendiente"],
  ["OE-08", "Informe E6, 10.6", "La referencia del piso A está intervenida (16.7 % de suelo desnudo) y la de los pisos C–D tiene 16 parcelas con confianza baja.", "Incorporar el ACR Bosques Secos de Salitral–Huarmaca y densificar referencias de altura.", "Pendiente"],
  ["OE-09", "Informe E6, 10.9; plan v2, 1.4 d) y Tabla 8", "Coexisten tres criterios de brecha: areal del E6 (8,341.82 ha), media del bloque de los volúmenes DT (7,637.14 ha en 50 bloques) y media del MSAVI por ecosistema con umbral por ecosistema que propone el plan institucional.", "Adoptar el criterio único con la aprobación del plan institucional y recalcular el indicador de los 117 bloques.", "Pendiente (decisión de SESDI)"],
  ["OE-10", "Informe E6, 10.3 y Nota Técnica", "Las coberturas 4 y 5 de fin de época seca incluyen bosque caducifolio sano sin hojas.", "Adquirir compuesto S2DR4 de época de lluvias (feb–abr 2027) y contrastar con referencias.", "Pendiente (estudio definitivo)"],
  ["OE-11", "Informe E6, 10.5", "Cerca del 90 % de las categorías del D.S. N° 043-2006-AG y los atributos de sequía, raíz y palatabilidad figuran «Por determinar»; tres especies se citan como CR «según fuentes del proyecto».", "Completar con consulta a especialistas y verificar contra el anexo oficial del decreto.", "Pendiente"],
  ["OE-12", "Informe E6, 10.3", "Sesgo de +15.9 pp de la cobertura S2DR4 frente a F-DT por el herbazal seco.", "Declarar el sesgo y su efecto en el IC.", "Atendida: declarado en el informe"],
  ["OE-13", "Plan v2, Tabla 5 (H7b)", "No se ha emitido la conformidad de IN/SESDI sobre la metodología de firma espectral y los ítems adaptados del índice R00.", "Emitir la conformidad sobre la base de la Nota Técnica N° 002-2026 (riesgo R04).", "Pendiente"],
  ["OE-14", "Plan v2, Tablas 4 y 12; matrices del plan", `El avance del E6 figura como 39 % en el plan (corte 08/10/2026) y como ${fmt(pond(4), 1)} % en la hoja de avance de las matrices; los diez componentes registran avances distintos.`, "Homologar el avance del E6 al mismo corte en el plan y en sus matrices.", "Pendiente"],
  ["OE-15", "Informe E6, 7.1", "El informe deja «por designar» a especialistas de flora, viveros, GdR y social; el plan institucional fija un equipo base sin contratación de personal.", "Alinear la nota del equipo al plan v2 y asignar la oferta de viveros según la reasignación de funciones (R05).", "Pendiente"],
  ["OE-16", "Informe E6, adecuación de ítems 9.5 y 10.7", "El informe presenta las parcelas virtuales como unidad de muestreo; el plan institucional define el píxel de 1 m como unidad de análisis.", "Presentar las parcelas virtuales como resumen estadístico de la cobertura clasificada, conforme a la hoja «E6_Firma_Espectral».", "Pendiente"],
];
a.push(tabla([{ w: 750, align: C }, { w: 1550 }, { w: 2900 }, { w: 2338 }, { w: 2100 }], ["CÓD.", "DOCUMENTO / SECCIÓN", "OBSERVACIÓN", "LEVANTAMIENTO REQUERIDO", "ESTADO AL CORTE"], OE, { size: 15 }));
a.push(fuente("Fuente: revisión técnica del suscrito sobre el Plan de Trabajo institucional v2, la Nota Técnica N° 002-2026 y el Informe E6 (versión para revisión de SESDI)."));
a.push(h2("3-A.6 Estado de la conformidad"));
a.push(para("El Diagnóstico del Estado de los Ecosistemas se encuentra en versión para revisión de SESDI, con redacción final programada al 18/11/2026. La conformidad está condicionada a: (i) la conformidad de IN/SESDI sobre la metodología de firma espectral y los ítems adaptados del índice R00 (hito H7b, pendiente; riesgo R04); (ii) el levantamiento de las observaciones OE-03 a OE-05, OE-07 a OE-09, OE-11 y OE-13 a OE-16; y (iii) la decisión sobre el criterio único de brecha. La matriz de confusión formal, la TAU, la estructura vertical medida y el compuesto de época de lluvias (OE-06 y OE-10) se trasladan como vacíos declarados al estudio definitivo, conforme al plan institucional."));

// ---------------- 3-B ----------------
const bb = [];
bb.push(h1("3-B. ESTUDIO DE OFERTA DE MATERIAL VEGETAL DE PROPAGACIÓN"));
bb.push(h2("3-B.1 Plan de trabajo del componente"));
bb.push(para("En el plan institucional el componente de oferta comprende cuatro líneas de la estructura del E6, que suman 35 de los 100 puntos del entregable: E6.5, lista corta de especies nativas para GRD y fichas técnicas (5); E6.6, demanda de plantones por especie y bloque para los 117 bloques (10); E6.7, directorio de viveros, fuentes semilleras y brecha de viveros (14); y E6.8, propuesta de viveros permanentes y temporales, propagación y costos (6). Sus productos para el E10 —especies priorizadas por piso altitudinal y brecha de viveros por escenarios— se entregan con el informe final del 18/11/2026 (hito H7c)."));
bb.push(para("Dos decisiones del plan condicionan el componente. La primera: la cifra de plantones para los bloques no se fija en esta etapa; se proporcionará en el E10 como parte del diseño preliminar de intervenciones (comentario N° 14 del Anexo 4 del plan). La segunda: con la plaza de viveros desierta y sin asignación presupuestal, la oferta y la brecha de viveros quedan a cargo del equipo base del E6 hasta que se formalice la reasignación de funciones entre los especialistas de Uso Sostenible, Infraestructura Verde e Infraestructura Marrón (plazo 23/10/2026)."));
bb.push(para([b("Método de cálculo. "), "El Informe E6 (numeral 9.7) plantea la demanda como D = A · d · p · (1 + r) y la brecha como B = D − O, donde A son las hectáreas por medida (restauración activa y enriquecimiento según el MSAVI), d la densidad de plantación, p la proporción de la especie en la mezcla, r la tasa de reposición y O la oferta efectiva. Los parámetros d, p y r los fija el E10; el censo remoto de viveros (encuesta y verificación en imagen) completará la oferta."]));
bb.push(h2("3-B.2 Revisión técnica de los insumos de oferta"));
bb.push(para("Los entregables de las OS 1147 (oferta de material vegetal) y 2726 (fitosanidad y propagación), revisados el 03/10/2026, aportan una priorización preliminar de 35 especies arbóreas y 34 arbustivas sobre ≈ 77 nativas, protocolos de propagación de 10 especies, un protocolo de vivero limpio y un costo paramétrico de S/ 12 por plantón. Sus limitaciones para el Perfil son que la demanda de plantones (2.94 millones) se calculó solo para Lotes SUS en 83 bloques, que las cinco fuentes de oferta identificadas en gabinete no están en los 15 distritos y que la única oferta confirmada son 6,000 plantones (Suyo)."));
bb.push(h2("3-B.3 Resultados del E6 sobre la oferta"));
bb.push(N.cuadro("Viveros e iniciativas de producción de plantones identificados"));
bb.push(tabla([{ w: 3000 }, { w: 2600 }, { w: 1700, align: C }, { w: 2338 }], ["VIVERO O INICIATIVA", "PROVINCIA / LUGAR", "CAPACIDAD DECLARADA", "ESPECIES"], [
  ["Viveros del GORE Piura en Hualtacal", "Huancabamba; Hualtacal (Canchaque)", "21,000 plantones", "Nativas: hualtaco, algarrobo, roble, laurel, aliso"],
  ["Vivero Forestal Hualtacal (Proyecto Pusmalca)", "Huancabamba; Canchaque", "No declarada", "—"],
  ["Vivero del GORE en Maraypampa", "Huancabamba; Maraypampa", "No declarada", "Material vegetal nativo"],
  ["Vivero forestal en el entorno de M22B1", "Huancabamba; San Miguel de El Faique", "Por determinar", "Aliso, queñua, cedro de altura, pino radiata"],
  ["Reforestación FONCODES (Haku Wiñay, N. E. Canchaque 02)", "Huancabamba; A. Miguel Grau y Huabal", "No declarada", "Roble y laurel"],
  ["Terreno ofrecido para vivero (Serrán)", "Morropón; Salitral", "No declarada", "—"],
  ["Vivero regional (GORE Piura)", "Morropón", "No declarada", "—"],
  ["Entrega de plantones ACR Salitral–Huarmaca", "Morropón; La Alberca, La Tranca, Palo Blanco–El Cerezo, Hornopampa", "No declarada", "—"],
  ["Iniciativas municipales de reforestación", "Morropón; Yamango", "No declarada", "Palo santo y bambú"],
  ["Proyecto de vivero del gobierno regional (Bigote)", "Morropón; San Juan de Bigote", "No declarada", "—"],
  ["Proyecto de viveros (Víctor Raúl / El Checo)", "Morropón; Yamango", "No declarada", "Café, cacao, frutales"],
  ["Huertos y viveros FONCODES (Nuevo Progreso)", "Morropón; Salitral", "36 usuarios (no es capacidad)", "—"],
], { size: 15 }));
bb.push(fuente("Fuente: Informe E6, Tabla 15, sobre fichas DT, matrices del E2 y consolidado DS."));
bb.push(para("La única capacidad cuantificada son los 21,000 plantones de los viveros del Gobierno Regional en Hualtacal, el mismo dato que recoge el plan institucional (numeral 7.2), frente a una demanda potencial que se dimensionará sobre 5,503.41 ha de restauración activa y 2,838.41 ha de enriquecimiento. El E6 incluye 16 fichas técnicas de especies nativas recomendadas: Vachellia macracantha (aromo), Eriotheca ruizii (pasallo), Loxopterygium huasango (hualtaco), Bursera graveolens (palo santo), Caesalpinia paipai (charán), Cordia lutea (overo), Ceiba trischistandra (ceibo), Colicodendron scabridum (sapote), Prosopis pallida (algarrobo), Handroanthus chrysanthus (guayacán), Caesalpinia spinosa (tara), Inga feuilleei (guaba), Alnus acuminata (aliso), Capparicordis crotonoides (bichayo), Baccharis latifolia (chilca) y Tecoma stans (huaranhuay). La demanda y la brecha se calculan en la hoja 16 de las matrices al ingresar las densidades, la mezcla y la reposición del E10; no se asumieron valores."));
bb.push(h2("3-B.4 Matriz de observaciones del Estudio de Oferta de Material Vegetal"));
bb.push(N.cuadro("Observaciones formuladas al Estudio de Oferta de Material Vegetal de Propagación"));
const OM = [
  ["OM-01", "Insumo OS 1147; plan v2, E6.6", "La demanda de plantones (2.94 M) se calculó solo para Lotes SUS en 83 bloques.", "Calcular la demanda de los 117 bloques en el E10 con las densidades, la mezcla y la reposición (hoja 16), conforme al plan institucional.", "Pendiente (E10)"],
  ["OM-02", "Insumos de oferta", "Ninguno de los viveros de gabinete está en los 15 distritos; oferta confirmada de 6,000 plantones.", "Ampliar el directorio al ámbito.", "Atendida parcialmente: 12 iniciativas; una con capacidad declarada (21,000)"],
  ["OM-03", "Insumo OS 2726", "Costo paramétrico de S/ 12 por plantón sin análisis de costos unitarios.", "Sustentar el costo por especie y tipo de vivero en el E10/E12.", "Pendiente"],
  ["OM-04", "Informe E6, 9.7; plan v2, E6.7", "El censo remoto de viveros (encuesta y verificación en imágenes) no se ha ejecutado.", "Ejecutar el censo con el equipo base y calcular la oferta efectiva por especie antes del 18/11/2026.", "Pendiente"],
  ["OM-05", "Plan v2, 1.4 c) y R05", "Plaza de Especialista en Infraestructura Rural para la Producción (viveros) desierta y sin asignación presupuestal.", "Formalizar la reasignación de funciones entre ESUS, EIV y EIM antes del 23/10/2026.", "Pendiente al corte de información del plan (16/10/2026)"],
  ["OM-06", "Plan v2, E6.7", "No se ha elaborado el mapa de fuentes de semilla (árboles semilleros y rodales).", "Elaborarlo con viveristas, saberes ancestrales y registros de SERFOR.", "Pendiente"],
  ["OM-07", "OS 1147 / E6 / E5", "La priorización de 35 arbóreas y 34 arbustivas (OS 1147) y la lista del E6 no están conciliadas con las especies de la matriz suelo–intervención del E5.", "Conciliar una lista única por piso, ecosistema y grupo de suelos para la entrega H7c.", "Pendiente"],
  ["OM-08", "Informe E6, Tabla 15", "El vivero del entorno de M22B1 produce pino radiata, especie exótica.", "Excluir de la oferta efectiva las especies no nativas o incompatibles.", "Criterio a aplicar en el cálculo"],
];
bb.push(tabla([{ w: 750, align: C }, { w: 1550 }, { w: 2900 }, { w: 2338 }, { w: 2100 }], ["CÓD.", "DOCUMENTO / SECCIÓN", "OBSERVACIÓN", "LEVANTAMIENTO REQUERIDO", "ESTADO AL CORTE"], OM, { size: 15 }));
bb.push(fuente("Fuente: revisión técnica del suscrito sobre los insumos de las OS 1147 y 2726, el Plan de Trabajo institucional v2 y el Informe E6."));
bb.push(h2("3-B.5 Estado de la conformidad"));
bb.push(para("El componente de oferta de material vegetal es el de menor avance del E6 (E6.7: 10 % en el plan, 50 % en sus matrices; E6.8: 20 % y 25 %): el directorio está identificado, pero el censo remoto, la oferta efectiva por especie, el mapa de fuentes semilleras y la brecha de viveros están pendientes. No corresponde conformidad al corte; su emisión requiere el levantamiento de las observaciones OM-02 a OM-07 y la entrega al E10 de las especies priorizadas y de la brecha de viveros por escenarios con el informe final (hito H7c, 18/11/2026). La cifra de plantones de los bloques (OM-01) se resolverá en el E10."));

// ---------------- Cierre ----------------
const z = [];
z.push(h1("4. CONCLUSIONES"));
[
  "El Diagnóstico del Estado de los Ecosistemas y el Estudio de Oferta de Material Vegetal se ejecutan como un solo Entregable 6 por Administración Directa con la metodología de firma espectral, conforme al Plan de Trabajo Detallado del Perfil – Actualización N° 01, Versión 2 (16/10/2026): equipo base, sin servicio externo ni contratación de personal, periodo 12/10–18/11/2026 y fuera de la ruta crítica.",
  "El plan del E6 se revisó y actualizó en tres estados —versión 1.1, versión 2.0 «In house» y plan institucional v2—; la programación de 10 semanas y la dotación de personal del E6 v2.0 quedan sin efecto, y los comentarios del suscrito sobre el E6 fueron absueltos en el Anexo 4 del plan institucional.",
  "La clasificación de 92 tomas S2DR4 concluyó el 13/10/2026 (hito H6g) y la Nota Técnica N° 002-2026 fue remitida; la conformidad de IN/SESDI sobre la metodología (hito H7b) es la condición principal del E6.",
  "El diagnóstico cuantifica 8,341.82 ha (68.0 %) bajo el umbral MSAVI y 5,503.41 ha para restauración activa, y clasifica 14 bloques como moderados y 6 como severos; la matriz de confusión formal, la TAU y la estructura vertical pasan al estudio definitivo.",
  "Se formularon 16 observaciones al diagnóstico (OE) y 8 a la oferta (OM), de las cuales 4 están atendidas total o parcialmente; cuatro observaciones —OE-03 reformulada y OE-14 a OE-16 nuevas— alinean el Informe E6 y las matrices con el plan institucional.",
  "La oferta de material vegetal identificada es mínima (21,000 plantones) frente a la escala de la intervención; la reasignación formal de las funciones de viveros y el censo de viveros son la ruta crítica del componente.",
].forEach((t, k) => z.push(num(`4.${k + 1}.`, t)));
z.push(h1("5. RECOMENDACIONES"));
[
  "Emitir la conformidad de IN/SESDI sobre la metodología de firma espectral y los cinco ítems adaptados del índice R00 (hito H7b), sobre la base de la Nota Técnica N° 002-2026.",
  "Alinear el informe final del E6 al plan institucional v2: referencia metodológica, nota del equipo, unidad de análisis (píxel de 1 m) y vacíos trasladados al estudio definitivo, y homologar el avance del E6 entre el plan y sus matrices.",
  "Adoptar el criterio único de brecha con la aprobación del plan institucional (media del MSAVI por ecosistema frente al umbral 0.4976, con umbral por ecosistema por consenso del equipo de especialistas) y recalcular el indicador de los 117 bloques.",
  "Densificar la lectura del estrato de matorral andino bajo umbral y las referencias de altura, e incorporar el ACR Bosques Secos de Salitral–Huarmaca como referencia del piso A.",
  "Formalizar la reasignación de las funciones de viveros, ejecutar con el equipo base el censo remoto y el mapa de fuentes semilleras, y entregar al E10 las especies priorizadas y la brecha de viveros por escenarios el 18/11/2026.",
  "Conciliar una lista única de especies por piso, ecosistema y grupo de suelos entre el E6, la matriz suelo–intervención del E5 y los Lotes SUS, en coordinación con el Especialista en Intervenciones Vegetativas de Uso Sostenible.",
].forEach((t, k) => z.push(num(`5.${k + 1}.`, t)));
z.push(h1("6. MEDIOS VERIFICADORES"));
z.push(N.cuadro("Medios verificadores de la actividad d)"));
z.push(tabla([{ w: 3500 }, { w: 4138 }, { w: 2000, align: C }], ["MEDIO VERIFICADOR", "UBICACIÓN / IDENTIFICADOR", "FECHA"], [
  ["Plan de Trabajo Detallado del Perfil – Actualización N° 01, Versión 2 (documento rector; numeral 7.2, Tablas 5, 6, 12, 13, 16, 17, 20 y Anexo 4)", "Google Drive ID 1QkhMG4cWkPFAyhxOH8qtaWG2YVCCJsGy", "16–17/10/2026"],
  ["Matrices del Plan de Trabajo (hojas «E6_Firma_Espectral», «Hitos», «Riesgos» y «Control_Cambios»)", "Google Drive ID 1EF3u3btVKiK2DcUKrHYcJZv-4co_arCc", "16/10/2026"],
  ["Nota Técnica N° 002-2026-ANIN/DIME/SESDI (original y revisada)", "Google Drive ID 1EZk6-sdLsbFfkiuI4jma4IxlTVsTSXyT; 137Syk2NZPG9bTZ8rwrxtpTmjbzOtLhTy; PDF 17Vt9Z1NBZPTGF933r4jhqFNu0N7MP-nR", "13–14/10/2026"],
  ["Clasificación no supervisada de las tomas S2DR4 (resultados y LEEME)", "Google Drive ID 1oZ-PLyuw4QyIBe-fupK3j7lQY_DE48J8; 1NRGviLNkpnvfyrLfv3dVNZms4w3SbfuS", "13–14/10/2026"],
  ["Informe E6 – versión para revisión de SESDI", "Google Drive ID 1NfYjSV3QXI3uA4eHsMNQNTw9vdY5RdDA", "15/10/2026"],
  ["Matrices del diagnóstico del ecosistema (E6)", "Google Drive ID 1Uie1il6FeHI4SD8QhKXsfDof7cm4NF8J", "15/10/2026"],
  ["E6 v2.0 «In house» – original y revisión del suscrito (antecedente, sin efecto en su programación)", "Google Drive ID 1B3QsYMkodQVnK5ZwQU-7BI0t9Yn6xmoR; 1rxdf63E-TOJpUEUWWazdkth6WjdwU4HY; 1lbr4-I5-OH8Nqd4hBccIEizEy3H-rV7V; 18FQxeA3WasWZOCjg4mVapnjQf6D_Est6", "08–14/10/2026"],
  ["Tablero de avance del Perfil (revisión de entregables OS 1147, 1148, 2357, 2726 y 2797)", "Google Drive ID 1wRdIvYH8ni8kbj3XsDKa0cgrgSIdVWn6", "03/10/2026"],
  ["Entregables de las OS insumo del E6", "Google Drive: PRIMER/SEGUNDO/TERCER_ENTREGABLE de las OS 1147, 1148, 2357, 2726 y 2797", "22/09/2026"],
  ["Referencia técnica de fuentes semilleras", "Google Drive ID 1hS8QlbTu_Ks05061BV4FAbAqbfUpikzy («Fuentes semilleras de los Bosques Secos»)", "—"],
  ["Matrices de observaciones en formato nativo", "Anexo 4 (Excel), hojas «Obs_E6_Ecosistemas» y «Obs_E6_Material_vegetal»", "Octubre 2026"],
], { size: 15 }));
z.push(fuente("Fuente: elaboración propia."));

if (require.main === module) {
  L.guardar([
    seccionCaratula(DEF.anexo3),
    L.seccion(x, { pie: PIE }),
    seccionCaratula(DEF.anexo3A),
    L.seccion(a, { pie: PIE }),
    seccionCaratula(DEF.anexo3B),
    L.seccion([...bb, ...z], { pie: PIE }),
  ], OUT, "Anexo N° 3 – Actividad d)");
}
module.exports = { OE, OM };
