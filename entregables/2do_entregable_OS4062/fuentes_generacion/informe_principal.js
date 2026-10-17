// Informe N° 008-2026-HSCM — Segundo Entregable OS 4062-2026 (carta + informe principal)
const L = require("./lib");
const { para, b, i, run, h1, h2, h3, literal, guion, num, fuente, tabla, meta, salto, vacio, centro, W, C, R, J, fmt, figura } = L;
const { Paragraph, AlignmentType } = require("docx");
const D = require("./datos.json");
const OUT = process.argv[2];
const FIG = "/tmp/claude-0/-home-user-Salomon1969/03b82ed8-c0e9-54ec-93c9-bff5f93b575b/scratchpad/build/fig/";
const H = D.prov["Huancabamba"], M = D.prov["Morropón"];
const N = L.numerador();
const PIE = "Informe N° 008-2026-HSCM · Segundo Entregable · Orden de Servicio N° 4062-2026";
const FECHA = "Lima, octubre de 2026";
const TOT_HA = H.ha + M.ha;

// ---------- CARTA ----------
const carta = [
  new Paragraph({ alignment: R, spacing: { before: 200, after: 240 }, children: [run(FECHA)] }),
  para([b("CARTA N° 014 – 2026 – HSCM")], { spacing: { after: 200 } }),
  para("Señor:", { spacing: { after: 0 } }),
  para([b("Ing. CÉSAR ALFREDO GONZÁLES ALFARO")], { spacing: { after: 0 } }),
  para("Ejecutivo (e) de la Subdirección de Estudios de Inversión", { spacing: { after: 0 } }),
  para("Dirección de Intervenciones Multisectoriales y de Emergencia", { spacing: { after: 0 } }),
  para("Autoridad Nacional de Infraestructura – ANIN", { spacing: { after: 0 } }),
  para("Presente.–", { spacing: { after: 240 } }),
  meta([["ASUNTO", "Presentación del Segundo Entregable – Orden de Servicio N° 4062-2026"],
        ["REFERENCIA", "a) Orden de Servicio N° 0004062, de fecha 28 de agosto de 2026\nb) Términos de Referencia – Anexo N° 3, Pedido de Servicio N° 8335, del 27 de agosto de 2026\nc) Carta N° 013-2026-HSCM – Primer Entregable (Informe N° 007-2026-HSCM)"]], 1900),
  ...vacio(1),
  para("Tengo el agrado de dirigirme a usted, con la finalidad de presentarle el Segundo Entregable correspondiente a la Orden de Servicio de la referencia: «Servicio especializado para la determinación de área geográfica para el componente de infraestructura verde del estudio de preinversión del proyecto de Infraestructura Natural Piura (Código de Idea: 337830)», para su trámite respectivo."),
  para("El presente entregable comprende el informe sobre las acciones y/o actividades de los puntos b), c) y d) descritos en el numeral 5.1 del ítem 5 de los Términos de Referencia, correspondiente al periodo comprendido en el segundo entregable, adjuntando los medios verificadores correspondientes."),
  para("Se adjunta a la presente la siguiente documentación:"),
  guion("Informe N° 008-2026-HSCM y sus Anexos 1, 2, 3 y 4"),
  guion("Recibo por Honorarios Electrónico"),
  guion("Validez del comprobante de pago"),
  guion("Orden de Servicio N° 4062-2026"),
  guion("Términos de Referencia"),
  guion("Constancia de afiliación a seguro (consulta SUSALUD vigente) y SCTR"),
  ...vacio(1),
  para("Sin otro particular, hago propicia la ocasión para expresarle los sentimientos de mi especial consideración."),
  para("Atentamente,"),
  ...L.firma(),
];

// ---------- INFORME ----------
const inf = [];
inf.push(centro([b("INFORME N° 008-2026-HSCM", { size: 28 })], { spacing: { before: 120, after: 240 } }));
inf.push(meta([
  ["Dirigido a", "Ing. CÉSAR ALFREDO GONZÁLES ALFARO\nEjecutivo (e) de la Subdirección de Estudios de Inversión\nDirección de Intervenciones Multisectoriales y de Emergencia"],
  ["De", "Ing. HÉCTOR SALOMÓN CAHUAS MILLER\nIngeniero Forestal – Especialista en Infraestructura Verde"],
  ["Asunto", "Remite Segundo Entregable"],
  ["Referencia", "Orden de Servicio N° 4062-2026"],
  ["Periodo", "Del 26 de septiembre al 27 de octubre de 2026 (días 31 a 60 de la Orden de Servicio)"],
  ["Fecha", FECHA]], 1900));
inf.push(...vacio(1));
inf.push(para("Es grato dirigirme a usted para saludarlo cordialmente y remitir adjunto al presente el informe correspondiente al Segundo Entregable del «Servicio especializado para la determinación de área geográfica para el componente de infraestructura verde del estudio de preinversión del proyecto de Infraestructura Natural Piura (Código de Idea: 337830)», para la elaboración del Estudio de Preinversión del proyecto Infraestructura Natural Piura, con Código Único de Inversión N° 2669244, en lo correspondiente al componente de infraestructura verde."));

// 1) INTRODUCCIÓN
inf.push(h1("1) INTRODUCCIÓN"));
inf.push(para("Conforme a lo señalado en el numeral 5 de los Términos de Referencia de la Orden de Servicio N° 4062-2026 (Características y Condiciones de la Contratación), numeral 5.1 Actividades, las actividades que realiza el suscrito durante la prestación del servicio son:"));
const ACT = {
  a: "a) Verificar en campo de la información biofísica de los bloques pendientes en los distritos de la provincia de Huancabamba.",
  b: "b) Actualizar, integrar y consolidar la información geoespacial y el diagnóstico territorial de los bloques preliminares de intervención ubicados en los distritos de la provincia de Huancabamba y Morropón.",
  c: "c) Realizar el seguimiento y reportar el avance en coordinación con el Especialista en Infraestructura Marrón, del Estudio de Reconocimiento de Suelos, incluyendo la revisión técnica, la formulación y el levantamiento de observaciones, así como la emisión de la conformidad —cuando corresponda— respecto del plan de trabajo, las actividades de campo y los informes presentados, verificando que la información generada cumpla con los requerimientos técnicos necesarios para su aplicación en el componente verde del proyecto IN Piura.",
  d: "d) Realizar el seguimiento y reportar el avance, en coordinación con el Especialista en Intervenciones de Uso Sostenible, del Diagnóstico del Estado de los Ecosistemas y del Estudio de Oferta de Material Vegetal de Propagación, incluyendo la revisión y actualización de sus respectivos planes de trabajo, el seguimiento de las actividades de campo y la revisión técnica de los entregables, así como la formulación y el levantamiento de observaciones y la emisión de la conformidad, cuando corresponda.",
};
Object.values(ACT).forEach((t) => inf.push(guion(t)));
inf.push(para("Asimismo, de acuerdo con el numeral 6 de los Términos de Referencia (Entregables), el Segundo Entregable comprende un informe sobre las acciones y/o actividades de los puntos b), c) y d) descritos en el numeral 5.1 del ítem 5, correspondiente al periodo comprendido en el segundo entregable, adjuntando medios verificadores, hasta los sesenta (60) días calendario contados a partir del día siguiente de notificada la Orden de Servicio."));
inf.push(para("En consecuencia, el presente informe desarrolla las actividades b), c) y d) del numeral 5.1 durante el periodo comprendido entre el 26 de septiembre y el 27 de octubre de 2026, cuyos productos y medios verificadores se relacionan en el Cuadro N° 1."));
inf.push(para("La finalidad del presente informe es fortalecer la determinación del área geográfica del componente de infraestructura verde en el marco de la formulación del estudio de preinversión del proyecto IN Piura, aportando la actualización e integración del diagnóstico territorial de las provincias de Huancabamba y Morropón con los estudios especializados disponibles al corte, el seguimiento técnico del Estudio de Reconocimiento de Suelos y el seguimiento del Diagnóstico del Estado de los Ecosistemas y del Estudio de Oferta de Material Vegetal de Propagación, que sustentan la delimitación de las áreas de reforestación y restauración ecosistémica."));
inf.push(para("El desarrollo del informe mantiene la regla de no fabricación de datos adoptada en el primer entregable: ningún valor ha sido estimado, interpolado ni transferido entre bloques; los campos sin sustento documental se consignan como «Por determinar», «Pendiente» o «s/d», y las inconsistencias entre fuentes se registran como observaciones trazables en lugar de resolverse en silencio."));
inf.push(N.cuadro("Correspondencia entre las actividades del Segundo Entregable y sus medios verificadores"));
inf.push(tabla([{ w: 800, align: C }, { w: 3900 }, { w: 4938 }], ["ÍTEM", "ACTIVIDAD SEGÚN NUMERAL 5.1 DEL TdR", "MEDIO VERIFICADOR / PRODUCTO"], [
  ["b)", ACT.b.slice(3), "Anexo 1 — Informe de actualización, integración y consolidación de la información geoespacial y del diagnóstico territorial de Huancabamba y Morropón (Diagnóstico Territorial y Social, Versión 5), con la consolidación por provincia, distrito y bloque, la situación predial de gabinete de 107 bloques, las matrices consolidadas y el registro de versiones publicadas."],
  ["c)", ACT.c.slice(3), "Anexo 2 — Informe de seguimiento técnico del Estudio de Suelos a Nivel de Reconocimiento: estado de los entregables, seguimiento de la fase de campo, revisión técnica del Entregable 2 con su matriz de observaciones, seguimiento del laboratorio y verificación de la aplicación al componente verde."],
  ["d)", ACT.d.slice(3), "Anexo 3 — Informe de seguimiento del Diagnóstico del Estado de los Ecosistemas (3-A) y del Estudio de Oferta de Material Vegetal de Propagación (3-B): revisión y actualización de planes de trabajo, revisión técnica de entregables, observaciones y estado de conformidades."],
  ["b), c), d)", "Información nativa de procesamiento y resultados", "Anexo 4 — Matrices de seguimiento en formato Excel: consolidado por bloque y distrito, situación predial, observaciones del Estudio de Suelos y del E6, y registro de medios verificadores."],
]));
inf.push(fuente("Fuente: elaboración propia sobre los Términos de Referencia, numerales 5.1 y 6."));

// 2) ANTECEDENTES
inf.push(h1("2) ANTECEDENTES"));
const ant = [
  "Mediante el INFORME-000180-2026-ANIN/DIME-SESDI-JCR, la Subdirección de Estudios de Inversión sustentó la necesidad de contratar el servicio especializado para la determinación del área geográfica del componente de infraestructura verde del Proyecto IN Piura, ante la necesidad de contar con información primaria verificada en campo y con la supervisión de estudios especializados que eviten que la delimitación de bloques se base en datos de gabinete no actualizados.",
  "Con fecha 27 de agosto de 2026 se emiten los Términos de Referencia correspondientes al Pedido de Servicio N° 8335 y con fecha 28 de agosto de 2026 se emite la Orden de Servicio N° 0004062 (Expediente SIAF N° 0000008734), por el monto de S/ 20,000.00 (veinte mil y 00/100 soles), con un plazo de ejecución de hasta sesenta (60) días calendario y dos (02) entregables de igual monto, bajo el marco de la Ley N° 32069, Ley General de Contrataciones Públicas, y su Reglamento aprobado por Decreto Supremo N° 009-2025-EF.",
  "Con Carta N° 013-2026-HSCM, del 25 de septiembre de 2026, el suscrito presentó el Primer Entregable (Informe N° 007-2026-HSCM), que desarrolló las actividades a), b) y c): la verificación biofísica de los bloques 79, M4B3 y M4B4 (618.79 ha), la consolidación del diagnóstico territorial de 107 bloques y 10,878.52 ha de Huancabamba y Morropón (Volúmenes I y II, corte al 21 de septiembre de 2026) y el seguimiento del Estudio de Suelos hasta la conformidad de su Entregable N° 1.",
  "El servicio se enmarca en la fase de preinversión (Estudio de Perfil) del Proyecto IN Piura, cuya Unidad Productora es el ecosistema en su totalidad, dentro del cual se delimitan los bloques preliminares de intervención; la intervención se dirige a las zonas de origen de los peligros de la Cuenca Alta del Río Piura. El catálogo vigente de bloques es la versión V6, con 117 bloques y 12,270.235 ha.",
  "El Estudio de Suelos a Nivel de Reconocimiento es ejecutado por GeoSIG Ingenieros E.I.R.L. (procedimiento CP-ABR-6-2026-ANIN/DAI-1, contrato del 17 de agosto de 2026), en cuatro entregables y 75 días calendario. Su Entregable N° 1 cuenta con conformidad de SESDI del 04 de septiembre de 2026 (Informe N° 000251-2026-ANIN/DIME-SESDI-ARV); la fase de campo se ejecutó entre el 03 y el 20 de septiembre de 2026 y su Informe de Avance (Entregable N° 2) fue presentado al vencimiento del día 40 del servicio (26 de septiembre de 2026).",
  "El procedimiento para el Diagnóstico del Ecosistema y Estudio de Oferta de Material Vegetal (CP-ABR-5-2026-ANIN/DAI-1) fue declarado desierto en septiembre de 2026. La Coordinación de Proyectos IN decidió ejecutar el Entregable 6 (E6) por Administración Directa, con el equipo del proyecto y a nivel de perfil. Concurrentemente, los Comunicados Oficiales ENFEN N° 16-2026 (14 de septiembre) y N° 17-2026 (28 de septiembre) mantuvieron la Alerta de El Niño Costero con probable magnitud extraordinaria, lo que obligó a sustituir la campaña de 157 parcelas de campo prevista en la versión 1.1 del plan del E6 por una inspección remota.",
  "La Especialidad de Infraestructura Marrón está a cargo del Ing. Pedro Talledo Hernández y la Especialidad de Intervenciones Vegetativas de Uso Sostenible, del Blgo. Abel Aucasime Orihuela; con ambos se coordinaron las actividades c) y d), respectivamente. La Coordinación de Proyectos IN está a cargo del Ing. Alejandro Falconí Valdivia.",
];
ant.forEach((t, k) => inf.push(num(`2.${k + 1}.`, t)));
inf.push(num("2.8.", "Los datos generales del Proyecto IN Piura se muestran en el Cuadro N° 2."));
inf.push(N.cuadro("Datos generales del Proyecto IN Piura"));
inf.push(tabla([{ w: 3000, bold: true }, { w: 6638 }], ["CAMPO", "DESCRIPCIÓN"], [
  ["Código de Idea", "337830"], ["Código Único de Inversión (CUI)", "2669244"],
  ["Nombre del proyecto", "Recuperación del servicio de regulación de riesgos naturales y recuperación de ecosistemas degradados en la Cuenca Alta del Río Piura"],
  ["Entidad", "ANIN – DIME – Subdirección de Estudios de Inversión (SESDI)"],
  ["Fase actual", "Preinversión – Estudio de Perfil (Invierte.pe · DGPMI-MEF)"],
  ["Modalidad de ejecución", "Administración Directa"],
  ["Ámbito del proyecto", "3 provincias (Morropón, Huancabamba y Ayabaca) y 15 distritos; catálogo V6: 117 bloques, 12,270.235 ha"],
  ["Ámbito de la actividad b)", `2 provincias (Huancabamba y Morropón); 107 bloques y ${fmt(TOT_HA)} ha en 14 distritos y ${H.nmc + M.nmc} microcuencas`],
  ["Área de intervención directa prevista", "Entre 10,000 y 15,000 ha"],
  ["Indicador de brecha", "Porcentaje de superficie de ecosistemas degradados que brindan servicios ecosistémicos que requieren de recuperación (R.M. N° 00213-2024-MINAM); umbral operativo MSAVI 2024 = 0.4976"],
  ["Sistema de referencia", "UTM WGS 84, Zona 17S (EPSG:32717)"],
]));
inf.push(fuente("Fuente: Plan de Trabajo Detallado del Perfil (Actualización N° 02, 16/10/2026), catálogo de bloques V6 y Diagnóstico Territorial y Social, Versión 5."));

// 3) SIGLAS
inf.push(h1("3) SIGLAS Y ACRÓNIMOS"));
inf.push(N.cuadro("Siglas y acrónimos empleados en el presente informe"));
const siglas = [
  ["AHP", "Proceso de Análisis Jerárquico (Analytic Hierarchy Process)"], ["ANIN", "Autoridad Nacional de Infraestructura"],
  ["Bes-cm", "Bosque estacionalmente seco de colina y montaña"], ["C. C.", "Comunidad campesina"], ["COFOPRI", "Organismo de Formalización de la Propiedad Informal"],
  ["CP", "Centro poblado"], ["CUI", "Código Único de Inversiones"], ["CUM", "Capacidad de Uso Mayor de las Tierras"],
  ["DIME", "Dirección de Intervenciones Multisectoriales y de Emergencia"], ["DT / DS", "Diagnóstico Territorial / Diagnóstico Social"],
  ["E1 … E15", "Entregables del Plan de Trabajo Detallado del Perfil (E2 diagnóstico social, E3 geología, E4 hidrología, E5 suelos, E6 ecosistemas y material vegetal, E8 análisis de riesgo, E10 análisis técnico)"],
  ["ENFEN", "Estudio Nacional del Fenómeno El Niño"], ["EPH", "Erosión Potencial Hídrica"],
  ["F-DT / F-DS / F-LA", "Fichas de Diagnóstico Territorial, de Diagnóstico Social y de Liberación de Áreas del aplicativo IN Piura"],
  ["FEN", "Fenómeno El Niño"], ["GDB", "Geodatabase"], ["GEORURAL", "Plataforma de catastro rural del MIDAGRI (comunidades campesinas)"],
  ["IC", "Índice de condición del ecosistema (cinco atributos frente a ecosistemas de referencia)"], ["INF", "Prueba de infiltración de doble anillo (método Muntz)"],
  ["IN", "Infraestructura Natural"], ["ISL", "Índice de Susceptibilidad Litológica (E3)"], ["MCA", "Análisis multicriterio"],
  ["MRR-CCC", "Medidas de Reducción del Riesgo en contexto de Cambio Climático"], ["MSAVI", "Índice de Vegetación Ajustado al Suelo Modificado"],
  ["NDVI", "Índice de Vegetación de Diferencia Normalizada"], ["OS", "Orden de Servicio"], ["PI", "Peligro Integrado"], ["PGI", "Peligro Generador de Inundaciones"],
  ["PISCO", "Peruvian Interpolated data of the SENAMHI's Climatological and Hydrological Observations"], ["PMM", "Peligro por Movimientos en Masa"],
  ["RPAS", "Aeronave pilotada a distancia (dron)"], ["S2DR4", "Sentinel-2 Deep Resolution, imágenes de 1 m"], ["SESDI", "Subdirección de Estudios de Inversión"],
  ["SP", "Lámina de mapa de situación predial por bloque (SP-001 a SP-117)"], ["SUS", "Lote de Uso Sostenible"], ["TAU", "Tendencia de degradación (Mann-Kendall)"],
  ["TdR", "Términos de Referencia"], ["UNALM", "Universidad Nacional Agraria La Molina"], ["UNF", "Universidad Nacional de Frontera"],
  ["UP", "Unidad Productora"], ["USLE", "Ecuación Universal de Pérdida de Suelo"], ["UTM", "Universal Transversal de Mercator"],
];
inf.push(tabla([{ w: 2200, bold: true }, { w: 7438 }], ["TÉRMINO", "DESCRIPCIÓN"], siglas));
inf.push(fuente("Fuente: elaboración propia."));

// 4) ANÁLISIS
inf.push(h1("4) ANÁLISIS Y DESARROLLO DE LAS ACTIVIDADES SOLICITADAS"));
inf.push(para("Descripción del servicio contratado: «Servicio especializado para la determinación de área geográfica para el componente de infraestructura verde del estudio de preinversión del proyecto de Infraestructura Natural Piura (Código de Idea: 337830)», para la elaboración del Estudio de Preinversión del proyecto Infraestructura Natural Piura con CUI N° 2669244, en lo correspondiente a infraestructura verde."));
inf.push(para("De acuerdo con lo señalado en el numeral 6 de los TdR, la información presentada en el Segundo Entregable comprende un informe sobre las acciones y/o actividades de los puntos b), c) y d) descritos en el numeral 5.1 del ítem 5, adjuntando medios verificadores."));

// 4.1 b)
inf.push(h2("4.1 Actividad b) según Términos de Referencia"));
inf.push(literal(ACT.b));
inf.push(para(`Durante el periodo del segundo entregable se actualizó, integró y consolidó la información geoespacial y el diagnóstico territorial de los 107 bloques preliminares de intervención de Huancabamba (48 bloques, ${fmt(H.ha)} ha) y Morropón (59 bloques, ${fmt(M.ha)} ha). La línea base del primer entregable —dos volúmenes de diagnóstico territorial con corte al 21 de septiembre de 2026— evolucionó al Diagnóstico Territorial y Social, Versión 5, que lee en una sola estructura por bloque los estudios especializados del Capítulo de Identificación disponibles al corte del 17 de octubre de 2026. El desarrollo íntegro se presenta en el Anexo 1.`));
inf.push(h3("4.1.1 Marco metodológico de la actualización"));
inf.push(para("La actualización adopta el bloque del catálogo V6 como unidad común de integración y aplica tres reglas: (i) cada cifra procede de una fuente primaria identificada y fechada; (ii) ningún valor ausente se imputa; y (iii) las divergencias entre fuentes se registran con código y tratamiento, sin modificar la fuente. Las capas incorporadas en el periodo se resumen en el Cuadro N° 4."));
inf.push(N.cuadro("Insumos incorporados a la consolidación geoespacial y al diagnóstico territorial en el periodo"));
inf.push(tabla([{ w: 2700 }, { w: 4738 }, { w: 2200, align: C }], ["CAPA O ESTUDIO", "CONTENIDO INTEGRADO POR BLOQUE", "CORTE"], [
  ["Diagnóstico Territorial (E1)", "Informes DT y fichas F-DT-01 a F-DT-05 de los 107 bloques; MSAVI 2024, régimen territorial, peligro declarado", "20–21/09/2026"],
  ["Peligro Integrado (MCA–AHP)", "Tabulación zonal de la capa PI_Bloque en cinco niveles CENEPRED; índice ponderado por bloque", "Capa PI_Bloque"],
  ["E3 Geología", "Índice de Susceptibilidad Litológica ISL-MM e ISL-EH por bloque (Vol. I y II del estudio)", "E3 culminado"],
  ["E4 Hidroclima", "Series SENAMHI PISCOp v3.0 (1981–2025) y PISCOt v1.2 (1981–2020): P, ETP, índice de aridez, MFI", "23/09/2026"],
  ["E5 Suelos (campo)", "Calicatas, chequeos e infiltración del Informe de Avance de la Fase de Campo; CUM y matriz suelo–intervención preliminares", "20/09/2026"],
  ["Ecosistemas y pendientes V6", "Atlas de ecosistemas MINAM 2018 por bloque (117 láminas), matriz de pendientes DEM Copernicus GLO-30, área de influencia AI_aprobado_2", "05–06/10/2026"],
  ["E6 Estado de los ecosistemas", "Coberturas S2DR4 1 m, MSAVI por clases, índice de condición, clase de estado, delimitación interna y especies por bloque", "10/2026"],
  ["Situación predial v2 (Subentregable 2.2)", "117 mapas A3 SP-001 a SP-117: comunidades campesinas (GEORURAL), predios COFOPRI y tierras sin registro; centros poblados asociados", "11/10/2026"],
  ["E2 Diagnóstico social", "Consolidado DS del aplicativo IN Piura: fichas F-DS-01 a F-DS-03 vigentes, tenencia declarada, actores, servicios y economía", "17/10/2026"],
]));
inf.push(fuente("Fuente: Diagnóstico Territorial y Social, Versión 5, Volúmenes I (Morropón, revisión 10) y II (Huancabamba, revisión integrada DT + DS, actualización 17/10/2026)."));
inf.push(h3("4.1.2 Ámbito y cifras consolidadas"));
inf.push(para(`El ámbito consolidado se mantiene en 107 bloques y ${fmt(TOT_HA)} ha, en 14 distritos. El número de microcuencas se corrige a ${H.nmc + M.nmc} (${M.nmc} en Morropón y ${H.nmc} en Huancabamba) conforme al catálogo, frente a las 29 consignadas en el primer entregable. Las cifras consolidadas por provincia se presentan en el Cuadro N° 5.`));
inf.push(N.cuadro("Cifras consolidadas del diagnóstico territorial por provincia (Versión 5)"));
const pc = (a, bb) => fmt(100 * a / bb, 1) + " %";
const e6cls = (P, k) => `${P.e6_cls[k] || 0} (${fmt(P.e6_clsha[k] || 0)} ha)`;
inf.push(tabla([{ w: 3638 }, { w: 2000, align: R }, { w: 2000, align: R }, { w: 2000, align: R }], ["INDICADOR", "HUANCABAMBA", "MORROPÓN", "TOTAL"], [
  ["Bloques / superficie (ha)", `48 / ${fmt(H.ha)}`, `59 / ${fmt(M.ha)}`, `107 / ${fmt(TOT_HA)}`],
  ["Distritos / microcuencas", `${H.nd} / ${H.nmc}`, `${M.nd} / ${M.nmc}`, `14 / ${H.nmc + M.nmc}`],
  ["Peligro Integrado Alto + Muy alto (ha)", `${fmt(H.h45)} (${fmt(H.p45, 1)} %)`, `${fmt(M.h45)} (${fmt(M.p45, 1)} %)`, fmt(H.h45 + M.h45)],
  ["Brecha por media del bloque (MSAVI < 0.4976)", `${H.brecha_media_n} bloques · ${fmt(H.brecha_media_ha)} ha`, `${M.brecha_media_n} bloques · ${fmt(M.brecha_media_ha)} ha`, `${H.brecha_media_n + M.brecha_media_n} bloques · ${fmt(H.brecha_media_ha + M.brecha_media_ha)} ha`],
  ["Superficie bajo umbral, criterio areal del E6 (ha)", fmt(H.e6_bajo_ha), fmt(M.e6_bajo_ha), fmt(H.e6_bajo_ha + M.e6_bajo_ha)],
  ["Índice de condición medio ponderado (E6)", fmt(H.e6_ic_pond, 3), fmt(M.e6_ic_pond, 3), "—"],
  ["Bloques en estado Moderado o Severo (E6)", `${(H.e6_cls.Moderado || 0) + (H.e6_cls.Severo || 0)}`, `${(M.e6_cls.Moderado || 0) + (M.e6_cls.Severo || 0)}`, `${(H.e6_cls.Moderado || 0) + (H.e6_cls.Severo || 0) + (M.e6_cls.Moderado || 0) + (M.e6_cls.Severo || 0)}`],
  ["Superficie en comunidad campesina (ha)", `${fmt(H.pred_cc)} (${pc(H.pred_cc, H.ha)})`, `${fmt(M.pred_cc)} (${pc(M.pred_cc, M.ha)})`, fmt(H.pred_cc + M.pred_cc)],
  ["Predios rurales COFOPRI interceptados", `${H.pred_npred}`, `${M.pred_npred}`, `${H.pred_npred + M.pred_npred}`],
  ["Calicatas del E5 asociadas a bloques", `${H.cal_n} en ${H.cal_bloques} bloques`, `${M.cal_n} en ${M.cal_bloques} bloques`, `${H.cal_n + M.cal_n} en ${H.cal_bloques + M.cal_bloques} bloques`],
  ["Cárcavas codificadas en capa de gabinete", `${H.cv_n} (${fmt(H.cv_m, 0)} m)`, `${M.cv_n} (${fmt(M.cv_m, 0)} m)`, `${H.cv_n + M.cv_n} (${fmt(H.cv_m + M.cv_m, 0)} m)`],
  ["Fichas sociales vigentes / bloques con ficha", "93 / 24 de 48", "252 / 52 de 59", "345 / 76 de 107"],
]));
inf.push(fuente("Fuente: matrices de los Volúmenes I y II, Versión 5 (Anexo 1 y Anexo 4); índice de mapas de situación predial v2 (11/10/2026); consolidado DS del aplicativo IN Piura (17/10/2026). Superficies en UTM WGS 84 Zona 17S."));
inf.push(h3("4.1.3 Resultados obtenidos"));
inf.push(para([b("Estado del ecosistema por bloque. "), `La integración del E6 añade al diagnóstico la condición de cada bloque frente a ecosistemas de referencia del mismo piso. En el ámbito, ${H.e6_cls["No degradado"] + M.e6_cls["No degradado"]} bloques resultan no degradados, ${H.e6_cls.Leve + M.e6_cls.Leve} leves, ${(H.e6_cls.Moderado || 0) + (M.e6_cls.Moderado || 0)} moderados y ${M.e6_cls.Severo || 0} severos; los seis severos son los bloques 77 y 83 a 87 de San Juan de Bigote, lo que confirma el conglomerado crítico identificado en el primer entregable. La delimitación interna del E6 asigna ${fmt(H.e6_delim.delim_rest_activa_ha + M.e6_delim.delim_rest_activa_ha)} ha a restauración activa, ${fmt(H.e6_delim.delim_enriquecimiento_ha + M.e6_delim.delim_enriquecimiento_ha)} ha a enriquecimiento, ${fmt(H.e6_delim.delim_conservacion_ha + M.e6_delim.delim_conservacion_ha)} ha a conservación y ${fmt(H.e6_delim.delim_tamizaje_agricola_ha + M.e6_delim.delim_tamizaje_agricola_ha)} ha a tamizaje agrícola.`]));
inf.push(figura(FIG + "fig_e6.png", 15.5));
inf.push(N.figura("Clase de estado del ecosistema por provincia (porcentaje de la superficie de los bloques)"));
inf.push(fuente("Fuente: Tablero E6 «Estado de los ecosistemas por bloque de intervención» integrado en la Versión 5."));
inf.push(para([b("Tamizaje predial de gabinete. "), `El vacío más crítico del primer entregable —ningún bloque con tamizaje predial— se atendió en gabinete con los 117 mapas de situación predial v2 (Subentregable 2.2, 11/10/2026), de los cuales 107 corresponden al ámbito de la actividad b). En Huancabamba, ${fmt(H.pred_cc)} ha (${pc(H.pred_cc, H.ha)}) están en comunidad campesina, ${fmt(H.pred_cof)} ha en predios COFOPRI y ${fmt(H.pred_est)} ha sin registro; en Morropón, ${fmt(M.pred_cc)} ha (${pc(M.pred_cc, M.ha)}) en comunidad campesina y ${fmt(M.pred_est)} ha sin registro. Cinco bloques de Huancabamba superan el máximo de 19 predios del Paso 5: ${H.pred_gt19.map(([bq, n]) => `${bq} (${n})`).join(", ")}, y deben sectorizarse o redelimitarse antes de las actas. La tenencia declarada en la F-DS-01 difiere del gabinete en 3 bloques de Huancabamba y en 17 de Morropón; el tamizaje de campo, a cargo del equipo de liberación de áreas que ingresa en la segunda quincena de octubre, debe resolverlo.`]));
inf.push(figura(FIG + "fig_predial.png", 15.5));
inf.push(N.figura("Situación predial de gabinete de los bloques por provincia"));
inf.push(fuente("Fuente: Índice de mapas de situación predial de los 117 bloques V6, v2 (GEORURAL–MIDAGRI y COFOPRI), 11/10/2026."));
inf.push(para([b("Suelos, social y riesgo. "), `La información de campo del E5 quedó asociada a cada bloque: ${H.cal_n + M.cal_n} calicatas en ${H.cal_bloques + M.cal_bloques} bloques, con su clasificación tentativa, CUM de campo y matriz suelo–intervención preliminar. El diagnóstico social integrado registra 345 fichas vigentes en 76 de los 107 bloques; continúan sin ficha 24 bloques de Huancabamba —entre ellos los cinco del núcleo de brecha M4B4, M12B1, M4B3, M20B1 y M2B8— y 7 de Morropón. La síntesis peligro × exposición califica ${H.riesgo.Alto + M.riesgo.Alto} bloques con riesgo Alto, ${H.riesgo.Medio + M.riesgo.Medio} Medio y ${H.riesgo.Bajo + M.riesgo.Bajo} Bajo, como valoración preliminar de gabinete.`]));
inf.push(para([b("Criterio de brecha. "), `La consolidación evidencia dos criterios de cálculo del indicador: la media del bloque frente al umbral (${H.brecha_media_n + M.brecha_media_n} bloques, ${fmt(H.brecha_media_ha + M.brecha_media_ha)} ha) y la distribución areal del E6 (${fmt(H.e6_bajo_ha + M.e6_bajo_ha)} ha). El indicador de la R.M. N° 00213-2024-MINAM no es sumable entre provincias mientras el criterio difiera; se recomienda adoptar un criterio único para los 117 bloques (numeral 6).`]));
inf.push(para("Los cambios respecto de la línea base del primer entregable se sintetizan en el Cuadro N° 6."));
inf.push(N.cuadro("Cambios en la información geoespacial y el diagnóstico territorial respecto del primer entregable"));
inf.push(tabla([{ w: 2300 }, { w: 3200 }, { w: 4138 }], ["COMPONENTE", "PRIMER ENTREGABLE (corte 21/09/2026)", "SEGUNDO ENTREGABLE (corte 17/10/2026)"], [
  ["Versión del diagnóstico", "Volúmenes I y II de Diagnóstico Territorial", "Diagnóstico Territorial y Social, Versión 5 (Vol. I revisión 10; Vol. II revisión integrada DT + DS), publicado en línea"],
  ["Estudios integrados por bloque", "DT, MSAVI 2024, PI", "DT, MSAVI, PI, E3 (ISL), E4 (hidroclima), E5 (suelos de campo), E6 (estado del ecosistema), E2 (social) y situación predial v2"],
  ["Microcuencas", "29", `${H.nmc + M.nmc} (corrección conforme al catálogo)`],
  ["Estado del ecosistema", "No disponible", `IC por bloque: ${H.e6_cls["No degradado"] + M.e6_cls["No degradado"]} no degradados, ${H.e6_cls.Leve + M.e6_cls.Leve} leves, 14 moderados, 6 severos`],
  ["Delimitación interna", "Unidades de manejo propuestas en el bloque 79", "Cuatro clases cuantificadas en los 107 bloques (restauración activa, enriquecimiento, conservación, tamizaje agrícola)"],
  ["Tamizaje predial", "0 de 107 bloques", "107 de 107 con situación predial de gabinete; 5 bloques > 19 predios; campo y actas pendientes"],
  ["Suelos", "Plan de muestreo aprobado", `${H.cal_n + M.cal_n} calicatas de campo asociadas a ${H.cal_bloques + M.cal_bloques} bloques; laboratorio pendiente`],
  ["Diagnóstico social", "No integrado", "345 fichas F-DS vigentes en 76 bloques; tenencia declarada contrastada con el gabinete"],
  ["Brecha", "50 bloques, 7,637 ha (media del bloque)", `Se mantiene y se contrasta con el criterio areal del E6 (${fmt(H.e6_bajo_ha + M.e6_bajo_ha)} ha)`],
  ["Filtros formales de exclusión", "No verificados", "Sin cambio: pendiente la intersección con SIGDA, INGEMMET, OSINERGMIN y SERNANP"],
]));
inf.push(fuente("Fuente: Informe N° 007-2026-HSCM (Primer Entregable) y Diagnóstico Territorial y Social, Versión 5."));
inf.push(para("El desarrollo íntegro de esta actividad —con la consolidación por distrito y por bloque, el detalle del tamizaje predial, la integración de los estudios, los vacíos de integración y los medios verificadores— se presenta en el Anexo 1, y las matrices nativas en el Anexo 4."));

// 4.2 c)
inf.push(h2("4.2 Actividad c) según Términos de Referencia"));
inf.push(literal(ACT.c));
inf.push(para("Durante el periodo del segundo entregable se ejerció el seguimiento técnico del Estudio de Suelos a Nivel de Reconocimiento en coordinación con el Especialista en Infraestructura Marrón, Ing. Pedro Talledo Hernández: se supervisó el cierre de la fase de campo, se revisó el Informe de Avance de la Fase de Campo (Entregable N° 2), se formularon las observaciones de consistencia, se dio seguimiento al envío de muestras al laboratorio (Entregable N° 3) y se verificó la aplicación de los resultados de campo al componente verde mediante su integración por bloque en el diagnóstico territorial. El desarrollo íntegro se presenta en el Anexo 2."));
inf.push(h3("4.2.1 Seguimiento de entregables y estado de conformidades"));
inf.push(N.cuadro("Seguimiento de los entregables del Estudio de Suelos a Nivel de Reconocimiento"));
inf.push(tabla([{ w: 1300, align: C }, { w: 2400 }, { w: 1300, align: C }, { w: 2838 }, { w: 1800, align: C }], ["ENTREGABLE", "CONTENIDO", "PLAZO", "ESTADO TÉCNICO", "CONDICIÓN"], [
  ["N° 1", "Plan de Trabajo y Plan de Muestreo georreferenciado", "Día 10 (27/08/2026)", "Revisado; observaciones metodológicas levantadas (primer entregable)", "CONFORMIDAD OTORGADA (04/09/2026)"],
  ["N° 2", "Informe de Avance – Fase de Campo: fichas, infiltración, registro fotográfico, GDB de puntos", "Día 40 (26/09/2026)", "Recibido y revisado al 100 % (reporte del 03/10/2026); 14 observaciones de consistencia formuladas (Anexo 2)", "REVISADO – conformidad sujeta al levantamiento de observaciones"],
  ["N° 3", "Resultados de laboratorio e interpretación preliminar", "Día 55 (11/10/2026)", "Muestras enviadas e ingresadas al laboratorio; a la espera de resultados", "EN PROCESO"],
  ["N° 4", "Informe final: CUM, Factor K, matriz suelo–intervención, estudio agronómico SUS y GDB", "Día 75 (31/10/2026)", "Pendiente; el Plan de Trabajo Detallado reprograma su recepción al 13/11/2026", "PENDIENTE"],
]));
inf.push(fuente("Fuente: Informe de Avance – Fase de Campo (GeoSIG Ingenieros, setiembre de 2026), Cuadros 1 y 22; Plan de Trabajo Detallado del Perfil, Actualización N° 02 (16/10/2026); tablero de avance del Perfil al 03/10/2026."));
inf.push(h3("4.2.2 Verificación de las actividades de campo"));
inf.push(para("La fase de campo se ejecutó entre el 03 y el 20 de septiembre de 2026 (días 17 a 34 del servicio) con tres brigadas, en lugar de las dos previstas, para recuperar el inicio diferido hasta la conformidad del Entregable N° 1. Se verificó el cumplimiento de la intensidad de muestreo exigida por el TdR del estudio, que se resume en el Cuadro N° 8."));
inf.push(N.cuadro("Cumplimiento de la intensidad de muestreo del Estudio de Suelos"));
inf.push(tabla([{ w: 2400 }, { w: 2200, align: C }, { w: 2000, align: C }, { w: 3038, align: C }], ["EXIGENCIA", "MÍNIMO DEL TdR", "PLAN DE MUESTREO", "EJECUTADO AL 20/09/2026"], [
  ["Calicatas", "1 por cada 200 ha (62)", "74 (1 por cada 166 ha)", "74 (100 %): 35 de bloque, 34 de referencia y 5 en Lotes SUS"],
  ["Barrenaciones de comprobación", "2 por calicata (148)", "148", "148 (100 %)"],
  ["Pruebas de infiltración", "1 por suelo dominante, 5 repeticiones", "12 sitios", "13 sitios (12 del plan, uno reubicado, más 1 adicional); 59 repeticiones medidas"],
  ["Muestreo de horizontes", "2 por calicata (148)", "148 como mínimo", "209 horizontes registrados en las fichas"],
  ["Georreferenciación", "UTM WGS 84 17S con altitud", "Sí", "74 calicatas, 148 chequeos y 13 pruebas en la GDB ENTREGABLE2_PUNTOS_MUESTREADOS"],
]));
inf.push(fuente("Fuente: Informe de Avance – Fase de Campo, Cuadros 9 y 10, verificados por el suscrito. La diferencia entre las 65 repeticiones del resumen ejecutivo y las 59 que suman los cuadros es objeto de observación."));
inf.push(para("De las 74 calicatas, 17 se reubicaron a más de 300 m del punto de gabinete por difícil acceso, siempre dentro de su unidad de suelo y zona homogénea; el mayor desplazamiento corresponde al par CAL-65/CAL-66 del páramo, trasladado al sector Gaspar dentro del mismo bloque M22B1. En el sector Hualcas–Chignia (Huarmaca), con actividad minera y rondas campesinas, el ingreso se coordinó previamente con autoridades locales y rondas sin necesidad de reubicar puntos, en concordancia con lo constatado en el diagnóstico territorial sobre el rol de las rondas como instancia que regula el acceso al territorio."));
inf.push(h3("4.2.3 Revisión técnica del Entregable N° 2 y formulación de observaciones"));
inf.push(para("La revisión técnica contrastó el informe contra el TdR del estudio, el Plan de Muestreo aprobado y el catálogo V6. El Entregable N° 2 cumple el alcance de la Fase 2 en calicatas, barrenaciones, descripción morfológica, muestreo, infiltración, georreferenciación y registro fotográfico; las observaciones formuladas son de consistencia y de alcance para los Entregables N° 3 y N° 4, y se sintetizan en el Cuadro N° 9 (matriz completa en el Anexo 2)."));
inf.push(N.cuadro("Síntesis de observaciones formuladas al Entregable N° 2 del Estudio de Suelos"));
inf.push(tabla([{ w: 1200, align: C }, { w: 5738 }, { w: 2700 }], ["CÓDIGO", "OBSERVACIÓN", "LEVANTAMIENTO"], [
  ["OS-01", "El informe consigna 118 bloques y 12,299 ha; el catálogo vigente V6 tiene 117 bloques y 12,270.235 ha.", "Homologar el universo en el Entregable N° 4"],
  ["OS-02", "El resumen ejecutivo reporta 65 repeticiones de infiltración; los Cuadros 8, 9 y 15.1 suman 59.", "Corregir en el Entregable N° 4"],
  ["OS-03", "El resumen ejecutivo reporta 152 horizontes muestreados; los Cuadros 9 y 10 y el numeral 8.3 reportan 209.", "Precisar muestras remitidas al laboratorio"],
  ["OS-04 / OS-05", "INF-11 figura como «Culcas Alto» (Cuadro 8) y «La Unión» (Cuadro 19), con 2 repeticiones; INF-08 tiene 4 e INF-02b, 3, frente a las 5 del TdR.", "Uniformizar; declarar valores indicativos, no usar para diseño sin repetición"],
  ["OS-07", "El laboratorio figura como UNALM (E2), LASPAF-UNALM (Plan de Trabajo) y «Valle Grande» (Plan de Trabajo Detallado del Perfil).", "Acreditar laboratorio, acreditación y fecha de envío"],
  ["OS-10", "La CUM de campo difiere de la CUM de gabinete en varias unidades (p. ej., CAL-17/18: P3 y X frente a C2se(r)).", "Explicar y resolver con laboratorio y D.S. N° 005-2022-MIDAGRI"],
  ["OS-11", "La matriz suelo–intervención es preliminar por grupo de suelos, no por unidad de suelo ni por bloque.", "Entregar por unidad de suelo y bloque en el Entregable N° 4"],
]));
inf.push(fuente("Fuente: revisión técnica del suscrito sobre el Informe de Avance – Fase de Campo. Matriz completa OS-01 a OS-14 en el Anexo 2."));
inf.push(h3("4.2.4 Verificación de la aplicación técnica en el componente verde"));
inf.push(para("Los resultados de campo ya condicionan el diseño del componente verde. Nueve de los trece sitios de infiltración tienen velocidad básica lenta a moderadamente lenta (0.33 a 1.37 cm/h): en esos suelos las zanjas de infiltración requieren mayor capacidad de almacenamiento y menor espaciamiento, y donde el contacto restrictivo está a menos de 50 cm deben priorizarse medidas de cobertura y retención superficial. La pre-clasificación de CUM ubica 34 calicatas en protección, 19 en forestal de calidad baja con alternativa de protección, 8 en forestal media o pastoreo, 8 en pastoreo, 4 en cultivo permanente y 1 en cultivo en limpio, coherente con la intervención en laderas degradadas. En coordinación con el Especialista en Infraestructura Marrón, la regla preliminar —zanjas en pendientes de 15 a 50 % con 50 cm o más de suelo; terrazas de formación lenta de 15 a 40 %; revegetación, exclusión de pastoreo y control de cárcavas sobre 50 %— se incorporó a la ficha de cada bloque del diagnóstico."));
inf.push(figura(FIG + "fig_infiltracion.png", 15.5));
inf.push(N.figura("Velocidad de infiltración básica representativa por sitio"));
inf.push(fuente("Fuente: Informe de Avance – Fase de Campo, Cuadro 19. INF-11 (2 repeticiones), INF-08 (4) e INF-02b (3) son indicativos (observación OS-04)."));

// 4.3 d)
inf.push(h2("4.3 Actividad d) según Términos de Referencia"));
inf.push(literal(ACT.d));
inf.push(para("Ambos estudios de la actividad d) constituyen un solo producto del Perfil: el Entregable 6 (E6) «Diagnóstico del estado del ecosistema y estudio de oferta de material vegetal», regido por el Índice de Contenidos Mínimos 999990-ANIN001-000-XX-SC-EN-000002-R00. Declarado desierto el procedimiento de contratación y ante la alerta de El Niño Costero, el seguimiento se ejerció, en coordinación con el Especialista en Intervenciones Vegetativas de Uso Sostenible, Blgo. Abel Aucasime Orihuela, sobre la reformulación del plan de trabajo, los insumos ya producidos por órdenes de servicio y los productos del E6 por Administración Directa. El desarrollo íntegro se presenta en el Anexo 3."));
inf.push(h3("4.3.1 Revisión y actualización de los planes de trabajo"));
inf.push(para("El plan del E6 se actualizó de la versión 1.1 (servicio externo de 16 semanas, S/ 526,900 referenciales, 157 parcelas de campo de 0.1 ha) a la versión 2.0 «In house» (08/10/2026): 10 semanas con equipo de ANIN, inspección remota con parcelas virtuales de 0.5 ha sobre imágenes de 0.6 a 1 m, ≈ 600 puntos de validación y meta de exactitud global ≥ 85 %, e incorporación del estudio de oferta de material vegetal que la versión 1.1 no contemplaba. La revisión del suscrito («rev_hscm») ajustó la denominación de la modalidad, asignó al especialista SIG y operador RPAS, precisó el perfil del especialista en viveros, corrigió el horizonte de registro del Perfil y amplió las solicitudes de información a los inventarios florísticos de las ACR. El Plan de Trabajo Detallado del Perfil (Actualización N° 01 del 11/10/2026 y N° 02 del 16/10/2026) incorporó el cambio de modalidad y fijó la redacción final del E6 al 18/11/2026."));
inf.push(h3("4.3.2 Seguimiento de actividades y revisión técnica de entregables"));
inf.push(para("Sin campaña nueva de campo, el seguimiento de las actividades de campo se aplicó a la evidencia ya levantada (fichas F-DT de los 117 bloques, calicatas del E5, fichas F-DS y vuelos RPAS) y a su uso como control de la inspección remota. Se revisaron técnicamente los insumos producidos por órdenes de servicio y los productos del E6, como se resume en el Cuadro N° 10."));
inf.push(N.cuadro("Revisión técnica de los productos del Diagnóstico del Estado de los Ecosistemas y del Estudio de Oferta de Material Vegetal"));
inf.push(tabla([{ w: 2600 }, { w: 1400, align: C }, { w: 3438 }, { w: 2200 }], ["PRODUCTO", "FECHA", "RESULTADO DE LA REVISIÓN", "ESTADO"], [
  ["Entregables de las OS 1148 y 2797 (flora y ecosistemas) y OS 1147, 2357 y 2726 (oferta de material vegetal, sistematización de bloques y fitosanidad)", "03/10/2026", "Listas florísticas con errores taxonómicos y sin depurar (~77 nativas y ~90 taxones); demanda de plantones solo para Lotes SUS en 83 bloques; ningún vivero identificado en los 15 distritos", "Observaciones atendidas parcialmente en el E6"],
  ["Plan del E6, versión 2.0 «In house» y su revisión", "08/10/2026", "Adecuación de los 14 capítulos del índice a inspección remota; cinco ítems requieren conformidad expresa de SESDI", "Revisado; conformidad de SESDI pendiente"],
  ["Nota Técnica N° 002-2026-ANIN/DIME/SESDI (firma espectral)", "14/10/2026", "Clasificación de 92 tomas S2DR4 de 1 m en 9 coberturas con un único modelo", "Remitida a la Coordinación IN; conformidad metodológica pendiente (hito H7b)"],
  ["Informe E6 y matrices (versión para revisión de SESDI)", "15/10/2026", "Brecha, índice de condición, delimitación interna, 69 especies, 16 fichas técnicas y oferta identificada", "En revisión; observaciones de consistencia formuladas (Anexo 3)"],
]));
inf.push(fuente("Fuente: tablero de avance del Perfil al 03/10/2026; E6 v2.0 y su revisión; Nota Técnica N° 002-2026; Informe E6 (octubre de 2026)."));
inf.push(h3("4.3.3 Resultados principales y estado de conformidades"));
inf.push(para("El E6 diagnosticó los 117 bloques (12,270.235 ha) con la firma espectral de 92 tomas S2DR4 de 1 m, el MSAVI 2024 y el NDVI, contrastados con la evidencia de campo 2026. Sus resultados principales se presentan en el Cuadro N° 11."));
inf.push(N.cuadro("Resultados principales del Entregable 6 en el ámbito del proyecto"));
inf.push(tabla([{ w: 4200 }, { w: 5438 }], ["INDICADOR", "RESULTADO"], [
  ["Superficie bajo el umbral MSAVI 0.4976 (brecha)", "8,341.82 ha (68.0 % de los bloques), de ellas 5,503.41 ha bajo 0.3813 (restauración activa)"],
  ["Concentración", "Morropón: 6,216.63 ha bajo umbral (84.8 % de su superficie)"],
  ["Clase de estado (índice de condición)", "53 no degradados (3,949.76 ha), 44 leves (6,675.47 ha), 14 moderados (1,493.62 ha) y 6 severos (151.39 ha)"],
  ["Lectura integrada MSAVI–NDVI", "36 bloques (6,298.90 ha, 51.3 %) con degradación convergente"],
  ["Inspección remota", "1,374 parcelas virtuales en bloques y 102 de referencia; todos los estratos cumplen el error de 15 % salvo el matorral andino bajo umbral"],
  ["Contraste con campo", "r = 0.72 frente a la cobertura de 82 fichas F-DT; sesgo +15.9 pp"],
  ["Flora", "Matriz de 69 especies nativas; 48 taxones confirmados en campo 2026; 16 fichas técnicas de especies recomendadas"],
  ["Oferta de material vegetal", "12 viveros o iniciativas identificadas; solo los viveros del GORE Piura en Hualtacal declaran capacidad (21,000 plantones)"],
]));
inf.push(fuente("Fuente: Informe E6 – Diagnóstico del estado del ecosistema y estudio de oferta de material vegetal, versión para revisión de SESDI (octubre de 2026)."));
inf.push(para("En cuanto a conformidades, el E6 se encuentra en versión para revisión de SESDI: la conformidad de la adecuación metodológica (memorando de SESDI e hito H7b) y la del informe final están pendientes. El componente de oferta de material vegetal es el de menor avance —censo remoto de viveros y cálculo de la brecha con los parámetros del E10 pendientes— y la plaza de Especialista en Infraestructura Rural para la Producción (viveros) quedó desierta, por lo que sus funciones deben reasignarse formalmente entre los especialistas de Intervenciones Vegetativas de Uso Sostenible, Infraestructura Verde e Infraestructura Marrón."));

// 5) CONCLUSIONES
inf.push(h1("5) CONCLUSIONES"));
let k = 0;
const CN = (t) => inf.push(num(`5.${++k}.`, t));
inf.push(para([b("Para la actividad b)")])); inf.push(literal(ACT.b));
CN(`Se actualizó, integró y consolidó la información geoespacial y el diagnóstico territorial de los 107 bloques preliminares de intervención de Huancabamba y Morropón (${fmt(TOT_HA)} ha, 14 distritos y ${H.nmc + M.nmc} microcuencas) en el Diagnóstico Territorial y Social, Versión 5, que integra por bloque nueve fuentes: DT, PI, E3, E4, E5, ecosistemas y pendientes V6, E6, situación predial v2 y E2.`);
CN(`El Peligro Integrado Alto y Muy alto cubre ${fmt(H.h45 + M.h45)} ha del ámbito consolidado (${fmt(H.p45, 1)} % en Huancabamba y ${fmt(M.p45, 1)} % en Morropón), lo que confirma la pertinencia de la intervención en la zona de origen de los peligros.`);
CN(`El estado del ecosistema se cuantificó por bloque: 14 bloques moderados y 6 severos; los seis severos (77 y 83 a 87 de San Juan de Bigote) ratifican el conglomerado crítico del primer entregable, y la delimitación interna asigna ${fmt(H.e6_delim.delim_rest_activa_ha + M.e6_delim.delim_rest_activa_ha)} ha a restauración activa y ${fmt(H.e6_delim.delim_enriquecimiento_ha + M.e6_delim.delim_enriquecimiento_ha)} ha a enriquecimiento.`);
CN(`Los 107 bloques cuentan con situación predial de gabinete: ${fmt(H.pred_cc + M.pred_cc)} ha en comunidades campesinas, ${fmt(H.pred_cof + M.pred_cof)} ha en predios COFOPRI y ${fmt(H.pred_est + M.pred_est)} ha sin registro; cinco bloques de Huancabamba (4, 38, 9, 35 y 43) exceden el máximo de 19 predios del Paso 5 y no hay actas suscritas al corte.`);
CN("Persisten dos criterios de cálculo de la brecha (media del bloque y distribución areal del E6), por lo que el indicador de la R.M. N° 00213-2024-MINAM no es aún sumable entre provincias.");
inf.push(para([b("Para la actividad c)")])); inf.push(literal(ACT.c));
CN("Se dio seguimiento técnico al Estudio de Suelos en coordinación con el Especialista en Infraestructura Marrón: el Entregable N° 1 mantiene su conformidad (04/09/2026), el Entregable N° 2 fue recibido y revisado al 100 %, el Entregable N° 3 se encuentra en laboratorio y el Entregable N° 4 está pendiente.");
CN("La fase de campo cumplió la intensidad del TdR: 74 calicatas (una por cada 166 ha), 148 barrenaciones, 13 sitios de infiltración con 59 repeticiones medidas y 209 horizontes registrados, con 17 calicatas reubicadas dentro de su unidad de suelo.");
CN("Se formularon catorce observaciones de consistencia y alcance al Entregable N° 2, ninguna de las cuales invalida la fase de campo; su levantamiento se exige en los Entregables N° 3 y N° 4.");
CN("La información edáfica ya se aplica al componente verde: las velocidades de infiltración lentas condicionan el dimensionamiento de zanjas, la CUM de campo ubica la mayor parte de las tierras en protección y forestal, y la regla preliminar suelo–intervención está integrada en la ficha de cada bloque.");
inf.push(para([b("Para la actividad d)")])); inf.push(literal(ACT.d));
CN("El Diagnóstico del Estado de los Ecosistemas y el Estudio de Oferta de Material Vegetal se ejecutan como E6 por Administración Directa; su plan de trabajo fue revisado y actualizado a la versión 2.0 «In house» y el Plan de Trabajo Detallado del Perfil fija su redacción final al 18/11/2026.");
CN("Se revisaron técnicamente los insumos de las OS 1147, 1148, 2357, 2726 y 2797, la Nota Técnica N° 002-2026 y el Informe E6 en versión para revisión de SESDI; las observaciones formuladas se registran en el Anexo 3.");
CN("El E6 cuantifica 8,341.82 ha (68.0 %) bajo el umbral MSAVI en los 117 bloques, con 5,503.41 ha para restauración activa; la oferta de material vegetal identificada (21,000 plantones) es mínima frente a la escala de la intervención, y la conformidad de SESDI del E6 está pendiente.");

// 6) RECOMENDACIONES
inf.push(h1("6) RECOMENDACIONES"));
inf.push(para([b("A. Recomendaciones de competencia del equipo formulador (ANIN – DIME – SESDI)")]));
k = 0;
const RC = (t) => inf.push(num(`6.${++k}.`, t));
inf.push(para([b("Para la actividad b)")]));
RC("Adoptar un criterio único de brecha para los 117 bloques —preferentemente la distribución areal del E6 bajo el umbral MSAVI 0.4976— y recalcular con él el indicador consolidado del Perfil.");
RC("Sectorizar o redelimitar los bloques 4, 38, 9, 35 y 43 antes de iniciar las actas, y resolver en el tamizaje de campo las discrepancias de tenencia entre la F-DS-01 y el gabinete (3 bloques en Huancabamba y 17 en Morropón).");
RC("Completar la ficha social en los 31 bloques sin ficha, con prioridad en el núcleo de brecha de Huancabamba (M4B4, M12B1, M4B3, M20B1 y M2B8), y aplicar la ficha F-DS-06 de percepción del riesgo como insumo del E8.");
RC("Ejecutar la intersección formal de los 107 polígonos con SIGDA, INGEMMET, OSINERGMIN y SERNANP, y levantar en campo el inventario de cárcavas donde la capa de gabinete no registra datos.");
inf.push(para([b("Para la actividad c)")]));
RC("Exigir el levantamiento de las observaciones OS-01 a OS-14 en los Entregables N° 3 y N° 4, y condicionar la conformidad del Entregable N° 4 a la homologación del universo de bloques con el catálogo V6.");
RC("Formalizar la ampliación de plazo o reprogramación del Entregable N° 4 si se confirma la recepción al 13/11/2026, y acreditar el laboratorio que analiza las muestras y su acreditación.");
RC("Requerir que la Matriz de Compatibilidad Suelo–Intervención se entregue por unidad de suelo y por bloque, con la máscara de profundidad efectiva menor de 50 cm, y que sus especies se concilien con la lista priorizada del E6.");
inf.push(para([b("Para la actividad d)")]));
RC("Emitir el memorando de SESDI que acepte la adecuación metodológica del E6 (inspección remota) y los ítems adaptados del índice de contenidos mínimos.");
RC("Ejecutar la validación ciega de 567 puntos, densificar el estrato de matorral andino bajo umbral y completar la consulta Delphi de atributos de especies antes de la conformidad del E6.");
RC("Reasignar formalmente las funciones de la plaza de viveros y ejecutar el censo remoto de viveros y fuentes semilleras, para entregar al E10 la oferta efectiva y la brecha de plantones por escenarios.");
RC("Homologar las fechas del E6 entre su plan v2.0 (18/12/2026) y el Plan de Trabajo Detallado (18/11/2026), y la referencia al catálogo V6 en la Nota Técnica N° 002-2026.");

// 7) ANEXOS
inf.push(h1("7) ANEXOS"));
[
  "Anexo 1 — Actualización, integración y consolidación de la información geoespacial y del diagnóstico territorial de las provincias de Huancabamba y Morropón (actividad b). Comprende el informe de actualización a la Versión 5, la consolidación por provincia, distrito y bloque, la situación predial de gabinete, los vacíos de integración y los medios verificadores.",
  "Anexo 2 — Seguimiento técnico del Estudio de Suelos a Nivel de Reconocimiento, en coordinación con el Especialista en Infraestructura Marrón (actividad c). Comprende el estado de los entregables, el seguimiento de campo, la revisión técnica del Entregable N° 2 con su matriz de observaciones, el seguimiento del laboratorio y la verificación de su aplicación al componente verde.",
  "Anexo 3 — Seguimiento del Diagnóstico del Estado de los Ecosistemas (3-A) y del Estudio de Oferta de Material Vegetal de Propagación (3-B), en coordinación con el Especialista en Intervenciones de Uso Sostenible (actividad d). Comprende la revisión y actualización de los planes de trabajo, la revisión técnica de los entregables, las observaciones y el estado de las conformidades.",
  "Anexo 4 — Matrices de seguimiento en formato Excel (información nativa): consolidado por bloque y por distrito, situación predial, observaciones del Estudio de Suelos y del E6, y registro de medios verificadores.",
  "La información nativa restante correspondiente al procesamiento, análisis y resultados se remitirá al enlace otorgado por la coordinación del proyecto de IN, conforme al numeral 6 de los TdR.",
].forEach((t) => inf.push(guion(t)));
inf.push(...L.firma(FECHA));

L.guardar([L.seccion(carta, { pie: PIE }), L.seccion(inf, { pie: PIE })], OUT, "Informe N° 008-2026-HSCM – Segundo Entregable");
