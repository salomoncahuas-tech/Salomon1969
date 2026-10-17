// Anexo N° 3 — Actividad d): Diagnóstico del Estado de los Ecosistemas (3-A) y Estudio de Oferta de Material Vegetal (3-B)
const L = require("./lib");
const { para, b, i, run, h1, h2, h3, guion, bullet, num, fuente, tabla, vacio, centro, W, C, R, fmt, figura } = L;
const { seccionCaratula } = require("./caratula");
const DEF = require("./anexos_def");
const OUT = process.argv[2];
const FIG = "/tmp/claude-0/-home-user-Salomon1969/03b82ed8-c0e9-54ec-93c9-bff5f93b575b/scratchpad/build/fig/";
const PIE = "Anexo N° 3 · Informe N° 008-2026-HSCM · Segundo Entregable · OS N° 4062-2026";
const N = L.numerador("A3.");

// ---------------- Parte común ----------------
const x = [];
x.push(centro([b("ANEXO N° 3", { size: 26 })], { spacing: { after: 60 } }));
x.push(centro([b("SEGUIMIENTO DEL DIAGNÓSTICO DEL ESTADO DE LOS ECOSISTEMAS Y DEL ESTUDIO DE OFERTA DE MATERIAL VEGETAL DE PROPAGACIÓN", { size: 22 })], { spacing: { after: 60 } }));
x.push(centro([i("Entregable 6 del Perfil (E6) — periodo del 26/09 al 27/10/2026", { size: 19 })], { spacing: { after: 240 } }));
x.push(h1("1. OBJETIVO Y ALCANCE"));
x.push(para("El presente anexo desarrolla la actividad d) del numeral 5.1 de los Términos de Referencia durante el periodo del segundo entregable. Documenta el seguimiento, en coordinación con el Especialista en Intervenciones Vegetativas de Uso Sostenible, Blgo. Abel Aucasime Orihuela, del Diagnóstico del Estado de los Ecosistemas (parte 3-A) y del Estudio de Oferta de Material Vegetal de Propagación (parte 3-B): la revisión y actualización de sus planes de trabajo, el seguimiento de las actividades de campo, la revisión técnica de los entregables, la formulación y el levantamiento de observaciones y el estado de las conformidades."));
x.push(para("Ambos estudios forman un solo producto del Perfil, el Entregable 6 «Diagnóstico del estado del ecosistema y estudio de oferta de material vegetal», regido por el Índice de Contenidos Mínimos 999990-ANIN001-000-XX-SC-EN-000002-R00 (SESDI, mayo de 2025), con tres componentes: estado del ecosistema, flora nativa para la gestión del riesgo de desastres (GRD) y oferta de material vegetal."));

x.push(h1("2. CONTEXTO: CAMBIO DE MODALIDAD DEL ESTUDIO"));
x.push(para("Tres hechos del periodo redefinieron el seguimiento de la actividad d):"));
x.push(bullet([b("Procedimiento desierto. "), "El procedimiento CP-ABR-5-2026-ANIN/DAI-1 para el Diagnóstico del Ecosistema y Estudio de Oferta de Material Vegetal fue declarado desierto en septiembre de 2026; la Coordinación de Proyectos IN desistió de reconvocarlo y dispuso su ejecución por Administración Directa, con el equipo del proyecto y a nivel de perfil."]));
x.push(bullet([b("Alerta de El Niño Costero. "), "Los Comunicados Oficiales ENFEN N° 16-2026 (14/09/2026: 48 % de probabilidad de magnitud extraordinaria y 45 % de fuerte para diciembre 2026–marzo 2027) y N° 17-2026 (28/09/2026: lluvias sobre lo normal en la costa norte desde octubre) hicieron inviable la campaña de 157 parcelas de campo, que caería en plena temporada de lluvias, cortes de vías y huaicos."]));
x.push(bullet([b("Plaza de viveros desierta. "), "La plaza de Especialista en Infraestructura Rural para la Producción (viveros) quedó desierta (octubre de 2026); la oferta y brecha de viveros del E6, el E10.2 y el E12.2 requieren la reasignación formal de funciones entre los especialistas de Intervenciones Vegetativas de Uso Sostenible, Infraestructura Verde e Infraestructura Marrón."]));
x.push(para("En consecuencia, el seguimiento se ejerció sobre (i) la reformulación del plan de trabajo, (ii) los insumos producidos por órdenes de servicio (OS 1147, 1148, 2357, 2726 y 2797), (iii) la Nota Técnica de resultados por firma espectral y (iv) el Informe E6 en versión para revisión de SESDI. El equipo del E6 lo integran el suscrito (responsable técnico), el Ing. Albert Flores Mendoza (teledetección y SIG; operador RPAS para el Plan B) y el Ing. Pedro Talledo Hernández (Especialista en Infraestructura Marrón, análisis), con el apoyo del Especialista en Intervenciones Vegetativas de Uso Sostenible para la flora de los Lotes SUS y la oferta de plantones."));

x.push(h1("3. REVISIÓN Y ACTUALIZACIÓN DE LOS PLANES DE TRABAJO"));
x.push(h2("3.1 Del plan versión 1.1 a la versión 2.0 «In house»"));
x.push(para(`El plan del E6 se actualizó a la versión 2.0 «In house» (08/10/2026), que ejecuta el estudio en 10 semanas, íntegramente en gabinete, y atiende los 14 capítulos del índice de contenidos mínimos. El Cuadro N° ${N.sigC()} compara ambas versiones.`));
x.push(N.cuadro("Cambios del plan de trabajo del E6: versión 1.1 frente a versión 2.0"));
x.push(tabla([{ w: 2400 }, { w: 3300 }, { w: 3938 }], ["COMPONENTE", "VERSIÓN 1.1", "VERSIÓN 2.0 «IN HOUSE»"], [
  ["Estructura", "Índice propio de 13 ítems", "Índice de Contenidos Mínimos R00, 14 capítulos"],
  ["Ejecución y plazo", "Servicio externo, 16 semanas; S/ 526,900 referenciales", "Equipo ANIN, 10 semanas (S1 = 12/10/2026, referencial)"],
  ["Unidades de muestreo", "157 parcelas de campo de 0.1 ha", "≈ 1,400–1,700 parcelas virtuales de 0.5 ha y ≈ 600 puntos de validación"],
  ["Firma espectral", "Campaña con espectrorradiómetro 350–2,500 nm", "Firmas espectro-temporales extraídas de imagen; librería UNF como referencia; espectrorradiometría diferida al estudio definitivo"],
  ["Dron", "≈ 30 vuelos nuevos", "Ortofotos ya tomadas como control; vuelos nuevos solo en el Plan B"],
  ["Composición florística", "Medida en parcelas", "Registros secundarios, especies declaradas en campo y validación de expertos (Delphi)"],
  ["Estructura vertical", "DAP y altura en parcelas", "Modelos de altura de dosel (GEDI, ETH 10 m, Meta 1 m) y fotointerpretación"],
  ["Ecosistemas de referencia", "Percentil 90 de referencias conservadas", "Identificados en imágenes de muy alta resolución y series 1985–2026"],
  ["Degradación", "Índice de condición en 4 clases", "Se mantiene, más el % de superficie bajo el umbral MSAVI 0.4976"],
  ["Oferta de material vegetal", "No incluida", "Directorio, encuesta remota a viveros, capacidad y brecha de viveros"],
]));
x.push(fuente("Fuente: E6 v2.0 «In house», sección «Cambios respecto a la v1.1» (08/10/2026)."));
x.push(para(`La versión 2.0 identifica cinco ítems del índice que requieren conformidad expresa de SESDI por su adecuación a la inspección remota (Cuadro N° ${N.sigC()}).`));
x.push(N.cuadro("Ítems del índice de contenidos mínimos adecuados que requieren conformidad de SESDI"));
x.push(tabla([{ w: 1100, align: C }, { w: 3300 }, { w: 5238 }], ["ÍTEM", "EXIGENCIA DEL ÍNDICE", "ADECUACIÓN PROPUESTA"], [
  ["7.2", "Razón social de la consultora", "ANIN – DIME – SESDI, por Administración Directa"],
  ["9.1", "GPS y formatos de campo", "Formatos de inspección remota; GPS solo en el Plan B"],
  ["9.5 y 10.7", "Número de unidades de muestreo para la inspección de campo", "Parcelas virtuales con cálculo estadístico; la evidencia de campo existente actúa como control"],
  ["9.6 y 10.8", "Mapas con validación en campo", "Validación con imágenes de muy alta resolución, ortofotos y fotos georreferenciadas ya tomadas"],
  ["10.3", "Levantamiento de información de campo", "Se reporta la evidencia ya levantada; no hay levantamiento nuevo"],
]));
x.push(fuente("Fuente: E6 v2.0, sección «Ítems del índice que requieren conformidad de SESDI»."));
x.push(h2("3.2 Revisión del plan por el suscrito"));
x.push(para(`La revisión del plan v2.0 (versión «rev_hscm») introdujo los ajustes del Cuadro N° ${N.sigC()}, contrastados contra la versión original del mismo día.`));
x.push(N.cuadro("Ajustes introducidos en la revisión del plan de trabajo del E6 v2.0"));
x.push(tabla([{ w: 2300 }, { w: 3400 }, { w: 3938 }], ["ELEMENTO", "VERSIÓN ORIGINAL", "VERSIÓN REVISADA (rev_hscm)"], [
  ["Denominación de la modalidad", "«Propuesta ajustada sin parcelas de campo»", "«Propuesta ajustada (Modalidad In house)»: la ausencia de parcelas deja de ser el rasgo definitorio"],
  ["Horizonte de registro del Perfil", "Enero de 2027", "Marzo de 2027, consistente con la reprogramación del Plan de Trabajo Detallado"],
  ["Especialista en teledetección y SIG (P2)", "Por designar", "Ing. Albert Flores Mendoza"],
  ["Operador RPAS, Plan B (P9)", "Por designar", "Ing. Albert Flores Mendoza"],
  ["Especialista en viveros y propagación (P4)", "Ing. Forestal o Agrónomo", "Ing. Forestal, Agrónomo o Biólogo con experiencia en viveros y propagación"],
  ["Solicitudes de información", "Documentos del ACR al Gobierno Regional Piura", "Documentos de inventario florístico de las ACR al Gobierno Regional Piura"],
  ["Activación de la brigada de campo", "Brigada D (dron y permisos); próximo comunicado ENFEN 15/10/2026", "Brigada D (dron), activación como medida de seguridad; próxima evaluación 19/10/2026"],
  ["Índice del documento", "Campo sin actualizar", "Índice con paginación de los 14 capítulos y secciones previas"],
]));
x.push(fuente("Fuente: comparación de «E6_v2.0_sin_parcelas_ANIN.docx» y «E6_v2.0_sin_parcelas_ANIN_rev_hscm.docx» (08/10/2026)."));
x.push(h2("3.3 Actualización del Plan de Trabajo Detallado del Perfil"));
x.push(para("La Actualización N° 01 (11/10/2026) y la Actualización N° 02 (16/10/2026) del Plan de Trabajo Detallado del Perfil, elaboradas por el suscrito para revisión de la Coordinación de Proyectos IN y conformidad de SESDI, formalizan para el E6: el descarte de la reconvocatoria y la ejecución por Administración Directa con la metodología de firma espectral, sin servicio externo ni contratación de personal; el avance de 39 % con estado «con retraso»; la redacción final del informe al 18/11/2026, con las especies priorizadas y la brecha de viveros por escenarios como entregas parciales al E10; y la reasignación de las funciones de la plaza de viveros antes del 23/10/2026."));
x.push(N.cuadro("Hitos del E6 según el plan v2.0 y el Plan de Trabajo Detallado"));
x.push(tabla([{ w: 1000, align: C }, { w: 4400 }, { w: 2100, align: C }, { w: 2138, align: C }], ["HITO", "DESCRIPCIÓN", "PLAN E6 v2.0 (S1 = 12/10)", "PLAN DETALLADO (ACT. 02)"], [
  ["H1", "Diseño muestral aprobado (piloto de 30 parcelas por estrato)", "Fin de S4", "—"],
  ["H2", "Especies priorizadas provisionales para el E10", "Fin de S6 (ratificadas en S7)", "18/11/2026 (H7c)"],
  ["H3", "Brecha de viveros por escenarios para el E10", "Fin de S8", "18/11/2026 (H7c)"],
  ["H4", "Borrador para revisión de SESDI", "Fin de S9", "Versión para revisión emitida (10/2026)"],
  ["H5", "Informe final con observaciones levantadas", "Fin de S10 (18/12/2026)", "18/11/2026 (redacción final)"],
  ["H6g", "Clasificación de 92 tomas S2DR4", "—", "Cumplido (13/10/2026)"],
  ["H7b", "Conformidad de SESDI / Coordinación IN sobre la metodología de firma espectral", "Semana 1 (memorando)", "Pendiente"],
]));
x.push(fuente("Fuente: E6 v2.0, numeral 10.1; Plan de Trabajo Detallado del Perfil, Actualización N° 02, cuadro de hitos."));

// ---------------- 3-A ----------------
const a = [];
a.push(h1("3-A. DIAGNÓSTICO DEL ESTADO DE LOS ECOSISTEMAS"));
a.push(h2("3-A.1 Seguimiento de las actividades"));
a.push(para("Sin levantamiento nuevo de campo, el seguimiento de las actividades de campo se aplicó a la depuración y uso como control de la evidencia de 2026: fichas F-DT de los 117 bloques (evidencia A en 101, B en 12 y C en 4), 222 registros de calicatas y chequeos del E5 (34 sitios de referencia), 346 fichas sociales con 143 registros de especies y ortofotos de vuelos RPAS. El Plan B (inspección rápida de campo) no se activó por la vigencia de la alerta ENFEN. El avance por fase al 14/10/2026 se presenta en el Cuadro N° " + N.sigC() + "."));
a.push(N.cuadro("Avance de las fases del E6 al 14/10/2026"));
a.push(tabla([{ w: 900, align: C }, { w: 5000 }, { w: 3738 }], ["FASE", "ACTIVIDAD", "ESTADO"], [
  ["1", "Clasificación de 92 tomas S2DR4 y Nota Técnica N° 002-2026", "Concluida"],
  ["2", "Regeneración y control de 30 ráster de cobertura", "Concluida (acuerdo 100 % en toma de control)"],
  ["3", "Recorte a 117 bloques, parcelas virtuales y referencias", "Concluida"],
  ["4", "Índice de condición, tipología MSAVI–NDVI y brecha", "Concluida"],
  ["5", "Composición florística y matriz de atributos", "Concluida con vacíos (consulta Delphi pendiente)"],
  ["6", "Censo remoto de viveros y brecha con parámetros del E10", "Pendiente"],
  ["7", "Validación ciega (≈ 600 puntos) y TAU", "Pendiente (estudio definitivo o adenda)"],
]));
a.push(fuente("Fuente: Informe E6, numeral 10.1."));

a.push(h2("3-A.2 Revisión técnica de los insumos producidos por órdenes de servicio"));
a.push(para("El 03/10/2026 se revisaron las carpetas de entregables de las órdenes de servicio que produjeron insumos para el E6, con el resultado del Cuadro N° " + N.sigC() + "."));
a.push(N.cuadro("Revisión técnica de los entregables de las órdenes de servicio insumo del E6"));
a.push(tabla([{ w: 1600 }, { w: 2500 }, { w: 3138 }, { w: 2400 }], ["ORDEN DE SERVICIO", "PRODUCTO", "HALLAZGO DE LA REVISIÓN", "TRATAMIENTO EN EL E6"], [
  ["OS 1148 y OS 2797", "Flora y ecosistemas", "≈ 90 taxones (OS 1148) en listas preliminares con errores taxonómicos y sin depurar", "Depuración contra F-DT, F-DS y chequeos de suelos; matriz de 69 especies nativas"],
  ["OS 1147", "Oferta de material vegetal", "≈ 77 nativas; priorización preliminar de 35 arbóreas y 34 arbustivas; demanda de 2.94 M plantones solo para Lotes SUS en 83 bloques", "Lista recalculada por piso y ecosistema; demanda a recalcular para 117 bloques con parámetros del E10"],
  ["OS 2726", "Fitosanidad y presiones", "Presiones en 58 bloques: sobrepastoreo 96.6 %, quemas 86.2 %, especies invasoras 67.2 %; propagación de 10 especies; protocolo de vivero limpio; costo paramétrico S/ 12 por plantón", "Atributo «ausencia de amenazas» del IC con F-DT-04 de los 117 bloques; costo a sustentar en E10/E12"],
  ["OS 2357", "Sistematización de bloques", "Sistematización de la evidencia de bloques", "Insumo de la clasificación de evidencia A/B/C"],
  ["Conjunto", "Oferta", "Cinco fuentes de gabinete, ninguna en los 15 distritos; 6,000 plantones confirmados (Suyo)", "Directorio ampliado a 12 iniciativas (parte 3-B)"],
], { size: 16 }));
a.push(fuente("Fuente: tablero de avance del Perfil al 03/10/2026 (componentes del E6 y vacíos de integración)."));

a.push(h2("3-A.3 Revisión de la Nota Técnica N° 002-2026-ANIN/DIME/SESDI"));
a.push(para("La Nota Técnica (14/10/2026), dirigida a la Coordinación de Proyectos IN, informa la clasificación de la cobertura vegetal de 92 tomas S2DR4 de 1 m (≈ 1,681 ha cada una; compuesto del 02/09 al 09/10/2026) con un único modelo K-means de 16 clases sobre 8 variables, agrupadas en 9 coberturas, y la asignación de especies probables por piso altitudinal. Al final de la época seca, el 50.9 % de la superficie analizada conserva vegetación verde, el 27.5 % está seca o defoliada y el 13.9 % sin cobertura vegetal efectiva; el piso A (< 1,000 msnm) concentra el suelo expuesto y el piso B (1,000–2,000 msnm) es el mejor conservado."));
a.push(para("La revisión precisó el texto respecto de la versión del 13/10/2026: numeración de la nota, ampliación de los ecosistemas a bosque montano relicto y páramos, y vínculo de las coberturas con los metrados. Se mantienen como limitaciones declaradas que las coberturas 4 y 5 incluyen bosque caducifolio sano sin hojas y que las superficies por toma incluyen traslapes; dos expresiones de la versión revisada requieren corrección (observaciones OE-04 y OE-05)."));

a.push(h2("3-A.4 Revisión técnica del Informe E6 (versión para revisión de SESDI)"));
a.push(para("El Informe E6 cubre los 14 capítulos del índice con tres líneas de evidencia —firma espectral de 92 tomas S2DR4, evidencia de campo 2026 e información secundaria— y declara su adecuación metodológica. Sus resultados principales por provincia se presentan en el Cuadro N° " + N.sigC() + "."));
a.push(N.cuadro("Síntesis del estado de los ecosistemas por provincia (E6)"));
a.push(tabla([{ w: 1500 }, { w: 800, align: C }, { w: 1200, align: R }, { w: 1000, align: R }, { w: 1300, align: R }, { w: 1000, align: R }, { w: 1638, align: R }, { w: 1200, align: R }], ["PROVINCIA", "BLOQ.", "ÁREA (ha)", "MSAVI POND.", "BAJO 0.4976 (ha)", "% BAJO UMBRAL", "RESTAURACIÓN ACTIVA ≤ 0.3813 (ha)", "IC POND."], [
  ["Ayabaca", "10", "1,391.70", "0.548", "525.21", "37.7 %", "42.28", "0.851"],
  ["Huancabamba", "48", "3,551.16", "0.509", "1,599.99", "45.1 %", "925.51", "0.733"],
  ["Morropón", "59", "7,327.37", "0.379", "6,216.63", "84.8 %", "4,535.62", "0.706"],
], { total: ["TOTAL", "117", "12,270.24", "0.436", "8,341.82", "68.0 %", "5,503.41", "0.730"] }));
a.push(fuente("Fuente: Informe E6, Tabla 1."));
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
a.push(N.cuadro("Inspección remota: parcelas virtuales por estrato y precisión alcanzada"));
a.push(tabla([{ w: 2700 }, { w: 1300 }, { w: 1000, align: R }, { w: 1300, align: R }, { w: 900, align: R }, { w: 1100, align: R }, { w: 1338, align: C }], ["ECOSISTEMA", "CONDICIÓN MSAVI", "PARCELAS", "% COB. VEG.", "CV (%)", "ERROR (%)", "¿CUMPLE 15 %?"], [
  ["Bosque estacionalmente seco (Bes-cm)", "Bajo umbral", "789", "79.4", "38.0", "2.7", "Sí"],
  ["Bosque estacionalmente seco (Bes-cm)", "Sobre umbral", "142", "96.9", "7.6", "1.2", "Sí"],
  ["Bosque relicto montano (Br-mvoc)", "Sobre umbral", "27", "96.4", "10.6", "4.0", "Sí"],
  ["Matorral andino (Ma)", "Bajo umbral", "10", "67.9", "34.0", "21.1", "No"],
  ["Matorral andino (Ma)", "Sobre umbral", "152", "96.2", "9.1", "1.4", "Sí"],
  ["Páramo (Pa)", "Sobre umbral", "12", "96.3", "8.2", "4.6", "Sí"],
  ["Zona agrícola / mosaico agroforestal", "Sobre umbral", "233", "97.7", "6.4", "0.8", "Sí"],
]));
a.push(fuente("Fuente: Informe E6, Tabla 11. Se interpretaron 1,374 parcelas en bloques (1,365 con lectura válida) y 102 de referencia. Contraste con la cobertura de 82 fichas F-DT: r = 0.72; error absoluto medio 19.5 pp; sesgo +15.9 pp."));
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
a.push(para("La revisión confirma que el E6 aporta por primera vez una medida de condición por bloque y una delimitación interna cuantificada (restauración activa 5,566.89 ha; enriquecimiento 5,032.34 ha; conservación 1,207.99 ha; tamizaje agrícola 463.01 ha en los 117 bloques), y que la regla «MSAVI para jerarquizar, NDVI para delimitar» se valida en el ámbito: los 16 bloques de Morropón que el NDVI 2025 mostraba enmascarados por verdor pasan a degradación convergente al final de la época seca. Las observaciones de consistencia y de alcance se registran en el numeral 3-A.5."));

a.push(h2("3-A.5 Matriz de observaciones del Diagnóstico del Estado de los Ecosistemas"));
a.push(N.cuadro("Observaciones formuladas al Diagnóstico del Estado de los Ecosistemas (E6) y a sus insumos"));
const OE = [
  ["OE-01", "Insumos OS 1148 / 2797", "Listas florísticas con errores taxonómicos y sin depurar.", "Depurar contra F-DT, F-DS y chequeos de suelos.", "Atendida parcialmente: matriz de 69 especies; 48 taxones confirmados en campo 2026"],
  ["OE-02", "Insumo OS 2726", "Presiones evaluadas solo en 58 bloques.", "Extender el atributo de amenazas a los 117 bloques.", "Atendida: IC con F-DT-04 de los 117 bloques"],
  ["OE-03", "Plan v2.0, 10.1; Plan Detallado", "El plan v2.0 cierra el E6 el 18/12/2026 y cita registro del Perfil en enero de 2027 (10.1) y en marzo de 2027 (sustento); el Plan Detallado fija la redacción final al 18/11/2026.", "Homologar fechas y horizonte en ambos documentos.", "Pendiente"],
  ["OE-04", "Nota Técnica 002-2026, numeral 1", "La versión revisada cita el «catálogo V7»; el catálogo vigente es V6.", "Corregir la referencia al catálogo V6.", "Pendiente"],
  ["OE-05", "Nota Técnica 002-2026, Cuadro 2", "La cobertura 7 no registra superficie ni denominación.", "Confirmar si se fusionó con otra clase y declararlo.", "Pendiente"],
  ["OE-06", "Informe E6, 9.6, 10.8 y 11", "Validación ciega de 567 puntos y matriz de confusión no ejecutadas; el plan v2.0 comprometió exactitud global ≥ 85 %.", "Ejecutar la validación o declararla formalmente como vacío aceptado por SESDI.", "Pendiente"],
  ["OE-07", "Informe E6, Tabla 11", "El estrato de matorral andino bajo umbral tiene 10 de 20 parcelas requeridas (error 21.1 %).", "Densificar el estrato hasta cumplir el error de 15 %.", "Pendiente"],
  ["OE-08", "Informe E6, 10.6", "La referencia del piso A está intervenida (16.7 % de suelo desnudo) y la de los pisos C–D tiene 16 parcelas con confianza baja.", "Incorporar el ACR Bosques Secos de Salitral–Huarmaca y densificar referencias de altura.", "Pendiente"],
  ["OE-09", "Informe E6, 10.9; Plan Detallado", "Coexisten tres criterios de brecha: areal del E6 (8,341.82 ha), media del bloque de los volúmenes DT y media por ecosistema propuesta en el Plan Detallado.", "Adoptar por decisión de SESDI un criterio único para los 117 bloques.", "Pendiente"],
  ["OE-10", "Informe E6, 10.3 y Nota Técnica", "Las coberturas 4 y 5 de fin de época seca incluyen bosque caducifolio sano sin hojas.", "Adquirir compuesto S2DR4 de época de lluvias (feb–abr 2027) y contrastar con referencias.", "Pendiente (estudio definitivo)"],
  ["OE-11", "Informe E6, 10.5", "Cerca del 90 % de las categorías del D.S. N° 043-2006-AG y los atributos de sequía, raíz y palatabilidad figuran «Por determinar»; tres especies se citan como CR «según fuentes del proyecto».", "Completar con consulta Delphi y verificar contra el anexo oficial del decreto.", "Pendiente"],
  ["OE-12", "Informe E6, 10.3", "Sesgo de +15.9 pp de la cobertura S2DR4 frente a F-DT por el herbazal seco.", "Declarar el sesgo y su efecto en el IC.", "Atendida: declarado en el informe"],
  ["OE-13", "Plan v2.0, 12; hito H7b", "No se ha emitido el memorando de SESDI que acepte la adecuación metodológica.", "Emitir el memorando de conformidad de la adecuación.", "Pendiente"],
];
a.push(tabla([{ w: 750, align: C }, { w: 1550 }, { w: 2900 }, { w: 2338 }, { w: 2100 }], ["CÓD.", "DOCUMENTO / SECCIÓN", "OBSERVACIÓN", "LEVANTAMIENTO REQUERIDO", "ESTADO AL CORTE"], OE, { size: 15 }));
a.push(fuente("Fuente: revisión técnica del suscrito sobre el plan v2.0, la Nota Técnica N° 002-2026 y el Informe E6 (versión para revisión de SESDI)."));
a.push(h2("3-A.6 Estado de la conformidad"));
a.push(para("El Diagnóstico del Estado de los Ecosistemas se encuentra en versión para revisión de SESDI. La conformidad está condicionada a: (i) la emisión del memorando de SESDI que acepte la adecuación metodológica (hito H7b, pendiente); (ii) el levantamiento de las observaciones OE-03 a OE-09, OE-11 y OE-13; y (iii) la decisión sobre el criterio único de brecha. Las observaciones OE-10 y la estructura vertical medida (DAP, biomasa, TAU) se trasladan como vacíos declarados al estudio definitivo."));

// ---------------- 3-B ----------------
const bb = [];
bb.push(h1("3-B. ESTUDIO DE OFERTA DE MATERIAL VEGETAL DE PROPAGACIÓN"));
bb.push(h2("3-B.1 Plan de trabajo del componente"));
bb.push(para("La versión 2.0 del plan incorpora la oferta de material vegetal, ausente en la versión 1.1, como un censo remoto de viveros verificado en imágenes, con la brecha calculada contra la demanda por especie y año que fije el E10. Sus pasos son: (i) directorio de viveros (SERFOR–ATFFS Piura, AGRO RURAL, Gobierno Regional, municipalidades, proyectos especiales Chira-Piura y Alto Piura, UNP, UNF, ONG, viveros privados y comunales); (ii) encuesta remota a todo el directorio por teléfono y KoboToolbox; (iii) verificación de camas y tinglados en imágenes de muy alta resolución; (iv) mapa de fuentes de semilla; (v) oferta efectiva por especie y accesibilidad por tiempo de viaje; y (vi) brecha por especie y año y brecha de viveros permanentes y temporales con un modelo de localización-asignación."));
bb.push(para([b("Demanda y brecha. "), "D(e,a) = Σₘ A(m,a) · dₘ · p(e,m) · (1 + r); B(e,a) = D(e,a) − O(e,a), donde A son las hectáreas por medida, d la densidad de plantación, p la proporción de la especie en la mezcla, r la tasa de reposición y O la oferta efectiva. Para la entrega parcial al E10 se calculan tres escenarios de demanda (bajo, medio y alto) con las hectáreas por bloque de la microlocalización y densidades de referencia declaradas como supuesto."]));
bb.push(h2("3-B.2 Revisión técnica de los insumos de oferta"));
bb.push(para("Los entregables de las OS 1147 (oferta de material vegetal) y 2726 (fitosanidad y propagación), revisados el 03/10/2026, aportan una priorización preliminar de 35 especies arbóreas y 34 arbustivas sobre ≈ 77 nativas, protocolos de propagación de 10 especies, un protocolo de vivero limpio y un costo paramétrico de S/ 12 por plantón. Sus limitaciones para el Perfil son que la demanda de plantones (2.94 millones) se calculó solo para Lotes SUS en 83 bloques, que las cinco fuentes de oferta identificadas en gabinete no están en los 15 distritos y que la única oferta confirmada son 6,000 plantones (Suyo). El tablero de avance al 03/10/2026 estimó el componente de oferta en ≈ 12 % frente a ≈ 45 % del diagnóstico del ecosistema."));
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
bb.push(para("La única capacidad cuantificada son los 21,000 plantones de los viveros del Gobierno Regional en Hualtacal, frente a una demanda potencial que se dimensionará sobre 5,503.41 ha de restauración activa y 2,838.41 ha de enriquecimiento. El E6 incluye 16 fichas técnicas de especies nativas recomendadas: Vachellia macracantha (aromo), Eriotheca ruizii (pasallo), Loxopterygium huasango (hualtaco), Bursera graveolens (palo santo), Caesalpinia paipai (charán), Cordia lutea (overo), Ceiba trischistandra (ceibo), Colicodendron scabridum (sapote), Prosopis pallida (algarrobo), Handroanthus chrysanthus (guayacán), Caesalpinia spinosa (tara), Inga feuilleei (guaba), Alnus acuminata (aliso), Capparicordis crotonoides (bichayo), Baccharis latifolia (chilca) y Tecoma stans (huaranhuay). La demanda y la brecha se calculan en la hoja 16 de las matrices al ingresar las densidades, la mezcla y la reposición del E10; no se asumieron valores."));
bb.push(h2("3-B.4 Matriz de observaciones del Estudio de Oferta de Material Vegetal"));
bb.push(N.cuadro("Observaciones formuladas al Estudio de Oferta de Material Vegetal de Propagación"));
const OM = [
  ["OM-01", "Insumo OS 1147", "La demanda de plantones (2.94 M) se calculó solo para Lotes SUS en 83 bloques.", "Recalcular para los 117 bloques con las densidades, la mezcla y la reposición del E10 (hoja 16).", "Pendiente"],
  ["OM-02", "Insumos de oferta", "Ninguno de los viveros de gabinete está en los 15 distritos; oferta confirmada de 6,000 plantones.", "Ampliar el directorio al ámbito.", "Atendida parcialmente: 12 iniciativas; una con capacidad declarada (21,000)"],
  ["OM-03", "Insumo OS 2726", "Costo paramétrico de S/ 12 por plantón sin análisis de costos unitarios.", "Sustentar el costo por especie y tipo de vivero en el E10/E12.", "Pendiente"],
  ["OM-04", "E6, fase 6", "El censo remoto de viveros (encuesta KoboToolbox y verificación en imágenes) no se ha ejecutado.", "Ejecutar el censo y calcular la oferta efectiva por especie.", "Pendiente"],
  ["OM-05", "Plan Detallado", "Plaza de Especialista en Infraestructura Rural para la Producción (viveros) desierta.", "Reasignar formalmente las funciones antes del 23/10/2026.", "Pendiente"],
  ["OM-06", "E6, 9.7", "No se ha elaborado el mapa de fuentes de semilla (árboles semilleros y rodales).", "Elaborarlo con viveristas, saberes ancestrales y registros de SERFOR.", "Pendiente"],
  ["OM-07", "OS 1147 / E6 / E5", "La priorización de 35 arbóreas y 34 arbustivas (OS 1147) y la lista del E6 no están conciliadas con las especies de la matriz suelo–intervención del E5.", "Conciliar una lista única por piso, ecosistema y grupo de suelos.", "Pendiente"],
  ["OM-08", "E6, Tabla 15", "El vivero del entorno de M22B1 produce pino radiata, especie exótica.", "Excluir de la oferta efectiva las especies no nativas o incompatibles.", "Criterio a aplicar en el cálculo"],
];
bb.push(tabla([{ w: 750, align: C }, { w: 1550 }, { w: 2900 }, { w: 2338 }, { w: 2100 }], ["CÓD.", "DOCUMENTO / SECCIÓN", "OBSERVACIÓN", "LEVANTAMIENTO REQUERIDO", "ESTADO AL CORTE"], OM, { size: 15 }));
bb.push(fuente("Fuente: revisión técnica del suscrito sobre los insumos de las OS 1147 y 2726, el plan v2.0 y el Informe E6."));
bb.push(h2("3-B.5 Estado de la conformidad"));
bb.push(para("El componente de oferta de material vegetal es el de menor avance del E6: el directorio está identificado, pero el censo remoto, la oferta efectiva por especie, el mapa de fuentes semilleras y la brecha de viveros están pendientes. No corresponde conformidad al corte; su emisión requiere el levantamiento de las observaciones OM-01 a OM-07 y la entrega parcial al E10 de las especies priorizadas y de la brecha por escenarios, programada para el 18/11/2026."));

// ---------------- Cierre ----------------
const z = [];
z.push(h1("4. CONCLUSIONES"));
[
  "El Diagnóstico del Estado de los Ecosistemas y el Estudio de Oferta de Material Vegetal se ejecutan como un solo Entregable 6 por Administración Directa, tras el procedimiento desierto y la alerta de El Niño Costero; su plan de trabajo fue revisado y actualizado de la versión 1.1 a la 2.0 «In house» y el Plan de Trabajo Detallado del Perfil fija su redacción final al 18/11/2026.",
  "Se revisaron técnicamente los insumos de las OS 1147, 1148, 2357, 2726 y 2797, la Nota Técnica N° 002-2026 y el Informe E6; se formularon 13 observaciones al diagnóstico (OE) y 8 a la oferta (OM), de las cuales 4 están atendidas total o parcialmente.",
  "El diagnóstico cuantifica 8,341.82 ha (68.0 %) bajo el umbral MSAVI y clasifica 14 bloques como moderados y 6 como severos; la validación ciega, la densificación del matorral andino y la consulta Delphi son condiciones para su conformidad.",
  "La oferta de material vegetal identificada es mínima (21,000 plantones) frente a la escala de la intervención; el censo de viveros, la demanda para 117 bloques y la reasignación de la plaza de viveros son la ruta crítica del componente.",
].forEach((t, k) => z.push(num(`4.${k + 1}.`, t)));
z.push(h1("5. RECOMENDACIONES"));
[
  "Emitir el memorando de SESDI que acepte la adecuación metodológica y los ítems adaptados del índice (hito H7b).",
  "Adoptar un criterio único de brecha para los 117 bloques y homologar las fechas del E6 entre su plan v2.0 y el Plan de Trabajo Detallado.",
  "Ejecutar la validación ciega de 567 puntos, densificar el estrato de matorral andino bajo umbral y las referencias de altura, e incorporar el ACR Bosques Secos de Salitral–Huarmaca como referencia del piso A.",
  "Reasignar formalmente la plaza de viveros, ejecutar el censo remoto y el mapa de fuentes semilleras, y entregar al E10 la oferta efectiva y la brecha de plantones por escenarios.",
  "Conciliar una lista única de especies por piso, ecosistema y grupo de suelos entre el E6, la matriz suelo–intervención del E5 y los Lotes SUS, en coordinación con el Especialista en Intervenciones Vegetativas de Uso Sostenible.",
].forEach((t, k) => z.push(num(`5.${k + 1}.`, t)));
z.push(h1("6. MEDIOS VERIFICADORES"));
z.push(N.cuadro("Medios verificadores de la actividad d)"));
z.push(tabla([{ w: 3500 }, { w: 4138 }, { w: 2000, align: C }], ["MEDIO VERIFICADOR", "UBICACIÓN / IDENTIFICADOR", "FECHA"], [
  ["E6 v2.0 «In house» – versión original", "Google Drive ID 1B3QsYMkodQVnK5ZwQU-7BI0t9Yn6xmoR", "08/10/2026"],
  ["E6 v2.0 – revisión del suscrito (Word, PDF y Markdown)", "Google Drive ID 1rxdf63E-TOJpUEUWWazdkth6WjdwU4HY; 1lbr4-I5-OH8Nqd4hBccIEizEy3H-rV7V; 18FQxeA3WasWZOCjg4mVapnjQf6D_Est6", "08–14/10/2026"],
  ["Nota Técnica N° 002-2026-ANIN/DIME/SESDI (original y revisada)", "Google Drive ID 1EZk6-sdLsbFfkiuI4jma4IxlTVsTSXyT; 137Syk2NZPG9bTZ8rwrxtpTmjbzOtLhTy; PDF 17Vt9Z1NBZPTGF933r4jhqFNu0N7MP-nR", "13–14/10/2026"],
  ["Informe E6 – versión para revisión de SESDI", "Google Drive ID 1NfYjSV3QXI3uA4eHsMNQNTw9vdY5RdDA", "15/10/2026"],
  ["Matrices del diagnóstico del ecosistema (E6)", "Google Drive ID 1Uie1il6FeHI4SD8QhKXsfDof7cm4NF8J", "15/10/2026"],
  ["Plan de Trabajo Detallado, Actualización N° 02, y matrices de planificación", "Google Drive ID 1QkhMG4cWkPFAyhxOH8qtaWG2YVCCJsGy; 1EF3u3btVKiK2DcUKrHYcJZv-4co_arCc", "16–17/10/2026"],
  ["Tablero de avance del Perfil (revisión de entregables OS 1147, 1148, 2357, 2726 y 2797)", "Google Drive ID 1wRdIvYH8ni8kbj3XsDKa0cgrgSIdVWn6", "03/10/2026"],
  ["Entregables de las OS insumo del E6", "Google Drive: PRIMER/SEGUNDO/TERCER_ENTREGABLE de las OS 1147, 1148, 2357, 2726 y 2797", "22/09/2026"],
  ["Referencia técnica de fuentes semilleras", "Google Drive ID 1hS8QlbTu_Ks05061BV4FAbAqbfUpikzy («Fuentes semilleras de los Bosques Secos»)", "—"],
  ["Matrices de observaciones en formato nativo", "Anexo 4 (Excel), hojas «Obs_E6_Ecosistemas» y «Obs_E6_Material_vegetal»", "Octubre 2026"],
], { size: 15 }));
z.push(fuente("Fuente: elaboración propia."));

L.guardar([
  seccionCaratula(DEF.anexo3),
  L.seccion(x, { pie: PIE }),
  seccionCaratula(DEF.anexo3A),
  L.seccion(a, { pie: PIE }),
  seccionCaratula(DEF.anexo3B),
  L.seccion([...bb, ...z], { pie: PIE }),
], OUT, "Anexo N° 3 – Actividad d)");
module.exports = { OE, OM };
