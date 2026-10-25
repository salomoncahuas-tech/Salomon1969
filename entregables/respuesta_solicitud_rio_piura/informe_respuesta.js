// Informe técnico de respuesta a la solicitud de información sobre el avance del Perfil IN Piura.
//
// Lee las cifras de salidas/cifras_informe.json (generado por generar_paquete_respuesta.py, las mismas que
// calculan las fórmulas del Excel de anexos) y escribe salidas/IN_Piura_Informe_Avance_Perfil_Respuesta.docx.
//
//   python entregables/respuesta_solicitud_rio_piura/generar_paquete_respuesta.py
//   NODE_PATH=$(npm root -g) node entregables/respuesta_solicitud_rio_piura/informe_respuesta.js
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, ImageRun } = require("docx");

const RAIZ = path.resolve(__dirname, "..", "..");
const A = require(path.join(RAIZ, ".claude", "skills", "anin-in-piura", "scripts", "anin_docx"));
const { W, L, R, C, f0, f1, f3, pct, run, b, para, h1, h2, bullet, cap, fuente, tabla, meta } = A;
const SAL = path.join(__dirname, "salidas");
const D = JSON.parse(fs.readFileSync(path.join(SAL, "cifras_informe.json"), "utf8"));

const PD = (t) => run(`[POR DEFINIR${t ? ": " + t : ""}]`, { bold: true, color: "9C2A00" });
const fecha = new Date(D.fecha + "T12:00:00");
const MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "setiembre", "octubre",
  "noviembre", "diciembre"];
const fechaTxt = `${fecha.getDate()} de ${MESES[fecha.getMonth()]} de ${fecha.getFullYear()}`;
const fechaCorta = (iso) => iso.split("-").reverse().join("/");

const BL = D.bloques, CP = D.ccpp, SO = D.social, MS = D.msavi, CA = D.carcavas;
const USD = 372442809 / 3.4;

function figura(archivo, ancho, alto) {
  return new Paragraph({ alignment: C, spacing: { before: 120, after: 60 }, keepNext: true,
    children: [new ImageRun({ type: "png", data: fs.readFileSync(archivo),
      transformation: { width: ancho, height: alto },
      altText: { title: path.basename(archivo), description: path.basename(archivo), name: path.basename(archivo) } })] });
}

// ── cuadros ──
const sumar = (xs, k) => xs.reduce((s, x) => s + x[k], 0);

const tProv = tabla(
  [{ w: 2000, align: L }, { w: 1300, align: R }, { w: 1800, align: R }, { w: 1100, align: R }, { w: 1700, align: R },
    { w: 1738, align: R }],
  ["Provincia", "N.° distritos", "Área de distritos (ha, ref.)", "N.° bloques", "Área de bloques (ha)", "Área de influencia (ha)"],
  D.por_provincia.map((p) => [p.prov, f0(p.nd), f1(sumar(D.por_distrito.filter((d) => d.prov === p.prov), "dist_ha")),
    f0(p.n), f3(p.ha), f3(p.ai)]),
  ["TOTAL", f0(15), f1(D.ambito_ha), f0(BL.n), f3(BL.ha), f3(BL.ai_ha)]);

const tDist = tabla(
  [{ w: 1700, align: L }, { w: 2600, align: L }, { w: 1000, align: R }, { w: 1500, align: R }, { w: 1500, align: R },
    { w: 1338, align: R }],
  ["Provincia", "Distrito", "Bloques", "Bloques (ha)", "AI (ha)", "Bloques + AI (ha)"],
  D.por_distrito.map((d) => [d.prov, d.dist, f0(d.n), f3(d.ha), f3(d.ai), f3(d.ha + d.ai)]),
  ["TOTAL", "15 distritos", f0(BL.n), f3(BL.ha), f3(BL.ai_ha), f3(BL.ha + BL.ai_ha)]);

const tCrit = tabla(
  [{ w: 760, align: C }, { w: 1700, align: L }, { w: 4540, align: L }, { w: 2638, align: L }],
  ["Paso", "Etapa", "Criterios y reglas", "Fuente de información"],
  [
    ["1", "Mesolocalización", "Peligro integrado (PI) por MCA-AHP de PMM (geología, pendiente, TPI, cobertura; desencadenante: precipitación acumulada multianual; tendencia a la degradación TAU/NDVI), EPH y PGI (pendiente, escorrentía total, NDVI); 5 niveles.", "Sentinel-2, Landsat, DEM, WorldClim, SENAMHI, INGEMMET"],
    ["2", "Digitalización", "Bloques en áreas de PI medio, alto y muy alto, sobre microcuencas priorizadas de la cuenca alta (zona de origen del peligro).", "Mapa de PI; imágenes satelitales"],
    ["3", "Filtros de exclusión", "Restos arqueológicos; concesiones mineras activas; infraestructura eléctrica (franja de 10 m); Áreas Naturales Protegidas.", "SIGDA (Min. Cultura), INGEMMET, OSINERGMIN, SERNANP"],
    ["4", "Idoneidad", "Accesibilidad; pendiente ≤ 75 %; cobertura de los 15 distritos priorizados; déficit hídrico futuro (MPI-ESM1-2HR, SSP5-8.5, 2021-2040).", "Red vial MTC, DEM, escenarios CMIP6"],
    ["5", "Tamizaje predial", "Máximo 19 predios rurales interceptados por bloque; exclusión de áreas agrícolas.", "Catastro rural; imágenes satelitales"],
    ["6", "Verificación de campo", "Accesibilidad real, uso actual, evidencia de movimientos en masa, riesgo social y aceptación de titulares; codificación M#-B#.", "Fichas F-DT, F-EE, F-DS y F-LA; vuelos de dron"],
  ]);

const tPob = tabla(
  [{ w: 1700, align: L }, { w: 1000, align: R }, { w: 1500, align: R }, { w: 1200, align: R }, { w: 1200, align: R },
    { w: 1538, align: R }, { w: 1500, align: R }],
  ["Provincia", "CCPP", "Población total", "Hombres", "Mujeres", "Población vulnerable", "Viviendas"],
  CP.por_provincia.map((p) => [p.prov, f0(p.n), f0(p.pob), f0(p.h), f0(p.m), f0(p.vul), f0(p.viv)]),
  ["TOTAL", f0(CP.n), f0(CP.pob), f0(CP.h), f0(CP.m), f0(CP.vul), f0(CP.viv)]);

const tPobDist = tabla(
  [{ w: 1700, align: L }, { w: 3100, align: L }, { w: 1200, align: R }, { w: 1900, align: R }, { w: 1738, align: R }],
  ["Provincia", "Distrito del centro poblado", "CCPP", "Población total", "% de la población"],
  CP.por_distrito.map((d) => [d.prov, d.dist, f0(d.n), f0(d.pob), pct(d.pob / CP.pob)]),
  ["TOTAL", `${CP.por_distrito.length} distritos`, f0(CP.n), f0(CP.pob), "100.0 %"]);

const tSocial = tabla(
  [{ w: 6638, align: L }, { w: 3000, align: R }],
  ["Indicador de avance (bloques V6)", "Valor"],
  [
    ["Registros centro poblado–bloque programados", f0(SO.registros)],
    ["Tamaño de muestra de hogares (N = 8,865.9 viviendas; 95 % de confianza; e = 5 %)", f0(SO.n)],
    ["Fichas F-DS-01 (socioeconómica de hogar) reportadas", f0(SO.ds01)],
    ["Registros con F-DS-01 completo", `${f0(SO.ds01_completo)} de ${f0(SO.registros)}`],
    ["Entrevistas F-DS-03 a autoridades y líderes reportadas", f0(SO.ds03)],
    ["Registros con F-DS-03 completo", `${f0(SO.ds03_completo)} de ${f0(SO.registros)}`],
    ["Registros con ficha de saberes ancestrales completa", `${f0(SO.saberes_completo)} de ${f0(SO.registros)}`],
    ["Centros poblados descartados o reemplazados por el equipo técnico", f0(SO.descartados)],
    ["Última fecha de evaluación registrada", fechaCorta(SO.fecha_ult)],
  ]);

const ecoFilas = D.ecosistemas.filter((e) => e.ha >= 0.01);
const tEco = tabla(
  [{ w: 3900, align: L }, { w: 1100, align: R }, { w: 1600, align: R }, { w: 1300, align: R }, { w: 1738, align: R }],
  ["Ecosistema (MINAM 2018)", "Bloques", "Área en bloques (ha)", "% de bloques", "Área en el ámbito (ha)"],
  ecoFilas.map((e) => [e.eco, f0(e.n), f3(e.ha), pct(e.pct), e.amb_ha ? f3(e.amb_ha) : "—"]),
  ["TOTAL", "", f3(sumar(ecoFilas, "ha")), "100.0 %", "—"]);

const DNS = [["dn1", "DN 1", "≤ 0.2650"], ["dn2", "DN 2", "0.2650 – 0.3813"], ["dn3", "DN 3", "0.3813 – 0.4976"],
  ["dn4", "DN 4", "0.4976 – 0.6139"], ["dn5", "DN 5", "> 0.6139"]];
const tMsavi = tabla(
  [{ w: 1600, align: L }, { w: 2800, align: C }, { w: 2600, align: R }, { w: 2638, align: R }],
  ["Clase", "Rango MSAVI 2024", "Superficie (ha)", "% de los bloques"],
  DNS.map(([k, n, r]) => [n, r, f3(MS.dn[k]), pct(MS.dn[k] / MS.total_ha)]),
  ["TOTAL", "", f3(MS.total_ha), "100.0 %"]);

const tInv = tabla(
  [{ w: 2700, align: L }, { w: 1150, align: C }, { w: 1150, align: R }, { w: 1400, align: R }, { w: 1500, align: L },
    { w: 1738, align: L }],
  ["Capa / producto", "Geometría", "Registros", "Superficie (ha)", "Formatos", "Estado"],
  D.inventario.map((r) => [r[1], r[3], r[4] == null ? "—" : f0(r[4]), r[5] == null ? "—" : f3(r[5]), r[6], r[8]]));

const INV = D.involucrados;
const tInvol = tabla(
  [{ w: 2100, align: L }, { w: 2500, align: L }, { w: 2500, align: L }, { w: 2538, align: L }],
  ["Grupo de involucrados", "Problemas percibidos", "Intereses", "Estrategias y acuerdos"],
  INV.map((r) => [r[0], r[2], r[3], `${r[4]}. Acuerdos: ${r[5]}`]));

// ── documento ──
const hijos = [
  new Paragraph({ alignment: C, spacing: { before: 60, after: 60 },
    children: [run("INFORME TÉCNICO N.° ", { bold: true, size: 26 }), run("[POR DEFINIR]", { bold: true, size: 26, color: "9C2A00" }),
      run("-2026-ANIN/DIME-SESDI", { bold: true, size: 26 })] }),
  new Paragraph({ alignment: C, spacing: { after: 200 },
    children: [run("AVANCE DE LA FORMULACIÓN DEL PERFIL DE INVERSIÓN «INFRAESTRUCTURA NATURAL» — PROYECTO IN PIURA",
      { bold: true, size: 22, color: A.VERDE })] }),
  meta([
    ["Para", "[POR DEFINIR: titular de la dependencia solicitante — Programa de Inversión del Río Piura]"],
    ["De", "Subdirección de Estudios de Inversión — SESDI / DIME"],
    ["Asunto", "Remisión de información actualizada sobre el avance del Perfil de Inversión «Infraestructura Natural» para su incorporación en el Programa de Inversión del Río Piura"],
    ["Referencia", "[POR DEFINIR: documento de solicitud, número y fecha]"],
    ["Proyecto", "Recuperación del servicio de regulación de riesgos naturales y recuperación de ecosistemas degradados en la Cuenca Alta del Río Piura — CUI 2669244"],
    ["Fase / estado", "Preinversión — Perfil en formulación (Sistema Invierte.pe)"],
    ["Fecha de corte", `Información disponible al ${fechaTxt}`],
    ["Elaborado por", "Ing. Héctor Salomón Cahuas Miller — Especialista en Infraestructura Verde, Proyecto IN Piura"],
  ], 2000),

  h1("1. ANTECEDENTES"),
  para(["Mediante el documento de la referencia se solicita información actualizada sobre los avances en la formulación del Perfil de Inversión denominado «Infraestructura Natural», a fin de incorporar su alcance en el Programa de Inversión del Río Piura y realizar las acciones técnicas correspondientes. Se solicitan seis componentes: (1) delimitación del ámbito de estudio, (2) criterios de selección, (3) diagnóstico de la población afectada, (4) diagnóstico de las Unidades Productoras, (5) diagnóstico de involucrados y (6) información georreferenciada y cartográfica en formatos nativos."]),
  para(["El Proyecto IN Piura (CUI 2669244) es formulado por la ANIN, a través de la DIME y la SESDI, en el marco del Sistema Nacional de Programación Multianual y Gestión de Inversiones (Invierte.pe). Su monto de inversión referencial es ", b("S/ 372,442,809"), ` (≈ USD ${f1(USD / 1e6)} millones a un tipo de cambio referencial de 3.40 S/ por USD), que comprende el estudio de preinversión, el expediente técnico, la ejecución (5 años) y la operación y mantenimiento (15 años), con financiamiento por `, b("Administración Directa"), "."]),
  para(["El proyecto atiende la brecha «Porcentaje de superficie de ecosistemas degradados que brindan servicios ecosistémicos que requieren de recuperación» (R.M. N.° 00213-2024-MINAM), mediante tres líneas de intervención: infraestructura natural (revegetación y reforestación con especies nativas, con viveros), infraestructura marrón y complementaria (zanjas de infiltración, terrazas, diques para el control de cárcavas, clausura y manejo de pastizales, barreras vivas y obras de estabilización) y gobernanza y gestión comunitaria."]),

  h1("2. OBJETIVO"),
  para("Remitir el estado de avance de la formulación del Perfil, con la información técnica y geoespacial disponible a la fecha de corte, para su incorporación en el Programa de Inversión del Río Piura, señalando de forma expresa la información que aún se encuentra en elaboración."),

  h1("3. ALCANCE Y FUENTES DE INFORMACIÓN"),
  para("La información corresponde a la versión vigente del catálogo de bloques de intervención (V6) y a los productos de gabinete y campo consolidados en el repositorio del proyecto. Las cifras de este informe se calculan directamente de esas fuentes y coinciden con las fórmulas del Excel de anexos:"),
  bullet([b("Bloques V6: "), `${BL.n} bloques de intervención con ${f3(BL.ha)} ha (catálogo V6 y capa Bloques V6), en ${BL.n_distritos} distritos y ${BL.n_microcuencas} microcuencas.`]),
  bullet([b("Área de influencia: "), `capa aprobada AI_aprobado_2 (${f0(BL.ai_n)} polígonos, ${f3(BL.ai_ha)} ha), sin los AI de bloques retirados.`]),
  bullet([b("Población: "), "capa de centros poblados del INEI con población censal y asociación centro poblado–bloque del equipo social (corte 24.09.2026); hoja de control del trabajo de campo social."]),
  bullet([b("Unidad Productora: "), "Mapa Nacional de Ecosistemas (MINAM 2018), MSAVI 2024 clasificado por bloque y caracterización de cárcavas."]),
  para(["Los registros de campo de las fichas de diagnóstico territorial (F-DT), elementos expuestos (F-EE), diagnóstico social (F-DS-02 a F-DS-07) y liberación de áreas (F-LA) se encuentran en la base de datos del aplicativo IN Piura y no forman parte de este corte; su consolidado se remitirá posteriormente ", PD(""), ". Los resultados del perfil son preliminares y pueden variar con la verificación de campo."]),

  h1("4. DELIMITACIÓN DEL ÁMBITO DE ESTUDIO (PUNTO 1)"),
  h2("4.1 Área de estudio"),
  para([`El área de estudio comprende la parte alta de la cuenca del río Piura en `, b("15 distritos de 3 provincias"), ` del departamento de Piura: Frías (Ayabaca); Canchaque, Huancabamba, Huarmaca, Lalaquiz y San Miguel de El Faique (Huancabamba); y Buenos Aires, Chalaco, Chulucanas, Morropón, Salitral, San Juan de Bigote, Santa Catalina de Mossa, Santo Domingo y Yamango (Morropón). La superficie de los 15 distritos suma ${f1(D.ambito_ha)} ha (capa INEI simplificada, referencial).`]),
  para("Criterios técnicos considerados para su delimitación:"),
  bullet([b("Hidrográfico y de origen del peligro: "), "la intervención se ubica en la cuenca alta, zona de origen de los movimientos en masa, la erosión y la escorrentía que generan inundaciones aguas abajo; no en la zona de impacto (valle y planicie costera)."]),
  bullet([b("Unidad Productora: "), "presencia de ecosistemas degradados (bosque estacionalmente seco de colina y montaña, matorral andino, bosque relicto montano y páramo) que brindan los servicios de regulación de riesgos naturales y regulación hídrica."]),
  bullet([b("Priorización por peligro: "), `microcuencas priorizadas por el modelamiento de peligro integrado de la mesolocalización (${BL.n_microcuencas} microcuencas con bloques V6).`]),
  bullet([b("Político-administrativo y de población: "), "distritos que contienen las microcuencas priorizadas y los centros poblados vinculados a los ecosistemas por intervenir."]),
  cap("Cuadro 1. Ámbito de estudio, bloques V6 y área de influencia por provincia"),
  tProv,
  fuente("Fuente: catálogo y capa de bloques V6; área de influencia aprobada AI_aprobado_2; límites distritales INEI (referencial). Elaboración: ANIN-DIME-SESDI."),
  h2("4.2 Área de influencia"),
  para([`El área de influencia comprende los `, b(`${BL.n} bloques de intervención V6 (${f3(BL.ha)} ha)`), `, donde se ejecutarán las intervenciones, y el `, b(`área de influencia aprobada (${f0(BL.ai_n)} polígonos, ${f3(BL.ai_ha)} ha)`), `, que rodea a cada bloque sin superponerse con él; en conjunto suman ${f3(BL.ha + BL.ai_ha)} ha. Cada polígono de influencia se codifica AI_<bloque> y se asigna al distrito y la provincia de su bloque.`]),
  para(["El área de intervención directa del proyecto se define en un rango de 10,000 a 15,000 ha; las ", f3(BL.ha), " ha corresponden a la suma de los bloques preliminares, que se ajustará con la verificación de campo y el saneamiento de áreas. El criterio técnico de delimitación del área de influencia aprobada (AI_aprobado_2) se documentará en el perfil ", PD("criterio y parámetros de delimitación del AI"), "."]),
  cap("Cuadro 2. Bloques V6 y área de influencia por distrito"),
  tDist,
  fuente("Fuente: catálogo y capa de bloques V6; área de influencia aprobada AI_aprobado_2. Detalle por bloque en el Anexo A1."),
  figura(path.join(RAIZ, "mapas", "salidas", "IN_Piura_01_mapa_general.png"), 640, 452),
  cap("Figura 1. Área de influencia y bloques de intervención V6 (Lámina 1, formato A3 en el paquete SIG)"),

  h1("5. CRITERIOS DE SELECCIÓN (PUNTO 2)"),
  para("Los ámbitos de intervención se identifican y priorizan mediante el análisis de riesgo territorial en contexto de cambio climático (AdR-CCC), según la Nota Técnica de Meso y Microlocalización de Infraestructura Natural (ANIN-DIME, 2023) y el Anexo 2 (GdR-CCC) de la Guía General para la identificación, formulación y evaluación de proyectos de inversión (DGPMI-MEF). El proceso combina criterios territoriales, de peligro, de exclusión, de idoneidad, prediales y sociales en seis pasos:"),
  cap("Cuadro 3. Criterios de selección y priorización de los ámbitos de intervención"),
  tCrit,
  fuente("Fuente: procedimiento AdR-CCC del Proyecto IN Piura. Matriz completa con productos por paso en el Anexo A2."),
  para([`Resultado vigente: `, b(`${BL.n} bloques (catálogo V6, ${f3(BL.ha)} ha)`), ` distribuidos en ${BL.n_microcuencas} microcuencas y ${BL.n_distritos} distritos. En la versión V6 se retiraron 12 códigos de versiones anteriores (1, 7, 25, 29, 32, 33, 46, 48, 68, 74, 75 y M18B5), que no se incluyen en ningún producto de este informe. Las microcuencas con mayor superficie de bloques son ${D.microcuencas.slice(0, 3).map((m) => `${m.cod} (${f3(m.ha)} ha, ${m.n} bloques)`).join(", ")}.`]),
  para(["Los resultados cuantitativos del modelamiento de peligro integrado por microcuenca y bloque se incorporarán con la remisión de las capas de peligro ", PD("resultados del PI por microcuenca y bloque"), "."]),

  h1("6. DIAGNÓSTICO DE LA POBLACIÓN AFECTADA (PUNTO 3)"),
  h2("6.1 Metodología"),
  bullet([b("Identificación: "), "asociación de cada bloque con los centros poblados de su entorno por el equipo social, sobre la capa de centros poblados del INEI (población total y por sexo, población vulnerable y viviendas particulares). Un centro poblado puede atender a varios bloques; se cuenta una sola vez."]),
  bullet([b("Muestreo: "), `diseño para población finita, estratificado por centro poblado con asignación proporcional a las viviendas: N = 8,865.9 viviendas, 95 % de confianza, margen de error de 5 % y p = q = 0.5, que da una muestra de ${f0(SO.n)} hogares.`]),
  bullet([b("Instrumentos: "), "F-DS-01 diagnóstico socioeconómico, F-DS-02 mapeo de actores, F-DS-03 entrevista a autoridades y líderes, F-DS-04 actas de talleres, F-DS-05 conflictos socioambientales, F-DS-06 percepción de peligros y cambio climático, F-DS-07 disposición a participar y consentimiento, y F-EE-02 elementos expuestos (población y viviendas)."]),
  h2("6.2 Resultados preliminares"),
  para([`Se identifican `, b(`${f0(CP.n)} centros poblados`), ` asociados a los bloques V6, con `, b(`${f0(CP.pob)} habitantes`), ` (${f0(CP.h)} hombres y ${f0(CP.m)} mujeres en los centros poblados con desagregación), ${f0(CP.vul)} personas en condición de vulnerabilidad y ${f0(CP.viv)} viviendas particulares. ${f0(CP.rurales)} centros poblados son rurales y ${f0(CP.multi_bloque)} están asociados a más de un bloque.`]),
  cap("Cuadro 4. Población de los centros poblados asociados a los bloques V6, por provincia"),
  tPob,
  fuente(`Fuente: INEI (capa de centros poblados); asociación centro poblado–bloque del equipo social (corte 24.09.2026). Hombres, mujeres, población vulnerable y viviendas excluyen ${f0(CP.pob_sin_desagregar)} habitantes de los centros poblados de los bloques 83 a 87, cuya fuente solo registra la población total.`),
  cap("Cuadro 5. Población de los centros poblados asociados, por distrito"),
  tPobDist,
  fuente("Fuente: ídem Cuadro 4. El distrito es el del centro poblado. Detalle por centro poblado, con coordenadas UTM, en el Anexo A3."),
  para(["Esta población corresponde a la de los centros poblados vinculados a los ecosistemas por recuperar (población directamente relacionada con la Unidad Productora), y no a la residente dentro de los polígonos. La estimación de la población beneficiaria indirecta aguas abajo y la proyección al horizonte de evaluación se desarrollarán en el análisis de demanda ", PD("población afectada y objetivo definitivas"), ". Quedan pendientes: centros poblados de los bloques ", CP.bloques_sin_cp.join(", "), "; coordenadas de ", CP.sin_coord, "; y el año censal de la capa INEI ", PD(""), "."]),
  h2("6.3 Avance del trabajo de campo social"),
  cap("Cuadro 6. Avance del diagnóstico social en los centros poblados asociados a bloques V6"),
  tSocial,
  fuente("Fuente: hoja de control del equipo social (CentrosPoblados-INEI-Bloques_v5, corte 24.09.2026). Detalle por centro poblado y bloque, con observaciones de campo, en el Anexo A3."),
  para([`El número de fichas F-DS-01 reportadas (${f0(SO.ds01)}) supera el tamaño de muestra calculado (${f0(SO.n)}). Las entrevistas a autoridades (F-DS-03) están completas en ${f0(SO.ds03_completo)} de ${f0(SO.registros)} registros y la ficha de saberes ancestrales está pendiente en todos los casos. Los centros poblados descartados se deben principalmente a su lejanía del bloque o a que comparten autoridad con otro caserío, y fueron reemplazados según las observaciones registradas.`]),
  figura(path.join(SAL, "IN_Piura_Mapa_Centros_Poblados_V6.png"), 640, 452),
  cap("Figura 2. Centros poblados asociados a los bloques de intervención V6"),

  h1("7. DIAGNÓSTICO DE LA UNIDAD PRODUCTORA (PUNTO 4)"),
  h2("7.1 Identificación"),
  para(["La Unidad Productora (UP) del proyecto no es una infraestructura convencional: ", b("es el ecosistema en su totalidad"), " —bosques estacionalmente secos de colina y montaña (Prosopis, Capparis, Loxopterygium, Bursera, Ceiba, entre otras), matorrales, pastizales y ecosistemas de transición andino-costeros, con sus suelos, cobertura vegetal, drenaje natural y red de quebradas—. Los bloques de intervención se ubican dentro del ecosistema. La UP provee los servicios ecosistémicos de regulación de riesgos naturales (movimientos en masa, erosión, inundaciones), regulación hídrica (interceptación, infiltración y recarga), control de erosión y conservación de suelos, y captura y almacenamiento de carbono."]),
  h2("7.2 Caracterización: ecosistemas en los bloques"),
  cap("Cuadro 7. Ecosistemas en los 117 bloques V6"),
  tEco,
  fuente(`Fuente: Mapa Nacional de Ecosistemas del Perú, MINAM 2018 (R.M. N.° 440-2018-MINAM); áreas prorrateadas al área oficial del catálogo V6. La capa recortada al ámbito cubre ${f3(D.eco_ambito.total_ha)} ha e incluye además: ${D.eco_ambito.otros.join(", ").toLowerCase()}. Detalle por bloque en el Anexo A4.`),
  para([`El bosque estacionalmente seco de colina y montaña ocupa el ${pct(D.ecosistemas[0].pct)} de los bloques. La «zona agrícola» (${pct(D.ecosistemas[1].pct)}) corresponde a superficie con uso agropecuario según el mapa regional, que se revisa en el tamizaje predial y en la verificación de campo (ficha F-DT-03), dado que la escala del mapa del MINAM es regional.`]),
  h2("7.3 Estado de la cobertura vegetal"),
  para([`Como indicador preliminar del estado de la UP se clasificó el índice de vegetación ajustado al suelo modificado (MSAVI 2024) en cinco clases. La superficie con MSAVI ≤ 0.4976 (clases DN 1 a 3, baja cobertura vegetal) suma `, b(`${f3(MS.bajo_umbral_ha)} ha (${pct(MS.pct_bajo)})`), ` y es mayoritaria en ${MS.bloques_mayoria_bajo} de los 117 bloques. La media de MSAVI por bloque varía entre ${MS.media_min.toFixed(4)} y ${MS.media_max.toFixed(4)}.`]),
  cap("Cuadro 8. Superficie de los bloques V6 por clase de MSAVI 2024"),
  tMsavi,
  fuente(`Fuente: MSAVI 2024 clasificado (exportación SIG AREAS_MSAVI_BLOQUES_V6) y traspaso a las plantillas por bloque; equivalencia DN→clase verificada en 117 bloques (r = ${MS.r}; error absoluto medio = ${MS.eam}). Detalle por bloque en el Anexo A4.`),
  para("La condición de degradación de cada bloque se confirma con la ficha F-DT-04 (causas e indicadores de degradación) y el procedimiento del MINAM para la identificación de áreas degradadas; el MSAVI no la sustituye."),
  h2("7.4 Procesos erosivos y peligros"),
  para([`Se caracterizaron en gabinete `, b(`${f0(CA.n)} cárcavas`), ` en ${f0(CA.bloques)} bloques, con una longitud total de ${f1(CA.long_m / 1000)} km, con atributos de altitud, pendiente, irregularidad y NDVI (Anexo A4 y capa 07 del paquete SIG). El inventario de campo se registra en la ficha F-DT-02.`]),
  para(["El análisis de peligros de la UP distingue el peligro por movimientos en masa (PMM), la erosión potencial hídrica (EPH) y el peligro generador de inundaciones (PGI), integrados por AHP en el peligro integrado (PI) de cinco niveles. Conforme al Anexo 2 de la Guía General, el análisis diferencia peligro, exposición, fragilidad y resiliencia, y se aplica tanto al ecosistema existente como a las medidas propuestas. Los resultados por microcuenca y bloque se remitirán con las capas de peligro ", PD(""), "."]),
  h2("7.5 Diagnóstico de campo"),
  para(["El diagnóstico de la UP en campo se realiza con las fichas F-DT-01 (datos generales y fisiografía), F-DT-02 (suelo y procesos erosivos, con inventario de cárcavas), F-DT-03 (ecosistema: composición, estructura y valor ecológico), F-DT-04 (causas e indicadores de degradación) y F-DT-05 (recursos hídricos y accesibilidad), complementadas con el inventario de elementos expuestos F-EE-01 a F-EE-07 y la verificación aérea con dron. El consolidado de campo por bloque se remitirá con la exportación del aplicativo IN Piura ", PD("número de bloques con F-DT y F-EE completas"), "."]),
  figura(path.join(RAIZ, "mapas", "salidas", "IN_Piura_01_ecosistemas_general.png"), 640, 452),
  cap("Figura 3. Ecosistemas y bloques de intervención V6 (Lámina 1, formato A3 en el paquete SIG)"),

  h1("8. DIAGNÓSTICO DE INVOLUCRADOS (PUNTO 5)"),
  para([`La identificación de involucrados se realiza con el mapeo de actores (F-DS-02: tipo, rol frente al proyecto, influencia, interés, posición y nivel territorial), las entrevistas a autoridades y líderes (F-DS-03; ${f0(SO.ds03)} reportadas), las actas de talleres participativos (F-DS-04), el inventario de conflictos socioambientales (F-DS-05), la evaluación de la disposición a participar y el consentimiento (F-DS-07) y el registro de prediantes del proceso de liberación de áreas (F-LA-01 a F-LA-06). El cuadro siguiente resume la matriz preliminar, elaborada en gabinete con el avance de campo; se validará con los registros F-DS-02 del aplicativo.`]),
  cap("Cuadro 9. Matriz preliminar de involucrados"),
  tInvol,
  fuente("Fuente: equipo formulador ANIN-DIME-SESDI, con base en el avance de campo F-DS. Matriz completa (actores identificados y posición preliminar) en el Anexo A5."),
  para(["Los acuerdos y compromisos formales (actas de talleres, actas de libre disponibilidad de predios y compromisos de operación y mantenimiento) se incorporarán a medida que se suscriban ", PD("relación de acuerdos suscritos"), "."]),

  h1("9. INFORMACIÓN GEORREFERENCIADA Y CARTOGRÁFICA (PUNTO 6)"),
  para(["Se remite el paquete ", b("IN_Piura_Paquete_SIG_Avance_Perfil.zip"), ", con cada capa en shapefile (con sus archivos .shp, .shx, .dbf, .prj y .cpg), KMZ (una capa por archivo y un KMZ integrado) y GeoPackage, además de las láminas A3 en PDF y un archivo LEEME con el diccionario de campos. Las capas SHP y GPKG están en UTM WGS 84 Zona 17S (EPSG:32717); los KMZ, en coordenadas geográficas WGS 84, formato propio de KML."]),
  cap("Cuadro 10. Inventario de información georreferenciada"),
  tInv,
  fuente("Fuente: paquete SIG del Proyecto IN Piura. Contenido, campos y fuentes de cada capa en el Anexo A6 y en el archivo LEEME."),

  h1("10. CONCLUSIONES"),
  bullet([`El ámbito de estudio comprende la cuenca alta del río Piura en 15 distritos de Ayabaca, Huancabamba y Morropón; el área de influencia la conforman ${BL.n} bloques V6 (${f3(BL.ha)} ha) y su área de influencia aprobada (${f3(BL.ai_ha)} ha), en ${BL.n_microcuencas} microcuencas.`]),
  bullet("Los bloques se seleccionan con un proceso de seis pasos: peligro integrado (MCA-AHP), digitalización, filtros de exclusión, idoneidad, tamizaje predial y verificación de campo."),
  bullet(`La población de los ${f0(CP.n)} centros poblados asociados a los bloques asciende a ${f0(CP.pob)} habitantes; el levantamiento socioeconómico superó la muestra calculada (${f0(SO.ds01)} fichas frente a ${f0(SO.n)}).`),
  bullet(`La UP es el ecosistema; el bosque estacionalmente seco de colina y montaña predomina en los bloques (${pct(D.ecosistemas[0].pct)}) y el ${pct(MS.pct_bajo)} de su superficie presenta baja cobertura vegetal según el MSAVI 2024.`),
  bullet("Se cuenta con una matriz preliminar de involucrados y con la información georreferenciada de ámbito, bloques, área de influencia, centros poblados, ecosistemas y cárcavas en formatos nativos."),

  h1("11. RECOMENDACIONES E INFORMACIÓN PENDIENTE"),
  para("Se recomienda remitir el presente informe y sus anexos a la dependencia solicitante, precisando que la información es preliminar y que se complementará con:"),
  bullet([PD("capas y resultados"), " del modelamiento de peligro (PMM, EPH, PGI y PI) por microcuenca y bloque, y límites de las 26 microcuencas priorizadas."]),
  bullet([PD("criterio"), " técnico de delimitación del área de influencia aprobada."]),
  bullet([PD("capas"), " de los filtros de exclusión e idoneidad, ráster MSAVI 2024 clasificado y ortomosaicos de dron."]),
  bullet([PD("consolidado"), " de las fichas F-DT, F-EE, F-DS-02 a F-DS-07 y F-LA exportado del aplicativo IN Piura."]),
  bullet([PD("población"), ` de los bloques ${CP.bloques_sin_cp.join(", ")}; coordenadas de los centros poblados de los bloques 83 a 87; población beneficiaria indirecta y proyección al horizonte del proyecto.`]),
  bullet([PD("acuerdos"), " suscritos con involucrados y matriz de involucrados definitiva."]),
  para("Asimismo, se recomienda coordinar con el Programa de Inversión del Río Piura la articulación de alcances, a fin de evitar duplicidad de intervenciones y aprovechar la complementariedad entre la infraestructura natural de la cuenca alta y las intervenciones aguas abajo."),

  h1("12. ANEXOS"),
  bullet([b("IN_Piura_Anexos_Avance_Perfil.xlsx: "), "A1 ámbito por distrito, microcuenca y bloque; A2 criterios de selección; A3 población por centro poblado y distrito, y avance social; A4 ecosistemas, MSAVI y cárcavas; A5 matriz de involucrados; A6 inventario SIG; notas y pendientes."]),
  bullet([b("IN_Piura_Paquete_SIG_Avance_Perfil.zip: "), "capas SHP, KMZ y GPKG, láminas A3 en PDF y LEEME."]),
  bullet([b("IN_Piura_Mapa_Centros_Poblados_V6.pdf: "), "lámina A3 de centros poblados asociados."]),
  new Paragraph({ spacing: { before: 400 }, children: [run(`Lima, ${fechaTxt}.`)] }),
];

const doc = new Document({
  ...A.estilos(),
  creator: "ANIN - DIME - SESDI",
  title: "Informe de avance del Perfil IN Piura",
  description: "Respuesta a la solicitud de información del Programa de Inversión del Río Piura",
  sections: [{ properties: A.pagina(), headers: { default: A.encabezado() }, footers: { default: A.pie() }, children: hijos }],
});

const salida = path.join(SAL, "IN_Piura_Informe_Avance_Perfil_Respuesta.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(salida, buf); console.log("→", salida); });
