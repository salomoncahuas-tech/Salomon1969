// Presentación de avance del Perfil IN Piura ante la SESDI · corte 03/11/2026
// Misma estructura que la presentación HTML (11 láminas), en formato PowerPoint nativo.
const path = require("path");
const fs = require("fs");
const A = require("/home/user/Salomon1969/.claude/skills/anin-in-piura/scripts/anin_pptx.js");
// Tema en azul institucional ANIN (#1B4F72), con la paleta dorado / cian de la presentación HTML.
A.THEME.name = "ANIN IN Piura · azul institucional";
Object.assign(A.THEME.colors, { dk1: "1B4F72", lt1: "FFFFFF", dk2: "133A55", lt2: "C4DCEB",
  accent1: "F2B447", accent2: "7FD3E6", accent3: "6CC58E", accent4: "2A6B96", accent5: "6FC3F2", accent6: "F28B7A",
  hlink: "7FD3E6", folHlink: "C4DCEB" });

const IMG = process.argv[2];           // carpeta con 00.jpg (portada) y 01..15.jpg (mapas)
const OUT = process.argv[3];

const { pres, C, tarjeta, txt, nombre } = A.nuevaPresentacion({ titulo: "Avance del Perfil IN Piura · SESDI · 03/11/2026" });
pres.subject = "Avance del Estudio de Perfil · CUI 2669244 · corte 03/11/2026";
const HEX = A.HEX;
const H = { panel: "133A55", line: "3E6A8A", muted: "C4DCEB", white: "FFFFFF", gold: "F2B447", teal: "7FD3E6", green: "6CC58E", idle: "8FA6B8", coral: "F28B7A", dark: "0F2B40" };
const PI = { b: "A6D96A", m: "FFD24C", a: "F46D43", ma: "B2182B" };

// ---------- utilidades de composición ----------
const SEC = (t) => pres.addSection({ title: t });
let secActual = "";
function lam(seccion, titulo, notas) {
  if (seccion !== secActual) { SEC(seccion); secActual = seccion; }
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: seccion });
  s.addText(titulo, { placeholder: "title" });
  if (notas) s.addNotes(notas);
  return s;
}
// Etiqueta pequeña tipo «eyebrow»
const kicker = (s, t, x, y, w, color = C.accent2) => txt(s, t.toUpperCase(), { x, y, w, h: 0.26, fontSize: 11, bold: true, color, charSpacing: 1 });
// Tarjeta de indicador
function kpi(s, x, y, w, h, label, value, unit, note, color = C.accent1) {
  tarjeta(s, x, y, w, h);
  txt(s, label.toUpperCase(), { x: x + 0.18, y: y + 0.14, w: w - 0.36, h: 0.3, fontSize: 10.5, bold: true, color: C.background2, charSpacing: 1 });
  txt(s, [{ text: value, options: { fontSize: 30, bold: true, color } }, { text: unit ? "  " + unit : "", options: { fontSize: 13, color: C.background2 } }],
    { x: x + 0.18, y: y + 0.45, w: w - 0.36, h: 0.62, valign: "middle" });
  if (note) txt(s, note, { x: x + 0.18, y: y + 1.1, w: w - 0.36, h: h - 1.2, fontSize: 12, color: C.background1 });
}
// Tarjeta con titular y texto
function card(s, x, y, w, h, k, titulo, cuerpo, kc = C.accent2, fs = 13) {
  tarjeta(s, x, y, w, h);
  s.addShape(pres.ShapeType.rect, { x, y, w, h: 0.06, fill: { color: kc }, line: { color: kc, width: 0 }, objectName: nombre("filete") });
  kicker(s, k, x + 0.2, y + 0.18, w - 0.4, kc);
  txt(s, titulo, { x: x + 0.2, y: y + 0.46, w: w - 0.4, h: 0.95, fontSize: 15, bold: true, color: C.background1 });
  txt(s, cuerpo, { x: x + 0.2, y: y + 1.45, w: w - 0.4, h: h - 1.55, fontSize: fs, color: C.background1, paraSpaceAfter: 4 });
}
// Tabla con cabecera
function tabla(s, cab, filas, o) {
  const fsz = o.fontSize || 12;
  const head = cab.map((c) => ({ text: c, options: { bold: true, color: C.text1, fill: { color: C.accent2 }, fontSize: fsz - 0.5, valign: "middle", align: o.alignHead || "left" } }));
  const body = filas.map((r, i) => r.map((c, j) => {
    const base = { color: C.background1, fill: { color: i % 2 ? C.text2 : C.text1 }, fontSize: fsz, valign: "top", align: (o.right || []).includes(j) ? "right" : "left" };
    return typeof c === "object" ? { text: c.text, options: { ...base, ...c.options } } : { text: String(c), options: base };
  }));
  s.addTable([head, ...body], { x: o.x, y: o.y, w: o.w, colW: o.colW, border: { type: "solid", pt: 0.5, color: H.line }, margin: 0.05, rowH: o.rowH, autoPage: false, objectName: nombre("tabla") });
}
const B = (t) => ({ text: t, options: { bold: true } });
const G = (t) => ({ text: t, options: { bold: true, color: C.accent1 } });
const fmt = (n, d = 2) => n.toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
const ejes = { catAxisLabelColor: H.muted, valAxisLabelColor: H.muted, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt", valGridLine: { color: H.line, size: 0.5 }, catGridLine: { style: "none" }, catAxisLineShow: false, valAxisLineShow: false };

// ======================= PORTADA =======================
SEC("Portada"); secActual = "Portada";
{
  const s = pres.addSlide({ sectionTitle: "Portada" });
  s.addImage({ path: path.join(IMG, "00.jpg"), x: 0, y: 0, w: 13.333, h: 7.5, altText: "Portada institucional de la Autoridad Nacional de Infraestructura (ANIN)", objectName: "portada-imagen" });
  const X = 13.333 * 0.634, W = 13.333 * (1 - 0.634 - 0.036);
  const az = "192E4D";
  txt(s, "Avance del Estudio de Preinversión del Proyecto IN Piura", { x: X, y: 0.85, w: W, h: 1.6, fontSize: 26, color: az });
  txt(s, [{ text: "CUI 2669244", options: { bold: true } }, { text: " · Estudio de Perfil", options: { breakLine: true } }, { text: "Presentación resumida ante la SESDI" }], { x: X, y: 2.55, w: W, h: 0.8, fontSize: 14, color: az });
  s.addShape(pres.ShapeType.rect, { x: X, y: 4.16, w: 0.43, h: 0.08, fill: { color: "FF0000" }, line: { color: "FF0000", width: 0 }, objectName: "portada-barra" });
  txt(s, "Dirección de Intervenciones Multisectoriales y de Emergencia – DIME\nSubdirección de Estudios de Inversión – SESDI", { x: X, y: 4.45, w: W, h: 1.0, fontSize: 15, color: az });
  txt(s, "Noviembre 2026 · corte al 03/11/2026", { x: X, y: 6.0, w: 11.65 - X, h: 0.35, fontSize: 14, color: az, align: "right" });
  s.addNotes("Avance del Estudio de Perfil del Proyecto IN Piura (CUI 2669244) ante la SESDI. Corte al 03/11/2026, con el Diagnóstico Integral IN Piura V3. Ejecución por Administración Directa; presupuesto total de la inversión S/ 372,442,809 (≈ USD 109.5 M a 3.40 S/ por USD).");
}

// ======================= 01 RESUMEN =======================
{
  const s = lam("01 Resumen", "Estado del Perfil: 48.5 % de avance ponderado",
    "Fuente: Diagnóstico Integral IN Piura V3 (03/11/2026). Avance ponderado con los pesos de los 15 entregables del Perfil. El 28/10 se presentó 45 %. Identificación (E1–E9): 71.9 %. El Peligro Integrado (MCA–AHP) combina movimientos en masa, erosión y escorrentía; el 93.9 % es solo por movimientos en masa (EVAR). La línea base del indicador propuesta es el criterio areal del E6.");
  const W = 2.93, g = 0.13, x0 = 0.6, y1 = 1.55, y2 = 4.25, h = 2.5;
  const k = [
    ["Avance del Perfil", "48.5", "%", "Ponderado de los 15 entregables · 45 % el 28/10", C.accent1],
    ["Capítulo de Identificación", "71.9", "%", "E1 a E9 · Geología, Suelos y Riesgo culminados", C.accent5],
    ["Bloques de intervención", "117", "", "12,157.47 ha netas (catálogo V8) · 15 distritos · 26 microcuencas", C.background1],
    ["Peligro Alto y Muy alto", "69.0", "%", "8,471.57 ha · Peligro Integrado; 93.9 % por movimientos en masa (EVAR)", C.accent1],
    ["Bajo umbral de cobertura", "68.0", "%", "8,341.83 ha · MSAVI areal (E6) · línea base del indicador", C.accent3],
    ["Diagnóstico social", "422", "fichas", "En 85 de 117 bloques · 18,651 hab. afectados", C.accent5],
    ["Libre disponibilidad", "0", "actas conformes", "22 predios inspeccionados en 10 bloques · 1 acta observada", C.accent1],
    ["Inversión + O&M 15 años", "95.8", "% del techo", "S/ 241.23 M + S/ 115.40 M · Matriz v1.6", C.accent3],
  ];
  k.forEach((d, i) => kpi(s, x0 + (i % 4) * (W + g), i < 4 ? y1 : y2, W, h, d[0], d[1], d[2], d[3], d[4]));
}
{
  const s = lam("01 Resumen", "Aspectos clave del avance",
    "Logros, decisiones y alertas de la etapa. Avance: E5 informe final 01/11/2026; EVAR final 19/10/2026; árbol de problemas v1.0 del 30/10/2026; E4 informe de avance (40 %). Decisión: adoptar la brecha areal del E6 y recalcularla sobre V8. Alerta: Matriz v1.6 dentro del techo (antes 167.7 % con el motor v1.4); el E5 solo sustenta 7.44 km de zanjas frente a 2,250.56 km del plano. Ruta crítica: Reporte_Liberacion_Areas al 03/11/2026.");
  const W = 2.93, g = 0.13;
  card(s, 0.6, 1.55, W, 5.25, "Avance", "Identificación al 71.9 %, con suelos y riesgo cerrados",
    "Geología, Suelos (informe final, 01/11) y Riesgo (EVAR final) culminados. El E7 pasa de 45 % a 80 % con el árbol de problemas v1.0. Hidrología entregó su informe de avance (E4 al 40 %) y el E10 sigue en 25 %, ahora con la Matriz v1.6.", C.accent3);
  card(s, 0.6 + (W + g), 1.55, W, 5.25, "Decisión", "Línea base del indicador sobre la base V8",
    "Adoptar el criterio areal del E6, 8,341.83 ha bajo el umbral MSAVI (68.0 %), como línea base del indicador de la R.M. N.° 00213-2024-MINAM, y recalcularlo sobre los polígonos V8, que descuentan 112.77 ha superpuestas con minería en trámite y un sitio arqueológico.", C.accent5);
  card(s, 0.6 + 2 * (W + g), 1.55, W, 5.25, "Alerta", "La cartera cabe en el techo, con condiciones",
    "Inversión (S/ 241.23 M) más O&M de 15 años (S/ 115.40 M) suman el 95.8 % de los S/ 372.44 M (antes 167.7 %). El margen exige validar en campo la terraza sin pirca y la infiltración: el plano proyecta 2,250.56 km de zanjas y el suelo del E5 solo sustenta 7.44 km.", C.accent1);
  card(s, 0.6 + 3 * (W + g), 1.55, W, 5.25, "Ruta crítica", "Primeras inspecciones prediales, sin actas conformes",
    "Al 03/11 el tamizaje registra 22 predios inspeccionados en 10 bloques de Morropón y 1 acta observada (bloque 85); ninguna conforme. Los otros 107 bloques y los 60 lotes SUS siguen sin datos de campo: sin actas no se fija la superficie disponible ni las metas.", C.accent1);
}
{
  const s = lam("01 Resumen", "Ámbito de intervención · catálogo V8",
    "Catálogo V8 (09/10/2026): parte de la relación V7 normalizada (12,270.235 ha, igual al catálogo V6) y descuenta 112.765 ha superpuestas con derechos mineros en trámite (bloques 8, 71, 22 y M17B1; 101.779 ha) y con el Sitio Arqueológico Juan Velasco Alvarado 4 (bloque 5; 10.986 ha). 26 microcuencas; las compartidas entre provincias se cuentan en cada una. El rango de intervención directa del proyecto es de 10,000 a 15,000 ha; la meta física se fija en Formulación.");
  tabla(s, ["Provincia", "Distritos", "Microcuencas", "Bloques", "V7 (ha)", "V8 neta (ha)", "% del ámbito"], [
    [B("Morropón"), "Buenos Aires, Chalaco, Chulucanas, Morropón, Salitral, San Juan de Bigote, Santa Catalina de Mossa, Santo Domingo, Yamango", "17", "59", "7,327.37", "7,313.55", "60.2"],
    [B("Huancabamba"), "Canchaque, Huancabamba, Huarmaca, Lalaquiz, San Miguel de El Faique", "11", "48", "3,551.16", "3,452.21", "28.4"],
    [B("Ayabaca"), "Frías", "4", "10", "1,391.70", "1,391.70", "11.4"],
    [B("Total"), B("15 distritos"), B("26"), G("117"), B("12,270.24"), G("12,157.47"), B("100.0")],
  ], { x: 0.6, y: 1.6, w: 12.13, colW: [1.6, 4.73, 1.2, 0.9, 1.25, 1.35, 1.1], right: [2, 3, 4, 5, 6], fontSize: 14 });
  txt(s, "Descuento V8: 101.78 ha de derechos mineros en trámite (bloques 8, 71, 22 y M17B1) y 10.99 ha del Sitio Arqueológico Juan Velasco Alvarado 4 (bloque 5). El bloque 8 conserva solo el 25.4 % de su área V7.",
    { x: 0.6, y: 5.6, w: 12.1, h: 0.9, fontSize: 14, color: C.background2 });
}

// ======================= 02 AVANCE =======================
const DEL = [
  ["E1 Microlocalización", 85, 85], ["E2 Diagnóstico social y liberación", 75, 75], ["E3 Geología", 100, 100], ["E4 Hidrología", 40, 60],
  ["E5 Suelos", 100, 63], ["E6 Ecosistemas y material vegetal", 39, 39], ["E7 Problema y planteamiento", 80, 45], ["E8 Análisis de riesgos", 100, 95],
  ["E9 Capítulo de Identificación", 40, 40], ["E10 Metas, costos y presupuesto", 25, 25], ["E11 Plan de gestión", 0, 0], ["E12 Plan de O&M", 0, 0],
  ["E13 Evaluación social", 0, 0], ["E14 Formulación y Evaluación", 0, 0], ["E15 Resumen ejecutivo", 0, 0],
];
{
  const s = lam("02 Avance", "Avance por entregable: 28/10 frente a 03/11",
    "Pesos del Perfil: E1 8, E2 10, E3 5, E4 6, E5 6, E6 8, E7 5, E8 8, E9 8, E10 10, E11 4, E12 4, E13 8, E14 5, E15 5. Cambios: E5 de 63 % a 100 % (informe final del 01/11/2026); E7 de 45 % a 80 % (árbol de problemas v1.0 como insumo principal del informe final); E8 de 95 % a 100 % (informe final EVAR). E4: 40 % según el informe de avance hidrológico; los 20 puntos que el 28/10 se anticipaban por el modelamiento se reconocerán con las lluvias de diseño, las curvas IDF y los caudales HEC-HMS.");
  const cats = DEL.map((d) => d[0]).reverse();
  s.addChart(pres.ChartType.bar, [
    { name: "28/10/2026", labels: cats, values: DEL.map((d) => d[2]).reverse() },
    { name: "03/11/2026", labels: cats, values: DEL.map((d) => d[1]).reverse() },
  ], { x: 0.6, y: 1.45, w: 8.4, h: 5.5, barDir: "bar", barGrouping: "clustered", barGapWidthPct: 45, chartColors: [H.idle, H.gold],
    valAxisMaxVal: 100, valAxisMinVal: 0, valAxisMajorUnit: 25, valAxisLabelFormatCode: "0\"%\"", catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 9, dataLabelColor: H.white, dataLabelFormatCode: "0\"%\";;;",
    showLegend: true, legendPos: "t", legendFontSize: 11, legendColor: H.white, ...ejes, objectName: "grafico-avance" });
  const caps = [["Capítulo I · Identificación", "71.9 %", "peso 64"], ["Capítulo II · Formulación", "13.9 %", "peso 18"], ["Capítulo III · Evaluación", "0.0 %", "peso 13"], ["Resumen y anexos", "0.0 %", "peso 5"]];
  tarjeta(s, 9.25, 1.5, 3.48, 5.4);
  kicker(s, "Avance ponderado", 9.45, 1.65, 3.1, C.accent2);
  txt(s, [{ text: "48.5", options: { fontSize: 44, bold: true, color: C.accent1 } }, { text: " %", options: { fontSize: 18, color: C.background2 } }], { x: 9.45, y: 1.95, w: 3.1, h: 0.85 });
  txt(s, "45 % el 28/10", { x: 9.45, y: 2.8, w: 3.1, h: 0.3, fontSize: 12, color: C.background2 });
  caps.forEach((c, i) => {
    const y = 3.3 + i * 0.88;
    txt(s, c[0], { x: 9.45, y, w: 3.1, h: 0.3, fontSize: 13, bold: true, color: C.background1 });
    txt(s, [{ text: c[1], options: { bold: true, color: C.accent5 } }, { text: "  ·  " + c[2], options: { color: C.background2 } }], { x: 9.45, y: y + 0.32, w: 3.1, h: 0.3, fontSize: 13 });
  });
}
{
  const s = lam("02 Avance", "Qué sustenta el avance al 03/11/2026",
    "Fuentes: Estudio de Suelos E5, informe final (GEOSIG, 01/11/2026); informe final EVAR (19/10/2026); Árboles de Problemas IN Piura v1.0 (30/10/2026); informe de avance del Estudio Hidrológico (láminas de octubre); Informe E2 B3 (03/11/2026); Matriz de Costos Paramétricos v1.6 (03/11/2026); E6, borrador del informe final (15/10/2026).");
  tabla(s, ["Entregable", "Avance", "Sustento"], [
    [B("E5 Suelos"), G("100 %  (63 %)"), "Informe final: 74 calicatas, 148 chequeos, 152 muestras y 13 pruebas de infiltración; 7 unidades de suelo, Factor K, CUM y matriz suelo–intervención."],
    [B("E8 Riesgo"), G("100 %  (95 %)"), "Informe final EVAR (Manual CENEPRED v02): peligro en 117 bloques y 85 áreas de aporte; vulnerabilidad y riesgo en 117 áreas de influencia."],
    [B("E7 Problema"), G("80 %  (45 %)"), "Árbol de problemas unificado v1.0: problema central, 3 causas directas y 10 indirectas; insumo principal del Informe Final."],
    [B("E4 Hidrología"), G("40 %  (60 %)"), "Informe de avance: 60 microcuencas y 264 nanocuencas; 15 nanocuencas priorizadas. Faltan lluvias de 24 h, curvas IDF y caudales HEC-HMS."],
    [B("E2 Social"), "75 %", "Volúmenes I–III en borrador B3: 422 fichas en 85 bloques. Tamizaje: 22 predios inspeccionados y 1 acta observada."],
    [B("E6 Ecosistemas"), "39 %", "Borrador del informe final (fases 1–5 de 7): 8,341.82 ha bajo el umbral MSAVI y 69 especies nativas."],
    [B("E10 Costos"), "25 %", "Matriz de Costos Paramétricos v1.6: inversión S/ 241.23 M y O&M S/ 115.40 M, el 95.8 % del techo."],
  ], { x: 0.6, y: 1.55, w: 12.13, colW: [1.9, 1.75, 8.48], fontSize: 13 });
  txt(s, "Entre paréntesis, el avance presentado el 28/10/2026.", { x: 0.6, y: 6.55, w: 12, h: 0.3, fontSize: 12, color: C.background2 });
}

// ======================= 03 TERRITORIO =======================
{
  const s = lam("03 Territorio", "Diagnóstico integral del territorio",
    "Volúmenes I–III del diagnóstico (versión 5; DT 21–23/09/2026, E6 de octubre de 2026), catálogo V8, E2 B3 y E5 final. Cada bloque se asigna a la clase de su índice PI medio (MCA–AHP; cortes 1.8 · 2.6 · 3.4 · 4.2); ningún bloque tiene clase Muy baja. Superficie de peligro sobre la geometría V7 hasta recortar los polígonos V8.");
  const k = [["Índice PI medio", "3.935", "", "Alto · MCA–AHP", C.accent1], ["Ecosistemas (E6)", "20", "bloques", "Moderados (14) o severos (6) · brecha 1,645.00 ha", C.accent3], ["Tierras de protección", "87.9", "%", "Del área estudiada · E5 final", C.accent3], ["Población afectada", "18,651", "hab.", "154 centros poblados · E2 al 02/11/2026", C.background1]];
  k.forEach((d, i) => kpi(s, 0.6 + i * 3.06, 1.5, 2.95, 1.75, d[0], d[1], d[2], d[3], d[4]));
  const provs = ["Morropón", "Huancabamba", "Ayabaca"];
  const ha = [7327.37, 3551.15, 1391.70];
  const cls = [[15.84, 1247.02, 2572.05, 3492.46], [131.05, 934.11, 1036.99, 1449.00], [146.24, 0, 1063.56, 181.90]];
  const pct = (j) => cls.map((c, i) => Math.round(1000 * c[j] / ha[i]) / 10);
  s.addChart(pres.ChartType.bar, ["Bajo", "Medio", "Alto", "Muy alto"].map((n, j) => ({ name: n, labels: provs, values: pct(j) })),
    { x: 0.6, y: 3.45, w: 6.6, h: 3.45, barDir: "bar", barGrouping: "percentStacked", chartColors: [PI.b, PI.m, PI.a, PI.ma], barGapWidthPct: 40,
      showValue: true, dataLabelPosition: "ctr", dataLabelFontSize: 10, dataLabelColor: "1A1A1A", dataLabelFormatCode: "[<4]\"\";0\"%\"",
      showTitle: true, title: "Peligro integrado por provincia (% de superficie)", titleColor: H.white, titleFontSize: 13,
      showLegend: true, legendPos: "b", legendColor: H.white, legendFontSize: 11, catAxisLabelFontSize: 12, valAxisHidden: true, ...ejes, objectName: "grafico-pi" });
  tarjeta(s, 7.4, 3.45, 5.33, 3.45);
  kicker(s, "Lo que muestra el diagnóstico", 7.6, 3.6, 5, C.accent2);
  txt(s, [
    { text: "69.0 % del ámbito con peligro integrado Alto o Muy alto: Morropón 72.8 %, Huancabamba 66.3 %, Ayabaca 56.2 %.", options: { bullet: true, breakLine: true } },
    { text: "8,341.83 ha bajo el umbral MSAVI: 5,503.41 ha de restauración activa y 2,838.41 ha de enriquecimiento.", options: { bullet: true, breakLine: true } },
    { text: "Suelos: infiltración lenta en 5 de 7 unidades; restringe las zanjas.", options: { bullet: true, breakLine: true } },
    { text: "Clima semiárido (544 mm/año, aridez 0.34) y 55 cárcavas (23.08 km).", options: { bullet: true } },
  ], { x: 7.6, y: 3.95, w: 4.95, h: 2.85, fontSize: 13, color: C.background1, paraSpaceAfter: 5 });
}
{
  const s = lam("03 Territorio", "Indicadores por provincia",
    "Medias ponderadas por la superficie del bloque. Superficie neta del catálogo V8 (09/10/2026); población afectada del E2 B3 (corte del aplicativo 02/11/2026); calicatas del informe final del E5 (01/11/2026). Regímenes territoriales: Morropón A Bosque seco del valle (37 bloques) y B Mosaico andino (22); Huancabamba A Bosque seco del piso bajo (8) y B Mosaico andino de ladera (40); Ayabaca A Bosque seco del piedemonte (4), B Mosaico agropecuario de ladera (4) y C Matorral y bosque relicto de cabecera (2).");
  tabla(s, ["Indicador", "Morropón", "Huancabamba", "Ayabaca", "Ámbito"], [
    ["Bloques · microcuencas", "59 · 17", "48 · 11", "10 · 4", G("117 · 26")],
    ["Superficie neta V8 (ha)", "7,313.55", "3,452.21", "1,391.70", G("12,157.47")],
    ["Peligro integrado Alto y Muy alto (ha · %)", "5,335.20 · 72.8", "2,353.73 · 66.3", "782.64 · 56.2", G("8,471.57 · 69.0")],
    ["Índice PI medio", "4.020", "3.898", "3.580", G("3.935")],
    ["Bajo umbral MSAVI, criterio areal E6 (ha)", "6,216.63", "1,599.99", "525.21", G("8,341.83")],
    ["Bloques moderados o severos (E6)", "14", "6", "0", G("20")],
    ["Precipitación (mm/año) · índice de aridez", "543 · 0.31", "541 · 0.38", "557 · 0.39", G("544 · 0.34")],
    ["Cárcavas registradas (n · km)", "39 · 18.02", "9 · 3.90", "7 · 1.15", G("55 · 23.08")],
    ["Calicatas E5 · bloques con calicata propia", "37 · 21", "28 · 15", "9 · 4", G("74 · 40")],
    ["Población afectada 2017 · centros poblados", "9,537 · 79", "6,573 · 58", "2,541 · 17", G("18,651 · 154")],
  ], { x: 0.6, y: 1.5, w: 12.13, colW: [4.53, 1.9, 1.9, 1.9, 1.9], right: [1, 2, 3, 4], fontSize: 13 });
}

// ======================= 04 MAPAS DISTRITALES =======================
const V8 = { "5": 284.740, "22": 76.574, "M17B1": 42.377, "8": 32.267, "71": 6.407 };
const MAPS = JSON.parse(fs.readFileSync(path.join(__dirname, "maps.json"), "utf8"));
MAPS.forEach((m, j) => {
  m.codes.forEach((c) => { if (V8[c[0]] != null) { c[3] = c[2]; c[2] = V8[c[0]]; } });
  const ha8 = m.codes.reduce((a, c) => a + c[2], 0);
  const s = lam("04 Mapas distritales", `Distrito de ${m.n} · ${m.p}`,
    `Mapa de distribución de bloques preliminares del distrito de ${m.n} (equipo SIG del Proyecto IN Piura, QGIS, UTM WGS 84 Zona 17S; delimitación V6). Superficie neta del catálogo V8 (09/10/2026).` +
    (m.codes.some((c) => c[3] != null) ? " Bloques con descuento V8: " + m.codes.filter((c) => c[3] != null).map((c) => `${c[0]} (V7 ${fmt(c[3])} ha → V8 ${fmt(c[2])} ha)`).join("; ") + "." : ""));
  s.addImage({ path: path.join(IMG, String(j + 1).padStart(2, "0") + ".jpg"), x: 0.6, y: 1.45, w: 7.85, h: 5.55, altText: `Mapa de distribución de bloques en el distrito de ${m.n}, provincia de ${m.p}: ${m.bl} bloques, ${fmt(ha8)} ha netas V8.`, objectName: nombre("mapa") });
  tarjeta(s, 8.65, 1.45, 4.08, 5.55);
  const st = [["Bloques", String(m.bl)], ["Superficie V8 neta", fmt(ha8) + " ha"], ["Del total V8", (100 * ha8 / 12157.47).toFixed(1) + " %"], ["Microcuencas", String(m.mc)], ["MSAVI promedio", m.ms.toFixed(3)]];
  st.forEach((d, i) => {
    txt(s, d[0].toUpperCase(), { x: 8.85, y: 1.6 + i * 0.6, w: 1.9, h: 0.5, fontSize: 10.5, bold: true, color: C.background2, charSpacing: 1, valign: "middle" });
    txt(s, d[1], { x: 10.6, y: 1.6 + i * 0.6, w: 1.95, h: 0.5, fontSize: 17, bold: true, color: i === 0 ? C.accent1 : C.background1, align: "right", valign: "middle" });
  });
  kicker(s, `Códigos · ${m.cod} definitivos M#B#`, 8.85, 4.7, 3.7, C.accent2);
  const runs = [];
  m.codes.forEach((c, i) => {
    runs.push({ text: (c[1] ? c[0] : "N.° " + c[0]), options: { bold: true, color: c[3] != null ? C.accent6 : c[1] ? C.accent1 : C.background1 } });
    runs.push({ text: ` ${fmt(c[2])}` + (i < m.codes.length - 1 ? "   " : ""), options: { color: C.background2 } });
  });
  txt(s, runs, { x: 8.85, y: 5.0, w: 3.7, h: 1.9, fontSize: m.codes.length > 12 ? 10 : 11.5 });
});

// ======================= 05 ESTUDIOS BÁSICOS =======================
{
  const s = lam("05 Estudios básicos", "Estudios básicos: dos culminados en la etapa",
    "E3 Geología: ISL-MM medio 3.81 e ISL-EH 3.68. E4: QSWAT sobre DEM ASTER de 30 m (71,787 ha); 15 nanocuencas priorizadas (5,936 ha; 9/5/1 por provincia); siguen lluvias máximas de 24 h (PISCO y 6 estaciones, 1964–2023), curvas IDF y HEC-HMS (SCS-CN). E5: GEOSIG, 01/11/2026; laboratorio LQA Valle Grande; Soil Taxonomy y WRB 2022; CUM D.S. 005-2022-MIDAGRI. E8: Manual CENEPRED v02.");
  const st = [
    ["E3 · Culminado", "Geología", "Unidades litológicas y susceptibilidad a movimientos en masa y erosión hídrica, integradas en los volúmenes.", C.accent3],
    ["E4 · 40 %", "Hidrología y clima", "60 microcuencas y 264 nanocuencas delimitadas; 15 priorizadas. Siguen lluvias de diseño, curvas IDF y caudales.", C.accent2],
    ["E5 · Culminado", "Suelos · informe final", "74 calicatas y 152 muestras; 7 unidades de suelo; 87.9 % tierras de protección; zanja compatible solo en 2 unidades.", C.accent3],
    ["E6 · 39 %", "Ecosistemas", "Borrador final de 117 bloques: 8,341.82 ha bajo el umbral; 69 especies nativas; faltan viveros y validación.", C.accent2],
    ["E8 · Culminado", "Análisis de riesgos", "EVAR final: peligro alto y muy alto en el 93.9 % de los bloques; riesgo alto en 7,132.39 ha.", C.accent3],
    ["E1 · 85 %", "Microlocalización · V8", "117 bloques y 12,157.47 ha netas; 45 con código M#-B# y 72 por codificar; recortar 5 polígonos.", C.accent2],
  ];
  st.forEach((d, i) => card(s, 0.6 + (i % 3) * 4.09, i < 3 ? 1.5 : 4.25, 3.95, 2.6, d[0], d[1], d[2], d[3], 13));
}
{
  const s = lam("05 Estudios básicos", "Suelos y ecosistemas: lo que cambia para el diseño",
    "E5: 13 pruebas de doble anillo (Kostiakov); velocidad de diseño = mitad de la básica; una zanja de 40 cm debe vaciarse en 72 h. Canchaque y El Ingenio: 6.6 cm/h (6 h); otras cinco unidades: 0.165–0.415 cm/h (96–242 h). Área apta para zanjas (suelo, pendiente 10–40 % y cobertura): 12.30 ha (7.44 km). E6: índice de condición (Gann et al., 2019): 53 bloques no degradados, 44 leves, 14 moderados y 6 severos.");
  kpi(s, 0.6, 1.5, 2.95, 1.85, "Tierras de protección", "87.9", "%", "12,074.90 ha (X) · pastos 7.7 % · forestal 4.3 %", C.accent3);
  kpi(s, 3.66, 1.5, 2.95, 1.85, "Zanjas aptas por suelo", "7.44", "km", "frente a 2,250.56 km del plano de obras", C.accent1);
  kpi(s, 6.72, 1.5, 2.95, 1.85, "Bajo umbral MSAVI", "8,341.82", "ha", "E6: 5,503.41 activa · 2,838.41 enriquecimiento", C.accent3);
  kpi(s, 9.78, 1.5, 2.95, 1.85, "Oferta de plantones", "21,000", "", "declarada (Hualtacal) · brecha en la fase 6", C.accent1);
  tabla(s, ["Unidad de suelo", "Superficie (ha)", "Factor K", "Infiltración diseño (cm/h)", "Vaciado zanja 40 cm (h)", "Zanja de infiltración"], [
    ["Morropón (Torriorthents)", "6,936.10", "0.0272 · Media", "0.175", "229", "No compatible"],
    ["Cruz de Piedra (Ustorthents)", "2,658.80", "0.0222 · Media", "0.165", "242", "No compatible"],
    ["Chillique (Haplocambids)", "2,111.10", "0.0235 · Media", "0.415", "96", "No compatible"],
    ["Ramón Castilla (Haplustepts)", "763.30", "0.0261 · Media", "0.165", "242", "No compatible"],
    ["Canchaque (Haplustolls)", "560.40", "0.0146 · Baja", "6.625", "6", G("Con restricción")],
    ["Huarmaca (Dystrustepts)", "513.10", "0.0142 · Baja", "0.165", "242", "No compatible"],
    ["El Ingenio (Humustepts)", "193.80", "0.0078 · Muy baja", "6.625", "6", G("Con restricción")],
  ], { x: 0.6, y: 3.55, w: 12.13, colW: [3.13, 1.6, 1.9, 2.1, 1.9, 1.5], right: [1, 3, 4], fontSize: 12 });
}

// ======================= 06 RIESGO (E8) =======================
{
  const s = lam("06 Riesgo (E8)", "Riesgo por movimientos en masa (EVAR final)",
    "Informe EVAR (19/10/2026, Manual CENEPRED v02). Peligro: capa PEL_MM (239,462 polígonos). Riesgo R = P × V con cortes 0.04375, 0.111 y 0.231. Riesgo muy alto solo en AI_73 (Chococa, 36.98 ha) y AI_41 (Cambruran, 19.94 ha), en Huarmaca. 106 centros poblados de referencia, 21,672 hab. sin duplicación (Censo 2017). Por provincia, peligro alto y muy alto: Morropón 90.7 %, Huancabamba 98.2 %, Ayabaca 99.4 % de los bloques. El EVAR no valora daños: beneficios por daños evitados [POR DEFINIR].");
  const labs = ["Peligro · bloques (12,270 ha)", "Peligro · áreas de aporte (7,538 ha)", "Peligro · áreas de influencia (17,132 ha)", "Riesgo · áreas de influencia (17,132 ha)"].reverse();
  const D = [[0.08, 6.04, 45.16, 48.72], [0, 3.12, 37.28, 59.60], [4.37, 45.69, 32.65, 17.29], [0, 58.04, 41.63, 0.33]].reverse();
  s.addChart(pres.ChartType.bar, ["Bajo", "Medio", "Alto", "Muy alto"].map((n, j) => ({ name: n, labels: labs, values: D.map((r) => r[j]) })),
    { x: 0.6, y: 1.45, w: 7.4, h: 5.45, barDir: "bar", barGrouping: "percentStacked", chartColors: [PI.b, PI.m, PI.a, PI.ma], barGapWidthPct: 45,
      showValue: true, dataLabelPosition: "ctr", dataLabelFontSize: 10, dataLabelColor: "1A1A1A", dataLabelFormatCode: "[<2.5]\"\";0.0", showLegend: true, legendPos: "b", legendColor: H.white, legendFontSize: 11,
      catAxisLabelFontSize: 11, valAxisHidden: true, ...ejes, objectName: "grafico-riesgo" });
  const k = [["Peligro alto y muy alto", "93.9", "%", "De los bloques · 96.9 % de las áreas de aporte", C.accent1], ["Riesgo alto", "7,132.4", "ha", "41.6 % de las áreas de influencia; predomina en 76", C.accent1], ["Riesgo muy alto", "56.9", "ha", "AI_73 Chococa y AI_41 Cambruran (Huarmaca)", C.accent1], ["Vulnerabilidad alta", "15", "de 117", "3 Morropón · 10 Huancabamba · 2 Ayabaca", C.accent5]];
  k.forEach((d, i) => kpi(s, 8.2, 1.45 + i * 1.39, 4.53, 1.3, d[0], d[1], d[2], null, d[4]));
  k.forEach((d, i) => txt(s, d[3], { x: 10.35, y: 1.55 + i * 1.39, w: 2.25, h: 1.1, fontSize: 11.5, color: C.background1, valign: "middle" }));
}
{
  const s = lam("06 Riesgo (E8)", "Control del riesgo: medidas y líneas del proyecto",
    "Medidas propuestas por el EVAR y su correspondencia con las líneas de intervención (MRR-CCC). La protección específica de elementos expuestos es una medida externa a la UP y requiere evaluación de rentabilidad propia según la Guía DGPMI. El EVAR usa un desencadenante fijo, sin escenarios de cambio climático: falta el escenario SSP5-8.5.");
  tabla(s, ["Medida propuesta (E8)", "Condición para su selección", "Línea de intervención", "Gestión"], [
    [B("Restauración de cobertura con especies adecuadas"), "Suelo, clima, uso y tenencia compatibles; diagnóstico de erosión y supervivencia", "Línea 1 · infraestructura natural (verde)", "Correctiva"],
    [B("Protección de suelos y control de erosión superficial"), "Sectores con erosión identificada y diseños que no concentren escorrentía", "Línea 2 · zanjas, terrazas y diques", "Correctiva"],
    [B("Manejo de drenaje superficial"), "Descarga segura; no derivar caudales hacia laderas inestables", "Línea 2.1 · drenaje", "Correctiva"],
    [B("Recuperación de fajas y cauces degradados"), "Compatibilidad con la dinámica fluvial y las fajas marginales", "Línea 2.1 · quebradas (con la ANA)", "Correctiva"],
    [B("Protección de elementos expuestos"), "Mecanismo, intensidad y diseño sustentados para el activo", "Medida externa a la UP", "Correctiva y reactiva"],
    [B("Medidas no estructurales"), "Alerta temprana por precipitación acumulada; rutas de evacuación", "Línea 3 · gobernanza", "Prospectiva y reactiva"],
  ], { x: 0.6, y: 1.55, w: 12.13, colW: [3.5, 4.33, 2.8, 1.5], fontSize: 13 });
}

// ======================= 07 SOCIAL Y LIBERACIÓN =======================
{
  const s = lam("07 Social y liberación", "Diagnóstico social: 422 fichas en 85 bloques",
    "E2 B3, Volúmenes I–III (03/11/2026) y Matrices_E2_IN_Piura_B3.xlsx (corte del aplicativo 02/11/2026 23:50). Fichas: F-DS-01 215, F-DS-02 103, F-DS-03 74, F-DS-06 30 (27 localidades de Morropón). Sin aplicar: F-DS-04, F-DS-05 y F-DS-07. Población: centros poblados asociados a los bloques (Censo 2017 más localidades sin código INEI); proyección 2026: 18,524 hab.");
  const k = [["Avance del E2", "75", "%", "Volúmenes I–III en borrador B3", C.accent1], ["Bloques con ficha", "85", "de 117", "422 fichas F-DS al 02/11/2026", C.accent5], ["Aceptación", "67", "de 76", "Autoridades de acuerdo · 8 condicionan · 0 en desacuerdo", C.accent3], ["Actores mapeados", "113", "", "108 cooperantes · 5 neutrales o sin posición", C.accent3]];
  k.forEach((d, i) => kpi(s, 0.6 + i * 3.06, 1.5, 2.95, 1.95, d[0], d[1], d[2], d[3], d[4]));
  tabla(s, ["Provincia", "Bloques", "Con ficha", "Fichas F-DS", "Cooperantes", "Población 2017 · CP", "Instrumentos"], [
    [B("Morropón"), "59", "53", "307", "71 de 76", "9,537 · 79", "4 de 7"],
    [B("Huancabamba"), "48", "26", "96", "29 de 29", "6,573 · 58", "3 de 7"],
    [B("Ayabaca"), "10", "6", "19", "8 de 8", "2,541 · 17", "2 de 7"],
    [B("Total"), G("117"), G("85"), G("422"), G("108 de 113"), G("18,651 · 154"), "—"],
  ], { x: 0.6, y: 3.65, w: 12.13, colW: [2.2, 1.3, 1.4, 1.6, 1.8, 2.33, 1.5], right: [1, 2, 3, 4, 5, 6], fontSize: 14 });
  txt(s, "Vacíos: 32 bloques sin ficha (22 en Huancabamba, Huarmaca concentra el vacío); F-DS-05 y F-DS-07 sin aplicar; Ayabaca sin F-DS-03.", { x: 0.6, y: 6.2, w: 12.1, h: 0.6, fontSize: 13, color: C.background2 });
}
{
  const s = lam("07 Social y liberación", "Liberación de áreas: primeras inspecciones",
    "Reporte_Liberacion_Areas_IN_Piura_20261103.xlsx, F-LA-03 y F-LA-04 al 03/11/2026. 177 unidades activas (117 bloques y 60 lotes SUS, 12,527.21 ha). Inspecciones 07/10–29/10/2026: M6B10, M19B2, M18B3, M2B1, M7B1, 85, M3B7, 49, 52 y 55. Seis inspecciones observadas por interferencia con bosque natural (M18B3, M2B1, M6B10) y una con vivienda habitada (52). Tablero del subentregable: 27.5 % (planificación e instrumentos al 100 %; campo y gabinete por consolidar).");
  const P = [["Paso 1", "Socialización", "F-LA-01", "0 de 117 bloques"], ["Paso 2", "Registro del titular", "F-LA-02", "0 de 117 bloques"], ["Paso 3", "Inspección", "F-LA-03", "22 predios en 10 bloques"], ["Paso 4", "Acuerdo (acta)", "F-LA-04", "1 acta observada"], ["Paso 5", "Verificación registral", "F-LA-05", "0 de 117 bloques"], ["Paso 6", "Expediente", "AX-04", "0 de 117 bloques"]];
  P.forEach((p, i) => {
    const x = 0.6 + i * 2.04, act = i === 2 || i === 3;
    tarjeta(s, x, 1.5, 1.94, 2.3, act ? C.accent4 : C.text2);
    kicker(s, p[0], x + 0.15, 1.62, 1.7, C.accent1);
    txt(s, p[1], { x: x + 0.15, y: 1.92, w: 1.68, h: 0.65, fontSize: 14, bold: true, color: C.background1 });
    txt(s, p[2], { x: x + 0.15, y: 2.62, w: 1.68, h: 0.3, fontSize: 11, color: C.background2 });
    txt(s, p[3], { x: x + 0.15, y: 2.95, w: 1.68, h: 0.75, fontSize: 13, bold: true, color: act ? C.accent1 : C.background1 });
  });
  kpi(s, 0.6, 4.05, 2.95, 2.85, "Actas conformes", "0", "", "Ninguna unidad en estado LA-4 al 03/11/2026", C.accent1);
  tarjeta(s, 3.66, 4.05, 9.07, 2.85);
  kicker(s, "Alertas del cruce con el diagnóstico social y el campo", 3.86, 4.18, 8.6, C.accent2);
  txt(s, [
    { text: "107 bloques y 60 lotes SUS (256.97 ha) siguen sin datos de campo; sin actas no se fija la superficie disponible.", options: { bullet: true, breakLine: true } },
    { text: "19 bloques con tenencia declarada distinta del gabinete (16 en Morropón); 5 bloques de Huancabamba (4, 9, 35, 38 y 43) con más de 19 predios.", options: { bullet: true, breakLine: true } },
    { text: "Seis inspecciones interfieren con bosque natural (solo manejo o enriquecimiento) y el acta del bloque 85 está observada por falta de plazo.", options: { bullet: true, breakLine: true } },
    { text: "Polígono comunal sin nombre (ID 2790, 10 bloques): identificar con la DRA Piura y SUNARP antes de las asambleas.", options: { bullet: true } },
  ], { x: 3.86, y: 4.5, w: 8.7, h: 2.35, fontSize: 13, color: C.background1, paraSpaceAfter: 4 });
}
{
  const s = lam("07 Social y liberación", "Situación predial de gabinete por provincia",
    "Capa SIG del 09/10/2026 (comunidades campesinas GEORURAL–MIDAGRI, predios rurales COFOPRI y superficie sin registro) sobre la superficie neta V8. Total: comunal 3,929.61 ha (32.3 %); Estado o posesionarios 7,928.97 ha (65.2 %); COFOPRI 298.88 ha y 328 predios. No reemplaza la verificación de campo ni las actas.");
  const provs = ["Morropón", "Huancabamba", "Ayabaca"];
  s.addChart(pres.ChartType.bar, [
    { name: "Comunidad campesina", labels: provs, values: [28.0, 20.8, 83.9] },
    { name: "Estado y/o posesionarios", labels: provs, values: [71.8, 71.0, 16.1] },
    { name: "Predios COFOPRI", labels: provs, values: [0.2, 8.2, 0] },
  ], { x: 0.6, y: 1.45, w: 6.2, h: 5.45, barDir: "bar", barGrouping: "percentStacked", chartColors: [H.green, H.gold, H.coral], barGapWidthPct: 40,
    showValue: true, dataLabelPosition: "ctr", dataLabelFontSize: 11, dataLabelColor: "1A1A1A", dataLabelFormatCode: "[<2.5]\"\";0.0", showLegend: true, legendPos: "b", legendColor: H.white, legendFontSize: 11,
    catAxisLabelFontSize: 12, valAxisHidden: true, ...ejes, objectName: "grafico-predial" });
  tabla(s, ["Provincia", "Comunal (ha)", "COFOPRI (ha · predios)", "> 19 predios", "Sin ficha social"], [
    [B("Morropón"), "2,044.75", "15.71 · 21", "0", "6"],
    [B("Huancabamba"), "717.84", "283.17 · 307", "5", "22"],
    [B("Ayabaca"), "1,167.02", "0.00 · 0", "0", "4"],
    [B("Total"), G("3,929.61"), G("298.88 · 328"), G("5"), G("32")],
  ], { x: 7.0, y: 1.6, w: 5.73, colW: [1.45, 1.2, 1.38, 0.8, 0.9], right: [1, 2, 3, 4], fontSize: 12 });
  txt(s, "La tenencia declarada en la F-DS-01 difiere de la clase SIG en 19 de 76 bloques con dato (16 de 50 en Morropón, 1 de 20 en Huancabamba y 2 de 6 en Ayabaca).", { x: 7.0, y: 4.4, w: 5.73, h: 1.2, fontSize: 13, color: C.background1 });
}

// ======================= 08 PROBLEMA Y ALTERNATIVAS =======================
{
  const s = lam("08 Problema y alternativas", "Árbol de problemas unificado v1.0",
    "Árboles de Problemas IN Piura v1.0 (30/10/2026), insumo principal del Informe Final del E7 (80 %). Sigue el modelo ANIN (ZZ-000005, Fig. 2). EF: incremento de los costos de reconstrucción y disminución de la resiliencia frente a movimientos en masa e inundaciones intensificados por el cambio climático. Causas indirectas: CD1 sobrepastoreo, tala, expansión agrícola, baja regeneración; CD2 trochas sin drenaje, laderas sin conservación y cárcavas, pisoteo; CD3 organización comunal, gobernanza institucional, tenencia sin sanear. Factores exógenos: pendiente, litología, El Niño y cambio climático.");
  const box = (x, y, w, h, k, t, fill, kc) => {
    tarjeta(s, x, y, w, h, fill);
    kicker(s, k, x + 0.15, y + 0.1, w - 0.3, kc);
    txt(s, t, { x: x + 0.15, y: y + 0.38, w: w - 0.3, h: h - 0.45, fontSize: 12.5, bold: true, color: C.background1 });
  };
  box(0.6, 1.45, 12.13, 0.85, "Efecto final", "Más costos de reconstrucción y menor resiliencia de la población y sus medios de vida frente a movimientos en masa e inundaciones intensificados por el cambio climático", C.text2, C.accent6);
  const ed = ["Mayor ocurrencia de movimientos en masa y huaicos en laderas y quebradas de cabecera", "Erosión hídrica acelerada y pérdida de suelo (laminar, surcos y cárcavas)", "Mayor escorrentía torrencial y transporte de sedimentos que activan las quebradas"];
  ed.forEach((t, i) => box(0.6 + i * 4.09, 2.42, 3.95, 1.0, "ED" + (i + 1), t, C.text2, C.accent1));
  tarjeta(s, 0.6, 3.55, 12.13, 1.15, C.accent4);
  kicker(s, "Problema central", 0.75, 3.62, 4, C.accent1);
  txt(s, "La población de la cuenca alta del río Piura no accede adecuadamente al servicio ecosistémico de regulación de riesgos naturales por la degradación de los ecosistemas en las zonas de origen de movimientos en masa e inundaciones",
    { x: 0.75, y: 3.9, w: 11.85, h: 0.75, fontSize: 14, bold: true, color: C.background1 });
  const cd = [["CD1 · cobertura", "Degradación de la cobertura vegetal protectora en las zonas de origen"], ["CD2 · suelos y laderas", "Reducida capacidad de los suelos y laderas para controlar la erosión y la escorrentía"], ["CD3 · gobernanza", "Inadecuada gobernanza y gestión local del ecosistema y del riesgo"]];
  cd.forEach((d, i) => box(0.6 + i * 4.09, 4.83, 3.95, 1.0, d[0], d[1], C.text2, C.accent2));
  txt(s, "Línea base del indicador (R.M. N.° 00213-2024-MINAM): 8,341.83 ha bajo el umbral (68.0 %). Población afectada: 18,651 hab. en 154 centros poblados. Superficie: 12,157.47 ha netas V8.",
    { x: 0.6, y: 6.0, w: 12.13, h: 0.8, fontSize: 13, color: C.background2 });
}
{
  const s = lam("08 Problema y alternativas", "Objetivo central y medios fundamentales",
    "Redacción preliminar del objetivo, alineada al problema central v1.0; el árbol de medios y fines debe ajustarse en el Informe Final del E7. La meta en hectáreas se fija en Formulación (E10), después del tamizaje predial.");
  tarjeta(s, 0.6, 1.45, 12.13, 1.05, C.accent4);
  txt(s, "La población de la cuenca alta del río Piura accede adecuadamente al servicio ecosistémico de regulación de riesgos naturales, con ecosistemas recuperados en las zonas de origen de movimientos en masa e inundaciones",
    { x: 0.8, y: 1.55, w: 11.75, h: 0.85, fontSize: 14, bold: true, color: C.background1, valign: "middle" });
  tabla(s, ["Código", "Línea", "Causa", "Medio fundamental", "Relación"], [
    [B("MF-1"), "IN verde", "CD1", "Reforestación con especies nativas en laderas degradadas", "Excluyente con MF-2"],
    [B("MF-2"), "IN verde", "CD1", "Revegetación y regeneración natural asistida, con viveros central y temporales", "Excluyente con MF-1"],
    [B("MF-3"), "IN verde", "CD1", "Lotes de sistemas de uso sostenible agroforestales y silvopastoriles", "Complementario"],
    [B("MF-4"), "IN marrón", "CD2", "Zanjas de infiltración y terrazas de formación lenta, según el suelo (E5)", "Técnicas excluyentes"],
    [B("MF-5"), "IN marrón", "CD2", "Diques para control de cárcavas y estabilización de taludes de trochas", "Complementario"],
    [B("MF-6"), "Gobernanza", "CD3", "Comités de gestión, planes de manejo comunitario y acuerdos de pastoreo", "Imprescindible"],
    [B("MF-7"), "Gobernanza", "CD3", "Capacidades, educación ambiental y alerta temprana comunitaria", "Imprescindible"],
    [B("MF-8"), "Gobernanza", "CD3", "Libre disponibilidad predial, saneamiento de la tenencia y MERESE hídrico", "Imprescindible"],
  ], { x: 0.6, y: 2.65, w: 12.13, colW: [0.95, 1.45, 0.85, 6.68, 2.2], fontSize: 12.5 });
}
{
  const s = lam("08 Problema y alternativas", "Alternativas de solución",
    "Las alternativas se pueden combinar por sector o régimen territorial. La selección se hará en Formulación con los costos de inversión y O&M (Matriz v1.6, E12 y E13), la compatibilidad suelo–intervención del E5 (Anexo 7) y el análisis de riesgo de las medidas (E8). Común a ambas: lotes de uso sostenible, dimensionamiento según el inventario de cárcavas y el E4, y actas de libre disponibilidad.");
  tabla(s, ["Línea", "Alternativa 1 · restauración activa con obras", "Alternativa 2 · restauración asistida con prácticas"], [
    [B("IN verde"), "Reforestación con plantación de nativas en macizos y fajas en contorno, con vivero central y temporales", "Revegetación de áreas críticas, enriquecimiento y regeneración natural asistida bajo clausura temporal"],
    [B("IN marrón"), "Zanjas de infiltración y diques de mampostería o gaviones en cárcavas", "Terrazas de formación lenta con barreras vivas, diques de piedra acomodada y estabilización biológica"],
    [B("Gobernanza"), "Comités de gestión, planes de manejo, capacitación y alerta temprana", "Lo mismo, más acuerdos de pastoreo y clausura y un MERESE hídrico con las JASS"],
    [B("Efecto · costo"), "Efecto inmediato; inversión y O&M mayores", "Efecto gradual; inversión y O&M menores si se cumplen los acuerdos"],
    [B("Exposición"), "Mayor: obras expuestas a huaicos y lluvias El Niño", "Menor; mayor tiempo de establecimiento"],
    [B("Suelo (E5)"), "Zanjas solo compatibles, con restricción, en Canchaque y El Ingenio", "Terrazas compatibles con restricción (30–70 %); reforestación y nucleación en todas las unidades"],
  ], { x: 0.6, y: 1.55, w: 12.13, colW: [1.9, 5.11, 5.12], fontSize: 13 });
}

// ======================= 09 METAS Y COSTOS (E10) =======================
{
  const s = lam("09 Metas y costos (E10)", "Matriz de Costos Paramétricos v1.6",
    "Costos_Parametros_IN_Piura_v1.6.xlsx (03/11/2026), generado desde el v1.5. Metas marrón del plano de obras de conservación de suelos (06/10/2026). Elimina el pircado con piedra y los riegos de revegetación de terrazas; desestima relleno compactado, eliminación de desmonte, podas y fertilización. Cifras preliminares, no son metas del Perfil: densidades y cerco por calibrar con el piloto (21, 31, 66, 73 y M6B10); cárcavas sin metrado; supervisión y expediente técnico 0 % por definir.");
  const k = [["Plantación R1 + R2", "8,341.87", "ha", "R1 plena 5,503.44 · R2 enriquecimiento 2,838.42", C.accent3], ["Plantas", "4.51", "M", "4,959,516 plantones con recalce de 10 %", C.accent3], ["Zanjas", "2,250.6", "km", "y 3,117.7 km de terrazas · plano de obras", C.accent1], ["Cerco perimétrico", "577.8", "km", "100 % del perímetro (supuesto)", C.background1],
    ["Inversión total", "S/ 241.23", "M", "S/ 19,660 por ha · v1.4: S/ 343.05 M", C.accent1], ["O&M 15 años", "S/ 115.40", "M", "Primer año S/ 19.45 M · v1.4: S/ 281.68 M", C.accent5], ["VAC social", "S/ 130.56", "M", "TSD 8 % · horizonte de 20 años", C.background1], ["Costo-efectividad", "S/ 15,652", "", "VAC social por ha plantada", C.accent5]];
  k.forEach((d, i) => kpi(s, 0.6 + (i % 4) * 3.06, i < 4 ? 1.5 : 4.25, 2.95, 2.5, d[0], d[1], d[2], d[3], d[4]));
}
{
  const s = lam("09 Metas y costos (E10)", "La cartera v1.6 cabe en el techo presupuestal",
    "Techo: presupuesto total de la inversión S/ 372,442,809 (incluye preinversión, expediente técnico, ejecución y O&M de 15 años). Matriz v1.6: S/ 356.63 M (95.8 %), margen S/ 15.81 M para estudios. Motor v1.4 (28/10): S/ 624.73 M (167.7 %). Escenarios de la Matriz v1.6 frente al techo: fichas ANIN completas 163.7 %; solo lo resaltado en amarillo 140.9 %; v1.6 con diques en 23.08 km de cárcavas 105.8 %; v1.6 con zanja de la ficha ANIN (0.40 × 0.30 m) 87.3 %; v1.6 sin obras en bloques con P < 400 mm 92.5 %.");
  const cats = ["Presupuesto total", "Matriz v1.6 (03/11)", "Motor v1.4 (28/10)"];
  s.addChart(pres.ChartType.bar, [
    { name: "Inversión", labels: cats, values: [372.44, 241.23, 343.05] },
    { name: "O&M 15 años", labels: cats, values: [0, 115.40, 281.68] },
  ], { x: 0.6, y: 1.45, w: 7.2, h: 5.45, barDir: "bar", barGrouping: "stacked", chartColors: [H.gold, H.teal], barGapWidthPct: 45,
    showValue: true, dataLabelPosition: "ctr", dataLabelFontSize: 11, dataLabelColor: H.dark, dataLabelFormatCode: "#,##0.0;;;",
    showTitle: true, title: "Millones de soles", titleColor: H.white, titleFontSize: 12, valAxisMinVal: 0, valAxisMaxVal: 650, valAxisMajorUnit: 100,
    showLegend: true, legendPos: "b", legendColor: H.white, legendFontSize: 11, catAxisLabelFontSize: 12, valAxisLabelFontSize: 10, ...ejes, objectName: "grafico-techo" });
  kpi(s, 8.0, 1.5, 4.73, 1.95, "Inversión + O&M v1.6", "95.8", "% del techo", "S/ 356.63 M · margen de S/ 15.81 M para estudios (antes 167.7 %)", C.accent3);
  tabla(s, ["Escenario de la Matriz v1.6", "% del techo"], [
    ["Valores por defecto (sin pirca ni relleno)", G("95.8 %")],
    ["Con la zanja de la ficha ANIN 0.40 × 0.30 m", "87.3 %"],
    ["Sin obras en bloques con P < 400 mm", "92.5 %"],
    ["Con diques en cárcavas (23.08 km)", "105.8 %"],
    ["Con las fichas ANIN completas", "163.7 %"],
  ], { x: 8.0, y: 3.65, w: 4.73, colW: [3.43, 1.3], right: [1], fontSize: 12 });
}
{
  const s = lam("09 Metas y costos (E10)", "Presupuesto por partida (precios de mercado)",
    "CU de mercado set-2026: base ANIN (S10, 10/12/2020), jornal rural S/ 100 por día y factor IPC Lima dic-2020 → set-2026. Plantones con recalce de 10 %. Flete rural con la metodología Sierra Azul. Gastos generales 15.08 % (S/ 31,606,286). Precios sociales: MO no calificada × 0.42, MO calificada × 0.79, bienes y servicios ÷ 1.18 (MEF, Anexo 11). Inversión a precios sociales: S/ 134,396,789 (55.7 %).");
  const R = (l, u, m, p, pc) => [l, u, m, p, pc];
  tabla(s, ["Partida", "Und", "Metrado", "Parcial (S/)", "% CD"], [
    R("Producción de plantones en vivero propio (recalce 10 %)", "und", "4,959,516", "6,612,951", "3.2"),
    R("Plantación y labores por planta", "und", "4,508,651", "30,933,320", "14.8"),
    R("Labores por hectárea", "ha", "8,342", "9,345,023", "4.5"),
    R("Flete rural y terrestre de plantones e insumos", "kg", "9,125,509", "2,268,047", "1.1"),
    R("Cerco perimétrico de alambre de púas", "m", "577,846", "12,695,154", "6.1"),
    R("Zanjas de infiltración 0.60 × 0.50 m y de derivación", "km", "2,250.56", "47,367,900", "22.6"),
    R("Terrazas de formación lenta sin pirca", "km", "3,117.65", "96,700,179", "46.1"),
    R("Barrera viva en bordos y camellones (sin riego)", "km", "5,069.97", "3,705,006", "1.8"),
    R("Control de cárcavas · meta por definir", "km", "sin metrado", "—", "—"),
    [B("Costo directo"), "", "", G("209,627,579"), B("100.0")],
    [B("Inversión total (con gastos generales)"), "", "", G("241,233,865"), ""],
  ], { x: 0.6, y: 1.5, w: 12.13, colW: [6.13, 0.8, 1.7, 2.1, 1.4], right: [2, 3, 4], fontSize: 12.5 });
}
{
  const s = lam("09 Metas y costos (E10)", "Inversión por tipología y piloto de calibración",
    "Tipologías T1–T6 = grupo de ecosistema dominante × degradación (Alta si ≥ 50 % del bloque está bajo el umbral MSAVI 0.4976). Simulación de la Matriz v1.6 por tipología con los valores por defecto. Piloto de calibración: bloques 21 (T6), 66 (T6), 73 (T6), 31 (T3) y M6B10 (T1): 538.32 ha, R1 + R2 310.45 ha, inversión S/ 8,271,185 (S/ 15,365 por ha); zanjas 97.50 km y terrazas 83.71 km.");
  const T = ["T1 BES alta (47)", "T2 BES moderada (7)", "T3 Andino alta (3)", "T4 Andino moderada (20)", "T5 Mosaico alta (3)", "T6 Mosaico moderada (37)"];
  s.addChart(pres.ChartType.bar, [{ name: "Inversión (S/ M)", labels: T.slice().reverse(), values: [184.87, 6.79, 1.86, 17.84, 1.83, 28.04].reverse() }],
    { x: 0.6, y: 1.45, w: 6.6, h: 5.45, barDir: "bar", chartColors: [H.gold], barGapWidthPct: 40, showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelColor: H.white, dataLabelFormatCode: "#,##0.00",
      showTitle: true, title: "Inversión por tipología · millones de soles (bloques)", titleColor: H.white, titleFontSize: 12, showLegend: false, catAxisLabelFontSize: 11, valAxisHidden: true, ...ejes, objectName: "grafico-tipologia" });
  tabla(s, ["Bloque piloto", "Tip.", "Área (ha)", "R1+R2 (ha)", "Inversión (S/)", "S/ por ha"], [
    [B("21"), "T6", "84.23", "4.36", "1,268,578", "15,061"],
    [B("66"), "T6", "102.34", "9.05", "696,398", "6,805"],
    [B("73"), "T6", "54.29", "10.28", "185,732", "3,421"],
    [B("31"), "T3", "47.58", "37.31", "658,297", "13,835"],
    [B("M6B10"), "T1", "249.89", "249.46", "5,462,180", "21,858"],
    [B("Piloto"), "5", G("538.32"), G("310.45"), G("8,271,185"), G("15,365")],
  ], { x: 7.4, y: 1.55, w: 5.33, colW: [1.05, 0.55, 0.85, 0.9, 1.13, 0.85], right: [2, 3, 4, 5], fontSize: 11.5 });
  txt(s, "T1 (bosque seco con degradación alta) concentra el 76.6 % de la inversión. El piloto cuesta S/ 15,365 por ha, frente a S/ 19,660 de la cartera.", { x: 7.4, y: 4.6, w: 5.33, h: 1.2, fontSize: 13, color: C.background1 });
}

// ======================= 10 PENDIENTES =======================
const GAP = [
  ["Libre disponibilidad (E2 · SE2)", "22 predios inspeccionados; 0 actas conformes; 107 bloques y 60 lotes SUS sin datos", "Ruta F-LA en bloques piloto y de mayor riesgo; resolver 7 observaciones", "Alta"],
  ["Diagnóstico social (E2 · SE1)", "32 bloques sin ficha; F-DS-04, 05 y 07 sin aplicar; Ayabaca sin F-DS-03", "Priorizar Huarmaca y Ayabaca; F-DS-05 y F-DS-07 en bloques comunales y 83–87", "Alta"],
  ["Línea base de la brecha", "Criterios mixtos; el E6 areal cubre 117 bloques sobre geometría V7", "Adoptar el E6 areal (8,341.83 ha) y recalcularlo en V8", "Alta"],
  ["Zanjas del plano frente al E5", "2,250.56 km proyectados; 7.44 km aptos por suelo, pendiente y cobertura", "Pruebas de infiltración; terraza o barrera viva donde no infiltre", "Alta"],
  ["Costos y evaluación social (E10)", "95.8 % del techo; terraza sin pirca por validar; beneficios por determinar", "Calibrar con el piloto; supervisión y expediente; VANS y TIR social", "Alta"],
  ["Bloques con P < 400 mm", "8, 20, 71, M4B3 y M4B4: infiltración y prendimiento inciertos", "Ensayos de infiltración y prendimiento; decidir las obras", "Media"],
  ["E4 Hidrología", "Sin lluvias de diseño, curvas IDF ni caudales", "Completar HEC-HMS; vuelo de dron para drenaje y cárcavas", "Media"],
  ["E6 Ecosistemas", "Fases 1–5; oferta declarada de solo 21,000 plantones", "Fase 6 (viveros y brecha de plantones) y fase 7 (validación)", "Media"],
  ["Geometría V8", "Polígonos de 5 bloques sin recortar; bloque 8 al 25.4 %", "Recortar y recalcular PI, MSAVI, EVAR y costos; decidir el bloque 8", "Media"],
  ["E8 · EVAR", "Sin escenarios de cambio climático; daños no valorados", "Escenario SSP5-8.5; inventario de elementos expuestos", "Media"],
  ["Cárcavas", "No figuran en el plano; 85 bloques sin registro", "Levantamiento F-DT-02 o dron antes de fijar diques", "Media"],
  ["E5 Suelos", "3 unidades sin infiltración propia; Factor K sin parcelas", "Repetir infiltración donde haya zanjas; validar el Factor K", "Media"],
  ["E7 Árbol de problemas", "v1.0 lista; medios y fines por alinear", "Redactar el Informe Final del E7", "Baja"],
  ["E1 Relieve", "Pendiente de campo por clase visual", "Medir con clinómetro en parcelas", "Baja"],
];
[[0, 5, "Pendientes de prioridad alta"], [5, 14, "Pendientes de prioridad media y baja"]].forEach(([a, b, t]) => {
  const s = lam("10 Pendientes", t, "Vacíos de integración en los 117 bloques al 03/11/2026 (Diagnóstico Integral V3, «Vacíos para cerrar el Perfil»). Salen de la lista el laboratorio de suelos (E5 final) y la cartera sobre el techo (Matriz v1.6). La prioridad indica el efecto sobre el indicador de brecha, las metas físicas o el costo.");
  tabla(s, ["Tema", "Situación al 03/11/2026", "Acción propuesta", "Prioridad"],
    GAP.slice(a, b).map((r) => [B(r[0]), r[1], r[2], r[3] === "Alta" ? G(r[3]) : r[3]]),
    { x: 0.6, y: 1.5, w: 12.13, colW: [2.75, 4.35, 4.03, 1.0], fontSize: b - a > 6 ? 12 : 13.5 });
});

// ======================= 11 PRÓXIMOS PASOS =======================
{
  const s = lam("11 Próximos pasos", "Proyección del avance del Perfil",
    "Avance ponderado con los pesos del Perfil. 28/10/2026: 45 % (presentación anterior). 03/11/2026: 48.5 %. Diciembre de 2026: ≥ 66.5 % al culminar el Capítulo de Identificación (E1 a E9 = 64 puntos, más el E10 al 25 %). Febrero de 2027: ≥ 95 % con los Capítulos de Formulación y Evaluación (E10 a E14 = 31 puntos). Marzo de 2027: 100 % con el resumen ejecutivo y anexos. Fechas referenciales, sujetas al campo y a la temporada de lluvias.");
  const cats = ["28/10/2026", "03/11/2026", "Dic 2026", "Feb 2027", "Mar 2027"];
  s.addChart(pres.ChartType.bar, [
    { name: "Logrado", labels: cats, values: [45.35, 48.52, 0, 0, 0] },
    { name: "Previsto (mínimo)", labels: cats, values: [0, 0, 66.5, 95, 100] },
  ], { x: 0.6, y: 1.45, w: 7.6, h: 5.45, barDir: "col", barGrouping: "stacked", chartColors: [H.gold, H.teal], barGapWidthPct: 55,
    valAxisMaxVal: 100, valAxisMinVal: 0, valAxisMajorUnit: 25, valAxisLabelFormatCode: "0\"%\"", showValue: true, dataLabelPosition: "inEnd", dataLabelFontSize: 12, dataLabelColor: H.dark, dataLabelFormatCode: "0.0\"%\";;;",
    showLegend: true, legendPos: "b", legendColor: H.white, legendFontSize: 11, catAxisLabelFontSize: 12, valAxisLabelFontSize: 10, ...ejes, objectName: "grafico-proyeccion" });
  kpi(s, 8.4, 1.5, 4.33, 1.7, "Avance al 03/11/2026", "48.5", "%", "E5 y E8 culminados · E7 al 80 %", C.accent1);
  kpi(s, 8.4, 3.35, 4.33, 1.7, "Al cerrar la Identificación", "≥ 66.5", "%", "Diciembre de 2026 · E1 a E9", C.accent5);
  kpi(s, 8.4, 5.2, 4.33, 1.7, "Perfil completo", "100", "%", "Marzo de 2027 · viabilidad en abril", C.accent3);
}
{
  const s = lam("11 Próximos pasos", "Próximos pasos: noviembre de 2026 a abril de 2027",
    "Cumplido en octubre: plano de obras (06/10), catálogo V8 (09/10), primeras inspecciones F-LA-03 (desde 07/10), borrador final del E6 (15/10), EVAR final (19/10), árbol de problemas v1.0 (30/10). Cumplido en noviembre: E5 final (01/11), E2 B3 y Matriz v1.6 (03/11). Fechas referenciales del estudio, sujetas a las condiciones de campo y a la temporada de lluvias.");
  const T = [
    ["Nov 2026", "En curso", ["Informe Final del E7 con medios y fines alineados", "Hidrología: lluvias de diseño, IDF y HEC-HMS", "Tamizaje: asambleas, fichas y actas (piloto y mayor riesgo)", "Infiltración donde hay zanjas; F-DS-05 y F-DS-07", "Inicio del análisis de mercado y del plan de O&M"]],
    ["Dic 2026", "Programado", ["Liberación de áreas y catálogo definitivo", "E6 fases 6 y 7", "Polígonos V8 recortados", "Capítulo de Identificación"]],
    ["Ene – Feb 2027", "Programado", ["Metas y costos calibrados con el piloto", "Plan de O&M y evaluación social", "Capítulos de Formulación y Evaluación"]],
    ["Mar – Abr 2027", "Programado", ["Perfil completo con resumen ejecutivo y anexos", "Revisión, registro en el Banco de Inversiones", "Declaración de viabilidad"]],
  ];
  T.forEach((r, i) => {
    const x = 0.6 + i * 3.06, w = 2.95;
    tarjeta(s, x, 1.5, w, 5.35, i === 0 ? C.accent4 : C.text2);
    txt(s, r[0], { x: x + 0.2, y: 1.65, w: w - 0.4, h: 0.45, fontSize: 18, bold: true, color: C.accent1 });
    kicker(s, r[1], x + 0.2, 2.12, w - 0.4, i === 0 ? C.accent5 : C.background2);
    txt(s, r[2].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < r[2].length - 1 } })), { x: x + 0.2, y: 2.5, w: w - 0.4, h: 4.2, fontSize: 13.5, color: C.background1, paraSpaceAfter: 6 });
  });
}

A.guardar(pres, OUT).then(() => console.log("ok", OUT));
