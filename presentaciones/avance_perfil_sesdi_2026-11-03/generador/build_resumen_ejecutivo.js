// Resumen ejecutivo para la Alta Dirección · Proyecto IN Piura · corte 03/11/2026
// Formato tomado del Resumen Ejecutivo del 14-SEP-2026 (A4, encabezado y pie ANIN, verde #1B4D2E, Arial).
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer, AlignmentType, WidthType,
  BorderStyle, ShadingType, LevelFormat, PageNumber, VerticalAlign,
} = require("docx");

const OUT = process.argv[2];
const VERDE = "1B4D2E", CLARO = "E8F0EA", GRIS = "595959", OSC = "1F1F1F", CORAL = "B03A2E", ORO = "8A5209";
const W = 11906 - 2 * 794; // ancho útil con los márgenes del modelo (DXA)
const L = AlignmentType.LEFT, R = AlignmentType.RIGHT, C = AlignmentType.CENTER, J = AlignmentType.JUSTIFIED;

const run = (t, o = {}) => new TextRun({ text: t, font: "Arial", size: 19, color: OSC, ...o });
const B = (t, o = {}) => run(t, { bold: true, ...o });
const R_ = (parts) => (Array.isArray(parts) ? parts : [parts]).map((x) => (typeof x === "string" ? run(x) : x));
const para = (parts, o = {}) => new Paragraph({ children: R_(parts), alignment: J, spacing: { after: 100, line: 264 }, ...o });
const h = (t) => new Paragraph({ children: [run(t, { bold: true, size: 21, color: VERDE })], spacing: { before: 180, after: 80 }, keepNext: true,
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: VERDE, space: 2 } } });
const bul = (parts) => new Paragraph({ numbering: { reference: "vi", level: 0 }, children: R_(parts), alignment: J, spacing: { after: 60, line: 259 } });
const nota = (t) => new Paragraph({ children: [run(t, { italics: true, size: 15, color: GRIS })], spacing: { before: 40, after: 80 } });

const fino = { style: BorderStyle.SINGLE, size: 4, color: "A6A6A6" };
const bordes = { top: fino, bottom: fino, left: fino, right: fino };
function celda(contenido, w, { head = false, fill, align = L, bold = false, size = 17, color } = {}) {
  const ch = Array.isArray(contenido) ? contenido : [contenido];
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, borders: bordes, verticalAlign: VerticalAlign.CENTER,
    shading: head ? { type: ShadingType.CLEAR, fill: VERDE, color: "auto" } : fill ? { type: ShadingType.CLEAR, fill, color: "auto" } : undefined,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: [new Paragraph({ alignment: head ? C : align, children: ch.map((x) => (typeof x === "string"
      ? run(x, { size, bold: head || bold, color: head ? "FFFFFF" : color || OSC }) : x)) })],
  });
}
function tabla(cols, cab, filas, total) {
  const rows = [new TableRow({ tableHeader: true, cantSplit: true, children: cab.map((t, i) => celda(t, cols[i].w, { head: true })) })];
  filas.forEach((f, k) => rows.push(new TableRow({ cantSplit: true, children: f.map((v, i) => celda(v, cols[i].w, { fill: k % 2 ? CLARO : undefined, align: cols[i].a || L, color: cols[i].c ? cols[i].c(v) : undefined })) })));
  if (total) rows.push(new TableRow({ cantSplit: true, children: total.map((v, i) => celda(v, cols[i].w, { fill: "D9E6DD", bold: true, align: cols[i].a || L })) }));
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: cols.map((c) => c.w), rows });
}
// Tablero de cuatro cifras
function tablero(items) {
  const w = W / items.length;
  const cell = ([k, v, n, col]) => new TableCell({
    width: { size: w, type: WidthType.DXA }, borders: bordes, shading: { type: ShadingType.CLEAR, fill: CLARO, color: "auto" },
    margins: { top: 70, bottom: 70, left: 110, right: 110 },
    children: [
      new Paragraph({ alignment: C, children: [run(k.toUpperCase(), { size: 14, bold: true, color: GRIS })] }),
      new Paragraph({ alignment: C, spacing: { before: 20, after: 20 }, children: [run(v, { size: 32, bold: true, color: col || VERDE })] }),
      new Paragraph({ alignment: C, children: [run(n, { size: 15, color: GRIS })] }),
    ],
  });
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: items.map(() => w), rows: [new TableRow({ cantSplit: true, children: items.map(cell) })] });
}

// ---------- Encabezado y pie (como el modelo) ----------
const header = new Header({ children: [
  new Paragraph({ alignment: C, spacing: { after: 0 }, children: [run("AUTORIDAD NACIONAL DE INFRAESTRUCTURA — ANIN", { bold: true, size: 17, color: VERDE })] }),
  new Paragraph({ alignment: C, spacing: { after: 0 }, children: [run("DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA — DIME · SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN — SESDI", { bold: true, size: 13, color: VERDE })] }),
  new Paragraph({ alignment: C, spacing: { after: 0 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: VERDE, space: 4 } },
    children: [run("Proyecto IN Piura | CUI 2669244 | Estudio de Preinversión a nivel de Perfil | Corte al 03 de noviembre de 2026", { size: 13, color: GRIS })] }),
] });
const footer = new Footer({ children: [new Paragraph({ alignment: C, border: { top: { style: BorderStyle.SINGLE, size: 4, color: "A6A6A6", space: 4 } },
  children: [run("Resumen Ejecutivo dirigido a la Alta Dirección · ANIN–DIME–SESDI · Página ", { size: 13, color: GRIS }),
    new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 13, color: GRIS }), run(" de ", { size: 13, color: GRIS }),
    new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 13, color: GRIS })] })] });

// ---------- Contenido ----------
const est = (v) => /Culminado/.test(v) ? VERDE : /actualizar|retraso/i.test(v) ? CORAL : /27\/11/.test(v) ? ORO : undefined;
const children = [
  new Paragraph({ alignment: C, spacing: { after: 40 }, children: [run("RESUMEN EJECUTIVO — ESTADO DE AVANCE DEL ESTUDIO DE PREINVERSIÓN (PERFIL)", { bold: true, size: 24, color: VERDE })] }),
  new Paragraph({ alignment: C, spacing: { after: 60 }, children: [run("«Recuperación del servicio de regulación de riesgos naturales y recuperación de ecosistemas degradados en la Cuenca Alta del Río Piura»", { italics: true, size: 18, color: OSC })] }),
  new Paragraph({ alignment: C, spacing: { after: 140 }, children: [run("CUI 2669244 | Inversión referencial: S/ 372,442,809 | Administración Directa | 3 provincias · 15 distritos · 117 bloques · 12,157.47 ha netas | Corte: 03-NOV-2026 (corte previo: 14-SEP-2026)", { size: 15, color: GRIS })] }),

  h("1. SITUACIÓN GENERAL"),
  para(["Al 03 de noviembre de 2026 el Perfil alcanza ", B("51.4 % de avance ponderado"), " (45 % al 28/10) y el Capítulo de Identificación ", B("76.4 %"), ". Se culminaron la Geología (E3), los Suelos (E5) y el Análisis de Riesgos (E8), este último ", B("preliminar"), ", pendiente de actualizar con la redelimitación del área de influencia directa (AID). Dos frentes requieren decisión: la ", B("cartera preliminar supera el presupuesto total en 5.8 %"), " y la ", B("liberación de áreas aún no tiene actas conformes"), "."]),
  tablero([
    ["Avance del Perfil", "51.4 %", "15 entregables ponderados · 45 % al 28/10"],
    ["Capítulo de Identificación", "76.4 %", "E3, E5 y E8 culminados · cierre en dic-2026"],
    ["Inversión + O&M 15 años", "105.8 %", "del presupuesto total · faltan S/ 21.68 M", CORAL],
    ["Actas de libre disponibilidad", "0", "conformes · 22 predios inspeccionados", CORAL],
  ]),

  h("2. AVANCE POR ENTREGABLE"),
  tabla([{ w: 4300 }, { w: 800, a: C }, { w: 1150, a: C }, { w: 1150, a: C }, { w: W - 7400, c: est }],
    ["ENTREGABLE", "PESO", "14-SEP", "03-NOV", "ESTADO"], [
      ["E1 · Microlocalización (catálogo V8)", "8", "—", "85 %", "En curso"],
      ["E2 · Población afectada, actores y liberación de áreas", "10", "30 % (1)", "75 %", "En curso: volúmenes I–III en borrador B3"],
      ["E3 · Estudio de Geología", "5", "100 %", "100 %", "Culminado"],
      ["E4 · Estudio de Hidrología", "6", "20 %", "40 %", "En curso: faltan lluvias de diseño, IDF y caudales"],
      ["E5 · Estudio de Suelos", "6", "40 %", "100 %", "Culminado (informe final 01/11)"],
      ["E6 · Ecosistemas y oferta de material vegetal", "8", "10 %", "75 %", "Versión final prevista al 27/11"],
      ["E7 · Problema, causas y planteamiento", "5", "—", "80 %", "En curso: árbol de problemas v1.0"],
      ["E8 · Análisis de riesgos (EVAR)", "8", "55 %", "100 %", "Preliminar: por actualizar con el AID"],
      ["E9 · Informe del Capítulo de Identificación", "8", "—", "40 %", "En curso"],
      ["E10 · Metas, costos y presupuesto", "10", "—", "25 %", "En curso: Matriz de Costos v1.6"],
      ["E11 a E15 · Formulación, evaluación y resumen", "26", "—", "0 %", "Programado (dic-2026 a feb-2027)"],
    ], ["AVANCE PONDERADO DEL PERFIL", "100", "—", "51.4 %", ""]),
  nota("(1) El 14-SEP se reportó el diagnóstico de la población afectada (30 %) y, por separado, el expediente de liberación de áreas (10 %; 27.5 % en el tablero al 17/10). «—»: no reportado en el resumen del 14-SEP."),

  h("3. RESULTADOS CLAVE"),
  bul([B("Territorio y brecha. "), "117 bloques y 12,157.47 ha netas (catálogo V8, que descuenta 112.77 ha superpuestas con minería en trámite y un sitio arqueológico); 69.0 % con peligro integrado alto o muy alto; línea base del indicador de brecha: 8,341.83 ha bajo el umbral de cobertura vegetal (68.0 %)."]),
  bul([B("Riesgo (preliminar). "), "Peligro por movimientos en masa alto y muy alto en el 93.9 % de los bloques; riesgo alto en 7,132.39 ha y muy alto en 56.92 ha, este último en dos áreas de influencia de Huarmaca."]),
  bul([B("Suelos. "), "El 87.9 % del área estudiada es tierra de protección y la infiltración es lenta en cinco de siete unidades de suelo, lo que limita el uso de zanjas de infiltración."]),
  bul([B("Social. "), "422 fichas en 85 de 117 bloques; 18,651 habitantes afectados en 154 centros poblados; 67 de 76 autoridades están de acuerdo con el proyecto, 8 lo condicionan y ninguna se opone."]),
  bul([B("Costos (preliminares). "), "Con las obras del plano de conservación de suelos y 23.08 km de diques en 55 cárcavas, la inversión es de S/ 259.85 M y la O&M de 15 años de S/ 134.27 M: S/ 394.12 M, el 105.8 % del presupuesto total."]),

  h("4. ALERTAS Y RUTA CRÍTICA"),
  bul([B("Liberación de áreas: "), "22 predios inspeccionados en 10 bloques y 1 acta observada (bloque 85), pendiente de conformidad; 107 bloques y 60 lotes de uso sostenible sin datos de campo. Sin actas no se pueden fijar las metas físicas."]),
  bul([B("Presupuesto: "), "la cartera vuelve al techo con la zanja de la ficha ANIN de 0.40 × 0.30 m (97.4 %) y, además, sin obras en los 5 bloques con lluvia menor de 400 mm (95.0 %)."]),
  bul([B("Estudios por cerrar: "), "actualización del E8 con el AID redelimitado; lluvias de diseño, curvas IDF y caudales del E4; versión final del E6 (27/11)."]),
  bul([B("Cronograma: "), "Capítulo de Identificación en diciembre de 2026 (avance mínimo de 66.5 %), Perfil completo en marzo de 2027 y declaración de viabilidad en abril de 2027."]),

  h("5. DECISIONES REQUERIDAS DE LA ALTA DIRECCIÓN"),
  tabla([{ w: 650, a: C }, { w: W - 650 - 1900 }, { w: 1900, a: C }], ["N°", "DECISIÓN SOLICITADA", "PLAZO SUGERIDO"], [
    ["D1", "Aprobar el criterio para ajustar la cartera al presupuesto total: zanja de la ficha ANIN (0.40 × 0.30 m) y evaluación de las obras en los bloques con lluvia menor de 400 mm.", "30-NOV-2026"],
    ["D2", "Priorizar la conformidad de las actas de libre disponibilidad y completar el equipo de campo de liberación de áreas (dos asistentes prediales por confirmar y el ámbito de Lalaquiz sin asignar).", "30-NOV-2026"],
    ["D3", "Disponer la redelimitación del AID y la actualización del Análisis de Riesgos (E8) antes del cierre del Capítulo de Identificación.", "15-DIC-2026"],
    ["D4", "Validar como línea base del indicador de brecha el criterio areal del E6 (8,341.83 ha) sobre la superficie neta del catálogo V8.", "15-DIC-2026"],
  ]),
  para(["Se recomienda adoptar D1 y D2 en noviembre de 2026, con seguimiento semanal, para cerrar el Capítulo de Identificación en diciembre de 2026 y entregar el Perfil completo en marzo de 2027 bajo el Sistema Invierte.pe."], { spacing: { before: 120, after: 160, line: 264 } }),

  new Paragraph({ spacing: { after: 0 }, children: [run("Atentamente,")] }),
  new Paragraph({ spacing: { before: 360, after: 0 }, children: [B("Ing. Héctor Salomón Cahuas Miller")] }),
  new Paragraph({ spacing: { after: 0 }, children: [run("Especialista en Proyectos de Infraestructura Natural", { size: 17 })] }),
  new Paragraph({ spacing: { after: 120 }, children: [run("ANIN — DIME — SESDI | Proyecto IN Piura | CUI 2669244", { size: 17 })] }),
  nota("Fuente: presentación de avance del Perfil ante la SESDI y Diagnóstico Integral IN Piura V3 (corte 03/11/2026); Resumen Ejecutivo del 14-SEP-2026. Cifras preliminares, sujetas a ajuste en la versión final del Perfil."),
];

const doc = new Document({
  creator: "ANIN - DIME - SESDI", title: "Resumen Ejecutivo IN Piura - 03/11/2026",
  styles: { default: { document: { run: { font: "Arial", size: 19 } } } },
  numbering: { config: [{ reference: "vi", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: L, style: { paragraph: { indent: { left: 400, hanging: 220 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1100, right: 794, bottom: 700, left: 794, header: 340, footer: 280 } } },
    headers: { default: header }, footers: { default: footer }, children,
  }],
});
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("ok", OUT); });
