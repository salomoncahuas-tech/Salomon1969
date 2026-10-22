// Carátulas de anexos — Segundo Entregable OS 4062-2026 (formato ANIN, Arial, verde #1B4D2E)
const { Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType,
  VerticalAlign, TableLayoutType, HeightRule, Footer, PageBreak } = require("docx");
const L_ = require("./lib");
const { VERDE, CLARO, GRIS, W, run } = L_;
const ORO = "B8962E", VERDE2 = "2E6B47", TINTA = "1F2D24";
const C = AlignmentType.CENTER, LEFT = AlignmentType.LEFT, J = AlignmentType.JUSTIFIED;
const nb = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const NB = { top: nb, bottom: nb, left: nb, right: nb };

const p = (children, o = {}) => new Paragraph({ spacing: { after: 0 }, ...o, children });
const esp = (pt) => p([run("", { size: 2 })], { spacing: { before: 0, after: pt * 20 } });

function cell(children, w, { fill, borders = NB, margins = { top: 120, bottom: 120, left: 280, right: 280 }, vAlign = VerticalAlign.CENTER, span } = {}) {
  return new TableCell({ width: { size: w, type: WidthType.DXA }, borders, columnSpan: span, verticalAlign: vAlign, margins,
    shading: fill ? { type: ShadingType.CLEAR, fill, color: "auto" } : undefined, children });
}
const tbl = (rows, cols) => new Table({ width: { size: cols.reduce((a, c) => a + c, 0), type: WidthType.DXA }, columnWidths: cols,
  layout: TableLayoutType.FIXED, rows });

/**
 * caratula({ numero: "1", actividad: "b)", textoActividad, titulo, subtitulo, coordinacion, comprende: [..], fecha, informe })
 * variante: "anexo" (página completa) | "separador" (sub-anexo, p. ej. 3-A)
 */
function caratula(o) {
  const out = [];
  const sep = o.variante === "separador";
  out.push(esp(sep ? 150 : (o.espIni ?? 30)));
  // Rótulo superior
  out.push(p([run("PROYECTO IN PIURA  ·  CUI 2669244  ·  CÓDIGO DE IDEA 337830", { size: 16, bold: true, color: ORO, characterSpacing: 30 })], { alignment: LEFT }));
  out.push(p([run("Segundo Entregable  ·  Orden de Servicio N° 4062-2026  ·  " + o.informe, { size: 16, color: GRIS })], { alignment: LEFT, spacing: { after: 160 } }));

  // Banda institucional: filete dorado + bloque verde con el número de anexo
  const banda = tbl([
    new TableRow({ height: { value: 70, rule: HeightRule.EXACT }, children: [cell([p([run("", { size: 2 })])], W, { fill: ORO, margins: { top: 0, bottom: 0, left: 0, right: 0 } })] }),
    new TableRow({ children: [cell([
      p([run(sep ? "ANEXO " + o.numero : "ANEXO N° " + o.numero, { size: sep ? 44 : 64, bold: true, color: "FFFFFF" })], { spacing: { after: 60 } }),
      p([run("ACTIVIDAD " + o.actividad + "  —  NUMERAL 5.1 DE LOS TÉRMINOS DE REFERENCIA", { size: 18, bold: true, color: "D9C27A", characterSpacing: 20 })]),
    ], W, { fill: VERDE, margins: { top: sep ? 260 : 420, bottom: sep ? 240 : 380, left: 400, right: 400 } })] }),
  ], [W]);
  out.push(banda);
  out.push(esp(sep ? 16 : 22));

  // Título del anexo
  out.push(p([run(o.titulo.toUpperCase(), { size: sep ? 30 : 34, bold: true, color: VERDE })], { alignment: LEFT, spacing: { after: 120, line: 300 } }));
  if (o.subtitulo) out.push(p([run(o.subtitulo, { size: 22, color: VERDE2 })], { alignment: LEFT, spacing: { after: 240, line: 280 } }));

  // Literal del TdR en recuadro con filete dorado a la izquierda
  if (o.textoActividad) {
    out.push(tbl([new TableRow({ children: [
      cell([p([run("", { size: 2 })])], 90, { fill: ORO, margins: { top: 0, bottom: 0, left: 0, right: 0 } }),
      cell([
        p([run("Literal del TdR — actividad " + o.actividad, { size: 15, bold: true, color: GRIS, characterSpacing: 20 })], { spacing: { after: 60 } }),
        p([run("«" + o.textoActividad + "»", { size: 18, italics: true, color: TINTA })], { alignment: J, spacing: { line: 264 } }),
      ], W - 90, { fill: "F3F6F2", margins: { top: 140, bottom: 140, left: 240, right: 240 } }),
    ] })], [90, W - 90]));
    out.push(esp(16));
  }

  // Contenido del anexo
  if (o.comprende && o.comprende.length) {
    out.push(p([run("CONTENIDO DEL ANEXO", { size: 17, bold: true, color: VERDE, characterSpacing: 30 })], {
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: VERDE, space: 3 } }, spacing: { after: 100 } }));
    o.comprende.forEach(([cod, txt]) => out.push(tbl([new TableRow({ children: [
      cell([p([run(cod, { size: 17, bold: true, color: VERDE })])], 900, { margins: { top: 40, bottom: 40, left: 0, right: 80 }, vAlign: VerticalAlign.TOP }),
      cell([p([run(txt, { size: 18, color: TINTA })], { alignment: J })], W - 900, { margins: { top: 40, bottom: 40, left: 80, right: 0 }, vAlign: VerticalAlign.TOP }),
    ] })], [900, W - 900])));
    out.push(esp(sep ? 14 : 20));
  }

  // Cuadro de datos
  const fino = { style: BorderStyle.SINGLE, size: 4, color: "B9C8BD" };
  const fb = { top: fino, bottom: fino, left: nb, right: nb };
  const filas = [
    ["Servicio", "Servicio especializado para la determinación de área geográfica para el componente de infraestructura verde del estudio de preinversión del proyecto de Infraestructura Natural Piura"],
    ["Estudio de preinversión", "Recuperación del servicio de regulación de riesgos naturales y recuperación de ecosistemas degradados en la Cuenca Alta del Río Piura — CUI 2669244"],
    ...(o.coordinacion ? [["En coordinación con", o.coordinacion]] : []),
    ...(o.ambito ? [["Ámbito", o.ambito]] : []),
    ["Elaborado por", "Ing. Héctor Salomón Cahuas Miller — Ingeniero Forestal, Especialista en Infraestructura Verde"],
    ["Periodo", o.periodo],
  ];
  out.push(tbl(filas.map(([k, v]) => new TableRow({ cantSplit: true, children: [
    cell([p([run(k.toUpperCase(), { size: 14, bold: true, color: VERDE, characterSpacing: 15 })])], 2300, { borders: fb, fill: CLARO, margins: { top: 70, bottom: 70, left: 140, right: 100 } }),
    cell([p([run(v, { size: 17, color: TINTA })], { alignment: LEFT })], W - 2300, { borders: fb, margins: { top: 70, bottom: 70, left: 160, right: 100 } }),
  ] })), [2300, W - 2300]));

  // Pie de carátula
  out.push(esp(sep ? 60 : (o.espFinal ?? 40)));
  out.push(p([run(o.fecha, { size: 19, bold: true, color: VERDE })], { alignment: C, spacing: { after: 30 } }));
  out.push(p([run("Autoridad Nacional de Infraestructura · DIME · SESDI · UTM WGS 84 Zona 17S (EPSG:32717)", { size: 15, color: GRIS })], { alignment: C }));
  return out;
}

// Sección de carátula: encabezado institucional, sin pie de página (la numeración arranca en el cuerpo del anexo)
function seccionCaratula(o) {
  return { properties: L_.pagina(false), headers: { default: L_.encabezado() },
    footers: { default: new Footer({ children: [p([run("")])] }) }, children: caratula(o) };
}

module.exports = { caratula, seccionCaratula, ORO };
