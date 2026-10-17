// Biblioteca de maquetación para el Segundo Entregable OS 4062-2026 (formato del 1er Entregable).
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer, AlignmentType, WidthType,
  BorderStyle, ShadingType, HeadingLevel, LevelFormat, PageNumber, PageBreak, ImageRun, VerticalAlign,
  PageOrientation, TableLayoutType,
} = require("docx");

const VERDE = "1B4D2E", AZUL = "1B4F72", CLARO = "E8F0EA", TOTAL = "D9E6DD", GRIS = "595959";
const W = 9638;          // ancho útil vertical (A4, 2 cm)
const WH = 14570;        // ancho útil horizontal
const L = AlignmentType.LEFT, R = AlignmentType.RIGHT, C = AlignmentType.CENTER, J = AlignmentType.JUSTIFIED;

const run = (t, o = {}) => new TextRun({ text: String(t), font: "Arial", size: 21, ...o });
const b = (t, o = {}) => run(t, { bold: true, ...o });
const i = (t, o = {}) => run(t, { italics: true, ...o });
const runs = (p) => (Array.isArray(p) ? p : [p]).map((x) => (typeof x === "string" ? run(x) : x));
const para = (p, o = {}) => new Paragraph({ children: runs(p), spacing: { after: 120, line: 276 }, alignment: J, ...o });
const centro = (p, o = {}) => para(p, { alignment: C, ...o });
const vacio = (n = 1) => Array.from({ length: n }, () => new Paragraph({ children: [run("")], spacing: { after: 0 } }));
const salto = () => new Paragraph({ children: [new PageBreak()] });

const h1 = (t, o = {}) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t })], keepNext: true, pageBreakBefore: !!o.nueva });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t })], keepNext: true });
const h3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text: t })], keepNext: true });
// Cita del literal del TdR (negrita cursiva, como el modelo)
const literal = (t) => new Paragraph({ children: [run(t, { bold: true, italics: true })], spacing: { after: 140, line: 276 }, alignment: J });
const guion = (p) => new Paragraph({ numbering: { reference: "guion", level: 0 }, spacing: { after: 80, line: 276 }, children: runs(p), alignment: J });
const bullet = (p) => new Paragraph({ numbering: { reference: "vi", level: 0 }, spacing: { after: 80, line: 276 }, children: runs(p), alignment: J });
const num = (n, p) => para([b(n + " "), ...runs(p)]);   // 5.1. Texto
const cap = (t) => new Paragraph({ children: [run(t, { bold: true, size: 19 })], spacing: { before: 160, after: 80 }, keepNext: true, keepLines: true });
const fuente = (t) => new Paragraph({ children: [run(t, { italics: true, size: 15, color: GRIS })], spacing: { before: 50, after: 200 }, alignment: J });

const fino = { style: BorderStyle.SINGLE, size: 4, color: "8C8C8C" };
const bordes = { top: fino, bottom: fino, left: fino, right: fino };
const sinB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const sinBordes = { top: sinB, bottom: sinB, left: sinB, right: sinB };

function celda(texto, w, { head = false, fill, align = L, bold = false, size = 17, color, span, vmerge, italics = false } = {}) {
  const partes = Array.isArray(texto) ? texto : String(texto ?? "").split("\n");
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, borders: bordes, columnSpan: span,
    verticalAlign: VerticalAlign.CENTER,
    shading: head ? { type: ShadingType.CLEAR, fill: VERDE, color: "auto" } : fill ? { type: ShadingType.CLEAR, fill, color: "auto" } : undefined,
    margins: { top: 45, bottom: 45, left: 80, right: 80 },
    children: partes.map((pt) => new Paragraph({ alignment: head ? C : align, spacing: { after: 0 },
      children: (typeof pt === "string" ? [run(pt, { size, bold: head || bold, italics, color: head ? "FFFFFF" : color })] : [pt]) })),
  });
}

// cols: [{w, align}] ; filas: [[...]] ; opts.total: fila final ; opts.ancho
function tabla(cols, cabecera, filas, opts = {}) {
  const size = opts.size || 17;
  const trs = [];
  const cabs = Array.isArray(cabecera[0]) ? cabecera : [cabecera];
  cabs.forEach((cab) => trs.push(new TableRow({ tableHeader: true, cantSplit: true,
    children: cab.map((h, k) => (typeof h === "object" && h !== null && !Array.isArray(h) && h.t !== undefined)
      ? celda(h.t, h.w ?? cols[k].w, { head: true, size, span: h.span })
      : celda(h, cols[k].w, { head: true, size })) })));
  filas.forEach((r, k) => trs.push(new TableRow({ cantSplit: opts.cantSplit !== false,
    children: r.map((v, j) => celda(v, cols[j].w, { fill: k % 2 ? CLARO : undefined, align: cols[j].align || L, size, bold: cols[j].bold })) })));
  if (opts.total) trs.push(new TableRow({ cantSplit: true,
    children: opts.total.map((v, j) => celda(v, cols[j].w, { fill: TOTAL, bold: true, align: cols[j].align || L, size })) }));
  return new Table({ width: { size: cols.reduce((s, c) => s + c.w, 0), type: WidthType.DXA },
    columnWidths: cols.map((c) => c.w), rows: trs, layout: TableLayoutType.FIXED });
}

// Cuadro de datos del informe (Dirigido a / De / Asunto ...): [[etiqueta, valor], ...]
function meta(filas, wE = 2000, ancho = W) {
  return new Table({ width: { size: ancho, type: WidthType.DXA }, columnWidths: [wE, ancho - wE], layout: TableLayoutType.FIXED,
    rows: filas.map(([k, v]) => new TableRow({ cantSplit: true, children: [
      celda(k, wE, { fill: CLARO, bold: true, size: 19 }),
      celda(String(v).split("\n").map((x, n) => (n === 0 ? ": " + x : "  " + x)), ancho - wE, { size: 19 })] })) });
}

function encabezado() {
  const lineas = [["AUTORIDAD NACIONAL DE INFRAESTRUCTURA — ANIN", 17], ["DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA — DIME", 14], ["SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN — SESDI", 14]];
  return new Header({ children: lineas.map(([t, s], k, a) => new Paragraph({ alignment: C, spacing: { after: 0 },
    border: k === a.length - 1 ? { bottom: { style: BorderStyle.SINGLE, size: 6, color: VERDE, space: 4 } } : undefined,
    children: [run(t, { bold: k === 0, size: s, color: VERDE })] })) });
}
function pie(texto) {
  return new Footer({ children: [new Paragraph({ alignment: C,
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: "999999", space: 4 } },
    children: [run(texto + " · Página ", { size: 15, color: GRIS }),
      new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 15, color: GRIS }),
      run(" de ", { size: 15, color: GRIS }),
      new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 15, color: GRIS })] })] });
}
function pagina(horizontal = false) {
  return { page: { size: { width: 11906, height: 16838, ...(horizontal ? { orientation: PageOrientation.LANDSCAPE } : {}) },
    margin: { top: 1560, bottom: 1200, left: 1134, right: 1134, header: 560, footer: 500 } } };
}
function estilos() {
  return {
    styles: {
      default: { document: { run: { font: "Arial", size: 21 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: "Arial", size: 24, bold: true, color: VERDE }, paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: "Arial", size: 22, bold: true, color: VERDE }, paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 1 } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: "Arial", size: 21, bold: true, color: "2E5E3E" }, paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } },
      ],
    },
    numbering: { config: [
      { reference: "vi", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: L, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
      { reference: "guion", levels: [{ level: 0, format: LevelFormat.BULLET, text: "—", alignment: L, style: { paragraph: { indent: { left: 640, hanging: 360 } } } }] },
    ] },
  };
}

// Figura: imagen PNG/JPG escalada al ancho indicado (cm)
function figura(ruta, anchoCm = 16, tipo) {
  const buf = fs.readFileSync(ruta);
  let wpx, hpx;
  if (buf.slice(1, 4).toString() === "PNG") { wpx = buf.readUInt32BE(16); hpx = buf.readUInt32BE(20); tipo = tipo || "png"; }
  else { // JPEG: buscar SOF
    let o = 2; while (o < buf.length) { const m = buf[o + 1]; const len = buf.readUInt16BE(o + 2);
      if (m >= 0xc0 && m <= 0xc3) { hpx = buf.readUInt16BE(o + 5); wpx = buf.readUInt16BE(o + 7); break; } o += 2 + len; }
    tipo = tipo || "jpg"; }
  const wPt = anchoCm * 37.8; const hPt = wPt * hpx / wpx;
  return new Paragraph({ alignment: C, spacing: { before: 80, after: 60 }, keepNext: true,
    children: [new ImageRun({ type: tipo, data: buf, transformation: { width: Math.round(wPt), height: Math.round(hPt) } })] });
}

function seccion(children, { pie: textoPie, horizontal = false } = {}) {
  return { properties: pagina(horizontal), headers: { default: encabezado() }, footers: { default: pie(textoPie) }, children };
}
async function guardar(secciones, ruta, titulo = "") {
  const doc = new Document({ creator: "Ing. Héctor Salomón Cahuas Miller", title: titulo, ...estilos(), sections: secciones });
  fs.writeFileSync(ruta, await Packer.toBuffer(doc));
  console.log("OK", ruta);
}

module.exports = { VERDE, AZUL, CLARO, TOTAL, GRIS, W, WH, L, R, C, J, run, b, i, runs, para, centro, vacio, salto, h1, h2, h3,
  literal, guion, bullet, num, cap, fuente, celda, tabla, meta, encabezado, pie, pagina, estilos, figura, seccion, guardar };

// Numeración automática de cuadros y figuras por documento
function numerador(prefijo = "") {
  let c = 0, f = 0;
  return {
    cuadro: (t) => cap(`Cuadro N° ${prefijo}${++c}. ${t}`),
    figura: (t) => cap(`Figura N° ${prefijo}${++f}. ${t}`),
    sigC: () => `${prefijo}${c + 1}`, sigF: () => `${prefijo}${f + 1}`,
  };
}
const fmt = (n, d = 2) => Number(n).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
// Firma
function firma(fecha) {
  const { Paragraph } = require("docx");
  return [...vacio(3),
    new Paragraph({ alignment: C, spacing: { after: 0 }, children: [run("_______________________________________")] }),
    new Paragraph({ alignment: C, spacing: { after: 0 }, children: [b("Ing. HÉCTOR SALOMÓN CAHUAS MILLER")] }),
    new Paragraph({ alignment: C, spacing: { after: 0 }, children: [run("Ingeniero Forestal – Especialista en Infraestructura Verde", { size: 19 })] }),
    new Paragraph({ alignment: C, spacing: { after: 0 }, children: [run("RUC N° 10096435270 · DNI N° 09643527", { size: 19 })] }),
    ...(fecha ? [new Paragraph({ alignment: C, spacing: { before: 120 }, children: [run(fecha, { size: 19 })] })] : [])];
}
module.exports.numerador = numerador; module.exports.fmt = fmt; module.exports.firma = firma;
