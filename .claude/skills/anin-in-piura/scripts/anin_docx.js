// Helpers docx-js con el formato institucional ANIN (Proyecto IN Piura).
//
//   const A = require("./anin_docx");           // NODE_PATH debe apuntar a las librerías globales (docx)
//   const doc = new Document({ ...A.estilos(), sections: [{
//     properties: A.pagina(),                    // A4 vertical, márgenes de 2 cm
//     headers: { default: A.encabezado() }, footers: { default: A.pie() },
//     children: [A.h1("1. ANTECEDENTES"), A.para("Texto…"), A.cap("Tabla 1. …"),
//                A.tabla(cols, cabecera, filas, filaTotal), A.fuente("Fuente: …")] }] });
//
// Reglas que ya resuelve: Arial, verde #1B4D2E, ancho de tabla = suma de columnas (DXA, requisito de
// Word/Google Docs), sombreado CLEAR (SOLID se ve negro), viñetas con numbering real, cabecera de tabla
// que se repite en cada página y filas que no se parten.
const {
  Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer, AlignmentType, WidthType, BorderStyle,
  ShadingType, HeadingLevel, LevelFormat, PageNumber, PositionalTab, PositionalTabAlignment,
  PositionalTabRelativeTo, PositionalTabLeader,
} = require("docx");

const VERDE = "1B4D2E", AZUL = "1B4F72", CLARO = "E8F0EA", TOTAL = "D9E6DD", GRIS = "595959";
const W = 9638; // ancho útil A4 con márgenes de 2 cm (DXA)
const L = AlignmentType.LEFT, R = AlignmentType.RIGHT, C = AlignmentType.CENTER;

const f0 = (n) => n.toLocaleString("en-US", { maximumFractionDigits: 0 });
const f1 = (n) => n.toLocaleString("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const f3 = (n) => n.toLocaleString("en-US", { minimumFractionDigits: 3, maximumFractionDigits: 3 });
const pct = (x) => f1(x * 100) + " %";

const run = (t, o = {}) => new TextRun({ text: t, font: "Arial", size: 21, ...o });
const b = (t) => run(t, { bold: true });
const runs = (parts) => (Array.isArray(parts) ? parts : [parts]).map((x) => (typeof x === "string" ? run(x) : x));
const para = (parts, o = {}) => new Paragraph({ children: runs(parts), spacing: { after: 120, line: 276 }, alignment: AlignmentType.JUSTIFIED, ...o });
// h1(texto, { nuevaPagina: true }) para iniciar una sección (p. ej. un anexo) en página nueva.
const h1 = (t, { nuevaPagina = false } = {}) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t })], keepNext: true, pageBreakBefore: nuevaPagina });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t })], keepNext: true });
const bullet = (parts) => new Paragraph({ numbering: { reference: "vi", level: 0 }, spacing: { after: 80, line: 276 },
  children: runs(parts), alignment: AlignmentType.JUSTIFIED });
const cap = (t) => new Paragraph({ children: [run(t, { bold: true, size: 19, color: VERDE })], spacing: { before: 160, after: 80 }, keepNext: true });
const fuente = (t) => new Paragraph({ children: [run(t, { italics: true, size: 16, color: GRIS })], spacing: { before: 60, after: 200 } });

const fino = { style: BorderStyle.SINGLE, size: 4, color: "999999" };
const bordes = { top: fino, bottom: fino, left: fino, right: fino };

function celda(texto, w, { head = false, fill, align = L, bold = false, color = VERDE } = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, borders: bordes,
    shading: head ? { type: ShadingType.CLEAR, fill: color, color: "auto" } : fill ? { type: ShadingType.CLEAR, fill, color: "auto" } : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    children: [new Paragraph({ alignment: head ? C : align,
      children: [run(String(texto), { size: 18, bold: head || bold, color: head ? "FFFFFF" : undefined })] })],
  });
}

// cols: [{w, align}] (la suma de w debe ser W para ocupar todo el ancho); cabecera: [..]; filas: [[..]]; total: [..] | null
function tabla(cols, cabecera, filas, total = null) {
  const trs = [new TableRow({ tableHeader: true, cantSplit: true, children: cabecera.map((h, i) => celda(h, cols[i].w, { head: true })) })];
  filas.forEach((r, k) => trs.push(new TableRow({ cantSplit: true,
    children: r.map((v, i) => celda(v, cols[i].w, { fill: k % 2 ? CLARO : undefined, align: cols[i].align })) })));
  if (total) trs.push(new TableRow({ cantSplit: true,
    children: total.map((v, i) => celda(v, cols[i].w, { fill: TOTAL, bold: true, align: cols[i].align })) }));
  return new Table({ width: { size: cols.reduce((s, c) => s + c.w, 0), type: WidthType.DXA },
    columnWidths: cols.map((c) => c.w), rows: trs });
}

// Cuadro de datos del informe: [[etiqueta, valor], ...]
function meta(filas, wEtiqueta = 2200) {
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [wEtiqueta, W - wEtiqueta],
    rows: filas.map(([k, v]) => new TableRow({ cantSplit: true,
      children: [celda(k, wEtiqueta, { fill: CLARO, bold: true }), celda(v, W - wEtiqueta)] })) });
}

function encabezado(color = VERDE) {
  const lineas = ["AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN",
    "DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME",
    "SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN - SESDI"];
  return new Header({ children: lineas.map((t, i, a) => new Paragraph({
    alignment: C, spacing: { after: 0 },
    border: i === a.length - 1 ? { bottom: { style: BorderStyle.SINGLE, size: 8, color, space: 4 } } : undefined,
    children: [run(t, { bold: true, size: i === 0 ? 18 : 15, color })] })) });
}

function pie(texto = "Proyecto IN Piura · CUI 2669244 · Cuenca Alta del Río Piura") {
  return new Footer({ children: [new Paragraph({
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: "999999", space: 4 } },
    children: [run(texto, { size: 15, color: GRIS }),
      new TextRun({ children: [new PositionalTab({ alignment: PositionalTabAlignment.RIGHT, relativeTo: PositionalTabRelativeTo.MARGIN, leader: PositionalTabLeader.NONE }),
        "Página ", PageNumber.CURRENT, " de ", PageNumber.TOTAL_PAGES], font: "Arial", size: 15, color: GRIS })] })] });
}

// A4 vertical con márgenes de 2 cm. Para matrices anchas: pagina({ horizontal: true }) (docx-js intercambia ancho y alto).
function pagina({ horizontal = false } = {}) {
  const { PageOrientation } = require("docx");
  return { page: { size: { width: 11906, height: 16838, ...(horizontal ? { orientation: PageOrientation.LANDSCAPE } : {}) },
    margin: { top: 1500, bottom: 1200, left: 1134, right: 1134, header: 600, footer: 500 } } };
}

// Estilos de título jerárquicos reales (Título 1/2), numeración de viñetas y fuente por defecto.
function estilos(color = VERDE) {
  return {
    styles: {
      default: { document: { run: { font: "Arial", size: 21 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: "Arial", size: 26, bold: true, color }, paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: "Arial", size: 23, bold: true, color }, paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 } },
      ],
    },
    numbering: { config: [{ reference: "vi", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: L,
      style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  };
}

module.exports = { VERDE, AZUL, CLARO, GRIS, W, L, R, C, f0, f1, f3, pct, run, b, para, h1, h2, bullet, cap, fuente,
  celda, tabla, meta, encabezado, pie, pagina, estilos };
