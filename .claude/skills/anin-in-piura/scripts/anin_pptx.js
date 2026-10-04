// Helpers pptxgenjs para presentaciones ANIN (Proyecto IN Piura): tema oscuro verde / teal / dorado, Arial.
//
//   const A = require("./anin_pptx");   // NODE_PATH debe incluir pptxgenjs; PPTX_SKILL = carpeta de la skill `pptx`
//   const { pres, C, tarjeta, circulo, txt, HEX } = A.nuevaPresentacion({ titulo: "…" });
//   pres.addSection({ title: "Portada" });
//   let s = pres.addSlide({ masterName: "PORTADA", sectionTitle: "Portada" });
//   s.addText("TÍTULO", { placeholder: "title" });  s.addText("subtítulo", { placeholder: "body" });
//   s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Contexto" });
//   s.addText("Título de la diapositiva", { placeholder: "title" });
//   tarjeta(s, 0.6, 1.6, 5.95, 2.1);  circulo(s, 0.9, 1.9, 0.7, 1);  txt(s, "texto", { x: 1, y: 2, w: 4, h: 1, fontSize: 16, color: C.background1 });
//   await A.guardar(pres, "salida.pptx");     // escribe el archivo y aplica el tema (colores) al .pptx
//
// Lienzo LAYOUT_WIDE (13.33" x 7.5"), márgenes de 0.6". Todos los colores son del tema (pres.SchemeColor):
//   text1 = fondo oscuro · text2 = paneles verdes · background1 = texto blanco · background2 = texto atenuado
//   accent1 = dorado · accent2 = teal · accent3 = verde · accent5 = teal claro (acentos de texto)
// Los gráficos exigen hex, no colores del tema: usa HEX.verde / HEX.teal / HEX.oro y tipografía "+mn-lt".
// Reglas de diseño que ya cumplen estos helpers: títulos de 34 pt (≤ ~45 caracteres en una línea), cuerpo ≥ 14 pt,
// sin franjas decorativas, pie y número de diapositiva en el layout (no en cada diapositiva).
const path = require("path");
const pptxgen = require("pptxgenjs");

const THEME = {
  name: "ANIN IN Piura", headFontFace: "Arial", bodyFontFace: "Arial",
  colors: { dk1: "0E2A1A", lt1: "FFFFFF", dk2: "1B4D2E", lt2: "B7CDBE",
    accent1: "D4A82F", accent2: "1C9A8A", accent3: "5FA777", accent4: "2E7D4F", accent5: "8FD3C8", accent6: "8C6D1F",
    hlink: "8FD3C8", folHlink: "B7CDBE" },
};
const HEX = { verde: "5FA777", teal: "1C9A8A", oro: "D4A82F", texto: "FFFFFF", fondo: "0E2A1A" };

function nuevaPresentacion({ titulo, pie = "Proyecto IN Piura · CUI 2669244 · ANIN - DIME - SESDI" } = {}) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  if (titulo) pres.title = titulo;
  pres.author = "ANIN - DIME - SESDI"; pres.company = "ANIN";
  const C = pres.SchemeColor;

  pres.defineSlideMaster({
    title: "PORTADA", background: { color: C.text1 },
    objects: [
      { text: { text: "AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN", options: { x: 0.6, y: 0.5, w: 8, h: 0.3, fontSize: 12, bold: true, color: C.accent1, margin: 0 } } },
      { text: { text: "DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME", options: { x: 0.6, y: 0.8, w: 9, h: 0.28, fontSize: 11, color: C.background2, margin: 0 } } },
      { text: { text: "SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN - SESDI", options: { x: 0.6, y: 1.08, w: 9, h: 0.28, fontSize: 11, color: C.background2, margin: 0 } } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 2.3, w: 8.2, h: 1.5, fontSize: 48, bold: true, color: C.background1, valign: "top", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 3.95, w: 7.6, h: 1.3, fontSize: 18, color: C.background2, valign: "top", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENIDO", background: { color: C.text1 },
    objects: [
      { text: { text: "AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN", options: { x: 0.6, y: 0.18, w: 8, h: 0.25, fontSize: 11, bold: true, color: C.accent1, margin: 0 } } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.5, w: 12.1, h: 0.85, fontSize: 34, bold: true, color: C.background1, valign: "middle", margin: 0 }, text: "" } },
      { text: { text: pie, options: { x: 0.6, y: 7.05, w: 8, h: 0.28, fontSize: 11, color: C.background2, margin: 0 } } },
    ],
    slideNumber: { x: 12.2, y: 7.05, w: 0.5, h: 0.28, fontSize: 11, color: C.background2, align: "right" },
  });

  let n = 0;
  const nombre = (p) => `${p}-${++n}`; // objectName único: ayuda a leer el panel de selección y la accesibilidad
  const tarjeta = (s, x, y, w, h, fill = C.text2) =>
    s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: nombre("tarjeta") });
  const circulo = (s, x, y, d, texto, fill = C.accent1, color = C.text1, size = 18) =>
    s.addText(String(texto), { x, y, w: d, h: d, shape: pres.ShapeType.ellipse, fill: { color: fill }, color, bold: true, fontSize: size,
      align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: nombre("circulo") });
  const txt = (s, t, o) => s.addText(t, { isTextBox: true, margin: 0, valign: "top", objectName: nombre("texto"), ...o });
  return { pres, C, tarjeta, circulo, txt, nombre };
}

// Escribe el .pptx y aplica el tema (pptxgenjs no puede escribir los colores del tema por sí solo).
async function guardar(pres, fileName) {
  const skill = process.env.PPTX_SKILL;
  if (!skill) throw new Error("Define PPTX_SKILL con la carpeta de la skill `pptx` (contiene scripts/apply_theme.js).");
  const { applyTheme } = require(path.join(skill, "scripts/apply_theme.js"));
  await pres.writeFile({ fileName });
  await applyTheme(fileName, THEME);
  return fileName;
}

module.exports = { THEME, HEX, nuevaPresentacion, guardar };
