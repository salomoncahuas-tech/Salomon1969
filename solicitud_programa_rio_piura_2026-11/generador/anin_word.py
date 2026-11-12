"""Helpers python-docx con formato institucional ANIN (Arial; azul #1B4F72 / #2E86AB)."""
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

AZUL = "1B4F72"
AZUL2 = "2E86AB"
FILA_ALT = "EAF2F8"
ENCABEZADO = [
    "AUTORIDAD NACIONAL DE INFRAESTRUCTURA – ANIN",
    "DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA – DIME",
    "SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN",
]
PIE = "Proyecto IN Piura · CUI 2669244 · Cuenca Alta del Río Piura · UTM WGS 84 Zona 17S (EPSG:32717)"


def _rgb(h):
    return RGBColor.from_string(h)


def _shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


_ORDEN_TBLPR = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW",
                "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
                "tblCaption", "tblDescription"]


def _insert_tblpr(tblPr, el):
    """Inserta un hijo de w:tblPr respetando el orden del esquema OOXML."""
    tag = el.tag.split("}")[1]
    pos = _ORDEN_TBLPR.index(tag)
    for i, ch in enumerate(list(tblPr)):
        t = ch.tag.split("}")[1]
        if t in _ORDEN_TBLPR and _ORDEN_TBLPR.index(t) > pos:
            tblPr.insert(i, el)
            return
    tblPr.append(el)


def _borders(table, color="7F8C8D", sz="4"):
    tbl = table._tbl
    tblPr = tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{e}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        b.append(el)
    _insert_tblpr(tblPr, b)


def _grid(table, anchos_emu):
    """Fija anchos reales (tblGrid + celdas) para Word y LibreOffice."""
    table.autofit = False
    tbl = table._tbl
    tblPr = tbl.tblPr
    grid = tbl.tblGrid
    for gc in list(grid):
        grid.remove(gc)
    for w in anchos_emu:
        gc = OxmlElement("w:gridCol"); gc.set(qn("w:w"), str(int(w / 635))); grid.append(gc)
    for row in table.rows:
        for i, c in enumerate(row.cells):
            c.width = int(anchos_emu[i])


def _field(run, instr):
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
    run._r.append(f1); run._r.append(it); run._r.append(f2)


def nuevo(horizontal=False):
    d = Document()
    st = d.styles["Normal"]
    st.font.name = "Arial"
    st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    st.paragraph_format.space_after = Pt(4)
    st.paragraph_format.line_spacing = 1.1
    for nivel, tam, col in ((1, 13, AZUL), (2, 11, AZUL), (3, 10, AZUL2)):
        h = d.styles[f"Heading {nivel}"]
        h.font.name = "Arial"; h.font.size = Pt(tam); h.font.bold = True; h.font.color.rgb = _rgb(col)
        rf = h.element.rPr.rFonts
        for att in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            if rf.get(qn(att)) is not None:
                del rf.attrib[qn(att)]
        for att in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rf.set(qn(att), "Arial")
        h.paragraph_format.space_before = Pt(10 if nivel == 1 else 6); h.paragraph_format.space_after = Pt(4)
    for s in ("List Bullet", "Caption"):
        d.styles[s].font.name = "Arial"
    z = d.settings.element.find(qn("w:zoom"))
    if z is not None and z.get(qn("w:percent")) is None:
        z.set(qn("w:percent"), "100")
    sec = d.sections[0]
    if horizontal:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    else:
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.2)
    sec.top_margin = Cm(2.6); sec.bottom_margin = Cm(2.0)
    hp = sec.header.paragraphs[0]
    for i, t in enumerate(ENCABEZADO):
        r = hp.add_run(t)
        r.bold = True; r.font.size = Pt(8 if i else 8.5); r.font.color.rgb = _rgb(AZUL); r.font.name = "Arial"
        if i < 2:
            r.add_break()
    pPr = hp._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr"); bot = OxmlElement("w:bottom")
    for k, v in (("val", "single"), ("sz", "8"), ("space", "1"), ("color", AZUL2)):
        bot.set(qn(f"w:{k}"), v)
    bdr.append(bot); pPr.append(bdr)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run(PIE + " · Página "); r.font.size = Pt(7.5); r.font.color.rgb = _rgb("566573")
    r2 = fp.add_run(); r2.font.size = Pt(7.5); _field(r2, "PAGE")
    return d


def ancho_util(d):
    s = d.sections[0]
    return s.page_width - s.left_margin - s.right_margin


def titulo_doc(d, titulo, subtitulo=None):
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(titulo); r.bold = True; r.font.size = Pt(14); r.font.color.rgb = _rgb(AZUL)
    if subtitulo:
        p2 = d.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p2.add_run(subtitulo); r.font.size = Pt(10); r.font.color.rgb = _rgb(AZUL2); r.bold = True


def datos(d, filas):
    """Cuadro de datos clave: lista de (etiqueta, valor)."""
    t = d.add_table(rows=0, cols=2)
    _borders(t)
    w = ancho_util(d)
    for et, val in filas:
        c = t.add_row().cells
        c[0].width, c[1].width = int(w * 0.25), int(w * 0.75)
        _shade(c[0], AZUL)
        rr = c[0].paragraphs[0].add_run(et); rr.bold = True; rr.font.color.rgb = _rgb("FFFFFF"); rr.font.size = Pt(9)
        rv = c[1].paragraphs[0].add_run(val); rv.font.size = Pt(9)
    _grid(t, [w * 0.25, w * 0.75])
    d.add_paragraph()
    return t


def h(d, texto, nivel=1):
    return d.add_heading(texto, level=nivel)


def p(d, texto, bold_prefix=None, italic=False, size=None):
    par = d.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if bold_prefix:
        r = par.add_run(bold_prefix); r.bold = True
        if size: r.font.size = Pt(size)
    r = par.add_run(texto); r.italic = italic
    if size: r.font.size = Pt(size)
    return par


def vinetas(d, items):
    for it in items:
        par = d.add_paragraph(style="List Bullet")
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if isinstance(it, tuple):
            r = par.add_run(it[0]); r.bold = True
            par.add_run(it[1])
        else:
            par.add_run(it)


def leyenda(d, texto):
    par = d.add_paragraph()
    r = par.add_run(texto); r.bold = True; r.font.size = Pt(9); r.font.color.rgb = _rgb(AZUL)
    par.paragraph_format.space_after = Pt(2)
    par.paragraph_format.keep_with_next = True


def fuente(d, texto):
    par = d.add_paragraph()
    r = par.add_run("Fuente: " + texto); r.italic = True; r.font.size = Pt(7.5); r.font.color.rgb = _rgb("566573")


def tabla(d, cabecera, filas, anchos=None, total=None, size=8.5, alinear_num=True):
    """Tabla con cabecera azul, filas alternas y bordes finos. anchos en fracciones del ancho útil."""
    t = d.add_table(rows=1, cols=len(cabecera))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borders(t)
    w = ancho_util(d)
    anchos = anchos or [1 / len(cabecera)] * len(cabecera)
    for i, txt in enumerate(cabecera):
        c = t.rows[0].cells[i]
        _shade(c, AZUL)
        r = c.paragraphs[0].add_run(str(txt)); r.bold = True; r.font.size = Pt(size); r.font.color.rgb = _rgb("FFFFFF")
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    # repetir cabecera
    trPr = t.rows[0]._tr.get_or_add_trPr(); th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true"); trPr.append(th)
    for k, fila in enumerate(filas + ([total] if total else [])):
        cells = t.add_row().cells
        es_total = total is not None and k == len(filas)
        for i, v in enumerate(fila):
            c = cells[i]
            if es_total:
                _shade(c, "D6EAF8")
            elif k % 2 == 1:
                _shade(c, FILA_ALT)
            txt = "" if v is None else str(v)
            r = c.paragraphs[0].add_run(txt); r.font.size = Pt(size); r.bold = es_total
            if alinear_num and isinstance(v, str) and v.replace(",", "").replace(".", "").replace("%", "").replace(" ", "").replace("-", "").isdigit():
                c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _grid(t, [w * a for a in anchos])
    return t


def salto(d):
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def firma(d, nombre="Ing. Héctor Salomón Cahuas Miller", cargo="Especialista en Infraestructura Verde – DIME / SESDI"):
    d.add_paragraph()
    for t, b in (("______________________________________", False), (nombre, True), (cargo, False)):
        par = d.add_paragraph(); par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = par.add_run(t); r.bold = b; r.font.size = Pt(9.5)


def n(x, dec=2):
    """Formato numérico peruano del proyecto: miles con coma, decimales con punto."""
    if x is None or x == "":
        return "Por determinar"
    if isinstance(x, (int,)) or (isinstance(x, float) and dec == 0):
        return f"{x:,.0f}"
    return f"{x:,.{dec}f}"


def pct(x, dec=1):
    return f"{x * 100:.{dec}f} %"
