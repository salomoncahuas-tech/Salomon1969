"""Genera los árboles de problemas del Proyecto IN Piura en formatos editables.

    python arboles_problemas/generar_arboles.py

Lee `arboles_problemas/arboles_IN_Piura.json` (modelo único de los 4 árboles) y escribe:
  - Arboles_Problemas_IN_Piura.pptx   PowerPoint A3 horizontal, una lámina por árbol, cajas y
                                       conectores nativos anclados (al mover una caja, la flecha la sigue)
  - Arboles_Problemas_IN_Piura.drawio  diagrams.net / draw.io, una página por árbol (abre también en
                                       app.diagrams.net y se importa en Visio o Lucidchart)

Ambos salen del mismo cálculo de diagramación, así que las láminas y el diagrama coinciden.
El texto se mide con Liberation Sans (métrica idéntica a Arial) para dimensionar las cajas.
"""
import copy
import json
import sys
from pathlib import Path
from xml.sax.saxutils import escape

from PIL import ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

AQUI = Path(__file__).resolve().parent
FUENTE = "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
FUENTE_B = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"

# Lámina A3 horizontal (pulgadas)
ANCHO, ALTO = 16.535, 11.693
MARGEN = 0.45
SEP_COL = 0.22          # separación horizontal entre cajas
SANGRIA = 0.30          # sangría de las causas indirectas bajo su causa directa
PAD_X, PAD_Y = 0.08, 0.05

# Paleta ANIN (azul institucional #1B4F72) para niveles del árbol
COL = {
    "verde": "1B4F72", "verde2": "2A6A98", "verde_claro": "E5EEF6", "borde_cd": "1B4F72",
    "ci_fondo": "F3F7FA", "ci_borde": "7B9BB5",
    "pc_fondo": "1B4F72", "pc_texto": "FFFFFF",
    "ed_fondo": "F6E7C8", "ed_borde": "B8862B",
    "ei_fondo": "FBF3E2", "ei_borde": "C9A24F",
    "ef_fondo": "8C5A12", "ef_texto": "FFFFFF",
    "ex_fondo": "FFFFFF", "ex_borde": "6B7A89",
    "texto": "1B2733", "gris": "55677A", "linea": "4A5867",
}
TAM_BASE = {"EF": 12, "EI": 9.5, "ED": 10, "PC": 13, "CD": 10, "CI": 9, "EX": 8.5}
TAM = dict(TAM_BASE)  # se escala por lámina para aprovechar la altura A3 (ver ajustar_escala)
NEGRITA = {"EF": True, "EI": False, "ED": True, "PC": True, "CD": True, "CI": False, "EX": False}

_fuentes = {}


def _font(pt, bold):
    clave = (round(pt * 4), bold)
    if clave not in _fuentes:
        # 1 pt = 1/72 in; medimos a 300 ppp para precisión
        _fuentes[clave] = ImageFont.truetype(FUENTE_B if bold else FUENTE, int(round(pt / 72 * 300)))
    return _fuentes[clave]


def lineas_texto(texto, ancho_in, pt, bold):
    """Parte el texto en líneas que caben en ancho_in (pulgadas) con Arial/Liberation Sans."""
    f = _font(pt, bold)
    limite = ancho_in * 300
    salida = []
    for parrafo in texto.split("\n"):
        actual = ""
        for palabra in parrafo.split():
            prueba = (actual + " " + palabra).strip()
            if f.getlength(prueba) <= limite or not actual:
                actual = prueba
            else:
                salida.append(actual)
                actual = palabra
        salida.append(actual)
    return salida


def alto_texto(texto, ancho_in, pt, bold, interlineado=1.2):
    n = len(lineas_texto(texto, ancho_in - 2 * PAD_X, pt, bold))
    return n * pt * interlineado / 72 + 2 * PAD_Y


def etiqueta(nodo):
    return f"{nodo['id']}  {nodo['texto']}"


# ---------------------------------------------------------------- diagramación
def diagramar(arbol):
    """Devuelve cajas {id: dict(x,y,w,h,nivel,texto)} y enlaces [(hijo, padre, tipo)] en pulgadas."""
    cajas, enlaces = {}, []
    util = ANCHO - 2 * MARGEN
    y = 1.05  # bajo la banda de encabezado

    def fila(nodos, nivel, y0, anchos=None):
        n = len(nodos)
        w = (util - (n - 1) * SEP_COL) / n if anchos is None else None
        hs = [alto_texto(etiqueta(nd), w, TAM[nivel], NEGRITA[nivel]) for nd in nodos]
        h = max(max(hs), 0.5)
        for i, nd in enumerate(nodos):
            cajas[nd["id"]] = dict(x=MARGEN + i * (w + SEP_COL), y=y0, w=w, h=h, nivel=nivel, nodo=nd)
        return h

    # Efecto final (centrado, ancho 60 %)
    ef = arbol["efecto_final"]
    w_ef = util * 0.62
    h_ef = max(alto_texto(ef["texto"], w_ef, TAM["EF"], True), 0.55)
    cajas[ef["id"]] = dict(x=MARGEN + (util - w_ef) / 2, y=y, w=w_ef, h=h_ef, nivel="EF", nodo=ef)
    y += h_ef + 0.42

    h = fila(arbol["efectos_indirectos"], "EI", y)
    y += h + 0.42
    h = fila(arbol["efectos_directos"], "ED", y)
    y += h + 0.45

    # Tronco: problema central al centro, factores exógenos a los lados
    pc = arbol["problema"]
    w_pc = util * 0.46
    texto_pc = pc["texto"]
    h_pc = max(alto_texto(texto_pc, w_pc, TAM["PC"], True) + 0.05, 0.85)
    ex = arbol.get("factores_exogenos", [])
    w_ex = (util - w_pc - 2 * 0.35) / 2
    h_ex = max([alto_texto(f["titulo"] + "\n" + f["texto"], w_ex, TAM["EX"], False) for f in ex] + [h_pc])
    h_tronco = max(h_pc, h_ex)
    cajas[pc["id"]] = dict(x=MARGEN + (util - w_pc) / 2, y=y + (h_tronco - h_pc) / 2, w=w_pc, h=h_pc, nivel="PC", nodo=pc)
    for i, f in enumerate(ex[:2]):
        x = MARGEN if i == 0 else ANCHO - MARGEN - w_ex
        cajas[f["id"]] = dict(x=x, y=y + (h_tronco - h_ex) / 2, w=w_ex, h=h_ex, nivel="EX", nodo=f)
        enlaces.append((f["id"], pc["id"], "exogeno"))
    y += h_tronco + 0.45

    # Causas directas y, debajo, sus causas indirectas apiladas
    cds = arbol["causas_directas"]
    n = len(cds)
    w = (util - (n - 1) * SEP_COL) / n
    h_cd = max(max(alto_texto(etiqueta(c), w, TAM["CD"], True) for c in cds), 0.55)
    y_ci0 = y + h_cd + 0.22
    for i, c in enumerate(cds):
        x = MARGEN + i * (w + SEP_COL)
        cajas[c["id"]] = dict(x=x, y=y, w=w, h=h_cd, nivel="CD", nodo=c)
        enlaces.append((c["id"], pc["id"], "causa"))
    # Altura uniforme por "fila" de causas indirectas para alinear columnas
    filas = max(len([ci for ci in arbol["causas_indirectas"] if ci["padre"] == c["id"]]) for c in cds)
    w_ci = w - SANGRIA
    alturas = []
    for k in range(filas):
        hs = []
        for c in cds:
            hijos = [ci for ci in arbol["causas_indirectas"] if ci["padre"] == c["id"]]
            if k < len(hijos):
                hs.append(alto_texto(etiqueta(hijos[k]), w_ci, TAM["CI"], False))
        alturas.append(max(hs + [0.42]))
    for i, c in enumerate(cds):
        hijos = [ci for ci in arbol["causas_indirectas"] if ci["padre"] == c["id"]]
        yy = y_ci0
        for k, ci in enumerate(hijos):
            cajas[ci["id"]] = dict(x=cajas[c["id"]]["x"] + SANGRIA, y=yy, w=w_ci, h=alturas[k], nivel="CI", nodo=ci)
            enlaces.append((ci["id"], c["id"], "indirecta"))
            yy += alturas[k] + 0.13
    y_fin = y_ci0 + sum(alturas) + 0.13 * (len(alturas) - 1)

    for e in arbol["efectos_directos"]:
        enlaces.append((pc["id"], e["id"], "efecto"))
    for e in arbol["efectos_indirectos"]:
        for p in e["padres"]:
            enlaces.append((p, e["id"], "efecto"))
        enlaces.append((e["id"], ef["id"], "efecto"))
    return cajas, enlaces, y_fin


# ---------------------------------------------------------------- PowerPoint
def _rgb(h):
    return RGBColor.from_string(h)


def _estilo(nivel):
    return {
        "EF": (COL["ef_fondo"], COL["ef_fondo"], COL["ef_texto"]),
        "EI": (COL["ei_fondo"], COL["ei_borde"], COL["texto"]),
        "ED": (COL["ed_fondo"], COL["ed_borde"], COL["texto"]),
        "PC": (COL["pc_fondo"], COL["pc_fondo"], COL["pc_texto"]),
        "CD": (COL["verde_claro"], COL["borde_cd"], COL["texto"]),
        "CI": (COL["ci_fondo"], COL["ci_borde"], COL["texto"]),
        "EX": (COL["ex_fondo"], COL["ex_borde"], COL["gris"]),
    }[nivel]


def _texto(tf, partes, size, color, align=PP_ALIGN.CENTER):
    """partes: lista de (texto, bold, italic, size_rel)."""
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(PAD_X)
    tf.margin_top = tf.margin_bottom = Inches(PAD_Y)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    primero = True
    for parrafo in partes:
        p = tf.paragraphs[0] if primero else tf.add_paragraph()
        primero = False
        p.alignment = align
        for (t, b, it, rel) in parrafo:
            r = p.add_run()
            r.text = t
            f = r.font
            f.name = "Arial"
            f.size = Pt(size * rel)
            f.bold = b
            f.italic = it
            f.color.rgb = _rgb(color)


def _caja(slide, c, nombre):
    nivel = c["nivel"]
    fondo, borde, color = _estilo(nivel)
    forma = MSO_SHAPE.ROUNDED_RECTANGLE if nivel in ("PC", "EF") else MSO_SHAPE.RECTANGLE
    sh = slide.shapes.add_shape(forma, Inches(c["x"]), Inches(c["y"]), Inches(c["w"]), Inches(c["h"]))
    if forma == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = 0.12
    sh.name = nombre
    sh.fill.solid()
    sh.fill.fore_color.rgb = _rgb(fondo)
    sh.line.color.rgb = _rgb(borde)
    sh.line.width = Pt(1.75 if nivel in ("CD", "ED") else 1.0)
    if nivel == "EX":
        sh.line.dash_style = 4  # guion
    sh.shadow.inherit = False
    nd = c["nodo"]
    if nivel == "EX":
        partes = [[(nd["titulo"], True, False, 1.0)], [(nd["texto"], False, False, 1.0)]]
        _texto(sh.text_frame, partes, TAM[nivel], color, PP_ALIGN.LEFT)
    else:
        partes = [(nd["texto"], NEGRITA[nivel], False, 1.0)]
        if nivel not in ("PC", "EF"):
            partes.insert(0, (nd["id"] + "  ", True, False, 0.78))  # código para cruzar con la matriz de consistencia
        _texto(sh.text_frame, [partes], TAM[nivel], color,
               PP_ALIGN.CENTER if nivel in ("PC", "EF", "ED", "CD") else PP_ALIGN.LEFT)
    return sh


def _conector(slide, a, b, sa, sb, tipo, desde, hasta):
    """Conector acodado anclado: de la caja `desde` (punto sa) a la caja `hasta` (punto sb).

    Puntos de conexión del rectángulo: 0 arriba, 1 izquierda, 2 abajo, 3 derecha.
    a y b son las coordenadas (pulgadas) de los puntos de anclaje; la flecha apunta a `hasta`."""
    (ax, ay), (bx, by) = a, b
    if tipo == "indirecta":
        # Horizontal-vertical-horizontal, del borde izquierdo de la causa directa al de la indirecta:
        # el tramo vertical queda en la sangría y parece salir del borde inferior de la causa directa.
        cx = slide.shapes.add_connector(MSO_CONNECTOR.ELBOW, Inches(bx), Inches(by), Inches(ax), Inches(ay))
        cx.begin_connect(hasta, sb)
        cx.end_connect(desde, sa)
        # sin punta de flecha: corchete de agrupación
    else:
        # Vertical-horizontal-vertical desde b (arriba) hasta a (abajo): bentConnector3 rotado 90°
        cx = slide.shapes.add_connector(MSO_CONNECTOR.ELBOW, Inches(bx), Inches(by), Inches(ax), Inches(ay))
        cx.begin_connect(hasta, sb)
        cx.end_connect(desde, sa)
        sp = cx._element
        xfrm = sp.spPr.find(qn("a:xfrm"))
        w = ay - by
        h = bx - ax
        flip_v = h < 0
        h = abs(h)
        cxm, cym = (ax + bx) / 2, (ay + by) / 2
        xfrm.set("rot", "5400000")
        for k in ("flipH", "flipV"):
            if k in xfrm.attrib:
                del xfrm.attrib[k]
        if flip_v:
            xfrm.set("flipV", "1")
        off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
        off.set("x", str(int(Inches(cxm - w / 2))))
        off.set("y", str(int(Inches(cym - h / 2))))
        ext.set("cx", str(int(Inches(w))))
        ext.set("cy", str(int(Inches(h))))
    ln = cx.line
    ln.color.rgb = _rgb(COL["linea"])
    ln.width = Pt(1.25)
    if tipo == "exogeno":
        ln.dash_style = 4
    lnx = cx._element.spPr.find(qn("a:ln"))
    if tipo != "indirecta":
        punta = etree.SubElement(lnx, qn("a:headEnd"))
        punta.set("type", "triangle")
        punta.set("w", "med")
        punta.set("len", "med")
    return cx


def _anclas(c, lado):
    x, y, w, h = c["x"], c["y"], c["w"], c["h"]
    return {0: (x + w / 2, y), 1: (x, y + h / 2), 2: (x + w / 2, y + h), 3: (x + w, y + h / 2)}[lado]


def _banda(slide, arbol, meta):
    ban = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(ANCHO), Inches(0.82))
    ban.fill.solid()
    ban.fill.fore_color.rgb = _rgb(COL["verde"])
    ban.line.fill.background()
    ban.shadow.inherit = False
    ban.name = "Encabezado institucional"
    tf = ban.text_frame
    _texto(tf, [
        [("AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN  ·  DIME  ·  SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN - SESDI", True, False, 0.75)],
        [(arbol["titulo"], True, False, 1.35)],
        [(arbol["subtitulo"], False, False, 0.8)],
    ], 12, "FFFFFF", PP_ALIGN.LEFT)
    tf.margin_left = Inches(MARGEN)
    pie = slide.shapes.add_textbox(Inches(MARGEN), Inches(ALTO - 0.42), Inches(ANCHO - 2 * MARGEN), Inches(0.3))
    pie.name = "Pie de lámina"
    _texto(pie.text_frame, [[(meta["pie"] + "  ·  " + arbol.get("fuente", ""), False, False, 1.0)]], 7.5, COL["gris"], PP_ALIGN.LEFT)


def _leyenda(slide, y):
    items = [("EF", "Efecto final"), ("EI", "Efecto indirecto"), ("ED", "Efecto directo"), ("PC", "Problema central"),
             ("CD", "Causa directa"), ("CI", "Causa indirecta"), ("EX", "Factor exógeno de peligro (no controlable por el PI)")]
    x = MARGEN
    for nivel, txt in items:
        fondo, borde, _ = _estilo(nivel)
        sq = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y + 0.04), Inches(0.18), Inches(0.13))
        sq.fill.solid()
        sq.fill.fore_color.rgb = _rgb(fondo)
        sq.line.color.rgb = _rgb(borde)
        sq.shadow.inherit = False
        if nivel == "EX":
            sq.line.dash_style = 4
        ancho = _font(7.5, False).getlength(txt) / 300 + 0.15
        tb = slide.shapes.add_textbox(Inches(x + 0.22), Inches(y), Inches(ancho), Inches(0.22))
        _texto(tb.text_frame, [[(txt, False, False, 1.0)]], 7.5, COL["gris"], PP_ALIGN.LEFT)
        tb.text_frame.margin_left = tb.text_frame.margin_right = 0
        x += 0.22 + ancho + 0.2


def ajustar_escala(arbol):
    """Elige el mayor tamaño de letra (hasta +40 %) con el que el árbol cabe en la lámina."""
    for f in [1.4, 1.35, 1.3, 1.25, 1.2, 1.15, 1.1, 1.05, 1.0, 0.95, 0.9, 0.85]:
        for k, v in TAM_BASE.items():
            TAM[k] = round(v * f * 4) / 4
        _, _, y_fin = diagramar(arbol)
        if y_fin + 0.12 + 0.3 <= ALTO - 0.5:
            return f
    return f


def lamina_arbol(prs, arbol, meta):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    ajustar_escala(arbol)
    cajas, enlaces, y_fin = diagramar(arbol)
    _banda(slide, arbol, meta)
    formas = {}
    for k, c in cajas.items():
        formas[k] = _caja(slide, c, f"{k} · {c['nivel']}")
    conectores = []
    for hijo, padre, tipo in enlaces:
        ch, cp = cajas[hijo], cajas[padre]
        if tipo == "indirecta":
            sa, sb = 1, 1
        elif tipo == "exogeno":
            # del lateral del factor exógeno al lateral del problema central
            izquierda = ch["x"] < cp["x"]
            sa, sb = (3, 1) if izquierda else (1, 3)
            a, b = _anclas(ch, sa), _anclas(cp, sb)
            cx = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(a[0]), Inches(a[1]), Inches(b[0]), Inches(b[1]))
            cx.begin_connect(formas[hijo], sa)
            cx.end_connect(formas[padre], sb)
            cx.line.color.rgb = _rgb(COL["ex_borde"])
            cx.line.width = Pt(1.25)
            cx.line.dash_style = 4
            lnx = cx._element.spPr.find(qn("a:ln"))
            t = etree.SubElement(lnx, qn("a:tailEnd"))
            t.set("type", "triangle")
            conectores.append(cx)
            continue
        else:
            sa, sb = 0, 2  # del borde superior del hijo al inferior del padre
        a, b = _anclas(ch, sa), _anclas(cp, sb)
        conectores.append(_conector(slide, a, b, sa, sb, tipo, formas[hijo], formas[padre]))
    # Conectores detrás de las cajas (el tramo que pasa bajo una caja queda oculto)
    arbol_xml = slide.shapes._spTree
    for cx in conectores:
        el = cx._element
        arbol_xml.remove(el)
        arbol_xml.insert(2, el)
    _leyenda(slide, min(y_fin + 0.12, ALTO - 0.72))
    notas = slide.notes_slide.notes_text_frame
    notas.text = arbol.get("notas", "")
    return slide


def generar_pptx(datos, salida):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(ANCHO), Inches(ALTO)
    prs.core_properties.title = datos["meta"]["titulo"]
    prs.core_properties.author = "ANIN - DIME - SESDI"
    prs.core_properties.subject = "Árboles de problemas, causas y efectos - Estudio de preinversión"
    for clave in datos["orden"]:
        lamina_arbol(prs, datos["arboles"][clave], datos["meta"])
    prs.save(salida)


# ---------------------------------------------------------------- draw.io
def generar_drawio(datos, salida):
    PX = 96  # px por pulgada
    paginas = []
    for clave in datos["orden"]:
        arbol = datos["arboles"][clave]
        ajustar_escala(arbol)
        cajas, enlaces, _ = diagramar(arbol)
        celdas = ['<mxCell id="0"/>', '<mxCell id="1" parent="0"/>']
        titulo = escape(arbol["titulo"] + " — " + arbol["subtitulo"])
        celdas.append(
            f'<mxCell id="titulo" value="{titulo}" style="text;html=0;fontFamily=Arial;fontSize=16;fontStyle=1;'
            f'fontColor=#1B4F72;align=left;verticalAlign=middle;whiteSpace=wrap;" vertex="1" parent="1">'
            f'<mxGeometry x="{MARGEN*PX:.0f}" y="{0.25*PX:.0f}" width="{(ANCHO-2*MARGEN)*PX:.0f}" height="{0.6*PX:.0f}" as="geometry"/></mxCell>')
        for k, c in cajas.items():
            fondo, borde, color = _estilo(c["nivel"])
            nd = c["nodo"]
            valor = (nd["titulo"] + "\n" + nd["texto"]) if c["nivel"] == "EX" else (
                nd["texto"] if c["nivel"] in ("PC", "EF") else nd["id"] + "  " + nd["texto"])
            redondo = 1 if c["nivel"] in ("PC", "EF") else 0
            guion = "dashed=1;" if c["nivel"] == "EX" else ""
            bold = 1 if NEGRITA[c["nivel"]] else 0
            alin = "center" if c["nivel"] in ("PC", "EF", "ED", "CD") else "left"
            ancho_borde = 2 if c["nivel"] in ("CD", "ED") else 1
            celdas.append(
                f'<mxCell id="{escape(k)}" value="{escape(valor)}" style="rounded={redondo};whiteSpace=wrap;html=0;'
                f'fillColor=#{fondo};strokeColor=#{borde};strokeWidth={ancho_borde};fontColor=#{color};fontFamily=Arial;'
                f'fontSize={TAM[c["nivel"]]*96/72:.1f};fontStyle={bold};align={alin};spacingLeft=6;spacingRight=6;{guion}" '
                f'vertex="1" parent="1"><mxGeometry x="{c["x"]*PX:.1f}" y="{c["y"]*PX:.1f}" '
                f'width="{c["w"]*PX:.1f}" height="{c["h"]*PX:.1f}" as="geometry"/></mxCell>')
        for i, (hijo, padre, tipo) in enumerate(enlaces):
            base = "edgeStyle=orthogonalEdgeStyle;rounded=0;html=0;strokeColor=#4A5867;strokeWidth=1.5;"
            if tipo == "indirecta":
                est = base + "exitX=0;exitY=0.5;entryX=0;entryY=0.5;endArrow=none;"
                # el borde izquierdo de la causa indirecta une al borde inferior-izquierdo de la directa
                est = base + f"exitX=0;exitY=0.5;exitDx=0;exitDy=0;entryX={SANGRIA/2/cajas[padre]['w']:.4f};entryY=1;endArrow=none;"
            elif tipo == "exogeno":
                est = "endArrow=block;dashed=1;html=0;strokeColor=#6B7A89;strokeWidth=1.5;"
            else:
                est = base + "exitX=0.5;exitY=0;entryX=0.5;entryY=1;endArrow=block;endFill=1;"
            celdas.append(f'<mxCell id="e{i}" style="{est}" edge="1" parent="1" source="{escape(hijo)}" target="{escape(padre)}">'
                          f'<mxGeometry relative="1" as="geometry"/></mxCell>')
        modelo = (f'<mxGraphModel dx="1600" dy="1100" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" '
                  f'fold="1" page="1" pageScale="1" pageWidth="{ANCHO*PX:.0f}" pageHeight="{ALTO*PX:.0f}" math="0" shadow="0">'
                  f'<root>{"".join(celdas)}</root></mxGraphModel>')
        paginas.append(f'<diagram id="{clave}" name="{escape(arbol["pestana"])}">{modelo}</diagram>')
    xml = ('<?xml version="1.0" encoding="UTF-8"?>\n<mxfile host="app.diagrams.net" agent="ANIN-SESDI" version="24.7.17">'
           + "".join(paginas) + "</mxfile>\n")
    Path(salida).write_text(xml, encoding="utf-8")


if __name__ == "__main__":
    ruta = Path(sys.argv[1]) if len(sys.argv) > 1 else AQUI / "arboles_IN_Piura.json"
    datos = json.loads(ruta.read_text(encoding="utf-8"))
    generar_pptx(datos, AQUI / "Arboles_Problemas_IN_Piura.pptx")
    generar_drawio(datos, AQUI / "Arboles_Problemas_IN_Piura.drawio")
    print("OK:", AQUI / "Arboles_Problemas_IN_Piura.pptx", "y .drawio")
