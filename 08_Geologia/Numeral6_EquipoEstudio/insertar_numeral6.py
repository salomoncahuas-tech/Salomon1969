# -*- coding: utf-8 -*-
"""
Inserta el numeral 6. EQUIPO DE ESTUDIO (6.1 Cuadro de responsables por actividad
y 6.2 Organigrama del equipo de estudio) en los informes del Entregable 3 -
Estudio de Geologia del Proyecto IN Piura (ANIN - DIME - SESDI).

Reemplaza los marcadores de posicion existentes respetando la estructura de
titulos y el formato institucional del propio informe (verde ANIN #1B4D2E en
encabezados de tabla, leyendas en azul #1B4F72, cuerpo a 7.5 pt, filas alternas).

Uso:
    python3 insertar_numeral6.py <informe_entrada.docx> <organigrama.png> <informe_salida.docx>
"""
import sys, os, re
from docx import Document
from docx.shared import Emu
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from estilo import parrafo, leyenda, tabla, VERDE_ANIN, GRIS_ALT
from contenido6 import (PARRAFOS_61, LEYENDA_TABLA, ENCABEZADO_PROF, ENCABEZADO_ROL,
                        FILAS, NOTAS, LEYENDA_FIGURA, TEXTO_62)

# ---------------------------------------------------------------- utilidades
def bloques(doc):
    """Devuelve los hijos de <w:body> que son parrafos o tablas, en orden."""
    return [c for c in doc.element.body.iterchildren()
            if c.tag in (qn('w:p'), qn('w:tbl'))]

def texto_de(el, doc):
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    if el.tag == qn('w:p'):
        return Paragraph(el, doc).text.strip()
    return Table(el, doc).rows[0].cells[0].text.strip()

RE_INI = re.compile(r'^6\s*[.\-)]?\s*EQUIPO\s+DE\s+ESTUDIO', re.I)
RE_FIN = re.compile(r'^7\s*[.\-)]\s')

def es_titulo(el, doc):
    from docx.text.paragraph import Paragraph
    if el.tag != qn('w:p'):
        return False
    st = (Paragraph(el, doc).style.name or '')
    return st.startswith('Heading') or st.startswith('Ttulo') or st.startswith('Título')

def rango_numeral6(doc):
    """Localiza [inicio, fin): desde el titulo del numeral 6 hasta el titulo 7
    (exclusivo). Ignora la entrada homonima del indice de contenidos."""
    els = bloques(doc)
    candidatos = [i for i, el in enumerate(els)
                  if es_titulo(el, doc) and RE_INI.match(texto_de(el, doc))]
    if not candidatos:
        raise SystemExit("ERROR: no se encontro el titulo '6. EQUIPO DE ESTUDIO'. "
                         "Revise que el informe conserve ese encabezado.")
    ini = candidatos[-1]          # el ultimo: el del cuerpo, no el del indice
    fin = len(els)
    for i in range(ini + 1, len(els)):
        if es_titulo(els[i], doc) and RE_FIN.match(texto_de(els[i], doc)):
            fin = i
            break
    else:
        raise SystemExit("ERROR: no se encontro el titulo '7.' que cierra el numeral 6.")
    return els, ini, fin

def estilo_de(doc, *nombres):
    for n in nombres:
        try:
            return doc.styles[n]
        except KeyError:
            continue
    return None

def titulo(doc, texto, nivel):
    """Parrafo de titulo usando el estilo Heading real del documento."""
    st = estilo_de(doc, f'Heading {nivel}', f'Ttulo {nivel}', f'Título {nivel}')
    p = OxmlElement('w:p')
    if st is not None:
        pPr = OxmlElement('w:pPr')
        pStyle = OxmlElement('w:pStyle'); pStyle.set(qn('w:val'), st.style_id)
        pPr.append(pStyle); p.append(pPr)
    r = OxmlElement('w:r'); t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve'); t.text = texto
    r.append(t); p.append(r)
    return p

def parrafo_imagen(doc, ruta_png, ancho_emu):
    """Parrafo centrado con la imagen incrustada, escalada a ancho_emu."""
    from docx.text.paragraph import Paragraph
    p = doc.add_paragraph()
    p.alignment = 1  # centrado
    run = p.add_run()
    run.add_picture(ruta_png, width=Emu(ancho_emu))
    el = p._p
    el.getparent().remove(el)   # lo sacamos del final; se reinsertara en su sitio
    return el

# ---------------------------------------------------------------- construccion
def construir_numeral6(doc, ruta_png, ancho_util):
    """Devuelve la lista de elementos XML que conforman el numeral 6 completo."""
    # Anchos en dxa (twips). ancho_util viene en dxa.
    n_prof = len(ENCABEZADO_PROF)
    col_prof = 620                          # columnas de profesionales
    col_act = ancho_util - col_prof * n_prof  # columna de actividades/tareas
    anchos = [col_act] + [col_prof] * n_prof

    out = []
    out.append(titulo(doc, '6. EQUIPO DE ESTUDIO', 1))
    out.append(titulo(doc, '6.1. Cuadro de responsables por actividad', 2))
    for tx in PARRAFOS_61:
        out.append(parrafo(tx, sz=17, jc='both', after=120))
    out.append(leyenda(LEYENDA_TABLA))

    filas = []
    # Fila 1: banda de titulo
    filas.append([
        {'texto': 'ESTUDIO DE GEOLOGÍA', 'bold': True, 'color': 'FFFFFF',
         'fill': VERDE_ANIN, 'jc': 'center', 'sz': 14},
        {'texto': 'PROFESIONAL / ESPECIALIDAD', 'bold': True, 'color': 'FFFFFF',
         'fill': VERDE_ANIN, 'jc': 'center', 'sz': 14, 'gridspan': n_prof},
    ])
    # Fila 2: nombres y especialidad
    filas.append([{'texto': 'ACTIVIDADES Y TAREAS DEL ENTREGABLE', 'bold': True,
                   'color': 'FFFFFF', 'fill': VERDE_ANIN, 'jc': 'center', 'sz': 14}] +
                 [{'texto': p, 'bold': True, 'fill': 'E8EFE9', 'jc': 'center', 'sz': 12}
                  for p in ENCABEZADO_PROF])
    # Fila 3: rol segun el flujograma
    filas.append([{'texto': 'Rol en el entregable, según el flujograma del Entregable 3 →',
                   'bold': True, 'fill': 'E8EFE9', 'sz': 12}] +
                 [{'texto': r, 'fill': 'E8EFE9', 'jc': 'center', 'sz': 12}
                  for r in ENCABEZADO_ROL])
    # Filas de actividades y tareas
    alterna = 0
    for etiqueta, marcas in FILAS:
        if marcas is None:      # banda de ACTIVIDAD
            filas.append([{'texto': etiqueta, 'bold': True, 'color': 'FFFFFF',
                           'fill': '2E6B45', 'sz': 13, 'gridspan': n_prof + 1}])
            alterna = 0
            continue
        fill = GRIS_ALT if alterna % 2 else None
        alterna += 1
        assert len(marcas) == n_prof, f"marcas invalidas en: {etiqueta[:40]}"
        filas.append([{'texto': etiqueta, 'sz': 13, 'fill': fill, 'jc': 'both'}] +
                     [{'texto': ('X' if m == 'X' else ''), 'bold': (m == 'X'),
                       'jc': 'center', 'sz': 14, 'fill': fill} for m in marcas])
    out.append(tabla(anchos, filas, header_rows=3))

    for n in NOTAS:
        out.append(parrafo(n, sz=13, jc='both', before=40, after=40, italic=True))

    # ---- 6.2
    out.append(titulo(doc, '6.2. Organigrama del equipo de estudio', 2))
    out.append(parrafo(TEXTO_62, sz=17, jc='both', after=120))
    out.append(parrafo_imagen(doc, ruta_png, int(ancho_util * 635)))  # dxa -> EMU
    out.append(leyenda(LEYENDA_FIGURA))
    return out

# ---------------------------------------------------------------- principal
def main():
    if len(sys.argv) != 4:
        raise SystemExit(__doc__)
    entrada, png, salida = sys.argv[1:4]
    for f in (entrada, png):
        if not os.path.exists(f):
            raise SystemExit(f"ERROR: no existe {f}")

    doc = Document(entrada)
    sec = doc.sections[0]
    ancho_util = int((sec.page_width - sec.left_margin - sec.right_margin) / 635)  # EMU->dxa

    els, ini, fin = rango_numeral6(doc)
    print(f"Numeral 6 localizado: {fin - ini} elementos a reemplazar "
          f"(posiciones {ini}..{fin - 1}).")

    nuevos = construir_numeral6(doc, png, ancho_util)

    ancla = els[ini]
    for el in nuevos:                 # insertar el bloque nuevo antes del titulo actual
        ancla.addprevious(el)
    for el in els[ini:fin]:           # y retirar el contenido anterior (marcadores)
        el.getparent().remove(el)

    doc.save(salida)
    print(f"OK -> {salida}  ({os.path.getsize(salida):,} bytes)")

if __name__ == '__main__':
    main()
