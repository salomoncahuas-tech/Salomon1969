# -*- coding: utf-8 -*-
"""Utilidades de estilo institucional ANIN para los informes del Entregable 3 - Geologia."""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

VERDE_ANIN = "1B4D2E"   # encabezado de tabla
AZUL_ANIN  = "1B4F72"   # leyendas (Tabla / Figura)
BORDE_EXT  = "7F9E8A"
BORDE_INT  = "BFCFC2"
GRIS_ALT   = "F2F2F2"

def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn('w:' + k), v)
    return e

def run(texto, sz=15, bold=False, color="000000", italic=False):
    r = OxmlElement('w:r'); rPr = OxmlElement('w:rPr')
    if bold:
        rPr.append(OxmlElement('w:b')); rPr.append(OxmlElement('w:bCs'))
    if italic:
        rPr.append(OxmlElement('w:i')); rPr.append(OxmlElement('w:iCs'))
    rPr.append(_el('w:color', val=color))
    rPr.append(_el('w:sz', val=str(sz))); rPr.append(_el('w:szCs', val=str(sz)))
    r.append(rPr)
    t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = texto
    r.append(t)
    return r

def parrafo(texto, sz=15, bold=False, color="000000", jc=None, before=0, after=80,
            keepNext=False, italic=False):
    p = OxmlElement('w:p'); pPr = OxmlElement('w:pPr')
    if keepNext:
        pPr.append(OxmlElement('w:keepNext'))
    pPr.append(_el('w:spacing', before=str(before), after=str(after)))
    if jc:
        pPr.append(_el('w:jc', val=jc))
    p.append(pPr)
    if texto:
        p.append(run(texto, sz=sz, bold=bold, color=color, italic=italic))
    return p

def leyenda(texto):
    """Leyenda de Tabla/Figura: negrita, azul institucional, 7.5 pt."""
    return parrafo(texto, sz=15, bold=True, color=AZUL_ANIN, before=160, after=60, keepNext=True)

def _tc(texto, ancho, sz=15, bold=False, color="000000", fill=None, jc=None,
        vmerge=None, gridspan=None):
    tc = OxmlElement('w:tc'); tcPr = OxmlElement('w:tcPr')
    tcPr.append(_el('w:tcW', w=str(ancho), type='dxa'))
    if gridspan:
        tcPr.append(_el('w:gridSpan', val=str(gridspan)))
    if vmerge is not None:
        tcPr.append(_el('w:vMerge', val=vmerge) if vmerge else OxmlElement('w:vMerge'))
    if fill:
        tcPr.append(_el('w:shd', val='clear', color='auto', fill=fill))
    mar = OxmlElement('w:tcMar')
    for lado, w in (('top', '60'), ('left', '90'), ('bottom', '60'), ('right', '90')):
        mar.append(_el('w:' + lado, w=w, type='dxa'))
    tcPr.append(mar)
    tcPr.append(_el('w:vAlign', val='center'))
    tc.append(tcPr)
    tc.append(parrafo(texto, sz=sz, bold=bold, color=color, jc=jc, before=20, after=20))
    return tc

def tabla(anchos, filas, header_rows=1):
    """filas: lista de listas de dicts {texto, bold, color, fill, jc, gridspan, vmerge}."""
    tbl = OxmlElement('w:tbl'); tblPr = OxmlElement('w:tblPr')
    tblPr.append(_el('w:tblW', w=str(sum(anchos)), type='dxa'))
    b = OxmlElement('w:tblBorders')
    for lado, sz_, col in (('top', '4', BORDE_EXT), ('left', '4', BORDE_EXT),
                           ('bottom', '4', BORDE_EXT), ('right', '4', BORDE_EXT),
                           ('insideH', '2', BORDE_INT), ('insideV', '2', BORDE_INT)):
        b.append(_el('w:' + lado, val='single', sz=sz_, space='0', color=col))
    tblPr.append(b)
    cm = OxmlElement('w:tblCellMar')
    cm.append(_el('w:left', w='10', type='dxa')); cm.append(_el('w:right', w='10', type='dxa'))
    tblPr.append(cm)
    tblPr.append(_el('w:tblLook', val='0000', firstRow='0', lastRow='0',
                     firstColumn='0', lastColumn='0', noHBand='0', noVBand='0'))
    tbl.append(tblPr)
    grid = OxmlElement('w:tblGrid')
    for a in anchos:
        grid.append(_el('w:gridCol', w=str(a)))
    tbl.append(grid)
    for i, fila in enumerate(filas):
        tr = OxmlElement('w:tr')
        if i < header_rows:
            trPr = OxmlElement('w:trPr'); trPr.append(OxmlElement('w:tblHeader')); tr.append(trPr)
        j = 0
        for celda in fila:
            gs = celda.get('gridspan')
            ancho = sum(anchos[j:j + (gs or 1)])
            tr.append(_tc(celda.get('texto', ''), ancho,
                          sz=celda.get('sz', 15), bold=celda.get('bold', False),
                          color=celda.get('color', '000000'), fill=celda.get('fill'),
                          jc=celda.get('jc'), gridspan=gs, vmerge=celda.get('vmerge')))
            j += (gs or 1)
        tbl.append(tr)
    return tbl
