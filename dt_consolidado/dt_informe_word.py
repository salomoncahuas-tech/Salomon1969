# -*- coding: utf-8 -*-
"""
Proyecto IN Piura — CUI 2669244 · ANIN / DIME / SESDI
Generador del Informe Tecnico Consolidado de Diagnostico Territorial,
una version por provincia, desagregada por distritos.

Estructura homologada con los informes precedentes de Morropon y de Frias
(Ayabaca), ampliada con la integracion del Indice de Susceptibilidad
Litologica (ISL) del E3 Estudio de Geologia y con el inventario de carcavas.
"""
import os
from collections import defaultdict

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

import dt_data as D

VERDE = RGBColor(0x1B, 0x4D, 0x2E)
AZUL = RGBColor(0x1B, 0x4F, 0x72)
GRIS = RGBColor(0x44, 0x44, 0x44)
VERDE_HEX = "1B4D2E"
GRIS_HEX = "F2F5F3"

NOMBRE_PROYECTO = ("«Recuperación del servicio de regulación de riesgos naturales y "
                   "recuperación de ecosistemas degradados en la Cuenca Alta del Río Piura»")


# ------------------------------------------------------------------ utilidades

def _estilos(doc):
    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for nivel, size, color in ((1, 14, VERDE), (2, 12, AZUL), (3, 11, AZUL)):
        s = doc.styles[f"Heading {nivel}"]
        s.font.name = "Arial"
        s.font.size = Pt(size)
        s.font.bold = True
        s.font.color.rgb = color
        s.paragraph_format.space_before = Pt(14 if nivel == 1 else 10)
        s.paragraph_format.space_after = Pt(6)


def _sombra(celda, hexcolor):
    tcPr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def _p(doc, texto, size=10, bold=False, italic=False, color=None,
       align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(texto)
    r.font.name = "Arial"
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    if color is not None:
        r.font.color.rgb = color
    return p


def _vineta(doc, texto, size=10):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(texto)
    r.font.name = "Arial"
    r.font.size = Pt(size)
    return p


def _tabla(doc, cabeceras, filas, titulo=None, fuente=None, size=8, anchos=None):
    if titulo:
        _p(doc, titulo, size=9, bold=True, color=AZUL, space_after=3)
    t = doc.add_table(rows=1, cols=len(cabeceras))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(cabeceras):
        hdr[i].text = ""
        p = hdr[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(str(h))
        r.font.name = "Arial"
        r.font.size = Pt(size)
        r.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        _sombra(hdr[i], VERDE_HEX)
    for k, fila in enumerate(filas):
        cells = t.add_row().cells
        for i, v in enumerate(fila):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.alignment = (WD_ALIGN_PARAGRAPH.RIGHT if isinstance(v, (int, float))
                           else WD_ALIGN_PARAGRAPH.LEFT)
            txt = f"{v:,.2f}" if isinstance(v, float) else str(v)
            r = p.add_run(txt)
            r.font.name = "Arial"
            r.font.size = Pt(size)
            if isinstance(fila[0], str) and fila[0].startswith(("TOTAL", "PROMEDIO")):
                r.bold = True
            if k % 2 == 1:
                _sombra(cells[i], GRIS_HEX)
    if anchos:
        for i, w in enumerate(anchos):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    if fuente:
        _p(doc, fuente, size=7, italic=True, color=GRIS, space_after=10)
    return t


def _portada(doc, prov, r, bs):
    for txt, size, bold, color in (
            ("AUTORIDAD NACIONAL DE INFRAESTRUCTURA — ANIN", 13, True, VERDE),
            ("DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA — DIME", 10, False, VERDE),
            ("SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN — SESDI", 10, False, VERDE)):
        _p(doc, txt, size=size, bold=bold, color=color,
           align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    doc.add_paragraph()
    _p(doc, "INFORME TÉCNICO CONSOLIDADO DE DIAGNÓSTICO TERRITORIAL", size=17, bold=True,
       color=VERDE, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    _p(doc, "Bloques preliminares de intervención verificados", size=12,
       align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)
    _p(doc, f"PROVINCIA DE {D.PROV_META[prov]['nombre'].upper()}", size=20, bold=True,
       color=AZUL, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    distritos = ", ".join(k for k, _ in D.distritos_de(bs))
    _p(doc, f"{r['distritos']} distrito{'s' if r['distritos'] > 1 else ''} · "
            f"{r['n']} bloques · {r['ha']:,.2f} ha", size=12, bold=True,
       align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    _p(doc, distritos, size=10, italic=True, color=GRIS,
       align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
    _p(doc, "Proyecto de Inversión IN PIURA · CUI 2669244", size=12, bold=True,
       align=WD_ALIGN_PARAGRAPH.CENTER, space_after=3)
    _p(doc, NOMBRE_PROYECTO, size=10, italic=True,
       align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
    _p(doc, "Fase de preinversión — Estudio de Perfil · Sistema Invierte.pe (DGPMI-MEF)",
       size=9, color=GRIS, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    _p(doc, "Sistema de referencia: UTM WGS 84 Zona 17S (EPSG:32717)", size=9, color=GRIS,
       align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    _p(doc, "Financiamiento: Administración Directa", size=9, color=GRIS,
       align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()


# ------------------------------------------------------------------ secciones

def _s1(doc, prov, r, bs, meta):
    doc.add_heading("1. Objeto, alcance y encuadre normativo", level=1)
    _p(doc, f"El presente informe consolida y sistematiza el Diagnóstico Territorial (DT) de "
            f"los {r['n']} bloques preliminares de intervención de la provincia de "
            f"{meta['nombre']}, en el marco de la formulación del Proyecto de Inversión "
            f"{NOMBRE_PROYECTO} (CUI 2669244), fase de preinversión — Estudio de Perfil.")
    _p(doc, "El documento se elabora como insumo directo del Capítulo de Identificación del "
            "Perfil, conforme a los Contenidos Mínimos del Anexo N.° 1 de la Directiva General "
            "del Sistema Nacional de Programación Multianual y Gestión de Inversiones "
            "(Invierte.pe) y a la Guía General para la Identificación, Formulación y Evaluación "
            "de Proyectos de Inversión (DGPMI-MEF, 2022), en particular su Anexo 2 sobre "
            "Gestión del Riesgo en contexto de Cambio Climático (GdR-CCC).")

    doc.add_heading("1.1 Alcance del análisis", level=2)
    distritos = ", ".join(k for k, _ in sorted(D.distritos_de(bs)))
    _vineta(doc, f"Ámbito: provincia de {meta['nombre']}, {r['distritos']} distrito"
                 f"{'s' if r['distritos'] > 1 else ''} ({distritos}).")
    _vineta(doc, f"Universo analizado: {r['n']} bloques preliminares de intervención, "
                 f"{r['ha']:,.2f} ha, correspondientes al catálogo maestro «Bloques V5/V6».")
    _vineta(doc, f"Base documental: informes técnicos de Diagnóstico Territorial por bloque y "
                 f"las {r['n']} plantillas de resumen validadas (Plantilla DT Campo Check "
                 f"Validada V5, versión V6 con MSAVI), con cobertura del 100 % de los bloques "
                 f"del ámbito.")
    _vineta(doc, f"Integración sectorial: {meta['volumen_geologia']}, del que se incorpora el "
                 f"Índice de Susceptibilidad Litológica (ISL-MM e ISL-EH) de los {r['n']} "
                 f"bloques; e inventario de cárcavas codificadas.")
    _vineta(doc, "Sistema de referencia: UTM WGS 84 Zona 17S (EPSG:32717).")
    _vineta(doc, "Marco del indicador de brecha: Resolución Ministerial N.° 00213-2024-MINAM, "
                 "«Porcentaje de superficie de ecosistemas degradados que brindan servicios "
                 "ecosistémicos que requieren de recuperación».")

    doc.add_heading("1.2 Naturaleza de la Unidad Productora", level=2)
    _p(doc, "La Unidad Productora (UP) de este proyecto no es una infraestructura convencional. "
            "La UP es el ecosistema en su totalidad: bosques estacionalmente secos de colina y "
            "montaña, matorrales, pastizales y ecosistemas de transición andino-costeros, junto "
            "con sus suelos, su cobertura vegetal, su sistema de drenaje natural y su red de "
            "quebradas. Los bloques preliminares de intervención son recortes operativos "
            "situados dentro de esa UP, seleccionados mediante criterios de exclusión e "
            "idoneidad; no son la UP en sí.")
    _p(doc, "Esta precisión tiene consecuencia metodológica directa: el análisis de riesgo se "
            "aplica tanto al ecosistema existente como a las medidas de reducción del riesgo en "
            "contexto de cambio climático (MRR-CCC) que se propongan, las cuales también deben "
            "evaluarse por exposición y fragilidad.")


def _s2(doc, prov, r, bs, meta):
    doc.add_heading("2. Metodología y fuentes", level=1)
    _p(doc, "El diagnóstico integra información de tres naturalezas explícitamente "
            "diferenciadas: registro de campo, producto de gabinete y catálogo oficial. La "
            "diferenciación no es formal: determina qué afirmaciones del informe admiten uso "
            "para fijar metas físicas y cuáles conservan carácter de hipótesis de trabajo.")
    _tabla(doc,
           ["Fuente", "Naturaleza", "Variables que aporta", "Tratamiento en el informe"],
           [["Catálogo maestro Bloques V5/V6", "Catálogo oficial",
             "Código, microcuenca, superficie, centroide, distrito",
             "Valor de referencia. Prevalece sobre la declaración de la ficha en caso de discrepancia."],
            ["Estadística zonal sobre el MDE", "Producto de gabinete",
             "Altitud mínima y máxima, amplitud, pendiente promedio",
             "Valor de bloque. No comparable con los perfiles de transecto de campo."],
            ["MSAVI 2024 y NDVI mediana 2025 (Sentinel-2)", "Producto de gabinete",
             "Vigor y densidad de la cobertura, distribución areal por clase",
             "Mide biomasa, no integridad ecosistémica. Lectura sujeta a la advertencia de estacionalidad."],
            ["Fichas F-DT-01 a F-DT-05 (Plantilla V5)", "Registro de campo",
             "Fisiografía, suelo, ecosistema, causas de degradación, hidrología, accesibilidad",
             "Registro primario. Se reproduce sin homologar la redacción entre brigadas."],
            [meta["volumen_geologia"], "Producto de gabinete (fuente secundaria)",
             "ISL-MM e ISL-EH, unidades geológicas por bloque",
             "Jerarquización del sustrato. No constituye evaluación de peligro ni de riesgo."],
            ["Capa «Cárcavas codificadas» y su caracterización", "Producto de gabinete",
             "Localización, longitud, morfometría y NDVI de cada cárcava",
             "Inventario parcial. La ausencia de registros no acredita ausencia de cárcavas."],
            ["Cruce INEI — Bloques V5", "Catálogo oficial",
             "Centros poblados asociados a cada bloque",
             "Base del diagnóstico social y del cálculo de población beneficiaria."]],
           titulo="Tabla 1. Fuentes de información utilizadas y su tratamiento",
           fuente="Fuente: elaboración propia sobre el acervo documental del Proyecto IN Piura.",
           anchos=[3.6, 2.6, 4.6, 6.2])
    _p(doc, "El procesamiento consistió en: (i) extracción y normalización de los campos de las "
            "cinco fichas F-DT desde las plantillas de resumen validadas de cada bloque; "
            "(ii) construcción de una base de datos con las variables geoespaciales del "
            "catálogo; (iii) incorporación del ISL por bloque desde el estudio de geología; "
            "(iv) incorporación del inventario de cárcavas; (v) contraste cruzado entre valores "
            "de campo, de gabinete y de catálogo; y (vi) cálculo de indicadores de completitud, "
            "concordancia y dispersión.")
    _p(doc, "PRINCIPIO DE NO INVENCIÓN DE DATOS. No se ha imputado ni estimado ningún valor "
            "ausente. Todo campo no reportado se conserva con la marca de origen «Por "
            "determinar» o «Por verificar» y se registra en la sección de vacíos de "
            "información. Las discrepancias detectadas entre fuentes no se han resuelto por "
            "criterio del redactor: se documentan, se declara el valor adoptado y se indica la "
            "acción requerida.", bold=True)


def _s3(doc, prov, r, bs, meta):
    doc.add_heading("3. Caracterización general del ámbito", level=1)
    doc.add_heading("3.1 Distribución de los bloques por distrito", level=2)
    dist = D.distritos_de(bs)
    filas = []
    for nombre, a in dist:
        sub = [b for b in bs if b["distrito"] == nombre]
        cd = D.con_dato(sub, "alt_min")
        rango = (f"{min(x['alt_min'] for x in cd):,.0f} – {max(x['alt_max'] for x in cd):,.0f}"
                 if cd else "s/d")
        pp = D.ponderada(sub, "pend_pct")
        filas.append([nombre, a["n"], round(a["ha"], 2), round(100 * a["ha"] / r["ha"], 2),
                      rango, round(pp, 2) if pp is not None else "s/d",
                      round(D.ponderada(sub, "msavi"), 4),
                      round(100 * sum(x["ha_bajo_umbral"] or 0 for x in sub) / a["ha"], 2)])
    filas.append(["TOTAL / MEDIA PONDERADA", r["n"], round(r["ha"], 2), 100.0,
                  f"{r['alt_min']:,.0f} – {r['alt_max']:,.0f}",
                  round(r["pend_ponderada"], 2), round(r["msavi_ponderado"], 4),
                  round(r["pct_brecha"], 2)])
    _tabla(doc, ["Distrito", "N.° de bloques", "Superficie (ha)", "% del ámbito",
                 "Rango altitudinal (msnm)", "Pendiente ponderada (%)", "MSAVI 2024 ponderado",
                 "% bajo umbral MSAVI"], filas,
           titulo=f"Tabla 2. Síntesis de bloques preliminares por distrito "
                  f"(provincia de {meta['nombre']})",
           fuente="Fuente: elaboración propia sobre el catálogo maestro Bloques V5/V6, la "
                  "estadística zonal del MDE y los compuestos Sentinel-2. Las medias se "
                  "ponderan por la superficie de cada bloque.",
           anchos=[3.4, 1.8, 2.2, 1.8, 2.8, 2.2, 2.2, 2.2])

    # Lectura de la asimetria distrital
    top = dist[0]
    areas = sorted((b["area_ha"], b["codigo"], b["distrito"]) for b in bs)
    med = D.mediana([b["area_ha"] for b in bs])
    if len(dist) > 1:
        acum, n_conc = 0.0, 0
        for _, a in dist:
            acum += a["ha"]
            n_conc += 1
            if acum >= 0.75 * r["ha"]:
                break
        _p(doc, f"La distribución es asimétrica. {n_conc} de los {len(dist)} distritos "
                f"concentran {acum:,.2f} ha, equivalentes al {100 * acum / r['ha']:.1f} % de la "
                f"superficie del ámbito provincial. El distrito de {top[0]} aporta por sí solo "
                f"{top[1]['ha']:,.2f} ha ({100 * top[1]['ha'] / r['ha']:.1f} %) en "
                f"{top[1]['n']} bloques.")
    _p(doc, f"El tamaño de los bloques varía entre {areas[0][0]:,.3f} ha (bloque "
            f"{areas[0][1]}, {areas[0][2]}) y {areas[-1][0]:,.2f} ha (bloque {areas[-1][1]}, "
            f"{areas[-1][2]}), con una mediana de {med:,.2f} ha. Esta dispersión es "
            f"determinante para el diseño de metas físicas y para el costeo unitario de las "
            f"MRR-CCC: el costo por hectárea de habilitación, transporte de plantones y "
            f"vigilancia no escala linealmente con la superficie del bloque.")

    doc.add_heading("3.2 Organización por microcuenca", level=2)
    _p(doc, f"Los {r['n']} bloques se distribuyen en {r['microcuencas']} microcuencas. La "
            f"microcuenca es la unidad funcional de planificación de las MRR-CCC: los bloques "
            f"que la comparten están sujetos al mismo régimen hídrico y a condicionantes de "
            f"ladera comunes, de modo que las intervenciones de infiltración, control de "
            f"escorrentía y estabilización deben dimensionarse en conjunto y no bloque a bloque.")
    agg = defaultdict(lambda: {"n": 0, "ha": 0.0, "bl": []})
    for b in bs:
        agg[b["microcuenca"]]["n"] += 1
        agg[b["microcuenca"]]["ha"] += b["area_ha"]
        agg[b["microcuenca"]]["bl"].append(b)
    filas = []
    for mc, a in sorted(agg.items(), key=lambda kv: -kv[1]["ha"]):
        bl = a["bl"]
        filas.append([mc, " · ".join(sorted({x["distrito"] for x in bl})), a["n"],
                      round(a["ha"], 2), round(100 * a["ha"] / r["ha"], 2),
                      round(sum(x["msavi"] * x["area_ha"] for x in bl) / a["ha"], 4),
                      round(100 * sum(x["ha_bajo_umbral"] or 0 for x in bl) / a["ha"], 2)])
    filas.append(["TOTAL", "", r["n"], round(r["ha"], 2), 100.0,
                  round(r["msavi_ponderado"], 4), round(r["pct_brecha"], 2)])
    _tabla(doc, ["Microcuenca", "Distrito(s)", "N.° de bloques", "Superficie (ha)",
                 "% del ámbito", "MSAVI ponderado", "% bajo umbral MSAVI"], filas,
           titulo="Tabla 3. Agregación de los bloques por microcuenca",
           fuente="Fuente: elaboración propia sobre el catálogo maestro Bloques V5/V6.",
           anchos=[2.8, 4.4, 1.8, 2.2, 1.8, 2.2, 2.4])


def _s4(doc, prov, r, bs, meta):
    doc.add_heading("4. Análisis geoespacial de los bloques", level=1)

    doc.add_heading("4.1 Rango altitudinal y pisos ecológicos", level=2)
    cda = D.con_dato(bs, "alt_min")
    lo = min(cda, key=lambda b: b["alt_min"])
    hi = max(cda, key=lambda b: b["alt_max"])
    amp = [b["amplitud"] for b in bs if b["amplitud"] is not None]
    _p(doc, f"El ámbito se despliega sobre un gradiente altitudinal de {r['alt_min']:,.0f} a "
            f"{r['alt_max']:,.0f} msnm, es decir, {r['alt_max'] - r['alt_min']:,.0f} m de "
            f"desnivel entre el bloque más bajo (bloque {lo['codigo']}, {lo['distrito']}) y el "
            f"más alto (bloque {hi['codigo']}, {hi['distrito']}). La amplitud altitudinal "
            f"interna media por bloque es de {sum(amp) / len(amp):,.0f} m "
            f"(mínimo {min(amp):,.0f} m, máximo {max(amp):,.0f} m).")
    filas = [[k, v["n"], round(v["ha"], 2), round(100 * v["ha"] / r["ha"], 2)]
             for k, v in D.conteo(bs, "piso")]
    filas.append(["TOTAL", r["n"], round(r["ha"], 2), 100.0])
    _tabla(doc, ["Piso altitudinal dominante", "N.° de bloques", "Superficie (ha)",
                 "% del ámbito"], filas,
           titulo="Tabla 4. Distribución de bloques por piso altitudinal dominante",
           fuente="Fuente: elaboración propia sobre la estadística zonal del MDE.",
           anchos=[7.0, 2.4, 2.6, 2.4])
    anchos_grandes = [b for b in bs if b["amplitud"] and b["amplitud"] >= 600]
    if anchos_grandes:
        cods = ", ".join(b["codigo"] for b in sorted(anchos_grandes,
                                                     key=lambda x: -x["amplitud"])[:8])
        _p(doc, f"{len(anchos_grandes)} bloques presentan una amplitud altitudinal interna igual "
                f"o superior a 600 m ({cods}{'…' if len(anchos_grandes) > 8 else ''}). Un bloque "
                f"que atraviesa más de un piso altitudinal no admite un elenco florístico ni una "
                f"prescripción de manejo únicos: requiere zonificación interna por franja de "
                f"cota antes de fijar metas físicas y seleccionar especies.")

    doc.add_heading("4.2 Pendiente media y clases de terreno", level=2)
    cp = D.con_dato(bs, "pend_pct")
    pmin = min(cp, key=lambda b: b["pend_pct"])
    pmax = max(cp, key=lambda b: b["pend_pct"])
    _p(doc, f"La pendiente media derivada del MDE para los {len(cp)} bloques con dato "
            f"disponible es de "
            f"{r['pend_media']:.2f} % (mediana {D.mediana([b['pend_pct'] for b in cp]):.2f} %), "
            f"con un mínimo de {pmin['pend_pct']:.2f} % (bloque {pmin['codigo']}, "
            f"{pmin['distrito']}) y un máximo de {pmax['pend_pct']:.2f} % (bloque "
            f"{pmax['codigo']}, {pmax['distrito']}). Ningún bloque del ámbito supera el "
            f"criterio de idoneidad de pendiente máxima del proyecto (75 %)."
            if pmax["pend_pct"] <= 75 else
            f"La pendiente media derivada del MDE es de {r['pend_media']:.2f} %, con un máximo "
            f"de {pmax['pend_pct']:.2f} % (bloque {pmax['codigo']}, {pmax['distrito']}), que "
            f"supera el criterio de idoneidad de pendiente máxima del proyecto (75 %).")
    filas = [[k, v["n"], round(v["ha"], 2), round(100 * v["ha"] / r["ha"], 2)]
             for k, v in D.conteo(bs, "clase_pendiente")]
    filas.append(["TOTAL", r["n"], round(r["ha"], 2), 100.0])
    _tabla(doc, ["Clase de pendiente (estadística zonal)", "N.° de bloques",
                 "Superficie (ha)", "% del ámbito"], filas,
           titulo="Tabla 5. Distribución de bloques por clase de pendiente media",
           fuente="Fuente: elaboración propia sobre la estadística zonal del MDE. La pendiente "
                  "se expresa en porcentaje.",
           anchos=[7.0, 2.4, 2.6, 2.4])
    discord = [b for b in bs if D.norm(b["clase_pendiente"]) != D.norm(b["pendiente_campo"])]
    _p(doc, f"En {len(discord)} de los {r['n']} bloques la clase de pendiente derivada de la "
            f"estadística zonal no coincide con el rango declarado por la brigada en la ficha "
            f"F-DT-01. La discrepancia no invalida ninguno de los dos registros: el valor zonal "
            f"se calcula sobre la totalidad del polígono y el de campo describe la traza "
            f"efectivamente recorrida, habitualmente la de menor pendiente por ser la "
            f"transitable. El detalle bloque a bloque figura en la hoja «4. Pendientes» del "
            f"anexo de matrices.")

    doc.add_heading("4.3 Lectura integrada de los índices de vegetación MSAVI y NDVI", level=2)
    mmin = min(bs, key=lambda b: b["msavi"])
    mmax = max(bs, key=lambda b: b["msavi"])
    _p(doc, f"El MSAVI 2024 medio de los {r['n']} bloques es de {r['msavi_medio']:.4f} (media "
            f"simple), {r['msavi_ponderado']:.4f} ponderado por superficie, con mediana "
            f"{D.mediana([b['msavi'] for b in bs]):.4f} y un rango entre {mmin['msavi']:.4f} "
            f"(bloque {mmin['codigo']}, {mmin['distrito']}) y {mmax['msavi']:.4f} (bloque "
            f"{mmax['codigo']}, {mmax['distrito']}).")
    sobre = [b for b in bs if b["msavi_condicion"] == "Sobre umbral"]
    ha_sobre = sum(b["area_ha"] for b in sobre)
    _p(doc, f"Aplicando el umbral de brecha del proyecto (MSAVI = 0.4976), {len(sobre)} bloques "
            f"({ha_sobre:,.2f} ha, {100 * ha_sobre / r['ha']:.1f} % del ámbito) presentan una "
            f"media por encima del umbral y {r['n'] - len(sobre)} por debajo. Ahora bien, la "
            f"media del bloque oculta la distribución interna: al descender a la estadística "
            f"areal por clase, {r['ha_brecha']:,.2f} ha —el {r['pct_brecha']:.2f} % de la "
            f"superficie provincial— se sitúan bajo el umbral. Esta superficie, y no el conteo "
            f"de bloques, es la base para el cómputo del indicador de brecha de la R.M. N.° "
            f"00213-2024-MINAM.", bold=True)
    filas = [[k, v["n"], round(v["ha"], 2), round(100 * v["ha"] / r["ha"], 2)]
             for k, v in D.conteo(bs, "msavi_clase")]
    filas.append(["TOTAL", r["n"], round(r["ha"], 2), 100.0])
    _tabla(doc, ["Clase de la media MSAVI 2024", "N.° de bloques", "Superficie (ha)",
                 "% del ámbito"], filas,
           titulo="Tabla 6. Distribución de bloques por clase de la media MSAVI 2024",
           fuente="Fuente: elaboración propia sobre el compuesto MSAVI 2024 (Sentinel-2) y la "
                  "estadística zonal del ráster clasificado.",
           anchos=[7.0, 2.4, 2.6, 2.4])

    # Distribucion areal NDVI agregada
    tot_ndvi = defaultdict(float)
    for b in bs:
        for k, (ha_k, _) in b["ndvi_clases"].items():
            if ha_k:
                tot_ndvi[k] += ha_k
    suma = sum(tot_ndvi.values())
    if suma:
        filas = [[k, round(tot_ndvi[k], 2), round(100 * tot_ndvi[k] / suma, 2)]
                 for k in ("Vegetación alta", "Vegetación mediana", "Vegetación ligera",
                           "Tierra desnuda") if k in tot_ndvi]
        filas.append(["TOTAL CLASIFICADO", round(suma, 2), 100.0])
        _tabla(doc, ["Clase NDVI mediana 2025", "Superficie (ha)", "% del área clasificada"],
               filas,
               titulo="Tabla 7. Distribución areal agregada del NDVI mediana 2025",
               fuente="Fuente: elaboración propia sobre la estadística zonal del NDVI mediana "
                      "2025 (Sentinel-2).",
               anchos=[6.4, 3.4, 3.4])
    _p(doc, "LECTURA CRÍTICA. Los índices MSAVI y NDVI miden vigor y densidad de biomasa, no "
            "composición ni integridad ecosistémica. Un valor alto no equivale a ausencia de "
            "degradación: en bloques con mosaico agrícola, pastizal cultivado o plantaciones, "
            "la respuesta espectral puede ser alta sobre una unidad productora sustituida en su "
            "composición. La diferencia entre ambos productos tampoco debe leerse como mejora "
            "entre 2024 y 2025: son índices distintos —el NDVI satura ante biomasa densa, "
            "mientras el MSAVI conserva sensibilidad al suelo de fondo— aplicados sobre "
            "compuestos temporales distintos.", bold=True)
    _p(doc, "ADVERTENCIA DE ESTACIONALIDAD. En los bloques de Bosque Estacionalmente Seco la "
            "caducifolia produce divergencia sistemática entre los compuestos de estiaje y las "
            "tomas de campo de temporada húmeda. Toda meta física dimensionada sobre índices de "
            "estiaje debe tratarse como cota superior, y su cierre exige al menos una captura "
            "de temporada húmeda (enero–abril).")

    doc.add_heading("4.4 Composición ecosistémica de la Unidad Productora", level=2)
    filas = [[k, v["n"], round(v["ha"], 2), round(100 * v["ha"] / r["ha"], 2)]
             for k, v in D.conteo(bs, "ecosistema")]
    filas.append(["TOTAL", r["n"], round(r["ha"], 2), 100.0])
    _tabla(doc, ["Tipo de ecosistema (UP)", "N.° de bloques", "Superficie (ha)", "% del ámbito"],
           filas,
           titulo="Tabla 8. Superficie de la Unidad Productora por tipo de ecosistema",
           fuente="Fuente: elaboración propia sobre la ficha F-DT-03 de cada bloque. La "
                  "superficie corresponde a la del bloque, no a la del polígono de ecosistema, "
                  "que se consigna «Por determinar» en las fichas.",
           anchos=[7.6, 2.2, 2.4, 2.2])
    _p(doc, "El desagregado de superficie por polígono de ecosistema dentro de cada bloque no "
            "consta en las fichas DT y se conserva como «Por determinar». Su cierre requiere el "
            "cruce con el Reporte de Ecosistemas por Bloques del MINAM y es requisito previo al "
            "cómputo definitivo del indicador de brecha, dado que la categoría «Zona agrícola» "
            "corresponde a un agroecosistema y no a un ecosistema natural.")

    doc.add_heading("4.5 Estado de conservación y erosión", level=2)
    for etiqueta, campo, titulo in (
            ("estado de conservación", "estado_conservacion",
             "Tabla 9. Distribución de bloques por estado de conservación"),
            ("nivel general de erosión", "erosion",
             "Tabla 10. Distribución de bloques por nivel general de erosión")):
        filas = [[k, v["n"], round(v["ha"], 2), round(100 * v["ha"] / r["ha"], 2)]
                 for k, v in D.conteo(bs, campo)]
        filas.append(["TOTAL", r["n"], round(r["ha"], 2), 100.0])
        _tabla(doc, [etiqueta.capitalize(), "N.° de bloques", "Superficie (ha)", "% del ámbito"],
               filas, titulo=titulo,
               fuente="Fuente: elaboración propia sobre las fichas F-DT-02 y F-DT-03.",
               anchos=[7.6, 2.2, 2.4, 2.2])


def _s5(doc, prov, r, bs, meta):
    doc.add_heading("5. Integración con el Estudio de Geología — Índice de Susceptibilidad "
                    "Litológica (ISL)", level=1)
    doc.add_heading("5.1 Alcance y método del ISL", level=2)
    _p(doc, f"El {meta['volumen_geologia']} desarrolla, ante la ausencia al momento de su "
            f"elaboración de la capa de pendientes y de los resultados del modelamiento de "
            f"peligro integrado, un Índice de Susceptibilidad Litológica (ISL) que jerarquiza "
            f"los bloques exclusivamente en función de las propiedades del sustrato. A cada "
            f"unidad geológica se le asigna una clase en escala de 1 a 5 según su competencia "
            f"mecánica, su anisotropía estructural, el tipo de perfil de meteorización que "
            f"desarrolla y su estado de consolidación. Se establecen dos escalas independientes: "
            f"ISL-MM para movimientos en masa e ISL-EH para erosión hídrica. El valor de cada "
            f"bloque resulta de ponderar por área las unidades geológicas que lo intersecan.")
    _tabla(doc, ["Rango del ISL", "Clase de susceptibilidad"],
           [["< 1.50", "Muy baja"], ["1.50 – 2.49", "Baja"], ["2.50 – 3.49", "Media"],
            ["3.50 – 4.49", "Alta"], ["≥ 4.50", "Muy alta"]],
           titulo="Tabla 11. Escala de clasificación del ISL",
           fuente=f"Fuente: {meta['volumen_geologia']}, sobre la Carta Geológica Nacional "
                  f"1:50 000 — INGEMMET.",
           anchos=[4.0, 5.0])
    _p(doc, "ALCANCE DEL ÍNDICE. El ISL no constituye una evaluación de peligro ni de riesgo. "
            "No integra pendiente, precipitación detonante, cobertura vegetal ni condiciones "
            "hidrogeológicas. Su integración con esos factores —para obtener el peligro "
            "propiamente dicho— corresponde al Entregable 8 (Análisis de riesgos para el "
            "proyecto). La verificación de campo y los ensayos geotécnicos se ejecutarán en la "
            "fase de estudio definitivo o expediente técnico.", bold=True)

    doc.add_heading("5.2 Resultados del ISL en el ámbito provincial", level=2)
    ha = r["ha"]
    isl_mm_p = sum(b["isl_mm"] * b["area_ha"] for b in bs) / ha
    isl_eh_p = sum(b["isl_eh"] * b["area_ha"] for b in bs) / ha
    _p(doc, f"El ISL-MM ponderado por superficie del ámbito provincial es de {isl_mm_p:.2f} y el "
            f"ISL-EH ponderado, de {isl_eh_p:.2f}. La distribución por clase se presenta a "
            f"continuación.")
    for etiqueta, campo in (("ISL-MM — movimientos en masa", "isl_clase_mm"),
                            ("ISL-EH — erosión hídrica", "isl_clase_eh")):
        agg = defaultdict(lambda: {"n": 0, "ha": 0.0, "cods": []})
        for b in bs:
            agg[b[campo]]["n"] += 1
            agg[b[campo]]["ha"] += b["area_ha"]
            agg[b[campo]]["cods"].append(b["codigo"])
        orden = ["Muy alta", "Alta", "Media", "Baja", "Muy baja"]
        filas = []
        for k in orden:
            if k in agg:
                cods = ", ".join(sorted(agg[k]["cods"])[:12])
                if len(agg[k]["cods"]) > 12:
                    cods += ", …"
                filas.append([k, agg[k]["n"], round(agg[k]["ha"], 2),
                              round(100 * agg[k]["ha"] / ha, 2), cods])
        filas.append(["TOTAL", r["n"], round(ha, 2), 100.0, ""])
        _tabla(doc, ["Clase de susceptibilidad", "N.° de bloques", "Superficie (ha)",
                     "% del ámbito", "Bloques"], filas,
               titulo=f"Tabla {12 if campo == 'isl_clase_mm' else 13}. "
                      f"Distribución de los bloques por clase de {etiqueta}",
               fuente=f"Fuente: elaboración propia sobre el {meta['volumen_geologia']}.",
               anchos=[2.8, 1.8, 2.2, 1.8, 6.4])

    unidades = D.conteo(bs, "geo_unidad")
    filas = [[k, v["n"], round(v["ha"], 2), round(100 * v["ha"] / ha, 2)] for k, v in unidades]
    filas.append(["TOTAL", r["n"], round(ha, 2), 100.0])
    _tabla(doc, ["Unidad geológica dominante del bloque", "N.° de bloques",
                 "Superficie (ha)", "% del ámbito"], filas,
           titulo="Tabla 14. Unidad geológica dominante de los bloques",
           fuente=f"Fuente: elaboración propia sobre el {meta['volumen_geologia']}. La unidad "
                  f"dominante es la de mayor superficie de intersección con el polígono del "
                  f"bloque; no agota su composición litológica.",
           anchos=[7.6, 2.2, 2.4, 2.2])

    doc.add_heading("5.3 Lectura cruzada susceptibilidad — cobertura — pendiente", level=2)
    _p(doc, "La susceptibilidad del sustrato adquiere significado operativo al cruzarse con el "
            "estado de la cobertura y con la pendiente. Un sustrato de alta susceptibilidad "
            "bajo dosel funcional y pendiente moderada plantea una prioridad de conservación; "
            "el mismo sustrato con cobertura bajo el umbral espectral y pendiente igual o "
            "superior al 25 % plantea una prioridad de control de erosión y estabilización.")
    conv = [b for b in bs if (b["isl_clase_mm"] in ("Alta", "Muy alta")
                              or b["isl_clase_eh"] in ("Alta", "Muy alta"))
            and (b["pct_bajo_umbral"] or 0) >= 50 and (b["pend_pct"] or 0) >= 25]
    ha_conv = sum(b["area_ha"] for b in conv)
    prot = [b for b in bs if (b["isl_clase_mm"] in ("Alta", "Muy alta")
                              or b["isl_clase_eh"] in ("Alta", "Muy alta"))
            and (b["pct_bajo_umbral"] or 0) < 50 and (b["pend_pct"] or 0) >= 25]
    ha_prot = sum(b["area_ha"] for b in prot)
    _p(doc, f"En el ámbito provincial, {len(conv)} bloques ({ha_conv:,.2f} ha, "
            f"{100 * ha_conv / ha:.1f} %) presentan convergencia de los tres factores: sustrato "
            f"de susceptibilidad alta o muy alta, más del 50 % de su superficie bajo el umbral "
            f"MSAVI 0.4976 y pendiente promedio igual o superior al 25 %. Constituyen la "
            f"prioridad máxima del ámbito para el control de erosión y la estabilización de "
            f"laderas."
            if conv else
            f"En el ámbito provincial ningún bloque presenta convergencia simultánea de "
            f"sustrato de susceptibilidad alta o muy alta, cobertura mayoritariamente bajo el "
            f"umbral MSAVI 0.4976 y pendiente promedio igual o superior al 25 %.")
    _p(doc, f"En contraste, {len(prot)} bloques ({ha_prot:,.2f} ha, {100 * ha_prot / ha:.1f} %) "
            f"combinan sustrato susceptible y pendiente igual o superior al 25 % con una "
            f"cobertura mayoritariamente sobre el umbral: en ellos la cobertura vegetal está "
            f"cumpliendo hoy la función de regulación que el proyecto busca recuperar, y la "
            f"medida pertinente es la conservación del dosel funcional antes que la "
            f"restauración activa."
            if prot else
            "No se registran en el ámbito bloques que combinen sustrato susceptible y pendiente "
            "igual o superior al 25 % con cobertura mayoritariamente sobre el umbral.")
    _p(doc, "ADVERTENCIA METODOLÓGICA. Esta lectura cruzada es una aproximación de gabinete "
            "destinada a ordenar la prioridad preliminar. No sustituye el modelamiento del "
            "peligro integrado por AHP (PMM + EPH + PGI) de la mesolocalización, que sigue "
            "siendo el insumo requerido para fijar la prioridad definitiva de intervención de "
            "cada bloque.", bold=True)


def _s6(doc, prov, r, bs, meta):
    doc.add_heading("6. Procesos erosivos y cárcavas", level=1)
    con = [b for b in bs if b["n_carcavas"]]
    if con:
        _p(doc, f"El inventario de cárcavas codificadas registra {r['n_carcavas']} rasgos en "
                f"{len(con)} de los {r['n']} bloques del ámbito, con una longitud acumulada de "
                f"{r['long_carcavas']:,.2f} m. La caracterización morfométrica de cada rasgo "
                f"—longitud, cotas, pendiente promedio, índice de irregularidad y NDVI sobre el "
                f"eje— permite discriminar los rasgos activos de los estabilizados por "
                f"vegetación.")
        filas = []
        for b in sorted(con, key=lambda x: -x["long_carcavas"]):
            filas.append([b["codigo"], b["distrito"], round(b["area_ha"], 2), b["n_carcavas"],
                          round(b["long_carcavas"], 2),
                          round(b["long_carcavas"] / b["area_ha"], 2),
                          b["erosion"], b["urgencia_erosion"]])
        filas.append(["TOTAL", "", round(sum(b["area_ha"] for b in con), 2),
                      r["n_carcavas"], round(r["long_carcavas"], 2), "", "", ""])
        _tabla(doc, ["Bloque", "Distrito", "Superficie (ha)", "N.° de cárcavas",
                     "Longitud total (m)", "Densidad (m/ha)", "Nivel de erosión (F-DT-02)",
                     "Urgencia de control"], filas,
               titulo="Tabla 15. Bloques con cárcavas codificadas inventariadas",
               fuente="Fuente: elaboración propia sobre la capa «Cárcavas codificadas» y su "
                      "caracterización morfométrica, cruzada con la ficha F-DT-02.",
               anchos=[1.6, 3.0, 2.2, 1.8, 2.2, 2.0, 2.6, 1.8])
        todas = [c for b in con for c in b["carcavas"]]
        largas = sorted(todas, key=lambda c: -c["longitud"])[:5]
        _p(doc, f"La longitud individual de los rasgos inventariados varía entre "
                f"{min(c['longitud'] for c in todas):,.2f} m y "
                f"{max(c['longitud'] for c in todas):,.2f} m. Los cinco de mayor desarrollo son "
                + ", ".join(f"{c['codigo']} ({c['longitud']:,.0f} m)" for c in largas) + ".")
        desnudo = [c for c in todas if c["clase_ndvi"] == "Tierra desnuda"]
        if desnudo:
            _p(doc, f"{len(desnudo)} de los {len(todas)} rasgos presentan clase NDVI «Tierra "
                    f"desnuda» sobre su eje, indicativa de actividad erosiva sin colonización "
                    f"vegetal: son los candidatos prioritarios a tratamiento con estructuras "
                    f"transversales de retención y revegetación de cabecera.")
    else:
        _p(doc, f"No se registran cárcavas digitalizadas y codificadas en los bloques de la "
                f"provincia de {meta['nombre']} dentro de la capa vectorial disponible.")
    _p(doc, "LIMITACIÓN SUSTANTIVA DEL INVENTARIO. La ausencia de registros en un bloque NO "
            "equivale a ausencia verificada de cárcavas. El levantamiento instrumental "
            "georreferenciado está pendiente en la mayor parte del ámbito y es requisito previo "
            "a la fijación de metas físicas de infraestructura gris (diques de contención, "
            "estructuras transversales, obras de estabilización). Las fichas DT consignan el "
            "campo «N.° de cárcavas registradas» como «Por determinar» en la generalidad de los "
            "bloques, y así se conserva.", bold=True)


def _s7(doc, prov, r, bs, meta):
    doc.add_heading("7. Caracterización diagnóstica — fichas F-DT consolidadas", level=1)

    doc.add_heading("7.1 F-DT-01 · Datos generales y fisiografía", level=2)
    for etiqueta, campo in (("forma predominante del terreno", "forma_terreno"),
                            ("posición fisiográfica", "posicion")):
        filas = [[k, v["n"], round(v["ha"], 2), round(100 * v["ha"] / r["ha"], 2)]
                 for k, v in D.conteo(bs, campo)]
        filas.append(["TOTAL", r["n"], round(r["ha"], 2), 100.0])
        _tabla(doc, [etiqueta.capitalize(), "N.° de bloques", "Superficie (ha)", "% del ámbito"],
               filas, fuente=None, anchos=[7.6, 2.2, 2.4, 2.2])
    aflo = len([b for b in bs if D.norm(b["afloramientos"]) == "SI"])
    esc = len([b for b in bs if D.norm(b["escarpes"]) == "SI"])
    rem = len([b for b in bs if D.norm(b["remociones"]) == "SI"])
    _p(doc, f"Las brigadas registran afloramientos rocosos en {aflo} bloques, escarpes activos "
            f"en {esc} y remociones en masa activas en {rem}, sobre un universo de {r['n']}. "
            f"El registro de escarpes y remociones activas es la evidencia de campo directa de "
            f"procesos de ladera en curso y debe cruzarse con el ISL-MM del sustrato "
            f"(sección 5) antes de descartar cualquier bloque del tratamiento de estabilización.")

    doc.add_heading("7.2 F-DT-02 · Suelo y procesos erosivos", level=2)
    filas = [[k, v["n"], round(v["ha"], 2), round(100 * v["ha"] / r["ha"], 2)]
             for k, v in D.conteo(bs, "urgencia_erosion")]
    filas.append(["TOTAL", r["n"], round(r["ha"], 2), 100.0])
    _tabla(doc, ["Urgencia de control de erosión", "N.° de bloques", "Superficie (ha)",
                 "% del ámbito"], filas,
           titulo="Tabla 16. Urgencia de control de erosión declarada en campo",
           fuente="Fuente: elaboración propia sobre la ficha F-DT-02.",
           anchos=[7.6, 2.2, 2.4, 2.2])
    sin_carc = len([b for b in bs if D.es_sin_dato(b["carcavas_ficha"])])
    _p(doc, f"El campo «N.° de cárcavas registradas» de la ficha F-DT-02 se consigna «Por "
            f"determinar» en {sin_carc} de los {r['n']} bloques. El inventario instrumental "
            f"pendiente se trata en la sección 6.")

    doc.add_heading("7.3 F-DT-03 · Ecosistema, cobertura y valor ecológico", level=2)
    tax = [b["taxones"] for b in bs if b["taxones"] is not None]
    bajo15 = len([t for t in tax if t < 15])
    _p(doc, f"El elenco florístico registrado promedia {sum(tax) / len(tax):.1f} taxones por "
            f"bloque (mínimo {min(tax):.0f}, máximo {max(tax):.0f}). En {bajo15} de los "
            f"{len(tax)} bloques con dato el registro no alcanza el mínimo de 15 taxones que "
            f"exige el formato V5. Se trata de registros florísticos preliminares, sin colecta "
            f"ni determinación botánica formal: tienen valor de hipótesis de trabajo y no de "
            f"inventario. Su cierre exige colecta botánica en la siguiente campaña, y es "
            f"requisito para definir el elenco de especies de revegetación y enriquecimiento.")
    filas = [[k, v["n"], round(v["ha"], 2), round(100 * v["ha"] / r["ha"], 2)]
             for k, v in D.conteo(bs, "uso_actual")]
    filas.append(["TOTAL", r["n"], round(r["ha"], 2), 100.0])
    _tabla(doc, ["Uso actual dominante del suelo", "N.° de bloques", "Superficie (ha)",
                 "% del ámbito"], filas,
           titulo="Tabla 17. Uso actual dominante del suelo",
           fuente="Fuente: elaboración propia sobre la ficha F-DT-03.",
           anchos=[7.6, 2.2, 2.4, 2.2])
    cob = len([b for b in bs if D.es_sin_dato(b["cobertura_pct"])])
    _p(doc, f"Los campos «Cobertura vegetal total» y «Suelo desnudo» medidos en campo se "
            f"consignan «Por determinar» en {cob} de los {r['n']} bloques. En su ausencia, la "
            f"única estimación disponible de cobertura procede de los índices espectrales, con "
            f"las limitaciones declaradas en la sección 4.3.")

    doc.add_heading("7.4 F-DT-04 · Causas e indicadores de degradación", level=2)
    for etiqueta, campo, titulo in (
            ("velocidad de degradación", "velocidad",
             "Tabla 18. Velocidad de degradación declarada"),
            ("reversibilidad técnica", "reversibilidad",
             "Tabla 19. Reversibilidad técnica declarada")):
        filas = [[k, v["n"], round(v["ha"], 2), round(100 * v["ha"] / r["ha"], 2)]
                 for k, v in D.conteo(bs, campo)]
        filas.append(["TOTAL", r["n"], round(r["ha"], 2), 100.0])
        _tabla(doc, [etiqueta.capitalize(), "N.° de bloques", "Superficie (ha)", "% del ámbito"],
               filas, titulo=titulo,
               fuente="Fuente: elaboración propia sobre la ficha F-DT-04.",
               anchos=[7.6, 2.2, 2.4, 2.2])
    causas = D.conteo(bs, "causa")[:6]
    _p(doc, "Las causas subyacentes de degradación consignadas con mayor frecuencia en el "
            "ámbito, ordenadas por superficie afectada, son:")
    for k, v in causas:
        _vineta(doc, f"{k} — {v['n']} bloque(s), {v['ha']:,.2f} ha "
                     f"({100 * v['ha'] / r['ha']:.1f} %).", size=9)
    _p(doc, "La redacción de las causas no ha sido homologada entre brigadas: se reproduce tal "
            "como consta en cada ficha. La homologación previa a la construcción del árbol de "
            "problemas del Perfil es una tarea pendiente que se registra en la sección 9.")

    doc.add_heading("7.5 F-DT-05 · Recursos hídricos y accesibilidad", level=2)
    rec = len([b for b in bs if D.norm(b["recarga"]) == "SI"])
    ha_rec = sum(b["area_ha"] for b in bs if D.norm(b["recarga"]) == "SI")
    _p(doc, f"{rec} de los {r['n']} bloques ({ha_rec:,.2f} ha, {100 * ha_rec / r['ha']:.1f} % "
            f"del ámbito) están declarados como zona de recarga hídrica. Esta condición es "
            f"determinante para la priorización: la función de regulación hídrica —"
            f"interceptación, infiltración y recarga de acuíferos— es uno de los servicios "
            f"ecosistémicos que el proyecto busca recuperar, y las medidas de infiltración "
            f"(zanjas, acequias de conducción, terrazas de formación lenta) rinden en estos "
            f"bloques su mayor retorno por hectárea intervenida.")
    filas = [[k, v["n"], round(v["ha"], 2), round(100 * v["ha"] / r["ha"], 2)]
             for k, v in D.conteo(bs, "acceso")]
    filas.append(["TOTAL", r["n"], round(r["ha"], 2), 100.0])
    _tabla(doc, ["Modalidad de acceso", "N.° de bloques", "Superficie (ha)", "% del ámbito"],
           filas, titulo="Tabla 20. Modalidad de acceso a los bloques",
           fuente="Fuente: elaboración propia sobre la ficha F-DT-05. La accesibilidad "
                  "determina el costo unitario de transporte de plantones e insumos y la "
                  "logística de vigilancia durante el período de prendimiento.",
           anchos=[7.6, 2.2, 2.4, 2.2])


def _s8(doc, prov, r, bs, meta):
    doc.add_heading("8. Análisis preliminar del riesgo en contexto de cambio climático "
                    "(AdR-CCC)", level=1)
    _p(doc, "Conforme al Anexo 2 de la Guía General DGPMI-MEF (2022), el análisis distingue "
            "peligro, exposición, fragilidad y resiliencia. El proyecto interviene en la ZONA "
            "DE ORIGEN —la cuenca alta—, no en la zona de impacto del valle y la planicie "
            "costera: los peligros que se abordan son los que se generan en la ladera.")

    doc.add_heading("8.1 Peligros identificados", level=2)
    _vineta(doc, "Peligro por Movimientos en Masa (PMM): deslizamientos de tierra y roca, flujos "
                 "de detritos y lodo (huaicos), caída de rocas y volcamientos, y reptación de "
                 "suelos en laderas degradadas. Factores condicionantes: geología, pendiente, "
                 "índice de posición topográfica y cobertura vegetal. Factor desencadenante "
                 "principal: precipitación acumulada multianual.")
    _vineta(doc, "Erosión Potencial Hídrica (EPH): pérdida de suelo por acción del agua, "
                 "carcavamiento, erosión en surcos y socavamiento de márgenes de quebradas.")
    _vineta(doc, "Peligro Generador de Inundaciones (PGI): zonas generadoras de escorrentía que "
                 "contribuyen a inundaciones aguas abajo. Variables: pendiente, escorrentía "
                 "total y NDVI.")
    _vineta(doc, "Peligro Integrado (PI): resultado de la integración ponderada por AHP de "
                 "PMM + EPH + PGI, en cinco niveles.")
    _p(doc, "ADVERTENCIA SUSTANTIVA. El nivel de PELIGRO INTEGRADO (MCA-AHP) por bloque NO está "
            "disponible en los insumos de este entregable, y así se consigna en la totalidad de "
            "las fichas del ámbito. Debe tomarse del modelamiento de mesolocalización antes de "
            "fijar la prioridad definitiva de intervención. Los descriptores que este informe "
            "aporta —ISL del sustrato, pendiente zonal, brecha espectral de cobertura, evidencia "
            "de campo de escarpes y remociones— son componentes de ese peligro, no el peligro "
            "integrado.", bold=True)

    doc.add_heading("8.2 Exposición", level=2)
    _p(doc, f"La UP expuesta en el ámbito provincial asciende a {r['ha']:,.2f} ha distribuidas "
            f"en {r['n']} bloques, {r['distritos']} distritos y {r['microcuencas']} "
            f"microcuencas. Los elementos expuestos comprenden el propio ecosistema —suelo, "
            f"cobertura vegetal y red de drenaje natural—, los centros poblados asociados "
            f"consignados en las fichas y las vías de acceso, así como las MRR-CCC que el "
            f"proyecto instale, que quedarán expuestas a los mismos peligros durante su período "
            f"de establecimiento.")

    doc.add_heading("8.3 Fragilidad", level=2)
    _p(doc, f"La fragilidad de la UP se expresa en tres registros convergentes. Primero, la "
            f"brecha espectral: {r['ha_brecha']:,.2f} ha ({r['pct_brecha']:.2f} % del ámbito) "
            f"se sitúan bajo el umbral MSAVI 0.4976, lo que indica cobertura insuficiente para "
            f"la función de protección de suelo. Segundo, la susceptibilidad del sustrato: el "
            f"ISL-EH ponderado del ámbito es de "
            f"{D.ponderada(bs, 'isl_eh'):.2f} y el ISL-MM ponderado, "
            f"de {D.ponderada(bs, 'isl_mm'):.2f}. Tercero, la "
            f"pendiente: la media ponderada del ámbito es de "
            f"{r['pend_ponderada']:.2f} % según el rótulo del catálogo, equivalente a "
            f"{D.ponderada(bs, 'pend_si_grados_pct'):.2f} % si el valor estuviera expresado en "
            f"grados (discrepancia D-P01, sección 10.2).")
    est = D.conteo(bs, "estado_conservacion")
    alterados = sum(v["ha"] for k, v in est if "lterad" in k or "egradad" in k)
    if alterados:
        _p(doc, f"A ello se suma el estado de conservación declarado en campo: {alterados:,.2f} "
                f"ha ({100 * alterados / r['ha']:.1f} % del ámbito) figuran con algún grado de "
                f"alteración o degradación en la ficha F-DT-03.")

    doc.add_heading("8.4 Resiliencia y factores favorables", level=2)
    reg = [b for b in bs if D.norm(b["regeneracion"]) in ("SI", "PRESENTE", "ABUNDANTE", "MODERADA")]
    sano = len([b for b in bs if D.norm(b["estado_sanitario"]) == "SANO"])
    revers = sum(v["ha"] for k, v in D.conteo(bs, "reversibilidad")
                 if "reversible" in k.lower() and "no" not in k.lower()[:3])
    _p(doc, f"Del lado de la resiliencia, {revers:,.2f} ha ({100 * revers / r['ha']:.1f} % del "
            f"ámbito) están calificadas en campo como técnicamente reversibles, total o "
            f"parcialmente, lo que sustenta la viabilidad de la restauración pasiva y del "
            f"enriquecimiento bajo dosel frente a la restauración activa integral, de costo "
            f"unitario muy superior. El estado sanitario de la vegetación se declara sano en "
            f"{sano} de los {r['n']} bloques.")
    if reg:
        _p(doc, f"La presencia de regeneración natural, registrada en {len(reg)} bloques, es el "
                f"indicador operativo más relevante para decidir entre clausura temporal —de "
                f"costo bajo y alto rendimiento donde el banco de semillas persiste— y "
                f"plantación, y debe completarse en la verificación pendiente.")
    else:
        _p(doc, "El campo «Regeneración natural» se consigna «Por determinar» en la generalidad "
                "de los bloques del ámbito. Su cierre es prioritario: es el indicador operativo "
                "que permite decidir entre clausura temporal —de costo bajo y alto rendimiento "
                "donde el banco de semillas persiste— y plantación, y esa decisión gobierna una "
                "fracción sustantiva del presupuesto de inversión.")

    doc.add_heading("8.5 Síntesis del riesgo y prioridad preliminar", level=2)
    filas = [[k, v["n"], round(v["ha"], 2), round(100 * v["ha"] / r["ha"], 2)]
             for k, v in D.conteo(bs, "urgencia")]
    filas.append(["TOTAL", r["n"], round(r["ha"], 2), 100.0])
    _tabla(doc, ["Urgencia de intervención declarada", "N.° de bloques", "Superficie (ha)",
                 "% del ámbito"], filas,
           titulo="Tabla 21. Urgencia de intervención declarada en campo",
           fuente="Fuente: elaboración propia sobre la síntesis GdR-CCC de las fichas DT. La "
                  "urgencia declarada es una apreciación de campo y no sustituye la prioridad "
                  "derivada del modelamiento de peligro integrado.",
           anchos=[7.6, 2.2, 2.4, 2.2])


def _s9(doc, prov, r, bs, meta):
    doc.add_heading("9. Medidas de reducción del riesgo y alternativas de infraestructura "
                    "natural (MRR-CCC)", level=1)
    _p(doc, "Las medidas que siguen se enuncian como alternativas técnicamente pertinentes al "
            "diagnóstico levantado. Su dimensionamiento físico y su costeo NO se realizan en "
            "este entregable: dependen del modelamiento de peligro integrado, del cierre de los "
            "vacíos de información de la sección 10 y del análisis técnico del Entregable 10.")

    doc.add_heading("9.1 Infraestructura natural (verde)", level=2)
    _vineta(doc, "Revegetación con especies nativas mediante enriquecimiento bajo dosel y "
                 "siembra directa, en los bloques con dosel funcional y regeneración natural "
                 "presente.")
    _vineta(doc, "Reforestación con especies forestales nativas en los sectores bajo el umbral "
                 "espectral sin regeneración, previa verificación del banco de semillas.")
    _vineta(doc, "Zanjas de infiltración y acequias de conducción, prioritariamente en los "
                 "bloques declarados zona de recarga hídrica.")
    _vineta(doc, "Terrazas de formación lenta en las laderas de pendiente moderada con uso "
                 "agrícola o mixto.")
    _vineta(doc, "Clausura temporal y manejo de pastizales, donde la causa dominante de "
                 "degradación sea la presión de pastoreo.")
    _vineta(doc, "Barreras vivas y cercos vivos como medida de contención perimetral y de "
                 "delimitación predial.")

    doc.add_heading("9.2 Infraestructura gris complementaria", level=2)
    _vineta(doc, "Obras de estabilización de taludes en los sectores con escarpes activos y "
                 "sustrato de ISL-MM alto o muy alto.")
    _vineta(doc, "Muros de contención y gaviones en los puntos críticos de ladera.")
    _vineta(doc, "Diques de contención y estructuras transversales de retención en las cárcavas "
                 "inventariadas, comenzando por los rasgos de mayor desarrollo y NDVI de suelo "
                 "desnudo sobre el eje.")
    _vineta(doc, "Obras de drenaje y control de escorrentía en las vías de acceso interiores, "
                 "cuya habilitación sin instrumento de manejo figura entre las causas de "
                 "degradación consignadas en campo.")
    _p(doc, "El dimensionamiento de la infraestructura gris está condicionado por el "
            "levantamiento instrumental de cárcavas pendiente (sección 6). No es posible fijar "
            "metas físicas de esta línea sobre el inventario parcial disponible.", bold=True)

    doc.add_heading("9.3 Gobernanza y gestión comunitaria", level=2)
    _vineta(doc, "Planes de manejo comunitario de ecosistemas, articulados con los titulares de "
                 "predios interceptados por cada bloque.")
    _vineta(doc, "Fortalecimiento de capacidades locales en gestión del riesgo de desastres.")
    _vineta(doc, "Sistemas de alerta temprana comunitarios en las microcuencas con generación de "
                 "escorrentía hacia centros poblados.")
    _vineta(doc, "Mecanismos de retribución por servicios ecosistémicos (MERESE) como "
                 "instrumento de sostenibilidad de la fase de funcionamiento.")

    doc.add_heading("9.4 Clasificación por tipo de gestión del riesgo", level=2)
    _tabla(doc, ["Tipo de gestión", "Medidas asociadas", "Momento de aplicación"],
           [["Prospectiva",
             "Clausura temporal, planes de manejo comunitario, delimitación con cercos vivos, "
             "ordenamiento del uso de vías interiores, MERESE",
             "Evita la generación de nuevo riesgo sobre los sectores aún funcionales de la UP"],
            ["Correctiva",
             "Revegetación y reforestación, zanjas de infiltración, terrazas de formación lenta, "
             "estabilización de taludes, tratamiento de cárcavas",
             "Reduce el riesgo existente sobre la superficie ya degradada"],
            ["Reactiva",
             "Sistemas de alerta temprana comunitarios, protocolos de respuesta ante "
             "movimientos en masa activos",
             "Prepara la respuesta ante la ocurrencia del peligro"]],
           titulo="Tabla 22. Clasificación de las MRR-CCC por tipo de gestión del riesgo",
           fuente="Fuente: elaboración propia conforme a la Ley 29664 (SINAGERD) y su "
                  "reglamento, D.S. 048-2011-PCM.",
           anchos=[2.6, 6.6, 5.4])

    doc.add_heading("9.5 Tratamiento de las MRR-CCC en la evaluación social", level=2)
    _p(doc, "Conforme a la Guía General DGPMI, los costos de inversión y de operación y "
            "mantenimiento de las MRR-CCC deben registrarse por separado e integrarse en los "
            "flujos de la evaluación social. Las medidas integradas —aquellas que forman parte "
            "de la solución técnica del proyecto— no se evalúan por separado. Las medidas "
            "externas —las que protegen a múltiples unidades productoras o a población fuera "
            "del ámbito del proyecto— sí requieren evaluación de rentabilidad social "
            "independiente. Dado el carácter intergeneracional de los beneficios de la "
            "recuperación ecosistémica, corresponde aplicar la Tasa Social de Descuento de "
            "Largo Plazo y realizar análisis de sensibilidad sobre las variables climáticas y "
            "ecológicas críticas.")


def _s10(doc, prov, r, bs, meta):
    doc.add_heading("10. Vacíos de información, consistencia y trazabilidad", level=1)

    doc.add_heading("10.1 Completitud de la información", level=2)
    campos = [("Peligro integrado (MCA-AHP)", "peligro_integrado"),
              ("Prioridad de intervención", "prioridad"),
              ("Superficie de ecosistema por polígono", "ecosistema_sup"),
              ("Cobertura vegetal total medida en campo", "cobertura_pct"),
              ("Suelo desnudo medido en campo", "suelo_desnudo"),
              ("Regeneración natural", "regeneracion"),
              ("N.° de cárcavas registradas en ficha", "carcavas_ficha"),
              ("Comunidad campesina", "comunidad")]
    filas = []
    for etiqueta, campo in campos:
        if campo == "ecosistema_sup":
            faltan = len(bs)  # el campo se consigna "Por determinar" en la plantilla
            faltan = len([b for b in bs if D.es_sin_dato(b.get("ecosistema_sup", "Por determinar"))])
        else:
            faltan = len([b for b in bs if D.es_sin_dato(b[campo])])
        filas.append([etiqueta, r["n"] - faltan, faltan,
                      round(100 * (r["n"] - faltan) / r["n"], 1)])
    _tabla(doc, ["Campo del diagnóstico", "Bloques con dato", "Bloques sin dato",
                 "Completitud (%)"], filas,
           titulo="Tabla 23. Completitud de los campos críticos del diagnóstico",
           fuente="Fuente: elaboración propia. «Sin dato» comprende las marcas de origen «Por "
                  "determinar» y «Por verificar». Ningún valor ha sido imputado.",
           anchos=[7.6, 2.2, 2.2, 2.4])

    doc.add_heading("10.2 Registro de discrepancias", level=2)
    cnt = defaultdict(int)
    tipos = defaultdict(int)
    for b in bs:
        for d in b["consistencia"]:
            cnt[d["calificacion"]] += 1
            if d["calificacion"] == "SUSTANTIVA":
                tipos[d["campo"]] += 1
    total = sum(cnt.values())
    filas = [[k, v, round(100 * v / total, 1)]
             for k, v in sorted(cnt.items(), key=lambda kv: -kv[1])]
    filas.append(["TOTAL DE VERIFICACIONES", total, 100.0])
    _tabla(doc, ["Calificación", "N.° de verificaciones", "% del total"], filas,
           titulo="Tabla 24. Resultado del control de consistencia del ámbito",
           fuente="Fuente: elaboración propia sobre la hoja «Control de consistencia» de cada "
                  "plantilla de bloque. Una calificación CONFORME acredita que la verificación "
                  "se ejecutó y no arrojó discrepancia; no equivale a validación de campo del "
                  "dato.",
           anchos=[5.6, 3.4, 3.0])
    if tipos:
        _p(doc, "Los campos que concentran las discrepancias calificadas como SUSTANTIVAS, "
                "ordenados por frecuencia, son:")
        for k, v in sorted(tipos.items(), key=lambda kv: -kv[1])[:8]:
            _vineta(doc, f"{k} — {v} bloque(s).", size=9)
    _p(doc, "El detalle bloque a bloque de las discrepancias, con su calificación y el "
            "tratamiento adoptado, figura en la hoja «9. Consistencia» del anexo de matrices "
            "que acompaña a este informe.")

    doc.add_heading("10.3 Discrepancias de alcance provincial detectadas en esta consolidación",
                    level=2)
    _p(doc, "Además de las discrepancias registradas bloque a bloque, la consolidación del "
            "ámbito ha puesto de manifiesto dos discrepancias que afectan a la totalidad de la "
            "cartera y que no pueden resolverse en este informe. Se declaran conforme al "
            "principio de trazabilidad: se documenta el conflicto, se indica el valor adoptado y "
            "se eleva la resolución a la Subdirección.")
    filas = [
        ["D-P01", "Unidad de la pendiente promedio del catálogo",
         "Las 117 plantillas de resumen V6 rotulan el valor de pendiente del catálogo como "
         "PORCENTAJE y derivan los grados como arcotangente de ese valor dividido entre 100. El "
         "Informe Técnico Consolidado de Diagnóstico Territorial del distrito de Frías (Ayabaca) "
         "concluye que el MISMO valor está expresado en GRADOS, tras contrastarlo con la "
         "distribución de clases de la capa «Pendientes vector»: error cuadrático medio de "
         "1,14° leyendo el catálogo en grados frente a 12,70° leyéndolo como porcentaje "
         "(r = +0,956). Se ha verificado que el valor numérico es idéntico en ambas fuentes para "
         "los diez bloques de Frías, de modo que el conflicto es de unidad y no de dato.",
         "MUY ALTO. Bajo la lectura en grados la pendiente del ámbito casi se duplica en "
         "términos de porcentaje, lo que reclasifica los bloques y cambia el dimensionamiento de "
         "toda partida de movimiento de tierras (zanjas de infiltración, terrazas de formación "
         "lenta). NO afecta la elegibilidad: bajo ambas lecturas ningún bloque supera el "
         "criterio de idoneidad del 75 %.",
         "Se conserva el valor del catálogo con el rótulo de las plantillas (porcentaje) y se "
         "publica en paralelo el equivalente bajo la lectura en grados, en el anexo de matrices. "
         "No se modifica ninguna fuente primaria.",
         "ESCALAR A SESDI con carácter prioritario. Debe emitirse pronunciamiento formal sobre la "
         "unidad del campo de pendiente del catálogo y homologarse en las 117 plantillas y en "
         "todos los informes DT antes de fijar metas físicas."],
        ["D-M01", "Bloques sin cobertura del modelo digital de elevación",
         "Los bloques 83, 84, 85, 86 y 87 (San Juan de Bigote, Morropón; 133,57 ha en conjunto) "
         "presentan altitud mínima = altitud máxima = 0 msnm, amplitud = 0 m y pendiente = 0 % en "
         "el reporte de estadística zonal. Un bloque no está a nivel del mar ni es perfectamente "
         "plano: el cero es el marcador de ausencia de cobertura del MDE, no una medición. Son "
         "los mismos cinco bloques que no figuran en la base general del estudio de geología y "
         "que se incorporaron mediante un archivo complementario.",
         "ALTO en el ámbito de Morropón. Leídos como medición, esos ceros arrastran a la baja la "
         "pendiente media provincial y sitúan falsamente el mínimo altitudinal del ámbito en "
         "0 msnm.",
         "Los cinco valores se anulan y se excluyen de todo promedio, mínimo y máximo del "
         "informe; los bloques se conservan en el universo de superficie. Su clase de pendiente "
         "se consigna «Sin cobertura del MDE».",
         "Ejecutar la estadística zonal de los cinco polígonos sobre el MDE vigente y reponer "
         "altitud, amplitud y pendiente antes del cierre del entregable."],
    ]
    sub = [f for f in filas if f[0] == "D-P01" or r["sin_mde"]]
    _tabla(doc, ["Cód.", "Materia", "Descripción de la discrepancia",
                 "Efecto sobre la formulación", "Valor adoptado en este informe",
                 "Acción requerida"],
           sub, fuente="Fuente: elaboración propia de esta consolidación.",
           anchos=[1.2, 2.8, 6.4, 4.0, 3.6, 3.6])

    doc.add_heading("10.4 Acciones requeridas antes del cierre del entregable", level=2)
    acciones = [
        ("Incorporar el peligro integrado (MCA-AHP) por bloque desde el modelamiento de "
         "mesolocalización", "Sin él no puede fijarse la prioridad definitiva de intervención "
         "ni jerarquizarse la cartera del ámbito.", "Especialista SIG / Mesolocalización"),
        ("Ejecutar el levantamiento instrumental georreferenciado de cárcavas en los bloques "
         "sin inventario", "Requisito previo a la fijación de metas físicas de infraestructura "
         "gris.", "Especialista de Infraestructura Marrón"),
        ("Cerrar el desagregado de superficie por polígono de ecosistema dentro de cada bloque",
         "Determina qué fracción de la superficie ingresa al cómputo del indicador de brecha de "
         "la R.M. N.° 00213-2024-MINAM; la categoría «Zona agrícola» es agroecosistema y no "
         "ecosistema natural.", "Especialista Verde / SIG"),
        ("Completar en campo la cobertura vegetal, el suelo desnudo y la regeneración natural",
         "Gobierna la decisión entre clausura temporal y plantación, y con ella una fracción "
         "sustantiva del presupuesto de inversión.", "Brigadas SESDI"),
        ("Ejecutar colecta y determinación botánica formal", "El elenco florístico actual es "
         "preliminar y no alcanza el mínimo del formato V5 en la mayoría de los bloques; "
         "condiciona la selección de especies de revegetación.", "Especialista botánico"),
        ("Confirmar la equivalencia entre el valor de clase DN del ráster MSAVI y los umbrales "
         "del proyecto", "La equivalencia fue establecida por contraste estadístico "
         "(r = 0.9965) y requiere ratificación del especialista.", "Especialista SIG"),
        ("Incorporar al menos una captura de temporada húmeda (enero–abril) de los índices de "
         "vegetación", "Las metas dimensionadas sobre compuestos de estiaje son cota superior "
         "en los bloques de bosque estacionalmente seco.", "Especialista SIG"),
        ("Homologar la redacción de las causas de degradación entre brigadas", "Requisito para "
         "la construcción del árbol de problemas del Perfil.", "SESDI"),
        ("Resolver la unidad del campo de pendiente del catálogo (discrepancia D-P01)",
         "Determina el dimensionamiento de toda partida de movimiento de tierras del proyecto.",
         "SESDI / Especialista SIG"),
        ("Reponer la estadística zonal de los bloques sin cobertura del MDE (discrepancia D-M01)",
         "Sin altitud ni pendiente no puede zonificarse internamente ni seleccionarse el elenco "
         "de especies de esos bloques.", "Especialista SIG"),
    ]
    _tabla(doc, ["Acción requerida", "Efecto sobre la formulación", "Responsable"],
           [[a, b, c] for a, b, c in acciones],
           titulo="Tabla 25. Acciones requeridas para el cierre del diagnóstico",
           fuente="Fuente: elaboración propia.",
           anchos=[5.4, 6.4, 2.8])


def _s11(doc, prov, r, bs, meta):
    doc.add_heading("11. Conclusiones", level=1)
    ha = r["ha"]
    isl_alta = [b for b in bs if b["isl_clase_mm"] in ("Alta", "Muy alta")
                or b["isl_clase_eh"] in ("Alta", "Muy alta")]
    ha_isl = sum(b["area_ha"] for b in isl_alta)
    conc = [
        f"El ámbito de la provincia de {meta['nombre']} comprende {r['n']} bloques preliminares "
        f"de intervención con una superficie total de {ha:,.2f} ha, distribuidos en "
        f"{r['distritos']} distrito{'s' if r['distritos'] > 1 else ''} y {r['microcuencas']} "
        f"microcuencas, sobre un gradiente altitudinal de {r['alt_min']:,.0f} a "
        f"{r['alt_max']:,.0f} msnm.",

        f"La brecha espectral del ámbito asciende a {r['ha_brecha']:,.2f} ha, el "
        f"{r['pct_brecha']:.2f} % de su superficie, medida como área bajo el umbral MSAVI "
        f"0.4976. Esta cifra —y no el conteo de bloques con media bajo umbral— es la base "
        f"pertinente para el cómputo del indicador de brecha de la R.M. N.° 00213-2024-MINAM, y "
        f"debe depurarse descontando la fracción de agroecosistema que el desagregado de "
        f"ecosistemas aún pendiente determine.",

        f"La integración con el {meta['volumen_geologia']} sitúa {len(isl_alta)} bloques "
        f"({ha_isl:,.2f} ha, {100 * ha_isl / ha:.1f} % del ámbito) en clase de susceptibilidad "
        f"litológica alta o muy alta en al menos uno de los dos índices. El ISL describe una "
        f"condición del sustrato y no un nivel de peligro: su conversión en peligro exige la "
        f"integración con pendiente, precipitación detonante y cobertura, propia del "
        f"Entregable 8.",

        f"La pendiente media ponderada del ámbito es de {r['pend_ponderada']:.2f} % según el "
        f"rótulo del catálogo y de {D.ponderada(bs, 'pend_si_grados_pct'):.2f} % bajo la lectura "
        f"en grados que sostiene el Informe DT Consolidado de Frías. Bajo CUALQUIERA de las dos "
        f"lecturas, ningún bloque del ámbito supera el criterio de idoneidad de pendiente máxima "
        f"del proyecto (75 %): el máximo provincial es de "
        f"{max(b['pend_si_grados_pct'] for b in D.con_dato(bs, 'pend_si_grados_pct')):.2f} % en "
        f"la lectura más exigente. La elegibilidad por pendiente no está en cuestión; sí lo está "
        f"el dimensionamiento de las partidas de movimiento de tierras, que depende de cuál de "
        f"las dos unidades sea la correcta (discrepancia D-P01).",

        f"El inventario de cárcavas codificadas cubre {r['bloques_con_carcavas']} de los "
        f"{r['n']} bloques, con {r['n_carcavas']} rasgos y {r['long_carcavas']:,.2f} m de "
        f"longitud acumulada. La cobertura parcial del inventario impide fijar metas físicas de "
        f"infraestructura gris para el conjunto del ámbito."
        if r["n_carcavas"] else
        f"No se dispone de cárcavas codificadas en los bloques del ámbito, lo que impide fijar "
        f"metas físicas de infraestructura gris. La ausencia de registros no acredita ausencia "
        f"de cárcavas.",

        "El vacío de información de mayor consecuencia es la falta del peligro integrado "
        "(MCA-AHP) por bloque, ausente en la totalidad de las fichas del ámbito. Mientras no se "
        "incorpore desde el modelamiento de mesolocalización, la jerarquización de la cartera "
        "descansa en descriptores parciales y no en el peligro propiamente dicho.",

        "El diagnóstico se ha elaborado bajo el principio de no invención de datos: ningún "
        "valor ausente ha sido imputado ni estimado, y las discrepancias entre campo, gabinete "
        "y catálogo se documentan con el valor adoptado y la acción requerida, sin resolverse "
        "por criterio del redactor.",
    ]
    for i, c in enumerate(conc, start=1):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(8)
        rr = p.add_run(f"{i}. ")
        rr.bold = True
        rr.font.name = "Arial"
        rr.font.size = Pt(10)
        r2 = p.add_run(c)
        r2.font.name = "Arial"
        r2.font.size = Pt(10)


def _anexo(doc, prov, r, bs, meta):
    doc.add_page_break()
    sec = doc.add_section()
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, m, Cm(1.5))
    doc.add_heading(f"Anexo A. Inventario geoespacial consolidado de los {r['n']} bloques",
                    level=1)
    _p(doc, "Sistema de referencia UTM WGS 84 Zona 17S (EPSG:32717). El detalle completo, con "
            "las matrices de lectura integrada de índices de vegetación, de integración con el "
            "ISL, de pendientes, de cárcavas, de síntesis F-DT y de consistencia, figura en el "
            "anexo de matrices en formato Excel que acompaña a este informe.", size=9)
    filas = []
    for b in bs:
        filas.append([b["codigo"], b["distrito"], b["microcuenca"], round(b["area_ha"], 2),
                      f"{b['este']:,.0f}", f"{b['norte']:,.0f}",
                      (f"{b['alt_min']:,.0f}–{b['alt_max']:,.0f}"
                       if b["alt_min"] is not None else "s/d"),
                      round(b["pend_pct"], 2) if b["pend_pct"] is not None else "s/d",
                      round(b["msavi"], 4), round(b["pct_bajo_umbral"] or 0, 2),
                      f"{b['isl_mm']:.2f}", f"{b['isl_eh']:.2f}", b["n_carcavas"],
                      b["urgencia"]])
    filas.append(["TOTAL", "", "", round(r["ha"], 2), "", "",
                  f"{r['alt_min']:,.0f}–{r['alt_max']:,.0f}",
                  round(r["pend_media"], 2), round(r["msavi_medio"], 4),
                  round(r["pct_brecha"], 2), "", "", r["n_carcavas"], ""])
    _tabla(doc, ["Bloque", "Distrito", "Microcuenca", "Sup. (ha)", "Este (m)", "Norte (m)",
                 "Altitud (msnm)", "Pend. (%)", "MSAVI 2024", "% bajo umbral", "ISL-MM",
                 "ISL-EH", "Cárcavas", "Urgencia"], filas, size=7,
           fuente="Fuente: elaboración propia sobre el catálogo maestro Bloques V5/V6, la "
                  "estadística zonal del MDE, los compuestos Sentinel-2, el "
                  f"{meta['volumen_geologia']} y la capa «Cárcavas codificadas».",
           anchos=[1.5, 2.6, 2.2, 1.5, 1.7, 1.9, 2.3, 1.4, 1.5, 1.6, 1.3, 1.3, 1.4, 1.6])


def generar(prov, bloques, destino):
    bs = D.por_provincia(bloques, prov)
    r = D.resumen(bs)
    meta = D.PROV_META[prov]

    doc = Document()
    _estilos(doc)
    for s in doc.sections:
        s.left_margin = s.right_margin = Cm(2.2)
        s.top_margin = s.bottom_margin = Cm(2.0)

    _portada(doc, prov, r, bs)
    doc.add_heading("CONTENIDO", level=1)
    for t in ["1. Objeto, alcance y encuadre normativo",
              "2. Metodología y fuentes",
              "3. Caracterización general del ámbito",
              "4. Análisis geoespacial de los bloques",
              "5. Integración con el Estudio de Geología — Índice de Susceptibilidad Litológica (ISL)",
              "6. Procesos erosivos y cárcavas",
              "7. Caracterización diagnóstica — fichas F-DT consolidadas",
              "8. Análisis preliminar del riesgo en contexto de cambio climático (AdR-CCC)",
              "9. Medidas de reducción del riesgo y alternativas de infraestructura natural (MRR-CCC)",
              "10. Vacíos de información, consistencia y trazabilidad",
              "11. Conclusiones",
              f"Anexo A. Inventario geoespacial consolidado de los {r['n']} bloques"]:
        _p(doc, t, size=10, space_after=3)
    doc.add_page_break()

    for fn in (_s1, _s2, _s3, _s4, _s5, _s6, _s7, _s8, _s9, _s10, _s11):
        fn(doc, prov, r, bs, meta)
    _anexo(doc, prov, r, bs, meta)

    ruta = os.path.join(destino, f"Informe_DT_Consolidado_{meta['archivo']}_IN_Piura.docx")
    doc.save(ruta)
    return ruta


if __name__ == "__main__":
    bl = D.cargar_bloques()
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "entregables")
    os.makedirs(out, exist_ok=True)
    for p in D.PROVINCIAS:
        print("OK:", generar(p, bl, out))
