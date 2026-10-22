"""Inserta los mapas geológicos por bloque en el Volumen I (Morropón) del
Entregable 3 - Estudio de Geología del Proyecto IN Piura.

El documento trae, para cada bloque de intervención, un marcador
"FIGURA PENDIENTE DE INSERCIÓN" (tabla de una celda) seguido del pie de
figura "Figura N. Bloque <CÓDIGO> — delimitación de unidades geológicas...".
El script reemplaza cada marcador por el mapa del bloque correspondiente,
tomado del ATLAS DE GEOLOGIA BLOQUES V5 PIURA.

Los bloques adicionales de San Juan de Bigote (83 a 87) no tienen marcador:
su lámina viene en PDF, se rasteriza y se inserta sobre el pie de figura.

Uso:
    python insertar_figuras_geologia.py ORIGINAL.docx ATLAS_DIR PDF_DIR SALIDA.docx
"""

import argparse
import io
import re
import sys
from pathlib import Path

from PIL import Image
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Emu, Pt

MARCADOR = "FIGURA PENDIENTE DE INSERCIÓN"
RE_PIE = re.compile(r"^Figura\s+\d+(?:\.\d+)*\.\s+Bloque\s+(.+?)\s+—\s+delimitación")

# Bloques adicionales cuya lámina llega en PDF y no tiene marcador en el documento.
RE_PDF_BLOQUE = re.compile(r"BLOQUE[_\s]+(\d+)", re.IGNORECASE)

DPI_PDF = 150          # rasterizado de las láminas PDF
JPEG_QUALITY = 92      # 4:4:4, sin pérdida visible en las etiquetas del mapa
ALTO_MAX_FRACCION = 0.92   # deja aire para el pie de figura en la misma página


def codigo_de_pie(parrafo):
    """Devuelve el código de bloque de un pie de figura, o None."""
    m = RE_PIE.match(parrafo.text.strip())
    return m.group(1).strip() if m else None


def es_marcador(tabla):
    return MARCADOR in tabla._element.xml


def a_jpeg(ruta, destino_cache, calidad=JPEG_QUALITY):
    """Convierte a JPEG optimizado. Los PNG del atlas son RGBA de ~2.7 MB;
    en JPEG 4:4:4 q92 pesan ~0.7 MB sin degradar las etiquetas."""
    if destino_cache.exists():
        return destino_cache
    with Image.open(ruta) as im:
        rgb = im.convert("RGB")
        rgb.save(destino_cache, "JPEG", quality=calidad,
                 subsampling=0, optimize=True)
    return destino_cache


def rasterizar_pdf(ruta_pdf, destino_cache, calidad=JPEG_QUALITY):
    if destino_cache.exists():
        return destino_cache
    import pymupdf
    doc = pymupdf.open(ruta_pdf)
    pix = doc[0].get_pixmap(dpi=DPI_PDF)
    im = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    im.save(destino_cache, "JPEG", quality=calidad,
            subsampling=0, optimize=True)
    doc.close()
    return destino_cache


def medidas(ruta_img, ancho_max, alto_max):
    """Escala la imagen para caber en el área útil conservando proporción."""
    with Image.open(ruta_img) as im:
        w, h = im.size
    ancho = ancho_max
    alto = int(ancho * h / w)
    if alto > alto_max:
        alto = alto_max
        ancho = int(alto * w / h)
    return Emu(ancho), Emu(alto)


def parrafo_imagen(doc, ruta_img, ancho, alto):
    """Crea un párrafo centrado con la imagen. Se añade al final del cuerpo;
    el llamador lo reubica con addprevious()."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(ruta_img), width=ancho, height=alto)
    return p._p


def renumerar_ids_dibujo(doc):
    """python-docx numera cada imagen desde 1, de modo que las nuevas chocan
    con las que ya traía el documento. Word considera dañado un archivo con
    identificadores wp:docPr repetidos, así que se renumeran todos."""
    NS_WP = ("{http://schemas.openxmlformats.org/drawingml/2006/"
             "wordprocessingDrawing}")
    NS_PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"
    siguiente = 1
    for anclaje in doc.element.body.iter():
        if anclaje.tag not in (f"{NS_WP}inline", f"{NS_WP}anchor"):
            continue
        docpr = anclaje.find(f"{NS_WP}docPr")
        if docpr is None:
            continue
        docpr.set("id", str(siguiente))
        nombre = docpr.get("name") or f"Imagen {siguiente}"
        cnvpr = anclaje.find(f".//{NS_PIC}nvPicPr/{NS_PIC}cNvPr")
        if cnvpr is not None:
            cnvpr.set("id", str(siguiente))
            cnvpr.set("name", nombre)
        siguiente += 1
    return siguiente - 1


def recorrer(doc):
    """Devuelve el cuerpo como lista ordenada de ('p'|'tbl', elemento)."""
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    items = []
    for hijo in doc.element.body.iterchildren():
        if hijo.tag.endswith("}p"):
            items.append(("p", Paragraph(hijo, doc)))
        elif hijo.tag.endswith("}tbl"):
            items.append(("tbl", Table(hijo, doc)))
    return items


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("original")
    ap.add_argument("atlas_dir")
    ap.add_argument("pdf_dir")
    ap.add_argument("salida")
    ap.add_argument("--cache", default=None,
                    help="carpeta para los JPEG intermedios")
    ap.add_argument("--calidad", type=int, default=JPEG_QUALITY,
                    help="calidad JPEG (por defecto 92); bájela para un "
                         "archivo más ligero de revisión")
    args = ap.parse_args()

    atlas_dir = Path(args.atlas_dir)
    pdf_dir = Path(args.pdf_dir)
    cache = Path(args.cache or (Path(args.salida).parent / "_cache_img"))
    cache.mkdir(parents=True, exist_ok=True)

    # Índice del atlas: el nombre del archivo es el código del bloque.
    atlas = {p.stem: p for p in atlas_dir.glob("*.png")}
    atlas.update({p.stem: p for p in atlas_dir.glob("*.jpg")})

    # Índice de las láminas PDF por número de bloque.
    pdfs = {}
    for p in pdf_dir.glob("*.pdf"):
        m = RE_PDF_BLOQUE.search(p.stem)
        if m:
            pdfs[str(int(m.group(1)))] = p

    doc = Document(args.original)
    seccion = doc.sections[0]
    ancho_max = seccion.page_width - seccion.left_margin - seccion.right_margin
    alto_max = int((seccion.page_height - seccion.top_margin
                    - seccion.bottom_margin) * ALTO_MAX_FRACCION)

    items = recorrer(doc)
    pies = [(i, obj) for i, (k, obj) in enumerate(items)
            if k == "p" and codigo_de_pie(obj)]
    marcadores = [i for i, (k, obj) in enumerate(items)
                  if k == "tbl" and es_marcador(obj)]

    print(f"Pies de figura: {len(pies)} | Marcadores: {len(marcadores)}")

    insertadas, faltantes = [], []

    for idx_pie, pie in pies:
        codigo = codigo_de_pie(pie)
        # ¿Hay un marcador inmediatamente antes de este pie?
        previos = [m for m in marcadores if 0 < idx_pie - m <= 3]
        marcador = items[previos[-1]][1] if previos else None

        if codigo in atlas:
            img = a_jpeg(atlas[codigo], cache / f"{codigo}.jpg",
                         args.calidad)
        elif codigo in pdfs:
            img = rasterizar_pdf(pdfs[codigo], cache / f"pdf_{codigo}.jpg",
                                 args.calidad)
        else:
            faltantes.append(codigo)
            continue

        ancho, alto = medidas(img, ancho_max, alto_max)
        nuevo = parrafo_imagen(doc, img, ancho, alto)

        if marcador is not None:
            # Reemplaza el marcador: se coloca la imagen en su lugar y se elimina.
            marcador._element.addprevious(nuevo)
            marcador._element.getparent().remove(marcador._element)
            origen = "marcador"
        else:
            # Bloque adicional: la imagen va justo encima del pie de figura.
            pie._p.addprevious(nuevo)
            origen = "sobre el pie"

        insertadas.append((codigo, origen, alto / 914400))

    total_ids = renumerar_ids_dibujo(doc)
    doc.save(args.salida)

    print(f"Insertadas: {len(insertadas)} | ids de dibujo renumerados: {total_ids}")
    if faltantes:
        print(f"SIN IMAGEN ({len(faltantes)}): {', '.join(faltantes)}",
              file=sys.stderr)

    # Verificación: no debe quedar ningún marcador en el documento final.
    rev = Document(args.salida)
    restantes = sum(1 for t in rev.tables if MARCADOR in t._element.xml)
    print(f"Marcadores restantes: {restantes}")
    graficos = rev.element.body.findall(
        ".//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline")
    print(f"Imágenes en el documento: {len(graficos)}")
    return 1 if (faltantes or restantes) else 0


if __name__ == "__main__":
    raise SystemExit(main())
