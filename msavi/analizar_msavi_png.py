#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Mide la distribucion areal de las clases MSAVI 2024 por bloque a partir de las
cartografias tematicas PNG de la carpeta Drive «MSAVI BLOQUES V6».

AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN / DIME / SESDI
Proyecto IN Piura | CUI 2669244 | UTM WGS 84 Zona 17S (EPSG:32717)

Procedimiento (detallado en METODOLOGIA_MSAVI.md):
  1. recorte al interior del marco cartografico, que deja fuera la leyenda;
  2. clasificacion de celdas por igualdad exacta con la paleta RdYlGn de 5 clases;
  3. agrupacion en componentes conexos y adjudicacion al bloque del componente
     que contiene el centro de la lamina (el atlas encuadra cada bloque);
  4. reparto porcentual; las hectareas se calculan despues, en
     `completar_msavi_plantillas.py`, sobre la superficie de catalogo del libro.

Uso:
    python analizar_msavi_png.py --png <carpeta_png> \
        --interseccion BLOQUES_V6_INTERSECC_MSAVI.md \
        --manifiesto ../datos/manifiesto_resumenes_117.json \
        --salida msavi_distribucion_117_bloques.csv
"""

import argparse
import collections
import csv
import glob
import json
import os
import re
import sys

import numpy as np
from PIL import Image
from scipy import ndimage

# Paleta RdYlGn de 5 clases del producto MSAVI_Piura_2024, Banda 1 (Gray)
PALETA = {1: (0xD7, 0x19, 0x1C),   # <= 0.2650   suelo desnudo / no vegetal
          2: (0xFD, 0xAE, 0x61),   # 0.2650-0.3813  vigor muy bajo
          3: (0xFF, 0xFF, 0xC0),   # 0.3813-0.4976  vigor bajo
          4: (0xA6, 0xD9, 0x6A),   # 0.4976-0.6139  vigor moderado
          5: (0x1A, 0x96, 0x41)}   # > 0.6139    vigor alto

# Interior del marco cartografico; la maqueta es identica en las 130 laminas.
IZQ, DER, SUP, INF = 745, 3409, 99, 2243

INTERPRETACION = {5: "Vigor alto", 4: "Vigor moderado", 3: "Vigor bajo",
                  2: "Vigor muy bajo", 1: "Suelo desnudo / no vegetal"}


def codigo_de_lamina(nombre):
    """'MSAVI Bloque 23.png' -> '23'; 'MSAVI M6B2-1.png' -> 'M6B2-1'."""
    txt = os.path.splitext(nombre)[0]
    txt = re.sub(r"^MSAVI\s+", "", txt, flags=re.I)
    txt = re.sub(r"^BL?oque\s+", "", txt, flags=re.I)
    return txt.strip()


def mide_lamina(ruta):
    """Devuelve la lista de componentes de la lamina, de mayor a menor."""
    img = np.array(Image.open(ruta).convert("RGB"))
    if img.shape[0] <= INF or img.shape[1] <= DER:
        raise ValueError("la lamina no tiene la maqueta esperada: %s" % (img.shape,))
    rec = img[SUP:INF + 1, IZQ:DER + 1]
    alto, ancho = rec.shape[:2]
    r, g, b = rec[:, :, 0], rec[:, :, 1], rec[:, :, 2]

    clases = np.zeros((alto, ancho), np.uint8)
    for clave, (cr, cg, cb) in PALETA.items():
        clases[(r == cr) & (g == cg) & (b == cb)] = clave

    ocupado = clases > 0
    # el cierre salva los cortes de grilla, curvas de nivel y rotulos
    solido = ndimage.binary_fill_holes(
        ndimage.binary_closing(ocupado, np.ones((11, 11))))
    etiquetas, cuantos = ndimage.label(solido, np.ones((3, 3)))
    if cuantos == 0:
        return []
    tam = ndimage.sum(ocupado, etiquetas, range(1, cuantos + 1))
    total = tam.sum()
    centros = ndimage.center_of_mass(solido, etiquetas, range(1, cuantos + 1))
    cy, cx = alto / 2.0, ancho / 2.0

    componentes = []
    for i in np.argsort(-tam):
        if tam[i] < max(total * 0.01, 300):
            break
        sel = etiquetas == i + 1
        componentes.append(dict(
            px=int(tam[i]),
            clases={k: int(((clases == k) & sel).sum()) for k in PALETA},
            distancia=float(((centros[i][0] - cy) ** 2 + (centros[i][1] - cx) ** 2) ** 0.5),
            contiene_centro=bool(solido[int(cy), int(cx)] and
                                 etiquetas[int(cy), int(cx)] == i + 1)))
    return componentes


def adjudica(componentes):
    """El bloque de la lamina es el componente que contiene el centro; si ninguno
    lo contiene, el de centroide mas proximo."""
    con_centro = [c for c in componentes if c["contiene_centro"]]
    return con_centro[0] if con_centro else min(componentes, key=lambda c: c["distancia"])


def lee_interseccion(ruta):
    """Superficie, microcuenca, distrito y provincia por bloque a partir del
    reporte de interseccion.  Los registros se parten en varias lineas cuando el
    nombre del distrito es largo, de ahi el ensamblado por buffer.

    AVISO: AREA_M2 es el area del poligono PADRE del bloque, repetida en cada
    pieza de interseccion; por eso solo se suman los valores distintos
    consecutivos (un poligono por valor) y NO se usa para repartir por clase.
    """
    saltar = ("Page ", "FID BLOQUE", "B6_Intersect", "Sum ", "Count ",
              "Mean ", "Max ", "Min ", "Standard")
    inicio = re.compile(r"^(\d+)\s")
    cola = re.compile(r"(\d+(?:\.\d+)?)\s+([1-5])$")
    registros = []
    buffer_ = None
    with open(ruta, encoding="utf-8") as fh:
        for linea in fh:
            txt = linea.strip()
            if not txt or txt.startswith(saltar):
                continue
            if inicio.match(txt) and buffer_ is None:
                buffer_ = txt
            elif buffer_ is not None:
                buffer_ += " " + txt
            else:
                continue
            if cola.search(buffer_) and len(buffer_.split()) >= 7:
                p = buffer_.split()
                registros.append((p[1], p[2], " ".join(p[3:-3]), p[-3], float(p[-2])))
                buffer_ = None

    por_bloque = collections.OrderedDict()
    for bloque, micro, distrito, provincia, area in registros:
        d = por_bloque.setdefault(bloque, dict(area_m2=0.0, previa=None, micro=micro,
                                               distrito=distrito, provincia=provincia))
        if d["previa"] is None or abs(area - d["previa"]) > 1e-9:
            d["area_m2"] += area
        d["previa"] = area
    return {b: dict(area_ha=d["area_m2"] / 1e4, micro=d["micro"],
                    distrito=d["distrito"], provincia=d["provincia"])
            for b, d in por_bloque.items()}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--png", required=True, help="carpeta con las cartografias PNG")
    ap.add_argument("--interseccion", required=True, help="BLOQUES_V6_INTERSECC_MSAVI.md")
    ap.add_argument("--manifiesto", required=True, help="manifiesto_resumenes_117.json")
    ap.add_argument("--salida", default="msavi_distribucion_117_bloques.csv")
    args = ap.parse_args(argv)

    padron = [b["codigo"] for b in json.load(open(args.manifiesto, encoding="utf-8"))["bloques"]]
    catastro = lee_interseccion(args.interseccion)

    laminas = {}
    for ruta in sorted(glob.glob(os.path.join(args.png, "*.png"))):
        nombre = os.path.basename(ruta)
        try:
            laminas[codigo_de_lamina(nombre)] = (nombre, mide_lamina(ruta))
        except Exception as exc:                       # noqa: BLE001
            print("  [!] %s: %s" % (nombre, exc), file=sys.stderr)
    print("laminas medidas: %d" % len(laminas))

    filas = []
    for codigo in padron:
        if codigo not in laminas:
            print("  [!] bloque %s sin cartografia" % codigo, file=sys.stderr)
            continue
        nombre, componentes = laminas[codigo]
        if not componentes:
            print("  [!] bloque %s sin poligono clasificado" % codigo, file=sys.stderr)
            continue
        elegido = adjudica(componentes)
        cuenta = elegido["clases"]
        total = sum(cuenta.values())
        if total < 500:
            print("  [!] bloque %s con solo %d celdas" % (codigo, total), file=sys.stderr)
            continue
        pct = {k: 100.0 * cuenta[k] / total for k in cuenta}
        info = catastro.get(codigo, {})
        area = info.get("area_ha", 0.0)
        bajo = pct[1] + pct[2] + pct[3]
        modal = max(pct, key=pct.get)
        filas.append(collections.OrderedDict([
            ("BLOQUE", codigo),
            ("MICROCUENCA", info.get("micro", "")),
            ("DISTRITO", info.get("distrito", "")),
            ("PROVINCIA", info.get("provincia", "")),
            ("AREA_HA", round(area, 3)),
            ("PCT_MAYOR_0_6139", round(pct[5], 2)),
            ("PCT_0_4976_0_6139", round(pct[4], 2)),
            ("PCT_0_3813_0_4976", round(pct[3], 2)),
            ("PCT_0_2650_0_3813", round(pct[2], 2)),
            ("PCT_MENOR_IGUAL_0_2650", round(pct[1], 2)),
            ("HA_MAYOR_0_6139", round(area * pct[5] / 100, 3)),
            ("HA_0_4976_0_6139", round(area * pct[4] / 100, 3)),
            ("HA_0_3813_0_4976", round(area * pct[3] / 100, 3)),
            ("HA_0_2650_0_3813", round(area * pct[2] / 100, 3)),
            ("HA_MENOR_IGUAL_0_2650", round(area * pct[1] / 100, 3)),
            ("PCT_BAJO_UMBRAL_0_4976", round(bajo, 2)),
            ("HA_BAJO_UMBRAL_0_4976", round(area * bajo / 100, 3)),
            ("PCT_SOBRE_UMBRAL_0_4976", round(100 - bajo, 2)),
            ("CLASE_MODAL", INTERPRETACION[modal]),
            ("PIXELES", total),
            ("POLIG_EN_MAPA", len(componentes)),
            ("MAPA_PNG", nombre),
        ]))

    filas.sort(key=lambda f: (len(f["BLOQUE"]), f["BLOQUE"]))
    with open(args.salida, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=list(filas[0].keys()), delimiter=";")
        w.writeheader()
        w.writerows(filas)

    total_ha = sum(f["AREA_HA"] for f in filas)
    bajo_ha = sum(f["HA_BAJO_UMBRAL_0_4976"] for f in filas)
    print("bloques: %d   superficie: %.1f ha" % (len(filas), total_ha))
    print("bajo el umbral 0.4976: %.1f ha (%.1f %%)" % (bajo_ha, 100 * bajo_ha / total_ha))
    print("escrito %s" % args.salida)
    return 0


if __name__ == "__main__":
    sys.exit(main())
