"""Atlas de ecosistemas de los 117 bloques de intervención V6 — Proyecto IN Piura.

Una ficha A4 horizontal por bloque (ordenadas por provincia, distrito y código), precedida de una portada con
índice. Cada ficha muestra los ecosistemas dentro del bloque y en su entorno, el área de influencia aprobada,
los bloques vecinos, la ubicación en el ámbito, el cuadro de ha por ecosistema y los datos del bloque.

Usa las mismas capas y la misma paleta que mapas/mapa_ecosistemas_bloques_v6.py (capa base de datos/gis/ más
las capas «complemento»). Salida: mapas/salidas/IN_Piura_Atlas_Ecosistemas_Bloques_V6.pdf
y, con --png, una imagen por bloque en mapas/salidas/atlas_ecosistemas/.

Uso:  python mapas/atlas_ecosistemas_bloques_v6.py [--png] [--bloques M27B1 3 ...]
"""
import argparse
import re
import sys
import textwrap
from pathlib import Path

import geopandas as gpd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.lines import Line2D
from matplotlib.patches import Patch, Rectangle
from shapely.geometry import box

sys.path.insert(0, str(Path(__file__).resolve().parent))
from mapa_area_influencia_bloques_v6 import (  # noqa: E402
    AI_EDGE, DORADO, NOMBRE_PROV, TOTAL_V6_HA, VERDE, ejes_utm, norte_y_escala, titulo_cdl)
from mapa_ecosistemas_bloques_v6 import SALIDA, SIN_DATO, preparar  # noqa: E402

A4 = (11.69, 8.27)
FUENTE = ("Fuente: Ecosistemas — Mapa Nacional de Ecosistemas del Perú, MINAM 2018 (R.M. N.° 440-2018-MINAM), recorte "
          "a microcuencas completado en M27B1 y M9B1. Bloques V6 y área de influencia aprobada (ANIN-DIME-SESDI, 2026). "
          "Elaboración: ANIN-DIME-SESDI.")
ORDEN_PROV = ["AYABACA", "HUANCABAMBA", "MORROPON"]


def clave_natural(cod):
    """'4' < '10' < 'M2B1' < 'M10B4': números como números, con los códigos M al final."""
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", cod)]


def encabezado(fig, titulo, subtitulo, derecha):
    fig.add_artist(Rectangle((0.008, 0.012), 0.984, 0.976, transform=fig.transFigure, fill=False,
                             edgecolor=VERDE, linewidth=1.6))
    fig.add_artist(Rectangle((0.008, 0.905), 0.984, 0.083, transform=fig.transFigure, facecolor=VERDE,
                             edgecolor=VERDE))
    fig.text(0.02, 0.962, "AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN", color="white", fontsize=10,
             fontweight="bold", va="center")
    fig.text(0.02, 0.938, "DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME", color="white",
             fontsize=6.8, va="center")
    fig.text(0.02, 0.919, "SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN - SESDI", color="white", fontsize=6.8, va="center")
    fig.text(0.98, 0.962, "PROYECTO IN PIURA · CUI 2669244", color=DORADO, fontsize=9.5, fontweight="bold",
             va="center", ha="right")
    fig.text(0.98, 0.93, "Atlas de ecosistemas — bloques de intervención V6", color="white", fontsize=7,
             va="center", ha="right")
    fig.text(0.02, 0.875, titulo, color=VERDE, fontsize=12.5, fontweight="bold", va="center")
    fig.text(0.02, 0.85, subtitulo, color="#333333", fontsize=7.8, va="center")
    fig.text(0.98, 0.875, derecha, color=VERDE, fontsize=9, fontweight="bold", ha="right", va="center")
    fig.text(0.98, 0.85, "Formato A4 · Octubre 2026", color="#555555", fontsize=6.8, ha="right", va="center")


def pie(fig):
    fig.text(0.02, 0.034, "Sistema de coordenadas: UTM WGS84 Zona 17S (EPSG:32717)", fontsize=6.5,
             color="#333333", fontweight="bold")
    fig.text(0.02, 0.02, FUENTE, fontsize=5.4, color="#555555")


def ventana(geom, aspecto, margen=0.45, minimo=1500):
    """Encuadre centrado en el bloque con margen proporcional y lado mínimo, ajustado al aspecto del eje."""
    xmin, ymin, xmax, ymax = geom.bounds
    cx, cy = (xmin + xmax) / 2, (ymin + ymax) / 2
    w = max(xmax - xmin, minimo) * (1 + 2 * margen)
    h = max(ymax - ymin, minimo) * (1 + 2 * margen)
    if w / h < aspecto:
        w = h * aspecto
    else:
        h = w / aspecto
    return (cx - w / 2, cx + w / 2), (cy - h / 2, cy + h / 2)


def ficha(r, b, ai, amb, ctx, eco, x, colores, eco_orden, n, total):
    cod = r.BLOQUE
    xb = x[x.BLOQUE == cod].sort_values("HA_OFICIAL", ascending=False)
    aib = ai[ai.BLOQUE == cod]
    fig = plt.figure(figsize=A4)
    encabezado(fig, f"ECOSISTEMAS DEL BLOQUE {cod}",
               f"Microcuenca {r.MICROC} · Distrito de {titulo_cdl(r.NOMBDIST)} · Provincia de "
               f"{NOMBRE_PROV[r.NOMBPROV]} · {r.AREA_HA:,.3f} ha", f"FICHA {n} de {total}")

    # ---- mapa
    ax = fig.add_axes([0.045, 0.085, 0.565, 0.735])
    xlim, ylim = ventana(r.geometry, aspecto=0.565 * A4[0] / (0.735 * A4[1]))
    caja = box(xlim[0], ylim[0], xlim[1], ylim[1])
    ax.set_facecolor("white")
    ctx.cx[xlim[0]:xlim[1], ylim[0]:ylim[1]].plot(ax=ax, facecolor="#F2F2F2", edgecolor="#BFBFBF",
                                                   linewidth=0.3, zorder=1)
    amb.cx[xlim[0]:xlim[1], ylim[0]:ylim[1]].plot(ax=ax, facecolor="white", edgecolor="#7F8C86",
                                                   linewidth=0.5, zorder=1.5)
    ev = gpd.clip(eco, caja)
    if len(ev):
        # entorno atenuado y, encima, el ecosistema dentro del bloque a color pleno
        ev.plot(ax=ax, color=ev.ECOSISTEMA.map(colores), alpha=0.38, edgecolor="none", zorder=2, rasterized=True)
        dentro = gpd.clip(ev, r.geometry)
        if len(dentro):
            dentro.plot(ax=ax, color=dentro.ECOSISTEMA.map(colores), edgecolor="none", zorder=3, rasterized=True)
    aib.plot(ax=ax, facecolor="none", edgecolor=AI_EDGE, linewidth=0.9, linestyle=(0, (4, 2)), zorder=4)
    vecinos = b[(b.BLOQUE != cod) & b.intersects(caja)]
    vecinos.plot(ax=ax, facecolor="none", edgecolor="#555555", linewidth=0.7, zorder=4.5)
    for v in vecinos.itertuples():
        p = v.geometry.representative_point()
        if caja.contains(p):
            ax.text(p.x, p.y, v.BLOQUE, fontsize=6, color="#333333", ha="center", va="center", zorder=6,
                    bbox=dict(boxstyle="round,pad=0.1", facecolor="white", edgecolor="none", alpha=0.7))
    gpd.GeoSeries([r.geometry]).plot(ax=ax, facecolor="none", edgecolor="white", linewidth=2.6, zorder=5)
    gpd.GeoSeries([r.geometry]).plot(ax=ax, facecolor="none", edgecolor="#C0392B", linewidth=1.4, zorder=5.5)
    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)
    w = xlim[1] - xlim[0]
    paso = next(p for p in (500, 1000, 2000, 5000) if w / p <= 7)
    ejes_utm(ax, paso=paso)
    ax.tick_params(labelsize=5.5)
    ax.xaxis.label.set_size(6.5)
    ax.yaxis.label.set_size(6.5)
    norte_y_escala(ax, next(k for k in (5, 2, 1, 0.5, 0.25) if k * 1000 <= w / 5))

    # ---- panel derecho
    panel = fig.add_axes([0.64, 0.06, 0.34, 0.76])
    panel.axis("off")
    ub = fig.add_axes([0.645, 0.6, 0.15, 0.215])
    amb.plot(ax=ub, facecolor="#E8E1C8", edgecolor="#9AA5A0", linewidth=0.3, rasterized=True)
    amb.dissolve().boundary.plot(ax=ub, color=VERDE, linewidth=0.8, rasterized=True)
    b.plot(ax=ub, facecolor="#C9C9C9", edgecolor="none", rasterized=True)
    p = r.geometry.representative_point()
    ub.plot(p.x, p.y, marker="*", color="#C0392B", markersize=11, markeredgecolor="white", markeredgewidth=0.6)
    ub.set_xticks([]); ub.set_yticks([]); ub.set_aspect("equal")
    ub.set_xlabel(""); ub.set_ylabel("")
    ub.set_title("Ubicación en el ámbito", fontsize=7, fontweight="bold", color=VERDE, pad=2)
    for s in ub.spines.values():
        s.set_edgecolor(VERDE)
    datos = [("Código", cod), ("Microcuenca", r.MICROC), ("Distrito", titulo_cdl(r.NOMBDIST)),
             ("Provincia", NOMBRE_PROV[r.NOMBPROV]), ("Área (catálogo)", f"{r.AREA_HA:,.3f} ha"),
             ("Área de influencia", f"{aib.AREA_HA.sum():,.3f} ha"),
             ("Punto interior", f"E {p.x:,.0f} · N {p.y:,.0f}"),
             ("Dominante", f"{xb.iloc[0].SIMBOLO or xb.iloc[0].ECOSISTEMA} ({xb.iloc[0].PCT:.0%})")]
    y = 0.995
    for k, v in datos:
        panel.text(0.49, y, k, fontsize=6.4, color="#555555", transform=panel.transAxes, va="top")
        panel.text(0.69, y, v, fontsize=6.4, color="#1B1B1B", fontweight="bold", transform=panel.transAxes,
                   va="top")
        y -= 0.034

    # cuadro de ecosistemas del bloque
    filas = [[e if len(e) <= 46 else e[:45] + "…", s or "—", f"{h:,.3f}", f"{pc:.1%}"]
             for e, s, h, pc in zip(xb.ECOSISTEMA, xb.SIMBOLO, xb.HA_OFICIAL, xb.PCT)]
    filas.append(["TOTAL", "", f"{xb.HA_OFICIAL.sum():,.3f}", "100.0%"])
    y0 = 0.67
    alto = 0.034 * (len(filas) + 1)
    panel.text(0.0, y0 + 0.012, "ECOSISTEMAS EN EL BLOQUE", fontsize=7.5, fontweight="bold", color=VERDE,
               transform=panel.transAxes)
    t = panel.table(cellText=filas, colLabels=["Ecosistema", "Símbolo", "Área (ha)", "%"],
                    colWidths=[0.58, 0.14, 0.16, 0.12], cellLoc="left", bbox=[0.0, y0 - alto, 1.0, alto])
    t.auto_set_font_size(False)
    t.set_fontsize(6)
    for (i, j), c in t.get_celld().items():
        c.set_edgecolor("#B0B7B3"); c.set_linewidth(0.4)
        if j >= 2:
            c.set_text_props(ha="right")
        if i == 0:
            c.set_facecolor(VERDE); c.set_text_props(color="white", fontweight="bold", ha="center")
        elif i == len(filas):
            c.set_facecolor("#D5E8D4"); c.set_text_props(fontweight="bold")
        elif i % 2 == 0:
            c.set_facecolor("#F2F7F2")

    # leyenda: solo lo que aparece en la ventana
    vistos = set(ev.ECOSISTEMA) if len(ev) else set()
    items = [Patch(facecolor=colores[e], edgecolor="#7F8C86", linewidth=0.3, label=lab)
             for e, lab in eco_orden if e in vistos and e != SIN_DATO]
    items += [
        Patch(facecolor="#BDBDBD", alpha=0.45, edgecolor="none", label="Color atenuado: ecosistema fuera del bloque"),
        Line2D([], [], color="#C0392B", linewidth=1.6, label=f"Bloque {cod}"),
        Line2D([], [], color=AI_EDGE, linewidth=1, linestyle=(0, (4, 2)), label="Área de influencia aprobada"),
        Line2D([], [], color="#555555", linewidth=0.8, label="Otros bloques V6"),
        Patch(facecolor="white", edgecolor="#7F8C86", linewidth=0.4, label="Ámbito sin cobertura de la capa"),
        Patch(facecolor="#F2F2F2", edgecolor="#BFBFBF", label="Fuera del ámbito"),
    ]
    if SIN_DATO in set(xb.ECOSISTEMA):
        items.insert(0, Patch(facecolor="white", edgecolor="#7F8C86", label=SIN_DATO + " (dentro del bloque)"))
    panel.legend(handles=items, loc="upper left", bbox_to_anchor=(0.0, y0 - alto - 0.02), fontsize=6,
                 frameon=False, title="LEYENDA", title_fontproperties={"weight": "bold", "size": 7},
                 alignment="left", handlelength=1.8, handleheight=0.8, labelspacing=0.28)
    pie(fig)
    return fig


def portada(b, x, total):
    fig = plt.figure(figsize=A4)
    encabezado(fig, "ATLAS DE ECOSISTEMAS DE LOS BLOQUES DE INTERVENCIÓN V6",
               f"{total} bloques · {TOTAL_V6_HA:,.3f} ha · 15 distritos de Ayabaca, Huancabamba y Morropón — "
               "Región Piura", "ÍNDICE")
    resumen = x.groupby("ECOSISTEMA").HA_OFICIAL.sum().sort_values(ascending=False)
    texto = "Superficie de los bloques por ecosistema: " + " · ".join(
        f"{e} {h:,.1f} ha ({h / TOTAL_V6_HA:.1%})" for e, h in resumen.items() if h >= 0.05)
    fig.text(0.02, 0.825, textwrap.fill(texto, 175), fontsize=6.3, color="#333333", va="top")
    dom = x.sort_values("HA_OFICIAL", ascending=False).drop_duplicates("BLOQUE").set_index("BLOQUE")
    lineas, pag = [], 2
    for prov in ORDEN_PROV:
        sub = b[b.NOMBPROV == prov]
        lineas.append((f"{NOMBRE_PROV[prov].upper()} — {len(sub)} bloques", "", "", True))
        for dist, g in sub.groupby("NOMBDIST", sort=False):
            lineas.append((f"  {titulo_cdl(dist)}", "", "", True))
            for r in g.itertuples():
                d = dom.loc[r.BLOQUE]
                lineas.append((f"     {r.BLOQUE}", f"{r.AREA_HA:,.2f} ha · {d.SIMBOLO or d.ECOSISTEMA} {d.PCT:.0%}",
                               f"{pag}", False))
                pag += 1
    filas_col = -(-len(lineas) // 3)
    for c in (1, 2):  # un título de distrito no queda solo al pie de una columna
        fin = c * filas_col - 1
        if fin < len(lineas) and lineas[fin][3]:
            lineas.insert(fin, ("", "", "", False))
    filas_col = -(-len(lineas) // 3)
    for i, (a, m, pg, neg) in enumerate(lineas):
        c, f = divmod(i, filas_col)
        x0, y = 0.02 + c * 0.325, 0.775 - f * (0.715 / filas_col)
        fig.text(x0, y, a, fontsize=5.6, fontweight="bold" if neg else "normal",
                 color=VERDE if neg else "#222222")
        if m:
            fig.text(x0 + 0.075, y, m, fontsize=5.4, color="#555555")
            fig.text(x0 + 0.3, y, pg, fontsize=5.6, ha="right", color="#222222")
    fig.text(0.98, 0.055, "Página del atlas · ecosistema dominante y su % del bloque", fontsize=5.6,
             color="#555555", ha="right")
    pie(fig)
    return fig


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--png", action="store_true", help="guardar también una imagen PNG por bloque")
    ap.add_argument("--bloques", nargs="*", help="solo estos códigos (prueba rápida)")
    ap.add_argument("--salida", default=str(SALIDA))
    a = ap.parse_args()
    salida = Path(a.salida)

    b, ai, amb, ctx, eco, x, colores, eco_orden = preparar()
    orden = sorted(range(len(b)), key=lambda i: (ORDEN_PROV.index(b.NOMBPROV.iat[i]), b.NOMBDIST.iat[i],
                                                 clave_natural(b.BLOQUE.iat[i])))
    b = b.iloc[orden]
    if a.bloques:
        faltan = set(a.bloques) - set(b.BLOQUE)
        if faltan:
            sys.exit(f"Códigos que no están en los bloques V6: {sorted(faltan)}")
        b_sel = b[b.BLOQUE.isin(a.bloques)]
    else:
        b_sel = b
    salida.mkdir(parents=True, exist_ok=True)
    dir_png = salida / "atlas_ecosistemas"
    if a.png:
        dir_png.mkdir(exist_ok=True)
    nombre = "IN_Piura_Atlas_Ecosistemas_Bloques_V6.pdf" if not a.bloques else "IN_Piura_Atlas_Ecosistemas_prueba.pdf"
    with PdfPages(salida / nombre) as pp:
        fig = portada(b, x, len(b))
        pp.savefig(fig, dpi=150)
        plt.close(fig)
        for n, r in enumerate(b_sel.itertuples(), start=1):
            fig = ficha(r, b, ai, amb, ctx, eco, x, colores, eco_orden, n, len(b_sel))
            pp.savefig(fig, dpi=150)  # ecosistemas en raster: el PDF queda liviano
            if a.png:
                fig.savefig(dir_png / f"Ecosistemas_Bloque_{r.BLOQUE}.png", dpi=150)
            plt.close(fig)
            if n % 20 == 0:
                print(f"  {n} fichas…")
    print(f"Atlas: {len(b_sel)} fichas + índice → {salida / nombre}")


if __name__ == "__main__":
    main()
