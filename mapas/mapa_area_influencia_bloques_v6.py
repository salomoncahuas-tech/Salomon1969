"""Mapa del área de influencia aprobada y los 117 bloques de intervención V6 — Proyecto IN Piura.

Lee datos/gis/IN_Piura_area_influencia_bloques_v6.gpkg (EPSG:32717):
  bloques_v6 (117), area_influencia_v6 (130 polígonos AI_<bloque>), ambito_distritos (15) y contexto_distritos.
Genera, en formato A3 horizontal:
  - Lámina 1: mapa general (bloques V6 + área de influencia sobre el ámbito de 15 distritos, con ubicación
    y cuadro por distrito).
  - Láminas 2-4: detalle por provincia (Ayabaca, Huancabamba, Morropón) con el código de cada bloque.
Salidas: mapas/salidas/*.png y un PDF con las 4 láminas.

Uso:  python mapas/mapa_area_influencia_bloques_v6.py
Requiere geopandas y matplotlib.
"""
from pathlib import Path

import geopandas as gpd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.lines import Line2D
from matplotlib.patches import Patch, Rectangle
from matplotlib.ticker import FuncFormatter
from shapely.geometry import box

RAIZ = Path(__file__).resolve().parents[1]
GPKG = RAIZ / "datos" / "gis" / "IN_Piura_area_influencia_bloques_v6.gpkg"
SALIDA = RAIZ / "mapas" / "salidas"

VERDE, AZUL, DORADO = "#1B4D2E", "#1B4F72", "#C9A227"
COLOR_PROV = {"AYABACA": "#E4F0DF", "HUANCABAMBA": "#DDEDEB", "MORROPON": "#F2EFE2"}
BLOQUE_FILL, BLOQUE_EDGE = "#C0392B", "#7B1E14"
AI_FILL, AI_EDGE = "#F0AE1E", "#8A6410"
NOMBRE_PROV = {"AYABACA": "Ayabaca", "HUANCABAMBA": "Huancabamba", "MORROPON": "Morropón"}
TOTAL_V6_HA = 12270.235
RETIRADOS = ["74", "75"]

plt.rcParams["font.family"] = ["Arial", "Liberation Sans", "DejaVu Sans"]

FUENTE = ("Fuente: Bloques de intervención V6 y área de influencia aprobada AI_aprobado_2 (ANIN-DIME-SESDI, 2026; "
          "se excluyen los bloques retirados 74 y 75 y sus AI). Límites distritales: INEI (capa simplificada, "
          "referencial). Elaboración: ANIN-DIME-SESDI.")


def titulo_cdl(nombre):
    """Nombre oficial en tipo título (los conectores van en minúscula)."""
    menores = {"de", "el", "la", "del"}
    pal = [p.capitalize() if i == 0 or p.lower() not in menores else p.lower()
           for i, p in enumerate(nombre.split())]
    t = " ".join(pal)
    return t.replace("Frias", "Frías").replace("Morropon", "Morropón").replace("de el Faique", "de El Faique")


def cargar():
    b = gpd.read_file(GPKG, layer="bloques_v6")
    ai = gpd.read_file(GPKG, layer="area_influencia_v6")
    amb = gpd.read_file(GPKG, layer="ambito_distritos")
    ctx = gpd.read_file(GPKG, layer="contexto_distritos")
    assert len(b) == 117, f"Se esperaban 117 bloques V6 y hay {len(b)}"
    assert not b.BLOQUE.isin(RETIRADOS).any(), "Hay bloques retirados en la capa"
    assert not ai.BLOQUE.isin(RETIRADOS).any(), "Hay AI de bloques retirados en la capa"
    assert set(ai.BLOQUE) == set(b.BLOQUE), "Hay bloques V6 sin área de influencia (o AI sin bloque)"
    assert abs(b.AREA_HA.sum() - TOTAL_V6_HA) < 0.01, f"Área total {b.AREA_HA.sum():.3f} ≠ {TOTAL_V6_HA}"
    return b, ai, amb, ctx


def ejes_utm(ax, paso=None):
    fmt = FuncFormatter(lambda v, _: f"{v:,.0f}")
    ax.xaxis.set_major_formatter(fmt)
    ax.yaxis.set_major_formatter(fmt)
    if paso:
        from matplotlib.ticker import MultipleLocator
        ax.xaxis.set_major_locator(MultipleLocator(paso))
        ax.yaxis.set_major_locator(MultipleLocator(paso))
    ax.tick_params(labelsize=7, direction="in", top=True, right=True, labeltop=True, labelright=True)
    plt.setp(ax.get_yticklabels(), rotation=90, va="center")
    ax.set_xlabel("ESTE (m)", fontsize=8)
    ax.set_ylabel("NORTE (m)", fontsize=8)
    ax.grid(True, color="#9AA5A0", linewidth=0.35, linestyle=(0, (4, 4)), zorder=0.5)
    ax.set_aspect("equal")
    for s in ax.spines.values():
        s.set_linewidth(1.2)


def norte_y_escala(ax, escala_km):
    x0, x1 = ax.get_xlim()
    y0, y1 = ax.get_ylim()
    w, h = x1 - x0, y1 - y0
    # escala gráfica alternada (exacta: el mapa está en metros)
    ex, ey = x0 + w * 0.04, y0 + h * 0.04
    seg = escala_km * 1000 / 4
    ax.add_patch(Rectangle((ex - w * 0.012, ey - h * 0.022), escala_km * 1000 + w * 0.05, h * 0.07,
                           facecolor="white", edgecolor="#555555", linewidth=0.5, alpha=0.9, zorder=6))
    for i in range(4):
        ax.add_patch(Rectangle((ex + i * seg, ey), seg, h * 0.008, facecolor="black" if i % 2 == 0 else "white",
                               edgecolor="black", linewidth=0.6, zorder=7))
    for i in range(5):
        ax.text(ex + i * seg, ey + h * 0.014, f"{i * escala_km / 4:g}", ha="center", fontsize=6.5, zorder=7)
    ax.text(ex + escala_km * 1000 + w * 0.006, ey + h * 0.004, "km", fontsize=6.5, va="center", zorder=7)
    ax.text(ex + escala_km * 500, ey - h * 0.015, "Escala gráfica", ha="center", fontsize=6.5, zorder=7)
    # norte
    nx, ny = x1 - w * 0.05, y1 - h * 0.13
    ax.add_patch(Rectangle((nx - w * 0.025, ny - h * 0.02), w * 0.05, h * 0.12, facecolor="white",
                           edgecolor="#555555", linewidth=0.5, alpha=0.9, zorder=6))
    ax.annotate("N", xy=(nx, ny + h * 0.075), xytext=(nx, ny), ha="center", va="bottom", fontsize=12,
                fontweight="bold", zorder=7,
                arrowprops=dict(facecolor="black", width=3, headwidth=10, headlength=12))


def marco(fig, titulo, subtitulo, lamina):
    """Encabezado institucional y cajetín inferior."""
    fig.add_artist(Rectangle((0.008, 0.012), 0.984, 0.976, transform=fig.transFigure, fill=False,
                             edgecolor=VERDE, linewidth=2))
    fig.add_artist(Rectangle((0.008, 0.915), 0.984, 0.073, transform=fig.transFigure, facecolor=VERDE,
                             edgecolor=VERDE))
    fig.text(0.02, 0.965, "AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN", color="white", fontsize=12,
             fontweight="bold", va="center")
    fig.text(0.02, 0.940, "DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME  ·  "
             "SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN - SESDI", color="white", fontsize=8.5, va="center")
    fig.text(0.98, 0.965, "PROYECTO IN PIURA · CUI 2669244", color=DORADO, fontsize=11, fontweight="bold",
             va="center", ha="right")
    fig.text(0.98, 0.940, "Recuperación de servicios ecosistémicos — Cuenca Alta del Río Piura", color="white",
             fontsize=8.5, va="center", ha="right")
    fig.text(0.02, 0.892, titulo, color=VERDE, fontsize=14, fontweight="bold", va="center")
    fig.text(0.02, 0.868, subtitulo, color="#333333", fontsize=9.5, va="center")
    fig.text(0.98, 0.885, f"LÁMINA {lamina}", color=VERDE, fontsize=12, fontweight="bold", ha="right",
             va="center")
    fig.text(0.98, 0.865, "Formato A3 · Octubre 2026", color="#555555", fontsize=8, ha="right", va="center")


def leyenda(ax_panel, y, extra=()):
    elementos = [
        Patch(facecolor=BLOQUE_FILL, edgecolor=BLOQUE_EDGE, label="Bloque de intervención V6 (117)"),
        Patch(facecolor=AI_FILL, edgecolor=AI_EDGE, label="Área de influencia aprobada (130 polígonos)"),
        Patch(facecolor=COLOR_PROV["AYABACA"], edgecolor="#7F8C86", label="Ámbito del proyecto – Ayabaca"),
        Patch(facecolor=COLOR_PROV["HUANCABAMBA"], edgecolor="#7F8C86", label="Ámbito del proyecto – Huancabamba"),
        Patch(facecolor=COLOR_PROV["MORROPON"], edgecolor="#7F8C86", label="Ámbito del proyecto – Morropón"),
        Line2D([], [], color=VERDE, linewidth=2.2, label="Límite del ámbito (15 distritos)"),
        Line2D([], [], color="#4D5656", linewidth=1.1, label="Límite provincial"),
        Line2D([], [], color="#95A5A6", linewidth=0.5, label="Límite distrital"),
        Patch(facecolor="#F2F2F2", edgecolor="#BFBFBF", label="Otros distritos (contexto)"),
        *extra,
    ]
    ax_panel.legend(handles=elementos, loc="upper left", bbox_to_anchor=(0.0, y), fontsize=8, frameon=False,
                    title="LEYENDA", title_fontproperties={"weight": "bold", "size": 9}, alignment="left",
                    handlelength=2.2, handleheight=1.0, labelspacing=0.4)


def dibujar_base(ax, b, ai, amb, ctx, xlim, ylim):
    ax.set_facecolor("white")  # fuera de la capa de contexto (sin datos)
    ctx.plot(ax=ax, facecolor="#F2F2F2", edgecolor="#BFBFBF", linewidth=0.3, zorder=1)
    for p, sub in amb.groupby("NOMBPROV"):
        sub.plot(ax=ax, facecolor=COLOR_PROV[p], edgecolor="#95A5A6", linewidth=0.5, zorder=2)
    prov = ctx[ctx.NOMBDEP == "PIURA"].dissolve("NOMBPROV")
    prov.boundary.plot(ax=ax, color="#4D5656", linewidth=1.1, zorder=3)
    amb.dissolve().boundary.plot(ax=ax, color=VERDE, linewidth=2.2, zorder=4)
    ai.plot(ax=ax, facecolor=AI_FILL, edgecolor=AI_EDGE, linewidth=0.45, alpha=0.9, zorder=5.5)
    b.plot(ax=ax, facecolor=BLOQUE_FILL, edgecolor=BLOQUE_EDGE, linewidth=0.6, zorder=6)
    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)


def etiquetas_distritos(ax, amb, xlim, ylim, tam=7.5):
    for _, r in amb.iterrows():
        p = r.geometry.representative_point()
        if xlim[0] < p.x < xlim[1] and ylim[0] < p.y < ylim[1]:
            ax.text(p.x, p.y, titulo_cdl(r.NOMBDIST).upper(), fontsize=tam, color="#2C3E50", ha="center",
                    va="center", style="italic", zorder=5, alpha=0.85,
                    bbox=dict(boxstyle="round,pad=0.15", facecolor="white", edgecolor="none", alpha=0.55))


def encuadre(geom, margen=0.06, aspecto=None):
    xmin, ymin, xmax, ymax = geom.total_bounds
    w, h = xmax - xmin, ymax - ymin
    xmin, xmax, ymin, ymax = xmin - w * margen, xmax + w * margen, ymin - h * margen, ymax + h * margen
    if aspecto:  # ancho / alto del eje en pantalla
        w, h = xmax - xmin, ymax - ymin
        if w / h < aspecto:
            d = (h * aspecto - w) / 2
            xmin, xmax = xmin - d, xmax + d
        else:
            d = (w / aspecto - h) / 2
            ymin, ymax = ymin - d, ymax + d
    return (xmin, xmax), (ymin, ymax)


def mapa_ubicacion(ax, ctx, amb):
    dep = ctx.dissolve("NOMBDEP")
    dep.plot(ax=ax, facecolor="#F2F2F2", edgecolor="#808B96", linewidth=0.5)
    dep.loc[["PIURA"]].plot(ax=ax, facecolor="#E8E1C8", edgecolor="#4D5656", linewidth=0.8)
    amb.dissolve().plot(ax=ax, facecolor=VERDE, edgecolor=VERDE, linewidth=0.5)
    for n, r in dep.iterrows():
        p = r.geometry.representative_point()
        if n == "PIURA":  # al oeste, fuera de la mancha del ámbito
            b0 = r.geometry.bounds
            p = type(p)(b0[0] + (b0[2] - b0[0]) * 0.22, b0[1] + (b0[3] - b0[1]) * 0.55)
        ax.text(p.x, p.y, n.title(), fontsize=6, ha="center", va="center", color="#333333",
                fontweight="bold" if n == "PIURA" else "normal")
    xmin, ymin, xmax, ymax = dep.loc[["PIURA", "TUMBES", "LAMBAYEQUE"]].total_bounds
    ax.set_xlim(xmin - 20000, xmax + 90000)
    ax.set_ylim(ymin - 15000, ymax + 15000)
    ax.text(xmin + 5000, ymin + 25000, "Océano\nPacífico", fontsize=6, style="italic", color="#1B4F72")
    ax.text(xmax + 10000, ymax - 6000, "ECUADOR", fontsize=6, color="#555555")
    ax.set_facecolor("white")
    ax.set_xticks([]); ax.set_yticks([])
    ax.set_xlabel(""); ax.set_ylabel("")
    ax.set_aspect("equal")
    ax.set_title("Ubicación en la región Piura", fontsize=8.5, fontweight="bold", color=VERDE, pad=3)
    for s in ax.spines.values():
        s.set_edgecolor(VERDE)


def resumen_distritos(b, ai):
    """N.° de bloques, ha de bloques y ha de AI por distrito (el AI se asigna al distrito de su bloque)."""
    rb = b.groupby(["NOMBPROV", "NOMBDIST"]).agg(n=("BLOQUE", "size"), ha=("AREA_HA", "sum"))
    ra = ai.groupby(["NOMBPROV", "NOMBDIST"]).agg(ai=("AREA_HA", "sum"))
    return rb.join(ra).fillna(0).reset_index().sort_values(["NOMBPROV", "NOMBDIST"])


def cuadro_distritos(ax_panel, b, ai, y_top):
    resumen = resumen_distritos(b, ai)
    filas = [[NOMBRE_PROV[r.NOMBPROV], titulo_cdl(r.NOMBDIST), f"{r.n}", f"{r.ha:,.2f}", f"{r.ai:,.2f}"]
             for r in resumen.itertuples()]
    filas.append(["TOTAL", "15 distritos", f"{resumen.n.sum()}", f"{resumen.ha.sum():,.3f}",
                  f"{resumen.ai.sum():,.3f}"])
    alto = 0.0195 * (len(filas) + 1)
    t = ax_panel.table(cellText=filas, colLabels=["Provincia", "Distrito", "Bloques", "Bloques (ha)", "AI (ha)"],
                       colWidths=[0.22, 0.37, 0.11, 0.16, 0.14], cellLoc="left",
                       bbox=[0.0, y_top - alto, 1.0, alto])
    t.auto_set_font_size(False)
    t.set_fontsize(7)
    for (i, j), c in t.get_celld().items():
        c.set_edgecolor("#B0B7B3")
        c.set_linewidth(0.4)
        if j >= 2:
            c.set_text_props(ha="right")
        if i == 0:
            c.set_facecolor(VERDE)
            c.set_text_props(color="white", fontweight="bold", ha="center")
        elif i == len(filas):
            c.set_facecolor("#D5E8D4")
            c.set_text_props(fontweight="bold")
        elif i % 2 == 0:
            c.set_facecolor("#F2F7F2")
    ax_panel.text(0.0, y_top + 0.012, "BLOQUES V6 Y ÁREA DE INFLUENCIA POR DISTRITO", fontsize=9, fontweight="bold", color=VERDE,
                  transform=ax_panel.transAxes)
    return y_top - alto


def pie(fig):
    fig.text(0.02, 0.034, "Sistema de coordenadas: UTM WGS84 Zona 17S (EPSG:32717)", fontsize=7.5,
             color="#333333", fontweight="bold")
    fig.text(0.02, 0.019, FUENTE, fontsize=6.5, color="#555555")


def lamina_general(b, ai, amb, ctx):
    fig = plt.figure(figsize=(16.54, 11.69))
    marco(fig, "MAPA DEL ÁREA DE INFLUENCIA Y BLOQUES DE INTERVENCIÓN V6",
          f"117 bloques de intervención ({TOTAL_V6_HA:,.3f} ha) y su área de influencia ({ai.AREA_HA.sum():,.3f} ha) "
          "en 15 distritos de Ayabaca, Huancabamba y Morropón — Región Piura", 1)
    ax = fig.add_axes([0.04, 0.075, 0.62, 0.755])
    xlim, ylim = encuadre(amb, 0.03, aspecto=0.62 * 16.54 / (0.755 * 11.69))
    dibujar_base(ax, b, ai, amb, ctx, xlim, ylim)
    etiquetas_distritos(ax, amb, xlim, ylim, tam=7.5)
    ejes_utm(ax, paso=20000)
    norte_y_escala(ax, 20)

    panel = fig.add_axes([0.69, 0.065, 0.29, 0.77])
    panel.axis("off")
    ax_ub = fig.add_axes([0.70, 0.60, 0.27, 0.23])
    mapa_ubicacion(ax_ub, ctx, amb)
    y = cuadro_distritos(panel, b, ai, 0.635)
    leyenda(panel, y - 0.01)
    area_ai = ai.AREA_HA.sum()
    panel.text(0.0, 0.012,
               f"Bloques V6: {TOTAL_V6_HA:,.3f} ha (117)\n"
               f"Área de influencia: {area_ai:,.3f} ha ({len(ai)} polígonos)\n"
               f"Bloques + área de influencia: {TOTAL_V6_HA + area_ai:,.3f} ha\n"
               f"Ámbito (15 distritos, referencial): {amb.AREA_HA.sum():,.0f} ha",
               fontsize=7.5, color="#333333", transform=panel.transAxes, va="bottom",
               bbox=dict(boxstyle="round,pad=0.4", facecolor="#F7F3E3", edgecolor=DORADO, linewidth=0.8))
    pie(fig)
    return fig


def lamina_provincia(b, ai, amb, ctx, prov, lamina):
    bp, ap, aip = b[b.NOMBPROV == prov], ai[ai.NOMBPROV == prov], amb[amb.NOMBPROV == prov]
    fig = plt.figure(figsize=(16.54, 11.69))
    marco(fig, f"ÁREA DE INFLUENCIA Y BLOQUES V6 — PROVINCIA DE {NOMBRE_PROV[prov].upper()}",
          f"{len(bp)} bloques de intervención ({bp.AREA_HA.sum():,.3f} ha) y su área de influencia "
          f"({ap.AREA_HA.sum():,.3f} ha, {len(ap)} polígonos) en {aip.NOMBDIST.nunique()} distrito(s)", lamina)
    ax = fig.add_axes([0.04, 0.075, 0.68, 0.755])
    # se reserva una franja inferior para la escala gráfica, así no tapa bloques
    xmin, ymin, xmax, ymax = aip.total_bounds
    caja = gpd.GeoSeries([box(xmin, ymin - (ymax - ymin) * 0.10, xmax, ymax)], crs=aip.crs)
    xlim, ylim = encuadre(caja, 0.04, aspecto=0.68 * 16.54 / (0.755 * 11.69))
    dibujar_base(ax, b, ai, amb, ctx, xlim, ylim)
    etiquetas_distritos(ax, amb, xlim, ylim, tam=8.5)
    # etiquetas de bloque: adjustText las separa entre sí y de los polígonos, con línea guía al bloque
    from adjustText import adjust_text
    w = xlim[1] - xlim[0]
    pts = bp.geometry.representative_point()
    textos = [ax.text(pt.x, pt.y, cod, fontsize=6.6, fontweight="bold", color="#5B1A12", ha="center",
                      va="center", zorder=8,
                      bbox=dict(boxstyle="round,pad=0.12", facecolor="white", edgecolor=BLOQUE_EDGE,
                                linewidth=0.4, alpha=0.9))
              for cod, pt in zip(bp.BLOQUE, pts)]
    adjust_text(textos, x=list(pts.x), y=list(pts.y), ax=ax, expand=(1.25, 1.4), force_text=(0.4, 0.6),
                force_static=(0.3, 0.4), ensure_inside_axes=True, max_move=None,
                arrowprops=dict(arrowstyle="-", color=BLOQUE_EDGE, linewidth=0.45, zorder=7))
    paso = 10000 if w < 120000 else 20000
    ejes_utm(ax, paso=paso)
    norte_y_escala(ax, next(k for k in (20, 10, 5, 2) if k * 1000 <= w / 7))

    panel = fig.add_axes([0.745, 0.065, 0.235, 0.77])
    panel.axis("off")
    ax_ub = fig.add_axes([0.75, 0.645, 0.225, 0.185])
    mapa_ubicacion(ax_ub, ctx, amb)
    ax_ub.set_title("Ubicación de la provincia", fontsize=8.5, fontweight="bold", color=VERDE, pad=3)
    xmin, ymin, xmax, ymax = aip.total_bounds
    ax_ub.add_patch(Rectangle((xmin, ymin), xmax - xmin, ymax - ymin, fill=False, edgecolor=BLOQUE_FILL,
                              linewidth=1.3, zorder=5))
    # lista de bloques de la provincia
    filas = bp.sort_values(["NOMBDIST", "BLOQUE"])
    lineas = []
    ai_dist = ap.groupby("NOMBDIST").AREA_HA.sum()
    for dist, sub in filas.groupby("NOMBDIST"):
        lineas.append((f"{titulo_cdl(dist)}: {len(sub)} bloques · {sub.AREA_HA.sum():,.2f} ha · "
                       f"AI {ai_dist.get(dist, 0):,.2f} ha", True))
        cods = list(sub.BLOQUE)
        for k in range(0, len(cods), 7):
            lineas.append(("   " + ", ".join(cods[k:k + 7]), False))
    y = 0.715
    panel.text(0.0, y, "BLOQUES V6 Y ÁREA DE INFLUENCIA (AI)", fontsize=9, fontweight="bold", color=VERDE,
               transform=panel.transAxes)
    y -= 0.025
    paso_y = min(0.019, 0.36 / max(len(lineas), 1))
    for txt, neg in lineas:
        panel.text(0.0, y, txt, fontsize=7.2 if neg else 6.8, fontweight="bold" if neg else "normal",
                   color="#1B4D2E" if neg else "#333333", transform=panel.transAxes)
        y -= paso_y
    leyenda(panel, y - 0.005)
    pie(fig)
    return fig


def main():
    b, ai, amb, ctx = cargar()
    SALIDA.mkdir(parents=True, exist_ok=True)
    figs = [("01_mapa_general", lamina_general(b, ai, amb, ctx))]
    for i, p in enumerate(["AYABACA", "HUANCABAMBA", "MORROPON"], start=2):
        figs.append((f"0{i}_{p.lower()}", lamina_provincia(b, ai, amb, ctx, p, i)))
    pdf = SALIDA / "IN_Piura_Mapa_Area_Influencia_Bloques_V6.pdf"
    with PdfPages(pdf) as pp:
        for nombre, fig in figs:
            fig.savefig(SALIDA / f"IN_Piura_{nombre}.png", dpi=200)
            pp.savefig(fig)
            plt.close(fig)
    print(f"{len(figs)} láminas → {SALIDA}")


if __name__ == "__main__":
    main()
