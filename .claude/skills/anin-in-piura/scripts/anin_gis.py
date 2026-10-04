"""Utilidades GIS del Proyecto IN Piura (UTM WGS84 Zona 17S, EPSG:32717).

    from anin_gis import puntos_utm, mapa_utm, guardar_gpkg
    gdf = puntos_utm(df, "ESTE_UTM", "NORTE_UTM")     # GeoDataFrame en EPSG:32717 con control de rangos
    mapa_utm(gdf, "mapa.png", "Título", columna="provincia", colores={"AYABACA": "#2E7D4F"})
    guardar_gpkg(gdf, "salida.gpkg", "capa")

Requiere geopandas, shapely, pyproj y matplotlib (pip install geopandas matplotlib).
Las áreas y distancias se calculan siempre sobre estas geometrías proyectadas, nunca en grados.
"""
import numpy as np
import pandas as pd
import geopandas as gpd
from pyproj import Transformer

from anin_utm import validar_utm

CRS_PROYECTO = "EPSG:32717"
_A_LL = Transformer.from_crs(CRS_PROYECTO, "EPSG:4326", always_xy=True)


def puntos_utm(df, col_este, col_norte):
    """DataFrame con coordenadas UTM 17S -> GeoDataFrame EPSG:32717.

    Agrega coord_ok / coord_msg (rangos del proyecto), lat y lon (grados decimales). Los registros
    con coordenadas vacías o no numéricas se conservan con geometría vacía y coord_ok = False, para que
    el control los reporte en vez de perderlos en silencio."""
    d = df.copy()
    d[col_este] = pd.to_numeric(d[col_este], errors="coerce")
    d[col_norte] = pd.to_numeric(d[col_norte], errors="coerce")
    res = [validar_utm(e, n) if pd.notna(e) and pd.notna(n) else (False, "coordenada vacía o no numérica")
           for e, n in zip(d[col_este], d[col_norte])]
    d["coord_ok"] = [r[0] for r in res]
    d["coord_msg"] = [r[1] for r in res]
    lon, lat = _A_LL.transform(d[col_este].fillna(0).values, d[col_norte].fillna(0).values)
    d["lat"], d["lon"] = np.where(d["coord_ok"], lat, np.nan), np.where(d["coord_ok"], lon, np.nan)
    geom = gpd.points_from_xy(d[col_este], d[col_norte])
    return gpd.GeoDataFrame(d, geometry=geom, crs=CRS_PROYECTO)


def guardar_gpkg(gdf, ruta, capa):
    """Escribe una capa en GeoPackage (un solo archivo, con CRS incluido)."""
    gdf.to_file(ruta, layer=capa, driver="GPKG")


def mapa_utm(gdf, ruta_png, titulo, columna=None, colores=None, marcadores=None, fuente="",
             escala_km=10, tam=(8.2, 7.8), dpi=200, titulo_leyenda=None):
    """Mapa esquemático de puntos con ticks en UTM, norte, escala gráfica, leyenda y fuente.

    columna:    campo que define las categorías (color); colores: {categoría: hex}; marcadores: {categoría: marcador}.
    La leyenda va debajo del mapa para no tapar puntos."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter

    g = gdf[gdf.geometry.notna() & ~gdf.geometry.is_empty]
    fig, ax = plt.subplots(figsize=tam, dpi=dpi)
    cats = [None] if columna is None else list(colores) if colores else sorted(g[columna].dropna().unique())
    for c in cats:
        sub = g if c is None else g[g[columna] == c]
        if sub.empty:
            continue
        ax.scatter(sub.geometry.x, sub.geometry.y, s=26, c=(colores or {}).get(c, "#1B4D2E"),
                   marker=(marcadores or {}).get(c, "o"), edgecolors="white", linewidths=0.4,
                   label=f"{c} ({len(sub)})" if c is not None else None, zorder=3)
    xmin, ymin, xmax, ymax = g.total_bounds
    px, py = (xmax - xmin) * 0.06 + 1500, (ymax - ymin) * 0.06 + 1500
    ax.set_xlim(xmin - px, xmax + px); ax.set_ylim(ymin - py, ymax + py)
    ax.set_aspect("equal")
    fmt = FuncFormatter(lambda v, _: f"{v:,.0f}")
    ax.xaxis.set_major_formatter(fmt); ax.yaxis.set_major_formatter(fmt)
    ax.tick_params(labelsize=7)
    plt.setp(ax.get_yticklabels(), rotation=90, va="center")
    ax.set_xlabel("ESTE (m) · UTM WGS84 Zona 17S", fontsize=8); ax.set_ylabel("NORTE (m)", fontsize=8)
    ax.grid(True, color="#DDDDDD", linewidth=0.5, zorder=0)
    ax.set_title(titulo, fontsize=11, fontweight="bold", color="#1B4D2E", loc="left")
    # escala gráfica (exacta: el mapa está en metros)
    x0, y0 = ax.get_xlim()[0] + (ax.get_xlim()[1] - ax.get_xlim()[0]) * 0.04, ax.get_ylim()[0] + (ax.get_ylim()[1] - ax.get_ylim()[0]) * 0.05
    ax.plot([x0, x0 + escala_km * 1000], [y0, y0], color="black", linewidth=3, solid_capstyle="butt", zorder=5)
    ax.text(x0 + escala_km * 500, y0 + (ax.get_ylim()[1] - ax.get_ylim()[0]) * 0.012, f"{escala_km} km", ha="center", fontsize=8, zorder=5)
    # norte
    xr, yr = ax.get_xlim()[1] - (ax.get_xlim()[1] - ax.get_xlim()[0]) * 0.05, ax.get_ylim()[1] - (ax.get_ylim()[1] - ax.get_ylim()[0]) * 0.17
    ax.annotate("N", xy=(xr, yr + (ax.get_ylim()[1] - ax.get_ylim()[0]) * 0.09), xytext=(xr, yr), ha="center", va="bottom",
                fontsize=11, fontweight="bold", arrowprops=dict(facecolor="black", width=2.5, headwidth=9, headlength=11), zorder=5)
    if columna is not None:
        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.09), ncol=1, fontsize=8, frameon=False,
                  title=titulo_leyenda, title_fontsize=8)
    fig.text(0.01, 0.005, fuente, fontsize=6.5, color="#555555", ha="left", va="bottom")
    fig.tight_layout(rect=(0, 0.02, 1, 1))
    fig.savefig(ruta_png, bbox_inches="tight")
    plt.close(fig)
    return ruta_png
