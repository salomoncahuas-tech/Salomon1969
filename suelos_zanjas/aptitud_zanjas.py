"""Aptitud de áreas para zanjas de infiltración por bloque — Proyecto IN Piura (CUI 2669244).

Cruza la geometría de los 117 bloques V6 con el mapa de suelos del Estudio de Suelos E5
(informe final, GEOSIG: UNIDADES_CARTOGRAFICAS_SUELO, 423 delineaciones) y con una capa
de pendientes, y devuelve por bloque la superficie en cada clase de pendiente de 5 %
(0–10, 10–15, …, 35–40, > 40 %) y la parte de esa superficie que está en suelos aptos
para la zanja de infiltración.

Fuentes de pendiente (en orden de preferencia):
  1. Vector de pendientes del proyecto («Pendientes vector», campo DN 1–7) si se pasa con
     --pendientes y su leyenda con --leyenda (JSON {"1": [lim_inf, lim_sup, "grados"|"%"], ...}).
  2. Clase de pendiente dominante de cada delineación del E5 (PEND_SIMB, D.S. N.° 005-2022-
     MIDAGRI, Anexo V, sobre el MDE de 12.5 m del proyecto).
Cuando una clase de pendiente cruza un límite de 10, 15, …, 40 %, su superficie se reparte
en proporción lineal a la amplitud de la clase en su propia unidad (grados o porcentaje).

Aptitud del suelo (E5, Cuadro 61 y Anexo 7, matriz suelo–intervención): una zanja de 40 cm
de profundidad debe vaciarse en 72 h con la velocidad de diseño (mitad de la básica, Guía 5
de Forest Trends 2024). Solo Canchaque y El Ingenio (6.625 cm/h de diseño, 6 h) lo cumplen;
Morropón (229 h), Cruz de Piedra (242 h), Ramón Castilla (242 h), Chillique (96 h) y
Huarmaca (242 h) no. La parte apta de cada delineación es la proporción de esos componentes
en su unidad cartográfica (COMP_SUELO, con inclusiones).

Uso:
    python suelos_zanjas/aptitud_zanjas.py --salida suelos_zanjas/salidas
    python suelos_zanjas/aptitud_zanjas.py --pendientes "Pendientes vector.shp" \
        --leyenda leyenda_pendientes.json --salida suelos_zanjas/salidas
"""
import argparse
import json
import math
import re
from collections import defaultdict
from pathlib import Path


import geopandas as gpd  # noqa: E402

RAIZ = Path(__file__).resolve().parents[1]
BLOQUES_GPKG = RAIZ / "datos" / "gis" / "IN_Piura_area_influencia_bloques_v6.gpkg"
SUELOS_SHP = RAIZ / "datos" / "gis" / "suelos_e5" / "UNIDADES_CARTOGRAFICAS_SUELO.shp"

# Clases de pendiente del motor (porcentaje)
BINS = [(0, 10), (10, 15), (15, 20), (20, 25), (25, 30), (30, 35), (35, 40), (40, None)]
# Clases de pendiente D.S. N.° 005-2022-MIDAGRI, Anexo V (porcentaje)
MIDAGRI = {"A": (0, 2), "B": (2, 4), "C": (4, 8), "D": (8, 15), "E": (15, 25),
           "F": (25, 50), "G": (50, 75), "H": (75, None)}
# Unidades de suelo que vacían una zanja de 40 cm en ≤ 72 h (E5, Cuadro 61)
SUELO_APTO = {"Canchaque", "El Ingenio"}
# Velocidad de diseño (cm/h, mitad de la básica) y vaciado de 40 cm (h) — E5, Tasa_basica_por_unidad_de_suelo.xlsx
INFILTRACION = {
    "Morropón": (0.175, 229), "Cruz de Piedra": (0.165, 242), "Ramón Castilla": (0.165, 242),
    "Chillique": (0.415, 96), "Canchaque": (6.625, 6), "Huarmaca": (0.165, 242), "El Ingenio": (6.625, 6),
}


def pct_a_grados(p):
    return math.degrees(math.atan(p / 100.0))


def reparto(lo, hi, unidad="%"):
    """Reparte una clase [lo, hi) en las clases BINS (en %), lineal en la unidad de la clase."""
    if unidad.startswith("g"):
        conv = pct_a_grados
    else:
        conv = (lambda p: p)
    out = [0.0] * len(BINS)
    if hi is None:  # clase abierta (p. ej. > 75 %): debe empezar en 40 % o más
        lo_pct = math.tan(math.radians(lo)) * 100 if unidad.startswith("g") else lo
        if lo_pct < BINS[-1][0] - 1e-9:
            raise ValueError(f"Clase abierta que empieza en {lo} {unidad} (< 40 %): defina su límite superior")
        out[-1] = 1.0
        return out
    w = hi - lo
    for i, (a, b) in enumerate(BINS):
        a_u = conv(a)
        b_u = math.inf if b is None else conv(b)
        s = max(0.0, min(hi, b_u) - max(lo, a_u))
        out[i] = s / w if w > 0 else 0.0
    return out


def componentes(comp):
    """'Chillique 50 %; Cruz de Piedra 38 %; Morropón 12 %' → {'Chillique': .5, ...}"""
    out = {}
    for parte in str(comp or "").split(";"):
        m = re.match(r"\s*(.+?)\s+(\d+(?:[.,]\d+)?)\s*%", parte)
        if m:
            out[m.group(1).strip()] = float(m.group(2).replace(",", ".")) / 100.0
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--bloques", default=str(BLOQUES_GPKG))
    ap.add_argument("--suelos", default=str(SUELOS_SHP))
    ap.add_argument("--pendientes", default=None, help="Vector de pendientes con campo DN")
    ap.add_argument("--leyenda", default=None, help="JSON con los rangos de cada DN")
    ap.add_argument("--salida", default=str(RAIZ / "suelos_zanjas" / "salidas"))
    a = ap.parse_args()
    Path(a.salida).mkdir(parents=True, exist_ok=True)

    blq = gpd.read_file(a.bloques, layer="bloques_v6")[["BLOQUE", "geometry"]]
    blq["BLOQUE"] = blq["BLOQUE"].astype(str).str.strip()
    blq = blq.dissolve("BLOQUE").reset_index()
    blq["ha_blq"] = blq.area / 1e4
    sue = gpd.read_file(a.suelos, encoding="utf-8")
    sue = sue[sue["AMBITO"] == "Bloque"][["COD_POL", "UNID_CARTO", "COMP_SUELO", "PEND_SIMB", "geometry"]]
    sue = sue.to_crs(blq.crs)

    x = gpd.overlay(blq[["BLOQUE", "geometry"]], sue, how="intersection", keep_geom_type=True)
    if a.pendientes:
        leyenda = {int(k): v for k, v in json.load(open(a.leyenda, encoding="utf-8")).items() if not k.startswith("_")}
        ruta = a.pendientes
        if ruta.lower().endswith(".zip"):
            ruta = f"zip://{ruta}!Pendientes vector/Pendientes vector.shp"
        pen = gpd.read_file(ruta)[["DN", "geometry"]].to_crs(blq.crs)
        x = gpd.overlay(x, pen, how="intersection", keep_geom_type=True)
        rep = {dn: reparto(v[0], v[1], v[2] if len(v) > 2 else "%") for dn, v in leyenda.items()}
        x["rep"] = x["DN"].map(rep)
        fuente = ("Vector de pendientes del proyecto («Pendientes vector», 57,255 polígonos, DN 1–7 en grados: "
                  "0–5°, 5–10°, 10–15°, 15–20°, 20–25°, 25–27°, > 27°)")
    else:
        rep = {k: reparto(lo, hi, "%") for k, (lo, hi) in MIDAGRI.items()}
        x["rep"] = x["PEND_SIMB"].map(rep)
        fuente = "Pendiente dominante de la delineación del E5 (PEND_SIMB, D.S. N.° 005-2022-MIDAGRI)"
    x["ha"] = x.area / 1e4

    res = {}
    acc = defaultdict(lambda: {"a": [0.0] * len(BINS), "s": [0.0] * len(BINS), "us": defaultdict(float), "uc": defaultdict(float), "pd": defaultdict(float)})
    for _, r in x.iterrows():
        comp = componentes(r["COMP_SUELO"])
        tot = sum(comp.values()) or 1.0
        apto = sum(v for k, v in comp.items() if k in SUELO_APTO) / tot
        b = acc[r["BLOQUE"]]
        for i, f in enumerate(r["rep"]):
            b["a"][i] += r["ha"] * f
            b["s"][i] += r["ha"] * f * apto
        for k, v in comp.items():
            b["us"][k] += r["ha"] * v / tot
        b["uc"][r["UNID_CARTO"]] += r["ha"]
        b["pd"][r["PEND_SIMB"]] += r["ha"]
    for _, rb in blq.iterrows():
        k = rb["BLOQUE"]
        b = acc.get(k)
        ha = rb["ha_blq"]
        if b is None:
            res[k] = {"ha": round(ha, 3), "cov": 0.0, "a": None, "s": None, "us": {}}
            continue
        cov = sum(b["a"])
        res[k] = {
            "ha": round(ha, 3),
            "cov": round(cov / ha, 4) if ha else 0.0,
            # fracciones de la superficie cartografiada del bloque
            "a": [round(v / cov, 5) for v in b["a"]],
            "s": [round(v / cov, 5) for v in b["s"]],
            "us": {u: round(v, 2) for u, v in sorted(b["us"].items(), key=lambda t: -t[1])},
            "uc": {u: round(v, 2) for u, v in sorted(b["uc"].items(), key=lambda t: -t[1])},
            "pd": {u: round(v, 2) for u, v in sorted(b["pd"].items())},
        }
    meta = {
        "fuente_pendiente": fuente,
        "fuente_suelos": "E5 informe final (GEOSIG), UNIDADES_CARTOGRAFICAS_SUELO, 423 delineaciones (388 en bloques)",
        "bloques": "Bloques V6 (IN_Piura_area_influencia_bloques_v6.gpkg, capa bloques_v6)",
        "bins": [[lo, hi] for lo, hi in BINS],
        "suelo_apto": sorted(SUELO_APTO),
        "infiltracion": INFILTRACION,
        "n_bloques": len(res),
        "sin_mapa": [k for k, v in res.items() if not v["cov"]],
    }
    # Atributos por bloque de la geodatabase del E5 (capa BLOQUES_INTERVENCION, geometría V7, 118 bloques)
    e5b = {}
    dbf = Path(a.suelos).with_name("BLOQUES_INTERVENCION.dbf")
    if dbf.exists():
        import pyogrio
        t = pyogrio.read_dataframe(dbf, read_geometry=False, encoding="utf-8")
        for _, r in t.iterrows():
            e5b[str(r["BLOQUE"]).strip()] = {k: (None if str(r[k]) == "nan" else str(r[k])) for k in
                                             ("COMP_SUELO", "CUM_COMP", "US_DOM", "UC_DOM", "CUM_DOM", "PEND_DOM", "SUBGR_DOM")}
    meta["e5_bloques_sin_v6"] = sorted(set(e5b) - set(res))
    out = {"meta": meta, "blk": res, "e5b": {k: v for k, v in e5b.items() if k in res}}
    p = Path(a.salida) / "aptitud_zanjas_bloques.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    print(p, len(res), "bloques ·", fuente)


if __name__ == "__main__":
    main()
