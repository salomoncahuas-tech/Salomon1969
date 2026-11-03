"""Aptitud de áreas para zanjas de infiltración y terrazas por bloque — Proyecto IN Piura (CUI 2669244).

Cruza la geometría de los 117 bloques V6 con:
  * el mapa de suelos del Estudio de Suelos E5 (informe final, GEOSIG: UNIDADES_CARTOGRAFICAS_SUELO,
    423 delineaciones);
  * el vector de pendientes del proyecto («Pendientes vector», DN 1–7 en grados);
  * el MSAVI vectorizado V6 (DN 1–5): R1 = DN 1–2 (plantación plena), R2 = DN 3 (enriquecimiento),
    C = DN 4–5 (protección);
y devuelve, por bloque y por clase R1/R2, la superficie en cada clase de pendiente del motor
(0–10, 10–15, 15–20, 20–25, 25–30, 30–35, 35–40, 40–70 y > 70 %) y la parte en suelos aptos para
la zanja de infiltración.

Pendiente. Cuando una clase de la leyenda cruza un límite de 10, 15, …, 40 o 70 %, su superficie se
reparte en proporción lineal a la amplitud de la clase en su propia unidad (grados). La clase abierta
del vector (> 27°, más de 51 %) no tiene límite superior: su superficie va a 40–70 %, salvo en las
delineaciones del E5 con pendiente dominante H (> 75 %), donde va a > 70 %. Sin vector de pendientes
se usa la clase dominante de cada delineación del E5 (PEND_SIMB, D.S. N.° 005-2022-MIDAGRI).

Suelo (E5, Cuadro 61 y Anexo 7): una zanja de 40 cm debe vaciarse en ≤ 72 h con la velocidad de diseño
(mitad de la básica; Guía 5 de Forest Trends 2024). Solo Canchaque y El Ingenio lo cumplen. La parte
apta de cada delineación es la proporción de esos componentes en su unidad cartográfica (COMP_SUELO).

Uso:
    python suelos_zanjas/aptitud_zanjas.py --pendientes datos/gis/pendientes/Pendientes_vector.zip \
        --leyenda suelos_zanjas/leyenda_pendientes.json --msavi datos/gis/msavi/MSAVI_vectorizado.zip \
        --salida suelos_zanjas/salidas
"""
import argparse
import json
import math
import re
from collections import defaultdict
from pathlib import Path

import geopandas as gpd

RAIZ = Path(__file__).resolve().parents[1]
BLOQUES_GPKG = RAIZ / "datos" / "gis" / "IN_Piura_area_influencia_bloques_v6.gpkg"
SUELOS_SHP = RAIZ / "datos" / "gis" / "suelos_e5" / "UNIDADES_CARTOGRAFICAS_SUELO.shp"

# Clases de pendiente del motor (porcentaje). 10–40 %: zanjas (cartilla MIDAGRI); 30–70 %: terrazas (E5).
BINS = [(0, 10), (10, 15), (15, 20), (20, 25), (25, 30), (30, 35), (35, 40), (40, 70), (70, None)]
I_4070 = 7
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
MSAVI_R = {1: "r1", 2: "r1", 3: "r2", 4: "c", 5: "c"}


def pct_a_grados(p):
    return math.degrees(math.atan(p / 100.0))


def reparto(lo, hi, unidad="%"):
    """Reparte una clase [lo, hi) en las clases BINS (en %), lineal en la unidad de la clase.
    Una clase abierta (hi = None) va entera a la clase del motor que contiene su límite inferior."""
    grados = unidad.startswith("g")
    conv = pct_a_grados if grados else (lambda p: p)
    out = [0.0] * len(BINS)
    if hi is None:
        lo_pct = math.tan(math.radians(lo)) * 100 if grados else lo
        for i, (a, b) in enumerate(BINS):
            if lo_pct >= a and (b is None or lo_pct < b):
                out[i] = 1.0
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


def leer(ruta, capa_zip):
    if str(ruta).lower().endswith(".zip"):
        return gpd.read_file(f"zip://{ruta}!{capa_zip}")
    return gpd.read_file(ruta)


def nuevo():
    return {"a": [0.0] * len(BINS), "s": [0.0] * len(BINS), "ha": 0.0}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--bloques", default=str(BLOQUES_GPKG))
    ap.add_argument("--suelos", default=str(SUELOS_SHP))
    ap.add_argument("--pendientes", default=None, help="Vector de pendientes con campo DN (.shp o .zip)")
    ap.add_argument("--leyenda", default=None, help="JSON con los rangos de cada DN")
    ap.add_argument("--msavi", default=None, help="MSAVI vectorizado con campo DN (.shp o .zip)")
    ap.add_argument("--salida", default=str(RAIZ / "suelos_zanjas" / "salidas"))
    a = ap.parse_args()
    Path(a.salida).mkdir(parents=True, exist_ok=True)

    blq = gpd.read_file(a.bloques, layer="bloques_v6")[["BLOQUE", "geometry"]]
    blq["BLOQUE"] = blq["BLOQUE"].astype(str).str.strip()
    blq = blq.dissolve("BLOQUE").reset_index()
    blq["ha_blq"] = blq.area / 1e4
    sue = gpd.read_file(a.suelos, encoding="utf-8")
    sue = sue[sue["AMBITO"] == "Bloque"][["COD_POL", "UNID_CARTO", "COMP_SUELO", "PEND_SIMB", "geometry"]].to_crs(blq.crs)

    x = gpd.overlay(blq[["BLOQUE", "geometry"]], sue, how="intersection", keep_geom_type=True)
    abierta = set()
    if a.pendientes:
        leyenda = {int(k): v for k, v in json.load(open(a.leyenda, encoding="utf-8")).items() if not k.startswith("_")}
        pen = leer(a.pendientes, "Pendientes vector/Pendientes vector.shp")[["DN", "geometry"]].to_crs(blq.crs)
        x = gpd.overlay(x, pen, how="intersection", keep_geom_type=True)
        rep = {dn: reparto(v[0], v[1], v[2] if len(v) > 2 else "%") for dn, v in leyenda.items()}
        abierta = {dn for dn, v in leyenda.items() if v[1] is None}
        x["clase"] = x["DN"]
        fuente = ("Vector de pendientes del proyecto («Pendientes vector», 57,255 polígonos, DN 1–7 en grados: "
                  "0–5°, 5–10°, 10–15°, 15–20°, 20–25°, 25–27°, > 27°)")
    else:
        rep = {k: reparto(lo, hi, "%") for k, (lo, hi) in MIDAGRI.items()}
        x["clase"] = x["PEND_SIMB"]
        fuente = "Pendiente dominante de la delineación del E5 (PEND_SIMB, D.S. N.° 005-2022-MIDAGRI)"
    if a.msavi:
        ms = leer(a.msavi, "MSAVI vectorizado/MSAVI vectorizado.shp")[["DN", "geometry"]].rename(columns={"DN": "MS"}).to_crs(blq.crs)
        x = gpd.overlay(x, ms, how="intersection", keep_geom_type=True)
        x["R"] = x["MS"].map(MSAVI_R)
        fuente_ms = "MSAVI vectorizado V6 (38,988 polígonos; R1 = DN 1–2, R2 = DN 3, C = DN 4–5), cruzado polígono a polígono"
    else:
        x["R"] = None
        fuente_ms = None
    x["ha"] = x.area / 1e4

    acc = defaultdict(lambda: {"blk": nuevo(), "r1": nuevo(), "r2": nuevo(), "c": nuevo(),
                               "us": defaultdict(float), "uc": defaultdict(float), "pd": defaultdict(float)})
    for _, r in x.iterrows():
        comp = componentes(r["COMP_SUELO"])
        tot = sum(comp.values()) or 1.0
        apto = sum(v for k, v in comp.items() if k in SUELO_APTO) / tot
        f = list(rep[r["clase"]])
        if r["clase"] in abierta and r["PEND_SIMB"] == "H":   # > 27° dentro de una delineación > 75 %
            f = [0.0] * len(BINS)
            f[-1] = 1.0
        b = acc[r["BLOQUE"]]
        dest = [b["blk"]] + ([b[r["R"]]] if r["R"] in ("r1", "r2", "c") else [])
        for d in dest:
            d["ha"] += r["ha"]
            for i, fi in enumerate(f):
                d["a"][i] += r["ha"] * fi
                d["s"][i] += r["ha"] * fi * apto
        for k, v in comp.items():
            b["us"][k] += r["ha"] * v / tot
        b["uc"][r["UNID_CARTO"]] += r["ha"]
        b["pd"][r["PEND_SIMB"]] += r["ha"]

    def frac(d):
        h = d["ha"]
        if not h:
            return None
        return {"ha": round(h, 3), "a": [round(v / h, 5) for v in d["a"]], "s": [round(v / h, 5) for v in d["s"]]}

    res = {}
    for _, rb in blq.iterrows():
        k = rb["BLOQUE"]
        ha = rb["ha_blq"]
        b = acc.get(k)
        if b is None:
            res[k] = {"ha": round(ha, 3), "cov": 0.0, "a": None, "s": None, "us": {}}
            continue
        fb = frac(b["blk"])
        res[k] = {
            "ha": round(ha, 3),
            "cov": round(b["blk"]["ha"] / ha, 4) if ha else 0.0,
            "a": fb["a"], "s": fb["s"],
            "r1": frac(b["r1"]), "r2": frac(b["r2"]), "c": frac(b["c"]),
            "us": {u: round(v, 2) for u, v in sorted(b["us"].items(), key=lambda t: -t[1])},
            "uc": {u: round(v, 2) for u, v in sorted(b["uc"].items(), key=lambda t: -t[1])},
            "pd": {u: round(v, 2) for u, v in sorted(b["pd"].items())},
        }
    meta = {
        "fuente_pendiente": fuente,
        "fuente_msavi": fuente_ms,
        "fuente_suelos": "E5 informe final (GEOSIG), UNIDADES_CARTOGRAFICAS_SUELO, 423 delineaciones (388 en bloques)",
        "bloques": "Bloques V6 (IN_Piura_area_influencia_bloques_v6.gpkg, capa bloques_v6)",
        "bins": [[lo, hi] for lo, hi in BINS],
        "suelo_apto": sorted(SUELO_APTO),
        "infiltracion": INFILTRACION,
        "n_bloques": len(res),
        "sin_mapa": [k for k, v in res.items() if not v["cov"]],
        "clase_abierta": "La clase > 27° va a 40–70 %, salvo en delineaciones del E5 con pendiente dominante H (> 75 %)",
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
    print(p, len(res), "bloques ·", fuente, "·", fuente_ms)


if __name__ == "__main__":
    main()
