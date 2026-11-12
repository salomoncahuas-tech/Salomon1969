"""Lectura de cifras desde las fuentes del proyecto (sin digitar valores)."""
from pathlib import Path

import geopandas as gpd
import openpyxl

PROVS = ("Morropón", "Huancabamba", "Ayabaca")


def _rows(path, sheet):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    return [r for r in wb[sheet].iter_rows(values_only=True)]


def _tabla(path, sheet, primera="Provincia"):
    rows = _rows(path, sheet)
    i = next(k for k, r in enumerate(rows) if r and r[0] == primera)
    hdr = [c for c in rows[i]]
    out = []
    for r in rows[i + 1:]:
        if r and r[0] not in (None, ""):
            out.append(dict(zip(hdr, r)))
    return hdr, out


def cargar(repo, drive):
    repo, drive = Path(repo), Path(drive)
    D = {}
    gp = repo / "datos/gis/IN_Piura_area_influencia_bloques_v6.gpkg"
    b = gpd.read_file(gp, layer="bloques_v6")
    ai = gpd.read_file(gp, layer="area_influencia_v6")
    amb = gpd.read_file(gp, layer="ambito_distritos")
    D["bloques_n"] = len(b)
    D["bloques_ha"] = float(b["AREA_HA"].sum())
    D["ai_n"] = len(ai)
    D["ai_bloques"] = ai["BLOQUE"].nunique()
    D["ai_ha"] = float(ai["AREA_HA"].sum())
    D["ai_geom_ha"] = float(ai["AREA_GEOM_HA"].sum())
    D["ambito_ha"] = float(amb["AREA_HA"].sum())
    D["microcuencas_bloques"] = b["MICROC"].nunique()
    prov = []
    for pv in sorted(b["NOMBPROV"].unique()):
        bb = b[b["NOMBPROV"] == pv]; aa = ai[ai["NOMBPROV"] == pv]
        prov.append(dict(prov=pv, dist=bb["NOMBDIST"].nunique(), bloques=len(bb), ha=float(bb["AREA_HA"].sum()),
                         ai_n=len(aa), ai_ha=float(aa["AREA_HA"].sum()), micro=bb["MICROC"].nunique()))
    D["por_provincia_v6"] = prov
    dist = []
    for (pv, ds), bb in b.groupby(["NOMBPROV", "NOMBDIST"]):
        aa = ai[(ai["NOMBPROV"] == pv) & (ai["NOMBDIST"] == ds)]
        dist.append(dict(prov=pv, dist=ds, bloques=len(bb), ha=float(bb["AREA_HA"].sum()), ai_n=len(aa),
                         ai_ha=float(aa["AREA_HA"].sum())))
    D["por_distrito_v6"] = dist

    mx = drive / "Matrices_E2_IN_Piura_B4.xlsx"
    _, dem = _tabla(mx, "08_Demanda_Distrito")
    D["demanda_distrito"] = [r for r in dem if r["Provincia"] in PROVS]
    D["demanda_total"] = next(r for r in dem if r["Provincia"] == "Total ámbito")
    _, pob = _tabla(mx, "05_Poblacion_Afectada")
    D["pob_distrito"] = [r for r in pob if r["Provincia"] in PROVS]
    D["pob_total"] = next(r for r in pob if r["Provincia"] == "Total ámbito")
    _, cob = _tabla(mx, "03_Cobertura_Distrito")
    D["cobertura"] = cob
    _, act = _tabla(mx, "13_Resumen_Actores")
    D["actores"] = act
    _, inv = _tabla(mx, "14_Involucrados_T1.7")
    D["involucrados"] = [r for r in inv if r["Provincia"] in PROVS]
    _, conc = _tabla(mx, "31_Conclusiones_Vacios")
    D["conclusiones_e2"] = [r for r in conc if r["Provincia"] in PROVS]
    _, riesgos = _tabla(mx, "15_Riesgos_Sociales") if _has(mx, "15_Riesgos_Sociales") else (None, [])
    D["riesgos_sociales"] = riesgos

    lb = drive / "Linea_Base_Brecha_E6_IN_Piura.xlsx"
    _, lbr = _tabla(lb, "Línea base")
    D["linea_base"] = lbr

    e6 = drive / "E6_Matriz_Ecosistemas_Bloques_V7.xlsx"
    rows = _rows(e6, "Resumen_Ecosistemas")
    i = next(k for k, r in enumerate(rows) if r and "Ecosistema" in [str(c) for c in r])
    hdr = list(rows[i]); eco = []
    for r in rows[i + 1:]:
        d = dict(zip(hdr, r))
        if d.get("Ecosistema"):
            eco.append(d)
    D["ecosistemas_v7"] = eco

    se = drive / "Interseccion_SERRN_AMA_Bloques_V7_IN_Piura.xlsx"
    rows = _rows(se, "Resumen_Distrito")
    tot = [c for c in next(r for r in rows if r and r[0] == "TOTAL") if c not in (None, "")]
    D["serrn_total"] = dict(bloques=tot[1], sin_superp=tot[2], area_bloques=tot[3], area_serrn=tot[4], pct=tot[5])

    di = drive / "Delimitacion_Interna_Bloques_E6_IN_Piura.xlsx"
    rows = _rows(di, "Resumen")
    clases = {}
    for r in rows:
        if r and r[0] in ("Restauración activa", "Enriquecimiento", "Conservación", "Tamizaje agrícola"):
            clases[r[0]] = (r[1], r[2])
    D["delim_interna"] = clases
    return D


def _has(path, sheet):
    wb = openpyxl.load_workbook(path, read_only=True)
    return sheet in wb.sheetnames
