"""
Paquete 06 - Informacion georreferenciada y cartografica del Perfil IN Piura.

Arma la carpeta 06 con:
  01_SHP   capas en shapefile (UTM WGS84 Zona 17S, EPSG:32717) con .shp .shx .dbf .prj .cpg
  02_KMZ   las capas principales en KMZ (WGS84 geograficas, EPSG:4326) para Google Earth
  03_GPKG  geopaquetes nativos del proyecto
  04_Mapas mapas tematicos (PDF/PNG) y mapas distritales de la presentacion del 03/11/2026
  05_Procesos scripts y leyendas usados para generar los productos

Uso:
  python build_gis.py --repo <raiz repo> --drive <carpeta descargas Drive (SHP sueltos)> --out <carpeta 06>

No inventa geometrias: solo copia o convierte capas existentes. Las capas sin .prj o con CRS
distinto se reproyectan a EPSG:32717 para el SHP y a EPSG:4326 para el KMZ.
"""
import argparse
import base64
import json
import os
import re
import shutil
import zipfile
from pathlib import Path

os.environ.setdefault("SHAPE_RESTORE_SHX", "YES")

import geopandas as gpd
import openpyxl
import pandas as pd
import pyogrio
from PIL import Image

CRS_UTM = 32717


def safe_name(s):
    import unicodedata
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = re.sub(r"[^\w\-]+", "_", s, flags=re.UNICODE).strip("_")
    return s


def write_shp(gdf, folder, name):
    folder = Path(folder) / name
    folder.mkdir(parents=True, exist_ok=True)
    g = gdf.copy()
    if g.crs is None:
        raise ValueError(f"{name}: capa sin CRS")
    g = g.to_crs(CRS_UTM)
    # nombres de campo <= 10 caracteres para DBF
    ren = {}
    used = set()
    for c in g.columns:
        if c == "geometry":
            continue
        n = c[:10]
        k = 1
        while n.upper() in used:
            n = c[: 10 - len(str(k))] + str(k)
            k += 1
        used.add(n.upper())
        if n != c:
            ren[c] = n
    g = g.rename(columns=ren)
    path = folder / f"{name}.shp"
    pyogrio.write_dataframe(g, path, driver="ESRI Shapefile", encoding="UTF-8")
    (folder / f"{name}.cpg").write_text("UTF-8")
    return path, ren


def write_kmz(gdf, folder, name, label_field=None):
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    g = gdf.to_crs(4326).copy()
    # KML: geometrias 2D
    g["geometry"] = g.geometry.force_2d()
    if label_field and label_field in g.columns:
        g.insert(0, "Name", g[label_field].astype(str))
    tmp_kml = folder / f"{name}.kml"
    pyogrio.write_dataframe(g, tmp_kml, driver="KML", layer=name)
    kmz = folder / f"{name}.kmz"
    with zipfile.ZipFile(kmz, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(tmp_kml, "doc.kml")
    tmp_kml.unlink()
    return kmz


def copy_shp_set(src_base, dst_folder, name):
    """Copia un shapefile nativo con todos sus complementarios."""
    src_base = Path(src_base)
    dst = Path(dst_folder) / name
    dst.mkdir(parents=True, exist_ok=True)
    n = 0
    for f in src_base.parent.iterdir():
        if f.name.startswith(src_base.name + ".") or f.stem == src_base.name:
            shutil.copy2(f, dst / f.name)
            n += 1
    return n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo", required=True)
    ap.add_argument("--drive", required=True, help="carpeta con SHP/GPKG descargados de Drive")
    ap.add_argument("--matrices-e2", required=True, help="Matrices_E2_IN_Piura_B4.xlsx")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    repo, drv, out = Path(a.repo), Path(a.drive), Path(a.out)
    if out.exists():
        shutil.rmtree(out)
    shp, kmz, gpkg, mapas, proc = (out / d for d in ("01_SHP", "02_KMZ", "03_GPKG", "04_Mapas", "05_Procesos"))
    for d in (shp, kmz, gpkg, mapas, proc):
        d.mkdir(parents=True)

    inv = []  # inventario de capas

    def reg(capa, formato, ruta, n, geom, fuente, nota=""):
        inv.append(dict(Capa=capa, Formato=formato, Ruta=str(Path(ruta).relative_to(out)), Registros=n,
                        Geometria=geom, Fuente=fuente, Nota=nota))

    ai_gpkg = repo / "datos/gis/IN_Piura_area_influencia_bloques_v6.gpkg"
    eco_gpkg = repo / "mapas/salidas/IN_Piura_ecosistemas_bloques_v6.gpkg"

    # ---------- capas del geopaquete de bloques y AI (repo) ----------
    capas_ai = [
        ("bloques_v6", "01_Bloques_V6_117", "BLOQUE", "Bloques preliminares de intervencion V6 (117; 12,270.235 ha)",
         "Bloques V6.kml - equipo SIG IN Piura; catalogo V6"),
        ("area_influencia_v6", "02_Area_Influencia_V6", "CODIGO",
         "Area de influencia directa aprobada (buffer 500 m recortado por microcuenca; 130 poligonos; 17,139.655 ha)",
         "AI_aprobado_2.shp - equipo SIG IN Piura (sin AI_74 y AI_75, bloques retirados)"),
        ("ambito_distritos", "03_Ambito_15_Distritos", "NOMBDIST", "Ambito politico de 15 distritos (referencial)",
         "INEI, limites distritales (capa simplificada)"),
    ]
    for lyr, name, lab, desc, src in capas_ai:
        g = gpd.read_file(ai_gpkg, layer=lyr)
        p, _ = write_shp(g, shp, name)
        reg(desc, "SHP", p, len(g), g.geom_type.iloc[0], src)
        k = write_kmz(g, kmz, name, lab)
        reg(desc, "KMZ", k, len(g), g.geom_type.iloc[0], src)

    # Bloques V7 (normalizada, misma superficie V6) - nativo Drive
    v7 = drv / "Bloques V7.zip"
    if v7.exists():
        d = shp / "04_Bloques_V7_normalizada"
        d.mkdir()
        with zipfile.ZipFile(v7) as z:
            z.extractall(d)
        g = gpd.read_file(d / "Bloques V7.shp")
        reg("Bloques V7 - relacion normalizada (base de E6, E8, E10)", "SHP nativo", d / "Bloques V7.shp", len(g),
            g.geom_type.iloc[0], "Equipo SIG IN Piura, 24/09/2026",
            "Catalogo V8 (12,157.47 ha netas) aun sin recorte de geometria en bloques 8, 71, 22, M17B1 y 5")
        k = write_kmz(g, kmz, "04_Bloques_V7_normalizada", [c for c in g.columns if c.upper().startswith("BLOQ")][0])
        reg("Bloques V7", "KMZ", k, len(g), g.geom_type.iloc[0], "Equipo SIG IN Piura")

    # Microcuencas / red hidrografica / carcavas (nativos Drive)
    nativos = [
        ("Microcuencas nivel 12", "05_Microcuencas_nivel_12", "CODIGO", "Microcuencas priorizadas (26) - unidades hidrograficas nivel 12",
         "HydroSHEDS / DMERNT-MINAM; recorte equipo SIG IN Piura"),
        ("Red hidrográfica nivel 12", "06_Red_Hidrografica", "NOMBRE", "Red hidrografica de las microcuencas",
         "IGN / equipo SIG IN Piura"),
        ("Cárcavas codificadas", "07_Carcavas_codificadas", "Codigo", "Carcavas identificadas en gabinete (55; 23.08 km)",
         "Equipo SIG IN Piura, 10/09/2026"),
    ]
    for base, name, lab, desc, src in nativos:
        sp = drv / f"{base}.shp"
        if not sp.exists():
            continue
        dst = shp / name
        dst.mkdir()
        for ext in (".shp", ".shx", ".dbf", ".prj"):
            f = drv / f"{base}{ext}"
            if f.exists():
                shutil.copy2(f, dst / f"{name}{ext}")
        (dst / f"{name}.cpg").write_text("UTF-8")
        g = gpd.read_file(dst / f"{name}.shp")
        reg(desc, "SHP nativo", dst / f"{name}.shp", len(g), g.geom_type.iloc[0], src)
        k = write_kmz(g, kmz, name, lab if lab in g.columns else None)
        reg(desc, "KMZ", k, len(g), g.geom_type.iloc[0], src)

    # Ecosistemas (repo)
    for lyr, name, lab, desc in [
        ("ecosistemas_ambito", "08_Ecosistemas_Ambito", "ECOSISTEMA",
         "Ecosistemas del ambito (Mapa Nacional de Ecosistemas MINAM 2018, recorte)"),
        ("ecosistemas_bloques_v6", "09_Ecosistemas_por_Bloque", "BLOQUE", "Ecosistemas dentro de los 117 bloques V6"),
    ]:
        g = gpd.read_file(eco_gpkg, layer=lyr)
        p, _ = write_shp(g, shp, name)
        reg(desc, "SHP", p, len(g), g.geom_type.iloc[0], "MINAM 2018 (R.M. N.° 440-2018-MINAM); proceso equipo IN Piura")
        k = write_kmz(g, kmz, name, lab)
        reg(desc, "KMZ", k, len(g), g.geom_type.iloc[0], "MINAM 2018; proceso equipo IN Piura")

    # Suelos E5 (repo, nativo)
    dst = shp / "10_Suelos_E5_Unidades_Cartograficas"
    n = copy_shp_set(repo / "datos/gis/suelos_e5/UNIDADES_CARTOGRAFICAS_SUELO", shp, "10_Suelos_E5_Unidades_Cartograficas")
    g = gpd.read_file(dst / "UNIDADES_CARTOGRAFICAS_SUELO.shp")
    reg("Unidades cartograficas de suelo - Estudio de Suelos E5 (423 delineaciones)", "SHP nativo",
        dst / "UNIDADES_CARTOGRAFICAS_SUELO.shp", len(g), g.geom_type.iloc[0], "E5 informe final (GEOSIG, 01/11/2026)")
    k = write_kmz(g, kmz, "10_Suelos_E5", "UNID_CARTO")
    reg("Unidades cartograficas de suelo E5", "KMZ", k, len(g), g.geom_type.iloc[0], "E5")

    # Pendientes vector, ecosistemas microcuencas y complemento (zip nativos del repo)
    for src_zip, name, desc, fuente in [
        ("datos/gis/pendientes/Pendientes_vector.zip", "11_Pendientes_vector",
         "Pendientes en clases (7 clases, DN 1-7; 57,255 poligonos) con leyenda JPEG", "Equipo SIG IN Piura, 03/09/2026"),
        ("datos/gis/ecosistemas_microcuencas.zip", "12_Ecosistemas_microcuencas_nativo",
         "Mapa Nacional de Ecosistemas recortado a microcuencas (EPSG:32718, nativo)", "MINAM 2018"),
        ("datos/gis/ecosistemas_complemento_M27B1_M9B1.zip", "13_Ecosistemas_complemento_M27B1_M9B1",
         "Complemento de ecosistemas para M27B1 y M9B1", "MINAM 2018; equipo SIG"),
    ]:
        d = shp / name
        d.mkdir()
        with zipfile.ZipFile(repo / src_zip) as z:
            z.extractall(d)
        shps = list(d.rglob("*.shp"))
        info = pyogrio.read_info(shps[0])
        reg(desc, "SHP nativo", shps[0], info["features"], info["geometry_type"], fuente,
            "Sin KMZ por volumen" if info["features"] > 5000 else "")

    # Lotes SUS y bloques descartados (Drive gpkg)
    sus = drv / "Lotes_SUS_Preliminares_UTM17S.gpkg"
    if sus.exists():
        for lyr, name, lab, desc in [
            ("lotes_sus_prelim", "14_Lotes_SUS_Preliminares", "cod_sus",
             "Lotes preliminares de sistemas de uso sostenible (60; 256.97 ha)"),
            ("bloques_descartados", "15_Bloques_Descartados", "bloque", "Bloques retirados (74 y 75)"),
        ]:
            g = gpd.read_file(sus, layer=lyr)
            p, _ = write_shp(g, shp, name)
            reg(desc, "SHP", p, len(g), g.geom_type.iloc[0], "Equipo SIG / SE2 liberacion de areas, 27/09/2026")
            k = write_kmz(g, kmz, name, lab)
            reg(desc, "KMZ", k, len(g), g.geom_type.iloc[0], "Equipo SIG")
        shutil.copy2(sus, gpkg / sus.name)
        reg("Lotes SUS + bloques V6 activos + descartados", "GPKG", gpkg / sus.name, "3 capas", "Polygon", "Equipo SIG")

    # Predios que intersectan bloques V7 (Drive zip nativo)
    pz = drv / "Predios_Interseccion_Bloques_V7.zip"
    if pz.exists():
        d = shp / "16_Predios_Interseccion_Bloques_V7"
        d.mkdir()
        with zipfile.ZipFile(pz) as z:
            z.extractall(d)
        s = list(d.rglob("*.shp"))[0]
        g = gpd.read_file(s)
        reg("Predios (COFOPRI / comunidades) que intersectan los bloques V7 - tamizaje predial de gabinete", "SHP nativo",
            s, len(g), g.geom_type.iloc[0], "Equipo SIG / SE2, 08/11/2026")
        k = write_kmz(g, kmz, "16_Predios_Interseccion_Bloques_V7")
        reg("Predios que intersectan bloques V7", "KMZ", k, len(g), g.geom_type.iloc[0], "Equipo SIG / SE2")

    # Interseccion SERRN y medidas E10 (Drive gpkg)
    for fn, name, lab, desc, fuente in [
        ("Interseccion_SERRN_AMA_Bloques_V7.gpkg", "17_Interseccion_SERRN_AMA_Bloques_V7", "BLOQUE",
         "Areas degradadas de nivel alto y muy alto que prestan SERRN dentro de los bloques V7 (3,930.43 ha)",
         "Capa SERRN (MINAM) x Bloques V7; equipo SIG, 07/11/2026"),
        ("E10_Medidas_IV_Bloques_V7.gpkg", "18_E10_Medidas_IV_Bloques_V7", "MEDIDA",
         "Distribucion preliminar de medidas de infraestructura verde por bloque (516 poligonos) - avance E10",
         "E10, equipo formulador, 10/11/2026 (preliminar, no son metas del Perfil)"),
    ]:
        f = drv / fn
        if not f.exists():
            continue
        lyr = pyogrio.list_layers(f)[0][0]
        g = gpd.read_file(f, layer=lyr)
        p, _ = write_shp(g, shp, name)
        reg(desc, "SHP", p, len(g), g.geom_type.iloc[0], fuente)
        k = write_kmz(g, kmz, name, lab)
        reg(desc, "KMZ", k, len(g), g.geom_type.iloc[0], fuente)
        shutil.copy2(f, gpkg / fn)
        reg(desc, "GPKG", gpkg / fn, len(g), g.geom_type.iloc[0], fuente)

    # Centros poblados del ambito (puntos) desde la Matriz E2 B4, hoja 09
    wb = openpyxl.load_workbook(a.matrices_e2, read_only=True, data_only=True)
    rows = list(wb["09_CP_Ambito"].iter_rows(values_only=True))
    hdr_i = next(i for i, r in enumerate(rows) if r and r[0] == "Provincia")
    hdr = list(rows[hdr_i])
    data = [r for r in rows[hdr_i + 1:] if r and r[0] in ("Morropón", "Huancabamba", "Ayabaca")]
    df = pd.DataFrame(data, columns=hdr)
    df = df[pd.to_numeric(df["Este (m)"], errors="coerce").notna()]
    df["Este (m)"] = df["Este (m)"].astype(float)
    df["Norte (m)"] = df["Norte (m)"].astype(float)
    ok = df["Este (m)"].between(450000, 750000) & df["Norte (m)"].between(9300000, 9600000)
    df = df[ok]
    cp = gpd.GeoDataFrame(df.rename(columns={
        "Provincia": "PROVINCIA", "Distrito (bloque)": "DIST_BLOQ", "Distrito (INEI)": "DIST_INEI",
        "Localidad": "LOCALIDAD", "Código INEI": "COD_INEI", "Bloque(s)": "BLOQUES", "Tipo": "TIPO",
        "Población 2017": "POB_2017", "Hombres": "HOMBRES", "Mujeres": "MUJERES",
        "Viviendas particulares": "VIV_PART", "Población vulnerable": "POB_VULN", "Este (m)": "ESTE",
        "Norte (m)": "NORTE", "Validación UTM": "VALID_UTM", "Microcuenca": "MICROC",
        "Dentro del AI aprobado": "EN_AI", "Con ficha social": "FICHA_DS"}),
        geometry=gpd.points_from_xy(df["Este (m)"], df["Norte (m)"]), crs=CRS_UTM)
    p, _ = write_shp(cp, shp, "19_Centros_Poblados_Ambito")
    reg(f"Localidades asociadas a los bloques (E2 B4, hoja 09; {len(cp)} con coordenada valida)", "SHP", p, len(cp),
        "Point", "INEI CPV 2017; catalogo CentrosPoblados-INEI-Bloques_v5 (CON SIN) 02.10.26; Matrices E2 B4 (07/11/2026)",
        f"{len(data) - len(cp)} localidades sin coordenada valida no se incluyen")
    k = write_kmz(cp, kmz, "19_Centros_Poblados_Ambito", "LOCALIDAD")
    reg("Localidades asociadas a los bloques", "KMZ", k, len(cp), "Point", "INEI CPV 2017; E2 B4")

    # CCPP INEI por distrito (zip nativo)
    cz = drv / "CENTROS POBLADOS DISTRITOS PIURA.zip"
    if cz.exists():
        d = shp / "20_Centros_Poblados_INEI_por_Distrito"
        d.mkdir()
        with zipfile.ZipFile(cz) as z:
            z.extractall(d)
        nn = sum(pyogrio.read_info(s)["features"] for s in d.rglob("*.shp"))
        reg("Centros poblados INEI por distrito (un shapefile por distrito)", "SHP nativo", d, nn, "Point",
            "INEI; compilacion equipo SIG IN Piura")

    # GPKG nativos del repo
    for f, desc in [(ai_gpkg, "Bloques V6, area de influencia, ambito y contexto distrital (4 capas)"),
                    (eco_gpkg, "Ecosistemas del ambito y por bloque (2 capas)")]:
        shutil.copy2(f, gpkg / f.name)
        reg(desc, "GPKG", gpkg / f.name, "varias", "Polygon", "Equipo formulador IN Piura")

    # ---------- mapas ----------
    for f in sorted((repo / "mapas/salidas").glob("*")):
        if f.suffix.lower() in (".pdf", ".png"):
            shutil.copy2(f, mapas / f.name)
            reg(f"Mapa tematico {f.stem}", f.suffix[1:].upper(), mapas / f.name, "", "", "Equipo formulador IN Piura")
    # mapas distritales embebidos en la presentacion del 03/11/2026
    html = (repo / "presentaciones/avance_perfil_sesdi_2026-11-03/Avance_Perfil_IN_Piura_SESDI_03_noviembre.html").read_text(encoding="utf8")
    maps = json.loads((repo / "presentaciones/avance_perfil_sesdi_2026-11-03/generador/maps.json").read_text(encoding="utf8"))
    figs = re.findall(r'<figure class="mp[^"]*" id="map-(\d+)"[^>]*><img src="data:image/webp;base64,([A-Za-z0-9+/=]+)"', html)
    dd = mapas / "Mapas_Distritales_Bloques"
    dd.mkdir()
    import io
    for i, b64 in figs:
        m = maps[int(i)]
        im = Image.open(io.BytesIO(base64.b64decode(b64))).convert("RGB")
        fn = f"Mapa_{int(i) + 1:02d}_{safe_name(m['p'])}_{safe_name(m['n'])}.jpg"
        im.save(dd / fn, quality=92, optimize=True)
        reg(f"Mapa distrital de bloques - {m['n']} ({m['p']})", "JPG", dd / fn, m["bl"], "", "Equipo SIG IN Piura (QGIS), base V6")

    # ---------- procesos ----------
    for f in ["mapas/mapa_area_influencia_bloques_v6.py", "mapas/mapa_ecosistemas_bloques_v6.py",
              "mapas/excel_bloques_ai_distrito.py", "mapas/README.md", "suelos_zanjas/aptitud_zanjas.py",
              "suelos_zanjas/leyenda_pendientes.json", "suelos_zanjas/README.md",
              "suelos_zanjas/salidas/aptitud_zanjas_bloques.json"]:
        src = repo / f
        dst = proc / f.replace("/", "__")
        shutil.copy2(src, dst)
        reg(f"Proceso / insumo: {f}", src.suffix[1:].upper(), dst, "", "", "Repositorio del equipo formulador")

    pd.DataFrame(inv).to_json(out / "_inventario_capas.json", orient="records", force_ascii=False, indent=1)
    print(f"{len(inv)} elementos inventariados en {out}")


if __name__ == "__main__":
    main()
