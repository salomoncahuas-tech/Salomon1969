#!/usr/bin/env python3
"""
Descarga mes a mes las tablas "_dato" del mapa de estaciones SENAMHI y arma
CSV diarios (fecha, caudal_m3s, nivel_m, estacion, codigo, fuente).

ESTADO: SIN PROBAR. Al 27/11/2026 el host www.senamhi.gob.pe estaba bloqueado
por el proxy de la sesión (HTTP 403 en CONNECT). Los parámetros del endpoint
(tipo02/tipo03, cate_esta, cod_old) son "Por verificar": revisar el JS de
https://www.senamhi.gob.pe/mapas/mapa-estaciones-2/ antes de usar.

No rellena ni interpola: solo guarda valores leídos de las tablas.
Uso:  python3 -I descargar_senamhi_diarios.py 2000 2026
"""
import io, sys, time, re
from pathlib import Path
import requests
import pandas as pd

OUT = Path(__file__).resolve().parent
RAW = OUT / "raw_senamhi"
BASE = "https://www.senamhi.gob.pe/mapas/mapa-estaciones-2/"
# Solo códigos confirmados en URLs de avisos SENAMHI. Añadir otros cuando se verifiquen.
ESTACIONES = {
    "tambogrande":   {"codigo": "47E0522C", "nombre": "Tambogrande"},
    "puente_nacara": {"codigo": "47E0415A", "nombre": "Puente Ñácara"},
}
ENDPOINTS = ["_dato_esta_tipo02.php", "_dato_esta_tipo03.php"]  # Por verificar
PARAMS = dict(t_e="H", estado="AUTOMATICA", cod_old="", cate_esta="EHA")  # Por verificar
DELAY = 1.5


def fetch_month(sess, ep, codigo, yyyymm):
    p = dict(PARAMS, estaciones=codigo, CBOFiltro=yyyymm)
    r = sess.get(BASE + ep, params=p, timeout=60)
    return r.status_code, r.url, r.text


def parse(html):
    try:
        tables = pd.read_html(io.StringIO(html), decimal=".", thousands=None)
    except ValueError:
        return None
    for t in tables:
        cols = " ".join(map(str, t.columns)).lower() + " " + " ".join(map(str, t.iloc[:2].values.ravel())).lower()
        if "caudal" in cols or "nivel" in cols:
            return t
    return None


def main(y0, y1):
    RAW.mkdir(exist_ok=True)
    sess = requests.Session()
    sess.headers["User-Agent"] = "Mozilla/5.0 (investigacion hidrologica)"
    log = []
    for key, est in ESTACIONES.items():
        rows = []
        for y in range(y0, y1 + 1):
            for m in range(1, 13):
                ym = f"{y}{m:02d}"
                for ep in ENDPOINTS:
                    try:
                        code, url, html = fetch_month(sess, ep, est["codigo"], ym)
                    except Exception as e:
                        log.append(f"{key}\t{ym}\t{ep}\tERROR {e}")
                        time.sleep(DELAY); continue
                    (RAW / f"{key}_{ym}_{ep}.html").write_text(html, encoding="utf-8")
                    t = parse(html) if code == 200 else None
                    log.append(f"{key}\t{ym}\t{ep}\tHTTP {code}\t{'tabla' if t is not None else 'sin tabla'}\t{url}")
                    time.sleep(DELAY)
                    if t is None:
                        continue
                    t.columns = [re.sub(r"\s+", " ", str(c)).strip().lower() for c in t.columns]
                    fcol = next((c for c in t.columns if "fecha" in c or "día" in c or "dia" in c), t.columns[0])
                    qcol = next((c for c in t.columns if "caudal" in c), None)
                    ncol = next((c for c in t.columns if "nivel" in c), None)
                    hcol = next((c for c in t.columns if "hora" in c), None)
                    df = pd.DataFrame({
                        "fecha": pd.to_datetime(t[fcol], errors="coerce", dayfirst=True),
                        "caudal_m3s": pd.to_numeric(t[qcol], errors="coerce") if qcol else pd.NA,
                        "nivel_m": pd.to_numeric(t[ncol], errors="coerce") if ncol else pd.NA,
                    }).dropna(subset=["fecha"])
                    if hcol:  # datos horarios -> media diaria SOLO de valores existentes
                        df = df.groupby(df["fecha"].dt.normalize()).agg(
                            caudal_m3s=("caudal_m3s", "mean"), nivel_m=("nivel_m", "mean"),
                            n_lecturas=("caudal_m3s", "count")).reset_index()
                    rows.append(df)
                    break
        if not rows:
            continue
        d = pd.concat(rows).drop_duplicates("fecha").sort_values("fecha")
        d["fecha"] = d["fecha"].dt.strftime("%Y-%m-%d")
        d["estacion"], d["codigo"] = est["nombre"], est["codigo"]
        d["fuente"] = "SENAMHI mapa-estaciones-2 (_dato), descargado " + time.strftime("%Y-%m-%d")
        d.to_csv(OUT / f"diario_{key}_senamhi.csv", index=False)
        # cobertura
        f = pd.to_datetime(d["fecha"]); full = pd.date_range(f.min(), f.max(), freq="D")
        ok = set(f[d["caudal_m3s"].notna().values])
        miss = [x for x in full if x not in ok]
        print(f"{est['nombre']}: {f.min().date()} a {f.max().date()}, faltantes {100*len(miss)/len(full):.2f}%")
        print(pd.Series([x.year for x in miss]).value_counts().sort_index().to_string())
    (OUT / "log_descarga_senamhi.tsv").write_text("\n".join(log), encoding="utf-8")


if __name__ == "__main__":
    main(int(sys.argv[1]), int(sys.argv[2]))
