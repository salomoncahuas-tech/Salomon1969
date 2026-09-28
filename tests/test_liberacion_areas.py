# -*- coding: utf-8 -*-
"""
Casos de prueba del módulo de Liberación de Áreas (datos SINTÉTICOS: nombres, DNI y geometrías ficticios).
- Pruebas de lógica pura: `python -m pytest tests/ -q`
- Prueba de integración con PostgreSQL: definir LA_TEST_DB="postgresql://…" (base VACÍA de pruebas, NUNCA la de
  producción: la prueba borra y recrea el esquema public). Crea `bloques` y `verificacion_campo_odk` con el DDL
  real de database.py y los 129 bloques del catálogo del aplicativo, y verifica el enlace con la_unidades.
Genera tests/casos_prueba_liberacion_areas.json con envíos simulados de F-LA-01/02/03/04/06.
"""
import json
import os
import sys
from pathlib import Path

import pandas as pd
from pyproj import Transformer
from shapely.geometry import box

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
from liberacion_areas import la_core as core  # noqa: E402

CSV = RAIZ / "datos" / "unidades_liberacion_areas.csv"
INV = Transformer.from_crs(32717, 4326, always_xy=True)

# Geometrías SINTÉTICAS (cuadrados) para las pruebas; en la app se cargan los polígonos reales (GeoJSON 32717)
E0, N0 = 620_000, 9_430_000
GEOMS = {"27": box(E0, N0, E0 + 884, N0 + 884),                     # ~78.15 ha
         "M20B1": box(E0 + 5000, N0, E0 + 5000 + 1672.4, N0 + 1672.4)}  # ~279.68 ha


def gp(e, n, prec=4.0):
    lon, lat = INV.transform(e, n)
    return f"{lat:.7f} {lon:.7f} 1500 {prec}"


def gs(e, n, lado):
    pts = [(e, n), (e + lado, n), (e + lado, n + lado), (e, n + lado), (e, n)]
    return ";".join(gp(x, y) for x, y in pts)


def catalogo():
    d = pd.read_csv(CSV, dtype=str)
    return {r["name"]: r for r in d.to_dict("records")}


def casos():
    base = {"hoy": "2026-10-05", "asistente": "AP-1", "provincia": "ayabaca", "distrito": "frias"}
    return {
        "f_la_01_reunion": [{**base, "meta/instanceID": "uuid:t01-a", "g_ubicacion/unidades": "27 3", "cod_unidad": "27 3",
                             "g_evento/tipo_evento": "asamblea_comunal", "g_evento/fecha_evento": "2026-10-05",
                             "g_evento/gps_evento": gp(E0 + 100, N0 - 300), "g_asistencia/asist_hombres": 20,
                             "g_asistencia/asist_mujeres": 12, "g_temas/aceptacion": "favorable", "g_temas/alertas": "ninguna",
                             "g_evidencias/a01_suscrita": "si", "g_evidencias/a02_suscrita": "si"}],
        "f_la_02_titular": [
            {**base, "meta/instanceID": "uuid:t02-a", "tipo_unidad": "bloque", "unidad": "27", "cod_unidad": "27", "n_predio": 1,
             "cod_predio": "27-P01", "consentimiento": "si", "tipo_titularidad": "privado", "tit_nombres": "JUAN (FICTICIO)",
             "tit_apellidos": "PRUEBA UNO", "tit_dni": "00000001", "tit_estado_civil": "soltero", "predio_area_unidad_ha": "40",
             "aceptacion": "favorable", "estado_la_propuesto": "LA-2"},
            {**base, "meta/instanceID": "uuid:t02-b", "tipo_unidad": "bloque", "unidad": "27", "cod_unidad": "27", "n_predio": 2,
             "cod_predio": "27-P02", "consentimiento": "si", "tipo_titularidad": "sucesion", "sucesion_estado": "sin_tramite",
             "tit_nombres": "ANA (FICTICIA)", "tit_apellidos": "PRUEBA DOS", "tit_dni": "00000002", "tit_estado_civil": "viudo",
             "predio_area_unidad_ha": "38", "aceptacion": "condicionada"}],
        "f_la_03_inspeccion": [
            {**base, "meta/instanceID": "uuid:t03-a", "tipo_unidad": "bloque", "unidad": "27", "cod_unidad": "27", "n_predio": 1,
             "cod_predio": "27-P01", "punto_interior": gp(E0 + 400, N0 + 400), "interferencias": "ninguna",
             "conclusion_campo": "sin_restriccion", "r_puntos": [{"r_puntos/pt_gps": gp(E0 + 450, N0 + 300), "r_puntos/pt_tipo": "vertice"}]},
            {**base, "meta/instanceID": "uuid:t03-b", "tipo_unidad": "bloque", "unidad": "27", "cod_unidad": "27", "n_predio": 2,
             "cod_predio": "27-P02", "punto_interior": gp(E0 - 120, N0 + 400), "interferencias": "ninguna"},     # FUERA 120 m
            {"hoy": "2026-10-06", "asistente": "AP-6", "provincia": "huancabamba", "distrito": "san_miguel_de_el_faique",
             "meta/instanceID": "uuid:t03-c", "tipo_unidad": "lote_sus", "unidad": "SUS-059", "cod_unidad": "SUS-059",
             "bloque_ref": "M20B1", "area_bloque_ha": "279.680", "n_predio": 1, "cod_predio": "SUS-059-P01",
             "punto_interior": gp(E0 + 5100, N0 + 100), "poligono_sus": gs(E0 + 5050, N0 + 50, 123),   # 1.513 ha → Σ > 10 %
             "posicion_sus_campo": "DENTRO", "interferencias": "ninguna"}],
        "f_la_04_actas": [
            {**base, "meta/instanceID": "uuid:t04-a", "tipo_unidad": "bloque", "unidad": "27", "cod_unidad": "27", "n_predio": 1,
             "cod_predio": "27-P01", "tipo_acta": "A-04", "cod_doc": "A04-27-P01", "fecha_acta": "2026-10-07",
             "area_comprometida_ha": "40", "checklist": "coordenadas croquis titularidad gratuidad plazo firmas fedatario dj datos legible",
             "r_firmantes": [{"r_firmantes/fir_nombre": "JUAN (FICTICIO) PRUEBA UNO", "r_firmantes/fir_dni": "00000001",
                              "r_firmantes/fir_rol": "titular", "r_firmantes/fir_huella": "si"}],
             "fedatario_tipo": "juez_paz", "estado_acta": "completa"},
            {**base, "meta/instanceID": "uuid:t04-b", "tipo_unidad": "bloque", "unidad": "27", "cod_unidad": "27", "n_predio": 2,
             "cod_predio": "27-P02", "tipo_acta": "A-04", "cod_doc": "A04-27-P02", "fecha_acta": "2026-10-07",
             "area_comprometida_ha": "38", "checklist": "firmas datos", "r_firmantes": [
                 {"r_firmantes/fir_nombre": "ANA (FICTICIA)", "r_firmantes/fir_dni": "00000002", "r_firmantes/fir_rol": "heredero",
                  "r_firmantes/fir_huella": "no"}], "fedatario_tipo": "ninguno", "estado_acta": "observada"}],
        "f_la_06_vivero": [{"hoy": "2026-10-08", "asistente": "ESP", "provincia": "morropon", "distrito": "chulucanas",
                            "meta/instanceID": "uuid:t06-a", "cod_vivero": "VIV-01", "cod_unidad": "VIV-01",
                            "alt_nombre": "ALTERNATIVA DE PRUEBA", "modalidad": "convenio", "tipo_titularidad": "estatal",
                            "titular_nombre": "ENTIDAD FICTICIA", "gps_centro": gp(E0 - 20000, N0 - 10000),
                            "area_ha": "1.2", "pendiente": "p0_5", "agua_fuente": "canal", "agua_caudal_ls": "2",
                            "agua_permanencia": "permanente", "acceso_tipo": "carretera", "acceso_camion": "si", "energia": "red",
                            "inundabilidad": "bajo", "deslizamiento": "nulo", "disposicion": "favorable"}],
    }


def _validar_todo(existentes=frozenset()):
    cat = catalogo()
    sus = core.sus_areas_desde_catalogo(cat)
    out = {}
    for fid, regs in casos().items():
        out[fid] = [core.validar_envio(fid, core.aplanar(r), cat, GEOMS, set(existentes), sus) for r in regs]
    return out


# ------------------------------------------------------------------ pruebas de lógica
def test_utm_y_rango():
    p = core.parse_geopoint(gp(E0, N0))
    assert abs(p["este"] - E0) < 0.05 and abs(p["norte"] - N0) < 0.05 and p["en_rango"]
    assert not core.parse_geopoint("-5.0 -60.0 0 3")["en_rango"]


def test_punto_dentro_y_fuera():
    r = _validar_todo()["f_la_03_inspeccion"]
    assert r[0].datos["validacion_espacial"] == "DENTRO" and r[0].estado_import == "NUEVO"
    assert r[1].datos["validacion_espacial"] == "FUERA_POLIGONO" and 115 < r[1].datos["distancia_m"] < 125
    assert r[1].estado_import == "OBSERVADO"


def test_regla_10_por_ciento_sus():
    r = _validar_todo()["f_la_03_inspeccion"][2]
    assert r.datos["posicion_sus_calc"] == "DENTRO"
    assert abs(r.datos["area_sus_ha"] - 1.513) < 0.01
    assert any("10 %" in m for m in r.motivos), r.motivos       # SUS-058 (26.689) + 1.513 > 27.968


def test_contiguo_50m():
    bloque = GEOMS["27"]
    pol = box(E0 + 884 + 30, N0 + 100, E0 + 884 + 150, N0 + 220)
    assert core.posicion_sus(pol, bloque) == ("CONTIGUO", 30.0)
    pol2 = box(E0 + 884 + 60, N0 + 100, E0 + 884 + 200, N0 + 220)
    assert core.posicion_sus(pol2, bloque)[0] == "FUERA"


def test_actas_y_duplicados():
    r = _validar_todo()["f_la_04_actas"]
    assert r[0].estado_import == "NUEVO", r[0].motivos
    assert r[1].estado_import == "OBSERVADO" and any("incompleta" in m for m in r[1].motivos)
    d = _validar_todo(existentes={"t04-a"})["f_la_04_actas"][0]
    assert d.estado_import == "DUPLICADO"


def test_estados_y_semaforo():
    assert core.estado_predio({"titular": True, "socializado": True, "inspeccion_dentro": True, "acta_conforme": True}) == "LA-4"
    assert core.estado_predio({"titular": True, "socializado": False, "inspeccion_dentro": True}) == "LA-1"
    assert core.estado_predio({"excepcion": "NEG", "titular": True}) == "NEG"
    assert core.semaforo_unidad([{"estado": "LA-4", "area_ha": 70}], 78.154)[0] == "VERDE"
    assert core.semaforo_unidad([{"estado": "LA-4", "area_ha": 40}], 78.154)[0] == "AMBAR"
    assert core.semaforo_unidad([{"estado": "LA-4", "area_ha": 70}, {"estado": "NEG", "area_ha": 5, "nucleo": True}], 78.154)[0] == "ROJO"
    assert core.siguiente_titular(["T0001", "T0009"]) == "T0010"


def _ddl_app(tabla):
    """CREATE TABLE real de database.py (se lee como texto: database.py exige st.secrets al importarse)."""
    import re
    fuente = (RAIZ / "database.py").read_text(encoding="utf-8")
    m = re.search(rf"(CREATE TABLE IF NOT EXISTS {tabla} \(.*?\n        \))", fuente, re.DOTALL)
    return m.group(1)


def _bloques_app():
    """Catálogo BLOQUES_V5 de streamlit_app.py: (id, codigo, microcuenca, area, provincia, distrito, …)."""
    import ast
    src = (RAIZ / "streamlit_app.py").read_text(encoding="utf-8")
    i = src.index("BLOQUES_V5 = [")
    j = src.index("\n]\n", i) + 2
    return ast.literal_eval(src[i + len("BLOQUES_V5 = "):j])


def _retirados_app():
    """BLOQUES_RETIRADOS de database.py (leído como texto)."""
    import ast
    import re
    fuente = (RAIZ / "database.py").read_text(encoding="utf-8")
    m = re.search(r"BLOQUES_RETIRADOS = (\(.*?\))\n", fuente, re.DOTALL)
    return set(ast.literal_eval(m.group(1)))


def test_catalogo_vigente_del_app_es_v6():
    """Todo bloque del catálogo del aplicativo que no está en V6 debe estar retirado, y ningún bloque V6 retirado."""
    v6 = set(pd.read_csv(CSV, dtype=str).query("tipo_unidad == 'bloque'")["name"])
    app = {str(b[1]) for b in _bloques_app()}
    ret = _retirados_app()
    assert v6 <= app
    assert app - v6 <= ret, sorted(app - v6 - ret)
    assert not (v6 & ret), sorted(v6 & ret)
    assert len(app - ret) == 117


def test_integracion_postgres():
    url = os.environ.get("LA_TEST_DB")
    if not url:
        return
    from liberacion_areas import la_db as db
    conn = db.conectar(url)
    with conn.cursor() as cur:
        cur.execute("DROP SCHEMA public CASCADE; CREATE SCHEMA public;")
        cur.execute(_ddl_app("bloques"))
        cur.execute(_ddl_app("inspecciones"))
        cur.execute(_ddl_app("verificacion_campo_odk"))
        for b in _bloques_app():
            cur.execute("""INSERT INTO bloques (codigo, tipo_intervencion, cuenca, distrito, utm_este, utm_norte,
                           area_hectareas, microcuenca, provincia, fecha_registro)
                           VALUES (%s,'Restauracion','Cuenca Alta del Rio Piura',%s,%s,%s,%s,%s,%s,'2026-09-28')""",
                        (b[1], b[5], b[8], b[9], b[3], b[2], b[4]))
        cur.execute("""INSERT INTO verificacion_campo_odk (clave_envio, codigo_bloque, fecha_visita, tenencia, n_predios,
                       titular, aceptacion_titular, fecha_registro) VALUES
                       ('t-p6','27','2026-09-20','Privada','3','JUAN (FICTICIO)','Acepta','2026-09-20')""")
    conn.commit()
    r = db.inicializar_la(conn, CSV)
    assert r["cargadas"] == 177 and r["enlazados"] == 117 and r["sin_enlace"] == [], r
    conc = db.conciliacion(conn)
    assert (conc.estado == "OK").sum() == 117, conc.estado.value_counts()
    fuera = sorted(conc[conc.estado == "ACTIVO_EN_APP_FUERA_DE_V6"].codigo)
    assert fuera == sorted(["1", "25", "29", "32", "33", "46", "48", "68", "7", "74", "75", "M18B5"]), fuera
    # misma instrucción que database.inicializar_bd() aplica en cada arranque
    with conn.cursor() as cur:
        cur.execute("UPDATE bloques SET activo = 0 WHERE codigo = ANY(%s) AND COALESCE(activo, 1) <> 0",
                    (sorted(_retirados_app()),))
    conn.commit()
    conc = db.conciliacion(conn)
    assert (conc.estado == "ACTIVO_EN_APP_FUERA_DE_V6").sum() == 0 and (conc.estado == "OK").sum() == 117
    lotes = db.df(conn, "SELECT count(*) n FROM la_unidades WHERE tipo_unidad='lote_sus' AND bloque_ref_id IS NULL")
    assert lotes["n"][0] == 0
    assert db.antecedentes_paso6(conn).iloc[0]["n_predios"] == 3
    # ON DELETE SET NULL: borrar un bloque del aplicativo no borra la unidad
    with conn.cursor() as cur:
        cur.execute("DELETE FROM bloques WHERE codigo='74'")
    conn.commit()

    cat = db.catalogo(conn)
    with conn.cursor() as cur:
        for cod, g in GEOMS.items():
            cur.execute("UPDATE la_unidades SET geom_wkt=%s WHERE codigo=%s", (g.wkt, cod))
    conn.commit()
    geoms = db.geometrias(conn)
    assert set(geoms) == set(GEOMS)
    orden = ["f_la_01_reunion", "f_la_02_titular", "f_la_03_inspeccion", "f_la_04_actas", "f_la_06_vivero"]
    for fid in orden:
        regs = casos()[fid]
        res = [core.validar_envio(fid, core.aplanar(x), cat, geoms, db.uuids_existentes(conn), db.sus_areas(conn, cat)) for x in regs]
        db.importar(conn, res, "TEST", fid, "pytest")
        db.guardar_adjunto(conn, res[0].kobo_uuid, res[0].cod_predio, "foto_1", f"{fid}_foto_1_1.jpg", "image/jpeg", "", b"\xff\xd8")
    res = [core.validar_envio("f_la_04_actas", core.aplanar(x), cat, geoms, db.uuids_existentes(conn)) for x in casos()["f_la_04_actas"]]
    assert all(x.estado_import == "DUPLICADO" for x in res)
    m = db.df(conn, "SELECT cod_predio, estado_la FROM la_predios ORDER BY cod_predio").set_index("cod_predio")
    assert m.loc["27-P01", "estado_la"] == "LA-4", m
    assert m.loc["27-P02", "estado_la"] == "LA-2", m
    assert db.df(conn, "SELECT count(*) n FROM la_adjuntos")["n"][0] == 5
    assert db.df(conn, "SELECT puntaje FROM la_vivero_alternativas")["puntaje"][0] > 80
    # re-inicializar es idempotente y no duplica el catálogo
    assert db.inicializar_la(conn, CSV)["cargadas"] == 0
    conn.close()


if __name__ == "__main__":
    salida = RAIZ / "tests" / "casos_prueba_liberacion_areas.json"
    salida.write_text(json.dumps(casos(), ensure_ascii=False, indent=1), encoding="utf-8")
    print("Casos de prueba escritos en", salida)
