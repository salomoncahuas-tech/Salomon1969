# -*- coding: utf-8 -*-
"""Acceso a Supabase/PostgreSQL del módulo de Liberación de Áreas.
Usa la MISMA conexión del aplicativo: st.secrets["DATABASE_URL"] (pooler de Supabase), con los
mismos parámetros de database.py (sslmode=require, keepalives, reintentos). Nunca hay credenciales
en el código."""
from __future__ import annotations

import json
import time
from pathlib import Path

import pandas as pd
import psycopg2
import psycopg2.extras
from shapely import wkt as shp_wkt

from . import la_core as core

SCHEMA_SQL = Path(__file__).resolve().parent / "la_schema.sql"

_CONNECT_KWARGS = {"connect_timeout": 15, "keepalives": 1, "keepalives_idle": 30,
                   "keepalives_interval": 10, "keepalives_count": 5}


def conectar(url: str | None = None, sslmode: str | None = None):
    """Sin argumentos usa st.secrets["DATABASE_URL"] con sslmode=require (igual que database.py).
    Con `url` explícita (pruebas) solo aplica sslmode si se indica."""
    if url is None:
        import streamlit as st
        url, sslmode = st.secrets["DATABASE_URL"], sslmode or "require"
    kw = dict(_CONNECT_KWARGS)
    if sslmode:
        kw["sslmode"] = sslmode
    ultimo = None
    for intento in range(4):
        try:
            conn = psycopg2.connect(url, **kw)
            conn.autocommit = False
            return conn
        except psycopg2.OperationalError as e:
            ultimo = e
            if intento < 3:
                time.sleep(2 ** intento)
    raise ultimo


def crear_esquema(conn):
    with conn.cursor() as cur:
        cur.execute(SCHEMA_SQL.read_text(encoding="utf-8"))
    conn.commit()


def inicializar_la(conn, csv_catalogo: Path | None = None) -> dict:
    """Crea/actualiza las tablas la_* (aditivo) y, si la_unidades está vacía, carga el catálogo
    V6 del repositorio (datos/unidades_liberacion_areas.csv). Luego enlaza con `bloques`."""
    crear_esquema(conn)
    n = df(conn, "SELECT count(*) AS n FROM la_unidades")["n"][0]
    cargadas = 0
    csv_catalogo = csv_catalogo or (Path(__file__).resolve().parent.parent / "datos" / "unidades_liberacion_areas.csv")
    if n == 0 and csv_catalogo.exists():
        cargadas = cargar_unidades_csv(conn, pd.read_csv(csv_catalogo, dtype=str))
    return {"cargadas": cargadas, **vincular_bloques(conn)}


def _existe(conn, tabla: str) -> bool:
    with conn.cursor() as cur:
        cur.execute("SELECT to_regclass(%s) IS NOT NULL", (f"public.{tabla}",))
        return bool(cur.fetchone()[0])


def vincular_bloques(conn) -> dict:
    """Enlaza la_unidades con bloques(id) por igualdad EXACTA de código (el mismo criterio de
    bloque_lookup: nunca se adivina un bloque). No modifica la tabla bloques."""
    if not _existe(conn, "bloques"):
        return {"enlazados": 0, "sin_enlace": []}
    with conn.cursor() as cur:
        cur.execute("""UPDATE la_unidades u SET bloque_id = b.id FROM bloques b
                       WHERE u.tipo_unidad = 'bloque' AND b.codigo = u.codigo
                         AND u.bloque_id IS DISTINCT FROM b.id""")
        cur.execute("""UPDATE la_unidades u SET bloque_ref_id = b.id FROM bloques b
                       WHERE u.tipo_unidad = 'lote_sus' AND b.codigo = u.bloque_ref
                         AND u.bloque_ref_id IS DISTINCT FROM b.id""")
        cur.execute("SELECT count(*) FROM la_unidades WHERE tipo_unidad='bloque' AND bloque_id IS NOT NULL")
        enl = cur.fetchone()[0]
        cur.execute("SELECT codigo FROM la_unidades WHERE tipo_unidad='bloque' AND bloque_id IS NULL ORDER BY codigo")
        sin = [r[0] for r in cur.fetchall()]
    conn.commit()
    return {"enlazados": enl, "sin_enlace": sin}


def conciliacion(conn) -> pd.DataFrame:
    if not _existe(conn, "la_v_conciliacion_bloques"):
        return pd.DataFrame()
    return df(conn, "SELECT * FROM la_v_conciliacion_bloques ORDER BY estado, codigo")


def antecedentes_paso6(conn) -> pd.DataFrame:
    if not _existe(conn, "la_v_antecedentes_paso6"):
        return pd.DataFrame()
    return df(conn, "SELECT * FROM la_v_antecedentes_paso6 ORDER BY codigo_bloque")


def guardar_adjunto(conn, kobo_uuid, cod_predio, campo, nombre, mimetype, url, contenido: bytes | None):
    with conn.cursor() as cur:
        cur.execute("""INSERT INTO la_adjuntos (kobo_uuid, cod_predio, campo, nombre_archivo, mimetype, url_original,
                       contenido, tamano_bytes) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
                       ON CONFLICT (kobo_uuid, nombre_archivo) DO NOTHING""",
                    (kobo_uuid, cod_predio, campo, nombre, mimetype, url,
                     psycopg2.Binary(contenido) if contenido else None, len(contenido or b"")))
    conn.commit()


def df(conn, sql: str, params=None) -> pd.DataFrame:
    with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
        cur.execute(sql, params)
        return pd.DataFrame(cur.fetchall())


# ------------------------------------------------------------------ catálogo
def cargar_unidades_csv(conn, unidades: pd.DataFrame, geoms: dict | None = None) -> int:
    """Carga / actualiza la_unidades desde unidades.csv (mismo archivo que usa Kobo) y, opcionalmente, geometrías."""
    geoms = geoms or {}
    filas = []
    for r in unidades.fillna("").to_dict("records"):
        area_b = core._num(r.get("area_bloque_ha"))
        area = core._num(r.get("area_ha"))
        g = geoms.get(str(r["name"]))
        filas.append((str(r["name"]), r.get("label"), r["tipo_unidad"], r["provincia"], r["distrito"],
                      str(r.get("bloque_ref") or "") or None, area, area_b,
                      round(area / area_b * 100, 2) if (r["tipo_unidad"] == "lote_sus" and area and area_b) else None,
                      r.get("posicion_sus") or None, r.get("asistente") or None, r.get("asistente_nombre") or None,
                      g.wkt if g is not None else None))
    sql = """INSERT INTO la_unidades (codigo,label,tipo_unidad,provincia,distrito,bloque_ref,area_ha,area_bloque_ha,
             pct_bloque,posicion_sus,asistente,asistente_nombre,geom_wkt) VALUES %s
             ON CONFLICT (codigo) DO UPDATE SET label=EXCLUDED.label, tipo_unidad=EXCLUDED.tipo_unidad,
             provincia=EXCLUDED.provincia, distrito=EXCLUDED.distrito, bloque_ref=EXCLUDED.bloque_ref,
             area_ha=EXCLUDED.area_ha, area_bloque_ha=EXCLUDED.area_bloque_ha, pct_bloque=EXCLUDED.pct_bloque,
             posicion_sus=EXCLUDED.posicion_sus, asistente=EXCLUDED.asistente, asistente_nombre=EXCLUDED.asistente_nombre,
             geom_wkt=COALESCE(EXCLUDED.geom_wkt, la_unidades.geom_wkt), actualizado=now()"""
    with conn.cursor() as cur:
        psycopg2.extras.execute_values(cur, sql, filas)
    conn.commit()
    return len(filas)


def cargar_geometrias(conn, geoms: dict) -> int:
    with conn.cursor() as cur:
        psycopg2.extras.execute_batch(cur, "UPDATE la_unidades SET geom_wkt=%s, actualizado=now() WHERE codigo=%s",
                                      [(g.wkt, c) for c, g in geoms.items()])
        n = cur.rowcount
    conn.commit()
    return n


def catalogo(conn) -> dict[str, dict]:
    d = df(conn, "SELECT * FROM la_unidades WHERE activo")
    return {r["codigo"]: r for r in d.to_dict("records")} if not d.empty else {}


def geometrias(conn) -> dict:
    d = df(conn, "SELECT codigo, geom_wkt FROM la_unidades WHERE geom_wkt IS NOT NULL")
    return {r["codigo"]: shp_wkt.loads(r["geom_wkt"]) for r in d.to_dict("records")} if not d.empty else {}


def uuids_existentes(conn) -> set[str]:
    d = df(conn, "SELECT kobo_uuid FROM la_envios_raw")
    return set(d["kobo_uuid"]) if not d.empty else set()


def sus_areas(conn, cat: dict) -> dict:
    """Área vigente por lote SUS: medición de campo más reciente (F-LA-03) o, si no hay, la de gabinete."""
    base = core.sus_areas_desde_catalogo(cat)
    d = df(conn, """SELECT DISTINCT ON (cod_unidad) cod_unidad, area_sus_ha FROM la_inspecciones
                    WHERE area_sus_ha IS NOT NULL ORDER BY cod_unidad, fecha DESC""")
    for r in (d.to_dict("records") if not d.empty else []):
        if r["cod_unidad"] in base:
            base[r["cod_unidad"]] = (base[r["cod_unidad"]][0], float(r["area_sus_ha"]))
    return base


# ------------------------------------------------------------------ importación
def _s(v):
    if isinstance(v, (list, dict)):
        return json.dumps(v, ensure_ascii=False)
    return None if v in ("", None) else v


ORIGEN_POR_FUENTE = {"API": "KOBO", "ARCHIVO": "KOBO", "APP": "APP", "PLANTILLA": "PLANTILLA"}


def importar(conn, resultados: list[core.Resultado], fuente: str, form_id: str, usuario: str,
             origen: str | None = None) -> dict:
    """Inserta solo NUEVO y OBSERVADO (los observados quedan registrados con su motivo); ignora DUPLICADO.
    Idempotente: INSERT … ON CONFLICT (kobo_uuid) DO NOTHING.
    fuente: API / ARCHIVO (Kobo), APP (digitado en el aplicativo), PLANTILLA (Excel ANIN), RESTAURACION."""
    origen = origen or ORIGEN_POR_FUENTE.get(fuente, "KOBO")
    nuevos = [r for r in resultados if r.estado_import in ("NUEVO", "OBSERVADO") and r.kobo_uuid]
    with conn.cursor() as cur:
        cur.execute("INSERT INTO la_import_log (fuente, form_id, leidos, nuevos, duplicados, observados, usuario) "
                    "VALUES (%s,%s,%s,%s,%s,%s,%s) RETURNING id",
                    (fuente, form_id, len(resultados), sum(r.estado_import == "NUEVO" for r in resultados),
                     sum(r.estado_import == "DUPLICADO" for r in resultados),
                     sum(r.estado_import == "OBSERVADO" for r in resultados), usuario))
        import_id = cur.fetchone()[0]
        for r in nuevos:
            d = r.datos
            cur.execute("""INSERT INTO la_envios_raw (kobo_uuid, form_id, cod_unidad, cod_predio, asistente, fecha_envio,
                           estado_import, motivos, payload, import_id, origen) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                           ON CONFLICT (kobo_uuid) DO NOTHING""",
                        (r.kobo_uuid, form_id, _s(r.cod_unidad), _s(r.cod_predio), d.get("asistente"),
                         _fecha_envio(d), r.estado_import, " | ".join(r.motivos),
                         json.dumps(d, ensure_ascii=False, default=str), import_id, origen))
            if cur.rowcount == 0:
                continue
            _insertar_formulario(cur, form_id, r)
    conn.commit()
    recalcular_estados(conn)
    return {"import_id": import_id, "insertados": len(nuevos)}


def _fecha_envio(d: dict):
    """Kobo trae _submission_time / fin; en el aplicativo y la plantilla se usa la fecha de registro."""
    return d.get("_submission_time") or d.get("fin") or d.get("hoy") or None


def _insertar_formulario(cur, form_id: str, r: core.Resultado, actualizar_titular: bool = False):
    d, u = r.datos, r.kobo_uuid
    fecha = (d.get("hoy") or d.get("fecha_evento") or d.get("fecha_insp") or d.get("fecha_acta"))
    if form_id == "f_la_01_reunion":
        cur.execute("""INSERT INTO la_reuniones VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                       ON CONFLICT DO NOTHING""",
                    (u, d.get("fecha_evento") or fecha, d.get("asistente"), d.get("distrito"), d.get("centro_poblado"),
                     d.get("comunidad"), d.get("tipo_evento"), d.get("unidades"), d.get("asist_hombres"), d.get("asist_mujeres"),
                     d.get("titulares_presentes"), d.get("aceptacion"), d.get("alertas"), d.get("acuerdos"),
                     d.get("a01_suscrita"), d.get("a02_suscrita"), d.get("este"), d.get("norte")))
    elif form_id == "f_la_02_titular":
        cod_tit = _registrar_titular(cur, d, actualizar=actualizar_titular)
        _asegurar_predio(cur, d)
        if cod_tit and d.get("cod_predio"):
            cur.execute("INSERT INTO la_predio_titular VALUES (%s,%s,'titular') ON CONFLICT DO NOTHING", (d["cod_predio"], cod_tit))
        cur.execute("INSERT INTO la_fichas_titular VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING",
                    (u, fecha, d.get("asistente"), d.get("cod_unidad"), d.get("cod_predio"), cod_tit, d.get("consentimiento"),
                     d.get("tipo_titularidad"), d.get("docs_exhibidos"), d.get("aceptacion"), d.get("estado_la_propuesto")))
    elif form_id == "f_la_03_inspeccion":
        _asegurar_predio(cur, d)
        cur.execute("""INSERT INTO la_inspecciones VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                       ON CONFLICT DO NOTHING""",
                    (u, d.get("fecha_insp") or fecha, d.get("asistente"), d.get("cod_unidad"), d.get("cod_predio"),
                     d.get("este"), d.get("norte"), d.get("precision_m"), d.get("validacion_espacial"), d.get("distancia_m"),
                     d.get("puntos_fuera"), d.get("interferencias"), d.get("conclusion_campo"), d.get("area_sus_ha"),
                     d.get("posicion_sus_calc"), d.get("distancia_bloque_m"), d.get("geom_wkt")))
    elif form_id == "f_la_04_actas":
        if (d.get("n_predio") not in (0, "0")):
            _asegurar_predio(cur, d)
        conforme = r.estado_import == "NUEVO" and d.get("estado_acta") == "completa"
        cur.execute("""INSERT INTO la_actas (kobo_uuid, cod_doc, tipo_acta, fecha, asistente, cod_unidad, cod_predio,
                       area_comprometida_ha, plazo, n_firmantes, fedatario_tipo, quorum_pct, checklist, estado_acta, conforme,
                       fecha_entrega_cd) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING""",
                    (u, d.get("cod_doc"), d.get("tipo_acta"), d.get("fecha_acta"), d.get("asistente"), d.get("cod_unidad"),
                     d.get("cod_predio"), core._num(d.get("area_comprometida_ha")), d.get("plazo_consignado"),
                     len(d.get("r_firmantes") or []), d.get("fedatario_tipo"), core._num(d.get("quorum_pct")),
                     d.get("checklist"), d.get("estado_acta"), conforme, d.get("fecha_entrega_cd") or None))
    elif form_id == "f_la_06_vivero":
        cur.execute("""INSERT INTO la_vivero_alternativas VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                       ON CONFLICT DO NOTHING""",
                    (u, d.get("cod_vivero"), fecha, d.get("distrito"), d.get("alt_nombre"), d.get("modalidad"),
                     d.get("tipo_titularidad"), d.get("titular_nombre"), core._num(d.get("area_ha")), core._num(d.get("altitud")),
                     d.get("pendiente"), d.get("agua_fuente"), core._num(d.get("agua_caudal_ls")), d.get("acceso_tipo"),
                     d.get("acceso_camion"), d.get("energia"), core._num(d.get("dist_bloques_km")), d.get("inundabilidad"),
                     d.get("deslizamiento"), d.get("disposicion"), d.get("firmaria_a07"), core.puntaje_vivero(d),
                     d.get("este"), d.get("norte")))


def _asegurar_predio(cur, d: dict):
    if not d.get("cod_predio") or d.get("tipo_unidad") == "vivero":
        return
    cur.execute("""INSERT INTO la_predios (cod_predio, cod_unidad, n_predio, nombre_predio, area_decl_ha, area_unidad_ha,
                   uso_actual, ocupacion, aceptacion, alertas) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                   ON CONFLICT (cod_predio) DO UPDATE SET
                     nombre_predio = COALESCE(EXCLUDED.nombre_predio, la_predios.nombre_predio),
                     area_decl_ha = COALESCE(EXCLUDED.area_decl_ha, la_predios.area_decl_ha),
                     area_unidad_ha = COALESCE(EXCLUDED.area_unidad_ha, la_predios.area_unidad_ha),
                     uso_actual = COALESCE(EXCLUDED.uso_actual, la_predios.uso_actual),
                     ocupacion = COALESCE(EXCLUDED.ocupacion, la_predios.ocupacion),
                     aceptacion = COALESCE(EXCLUDED.aceptacion, la_predios.aceptacion),
                     alertas = COALESCE(EXCLUDED.alertas, la_predios.alertas), actualizado = now()""",
                (d["cod_predio"], d.get("cod_unidad"), int(d.get("n_predio") or 0), d.get("predio_nombre"),
                 core._num(d.get("predio_area_decl_ha")), core._num(d.get("predio_area_unidad_ha")),
                 d.get("uso_actual") or d.get("uso_observado"), d.get("ocupacion"), d.get("aceptacion"),
                 d.get("conflicto_linderos_det") or d.get("interf_detalle")))


def _registrar_titular(cur, d: dict, actualizar: bool = False) -> str | None:
    """Deduplica por DNI/RUC; asigna T{nnnn} correlativo si es nuevo.
    Con `actualizar` (edición de una ficha F-LA-02) corrige los datos del titular ya existente."""
    if d.get("consentimiento") != "si":
        return None
    t = d.get("tipo_titularidad")
    if t == "comunal":
        nombre, doc = d.get("cc_nombre"), d.get("cc_ruc")
        rep, rep_dni = d.get("cc_presidente"), d.get("cc_presidente_dni")
    elif t == "estatal":
        nombre, doc, rep, rep_dni = d.get("est_entidad_otra") or d.get("est_entidad"), None, d.get("est_contacto"), None
    elif t == "sin_titular":
        return None
    else:
        nombre = f"{d.get('tit_nombres', '')} {d.get('tit_apellidos', '')}".strip()
        doc, rep, rep_dni = d.get("tit_dni"), None, None
    if d.get("cod_titular_prev"):
        return d["cod_titular_prev"]
    if doc:
        cur.execute("SELECT cod_titular FROM la_titulares WHERE dni_ruc=%s", (doc,))
        x = cur.fetchone()
        if x:
            if actualizar:
                cur.execute("""UPDATE la_titulares SET tipo_titularidad=%s, nombre=%s, celular=COALESCE(%s, celular),
                               representante=COALESCE(%s, representante), representante_dni=COALESCE(%s, representante_dni),
                               conyuge=COALESCE(%s, conyuge), conyuge_dni=COALESCE(%s, conyuge_dni),
                               partida=COALESCE(%s, partida) WHERE cod_titular=%s""",
                            (t, nombre or "[DATO FALTANTE]", d.get("tit_celular") or d.get("cc_presidente_cel"), rep, rep_dni,
                             d.get("cony_nombre"), d.get("cony_dni"), d.get("doc_partida") or d.get("cc_partida"), x[0]))
            return x[0]
    cur.execute("SELECT cod_titular FROM la_titulares")
    cod = core.siguiente_titular([x[0] for x in cur.fetchall()])
    cur.execute("""INSERT INTO la_titulares (cod_titular, tipo_titularidad, nombre, dni_ruc, celular, representante,
                   representante_dni, conyuge, conyuge_dni, partida) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                (cod, t, nombre or "[DATO FALTANTE]", doc, d.get("tit_celular") or d.get("cc_presidente_cel"), rep, rep_dni,
                 d.get("cony_nombre"), d.get("cony_dni"), d.get("doc_partida") or d.get("cc_partida")))
    return cod


# ------------------------------------------------------------------ estados
def recalcular_estados(conn):
    """Recalcula estado_la y clasificación de cada predio a partir de los formularios importados."""
    pred = df(conn, "SELECT * FROM la_predios")
    if pred.empty:
        return
    tit = df(conn, "SELECT DISTINCT cod_predio FROM la_predio_titular")
    fich = df(conn, "SELECT cod_predio, consentimiento, tipo_titularidad, aceptacion FROM la_fichas_titular")
    reun = df(conn, "SELECT unidades FROM la_reuniones")
    insp = df(conn, "SELECT cod_predio, validacion_espacial FROM la_inspecciones")
    actas = df(conn, "SELECT cod_unidad, cod_predio, tipo_acta, conforme FROM la_actas")
    raw = df(conn, "SELECT cod_predio, motivos FROM la_envios_raw WHERE estado_import='OBSERVADO'")
    socializadas = set()
    for u in (reun["unidades"] if not reun.empty else []):
        socializadas.update(str(u or "").split())
    con_tit = set(tit["cod_predio"]) if not tit.empty else set()
    filas = []
    for p in pred.to_dict("records"):
        cp, cu = p["cod_predio"], p["cod_unidad"]
        f = fich[fich["cod_predio"] == cp] if not fich.empty else fich
        i = insp[insp["cod_predio"] == cp] if not insp.empty else insp
        a = actas[((actas["cod_predio"] == cp) & actas["tipo_acta"].isin(["A-04"])) |
                  ((actas["cod_unidad"] == cu) & (actas["tipo_acta"] == "A-03"))] if not actas.empty else actas
        motivos = list(raw[raw["cod_predio"] == cp]["motivos"]) if not raw.empty else []
        exc = p.get("excepcion_manual")
        if not exc and not f.empty and (f["aceptacion"] == "rechazo").any():
            exc = "NEG"
        if not exc and not f.empty and f["tipo_titularidad"].isin(["sin_titular"]).any():
            exc = "OBS"
        ev = {"excepcion": exc, "titular": cp in con_tit,
              "socializado": cu in socializadas or (not f.empty and (f["consentimiento"] == "si").any()),
              "inspeccion_dentro": not i.empty and (i["validacion_espacial"] == "DENTRO").any(),
              "acta_conforme": not a.empty and bool(a["conforme"].fillna(False).any()),
              "docs_completos": bool(p.get("docs_completos")), "expediente_conforme": bool(p.get("expediente_conforme"))}
        est = core.estado_predio(ev)
        filas.append((est, core.clasificacion_matriz(est, motivos), cp))
    with conn.cursor() as cur:
        psycopg2.extras.execute_batch(cur, "UPDATE la_predios SET estado_la=%s, clasificacion=%s, actualizado=now() WHERE cod_predio=%s", filas)
    conn.commit()


def registrar_documento(conn, datos: dict) -> str:
    """F-LA-05: registra una constancia de búsqueda / documento. Código CBU-{UNIDAD}-P{nn}-{n}."""
    with conn.cursor() as cur:
        base = f"CBU-{datos.get('cod_predio') or datos['cod_unidad'] + '-P00'}"
        # correlativo sobre los códigos vigentes y los eliminados (bitácora): un código nunca se reutiliza
        cur.execute("""SELECT cod_doc FROM la_documentos WHERE cod_doc LIKE %s
                       UNION SELECT clave FROM la_bitacora WHERE tabla = 'la_documentos' AND clave LIKE %s""",
                    (base + "-%", base + "-%"))
        usados = [int(x[0].rsplit("-", 1)[-1]) for x in cur.fetchall() if x[0].rsplit("-", 1)[-1].isdigit()]
        cod = f"{base}-{max(usados, default=0) + 1}"
        cur.execute("""INSERT INTO la_documentos (cod_doc, cod_unidad, cod_predio, tipo, entidad, fecha, resultado, n_partida,
                       descripcion, archivo_url, registrado_por) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                    (cod, datos["cod_unidad"], datos.get("cod_predio"), datos["tipo"], datos.get("entidad"), datos.get("fecha"),
                     datos.get("resultado"), datos.get("n_partida"), datos.get("descripcion"), datos.get("archivo_url"),
                     datos.get("registrado_por")))
    conn.commit()
    return cod


# ------------------------------------------------------------------ consulta, edición, eliminación y restauración
# Toda edición o eliminación guarda antes una copia completa en la_bitacora (permite restaurar).
# Las fotos de un envío eliminado NO se borran: la llave la_adjuntos.kobo_uuid queda en NULL
# (ON DELETE SET NULL) y se vuelven a enlazar si el envío se restaura.

def _json(v) -> str:
    return json.dumps(v, ensure_ascii=False, default=str)


def _bitacora(cur, usuario, accion, tabla, clave, form_id, datos):
    cur.execute("INSERT INTO la_bitacora (usuario, accion, tabla, clave, form_id, datos) VALUES (%s,%s,%s,%s,%s,%s)",
                (usuario, accion, tabla, clave, form_id, _json(datos)))


def listar_envios(conn, form_id: str | None = None) -> pd.DataFrame:
    sql = """SELECT kobo_uuid, form_id, cod_unidad, cod_predio, asistente, fecha_envio, estado_import, motivos,
                    COALESCE(origen, 'KOBO') AS origen, creado, editado, editado_por
             FROM la_envios_raw {w} ORDER BY creado DESC, kobo_uuid"""
    if form_id:
        return df(conn, sql.format(w="WHERE form_id = %s"), (form_id,))
    return df(conn, sql.format(w=""))


def obtener_envio(conn, kobo_uuid: str) -> dict | None:
    d = df(conn, "SELECT * FROM la_envios_raw WHERE kobo_uuid = %s", (kobo_uuid,))
    if d.empty:
        return None
    r = d.to_dict("records")[0]
    if isinstance(r.get("payload"), str):
        r["payload"] = json.loads(r["payload"])
    return r


def fila_formulario(conn, form_id: str, kobo_uuid: str) -> dict:
    tabla = core.FORMULARIOS[form_id]["tabla"]
    d = df(conn, f"SELECT * FROM {tabla} WHERE kobo_uuid = %s", (kobo_uuid,))
    return d.to_dict("records")[0] if not d.empty else {}


def adjuntos_envio(conn, kobo_uuid: str) -> pd.DataFrame:
    return df(conn, "SELECT id, campo, nombre_archivo, mimetype, url_original, contenido FROM la_adjuntos "
                    "WHERE kobo_uuid = %s ORDER BY nombre_archivo", (kobo_uuid,))


def _quitar_formulario(cur, form_id: str, kobo_uuid: str) -> dict:
    """Borra la fila del formulario (no el envío crudo). En F-LA-02 retira además el enlace predio–titular si
    ninguna otra ficha lo sostiene. Devuelve la fila borrada."""
    tabla = core.FORMULARIOS[form_id]["tabla"]
    cur.execute(f"SELECT row_to_json(t) FROM {tabla} t WHERE kobo_uuid = %s", (kobo_uuid,))
    x = cur.fetchone()
    fila = x[0] if x else {}
    cur.execute(f"DELETE FROM {tabla} WHERE kobo_uuid = %s", (kobo_uuid,))
    if form_id == "f_la_02_titular" and fila.get("cod_predio") and fila.get("cod_titular"):
        cur.execute("""DELETE FROM la_predio_titular pt WHERE pt.cod_predio = %s AND pt.cod_titular = %s AND pt.rol = 'titular'
                       AND NOT EXISTS (SELECT 1 FROM la_fichas_titular f WHERE f.cod_predio = pt.cod_predio
                                       AND f.cod_titular = pt.cod_titular)""",
                    (fila["cod_predio"], fila["cod_titular"]))
    return fila


def actualizar_envio(conn, r: core.Resultado, form_id: str, usuario: str) -> None:
    """Reemplaza un envío ya registrado (de Kobo, del aplicativo o de la plantilla) por su versión corregida,
    conservando el mismo kobo_uuid, sus adjuntos y una copia previa en la bitácora."""
    previo = obtener_envio(conn, r.kobo_uuid)
    if previo is None:
        raise ValueError(f"El envío {r.kobo_uuid} ya no existe.")
    if previo["form_id"] != form_id:
        raise ValueError("El formulario del envío no coincide.")
    d = r.datos
    with conn.cursor() as cur:
        fila = _quitar_formulario(cur, form_id, r.kobo_uuid)
        _bitacora(cur, usuario, "EDITAR", "la_envios_raw", r.kobo_uuid, form_id,
                  {"envio": {k: v for k, v in previo.items() if k != "payload"}, "payload": previo["payload"], "fila": fila})
        cur.execute("""UPDATE la_envios_raw SET cod_unidad=%s, cod_predio=%s, asistente=%s, estado_import=%s, motivos=%s,
                       payload=%s, editado=now(), editado_por=%s WHERE kobo_uuid=%s""",
                    (_s(r.cod_unidad), _s(r.cod_predio), d.get("asistente"), r.estado_import, " | ".join(r.motivos),
                     _json(d), usuario, r.kobo_uuid))
        _insertar_formulario(cur, form_id, r, actualizar_titular=True)
        if form_id == "f_la_04_actas" and fila.get("recibido_cd"):          # control documentario ya registrado
            cur.execute("UPDATE la_actas SET recibido_cd = TRUE WHERE kobo_uuid = %s", (r.kobo_uuid,))
    conn.commit()
    recalcular_estados(conn)


def eliminar_envio(conn, kobo_uuid: str, usuario: str) -> None:
    previo = obtener_envio(conn, kobo_uuid)
    if previo is None:
        return
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM la_adjuntos WHERE kobo_uuid = %s", (kobo_uuid,))
        adj = [x[0] for x in cur.fetchall()]
        fila = _quitar_formulario(cur, previo["form_id"], kobo_uuid)
        _bitacora(cur, usuario, "ELIMINAR", "la_envios_raw", kobo_uuid, previo["form_id"],
                  {"envio": {k: v for k, v in previo.items() if k != "payload"}, "payload": previo["payload"],
                   "fila": fila, "adjuntos": adj})
        cur.execute("DELETE FROM la_envios_raw WHERE kobo_uuid = %s", (kobo_uuid,))   # adjuntos → kobo_uuid NULL
    conn.commit()
    recalcular_estados(conn)


def bitacora(conn, solo_eliminados: bool = False, limite: int = 500) -> pd.DataFrame:
    w = "WHERE accion = 'ELIMINAR' AND NOT restaurado" if solo_eliminados else ""
    return df(conn, f"""SELECT id, fecha, usuario, accion, tabla, clave, form_id, restaurado FROM la_bitacora {w}
                        ORDER BY id DESC LIMIT %s""", (limite,))


def restaurar(conn, id_bitacora: int, usuario: str) -> str:
    """Vuelve a registrar un envío / documento / predio eliminado a partir de su copia en la bitácora.
    Los envíos pasan otra vez por core.validar_envio con el catálogo y geometrías vigentes."""
    b = df(conn, "SELECT * FROM la_bitacora WHERE id = %s", (id_bitacora,))
    if b.empty:
        raise ValueError("Registro de bitácora inexistente.")
    b = b.to_dict("records")[0]
    datos = b["datos"] if isinstance(b["datos"], dict) else json.loads(b["datos"])
    if b["accion"] != "ELIMINAR" or b["restaurado"]:
        raise ValueError("Solo se restauran eliminaciones no restauradas.")
    if b["tabla"] == "la_envios_raw":
        if obtener_envio(conn, b["clave"]):
            raise ValueError(f"El envío {b['clave']} ya existe.")
        cat = catalogo(conn)
        res = core.validar_envio(b["form_id"], datos["payload"], cat, geometrias(conn), uuids_existentes(conn),
                                 sus_areas(conn, cat))
        res.kobo_uuid = b["clave"]
        importar(conn, [res], "RESTAURACION", b["form_id"], usuario, origen=(datos.get("envio") or {}).get("origen"))
        with conn.cursor() as cur:
            if datos.get("adjuntos"):
                cur.execute("UPDATE la_adjuntos SET kobo_uuid = %s WHERE kobo_uuid IS NULL AND id = ANY(%s)",
                            (b["clave"], datos["adjuntos"]))
    elif b["tabla"] == "la_documentos":
        f = datos["fila"]
        with conn.cursor() as cur:
            cur.execute("""INSERT INTO la_documentos (cod_doc, cod_unidad, cod_predio, tipo, entidad, fecha, resultado,
                           n_partida, descripcion, archivo_url, registrado_por) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                        tuple(f.get(k) for k in ("cod_doc", "cod_unidad", "cod_predio", "tipo", "entidad", "fecha",
                                                 "resultado", "n_partida", "descripcion", "archivo_url", "registrado_por")))
    elif b["tabla"] == "la_predios":
        f = datos["fila"]
        cols = [k for k in f if k != "actualizado"]
        with conn.cursor() as cur:
            cur.execute(f"INSERT INTO la_predios ({', '.join(cols)}) VALUES ({', '.join(['%s'] * len(cols))}) "
                        "ON CONFLICT (cod_predio) DO NOTHING", tuple(f[k] for k in cols))
            for t in datos.get("titulares", []):
                cur.execute("INSERT INTO la_predio_titular VALUES (%s,%s,%s) ON CONFLICT DO NOTHING",
                            (f["cod_predio"], t["cod_titular"], t["rol"]))
    with conn.cursor() as cur:
        cur.execute("UPDATE la_bitacora SET restaurado = TRUE WHERE id = %s", (id_bitacora,))
        _bitacora(cur, usuario, "RESTAURAR", b["tabla"], b["clave"], b["form_id"], {"id_bitacora": id_bitacora})
    conn.commit()
    recalcular_estados(conn)
    return b["clave"]


# ------------------------------------------------------------------ F-LA-05 documentos
COLS_DOC = ("cod_unidad", "cod_predio", "tipo", "entidad", "fecha", "resultado", "n_partida", "descripcion",
            "archivo_url", "registrado_por")


def documento_existe(conn, datos: dict) -> str | None:
    """Mismo documento ya registrado (unidad, predio, tipo, entidad, fecha y n.° de partida) → su código."""
    d = df(conn, """SELECT cod_doc FROM la_documentos WHERE cod_unidad = %s AND cod_predio IS NOT DISTINCT FROM %s
                    AND tipo = %s AND entidad IS NOT DISTINCT FROM %s AND fecha IS NOT DISTINCT FROM %s
                    AND COALESCE(n_partida, '') = COALESCE(%s, '')""",
           (datos.get("cod_unidad"), datos.get("cod_predio"), datos.get("tipo"), datos.get("entidad"), datos.get("fecha"),
            datos.get("n_partida")))
    return None if d.empty else d["cod_doc"][0]


def obtener_documento(conn, id_doc: int) -> dict | None:
    d = df(conn, "SELECT * FROM la_documentos WHERE id = %s", (id_doc,))
    return None if d.empty else d.to_dict("records")[0]


def actualizar_documento(conn, id_doc: int, datos: dict, usuario: str) -> None:
    previo = obtener_documento(conn, id_doc)
    if previo is None:
        raise ValueError("El documento ya no existe.")
    with conn.cursor() as cur:
        _bitacora(cur, usuario, "EDITAR", "la_documentos", previo["cod_doc"], "f_la_05_documentos", {"fila": previo})
        cur.execute(f"UPDATE la_documentos SET {', '.join(f'{c}=%s' for c in COLS_DOC)}, editado=now(), editado_por=%s "
                    "WHERE id=%s", tuple(datos.get(c) for c in COLS_DOC) + (usuario, id_doc))
    conn.commit()


def eliminar_documento(conn, id_doc: int, usuario: str) -> None:
    previo = obtener_documento(conn, id_doc)
    if previo is None:
        return
    with conn.cursor() as cur:
        _bitacora(cur, usuario, "ELIMINAR", "la_documentos", previo["cod_doc"], "f_la_05_documentos", {"fila": previo})
        cur.execute("DELETE FROM la_documentos WHERE id = %s", (id_doc,))
    conn.commit()


# ------------------------------------------------------------------ predios (matriz)
CAMPOS_PREDIO_EDITABLES = ("nombre_predio", "area_decl_ha", "area_unidad_ha", "uso_actual", "ocupacion", "nucleo",
                           "docs_completos", "expediente_conforme", "excepcion_manual", "alertas")


def actualizar_predios(conn, cambios: list[dict], usuario: str) -> int:
    """cambios = [{cod_predio, campo: valor…}] con campos de CAMPOS_PREDIO_EDITABLES. Guarda solo lo que cambió."""
    n = 0
    with conn.cursor() as cur:
        for c in cambios:
            cur.execute("SELECT row_to_json(p) FROM la_predios p WHERE cod_predio = %s", (c["cod_predio"],))
            x = cur.fetchone()
            if not x:
                continue
            previo = x[0]
            nuevos = {k: v for k, v in c.items() if k in CAMPOS_PREDIO_EDITABLES and _distinto(previo.get(k), v)}
            if not nuevos:
                continue
            _bitacora(cur, usuario, "EDITAR", "la_predios", c["cod_predio"], None, {"fila": previo, "cambios": nuevos})
            cur.execute(f"UPDATE la_predios SET {', '.join(f'{k}=%s' for k in nuevos)}, actualizado=now() WHERE cod_predio=%s",
                        tuple(nuevos.values()) + (c["cod_predio"],))
            n += 1
    conn.commit()
    recalcular_estados(conn)
    return n


def _distinto(a, b) -> bool:
    if a in (None, "") and b in (None, ""):
        return False
    try:
        return abs(float(a) - float(b)) > 1e-9
    except (TypeError, ValueError):
        return str(a) != str(b)


def referencias_predio(conn, cod_predio: str) -> dict[str, int]:
    out = {}
    for t in ("la_fichas_titular", "la_inspecciones", "la_actas", "la_documentos"):
        n = df(conn, f"SELECT count(*) AS n FROM {t} WHERE cod_predio = %s", (cod_predio,))["n"][0]
        if n:
            out[t] = int(n)
    return out


def eliminar_predio(conn, cod_predio: str, usuario: str) -> None:
    """Solo si ningún formulario ni documento lo referencia (primero se eliminan / corrigen esos registros)."""
    refs = referencias_predio(conn, cod_predio)
    if refs:
        raise ValueError("El predio tiene registros asociados: " + ", ".join(f"{k} ({v})" for k, v in refs.items()))
    with conn.cursor() as cur:
        cur.execute("SELECT row_to_json(p) FROM la_predios p WHERE cod_predio = %s", (cod_predio,))
        x = cur.fetchone()
        if not x:
            return
        cur.execute("SELECT cod_titular, rol FROM la_predio_titular WHERE cod_predio = %s", (cod_predio,))
        tit = [{"cod_titular": a, "rol": b} for a, b in cur.fetchall()]
        _bitacora(cur, usuario, "ELIMINAR", "la_predios", cod_predio, None, {"fila": x[0], "titulares": tit})
        cur.execute("DELETE FROM la_predios WHERE cod_predio = %s", (cod_predio,))
    conn.commit()
