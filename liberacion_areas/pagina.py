# -*- coding: utf-8 -*-
"""
Página «Liberación de Áreas» del aplicativo IN Piura (CUI 2669244).
Se integra en el menú de streamlit_app.py:
    from liberacion_areas.pagina import render as pagina_liberacion_areas
    ...
    elif pagina == "Liberacion de Areas": pagina_liberacion_areas()
Usa los mismos secrets del aplicativo: DATABASE_URL, KOBO_TOKEN y (opcional) KOBO_SERVER.
Las tablas la_* se crean solas la primera vez (migración aditiva) y se enlazan con `bloques`.
"""
from __future__ import annotations

import json
from datetime import date

import pandas as pd
import streamlit as st

from . import la_core as core
from . import la_db as db
from . import la_export as ex
from . import la_kobo as kb

ESTILO_ESTADO = {"NUEVO": "background-color:#C6EFCE", "DUPLICADO": "background-color:#D9D9D9",
                 "OBSERVADO": "background-color:#FFEB9C"}


@st.cache_resource(show_spinner="Preparando tablas de liberación de áreas…")
def _inicializar_una_vez():
    """Crea/actualiza las tablas la_* y las enlaza con `bloques` una sola vez por proceso."""
    conn = db.conectar()
    try:
        return db.inicializar_la(conn)
    finally:
        conn.close()


def render():
    st.header("Liberación de Áreas – Tamizaje predial")
    st.caption("Expediente preliminar de liberación de áreas · Bloques V6, lotes SUS y Vivero Central · UTM WGS84 17S")
    try:
        _inicializar_una_vez()
        conn = db.conectar()           # una conexión por ejecución, como el resto del aplicativo
    except Exception as e:  # noqa: BLE001
        st.error(f"No se pudo conectar a la base de datos: {e}")
        return
    try:
        t = st.tabs(["1 · Importar Kobo", "2 · F-LA-05 Documentos", "3 · Matriz predial", "4 · Avance y semáforo",
                     "5 · Vivero Central", "6 · Catálogo y enlace con bloques"])
        with t[0]:
            _tab_importar(conn)
        with t[1]:
            _tab_documentos(conn)
        with t[2]:
            _tab_matriz(conn)
        with t[3]:
            _tab_avance(conn)
        with t[4]:
            _tab_vivero(conn)
        with t[5]:
            _tab_catalogo(conn)
    finally:
        conn.close()


# ------------------------------------------------------------------ 1. importar
def _tab_importar(conn):
    st.subheader("Importar envíos de KoboToolbox")
    fuente = st.radio("Fuente", ["API KoboToolbox", "Archivo exportado (XLSX / JSON)"], horizontal=True)
    form_id = st.selectbox("Formulario", list(core.FORMULARIOS),
                           format_func=lambda f: f"{core.FORMULARIOS[f]['codigo']} – {core.FORMULARIOS[f]['nombre']}")
    registros: list[dict] = []
    kobo = None
    if fuente.startswith("API"):
        try:
            kobo = kb.cliente()
        except ValueError as e:
            st.warning(f"{e} Es el mismo token que usa la página «ODK / KoBoToolbox».")
            return
        if st.button("Listar formularios F-LA"):
            with st.spinner("Consultando KoboToolbox…"):
                st.session_state["la_forms"] = kb.formularios_la(kobo)
        forms = st.session_state.get("la_forms", [])
        uid = st.selectbox("Formulario en KoboToolbox", [f["uid"] for f in forms],
                           format_func=lambda u: next((f"{f['nombre']} · {f['envios']} envíos" for f in forms if f["uid"] == u), u)) if forms else \
            st.text_input("UID del formulario (asset uid)", help="Aparece en la URL del formulario: /#/forms/<uid>/…")
        if st.button("Descargar envíos", disabled=not uid):
            with st.spinner("Descargando todas las páginas de envíos…"):
                st.session_state["la_raw"] = kobo.obtener_envios(uid)
    else:
        arch = st.file_uploader("Exportación de Kobo (XLSX con nombres XML, o JSON)", type=["xlsx", "json"])
        if arch is not None:
            st.session_state["la_raw"] = (json.load(arch) if arch.name.endswith(".json") else kb.leer_exportacion_xlsx(arch))
            if isinstance(st.session_state["la_raw"], dict):
                st.session_state["la_raw"] = st.session_state["la_raw"].get("results", [])
    registros = st.session_state.get("la_raw", [])
    if not registros:
        st.info("Sin envíos cargados.")
        return

    cat, geoms = db.catalogo(conn), db.geometrias(conn)
    existentes = db.uuids_existentes(conn)
    sus = db.sus_areas(conn, cat)
    resultados = [core.validar_envio(form_id, core.aplanar(r), cat, geoms, existentes, sus) for r in registros]
    vista = pd.DataFrame([{"estado_import": r.estado_import, "motivo": " | ".join(r.motivos), "kobo_uuid": r.kobo_uuid,
                           "asistente": r.datos.get("asistente"), "unidad": r.cod_unidad, "predio": r.cod_predio,
                           "fecha": r.datos.get("hoy"), "este": r.datos.get("este"), "norte": r.datos.get("norte"),
                           "validación espacial": r.datos.get("validacion_espacial"),
                           "SUS (ha)": r.datos.get("area_sus_ha"), "posición SUS": r.datos.get("posicion_sus_calc")}
                          for r in resultados])
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Leídos", len(vista))
    c2.metric("Nuevos", int((vista.estado_import == "NUEVO").sum()))
    c3.metric("Duplicados", int((vista.estado_import == "DUPLICADO").sum()))
    c4.metric("Observados", int((vista.estado_import == "OBSERVADO").sum()))
    st.dataframe(vista.style.map(lambda v: ESTILO_ESTADO.get(v, ""), subset=["estado_import"]), width="stretch", hide_index=True)
    st.caption("Los OBSERVADOS se cargan con su motivo para que el Especialista Predial los resuelva; los DUPLICADOS se ignoran.")
    fotos = st.checkbox("Descargar fotos y páginas de actas (versión mediana, se guardan en la base)", value=True,
                        disabled=kobo is None)
    usuario = st.text_input("Usuario que confirma la carga", value=st.session_state.get("usuario", ""))
    if st.button("✅ Confirmar importación", type="primary", disabled=not usuario):
        res = db.importar(conn, resultados, "API" if kobo else "ARCHIVO", form_id, usuario)
        if fotos and kobo:
            n = 0
            with st.spinner("Descargando adjuntos…"):
                for r in resultados:
                    if r.estado_import == "DUPLICADO":
                        continue
                    for a in kb.descargar_adjuntos(kobo, r.datos):
                        db.guardar_adjunto(conn, r.kobo_uuid, r.cod_predio, a["campo"], a["nombre"], a["mimetype"],
                                           a["url"], a["contenido"])
                        n += 1
            st.info(f"Adjuntos guardados: {n}")
        st.success(f"Importación N.° {res['import_id']}: {res['insertados']} envíos registrados. Estados LA recalculados.")
        st.session_state.pop("la_raw", None)


# ------------------------------------------------------------------ 2. F-LA-05
def _tab_documentos(conn):
    st.subheader("F-LA-05 · Constancia de búsqueda documental y documentos de titularidad")
    cat = db.catalogo(conn)
    if not cat:
        st.info("Cargue primero el catálogo de unidades (pestaña 6).")
        return
    pred = db.df(conn, "SELECT cod_predio, cod_unidad FROM la_predios ORDER BY cod_predio")
    with st.form("f_la_05", clear_on_submit=True):
        c1, c2 = st.columns(2)
        unidad = c1.selectbox("Unidad", sorted(cat))
        opciones = ["(toda la unidad)"] + (list(pred[pred.cod_unidad == unidad].cod_predio) if not pred.empty else [])
        predio = c2.selectbox("Predio", opciones)
        tipo = c1.selectbox("Tipo de documento", ["constancia_busqueda", "certificado_busqueda_catastral", "partida_registral",
                                                 "titulo", "vigencia_poder", "acta_asamblea_certificada", "constancia_posesion",
                                                 "sucesion_intestada", "consulta_sinabip", "dni", "otro"])
        entidad = c2.selectbox("Entidad", ["SUNARP", "SBN", "GORE / DRA Piura", "COFOPRI", "Municipalidad", "SERFOR / ATFFS",
                                           "Min. Cultura", "INGEMMET", "Juzgado de paz", "Otra"])
        fecha = c1.date_input("Fecha", value=date.today())
        resultado = c2.selectbox("Resultado", ["positivo", "negativo", "en_tramite"])
        partida = c1.text_input("N.° de partida / expediente")
        url = c2.text_input("Enlace al archivo (Storage / Drive)")
        desc = st.text_area("Descripción / hallazgos")
        por = st.text_input("Registrado por")
        if st.form_submit_button("Registrar"):
            cod = db.registrar_documento(conn, {"cod_unidad": unidad, "cod_predio": None if predio.startswith("(") else predio,
                                                "tipo": tipo, "entidad": entidad, "fecha": fecha, "resultado": resultado,
                                                "n_partida": partida, "descripcion": desc, "archivo_url": url, "registrado_por": por})
            st.success(f"Registrado con código {cod}")
    st.dataframe(db.df(conn, "SELECT cod_doc, cod_unidad, cod_predio, tipo, entidad, fecha, resultado, n_partida FROM la_documentos ORDER BY id DESC"),
                 width="stretch", hide_index=True)


# ------------------------------------------------------------------ 3. matriz
SQL_MATRIZ = """
SELECT p.cod_predio, p.cod_unidad, u.tipo_unidad, u.provincia, u.distrito, u.asistente, p.nombre_predio,
       string_agg(DISTINCT t.cod_titular || ' ' || t.nombre, '; ') AS titulares,
       string_agg(DISTINCT t.tipo_titularidad, '; ') AS tipo_titularidad,
       p.area_unidad_ha, p.uso_actual, p.aceptacion, p.estado_la, p.clasificacion, p.excepcion_manual,
       p.docs_completos, p.expediente_conforme, p.alertas
FROM la_predios p JOIN la_unidades u ON u.codigo = p.cod_unidad
LEFT JOIN la_predio_titular pt ON pt.cod_predio = p.cod_predio
LEFT JOIN la_titulares t ON t.cod_titular = pt.cod_titular
GROUP BY p.cod_predio, u.codigo ORDER BY u.provincia, u.distrito, p.cod_predio"""


def _tab_matriz(conn):
    st.subheader("Matriz predial")
    m = db.df(conn, SQL_MATRIZ)
    if m.empty:
        st.info("Aún no hay predios registrados.")
        return
    c1, c2, c3 = st.columns(3)
    fa = c1.multiselect("Asistente", sorted(m.asistente.dropna().unique()))
    fd = c2.multiselect("Distrito", sorted(m.distrito.unique()))
    fe = c3.multiselect("Estado LA", core.ESTADOS_LA + core.ESTADOS_EXCEPCION)
    mask = pd.Series(True, index=m.index)
    if fa:
        mask &= m.asistente.isin(fa)
    if fd:
        mask &= m.distrito.isin(fd)
    if fe:
        mask &= m.estado_la.isin(fe)
    v = m[mask]
    st.caption("Especialista Predial: marque LA-5 (documentación completa). Especialista Legal: LA-6 (expediente conforme). "
               "Excepción manual: NEG / OBS / EXC.")
    ed = st.data_editor(v, hide_index=True, width="stretch", disabled=[c for c in v.columns if c not in
                        ("docs_completos", "expediente_conforme", "excepcion_manual")],
                        column_config={"excepcion_manual": st.column_config.SelectboxColumn(options=[None, "NEG", "OBS", "EXC"])})
    if st.button("Guardar validaciones"):
        with conn.cursor() as cur:
            for r in ed.to_dict("records"):
                cur.execute("UPDATE la_predios SET docs_completos=%s, expediente_conforme=%s, excepcion_manual=%s WHERE cod_predio=%s",
                            (bool(r["docs_completos"]), bool(r["expediente_conforme"]), r["excepcion_manual"] or None, r["cod_predio"]))
        conn.commit()
        db.recalcular_estados(conn)
        st.success("Guardado y recalculado.")
    # datos reservados: la matriz de circulación general no incluye DNI ni teléfonos (Ley 29733)
    ant = db.antecedentes_paso6(conn)
    if not ant.empty:
        with st.expander(f"Consulta inicial a prediantes – verificación de campo Paso 6 ({len(ant)} bloques)"):
            st.caption("Antecedente registrado en la página «ODK / KoBoToolbox» (tenencia, n.° de predios, titular, "
                       "aceptación). Sirve para priorizar la visita; no reemplaza la ficha F-LA-02.")
            codigos = sorted(set(v.cod_unidad)) if not v.empty else []
            st.dataframe(ant[ant.codigo_bloque.isin(codigos)] if codigos else ant, hide_index=True, width="stretch")
    st.download_button("⬇️ Exportar matriz predial (Excel ANIN)",
                       ex.exportar({"Matriz_predial": ("MATRIZ PREDIAL – EXPEDIENTE PRELIMINAR DE LIBERACIÓN DE ÁREAS", v, ["area_unidad_ha"])}),
                       file_name=f"Matriz_Predial_IN_Piura_{date.today():%Y%m%d}.xlsx")


# ------------------------------------------------------------------ 4. avance
def _tab_avance(conn):
    st.subheader("Avance por asistente, distrito y estado · semáforo por unidad")
    u = db.df(conn, "SELECT codigo, tipo_unidad, distrito, asistente, area_ha FROM la_unidades WHERE activo")
    p = db.df(conn, "SELECT cod_predio, cod_unidad, estado_la, area_unidad_ha, nucleo FROM la_predios")
    a = db.df(conn, "SELECT asistente, fecha, conforme FROM la_actas")
    if u.empty:
        st.info("Cargue el catálogo.")
        return
    filas = []
    for x in u.to_dict("records"):
        pp = p[p.cod_unidad == x["codigo"]] if not p.empty else p
        sem, pct = core.semaforo_unidad([{"estado": r["estado_la"], "area_ha": r["area_unidad_ha"], "nucleo": r["nucleo"]}
                                        for r in pp.to_dict("records")], float(x["area_ha"] or 0))
        filas.append({**x, "predios": len(pp), "pct_area_LA4": round(pct * 100, 1), "semaforo": sem})
    s = pd.DataFrame(filas)
    por_asis = s.groupby("asistente").agg(unidades=("codigo", "count"), ha=("area_ha", "sum"), predios=("predios", "sum"),
                                          verdes=("semaforo", lambda z: (z == "VERDE").sum()),
                                          ambar=("semaforo", lambda z: (z == "AMBAR").sum()),
                                          rojas=("semaforo", lambda z: (z == "ROJO").sum()),
                                          sin_datos=("semaforo", lambda z: (z == "SIN_DATOS").sum())).reset_index()
    if not a.empty:
        a["semana"] = pd.to_datetime(a["fecha"]).dt.strftime("%G-S%V")
        act = a.groupby(["asistente", "semana"]).size().rename("actas").reset_index()
        st.markdown("**Actas por semana y asistente**")
        st.bar_chart(act, x="semana", y="actas", color="asistente")
    st.markdown("**Resumen por asistente**")
    st.dataframe(por_asis, hide_index=True, width="stretch")
    if not p.empty:
        st.markdown("**Predios por estado LA**")
        st.dataframe(pd.crosstab(p.merge(u, left_on="cod_unidad", right_on="codigo").asistente, p.estado_la),
                     width="stretch")
    st.markdown("**Semáforo por unidad**")
    st.dataframe(s.style.map(lambda v: {"VERDE": "background-color:#C6EFCE", "AMBAR": "background-color:#FFEB9C",
                                        "ROJO": "background-color:#FFC7CE"}.get(v, ""), subset=["semaforo"]),
                 hide_index=True, width="stretch")
    st.caption(f"Umbrales [POR CONFIRMAR]: verde ≥ {core.SEMAFORO_VERDE:.0%}, ámbar ≥ {core.SEMAFORO_AMBAR:.0%} "
               "del área con estado ≥ LA-4; rojo si hay NEG sobre área núcleo.")
    st.download_button("⬇️ Reporte de avance (Excel ANIN)",
                       ex.exportar({"Avance_asistente": ("AVANCE POR ASISTENTE", por_asis, ["unidades", "ha", "predios"]),
                                    "Semaforo_unidad": ("SEMÁFORO POR UNIDAD", s, ["area_ha", "predios"])}),
                       file_name=f"Avance_LA_IN_Piura_{date.today():%Y%m%d}.xlsx")


# ------------------------------------------------------------------ 5. vivero
def _tab_vivero(conn):
    st.subheader("Vivero Central – comparación de alternativas (F-LA-06)")
    v = db.df(conn, "SELECT * FROM la_vivero_alternativas ORDER BY puntaje DESC NULLS LAST")
    if v.empty:
        st.info("Sin alternativas registradas.")
        return
    st.dataframe(v.drop(columns=["kobo_uuid"]), hide_index=True, width="stretch")
    st.caption(f"Puntaje 0–100 con pesos {core.PESOS_VIVERO} [SUPUESTO: validar con el Especialista en IV]. Sin tasación en esta fase.")


# ------------------------------------------------------------------ 6. catálogo
def _tab_catalogo(conn):
    st.subheader("Catálogo de unidades y enlace con la tabla de bloques")
    st.caption("El catálogo V6 (datos/unidades_liberacion_areas.csv) se carga solo la primera vez. Cada unidad de tipo "
               "bloque se enlaza con bloques(id) por código exacto; los lotes SUS, con su bloque de referencia. "
               "No se modifica la tabla bloques.")
    conc = db.conciliacion(conn)
    if not conc.empty:
        cuenta = conc["estado"].value_counts()
        cols = st.columns(5)
        for c, (k, lbl) in zip(cols, [("OK", "Enlazados OK"), ("AREA_DISTINTA", "Área distinta"),
                                      ("NO_EXISTE_EN_APP", "V6 sin bloque en app"), ("RETIRADO_EN_APP", "V6 retirado en app"),
                                      ("ACTIVO_EN_APP_FUERA_DE_V6", "Activos en app fuera de V6")]):
            c.metric(lbl, int(cuenta.get(k, 0)))
        pend = conc[conc.estado != "OK"]
        if not pend.empty:
            st.warning("Diferencias entre el catálogo V6 y la tabla bloques del aplicativo. Los bloques «activos en app "
                       "fuera de V6» siguen apareciendo en las demás páginas; para ocultarlos sin perder datos, agréguelos "
                       "a BLOQUES_RETIRADOS en database.py (decisión del equipo).")
            st.dataframe(pend, hide_index=True, width="stretch")
    c1, c2 = st.columns(2)
    if c1.button("Re-enlazar con bloques"):
        r = db.vincular_bloques(conn)
        st.success(f"{r['enlazados']} bloques enlazados. Sin enlace: {', '.join(r['sin_enlace']) or 'ninguno'}")
    if c2.button("Aplicar esquema la_* (aditivo)"):
        db.crear_esquema(conn)
        st.success("Esquema aplicado.")
    csv = st.file_uploader("Actualizar unidades.csv (el mismo que se adjunta a los formularios Kobo)", type=["csv"])
    if csv is not None and st.button("Cargar catálogo"):
        n = db.cargar_unidades_csv(conn, pd.read_csv(csv, dtype=str))
        r = db.vincular_bloques(conn)
        st.success(f"{n} unidades cargadas / actualizadas; {r['enlazados']} bloques enlazados.")
    gj = st.file_uploader("Polígonos de bloques y lotes SUS (GeoJSON exportado de QGIS en EPSG:32717)", type=["geojson", "json"])
    campo = st.text_input("Campo con el código de unidad en el GeoJSON", value="name")
    if gj is not None and st.button("Cargar geometrías"):
        try:
            n = db.cargar_geometrias(conn, core.geojson_a_unidades(json.load(gj), campo))
            st.success(f"{n} geometrías actualizadas.")
        except ValueError as e:
            st.error(str(e))
    st.dataframe(db.df(conn, "SELECT codigo, tipo_unidad, distrito, bloque_ref, area_ha, posicion_sus, asistente, "
                             "bloque_id, bloque_ref_id, (geom_wkt IS NOT NULL) AS con_geometria "
                             "FROM la_unidades ORDER BY distrito, codigo"),
                 hide_index=True, width="stretch")
    log = db.df(conn, "SELECT * FROM la_import_log ORDER BY id DESC LIMIT 50")
    if not log.empty:
        st.markdown("**Bitácora de importaciones**")
        st.dataframe(log, hide_index=True, width="stretch")
