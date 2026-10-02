# -*- coding: utf-8 -*-
"""
Página «Liberación de Áreas» del aplicativo IN Piura (CUI 2669244).
Se integra en el menú de streamlit_app.py:
    from liberacion_areas.pagina import render as pagina_liberacion_areas
    ...
    elif pagina == "Liberacion de Areas": pagina_liberacion_areas()
Usa los mismos secrets del aplicativo: DATABASE_URL, KOBO_TOKEN y (opcional) KOBO_SERVER.
Las tablas la_* se crean solas la primera vez (migración aditiva) y se enlazan con `bloques`.

Tres vías de registro equivalentes (mismos campos, mismas validaciones, mismas tablas):
  1. KoboToolbox (API o exportación)   → pestaña «Importar KoboToolbox»
  2. Digitación en el aplicativo       → pestaña «Registro en campo»
  3. Plantilla Excel ANIN              → pestaña «Plantillas Excel / Kobo»
Los registros se consultan, editan, eliminan (con bitácora y restauración) en «Historial / Edición».
"""
from __future__ import annotations

import json
from datetime import date

import pandas as pd
import streamlit as st

from . import la_campos as lc
from . import la_core as core
from . import la_db as db
from . import la_export as ex
from . import la_kobo as kb
from . import la_plantillas as lp
from . import la_reportes as rep

ESTILO_ESTADO = {"NUEVO": "background-color:#C6EFCE", "DUPLICADO": "background-color:#D9D9D9",
                 "OBSERVADO": "background-color:#FFEB9C"}
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
POR_PAGINA = 15


@st.cache_resource(show_spinner="Preparando tablas de liberación de áreas…")
def _inicializar_una_vez():
    """Crea/actualiza las tablas la_* y las enlaza con `bloques` una sola vez por proceso."""
    conn = db.conectar()
    try:
        return db.inicializar_la(conn)
    finally:
        conn.close()


def _flash(msg: str, tipo: str = "success"):
    st.session_state["la_flash"] = (tipo, msg)


def _mostrar_flash(contenedor):
    """El mensaje va en un contenedor que existe siempre: así las pestañas no cambian de posición en la página y
    Streamlit conserva la pestaña activa después de guardar, eliminar o restaurar."""
    f = st.session_state.pop("la_flash", None)
    if f:
        getattr(contenedor, f[0], contenedor.info)(f[1])


def _usuario() -> str:
    return (st.session_state.get("la_usuario") or "").strip()


def _form_label(fid: str) -> str:
    f = lc.FORMULARIOS[fid]
    return f"{f.codigo} – {f.nombre}"


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
        _aplicar_pendiente()
        c1, c2 = st.columns([1, 2])
        c1.text_input("Usuario responsable", key="la_usuario", value=st.session_state.get("usuario", ""),
                      help="Queda registrado en cada registro, edición, eliminación e importación (bitácora).")
        c2.caption("Tres vías de registro equivalentes: **KoboToolbox**, **digitación en el aplicativo** y **plantilla "
                   "Excel**. Todas pasan por las mismas validaciones (catálogo V6, asistente, UTM 17S, lotes SUS, actas) "
                   "y se consultan, editan o eliminan en **Historial / Edición**.")
        _mostrar_flash(st.container())
        cat = db.catalogo(conn)
        t = st.tabs(["1 · Registro en campo", "2 · Plantillas Excel / Kobo", "3 · Importar KoboToolbox",
                     "4 · Historial / Edición", "5 · Matriz predial", "6 · Avance y semáforo", "7 · Reportes",
                     "8 · Vivero Central", "9 · Catálogo y enlace con bloques"])
        with t[0]:
            _tab_registro(conn, cat)
        with t[1]:
            _tab_plantillas(conn, cat)
        with t[2]:
            _tab_importar(conn)
        with t[3]:
            _tab_historial(conn, cat)
        with t[4]:
            _tab_matriz(conn)
        with t[5]:
            _tab_avance(conn)
        with t[6]:
            _tab_reportes(conn, cat)
        with t[7]:
            _tab_vivero(conn)
        with t[8]:
            _tab_catalogo(conn)
    finally:
        conn.close()


# ================================================================== 1. registro en campo (digitación)
def _nonce() -> int:
    return st.session_state.get("la_nonce", 0)


def _k(fid: str, nombre: str) -> str:
    return f"la_f{_nonce()}_{fid}_{nombre}"


def _a_fecha(v):
    if isinstance(v, date):
        return v
    s = lc._fecha_txt(v)
    try:
        return date.fromisoformat(s) if s else None
    except ValueError:
        return None


def _cargar_edicion(fid: str, valores: dict, reps: dict[str, list[dict]], edit: dict):
    """Precarga los widgets del formulario (modo edición) en el siguiente rerun."""
    st.session_state["la_nonce"] = _nonce() + 1
    st.session_state["la_form_sel"] = fid
    st.session_state["la_edit"] = edit
    f = lc.FORMULARIOS[fid]
    for c in f.campos:
        if c.tipo == "utm":
            for suf in lc.UTM_SUFIJOS:
                st.session_state[_k(fid, f"{c.nombre}_{suf}")] = core._num(valores.get(f"{c.nombre}_{suf}"))
            continue
        v = valores.get(c.nombre)
        if c.tipo == "date":
            v = _a_fecha(v)
        elif c.tipo == "int":
            v = int(core._num(v)) if core._num(v) is not None else None
        elif c.tipo == "decimal":
            v = core._num(v)
        elif c.tipo == "select_multiple":
            v = list(v or [])
        elif c.tipo == "unidades":
            v = str(v or "").split()
        elif c.tipo == "select_one":
            v = lc.codigo_opcion(c.lista, v)
        st.session_state[_k(fid, c.nombre)] = v
    st.session_state["la_rep_init"] = {rp.nombre: reps.get(rp.nombre) or [] for rp in f.repeats}


def _aplicar_pendiente():
    """La precarga de edición se aplica al inicio del rerun, antes de crear los widgets (regla de Streamlit)."""
    p = st.session_state.pop("la_pending", None)
    if p:
        _cargar_edicion(*p)


def _cancelar_edicion():
    st.session_state.pop("la_edit", None)
    st.session_state.pop("la_rep_init", None)
    st.session_state["la_nonce"] = _nonce() + 1


def _widget(c: lc.Campo, fid: str, col, cat: dict, valores: dict):
    key = _k(fid, c.nombre)
    etiqueta = c.etiqueta + (" *" if c.requerido else "")
    ayuda = c.ayuda or None
    actual = st.session_state.get(key)
    if c.tipo == "text":
        return col.text_input(etiqueta, key=key, help=ayuda)
    if c.tipo == "textarea":
        return col.text_area(etiqueta, key=key, help=ayuda, height=80)
    if c.tipo == "int":
        return col.number_input(etiqueta, key=key, help=ayuda, step=1, value=None,
                                min_value=int(c.minimo) if c.minimo is not None else None,
                                max_value=int(c.maximo) if c.maximo is not None else None)
    if c.tipo == "decimal":
        return col.number_input(etiqueta, key=key, help=ayuda, value=None, format="%.3f",
                                min_value=float(c.minimo) if c.minimo is not None else None,
                                max_value=float(c.maximo) if c.maximo is not None else None)
    if c.tipo == "date":
        if key not in st.session_state and c.nombre in ("hoy", "fecha_evento", "fecha_insp", "fecha_acta", "fecha"):
            st.session_state[key] = date.today()
        return col.date_input(etiqueta, key=key, help=ayuda, format="DD/MM/YYYY", value=None)
    if c.tipo == "select_one":
        codigos = [x for x, _ in lc.OPCIONES[c.lista]]
        if actual and actual not in codigos:
            codigos.append(actual)
        return col.selectbox(etiqueta, [None] + codigos, key=key, help=ayuda,
                             format_func=lambda x: "— seleccione —" if x is None else lc.etiqueta_opcion(c.lista, x))
    if c.tipo == "select_multiple":
        codigos = [x for x, _ in lc.OPCIONES[c.lista]] + [x for x in (actual or []) if x not in dict(lc.OPCIONES[c.lista])]
        return col.multiselect(etiqueta, codigos, key=key, help=ayuda, format_func=lambda x: lc.etiqueta_opcion(c.lista, x))
    if c.tipo in ("unidad", "unidades"):
        asis = valores.get("asistente")
        codigos = sorted(u for u, r in cat.items() if not asis or asis == "ESP" or r.get("asistente") == asis)
        extra = actual if isinstance(actual, list) else ([actual] if actual else [])
        codigos += [x for x in extra if x not in codigos]
        fmt = (lambda u: "— seleccione —" if u is None else f"{u} · {cat.get(u, {}).get('label') or u}")
        if c.tipo == "unidad":
            return col.selectbox(etiqueta, [None] + codigos, key=key, format_func=fmt,
                                 help="Lista filtrada por el asistente elegido (ESP ve todas).")
        return col.multiselect(etiqueta, codigos, key=key, format_func=fmt, help=ayuda)
    return None


def _widget_utm(c: lc.Campo, fid: str) -> dict:
    st.markdown(f"**{c.etiqueta}{' *' if c.requerido else ''}** · UTM WGS84 Zona 17S")
    cols = st.columns(4)
    out = {}
    for col, (suf, et, fmt) in zip(cols, [("este", "Este (m)", "%.2f"), ("norte", "Norte (m)", "%.2f"),
                                          ("alt", "Altitud (m)", "%.1f"), ("prec", "Precisión GPS (m)", "%.1f")]):
        out[f"{c.nombre}_{suf}"] = col.number_input(et, key=_k(fid, f"{c.nombre}_{suf}"), value=None, format=fmt)
    e, n = out[f"{c.nombre}_este"], out[f"{c.nombre}_norte"]
    if e is not None and n is not None:
        if lc.utm_en_rango(e, n):
            gp = core.parse_geopoint(lc.utm_a_geopoint(e, n))
            st.caption(f"✅ En rango del ámbito · Lat {gp['lat']:.6f}, Lon {gp['lon']:.6f}")
        else:
            st.warning(f"Coordenada fuera del rango UTM 17S del ámbito (Este {core.E_MIN:,}–{core.E_MAX:,} m; "
                       f"Norte {core.N_MIN:,}–{core.N_MAX:,} m).")
    return out


def _editor_repeat(rp: lc.Repeat, fid: str) -> list[dict]:
    """Tabla editable de un repeat (puntos, vértices, firmantes)."""
    cols, cfg = [], {}
    for c in rp.campos:
        if c.tipo == "utm":
            cols += [f"{c.nombre}_este", f"{c.nombre}_norte"]
            cfg[f"{c.nombre}_este"] = st.column_config.NumberColumn("Este (m)", format="%.2f",
                                                                    min_value=core.E_MIN, max_value=core.E_MAX)
            cfg[f"{c.nombre}_norte"] = st.column_config.NumberColumn("Norte (m)", format="%.2f",
                                                                     min_value=core.N_MIN, max_value=core.N_MAX)
        elif c.tipo == "select_one":
            cols.append(c.nombre)
            cfg[c.nombre] = st.column_config.SelectboxColumn(c.etiqueta, options=[x for x, _ in lc.OPCIONES[c.lista]])
        elif c.tipo == "int":
            cols.append(c.nombre)
            cfg[c.nombre] = st.column_config.NumberColumn(c.etiqueta, step=1)
        else:
            cols.append(c.nombre)
            cfg[c.nombre] = st.column_config.TextColumn(c.etiqueta)
    ini = (st.session_state.get("la_rep_init") or {}).get(rp.nombre) or []
    datos = pd.DataFrame(ini).reindex(columns=cols) if ini else pd.DataFrame(columns=cols)
    for c in rp.campos:
        if c.tipo == "select_one" and c.nombre in datos:
            datos[c.nombre] = datos[c.nombre].map(lambda v, c=c: lc.codigo_opcion(c.lista, v))
    numericas = [f"{c.nombre}_{x}" for c in rp.campos if c.tipo == "utm" for x in ("este", "norte")] + \
        [c.nombre for c in rp.campos if c.tipo == "int"]
    for col in numericas:
        datos[col] = pd.to_numeric(datos[col], errors="coerce")
    ed = st.data_editor(datos, num_rows="dynamic", hide_index=True, width="stretch", column_config=cfg,
                        key=_k(fid, f"rep_{rp.nombre}"))
    return ed.to_dict("records")


def _tab_registro(conn, cat: dict):
    st.subheader("Registro en campo – digitación directa en el aplicativo")
    if not cat:
        st.info("Cargue primero el catálogo de unidades (pestaña 9).")
        return
    edit = st.session_state.get("la_edit")
    if edit:
        st.markdown(f"#### ✏️ Modo edición · {_form_label(edit['form_id'])} · `{edit.get('clave')}`")
        st.caption("Corrija los campos y pulse **Actualizar registro**. La versión anterior queda en la bitácora.")
        if st.button("Cancelar edición (nuevo registro)", key="la_cancel_edit"):
            _cancelar_edicion()
            st.rerun()
    fid = st.selectbox("Formulario", list(lc.FORMULARIOS), key="la_form_sel", format_func=_form_label,
                       disabled=bool(edit))
    f = lc.FORMULARIOS[fid]
    if f.kobo:
        st.caption("Mismos campos que el formulario KoboToolbox y la plantilla Excel. Coordenadas en UTM WGS84 17S.")
    valores: dict = {}
    cols = st.columns(3)
    i = 0
    for c in f.campos:
        tipo_u = cat.get(valores.get("unidad") or "", {}).get("tipo_unidad")
        contexto = {**valores, "tipo_unidad": tipo_u}
        if not lc.es_relevante(c, contexto):
            continue
        if c.tipo == "utm":
            valores.update(_widget_utm(c, fid))
            cols = st.columns(3)
            i = 0
            continue
        if c.tipo == "textarea":
            valores[c.nombre] = _widget(c, fid, st, cat, valores)
            cols = st.columns(3)
            i = 0
            continue
        valores[c.nombre] = _widget(c, fid, cols[i % 3], cat, valores)
        i += 1
        if c.tipo == "unidad" and valores.get("unidad"):
            u = cat.get(valores["unidad"], {})
            st.info(f"**{u.get('label') or valores['unidad']}** · {u.get('tipo_unidad')} · provincia "
                    f"{str(u.get('provincia', '')).title()} · asistente {u.get('asistente')} {u.get('asistente_nombre') or ''}"
                    + (f" · bloque de referencia {u.get('bloque_ref')}" if u.get("tipo_unidad") == "lote_sus" else ""))
            cols = st.columns(3)
            i = 0
    if f.con_predio and valores.get("unidad") and valores.get("n_predio") is not None:
        st.caption(f"Código de predio: **{lc.cod_predio(valores['unidad'], valores['n_predio'])}**")
    reps = {}
    tipo_u = cat.get(valores.get("unidad") or "", {}).get("tipo_unidad")
    for rp in f.repeats:
        if not lc.es_relevante(rp, {**valores, "tipo_unidad": tipo_u}):
            continue
        st.markdown(f"**{rp.etiqueta}** (agregue filas con ➕)")
        reps[rp.nombre] = _editor_repeat(rp, fid)

    etiqueta = "💾 Actualizar registro" if edit else "💾 Guardar registro"
    if st.button(etiqueta, type="primary", key="la_guardar"):
        if not _usuario():
            st.warning("Ingrese el usuario responsable (arriba).")
            return
        _guardar(conn, cat, fid, valores, reps, edit)


def _valores_limpios(valores: dict) -> dict:
    out = {}
    for k, v in valores.items():
        if isinstance(v, date):
            v = v.isoformat()
        out[k] = v
    return out


def _guardar(conn, cat: dict, fid: str, valores: dict, reps: dict, edit: dict | None):
    valores = _valores_limpios(valores)
    if fid == "f_la_05_documentos":
        reg, avisos = lc.construir_registro(fid, valores, cat)
        faltan = [a for a in avisos if a.startswith("Falta")]
        if faltan:
            st.error("No se guardó: " + "; ".join(faltan))
            return
        datos = lp._documento_desde_registro(reg)
        if edit:
            db.actualizar_documento(conn, edit["id"], datos, _usuario())
            _flash(f"Documento {edit['clave']} actualizado.")
        else:
            previo = db.documento_existe(conn, datos)
            if previo:
                st.warning(f"Ese documento ya está registrado ({previo}). Use Historial / Edición para modificarlo.")
                return
            _flash(f"Documento registrado con código {db.registrar_documento(conn, datos)}.")
        _cancelar_edicion()
        st.rerun()
        return

    uid = edit["clave"] if edit else lc.nuevo_uuid("app")
    base = None
    if edit:
        env = db.obtener_envio(conn, uid)
        base = (env or {}).get("payload")
    reg, avisos = lc.construir_registro(fid, valores, cat, reps, kobo_uuid=uid, base=base)
    faltan = [a for a in avisos if a.startswith("Falta")]
    if faltan:
        st.error("No se guardó: " + "; ".join(faltan))
        return
    existentes = db.uuids_existentes(conn) - {uid}
    r = core.validar_envio(fid, reg, cat, db.geometrias(conn), existentes, db.sus_areas(conn, cat))
    r.kobo_uuid = uid
    for a in avisos:
        if a not in r.motivos:
            r.observar(a)
    if edit:
        db.actualizar_envio(conn, r, fid, _usuario())
        accion = "actualizado"
    else:
        db.importar(conn, [r], "APP", fid, _usuario())
        accion = "registrado"
    msg = f"{lc.FORMULARIOS[fid].codigo} {accion}" + (f" · predio {r.cod_predio}" if r.cod_predio else "") + \
          f" · estado de importación: {r.estado_import}."
    if r.motivos:
        msg += " Motivos: " + " | ".join(r.motivos)
    _flash(msg, "warning" if r.estado_import == "OBSERVADO" else "success")
    _cancelar_edicion()
    st.rerun()


# ================================================================== 2. plantillas
def _tab_plantillas(conn, cat: dict):
    st.subheader("Plantillas Excel de Liberación de Áreas y formularios KoboToolbox")
    if not cat:
        st.info("Cargue primero el catálogo de unidades (pestaña 9).")
        return
    st.markdown("**1. Descargar plantilla Excel (formato ANIN, con listas desplegables)**")
    st.caption("Una hoja por formulario (F-LA-01 a F-LA-06), hojas de detalle (puntos, vértices SUS, firmantes), "
               "catálogo de unidades V6 e instrucciones. Coordenadas en UTM WGS84 17S.")
    c1, c2 = st.columns([1, 2])
    asis = c1.selectbox("Preparar para", ["Todos"] + [x for x, _ in lc.OPCIONES["asistente"] if x != "ESP"], key="la_pl_asis")
    if c2.button("Generar plantilla Excel", key="la_pl_gen"):
        st.session_state["la_pl_bytes"] = (asis, lp.generar_plantilla_excel(cat, None if asis == "Todos" else asis))
    if st.session_state.get("la_pl_bytes"):
        a, b = st.session_state["la_pl_bytes"]
        sufijo = "" if a == "Todos" else f"_{a}"
        st.download_button("⬇️ Descargar plantilla Excel", b, key="la_pl_dl", mime=MIME_XLSX,
                           file_name=f"Plantilla_Liberacion_Areas_IN_Piura{sufijo}_{date.today():%Y%m%d}.xlsx")

    st.markdown("---")
    st.markdown("**2. Formularios KoboToolbox (XLSForm) con los mismos campos**")
    st.caption("ZIP con los XLSForm F-LA-01/02/03/04/06, el archivo unidades.csv (multimedia de cada formulario) y "
               "un LEEME. Cárguelos en KoboToolbox si aún no están desplegados o si cambió el catálogo.")
    c1, c2 = st.columns(2)
    if c1.button("Generar paquete KoboToolbox", key="la_kobo_gen"):
        st.session_state["la_kobo_zip"] = lp.paquete_kobo(cat)
    if st.session_state.get("la_kobo_zip"):
        c1.download_button("⬇️ Descargar XLSForms (ZIP)", st.session_state["la_kobo_zip"], key="la_kobo_dl",
                           file_name=f"Formularios_Kobo_Liberacion_Areas_{date.today():%Y%m%d}.zip", mime="application/zip")
    c2.download_button("⬇️ unidades.csv (catálogo vigente)", lp.unidades_csv(cat), key="la_csv_dl",
                       file_name="unidades.csv", mime="text/csv")

    st.markdown("---")
    st.markdown("**3. Importar plantilla Excel llenada**")
    arch = st.file_uploader("Plantilla de Liberación de Áreas (.xlsx, máx. 25 MB)", type=["xlsx"], key="la_pl_up")
    if arch is None:
        return
    if arch.size > 25 * 1024 * 1024:
        st.error("El archivo excede 25 MB.")
        return
    try:
        leido = lp.leer_plantilla_excel(arch, cat)
    except Exception as e:  # noqa: BLE001
        st.error(f"No se pudo leer el archivo: {e}")
        return
    st.caption("Hojas leídas: " + ", ".join(leido["hojas_leidas"]))
    for e in leido["errores"]:
        st.error(e)
    geoms, existentes, sus = db.geometrias(conn), db.uuids_existentes(conn), db.sus_areas(conn, cat)
    validados: dict[str, list] = {}
    filas_vista = []
    for fid, items in leido["envios"].items():
        res = []
        for it in items:
            r = core.validar_envio(fid, it["registro"], cat, geoms, existentes, sus)
            if r.estado_import != "DUPLICADO":
                for a in it["avisos"]:
                    if a not in r.motivos:
                        r.observar(a)
            res.append(r)
            filas_vista.append({"hoja": it["hoja"], "fila": it["fila"], "estado_import": r.estado_import,
                                "motivo": " | ".join(r.motivos), "unidad": r.cod_unidad, "predio": r.cod_predio,
                                "asistente": r.datos.get("asistente"), "fecha": r.datos.get("hoy")})
        validados[fid] = res
    docs = []
    for d in leido["documentos"]:
        previo = db.documento_existe(conn, d["datos"])
        estado = "DUPLICADO" if previo else ("OBSERVADO" if d["avisos"] else "NUEVO")
        docs.append({**d, "estado": estado})
        filas_vista.append({"hoja": "F-LA-05", "fila": d["fila"], "estado_import": estado,
                            "motivo": (f"Ya registrado ({previo})" if previo else " | ".join(d["avisos"])),
                            "unidad": d["datos"].get("cod_unidad"), "predio": d["datos"].get("cod_predio"),
                            "asistente": None, "fecha": d["datos"].get("fecha")})
    if not filas_vista:
        st.info("La plantilla no tiene filas llenadas.")
        return
    vista = pd.DataFrame(filas_vista)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Filas leídas", len(vista))
    c2.metric("Nuevas", int((vista.estado_import == "NUEVO").sum()))
    c3.metric("Duplicadas", int((vista.estado_import == "DUPLICADO").sum()))
    c4.metric("Observadas", int((vista.estado_import == "OBSERVADO").sum()))
    st.dataframe(vista.style.map(lambda v: ESTILO_ESTADO.get(v, ""), subset=["estado_import"]), width="stretch",
                 hide_index=True)
    st.caption("Las filas OBSERVADAS se registran con su motivo (igual que en Kobo) para que el Especialista Predial "
               "las resuelva; las DUPLICADAS se ignoran. F-LA-05 observadas (faltan datos obligatorios) no se cargan.")
    por_cargar = int(vista.estado_import.isin(["NUEVO", "OBSERVADO"]).sum())
    if st.button(f"✅ Confirmar importación de la plantilla ({por_cargar} filas)", type="primary", key="la_pl_conf",
                 disabled=not _usuario() or por_cargar == 0):
        total = 0
        for fid, res in validados.items():
            total += db.importar(conn, res, "PLANTILLA", fid, _usuario())["insertados"]
        nd = 0
        for d in docs:
            if d["estado"] == "NUEVO":
                db.registrar_documento(conn, d["datos"])
                nd += 1
        _flash(f"Plantilla importada: {total} registros de formularios y {nd} documentos F-LA-05. Estados LA recalculados.")
        st.rerun()
    if not _usuario():
        st.caption("Ingrese el usuario responsable (arriba) para confirmar.")


# ================================================================== 3. importar KoboToolbox
def _tab_importar(conn):
    st.subheader("Importar envíos de KoboToolbox")
    fuente = st.radio("Fuente", ["API KoboToolbox", "Archivo exportado (XLSX / JSON)"], horizontal=True)
    form_id = st.selectbox("Formulario", list(core.FORMULARIOS),
                           format_func=lambda f: f"{core.FORMULARIOS[f]['codigo']} – {core.FORMULARIOS[f]['nombre']}")
    kobo = None
    if fuente.startswith("API"):
        from odk_kobo import CLAVE_TOKEN_SESION, normalizar_token, token_kobo
        token, fuente_token = token_kobo()
        if fuente_token == "secrets":
            st.caption("Token API leído de la configuración segura (secrets).")
        else:
            st.warning("No se encontró KOBO_TOKEN en los secrets del aplicativo. Puede escribir el token aquí "
                       "(solo se usa en esta sesión) o agregarlo en Streamlit Cloud → Settings → Secrets como "
                       "`KOBO_TOKEN = \"...\"` en una línea propia, al inicio del cuadro.")
            escrito = st.text_input("Token API de KoboToolbox", type="password", value=token, key="la_token_input",
                                    help="KoboToolbox → Account Settings → Security → API Key. "
                                         "No se guarda en la base de datos.")
            token = normalizar_token(escrito)
            st.session_state[CLAVE_TOKEN_SESION] = token
            if not token:
                return
        try:
            kobo = kb.cliente(token=token)
        except ValueError as e:
            st.error(str(e))
            return
        if st.button("Listar formularios F-LA"):
            with st.spinner("Consultando KoboToolbox…"):
                try:
                    st.session_state["la_forms"] = kb.formularios_la(kobo)
                except Exception as e:  # noqa: BLE001 – token inválido, sin red, etc.
                    st.error(f"No se pudo consultar KoboToolbox: {e}")
        forms = st.session_state.get("la_forms", [])
        uid = st.selectbox("Formulario en KoboToolbox", [f["uid"] for f in forms],
                           format_func=lambda u: next((f"{f['nombre']} · {f['envios']} envíos" for f in forms if f["uid"] == u), u)) if forms else \
            st.text_input("UID del formulario (asset uid)", help="Aparece en la URL del formulario: /#/forms/<uid>/…")
        if st.button("Descargar envíos", disabled=not uid):
            with st.spinner("Descargando todas las páginas de envíos…"):
                try:
                    st.session_state["la_raw"] = kobo.obtener_envios(uid)
                except Exception as e:  # noqa: BLE001
                    st.error(f"No se pudieron descargar los envíos: {e}")
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
    if st.button("✅ Confirmar importación", type="primary", disabled=not _usuario()):
        res = db.importar(conn, resultados, "API" if kobo else "ARCHIVO", form_id, _usuario())
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
    if not _usuario():
        st.caption("Ingrese el usuario responsable (arriba) para confirmar.")


# ================================================================== 4. historial / edición
def _paginar(d: pd.DataFrame, clave: str) -> pd.DataFrame:
    total = max(1, -(-len(d) // POR_PAGINA))
    pag = st.session_state.get(clave, 1)
    pag = min(max(1, pag), total)
    if total > 1:
        c1, c2, c3 = st.columns([1, 2, 1])
        if c1.button("◀ Anterior", key=f"{clave}_prev", disabled=pag <= 1):
            st.session_state[clave] = pag - 1
            st.rerun()
        c2.markdown(f"<div style='text-align:center'>Página {pag} de {total} · {len(d)} registros</div>",
                    unsafe_allow_html=True)
        if c3.button("Siguiente ▶", key=f"{clave}_next", disabled=pag >= total):
            st.session_state[clave] = pag + 1
            st.rerun()
    return d.iloc[(pag - 1) * POR_PAGINA: pag * POR_PAGINA]


def _humanizar(fid: str, payload: dict) -> pd.DataFrame:
    """Campos del formulario con su etiqueta y el valor legible (sin DNI ni celular)."""
    f = lc.FORMULARIOS[fid]
    filas = []
    for c in f.campos:
        v = payload.get(c.nombre)
        if c.nombre == "unidad" and not v:
            v = payload.get("cod_unidad")
        if v in (None, "", []):
            continue
        if c.nombre.endswith(("_dni", "_celular", "_cel", "_ruc")):
            v = "[dato reservado – Ley 29733]"
        elif c.tipo == "select_one":
            v = lc.etiqueta_opcion(c.lista, v)
        elif c.tipo == "select_multiple":
            v = ", ".join(lc.etiqueta_opcion(c.lista, x) for x in lc.codigos_multiples(c.lista, v))
        elif c.tipo == "utm":
            u = lc.geopoint_a_utm(v)
            v = f"E {u.get('este'):,.2f} · N {u.get('norte'):,.2f} (prec. {u.get('prec') or '-'} m)" if u else v
        filas.append({"campo": c.etiqueta, "valor": str(v)})
    for k in ("cod_predio", "tipo_unidad", "bloque_ref", "validacion_espacial", "distancia_m", "area_sus_ha",
              "posicion_sus_calc"):
        if payload.get(k) not in (None, ""):
            filas.append({"campo": k, "valor": str(payload[k])})
    return pd.DataFrame(filas)


def _tab_historial(conn, cat: dict):
    st.subheader("Historial / Edición de registros")
    st.caption("Registros de las tres vías (Kobo, aplicativo, plantilla). **Editar** los carga en «Registro en campo»; "
               "**Eliminar** guarda una copia en la bitácora y puede restaurarse. Las fotos nunca se borran.")
    if st.session_state.get("la_edit"):
        st.info("✏️ Hay un registro cargado en **modo edición**: abra la pestaña **1 · Registro en campo**.")
    conteo = db.df(conn, "SELECT form_id, count(*) AS n FROM la_envios_raw GROUP BY form_id")
    n_por = dict(zip(conteo.form_id, conteo.n)) if not conteo.empty else {}
    n_por["f_la_05_documentos"] = int(db.df(conn, "SELECT count(*) AS n FROM la_documentos")["n"][0])
    fid = st.selectbox("Formulario", list(lc.FORMULARIOS), key="la_h_form",
                       format_func=lambda x: f"{_form_label(x)} ({n_por.get(x, 0)})")
    es_doc = fid == "f_la_05_documentos"
    d = rep.tabla_formulario(conn, fid) if es_doc else db.listar_envios(conn, fid)
    if d.empty:
        st.info("Sin registros para este formulario.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        fu = c1.multiselect("Unidad", sorted(d["cod_unidad"].dropna().astype(str).unique()), key="la_h_u")
        texto = c4.text_input("Buscar (predio / código)", key="la_h_q")
        mask = pd.Series(True, index=d.index)
        if fu:
            mask &= d["cod_unidad"].astype(str).apply(lambda v: any(u in v.split() for u in fu))
        if not es_doc:
            fa = c2.multiselect("Asistente", sorted(d["asistente"].dropna().unique()), key="la_h_a")
            fe = c3.multiselect("Estado / origen", ["NUEVO", "OBSERVADO", "KOBO", "APP", "PLANTILLA"], key="la_h_e")
            if fa:
                mask &= d["asistente"].isin(fa)
            if fe:
                mask &= d["estado_import"].isin(fe) | d["origen"].isin(fe)
        if texto:
            mask &= d.astype(str).apply(lambda r: r.str.contains(texto, case=False, regex=False).any(), axis=1)
        v = d[mask]
        _lista_registros(conn, cat, fid, v, es_doc)
        st.download_button("⬇️ Exportar registros filtrados (Excel ANIN)",
                           ex.exportar({lc.FORMULARIOS[fid].codigo: (f"{_form_label(fid).upper()} – REGISTROS",
                                                                     _exportable(conn, fid, v, es_doc), None)}),
                           file_name=f"{lc.FORMULARIOS[fid].codigo}_IN_Piura_{date.today():%Y%m%d}.xlsx", mime=MIME_XLSX,
                           key="la_h_xls")
        _detalle(conn, fid, v, es_doc)
    _papelera(conn)


def _exportable(conn, fid: str, v: pd.DataFrame, es_doc: bool) -> pd.DataFrame:
    if es_doc:
        return v
    t = rep.tabla_formulario(conn, fid)
    return t[t["kobo_uuid"].isin(set(v["kobo_uuid"]))] if not t.empty else t


def _lista_registros(conn, cat: dict, fid: str, v: pd.DataFrame, es_doc: bool):
    pagina = _paginar(v, f"la_h_pag_{fid}")
    anchos = [1.1, 1.0, 0.8, 0.9, 0.9, 1.6, 0.6, 0.7]
    enc = ["Código" if es_doc else "Predio", "Unidad", "Asistente" if not es_doc else "Tipo", "Fecha",
           "Origen" if not es_doc else "Resultado", "Estado / motivos" if not es_doc else "Entidad", "", ""]
    for col, h in zip(st.columns(anchos), enc):
        col.markdown(f"**{h}**")
    confirmar = st.session_state.get("la_h_del")
    for r in pagina.to_dict("records"):
        clave = int(r["id"]) if es_doc else r["kobo_uuid"]
        fila = st.columns(anchos)
        if es_doc:
            vals = [r["cod_doc"], r["cod_unidad"], lc.etiqueta_opcion("doc_tipo", r["tipo"]), r["fecha"], r["resultado"],
                    r["entidad"]]
        else:
            vals = [r["cod_predio"] or "—", r["cod_unidad"], r["asistente"],
                    (pd.to_datetime(r["fecha_envio"]).strftime("%d/%m/%Y") if r["fecha_envio"] is not None and not pd.isna(r["fecha_envio"]) else "—"),
                    r["origen"], ("✅ " if r["estado_import"] == "NUEVO" else "⚠️ ") + (r["motivos"] or r["estado_import"])[:90]]
        for col, x in zip(fila[:6], vals):
            col.write("—" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x))
        if fila[6].button("Editar", key=f"la_ed_{clave}", type="primary"):
            _editar(conn, fid, r, es_doc)
            st.rerun()
        if confirmar == clave:
            if fila[7].button("Confirmar", key=f"la_delok_{clave}", type="primary"):
                if not _usuario():
                    st.warning("Ingrese el usuario responsable (arriba).")
                else:
                    (db.eliminar_documento if es_doc else db.eliminar_envio)(conn, clave, _usuario())
                    st.session_state.pop("la_h_del", None)
                    _flash(f"Registro {r.get('cod_doc') or r.get('cod_predio') or clave} eliminado. Puede restaurarlo "
                           "desde «Registros eliminados».")
                    st.rerun()
        elif fila[7].button("Eliminar", key=f"la_del_{clave}"):
            st.session_state["la_h_del"] = clave
            st.rerun()
    if confirmar is not None:
        st.warning("Confirme la eliminación pulsando **Confirmar** en su fila. Se guarda una copia en la bitácora.")
        if st.button("Cancelar eliminación", key="la_h_del_cancel"):
            st.session_state.pop("la_h_del", None)
            st.rerun()


def _editar(conn, fid: str, r: dict, es_doc: bool):
    if es_doc:
        doc = db.obtener_documento(conn, int(r["id"]))
        n = None
        if doc.get("cod_predio"):
            try:
                n = int(str(doc["cod_predio"]).rsplit("-P", 1)[-1])
            except ValueError:
                n = None
        valores = {"unidad": doc["cod_unidad"], "n_predio": n, **{k: doc.get(k) for k in
                   ("tipo", "entidad", "fecha", "resultado", "n_partida", "descripcion", "archivo_url", "registrado_por")}}
        st.session_state["la_pending"] = (fid, valores, {}, {"form_id": fid, "id": int(doc["id"]), "clave": doc["cod_doc"]})
    else:
        env = db.obtener_envio(conn, r["kobo_uuid"])
        if env is None:
            _flash("El registro ya no existe.", "warning")
            return
        valores, reps = lc.registro_a_valores(fid, env["payload"])
        st.session_state["la_pending"] = (fid, valores, reps, {"form_id": fid, "clave": env["kobo_uuid"]})
    _flash("Registro cargado en modo edición: abra la pestaña **1 · Registro en campo**.", "info")


def _detalle(conn, fid: str, v: pd.DataFrame, es_doc: bool):
    st.markdown("#### Detalle del registro")
    if es_doc:
        opciones = {f"{r['cod_doc']} · {r['tipo']}": int(r["id"]) for r in v.to_dict("records")}
    else:
        opciones = {f"{r['cod_predio'] or r['cod_unidad']} · {r['origen']} · {r['kobo_uuid'][:13]}": r["kobo_uuid"]
                    for r in v.to_dict("records")}
    sel = st.selectbox("Seleccione un registro", [""] + list(opciones), key=f"la_det_{fid}")
    if not sel:
        return
    if es_doc:
        doc = db.obtener_documento(conn, opciones[sel])
        st.dataframe(pd.DataFrame([{"campo": k, "valor": str(x)} for k, x in doc.items() if x not in (None, "")]),
                     hide_index=True, width="stretch")
        clave = doc["cod_doc"]
        tabla = "la_documentos"
    else:
        env = db.obtener_envio(conn, opciones[sel])
        if env is None:
            return
        if env.get("motivos"):
            st.warning("Motivos: " + env["motivos"])
        st.dataframe(_humanizar(fid, env["payload"]), hide_index=True, width="stretch")
        for rp in lc.FORMULARIOS[fid].repeats:
            if rp.nombre == "poligono_sus":
                continue
            filas = env["payload"].get(rp.nombre) or []
            if filas:
                st.markdown(f"**{rp.etiqueta}**")
                tabla_rep = pd.DataFrame(filas)
                if "pt_gps" in tabla_rep:
                    utm = tabla_rep["pt_gps"].map(lc.geopoint_a_utm)
                    tabla_rep["este"] = utm.map(lambda u: u.get("este"))
                    tabla_rep["norte"] = utm.map(lambda u: u.get("norte"))
                    tabla_rep = tabla_rep.drop(columns=["pt_gps"])
                if "fir_dni" in tabla_rep:
                    tabla_rep = tabla_rep.drop(columns=["fir_dni"])
                st.dataframe(tabla_rep, hide_index=True, width="stretch")
        adj = db.adjuntos_envio(conn, env["kobo_uuid"])
        if not adj.empty:
            st.markdown(f"**Adjuntos ({len(adj)})**")
            cols = st.columns(4)
            for i, a in enumerate(adj.to_dict("records")):
                if a["contenido"] is not None and str(a["mimetype"] or "").startswith("image"):
                    cols[i % 4].image(bytes(a["contenido"]), caption=a["nombre_archivo"])
                elif a["url_original"]:
                    cols[i % 4].markdown(f"[{a['nombre_archivo']}]({a['url_original']})")
        st.caption(f"Origen {env.get('origen')} · registrado {env.get('creado')}"
                   + (f" · editado {env.get('editado')} por {env.get('editado_por')}" if env.get("editado") else ""))
        clave = env["kobo_uuid"]
        tabla = "la_envios_raw"
    hist = db.df(conn, "SELECT fecha, usuario, accion FROM la_bitacora WHERE tabla = %s AND clave = %s ORDER BY id DESC",
                 (tabla, clave))
    if not hist.empty:
        with st.expander(f"Historial de cambios ({len(hist)})"):
            st.dataframe(hist, hide_index=True, width="stretch")


def _papelera(conn):
    st.markdown("---")
    elim = db.bitacora(conn, solo_eliminados=True)
    with st.expander(f"🗑️ Registros eliminados – restaurables ({len(elim)})"):
        if elim.empty:
            st.caption("No hay registros eliminados.")
        for r in elim.to_dict("records"):
            c1, c2, c3, c4 = st.columns([1.2, 1.5, 1.2, 0.8])
            c1.write(f"{pd.to_datetime(r['fecha']):%d/%m/%Y %H:%M}")
            c2.write(f"{lc.FORMULARIOS.get(r['form_id']).codigo if r['form_id'] in lc.FORMULARIOS else r['tabla']} · {r['clave'][:30]}")
            c3.write(r["usuario"] or "—")
            if c4.button("Restaurar", key=f"la_rest_{r['id']}"):
                if not _usuario():
                    st.warning("Ingrese el usuario responsable (arriba).")
                else:
                    try:
                        clave = db.restaurar(conn, int(r["id"]), _usuario())
                        env = db.obtener_envio(conn, clave) if r["tabla"] == "la_envios_raw" else None
                        cod = lc.FORMULARIOS[r["form_id"]].codigo if r["form_id"] in lc.FORMULARIOS else r["tabla"]
                        _flash(f"{cod} restaurado: {(env or {}).get('cod_predio') or (env or {}).get('cod_unidad') or clave}"
                               + (f" · estado {env['estado_import']}" if env else "") + ".")
                    except ValueError as e:
                        _flash(str(e), "error")
                    st.rerun()
    with st.expander("📜 Bitácora de cambios (últimos 500)"):
        b = db.bitacora(conn)
        if b.empty:
            st.caption("Sin cambios registrados.")
        else:
            st.dataframe(b, hide_index=True, width="stretch")
            st.download_button("⬇️ Exportar bitácora", ex.exportar({"Bitacora": ("BITÁCORA DE CAMBIOS – LIBERACIÓN DE ÁREAS", b, None)}),
                               file_name=f"Bitacora_LA_IN_Piura_{date.today():%Y%m%d}.xlsx", mime=MIME_XLSX, key="la_bit_xls")


# ================================================================== 5. matriz
EDITABLES_MATRIZ = ("nombre_predio", "area_unidad_ha", "nucleo", "docs_completos", "expediente_conforme",
                    "excepcion_manual", "alertas")


def _tab_matriz(conn):
    st.subheader("Matriz predial")
    m = rep.matriz(conn)
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
    st.caption("Editables: nombre, área dentro de la unidad, **núcleo** (predio sobre el área núcleo: una negativa pone "
               "la unidad en ROJO), LA-5 documentación completa (Especialista Predial), LA-6 expediente conforme "
               "(Especialista Legal), excepción manual NEG / OBS / EXC y alertas. Cada cambio queda en la bitácora.")
    ed = st.data_editor(v, hide_index=True, width="stretch", key="la_mz_ed",
                        disabled=[c for c in v.columns if c not in EDITABLES_MATRIZ],
                        column_config={"excepcion_manual": st.column_config.SelectboxColumn(options=[None, "NEG", "OBS", "EXC"]),
                                       "area_unidad_ha": st.column_config.NumberColumn(format="%.3f", min_value=0)})
    if st.button("Guardar cambios de la matriz", disabled=not _usuario()):
        cambios = [{"cod_predio": r["cod_predio"], **{k: (None if isinstance(r[k], float) and pd.isna(r[k]) else r[k])
                                                      for k in EDITABLES_MATRIZ}} for r in ed.to_dict("records")]
        for c in cambios:
            for k in ("nucleo", "docs_completos", "expediente_conforme"):
                c[k] = bool(c[k])
            c["excepcion_manual"] = c["excepcion_manual"] or None
        n = db.actualizar_predios(conn, cambios, _usuario())
        _flash(f"{n} predio(s) actualizados y estados recalculados.")
        st.rerun()
    # datos reservados: la matriz de circulación general no incluye DNI ni teléfonos (Ley 29733)
    ant = db.antecedentes_paso6(conn)
    if not ant.empty:
        with st.expander(f"Consulta inicial a prediantes – verificación de campo Paso 6 ({len(ant)} bloques)"):
            st.caption("Antecedente registrado en la página «ODK / KoBoToolbox» (tenencia, n.° de predios, titular, "
                       "aceptación). Sirve para priorizar la visita; no reemplaza la ficha F-LA-02.")
            codigos = sorted(set(v.cod_unidad)) if not v.empty else []
            st.dataframe(ant[ant.codigo_bloque.isin(codigos)] if codigos else ant, hide_index=True, width="stretch")
    with st.expander("Eliminar un predio (solo si no tiene fichas, inspecciones, actas ni documentos)"):
        p = st.selectbox("Predio", [""] + list(v.cod_predio), key="la_mz_del")
        if p:
            refs = db.referencias_predio(conn, p)
            if refs:
                st.info("Registros asociados: " + ", ".join(f"{k} ({n})" for k, n in refs.items())
                        + ". Elimínelos o corrija su predio en «Historial / Edición» antes de eliminar el predio.")
            elif st.button(f"Eliminar predio {p}", disabled=not _usuario(), key="la_mz_del_btn"):
                db.eliminar_predio(conn, p, _usuario())
                _flash(f"Predio {p} eliminado (restaurable desde la bitácora).")
                st.rerun()
    st.download_button("⬇️ Exportar matriz predial (Excel ANIN)",
                       ex.exportar({"Matriz_predial": ("MATRIZ PREDIAL – EXPEDIENTE PRELIMINAR DE LIBERACIÓN DE ÁREAS", v,
                                                       ["area_decl_ha", "area_unidad_ha"])}),
                       file_name=f"Matriz_Predial_IN_Piura_{date.today():%Y%m%d}.xlsx", mime=MIME_XLSX)


# ================================================================== 6. avance
def _tab_avance(conn):
    st.subheader("Avance por asistente, distrito y estado · semáforo por unidad")
    s, por_asis = rep.semaforo(conn)
    if s.empty:
        st.info("Cargue el catálogo.")
        return
    p = db.df(conn, "SELECT cod_predio, cod_unidad, estado_la FROM la_predios")
    a = db.df(conn, "SELECT asistente, fecha FROM la_actas WHERE fecha IS NOT NULL")
    if not a.empty:
        a["semana"] = pd.to_datetime(a["fecha"]).dt.strftime("%G-S%V")
        act = a.groupby(["asistente", "semana"]).size().rename("actas").reset_index()
        st.markdown("**Actas por semana y asistente**")
        st.bar_chart(act, x="semana", y="actas", color="asistente")
    st.markdown("**Resumen por asistente**")
    st.dataframe(por_asis, hide_index=True, width="stretch")
    if not p.empty:
        st.markdown("**Predios por estado LA**")
        st.dataframe(pd.crosstab(p.merge(s, left_on="cod_unidad", right_on="codigo").asistente, p.estado_la),
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
                       file_name=f"Avance_LA_IN_Piura_{date.today():%Y%m%d}.xlsx", mime=MIME_XLSX)


# ================================================================== 7. reportes
def _tab_reportes(conn, cat: dict):
    st.subheader("Reportes de Liberación de Áreas")
    s, _ = rep.semaforo(conn)
    m = rep.matriz(conn)
    if not s.empty:
        c = st.columns(5)
        c[0].metric("Unidades", len(s))
        c[1].metric("Predios", len(m))
        c[2].metric("Predios ≥ LA-4", int(m["estado_la"].isin(["LA-4", "LA-5", "LA-6"]).sum()) if not m.empty else 0)
        c[3].metric("Unidades en verde", int((s.semaforo == "VERDE").sum()))
        c[4].metric("Envíos observados", int(db.df(conn, "SELECT count(*) n FROM la_envios_raw WHERE estado_import='OBSERVADO'")["n"][0]))

    st.markdown("**1. Reporte consolidado (Excel ANIN)**")
    st.caption("Resumen, matriz predial, semáforo, avance, una hoja por formulario, documentos F-LA-05, titulares "
               "(sin DNI ni celular – Ley 29733) y envíos observados.")
    c1, c2, c3 = st.columns([1, 1, 1])
    asis = c1.selectbox("Asistente", ["Todos"] + [x for x, _ in lc.OPCIONES["asistente"] if x != "ESP"], key="la_r_asis")
    distritos = sorted(s["distrito"].dropna().unique()) if not s.empty else []
    dist = c2.selectbox("Distrito", ["Todos"] + distritos, key="la_r_dist")
    if c3.button("Generar reporte consolidado", key="la_r_gen"):
        with st.spinner("Generando…"):
            st.session_state["la_r_xls"] = rep.excel_consolidado(conn, None if asis == "Todos" else asis,
                                                                 None if dist == "Todos" else dist)
    if st.session_state.get("la_r_xls"):
        st.download_button("⬇️ Descargar reporte consolidado", st.session_state["la_r_xls"], mime=MIME_XLSX,
                           file_name=f"Reporte_Liberacion_Areas_IN_Piura_{date.today():%Y%m%d}.xlsx", key="la_r_dl")

    st.markdown("---")
    st.markdown("**2. Expediente preliminar por unidad (PDF)**")
    st.caption("Datos de la unidad, semáforo, matriz predial, socialización, inspecciones (UTM 17S), actas, "
               "búsqueda documental y observaciones por resolver.")
    con_datos = sorted(set(m["cod_unidad"])) if not m.empty else []
    todas = con_datos + sorted(u for u in cat if u not in con_datos)
    c1, c2 = st.columns([2, 1])
    u = c1.selectbox("Unidad (primero las que tienen predios)", todas, key="la_r_u",
                     format_func=lambda x: f"{x} · {cat.get(x, {}).get('label') or ''}" + (" ●" if x in con_datos else ""))
    if c2.button("Generar PDF", key="la_r_pdf") and u:
        try:
            st.session_state["la_r_pdf_b"] = (u, rep.pdf_expediente_unidad(conn, u))
        except Exception as e:  # noqa: BLE001
            st.error(f"No se pudo generar el PDF: {e}")
    if st.session_state.get("la_r_pdf_b") and st.session_state["la_r_pdf_b"][0] == u:
        st.download_button(f"⬇️ Expediente de la unidad {u} (PDF)", st.session_state["la_r_pdf_b"][1],
                           file_name=f"Expediente_LA_{u}_{date.today():%Y%m%d}.pdf", mime="application/pdf", key="la_r_pdf_dl")
    if con_datos:
        if st.button(f"Generar ZIP con los expedientes de las {len(con_datos)} unidades con predios", key="la_r_zip"):
            with st.spinner("Generando expedientes…"):
                st.session_state["la_r_zip_b"] = rep.zip_expedientes(conn, con_datos)
        if st.session_state.get("la_r_zip_b"):
            st.download_button("⬇️ Descargar expedientes (ZIP)", st.session_state["la_r_zip_b"], mime="application/zip",
                               file_name=f"Expedientes_LA_IN_Piura_{date.today():%Y%m%d}.zip", key="la_r_zip_dl")


# ================================================================== 8. vivero
def _tab_vivero(conn):
    st.subheader("Vivero Central – comparación de alternativas (F-LA-06)")
    v = db.df(conn, "SELECT * FROM la_vivero_alternativas ORDER BY puntaje DESC NULLS LAST")
    if v.empty:
        st.info("Sin alternativas registradas. Regístrelas con F-LA-06 (Kobo, aplicativo o plantilla).")
        return
    st.dataframe(v.drop(columns=["kobo_uuid"]), hide_index=True, width="stretch")
    st.caption(f"Puntaje 0–100 con pesos {core.PESOS_VIVERO} [SUPUESTO: validar con el Especialista en IV]. Sin tasación en esta fase.")
    st.download_button("⬇️ Exportar alternativas (Excel ANIN)",
                       ex.exportar({"Vivero_Central": ("VIVERO CENTRAL – ALTERNATIVAS (F-LA-06)", v.drop(columns=["kobo_uuid"]),
                                                       ["area_ha"])}),
                       file_name=f"Vivero_Central_IN_Piura_{date.today():%Y%m%d}.xlsx", mime=MIME_XLSX)


# ================================================================== 9. catálogo
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
