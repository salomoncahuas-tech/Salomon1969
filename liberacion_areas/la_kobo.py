# -*- coding: utf-8 -*-
"""KoboToolbox para el módulo de Liberación de Áreas.
Reutiliza el cliente del aplicativo (odk_kobo.KoBoClient: paginación completa, SSL verificado,
token enviado solo al dominio del servidor) y los mismos secrets: KOBO_TOKEN y KOBO_SERVER."""
from __future__ import annotations

import re

import pandas as pd

from . import la_core as core

SERVIDOR_POR_DEFECTO = "https://kf.kobotoolbox.org"
# Servidores oficiales de KoboToolbox (los mismos que ofrece la página «ODK / KoBoToolbox»).
# Cada cuenta y su token existen en UNO solo: un token de eu.kobotoolbox.org es inválido en kf.kobotoolbox.org.
SERVIDORES_KOBO = ("https://kf.kobotoolbox.org", "https://eu.kobotoolbox.org", "https://kobo.humanitarianresponse.info")
CLAVE_SERVIDOR_SESION = "la_kobo_servidor"


def normalizar_servidor(url: str | None) -> str:
    """'kf.kobotoolbox.org', 'https://kf.kobotoolbox.org/#/forms/aXyz' o '…/api/v2/' → 'https://kf.kobotoolbox.org'."""
    t = str(url or "").strip().strip('"').strip("'").strip()
    if not t:
        return ""
    if "://" not in t:
        t = "https://" + t
    m = re.match(r"^(https?)://([^/#?\s]+)", t, re.I)
    return f"{m.group(1).lower()}://{m.group(2).lower()}" if m else t


def servidor_actual() -> tuple[str, str]:
    """(servidor, origen): el elegido en esta sesión, el de KOBO_SERVER en los secrets o el predeterminado."""
    from odk_kobo import secreto_kobo
    try:
        import streamlit as st
        elegido = st.session_state.get(CLAVE_SERVIDOR_SESION)
    except Exception:  # noqa: BLE001 – fuera de Streamlit
        elegido = None
    if elegido:
        return normalizar_servidor(elegido), "sesion"
    configurado = secreto_kobo("KOBO_SERVER")
    if configurado:
        return normalizar_servidor(configurado), "secrets"
    return SERVIDOR_POR_DEFECTO, "predeterminado"


def cliente(servidor: str | None = None, token: str | None = None):
    """KoBoClient con los secrets del aplicativo (KOBO_SERVER, KOBO_TOKEN).

    El token se busca igual que en la página «ODK / KoBoToolbox»: nivel superior de los
    secrets, sección [kobo] o variable de entorno; si no está, el escrito en esta sesión."""
    from odk_kobo import KoBoClient, normalizar_token, token_kobo
    servidor = normalizar_servidor(servidor) or servidor_actual()[0]
    token = normalizar_token(token) if token else token_kobo()[0]
    if not token:
        raise ValueError("Falta KOBO_TOKEN en los secrets del aplicativo.")
    return KoBoClient(servidor, token)


def es_error_token(error) -> bool:
    return "Token inválido" in str(error)


def buscar_servidor(token: str, excluir: str | None = None) -> str | None:
    """Servidor oficial de KoboToolbox donde el token es válido (None si en ninguno).
    El token solo se envía a los servidores oficiales de SERVIDORES_KOBO."""
    from odk_kobo import KoBoClient
    for sv in SERVIDORES_KOBO:
        if sv == normalizar_servidor(excluir):
            continue
        try:
            if KoBoClient(sv, token, timeout=20).test_conexion()[0]:
                return sv
        except Exception:  # noqa: BLE001
            continue
    return None


def describir_token(token: str) -> str:
    """Huella del token para diagnóstico sin revelarlo: longitud y últimos 4 caracteres."""
    t = str(token or "")
    aviso = "" if len(t) == 40 else " — los tokens de KoboToolbox tienen 40 caracteres: revise que esté completo"
    return f"{len(t)} caracteres, termina en «…{t[-4:]}»{aviso}" if t else "vacío"


# ------------------------------------------------------------------ formulario vinculado (Formato de Liberación de Áreas)
# Formulario de KoboToolbox enlazado al aplicativo: la ficha técnica predial completa (FP) de Liberación de Áreas.
# Se puede cambiar sin tocar el código con KOBO_LA_FORM_URL en los secrets (enlace copiado del navegador) o, solo en
# la sesión, desde la pestaña «3 · Importar KoboToolbox». Sus envíos entran como F-LA-02 + F-LA-03 (la_fp).
FORMULARIO_VINCULADO_URL = "https://kf.kobotoolbox.org/#/forms/aCyBhgqeCTZ3UKzmBW8tx2/summary"
CLAVE_VINCULO_SESION = "la_kobo_vinculo"
_PATRON_UID_URL = re.compile(r"/forms/([A-Za-z0-9]{10,40})(?![A-Za-z0-9])")
_PATRON_UID_SOLO = re.compile(r"^a[A-Za-z0-9]{9,39}$")
# Vistas del formulario en la interfaz web de KoboToolbox
VISTAS_KOBO = {"summary": "Resumen", "landing": "Formulario", "data/table": "Datos", "settings": "Configuración"}


def leer_url_formulario(texto: str | None, servidor_defecto: str = SERVIDOR_POR_DEFECTO) -> tuple[str, str] | None:
    """(servidor, uid) de un enlace de formulario de KoboToolbox; None si no se reconoce.
    'https://kf.kobotoolbox.org/#/forms/aXyz…/summary' → ('https://kf.kobotoolbox.org', 'aXyz…').
    También acepta el UID solo (se asume el servidor por defecto) o el enlace de la API (/api/v2/assets/<uid>/)."""
    t = str(texto or "").strip().strip('"').strip("'").strip()
    if not t:
        return None
    if _PATRON_UID_SOLO.match(t):
        return normalizar_servidor(servidor_defecto), t
    m = _PATRON_UID_URL.search(t) or re.search(r"/assets/([A-Za-z0-9]{10,40})(?![A-Za-z0-9])", t)
    if not m:
        return None
    servidor = normalizar_servidor(t)
    if not re.match(r"^https?://[a-z0-9.-]+\.[a-z]{2,}(:\d+)?$", servidor):   # enlace sin dominio: '/#/forms/…'
        servidor = normalizar_servidor(servidor_defecto)
    return servidor, m.group(1)


def url_formulario(servidor: str, uid: str, vista: str = "summary") -> str:
    """Enlace a una vista del formulario en la interfaz web de KoboToolbox."""
    return f"{normalizar_servidor(servidor)}/#/forms/{uid}/{vista}"


def formulario_vinculado() -> dict:
    """{servidor, uid, url, origen}: el elegido en esta sesión, KOBO_LA_FORM_URL de los secrets o el predeterminado."""
    candidatos = []
    try:
        import streamlit as st
        candidatos.append((st.session_state.get(CLAVE_VINCULO_SESION), "sesion"))
    except Exception:  # noqa: BLE001 – fuera de Streamlit
        pass
    try:
        from odk_kobo import secreto_kobo
        candidatos.append((secreto_kobo("KOBO_LA_FORM_URL"), "secrets"))
    except Exception:  # noqa: BLE001
        pass
    candidatos.append((FORMULARIO_VINCULADO_URL, "predeterminado"))
    for texto, origen in candidatos:
        leido = leer_url_formulario(texto)
        if leido:
            servidor, uid = leido
            return {"servidor": servidor, "uid": uid, "url": url_formulario(servidor, uid), "origen": origen}
    raise AssertionError("FORMULARIO_VINCULADO_URL no es un enlace de formulario válido")


def formularios_cuenta(kobo) -> list[dict]:
    """Formularios de la cuenta más el formulario vinculado si no figura en el listado (p. ej. compartido por otra
    cuenta): se consulta directamente por su UID. Si el token no tiene permiso sobre él, se omite aquí y el panel
    del formulario vinculado explica el motivo."""
    todos = list(kobo.listar_formularios())
    v = formulario_vinculado()
    obtener = getattr(kobo, "obtener_formulario", None)
    servidor = normalizar_servidor(getattr(kobo, "url_servidor", "") or v["servidor"])
    if obtener and servidor == v["servidor"] and v["uid"] not in {f.get("uid") for f in todos}:
        try:
            todos.append(obtener(v["uid"]))
        except Exception:  # noqa: BLE001 – sin permiso o sin red
            pass
    return todos


def _pendientes(form_id: str, registros: list[dict], ya: set[str]) -> list[dict]:
    """Envíos aún no importados. Una ficha FP está pendiente si falta su F-LA-02 o su F-LA-03."""
    from . import la_fp as fp
    if form_id == fp.FORM_ID:
        return [r for r in registros if not set(fp.uuids_partes(core.kobo_uuid(core.aplanar(r)))) <= ya]
    return [r for r in registros if (core.kobo_uuid(core.aplanar(r)) or "") not in ya]


def estado_vinculado(kobo, conn) -> dict:
    """Estado del formulario vinculado: datos del formulario en KoboToolbox, envíos y pendientes de importar.
    No escribe en la base. La fila `revision` sirve tal cual para sincronizar_envios."""
    from . import la_db as db
    v = formulario_vinculado()
    out = {**v, "nombre": "", "desplegado": None, "envios": None, "ultimo_envio": "", "propietario": "",
           "enlaces_llenado": {}, "importados": 0, "pendientes": 0, "revision": None, "error": "", "aviso": ""}
    servidor = normalizar_servidor(getattr(kobo, "url_servidor", "") or v["servidor"])
    if servidor != v["servidor"]:
        out["error"] = (f"El formulario vinculado está en {v['servidor']}, pero el aplicativo consulta {servidor}. "
                        f"Elija {v['servidor']} como servidor (o KOBO_SERVER en los secrets) con un token de esa cuenta.")
        return out
    try:
        info = kobo.obtener_formulario(v["uid"])
    except Exception as e:  # noqa: BLE001
        out["error"] = (f"No se pudo leer el formulario: {e}. Verifique que la cuenta del token sea la propietaria del "
                        "formulario o que el formulario esté compartido con ella (KoboToolbox › Compartir › «Ver envíos»)."
                        if es_error_token(e) or "404" in str(e) else f"No se pudo leer el formulario: {e}")
        return out
    links = info.get("enlaces") or {}
    out.update({"nombre": info.get("nombre", ""), "desplegado": bool(info.get("desplegado")),
                "envios": info.get("envios"), "ultimo_envio": info.get("ultimo_envio", ""),
                "propietario": info.get("propietario", ""),
                "enlaces_llenado": {et: links[k] for k, et in ENLACES_LLENADO if links.get(k)}})
    fid = form_id_kobo(out["nombre"], v["uid"])
    if not out["desplegado"] and not out["envios"]:
        out["aviso"] = "El formulario no está desplegado en KoboToolbox: despliéguelo para recibir envíos."
        return out
    try:
        registros = kobo.obtener_envios(v["uid"])
    except Exception as e:  # noqa: BLE001
        out["error"] = f"No se pudieron leer los envíos: {e}"
        return out
    ya = db.uuids_existentes(conn) | db.uuids_eliminados(conn)
    pend = _pendientes(fid, registros, ya)
    out.update({"envios": len(registros), "pendientes": len(pend), "importados": len(registros) - len(pend),
                "form_id": fid,
                "revision": {"uid": v["uid"], "nombre": out["nombre"], "form_id": fid, "en_kobo": len(registros),
                             "pendientes": pend, "error": ""}})
    return out


# «F-LA-01», «F-LA01», «F_LA_01», «FLA 01 – Reunión…» → f_la_01_reunion (F-LA-05 no tiene formulario Kobo)
_PATRON_FLA = re.compile(r"F[\s_-]*LA[\s_-]*0?([1-6])(?!\d)")
_FORM_POR_NUMERO = {"1": "f_la_01_reunion", "2": "f_la_02_titular", "3": "f_la_03_inspeccion", "4": "f_la_04_actas",
                    "6": "f_la_06_vivero"}
# Campos propios de cada formulario, para reconocer envíos de un formulario con otro nombre
_CAMPOS_CLAVE = {"f_la_01_reunion": {"tipo_evento", "fecha_evento", "unidades"},
                 "f_la_02_titular": {"consentimiento", "estado_la_propuesto", "tit_nombres"},
                 "f_la_03_inspeccion": {"punto_interior", "fecha_insp"},
                 "f_la_04_actas": {"tipo_acta", "checklist", "fecha_acta"},
                 "f_la_06_vivero": {"cod_vivero", "gps_centro"},
                 # ficha técnica predial completa (F-LA-02 + F-LA-03): campos que solo tiene la ficha
                 "fp_ficha_predial": {"personeria", "predio_provincia", "metodo_coord", "col_norte", "dom_direccion",
                                      "tipo_predio"}}


def form_id_por_nombre(nombre: str | None) -> str | None:
    """Formulario del aplicativo que corresponde al nombre del formulario en KoboToolbox."""
    m = _PATRON_FLA.search(str(nombre or "").upper())
    return _FORM_POR_NUMERO.get(m.group(1)) if m else None


def form_id_kobo(nombre: str | None, uid: str | None = None) -> str | None:
    """Como form_id_por_nombre, reconociendo además la ficha técnica predial completa (FP) por título o UID.
    El formulario vinculado (Formato de Liberación de Áreas) es la ficha FP, salvo que su nombre indique un F-LA-0x."""
    from . import la_fp as fp
    if fp.es_formulario_fp(nombre, uid):
        return fp.FORM_ID
    fid = form_id_por_nombre(nombre)
    if not fid and uid and uid == formulario_vinculado()["uid"]:
        return fp.FORM_ID
    return fid


def codigo_formulario(form_id: str) -> str:
    from . import la_fp as fp
    return fp.CODIGO if form_id == fp.FORM_ID else core.FORMULARIOS[form_id]["codigo"]


def form_id_por_campos(registros: list[dict]) -> str | None:
    """Formulario cuyos campos propios aparecen en los envíos (exportación sin nombre de formulario)."""
    claves: set[str] = set()
    for r in registros[:50]:
        claves |= set(core.aplanar(r))
    votos = {fid: len(campos & claves) for fid, campos in _CAMPOS_CLAVE.items()}
    mejor = max(votos, key=votos.get)
    return mejor if votos[mejor] else None


def formularios_la(kobo) -> list[dict]:
    """Formularios desplegados cuyo nombre corresponde a la serie F-LA (si no hay ninguno, todos)."""
    todos = formularios_cuenta(kobo)
    return [f for f in todos if form_id_kobo(f.get("nombre"), f.get("uid"))] or todos


# ------------------------------------------------------------------ enlaces de llenado (Enketo)
# Claves de deployment__links del asset (API v2). Los enlaces de envíos múltiples dejan el formulario en blanco y
# vuelven al inicio después de «Enviar»; el de envío único (single_url, /single/…) termina en la página de
# agradecimiento y obliga a recargar para registrar otro prediante.
ENLACES_LLENADO = (("offline_url", "En línea-Fuera de línea (envíos múltiples)"),
                   ("url", "En línea (envíos múltiples)"))


def enlaces_llenado(kobo, uid: str) -> dict[str, str]:
    """{etiqueta: url} de los enlaces web de envíos múltiples de un formulario desplegado ({} si no está desplegado)."""
    import urllib.parse
    a = kobo._get_json(f"assets/{urllib.parse.quote(uid, safe='')}/")
    links = a.get("deployment__links") or {}
    return {et: links[k] for k, et in ENLACES_LLENADO if links.get(k)}


def enlaces_formularios_la(kobo) -> list[dict]:
    """Enlaces de registro continuo de cada formulario F-LA desplegado. Un formulario que falla no detiene el resto."""
    out = []
    for f in formularios_cuenta(kobo):
        fid = form_id_kobo(f.get("nombre"), f.get("uid"))
        if not fid or not f.get("desplegado"):
            continue
        fila = {"form_id": fid, "nombre": f.get("nombre"), "enlaces": {}, "error": ""}
        try:
            fila["enlaces"] = enlaces_llenado(kobo, f["uid"])
        except Exception as e:  # noqa: BLE001 – se informa por formulario
            fila["error"] = str(e)
        out.append(fila)
    return sorted(out, key=lambda x: (x["form_id"], str(x["nombre"])))


# ------------------------------------------------------------------ sincronización de envíos
def revisar_envios(kobo, conn) -> list[dict]:
    """Por cada formulario F-LA de la cuenta: envíos en KoboToolbox y cuáles faltan importar.
    No escribe en la base. Los envíos eliminados en el aplicativo (bitácora) no cuentan como pendientes."""
    from . import la_db as db
    ya = db.uuids_existentes(conn) | db.uuids_eliminados(conn)
    out = []
    for f in formularios_cuenta(kobo):
        fid = form_id_kobo(f.get("nombre"), f.get("uid"))
        if not fid or not (f.get("desplegado") or f.get("envios")):
            continue
        fila = {"uid": f["uid"], "nombre": f.get("nombre"), "form_id": fid, "en_kobo": 0, "pendientes": [], "error": ""}
        try:
            registros = kobo.obtener_envios(f["uid"])
        except Exception as e:  # noqa: BLE001 – se informa por formulario
            fila["error"] = str(e)
            out.append(fila)
            continue
        fila["en_kobo"] = len(registros)
        fila["pendientes"] = _pendientes(fid, registros, ya)
        out.append(fila)
    return out


def guardar_adjuntos(kobo, conn, resultados: list) -> tuple[int, int]:
    """Descarga y guarda fotos y actas de los envíos importados. Un adjunto que falla no detiene el resto.
    Devuelve (guardados, fallidos)."""
    from . import la_db as db
    ok = fallos = 0
    for r in resultados:
        if r.estado_import == "DUPLICADO":
            continue
        try:
            adjuntos = descargar_adjuntos(kobo, r.datos)
        except Exception:  # noqa: BLE001
            fallos += 1
            continue
        for a in adjuntos:
            if a["contenido"] is None:
                fallos += 1
            db.guardar_adjunto(conn, r.kobo_uuid, r.cod_predio, a["campo"], a["nombre"], a["mimetype"], a["url"],
                               a["contenido"])
            ok += a["contenido"] is not None
    return ok, fallos


def sincronizar_envios(kobo, conn, usuario: str, revision: list[dict] | None = None, fotos: bool = True) -> list[dict]:
    """Importa los envíos pendientes de todos los formularios F-LA (mismas validaciones que la importación manual).
    Devuelve un resumen por formulario."""
    from . import la_db as db
    from . import la_fp as fp
    resumen = []
    for f in (revision if revision is not None else revisar_envios(kobo, conn)):
        fila = {"Formulario Kobo": f["nombre"], "Formulario": codigo_formulario(f["form_id"]),
                "En Kobo": f["en_kobo"], "Importados": 0, "Observados": 0, "Adjuntos": 0, "Detalle": f["error"]}
        if f["pendientes"] and not f["error"]:
            try:
                cat, geoms = db.catalogo(conn), db.geometrias(conn)
                if f["form_id"] == fp.FORM_ID:     # cada ficha FP → F-LA-02 + F-LA-03 (sin las partes eliminadas)
                    ya, sus = db.uuids_existentes(conn), db.sus_areas(conn, cat)
                    omitir = ya | db.uuids_eliminados(conn)
                    res = [x for r in f["pendientes"]
                           for x in fp.validar(core.aplanar(r), cat, geoms, ya, sus, omitir=omitir)]
                    n = fp.importar(conn, res, "API", usuario)
                else:
                    res = [core.validar_envio(f["form_id"], core.aplanar(r), cat, geoms, db.uuids_existentes(conn),
                                              db.sus_areas(conn, cat)) for r in f["pendientes"]]
                    n = db.importar(conn, res, "API", f["form_id"], usuario)["insertados"]
                fila["Importados"] = n
                fila["Observados"] = sum(r.estado_import == "OBSERVADO" for r in res)
                if fotos:
                    ok, fallos = guardar_adjuntos(kobo, conn, res)
                    fila["Adjuntos"] = ok
                    if fallos:
                        fila["Detalle"] = f"{fallos} adjunto(s) no se pudieron descargar"
            except Exception as e:  # noqa: BLE001
                conn.rollback()
                fila["Detalle"] = f"Error: {e}"
        resumen.append(fila)
    return resumen


def leer_exportacion_xlsx(archivo) -> list[dict]:
    """Exportación XLSX de Kobo ('XML values and headers', grupos con separador '/').
    Hoja 1 = envíos; hojas siguientes = repeats (unidas por _parent_index)."""
    hojas = pd.read_excel(archivo, sheet_name=None, dtype=str)
    nombres = list(hojas)
    principal = hojas[nombres[0]].where(pd.notna(hojas[nombres[0]]), None).to_dict("records")
    for r in principal:
        r["_index"] = int(float(r["_index"])) if r.get("_index") else None
    repeats = {}
    for n in nombres[1:]:
        filas = hojas[n].where(pd.notna(hojas[n]), None).to_dict("records")
        for f in filas:
            f["_parent_index"] = int(float(f["_parent_index"])) if f.get("_parent_index") else None
        repeats[n.split("/")[-1]] = [core.aplanar(f) for f in filas]
    return core.unir_repeats_xlsx(principal, repeats)


def _slug(t: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]", "_", t or "")


def descargar_adjuntos(kobo, registro: dict) -> list[dict]:
    """Descarga fotos y páginas de actas (versión mediana, como el Paso 6) con nombre
    {cod_predio}_{campo}_{n}.{ext}. Devuelve dicts listos para la_db.guardar_adjunto."""
    base = _slug(registro.get("cod_predio") or registro.get("cod_unidad") or registro.get("cod_vivero") or "SIN_COD")
    cont: dict[str, int] = {}
    out = []
    for a in registro.get("_attachments", []) or []:
        campo = (a.get("question_xpath") or a.get("filename", "adjunto")).split("/")[-1]
        cont[campo] = cont.get(campo, 0) + 1
        mime = a.get("mimetype", "image/jpeg")
        ext = (mime.split("/")[-1] or "jpg").replace("jpeg", "jpg")
        url = a.get("download_medium_url") or a.get("download_large_url") or a.get("download_url")
        try:
            contenido = kobo.descargar(url) if url else None
        except Exception:  # noqa: BLE001 – queda la URL del original; se informa como adjunto fallido
            contenido = None
        out.append({"campo": campo, "nombre": f"{base}_{_slug(campo)}_{cont[campo]}.{ext}", "mimetype": mime,
                    "url": a.get("download_url") or url, "contenido": contenido})
    return out
