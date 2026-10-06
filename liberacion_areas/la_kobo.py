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


# «F-LA-01», «F-LA01», «F_LA_01», «FLA 01 – Reunión…» → f_la_01_reunion (F-LA-05 no tiene formulario Kobo)
_PATRON_FLA = re.compile(r"F[\s_-]*LA[\s_-]*0?([1-6])(?!\d)")
_FORM_POR_NUMERO = {"1": "f_la_01_reunion", "2": "f_la_02_titular", "3": "f_la_03_inspeccion", "4": "f_la_04_actas",
                    "6": "f_la_06_vivero"}
# Campos propios de cada formulario, para reconocer envíos de un formulario con otro nombre
_CAMPOS_CLAVE = {"f_la_01_reunion": {"tipo_evento", "fecha_evento", "unidades"},
                 "f_la_02_titular": {"consentimiento", "estado_la_propuesto", "tit_nombres"},
                 "f_la_03_inspeccion": {"punto_interior", "fecha_insp"},
                 "f_la_04_actas": {"tipo_acta", "checklist", "fecha_acta"},
                 "f_la_06_vivero": {"cod_vivero", "gps_centro"}}


def form_id_por_nombre(nombre: str | None) -> str | None:
    """Formulario del aplicativo que corresponde al nombre del formulario en KoboToolbox."""
    m = _PATRON_FLA.search(str(nombre or "").upper())
    return _FORM_POR_NUMERO.get(m.group(1)) if m else None


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
    todos = kobo.listar_formularios()
    return [f for f in todos if form_id_por_nombre(f.get("nombre"))] or todos


# ------------------------------------------------------------------ sincronización de envíos
def revisar_envios(kobo, conn) -> list[dict]:
    """Por cada formulario F-LA de la cuenta: envíos en KoboToolbox y cuáles faltan importar.
    No escribe en la base. Los envíos eliminados en el aplicativo (bitácora) no cuentan como pendientes."""
    from . import la_db as db
    ya = db.uuids_existentes(conn) | db.uuids_eliminados(conn)
    out = []
    for f in kobo.listar_formularios():
        fid = form_id_por_nombre(f.get("nombre"))
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
        fila["pendientes"] = [r for r in registros if (core.kobo_uuid(core.aplanar(r)) or "") not in ya]
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
    resumen = []
    for f in (revision if revision is not None else revisar_envios(kobo, conn)):
        fila = {"Formulario Kobo": f["nombre"], "Formulario": core.FORMULARIOS[f["form_id"]]["codigo"],
                "En Kobo": f["en_kobo"], "Importados": 0, "Observados": 0, "Adjuntos": 0, "Detalle": f["error"]}
        if f["pendientes"] and not f["error"]:
            try:
                cat, geoms = db.catalogo(conn), db.geometrias(conn)
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
