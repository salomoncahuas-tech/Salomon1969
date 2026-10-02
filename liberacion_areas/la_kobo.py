# -*- coding: utf-8 -*-
"""KoboToolbox para el módulo de Liberación de Áreas.
Reutiliza el cliente del aplicativo (odk_kobo.KoBoClient: paginación completa, SSL verificado,
token enviado solo al dominio del servidor) y los mismos secrets: KOBO_TOKEN y KOBO_SERVER."""
from __future__ import annotations

import re

import pandas as pd

from . import la_core as core

SERVIDOR_POR_DEFECTO = "https://kf.kobotoolbox.org"


def cliente(servidor: str | None = None, token: str | None = None):
    """KoBoClient con los secrets del aplicativo (KOBO_SERVER, KOBO_TOKEN).

    El token se busca igual que en la página «ODK / KoBoToolbox»: nivel superior de los
    secrets, sección [kobo] o variable de entorno; si no está, el escrito en esta sesión."""
    from odk_kobo import KoBoClient, normalizar_token, secreto_kobo, token_kobo
    servidor = servidor or secreto_kobo("KOBO_SERVER", SERVIDOR_POR_DEFECTO)
    token = normalizar_token(token) if token else token_kobo()[0]
    if not token:
        raise ValueError("Falta KOBO_TOKEN en los secrets del aplicativo.")
    return KoBoClient(servidor, token)


def formularios_la(kobo) -> list[dict]:
    """Formularios desplegados cuyo nombre corresponde a la serie F-LA."""
    return [f for f in kobo.listar_formularios() if "F-LA" in f.get("nombre", "").upper()] or kobo.listar_formularios()


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
        contenido = kobo.descargar(url) if url else None
        out.append({"campo": campo, "nombre": f"{base}_{_slug(campo)}_{cont[campo]}.{ext}", "mimetype": mime,
                    "url": a.get("download_url") or url, "contenido": contenido})
    return out
