# -*- coding: utf-8 -*-
"""
Formulario vinculado de KoboToolbox («Formato de Liberación de Áreas», kf.kobotoolbox.org, UID aCyBhgqeCTZ3UKzmBW8tx2).
Cliente Kobo SIMULADO (sin red) y datos SINTÉTICOS.
`python -m pytest tests/test_la_kobo_vinculo.py -q`
"""
import copy
import json
import sys
from pathlib import Path

import types

import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
# database.py exige st.secrets["DATABASE_URL"] al importarse (lo usa odk_kobo); estas pruebas no tocan esa base.
try:
    import database  # noqa: F401
except Exception:  # noqa: BLE001
    sys.modules["database"] = types.ModuleType("database")
from liberacion_areas import la_db as db  # noqa: E402
from liberacion_areas import la_fp as fp  # noqa: E402
from liberacion_areas import la_kobo as kb  # noqa: E402

UID = "aCyBhgqeCTZ3UKzmBW8tx2"


@pytest.fixture(autouse=True)
def _sin_secrets(monkeypatch):
    monkeypatch.delenv("KOBO_LA_FORM_URL", raising=False)
    monkeypatch.delenv("KOBO_FP_UID", raising=False)


class KoboFalso:
    """Imita odk_kobo.KoBoClient: listar_formularios, obtener_formulario y obtener_envios."""

    def __init__(self, listado, assets, envios, url_servidor="https://kf.kobotoolbox.org"):
        self.listado, self.assets, self.envios, self.url_servidor = listado, assets, envios, url_servidor

    def listar_formularios(self):
        return copy.deepcopy(self.listado)

    def obtener_formulario(self, uid):
        if uid not in self.assets:
            raise ConnectionError("Token inválido o sin permiso sobre el formulario (HTTP 403 en kf.kobotoolbox.org)")
        return copy.deepcopy(self.assets[uid])

    def obtener_envios(self, uid):
        return copy.deepcopy(self.envios.get(uid, []))


def _asset(uid=UID, nombre="Formato de Liberación de Áreas – IN Piura", envios=3):
    return {"uid": uid, "nombre": nombre, "envios": envios, "desplegado": True, "fecha_modificacion": "",
            "ultimo_envio": "2026-11-02T15:04:05Z", "propietario": "anin_in_piura",
            "enlaces": {"offline_url": "https://ee.kobotoolbox.org/x/abc", "url": "https://ee.kobotoolbox.org/x/def",
                        "single_url": "https://ee.kobotoolbox.org/single/ghi"}}


def _envio(uuid):
    return {"_uuid": uuid, "meta/instanceID": f"uuid:{uuid}"}


# ------------------------------------------------------------------ enlace del formulario
@pytest.mark.parametrize("texto, esperado", [
    (f"https://kf.kobotoolbox.org/#/forms/{UID}/summary", ("https://kf.kobotoolbox.org", UID)),
    (f"kf.kobotoolbox.org/#/forms/{UID}/data/table", ("https://kf.kobotoolbox.org", UID)),
    (f"  {UID}  ", ("https://kf.kobotoolbox.org", UID)),
    (f"https://eu.kobotoolbox.org/api/v2/assets/{UID}/?format=json", ("https://eu.kobotoolbox.org", UID)),
    (f"/#/forms/{UID}/landing", ("https://kf.kobotoolbox.org", UID)),
    ("https://kf.kobotoolbox.org/", None),
    ("formulario de liberación", None),
    ("", None),
])
def test_leer_url_formulario(texto, esperado):
    assert kb.leer_url_formulario(texto) == esperado


def test_formulario_vinculado_predeterminado_y_secrets(monkeypatch):
    v = kb.formulario_vinculado()
    assert (v["servidor"], v["uid"], v["origen"]) == ("https://kf.kobotoolbox.org", UID, "predeterminado")
    assert v["url"] == f"https://kf.kobotoolbox.org/#/forms/{UID}/summary"
    assert kb.url_formulario(v["servidor"], UID, "data/table") == f"https://kf.kobotoolbox.org/#/forms/{UID}/data/table"
    monkeypatch.setenv("KOBO_LA_FORM_URL", "https://kf.kobotoolbox.org/#/forms/aNuevoFormulario123/summary")
    v = kb.formulario_vinculado()
    assert (v["uid"], v["origen"]) == ("aNuevoFormulario123", "secrets")
    monkeypatch.setenv("KOBO_LA_FORM_URL", "enlace mal copiado")       # inválido → se usa el predeterminado
    assert kb.formulario_vinculado()["uid"] == UID


def test_form_id_del_formulario_vinculado(monkeypatch):
    assert kb.form_id_kobo("Formato de Liberación de Áreas", UID) == fp.FORM_ID
    monkeypatch.setenv("KOBO_LA_FORM_URL", "aOtroVinculado12345")
    assert kb.form_id_kobo("Formato de Liberación de Áreas", "aOtroVinculado12345") == fp.FORM_ID
    assert kb.form_id_kobo("F-LA-01 Reunión informativa", "aOtroVinculado12345") == "f_la_01_reunion"
    assert kb.form_id_kobo("Encuesta ajena", "aCualquiera1234567") is None


# ------------------------------------------------------------------ formularios de la cuenta
def test_formularios_cuenta_agrega_el_vinculado_compartido():
    kobo = KoboFalso([{"uid": "aOtro", "nombre": "Encuesta ajena", "envios": 1, "desplegado": True}],
                     {UID: _asset()}, {})
    uids = [f["uid"] for f in kb.formularios_cuenta(kobo)]
    assert uids == ["aOtro", UID]
    # ya listado: no se duplica
    kobo.listado.append({"uid": UID, "nombre": "x", "envios": 0, "desplegado": True})
    assert [f["uid"] for f in kb.formularios_cuenta(kobo)].count(UID) == 1
    # sin permiso o en otro servidor: se omite sin error
    assert [f["uid"] for f in kb.formularios_cuenta(KoboFalso([], {}, {}))] == []
    assert kb.formularios_cuenta(KoboFalso([], {UID: _asset()}, {}, url_servidor="https://eu.kobotoolbox.org")) == []


def test_revisar_envios_incluye_el_vinculado_no_listado(monkeypatch):
    monkeypatch.setattr(db, "uuids_existentes", lambda conn: {"fp-1-F02", "fp-1-F03", "fp-2-F02"})
    monkeypatch.setattr(db, "uuids_eliminados", lambda conn: set())
    kobo = KoboFalso([], {UID: _asset()}, {UID: [_envio("fp-1"), _envio("fp-2"), _envio("fp-3")]})
    rev = kb.revisar_envios(kobo, None)
    assert [(f["uid"], f["form_id"], f["en_kobo"]) for f in rev] == [(UID, fp.FORM_ID, 3)]
    # fp-2 tiene solo su F-LA-02: sigue pendiente
    assert sorted(r["_uuid"] for r in rev[0]["pendientes"]) == ["fp-2", "fp-3"]
    assert [f["uid"] for f in kb.formularios_la(kobo)] == [UID]


# ------------------------------------------------------------------ estado del formulario vinculado
def test_estado_vinculado(monkeypatch):
    monkeypatch.setattr(db, "uuids_existentes", lambda conn: {"fp-1-F02", "fp-1-F03"})
    monkeypatch.setattr(db, "uuids_eliminados", lambda conn: {"fp-9-F02", "fp-9-F03"})
    kobo = KoboFalso([], {UID: _asset()}, {UID: [_envio("fp-1"), _envio("fp-2"), _envio("fp-9")]})
    est = kb.estado_vinculado(kobo, None)
    assert not est["error"] and est["desplegado"] and est["propietario"] == "anin_in_piura"
    assert (est["envios"], est["importados"], est["pendientes"]) == (3, 2, 1)
    assert list(est["enlaces_llenado"]) == [et for _, et in kb.ENLACES_LLENADO]     # sin el enlace /single/
    assert "single" not in json.dumps(est["enlaces_llenado"])
    rev = est["revision"]
    assert rev["form_id"] == fp.FORM_ID and [r["_uuid"] for r in rev["pendientes"]] == ["fp-2"]
    assert set(rev) >= {"uid", "nombre", "form_id", "en_kobo", "pendientes", "error"}   # fila de sincronizar_envios


def test_estado_vinculado_errores(monkeypatch):
    monkeypatch.setattr(db, "uuids_existentes", lambda conn: set())
    monkeypatch.setattr(db, "uuids_eliminados", lambda conn: set())
    sin_permiso = kb.estado_vinculado(KoboFalso([], {}, {}), None)
    assert "Compartir" in sin_permiso["error"] and sin_permiso["revision"] is None
    otro_servidor = kb.estado_vinculado(KoboFalso([], {UID: _asset()}, {}, "https://eu.kobotoolbox.org"), None)
    assert "kf.kobotoolbox.org" in otro_servidor["error"]
    a = _asset(envios=0)
    a["desplegado"] = False
    no_desplegado = kb.estado_vinculado(KoboFalso([], {UID: a}, {}), None)
    assert no_desplegado["aviso"] and not no_desplegado["error"] and no_desplegado["pendientes"] == 0


def test_cliente_obtener_formulario(monkeypatch):
    from odk_kobo import KoBoClient
    c = KoBoClient("https://kf.kobotoolbox.org", "t" * 40)
    pedido = {}

    def _get_json(endpoint, params=None):
        pedido["endpoint"] = endpoint
        return {"uid": UID, "name": "Formato LA", "asset_type": "survey", "has_deployment": True,
                "deployment__submission_count": 7, "deployment__last_submission_time": "2026-11-01T10:00:00Z",
                "owner__username": "anin", "deployment__links": {"offline_url": "https://ee.kobotoolbox.org/x/abc"}}

    monkeypatch.setattr(c, "_get_json", _get_json)
    f = c.obtener_formulario(UID)
    assert pedido["endpoint"] == f"assets/{UID}/"
    assert (f["uid"], f["nombre"], f["envios"], f["desplegado"], f["propietario"]) == (UID, "Formato LA", 7, True, "anin")
    assert f["enlaces"]["offline_url"].startswith("https://ee.kobotoolbox.org/")
