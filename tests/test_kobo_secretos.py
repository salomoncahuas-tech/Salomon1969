# -*- coding: utf-8 -*-
"""Lectura tolerante de KOBO_TOKEN / KOBO_SERVER / KOBO_FORM_UID desde los secrets (sin red)."""
import sys
import types
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
# database.py exige st.secrets["DATABASE_URL"] al importarse; estas pruebas no tocan la base.
try:
    import database  # noqa: F401
except Exception:  # noqa: BLE001
    sys.modules["database"] = types.ModuleType("database")
import odk_kobo as ok  # noqa: E402
from liberacion_areas import la_kobo as kb  # noqa: E402


def test_nivel_superior():
    assert ok.leer_secreto_kobo({"KOBO_TOKEN": " abc123 "}, "KOBO_TOKEN") == "abc123"


def test_mayusculas_indistintas():
    assert ok.leer_secreto_kobo({"kobo_token": "abc"}, "KOBO_TOKEN") == "abc"


def test_pegado_debajo_de_otra_seccion():
    # TOML: lo escrito debajo de [connections.x] queda dentro de esa seccion.
    s = {"DATABASE_URL": "postgresql://x", "connections": {"url": "y", "KOBO_TOKEN": "abc"}}
    assert ok.leer_secreto_kobo(s, "KOBO_TOKEN") == "abc"


def test_seccion_con_punto_anidada():
    # TOML: [connections.postgresql] crea dos niveles de anidamiento.
    s = {"DATABASE_URL": "x", "connections": {"postgresql": {"url": "y", "KOBO_TOKEN": "abc"}}}
    assert ok.leer_secreto_kobo(s, "KOBO_TOKEN") == "abc"


def test_seccion_kobo_con_alias():
    s = {"kobo": {"token": "abc", "server": "https://eu.kobotoolbox.org", "form_uid": "aAv"}}
    assert ok.leer_secreto_kobo(s, "KOBO_TOKEN") == "abc"
    assert ok.leer_secreto_kobo(s, "KOBO_SERVER") == "https://eu.kobotoolbox.org"
    assert ok.leer_secreto_kobo(s, "KOBO_FORM_UID") == "aAv"


def test_alias_solo_en_seccion_kobo():
    # «url» de otra seccion no debe tomarse como servidor de KoBo.
    assert ok.leer_secreto_kobo({"connections": {"url": "postgresql://x"}}, "KOBO_SERVER", "def") == "def"


def test_ausente_o_vacio():
    assert ok.leer_secreto_kobo({}, "KOBO_TOKEN") == ""
    assert ok.leer_secreto_kobo({"KOBO_TOKEN": "  "}, "KOBO_TOKEN", "x") == "x"
    assert ok.leer_secreto_kobo(None, "KOBO_TOKEN", "x") == "x"


@pytest.mark.parametrize("entrada", ["abc", " abc ", '"abc"', "Token abc", "token  abc"])
def test_normalizar_token(entrada):
    assert ok.normalizar_token(entrada) == "abc"


def test_cliente_la_con_token_explicito(monkeypatch):
    monkeypatch.delenv("KOBO_TOKEN", raising=False)
    c = kb.cliente(token="Token abc")
    assert c.token_api == "abc" and c.url_servidor == "https://kf.kobotoolbox.org"


def test_cliente_la_sin_token(monkeypatch):
    monkeypatch.delenv("KOBO_TOKEN", raising=False)
    monkeypatch.setattr(ok, "token_kobo", lambda: ("", ""))
    with pytest.raises(ValueError, match="KOBO_TOKEN"):
        kb.cliente()


def test_cliente_la_desde_entorno(monkeypatch):
    monkeypatch.setenv("KOBO_TOKEN", "envtok")
    assert kb.cliente().token_api == "envtok"
