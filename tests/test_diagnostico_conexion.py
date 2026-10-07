# -*- coding: utf-8 -*-
"""Diagnóstico del error de conexión a Supabase (pantalla «No se pudo conectar a la base de datos»).
La contraseña nunca debe aparecer en el mensaje."""
import sys
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
# database.py exige st.secrets["DATABASE_URL"] al importarse; estas pruebas no tocan la base.
st.secrets._secrets = {"DATABASE_URL": "postgresql://x:y@localhost:5432/postgres"}
import database as db  # noqa: E402

URL = "postgresql://postgres.maiizkcpepuwlevbxlxw:Cl4veSecreta99@aws-0-us-west-2.pooler.supabase.com:6543/postgres"


def _diag(msg):
    return db.diagnosticar_error_conexion(Exception(msg), URL)


def test_describir_url_oculta_contrasena():
    d = db.describir_url(URL)
    assert d == ("postgres.maiizkcpepuwlevbxlxw@aws-0-us-west-2.pooler.supabase.com:6543/postgres "
                 "(contraseña: 14 caracteres)")
    assert "Cl4ve" not in d


def test_causas():
    casos = {
        'FATAL:  password authentication failed for user "postgres"': "no coincide",
        "FATAL:  Circuit breaker open: Too many authentication errors": "bloqueó temporalmente",
        "FATAL:  Tenant or user not found": "postgres.<código del proyecto>",
        'could not translate host name "aws-0-us-west-2.pooler.supabse.com" to address': "mal escrito",
        'invalid dsn: missing "=" after "x" in connection info string': "formato no válido",
        "connection to server timed out": "pausado",
    }
    for msg, esperado in casos.items():
        assert esperado in _diag(msg)["causa"], msg


def test_contrasena_nunca_en_detalle():
    d = _diag("invalid dsn: postgresql://postgres:Cl4veSecreta99@host")
    assert "Cl4veSecreta99" not in str(d) and "****" in d["detalle"]
