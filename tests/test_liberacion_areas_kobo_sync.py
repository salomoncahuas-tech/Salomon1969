# -*- coding: utf-8 -*-
"""
Sincronización de envíos KoboToolbox → Liberación de Áreas (aviso «envíos aún no importados» e «Importar ahora»).
Cliente Kobo SIMULADO (sin red) y datos SINTÉTICOS.
- Lógica pura: `python -m pytest tests/test_liberacion_areas_kobo_sync.py -q`
- Integración PostgreSQL: LA_TEST_DB="postgresql://…" (base VACÍA de pruebas: se borra el esquema public).
"""
import copy
import os
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
sys.path.insert(0, str(RAIZ / "tests"))
from liberacion_areas import la_core as core  # noqa: E402
from liberacion_areas import la_kobo as kb  # noqa: E402
import test_liberacion_areas as base  # noqa: E402


class KoboFalso:
    """Imita odk_kobo.KoBoClient: listar_formularios, obtener_envios y descargar."""

    def __init__(self, formularios: list[dict], envios: dict[str, list[dict]]):
        self.formularios, self.envios, self.descargas = formularios, envios, 0

    def listar_formularios(self):
        return self.formularios

    def obtener_envios(self, uid):
        return copy.deepcopy(self.envios.get(uid, []))

    def descargar(self, url):
        self.descargas += 1
        if "falla" in url:
            raise ConnectionError("Error HTTP 404")
        return b"\xff\xd8foto"


def _envio_la01(uuid: str) -> dict:
    """Envío F-LA-01 tal como lo entrega la API v2 de Kobo (con metadatos y adjuntos)."""
    r = copy.deepcopy(base.casos()["f_la_01_reunion"][0])
    r.update({"meta/instanceID": f"uuid:{uuid}", "_uuid": uuid, "_id": 1, "formhub/uuid": "x", "__version__": "v1",
              "_submission_time": "2026-10-05T14:31:02", "_geolocation": [-5.15, -79.91], "_tags": [], "_notes": [],
              "_validation_status": {}, "_attachments": [
                  {"question_xpath": "g_evidencias/foto_evento", "mimetype": "image/jpeg",
                   "download_url": f"https://kf.kobotoolbox.org/{uuid}.jpg",
                   "download_medium_url": f"https://kf.kobotoolbox.org/{uuid}_m.jpg"}]})
    return r


def _kobo_dos_envios() -> KoboFalso:
    forms = [{"uid": "aLA01", "nombre": "F-LA01 Reunión informativa", "envios": 2, "desplegado": True},
             {"uid": "aP6", "nombre": "Paso 6 – Verificación de campo", "envios": 40, "desplegado": True},
             {"uid": "aLA03", "nombre": "F-LA-03 – Inspección", "envios": 0, "desplegado": True}]
    return KoboFalso(forms, {"aLA01": [_envio_la01("s-001"), _envio_la01("s-002")], "aLA03": []})


# ------------------------------------------------------------------ lógica pura
def test_form_id_por_nombre():
    casos = {"F-LA01": "f_la_01_reunion", "F-LA-01 – Reunión / asamblea informativa": "f_la_01_reunion",
             "f_la_02 ficha predial": "f_la_02_titular", "FLA 03": "f_la_03_inspeccion", "F-LA-04 actas": "f_la_04_actas",
             "F-LA-06 Vivero": "f_la_06_vivero", "F-LA-05 documentos": None, "Paso 6 – Verificación": None,
             "F-LA-010": None, "": None, None: None}
    for nombre, esperado in casos.items():
        assert kb.form_id_por_nombre(nombre) == esperado, nombre


def test_form_id_por_campos():
    c = base.casos()
    for fid in ("f_la_01_reunion", "f_la_02_titular", "f_la_03_inspeccion", "f_la_04_actas", "f_la_06_vivero"):
        assert kb.form_id_por_campos(c[fid]) == fid, fid
    assert kb.form_id_por_campos([{"_uuid": "x", "otro": 1}]) is None


def test_formularios_la_reconoce_f_la01():
    k = _kobo_dos_envios()
    assert [f["uid"] for f in kb.formularios_la(k)] == ["aLA01", "aLA03"]


def test_unidad_f_la_01_sin_calculo():
    """Un F-LA-01 sin el cálculo cod_unidad toma la unidad de «unidades» (Historial y filtros por unidad)."""
    r = _envio_la01("s-009")
    r.pop("cod_unidad")
    res = core.validar_envio("f_la_01_reunion", core.aplanar(r), base.catalogo(), base.GEOMS, set())
    assert res.cod_unidad == "27 3" and res.estado_import == "NUEVO", res.motivos


def test_adjunto_fallido_no_detiene():
    k = KoboFalso([], {})
    r = core.aplanar(_envio_la01("s-010"))
    r["_attachments"].append({"question_xpath": "g/acta", "mimetype": "image/jpeg",
                              "download_medium_url": "https://kf.kobotoolbox.org/falla.jpg"})
    adj = kb.descargar_adjuntos(k, r)
    assert [a["contenido"] is not None for a in adj] == [True, False]


# ------------------------------------------------------------------ integración PostgreSQL
def test_integracion_sincronizacion():
    url = os.environ.get("LA_TEST_DB")
    if not url:
        return
    from liberacion_areas import la_db as db
    base.test_integracion_postgres()               # esquema limpio + catálogo V6 enlazado a bloques
    conn = db.conectar(url)
    try:
        conn.rollback()
        with conn.cursor() as cur:
            cur.execute("DELETE FROM la_adjuntos; DELETE FROM la_reuniones; DELETE FROM la_envios_raw "
                        "WHERE form_id = 'f_la_01_reunion'")
        conn.commit()
        k = _kobo_dos_envios()
        rev = kb.revisar_envios(k, conn)
        assert [(f["form_id"], f["en_kobo"], len(f["pendientes"])) for f in rev] == \
               [("f_la_01_reunion", 2, 2), ("f_la_03_inspeccion", 0, 0)]

        res = kb.sincronizar_envios(k, conn, "prueba", revision=rev)
        assert res[0]["Importados"] == 2 and res[0]["Adjuntos"] == 2, res
        env = db.listar_envios(conn, "f_la_01_reunion")
        assert sorted(env.kobo_uuid) == ["s-001", "s-002"]
        assert set(env.origen) == {"KOBO"} and set(env.cod_unidad) == {"27 3"}
        assert db.df(conn, "SELECT count(*) AS n FROM la_reuniones")["n"][0] == 2

        # Segunda revisión: nada pendiente; volver a sincronizar no duplica
        assert sum(len(f["pendientes"]) for f in kb.revisar_envios(k, conn)) == 0
        assert kb.sincronizar_envios(k, conn, "prueba")[0]["Importados"] == 0

        # Un envío eliminado en el aplicativo no vuelve a aparecer como pendiente
        db.eliminar_envio(conn, "s-001", "prueba")
        assert sum(len(f["pendientes"]) for f in kb.revisar_envios(k, conn)) == 0
        assert sorted(db.listar_envios(conn, "f_la_01_reunion").kobo_uuid) == ["s-002"]

        # Un envío nuevo en Kobo sí aparece y se importa
        k.envios["aLA01"].append(_envio_la01("s-003"))
        rev = kb.revisar_envios(k, conn)
        assert [core.kobo_uuid(core.aplanar(r)) for r in rev[0]["pendientes"]] == ["s-003"]
        assert kb.sincronizar_envios(k, conn, "prueba", revision=rev)[0]["Importados"] == 1
    finally:
        conn.close()


# ------------------------------------------------------------------ servidor y token
def test_normalizar_servidor():
    casos = {"https://kf.kobotoolbox.org": "https://kf.kobotoolbox.org",
             "https://kf.kobotoolbox.org/": "https://kf.kobotoolbox.org",
             "eu.kobotoolbox.org": "https://eu.kobotoolbox.org",
             " 'https://EU.kobotoolbox.org/#/forms/aXyz/summary' ": "https://eu.kobotoolbox.org",
             "https://kf.kobotoolbox.org/api/v2/": "https://kf.kobotoolbox.org", "": "", None: ""}
    for entrada, esperado in casos.items():
        assert kb.normalizar_servidor(entrada) == esperado, entrada


def test_servidor_desde_secrets(monkeypatch):
    import odk_kobo as ok
    monkeypatch.setattr(ok, "secreto_kobo", lambda nombre, defecto="": "eu.kobotoolbox.org/" if nombre == "KOBO_SERVER" else defecto)
    assert kb.servidor_actual() == ("https://eu.kobotoolbox.org", "secrets")
    assert kb.cliente(token="abc").url_servidor == "https://eu.kobotoolbox.org"


def test_buscar_servidor_del_token(monkeypatch):
    """Un token de eu.kobotoolbox.org rechazado en kf se encuentra en eu; uno inválido en todos → None."""
    import odk_kobo as ok
    monkeypatch.setattr(ok.KoBoClient, "test_conexion",
                        lambda self: (self.url_servidor == "https://eu.kobotoolbox.org" and self.token_api == "tok-eu", ""))
    assert kb.buscar_servidor("tok-eu", excluir="https://kf.kobotoolbox.org") == "https://eu.kobotoolbox.org"
    assert kb.buscar_servidor("malo", excluir="https://kf.kobotoolbox.org") is None


def test_describir_token_sin_revelarlo():
    t = "a" * 36 + "b9f2"
    assert kb.describir_token(t) == "40 caracteres, termina en «…b9f2»"
    assert "revise que esté completo" in kb.describir_token("abc123")
    assert kb.describir_token("") == "vacío"
    assert kb.es_error_token(ConnectionError("Token inválido o sin permiso sobre el formulario (HTTP 401 en kf)"))
