# -*- coding: utf-8 -*-
"""
Pruebas del XLSForm de la Ficha Técnica Predial completa (liberacion_areas/la_xlsform_ficha.py).
`python -m pytest tests/test_xlsform_ficha.py -q`  (requiere pyxform, el mismo conversor de KoboToolbox)
"""
import sys
from collections import Counter
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
from liberacion_areas import la_campos as lc  # noqa: E402
from liberacion_areas import la_ficha_predial as fp  # noqa: E402
from liberacion_areas import la_xlsform_ficha as xf  # noqa: E402


def _survey():
    return xf.hojas()["survey"][1]


def _choices():
    return xf.hojas()["choices"][1]


def test_convierte_con_pyxform(tmp_path):
    pyxform = pytest.importorskip("pyxform.xls2xform")
    ruta = xf.generar_xlsform(tmp_path / xf.ARCHIVO)
    avisos = []
    r = pyxform.convert(xlsform=str(ruta), validate=False, warnings=avisos)
    assert "jr://file-csv/unidades.csv" in r.xform
    assert avisos == []


def test_nombres_unicos_y_listas_definidas():
    s = _survey()
    nombres = [f["name"] for f in s if f["type"] not in ("end_group", "end_repeat")]
    assert [n for n, k in Counter(nombres).items() if k > 1] == []
    listas = {c["list_name"] for c in _choices()}
    usadas = {f["type"].split()[1] for f in s if f["type"].startswith(("select_one ", "select_multiple "))
              and "from_file" not in f["type"]}
    assert usadas <= listas, usadas - listas
    assert Counter(f["type"] for f in s if f["type"] in ("begin_group", "end_group"))["begin_group"] == \
        Counter(f["type"] for f in s)["end_group"]


def test_campos_y_opciones_iguales_a_f_la():
    """Los datos que ya existen en F-LA-02/03 conservan nombre, lista y códigos de opción."""
    s = {f["name"]: f for f in _survey() if f["type"] != "end_repeat"}
    for fid in ("f_la_02_titular", "f_la_03_inspeccion"):
        for c in lc.FORMULARIOS[fid].campos:
            if c.nombre in s and c.lista:
                assert s[c.nombre]["type"].endswith(" " + c.lista), c.nombre
    por_lista = {}
    for c in _choices():
        por_lista.setdefault(c["list_name"], []).append((c["name"], c["label"]))
    for ln in ("tipo_titularidad", "uso_actual", "estado_civil", "aceptacion", "asistente", "docs_exhibidos"):
        assert por_lista[ln] == lc.OPCIONES[ln], ln
    assert ("otro", "Otro") in por_lista["uso_actual"]
    assert por_lista["personeria"] == [("natural", "Natural"), ("juridica", "Jurídica")]


def test_cascada_provincia_distrito_y_unidad():
    ch = _choices()
    provincias = {c["name"] for c in ch if c["list_name"] == "provincia"}
    distritos = [c for c in ch if c["list_name"] == "distrito"]
    assert provincias == set(fp.PROVINCIAS) and len(distritos) == 15
    assert {d["provincia"] for d in distritos} == provincias
    assert {d["name"] for d in distritos} == set(fp.DISTRITOS)          # mismos códigos que unidades.csv
    s = {f["name"]: f for f in _survey() if f["type"] != "end_repeat"}
    assert s["predio_distrito"]["choice_filter"] == "provincia=${predio_provincia}"
    assert s["unidad"]["type"] == "select_one_from_file unidades.csv"
    assert s["unidad"]["choice_filter"] == "asistente=${asistente} or ${asistente}='ESP'"


def test_sincronizacion_no_lo_toma_como_f_la():
    from liberacion_areas import la_kobo
    assert la_kobo.form_id_por_nombre(xf.TITULO) is None


def test_puntos_de_control_con_numero_fijo():
    """r_puntos lleva repeat_count: Kobo no pide «agregar otro» sin fin y la ficha continúa con las fotos."""
    s = {f["name"]: f for f in _survey() if f["type"] != "end_repeat"}
    n = s["n_puntos"]
    assert n["type"] == "integer" and n["required"] == "yes" and "<= 50" in n["constraint"]
    assert s["r_puntos"]["repeat_count"] == "${n_puntos}" and s["r_puntos"]["relevant"] == "${n_puntos} > 0"


def test_afectaciones_con_numero_fijo_y_ningun_repeat_abierto():
    s = {f["name"]: f for f in _survey() if f["type"] != "end_repeat"}
    n = s["n_afectaciones"]
    assert n["relevant"] == "${tiene_afectaciones}='si'" and n["constraint"] == ". >= 1 and . <= 20"
    assert s["r_afectaciones"]["repeat_count"] == "${n_afectaciones}"
    assert all(f.get("repeat_count") for f in _survey() if f["type"] == "begin_repeat")
