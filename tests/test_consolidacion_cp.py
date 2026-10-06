"""Pruebas de la consolidacion del Diagnostico Social por centro poblado.

Reproducen las inconsistencias que observaron las analistas sociales en los
graficos y libros Excel (octubre 2026) con los datos reales que las causaban:

  - La Alberca: 14 fichas F-DS-01 del mismo CP (el responsable escrito con
    variantes) sumaban 1 400 % de cobertura y 21 000 habitantes.
  - Coyona: dos CP homonimos (Canchaque y San Miguel de El Faique) se
    apilaban en una sola barra (243 + 200 + 1 100 hab., 200 % de cobertura).
  - Chacayo: dos fichas con 100 % y 60 % de tierras tituladas daban 160 %.
  - Buenos Aires: 11 caserios y los graficos contaban 17, 18 y 19 "CP".
  - Excel: totales que suman porcentajes y columna "Poblacion total" en 0.
"""

import io
import json
import os
import sys
import unittest

from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import analitica_social as an  # noqa: E402


_ID = [0]


def _reg(ficha, form, cp, bloque="M7B2", distrito="Salitral", **extra):
    _ID[0] += 1
    num = ficha.split("-")[-1]
    registro = {
        "id": extra.pop("id", _ID[0]), "ficha": ficha, "bloque_codigo": bloque,
        "centro_poblado": cp,
        "comunidad_campesina": extra.pop("comunidad_campesina", "Ninguna"),
        "distrito": distrito, "provincia": extra.pop("provincia", "Morropon"),
        "fecha_evaluacion": extra.pop("fecha_evaluacion", "2026-08-04"),
        "evaluador": extra.pop("evaluador", "Stefany Campos"),
        "nombre_entrevistado": extra.pop("nombre_entrevistado", ""),
        f"ds{num}_data_v3": json.dumps(form, ensure_ascii=False),
    }
    registro.update(extra)
    return registro


def _serie(informe, id_):
    for seccion in informe["secciones"]:
        for serie in seccion["series"]:
            if serie["id"] == id_:
                return serie
    return None


def _tabla(serie):
    """{categoria: {subclase: valor}} tal como lo dibujan grafico y Excel."""
    categorias, subclases, matriz = an._pivote(serie)
    return {c: dict(zip(subclases, fila)) for c, fila in zip(categorias, matriz)}


def _la_alberca():
    """14 fichas del mismo CP: el responsable cambia en cada una."""
    energias = ["90%", "90%", "90%", "80%", "90%", "90%", "90%", "70%",
                "90%", "90%", "90%", "80%", "80%", "80%"]
    variantes = ["Stefany Campos", "Stefany Campos A", "Stefany Campos Ab",
                 "Stefany Campos Aba", "Stefany Campos AbaD", "Stefany.Campos.Abad.",
                 "stefany.campos.abad", "Stefany CamposA", "Stefany Campos Abad.",
                 "Stefany Campos Abad..", "Stefany.Campos.abad", "Stefany Campos AbaDD",
                 "Stefany.Campos Abad.", "Stefany.Campos.Abad.."]
    return [_reg("F-DS-01", {"f1_nfam": "500", "f1_pob_t": "1500", "f1_pob_h": "220",
                             "f1_pob_m": "230", "f1_agua_cob": "100%",
                             "f1_energia_cob": energia,
                             "f1_agua": ["JASS / Sistema local"],
                             "f1_sanea": ["Letrina seca", "Pozo séptico"],
                             "f1_energia": ["Red pública"],
                             "f1_tenencia": "Estatal"},
                 "La Alberca", evaluador=ev, nombre_entrevistado=f"Informante {i}")
            for i, (energia, ev) in enumerate(zip(energias, variantes))]


def _coyona():
    """Dos CP homonimos de distritos distintos (catalogo INEI: bloques 24 y 43)."""
    return [
        _reg("F-DS-01", {"f1_nfam": "185", "f1_pob_t": "1100",
                         "f1_agua_cob": "100%", "f1_energia_cob": "100 %"},
             "Coyona", bloque="24", distrito="Canchaque",
             fecha_evaluacion="2026-07-22"),
        _reg("F-DS-01", {"f1_nfam": "185", "f1_pob_t": "443", "f1_pob_h": "243",
                         "f1_pob_m": "200", "f1_agua_cob": "100%",
                         "f1_energia_cob": "100%"},
             "Coyona", bloque="43", distrito="San Miguel de El Faique",
             fecha_evaluacion="2026-08-21"),
    ]


def _chacayo():
    return [
        _reg("F-DS-01", {"f1_pob_t": "200", "f1_pob_h": "120", "f1_pob_m": "80",
                         "f1_pct_tituladas": "100", "f1_tenencia": "Comunal (tierras comunales)"},
             "Chacayo", bloque="12", distrito="Santo Domingo",
             fecha_evaluacion="2026-09-10"),
        _reg("F-DS-01", {"f1_pob_t": "180", "f1_pob_h": "80", "f1_pob_m": "100",
                         "f1_pct_tituladas": "60", "f1_tenencia": "Comunal (tierras comunales)"},
             "Chacayo", bloque="12", distrito="Santo Domingo",
             fecha_evaluacion="2026-06-09"),
    ]


# ══════════════════════════════════════════════════════════════════════════
# Lectura de los campos de texto libre
# ══════════════════════════════════════════════════════════════════════════

class LecturaDeCampos(unittest.TestCase):

    def test_porcentaje_con_aclaraciones(self):
        """Manda la cifra con el signo %: antes 'cada 15 días 50%' daba 15."""
        self.assertEqual(an._pct("cada 15 días 50%"), 50)
        self.assertEqual(an._pct("50% 15 días al año"), 50)
        self.assertEqual(an._pct("80% (de 8am a 6pm)"), 80)
        self.assertEqual(an._pct("100 %"), 100)
        self.assertEqual(an._pct("90"), 90)

    def test_texto_que_no_es_porcentaje(self):
        self.assertIsNone(an._pct("80'/día"))
        self.assertIsNone(an._pct("30 min"))
        self.assertIsNone(an._pct(""))

    def test_porcentaje_fuera_de_rango_no_se_grafica(self):
        self.assertEqual(an._pct("160"), 160)
        self.assertIsNone(an._pct_valido("160"))
        self.assertEqual(an._pct_valido("100"), 100)

    def test_miles_con_punto_en_conteos(self):
        self.assertEqual(an._entero("1.500"), 1500)
        self.assertEqual(an._entero("1,500"), 1500)
        self.assertEqual(an._entero("233"), 233)

    def test_duracion_en_horas(self):
        """'1 hora' se leia como 1 minuto."""
        self.assertEqual(an._minutos("1 hora"), 60)
        self.assertEqual(an._minutos("1hora"), 60)
        self.assertEqual(an._minutos("1 hora 30m"), 90)
        self.assertEqual(an._minutos("una hora y media"), 90)
        self.assertEqual(an._minutos("45 min"), 45)
        self.assertEqual(an._minutos("60"), 60)

    def test_multiselect_guardado_como_texto(self):
        self.assertEqual(
            an._lista("Letrina seca / Pozo séptico", opciones=an.FL.L_SANEA),
            ["Letrina seca", "Pozo séptico"])
        self.assertEqual(
            an._lista("JASS / Sistema local", opciones=an.FL.L_AGUA),
            ["JASS / Sistema local"])

    def test_comunidad_ninguna_no_es_ambito(self):
        for texto in ("Ninguna", "Ninguno", "NO PERTENECEN A NINGUNA COMUNIDAD CAMPESINA"):
            self.assertEqual(an._comunidad({"comunidad_campesina": texto}), "")
        self.assertEqual(an._comunidad({"comunidad_campesina": "CC Andanjo"}), "CC Andanjo")


# ══════════════════════════════════════════════════════════════════════════
# Un centro poblado se cuenta una sola vez
# ══════════════════════════════════════════════════════════════════════════

class UnCentroPobladoUnaVez(unittest.TestCase):

    def test_la_alberca_es_un_solo_cp(self):
        unidades = an.unidades_por_cp(_la_alberca())
        self.assertEqual(len(unidades), 1)
        self.assertEqual(len(unidades[0]["fichas"]), 14)

    def test_la_alberca_cobertura_no_supera_100(self):
        """Antes: agua 1 400 % y energía 1 200 %."""
        informe = an.indicadores_bloque({"codigo": "M7B2"}, _la_alberca())
        cobertura = _tabla(_serie(informe, "f1_cobertura"))
        self.assertEqual(cobertura["La Alberca"]["Agua para consumo"], 100)
        # Moda de las 14 fichas: 90 % (8 fichas), no la suma.
        self.assertEqual(cobertura["La Alberca"]["Energía eléctrica"], 90)

    def test_la_alberca_poblacion_no_se_multiplica(self):
        """Antes: 14 × 1 500 = 21 000 habitantes."""
        informe = an.indicadores_bloque({"codigo": "M7B2"}, _la_alberca())
        metricas = {m["etiqueta"]: m["valor"] for m in informe["metricas"]}
        self.assertEqual(metricas["Población del ámbito"], "1 500")
        poblacion = _tabla(_serie(informe, "f1_poblacion"))["La Alberca"]
        self.assertEqual(poblacion["Hombres"] + poblacion["Mujeres"]
                         + poblacion["Sin desagregar por sexo"], 1500)

    def test_marcado_multiple_cuenta_centros_poblados(self):
        """Antes: 14 'CP' con JASS para un solo caserío."""
        informe = an.indicadores_bloque({"codigo": "M7B2"}, _la_alberca())
        agua = _serie(informe, "f1_fuentes_agua")
        self.assertEqual(agua["filas"], [{"clase": "JASS / Sistema local", "valor": 1}])
        self.assertEqual(agua["pct_base"], 1)
        tenencia = _serie(informe, "f1_tenencia")
        self.assertEqual(tenencia["filas"], [{"clase": "Estatal", "valor": 1}])

    def test_variantes_de_escritura_del_cp(self):
        regs = [_reg("F-DS-01", {"f1_pob_t": "250"}, "Dótor", bloque="85",
                     distrito="San Juan de Bigote"),
                _reg("F-DS-01", {"f1_pob_t": "250"}, "Dotor", bloque="85",
                     distrito="San Juan de Bigote")]
        self.assertEqual(len(an.unidades_por_cp(regs)), 1)

    def test_cp_compartido_por_dos_bloques(self):
        """Miguel Pampa figura en los bloques 77, 83 y 84: es un solo CP."""
        regs = [_reg("F-DS-01", {"f1_pob_t": "480"}, "Miguel Pampa", bloque=b,
                     distrito="San Juan de Bigote")
                for b in ("77", "83", "84")]
        informe = an.indicadores_consolidado(regs)
        metricas = {m["etiqueta"]: m["valor"] for m in informe["metricas"]}
        self.assertEqual(metricas["Centros poblados cubiertos"], "1")
        self.assertEqual(metricas["Población del ámbito"], "480")


class Homonimos(unittest.TestCase):

    def setUp(self):
        self.informe = an.indicadores_consolidado(_coyona())

    def test_coyona_son_dos_centros_poblados(self):
        poblacion = _tabla(_serie(self.informe, "f1_poblacion"))
        self.assertEqual(set(poblacion),
                         {"Coyona (Canchaque)", "Coyona (San Miguel de El Faique)"})

    def test_no_se_apilan_los_datos_de_otro_cp(self):
        """Antes: 243 + 200 + 1 100 = 1 543 en una sola barra."""
        poblacion = _tabla(_serie(self.informe, "f1_poblacion"))
        canchaque = poblacion["Coyona (Canchaque)"]
        faique = poblacion["Coyona (San Miguel de El Faique)"]
        self.assertEqual(canchaque["Sin desagregar por sexo"], 1100)
        self.assertIsNone(canchaque["Hombres"])          # sin dato, no 0
        self.assertEqual((faique["Hombres"], faique["Mujeres"]), (243, 200))

    def test_cobertura_de_cada_coyona(self):
        cobertura = _tabla(_serie(self.informe, "f1_cobertura"))
        for fila in cobertura.values():
            self.assertTrue(all(v <= 100 for v in fila.values() if v is not None))

    def test_misma_etiqueta_en_todas_las_secciones(self):
        """Una entrevista solo en Coyona de Canchaque se rotula igual que en
        la F-DS-01, aunque en la F-DS-03 no haya homonimo."""
        regs = _coyona() + [_reg("F-DS-03", {"f3_nombre": "Walter Facundo",
                                              "f3_cargo": "Teniente Gobernador"},
                                 "Coyona", bloque="24", distrito="Canchaque")]
        informe = an.indicadores_consolidado(regs)
        entrevistas = next(s for s in informe["secciones"] if s["id"] == "F-DS-03")
        detalle = dict(entrevistas["tablas"])["Entrevistas"]
        self.assertEqual(detalle[0]["Centro poblado / ámbito"], "Coyona (Canchaque)")

    def test_se_contrasta_con_el_inei(self):
        """La ficha del bloque 43 declara 443 hab.; el INEI le da 74."""
        detalles = [o["Detalle"] for o in self.informe["control"]["tablas"][0][1]
                    if o["Centro poblado / ámbito"] == "Coyona (San Miguel de El Faique)"]
        self.assertTrue(any("INEI" in d for d in detalles), detalles)


class FichasQueNoCoinciden(unittest.TestCase):

    def test_chacayo_tituladas_no_supera_100(self):
        """Antes: 100 + 60 = 160 %."""
        informe = an.indicadores_bloque({"codigo": "12"}, _chacayo())
        tituladas = _serie(informe, "f1_tituladas")
        self.assertEqual(len(tituladas["filas"]), 1)
        self.assertEqual(tituladas["filas"][0]["valor"], 100)   # la más reciente

    def test_la_discrepancia_queda_en_el_control_de_calidad(self):
        informe = an.indicadores_bloque({"codigo": "12"}, _chacayo())
        control = dict(informe["control"]["tablas"])["Control de calidad"]
        tituladas = [o for o in control if "Tierras tituladas" in o["Detalle"]]
        self.assertEqual(len(tituladas), 1)
        self.assertIn("60", tituladas[0]["Detalle"])

    def test_la_poblacion_sale_de_una_misma_ficha(self):
        """Hombres, mujeres y total no se mezclan entre informantes."""
        informe = an.indicadores_bloque({"codigo": "12"}, _chacayo())
        chacayo = _tabla(_serie(informe, "f1_poblacion"))["Chacayo"]
        self.assertEqual((chacayo["Hombres"], chacayo["Mujeres"]), (120, 80))

    def test_moda_entre_informantes(self):
        regs = [_reg("F-DS-01", {"f1_pob_t": t}, "Quemazon", bloque="M11B3",
                     distrito="San Juan de Bigote", nombre_entrevistado=f"I{i}")
                for i, t in enumerate(("1000", "700", "350", "1000"))]
        unidad = an.unidades_por_cp(regs)[0]
        self.assertEqual(an._entero(unidad["form"]["f1_pob_t"]), 1000)

    def test_hombres_mas_mujeres_mayor_que_el_total(self):
        regs = [_reg("F-DS-01", {"f1_pob_t": "180", "f1_pob_h": "180", "f1_pob_m": "120"},
                     "Nueva Esperanza", bloque="M28B4", distrito="Yamango")]
        informe = an.indicadores_bloque({"codigo": "M28B4"}, regs)
        fila = _tabla(_serie(informe, "f1_poblacion"))["Nueva Esperanza"]
        self.assertEqual(fila["Sin desagregar por sexo"], 180)
        control = dict(informe["control"]["tablas"])["Control de calidad"]
        self.assertTrue(any("no coincide" in o["Detalle"] for o in control))


class TenenciaPorBloque(unittest.TestCase):
    """La seccion 4 de la F-DS-01 es "Tenencia de la tierra relacionada al
    bloque": un valor por bloque, no por centro poblado ni por ficha."""

    def setUp(self):
        regs = [
            _reg("F-DS-01", {"f1_tenencia": "Comunal (tierras comunales)",
                             "f1_pct_tituladas": "40", "f1_superpone": "Sí"},
                 cp, bloque="M6B10", distrito="Buenos Aires",
                 fecha_evaluacion=fecha, nombre_entrevistado=cp)
            for cp, fecha in (("La Pilca", "2026-07-14"), ("La Maravilla", "2026-06-17"))]
        regs.append(_reg("F-DS-01", {"f1_tenencia": "Estatal", "f1_pct_tituladas": "0"},
                         "Ingenio de Buenos Aires", bloque="M6B10", distrito="Buenos Aires",
                         fecha_evaluacion="2026-06-17"))
        regs.append(_reg("F-DS-01", {"f1_tenencia": "Estatal"}, "Rio Seco",
                         bloque="M1B1", distrito="Buenos Aires"))
        self.informe = an.indicadores_consolidado(regs)

    def test_un_regimen_por_bloque(self):
        tenencia = {f["clase"]: f["valor"]
                    for f in _serie(self.informe, "f1_tenencia")["filas"]}
        # M6B10: Comunal (2 de 3 fichas); M1B1: Estatal. Antes: 3 + 1 "CP".
        self.assertEqual(tenencia, {"Comunal (tierras comunales)": 1, "Estatal": 1})
        self.assertEqual(_serie(self.informe, "f1_tenencia")["unidad"], "bloques")

    def test_tituladas_una_barra_por_bloque(self):
        tituladas = _serie(self.informe, "f1_tituladas")["filas"]
        self.assertEqual(tituladas, [{"clase": "Bloque M6B10", "valor": 40}])

    def test_fichas_del_bloque_que_no_coinciden(self):
        control = dict(self.informe["control"]["tablas"])["Control de calidad"]
        temas = [o for o in control if o["Tema"].startswith("Tenencia del bloque")
                 and "M6B10" in o["Bloque(s)"]]
        self.assertTrue(any("Estatal (1)" in o["Detalle"] for o in temas), temas)


class OtrasFichas(unittest.TestCase):

    def test_dos_entrevistas_distintas_se_conservan(self):
        """La deduplicacion anterior descartaba la segunda entrevista del CP."""
        regs = [_reg("F-DS-03", {"f3_nombre": "Abimel Córdova", "f3_dur": "1 hora"},
                     "San Lorenzo / Taspa", bloque="57", evaluador="Leslie Carrión"),
                _reg("F-DS-03", {"f3_nombre": "Sabino Guerrero", "f3_dur": "45 min"},
                     "San Lorenzo / Taspa", bloque="57", evaluador="Leslie Carrión",
                     fecha_evaluacion="2026-07-07")]
        self.assertEqual(len(an.deduplicar(regs)), 2)

    def test_la_misma_entrevista_registrada_dos_veces(self):
        regs = [_reg("F-DS-03", {"f3_nombre": "Ricardo Campos Carrasco", "f3_dur": "1 hora"},
                     "Chacchacal", bloque="40", evaluador=ev)
                for ev in ("Stefany Campos Abad", "Stefany Campos Abad..")]
        self.assertEqual(len(an.deduplicar(regs)), 1)

    def test_duracion_media_en_minutos(self):
        regs = [_reg("F-DS-03", {"f3_nombre": "A", "f3_dur": "1 hora"}, "X", bloque="40"),
                _reg("F-DS-03", {"f3_nombre": "B", "f3_dur": "30 min"}, "X", bloque="40")]
        seccion = an._seccion_entrevistas(an.deduplicar(regs))
        self.assertIn("duración media 45 min", seccion["descripcion"])

    def test_actor_en_dos_fichas_cuenta_una_vez(self):
        actor = {"Nombre del actor": "Gabriel Padilla Julca", "Cargo": "Agente Municipal",
                 "Tipo": "Gobierno Local (Municipalidad)", "Influencia": "Alto",
                 "Interes": "Alto", "Posicion": "A favor del proyecto"}
        regs = [_reg("F-DS-02", {"f2_actores": [actor]}, "Hualtacal", bloque="81",
                     distrito="Canchaque", evaluador=ev, fecha_evaluacion=f)
                for ev, f in (("Stefany Campos Abad", "2026-07-20"),
                              ("Leslie Carrion Lastra.", "2026-09-26"))]
        informe = an.indicadores_bloque({"codigo": "81"}, regs)
        metricas = {m["etiqueta"]: m["valor"] for m in informe["metricas"]}
        self.assertEqual(metricas["Actores mapeados"], "1")

    def test_peligro_reportado_por_dos_fichas_del_mismo_cp(self):
        peligro = {"Peligro observado": "Huaycos", "¿Ocurre?": "Sí",
                   "Frecuencia": "F (Frecuente)", "Magnitud": "A (Alta)"}
        regs = [_reg("F-DS-06", {"f6_peligros": [peligro]}, "Bigote", bloque="M3B3",
                     distrito="San Juan de Bigote", nombre_entrevistado=n)
                for n in ("Uno", "Dos")]
        informe = an.indicadores_bloque({"codigo": "M3B3"}, regs)
        frecuencia = _serie(informe, "f6_frecuencia")
        self.assertEqual(sum(f["valor"] for f in frecuencia["filas"]), 1)


class HallazgosDeLaRevision(unittest.TestCase):

    def test_ficha_sin_cp_que_repite_a_un_cp_no_se_suma(self):
        """M10B4: Rio Seco Alto y la ficha de la comunidad Carlos Augusto
        Rivera declaran los mismos 233 hab.; antes la poblacion daba 466."""
        regs = [_reg("F-DS-01", {"f1_pob_t": "233", "f1_nfam": "113"}, "Rio Seco Alto",
                     bloque="M10B4", distrito="Chulucanas"),
                _reg("F-DS-01", {"f1_pob_t": "233", "f1_nfam": "113"}, "",
                     comunidad_campesina="Carlos Augusto Rivera", bloque="M10B4",
                     distrito="Chulucanas")]
        informe = an.indicadores_bloque({"codigo": "M10B4"}, regs)
        metricas = {m["etiqueta"]: m["valor"] for m in informe["metricas"]}
        self.assertEqual(metricas["Población del ámbito"], "233")
        control = dict(informe["control"]["tablas"])["Control de calidad"]
        self.assertTrue(any(o["Tema"] == "Ámbito repetido (no se suma)" for o in control))

    def test_ambito_compuesto_cuyos_cp_tienen_ficha_propia(self):
        regs = [_reg("F-DS-01", {"f1_pob_t": "100"}, "A", bloque="X1", distrito="D"),
                _reg("F-DS-01", {"f1_pob_t": "50"}, "B", bloque="X1", distrito="D"),
                _reg("F-DS-01", {"f1_pob_t": "150"}, "A / B", bloque="X1", distrito="D")]
        informe = an.indicadores_bloque({"codigo": "X1"}, regs)
        metricas = {m["etiqueta"]: m["valor"] for m in informe["metricas"]}
        self.assertEqual(metricas["Población del ámbito"], "150")
        self.assertNotIn("A / B", _tabla(_serie(informe, "f1_poblacion")))

    def test_homonimos_del_mismo_distrito_no_se_funden(self):
        catalogo = {"A1": {"demografia": [{"centro_poblado": "Santa Rosa", "poblacion_total": 300,
                                           "utm_este": 600000, "utm_norte": 9400000}]},
                    "B1": {"demografia": [{"centro_poblado": "Santa Rosa", "poblacion_total": 80,
                                           "utm_este": 610000, "utm_norte": 9410000}]}}
        original = an._catalogo_bloque
        an._catalogo_bloque = lambda codigo: catalogo.get(codigo, {})
        try:
            regs = [_reg("F-DS-01", {"f1_pob_t": "310"}, "Santa Rosa", bloque="A1", distrito="D1"),
                    _reg("F-DS-01", {"f1_pob_t": "75"}, "Santa Rosa", bloque="B1", distrito="D1")]
            informe = an.indicadores_consolidado(regs)
        finally:
            an._catalogo_bloque = original
        metricas = {m["etiqueta"]: m["valor"] for m in informe["metricas"]}
        self.assertEqual(metricas["Población del ámbito"], "385")
        self.assertEqual(len(_tabla(_serie(informe, "f1_poblacion"))), 2)

    def test_fila_precargada_sin_datos_no_tapa_la_ficha_anterior(self):
        llena = {"Peligro observado": "Huaycos", "¿Ocurre?": "Sí",
                 "Frecuencia": "F (Frecuente)", "Magnitud": "A (Alta)"}
        vacia = {"Peligro observado": "Huaycos", "¿Ocurre?": "", "Frecuencia": ""}
        regs = [_reg("F-DS-06", {"f6_peligros": [llena]}, "A", bloque="X1",
                     fecha_evaluacion="2026-07-01"),
                _reg("F-DS-06", {"f6_peligros": [vacia]}, "A", bloque="X1",
                     fecha_evaluacion="2026-07-05")]
        informe = an.indicadores_bloque({"codigo": "X1"}, regs)
        self.assertIsNotNone(_serie(informe, "f6_frecuencia"))

    def test_union_del_marcado_multiple_con_texto_heredado(self):
        regs = [_reg("F-DS-01", {"f1_sanea": ["Letrina seca"]}, "Chacayo", bloque="12",
                     fecha_evaluacion="2026-09-10"),
                _reg("F-DS-01", {"f1_sanea": "Letrina seca / Pozo séptico"}, "Chacayo",
                     bloque="12", fecha_evaluacion="2026-06-09")]
        informe = an.indicadores_bloque({"codigo": "12"}, regs)
        sanea = {f["clase"]: f["valor"] for f in _serie(informe, "f1_saneamiento")["filas"]}
        self.assertEqual(sanea, {"Letrina seca": 1, "Pozo séptico": 1})

    def test_prioridad_de_peligros_de_una_misma_ficha(self):
        pares = (("Sequía", "Helada"), ("Helada", "Sequía"), ("Sequía", "Lluvias"),
                 ("Lluvias", "Sequía"))
        regs = [_reg("F-DS-06", {"f6_p1": p1, "f6_p2": p2}, "A", bloque="X1",
                     nombre_entrevistado=f"I{i}")
                for i, (p1, p2) in enumerate(pares)]
        informe = an.indicadores_bloque({"codigo": "X1"}, regs)
        puntos = {f["clase"]: f["valor"] for f in _serie(informe, "f6_prioridad")["filas"]}
        self.assertEqual(sum(puntos.values()), 5)          # 3 + 2, una vez por CP
        self.assertTrue(all(v in (2, 3) for v in puntos.values()), puntos)

    def test_asistencia_sin_convocados_no_supera_100(self):
        regs = [_reg("F-DS-04", {"f4_fecha": "2026-07-01", "f4_lugar": "Local",
                                 "f4_conv_n": "80", "f4_tot": "62"}, "Chungayo", bloque="X1"),
                _reg("F-DS-04", {"f4_fecha": "2026-07-02", "f4_lugar": "Local",
                                 "f4_tot": "62"}, "Sapalache", bloque="X1")]
        informe = an.indicadores_bloque({"codigo": "X1"}, regs)
        serie = _serie(informe, "f4_asistencia")
        categorias, subclases, matriz = an._pivote(serie)
        pct, total = an._matriz_porcentaje(("ref", "Convocados"), subclases, matriz)
        self.assertLessEqual(max(t for t in total if t is not None), 1.0)
        fila_sin = pct[categorias.index(next(c for c in categorias if "Sapalache" in c))]
        self.assertTrue(all(v is None for v in fila_sin))

    def test_control_de_una_seccion_sin_graficos(self):
        regs = [_reg("F-DS-01", {"f1_agua_cob": "150%"}, "A", bloque="X1",
                     fecha_evaluacion="2026-07-02"),
                _reg("F-DS-01", {"f1_agua_cob": "sin servicio"}, "A", bloque="X1",
                     fecha_evaluacion="2026-07-01")]
        informe = an.indicadores_bloque({"codigo": "X1"}, regs)
        control = dict(informe["control"]["tablas"])["Control de calidad"]
        self.assertTrue(any("150%" in o["Detalle"] for o in control), control)


class CatalogoDelAmbito(unittest.TestCase):

    def test_buenos_aires_tiene_11_caserios(self):
        """Los bloques del distrito Buenos Aires suman 11 CP del catálogo
        INEI; una ficha por CP y no 17 o 19."""
        bloques = ["5", "M18B1", "M1B1", "M6B10", "M6B2-1", "M6B2-2", "M6B2-3"]
        regs = []
        for cp, bloque in (("Rio Seco", "M1B1"), ("La Pilca", "M6B10"),
                           ("Hualas", "M18B1")):
            for i in range(4):
                regs.append(_reg("F-DS-01", {"f1_pob_t": "100", "f1_agua": ["JASS / Sistema local"],
                                             "f1_sanea": ["Letrina seca"]},
                                 cp, bloque=bloque, distrito="Buenos Aires",
                                 nombre_entrevistado=f"I{i}"))
        informe = an.indicadores_consolidado(regs, etiqueta="Buenos Aires",
                                             bloques_ambito=bloques)
        catalogo = {f["clase"]: f["valor"]
                    for f in _serie(informe, "cob_catalogo")["filas"]}
        self.assertEqual(sum(catalogo.values()), 11)
        self.assertEqual(catalogo["Con al menos una ficha social"], 3)
        sanea = _serie(informe, "f1_saneamiento")
        self.assertEqual(sanea["filas"][0]["valor"], 3)     # antes: 12
        metricas = {m["etiqueta"]: m for m in informe["metricas"]}
        self.assertIn("de 11 del catálogo", metricas["Centros poblados cubiertos"]["detalle"])

    def test_cp_de_un_ambito_compuesto_cuenta_como_cubierto(self):
        regs = [_reg("F-DS-02", {}, "La Laja / Overazal", bloque="14",
                     distrito="Santo Domingo")]
        catalogo = [{"nombre": "La Laja", "bloques": ["14"]},
                    {"nombre": "Overazal", "bloques": ["14"]},
                    {"nombre": "Otro", "bloques": ["14"]}]
        cubiertos, pendientes = an._cobertura_catalogo(regs, catalogo)
        self.assertEqual([c["nombre"] for c in pendientes], ["Otro"])


# ══════════════════════════════════════════════════════════════════════════
# Salidas: tabla de la app y libro Excel
# ══════════════════════════════════════════════════════════════════════════

class TablaDeLaApp(unittest.TestCase):

    def test_cobertura_sin_columna_total(self):
        """El CSV exportado traía 'Total' = agua % + energía %."""
        informe = an.indicadores_consolidado(_coyona())
        df = an.tabla_serie(_serie(informe, "f1_cobertura"))
        self.assertNotIn("Total", df.columns)
        self.assertEqual(df.columns[0], "Centro poblado")

    def test_tituladas_sin_porcentaje_del_total(self):
        informe = an.indicadores_bloque({"codigo": "12"}, _chacayo())
        df = an.tabla_serie(_serie(informe, "f1_tituladas"))
        self.assertNotIn("% del total", df.columns)
        self.assertEqual(len(df), 1)

    def test_poblacion_con_total_por_centro_poblado(self):
        informe = an.indicadores_consolidado(_coyona())
        df = an.tabla_serie(_serie(informe, "f1_poblacion"))
        totales = dict(zip(df["Centro poblado"], df["Población total"]))
        self.assertEqual(totales["Coyona (Canchaque)"], 1100)
        self.assertEqual(totales["Coyona (San Miguel de El Faique)"], 443)


def _bloque_tabla(ws, titulo):
    """Filas desde la cabecera de la tabla `titulo` hasta su fila de cierre."""
    cabeceras = ("Clase", "Centro poblado", "Programa social",
                 "Capacidad evaluada", "Actividad / rubro")
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == titulo:
            cab = r + 1
            while ws.cell(cab, 1).value not in cabeceras:
                cab += 1
            filas, fila = [], cab
            while ws.cell(fila, 1).value is not None:
                filas.append([ws.cell(fila, c).value for c in range(1, 8)])
                if str(filas[-1][0]).startswith(("TOTAL", "PROMEDIO", "Base:")):
                    break
                fila += 1
            return filas
    raise AssertionError(f"No se encontró la tabla {titulo!r}")


class LibroExcel(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        regs = _la_alberca() + _coyona() + _chacayo()
        for r in regs:
            form = json.loads(r["ds01_data_v3"])
            form.setdefault("f1_juntos", "20")
            form.setdefault("f1_pension65", "5")
            r["ds01_data_v3"] = json.dumps(form)
        cls.informe = an.indicadores_consolidado(regs)
        cls.wb = load_workbook(io.BytesIO(an.generar_excel_social(cls.informe)))
        cls.ws = next(ws for ws in cls.wb.worksheets if ws.title.startswith("F-DS-01"))

    def _tabla(self, prefijo):
        for r in range(1, self.ws.max_row + 1):
            valor = self.ws.cell(r, 1).value
            if isinstance(valor, str) and valor[3:].startswith(prefijo):
                return _bloque_tabla(self.ws, valor)
        raise AssertionError(prefijo)

    def test_poblacion_total_por_fila_y_columna_con_formula(self):
        tabla = self._tabla("Población por centro poblado")
        cab = tabla[0]
        self.assertEqual(cab[:5], ["Centro poblado", "Hombres", "Mujeres",
                                   "Sin desagregar por sexo", "Población total"])
        self.assertTrue(str(tabla[1][4]).startswith("=SUM(B"))
        self.assertEqual(tabla[-1][0], "TOTAL")
        self.assertTrue(str(tabla[-1][4]).startswith("=SUM(E"))

    def test_porcentajes_con_promedio_y_no_con_suma(self):
        tabla = self._tabla("Cobertura de agua y energía")
        self.assertEqual(tabla[-1][0], "PROMEDIO (CP con dato)")
        self.assertTrue(str(tabla[-1][1]).startswith("=IFERROR(AVERAGE("))
        for fila in tabla[1:-1]:
            for valor in fila[1:3]:
                self.assertTrue(valor is None or 0 <= valor <= 100, fila)

    def test_marcado_multiple_con_base_en_cp(self):
        tabla = self._tabla("Fuentes de agua para consumo")
        self.assertTrue(tabla[-1][0].startswith("Base: "))
        self.assertIn("centro(s) poblado(s)", tabla[-1][0])
        self.assertTrue(str(tabla[1][2]).startswith("=IFERROR(B"))
        self.assertIn("/$B$", tabla[1][2])         # referencia a la celda base

    def test_unidades_mixtas_sin_total(self):
        tabla = self._tabla("Cobertura de programas sociales")
        self.assertFalse(any(str(f[0]).startswith("TOTAL") for f in tabla))

    def test_hoja_de_control_de_calidad(self):
        nombres = " | ".join(self.wb.sheetnames)
        self.assertIn("Control de calidad", nombres)
        self.assertIn("Fichas F-DS-01 por CP", nombres)

    def test_demografia_con_formulas_de_control(self):
        ws = next(w for w in self.wb.worksheets if w.title.startswith("T Demograf"))
        cabeceras = [c.value for c in ws[7]]
        self.assertIn("Hombres + mujeres", cabeceras)
        self.assertIn("Diferencia: total − (H + M)", cabeceras)
        col = cabeceras.index("Hombres + mujeres") + 1
        self.assertTrue(str(ws.cell(8, col).value).startswith("=IF(AND(ISNUMBER("))
        ultima = max(r for r in range(8, ws.max_row + 1) if ws.cell(r, 1).value)
        self.assertEqual(ws.cell(ultima, 1).value, "TOTAL")


if __name__ == "__main__":
    unittest.main()
