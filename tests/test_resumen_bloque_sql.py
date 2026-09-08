"""Consistencia del esquema y del CRUD de `resumen_bloque_excel`.

database.py no puede importarse sin `st.secrets["DATABASE_URL"]`, de modo que
estas pruebas leen el modulo como texto y verifican estaticamente lo que en
produccion solo fallaria contra PostgreSQL: que las columnas del INSERT, sus
marcadores y la tupla de valores concuerden, que el UPDATE toque las mismas
columnas y que todas existan en el CREATE TABLE.
"""

import os
import re
import unittest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FUENTE = open(os.path.join(RAIZ, "database.py"), encoding="utf-8").read()

TABLA = "resumen_bloque_excel"


def _bloque(patron):
    m = re.search(patron, FUENTE, re.DOTALL | re.IGNORECASE)
    assert m, f"No se encontro el fragmento SQL: {patron}"
    return m.group(1)


class TestEsquemaResumenBloque(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cuerpo = _bloque(
            rf"CREATE TABLE IF NOT EXISTS {TABLA} \((.*?)\n        \)\s*\"\"\"")
        cls.columnas = []
        for linea in cuerpo.splitlines():
            linea = linea.strip().rstrip(",")
            if not linea or linea.startswith("--"):
                continue
            nombre = linea.split()[0]
            if nombre.upper() in ("PRIMARY", "UNIQUE", "FOREIGN", "CHECK"):
                continue
            cls.columnas.append(nombre)

    def test_columnas_esperadas(self):
        for columna in ("id", "codigo_bloque", "bloque_id", "nombre_archivo",
                        "datos_json", "archivo_xlsx", "tamano_bytes",
                        "fecha_carga", "n_sustantivas", "validacion_utm"):
            self.assertIn(columna, self.columnas)

    def test_archivo_original_se_guarda_como_binario(self):
        self.assertRegex(FUENTE, r"archivo_xlsx\s+BYTEA")

    def test_indice_unico_por_bloque(self):
        """Recargar el mismo bloque debe actualizar, no duplicar."""
        self.assertRegex(
            FUENTE, r"CREATE UNIQUE INDEX uq_resumen_bloque_codigo\s+"
                    rf"ON {TABLA} \(codigo_bloque\)")

    def test_creacion_idempotente(self):
        self.assertIn(f"CREATE TABLE IF NOT EXISTS {TABLA}", FUENTE)

    def test_no_borra_datos_existentes(self):
        """La migracion no puede traer sentencias destructivas."""
        for peligrosa in ("DROP TABLE", "TRUNCATE", "DROP COLUMN",
                          "DELETE FROM bloques", "DROP DATABASE"):
            self.assertNotIn(peligrosa, FUENTE.upper())


class TestInsertUpdate(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        # Se acota al INSERT de esta tabla: el modulo tiene otros INSERT con
        # su propia lista de marcadores.
        sentencia = _bloque(rf"(INSERT INTO {TABLA} \(.*?VALUES \([?,]+\))")
        insert = re.search(r"\((.*?)\)\s*\n\s*VALUES", sentencia,
                           re.DOTALL).group(1)
        cls.cols_insert = [c.strip() for c in insert.replace("\n", " ").split(",")
                           if c.strip()]
        cls.marcadores = sentencia.rsplit("VALUES", 1)[1].count("?")
        update = _bloque(r"UPDATE resumen_bloque_excel SET\s*\n(.*?)\n\s*WHERE")
        cls.cols_update = re.findall(r"(\w+)=\?", update)
        valores = _bloque(r"valores = \((.*?)\n    \)")
        # Cada elemento de la tupla ocupa una posicion; se cuentan las comas
        # de primer nivel para no confundir las de las llamadas anidadas.
        cls.n_valores = _contar_elementos(valores)

    def test_insert_columnas_igual_a_marcadores(self):
        self.assertEqual(len(self.cols_insert), self.marcadores)

    def test_insert_columnas_igual_a_valores(self):
        self.assertEqual(len(self.cols_insert), self.n_valores)

    def test_update_cubre_las_mismas_columnas_salvo_la_clave(self):
        """El UPDATE recibe valores[1:] + (codigo,): debe fijar todo menos
        codigo_bloque, que es la clave de busqueda."""
        self.assertEqual(self.cols_update, self.cols_insert[1:])

    def test_columnas_existen_en_la_tabla(self):
        declaradas = TestEsquemaResumenBloque.__dict__.get("columnas")
        if declaradas is None:
            TestEsquemaResumenBloque.setUpClass()
            declaradas = TestEsquemaResumenBloque.columnas
        for columna in self.cols_insert:
            self.assertIn(columna, declaradas)

    def test_binario_envuelto_para_psycopg2(self):
        """El BYTEA debe viajar como psycopg2.Binary, no como bytes crudos."""
        self.assertRegex(FUENTE, r"_pg\.Binary\(contenido_xlsx\)")

    def test_lectura_convierte_memoryview_a_bytes(self):
        self.assertIn('bytes(row["archivo_xlsx"])', FUENTE)

    def test_listado_no_arrastra_el_binario(self):
        """El catalogo de 117 libros no debe traer los BYTEA en cada recarga."""
        cols = _bloque(r"_COLS_RESUMEN = \"\"\"(.*?)\"\"\"")
        self.assertNotIn("archivo_xlsx", cols)
        self.assertNotIn("datos_json", cols)
        self.assertIn("codigo_bloque", cols)


def _contar_elementos(texto):
    """Cuenta elementos de primer nivel en el cuerpo de una tupla."""
    profundidad = 0
    elementos = 1
    for caracter in texto:
        if caracter in "([{":
            profundidad += 1
        elif caracter in ")]}":
            profundidad -= 1
        elif caracter == "," and profundidad == 0:
            elementos += 1
    # Una coma final deja un elemento vacio: se descuenta.
    return elementos - 1 if texto.rstrip().endswith(",") else elementos


if __name__ == "__main__":
    unittest.main(verbosity=2)
