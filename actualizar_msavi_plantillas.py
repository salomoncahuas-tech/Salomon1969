# -*- coding: utf-8 -*-
"""Traspaso automatizado de las áreas MSAVI por clase DN a las 117 plantillas Excel de bloque.

Proyecto IN Piura — CUI 2669244 — ANIN / DIME / SESDI.

Fuente de datos
---------------
``AREAS_MSAVI_BLOQUES_V6_REV_HSCM.xlsx``
    hoja ``Resumen_por_Bloque``: área en hectáreas por bloque y clase DN (1 a 5)
    del ráster MSAVI 2024 clasificado, y número de polígonos por bloque.

Destino
-------
Las 117 plantillas ``Plantilla_Excel_Bloque_<código>_IN_Piura.xlsx``. En cada una
se actualizan las hojas:

* ``Cobertura MSAVI-NDVI`` — sección A: superficie y porcentaje por clase MSAVI,
  totales, superficie sobre/bajo el umbral 0.4976 y contraste con el catálogo.
* ``Resumen`` — sección 3: superficie clasificada MSAVI, desviación frente al
  catálogo, superficie bajo umbral (brecha), clase DN dominante y n.° de polígonos.
* ``Control de consistencia`` — verificaciones D-xx del traspaso y recuento.
* ``Microcuenca`` — columna comparativa «% bajo umbral MSAVI» de los bloques de
  la misma microcuenca.

Equivalencia DN → clase MSAVI
-----------------------------
El archivo fuente no documenta la semántica del DN. La correspondencia se
estableció por contraste con la media MSAVI 2024 del catálogo maestro de los 117
bloques: reconstruyendo la media de cada bloque con los puntos medios de clase
bajo el orden DN 1 = clase inferior … DN 5 = clase superior se obtiene r = 0.9965
y un error absoluto medio de 0.0091 frente a la media de catálogo; el orden
inverso da r = -0.9975. La verificación se rehace en cada ejecución
(``--verificar-dn``) y se reporta en la hoja de control de consistencia.

Uso
---
    python3 actualizar_msavi_plantillas.py \
        --fuente AREAS_MSAVI_BLOQUES_V6_REV_HSCM.xlsx \
        --plantillas "Plantillas Excel 117 bloques" \
        --salida plantillas_117_msavi_v6

El script es idempotente: una segunda ejecución sobre archivos ya actualizados
reescribe los valores sin duplicar filas ni columnas.
"""

from __future__ import annotations

import argparse
import copy
import json
import math
import os
import re
import shutil
import statistics
import sys
import tempfile
import zipfile
from xml.etree import ElementTree
from dataclasses import dataclass, field

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Constantes institucionales y de clasificación
# --------------------------------------------------------------------------

VERDE_ANIN = "1B4D2E"
VERDE_CLARO = "E8F0EA"
AMBAR = "FFF2CC"
UMBRAL_BRECHA = 0.4976

# Clases del MSAVI 2024 clasificado, de la superior a la inferior, tal como
# figuran en la sección A de la hoja «Cobertura MSAVI-NDVI» de las plantillas.
CLASES_MSAVI = [
    # (etiqueta en plantilla, DN, límite inferior, límite superior, interpretación)
    ("> 0.6139", 5, 0.6139, None, "Vigor alto"),
    ("0.4976 - 0.6139", 4, 0.4976, 0.6139, "Vigor moderado"),
    ("0.3813 - 0.4976", 3, 0.3813, 0.4976, "Vigor bajo"),
    ("0.2650 - 0.3813", 2, 0.2650, 0.3813, "Vigor muy bajo"),
    ("<= 0.2650", 1, None, 0.2650, "Suelo desnudo / no vegetal"),
]

# Puntos medios por DN (1 a 5) para la verificación de la equivalencia DN → clase.
PUNTOS_MEDIOS = {1: 0.20670, 2: 0.32315, 3: 0.44945, 4: 0.55575, 5: 0.67210}

FILA_TOTAL_MSAVI = "TOTAL CLASIFICADO MSAVI 2024"
FILA_SOBRE_UMBRAL = "Superficie SOBRE umbral 0.4976 (ha)"
FILA_BAJO_UMBRAL = "Superficie BAJO umbral 0.4976 (ha) — brecha"
FILA_CATALOGO = "Superficie de catálogo (V5/V6) — contraste MSAVI"
ETQ_SUP_MSAVI = "Superficie clasificada MSAVI 2024 (ha)"
ETQ_BAJO_MSAVI = "Superficie bajo umbral MSAVI 0.4976 (ha)"
ETQ_CLASE_DOM = "Clase DN dominante (MSAVI 2024)"
ETQ_COL_MICRO = "MSAVI — % bajo umbral 0.4976"

VERIFICACIONES = [
    ("Distribución areal MSAVI 2024", "traspaso"),
    ("Superficie MSAVI vs. catálogo", "superficie"),
    ("Media MSAVI vs. clase DN dominante", "coherencia"),
]

RE_PLANTILLA = re.compile(r"^Plantilla_Excel_Bloque_(.+)_IN_Piura\.xlsx$")
RE_CODIGO_D = re.compile(r"^D-(\d+)$")
# Algunos libros incorporan una segunda serie de verificaciones (C-01, C-02:
# hallazgos sobre el propio archivo de la ficha DT). Se conservan con su
# código, pero cuentan en el resumen del control.
RE_CODIGO_VERIFICACION = re.compile(r"^[A-Z]{1,3}-\d+$")


# --------------------------------------------------------------------------
# Estructuras de datos
# --------------------------------------------------------------------------


@dataclass
class DatosMSAVI:
    """Áreas MSAVI de un bloque, en hectáreas, por clase DN."""

    bloque: str
    poligonos: int
    areas: dict  # {DN: ha}
    total: float

    @property
    def sobre_umbral(self) -> float:
        return self.areas[4] + self.areas[5]

    @property
    def bajo_umbral(self) -> float:
        return self.areas[1] + self.areas[2] + self.areas[3]

    @property
    def pct_bajo_umbral(self) -> float:
        return 100.0 * self.bajo_umbral / self.total if self.total else 0.0

    @property
    def dn_dominante(self) -> int:
        return max(self.areas, key=lambda dn: self.areas[dn])

    def pct(self, dn: int) -> float:
        return 100.0 * self.areas[dn] / self.total if self.total else 0.0


@dataclass
class Reporte:
    """Bitácora de la ejecución, para el cotejo posterior."""

    filas: list = field(default_factory=list)
    avisos: list = field(default_factory=list)


# --------------------------------------------------------------------------
# Lectura de la fuente
# --------------------------------------------------------------------------


def leer_fuente(ruta: str) -> dict:
    """Devuelve {código de bloque: DatosMSAVI} a partir de Resumen_por_Bloque."""
    wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
    ws = wb["Resumen_por_Bloque"]
    encabezado = None
    datos = {}
    for fila in ws.iter_rows(values_only=True):
        primera = fila[0]
        if encabezado is None:
            if isinstance(primera, str) and primera.strip().upper() == "BLOQUE":
                encabezado = True
            continue
        if primera is None:
            continue
        codigo = str(primera).strip()
        if codigo.upper() == "TOTAL":
            continue
        areas = {dn: float(fila[1 + dn] or 0.0) for dn in range(1, 6)}
        total = float(fila[7] or 0.0)
        if total <= 0:
            total = sum(areas.values())
        datos[codigo] = DatosMSAVI(
            bloque=codigo,
            poligonos=int(fila[1] or 0),
            areas=areas,
            total=total,
        )
    wb.close()
    return datos


def clase_de_media(media: float) -> int:
    """DN de la clase en la que cae una media MSAVI."""
    for _etq, dn, inferior, superior in ((c[0], c[1], c[2], c[3]) for c in CLASES_MSAVI):
        if superior is None:
            if media > inferior:
                return dn
        elif inferior is None:
            if media <= superior:
                return dn
        elif inferior < media <= superior:
            return dn
    return 1


def etiqueta_clase(dn: int) -> str:
    etq, _dn, _i, _s, interp = next(c for c in CLASES_MSAVI if c[1] == dn)
    return f"DN {dn} · {etq} ({interp})"


def verificar_equivalencia_dn(datos: dict, medias: dict) -> dict:
    """Contrasta el orden DN 1..5 contra las medias MSAVI del catálogo."""
    directo, inverso, observado = [], [], []
    for codigo, media in medias.items():
        d = datos.get(codigo)
        if not d or not d.total or media is None:
            continue
        est = sum(d.areas[dn] * PUNTOS_MEDIOS[dn] for dn in range(1, 6)) / d.total
        est_inv = sum(d.areas[dn] * PUNTOS_MEDIOS[6 - dn] for dn in range(1, 6)) / d.total
        directo.append(est)
        inverso.append(est_inv)
        observado.append(float(media))

    def r(xs, ys):
        mx, my = statistics.mean(xs), statistics.mean(ys)
        num = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
        den = math.sqrt(sum((x - mx) ** 2 for x in xs) * sum((y - my) ** 2 for y in ys))
        return num / den if den else float("nan")

    return {
        "n": len(observado),
        "r_directo": round(r(observado, directo), 4),
        "r_inverso": round(r(observado, inverso), 4),
        "eam_directo": round(statistics.mean(abs(a - b) for a, b in zip(observado, directo)), 4),
        "eam_inverso": round(statistics.mean(abs(a - b) for a, b in zip(observado, inverso)), 4),
    }


# --------------------------------------------------------------------------
# Utilidades de hoja de cálculo
# --------------------------------------------------------------------------


def insertar_filas(ws, idx: int, cantidad: int) -> None:
    """Inserta filas conservando combinaciones de celdas y alturas de fila.

    ``openpyxl`` desplaza las celdas con su estilo, pero no las celdas
    combinadas ni las alturas de fila; aquí se corrigen ambas.
    """
    combinadas = [(m.min_col, m.min_row, m.max_col, m.max_row) for m in ws.merged_cells.ranges]
    alturas = {r: d.height for r, d in ws.row_dimensions.items() if d.height is not None}
    formulas = {
        (celda.row, celda.column): celda.value
        for fila in ws.iter_rows(min_row=idx)
        for celda in fila
        if isinstance(celda.value, str) and celda.value.startswith("=")
    }

    ws.insert_rows(idx, cantidad)

    # ``insert_rows`` no reescribe las referencias: las fórmulas desplazadas se
    # retraducen a su nueva posición (las secciones se mueven completas, de modo
    # que sus referencias internas se desplazan con ellas).
    for (fila, columna), formula in formulas.items():
        origen = f"{get_column_letter(columna)}{fila}"
        destino = f"{get_column_letter(columna)}{fila + cantidad}"
        ws.cell(fila + cantidad, columna).value = Translator(formula, origin=origen).translate_formula(destino)

    # Las combinaciones se reasignan sobre la lista de rangos: las API
    # unmerge/merge eliminarían celdas ya desplazadas por la inserción.
    ws.merged_cells.ranges = [
        CellRange(
            min_col=min_col,
            min_row=min_row + cantidad if min_row >= idx else min_row,
            max_col=max_col,
            max_row=max_row + cantidad if max_row >= idx else max_row,
        )
        for min_col, min_row, max_col, max_row in combinadas
    ]

    for fila in list(alturas):
        ws.row_dimensions[fila].height = None
    for fila, altura in alturas.items():
        destino = fila + cantidad if fila >= idx else fila
        ws.row_dimensions[destino].height = altura


def num(valor, decimales: int = 0) -> str:
    """Número con separador de millares en espacio fino, uso peruano."""
    texto = f"{valor:,.{decimales}f}"
    return texto.replace(",", "\u202f")


# Formas de fórmula que emplean estas plantillas. Se evalúan en Python para
# poder guardar el valor en caché junto a la fórmula: openpyxl escribe la
# fórmula sin resultado, y cualquier lector de valores (el aplicativo IN Piura,
# pandas, una vista previa) leería la celda como vacía hasta abrirla en Excel.
_RE_SUMA = re.compile(r"^=SUM\((\$?[A-Z]+)\$?(\d+):(\$?[A-Z]+)\$?(\d+)\)$", re.I)
_RE_PROMEDIO = re.compile(
    r"^=ROUND\(AVERAGE\((\$?[A-Z]+)\$?(\d+):(\$?[A-Z]+)\$?(\d+)\),(\d+)\)$", re.I)
_RE_PORCENTAJE = re.compile(
    r"^=ROUND\(\$?([A-Z]+)\$?(\d+)/\$?([A-Z]+)\$?(\d+)\*100,(\d+)\)$", re.I)


def _valor_celda(ws, columna: str, fila: int, resueltos: dict):
    """Valor numérico de una celda: literal o fórmula ya evaluada."""
    coord = f"{columna.replace('$', '')}{fila}"
    if coord in resueltos:
        return resueltos[coord]
    valor = ws[coord].value
    if isinstance(valor, (int, float)):
        return float(valor)
    return None


def _evaluar_formula(ws, formula: str, resueltos: dict):
    """Evalúa las formas de fórmula presentes en las plantillas."""
    m = _RE_SUMA.match(formula)
    if m:
        col, desde, _col2, hasta = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
        valores = [_valor_celda(ws, col, f, resueltos) for f in range(desde, hasta + 1)]
        presentes = [v for v in valores if v is not None]
        return sum(presentes) if presentes else None

    m = _RE_PROMEDIO.match(formula)
    if m:
        col, desde, _col2, hasta = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
        decimales = int(m.group(5))
        valores = [_valor_celda(ws, col, f, resueltos) for f in range(desde, hasta + 1)]
        presentes = [v for v in valores if v is not None]
        return round(statistics.mean(presentes), decimales) if presentes else None

    m = _RE_PORCENTAJE.match(formula)
    if m:
        numerador = _valor_celda(ws, m.group(1), int(m.group(2)), resueltos)
        denominador = _valor_celda(ws, m.group(3), int(m.group(4)), resueltos)
        if numerador is None or not denominador:
            return None
        return round(numerador / denominador * 100, int(m.group(5)))

    return None


def calcular_valores_cacheados(wb) -> dict:
    """{hoja: {coordenada: valor}} para toda fórmula evaluable del libro.

    Se repite hasta estabilizar porque hay fórmulas encadenadas (el porcentaje
    de cada clase divide entre el total, que es a su vez una suma).
    """
    cacheados = {}
    for ws in wb.worksheets:
        resueltos = {}
        formulas = {
            celda.coordinate: celda.value
            for fila in ws.iter_rows() for celda in fila
            if isinstance(celda.value, str) and celda.value.startswith("=")
        }
        for _ in range(len(formulas) + 1):
            pendientes = [c for c in formulas if c not in resueltos]
            if not pendientes:
                break
            avance = False
            for coord in pendientes:
                valor = _evaluar_formula(ws, formulas[coord], resueltos)
                if valor is not None:
                    resueltos[coord] = valor
                    avance = True
            if not avance:
                break
        sin_resolver = sorted(set(formulas) - set(resueltos))
        if sin_resolver:
            raise ValueError(
                f"Fórmulas sin valor calculable en «{ws.title}»: "
                f"{', '.join(sin_resolver[:5])}"
            )
        if resueltos:
            cacheados[ws.title] = resueltos
    return cacheados


def _hojas_del_paquete(zf) -> dict:
    """{título de hoja: ruta del XML} leyendo workbook.xml y sus relaciones."""
    ns_rel = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    libro = ElementTree.fromstring(zf.read("xl/workbook.xml"))
    relaciones = ElementTree.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    destinos = {
        rel.get("Id"): rel.get("Target")
        for rel in relaciones
    }
    rutas = {}
    for hoja in libro.iter():
        if not hoja.tag.endswith("}sheet"):
            continue
        destino = destinos.get(hoja.get(f"{ns_rel}id"), "")
        if destino:
            rutas[hoja.get("name")] = "xl/" + destino.lstrip("/").replace("xl/", "", 1)
    return rutas


def inyectar_valores_cacheados(ruta: str, cacheados: dict) -> int:
    """Escribe el resultado de cada fórmula como valor en caché del .xlsx.

    La fórmula se conserva intacta: Excel la recalcula al abrir (el libro
    declara fullCalcOnLoad) y, entre tanto, quien lea valores obtiene el
    número correcto en vez de una celda vacía.
    """
    escritos = 0
    with zipfile.ZipFile(ruta) as zf:
        rutas = _hojas_del_paquete(zf)
        partes = {nombre: zf.read(nombre) for nombre in zf.namelist()}
        orden = zf.namelist()

    for hoja, valores in cacheados.items():
        destino = rutas.get(hoja)
        if not destino or destino not in partes:
            raise ValueError(f"No se ubicó el XML de la hoja «{hoja}»")
        xml = partes[destino].decode("utf-8")
        for coord, valor in valores.items():
            patron = re.compile(
                r'(<c r="%s"[^>]*>)(<f[^>]*>.*?</f>)(?:<v\s*/>|<v>.*?</v>)?(</c>)' % coord,
                re.S)
            texto = repr(round(float(valor), 10)) if isinstance(valor, float) else str(valor)
            xml, n = patron.subn(
                lambda m: f"{m.group(1)}{m.group(2)}<v>{texto}</v>{m.group(3)}", xml)
            escritos += n
        partes[destino] = xml.encode("utf-8")

    with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as zf:
        for nombre in orden:
            zf.writestr(nombre, partes[nombre])
    return escritos


def copiar_estilo(origen, destino) -> None:
    destino.font = copy.copy(origen.font)
    destino.fill = copy.copy(origen.fill)
    destino.border = copy.copy(origen.border)
    destino.alignment = copy.copy(origen.alignment)
    destino.number_format = origen.number_format


def clonar_fila_estilo(ws, fila_origen: int, fila_destino: int, columnas: int) -> None:
    for col in range(1, columnas + 1):
        copiar_estilo(ws.cell(fila_origen, col), ws.cell(fila_destino, col))


def buscar_fila(ws, texto: str, columna: int = 1, exacto: bool = True):
    for r in range(1, ws.max_row + 1):
        valor = ws.cell(r, columna).value
        if isinstance(valor, str):
            v = valor.strip()
            if (v == texto) if exacto else v.startswith(texto):
                return r
    return None


def bordes_finos() -> Border:
    lado = Side(style="thin", color="B7C4BA")
    return Border(left=lado, right=lado, top=lado, bottom=lado)


# --------------------------------------------------------------------------
# Hoja «Cobertura MSAVI-NDVI»
# --------------------------------------------------------------------------


def actualizar_cobertura(wb, d: DatosMSAVI, media: float, sup_catalogo: float, rep: Reporte) -> dict:
    ws = wb["Cobertura MSAVI-NDVI"]

    filas_clase = {}
    for etiqueta, dn, _inf, _sup, _interp in CLASES_MSAVI:
        fila = None
        for r in range(1, ws.max_row + 1):
            valor = ws.cell(r, 1).value
            if isinstance(valor, str) and valor.strip().split("·")[0].strip() == etiqueta:
                fila = r
                break
        if fila is None:
            raise ValueError(f"No se ubicó la clase «{etiqueta}» en la hoja de cobertura")
        filas_clase[dn] = fila

    primera = min(filas_clase.values())
    ultima = max(filas_clase.values())
    fila_media = buscar_fila(ws, "MSAVI 2024 — MEDIA", exacto=False)
    dn_media = clase_de_media(float(media))

    # -- filas de clase: etiqueta con el DN, superficie y porcentaje ---------
    for etiqueta, dn, _inf, _sup, _interp in CLASES_MSAVI:
        r = filas_clase[dn]
        ws.cell(r, 1).value = f"{etiqueta} · DN {dn}"
        celda_ha = ws.cell(r, 2)
        celda_ha.value = round(d.areas[dn], 4)
        celda_ha.number_format = "0.0000"
        celda_pct = ws.cell(r, 3)
        celda_pct.value = f"=ROUND(B{r}/$B${ultima + 1}*100,2)"
        celda_pct.number_format = "0.00"
        # Realce de la clase en la que cae la media del bloque; el resto
        # recupera el bandeado original de la tabla.
        resaltada = dn == dn_media
        for col in range(1, 6):
            celda = ws.cell(r, col)
            if resaltada:
                celda.fill = PatternFill("solid", fgColor=AMBAR)
                celda.font = Font(name="Arial", size=9, bold=True)
            else:
                sombreada = (r - primera) % 2 == 1
                celda.fill = (
                    PatternFill("solid", fgColor=VERDE_CLARO) if sombreada else PatternFill()
                )
                celda.font = Font(name="Arial", size=9, bold=False)

    # -- filas de síntesis, insertadas antes de la media del bloque ---------
    fila_total = buscar_fila(ws, FILA_TOTAL_MSAVI)
    if fila_total is None:
        insertar_filas(ws, ultima + 1, 4)
        fila_media += 4
        for i in range(4):
            clonar_fila_estilo(ws, fila_media, ultima + 1 + i, 5)
        fila_total = ultima + 1
    fila_sobre, fila_bajo, fila_cat = fila_total + 1, fila_total + 2, fila_total + 3

    desviacion = 100.0 * (d.total - sup_catalogo) / sup_catalogo if sup_catalogo else 0.0
    if abs(desviacion) < 0.005:
        desviacion = 0.0

    sintesis = [
        (
            fila_total,
            FILA_TOTAL_MSAVI,
            f"=SUM(B{primera}:B{ultima})",
            f"=SUM(C{primera}:C{ultima})",
            f"Suma de la distribución areal · {num(d.poligonos)} polígonos",
            "—",
            True,
        ),
        (
            fila_sobre,
            FILA_SOBRE_UMBRAL,
            f"=SUM(B{filas_clase[5]}:B{filas_clase[4]})",
            f"=ROUND(B{fila_sobre}/B{fila_total}*100,2)",
            "Vigor alto y moderado (DN 4 y DN 5)",
            "Sobre umbral 0.4976",
            False,
        ),
        (
            fila_bajo,
            FILA_BAJO_UMBRAL,
            f"=SUM(B{filas_clase[3]}:B{filas_clase[1]})",
            f"=ROUND(B{fila_bajo}/B{fila_total}*100,2)",
            "Vigor bajo, muy bajo y suelo desnudo (DN 1 a DN 3)",
            "Base para el indicador de brecha (R.M. N.° 00213-2024-MINAM)",
            False,
        ),
        (
            fila_cat,
            FILA_CATALOGO,
            round(float(sup_catalogo), 4),
            "—",
            "Desviación planimétrica frente al catálogo",
            f"{desviacion:+.2f} %",
            False,
        ),
    ]

    for fila, etiqueta, valor_b, valor_c, valor_d, valor_e, resaltar in sintesis:
        ws.cell(fila, 1).value = etiqueta
        ws.cell(fila, 2).value = valor_b
        ws.cell(fila, 2).number_format = "0.0000"
        ws.cell(fila, 3).value = valor_c
        ws.cell(fila, 3).number_format = "0.00" if isinstance(valor_c, str) and valor_c.startswith("=") else "General"
        ws.cell(fila, 4).value = valor_d
        ws.cell(fila, 5).value = valor_e
        for col in range(1, 6):
            celda = ws.cell(fila, col)
            celda.border = bordes_finos()
            celda.font = Font(name="Arial", size=9, bold=resaltar)
            celda.fill = PatternFill("solid", fgColor=VERDE_CLARO) if resaltar else PatternFill()
            celda.alignment = Alignment(
                horizontal="left" if col in (1, 4, 5) else "center",
                vertical="center",
                wrap_text=True,
            )
        ws.row_dimensions[fila].height = 30

    # -- nota metodológica de la sección A ----------------------------------
    fila_secB = buscar_fila(ws, "B. NDVI", exacto=False)
    fila_nota = None
    for r in range(fila_media, fila_secB or ws.max_row + 1):
        valor = ws.cell(r, 1).value
        if isinstance(valor, str) and valor.startswith("NOTA METODOL"):
            fila_nota = r
            break
    if fila_nota:
        ws.cell(fila_nota, 1).value = (
            "NOTA METODOLÓGICA. La distribución areal por clase de MSAVI 2024 procede de la estadística "
            "zonal por polígono del ráster MSAVI clasificado (archivo AREAS_MSAVI_BLOQUES_V6_REV_HSCM.xlsx, "
            f"hojas «Datos_Detalle» y «Resumen_por_Bloque»): {num(d.poligonos)} polígonos agregados para este "
            "bloque, con superficies obtenidas como Area_m2 / 10 000. El archivo fuente no documenta la "
            "semántica del valor de clase DN; la equivalencia DN 1 = clase inferior … DN 5 = clase superior "
            "se estableció por contraste con la media MSAVI 2024 del catálogo maestro de los 117 bloques "
            "(coeficiente de correlación 0.9965 y error absoluto medio 0.0091 al reconstruir la media con los "
            "puntos medios de clase; el orden inverso arroja correlación negativa), y debe ser confirmada por "
            f"el especialista SIG del proyecto. La media del bloque ({float(media):.4f}) se conserva del "
            "catálogo maestro y es dato oficial; la fila resaltada indica la clase en la que cae dicha media. "
            "Los totales y porcentajes se calculan con fórmulas sobre las filas del cuadro."
        )
        ws.cell(fila_nota, 1).alignment = Alignment(horizontal="justify", vertical="top", wrap_text=True)

    return {
        "dn_media": dn_media,
        "desviacion": desviacion,
        "fila_total": fila_total,
    }


# --------------------------------------------------------------------------
# Hoja «Resumen»
# --------------------------------------------------------------------------


def actualizar_resumen(wb, d: DatosMSAVI, sup_catalogo: float, desviacion: float) -> None:
    ws = wb["Resumen"]
    fila_ndvi = buscar_fila(ws, "Superficie clasificada NDVI", exacto=False)
    if fila_ndvi is None:
        raise ValueError("No se ubicó la fila de superficie clasificada NDVI en la hoja Resumen")

    fila_sup = buscar_fila(ws, ETQ_SUP_MSAVI)
    if fila_sup is None:
        insertar_filas(ws, fila_ndvi + 1, 3)
        for i in range(3):
            clonar_fila_estilo(ws, fila_ndvi, fila_ndvi + 1 + i, 6)
        fila_sup = fila_ndvi + 1
    fila_bajo, fila_clase = fila_sup + 1, fila_sup + 2

    valores = [
        (
            fila_sup,
            ETQ_SUP_MSAVI,
            round(d.total, 3),
            "Desviación MSAVI frente al catálogo (%)",
            round(desviacion, 2),
        ),
        (
            fila_bajo,
            ETQ_BAJO_MSAVI,
            round(d.bajo_umbral, 3),
            "% del bloque bajo umbral (brecha espectral)",
            round(d.pct_bajo_umbral, 2),
        ),
        (
            fila_clase,
            ETQ_CLASE_DOM,
            etiqueta_clase(d.dn_dominante),
            "N.° de polígonos MSAVI del bloque",
            d.poligonos,
        ),
    ]
    for fila, etq_a, val_b, etq_c, val_d in valores:
        ws.cell(fila, 1).value = etq_a
        ws.cell(fila, 2).value = val_b
        ws.cell(fila, 3).value = etq_c
        ws.cell(fila, 4).value = val_d
        ws.cell(fila, 2).number_format = "General" if isinstance(val_b, str) else "0.000"
        ws.cell(fila, 4).number_format = "0" if isinstance(val_d, int) else "0.00"

    fila_decl = buscar_fila(ws, "DECLARACIÓN DE INTEGRIDAD", exacto=False)
    if fila_decl:
        texto = ws.cell(fila_decl, 1).value or ""
        marca = "La distribución areal por clase de MSAVI 2024"
        if marca not in texto:
            ws.cell(fila_decl, 1).value = texto.rstrip() + (
                " La distribución areal por clase de MSAVI 2024 procede de la estadística zonal del ráster "
                "MSAVI clasificado (AREAS_MSAVI_BLOQUES_V6_REV_HSCM) y se detalla en la hoja "
                "«Cobertura MSAVI-NDVI»; la equivalencia entre el valor de clase DN y los umbrales del "
                "proyecto se declara en la nota metodológica de esa hoja."
            )


# --------------------------------------------------------------------------
# Hoja «Control de consistencia»
# --------------------------------------------------------------------------


def actualizar_control(wb, d: DatosMSAVI, media: float, sup_catalogo: float,
                       desviacion: float, dn_media: int, verificacion: dict) -> dict:
    ws = wb["Control de consistencia"]
    fila_resumen = buscar_fila(ws, "RESUMEN")
    if fila_resumen is None:
        raise ValueError("No se ubicó la fila RESUMEN en la hoja de control de consistencia")

    filas_d = [
        r for r in range(1, ws.max_row + 1)
        if isinstance(ws.cell(r, 1).value, str) and RE_CODIGO_D.match(str(ws.cell(r, 1).value).strip())
    ]
    ultima_d = max(filas_d)

    conforme_area = abs(desviacion) <= 2.0
    coherente = dn_media == d.dn_dominante

    contenidos = [
        (
            "Distribución areal MSAVI 2024",
            "La distribución por clase de MSAVI no figuraba en los insumos del entregable y se consignaba "
            "«Por determinar»; se incorpora la estadística zonal del ráster MSAVI clasificado "
            f"({num(d.poligonos)} polígonos, {num(d.total, 3)} ha)",
            "CORREGIDO",
            "Se completó la sección A de la hoja «Cobertura MSAVI-NDVI» con superficie y porcentaje por "
            "clase DN 1 a DN 5, totales, superficie sobre y bajo el umbral 0.4976 y contraste con el "
            "catálogo. La equivalencia DN → clase se sustenta en la nota metodológica de esa hoja "
            f"(r = {verificacion['r_directo']}, error absoluto medio = {verificacion['eam_directo']}, "
            f"n = {verificacion['n']} bloques).",
        ),
        (
            "Superficie MSAVI vs. catálogo",
            f"La superficie total clasificada por el MSAVI 2024 ({num(d.total, 3)} ha) frente a la de "
            f"catálogo V5/V6 ({num(sup_catalogo, 3)} ha): desviación de {desviacion:+.2f} %",
            "CONFORME" if conforme_area else "SUSTANTIVA",
            "Sin acción: la segmentación del ráster valida la geometría del bloque dentro del ±2 %."
            if conforme_area
            else "Debe revisarse la geometría del bloque o el recorte del ráster: la desviación excede el "
                 "±2 % admitido para el contraste planimétrico.",
        ),
        (
            "Media MSAVI vs. clase DN dominante",
            f"La media del bloque ({float(media):.4f}) cae en la clase «{etiqueta_clase(dn_media)}»; la clase "
            f"de mayor superficie es «{etiqueta_clase(d.dn_dominante)}», con {d.pct(d.dn_dominante):.2f} % "
            f"del bloque. Superficie bajo el umbral 0.4976: {num(d.bajo_umbral, 3)} ha "
            f"({d.pct_bajo_umbral:.2f} %)",
            "CONFORME" if coherente else "NO SUSTANTIVA",
            "Sin acción: la media del bloque y la clase de mayor superficie coinciden."
            if coherente
            else "Se declara la divergencia: en bloques de distribución asimétrica la media puede caer en una "
                 "clase distinta de la modal. Para el dimensionamiento de metas físicas debe usarse la "
                 "distribución areal, no la media.",
        ),
    ]

    nuevas = [c for c in contenidos if buscar_fila(ws, c[0], columna=2) is None]
    if nuevas:
        insertar_filas(ws, ultima_d + 1, len(nuevas))
        for i in range(len(nuevas)):
            clonar_fila_estilo(ws, ultima_d, ultima_d + 1 + i, 5)
            ws.row_dimensions[ultima_d + 1 + i].height = ws.row_dimensions[ultima_d].height or 48

    for campo, discrepancia, calificacion, tratamiento in contenidos:
        fila = buscar_fila(ws, campo, columna=2)
        if fila is None:
            fila = ultima_d + 1 + [c[0] for c in nuevas].index(campo)
        ws.cell(fila, 2).value = campo
        ws.cell(fila, 3).value = discrepancia
        ws.cell(fila, 4).value = calificacion
        ws.cell(fila, 5).value = tratamiento

    # Renumeración correlativa de los códigos D-xx y recuento del resumen.
    campos_nuevos = [c[0] for c in contenidos]

    def codigo_de(fila):
        valor = ws.cell(fila, 1).value
        return str(valor).strip() if isinstance(valor, str) else ""

    # Toda fila de verificación cuenta en el resumen; solo se renumera la
    # serie D, que es la que crece con este traspaso.
    filas_verificacion = sorted(
        r for r in range(1, ws.max_row + 1)
        if RE_CODIGO_VERIFICACION.match(codigo_de(r))
        or ws.cell(r, 2).value in campos_nuevos
    )
    filas_d = [r for r in filas_verificacion
               if RE_CODIGO_D.match(codigo_de(r)) or not codigo_de(r)]

    for i, fila in enumerate(filas_d, start=1):
        ws.cell(fila, 1).value = f"D-{i:02d}"

    recuento = {}
    for fila in filas_verificacion:
        calificacion = str(ws.cell(fila, 4).value or "").strip()
        recuento[calificacion] = recuento.get(calificacion, 0) + 1

    fila_resumen = buscar_fila(ws, "RESUMEN")
    ws.cell(fila_resumen, 2).value = f"{len(filas_verificacion)} verificaciones"
    orden = ["CONFORME", "CORREGIDO", "NO SUSTANTIVA", "SUSTANTIVA"]
    partes = [f"{k}: {recuento[k]}" for k in orden if recuento.get(k)]
    partes += [f"{k}: {v}" for k, v in recuento.items() if k not in orden and k]
    ws.cell(fila_resumen, 3).value = " · ".join(partes)
    return recuento


# --------------------------------------------------------------------------
# Hoja «Microcuenca»
# --------------------------------------------------------------------------


def actualizar_microcuenca(wb, datos: dict, codigo: str, d: DatosMSAVI) -> None:
    ws = wb["Microcuenca"]
    fila_hdr = buscar_fila(ws, "Bloque")
    fila_total = buscar_fila(ws, "TOTAL / PROMEDIO", exacto=False)
    if fila_hdr is None or fila_total is None:
        raise ValueError("No se ubicó el cuadro intramicrocuenca")

    col = 9  # columna I
    if ws.cell(fila_hdr, col).value not in (ETQ_COL_MICRO, None):
        col = ws.max_column + 1
    copiar_estilo(ws.cell(fila_hdr, col - 1), ws.cell(fila_hdr, col))
    ws.cell(fila_hdr, col).value = ETQ_COL_MICRO
    ws.column_dimensions[get_column_letter(col)].width = 22.0

    valores = []
    for fila in range(fila_hdr + 1, fila_total):
        etiqueta = ws.cell(fila, 1).value
        if not isinstance(etiqueta, str):
            continue
        clave = etiqueta.replace("►", "").strip()
        copiar_estilo(ws.cell(fila, col - 1), ws.cell(fila, col))
        registro = datos.get(clave)
        if registro:
            ws.cell(fila, col).value = round(registro.pct_bajo_umbral, 2)
            ws.cell(fila, col).number_format = "0.00"
            valores.append(registro.pct_bajo_umbral)
        else:
            ws.cell(fila, col).value = "s/d"
            ws.cell(fila, col).number_format = "General"

    copiar_estilo(ws.cell(fila_total, col - 1), ws.cell(fila_total, col))
    letra = get_column_letter(col)
    ws.cell(fila_total, col).value = f"=ROUND(AVERAGE({letra}{fila_hdr + 1}:{letra}{fila_total - 1}),2)"
    ws.cell(fila_total, col).number_format = "0.00"

    # Las notas al pie se extienden a la nueva columna.
    for rango in list(ws.merged_cells.ranges):
        if rango.max_col == col - 1 and rango.min_col == 1:
            rango.max_col = col

    promedio = statistics.mean(valores) if valores else None
    fila_lectura = buscar_fila(ws, "LECTURA INTRAMICROCUENCA", exacto=False)
    if fila_lectura and promedio is not None:
        texto = ws.cell(fila_lectura, 1).value or ""
        marca = "BRECHA ESPECTRAL."
        base = texto.split(marca)[0].rstrip()
        comparacion = (
            "por encima del promedio local"
            if d.pct_bajo_umbral > promedio
            else "por debajo del promedio local"
            if d.pct_bajo_umbral < promedio
            else "igual al promedio local"
        )
        ws.cell(fila_lectura, 1).value = (
            f"{base} {marca} El {d.pct_bajo_umbral:.2f} % de la superficie del bloque "
            f"({num(d.bajo_umbral, 3)} ha de {num(d.total, 3)} ha) se sitúa bajo el umbral MSAVI 0.4976, "
            f"{comparacion} de la microcuenca ({promedio:.2f} %) calculado sobre los bloques con dato MSAVI "
            "disponible."
        )

    fila_nota = buscar_fila(ws, "NOTA METODOLÓGICA", exacto=False)
    if fila_nota:
        texto = ws.cell(fila_nota, 1).value or ""
        marca = "El porcentaje bajo el umbral MSAVI 0.4976"
        if marca not in texto:
            ws.cell(fila_nota, 1).value = texto.rstrip() + (
                f" {marca} procede de la estadística zonal del ráster MSAVI 2024 clasificado "
                "(AREAS_MSAVI_BLOQUES_V6_REV_HSCM) y corresponde a la suma de las clases DN 1 a DN 3 sobre la "
                "superficie clasificada de cada bloque; «s/d» indica bloques del catálogo sin cobertura en ese "
                "archivo."
            )


# --------------------------------------------------------------------------
# Proceso por plantilla
# --------------------------------------------------------------------------


def procesar_plantilla(ruta: str, codigo: str, datos: dict, verificacion: dict,
                       destino: str, rep: Reporte) -> None:
    d = datos[codigo]
    wb = openpyxl.load_workbook(ruta)

    media = wb["Cobertura MSAVI-NDVI"][
        f"B{buscar_fila(wb['Cobertura MSAVI-NDVI'], 'MSAVI 2024 — MEDIA', exacto=False)}"
    ].value
    fila_sup = buscar_fila(wb["Resumen"], "Superficie de catálogo", exacto=False)
    sup_catalogo = float(wb["Resumen"].cell(fila_sup, 2).value)

    info = actualizar_cobertura(wb, d, float(media), sup_catalogo, rep)
    actualizar_resumen(wb, d, sup_catalogo, info["desviacion"])
    recuento = actualizar_control(
        wb, d, float(media), sup_catalogo, info["desviacion"], info["dn_media"], verificacion
    )
    actualizar_microcuenca(wb, datos, codigo, d)

    # Excel recalcula al abrir; el valor en caché sirve a quien lea valores.
    wb.calculation.fullCalcOnLoad = True
    cacheados = calcular_valores_cacheados(wb)
    wb.save(destino)
    escritos = inyectar_valores_cacheados(destino, cacheados)
    esperados = sum(len(v) for v in cacheados.values())
    if escritos != esperados:
        raise ValueError(
            f"{codigo}: se cachearon {escritos} de {esperados} fórmulas")

    rep.filas.append(
        {
            "bloque": codigo,
            "archivo": os.path.basename(destino),
            "poligonos": d.poligonos,
            "dn1_ha": round(d.areas[1], 4),
            "dn2_ha": round(d.areas[2], 4),
            "dn3_ha": round(d.areas[3], 4),
            "dn4_ha": round(d.areas[4], 4),
            "dn5_ha": round(d.areas[5], 4),
            "total_ha": round(d.total, 4),
            "catalogo_ha": round(sup_catalogo, 4),
            "desviacion_pct": round(info["desviacion"], 3),
            "media_msavi": round(float(media), 6),
            "dn_media": info["dn_media"],
            "dn_dominante": d.dn_dominante,
            "bajo_umbral_ha": round(d.bajo_umbral, 4),
            "bajo_umbral_pct": round(d.pct_bajo_umbral, 2),
            "verificaciones": recuento,
        }
    )
    if abs(info["desviacion"]) > 2.0:
        rep.avisos.append(f"Bloque {codigo}: desviación de superficie {info['desviacion']:+.2f} % (> ±2 %)")
    if info["dn_media"] != d.dn_dominante:
        rep.avisos.append(
            f"Bloque {codigo}: la media cae en DN {info['dn_media']} y la clase modal es DN {d.dn_dominante}"
        )


# --------------------------------------------------------------------------
# Reporte consolidado
# --------------------------------------------------------------------------


def escribir_reporte(rep: Reporte, verificacion: dict, ruta: str, fuente: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Traspaso MSAVI"

    encabezados_inst = [
        "AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN",
        "DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME",
        "SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN - SESDI",
        "REPORTE DE TRASPASO DE ÁREAS MSAVI POR CLASE DN A LAS PLANTILLAS DE BLOQUE",
        "Proyecto IN Piura — CUI 2669244 — Cuenca Alta del Río Piura",
        f"Fuente: {os.path.basename(fuente)} · hoja Resumen_por_Bloque · superficies en hectáreas",
        (
            f"Verificación de la equivalencia DN → clase MSAVI sobre {verificacion['n']} bloques: "
            f"r = {verificacion['r_directo']} y error absoluto medio = {verificacion['eam_directo']} "
            f"con el orden DN 1 = clase inferior … DN 5 = clase superior "
            f"(orden inverso: r = {verificacion['r_inverso']}, eam = {verificacion['eam_inverso']})"
        ),
    ]
    for i, texto in enumerate(encabezados_inst, start=1):
        celda = ws.cell(i, 1, texto)
        celda.font = Font(name="Arial", size=11 if i <= 3 else 10, bold=True,
                          color="FFFFFF" if i <= 3 else "1B4D2E")
        if i <= 3:
            celda.fill = PatternFill("solid", fgColor=VERDE_ANIN)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=17)
        ws.row_dimensions[i].height = 18 if i <= 3 else 24
    ws.row_dimensions[7].height = 40
    ws.cell(7, 1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    columnas = [
        ("BLOQUE", "bloque", 12, "General"),
        ("Archivo actualizado", "archivo", 42, "General"),
        ("N° Polígonos", "poligonos", 12, "#,##0"),
        ("DN 1 ≤ 0.2650 (ha)", "dn1_ha", 15, "0.0000"),
        ("DN 2 0.2650-0.3813 (ha)", "dn2_ha", 16, "0.0000"),
        ("DN 3 0.3813-0.4976 (ha)", "dn3_ha", 16, "0.0000"),
        ("DN 4 0.4976-0.6139 (ha)", "dn4_ha", 16, "0.0000"),
        ("DN 5 > 0.6139 (ha)", "dn5_ha", 15, "0.0000"),
        ("Total MSAVI (ha)", "total_ha", 14, "0.0000"),
        ("Catálogo V5/V6 (ha)", "catalogo_ha", 14, "0.0000"),
        ("Desviación (%)", "desviacion_pct", 12, "0.000"),
        ("MSAVI media", "media_msavi", 12, "0.000000"),
        ("DN de la media", "dn_media", 12, "0"),
        ("DN dominante", "dn_dominante", 12, "0"),
        ("Bajo umbral 0.4976 (ha)", "bajo_umbral_ha", 16, "0.0000"),
        ("Bajo umbral (%)", "bajo_umbral_pct", 13, "0.00"),
        ("Verificaciones de consistencia", "verif", 34, "General"),
    ]
    fila_hdr = 9
    for j, (titulo, _clave, ancho, _fmt) in enumerate(columnas, start=1):
        celda = ws.cell(fila_hdr, j, titulo)
        celda.font = Font(name="Arial", size=9, bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=VERDE_ANIN)
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = bordes_finos()
        ws.column_dimensions[get_column_letter(j)].width = ancho
    ws.row_dimensions[fila_hdr].height = 30

    def clave_orden(f):
        m = re.match(r"^(\d+)$", f["bloque"])
        return (0, int(m.group(1)), "") if m else (1, 0, f["bloque"])

    for i, registro in enumerate(sorted(rep.filas, key=clave_orden)):
        fila = fila_hdr + 1 + i
        for j, (_titulo, clave, _ancho, fmt) in enumerate(columnas, start=1):
            if clave == "verif":
                orden = ["CONFORME", "CORREGIDO", "NO SUSTANTIVA", "SUSTANTIVA"]
                conteo = registro["verificaciones"]
                claves = [k for k in orden if conteo.get(k)] + [k for k in conteo if k and k not in orden]
                valor = " · ".join(f"{k}: {conteo[k]}" for k in claves)
            else:
                valor = registro[clave]
            celda = ws.cell(fila, j, valor)
            celda.font = Font(name="Arial", size=9)
            celda.number_format = fmt
            celda.border = bordes_finos()
            celda.alignment = Alignment(
                horizontal="left" if j in (1, 2, 17) else "center", vertical="center", wrap_text=False
            )
            if i % 2 == 1:
                celda.fill = PatternFill("solid", fgColor=VERDE_CLARO)

    fila_tot = fila_hdr + 1 + len(rep.filas)
    ws.cell(fila_tot, 1, "TOTAL")
    for j, (_titulo, clave, _ancho, fmt) in enumerate(columnas, start=1):
        celda = ws.cell(fila_tot, j)
        celda.font = Font(name="Arial", size=9, bold=True)
        celda.fill = PatternFill("solid", fgColor=VERDE_CLARO)
        celda.border = bordes_finos()
        celda.alignment = Alignment(horizontal="left" if j in (1, 2, 17) else "center", vertical="center")
        if clave in {"poligonos", "dn1_ha", "dn2_ha", "dn3_ha", "dn4_ha", "dn5_ha",
                     "total_ha", "catalogo_ha", "bajo_umbral_ha"}:
            letra = get_column_letter(j)
            celda.value = f"=SUM({letra}{fila_hdr + 1}:{letra}{fila_tot - 1})"
            celda.number_format = fmt
    ws.freeze_panes = f"A{fila_hdr + 1}"

    ws.cell(fila_tot + 2, 1, (
        "NOTA. Cada fila corresponde a una plantilla de bloque actualizada. Las superficies por clase DN "
        "proceden íntegramente del archivo fuente y no han sido estimadas ni redistribuidas. La columna "
        "«Desviación» contrasta la superficie clasificada por el MSAVI con la superficie de catálogo V5/V6 "
        "declarada en cada plantilla; el criterio de conformidad es ±2 %."
    ))
    ws.merge_cells(start_row=fila_tot + 2, start_column=1, end_row=fila_tot + 2, end_column=17)
    ws.cell(fila_tot + 2, 1).font = Font(name="Arial", size=8, color="333333")
    ws.cell(fila_tot + 2, 1).alignment = Alignment(horizontal="justify", vertical="top", wrap_text=True)
    ws.row_dimensions[fila_tot + 2].height = 40

    if rep.avisos:
        ws2 = wb.create_sheet("Avisos")
        ws2.column_dimensions["A"].width = 120
        ws2.cell(1, 1, "AVISOS DE LA EJECUCIÓN").font = Font(name="Arial", size=10, bold=True, color="1B4D2E")
        for i, aviso in enumerate(rep.avisos, start=3):
            ws2.cell(i, 1, aviso).font = Font(name="Arial", size=9)

    wb.save(ruta)


# --------------------------------------------------------------------------
# Programa principal
# --------------------------------------------------------------------------


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--fuente", default="AREAS_MSAVI_BLOQUES_V6_REV_HSCM.xlsx")
    parser.add_argument("--plantillas", default="Plantillas Excel 117 bloques",
                        help="carpeta con las plantillas o ruta del ZIP que las contiene")
    parser.add_argument("--salida", default="plantillas_117_msavi_v6")
    parser.add_argument("--reporte", default="REPORTE_TRASPASO_MSAVI_117_BLOQUES.xlsx")
    parser.add_argument("--zip", dest="zip_salida", default=None,
                        help="ruta del ZIP a generar con las plantillas actualizadas")
    parser.add_argument("--json", dest="json_salida", default=None)
    args = parser.parse_args(argv)

    datos = leer_fuente(args.fuente)

    origen = args.plantillas
    temporal = None
    if origen.lower().endswith(".zip"):
        temporal = tempfile.mkdtemp(prefix="plantillas_msavi_")
        with zipfile.ZipFile(origen) as zf:
            zf.extractall(temporal)
        origen = temporal
    carpetas = [
        raiz for raiz, _dirs, archivos in os.walk(origen)
        if any(RE_PLANTILLA.match(a) for a in archivos)
    ]
    rutas = sorted(
        os.path.join(carpeta, n)
        for carpeta in carpetas
        for n in os.listdir(carpeta)
        if RE_PLANTILLA.match(n)
    )
    if not rutas:
        print(f"No se hallaron plantillas en «{args.plantillas}»", file=sys.stderr)
        return 1

    # Medias MSAVI del catálogo, para verificar la equivalencia DN → clase.
    medias, codigos = {}, {}
    for ruta in rutas:
        codigo = RE_PLANTILLA.match(os.path.basename(ruta)).group(1)
        codigos[ruta] = codigo
        wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
        ws = wb["Cobertura MSAVI-NDVI"]
        for fila in ws.iter_rows(max_col=2, values_only=True):
            if isinstance(fila[0], str) and fila[0].startswith("MSAVI 2024 — MEDIA"):
                medias[codigo] = fila[1]
                break
        wb.close()

    faltantes = [c for c in codigos.values() if c not in datos]
    if faltantes:
        print(f"Sin datos MSAVI para los bloques: {', '.join(faltantes)}", file=sys.stderr)
        return 1

    verificacion = verificar_equivalencia_dn(datos, medias)
    print(
        f"Equivalencia DN → clase verificada sobre {verificacion['n']} bloques: "
        f"r = {verificacion['r_directo']} (eam {verificacion['eam_directo']}) frente a "
        f"r = {verificacion['r_inverso']} (eam {verificacion['eam_inverso']}) del orden inverso"
    )
    if verificacion["r_directo"] < 0.9:
        print("La equivalencia DN → clase no se sustenta; se aborta el traspaso.", file=sys.stderr)
        return 2

    os.makedirs(args.salida, exist_ok=True)
    rep = Reporte()
    for ruta in rutas:
        codigo = codigos[ruta]
        destino = os.path.join(args.salida, os.path.basename(ruta))
        procesar_plantilla(ruta, codigo, datos, verificacion, destino, rep)
        print(f"  · Bloque {codigo}: {len(datos[codigo].areas)} clases DN escritas → {os.path.basename(destino)}")

    escribir_reporte(rep, verificacion, args.reporte, args.fuente)
    print(f"Plantillas actualizadas: {len(rep.filas)} → «{args.salida}»")
    print(f"Reporte de traspaso: «{args.reporte}»")

    if args.json_salida:
        with open(args.json_salida, "w", encoding="utf-8") as fh:
            json.dump({"verificacion": verificacion, "bloques": rep.filas, "avisos": rep.avisos},
                      fh, ensure_ascii=False, indent=2)
        print(f"Bitácora JSON: «{args.json_salida}»")

    if args.zip_salida:
        with zipfile.ZipFile(args.zip_salida, "w", zipfile.ZIP_DEFLATED) as zf:
            for nombre in sorted(os.listdir(args.salida)):
                zf.write(os.path.join(args.salida, nombre), os.path.join(os.path.basename(args.salida), nombre))
        print(f"ZIP para carga en Drive: «{args.zip_salida}»")

    for aviso in rep.avisos:
        print(f"  AVISO: {aviso}")

    if temporal:
        shutil.rmtree(temporal, ignore_errors=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
