#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Completa la informacion del indice MSAVI 2024 en las plantillas Excel de los
117 bloques del Proyecto IN Piura.

AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN
DIRECCION DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME
SUBDIRECCION DE ESTUDIOS DE INVERSION - SESDI
Proyecto IN Piura | CUI 2669244 | UTM WGS 84 Zona 17S (EPSG:32717)

Escribe, en cada libro `Plantilla_Excel_Bloque_<codigo>_IN_Piura.xlsx`:

  1. Hoja "Cobertura MSAVI-NDVI": seccion A, las cinco filas de clase MSAVI
     ("Superficie (ha)" y "% del area clasificada") que hoy figuran como
     "Por determinar", y reemplaza la NOTA METODOLOGICA de la seccion.
  2. Hoja "Resumen": seccion 3, enriquece "Condicion frente al umbral 0.4976"
     y agrega las filas de distribucion areal del MSAVI.
  3. Hoja "Control de consistencia": agrega la verificacion D-xx que deja
     constancia del origen del dato y actualiza la fila RESUMEN.

Las hectareas NO se toman del CSV: se calculan sobre la "Superficie de catalogo
(V5/V6), ha" que declara cada libro, de modo que el total por clase cuadra
exactamente con la superficie oficial del bloque. Del CSV solo se toma el
reparto porcentual medido sobre la cartografia MSAVI.

Uso:
    python completar_msavi_plantillas.py --dir <carpeta_con_los_117_xlsx>
    python completar_msavi_plantillas.py --dir salida --solo 38 M22B1
    python completar_msavi_plantillas.py --dir salida --simular      # no escribe
    python completar_msavi_plantillas.py --dir salida --sin-insertar # no crea filas
"""

import argparse
import copy
import csv
import os
import re
import shutil
import sys
import unicodedata

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.exit("Falta openpyxl.  Instalar con:  pip install openpyxl")

CSV_POR_DEFECTO = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               "msavi_distribucion_117_bloques.csv")

UMBRAL = 0.4976

# Etiqueta de la fila en la hoja "Cobertura MSAVI-NDVI" -> (columna % en el CSV,
# columna ha en el CSV).  El orden es el de la propia hoja.
CLASES = [
    ("> 0.6139",        "PCT_MAYOR_0_6139",       "Vigor alto"),
    ("0.4976 - 0.6139", "PCT_0_4976_0_6139",      "Vigor moderado"),
    ("0.3813 - 0.4976", "PCT_0_3813_0_4976",      "Vigor bajo"),
    ("0.2650 - 0.3813", "PCT_0_2650_0_3813",      "Vigor muy bajo"),
    ("<= 0.2650",       "PCT_MENOR_IGUAL_0_2650", "Suelo desnudo / no vegetal"),
]

NOTA_NUEVA = (
    "NOTA METODOLOGICA. La media del MSAVI 2024 del bloque procede del catalogo "
    "maestro de bloques V5/V6 y es dato oficial. La DISTRIBUCION AREAL por clase "
    "se obtuvo por conteo de celdas sobre la cartografia tematica MSAVI 2024 del "
    "bloque (carpeta Drive «MSAVI BLOQUES V6»), clasificada con la paleta RdYlGn "
    "de cinco clases y los mismos umbrales del proyecto. El reparto porcentual es "
    "el medido sobre el poligono del bloque; las hectareas se obtienen aplicando "
    "ese reparto a la superficie de catalogo (V5/V6) declarada en la hoja "
    "«Resumen», de modo que el total por clase cuadra con la superficie oficial. "
    "Umbral de brecha 0.4976 conforme a la R.M. N.o 00213-2024-MINAM. "
    "Contraste independiente: la tabla de interseccion vectorial "
    "«BLOQUES_V6_INTERSECC_MSAVI» reproduce el mismo ordenamiento de bloques "
    "(correlacion de Pearson 0.949 sobre el porcentaje bajo umbral en los 117 "
    "bloques); no se usa para el metrado porque su campo AREA_M2 conserva el area "
    "del poligono padre y no la de cada pieza de interseccion."
)

TXT_CONSISTENCIA = (
    "La hoja «Cobertura MSAVI-NDVI» consignaba «Por determinar» en la distribucion "
    "areal de clases MSAVI 2024; se incorpora el metrado por clase medido sobre la "
    "cartografia tematica MSAVI 2024 del bloque"
)


# --------------------------------------------------------------------------- #
# utilidades
# --------------------------------------------------------------------------- #
def normaliza(texto):
    """Minusculas, sin acentos y con espacios colapsados, para comparar rotulos."""
    if texto is None:
        return ""
    txt = unicodedata.normalize("NFKD", str(texto))
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    txt = txt.replace(" ", " ").replace("—", "-").replace("–", "-")
    return re.sub(r"\s+", " ", txt).strip().lower()


def busca_fila(hoja, rotulo, col_max=8):
    """Devuelve (fila, columna) de la primera celda cuyo texto normalizado
    empieza por `rotulo` normalizado.  None si no aparece."""
    objetivo = normaliza(rotulo)
    for fila in hoja.iter_rows(min_row=1, max_row=hoja.max_row, max_col=col_max):
        for celda in fila:
            if celda.value is None:
                continue
            if normaliza(celda.value).startswith(objetivo):
                return celda.row, celda.column
    return None


def celda_escribible(hoja, fila, columna):
    """Devuelve la celda ancla: si (fila, columna) esta dentro de una combinacion,
    openpyxl solo admite escritura en la esquina superior izquierda."""
    for rango in hoja.merged_cells.ranges:
        if (rango.min_row <= fila <= rango.max_row
                and rango.min_col <= columna <= rango.max_col):
            return hoja.cell(row=rango.min_row, column=rango.min_col)
    return hoja.cell(row=fila, column=columna)


def escribe(hoja, fila, columna, valor, formato=None):
    celda = celda_escribible(hoja, fila, columna)
    celda.value = valor
    if formato:
        celda.number_format = formato
    return celda


def copia_estilo(origen, destino):
    if origen.has_style:
        destino.font = copy.copy(origen.font)
        destino.border = copy.copy(origen.border)
        destino.fill = copy.copy(origen.fill)
        destino.alignment = copy.copy(origen.alignment)
        destino.number_format = origen.number_format


def inserta_fila_como(hoja, fila_modelo, fila_destino, col_max=8):
    """Inserta una fila en `fila_destino` heredando el estilo de `fila_modelo`."""
    hoja.insert_rows(fila_destino)
    modelo = fila_modelo if fila_modelo < fila_destino else fila_modelo + 1
    for col in range(1, col_max + 1):
        copia_estilo(hoja.cell(row=modelo, column=col),
                     hoja.cell(row=fila_destino, column=col))
    return fila_destino


def a_float(valor):
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    txt = str(valor).strip().replace(" ", "").replace(",", "")
    m = re.search(r"-?\d+(?:\.\d+)?", txt)
    return float(m.group()) if m else None


# --------------------------------------------------------------------------- #
# hojas
# --------------------------------------------------------------------------- #
def superficie_catalogo(libro):
    """Superficie de catalogo (V5/V6) declarada en la hoja Resumen, en ha."""
    if "Resumen" not in libro.sheetnames:
        return None
    hoja = libro["Resumen"]
    pos = busca_fila(hoja, "Superficie de catalogo")
    if not pos:
        return None
    fila, col = pos
    return a_float(hoja.cell(row=fila, column=col + 1).value)


def completa_cobertura(libro, datos, area_ha, avisos):
    if "Cobertura MSAVI-NDVI" not in libro.sheetnames:
        avisos.append("no existe la hoja 'Cobertura MSAVI-NDVI'")
        return 0
    hoja = libro["Cobertura MSAVI-NDVI"]
    escritas = 0
    for rotulo, columna_csv, _interpretacion in CLASES:
        pos = busca_fila(hoja, rotulo, col_max=2)
        if not pos:
            avisos.append("no se hallo la fila de clase %r" % rotulo)
            continue
        fila, col = pos
        pct = float(datos[columna_csv])
        escribe(hoja, fila, col + 1, round(area_ha * pct / 100.0, 4), "0.0000")
        escribe(hoja, fila, col + 2, round(pct, 2), "0.00")
        escritas += 1

    pos = busca_fila(hoja, "NOTA METODOLOGICA", col_max=8)
    if pos:
        escribe(hoja, pos[0], pos[1], NOTA_NUEVA)
    else:
        avisos.append("no se hallo la NOTA METODOLOGICA de la hoja de cobertura")
    return escritas


def completa_resumen(libro, datos, area_ha, insertar, avisos):
    if "Resumen" not in libro.sheetnames:
        avisos.append("no existe la hoja 'Resumen'")
        return 0
    hoja = libro["Resumen"]
    bajo = float(datos["PCT_BAJO_UMBRAL_0_4976"])
    ha_bajo = round(area_ha * bajo / 100.0, 3)
    condicion = "Sobre umbral" if bajo < 50 else "Bajo umbral"

    pos = busca_fila(hoja, "Condicion frente al umbral")
    if not pos:
        avisos.append("no se hallo 'Condicion frente al umbral' en Resumen")
        return 0
    fila, col = pos
    escribe(hoja, fila, col + 1,
            "%s - distribucion areal: %.2f %% del bloque BAJO el umbral "
            "(%.3f ha de %.3f ha)" % (condicion, bajo, ha_bajo, area_ha))

    if not insertar:
        return 1

    nuevas = [
        ("MSAVI 2024 - superficie BAJO el umbral 0.4976 (ha)", ha_bajo,
         "MSAVI 2024 - % del bloque BAJO el umbral", round(bajo, 2)),
        ("MSAVI 2024 - clase areal dominante", datos["CLASE_MODAL"],
         "MSAVI 2024 - fuente de la distribucion areal",
         "Conteo de celdas sobre cartografia MSAVI 2024 del bloque (%s)"
         % datos["MAPA_PNG"]),
    ]
    for desplazamiento, (rot_a, val_a, rot_b, val_b) in enumerate(nuevas):
        destino = fila + 1 + desplazamiento
        inserta_fila_como(hoja, fila, destino)
        escribe(hoja, destino, col, rot_a)
        escribe(hoja, destino, col + 1, val_a)
        escribe(hoja, destino, col + 2, rot_b)
        escribe(hoja, destino, col + 3, val_b)
    return 3


def completa_consistencia(libro, datos, area_ha, insertar, avisos):
    if "Control de consistencia" not in libro.sheetnames:
        avisos.append("no existe la hoja 'Control de consistencia'")
        return 0
    hoja = libro["Control de consistencia"]
    pos = busca_fila(hoja, "RESUMEN", col_max=2)
    if not pos:
        avisos.append("no se hallo la fila RESUMEN en Control de consistencia")
        return 0
    fila_resumen, col = pos

    # ultimo correlativo D-xx por encima de la fila RESUMEN
    ultimo = 0
    for f in range(1, fila_resumen):
        m = re.fullmatch(r"D-(\d+)", str(hoja.cell(row=f, column=col).value or "").strip())
        if m:
            ultimo = max(ultimo, int(m.group(1)))
    codigo = "D-%02d" % (ultimo + 1)

    bajo = float(datos["PCT_BAJO_UMBRAL_0_4976"])
    detalle = ("%s: %.2f %% del bloque (%.3f ha) bajo el umbral 0.4976 y %.2f %% "
               "sobre el umbral; clase areal dominante «%s»."
               % (TXT_CONSISTENCIA, bajo, area_ha * bajo / 100.0, 100 - bajo,
                  datos["CLASE_MODAL"]))
    tratamiento = ("Se sustituye «Por determinar» por el metrado por clase. Las "
                   "hectareas se derivan del reparto porcentual medido aplicado a la "
                   "superficie de catalogo (V5/V6) del bloque, por lo que la suma de "
                   "las cinco clases iguala dicha superficie. Pendiente de refrendo "
                   "con estadistica zonal directa sobre el raster MSAVI 2024 "
                   "cuando se disponga del insumo.")

    if insertar:
        inserta_fila_como(hoja, fila_resumen - 1, fila_resumen)
        escribe(hoja, fila_resumen, col, codigo)
        escribe(hoja, fila_resumen, col + 1, "Distribucion areal de clases MSAVI 2024")
        escribe(hoja, fila_resumen, col + 2, detalle)
        escribe(hoja, fila_resumen, col + 3, "CORREGIDO")
        escribe(hoja, fila_resumen, col + 4, tratamiento)
        fila_resumen += 1

    # actualiza los conteos de la fila RESUMEN
    conteo = {"CONFORME": 0, "NO SUSTANTIVA": 0, "SUSTANTIVA": 0, "CORREGIDO": 0}
    total = 0
    for f in range(1, fila_resumen):
        if not re.fullmatch(r"D-\d+", str(hoja.cell(row=f, column=col).value or "").strip()):
            continue
        total += 1
        cal = normaliza(hoja.cell(row=f, column=col + 3).value)
        for clave in conteo:
            if cal == normaliza(clave):
                conteo[clave] += 1
    if total:
        escribe(hoja, fila_resumen, col + 1, "%d verificaciones" % total)
        escribe(hoja, fila_resumen, col + 2,
                "CONFORME: %d - NO SUSTANTIVA: %d - SUSTANTIVA: %d - CORREGIDO: %d"
                % (conteo["CONFORME"], conteo["NO SUSTANTIVA"],
                   conteo["SUSTANTIVA"], conteo["CORREGIDO"]))
    return 1


# --------------------------------------------------------------------------- #
def procesa_libro(ruta, datos, insertar, simular):
    avisos = []
    libro = openpyxl.load_workbook(ruta)
    area = superficie_catalogo(libro)
    if area is None or area <= 0:
        area = float(datos["AREA_HA"])
        avisos.append("sin 'Superficie de catalogo' legible; se usa el area de "
                      "la tabla de interseccion (%.3f ha)" % area)

    n1 = completa_cobertura(libro, datos, area, avisos)
    n2 = completa_resumen(libro, datos, area, insertar, avisos)
    n3 = completa_consistencia(libro, datos, area, insertar, avisos)

    if not simular:
        respaldo = ruta + ".bak"
        if not os.path.exists(respaldo):
            shutil.copy2(ruta, respaldo)
        libro.save(ruta)
    libro.close()
    return n1, n2, n3, area, avisos


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--dir", required=True, help="carpeta con los libros .xlsx")
    ap.add_argument("--csv", default=CSV_POR_DEFECTO, help="CSV de distribucion MSAVI")
    ap.add_argument("--solo", nargs="*", default=None, help="codigos de bloque a procesar")
    ap.add_argument("--simular", action="store_true", help="no escribe los libros")
    ap.add_argument("--sin-insertar", dest="insertar", action="store_false",
                    help="no crea filas nuevas; solo rellena celdas existentes")
    args = ap.parse_args(argv)

    with open(args.csv, encoding="utf-8") as fh:
        tabla = {fila["BLOQUE"]: fila for fila in csv.DictReader(fh, delimiter=";")}

    codigos = args.solo if args.solo else sorted(tabla)
    ok = fallos = 0
    for codigo in codigos:
        datos = tabla.get(codigo)
        if datos is None:
            print("  [!] %-8s sin fila en el CSV" % codigo)
            fallos += 1
            continue
        ruta = os.path.join(args.dir, "Plantilla_Excel_Bloque_%s_IN_Piura.xlsx" % codigo)
        if not os.path.exists(ruta):
            print("  [!] %-8s no se encuentra %s" % (codigo, os.path.basename(ruta)))
            fallos += 1
            continue
        try:
            n1, n2, n3, area, avisos = procesa_libro(ruta, datos, args.insertar, args.simular)
        except Exception as exc:                      # noqa: BLE001
            print("  [X] %-8s ERROR: %s" % (codigo, exc))
            fallos += 1
            continue
        estado = "simulado" if args.simular else "escrito "
        print("  [%s] %-8s %.3f ha | clases=%d resumen=%d consistencia=%d%s"
              % (estado, codigo, area, n1, n2, n3,
                 ("  <- " + "; ".join(avisos)) if avisos else ""))
        ok += 1

    print("\nLibros procesados: %d   con incidencia: %d" % (ok, fallos))
    return 1 if fallos else 0


if __name__ == "__main__":
    sys.exit(main())
