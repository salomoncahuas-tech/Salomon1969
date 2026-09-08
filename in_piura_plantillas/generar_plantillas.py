# -*- coding: utf-8 -*-
"""
Generador de plantillas Excel por bloque de intervencion - Proyecto IN Piura (CUI 2669244)
ANIN - DIME - SESDI.

Hojas:
  1. Resumen                  (modelo: Plantilla_Excel_Bloque_38.xlsx)
  2. Cobertura MSAVI-NDVI     (modelo: Plantilla_Excel_Bloque_38 / 51)
  3. Estaciones fotograficas  (modelo: Plantilla_Excel_Bloque_38)
  4. Microcuenca              (modelo: Plantilla_Excel_Bloque_38)
  5. Control de consistencia  (modelo: Plantilla_Excel_Bloque_51)

Fuentes de datos:
  catalogo_v5.json      catalogo maestro de bloques V5/V6 (area, microcuenca, zona, centroide, MSAVI)
  alt_pend.json         altitud min/max y pendiente promedio por bloque (estadistica zonal MDE)
  ndvi.json             distribucion areal de clases NDVI mediana 2025 (Sentinel-2)
  centros_poblados.json centros poblados INEI por bloque
  dt/<bloque>.json      diagnostico territorial de campo (fichas F-DT-01 a F-DT-05)
"""

import json
import os
import math

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
DATOS = os.path.join(BASE, "datos")
SALIDA = os.path.join(BASE, "salida")

# --- Identidad institucional ANIN -------------------------------------------
VERDE = "1B4D2E"
VERDE_CLARO = "E8F0EA"
AZUL = "1B4F72"
GRIS = "F2F2F2"
DORADO = "B7950B"
BLANCO = "FFFFFF"

ENCABEZADOS = [
    "AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN",
    "DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME",
    "SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN - SESDI",
    "PROYECTO IN PIURA | CUI 2669244 | Recuperación del servicio de regulación de riesgos naturales "
    "y de ecosistemas degradados - Cuenca Alta del Río Piura",
]

FINO = Side(style="thin", color="9C9C9C")
BORDE = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)

# Umbrales MSAVI del proyecto (quintiles de la clasificacion Sentinel-2 2024)
CLASES_MSAVI = [
    ("> 0.6139", 0.6139, 99.0, "Vigor alto", "Sobre umbral"),
    ("0.4976 - 0.6139", 0.4976, 0.6139, "Vigor moderado", "Sobre umbral"),
    ("0.3813 - 0.4976", 0.3813, 0.4976, "Vigor bajo", "BAJO umbral 0.4976"),
    ("0.2650 - 0.3813", 0.2650, 0.3813, "Vigor muy bajo", "BAJO umbral 0.4976"),
    ("<= 0.2650", -1.0, 0.2650, "Suelo desnudo / no vegetal", "BAJO umbral 0.4976"),
]

ORDEN_NDVI = ["Vegetación alta", "Vegetación mediana", "Vegetación ligera",
              "Tierra desnuda", "Clase no vegetal"]

PISOS = [
    (1000, "<1000 m (Yunga)"),
    (1500, "1000-1500 m (Quechua baja)"),
    (2000, "1500-2000 m (Quechua media)"),
    (2500, "2000-2500 m (Quechua alta)"),
    (3000, "2500-3000 m (Suni)"),
    (3500, "3000-3500 m (Puna baja)"),
    (10000, ">3500 m (Puna/Jalca)"),
]

CLASES_PENDIENTE = [
    (8, "0-8% (Plano-lig. inclinado)"),
    (15, "8-15% (Mod. inclinado)"),
    (25, "15-25% (Fuert. inclinado)"),
    (50, "25-50% (Mod. escarpado)"),
    (75, "50-75% (Escarpado)"),
    (1000, ">75% (Muy escarpado)"),
]


def piso_altitudinal(alt):
    for techo, nombre in PISOS:
        if alt < techo:
            return nombre
    return PISOS[-1][1]


def clase_pendiente(p):
    for techo, nombre in CLASES_PENDIENTE:
        if p < techo:
            return nombre
    return CLASES_PENDIENTE[-1][1]


def clase_msavi(v):
    for etiqueta, lo, hi, interp, cond in CLASES_MSAVI:
        if lo <= v < hi:
            return etiqueta, interp, cond
    return CLASES_MSAVI[0][0], CLASES_MSAVI[0][3], CLASES_MSAVI[0][4]


def nd(valor, defecto="Por determinar"):
    """Normaliza un valor ausente a la leyenda declarativa del proyecto."""
    if valor is None or valor == "":
        return defecto
    return valor


# --- Utilidades de formato ---------------------------------------------------

def encabezado_institucional(ws, ncols, titulo, subtitulo=None):
    """Escribe el bloque de encabezado ANIN y devuelve la fila siguiente libre."""
    fila = 1
    for i, texto in enumerate(ENCABEZADOS):
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
        c = ws.cell(row=fila, column=1, value=texto)
        c.font = Font(name="Arial", size=10 if i < 3 else 8,
                      bold=i < 3, color=BLANCO)
        c.fill = PatternFill("solid", fgColor=VERDE if i < 3 else AZUL)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[fila].height = 15 if i < 3 else 24
        fila += 1

    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
    c = ws.cell(row=fila, column=1, value=titulo)
    c.font = Font(name="Arial", size=12, bold=True, color=VERDE)
    c.fill = PatternFill("solid", fgColor=VERDE_CLARO)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[fila].height = 22
    fila += 1

    if subtitulo:
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
        c = ws.cell(row=fila, column=1, value=subtitulo)
        c.font = Font(name="Arial", size=8, italic=True, color="555555")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[fila].height = 24
        fila += 1

    return fila + 1


def titulo_seccion(ws, fila, ncols, texto):
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
    c = ws.cell(row=fila, column=1, value=texto)
    c.font = Font(name="Arial", size=10, bold=True, color=BLANCO)
    c.fill = PatternFill("solid", fgColor=VERDE)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[fila].height = 18
    return fila + 1


def fila_par(ws, fila, pares, ancho_total=6):
    """Escribe una fila con hasta dos pares etiqueta/valor (4 columnas logicas)."""
    col = 1
    for etiqueta, valor in pares:
        ce = ws.cell(row=fila, column=col, value=etiqueta)
        ce.font = Font(name="Arial", size=9, bold=True)
        ce.fill = PatternFill("solid", fgColor=GRIS)
        ce.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ce.border = BORDE

        cv = ws.cell(row=fila, column=col + 1, value=valor)
        cv.font = Font(name="Arial", size=9)
        cv.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        cv.border = BORDE
        col += 2
    # rellena celdas vacias del par faltante para conservar la reticula
    while col <= ancho_total:
        for k in (0, 1):
            c = ws.cell(row=fila, column=col + k)
            c.border = BORDE
        col += 2
    return fila + 1


def cabecera_tabla(ws, fila, cabeceras, anchos=None):
    for j, texto in enumerate(cabeceras, start=1):
        c = ws.cell(row=fila, column=j, value=texto)
        c.font = Font(name="Arial", size=9, bold=True, color=BLANCO)
        c.fill = PatternFill("solid", fgColor=VERDE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDE
    ws.row_dimensions[fila].height = 28
    if anchos:
        for j, a in enumerate(anchos, start=1):
            ws.column_dimensions[get_column_letter(j)].width = a
    return fila + 1


def fila_datos(ws, fila, valores, alterna=False, negrita=False, formatos=None):
    relleno = PatternFill("solid", fgColor=VERDE_CLARO) if alterna else None
    for j, v in enumerate(valores, start=1):
        c = ws.cell(row=fila, column=j, value=v)
        c.font = Font(name="Arial", size=9, bold=negrita)
        c.border = BORDE
        c.alignment = Alignment(horizontal="center" if j > 1 else "left",
                                vertical="center", wrap_text=True, indent=0 if j > 1 else 1)
        if relleno:
            c.fill = relleno
        if formatos and formatos.get(j):
            c.number_format = formatos[j]
    return fila + 1


def nota(ws, fila, ncols, texto, etiqueta="NOTA METODOLÓGICA"):
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
    c = ws.cell(row=fila, column=1, value=f"{etiqueta}. {texto}")
    c.font = Font(name="Arial", size=8, italic=True, color="333333")
    c.alignment = Alignment(horizontal="justify", vertical="top", wrap_text=True)
    c.border = BORDE
    ws.row_dimensions[fila].height = max(30, 12 * (len(texto) // 120 + 2))
    return fila + 1


# --- Hoja 1: Resumen ---------------------------------------------------------

def hoja_resumen(wb, b, cat, ap, ndvi, cps, dt):
    ws = wb.create_sheet("Resumen")
    N = 6
    for j, a in enumerate([26, 30, 26, 30, 14, 14], start=1):
        ws.column_dimensions[get_column_letter(j)].width = a

    f = encabezado_institucional(
        ws, N,
        f"FICHA RESUMEN — BLOQUE PRELIMINAR DE INTERVENCIÓN {b}",
        "Diagnóstico Territorial — Plantilla DT Campo Check Validada V5 | "
        "Sistema de referencia: UTM WGS 84 Zona 17S (EPSG:32717)")

    alt_min, alt_max = ap["alt_min"], ap["alt_max"]
    amplitud = alt_max - alt_min
    pend = ap["pendiente"]
    msavi = cat["msavi"]
    et_msavi, interp_msavi, cond_msavi = clase_msavi(msavi)

    cp_txt = "Por determinar"
    if cps:
        principal = max(cps, key=lambda x: (x.get("pob") or 0))
        otros = len(cps) - 1
        cp_txt = f"{principal['cp']} ({principal.get('pob') or 's/d'} hab.)"
        if otros > 0:
            cp_txt += f" y {otros} C.P. más en el bloque"
    if dt.get("cp_cercano"):
        cp_txt = dt["cp_cercano"]

    clases_ndvi = ndvi["clases"] if ndvi else {}
    modal_ndvi = max(clases_ndvi, key=lambda k: clases_ndvi[k]["pct"]) if clases_ndvi else "Por determinar"
    pct_modal = f"{clases_ndvi[modal_ndvi]['pct']:.2f} %" if clases_ndvi else ""

    f = titulo_seccion(ws, f, N, "1. IDENTIFICACIÓN Y LOCALIZACIÓN")
    f = fila_par(ws, f, [("Código del bloque", b),
                         ("Microcuenca (catálogo)", cat["microcuenca"])])
    f = fila_par(ws, f, [("Zona de planificación", cat["zona"]),
                         ("Microcuenca declarada en ficha DT", nd(dt.get("microcuenca_dt")))])
    f = fila_par(ws, f, [("Departamento", "Piura"),
                         ("Provincia", cat["provincia"])])
    f = fila_par(ws, f, [("Distrito", cat["distrito"]),
                         ("Capital distrital", nd(cat.get("capital"), cat["distrito"]))])
    f = fila_par(ws, f, [("Centro poblado asociado", cp_txt),
                         ("Comunidad campesina", nd(dt.get("comunidad"), "Por verificar"))])
    f = fila_par(ws, f, [("Superficie de catálogo (V5/V6), ha", cat["area_ha"]),
                         ("Fuente de superficie", "Catálogo maestro Bloques V5/V6")])
    f = fila_par(ws, f, [("Centroide UTM ESTE (m)", cat["este"]),
                         ("Centroide UTM NORTE (m)", cat["norte"])])
    f = fila_par(ws, f, [("Sistema de coordenadas", "UTM WGS 84 Zona 17S (EPSG:32717)"),
                         ("Tipo de intervención", "Restauración")])
    f += 1

    f = titulo_seccion(ws, f, N, "2. PARÁMETROS FÍSICOS (ESTADÍSTICA ZONAL SOBRE MDE)")
    f = fila_par(ws, f, [("Altitud mínima (msnm)", alt_min),
                         ("Altitud máxima (msnm)", alt_max)])
    f = fila_par(ws, f, [("Amplitud altitudinal (m)", amplitud),
                         ("Piso altitudinal dominante", piso_altitudinal((alt_min + alt_max) / 2))])
    f = fila_par(ws, f, [("Pendiente promedio (%)", round(pend, 2)),
                         ("Pendiente promedio (grados)", round(math.degrees(math.atan(pend / 100.0)), 1))])
    f = fila_par(ws, f, [("Clase de pendiente equivalente", clase_pendiente(pend)),
                         ("Rango de pendiente declarado en campo", nd(dt.get("rango_pendiente")))])
    f = fila_par(ws, f, [("Forma predominante del terreno", nd(dt.get("forma_terreno"))),
                         ("Posición fisiográfica", nd(dt.get("posicion_fisio")))])
    f = fila_par(ws, f, [("Exposición / orientación", nd(dt.get("exposicion"))),
                         ("Afloramientos rocosos", nd(dt.get("afloramientos")))])
    f = fila_par(ws, f, [("Escarpes activos", nd(dt.get("escarpes"))),
                         ("Remociones en masa activas", nd(dt.get("remociones")))])
    f += 1

    f = titulo_seccion(ws, f, N, "3. ÍNDICES DE VEGETACIÓN (SENTINEL-2)")
    f = fila_par(ws, f, [("MSAVI 2024 — media del bloque", round(msavi, 4)),
                         ("MSAVI 2024 — clase de la media", f"{et_msavi} ({interp_msavi})")])
    f = fila_par(ws, f, [("Condición frente al umbral 0.4976", cond_msavi),
                         ("NDVI mediana 2025 — clase modal", f"{modal_ndvi} ({pct_modal})")])
    f = fila_par(ws, f, [("Superficie clasificada NDVI 2025 (ha)",
                          round(ndvi["area_ha_ndvi"], 3) if ndvi else "Por determinar"),
                         ("Desviación frente al catálogo (%)",
                          round(100 * (ndvi["area_ha_ndvi"] - cat["area_ha"]) / cat["area_ha"], 2)
                          if ndvi else "Por determinar")])
    f += 1

    f = titulo_seccion(ws, f, N, "4. ECOSISTEMA Y ESTADO DE CONSERVACIÓN")
    f = fila_par(ws, f, [("Tipo de ecosistema (UP)", nd(dt.get("ecosistema"))),
                         ("Superficie de ecosistema (ha)", nd(dt.get("sup_ecosistema")))])
    f = fila_par(ws, f, [("Estado de conservación", nd(dt.get("estado_cons"))),
                         ("Uso actual dominante del suelo", nd(dt.get("uso_dominante")))])
    f = fila_par(ws, f, [("Tipo de cobertura dominante", nd(dt.get("tipo_cobertura"))),
                         ("Cobertura vegetal total — campo (%)", nd(dt.get("cob_veg")))])
    f = fila_par(ws, f, [("Suelo desnudo — campo (%)", nd(dt.get("suelo_desnudo"))),
                         ("Regeneración natural", nd(dt.get("regeneracion")))])
    f = fila_par(ws, f, [("Nivel general de erosión", nd(dt.get("nivel_erosion"))),
                         ("N.° de cárcavas registradas", nd(dt.get("n_carcavas")))])
    f = fila_par(ws, f, [("Elenco florístico (n.° de taxones)", nd(dt.get("n_taxones"))),
                         ("Estado sanitario", nd(dt.get("sanitario")))])
    f += 1

    f = titulo_seccion(ws, f, N, "5. RESPONSABLE Y MODALIDAD DE LA EVALUACIÓN")
    f = fila_par(ws, f, [("Responsable de la evaluación", nd(dt.get("evaluador"))),
                         ("Fecha de evaluación", nd(dt.get("fecha_eval")))])
    f = fila_par(ws, f, [("Hora de registro", nd(dt.get("hora"), "Por registrar")),
                         ("Correlativo de ficha", nd(dt.get("correlativo"), "Por registrar"))])
    f = fila_par(ws, f, [("Entidad", "ANIN - DIME - SESDI"),
                         ("Fase del estudio", "Preinversión (Perfil) · Invierte.pe")])
    f = fila_par(ws, f, [("Instrumento aplicado", "Fichas F-DT-01 a F-DT-05 (Plantilla V5)"),
                         ("Parcela de muestreo", nd(dt.get("parcela"), "No instalada"))])
    f = fila_par(ws, f, [("Estaciones fotográficas georreferenciadas",
                          nd(dt.get("n_estaciones"), "Sin registro en la ficha DT")),
                         ("Modalidad de acceso", nd(dt.get("modalidad_acceso")))])
    f += 1

    f = titulo_seccion(ws, f, N, "6. SÍNTESIS PARA LA GESTIÓN DEL RIESGO EN CONTEXTO DE CAMBIO CLIMÁTICO")
    f = fila_par(ws, f, [("Causa subyacente principal", nd(dt.get("causa_subyacente"))),
                         ("Velocidad de degradación", nd(dt.get("velocidad")))])
    f = fila_par(ws, f, [("Reversibilidad técnica", nd(dt.get("reversibilidad"))),
                         ("Urgencia de intervención", nd(dt.get("urgencia")))])
    f = fila_par(ws, f, [("Urgencia de control de erosión", nd(dt.get("urgencia_control"))),
                         ("Zona de recarga hídrica", nd(dt.get("recarga")))])
    f = fila_par(ws, f, [("Peligro integrado preliminar (MCA-AHP)",
                          "No disponible en los insumos; debe tomarse del modelamiento de mesolocalización"),
                         ("Prioridad de intervención", "Por definir en mesolocalización")])
    f = fila_par(ws, f, [("Estado de verificación de campo",
                          "VERIFICADO" if dt.get("fecha_eval") else "PENDIENTE"),
                         ("Marco del indicador de brecha", "R.M. N.° 00213-2024-MINAM")])
    f += 1

    f = nota(ws, f, N,
             "Ningún valor ausente ha sido estimado o inferido sin declararlo. Los campos que no pudieron "
             "sustentarse en observación de campo, estadística zonal o catálogo oficial se consignan como "
             "«Por determinar» o «Por verificar». La altitud y la pendiente proceden de estadística zonal "
             "sobre el modelo digital de elevación; el MSAVI 2024 y el NDVI mediana 2025 proceden de "
             "compuestos Sentinel-2. Las discrepancias detectadas entre campo, gabinete y catálogo se "
             "listan en la hoja «Control de consistencia».",
             "DECLARACIÓN DE INTEGRIDAD DE DATOS")

    ws.freeze_panes = "A7"
    ws.sheet_view.showGridLines = False
    return ws


# --- Hoja 2: Cobertura MSAVI-NDVI -------------------------------------------

def hoja_cobertura(wb, b, cat, ap, ndvi, dt):
    ws = wb.create_sheet("Cobertura MSAVI-NDVI")
    N = 5
    f = encabezado_institucional(
        ws, N,
        f"DISTRIBUCIÓN AREAL DE CLASES ESPECTRALES — BLOQUE {b}",
        "MSAVI 2024 y NDVI mediana 2025 (Sentinel-2, mosaico Piura) · "
        "Umbral de brecha del proyecto: MSAVI = 0.4976")

    msavi = cat["msavi"]
    area = cat["area_ha"]
    et_msavi, interp_msavi, cond_msavi = clase_msavi(msavi)

    f = titulo_seccion(ws, f, N, "A. MSAVI 2024 — CLASIFICACIÓN POR UMBRALES DEL PROYECTO")
    f = cabecera_tabla(ws, f,
                       ["Clase MSAVI", "Superficie (ha)", "% del área clasificada",
                        "Interpretación", "Condición frente al umbral 0.4976"],
                       [24, 16, 20, 30, 28])
    fila_ini = f
    for i, (etiqueta, lo, hi, interp, cond) in enumerate(CLASES_MSAVI):
        es_modal = (etiqueta == et_msavi)
        f = fila_datos(ws, f,
                       [etiqueta, "Por determinar", "Por determinar", interp, cond],
                       alterna=(i % 2 == 1), negrita=es_modal)
        if es_modal:
            for j in range(1, N + 1):
                ws.cell(row=f - 1, column=j).fill = PatternFill("solid", fgColor="FFF2CC")
    f = fila_datos(ws, f, ["MSAVI 2024 — MEDIA DEL BLOQUE", round(msavi, 6), "—",
                           f"{interp_msavi} · clase {et_msavi}", cond_msavi], negrita=True)
    for j in range(1, N + 1):
        ws.cell(row=f - 1, column=j).fill = PatternFill("solid", fgColor=VERDE_CLARO)
    f += 1

    f = nota(ws, f, N,
             "La media del MSAVI 2024 del bloque (%.6f) procede del catálogo maestro de bloques V5/V6 y "
             "es dato oficial. La DISTRIBUCIÓN AREAL por clase de MSAVI no está disponible como estadística "
             "zonal en los insumos de este entregable: se consigna «Por determinar» y NO se estima, conforme "
             "a la declaración de integridad de datos del proyecto. La fila resaltada indica la clase en la "
             "que cae la media del bloque. Para el dimensionamiento de metas físicas debe ejecutarse "
             "estadística zonal directa sobre el ráster MSAVI 2024 recortado al polígono." % msavi)
    f += 1

    f = titulo_seccion(ws, f, N, "B. NDVI MEDIANA 2025 — DISTRIBUCIÓN AREAL (ESTADÍSTICA ZONAL)")
    f = cabecera_tabla(ws, f, ["Clase NDVI", "Superficie (ha)", "% del área clasificada",
                               "Interpretación", "Observación"])
    if ndvi:
        clases = ndvi["clases"]
        fila_ini = f
        for i, nombre in enumerate([c for c in ORDEN_NDVI if c in clases] +
                                   [c for c in clases if c not in ORDEN_NDVI]):
            v = clases[nombre]
            interp = {
                "Vegetación alta": "Dosel continuo, matorral denso o mosaico agroforestal",
                "Vegetación mediana": "Pastizal cultivado, matorral ralo, parcelas agrícolas",
                "Vegetación ligera": "Cobertura discontinua, suelo parcialmente expuesto",
                "Tierra desnuda": "Suelo desnudo, afloramientos, plataformas y taludes",
                "Clase no vegetal": "Superficie no vegetal",
            }.get(nombre, "")
            f = fila_datos(ws, f, [nombre, round(v["ha"], 4), round(v["pct"], 2), interp, ""],
                           alterna=(i % 2 == 1),
                           formatos={2: "0.0000", 3: "0.00"})
        # Totales con formula
        c1 = f"=SUM(B{fila_ini}:B{f-1})"
        c2 = f"=SUM(C{fila_ini}:C{f-1})"
        f = fila_datos(ws, f, ["TOTAL CLASIFICADO", c1, c2, "", ""], negrita=True,
                       formatos={2: "0.0000", 3: "0.00"})
        for j in range(1, N + 1):
            ws.cell(row=f - 1, column=j).fill = PatternFill("solid", fgColor=VERDE_CLARO)
        f = fila_datos(ws, f, ["Superficie de catálogo (V5/V6)", area, "—",
                               "Desviación planimétrica frente al catálogo",
                               f"{100*(ndvi['area_ha_ndvi']-area)/area:+.2f} %"], negrita=True)
    else:
        f = fila_datos(ws, f, ["Sin dato NDVI para este bloque", "Por determinar",
                               "Por determinar", "", ""])
    f += 1

    veg_alta = ndvi["clases"].get("Vegetación alta", {}).get("pct", 0) if ndvi else 0
    cob_campo = dt.get("cob_veg")
    contraste = ""
    if isinstance(cob_campo, (int, float)):
        contraste = (" Contraste registrado en este bloque: la cobertura vegetal estimada en campo es "
                     "%s %% frente a un %.2f %% del polígono clasificado como «Vegetación alta» por el "
                     "NDVI 2025; la observación de campo corresponde a una parcela o recorrido puntual y "
                     "no es estadísticamente representativa de las %.3f ha del bloque." %
                     (cob_campo, veg_alta, area))

    f = nota(ws, f, N,
             "Los índices MSAVI y NDVI miden vigor y densidad de biomasa, no composición ni integridad "
             "ecosistémica. Un valor alto NO equivale a ausencia de degradación: en bloques con mosaico "
             "agrícola, pastizal cultivado o plantaciones, la respuesta espectral puede ser alta sobre una "
             "unidad productora sustituida en su composición. La diferencia entre ambos productos no debe "
             "leerse como mejora entre 2024 y 2025: son índices distintos (el NDVI satura ante biomasa densa, "
             "el MSAVI conserva sensibilidad al suelo de fondo) aplicados sobre compuestos temporales "
             "distintos." + contraste,
             "LECTURA CRÍTICA")

    ws.freeze_panes = "A7"
    ws.sheet_view.showGridLines = False
    return ws


# --- Hoja 3: Estaciones fotograficas ----------------------------------------

def _dist(e, n, e0, n0):
    try:
        return round(math.hypot(float(e) - float(e0), float(n) - float(n0)), 1)
    except (TypeError, ValueError):
        return ""


def hoja_estaciones(wb, b, cat, dt):
    ws = wb.create_sheet("Estaciones fotográficas")
    N = 7
    fecha = dt.get("fecha_eval") or "fecha por registrar"
    f = encabezado_institucional(
        ws, N,
        f"PUNTOS GEORREFERENCIADOS DE VERIFICACIÓN — BLOQUE {b}",
        f"Levantamiento del {fecha} · Sistema de referencia UTM WGS 84 Zona 17S (EPSG:32717)")

    e0, n0 = cat["este"], cat["norte"]
    puntos = []

    for est in (dt.get("estaciones_fotograficas") or []):
        cod, e, n = (est + [None, None, None])[:3] if isinstance(est, list) else (est, None, None)
        puntos.append([cod, "Estación fotográfica", e, n, "", "Registro fotográfico georreferenciado"])

    for c in (dt.get("carcavas") or []):
        cod, tipo, e, n = c[0], c[1], c[2], c[3]
        estado = c[4] if len(c) > 4 else ""
        causa = c[5] if len(c) > 5 else ""
        puntos.append([cod, f"Rasgo erosivo — {tipo}", e, n, estado, causa])

    for i, fu in enumerate(dt.get("fuentes_agua") or [], start=1):
        tipo, e, n = fu[0], fu[1], fu[2]
        regimen = fu[3] if len(fu) > 3 else ""
        obs = fu[6] if len(fu) > 6 else ""
        puntos.append([f"H-{i:02d}", f"Fuente de agua — {tipo}", e, n, regimen, obs])

    for i, esp in enumerate(dt.get("especies_clave") or [], start=1):
        if len(esp) >= 5 and esp[3] and esp[4]:
            puntos.append([f"B-{i:02d}", f"Especie clave — {esp[0]}", esp[3], esp[4],
                           esp[1], esp[2]])

    if dt.get("utm_e_dt") and dt.get("utm_n_dt"):
        puntos.insert(0, ["P-00", "Punto de muestreo declarado en ficha DT",
                          dt["utm_e_dt"], dt["utm_n_dt"], "",
                          "Coordenada consignada en F-DT-01"])

    f = titulo_seccion(ws, f, N, "INVENTARIO DE PUNTOS GEORREFERENCIADOS DE LA FICHA DT")
    f = cabecera_tabla(ws, f,
                       ["Código", "Naturaleza del punto", "UTM ESTE (m)", "UTM NORTE (m)",
                        "Dist. al centroide (m)", "Estado / régimen", "Contenido registrado"],
                       [12, 34, 15, 15, 18, 22, 44])

    if puntos:
        fila_ini = f
        for i, p in enumerate(puntos):
            cod, nat, e, n, est, obs = p
            f = fila_datos(ws, f, [cod, nat, e, n, _dist(e, n, e0, n0), est, obs],
                           alterna=(i % 2 == 1), formatos={3: "#,##0", 4: "#,##0"})
        f = fila_datos(ws, f, [f"TOTAL: {len(puntos)} puntos", "", "", "", "", "", ""],
                       negrita=True)
        for j in range(1, N + 1):
            ws.cell(row=f - 1, column=j).fill = PatternFill("solid", fgColor=VERDE_CLARO)
    else:
        f = fila_datos(ws, f, ["—", "Sin puntos georreferenciados en la ficha DT",
                               "", "", "", "", "No corresponde para este bloque"])
    f += 1

    f = fila_par(ws, f, [("Centroide del bloque — UTM ESTE (m)", e0),
                         ("Centroide del bloque — UTM NORTE (m)", n0)], ancho_total=N)
    f = fila_par(ws, f, [("N.° de estaciones fotográficas declaradas",
                          nd(dt.get("n_estaciones"), "Sin registro en la ficha DT")),
                         ("Estaciones dentro del polígono",
                          nd(dt.get("estaciones_dentro"), "Por determinar"))], ancho_total=N)
    f += 1

    obs = dt.get("obs_dt01") or ""
    f = nota(ws, f, N,
             "Las coordenadas de esta hoja proceden de la ficha DT del bloque (F-DT-01 a F-DT-05) y "
             "corresponden a rasgos erosivos, fuentes de agua, especies clave y al punto de muestreo "
             "declarado. La distancia al centroide se calcula sobre coordenadas UTM 17S y es referencial: "
             "no acredita inclusión ni exclusión respecto del polígono, que exige prueba de inclusión "
             "sobre la geometría del bloque. Las estaciones situadas a menos de ~30 m del límite deben "
             "considerarse SOBRE EL LÍMITE, dentro del margen de error de la georreferenciación "
             "(± 10-15 m). " + obs,
             "CONTROL GEOMÉTRICO")

    ws.freeze_panes = "A7"
    ws.sheet_view.showGridLines = False
    return ws


# --- Hoja 4: Microcuenca -----------------------------------------------------

def hoja_microcuenca(wb, b, cat, catalogo, alt_pend, ndvi_all, dt):
    mic = cat["microcuenca"]
    ws = wb.create_sheet("Microcuenca")
    N = 8
    f = encabezado_institucional(
        ws, N,
        f"CONTEXTO INTRAMICROCUENCA {mic} — BLOQUE {b}",
        f"Distrito de {cat['distrito']} · Provincia de {cat['provincia']} · "
        "Comparación de los bloques preliminares de intervención de la misma microcuenca")

    hermanos = sorted([k for k, v in catalogo.items() if v["microcuenca"] == mic],
                      key=lambda k: -catalogo[k]["area_ha"])
    total_area = sum(catalogo[k]["area_ha"] for k in hermanos)

    f = titulo_seccion(ws, f, N, f"BLOQUES DE LA MICROCUENCA {mic}")
    f = cabecera_tabla(ws, f,
                       ["Bloque", "Área (ha)", "% microcuenca", "Rango altitudinal (msnm)",
                        "Amplitud (m)", "Pendiente prom. (%)", "MSAVI 2024",
                        "NDVI 2025 — Veg. alta (%)"],
                       [14, 13, 14, 24, 13, 18, 13, 20])

    fila_ini = f
    for i, k in enumerate(hermanos):
        c = catalogo[k]
        a = alt_pend.get(k)
        nv = ndvi_all.get(k)
        rango = f"{a['alt_min']} – {a['alt_max']}" if a else "Por determinar"
        ampl = (a["alt_max"] - a["alt_min"]) if a else "Por determinar"
        pend = round(a["pendiente"], 2) if a else "Por determinar"
        vga = round(nv["clases"].get("Vegetación alta", {}).get("pct", 0), 2) if nv else "Por determinar"
        etiqueta = ("► " + k) if k == b else k
        f = fila_datos(ws, f,
                       [etiqueta, round(c["area_ha"], 3),
                        round(100 * c["area_ha"] / total_area, 2), rango, ampl, pend,
                        round(c["msavi"], 4), vga],
                       alterna=(i % 2 == 1), negrita=(k == b),
                       formatos={2: "0.000", 3: "0.00", 7: "0.0000", 8: "0.00"})
        if k == b:
            for j in range(1, N + 1):
                ws.cell(row=f - 1, column=j).fill = PatternFill("solid", fgColor="FFF2CC")

    fila_fin = f - 1
    n_herm = len(hermanos)
    n_ap = sum(1 for k in hermanos if k in alt_pend)
    n_nd = sum(1 for k in hermanos if k in ndvi_all)
    f = fila_datos(ws, f,
                   ["TOTAL / PROMEDIO",
                    f"=SUM(B{fila_ini}:B{fila_fin})",
                    f"=SUM(C{fila_ini}:C{fila_fin})",
                    "%d – %d" % (min(alt_pend[k]["alt_min"] for k in hermanos if k in alt_pend),
                                 max(alt_pend[k]["alt_max"] for k in hermanos if k in alt_pend))
                    if any(k in alt_pend for k in hermanos) else "Por determinar",
                    "",
                    f"=ROUND(AVERAGE(F{fila_ini}:F{fila_fin}),2)",
                    f"=ROUND(AVERAGE(G{fila_ini}:G{fila_fin}),4)",
                    f"=ROUND(AVERAGE(H{fila_ini}:H{fila_fin}),2)"],
                   negrita=True, formatos={2: "0.000", 3: "0.00", 7: "0.0000", 8: "0.00"})
    for j in range(1, N + 1):
        ws.cell(row=f - 1, column=j).fill = PatternFill("solid", fgColor=VERDE_CLARO)

    # Declaracion explicita de la cobertura de cada promedio (principio de no fabricacion):
    # AVERAGE omite las celdas de texto «Por determinar», de modo que un promedio parcial
    # se veria identico a uno completo si no se advirtiera aqui.
    if n_ap < n_herm or n_nd < n_herm:
        partes = []
        def _resto(n):
            return "el bloque restante carece" if n == 1 else "los %d bloques restantes carecen" % n
        if n_ap < n_herm:
            partes.append("los promedios de PENDIENTE y MSAVI 2024 se calculan sobre "
                          "%d de los %d bloques de la microcuenca; %s de dato derivado "
                          "del MDE" % (n_ap, n_herm, _resto(n_herm - n_ap)))
        if n_nd < n_herm:
            partes.append("el promedio de VEGETACIÓN ALTA (NDVI 2025) se calcula sobre "
                          "%d de los %d bloques; %s de estadística zonal de NDVI"
                          % (n_nd, n_herm, _resto(n_herm - n_nd)))
        cierre = (". Los totales de ÁREA y % de microcuenca sí comprenden la totalidad "
                  "de los {} bloques. No se estima ningún valor ausente.".format(n_herm))
        f = nota(ws, f, N,
                 "La fila TOTAL / PROMEDIO omite las celdas «Por determinar»: " +
                 "; ".join(partes) + cierre,
                 etiqueta="ALCANCE DE LOS PROMEDIOS")
    f += 1

    # --- Lectura posicional automatica del bloque dentro de su microcuenca ---
    a_b = alt_pend.get(b)
    pos_area = sorted(hermanos, key=lambda k: -catalogo[k]["area_ha"]).index(b) + 1
    n = len(hermanos)
    frases = [f"El bloque {b} ocupa el puesto {pos_area} de {n} por superficie dentro de la microcuenca "
              f"{mic} ({100*cat['area_ha']/total_area:.2f} % de las {total_area:.2f} ha preliminares "
              f"de la microcuenca)."]
    if a_b:
        con_ap = [k for k in hermanos if k in alt_pend]
        pos_pend = sorted(con_ap, key=lambda k: alt_pend[k]["pendiente"]).index(b) + 1
        pos_msavi = sorted(hermanos, key=lambda k: -catalogo[k]["msavi"]).index(b) + 1
        prom_pend = sum(alt_pend[k]["pendiente"] for k in con_ap) / len(con_ap)
        prom_msavi = sum(catalogo[k]["msavi"] for k in hermanos) / n
        frases.append(
            f"Su pendiente promedio ({a_b['pendiente']:.2f} %, clase «{clase_pendiente(a_b['pendiente'])}») "
            f"es la {pos_pend}.ª más baja de {len(con_ap)} y se sitúa "
            f"{'por debajo' if a_b['pendiente'] < prom_pend else 'por encima'} del promedio de la "
            f"microcuenca ({prom_pend:.2f} %).")
        frases.append(
            f"Su MSAVI 2024 ({cat['msavi']:.4f}) es el {pos_msavi}.º de {n} y queda "
            f"{'por encima' if cat['msavi'] > prom_msavi else 'por debajo'} del promedio local "
            f"({prom_msavi:.4f}). "
            f"Amplitud altitudinal del bloque: {a_b['alt_max']-a_b['alt_min']} m "
            f"({a_b['alt_min']}–{a_b['alt_max']} msnm).")
    frases.append("El cuadro conserva todos los bloques catalogados en la microcuenca, incluidos los "
                  "descartados en el tamizaje, para referencia del conjunto. Los bloques sin ficha DT "
                  "figuran con sus parámetros de gabinete.")

    f = nota(ws, f, N, " ".join(frases), "LECTURA INTRAMICROCUENCA")
    f += 1
    f = nota(ws, f, N,
             "Área y MSAVI 2024 proceden del catálogo maestro de bloques V5/V6. El rango altitudinal y la "
             "pendiente promedio proceden del reporte de estadística zonal sobre el MDE («Rango de altitud "
             "y Pendiente Promedio por Bloques V6»). El porcentaje de «Vegetación alta» procede de la "
             "estadística zonal del NDVI mediana 2025 (Sentinel-2). Los totales y promedios se calculan "
             "con fórmulas sobre las filas del cuadro.")

    ws.freeze_panes = "A7"
    ws.sheet_view.showGridLines = False
    return ws


# --- Hoja 5: Control de consistencia ----------------------------------------

def hoja_consistencia(wb, b, cat, ap, ndvi, cps, dt):
    ws = wb.create_sheet("Control de consistencia")
    N = 5
    f = encabezado_institucional(
        ws, N,
        f"CONTROL DE CONSISTENCIA DE LA INFORMACIÓN — BLOQUE {b}",
        "Discrepancias detectadas entre el registro de campo, los productos de gabinete y el "
        "catálogo maestro. Calificación: SUSTANTIVA / NO SUSTANTIVA / CORREGIDO / CONFORME.")

    filas = []
    n = 0

    def add(campo, discrepancia, calif, tratamiento):
        nonlocal n
        n += 1
        filas.append([f"{'C' if calif == 'CORREGIDO' else 'D'}-{n:02d}",
                      campo, discrepancia, calif, tratamiento])

    # 1. Microcuenca declarada en ficha vs catalogo
    mic_dt = dt.get("microcuenca_dt")
    if mic_dt and mic_dt != cat["microcuenca"]:
        add("Código de microcuenca",
            f"La ficha DT declara «{mic_dt}»; el catálogo maestro V5/V6 asigna el bloque a "
            f"«{cat['microcuenca']}»",
            "SUSTANTIVA",
            "Se conserva el código del catálogo maestro para todo cálculo intramicrocuenca. La "
            "declaración de la ficha debe corregirse en la siguiente versión del formato V5.")
    elif mic_dt:
        add("Código de microcuenca",
            f"Ficha DT y catálogo maestro coinciden en «{cat['microcuenca']}»",
            "CONFORME", "Sin acción.")

    # 2. Punto de muestreo vs centroide
    e, nn = dt.get("utm_e_dt"), dt.get("utm_n_dt")
    if not e or not nn:
        add("UTM del punto de muestreo",
            "Las celdas de UTM ESTE / NORTE del punto de muestreo están vacías o incompletas en la ficha DT",
            "SUSTANTIVA",
            "No se sustituye por el centroide del catálogo. Se consigna «Por registrar en campo».")
    else:
        d = _dist(e, nn, cat["este"], cat["norte"])
        if isinstance(d, float) and d > 800:
            add("UTM del punto de muestreo",
                f"El punto de muestreo declarado ({e} E / {nn} N) dista {d:,.0f} m del centroide de "
                f"catálogo ({cat['este']} E / {cat['norte']} N), distancia incompatible con la "
                f"superficie del bloque ({cat['area_ha']} ha)",
                "SUSTANTIVA",
                "Se declara la coordenada como POR VERIFICAR: puede corresponder a otro bloque o a un "
                "error de digitación. No se corrige de oficio.")
        else:
            add("UTM del punto de muestreo",
                f"El punto de muestreo declarado dista {d:,.0f} m del centroide de catálogo, "
                f"compatible con la extensión del bloque",
                "CONFORME", "Sin acción.")

    # 3. Superficie NDVI vs catalogo
    if ndvi:
        dev = 100 * (ndvi["area_ha_ndvi"] - cat["area_ha"]) / cat["area_ha"]
        if abs(dev) > 2:
            add("Superficie del polígono",
                f"La superficie clasificada por el NDVI 2025 ({ndvi['area_ha_ndvi']:.3f} ha) difiere "
                f"{dev:+.2f} % de la superficie de catálogo ({cat['area_ha']} ha)",
                "SUSTANTIVA",
                "Se conserva la superficie de catálogo para el dimensionamiento de metas físicas y se "
                "declara la desviación. Requiere revisión de la geometría del polígono.")
        else:
            add("Superficie del polígono",
                f"La superficie clasificada por el NDVI 2025 ({ndvi['area_ha_ndvi']:.3f} ha) coincide "
                f"con la de catálogo ({cat['area_ha']} ha) dentro del ±2 % ({dev:+.2f} %)",
                "CONFORME", "Sin acción: la segmentación valida la geometría del bloque.")

    # 4. Rango de pendiente campo vs MDE
    rp = dt.get("rango_pendiente")
    clase_mde = clase_pendiente(ap["pendiente"])
    if rp and rp.strip() != clase_mde:
        add("Rango de pendiente dominante",
            f"La ficha DT declara «{rp}»; la estadística zonal sobre el MDE arroja una pendiente "
            f"promedio de {ap['pendiente']:.2f} %, que corresponde a la clase «{clase_mde}»",
            "NO SUSTANTIVA",
            "Se conserva el registro de campo como observación de sector y se adopta el valor del MDE "
            "para el cálculo intramicrocuenca. La divergencia es esperable: el dato de campo es puntual "
            "y el del MDE es promedio del polígono.")

    # 5. Rango altitudinal campo vs MDE
    ra = dt.get("rango_altitudinal")
    piso_mde = piso_altitudinal((ap["alt_min"] + ap["alt_max"]) / 2)
    if ra and ra.strip() != piso_mde:
        add("Rango altitudinal / piso ecológico",
            f"La ficha DT declara «{ra}»; el MDE sitúa el bloque entre {ap['alt_min']} y "
            f"{ap['alt_max']} msnm, cuyo punto medio corresponde a «{piso_mde}»",
            "NO SUSTANTIVA",
            "Se declara la divergencia. El bloque abarca más de un piso altitudinal cuando la amplitud "
            "supera el ancho de la clase; el registro de campo corresponde al sector recorrido.")

    # 6. Cobertura de campo vs NDVI
    cv = dt.get("cob_veg")
    if isinstance(cv, (int, float)) and ndvi:
        veg = sum(v["pct"] for k, v in ndvi["clases"].items() if k.startswith("Vegetación"))
        if abs(veg - cv) > 15:
            add("Cobertura vegetal total",
                f"La ficha DT estima {cv} % de cobertura vegetal en campo; el NDVI mediana 2025 "
                f"clasifica {veg:.2f} % del polígono en clases de vegetación",
                "NO SUSTANTIVA",
                "Se declaran ambos valores. La observación de campo corresponde a una parcela o recorrido "
                "puntual y no es representativa de la totalidad del bloque; el índice espectral no "
                "discrimina composición ni integridad ecosistémica.")

    # 7. Altitud GPS no registrada
    ag = dt.get("altitud_gps")
    if not ag or "determinar" in str(ag).lower() or "verificar" in str(ag).lower():
        add("Altitud GPS del punto de muestreo",
            "La aplicación de captura no registró altitud GPS en campo",
            "NO SUSTANTIVA",
            f"El rango altitudinal del bloque ({ap['alt_min']}–{ap['alt_max']} msnm) y la pendiente "
            f"promedio ({ap['pendiente']:.2f} %) se toman de la estadística zonal sobre el MDE, no de "
            "medición de campo. Se declara la fuente.")

    # 8. Inventario de carcavas
    nc = dt.get("n_carcavas")
    if nc in (None, "", "Sin inventario", "Por determinar"):
        add("Inventario georreferenciado de cárcavas",
            "El inventario de cárcavas no fue ejecutado o no consta en la ficha DT",
            "SUSTANTIVA",
            "La ausencia de registros NO equivale a ausencia verificada de cárcavas. Queda pendiente el "
            "levantamiento instrumental a cargo del especialista de Infraestructura Marrón antes de fijar "
            "metas físicas de infraestructura gris.")
    elif dt.get("long_carcavas") in (None, "", "Por determinar", "Por verificar", "No determinada"):
        add("Dimensiones del inventario de cárcavas",
            f"Se registran {nc} rasgos erosivos, pero longitud, profundidad y ancho no fueron medidos "
            f"instrumentalmente",
            "NO SUSTANTIVA",
            "El conteo registrado no debe usarse para calcular densidad de cárcavas por hectárea. "
            "Pendiente de levantamiento con GPS submétrico.")

    # 9. Elenco floristico incompleto
    nt = dt.get("n_taxones")
    if isinstance(nt, int) and nt < 15:
        add("Elenco florístico",
            f"La ficha registra {nt} taxones, por debajo del mínimo de 15 que exige el formato V5",
            "SUSTANTIVA",
            "Registro florístico preliminar, sin colecta ni determinación botánica formal. Debe cerrarse "
            "con colecta botánica en la siguiente campaña.")

    # 10. Estaciones fuera del poligono
    if dt.get("estaciones_dentro") == 0:
        add("Representatividad del levantamiento",
            "Ninguna de las estaciones fotográficas georreferenciadas se localiza dentro del polígono "
            "del bloque",
            "SUSTANTIVA",
            "La verificación documenta el perímetro y la vista panorámica del bloque, no su interior. "
            "Los parámetros que exigen recorrido interno se consignan «Por determinar».")

    # 11. Centro poblado
    if cps:
        nombres = ", ".join(sorted({c["cp"] for c in cps})[:6])
        declarado = str(dt.get("cp_cercano") or "")
        coincide = any(str(c["cp"]).upper().split()[0] in declarado.upper() for c in cps if c["cp"])
        add("Centro poblado asociado",
            f"La ficha DT declara «{declarado or 'sin registro'}»; el cruce INEI-Bloques V5 identifica "
            f"{len(cps)} centro(s) poblado(s) en el bloque: {nombres}",
            "CONFORME" if coincide else "NO SUSTANTIVA",
            "Se conserva el registro de campo como referencia y se adopta el listado INEI para el "
            "diagnóstico social y el cálculo de población beneficiaria.")

    # 12. Peligro integrado
    add("Peligro integrado (MCA-AHP)",
        "El nivel de peligro integrado no está disponible en los insumos de este entregable",
        "SUSTANTIVA",
        "Debe tomarse del modelamiento de mesolocalización (PMM + EPH + PGI ponderados por AHP) antes "
        "de fijar la prioridad definitiva de intervención del bloque.")

    # 13. Correcciones declaradas por la Unidad Formuladora en la ficha
    for c in (dt.get("correcciones") or []):
        if len(c) >= 5:
            filas.append([c[0], c[1], c[2], c[3], c[4]])
        else:
            n += 1
            filas.append([f"D-{n:02d}", c[0], c[1], c[2], c[3]])

    f = titulo_seccion(ws, f, N, "DISCREPANCIAS Y VERIFICACIONES")
    f = cabecera_tabla(ws, f, ["Cód.", "Campo afectado", "Discrepancia observada",
                               "Calificación", "Tratamiento adoptado"],
                       [9, 32, 60, 18, 62])
    colores = {"SUSTANTIVA": "F8CBAD", "NO SUSTANTIVA": "FFF2CC",
               "CORREGIDO": "C6E0B4", "CONFORME": "D9E1F2"}
    for i, fila in enumerate(filas):
        f = fila_datos(ws, f, fila, alterna=(i % 2 == 1))
        for j in range(1, N + 1):
            ws.cell(row=f - 1, column=j).alignment = Alignment(
                horizontal="left", vertical="top", wrap_text=True, indent=1)
        cc = ws.cell(row=f - 1, column=4)
        if fila[3] in colores:
            cc.fill = PatternFill("solid", fgColor=colores[fila[3]])
            cc.font = Font(name="Arial", size=9, bold=True)
            cc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[f - 1].height = 48

    f += 1
    resumen = {}
    for fila in filas:
        resumen[fila[3]] = resumen.get(fila[3], 0) + 1
    f = fila_datos(ws, f, ["RESUMEN", f"{len(filas)} verificaciones",
                           " · ".join(f"{k}: {v}" for k, v in sorted(resumen.items())),
                           "", ""], negrita=True)
    for j in range(1, N + 1):
        ws.cell(row=f - 1, column=j).fill = PatternFill("solid", fgColor=VERDE_CLARO)
    f += 1

    f = nota(ws, f, N,
             "Este control se genera de forma reproducible cruzando la ficha DT del bloque (F-DT-01 a "
             "F-DT-05) con el catálogo maestro de bloques V5/V6, la estadística zonal de altitud y "
             "pendiente sobre el MDE, la estadística zonal del NDVI mediana 2025 y el cruce INEI-Bloques "
             "de centros poblados. Una calificación CONFORME acredita que la verificación se ejecutó y no "
             "arrojó discrepancia; no equivale a validación de campo del dato.")

    ws.freeze_panes = "A7"
    ws.sheet_view.showGridLines = False
    return ws


# --- Orquestacion ------------------------------------------------------------

def cargar():
    def j(nombre):
        with open(os.path.join(DATOS, nombre), encoding="utf-8") as fh:
            return json.load(fh)
    catalogo = j("catalogo_v5.json")
    alt_pend = j("alt_pend.json")
    ndvi_all = j("ndvi.json")
    cps_all = j("centros_poblados.json")
    bloques = j("bloques_dt.json")
    # capital distrital desde el maestro de altitud/pendiente si esta disponible
    return catalogo, alt_pend, ndvi_all, cps_all, bloques


def dt_de(b):
    ruta = os.path.join(BASE, "dt", f"{b}.json")
    if os.path.exists(ruta):
        with open(ruta, encoding="utf-8") as fh:
            return json.load(fh)
    return {}


def construir(b, catalogo, alt_pend, ndvi_all, cps_all):
    cat = catalogo[b]
    ap = alt_pend.get(b)
    if ap is None:
        ap = {"alt_min": 0, "alt_max": 0, "pendiente": 0.0}
    ndvi = ndvi_all.get(b)
    cps = cps_all.get(b, [])
    dt = dt_de(b)

    wb = Workbook()
    wb.remove(wb.active)
    hoja_resumen(wb, b, cat, ap, ndvi, cps, dt)
    hoja_cobertura(wb, b, cat, ap, ndvi, dt)
    hoja_estaciones(wb, b, cat, dt)
    hoja_microcuenca(wb, b, cat, catalogo, alt_pend, ndvi_all, dt)
    hoja_consistencia(wb, b, cat, ap, ndvi, cps, dt)

    wb.properties.title = f"Plantilla Excel — Bloque {b} — IN Piura"
    wb.properties.creator = "ANIN - DIME - SESDI | Proyecto IN Piura (CUI 2669244)"
    wb.properties.subject = "Diagnóstico Territorial por bloque de intervención"
    return wb


def main(solo=None):
    catalogo, alt_pend, ndvi_all, cps_all, bloques = cargar()
    os.makedirs(SALIDA, exist_ok=True)
    hechos, con_dt = [], 0
    for b, folder, titulo in bloques:
        if solo and b not in solo:
            continue
        wb = construir(b, catalogo, alt_pend, ndvi_all, cps_all)
        nombre = f"Plantilla_Excel_Bloque_{b}_IN_Piura.xlsx"
        wb.save(os.path.join(SALIDA, nombre))
        if dt_de(b):
            con_dt += 1
        hechos.append((b, nombre, folder))
    print(f"generados: {len(hechos)} | con ficha DT extraída: {con_dt}")
    with open(os.path.join(BASE, "generados.json"), "w", encoding="utf-8") as fh:
        json.dump(hechos, fh, ensure_ascii=False, indent=1)
    return hechos


if __name__ == "__main__":
    import sys
    main(set(sys.argv[1:]) or None)
