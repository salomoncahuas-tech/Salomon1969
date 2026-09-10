# -*- coding: utf-8 -*-
"""
Proyecto IN Piura — CUI 2669244 · ANIN / DIME / SESDI
Capa de consolidacion de datos del Diagnostico Territorial de los 117 bloques.

Integra, sin imputar ni estimar ningun valor:
  1. Las 117 plantillas Excel de resumen por bloque (V6 con MSAVI), hojas
     Resumen, Cobertura MSAVI-NDVI, Microcuenca, Estaciones fotograficas y
     Control de consistencia.
  2. El Indice de Susceptibilidad Litologica (ISL-MM / ISL-EH) del E3 Estudio
     de Geologia, volumenes I (Morropon), II (Huancabamba) y III (Ayabaca).
  3. El inventario de carcavas codificadas (Caracterizacion de carcavas).

Principio de no invencion: todo campo no sustentado en fuente primaria se
conserva con la marca textual de origen ("Por determinar" / "Por verificar").
"""
import csv
import json
import math
import os
import re
import unicodedata
from collections import Counter, OrderedDict, defaultdict

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "datos_fuente")

PROVINCIAS = ("Morropon", "Huancabamba", "Ayabaca")

PROV_META = {
    "Morropon": {
        "nombre": "Morropón",
        "volumen_geologia": "E3 Estudio de Geología, Volumen I — Morropón (R03)",
        "archivo": "Morropon",
    },
    "Huancabamba": {
        "nombre": "Huancabamba",
        "volumen_geologia": "E3 Estudio de Geología, Volumen II — Huancabamba (R03)",
        "archivo": "Huancabamba",
    },
    "Ayabaca": {
        "nombre": "Ayabaca",
        "volumen_geologia": "E3 Estudio de Geología, Volumen III — Ayabaca (R03)",
        "archivo": "Ayabaca",
    },
}

# Marcas textuales que la plantilla emplea para declarar ausencia de dato.
SIN_DATO = ("Por determinar", "Por verificar", "s/d", "No disponible")


def es_sin_dato(v):
    return v is None or (isinstance(v, str) and v.strip() in SIN_DATO)


def num(v):
    """Devuelve float si el valor es numerico; None si es marca de ausencia."""
    if es_sin_dato(v):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def norm(s):
    """Normaliza texto para comparacion (sin tildes, mayusculas)."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.upper().strip()


def _load_json(nombre):
    with open(os.path.join(DATA, nombre), encoding="utf-8") as fh:
        return json.load(fh)


def cargar_bloques():
    """Devuelve la lista de los 117 bloques con todas las fuentes integradas."""
    master = _load_json("master_117.json")
    isl = _load_json("isl_117.json")

    carcavas = defaultdict(list)
    with open(os.path.join(DATA, "carcavas.csv"), encoding="utf-8") as fh:
        for fila in csv.DictReader(fh):
            carcavas[fila["Bloque"].strip()].append(fila)

    bloques = []
    for reg in master:
        cod = str(reg["Código del bloque"]).strip()
        b = OrderedDict()
        b["codigo"] = cod
        b["provincia"] = reg["Provincia"]
        b["distrito"] = reg["Distrito"]
        b["capital"] = reg["Capital distrital"]
        b["microcuenca"] = reg["Microcuenca (catálogo)"]
        b["microcuenca_ficha"] = reg["Microcuenca declarada en ficha DT"]
        b["zona_planificacion"] = reg["Zona de planificación"]
        b["centro_poblado"] = reg["Centro poblado asociado"]
        b["comunidad"] = reg["Comunidad campesina"]
        b["area_ha"] = num(reg["Superficie de catálogo (V5/V6), ha"])
        b["este"] = num(reg["Centroide UTM ESTE (m)"])
        b["norte"] = num(reg["Centroide UTM NORTE (m)"])
        b["tipo_intervencion"] = reg["Tipo de intervención"]

        # 2. Parametros fisicos (estadistica zonal sobre MDE)
        b["alt_min"] = num(reg["Altitud mínima (msnm)"])
        b["alt_max"] = num(reg["Altitud máxima (msnm)"])
        b["amplitud"] = num(reg["Amplitud altitudinal (m)"])
        b["piso"] = reg["Piso altitudinal dominante"]
        b["pend_pct"] = num(reg["Pendiente promedio (%)"])
        b["pend_grados"] = num(reg["Pendiente promedio (grados)"])

        # MARCADOR DE AUSENCIA DEL MDE. La estadistica zonal devuelve 0 cuando el
        # poligono no tiene cobertura del modelo digital de elevacion. Un bloque
        # con altitud minima = maxima = 0 y pendiente = 0 no esta a nivel del mar
        # ni es plano: carece de dato. Se anula para no contaminar los agregados
        # y se declara en el control de consistencia.
        b["mde_sin_dato"] = (b["alt_min"] == 0 and b["alt_max"] == 0
                             and b["pend_pct"] == 0)
        if b["mde_sin_dato"]:
            b["alt_min"] = b["alt_max"] = b["amplitud"] = None
            b["pend_pct"] = b["pend_grados"] = None

        # CONFLICTO DE UNIDAD DE LA PENDIENTE (discrepancia D-P01). Las 117
        # plantillas V6 rotulan el valor del catalogo como PORCENTAJE y derivan
        # los grados como atan(v/100). El Informe DT Consolidado de Frias
        # (Ayabaca) concluyo, contrastando el catalogo con la capa «Pendientes
        # vector», que el MISMO valor esta expresado en GRADOS (error cuadratico
        # medio 1.14 grados frente a 12.70 leyendolo como porcentaje;
        # r = +0.956). No se resuelve aqui: se conservan las dos lecturas.
        b["pend_valor_catalogo"] = b["pend_pct"]
        b["pend_si_grados_pct"] = (round(math.tan(math.radians(b["pend_pct"])) * 100, 2)
                                   if b["pend_pct"] is not None else None)

        b["clase_pendiente"] = ("Sin cobertura del MDE" if b["mde_sin_dato"]
                                else reg["Clase de pendiente equivalente"])
        b["pendiente_campo"] = reg["Rango de pendiente declarado en campo"]
        b["forma_terreno"] = reg["Forma predominante del terreno"]
        b["posicion"] = reg["Posición fisiográfica"]
        b["exposicion"] = reg["Exposición / orientación"]
        b["afloramientos"] = reg["Afloramientos rocosos"]
        b["escarpes"] = reg["Escarpes activos"]
        b["remociones"] = reg["Remociones en masa activas"]

        # 3. Indices de vegetacion (Sentinel-2)
        b["msavi"] = num(reg["MSAVI 2024 — media del bloque"])
        b["msavi_clase"] = reg["MSAVI 2024 — clase de la media"]
        b["msavi_condicion"] = reg["Condición frente al umbral 0.4976"]
        b["ndvi_modal"] = reg["NDVI mediana 2025 — clase modal"]
        b["sup_ndvi"] = num(reg["Superficie clasificada NDVI 2025 (ha)"])
        b["desv_ndvi"] = num(reg["Desviación frente al catálogo (%)"])
        b["sup_msavi"] = num(reg["Superficie clasificada MSAVI 2024 (ha)"])
        b["desv_msavi"] = num(reg["Desviación MSAVI frente al catálogo (%)"])
        b["ha_bajo_umbral"] = num(reg["Superficie bajo umbral MSAVI 0.4976 (ha)"])
        b["pct_bajo_umbral"] = num(reg["% del bloque bajo umbral (brecha espectral)"])
        b["dn_dominante"] = reg["Clase DN dominante (MSAVI 2024)"]
        b["n_poligonos"] = num(reg["N.° de polígonos MSAVI del bloque"])

        # 4. Ecosistema y estado de conservacion
        b["ecosistema"] = reg["Tipo de ecosistema (UP)"]
        b["estado_conservacion"] = reg["Estado de conservación"]
        b["uso_actual"] = reg["Uso actual dominante del suelo"]
        b["cobertura_tipo"] = reg["Tipo de cobertura dominante"]
        b["cobertura_pct"] = reg["Cobertura vegetal total — campo (%)"]
        b["suelo_desnudo"] = reg["Suelo desnudo — campo (%)"]
        b["regeneracion"] = reg["Regeneración natural"]
        b["erosion"] = reg["Nivel general de erosión"]
        b["carcavas_ficha"] = reg["N.° de cárcavas registradas"]
        b["taxones"] = num(reg["Elenco florístico (n.° de taxones)"])
        b["estado_sanitario"] = reg["Estado sanitario"]

        # 5. Responsable y modalidad de la evaluacion
        b["responsable"] = reg["Responsable de la evaluación"]
        b["fecha"] = str(reg["Fecha de evaluación"])[:10]
        b["correlativo"] = reg["Correlativo de ficha"]
        b["instrumento"] = reg["Instrumento aplicado"]
        b["parcela"] = reg["Parcela de muestreo"]
        b["estaciones"] = num(reg["Estaciones fotográficas georreferenciadas"])
        b["acceso"] = reg["Modalidad de acceso"]

        # 6. Sintesis para la gestion del riesgo en contexto de cambio climatico
        b["causa"] = reg["Causa subyacente principal"]
        b["velocidad"] = reg["Velocidad de degradación"]
        b["reversibilidad"] = reg["Reversibilidad técnica"]
        b["urgencia"] = reg["Urgencia de intervención"]
        b["urgencia_erosion"] = reg["Urgencia de control de erosión"]
        b["recarga"] = reg["Zona de recarga hídrica"]
        b["peligro_integrado"] = reg["Peligro integrado preliminar (MCA-AHP)"]
        b["prioridad"] = reg["Prioridad de intervención"]
        b["verificacion"] = reg["Estado de verificación de campo"]

        # Distribucion areal por clase espectral
        b["msavi_clases"] = _clases_msavi(reg.get("_msavi_clases", {}))
        b["ndvi_clases"] = _clases_ndvi(reg.get("_ndvi_clases", {}))

        # Control de consistencia de la plantilla
        b["consistencia"] = [
            {"codigo": f[0], "campo": f[1], "discrepancia": f[2],
             "calificacion": f[3], "tratamiento": f[4]}
            for f in reg.get("_consistencia", []) if f and f[0]
        ]

        # Puntos georreferenciados de la ficha DT
        b["puntos"] = [
            {"codigo": p[0], "naturaleza": p[1], "este": p[2], "norte": p[3],
             "dist_centroide": p[4], "estado": p[5], "contenido": p[6]}
            for p in reg.get("_puntos", []) if p and p[0]
        ]

        # --- Integracion con el E3 Estudio de Geologia (ISL) ---
        g = isl.get(cod)
        if g:
            b["isl_mm"] = g["isl_mm"]
            b["isl_clase_mm"] = g["clase_mm"]
            b["isl_eh"] = g["isl_eh"]
            b["isl_clase_eh"] = g["clase_eh"]
            b["geo_unidad"] = g["unidad"]
            b["geo_n_unidades"] = g["nunid"]
            b["geo_area"] = g["area"]
        else:  # no debe ocurrir: la cobertura verificada es 117/117
            b["isl_mm"] = b["isl_eh"] = None
            b["isl_clase_mm"] = b["isl_clase_eh"] = "Sin dato en el E3 Geología"
            b["geo_unidad"] = "Sin dato en el E3 Geología"
            b["geo_n_unidades"] = b["geo_area"] = None

        # --- Inventario de carcavas codificadas ---
        cv = carcavas.get(cod, [])
        b["carcavas"] = [
            {"codigo": c["Codigo"], "este": float(c["Este"]), "norte": float(c["Norte"]),
             "longitud": float(c["Long"]), "alt_min": float(c["Alt_min"]),
             "alt_max": float(c["Alt_max"]), "alt_prom": float(c["Alt_prom"]),
             "pend_prom": float(c["Pend_prom"]), "iirreg": float(c["Iirreg"]),
             "ndvi": float(c["NDVI"]), "clase_morf": c["Cl_Morf"], "clase_ndvi": c["Cl_NDVI"],
             "dist_fuente": c["Dist"], "prov_fuente": c["Prov"]}
            for c in cv
        ]
        b["n_carcavas"] = len(cv)
        b["long_carcavas"] = round(sum(c["longitud"] for c in b["carcavas"]), 2) if cv else 0.0

        bloques.append(b)

    bloques.sort(key=lambda x: (-(x["area_ha"] or 0),))
    return bloques


def _clases_msavi(d):
    """Extrae la distribucion areal por clase DN del MSAVI 2024."""
    orden = ["> 0.6139 · DN 5", "0.4976 - 0.6139 · DN 4", "0.3813 - 0.4976 · DN 3",
             "0.2650 - 0.3813 · DN 2", "<= 0.2650 · DN 1"]
    out = OrderedDict()
    for k in orden:
        if k in d:
            out[k] = (num(d[k][0]), num(d[k][1]))
    return out


def _clases_ndvi(d):
    orden = ["Vegetación alta", "Vegetación mediana", "Vegetación ligera", "Tierra desnuda"]
    out = OrderedDict()
    for k in orden:
        if k in d:
            out[k] = (num(d[k][0]), num(d[k][1]))
    return out


def por_provincia(bloques, provincia):
    return [b for b in bloques if b["provincia"] == provincia]


def distritos_de(bs):
    """Distritos ordenados por superficie descendente."""
    agg = defaultdict(lambda: {"n": 0, "ha": 0.0})
    for b in bs:
        agg[b["distrito"]]["n"] += 1
        agg[b["distrito"]]["ha"] += b["area_ha"] or 0
    return sorted(agg.items(), key=lambda kv: -kv[1]["ha"])


def ponderada(bs, campo):
    """Media ponderada por superficie sobre los bloques que tienen el dato."""
    sub = [b for b in bs if b[campo] is not None and b["area_ha"]]
    ha = sum(b["area_ha"] for b in sub)
    if not ha:
        return None
    return sum(b[campo] * b["area_ha"] for b in sub) / ha


def con_dato(bs, campo):
    return [b for b in bs if b[campo] is not None]


def resumen(bs):
    """Estadisticos agregados de un conjunto de bloques."""
    ha = sum(b["area_ha"] or 0 for b in bs)
    msavi = [b["msavi"] for b in bs if b["msavi"] is not None]
    pend = [b["pend_pct"] for b in bs if b["pend_pct"] is not None]
    brecha = sum(b["ha_bajo_umbral"] or 0 for b in bs)
    res = {
        "n": len(bs),
        "ha": ha,
        "distritos": len({b["distrito"] for b in bs}),
        "microcuencas": len({b["microcuenca"] for b in bs}),
        "alt_min": min((b["alt_min"] for b in bs if b["alt_min"] is not None), default=None),
        "alt_max": max((b["alt_max"] for b in bs if b["alt_max"] is not None), default=None),
        "msavi_medio": sum(msavi) / len(msavi) if msavi else None,
        "msavi_ponderado": (sum(b["msavi"] * b["area_ha"] for b in bs
                                if b["msavi"] is not None and b["area_ha"]) / ha) if ha else None,
        "msavi_min": min(msavi) if msavi else None,
        "msavi_max": max(msavi) if msavi else None,
        "pend_media": sum(pend) / len(pend) if pend else None,
        "pend_min": min(pend) if pend else None,
        "pend_max": max(pend) if pend else None,
        "ha_brecha": brecha,
        "pct_brecha": 100 * brecha / ha if ha else None,
        "n_carcavas": sum(b["n_carcavas"] for b in bs),
        "long_carcavas": sum(b["long_carcavas"] for b in bs),
        "bloques_con_carcavas": len([b for b in bs if b["n_carcavas"]]),
        "sin_mde": [b["codigo"] for b in bs if b.get("mde_sin_dato")],
        "ha_sin_mde": sum(b["area_ha"] for b in bs if b.get("mde_sin_dato")),
        "n_con_mde": len([b for b in bs if not b.get("mde_sin_dato")]),
    }
    res["pend_ponderada"] = ponderada(bs, "pend_pct")
    res["msavi_ponderada"] = ponderada(bs, "msavi")
    return res


def mediana(vals):
    v = sorted(x for x in vals if x is not None)
    if not v:
        return None
    m = len(v) // 2
    return v[m] if len(v) % 2 else (v[m - 1] + v[m]) / 2


def conteo(bs, campo):
    """Conteo de bloques y superficie por valor de un campo categorico."""
    agg = defaultdict(lambda: {"n": 0, "ha": 0.0})
    for b in bs:
        agg[b[campo]]["n"] += 1
        agg[b[campo]]["ha"] += b["area_ha"] or 0
    return sorted(agg.items(), key=lambda kv: -kv[1]["ha"])


if __name__ == "__main__":
    bl = cargar_bloques()
    print("Bloques cargados:", len(bl))
    for p in PROVINCIAS:
        bs = por_provincia(bl, p)
        r = resumen(bs)
        print(f"{p:14s} n={r['n']:3d}  ha={r['ha']:10,.2f}  distritos={r['distritos']:2d} "
              f"MSAVI={r['msavi_medio']:.4f}  brecha={r['pct_brecha']:.2f}%  "
              f"carcavas={r['n_carcavas']}")
