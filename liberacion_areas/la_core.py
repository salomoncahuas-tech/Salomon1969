# -*- coding: utf-8 -*-
"""
Lógica pura (sin Streamlit ni base de datos) del módulo de Liberación de Áreas – Proyecto IN Piura.
- Aplanado de envíos KoboToolbox (API v2 JSON o exportación XLSX con nombres XML).
- Conversión lat/lon → UTM WGS84 Zona 17S (EPSG:32717) y validación de rangos.
- Validación espacial de puntos y polígonos contra la unidad (shapely).
- Reglas de lotes SUS (≥ 1 ha, Σ ≤ 10 % del bloque, DENTRO / CONTIGUO ≤ 50 m).
- Estado de importación (NUEVO / DUPLICADO / OBSERVADO) y estado LA por predio.
Todas las geometrías se manejan en EPSG:32717.
"""
from __future__ import annotations

import json
import math
from dataclasses import dataclass, field
from typing import Any, Iterable

from pyproj import Transformer
from shapely.geometry import Point, Polygon, shape
from shapely.ops import transform as shp_transform

# ------------------------------------------------------------------ constantes
EPSG_UTM = 32717
E_MIN, E_MAX = 450_000, 750_000
N_MIN, N_MAX = 9_300_000, 9_600_000
GPS_PREC_MAX_M = 10.0
SUS_MIN_HA = 1.0
SUS_MAX_PCT_BLOQUE = 0.10
SUS_CONTIGUO_MAX_M = 50.0
BLOQUE_MIN_PARA_SUS_HA = 10.0
TOL_DENTRO_M2 = 1.0            # tolerancia numérica para considerar "100 % dentro"
SEMAFORO_VERDE = 0.80          # [POR CONFIRMAR]
SEMAFORO_AMBAR = 0.50          # [POR CONFIRMAR]

ESTADOS_LA = ["LA-0", "LA-1", "LA-2", "LA-3", "LA-4", "LA-5", "LA-6"]
ESTADOS_EXCEPCION = ["NEG", "OBS", "EXC"]

FORMULARIOS = {
    "f_la_01_reunion": {"codigo": "F-LA-01", "tabla": "la_reuniones", "nombre": "Reunión / asamblea informativa"},
    "f_la_02_titular": {"codigo": "F-LA-02", "tabla": "la_fichas_titular", "nombre": "Ficha predial y socio-territorial"},
    "f_la_03_inspeccion": {"codigo": "F-LA-03", "tabla": "la_inspecciones", "nombre": "Inspección in situ y croquis"},
    "f_la_04_actas": {"codigo": "F-LA-04", "tabla": "la_actas", "nombre": "Registro de suscripción de actas"},
    "f_la_06_vivero": {"codigo": "F-LA-06", "tabla": "la_vivero_alternativas", "nombre": "Evaluación de terreno – Vivero Central"},
}

_TR = Transformer.from_crs(4326, EPSG_UTM, always_xy=True)


# ------------------------------------------------------------------ aplanado
def aplanar(registro: dict) -> dict:
    """Quita los prefijos de grupo ('g_ubicacion/asistente' → 'asistente').
    Los repeats (listas de dicts) se conservan aplanados en su interior."""
    out: dict[str, Any] = {}
    for k, v in registro.items():
        clave = k if k.startswith("_") or k.startswith("meta/") else k.split("/")[-1]
        if k.startswith("meta/"):
            clave = k.replace("meta/", "meta_")
        if isinstance(v, list) and v and isinstance(v[0], dict) and not k.startswith("_"):
            out[clave] = [aplanar(x) for x in v]
        else:
            out[clave] = v
    return out


def kobo_uuid(reg: dict) -> str | None:
    """Clave única del envío: meta/instanceID (uuid:...) o _uuid."""
    iid = reg.get("meta_instanceID") or reg.get("meta/instanceID") or reg.get("instanceID")
    if iid:
        return str(iid).replace("uuid:", "")
    u = reg.get("_uuid")
    return str(u) if u else None


def unir_repeats_xlsx(principal: list[dict], repeats: dict[str, list[dict]]) -> list[dict]:
    """Exportación XLSX de Kobo: cada repeat viene en otra hoja con _parent_index = _index del principal."""
    por_idx = {r.get("_index"): r for r in principal}
    for nombre, filas in repeats.items():
        for f in filas:
            padre = por_idx.get(f.get("_parent_index"))
            if padre is not None:
                padre.setdefault(nombre, []).append({k: v for k, v in f.items() if not k.startswith("_")})
    return principal


# ------------------------------------------------------------------ geometría
def _num(x) -> float | None:
    try:
        if x is None or (isinstance(x, float) and math.isnan(x)) or str(x).strip() == "":
            return None
        return float(x)
    except (TypeError, ValueError):
        return None


def parse_geopoint(txt: str | None) -> dict | None:
    """'lat lon alt precision' (ODK) → dict con lat, lon, alt, prec, este, norte, en_rango."""
    if not txt or not isinstance(txt, str):
        return None
    p = txt.strip().split()
    if len(p) < 2:
        return None
    lat, lon = float(p[0]), float(p[1])
    alt = float(p[2]) if len(p) > 2 else None
    prec = float(p[3]) if len(p) > 3 else None
    e, n = _TR.transform(lon, lat)
    return {"lat": lat, "lon": lon, "alt": alt, "prec": prec, "este": round(e, 2), "norte": round(n, 2),
            "en_rango": E_MIN <= e <= E_MAX and N_MIN <= n <= N_MAX}


def parse_geoshape(txt: str | None) -> Polygon | None:
    """'lat lon alt prec;lat lon alt prec;…' → Polygon en EPSG:32717 (None si < 3 vértices o inválido)."""
    if not txt or not isinstance(txt, str):
        return None
    coords = []
    for tramo in txt.strip().strip(";").split(";"):
        p = tramo.strip().split()
        if len(p) >= 2:
            coords.append(_TR.transform(float(p[1]), float(p[0])))
    if len(coords) < 3:
        return None
    pol = Polygon(coords)
    if not pol.is_valid:
        pol = pol.buffer(0)
    return pol if not pol.is_empty else None


def geojson_a_unidades(geojson: dict, campo_codigo: str = "name") -> dict[str, Any]:
    """Carga un GeoJSON exportado desde QGIS EN EPSG:32717 → {codigo: geometría shapely}.
    Si el GeoJSON está en EPSG:4326 (lon/lat), se reproyecta automáticamente."""
    crs = json.dumps(geojson.get("crs", {}))
    es_utm17 = "32717" in crs
    out = {}
    for f in geojson.get("features", []):
        cod = str(f["properties"].get(campo_codigo, "")).strip()
        if not cod:
            continue
        g = shape(f["geometry"])
        if not es_utm17:
            minx, miny, maxx, maxy = g.bounds
            if -180 <= minx <= 180 and -90 <= miny <= 90:
                g = shp_transform(lambda x, y, z=None: _TR.transform(x, y), g)
            else:
                raise ValueError("El GeoJSON no está en EPSG:32717 ni en EPSG:4326. Reproyecte en QGIS antes de cargar "
                                 "(la capa de idoneidad SUS original está en 18S).")
        out[cod] = g
    return out


def validar_punto(pt: dict | None, geom_unidad) -> tuple[str, float | None]:
    """→ ('DENTRO' | 'FUERA_POLIGONO' | 'SIN_GEOMETRIA' | 'SIN_PUNTO' | 'FUERA_RANGO', distancia_m)."""
    if pt is None:
        return "SIN_PUNTO", None
    if not pt["en_rango"]:
        return "FUERA_RANGO", None
    if geom_unidad is None:
        return "SIN_GEOMETRIA", None
    p = Point(pt["este"], pt["norte"])
    if geom_unidad.contains(p) or geom_unidad.touches(p):
        return "DENTRO", 0.0
    return "FUERA_POLIGONO", round(geom_unidad.distance(p), 1)


def posicion_sus(pol_sus: Polygon, geom_bloque) -> tuple[str, float]:
    """DENTRO = 100 % dentro del bloque; CONTIGUO = fuera, a ≤ 50 m; FUERA = > 50 m (no admisible)."""
    fuera = pol_sus.difference(geom_bloque).area
    if fuera <= TOL_DENTRO_M2:
        return "DENTRO", 0.0
    dist = round(pol_sus.distance(geom_bloque), 1)
    return ("CONTIGUO" if dist <= SUS_CONTIGUO_MAX_M else "FUERA"), dist


# ------------------------------------------------------------------ validaciones por formulario
@dataclass
class Resultado:
    kobo_uuid: str | None
    form_id: str
    cod_unidad: str | None
    cod_predio: str | None
    estado_import: str = "NUEVO"
    motivos: list[str] = field(default_factory=list)
    datos: dict = field(default_factory=dict)

    def observar(self, motivo: str):
        self.estado_import = "OBSERVADO"
        self.motivos.append(motivo)


def validar_envio(form_id: str, reg: dict, catalogo: dict[str, dict], geoms: dict[str, Any],
                  uuids_existentes: set[str], sus_areas: dict[str, tuple[str, float]] | None = None) -> Resultado:
    """Valida un envío ya aplanado.
    catalogo: {codigo: fila de unidades.csv}; geoms: {codigo: geometría 32717}.
    sus_areas: {codigo_lote: (bloque_ref, area_ha)} con el área vigente de cada lote SUS (catálogo o
    medición de campo). Es mutable: cada lote validado actualiza su área, de modo que la Σ por bloque
    considera también los envíos del mismo lote de importación."""
    uid = kobo_uuid(reg)
    r = Resultado(uid, form_id, reg.get("cod_unidad"), reg.get("cod_predio"), datos=reg)
    if not uid:
        r.observar("Envío sin instanceID")
    elif uid in uuids_existentes:
        r.estado_import = "DUPLICADO"
        r.motivos.append("Ya importado (kobo_uuid)")
        return r

    unidad = reg.get("cod_unidad") or reg.get("unidad")
    tipo = reg.get("tipo_unidad")
    if form_id != "f_la_01_reunion" and tipo != "vivero" and form_id != "f_la_06_vivero":
        if unidad not in catalogo:
            r.observar(f"Unidad '{unidad}' no existe en el catálogo vigente")
        else:
            cat = catalogo[unidad]
            asis = reg.get("asistente")
            if asis not in ("ESP", cat.get("asistente")):
                r.observar(f"Asistente {asis} distinto al asignado en catálogo ({cat.get('asistente')})")

    # DNI en repeats y campos
    for campo in ("tit_dni", "cony_dni", "cc_presidente_dni", "cc_comunero_dni"):
        v = reg.get(campo)
        if v and not (str(v).isdigit() and len(str(v)) == 8):
            r.observar(f"{campo} inválido")

    # --- puntos GPS
    campos_pt = {"f_la_01_reunion": "gps_evento", "f_la_03_inspeccion": "punto_interior", "f_la_06_vivero": "gps_centro"}
    if form_id in campos_pt:
        pt = parse_geopoint(reg.get(campos_pt[form_id]))
        if pt:
            r.datos.update({"este": pt["este"], "norte": pt["norte"], "altitud": pt["alt"], "precision_m": pt["prec"]})
            if not pt["en_rango"]:
                r.observar(f"Coordenada fuera de rango UTM 17S (E {pt['este']:.0f}, N {pt['norte']:.0f})")
            if pt["prec"] is not None and pt["prec"] > GPS_PREC_MAX_M:
                r.observar(f"Precisión GPS {pt['prec']:.1f} m > {GPS_PREC_MAX_M} m")
        if form_id == "f_la_03_inspeccion":
            est, dist = validar_punto(pt, geoms.get(unidad))
            r.datos["validacion_espacial"] = est
            r.datos["distancia_m"] = dist
            if est == "FUERA_POLIGONO":
                r.observar(f"FUERA_POLIGONO: punto interior a {dist} m de la unidad")
            elif est in ("SIN_PUNTO", "FUERA_RANGO"):
                r.observar(f"Punto interior: {est}")
            # puntos del repeat
            fuera = 0
            for p in reg.get("r_puntos", []) or []:
                gp = parse_geopoint(p.get("pt_gps"))
                e2, _ = validar_punto(gp, geoms.get(unidad))
                if e2 == "FUERA_POLIGONO":
                    fuera += 1
            r.datos["puntos_fuera"] = fuera

    # --- lotes SUS (F-LA-03)
    if form_id == "f_la_03_inspeccion" and tipo == "lote_sus":
        _validar_sus(r, reg, catalogo, geoms, sus_areas if sus_areas is not None else {})
    if form_id == "f_la_03_inspeccion" and "bosque_natural" in str(reg.get("interferencias", "")).split():
        r.observar("Interferencia con BOSQUE NATURAL: SUS solo como manejo / enriquecimiento (Ley 29763 [VERIFICAR])")

    # --- actas (F-LA-04)
    if form_id == "f_la_04_actas":
        _validar_acta(r, reg, catalogo)

    # --- ficha titular (F-LA-02)
    if form_id == "f_la_02_titular":
        if reg.get("consentimiento") == "no":
            r.observar("Sin consentimiento de datos personales")
        if reg.get("tipo_titularidad") == "sucesion" and reg.get("sucesion_estado") != "declarada":
            r.observar("Sucesión no declarada: requiere firma de todos los herederos → derivar a Especialista Legal")
        if reg.get("tipo_titularidad") == "sin_titular":
            r.observar("Sin titular identificado / en conflicto → Especialista Legal")
    return r


def sus_areas_desde_catalogo(catalogo: dict[str, dict]) -> dict[str, tuple[str, float]]:
    return {c: (str(f.get("bloque_ref")), _num(f.get("area_ha")) or 0.0)
            for c, f in catalogo.items() if f.get("tipo_unidad") == "lote_sus"}


def _validar_sus(r: Resultado, reg, catalogo, geoms, sus_areas):
    bloque = reg.get("bloque_ref") or catalogo.get(reg.get("cod_unidad"), {}).get("bloque_ref")
    area_bloque = _num(reg.get("area_bloque_ha")) or _num(catalogo.get(bloque, {}).get("area_ha"))
    pol = parse_geoshape(reg.get("poligono_sus"))
    if pol is None:
        r.observar("Lote SUS sin polígono válido")
        return
    area = round(pol.area / 10_000, 3)
    r.datos["area_sus_ha"] = area
    r.datos["geom_wkt"] = pol.wkt
    if area_bloque is not None and area_bloque < BLOQUE_MIN_PARA_SUS_HA:
        r.datos["estado_sus"] = "SIN_SUS"
        r.observar(f"Bloque {bloque} de {area_bloque} ha (< 10 ha): SIN_SUS salvo excepción ANIN")
    if area < SUS_MIN_HA:
        r.observar(f"Lote SUS de {area} ha < 1 ha")
    if area_bloque:
        lote = reg.get("cod_unidad")
        suma = sum(a for c, (b, a) in sus_areas.items() if b == str(bloque) and c != lote) + area
        tope = SUS_MAX_PCT_BLOQUE * area_bloque
        r.datos["pct_bloque"] = round(area / area_bloque * 100, 2)
        if suma > tope + 1e-6:
            r.observar(f"Σ lotes SUS del bloque {bloque} = {suma:.3f} ha > 10 % ({tope:.3f} ha); exceso {suma - tope:.3f} ha")
        sus_areas[lote] = (str(bloque), area)
    gb = geoms.get(str(bloque))
    if gb is not None:
        pos, dist = posicion_sus(pol, gb)
        r.datos["posicion_sus_calc"] = pos
        r.datos["distancia_bloque_m"] = dist
        if pos == "FUERA":
            r.observar(f"Lote SUS a {dist} m del bloque (> 50 m): no admisible")
        declarada = reg.get("posicion_sus_campo")
        if declarada and declarada != pos:
            r.motivos.append(f"Aviso: posición declarada {declarada} ≠ calculada {pos}")
    else:
        r.motivos.append(f"Aviso: sin geometría del bloque {bloque}; posición no verificada")


REQ_ACTA = {"A-03": ["coordenadas", "croquis", "titularidad", "gratuidad", "plazo", "firmas", "fedatario", "dj", "datos"],
            "A-04": ["coordenadas", "croquis", "titularidad", "gratuidad", "plazo", "firmas", "fedatario", "dj", "datos"],
            "A-07": ["coordenadas", "titularidad", "firmas", "datos"]}


def _validar_acta(r: Resultado, reg, catalogo):
    tipo = reg.get("tipo_acta")
    chk = set(str(reg.get("checklist", "")).split())
    faltan = [c for c in REQ_ACTA.get(tipo, ["firmas", "datos"]) if c not in chk]
    if faltan:
        r.observar(f"Acta {tipo} incompleta: falta {', '.join(faltan)}")
    firmantes = reg.get("r_firmantes", []) or []
    roles = [f.get("fir_rol") for f in firmantes]
    if tipo == "A-04" and "titular" not in roles and "representante" not in roles and "heredero" not in roles:
        r.observar("A-04 sin firma del titular / representante / herederos")
    if tipo == "A-03":
        if "presidente_cc" not in roles:
            r.observar("A-03 sin firma del presidente comunal")
        if reg.get("libro_actas") != "si":
            r.observar("A-03 no transcrita al libro de actas")
    if any(f.get("fir_huella") == "no" for f in firmantes):
        r.observar("Firmante(s) sin huella digital")
    area = _num(reg.get("area_comprometida_ha"))
    area_u = _num(catalogo.get(reg.get("cod_unidad"), {}).get("area_ha"))
    if area and area_u and area > area_u * 1.05:
        r.observar(f"Área comprometida {area} ha > área de la unidad {area_u} ha")
    if reg.get("estado_acta") == "observada":
        r.observar("Acta marcada como OBSERVADA en campo")


# ------------------------------------------------------------------ estado LA y semáforo
def estado_predio(ev: dict) -> str:
    """ev: indicadores del predio (bool) y excepción opcional.
    Claves: excepcion, titular, socializado, inspeccion_dentro, acta_conforme, docs_completos, expediente_conforme."""
    if ev.get("excepcion") in ESTADOS_EXCEPCION:
        return ev["excepcion"]
    nivel = 0
    for i, clave in enumerate(["titular", "socializado", "inspeccion_dentro", "acta_conforme", "docs_completos", "expediente_conforme"], 1):
        if ev.get(clave):
            nivel = i
        else:
            break
    return ESTADOS_LA[nivel]


def semaforo_unidad(predios: Iterable[dict], area_unidad_ha: float) -> tuple[str, float]:
    """predios: [{estado, area_ha, nucleo(bool)}] → (VERDE/AMBAR/ROJO/SIN_DATOS, % área con estado ≥ LA-4)."""
    predios = list(predios)
    if not predios or not area_unidad_ha:
        return "SIN_DATOS", 0.0
    ok = sum((_num(p.get("area_ha")) or 0) for p in predios if p.get("estado") in ("LA-4", "LA-5", "LA-6"))
    pct = min(ok / area_unidad_ha, 1.0)
    if any(p.get("estado") == "NEG" and p.get("nucleo") for p in predios):
        return "ROJO", pct
    if pct >= SEMAFORO_VERDE:
        return "VERDE", pct
    if pct >= SEMAFORO_AMBAR:
        return "AMBAR", pct
    return "ROJO", pct


def clasificacion_matriz(estado: str, motivos: list[str] | None = None) -> str:
    """Clasificación de la matriz predial (TdR Especialista Predial)."""
    m = " ".join(motivos or []).lower()
    if estado in ("NEG",) or "conflicto" in m or "superposici" in m:
        return "Posible conflicto"
    if "sucesi" in m or "legal" in m or estado == "OBS":
        return "Requiere gestión legal específica"
    if estado in ("LA-4", "LA-5", "LA-6"):
        return "Disponibilidad preliminar"
    return "Requiere verificación adicional"


def siguiente_titular(codigos: Iterable[str]) -> str:
    nums = [int(c[1:]) for c in codigos if c and c.startswith("T") and c[1:].isdigit()]
    return f"T{(max(nums) + 1) if nums else 1:04d}"


# ------------------------------------------------------------------ vivero: puntaje multicriterio [SUPUESTO]
PESOS_VIVERO = {"agua": 30, "acceso": 20, "pendiente": 15, "riesgo": 15, "energia": 5, "disposicion": 15}


def puntaje_vivero(reg: dict) -> float:
    """Puntaje 0–100 para comparar alternativas de Vivero Central (pesos [SUPUESTO]: validar con el Especialista en IV)."""
    caudal = _num(reg.get("agua_caudal_ls")) or 0
    agua = 0 if reg.get("agua_fuente") in (None, "ninguna") else min(caudal / 1.0, 1.0) * (1 if reg.get("agua_permanencia") == "permanente" else 0.5)
    acceso = {"carretera": 1, "trocha": 0.7, "herradura": 0.2, "sin_camino": 0}.get(reg.get("acceso_tipo"), 0) * (1 if reg.get("acceso_camion") == "si" else 0.6)
    pend = {"p0_5": 1, "p5_15": 0.7, "p15_30": 0.3, "p30": 0}.get(reg.get("pendiente"), 0)
    niv = {"nulo": 1, "bajo": 0.75, "medio": 0.3, "alto": 0}
    riesgo = min(niv.get(reg.get("inundabilidad"), 0), niv.get(reg.get("deslizamiento"), 0))
    energia = {"red": 1, "cerca": 0.6, "no": 0}.get(reg.get("energia"), 0)
    disp = {"favorable": 1, "condicionada": 0.5, "negativa": 0}.get(reg.get("disposicion"), 0)
    v = {"agua": agua, "acceso": acceso, "pendiente": pend, "riesgo": riesgo, "energia": energia, "disposicion": disp}
    return round(sum(PESOS_VIVERO[k] * v[k] for k in PESOS_VIVERO), 1)
