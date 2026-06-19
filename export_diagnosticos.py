"""
Exportacion consolidada a Excel de los Diagnosticos Territorial (FDT) y
Social (FDS) del Proyecto IN Piura.

Genera UN solo archivo .xlsx por cada caso con varias hojas de calculo:

  FDT (exportar_fdt_consolidado):
    - "Diagnosticos"      : una fila por registro con todos los campos escalares.
    - "F-DT-02 Carcavas"  : filas de la matriz de carcavas (dt02_carcavas_json).
    - "F-DT-03 Floristica": inventario floristico (dt03_floristica_json).
    - "F-DT-03 Esp. clave": especies clave / indicadoras.
    - "F-DT-04 Causas"     : matriz de causas de degradacion.
    - "F-DT-04 Indicadores": indicadores cuantitativos.
    - "F-DT-05 Fuentes agua": fuentes de agua registradas.

  FDS (exportar_fds_consolidado):
    - "Resumen"           : una fila por registro (cabecera).
    - "F-DS-01".."F-DS-07": campos escalares de cada ficha.
    - Hojas de tablas      : actividades, actores, participantes, agenda,
                             acuerdos, conflictos, oportunidades, peligros,
                             cambios climaticos.

Solo depende de pandas + openpyxl (ya presentes en requirements.txt).
"""

import io
import json
import re

import pandas as pd


# ─── Utilidades ────────────────────────────────────────────────────────────

def _load_json(raw, default):
    """Carga un valor JSON guardado como texto; tolera valores vacios."""
    if raw in (None, "", "[]", "null"):
        return default
    if isinstance(raw, (list, dict)):
        return raw
    try:
        val = json.loads(raw)
        return val if val not in (None, "") else default
    except (json.JSONDecodeError, TypeError):
        return default


def _humanizar(key):
    """Convierte una clave tecnica (dt02_nivel_erosion) en etiqueta legible."""
    s = re.sub(r"^(dt0[1-5]_|ds0[1-7]_|f[1-7]_)", "", str(key))
    s = s.replace("_json", "").replace("_", " ").strip()
    return s[:1].upper() + s[1:] if s else key


def _sanear_hoja(nombre, usados):
    """Devuelve un nombre de hoja valido (<=31 chars, sin caracteres ilegales,
    unico dentro del libro)."""
    nombre = re.sub(r"[\\/?*\[\]:]", " ", str(nombre)).strip() or "Hoja"
    nombre = nombre[:31]
    base = nombre
    i = 2
    while nombre.lower() in usados:
        sufijo = f" {i}"
        nombre = base[:31 - len(sufijo)] + sufijo
        i += 1
    usados.add(nombre.lower())
    return nombre


def _autoajustar(ws):
    """Ajusta el ancho de columnas de forma aproximada."""
    for col in ws.columns:
        try:
            longitud = max(
                (len(str(c.value)) for c in col if c.value is not None),
                default=0)
        except ValueError:
            longitud = 0
        letra = col[0].column_letter
        ws.column_dimensions[letra].width = min(max(longitud + 2, 10), 60)


def _escribir_libro(hojas):
    """hojas: lista de (nombre, DataFrame). Devuelve bytes del .xlsx.
    Siempre escribe al menos una hoja para que el archivo sea valido."""
    if not hojas:
        hojas = [("Sin datos", pd.DataFrame({"Aviso": ["No hay registros"]}))]
    buf = io.BytesIO()
    usados = set()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for nombre, df in hojas:
            if df is None or df.empty:
                continue
            hoja = _sanear_hoja(nombre, usados)
            df.to_excel(writer, sheet_name=hoja, index=False)
            _autoajustar(writer.sheets[hoja])
        # Si todas estaban vacias, escribir una hoja de aviso.
        if not writer.sheets:
            pd.DataFrame({"Aviso": ["No hay registros para exportar"]}).to_excel(
                writer, sheet_name="Sin datos", index=False)
    buf.seek(0)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════
# DIAGNOSTICO TERRITORIAL (FDT)
# ═══════════════════════════════════════════════════════════════════════════

# Columnas escalares principales y su etiqueta legible. Las claves *_json se
# tratan aparte (hojas de tablas).
_FDT_CABECERA = [
    ("id", "ID"),
    ("bloque_codigo", "Bloque"),
    ("ficha", "Fichas"),
    ("fecha_evaluacion", "Fecha"),
    ("evaluador", "Evaluador"),
    ("brigada", "Brigada"),
    ("ficha_correlativo", "Ficha N°"),
    ("microcuenca", "Microcuenca"),
    ("distrito", "Distrito"),
    ("centro_poblado_cercano", "Centro Poblado"),
    ("comunidad_campesina_dt", "Comunidad Campesina"),
    ("hora_registro", "Hora"),
    ("utm_este_dt", "UTM Este"),
    ("utm_norte_dt", "UTM Norte"),
    ("altitud_gps", "Altitud GPS"),
]

# Campos escalares de cada ficha F-DT (orden de aparicion en el aplicativo).
_FDT_CAMPOS = [
    # F-DT-01
    "forma_terreno", "pendiente", "posicion_fisiografica",
    "exposicion_orientacion", "rango_altitudinal", "paisaje_dominante",
    "dt01_afloramientos_rocosos", "dt01_escarpes_activos",
    "dt01_reptacion_suelo", "dt01_deslizamientos_antiguos",
    "dt01_remociones_masa_activas", "dt01_observaciones",
    # F-DT-02
    "dt02_sellamiento_costra", "dt02_compactacion_pisoteo",
    "dt02_raices_expuestas", "dt02_nivel_erosion_general",
    "dt02_nivel_erosion_sintesis", "dt02_num_carcavas",
    "dt02_longitud_total_carcavas", "dt02_pct_bloque_carcavas",
    "dt02_erosion_laminar_pct", "dt02_patron_carcavas",
    "dt02_socavamiento_cauce", "dt02_urgencia_control", "dt02_observaciones",
    # F-DT-03
    "dt03_parcela_muestreo", "dt03_dim_parcela", "dt03_pendiente_parcela",
    "dt03_cobertura_total", "dt03_tipo_ecosistema",
    "dt03_superficie_ecosistema", "dt03_estado_conservacion_eco",
    "dt03_uso_dominante", "dt03_cobertura_dosel", "dt03_cobertura_arbustiva",
    "dt03_cobertura_herbacea", "dt03_cobertura_hojarasca",
    "dt03_suelo_desnudo", "dt03_altura_estrato_dom", "dt03_altura_max",
    "dt03_dap_promedio", "dt03_regeneracion_natural",
    "dt03_estado_sanitario", "dt03_presencia_epifitas",
    "dt03_fenologia_dominante", "dt03_tipo_cobertura_dom", "dt03_observaciones",
    # F-DT-04
    "dt04_causas_directas_texto", "dt04_causa_subyacente",
    "dt04_velocidad_degradacion", "dt04_reversibilidad",
    "dt04_urgencia_intervencion", "dt04_observaciones",
    # F-DT-05
    "dt05_zona_recarga", "dt05_humedad_persistente",
    "dt05_escorrentia_concentrada", "dt05_dist_captacion",
    "dt05_jass_captacion", "dt05_interferencia_riego",
    "dt05_sistema_riego_nombre", "dt05_modalidad_acceso",
    "dt05_via_principal", "dt05_tipo_via_final",
    "dt05_transitabilidad_seca", "dt05_transitabilidad_lluviosa",
    "dt05_tiempo_dist_capital", "dt05_tiempo_prov_capital",
    "dt05_senal_celular", "dt05_operador_celular", "dt05_alojamiento",
    "dt05_requiere_ronda", "dt05_contacto_ronda", "dt05_observaciones",
    "observaciones_generales",
]

# Hojas de tablas (columna JSON -> titulo de hoja).
_FDT_TABLAS = [
    ("dt02_carcavas_json", "F-DT-02 Carcavas"),
    ("dt03_floristica_json", "F-DT-03 Floristica"),
    ("dt03_especies_clave_json", "F-DT-03 Esp clave"),
    ("dt04_causas_json", "F-DT-04 Causas"),
    ("dt04_indicadores_json", "F-DT-04 Indicadores"),
    ("dt05_fuentes_agua_json", "F-DT-05 Fuentes agua"),
]


def exportar_fdt_consolidado(registros):
    """registros: lista de dicts (salida de obtener_todos_diagnosticos()).
    Devuelve bytes del .xlsx consolidado."""
    registros = registros or []
    hojas = []

    # Hoja principal: una fila por diagnostico.
    filas = []
    for r in registros:
        fila = {}
        for col, etiqueta in _FDT_CABECERA:
            fila[etiqueta] = r.get(col, "")
        for col in _FDT_CAMPOS:
            fila[_humanizar(col)] = r.get(col, "")
        filas.append(fila)
    hojas.append(("Diagnosticos", pd.DataFrame(filas)))

    # Hojas de tablas (matriz de carcavas, floristica, causas, etc.).
    for col_json, titulo in _FDT_TABLAS:
        filas_tabla = []
        for r in registros:
            items = _load_json(r.get(col_json, ""), [])
            if not isinstance(items, list):
                continue
            for it in items:
                if not isinstance(it, dict):
                    continue
                base = {
                    "ID Diagnostico": r.get("id", ""),
                    "Bloque": r.get("bloque_codigo", ""),
                }
                base.update(it)
                filas_tabla.append(base)
        if filas_tabla:
            hojas.append((titulo, pd.DataFrame(filas_tabla)))

    return _escribir_libro(hojas)


# ═══════════════════════════════════════════════════════════════════════════
# DIAGNOSTICO SOCIAL (FDS)
# ═══════════════════════════════════════════════════════════════════════════

_FDS_CABECERA = [
    ("id", "ID"),
    ("bloque_codigo", "Bloque"),
    ("ficha", "Ficha"),
    ("ficha_numero", "Ficha N°"),
    ("fecha_evaluacion", "Fecha"),
    ("evaluador", "Responsable"),
    ("microcuenca", "Microcuenca"),
    ("centro_poblado", "Centro Poblado"),
    ("comunidad_campesina", "Comunidad Campesina"),
    ("distrito", "Distrito"),
    ("provincia", "Provincia"),
    ("nombre_entrevistado", "Entrevistado"),
    ("dni_entrevistado", "DNI"),
    ("oficio_ocupacion", "Oficio / Ocupacion"),
    ("observaciones_generales", "Observaciones"),
]

# Slots de tablas (coinciden con streamlit_app._DS_TABLE_SLOTS).
_FDS_TABLAS = {
    "f1_activ": "Actividades economicas",
    "f2_actores": "Actores clave",
    "f4_part": "Participantes taller",
    "f4_agenda": "Agenda taller",
    "f4_acuerdos": "Acuerdos taller",
    "f5_conflictos": "Conflictos",
    "f5_oportunidades": "Oportunidades",
    "f6_peligros": "Peligros naturales",
    "f6_cambios": "Cambios climaticos",
}

_FICHAS_DS = ["F-DS-01", "F-DS-02", "F-DS-03", "F-DS-04",
              "F-DS-05", "F-DS-06", "F-DS-07"]


def _fds_form(reg):
    """Extrae el dict del formulario V4 (dsNN_data_v3) de un registro."""
    ficha = reg.get("ficha", "") or ""
    num = ficha.split("-")[-1] if ficha else ""
    raw = reg.get(f"ds{num}_data_v3", "") or ""
    val = _load_json(raw, {})
    return val if isinstance(val, dict) else {}


def exportar_fds_consolidado(registros):
    """registros: lista de dicts (salida de obtener_todos_diagnosticos_sociales()).
    Devuelve bytes del .xlsx consolidado."""
    registros = registros or []
    hojas = []

    # Hoja resumen: una fila por registro.
    filas = []
    for r in registros:
        fila = {}
        for col, etiqueta in _FDS_CABECERA:
            fila[etiqueta] = r.get(col, "")
        filas.append(fila)
    hojas.append(("Resumen", pd.DataFrame(filas)))

    # Una hoja por tipo de ficha con sus campos escalares.
    for ficha in _FICHAS_DS:
        filas_ficha = []
        for r in registros:
            if (r.get("ficha", "") or "") != ficha:
                continue
            form = _fds_form(r)
            fila = {
                "ID": r.get("id", ""),
                "Bloque": r.get("bloque_codigo", ""),
                "Fecha": r.get("fecha_evaluacion", ""),
                "Responsable": r.get("evaluador", ""),
                "Centro Poblado": r.get("centro_poblado", ""),
                "Entrevistado": r.get("nombre_entrevistado", ""),
                "DNI": r.get("dni_entrevistado", ""),
                "Oficio / Ocupacion": r.get("oficio_ocupacion", ""),
            }
            for k, v in form.items():
                if k in _FDS_TABLAS:
                    continue
                if isinstance(v, list):
                    v = "; ".join(str(x) for x in v)
                fila[_humanizar(k)] = v
            filas_ficha.append(fila)
        if filas_ficha:
            hojas.append((ficha, pd.DataFrame(filas_ficha)))

    # Hojas de tablas, consolidando todas las filas de todos los registros.
    for slot, titulo in _FDS_TABLAS.items():
        filas_tabla = []
        for r in registros:
            form = _fds_form(r)
            items = form.get(slot, [])
            if not isinstance(items, list):
                continue
            for it in items:
                if not isinstance(it, dict):
                    continue
                base = {
                    "ID Diagnostico": r.get("id", ""),
                    "Bloque": r.get("bloque_codigo", ""),
                    "Ficha": r.get("ficha", ""),
                }
                base.update(it)
                filas_tabla.append(base)
        if filas_tabla:
            hojas.append((titulo, pd.DataFrame(filas_tabla)))

    return _escribir_libro(hojas)
