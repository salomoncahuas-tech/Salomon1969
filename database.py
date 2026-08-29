"""
IN Piura - Plan de Ingreso / Verificacion de Campo
Modulo de base de datos con PERSISTENCIA EN LA NUBE.
Usa PostgreSQL (Supabase) - los datos sobreviven al ciclo
de sueno/despertar de Streamlit Community Cloud.
Cuenca alta del rio Piura, Peru.

CAMBIO: SQLite local -> PostgreSQL en Supabase.
Capa de compatibilidad automatica: el resto del codigo NO cambia.

CONFIGURACION en .streamlit/secrets.toml:
  DATABASE_URL = "postgresql://postgres.xxxx:password@host:port/postgres"
"""

import re
import time
import streamlit as st
import psycopg2
from psycopg2.extras import RealDictCursor
from datetime import datetime

DATABASE_URL = st.secrets["DATABASE_URL"]

_CONNECT_KWARGS = {
    "connect_timeout": 15,
    "keepalives": 1,
    "keepalives_idle": 30,
    "keepalives_interval": 10,
    "keepalives_count": 5,
    "sslmode": "require",
}


def _connect_with_retry(dsn, **kwargs):
    """Conecta a PostgreSQL con reintentos y backoff exponencial."""
    merged = {**_CONNECT_KWARGS, **kwargs}
    last_err = None
    for attempt in range(4):
        try:
            return psycopg2.connect(dsn, **merged)
        except psycopg2.OperationalError as e:
            last_err = e
            if attempt < 3:
                time.sleep(2 ** attempt)
    raise last_err


# == CAPA DE COMPATIBILIDAD SQLite -> PostgreSQL ===========================
# Traduce automaticamente:
#   ? -> %s  |  AUTOINCREMENT -> (nada)  |  GROUP_CONCAT -> STRING_AGG
#   cursor.lastrowid -> via RETURNING id
# ==========================================================================

def _translate_sql(sql):
    """Traduce SQL de SQLite a PostgreSQL."""
    sql = sql.replace('?', '%s')
    sql = sql.replace('AUTOINCREMENT', '')
    if 'PRAGMA' in sql.upper():
        return None
    sql = re.sub(
        r"GROUP_CONCAT\(([^)]+)\)",
        r"STRING_AGG(\1, ',')",
        sql)
    return sql


class _CursorWrapper:
    """Envuelve psycopg2 cursor para simular sqlite3.Cursor."""
    def __init__(self, pg_cursor):
        self._c = pg_cursor
        self.lastrowid = None
        self.description = None

    def execute(self, sql, params=None):
        sql = _translate_sql(sql)
        if sql is None:
            return
        is_insert = sql.strip().upper().startswith('INSERT')
        if is_insert and 'RETURNING' not in sql.upper():
            sql = sql.rstrip().rstrip(';') + ' RETURNING id'
        if params:
            self._c.execute(sql, params)
        else:
            self._c.execute(sql)
        self.description = self._c.description
        if is_insert:
            row = self._c.fetchone()
            self.lastrowid = row['id'] if row else None

    def fetchall(self):
        return self._c.fetchall()

    def fetchone(self):
        return self._c.fetchone()


class _ConnectionWrapper:
    """Envuelve psycopg2 connection para simular sqlite3.Connection."""
    def __init__(self):
        self._conn = _connect_with_retry(
            DATABASE_URL, cursor_factory=RealDictCursor)

    def cursor(self):
        return _CursorWrapper(self._conn.cursor())

    def execute(self, sql, params=None):
        sql = _translate_sql(sql)
        if sql is None:
            return
        cur = self._conn.cursor()
        if params:
            cur.execute(sql, params)
        else:
            cur.execute(sql)

    def commit(self):
        self._conn.commit()

    def close(self):
        self._conn.close()


# == INTERFAZ PUBLICA (identica a version SQLite) ==========================

def get_connection():
    """Retorna conexion a PostgreSQL (Supabase)."""
    return _ConnectionWrapper()


def _dictfetch(cursor):
    """Convierte los resultados del cursor en lista de dicts."""
    rows = cursor.fetchall()
    return [dict(row) for row in rows]


def _dictfetchone(cursor):
    """Convierte un resultado del cursor en dict."""
    row = cursor.fetchone()
    return dict(row) if row else None

def inicializar_bd():
    """Crea las tablas si no existen (idempotente)."""
    conn = _connect_with_retry(DATABASE_URL, sslmode="require")
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bloques (
            id SERIAL PRIMARY KEY,
            codigo TEXT UNIQUE NOT NULL,
            tipo_intervencion TEXT NOT NULL,
            cuenca TEXT NOT NULL,
            distrito TEXT NOT NULL,
            utm_este REAL NOT NULL,
            utm_norte REAL NOT NULL,
            utm_zona TEXT NOT NULL DEFAULT '17S',
            altitud REAL DEFAULT 0,
            area_hectareas REAL NOT NULL,
            responsable TEXT DEFAULT '',
            estado TEXT NOT NULL DEFAULT 'Pendiente',
            microcuenca TEXT DEFAULT '',
            provincia TEXT DEFAULT '',
            fecha_registro TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS inspecciones (
            id SERIAL PRIMARY KEY,
            bloque_id INTEGER NOT NULL REFERENCES bloques(id) ON DELETE CASCADE,
            fecha_visita TEXT NOT NULL,
            inspector TEXT NOT NULL,
            condiciones_climaticas TEXT,
            avance_fisico REAL DEFAULT 0,
            observaciones TEXT,
            desviaciones TEXT,
            registro_fotografico TEXT,
            codigo_verificacion TEXT,
            microcuenca TEXT DEFAULT '',
            archivos_pdf TEXT DEFAULT '',
            fecha_registro TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS indicadores_calidad (
            id SERIAL PRIMARY KEY,
            bloque_id INTEGER NOT NULL REFERENCES bloques(id) ON DELETE CASCADE,
            inspeccion_id INTEGER NOT NULL REFERENCES inspecciones(id) ON DELETE CASCADE,
            cobertura_vegetal_planificada REAL DEFAULT 0,
            cobertura_vegetal_lograda REAL DEFAULT 0,
            sobrevivencia_especies REAL DEFAULT 0,
            longitud_zanjas_ejecutada REAL DEFAULT 0,
            volumen_retencion_sedimentos REAL DEFAULT 0,
            porcentaje_cobertura_vegetal REAL DEFAULT 0,
            tipo_cobertura_vegetal TEXT DEFAULT '',
            vigor_cobertura_vegetal TEXT DEFAULT '',
            microcuenca TEXT DEFAULT '',
            fecha_registro TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS presupuesto (
            id SERIAL PRIMARY KEY,
            bloque_id INTEGER NOT NULL REFERENCES bloques(id) ON DELETE CASCADE,
            categoria TEXT NOT NULL,
            descripcion TEXT DEFAULT '',
            monto_planificado REAL DEFAULT 0,
            monto_ejecutado REAL DEFAULT 0,
            fuente_financiamiento TEXT DEFAULT '',
            fecha_registro TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cronograma (
            id SERIAL PRIMARY KEY,
            bloque_id INTEGER NOT NULL REFERENCES bloques(id) ON DELETE CASCADE,
            actividad TEXT NOT NULL,
            fecha_inicio_plan TEXT NOT NULL,
            fecha_fin_plan TEXT NOT NULL,
            fecha_inicio_real TEXT DEFAULT '',
            fecha_fin_real TEXT DEFAULT '',
            porcentaje_avance REAL DEFAULT 0,
            responsable TEXT DEFAULT '',
            observaciones TEXT DEFAULT '',
            estado TEXT NOT NULL DEFAULT 'Programado',
            fecha_registro TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS personal (
            id SERIAL PRIMARY KEY,
            nombre TEXT NOT NULL,
            cargo TEXT NOT NULL,
            especialidad TEXT DEFAULT '',
            telefono TEXT DEFAULT '',
            email TEXT DEFAULT '',
            bloque_asignado INTEGER REFERENCES bloques(id) ON DELETE SET NULL,
            activo INTEGER DEFAULT 1,
            fecha_registro TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS diagnostico_territorial (
            id SERIAL PRIMARY KEY,
            bloque_id INTEGER NOT NULL REFERENCES bloques(id) ON DELETE CASCADE,
            inspeccion_id INTEGER REFERENCES inspecciones(id) ON DELETE SET NULL,
            ficha TEXT NOT NULL,
            microcuenca TEXT DEFAULT '',
            fecha_evaluacion TEXT NOT NULL,
            evaluador TEXT DEFAULT '',
            forma_terreno TEXT DEFAULT '',
            pendiente TEXT DEFAULT '',
            posicion_fisiografica TEXT DEFAULT '',
            exposicion_orientacion TEXT DEFAULT '',
            paisaje_dominante TEXT DEFAULT '',
            rango_altitudinal TEXT DEFAULT '',
            precipitacion_anual TEXT DEFAULT '',
            temperatura_media TEXT DEFAULT '',
            humedad_relativa TEXT DEFAULT '',
            zona_vida TEXT DEFAULT '',
            presencia_heladas TEXT DEFAULT '',
            regimen_vientos TEXT DEFAULT '',
            textura_suelo TEXT DEFAULT '',
            color_suelo TEXT DEFAULT '',
            profundidad_efectiva TEXT DEFAULT '',
            pedregosidad TEXT DEFAULT '',
            drenaje TEXT DEFAULT '',
            presencia_erosion TEXT DEFAULT '',
            materia_organica TEXT DEFAULT '',
            tipo_cobertura TEXT DEFAULT '',
            densidad_cobertura TEXT DEFAULT '',
            estado_conservacion TEXT DEFAULT '',
            uso_actual_suelo TEXT DEFAULT '',
            conflicto_uso TEXT DEFAULT '',
            fuente_agua TEXT DEFAULT '',
            regimen_hidrico TEXT DEFAULT '',
            calidad_agua TEXT DEFAULT '',
            distancia_fuente_agua TEXT DEFAULT '',
            uso_recurso_hidrico TEXT DEFAULT '',
            tenencia_tierra TEXT DEFAULT '',
            organizacion_comunal TEXT DEFAULT '',
            actividad_economica TEXT DEFAULT '',
            accesibilidad_via TEXT DEFAULT '',
            distancia_centro_poblado TEXT DEFAULT '',
            servicios_basicos TEXT DEFAULT '',
            observaciones_generales TEXT DEFAULT '',
            fecha_registro TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS diagnostico_social (
            id SERIAL PRIMARY KEY,
            bloque_id INTEGER NOT NULL REFERENCES bloques(id) ON DELETE CASCADE,
            inspeccion_id INTEGER REFERENCES inspecciones(id) ON DELETE SET NULL,
            ficha TEXT NOT NULL,
            ficha_numero TEXT DEFAULT '',
            microcuenca TEXT DEFAULT '',
            fecha_evaluacion TEXT NOT NULL,
            evaluador TEXT DEFAULT '',
            provincia TEXT DEFAULT '',
            distrito TEXT DEFAULT '',
            centro_poblado TEXT DEFAULT '',
            comunidad_campesina TEXT DEFAULT '',
            coordenada_este REAL DEFAULT 0,
            coordenada_norte REAL DEFAULT 0,
            altitud REAL DEFAULT 0,
            codigo_ubigeo TEXT DEFAULT '',
            nombre_entrevistado TEXT DEFAULT '',
            dni_entrevistado TEXT DEFAULT '',
            oficio_ocupacion TEXT DEFAULT '',
            ds01_num_familias TEXT DEFAULT '',
            ds01_poblacion_hombres TEXT DEFAULT '',
            ds01_poblacion_mujeres TEXT DEFAULT '',
            ds01_poblacion_total TEXT DEFAULT '',
            ds01_idioma TEXT DEFAULT '',
            ds01_nivel_educativo TEXT DEFAULT '',
            ds01_tasa_migracion TEXT DEFAULT '',
            ds01_destino_migracion TEXT DEFAULT '',
            ds01_organizacion_comunal TEXT DEFAULT '',
            ds01_junta_directiva TEXT DEFAULT '',
            ds01_presidente_junta TEXT DEFAULT '',
            ds01_agua_potable_tipo TEXT DEFAULT '',
            ds01_agua_potable_cobertura TEXT DEFAULT '',
            ds01_saneamiento TEXT DEFAULT '',
            ds01_energia_tipo TEXT DEFAULT '',
            ds01_energia_cobertura TEXT DEFAULT '',
            ds01_telecomunicaciones TEXT DEFAULT '',
            ds01_telecom_operador TEXT DEFAULT '',
            ds01_acceso_vial TEXT DEFAULT '',
            ds01_distancia_capital TEXT DEFAULT '',
            ds01_transporte TEXT DEFAULT '',
            ds01_salud_tipo TEXT DEFAULT '',
            ds01_salud_distancia TEXT DEFAULT '',
            ds01_educacion TEXT DEFAULT '',
            ds01_actividades_economicas TEXT DEFAULT '',
            ds01_fuente_agua TEXT DEFAULT '',
            ds01_problemas_agua TEXT DEFAULT '',
            ds01_uso_recursos_forestales TEXT DEFAULT '',
            ds01_frecuencia_uso_forestal TEXT DEFAULT '',
            ds01_percepcion_cambios TEXT DEFAULT '',
            ds01_disposicion_participar TEXT DEFAULT '',
            ds01_comentario_disposicion TEXT DEFAULT '',
            ds01_activos_asociados TEXT DEFAULT '',
            ds01_tenencia_comunal_ha TEXT DEFAULT '',
            ds01_tenencia_privada_ha TEXT DEFAULT '',
            ds01_tenencia_estatal_ha TEXT DEFAULT '',
            ds02_registro_actores TEXT DEFAULT '',
            ds02_actores_gob_local TEXT DEFAULT '',
            ds02_actores_gob_regional TEXT DEFAULT '',
            ds02_actores_gob_nacional TEXT DEFAULT '',
            ds02_actores_comunidades TEXT DEFAULT '',
            ds02_actores_juntas_riego TEXT DEFAULT '',
            ds02_actores_comites_cuenca TEXT DEFAULT '',
            ds02_actores_ong TEXT DEFAULT '',
            ds02_actores_empresa TEXT DEFAULT '',
            ds02_actores_educacion TEXT DEFAULT '',
            ds02_actores_org_base TEXT DEFAULT '',
            ds03_nombre_entrevistado TEXT DEFAULT '',
            ds03_cargo_funcion TEXT DEFAULT '',
            ds03_institucion TEXT DEFAULT '',
            ds03_telefono_correo TEXT DEFAULT '',
            ds03_duracion TEXT DEFAULT '',
            ds03_resp_recursos_naturales TEXT DEFAULT '',
            ds03_resp_cambios_ambiente TEXT DEFAULT '',
            ds03_resp_problemas_ambientales TEXT DEFAULT '',
            ds03_resp_zonas_conservacion TEXT DEFAULT '',
            ds03_resp_actividades_economicas TEXT DEFAULT '',
            ds03_resp_abastecimiento_agua TEXT DEFAULT '',
            ds03_resp_productos_bosque TEXT DEFAULT '',
            ds03_resp_cadenas_productivas TEXT DEFAULT '',
            ds03_resp_organizaciones TEXT DEFAULT '',
            ds03_resp_decisiones_territorio TEXT DEFAULT '',
            ds03_resp_conflictos TEXT DEFAULT '',
            ds03_resp_proyectos_anteriores TEXT DEFAULT '',
            ds03_resp_experiencia_reforestacion TEXT DEFAULT '',
            ds03_resp_conocimiento_restauracion TEXT DEFAULT '',
            ds03_resp_expectativas TEXT DEFAULT '',
            ds03_resp_disposicion_participar TEXT DEFAULT '',
            ds03_resp_condiciones TEXT DEFAULT '',
            ds03_resp_conocimiento_merese TEXT DEFAULT '',
            ds03_resp_beneficiarios TEXT DEFAULT '',
            ds03_resp_instituciones_contribuyentes TEXT DEFAULT '',
            ds03_resp_experiencias_pago TEXT DEFAULT '',
            ds04_lugar_taller TEXT DEFAULT '',
            ds04_hora_inicio TEXT DEFAULT '',
            ds04_hora_fin TEXT DEFAULT '',
            ds04_convocante TEXT DEFAULT '',
            ds04_objetivo TEXT DEFAULT '',
            ds04_lista_participantes TEXT DEFAULT '',
            ds04_presentacion TEXT DEFAULT '',
            ds04_intervenciones TEXT DEFAULT '',
            ds04_preguntas_respuestas TEXT DEFAULT '',
            ds04_acuerdos TEXT DEFAULT '',
            ds04_observaciones TEXT DEFAULT '',
            ds05_conflictos TEXT DEFAULT '',
            ds05_oportunidades TEXT DEFAULT '',
            archivos_adjuntos TEXT DEFAULT '',
            observaciones_generales TEXT DEFAULT '',
            fecha_registro TEXT NOT NULL
        )
    """)

    # ── Tabla de Elementos Expuestos (AdR / Riesgos) ────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS elementos_expuestos (
            id SERIAL PRIMARY KEY,
            bloque_id INTEGER NOT NULL REFERENCES bloques(id) ON DELETE CASCADE,
            inspeccion_id INTEGER REFERENCES inspecciones(id) ON DELETE SET NULL,
            ficha TEXT NOT NULL,
            fecha_campo TEXT DEFAULT '',
            responsable_brigada TEXT DEFAULT '',
            centro_poblado TEXT DEFAULT '',
            coordenada_este TEXT DEFAULT '',
            coordenada_norte TEXT DEFAULT '',
            altitud TEXT DEFAULT '',
            ee01_registros TEXT DEFAULT '',
            ee02_registros TEXT DEFAULT '',
            ee02_total_viviendas TEXT DEFAULT '',
            ee02_total_poblacion TEXT DEFAULT '',
            ee03_registros TEXT DEFAULT '',
            ee04_registros TEXT DEFAULT '',
            ee05_tipo_ecosistema TEXT DEFAULT '',
            ee05_zona_vida TEXT DEFAULT '',
            ee05_cobertura_vegetal TEXT DEFAULT '',
            ee05_pct_cobertura TEXT DEFAULT '',
            ee05_especies_dominantes TEXT DEFAULT '',
            ee05_evidencia_degradacion TEXT DEFAULT '',
            ee05_tipo_degradacion TEXT DEFAULT '',
            ee05_nivel_degradacion TEXT DEFAULT '',
            ee05_pendiente TEXT DEFAULT '',
            ee05_tipo_suelo TEXT DEFAULT '',
            ee05_profundidad_efectiva TEXT DEFAULT '',
            ee05_presencia_carcavas TEXT DEFAULT '',
            ee05_presencia_quebrada TEXT DEFAULT '',
            ee05_nombre_quebrada TEXT DEFAULT '',
            ee05_fuentes_agua TEXT DEFAULT '',
            ee05_peligros_observados TEXT DEFAULT '',
            ee06_cuantificacion TEXT DEFAULT '',
            ee06_valoracion_vulnerabilidad TEXT DEFAULT '',
            ee06_nivel_vulnerabilidad TEXT DEFAULT '',
            ee06_nivel_riesgo TEXT DEFAULT '',
            ee07_registros TEXT DEFAULT '',
            observaciones_generales TEXT DEFAULT '',
            archivos_adjuntos TEXT DEFAULT '',
            fecha_registro TEXT NOT NULL
        )
    """)

    # ── Indices UNIQUE para evitar duplicados ────────────────────────────
    cursor.execute("""
        DO $$ BEGIN
            IF NOT EXISTS (SELECT 1 FROM pg_indexes WHERE indexname = 'uq_inspeccion_bloque_fecha_inspector') THEN
                CREATE UNIQUE INDEX uq_inspeccion_bloque_fecha_inspector
                ON inspecciones (bloque_id, fecha_visita, inspector);
            END IF;
        END $$
    """)
    cursor.execute("""
        DO $$ BEGIN
            IF NOT EXISTS (SELECT 1 FROM pg_indexes WHERE indexname = 'uq_dt_bloque_fecha_evaluador') THEN
                CREATE UNIQUE INDEX uq_dt_bloque_fecha_evaluador
                ON diagnostico_territorial (bloque_id, fecha_evaluacion, evaluador);
            END IF;
        END $$
    """)
    cursor.execute("""
        DO $$ BEGIN
            IF NOT EXISTS (SELECT 1 FROM pg_indexes WHERE indexname = 'uq_ds_bloque_ficha_fecha_evaluador') THEN
                CREATE UNIQUE INDEX uq_ds_bloque_ficha_fecha_evaluador
                ON diagnostico_social (bloque_id, ficha, fecha_evaluacion, evaluador);
            END IF;
        END $$
    """)
    cursor.execute("""
        DO $$ BEGIN
            IF NOT EXISTS (SELECT 1 FROM pg_indexes WHERE indexname = 'uq_presupuesto_bloque_cat_desc') THEN
                CREATE UNIQUE INDEX uq_presupuesto_bloque_cat_desc
                ON presupuesto (bloque_id, categoria, descripcion);
            END IF;
        END $$
    """)
    cursor.execute("""
        DO $$ BEGIN
            IF NOT EXISTS (SELECT 1 FROM pg_indexes WHERE indexname = 'uq_cronograma_bloque_actividad_inicio') THEN
                CREATE UNIQUE INDEX uq_cronograma_bloque_actividad_inicio
                ON cronograma (bloque_id, actividad, fecha_inicio_plan);
            END IF;
        END $$
    """)
    cursor.execute("""
        DO $$ BEGIN
            IF NOT EXISTS (SELECT 1 FROM pg_indexes WHERE indexname = 'uq_ee_bloque_ficha_fecha_responsable') THEN
                CREATE UNIQUE INDEX uq_ee_bloque_ficha_fecha_responsable
                ON elementos_expuestos (bloque_id, ficha, fecha_campo, responsable_brigada);
            END IF;
        END $$
    """)

    # ── Migracion: agregar columnas nuevas a tablas existentes ─────────────
    nuevas_columnas_ds = [
        ("ds01_activos_asociados", "TEXT DEFAULT ''"),
        ("ds01_tenencia_comunal_ha", "TEXT DEFAULT ''"),
        ("ds01_tenencia_privada_ha", "TEXT DEFAULT ''"),
        ("ds01_tenencia_estatal_ha", "TEXT DEFAULT ''"),
        ("ds03_resp_experiencia_reforestacion", "TEXT DEFAULT ''"),
        # ── Datos del entrevistado (Datos Generales y Localizacion) ────────
        ("nombre_entrevistado", "TEXT DEFAULT ''"),
        ("dni_entrevistado", "TEXT DEFAULT ''"),
        ("oficio_ocupacion", "TEXT DEFAULT ''"),
        # ── Diagnostico Social V3 (Plantilla validada, F-DS-01..07) ────────
        # Cada ficha guarda su formulario completo como JSON en una sola
        # columna dsNN_data_v3. Las columnas legacy (arriba) se siguen
        # alimentando donde aplica para compatibilidad con reportes.
        ("ds01_data_v3", "TEXT DEFAULT ''"),
        ("ds02_data_v3", "TEXT DEFAULT ''"),
        ("ds03_data_v3", "TEXT DEFAULT ''"),
        ("ds04_data_v3", "TEXT DEFAULT ''"),
        ("ds05_data_v3", "TEXT DEFAULT ''"),
        ("ds06_data_v3", "TEXT DEFAULT ''"),
        ("ds07_data_v3", "TEXT DEFAULT ''"),
    ]
    for col_name, col_type in nuevas_columnas_ds:
        cursor.execute(f"""
            DO $$ BEGIN
                ALTER TABLE diagnostico_social ADD COLUMN {col_name} {col_type};
            EXCEPTION WHEN duplicate_column THEN NULL;
            END $$
        """)

    # ── Diagnostico Territorial V5: nuevas columnas (Plantilla DT V5) ──────
    # F-DT-01..05 V5 reemplazan a las 6 fichas previas. Las columnas viejas
    # (textura_suelo, condiciones climaticas, etc.) se mantienen para
    # historico pero la nueva UI no las escribe.
    nuevas_columnas_dt = [
        # Comunes V5
        ("brigada", "TEXT DEFAULT ''"),
        ("ficha_correlativo", "TEXT DEFAULT ''"),
        ("altitud_gps", "TEXT DEFAULT ''"),
        ("centro_poblado_cercano", "TEXT DEFAULT ''"),
        ("comunidad_campesina_dt", "TEXT DEFAULT ''"),
        ("hora_registro", "TEXT DEFAULT ''"),
        ("utm_este_dt", "TEXT DEFAULT ''"),
        ("utm_norte_dt", "TEXT DEFAULT ''"),
        # F-DT-01 (reusa forma_terreno, pendiente, posicion_fisiografica,
        # exposicion_orientacion, rango_altitudinal, paisaje_dominante)
        ("dt01_afloramientos_rocosos", "TEXT DEFAULT ''"),
        ("dt01_escarpes_activos", "TEXT DEFAULT ''"),
        ("dt01_reptacion_suelo", "TEXT DEFAULT ''"),
        ("dt01_deslizamientos_antiguos", "TEXT DEFAULT ''"),
        ("dt01_remociones_masa_activas", "TEXT DEFAULT ''"),
        ("dt01_observaciones", "TEXT DEFAULT ''"),
        # F-DT-02
        ("dt02_sellamiento_costra", "TEXT DEFAULT ''"),
        ("dt02_compactacion_pisoteo", "TEXT DEFAULT ''"),
        ("dt02_raices_expuestas", "TEXT DEFAULT ''"),
        ("dt02_nivel_erosion_general", "TEXT DEFAULT ''"),
        ("dt02_carcavas_json", "TEXT DEFAULT ''"),
        ("dt02_nivel_erosion_sintesis", "TEXT DEFAULT ''"),
        ("dt02_num_carcavas", "TEXT DEFAULT ''"),
        ("dt02_longitud_total_carcavas", "TEXT DEFAULT ''"),
        ("dt02_pct_bloque_carcavas", "TEXT DEFAULT ''"),
        ("dt02_erosion_laminar_pct", "TEXT DEFAULT ''"),
        ("dt02_patron_carcavas", "TEXT DEFAULT ''"),
        ("dt02_socavamiento_cauce", "TEXT DEFAULT ''"),
        ("dt02_urgencia_control", "TEXT DEFAULT ''"),
        ("dt02_observaciones", "TEXT DEFAULT ''"),
        # F-DT-03
        ("dt03_parcela_muestreo", "TEXT DEFAULT ''"),
        ("dt03_dim_parcela", "TEXT DEFAULT ''"),
        ("dt03_pendiente_parcela", "TEXT DEFAULT ''"),
        ("dt03_cobertura_total", "TEXT DEFAULT ''"),
        ("dt03_tipo_ecosistema", "TEXT DEFAULT ''"),
        ("dt03_superficie_ecosistema", "TEXT DEFAULT ''"),
        ("dt03_estado_conservacion_eco", "TEXT DEFAULT ''"),
        ("dt03_uso_dominante", "TEXT DEFAULT ''"),
        ("dt03_cobertura_dosel", "TEXT DEFAULT ''"),
        ("dt03_cobertura_arbustiva", "TEXT DEFAULT ''"),
        ("dt03_cobertura_herbacea", "TEXT DEFAULT ''"),
        ("dt03_cobertura_hojarasca", "TEXT DEFAULT ''"),
        ("dt03_suelo_desnudo", "TEXT DEFAULT ''"),
        ("dt03_altura_estrato_dom", "TEXT DEFAULT ''"),
        ("dt03_altura_max", "TEXT DEFAULT ''"),
        ("dt03_dap_promedio", "TEXT DEFAULT ''"),
        ("dt03_regeneracion_natural", "TEXT DEFAULT ''"),
        ("dt03_estado_sanitario", "TEXT DEFAULT ''"),
        ("dt03_presencia_epifitas", "TEXT DEFAULT ''"),
        ("dt03_fenologia_dominante", "TEXT DEFAULT ''"),
        ("dt03_tipo_cobertura_dom", "TEXT DEFAULT ''"),
        ("dt03_floristica_json", "TEXT DEFAULT ''"),
        ("dt03_especies_clave_json", "TEXT DEFAULT ''"),
        ("dt03_observaciones", "TEXT DEFAULT ''"),
        # F-DT-04
        ("dt04_causas_json", "TEXT DEFAULT ''"),
        ("dt04_indicadores_json", "TEXT DEFAULT ''"),
        ("dt04_causas_directas_texto", "TEXT DEFAULT ''"),
        ("dt04_causa_subyacente", "TEXT DEFAULT ''"),
        ("dt04_velocidad_degradacion", "TEXT DEFAULT ''"),
        ("dt04_reversibilidad", "TEXT DEFAULT ''"),
        ("dt04_urgencia_intervencion", "TEXT DEFAULT ''"),
        ("dt04_observaciones", "TEXT DEFAULT ''"),
        # F-DT-05
        ("dt05_fuentes_agua_json", "TEXT DEFAULT ''"),
        ("dt05_zona_recarga", "TEXT DEFAULT ''"),
        ("dt05_humedad_persistente", "TEXT DEFAULT ''"),
        ("dt05_escorrentia_concentrada", "TEXT DEFAULT ''"),
        ("dt05_dist_captacion", "TEXT DEFAULT ''"),
        ("dt05_jass_captacion", "TEXT DEFAULT ''"),
        ("dt05_interferencia_riego", "TEXT DEFAULT ''"),
        ("dt05_sistema_riego_nombre", "TEXT DEFAULT ''"),
        ("dt05_modalidad_acceso", "TEXT DEFAULT ''"),
        ("dt05_via_principal", "TEXT DEFAULT ''"),
        ("dt05_tipo_via_final", "TEXT DEFAULT ''"),
        ("dt05_transitabilidad_seca", "TEXT DEFAULT ''"),
        ("dt05_transitabilidad_lluviosa", "TEXT DEFAULT ''"),
        ("dt05_tiempo_dist_capital", "TEXT DEFAULT ''"),
        ("dt05_tiempo_prov_capital", "TEXT DEFAULT ''"),
        ("dt05_senal_celular", "TEXT DEFAULT ''"),
        ("dt05_operador_celular", "TEXT DEFAULT ''"),
        ("dt05_alojamiento", "TEXT DEFAULT ''"),
        ("dt05_requiere_ronda", "TEXT DEFAULT ''"),
        ("dt05_contacto_ronda", "TEXT DEFAULT ''"),
        ("dt05_observaciones", "TEXT DEFAULT ''"),
    ]
    for col_name, col_type in nuevas_columnas_dt:
        cursor.execute(f"""
            DO $$ BEGIN
                ALTER TABLE diagnostico_territorial ADD COLUMN {col_name} {col_type};
            EXCEPTION WHEN duplicate_column THEN NULL;
            END $$
        """)

    conn.commit()
    conn.close()

# ── Bloques ───────────────────────────────────────────────────────────────

def insertar_bloque(codigo, tipo_intervencion, cuenca, distrito,
                    utm_este, utm_norte, utm_zona, area_hectareas, estado,
                    altitud=0, responsable="", microcuenca="", provincia=""):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO bloques (codigo, tipo_intervencion, cuenca, distrito,
                             utm_este, utm_norte, utm_zona, altitud,
                             area_hectareas, responsable, estado, microcuenca,
                             provincia, fecha_registro)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (codigo, tipo_intervencion, cuenca, distrito,
          utm_este, utm_norte, utm_zona, altitud,
          area_hectareas, responsable, estado, microcuenca, provincia,
          datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    bloque_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return bloque_id


def actualizar_bloque(bloque_id, codigo, tipo_intervencion, cuenca, distrito,
                      utm_este, utm_norte, utm_zona, area_hectareas, estado,
                      altitud=0, responsable="", microcuenca="", provincia=""):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE bloques SET codigo=?, tipo_intervencion=?, cuenca=?, distrito=?,
                           utm_este=?, utm_norte=?, utm_zona=?, altitud=?,
                           area_hectareas=?, responsable=?, estado=?, microcuenca=?,
                           provincia=?
        WHERE id=?
    """, (codigo, tipo_intervencion, cuenca, distrito,
          utm_este, utm_norte, utm_zona, altitud,
          area_hectareas, responsable, estado, microcuenca, provincia,
          bloque_id))
    conn.commit()
    conn.close()


def bloques_con_coordenadas_faltantes(bloques):
    """Bloques que ya existen en la BD con coordenadas en 0 y para los que el
    catalogo si tiene centroide. Devuelve [(codigo, utm_este, utm_norte), ...]."""
    en_bd = {b["codigo"]: b for b in obtener_bloques()}
    pendientes = []
    for b in bloques or []:
        actual = en_bd.get(b.get("codigo"))
        if not actual:
            continue
        try:
            este_cat = float(b.get("utm_este") or 0)
            norte_cat = float(b.get("utm_norte") or 0)
            este_bd = float(actual.get("utm_este") or 0)
            norte_bd = float(actual.get("utm_norte") or 0)
        except (TypeError, ValueError):
            continue
        if este_cat and norte_cat and not (este_bd and norte_bd):
            pendientes.append((b["codigo"], este_cat, norte_cat))
    return pendientes


def sincronizar_bloques_catalogo(bloques, cuenca="Cuenca Alta del Rio Piura",
                                 tipo_intervencion="Restauracion", utm_zona="17S",
                                 completar_coordenadas=True):
    """Alta ADITIVA de bloques del catalogo que aun no existen en la BD.

    `bloques` es una lista de dicts con las claves: codigo, microcuenca,
    area_ha, provincia, distrito, utm_este, utm_norte.

    NO borra ni modifica ningun registro existente: los codigos ya presentes
    se omiten (el UNIQUE de `bloques.codigo` es la salvaguarda final). Toda la
    operacion corre en una sola transaccion, de modo que un fallo parcial deja
    la base intacta.

    Con `completar_coordenadas` (por defecto), ademas RELLENA el centroide de
    los bloques que ya existen con UTM en 0 y para los que el catalogo si trae
    coordenadas. Solo escribe sobre ceros: jamas pisa una coordenada cargada.

    Devuelve {"insertados": [...], "existentes": [...], "coords_actualizadas": [...]}.
    """
    codigos_bd = {b["codigo"] for b in obtener_bloques()}
    faltantes = [b for b in (bloques or []) if b.get("codigo") not in codigos_bd]
    existentes = [b.get("codigo") for b in (bloques or []) if b.get("codigo") in codigos_bd]
    pendientes_coord = (bloques_con_coordenadas_faltantes(bloques)
                        if completar_coordenadas else [])

    if not faltantes and not pendientes_coord:
        return {"insertados": [], "existentes": existentes, "coords_actualizadas": []}

    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_connection()
    insertados = []
    coords_actualizadas = []
    try:
        cursor = conn.cursor()
        for codigo_p, este_p, norte_p in pendientes_coord:
            # El WHERE repite la condicion de "coordenada vacia" para que la
            # actualizacion sea inofensiva aunque otro proceso las haya
            # cargado entre la lectura y esta escritura.
            cursor.execute("""
                UPDATE bloques SET utm_este=?, utm_norte=?
                WHERE codigo=?
                  AND (utm_este IS NULL OR utm_este=0)
                  AND (utm_norte IS NULL OR utm_norte=0)
            """, (float(este_p), float(norte_p), codigo_p))
            coords_actualizadas.append(codigo_p)
        for b in faltantes:
            cursor.execute("""
                INSERT INTO bloques (codigo, tipo_intervencion, cuenca, distrito,
                                     utm_este, utm_norte, utm_zona, altitud,
                                     area_hectareas, responsable, estado,
                                     microcuenca, provincia, fecha_registro)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT (codigo) DO NOTHING
            """, (
                b.get("codigo", ""), tipo_intervencion, cuenca, b.get("distrito", ""),
                float(b.get("utm_este") or 0.0), float(b.get("utm_norte") or 0.0),
                utm_zona, 0.0, float(b.get("area_ha") or 0.0), "", "Pendiente",
                b.get("microcuenca", ""), b.get("provincia", ""), fecha,
            ))
            # lastrowid queda en None si ON CONFLICT descarto la fila (otro
            # proceso la inserto entre la lectura y este INSERT).
            if cursor.lastrowid is not None:
                insertados.append(b.get("codigo", ""))
            else:
                existentes.append(b.get("codigo", ""))
        conn.commit()
    except Exception:
        try:
            conn._conn.rollback()
        except Exception:
            pass
        raise
    finally:
        conn.close()

    return {"insertados": insertados, "existentes": existentes,
            "coords_actualizadas": coords_actualizadas}


def eliminar_bloque(bloque_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM bloques WHERE id=?", (bloque_id,))
        conn.commit()
    except Exception:
        conn._conn.rollback()
        raise
    finally:
        conn.close()


# Tablas que componen un respaldo completo del aplicativo, en orden de lectura.
TABLAS_RESPALDO = [
    "bloques", "inspecciones", "indicadores_calidad", "diagnostico_territorial",
    "diagnostico_social", "elementos_expuestos", "presupuesto", "cronograma",
    "personal",
]


def respaldo_completo(tablas=None):
    """Lectura integra de las tablas del aplicativo para armar un respaldo.

    Solo lee (SELECT *): no modifica nada. Una tabla que no exista se omite en
    silencio para que el respaldo no falle por un esquema mas antiguo.

    Devuelve un dict {nombre_tabla: [dict por fila]}.
    """
    datos = {}
    conn = get_connection()
    try:
        for tabla in (tablas or TABLAS_RESPALDO):
            cursor = conn.cursor()
            try:
                cursor.execute(f"SELECT * FROM {tabla}")
                datos[tabla] = [dict(r) for r in _dictfetch(cursor)]
            except Exception:
                # Tabla inexistente: la transaccion queda abortada en
                # PostgreSQL, hay que revertirla antes de seguir leyendo.
                try:
                    conn._conn.rollback()
                except Exception:
                    pass
    finally:
        conn.close()
    return datos


def contar_registros_vinculados(bloque_id):
    """Cuenta los registros que se borrarian en cascada al eliminar un bloque.

    Devuelve un dict {nombre_tabla: n} solo con las tablas que tienen registros.
    """
    conteos = {}
    conn = get_connection()
    try:
        for tabla in ("inspecciones", "indicadores_calidad", "diagnostico_territorial",
                      "diagnostico_social", "elementos_expuestos", "presupuesto",
                      "cronograma"):
            cursor = conn.cursor()
            try:
                cursor.execute(f"SELECT COUNT(*) AS n FROM {tabla} WHERE bloque_id=?",
                               (bloque_id,))
                fila = cursor.fetchone()
                n = 0
                if fila is not None:
                    n = list(fila.values())[0] if hasattr(fila, "values") else fila[0]
                if n:
                    conteos[tabla] = int(n)
            except Exception:
                try:
                    conn._conn.rollback()
                except Exception:
                    pass
    finally:
        conn.close()
    return conteos


def obtener_bloques():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM bloques ORDER BY codigo")
    rows = _dictfetch(cursor)
    conn.close()
    return rows


def obtener_bloque_por_id(bloque_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM bloques WHERE id=?", (bloque_id,))
    row = _dictfetchone(cursor)
    conn.close()
    return row


def obtener_bloque_por_codigo(codigo):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM bloques WHERE codigo=?", (codigo,))
    row = _dictfetchone(cursor)
    conn.close()
    return row


# ── Inspecciones ──────────────────────────────────────────────────────────

def insertar_inspeccion(bloque_id, fecha_visita, inspector,
                        condiciones_climaticas, avance_fisico,
                        observaciones, desviaciones,
                        registro_fotografico, codigo_verificacion,
                        microcuenca="", archivos_pdf=""):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO inspecciones (bloque_id, fecha_visita, inspector,
                                  condiciones_climaticas, avance_fisico,
                                  observaciones, desviaciones,
                                  registro_fotografico, codigo_verificacion,
                                  microcuenca, archivos_pdf, fecha_registro)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
    """, (bloque_id, fecha_visita, inspector,
          condiciones_climaticas, avance_fisico,
          observaciones, desviaciones,
          registro_fotografico, codigo_verificacion,
          microcuenca, archivos_pdf,
          datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    inspeccion_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return inspeccion_id


def obtener_inspecciones_por_bloque(bloque_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT * FROM inspecciones WHERE bloque_id=? ORDER BY fecha_visita DESC
    """, (bloque_id,))
    rows = _dictfetch(cursor)
    conn.close()
    return rows


def obtener_todas_inspecciones():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT i.*, b.codigo AS bloque_codigo, b.tipo_intervencion,
               b.cuenca, b.distrito, b.utm_este, b.utm_norte, b.utm_zona,
               b.area_hectareas, b.estado AS bloque_estado
        FROM inspecciones i
        JOIN bloques b ON i.bloque_id = b.id
        ORDER BY i.fecha_visita DESC
    """)
    rows = _dictfetch(cursor)
    conn.close()
    return rows


def obtener_inspeccion_por_id(inspeccion_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT i.*, b.codigo AS bloque_codigo, b.tipo_intervencion,
               b.cuenca, b.distrito, b.utm_este, b.utm_norte, b.utm_zona,
               b.area_hectareas, b.estado AS bloque_estado
        FROM inspecciones i
        JOIN bloques b ON i.bloque_id = b.id
        WHERE i.id=?
    """, (inspeccion_id,))
    row = _dictfetchone(cursor)
    conn.close()
    return row


def actualizar_archivos_pdf_inspeccion(inspeccion_id, archivos_pdf):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE inspecciones SET archivos_pdf=? WHERE id=?",
                   (archivos_pdf, inspeccion_id))
    conn.commit()
    conn.close()


def actualizar_inspeccion(inspeccion_id, fecha_visita, inspector,
                          condiciones_climaticas, avance_fisico,
                          observaciones, desviaciones,
                          codigo_verificacion, microcuenca=""):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE inspecciones SET fecha_visita=?, inspector=?,
               condiciones_climaticas=?, avance_fisico=?,
               observaciones=?, desviaciones=?,
               codigo_verificacion=?, microcuenca=?
        WHERE id=?
    """, (fecha_visita, inspector, condiciones_climaticas, avance_fisico,
          observaciones, desviaciones, codigo_verificacion, microcuenca,
          inspeccion_id))
    conn.commit()
    conn.close()


def eliminar_inspeccion(inspeccion_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM inspecciones WHERE id=?", (inspeccion_id,))
    conn.commit()
    conn.close()


# ── Indicadores de Calidad ────────────────────────────────────────────────

def insertar_indicadores(bloque_id, inspeccion_id, cobertura_vegetal_planificada,
                         cobertura_vegetal_lograda, sobrevivencia_especies,
                         longitud_zanjas_ejecutada, volumen_retencion_sedimentos,
                         porcentaje_cobertura_vegetal=0,
                         tipo_cobertura_vegetal="",
                         vigor_cobertura_vegetal="",
                         microcuenca=""):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO indicadores_calidad (bloque_id, inspeccion_id,
                                         cobertura_vegetal_planificada, cobertura_vegetal_lograda,
                                         sobrevivencia_especies, longitud_zanjas_ejecutada,
                                         volumen_retencion_sedimentos,
                                         porcentaje_cobertura_vegetal,
                                         tipo_cobertura_vegetal, vigor_cobertura_vegetal,
                                         microcuenca, fecha_registro)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
    """, (bloque_id, inspeccion_id, cobertura_vegetal_planificada, cobertura_vegetal_lograda,
          sobrevivencia_especies, longitud_zanjas_ejecutada, volumen_retencion_sedimentos,
          porcentaje_cobertura_vegetal, tipo_cobertura_vegetal, vigor_cobertura_vegetal,
          microcuenca, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    indicador_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return indicador_id


def actualizar_indicadores(indicador_id, porcentaje_cobertura_vegetal=0,
                           tipo_cobertura_vegetal="", vigor_cobertura_vegetal="",
                           microcuenca=""):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE indicadores_calidad SET porcentaje_cobertura_vegetal=?,
               tipo_cobertura_vegetal=?, vigor_cobertura_vegetal=?, microcuenca=?
        WHERE id=?
    """, (porcentaje_cobertura_vegetal, tipo_cobertura_vegetal,
          vigor_cobertura_vegetal, microcuenca, indicador_id))
    conn.commit()
    conn.close()


def eliminar_indicadores(indicador_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM indicadores_calidad WHERE id=?", (indicador_id,))
    conn.commit()
    conn.close()


def obtener_indicadores_por_bloque(bloque_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT ic.*, i.fecha_visita, i.inspector
        FROM indicadores_calidad ic
        JOIN inspecciones i ON ic.inspeccion_id = i.id
        WHERE ic.bloque_id=?
        ORDER BY i.fecha_visita DESC
    """, (bloque_id,))
    rows = _dictfetch(cursor)
    conn.close()
    return rows


def obtener_indicadores_por_inspeccion(inspeccion_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM indicadores_calidad WHERE inspeccion_id=?", (inspeccion_id,))
    row = _dictfetchone(cursor)
    conn.close()
    return row


def obtener_resumen_bloques():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT b.*,
               (SELECT COUNT(*) FROM inspecciones WHERE bloque_id = b.id) AS total_inspecciones,
               (SELECT MAX(fecha_visita) FROM inspecciones WHERE bloque_id = b.id) AS ultima_visita,
               (SELECT avance_fisico FROM inspecciones WHERE bloque_id = b.id
                ORDER BY fecha_visita DESC LIMIT 1) AS ultimo_avance,
               (SELECT sobrevivencia_especies FROM indicadores_calidad WHERE bloque_id = b.id
                ORDER BY fecha_registro DESC LIMIT 1) AS ultima_sobrevivencia
        FROM bloques b
        ORDER BY b.codigo
    """)
    rows = _dictfetch(cursor)
    conn.close()
    return rows


def buscar_bloques(texto_busqueda):
    conn = get_connection()
    cursor = conn.cursor()
    patron = f"%{texto_busqueda}%"
    cursor.execute("""
        SELECT * FROM bloques
        WHERE codigo LIKE ? OR distrito LIKE ?
              OR tipo_intervencion LIKE ? OR responsable LIKE ?
        ORDER BY codigo
    """, (patron, patron, patron, patron))
    rows = _dictfetch(cursor)
    conn.close()
    return rows


# ── Presupuesto ───────────────────────────────────────────────────────────

def insertar_presupuesto(bloque_id, categoria, descripcion,
                         monto_planificado, monto_ejecutado, fuente_financiamiento):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO presupuesto (bloque_id, categoria, descripcion,
                                 monto_planificado, monto_ejecutado,
                                 fuente_financiamiento, fecha_registro)
        VALUES (?,?,?,?,?,?,?)
    """, (bloque_id, categoria, descripcion, monto_planificado, monto_ejecutado,
          fuente_financiamiento, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    pid = cursor.lastrowid
    conn.commit()
    conn.close()
    return pid


def actualizar_presupuesto(presupuesto_id, categoria, descripcion,
                           monto_planificado, monto_ejecutado, fuente_financiamiento):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE presupuesto SET categoria=?, descripcion=?,
                               monto_planificado=?, monto_ejecutado=?,
                               fuente_financiamiento=?
        WHERE id=?
    """, (categoria, descripcion, monto_planificado, monto_ejecutado,
          fuente_financiamiento, presupuesto_id))
    conn.commit()
    conn.close()


def eliminar_presupuesto(presupuesto_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM presupuesto WHERE id=?", (presupuesto_id,))
    conn.commit()
    conn.close()


def obtener_presupuesto_por_bloque(bloque_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM presupuesto WHERE bloque_id=? ORDER BY categoria", (bloque_id,))
    rows = _dictfetch(cursor)
    conn.close()
    return rows


def obtener_resumen_presupuesto():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT b.codigo, b.tipo_intervencion, b.distrito,
               COALESCE(SUM(p.monto_planificado), 0) AS total_planificado,
               COALESCE(SUM(p.monto_ejecutado), 0) AS total_ejecutado,
               COUNT(p.id) AS num_partidas
        FROM bloques b
        LEFT JOIN presupuesto p ON p.bloque_id = b.id
        GROUP BY b.id, b.codigo, b.tipo_intervencion, b.distrito
        ORDER BY b.codigo
    """)
    rows = _dictfetch(cursor)
    conn.close()
    return rows


def obtener_presupuesto_total():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT COALESCE(SUM(monto_planificado), 0) AS total_planificado,
               COALESCE(SUM(monto_ejecutado), 0) AS total_ejecutado
        FROM presupuesto
    """)
    row = _dictfetchone(cursor)
    conn.close()
    return row if row else {"total_planificado": 0, "total_ejecutado": 0}


# ── Cronograma ────────────────────────────────────────────────────────────

def insertar_actividad(bloque_id, actividad, fecha_inicio_plan, fecha_fin_plan,
                       fecha_inicio_real="", fecha_fin_real="",
                       porcentaje_avance=0, responsable="",
                       observaciones="", estado="Programado"):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO cronograma (bloque_id, actividad, fecha_inicio_plan,
                                fecha_fin_plan, fecha_inicio_real, fecha_fin_real,
                                porcentaje_avance, responsable, observaciones,
                                estado, fecha_registro)
        VALUES (?,?,?,?,?,?,?,?,?,?,?)
    """, (bloque_id, actividad, fecha_inicio_plan, fecha_fin_plan,
          fecha_inicio_real, fecha_fin_real, porcentaje_avance,
          responsable, observaciones, estado,
          datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    aid = cursor.lastrowid
    conn.commit()
    conn.close()
    return aid


def actualizar_actividad(actividad_id, actividad, fecha_inicio_plan,
                         fecha_fin_plan, fecha_inicio_real, fecha_fin_real,
                         porcentaje_avance, responsable, observaciones, estado):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE cronograma SET actividad=?, fecha_inicio_plan=?,
                              fecha_fin_plan=?, fecha_inicio_real=?,
                              fecha_fin_real=?, porcentaje_avance=?,
                              responsable=?, observaciones=?, estado=?
        WHERE id=?
    """, (actividad, fecha_inicio_plan, fecha_fin_plan,
          fecha_inicio_real, fecha_fin_real, porcentaje_avance,
          responsable, observaciones, estado, actividad_id))
    conn.commit()
    conn.close()


def eliminar_actividad(actividad_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM cronograma WHERE id=?", (actividad_id,))
    conn.commit()
    conn.close()


def obtener_actividades_por_bloque(bloque_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT * FROM cronograma WHERE bloque_id=? ORDER BY fecha_inicio_plan
    """, (bloque_id,))
    rows = _dictfetch(cursor)
    conn.close()
    return rows


def obtener_todas_actividades():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT c.*, b.codigo AS bloque_codigo
        FROM cronograma c
        JOIN bloques b ON c.bloque_id = b.id
        ORDER BY c.fecha_inicio_plan
    """)
    rows = _dictfetch(cursor)
    conn.close()
    return rows


def obtener_resumen_cronograma():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT estado, COUNT(*) AS cantidad FROM cronograma GROUP BY estado
    """)
    rows = _dictfetch(cursor)
    conn.close()
    return {r["estado"]: r["cantidad"] for r in rows}


# ── Personal ──────────────────────────────────────────────────────────────

def insertar_personal(nombre, cargo, especialidad="", telefono="",
                      email="", bloque_asignado=None):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO personal (nombre, cargo, especialidad, telefono,
                              email, bloque_asignado, activo, fecha_registro)
        VALUES (?,?,?,?,?,?,1,?)
    """, (nombre, cargo, especialidad, telefono, email,
          bloque_asignado, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    pid = cursor.lastrowid
    conn.commit()
    conn.close()
    return pid


def actualizar_personal(personal_id, nombre, cargo, especialidad,
                        telefono, email, bloque_asignado, activo):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE personal SET nombre=?, cargo=?, especialidad=?,
                            telefono=?, email=?, bloque_asignado=?, activo=?
        WHERE id=?
    """, (nombre, cargo, especialidad, telefono, email,
          bloque_asignado, activo, personal_id))
    conn.commit()
    conn.close()


def eliminar_personal(personal_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM personal WHERE id=?", (personal_id,))
    conn.commit()
    conn.close()


def obtener_todo_personal():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT p.*, b.codigo AS bloque_codigo
        FROM personal p
        LEFT JOIN bloques b ON p.bloque_asignado = b.id
        ORDER BY p.nombre
    """)
    rows = _dictfetch(cursor)
    conn.close()
    return rows


def obtener_personal_activo():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT p.*, b.codigo AS bloque_codigo
        FROM personal p
        LEFT JOIN bloques b ON p.bloque_asignado = b.id
        WHERE p.activo = 1
        ORDER BY p.nombre
    """)
    rows = _dictfetch(cursor)
    conn.close()
    return rows


# ── Dashboard ─────────────────────────────────────────────────────────────

def obtener_estadisticas_generales():
    conn = get_connection()
    cursor = conn.cursor()
    stats = {}

    cursor.execute("SELECT estado, COUNT(*) AS cantidad FROM bloques GROUP BY estado")
    stats["bloques_por_estado"] = {r["estado"]: r["cantidad"] for r in _dictfetch(cursor)}

    cursor.execute("SELECT tipo_intervencion, COUNT(*) AS cantidad FROM bloques GROUP BY tipo_intervencion")
    stats["bloques_por_tipo"] = {r["tipo_intervencion"]: r["cantidad"] for r in _dictfetch(cursor)}

    cursor.execute("SELECT COUNT(*) AS total FROM bloques")
    stats["total_bloques"] = _dictfetchone(cursor)["total"]

    cursor.execute("SELECT COALESCE(SUM(area_hectareas), 0) AS total FROM bloques")
    stats["area_total_ha"] = _dictfetchone(cursor)["total"]

    cursor.execute("SELECT COUNT(*) AS total FROM inspecciones")
    stats["total_inspecciones"] = _dictfetchone(cursor)["total"]

    cursor.execute("""
        SELECT COALESCE(AVG(sub.ultimo_avance), 0) AS promedio FROM (
            SELECT (SELECT avance_fisico FROM inspecciones
                    WHERE bloque_id = b.id ORDER BY fecha_visita DESC LIMIT 1)
                   AS ultimo_avance
            FROM bloques b
        ) sub WHERE sub.ultimo_avance IS NOT NULL
    """)
    stats["avance_promedio"] = _dictfetchone(cursor)["promedio"]

    cursor.execute("""
        SELECT COALESCE(SUM(monto_planificado), 0) AS planificado,
               COALESCE(SUM(monto_ejecutado), 0) AS ejecutado
        FROM presupuesto
    """)
    row = _dictfetchone(cursor)
    stats["presupuesto_planificado"] = row["planificado"]
    stats["presupuesto_ejecutado"] = row["ejecutado"]

    cursor.execute("SELECT estado, COUNT(*) AS cantidad FROM cronograma GROUP BY estado")
    stats["actividades_por_estado"] = {r["estado"]: r["cantidad"] for r in _dictfetch(cursor)}

    cursor.execute("SELECT COUNT(*) AS total FROM personal WHERE activo = 1")
    stats["personal_activo"] = _dictfetchone(cursor)["total"]

    cursor.execute("SELECT COUNT(*) AS total FROM diagnostico_territorial")
    stats["total_diagnosticos"] = _dictfetchone(cursor)["total"]

    cursor.execute("SELECT COUNT(*) AS total FROM diagnostico_social")
    stats["total_diagnosticos_sociales"] = _dictfetchone(cursor)["total"]

    conn.close()
    return stats


# ── Diagnostico Territorial ───────────────────────────────────────────────

def insertar_diagnostico_territorial(bloque_id, ficha, fecha_evaluacion, evaluador="",
                                     inspeccion_id=None, microcuenca="",
                                     forma_terreno="", pendiente="",
                                     posicion_fisiografica="", exposicion_orientacion="",
                                     paisaje_dominante="", rango_altitudinal="",
                                     precipitacion_anual="", temperatura_media="",
                                     humedad_relativa="", zona_vida="",
                                     presencia_heladas="", regimen_vientos="",
                                     textura_suelo="", color_suelo="",
                                     profundidad_efectiva="", pedregosidad="",
                                     drenaje="", presencia_erosion="",
                                     materia_organica="",
                                     tipo_cobertura="", densidad_cobertura="",
                                     estado_conservacion="", uso_actual_suelo="",
                                     conflicto_uso="",
                                     fuente_agua="", regimen_hidrico="",
                                     calidad_agua="", distancia_fuente_agua="",
                                     uso_recurso_hidrico="",
                                     tenencia_tierra="", organizacion_comunal="",
                                     actividad_economica="", accesibilidad_via="",
                                     distancia_centro_poblado="", servicios_basicos="",
                                     observaciones_generales=""):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO diagnostico_territorial (
            bloque_id, inspeccion_id, ficha, microcuenca, fecha_evaluacion, evaluador,
            forma_terreno, pendiente, posicion_fisiografica, exposicion_orientacion,
            paisaje_dominante, rango_altitudinal,
            precipitacion_anual, temperatura_media, humedad_relativa, zona_vida,
            presencia_heladas, regimen_vientos,
            textura_suelo, color_suelo, profundidad_efectiva, pedregosidad,
            drenaje, presencia_erosion, materia_organica,
            tipo_cobertura, densidad_cobertura, estado_conservacion,
            uso_actual_suelo, conflicto_uso,
            fuente_agua, regimen_hidrico, calidad_agua,
            distancia_fuente_agua, uso_recurso_hidrico,
            tenencia_tierra, organizacion_comunal, actividad_economica,
            accesibilidad_via, distancia_centro_poblado, servicios_basicos,
            observaciones_generales, fecha_registro
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,
                  ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (bloque_id, inspeccion_id, ficha, microcuenca, fecha_evaluacion, evaluador,
          forma_terreno, pendiente, posicion_fisiografica, exposicion_orientacion,
          paisaje_dominante, rango_altitudinal,
          precipitacion_anual, temperatura_media, humedad_relativa, zona_vida,
          presencia_heladas, regimen_vientos,
          textura_suelo, color_suelo, profundidad_efectiva, pedregosidad,
          drenaje, presencia_erosion, materia_organica,
          tipo_cobertura, densidad_cobertura, estado_conservacion,
          uso_actual_suelo, conflicto_uso,
          fuente_agua, regimen_hidrico, calidad_agua,
          distancia_fuente_agua, uso_recurso_hidrico,
          tenencia_tierra, organizacion_comunal, actividad_economica,
          accesibilidad_via, distancia_centro_poblado, servicios_basicos,
          observaciones_generales, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    did = cursor.lastrowid
    conn.commit()
    conn.close()
    return did


def obtener_diagnosticos_por_bloque(bloque_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT dt.*, b.codigo AS bloque_codigo
        FROM diagnostico_territorial dt
        JOIN bloques b ON dt.bloque_id = b.id
        WHERE dt.bloque_id=? ORDER BY dt.fecha_evaluacion DESC
    """, (bloque_id,))
    rows = _dictfetch(cursor)
    conn.close()
    return rows


def obtener_todos_diagnosticos():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT dt.*, b.codigo AS bloque_codigo, b.tipo_intervencion, b.distrito
        FROM diagnostico_territorial dt
        JOIN bloques b ON dt.bloque_id = b.id
        ORDER BY dt.fecha_evaluacion DESC
    """)
    rows = _dictfetch(cursor)
    conn.close()
    return rows


def obtener_diagnostico_por_id(diagnostico_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT dt.*, b.codigo AS bloque_codigo, b.tipo_intervencion, b.distrito
        FROM diagnostico_territorial dt
        JOIN bloques b ON dt.bloque_id = b.id
        WHERE dt.id=?
    """, (diagnostico_id,))
    row = _dictfetchone(cursor)
    conn.close()
    return row


def actualizar_diagnostico_territorial(diagnostico_id, ficha="", fecha_evaluacion="",
                                       evaluador="", microcuenca="",
                                       forma_terreno="", pendiente="",
                                       posicion_fisiografica="", exposicion_orientacion="",
                                       paisaje_dominante="", rango_altitudinal="",
                                       precipitacion_anual="", temperatura_media="",
                                       humedad_relativa="", zona_vida="",
                                       presencia_heladas="", regimen_vientos="",
                                       textura_suelo="", color_suelo="",
                                       profundidad_efectiva="", pedregosidad="",
                                       drenaje="", presencia_erosion="",
                                       materia_organica="",
                                       tipo_cobertura="", densidad_cobertura="",
                                       estado_conservacion="", uso_actual_suelo="",
                                       conflicto_uso="",
                                       fuente_agua="", regimen_hidrico="",
                                       calidad_agua="", distancia_fuente_agua="",
                                       uso_recurso_hidrico="",
                                       tenencia_tierra="", organizacion_comunal="",
                                       actividad_economica="", accesibilidad_via="",
                                       distancia_centro_poblado="", servicios_basicos="",
                                       observaciones_generales=""):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE diagnostico_territorial SET
            ficha=?, fecha_evaluacion=?, evaluador=?, microcuenca=?,
            forma_terreno=?, pendiente=?, posicion_fisiografica=?, exposicion_orientacion=?,
            paisaje_dominante=?, rango_altitudinal=?,
            precipitacion_anual=?, temperatura_media=?, humedad_relativa=?, zona_vida=?,
            presencia_heladas=?, regimen_vientos=?,
            textura_suelo=?, color_suelo=?, profundidad_efectiva=?, pedregosidad=?,
            drenaje=?, presencia_erosion=?, materia_organica=?,
            tipo_cobertura=?, densidad_cobertura=?, estado_conservacion=?,
            uso_actual_suelo=?, conflicto_uso=?,
            fuente_agua=?, regimen_hidrico=?, calidad_agua=?,
            distancia_fuente_agua=?, uso_recurso_hidrico=?,
            tenencia_tierra=?, organizacion_comunal=?, actividad_economica=?,
            accesibilidad_via=?, distancia_centro_poblado=?, servicios_basicos=?,
            observaciones_generales=?
        WHERE id=?
    """, (ficha, fecha_evaluacion, evaluador, microcuenca,
          forma_terreno, pendiente, posicion_fisiografica, exposicion_orientacion,
          paisaje_dominante, rango_altitudinal,
          precipitacion_anual, temperatura_media, humedad_relativa, zona_vida,
          presencia_heladas, regimen_vientos,
          textura_suelo, color_suelo, profundidad_efectiva, pedregosidad,
          drenaje, presencia_erosion, materia_organica,
          tipo_cobertura, densidad_cobertura, estado_conservacion,
          uso_actual_suelo, conflicto_uso,
          fuente_agua, regimen_hidrico, calidad_agua,
          distancia_fuente_agua, uso_recurso_hidrico,
          tenencia_tierra, organizacion_comunal, actividad_economica,
          accesibilidad_via, distancia_centro_poblado, servicios_basicos,
          observaciones_generales, diagnostico_id))
    conn.commit()
    conn.close()


# ── Diagnostico Territorial V5 (Plantilla DT V5: F-DT-01..05) ────────────
# Lista canonica de columnas escribibles por la UI V5. Todas son TEXT.
DT_V5_COLUMNS = [
    "ficha", "fecha_evaluacion", "evaluador", "microcuenca",
    # Datos generales V5
    "brigada", "ficha_correlativo", "altitud_gps",
    "centro_poblado_cercano", "comunidad_campesina_dt", "hora_registro",
    "utm_este_dt", "utm_norte_dt",
    # F-DT-01 (reusa columnas legacy + nuevas dt01_*)
    "forma_terreno", "pendiente", "posicion_fisiografica",
    "exposicion_orientacion", "rango_altitudinal", "paisaje_dominante",
    "dt01_afloramientos_rocosos", "dt01_escarpes_activos",
    "dt01_reptacion_suelo", "dt01_deslizamientos_antiguos",
    "dt01_remociones_masa_activas", "dt01_observaciones",
    # F-DT-02
    "dt02_sellamiento_costra", "dt02_compactacion_pisoteo",
    "dt02_raices_expuestas", "dt02_nivel_erosion_general",
    "dt02_carcavas_json", "dt02_nivel_erosion_sintesis",
    "dt02_num_carcavas", "dt02_longitud_total_carcavas",
    "dt02_pct_bloque_carcavas", "dt02_erosion_laminar_pct",
    "dt02_patron_carcavas", "dt02_socavamiento_cauce",
    "dt02_urgencia_control", "dt02_observaciones",
    # F-DT-03
    "dt03_parcela_muestreo", "dt03_dim_parcela", "dt03_pendiente_parcela",
    "dt03_cobertura_total", "dt03_tipo_ecosistema",
    "dt03_superficie_ecosistema", "dt03_estado_conservacion_eco",
    "dt03_uso_dominante", "dt03_cobertura_dosel", "dt03_cobertura_arbustiva",
    "dt03_cobertura_herbacea", "dt03_cobertura_hojarasca",
    "dt03_suelo_desnudo", "dt03_altura_estrato_dom", "dt03_altura_max",
    "dt03_dap_promedio", "dt03_regeneracion_natural",
    "dt03_estado_sanitario", "dt03_presencia_epifitas",
    "dt03_fenologia_dominante", "dt03_tipo_cobertura_dom",
    "dt03_floristica_json", "dt03_especies_clave_json", "dt03_observaciones",
    # F-DT-04
    "dt04_causas_json", "dt04_indicadores_json",
    "dt04_causas_directas_texto", "dt04_causa_subyacente",
    "dt04_velocidad_degradacion", "dt04_reversibilidad",
    "dt04_urgencia_intervencion", "dt04_observaciones",
    # F-DT-05
    "dt05_fuentes_agua_json", "dt05_zona_recarga",
    "dt05_humedad_persistente", "dt05_escorrentia_concentrada",
    "dt05_dist_captacion", "dt05_jass_captacion",
    "dt05_interferencia_riego", "dt05_sistema_riego_nombre",
    "dt05_modalidad_acceso", "dt05_via_principal", "dt05_tipo_via_final",
    "dt05_transitabilidad_seca", "dt05_transitabilidad_lluviosa",
    "dt05_tiempo_dist_capital", "dt05_tiempo_prov_capital",
    "dt05_senal_celular", "dt05_operador_celular", "dt05_alojamiento",
    "dt05_requiere_ronda", "dt05_contacto_ronda", "dt05_observaciones",
    # Observaciones generales (compartidas)
    "observaciones_generales",
]


def insertar_diagnostico_territorial_v5(bloque_id, data, inspeccion_id=None):
    """Inserta un diagnostico territorial V5. `data` es un dict; las claves
    no presentes se almacenan como cadena vacia. Retorna el id insertado."""
    cols = ["bloque_id", "inspeccion_id", "fecha_registro"] + DT_V5_COLUMNS
    placeholders = ",".join(["?"] * len(cols))
    cols_sql = ", ".join(cols)
    fecha_reg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    values = [bloque_id, inspeccion_id, fecha_reg] + [
        str(data.get(c, "") or "") for c in DT_V5_COLUMNS
    ]
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        f"INSERT INTO diagnostico_territorial ({cols_sql}) VALUES ({placeholders})",
        tuple(values),
    )
    did = cursor.lastrowid
    conn.commit()
    conn.close()
    return did


def actualizar_diagnostico_territorial_v5(diagnostico_id, data):
    """Actualiza un diagnostico territorial V5 a partir de un dict."""
    set_clause = ", ".join(f"{c}=?" for c in DT_V5_COLUMNS)
    values = [str(data.get(c, "") or "") for c in DT_V5_COLUMNS] + [diagnostico_id]
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        f"UPDATE diagnostico_territorial SET {set_clause} WHERE id=?",
        tuple(values),
    )
    conn.commit()
    conn.close()


def eliminar_diagnostico(diagnostico_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM diagnostico_territorial WHERE id=?", (diagnostico_id,))
    conn.commit()
    conn.close()


def obtener_resumen_diagnosticos():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT b.codigo, b.tipo_intervencion, b.distrito,
               COUNT(dt.id) AS total_fichas,
               GROUP_CONCAT(DISTINCT dt.ficha) AS fichas_completadas
        FROM bloques b
        LEFT JOIN diagnostico_territorial dt ON dt.bloque_id = b.id
        GROUP BY b.id, b.codigo, b.tipo_intervencion, b.distrito
        ORDER BY b.codigo
    """)
    rows = _dictfetch(cursor)
    conn.close()
    return rows


# ── Diagnostico Social ────────────────────────────────────────────────────

def insertar_diagnostico_social(datos):
    conn = get_connection()
    cursor = conn.cursor()
    datos["fecha_registro"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    columnas = list(datos.keys())
    placeholders = ",".join(["?"] * len(columnas))
    sql = f"INSERT INTO diagnostico_social ({','.join(columnas)}) VALUES ({placeholders})"
    cursor.execute(sql, list(datos.values()))
    did = cursor.lastrowid
    conn.commit()
    conn.close()
    return did


def actualizar_diagnostico_social(diagnostico_id, datos):
    conn = get_connection()
    cursor = conn.cursor()
    datos["fecha_registro"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sets = ", ".join([f"{k}=?" for k in datos.keys()])
    sql = f"UPDATE diagnostico_social SET {sets} WHERE id=?"
    cursor.execute(sql, list(datos.values()) + [diagnostico_id])
    conn.commit()
    conn.close()


def obtener_diagnosticos_sociales_por_bloque(bloque_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT ds.*, b.codigo AS bloque_codigo
        FROM diagnostico_social ds
        JOIN bloques b ON ds.bloque_id = b.id
        WHERE ds.bloque_id=? ORDER BY ds.fecha_evaluacion DESC
    """, (bloque_id,))
    rows = _dictfetch(cursor)
    conn.close()
    return rows


def obtener_todos_diagnosticos_sociales():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT ds.*, b.codigo AS bloque_codigo, b.tipo_intervencion, b.distrito
        FROM diagnostico_social ds
        JOIN bloques b ON ds.bloque_id = b.id
        ORDER BY ds.fecha_evaluacion DESC
    """)
    rows = _dictfetch(cursor)
    conn.close()
    return rows


def obtener_diagnostico_social_por_id(diagnostico_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT ds.*, b.codigo AS bloque_codigo, b.tipo_intervencion, b.distrito
        FROM diagnostico_social ds
        JOIN bloques b ON ds.bloque_id = b.id
        WHERE ds.id=?
    """, (diagnostico_id,))
    row = _dictfetchone(cursor)
    conn.close()
    return row


def eliminar_diagnostico_social(diagnostico_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM diagnostico_social WHERE id=?", (diagnostico_id,))
    conn.commit()
    conn.close()


def obtener_resumen_diagnosticos_sociales():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT b.codigo, b.tipo_intervencion, b.distrito,
               COUNT(ds.id) AS total_fichas,
               GROUP_CONCAT(DISTINCT ds.ficha) AS fichas_completadas
        FROM bloques b
        LEFT JOIN diagnostico_social ds ON ds.bloque_id = b.id
        GROUP BY b.id, b.codigo, b.tipo_intervencion, b.distrito
        ORDER BY b.codigo
    """)
    rows = _dictfetch(cursor)
    conn.close()
    return rows


# ── Elementos Expuestos (AdR / Riesgos) ──────────────────────────────────

def insertar_elementos_expuestos(datos):
    conn = get_connection()
    cursor = conn.cursor()
    datos["fecha_registro"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    columnas = list(datos.keys())
    placeholders = ",".join(["?"] * len(columnas))
    sql = f"INSERT INTO elementos_expuestos ({','.join(columnas)}) VALUES ({placeholders})"
    cursor.execute(sql, list(datos.values()))
    eid = cursor.lastrowid
    conn.commit()
    conn.close()
    return eid


def actualizar_elementos_expuestos(elemento_id, datos):
    conn = get_connection()
    cursor = conn.cursor()
    datos["fecha_registro"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sets = ", ".join([f"{k}=?" for k in datos.keys()])
    sql = f"UPDATE elementos_expuestos SET {sets} WHERE id=?"
    cursor.execute(sql, list(datos.values()) + [elemento_id])
    conn.commit()
    conn.close()


def obtener_elementos_expuestos_por_bloque(bloque_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT ee.*, b.codigo AS bloque_codigo
        FROM elementos_expuestos ee
        JOIN bloques b ON ee.bloque_id = b.id
        WHERE ee.bloque_id=? ORDER BY ee.fecha_campo DESC
    """, (bloque_id,))
    rows = _dictfetch(cursor)
    conn.close()
    return rows


def obtener_todos_elementos_expuestos():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT ee.*, b.codigo AS bloque_codigo, b.tipo_intervencion, b.distrito
        FROM elementos_expuestos ee
        JOIN bloques b ON ee.bloque_id = b.id
        ORDER BY ee.fecha_campo DESC
    """)
    rows = _dictfetch(cursor)
    conn.close()
    return rows


def obtener_elementos_expuestos_por_id(elemento_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT ee.*, b.codigo AS bloque_codigo, b.tipo_intervencion, b.distrito
        FROM elementos_expuestos ee
        JOIN bloques b ON ee.bloque_id = b.id
        WHERE ee.id=?
    """, (elemento_id,))
    row = _dictfetchone(cursor)
    conn.close()
    return row


def eliminar_elementos_expuestos(elemento_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM elementos_expuestos WHERE id=?", (elemento_id,))
    conn.commit()
    conn.close()


def obtener_resumen_elementos_expuestos():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT b.codigo, b.tipo_intervencion, b.distrito,
               COUNT(ee.id) AS total_fichas,
               GROUP_CONCAT(DISTINCT ee.ficha) AS fichas_completadas
        FROM bloques b
        LEFT JOIN elementos_expuestos ee ON ee.bloque_id = b.id
        GROUP BY b.id, b.codigo, b.tipo_intervencion, b.distrito
        ORDER BY b.codigo
    """)
    rows = _dictfetch(cursor)
    conn.close()
    return rows
