"""
IN Piura - Plan de Ingreso / Verificación de Campo
Módulo de base de datos SQLite para registro y consulta de bloques,
inspecciones, indicadores de calidad, presupuesto, cronograma y personal.
Cuenca alta del río Piura, Perú.
"""

import sqlite3
import os
from datetime import datetime

DB_NAME = "in_piura.db"


def get_connection(db_path=None):
    """Retorna una conexión a la base de datos SQLite."""
    if db_path is None:
        db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), DB_NAME)
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON")
    conn.row_factory = sqlite3.Row
    return conn


def inicializar_bd():
    """Crea las tablas si no existen."""
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bloques (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
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
            fecha_registro TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS inspecciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bloque_id INTEGER NOT NULL,
            fecha_visita TEXT NOT NULL,
            inspector TEXT NOT NULL,
            condiciones_climaticas TEXT,
            avance_fisico REAL DEFAULT 0,
            observaciones TEXT,
            desviaciones TEXT,
            registro_fotografico TEXT,
            codigo_verificacion TEXT,
            fecha_registro TEXT NOT NULL,
            FOREIGN KEY (bloque_id) REFERENCES bloques(id) ON DELETE CASCADE
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS indicadores_calidad (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bloque_id INTEGER NOT NULL,
            inspeccion_id INTEGER NOT NULL,
            cobertura_vegetal_planificada REAL DEFAULT 0,
            cobertura_vegetal_lograda REAL DEFAULT 0,
            sobrevivencia_especies REAL DEFAULT 0,
            longitud_zanjas_ejecutada REAL DEFAULT 0,
            volumen_retencion_sedimentos REAL DEFAULT 0,
            fecha_registro TEXT NOT NULL,
            FOREIGN KEY (bloque_id) REFERENCES bloques(id) ON DELETE CASCADE,
            FOREIGN KEY (inspeccion_id) REFERENCES inspecciones(id) ON DELETE CASCADE
        )
    """)

    # Tabla de presupuesto por bloque
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS presupuesto (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bloque_id INTEGER NOT NULL,
            categoria TEXT NOT NULL,
            descripcion TEXT DEFAULT '',
            monto_planificado REAL DEFAULT 0,
            monto_ejecutado REAL DEFAULT 0,
            fuente_financiamiento TEXT DEFAULT '',
            fecha_registro TEXT NOT NULL,
            FOREIGN KEY (bloque_id) REFERENCES bloques(id) ON DELETE CASCADE
        )
    """)

    # Tabla de cronograma / hitos
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cronograma (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bloque_id INTEGER NOT NULL,
            actividad TEXT NOT NULL,
            fecha_inicio_plan TEXT NOT NULL,
            fecha_fin_plan TEXT NOT NULL,
            fecha_inicio_real TEXT DEFAULT '',
            fecha_fin_real TEXT DEFAULT '',
            porcentaje_avance REAL DEFAULT 0,
            responsable TEXT DEFAULT '',
            observaciones TEXT DEFAULT '',
            estado TEXT NOT NULL DEFAULT 'Programado',
            fecha_registro TEXT NOT NULL,
            FOREIGN KEY (bloque_id) REFERENCES bloques(id) ON DELETE CASCADE
        )
    """)

    # Tabla de personal asignado al proyecto
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS personal (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            cargo TEXT NOT NULL,
            especialidad TEXT DEFAULT '',
            telefono TEXT DEFAULT '',
            email TEXT DEFAULT '',
            bloque_asignado INTEGER,
            activo INTEGER DEFAULT 1,
            fecha_registro TEXT NOT NULL,
            FOREIGN KEY (bloque_asignado) REFERENCES bloques(id) ON DELETE SET NULL
        )
    """)

    # Migrar columnas nuevas en bloques si la tabla ya existe
    _migrar_bloques(cursor)

    conn.commit()
    conn.close()


def _migrar_bloques(cursor):
    """Agrega columnas nuevas a la tabla bloques si no existen."""
    cursor.execute("PRAGMA table_info(bloques)")
    columnas = {col[1] for col in cursor.fetchall()}
    if "altitud" not in columnas:
        cursor.execute("ALTER TABLE bloques ADD COLUMN altitud REAL DEFAULT 0")
    if "responsable" not in columnas:
        cursor.execute("ALTER TABLE bloques ADD COLUMN responsable TEXT DEFAULT ''")
    if "microcuenca" not in columnas:
        cursor.execute("ALTER TABLE bloques ADD COLUMN microcuenca TEXT DEFAULT ''")
    if "provincia" not in columnas:
        cursor.execute("ALTER TABLE bloques ADD COLUMN provincia TEXT DEFAULT ''")

    # Migrar columnas en inspecciones
    cursor.execute("PRAGMA table_info(inspecciones)")
    cols_insp = {col[1] for col in cursor.fetchall()}
    if "microcuenca" not in cols_insp:
        cursor.execute("ALTER TABLE inspecciones ADD COLUMN microcuenca TEXT DEFAULT ''")

    # Migrar columnas en indicadores_calidad
    cursor.execute("PRAGMA table_info(indicadores_calidad)")
    cols_ind = {col[1] for col in cursor.fetchall()}
    if "porcentaje_cobertura_vegetal" not in cols_ind:
        cursor.execute("ALTER TABLE indicadores_calidad ADD COLUMN porcentaje_cobertura_vegetal REAL DEFAULT 0")
    if "tipo_cobertura_vegetal" not in cols_ind:
        cursor.execute("ALTER TABLE indicadores_calidad ADD COLUMN tipo_cobertura_vegetal TEXT DEFAULT ''")
    if "vigor_cobertura_vegetal" not in cols_ind:
        cursor.execute("ALTER TABLE indicadores_calidad ADD COLUMN vigor_cobertura_vegetal TEXT DEFAULT ''")
    if "microcuenca" not in cols_ind:
        cursor.execute("ALTER TABLE indicadores_calidad ADD COLUMN microcuenca TEXT DEFAULT ''")


# ── Operaciones CRUD para Bloques ──────────────────────────────────────────

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
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (codigo, tipo_intervencion, cuenca, distrito,
          utm_este, utm_norte, utm_zona, altitud,
          area_hectareas, responsable, estado, microcuenca, provincia,
          datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    bloque_id = cursor.lastrowid
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


def eliminar_bloque(bloque_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM bloques WHERE id=?", (bloque_id,))
    conn.commit()
    conn.close()


def obtener_bloques():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM bloques ORDER BY codigo")
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def obtener_bloque_por_id(bloque_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM bloques WHERE id=?", (bloque_id,))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def obtener_bloque_por_codigo(codigo):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM bloques WHERE codigo=?", (codigo,))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


# ── Operaciones CRUD para Inspecciones ─────────────────────────────────────

def insertar_inspeccion(bloque_id, fecha_visita, inspector,
                        condiciones_climaticas, avance_fisico,
                        observaciones, desviaciones,
                        registro_fotografico, codigo_verificacion,
                        microcuenca=""):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO inspecciones (bloque_id, fecha_visita, inspector,
                                  condiciones_climaticas, avance_fisico,
                                  observaciones, desviaciones,
                                  registro_fotografico, codigo_verificacion,
                                  microcuenca, fecha_registro)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (bloque_id, fecha_visita, inspector,
          condiciones_climaticas, avance_fisico,
          observaciones, desviaciones,
          registro_fotografico, codigo_verificacion,
          microcuenca,
          datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    inspeccion_id = cursor.lastrowid
    conn.close()
    return inspeccion_id


def obtener_inspecciones_por_bloque(bloque_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT * FROM inspecciones WHERE bloque_id=? ORDER BY fecha_visita DESC
    """, (bloque_id,))
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]


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
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]


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
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


# ── Operaciones CRUD para Indicadores de Calidad ───────────────────────────

def insertar_indicadores(bloque_id, inspeccion_id, cobertura_vegetal_planificada,
                         cobertura_vegetal_lograda, sobrevivencia_especies,
                         longitud_zanjas_ejecutada,
                         volumen_retencion_sedimentos,
                         porcentaje_cobertura_vegetal=0,
                         tipo_cobertura_vegetal="",
                         vigor_cobertura_vegetal="",
                         microcuenca=""):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO indicadores_calidad (bloque_id, inspeccion_id,
                                         cobertura_vegetal_planificada, cobertura_vegetal_lograda,
                                         sobrevivencia_especies,
                                         longitud_zanjas_ejecutada,
                                         volumen_retencion_sedimentos,
                                         porcentaje_cobertura_vegetal,
                                         tipo_cobertura_vegetal,
                                         vigor_cobertura_vegetal,
                                         microcuenca,
                                         fecha_registro)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (bloque_id, inspeccion_id, cobertura_vegetal_planificada, cobertura_vegetal_lograda,
          sobrevivencia_especies, longitud_zanjas_ejecutada,
          volumen_retencion_sedimentos,
          porcentaje_cobertura_vegetal,
          tipo_cobertura_vegetal,
          vigor_cobertura_vegetal,
          microcuenca,
          datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    indicador_id = cursor.lastrowid
    conn.close()
    return indicador_id


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
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def obtener_indicadores_por_inspeccion(inspeccion_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT * FROM indicadores_calidad WHERE inspeccion_id=?
    """, (inspeccion_id,))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def obtener_resumen_bloques():
    """Retorna un resumen de todos los bloques con su última inspección e indicadores."""
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
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def buscar_bloques(texto_busqueda):
    """Busca bloques por código, distrito o tipo de intervención."""
    conn = get_connection()
    cursor = conn.cursor()
    patron = f"%{texto_busqueda}%"
    cursor.execute("""
        SELECT * FROM bloques
        WHERE codigo LIKE ? OR distrito LIKE ? OR tipo_intervencion LIKE ?
              OR responsable LIKE ?
        ORDER BY codigo
    """, (patron, patron, patron, patron))
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]


# ── Operaciones CRUD para Presupuesto ─────────────────────────────────────

def insertar_presupuesto(bloque_id, categoria, descripcion,
                         monto_planificado, monto_ejecutado,
                         fuente_financiamiento):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO presupuesto (bloque_id, categoria, descripcion,
                                 monto_planificado, monto_ejecutado,
                                 fuente_financiamiento, fecha_registro)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (bloque_id, categoria, descripcion,
          monto_planificado, monto_ejecutado, fuente_financiamiento,
          datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    pid = cursor.lastrowid
    conn.close()
    return pid


def actualizar_presupuesto(presupuesto_id, categoria, descripcion,
                           monto_planificado, monto_ejecutado,
                           fuente_financiamiento):
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
    cursor.execute("""
        SELECT * FROM presupuesto WHERE bloque_id=? ORDER BY categoria
    """, (bloque_id,))
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def obtener_resumen_presupuesto():
    """Retorna un resumen del presupuesto agrupado por bloque."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT b.codigo, b.tipo_intervencion, b.distrito,
               COALESCE(SUM(p.monto_planificado), 0) AS total_planificado,
               COALESCE(SUM(p.monto_ejecutado), 0) AS total_ejecutado,
               COUNT(p.id) AS num_partidas
        FROM bloques b
        LEFT JOIN presupuesto p ON p.bloque_id = b.id
        GROUP BY b.id
        ORDER BY b.codigo
    """)
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def obtener_presupuesto_total():
    """Retorna totales globales del presupuesto."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT COALESCE(SUM(monto_planificado), 0) AS total_planificado,
               COALESCE(SUM(monto_ejecutado), 0) AS total_ejecutado
        FROM presupuesto
    """)
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else {"total_planificado": 0, "total_ejecutado": 0}


# ── Operaciones CRUD para Cronograma ──────────────────────────────────────

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
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (bloque_id, actividad, fecha_inicio_plan, fecha_fin_plan,
          fecha_inicio_real, fecha_fin_real, porcentaje_avance,
          responsable, observaciones, estado,
          datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    aid = cursor.lastrowid
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
        SELECT * FROM cronograma WHERE bloque_id=?
        ORDER BY fecha_inicio_plan
    """, (bloque_id,))
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def obtener_todas_actividades():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT c.*, b.codigo AS bloque_codigo
        FROM cronograma c
        JOIN bloques b ON c.bloque_id = b.id
        ORDER BY c.fecha_inicio_plan
    """)
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def obtener_resumen_cronograma():
    """Retorna estadísticas del cronograma."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT estado, COUNT(*) AS cantidad
        FROM cronograma
        GROUP BY estado
    """)
    rows = cursor.fetchall()
    conn.close()
    return {r["estado"]: r["cantidad"] for r in rows}


# ── Operaciones CRUD para Personal ────────────────────────────────────────

def insertar_personal(nombre, cargo, especialidad="", telefono="",
                      email="", bloque_asignado=None):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO personal (nombre, cargo, especialidad, telefono,
                              email, bloque_asignado, activo, fecha_registro)
        VALUES (?, ?, ?, ?, ?, ?, 1, ?)
    """, (nombre, cargo, especialidad, telefono, email,
          bloque_asignado, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    pid = cursor.lastrowid
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
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]


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
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]


# ── Estadísticas para Dashboard ───────────────────────────────────────────

def obtener_estadisticas_generales():
    """Retorna estadísticas generales del proyecto para el dashboard."""
    conn = get_connection()
    cursor = conn.cursor()

    stats = {}

    # Bloques por estado
    cursor.execute("""
        SELECT estado, COUNT(*) AS cantidad FROM bloques GROUP BY estado
    """)
    stats["bloques_por_estado"] = {r["estado"]: r["cantidad"] for r in cursor.fetchall()}

    # Bloques por tipo
    cursor.execute("""
        SELECT tipo_intervencion, COUNT(*) AS cantidad FROM bloques GROUP BY tipo_intervencion
    """)
    stats["bloques_por_tipo"] = {r["tipo_intervencion"]: r["cantidad"] for r in cursor.fetchall()}

    # Totales
    cursor.execute("SELECT COUNT(*) AS total FROM bloques")
    stats["total_bloques"] = cursor.fetchone()["total"]

    cursor.execute("SELECT COALESCE(SUM(area_hectareas), 0) AS total FROM bloques")
    stats["area_total_ha"] = cursor.fetchone()["total"]

    cursor.execute("SELECT COUNT(*) AS total FROM inspecciones")
    stats["total_inspecciones"] = cursor.fetchone()["total"]

    # Avance promedio
    cursor.execute("""
        SELECT COALESCE(AVG(sub.ultimo_avance), 0) AS promedio FROM (
            SELECT (SELECT avance_fisico FROM inspecciones
                    WHERE bloque_id = b.id ORDER BY fecha_visita DESC LIMIT 1)
                   AS ultimo_avance
            FROM bloques b
        ) sub WHERE sub.ultimo_avance IS NOT NULL
    """)
    stats["avance_promedio"] = cursor.fetchone()["promedio"]

    # Presupuesto
    cursor.execute("""
        SELECT COALESCE(SUM(monto_planificado), 0) AS planificado,
               COALESCE(SUM(monto_ejecutado), 0) AS ejecutado
        FROM presupuesto
    """)
    row = cursor.fetchone()
    stats["presupuesto_planificado"] = row["planificado"]
    stats["presupuesto_ejecutado"] = row["ejecutado"]

    # Cronograma
    cursor.execute("""
        SELECT estado, COUNT(*) AS cantidad FROM cronograma GROUP BY estado
    """)
    stats["actividades_por_estado"] = {r["estado"]: r["cantidad"] for r in cursor.fetchall()}

    # Personal activo
    cursor.execute("SELECT COUNT(*) AS total FROM personal WHERE activo = 1")
    stats["personal_activo"] = cursor.fetchone()["total"]

    conn.close()
    return stats
