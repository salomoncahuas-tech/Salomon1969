"""
IN Piura - Plan de Verificación de Campo
Módulo de base de datos SQLite para registro y consulta de bloques,
inspecciones e indicadores de calidad.
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
            area_hectareas REAL NOT NULL,
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
            densidad_planificada REAL DEFAULT 0,
            densidad_lograda REAL DEFAULT 0,
            sobrevivencia_especies REAL DEFAULT 0,
            longitud_zanjas_ejecutada REAL DEFAULT 0,
            volumen_retencion_sedimentos REAL DEFAULT 0,
            fecha_registro TEXT NOT NULL,
            FOREIGN KEY (bloque_id) REFERENCES bloques(id) ON DELETE CASCADE,
            FOREIGN KEY (inspeccion_id) REFERENCES inspecciones(id) ON DELETE CASCADE
        )
    """)

    conn.commit()
    conn.close()


# ── Operaciones CRUD para Bloques ──────────────────────────────────────────

def insertar_bloque(codigo, tipo_intervencion, cuenca, distrito,
                    utm_este, utm_norte, utm_zona, area_hectareas, estado):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO bloques (codigo, tipo_intervencion, cuenca, distrito,
                             utm_este, utm_norte, utm_zona, area_hectareas,
                             estado, fecha_registro)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (codigo, tipo_intervencion, cuenca, distrito,
          utm_este, utm_norte, utm_zona, area_hectareas, estado,
          datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    bloque_id = cursor.lastrowid
    conn.close()
    return bloque_id


def actualizar_bloque(bloque_id, codigo, tipo_intervencion, cuenca, distrito,
                      utm_este, utm_norte, utm_zona, area_hectareas, estado):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE bloques SET codigo=?, tipo_intervencion=?, cuenca=?, distrito=?,
                           utm_este=?, utm_norte=?, utm_zona=?, area_hectareas=?,
                           estado=?
        WHERE id=?
    """, (codigo, tipo_intervencion, cuenca, distrito,
          utm_este, utm_norte, utm_zona, area_hectareas, estado, bloque_id))
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
                        registro_fotografico, codigo_verificacion):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO inspecciones (bloque_id, fecha_visita, inspector,
                                  condiciones_climaticas, avance_fisico,
                                  observaciones, desviaciones,
                                  registro_fotografico, codigo_verificacion,
                                  fecha_registro)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (bloque_id, fecha_visita, inspector,
          condiciones_climaticas, avance_fisico,
          observaciones, desviaciones,
          registro_fotografico, codigo_verificacion,
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

def insertar_indicadores(bloque_id, inspeccion_id, densidad_planificada,
                         densidad_lograda, sobrevivencia_especies,
                         longitud_zanjas_ejecutada,
                         volumen_retencion_sedimentos):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO indicadores_calidad (bloque_id, inspeccion_id,
                                         densidad_planificada, densidad_lograda,
                                         sobrevivencia_especies,
                                         longitud_zanjas_ejecutada,
                                         volumen_retencion_sedimentos,
                                         fecha_registro)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (bloque_id, inspeccion_id, densidad_planificada, densidad_lograda,
          sobrevivencia_especies, longitud_zanjas_ejecutada,
          volumen_retencion_sedimentos,
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
