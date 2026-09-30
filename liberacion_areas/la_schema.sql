-- =====================================================================
-- Proyecto IN Piura (CUI 2669244) – Módulo de Liberación de Áreas
-- Esquema PostgreSQL / Supabase · versión 2026-09-28 (enlazado a `bloques`)
-- Lo ejecuta la_db.inicializar_la() al abrir la página (idempotente), o a
-- mano en Supabase > SQL Editor.
-- Migración SOLO ADITIVA, igual que el Paso 6 ODK: crea tablas la_* nuevas y
-- no modifica ni borra ninguna tabla ni fila existente. Las llaves hacia
-- `bloques` usan ON DELETE SET NULL (nunca CASCADE).
-- Geometrías en EPSG:32717 guardadas como WKT (texto), sin PostGIS.
-- =====================================================================

-- 1. Catálogo de unidades (bloques V6, lotes SUS, viveros)
CREATE TABLE IF NOT EXISTS la_unidades (
    codigo            TEXT PRIMARY KEY,               -- '27', 'M17B1', 'SUS-058', 'VIV-01'
    label             TEXT,
    tipo_unidad       TEXT NOT NULL CHECK (tipo_unidad IN ('bloque','lote_sus','vivero')),
    provincia         TEXT NOT NULL,
    distrito          TEXT NOT NULL,
    bloque_ref        TEXT,
    area_ha           NUMERIC(12,3),
    area_bloque_ha    NUMERIC(12,3),
    pct_bloque        NUMERIC(6,2),
    posicion_sus      TEXT CHECK (posicion_sus IS NULL OR posicion_sus IN ('DENTRO','CONTIGUO')),
    idon_sus          TEXT,
    estado_sus        TEXT,                            -- COMPLETO / PARCIAL / SIN_CANDIDATO / SIN_SUS
    asistente         TEXT,                            -- AP-1 … AP-6 / PENDIENTE
    asistente_nombre  TEXT,
    geom_wkt          TEXT,                            -- EPSG:32717
    activo            BOOLEAN DEFAULT TRUE,
    bloque_id         INTEGER,                         -- bloques(id) del aplicativo (tipo 'bloque')
    bloque_ref_id     INTEGER,                         -- bloques(id) del bloque asociado (lote SUS)
    actualizado       TIMESTAMPTZ DEFAULT now()
);
-- Bases creadas con la versión anterior del módulo
ALTER TABLE la_unidades ADD COLUMN IF NOT EXISTS bloque_id INTEGER;
ALTER TABLE la_unidades ADD COLUMN IF NOT EXISTS bloque_ref_id INTEGER;

-- 2. Envíos crudos de Kobo (trazabilidad total; clave única kobo_uuid)
CREATE TABLE IF NOT EXISTS la_envios_raw (
    kobo_uuid      TEXT PRIMARY KEY,
    form_id        TEXT NOT NULL,
    cod_unidad     TEXT,
    cod_predio     TEXT,
    asistente      TEXT,
    fecha_envio    TIMESTAMPTZ,
    estado_import  TEXT NOT NULL,                      -- NUEVO / OBSERVADO
    motivos        TEXT,
    payload        JSONB NOT NULL,
    import_id      BIGINT,
    creado         TIMESTAMPTZ DEFAULT now()
);

-- 3. Titulares (código único T0000; un titular puede tener varios predios)
CREATE TABLE IF NOT EXISTS la_titulares (
    cod_titular      TEXT PRIMARY KEY,                -- T0001
    tipo_titularidad TEXT NOT NULL,                   -- comunal/privado/posesionario/sucesion/estatal/sin_titular
    nombre           TEXT NOT NULL,                   -- persona o comunidad / entidad
    dni_ruc          TEXT UNIQUE,                     -- DNI (8) o RUC (11); dato reservado
    celular          TEXT,                            -- dato reservado
    representante    TEXT,
    representante_dni TEXT,
    conyuge          TEXT,
    conyuge_dni      TEXT,
    partida          TEXT,
    observaciones    TEXT,
    creado           TIMESTAMPTZ DEFAULT now()
);

-- 4. Predios ({UNIDAD}-P{nn}) – matriz predial
CREATE TABLE IF NOT EXISTS la_predios (
    cod_predio        TEXT PRIMARY KEY,               -- 'M17B1-P03'
    cod_unidad        TEXT NOT NULL REFERENCES la_unidades(codigo),
    n_predio          INTEGER NOT NULL,
    nombre_predio     TEXT,
    area_decl_ha      NUMERIC(12,3),
    area_unidad_ha    NUMERIC(12,3),                  -- superficie del predio dentro de la unidad
    uso_actual        TEXT,
    ocupacion         TEXT,
    aceptacion        TEXT,
    estado_la         TEXT NOT NULL DEFAULT 'LA-0'
                      CHECK (estado_la IN ('LA-0','LA-1','LA-2','LA-3','LA-4','LA-5','LA-6','NEG','OBS','EXC')),
    clasificacion     TEXT,                           -- Disponibilidad preliminar / Requiere verificación adicional / …
    nucleo            BOOLEAN DEFAULT FALSE,          -- predio sobre área núcleo de la unidad
    docs_completos    BOOLEAN DEFAULT FALSE,          -- lo marca el Especialista Predial (LA-5)
    expediente_conforme BOOLEAN DEFAULT FALSE,        -- lo marca el Especialista Legal (LA-6)
    excepcion_manual  TEXT CHECK (excepcion_manual IS NULL OR excepcion_manual IN ('NEG','OBS','EXC')),
    alertas           TEXT,
    geom_wkt          TEXT,
    actualizado       TIMESTAMPTZ DEFAULT now(),
    UNIQUE (cod_unidad, n_predio)
);

CREATE TABLE IF NOT EXISTS la_predio_titular (
    cod_predio   TEXT REFERENCES la_predios(cod_predio) ON DELETE CASCADE,
    cod_titular  TEXT REFERENCES la_titulares(cod_titular),
    rol          TEXT DEFAULT 'titular',              -- titular / conyuge / copropietario / heredero / comunero
    PRIMARY KEY (cod_predio, cod_titular, rol)
);

-- 5. Formularios de campo
CREATE TABLE IF NOT EXISTS la_reuniones (             -- F-LA-01
    kobo_uuid TEXT PRIMARY KEY REFERENCES la_envios_raw(kobo_uuid),
    fecha DATE, asistente TEXT, distrito TEXT, centro_poblado TEXT, comunidad TEXT, tipo_evento TEXT,
    unidades TEXT, asist_hombres INT, asist_mujeres INT, titulares_presentes INT,
    aceptacion TEXT, alertas TEXT, acuerdos TEXT, a01_suscrita TEXT, a02_suscrita TEXT,
    este NUMERIC(12,2), norte NUMERIC(12,2)
);
CREATE TABLE IF NOT EXISTS la_fichas_titular (        -- F-LA-02
    kobo_uuid TEXT PRIMARY KEY REFERENCES la_envios_raw(kobo_uuid),
    fecha DATE, asistente TEXT, cod_unidad TEXT, cod_predio TEXT, cod_titular TEXT,
    consentimiento TEXT, tipo_titularidad TEXT, docs_exhibidos TEXT, aceptacion TEXT, estado_la_propuesto TEXT
);
CREATE TABLE IF NOT EXISTS la_inspecciones (          -- F-LA-03
    kobo_uuid TEXT PRIMARY KEY REFERENCES la_envios_raw(kobo_uuid),
    fecha DATE, asistente TEXT, cod_unidad TEXT, cod_predio TEXT,
    este NUMERIC(12,2), norte NUMERIC(12,2), precision_m NUMERIC(8,2),
    validacion_espacial TEXT, distancia_m NUMERIC(10,1), puntos_fuera INT,
    interferencias TEXT, conclusion_campo TEXT,
    area_sus_ha NUMERIC(12,3), posicion_sus_calc TEXT, distancia_bloque_m NUMERIC(10,1), geom_wkt TEXT
);
CREATE TABLE IF NOT EXISTS la_actas (                 -- F-LA-04
    kobo_uuid TEXT PRIMARY KEY REFERENCES la_envios_raw(kobo_uuid),
    cod_doc TEXT, tipo_acta TEXT, fecha DATE, asistente TEXT, cod_unidad TEXT, cod_predio TEXT,
    area_comprometida_ha NUMERIC(12,3), plazo TEXT, n_firmantes INT, fedatario_tipo TEXT,
    quorum_pct NUMERIC(5,1), checklist TEXT, estado_acta TEXT, conforme BOOLEAN,
    fecha_entrega_cd DATE, recibido_cd BOOLEAN DEFAULT FALSE      -- control documentario
);
CREATE TABLE IF NOT EXISTS la_vivero_alternativas (   -- F-LA-06
    kobo_uuid TEXT PRIMARY KEY REFERENCES la_envios_raw(kobo_uuid),
    cod_vivero TEXT, fecha DATE, distrito TEXT, alt_nombre TEXT, modalidad TEXT, tipo_titularidad TEXT,
    titular_nombre TEXT, area_ha NUMERIC(12,3), altitud NUMERIC(8,1), pendiente TEXT, agua_fuente TEXT,
    agua_caudal_ls NUMERIC(8,2), acceso_tipo TEXT, acceso_camion TEXT, energia TEXT, dist_bloques_km NUMERIC(8,2),
    inundabilidad TEXT, deslizamiento TEXT, disposicion TEXT, firmaria_a07 TEXT, puntaje NUMERIC(5,1),
    este NUMERIC(12,2), norte NUMERIC(12,2)
);

-- 6. F-LA-05 Constancias de búsqueda documental (se registran en el aplicativo)
CREATE TABLE IF NOT EXISTS la_documentos (
    id            BIGSERIAL PRIMARY KEY,
    cod_doc       TEXT UNIQUE,                         -- CBU-{UNIDAD}-P{nn}-{n}
    cod_unidad    TEXT REFERENCES la_unidades(codigo),
    cod_predio    TEXT,
    tipo          TEXT NOT NULL,                       -- constancia_busqueda / partida / titulo / vigencia_poder / dni / otro
    entidad       TEXT,                                -- SUNARP / SBN / DRA / COFOPRI / municipalidad / otra
    fecha         DATE,
    resultado     TEXT,                                -- positivo / negativo / en_tramite
    n_partida     TEXT,
    descripcion   TEXT,
    archivo_url   TEXT,
    registrado_por TEXT,
    creado        TIMESTAMPTZ DEFAULT now()
);

-- 7. Adjuntos de Kobo (fotos y páginas de actas). Igual que adjuntos_odk:
--    versión mediana (~640 px) en BYTEA; la URL del original queda registrada.
CREATE TABLE IF NOT EXISTS la_adjuntos (
    id             BIGSERIAL PRIMARY KEY,
    kobo_uuid      TEXT REFERENCES la_envios_raw(kobo_uuid) ON DELETE SET NULL,
    cod_predio     TEXT,
    campo          TEXT,
    nombre_archivo TEXT NOT NULL,                       -- {cod_predio}_{campo}_{n}.jpg
    mimetype       TEXT,
    url_original   TEXT,
    contenido      BYTEA,
    tamano_bytes   INTEGER DEFAULT 0,
    creado         TIMESTAMPTZ DEFAULT now(),
    UNIQUE (kobo_uuid, nombre_archivo)
);

-- 8. Bitácora de importaciones
CREATE TABLE IF NOT EXISTS la_import_log (
    id          BIGSERIAL PRIMARY KEY,
    fecha       TIMESTAMPTZ DEFAULT now(),
    fuente      TEXT,                                   -- API / XLSX / JSON
    form_id     TEXT,
    leidos      INT, nuevos INT, duplicados INT, observados INT,
    usuario     TEXT
);

CREATE INDEX IF NOT EXISTS ix_predios_unidad ON la_predios(cod_unidad);
CREATE INDEX IF NOT EXISTS ix_raw_form ON la_envios_raw(form_id);
CREATE INDEX IF NOT EXISTS ix_actas_unidad ON la_actas(cod_unidad);

-- 9. Vista de avance por asistente
CREATE OR REPLACE VIEW la_v_avance_asistente AS
SELECT u.asistente,
       COUNT(DISTINCT u.codigo)                                        AS unidades,
       COUNT(p.cod_predio)                                             AS predios,
       COUNT(p.cod_predio) FILTER (WHERE p.estado_la IN ('LA-4','LA-5','LA-6')) AS predios_con_acta,
       COUNT(p.cod_predio) FILTER (WHERE p.estado_la = 'OBS')          AS observados,
       COUNT(p.cod_predio) FILTER (WHERE p.estado_la = 'NEG')          AS negativas
FROM la_unidades u LEFT JOIN la_predios p ON p.cod_unidad = u.codigo
WHERE u.activo
GROUP BY u.asistente;

-- 10. Seguridad (Ley 29733): activar RLS; el aplicativo se conecta con un rol
--     de servicio. Nunca exponer estas tablas con la clave anónima.
ALTER TABLE la_titulares   ENABLE ROW LEVEL SECURITY;
ALTER TABLE la_envios_raw  ENABLE ROW LEVEL SECURITY;
ALTER TABLE la_fichas_titular ENABLE ROW LEVEL SECURITY;
ALTER TABLE la_adjuntos    ENABLE ROW LEVEL SECURITY;
-- (Sin políticas para 'anon' = acceso denegado por defecto. El usuario postgres
--  del pooler, que usa el aplicativo, no está sujeto a RLS.)

-- 10b. Registro en el aplicativo / plantilla Excel, edición y eliminación con trazabilidad (versión 2026-10)
--     Solo columnas y tablas nuevas. Nada se borra al migrar.
ALTER TABLE la_envios_raw ADD COLUMN IF NOT EXISTS origen TEXT;          -- KOBO / APP / PLANTILLA
ALTER TABLE la_envios_raw ADD COLUMN IF NOT EXISTS editado TIMESTAMPTZ;
ALTER TABLE la_envios_raw ADD COLUMN IF NOT EXISTS editado_por TEXT;
ALTER TABLE la_documentos ADD COLUMN IF NOT EXISTS editado TIMESTAMPTZ;
ALTER TABLE la_documentos ADD COLUMN IF NOT EXISTS editado_por TEXT;
UPDATE la_envios_raw SET origen = CASE WHEN kobo_uuid LIKE 'app-%' THEN 'APP'
                                       WHEN kobo_uuid LIKE 'xls-%' THEN 'PLANTILLA' ELSE 'KOBO' END
 WHERE origen IS NULL;

-- Bitácora: copia completa del registro ANTES de editarlo o eliminarlo (permite restaurar).
CREATE TABLE IF NOT EXISTS la_bitacora (
    id          BIGSERIAL PRIMARY KEY,
    fecha       TIMESTAMPTZ DEFAULT now(),
    usuario     TEXT,
    accion      TEXT NOT NULL,                          -- EDITAR / ELIMINAR / RESTAURAR
    tabla       TEXT NOT NULL,                          -- la_envios_raw / la_documentos / la_predios
    clave       TEXT NOT NULL,                          -- kobo_uuid / cod_doc / cod_predio
    form_id     TEXT,
    datos       JSONB,
    restaurado  BOOLEAN DEFAULT FALSE
);
CREATE INDEX IF NOT EXISTS ix_bitacora_clave ON la_bitacora(tabla, clave);
ALTER TABLE la_bitacora ENABLE ROW LEVEL SECURITY;

-- 11. OPCIONAL con PostGIS:
-- CREATE EXTENSION IF NOT EXISTS postgis;
-- ALTER TABLE la_unidades ADD COLUMN IF NOT EXISTS geom geometry(MultiPolygon, 32717);
-- UPDATE la_unidades SET geom = ST_Multi(ST_GeomFromText(geom_wkt, 32717)) WHERE geom_wkt IS NOT NULL;

-- 12. Enlace con las tablas existentes del aplicativo (solo si existen).
--     Se agrega sin tocar `bloques` ni `verificacion_campo_odk`.
DO $$ BEGIN
  IF to_regclass('public.bloques') IS NOT NULL THEN
    IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'fk_la_unidades_bloque') THEN
      ALTER TABLE la_unidades ADD CONSTRAINT fk_la_unidades_bloque
        FOREIGN KEY (bloque_id) REFERENCES bloques(id) ON DELETE SET NULL;
    END IF;
    IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'fk_la_unidades_bloque_ref') THEN
      ALTER TABLE la_unidades ADD CONSTRAINT fk_la_unidades_bloque_ref
        FOREIGN KEY (bloque_ref_id) REFERENCES bloques(id) ON DELETE SET NULL;
    END IF;

    -- Conciliación catálogo V6 (la_unidades) ↔ tabla bloques del aplicativo
    EXECUTE $v$
      CREATE OR REPLACE VIEW la_v_conciliacion_bloques AS
      SELECT u.codigo, u.distrito AS distrito_v6, u.area_ha::double precision AS area_v6_ha,
             b.id AS bloque_id, b.distrito AS distrito_app, b.area_hectareas::double precision AS area_app_ha,
             COALESCE(b.activo, 1) AS activo_app,
             CASE WHEN b.id IS NULL THEN 'NO_EXISTE_EN_APP'
                  WHEN COALESCE(b.activo, 1) = 0 THEN 'RETIRADO_EN_APP'
                  WHEN abs(b.area_hectareas - u.area_ha::double precision) > 0.01 THEN 'AREA_DISTINTA'
                  ELSE 'OK' END AS estado
      FROM la_unidades u LEFT JOIN bloques b ON b.codigo = u.codigo
      WHERE u.tipo_unidad = 'bloque'
      UNION ALL
      SELECT b.codigo, NULL, NULL, b.id, b.distrito, b.area_hectareas::double precision,
             COALESCE(b.activo, 1), 'ACTIVO_EN_APP_FUERA_DE_V6'
      FROM bloques b
      WHERE COALESCE(b.activo, 1) = 1
        AND NOT EXISTS (SELECT 1 FROM la_unidades u WHERE u.tipo_unidad = 'bloque' AND u.codigo = b.codigo)
    $v$;
  END IF;

  -- Consulta inicial a prediantes registrada en la verificación Paso 6 (ODK)
  IF to_regclass('public.verificacion_campo_odk') IS NOT NULL THEN
    EXECUTE $v$
      CREATE OR REPLACE VIEW la_v_antecedentes_paso6 AS
      SELECT DISTINCT ON (v.codigo_bloque)
             v.codigo_bloque, v.fecha_visita, v.inspector, v.tenencia, v.n_predios,
             v.titular, v.aceptacion_titular, v.condiciones_aceptacion, v.acta_firmada,
             v.riesgo_social, v.dictamen
      FROM verificacion_campo_odk v
      ORDER BY v.codigo_bloque, v.fecha_visita DESC, v.id DESC
    $v$;
  END IF;
END $$;
