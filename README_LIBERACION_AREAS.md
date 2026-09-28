# Módulo «Liberación de Areas» – integrado al aplicativo IN Piura

Rama base: `claude/field-verification-desktop-app-mL0rL` · 28-sep-2026

## Qué cambia en el repo

| Archivo | Cambio |
|---|---|
| `streamlit_app.py` | +1 import, +1 opción de menú («Liberacion de Areas»), +1 ruta `elif` |
| `requirements.txt` | + `shapely>=2.0.0` (pyproj ya estaba) |
| `liberacion_areas/` | Paquete nuevo: `la_core`, `la_db`, `la_kobo`, `la_export`, `pagina`, `la_schema.sql` |
| `datos/unidades_liberacion_areas.csv` | Catálogo V6: 117 bloques + 60 lotes SUS, con asistente (mismo CSV de los formularios Kobo) |
| `tests/test_liberacion_areas.py` | 7 pruebas (lógica + integración PostgreSQL con el DDL real de `database.py`) |
| `tests/casos_prueba_liberacion_areas.json` | Envíos ficticios F-LA-01/02/03/04/06 |

Ninguna tabla ni fila existente se modifica o borra.

## Cómo se enlaza con lo existente

- **Conexión**: la misma `DATABASE_URL` de `st.secrets`, con `sslmode=require`, keepalives y reintentos como `database.py`. Una conexión por ejecución de la página.
- **Migración**: al abrir la página, `inicializar_la()` crea las tablas `la_*` si faltan (aditivo, idempotente), carga el catálogo V6 si `la_unidades` está vacía y enlaza con `bloques`.
- **Bloques**: `la_unidades.bloque_id → bloques(id)` para cada bloque V6 y `la_unidades.bloque_ref_id → bloques(id)` para cada lote SUS, por **igualdad exacta de código** (mismo criterio que `bloque_lookup.py`). Llaves `ON DELETE SET NULL`, como en el Paso 6.
- **Vista `la_v_conciliacion_bloques`**: compara el catálogo V6 con la tabla `bloques` (OK / AREA_DISTINTA / NO_EXISTE_EN_APP / RETIRADO_EN_APP / ACTIVO_EN_APP_FUERA_DE_V6). Se muestra en la pestaña 6.
- **Vista `la_v_antecedentes_paso6`**: último registro de `verificacion_campo_odk` por bloque (tenencia, n.° de predios, titular, aceptación, acta, dictamen) como consulta inicial a prediantes, visible en la matriz predial.
- **KoboToolbox**: reutiliza `odk_kobo.KoBoClient` (paginación completa, SSL, token solo a su dominio) y los secrets `KOBO_TOKEN` / `KOBO_SERVER` que ya usa la página «ODK / KoBoToolbox». Los formularios F-LA se listan desde la cuenta; no hace falta copiar UID.
- **Fotos y actas**: versión mediana en `la_adjuntos` (BYTEA), igual que `adjuntos_odk`; queda la URL del original.

## Hallazgo del cruce con el catálogo del aplicativo

Los 117 bloques V6 existen en `BLOQUES_V5` de `streamlit_app.py` con el mismo código, distrito y área (0 diferencias). Pero el aplicativo lista **12 bloques que ya no están en V6** y no figuran en `BLOQUES_RETIRADOS`: `1, 7, 25, 29, 32, 33, 46, 48, 68, 74, 75, M18B5`. Si el equipo confirma que están descartados, se ocultan sin perder datos agregándolos a `BLOQUES_RETIRADOS` en `database.py` (no se hizo en este cambio).

## Pruebas

```bash
python -m pytest tests/test_liberacion_areas.py -q
LA_TEST_DB="postgresql://…/base_vacia_de_pruebas" python -m pytest tests/test_liberacion_areas.py -q
```

Resultado: 7/7. La suite completa queda igual que antes del cambio (las 15 fallas previas de `test_fichas_dt.py` y `test_resumenes_bloques.py` se deben a que los 117 libros Excel no vienen en el ZIP de la rama).
