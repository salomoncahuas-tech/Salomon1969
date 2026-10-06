# Módulo «Liberación de Areas» – integrado al aplicativo IN Piura

Rama base: `claude/field-verification-desktop-app-mL0rL` · 28-sep-2026

## Qué cambia en el repo

| Archivo | Cambio |
|---|---|
| `streamlit_app.py` | +1 import, +1 opción de menú («Liberacion de Areas»), +1 ruta `elif`; catálogo vigente sin retirados |
| `database.py` | +12 códigos en `BLOQUES_RETIRADOS` (descartados en V6) |
| `requirements.txt` | + `shapely>=2.0.0` (pyproj ya estaba) |
| `liberacion_areas/` | Paquete nuevo: `la_core`, `la_db`, `la_kobo`, `la_export`, `pagina`, `la_schema.sql` |
| `datos/unidades_liberacion_areas.csv` | Catálogo V6: 117 bloques + 60 lotes SUS, con asistente (mismo CSV de los formularios Kobo) |
| `tests/test_liberacion_areas.py` | 8 pruebas (lógica + integración PostgreSQL con el DDL real de `database.py`) |
| `tests/casos_prueba_liberacion_areas.json` | Envíos ficticios F-LA-01/02/03/04/06 |

Ninguna tabla ni fila existente se borra; el único cambio sobre datos existentes es `activo = 0` en los 12 bloques retirados.

## Cómo se enlaza con lo existente

- **Conexión**: la misma `DATABASE_URL` de `st.secrets`, con `sslmode=require`, keepalives y reintentos como `database.py`. Una conexión por ejecución de la página.
- **Migración**: al abrir la página, `inicializar_la()` crea las tablas `la_*` si faltan (aditivo, idempotente), carga el catálogo V6 si `la_unidades` está vacía y enlaza con `bloques`.
- **Bloques**: `la_unidades.bloque_id → bloques(id)` para cada bloque V6 y `la_unidades.bloque_ref_id → bloques(id)` para cada lote SUS, por **igualdad exacta de código** (mismo criterio que `bloque_lookup.py`). Llaves `ON DELETE SET NULL`, como en el Paso 6.
- **Vista `la_v_conciliacion_bloques`**: compara el catálogo V6 con la tabla `bloques` (OK / AREA_DISTINTA / NO_EXISTE_EN_APP / RETIRADO_EN_APP / ACTIVO_EN_APP_FUERA_DE_V6). Se muestra en la pestaña 6.
- **Vista `la_v_antecedentes_paso6`**: último registro de `verificacion_campo_odk` por bloque (tenencia, n.° de predios, titular, aceptación, acta, dictamen) como consulta inicial a prediantes, visible en la matriz predial.
- **KoboToolbox**: reutiliza `odk_kobo.KoBoClient` (paginación completa, SSL, token solo a su dominio) y los secrets `KOBO_TOKEN` / `KOBO_SERVER` que ya usa la página «ODK / KoBoToolbox». Los formularios F-LA se listan desde la cuenta; no hace falta copiar UID.
- **Fotos y actas**: versión mediana en `la_adjuntos` (BYTEA), igual que `adjuntos_odk`; queda la URL del original.

## Bloques retirados (catálogo vigente V6)

Los 117 bloques V6 coinciden en código, distrito y área con `BLOQUES_V5` de `streamlit_app.py`. Los 12 bloques del catálogo del aplicativo que ya no están en V6 (`1, 7, 25, 29, 32, 33, 46, 48, 68, 74, 75, M18B5`) se agregaron a `BLOQUES_RETIRADOS` en `database.py` (confirmado el 28-sep-2026):

- En el siguiente arranque, `inicializar_bd()` los marca con `activo = 0`: dejan de listarse en todo el aplicativo, pero sus filas y todo lo registrado para ellos (inspecciones, diagnósticos, etc.) quedan intactos. Es reversible: basta quitarlos de la tupla y poner `activo = 1`.
- `streamlit_app.py` separa `BLOQUES_V5_COMPLETO` (todos, para resolver microcuenca y zona de registros ya ingresados) de `BLOQUES_V5` (vigentes: 117), que alimenta desplegables, plantillas y la sincronización. Así un bloque retirado no se vuelve a ofrecer ni aparece como «faltante».

## Pruebas

```bash
python -m pytest tests/test_liberacion_areas.py -q
LA_TEST_DB="postgresql://…/base_vacia_de_pruebas" python -m pytest tests/test_liberacion_areas.py -q
```

Resultado: 8/8. La suite completa queda igual que antes del cambio (las 15 fallas previas de `test_fichas_dt.py` y `test_resumenes_bloques.py` se deben a que los 117 libros Excel no vienen en el ZIP de la rama).

---

## Actualización 30-sep-2026 · Tres vías de registro, plantillas, edición y reportes

### Tres vías de registro equivalentes

| Vía | Dónde | Cómo entra al sistema |
|---|---|---|
| 1. KoboToolbox | Pestaña **3 · Importar KoboToolbox** (API o exportación XLSX/JSON) | Sin cambios |
| 2. Digitación en el aplicativo | Pestaña **1 · Registro en campo** | Formulario construido con los mismos campos; coordenadas en UTM 17S |
| 3. Plantilla Excel ANIN | Pestaña **2 · Plantillas Excel / Kobo** (descargar → llenar → importar) | Una hoja por formulario + hojas de detalle |

Las tres usan un único diccionario de campos (`liberacion_areas/la_campos.py`) y terminan en el mismo envío
«aplanado» que produce Kobo, que pasa por `la_core.validar_envio` y `la_db.importar`: mismas reglas (catálogo V6,
asistente asignado, rango UTM 17S, lotes SUS, checklist de actas) y mismos estados NUEVO / DUPLICADO / OBSERVADO.
El origen de cada envío queda en `la_envios_raw.origen` (KOBO / APP / PLANTILLA).

- **Plantilla Excel** (`la_plantillas.generar_plantilla_excel`): F-LA-01 a F-LA-06, `F-LA-03_puntos`,
  `F-LA-03_vertices_SUS`, `F-LA-04_firmantes`, catálogo de unidades e instrucciones; listas desplegables, validación
  de rangos UTM, fechas y DNI. Puede prepararse por asistente (el desplegable muestra solo sus unidades).
  Re-importar el mismo archivo no duplica: cada fila tiene un identificador determinista (`xls-…`).
- **XLSForm para Kobo** (`la_plantillas.paquete_kobo`): ZIP con F-LA-01/02/03/04/06 y `unidades.csv`, generado del
  mismo diccionario (validado con pyxform). Úselo si aún no están desplegados o si cambió el catálogo; si ya tiene
  formularios F-LA en la cuenta, compare los nombres de campo antes de reemplazarlos.
- **F-LA-05** (búsqueda documental) se registra en el aplicativo o en la plantilla (no tiene formulario Kobo).

### Historial, edición, eliminación y restauración (pestaña 4)

- Lista por formulario con filtros (unidad, asistente, estado, origen), **Editar** / **Eliminar** con confirmación,
  detalle con fotos y exportación a Excel ANIN — igual que en Diagnóstico Territorial y Social.
- **Editar** carga el registro (de cualquier vía) en «Registro en campo»; al guardar se re-valida, se conserva el
  mismo identificador y sus fotos, y se recalculan los estados LA.
- **Eliminar** guarda antes una copia completa en la tabla nueva `la_bitacora`; las fotos no se borran. Los
  registros eliminados (envíos, documentos, predios) se pueden **Restaurar**. La bitácora se puede exportar.
- **Matriz predial**: ahora también se editan nombre, área dentro de la unidad, **núcleo**, alertas; y se puede
  eliminar un predio sin registros asociados. Cada cambio queda en la bitácora.

### Reportes (pestaña 7)

- Excel consolidado ANIN (resumen, matriz, semáforo, avance, cada formulario, documentos, titulares **sin DNI ni
  celular**, observados), filtrable por asistente o distrito.
- Expediente preliminar PDF por unidad y ZIP con los expedientes de todas las unidades con predios.

### Migración

Solo aditiva (se aplica sola al abrir la página): columnas `origen`, `editado`, `editado_por` en `la_envios_raw` y
`editado`, `editado_por` en `la_documentos`; tabla nueva `la_bitacora`. Ningún registro existente se borra ni se
modifica (salvo rellenar `origen` de los envíos ya importados).

### Pruebas

```bash
python -m pytest tests/test_liberacion_areas.py tests/test_liberacion_areas_registro.py -q
LA_TEST_DB="postgresql://…/base_vacia_de_pruebas" python -m pytest tests -q
```

---

## Actualización 06-oct-2026 · Los envíos de KoboToolbox no aparecían en la página

### Causa

La página no consultaba KoboToolbox por sí sola: un envío F-LA quedaba solo en Kobo hasta que alguien lo
importaba a mano en la pestaña **3 · Importar KoboToolbox** (listar → descargar → elegir el formulario del
aplicativo → escribir el usuario → confirmar). Mientras tanto no aparecía en Historial, matriz predial ni reportes.
Además, el formulario del aplicativo se elegía aparte del formulario Kobo, y el botón «Confirmar importación»
quedaba deshabilitado sin explicación clara si faltaba el usuario responsable.

### Corrección

- **Aviso al abrir la página** (si hay `KOBO_TOKEN`): revisa los formularios F-LA de la cuenta (cada 5 minutos por
  sesión) y muestra «N envío(s) de KoboToolbox aún no importados (F-LA-01: N…)» con el botón **Importar ahora**.
  Importa con las mismas validaciones y estados (NUEVO / OBSERVADO) y descarga fotos y actas; un adjunto que falla
  ya no detiene la importación.
- **Formulario reconocido automáticamente** por el nombre en Kobo («F-LA-01», «F-LA01», «F_LA_01», «FLA 01»…) o, en
  exportaciones, por sus campos. La importación manual de la pestaña 3 sigue disponible.
- Los envíos **eliminados en el aplicativo** (bitácora) no se vuelven a traer en la sincronización.
- F-LA-01 de formularios sin el cálculo `cod_unidad` toma la unidad del campo «unidades».

Archivos: `la_kobo.py` (`form_id_por_nombre`, `form_id_por_campos`, `revisar_envios`, `sincronizar_envios`,
`guardar_adjuntos`), `la_db.uuids_eliminados`, `la_core.validar_envio`, `pagina.py`,
`tests/test_liberacion_areas_kobo_sync.py` (6 pruebas). Sin migración de base de datos.

### «No se pudo revisar KoboToolbox: Token inválido o sin permiso sobre el formulario»

KoboToolbox respondió HTTP 401/403 al token. Las causas habituales son: la cuenta está en otro servidor (cada token
vale en UN solo servidor: `kf.kobotoolbox.org`, `eu.kobotoolbox.org` o `kobo.humanitarianresponse.info`) y esta
página usaba siempre `KOBO_SERVER` o, si faltaba, `kf.kobotoolbox.org`; o bien el `KOBO_TOKEN` de los secrets está
incompleto o fue regenerado en Kobo.

- Si el token es rechazado, la página **prueba los servidores oficiales** y, si es válido en otro, lo usa en la
  sesión y muestra la línea `KOBO_SERVER = "…"` que conviene agregar a los secrets.
- Si no es válido en ninguno, el aviso indica el servidor consultado, de dónde se tomó el token y su huella (longitud
  y últimos 4 caracteres, sin revelarlo) con los pasos para reemplazarlo.
- La pestaña 3 tiene un selector **Servidor KoboToolbox**; `KOBO_SERVER` se normaliza (acepta `eu.kobotoolbox.org`,
  barras finales o una URL copiada del navegador).
