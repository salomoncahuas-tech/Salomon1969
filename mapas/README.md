# Mapas — Proyecto IN Piura

## Área de influencia y bloques de intervención V6

`python mapas/mapa_area_influencia_bloques_v6.py` genera en `mapas/salidas/` cuatro láminas A3 (PNG 200 ppp y un PDF):

| Lámina | Contenido |
|---|---|
| 1 | Mapa general: bloques V6 + área de influencia aprobada sobre el ámbito de 15 distritos, ubicación regional y cuadro por distrito (bloques y AI en ha) |
| 2-4 | Detalle por provincia (Ayabaca, Huancabamba, Morropón) con el código de cada bloque |

**Datos** (`datos/gis/IN_Piura_area_influencia_bloques_v6.gpkg`, EPSG:32717):

| Capa | Origen | Registros | Área |
|---|---|---|---|
| `bloques_v6` | `Bloques V6.kml` | 117 | 12,270.235 ha |
| `area_influencia_v6` | `AI_aprobado_2.shp` (CODIGO = `AI_<bloque>`) | 130 polígonos, los 117 bloques | 17,139.655 ha (campo AREA_HA) |
| `ambito_distritos` | INEI, capa simplificada (referencial) | 15 distritos | ≈ 642,928 ha |
| `contexto_distritos` | INEI, capa simplificada | Piura y regiones vecinas | solo fondo |

- Se excluyeron los bloques **74 y 75** (retirados en V6, ver `README_LIBERACION_AREAS.md`) y sus áreas de influencia
  `AI_74` (30.073 ha) y `AI_75` (125.819 ha). Por eso el AI vigente es 17,139.655 ha y no 17,295.547 ha (total del shapefile).
- Cada polígono de AI colinda con su bloque sin superponerse. A cada AI se le asignan el distrito y la provincia de su bloque.
- El campo `AREA_HA` del AI suma 17,139.655 ha; el área recalculada de las geometrías (`AREA_GEOM_HA`) suma 17,132.532 ha.
  Los cuadros usan `AREA_HA`.

## Excel de bloques y AI por distrito

`python mapas/excel_bloques_ai_distrito.py` genera `mapas/salidas/IN_Piura_Bloques_AI_por_Distrito_V6.xlsx`:

| Hoja | Contenido |
|---|---|
| `Resumen_Distrito` | 15 distritos: N.° de bloques, ha de bloques, polígonos AI, ha de AI, bloques + AI, AI/bloques, % del total |
| `Resumen_Provincia` | Lo mismo por provincia |
| `Bloques_AI` | 117 bloques con su AI (SUMIF), punto interior en UTM 17S |
| `AI_Poligonos` | 130 polígonos AI con área de atributo y área geométrica |
| `Notas` | Fuentes, exclusiones (74, 75) y controles |

Los resúmenes son fórmulas (`COUNTIFS`/`SUMIFS`) sobre las hojas de detalle, con celdas de control (OK/REVISAR).

## Ecosistemas y bloques V6

`python mapas/mapa_ecosistemas_bloques_v6.py` cruza los 117 bloques V6 con una capa de ecosistemas
(Mapa Nacional de Ecosistemas del Perú, MINAM 2018, R.M. N.° 440-2018-MINAM, o un recorte al ámbito) y genera en `mapas/salidas/`:

| Salida | Contenido |
|---|---|
| `IN_Piura_Mapa_Ecosistemas_Bloques_V6.pdf` + PNG | Lámina 1: ecosistemas del ámbito de 15 distritos con los 117 bloques, ubicación y cuadro de ha por ecosistema. Láminas 2-4: detalle por provincia con el código de cada bloque |
| `IN_Piura_Ecosistemas_por_Bloque_V6.xlsx` | `Resumen_Ecosistema`, `Ecosistema_Provincia`, `Bloques` (ecosistema dominante), `Bloque_Ecosistema` (detalle) y `Notas`; porcentajes y totales con fórmulas, control OK/REVISAR contra 12,270.235 ha |
| `IN_Piura_ecosistemas_bloques_v6.gpkg` | Capas `ecosistemas_ambito` y `ecosistemas_bloques_v6` (EPSG:32717) |

**La capa de ecosistemas no está en el repositorio.** Cópiela en `datos/gis/` con un nombre que contenga
`ecosistemas` (`.zip` con el shapefile y su `.prj`, `.gpkg`, `.shp` o `.geojson`), o indíquela con `--ecosistemas`.
El campo del nombre (`ECOSISTEMA`, `NOMBRE`, ...) y el del símbolo (`SIMBOLO`, `SIMB_ECOS`, ...) se detectan solos;
si no, use `--campo` y `--campo-simbolo`. Se reproyecta a UTM 17S y se recorta al ámbito más los bloques.

- El área de cada tramo bloque × ecosistema se calcula sobre la geometría y se **prorratea al área oficial del
  catálogo**, así los resúmenes suman 12,270.235 ha.
- Lo que la capa no cubra dentro de un bloque aparece como «Sin información de ecosistema».
- El mapa del MINAM es de escala regional: el ecosistema de cada bloque se confirma en campo (F-DT-03).
