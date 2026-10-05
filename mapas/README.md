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
