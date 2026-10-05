# Mapas — Proyecto IN Piura

## Área de influencia y bloques de intervención V6

`python mapas/mapa_area_influencia_bloques_v6.py` genera en `mapas/salidas/` cuatro láminas A3 (PNG 200 ppp y un PDF):

| Lámina | Contenido |
|---|---|
| 1 | Mapa general: área de influencia (15 distritos) + 117 bloques V6, ubicación regional y cuadro por distrito |
| 2-4 | Detalle por provincia (Ayabaca, Huancabamba, Morropón) con el código de cada bloque |

**Datos** (`datos/gis/IN_Piura_area_influencia_bloques_v6.gpkg`, EPSG:32717):

- `bloques_v6`: 117 bloques V6 (12,270.235 ha). Proviene de `Bloques V6.kml`; se excluyeron los bloques **74 y 75**, retirados en V6 (ver `README_LIBERACION_AREAS.md`).
- `area_influencia_distritos`: los 15 distritos del ámbito, tomados como área de influencia (≈ 642,928 ha).
- `contexto_distritos`: distritos de Piura y regiones vecinas (solo para el fondo del mapa).

Los límites distritales son una capa INEI simplificada (referencial). Si el estudio define otra área de influencia
(microcuencas, buffer, etc.), reemplaza la capa `area_influencia_distritos` y vuelve a correr el script.
