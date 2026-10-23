# Respuesta a la solicitud de información — Programa de Inversión del Río Piura

Avance del Perfil de Inversión «Infraestructura Natural» (Proyecto IN Piura, CUI 2669244) para su incorporación
en el Programa de Inversión del Río Piura. Atiende los seis puntos de la solicitud.

| Punto solicitado | Informe Word | Excel de anexos | Paquete SIG |
|---|---|---|---|
| 1. Delimitación del ámbito de estudio y del área de influencia | Cap. 4 | A1 (distrito, microcuenca, bloque) | capas 01, 02, 03, 08 |
| 2. Criterios de selección | Cap. 5 | A2 | — |
| 3. Población afectada | Cap. 6 | A3 (CCPP, distrito, avance social) | capa 04 + lámina de CCPP |
| 4. Unidad Productora | Cap. 7 | A4 (ecosistemas, MSAVI, cárcavas) | capas 05, 06, 07 |
| 5. Involucrados | Cap. 8 | A5 | — |
| 6. Información georreferenciada | Cap. 9 | A6 | SHP, KMZ, GPKG, PDF, LEEME |

## Archivos (`salidas/`)

- `IN_Piura_Informe_Avance_Perfil_Respuesta.docx` — informe técnico con formato ANIN.
- `IN_Piura_Anexos_Avance_Perfil.xlsx` — anexos A1 a A6; resúmenes y totales con fórmulas y celdas de control.
- `IN_Piura_Paquete_SIG_Avance_Perfil.zip` — 8 capas en SHP (con .prj y .cpg) y GPKG en UTM WGS 84 17S
  (EPSG:32717), 7 KMZ más uno integrado (WGS 84), 3 mapas A3 en PDF y `LEEME.txt` con el diccionario de campos.
- `IN_Piura_Mapa_Centros_Poblados_V6.pdf/.png` — lámina A3 de los centros poblados asociados.
- `cifras_informe.json` — cifras que usa el informe (las mismas que calculan las fórmulas del Excel).

## Regenerar

```bash
pip install geopandas pyogrio shapely pyproj matplotlib openpyxl
python entregables/respuesta_solicitud_rio_piura/generar_paquete_respuesta.py
NODE_PATH=$(npm root -g) node entregables/respuesta_solicitud_rio_piura/informe_respuesta.js
```

## Cifras principales (corte 23.10.2026)

- 117 bloques V6, 12,270.235 ha, en 26 microcuencas y 15 distritos; área de influencia aprobada de 17,139.655 ha
  (130 polígonos). Ámbito de los 15 distritos: 642,927.8 ha (INEI, referencial).
- 152 centros poblados asociados (únicos), 18,071 habitantes, 6,227 viviendas.
- Diagnóstico social: 400 fichas F-DS-01 frente a una muestra de 368 hogares; F-DS-03 completo en 144 de 171 registros.
- UP: bosque estacionalmente seco de colina y montaña en el 68.0 % de los bloques; 8,341.824 ha (68.0 %) con
  MSAVI 2024 ≤ 0.4976; 55 cárcavas caracterizadas (23.1 km).

## Pendientes marcados `[POR DEFINIR]`

Número y fecha del documento de solicitud y del informe; criterio de delimitación del área de influencia; capas y
resultados de peligro (PMM, EPH, PGI, PI); límites de las microcuencas; capas de exclusión e idoneidad; consolidado
de las fichas F-DT, F-EE, F-DS-02 a 07 y F-LA del aplicativo; centros poblados de los bloques 31, 50, M17B10, M30B5 y
M6B2-1; coordenadas de los centros poblados de los bloques 83 a 87; año censal de la capa INEI; acuerdos suscritos.
