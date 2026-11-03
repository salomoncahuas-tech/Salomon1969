# Aptitud de áreas para zanjas de infiltración — Proyecto IN Piura

Cruce de los 117 bloques V6 con el mapa de suelos del **Estudio de Suelos E5 (informe final, GEOSIG)** y con una
capa de pendientes para decidir **dónde** se pueden proyectar zanjas de infiltración y **con qué espaciamiento**.
El resultado (`salidas/aptitud_zanjas_bloques.json`) alimenta el motor de costos v1.5 del artefacto
«Diagnóstico Integral IN Piura V3» (pestañas *Suelos E5*, *Metrados y costos* y *Supuestos del motor*).

## Criterios

| Criterio | Condición | Fuente |
|---|---|---|
| Pendiente | 10 a 40 % (5.71° a 21.80°; % = 100·tan °) | Cartilla «Zanjas de infiltración», MINAGRI – AGRORURAL (2014) |
| Suelo y textura | Textura franca, no suelta | Cartilla; E5, matriz suelo–intervención (Anexo 7) |
| Infiltración | Vaciar 40 cm en ≤ 72 h con la velocidad de diseño (½ de la básica) | E5, Cuadro 61; `Tasa_basica_por_unidad_de_suelo.xlsx` |
| Cobertura | Sin cobertura: tabla hasta 30 %; con cobertura: hasta 40 % | Cartilla, tabla de distanciamiento |

Solo **Canchaque** y **El Ingenio** (6.625 cm/h de diseño, 6 h de vaciado) cumplen la condición de infiltración;
Chillique (96 h), Morropón (229 h), Cruz de Piedra, Ramón Castilla y Huarmaca (242 h) no.

Espaciamiento (zanja 0.40 × 0.40 m, cartilla MIDAGRI), interpolado en el centro de cada clase de 5 %:

| Pendiente (%) | 10 | 15 | 20 | 25 | 30 | 35 | 40 |
|---|---|---|---|---|---|---|---|
| Sin cobertura (m) | 30 | 20 | 15 | 13 | 11 | — | — |
| Con cobertura (m) | 45 | 30 | 23 | 20 | 17 | 14 | 12 |

En el motor, R1 (MSAVI DN 1–2) se trata **sin cobertura** y R2 (DN 3) **con cobertura**.

## Uso

```bash
pip install geopandas pyogrio
python suelos_zanjas/aptitud_zanjas.py --salida suelos_zanjas/salidas
```

Con el vector de pendientes del proyecto (salida vigente):

```bash
python suelos_zanjas/aptitud_zanjas.py --pendientes datos/gis/pendientes/Pendientes_vector.zip \
    --leyenda suelos_zanjas/leyenda_pendientes.json --salida suelos_zanjas/salidas
```

`leyenda_pendientes.json` transcribe la imagen «PENDIENTES VALORES DE VECTOR.jpeg» (incluida en el .zip): DN 1 = 0–5°,
2 = 5–10°, 3 = 10–15°, 4 = 15–20°, 5 = 20–25°, 6 = 25–27°, 7 = > 27°. Equivalencias usadas: 10 % = 5.71°,
15 % = 8.53°, 20 % = 11.31°, 25 % = 14.04°, 30 % = 16.70°, 35 % = 19.29°, 40 % = 21.80°. Cuando una clase cruza uno
de esos límites, su superficie se reparte en proporción lineal a su amplitud en grados (p. ej., DN 2: 14 % bajo 10 %,
56 % en 10–15 % y 29 % en 15–17.6 %). Control: la pendiente media por bloque que resulta de las clases coincide con la
de «Rango de altitud y Pendiente Promedio por Bloques V6» (bloque 5: 18.6° frente a 18.57°; M9B1: 10.2° frente a 10.17°).

Sin `--pendientes`, el script usa la clase de pendiente dominante de cada delineación del E5 (`PEND_SIMB`).

## Estado

- **Pendiente:** vector del proyecto «Pendientes vector» (57,255 polígonos), en `datos/gis/pendientes/Pendientes_vector.zip`.
- **Resultado vigente (cartera de 117 bloques):** 5,992 ha de los bloques están entre 10 y 40 % de pendiente; 244 ha de
  ellas en suelos aptos. Con las celdas R1/R2 de bosque seco y andino y la regla de cobertura, el área apta para zanjas
  es 13.35 ha y el metrado 8.06 km (frente a 17,971.71 km en el Excel v1.4). Sin la restricción de suelo serían
  3,605.99 ha y 2,220.44 km.
- **Cobertura:** la intersección celda a celda MSAVI × pendiente requiere la geometría de «MSAVI vectorizado»
  (13.9 MB); por ahora el motor reparte R1 y R2 dentro del bloque en proporción a la superficie.
- **Costo unitario:** el CU de zanja del motor es el de la ficha ANIN (0.40 × 0.30 m); la sección de la cartilla
  (0.40 × 0.40 m) excava un tercio más por metro. Ajustar el APU en el expediente técnico.

## Datos

- `datos/gis/suelos_e5/UNIDADES_CARTOGRAFICAS_SUELO.*` — mapa de suelos del E5 (423 delineaciones, EPSG:32717).
- `datos/gis/suelos_e5/BLOQUES_INTERVENCION.dbf` — composición de suelos y CUM por bloque (118 bloques, geometría V7).
- `datos/gis/IN_Piura_area_influencia_bloques_v6.gpkg` (capa `bloques_v6`) — geometría de los 117 bloques V6.
