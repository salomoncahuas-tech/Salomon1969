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

Con el vector de pendientes del proyecto (Drive, carpeta «Pendientes vector», campo `DN` 1–7):

```bash
python suelos_zanjas/aptitud_zanjas.py --pendientes "datos/gis/pendientes/Pendientes vector.shp" \
    --leyenda suelos_zanjas/leyenda_pendientes.json --salida suelos_zanjas/salidas
```

`leyenda_pendientes.json` define el rango de cada DN: `{"1": [0, 5, "grados"], ..., "7": [35, null, "grados"]}`
(transcribir la imagen «PENDIENTES VALORES DE VECTOR.jpeg»). Cuando una clase cruza un límite de 10, 15, …, 40 %,
su superficie se reparte en proporción lineal a su amplitud.

## Estado

- **Pendiente provisional:** mientras no se incorpore la geometría de «Pendientes vector» (14.7 MB; el conector de
  Drive solo baja archivos de hasta 10 MB), el script usa la **clase de pendiente dominante de cada delineación del E5**
  (`PEND_SIMB`, D.S. N.° 005-2022-MIDAGRI, MDE de 12.5 m del proyecto).
- **Cobertura:** la intersección celda a celda MSAVI × pendiente requiere la geometría de «MSAVI vectorizado»
  (13.9 MB); por ahora el motor reparte R1 y R2 dentro del bloque en proporción a la superficie.
- **Costo unitario:** el CU de zanja del motor es el de la ficha ANIN (0.40 × 0.30 m); la sección de la cartilla
  (0.40 × 0.40 m) excava un tercio más por metro. Ajustar el APU en el expediente técnico.

## Datos

- `datos/gis/suelos_e5/UNIDADES_CARTOGRAFICAS_SUELO.*` — mapa de suelos del E5 (423 delineaciones, EPSG:32717).
- `datos/gis/suelos_e5/BLOQUES_INTERVENCION.dbf` — composición de suelos y CUM por bloque (118 bloques, geometría V7).
- `datos/gis/IN_Piura_area_influencia_bloques_v6.gpkg` (capa `bloques_v6`) — geometría de los 117 bloques V6.
