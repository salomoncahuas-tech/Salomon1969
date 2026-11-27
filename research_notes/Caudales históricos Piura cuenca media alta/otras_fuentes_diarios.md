# Caudales diarios de la cuenca media y alta del río Piura en fuentes distintas del portal SENAMHI

Fecha de trabajo: 27/11/2026. Ámbito: cuenca media y alta del río Piura (Tambogrande, Puente Ñácara/Chulucanas, Malacasí, Paltashaco, Barrios, San Pedro y tributarios). Puente Sánchez Cerro aparece solo como referencia aguas abajo, y así se indica en todo el documento.

**Condiciones de acceso en esta sesión.** El proxy de red de la sesión bloqueó todas las descargas directas: snirh.ana.gob.pe, ana.gob.pe, repositorio.ana.gob.pe, portal.grdc.bafg.de, flood-api.open-meteo.com, iridl.senamhi.gob.pe, zenodo.org, chirapiura.gob.pe, pirhua.udep.edu.pe, cds/ewds.climate.copernicus.eu y datosabiertos.gob.pe. Las respuestas fueron `HTTP 403 (agent proxy: no rule or network setting allows this host)` en curl y `getaddrinfo ENOTFOUND` en WebFetch. Firecrawl tampoco pudo usarse porque la cuenta no tenía créditos. GitHub sí estaba accesible (git clone). Por eso, **todos los datos diarios recuperados vienen de una copia pública en GitHub de las series que publica el servicio web del SNIRH-ANA**, y no de una descarga directa desde el SNIRH.

**Archivos CSV guardados en esta carpeta.** Todos tienen las columnas `fecha,caudal_m3s,estacion,fuente,tipo` y `tipo = observado`. No hay ninguna serie simulada, porque no fue posible extraerlas (ver la pregunta 3).

| Archivo | Estación (código SNIRH / operador) | Río según SNIRH | Primera fecha | Última fecha | N.º de días con dato / días del periodo | % faltante | Vacíos principales |
|---|---|---|---|---|---|---|---|
| diario_tambogrande_ana-snirh.csv | Tambo Grande (1586 / op. 63) | "Río San Francisco" (Por verificar; debería ser el río Piura) | 1982-09-01 | 2023-07-18 | 566 / 14,931 | 96.2 % | 1987-02-01 a 2021-02-27 (12,446 días); 1983-01-27 a 1985-12-31 (1,070 días); 2021-06-10 a 2023-03-14 (643 días). Datos útiles: set. 1982–ene. 1983, ene. 1986–ene. 1987, mar.–jun. 2021 (dispersos) y mar.–jul. 2023 (solo 5 días) |
| diario_puente_nacara_auto9609_ana-snirh.csv | Puente Ñácara, automática (9609 / op. 72) | Río Piura | 2023-02-21 | 2026-11-25 | 803 / 1,374 | 41.6 % | 2024-05-16 a 2025-01-03 (233 días); 2025-07-02 a 2026-01-06 (189 días); 2026-08-04 a 2026-10-26 (84 días) |
| diario_puente_nacara_auto3554_ana-snirh.csv | Puente Ñácara, automática (3554 / op. 72) | Río Piura | 2023-02-21 | 2026-05-04 | 816 / 1,169 | 30.2 % | 2025-07-01 a 2025-11-17 (140 días); 2024-09-16 a 2024-12-27 (103 días); 2024-05-16 a 2024-07-01 (47 días) |
| diario_puente_nacara_conv1196_ana-snirh.csv | Puente Ñácara (Chulucanas), convencional (1196 / op. 72) | Río Piura | 1984-06-30 | 1991-08-31 | 521 / 2,619 | 80.1 % | 1986-12-01 a 1991-02-28 (1,551 días); 1985-01-01 a 1985-11-30 (334 días) |
| diario_puente_chulucanas_ana-snirh.csv | Puente Chulucanas (1154 / op. 72), cerrada | Río Piura | 1964-09-01 | 1969-08-31 | 1,826 / 1,826 | 0.0 % | Sin vacíos (1,384 días con 0 m³/s) |
| diario_la_greda_ana-snirh.csv | La Greda (714 / op. 72), cerrada | Río Piura | 1971-09-01 | 1975-12-31 | 1,583 / 1,583 | 0.0 % | Sin vacíos |
| diario_carrasquillo_ana-snirh.csv | Carrasquillo (223 / op. 72), cerrada | Río Piura | 1978-09-01 | 1982-12-31 | 1,583 / 1,583 | 0.0 % | Sin vacíos (796 días con 0) |
| diario_puente_paltashaco_ana-snirh.csv | Puente Paltashaco (1200 / op. 72), cerrada | "Río Los Gallegos" (Por verificar) | 1978-09-01 | 1982-12-31 | 1,583 / 1,583 | 0.0 % | Sin vacíos |
| diario_barrios_ana-snirh.csv | Barrios (111 / op. 63), cerrada | Río Bigote | 1988-09-01 | 1992-12-31 | 1,583 / 1,583 | 0.0 % | Sin vacíos |
| diario_hacienda_barrios_3557_ana-snirh.csv | Hacienda Barrios (3557 / op. 72) | s/d (cuenca Bigote) | 2023-01-01 | 2026-04-17 | 1,184 / 1,203 | 1.6 % | 2025-12-21 a 2026-01-05 (16 días) |
| diario_hacienda_barrios_9610_ana-snirh.csv | Hacienda Barrios (9610 / op. 72) | s/d | 2023-01-01 | 2026-11-25 | 1,369 / 1,425 | 3.9 % | 2026-04-19 a 2026-06-10 (53 días) |
| diario_san_pedro_9611_ana-snirh.csv | San Pedro (9611 / op. 72) | s/d | 2023-01-01 | 2026-11-25 | 1,416 / 1,425 | 0.6 % | Solo días sueltos en feb.–mar. 2025 |
| diario_san_pedro_3558_ana-snirh.csv | San Pedro (3558 / op. 72) | s/d | 2023-01-01 | 2026-05-04 | 1,146 / 1,220 | 6.1 % | 2025-02-12 a 2025-04-07 (55 días) |
| diario_san_pedro_1804_ana-snirh.csv | San Pedro (1804 / op. 72), cerrada | Río San Jorge | 2011-09-02 | 2017-09-30 | 337 / 2,221 | 84.8 % | 2014-08-16 a 2017-01-31 (900 días); 215 vacíos en total |
| diario_san_pedro_1442_ana-snirh.csv | San Pedro (1442 / op. 63), cerrada | "Río Yapatero" (coordenadas Por verificar) | 1988-09-01 | 1993-04-30 | 1,703 / 1,703 | 0.0 % | Sin vacíos |
| diario_chililique_ana-snirh.csv | Chililique (302 / op. 63), cerrada | Río Yapatera | 1987-09-01 | 1991-12-31 | 1,583 / 1,583 | 0.0 % | Sin vacíos |
| diario_corral_del_medio_ana-snirh.csv | Corral del Medio (406 / op. 72), cerrada | Río Corrales | 1925-09-01 | 1992-04-29 | 1,133 / 24,348 | 95.3 % | Datos completos de 1925-09-01 a 1928-08-31, más dic. 1942 y 6 días de abr. 1992 |
| diario_teodulo_pena_ana-snirh.csv | Teódulo Peña (1610 / op. 72), cerrada | Río Corrales | 1991-03-01 | 1995-04-30 | 366 / 1,522 | 76.0 % | 1992-09-27 a 1995-01-31 (857 días) |
| diario_puente_sanchez_cerro_REFERENCIA_ana-snirh.csv | **REFERENCIA AGUAS ABAJO (bajo Piura)**: Puente Sánchez Cerro (1212 / op. 63) | Río Piura | 2022-09-01 | 2026-11-25 | 1,496 / 1,547 | 3.3 % | 2025-11-01 a 2025-11-26 (26 días) |

Notas de control de calidad: ninguna serie tiene fechas duplicadas ni valores negativos. El % faltante se calculó entre la primera y la última fecha con dato. La tabla no incluye las estaciones del SNIRH en la cuenca Piura que figuran en el inventario **sin serie**: Malacasí (851, op. 63), Salitral (3556, 9590) y Corral del Medio automática (3559).

## 1. Fuentes que publican caudal diario observado descargable para estaciones de la cuenca media y alta (URL, códigos, periodos)

### Takeaway
La única fuente fuera del portal SENAMHI que **efectivamente** dio datos diarios observados fue el servicio web del **SNIRH-ANA** (visor por cuenca, método `CaudalSerie`), recuperado a través de una copia pública en GitHub con corte al 26/11/2026. Ese servicio entrega solo los **últimos ~5 años hidrológicos con datos** de cada estación, por lo que en las estaciones históricas se obtienen tramos de 4 a 5 años, no series largas. **CAMELS-PE** (Zenodo) sí contiene caudal diario observado 1981–2025 para Tambogrande, Puente Ñácara, San Pedro y Hacienda Barrios, pero no pudo descargarse porque Zenodo estaba bloqueado. GRDC no publica estaciones de Piura en abierto, o no se pudo confirmar que lo haga.

### Cited Findings
- **SNIRH-ANA, visor por cuenca.** La URL es `https://snirh.ana.gob.pe/visorPorCuenca/?IdVar=263` y tiene los servicios ASMX `ServicioGeneral.asmx/ListarMapaGEO` (inventario de estaciones, capa 263 = hidrométricas), `ServicioGeneral.asmx/ListarInfoGEO` (operador, cuenca, AAA, ALA) y `Principal.asmx/CaudalSerie` (serie de caudal; parámetros JSON `{"pIdEstacion": <código>, "pIdOperador": <operador>}`). — [cesarmiguel10/dashboard-bi-piura, README y etl/](https://github.com/cesarmiguel10/dashboard-bi-piura)
- `CaudalSerie` devuelve por estación una serie "Prom. Histórico" (climatología diaria por día del año hidrológico) y una serie por año hidrológico "AAAA-AAAA" (del 1 de setiembre al 31 de agosto), con "los últimos ~5 años hidrológicos" diarios. — [cesarmiguel10/dashboard-bi-piura, etl/snirh_caudales.py](https://github.com/cesarmiguel10/dashboard-bi-piura/blob/main/etl/snirh_caudales.py)
- La copia usada está en el commit `d02a6886c56a349da83d819e848efb8f8c72c014` del 26/11/2026 09:41 (-05:00), archivo `data/caudales_diarios.parquet`. El inventario de Piura tiene 63 estaciones. — [cesarmiguel10/dashboard-bi-piura, data/](https://github.com/cesarmiguel10/dashboard-bi-piura/tree/main/data)
- El operador 63 del SNIRH corresponde al **Proyecto Especial Chira Piura**. La estación Puente Sánchez Cerro aparece como `ESTACION, OPERADOR = 1212, 63` con el comentario "Puente Sánchez Cerro (201201), Proyecto Especial Chira Piura". El mismo script indica que el servicio del SNIRH "no responde a los servidores de GitHub desde el 17-sep-2026" y que en un año vacío devuelve el JSON malformado `"data":}`. — [carlosbr123/visor-bajo-piura, actualizar_caudal.py](https://github.com/carlosbr123/visor-bajo-piura/blob/main/actualizar_caudal.py)
- Códigos del SNIRH relevantes, con coordenadas del inventario (lat, lon):
  - Tambo Grande 1586: -4.95, -80.3167; op. 63; "Funcionando".
  - Puente Ñácara: 9609 (-5.1129, -80.1730), 3554 (-5.1128, -80.1730) y 1196 "Puente Ñacara (Chulucanas)" (-5.1097, -80.1697).
  - Puente Chulucanas 1154: -5.10, -80.1833.
  - Malacasí 851: -5.3167, -79.8833; op. 63; cerrada; **sin serie**.
  - Puente Paltashaco 1200: -5.1167, -79.90.
  - Barrios 111: -5.2833, -79.70.
  - Hacienda Barrios: 3557 y 9610 (-5.2895, -79.6991).
  - San Pedro: 1804, 3558 y 9611 (-5.0796, -80.0162), y 1442 (-5.1096, -80.1712).
  - Chililique 302: -5.0333, -80.0667.
  - Carrasquillo 223: -5.2167, -80.0167.
  - La Greda 714: -4.9667, -80.30.
  - Corral del Medio: 406 y 3559.
  - Teódulo Peña 1610.
  - Salitral: 3556 y 9590.

  — [cesarmiguel10/dashboard-bi-piura, data/estaciones.parquet](https://github.com/cesarmiguel10/dashboard-bi-piura/tree/main/data)
- Máximos diarios en las series recuperadas:
  - Puente Ñácara 9609: 1,105.89 m³/s el 14/03/2023.
  - Tambo Grande 1586: 909.50 m³/s el 15/03/2023; 483.90 m³/s el 26/01/1983.
  - Puente Chulucanas 1154: 1,423.58 m³/s el 29/03/1965.
  - Carrasquillo 223: 614.00 m³/s el 12/03/1981.
  - La Greda 714: 533.00 m³/s el 19/03/1972.
  - Corral del Medio 406: 675.00 m³/s el 09/03/1926.
  - Referencia Sánchez Cerro 1212: 1,561.54 m³/s el 15/03/2023.

  Cálculo propio sobre los datos de [cesarmiguel10/dashboard-bi-piura](https://github.com/cesarmiguel10/dashboard-bi-piura).
- **CAMELS-PE v1.0.1** (SENAMHI; Llauca, Montesinos-Cáceres, Gutiérrez-Reynaga y Lavado-Casimiro, 2026) tiene 136 cuencas con series diarias 1981–2025. Su variable `flow_obs` ("Observed streamflow") está en **mm/día** y `flow_sim` es PISCO-ARNOVIC v1.1. El DOI es 10.5281/zenodo.21195425 y la descarga es `https://zenodo.org/api/records/21195425/files/CAMELS-PE_v1.0.1.zip/content` (~120 MB). — [rcamelspe README y R/download.R](https://github.com/PaulESantos/rcamelspe); [ESSD preprint essd-2026-386](https://essd.copernicus.org/preprints/essd-2026-386/)
- Estaciones de CAMELS-PE en la cuenca Piura, según su gpkg de estaciones (gauge_id, coordenadas, COMID, área):
  - TAMBOGRANDE: PE_47E0522C; -4.938, -80.344; COMID 9042297; 6,392.09 km².
  - PUENTE NACARA: PE_47E0415A; -5.113, -80.173; COMID 9043133; 4,641.44 km².
  - SAN PEDRO: PE_47E08444; -5.080, -80.016; COMID 9043206; 157.58 km².
  - HACIENDA BARRIOS: PE_47E074C0; -5.289, -79.702; COMID 9044140; 423.87 km².
  - Referencia PUENTE SANCHEZ CERRO: PE_201201; COMID 9043566; 7,507.31 km².

  Los archivos por cuenca existen en el paquete completo, por ejemplo `03_timeseries/by_catchment/PE_47E0522C.csv` con 1,163,624 bytes. — [KURONO23/CAMELS-PE, shp/camels_pe_gauges.gpkg](https://github.com/KURONO23/CAMELS-PE); [rcamelspe, inst/benchmarks/results/data-manifest.csv](https://github.com/PaulESantos/rcamelspe/blob/main/inst/benchmarks/results/data-manifest.csv)
- CAMELS-PE retuvo solo cuencas con observaciones diarias en 1981–2025. Las observaciones vienen sobre todo de estaciones de SENAMHI, complementadas con estaciones de otras instituciones. — [ESSD preprint essd-2026-386](https://essd.copernicus.org/preprints/essd-2026-386/)
- **Datos Abiertos Perú.** Hay un conjunto del Gobierno Regional Piura / PECHP con `CAUDAL07H`, `PROMEDIO24H`, `MAXIMA24H` y `PRECIP24H` por estación (hidrológicas, hidrométricas y presas). Descarga: `https://www.datosabiertos.gob.pe/node/10105/download`. — [stephany-toribio/piura-hydrology-analytics, main.py y README](https://github.com/stephany-toribio/piura-hydrology-analytics). No pude descargarlo porque el proxy devolvió 403 para www.datosabiertos.gob.pe. Su periodo y estaciones quedan **Por verificar**.
- **GRDC.** Los datos diarios de GRDC figuran como "PRIVATE", con uso limitado por un acuerdo JRC–GRDC. — [JRC AquaKnow](https://aquaknow.jrc.ec.europa.eu/node/18911). No encontré ningún número GRDC de estaciones del río Piura en las búsquedas web ni en GitHub, y portal.grdc.bafg.de estaba bloqueado.
- Los reportes hidrológicos diarios de la DZ01 Piura de SENAMHI usan datos del Proyecto Especial Chira Piura. En 2026 reportan Puente Ñácara a las 08:00 y Tambogrande como "S/D" (sin dato). — [SENAMHI 03522SENA-641](https://www.senamhi.gob.pe/load/file/03522SENA-641.pdf). Es una fuente del portal SENAMHI y se cita solo como contexto sobre el origen PECHP.

### Inferences
- Lo que entrega el SNIRH depende de la estación. Para las estaciones cerradas (Chulucanas 1964–69, La Greda 1971–75, Carrasquillo y Paltashaco 1978–82, Barrios 1988–92, Chililique 1987–91) son los **últimos 4–5 años del registro**. Que el servicio también entregue una climatología diaria "Prom. Histórico" sugiere que la base de datos de ANA guarda registros más largos que el servicio no expone día a día. Esto es una inferencia y está **Por verificar**.
- Hay pares de códigos que parecen la misma estación con distinto procesamiento: 9609 y 3554 (Ñácara; r = 0.986, 171 de 677 días idénticos), 3558 y 9611 (San Pedro; r = 0.992) y 3557 y 9610 (Hacienda Barrios; r = 0.799, con medianas de 3.01 frente a 0.97 m³/s). Conviene preguntar a ANA o PECHP cuál es la serie validada antes de usarlas (**Por verificar**).
- La fecha de cada valor se reconstruye como 1 de setiembre del año inicial + índice. Las series completas cuadran exactamente con 365 o 366 días por año hidrológico, así que el desfase por años bisiestos parece correcto. Aun así, la asignación de fechas la hizo un tercero (**Por verificar** contra una fuente oficial).
- El código de operador 72 probablemente corresponde a SENAMHI (**Por verificar**). El 63 está confirmado como PECHP por la fuente citada.
- Inconsistencias de metadatos en el SNIRH:
  - Tambo Grande 1586 aparece con río "San Francisco" y coordenadas (-4.95, -80.317) distintas de las de CAMELS-PE (-4.938, -80.344).
  - San Pedro 1442 ("Río Yapatero") tiene prácticamente las mismas coordenadas que Puente Ñácara.
  - Paltashaco 1200 aparece como "Río Los Gallegos".

  Todo esto está **Por verificar**.

### Gaps
- No pude descargar directamente del SNIRH, de Datos Abiertos, de Zenodo (CAMELS-PE) ni de GRDC: el proxy bloqueó todos esos hosts. Para reproducirlo en una red sin restricciones:
  - SNIRH: hacer POST a `https://snirh.ana.gob.pe/visorPorCuenca/Principal.asmx/CaudalSerie` con `Content-Type: application/json; charset=UTF-8` y el cuerpo `{"pIdEstacion":1586,"pIdOperador":63}`.
  - CAMELS-PE: descargar el zip de Zenodo y convertir con Q [m³/s] = q [mm/día] × A [km²] / 86.4.
- No hay datos de Malacasí 851 en el servicio. Tampoco encontré series de Tambogrande ni de Ñácara para 1970–1981 ni para 1987–2020 en ninguna fuente abierta.
- No pude confirmar si GRDC tiene alguna estación del río Piura. Habría que revisar el catálogo de estaciones de GRDC (portal.grdc.bafg.de) desde una red sin bloqueo.
- No encontré un "Observatorio del Agua Chira-Piura" con descarga de series. Solo encontré la página del Consejo de Recursos Hídricos de Cuenca, sin datos descargables visibles: [crhc.ana.gob.pe/node/280](https://crhc.ana.gob.pe/node/280) (**Por verificar**).

## 2. ¿Existe publicada una serie diaria histórica larga de Tambogrande o Ñácara (ONERN, DEPECHP, tesis de UDEP, UNP o UNALM, anexos de estudios de ANA)?

### Takeaway
No encontré ninguna serie **diaria** larga de Tambogrande ni de Ñácara publicada y accesible. Lo más largo que pude recuperar en abierto es esto:
- Ñácara: 1984–1991, fragmentario, más 2023–2026.
- Tambogrande: 1982–1987 y 2021–2023, fragmentario.
- Chulucanas: 1964–1969, completo.

CAMELS-PE (1981–2025) es el mejor candidato publicado para Tambogrande y Ñácara, pero falta verificar su cobertura real, porque solo pude ver los metadatos.

### Cited Findings
- La tesis UDEP de Claudia L. Palacios Santa Cruz, "Caudales de diseño en el río Piura y su variación histórica ante el Fenómeno El Niño" (abril de 2010), modela caudales de diseño para Los Ejidos, Tambogrande y Ñácara, "las estaciones más importantes del río Piura". Su índice incluye la "estación de aforo Tambogrande" y la "estación limnigráfica y de aforo Ñácara". Por su enfoque, trabaja con máximos y distribuciones de probabilidad, no necesariamente con series diarias. — [PIRHUA ICI_184.pdf](https://pirhua.udep.edu.pe/bitstream/handle/11042/2713/ICI_184.pdf); [Repositorios Latinoamericanos, registro 2250/6501961](https://repositorioslatinoamericanos.uchile.cl/handle/2250/6501961). No pude leer el PDF porque pirhua.udep.edu.pe estaba bloqueado. Si tiene anexos diarios queda **Por verificar**.
- Una tesis sobre la red de la cuenca Piura indica que la red la gestionan SENAMHI, PECHP y ANA, y que los caudales máximos se presentan en el área de influencia de la estación hidrométrica Tambogrande. — [Repositorios Latinoamericanos 2250/6504296](https://repositorioslatinoamericanos.uchile.cl/handle/2250/6504296)
- Hay una tesis UDEP con una propuesta de rediseño de la red hidrométrica Chira-Piura que plantea reactivar 2 estaciones hidrométricas en la cuenca Piura. Esto confirma que la red tiene estaciones paralizadas. — [Repositorios Latinoamericanos 2250/6504544](https://repositorioslatinoamericanos.uchile.cl/handle/2250/6504544)
- En CAMELS-PE, el archivo diario de TAMBOGRANDE (PE_47E0522C) y el de PUENTE NACARA (PE_47E0415A) tienen tamaños parecidos a los demás, de unos 1.16 MB. Eso es coherente con el calendario completo 1981–2025, pero `flow_obs` puede tener muchos vacíos: en la muestra del paquete, la estación Picoy tiene solo 9.63 % de registros válidos. — [rcamelspe data-manifest.csv y sample stations.csv](https://github.com/PaulESantos/rcamelspe)
- Los informes de JICA sobre las cuencas Chira y Piura incluyen un mapa de la red hidrometeorológica existente. — [JICA 12123873_05.pdf](https://openjicareport.jica.go.jp/pdf/12123873_05.pdf). No pude revisar si tiene anexos con datos diarios (**Por verificar**).

### Inferences
- Que el SNIRH entregue para Tambo Grande datos de 1982–1987, luego un salto hasta 2021 y finalmente 2023, sugiere que la base de ANA para la estación 1586 tiene grandes vacíos, o que el servicio expone solo 5 años hidrológicos con datos. Las dos explicaciones son posibles (**Por verificar**).
- Para series largas 1970s–2000s de Tambogrande y Ñácara, lo más probable es que estén en los archivos de PECHP (antes DEPECHP) y en los anexos de los estudios hidrológicos de JICA/PECHP y ANA. No encontré confirmación en línea.

### Gaps
- No encontré tesis de UNP ni de UNALM con anexos de caudales diarios de Tambogrande o Ñácara, ni registros ONERN digitalizados.
- No pude abrir los PDF de PIRHUA (dominio bloqueado).

## 3. Productos simulados (PISCO HyD ARNOVIC, GloFAS): cómo extraer el caudal diario en Tambogrande y Ñácara, periodo cubierto y si se pudo extraer

### Takeaway
**No pude extraer ninguna serie simulada**, porque HydroShare, Open-Meteo y Copernicus EWDS estaban bloqueados. Abajo dejo los identificadores exactos (COMID y coordenadas) y el procedimiento para hacerlo. Estas series son **simuladas, no observadas**.

### Cited Findings
- **PISCO-ARNOVIC (PISCO_HyD_ARNOVIC) v1.1** (SENAMHI, H. Llauca) es caudal diario simulado a escala nacional. Usa PISCO (precipitación y evapotranspiración, 0.1°) con el modelo ARNO/VIC y el ruteo RAPID, calibrado y validado con 43 estaciones en 1981–2020. Tiene licencia CC BY. Archivos: `PISCO_ARNOVIC_v1.1.nc`, `riv_pisco_arnovic_v1.1.gpkg` y `cat_pisco_arnovic_v1.1.gpkg`; los tramos se identifican por COMID. — [HydroShare f723d6c762ca45b6936dd9489bc44842](https://www.hydroshare.org/resource/f723d6c762ca45b6936dd9489bc44842)
- La v1.0 cubre del 01/01/1981 al 31/12/2022 según el perfil del autor. El periodo exacto de la v1.1 queda **Por verificar**. — [hllauca.github.io/resources.html](https://hllauca.github.io/resources.html); [HydroShare v1.0, 7046f3efa5db4cb78296fa06d833d721](https://www.hydroshare.org/resource/7046f3efa5db4cb78296fa06d833d721/)
- Hay un buscador de COMID en `https://hllauca.github.io/PISCO_HyD_ARNOVIC_map/Identificar_COMID.html`. En el HTML descargado de ese mapa aparecen los tramos COMID 9042297 y 9043133. — [hllauca/PISCO_HyD_ARNOVIC_map](https://github.com/hllauca/PISCO_HyD_ARNOVIC_map)
- COMID de los tramos en las estaciones, según CAMELS-PE, que usa la misma red PISCO-ARNOVIC:
  - Tambogrande: **9042297**.
  - Puente Ñácara: **9043133**.
  - San Pedro: **9043206**.
  - Hacienda Barrios: **9044140**.
  - Referencia Sánchez Cerro: 9043566.

  — [KURONO23/CAMELS-PE gauges gpkg](https://github.com/KURONO23/CAMELS-PE). Otro proyecto también usa el tramo COMID 9043566 de SONICS para Sánchez Cerro. — [carlosbr123/visor-bajo-piura, datos/caudal_externo.json](https://github.com/carlosbr123/visor-bajo-piura)
- CAMELS-PE incluye `flow_sim` (PISCO-ARNOVIC v1.1, en mm/día) para esas mismas estaciones. Es la forma más simple de obtener la serie simulada ya recortada, junto con la observada. — [rcamelspe, data_dictionary.csv](https://github.com/PaulESantos/rcamelspe/blob/main/inst/extdata/sample_camels_pe/01_metadata/data_dictionary.csv)
- **GloFAS v4 vía Open-Meteo.** El endpoint `/v1/flood` devuelve el caudal del río más grande dentro de un área de 5 km alrededor de la coordenada. Por defecto usa GloFAS v4, con datos continuos desde 1984 hasta 7 meses de pronóstico. La documentación lista el reanálisis v4 a 0.05° (~5 km), diario, de 1984 a julio de 2022. Advierte que, por la resolución, el río elegido puede no ser el correcto, y recomienda variar la coordenada en ±0.1°. — [Open-Meteo Flood API docs](https://open-meteo.com/en/docs/flood-api)
- GloFAS v4.0 trabaja en una grilla de 0.05° y cubre desde 1980 según Copernicus, mientras Open-Meteo indica 1984. Esta discrepancia no está resuelta. — [ECMWF, lanzamiento del reanálisis GloFAS v4.0](https://www.ecmwf.int/node/28201). La descarga oficial es el conjunto `cems-glofas-historical` en EWDS, que requiere cuenta y clave API. — [EWDS cems-glofas-historical](https://ewds.climate.copernicus.eu/datasets/cems-glofas-historical)
- Un proyecto que corrige GloFAS para Sánchez Cerro usó la celda (-5.125, -80.675) del producto "consolidated_v4" de Open-Meteo, con un desfase de -1 día y mapeo de cuantiles 2014–2020. Esto muestra que hay que elegir la celda a mano. — [carlosbr123/visor-bajo-piura, datos/glofas_correccion.json](https://github.com/carlosbr123/visor-bajo-piura)

### Inferences
- Consultas listas para ejecutar desde una red sin bloqueo (simulado, GloFAS v4):
  - Tambogrande: `https://flood-api.open-meteo.com/v1/flood?latitude=-4.938&longitude=-80.344&daily=river_discharge&start_date=1984-01-01&end_date=2026-11-26`
  - Puente Ñácara: `...latitude=-5.113&longitude=-80.173...`
  - Conviene probar celdas vecinas (±0.05°) y comparar el caudal medio y las crecidas de 1998, 2017 y 2023 con las observaciones del SNIRH para confirmar que la celda cae en el cauce del río Piura.
  - Como referencia de áreas de drenaje: Tambogrande ≈ 6,392 km² y Ñácara ≈ 4,641 km² (CAMELS-PE).
- PISCO-ARNOVIC: abrir el NetCDF con xarray y seleccionar los COMID 9042297 y 9043133. Los nombres de las dimensiones y variables no están verificados (**Por verificar**). El producto queda limitado al periodo 1981–2022 o al que indique la v1.1.

### Gaps
- No se extrajo ninguna serie simulada; no hay CSV de tipo "simulado".
- No verifiqué la estructura interna del NetCDF de PISCO-ARNOVIC v1.1 ni su periodo exacto.

## 4. Procedimiento formal para solicitar datos a ANA y PECHP

### Takeaway
No encontré en fuentes en línea un procedimiento específico publicado para pedir series hidrométricas a ANA o a PECHP. Todo lo de esta sección queda **Por verificar**.

### Cited Findings
- La cuenca Piura pertenece a la AAA Jequetepeque-Zarumilla, ALA Alto Piura (estaciones de la cuenca media y alta) y ALA Medio y Bajo Piura (Sánchez Cerro). Así figura en los metadatos del SNIRH de cada estación. — [cesarmiguel10/dashboard-bi-piura, data/estaciones.parquet](https://github.com/cesarmiguel10/dashboard-bi-piura/tree/main/data)
- ANA tiene 13 Autoridades Administrativas del Agua, entre ellas la AAA Jequetepeque-Zarumilla. — [Congreso, presentación ANA 2025](https://www.congreso.gob.pe/Docs/comisiones2021/CES-emergencias-desastres-covid19/files/ppt2025/4)_gestion_integrada_de_recursos_hidricos_-_ana.pdf)
- PECHP opera estaciones del SNIRH (operador 63), entre ellas Puente Sánchez Cerro, Tambo Grande, Barrios, Chililique y San Pedro 1442. — [carlosbr123/visor-bajo-piura](https://github.com/carlosbr123/visor-bajo-piura/blob/main/actualizar_caudal.py); [cesarmiguel10/dashboard-bi-piura](https://github.com/cesarmiguel10/dashboard-bi-piura)

### Inferences
- La vía razonable, **Por verificar**, es una solicitud de acceso a la información pública (Ley N.º 27806) presentada por mesa de partes virtual:
  - A ANA (sede central o AAA Jequetepeque-Zarumilla / ALA Alto Piura), pidiendo las series diarias completas de las estaciones del SNIRH con los códigos de la tabla de arriba.
  - A PECHP (Gobierno Regional Piura), pidiendo las estaciones con operador 63, en particular Tambo Grande 1586.

  En la solicitud conviene citar el código SNIRH y el operador de cada estación.

### Gaps
- No encontré el formulario, el plazo ni la oficina exacta de atención en las fuentes consultadas.
- No pude abrir ana.gob.pe ni chirapiura.gob.pe (bloqueados) para confirmar cómo funciona su portal de transparencia.
