# Fragmentos diarios esconden tres décadas del alto Piura

**No existe hoy en acceso abierto una serie diaria continua y larga de caudales del río Piura en su cuenca media y alta: lo que se pudo recuperar son 18 tramos diarios observados, uno por cada código del SNIRH-ANA (13 estaciones con nombre propio), que suman 21,101 días-estación entre el 01/09/1925 y el 25/11/2026, pero cubren solo 10,693 fechas distintas (Por verificar).** En el cauce principal (Tambogrande, Puente Ñácara, Puente Chulucanas, La Greda y Carrasquillo) hay 6,742 fechas con dato, unos 18.5 años equivalentes. Entre el 31/08/1991 y el 28/02/2021 no hay ningún día publicado (cálculo propio sobre los CSV; Por verificar). Por eso **ninguna serie diaria abierta del tramo medio registra las grandes crecidas de 1983, 1998 o 2017**: el tramo de Tambogrande se corta el 26/01/1983, justo antes del pico de El Niño 1982-83. La historia documentada de la red es mucho más larga que los datos disponibles. Según el estudio PNUD, Puente Piura (hoy Sánchez Cerro) mide desde 1925, Ñácara desde 1942 y Tambogrande desde 1954 ([PNUD / SIAR Piura 204](https://siar.regionpiura.gob.pe/documentos/normativa/204.pdf)). El Proyecto Especial Chira-Piura (PECHP) administra los aforos desde 1971 ([MINAM](https://repositoriodigital.minam.gob.pe/bitstream/handle/123456789/305/BIV00255.pdf?sequ=)). Hoy SENAMHI opera las estaciones automáticas Tambogrande (47E0522C) y Puente Ñácara (47E0415A). Los registros largos están, con toda probabilidad, en los archivos de ANA y PECHP. **CAMELS-PE se descargó el 27/11/2026 y llena solo parte del vacío:** su caudal observado en la cuenca media y alta empieza recién en 2014 (Tambogrande, 1,793 días entre 21/12/2014 y 12/04/2025) y su serie simulada 1981–2025 está recortada en el máximo observado, por lo que no sirve para picos (ver la sección de CAMELS-PE). Una advertencia vale para todo el informe: los CSV diarios vienen de una copia de terceros en GitHub del servicio web SNIRH-ANA (repositorio `cesarmiguel10/dashboard-bi-piura`, commit `d02a6886c56a349da83d819e848efb8f8c72c014`, corte al 26/11/2026) ([GitHub](https://github.com/cesarmiguel10/dashboard-bi-piura)). La red de la sesión bloqueó todos los servidores primarios (SNIRH, ANA, SENAMHI, PECHP, Zenodo, Datos Abiertos, UDEP, SIGRID). En consecuencia, **toda cifra de este informe es «Por verificar» contra el documento primario**, salvo que se indique lo contrario.

> **Convenciones.** Números con coma de miles y punto decimal (1,234.56). Fechas en dd/mm/aaaa. «Observado» significa medido en estación. «Simulado» significa salida de modelo (PISCO-ARNOVIC, GloFAS); ningún dato simulado se mezcla con los observados. **Puente Sánchez Cerro y la presa Los Ejidos están en el bajo Piura y se citan solo como referencia aguas abajo.** Las coordenadas se transcriben tal como se publicaron. Ninguna fuente consultada declara su datum, así que el datum figura como «Por verificar».

## La red se instaló entre 1925 y 1967 y hoy solo mantiene unas pocas estaciones

La medición del río Piura empezó aguas abajo, en el Puente Piura, con la Dirección de Aguas. El INDECI no tiene caudal medido para 1925: reconstruye el volumen de ese año por regresión lineal con los datos «que la Dirección de Aguas registra para el año 1926» ([INDECI-OEA, Mapa de Peligros Vol. I](https://sigrid.cenepred.gob.pe/docs/PARA%20PUBLICAR/INDECI/INFORME%20PRINCIPAL%20DEL%20%20ESTUDIO%20MAPA%20DE%20PELIGROS%20DE%20PIURA%20(VOLUMEN%20I).pdf)). En la cuenca media y alta, el estudio PNUD da estos periodos de la red antigua: **Puente Ñácara de 1942 a 1966, Tambogrande de 1954 a 1967 y Puente Piura de 1925 a 1966**. El mismo estudio presenta módulos mensuales de Malacasí, Carrasquillo y Ñácara ([PNUD / SIAR Piura 204](https://siar.regionpiura.gob.pe/documentos/normativa/204.pdf)). Otras fuentes agregan las siguientes estaciones:

- **Malacasí**: registro de 1967 a 1993 (27 años) y caudal medio de 12.1 m³/s. **Huarmaca**, en el tributario del mismo nombre: de 1966 a 1986, con 2.1 m³/s. Ambos datos son Por verificar porque vienen del resumen de un cuadro no abierto ([MINAM, Diagnóstico de la cuenca del río Piura](https://repositoriodigital.minam.gob.pe/bitstream/handle/123456789/305/BIV00255.pdf?sequ=)).
- **Puente Paltashaco**: caudales medios mensuales de 1950 a 1991 ([ANA, repositorio ANA0004092](https://repositorio.ana.gob.pe/bitstream/handle/20.500.12543/5541/ANA0004092_1.pdf.txt;jsessionid=8C635B9D0B35D11E619A45F95C137037?sequence=5)). Está en el río La Gallega, no en el cauce principal ([SIAR Piura, Cap. III](https://siar.regionpiura.gob.pe/documentos/normativa/571.pdf)).

En cuanto a los operadores, las fuentes solo permiten reconstruir una parte de la historia. El diagnóstico del MINAM atribuye los aforos al PECHP desde 1971 ([MINAM](https://repositoriodigital.minam.gob.pe/bitstream/handle/123456789/305/BIV00255.pdf?sequ=)). En el SNIRH, el **operador 63 corresponde al PECHP**, según un script que identifica la estación «1212, 63» como «Puente Sánchez Cerro (201201), Proyecto Especial Chira Piura» ([carlosbr123/visor-bajo-piura](https://github.com/carlosbr123/visor-bajo-piura/blob/main/actualizar_caudal.py)). **Que el operador 72 sea SENAMHI es una inferencia (Por verificar).** No se encontró ningún documento que fije la secuencia ONERN → DEPECHP/PECHP → SENAMHI/ANA, las fechas de traspaso ni la destrucción y reinstalación de estaciones en 1983 o 1998.

La red se fue reduciendo con los años. El MINAM cuenta 13 de 19 estaciones de aforo fuera de servicio ([MINAM](https://repositoriodigital.minam.gob.pe/bitstream/handle/123456789/305/BIV00255.pdf?sequ=)). Un documento regional sin fecha dice que, de 20 estaciones, solo seguían operativas «Sánchez Cerro/Los Ejidos, puente Ñácara y Puente Carrasquillo» ([SIAR Piura, Cap. I](https://siar.regionpiura.gob.pe/documentos/normativa/574.pdf)). La UDEP recomendó en 2022 reactivar Puente Paltashaco ([UDEP Hoy](https://www.udep.edu.pe/hoy/2022/05/modelacion-hidrologica-de-los-caudales-de-los-rios-chira-y-piura/)). Hoy SENAMHI presenta a **Tambogrande (código 47E0522C, distrito de Tambogrande) y Puente Ñácara (47E0415A, distrito de Chulucanas, Morropón)** como «punto de control representativo del río Piura». Los códigos figuran en las URLs de sus avisos ([SENAMHI, aviso Tambogrande](https://www.senamhi.gob.pe/?p=avisos-detalle-hidrologicos&ca=28032&ce=47E0522C); [SENAMHI, aviso Ñácara](https://senamhi.gob.pe/?ca=25617&ce=47E0415A&p=avisos-detalle-hidrologicos)). Ambas aparecen como estaciones hidrológicas automáticas (EHA), y SENAMHI también publica Hacienda Barrios, en el río Bigote ([SENAMHI, Monitoreo Piura](https://www.senamhi.gob.pe/?p=monitoreo-piura)). El sistema de alerta temprana opera desde 2002, con cinco estaciones hidrométricas activas. Entre Tambogrande y Los Ejidos hay un desfase de unas 18 horas ([Farías et al., LACCEI 2020](https://proyectofortalecimientodelsinacti.prociencia.gob.pe/wp-content/uploads/2020/10/LACCEI_2020_paper_276-Marina-Farias.pdf)).

**Coordenadas publicadas, todas con datum Por verificar:**

- **SNIRH** (lat, lon en grados decimales): Tambo Grande 1586: −4.95, −80.3167; Puente Ñácara 9609: −5.1129, −80.1730; Malacasí 851: −5.3167, −79.8833 ([dashboard-bi-piura, data/](https://github.com/cesarmiguel10/dashboard-bi-piura/tree/main/data)).
- **CAMELS-PE**: Tambogrande −4.938, −80.344; Puente Ñácara −5.113, −80.173 ([KURONO23/CAMELS-PE](https://github.com/KURONO23/CAMELS-PE)).
- **LACCEI 2020**: Puente Ñácara 05°07' S, 80°10' W, 149 msnm ([LACCEI 2020, FP276](https://laccei.org/LACCEI2020-VirtualEdition/full_papers/FP276.pdf)).

En Tambogrande, el SNIRH y CAMELS-PE difieren en unos 0.03° de longitud.

## Dieciocho tramos diarios observados, pero el cauce principal no tiene datos de 1991 a 2021

El servicio `CaudalSerie` del visor por cuenca del SNIRH entrega, por estación, una climatología diaria («Prom. Histórico») y **solo los últimos ~5 años hidrológicos con datos**, cada uno de setiembre a agosto ([dashboard-bi-piura, etl/snirh_caudales.py](https://github.com/cesarmiguel10/dashboard-bi-piura/blob/main/etl/snirh_caudales.py)). Por eso las estaciones cerradas aparecen como bloques de 4 o 5 años, y no como su registro completo. La copia de terceros guarda el inventario de 63 estaciones de la cuenca y la serie en `data/caudales_diarios.parquet`. Las fechas las reconstruyó ese tercero sumando el índice del día al 1 de setiembre del año inicial ([dashboard-bi-piura](https://github.com/cesarmiguel10/dashboard-bi-piura)). Esa asignación es coherente, porque los años completos tienen exactamente 365 o 366 días, pero es Por verificar.

La tabla resume la cobertura. La calculé a partir de los CSV de la carpeta de notas, sin modificarlos: comprobé sus sumas MD5 antes y después del cálculo. El porcentaje de faltantes se mide entre la primera y la última fecha con dato.

| Estación (código SNIRH / operador) | Río según SNIRH | Periodo con dato | Días con dato / días del periodo | % faltante | Máx. diario (fecha) | Vacíos principales |
|---|---|---|---|---|---|---|
| **Cauce principal** | | | | | | |
| Puente Chulucanas (1154 / 72), cerrada | Piura | 01/09/1964 – 31/08/1969 | 1,826 / 1,826 | 0.0 % | 1,423.58 (29/03/1965) | Ninguno; 1,384 días con 0 m³/s |
| La Greda (714 / 72), cerrada | Piura | 01/09/1971 – 31/12/1975 | 1,583 / 1,583 | 0.0 % | 533.00 (19/03/1972) | Ninguno |
| Carrasquillo (223 / 72), cerrada | Piura | 01/09/1978 – 31/12/1982 | 1,583 / 1,583 | 0.0 % | 614.00 (12/03/1981) | Ninguno; 796 días con 0 |
| Tambo Grande (1586 / 63) | «San Francisco» (Por verificar) | 01/09/1982 – 18/07/2023 | 566 / 14,931 | 96.2 % | 909.50 (15/03/2023) | 01/02/1987–27/02/2021 (12,446 días); 27/01/1983–31/12/1985 (1,070); 10/06/2021–14/03/2023 (643) |
| Puente Ñácara (Chulucanas), conv. (1196 / 72) | Piura | 30/06/1984 – 31/08/1991 | 521 / 2,619 | 80.1 % | 74.00 (22/04/1986) | 01/12/1986–28/02/1991 (1,551); 01/01/1985–30/11/1985 (334) |
| Puente Ñácara, auto. (9609 / 72) | Piura | 21/02/2023 – 25/11/2026 | 803 / 1,374 | 41.6 % | 1,105.89 (14/03/2023) | 16/05/2024–03/01/2025 (233); 02/07/2025–06/01/2026 (189); 04/08/2026–26/10/2026 (84) |
| Puente Ñácara, auto. (3554 / 72) | Piura | 21/02/2023 – 04/05/2026 | 816 / 1,169 | 30.2 % | 1,105.89 (14/03/2023) | 01/07/2025–17/11/2025 (140); 16/09/2024–27/12/2024 (103) |
| **Tributarios** | | | | | | |
| Puente Paltashaco (1200 / 72), cerrada | «Los Gallegos» (Por verificar) | 01/09/1978 – 31/12/1982 | 1,583 / 1,583 | 0.0 % | 21.85 (30/12/1982) | Ninguno |
| Chililique (302 / 63), cerrada | Yapatera | 01/09/1987 – 31/12/1991 | 1,583 / 1,583 | 0.0 % | 22.20 (11/03/1991) | Ninguno |
| Barrios (111 / 63), cerrada | Bigote | 01/09/1988 – 31/12/1992 | 1,583 / 1,583 | 0.0 % | 170.00 (05/03/1989) | Ninguno |
| San Pedro (1442 / 63), cerrada | «Yapatero» (coordenadas Por verificar) | 01/09/1988 – 30/04/1993 | 1,703 / 1,703 | 0.0 % | 100.00 (16/02/1989) | Ninguno |
| San Pedro (1804 / 72), cerrada | San Jorge | 02/09/2011 – 30/09/2017 | 337 / 2,221 | 84.8 % | 17.48 (26/03/2017) | 16/08/2014–31/01/2017 (900); 215 vacíos |
| San Pedro (9611 / 72) | s/d | 01/01/2023 – 25/11/2026 | 1,416 / 1,425 | 0.6 % | 26.40 (13/03/2023) | Días sueltos en feb.–mar. 2025 |
| San Pedro (3558 / 72) | s/d | 01/01/2023 – 04/05/2026 | 1,146 / 1,220 | 6.1 % | 26.40 (13/03/2023) | 12/02/2025–07/04/2025 (55) |
| Hacienda Barrios (9610 / 72) | s/d (cuenca Bigote) | 01/01/2023 – 25/11/2026 | 1,369 / 1,425 | 3.9 % | 174.33 (22/02/2024) | 19/04/2026–10/06/2026 (53) |
| Hacienda Barrios (3557 / 72) | s/d (cuenca Bigote) | 01/01/2023 – 17/04/2026 | 1,184 / 1,203 | 1.6 % | 174.73 (22/02/2024) | 21/12/2025–05/01/2026 (16) |
| Corral del Medio (406 / 72), cerrada | Corrales (ubicación Por verificar) | 01/09/1925 – 29/04/1992 | 1,133 / 24,348 | 95.3 % | 675.00 (09/03/1926) | Completo solo de 01/09/1925 a 31/08/1928; además dic. 1942 y 6 días de abr. 1992 |
| Teódulo Peña (1610 / 72), cerrada | Corrales (ubicación Por verificar) | 01/03/1991 – 30/04/1995 | 366 / 1,522 | 76.0 % | 220.00 (25/04/1992) | 27/09/1992–31/01/1995 (857) |
| **Referencia aguas abajo (bajo Piura, fuera del ámbito)** | | | | | | |
| Puente Sánchez Cerro (1212 / 63) | Piura | 01/09/2022 – 25/11/2026 | 1,496 / 1,547 | 3.3 % | 1,561.54 (15/03/2023) | 01/11/2025–26/11/2025 (26) |

*Todos los valores son caudal medio diario observado en m³/s (columna `tipo = observado`) y son Por verificar contra el SNIRH. Ningún CSV tiene fechas duplicadas ni valores negativos. El inventario del SNIRH incluye además Malacasí (851, op. 63), Salitral (3556 y 9590) y Corral del Medio automática (3559), pero **sin ninguna serie** ([dashboard-bi-piura, data/](https://github.com/cesarmiguel10/dashboard-bi-piura/tree/main/data)).*

Si se unen todas las fechas, los años con al menos un dato en el cauce principal son 1964–1969, 1971–1975, 1978–1987, 1991, 2021 y 2023–2026. No hay ninguno en 1970, 1976–1977, 1988–1990, 1992–2020 ni 2022. Los tributarios llenan parcialmente 1988–1995 y 2011–2017. Aun así, **en 1996–2010 no hay ninguna fecha observada abierta en toda la cuenca media y alta**. Esa ventana incluye El Niño 1997-98.

Los tramos vacíos tienen un patrón que conviene notar. El de Tambogrande va del 01/09/1982 al 26/01/1983 y sube de 0.2 m³/s a 483.90 m³/s el último día, justo cuando empezaba la crecida de 1983. En 2023 solo tiene 5 días, y el primero es el 15/03/2023: falta el 14/03/2023, el día del desborde. Las estaciones más recientes tienen un problema distinto, los **códigos duplicados**. Los pares 9609/3554 (Ñácara), 9611/3558 (San Pedro) y 9610/3557 (Hacienda Barrios) parecen ser la misma estación procesada de dos maneras. En las fechas comunes calculé estas correlaciones:

| Par | Correlación (r) | Días comunes | Días idénticos | Medianas (m³/s) |
|---|---|---|---|---|
| Ñácara 9609 / 3554 | 0.986 | 677 | 171 | 18.25 / 30.93 |
| San Pedro 9611 / 3558 | 0.992 | 1,146 | 906 | 2.09 / 2.08 |
| Hacienda Barrios 9610 / 3557 | 0.799 | 1,181 | 348 | 0.90 / 3.06 |

En Hacienda Barrios los dos códigos son idénticos en el año hidrológico 2022–2023 y se separan después: la media de 2023–2024 es 2.49 m³/s frente a 7.57 m³/s. Eso sugiere un cambio de procesamiento o de curva de descarga (Por verificar). Antes de usar cualquiera de estos pares hay que preguntar a ANA o PECHP cuál es la serie validada.

## Los picos documentados superan con creces lo que registran las series diarias

**Los extremos históricos de la cuenca media se conocen por documentos y por la prensa, no por series diarias abiertas.** La tabla los separa por estación y tipo de valor. Un máximo diario (media de 24 horas) no es comparable con un máximo instantáneo.

| Estación | Fecha | Tipo | Caudal (m³/s) | Fuente | Estado |
|---|---|---|---|---|---|
| Puente Chulucanas 1154 | 29/03/1965 | Máx. diario observado (SNIRH) | 1,423.58 | [dashboard-bi-piura](https://github.com/cesarmiguel10/dashboard-bi-piura) | Por verificar |
| Tambogrande | 1983 | Máx. promedio diario anual | 1,820.14 | [ANA, Expediente técnico](https://www.ana.gob.pe/sites/default/files/normatividad/files/EXPEDIETE%20TECNICO%20RIO%20PIURA%20(1).pdf) | Por verificar (columna dudosa) |
| Puente Ñácara | 1983 | Máx. promedio diario anual | 1,046.70 | [ANA, Expediente técnico](https://www.ana.gob.pe/sites/default/files/normatividad/files/EXPEDIETE%20TECNICO%20RIO%20PIURA%20(1).pdf) | Por verificar |
| Puente Ñácara | 31/03/1998 | Máximo (tipo no confirmado) | 2,165.0 | [SENAMHI/ANA, Condiciones hidrológicas El Niño 2023](https://cdn.www.gob.pe/uploads/document/file/5019575/condiciones%20hidrol%C3%B3gicas%20El%20Ni%C3%B1o%202023.pdf) | Por verificar |
| Puente Ñácara | 12/04/1998 | Máximo (tipo no confirmado) | 1,334.0 | Mismo documento | Por verificar |
| Tambogrande | 03/2017 | Lectura de prensa | 2,852 y 2,803 | [Diario Correo](https://diariocorreo.pe/edicion/piura/piura-el-caudal-maximo-para-piura-se-registrara-aun-en-marzo-733688/) | Por verificar (día exacto no confirmado) |
| Puente Ñácara | 14/03/2023 | Lectura horaria de las 07:00 | 1,110.69 | [INDECI, Boletín 986-2023](https://portal.indeci.gob.pe/wp-content/uploads/2023/03/BOLET%C3%8DN-INFORMATIVO-DE-AVISO-HIDROL%C3%93GICO-N%C2%B0986-2023-Del-14-03-2023.pdf) | Por verificar |
| Puente Ñácara | 14/03/2023 | Máx. diario observado (SNIRH 9609/3554) | 1,105.89 | [dashboard-bi-piura](https://github.com/cesarmiguel10/dashboard-bi-piura) | Por verificar |
| Puente Ñácara / Tambogrande | 03/2023 | Lectura de las 14:00 | 1,345.16 / 1,173.39 | [Gestión](https://gestion.pe/peru/lluvias-en-peru-ciclon-yaku-rio-piura-se-desborda-en-tambogrande-y-familias-son-evacuadas-a-albergues-inundaciones-noticia/) | Por verificar (día 13 o 14) |
| Tambogrande | 14/03/2023 | Pico citado por la prensa | 1,234 | [La República](https://larepublica.pe/sociedad/2023/03/14/piura-ciclon-yaku-caudal-del-rio-piura-llegaria-al-casco-urbano-con-1350-metros-cubicos-por-segundo-advirtio-senamhi-lrnd-151116) | Por verificar |
| *Ref. Puente Piura (hoy Sánchez Cerro)* | 31/03/1965 | Descarga máxima | 2,540 | [PNUD / SIAR 204](https://siar.regionpiura.gob.pe/documentos/normativa/204.pdf) | Por verificar (tipo no confirmado) |
| *Ref. Sánchez Cerro* | 07/05/1983 | Máx. instantáneo | 3,200 (INDECI: 3,500) | [ResearchGate](https://www.researchgate.net/publication/327671499_MODELACION_NUMERICA_DEL_RIO_PIURA); [INDECI](https://sigrid.cenepred.gob.pe/docs/PARA%20PUBLICAR/INDECI/INFORME%20PRINCIPAL%20DEL%20%20ESTUDIO%20MAPA%20DE%20PELIGROS%20DE%20PIURA%20(VOLUMEN%20I).pdf) | Por verificar (fuentes en conflicto) |
| *Ref. Sánchez Cerro* | 12/03/1998 (otra fuente: 10/03/1998) | Máx. instantáneo | 4,424 | [SIAR, Estudio de Hidrología](https://siar.regionpiura.gob.pe/documentos/normativa/200.pdf) | Por verificar |
| *Ref. Los Ejidos* | 27/03/2017, 12:00 h | Máx. instantáneo | 3,468 | [Reyes et al.](https://www.researchgate.net/publication/344967173_Pronostico_de_caudales_del_rio_Piura_calibrado_con_El_Nino_Costero_2017) | Por verificar |
| *Ref. Sánchez Cerro* | 15/03/2023 | Máx. diario observado (SNIRH 1212) | 1,561.54 | [dashboard-bi-piura](https://github.com/cesarmiguel10/dashboard-bi-piura) | Por verificar |

### Qué muestran estos valores sobre las series diarias

El dato de Ñácara en 2023 sirve para calibrar los demás. La media diaria del SNIRH para el 14/03/2023 es 1,105.89 m³/s, casi igual a la lectura horaria de 1,110.69 m³/s de las 07:00. En Sánchez Cerro la media diaria llega a 1,561.54 m³/s un día después, el 15/03/2023, un desfase coherente con las 12 a 18 horas de tránsito descritas para el sistema de pronóstico. En 1965 se ve el mismo patrón: 1,423.58 m³/s diarios en Chulucanas el 29/03/1965 y 2,540 m³/s en Puente Piura el 31/03/1965. La comparación tiene límites, porque es probable que el valor de Puente Piura sea instantáneo y porque entre ambos puntos entran los aportes de los tributarios.

Las series diarias abiertas, en cambio, no alcanzan los picos de los grandes eventos. Los valores de 1983, 1998 y 2017 en Tambogrande y Ñácara se conocen solo por cuadros no transcritos o por la prensa. Las lecturas de 2023 son mediciones horarias de boletines y no picos oficiales. Para 1925, 1932, 1939–1941, 1953, 1972, 1987, 1992 y 2012 no se encontró ningún pico publicado en la cuenca media. La única excepción es La Greda, con 533.00 m³/s el 19/03/1972 en la serie SNIRH (Por verificar). Tampoco hay mínimos de estiaje documentados.

Las series diarias sí cuentan la otra cara del río: en Chulucanas, el **75.8 % de los días de 1964 a 1969 tuvo caudal nulo**, y todo el año hidrológico 1967–1968 registró 0.00 m³/s. Esto coincide con la descripción de SENAMHI de «caudales mínimos o nulos» en estiaje ([SENAMHI, Monitoreo Piura](https://www.senamhi.gob.pe/?p=monitoreo-piura)). Los umbrales operativos dan contexto: alerta amarilla de 647.0 m³/s en Ñácara y 760.0 m³/s en Tambogrande ([SENAMHI, 03522SENA-641](https://www.senamhi.gob.pe/load/file/03522SENA-641.pdf)), y umbral rojo de 1,100 m³/s en Ñácara en 2023 ([INDECI, Boletín 986-2023](https://portal.indeci.gob.pe/wp-content/uploads/2023/03/BOLET%C3%8DN-INFORMATIVO-DE-AVISO-HIDROL%C3%93GICO-N%C2%B0986-2023-Del-14-03-2023.pdf)). Todos estos umbrales son Por verificar.

## La calidad de los datos limita su uso estadístico

El juicio más severo sobre los datos históricos viene del estudio PNUD: los caudales máximos de **Malacasí, Ñácara y Tambogrande «no tienen el período de observación suficientemente largo y su calidad no es adecuada»**. Por esa razón se excluyeron del análisis estadístico, y el máximo de 1982 en Tambogrande se calificó de «no lógico» ([PNUD / SIAR 204](https://siar.regionpiura.gob.pe/documentos/normativa/204.pdf)). El expediente de faja marginal de la ANA trabaja con máximos *promedio diarios* que convierte a instantáneos con una fórmula empírica ([ANA, Expediente técnico](https://www.ana.gob.pe/sites/default/files/normatividad/files/EXPEDIETE%20TECNICO%20RIO%20PIURA%20(1).pdf)). En otras palabras, los picos de diseño no son medidos.

En Sánchez Cerro, la referencia aguas abajo, las fuentes chocan entre sí. Una serie de máximos instantáneos incluye 14 m³/s como máximo de 2004 ([ClubEnsayos](https://www.clubensayos.com/Ciencia/ANALISIS-HIDROLOGICO-DE-LOS-CAUDALES-M%C3%81XIMOS-INSTANT%C3%81NEOS/4786199.html)). Para 1983 se citan 3,200 y 3,500 m³/s. Para 2017 circulan tres cifras: 3,468 m³/s medidos, 3,585 m³/s calculados ([Terrones, UDEP](https://pirhua.udep.edu.pe/handle/11042/4396)) y 3,016 m³/s en el desborde urbano ([REDER](https://www.revistareder.com/ojs/index.php/reder/article/download/201/228)). Se encontró una figura de la curva de calibración de Tambogrande en una tesis de la UNP, pero no su ecuación ni su vigencia ([UNP](https://repositorio.unp.edu.pe/server/api/core/bitstreams/4b950354-3015-4c4e-a599-4185119f8cb2/content)). Un estudio regional recomienda actualizar la curva altura-caudal cada año ([SIAR, Estudio de Hidrología](https://siar.regionpiura.gob.pe/documentos/normativa/200.pdf)).

Los metadatos del SNIRH también tienen errores. Tambo Grande 1586 figura en el «río San Francisco». San Pedro 1442 («río Yapatero») tiene casi las mismas coordenadas que Puente Ñácara (−5.1096, −80.1712). Paltashaco figura en el «río Los Gallegos» ([dashboard-bi-piura, data/](https://github.com/cesarmiguel10/dashboard-bi-piura/tree/main/data)). Todo esto es Por verificar.

Esa fragilidad explica por qué los análisis de frecuencia no concuerdan entre sí:

| Estudio | Estación | Serie | Estimación |
|---|---|---|---|
| [Tesis UCV, Cotrado](https://repositorio.ucv.edu.pe/bitstream/handle/20.500.12692/35629/Cotrado_ANY-SD.pdf?sequence=5) | Tambogrande | «1925–2008» | Q50 = 3,483.5 m³/s; Q100 = 4,299.7 m³/s |
| [Tesis UNP (CORE)](https://core.ac.uk/download/pdf/250077934.pdf) | No identificada | No indicada | Q100 = 2,663.423 m³/s (Gumbel) y 2,435.599 m³/s (Pearson III) |
| [PNUD / SIAR 204](https://siar.regionpiura.gob.pe/documentos/normativa/204.pdf) | Sánchez Cerro (referencia) | No indicada | 5,500 m³/s para Tr 200 |

Todas estas cifras son Por verificar. La serie de la tesis UCV no puede ser solo de Tambogrande, porque esa estación mide desde 1954. Los años 1925–1953 tuvieron que transferirse o rellenarse desde Puente Piura: es una serie compuesta, no observada (inferencia; Por verificar).

## CAMELS-PE, SNIRH, SENAMHI y la Ley 27806 pueden llenar el vacío

**Observado y simulado tienen que ir por separado.**

### CAMELS-PE: descargado, pero con observaciones solo desde 2014

**CAMELS-PE v1.0.1** (SENAMHI; Llauca et al., 2026; DOI 10.5281/zenodo.21195425; licencia CC-BY 4.0; publicado el 04/07/2026) se descargó desde Zenodo el 27/11/2026 (archivo `CAMELS-PE_v1.0.1.zip`, 121,352,279 bytes). Trae, por estación, `flow_obs` (observado, SENAMHI) y `flow_sim` (simulado, PISCO-ARNOVIC v1.1) diarios de 01/01/1981 a 31/12/2025, en mm/día. Los pasé a m³/s con Q = q × A × 1,000 / 86,400 (q en mm/día, A en km² del archivo `topographic_attributes.csv`). Los CSV resultantes están en la carpeta de notas como `diario_<estacion>_camels-pe.csv`.

| Estación (gauge_id) | Área (km²) | Lat / lon publicadas (datum Por verificar) | Observado: periodo | Días con dato | Q medio obs. (m³/s) | Máx. diario obs. (fecha) |
|---|---|---|---|---|---|---|
| Tambogrande (PE_47E0522C) | 6,392.09 | −4.938 / −80.344 | 21/12/2014 – 12/04/2025 | 1,793 | 79.57 | 1,660.09 (11/03/2017) |
| Puente Ñácara (PE_47E0415A) | 4,641.44 | −5.113 / −80.173 | 03/03/2021 – 01/07/2025 | 635 | 75.01 | 1,105.89 (14/03/2023) |
| Hacienda Barrios, río Bigote (PE_47E074C0) | 423.87 | −5.289 / −79.702 | 19/01/2015 – 31/12/2025 | 3,332 | 8.40 | 168.13 (09/03/2016) |
| San Pedro (PE_47E08444) | 157.59 | −5.080 / −80.016 | 01/09/2020 – 31/12/2025 | 1,754 | 2.77 | 26.40 (13/03/2023) |
| *Ref. Puente Sánchez Cerro (PE_201201)* | 7,507.31 | −5.194 / −80.623 | 01/03/1997 – 22/05/2023 | 4,878 | 102.87 | 3,366.99 (02/04/1998) |

*Cálculo propio sobre los archivos de CAMELS-PE; las medias se refieren solo a los días con dato, no a años completos.*

Lo que aporta: **Tambogrande gana 1,793 días observados entre 2015 y 2025** (años con más de 250 días: 2015, 2016, 2019, 2023), y Hacienda Barrios una serie casi continua de 11 años. Con esto, el vacío del cauce principal queda en 1992–2014, no en 1991–2021. Lo que no aporta: ninguna estación de la cuenca media y alta tiene observaciones antes de 2014, y **en 2017 Tambogrande solo tiene 134 días**: su máximo diario de 1,660.09 m³/s (11/03/2017) no incluye el pico del 27/03/2017. El máximo de Ñácara (1,105.89 m³/s el 14/03/2023) coincide exactamente con el de la serie SNIRH 9609/3554, lo que confirma que ambas fuentes vienen del mismo registro SENAMHI.

**La serie simulada (`flow_sim`) no sirve para extremos.** Cubre todo 1981–2025 sin vacíos, pero su máximo anual se repite exactamente en el máximo observado de cada estación: Tambogrande 1,660.09 m³/s en 1983, 1998 y 2023; San Pedro 26.40 m³/s todos los años; Sánchez Cerro 3,366.99 m³/s en 1998. Eso indica un recorte (probablemente por la corrección de sesgo) en el máximo observado. Además, la media simulada es bastante menor que la observada en los días comunes (Tambogrande 46.58 frente a 79.57 m³/s; Ñácara 24.95 frente a 75.01 m³/s, aunque los periodos no son iguales). Las correlaciones diarias obs.–sim. en los días comunes van de 0.63 (Hacienda Barrios) a 0.85 (Sánchez Cerro). Por eso la serie simulada solo sirve, con cautela, para el régimen y los volúmenes de 1981–2014, nunca para caudales máximos ni de diseño (inferencia; Por verificar).

### Otras fuentes de series observadas

El servicio SNIRH se puede consultar directamente con un POST a `https://snirh.ana.gob.pe/visorPorCuenca/Principal.asmx/CaudalSerie` con el cuerpo `{"pIdEstacion":1586,"pIdOperador":63}`. Hay dos salvedades: según un tercero, el servicio no responde a los servidores de GitHub desde el 17/09/2026, y en años vacíos devuelve JSON malformado ([carlosbr123/visor-bajo-piura](https://github.com/carlosbr123/visor-bajo-piura/blob/main/actualizar_caudal.py)). Datos Abiertos Perú tiene un conjunto del GORE Piura/PECHP con `PROMEDIO24H` y `MAXIMA24H` por estación ([stephany-toribio/piura-hydrology-analytics](https://github.com/stephany-toribio/piura-hydrology-analytics)). Su periodo y sus estaciones son Por verificar. Los datos diarios de GRDC son de uso restringido ([JRC AquaKnow](https://aquaknow.jrc.ec.europa.eu/node/18911)), y no se encontró ninguna estación del río Piura en su catálogo.

### Productos simulados (no observados)

**PISCO-ARNOVIC v1.1** (SENAMHI, ARNO/VIC con ruteo RAPID, licencia CC BY) se consulta por COMID ([HydroShare](https://www.hydroshare.org/resource/f723d6c762ca45b6936dd9489bc44842)). La v1.0 cubre del 01/01/1981 al 31/12/2022 ([hllauca](https://hllauca.github.io/resources.html)). CAMELS-PE trae la v1.1 por estación como `flow_sim` hasta el 31/12/2025, con el recorte de picos descrito arriba.

**GloFAS v4** tiene un reanálisis diario a 0.05° y se puede consultar con la Flood API de Open-Meteo. Esa API toma el río más grande en un radio de 5 km del punto, por lo que la celda debe validarse a mano ([Open-Meteo](https://open-meteo.com/en/docs/flood-api)). Open-Meteo indica que los datos empiezan en 1984, mientras que ECMWF habla de 1980 ([ECMWF](https://www.ecmwf.int/node/28201)). La discrepancia está sin resolver.

### Solicitud formal de datos

SENAMHI atiende pedidos de datos históricos como servicio TUSNE, regulado por la RPE N.º 041-2024-SENAMHI/PREJ. El formulario pide el **código antiguo** de la estación y su tipo (CONV/AUTO) ([SENAMHI, TUSNE](https://www.senamhi.gob.pe/?p=tusne); [SENAMHI Piura, solicitud](https://www.senamhi.gob.pe/main.php?dp=piura&p=solicitud-servicio)). La única tarifa conocida es de S/ 87.20 por variable diaria y estación, y es de 2021, así que solo sirve como referencia ([RPE 050-2021](https://www.senamhi.gob.pe/load/file/00701SENA-1374.pdf)); la vigente es Por verificar.

Para las series largas en poder de la ANA (AAA Jequetepeque-Zarumilla, ALA Alto Piura) y del PECHP (operador 63: Tambo Grande, Barrios, Chililique, San Pedro 1442), la vía razonable es una solicitud de acceso a la información pública (Ley N.º 27806) que cite el código SNIRH y el operador de cada estación. El procedimiento y los plazos son Por verificar ([dashboard-bi-piura, data/](https://github.com/cesarmiguel10/dashboard-bi-piura/tree/main/data)).

También hay que releer varios documentos con los dominios desbloqueados: el cuadro de máximos diarios del expediente ANA, el PNUD/SIAR 204, la tesis UDEP ICI_184 de Palacios (2010) sobre Los Ejidos, Tambogrande y Ñácara ([PIRHUA](https://pirhua.udep.edu.pe/bitstream/handle/11042/2713/ICI_184.pdf)) y los anexos de JICA ([JICA](https://openjicareport.jica.go.jp/pdf/12123873_05.pdf)). Antes hay que permitir en la red los hosts `snirh.ana.gob.pe`, `www.senamhi.gob.pe`, `zenodo.org`, `www.datosabiertos.gob.pe`, `pirhua.udep.edu.pe`, `siar.regionpiura.gob.pe`, `www.ana.gob.pe` y `sigrid.cenepred.gob.pe`, que el proxy de la sesión rechazó.

## Conclusión

La pregunta «¿cuál es la serie diaria más larga?» tiene una respuesta incómoda. La serie larga existe en papel desde 1942–1954, pero el archivo abierto la entrega por tramos, y le faltan justo los años que importan para el diseño. Los periodos 1983, 1997–1998 y 2017 en Tambogrande y Ñácara no tienen ni un solo día observado en acceso abierto. Por eso cualquier análisis de frecuencia hecho hoy con estas series descansa en máximos anuales transcritos de cuadros o en series compuestas, no en registros diarios verificables. La dispersión del Q100 publicado, entre ~2,400 y ~4,300 m³/s (Por verificar), es el síntoma de ese problema. Otro riesgo es menos visible. En los tres pares de estaciones con dos códigos, las medianas cambian hasta 3.4 veces según el código que se use. Elegir sin preguntar a ANA o PECHP cuál es la serie validada cambia los resultados de cualquier balance o calibración.

CAMELS-PE ya se descargó y su aporte es limitado: observaciones solo desde 2014 en la cuenca media y alta, y una serie simulada 1981–2025 con los picos recortados. La salida práctica, por tanto, es presentar solicitudes formales a ANA y PECHP (Ley 27806) y a SENAMHI (TUSNE) por los registros de 1942–2014 que el SNIRH solo expone como climatología. Mientras tanto, los productos simulados (PISCO-ARNOVIC y GloFAS) pueden ayudar a reconstruir 1992–2020, siempre como datos rotulados «simulado», nunca mezclados con lo observado. Conviene validarlos contra los pocos tramos observados de 2021–2026, que sí registran el evento Yaku de 2023 en Ñácara, Sánchez Cerro y los tributarios.
