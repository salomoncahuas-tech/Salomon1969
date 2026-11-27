# Caudales diarios SENAMHI – estaciones hidrométricas de la cuenca media y alta del río Piura

> **Estado de la descarga (27/11/2026): NO se descargó ningún dato diario.** No hay archivos `diario_<estacion>_senamhi.csv` en esta carpeta porque no se recuperó ningún valor (no se inventaron ni rellenaron datos). El motivo es un bloqueo de red de la sesión de trabajo, no una limitación probada del portal SENAMHI (detalle abajo).
>
> Se dejó en esta carpeta un script de descarga **sin probar** (`descargar_senamhi_diarios.py`), para ejecutarlo cuando el host `www.senamhi.gob.pe` esté permitido.

## 1. ¿Qué estaciones hidrométricas SENAMHI hay en la cuenca media y alta del Piura (código, tipo, coordenadas, altitud, estado, periodo)?

### Takeaway
Se confirmaron dos estaciones hidrológicas automáticas (EHA) SENAMHI en el tramo medio, con su código nuevo: **Tambogrande (47E0522C)** y **Puente Ñácara (47E0415A)**. Puente Sánchez Cerro es la referencia aguas abajo, pero su código no se confirmó. Malacasí, Huarmaca, Paltashaco, Barrios y San Pedro solo aparecen en fuentes secundarias, como estaciones históricas o inactivas. Sus códigos, coordenadas oficiales y operador actual quedan **Por verificar**.

### Cited Findings
- **Tambogrande (río Piura):** según SENAMHI es «un punto de control representativo del río Piura», ubicado en el distrito de Tambogrande, provincia de Piura. Sus avisos hidrológicos usan el código de estación `ce=47E0522C` en la URL. — [SENAMHI – aviso "situación actual del caudal del río Piura – estación Tambogrande"](https://www.senamhi.gob.pe/?p=avisos-detalle-hidrologicos&ca=28032&ce=47E0522C); [SENAMHI – aviso "incremento del caudal del río Piura – estación Tambogrande"](https://www.senamhi.gob.pe/?p=avisos-detalle-hidrologicos&ca=67606&ce=47E0522C); [SENAMHI – Monitoreo Piura](https://www.senamhi.gob.pe/?p=monitoreo-piura)
- **Tambogrande, caudales y umbrales:** un aviso de marzo de 2026 reportó un caudal de 1,156.22 m³/s (umbral NARANJA). El umbral rojo de la tabla es 1,600 m³/s. Estas cifras vienen de un resumen del buscador; no se pudo abrir la página. — [SENAMHI – aviso Tambogrande](https://www.senamhi.gob.pe/?p=avisos-detalle-hidrologicos&ca=28032&ce=47E0522C)
- **Puente Ñácara (río Piura):** sus avisos hidrológicos usan el código `ce=47E0415A`. SENAMHI la describe como punto de control representativo del río Piura, en el distrito de Chulucanas, provincia de Morropón. — [SENAMHI – Avisos detalle hidrológicos (Pte. Ñácara)](https://senamhi.gob.pe/?ca=25617&ce=47E0415A&p=avisos-detalle-hidrologicos); [SENAMHI (servicios) – aviso ca=25245, ce=47E0415A](https://www.senamhi.gob.pe/servicios/?p=avisos-detalle-hidrologicos&ca=25245&ce=47E0415A); [SENAMHI – Monitoreo Piura](https://www.senamhi.gob.pe/?p=monitoreo-piura)
- **Tipo de estación:** los reportes hidrológicos de la DZ Piura nombran «EHA Tambogrande» y «EHA Pte Ñácara», es decir, estaciones hidrológicas automáticas. — [SENAMHI – informe técnico 03522SENA-641](https://www.senamhi.gob.pe/load/file/03522SENA-641.pdf) (dato tomado del resumen del buscador; PDF no abierto)
- **Lecturas recientes (2026):** el reporte hidrológico Piura N.º 065 (13/04/2026) da 32.93 m³/s en Puente Ñácara a las 08:00 y «S/D» (sin dato) en Tambogrande. El reporte N.º 029 (18/02/2026) da 283.80 m³/s en Puente Ñácara y S/D en Tambogrande. Los reportes citan como fuentes al Proyecto Especial Chira-Piura y a la Junta de Usuarios San Lorenzo, y remiten a la plataforma PHISIS de SENAMHI. — [SENAMHI – reportes 03525SENA-1xx](https://www.senamhi.gob.pe/load/file/03525SENA-142.pdf) (resumen del buscador; no se identificó qué PDF corresponde a cada número)
- **Coordenadas de Puente Ñácara (fuente secundaria, no SENAMHI):** 05°07' S, 80°10' W, a 149 msnm. No indica datum. — [LACCEI 2020, FP276 – modelo de pronóstico de caudales del río Piura](https://laccei.org/LACCEI2020-VirtualEdition/full_papers/FP276.pdf). Un informe estudiantil da otra altitud (100 msnm), sin atarla claramente al punto de la estación. — [Studocu – Informe defensa ribereña](https://www.studocu.com/pe/document/institucion-educativa-privada-familia-universal-school/economia-politica-economico-ammistrazione-finanza-e-marketing-editar/informedel-logro-obras-hidraulicas/92469539)
- **Malacasí (río Piura):** registro de 1967 a 1993 (27 años con dato) y caudal medio de 12.1 m³/s, según una tabla del diagnóstico de la cuenca del río Piura (MINAM). — [MINAM – Diagnóstico de la cuenca del río Piura con enfoque de gestión de riesgo](https://repositoriodigital.minam.gob.pe/bitstream/handle/123456789/305/BIV00255.pdf?sequ=) (resumen del buscador; PDF no abierto)
- **Malacasí, calidad del registro:** otro informe advierte que sus caudales máximos instantáneos no tienen un periodo suficientemente largo ni calidad adecuada para análisis estadístico. — [Gobierno Regional Piura / SIAR – Proyecto PNUD (doc. 204)](https://siar.regionpiura.gob.pe/documentos/normativa/204.pdf)
- **Huarmaca (río Huarmaca, tributario):** registro de 1966 a 1986 (21 años) y caudal medio de 2.1 m³/s. — [MINAM – Diagnóstico cuenca Piura](https://repositoriodigital.minam.gob.pe/bitstream/handle/123456789/305/BIV00255.pdf?sequ=)
- **Paltashaco:** un estudio de la UDEP recomienda «reactivar» la estación Puente Paltashaco, lo que indica que no estaba operativa en 2022. — [UDEP Hoy, 2022 – Modelación hidrológica de los caudales de los ríos Chira y Piura](https://www.udep.edu.pe/hoy/2022/05/modelacion-hidrologica-de-los-caudales-de-los-rios-chira-y-piura/)
- **Barrios y San Pedro:** aparecen como subcuencas con estación hidrométrica en una tesis, sin periodo de registro en el resumen. — [SlideShare – Análisis hidrológico (tesis)](https://www.slideshare.net/slideshow/analisis-hidrologicotesis/245088081)
- **Red histórica:** 13 de 19 estaciones de aforo de la cuenca dejaron de funcionar. — [MINAM – Diagnóstico cuenca Piura](https://repositoriodigital.minam.gob.pe/bitstream/handle/123456789/305/BIV00255.pdf?sequ=). Las mediciones en el río Piura empezaron en 1925, y el Proyecto Especial Chira-Piura hace los aforos desde 1971. — [VSIP – Resumen Ejecutivo Cuenca Piura](https://vsip.info/resumen-ejecutivo-cuenca-piura-pdf-free.html)
- **Puente Sánchez Cerro (referencia aguas abajo, fuera del ámbito):** umbrales de alerta de 900 m³/s (amarilla), 1,300 m³/s (naranja) y 1,900 m³/s (roja). — [Ahora.com.pe](https://ahora.com.pe/policiales/senamhi-alerta-huaicos-piura-quebradas); [Norte Sostenible](https://nortesostenible.com/rio-piura-senamhi-proyecta-aumento-caudal-noviembre/). El aviso N.º 054 (marzo de 2017, nivel rojo) proyectaba entre 2,800 y 3,100 m³/s en Sánchez Cerro. — [SENAMHI – Aviso hidrológico 2017-054](https://senamhi.gob.pe/?a=2017&b=054&c=026&d=SENA&p=aviso-hidrologico-detalle)
- **Formulario de solicitud:** en el formulario SENAMHI de solicitud de datos, el campo «Código» es el **código antiguo** de la estación. Cada estación figura como convencional (CONV) o automática (AUTO), con su estado de funcionamiento. — [SENAMHI Piura – Solicitud de servicio](https://www.senamhi.gob.pe/main.php?dp=piura&p=solicitud-servicio)
- Existe un listado oficial de estaciones SENAMHI publicado en Scribd (más de 150 estaciones, con código, cuenca y coordenadas). No se pudo verificar su contenido. — [Scribd – Lista oficial de estaciones SENAMHI](https://www.scribd.com/doc/315197407/Lista-Oficial-de-Estaciones-Senamhi)

### Inferences
- Los códigos 47E0522C y 47E0415A siguen el formato de las estaciones automáticas SENAMHI (8 caracteres hexadecimales, prefijo 47…). Son los valores que debería recibir el parámetro `estaciones=` del endpoint `_dato_esta_tipo0X.php`. Esto es **Por verificar**, porque el endpoint no se pudo consultar.
- Las estaciones convencionales históricas (Malacasí, Huarmaca, Paltashaco, Barrios, San Pedro) probablemente tienen código numérico antiguo (tipo 2xxxxx). Varias las operó el Proyecto Especial Chira-Piura y no SENAMHI, así que su serie diaria podría no estar en el portal SENAMHI. **Por verificar.**
- El periodo «automático» de Tambogrande y Puente Ñácara probablemente es reciente (EHA). Las series largas (desde 1925 o 1971) parecen estar en manos del PECHP o de la ANA, no del portal SENAMHI. **Por verificar.**

### Gaps
- Faltan, para todas las estaciones: código oficial (nuevo y antiguo), coordenadas oficiales con datum, altitud, estado y fecha de inicio. No se pudo abrir el mapa de estaciones de SENAMHI (host bloqueado).
- No se confirmó el código de Puente Sánchez Cerro, Malacasí, Paltashaco, Barrios, San Pedro ni Huarmaca.
- No se halló una fuente que confirme si Malacasí, Huarmaca, Barrios o San Pedro tienen equipo SENAMHI activo hoy.

## 2. ¿Se pueden descargar los datos diarios desde el portal SENAMHI? ¿Cuál es el primer y último registro y qué periodos faltan?

### Takeaway
No se pudo comprobar en esta sesión. Todas las conexiones a `www.senamhi.gob.pe` (y a `senamhi.gob.pe`, `idesep.senamhi.gob.pe` y `repositorio.senamhi.gob.pe`) fueron rechazadas por el proxy de red: HTTP 403 en el CONNECT, o error de resolución DNS en WebFetch. Por eso no se descargó ninguna serie ni se calculó cobertura alguna (primer y último dato, % de faltantes, vacíos por año).

### Cited Findings
- **Intentos fallidos el 27/11/2026, con el error exacto:**
  - `curl https://www.senamhi.gob.pe/mapas/mapa-estaciones-2/` → `curl: (56) CONNECT tunnel failed, response 403` (proxy de la sesión).
  - `curl https://www.senamhi.gob.pe/?p=estaciones` → `CONNECT tunnel failed, response 403`.
  - `https://senamhi.gob.pe/`, `https://www.senamhi.gob.pe/site/descarga-datos/`, `https://idesep.senamhi.gob.pe/`, `https://repositorio.senamhi.gob.pe/` → sin conexión (código 000).
  - WebFetch `https://www.senamhi.gob.pe/mapas/mapa-estaciones-2/` → `getaddrinfo ENOTFOUND www.senamhi.gob.pe` (host fuera de los dominios permitidos de la sesión).
  - Hosts alternativos también bloqueados: `snirh.ana.gob.pe`, `www.ana.gob.pe`, `datosabiertos.gob.pe`, `www.gob.pe`, `web.archive.org`, `zenodo.org`, `iridl.ldeo.columbia.edu`, `laccei.org`, `repositoriodigital.minam.gob.pe`, `siar.regionpiura.gob.pe`.
  - El scraper externo (Firecrawl) no se ejecutó porque la cuenta no tenía créditos.
  - **No se probaron:** el endpoint `https://www.senamhi.gob.pe/mapas/mapa-estaciones-2/_dato_esta_tipo02.php?estaciones=<CODIGO>&CBOFiltro=<AAAAMM>&t_e=H&estado=REAL&cod_old=&cate_esta=...` y las URLs de `?p=estaciones`. No hubo conexión al host.
- **Descarga en línea:** según gob.pe, se pueden descargar datos diarios de estaciones meteorológicas nacionales en formato .txt separado por comas. Hay que registrarse antes (nombre, DNI, correo y teléfono); la contraseña llega por correo. La página tiene unos 2,300 días de antigüedad y habla de estaciones *meteorológicas*; no está confirmado que incluya caudal. — [gob.pe – Conocer los datos hidrometeorológicos del Perú](https://www.gob.pe/9312-conocer-los-datos-hidrometeorologicos-del-peru-descargar-los-datos-meteorologicos)
- **Páginas de descarga existentes:** — [SENAMHI – Descarga de datos](https://www.senamhi.gob.pe/site/descarga-datos/); [SENAMHI – Descarga datos hidrometeorológicos](https://www.senamhi.gob.pe/?dp=tumbes&p=descarga-datos-hidrometeorologicos); [SENAMHI – Datos hidrometeorológicos a nivel nacional](https://www.senamhi.gob.pe/servicios/?p=estaciones)
- **Pronósticos de Puente Ñácara:** SENAMHI publica pronósticos mensuales de caudal en Puente Ñácara (cuenca del Piura). — [SENAMHI Piura – Pronóstico hidrológico](https://www.senamhi.gob.pe/main.php?dp=piura&p=pronostico-hidrologico)

### Inferences
- **Para desbloquear la descarga:** hay que permitir el host `www.senamhi.gob.pe` en la red de la sesión. Hoy el proxy lo rechaza. Después se puede correr `descargar_senamhi_diarios.py`, que recorre mes a mes con 1.5 s de pausa.
- Las tablas «_dato» del mapa de estaciones suelen mostrar lecturas horarias o diarias (nivel y caudal) de estaciones automáticas y datos diarios de convencionales. **Por verificar** si para Tambogrande y Puente Ñácara entregan caudal diario o solo nivel u horario.

### Gaps
- Primer y último registro diario por estación: **sin dato** (no se descargó nada).
- % de días faltantes y vacíos por año: **sin dato**.
- Parámetros reales del endpoint `_dato` (`cate_esta`, `tipo02` frente a `tipo03` u otro, `cod_old`): **Por verificar** inspeccionando el JS de `mapa-estaciones-2` cuando el host esté accesible.

## 3. Si los datos diarios solo se obtienen por solicitud formal, ¿cuál es el procedimiento, formulario, oficina y costo?

### Takeaway
SENAMHI atiende la solicitud de datos hidrometeorológicos históricos como servicio **TUSNE** (Texto Único de Servicios No Exclusivos), es decir, de pago. Se pide con un formulario en línea o en mesa de partes, también en la Dirección Zonal Piura. No se pudo confirmar el monto vigente por estación y variable para caudal diario. Una tarifa de 2021 de S/ 87.20 (IGV incluido) por datos diarios de una variable en una estación es solo referencial y está **Por verificar**.

### Cited Findings
- **Formulario TUSNE:** el formulario de solicitud cubre datos hidrometeorológicos históricos (precipitación, temperatura, humedad, etc.). Los datos personales o de la empresa se usan para emitir boleta o factura «si paga por el servicio». — [SENAMHI – Solicitud de servicio](https://www.senamhi.gob.pe/?p=solicitud-servicio); [SENAMHI Piura – Solicitud de servicio](https://www.senamhi.gob.pe/main.php?dp=piura&p=solicitud-servicio)
- **Formulario, campos:** se indica el «Código» antiguo de la estación, su tipo (CONV o AUTO) y su estado. — [SENAMHI Piura – Solicitud de servicio](https://www.senamhi.gob.pe/main.php?dp=piura&p=solicitud-servicio)
- **Canales de atención:** canal virtual, o presencial en la Unidad de Atención al Ciudadano y Gestión Documental (UACGD), Jr. Cahuide 721, Jesús María, Lima, de lunes a viernes de 8:30 a 17:15. — [SENAMHI – TUSNE](https://www.senamhi.gob.pe/?p=tusne)
- **Oficina regional (Piura/Tumbes):** teléfono (073) 294967, de lunes a viernes de 8:30 a 17:15. — [SENAMHI – Solicitud de servicio](https://www.senamhi.gob.pe/?p=solicitud-servicio) (resumen del buscador). La dirección física de la DZ1 Piura queda **Por verificar**.
- **Norma vigente:** la página TUSNE remite a la Resolución de Presidencia Ejecutiva N.º 041-2024-SENAMHI/PREJ para la solicitud de datos históricos. — [SENAMHI – TUSNE](https://www.senamhi.gob.pe/?p=tusne); [TUSNE SENAMHI 2024 (PDF)](https://www.senamhi.gob.pe/pdf/trans/tusne_senamhi_2024.pdf)
- **Tarifa 2021 (referencial, no vigente):** la RPE N.º 050-2021-SENAMHI/PREJ fija S/ 87.20 con IGV por datos diarios de una variable en una estación. La tabla se extrajo incompleta. — [RPE 050-2021-SENAMHI/PREJ](https://www.senamhi.gob.pe/load/file/00701SENA-1374.pdf)
- **Monto en el TUSNE 2024:** el PDF muestra S/ 1,354.99 bajo «Dirección de Redes de Observación y Datos», sin que quede claro a qué servicio corresponde. — [TUSNE SENAMHI 2024](https://www.senamhi.gob.pe/pdf/trans/tusne_senamhi_2024.pdf) (**Por verificar**; fragmento mal extraído)

### Inferences
- **Vía formal sugerida, si la descarga web no da caudal diario:**
  1. Llenar el formulario TUSNE (en línea o en la DZ1 Piura) con los códigos antiguos de las estaciones, la variable «caudal medio diario (m³/s)» y el periodo.
  2. Pagar según el TUSNE 2024.
  3. En paralelo, presentar una solicitud de acceso a la información pública (Ley N.º 27806) para datos que SENAMHI ya tenga procesados.

  El costo vía Ley de Transparencia suele limitarse a la reproducción, pero eso es **Por verificar** para SENAMHI.
- Para las series largas históricas (Malacasí, Huarmaca, Paltashaco, Barrios, San Pedro), conviene pedir también al Proyecto Especial Chira-Piura y a la ANA (AAA Jequetepeque-Zarumilla / ALA Alto Piura). Es posible que SENAMHI no tenga esas series.

### Gaps
- Tarifa vigente exacta (TUSNE 2024, RPE 041-2024) para caudal diario por estación y año: no se pudo leer el PDF (host bloqueado).
- Plazo de atención y formato de entrega (Excel o txt): no se encontró.
- Procedimiento y costo vía Ley de Transparencia ante SENAMHI: no se encontró una fuente específica.
