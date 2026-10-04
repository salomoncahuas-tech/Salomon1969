# Formatos de entrega ANIN

## Identidad institucional
- **Color principal:** verde `#1B4D2E` (alternativa: azul institucional `#1B4F72`). Fila alterna / fondo suave: `#E8F0EA`; neutro `#F2F2F2`.
- **Fuente:** Arial.
- **Encabezado (3 líneas, en este orden):**
  1. AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN
  2. DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME
  3. SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN - SESDI
- **Subtítulo de proyecto:** `PROYECTO IN PIURA | CUI 2669244 | Cuenca Alta del Río Piura | UTM WGS 84 Zona 17S (EPSG:32717)`
- Idioma: español (Perú). Decimales con punto y miles con coma, como en el resto de los archivos del proyecto.

## Excel (matrices, plantillas, reportes)
Usa `openpyxl` y `scripts/anin_excel.py`. Reglas:
- Encabezados institucionales en las filas superiores (combinados), luego título del reporte y fuente de datos.
- Fila de cabecera de tabla: fondo verde, texto blanco, Arial negrita.
- Filas alternas, bordes finos, paneles congelados bajo la cabecera, filtro automático.
- **Totales con fórmulas** (`=SUM(...)`), nunca valores pegados: el revisor debe poder auditar.
- Hoja final de notas o avisos cuando haya datos pendientes o supuestos.
- Si el libro tendrá fórmulas, recalcúlalo y revisa que no haya errores (`#REF!`, `#DIV/0!`) antes de entregar. Los archivos creados con openpyxl no guardan el resultado de las fórmulas hasta que se abren en Excel; si LibreOffice no está disponible para recalcular, compara los resultados esperados calculados en Python con las fórmulas escritas y avísalo. Para tareas Excel complejas, apóyate también en la skill `xlsx`.

## Word (informes)
Usa la biblioteca `docx` (ver skill `docx`) con los helpers de `scripts/anin_docx.js`, que ya resuelven el formato de abajo. Reglas:
- Orientación según contenido: vertical para texto, horizontal si hay matrices anchas.
- Encabezado institucional en la cabecera de página; pie con proyecto y número de página.
- Estilos de título jerárquicos reales (Título 1/2/3), no texto en negrita simulando títulos.
- Tablas con bordes y cabecera en verde con texto blanco; figuras y cuadros numerados con fuente.
- Estructura típica de informe técnico: antecedentes, objetivo, alcance y metodología, resultados, conclusiones, recomendaciones, anexos.
- Las cifras del texto se calculan desde los datos (catálogo, Excel), no se digitan: así el texto, las tablas y las figuras no se contradicen.
- Lo que falta se marca `[POR DEFINIR]` y se lista en recomendaciones; no se rellena con supuestos.
- **Verificación:** valida con `python <skill docx>/scripts/office/validate.py archivo.docx` (requiere `pip install defusedxml`) y, si LibreOffice puede abrir archivos, renderiza a PDF para revisar el diseño. Si LibreOffice falla ("source file could not be loaded" incluso con un .txt), el entorno está roto: revisa el contenido con `python-docx`, confirma que cada tabla suma el ancho útil y di al usuario que no se pudo revisar el diseño visualmente.

## PowerPoint
Usa `pptxgenjs` (ver skill `pptx`). Tema oscuro institucional, paleta verde / teal / dorado, **máximo 2 fuentes por diapositiva** (Arial + una complementaria). Una idea por diapositiva; cifras clave grandes; fuente de datos al pie. Para solicitudes de financiamiento: problema, alcance, líneas de intervención, presupuesto, cronograma, indicadores.

## GIS y coordenadas
- **CRS de trabajo:** EPSG:32717 (UTM 17S, WGS84). Convierte lat/lon con `scripts/anin_utm.py` (pyproj, `always_xy`).
- **Validación obligatoria:** ESTE entre 450,000 y 750,000 m; NORTE entre 9,300,000 y 9,600,000 m. Un valor fuera de rango casi siempre es lat/lon invertidos, un signo perdido o una zona UTM equivocada; el script lo advierte.
- **Áreas y longitudes** se calculan en UTM 17S, nunca en grados. Reproyecta primero y reporta en hectáreas con 3 decimales (como las áreas de los bloques).
- Con `geopandas`/`shapely`/`rasterio`: confirma el CRS de cada capa al leerla (`gdf.crs`), reproyecta con `to_crs(32717)` y no asumas el CRS de un shapefile sin `.prj`.
- Controles de consistencia frente al catálogo: 117 bloques vigentes y 12,270.235 ha en total; códigos sin duplicados. Los códigos del catálogo mezclan numéricos (`27`) y `M17B10` sin guion; respétalos tal como figuran y no los «normalices» a `M#-B#` sin que el usuario lo pida.
- Mapas: cuadrícula o ticks en UTM, norte, escala gráfica, leyenda, fuente y sistema de coordenadas al pie. Para el peligro integrado usa una rampa secuencial de 5 clases (Muy Bajo a Muy Alto).
- Salidas GIS habituales: GeoPackage o shapefile (con `.prj`), Excel de atributos con coordenadas UTM, mapa en PNG/PDF para el informe.
