# Solicitud de información — Programa de Inversión del Río Piura

Respuesta a la solicitud de información actualizada sobre los avances del Perfil «Infraestructura Natural»
(CUI 2669244) para incorporar su alcance en el Programa de Inversión del Río Piura. Corte: 12/11/2026.

## Entrega (`entrega/`)

| Archivo | Contenido |
|---|---|
| `00_Informe_Remision_Avance_Perfil_IN_Piura.docx` | Informe de remisión (número de informe por definir) |
| `00_Indice_Informacion_Remitida_IN_Piura.xlsx` | Índice de los 520 archivos remitidos, con totales por fórmula |
| `01_Delimitacion_Ambito_Estudio.zip` | Área de estudio, de intervención (117 bloques, V8) y de influencia; criterios; NT 001-2026 |
| `02_Criterios_Seleccion.zip` | Meso y microlocalización: PI MCA-AHP, exclusión, idoneidad, tamizaje, campo; matriz por bloque |
| `03_Diagnostico_Poblacion_Afectada.zip` | Nota técnica, Informes E2 B4 (Vol. I–III), Matrices E2 B4, metodología e instrumentos |
| `04_Diagnostico_Unidades_Productoras.zip` | UP = ecosistemas: E6, línea base de brecha, EVAR (E8), Suelos E5, SERRN, ecosistemas |
| `05_Diagnostico_Involucrados.zip` | Actores (F-DS-02), matriz de involucrados T1.7, riesgos sociales, tamizaje predial |
| `06A_Informacion_Georreferenciada_Capas.zip` | 20 grupos de capas SHP (58 shapefiles con .prj/.cpg), 17 KMZ, 5 GPKG, procesos e inventario |
| `06B_Informacion_Georreferenciada_Mapas.zip` | Mapas temáticos (AI, ecosistemas) y 15 mapas distritales de bloques |

## Regenerar

Los insumos de Google Drive (informes E2/E5/E6/E8, matrices, capas) se descargan a una carpeta local; luego:

```bash
python generador/build_gis.py --repo . --drive <drive_raw> --matrices-e2 <drive>/Matrices_E2_IN_Piura_B4.xlsx \
    --out <paquetes>/06_Informacion_Georreferenciada_Cartografica
python generador/build_docs.py --repo . --drive <drive> --paquetes <paquetes>
python generador/empaquetar.py --paquetes <paquetes> --salida solicitud_programa_rio_piura_2026-11/entrega
```

Las cifras de las notas se leen de las fuentes (`generador/datos.py`); no se digitan.

## Por adjuntar (no incluidos)

Archivos de Drive mayores de 10 MB (límite del conector) o no ubicados: Nota Técnica de Meso y Microlocalización
(2023, 18 MB); capa de peligro por movimientos en masa del E8 (112 MB); parcelas virtuales E6 (17.9 MB);
mapas de ecosistemas de 117 bloques V7 (58 MB); mapas SERRN (71 MB); MSAVI vectorizado (13.9 MB);
proyectos QGIS y rásteres del modelamiento (por verificar); polígonos recortados del catálogo V8 (pendiente).
