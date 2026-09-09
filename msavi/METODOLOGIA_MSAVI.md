# Distribución areal del índice MSAVI 2024 por bloque — Proyecto IN Piura

**AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN**
**DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME**
**SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN - SESDI**

Proyecto IN Piura | CUI 2669244 | UTM WGS 84 Zona 17S (EPSG:32717)
Fase: Preinversión (Estudio de Perfil) · Sistema Invierte.pe

---

## 1. Objeto

Cuantificar, para cada uno de los **117 bloques preliminares de intervención**
que cuentan con ficha de Diagnóstico Territorial, la **distribución de la
superficie del bloque entre las cinco clases del índice MSAVI 2024**, y con ello
completar las hojas «Cobertura MSAVI-NDVI», «Resumen» y «Control de
consistencia» de los libros Excel del entregable.

Las cinco clases son las del propio producto cartográfico `MSAVI_Piura_2024`,
Banda 1 (Gray), y coinciden con el umbral de brecha del proyecto:

| Clase | Rango MSAVI 2024 | Interpretación | Frente al umbral 0.4976 |
|---|---|---|---|
| 5 | > 0.6139 | Vigor alto | Sobre umbral |
| 4 | 0.4976 – 0.6139 | Vigor moderado | Sobre umbral |
| 3 | 0.3813 – 0.4976 | Vigor bajo | **BAJO umbral** |
| 2 | 0.2650 – 0.3813 | Vigor muy bajo | **BAJO umbral** |
| 1 | ≤ 0.2650 | Suelo desnudo / no vegetal | **BAJO umbral** |

Umbral de brecha 0.4976 conforme a la **R.M. N.° 00213-2024-MINAM**.

## 2. Insumos

| Insumo | Origen | Uso |
|---|---|---|
| 130 cartografías temáticas `*.png` | Carpeta Drive `MSAVI BLOQUES V6` | Medición del reparto por clase |
| `BLOQUES_V6_INTERSECC_MSAVI.md` (90 175 registros) | Misma carpeta | Superficie del bloque y contraste independiente |
| `manifiesto_resumenes_117.json` | Repositorio, `datos/` | Padrón de los 117 bloques con ficha DT |
| 117 libros `Plantilla_Excel_Bloque_*.xlsx` | Drive `Plantillas Excel 117 bloques` | Destino; aportan la superficie de catálogo V5/V6 |

Los 117 bloques del padrón tienen cartografía propia: la correspondencia
`bloque → PNG` se resolvió por igualdad exacta del código, sin coincidencias
parciales ni duplicados.

## 3. Procedimiento de medición

1. **Recorte al marco cartográfico.** Todas las láminas comparten maqueta
   (3507 × 2480 px). Se recorta el interior del marco —columnas 745–3409,
   filas 99–2243— con lo que quedan fuera la leyenda, la rosa de los vientos y
   la barra de escala. Es indispensable: las cinco muestras de color de la
   leyenda, si se cuentan, saturan por sí solas las clases minoritarias.
2. **Clasificación por color exacto.** La simbología es la paleta RdYlGn de
   cinco clases; cada celda se asigna por igualdad exacta de RGB:
   `#d7191c` (1), `#fdae61` (2), `#ffffc0` (3), `#a6d96a` (4), `#1a9641` (5).
   Todo lo demás —curvas de nivel `#d79a38`, grilla, rótulos, contorno del
   bloque `#23ff23`, hidrografía— queda excluido del recuento.
3. **Separación de bloques dentro de la lámina.** 47 de los 117 bloques
   comparten lámina con bloques vecinos. Las celdas clasificadas se agrupan en
   componentes conexos (cierre morfológico de 11 px y relleno de huecos, que
   salvan los cortes de grilla y curvas de nivel). Se adjudica al bloque el
   componente que contiene el centro de la lámina y, si ninguno lo contiene, el
   de centroide más próximo: el atlas encuadra cada lámina sobre su bloque.
4. **Reparto porcentual.** El resultado por bloque es el porcentaje de celdas de
   cada clase sobre el total de celdas clasificadas del componente adjudicado.
5. **Metrado en hectáreas.** Las hectáreas **no** se derivan del tamaño en
   píxeles: se obtienen aplicando el reparto porcentual a la **superficie de
   catálogo (V5/V6)** que declara cada libro. Así la suma de las cinco clases
   iguala exactamente la superficie oficial del bloque y no se introduce una
   segunda medición de superficie que compita con el catálogo maestro.

## 4. Verificación

**a) Contra la media MSAVI de catálogo.** El bloque 38 declara media 0.6725
(clase «> 0.6139»). La media reconstruida con las marcas de clase de la
distribución medida es 0.6573 (−0.0152). La distribución sitúa el 99.55 % del
bloque sobre el umbral, coherente con «Sobre umbral».

**b) Contra la intersección vectorial, en los 117 bloques.** Se procesó
`BLOQUES_V6_INTERSECC_MSAVI.md` (90 175 registros, íntegros: el recuento y los
FID 1–90 175 cuadran con el pie del propio reporte). La correlación de Pearson
entre el porcentaje bajo umbral medido sobre la cartografía y el derivado de la
tabla es **0.949**, lo que confirma la adjudicación bloque→polígono.

**c) Contra las superficies oficiales.** La suma de los polígonos de cada bloque
en la tabla de intersección reproduce la superficie de catálogo con desviaciones
de 0.0 % a 1.0 % (bloques 38, 55, 84, 70, 2, 3, 23, M3B8), lo que valida a su
vez el padrón de bloques usado.

**d) Segmentación.** En la lámina de los bloques 6 y 36 los dos componentes
miden 169.10 ha y 24.13 ha frente a 168.41 ha y 24.84 ha de catálogo (0.4 % y
2.9 %).

## 5. Por qué no se usó la tabla de intersección para el metrado

El campo `AREA_M2` de `BLOQUES_V6_INTERSECC_MSAVI` **no es el área de cada pieza
de intersección**: conserva el área del polígono padre del bloque, repetida en
todas las filas que esa pieza generó al cruzarse con el ráster. Se comprueba
directamente: el bloque 55 tiene 9 registros con clases distintas (5, 4, 3, 4, 5,
4, 5, 4, 4) y un único valor de área, 8 594.677 m², que es la superficie del
bloque completo (0.86 ha de catálogo).

Repartir esa área a partes iguales entre las piezas —la única estimación posible
con ese campo— arroja desviaciones de 10 a 21 puntos porcentuales frente al
reparto real observado en la propia cartografía:

| Lámina | Clase | Medido sobre la cartografía | Reparto equitativo de la tabla |
|---|---|---|---|
| 6 + 36 | 0.4976–0.6139 | 46.1 % | 35.1 % |
| 10 + 16 | 0.4976–0.6139 | 48.1 % | 34.7 % |
| 23 | 0.3813–0.4976 | 17.5 % | 31.8 % |
| M3B8 + M3B9 | 0.2650–0.3813 | 55.5 % | 34.5 % |

**Recomendación.** El camino definitivo es recalcular la geometría después del
`Intersect` (`$area` sobre la capa resultante) o ejecutar directamente una
estadística zonal del ráster MSAVI 2024 recortado a `Bloques_V6`. Mientras ese
insumo no exista, la medición sobre la cartografía temática es la mejor
aproximación disponible y así queda declarada en cada libro.

## 6. Resultado agregado (117 bloques, 12 230.4 ha)

| Clase MSAVI 2024 | Superficie | % del ámbito |
|---|---:|---:|
| > 0.6139 — Vigor alto | 1 950.6 ha | 15.9 % |
| 0.4976 – 0.6139 — Vigor moderado | 1 958.1 ha | 16.0 % |
| 0.3813 – 0.4976 — Vigor bajo | 2 815.1 ha | 23.0 % |
| 0.2650 – 0.3813 — Vigor muy bajo | 4 259.6 ha | 34.8 % |
| ≤ 0.2650 — Suelo desnudo / no vegetal | 1 246.9 ha | 10.2 % |
| **BAJO el umbral 0.4976** | **8 321.6 ha** | **68.0 %** |

El 68.0 % del ámbito con ficha DT se sitúa por debajo del umbral de brecha, lo
que sustenta el indicador de brecha del proyecto (*porcentaje de superficie de
ecosistemas degradados que requieren recuperación*, R.M. N.° 00213-2024-MINAM).

## 7. Declaración de integridad de datos

Ningún valor ausente ha sido estimado ni inferido sin declararlo. La distribución
areal procede de una medición sobre el producto cartográfico del propio
entregable y así se consigna en la nota metodológica de cada libro y en la hoja
«Control de consistencia», calificada **CORREGIDO** y con el refrendo por
estadística zonal directa señalado como pendiente. Las hectáreas se anclan a la
superficie de catálogo V5/V6 declarada en cada libro.

## 8. Archivos

| Archivo | Contenido |
|---|---|
| `msavi_distribucion_117_bloques.csv` | 117 bloques × 5 clases: % y ha, superficie, % bajo/sobre umbral, clase dominante, lámina de origen |
| `analizar_msavi_png.py` | Medición reproducible sobre las cartografías PNG |
| `completar_msavi_plantillas.py` | Escribe las tres hojas en los 117 libros |
| `test_completar_msavi.py` | Prueba del anterior sobre un libro que replica la estructura |

### Uso

```bash
pip install openpyxl
python completar_msavi_plantillas.py --dir <carpeta_con_los_117_xlsx> --simular
python completar_msavi_plantillas.py --dir <carpeta_con_los_117_xlsx>
```

Deja un respaldo `.bak` de cada libro antes de escribir.
