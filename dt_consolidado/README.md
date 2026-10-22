# Diagnóstico Territorial consolidado — Proyecto IN Piura (CUI 2669244)

Generación reproducible de los tres informes técnicos consolidados de Diagnóstico
Territorial (uno por provincia, desagregado por distritos), sus anexos de matrices
en Excel y los artefactos de presentación de resultados.

**ANIN — DIME — SESDI** · Fase de preinversión (Perfil) · Sistema Invierte.pe

## Ámbito

| Provincia | Bloques | Superficie (ha) | Distritos | Microcuencas |
|---|---:|---:|---:|---:|
| Morropón | 59 | 7 327,37 | 9 | 17 |
| Huancabamba | 48 | 3 551,15 | 5 | — |
| Ayabaca | 10 | 1 391,70 | 1 | 4 |
| **Total** | **117** | **12 270,22** | **15** | — |

## Fuentes integradas

1. Las **117 plantillas** de resumen por bloque (V6 con MSAVI) — hojas Resumen,
   Cobertura MSAVI-NDVI, Estaciones fotográficas, Microcuenca y Control de consistencia.
2. El **Índice de Susceptibilidad Litológica (ISL-MM / ISL-EH)** del E3 Estudio de
   Geología, volúmenes I (Morropón), II (Huancabamba) y III (Ayabaca). Cobertura
   verificada: 117 de 117 bloques.
3. El inventario de **cárcavas codificadas** con su caracterización morfométrica
   (55 rasgos en 32 bloques).

## Uso

```bash
pip install openpyxl python-docx
python3 dt_anexo_excel.py    # 3 anexos Excel de 12 hojas
python3 dt_informe_word.py   # 3 informes Word
python3 dt_artefacto.py      # 3 artefactos HTML
```

`dt_data.py` es la capa de consolidación; ejecutarla sola imprime el resumen por
provincia como verificación.

## Principio de no invención de datos

Ningún valor ausente ha sido imputado ni estimado. Los campos no sustentados en
observación de campo, estadística zonal o catálogo oficial se conservan con su
marca de origen («Por determinar» / «Por verificar»). Las discrepancias entre
fuentes se documentan con el valor adoptado y la acción requerida, sin resolverse
por criterio del redactor.

Dos discrepancias de alcance provincial detectadas en esta consolidación:

- **D-P01 — Unidad de la pendiente del catálogo.** Las plantillas V6 rotulan el
  valor como porcentaje; el Informe DT Consolidado de Frías concluye que el mismo
  valor está en grados (ECM 1,14° frente a 12,70°; r = +0,956). Se verificó que el
  número es idéntico en ambas fuentes: el conflicto es de unidad, no de dato. Se
  publican **las dos lecturas** en paralelo. Bajo ninguna de ellas hay bloques por
  encima del criterio de idoneidad del 75 %.
- **D-M01 — Bloques sin cobertura del MDE.** Los bloques 83, 84, 85, 86 y 87
  (San Juan de Bigote, 133,57 ha) traen 0 msnm y 0 % de pendiente, que son
  marcadores de ausencia y no mediciones. Se anulan y se excluyen de todo
  promedio, mínimo y máximo.
