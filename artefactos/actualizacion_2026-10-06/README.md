# Actualización 06/10/2026 de los Volúmenes I–III del Diagnóstico Territorial y Social

Capa común que se añade a los tres artefactos («Diagnóstico Morropón / Huancabamba / Ayabaca Versión 4 + Suelos») sin cambiar sus pestañas ni su arquitectura.

| Archivo | Función |
|---|---|
| `parse_ds.py` | Lee el libro «Gráficos DS Consolidado» del aplicativo (valores y %) y lo pasa a JSON por tablas. |
| `build_upd.py` | Arma `window.UPD` por provincia: consolidado DS sin nombres de informantes ni responsables, ecosistemas V6 por bloque (`mapas/salidas/IN_Piura_Ecosistemas_por_Bloque_V6.xlsx`), área de influencia (`IN_Piura_Bloques_AI_por_Distrito_V6.xlsx`) y matriz de pendientes V6 (`datos/pendientes_v6.csv`). |
| `patch.py` | Inserta la capa en el HTML publicado de cada volumen; actualiza el corte DS, el avance del E2, los textos de Integralidad, el pie de fuentes y marca en el registro de entrevistados los bloques retirados del catálogo V6. |
| `capa_0610.css` / `capa_0610.js` | Fondo azul institucional y bloques nuevos en Hallazgos sociales, Cobertura, Localidades, Servicios y economía, Actores, Riesgos y oportunidades, Discrepancias, Condicionantes, Geoespacial y Ficha por bloque. |
| `datos/` | `pendientes_v6.csv` (transcripción de *Matriz_Pendientes_Bloques_V6_IN_Piura.xlsx*, Drive, 05/10/2026; control: 117 bloques, 12,270.235 ha, 822.9 ha > 75 %) y los `upd_*.json` publicados. |

Uso: `python3 -I parse_ds.py <libro.xlsx> ds_<prov>.json` → `python3 -I build_upd.py <dir_trabajo> <repo>` → `python3 -I patch.py <dir_trabajo>` (las rutas de entrada del HTML publicado están al inicio de `patch.py`).

Avance del E2 recalculado: digitación 80/117 bloques = 68.4 %; análisis y redacción 50 % (estimación de gabinete, ajustable); E2 = 53.6 % (52.0 % al 03/10).
