# Actualización de los Volúmenes I–III del Diagnóstico Territorial y Social · corte 17/10/2026

Scripts que regeneran los artefactos «Diagnóstico Morropón / Huancabamba / Ayabaca (Copy)» con:

1. **Capa social (window.UPD)** desde el libro «Gráficos DS Consolidado» VAL_PCT del 17/10/2026
   (`parse_ds.py` → `build_upd.py`). Conserva las convenciones del corte 07/10 (modo de % por tabla,
   columnas calculadas) y **elimina las columnas con nombres de personas** que el libro nuevo incluye
   (Nombre del actor, Responsable, Entrevistado/a; nombre del responsable en «Ficha de referencia»),
   conforme a la Ley N.° 29733. Validado: Ayabaca (sin fichas nuevas) reproduce exactamente el UPD del 07/10.
2. **Pestaña «Ecosistemas»** con los datos del Tablero E6 (estado de los ecosistemas por bloque).
3. **Pestaña «Tamizaje predial»** con el índice de mapas de situación predial v2 (incluido aquí) y las láminas
   SP-001 a SP-117 de Drive (SP-059, bloque 5, no estaba entre los archivos extraídos).
4. **Integración** (`build_x.py`, `capa_eco_tam.js`): cruce por bloque PI (clase del índice `ipp`) × estado E6
   × situación predial × tenencia declarada en la F-DS-01; paneles en «Ficha por bloque», «Integralidad»,
   «Consistencia», «Geoespacial» y «Riesgos y oportunidades»; avance del E2 y vacíos de integración.

Estructura esperada del directorio de trabajo: `src/{A,H,M}.html` (artefacto leído), `xl/*.xlsx` (consolidados),
`work/E6_D.json` (datos del Tablero E6), `work/SP_index.json`, `sp/` y `sp2/` (láminas). Orden:
`build_upd.py` → `build_x.py` → `assemble.py`; pruebas con `check.js` e `inter.js` (Playwright).
`nov.py` contiene las novedades redactadas a partir de `diffs.py`; si cambia el consolidado, deben revisarse.
