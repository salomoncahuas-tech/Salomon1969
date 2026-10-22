# Pestaña «Suelos (E5)» de los diagnósticos por provincia

Genera la pestaña Estudios básicos › Suelos de los Diagnósticos Morropón, Huancabamba y Ayabaca (Versión 2)
a partir del Entregable 2 (fase de campo) del Estudio de Suelos de GeoSIG Ingenieros (datos al 20/09/2026).

- `data.py`: lee el informe (texto), las capas CALICATAS_EJECUTADAS y BLOQUES_INTERVENCION (DBF) y produce `sue.json`
  (74 calicatas, 148 chequeos, 13 pruebas de infiltración con 59 repeticiones, 118 bloques y 28 unidades de suelo muestreadas).
  El color Munsell del horizonte superficial se convierte a sRGB con `colour-science`.
- `sue_prep.js`, `sue_tab.js`: módulo que se inserta en la capa común (Integralidad) de cada volumen.
- `sue.css`, `section.html`: estilos y marcado de la pestaña.
- `assemble.py`: inyecta el módulo en cada artefacto, añade la séptima línea de evidencia (suelo limitante) a la
  convergencia y actualiza el diagrama, la Tabla 8, la cadena de evidencia, la matriz de correlación y los vacíos.

Cuando lleguen los Entregables 3 y 4 (laboratorio, CUM y Factor K), actualizar `data.py` y volver a ensamblar.
