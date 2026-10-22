# Diagnóstico Integral IN Piura V1

Fuentes y scripts del artefacto **Diagnóstico Integral IN Piura V1**
(https://claude.ai/artifact/RRb5qHVT11J6wU6o3LnVVK). El artefacto reúne en un solo enlace:

- la **portada integral**, con estas pestañas: Panorama integral, Comparativo provincial, Los 117 bloques, Metrados y costos, Supuestos del motor;
- los tres volúmenes del diagnóstico territorial y social, sin cambios en su contenido: Vol. I Morropón, Vol. II Huancabamba y Vol. III Ayabaca, versión 5.

Proyecto IN Piura · CUI 2669244 · ANIN – DIME – SESDI · UTM WGS 84 Zona 17S.

## Contenido

| Ruta | Qué es |
|---|---|
| `fuentes/Costos_Parametros_IN_Piura_v1.3.xlsx` | Motor de costos paramétrico (18/10/2026) |
| `fuentes/vol*_v5.html` | Los tres volúmenes tal como estaban publicados |
| `src/index.tpl.html` | Plantilla de la portada integral (estructura y estilos propios) |
| `src/app.js` | Lógica de la portada: indicadores, gráficos, tablas y simulador de costos |
| `src/engine.js` | Réplica en JavaScript del motor del Excel (BLOQUES → PAQUETES → CALCULO ← CU_COMP → OYM → FLUJO) |
| `src/base.css` | Estilos comunes de los volúmenes (extraídos del Vol. II) |
| `datos/` | Datos extraídos: `data.json` (va incrustado en la portada), `costos_data.json` y `ref.json` (control) |
| `extraer_datos.py` | Extrae `datos/` y `src/base.css` desde `fuentes/` |
| `construir.py` | Genera `publicar/` (portada + tres volúmenes con la navegación integral) |
| `probar_motor.js` | Compara el motor JavaScript con el Excel |
| `archivos_publicacion.py` | Genera las listas de archivos para publicar el artefacto |

`publicar/` no se versiona: se regenera con `construir.py`.

## Cómo regenerar

```bash
cd diagnostico_integral
python extraer_datos.py        # requiere openpyxl y LibreOffice (soffice)
node probar_motor.js           # debe terminar sin diferencias
python construir.py
python archivos_publicacion.py
```

Para actualizar el artefacto, pide a Claude que publique `publicar/index.html` en la URL del artefacto en dos partes:
- con `archivos_parte1.json`: páginas de los volúmenes y láminas de los Vol. II y III;
- con `archivos_parte2.json`: láminas del Vol. I.

Las 357 láminas (mapas de PI, ecosistemas y situación predial) no están en el repositorio. Se copian en el servidor desde los artefactos originales de cada volumen.

## Motor de costos

- Con los valores por defecto, `engine.js` reproduce el Excel en los 117 bloques, el piloto (21, 66, 73, 31, M6B10) y la cartera: inversión, precios sociales, O&M de 15 años, VAC social y costo-efectividad. El error relativo máximo es de 1e-14.
- Los costos unitarios son lineales en el jornal. Se toman del Excel (S/ 100) y de una copia recalculada con LibreOffice a S/ 120, y se interpolan.
- La participación de la mano de obra no calificada en la O&M se recalcula de forma exacta para cada jornal. `probar_motor.js` lo verifica también con S/ 80.
- Desplegables del simulador:
  - **Ámbito**: cartera, piloto, provincia, distrito, microcuenca, tipología o bloque.
  - **Escenario**: jornal, densidad R1, enriquecimiento R2, espaciamiento de zanjas y de terrazas, cerco, cárcavas, gastos generales, supervisión, expediente técnico, cronograma de inversión y tasa social de descuento.
- La opción de cárcavas usa la capa de cárcavas de los volúmenes (23.08 km) como escenario. Por defecto está apagada, porque la columna AD del Excel está vacía.
- Las superficies del motor (Bloques_V7_Normalizado, 09/10/2026) pueden diferir de las de los volúmenes y se mantienen así hasta la próxima actualización. La pestaña lo advierte y muestra la comparación para cada ámbito.

## Al actualizar

- **Nueva versión del Excel**: reemplaza `fuentes/Costos_Parametros_IN_Piura_v1.3.xlsx`, ajusta el nombre en `extraer_datos.py` si cambia y vuelve a correr los pasos. Si el libro cambia de estructura (filas, hojas o nombres definidos), revisa `extraer_datos.py` y `engine.js` antes de publicar; `probar_motor.js` detecta cualquier diferencia con el Excel.
- **Nueva versión de un volumen**: reemplaza su archivo en `fuentes/`. Si cambian sus láminas, actualiza las listas de `archivos_publicacion.py`.
