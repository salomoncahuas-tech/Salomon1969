# Matriz de Costos Paramétricos — Proyecto IN Piura (CUI 2669244)

Libro vigente: `Costos_Parametros_IN_Piura_v1.6.xlsx` (raíz del repositorio), generado desde
`Costos_Parametros_IN_Piura_v1.5.xlsx` (libro con las partidas a eliminar resaltadas en amarillo).

## Regenerar

```bash
python3 costos/generar_matriz_costos_v16.py Costos_Parametros_IN_Piura_v1.5.xlsx /tmp/mcp_v16.xlsx
soffice --headless --calc --convert-to xlsx --outdir /tmp/recalc /tmp/mcp_v16.xlsx
cp /tmp/recalc/mcp_v16.xlsx Costos_Parametros_IN_Piura_v1.6.xlsx
```

El paso de LibreOffice recalcula y guarda los valores de las fórmulas.

## Qué cambia en v1.6

| Tema | v1.5 | v1.6 |
|---|---|---|
| Partidas en amarillo | Pircado (03.01.03), riego con cisterna (03.02.03), riego de cama (03.03.05) y riego con balde (03.03.07) de terrazas | Eliminadas de PARTIDAS y APU_BASE (36 líneas, con 5 subpartidas que solo ellas usaban) |
| Metas marrón | Espaciamiento de las fichas ANIN sobre celdas R1/R2 | Plano de obras de conservación de suelos (06/10/2026), por bloque: hoja `METAS_MARRON` |
| Zanja | Ficha ANIN 0.40 × 0.30 m | Sección del plano 0.60 × 0.50 m (apertura y camellón × 2.5; taludes × 1.67) |
| Terraza de formación lenta | Excavación, relleno compactado, pircado, eliminación | Trazo y excavación del bordo; relleno compactado y eliminación desestimados (sustento en CU_COMP) |
| Barrera viva | Sin costo | Partidas ANIN 03.02 y 03.03 sin riego, por km de terraza y de zanja |
| O&M | Con podas y fertilización | Podas y fertilización desestimadas (años = 0) |
| Bloques con P < 400 mm | — | `P_ARIDO` (1 = obras sin barrera viva, por defecto) |
| Cárcavas | `BLOQUES!AD` vacío | Igual: meta por definir con levantamiento de campo |

`obras_planos.py` contiene las matrices M1 y M2 del plano transcritas y validadas contra los
totales impresos (119 hojas; los bloques 74 y 75 están retirados en V6 y no se costean).

## Resultado con los valores por defecto (117 bloques)

- Inversión: S/ 241,233,865
- O&M 15 años: S/ 115,399,348
- Inversión + O&M: S/ 356,633,213 (95.8 % del techo de S/ 372,442,809)
- Margen para preinversión y expediente técnico: S/ 15,809,596

## Archivos auxiliares del tablero

- `matriz_costos_engine_v16.js`: motor JavaScript del tablero «Diagnóstico Integral IN Piura V3». Reproduce al céntimo el libro v1.6 con los valores por defecto (inversión, O&M, VAC social, metas y barrera viva).
- `extraer_datos_artefacto.py`: extrae del libro recalculado a jornal de S/ 100 y S/ 120 los datos que usa el motor. Las rutas de entrada están fijadas al entorno de trabajo donde se generó: ajústelas antes de usarlo.
