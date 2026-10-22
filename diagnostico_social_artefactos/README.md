# Actualización del Diagnóstico Social de los artefactos provinciales

Pipeline que actualiza el componente social (fichas F-DS) de los tres volúmenes del
**Diagnóstico Territorial y Social** del Proyecto IN Piura (CUI 2669244) y aplica la
terminología de MRR-CCC del proyecto:

- Vol. I · Morropón — https://claude.ai/artifact/QSAZwL5aCbHUXFuzgArxQC
- Vol. II · Huancabamba — https://claude.ai/artifact/2MR2LbK1kpBxo45QMJKuat
- Vol. III · Ayabaca — https://claude.ai/artifact/CPGxFSAkDpK3EcSviwTxLB

## Entradas (no versionadas: contienen datos personales)

| Archivo | Origen |
|---|---|
| `respaldo.xlsx` | Respaldo del aplicativo IN Piura (`Respaldo_IN_Piura_20260926_211810.xlsx`, hoja «Diagnostico social», 275 fichas) |
| `aya_orig.html`, `hua_orig.html`, `mor_orig.html` | Versión vigente de cada artefacto antes de la actualización |

Los libros `Graficos_DS_Consolidado_<Provincia>_20260926.xlsx` del aplicativo se usaron para
contrastar las cifras de cabecera (fichas vigentes, familias, programas sociales, tenencia).

## Ejecución

```bash
python3 prep.py      # ds.pkl y JSON de datos de cada artefacto
sh build_all.sh      # hua_new.html, mor_new.html, aya_new.html
node test.js $PWD/hua_new.html   # recorre las 23 pestañas y reporta errores JS
```

## Reglas aplicadas

- **Depuración (igual en los tres volúmenes):** una F-DS-01 por centro poblado declarado
  (valores modales; en Morropón, población y familias por mediana, como en el volumen);
  F-DS-02 por ficha y actor únicos; F-DS-03 por entrevistado único. Se informa el número de
  informantes F-DS-01 de cada localidad.
- **Fragilidad y resiliencia:** las reglas de seis condiciones y seis capacidades ya
  publicadas en cada volumen, recalculadas sobre las fichas depuradas.
- **Percepción (Morropón):** codificación por palabras clave de F-DS-03 (numerales 1.1–1.3 y
  4.2); en Huancabamba las localidades nuevas se codificaron a mano con el mismo catálogo.
- **Terminología MRR-CCC:** «Infraestructura gris» pasa a «Infraestructura complementaria»; la
  infraestructura natural se separa en **verde** (revegetación, reforestación,
  enriquecimiento, regeneración natural, restauración activa) y **marrón** (zanjas de
  infiltración, terrazas de formación lenta, diques para control de cárcavas, clausura y
  manejo de pastizales, barreras y cercos vivos). Los diques en quebradas, taludes, muros,
  gaviones y drenaje son infraestructura complementaria (`term.py`).
- Ningún valor ausente se estima; no se reproducen nombres, DNI ni teléfonos de informantes.
