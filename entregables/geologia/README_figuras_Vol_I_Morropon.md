# Inserción de figuras — E3 Geología, Volumen I: Morropón (Rev. R02)

**Proyecto IN Piura** · CUI 2669244 · ANIN – DIME – SESDI
Entregable 3 — Estudio de Geología · Fase de preinversión (perfil)

## Qué se hizo

Se insertaron las **59 figuras** de delimitación de unidades geológicas
del Volumen I, una por cada bloque de intervención de la provincia de
Morropón.

| Origen | Figuras | Cómo estaban en el documento |
| :-- | --: | :-- |
| `ATLAS DE GEOLOGIA BLOQUES V5 PIURA` (PNG) | 54 | Marcador «FIGURA PENDIENTE DE INSERCIÓN» |
| Láminas PDF de San Juan de Bigote (bloques 83–87) | 5 | Sin marcador; solo el pie de figura |
| **Total** | **59** | |

Los 54 marcadores fueron reemplazados por su mapa y eliminados; en los
5 bloques adicionales la lámina se insertó inmediatamente encima de su
pie de figura, respetando la numeración existente (9.3.6.8 a 9.3.6.12).

## Distribución por distrito

| Distrito | Bloques |
| :-- | --: |
| Buenos Aires | 7 |
| Chalaco | 2 |
| Chulucanas | 7 |
| Morropón | 2 |
| Salitral | 9 |
| San Juan de Bigote | 12 (7 del atlas + 5 adicionales) |
| Santa Catalina de Mossa | 4 |
| Santo Domingo | 10 |
| Yamango | 6 |
| **Total** | **59** |

## Criterios de composición

- **Ancho de caja:** 6.69 in (área útil A4 con los márgenes del documento).
- **Mapas del atlas** (apaisados, 1520 × 1074 px): 6.69 × 4.73 in.
- **Láminas PDF** (A4 vertical, rasterizadas a 150 dpi): 6.32 × 8.95 in,
  limitadas al 92 % del alto útil para que el pie de figura quede en la
  misma página.
- Imágenes centradas, con `keep_with_next` para que no se separen del pie.
- **Compresión:** JPEG calidad 92 sin submuestreo de croma (4:4:4). Se
  verificó a escala 1:1 que las etiquetas de unidades geológicas no
  presentan degradación. Los PNG originales son RGBA de ~2.7 MB; en JPEG
  quedan en ~0.7 MB, lo que reduce el documento de ~180 MB a 42 MB.

## Verificaciones ejecutadas

- Integridad del paquete ZIP y buen formato de los 16 XML del OOXML.
- 64 relaciones de imagen (`r:embed`) resueltas, 0 rotas.
- 0 identificadores `wp:docPr` duplicados (se renumeraron; python-docx
  numera cada imagen desde 1 y chocaba con las 5 imágenes preexistentes).
- 0 marcadores «FIGURA PENDIENTE DE INSERCIÓN» restantes.
- **Comprobación de correspondencia:** para cada uno de los 59 pies de
  figura se recuperó la imagen que lo precede y se comparó su SHA-256 con
  el del mapa esperado para ese código de bloque. Resultado: 59/59.

## Reproducir

```bash
python insertar_figuras_geologia.py \
    E3_Geologia_Vol_I_Morropon_Rev_R02.docx \
    "ATLAS DE GEOLOGIA BLOQUES V5 PIURA/" \
    "Geologia_Bloquesfaltantes/" \
    E3_Geologia_Vol_I_Morropon_Rev_R02_con_figuras.docx
```

Requiere `python-docx`, `Pillow` y `PyMuPDF`. El script identifica cada
bloque por el código del pie de figura y lo busca por nombre de archivo en
el atlas (`M6B2-1.png`, `55.png`, …) o, para los bloques adicionales, por
el número en el nombre del PDF (`SAN JUAN DE BIGOTE_BLOQUE_83.pdf`). Sirve
igual para los volúmenes II (Huancabamba) y III (Ayabaca).

## Pendiente

El índice de figuras y la paginación son campos de Word: abrir el
documento, seleccionar todo (Ctrl+E) y actualizar con F9.
