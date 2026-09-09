# Numeral 6. EQUIPO DE ESTUDIO
### Entregable 3 – Estudio de Geología · Proyecto IN Piura (CUI 2669244)
**ANIN – DIME – SESDI** · Modalidad de ejecución: Administración Directa

## Corrección aplicada (revisión hscm)

Se corrigieron los datos del segundo profesional del equipo, que figuraban con tres
errores en el cuadro 6.1, en el organigrama y en las versiones DOCX/MD del acápite:

| Campo | Decía (incorrecto) | Debe decir (correcto) |
|---|---|---|
| Nombre | Christian Camacho **Flores** | Christian Camacho **Aponte** |
| Grado / profesión | **Bach.** Christian… / **Bach. en** Ing. Agrícola | **Ing.** Christian… / **Ing. Agrícola** |
| Cargo | **Asistente SIG** | **Asistente en SIG** |

El rol según el flujograma del Entregable 3 (*Apoyo SIG-CAD*) no cambia.

La corrección está en un solo lugar —las listas `ENCABEZADO_PROF` y `ENCABEZADO_ROL`
de `contenido6.py`— del que derivan el DOCX y el MD, y en `organigrama.py` para la
figura. Así los tres formatos no pueden volver a divergir.

## Archivos del paquete

| Archivo | Contenido |
|---|---|
| `contenido6.py` | Fuente única de datos: párrafos, equipo, matriz de 27 tareas y notas |
| `estilo.py` | Primitivas de formato institucional ANIN (verde #1B4D2E, azul #1B4F72) |
| `organigrama.py` | Genera `Organigrama_E3_Geologia.png` (Figura 6.1) |
| `insertar_numeral6.py` | Reemplaza el numeral 6 dentro de un informe provincial |
| `generar_numeral6_docx.py` | Genera el acápite 6 como documento Word autónomo |
| `generar_numeral6_md.py` | Genera el acápite 6 en Markdown |
| `verificar.py` | Comprueba que el DOCX lleve los datos corregidos |

## Requisitos

    pip install python-docx Pillow

## Uso

### 1. Regenerar el organigrama (Figura 6.1)

    python3 organigrama.py Organigrama_E3_Geologia.png

Usa Arial si está instalada; si no, Liberation Sans, métricamente compatible.

### 2. Regenerar el acápite autónomo (DOCX y MD)

    python3 generar_numeral6_docx.py Organigrama_E3_Geologia.png E3_Geologia_6_Equipo_de_Estudio.docx
    python3 generar_numeral6_md.py E3_Geologia_6_Equipo_de_Estudio.md
    python3 verificar.py E3_Geologia_6_Equipo_de_Estudio.docx

### 3. Reinsertar el numeral 6 en los tres informes provinciales

Ejecutar sobre el informe R03 vigente de cada volumen, tomándolo de su carpeta en
`08_Geologia`:

    python3 insertar_numeral6.py "E3_Geologia_Vol_I_Morropon_Rev_R03_Rev_hscm.docx" \
        Organigrama_E3_Geologia.png "E3_Geologia_Vol_I_Morropon_R04.docx"

    python3 insertar_numeral6.py "E3_Geologia_Vol_II_Huancabamba_R03_Rev_hscm.docx" \
        Organigrama_E3_Geologia.png "E3_Geologia_Vol_II_Huancabamba_R04.docx"

    python3 insertar_numeral6.py "E3_Geologia_Vol_III_Ayabaca_R03_Rev_hscm.docx" \
        Organigrama_E3_Geologia.png "E3_Geologia_Vol_III_Ayabaca_R04.docx"

El script localiza el numeral 6 del cuerpo (ignora la entrada del índice), lo
reemplaza íntegro y deja intacto el numeral 7 y todo lo posterior.

## Formato aplicado

| Elemento | Formato |
|---|---|
| Títulos 6, 6.1, 6.2 | estilos `Heading 1` / `Heading 2` del propio informe |
| Encabezado de tabla | verde institucional **#1B4D2E**, texto blanco negrita |
| Bandas de ACTIVIDAD | verde **#2E6B45**, texto blanco negrita |
| Cuerpo de tabla | 6.5–7.5 pt, filas alternas **#F2F2F2**, bordes #7F9E8A / #BFCFC2 |
| Leyendas Tabla/Figura | negrita, azul institucional **#1B4F72**, 7.5 pt |
| Fuente | Arial |
| Numeración | `Tabla 6.1.` y `Figura 6.1.`, según la convención del informe |

## Nota sobre el índice

Si el informe tiene tabla de contenidos automática, actualícela en Word con
**Ctrl + E** (seleccionar todo) y **F9** después de abrir el archivo generado.
