# Python desde cero: curso autodidacta

Material para aprender Python **sin conocimientos previos de programación**, con ejemplos del trabajo forestal, ambiental y SIG.

| | |
|---|---|
| Módulos | 15 (M00–M14) |
| Lecciones | 79, cada una con explicación y ejemplo ejecutable |
| Ejercicios | 60, con pista, comprobación automática y solución |
| Autoevaluación | 60 preguntas (4 por módulo) con explicación |
| Dedicación | ≈ 37.5 h (8 semanas a unas 5 h por semana) |

## Contenido de la carpeta

```
curso_python/
├── curso_python_interactivo.html   ← curso completo; ejecuta Python en el navegador
├── cuadernos/                      ← un cuaderno Jupyter (.ipynb) por módulo
├── soluciones/                     ← soluciones comprobadas de cada módulo
└── _fuente/                        ← contenido y script que genera todo lo anterior
```

## Por dónde empezar

1. **Abra `curso_python_interactivo.html`** con doble clic (Chrome o Edge). Necesita internet la primera vez para cargar el intérprete de Python. Empiece por el Módulo 0.
2. Cuando llegue al **Módulo 13** (o antes, si lo prefiere), instale **Anaconda** como explica el Módulo 0 y abra los cuadernos de la carpeta `cuadernos/` con Jupyter.
3. Alternativa sin instalar: suba los cuadernos `.ipynb` a **Google Colab** (*Archivo → Subir cuaderno*).

### Cómo trabajar cada módulo

1. Lea cada lección y ejecute el ejemplo. Luego **modifíquelo**: cambie valores y vuelva a ejecutar.
2. Resuelva los ejercicios y pulse **Comprobar** (en Jupyter: ejecute la celda de comprobación).
3. Si se atasca, abra la **pista**. Consulte la **solución** solo después de intentarlo.
4. Responda la **autoevaluación**. Si falla una pregunta, repase la lección.

## Plan de estudio sugerido

| Semana | Módulos | Horas |
|---|---|---|
| 1 | M00 ¿Qué es programar? · M01 Primeros pasos · M02 Variables y tipos de datos | 5 |
| 2 | M03 Comparaciones y lógica · M04 Texto · M05 Decisiones (if) | 5.5 |
| 3 | M06 Bucles for y while · M07 Listas y tuplas | 5 |
| 4 | M08 Diccionarios y conjuntos · M09 Funciones | 5.5 |
| 5 | M10 Errores y módulos · M11 Archivos CSV y JSON | 5.5 |
| 6 | M12 Clases y objetos | 3 |
| 7 | M13 Análisis de datos con NumPy, pandas y matplotlib | 4 |
| 8 | M14 Python aplicado a SIG y proyecto integrador | 4 |

## Módulos

| Código | Tema | Dónde se ejecuta |
|---|---|---|
| M00 | Antes de empezar: qué es programar, dónde escribir código, cómo leer errores | Navegador / Jupyter |
| M01 | `print()`, comentarios y operaciones aritméticas | Navegador / Jupyter |
| M02 | Variables, tipos de datos y conversiones | Navegador / Jupyter |
| M03 | Comparaciones y operadores lógicos | Navegador / Jupyter |
| M04 | Texto: f-strings, índices, métodos, `input()` | Navegador / Jupyter |
| M05 | Decisiones: `if`, `elif`, `else` | Navegador / Jupyter |
| M06 | Bucles `for` y `while` | Navegador / Jupyter |
| M07 | Listas, tuplas y comprensiones | Navegador / Jupyter |
| M08 | Diccionarios y conjuntos | Navegador / Jupyter |
| M09 | Funciones | Navegador / Jupyter |
| M10 | Errores, excepciones y módulos | Navegador / Jupyter |
| M11 | Archivos TXT, CSV y JSON | Navegador (ejercicios) / Jupyter (lectura y escritura de archivos) |
| M12 | Clases y objetos | Navegador / Jupyter |
| M13 | NumPy, pandas y matplotlib | Jupyter / Colab |
| M14 | Coordenadas UTM 17S (EPSG:32717), distancias, áreas y proyecto integrador | Navegador / Jupyter (pyproj solo en Jupyter) |

## Notas

- Los datos de los ejemplos son **ficticios** y sirven solo para practicar. El factor de forma 0.5 que se usa para calcular volúmenes es un valor supuesto para los ejercicios, no un valor de referencia.
- El HTML ejecuta Python con **Brython**, que funciona dentro del navegador. Algunos mensajes de error se redactan distinto que en Python oficial. Un programa que tarde más de 10 segundos se detiene solo (protección contra bucles infinitos).
- El avance (ejercicios resueltos y respuestas) se guarda **solo en el navegador** donde se abrió el curso.

## Regenerar los materiales

El contenido está en `_fuente/contenido_a.py`, `contenido_b.py` y `contenido_c.py`. Después de editarlo:

```bash
pip install nbformat markdown pygments numpy pandas matplotlib pyproj
python _fuente/construir.py
```

El script ejecuta todos los ejemplos, comprueba que cada solución pase su prueba (y que ninguna plantilla vacía la pase), y vuelve a generar los cuadernos, las soluciones y el HTML.
