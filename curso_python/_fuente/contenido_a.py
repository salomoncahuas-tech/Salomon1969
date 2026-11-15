# -*- coding: utf-8 -*-
"""Contenido del curso: módulos 0 a 4 (fundamentos)."""

MODULOS = []

# ---------------------------------------------------------------------------
MODULOS.append({
"id": "M00",
"titulo": "Antes de empezar: ¿qué es programar?",
"duracion": "1 h",
"objetivos": [
    "Entender qué es un programa y qué es un lenguaje de programación.",
    "Conocer qué es Python y para qué sirve en el trabajo técnico.",
    "Saber dónde escribir y ejecutar código (sin instalar nada y con Jupyter).",
    "Perderle el miedo a los mensajes de error.",
],
"lecciones": [
{"titulo": "¿Qué es programar?",
"texto": r'''Programar es **escribir instrucciones para que una computadora las ejecute**.

Piense en un **protocolo de campo** para medir una parcela: "1) ubicar el vértice A, 2) medir 20 m al norte, 3) colocar la estaca…". Si el protocolo es claro y está en orden, cualquier brigada obtiene el mismo resultado.

Un programa es lo mismo, con una diferencia importante: **la computadora no interpreta ni adivina**. Hace exactamente lo que se le dice, al pie de la letra. Por eso los lenguajes de programación tienen reglas estrictas de escritura, llamadas **sintaxis**.

| Concepto | En la vida diaria | En programación |
|---|---|---|
| Instrucción | "Mida el DAP" | `print("Hola")` |
| Secuencia | Pasos del protocolo en orden | Líneas de código de arriba hacia abajo |
| Programa | Protocolo completo | Archivo `.py` o cuaderno `.ipynb` |
| Error | Paso ambiguo o mal escrito | Mensaje de error de Python |
'''},
{"titulo": "¿Qué es Python y por qué aprenderlo?",
"texto": r'''**Python** es un lenguaje de programación gratuito, de código abierto y fácil de leer. Fue diseñado para parecerse al inglés sencillo.

Se usa ampliamente en ciencia, ingeniería y sistemas de información geográfica:

| Herramienta | Uso típico |
|---|---|
| **pandas** | Tablas de datos (como Excel, pero automatizado) |
| **openpyxl** | Crear y editar archivos Excel |
| **matplotlib** | Gráficos |
| **pyproj / geopandas** | Coordenadas y capas geográficas |
| **PyQGIS / arcpy** | Automatizar QGIS y ArcGIS |

Python es un lenguaje **interpretado**: un programa llamado *intérprete* lee el código línea por línea y lo ejecuta de inmediato. Esto permite probar ideas rápidamente.
'''},
{"titulo": "¿Dónde escribo el código?",
"texto": r'''Tiene tres opciones. **Recomendación:** empiece con la opción A y pase a la C cuando llegue al Módulo 13.

**A. Curso interactivo (HTML) — sin instalar nada.** Cada ejemplo tiene un botón **Ejecutar** y cada ejercicio un botón **Comprobar**. Funciona para los módulos 0 a 12 y parte del 14.

**B. Google Colab** (colab.research.google.com). Funciona en el navegador con una cuenta de Google. Permite abrir los cuadernos `.ipynb` de este curso (*Archivo → Subir cuaderno*).

**C. Anaconda + Jupyter (instalación local, recomendada para trabajo real).**

1. Ingrese a `anaconda.com/download` y descargue el instalador para su sistema (Windows 64 bits).
2. Ejecute el instalador y acepte las opciones por defecto.
3. Abra **Anaconda Navigator** desde el menú Inicio y pulse **Launch** en *Jupyter Notebook* (o *JupyterLab*).
4. Se abrirá el navegador con sus carpetas: ubique la carpeta del curso y abra `M00_....ipynb`.

**Cómo se usa un cuaderno Jupyter:**

- El cuaderno está formado por **celdas**. Hay celdas de **texto** (como esta) y celdas de **código**.
- Para ejecutar una celda de código: haga clic en ella y pulse **Shift + Enter**.
- El resultado aparece debajo de la celda.
- Ejecute las celdas **en orden**, de arriba hacia abajo.
- Si algo se comporta raro: menú *Kernel → Restart & Run All* (reinicia y ejecuta todo).
'''},
{"titulo": "Su primera línea de código",
"texto": r'''La instrucción más básica es `print()`, que **muestra en pantalla** lo que se coloca entre paréntesis.

- `print` es el nombre de la instrucción (en Python se llama **función**).
- Los paréntesis `( )` encierran lo que se le entrega a la función.
- Las comillas `" "` indican que se trata de un **texto**.

Ejecute la celda y observe el resultado.''',
"codigo": r'''print("Hola, mundo")
print("Estoy aprendiendo Python")'''},
{"titulo": "Los errores son normales (y útiles)",
"texto": r'''Todos los programadores cometen errores todos los días. Python avisa con un **mensaje de error**. Lea siempre **la última línea**: dice el *tipo* de error y una pista.

En el siguiente ejemplo falta cerrar las comillas. Ejecútelo a propósito y lea el mensaje.

**Lista de verificación ante un error:**

1. ¿Cerré todas las comillas `" "` y paréntesis `( )`?
2. ¿Escribí bien los nombres? Python distingue **mayúsculas y minúsculas** (`print` ≠ `Print`).
3. ¿Usé punto decimal (`3.5`) y no coma (`3,5`)?
4. ¿Respeté la sangría (espacios al inicio de la línea)? Lo verá en el Módulo 5.''',
"codigo": r'''print("Hola, mundo)''',
"error": True},
],
"ejercicios": [
{"titulo": "Su primer mensaje", "nivel": 1,
"enunciado": "Guarde el texto `Hola, soy aprendiz de Python` en `mensaje` (ya está escrito el inicio) y muéstrelo con `print`.",
"plantilla": r'''mensaje = "Escriba aquí el texto"
print(mensaje)''',
"pista": "Reemplace solo lo que está entre comillas. Respete mayúsculas, la coma y los espacios.",
"solucion": r'''mensaje = "Hola, soy aprendiz de Python"
print(mensaje)''',
"prueba": r'''assert mensaje == "Hola, soy aprendiz de Python", "El texto no coincide exactamente: revise mayúsculas, coma y espacios."'''},
{"titulo": "Corrija el error", "nivel": 1,
"enunciado": "El siguiente código tiene un error de sintaxis. Ejecútelo, lea el mensaje y corríjalo.",
"plantilla": r'''frase = "Python es fácil de aprender
print(frase)''',
"pista": "Mire dónde empieza el texto y dónde debería terminar. ¿Falta algún símbolo?",
"solucion": r'''frase = "Python es fácil de aprender"
print(frase)''',
"prueba": r'''assert frase == "Python es fácil de aprender", "Revise el texto de la frase."'''},
{"titulo": "Python como calculadora", "nivel": 1,
"enunciado": "En un vivero se produjeron 125 plantones en enero y 375 en febrero. Calcule el total y guárdelo en `total`.",
"plantilla": r'''total = 0   # reemplace el 0 por la operación
print(total)''',
"pista": "La suma se escribe con el signo +, igual que en una calculadora: `125 + 375`.",
"solucion": r'''total = 125 + 375
print(total)''',
"prueba": r'''assert total == 500, f"Se esperaba 500 y se obtuvo {total}."'''},
{"titulo": "Horas de estudio", "nivel": 1,
"enunciado": "Si estudia 5 horas por semana durante 6 semanas, ¿cuántas horas estudiará? Guarde el resultado en `horas_totales`.",
"plantilla": r'''horas_totales = 0   # use el signo * para multiplicar
print(horas_totales)''',
"pista": "La multiplicación se escribe con asterisco: `5 * 6`.",
"solucion": r'''horas_totales = 5 * 6
print(horas_totales)''',
"prueba": r'''assert horas_totales == 30, f"Se esperaba 30 y se obtuvo {horas_totales}."'''},
],
"quiz": [
{"pregunta": "¿Qué es un programa?",
 "opciones": ["Un conjunto de instrucciones que la computadora ejecuta", "Un archivo de Excel", "Un tipo de computadora", "Una página web"],
 "correcta": 0, "explicacion": "Un programa es una secuencia de instrucciones escritas en un lenguaje que la computadora puede ejecutar."},
{"pregunta": "En Jupyter, ¿cómo se ejecuta una celda de código?",
 "opciones": ["Enter", "Ctrl + S", "Shift + Enter", "Doble clic"],
 "correcta": 2, "explicacion": "Shift + Enter ejecuta la celda y pasa a la siguiente."},
{"pregunta": "Ante un mensaje de error, ¿qué parte conviene leer primero?",
 "opciones": ["La primera línea", "La última línea", "Ninguna, se reinicia la computadora", "Solo el número de línea"],
 "correcta": 1, "explicacion": "La última línea indica el tipo de error y una descripción breve del problema."},
{"pregunta": "¿Cuál de estas instrucciones es correcta?",
 "opciones": ["Print(\"Hola\")", "print(\"Hola\")", "print(Hola\")", "PRINT(\"Hola\")"],
 "correcta": 1, "explicacion": "Python distingue mayúsculas: la función se llama print, en minúsculas, y el texto va entre comillas."},
],
"glosario": {
 "Programa": "Conjunto de instrucciones que la computadora ejecuta en orden.",
 "Sintaxis": "Reglas de escritura de un lenguaje de programación.",
 "Intérprete": "Programa que lee y ejecuta código Python línea por línea.",
 "Celda": "Bloque de un cuaderno Jupyter; puede contener texto o código.",
 "Función": "Instrucción con nombre que realiza una tarea, como print().",
},
})

# ---------------------------------------------------------------------------
MODULOS.append({
"id": "M01",
"titulo": "Primeros pasos: mostrar resultados y hacer cálculos",
"duracion": "2 h",
"objetivos": [
    "Usar print() para mostrar textos y números.",
    "Escribir comentarios para documentar el código.",
    "Realizar operaciones aritméticas y conocer el orden de las operaciones.",
],
"lecciones": [
{"titulo": "La función print() en detalle",
"texto": r'''`print()` puede mostrar **textos** (entre comillas) y **números** (sin comillas). Si se le entregan varios valores separados por **comas**, los muestra en la misma línea con un espacio entre ellos.

- `sep=` cambia el separador entre valores.
- `print()` vacío imprime una línea en blanco.''',
"codigo": r'''print("Parcela", 5)
print("Área:", 2.5, "ha")
print("Norte", "Sur", "Este", sep=" | ")
print()
print("Fin del reporte")'''},
{"titulo": "Comentarios: notas para personas",
"texto": r'''Todo lo que va después del símbolo `#` es un **comentario**: Python lo ignora. Sirve para explicar qué hace el código, a usted mismo en el futuro o a sus colegas.

Buena práctica: comente el **porqué**, no lo obvio.''',
"codigo": r'''# Cálculo de la superficie de una parcela rectangular de 20 m x 50 m
print(20 * 50)   # resultado en metros cuadrados'''},
{"titulo": "Python como calculadora",
"texto": r'''| Operación | Símbolo | Ejemplo | Resultado |
|---|---|---|---|
| Suma | `+` | `7 + 2` | `9` |
| Resta | `-` | `7 - 2` | `5` |
| Multiplicación | `*` | `7 * 2` | `14` |
| División | `/` | `7 / 2` | `3.5` |
| División entera | `//` | `7 // 2` | `3` |
| Resto (módulo) | `%` | `7 % 2` | `1` |
| Potencia | `**` | `7 ** 2` | `49` |

**Orden de las operaciones** (igual que en matemáticas): primero paréntesis, luego potencias, luego multiplicación y división, y al final suma y resta. Ante la duda, **use paréntesis**.''',
"codigo": r'''print(7 + 2, 7 - 2, 7 * 2, 7 / 2)
print(7 // 2, 7 % 2, 7 ** 2)
print(2 + 3 * 4)      # primero la multiplicación
print((2 + 3) * 4)    # primero el paréntesis'''},
{"titulo": "Números con decimales",
"texto": r'''En Python el separador decimal es el **punto**, nunca la coma: `3.75`.

- `round(valor, n)` redondea a *n* decimales.
- Las computadoras guardan los decimales en binario, por eso a veces aparecen resultados como `0.30000000000000004`. Es normal; se resuelve redondeando al mostrar.''',
"codigo": r'''print(0.1 + 0.2)
print(round(0.1 + 0.2, 2))
print(round(3.14159, 3))
print(10 / 4)'''},
],
"ejercicios": [
{"titulo": "Superficie de una parcela", "nivel": 1,
"enunciado": "Una parcela mide 20 m por 50 m. Calcule su superficie en m² (`area_m2`) y en hectáreas (`area_ha`). Recuerde: 1 ha = 10 000 m².",
"plantilla": r'''area_m2 = 0
area_ha = 0
print(area_m2, "m2")
print(area_ha, "ha")''',
"pista": "`area_ha` se obtiene dividiendo `area_m2` entre 10000. Puede usar la variable ya calculada.",
"solucion": r'''area_m2 = 20 * 50
area_ha = area_m2 / 10000
print(area_m2, "m2")
print(area_ha, "ha")''',
"prueba": r'''assert area_m2 == 1000, "La superficie en m2 debe ser 1000."
assert abs(area_ha - 0.1) < 1e-9, "La superficie en ha debe ser 0.1."'''},
{"titulo": "Altura promedio", "nivel": 1,
"enunciado": "Se midieron tres árboles: 12.5 m, 14.0 m y 13.6 m. Calcule la altura promedio en `promedio`.",
"plantilla": r'''promedio = 0
print(round(promedio, 2))''',
"pista": "Sume las tres alturas **dentro de paréntesis** y luego divida entre 3.",
"solucion": r'''promedio = (12.5 + 14.0 + 13.6) / 3
print(round(promedio, 2))''',
"prueba": r'''assert abs(promedio - 13.3667) < 0.001, "Revise los paréntesis: primero se suma y luego se divide."'''},
{"titulo": "Cajas de plantones", "nivel": 2,
"enunciado": "Hay 1100 plantones y cada caja lleva 40. Calcule cuántas cajas se llenan completas (`cajas_llenas`) y cuántos plantones sobran (`sobrantes`).",
"plantilla": r'''cajas_llenas = 0
sobrantes = 0
print(cajas_llenas, sobrantes)''',
"pista": "Use la división entera `//` para las cajas y el resto `%` para los sobrantes.",
"solucion": r'''cajas_llenas = 1100 // 40
sobrantes = 1100 % 40
print(cajas_llenas, sobrantes)''',
"prueba": r'''assert cajas_llenas == 27, "Las cajas llenas deben ser 27."
assert sobrantes == 20, "Deben sobrar 20 plantones."'''},
{"titulo": "De Celsius a Fahrenheit", "nivel": 2,
"enunciado": "Convierta 15 °C a grados Fahrenheit con la fórmula F = C × 9/5 + 32. Guarde el resultado en `fahrenheit`.",
"plantilla": r'''celsius = 15
fahrenheit = 0
print(fahrenheit)''',
"pista": "Escriba la fórmula usando la variable: `celsius * 9 / 5 + 32`.",
"solucion": r'''celsius = 15
fahrenheit = celsius * 9 / 5 + 32
print(fahrenheit)''',
"prueba": r'''assert abs(fahrenheit - 59) < 1e-9, "15 °C equivalen a 59 °F."'''},
],
"quiz": [
{"pregunta": "¿Cuál es el resultado de `7 / 2`?", "opciones": ["3", "3.5", "1", "4"], "correcta": 1,
 "explicacion": "La división / siempre devuelve un número decimal: 3.5."},
{"pregunta": "¿Cuál es el resultado de `7 // 2`?", "opciones": ["3.5", "1", "3", "4"], "correcta": 2,
 "explicacion": "// es la división entera: descarta la parte decimal."},
{"pregunta": "¿Cuál es el resultado de `2 + 3 * 4`?", "opciones": ["20", "14", "24", "9"], "correcta": 1,
 "explicacion": "Primero se multiplica (3 × 4 = 12) y luego se suma 2."},
{"pregunta": "¿Qué hace el símbolo # en Python?", "opciones": ["Multiplica", "Inicia un comentario", "Define un número", "Muestra un resultado"], "correcta": 1,
 "explicacion": "Todo lo que sigue a # en la línea es un comentario y Python lo ignora."},
],
"glosario": {
 "Comentario": "Texto después de # que Python ignora; sirve para documentar.",
 "Operador": "Símbolo que realiza una operación, como +, -, * o /.",
 "Módulo (operador %)": "Resto de una división entera.",
},
})

# ---------------------------------------------------------------------------
MODULOS.append({
"id": "M02",
"titulo": "Variables y tipos de datos",
"duracion": "2 h",
"objetivos": [
    "Crear variables y asignarles valores.",
    "Nombrar variables correctamente.",
    "Reconocer los tipos de datos básicos: int, float, str, bool y None.",
    "Convertir valores de un tipo a otro.",
],
"lecciones": [
{"titulo": "¿Qué es una variable?",
"texto": r'''Una **variable** es un **nombre que guarda un valor**, como una etiqueta pegada a una caja. Se crea con el signo `=`, que en Python significa **"guardar"** (asignar), no "es igual a".

```
nombre_de_la_variable = valor
```

Una variable puede cambiar de valor: el nuevo valor reemplaza al anterior.''',
"codigo": r'''altura = 12.5          # guardo 12.5 en "altura"
print(altura)
altura = 13.0          # reemplazo el valor
print(altura)
altura_en_cm = altura * 100
print(altura_en_cm)'''},
{"titulo": "Reglas para nombrar variables",
"texto": r'''| Regla | Correcto | Incorrecto |
|---|---|---|
| Solo letras, números y guion bajo `_` | `dap_cm` | `dap-cm` |
| No puede empezar con número | `parcela1` | `1parcela` |
| Sin espacios | `area_total` | `area total` |
| Distingue mayúsculas | `Area` y `area` son distintas | — |
| No usar palabras reservadas | `clase_suelo` | `class`, `if`, `for` |

**Convención en Python (snake_case):** minúsculas y palabras separadas por guion bajo, con nombres que describan el contenido: `precipitacion_mm` es mejor que `p`. Evite tildes y la ñ en los nombres (son válidas, pero generan problemas al compartir código).'''},
{"titulo": "Tipos de datos básicos",
"texto": r'''Cada valor tiene un **tipo**. La función `type()` dice cuál es.

| Tipo | Nombre | Ejemplo |
|---|---|---|
| `int` | Entero | `25` |
| `float` | Decimal | `12.5` |
| `str` | Texto (cadena) | `"Algarrobo"` |
| `bool` | Lógico (verdadero/falso) | `True`, `False` |
| `NoneType` | Ausencia de valor | `None` |

`None` es útil para representar un dato **por determinar** (por ejemplo, una medición que aún no se hizo).''',
"codigo": r'''especie = "Prosopis pallida"
n_arboles = 25
dap_cm = 32.5
es_nativa = True
altura_comercial = None   # dato por determinar

print(type(especie))
print(type(n_arboles))
print(type(dap_cm))
print(type(es_nativa))
print(type(altura_comercial))'''},
{"titulo": "Convertir entre tipos",
"texto": r'''Es frecuente recibir números **como texto** (por ejemplo, desde un formulario o archivo). Para operar con ellos hay que convertirlos:

- `int("45")` → `45` (entero)
- `float("45.8")` → `45.8` (decimal)
- `str(45)` → `"45"` (texto)

Observe la diferencia: `"5" + "3"` **une textos** (`"53"`), mientras que `5 + 3` **suma** (`8`). Sumar un texto con un número produce un `TypeError`.''',
"codigo": r'''texto = "45.8"
numero = float(texto)
print(numero + 10)

print("5" + "3")          # une textos
print(int("5") + int("3")) # suma números
print("Parcela " + str(7)) # convierto el número a texto para unir'''},
{"titulo": "Atajos de asignación",
"texto": r'''- `x += 5` equivale a `x = x + 5` (también existen `-=`, `*=`, `/=`).
- Se pueden asignar varias variables en una línea: `a, b = 1, 2`.
- Intercambiar valores: `a, b = b, a`.''',
"codigo": r'''contador = 0
contador += 10    # contador = contador + 10
contador -= 3
print(contador)

este, norte = 545120, 9425300   # coordenadas de ejemplo
print(este, norte)'''},
],
"ejercicios": [
{"titulo": "Ficha de un árbol", "nivel": 1,
"enunciado": "Cree cuatro variables: `especie` con el texto `Algarrobo`, `dap_cm` con 32.5, `altura_m` con el entero 8 y `es_nativa` con True.",
"plantilla": r'''# Escriba sus variables aquí

print(especie, dap_cm, altura_m, es_nativa)''',
"pista": "Los textos van entre comillas; True se escribe con T mayúscula y sin comillas.",
"solucion": r'''especie = "Algarrobo"
dap_cm = 32.5
altura_m = 8
es_nativa = True
print(especie, dap_cm, altura_m, es_nativa)''',
"prueba": r'''assert especie == "Algarrobo", "especie debe ser el texto 'Algarrobo'."
assert dap_cm == 32.5 and isinstance(dap_cm, float), "dap_cm debe ser el decimal 32.5."
assert altura_m == 8 and isinstance(altura_m, int), "altura_m debe ser el entero 8."
assert es_nativa is True, "es_nativa debe ser True (sin comillas)."'''},
{"titulo": "Texto a número", "nivel": 1,
"enunciado": "Un formulario entregó la precipitación como texto: `\"125.4\"`. Conviértala a decimal en `lluvia_mm` y súmele 10.6 en `lluvia_total`.",
"plantilla": r'''dato = "125.4"
lluvia_mm = dato          # convierta aquí
lluvia_total = 0
print(lluvia_total)''',
"pista": "Use `float(dato)`.",
"solucion": r'''dato = "125.4"
lluvia_mm = float(dato)
lluvia_total = lluvia_mm + 10.6
print(lluvia_total)''',
"prueba": r'''assert isinstance(lluvia_mm, float), "lluvia_mm debe ser un float."
assert abs(lluvia_total - 136.0) < 1e-9, "lluvia_total debe ser 136.0."'''},
{"titulo": "Contador de árboles", "nivel": 1,
"enunciado": "Empiece con `arboles = 0`. Sume 15 (día 1), sume 23 (día 2) y reste 4 (árboles mal registrados) usando `+=` y `-=`.",
"plantilla": r'''arboles = 0
# día 1

# día 2

# corrección

print(arboles)''',
"pista": "`arboles += 15` suma 15 al valor actual.",
"solucion": r'''arboles = 0
arboles += 15
arboles += 23
arboles -= 4
print(arboles)''',
"prueba": r'''assert arboles == 34, f"Se esperaba 34 y se obtuvo {arboles}."'''},
{"titulo": "Intercambio de valores", "nivel": 2,
"enunciado": "Por error se registró `este = 9425300` y `norte = 545120`. Intercambie los valores en una sola línea.",
"plantilla": r'''este = 9425300
norte = 545120
# intercambie aquí

print(este, norte)''',
"pista": "`a, b = b, a` intercambia dos variables.",
"solucion": r'''este = 9425300
norte = 545120
este, norte = norte, este
print(este, norte)''',
"prueba": r'''assert este == 545120 and norte == 9425300, "Los valores no se intercambiaron."'''},
],
"quiz": [
{"pregunta": "¿Qué significa `x = 5` en Python?", "opciones": ["x es igual a 5 (comparación)", "Guardar 5 en la variable x", "Multiplicar x por 5", "Error"], "correcta": 1,
 "explicacion": "El signo = asigna (guarda) un valor en una variable."},
{"pregunta": "¿Cuál es un nombre de variable válido?", "opciones": ["2area", "area total", "area_total", "area-total"], "correcta": 2,
 "explicacion": "Solo letras, números y guion bajo, y no puede empezar con número."},
{"pregunta": "¿Qué tipo tiene el valor `\"12\"`?", "opciones": ["int", "float", "str", "bool"], "correcta": 2,
 "explicacion": "Está entre comillas, por lo tanto es un texto (str)."},
{"pregunta": "¿Cuál es el resultado de `\"5\" + \"3\"`?", "opciones": ["8", "\"53\"", "Error", "15"], "correcta": 1,
 "explicacion": "Con textos, + une (concatena) en lugar de sumar."},
],
"glosario": {
 "Variable": "Nombre que guarda un valor en la memoria.",
 "Asignación": "Acción de guardar un valor en una variable con =.",
 "int / float": "Números enteros / números con decimales.",
 "str": "Cadena de texto, siempre entre comillas.",
 "bool": "Valor lógico: True o False.",
 "None": "Valor especial que indica ausencia de dato.",
},
})

# ---------------------------------------------------------------------------
MODULOS.append({
"id": "M03",
"titulo": "Comparaciones y lógica",
"duracion": "1.5 h",
"objetivos": [
    "Comparar valores y obtener resultados verdadero/falso.",
    "Combinar condiciones con and, or y not.",
    "Preparar la base para tomar decisiones (Módulo 5).",
],
"lecciones": [
{"titulo": "Operadores de comparación",
"texto": r'''Una comparación siempre responde **True** (verdadero) o **False** (falso).

| Operador | Significado | Ejemplo | Resultado |
|---|---|---|---|
| `==` | igual a | `5 == 5` | `True` |
| `!=` | distinto de | `5 != 3` | `True` |
| `>` | mayor que | `5 > 8` | `False` |
| `<` | menor que | `5 < 8` | `True` |
| `>=` | mayor o igual | `5 >= 5` | `True` |
| `<=` | menor o igual | `4 <= 3` | `False` |

**Atención:** `=` guarda un valor; `==` compara. Confundirlos es el error más común al empezar.''',
"codigo": r'''dap_cm = 12.3
print(dap_cm >= 10)      # ¿el árbol supera el diámetro mínimo de 10 cm?
print(dap_cm == 12.3)
print("algarrobo" == "Algarrobo")   # mayúsculas cuentan'''},
{"titulo": "Operadores lógicos: and, or, not",
"texto": r'''- `A and B` → True solo si **ambas** son verdaderas.
- `A or B` → True si **al menos una** es verdadera.
- `not A` → invierte el valor.

| A | B | A and B | A or B |
|---|---|---|---|
| True | True | True | True |
| True | False | False | True |
| False | True | False | True |
| False | False | False | False |''',
"codigo": r'''pendiente = 35        # porcentaje
lluvia_mm = 620       # anual (valores de ejemplo)

print(pendiente > 30 and lluvia_mm > 500)
print(pendiente > 50 or lluvia_mm > 500)
print(not pendiente > 30)'''},
{"titulo": "Comparaciones encadenadas",
"texto": r'''Python permite escribir rangos de forma natural: `15 <= x <= 30` equivale a `x >= 15 and x <= 30`.''',
"codigo": r'''temperatura = 24
print(18 <= temperatura <= 28)   # ¿está en el rango?
print(0 < -5 < 10)'''},
{"titulo": "Valores que cuentan como verdadero o falso",
"texto": r'''Cualquier valor puede evaluarse como lógico con `bool()`:

- Se consideran **falsos**: `0`, `0.0`, `""` (texto vacío), `None`, listas vacías.
- Todo lo demás es **verdadero**.

Esto será útil para verificar si un campo está vacío.''',
"codigo": r'''print(bool(0), bool(15))
print(bool(""), bool("Faique"))
print(bool(None))'''},
],
"ejercicios": [
{"titulo": "¿Se inventaría el árbol?", "nivel": 1,
"enunciado": "Con `dap_cm = 8.7`, guarde en `inventariable` si el árbol tiene un DAP mayor o igual a 10 cm (criterio de ejemplo).",
"plantilla": r'''dap_cm = 8.7
inventariable = None   # escriba la comparación
print(inventariable)''',
"pista": "La comparación `dap_cm >= 10` ya da True o False; guárdela directamente.",
"solucion": r'''dap_cm = 8.7
inventariable = dap_cm >= 10
print(inventariable)''',
"prueba": r'''assert inventariable is False, "Con 8.7 cm el resultado debe ser False."'''},
{"titulo": "Dos condiciones a la vez", "nivel": 1,
"enunciado": "Guarde en `requiere_zanjas` si la pendiente es mayor que 30 **y** la lluvia mayor que 500 mm (criterio ficticio de práctica).",
"plantilla": r'''pendiente = 42
lluvia_mm = 480
requiere_zanjas = None
print(requiere_zanjas)''',
"pista": "Use `and` entre las dos comparaciones.",
"solucion": r'''pendiente = 42
lluvia_mm = 480
requiere_zanjas = pendiente > 30 and lluvia_mm > 500
print(requiere_zanjas)''',
"prueba": r'''assert requiere_zanjas is False, "La lluvia no supera 500, por eso debe ser False."'''},
{"titulo": "Alerta de lluvia", "nivel": 1,
"enunciado": "Guarde en `alerta` si la lluvia diaria supera 50 mm **o** el caudal supera 100 m³/s.",
"plantilla": r'''lluvia_dia = 38
caudal = 145
alerta = None
print(alerta)''',
"pista": "Con `or` basta que una condición sea verdadera.",
"solucion": r'''lluvia_dia = 38
caudal = 145
alerta = lluvia_dia > 50 or caudal > 100
print(alerta)''',
"prueba": r'''assert alerta is True, "El caudal supera 100, por eso la alerta debe ser True."'''},
{"titulo": "Dentro del rango", "nivel": 2,
"enunciado": "Guarde en `en_rango` si `ph` está entre 6.0 y 7.5 (inclusive), usando una comparación encadenada.",
"plantilla": r'''ph = 7.8
en_rango = None
print(en_rango)''',
"pista": "`6.0 <= ph <= 7.5`",
"solucion": r'''ph = 7.8
en_rango = 6.0 <= ph <= 7.5
print(en_rango)''',
"prueba": r'''assert en_rango is False, "7.8 está fuera del rango 6.0–7.5."'''},
],
"quiz": [
{"pregunta": "¿Qué devuelve `5 == 5.0`?", "opciones": ["True", "False", "Error", "None"], "correcta": 0,
 "explicacion": "Python compara el valor numérico: 5 y 5.0 valen lo mismo."},
{"pregunta": "¿Qué devuelve `True and False`?", "opciones": ["True", "False"], "correcta": 1,
 "explicacion": "and exige que ambas partes sean verdaderas."},
{"pregunta": "¿Qué devuelve `not (3 > 1)`?", "opciones": ["True", "False"], "correcta": 1,
 "explicacion": "3 > 1 es True y not lo invierte a False."},
{"pregunta": "¿Cuál de estos valores se evalúa como falso?", "opciones": ["\"0\"", "\" \"", "0", "-1"], "correcta": 2,
 "explicacion": "El número 0 es falso. \"0\" y \" \" son textos no vacíos, por tanto verdaderos."},
],
"glosario": {
 "Booleano": "Valor lógico True o False; resultado de una comparación.",
 "Operador lógico": "and, or y not; combinan o invierten condiciones.",
},
})

# ---------------------------------------------------------------------------
MODULOS.append({
"id": "M04",
"titulo": "Trabajando con texto",
"duracion": "2 h",
"objetivos": [
    "Crear, unir y formatear textos.",
    "Usar f-strings para construir mensajes y reportes.",
    "Extraer partes de un texto con índices y rebanadas.",
    "Aplicar métodos de texto para limpiar datos.",
],
"lecciones": [
{"titulo": "Crear textos",
"texto": r'''Un texto (*string* o cadena) se escribe entre comillas simples `' '` o dobles `" "`. Para textos de varias líneas se usan comillas triples `""" """`.

- `len(texto)` cuenta los caracteres (incluye espacios).
- `\n` dentro de un texto significa **salto de línea**.''',
"codigo": r'''especie = "Faique"
nota = 'Árbol con "espinas" blancas'
print(especie, len(especie))
print(nota)
print("Línea 1\nLínea 2")'''},
{"titulo": "Unir textos y f-strings",
"texto": r'''La forma más cómoda de combinar texto y variables es la **f-string**: se antepone `f` a las comillas y las variables van entre llaves `{ }`.

Formatos útiles dentro de las llaves:

| Formato | Resultado | Uso |
|---|---|---|
| `{x:.2f}` | 2 decimales | `12.50` |
| `{x:,.2f}` | separador de miles | `12,345.68` |
| `{x:.1%}` | porcentaje | `85.0%` |
| `{texto:>10}` | alinear a la derecha en 10 espacios | tablas |''',
"codigo": r'''parcela = 3
area_ha = 12.5
costo = 45870.456
avance = 0.85

print("Parcela " + str(parcela))                 # unión con +
print(f"Parcela {parcela}: {area_ha:.2f} ha")    # f-string
print(f"Costo: S/ {costo:,.2f}")
print(f"Avance: {avance:.1%}")'''},
{"titulo": "Índices y rebanadas",
"texto": r'''Cada carácter tiene una **posición (índice)** que empieza en **0**. Los índices negativos cuentan desde el final.

```
 texto:   A  L  G  -  0  4  5
 índice:  0  1  2  3  4  5  6
 negativo:-7 -6 -5 -4 -3 -2 -1
```

Una **rebanada** `texto[inicio:fin]` toma desde *inicio* hasta *fin*, **sin incluir** *fin*.''',
"codigo": r'''codigo = "ALG-045"
print(codigo[0])      # primer carácter
print(codigo[-1])     # último carácter
print(codigo[0:3])    # posiciones 0, 1 y 2
print(codigo[4:])     # desde la posición 4 hasta el final'''},
{"titulo": "Métodos útiles para limpiar datos",
"texto": r'''Un **método** es una función que pertenece a un valor y se usa con un punto: `texto.upper()`. Los textos **no se modifican**: el método devuelve un texto nuevo.

| Método | Qué hace |
|---|---|
| `.upper()` / `.lower()` | mayúsculas / minúsculas |
| `.title()` / `.capitalize()` | Cada Palabra / Primera letra |
| `.strip()` | quita espacios al inicio y al final |
| `.replace(a, b)` | reemplaza a por b |
| `.split(sep)` | divide en partes (lista) |
| `sep.join(lista)` | une partes con un separador |
| `.startswith(x)` | ¿empieza con x? |
| `.count(x)` | cuántas veces aparece x |
| `x in texto` | ¿contiene x? |''',
"codigo": r'''dato = "  prosopis PALLIDA  "
limpio = dato.strip().capitalize()
print(f"[{limpio}]")

coord = "545120,9425300"
partes = coord.split(",")
print(partes)
print(float(partes[0]) + 100)

print("-".join(["M3", "B12"]))
print("pallida" in limpio)'''},
{"titulo": "Pedir datos al usuario: input()",
"texto": r'''`input("mensaje")` muestra un mensaje y espera que el usuario escriba algo. **Siempre devuelve texto**, por eso hay que convertirlo si es un número.

En el curso interactivo, escriba las respuestas en el cuadro **Entrada** (una por línea) antes de ejecutar.''',
"codigo": r'''nombre = input("¿Cuál es su nombre? ")
edad = int(input("¿Cuántos años tiene? "))
print(f"Hola, {nombre}. En 5 años tendrá {edad + 5}.")''',
"entrada": "Ana\n30"},
],
"ejercicios": [
{"titulo": "Nombre completo", "nivel": 1,
"enunciado": "Una `nombre` y `apellido` con un espacio en medio y guarde el resultado en `completo`.",
"plantilla": r'''nombre = "Ana"
apellido = "Torres"
completo = ""
print(completo)''',
"pista": "Use una f-string: `f\"{nombre} {apellido}\"`.",
"solucion": r'''nombre = "Ana"
apellido = "Torres"
completo = f"{nombre} {apellido}"
print(completo)''',
"prueba": r'''assert completo == "Ana Torres", f"Se esperaba 'Ana Torres' y se obtuvo '{completo}'."'''},
{"titulo": "Limpiar un dato de campo", "nivel": 1,
"enunciado": "El texto `\"   ALGARROBO \"` viene con espacios y en mayúsculas. Guarde en `limpio` la versión `Algarrobo`.",
"plantilla": r'''dato = "   ALGARROBO "
limpio = dato
print(f"[{limpio}]")''',
"pista": "Encadene métodos: primero `.strip()` y luego `.capitalize()`.",
"solucion": r'''dato = "   ALGARROBO "
limpio = dato.strip().capitalize()
print(f"[{limpio}]")''',
"prueba": r'''assert limpio == "Algarrobo", f"Se obtuvo '{limpio}'."'''},
{"titulo": "Descomponer un código", "nivel": 2,
"enunciado": "Del código `\"PI-ALG-0045\"`, guarde en `especie` el texto `ALG` y en `numero` el entero `45`.",
"plantilla": r'''codigo = "PI-ALG-0045"
partes = codigo.split("-")
especie = ""
numero = 0
print(especie, numero)''',
"pista": "`partes` es una lista de 3 elementos: posiciones 0, 1 y 2. Convierta la última con `int()`.",
"solucion": r'''codigo = "PI-ALG-0045"
partes = codigo.split("-")
especie = partes[1]
numero = int(partes[2])
print(especie, numero)''',
"prueba": r'''assert especie == "ALG", "especie debe ser 'ALG'."
assert numero == 45 and isinstance(numero, int), "numero debe ser el entero 45."'''},
{"titulo": "Línea de reporte", "nivel": 2,
"enunciado": "Construya en `reporte` el texto exacto `Parcela 3: 12.50 ha` usando una f-string con 2 decimales.",
"plantilla": r'''parcela = 3
area = 12.5
reporte = ""
print(reporte)''',
"pista": "Dentro de las llaves use `{area:.2f}`.",
"solucion": r'''parcela = 3
area = 12.5
reporte = f"Parcela {parcela}: {area:.2f} ha"
print(reporte)''',
"prueba": r'''assert reporte == "Parcela 3: 12.50 ha", f"Se obtuvo '{reporte}'."'''},
],
"quiz": [
{"pregunta": "Si `t = \"Python\"`, ¿qué es `t[0]`?", "opciones": ["\"P\"", "\"y\"", "\"n\"", "Error"], "correcta": 0,
 "explicacion": "Los índices empiezan en 0, que corresponde al primer carácter."},
{"pregunta": "Si `t = \"Python\"`, ¿qué es `t[1:4]`?", "opciones": ["\"Pyt\"", "\"yth\"", "\"ytho\"", "\"tho\""], "correcta": 1,
 "explicacion": "Toma las posiciones 1, 2 y 3; la 4 no se incluye."},
{"pregunta": "¿Qué imprime `print(f\"{3.14159:.2f}\")`?", "opciones": ["3.14159", "3.1", "3.14", "3"], "correcta": 2,
 "explicacion": ".2f muestra el número con 2 decimales."},
{"pregunta": "¿Qué tipo de dato devuelve siempre `input()`?", "opciones": ["int", "float", "str", "Depende de lo que se escriba"], "correcta": 2,
 "explicacion": "input() siempre devuelve texto; hay que convertirlo para operar."},
],
"glosario": {
 "Cadena (string)": "Secuencia de caracteres; tipo str.",
 "f-string": "Texto con prefijo f que permite insertar variables entre llaves.",
 "Índice": "Posición de un elemento; empieza en 0.",
 "Rebanada (slice)": "Parte de una secuencia: s[inicio:fin], sin incluir fin.",
 "Método": "Función asociada a un valor que se llama con un punto: texto.upper().",
},
})
