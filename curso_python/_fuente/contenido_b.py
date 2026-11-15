# -*- coding: utf-8 -*-
"""Contenido del curso: módulos 5 a 9 (control de flujo, estructuras y funciones)."""

MODULOS = []

# ---------------------------------------------------------------------------
MODULOS.append({
"id": "M05",
"titulo": "Tomando decisiones: if, elif y else",
"duracion": "2 h",
"objetivos": [
    "Ejecutar código solo cuando se cumple una condición.",
    "Elegir entre varias alternativas con elif y else.",
    "Comprender la sangría como parte de la sintaxis de Python.",
],
"lecciones": [
{"titulo": "La instrucción if",
"texto": r'''`if` (si) ejecuta un bloque de código **solo si** la condición es verdadera.

```
if condición:
    instrucción 1     # 4 espacios de sangría
    instrucción 2
siguiente línea       # sin sangría: se ejecuta siempre
```

Dos detalles obligatorios:

1. La línea del `if` termina con **dos puntos** `:`.
2. Las instrucciones del bloque llevan **sangría de 4 espacios** (la tecla Tab en Jupyter la inserta). La sangría indica qué líneas "pertenecen" al `if`.''',
"codigo": r'''lluvia_mm = 72
if lluvia_mm > 50:
    print("Alerta: lluvia intensa")
    print("Revisar obras de drenaje")
print("Fin de la evaluación")'''},
{"titulo": "if … else",
"texto": r'''`else` (si no) define qué hacer cuando la condición es falsa. No lleva condición.''',
"codigo": r'''dap_cm = 8.4
if dap_cm >= 10:
    print("Registrar en el inventario")
else:
    print("Registrar como regeneración")'''},
{"titulo": "Varias alternativas: elif",
"texto": r'''`elif` (si no, si…) permite evaluar condiciones en cadena. Python revisa **de arriba hacia abajo** y ejecuta **solo el primer bloque verdadero**. Por eso el orden importa.

Los rangos del ejemplo son ilustrativos, no normativos.''',
"codigo": r'''pendiente = 22   # porcentaje

if pendiente < 15:
    clase = "Suave"
elif pendiente < 30:
    clase = "Moderada"
elif pendiente < 50:
    clase = "Fuerte"
else:
    clase = "Muy fuerte"

print(f"Pendiente {pendiente}% -> {clase}")'''},
{"titulo": "Condiciones combinadas y anidadas",
"texto": r'''Puede combinar condiciones con `and`/`or` (Módulo 3) o colocar un `if` dentro de otro (anidado). Prefiera las condiciones combinadas: el código queda más plano y legible.''',
"codigo": r'''pendiente = 40
cobertura = 20   # % de cobertura vegetal

# Anidado
if pendiente > 30:
    if cobertura < 30:
        print("Prioridad alta (anidado)")

# Combinado: mismo resultado, más legible
if pendiente > 30 and cobertura < 30:
    print("Prioridad alta (combinado)")'''},
{"titulo": "Decisión en una sola línea",
"texto": r'''Para asignar un valor según una condición simple existe la **expresión condicional**:

```
variable = valor_si_verdadero if condición else valor_si_falso
```''',
"codigo": r'''nota = 13
estado = "Aprobado" if nota >= 11 else "Desaprobado"
print(estado)'''},
],
"ejercicios": [
{"titulo": "Par o impar", "nivel": 1,
"enunciado": "Guarde en `tipo` el texto `par` o `impar` según el valor de `numero`.",
"plantilla": r'''numero = 17
# escriba aquí su if / else

print(tipo)''',
"pista": "Un número es par si el resto de dividirlo entre 2 es 0: `numero % 2 == 0`.",
"solucion": r'''numero = 17
if numero % 2 == 0:
    tipo = "par"
else:
    tipo = "impar"
print(tipo)''',
"prueba": r'''assert tipo == "impar", "17 es impar."'''},
{"titulo": "Aprobado o desaprobado", "nivel": 1,
"enunciado": "En escala vigesimal se aprueba con 11 o más. Guarde en `estado` el texto `Aprobado` o `Desaprobado` para `nota = 10`.",
"plantilla": r'''nota = 10
estado = ""
print(estado)''',
"pista": "Puede usar un if/else o la expresión condicional en una línea.",
"solucion": r'''nota = 10
estado = "Aprobado" if nota >= 11 else "Desaprobado"
print(estado)''',
"prueba": r'''assert estado == "Desaprobado", "Con 10 el estado debe ser 'Desaprobado'."'''},
{"titulo": "Clasificar la lluvia del día", "nivel": 2,
"enunciado": "Clasifique `lluvia` en `clase` con estos umbrales de ejemplo: 0 → `Sin lluvia`; hasta 10 → `Ligera`; hasta 30 → `Moderada`; mayor a 30 → `Intensa`.",
"plantilla": r'''lluvia = 18.5
# escriba su if / elif / else

print(clase)''',
"pista": "Empiece por el caso `lluvia == 0`, luego `lluvia <= 10`, luego `lluvia <= 30` y al final `else`.",
"solucion": r'''lluvia = 18.5
if lluvia == 0:
    clase = "Sin lluvia"
elif lluvia <= 10:
    clase = "Ligera"
elif lluvia <= 30:
    clase = "Moderada"
else:
    clase = "Intensa"
print(clase)''',
"prueba": r'''assert clase == "Moderada", f"Con 18.5 mm se esperaba 'Moderada' y se obtuvo '{clase}'."'''},
{"titulo": "Prioridad de intervención", "nivel": 2,
"enunciado": "Asigne `prioridad`: `Alta` si pendiente > 30 y cobertura < 30; `Media` si solo se cumple una de las dos; `Baja` en otro caso.",
"plantilla": r'''pendiente = 45
cobertura = 55
prioridad = ""
print(prioridad)''',
"pista": "Evalúe primero el caso con `and`; después, el caso con `or`.",
"solucion": r'''pendiente = 45
cobertura = 55
if pendiente > 30 and cobertura < 30:
    prioridad = "Alta"
elif pendiente > 30 or cobertura < 30:
    prioridad = "Media"
else:
    prioridad = "Baja"
print(prioridad)''',
"prueba": r'''assert prioridad == "Media", f"Solo se cumple la pendiente: se esperaba 'Media' y se obtuvo '{prioridad}'."'''},
],
"quiz": [
{"pregunta": "¿Qué símbolo debe ir al final de la línea `if x > 5`?", "opciones": [";", ":", ".", "Ninguno"], "correcta": 1,
 "explicacion": "Las líneas if, elif, else, for, while y def terminan con dos puntos."},
{"pregunta": "¿Para qué sirve la sangría en Python?", "opciones": ["Solo para que se vea ordenado", "Indica qué líneas pertenecen a un bloque", "Es opcional", "Para escribir comentarios"], "correcta": 1,
 "explicacion": "En Python la sangría es parte de la sintaxis: define los bloques."},
{"pregunta": "Si varias condiciones de un if/elif son verdaderas, ¿qué bloques se ejecutan?", "opciones": ["Todos", "Solo el primero verdadero", "Solo el último", "Ninguno"], "correcta": 1,
 "explicacion": "Python se detiene en el primer bloque cuya condición es verdadera."},
{"pregunta": "¿Cuándo se ejecuta el bloque else?", "opciones": ["Siempre", "Cuando ninguna condición anterior fue verdadera", "Cuando la primera fue verdadera", "Nunca"], "correcta": 1,
 "explicacion": "else es el caso por defecto."},
],
"glosario": {
 "Condicional": "Estructura que ejecuta código según una condición (if/elif/else).",
 "Sangría (indentación)": "Espacios al inicio de la línea que definen un bloque; se usan 4.",
 "Bloque": "Conjunto de líneas con la misma sangría que se ejecutan juntas.",
},
})

# ---------------------------------------------------------------------------
MODULOS.append({
"id": "M06",
"titulo": "Repetir tareas: bucles for y while",
"duracion": "2.5 h",
"objetivos": [
    "Repetir instrucciones un número definido de veces con for y range().",
    "Recorrer los elementos de una lista o texto.",
    "Usar acumuladores y contadores.",
    "Repetir mientras se cumpla una condición con while.",
],
"lecciones": [
{"titulo": "El bucle for y range()",
"texto": r'''Un **bucle** repite instrucciones. `for` recorre una secuencia de valores, uno por uno, y en cada vuelta la variable toma el siguiente valor.

`range()` genera secuencias de números enteros:

| Expresión | Valores |
|---|---|
| `range(5)` | 0, 1, 2, 3, 4 |
| `range(1, 6)` | 1, 2, 3, 4, 5 |
| `range(0, 10, 2)` | 0, 2, 4, 6, 8 |

Igual que el `if`, la línea termina en `:` y el bloque lleva sangría.''',
"codigo": r'''for i in range(1, 6):
    print("Parcela", i)
print("Recorrido terminado")'''},
{"titulo": "Recorrer una lista o un texto",
"texto": r'''Una **lista** (se estudia a fondo en el Módulo 7) es una colección de valores entre corchetes `[ ]`. `for` puede recorrerla directamente. `enumerate()` entrega además la posición.''',
"codigo": r'''especies = ["Algarrobo", "Faique", "Hualtaco"]
for especie in especies:
    print("Especie:", especie)

for n, especie in enumerate(especies, start=1):
    print(n, especie)'''},
{"titulo": "Acumuladores y contadores",
"texto": r'''Patrón muy usado:

1. **Antes** del bucle, crear la variable en 0.
2. **Dentro** del bucle, sumarle algo en cada vuelta.
3. **Después** del bucle, usar el resultado.''',
"codigo": r'''alturas = [8.5, 12.1, 10.0, 15.3, 9.9]

total = 0          # acumulador
mayores = 0        # contador
for h in alturas:
    total += h
    if h > 10:
        mayores += 1

print("Suma de alturas:", round(total, 2))
print("Promedio:", round(total / len(alturas), 2))
print("Árboles de más de 10 m:", mayores)'''},
{"titulo": "El bucle while",
"texto": r'''`while` (mientras) repite **mientras la condición sea verdadera**. Se usa cuando no se sabe de antemano cuántas vueltas habrá.

**Cuidado:** dentro del bucle algo debe cambiar para que la condición llegue a ser falsa; si no, el bucle sería infinito (en Jupyter se detiene con el botón ■ *Interrupt*).''',
"codigo": r'''altura_cm = 50          # plantón recién instalado (valores de ejemplo)
crecimiento_anual = 60
anios = 0
while altura_cm < 300:
    altura_cm += crecimiento_anual
    anios += 1
    print(f"Año {anios}: {altura_cm} cm")
print(f"Supera los 3 m en {anios} años")'''},
{"titulo": "break y continue",
"texto": r'''- `break` **termina** el bucle inmediatamente.
- `continue` **salta** a la siguiente vuelta sin ejecutar el resto del bloque.''',
"codigo": r'''lecturas = [12, 15, -1, 18, 999, 20]
for valor in lecturas:
    if valor < 0:
        print("Lectura inválida, se omite")
        continue
    if valor == 999:
        print("Código de fin de datos")
        break
    print("Lectura:", valor)'''},
],
"ejercicios": [
{"titulo": "Suma del 1 al 100", "nivel": 1,
"enunciado": "Use un bucle `for` con `range()` para sumar los números del 1 al 100 en `total`.",
"plantilla": r'''total = 0
# su bucle aquí

print(total)''',
"pista": "`range(1, 101)` incluye el 100. Dentro del bucle: `total += i`.",
"solucion": r'''total = 0
for i in range(1, 101):
    total += i
print(total)''',
"prueba": r'''assert total == 5050, f"Se esperaba 5050 y se obtuvo {total}."'''},
{"titulo": "Contar árboles altos", "nivel": 1,
"enunciado": "Cuente en `altos` cuántas alturas de la lista son mayores a 10 m.",
"plantilla": r'''alturas = [8.5, 12.1, 10.0, 15.3, 9.9, 11.2]
altos = 0
# su bucle aquí

print(altos)''',
"pista": "Recorra la lista con `for h in alturas:` y use un `if` para sumar 1 al contador.",
"solucion": r'''alturas = [8.5, 12.1, 10.0, 15.3, 9.9, 11.2]
altos = 0
for h in alturas:
    if h > 10:
        altos += 1
print(altos)''',
"prueba": r'''assert altos == 3, f"Son 3 (12.1, 15.3 y 11.2); se obtuvo {altos}."'''},
{"titulo": "¿Cuántos años?", "nivel": 2,
"enunciado": "Un árbol mide 200 cm y crece 80 cm por año. Con `while`, calcule en `anios` cuántos años tarda en alcanzar al menos 1000 cm.",
"plantilla": r'''altura = 200
anios = 0
# su bucle while aquí

print(anios)''',
"pista": "Condición: `while altura < 1000:`; dentro, sume 80 a la altura y 1 a los años.",
"solucion": r'''altura = 200
anios = 0
while altura < 1000:
    altura += 80
    anios += 1
print(anios)''',
"prueba": r'''assert anios == 10, f"Se esperaban 10 años y se obtuvo {anios}."'''},
{"titulo": "Factorial", "nivel": 2,
"enunciado": "Calcule en `producto` el factorial de 5 (1 × 2 × 3 × 4 × 5) con un bucle.",
"plantilla": r'''producto = 1
# su bucle aquí

print(producto)''',
"pista": "El acumulador de una multiplicación empieza en 1, no en 0. Use `producto *= i`.",
"solucion": r'''producto = 1
for i in range(1, 6):
    producto *= i
print(producto)''',
"prueba": r'''assert producto == 120, f"5! = 120; se obtuvo {producto}."'''},
],
"quiz": [
{"pregunta": "¿Qué valores genera `range(3)`?", "opciones": ["1, 2, 3", "0, 1, 2", "0, 1, 2, 3", "3"], "correcta": 1,
 "explicacion": "range(n) empieza en 0 y termina en n-1."},
{"pregunta": "¿Cuándo conviene usar while en lugar de for?", "opciones": ["Cuando se conoce el número exacto de repeticiones", "Cuando se repite mientras se cumpla una condición", "Nunca", "Para recorrer listas"], "correcta": 1,
 "explicacion": "while es para repeticiones cuyo número depende de una condición."},
{"pregunta": "¿Qué hace `break`?", "opciones": ["Salta a la siguiente vuelta", "Termina el bucle", "Detiene el programa con error", "Reinicia el bucle"], "correcta": 1,
 "explicacion": "break sale inmediatamente del bucle."},
{"pregunta": "¿Con qué valor debe empezar un acumulador de suma?", "opciones": ["1", "0", "None", "Con el primer elemento siempre"], "correcta": 1,
 "explicacion": "Para sumar se empieza en 0; para multiplicar, en 1."},
],
"glosario": {
 "Bucle": "Estructura que repite un bloque de código (for, while).",
 "Iteración": "Cada una de las vueltas de un bucle.",
 "Acumulador": "Variable que va sumando (o multiplicando) valores dentro de un bucle.",
 "range()": "Función que genera secuencias de números enteros.",
},
})

# ---------------------------------------------------------------------------
MODULOS.append({
"id": "M07",
"titulo": "Listas y tuplas",
"duracion": "2.5 h",
"objetivos": [
    "Crear y modificar listas.",
    "Ordenar, buscar y resumir datos de una lista.",
    "Crear listas nuevas con comprensiones.",
    "Usar tuplas para datos fijos como coordenadas.",
],
"lecciones": [
{"titulo": "Crear listas y acceder a sus elementos",
"texto": r'''Una **lista** guarda varios valores en orden, entre corchetes y separados por comas. Puede mezclar tipos, aunque lo habitual es que sean del mismo tipo.

Se accede a cada elemento por su **índice** (empieza en 0), igual que en los textos. A diferencia de los textos, las listas **sí se pueden modificar**.''',
"codigo": r'''especies = ["Algarrobo", "Faique", "Hualtaco", "Palo santo"]
print(especies[0])
print(especies[-1])
print(len(especies))

especies[1] = "Sapote"       # modifico un elemento
print(especies)'''},
{"titulo": "Agregar y quitar elementos",
"texto": r'''| Método | Qué hace |
|---|---|
| `lista.append(x)` | agrega x al final |
| `lista.insert(i, x)` | inserta x en la posición i |
| `lista.remove(x)` | elimina el primer x encontrado |
| `lista.pop()` | quita y devuelve el último elemento |
| `lista.pop(i)` | quita y devuelve el elemento de la posición i |''',
"codigo": r'''plantones = ["Algarrobo", "Faique"]
plantones.append("Hualtaco")
plantones.insert(0, "Ceibo")
print(plantones)

plantones.remove("Faique")
ultimo = plantones.pop()
print(plantones, "| retirado:", ultimo)'''},
{"titulo": "Ordenar, buscar y resumir",
"texto": r'''- `sum()`, `min()`, `max()` y `len()` resumen listas de números.
- `sorted(lista)` devuelve una **copia ordenada**; `lista.sort()` ordena **la misma lista**.
- `x in lista` verifica si un elemento existe; `lista.index(x)` dice su posición.''',
"codigo": r'''dap = [32.5, 18.0, 45.2, 22.7, 27.1]
print("Mínimo:", min(dap), "Máximo:", max(dap))
print("Promedio:", round(sum(dap) / len(dap), 2))
print("Ordenado:", sorted(dap))
print("De mayor a menor:", sorted(dap, reverse=True))
print(45.2 in dap, dap.index(45.2))'''},
{"titulo": "Rebanadas y recorrido",
"texto": r'''Las rebanadas funcionan igual que en los textos: `lista[inicio:fin]`.''',
"codigo": r'''meses = ["ene", "feb", "mar", "abr", "may", "jun"]
print(meses[:3])     # primer trimestre
print(meses[3:])     # segundo trimestre
for m in meses[::2]: # uno sí, uno no
    print(m)'''},
{"titulo": "Comprensiones de listas",
"texto": r'''Una **comprensión** crea una lista nueva a partir de otra en una sola línea:

```
nueva = [expresión for elemento in lista]
nueva = [expresión for elemento in lista if condición]
```

Equivale a un `for` con `append`, pero más compacto. Empiece por el bucle tradicional y, cuando lo entienda, pase a la comprensión.''',
"codigo": r'''dap_cm = [32.5, 18.0, 45.2, 22.7]

# Forma tradicional
dap_m = []
for d in dap_cm:
    dap_m.append(d / 100)
print(dap_m)

# Comprensión equivalente
dap_m = [d / 100 for d in dap_cm]
print(dap_m)

# Con filtro: solo DAP mayores a 25 cm
grandes = [d for d in dap_cm if d > 25]
print(grandes)'''},
{"titulo": "Tuplas: datos que no cambian",
"texto": r'''Una **tupla** es como una lista, pero se escribe con paréntesis y **no se puede modificar**. Es ideal para datos fijos, como un par de coordenadas (Este, Norte).

El **desempaquetado** asigna cada elemento a una variable.''',
"codigo": r'''punto = (545120, 9425300)     # (Este, Norte) UTM 17S, punto de ejemplo
este, norte = punto           # desempaquetado
print("Este:", este)
print("Norte:", norte)

vertices = [(0, 0), (100, 0), (100, 50)]
for x, y in vertices:
    print(x, y)'''},
],
"ejercicios": [
{"titulo": "Actualizar la lista de especies", "nivel": 1,
"enunciado": "Agregue `Faique` al final de la lista y elimine `Eucalipto`.",
"plantilla": r'''especies = ["Algarrobo", "Eucalipto", "Hualtaco"]
# sus instrucciones aquí

print(especies)''',
"pista": "Use `.append()` y `.remove()`.",
"solucion": r'''especies = ["Algarrobo", "Eucalipto", "Hualtaco"]
especies.append("Faique")
especies.remove("Eucalipto")
print(especies)''',
"prueba": r'''assert especies == ["Algarrobo", "Hualtaco", "Faique"], f"Se obtuvo {especies}."'''},
{"titulo": "Resumen de alturas", "nivel": 1,
"enunciado": "Calcule `maximo`, `minimo` y `promedio` de la lista de alturas.",
"plantilla": r'''alturas = [12.5, 9.8, 15.1, 11.0, 13.6]
maximo = 0
minimo = 0
promedio = 0
print(maximo, minimo, round(promedio, 2))''',
"pista": "Use `max()`, `min()` y `sum() / len()`.",
"solucion": r'''alturas = [12.5, 9.8, 15.1, 11.0, 13.6]
maximo = max(alturas)
minimo = min(alturas)
promedio = sum(alturas) / len(alturas)
print(maximo, minimo, round(promedio, 2))''',
"prueba": r'''assert maximo == 15.1 and minimo == 9.8, "Revise máximo y mínimo."
assert abs(promedio - 12.4) < 1e-9, "El promedio debe ser 12.4."'''},
{"titulo": "De centímetros a metros", "nivel": 2,
"enunciado": "Con una comprensión de lista, cree `dap_m` con los diámetros convertidos a metros.",
"plantilla": r'''dap_cm = [20, 35, 50]
dap_m = []
print(dap_m)''',
"pista": "`[d / 100 for d in dap_cm]`",
"solucion": r'''dap_cm = [20, 35, 50]
dap_m = [d / 100 for d in dap_cm]
print(dap_m)''',
"prueba": r'''assert dap_m == [0.2, 0.35, 0.5], f"Se obtuvo {dap_m}."'''},
{"titulo": "Distancia entre dos puntos", "nivel": 2,
"enunciado": "Desempaquete las tuplas y calcule en `distancia` la distancia en metros entre los puntos con el teorema de Pitágoras: √((x2−x1)² + (y2−y1)²).",
"plantilla": r'''p1 = (500000, 9400000)
p2 = (500300, 9400400)
distancia = 0
print(distancia)''',
"pista": "La raíz cuadrada se obtiene elevando a 0.5: `(...) ** 0.5`.",
"solucion": r'''p1 = (500000, 9400000)
p2 = (500300, 9400400)
x1, y1 = p1
x2, y2 = p2
distancia = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
print(distancia)''',
"prueba": r'''assert abs(distancia - 500) < 1e-9, f"La distancia debe ser 500 m; se obtuvo {distancia}."'''},
],
"quiz": [
{"pregunta": "Si `l = [10, 20, 30]`, ¿qué devuelve `l[-1]`?", "opciones": ["10", "20", "30", "Error"], "correcta": 2,
 "explicacion": "El índice -1 es el último elemento."},
{"pregunta": "¿Qué método agrega un elemento al final de una lista?", "opciones": [".add()", ".append()", ".insert()", ".push()"], "correcta": 1,
 "explicacion": "append() agrega al final."},
{"pregunta": "¿Qué diferencia principal hay entre lista y tupla?", "opciones": ["La tupla no se puede modificar", "La lista no admite números", "No hay diferencia", "La tupla solo guarda textos"], "correcta": 0,
 "explicacion": "Las tuplas son inmutables."},
{"pregunta": "¿Qué produce `[x * 2 for x in [1, 2, 3]]`?", "opciones": ["[1, 2, 3, 1, 2, 3]", "[2, 4, 6]", "[1, 4, 9]", "6"], "correcta": 1,
 "explicacion": "La comprensión multiplica cada elemento por 2."},
],
"glosario": {
 "Lista": "Colección ordenada y modificable de valores: [1, 2, 3].",
 "Tupla": "Colección ordenada que no se puede modificar: (1, 2).",
 "Comprensión de lista": "Forma compacta de crear una lista: [x for x in datos].",
 "Desempaquetado": "Asignar los elementos de una secuencia a varias variables: a, b = (1, 2).",
},
})

# ---------------------------------------------------------------------------
MODULOS.append({
"id": "M08",
"titulo": "Diccionarios y conjuntos",
"duracion": "2.5 h",
"objetivos": [
    "Guardar datos con nombre usando diccionarios (clave: valor).",
    "Recorrer y modificar diccionarios.",
    "Representar una tabla como lista de diccionarios.",
    "Usar conjuntos para obtener valores únicos y comparar grupos.",
],
"lecciones": [
{"titulo": "¿Qué es un diccionario?",
"texto": r'''Un **diccionario** guarda pares **clave: valor** entre llaves `{ }`. En lugar de acceder por posición, se accede **por nombre** (clave). Es como una ficha de campo: cada campo tiene un nombre y un valor.

- `dic["clave"]` obtiene el valor (error si no existe).
- `dic.get("clave", defecto)` obtiene el valor o un valor por defecto si no existe.''',
"codigo": r'''arbol = {
    "codigo": "A-001",
    "especie": "Algarrobo",
    "dap_cm": 32.5,
    "altura_m": 9.0,
}
print(arbol["especie"])
print(arbol.get("estado", "Por verificar"))'''},
{"titulo": "Modificar, agregar y eliminar",
"texto": r'''- Asignar a una clave existente **modifica** su valor; a una clave nueva, la **agrega**.
- `del dic["clave"]` elimina un par.
- `"clave" in dic` verifica si existe la clave.''',
"codigo": r'''parcela = {"nombre": "P-01", "area_ha": 2.5}
parcela["area_ha"] = 2.8          # modifica
parcela["pendiente"] = 25         # agrega
print(parcela)
del parcela["pendiente"]
print("pendiente" in parcela)'''},
{"titulo": "Recorrer un diccionario",
"texto": r'''- `.keys()` → claves, `.values()` → valores, `.items()` → pares (clave, valor).
- Lo más usado es recorrer con `.items()`.''',
"codigo": r'''lluvia_mensual = {"ene": 45.2, "feb": 120.8, "mar": 210.5}
for mes, mm in lluvia_mensual.items():
    print(f"{mes}: {mm:>6.1f} mm")
print("Total:", sum(lluvia_mensual.values()))'''},
{"titulo": "Una tabla: lista de diccionarios",
"texto": r'''Una lista de diccionarios equivale a una **tabla**: cada diccionario es una **fila** y cada clave una **columna**. Esta estructura es el puente hacia los archivos CSV (Módulo 11) y pandas (Módulo 13).

*Datos ficticios de ejemplo.*''',
"codigo": r'''parcelas = [
    {"id": "P-01", "area_ha": 2.5, "especie": "Algarrobo"},
    {"id": "P-02", "area_ha": 1.8, "especie": "Faique"},
    {"id": "P-03", "area_ha": 3.2, "especie": "Algarrobo"},
]
total = 0
for p in parcelas:
    print(p["id"], p["area_ha"])
    total += p["area_ha"]
print("Área total:", round(total, 2), "ha")'''},
{"titulo": "Contar frecuencias",
"texto": r'''Un diccionario es ideal para contar cuántas veces aparece cada valor. Patrón: `conteo[x] = conteo.get(x, 0) + 1`.''',
"codigo": r'''registros = ["Algarrobo", "Faique", "Algarrobo", "Hualtaco", "Algarrobo", "Faique"]
conteo = {}
for especie in registros:
    conteo[especie] = conteo.get(especie, 0) + 1
print(conteo)'''},
{"titulo": "Conjuntos (set)",
"texto": r'''Un **conjunto** guarda valores **sin repetir y sin orden**. Útil para obtener valores únicos y comparar grupos:

- **Unión** `a | b`: elementos en a o en b.
- **Intersección** `a & b`: elementos presentes en ambos.
- **Diferencia** `a - b`: elementos en a pero no en b.''',
"codigo": r'''parcela_1 = {"Algarrobo", "Faique", "Hualtaco"}
parcela_2 = {"Algarrobo", "Palo santo", "Faique"}
print("En ambas:", sorted(parcela_1 & parcela_2))
print("Solo en la 1:", sorted(parcela_1 - parcela_2))
print("Todas:", sorted(parcela_1 | parcela_2))
print(set([1, 1, 2, 3, 3]))'''},
],
"ejercicios": [
{"titulo": "Completar la ficha", "nivel": 1,
"enunciado": "Agregue la clave `pendiente` con valor 25 y cambie `area_ha` a 3.0.",
"plantilla": r'''parcela = {"nombre": "P-07", "area_ha": 2.5}
# sus instrucciones aquí

print(parcela)''',
"pista": "`parcela[\"pendiente\"] = 25`",
"solucion": r'''parcela = {"nombre": "P-07", "area_ha": 2.5}
parcela["pendiente"] = 25
parcela["area_ha"] = 3.0
print(parcela)''',
"prueba": r'''assert parcela == {"nombre": "P-07", "area_ha": 3.0, "pendiente": 25}, f"Se obtuvo {parcela}."'''},
{"titulo": "Área total de la tabla", "nivel": 1,
"enunciado": "Sume en `total` el área de todas las parcelas.",
"plantilla": r'''parcelas = [
    {"id": "P-01", "area_ha": 2.5},
    {"id": "P-02", "area_ha": 1.8},
    {"id": "P-03", "area_ha": 3.2},
]
total = 0
# su bucle aquí

print(total)''',
"pista": "Recorra con `for p in parcelas:` y acumule `p[\"area_ha\"]`.",
"solucion": r'''parcelas = [
    {"id": "P-01", "area_ha": 2.5},
    {"id": "P-02", "area_ha": 1.8},
    {"id": "P-03", "area_ha": 3.2},
]
total = 0
for p in parcelas:
    total += p["area_ha"]
print(total)''',
"prueba": r'''assert abs(total - 7.5) < 1e-9, f"El total debe ser 7.5; se obtuvo {total}."'''},
{"titulo": "Conteo por especie", "nivel": 2,
"enunciado": "Cree el diccionario `conteo` con el número de registros de cada especie.",
"plantilla": r'''registros = ["Faique", "Algarrobo", "Faique", "Sapote", "Faique"]
conteo = {}
# su bucle aquí

print(conteo)''',
"pista": "`conteo[e] = conteo.get(e, 0) + 1`",
"solucion": r'''registros = ["Faique", "Algarrobo", "Faique", "Sapote", "Faique"]
conteo = {}
for e in registros:
    conteo[e] = conteo.get(e, 0) + 1
print(conteo)''',
"prueba": r'''assert conteo == {"Faique": 3, "Algarrobo": 1, "Sapote": 1}, f"Se obtuvo {conteo}."'''},
{"titulo": "Especies compartidas", "nivel": 2,
"enunciado": "Guarde en `comunes` el conjunto de especies presentes en ambas parcelas.",
"plantilla": r'''p1 = ["Algarrobo", "Faique", "Hualtaco", "Faique"]
p2 = ["Sapote", "Faique", "Algarrobo"]
comunes = set()
print(comunes)''',
"pista": "Convierta cada lista en conjunto con `set()` y use `&`.",
"solucion": r'''p1 = ["Algarrobo", "Faique", "Hualtaco", "Faique"]
p2 = ["Sapote", "Faique", "Algarrobo"]
comunes = set(p1) & set(p2)
print(comunes)''',
"prueba": r'''assert comunes == {"Algarrobo", "Faique"}, f"Se obtuvo {comunes}."'''},
],
"quiz": [
{"pregunta": "¿Cómo se obtiene el valor de la clave \"area\" en el diccionario d?", "opciones": ["d(area)", "d[\"area\"]", "d.area", "d[0]"], "correcta": 1,
 "explicacion": "Se usa la clave entre corchetes."},
{"pregunta": "¿Qué devuelve `d.get(\"x\", 0)` si la clave \"x\" no existe?", "opciones": ["Error", "None", "0", "\"x\""], "correcta": 2,
 "explicacion": "get devuelve el valor por defecto indicado."},
{"pregunta": "¿Qué método recorre pares clave-valor?", "opciones": [".keys()", ".values()", ".items()", ".pairs()"], "correcta": 2,
 "explicacion": "items() entrega tuplas (clave, valor)."},
{"pregunta": "¿Qué resultado da `len({1, 1, 2, 2, 3})`?", "opciones": ["5", "3", "2", "Error"], "correcta": 1,
 "explicacion": "Un conjunto elimina repetidos: {1, 2, 3}."},
],
"glosario": {
 "Diccionario": "Colección de pares clave: valor; se accede por la clave.",
 "Clave": "Nombre único que identifica un valor en un diccionario.",
 "Conjunto (set)": "Colección sin orden y sin elementos repetidos.",
},
})

# ---------------------------------------------------------------------------
MODULOS.append({
"id": "M09",
"titulo": "Funciones: cree sus propias herramientas",
"duracion": "3 h",
"objetivos": [
    "Definir funciones con def, parámetros y return.",
    "Distinguir entre mostrar (print) y devolver (return) un resultado.",
    "Usar parámetros con valores por defecto y documentar funciones.",
    "Entender el alcance de las variables.",
],
"lecciones": [
{"titulo": "¿Qué es una función?",
"texto": r'''Ya ha usado funciones: `print()`, `len()`, `round()`. Una función es un **bloque de código con nombre** que realiza una tarea y se puede **reutilizar** cuantas veces se necesite.

```
def nombre_funcion(parametro1, parametro2):
    instrucciones
    return resultado
```

- `def` define la función (no la ejecuta).
- La función se ejecuta al **llamarla**: `nombre_funcion(valor1, valor2)`.''',
"codigo": r'''def saludar(nombre):
    print(f"Hola, {nombre}. Bienvenido al curso.")

saludar("Ana")
saludar("Luis")'''},
{"titulo": "Devolver resultados con return",
"texto": r'''`return` **entrega** un valor a quien llamó la función, para guardarlo o seguir operando con él. `print` solo lo **muestra**.

Regla práctica: las funciones de cálculo deben usar `return`, no `print`.''',
"codigo": r'''def area_basal(dap_cm):
    """Área basal (m²) a partir del DAP en cm: AB = π/4 × DAP²."""
    pi = 3.141592653589793
    dap_m = dap_cm / 100
    return pi / 4 * dap_m ** 2

ab = area_basal(30)
print(round(ab, 4), "m2")

total = area_basal(30) + area_basal(25) + area_basal(40)
print("AB total:", round(total, 4), "m2")'''},
{"titulo": "Parámetros por defecto y argumentos con nombre",
"texto": r'''Un parámetro puede tener un **valor por defecto**: si no se entrega, se usa ese. Al llamar, puede indicar los argumentos **por nombre**, lo que hace el código más claro.

Una función también puede devolver **varios valores** (como tupla).''',
"codigo": r'''def volumen(dap_cm, altura_m, ff=0.5):
    """Volumen (m³) = AB × altura × factor de forma (ff).
    ff = 0.5 es un valor supuesto para este ejemplo."""
    ab = 3.141592653589793 / 4 * (dap_cm / 100) ** 2
    return ab * altura_m * ff

print(round(volumen(30, 12), 4))
print(round(volumen(30, 12, ff=0.6), 4))
print(round(volumen(altura_m=12, dap_cm=30), 4))

def min_max(datos):
    return min(datos), max(datos)

menor, mayor = min_max([4, 9, 2, 7])
print(menor, mayor)'''},
{"titulo": "Documentar: docstrings",
"texto": r'''El texto entre comillas triples justo debajo de `def` es el **docstring**: explica qué hace la función, qué recibe y qué devuelve. `help(funcion)` lo muestra.

Buenas prácticas: nombres de funciones con verbos o sustantivos claros (`calcular_volumen`), una sola tarea por función y funciones cortas.'''},
{"titulo": "Alcance de las variables",
"texto": r'''Las variables creadas **dentro** de una función son **locales**: solo existen mientras la función se ejecuta. Para sacar un valor de la función, use `return`.''',
"codigo": r'''def calcular():
    resultado_local = 42
    return resultado_local

valor = calcular()
print(valor)
# print(resultado_local)  # daría NameError: no existe fuera de la función'''},
{"titulo": "Funciones como herramienta de orden: key y lambda",
"texto": r'''Algunas funciones reciben **otra función** como argumento. Por ejemplo, `sorted(lista, key=...)` ordena según lo que devuelva esa función.

`lambda` crea una función pequeña de una sola línea, sin nombre: `lambda x: x["dap"]`.''',
"codigo": r'''arboles = [
    {"id": "A1", "dap": 32.5},
    {"id": "A2", "dap": 18.0},
    {"id": "A3", "dap": 45.2},
]
ordenados = sorted(arboles, key=lambda a: a["dap"], reverse=True)
for a in ordenados:
    print(a["id"], a["dap"])'''},
],
"ejercicios": [
{"titulo": "Metros cuadrados a hectáreas", "nivel": 1,
"enunciado": "Complete la función `a_hectareas(m2)` para que **devuelva** el área en hectáreas.",
"plantilla": r'''def a_hectareas(m2):
    # complete aquí
    pass

print(a_hectareas(25000))''',
"pista": "Reemplace `pass` por `return m2 / 10000`.",
"solucion": r'''def a_hectareas(m2):
    return m2 / 10000

print(a_hectareas(25000))''',
"prueba": r'''assert a_hectareas(25000) == 2.5, "25 000 m2 = 2.5 ha"
assert a_hectareas(0) == 0, "0 m2 = 0 ha"'''},
{"titulo": "Área basal", "nivel": 1,
"enunciado": "Complete `area_basal(dap_cm)` para que devuelva el área basal en m²: AB = π/4 × (DAP en m)².",
"plantilla": r'''PI = 3.141592653589793

def area_basal(dap_cm):
    # complete aquí
    pass

print(area_basal(30))''',
"pista": "Primero convierta a metros: `dap_m = dap_cm / 100`. Luego `return PI / 4 * dap_m ** 2`.",
"solucion": r'''PI = 3.141592653589793

def area_basal(dap_cm):
    dap_m = dap_cm / 100
    return PI / 4 * dap_m ** 2

print(area_basal(30))''',
"prueba": r'''assert abs(area_basal(30) - 0.0706858) < 1e-6, "Para DAP 30 cm, AB ≈ 0.0707 m2."
assert abs(area_basal(100) - 0.7853982) < 1e-6, "Para DAP 100 cm, AB ≈ 0.7854 m2."'''},
{"titulo": "Clasificador reutilizable", "nivel": 2,
"enunciado": "Escriba `clasificar_pendiente(p)` que devuelva `Suave` (< 15), `Moderada` (< 30), `Fuerte` (< 50) o `Muy fuerte`.",
"plantilla": r'''def clasificar_pendiente(p):
    # complete aquí
    pass

for p in [5, 15, 29.9, 50]:
    print(p, clasificar_pendiente(p))''',
"pista": "Use if/elif/else y un `return` en cada rama.",
"solucion": r'''def clasificar_pendiente(p):
    if p < 15:
        return "Suave"
    elif p < 30:
        return "Moderada"
    elif p < 50:
        return "Fuerte"
    else:
        return "Muy fuerte"

for p in [5, 15, 29.9, 50]:
    print(p, clasificar_pendiente(p))''',
"prueba": r'''casos = {5: "Suave", 15: "Moderada", 29.9: "Moderada", 30: "Fuerte", 50: "Muy fuerte", 80: "Muy fuerte"}
for valor, esperado in casos.items():
    obtenido = clasificar_pendiente(valor)
    assert obtenido == esperado, f"Para {valor} se esperaba '{esperado}' y se obtuvo '{obtenido}'."'''},
{"titulo": "Volumen con valor por defecto", "nivel": 3,
"enunciado": "Escriba `volumen(dap_cm, altura_m, ff=0.5)` que devuelva AB × altura × ff, reutilizando `area_basal`.",
"plantilla": r'''PI = 3.141592653589793

def area_basal(dap_cm):
    return PI / 4 * (dap_cm / 100) ** 2

def volumen(dap_cm, altura_m, ff=0.5):
    # complete aquí
    pass

print(round(volumen(30, 12), 4))''',
"pista": "Dentro de `volumen` llame a `area_basal(dap_cm)` y multiplique.",
"solucion": r'''PI = 3.141592653589793

def area_basal(dap_cm):
    return PI / 4 * (dap_cm / 100) ** 2

def volumen(dap_cm, altura_m, ff=0.5):
    return area_basal(dap_cm) * altura_m * ff

print(round(volumen(30, 12), 4))''',
"prueba": r'''assert abs(volumen(30, 12) - 0.4241150) < 1e-6, "Con ff por defecto (0.5) se esperaba ≈ 0.4241."
assert abs(volumen(30, 12, ff=0.7) - 0.5937610) < 1e-6, "Con ff = 0.7 se esperaba ≈ 0.5938."'''},
],
"quiz": [
{"pregunta": "¿Qué palabra se usa para definir una función?", "opciones": ["function", "def", "func", "define"], "correcta": 1,
 "explicacion": "Las funciones se definen con def."},
{"pregunta": "¿Qué diferencia hay entre print y return?", "opciones": ["Ninguna", "print muestra; return entrega el valor para seguir usándolo", "return muestra en pantalla", "print es más rápido"], "correcta": 1,
 "explicacion": "return devuelve el valor a quien llamó la función."},
{"pregunta": "Con `def f(a, b=2): return a * b`, ¿qué devuelve `f(5)`?", "opciones": ["5", "7", "10", "Error"], "correcta": 2,
 "explicacion": "b toma su valor por defecto, 2: 5 × 2 = 10."},
{"pregunta": "¿Qué devuelve una función que no tiene return?", "opciones": ["0", "None", "Error", "\"\""], "correcta": 1,
 "explicacion": "Sin return, la función devuelve None."},
],
"glosario": {
 "def": "Palabra clave para definir una función.",
 "Parámetro / argumento": "Parámetro: nombre en la definición. Argumento: valor entregado al llamar.",
 "return": "Instrucción que devuelve un valor y termina la función.",
 "Docstring": "Texto de documentación al inicio de una función.",
 "lambda": "Función anónima de una sola expresión.",
 "Alcance (scope)": "Zona del programa donde una variable existe.",
},
})
