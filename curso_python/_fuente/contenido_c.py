# -*- coding: utf-8 -*-
"""Contenido del curso: módulos 10 a 14 (errores, archivos, clases, datos y SIG)."""

MODULOS = []

# ---------------------------------------------------------------------------
MODULOS.append({
"id": "M10",
"titulo": "Errores, excepciones y módulos",
"duracion": "2.5 h",
"objetivos": [
    "Reconocer los errores más frecuentes y su causa.",
    "Controlar errores con try / except para que el programa no se detenga.",
    "Importar y usar módulos de la biblioteca estándar.",
    "Saber cómo instalar paquetes externos.",
],
"lecciones": [
{"titulo": "Errores frecuentes",
"texto": r'''| Error | Causa típica | Ejemplo |
|---|---|---|
| `SyntaxError` | Código mal escrito | falta `:` o una comilla |
| `IndentationError` | Sangría incorrecta | bloque sin espacios |
| `NameError` | Variable o función inexistente | `prin("x")` |
| `TypeError` | Operación entre tipos incompatibles | `"5" + 3` |
| `ValueError` | Valor inválido para la conversión | `int("abc")` |
| `ZeroDivisionError` | División entre cero | `10 / 0` |
| `IndexError` | Índice fuera de la lista | `[1, 2][5]` |
| `KeyError` | Clave inexistente en diccionario | `d["x"]` |

Ejecute el ejemplo y observe el tipo de error en la última línea.''',
"codigo": r'''datos = [10, 20, 30]
print(datos[5])''',
"error": True},
{"titulo": "Capturar errores con try / except",
"texto": r'''Cuando un error es **previsible** (por ejemplo, datos de campo mal digitados), se puede **capturar** para que el programa continúe:

```
try:
    código que podría fallar
except TipoDeError:
    qué hacer si falla
```

Capture siempre errores **específicos** (`ValueError`, `ZeroDivisionError`…), no todos a ciegas.''',
"codigo": r'''valores = ["12.5", "8,3", "abc", "15"]
limpios = []
for v in valores:
    try:
        limpios.append(float(v))
    except ValueError:
        print(f"Valor no numérico, por verificar: {v!r}")
print(limpios)'''},
{"titulo": "else, finally y raise",
"texto": r'''- `else`: se ejecuta si **no** hubo error.
- `finally`: se ejecuta **siempre** (útil para cerrar recursos).
- `raise`: **provoca** un error a propósito, por ejemplo para rechazar datos imposibles.''',
"codigo": r'''def validar_dap(dap_cm):
    if dap_cm <= 0:
        raise ValueError(f"DAP inválido: {dap_cm}")
    return dap_cm

for dato in [25, -3]:
    try:
        validar_dap(dato)
    except ValueError as e:
        print("Error:", e)
    else:
        print("Dato correcto:", dato)
    finally:
        print("-- revisión terminada --")'''},
{"titulo": "Módulos de la biblioteca estándar",
"texto": r'''Un **módulo** es un archivo con funciones listas para usar. Python incluye muchos (la *biblioteca estándar*):

| Módulo | Para qué |
|---|---|
| `math` | funciones matemáticas (`pi`, `sqrt`, `log`) |
| `statistics` | media, mediana, desviación estándar |
| `random` | números aleatorios, muestreo |
| `datetime` | fechas y horas |
| `csv`, `json` | leer y escribir datos (Módulo 11) |

Formas de importar: `import math`, `import statistics as st`, `from math import pi, sqrt`.''',
"codigo": r'''import math
import statistics as st
from datetime import date

print(math.pi, math.sqrt(16))
alturas = [12.5, 9.8, 15.1, 11.0, 13.6]
print("Media:", st.mean(alturas))
print("Mediana:", st.median(alturas))
print("Desv. estándar:", round(st.stdev(alturas), 3))

inicio = date(2026, 1, 15)
fin = date(2026, 3, 1)
print("Días transcurridos:", (fin - inicio).days)'''},
{"titulo": "Instalar paquetes externos",
"texto": r'''Los paquetes que no vienen con Python (pandas, openpyxl, pyproj, geopandas…) se instalan **una sola vez** desde la terminal:

- Con Anaconda (recomendado): abra *Anaconda Prompt* y escriba `conda install pandas openpyxl pyproj`
- Con pip: `pip install pandas openpyxl pyproj`
- Dentro de Jupyter o Colab: `%pip install pyproj` en una celda.

Luego se importan como cualquier módulo: `import pandas as pd`.

También puede crear **sus propios módulos**: guarde funciones en un archivo, por ejemplo `forestal.py`, y en otro archivo de la misma carpeta escriba `from forestal import area_basal`.'''},
],
"ejercicios": [
{"titulo": "Conversión segura", "nivel": 1,
"enunciado": "Complete `a_numero(texto)`: devuelve el número decimal o `None` si el texto no es numérico.",
"plantilla": r'''def a_numero(texto):
    # use try / except
    pass

print(a_numero("12.5"), a_numero("abc"))''',
"pista": "Dentro de `try` haga `return float(texto)`; en `except ValueError` haga `return None`.",
"solucion": r'''def a_numero(texto):
    try:
        return float(texto)
    except ValueError:
        return None

print(a_numero("12.5"), a_numero("abc"))''',
"prueba": r'''assert a_numero("12.5") == 12.5, "'12.5' debe convertirse en 12.5"
assert a_numero("abc") is None, "'abc' debe devolver None"
assert a_numero("7") == 7.0, "'7' debe convertirse en 7.0"'''},
{"titulo": "División segura", "nivel": 1,
"enunciado": "Complete `densidad(arboles, area_ha)`: devuelve árboles por hectárea, o `None` si el área es 0.",
"plantilla": r'''def densidad(arboles, area_ha):
    pass

print(densidad(450, 0.5), densidad(10, 0))''',
"pista": "Capture `ZeroDivisionError`.",
"solucion": r'''def densidad(arboles, area_ha):
    try:
        return arboles / area_ha
    except ZeroDivisionError:
        return None

print(densidad(450, 0.5), densidad(10, 0))''',
"prueba": r'''assert densidad(450, 0.5) == 900, "450 árboles en 0.5 ha = 900 árb/ha"
assert densidad(10, 0) is None, "Con área 0 debe devolver None"'''},
{"titulo": "DAP a partir de la circunferencia", "nivel": 2,
"enunciado": "En campo se mide la circunferencia a la altura del pecho (CAP). Complete `dap_desde_cap(cap_cm)` = CAP / π usando `math.pi`.",
"plantilla": r'''import math

def dap_desde_cap(cap_cm):
    pass

print(round(dap_desde_cap(100), 2))''',
"pista": "`return cap_cm / math.pi`",
"solucion": r'''import math

def dap_desde_cap(cap_cm):
    return cap_cm / math.pi

print(round(dap_desde_cap(100), 2))''',
"prueba": r'''import math
assert abs(dap_desde_cap(100) - 31.8310) < 1e-3, "CAP 100 cm → DAP ≈ 31.83 cm"
assert abs(dap_desde_cap(math.pi * 20) - 20) < 1e-9, "CAP de un DAP de 20 cm debe dar 20"'''},
{"titulo": "Estadística rápida", "nivel": 2,
"enunciado": "Con el módulo `statistics`, guarde en `mediana` la mediana y en `media` el promedio de las lluvias.",
"plantilla": r'''import statistics

lluvias = [0, 2.5, 12.0, 45.8, 3.1, 0, 22.4]
mediana = None
media = None
print(mediana, round(media, 2) if media is not None else media)''',
"pista": "`statistics.median(lluvias)` y `statistics.mean(lluvias)`.",
"solucion": r'''import statistics

lluvias = [0, 2.5, 12.0, 45.8, 3.1, 0, 22.4]
mediana = statistics.median(lluvias)
media = statistics.mean(lluvias)
print(mediana, round(media, 2))''',
"prueba": r'''assert mediana == 3.1, f"La mediana debe ser 3.1; se obtuvo {mediana}"
assert abs(media - 12.2571) < 1e-3, "La media debe ser ≈ 12.26"'''},
],
"quiz": [
{"pregunta": "¿Qué error produce `int(\"abc\")`?", "opciones": ["TypeError", "ValueError", "NameError", "SyntaxError"], "correcta": 1,
 "explicacion": "El tipo es correcto (texto) pero el valor no se puede convertir: ValueError."},
{"pregunta": "¿Cuándo se ejecuta el bloque finally?", "opciones": ["Solo si hay error", "Solo si no hay error", "Siempre", "Nunca"], "correcta": 2,
 "explicacion": "finally se ejecuta en todos los casos."},
{"pregunta": "¿Cómo se importa solo la función sqrt del módulo math?", "opciones": ["import sqrt", "from math import sqrt", "import math.sqrt()", "math import sqrt"], "correcta": 1,
 "explicacion": "from módulo import nombre."},
{"pregunta": "¿Con qué comando se instala pandas usando pip?", "opciones": ["python pandas", "pip install pandas", "import pandas", "install pandas"], "correcta": 1,
 "explicacion": "pip install nombre_paquete, desde la terminal."},
],
"glosario": {
 "Excepción": "Error que ocurre durante la ejecución y puede capturarse.",
 "try / except": "Estructura para capturar y manejar excepciones.",
 "Módulo": "Archivo .py con funciones reutilizables que se importa con import.",
 "Paquete": "Conjunto de módulos que se instala con pip o conda (p. ej., pandas).",
 "Biblioteca estándar": "Módulos que vienen incluidos con Python.",
},
})

# ---------------------------------------------------------------------------
MODULOS.append({
"id": "M11",
"titulo": "Archivos y datos: TXT, CSV y JSON",
"duracion": "3 h",
"objetivos": [
    "Leer y escribir archivos de texto con with open().",
    "Procesar archivos CSV con el módulo csv.",
    "Guardar y leer datos estructurados en JSON.",
    "Manejar rutas de archivos con pathlib.",
],
"lecciones": [
{"titulo": "Escribir y leer un archivo de texto",
"texto": r'''`open(ruta, modo, encoding="utf-8")` abre un archivo. Modos: `"r"` leer, `"w"` escribir (borra el contenido anterior), `"a"` agregar al final.

Use siempre `with`: cierra el archivo automáticamente al terminar el bloque. Use `encoding="utf-8"` para que las tildes y la ñ se guarden correctamente.

*Este ejemplo crea un archivo en la carpeta donde está el cuaderno; en el curso interactivo (navegador) no se puede ejecutar.*''',
"codigo": r'''with open("bitacora.txt", "w", encoding="utf-8") as f:
    f.write("Día 1: instalación de parcelas\n")
    f.write("Día 2: medición de DAP\n")

with open("bitacora.txt", "r", encoding="utf-8") as f:
    for linea in f:
        print(linea.strip())''',
"web": False},
{"titulo": "Archivos CSV",
"texto": r'''Un **CSV** (valores separados por comas) es una tabla en texto plano; Excel lo abre y lo guarda. El módulo `csv` lo lee fila por fila:

- `csv.reader` entrega cada fila como **lista**.
- `csv.DictReader` entrega cada fila como **diccionario** (usa la primera fila como encabezados). Es lo más práctico.

**Importante:** los valores leídos son **texto**; conviértalos con `float()` o `int()`.

Para practicar sin archivos, `io.StringIO(texto)` hace que un texto se comporte como un archivo abierto. *Datos ficticios.*''',
"codigo": r'''import csv
import io

contenido = """codigo,especie,dap_cm,altura_m
A-01,Algarrobo,32.5,9.0
A-02,Faique,18.0,6.5
A-03,Algarrobo,45.2,11.0
"""

lector = csv.DictReader(io.StringIO(contenido))
for fila in lector:
    dap = float(fila["dap_cm"])
    print(fila["codigo"], fila["especie"], dap)

# Con un archivo real sería:
# with open("inventario.csv", encoding="utf-8") as f:
#     for fila in csv.DictReader(f):
#         ...'''},
{"titulo": "Escribir un CSV",
"texto": r'''`csv.DictWriter` escribe una lista de diccionarios como tabla. Use `newline=""` al abrir el archivo para evitar líneas en blanco en Windows. Si el CSV se abrirá en Excel con configuración regional de Perú, puede usar `delimiter=";"`.''',
"codigo": r'''import csv

resultados = [
    {"especie": "Algarrobo", "n": 12, "ab_m2": 1.284},
    {"especie": "Faique", "n": 7, "ab_m2": 0.415},
]
with open("resumen.csv", "w", newline="", encoding="utf-8") as f:
    escritor = csv.DictWriter(f, fieldnames=["especie", "n", "ab_m2"])
    escritor.writeheader()
    escritor.writerows(resultados)

with open("resumen.csv", encoding="utf-8") as f:
    print(f.read())''',
"web": False},
{"titulo": "JSON: datos estructurados",
"texto": r'''**JSON** es un formato de texto para guardar diccionarios y listas (muy usado por aplicaciones web y formularios como KoBoToolbox).

- `json.dumps(objeto)` → texto JSON; `json.loads(texto)` → objeto Python.
- `json.dump(objeto, archivo)` y `json.load(archivo)` trabajan directamente con archivos.''',
"codigo": r'''import json

parcela = {"id": "P-01", "area_ha": 2.5, "especies": ["Algarrobo", "Faique"]}
texto = json.dumps(parcela, ensure_ascii=False, indent=2)
print(texto)

de_vuelta = json.loads(texto)
print(de_vuelta["especies"][0])'''},
{"titulo": "Rutas con pathlib",
"texto": r'''`pathlib.Path` maneja rutas de forma segura en Windows, Mac y Linux:

- `Path("datos") / "inventario.csv"` construye la ruta con el separador correcto.
- `.exists()`, `.name`, `.suffix`, `.stem` y `.glob("*.csv")` para buscar archivos.

En Windows, si escribe rutas como texto, use barras normales `"C:/Proyectos/datos.csv"` o una cadena *raw* `r"C:\Proyectos\datos.csv"`.''',
"codigo": r'''from pathlib import Path

carpeta = Path(".")
ruta = carpeta / "resumen.csv"
print(ruta.name, ruta.suffix, ruta.stem, ruta.exists())

for archivo in sorted(carpeta.glob("*.csv")):
    print("Encontrado:", archivo.name)''',
"web": False},
],
"ejercicios": [
{"titulo": "Contar líneas", "nivel": 1,
"enunciado": "Guarde en `n_lineas` cuántas líneas con contenido tiene el texto.",
"plantilla": r'''texto = """Parcela 1 instalada
Parcela 2 instalada
Parcela 3 pendiente
"""
n_lineas = 0
print(n_lineas)''',
"pista": "`texto.strip().split(\"\\n\")` produce una lista de líneas; cuente con `len()`.",
"solucion": r'''texto = """Parcela 1 instalada
Parcela 2 instalada
Parcela 3 pendiente
"""
n_lineas = len(texto.strip().split("\n"))
print(n_lineas)''',
"prueba": r'''assert n_lineas == 3, f"Son 3 líneas; se obtuvo {n_lineas}."'''},
{"titulo": "Total de árboles desde un CSV", "nivel": 2,
"enunciado": "Lea el CSV con `csv.DictReader` y sume en `total` la columna `arboles`.",
"plantilla": r'''import csv, io

datos = """parcela,arboles
P-01,45
P-02,38
P-03,52
"""
total = 0
# su código aquí

print(total)''',
"pista": "Recorra `csv.DictReader(io.StringIO(datos))` y sume `int(fila[\"arboles\"])`.",
"solucion": r'''import csv, io

datos = """parcela,arboles
P-01,45
P-02,38
P-03,52
"""
total = 0
for fila in csv.DictReader(io.StringIO(datos)):
    total += int(fila["arboles"])
print(total)''',
"prueba": r'''assert total == 135, f"El total debe ser 135; se obtuvo {total}."'''},
{"titulo": "Promedio por especie", "nivel": 3,
"enunciado": "Calcule en `promedios` un diccionario con el DAP promedio de cada especie (redondeado a 2 decimales).",
"plantilla": r'''import csv, io

datos = """especie,dap_cm
Algarrobo,30
Faique,18
Algarrobo,40
Faique,22
Algarrobo,35
"""
promedios = {}
# su código aquí

print(promedios)''',
"pista": "Acumule en dos diccionarios: `suma[especie]` y `cuenta[especie]`. Al final, divida.",
"solucion": r'''import csv, io

datos = """especie,dap_cm
Algarrobo,30
Faique,18
Algarrobo,40
Faique,22
Algarrobo,35
"""
suma = {}
cuenta = {}
for fila in csv.DictReader(io.StringIO(datos)):
    e = fila["especie"]
    suma[e] = suma.get(e, 0) + float(fila["dap_cm"])
    cuenta[e] = cuenta.get(e, 0) + 1
promedios = {e: round(suma[e] / cuenta[e], 2) for e in suma}
print(promedios)''',
"prueba": r'''assert promedios == {"Algarrobo": 35.0, "Faique": 20.0}, f"Se obtuvo {promedios}."'''},
{"titulo": "Leer un JSON", "nivel": 1,
"enunciado": "Convierta el texto JSON en un diccionario y guarde en `responsable` el nombre del responsable.",
"plantilla": r'''import json

texto = '{"brigada": 2, "responsable": "J. Pérez", "parcelas": [1, 2, 3]}'
responsable = ""
print(responsable)''',
"pista": "`datos = json.loads(texto)` y luego `datos[\"responsable\"]`.",
"solucion": r'''import json

texto = '{"brigada": 2, "responsable": "J. Pérez", "parcelas": [1, 2, 3]}'
datos = json.loads(texto)
responsable = datos["responsable"]
print(responsable)''',
"prueba": r'''assert responsable == "J. Pérez", f"Se obtuvo {responsable!r}."'''},
],
"quiz": [
{"pregunta": "¿Qué ventaja tiene usar `with open(...)`?", "opciones": ["Lee más rápido", "Cierra el archivo automáticamente", "Convierte a número", "Permite abrir Excel"], "correcta": 1,
 "explicacion": "with garantiza que el archivo se cierre al terminar el bloque, incluso si hay un error."},
{"pregunta": "¿Qué hace el modo \"w\"?", "opciones": ["Lee", "Escribe y reemplaza el contenido", "Agrega al final", "Elimina el archivo"], "correcta": 1,
 "explicacion": "w crea el archivo o borra su contenido anterior antes de escribir."},
{"pregunta": "Los valores que entrega csv.DictReader son de tipo…", "opciones": ["int", "float", "str", "Depende de la columna"], "correcta": 2,
 "explicacion": "Todo se lee como texto; hay que convertir con float() o int()."},
{"pregunta": "¿Qué función convierte un texto JSON en un objeto Python?", "opciones": ["json.dumps", "json.loads", "json.text", "json.read"], "correcta": 1,
 "explicacion": "loads = load string (cargar desde texto)."},
],
"glosario": {
 "CSV": "Archivo de texto con valores separados por comas (o punto y coma).",
 "JSON": "Formato de texto para datos estructurados (diccionarios y listas).",
 "Codificación (encoding)": "Forma de guardar caracteres; use utf-8 para tildes y ñ.",
 "with": "Bloque que administra un recurso (p. ej., un archivo) y lo cierra al terminar.",
 "Ruta (path)": "Ubicación de un archivo en el disco.",
},
})

# ---------------------------------------------------------------------------
MODULOS.append({
"id": "M12",
"titulo": "Clases y objetos (programación orientada a objetos)",
"duracion": "3 h",
"objetivos": [
    "Comprender qué son una clase y un objeto.",
    "Definir clases con atributos y métodos.",
    "Personalizar cómo se muestra un objeto.",
    "Reutilizar código con herencia.",
],
"lecciones": [
{"titulo": "¿Qué es un objeto?",
"texto": r'''Ya ha usado objetos sin saberlo: un texto es un objeto de la clase `str` y tiene métodos como `.upper()`; una lista es un objeto de la clase `list` con métodos como `.append()`.

- Una **clase** es un **molde** o plantilla (por ejemplo, el formato de una ficha de árbol).
- Un **objeto** (instancia) es un **ejemplar concreto** hecho con ese molde (la ficha llena del árbol A-001).
- Los **atributos** son los datos del objeto (especie, DAP); los **métodos** son lo que el objeto sabe hacer (calcular su área basal).''',
"codigo": r'''texto = "algarrobo"
print(type(texto))
print(texto.upper())   # método de la clase str'''},
{"titulo": "Definir una clase",
"texto": r'''```
class NombreClase:
    def __init__(self, parametro):
        self.atributo = parametro
```

- `class` define la clase; por convención el nombre empieza con **mayúscula**.
- `__init__` es el **constructor**: se ejecuta al crear cada objeto y guarda sus datos.
- `self` representa **al propio objeto**; es siempre el primer parámetro de los métodos.''',
"codigo": r'''class Arbol:
    def __init__(self, codigo, especie, dap_cm, altura_m):
        self.codigo = codigo
        self.especie = especie
        self.dap_cm = dap_cm
        self.altura_m = altura_m

a1 = Arbol("A-001", "Algarrobo", 32.5, 9.0)
a2 = Arbol("A-002", "Faique", 18.0, 6.5)
print(a1.especie, a1.dap_cm)
print(a2.especie, a2.dap_cm)'''},
{"titulo": "Métodos: lo que el objeto sabe hacer",
"texto": r'''Un método es una función definida dentro de la clase. Accede a los datos del objeto mediante `self`.''',
"codigo": r'''import math

class Arbol:
    def __init__(self, codigo, especie, dap_cm, altura_m):
        self.codigo = codigo
        self.especie = especie
        self.dap_cm = dap_cm
        self.altura_m = altura_m

    def area_basal(self):
        return math.pi / 4 * (self.dap_cm / 100) ** 2

    def volumen(self, ff=0.5):   # ff supuesto para el ejemplo
        return self.area_basal() * self.altura_m * ff

a = Arbol("A-001", "Algarrobo", 32.5, 9.0)
print(f"AB: {a.area_basal():.4f} m2")
print(f"Volumen: {a.volumen():.4f} m3")'''},
{"titulo": "Mostrar un objeto: __str__",
"texto": r'''Si imprime un objeto sin más, Python muestra algo poco útil como `<__main__.Arbol object at 0x...>`. El método especial `__str__` define el texto que se muestra con `print()`.''',
"codigo": r'''class Parcela:
    def __init__(self, nombre, area_ha):
        self.nombre = nombre
        self.area_ha = area_ha
        self.arboles = []          # lista que empieza vacía

    def agregar(self, especie):
        self.arboles.append(especie)

    def __str__(self):
        return f"Parcela {self.nombre} ({self.area_ha} ha, {len(self.arboles)} árboles)"

p = Parcela("P-01", 0.5)
p.agregar("Algarrobo")
p.agregar("Faique")
print(p)'''},
{"titulo": "Herencia",
"texto": r'''Una clase puede **heredar** de otra: recibe todos sus atributos y métodos y puede agregar o modificar algunos. `super().__init__(...)` llama al constructor de la clase padre.''',
"codigo": r'''class Arbol:
    def __init__(self, especie, dap_cm):
        self.especie = especie
        self.dap_cm = dap_cm

    def describir(self):
        return f"{self.especie}, DAP {self.dap_cm} cm"

class ArbolFrutal(Arbol):
    def __init__(self, especie, dap_cm, produccion_kg):
        super().__init__(especie, dap_cm)
        self.produccion_kg = produccion_kg

    def describir(self):
        return super().describir() + f", produce {self.produccion_kg} kg/año"

print(Arbol("Faique", 20).describir())
print(ArbolFrutal("Mango", 35, 120).describir())'''},
],
"ejercicios": [
{"titulo": "Clase Parcela", "nivel": 1,
"enunciado": "Complete la clase `Parcela` con los atributos `nombre` y `area_ha`, y el método `area_m2()` que devuelva el área en m².",
"plantilla": r'''class Parcela:
    def __init__(self, nombre, area_ha):
        # guarde los atributos con self.
        pass

    def area_m2(self):
        pass

p = Parcela("P-01", 2.5)
print(p.nombre, p.area_m2())''',
"pista": "`self.nombre = nombre`, `self.area_ha = area_ha` y `return self.area_ha * 10000`.",
"solucion": r'''class Parcela:
    def __init__(self, nombre, area_ha):
        self.nombre = nombre
        self.area_ha = area_ha

    def area_m2(self):
        return self.area_ha * 10000

p = Parcela("P-01", 2.5)
print(p.nombre, p.area_m2())''',
"prueba": r'''p = Parcela("P-09", 1.2)
assert p.nombre == "P-09", "El atributo nombre no se guardó."
assert abs(p.area_m2() - 12000) < 1e-9, "1.2 ha = 12 000 m2"'''},
{"titulo": "Árbol con métodos", "nivel": 2,
"enunciado": "Complete `area_basal()` (π/4 × DAP² en m) y `volumen(ff=0.5)` (AB × altura × ff).",
"plantilla": r'''import math

class Arbol:
    def __init__(self, dap_cm, altura_m):
        self.dap_cm = dap_cm
        self.altura_m = altura_m

    def area_basal(self):
        pass

    def volumen(self, ff=0.5):
        pass

a = Arbol(30, 12)
print(a.area_basal(), a.volumen())''',
"pista": "En `volumen` llame a `self.area_basal()`.",
"solucion": r'''import math

class Arbol:
    def __init__(self, dap_cm, altura_m):
        self.dap_cm = dap_cm
        self.altura_m = altura_m

    def area_basal(self):
        return math.pi / 4 * (self.dap_cm / 100) ** 2

    def volumen(self, ff=0.5):
        return self.area_basal() * self.altura_m * ff

a = Arbol(30, 12)
print(a.area_basal(), a.volumen())''',
"prueba": r'''a = Arbol(30, 12)
assert abs(a.area_basal() - 0.0706858) < 1e-6, "AB de DAP 30 cm ≈ 0.0707 m2"
assert abs(a.volumen() - 0.4241150) < 1e-6, "Volumen con ff 0.5 ≈ 0.4241 m3"
assert abs(a.volumen(ff=0.7) - 0.5937610) < 1e-6, "Volumen con ff 0.7 ≈ 0.5938 m3"'''},
{"titulo": "Inventario", "nivel": 3,
"enunciado": "Complete la clase `Inventario`: `agregar(dap_cm)` guarda un DAP en la lista y `promedio()` devuelve el DAP promedio (0 si está vacío).",
"plantilla": r'''class Inventario:
    def __init__(self):
        self.daps = []

    def agregar(self, dap_cm):
        pass

    def promedio(self):
        pass

inv = Inventario()
for d in [20, 30, 40]:
    inv.agregar(d)
print(inv.promedio())''',
"pista": "Use `self.daps.append(dap_cm)`. En `promedio`, si `len(self.daps) == 0` devuelva 0.",
"solucion": r'''class Inventario:
    def __init__(self):
        self.daps = []

    def agregar(self, dap_cm):
        self.daps.append(dap_cm)

    def promedio(self):
        if len(self.daps) == 0:
            return 0
        return sum(self.daps) / len(self.daps)

inv = Inventario()
for d in [20, 30, 40]:
    inv.agregar(d)
print(inv.promedio())''',
"prueba": r'''vacio = Inventario()
assert vacio.promedio() == 0, "Un inventario vacío debe devolver 0."
inv = Inventario()
for d in [20, 30, 40]:
    inv.agregar(d)
assert inv.daps == [20, 30, 40], "agregar() no guarda los valores."
assert inv.promedio() == 30, "El promedio de 20, 30 y 40 es 30."'''},
{"titulo": "Mostrar con __str__", "nivel": 2,
"enunciado": "Agregue `__str__` a la clase para que `str(Especie(\"Algarrobo\", \"Prosopis pallida\"))` devuelva `Algarrobo (Prosopis pallida)`.",
"plantilla": r'''class Especie:
    def __init__(self, comun, cientifico):
        self.comun = comun
        self.cientifico = cientifico

    # defina __str__ aquí

print(Especie("Algarrobo", "Prosopis pallida"))''',
"pista": "`def __str__(self): return f\"{self.comun} ({self.cientifico})\"`",
"solucion": r'''class Especie:
    def __init__(self, comun, cientifico):
        self.comun = comun
        self.cientifico = cientifico

    def __str__(self):
        return f"{self.comun} ({self.cientifico})"

print(Especie("Algarrobo", "Prosopis pallida"))''',
"prueba": r'''assert str(Especie("Algarrobo", "Prosopis pallida")) == "Algarrobo (Prosopis pallida)", "Revise el formato del texto."'''},
],
"quiz": [
{"pregunta": "¿Qué es un objeto?", "opciones": ["Un molde", "Un ejemplar concreto creado a partir de una clase", "Un archivo", "Un tipo de bucle"], "correcta": 1,
 "explicacion": "La clase es el molde; el objeto, el ejemplar."},
{"pregunta": "¿Cuándo se ejecuta `__init__`?", "opciones": ["Al importar el módulo", "Al crear cada objeto", "Al imprimir el objeto", "Nunca de forma automática"], "correcta": 1,
 "explicacion": "Es el constructor: se ejecuta al crear una instancia."},
{"pregunta": "¿Qué representa `self`?", "opciones": ["La clase padre", "El propio objeto", "Una variable global", "El módulo"], "correcta": 1,
 "explicacion": "self es la referencia al objeto sobre el que se trabaja."},
{"pregunta": "¿Para qué sirve la herencia?", "opciones": ["Para copiar archivos", "Para crear una clase basada en otra y reutilizar su código", "Para borrar objetos", "Para importar módulos"], "correcta": 1,
 "explicacion": "La clase hija recibe atributos y métodos de la clase padre."},
],
"glosario": {
 "Clase": "Molde que define atributos y métodos de un tipo de objeto.",
 "Objeto (instancia)": "Ejemplar concreto creado a partir de una clase.",
 "Atributo": "Dato que pertenece a un objeto (self.dap_cm).",
 "Constructor (__init__)": "Método que inicializa un objeto al crearlo.",
 "self": "Referencia al propio objeto dentro de sus métodos.",
 "Herencia": "Mecanismo para crear una clase a partir de otra.",
},
})

# ---------------------------------------------------------------------------
MODULOS.append({
"id": "M13",
"titulo": "Análisis de datos con NumPy, pandas y matplotlib",
"duracion": "4 h",
"web": False,
"objetivos": [
    "Hacer cálculos sobre columnas completas con NumPy.",
    "Crear, explorar, filtrar y resumir tablas con pandas.",
    "Leer y escribir CSV y Excel con pandas.",
    "Elaborar gráficos básicos con matplotlib.",
],
"lecciones": [
{"titulo": "Preparar el entorno",
"texto": r'''Este módulo usa paquetes externos y **no se ejecuta en el curso interactivo del navegador**: utilice **Jupyter (Anaconda)** o **Google Colab**, donde ya vienen instalados. Si faltara alguno: `%pip install numpy pandas matplotlib openpyxl`.

Por convención se importan con alias cortos: `np`, `pd` y `plt`.''',
"codigo": r'''import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
print("numpy", np.__version__)
print("pandas", pd.__version__)'''},
{"titulo": "NumPy: cálculos sobre muchos valores a la vez",
"texto": r'''Un **arreglo** (`np.array`) aplica una operación a **todos los elementos sin escribir bucles** (cálculo vectorizado). Es mucho más rápido que recorrer una lista.''',
"codigo": r'''import numpy as np

dap_cm = np.array([32.5, 18.0, 45.2, 22.7, 27.1])
ab = np.pi / 4 * (dap_cm / 100) ** 2      # sin bucle
print(np.round(ab, 4))
print("AB total:", round(ab.sum(), 4), "m2")
print("DAP medio:", dap_cm.mean(), "| máximo:", dap_cm.max())
print(dap_cm[dap_cm > 25])                 # filtro'''},
{"titulo": "pandas: la tabla (DataFrame)",
"texto": r'''Un **DataFrame** es una tabla con filas y columnas con nombre, como una hoja de Excel. Cada columna es una **Serie**.

| Instrucción | Qué muestra |
|---|---|
| `df.head()` | primeras 5 filas |
| `df.shape` | (filas, columnas) |
| `df.columns` | nombres de columnas |
| `df.dtypes` | tipo de cada columna |
| `df.describe()` | estadísticas de columnas numéricas |

*Datos ficticios de ejemplo.*''',
"codigo": r'''import pandas as pd

df = pd.DataFrame({
    "codigo": ["A-01", "A-02", "A-03", "A-04", "A-05", "A-06"],
    "especie": ["Algarrobo", "Faique", "Algarrobo", "Hualtaco", "Faique", "Algarrobo"],
    "dap_cm": [32.5, 18.0, 45.2, 22.7, 20.4, 38.1],
    "altura_m": [9.0, 6.5, 11.0, 8.2, 7.0, 10.4],
})
print(df.shape)
print(df.head())
print(df.describe().round(2))'''},
{"titulo": "Seleccionar, filtrar y crear columnas",
"texto": r'''- Una columna: `df["dap_cm"]`; varias: `df[["codigo", "dap_cm"]]`.
- Filtrar filas: `df[df["dap_cm"] > 30]`. Combinar condiciones con `&` (y) y `|` (o), cada una entre paréntesis.
- Crear una columna: `df["nueva"] = cálculo`, que se aplica a todas las filas.''',
"codigo": r'''import numpy as np
import pandas as pd

df = pd.DataFrame({
    "codigo": ["A-01", "A-02", "A-03", "A-04", "A-05", "A-06"],
    "especie": ["Algarrobo", "Faique", "Algarrobo", "Hualtaco", "Faique", "Algarrobo"],
    "dap_cm": [32.5, 18.0, 45.2, 22.7, 20.4, 38.1],
    "altura_m": [9.0, 6.5, 11.0, 8.2, 7.0, 10.4],
})
df["ab_m2"] = np.pi / 4 * (df["dap_cm"] / 100) ** 2
df["vol_m3"] = df["ab_m2"] * df["altura_m"] * 0.5      # ff supuesto 0.5

grandes = df[(df["dap_cm"] > 30) & (df["especie"] == "Algarrobo")]
print(grandes[["codigo", "dap_cm", "vol_m3"]].round(4))'''},
{"titulo": "Agrupar y resumir",
"texto": r'''`groupby` agrupa filas por una columna y resume las demás (equivale a una **tabla dinámica** de Excel).

`agg` permite nombrar varias estadísticas a la vez.''',
"codigo": r'''import numpy as np
import pandas as pd

df = pd.DataFrame({
    "especie": ["Algarrobo", "Faique", "Algarrobo", "Hualtaco", "Faique", "Algarrobo"],
    "dap_cm": [32.5, 18.0, 45.2, 22.7, 20.4, 38.1],
    "altura_m": [9.0, 6.5, 11.0, 8.2, 7.0, 10.4],
})
df["ab_m2"] = np.pi / 4 * (df["dap_cm"] / 100) ** 2

resumen = df.groupby("especie").agg(
    n_arboles=("dap_cm", "count"),
    dap_medio=("dap_cm", "mean"),
    ab_total=("ab_m2", "sum"),
).round(3)
print(resumen.sort_values("ab_total", ascending=False))'''},
{"titulo": "Leer y guardar CSV y Excel",
"texto": r'''- `pd.read_csv("archivo.csv")` y `pd.read_excel("archivo.xlsx", sheet_name="Hoja1")` leen archivos.
- `df.to_csv("salida.csv", index=False)` y `df.to_excel("salida.xlsx", index=False)` los guardan (Excel requiere el paquete `openpyxl`).
- Para un CSV con punto y coma y coma decimal (configuración regional de Perú): `pd.read_csv("archivo.csv", sep=";", decimal=",")`.''',
"codigo": r'''import io
import pandas as pd

texto = """parcela;arboles;area_ha
P-01;45;0,5
P-02;38;0,5
P-03;52;0,4
"""
df = pd.read_csv(io.StringIO(texto), sep=";", decimal=",")
df["densidad_ha"] = df["arboles"] / df["area_ha"]
print(df)
# df.to_excel("densidad.xlsx", index=False)   # guardaría en Excel'''},
{"titulo": "Gráficos con matplotlib",
"texto": r'''Pasos básicos: crear la figura, dibujar, rotular ejes y título, y mostrar (`plt.show()`). Siempre indique **unidades** en los ejes.''',
"codigo": r'''import matplotlib.pyplot as plt

meses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun"]
lluvia = [45.2, 120.8, 210.5, 95.3, 12.0, 3.4]   # datos ficticios

fig, ax = plt.subplots(figsize=(7, 3.5))
ax.bar(meses, lluvia, color="#2E86AB")
ax.set_xlabel("Mes")
ax.set_ylabel("Precipitación (mm)")
ax.set_title("Precipitación mensual (datos de ejemplo)")
ax.grid(axis="y", alpha=0.3)
plt.tight_layout()
plt.show()'''},
],
"ejercicios": [
{"titulo": "Área basal con NumPy", "nivel": 1,
"enunciado": "Calcule en `ab_total` la suma del área basal (m²) de todos los árboles, sin bucles.",
"plantilla": r'''import numpy as np

dap_cm = np.array([20, 30, 40, 50])
ab_total = 0
print(round(ab_total, 4))''',
"pista": "`(np.pi / 4 * (dap_cm / 100) ** 2).sum()`",
"solucion": r'''import numpy as np

dap_cm = np.array([20, 30, 40, 50])
ab_total = (np.pi / 4 * (dap_cm / 100) ** 2).sum()
print(round(ab_total, 4))''',
"prueba": r'''assert abs(ab_total - 0.4241150) < 1e-6, f"Se esperaba ≈ 0.4241; se obtuvo {ab_total}."'''},
{"titulo": "Filtrar un DataFrame", "nivel": 1,
"enunciado": "Guarde en `n_altos` cuántos árboles tienen altura mayor a 8 m.",
"plantilla": r'''import pandas as pd

df = pd.DataFrame({"altura_m": [9.0, 6.5, 11.0, 8.2, 7.0, 10.4]})
n_altos = 0
print(n_altos)''',
"pista": "`len(df[df[\"altura_m\"] > 8])`",
"solucion": r'''import pandas as pd

df = pd.DataFrame({"altura_m": [9.0, 6.5, 11.0, 8.2, 7.0, 10.4]})
n_altos = len(df[df["altura_m"] > 8])
print(n_altos)''',
"prueba": r'''assert n_altos == 4, f"Son 4 árboles; se obtuvo {n_altos}."'''},
{"titulo": "Promedio por especie", "nivel": 2,
"enunciado": "Guarde en `medias` una Serie con el DAP medio por especie usando `groupby`.",
"plantilla": r'''import pandas as pd

df = pd.DataFrame({
    "especie": ["Algarrobo", "Faique", "Algarrobo", "Faique"],
    "dap_cm": [30, 18, 40, 22],
})
medias = None
print(medias)''',
"pista": "`df.groupby(\"especie\")[\"dap_cm\"].mean()`",
"solucion": r'''import pandas as pd

df = pd.DataFrame({
    "especie": ["Algarrobo", "Faique", "Algarrobo", "Faique"],
    "dap_cm": [30, 18, 40, 22],
})
medias = df.groupby("especie")["dap_cm"].mean()
print(medias)''',
"prueba": r'''assert medias["Algarrobo"] == 35 and medias["Faique"] == 20, "Revise el agrupamiento."'''},
{"titulo": "Columna calculada", "nivel": 2,
"enunciado": "Agregue la columna `densidad` (árboles / área) y guarde en `max_parcela` el nombre de la parcela con mayor densidad.",
"plantilla": r'''import pandas as pd

df = pd.DataFrame({
    "parcela": ["P-01", "P-02", "P-03"],
    "arboles": [45, 38, 52],
    "area_ha": [0.5, 0.25, 0.8],
})
max_parcela = ""
print(df)
print(max_parcela)''',
"pista": "`df[\"densidad\"] = df[\"arboles\"] / df[\"area_ha\"]`; luego `df.loc[df[\"densidad\"].idxmax(), \"parcela\"]`.",
"solucion": r'''import pandas as pd

df = pd.DataFrame({
    "parcela": ["P-01", "P-02", "P-03"],
    "arboles": [45, 38, 52],
    "area_ha": [0.5, 0.25, 0.8],
})
df["densidad"] = df["arboles"] / df["area_ha"]
max_parcela = df.loc[df["densidad"].idxmax(), "parcela"]
print(df)
print(max_parcela)''',
"prueba": r'''assert "densidad" in df.columns, "Falta la columna densidad."
assert list(df["densidad"]) == [90.0, 152.0, 65.0], "Revise el cálculo de densidad."
assert max_parcela == "P-02", f"Se esperaba 'P-02'; se obtuvo {max_parcela!r}."'''},
],
"quiz": [
{"pregunta": "¿Qué ventaja tiene NumPy frente a una lista con bucle?", "opciones": ["Ninguna", "Opera sobre todos los elementos a la vez y es más rápido", "Solo guarda textos", "No requiere importarse"], "correcta": 1,
 "explicacion": "Las operaciones vectorizadas evitan bucles y son mucho más rápidas."},
{"pregunta": "¿Cómo se filtran las filas con dap_cm mayor a 30?", "opciones": ["df.dap_cm > 30", "df[df[\"dap_cm\"] > 30]", "df.filter(30)", "df[30]"], "correcta": 1,
 "explicacion": "La condición se coloca dentro de df[...] para seleccionar filas."},
{"pregunta": "¿A qué herramienta de Excel equivale groupby?", "opciones": ["Buscarv", "Tabla dinámica", "Formato condicional", "Gráfico"], "correcta": 1,
 "explicacion": "groupby agrupa y resume, como una tabla dinámica."},
{"pregunta": "¿Qué paquete adicional necesita pandas para guardar en Excel (.xlsx)?", "opciones": ["numpy", "openpyxl", "matplotlib", "json"], "correcta": 1,
 "explicacion": "to_excel usa openpyxl como motor para .xlsx."},
],
"glosario": {
 "NumPy": "Paquete para cálculo numérico con arreglos (arrays).",
 "Vectorización": "Aplicar una operación a todos los elementos sin bucles explícitos.",
 "pandas": "Paquete para manejar tablas de datos (DataFrame).",
 "DataFrame": "Tabla de pandas con filas y columnas con nombre.",
 "groupby": "Operación de agrupamiento y resumen por categorías.",
 "matplotlib": "Paquete para elaborar gráficos.",
},
})

# ---------------------------------------------------------------------------
MODULOS.append({
"id": "M14",
"titulo": "Python aplicado a SIG y proyecto integrador",
"duracion": "4 h",
"objetivos": [
    "Diferenciar coordenadas geográficas y UTM.",
    "Convertir coordenadas a UTM WGS84 Zona 17S (EPSG:32717) con pyproj.",
    "Calcular distancias, perímetros y áreas a partir de coordenadas UTM.",
    "Integrar lo aprendido en un pequeño procesador de inventario.",
],
"lecciones": [
{"titulo": "Coordenadas geográficas y UTM",
"texto": r'''- **Geográficas (EPSG:4326, WGS84):** latitud y longitud en grados. En el Perú la latitud es negativa (hemisferio sur) y la longitud también (oeste).
- **UTM (EPSG:32717, WGS84 Zona 17S):** Este y Norte en **metros**. Permite medir distancias y áreas directamente con geometría plana.

Para calcular distancias o áreas, primero convierta a UTM. El orden de los ejes importa: con `always_xy=True`, pyproj recibe **(longitud, latitud)** y devuelve **(Este, Norte)**.'''},
{"titulo": "Convertir coordenadas con pyproj",
"texto": r'''`pyproj` es el paquete estándar para transformar coordenadas (lo usan QGIS y geopandas internamente). Requiere instalación (`%pip install pyproj`), por lo que este ejemplo **solo se ejecuta en Jupyter o Colab**.

*Los puntos son de ejemplo.*''',
"codigo": r'''from pyproj import Transformer

a_utm = Transformer.from_crs("EPSG:4326", "EPSG:32717", always_xy=True)

puntos = {"Punto 1": (-80.63, -5.20), "Punto 2": (-79.97, -5.09)}  # (lon, lat)
for nombre, (lon, lat) in puntos.items():
    este, norte = a_utm.transform(lon, lat)
    print(f"{nombre}: E = {este:,.1f} m   N = {norte:,.1f} m")

# Conversión inversa: UTM -> geográficas
a_geo = Transformer.from_crs("EPSG:32717", "EPSG:4326", always_xy=True)
lon, lat = a_geo.transform(540000, 9425000)
print(f"lon = {lon:.6f}, lat = {lat:.6f}")''',
"web": False},
{"titulo": "Distancia y perímetro en UTM",
"texto": r'''En UTM la distancia entre dos puntos es la **distancia euclidiana** (Pitágoras). `math.dist(p1, p2)` la calcula directamente. El perímetro de un polígono es la suma de las distancias entre vértices consecutivos, cerrando con el primero.''',
"codigo": r'''import math

vertices = [(540000, 9425000), (540200, 9425000), (540200, 9425150), (540000, 9425150)]

perimetro = 0
for i in range(len(vertices)):
    p1 = vertices[i]
    p2 = vertices[(i + 1) % len(vertices)]   # el último se une con el primero
    perimetro += math.dist(p1, p2)
print(f"Perímetro: {perimetro:,.1f} m")'''},
{"titulo": "Área de un polígono (fórmula del zapatero)",
"texto": r'''Con los vértices en UTM, el área de un polígono simple se obtiene con la **fórmula del zapatero (shoelace)**:

Área = ½ · |Σ (xᵢ · yᵢ₊₁ − xᵢ₊₁ · yᵢ)|

Los vértices deben estar **en orden** (horario o antihorario) y el polígono no debe cruzarse a sí mismo.''',
"codigo": r'''def area_poligono(vertices):
    """Área (m²) de un polígono simple con vértices UTM ordenados."""
    n = len(vertices)
    suma = 0
    for i in range(n):
        x1, y1 = vertices[i]
        x2, y2 = vertices[(i + 1) % n]
        suma += x1 * y2 - x2 * y1
    return abs(suma) / 2

lote = [(540000, 9425000), (540200, 9425000), (540200, 9425150), (540000, 9425150)]
a = area_poligono(lote)
print(f"Área: {a:,.0f} m2 = {a / 10000:.2f} ha")'''},
{"titulo": "Ecosistema de Python para SIG",
"texto": r'''| Paquete | Uso | Ejemplo de tarea |
|---|---|---|
| **pyproj** | Sistemas de referencia | Convertir lat/lon a UTM 17S |
| **shapely** | Geometrías | Áreas, buffers, intersecciones |
| **geopandas** | Capas vectoriales como tablas | Leer shapefile/GeoPackage, `to_crs(32717)` |
| **rasterio** | Rásteres | Leer un DEM o un índice (NDVI, MSAVI) |
| **PyQGIS** | Automatizar QGIS | Procesos por lotes en la consola de QGIS |
| **openpyxl** | Excel | Reportes con formato |

Ruta sugerida después de este curso: geopandas → rasterio → PyQGIS.'''},
{"titulo": "Proyecto integrador: procesador de inventario",
"texto": r'''Este programa reúne lo aprendido: **lectura de CSV** (M11), **funciones** (M09), **diccionarios** (M08), **control de errores** (M10) y **f-strings** (M04). Lee un inventario, valida los datos, calcula área basal y volumen y emite un resumen por especie.

**Reto final:** adáptelo para (1) leer un archivo real con `open()`, (2) exportar el resumen a CSV, y (3) rehacerlo con pandas (M13).

*Datos ficticios; ff = 0.5 supuesto.*''',
"codigo": r'''import csv
import io
import math

DATOS = """codigo,especie,dap_cm,altura_m
A-01,Algarrobo,32.5,9.0
A-02,Faique,18.0,6.5
A-03,Algarrobo,45.2,11.0
A-04,Hualtaco,22.7,8.2
A-05,Faique,abc,7.0
A-06,Algarrobo,38.1,10.4
"""
FF = 0.5   # factor de forma supuesto

def area_basal(dap_cm):
    return math.pi / 4 * (dap_cm / 100) ** 2

def procesar(texto_csv):
    resumen = {}
    observados = []
    for fila in csv.DictReader(io.StringIO(texto_csv)):
        try:
            dap = float(fila["dap_cm"])
            alt = float(fila["altura_m"])
        except ValueError:
            observados.append(fila["codigo"])
            continue
        ab = area_basal(dap)
        r = resumen.setdefault(fila["especie"], {"n": 0, "ab": 0.0, "vol": 0.0})
        r["n"] += 1
        r["ab"] += ab
        r["vol"] += ab * alt * FF
    return resumen, observados

resumen, observados = procesar(DATOS)
print(f"{'Especie':<12}{'N':>4}{'AB (m2)':>10}{'Vol (m3)':>10}")
for especie, r in sorted(resumen.items()):
    print(f"{especie:<12}{r['n']:>4}{r['ab']:>10.4f}{r['vol']:>10.4f}")
print("Registros por verificar:", observados)'''},
],
"ejercicios": [
{"titulo": "Función distancia", "nivel": 1,
"enunciado": "Complete `distancia(p1, p2)` que devuelva la distancia en metros entre dos puntos UTM (tuplas).",
"plantilla": r'''def distancia(p1, p2):
    pass

print(distancia((540000, 9425000), (540300, 9425400)))''',
"pista": "Desempaquete: `x1, y1 = p1`, y aplique Pitágoras; o use `math.dist(p1, p2)`.",
"solucion": r'''import math

def distancia(p1, p2):
    return math.dist(p1, p2)

print(distancia((540000, 9425000), (540300, 9425400)))''',
"prueba": r'''assert abs(distancia((540000, 9425000), (540300, 9425400)) - 500) < 1e-9, "Debe ser 500 m."
assert distancia((0, 0), (0, 0)) == 0, "La distancia de un punto a sí mismo es 0."'''},
{"titulo": "Área en hectáreas", "nivel": 2,
"enunciado": "Complete `area_ha(vertices)` que devuelva el área del polígono en **hectáreas** con la fórmula del zapatero.",
"plantilla": r'''def area_ha(vertices):
    pass

triangulo = [(540000, 9425000), (540400, 9425000), (540000, 9425300)]
print(area_ha(triangulo))''',
"pista": "Reutilice el ciclo de la lección y divida el resultado final entre 10 000.",
"solucion": r'''def area_ha(vertices):
    n = len(vertices)
    suma = 0
    for i in range(n):
        x1, y1 = vertices[i]
        x2, y2 = vertices[(i + 1) % n]
        suma += x1 * y2 - x2 * y1
    return abs(suma) / 2 / 10000

triangulo = [(540000, 9425000), (540400, 9425000), (540000, 9425300)]
print(area_ha(triangulo))''',
"prueba": r'''tri = [(540000, 9425000), (540400, 9425000), (540000, 9425300)]
assert abs(area_ha(tri) - 6.0) < 1e-9, "El triángulo de 400 x 300 m tiene 6 ha."
cuadrado = [(0, 0), (100, 0), (100, 100), (0, 100)]
assert abs(area_ha(cuadrado) - 1.0) < 1e-9, "Un cuadrado de 100 m de lado tiene 1 ha."
assert abs(area_ha(list(reversed(cuadrado))) - 1.0) < 1e-9, "El sentido de los vértices no debe cambiar el área."'''},
{"titulo": "Perímetro", "nivel": 2,
"enunciado": "Complete `perimetro(vertices)` que devuelva el perímetro en metros de un polígono cerrado.",
"plantilla": r'''import math

def perimetro(vertices):
    pass

print(perimetro([(0, 0), (300, 0), (300, 400)]))''',
"pista": "Sume `math.dist(vertices[i], vertices[(i + 1) % n])` para cada i.",
"solucion": r'''import math

def perimetro(vertices):
    n = len(vertices)
    total = 0
    for i in range(n):
        total += math.dist(vertices[i], vertices[(i + 1) % n])
    return total

print(perimetro([(0, 0), (300, 0), (300, 400)]))''',
"prueba": r'''assert abs(perimetro([(0, 0), (300, 0), (300, 400)]) - 1200) < 1e-9, "300 + 400 + 500 = 1200 m."'''},
{"titulo": "Centro aproximado", "nivel": 2,
"enunciado": "Complete `promedio_vertices(vertices)` que devuelva la tupla (Este, Norte) promedio de los vértices. *Nota: en polígonos irregulares no es el centroide geométrico exacto.*",
"plantilla": r'''def promedio_vertices(vertices):
    pass

print(promedio_vertices([(0, 0), (200, 0), (200, 100), (0, 100)]))''',
"pista": "Use comprensiones: `sum(x for x, y in vertices) / len(vertices)` y lo mismo para y.",
"solucion": r'''def promedio_vertices(vertices):
    n = len(vertices)
    este = sum(x for x, y in vertices) / n
    norte = sum(y for x, y in vertices) / n
    return (este, norte)

print(promedio_vertices([(0, 0), (200, 0), (200, 100), (0, 100)]))''',
"prueba": r'''assert promedio_vertices([(0, 0), (200, 0), (200, 100), (0, 100)]) == (100, 50), "Se esperaba (100, 50)."'''},
],
"quiz": [
{"pregunta": "¿Qué código EPSG corresponde a UTM WGS84 Zona 17S?", "opciones": ["4326", "32717", "32617", "3857"], "correcta": 1,
 "explicacion": "EPSG:32717 = WGS84 / UTM zona 17S (326xx es hemisferio norte)."},
{"pregunta": "¿Por qué conviene convertir a UTM antes de medir áreas?", "opciones": ["Porque UTM usa metros y permite geometría plana", "Porque ocupa menos memoria", "Porque lo exige Python", "No conviene"], "correcta": 0,
 "explicacion": "Las coordenadas UTM están en metros; los grados no son unidades de longitud constantes."},
{"pregunta": "Con `always_xy=True`, ¿en qué orden recibe pyproj las coordenadas geográficas?", "opciones": ["(latitud, longitud)", "(longitud, latitud)", "(norte, este)", "Indistinto"], "correcta": 1,
 "explicacion": "always_xy fuerza el orden x, y: longitud, latitud."},
{"pregunta": "¿Qué requisito tiene la fórmula del zapatero?", "opciones": ["Vértices ordenados y polígono sin cruces", "Solo triángulos", "Coordenadas en grados", "Exactamente 4 vértices"], "correcta": 0,
 "explicacion": "Funciona para polígonos simples con vértices en orden."},
],
"glosario": {
 "EPSG": "Código que identifica un sistema de referencia de coordenadas.",
 "UTM": "Proyección que expresa coordenadas en metros (Este, Norte) por zonas.",
 "pyproj": "Paquete de Python para transformar coordenadas.",
 "Fórmula del zapatero": "Método para calcular el área de un polígono a partir de sus vértices.",
 "geopandas": "Extensión de pandas para capas geográficas vectoriales.",
},
})
