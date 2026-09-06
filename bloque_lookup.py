"""
Resolucion del codigo de bloque contra la lista de bloques del aplicativo.

Los codigos de bloque son TEXTO y de formato mixto: numericos ('1', '4', '14',
'128') y con prefijo de microcuenca ('M9B1', 'M6B2-3'). Buscar el bloque con
una coincidencia parcial (`codigo in label`) es incorrecto: el codigo '4' esta
contenido en '14', '24', '34', '40'..., y en un desplegable ordenado
alfabeticamente el primer candidato es '14'. Por eso una ficha del bloque 4
terminaba enlazada al bloque 14.

Este modulo resuelve SIEMPRE por igualdad exacta sobre el codigo normalizado.
Si no hay coincidencia exacta devuelve None: nunca se adivina un bloque.
"""

import re
import unicodedata


def normalizar_codigo_bloque(valor):
    """Normaliza un codigo de bloque para poder compararlo por igualdad.

    Acepta int, float ('4.0' que devuelve Excel) y str. Quita espacios,
    acentos y el prefijo redundante 'Bloque'; mayusculiza; y en los codigos
    puramente numericos elimina los ceros a la izquierda ('04' -> '4').
    """
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    txt = str(valor).strip()
    if not txt:
        return ""
    txt = unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("ascii")
    txt = txt.upper().replace(" ", "").replace("_", "")
    txt = re.sub(r"^BLOQUES?[-]?", "", txt)
    if txt.isdigit():
        txt = str(int(txt))
    return txt


def codigo_de_label(label):
    """Extrae el codigo del label del desplegable, con formato 'CODIGO - Tipo'.

    Los codigos con guion ('M6B2-3') se conservan intactos porque el separador
    del label es ' - ' (con espacios).
    """
    texto = str(label or "")
    return texto.split(" - ")[0] if " - " in texto else texto


def buscar_label_bloque(codigo, bloques_map):
    """Devuelve el label de `bloques_map` cuyo codigo coincide exactamente.

    `bloques_map` es cualquier iterable de labels 'CODIGO - Tipo' (por ejemplo
    el dict {label: id} que arma el aplicativo). Devuelve None si el codigo no
    existe en la lista.
    """
    objetivo = normalizar_codigo_bloque(codigo)
    if not objetivo or not bloques_map:
        return None
    for label in bloques_map:
        if normalizar_codigo_bloque(codigo_de_label(label)) == objetivo:
            return label
    return None
