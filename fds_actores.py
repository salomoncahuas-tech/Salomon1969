"""
F-DS-02 — Registro de actores: columnas y compatibilidad con registros antiguos.
Proyecto IN Piura CUI 2669244 | ANIN - DIME - SESDI.

La tabla de actores tenia una sola columna "Nombre del actor / Organizacion";
la barra hacia que los reportes la trataran como un unico campo (el nombre).
Desde esta version se registran dos columnas separadas:

    "Nombre del actor"  y  "Cargo"

Los registros guardados con la columna antigua NO se reescriben en la base de
datos: se normalizan al leerlos (formulario de edicion, historial, PDF, Excel
y analitica). El texto antiguo se conserva integro en "Nombre del actor" y
"Cargo" queda vacio para completarlo al editar la ficha.
"""

COL_NOMBRE = "Nombre del actor"
COL_CARGO = "Cargo"

# Encabezados que tuvo la primera columna en versiones anteriores.
COLS_LEGACY = ("Nombre del actor / Organizacion",
               "Nombre del actor / Organización")

# Columnas del editor (st.data_editor) en el orden en que se muestran.
COLUMNAS = [
    COL_NOMBRE, COL_CARGO, "Tipo", "Rol / Funcion frente al proyecto",
    "Influencia", "Interes", "Posicion", "Nivel territorial", "Telefono",
    "Correo / Contacto", "Observaciones / Historial",
]

SLOT = "f2_actores"


def _txt(valor):
    return "" if valor is None else str(valor).strip()


def migrar_fila(fila):
    """Devuelve la fila con "Nombre del actor" y "Cargo" separados.

    Acepta filas nuevas, filas con la columna antigua y filas del formato
    legacy de claves cortas (`nombre`, `cargo`). No pierde ningun campo: las
    columnas que no conoce se conservan al final en su orden original.
    """
    if not isinstance(fila, dict):
        return fila
    nombre = _txt(fila.get(COL_NOMBRE))
    if not nombre:
        for col in COLS_LEGACY + ("nombre",):
            nombre = _txt(fila.get(col))
            if nombre:
                break
    cargo = _txt(fila.get(COL_CARGO)) or _txt(fila.get("cargo"))
    salida = {COL_NOMBRE: nombre, COL_CARGO: cargo}
    for k, v in fila.items():
        if k in COLS_LEGACY or k in (COL_NOMBRE, COL_CARGO, "nombre", "cargo"):
            continue
        salida[k] = v
    return salida


def migrar_filas(filas):
    """Normaliza una lista de filas de actores (ver `migrar_fila`)."""
    if not isinstance(filas, list):
        return filas
    return [migrar_fila(f) for f in filas]


def migrar_tabla(slot, filas):
    """Aplica la migracion solo a la tabla de actores; el resto pasa igual."""
    return migrar_filas(filas) if slot == SLOT else filas
