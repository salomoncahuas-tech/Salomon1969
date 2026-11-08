"""Lectura del catálogo V6 de bloques del Proyecto IN Piura.

    from anin_catalogo import cargar_bloques
    bloques = cargar_bloques()          # lista de dicts, solo los 117 bloques vigentes

El catálogo (datos/unidades_liberacion_areas.csv) trae provincia y distrito como
«slugs» sin tildes (san_miguel_de_el_faique); aquí se convierten a su nombre oficial.
También incluye 60 lotes SUS (tipo_unidad = lote_sus) que no son bloques y se excluyen
por defecto.
"""
import csv
from pathlib import Path

PROVINCIAS = {"ayabaca": "Ayabaca", "huancabamba": "Huancabamba", "morropon": "Morropón"}
DISTRITOS = {
    "frias": "Frías", "canchaque": "Canchaque", "huarmaca": "Huarmaca",
    "huancabamba": "Huancabamba", "lalaquiz": "Lalaquiz",
    "san_miguel_de_el_faique": "San Miguel de El Faique", "buenos_aires": "Buenos Aires",
    "chalaco": "Chalaco", "chulucanas": "Chulucanas", "morropon": "Morropón",
    "salitral": "Salitral", "san_juan_de_bigote": "San Juan de Bigote",
    "santa_catalina_de_mossa": "Santa Catalina de Mossa", "santo_domingo": "Santo Domingo",
    "yamango": "Yamango",
}
AREA_TOTAL_HA = 12270.235  # suma de los 117 bloques preliminares


def _raiz_repo():
    # .../.claude/skills/anin-in-piura/scripts/anin_catalogo.py -> raíz del repositorio
    return Path(__file__).resolve().parents[4]


def cargar_bloques(ruta=None, incluir_lotes_sus=False):
    """Devuelve los bloques del catálogo con nombres oficiales y área como float.

    Cada elemento: codigo, provincia, distrito, area_ha, asistente, asistente_nombre,
    tipo_unidad. Ordenados por provincia, distrito y código."""
    ruta = Path(ruta) if ruta else _raiz_repo() / "datos" / "unidades_liberacion_areas.csv"
    salida = []
    with open(ruta, encoding="utf-8", newline="") as f:
        for r in csv.DictReader(f):
            if r["tipo_unidad"] != "bloque" and not incluir_lotes_sus:
                continue
            salida.append({
                "codigo": r["name"],
                "provincia": PROVINCIAS[r["provincia"]],
                "distrito": DISTRITOS[r["distrito"]],
                "area_ha": float(r["area_ha"]),
                "asistente": r["asistente"],
                "asistente_nombre": r["asistente_nombre"],
                "tipo_unidad": r["tipo_unidad"],
            })
    salida.sort(key=lambda b: (b["provincia"], b["distrito"], b["codigo"]))
    return salida


if __name__ == "__main__":
    b = cargar_bloques()
    total = sum(x["area_ha"] for x in b)
    print(f"{len(b)} bloques · {total:,.3f} ha (esperado {AREA_TOTAL_HA:,.3f}) · "
          f"{len({x['distrito'] for x in b})} distritos")
