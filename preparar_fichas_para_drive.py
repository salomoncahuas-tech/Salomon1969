"""
IN Piura - Prepara las fichas PDF para subirlas a Google Drive
================================================================================
AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN
DIRECCION DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME
SUBDIRECCION DE ESTUDIOS DE INVERSION

Copia UNICAMENTE los PDF que empiezan con "ficha" desde la carpeta
"DIAGNOSTICOS TERRITORIALES BLOQUES" hacia una carpeta nueva y ligera,
conservando una subcarpeta por bloque. Asi se sube a Google Drive solo lo
necesario (unos pocos MB) en lugar de shapefiles, fotos y rasters.

Tambien copia el .md ya existente de cada ficha, si lo hubiera, para que
esas fichas se reconozcan como ya convertidas y no se reprocesen.

Uso:

    python preparar_fichas_para_drive.py "C:\\...\\DIAGNOSTICOS TERRITORIALES BLOQUES"

Por defecto crea la carpeta destino "FICHAS_PARA_DRIVE" junto a la carpeta
de origen. Con --destino se puede indicar otra ruta.

Opciones:
    --destino RUTA   Carpeta de salida (por defecto: FICHAS_PARA_DRIVE).
    --simular        Solo informa que copiaria, sin escribir nada.
    --patron ficha   Prefijo de nombre a buscar (por defecto "ficha").
    --sin-md         No copia los .md ya existentes.

Luego, en drive.google.com, arrastre la carpeta generada a "Mi unidad".
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

# Reutiliza las utilidades del conversor (mismo directorio del proyecto)
sys.path.insert(0, str(Path(__file__).resolve().parent))
from fichas_pdf_a_md import (  # noqa: E402
    PATRON_PREFIJO_DEFECTO,
    buscar_fichas,
    clave_natural,
)

NOMBRE_DESTINO_DEFECTO = "FICHAS_PARA_DRIVE"


def preparar(raiz: Path, destino: Path, prefijo: str, simular: bool,
             copiar_md: bool) -> int:
    """Copia las fichas (y sus .md) a la carpeta destino. Retorna codigo salida."""
    fichas = buscar_fichas(raiz, prefijo, recursivo=True)
    if not fichas:
        print("No se encontro ningun PDF que empiece con el prefijo indicado.")
        return 1

    print(f"Carpeta origen  : {raiz}")
    print(f"Carpeta destino : {destino}")
    print(f"Fichas '{prefijo}*': {len(fichas)} archivo(s) PDF")
    if simular:
        print("MODO SIMULACION : no se copiara ningun archivo")
    print("-" * 78)

    copiados = 0
    md_copiados = 0
    bytes_totales = 0

    for pdf_path in fichas:
        # Nombre de la subcarpeta de bloque tal como esta en el origen
        try:
            sub = pdf_path.parent.relative_to(raiz)
        except ValueError:
            sub = Path(pdf_path.parent.name)
        carpeta_destino = destino / sub if str(sub) != "." else destino

        md_origen = pdf_path.with_suffix(".md")
        tiene_md = copiar_md and md_origen.exists() and md_origen.stat().st_size > 0
        tam = pdf_path.stat().st_size
        bytes_totales += tam

        if not simular:
            carpeta_destino.mkdir(parents=True, exist_ok=True)
            shutil.copy2(pdf_path, carpeta_destino / pdf_path.name)
            if tiene_md:
                shutil.copy2(md_origen, carpeta_destino / md_origen.name)

        copiados += 1
        md_copiados += 1 if tiene_md else 0
        marca = " + .md" if tiene_md else ""
        print(f"[{copiados:3}/{len(fichas)}] {sub}/{pdf_path.name}"
              f"  ({tam / 1024:,.0f} KB){marca}")

    print("-" * 78)
    subcarpetas = len({p.parent for p in fichas})
    print(f"Fichas copiadas: {copiados} | .md ya convertidos: {md_copiados} | "
          f"Subcarpetas: {subcarpetas}")
    print(f"Tamano total   : {bytes_totales / (1024 * 1024):,.1f} MB")
    if not simular:
        print(f"\nListo. Arrastre la carpeta a drive.google.com > Mi unidad:\n  {destino}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Copia solo las fichas PDF (y sus .md) a una carpeta ligera "
                    "lista para subir a Google Drive.")
    parser.add_argument("raiz",
                        help="Carpeta 'DIAGNOSTICOS TERRITORIALES BLOQUES'.")
    parser.add_argument("--destino", default=None,
                        help=f"Carpeta de salida (por defecto '{NOMBRE_DESTINO_DEFECTO}' "
                             "junto a la carpeta de origen).")
    parser.add_argument("--patron", default=PATRON_PREFIJO_DEFECTO,
                        help="Prefijo del nombre de archivo (por defecto 'ficha').")
    parser.add_argument("--simular", action="store_true",
                        help="No copia nada; solo informa (dry-run).")
    parser.add_argument("--sin-md", dest="copiar_md", action="store_false",
                        help="No copia los .md ya existentes.")
    args = parser.parse_args(argv)

    raiz = Path(args.raiz).expanduser().resolve()
    if not raiz.is_dir():
        print(f"ERROR: la ruta no existe o no es una carpeta: {raiz}", file=sys.stderr)
        return 3

    destino = (Path(args.destino).expanduser().resolve() if args.destino
               else raiz.parent / NOMBRE_DESTINO_DEFECTO)
    if destino == raiz or raiz in destino.parents:
        print("ERROR: la carpeta destino no puede estar dentro de la de origen.",
              file=sys.stderr)
        return 3

    return preparar(raiz, destino, args.patron, args.simular, args.copiar_md)


if __name__ == "__main__":
    raise SystemExit(main())
