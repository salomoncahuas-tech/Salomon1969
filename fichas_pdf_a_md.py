"""
IN Piura - Conversor masivo de FICHAS de diagnostico territorial: PDF -> Markdown
================================================================================
AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN
DIRECCION DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME
SUBDIRECCION DE ESTUDIOS DE INVERSION

Recorre la carpeta "DIAGNOSTICOS TERRITORIALES BLOQUES" (una subcarpeta por
bloque) y convierte a Markdown todo archivo PDF cuyo nombre empiece con
"ficha". El .md se guarda en la MISMA subcarpeta del PDF de origen.

Si el .md ya existe (y no esta vacio) el archivo se OMITE y el proceso
continua con los demas. Con --forzar se regeneran todos.

Uso tipico (Windows, carpeta sincronizada de OneDrive):

    python fichas_pdf_a_md.py "C:\\Users\\Hector\\OneDrive\\DIAGNOSTICOS\\DIAGNOSTICOS TERRITORIALES BLOQUES"

Sin argumento intenta autodetectar la carpeta dentro de OneDrive:

    python fichas_pdf_a_md.py

Opciones:
    --forzar            Regenera el .md aunque ya exista.
    --simular           No escribe nada; solo informa que haria (dry-run).
    --patron ficha      Prefijo de nombre a buscar (por defecto "ficha").
    --sin-recursivo     Solo el primer nivel de subcarpetas.
    --sin-reporte       No genera el reporte consolidado en la carpeta raiz.
    --ocr               Intenta OCR en paginas sin texto (requiere pytesseract).

Dependencias: pdfplumber  (pip install pdfplumber)
Opcional OCR: pytesseract, pdf2image + Tesseract y Poppler instalados.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

try:
    import pdfplumber
    PDFPLUMBER_OK = True
except ImportError:  # pragma: no cover - depende del entorno del usuario
    PDFPLUMBER_OK = False


# ── Constantes institucionales ANIN ──────────────────────────────────────
ENCABEZADO_ANIN = (
    "AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN",
    "DIRECCION DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME",
    "SUBDIRECCION DE ESTUDIOS DE INVERSION",
)
PROYECTO = "Proyecto IN Piura - CUI 2669244 - Cuenca Alta del Rio Piura"

PATRON_PREFIJO_DEFECTO = "ficha"
NOMBRE_REPORTE = "_REPORTE_CONVERSION_FICHAS.md"

# Estados posibles de cada archivo procesado
OK = "convertido"
OMITIDO = "omitido (md ya existe)"
SIN_TEXTO = "sin texto extraible (posible PDF escaneado)"
ERROR = "error"


# ── Utilidades de texto ──────────────────────────────────────────────────
def sin_tildes(texto: str) -> str:
    """Devuelve el texto en minusculas y sin tildes, para comparaciones."""
    norm = unicodedata.normalize("NFD", texto)
    return "".join(c for c in norm if unicodedata.category(c) != "Mn").lower()


def es_ficha(nombre_archivo: str, prefijo: str) -> bool:
    """True si el nombre del PDF empieza con el prefijo buscado ('ficha...')."""
    base = sin_tildes(nombre_archivo).lstrip("_-. ")
    return base.startswith(sin_tildes(prefijo))


def escapar_md(texto: str) -> str:
    """Escapa los caracteres que romperian una tabla Markdown."""
    return (texto or "").replace("|", "\\|").replace("\n", "<br>").strip()


def limpiar_lineas(texto: str) -> list[str]:
    """Normaliza espacios y elimina lineas en blanco consecutivas."""
    lineas: list[str] = []
    for cruda in (texto or "").splitlines():
        linea = re.sub(r"[ \t\u00a0]+", " ", cruda).strip()
        if not linea and lineas and not lineas[-1]:
            continue  # evita dobles lineas vacias
        lineas.append(linea)
    while lineas and not lineas[-1]:
        lineas.pop()
    return lineas


_RE_TITULO_NUM = re.compile(r"^(?:[IVXLC]+|\d+(?:\.\d+)*)[.)\-]\s+\S")


def es_titulo(linea: str) -> bool:
    """Heuristica: detecta encabezados de seccion tipicos de las fichas."""
    if not (3 < len(linea) <= 90) or linea.endswith((".", ",", ";", ":")):
        return bool(linea.endswith(":") and len(linea) <= 60 and linea.isupper())
    if _RE_TITULO_NUM.match(linea):
        return True
    letras = [c for c in linea if c.isalpha()]
    return bool(letras) and all(c.isupper() for c in letras) and len(letras) >= 4


def formatear_parrafos(texto: str) -> list[str]:
    """Convierte el texto plano de una franja del PDF en lineas Markdown."""
    salida: list[str] = []
    for linea in limpiar_lineas(texto):
        if not linea:
            if salida and salida[-1]:
                salida.append("")
            continue
        if es_titulo(linea):
            if salida and salida[-1]:
                salida.append("")
            salida.append(f"### {linea.rstrip(':')}")
            salida.append("")
        else:
            salida.append(linea)
    return salida


def tabla_a_markdown(tabla: list[list]) -> list[str]:
    """Renderiza una tabla extraida por pdfplumber como tabla Markdown."""
    filas = [
        [escapar_md("" if celda is None else str(celda)) for celda in fila]
        for fila in (tabla or [])
        if fila is not None
    ]
    filas = [f for f in filas if any(c for c in f)]
    if not filas:
        return []

    ancho = max(len(f) for f in filas)
    filas = [f + [""] * (ancho - len(f)) for f in filas]

    cabecera = filas[0]
    cuerpo = filas[1:]
    if not any(cabecera):  # primera fila vacia -> cabecera generica
        cabecera = [f"Col {i + 1}" for i in range(ancho)]
        cuerpo = filas
    if not cuerpo:  # tabla de una sola fila: se muestra como datos
        cuerpo = [cabecera]
        cabecera = [f"Col {i + 1}" for i in range(ancho)]

    lineas = ["| " + " | ".join(cabecera) + " |",
              "|" + "|".join([" --- "] * ancho) + "|"]
    lineas += ["| " + " | ".join(fila) + " |" for fila in cuerpo]
    lineas.append("")
    return lineas


# ── Extraccion del PDF ───────────────────────────────────────────────────
def _texto_franja(pagina, y0: float, y1: float) -> str:
    """Texto de una franja horizontal de la pagina, tolerante a errores."""
    if y1 - y0 < 2:
        return ""
    try:
        franja = pagina.crop((0, max(0, y0), pagina.width, min(pagina.height, y1)))
        return franja.extract_text() or ""
    except Exception:
        return ""


def contenido_pagina(pagina, ocr: bool = False) -> tuple[list[str], bool]:
    """
    Devuelve las lineas Markdown de una pagina respetando el orden vertical
    entre parrafos y tablas, y si se hallo algun contenido.
    """
    try:
        tablas = sorted(pagina.find_tables(), key=lambda t: t.bbox[1])
    except Exception:
        tablas = []

    lineas: list[str] = []
    hubo_contenido = False
    cursor = 0.0

    for tabla in tablas:
        _, top, _, bottom = tabla.bbox
        previo = _texto_franja(pagina, cursor, top)
        if previo.strip():
            lineas += formatear_parrafos(previo)
            hubo_contenido = True
        try:
            md_tabla = tabla_a_markdown(tabla.extract())
        except Exception:
            md_tabla = []
        if md_tabla:
            if lineas and lineas[-1]:
                lineas.append("")
            lineas += md_tabla
            hubo_contenido = True
        cursor = max(cursor, bottom)

    final = _texto_franja(pagina, cursor, pagina.height) if tablas else (
        pagina.extract_text() or ""
    )
    if final.strip():
        lineas += formatear_parrafos(final)
        hubo_contenido = True

    if not hubo_contenido and ocr:
        texto_ocr = _ocr_pagina(pagina)
        if texto_ocr.strip():
            lineas.append("> _Contenido obtenido por OCR; verificar antes de citar._")
            lineas.append("")
            lineas += formatear_parrafos(texto_ocr)
            hubo_contenido = True

    while lineas and not lineas[-1]:
        lineas.pop()
    return lineas, hubo_contenido


def _ocr_pagina(pagina) -> str:
    """OCR opcional de una pagina (solo si pytesseract esta disponible)."""
    try:
        import pytesseract  # type: ignore
    except ImportError:
        return ""
    try:
        imagen = pagina.to_image(resolution=300).original
        return pytesseract.image_to_string(imagen, lang="spa")
    except Exception:
        return ""


def codigo_bloque(carpeta: Path) -> str:
    """Extrae 'BLOQUE 12' del nombre de la subcarpeta, si es posible."""
    m = re.search(r"bloque\s*[-_]?\s*(\d+)", sin_tildes(carpeta.name))
    return f"BLOQUE {int(m.group(1))}" if m else carpeta.name


def construir_markdown(pdf_path: Path, ocr: bool = False) -> tuple[str, int, int]:
    """
    Convierte un PDF de ficha a texto Markdown.
    Retorna (contenido, total_paginas, paginas_con_texto).
    """
    partes: list[str] = []
    paginas_con_texto = 0

    with pdfplumber.open(str(pdf_path)) as pdf:
        total = len(pdf.pages)
        for i, pagina in enumerate(pdf.pages, start=1):
            lineas, hubo = contenido_pagina(pagina, ocr=ocr)
            if hubo:
                paginas_con_texto += 1
            partes.append(f"## Pagina {i} de {total}")
            partes.append("")
            partes += lineas if lineas else ["_(pagina sin texto extraible)_"]
            partes.append("")
            if i < total:
                partes.append("---")
                partes.append("")

    encabezado = [
        f"# {pdf_path.stem}",
        "",
        f"> **{ENCABEZADO_ANIN[0]}**  ",
        f"> {ENCABEZADO_ANIN[1]}  ",
        f"> {ENCABEZADO_ANIN[2]}  ",
        f"> {PROYECTO}",
        "",
        "| Campo | Valor |",
        "| --- | --- |",
        f"| Bloque / carpeta | {codigo_bloque(pdf_path.parent)} |",
        f"| Archivo de origen | `{pdf_path.name}` |",
        f"| Paginas | {total} |",
        f"| Paginas con texto | {paginas_con_texto} |",
        f"| Fecha de conversion | {datetime.now():%d/%m/%Y %H:%M} |",
        "",
        "---",
        "",
    ]
    return "\n".join(encabezado + partes).rstrip() + "\n", total, paginas_con_texto


# ── Recorrido de carpetas ────────────────────────────────────────────────
def clave_natural(texto: str) -> list:
    """Clave de ordenamiento natural: 'BLOQUE 2' antes que 'BLOQUE 18'."""
    return [int(p) if p.isdigit() else p
            for p in re.split(r"(\d+)", sin_tildes(texto))]


def buscar_fichas(raiz: Path, prefijo: str, recursivo: bool) -> list[Path]:
    """Lista los PDF de ficha bajo la raiz, ordenados por carpeta y nombre."""
    patron = "**/*" if recursivo else "*/*"
    encontrados = [
        p for p in raiz.glob(patron)
        if p.is_file() and p.suffix.lower() == ".pdf" and es_ficha(p.name, prefijo)
    ]
    # Tambien PDFs sueltos en la propia raiz
    encontrados += [
        p for p in raiz.glob("*.pdf")
        if p.is_file() and es_ficha(p.name, prefijo)
    ]
    unicos = {p.resolve(): p for p in encontrados}
    return sorted(unicos.values(),
                  key=lambda p: (clave_natural(str(p.parent)), clave_natural(p.name)))


def autodetectar_raiz() -> Path | None:
    """Busca 'DIAGNOSTICOS TERRITORIALES BLOQUES' en las rutas de OneDrive."""
    candidatos: list[Path] = []
    for var in ("OneDrive", "OneDriveConsumer", "OneDriveCommercial"):
        valor = os.environ.get(var)
        if valor:
            candidatos.append(Path(valor))
    inicio = Path.home()
    candidatos += [p for p in inicio.glob("OneDrive*") if p.is_dir()]
    candidatos.append(inicio)

    objetivo = sin_tildes("diagnosticos territoriales bloques")
    vistos: set[Path] = set()
    for base in candidatos:
        base = base.expanduser()
        if not base.is_dir() or base in vistos:
            continue
        vistos.add(base)
        for hallado in base.glob("**/*"):
            try:
                if hallado.is_dir() and sin_tildes(hallado.name) == objetivo:
                    return hallado
            except OSError:
                continue
    return None


def escribir_reporte(raiz: Path, resultados: list[dict], simular: bool) -> Path | None:
    """Genera el reporte consolidado de la corrida en la carpeta raiz."""
    if simular:
        return None
    resumen = {estado: 0 for estado in (OK, OMITIDO, SIN_TEXTO, ERROR)}
    for r in resultados:
        resumen[r["estado"]] += 1

    lineas = [
        "# Reporte de conversion de fichas PDF a Markdown",
        "",
        f"> **{ENCABEZADO_ANIN[0]}**  ",
        f"> {ENCABEZADO_ANIN[1]}  ",
        f"> {ENCABEZADO_ANIN[2]}",
        "",
        f"- Carpeta procesada: `{raiz}`",
        f"- Fecha de ejecucion: {datetime.now():%d/%m/%Y %H:%M}",
        f"- Archivos evaluados: {len(resultados)}",
        "",
        "## Resumen",
        "",
        "| Estado | Archivos |",
        "| --- | --- |",
    ]
    lineas += [f"| {estado} | {cantidad} |" for estado, cantidad in resumen.items()]
    lineas += [
        "",
        "## Detalle",
        "",
        "| Bloque / carpeta | Archivo PDF | Estado | Paginas | Con texto | Observacion |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for r in resultados:
        lineas.append(
            f"| {r['bloque']} | {r['archivo']} | {r['estado']} | "
            f"{r['paginas'] or '-'} | {r['con_texto'] or '-'} | "
            f"{escapar_md(r['detalle']) or '-'} |"
        )
    destino = raiz / NOMBRE_REPORTE
    destino.write_text("\n".join(lineas) + "\n", encoding="utf-8")
    return destino


def procesar(raiz: Path, prefijo: str, forzar: bool, simular: bool,
             recursivo: bool, reporte: bool, ocr: bool) -> int:
    """Ejecuta la conversion completa. Retorna el codigo de salida."""
    fichas = buscar_fichas(raiz, prefijo, recursivo)
    subcarpetas = sorted({p.parent for p in fichas})
    print(f"Carpeta raiz     : {raiz}")
    print(f"Subcarpetas      : {len(subcarpetas)} con fichas")
    print(f"Fichas '{prefijo}*' : {len(fichas)} archivo(s) PDF")
    if simular:
        print("MODO SIMULACION  : no se escribira ningun archivo")
    print("-" * 78)

    if not fichas:
        print("No se encontro ningun PDF que empiece con el prefijo indicado.")
        return 1

    resultados: list[dict] = []
    for n, pdf_path in enumerate(fichas, start=1):
        md_path = pdf_path.with_suffix(".md")
        rel = pdf_path.relative_to(raiz) if raiz in pdf_path.parents else pdf_path
        registro = {
            "bloque": codigo_bloque(pdf_path.parent),
            "archivo": pdf_path.name,
            "estado": ERROR,
            "paginas": 0,
            "con_texto": 0,
            "detalle": "",
        }

        if md_path.exists() and md_path.stat().st_size > 0 and not forzar:
            registro["estado"] = OMITIDO
            registro["detalle"] = f"ya existe {md_path.name}"
            resultados.append(registro)
            print(f"[{n:3}/{len(fichas)}] OMITIDO    {rel}  (ya existe {md_path.name})")
            continue

        try:
            contenido, paginas, con_texto = construir_markdown(pdf_path, ocr=ocr)
            registro["paginas"] = paginas
            registro["con_texto"] = con_texto
            registro["estado"] = OK if con_texto else SIN_TEXTO
            if not simular:
                md_path.write_text(contenido, encoding="utf-8")
            if not con_texto:
                registro["detalle"] = "revisar manualmente / considerar --ocr"
            etiqueta = "CONVERTIDO" if con_texto else "SIN TEXTO "
            print(f"[{n:3}/{len(fichas)}] {etiqueta} {rel}  "
                  f"({con_texto}/{paginas} pag. con texto)")
        except Exception as exc:  # PDF corrupto, protegido, etc.
            registro["estado"] = ERROR
            registro["detalle"] = str(exc)
            print(f"[{n:3}/{len(fichas)}] ERROR      {rel}  -> {exc}")
        resultados.append(registro)

    print("-" * 78)
    convertidos = sum(1 for r in resultados if r["estado"] == OK)
    omitidos = sum(1 for r in resultados if r["estado"] == OMITIDO)
    sin_texto = sum(1 for r in resultados if r["estado"] == SIN_TEXTO)
    errores = sum(1 for r in resultados if str(r["estado"]).startswith(ERROR))
    print(f"Convertidos: {convertidos} | Omitidos: {omitidos} | "
          f"Sin texto: {sin_texto} | Errores: {errores}")

    if reporte:
        destino = escribir_reporte(raiz, resultados, simular)
        if destino:
            print(f"Reporte generado: {destino}")

    return 0 if errores == 0 else 2


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Convierte a Markdown las fichas PDF de diagnostico "
                    "territorial por bloque (Proyecto IN Piura - ANIN).")
    parser.add_argument("raiz", nargs="?", default=None,
                        help="Carpeta 'DIAGNOSTICOS TERRITORIALES BLOQUES'. "
                             "Si se omite se intenta autodetectar en OneDrive.")
    parser.add_argument("--patron", default=PATRON_PREFIJO_DEFECTO,
                        help="Prefijo del nombre de archivo (por defecto 'ficha').")
    parser.add_argument("--forzar", action="store_true",
                        help="Regenera el .md aunque ya exista.")
    parser.add_argument("--simular", action="store_true",
                        help="No escribe archivos; solo informa (dry-run).")
    parser.add_argument("--sin-recursivo", dest="recursivo", action="store_false",
                        help="Solo el primer nivel de subcarpetas.")
    parser.add_argument("--sin-reporte", dest="reporte", action="store_false",
                        help="No genera el reporte consolidado.")
    parser.add_argument("--ocr", action="store_true",
                        help="Intenta OCR en paginas sin texto (requiere pytesseract).")
    args = parser.parse_args(argv)

    if not PDFPLUMBER_OK:
        print("ERROR: falta la libreria pdfplumber. Instalela con:\n"
              "       pip install pdfplumber", file=sys.stderr)
        return 3

    if args.raiz:
        raiz = Path(args.raiz).expanduser()
    else:
        print("Buscando la carpeta 'DIAGNOSTICOS TERRITORIALES BLOQUES' en OneDrive...")
        detectada = autodetectar_raiz()
        if detectada is None:
            print("ERROR: no se pudo autodetectar la carpeta. Indiquela como "
                  "argumento:\n       python fichas_pdf_a_md.py \"<ruta de la "
                  "carpeta>\"", file=sys.stderr)
            return 3
        raiz = detectada

    if not raiz.is_dir():
        print(f"ERROR: la ruta no existe o no es una carpeta: {raiz}", file=sys.stderr)
        return 3

    return procesar(raiz, args.patron, args.forzar, args.simular,
                    args.recursivo, args.reporte, args.ocr)


if __name__ == "__main__":
    raise SystemExit(main())
