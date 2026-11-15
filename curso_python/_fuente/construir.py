# -*- coding: utf-8 -*-
"""Genera los materiales del curso "Python desde cero" a partir del contenido.

Salidas (en la carpeta curso_python/):
  cuadernos/MNN_*.ipynb             lecciones, ejemplos y ejercicios
  soluciones/MNN_*_soluciones.ipynb soluciones comprobadas
  curso_python_interactivo.html     curso interactivo (Python en el navegador)

Uso:  python construir.py
Requisitos: nbformat, markdown, pygments, numpy, pandas, matplotlib, pyproj
"""
import base64
import builtins
import contextlib
import io
import json
import linecache
import os
import re
import sys
import tempfile
import traceback
import unicodedata
import warnings

os.environ.setdefault("MPLBACKEND", "Agg")
warnings.filterwarnings("ignore")

import markdown
import nbformat
from nbformat.v4 import new_code_cell, new_markdown_cell, new_notebook, new_output
from pygments import highlight
from pygments.formatters import HtmlFormatter
from pygments.lexers import PythonLexer

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
sys.path.insert(0, AQUI)

from contenido_a import MODULOS as _A  # noqa: E402
from contenido_b import MODULOS as _B  # noqa: E402
from contenido_c import MODULOS as _C  # noqa: E402

MODULOS = _A + _B + _C
NIVELES = {1: "Básico", 2: "Intermedio", 3: "Reto"}
ESTRELLAS = {1: "★☆☆", 2: "★★☆", 3: "★★★"}
TITULO_CURSO = "Python desde cero"


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def slug(texto):
    t = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    t = re.sub(r"[^a-zA-Z0-9]+", "_", t).strip("_").lower()
    return t[:40].rstrip("_")


def num(mod):
    return int(mod["id"][1:])


def md_html(texto):
    return markdown.markdown(texto, extensions=["tables", "fenced_code", "sane_lists"])


def resaltar(codigo):
    return highlight(codigo, PythonLexer(), HtmlFormatter(nowrap=True))


def ejecutar(codigo, entrada=None):
    """Ejecuta código en un espacio limpio. Devuelve (salidas_nb, texto, error, imagenes)."""
    import matplotlib.pyplot as plt

    plt.close("all")
    linecache.cache["<celda>"] = (len(codigo), None, codigo.splitlines(True), "<celda>")
    ns = {"__name__": "__main__"}
    buf = io.StringIO()
    error = None
    err_nb = None
    lineas = (entrada or "").split("\n")
    input_original = builtins.input

    def _input(msg=""):
        valor = lineas.pop(0)
        print(msg + valor)
        return valor

    builtins.input = _input
    try:
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            exec(compile(codigo, "<celda>", "exec"), ns)
    except SyntaxError as e:
        tb = "".join(traceback.format_exception_only(type(e), e))
        error = tb.rstrip()
        err_nb = (type(e).__name__, str(e.msg), tb.rstrip().split("\n"))
    except Exception as e:  # noqa: BLE001
        marcos = [f for f in traceback.extract_tb(e.__traceback__) if f.filename == "<celda>"]
        tb = "Traceback (most recent call last):\n" + "".join(traceback.format_list(marcos))
        tb += "".join(traceback.format_exception_only(type(e), e))
        error = tb.rstrip()
        err_nb = (type(e).__name__, str(e), tb.rstrip().split("\n"))
    finally:
        builtins.input = input_original

    texto = buf.getvalue()
    salidas = []
    if texto:
        salidas.append(new_output("stream", name="stdout", text=texto))
    imagenes = []
    for n in plt.get_fignums():
        fig = plt.figure(n)
        png = io.BytesIO()
        fig.savefig(png, format="png", dpi=110, bbox_inches="tight")
        b64 = base64.b64encode(png.getvalue()).decode()
        imagenes.append(b64)
        salidas.append(new_output("display_data", data={"image/png": b64, "text/plain": "<Figure>"}))
    plt.close("all")
    if err_nb:
        salidas.append(new_output("error", ename=err_nb[0], evalue=err_nb[1], traceback=err_nb[2]))
    return salidas, texto, error, imagenes, ns


def comprobar(codigo, prueba):
    """True si el código supera la prueba."""
    _, _, error, _, ns = ejecutar(codigo)
    if error:
        return False, error
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            exec(compile(prueba, "<prueba>", "exec"), ns)
        return True, None
    except Exception as e:  # noqa: BLE001
        return False, f"{type(e).__name__}: {e}"


# ---------------------------------------------------------------------------
# 1. Ejecutar ejemplos y validar ejercicios
# ---------------------------------------------------------------------------
def procesar():
    problemas = []
    dir_trabajo = tempfile.mkdtemp(prefix="curso_py_")
    cwd = os.getcwd()
    os.chdir(dir_trabajo)
    try:
        for m in MODULOS:
            m.setdefault("web", True)
            for i, lec in enumerate(m["lecciones"]):
                lec.setdefault("web", m["web"])
                if "codigo" not in lec:
                    continue
                salidas, texto, error, imgs, _ = ejecutar(lec["codigo"], lec.get("entrada"))
                lec["_salidas"], lec["_texto"], lec["_error"], lec["_imgs"] = salidas, texto, error, imgs
                if error and not lec.get("error"):
                    problemas.append(f"{m['id']} lección {i + 1}: error inesperado\n{error}")
                if lec.get("error") and not error:
                    problemas.append(f"{m['id']} lección {i + 1}: se esperaba un error")
            for j, ej in enumerate(m["ejercicios"]):
                ej.setdefault("web", m["web"])
                ok, msg = comprobar(ej["solucion"], ej["prueba"])
                if not ok:
                    problemas.append(f"{m['id']} ejercicio {j + 1}: la solución NO pasa la prueba: {msg}")
                ok_p, _ = comprobar(ej["plantilla"], ej["prueba"])
                if ok_p:
                    problemas.append(f"{m['id']} ejercicio {j + 1}: la plantilla ya pasa la prueba")
                salidas, texto, error, imgs, _ = ejecutar(ej["solucion"])
                ej["_salidas_sol"], ej["_texto_sol"] = salidas, texto
            for q in m["quiz"]:
                assert 0 <= q["correcta"] < len(q["opciones"]), (m["id"], q["pregunta"])
    finally:
        os.chdir(cwd)
    return problemas


# ---------------------------------------------------------------------------
# 2. Cuadernos Jupyter
# ---------------------------------------------------------------------------
def nb_meta(nb):
    nb.metadata["kernelspec"] = {"display_name": "Python 3", "language": "python", "name": "python3"}
    nb.metadata["language_info"] = {"name": "python"}
    return nb


NOMBRES = {
    "M00": "antes_de_empezar", "M01": "primeros_pasos", "M02": "variables_y_tipos",
    "M03": "comparaciones_y_logica", "M04": "texto", "M05": "decisiones_if",
    "M06": "bucles_for_while", "M07": "listas_y_tuplas", "M08": "diccionarios_y_conjuntos",
    "M09": "funciones", "M10": "errores_y_modulos", "M11": "archivos_csv_json",
    "M12": "clases_y_objetos", "M13": "analisis_de_datos_pandas", "M14": "sig_y_proyecto_integrador",
}


def nombre_archivo(m):
    return f"{m['id']}_{NOMBRES.get(m['id'], slug(m['titulo']))}"


def cuaderno_modulo(m):
    celdas = []
    n_mod = num(m)
    objetivos = "\n".join(f"- {o}" for o in m["objetivos"])
    aviso = ""
    if not m["web"]:
        aviso = "\n> **Nota:** este módulo usa paquetes externos. Ejecútelo en Jupyter (Anaconda) o Google Colab.\n"
    celdas.append(new_markdown_cell(
        f"# Módulo {n_mod} · {m['titulo']}\n\n"
        f"**Curso:** {TITULO_CURSO} · **Duración estimada:** {m['duracion']}\n{aviso}\n"
        f"## Objetivos\n\nAl terminar este módulo usted podrá:\n\n{objetivos}\n\n"
        "## Cómo usar este cuaderno\n\n"
        "1. Lea cada lección y ejecute los ejemplos con **Shift + Enter**.\n"
        "2. Modifique los ejemplos: cambie valores y observe qué pasa. Así se aprende.\n"
        "3. Resuelva los ejercicios en la celda indicada y ejecute la **celda de comprobación**.\n"
        "4. Si se atasca, abra la **pista**. Las soluciones están en la carpeta `soluciones/`.\n"
        "5. Termine con la **autoevaluación**."))
    for i, lec in enumerate(m["lecciones"], 1):
        txt = f"## {i}. {lec['titulo']}\n\n{lec['texto']}"
        if lec.get("error"):
            txt += "\n\n> Esta celda **produce un error a propósito**. Después de leerlo, continúe con la siguiente celda."
        if lec.get("entrada"):
            datos = lec["entrada"].replace("\n", "`, `")
            txt += f"\n\n> Al ejecutar, Python le pedirá datos. Por ejemplo, escriba: `{datos}` (pulse Enter tras cada uno)."
        celdas.append(new_markdown_cell(txt))
        if "codigo" in lec:
            c = new_code_cell(lec["codigo"])
            if not lec.get("entrada"):
                c.outputs = lec["_salidas"]
            celdas.append(c)
    celdas.append(new_markdown_cell(
        "---\n# Ejercicios\n\n"
        "Escriba su solución en la celda de código de cada ejercicio, ejecútela y luego ejecute la "
        "**celda de comprobación** que está debajo. Si todo está bien verá `✔ ¡Correcto!`; si no, "
        "un mensaje le indicará qué revisar.\n\n"
        "Niveles: ★☆☆ básico · ★★☆ intermedio · ★★★ reto."))
    for j, ej in enumerate(m["ejercicios"], 1):
        celdas.append(new_markdown_cell(
            f"### Ejercicio {j} · {ej['titulo']}  {ESTRELLAS[ej['nivel']]}\n\n{ej['enunciado']}\n\n"
            f"<details><summary><b>Ver pista</b></summary>\n\n{ej['pista']}\n\n</details>"))
        celdas.append(new_code_cell(ej["plantilla"]))
        celdas.append(new_code_cell(
            "# Comprobación del ejercicio (no modifique esta celda)\n"
            + ej["prueba"] + "\nprint(\"✔ ¡Correcto! Ejercicio resuelto.\")"))
    preguntas = []
    letras = "abcd"
    for k, q in enumerate(m["quiz"], 1):
        ops = "\n".join(f"   {letras[x]}) {o}" for x, o in enumerate(q["opciones"]))
        preguntas.append(
            f"**{k}. {q['pregunta']}**\n\n{ops}\n\n"
            f"<details><summary>Ver respuesta</summary>\n\n"
            f"**{letras[q['correcta']]})** {q['opciones'][q['correcta']]}. {q['explicacion']}\n\n</details>")
    celdas.append(new_markdown_cell("---\n# Autoevaluación\n\nResponda mentalmente y luego despliegue la respuesta.\n\n" + "\n\n".join(preguntas)))
    filas = "\n".join(f"| **{t}** | {d} |" for t, d in m["glosario"].items())
    celdas.append(new_markdown_cell(f"# Glosario del módulo\n\n| Término | Significado |\n|---|---|\n{filas}"))
    lista = "\n".join(f"- [ ] {o}" for o in m["objetivos"])
    sig = ""
    idx = MODULOS.index(m)
    if idx + 1 < len(MODULOS):
        s = MODULOS[idx + 1]
        sig = f"\n\n**Siguiente:** `{nombre_archivo(s)}.ipynb` — {s['titulo']}."
    celdas.append(new_markdown_cell(f"# Antes de continuar\n\nMarque lo que ya domina:\n\n{lista}{sig}"))
    nb = new_notebook(cells=celdas)
    return nb_meta(nb)


def cuaderno_soluciones(m):
    celdas = [new_markdown_cell(
        f"# Soluciones · Módulo {num(m)} · {m['titulo']}\n\n"
        "Consulte las soluciones **después** de intentar cada ejercicio. Suele haber varias "
        "soluciones correctas: si la suya pasa la comprobación, es válida.")]
    for j, ej in enumerate(m["ejercicios"], 1):
        celdas.append(new_markdown_cell(f"## Ejercicio {j} · {ej['titulo']}\n\n{ej['enunciado']}"))
        c = new_code_cell(ej["solucion"])
        c.outputs = ej["_salidas_sol"]
        celdas.append(c)
        t = new_code_cell("# Comprobación\n" + ej["prueba"] + "\nprint(\"✔ ¡Correcto! Ejercicio resuelto.\")")
        t.outputs = [new_output("stream", name="stdout", text="✔ ¡Correcto! Ejercicio resuelto.\n")]
        celdas.append(t)
    return nb_meta(new_notebook(cells=celdas))


def escribir_cuadernos():
    d_c = os.path.join(RAIZ, "cuadernos")
    d_s = os.path.join(RAIZ, "soluciones")
    os.makedirs(d_c, exist_ok=True)
    os.makedirs(d_s, exist_ok=True)
    for m in MODULOS:
        nbformat.write(cuaderno_modulo(m), os.path.join(d_c, nombre_archivo(m) + ".ipynb"))
        nbformat.write(cuaderno_soluciones(m), os.path.join(d_s, nombre_archivo(m) + "_soluciones.ipynb"))


# ---------------------------------------------------------------------------
# 3. HTML interactivo
# ---------------------------------------------------------------------------
def datos_web():
    mods = []
    glosario = []
    for m in MODULOS:
        lecs = []
        for lec in m["lecciones"]:
            d = {"t": lec["titulo"], "h": md_html(lec["texto"]), "web": lec["web"]}
            if "codigo" in lec:
                d.update({
                    "c": lec["codigo"], "ch": resaltar(lec["codigo"]),
                    "o": lec.get("_texto", ""), "e": lec.get("_error"),
                    "img": lec.get("_imgs", []), "in": lec.get("entrada", ""),
                    "err": bool(lec.get("error")),
                })
            lecs.append(d)
        ejs = []
        for j, ej in enumerate(m["ejercicios"], 1):
            ejs.append({
                "id": f"{m['id']}-E{j}", "t": ej["titulo"], "n": ej["nivel"],
                "h": md_html(ej["enunciado"]), "p": ej["plantilla"], "pista": md_html(ej["pista"]),
                "s": ej["solucion"], "sh": resaltar(ej["solucion"]), "so": ej.get("_texto_sol", ""),
                "test": ej["prueba"], "web": ej["web"],
            })
        quiz = [{"q": q["pregunta"], "o": q["opciones"], "c": q["correcta"], "x": q["explicacion"]} for q in m["quiz"]]
        for t, d in m["glosario"].items():
            glosario.append({"t": t, "d": d, "m": m["id"]})
        mods.append({
            "id": m["id"], "n": num(m), "t": m["titulo"], "dur": m["duracion"], "obj": m["objetivos"],
            "web": m["web"], "lec": lecs, "ej": ejs, "quiz": quiz,
            "glos": [{"t": t, "d": d} for t, d in m["glosario"].items()],
            "nb": nombre_archivo(m) + ".ipynb",
        })
    glosario.sort(key=lambda g: unicodedata.normalize("NFKD", g["t"].lower()))
    return {"mods": mods, "glos": glosario}


def escribir_html():
    with open(os.path.join(AQUI, "plantilla.html"), encoding="utf-8") as f:
        plantilla = f.read()
    with open(os.path.join(AQUI, "runner_web.py"), encoding="utf-8") as f:
        runner = f.read()
    datos = json.dumps(datos_web(), ensure_ascii=False, separators=(",", ":"))
    datos = datos.replace("</", "<\\/")
    runner_js = json.dumps(runner).replace("</", "<\\/")
    html = plantilla.replace("/*__DATOS__*/null", datos).replace("/*__RUNNER__*/\"\"", runner_js)
    destino = os.path.join(RAIZ, "curso_python_interactivo.html")
    with open(destino, "w", encoding="utf-8") as f:
        f.write(html)
    return destino


def main():
    problemas = procesar()
    total_ej = sum(len(m["ejercicios"]) for m in MODULOS)
    total_lec = sum(len(m["lecciones"]) for m in MODULOS)
    total_q = sum(len(m["quiz"]) for m in MODULOS)
    print(f"Módulos: {len(MODULOS)} · Lecciones: {total_lec} · Ejercicios: {total_ej} · Preguntas: {total_q}")
    if problemas:
        print("\nPROBLEMAS ENCONTRADOS:")
        for p in problemas:
            print(" -", p)
        sys.exit(1)
    print("Todos los ejemplos se ejecutaron y todas las soluciones pasaron sus pruebas.")
    escribir_cuadernos()
    if os.path.exists(os.path.join(AQUI, "plantilla.html")):
        print("HTML:", escribir_html())


if __name__ == "__main__":
    main()
