# Ejecutor de Python (Brython) que corre dentro de un Web Worker del curso interactivo.
import sys
import json
import traceback
from browser import self as _worker


class _Salida:
    def __init__(self):
        self.buf = []

    def write(self, s):
        self.buf.append(str(s))

    def flush(self):
        pass


def _formatear_error(e):
    nombre = type(e).__name__
    linea = getattr(e, "lineno", None)
    if linea is None:
        try:
            for fr in traceback.extract_tb(e.__traceback__):
                if fr.filename == "<celda>":
                    linea = fr.lineno
        except Exception:
            pass
    msg = str(e)
    if isinstance(e, SyntaxError) and getattr(e, "msg", None):
        msg = e.msg
    prefijo = f"Línea {linea}: " if linea else ""
    return f"{prefijo}{nombre}: {msg}"


def py_run(src, entrada, prueba):
    salida = _Salida()
    anterior = (sys.stdout, sys.stderr)
    sys.stdout = salida
    sys.stderr = salida
    lineas = entrada.split("\n") if entrada else []

    def _input(mensaje=""):
        salida.write(str(mensaje))
        if not lineas:
            raise EOFError("no hay más datos en el cuadro Entrada")
        valor = lineas.pop(0)
        salida.write(valor + "\n")
        return valor

    ns = {"__name__": "__main__", "input": _input}
    res = {"ok": True, "error": None, "test": None, "test_msg": None}
    try:
        exec(compile(src, "<celda>", "exec"), ns)
    except BaseException as e:
        res["ok"] = False
        res["error"] = _formatear_error(e)
    if res["ok"] and prueba:
        try:
            exec(compile(prueba, "<prueba>", "exec"), ns)
            res["test"] = True
        except AssertionError as e:
            res["test"] = False
            res["test_msg"] = str(e) or "Una comprobación no se cumplió."
        except BaseException as e:
            res["test"] = False
            res["test_msg"] = _formatear_error(e)
    sys.stdout, sys.stderr = anterior
    res["out"] = "".join(salida.buf)
    return json.dumps(res)


_worker.pyRun = py_run
