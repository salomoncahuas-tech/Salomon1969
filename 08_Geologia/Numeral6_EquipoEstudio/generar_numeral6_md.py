# -*- coding: utf-8 -*-
"""Genera la version Markdown del numeral 6. EQUIPO DE ESTUDIO.

Entregable 3 - Estudio de Geologia. Proyecto IN Piura (CUI 2669244) - ANIN / DIME / SESDI.
Toma los mismos datos que el DOCX (contenido6.py), de modo que ambos formatos no
puedan divergir.

Uso:
    python3 generar_numeral6_md.py [salida.md]
"""
import sys

from contenido6 import (ENCABEZADO_PROF, ENCABEZADO_ROL, FILAS, LEYENDA_FIGURA,
                        LEYENDA_TABLA, NOTAS, PARRAFOS_61, TEXTO_62)

CABECERA = """# AUTORIDAD NACIONAL DE INFRAESTRUCTURA – ANIN
## DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA – DIME
## SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN – SESDI

*Proyecto IN Piura (CUI 2669244) · Entregable 3 – Estudio de Geología ·
Modalidad de ejecución: Administración Directa*
"""


def celda(txt):
    """Escapa el separador de columnas y colapsa los saltos de linea."""
    return txt.replace("|", "\\|").replace("\n", " – ")


def main():
    salida = sys.argv[1] if len(sys.argv) > 1 else "E3_Geologia_6_Equipo_de_Estudio.md"
    n = len(ENCABEZADO_PROF)
    out = [CABECERA, "\n## 6. EQUIPO DE ESTUDIO\n",
           "### 6.1. Cuadro de responsables por actividad\n"]
    out += [p + "\n" for p in PARRAFOS_61]
    out.append(f"\n**{LEYENDA_TABLA}**\n")

    out.append("| ACTIVIDADES Y TAREAS DEL ENTREGABLE | " +
               " | ".join(celda(p) for p in ENCABEZADO_PROF) + " |")
    out.append("|" + " --- |" * (n + 1))
    out.append("| *Rol en el entregable, según el flujograma del Entregable 3 →* | " +
               " | ".join(f"*{celda(r)}*" for r in ENCABEZADO_ROL) + " |")

    for etiqueta, marcas in FILAS:
        if marcas is None:
            out.append(f"| **{celda(etiqueta)}** |" + " |" * n)
            continue
        marcadas = " | ".join("**X**" if m == "X" else "" for m in marcas)
        out.append(f"| {celda(etiqueta)} | {marcadas} |")

    out.append("")
    out += [f"*{nota}*\n" for nota in NOTAS]
    out.append("\n### 6.2. Organigrama del equipo de estudio\n")
    out.append(TEXTO_62 + "\n")
    out.append("![Organigrama del equipo de estudio](Organigrama_E3_Geologia.png)\n")
    out.append(f"**{LEYENDA_FIGURA}**")

    texto = "\n".join(out) + "\n"
    with open(salida, "w", encoding="utf-8") as fh:
        fh.write(texto)
    print(f"OK -> {salida} ({len(texto.encode('utf-8')):,} bytes)")


if __name__ == "__main__":
    main()
