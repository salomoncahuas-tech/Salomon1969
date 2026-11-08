# -*- coding: utf-8 -*-
"""Artefacto HTML del Plan de Contingencia: inyecta los datos maestros en la plantilla."""
import json, sys, os
sys.path.insert(0, os.path.dirname(__file__))
from datos_plan import DATOS
OUT = sys.argv[1] if len(sys.argv) > 1 else "Plan_Contingencia_IN_Piura.html"
data = json.dumps(DATOS, ensure_ascii=False).replace("</", "<\\/")
html = open(os.path.join(os.path.dirname(__file__), "plantilla.html"), encoding="utf-8").read().replace("/*__DATOS__*/", data)
open(OUT, "w", encoding="utf-8").write(html); print("ok", OUT, len(html)//1024, "KB")
