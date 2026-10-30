"""Construye todos los entregables de los árboles de problemas desde `contenido_arboles.py`.

    python arboles_problemas/construir.py [carpeta_html]

Escribe en arboles_problemas/:
  arboles_IN_Piura.json                 modelo de datos (4 árboles)
  Arboles_Problemas_IN_Piura.pptx       láminas A3 editables (PowerPoint)
  Arboles_Problemas_IN_Piura.drawio     diagrama editable (diagrams.net)
  Arboles_Problemas_IN_Piura.html       página interactiva (la misma del artefacto)
"""
import html
import json
import sys
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import contenido_arboles as C  # noqa: E402
import generar_arboles as G  # noqa: E402

e = html.escape
PROVN = {"M": "Morropón", "H": "Huancabamba", "A": "Ayabaca"}


def chips(prov):
    return "".join(f'<span class="pv {k} d{v}" title="{PROVN[k]}">{k}{"·" * v}</span>' for k, v in prov.items() if v > 0) or "—"


def tabla_provincias():
    filas = [
        ("Bloques · distritos · microcuencas", "59 · 9 · 17", "48 · 5 · 11", "10 · 1 · 4", "117 · 15 · 32"),
        ("Superficie de bloques (ha)", "7,327.37", "3,551.16", "1,391.70", "12,270.24"),
        ("Superficie neta V8 (ha)", "7,313.55", "3,452.21", "1,391.70", "12,157.47"),
        ("Centros poblados · población 2017", "79 · 9,537", "58 · 6,573", "17 · 2,541", "154 · 18,651"),
        ("Altitud media (msnm) · pendiente media (%, DEM GLO-30)", "498 · 34.7", "1,317 · 49.0", "1,420 · 57.5", "840 · 41.4"),
        ("Peligro integrado alto + muy alto (ha · %)", "5,335.20 · 72.8", "2,353.73 · 66.3", "782.64 · 56.2", "8,471.57 · 69.0"),
        ("Índice de peligro integrado", "4.020", "3.898", "3.580", "3.935"),
        ("Peligro por movimientos en masa alto + muy alto en bloques (EVAR, %)", "90.7", "98.2", "99.4", "93.9"),
        ("Riesgo alto + muy alto en áreas de influencia (EVAR, %)", "28.9", "80.0", "90.9", "42.0"),
        ("Bajo el umbral MSAVI, criterio areal E6 (ha · %)", "6,216.63 · 84.8", "1,599.99 · 45.1", "525.21 · 37.7", "8,341.83 · 68.0"),
        ("MSAVI medio ponderado", "0.379", "0.508", "0.548", "0.435"),
        ("Bloques en estado moderado o severo (E6)", "14", "6", "0", "20"),
        ("Cárcavas registradas (n · km)", "39 · 18.02", "9 · 3.90", "7 · 1.15", "55 · 23.08"),
        ("Lluvia (mm/año) · índice de aridez", "543 · 0.31", "541 · 0.38", "557 · 0.39", "544 · 0.34"),
        ("Lluvia en años El Niño frente a la media", "×3.0–3.8 (hasta ×6.1)", "×2.0–2.7 (×3.3 en bosque seco)", "×2.7–4.9", "×2 a ×6"),
        ("Superficie sin registro predial (%)", "71.8", "71.0", "16.1 (83.9 % comunal)", "—"),
        ("Localidades con comité de recursos naturales", "2 de 56", "3 de 22", "0 de 7", "5 de 85"),
        ("Localidades con alerta temprana formal", "0", "0", "0", "0"),
    ]
    cab = "<tr><th>Variable</th><th>Morropón</th><th>Huancabamba</th><th>Ayabaca</th><th>Total</th></tr>"
    cuerpo = "".join(f"<tr><td>{e(a)}</td>" + "".join(f'<td class="num">{e(x)}</td>' for x in r) + "</tr>" for a, *r in filas)
    return f'<div class="tabla-scroll"><table><thead>{cab}</thead><tbody>{cuerpo}</tbody></table></div>'


def tabla_consistencia():
    u = C.unificado
    filas = []
    for cd in u["causas_directas"]:
        filas.append(f'<tr><td style="font-family:var(--mono)"><b>{e(cd["id"])}</b></td><td><b>{e(cd["texto"])}</b></td>'
                     f'<td>{chips(cd["prov"])}</td><td>{e(cd.get("adr", ""))}</td><td>{e(cd.get("linea", ""))}</td></tr>')
        for ci in [x for x in u["causas_indirectas"] if x["padre"] == cd["id"]]:
            filas.append(f'<tr><td style="font-family:var(--mono)">{e(ci["id"])}</td><td>{e(ci["texto"])}</td>'
                         f'<td>{chips(ci["prov"])}</td><td>{e(ci.get("adr", ""))}</td><td>{e(ci.get("linea", ""))}</td></tr>')
    cab = "<tr><th>Código</th><th>Causa</th><th>Provincias</th><th>Lectura AdR-CCC</th><th>Componente que la atiende</th></tr>"
    return f'<div class="tabla-scroll"><table><thead>{cab}</thead><tbody>{"".join(filas)}</tbody></table></div>'


def metodo_html():
    return f"""
<section>
  <h2>Cómo se construyeron los árboles</h2>
  <ol class="pasos">
    <li><div><b>Evidencia.</b> Se revisaron los informes del Entregable 2 por provincia (Vol. I y II del 28-oct, Vol. III del 30-oct-2026), las Matrices E2 B2 (24-oct), los diagnósticos V5 por provincia (117 informes territoriales, peligro integrado MCA-AHP, suelos E5, ecosistemas E6, hidroclima PISCOp 1981–2025), el Diagnóstico Social (374 fichas en 154 centros poblados; 82 de 117 bloques con ficha), el Informe EVAR del 19-oct y el Avance del Perfil del 28-oct.</div></li>
    <li><div><b>Lectura con la cadena del peligro (CENEPRED).</b> Lo que el proyecto puede modificar y aumenta el peligro (cobertura, suelo, drenaje) o la vulnerabilidad (organización, tenencia) es <i>causa</i>. La manifestación del peligro en la zona de origen (movimientos en masa, erosión, escorrentía torrencial) es <i>efecto directo</i>. Los daños a elementos expuestos son <i>efectos indirectos</i> y la consecuencia agregada es el <i>efecto final</i>. Pendiente, litología y lluvias extremas no son modificables: se muestran como factores exógenos a los lados del problema, como pide la Nota Técnica de meso y microlocalización.</div></li>
    <li><div><b>Filtro de tipología.</b> El servicio es la regulación de riesgos naturales (Formato 04-A, RM N° 00213-2024-MINAM). La regulación hídrica y la sequía, que la población menciona primero (82–89 % de localidades), corresponden a otra tipología con su propia brecha. Por eso salen del problema central y de los efectos directos y quedan solo como efecto colateral (EI), igual que en el modelo ANIN.</div></li>
    <li><div><b>Alineación con el modelo de referencia ANIN</b> (ZZ-000005, Fig. 2). Se mantiene su problema central y su efecto final. La causa «Alteración de los factores condicionantes en zonas de origen» se desagrega en CD1 (cobertura) y CD2 (suelos y laderas), y «Inadecuada gobernanza y gestión del ecosistema» es CD3. Así cada causa directa lleva a un componente del proyecto: infraestructura verde, infraestructura marrón y gobernanza (ZZ-000006, num. 10.1.7).</div></li>
    <li><div><b>Árbol unificado por convergencia.</b> Entra una causa o un efecto que tiene evidencia en al menos dos provincias o que es generalizado en una. Las marcas M, H y A muestran dónde aparece: · puntual, ·· frecuente, ··· generalizada. Lo propio de una sola provincia queda en su árbol provincial.</div></li>
    <li><div><b>Control de calidad.</b> Se verificó que el problema esté redactado como situación negativa y no como falta de una solución, que cada causa tenga evidencia y fuente, que no haya causas repetidas entre ramas, que la lógica «si ocurre la causa, entonces aparece el efecto» se cumpla en cada flecha y que toda cifra provenga de un documento del proyecto.</div></li>
  </ol>
</section>

<section>
  <h2>Comparación de las tres provincias</h2>
  <p class="nota">Cifras del Avance del Perfil (28-oct) y de las Matrices E2 B2; el peligro y riesgo por movimientos en masa se calcularon por provincia con los anexos del Informe EVAR. El peligro integrado (MCA-AHP: movimientos en masa, erosión y potencial generador de inundaciones) y el EVAR (solo movimientos en masa, método CENEPRED) no son comparables entre sí.</p>
  {tabla_provincias()}
</section>

<section>
  <h2>Matriz de consistencia del árbol unificado</h2>
  <p class="nota">Cada causa directa e indirecta, dónde se presenta, cómo se lee en el análisis de riesgo y qué componente del proyecto la atiende. Es la base para el árbol de medios y fines del capítulo de Identificación.</p>
  {tabla_consistencia()}
</section>

<section>
  <h2>Ajustes frente a la propuesta preliminar del E7 (corte 03-oct-2026)</h2>
  <ul>
    <li><b>Problema central.</b> Decía «servicios ecosistémicos de regulación de riesgos naturales y de regulación hídrica». Ahora se limita a la regulación de riesgos naturales y adopta la forma del modelo ANIN: la población no accede adecuadamente al servicio por la degradación de los ecosistemas en las zonas de origen.</li>
    <li><b>Efecto directo 3.</b> «Menor infiltración y regulación hídrica» pasa a «Mayor escorrentía torrencial y transporte de sedimentos», que es la dimensión de zona generadora de inundaciones del peligro integrado.</li>
    <li><b>Causa directa 2.</b> «Uso del suelo y de las vías que acelera la erosión» era en realidad un conjunto de causas. Ahora describe el estado del suelo y de las laderas, y las vías, el pisoteo y la falta de conservación pasan a ser sus causas indirectas.</li>
    <li><b>Pastoreo.</b> Su efecto principal es impedir la regeneración (CD1). El pisoteo y los senderos que concentran escorrentía quedan en CD2, porque son otro mecanismo con su propia medida.</li>
    <li><b>Gobernanza.</b> Se alinea con las dos causas del modelo ANIN (conocimiento y organización comunal; gobernanza de entidades públicas y privadas) y se agrega la tenencia sin sanear, que condiciona la libre disponibilidad.</li>
    <li><b>Efecto final.</b> Se adopta el del modelo ANIN (costos de reconstrucción y resiliencia), que conecta con la valoración de beneficios por costos evitados.</li>
    <li><b>Factores exógenos.</b> Pendiente, litología, El Niño y cambio climático se muestran a los lados del problema y fuera de la cadena causal, porque el proyecto no puede modificarlos.</li>
  </ul>
</section>

<section>
  <h2>Pendientes para cerrar los árboles</h2>
  <ul>
    <li><span class="pend">[POR DEFINIR]</span> Criterio oficial de la línea base de brecha. El Avance del 28-oct propone el criterio areal E6 (8,341.83 ha; 68.0 %), que es el que usan estos árboles. Las Matrices E2 B2 usan la media de bloque (7,805.23 ha; 64.2 %).</li>
    <li><span class="pend">[POR DEFINIR]</span> Peligro por componente (movimientos en masa, erosión y potencial generador de inundaciones) modelado por bloque. Hoy solo existe el peligro integrado y la calificación cualitativa de campo.</li>
    <li><span class="pend">[POR DEFINIR]</span> Inventario de daños por evento (El Niño 1983, 1998, 2017 y Yaku 2023) con SIGRID e INDECI, y valoración de pérdidas para el efecto final.</li>
    <li><span class="pend">[POR DEFINIR]</span> Delimitación de la población y los elementos expuestos aguas abajo (80 bloques «por delimitar») y cuantificación de sedimentos con el E4.</li>
    <li><span class="pend">[POR DEFINIR]</span> Laboratorio de suelos (materia orgánica, textura, pH, factor K), carga ganadera y frecuencia de quemas.</li>
    <li><span class="pend">[POR DEFINIR]</span> Pobreza, NBI y PEA; fichas F-DS-04, 05 y 07; F-DS-03 y 06 en Ayabaca; fichas sociales en Huarmaca (21 bloques con 3 fichas).</li>
    <li><span class="pend">[POR DEFINIR]</span> Escenarios de cambio climático en el EVAR (hoy usa un desencadenante fijo).</li>
  </ul>
</section>

<section>
  <h2>Marco normativo y metodológico</h2>
  <ul>
    <li>Sistema Nacional de Programación Multianual y Gestión de Inversiones (Invierte.pe): Directiva N° 001-2019-EF/63.01 y Guía General para la identificación, formulación y evaluación de proyectos (DGPMI-MEF), con su anexo de gestión del riesgo de desastres en contexto de cambio climático.</li>
    <li>RM N° 178-2019-MINAM, lineamientos para proyectos en ecosistemas; RM N° 00213-2024-MINAM, indicador de brecha y Formato 04-A (definición del servicio ecosistémico de regulación de riesgos naturales).</li>
    <li>ANIN: contenidos mínimos del E7 (ZZ-000005, Figuras 2 y 3), del capítulo de Identificación (ZZ-000006) y de Formulación (ZZ-000007); Nota Técnica de meso y microlocalización de infraestructura natural.</li>
    <li>CENEPRED, Manual para la evaluación de riesgos originados por fenómenos naturales (versión 2): peligro como combinación de factores condicionantes y desencadenantes; vulnerabilidad como exposición, fragilidad y resiliencia.</li>
    <li>Metodología complementaria: árbol de problemas del Marco Lógico (CEPAL-ILPES), enfoque de servicios ecosistémicos y triangulación de evidencia territorial, social y de modelamiento.</li>
  </ul>
  <p class="nota">Las citas de normas y guías se tomaron de los documentos del proyecto en Drive. La Guía General DGPMI no está en la carpeta del proyecto: conviene confirmar sus numerales antes de citarlos en el informe.</p>
</section>

<section>
  <h2>Archivos editables</h2>
  <p>Los cuatro árboles están en versión editable en el repositorio del proyecto, carpeta <code>arboles_problemas/</code>:</p>
  <ul>
    <li><b>Arboles_Problemas_IN_Piura.pptx</b>: PowerPoint en lámina A3 horizontal, un árbol por lámina. Cajas y conectores nativos: al mover una caja, sus flechas la siguen. El sustento de cada árbol está en las notas de la lámina.</li>
    <li><b>Arboles_Problemas_IN_Piura.drawio</b>: diagrama para diagrams.net (gratuito, en línea o de escritorio), una página por árbol. También se importa en Visio y Lucidchart.</li>
    <li><b>arboles_IN_Piura.json</b> y <b>contenido_arboles.py</b>: el contenido con su evidencia; <b>construir.py</b> regenera todos los formatos si se cambia un texto.</li>
  </ul>
</section>
"""


def datos():
    meta = {
        "titulo": "Árboles de problemas, causas y efectos – Proyecto IN Piura",
        "titulo_pagina": "Árboles de problemas, causas y efectos",
        "subtitulo": "Proyecto IN Piura · CUI 2669244 · Cuenca Alta del Río Piura · Estudio de preinversión a nivel de perfil",
        "version": "Versión 1.0 · 30 de octubre de 2026 · Insumos: E2 Vol. I–III (28 y 30-oct), Matrices E2 B2 (24-oct), EVAR (19-oct), Avance del Perfil (28-oct)",
        "pie": "ANIN - DIME - SESDI · Proyecto IN Piura · CUI 2669244 · Versión 1.0, 30-oct-2026",
    }
    return {"meta": meta, "orden": C.ORDEN, "arboles": C.ARBOLES, "metodo_html": metodo_html()}


def validar(d):
    for k, t in d["arboles"].items():
        ids = [t["efecto_final"]["id"], t["problema"]["id"]] + [x["id"] for x in t["efectos_indirectos"] + t["efectos_directos"]
                                                                  + t["causas_directas"] + t["causas_indirectas"] + t["factores_exogenos"]]
        assert len(ids) == len(set(ids)), f"{k}: ids repetidos"
        cds = {c["id"] for c in t["causas_directas"]}
        eds = {x["id"] for x in t["efectos_directos"]}
        for ci in t["causas_indirectas"]:
            assert ci["padre"] in cds, f"{k}: {ci['id']} sin padre"
        for ei in t["efectos_indirectos"]:
            assert set(ei["padres"]) <= eds, f"{k}: {ei['id']} con padre inexistente"
        for ed in eds:
            assert any(ed in ei["padres"] for ei in t["efectos_indirectos"]), f"{k}: {ed} no conduce a ningún efecto indirecto"
        for cd in cds:
            assert any(ci["padre"] == cd for ci in t["causas_indirectas"]), f"{k}: {cd} sin causas indirectas"
        texto = json.dumps(t, ensure_ascii=False).lower()
        assert "banco mundial" not in texto


def main():
    d = datos()
    validar(d)
    (AQUI / "arboles_IN_Piura.json").write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    plantilla = (AQUI / "plantilla_artefacto.html").read_text(encoding="utf-8")
    pagina = plantilla.replace("/*__DATA__*/null", json.dumps(d, ensure_ascii=False).replace("</", "<\\/"))
    # Documento completo: se abre tal cual en un navegador y es lo que se publica como artefacto
    corte = pagina.index("</style>") + len("</style>")
    reinicio = ("<style>:root{color-scheme:light;padding-top:env(safe-area-inset-top,0px);"
                "padding-bottom:env(safe-area-inset-bottom,0px)}body{margin:0}img{max-width:100%}"
                "[hidden]{display:none!important}</style>\n")
    pagina = ('<!DOCTYPE html>\n<html lang="es">\n<head>\n<meta charset="utf-8">\n'
              '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
              + reinicio + pagina[:corte] + "\n</head>\n<body>\n" + pagina[corte:] + "\n</body>\n</html>\n")
    destinos = [AQUI / "Arboles_Problemas_IN_Piura.html"] + [Path(p) / "arboles_problemas_in_piura.html" for p in sys.argv[1:]]
    for ruta in destinos:
        ruta.write_text(pagina, encoding="utf-8")
    G.generar_pptx(d, AQUI / "Arboles_Problemas_IN_Piura.pptx")
    G.generar_drawio(d, AQUI / "Arboles_Problemas_IN_Piura.drawio")
    print("OK:", ", ".join(str(p) for p in destinos), "+ pptx, drawio, json")


if __name__ == "__main__":
    main()
