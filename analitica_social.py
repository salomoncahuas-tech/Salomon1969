"""
Analitica del Diagnostico Social (F-DS-01..F-DS-07) - Proyecto IN Piura.

Genera, a partir de las fichas sociales registradas en el aplicativo, los
mismos tres productos que ya existen para el Diagnostico Territorial:

  1. Series listas para graficos interactivos (Altair) en el aplicativo.
  2. Un libro Excel con hojas de datos y graficos nativos de Excel,
     editables por el usuario, con encabezado institucional ANIN.
  3. Un anexo grafico en PDF que acompana a la ficha PDF del bloque.

El modulo NO depende de Streamlit: recibe los registros tal como los
devuelve `database.obtener_diagnosticos_sociales_por_bloque()` y devuelve
estructuras de datos puras, de modo que puede probarse sin levantar la
aplicacion.

Fuente de los datos: columna `dsNN_data_v3` de cada registro (JSON del
formulario V4) mas las columnas de cabecera (bloque, centro poblado,
distrito, responsable, fecha).

ANIN - DIME - SESDI | CUI 2669244 | UTM WGS 84 Zona 17S (EPSG:32717).
"""

import io
import json
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

import fds_listas as FL

# El motor de series -paleta institucional, modelo de serie y los tres
# renderizadores (Altair, Excel nativo y PDF vectorial)- es comun a este
# modulo y al del Diagnostico Territorial. Aqui solo viven la lectura de las
# fichas sociales y los constructores de sus secciones.
from analitica_series import (
    DIVERGENTE, RAMPA_CRITICA, RAMPA_FAVORABLE, RAMPA_NEUTRA, SIN_DATO,
    TINTAS, _apiladas, _clave, _colores_sino, _conteo, _escribir_tabla,
    _hoja_seccion, _hoja_tablas, _lista, _num, _pdf_barras_apiladas,
    _pdf_barras_simples, _rampa, _serie, _sino, _titulo_hoja, _txt,
)


# ══════════════════════════════════════════════════════════════════════════
# LECTURA DE LAS FICHAS
# ══════════════════════════════════════════════════════════════════════════

FICHAS_DS = ["F-DS-01", "F-DS-02", "F-DS-03", "F-DS-04",
             "F-DS-05", "F-DS-06", "F-DS-07"]

FICHAS_DS_TITULOS = {
    "F-DS-01": "Diagnostico Socioeconomico del CP / Comunidad",
    "F-DS-02": "Mapeo y Analisis de Actores Clave",
    "F-DS-03": "Entrevista Semiestructurada a Autoridades / Lideres",
    "F-DS-04": "Acta de Taller Participativo",
    "F-DS-05": "Conflictos Socioambientales y Oportunidades",
    "F-DS-06": "Percepcion Local de Peligros y Cambio Climatico",
    "F-DS-07": "Disposicion a Participar y Consentimiento Previo Informado",
}

def formulario(registro):
    """Dict del formulario V4 guardado en `dsNN_data_v3`."""
    ficha = registro.get("ficha", "") or ""
    num = ficha.split("-")[-1] if ficha else ""
    crudo = registro.get(f"ds{num}_data_v3", "") or ""
    if isinstance(crudo, dict):
        return crudo
    if not crudo:
        return {}
    try:
        valor = json.loads(crudo)
    except (json.JSONDecodeError, TypeError):
        return {}
    return valor if isinstance(valor, dict) else {}


def _tabla(form, slot):
    """Filas de un slot de tabla (st.data_editor) del formulario."""
    filas = form.get(slot, [])
    if not isinstance(filas, list):
        return []
    return [f for f in filas if isinstance(f, dict)]


def _col(fila, *nombres):
    """Valor de la primera columna presente (los encabezados cambian de
    plantilla a plantilla; se aceptan alias)."""
    for nombre in nombres:
        if nombre in fila:
            return _txt(fila[nombre])
    # Busqueda laxa: sin tildes, sin puntuacion y en minusculas.
    claves = {_clave(k): v for k, v in fila.items()}
    for nombre in nombres:
        valor = claves.get(_clave(nombre))
        if valor is not None:
            return _txt(valor)
    return ""


def deduplicar(registros):
    """Conserva el registro mas reciente por (ficha, centro poblado, responsable).

    `obtener_*_por_bloque` devuelve primero el mas reciente, de modo que las
    reediciones antiguas no inflan los conteos de los graficos.
    """
    vistos, salida = set(), []
    for reg in registros or []:
        clave = (reg.get("ficha", ""), (reg.get("centro_poblado", "") or "").strip().lower(),
                 reg.get("evaluador", ""))
        if clave not in vistos:
            vistos.add(clave)
            salida.append(reg)
    return salida


def _por_ficha(registros, ficha):
    return [r for r in registros if (r.get("ficha", "") or "") == ficha]


def _ambito(registro):
    """Etiqueta del ambito de un registro: centro poblado o, a falta de el,
    la comunidad campesina o el bloque."""
    return (_txt(registro.get("centro_poblado"))
            or _txt(registro.get("comunidad_campesina"))
            or _txt(registro.get("bloque_codigo"))
            or "Sin ambito consignado")


def _bateria_sino(registros, campos, orden=("Sí", "No", "No aplica", "Parcial")):
    """Barras apiladas de una bateria de preguntas Si/No.

    `campos` es una lista de (clave, etiqueta). Devuelve filas {cat, sub, valor}
    donde `cat` es la pregunta y `sub` la respuesta.
    """
    filas = []
    for clave, etiqueta in campos:
        conteo = {}
        for reg in registros:
            respuesta = _sino(formulario(reg).get(clave))
            if respuesta:
                conteo[respuesta] = conteo.get(respuesta, 0) + 1
        for sub in list(orden) + sorted(k for k in conteo if k not in orden):
            if conteo.get(sub):
                filas.append({"cat": etiqueta, "sub": sub, "valor": conteo[sub]})
    return filas


# ── F-DS-01: socioeconomico ───────────────────────────────────────────────

_ORDEN_ABC = ["Alto", "Medio", "Bajo"]
_ORDEN_BRM = ["Bueno", "Regular", "Malo"]


def _seccion_socioeconomica(registros, tema="claro"):
    regs = _por_ficha(registros, "F-DS-01")
    if not regs:
        return None
    series, tablas = [], []

    # 1. Poblacion por ambito, desagregada por sexo.
    filas_pob = []
    detalle_pob = []
    for reg in regs:
        f = formulario(reg)
        ambito = _ambito(reg)
        hombres, mujeres = _num(f.get("f1_pob_h")), _num(f.get("f1_pob_m"))
        total = _num(f.get("f1_pob_t"))
        if hombres is not None:
            filas_pob.append({"cat": ambito, "sub": "Hombres", "valor": hombres})
        if mujeres is not None:
            filas_pob.append({"cat": ambito, "sub": "Mujeres", "valor": mujeres})
        # Solo se grafica el total cuando no hay desagregacion: un total
        # apilado junto a sus partes duplicaria la poblacion.
        if hombres is None and mujeres is None and total is not None:
            filas_pob.append({"cat": ambito, "sub": "Población total", "valor": total})
        detalle_pob.append({
            "Centro poblado / ámbito": ambito,
            "Familias / viviendas": _num(f.get("f1_nfam")),
            "Población total (hab.)": total,
            "Hombres": hombres, "Mujeres": mujeres,
            "Menores de 18 años": _num(f.get("f1_pob_men18")),
            "Mayores de 65 años": _num(f.get("f1_pob_may65")),
            "Población originaria": _num(f.get("f1_pob_orig")),
            "Idioma predominante": _txt(f.get("f1_idioma")),
            "Nivel educativo predominante": _txt(f.get("f1_nivel_edu")),
        })
    if filas_pob:
        orden_sub = ["Hombres", "Mujeres", "Población total"]
        presentes, colores = _apiladas(filas_pob, orden_sub, categorica=True, tema=tema)
        series.append(_serie(
            "f1_poblacion", "Población por centro poblado, según sexo",
            "apiladas", filas_pob, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="hab.",
            eje_x="Habitantes", escala="categorica",
            descripcion="Población declarada en la ficha F-DS-01 para cada "
                        "centro poblado o comunidad del bloque.",
            nota="F-DS-01, numeral 2 (Datos demográficos). La desagregación "
                 "por sexo se grafica solo donde la ficha la consigna."))
    if detalle_pob:
        tablas.append(("Demografía por centro poblado", detalle_pob))

    # 2. Estructura etaria declarada.
    filas_edad = []
    for reg in regs:
        f = formulario(reg)
        total = _num(f.get("f1_pob_t"))
        men, may = _num(f.get("f1_pob_men18")), _num(f.get("f1_pob_may65"))
        if total is None or (men is None and may is None):
            continue
        ambito = _ambito(reg)
        men, may = men or 0, may or 0
        intermedia = max(total - men - may, 0)
        filas_edad += [
            {"cat": ambito, "sub": "Menores de 18 años", "valor": men},
            {"cat": ambito, "sub": "Entre 18 y 65 años", "valor": intermedia},
            {"cat": ambito, "sub": "Mayores de 65 años", "valor": may},
        ]
    if filas_edad:
        orden_sub = ["Menores de 18 años", "Entre 18 y 65 años", "Mayores de 65 años"]
        presentes, colores = _apiladas(filas_edad, orden_sub, rampa=RAMPA_NEUTRA)
        series.append(_serie(
            "f1_etaria", "Estructura etaria de la población",
            "apiladas", filas_edad, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="hab.",
            eje_x="Habitantes",
            descripcion="Población dependiente (menores de 18 y mayores de 65) "
                        "frente a la población en edad de trabajar.",
            nota="El tramo intermedio se obtiene por diferencia con la "
                 "población total declarada; no se estima donde la ficha no "
                 "consigna total."))

    # 3. Cobertura de servicios basicos.
    filas_cob = []
    for reg in regs:
        f = formulario(reg)
        ambito = _ambito(reg)
        for clave, etiqueta in (("f1_agua_cob", "Agua para consumo"),
                                ("f1_energia_cob", "Energía eléctrica")):
            valor = _num(f.get(clave))
            if valor is not None:
                filas_cob.append({"cat": ambito, "sub": etiqueta,
                                  "valor": min(valor, 100.0)})
    if filas_cob:
        presentes, colores = _apiladas(
            filas_cob, ["Agua para consumo", "Energía eléctrica"],
            categorica=True, tema=tema)
        series.append(_serie(
            "f1_cobertura", "Cobertura de agua y energía eléctrica (%)",
            "agrupadas", filas_cob, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="%",
            eje_x="Cobertura (%)", maximo=100,
            descripcion="Porcentaje de viviendas con acceso declarado en la "
                        "ficha, por centro poblado.",
            nota="F-DS-01, numeral 5 (Servicios básicos e infraestructura social)."))

    # 4. Fuentes de agua, saneamiento y energia (multiseleccion).
    for clave, titulo, lista, id_ in (
            ("f1_agua", "Fuentes de agua para consumo", FL.L_AGUA, "f1_fuentes_agua"),
            ("f1_sanea", "Tipos de saneamiento", FL.L_SANEA, "f1_saneamiento"),
            ("f1_energia", "Fuentes de energía", FL.L_ENERG, "f1_fuentes_energia")):
        valores = []
        for reg in regs:
            valores += _lista(formulario(reg).get(clave))
        filas = _conteo(valores, orden=lista)
        if filas:
            series.append(_serie(
                id_, f"{titulo} (centros poblados que la reportan)",
                "barras_h", filas, "clase", "valor", escala="neutra",
                unidad="CP", eje_x="Centros poblados",
                descripcion="Cada centro poblado puede reportar más de una "
                            "opción, por lo que el total supera el número de "
                            "centros poblados.",
                nota="F-DS-01, numeral 5. Marcado múltiple en la ficha."))

    # 5. Actividades economicas.
    filas_act, detalle_act = [], []
    for reg in regs:
        for fila in _tabla(formulario(reg), "f1_activ"):
            actividad = _col(fila, "Actividad / Rubro", "Actividad")
            if not actividad:
                continue
            familias = _num(_col(fila, "N fam.", "N familias", "Nfam"))
            destino = _col(fila, "Destino")
            filas_act.append({"cat": actividad, "sub": destino or "Sin destino consignado",
                              "valor": familias or 0})
            detalle_act.append({
                "Centro poblado / ámbito": _ambito(reg),
                "Actividad / Rubro": actividad,
                "N.° de familias": familias,
                "Productos principales": _col(fila, "Productos principales"),
                "Destino de la producción": destino,
                "Ingreso (S/ / mes)": _num(_col(fila, "Ingreso (S/./mes)", "Ingreso")),
            })
    filas_act = [f for f in filas_act if f["valor"] > 0]
    if filas_act:
        presentes, colores = _apiladas(filas_act, FL.L_DESTINO, rampa=RAMPA_NEUTRA)
        series.append(_serie(
            "f1_actividades", "Familias por actividad económica y destino de la producción",
            "apiladas", filas_act, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="familias",
            eje_x="Familias dedicadas",
            descripcion="El color ordena el destino de la producción, del "
                        "autoconsumo (claro) al mercado (oscuro): es el "
                        "indicador de articulación de los medios de vida.",
            nota="F-DS-01, numeral 6 (Actividades económicas y medios de vida)."))
    if detalle_act:
        tablas.append(("Actividades económicas", detalle_act))

    # 6. Programas sociales.
    programas = [("f1_juntos", "JUNTOS (familias)"),
                 ("f1_pension65", "Pensión 65 (personas)"),
                 ("f1_beca18", "Beca 18 (personas)"),
                 ("f1_qaliwarma", "Qali Warma (IIEE)")]
    filas_prog = []
    for clave, etiqueta in programas:
        total = sum(_num(formulario(r).get(clave)) or 0 for r in regs)
        if total:
            filas_prog.append({"clase": etiqueta, "valor": total})
    if filas_prog:
        series.append(_serie(
            "f1_programas", "Cobertura de programas sociales en el ámbito del bloque",
            "barras_h", filas_prog, "clase", "valor", escala="neutra",
            eje_x="Beneficiarios declarados",
            descripcion="Suma de los beneficiarios declarados por los centros "
                        "poblados del bloque.",
            nota="F-DS-01, numeral 7. Las unidades difieren por programa "
                 "(familias, personas o instituciones educativas)."))

    # 7. Gobernanza comunal y presencia institucional.
    campos_gob = [
        ("f1_junta_vig", "Junta Directiva vigente"),
        ("f1_ronda", "Ronda Campesina activa"),
        ("f1_reglamento", "Reglamento interno"),
        ("f1_comite_rrnn", "Comité de Recursos Naturales"),
        ("f1_pdc", "Plan de Desarrollo Concertado vigente"),
        ("f1_ongs", "ONG operando en la zona"),
        ("f1_agrorural", "Proyectos AGRORURAL"),
        ("f1_prodern", "PRODERN / FONCODES"),
    ]
    filas_gob = _bateria_sino(regs, campos_gob)
    if filas_gob:
        series.append(_serie(
            "f1_gobernanza", "Capacidades de gobernanza comunal e institucional",
            "apiladas", filas_gob, "cat", "valor", sub="sub",
            orden_sub=["Sí", "No", "No aplica"], colores=_colores_sino(tema),
            unidad="CP", eje_x="Centros poblados", escala="sino",
            descripcion="Presencia efectiva de cada capacidad, contada sobre "
                        "los centros poblados que respondieron.",
            nota="F-DS-01, numerales 3 y 7. Insumo directo de la línea de "
                 "gobernanza y gestión comunitaria del proyecto."))

    # 8. Tenencia de la tierra.
    tenencias = [_txt(formulario(r).get("f1_tenencia")) for r in regs]
    filas_ten = _conteo(tenencias, orden=FL.L_TENENCIA)
    if filas_ten:
        series.append(_serie(
            "f1_tenencia", "Régimen predominante de tenencia de la tierra",
            "barras_h", filas_ten, "clase", "valor", escala="neutra",
            unidad="CP", eje_x="Centros poblados",
            descripcion="Determina con quién se suscriben los acuerdos de "
                        "intervención en cada bloque.",
            nota="F-DS-01, numeral 4 (Tenencia de la tierra)."))

    filas_tit = []
    for reg in regs:
        pct = _num(formulario(reg).get("f1_pct_tituladas"))
        if pct is not None:
            filas_tit.append({"clase": _ambito(reg), "valor": min(pct, 100.0)})
    if filas_tit:
        series.append(_serie(
            "f1_tituladas", "Tierras tituladas por centro poblado (%)",
            "barras_h", filas_tit, "clase", "valor", escala="favorable",
            unidad="%", eje_x="Tierras tituladas (%)", maximo=100, decimales=1,
            descripcion="A mayor porcentaje titulado, menor riesgo de "
                        "observaciones prediales en el tamizaje del bloque.",
            nota="F-DS-01, numeral 4.1."))

    # 9. Percepciones ordinales del ambito.
    filas_perc = []
    etiquetas_perc = [("f1_presencia_estatal", "Percepción de presencia estatal"),
                      ("f1_migracion", "Tasa de migración juvenil")]
    for clave, etiqueta in etiquetas_perc:
        for fila in _conteo([_txt(formulario(r).get(clave)) for r in regs],
                            orden=_ORDEN_ABC):
            filas_perc.append({"cat": etiqueta, "sub": fila["clase"],
                               "valor": fila["valor"]})
    if filas_perc:
        presentes, colores = _apiladas(filas_perc, _ORDEN_ABC,
                                       rampa=RAMPA_NEUTRA, descendente=True)
        series.append(_serie(
            "f1_percepciones", "Percepciones declaradas sobre el ámbito",
            "apiladas", filas_perc, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="CP",
            eje_x="Centros poblados",
            descripcion="Escala Alto / Medio / Bajo de la ficha.",
            nota="F-DS-01, numerales 2.5 y 7.1."))

    return {"id": "F-DS-01", "titulo": "F-DS-01 · Diagnóstico socioeconómico",
            "descripcion": f"{len(regs)} ficha(s) registrada(s).",
            "series": series, "tablas": tablas}


# ── F-DS-02: actores clave ────────────────────────────────────────────────

# Escala ordenada de posicion frente al proyecto: se colorea con la paleta
# divergente, con "Neutral" como punto muerto.
_ORDEN_POSICION = ["A favor del proyecto", "Neutral / Sin posición definida",
                   "Reticente (requiere persuasión)", "En contra del proyecto",
                   "No determinada"]
_COLOR_POSICION = {
    "A favor del proyecto": DIVERGENTE[0],
    "Neutral / Sin posición definida": DIVERGENTE[2],
    "Reticente (requiere persuasión)": DIVERGENTE[3],
    "En contra del proyecto": DIVERGENTE[4],
    "No determinada": SIN_DATO,
}


def _seccion_actores(registros, tema="claro"):
    regs = _por_ficha(registros, "F-DS-02")
    if not regs:
        return None
    actores, detalle = [], []
    for reg in regs:
        for fila in _tabla(formulario(reg), "f2_actores"):
            nombre = _col(fila, "Nombre del actor / Organizacion",
                          "Nombre del actor / Organización", "Nombre del actor")
            tipo = _col(fila, "Tipo")
            if not (nombre or tipo):
                continue
            actor = {
                "ambito": _ambito(reg), "nombre": nombre, "tipo": tipo,
                "influencia": _col(fila, "Influencia"),
                "interes": _col(fila, "Interes", "Interés"),
                "posicion": _col(fila, "Posicion", "Posición"),
                "nivel": _col(fila, "Nivel territorial"),
                "rol": _col(fila, "Rol / Funcion frente al proyecto",
                            "Rol / Función frente al proyecto", "Rol"),
            }
            actores.append(actor)
            detalle.append({
                "Centro poblado / ámbito": actor["ambito"],
                "Actor / Organización": actor["nombre"], "Tipo": actor["tipo"],
                "Rol frente al proyecto": actor["rol"],
                "Influencia": actor["influencia"], "Interés": actor["interes"],
                "Posición": actor["posicion"], "Nivel territorial": actor["nivel"],
            })
    if not actores:
        return None

    series = []

    # 1. Actores por tipo.
    filas_tipo = _conteo([a["tipo"] for a in actores], orden=FL.L_TIPO_ACTOR)
    if filas_tipo:
        series.append(_serie(
            "f2_tipos", "Actores identificados por tipo de organización",
            "barras_h", filas_tipo, "clase", "valor", escala="neutra",
            unidad="actores", eje_x="N.° de actores",
            descripcion="Composición del mapa de actores del ámbito del bloque.",
            nota="F-DS-02, numeral 3 (Registro de actores identificados)."))

    # 2. Posicion frente al proyecto (escala ordenada, barra divergente).
    filas_pos = _conteo([a["posicion"] for a in actores], orden=_ORDEN_POSICION)
    if filas_pos:
        series.append(_serie(
            "f2_posicion", "Posición de los actores frente al proyecto",
            "barras_h", filas_pos, "clase", "valor",
            colores=_COLOR_POSICION, escala="posicion",
            unidad="actores", eje_x="N.° de actores",
            descripcion="Del verde (a favor) al rojo (en contra); el gris "
                        "reúne a quienes no declararon posición.",
            nota="F-DS-02. Define la estrategia de relacionamiento y el "
                 "orden de acercamiento a cada bloque."))

    # 3. Matriz influencia x interes.
    celdas = {}
    nombres = {}
    for a in actores:
        inf, ints = a["influencia"], a["interes"]
        if not (inf and ints):
            continue
        clave = (inf, ints)
        celdas[clave] = celdas.get(clave, 0) + 1
        nombres.setdefault(clave, []).append(a["nombre"] or "(sin nombre)")
    if celdas:
        filas_matriz = [{"cat": inf, "sub": ints, "valor": n,
                         "detalle": ", ".join(nombres[(inf, ints)][:8])}
                        for (inf, ints), n in celdas.items()]
        series.append(_serie(
            "f2_matriz", "Matriz de influencia × interés de los actores",
            "mapa_calor", filas_matriz, "cat", "valor", sub="sub",
            orden_cat=_ORDEN_ABC, orden_sub=_ORDEN_ABC, escala="neutra",
            unidad="actores", eje_x="Interés en el proyecto",
            eje_y="Influencia sobre el territorio",
            descripcion="El cuadrante influencia alta / interés alto concentra "
                        "a los actores con los que debe negociarse primero; "
                        "influencia alta / interés bajo, a quienes hay que "
                        "informar para que no se conviertan en oposición.",
            nota="F-DS-02. Clasificación clásica de mapeo de actores."))

    # 4. Nivel territorial.
    filas_niv = _conteo([a["nivel"] for a in actores], orden=FL.L_NIV_TERR)
    if filas_niv:
        series.append(_serie(
            "f2_nivel", "Actores por nivel territorial de actuación",
            "barras_h", filas_niv, "clase", "valor", escala="neutra",
            unidad="actores", eje_x="N.° de actores",
            descripcion="Permite ver si la gobernanza del bloque descansa en "
                        "actores comunales o requiere articulación con niveles "
                        "distrital, provincial o regional.",
            nota="F-DS-02, numeral 3."))

    return {"id": "F-DS-02", "titulo": "F-DS-02 · Mapeo de actores clave",
            "descripcion": f"{len(actores)} actor(es) registrado(s) en "
                           f"{len(regs)} ficha(s).",
            "series": series,
            "tablas": [("Actores clave", detalle)]}


# ── F-DS-03: entrevistas a autoridades ────────────────────────────────────

def _rango_edad(valor):
    edad = _num(valor)
    if edad is None:
        return ""
    if edad < 30:
        return "Menos de 30 años"
    if edad < 45:
        return "De 30 a 44 años"
    if edad < 60:
        return "De 45 a 59 años"
    return "60 años o más"


_ORDEN_EDAD = ["Menos de 30 años", "De 30 a 44 años", "De 45 a 59 años",
               "60 años o más"]


def _seccion_entrevistas(registros, tema="claro"):
    regs = _por_ficha(registros, "F-DS-03")
    if not regs:
        return None
    series, detalle = [], []
    filas_edad, generos, duraciones = [], [], []
    for reg in regs:
        f = formulario(reg)
        genero = {"M": "Hombre", "F": "Mujer"}.get(_txt(f.get("f3_genero")),
                                                   _txt(f.get("f3_genero")))
        rango = _rango_edad(f.get("f3_edad"))
        if genero:
            generos.append(genero)
        if rango:
            filas_edad.append({"cat": rango, "sub": genero or "Sin consignar",
                               "valor": 1})
        duracion = _num(f.get("f3_dur"))
        if duracion:
            duraciones.append(duracion)
        detalle.append({
            "Centro poblado / ámbito": _ambito(reg),
            "Entrevistado/a": _txt(f.get("f3_nombre")),
            "Cargo / Rol": _txt(f.get("f3_cargo")),
            "Institución / Organización": _txt(f.get("f3_inst")),
            "Género": genero, "Edad": _num(f.get("f3_edad")),
            "Años en el cargo": _num(f.get("f3_anios")),
            "Duración (min)": duracion,
            "Consiente uso de nombre": _sino(f.get("f3_c_nom")),
            "Consiente fotografías": _sino(f.get("f3_c_foto")),
        })

    if filas_edad:
        # Se agregan las filas unitarias antes de graficar.
        acumulado = {}
        for fila in filas_edad:
            clave = (fila["cat"], fila["sub"])
            acumulado[clave] = acumulado.get(clave, 0) + 1
        filas = [{"cat": c, "sub": s, "valor": v}
                 for (c, s), v in acumulado.items()]
        presentes, colores = _apiladas(filas, ["Hombre", "Mujer", "Sin consignar"],
                                       categorica=True, tema=tema)
        series.append(_serie(
            "f3_perfil", "Perfil de las autoridades y líderes entrevistados",
            "apiladas", filas, "cat", "valor", sub="sub",
            orden_cat=_ORDEN_EDAD, orden_sub=presentes, colores=colores,
            unidad="entrevistas", eje_x="N.° de entrevistas", escala="categorica",
            descripcion="Composición por rango de edad y género: evidencia la "
                        "representatividad del recojo de información.",
            nota="F-DS-03, numeral 2 (Datos del entrevistado/a)."))

    cargos = _conteo([_txt(formulario(r).get("f3_cargo")) for r in regs])
    if cargos:
        series.append(_serie(
            "f3_cargos", "Entrevistas por cargo o rol del informante",
            "barras_h", cargos, "clase", "valor", escala="neutra",
            unidad="entrevistas", eje_x="N.° de entrevistas",
            descripcion="Cargos efectivamente cubiertos por el equipo social.",
            nota="F-DS-03, numeral 2."))

    filas_cons = _bateria_sino(regs, [
        ("f3_c_nom", "Consiente el uso de su nombre en el informe"),
        ("f3_c_foto", "Consiente la toma de fotografías")])
    if filas_cons:
        series.append(_serie(
            "f3_consentimiento", "Consentimientos otorgados por los entrevistados",
            "apiladas", filas_cons, "cat", "valor", sub="sub",
            orden_sub=["Sí", "No", "No aplica"], colores=_colores_sino(tema),
            unidad="entrevistas", eje_x="N.° de entrevistas", escala="sino",
            descripcion="Condiciona el tratamiento de los datos personales en "
                        "los entregables del estudio.",
            nota="F-DS-03, numeral 2. Ley 29733 de Protección de Datos "
                 "Personales."))

    return {"id": "F-DS-03", "titulo": "F-DS-03 · Entrevistas a autoridades y líderes",
            "descripcion": f"{len(regs)} entrevista(s) registrada(s)"
                           + (f"; duración media {sum(duraciones)/len(duraciones):.0f} min."
                              if duraciones else "."),
            "series": series, "tablas": [("Entrevistas", detalle)]}


# ── F-DS-04: talleres participativos ──────────────────────────────────────

def _seccion_talleres(registros, tema="claro"):
    regs = _por_ficha(registros, "F-DS-04")
    if not regs:
        return None
    series, detalle = [], []
    filas_asist, filas_sexo, filas_etaria, metodologias = [], [], [], []
    acuerdos_total = participantes_total = 0

    for reg in regs:
        f = formulario(reg)
        etiqueta = _ambito(reg)
        fecha = _txt(f.get("f4_fecha")) or _txt(reg.get("fecha_evaluacion"))
        if fecha:
            etiqueta = f"{etiqueta} ({fecha})"
        convocados = _num(f.get("f4_conv_n"))
        hombres, mujeres = _num(f.get("f4_h")), _num(f.get("f4_m"))
        asistentes = _num(f.get("f4_tot"))
        if asistentes is None and (hombres is not None or mujeres is not None):
            asistentes = (hombres or 0) + (mujeres or 0)
        if convocados is not None:
            filas_asist.append({"cat": etiqueta, "sub": "Convocados",
                                "valor": convocados})
        if asistentes is not None:
            filas_asist.append({"cat": etiqueta, "sub": "Asistentes",
                                "valor": asistentes})
        if hombres is not None:
            filas_sexo.append({"cat": etiqueta, "sub": "Hombres", "valor": hombres})
        if mujeres is not None:
            filas_sexo.append({"cat": etiqueta, "sub": "Mujeres", "valor": mujeres})
        jovenes, mayores = _num(f.get("f4_jov")), _num(f.get("f4_am"))
        if asistentes and (jovenes is not None or mayores is not None):
            jovenes, mayores = jovenes or 0, mayores or 0
            filas_etaria += [
                {"cat": etiqueta, "sub": "Jóvenes (menores de 30)", "valor": jovenes},
                {"cat": etiqueta, "sub": "Población adulta", "valor":
                    max(asistentes - jovenes - mayores, 0)},
                {"cat": etiqueta, "sub": "Adultos mayores (más de 60)", "valor": mayores},
            ]
        metodologias += _lista(f.get("f4_metod"))
        n_part = len(_tabla(f, "f4_part"))
        n_acuerdos = len(_tabla(f, "f4_acuerdos"))
        participantes_total += n_part
        acuerdos_total += n_acuerdos
        tasa = (100.0 * asistentes / convocados
                if convocados and asistentes is not None else None)
        detalle.append({
            "Centro poblado / ámbito": _ambito(reg), "Fecha del taller": fecha,
            "Lugar": _txt(f.get("f4_lugar")),
            "Entidad convocante": _txt(f.get("f4_conv")),
            "Convocados": convocados, "Asistentes": asistentes,
            "Tasa de asistencia (%)": round(tasa, 1) if tasa is not None else None,
            "Hombres": hombres, "Mujeres": mujeres,
            "Jóvenes (<30)": _num(f.get("f4_jov")),
            "Adultos mayores (>60)": _num(f.get("f4_am")),
            "Participantes en lista": n_part, "Acuerdos registrados": n_acuerdos,
            "Idioma de la facilitación": _txt(f.get("f4_idioma")),
        })

    if filas_asist:
        presentes, colores = _apiladas(filas_asist, ["Convocados", "Asistentes"],
                                       categorica=True, tema=tema)
        series.append(_serie(
            "f4_asistencia", "Convocatoria frente a asistencia efectiva por taller",
            "agrupadas", filas_asist, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="personas",
            eje_x="N.° de personas", escala="categorica",
            descripcion="La brecha entre ambas barras mide la capacidad real "
                        "de convocatoria en cada centro poblado.",
            nota="F-DS-04, numeral 2.1 (Asistencia)."))

    if filas_sexo:
        presentes, colores = _apiladas(filas_sexo, ["Hombres", "Mujeres"],
                                       categorica=True, tema=tema)
        series.append(_serie(
            "f4_sexo", "Participación por sexo en los talleres",
            "apiladas", filas_sexo, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="personas",
            eje_x="N.° de asistentes", escala="categorica",
            descripcion="Indicador de participación equitativa exigido a los "
                        "procesos participativos del proyecto.",
            nota="F-DS-04, numeral 2.1."))

    if filas_etaria:
        presentes, colores = _apiladas(
            filas_etaria,
            ["Jóvenes (menores de 30)", "Población adulta",
             "Adultos mayores (más de 60)"], rampa=RAMPA_NEUTRA)
        series.append(_serie(
            "f4_etaria", "Composición etaria de los asistentes",
            "apiladas", filas_etaria, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="personas",
            eje_x="N.° de asistentes",
            descripcion="El tramo adulto se obtiene por diferencia con el "
                        "total de asistentes declarado.",
            nota="F-DS-04, numeral 2.1."))

    filas_metod = _conteo(metodologias)
    if filas_metod:
        series.append(_serie(
            "f4_metodologias", "Metodologías participativas empleadas",
            "barras_h", filas_metod, "clase", "valor", escala="neutra",
            unidad="talleres", eje_x="N.° de talleres",
            descripcion="Cada taller puede emplear más de una metodología.",
            nota="F-DS-04, numeral 2.2."))

    filas_acu = [{"clase": d["Centro poblado / ámbito"],
                  "valor": d["Acuerdos registrados"]}
                 for d in detalle if d["Acuerdos registrados"]]
    if filas_acu:
        series.append(_serie(
            "f4_acuerdos", "Acuerdos y compromisos suscritos por taller",
            "barras_h", filas_acu, "clase", "valor", escala="favorable",
            unidad="acuerdos", eje_x="N.° de acuerdos",
            descripcion="Los acuerdos son el medio de verificación de la "
                        "aceptación comunal del bloque.",
            nota="F-DS-04, numeral 5 (Acuerdos y compromisos)."))

    return {"id": "F-DS-04", "titulo": "F-DS-04 · Talleres participativos",
            "descripcion": f"{len(regs)} taller(es); {participantes_total} "
                           f"participante(s) en lista y {acuerdos_total} acuerdo(s).",
            "series": series, "tablas": [("Talleres participativos", detalle)]}


# ── F-DS-05: conflictos y oportunidades ───────────────────────────────────

_ORDEN_CONFLICTIVIDAD = ["Muy bajo", "Bajo", "Medio", "Alto", "Muy alto"]
_ORDEN_POLARIZACION = ["Muy bajo / Inexistente", "Bajo", "Medio", "Alto", "Muy alto"]
_ORDEN_VIABILIDAD = [
    "Alta — viable para intervencion inmediata",
    "Media — requiere acercamiento reforzado",
    "Baja — requiere mesa de dialogo previa",
    "Muy baja — reevaluar inclusion del bloque",
]
_ORDEN_PLAZO = ["Inmediato (<1 mes)", "Corto (1-3 meses)", "Medio (3-6 meses)",
                "Largo (6-12 meses)", "Requiere >12 meses"]


def _seccion_conflictos(registros, tema="claro"):
    regs = _por_ficha(registros, "F-DS-05")
    if not regs:
        return None
    series = []
    conflictos, oportunidades, detalle_sintesis = [], [], []

    for reg in regs:
        f = formulario(reg)
        ambito = _ambito(reg)
        for fila in _tabla(f, "f5_conflictos"):
            tipo = _col(fila, "Tipo")
            if not tipo:
                continue
            conflictos.append({
                "Centro poblado / ámbito": ambito, "Tipo": tipo,
                "Actores involucrados": _col(fila, "Actores involucrados"),
                "Estado": _col(fila, "Estado"),
                "Antigüedad": _col(fila, "Antiguedad", "Antigüedad"),
                "Descripción / Causa raíz": _col(fila, "Descripcion / Causa raiz",
                                                 "Descripción / Causa raíz"),
                "Impacto potencial en el proyecto":
                    _col(fila, "Impacto potencial en el proyecto"),
            })
        for fila in _tabla(f, "f5_oportunidades"):
            oportunidad = _col(fila, "Oportunidad identificada")
            if not oportunidad:
                continue
            oportunidades.append({
                "Centro poblado / ámbito": ambito,
                "Oportunidad identificada": oportunidad,
                "Actores relacionados": _col(fila, "Actores relacionados"),
                "Tipo": _col(fila, "Tipo (alianza / plataforma / proy.)", "Tipo"),
                "Potencial": _col(fila, "Potencial"),
                "Cómo aprovecharla": _col(fila, "Como aprovecharla",
                                          "Cómo aprovecharla"),
            })
        detalle_sintesis.append({
            "Centro poblado / ámbito": ambito,
            "Nivel global de conflictividad": _txt(f.get("f5_confglob")),
            "Nivel de polarización actual": _txt(f.get("f5_polar")),
            "Viabilidad social preliminar": _txt(f.get("f5_viab")),
            "Plazo estimado de aceptación comunal": _txt(f.get("f5_plazo")),
            "¿Requiere mesa de diálogo?": _sino(f.get("f5_mesa")),
            "Vínculo con el conflicto Río Blanco": _sino(f.get("f5_rb1")),
            "Estrategia de acercamiento recomendada": _txt(f.get("f5_estrategia")),
        })

    # 1. Conflictos por tipo y estado.
    if conflictos:
        acumulado = {}
        for c in conflictos:
            clave = (c["Tipo"], c["Estado"] or "Sin estado consignado")
            acumulado[clave] = acumulado.get(clave, 0) + 1
        filas = [{"cat": t, "sub": e, "valor": n} for (t, e), n in acumulado.items()]
        presentes, colores = _apiladas(filas, FL.L_NIVEL_CONFL, rampa=RAMPA_CRITICA)
        # El estado "Resuelto / Inactivo" no es un agravante: se saca de la
        # rampa critica y se pinta en verde para no leerse como riesgo vivo.
        for clave in list(colores):
            if clave.lower().startswith("resuelto"):
                colores[clave] = RAMPA_FAVORABLE[3]
        series.append(_serie(
            "f5_tipos", "Conflictos identificados por tipo y estado",
            "apiladas", filas, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="conflictos",
            eje_x="N.° de conflictos",
            descripcion="El color va del latente (claro) al activo (oscuro); "
                        "los resueltos se muestran en verde.",
            nota="F-DS-05, numeral 2 (Identificación de conflictos)."))

        filas_ant = _conteo([c["Antigüedad"] for c in conflictos],
                            orden=FL.L_ANTIG_CONFL)
        if filas_ant:
            series.append(_serie(
                "f5_antiguedad", "Antigüedad de los conflictos registrados",
                "barras_h", filas_ant, "clase", "valor", escala="critica",
                orden_cat=FL.L_ANTIG_CONFL,
                unidad="conflictos", eje_x="N.° de conflictos",
                descripcion="Los conflictos históricos exigen mesas de diálogo "
                            "previas; los recientes admiten acercamiento directo.",
                nota="F-DS-05, numeral 2."))

    # 2. Oportunidades por potencial.
    if oportunidades:
        filas_op = _conteo([o["Potencial"] for o in oportunidades],
                           orden=_ORDEN_ABC)
        if filas_op:
            series.append(_serie(
                "f5_oportunidades", "Oportunidades identificadas según potencial",
                "barras_h", filas_op, "clase", "valor", escala="favorable",
                unidad="oportunidades", eje_x="N.° de oportunidades",
                descripcion="Alianzas, plataformas y proyectos con los que el "
                            "bloque puede apalancarse.",
                nota="F-DS-05, numeral 4 (Identificación de oportunidades)."))

    # 3. Sintesis estrategica por ambito (escalas ordenadas).
    escalas = [
        ("f5_confglob", "Nivel global de conflictividad", _ORDEN_CONFLICTIVIDAD, RAMPA_CRITICA),
        ("f5_polar", "Nivel de polarización actual", _ORDEN_POLARIZACION, RAMPA_CRITICA),
        ("f5_viab", "Viabilidad social preliminar", _ORDEN_VIABILIDAD, RAMPA_CRITICA),
        ("f5_plazo", "Plazo estimado de aceptación comunal", _ORDEN_PLAZO, RAMPA_CRITICA),
    ]
    for clave, etiqueta, orden, rampa in escalas:
        filas = _conteo([_txt(formulario(r).get(clave)) for r in regs], orden=orden)
        if not filas:
            continue
        clases = [f["clase"] for f in filas]
        colores = dict(zip(clases, _rampa(clases, rampa)))
        series.append(_serie(
            f"f5_{clave}", etiqueta, "barras_h", filas, "clase", "valor",
            colores=colores, escala="critica", unidad="CP",
            eje_x="Centros poblados",
            descripcion="Escala ordenada de la ficha; a mayor intensidad de "
                        "color, mayor exigencia de gestión social previa.",
            nota="F-DS-05, numeral 5 (Síntesis estratégica)."))

    filas_rb = _bateria_sino(regs, [
        ("f5_rb1", "Vínculo histórico con el conflicto Río Blanco"),
        ("f5_rb2", "Participó en la consulta de 2007"),
        ("f5_rb3", "Persiste sentimiento anti-minero fuerte"),
        ("f5_rb4", "Liderazgos activos anti-mineros"),
        ("f5_rb5", "Se diferencia el proyecto IN del contexto minero")])
    if filas_rb:
        series.append(_serie(
            "f5_riobranco", "Contexto del conflicto minero Río Blanco",
            "apiladas", filas_rb, "cat", "valor", sub="sub",
            orden_sub=["Sí", "No", "No aplica"], colores=_colores_sino(tema),
            unidad="CP", eje_x="Centros poblados", escala="sino",
            descripcion="Antecedente determinante de la aceptación social en "
                        "Huancabamba y Ayabaca: el proyecto debe diferenciarse "
                        "explícitamente de toda actividad extractiva.",
            nota="F-DS-05, numeral 3 (Contexto específico)."))

    tablas = []
    if conflictos:
        tablas.append(("Conflictos", conflictos))
    if oportunidades:
        tablas.append(("Oportunidades", oportunidades))
    tablas.append(("Síntesis estratégica", detalle_sintesis))

    return {"id": "F-DS-05",
            "titulo": "F-DS-05 · Conflictos socioambientales y oportunidades",
            "descripcion": f"{len(conflictos)} conflicto(s) y "
                           f"{len(oportunidades)} oportunidad(es) en "
                           f"{len(regs)} ficha(s).",
            "series": series, "tablas": tablas}


# ── F-DS-06: peligros y cambio climatico ──────────────────────────────────

_ORDEN_FRECUENCIA = FL.FDS06_FRECUENCIA           # MF, F, O, R, SR
_ORDEN_MAGNITUD = FL.FDS06_MAGNITUD               # A, M, B
_ORDEN_INTENSIDAD = FL.FDS06_INTENSIDAD           # Alta, Media, Baja
_ORDEN_TENDENCIA = FL.FDS06_TENDENCIA

# La tendencia no es una magnitud sino un signo: aumentar es adverso,
# disminuir favorable y estable, neutro.
_COLOR_TENDENCIA = {
    "Sube (Aumentando)": DIVERGENTE[4],
    "I (Irregular)": DIVERGENTE[3],
    "E (Estable)": DIVERGENTE[2],
    "Baja (Disminuyendo)": DIVERGENTE[0],
    "NS (No sabe)": SIN_DATO,
}


def _seccion_peligros(registros, tema="claro"):
    regs = _por_ficha(registros, "F-DS-06")
    if not regs:
        return None
    series = []
    peligros, cambios = [], []
    prioridades = {}

    for reg in regs:
        f = formulario(reg)
        ambito = _ambito(reg)
        for fila in _tabla(f, "f6_peligros"):
            peligro = _col(fila, "Peligro observado")
            if not peligro:
                continue
            peligros.append({
                "Centro poblado / ámbito": ambito, "Peligro observado": peligro,
                "¿Ocurre?": _sino(_col(fila, "¿Ocurre?", "Ocurre")),
                "Frecuencia": _col(fila, "Frecuencia"),
                "Magnitud": _col(fila, "Magnitud"),
                "Tendencia": _col(fila, "Tendencia"),
                "Último evento (año)": _col(fila, "Ultimo evento (año)",
                                            "Último evento (año)"),
                "Principales daños observados":
                    _col(fila, "Principales daños observados"),
            })
        for fila in _tabla(f, "f6_cambios"):
            cambio = _col(fila, "Cambio observado")
            if not cambio:
                continue
            cambios.append({
                "Centro poblado / ámbito": ambito, "Cambio observado": cambio,
                "¿Se percibe?": _sino(_col(fila, "¿Se percibe?", "Se percibe")),
                "Intensidad": _col(fila, "Intensidad"),
                "Año aproximado de inicio": _col(fila, "Año aprox. de inicio"),
                "Impacto en la comunidad / territorio":
                    _col(fila, "Impacto en la comunidad / territorio"),
            })
        # Priorizacion local: el primer peligro pesa 3, el segundo 2, el tercero 1.
        for clave, peso in (("f6_p1", 3), ("f6_p2", 2), ("f6_p3", 1)):
            nombre = _txt(f.get(clave))
            if nombre:
                prioridades[nombre] = prioridades.get(nombre, 0) + peso

    ocurren = [p for p in peligros if p["¿Ocurre?"] == "Sí"]

    # 1. Peligros que ocurren, por frecuencia.
    if ocurren:
        acumulado = {}
        for p in ocurren:
            clave = (p["Peligro observado"],
                     p["Frecuencia"] or "Sin frecuencia consignada")
            acumulado[clave] = acumulado.get(clave, 0) + 1
        filas = [{"cat": c, "sub": s, "valor": v} for (c, s), v in acumulado.items()]
        presentes, colores = _apiladas(filas, _ORDEN_FRECUENCIA,
                                       rampa=RAMPA_CRITICA, descendente=True)
        series.append(_serie(
            "f6_frecuencia", "Peligros que ocurren en el territorio, según frecuencia",
            "apiladas", filas, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="CP",
            eje_x="Centros poblados que lo reportan",
            descripcion="Solo se cuentan los peligros marcados como "
                        "ocurrentes. A mayor oscuridad, mayor frecuencia.",
            nota="F-DS-06, numeral 2. Corrobora en campo el peligro integrado "
                 "modelado por MCA-AHP en la mesolocalización."))

        acumulado = {}
        for p in ocurren:
            clave = (p["Peligro observado"], p["Magnitud"] or "Sin magnitud consignada")
            acumulado[clave] = acumulado.get(clave, 0) + 1
        filas = [{"cat": c, "sub": s, "valor": v} for (c, s), v in acumulado.items()]
        presentes, colores = _apiladas(filas, _ORDEN_MAGNITUD,
                                       rampa=RAMPA_CRITICA, descendente=True)
        series.append(_serie(
            "f6_magnitud", "Magnitud percibida de los peligros ocurrentes",
            "apiladas", filas, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="CP",
            eje_x="Centros poblados que lo reportan",
            descripcion="Magnitud declarada por la población para cada peligro.",
            nota="F-DS-06, numeral 2."))

        acumulado = {}
        for p in ocurren:
            clave = (p["Peligro observado"], p["Tendencia"] or "Sin tendencia consignada")
            acumulado[clave] = acumulado.get(clave, 0) + 1
        filas = [{"cat": c, "sub": s, "valor": v} for (c, s), v in acumulado.items()]
        presentes, _ = _apiladas(filas, _ORDEN_TENDENCIA)
        colores = {s: _COLOR_TENDENCIA.get(s, SIN_DATO) for s in presentes}
        series.append(_serie(
            "f6_tendencia", "Tendencia percibida de los peligros",
            "apiladas", filas, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="CP",
            eje_x="Centros poblados que lo reportan", escala="divergente",
            descripcion="Rojo: el peligro va en aumento. Verde: disminuye. "
                        "Gris: la población no sabe.",
            nota="F-DS-06, numeral 2. Evidencia local de la tendencia a la "
                 "degradación que el proyecto busca revertir."))

    # 2. Cambios climaticos percibidos.
    percibidos = [c for c in cambios if c["¿Se percibe?"] == "Sí"]
    if percibidos:
        acumulado = {}
        for c in percibidos:
            clave = (c["Cambio observado"], c["Intensidad"] or "Sin intensidad consignada")
            acumulado[clave] = acumulado.get(clave, 0) + 1
        filas = [{"cat": c, "sub": s, "valor": v} for (c, s), v in acumulado.items()]
        presentes, colores = _apiladas(filas, _ORDEN_INTENSIDAD,
                                       rampa=RAMPA_CRITICA, descendente=True)
        series.append(_serie(
            "f6_cambios", "Cambios climáticos percibidos en los últimos 10-15 años",
            "apiladas", filas, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="CP",
            eje_x="Centros poblados que lo perciben",
            descripcion="Percepción local contrastable con los escenarios "
                        "climáticos del estudio (MPI-ESM1-2HR, SSP 5-8.5).",
            nota="F-DS-06, numeral 3. Ley 30754, Ley Marco sobre Cambio Climático."))

    # 3. Priorizacion local de peligros.
    if prioridades:
        filas = sorted(({"clase": k, "valor": v} for k, v in prioridades.items()),
                       key=lambda r: -r["valor"])
        series.append(_serie(
            "f6_prioridad", "Priorización local de peligros (índice ponderado)",
            "barras_h", filas, "clase", "valor", escala="critica",
            unidad="pts", eje_x="Índice ponderado",
            descripcion="Índice construido por el aplicativo: el peligro más "
                        "grave señalado en cada ficha suma 3 puntos, el "
                        "segundo 2 y el tercero 1.",
            nota="F-DS-06, numeral 5. El índice ordena la percepción local; "
                 "no sustituye la evaluación técnica del peligro."))

    # 4. Capacidad de respuesta local.
    filas_cap = _bateria_sino(regs, [
        ("f6_medidas", "La comunidad ha tomado medidas"),
        ("f6_alerta", "Sistemas de alerta temprana comunitarios"),
        ("f6_saberes", "Saberes tradicionales de predicción"),
        ("f6_apoyo", "Requiere apoyo externo para la adaptación")])
    if filas_cap:
        series.append(_serie(
            "f6_capacidad", "Capacidad local de respuesta y adaptación",
            "apiladas", filas_cap, "cat", "valor", sub="sub",
            orden_sub=["Sí", "No", "No aplica"], colores=_colores_sino(tema),
            unidad="CP", eje_x="Centros poblados", escala="sino",
            descripcion="Resiliencia declarada de la población: insumo del "
                        "análisis de vulnerabilidad (fragilidad y resiliencia) "
                        "del AdR-CCC.",
            nota="F-DS-06, numeral 4. Anexo 2 de la Guía General DGPMI-MEF."))

    tablas = []
    if peligros:
        tablas.append(("Peligros percibidos", peligros))
    if cambios:
        tablas.append(("Cambios climáticos", cambios))

    return {"id": "F-DS-06",
            "titulo": "F-DS-06 · Percepción de peligros y cambio climático",
            "descripcion": f"{len(ocurren)} registro(s) de peligros ocurrentes y "
                           f"{len(percibidos)} de cambios percibidos en "
                           f"{len(regs)} ficha(s).",
            "series": series, "tablas": tablas}


# ── F-DS-07: disposicion a participar y consentimiento ────────────────────

_PUNTOS_CPI = [
    ("f7_info_anin", "Qué es la ANIN"),
    ("f7_info_objetivo", "Objetivo del proyecto IN Piura"),
    ("f7_info_no_minero", "No es actividad minera ni extractiva"),
    ("f7_info_medidas", "Medidas que podrían implementarse en el predio"),
    ("f7_info_temporalidad", "Temporalidad del proyecto"),
    ("f7_info_voluntaria", "La participación es voluntaria"),
    ("f7_info_actualizada", "Derecho a información actualizada"),
    ("f7_info_confidencialidad", "Tratamiento confidencial de los datos"),
    ("f7_info_preguntas", "Se absolvieron todas las preguntas"),
    ("f7_info_material", "Se entregó material informativo"),
]

_COLOR_DISPOSICION = {
    "Totalmente dispuesto/a": DIVERGENTE[0],
    "Dispuesto/a con condiciones": DIVERGENTE[1],
    "Requiere más información": DIVERGENTE[2],
    "No dispuesto/a por el momento": DIVERGENTE[3],
    "No dispuesto/a rotundamente": DIVERGENTE[4],
}


def _seccion_consentimiento(registros, tema="claro"):
    regs = _por_ficha(registros, "F-DS-07")
    if not regs:
        return None
    series, detalle = [], []
    superficies, documentos = [], []

    for reg in regs:
        f = formulario(reg)
        superficie = _num(f.get("f7_superficie"))
        if superficie:
            superficies.append({"clase": _txt(f.get("f7_nombre")) or "(sin nombre)",
                                "valor": superficie})
        documentos += _lista(f.get("f7_docs"))
        informados = sum(1 for clave, _ in _PUNTOS_CPI
                         if _sino(f.get(clave)) == "Sí")
        detalle.append({
            "Centro poblado / ámbito": _ambito(reg),
            "Titular / Representante": _txt(f.get("f7_nombre")),
            "DNI": _txt(f.get("f7_dni")),
            "Tipo de propietario": _txt(f.get("f7_tipo_prop")),
            "Superficie en el bloque (ha)": superficie,
            "Puntos del CPI explicados (de 10)": informados,
            "Disposición a participar": _txt(f.get("f7_disp")),
            "Autoriza ingreso preliminar": _sino(f.get("f7_aut_ing")),
            "Autoriza fotografías": _sino(f.get("f7_aut_foto")),
            "Autoriza uso de su nombre": _sino(f.get("f7_aut_nom")),
            "Conflictos de linderos": _sino(f.get("f7_linderos")),
            "Residente permanente": _sino(f.get("f7_residente")),
        })

    # 1. Disposicion a participar (escala ordenada favorable <-> desfavorable).
    filas_disp = _conteo([_txt(formulario(r).get("f7_disp")) for r in regs],
                         orden=FL.L_DISPOSICION)
    if filas_disp:
        series.append(_serie(
            "f7_disposicion", "Disposición de los titulares a participar",
            "barras_h", filas_disp, "clase", "valor",
            colores=_COLOR_DISPOSICION, escala="divergente",
            unidad="titulares", eje_x="N.° de titulares",
            descripcion="Del verde (totalmente dispuesto) al rojo (no "
                        "dispuesto); el gris marca a quienes requieren más "
                        "información antes de decidir.",
            nota="F-DS-07, numeral 4. Es el indicador de aceptación predial "
                 "que condiciona la inclusión definitiva del bloque."))

    # 2. Consentimiento previo informado: cobertura de los 10 puntos.
    filas_cpi = _bateria_sino(regs, _PUNTOS_CPI)
    if filas_cpi:
        series.append(_serie(
            "f7_cpi", "Consentimiento previo informado: puntos efectivamente explicados",
            "apiladas", filas_cpi, "cat", "valor", sub="sub",
            orden_sub=["Sí", "No", "No aplica"], colores=_colores_sino(tema),
            unidad="titulares", eje_x="N.° de titulares", escala="sino",
            descripcion="Un punto en rojo es una brecha del proceso de "
                        "consentimiento que debe subsanarse antes de "
                        "formalizar el acuerdo.",
            nota="F-DS-07, numeral 3. Convenio 169 de la OIT y estándares "
                 "de salvaguardas sociales."))

    # 3. Autorizaciones otorgadas.
    filas_aut = _bateria_sino(regs, [
        ("f7_aut_ing", "Autoriza el ingreso preliminar al predio"),
        ("f7_aut_foto", "Autoriza fotografías del predio"),
        ("f7_aut_nom", "Autoriza el uso de su nombre"),
        ("f7_consultar", "Requiere consultar con familia o asamblea"),
        ("f7_linderos", "Declara conflictos de linderos"),
        ("f7_residente", "Es residente permanente")])
    if filas_aut:
        series.append(_serie(
            "f7_autorizaciones", "Autorizaciones y condiciones declaradas por los titulares",
            "apiladas", filas_aut, "cat", "valor", sub="sub",
            orden_sub=["Sí", "No", "No aplica"], colores=_colores_sino(tema),
            unidad="titulares", eje_x="N.° de titulares", escala="sino",
            descripcion="Las tres primeras habilitan el trabajo de campo; las "
                        "tres últimas describen condiciones del predio.",
            nota="F-DS-07, numerales 2 y 4."))

    # 4. Tipo de propietario y documentacion de tenencia.
    filas_tipo = _conteo([_txt(formulario(r).get("f7_tipo_prop")) for r in regs],
                         orden=FL.L_TIPO_PROPIETARIO)
    if filas_tipo:
        series.append(_serie(
            "f7_tipo_prop", "Titulares y representantes por tipo",
            "barras_h", filas_tipo, "clase", "valor", escala="neutra",
            unidad="titulares", eje_x="N.° de titulares",
            descripcion="Define el instrumento jurídico aplicable a cada "
                        "acuerdo de intervención.",
            nota="F-DS-07, numeral 2.1."))

    filas_doc = _conteo(documentos)
    if filas_doc:
        series.append(_serie(
            "f7_documentos", "Documentación de tenencia disponible",
            "barras_h", filas_doc, "clase", "valor", escala="neutra",
            unidad="menciones", eje_x="N.° de menciones",
            descripcion="Un mismo titular puede presentar más de un documento.",
            nota="F-DS-07, numeral 2.2. Insumo del tamizaje predial del bloque."))

    if superficies:
        total_ha = sum(s["valor"] for s in superficies)
        series.append(_serie(
            "f7_superficies", "Superficie predial comprometida por titular (ha)",
            "barras_h", sorted(superficies, key=lambda r: -r["valor"]),
            "clase", "valor", escala="favorable", unidad="ha",
            eje_x="Superficie (ha)", decimales=2,
            descripcion=f"Suma declarada: {total_ha:,.2f} ha dentro del bloque."
                        .replace(",", " "),
            nota="F-DS-07, numeral 2. Superficie declarada por el titular; no "
                 "sustituye la medición predial del expediente."))

    return {"id": "F-DS-07",
            "titulo": "F-DS-07 · Disposición a participar y consentimiento previo",
            "descripcion": f"{len(regs)} titular(es) / representante(s) registrado(s).",
            "series": series, "tablas": [("Titulares y consentimiento", detalle)]}


# ══════════════════════════════════════════════════════════════════════════
# COBERTURA DEL DIAGNOSTICO Y ARMADO DEL INFORME
# ══════════════════════════════════════════════════════════════════════════

def _seccion_cobertura(registros, centros_catalogo=(), tema="claro"):
    """Avance del levantamiento: que fichas hay y en que centros poblados."""
    if not registros:
        return None
    series = []

    # 1. Fichas registradas por tipo.
    conteo_ficha = {f: len(_por_ficha(registros, f)) for f in FICHAS_DS}
    filas = [{"clase": f"{f} · {FICHAS_DS_TITULOS[f]}", "valor": conteo_ficha[f]}
             for f in FICHAS_DS]
    series.append(_serie(
        "cob_fichas", "Fichas sociales registradas por tipo",
        "barras_h", filas, "clase", "valor", escala="favorable",
        unidad="fichas", eje_x="N.° de fichas registradas", orden_cat=
        [f["clase"] for f in filas],
        descripcion="Las siete fichas de la Plantilla V4. Una barra vacía "
                    "señala un instrumento aún no aplicado en el bloque.",
        nota="Conteo tras descartar reediciones (se conserva el registro más "
             "reciente por ficha, centro poblado y responsable)."))

    # 2. Fichas por centro poblado.
    filas_cp = []
    acumulado = {}
    for reg in registros:
        clave = (_ambito(reg), reg.get("ficha", "") or "Sin ficha")
        acumulado[clave] = acumulado.get(clave, 0) + 1
    for (ambito, ficha), n in acumulado.items():
        filas_cp.append({"cat": ambito, "sub": ficha, "valor": n})
    if filas_cp:
        presentes, colores = _apiladas(filas_cp, FICHAS_DS, categorica=True,
                                       tema=tema)
        series.append(_serie(
            "cob_cp", "Cobertura del diagnóstico social por centro poblado",
            "apiladas", filas_cp, "cat", "valor", sub="sub",
            orden_sub=presentes, colores=colores, unidad="fichas",
            eje_x="N.° de fichas", escala="categorica",
            descripcion="Cada color es una ficha distinta: permite ver de un "
                        "vistazo qué centro poblado quedó incompleto.",
            nota="Ámbito declarado en cada ficha (centro poblado, comunidad "
                 "campesina o, a falta de ambos, el propio bloque)."))

    # 3. Centros poblados del catalogo aun sin ficha.
    if centros_catalogo:
        registrados = {_ambito(r).strip().lower() for r in registros}
        filas_pend = [
            {"clase": "Con al menos una ficha social",
             "valor": sum(1 for c in centros_catalogo
                          if c.strip().lower() in registrados)},
            {"clase": "Sin ficha social registrada",
             "valor": sum(1 for c in centros_catalogo
                          if c.strip().lower() not in registrados)},
        ]
        if filas_pend[1]["valor"] or filas_pend[0]["valor"]:
            series.append(_serie(
                "cob_catalogo", "Centros poblados del catálogo oficial cubiertos",
                "barras_h", filas_pend, "clase", "valor",
                colores={"Con al menos una ficha social": TINTAS[tema]["si"],
                         "Sin ficha social registrada": TINTAS[tema]["no"]},
                escala="sino", unidad="CP", eje_x="Centros poblados",
                descripcion="Comparación contra la relación oficial de centros "
                            "poblados asociados al bloque (INEI).",
                nota="El cotejo es por nombre; un centro poblado escrito de "
                     "distinta forma en campo aparecerá como no cubierto."))

    return {"id": "COBERTURA", "titulo": "Cobertura del diagnóstico social",
            "descripcion": f"{len(registros)} ficha(s) vigente(s).",
            "series": series, "tablas": []}


_CONSTRUCTORES = [
    _seccion_socioeconomica, _seccion_actores, _seccion_entrevistas,
    _seccion_talleres, _seccion_conflictos, _seccion_peligros,
    _seccion_consentimiento,
]


def _metricas(registros, datos_cp, secciones):
    """Cifras de cabecera del informe (fila de indicadores del aplicativo)."""
    regs01 = _por_ficha(registros, "F-DS-01")
    poblacion = sum(_num(formulario(r).get("f1_pob_t")) or 0 for r in regs01)
    if not poblacion:
        poblacion = (datos_cp or {}).get("poblacion_total", 0) or 0
    familias = sum(_num(formulario(r).get("f1_nfam")) or 0 for r in regs01)

    actores = sum(len(_tabla(formulario(r), "f2_actores"))
                  for r in _por_ficha(registros, "F-DS-02"))
    a_favor = 0
    for r in _por_ficha(registros, "F-DS-02"):
        for fila in _tabla(formulario(r), "f2_actores"):
            if _col(fila, "Posicion", "Posición").lower().startswith("a favor"):
                a_favor += 1

    conflictos = sum(len(_tabla(formulario(r), "f5_conflictos"))
                     for r in _por_ficha(registros, "F-DS-05"))
    activos = 0
    for r in _por_ficha(registros, "F-DS-05"):
        for fila in _tabla(formulario(r), "f5_conflictos"):
            estado = _col(fila, "Estado").lower()
            if estado.startswith("activo") or estado.startswith("en escalada"):
                activos += 1

    regs07 = _por_ficha(registros, "F-DS-07")
    dispuestos = sum(1 for r in regs07
                     if _txt(formulario(r).get("f7_disp")).lower()
                     .startswith(("totalmente dispuesto", "dispuesto")))
    superficie = sum(_num(formulario(r).get("f7_superficie")) or 0 for r in regs07)

    asistentes = 0
    for r in _por_ficha(registros, "F-DS-04"):
        f = formulario(r)
        total = _num(f.get("f4_tot"))
        if total is None:
            total = (_num(f.get("f4_h")) or 0) + (_num(f.get("f4_m")) or 0)
        asistentes += total or 0

    ambitos = {_ambito(r) for r in registros}
    fichas_con_datos = {r.get("ficha", "") for r in registros}

    metricas = [
        {"etiqueta": "Fichas sociales vigentes", "valor": f"{len(registros)}",
         "detalle": f"{len(fichas_con_datos)} de 7 tipos aplicados"},
        {"etiqueta": "Centros poblados cubiertos", "valor": f"{len(ambitos)}",
         "detalle": "ámbitos con al menos una ficha"},
        {"etiqueta": "Población del ámbito", "valor": f"{int(poblacion):,}".replace(",", " "),
         "detalle": (f"{int(familias):,} familias".replace(",", " ")
                     if familias else "hab. declarados en F-DS-01")},
        {"etiqueta": "Asistentes a talleres", "valor": f"{int(asistentes):,}".replace(",", " "),
         "detalle": f"{len(_por_ficha(registros, 'F-DS-04'))} taller(es)"},
        {"etiqueta": "Actores mapeados", "valor": f"{actores}",
         "detalle": f"{a_favor} declaradamente a favor",
         "tono": "favorable" if a_favor else "neutro"},
        {"etiqueta": "Conflictos identificados", "valor": f"{conflictos}",
         "detalle": f"{activos} activo(s) o en escalada",
         "tono": "critico" if activos else "favorable"},
        {"etiqueta": "Titulares dispuestos a participar",
         "valor": f"{dispuestos} / {len(regs07)}" if regs07 else "s/d",
         "detalle": "F-DS-07 (consentimiento previo informado)",
         "tono": "favorable" if regs07 and dispuestos >= len(regs07) / 2 else "critico"},
        {"etiqueta": "Superficie predial comprometida",
         "valor": f"{superficie:,.2f} ha".replace(",", " ") if superficie else "s/d",
         "detalle": "declarada por los titulares en F-DS-07"},
    ]
    return metricas


def indicadores_bloque(bloque, registros, datos_cp=None, tema="claro"):
    """Informe analitico del Diagnostico Social de UN bloque.

    `bloque`   dict del bloque (database.obtener_bloque_por_id).
    `registros` fichas sociales del bloque, tal como las devuelve la BD.
    `datos_cp` entrada del catalogo de centros poblados del bloque.

    Devuelve un dict con metricas de cabecera, secciones (cada una con sus
    series graficables y sus tablas de respaldo) y avisos.
    """
    bloque = bloque or {}
    datos_cp = datos_cp or {}
    registros = deduplicar(registros)
    centros = [c for c in (datos_cp.get("centros_poblados") or []) if _txt(c)]

    secciones = []
    cobertura = _seccion_cobertura(registros, centros, tema)
    if cobertura:
        secciones.append(cobertura)
    for constructor in _CONSTRUCTORES:
        seccion = constructor(registros, tema)
        if seccion and seccion["series"]:
            secciones.append(seccion)

    avisos = []
    faltantes = [f for f in FICHAS_DS if not _por_ficha(registros, f)]
    if faltantes:
        avisos.append("Sin registros de " + ", ".join(faltantes) +
                      ": las secciones correspondientes no se grafican.")
    if not registros:
        avisos.append("El bloque no tiene fichas sociales registradas.")

    return {
        "alcance": "bloque",
        "codigo": _txt(bloque.get("codigo")) or _txt(datos_cp.get("codigo")),
        "bloque": bloque,
        "centros_poblados": centros,
        "n_registros": len(registros),
        "metricas": _metricas(registros, datos_cp, secciones),
        "secciones": secciones,
        "avisos": avisos,
        "registros": registros,
        "generado": datetime.now(),
        "tema": tema,
    }


def indicadores_consolidado(registros, etiqueta="", tema="claro"):
    """Informe analitico de un conjunto de bloques (todo el ambito o un filtro).

    Agrega ademas la distribucion de fichas por bloque, provincia y distrito,
    que es lo que distingue la mirada consolidada de la de un solo bloque.
    """
    registros = deduplicar(registros)
    secciones = []
    cobertura = _seccion_cobertura(registros, (), tema)
    if cobertura:
        # En la mirada consolidada interesa el reparto por bloque y por
        # distrito, no el centro poblado individual.
        for clave, titulo, id_ in (("bloque_codigo", "Fichas sociales por bloque", "cons_bloque"),
                                   ("distrito", "Fichas sociales por distrito", "cons_distrito"),
                                   ("provincia", "Fichas sociales por provincia", "cons_provincia")):
            filas = _conteo([_txt(r.get(clave)) for r in registros],
                            agrupar_otros=(clave != "bloque_codigo"))
            if filas:
                cobertura["series"].append(_serie(
                    id_, titulo, "barras_h", filas, "clase", "valor",
                    escala="neutra", unidad="fichas", eje_x="N.° de fichas",
                    descripcion="Distribución del levantamiento social en el "
                                "ámbito seleccionado.",
                    nota="Cabecera de cada ficha registrada."))
        secciones.append(cobertura)
    for constructor in _CONSTRUCTORES:
        seccion = constructor(registros, tema)
        if seccion and seccion["series"]:
            secciones.append(seccion)

    bloques = sorted({_txt(r.get("bloque_codigo")) for r in registros if _txt(r.get("bloque_codigo"))})
    metricas = _metricas(registros, {}, secciones)
    metricas.insert(0, {"etiqueta": "Bloques con diagnóstico social",
                        "valor": f"{len(bloques)}",
                        "detalle": etiqueta or "ámbito seleccionado"})
    return {
        "alcance": "consolidado",
        "codigo": etiqueta or f"{len(bloques)} bloques",
        "bloque": {}, "centros_poblados": [],
        "n_registros": len(registros),
        "metricas": metricas, "secciones": secciones,
        "avisos": ([] if registros else
                   ["No hay fichas sociales en el ámbito seleccionado."]),
        "registros": registros, "bloques": bloques,
        "generado": datetime.now(), "tema": tema,
    }


def _hoja_resumen(wb, informe):
    """Portada del libro: identificacion, cifras de cabecera y avisos."""
    ws = wb.active
    ws.title = "Resumen DS"
    titulo = (f"DIAGNOSTICO SOCIAL - BLOQUE {informe['codigo']}"
              if informe.get("alcance") == "bloque"
              else f"DIAGNOSTICO SOCIAL CONSOLIDADO - {informe['codigo']}")
    fila = _titulo_hoja(ws, titulo, ancho=6)
    for col, ancho in zip("ABCDEF", (42, 24, 34, 16, 16, 16)):
        ws.column_dimensions[col].width = ancho

    bloque = informe.get("bloque") or {}
    generales = [
        ("Código del bloque", _txt(bloque.get("codigo")) or informe["codigo"]),
        ("Microcuenca", _txt(bloque.get("microcuenca"))),
        ("Provincia", _txt(bloque.get("provincia"))),
        ("Distrito", _txt(bloque.get("distrito"))),
        ("Centros poblados del catálogo",
         ", ".join(informe.get("centros_poblados") or []) or "—"),
        ("Fichas sociales vigentes", informe.get("n_registros", 0)),
        ("Fecha de emisión", informe["generado"].strftime("%d/%m/%Y %H:%M")),
    ]
    fila_cab, fila = _escribir_tabla(
        ws, fila, "1. Identificación", ["Campo", "Valor"],
        [[e, v] for e, v in generales if v not in ("", None)])
    fila += 1

    fila_cab, fila = _escribir_tabla(
        ws, fila, "2. Cifras de cabecera",
        ["Indicador", "Valor", "Detalle"],
        [[m["etiqueta"], m["valor"], m.get("detalle", "")]
         for m in informe.get("metricas", [])])
    fila += 1

    inventario = [[s["titulo"], len(s.get("series", [])),
                   s.get("descripcion", "")]
                  for s in informe.get("secciones", [])]
    if inventario:
        fila_cab, fila = _escribir_tabla(
            ws, fila, "3. Contenido del libro",
            ["Sección", "N.° de gráficos", "Alcance"], inventario,
            subtitulo="Cada sección ocupa una hoja con su tabla de datos y su "
                      "gráfico nativo de Excel, editable por el usuario.")
        fila += 1

    if informe.get("avisos"):
        for aviso in informe["avisos"]:
            celda = ws.cell(fila, 1, "Aviso: " + aviso)
            celda.font = Font(name="Arial", size=9, italic=True, color="9C5700")
            ws.merge_cells(start_row=fila, start_column=1, end_row=fila,
                           end_column=6)
            fila += 1
        fila += 1

    celda = ws.cell(fila, 1,
                    "Libro generado por el aplicativo IN Piura sobre las fichas "
                    "F-DS-01 a F-DS-07 registradas. Solo se grafica lo "
                    "declarado en campo: los campos sin respuesta no se "
                    "estiman ni se completan por analogía, conforme a la "
                    "declaración de integridad de datos del proyecto.")
    celda.font = Font(name="Arial", size=8, italic=True)
    celda.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila + 2, end_column=6)
    ws.sheet_view.showGridLines = False


def generar_excel_social(informe):
    """Libro Excel del Diagnostico Social con graficos nativos. Devuelve bytes."""
    wb = Workbook()
    usados = {"resumen ds"}
    _hoja_resumen(wb, informe)
    for seccion in informe.get("secciones", []):
        if seccion.get("series"):
            _hoja_seccion(wb, seccion, usados)
    for seccion in informe.get("secciones", []):
        _hoja_tablas(wb, seccion, usados)
    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue()


def nombre_excel(informe):
    marca = datetime.now().strftime("%Y%m%d_%H%M%S")
    if informe.get("alcance") == "bloque":
        return f"Graficos_DS_Bloque_{informe['codigo']}_IN_Piura_{marca}.xlsx"
    return f"Graficos_DS_Consolidado_IN_Piura_{marca}.xlsx"


def generar_pdf_social(informe):
    """Anexo grafico en PDF del Diagnostico Social. Devuelve bytes."""
    from resumenes_bloques import _PDFResumen, _pdf_bytes

    subtitulo = ("ANEXO GRAFICO DEL DIAGNOSTICO SOCIAL - BLOQUE "
                 f"{informe['codigo']}" if informe.get("alcance") == "bloque"
                 else "ANEXO GRAFICO DEL DIAGNOSTICO SOCIAL CONSOLIDADO")
    pdf = _PDFResumen(subtitulo=subtitulo)
    pdf.alias_nb_pages()
    pdf.add_page()

    bloque = informe.get("bloque") or {}
    pdf.seccion("1. Identificacion")
    generales = [
        ("Codigo del bloque", _txt(bloque.get("codigo")) or informe["codigo"]),
        ("Microcuenca", _txt(bloque.get("microcuenca"))),
        ("Provincia", _txt(bloque.get("provincia"))),
        ("Distrito", _txt(bloque.get("distrito"))),
        ("Centros poblados", ", ".join(informe.get("centros_poblados") or []) or "-"),
        ("Fichas sociales vigentes", str(informe.get("n_registros", 0))),
        ("Fecha de emision", informe["generado"].strftime("%d/%m/%Y %H:%M")),
    ]
    pdf.tabla(["Campo", "Valor"], [[e, v] for e, v in generales if v],
              anchos_rel=[1, 2], alineaciones=["L", "L"])

    pdf.seccion("2. Cifras de cabecera")
    pdf.tabla(["Indicador", "Valor", "Detalle"],
              [[m["etiqueta"], m["valor"], m.get("detalle", "")]
               for m in informe.get("metricas", [])],
              anchos_rel=[3, 2, 4], alineaciones=["L", "R", "L"])

    for i, seccion in enumerate(informe.get("secciones", []), start=3):
        series = seccion.get("series") or []
        if not series:
            continue
        # El titulo de seccion no puede quedar solo al pie de la pagina: se
        # reserva sitio para el primer grafico antes de escribirlo.
        pdf._salto_si_falta(48)
        pdf.seccion(f"{i}. {seccion['titulo'].replace('·', '-')}")
        for serie in series:
            if serie.get("forma") in ("apiladas", "agrupadas", "mapa_calor"):
                _pdf_barras_apiladas(pdf, serie)
            else:
                _pdf_barras_simples(pdf, serie)

    if informe.get("avisos"):
        pdf.seccion("Avisos")
        for aviso in informe["avisos"]:
            pdf.nota(aviso)
    pdf.nota("Anexo generado por el aplicativo IN Piura sobre las fichas "
             "F-DS-01 a F-DS-07 registradas. Solo se grafica lo declarado en "
             "campo; los campos sin respuesta no se estiman.")
    return _pdf_bytes(pdf)


# Nombres comunes a los dos modulos de analitica, para que el aplicativo
# pueda pedirle el libro o el anexo a cualquiera de ellos sin distinguir.
generar_excel = generar_excel_social
generar_pdf = generar_pdf_social


def nombre_pdf(informe):
    marca = datetime.now().strftime("%Y%m%d_%H%M%S")
    if informe.get("alcance") == "bloque":
        return f"Anexo_Graficos_DS_Bloque_{informe['codigo']}_IN_Piura_{marca}.pdf"
    return f"Anexo_Graficos_DS_Consolidado_IN_Piura_{marca}.pdf"
