"""
Motor de series analiticas - Proyecto IN Piura.

Reune lo que comparten los informes analiticos del Diagnostico Social y del
Diagnostico Territorial: la paleta institucional validada, el modelo de
`serie` -un dict autodescriptivo que sirve por igual al grafico interactivo,
al grafico nativo de Excel y al anexo PDF- y los tres renderizadores que lo
consumen.

Un modulo de analitica solo tiene que construir sus secciones como listas de
series; de dibujarlas se encarga este. Asi los dos diagnosticos salen con la
misma factura y una mejora de presentacion beneficia a ambos a la vez.

No depende de Streamlit ni de la base de datos: recibe estructuras de datos
puras y devuelve estructuras de datos, graficos de Altair o bytes.

ANIN - DIME - SESDI | CUI 2669244 | UTM WGS 84 Zona 17S (EPSG:32717).
"""

import math
import re

from openpyxl.chart import BarChart, Reference
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ══════════════════════════════════════════════════════════════════════════
# PALETA
# ══════════════════════════════════════════════════════════════════════════
# Los colores no se eligieron a ojo: se validaron con el verificador de
# paletas (banda de luminosidad, piso de croma, separacion bajo simulacion
# de daltonismo y contraste contra el fondo real de cada tema). Cada juego
# declara al lado el peor par adyacente medido, para que quien lo modifique
# vuelva a correr la validacion y no rompa la legibilidad.
#
# Regla de uso (una sola por grafico):
#   - magnitud sin signo  -> rampa secuencial de un solo tono (NEUTRA)
#   - magnitud adversa    -> rampa CRITICA (a mas oscuro, mas grave)
#   - magnitud favorable  -> rampa FAVORABLE
#   - identidad           -> paleta CATEGORICA, en orden fijo, nunca ciclada
#   - escala ordenada     -> paleta DIVERGENTE, centrada en el neutro
#   - estado si/no        -> par SI/NO (siempre con leyenda y tabla al lado)

ANIN_VERDE = "1B4D2E"
ANIN_AZUL = "1B4F72"
ANIN_VERDE_CLARO = "E8F0EA"
ANIN_GRIS = "F2F2F2"

ENCABEZADOS_ANIN = [
    "AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN",
    "DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME",
    "SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN - SESDI",
]

SUBTITULO_PROYECTO = (
    "PROYECTO IN PIURA | CUI 2669244 | Recuperación del servicio de regulación "
    "de riesgos naturales y recuperación de ecosistemas degradados en la "
    "Cuenca Alta del Río Piura | UTM WGS 84 Zona 17S (EPSG:32717)"
)

# Identidad (8 ranuras, orden fijo). Claro: peor par adyacente DeltaE CVD 14.0,
# vision normal 17.3 sobre #FFFFFF. Oscuro: 8.6 / 19.3 sobre #0E1117.
CATEGORICA_CLARA = ["#2E7D4F", "#1F6FB2", "#C4A03C", "#C0392B",
                    "#6A4C93", "#1BA39C", "#D2792E", "#C2477F"]
CATEGORICA_OSCURA = ["#199E70", "#3987E5", "#D95926", "#9085E9",
                     "#E66767", "#008300", "#D55181", "#C98500"]

# Rampas ordinales de un solo tono (claro -> oscuro = menos -> mas).
RAMPA_NEUTRA = ["#93BCDD", "#6BA1CD", "#4785B8", "#28679F", "#164C77"]
RAMPA_CRITICA = ["#E8A08E", "#D97B62", "#C55437", "#A63A22", "#8A2F18"]
RAMPA_FAVORABLE = ["#8FC2A0", "#66AC7C", "#40945C", "#2A7A48", "#14592F"]

# Divergente para escalas ordenadas (favorable <-> desfavorable) con neutro
# gris al centro. Peor par adyacente DeltaE CVD 10.1, vision normal 15.1.
DIVERGENTE = ["#14592F", "#66AC7C", "#C2BFB4", "#D97B62", "#8A2F18"]

SIN_DATO = "#8A8F98"

TINTAS = {
    "claro": {
        "categorica": CATEGORICA_CLARA,
        "si": "#2E7D4F", "no": "#C0392B", "sin_dato": SIN_DATO,
        "fondo": "#FFFFFF", "tinta": "#1A1A18", "tinta_2": "#52514E",
        "apagado": "#7A7975", "rejilla": "#E4E3DC",
    },
    "oscuro": {
        "categorica": CATEGORICA_OSCURA,
        "si": "#199E70", "no": "#E66767", "sin_dato": SIN_DATO,
        "fondo": "#0E1117", "tinta": "#F5F5F2", "tinta_2": "#C3C2B7",
        "apagado": "#898781", "rejilla": "#2C2C2A",
    },
}

FUENTE = "Arial, Helvetica, sans-serif"

# Maximo de clases con color propio: pasado ese punto las clases adyacentes
# se confunden, de modo que el resto se agrupa en "Otros".
MAX_CLASES = 8


def _rampa(valores, rampa):
    """Reparte una rampa de 5 pasos entre `n` clases ordenadas."""
    n = len(valores)
    if n == 0:
        return []
    if n == 1:
        return [rampa[len(rampa) // 2]]
    return [rampa[round(i * (len(rampa) - 1) / (n - 1))] for i in range(n)]


def _divergente(valores):
    """Reparte la paleta divergente entre `n` clases ordenadas."""
    return _rampa(valores, DIVERGENTE)


def _categorica(valores, tema="claro"):
    paleta = TINTAS[tema]["categorica"]
    return [paleta[i % len(paleta)] for i in range(len(valores))]


_SI = {"si", "sí", "s", "yes", "1", "true", "x"}
_NO = {"no", "n", "0", "false"}


def _txt(valor):
    """Texto limpio; los vacios y marcadores de ausencia devuelven ''."""
    if valor is None:
        return ""
    texto = str(valor).strip()
    if texto.lower() in ("", "nan", "none", "-", "s/d", "sin dato"):
        return ""
    return texto


def _num(valor):
    """Numero a partir del texto libre de campo (tolera %, unidades y miles).

    Los campos numericos de las fichas son cajas de texto: el tecnico escribe
    "1,250", "35 %", "12 ha" o "aprox. 40". Se recupera la primera cifra y se
    devuelve None cuando no hay ninguna, para no confundir "sin dato" con 0.
    """
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return None if isinstance(valor, float) and math.isnan(valor) else float(valor)
    texto = str(valor).strip()
    if not texto:
        return None
    texto = texto.replace(" ", "")
    # Separador de miles con coma y decimal con punto (1,250.5) o al reves.
    if re.search(r",\d{3}(\D|$)", texto):
        texto = texto.replace(",", "")
    else:
        texto = texto.replace(",", ".")
    m = re.search(r"-?\d+(?:\.\d+)?", texto)
    if not m:
        return None
    try:
        return float(m.group(0))
    except ValueError:
        return None


def _sino(valor):
    """Normaliza a 'Si', 'No', 'No aplica' o '' (sin responder)."""
    texto = _txt(valor).lower()
    if not texto:
        return ""
    if texto in _SI:
        return "Sí"
    if texto in _NO:
        return "No"
    if texto.startswith("no aplica") or texto in ("na", "n/a"):
        return "No aplica"
    if texto.startswith("parcial"):
        return "Parcial"
    return _txt(valor)


def _lista(valor):
    """Valor de un multiselect: lista de opciones marcadas."""
    if isinstance(valor, (list, tuple, set)):
        return [_txt(v) for v in valor if _txt(v)]
    texto = _txt(valor)
    if not texto:
        return []
    if ";" in texto:
        return [p.strip() for p in texto.split(";") if p.strip()]
    return [texto]


_TILDES = str.maketrans("áéíóúüñÁÉÍÓÚÜÑ", "aeiouunAEIOUUN")


def _clave(texto):
    return re.sub(r"[^a-z0-9]", "", str(texto).translate(_TILDES).lower())


# ══════════════════════════════════════════════════════════════════════════
# CONSTRUCTORES DE SERIES
# ══════════════════════════════════════════════════════════════════════════
# Cada serie es un dict autodescriptivo que sirve por igual al grafico
# interactivo (Altair), al grafico nativo de Excel y al anexo PDF.

def _serie(id_, titulo, forma, filas, cat, val, **extra):
    serie = {
        "id": id_, "titulo": titulo, "forma": forma, "filas": filas,
        "cat": cat, "val": val, "sub": None, "orden_cat": None,
        "orden_sub": None, "escala": "neutra", "colores": None,
        "unidad": "", "nota": "", "descripcion": "", "eje_x": "", "eje_y": "",
        "decimales": 0,
    }
    serie.update(extra)
    # Una serie con color declarado por clase viene de una escala ordenada
    # (disposicion, posicion, nivel de conflictividad): el eje debe respetar
    # ese orden, no reordenarse por frecuencia.
    if serie.get("colores") and not serie.get("orden_cat") and not serie.get("sub"):
        serie["orden_cat"] = list(serie["colores"])
    return serie


def _conteo(valores, orden=None, agrupar_otros=True):
    """Frecuencia de cada valor, respetando un orden dado si se entrega."""
    conteo = {}
    for v in valores:
        texto = _txt(v)
        if not texto:
            continue
        conteo[texto] = conteo.get(texto, 0) + 1
    if orden:
        claves = [o for o in orden if o in conteo]
        claves += sorted(k for k in conteo if k not in orden)
    else:
        claves = sorted(conteo, key=lambda k: (-conteo[k], k))
        if agrupar_otros and len(claves) > MAX_CLASES:
            # Nunca se generan colores nuevos: lo que excede la paleta se
            # agrupa en "Otros" y el detalle queda en la tabla.
            cabeza, cola = claves[:MAX_CLASES - 1], claves[MAX_CLASES - 1:]
            resto = sum(conteo[k] for k in cola)
            conteo = {k: conteo[k] for k in cabeza}
            conteo[f"Otros ({len(cola)} clases)"] = resto
            claves = cabeza + [f"Otros ({len(cola)} clases)"]
    return [{"clase": k, "valor": conteo[k]} for k in claves]


def _apiladas(filas_grupo, orden_sub, rampa=None, categorica=False,
              tema="claro", descendente=False):
    """Normaliza filas {cat, sub, valor} y fija el color de cada subclase.

    El tono se reparte sobre el dominio COMPLETO declarado en la ficha, no
    solo sobre las clases presentes: asi una misma clase conserva su color
    entre un bloque y otro y los graficos se pueden comparar entre si.
    """
    dominio = list(dict.fromkeys(orden_sub))
    if categorica:
        tonos = dict(zip(dominio, _categorica(dominio, tema)))
    else:
        pasos = _rampa(dominio, rampa or RAMPA_NEUTRA)
        # Las listas de la ficha van de mayor a menor ("Alto, Medio, Bajo") y
        # la rampa de claro a oscuro: sin invertir, el nivel alto saldria en
        # el tono mas palido y se leeria al reves.
        tonos = dict(zip(dominio, reversed(pasos) if descendente else pasos))
    vistos = {f["sub"] for f in filas_grupo}
    presentes = [s for s in dominio if s in vistos]
    presentes += sorted(vistos - set(dominio))
    return presentes, {s: tonos.get(s, SIN_DATO) for s in presentes}


def _divergente_centrada(filas_grupo, orden_sub, indice_neutro):
    """Convierte una escala ordenada en barras apiladas divergentes.

    Las clases favorables se apilan hacia la derecha, las desfavorables hacia
    la izquierda y la clase neutra queda a caballo del cero, que es lo que
    permite comparar de un vistazo el saldo de cada ambito. Devuelve las filas
    con los extremos `ini`/`fin` ya calculados en porcentaje.
    """
    salida = []
    grupos = {}
    for fila in filas_grupo:
        grupos.setdefault(fila["cat"], []).append(fila)
    for cat, filas in grupos.items():
        total = sum(f["valor"] for f in filas) or 1
        pct = {f["sub"]: 100.0 * f["valor"] / total for f in filas}
        crudo = {f["sub"]: f["valor"] for f in filas}
        neutro = pct.get(orden_sub[indice_neutro], 0.0)
        # Tramo desfavorable: hacia la izquierda desde el borde del neutro.
        cursor = -neutro / 2
        for sub in reversed(orden_sub[:indice_neutro]):
            ancho = pct.get(sub, 0.0)
            if ancho:
                salida.append({"cat": cat, "sub": sub, "valor": crudo.get(sub, 0),
                               "pct": ancho, "ini": cursor - ancho, "fin": cursor})
            cursor -= ancho
        if neutro:
            salida.append({"cat": cat, "sub": orden_sub[indice_neutro],
                           "valor": crudo.get(orden_sub[indice_neutro], 0),
                           "pct": neutro, "ini": -neutro / 2, "fin": neutro / 2})
        cursor = neutro / 2
        for sub in orden_sub[indice_neutro + 1:]:
            ancho = pct.get(sub, 0.0)
            if ancho:
                salida.append({"cat": cat, "sub": sub, "valor": crudo.get(sub, 0),
                               "pct": ancho, "ini": cursor, "fin": cursor + ancho})
            cursor += ancho
    return salida


def _colores_sino(tema="claro"):
    t = TINTAS[tema]
    return {"Sí": t["si"], "No": t["no"], "No aplica": t["sin_dato"],
            "Parcial": "#C4A03C"}


# ══════════════════════════════════════════════════════════════════════════
# GRAFICOS INTERACTIVOS (Altair / Vega-Lite)
# ══════════════════════════════════════════════════════════════════════════
# Altair viaja con Streamlit; se importa de forma diferida para que el resto
# del modulo (lectura de fichas y Excel) siga funcionando sin el.

def _alt():
    import altair as alt
    return alt


def _df(filas):
    import pandas as pd
    return pd.DataFrame(filas)


def _escala_valor(serie):
    """Rampa secuencial que corresponde al signo de la magnitud graficada."""
    return {"critica": RAMPA_CRITICA, "favorable": RAMPA_FAVORABLE,
            "neutra": RAMPA_NEUTRA}.get(serie.get("escala"), RAMPA_NEUTRA)


def _formato(serie):
    return f",.{serie.get('decimales', 0)}f"


def _eje_y(alt, serie, campo, orden):
    """Eje de categorias: sin regla ni marcas, y con aire antes de la barra."""
    return alt.Y(f"{campo}:N", sort=orden, title=None,
                 axis=alt.Axis(labelLimit=320, labelFontSize=11,
                               labelPadding=10, domain=False, ticks=False))


def grafico_altair(serie, tema="claro", altura=None):
    """Construye el grafico interactivo de una serie.

    Devuelve un objeto Altair listo para `st.altair_chart(..., use_container_width=True)`.
    Todos los graficos llevan tooltip; las barras llevan ademas el valor
    rotulado al extremo, que es lo que sostiene la lectura cuando el color
    por si solo no alcanza el contraste minimo sobre el fondo.
    """
    alt = _alt()
    t = TINTAS.get(tema, TINTAS["claro"])
    filas = serie.get("filas") or []
    if not filas:
        return None
    df, categorias, subclases = _datos_grafico(serie)
    if df.empty:
        return None
    forma = serie.get("forma", "barras_h")
    base = alt.Chart(df)
    if forma == "mapa_calor":
        grafico = _mapa_calor(alt, base, serie, t)
        alto = altura or 62 * max(len(serie.get("orden_cat") or categorias), 3) + 40
    elif forma in ("apiladas", "agrupadas"):
        grafico = _barras_multiples(alt, base, serie, t, forma, categorias,
                                    subclases)
        alto = altura or _alto_barras(len(categorias),
                                      agrupadas=(forma == "agrupadas"),
                                      n_sub=len(subclases))
    else:
        grafico = _barras_simples(alt, base, df, serie, t, categorias)
        alto = altura or _alto_barras(len(categorias))

    return (grafico
            # "fit-x" ajusta solo el ancho al contenedor. Con el "fit" que
            # Streamlit aplica por defecto, la leyenda inferior se descuenta
            # tambien del alto y las barras apiladas quedan sin dibujar.
            .properties(autosize=alt.AutoSizeParams(type="fit-x",
                                                    contains="padding"),
                        height=alto, title=alt.TitleParams(
                text=serie.get("titulo", ""),
                subtitle=_envolver(serie.get("descripcion", "")),
                anchor="start", fontSize=14, font=FUENTE, color=t["tinta"],
                subtitleFontSize=11, subtitleFont=FUENTE,
                subtitleColor=t["tinta_2"], offset=10))
            .configure_view(stroke=None)
            .configure_axis(grid=True, gridColor=t["rejilla"], gridWidth=1,
                            domainColor=t["rejilla"], tickColor=t["rejilla"],
                            labelColor=t["tinta_2"], titleColor=t["tinta_2"],
                            labelFont=FUENTE, titleFont=FUENTE,
                            titleFontSize=11, titleFontWeight="normal")
            .configure_legend(labelColor=t["tinta_2"], titleColor=t["tinta_2"],
                              labelFont=FUENTE, titleFont=FUENTE,
                              labelFontSize=11, titleFontSize=11,
                              symbolType="square", symbolSize=110,
                              orient="bottom", direction="horizontal",
                              columns=3, labelLimit=280)
            .configure_title(font=FUENTE))


def _datos_grafico(serie):
    """DataFrame acumulado del grafico, con el orden de apilado resuelto.

    Devuelve tambien el orden de categorias y subclases para que el eje, la
    leyenda y el apilado usen exactamente el mismo que el Excel y el PDF.
    """
    import pandas as pd
    cat, val, sub = serie["cat"], serie["val"], serie.get("sub")
    acumulado, extras = _acumular(serie)
    categorias = _orden_categorias(serie, acumulado)
    subclases = []
    if sub:
        subclases = [x for x in (serie.get("orden_sub") or [])
                     if any(k[1] == x for k in acumulado)]
        subclases += sorted({k[1] for k in acumulado} - set(subclases))
    if serie.get("forma") == "mapa_calor":
        for c in (serie.get("orden_cat") or categorias):
            for sc in (serie.get("orden_sub") or subclases):
                acumulado.setdefault((c, sc), 0)
        categorias = serie.get("orden_cat") or categorias
    filas = []
    for (c, sc), valor in acumulado.items():
        fila = {cat: c, val: valor}
        if sub:
            # El indice explicito fija el orden de los tramos: sin el, Vega
            # los apila por orden alfabetico y el "No" puede quedar antes que
            # el "Si" en unas barras y despues en otras.
            fila[sub] = sc
            fila["_orden"] = subclases.index(sc) if sc in subclases else 99
        if extras.get((c, sc)):
            fila["detalle"] = extras[(c, sc)]
        filas.append(fila)
    df = pd.DataFrame(filas)
    if not df.empty:
        df = df.sort_values([cat] + (["_orden"] if sub else []))
    return df, categorias, subclases


def _alto_barras(n_categorias, agrupadas=False, n_sub=1):
    """Altura util: barras finas, con aire entre ellas."""
    paso = 30 if not agrupadas else max(30, 16 * max(n_sub, 1))
    return int(min(760, max(160, n_categorias * paso + 60)))


def _envolver(texto, ancho=118):
    """Parte la descripcion en lineas para el subtitulo del grafico."""
    if not texto:
        return ""
    palabras, lineas, actual = texto.split(), [], ""
    for palabra in palabras:
        if len(actual) + len(palabra) + 1 > ancho:
            lineas.append(actual)
            actual = palabra
        else:
            actual = f"{actual} {palabra}".strip()
    if actual:
        lineas.append(actual)
    return lineas[:3]


def _eje_x(alt, serie, escala, tope=0):
    """Eje de valores.

    Donde se cuentan fichas, actores o titulares las marcas tienen que caer en
    enteros: con el reparto automatico un eje que llega a 2 se rotula
    "0 1 1 2". Por debajo de una docena se fijan los enteros uno a uno y por
    encima basta con exigir el paso minimo de 1.
    """
    fmt = _formato(serie)
    ejes = {"format": fmt, "labelFontSize": 10, "labelFlush": False}
    if serie.get("decimales", 0) == 0:
        ejes["tickMinStep"] = 1
        if 0 < tope <= 12:
            ejes["values"] = list(range(0, int(math.ceil(tope)) + 1))
        else:
            # Un eje de poblacion llega a cuatro cifras: sin limitar el numero
            # de marcas, Vega reparte una cada 50 y los rotulos se amontonan.
            ejes["tickCount"] = 8
    return alt.X(f"{serie['val']}:Q", title=serie.get("eje_x") or None,
                 scale=escala, axis=alt.Axis(**ejes))


def _barras_simples(alt, base, df, serie, t, categorias):
    """Barras horizontales de una sola serie.

    El color codifica la propia magnitud (rampa de un solo tono) salvo que la
    serie traiga un color explicito por clase, que es el caso de las escalas
    ordenadas favorable <-> desfavorable.
    """
    cat, val = serie["cat"], serie["val"]
    unidad, fmt = serie.get("unidad", ""), _formato(serie)
    colores = serie.get("colores")
    if colores:
        dominio = [c for c in categorias if c in colores]
        color = alt.Color(f"{cat}:N", legend=None, scale=alt.Scale(
            domain=dominio, range=[colores[c] for c in dominio]))
    else:
        color = alt.Color(f"{val}:Q", legend=None, scale=alt.Scale(
            range=_escala_valor(serie), type="linear"))

    tope = float(df[val].max() or 0)
    escala_x = alt.Scale(domain=[0, serie.get("maximo") or max(tope * 1.16, 1)],
                         nice=False)
    eje_x = _eje_x(alt, serie, escala_x, tope)
    barras = base.mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4,
                           height=alt.RelativeBandSize(0.6)).encode(
        x=eje_x, y=_eje_y(alt, serie, cat, categorias), color=color,
        tooltip=[alt.Tooltip(f"{cat}:N", title="Clase"),
                 alt.Tooltip(f"{val}:Q", title=unidad or "Valor", format=fmt)])
    # El valor al extremo de la barra es lo que sostiene la lectura cuando el
    # tono no alcanza 3:1 de contraste contra el fondo.
    etiquetas = base.mark_text(align="left", dx=7, fontSize=11, font=FUENTE,
                               color=t["tinta_2"], fontWeight="bold").encode(
        x=alt.X(f"{val}:Q", scale=escala_x, title=None, axis=None),
        y=_eje_y(alt, serie, cat, categorias),
        text=alt.Text(f"{val}:Q", format=fmt))
    return (barras + etiquetas)


def _barras_multiples(alt, base, serie, t, forma, categorias, subclases):
    """Barras apiladas (parte-todo) o agrupadas (comparacion de dos medidas)."""
    cat, val, sub = serie["cat"], serie["val"], serie["sub"]
    unidad, fmt = serie.get("unidad", ""), _formato(serie)
    colores = serie.get("colores") or {}
    # La leyenda solo declara las clases que aparecen: una entrada sin tramo
    # en ninguna barra hace dudar de si el dato falta o vale cero.
    color = alt.Color(f"{sub}:N", title=None, sort=subclases,
                      scale=alt.Scale(domain=subclases,
                                      range=[colores.get(x, SIN_DATO)
                                             for x in subclases]))
    tooltip = [alt.Tooltip(f"{cat}:N", title=serie.get("eje_y") or "Ámbito"),
               alt.Tooltip(f"{sub}:N", title="Clase"),
               alt.Tooltip(f"{val}:Q", title=unidad or "Valor", format=fmt)]

    if forma == "apiladas":
        totales = base.data.groupby(cat)[val].sum()
        tope, escala_x = float(totales.max() or 0), alt.Scale(nice=False)
    else:
        tope, escala_x = float(base.data[val].max() or 0), alt.Scale()
    codificacion = {
        "x": _eje_x(alt, serie, escala_x, tope).copy(),
        "y": _eje_y(alt, serie, cat, categorias),
        "color": color,
        "tooltip": tooltip,
        "order": alt.Order("_orden:Q", sort="ascending"),
    }
    codificacion["x"]["stack"] = True if forma == "apiladas" else None
    if forma == "agrupadas":
        codificacion["yOffset"] = alt.YOffset(f"{sub}:N", sort=subclases)
    # El filete del color del fondo separa los tramos contiguos: sin el, dos
    # tramos de tonos vecinos se leen como uno solo.
    return base.mark_bar(stroke=t["fondo"], strokeWidth=2,
                         cornerRadiusTopRight=3, cornerRadiusBottomRight=3,
                         height=alt.RelativeBandSize(0.68)).encode(**codificacion)


def _mapa_calor(alt, base, serie, t):
    """Rejilla categoria x subcategoria con el conteo rotulado en cada celda.

    El dominio de ambos ejes es el declarado completo, no el observado: una
    matriz de influencia x interes tiene que mostrar los nueve cuadrantes,
    incluidos los vacios, porque el cuadrante sin actores tambien informa.
    """
    cat, val, sub = serie["cat"], serie["val"], serie.get("sub")
    orden_cat = serie.get("orden_cat") or sorted(base.data[cat].unique())
    orden_sub = serie.get("orden_sub") or sorted(base.data[sub].unique())
    tope = float(base.data[val].max() or 1)
    escala_color = alt.Scale(range=_escala_valor(serie), domain=[0, tope])
    eje_comun = dict(labelFontSize=11, domain=False, ticks=False, grid=False,
                     labelLimit=200)

    codificacion = {
        "x": alt.X(f"{sub}:N", sort=orden_sub, scale=alt.Scale(domain=orden_sub),
                   title=serie.get("eje_x") or None,
                   axis=alt.Axis(labelAngle=0, orient="top", **eje_comun)),
        "y": alt.Y(f"{cat}:N", sort=orden_cat, scale=alt.Scale(domain=orden_cat),
                   title=serie.get("eje_y") or None, axis=alt.Axis(**eje_comun)),
    }
    tooltip = [alt.Tooltip(f"{cat}:N", title=serie.get("eje_y") or "Fila"),
               alt.Tooltip(f"{sub}:N", title=serie.get("eje_x") or "Columna"),
               alt.Tooltip(f"{val}:Q", title=serie.get("unidad") or "Valor")]
    if "detalle" in base.data.columns:
        tooltip.append(alt.Tooltip("detalle:N", title="Actores"))

    celdas = base.mark_rect(stroke=t["fondo"], strokeWidth=4,
                            cornerRadius=6).encode(
        color=alt.condition(
            alt.datum[val] > 0,
            alt.Color(f"{val}:Q", title=serie.get("unidad") or None,
                      scale=escala_color,
                      legend=alt.Legend(orient="right", direction="vertical",
                                        gradientLength=120, format=",.0f",
                                        # Con pocos actores la escala es
                                        # degenerada: se rotulan los enteros
                                        # en vez de repetir "1 / 1 / 0".
                                        values=sorted({0, *range(1, int(tope) + 1)})
                                        if tope <= 6 else alt.Undefined,
                                        tickMinStep=1)),
            alt.value(t["rejilla"])),
        tooltip=tooltip, **codificacion)
    # El numero va sobre la celda; el rotulo pasa a blanco en la mitad oscura
    # de la rampa, donde la tinta secundaria dejaria de leerse.
    rotulos = (base.transform_filter(alt.datum[val] > 0)
               .mark_text(fontSize=15, fontWeight="bold", font=FUENTE)
               .encode(text=alt.Text(f"{val}:Q", format=",.0f"),
                       color=alt.condition(alt.datum[val] > tope / 2,
                                           alt.value("#FFFFFF"),
                                           alt.value(t["tinta"])),
                       **codificacion))
    return (celdas + rotulos).properties(
        width=alt.Step(150), height=alt.Step(58))


def tabla_serie(serie):
    """DataFrame de respaldo de una serie, con las mismas cifras del grafico.

    Es la vista tabular que acompana a cada grafico: sostiene la lectura de
    los tonos que no alcanzan 3:1 de contraste sobre el fondo y permite
    copiar los valores sin exportar el libro.
    """
    import pandas as pd
    filas = serie.get("filas") or []
    if not filas:
        return pd.DataFrame()
    cat, val, sub = serie["cat"], serie["val"], serie.get("sub")
    if serie.get("forma") == "mapa_calor":
        df = pd.DataFrame(filas).pivot_table(
            index=cat, columns=sub, values=val, aggfunc="sum", fill_value=0)
        orden_cat = [c for c in (serie.get("orden_cat") or []) if c in df.index]
        orden_sub = [c for c in (serie.get("orden_sub") or []) if c in df.columns]
        if orden_cat:
            df = df.reindex(orden_cat + [i for i in df.index if i not in orden_cat])
        if orden_sub:
            df = df[orden_sub + [c for c in df.columns if c not in orden_sub]]
        return df.reset_index()
    if sub:
        df = pd.DataFrame(filas).pivot_table(
            index=cat, columns=sub, values=val, aggfunc="sum", fill_value=0)
        orden_sub = [c for c in (serie.get("orden_sub") or []) if c in df.columns]
        df = df[orden_sub + [c for c in df.columns if c not in orden_sub]]
        df["Total"] = df.sum(axis=1)
        return df.sort_values("Total", ascending=False).reset_index()
    etiqueta = serie.get("unidad") or "Valor"
    df = pd.DataFrame([{"Clase": f[cat], etiqueta: f[val]} for f in filas])
    total = df[etiqueta].sum()
    if total:
        df["% del total"] = (100.0 * df[etiqueta] / total).round(1)
    return df


# ══════════════════════════════════════════════════════════════════════════
# LIBRO EXCEL CON GRAFICOS NATIVOS
# ══════════════════════════════════════════════════════════════════════════
# Los graficos se generan como objetos nativos de Excel (no como imagenes):
# el usuario puede reordenarlos, cambiar la escala o copiarlos a Word sin
# perder calidad, y las series quedan a la vista en la misma hoja.

_BORDE = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
_ALTO_GRAFICO_FILAS = 16          # filas que ocupa un grafico de 8 cm


def _hex(color):
    """Normaliza '#1B4D2E' a '1B4D2E', que es lo que espera openpyxl."""
    return str(color or "").lstrip("#").upper() or "808080"


def _titulo_hoja(ws, titulo, ancho=10):
    """Encabezado institucional ANIN al inicio de la hoja."""
    fila = 1
    for texto in ENCABEZADOS_ANIN:
        celda = ws.cell(fila, 1, texto)
        celda.font = Font(name="Arial", size=9, bold=True, color=ANIN_VERDE)
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ancho)
        fila += 1
    celda = ws.cell(fila, 1, SUBTITULO_PROYECTO)
    celda.font = Font(name="Arial", size=8, italic=True)
    celda.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ancho)
    ws.row_dimensions[fila].height = 24
    fila += 1
    celda = ws.cell(fila, 1, titulo)
    celda.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    celda.fill = PatternFill("solid", fgColor=ANIN_VERDE)
    celda.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ancho)
    ws.row_dimensions[fila].height = 22
    return fila + 2


def _escribir_tabla(ws, fila, titulo, cabeceras, filas, subtitulo=""):
    """Tabla con estilo ANIN. Devuelve (fila_cabecera, fila_siguiente)."""
    if titulo:
        celda = ws.cell(fila, 1, titulo)
        celda.font = Font(name="Arial", size=10, bold=True, color=ANIN_AZUL)
        fila += 1
    if subtitulo:
        celda = ws.cell(fila, 1, subtitulo)
        celda.font = Font(name="Arial", size=8, italic=True, color="595959")
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila,
                       end_column=max(len(cabeceras), 4))
        fila += 1
    fila_cab = fila
    for i, texto in enumerate(cabeceras, start=1):
        celda = ws.cell(fila, i, texto)
        celda.font = Font(name="Arial", size=9, bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=ANIN_VERDE)
        celda.alignment = Alignment(horizontal="center", wrap_text=True,
                                    vertical="center")
        celda.border = _BORDE
    fila += 1
    for j, registro in enumerate(filas):
        for i, valor in enumerate(registro, start=1):
            celda = ws.cell(fila, i, valor)
            celda.font = Font(name="Arial", size=9)
            celda.border = _BORDE
            celda.alignment = Alignment(wrap_text=isinstance(valor, str)
                                        and len(str(valor)) > 40,
                                        vertical="top")
            if j % 2:
                celda.fill = PatternFill("solid", fgColor=ANIN_GRIS)
            if isinstance(valor, float):
                celda.number_format = "#,##0.00"
            elif isinstance(valor, int):
                celda.number_format = "#,##0"
        fila += 1
    return fila_cab, fila


def _acumular(serie):
    """Suma las filas repetidas de una misma (categoria, subclase).

    Varios centros poblados declaran la misma actividad o el mismo peligro: sin
    este paso el apilado dibujaria dos tramos contiguos del mismo color en vez
    de uno solo con el total.
    """
    cat, val, sub = serie["cat"], serie["val"], serie.get("sub")
    acumulado, extras = {}, {}
    for f in serie.get("filas") or []:
        clave = (f[cat], f[sub]) if sub else (f[cat], None)
        acumulado[clave] = acumulado.get(clave, 0) + (f[val] or 0)
        # El tooltip del mapa de calor lista los actores de cada celda.
        if f.get("detalle") and clave not in extras:
            extras[clave] = f["detalle"]
    return acumulado, extras


def _orden_categorias(serie, acumulado):
    """Orden del eje de categorias, comun a los tres formatos de salida."""
    declarado = serie.get("orden_cat")
    presentes = list(dict.fromkeys(c for c, _ in acumulado))
    if declarado:
        orden = [c for c in declarado if c in set(presentes)]
        return orden + [c for c in presentes if c not in set(orden)]
    if serie.get("escala") == "sino" and serie.get("orden_sub"):
        # Bateria de preguntas: manda el numero de respuestas afirmativas, que
        # es la lectura util (que capacidades existen y cuales faltan).
        favorable = serie["orden_sub"][0]
        return sorted(presentes,
                      key=lambda c: (-acumulado.get((c, favorable), 0), c))
    totales = {}
    for (c, _s), v in acumulado.items():
        totales[c] = totales.get(c, 0) + v
    return sorted(presentes, key=lambda c: (-totales[c], str(c)))


def _pivote(serie):
    """Convierte una serie en (categorias, subclases, matriz de valores)."""
    sub = serie.get("sub")
    acumulado, _extras = _acumular(serie)
    categorias = _orden_categorias(serie, acumulado)
    if not sub:
        return categorias, [serie.get("unidad") or "Valor"], \
            [[acumulado.get((c, None), 0)] for c in categorias]
    subclases = [s for s in (serie.get("orden_sub") or [])
                 if any(k[1] == s for k in acumulado)]
    subclases += sorted({k[1] for k in acumulado} - set(subclases))
    matriz = [[acumulado.get((c, s), 0) for s in subclases] for c in categorias]
    return categorias, subclases, matriz


def _colores_serie(serie, subclases, categorias):
    """Color de cada serie (o de cada punto) del grafico de Excel."""
    colores = serie.get("colores") or {}
    if serie.get("sub"):
        return [_hex(colores.get(s) or _rampa(subclases, _escala_valor(serie))[i])
                for i, s in enumerate(subclases)]
    if colores:
        return [_hex(colores.get(c, SIN_DATO)) for c in categorias]
    # Sin color por clase, la propia magnitud ordena la rampa: la barra mas
    # larga recibe el paso mas oscuro.
    rampa = _escala_valor(serie)
    filas = serie.get("filas") or []
    valores = {f[serie["cat"]]: f[serie["val"]] for f in filas}
    orden = sorted(categorias, key=lambda c: valores.get(c, 0))
    paso = {c: rampa[min(len(rampa) - 1,
                         int(i * len(rampa) / max(len(orden), 1)))]
            for i, c in enumerate(orden)}
    return [_hex(paso.get(c, rampa[2])) for c in categorias]


def _grafico_de_serie(ws, serie, fila_cab, categorias, subclases, colores,
                      ancla):
    """Agrega a la hoja el grafico nativo que corresponde a la serie."""
    n_cat, n_sub = len(categorias), len(subclases)
    if not n_cat or not n_sub:
        return
    graf = BarChart()
    graf.type = "bar"                      # barras horizontales, como en la app
    graf.style = 2
    graf.title = serie.get("titulo", "")
    graf.y_axis.title = serie.get("unidad") or serie.get("eje_x") or ""
    graf.x_axis.title = ""
    graf.gapWidth = 60
    graf.height = 8.5
    graf.width = 17 if n_sub > 1 else 15
    graf.grouping = "stacked" if serie.get("forma") == "apiladas" else "clustered"
    if graf.grouping == "stacked":
        graf.overlap = 100

    datos = Reference(ws, min_col=2, max_col=1 + n_sub,
                      min_row=fila_cab, max_row=fila_cab + n_cat)
    cats = Reference(ws, min_col=1, min_row=fila_cab + 1,
                     max_row=fila_cab + n_cat)
    graf.add_data(datos, titles_from_data=True)
    graf.set_categories(cats)

    if n_sub > 1:
        for i, s in enumerate(graf.series):
            s.graphicalProperties = GraphicalProperties(solidFill=colores[i])
            s.graphicalProperties.line.solidFill = "FFFFFF"
            s.graphicalProperties.line.width = 12700        # 1 pt
        graf.legend.position = "b"
    else:
        graf.legend = None
        serie_excel = graf.series[0]
        serie_excel.graphicalProperties = GraphicalProperties(
            solidFill=colores[0] if colores else _hex(RAMPA_NEUTRA[3]))
        # Un color por barra: en Excel eso se expresa como puntos de datos.
        serie_excel.data_points = [
            DataPoint(idx=i, spPr=GraphicalProperties(solidFill=color))
            for i, color in enumerate(colores)]
        serie_excel.dLbls = DataLabelList()
        serie_excel.dLbls.showVal = True
        serie_excel.dLbls.showSerName = False
        serie_excel.dLbls.showCatName = False
        serie_excel.dLbls.showLegendKey = False
    ws.add_chart(graf, ancla)


def _hoja_seccion(wb, seccion, usados):
    """Una hoja por seccion: cada serie con su tabla y su grafico al lado."""
    ws = wb.create_sheet(_nombre_hoja(seccion["titulo"], usados))
    fila = _titulo_hoja(ws, seccion["titulo"].upper(), ancho=10)
    for col, ancho in zip("ABCDEFGHIJ", (46, 15, 15, 15, 15, 15, 15, 15, 15, 15)):
        ws.column_dimensions[col].width = ancho
    if seccion.get("descripcion"):
        celda = ws.cell(fila, 1, seccion["descripcion"])
        celda.font = Font(name="Arial", size=9, italic=True, color="595959")
        fila += 2

    for i, serie in enumerate(seccion.get("series", []), start=1):
        categorias, subclases, matriz = _pivote(serie)
        if not categorias:
            continue
        colores = _colores_serie(serie, subclases, categorias)
        cabeceras = [serie.get("eje_y") or "Clase"] + list(subclases)
        filas = [[c] + [_celda_num(v) for v in valores]
                 for c, valores in zip(categorias, matriz)]
        fila_cab, fila_fin = _escribir_tabla(
            ws, fila, f"{chr(64 + i)}. {serie['titulo']}", cabeceras, filas,
            subtitulo=serie.get("descripcion", ""))
        # Totales por columna, como formula: el usuario puede filtrar y ver
        # el recalculo, que es lo que se pierde si se graban valores fijos.
        ws.cell(fila_fin, 1, "TOTAL").font = Font(name="Arial", size=9, bold=True)
        for j in range(len(subclases)):
            letra = get_column_letter(2 + j)
            celda = ws.cell(fila_fin, 2 + j,
                            f"=SUM({letra}{fila_cab + 1}:{letra}{fila_fin - 1})")
            celda.font = Font(name="Arial", size=9, bold=True)
            celda.number_format = "#,##0.00"
        ancla = f"{get_column_letter(3 + len(subclases))}{fila_cab}"
        _grafico_de_serie(ws, serie, fila_cab, categorias, subclases, colores,
                          ancla)
        if serie.get("nota"):
            celda = ws.cell(fila_fin + 1, 1, "Fuente: " + serie["nota"])
            celda.font = Font(name="Arial", size=8, italic=True, color="595959")
            ws.merge_cells(start_row=fila_fin + 1, start_column=1,
                           end_row=fila_fin + 1, end_column=8)
        fila = max(fila_fin + 3,
                   fila_cab + _ALTO_GRAFICO_FILAS + 2)
    ws.sheet_view.showGridLines = False
    return ws


def _celda_num(valor):
    """Entero cuando el valor es entero: evita '12.00' donde se cuentan fichas."""
    if isinstance(valor, float) and valor.is_integer():
        return int(valor)
    return valor


def _nombre_hoja(nombre, usados):
    """Nombre de hoja valido y unico (Excel: 31 caracteres, sin : \\ / ? * [ ])."""
    limpio = re.sub(r"[\\/?*\[\]:]", " ", str(nombre)).strip() or "Hoja"
    limpio = limpio.replace("·", "-")[:31].strip()
    base, i = limpio, 2
    while limpio.lower() in usados:
        sufijo = f" {i}"
        limpio = base[:31 - len(sufijo)] + sufijo
        i += 1
    usados.add(limpio.lower())
    return limpio


def _hoja_tablas(wb, seccion, usados):
    """Hojas de respaldo con el detalle fila a fila de las fichas."""
    for titulo, filas in seccion.get("tablas", []):
        if not filas:
            continue
        ws = wb.create_sheet(_nombre_hoja(f"T {titulo}", usados))
        fila = _titulo_hoja(ws, titulo.upper(), ancho=min(len(filas[0]), 12) or 6)
        cabeceras = list(filas[0].keys())
        for i, cab in enumerate(cabeceras, start=1):
            ws.column_dimensions[get_column_letter(i)].width = \
                min(max(len(str(cab)) + 4, 14), 46)
        fila_cab, _ = _escribir_tabla(
            ws, fila, "", cabeceras,
            [[_celda_num(f.get(c)) if isinstance(f.get(c), (int, float))
              else _txt(f.get(c)) for c in cabeceras] for f in filas])
        ws.freeze_panes = ws.cell(fila_cab + 1, 1)
        ws.auto_filter.ref = (f"A{fila_cab}:"
                              f"{get_column_letter(len(cabeceras))}"
                              f"{fila_cab + len(filas)}")
        ws.sheet_view.showGridLines = False


# ══════════════════════════════════════════════════════════════════════════
# ANEXO GRAFICO EN PDF
# ══════════════════════════════════════════════════════════════════════════
# Reutiliza el PDF institucional de los resumenes territoriales (cabecera,
# pie, tablas y graficos vectoriales) para que la ficha social y la
# territorial salgan con la misma factura.

def _rgb(color):
    """'#2E7D4F' -> (46, 125, 79)."""
    texto = str(color or "").lstrip("#")
    if len(texto) != 6:
        return (120, 120, 120)
    return tuple(int(texto[i:i + 2], 16) for i in (0, 2, 4))


def _recortar(texto, largo):
    texto = str(texto or "")
    return texto if len(texto) <= largo else texto[:largo - 1] + "…"


def _pdf_barras_apiladas(pdf, serie, alto_fila=6.2):
    """Barras horizontales apiladas dibujadas con primitivas de fpdf2.

    El PDF institucional de los resumenes solo trae barras simples y tortas;
    las escalas ordenadas del diagnostico social necesitan el apilado para
    que se lea el reparto dentro de cada ambito.
    """
    from resumenes_bloques import _s
    categorias, subclases, matriz = _pivote(serie)
    if not categorias or not subclases:
        return
    colores = serie.get("colores") or {}
    rampa = _rampa(subclases, _escala_valor(serie))
    tintas = [_rgb(colores.get(s) or rampa[i]) for i, s in enumerate(subclases)]

    filas_visibles = categorias[:18]
    alto_total = len(filas_visibles) * alto_fila + 12
    pdf._salto_si_falta(alto_total + 12)
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(0, 5, _s(serie.get("titulo", "")), 0, 1, "L")
    if serie.get("descripcion"):
        pdf.set_font("Helvetica", "I", 6.5)
        pdf.set_text_color(80, 80, 80)
        pdf.multi_cell(0, 3.2, _s(serie["descripcion"]))
        pdf.set_text_color(0, 0, 0)

    x0, y0 = 12.0, pdf.get_y() + 1
    ancho = pdf.w - 24
    ancho_etq = ancho * 0.36
    ancho_barra = ancho - ancho_etq - 16
    maximo = max((sum(f) for f in matriz), default=0) or 1.0

    for i, categoria in enumerate(filas_visibles):
        valores = matriz[categorias.index(categoria)]
        y = y0 + i * alto_fila
        pdf.set_xy(x0, y)
        pdf.set_font("Helvetica", "", 6.3)
        pdf.cell(ancho_etq, alto_fila, _s(_recortar(categoria, 46)), 0, 0, "L")
        x = x0 + ancho_etq
        for valor, tinta in zip(valores, tintas):
            if not valor:
                continue
            largo = ancho_barra * valor / maximo
            pdf.set_fill_color(*tinta)
            pdf.rect(x, y + alto_fila * 0.18, max(largo, 0.3),
                     alto_fila * 0.64, "F")
            x += largo
        pdf.set_xy(x + 1.2, y)
        pdf.set_font("Helvetica", "B", 6)
        pdf.cell(14, alto_fila, _s(_fmt_pdf(sum(valores))), 0, 0, "L")

    y = y0 + len(filas_visibles) * alto_fila + 1.5
    pdf.set_y(y)
    # Leyenda: sin ella el color no identifica nada, y los pares verde/rojo
    # del diagnostico necesitan ese segundo canal.
    pdf.set_font("Helvetica", "", 6.3)
    x = x0
    for subclase, tinta in zip(subclases, tintas):
        etiqueta = _recortar(subclase, 34)
        ancho_item = pdf.get_string_width(_s(etiqueta)) + 8
        if x + ancho_item > pdf.w - 12:
            x = x0
            y += 4.2
            pdf.set_y(y)
        pdf.set_fill_color(*tinta)
        pdf.rect(x, y + 1.0, 2.8, 2.8, "F")
        pdf.set_xy(x + 4, y)
        pdf.cell(ancho_item - 4, 4.6, _s(etiqueta), 0, 0, "L")
        x += ancho_item
    pdf.set_y(y + 6)
    if len(categorias) > len(filas_visibles):
        pdf.nota(f"Se grafican los {len(filas_visibles)} ámbitos con mayor "
                 f"registro de un total de {len(categorias)}; el detalle "
                 f"completo está en el libro Excel.")
    if serie.get("nota"):
        pdf.nota("Fuente: " + serie["nota"])


def _fmt_pdf(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and not valor.is_integer():
        return f"{valor:,.2f}".replace(",", " ")
    return f"{int(valor):,}".replace(",", " ")


def _pdf_barras_simples(pdf, serie):
    """Barras horizontales de una serie simple, con su color por clase."""
    categorias, _sub, matriz = _pivote(serie)
    if not categorias:
        return
    colores = _colores_serie(serie, [serie.get("unidad") or "Valor"], categorias)
    mapa = {str(c).strip().upper(): _rgb("#" + colores[i])
            for i, c in enumerate(categorias)}
    valores = [fila[0] for fila in matriz]
    pdf.set_font("Helvetica", "I", 6.5)
    pdf.grafico_barras(serie.get("titulo", ""), categorias, valores,
                       unidad=(" " + serie["unidad"]) if serie.get("unidad") else "",
                       alto=min(96, max(24, 6.4 * len(categorias))),
                       horizontal=True, colores=mapa)
    if serie.get("descripcion"):
        pdf.nota(serie["descripcion"])
    if serie.get("nota"):
        pdf.nota("Fuente: " + serie["nota"])
