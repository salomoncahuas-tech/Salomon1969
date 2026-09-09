# -*- coding: utf-8 -*-
"""Genera el organigrama del equipo de estudio del Entregable 3 - Estudio de Geologia.

Proyecto IN Piura (CUI 2669244) - ANIN / DIME / SESDI.
Identidad institucional: verde ANIN #1B4D2E, azul institucional #1B4F72, fuente Arial
(se usa Liberation Sans, metricamente compatible, cuando Arial no esta instalada).

Uso:
    python3 organigrama.py [salida.png]
"""
import os
import sys

from PIL import Image, ImageDraw, ImageFont

# ------------------------------------------------------------------ identidad
VERDE_ANIN = (27, 77, 46)        # #1B4D2E
VERDE_MEDIO = (46, 107, 69)      # #2E6B45
VERDE_CLARO = (232, 239, 233)    # #E8EFE9
AZUL_ANIN = (27, 79, 114)        # #1B4F72
GRIS_TEXTO = (60, 60, 60)
GRIS_LINEA = (127, 158, 138)     # #7F9E8A
BLANCO = (255, 255, 255)
FONDO = (255, 255, 255)

ESCALA = 2                        # supersampling para suavizar trazos y texto
ANCHO, ALTO = 1500, 756

RUTAS_FUENTE = [
    "/usr/share/fonts/truetype/msttcorefonts/Arial{suf}.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans{suf}.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans{suf2}.ttf",
]


def fuente(tam, negrita=False, cursiva=False):
    """Devuelve Arial si existe; si no, la alternativa metricamente compatible."""
    if negrita and cursiva:
        suf, suf2 = "-BoldItalic", "-BoldOblique"
    elif negrita:
        suf, suf2 = "-Bold", "-Bold"
    elif cursiva:
        suf, suf2 = "-Italic", "-Oblique"
    else:
        suf, suf2 = "-Regular", ""
    for plantilla in RUTAS_FUENTE:
        ruta = plantilla.format(suf=suf, suf2=suf2)
        if os.path.exists(ruta):
            return ImageFont.truetype(ruta, tam * ESCALA)
        # Arial usa nombres sin guion: Arial.ttf / Arial_Bold.ttf
        alt = plantilla.format(suf=suf.replace("-Regular", "").replace("-", "_"), suf2=suf2)
        if os.path.exists(alt):
            return ImageFont.truetype(alt, tam * ESCALA)
    return ImageFont.load_default()


# ------------------------------------------------------------------ contenido
COORDINACION = {
    "nombre": "Ing. Alejandro Falconi Valdivia",
    "cargo": "Coordinador de Proyectos IN",
    "rol": "DIME – SESDI  |  Gerente de Proyecto",
}

CONDUCCION = {
    "nombre": "Ing. Héctor Salomón Cahuas Miller",
    "cargo": "Responsable del Proyecto IN Piura",
    "rol": "DIME – SESDI  |  Gerente de Diseño",
}

EJECUCION = [
    {
        "nombre": "Ing. Albert Flores Mendoza",
        "cargo": "Analista SIG Proyectos IN",
        "area": "DIME – SESDI",
        "rol": "Especialista SIG-CAD",
        "clave": False,
    },
    {
        "nombre": "Ing. Christian Camacho Aponte",
        "cargo": "Asistente en SIG Proyectos IN",
        "area": "DIME – SESDI",
        "rol": "Apoyo SIG-CAD",
        "clave": False,
    },
    {
        "nombre": "Ing. Pedro Talledo Hernández",
        "cargo": "Especialista IM Proyectos IN",
        "area": "DIME – SESDI",
        "rol": "Especialista en Geología",
        "clave": True,
    },
]

ARTICULACION = {
    "nombre": "Eco. Luis García Iguía",
    "cargo": "Formulador de Proyectos IN",
    "area": "DIME – SESDI",
    "rol": "Articulación con el E10",
}

BANDAS = [
    "NIVEL DE COORDINACIÓN",
    "CONDUCCIÓN TÉCNICA",
    "EJECUCIÓN TÉCNICA DEL ENTREGABLE",
]

NOTA = ("Modalidad de ejecución: Administración Directa ANIN – DIME – SESDI. Los roles del "
        "flujograma del Entregable 3 (Especialista en Geología, Especialista SIG-CAD, "
        "Especialista Marrón, Gerente de Diseño y Gerente de Proyecto) son asumidos por el "
        "personal profesional de la SESDI.")


# ------------------------------------------------------------------ primitivas
def e(v):
    """Escala una coordenada al lienzo de supersampling."""
    return int(v * ESCALA)


def texto_centrado(d, cx, y, txt, fnt, color):
    ancho = d.textlength(txt, font=fnt)
    d.text((cx * ESCALA - ancho / 2, e(y)), txt, font=fnt, fill=color)


def caja(d, x, y, w, h, relleno, borde, grosor=2, radio=8):
    d.rounded_rectangle([e(x), e(y), e(x + w), e(y + h)], radius=e(radio),
                        fill=relleno, outline=borde, width=e(grosor))


def linea(d, x1, y1, x2, y2, color=VERDE_MEDIO, grosor=2):
    d.line([e(x1), e(y1), e(x2), e(y2)], fill=color, width=e(grosor))


def linea_punteada(d, x1, y1, x2, y2, color=AZUL_ANIN, grosor=2, tramo=8, hueco=6):
    """Traza una linea horizontal o vertical discontinua."""
    if y1 == y2:
        x = x1
        while x < x2:
            d.line([e(x), e(y1), e(min(x + tramo, x2)), e(y2)], fill=color, width=e(grosor))
            x += tramo + hueco
    else:
        y = y1
        while y < y2:
            d.line([e(x1), e(y), e(x2), e(min(y + tramo, y2))], fill=color, width=e(grosor))
            y += tramo + hueco


def bloque_persona(d, x, y, w, h, datos, acento, clave=False):
    """Dibuja una caja de persona: franja de acento, nombre, cargo, area y rol."""
    caja(d, x, y, w, h, BLANCO, acento, grosor=2)
    # franja superior de acento
    d.rounded_rectangle([e(x), e(y), e(x + w), e(y + 26)], radius=e(8), fill=acento)
    d.rectangle([e(x), e(y + 16), e(x + w), e(y + 26)], fill=acento)

    cx = x + w / 2
    texto_centrado(d, cx, y + 6, datos["nombre"], fuente(11, negrita=True), BLANCO)
    ty = y + 36
    texto_centrado(d, cx, ty, datos["cargo"], fuente(10, negrita=True), VERDE_ANIN)
    if datos.get("area"):
        texto_centrado(d, cx, ty + 17, datos["area"], fuente(9), GRIS_TEXTO)
        ty += 17
    # rol del flujograma, destacado en azul institucional
    texto_centrado(d, cx, ty + 19, datos["rol"], fuente(9, negrita=True, cursiva=True), AZUL_ANIN)

    if clave:
        # marca de rol clave del Entregable 3
        d.ellipse([e(x + w - 20), e(y + h - 20), e(x + w - 8), e(y + h - 8)],
                  fill=VERDE_MEDIO)


# ------------------------------------------------------------------ dibujo
def construir():
    img = Image.new("RGB", (ANCHO * ESCALA, ALTO * ESCALA), FONDO)
    d = ImageDraw.Draw(img)

    # ---- encabezado institucional
    d.rectangle([0, 0, ANCHO * ESCALA, e(96)], fill=VERDE_ANIN)
    texto_centrado(d, ANCHO / 2, 14, "AUTORIDAD NACIONAL DE INFRAESTRUCTURA – ANIN",
                   fuente(14, negrita=True), BLANCO)
    texto_centrado(d, ANCHO / 2, 36,
                   "DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA – DIME",
                   fuente(11, negrita=True), VERDE_CLARO)
    texto_centrado(d, ANCHO / 2, 54, "SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN – SESDI",
                   fuente(11, negrita=True), VERDE_CLARO)
    texto_centrado(d, ANCHO / 2, 74,
                   "Proyecto IN Piura (CUI 2669244) · Entregable 3 – Estudio de Geología",
                   fuente(10, cursiva=True), VERDE_CLARO)

    texto_centrado(d, ANCHO / 2, 116, "ORGANIGRAMA DEL EQUIPO DE ESTUDIO",
                   fuente(16, negrita=True), VERDE_ANIN)
    linea(d, 520, 142, 980, 142, VERDE_MEDIO, 2)

    # ---- bandas de nivel (etiquetas verticales a la izquierda)
    x_banda, w_banda = 40, 34
    bandas_y = [(168, 268), (288, 400), (420, 552)]
    for (y0, y1), etiqueta in zip(bandas_y, BANDAS):
        alto_banda = y1 - y0
        caja(d, x_banda, y0, w_banda, alto_banda, VERDE_CLARO, GRIS_LINEA, grosor=1, radio=5)
        # el mayor cuerpo (7 a 10 pt) cuyo texto rotado cabe dentro de la banda
        for tam in (10, 9, 8, 7):
            fnt = fuente(tam, negrita=True)
            ancho_txt = d.textlength(etiqueta, font=fnt)
            if ancho_txt <= e(alto_banda - 10):
                break
        tira = Image.new("RGB", (int(ancho_txt) + e(6), e(w_banda)), VERDE_CLARO)
        ImageDraw.Draw(tira).text((e(3), e(9)), etiqueta, font=fnt, fill=VERDE_ANIN)
        tira = tira.rotate(90, expand=True)
        img.paste(tira, (e(x_banda), e(y0) + (e(alto_banda) - tira.height) // 2))

    # ---- nivel 1: coordinacion
    x1, w1 = 545, 410
    bloque_persona(d, x1, 175, w1, 86, COORDINACION, VERDE_ANIN)

    # ---- nivel 2: conduccion tecnica
    bloque_persona(d, x1, 295, w1, 86, CONDUCCION, VERDE_ANIN)
    # linea jerarquica coordinacion -> conduccion
    linea(d, 750, 261, 750, 295, VERDE_MEDIO, 3)

    # ---- nivel 3: ejecucion tecnica
    y3, h3 = 430, 116
    w3 = 300
    huecos = 34
    total = len(EJECUCION) * w3 + (len(EJECUCION) - 1) * huecos
    x0 = (ANCHO - total) / 2 - 60
    centros = []
    for i, persona in enumerate(EJECUCION):
        x = x0 + i * (w3 + huecos)
        bloque_persona(d, x, y3, w3, h3, persona, VERDE_MEDIO, clave=persona["clave"])
        centros.append(x + w3 / 2)

    # conector en peine desde la conduccion tecnica
    y_bus = 405
    linea(d, 750, 381, 750, y_bus, VERDE_MEDIO, 3)
    linea(d, centros[0], y_bus, centros[-1], y_bus, VERDE_MEDIO, 3)
    for cx in centros:
        linea(d, cx, y_bus, cx, y3, VERDE_MEDIO, 3)

    # ---- articulacion con el E10 (linea punteada, dependencia funcional)
    xa, wa = 1130, 290
    bloque_persona(d, xa, 295, wa, 100, ARTICULACION, AZUL_ANIN)
    linea_punteada(d, x1 + w1, 338, xa, 338, AZUL_ANIN, 2)

    # ---- leyenda
    ly = 580
    caja(d, 545, ly, 410, 96, (250, 251, 250), GRIS_LINEA, grosor=1, radio=6)
    texto_centrado(d, 750, ly + 8, "LEYENDA", fuente(9, negrita=True), VERDE_ANIN)
    fnt_l = fuente(9)
    filas = [
        ("linea", VERDE_MEDIO, "Línea de dependencia jerárquica"),
        ("punteada", AZUL_ANIN, "Articulación con el Entregable 10"),
        ("punto", VERDE_MEDIO, "Rol clave del Entregable 3 – Estudio de Geología"),
    ]
    for i, (tipo, color, etiqueta) in enumerate(filas):
        y = ly + 30 + i * 20
        if tipo == "linea":
            linea(d, 562, y + 5, 596, y + 5, color, 3)
        elif tipo == "punteada":
            linea_punteada(d, 562, y + 5, 596, y + 5, color, 2, tramo=6, hueco=4)
        else:
            d.ellipse([e(573), e(y), e(585), e(y + 12)], fill=color)
        d.text((e(608), e(y - 1)), etiqueta, font=fnt_l, fill=GRIS_TEXTO)

    # ---- nota al pie
    ny = 690
    caja(d, 40, ny, ANCHO - 80, 46, VERDE_CLARO, GRIS_LINEA, grosor=1, radio=6)
    fnt_n = fuente(9, cursiva=True)
    palabras, lineas, actual = NOTA.split(), [], ""
    limite = (ANCHO - 130) * ESCALA
    for p in palabras:
        prueba = (actual + " " + p).strip()
        if d.textlength(prueba, font=fnt_n) > limite:
            lineas.append(actual)
            actual = p
        else:
            actual = prueba
    lineas.append(actual)
    for i, ln in enumerate(lineas):
        d.text((e(64), e(ny + 14 + i * 16)), ln, font=fnt_n, fill=VERDE_ANIN)

    # La leyenda "Figura 6.1." no se incrusta en la imagen: el informe la agrega como
    # parrafo debajo (ver LEYENDA_FIGURA en insertar_numeral6.py), y duplicarla sobraria.

    # borde exterior
    d.rectangle([0, 0, ANCHO * ESCALA - 1, ALTO * ESCALA - 1], outline=GRIS_LINEA, width=e(1))

    return img.resize((ANCHO, ALTO), Image.LANCZOS)


def main():
    salida = sys.argv[1] if len(sys.argv) > 1 else "Organigrama_E3_Geologia.png"
    img = construir()
    # paleta reducida: colores planos, archivo compacto y sin perdida visible
    img.convert("P", palette=Image.ADAPTIVE, colors=128).save(salida, optimize=True)
    print(f"OK -> {salida} ({os.path.getsize(salida):,} bytes)")


if __name__ == "__main__":
    main()
