"""Construye el Diagnóstico Integral IN Piura V1 en la carpeta publicar/.

  publicar/index.html                     portada integral (incrusta datos/data.json)
  publicar/vol1-morropon/index.html       Volumen I con la navegación integral
  publicar/vol2-huancabamba/index.html    Volumen II
  publicar/vol3-ayabaca/index.html        Volumen III

Las láminas de cada volumen (eco/, mapas/ o maps/, sp/) no están en el
repositorio: viven en el artefacto publicado. Ver README.md.

Uso:  python construir.py
"""
import os
import re

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, 'src')
OUT = os.path.join(BASE, 'publicar')
VOLS = [('mor', 'vol1-morropon', 'I', 'Morropón', 'vol1_morropon_v5.html'),
        ('hua', 'vol2-huancabamba', 'II', 'Huancabamba', 'vol2_huancabamba_v5.html'),
        ('aya', 'vol3-ayabaca', 'III', 'Ayabaca', 'vol3_ayabaca_v5.html')]


def leer(*p):
    return open(os.path.join(*p), encoding='utf8').read()


def enlaces(actual, pre):
    """Barra de volúmenes común: Integral · Vol. I · Vol. II · Vol. III · Metrados y costos."""
    items = [('Integral', pre + 'index.html', actual == 'int')]
    items += [(f'Vol. {v} · {n}', f'{pre}{d}/index.html', actual == k) for k, d, v, n, _ in VOLS]
    h = ''.join(f'<a href="{u}"' + (' aria-current="page"' if c else '') + f'>{t}</a>' for t, u, c in items)
    return h + f'<a href="{pre}index.html#costos" class="cst">Metrados y costos</a>'


def main():
    os.makedirs(OUT, exist_ok=True)
    datos = leer(BASE, 'datos', 'data.json').replace('</', '<\\/')
    motor = leer(SRC, 'engine.js').replace("if(typeof module!=='undefined') module.exports=cpEngine;", '')
    html = (leer(SRC, 'index.tpl.html')
            .replace('/*__CSS__*/', leer(SRC, 'base.css'))
            .replace('/*__DATA__*/', datos)
            .replace('/*__ENGINE__*/', motor)
            .replace('/*__APP__*/', leer(SRC, 'app.js'))
            .replace('/*__VOLS__*/', enlaces('int', ''))
            .replace('/*__FVOLS__*/', enlaces('int', '')))
    html = '<meta name="claude:viewport-fit-sides" content="cover">\n' + html
    open(os.path.join(OUT, 'index.html'), 'w', encoding='utf8').write(html)
    print('index.html', len(html))

    for k, carpeta, v, n, fuente in VOLS:
        s = leer(BASE, 'fuentes', fuente)
        s, c1 = re.subn(r'<nav class="vols"[^>]*>.*?</nav>',
                        '<nav class="vols" aria-label="Diagnóstico Integral IN Piura">' + enlaces(k, '../') + '</nav>',
                        s, count=1, flags=re.S)
        s, c2 = re.subn(r'<div class="fvols">.*?</div>',
                        '<div class="fvols"><span>Diagnóstico Integral IN Piura V1</span>' + enlaces(k, '../') + '</div>',
                        s, count=1, flags=re.S)
        s, c3 = re.subn(r'<title>Diagnóstico [^<]*</title>',
                        f'<title>Vol. {v} {n} · Diagnóstico Integral IN Piura V1</title>', s, count=1)
        if (c1, c2, c3) != (1, 1, 1) or 'claude.ai/artifact' in s:
            raise SystemExit(f'{carpeta}: no se pudo reemplazar la navegación ({c1}, {c2}, {c3})')
        s = s.replace('<header class="hero">', '<style>.vols a.cst{border-style:dashed}</style>\n<header class="hero">', 1)
        os.makedirs(os.path.join(OUT, carpeta), exist_ok=True)
        open(os.path.join(OUT, carpeta, 'index.html'), 'w', encoding='utf8').write(s)
        print(f'{carpeta}/index.html', len(s))


if __name__ == '__main__':
    main()
