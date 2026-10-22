"""Ensambla los Volúmenes I–III: capa social 17/10, pestañas Ecosistemas (E6) y Tamizaje predial, integración."""
import json, re, os, sys
sys.path.insert(0, 'bin')
from nov import NOV

PM = {'A': 'Ayabaca', 'H': 'Huancabamba', 'M': 'Morropón'}
VOL = {'M': 'I', 'H': 'II', 'A': 'III'}
COPY = {'M': 'https://claude.ai/artifact/Y5QRXCQRsXFvwFgYL1oZ64', 'H': 'https://claude.ai/artifact/9828ogNDewfGQLgQDCq7kn',
        'A': 'https://claude.ai/artifact/4kF1tRzj2FkuDgbcuTcBye'}
ORIG = {'M': 'https://claude.ai/artifact/QSAZwL5aCbHUXFuzgArxQC', 'H': 'https://claude.ai/artifact/2MR2LbK1kpBxo45QMJKuat',
        'A': 'https://claude.ai/artifact/CPGxFSAkDpK3EcSviwTxLB'}
E6 = json.load(open('work/E6_D.json'))['b']
SPDIR = ['sp', 'sp2']

def jdump(o):
    return json.dumps(o, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')

def replace_json(s, var, obj):
    key = f'window.{var}='
    i = s.index(key) + len(key)
    _, end = json.JSONDecoder().raw_decode(s[i:])
    return s[:i] + jdump(obj) + s[i + end:]

def get_json(s, var):
    key = f'window.{var}='
    i = s.index(key) + len(key)
    return json.JSONDecoder().raw_decode(s[i:])[0]

def sub1(s, old, new, count=1):
    assert old in s, ('NO ENCONTRADO', old[:90])
    return s.replace(old, new, count)

def bbox(paths):
    xs, ys = [], []
    for p in paths:
        nums = [float(v) for v in re.findall(r'-?\d+(?:\.\d+)?', p)]
        xs += nums[0::2]; ys += nums[1::2]
    x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
    w, h = x1 - x0, y1 - y0; pad = max(w, h) * 0.04 + 4
    return [round(x0 - pad, 1), round(y0 - pad, 1), round(w + 2 * pad, 1), round(h + 2 * pad, 1)]

def fmt2(v):
    return f'{v:,.2f}'

def build(p):
    prov = PM[p]
    s = open(f'src/{p}.html').read()
    U = get_json(s, 'UPD'); N = json.load(open(f'work/{p}_new.json'))
    X = json.load(open(f'work/{p}_X.json'))
    # ---------- 1. capa social: UPD 17/10 ----------
    U['corte'] = '17/10/2026'; U['res'] = N['res']; U['ds'] = N['ds']; U['T'] = N['T']; U['nov'] = NOV[p]
    s = replace_json(s, 'UPD', U)
    s = sub1(s, 'CAPA DE ACTUALIZACIÓN — corte 07/10/2026 (común a los Volúmenes I, II y III)',
             'CAPA DE ACTUALIZACIÓN — corte 17/10/2026 (común a los Volúmenes I, II y III)')
    s = sub1(s, 'Lee window.UPD: libro «Gráficos DS Consolidado» del aplicativo IN Piura\n   (07/10/2026; actividades económicas en % de las familias de cada CP),',
             'Lee window.UPD: libro «Gráficos DS Consolidado» (VAL_PCT) del aplicativo IN Piura\n   (17/10/2026; valores absolutos y %; actividades económicas en % de las familias de cada CP),')
    s = sub1(s, 'MINAM 2018 · área de influencia aprobada · Copernicus GLO-30 · ${CORTE}', 'MINAM 2018 · área de influencia aprobada · Copernicus GLO-30 · 05/10/2026')
    s = sub1(s, 'nav.tabs button[data-upd]::after{content:"07/10"', 'nav.tabs button[data-upd]::after{content:"17/10"')
    # ---------- 2. avance del E2 (SOC.e2, común) ----------
    S = get_json(s, 'SOC')
    c = S['e2']['comps']
    c[1][2], c[1][3] = c[1][3], round(82 / 117 * 100, 1)
    c[1][4] = ('Fichas en 82 de 117 bloques V6 (52 Morropón, 24 Huancabamba, 6 Ayabaca); 80 al 07/10. Consolidado del aplicativo al '
               '17/10/2026: 364 fichas vigentes (194 F-DS-01, 98 F-DS-02 y 72 F-DS-03; 339 al 07/10). Nuevos bloques con ficha: 35 y 60 (Huancabamba)')
    for k in (0, 2, 3, 4, 5):
        c[k][2] = c[k][3]
    c[4][4] = ('Volúmenes I a III actualizados al 17/10/2026 con el consolidado DS (valores y %), el estado de los ecosistemas por bloque (E6) '
               'y los mapas de situación predial v2, cruzados con la tenencia declarada en la F-DS-01; avance de análisis sin variación (estimación de gabinete)')
    c[5][4] = ('Mapas de situación predial v2 de los 117 bloques (gabinete, 11/10/2026) incorporados en la pestaña «Tamizaje predial»; '
               'actividades previas de tamizaje en Yamango y San Juan de Bigote–Salitral programadas para el 06 y 08/10; avance sin variación al no constar actas suscritas')
    S['e2']['prev'] = S['e2']['new']
    S['e2']['new'] = round(sum(r[1] * r[3] for r in c) / 100, 1)
    lib = S['lib']['comps']
    lib[4][3] = ('Mapas de situación predial v2 de los 117 bloques (A3, gabinete GEORURAL–MIDAGRI y COFOPRI, 11/10/2026): '
                 'comunidades campesinas, predios COFOPRI y Tierras del Estado y/o posesionarios por bloque. Expedientes físico-legales pendientes')
    s = replace_json(s, 'SOC', S)
    for a, b in [('${n1(E.prev)} % al 03/10/2026 · suma ponderada', '${n1(E.prev)} % al 07/10/2026 · suma ponderada'),
                 ('sin variación respecto del 03/10/2026 · 30 % del peso del E2', 'sin variación respecto del 07/10/2026 · 30 % del peso del E2'),
                 ('<th>Peso</th><th>03/10</th><th>Avance</th><th>06/10</th>', '<th>Peso</th><th>07/10</th><th>Avance</th><th>17/10</th>'),
                 ('La barra gris es el avance al 03/10; la azul, el del 06/10.', 'La barra gris es el avance al 07/10; la azul, el del 17/10.'),
                 ('<th class="l">Evidencia al 06/10/2026</th>', '<th class="l">Evidencia (06/10/2026; gabinete predial al 11/10/2026)</th>')]:
        s = sub1(s, a, b)
    # ---------- 3. Integralidad (INT) ----------
    I = get_json(s, 'INT')
    res = {r[0]: r for r in N['res']['cab']}
    nf = res['Fichas sociales vigentes'][1]; act = res['Actores mapeados']
    favor = act[2].split(' ')[0]
    for k in I['kpi']:
        if k[0] == 'Diagnóstico social':
            k[1] = str(nf); k[2] = f'fichas F-DS vigentes · {favor} de {act[1]} actores a favor'
    A = sum(b['area_ha'] for b in E6 if b['provincia'] == prov)
    Bp = [b for b in E6 if b['provincia'] == prov]
    BU = sum(b['msavi_bajo_umbral_ha'] for b in Bp)
    ic = sum(b['ic_bloque'] * b['area_ha'] for b in Bp) / A
    nms = sum(1 for b in Bp if b['clase_estado'] in ('Moderado', 'Severo'))
    cc = sum(r['cc'] for r in X); ten = sum(1 for r in X if any(f[0] == 'ten' for f in r['f']))
    p19 = sum(1 for r in X if (r['npc'] or 0) > 19)
    if not any(k[0] == 'Ecosistemas E6' for k in I['kpi']):
        I['kpi'].append(['Ecosistemas E6', f'{ic:.2f}', f'IC medio · {nms} bloques moderado o severo · {fmt2(BU)} ha bajo umbral', '--st-veg'])
        I['kpi'].append(['Situación predial', f'{cc / A * 100:.1f} %', f'en comunidad campesina · {ten} bloques con tenencia ≠ F-DS-01', '--st-soc'])
    blq = res['Bloques con diagnóstico social'][1]; nb = len(X)
    tipos = {r[0][:8]: r[1] for r in N['ds']['cob'][0]['r']}
    cps = res['Centros poblados cubiertos'][2].split(' con ficha')[0]
    I['dsTxt'] = (f"{nf} fichas F-DS vigentes ({tipos['F-DS-01 ']} F-DS-01, {tipos['F-DS-02 ']} F-DS-02"
                  + (f" y {tipos['F-DS-03 ']} F-DS-03" if tipos.get('F-DS-03 ') else '') + f") con ficha propia en {blq} de {nb} bloques V6 y {cps} C.P. (consolidado del 17/10/2026)")
    kb = next((k for k in I['kpi'] if k[0] == 'Brecha MSAVI'), None)
    kbv = float(kb[1].replace(',', '')) if kb else None
    if kbv is not None and abs(kbv - BU) < 0.05:
        brtxt = f'Coincide con la brecha de este volumen ({kb[1]} ha), que ya usa la distribución areal.'
    else:
        brtxt = (f'La brecha de este volumen ({kb[1]} ha) se calcula por la media del bloque y difiere en {fmt2(abs(BU - kbv))} ha: '
                 'la cifra del E6 es la que puede sumarse con las otras dos provincias.') if kb else ''
    for g in I['gaps']:
        if g[0] == 'Índices de vegetación':
            g[3] = 'Adoptar la brecha areal del E6 (mismo criterio en los 117 bloques) como cifra del indicador y recalcular la brecha consolidada.'
        if g[0] == 'E2 Diagnóstico social':
            sinf = [r['b'] for r in X if not r['nf']]
            g[1] = (f"Se aplicaron {'3' if tipos.get('F-DS-03 ') else '2'} de 7 instrumentos; fichas en {blq} de {nb} bloques al 17/10/2026"
                    + (f" (sin ficha: {', '.join(sorted(sinf, key=lambda x: (len(x), x)))})" if len(sinf) <= 10 else f'; {len(sinf)} bloques sin ficha')
                    + '. La percepción de peligros y cambio climático (F-DS-06) no se aplicó y la tenencia declarada difiere del gabinete predial en '
                    + f'{ten} bloques.')
    if not any(g[0] == 'E6 Ecosistemas' for g in I['gaps']):
        nev = sum(1 for b in Bp if b['clase_evidencia'] != 'A')
        I['gaps'].append(['E6 Ecosistemas',
            f'Incorporado (Tablero E6): estado de los {nb} bloques por índice de condición (cinco atributos, Gann et al., 2019) y MSAVI 2024 por distribución areal: {fmt2(BU)} ha bajo el umbral 0.4976. {nev} bloques con evidencia de campo B o C; las especies «esperadas» provienen de la firma espectral, no de inventario.',
            'El ecosistema de campo (F-DT) difiere del MINAM 2018 en parte de los bloques; la clase de estado no sustituye la evaluación del ecosistema de referencia por piso.',
            'Confirmar en campo los bloques con evidencia B o C y levantar inventario florístico donde las especies son solo «esperadas».', 'Media'])
        nsin = sum(1 for r in X if r.get('noimg'))
        I['gaps'].append(['Tamizaje predial',
            f'Mapas de situación predial v2 de gabinete (GEORURAL–MIDAGRI y COFOPRI, tolerancia 5 %), sin verificación de campo ni actas suscritas. {ten} bloques con tenencia declarada distinta del gabinete y {p19} con más de 19 predios COFOPRI.'
            + (' La lámina SP-059 (bloque 5) no está entre los archivos extraídos de la carpeta compartida.' if nsin else ''),
            'Las metas físicas y el área de intervención pueden reducirse si la libre disponibilidad no se acredita; el Paso 5 limita a 19 predios por bloque.',
            'Tamizaje de campo con las juntas directivas comunales y los titulares; redelimitar o sectorizar los bloques sobre 19 predios; excluir el área agrícola delimitada por el E6.', 'Alta'])
    s = replace_json(s, 'INT', I)
    # ---------- 4. datos de las nuevas pestañas ----------
    keep = ['bloque', 'distrito', 'microcuenca', 'area_ha', 'msavi_2024', 'ecosistema', 'alt_media', 'pendiente_media_pct', 'piso_dom',
            'cob1_pct', 'cob2_pct', 'cob3_pct', 'cob4_pct', 'cob5_pct', 'cob6_pct', 'cob8_pct', 'cob9_pct', 'DN1_ha', 'DN2_ha', 'DN3_ha', 'DN4_ha', 'DN5_ha',
            'msavi_bajo_umbral_ha', 'msavi_bajo_umbral_pct', 'msavi_rest_activa_ha', 'n_parcelas', 'ic_amenazas', 'ic_cond_fisicas', 'ic_div_estructural',
            'ic_funciones', 'ic_intercambios', 'ic_bloque', 'clase_estado', 'tipologia_S2DR4', 'tipologia_NDVI2025', 'delim_rest_activa_ha',
            'delim_enriquecimiento_ha', 'delim_conservacion_ha', 'delim_tamizaje_agricola_ha', 'dt_cobveg_campo', 'clase_evidencia', 'fecha_verificacion',
            'n_especies_esperadas', 'n_especies_confirmadas', 'path', 'utm', 'esp']
    ECO6 = dict(prov=prov, vb=bbox([b['path'] for b in Bp]), b=[{k: b[k] for k in keep} for b in Bp])
    xkeep = ['b', 'd', 'ha', 'pi', 'ipp', 'eco', 'v6', 'v6p', 'cls', 'bu', 'buha', 'sit', 'pcc', 'pcof', 'pest', 'cc', 'cof', 'est', 'ccn', 'npc',
             'nccpp', 'ccpp', 'pob', 'asis', 'sp', 'img', 'esc', 'nf', 'regd', 'npd', 'supcc', 'lind', 'tit', 'f']
    XX = []
    for r in X:
        o = {k: r[k] for k in xkeep}
        o['pcc'], o['pcof'], o['pest'] = round(o['pcc'], 2), round(o['pcof'], 2), round(o['pest'], 2)
        if not any(os.path.exists(os.path.join(d, r['img'])) for d in SPDIR):
            o['noimg'] = 1
        XX.append(o)
    SPX = dict(prov=prov, corte='17/10/2026', brTxt=brtxt, x=XX)
    # ---------- 5. pestañas y secciones ----------
    s = sub1(s, '<button role="tab" data-g="Estudios básicos" data-p="sue" aria-selected="false">Suelos</button>',
             '<button role="tab" data-g="Estudios básicos" data-p="sue" aria-selected="false">Suelos</button>\n  <button role="tab" data-g="Estudios básicos" data-p="eco" aria-selected="false">Ecosistemas</button>')
    s = sub1(s, '<button role="tab" data-g="Gestión" data-p="cond" aria-selected="false">Condicionantes</button>',
             '<button role="tab" data-g="Gestión" data-p="cond" aria-selected="false">Condicionantes</button>\n  <button role="tab" data-g="Gestión" data-p="tam" aria-selected="false">Tamizaje predial</button>')
    nB = len(X)
    sec_eco = f'''<section id="p-eco" hidden>
<div class="sec-h"><h2>Estado de los ecosistemas por bloque</h2><span class="tag">Entregable 6 · S2DR4 1 m · MSAVI 2024 · índice de condición · corte 10/2026</span></div>
<p class="lede">La unidad productora del proyecto es el ecosistema. El Entregable 6 evalúa el estado de cada uno de los {nB} bloques de {prov} con tres líneas de evidencia: la clasificación de coberturas de 92 tomas S2DR4 de 1 m, el MSAVI 2024 por clases (que mide la brecha frente al umbral 0.4976 de la R.M. N.° 00213-2024-MINAM) y un índice de condición de cinco atributos contrastado con referencias del mismo piso altitudinal y con las fichas F-DT de campo. Su delimitación interna asigna cada sector del bloque a restauración activa, enriquecimiento, conservación o exclusión agrícola.</p>
<div id="e6Root"></div>
</section>
'''
    sec_tam = f'''<section id="p-tam" hidden>
<div class="sec-h"><h2>Tamizaje predial por bloque</h2><span class="tag">Paso 5 · mapas de situación predial v2 · gabinete 11/10/2026 · F-DS-01 17/10/2026</span></div>
<p class="lede">Situación predial de gabinete de los {nB} bloques de {prov}: superficie en comunidades campesinas (GEORURAL–MIDAGRI), en predios rurales COFOPRI fuera de ellas y sin registro (Tierras del Estado y/o posesionarios), con los centros poblados asociados y la lámina A3 de cada bloque. Se cruza con la tenencia que declara la población en la F-DS-01 y con el criterio del Paso 5 (máximo 19 predios rurales por bloque, exclusión de áreas agrícolas) para anticipar con quién se suscribe cada acta de libre disponibilidad.</p>
<div id="spRoot"></div>
</section>
'''
    s = sub1(s, '<section id="p-brec" hidden>', sec_eco + '<section id="p-brec" hidden>')
    s = sub1(s, '<section id="p-con" hidden>', sec_tam + '<section id="p-con" hidden>')
    # ---------- 6. cabecera, pie y enlaces entre volúmenes (copias) ----------
    eb = re.search(r'(<p class="eyebrow">Proyecto IN Piura · CUI 2669244 · Volumen [IV]+ · )([^<]*)(</p>)', s)
    txt = eb.group(2)
    txt = re.sub(r'Revisión (\d+)', lambda m: f'Revisión {int(m.group(1)) + 1}', txt)
    txt = txt.replace('Revisión integrada DT + DS (actualización 07/10)', 'Revisión integrada DT + DS (actualización 17/10)')
    txt = txt.replace('DS 07/10/2026', 'DS 17/10/2026 · E6 ecosistemas 10/2026 · situación predial v2 11/10/2026')
    s = s[:eb.start(2)] + txt + s[eb.end(2):]
    s = re.sub(r'(<ul class="hero-lines" aria-label="Estudios integrados">.*?)(</ul>)',
               lambda m: m.group(1) + '<li><i style="background:var(--e6-ok)"></i>E6 Ecosistemas</li><li><i style="background:var(--sp-cc)"></i>Tamizaje predial</li>' + m.group(2),
               s, count=1, flags=re.S)
    for q in 'MHA':
        s = s.replace(f'href="{ORIG[q]}"', f'href="{COPY[q]}"')
    upd = ('<div><b>Actualización 17/10/2026:</b> libro «Gráficos DS Consolidado» (VAL_PCT) del aplicativo IN Piura, emitido el '
           + N['res']['id']['Fecha de emisión'] + ', en valores absolutos y % (reemplaza al del 07/10; se omiten los nombres de responsables, informantes y actores que el libro incluye) · '
           'Tablero E6 «Estado de los ecosistemas por bloque de intervención» (S2DR4 1 m, MSAVI 2024, índice de condición, delimitación interna y especies) · '
           'Índice y mapas de situación predial de los 117 bloques V6 v2 (A3, GEORURAL–MIDAGRI y COFOPRI, Subentregable 2.2, 11/10/2026)</div>\n  ')
    s = sub1(s, '<div><b>Actualización 07/10/2026:</b>', upd + '<div><b>Actualización 07/10/2026:</b>')
    # ---------- 7. CSS, datos y módulo ----------
    css = open('bin/capa_eco_tam.css').read()
    s = sub1(s, '<header class="hero">', css + '\n<header class="hero">')
    js = open('bin/capa_eco_tam.js').read()
    data = f'<script>window.ECO6={jdump(ECO6)};window.SPX={jdump(SPX)};</script>\n'
    i = s.rindex('</script>') + len('</script>')
    s = s[:i] + '\n' + data + js + s[i:]
    os.makedirs('out', exist_ok=True)
    open(f'out/{p}.html', 'w').write(s)
    return len(s)

if __name__ == '__main__':
    for p in (sys.argv[1:] or ['A', 'H', 'M']):
        print(p, build(p))
