"""Aplica la actualización 06/10/2026 a los tres volúmenes del Diagnóstico Territorial y Social."""
import json, re, sys, collections
S = sys.argv[1]
TR = '/root/.claude/projects/-home-user-Salomon1969/bb473581-4e75-5721-b1f1-c28ac0ad757c/tool-results'
SRC = {'aya': f'{TR}/artifact-5c30b825-1791078202-5918.html',
       'hua': f'{TR}/artifact-0af2d19c-1791078186-5e74.html',
       'mor': f'{TR}/artifact-bdc5ae0a-1791078175-4251.html'}
CSS = open(f'{S}/tools/capa_0610.css', encoding='utf-8').read()
JS = open(f'{S}/tools/capa_0610.js', encoding='utf-8').read()
f0 = lambda v: f'{v:,.0f}'
f1 = lambda v: f'{v:,.1f}'
f2 = lambda v: f'{v:,.2f}'
WARN = []

def rep(s, old, new, label, count=1):
    n = s.count(old)
    if n == 0:
        WARN.append(f'NO ENCONTRADO [{label}]: {old[:90]}')
        return s
    if count and n != count:
        WARN.append(f'{n} coincidencias [{label}] (se esperaba {count})')
    return s.replace(old, new)

def tab(u, sec, letter):
    for t in u['ds'][sec]:
        if t['t'].startswith(letter + '.'): return t
    return None

def stats(k, u):
    st = {}
    A = tab(u, 'cob', 'A'); st['reg'] = {r[0].split(' · ')[0]: r[1] for r in A['r']}
    D = tab(u, 'cob', 'D'); st['bl'] = {str(r[0]): r[1] for r in D['r']}
    C = tab(u, 'cob', 'C'); st['cpcov'] = C['r'][0][1]; st['cpsin'] = C['r'][1][1]; st['cpnote'] = ' '.join(C['n'])
    st['cab'] = {r[0]: (r[1], r[2] if len(r) > 2 else '') for r in u['res']['cab']}
    st['sin'] = sorted(set(u['pend']) - set(st['bl']), key=lambda x: (len(x), x))
    # ecosistemas y pendientes
    tot = collections.defaultdict(float); A_ = 0; AI = 0
    for b, e in u['eco'].items():
        for n, s, ha, p in e['e']: tot[n] += ha
        A_ += sum(x[2] for x in e['e']); AI += e['ai']
    st['eco'] = sorted(tot.items(), key=lambda x: -x[1]); st['A'] = A_; st['AI'] = AI
    P = u['pend']; st['ha75'] = sum(p[17] for p in P.values()); st['Ap'] = sum(p[2] for p in P.values())
    st['med'] = sum(p[2] * p[5] for p in P.values()) / st['Ap']
    st['hH'] = sorted([(b, p[16], p[17]) for b, p in P.items() if p[16] >= 25], key=lambda x: -x[1])
    cal = u['T']['cal']; i = cal['h'].index('Tema'); st['cal'] = collections.Counter(r[i] for r in cal['r']); st['ncal'] = len(cal['r'])
    return st

def eco_txt(st):
    return '; '.join(f'{n.lower() if n == "Zona agrícola" else n} {f2(ha)} ha ({f1(ha / st["A"] * 100)} %)' for n, ha in st['eco'])

def nov(k, u, st, soc):
    R = st['reg']; old = soc['reg']; cab = st['cab']
    reg_txt = f"{R.get('F-DS-01', 0)} F-DS-01, {R.get('F-DS-02', 0)} F-DS-02 y {R.get('F-DS-03', 0)} F-DS-03"
    old_txt = f"{old.get('F-DS-01', 0)}, {old.get('F-DS-02', 0)} y {old.get('F-DS-03', 0)}"
    nf = sum(R.values()); nb = len(st['bl']); NB = len(u['pend'])
    pob, pobd = cab['Población del ámbito']; acts, actd = cab['Actores mapeados']
    out = []
    if k == 'aya':
        out.append(['cri', 'El consolidado del 06/10 confirma las mismas 19 fichas: Ayabaca sigue sin campo desde mediados de septiembre',
            [f'El libro consolidado del aplicativo emitido el 06/10/2026 registra <strong>{nf} fichas vigentes</strong> ({reg_txt}), igual que el respaldo del 03/10, en <strong>{nb} de {NB} bloques</strong>. Siguen sin ficha los bloques {", ".join(st["sin"])}.',
             f'Cubre {st["cpcov"]} de los 17 centros poblados del catálogo INEI. {st["cpnote"].replace("Fuente: ", "")}',
             'La segunda campaña en Ayabaca y la aplicación de la F-DS-03 siguen pendientes (P-DS-00).'], 'Gráficos DS Consolidado, aplicativo IN Piura, 06/10/2026.'])
    elif k == 'hua':
        newb = [b for b in st['bl'] if b not in {r[4] for r in soc['ent']}]
        out.append(['ok', f'Huancabamba suma {nf - sum(old.values())} fichas y llega a {nb} de {NB} bloques',
            [f'El consolidado del 06/10/2026 registra <strong>{nf} fichas vigentes</strong> ({reg_txt}; al 03/10: {old_txt}). Se incorporan los bloques <strong>{", ".join(newb)}</strong>: Huamala Alto (21), Abalque (28), Pariamarca Centro (64) y Huamala Baja (66). Abalque y Huamala Baja se levantaron el 06 y el 05/10.',
             'Pariamarca Centro (bloque 64) es la <strong>primera localidad del distrito de Huancabamba con diagnóstico social</strong>: 4 fichas F-DS-01 y una entrevista al presidente de la JASS.',
             f'Siguen sin ficha {len(st["sin"])} bloques, entre ellos los cinco del núcleo de brecha de Huarmaca y San Miguel de El Faique (M4B4, M12B1, M4B3, M20B1, M2B8) y M30B5 en el distrito de Huancabamba.'], 'Gráficos DS Consolidado, aplicativo IN Piura, 06/10/2026.'])
    else:
        out.append(['ok', 'Morropón mantiene sus 226 fichas, ahora asignadas al catálogo V6',
            [f'El consolidado del 06/10/2026 registra <strong>{nf} fichas vigentes</strong> ({reg_txt}; al 03/10: {old_txt}) con ficha propia en <strong>{nb} de {NB} bloques V6</strong>. Las fichas que el respaldo cargaba en los bloques retirados 1, 46 y M18B5 se cuentan en su bloque vigente; M8B2 aparece con 3 fichas.',
             f'Siguen sin ficha {len(st["sin"])} bloques: {", ".join(st["sin"])}. El conglomerado 83–87 ya tiene F-DS-03 en los bloques 85, 86 y 87.',
             f'Cubre {st["cpcov"]} de los 78 centros poblados del catálogo INEI ({st["cpsin"]} sin ficha).'], 'Gráficos DS Consolidado, aplicativo IN Piura, 06/10/2026.'])
    out.append(['adv', f'Población declarada: {f0(pob)} habitantes, con cada centro poblado contado una sola vez',
        [f'La F-DS-01 consigna población total en {pobd.split(":")[1].split("·")[0].strip()} y {pobd.split("·")[1].strip()}. El consolidado toma, para cada CP, la ficha que más se repite (en empate, la más reciente) y no suma los ámbitos repetidos ni los agrupados.',
         'Esta cifra es declarativa y no reemplaza a la del catálogo INEI de las demás pestañas: las diferencias de más del doble respecto del INEI se listan en «Discrepancias».'], 'Consolidado DS, F-DS-01 numeral 2.'])
    # tenencia y gobernanza
    J = next(t for t in u['ds']['f01'] if 'Régimen predominante' in t['t'])
    reg = '; '.join(f'{r[0][:1].lower()+r[0][1:]} en {r[1]}' for r in J['r'])
    G = next(t for t in u['ds']['f01'] if 'Capacidades' in t['t'])
    g = {r[0]: r for r in G['r']}
    M = next(t for t in u['ds']['f01'] if 'Percepciones' in t['t'])
    mig = next(r for r in M['r'] if r[0].startswith('Tasa'))
    nmig = sum(v or 0 for v in mig[1:4])
    ten_extra = ''
    if k == 'aya':
        ten_extra = ' Dos bloques de régimen comunal (56 y M17B6) declaran a la vez que no se superponen a tierras comunales: debe aclararse en el tamizaje predial.'
    out.append(['adv', 'Tenencia y gobernanza: con quién se firman las actas',
        [f'Régimen predominante por bloque (F-DS-01, numeral 4): {reg} bloque(s).{ten_extra}',
         f'Juntas directivas vigentes en {g["Junta Directiva vigente"][1]} y rondas campesinas activas en {g["Ronda Campesina activa"][1]} de los centros poblados que responden; comité de recursos naturales en {g["Comité de Recursos Naturales"][1]}. La migración juvenil es alta en {mig[1]} de {nmig} centros poblados: condiciona la mano de obra para las MRR-CCC.'], 'Consolidado DS, F-DS-01 numerales 3, 4 y 7.'])
    calt = ', '.join(f'{v} «{t.lower()}»' for t, v in st['cal'].most_common(4))
    out.append(['', f'El consolidado registra {st["ncal"]} observaciones de control de calidad',
        [f'Las más frecuentes: {calt}. Se muestran en «Discrepancias» con la fecha de la ficha de referencia, sin nombres.'], 'Hoja «Control de calidad de los datos», 06/10/2026.'])
    hH = ', '.join(f'{b} ({f1(h)} %)' for b, h, _ in st['hH'][:6])
    out.append(['', 'Territorio: ecosistemas V6, área de influencia y pendientes por bloque',
        [f'Ecosistemas en los {NB} bloques ({f2(st["A"])} ha): {eco_txt(st)}. Área de influencia aprobada: {f2(st["AI"])} ha. Láminas por bloque en «Geoespacial» y en «Ficha por bloque».',
         f'La matriz de pendientes V6 (Copernicus GLO-30) da una media ponderada de {f1(st["med"])} % y <strong>{f2(st["ha75"])} ha con pendiente mayor de 75 %</strong> ({f1(st["ha75"] / st["Ap"] * 100)} % del ámbito)' + (f'; los bloques con un cuarto o más de su superficie en esa clase son {hH}.' if hH else '; ningún bloque tiene un cuarto de su superficie en esa clase.')], 'Atlas de ecosistemas V6 y Matriz de pendientes V6, ANIN-DIME-SESDI, 05/10/2026.'])
    return out

def strip_skeleton(h):
    i = h.find('<body>')
    if h.startswith('<!doctype') and i > 0: h = h[i + len('<body>'):].lstrip('\n')
    h = re.sub(r'\s*</body></html>\s*$', '\n', h)
    return h

E2_NEW = {'dig': 68.4, 'ana': 50.0}

def patch_soc(soc, st_all, k):
    soc['corte'] = '06/10/2026'
    st = st_all[k]
    soc['reg'] = {f: st['reg'].get(f, 0) for f in ('F-DS-01', 'F-DS-02', 'F-DS-03')}
    nb = {p: len(st_all[p]['bl']) for p in st_all}
    tot_b = sum(nb.values())
    comps = soc['e2']['comps']
    new = []
    for c in comps:
        name, w, a, b, nota = c
        a = b
        if name.startswith('Digitación'):
            b = round(tot_b / 117 * 100, 1)
            nota = f'Fichas en {tot_b} de 117 bloques V6 ({nb["mor"]} Morropón, {nb["hua"]} Huancabamba, {nb["aya"]} Ayabaca); 75 al 03/10. Consolidado del aplicativo al 06/10/2026: {sum(sum(st_all[p]["reg"].values()) for p in st_all)} fichas vigentes'
        elif name.startswith('Análisis y redacción'):
            b = E2_NEW['ana']
            nota = 'Volúmenes I a III actualizados al 06/10/2026 con el consolidado DS (valores y %), la tenencia por bloque, el control de calidad, los ecosistemas V6 y la matriz de pendientes V6 (estimación de gabinete)'
        elif name.startswith('Expediente preliminar'):
            nota = 'Instrumentos y manuales listos; actividades previas de tamizaje predial programadas en Yamango (06/10) y San Juan de Bigote–Salitral (08/10); contratación del equipo predial y legal en marcha'
        elif name.startswith('Registro de saberes'):
            pass
        new.append([name, w, a, b, nota])
    soc['e2']['comps'] = new
    soc['e2']['prev'] = soc['e2']['new']
    soc['e2']['new'] = round(sum(c[1] * c[3] for c in new) / 100, 1)
    L = soc['lib']['comps']
    for c in L:
        if c[0].startswith('Instrumentos y herramientas'):
            c[3] = c[3] + '; manuales de usuario del tamizaje predial y del aplicativo de liberación de áreas, evaluación de ejercicios, guía rápida de bolsillo, agenda y registro de asistencia y fichas F-LA en papel de respaldo (05/10/2026)'
        if c[0].startswith('Trabajo de campo'):
            c[3] = 'Actividades previas de tamizaje predial programadas: Yamango (06/10) y Dótor, Tórtola, Hualas y Hornopampa en San Juan de Bigote–Salitral (08/10); ningún acta suscrita al corte'
    RET = {'1', '7', '25', '29', '32', '33', '46', '48', '68', '74', '75', 'M18B5'}
    for r in soc['ent']:
        if r[4] in RET and not r[9]: r[9] = 'Bloque retirado del catálogo V6'
    return soc

def main():
    U = {k: json.load(open(f'{S}/work/upd_{k}.json')) for k in SRC}
    SOC = {k: json.load(open(f'{S}/work/{k}_soc.json')) for k in SRC}
    ST = {k: stats(k, U[k]) for k in SRC}
    for k, path in SRC.items():
        h = open(path, encoding='utf-8').read()
        h = strip_skeleton(h)
        u = U[k]; st = ST[k]
        u['nov'] = nov(k, u, st, SOC[k])
        # ---- window.SOC ----
        m = re.search(r'<script>window\.SOC=(.*?);?</script>', h, re.S)
        old_js = m.group(1)
        soc = patch_soc(json.loads(old_js.rstrip(';')), ST, k)
        h = h.replace(m.group(0), '<script>window.SOC=' + json.dumps(soc, ensure_ascii=False, separators=(',', ':')) + ';</script>')
        # ---- textos de la capa social (E2) ----
        h = rep(h, "% al 27/09/2026 · suma ponderada", "% al 03/10/2026 · suma ponderada", 'e2 kpi')
        h = rep(h, "<small>10 % al 27/09/2026 · 30 % del peso del E2</small>", "<small>sin variación respecto del 03/10/2026 · 30 % del peso del E2</small>", 'lib kpi')
        h = rep(h, "<th>27/09</th><th>Avance</th><th>03/10</th>", "<th>03/10</th><th>Avance</th><th>06/10</th>", 'e2 th')
        h = rep(h, "La barra gris es el avance al 27/09; la azul, el del 03/10.", "La barra gris es el avance al 03/10; la azul, el del 06/10.", 'e2 note')
        h = rep(h, "Evidencia al 03/10/2026</", "Evidencia al 06/10/2026</", 'lib th')
        h = rep(h, "${bar(10,L.tot)}", "${bar(27.5,L.tot)}", 'lib bar')
        h = rep(h, "Sin entrevistas de saberes ancestrales en esta provincia al 03/10/2026.", "Sin entrevistas de saberes ancestrales en esta provincia al 06/10/2026.", 'sa', 0)
        h = rep(h, 'Diagnóstico de involucrados y expediente preliminar de liberación de áreas · 03/10/2026</span>', 'Diagnóstico de involucrados y expediente preliminar de liberación de áreas · 06/10/2026</span>', 'cob tag')
        if k == 'aya': h = rep(h, "Fichas vigentes (corte 03/10/2026)", "Fichas vigentes (corte 03/10/2026; sin variación al 06/10)", 'cIns label')
        # ---- cabecera y textos de integralidad ----
        h = rep(h, ' · DS 03/10/2026 · ', ' · DS 06/10/2026 · ecosistemas y pendientes V6 05/10/2026 · ', 'eyebrow')
        if k == 'aya':
            h = rep(h, 'Volumen III · Revisión 6 ·', 'Volumen III · Revisión 7 ·', 'rev')
            h = rep(h, '"dsTxt":"19 fichas F-DS (11 F-DS-01 y 8 F-DS-02) en 11 de 17 C.P."', '"dsTxt":"19 fichas F-DS (11 F-DS-01 y 8 F-DS-02) en 11 de 17 C.P., sin variación al 06/10/2026"', 'dsTxt')
            h = rep(h, 'Fichas en 6 de 10 bloques; la percepción del riesgo', 'Fichas en 6 de 10 bloques, sin variación al 06/10/2026 (sin ficha: 27, 36, M17B5 y M17B10); la percepción del riesgo', 'gap')
            h = rep(h, '(cortes del 23, 26 y 27/09 y del 03/10)', '(cortes del 23, 26 y 27/09, del 03/10 y consolidado del 06/10)', 'P-DS-00')
        elif k == 'hua':
            h = rep(h, 'Volumen II · Revisión integrada DT + DS ·', 'Volumen II · Revisión integrada DT + DS (actualización 06/10) ·', 'rev')
            h = rep(h, '"dsTxt":"73 registros F-DS (47 depurados) con cobertura en 18 de 48 bloques y 24 de 58 C.P. (corte 03/10/2026)"',
                    '"dsTxt":"84 fichas F-DS vigentes (44 F-DS-01, 24 F-DS-02 y 16 F-DS-03) con cobertura en 22 de 48 bloques y 26 de 58 C.P. (consolidado del 06/10/2026)"', 'dsTxt')
            h = rep(h, 'fichas en 18 de 48 bloques al 03/10/2026 (14 al 27/09), sin ninguna en el núcleo de brecha ni en el distrito de Huancabamba.',
                    'fichas en 22 de 48 bloques al 06/10/2026 (18 al 03/10 y 14 al 27/09); la primera del distrito de Huancabamba es Pariamarca Centro (bloque 64) y el núcleo de brecha sigue sin ninguna.', 'gap')
            h = rep(h, 'Levantar el diagnóstico social de los 5 bloques del núcleo de brecha y del distrito de Huancabamba (64, M30B5), que siguen sin ficha al 03/10.',
                    'Levantar el diagnóstico social de los 5 bloques del núcleo de brecha y de M30B5 (distrito de Huancabamba), que siguen sin ficha al 06/10; el bloque 64 ya tiene F-DS-01 y F-DS-03 (Pariamarca Centro).', 'P-02')
        else:
            h = rep(h, 'Volumen I · Revisión 7 ·', 'Volumen I · Revisión 8 ·', 'rev')
            h = rep(h, '"dsTxt":"146 fichas F-DS depuradas de 226 registros, con ficha propia en 51 de 59 bloques (corte 03/10/2026, sin variación desde el 27/09)"',
                    '"dsTxt":"226 fichas F-DS vigentes (128 F-DS-01, 57 F-DS-02 y 41 F-DS-03) con ficha propia en 52 de 59 bloques V6 y 68 de 78 C.P. (consolidado del 06/10/2026)"', 'dsTxt')
            h = rep(h, '8 bloques sin ficha propia al 03/10/2026 (el conglomerado 83–87 ya tiene F-DS-01, pero ninguna F-DS-03).',
                    '7 bloques V6 sin ficha propia al 06/10/2026 (55, 58, M2B1, M17B4, M36B2, M6B2-1 y M6B2-3); el conglomerado 83–87 ya tiene F-DS-03 en 85, 86 y 87.', 'gap')
        # ---- pie de fuentes ----
        foot = ('<div><b>Actualización 06/10/2026:</b> libro «Gráficos DS Consolidado» del aplicativo IN Piura (valores absolutos y %, emitido el '
                + str(u['res']['id'].get('Fecha de emisión', '06/10/2026')) +
                '; fichas F-DS-01 a F-DS-03 vigentes, un centro poblado contado una sola vez; sin nombres de informantes ni de responsables) · Atlas de ecosistemas — bloques de intervención V6 (117 láminas A4) e IN_Piura_Ecosistemas_por_Bloque_V6, sobre el Mapa Nacional de Ecosistemas del Perú (MINAM 2018, R.M. N.° 440-2018-MINAM) · Bloques V6 y área de influencia aprobada (AI_aprobado_2) · Matriz de pendientes por bloque V6 (DEM Copernicus GLO-30, 05/10/2026) · Programación de campo del 05 al 09/10/2026 y manuales del tamizaje predial y del aplicativo de liberación de áreas (05/10/2026)</div>\n  ')
        anchor = '<div><b>Principio de no invención de datos:</b>' if '<div><b>Principio de no invención de datos:</b>' in h else '<div><b>Regla de datos:</b>'
        h = rep(h, anchor, foot + anchor, 'footer')
        # ---- capa nueva ----
        h = rep(h, '<header class="hero">', CSS + '\n<header class="hero">', 'css')
        data = '<script>window.UPD=' + json.dumps(u, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/') + ';</script>\n'
        h = h.rstrip() + '\n' + data + JS
        open(f'{S}/out/{k}/index.html', 'w', encoding='utf-8').write(h)
        print(k, len(h) // 1024, 'KB', 'E2', soc['e2']['prev'], '->', soc['e2']['new'])
    for w in WARN: print('WARN', w)

if __name__ == '__main__':
    main()
