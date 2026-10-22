"""Construye window.UPD (actualización 06/10/2026) por provincia."""
import json, csv, re, sys, collections, openpyxl
S = sys.argv[1]; REPO = sys.argv[2]
PROVS = {'aya': 'Ayabaca', 'hua': 'Huancabamba', 'mor': 'Morropón'}

def ds_tables(d):
    out = {}
    names = {'Cobertura': 'cob', 'F-DS-01': 'f01', 'F-DS-02': 'f02', 'F-DS-03': 'f03'}
    tsheets = {}
    for sh in d['sheets']:
        key = next((v for k, v in names.items() if sh['name'].startswith(k)), None)
        if sh['name'].startswith('T '):
            tsheets[sh['name']] = sh['tables'][0] if sh['tables'] else None
            continue
        if key is None: continue
        tabs = [t for t in sh['tables'] if t['header']]
        titles = {t['title'] for t in tabs}
        lst = []
        for t in tabs:
            if t['title'].endswith(' (%)') and t['title'][:-4] in titles: continue
            comp = next((x for x in tabs if x['title'] == t['title'] + ' (%)'), None)
            mode = 'none'
            if comp:
                dd = comp['desc']
                if 'cada columna' in dd: mode = 'col'
                elif 'cada fila' in dd: mode = 'row'
                elif 'total general' in dd: mode = 'all'
                elif 'Porcentaje sobre' in dd: mode = 'base'
            if mode == 'none' and not t['title'].endswith('(%)'): mode = 'val'
            hdr = t['header']; rows = t['rows']
            base = None
            if '% de la base' in hdr:
                mode = 'base'
                br = [r for r in rows if str(r[0]).startswith('Base')]
                if br: base = br[0][1]
                rows = [r for r in rows if not str(r[0]).startswith('Base')]
                hdr = hdr[:2]
            # columnas calculadas (Total, CP que responden...) -> se recalculan en la página
            calc = [i for i, h in enumerate(hdr) if i > 0 and h in ('Total', 'Total de fichas', 'Población total', 'CP que responden', 'Bloques que responden', 'Entrevistas', 'Total de actores')]
            lst.append({'t': t['title'], 'd': t['desc'], 'h': hdr, 'r': rows, 'n': t['notes'], 'tot': t['total'], 'm': mode, 'base': base, 'calc': calc})
        out[key] = (out.get(key) or []) + [x for x in lst]
    # el primer "título" de cada hoja es el encabezado de la hoja (sin cabecera)
    return out, tsheets

def dateonly(v):
    return str(v).split(' · ')[0] if v else v

def sanitize(tsheets):
    T = {}
    def tab(name):
        t = tsheets.get(name)
        return (t['header'], t['rows']) if t else (None, [])
    h, r = tab('T Demografía por centro poblado')
    if h:
        i = h.index('Ficha de referencia'); h = list(h); h[i] = 'Fecha de la ficha de referencia'
        T['demo'] = {'h': h, 'r': [[dateonly(x) if j == i else x for j, x in enumerate(row)] for row in r]}
    h, r = tab('T Actividades económicas')
    if h: T['act'] = {'h': h, 'r': r}
    h, r = tab('T Tenencia por bloque')
    if h: T['ten'] = {'h': h, 'r': r}
    h, r = tab('T Actores clave')
    if h:
        iN = h.index('Nombre del actor'); iC = h.index('Cargo')
        hh = [x for j, x in enumerate(h) if j != iN]
        rr = []
        for row in r:
            row = list(row) + [None] * (len(h) - len(row))
            nm = str(row[iN] or ''); cg = row[iC]
            if not cg and '/' in nm: cg = nm.split('/', 1)[1].strip()
            if cg: cg = cg[:1].upper() + cg[1:].lower() if cg.isupper() else cg
            row[iC] = cg or '—'
            rr.append([x for j, x in enumerate(row) if j != iN])
        T['actores'] = {'h': hh, 'r': rr}
    h, r = tab('T Entrevistas')
    if h:
        iN = h.index('Entrevistado/a')
        T['entrev'] = {'h': [x for j, x in enumerate(h) if j != iN], 'r': [[x for j, x in enumerate(list(row) + [None] * (len(h) - len(row))) if j != iN] for row in r]}
    h, r = tab('T Control de calidad')
    if h:
        i = h.index('Ficha de referencia'); h = list(h); h[i] = 'Fecha de la ficha de referencia'
        T['cal'] = {'h': h, 'r': [[dateonly(x) if j == i else x for j, x in enumerate(row)] for row in r]}
    h, r = tab('T Fichas F-DS-01 por CP')
    if h:
        drop = {h.index('Responsable'), h.index('Entrevistado')}
        T['fichas'] = {'h': [x for j, x in enumerate(h) if j not in drop], 'r': [[x for j, x in enumerate(list(row) + [None] * (len(h) - len(row))) if j not in drop] for row in r]}
    return T

def resumen(d):
    pre = d['sheets'][0]['pre']
    ident, cab, cont, avisos = {}, [], [], []
    mode = None
    for row in pre:
        s = str(row[0])
        if s.startswith('1. '): mode = 'id'; continue
        if s.startswith('2. '): mode = 'cab'; continue
        if s.startswith('3. '): mode = 'cont'; continue
        if row[0] in ('Campo', 'Indicador', 'Sección'): continue
        if s.startswith('Aviso') or s.startswith('Libro') or s.startswith('Cada '): avisos.append(s); continue
        if mode == 'id' and len(row) >= 2: ident[s] = row[1]
        elif mode == 'cab' and len(row) >= 2: cab.append(row)
        elif mode == 'cont' and len(row) >= 2: cont.append(row)
    return {'id': ident, 'cab': cab, 'cont': cont, 'av': avisos}

# ---- ecosistemas V6 y AI ----
wb = openpyxl.load_workbook(f'{REPO}/mapas/salidas/IN_Piura_Ecosistemas_por_Bloque_V6.xlsx', data_only=True)
ws = wb['Bloque_Ecosistema']
eco = collections.defaultdict(list); cat = {}
for row in ws.iter_rows(min_row=8, values_only=True):
    if not row or row[1] is None or not isinstance(row[0], (int, float)): continue
    _, b, mc, pv, dist, ecos, sym, geo, _, area_cat, _ = row[:11]
    eco[str(b)].append([ecos, sym or '', float(geo or 0)]); cat[str(b)] = float(area_cat)
ECO = {}
for b, lst in eco.items():
    tg = sum(x[2] for x in lst)
    items = sorted([[e, s, round(g / tg * cat[b], 3), round(g / tg * 100, 2)] for e, s, g in lst], key=lambda x: -x[2])
    ECO[b] = items
wb2 = openpyxl.load_workbook(f'{REPO}/mapas/salidas/IN_Piura_Bloques_AI_por_Distrito_V6.xlsx', data_only=True)
AI = collections.defaultdict(float); AIN = collections.Counter()
for row in wb2['AI_Poligonos'].iter_rows(min_row=8, values_only=True):
    if not row or not isinstance(row[0], (int, float)) or row[2] is None: continue
    AI[str(row[2])] += float(row[6] or 0); AIN[str(row[2])] += 1
# ---- pendientes ----
PEND = {r['cod']: r for r in csv.DictReader(open(f'{S}/work/pendientes_v6.csv'))}

for k, prov in PROVS.items():
    d = json.load(open(f'{S}/work/ds_{k}.json'))
    secs, tsh = ds_tables(d)
    for key in secs:  # descarta el encabezado de hoja sin tabla
        secs[key] = [t for t in secs[key] if t['h']]
    rs = resumen(d)
    T = sanitize(tsh)
    blocks = [r['cod'] for r in PEND.values() if r['prov'] == prov]
    eco_p = {b: {'e': ECO.get(b, []), 'ai': round(AI.get(b, 0), 3), 'nai': AIN.get(b, 0)} for b in blocks}
    cols = ['area', 'este', 'norte', 'media', 'mediana', 'p90', 'max', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'ha75']
    pend_p = {b: [PEND[b]['dist'], PEND[b]['mc']] + [float(PEND[b][c]) for c in cols] for b in blocks}
    U = {'prov': prov, 'corte': '06/10/2026', 'res': rs, 'ds': secs, 'T': T, 'eco': eco_p, 'pend': pend_p}
    json.dump(U, open(f'{S}/work/upd_{k}.json', 'w'), ensure_ascii=False, separators=(',', ':'))
    print(k, prov, len(blocks), 'bloques', {kk: len(v) for kk, v in secs.items()}, list(T.keys()), len(json.dumps(U, ensure_ascii=False)) // 1024, 'KB')
    missing = [b for b in blocks if b not in ECO]
    if missing: print('  sin ecosistema:', missing)
