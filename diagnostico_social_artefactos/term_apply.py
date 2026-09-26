import json, re
exec(open('term.py').read())
exec(open('aya_rec.py').read())

CSS = """
.mrrc{display:inline-block;font:600 10px/1.45 "IBM Plex Mono",monospace;letter-spacing:.04em;text-transform:uppercase;padding:1px 6px;border-radius:2px;margin:0 6px 0 0;white-space:nowrap;vertical-align:1px}
.mrrc.v{background:var(--moss-wash);color:var(--moss)} .mrrc.m{background:var(--ochre-wash);color:var(--ochre)}
.mrrc.c{background:var(--anin-wash);color:var(--anin)} .mrrc.g{background:var(--surface-2);color:var(--ink-2);border:1px solid var(--line)}
.mrrl{list-style:none;margin:0;padding:0;display:grid;gap:6px;font-size:13px;text-align:left}
.fb-row.c4{grid-template-columns:repeat(4,minmax(0,1fr))}
@media (max-width:860px){.fb-row.c4{grid-template-columns:1fr}}
.mrrleg{display:flex;flex-wrap:wrap;gap:4px 12px;font-size:11.5px;color:var(--ink-3);margin:0 0 10px}
"""


def walk(o):
    if isinstance(o, dict):
        return {k: walk(v) for k, v in o.items()}
    if isinstance(o, list):
        return [walk(v) for v in o]
    if isinstance(o, str):
        return gris_a_complementaria(o)
    return o


def data_line(L, idx, prefix):
    s = L[idx]
    assert s.startswith(prefix), (idx, s[:30])
    body = s[len(prefix):].strip().rstrip(';')
    return json.loads(body)


def put(L, idx, prefix, obj, sep=None):
    L[idx] = prefix + json.dumps(obj, ensure_ascii=False, separators=sep or (',', ':')) + ';'


def rep(s, a, b):
    assert s.count(a) == 1, (s.count(a), a[:80])
    return s.replace(a, b)

# ------------------ Huancabamba ------------------
L = open('hua_new.html').read().split('\n')
D = data_line(L, 1017, 'const D=')
for b in D['B']:
    g, t = reagrupar(b.get('s_mrr'))
    assert OTRAS not in g, b['b']
    b['s_mrr'] = t
D = walk(D)
put(L, 1017, 'const D=', D)
s = '\n'.join(L)
s = rep(s, """    <div><p class="fb-txt"><b>Infraestructura natural (verde)</b></p>${lst(mrr[g('infraestructura natural')])}</div>
    <div><p class="fb-txt"><b>Infraestructura gris (complementaria)</b></p>${lst(mrr[g('infraestructura gris')])}</div>
    <div><p class="fb-txt"><b>Gobernanza y gestión comunitaria</b></p>${lst(mrr[g('gobernanza')])}</div>""",
"""    <div><p class="fb-txt"><span class="mrrc v">Verde</span><b>Infraestructura natural verde</b></p>${lst(mrr[g('infraestructura natural verde')])}</div>
    <div><p class="fb-txt"><span class="mrrc m">Marrón</span><b>Infraestructura natural marrón</b></p>${lst(mrr[g('infraestructura natural marr')])}</div>
    <div><p class="fb-txt"><span class="mrrc c">Compl.</span><b>Infraestructura complementaria</b></p>${lst(mrr[g('infraestructura complementaria')])}</div>
    <div><p class="fb-txt"><span class="mrrc g">Gob.</span><b>Gobernanza y gestión comunitaria</b></p>${lst(mrr[g('gobernanza')])}</div>""")
s = rep(s, """<div class="fb-sec"><div class="fb-h">MRR-CCC identificadas<small>orientativas · ${esc(b.metas).toLowerCase()}</small></div>
   <div class="fb-row c3" style="border:0;margin:0 -22px">""",
"""<div class="fb-sec"><div class="fb-h">MRR-CCC identificadas<small>orientativas · ${esc(b.metas).toLowerCase()}</small></div>
   <p class="mrrleg">Verde: revegetación, reforestación, enriquecimiento, regeneración natural, restauración activa · Marrón: zanjas, terrazas, diques de cárcavas, clausura y manejo de pastizales, barreras y cercos vivos · Complementaria: estabilización de taludes, muros, gaviones, diques en quebradas, drenaje</p>
   <div class="fb-row c4" style="border:0;margin:0 -22px">""")
s = gris_a_complementaria(s)
s = s.replace('</style>\n<header class="hero">', CSS + '</style>\n<header class="hero">', 1)
assert CSS in s
open('hua_new.html', 'w').write(s)

# ------------------ Morropón ------------------
L = open('mor_new.html').read().split('\n')
D = data_line(L, 957, 'const D=')
for b in D['B']:
    lst_, t = etiquetar(b.get('mrr'))
    b['mrr_cls'] = [[{VERDE: 'v', MARRON: 'm', COMP: 'c', GOB: 'g'}.get(c, 'g'), x] for c, x in lst_]
    if t:
        b['mrr'] = t
    if b.get('s_mrr'):
        b['s_mrr'] = etiquetar(b['s_mrr'])[1] + ('…' if b['s_mrr'].rstrip().endswith('…') and not etiquetar(b['s_mrr'])[1].endswith('…') else '')
D = walk(D)
put(L, 957, 'const D=', D)
s = '\n'.join(L)
s = rep(s, "<div><h4>MRR-CCC identificadas</h4>${kv([['Medidas',sd(b.s_mrr||b.mrr),1],['Metas físicas',sd(b.metas),1]])}</div>",
    "<div><h4>MRR-CCC identificadas</h4><p class=\"mrrleg\"><span><span class=\"mrrc v\">Verde</span>revegetación, enriquecimiento, regeneración</span><span><span class=\"mrrc m\">Marrón</span>zanjas, terrazas, cárcavas, clausura, cercos vivos</span><span><span class=\"mrrc c\">Compl.</span>taludes, muros, drenaje, diques en quebradas</span><span><span class=\"mrrc g\">Gob.</span>gobernanza</span></p>${b.mrr_cls&&b.mrr_cls.length?'<ul class=\"mrrl\">'+b.mrr_cls.map(([c,t])=>`<li><span class=\"mrrc ${c}\">${({v:'Verde',m:'Marrón',c:'Compl.',g:'Gob.'})[c]}</span>${esc(t)}</li>`).join('')+'</ul>':sd(null)}${kv([['Metas físicas',sd(b.metas),1]])}</div>")
s = gris_a_complementaria(s)
s = s.replace('</style>\n<header class="hero">', CSS + '</style>\n<header class="hero">', 1)
assert CSS in s
open('mor_new.html', 'w').write(s)

# ------------------ Ayabaca ------------------
L = open('aya_new.html').read().split('\n')
X = data_line(L, 946, 'const X = ')
for k, v in REC.items():
    X['PDEC'][k]['rec'] = v
X = walk(X)
L[946] = 'const X = ' + json.dumps(X, ensure_ascii=False) + ';'
DA = data_line(L, 945, 'const DATA = ')
DA = walk(DA)
L[945] = 'const DATA = ' + json.dumps(DA, ensure_ascii=False) + ';'
s = '\n'.join(L)
s = gris_a_complementaria(s)
open('aya_new.html', 'w').write(s)

for f in ('aya_new.html', 'hua_new.html', 'mor_new.html'):
    t = open(f).read()
    print(f, 'gris restantes:', re.findall(r'(?i).{40}\b(?:infraestructura|obras?) gris.{20}', t)[:5], '| (verde):', len(re.findall(r'natural \(verde\)', t)))
