import json
src = open('mor_orig.html').read()
lines = src.split('\n')
D = json.load(open('mor_D_new.json'))
C = D['C']
assert lines[957].startswith('const D=')
lines[957] = 'const D=' + json.dumps(D, ensure_ascii=False, separators=(',', ':')) + ';'
s = '\n'.join(lines)

def rep(a, b, multi=False):
    global s
    n = s.count(a)
    assert n >= 1, ('NO ENCONTRADO', a[:90])
    assert multi or n == 1, ('MULTIPLE', n, a[:90])
    s = s.replace(a, b)

rep('Revisión 4 · Cortes: DT 20/09/2026 · DS e hidroclima 23/09/2026', 'Revisión 5 · Cortes: DT 20/09/2026 · hidroclima 23/09/2026 · DS 26/09/2026')
rep('Resultado de depurar y consolidar por localidad las fichas F-DS del aplicativo IN Piura (exportación por bloque del 23/09/2026) y cruzarlas con el catálogo INEI–bloques y el libro consolidado.',
    'Resultado de depurar y consolidar por localidad las fichas F-DS del aplicativo IN Piura (respaldo completo y consolidado por provincia del 26/09/2026) y cruzarlas con el catálogo INEI–bloques.')
rep('<span><i style="background:var(--line-2)"></i>Registros únicos exportados</span>', '<span><i style="background:var(--line-2)"></i>Registros repetidos retirados</span>')
rep('No existe muestra programada en la exportación: el avance se mide contra los centros poblados asociados a los bloques.',
    'No existe muestra programada en el aplicativo: el avance se mide contra los centros poblados asociados a los bloques.')
rep('Las localidades sin cifra INEI (Buenos Aires capital, Serrán, La Alberca, La Pilca, Ingenio de Buenos Aires, Huasimal) no se grafican.',
    'Las localidades sin cifra INEI (Buenos Aires capital, Serrán, La Alberca, La Pilca, Ingenio de Buenos Aires, Huasimal, Quemazón) no se grafican.')
rep("<span class=\"cp\">Fichas depuradas F-DS-01/02/03: ${l.n01} / ${l.n02} / ${l.n03}${l.nrep?` · ${l.nrep} reediciones retiradas`:''}",
    "<span class=\"cp\">Fichas depuradas F-DS-01/02/03: ${l.n01} / ${l.n02} / ${l.n03}${l.nrep?` · ${l.nrep} registros repetidos retirados`:''}${l.ninf01>1?` · ${l.ninf01} informantes F-DS-01`:''}")
rep("['Menores de 18 / mayores de 65',(l.men18==null&&l.may65==null)?sd(null):`${n0(l.men18)} / ${n0(l.may65)}`]])}</div>",
    "['Menores de 18 / mayores de 65',(l.men18==null&&l.may65==null)?sd(null):`${n0(l.men18)} / ${n0(l.may65)}`],['Tenencia de la tierra',sd(l.tenencia)+(l.ptit!=null?` <span class=\"muted sm\">(${l.ptit} % titulado)</span>`:'')],['Programas sociales',l.progsoc&&Object.keys(l.progsoc).length?esc(Object.entries(l.progsoc).map(([k,v])=>k+' '+n0(v)).join(' · ')):sd(null)]])}</div>")
rep("{h:'Duración (min)',k:'min',f:r=>r.min===1?'<span class=\"pill p-adv\">1 · por verificar</span>':sd(r.min)}",
    "{h:'Duración (min)',k:'min',f:r=>r.min===1?'<span class=\"pill p-adv\">1 · por verificar</span>':(r.min==null?sd(null):n0(r.min))}")
rep('exportación por bloque del aplicativo IN Piura, fichas F-DS-01, F-DS-02 y F-DS-03 (corte 23/09/2026) · libro consolidado DS del aplicativo (actividades, actores, entrevistas)',
    'respaldo completo del aplicativo IN Piura, fichas F-DS-01, F-DS-02 y F-DS-03 (corte 26/09/2026) · consolidado DS de la provincia (26/09/2026)')
rep('<span class="tag">Libro consolidado · sin nombres</span>', '<span class="tag">F-DS-03 · sin nombres</span>')
rep('Fuente: hoja de actividades económicas del libro consolidado del aplicativo.', 'Fuente: numeral 6 de las F-DS-01 depuradas (respaldo del aplicativo, 26/09/2026).')
# Integralidad
dep = sum(C['CDEP'].values())
rep('["Diagnóstico social","120","fichas depuradas · 49 de 53 actores a favor","--st-soc"]',
    '["Diagnóstico social","%d","fichas depuradas · %d de %d actores a favor","--st-soc"]' % (dep, C['POS']['A favor del proyecto'], C['NLACT']))
rep('"dsTxt":"120 fichas F-DS depuradas de 150 registros, con ficha propia en 46 de 59 bloques"',
    '"dsTxt":"%d fichas F-DS depuradas de 207 registros, con ficha propia en %d de 59 bloques (corte 26/09/2026)"' % (dep, C['CTRL']['bl_ficha']))
rep('"Se aplicaron 3 de 7 instrumentos (F-DS-01 a F-DS-03); 7 bloques sin ficha propia.',
    '"Se aplicaron 3 de 7 instrumentos (F-DS-01 a F-DS-03); %d bloques sin ficha propia al 26/09/2026 (Dótor y Nueva Esperanza en 85–86).' % (59 - C['CTRL']['bl_ficha']))
open('mor_new.html', 'w').write(s)
print('ok')
