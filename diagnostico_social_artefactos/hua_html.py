import json, re
src = open('hua_orig.html').read()
lines = src.split('\n')
D = json.load(open('hua_D_new.json'))
assert lines[1017].startswith('const D=')
lines[1017] = 'const D=' + json.dumps(D, ensure_ascii=False, separators=(',', ':')) + ';'
s = '\n'.join(lines)

def rep(a, b, count=1):
    global s
    n = s.count(a)
    assert n >= 1, ('NO ENCONTRADO', a[:80])
    if count == 1:
        assert n == 1, ('MULTIPLE', n, a[:80])
    s = s.replace(a, b)

# Cabecera
rep('Revisión integrada DT + DS · Cortes: DT 21/09/2026 · DS e hidroclima 23/09/2026',
    'Revisión integrada DT + DS · Cortes: DT 21/09/2026 · hidroclima 23/09/2026 · DS 26/09/2026')
rep("['Fichas sociales',CT.nreg,`${CT.ndep} depuradas · 13 bloques`],['Actores a favor',12,'de 12']",
    "['Fichas sociales',CT.nreg,`${CT.ndep} depuradas · ${CT.bl_ficha} bloques`],['Actores a favor',C.POS['A favor del proyecto']||0,`de ${C.NLACT}`]")
# Hallazgos
rep('(consolidado y exportación por bloque del 23/09/2026)', '(consolidado y exportación por bloque del 26/09/2026)')
# Cobertura
rep('<span class="tag">Hoja de control 17/09 · aplicativo 23/09</span>', '<span class="tag">Hoja de control 17/09 · aplicativo 26/09</span>')
rep("reporta <strong>${CT.ctrl01}</strong> aplicadas y ${CT.ctrl03} F-DS-03; el aplicativo registra 13 y 4.",
    "reporta <strong>${CT.ctrl01}</strong> aplicadas y ${CT.ctrl03} F-DS-03; el aplicativo registra ${C.CREG['F-DS-01']||0} y ${C.CREG['F-DS-03']||0} al ${CT.corte} (13 y 4 al 23/09).")
# Localidades
rep('Yumbe / Cruz Roja tiene ficha pero sin datos demográficos y no se grafica. Fuente: F-DS-01 depuradas (aplicativo IN Piura, 23/09/2026) y catálogo INEI–bloques v5.',
    'Yumbe / Cruz Roja tiene ficha pero sin datos demográficos, y Maray Grande (urbano) no tiene cifra INEI ni de control: no se grafican. Fuente: F-DS-01 depuradas (aplicativo IN Piura, 26/09/2026) y catálogo INEI–bloques v5.')
# Ficha por localidad: informantes, programas sociales, tenencia
rep("<span class=\"cp\">Fichas depuradas F-DS-01/02/03: ${l.n01} / ${l.n02} / ${l.n03}${l.nrep?` · ${l.nrep} reediciones retiradas`:''}",
    "<span class=\"cp\">Fichas depuradas F-DS-01/02/03: ${l.n01} / ${l.n02} / ${l.n03}${l.nrep?` · ${l.nrep} registros repetidos retirados`:''}${l.ninf01>1?` · ${l.ninf01} informantes F-DS-01`:''}")
rep("['Menores de 18 / mayores de 65',(l.men18==null&&l.may65==null)?sd(null):`${n0(l.men18)} / ${n0(l.may65)}`]])}</div>",
    "['Menores de 18 / mayores de 65',(l.men18==null&&l.may65==null)?sd(null):`${n0(l.men18)} / ${n0(l.may65)}`],['Tenencia de la tierra',sd(l.tenencia)+(l.ptit!=null?` <span class=\"muted sm\">(${l.ptit} % titulado)</span>`:'')],['Programas sociales',l.progsoc&&Object.keys(l.progsoc).length?esc(Object.entries(l.progsoc).map(([k,v])=>k+' '+n0(v)).join(' · ')):sd(null)]])}</div>")
# Servicios y economía
rep('Café, cacao, caña, maíz, frijol y plátano son los productos declarados.',
    'Café, cacao, caña, maíz, frijol, plátano, yuca, arroz y soya son los productos declarados.')
# Actores
rep("Los <strong>${A.length} actores</strong> registrados en F-DS-02 son de nivel comunal: 6 líderes o autoridades comunales (tenientes gobernadores), 2 agentes municipales, 2 JASS y 2 rondas campesinas; todos declaran posición <strong>a favor</strong> e influencia e interés «Alto» (D-HS05).",
    "Los <strong>${A.length} actores</strong> registrados en F-DS-02 son de nivel comunal: ${Object.entries(A.reduce((o,a)=>(o[a.tipo]=(o[a.tipo]||0)+1,o),{})).sort((x,y)=>y[1]-x[1]).map(([k,v])=>v+' '+k.toLowerCase().replace('líder / autoridad comunal','líderes o autoridades comunales').replace('gobierno local (municipalidad)','agentes municipales').replace('jass (saneamiento)','JASS').replace('ronda campesina','rondas campesinas')).join(', ')}; todos declaran posición <strong>a favor</strong> e influencia e interés «Alto» (D-HS05).")
# Percepción
rep('<span class="tag">F-DS-03 · codificación propia · 4 localidades</span>', '<span class="tag">F-DS-03 · codificación propia · 8 localidades</span>')
rep('Lo que sigue codifica las respuestas de las 4 entrevistas F-DS-03 (numerales 1.1 a 1.4 y 4.2), todas en Lalaquiz; cada categoría se cuenta una vez por localidad.',
    'Lo que sigue codifica las respuestas de las 8 entrevistas F-DS-03 (numerales 1.1 a 1.4 y 4.2), 6 en Lalaquiz y 2 en Canchaque; cada categoría se cuenta una vez por localidad.')
rep("lab:`${v} de ${C.NL03}`})),{lw:290,rw:90,axis:true,max:4});", "lab:`${v} de ${C.NL03}`})),{lw:290,rw:90,axis:true,max:C.NL03});")
rep('<span class="tag">4 de 4 entrevistas de acuerdo</span>', '<span class="tag">8 de 8 entrevistas de acuerdo</span>')
rep("const proc=/Derrumbes|quebradas|Huaicos|Erosión|agua/;", "const proc=/Derrumbes|quebradas|Huaicos|Erosión|agua|Sequía/;")
# Pie
rep('fichas F-DS-01, F-DS-02 y F-DS-03 (corte 23/09/2026)', 'fichas F-DS-01, F-DS-02 y F-DS-03 (corte 26/09/2026)')
# Integralidad
rep('["Diagnóstico social","29","registros F-DS · 12 de 12 actores a favor","--st-soc"]',
    '["Diagnóstico social","49","registros F-DS · 14 de 14 actores a favor","--st-soc"]')
rep('"dsTxt":"29 registros F-DS (22 depurados) con cobertura en 13 de 48 bloques y 16 de 58 C.P."',
    '"dsTxt":"49 registros F-DS (33 depurados) con cobertura en 14 de 48 bloques y 18 de 58 C.P. (corte 26/09/2026)"')
rep('"Se aplicaron 3 de 7 instrumentos (F-DS-01 a F-DS-03); fichas en 13 de 48 bloques. La percepción de peligros y cambio climático (F-DS-06) no se aplicó."',
    '"Se aplicaron 3 de 7 instrumentos (F-DS-01 a F-DS-03); fichas en 14 de 48 bloques al 26/09/2026, sin ninguna en el núcleo de brecha ni en el distrito de Huancabamba. La percepción de peligros y cambio climático (F-DS-06) no se aplicó."')
open('hua_new.html', 'w').write(s)
print('ok', len(s))
