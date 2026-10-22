import json,sys
from collections import Counter
def tab(U,k,pref):
    for t in U['ds'].get(k,[]):
        if t['t'].split('. ',1)[1].startswith(pref): return t
def summ(U):
    o={}
    o['cab']={c[0]:(c[1],c[2]) for c in U['res']['cab']}
    A=tab(U,'cob','Fichas sociales registradas'); o['tipo']={r[0][:8]:r[1] for r in A['r']}
    D=tab(U,'cob','Fichas sociales por bloque'); o['bloq']={str(r[0]):r[1] for r in D['r']}
    E=tab(U,'cob','Fichas sociales por distrito'); o['dist']={r[0]:r[1] for r in E['r']}
    B=tab(U,'cob','Cobertura del diagnóstico social por centro'); o['cp']={r[0]:sum(x or 0 for x in r[1:len(B['h'])-1]) for r in B['r']}
    C=tab(U,'cob','Centros poblados del catálogo'); o['cat']={r[0]:r[1] for r in C['r']}; o['catn']=C['n']
    cal=U['T']['cal']; i=cal['h'].index('Tema'); o['cal']=Counter(r[i] for r in cal['r'])
    o['T']={k:len(v['r']) for k,v in U['T'].items()}
    G=tab(U,'f01','Familias por actividad'); o['act']={r[0]:round(sum(x or 0 for x in r[1:7]),1) for r in G['r']} if G else {}
    o['actd']=G['d'] if G else ''
    ten=U['T']['ten']; o['ten']={r[0]:r[3] for r in ten['r']}
    return o
p=sys.argv[1]
O=summ(json.load(open(f'work/{p}_UPD_old.json'))); N=summ(json.load(open(f'work/{p}_new.json')))
for k in O:
    if O[k]!=N[k]:
        if isinstance(O[k],dict):
            ks=set(O[k])|set(N[k])
            print(k,':',{x:(O[k].get(x),N[k].get(x)) for x in sorted(ks,key=str) if O[k].get(x)!=N[k].get(x)})
        else: print(k,':',O[k],'->',N[k])
print('NEW actd:',N['actd']); print('NEW act:',N['act']); print('catn',N['catn'])
print('cab',N['cab'])
