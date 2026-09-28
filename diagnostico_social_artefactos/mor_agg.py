import json
exec(open('mor_rebuild.py').read().split("if __name__")[0])
exec(open('aggb.py').read())
rb={};rf={}
for _,r in m.iterrows():
    rb.setdefault(str(r.bloque),[]).append(r.to_dict()); rf.setdefault(r.ficha,[]).append(r.to_dict())
C=recalcular(D,rb,rf,via='Localidad caracterizada vía otro bloque',sin='Sin caracterización social',extra=D['BL'])
for x in D['BL']:
    for k in ('frag','res','caract'): x.pop(k,None)
B=D['B'];LOC=D['LOC']
LF=[l for l in LOC if l['nreg']]
V6s={b['b'] for b in B}
LV6=[l for l in LOC if any(x in V6s for x in l['bl'])]
LFV6=[l for l in LV6 if l['nreg']]
print('CREG',C['CREG'],'CDEP',C['CDEP'],'NF01',C['NF01'],'NLACT',C['NLACT'],'N03',C['N03'],'NL03',C['NL03'],'ACU',C['ACU'])
print('POS',C['POS'],'GEN',C['GEN'],'FRAG',C['FRAG'],'RES',C['RES'],'RATIO',C['RATIO'])
print('HZ',C['HZ']);print('SPP',C['SPP']);print('BEN',C['BEN']);print('ACTS',C['ACTS'])
from collections import Counter
print(Counter(b['ds']['caract'] for b in B))
print('sin',[b['b'] for b in B if b['ds']['caract'].startswith('Sin')],'via',[b['b'] for b in B if b['ds']['caract'].startswith('Loc')])
print('locs_v6',len(LV6),'pob6',sum(l['pob'] or 0 for l in LV6),'locs_car',len(LFV6),'pob_car',sum(l['pob'] or 0 for l in LFV6),'LF',len(LF))
json.dump(D,open('mor_D_new.json','w'),ensure_ascii=False)
