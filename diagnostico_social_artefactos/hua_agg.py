import json
exec(open('dsbuild.py').read()); exec(open('aggb.py').read())
D=json.load(open('hua_D_new.json'))
h=ds[ds.provincia=='Huancabamba']
rb={};rf={}
for _,r in h.iterrows():
    rb.setdefault(str(r.bloque),[]).append(r.to_dict()); rf.setdefault(r.ficha,[]).append(r.to_dict())
C=recalcular(D,rb,rf)
LOC=D['LOC']; B=D['B']
LF=[l for l in LOC if l['nreg']]
cpf=sum(len(l['inei']) for l in LF)
print('CREG',C['CREG'],'CDEP',C['CDEP'],'NF01',C['NF01'],'NLACT',C['NLACT'],'N03',C['N03'],'NL03',C['NL03'],'ACU',C['ACU'])
print('POS',C['POS'],'GEN',C['GEN'],'FRAG',C['FRAG'],'RES',C['RES'],'RATIO',C['RATIO'])
print('ACTS',C['ACTS']);print('FAM',C['FAM']);print('HZ',C['HZ']);print('SPP',C['SPP']);print('BEN',C['BEN'])
print('locs_ficha',len(LF),'cp_ficha',cpf,'pob_ficha',sum(l['pob'] for l in LF))
print('bl_ficha',sum(1 for b in B if b['ds']['caract']=='Con ficha propia'),'bl_via',sum(1 for b in B if b['ds']['caract'].startswith('Localidad')))
print('sincar',[b['b'] for b in B if b['ds']['caract'].startswith('Sin')])
from collections import Counter
print(Counter(b['d'] for b in B if b['ds']['caract']=='Con ficha propia'))
print([ (b['b'],b['ds']['n01'],b['ds']['n02'],b['ds']['n03'],b['ds']['d01'],b['ds']['d02'],b['ds']['d03'],b['ds']['frag'],b['ds']['res']) for b in B if b['ds']['caract']!='Sin ficha en el aplicativo'])
json.dump(D,open('hua_D_new.json','w'),ensure_ascii=False)
