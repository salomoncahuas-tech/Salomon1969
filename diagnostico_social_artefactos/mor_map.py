import json
exec(open('dsbuild.py').read())
D=json.load(open('mor_D.json'))
LL={l['loc']:l for l in D['LOC']}
ALIAS={('Buenos Aires','Buenos Aires'):'Buenos Aires (capital distrital)','Rio Seco':'Río Seco','San Lorenzo':'San Lorenzo / Taspa','Taspa':'San Lorenzo / Taspa',
'Rio Seco Alto':'Río Seco Alto','Mangomanguia':'Mangomanguía','Palo Blanco':'Palo Blanco–El Cerezo','Palo Blanco-el Cerezo':'Palo Blanco–El Cerezo',
('Salitral','Santa Rosa'):'Santa Rosa (Salitral)','Serran':'Serrán','Tortola':'Tórtola','Victor Raul':'Víctor Raúl','Alan Garcia':'Alan García / Bigote','Bigote':'Alan García / Bigote',
'Bado de Garzas':'Manzanares / Bado de Garzas','Manzanares':'Manzanares / Bado de Garzas','Polluco':'Polluco / Sinaí','Sinai':'Polluco / Sinaí','Quemazon':'Quemazón',
('San Juan de Bigote','Santa Rosa'):'Santa Rosa (San Juan de Bigote)','La Laja':'La Laja / Overazal','Jacanacas / Nueva Esperanza':'Nueva Esperanza (Santo Domingo)',
('Santo Domingo','Nueva Esperanza'):'Nueva Esperanza (Santo Domingo)',('Yamango','Nueva Esperanza'):'Nueva Esperanza / Mambluque / Alto Mambluque','Mambluque':'Nueva Esperanza / Mambluque / Alto Mambluque',
'Alto Mambluque':'Nueva Esperanza / Mambluque / Alto Mambluque','Rircardo Palma':'Ricardo Palma','Victor Raul (el Checo)':'Víctor Raúl (El Checo)'}
def locname(r):
    if not isinstance(r['centro_poblado'],str): return {'M10B4':'Río Seco Alto'}[str(r['bloque'])]
    k=(r['distrito'],r['centro_poblado'])
    a=ALIAS.get(k) or ALIAS.get(r['centro_poblado'])
    if a: return a
    NN={norm(x):x for x in LL}
    return NN.get(norm(r['centro_poblado']), r['centro_poblado'])
m=ds[ds.provincia=='Morropon']
recs_por={}
for _,r in m.iterrows():
    r=r.to_dict(); n=locname(r); assert n in LL,(n,r['distrito'])
    recs_por.setdefault(n,[]).append(r)
