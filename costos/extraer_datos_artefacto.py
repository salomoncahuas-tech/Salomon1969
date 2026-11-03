"""Extrae los datos de la Matriz de Costos Paramétricos v1.6 para el artefacto (DATA.cost)."""
import json, sys, importlib.util
import openpyxl
SP='/tmp/claude-0/-home-user-Salomon1969/8a54d1d3-2d8d-5efa-bba5-79ed50668a60/scratchpad'
spec=importlib.util.spec_from_file_location('o','/home/user/Salomon1969/costos/obras_planos.py'); o=importlib.util.module_from_spec(spec); spec.loader.exec_module(o)
obras=o.load()
old=json.loads(open(SP+'/build/old_cost.json').read())
w100=openpyxl.load_workbook(SP+'/recalc1/mcp_v16_raw.xlsx',data_only=True)
w120=openpyxl.load_workbook(SP+'/recalc120/v16_j120.xlsx',data_only=True)
v15_100=openpyxl.load_workbook('/home/user/Salomon1969/Costos_Parametros_IN_Piura_v1.5.xlsx',data_only=True)
v15_120=openpyxl.load_workbook(SP+'/recalc120/v15_j120.xlsx',data_only=True)
def part(wb,key):
    ws=wb['PARTIDAS']
    for r in range(7,ws.max_row+1):
        if ws.cell(r,1).value==key: return [ws.cell(r,c).value for c in (8,9,10)], ws.cell(r,4).value, ws.cell(r,5).value
    raise KeyError(key)
def P(key,src='v16'):
    a,d,u=part(w100 if src=='v16' else v15_100,key); b,_,_=part(w120 if src=='v16' else v15_120,key)
    return dict(h=[a[0],b[0]],nc=[a[1],b[1]],c=[a[2],b[2]],d=d,u=u)
cu={}
def comp(wb):
    ws=wb['CU_COMP']; out={}
    for r in range(60,ws.max_row+1):
        n=ws.cell(r,2).value
        if isinstance(n,str) and n.startswith('CU_') and ws.cell(r,8).value is not None: out[n[3:]]=(ws.cell(r,8).value,ws.cell(r,9).value,ws.cell(r,10).value,ws.cell(r,3).value)
    return out
c1,c2=comp(w100),comp(w120)
for k in c1:
    cu[k]=dict(h=[c1[k][0],c2[k][0]],nc=[c1[k][1],c2[k][1]],c=[c1[k][2],c2[k][2]],desc=c1[k][3])
zp=[dict(k='ZANJAS|01.04',m=1,f='',st='inc'),dict(k='ZANJAS|03.01.01',m=1,f='v',st='inc'),dict(k='ZANJAS|03.01.02',m=600,f='t',st='inc'),
    dict(k='ZANJAS|03.01.03',m=100,f='v',st='inc'),dict(k='ZANJAS|03.01.04',m=25,f='v',st='des',id='elimZ')]
tp=[dict(k='TERRAZAS|01.04',m=1,f='',st='inc'),dict(k='TERRAZAS|03.01.01',m=1150,f='',st='inc'),dict(k='TERRAZAS|03.01.02',m=1100,f='',st='des',id='rell'),
    dict(k='TERRAZAS|03.01.03',m=240,f='',st='eli',id='pirc',src='v15'),dict(k='TERRAZAS|03.01.04',m=62.5,f='',st='des',id='elimT')]
for L in (zp,tp):
    for x in L: x.update(P(x['k'],x.get('src','v16')))
# O&M
ws=w100['OYM']; ws2=w120['OYM']; oym=[]
for i,r in enumerate(range(7,21)):
    od=old['oym'][i]; assert od['act']==ws.cell(r,2).value, (od['act'],ws.cell(r,2).value)
    od=dict(od); od['f100']=ws.cell(r,6).value; od['f120']=ws2.cell(r,6).value
    if ws.cell(r,7).value==0: od['st']='des'; od['id']='podas' if od['act'].startswith('Podas') else 'fert'; od['anios0']=old['oym'][i]['anios']
    oym.append(od)
# parámetros
par=dict(old['par'])
ps=w100['PARAMETROS']
for r in range(50,58):
    n=ps.cell(r,1).value
    if n and n.startswith('P_'): par[n]=ps.cell(r,3).value
parT=[]
for r in range(7,58):
    a,b,v,u,f=[ps.cell(r,c).value for c in range(1,6)]
    if a is None: continue
    if b is None and v is None: parT.append(dict(sec=a)); continue
    if not str(a).startswith('P_'): continue
    parT.append(dict(n=a,p=b,v=v,u=u,f=f))
# leeme / lectura
lm=w100['LEEME']; leeme=[[lm.cell(r,1).value,lm.cell(r,2).value] for r in range(7,22) if lm.cell(r,1).value]
rs=w100['RESUMEN']; lect=[rs.cell(r,1).value for r in range(47,53)]
# desestimadas (tabla final de CU_COMP)
cc=w100['CU_COMP']; des=[]; start=None
for r in range(1,cc.max_row+1):
    if cc.cell(r,1).value=='Tratamiento': start=r+1
for r in range(start,cc.max_row+1):
    if cc.cell(r,2).value: des.append(dict(t=cc.cell(r,1).value,k=cc.cell(r,2).value,d=cc.cell(r,3).value,u=cc.cell(r,4).value,m=cc.cell(r,5).value,cu=cc.cell(r,6).value,km=cc.cell(r,7).value,s=cc.cell(r,8).value))
# bloques con metas del plano
blocks=[]
for b in old['blocks']:
    m=obras[b['id']]
    b=dict(b); b.update(zi=m['zi_m']/1000,tfl=m['tfl_m']/1000,der=m['deriv_m']/1000,av=1 if m['aviso'] else 0,cl=m['clase'],pp=m['P'],etp=m['ETP'],ia=m['IA'],bv=m['barrera'],
                        obra=m['obra'],s27=m['sup27'],supP=m['sup'],tz=m['tfl_zona'],zz=m['zi_zona'],tt=m['tfl_traz'],te=m['tfl_est'],zt=m['zi_traz'],ze=m['zi_est'],gr=m['grad'],hoja=list(obras).index(b['id'])+2)
    blocks.append(b)
ret=[dict(id=k,hoja=list(obras).index(k)+2,dist=obras[k]['dist'],sup=obras[k]['sup'],tfl=obras[k]['tfl_m']/1000,zi=obras[k]['zi_m']/1000) for k in o.RETIRADOS]
# totales de control del Excel
rs=w100['RESUMEN']
ctrl=dict(inv=rs['C18'].value,cd=rs['C17'].value,om1=rs['C22'].value,om15=rs['C23'].value,tot=rs['C24'].value,margen=rs['C30'].value,kz=rs['C13'].value,kt=rs['C14'].value,
          bv=w100['CALCULO']['AL124'].value,bvt=w100['CALCULO']['AJ124'].value,bvz=w100['CALCULO']['AK124'].value,vacS=rs['C27'].value)
cost=dict(old); cost.update(cu=cu,zp=zp,tp=tp,oym=oym,par=par,parT=parT,leeme=leeme,lect=lect,des=des,blocks=blocks,ret=ret,ctrl=ctrl,ver='1.6',
                            bvref=o.BARRERA,fuente='Costos_Parametros_IN_Piura_v1.6.xlsx')
json.dump(cost,open(SP+'/build/cost16.json','w'),ensure_ascii=False,separators=(',',':'))
print('cu',{k:round(v['h'][0],2) for k,v in cu.items()}); print('ctrl',{k:round(v,2) if isinstance(v,float) else v for k,v in ctrl.items()})
print('zp',[(x['k'],round(x['h'][0],3)) for x in zp]); print('tp',[(x['k'],round(x['h'][0],3),round(x['h'][1],3)) for x in tp])
print('des',len(des),'leeme',len(leeme),'parT',len(parT), [p['n'] for p in parT if p.get('n','').startswith('P_Z') or p.get('n') in ('P_ARIDO','P_BVIVA','P_DERIV')])
