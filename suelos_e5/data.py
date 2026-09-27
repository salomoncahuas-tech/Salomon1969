import re, json, shapefile, warnings, statistics as st
warnings.filterwarnings('ignore')
import colour, numpy as np
S='/tmp/claude-0/-home-user-Salomon1969/232a4de4-cdc3-5473-be12-eb7b421c81bc/scratchpad/'
TXT=open(S+'suelos/informe.txt',encoding='utf8').read().split('\n')

def table_after(prefix):
    i=next(k for k,l in enumerate(TXT) if l.startswith(prefix))
    while not TXT[i].startswith('|'): i+=1
    rows=[]
    while i<len(TXT) and TXT[i].startswith('|'):
        cells=[c.strip().replace('\\<','<') for c in TXT[i].strip().strip('|').split('|')]
        rows.append(cells); i+=1
    return rows[0],rows[2:]

C=[x.as_dict() for x in shapefile.Reader(dbf=open(S+'suelos/calicatas.dbf','rb')).records()]
BL=[x.as_dict() for x in shapefile.Reader(dbf=open(S+'suelos/bloques.dbf','rb')).records()]

# Cuadro 13 morfología superficial
h,r13=table_after('Cuadro 13.'); M13={r[0]:r for r in r13}
# Cuadro 20 códigos Factor K
h,r20=table_after('Cuadro 20.'); M20={r[0]:r for r in r20}
# Cuadro 18 repeticiones, 19 sitios, 17 pruebas
h,r18=table_after('Cuadro 18.'); h,r19=table_after('Cuadro 19.'); h,r17=table_after('Cuadro 17.')
h,r8=table_after('Cuadro 8.')
# Anexo A chequeos
h,rA=table_after('Cuadro 23.')

def munsell_hex(m):
    try:
        xyY=colour.munsell_colour_to_xyY(m.replace(' ',' ').strip())
        rgb=colour.XYZ_to_sRGB(colour.xyY_to_XYZ(xyY),illuminant=colour.CCS_ILLUMINANTS['CIE 1931 2 Degree Standard Observer']['C'])
        rgb=np.clip(rgb,0,1); return '#%02x%02x%02x'%tuple(int(round(v*255)) for v in rgb)
    except Exception as e:
        print('munsell fail',m,e); return None

def parse_samples(s):
    s=s.replace('–','-')
    out=[]
    m=re.search(r':\s*(.*)cm',s)
    if m:
        for part in m.group(1).split(','):
            a=re.findall(r'(\d+)\s*-\s*(\d+)',part)
            if a: out.append([int(a[0][0]),int(a[0][1])])
    else:
        for a,b in re.findall(r'\((\d+)-(\d+) cm\)',s): out.append([int(a),int(b)])
    return out

def pend_num(p):
    m=re.match(r'\((\d+)(?:\s*-\s*(\d+))?\s*%',p.replace('–','-'))
    if not m: return None
    return float(m.group(1)) if not m.group(2) else None

ERO={'Muy ligera':1,'Ligera':2,'Moderada':3,'Severa':4}
cals=[]
for c in C:
    k=c['CODIGO']; m=M13[k]; kk=M20[k]
    sam=parse_samples(c['MUESTRAS'])
    mun=m[3]
    cals.append(dict(
        c=k,t={'Bloque':'B','Referencia':'R','Lote SUS':'S'}[c['TIPO']],par=c['PAR'] or None,eco=c['ECOSISTEMA'],zh=c['ZONA_HOMOG'],
        b=c['BLOQUE'],us=c['ID_US'],u=c['UNIDAD_SUE'],lug=c['LUGAR'],fe=c['FECHA'].strftime('%d/%m'),br=c['BRIGADA'],
        e=int(c['ESTE']),n=int(c['NORTE']),z=int(c['ALTITUD']),pos=c['POSICION'],pend=c['PENDIENTE'],pn=pend_num(c['PENDIENTE']),pc=c['PEND_CLASE'],
        mic=c['MICRORELIE'],uso=c['USO_ACTUAL'],veg=c['VEGETACION'],mat=c['MAT_PARENT'],ped=c['PEDREGOSID'],ero=c['EROSION'],eroN=ERO.get(c['EROSION']),
        dren=c['DRENAJE'],pe=c['PROF_EFECT'],pcm=c['PROF_CM'] or None,con=c['CONTACTO'],ccm=c['CONTACTO_C'] or None,hz=c['HORIZONTES'].split('-'),
        epi=c['EPIPEDON'],sub=c['HZ_SUBSUP'],tax=c['SOIL_TAXON'],wrb=c['WRB'],g=c['GRUPO_SUEL'],cum=c['CUM_CAMPO'],cumu=c['CUM_UNIDAD'],
        nm=c['N_MUESTRAS'],mu=c['MUESTRAS'],sam=sam,ep=int(c['ESTE_PLAN']),np_=int(c['NORTE_PLAN']),dz=int(c['DESPLAZ_M']),mot=c['MOTIVO_REU'],obs=c['OBSERVACIO'],
        hs=m[2],mun=mun,col=munsell_hex(mun),tex=m[4],est=m[5],hcl=m[6],frag=m[7],rai=m[8],
        ke=int(kk[4]),kp=int(kk[5])))
# chequeos
chk=[]
for r in rA:
    chk.append(dict(c=r[0],cal=r[0][:6],fe=r[1],br=r[2],lug=r[3],e=int(r[4]),n=int(r[5]),z=int(r[6]),pc=r[7],mat=r[8],pe=r[9],con=r[10],hz=r[11],dg=r[12].strip()))
# infiltración
site={}
for r in r19:
    site[r[0]]=dict(id=r[0],sit=r[1],cal=r[2],u=r[3],nr=int(r[4]),ruh=r[5],rib=r[6],ib=float(r[7]),cl=r[8])
for r in r17:
    s=site[r[0]]; s.update(plan=r[1],sitio=r[2],fe=r[5],e=int(r[6].replace(' ','')),n=int(r[7].replace(' ','')),z=int(r[8]))
for r in r8:
    pass
reps=[]
for r in r18:
    reps.append(dict(i=r[0],r=r[1],d=int(r[2]),L=float(r[3]),vu=float(r[4]),k=float(r[5]),a=float(r[6]),r2=float(r[7]),tb=int(r[8]),ib=float(r[9]),cl=r[10]))
# SHP median IB for QC
shp_ib={'INF-01':0.4142,'INF-02':0.7553,'INF-02b':2.1672,'INF-03':0.4476,'INF-04':0.6849,'INF-05':0.3187,'INF-06':1.6603,'INF-07':0.8690,'INF-08':13.714,'INF-09':0.2918,'INF-10':2.0125,'INF-11':73.48,'INF-12':0.9680}
for k,v in shp_ib.items(): site[k]['ibshp']=v
# bloques
blocks={}
for b in BL:
    u=[x.strip() for x in b['UNIDADES_S'].split(',') if x.strip()]
    blocks[b['BLOQUE']]=dict(p=b['PROVINCIA'].title().replace('Morropon','Morropón'),d=b['DISTRITO'].title().replace(' De ',' de ').replace('Frias','Frías').replace('Morropon','Morropón'),ha=b['HA'],eco=b['ECOSISTEMA'],ti=b['TIPO_INTER'],zv=b['ZONA_VIDA_'],sp=b['SUBPAISAJE'],pd=b['PEND_DOM'],us=u,cal=[x for x in b['CALICATAS'].split(',') if x.strip()],inf=b['PRUEBA_INF'] or None,er=b['ER_PAREADA'] or None)
units={}
for c in cals:
    units.setdefault(c['us'],{'s':c['u'],'c':[]})['c'].append(c['c'])
D=dict(C=cals,K=chk,I=list(site.values()),R=reps,BL=blocks,U=units)
json.dump(D,open(S+'build/sue.json','w'),ensure_ascii=False,separators=(',',':'))
print(len(json.dumps(D,ensure_ascii=False,separators=(',',':'))))
print(len(cals),len(chk),len(site),len(reps),len(blocks),len(units))
print({c['c']:c['col'] for c in cals if not c['col']})
for c in cals:
    if len(c['sam'])==0: print('nosam',c['c'])
