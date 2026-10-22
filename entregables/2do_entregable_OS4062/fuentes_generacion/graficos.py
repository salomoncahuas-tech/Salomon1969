import json, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import rcParams
S='/tmp/claude-0/-home-user-Salomon1969/03b82ed8-c0e9-54ec-93c9-bff5f93b575b/scratchpad'
d=json.load(open(S+'/build/datos.json'))
rcParams.update({'font.family':'DejaVu Sans','font.size':9,'axes.edgecolor':'#9a9a9a','axes.linewidth':0.6,
  'xtick.color':'#52514e','ytick.color':'#0b0b0b','axes.labelcolor':'#52514e'})
INK='#0b0b0b'; INK2='#52514e'; GRID='#e4e4e0'
CAT=['#2a78d6','#eb6834','#1baf7a']
PROV=['Huancabamba','Morropón']
def fin(fig,ax,nombre):
    for s in ['top','right']: ax.spines[s].set_visible(False)
    ax.grid(axis='x',color=GRID,lw=0.6); ax.set_axisbelow(True)
    fig.savefig(S+f'/build/fig/{nombre}.png',dpi=200,bbox_inches='tight',facecolor='white'); plt.close(fig)

def apilado(nombre, cats, colores, valores, titulo, unidad='ha', etiqueta_min=4.0):
    fig,ax=plt.subplots(figsize=(7.2,2.0))
    for i,p in enumerate(PROV):
        tot=sum(valores[p]); x=0
        for j,c in enumerate(cats):
            v=valores[p][j]; w=100*v/tot
            ax.barh(i,w,left=x,color=colores[j],height=0.55,edgecolor='white',linewidth=1.5)
            if w>=etiqueta_min:
                lum=sum(int(colores[j][k:k+2],16) for k in (1,3,5))/3
                ax.text(x+w/2,i,f'{w:.1f} %',ha='center',va='center',fontsize=8,color='white' if lum<150 else INK)
            x+=w
        ax.text(101,i,f'{tot:,.2f} {unidad}',va='center',fontsize=8,color=INK2)
    ax.set_yticks(range(len(PROV))); ax.set_yticklabels(PROV); ax.invert_yaxis()
    ax.set_xlim(0,100); ax.set_xlabel('% de la superficie de los bloques de la provincia')
    ax.legend([plt.Rectangle((0,0),1,1,color=c) for c in colores],cats,ncol=len(cats),frameon=False,
              loc='upper left',bbox_to_anchor=(0,-0.32),fontsize=8,handlelength=1.2)
    ax.set_title(titulo,loc='left',fontsize=9.5,color=INK,pad=8,fontweight='bold')
    fin(fig,ax,nombre)

P=d['prov']
apilado('fig_predial',['Comunidad campesina (GEORURAL)','Predios COFOPRI fuera de C. C.','Tierras del Estado y/o posesionarios'],CAT,
  {p:[P[p]['pred_cc'],P[p]['pred_cof'],P[p]['pred_est']] for p in PROV},'Situación predial de gabinete de los bloques (mapas SP v2, 11/10/2026)')
SEQ_PI=['#fde3c8','#f9b98a','#f08a4b','#d4561f','#9c2f0e']
apilado('fig_pi',['Muy bajo','Bajo','Medio','Alto','Muy alto'],SEQ_PI,{p:P[p]['pi_q'] for p in PROV},
  'Peligro Integrado modelado (MCA–AHP) por nivel', etiqueta_min=5)
SEQ_E6=['#e8dcc6','#c9a46b','#9a6b33','#5c3a17']
ord_e6=['No degradado','Leve','Moderado','Severo']
apilado('fig_e6',ord_e6,SEQ_E6,{p:[P[p]['e6_clsha'].get(k,0) for k in ord_e6] for p in PROV},
  'Clase de estado del ecosistema (índice de condición, E6)', etiqueta_min=4)

# Infiltración básica (Informe de Avance E5, Cuadro 19)
inf=[('INF-05 Santa Rosa de Chirimoyos',0.33),('INF-09 El Guabo',0.35),('INF-03 Sahuate Hualanga',0.41),('INF-01 Ricardo Palma',0.59),
 ('INF-04 Chililique Alto',0.78),('INF-02 Hualas',0.79),('INF-12 Pampa Quemada',0.83),('INF-07 La Peña',0.93),('INF-06 Quemazón',1.37),
 ('INF-10 La Laguna',2.20),('INF-02b El Ala',6.88),('INF-08 Cruz de Piedra',13.25),('INF-11 La Unión / Culcas Alto',73.48)]
fig,ax=plt.subplots(figsize=(7.2,3.6))
ys=range(len(inf))
ax.barh(list(ys),[v for _,v in inf],color='#2a78d6',height=0.6)
for i,(n,v) in enumerate(inf): ax.text(v*1.08,i,f'{v:.2f}',va='center',fontsize=8,color=INK)
ax.set_yticks(list(ys)); ax.set_yticklabels([n for n,_ in inf],fontsize=8); ax.invert_yaxis()
ax.set_xscale('log'); ax.set_xlim(0.2,200); ax.set_xticks([0.2,0.5,1,2,5,10,20,50,100,200]); ax.set_xticklabels(['0.2','0.5','1','2','5','10','20','50','100','200']); ax.set_xlabel('Velocidad de infiltración básica representativa (cm/h, escala logarítmica)')
ax.set_title('Infiltración básica por sitio — doble anillo (Muntz), ajuste de Kostiakov',loc='left',fontsize=9.5,fontweight='bold',color=INK)
fin(fig,ax,'fig_infiltracion')

# Coberturas S2DR4 en los 117 bloques (E6, Tabla 7)
cob=[('Herbazal seco / pastizal senescente',3696.83),('Dosel caducifolio defoliado / matorral seco',2384.07),('Bosque ralo / matorral abierto',2251.87),
 ('Suelo desnudo o muy ralo, roca, vías',1570.83),('Dosel siempreverde denso',1105.19),('Sombra topográfica',633.62),
 ('Vegetación verde continua de porte bajo',413.86),('Nubes / neblina residual',213.96)]
fig,ax=plt.subplots(figsize=(7.2,3.0))
ax.barh(range(len(cob)),[v for _,v in cob],color='#1baf7a',height=0.6)
for i,(n,v) in enumerate(cob): ax.text(v+40,i,f'{v:,.0f} ha · {100*v/12270.24:.1f} %',va='center',fontsize=8,color=INK)
ax.set_yticks(range(len(cob))); ax.set_yticklabels([n for n,_ in cob],fontsize=8); ax.invert_yaxis(); ax.set_xlim(0,4800)
ax.set_xlabel('Superficie (ha) en los 117 bloques (12,270.24 ha)')
ax.set_title('Cobertura al final de la época seca 2026 — 92 tomas S2DR4 de 1 m',loc='left',fontsize=9.5,fontweight='bold',color=INK)
fin(fig,ax,'fig_coberturas')
print('ok')
