# Calcula, desde las fuentes, todas las cifras y cuadros del Segundo Entregable (actividad b).
import json, collections, pandas as pd
S='/tmp/claude-0/-home-user-Salomon1969/03b82ed8-c0e9-54ec-93c9-bff5f93b575b/scratchpad'
J=S+'/territorial/json/'
L=lambda p,k: json.load(open(J+f'{p}_{k}.json'))
TH=0.4976
pred=pd.read_csv(S+'/territorial/predial_v6.csv')
out={'prov':{}, 'dist':[], 'bloques':[]}
PN={'HUA':'Huancabamba','MOR':'Morropón'}
for p in ['HUA','MOR']:
    D=L(p,'D'); E=L(p,'ECO6'); X=L(p,'SPX'); SU=L(p,'SUE'); I=L(p,'INT'); U=L(p,'UPD'); SO=L(p,'SOC')
    B=D['B']; Eb={e['bloque']:e for e in E['b']}; Xb={x['b']:x for x in X['x']}
    prov=PN[p]
    cal=[c for c in SU['C'] if SU['BL'].get(c['b'],{}).get('p')==prov]
    calB={c['b'] for c in cal if c['t']=='B'}; calAny={c['b'] for c in cal}
    q=[0]*5
    for b in B:
        for i,v in enumerate(b['pi']['q']): q[i]+=v
    cls=collections.Counter(e['clase_estado'] for e in E['b'])
    clsha={k:round(sum(e['area_ha'] for e in E['b'] if e['clase_estado']==k),2) for k in cls}
    pr=pred[pred.prov==prov]
    sitc=pr.groupby('sit').agg(n=('sup','size'),ha=('sup','sum')).round(2)
    kp={r[0]:r[1:] for r in I['kpi']}
    cab={r[0]:r[1:] for r in U['res']['cab']}
    out['prov'][prov]={
      'n':len(B),'ha':round(sum(b['ha'] for b in B),6),'nd':I['nd'],'nmc':I['nmc'],
      'pi_q':[round(v,6) for v in q],'pi_cls':dict(collections.Counter(b['pi']['cl'] for b in B)),
      'pi_idx':I['agg']['ipp'],'p45':I['agg']['p45'],'h45':sum(b['pi']['h45'] for b in B),
      'conc':dict(collections.Counter(b.get('conc') for b in B)),
      'brecha_media_n':sum(1 for b in B if b['ms']<TH),'brecha_media_ha':round(sum(b['ha'] for b in B if b['ms']<TH),6),
      'e6_bajo_ha':round(sum(e['msavi_bajo_umbral_ha'] for e in E['b']),6),
      'e6_rest_msavi_ha':round(sum(e['msavi_rest_activa_ha'] or 0 for e in E['b']),6),
      'e6_cls':dict(cls),'e6_clsha':clsha,
      'e6_ic_pond':round(sum(e['ic_bloque']*e['area_ha'] for e in E['b'])/sum(e['area_ha'] for e in E['b']),3),
      'e6_delim':{k:round(sum(e[k] or 0 for e in E['b']),6) for k in ['delim_rest_activa_ha','delim_enriquecimiento_ha','delim_conservacion_ha','delim_tamizaje_agricola_ha']},
      'e6_evid':dict(collections.Counter(e['clase_evidencia'] for e in E['b'])),
      'e6_parcelas':sum(e['n_parcelas'] or 0 for e in E['b']),
      'e6_eco':{k:[sum(1 for e in E['b'] if e['ecosistema']==k), round(sum(e['area_ha'] for e in E['b'] if e['ecosistema']==k),6)] for k in set(e['ecosistema'] for e in E['b'])},
      'e6_tipo':dict(collections.Counter(e['tipologia_S2DR4'] for e in E['b'])),
      'pred_cc':round(pr.cc.sum(),6),'pred_cof':round(pr.priv.sum(),6),'pred_est':round(pr.est.sum(),6),'pred_npred':int(pr.npred.sum()),
      'pred_sit':{k:[int(v.n),float(v.ha)] for k,v in sitc.iterrows()},
      'pred_gt19':[[r.bloque,int(r.npred)] for r in pr.itertuples() if r.npred>19],
      'pred_nccpp':int(pr.nccpp.sum()),
      'cal_n':len(cal),'cal_t':dict(collections.Counter(c['t'] for c in cal)),'cal_bloques':len(calAny),'calB_bloques':len(calB),
      'cv_n':sum(len(b.get('cv') or []) for b in B),'cv_m':round(sum(c['lon'] for b in B for c in (b.get('cv') or [])),1),'cv_bloques':sum(1 for b in B if b.get('cv')),
      'riesgo':dict(collections.Counter(b['rk']['nivel'] for b in B)),
      'expo':dict(collections.Counter(b.get('exp') for b in B)),
      'kpi':kp,'ds_cab':cab,'nov':[[n[0],n[1]] for n in U['nov']],'gaps':I['gaps'],
      'lotes':SO['lotes'],'regTxt':I['regTxt'],
    }
    # distritos
    dist=collections.defaultdict(list)
    for b in B: dist[b['d']].append(b)
    for d,bs in sorted(dist.items(), key=lambda kv:-sum(b['ha'] for b in kv[1])):
        es=[Eb[b['b']] for b in bs]; xs=[Xb[b['b']] for b in bs]
        ha=sum(b['ha'] for b in bs)
        out['dist'].append({'prov':prov,'d':d,'n':len(bs),'ha':round(ha,6),
          'bajo_n':sum(1 for b in bs if b['ms']<TH),
          'e6_bajo':round(sum(e['msavi_bajo_umbral_ha'] for e in es),6),
          'pi45':round(sum(b['pi']['h45'] for b in bs),6),
          'ic':round(sum(e['ic_bloque']*e['area_ha'] for e in es)/sum(e['area_ha'] for e in es),6),
          'modsev':sum(1 for e in es if e['clase_estado'] in ('Moderado','Severo')),
          'calB':sum(1 for b in bs if b['b'] in calB),
          'cc':round(sum(x['cc'] or 0 for x in xs),6),'cof':round(sum(x['cof'] or 0 for x in xs),6),'est':round(sum(x['est'] or 0 for x in xs),6),
          'npred':sum(x['npc'] or 0 for x in xs)})
    for b in B:
        e=Eb[b['b']]; x=Xb[b['b']]
        out['bloques'].append({'prov':prov,'b':b['b'],'d':b['d'],'mc':b['mc'],'ha':b['ha'],'e':b['e'],'n':b['n'],'reg':b['reg'],
          'ms':b['ms'],'pi':b['pi']['cl'],'ipp':b['pi']['ipp'],'ncv':len(b.get('cv') or []),'ic':e['ic_bloque'],'est':e['clase_estado'],'tip':e['tipologia_S2DR4'],
          'bu':e['msavi_bajo_umbral_pct'],'sit':x['sit'],'npred':x['npc'] or 0,'cal':'Sí' if b['b'] in calB else ('Ref.' if b['b'] in calAny else 'No'),
          'ds':'Sí' if x['nf'] else 'No','rk':b['rk']['nivel'],'sp':x['sp']})
json.dump(out,open(S+'/build/datos.json','w'),ensure_ascii=False,indent=1)
for k,v in out['prov'].items():
    print(k, {kk:v[kk] for kk in ['n','ha','nd','nmc','pi_q','pi_cls','brecha_media_n','brecha_media_ha','e6_bajo_ha','e6_cls','e6_clsha','e6_ic_pond','e6_delim','pred_sit','pred_gt19','cal_n','cal_t','cal_bloques','calB_bloques','cv_n','riesgo','e6_parcelas','e6_tipo']})
print(len(out['dist']), len(out['bloques'])); print(out['bloques'][0])
