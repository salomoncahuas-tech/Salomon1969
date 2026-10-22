import json, re
S='/tmp/claude-0/-home-user-Salomon1969/232a4de4-cdc3-5473-be12-eb7b421c81bc/scratchpad/'
B_=S+'build/'
SUE=json.load(open(B_+'sue.json'))
prep=open(B_+'sue_prep.js').read(); tab=open(B_+'sue_tab.js').read()
css=open(B_+'sue.css').read(); sec=open(B_+'section.html').read()

def rep(s,a,b,count=1):
    n=s.count(a)
    if n<count or n==0: raise SystemExit(f'NOT FOUND ({n}): {a[:90]}')
    return s.replace(a,b) if count=='all' else s.replace(a,b,count)

def common(c):
    c=rep(c,"const SVGNS='http://www.w3.org/2000/svg';","const SVGNS='http://www.w3.org/2000/svg';\n"+prep)
    c=rep(c,"  {k:'CV',l:'CV lluvia',g:'E4 Hidrología',c:'--st-hid'},",
          "  {k:'CV',l:'CV lluvia',g:'E4 Hidrología',c:'--st-hid'},\n  {k:'prof',l:'Prof. efectiva',g:'E5 Suelos',c:'--st-sue'},\n  {k:'ero',l:'Erosión campo',g:'E5 Suelos',c:'--st-sue'},")
    c=rep(c,"v:b=>b.rk||'—'}\n];",
          "v:b=>b.rk||'—'},\n  {k:'suelo',sh:'Suelo',l:'Suelo limitante',s:'E5 Suelos',c:'--st-sue',crit:'Calicata de bloque con profundidad efectiva < 50 cm o erosión hídrica severa (ficha edáfica E5)',f:b=>b.sLim===true,v:b=>b.sCal?(b.prof!=null?b.prof+' cm':'s/d')+(b.ero===4?' · sev.':''):'s/c'}\n];")
    c=rep(c,"['Estudios integrados',n0(I.nstudies),`de ${n0(I.nstudies+1)} con datos por bloque · falta E5 Suelos`,''],",
          "['Estudios integrados',n0(I.nstudies),`con datos por bloque · E5 Suelos con ${CV.length} calicatas de campo`,''],")
    c=rep(c,"['Convergencia ≥ 4 de 6'","['Convergencia ≥ 4 de 7'")
    c=rep(c,"{id:'E5',c:1,y:248,code:'E5',t:'Suelos',s:'Retención hídrica · erodabilidad',v:'No incorporado',col:'--st-sue',tab:null,gap:true},",
          "{id:'E5',c:1,y:248,code:'E5',t:'Suelos',s:'Calicatas · infiltración · CUM',v:`${CV.length} calicatas · ${IV.length} prueba${IV.length===1?'':'s'} inf.`,col:'--st-sue',tab:'sue'},")
    c=rep(c,"['E1','E5',1]","['E1','E5']")
    c=rep(c,"['E5','E6',1],['E5','E8',1]","['E5','E6',0,['prof','ms']],['E5','E8']")
    c=rep(c,"""     `<span class="gap">No incorporado a la base de este volumen.</span> La erosión observada en las fichas F-DT y la capa de cárcavas son los únicos indicadores edáficos disponibles.`,
     ['Diagnóstico del territorio','Diagnóstico de la UP'],null],""",
          """     `${CV.length} calicatas (${CB.length} de bloque) en ${B.filter(b=>b.sAll.length).length} de ${B.length} bloques; profundidad efectiva mediana ${CB.length?n0(median(CB.map(c=>c.pcm)))+' cm':'—'}; ${SLIM.length} bloques con suelo limitante (< 50 cm o erosión severa). ${IV.length} pruebas de infiltración${IV.length?' ('+IV.map(s=>fmt(s.ib,2)).join(', ')+' cm/h)':''}. Con la cobertura: ${rt('prof','ms')}; con el peligro: ${rt('prof','ipp')}. <span class="gap">Laboratorio pendiente.</span>`,
     ['Diagnóstico del territorio','Diagnóstico de la UP','GdR-CCC'],'sue'],""")
    c=rep(c,"reúnen 4 o más de las 6 líneas de evidencia","reúnen 4 o más de las 7 líneas de evidencia")
    c=rep(c,"""    ['--st-hid','Clima','E4',""","""    ['--st-sue','Suelo','E5',CB.length?n0(median(CB.map(c=>c.pcm))):'—','cm',`Profundidad efectiva mediana de ${CB.length} calicatas de bloque; ${CB.filter(c=>c.pcm!=null&&c.pcm<50).length} con contacto antes de 50 cm.`],
    ['--st-hid','Clima','E4',""")
    c=rep(c,"const L=[['eh','ms','El sustrato frente a la cobertura'],","const L=[['eh','ms','El sustrato frente a la cobertura'],['prof','ms','El suelo frente a la cobertura'],")
    c=rep(c,"const by=[0,1,2,3,4,5,6].map","const by=[0,1,2,3,4,5,6,7].map")
    c=rep(c,"const bw=(W-L-Rr)/7;","const bw=(W-L-Rr)/8;")
    c=rep(c,"${d.k} de 6</text>","${d.k} de 7</text>")
    c=rep(c,"<b>${d.k} de 6 líneas</b>","<b>${d.k} de 7 líneas</b>")
    c=rep(c,"[0,1,2,3,4,5].map(i=>`<span class=\"${i<b.sc?'f':''}\"></span>`).join('')}<em>${b.sc}/6</em>","[0,1,2,3,4,5,6].map(i=>`<span class=\"${i<b.sc?'f':''}\"></span>`).join('')}<em>${b.sc}/7</em>")
    # hallazgo de suelos en la lectura integrada
    c=rep(c,"""  $('intFinds').innerHTML=F.map(""","""  { const ab=CB.map(x=>({x,a:apt(x)})), cz=ab.filter(o=>o.a.z==='Compatible').length, rz=ab.filter(o=>o.a.z.startsWith('Restringida')).length, nz=ab.filter(o=>o.a.z.startsWith('No ·')).length, sp=SLIM.filter(b=>b.ev[1]);
    if(CB.length) F.push(['adv','El suelo acota la infraestructura marrón',`El E5 describió calicata de bloque en ${CB.length} de ${B.length} bloques. ${SLIM.length} tienen suelo limitante (contacto antes de 50 cm o erosión severa): ${lst([...SLIM])}${sp.length?`; ${sp.length} de ellos con peligro crítico, donde la cobertura y la retención en superficie reemplazan a la excavación`:''}. Según la matriz suelo–intervención preliminar, las zanjas de infiltración son compatibles en ${cz} calicatas de bloque, restringidas en ${rz} y descartadas por pendiente mayor de 50 % en ${nz}. La velocidad básica medida (${IV.length?IV.map(s=>fmt(s.ib,2)).join(', ')+' cm/h':'sin prueba en este volumen'}) fija el espaciamiento y la capacidad de las zanjas; donde es lenta, se requieren zanjas más próximas y de mayor almacenamiento.`]); }
  $('intFinds').innerHTML=F.map(""")
    c=rep(c,"render();\n// abrir directamente #int\ntry{ if((location.hash||'').slice(1)==='int') goTab('int'); }catch(e){}",
          tab+"\nrender();\n[sueTab,fichaSue,panoSue].forEach(f=>{try{f();}catch(e){console.error('E5',e);}});\n// abrir directamente #int o #sue\ntry{ const hh=(location.hash||'').slice(1); if(hh==='int'||hh==='sue') goTab(hh); }catch(e){}")
    return c

VOL=[('aya','Ayabaca',2028,(2029,2450)),('hua','Huancabamba',1602,(1603,2024)),('mor','Morropón',1484,(1485,1906))]
for f,prov,lnI,(a,b) in VOL:
    L=open(S+'art/'+f+'_orig.html',encoding='utf8').read().split('\n')
    s=L[lnI-1]; I=json.loads(s[s.index('window.INT=')+11:s.rindex('</script>')].rstrip().rstrip(';'))
    bset={x['b'] for x in I['B']}
    cv=[c for c in SUE['C'] if c['b'] in bset]; nb=len({c['b'] for c in cv})
    na=sum(1 for x in I['B'] if any(u in SUE['U'] for u in SUE['BL'][x['b']]['us']))
    I['nstudies']=9
    I['gaps']=[g for g in I['gaps'] if g[0]!='E5 Suelos']
    I['gaps'].insert(3,['E5 Suelos',f'Incorporado con datos de campo (Entregable 2): {len(cv)} calicatas en {nb} de {len(I["B"])} bloques; {na} bloques con alguna unidad de suelo muestreada. Faltan textura, materia orgánica, pH, saturación de bases, Factor K y CUM definitiva (laboratorio UNALM).',
        'La línea de suelo de la convergencia solo se activa donde hay calicata de bloque; el dimensionamiento de zanjas y la CUM por bloque siguen siendo preliminares.',
        'Incorporar el Entregable 3 (laboratorio, 11/10/2026) y el Entregable 4 (CUM, Factor K y matriz suelo–intervención por unidad) y extender la caracterización por analogía a los bloques sin calicata.','Media'])
    I['kpi'].append(['Suelos E5',str(len(cv)),f'calicatas en {nb} bloques · campo','--st-sue'])
    c='\n'.join(L[a-1:b])
    c=common(c)
    head='\n'.join(L[:lnI-1]); tail='\n'.join(L[b:])
    head=rep(head,'<header class="hero">',css+'\n<header class="hero">')
    head=rep(head,'<li><i style="background:var(--st-hid)"></i>E4 Hidrología</li>','<li><i style="background:var(--st-hid)"></i>E4 Hidrología</li><li><i style="background:var(--st-sue)"></i>E5 Suelos</li>')
    head=rep(head,'DS 26/09/2026</p>','DS 26/09/2026 · suelos (campo) 20/09/2026</p>')
    head=rep(head,'<button role="tab" data-g="Estudios básicos" data-p="hid" aria-selected="false">Hidrología</button>','<button role="tab" data-g="Estudios básicos" data-p="hid" aria-selected="false">Hidrología</button>\n  <button role="tab" data-g="Estudios básicos" data-p="sue" aria-selected="false">Suelos</button>')
    head=rep(head,'<section id="p-brec" hidden>',sec+'\n<section id="p-brec" hidden>')
    head=rep(head,'Seis líneas · clic en la fila abre la ficha','Siete líneas · clic en la fila abre la ficha')
    head=rep(head,'Plena: 5–6 líneas; alta: 3–4; parcial: 0–2.','Plena: 5–7 líneas; alta: 3–4; parcial: 0–2. La línea de suelo solo puede activarse en los bloques con calicata de bloque («s/c»: sin calicata; no se imputa).')
    head=rep(head,'nivel de riesgo social)','nivel de riesgo social, profundidad efectiva de la calicata)')
    head=rep(head,'brecha, peligro, sustrato, clima, erosión y riesgo social','brecha, peligro, sustrato, suelo, clima, erosión y riesgo social')
    head=rep(head,'Línea discontinua: vínculo con el Estudio de Suelos (E5), no incorporado a este volumen, y relación recíproca Geología–Hidrología–Suelos.','Línea discontinua: relación recíproca Geología–Hidrología–Suelos.')
    head=rep(head,'<div><b>Fuentes sociales:</b>','<div><b>Fuentes edáficas:</b> GeoSIG Ingenieros E.I.R.L., Estudio de Suelos a nivel de reconocimiento, Informe de avance de la Fase de Campo (Entregable 2, datos al 20/09/2026): 74 fichas de calicata, 148 de chequeo y 13 libros de cálculo de infiltración (Cuadros 4 a 23) · geodatabase ENTREGABLE2_PUNTOS_MUESTREADOS (capas CALICATAS_EJECUTADAS y BLOQUES_INTERVENCION, EPSG:32717) · TDR y Plan de Trabajo del servicio · color Munsell convertido a sRGB con la renotación Munsell (iluminante C)</div>\n  <div><b>Fuentes sociales:</b>')
    intl='<script>window.INT='+json.dumps(I,ensure_ascii=False,separators=(',',':'))+';</script>\n<script>window.SUE='+json.dumps(SUE,ensure_ascii=False,separators=(',',':'))+';</script>'
    out=head+'\n'+intl+'\n'+c+'\n'+tail
    open(S+'build/'+f+'.html','w',encoding='utf8').write(out)
    print(f,len(out),len(cv),nb,na)
