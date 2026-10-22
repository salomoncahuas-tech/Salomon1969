/* =====================================================================
   DIAGNÓSTICO INTEGRAL IN PIURA V1 — portada integral y simulador
   ===================================================================== */
const $=id=>document.getElementById(id);
const esc=s=>String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const NF={};
const fmt=(v,d=0)=>{if(v==null||!isFinite(v)) return '—'; const k=d; NF[k]=NF[k]||new Intl.NumberFormat('en-US',{minimumFractionDigits:d,maximumFractionDigits:d}); return NF[k].format(v);};
const n0=v=>fmt(v,0), n2=v=>fmt(v,2), n3=v=>fmt(v,3);
const sol=v=>'S/ '+fmt(v,0), solM=v=>'S/ '+fmt(v/1e6,2)+' M', pct=(v,d=1)=>fmt(v*100,d)+' %';
const sum=(a,f)=>a.reduce((s,x)=>s+(+f(x)||0),0);
const wmean=(a,f,w)=>{const W=sum(a,w); return W?sum(a,x=>f(x)*w(x))/W:null;};
const pnum=s=>{const m=String(s).replace(/,/g,'').match(/-?\d+(\.\d+)?/); return m?+m[0]:null;};

const VK=['mor','hua','aya'];
const VM={mor:{name:'Morropón',vol:'I',href:'vol1-morropon/index.html',col:'var(--pv-mor)',hcol:'var(--pvh-mor)',X:'MORROPON'},
          hua:{name:'Huancabamba',vol:'II',href:'vol2-huancabamba/index.html',col:'var(--pv-hua)',hcol:'var(--pvh-hua)',X:'HUANCABAMBA'},
          aya:{name:'Ayabaca',vol:'III',href:'vol3-ayabaca/index.html',col:'var(--pv-aya)',hcol:'var(--pvh-aya)',X:'AYABACA'}};
const XV={MORROPON:'mor',HUANCABAMBA:'hua',AYABACA:'aya'};
const V=DATA.vols, C=DATA.cost;
const ALL=VK.flatMap(k=>V[k].B.map(b=>({...b,v:k})));
const PIN=['Muy bajo','Bajo','Medio','Alto','Muy alto'];
const PIC=['var(--pi-1)','var(--pi-2)','var(--pi-3)','var(--pi-4)','var(--pi-5)'];
const piCls=v=>v==null?null:v<1.8?0:v<2.6?1:v<3.4?2:v<4.2?3:4;
const piPill=v=>{const k=piCls(v); return k==null?'—':`<span class="pi-pill pi-${k+1}" style="background:${PIC[k]}">${n3(v)} · ${PIN[k]}</span>`;};

/* ---------- tooltip ---------- */
const tip=$('tip');
function showTip(e,h){tip.innerHTML=h; tip.hidden=false; const w=tip.offsetWidth,hh=tip.offsetHeight;
  let x=e.clientX+14,y=e.clientY+14; if(x+w>innerWidth-8) x=e.clientX-w-14; if(y+hh>innerHeight-8) y=e.clientY-hh-14; tip.style.left=Math.max(8,x)+'px'; tip.style.top=Math.max(8,y)+'px';}
function hideTip(){tip.hidden=true;}
const tr=(k,v)=>`<div class="r"><span>${k}</span><span>${v}</span></div>`;

/* ---------- agregados por volumen ---------- */
const AG={};
VK.forEach(k=>{const B=V[k].B, ha=sum(B,b=>b.ha), kp=V[k].kpi;
  const get=lbl=>kp.find(x=>x[0]===lbl)||['','','',''];
  const ds=V[k].gaps.find(g=>g[0]==='E2 Diagnóstico social'); const m=ds&&ds[1].match(/fichas en (\d+) de (\d+) bloques/);
  const act=get('Diagnóstico social')[2].match(/(\d+) de (\d+) actores/);
  const cal=get('Suelos E5'); const calB=cal[2].match(/en (\d+) bloques/);
  const e6=get('Ecosistemas E6')[2]; const e6m=e6.match(/(\d+) bloques moderado o severo · ([\d,\.]+) ha/);
  AG[k]={n:B.length,ha,piA:pnum(get('PI Alto + Muy alto')[1]),ipp:V[k].agg.ipp,P:V[k].agg.P,AI:V[k].agg.AI,pob:pnum(get('Población INEI')[1]),
    fds:pnum(get('Diagnóstico social')[1]),fdsB:m?+m[1]:null,actF:act?+act[1]:null,actT:act?+act[2]:null,
    cal:pnum(cal[1]),calB:calB?+calB[1]:null,e6b:e6m?+e6m[1]:null,e6ha:e6m?pnum(e6m[2]):null,
    brecha:pnum(get('Brecha MSAVI')[1]),nd:V[k].nd,nmc:V[k].nmc,ncv:sum(B,b=>b.ncv),lcv:sum(B,b=>b.lcv)};
});
const TOT={n:sum(VK,k=>AG[k].n),ha:sum(VK,k=>AG[k].ha),piA:sum(VK,k=>AG[k].piA),pob:sum(VK,k=>AG[k].pob),fds:sum(VK,k=>AG[k].fds),
  fdsB:sum(VK,k=>AG[k].fdsB),actF:sum(VK,k=>AG[k].actF),actT:sum(VK,k=>AG[k].actT),cal:sum(VK,k=>AG[k].cal),calB:sum(VK,k=>AG[k].calB),
  e6b:sum(VK,k=>AG[k].e6b),e6ha:sum(VK,k=>AG[k].e6ha),nd:sum(VK,k=>AG[k].nd),nmc:sum(VK,k=>AG[k].nmc),ncv:sum(VK,k=>AG[k].ncv),lcv:sum(VK,k=>AG[k].lcv)};
TOT.ipp=sum(VK,k=>AG[k].ipp*AG[k].ha)/TOT.ha; TOT.P=sum(VK,k=>AG[k].P*AG[k].ha)/TOT.ha; TOT.AI=sum(VK,k=>AG[k].AI*AG[k].ha)/TOT.ha;

/* ---------- motor de costos ---------- */
const ENG=cpEngine(C);
const CB=C.blocks;
const IDS_ALL=new Set(CB.map(b=>b.id));
const BASE=ENG.run(IDS_ALL,ENG.defaults());
const intHa={}; ALL.forEach(b=>intHa[b.b]=b.ha);

/* =====================================================================
   PORTADA
   ===================================================================== */
const volLinks=cur=>[['Integral','index.html',cur==='int'],...VK.map(k=>[`Vol. ${VM[k].vol} · ${VM[k].name}`,VM[k].href,false])];
$('sub').textContent=`Los ${TOT.n} bloques preliminares de intervención —${n2(TOT.ha)} ha en ${TOT.nd} distritos de Morropón, Huancabamba y Ayabaca— con su caracterización biofísica, peligro, brecha y diagnóstico social, leídos con un solo formato, y su traducción a metas físicas y costos con el motor paramétrico del Estudio de Perfil.`;

(function heroMap(){
  const W=460,H=330,pad=16, es=ALL.map(b=>b.e), ns=ALL.map(b=>b.n);
  const x0=Math.min(...es),x1=Math.max(...es),y0=Math.min(...ns),y1=Math.max(...ns);
  const s=Math.min((W-2*pad)/(x1-x0),(H-2*pad)/(y1-y0)), ox=(W-(x1-x0)*s)/2, oy=(H-(y1-y0)*s)/2;
  const X=e=>ox+(e-x0)*s, Y=n=>H-oy-(n-y0)*s, R=ha=>Math.max(2.2,Math.sqrt(ha)*0.55);
  const dots=[...ALL].sort((a,b)=>b.ha-a.ha).map(b=>`<circle cx="${X(b.e).toFixed(1)}" cy="${Y(b.n).toFixed(1)}" r="${R(b.ha).toFixed(1)}" data-b="${esc(b.b)}" style="fill:${VM[b.v].hcol};fill-opacity:.78;stroke:rgba(6,24,40,.55);stroke-width:.8"></circle>`).join('');
  const km=20000*s;
  $('heroMap').innerHTML=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Centroides de los 117 bloques en UTM 17S, coloreados por provincia">${dots}
    <g transform="translate(${W-pad-km},${H-10})"><rect width="${km}" height="3" style="fill:#F4F9FC"></rect><text x="${km/2}" y="-5" text-anchor="middle" style="fill:#D2E4F0;font:500 10px 'IBM Plex Mono',monospace">20 km</text></g>
    <g transform="translate(${pad+4},${pad+8})"><path d="M0 14 L6 0 L12 14 L6 10 Z" style="fill:#F4F9FC"></path><text x="6" y="26" text-anchor="middle" style="fill:#D2E4F0;font:600 10px 'IBM Plex Mono',monospace">N</text></g></svg>`;
  $('heroMap').querySelectorAll('circle').forEach(c=>{const b=ALL.find(x=>x.b===c.dataset.b);
    c.addEventListener('mousemove',e=>showTip(e,`<b>Bloque ${esc(b.b)}</b>${tr('Provincia',VM[b.v].name)}${tr('Distrito',esc(b.d))}${tr('Superficie',n2(b.ha)+' ha')}${tr('Índice PI',n3(b.ipp))}${tr('UTM 17S',n0(b.e)+' E · '+n0(b.n)+' N')}`));
    c.addEventListener('mouseleave',hideTip); c.style.cursor='pointer'; c.addEventListener('click',()=>{location.href=VM[b.v].href;});});
  $('heroMapCap').innerHTML=`<span>${ALL.length} centroides · UTM WGS 84 Zona 17S · área ∝ superficie</span>`+VK.map(k=>`<span><i style="background:${VM[k].hcol}"></i>${VM[k].name}</span>`).join('');
})();

(function strip(){
  const kp=[
    ['Bloques',n0(TOT.n),`3 provincias · ${TOT.nd} distritos · ${TOT.nmc} microcuencas`,'--anin'],
    ['Superficie',n2(TOT.ha),'ha · UTM WGS 84 Zona 17S','--anin'],
    ['PI Alto + Muy alto',n2(TOT.piA),`ha · ${pct(TOT.piA/TOT.ha)} del ámbito`,'--st-pi'],
    ['Índice PI',n3(TOT.ipp),PIN[piCls(TOT.ipp)]+' · media ponderada','--st-pi'],
    ['Clima',n0(TOT.P),`mm/año · IA ${n2(TOT.AI)} semiárido`,'--st-hid'],
    ['Bajo umbral MSAVI',n2(TOT.e6ha),`ha · E6, criterio areal · ${pct(TOT.e6ha/TOT.ha)}`,'--st-veg'],
    ['Población INEI',n0(TOT.pob),'hab. · suma de los volúmenes','--st-soc'],
    ['Diagnóstico social',n0(TOT.fds),`fichas F-DS · ${TOT.fdsB} de ${TOT.n} bloques`,'--st-soc'],
    ['Suelos E5',n0(TOT.cal),`calicatas en ${TOT.calB} bloques · campo`,'--st-sue'],
    ['Inversión',solM(BASE.t.AA),`motor v1.3 · ${sol(BASE.t.AA/BASE.t.ha)} por ha`,'--gold'],
    ['O&M 15 años',solM(BASE.ind.om15),`sin descontar · VAC social ${solM(BASE.ind.vacS)}`,'--gold']];
  $('strip').innerHTML=kp.map(([k,v,sm,c])=>`<div><dt><i style="background:var(${c})"></i>${k}</dt><dd>${v} <small>${sm}</small></dd></div>`).join('');
})();

function setVolsNav(){
  const html=volLinks('int').map(([t,h,cur])=>`<a href="${h}"${cur?' aria-current="page"':''}>${t}</a>`).join('')+`<a href="#costos" class="cst">Metrados y costos</a>`;
  document.querySelector('nav.vols').innerHTML=html;
}
setVolsNav();

/* ---------- navegación en dos niveles ---------- */
const tabs=[...document.querySelectorAll('nav.tabs button[role="tab"]')];
const groups=[...new Set(tabs.map(b=>b.dataset.g))];
const grow=$('grow');
function setGroup(g){grow.querySelectorAll('.gbtn').forEach(x=>x.setAttribute('aria-pressed',String(x.dataset.g===g)));
  tabs.forEach(b=>{if(b.dataset.g===g) b.removeAttribute('data-off'); else b.setAttribute('data-off','');});}
grow.innerHTML=groups.map(g=>`<button class="gbtn" type="button" data-g="${esc(g)}" aria-pressed="false">${esc(g)} <b>${tabs.filter(t=>t.dataset.g===g).length}</b></button>`).join('');
grow.addEventListener('click',e=>{const x=e.target.closest('.gbtn'); if(!x) return; const f=tabs.find(b=>b.dataset.g===x.dataset.g); if(f) show(f.dataset.p,true);});
function show(p,user){const b=tabs.find(x=>x.dataset.p===p)||tabs[0];
  tabs.forEach(t=>t.setAttribute('aria-selected',String(t===b)));
  document.querySelectorAll('main section').forEach(s=>s.hidden=s.id!=='p-'+b.dataset.p);
  setGroup(b.dataset.g);
  try{history.replaceState(null,'','#'+b.dataset.p);}catch(e){}
  try{localStorage.setItem('inp-int-tab',b.dataset.p);}catch(e){}
  if(user){const nav=document.querySelector('nav.tabs'); const top=nav.getBoundingClientRect().top+scrollY; if(scrollY>top) scrollTo({top,behavior:'instant'});}
  if(b.dataset.p==='costos'||b.dataset.p==='sup') renderCost();
}
tabs.forEach(b=>b.addEventListener('click',()=>show(b.dataset.p,true)));
addEventListener('hashchange',()=>{const h=(location.hash||'').slice(1); if(tabs.some(t=>t.dataset.p===h)) show(h,true);});

/* =====================================================================
   PANORAMA INTEGRAL
   ===================================================================== */
(function finds(){
  const share=VK.map(k=>`${VM[k].name} ${pct(AG[k].ha/TOT.ha)}`).join(', ');
  const piTxt=VK.map(k=>`${VM[k].name} ${pct(AG[k].piA/AG[k].ha)}`).join(' · ');
  const tot15=BASE.t.AA+BASE.ind.om15, techo=C.par.P_TECHO;
  const F=[
    ['neu','Un ámbito de 117 bloques, desigual entre provincias',`<p>Los tres volúmenes suman <strong>${TOT.n} bloques y ${n2(TOT.ha)} ha</strong> en ${TOT.nd} distritos y ${TOT.nmc} microcuencas. La superficie se reparte así: ${share}. Morropón reúne más de la mitad del ámbito y, con ello, la mayor parte de las metas físicas y del costo.</p><p class="src">El catálogo V6 registra 12,270.235 ha; la diferencia con la suma de los volúmenes es de redondeo por bloque.</p>`],
    ['cri','El peligro alto y muy alto domina en las tres provincias',`<p>El Peligro Integrado alto y muy alto cubre <strong>${n2(TOT.piA)} ha (${pct(TOT.piA/TOT.ha)} del ámbito)</strong>: ${piTxt}. El índice PI medio ponderado es ${n3(TOT.ipp)} (${PIN[piCls(TOT.ipp)]}).</p>`],
    ['adv','La brecha MSAVI todavía no es sumable entre volúmenes',`<p>El Volumen III calcula la brecha por distribución areal del MSAVI y los Volúmenes I y II por la media del bloque, de modo que sus cifras (${VK.map(k=>VM[k].name+' '+n2(AG[k].brecha)+' ha').join(', ')}) no se suman. El E6 sí aplica un mismo criterio areal en los 117 bloques: <strong>${n2(TOT.e6ha)} ha bajo el umbral</strong> (${pct(TOT.e6ha/TOT.ha)}).</p><p class="src">Acción propuesta en los tres volúmenes: adoptar la brecha areal del E6 como cifra del indicador (R.M. N.° 00213-2024-MINAM).</p>`],
    ['cri','Con los supuestos v1.3, la cartera supera el presupuesto total de la inversión',`<p>El motor estima una <strong>inversión de ${solM(BASE.t.AA)}</strong> y una operación y mantenimiento de ${solM(BASE.ind.om15)} en 15 años. Juntas suman ${solM(tot15)}, el <strong>${pct(tot15/techo)} del presupuesto total</strong> de S/ ${n0(techo)}, que además incluye preinversión y expediente técnico. Las zanjas y terrazas explican el ${pct(BASE.t.AH/BASE.t.AA)} de la inversión.</p><p class="src">Densidades, espaciamiento de zanjas y terrazas, cerco y años de O&amp;M son supuestos por calibrar con el piloto. Pruebe otros valores en la pestaña <a href="#costos" class="golink">Metrados y costos →</a></p>`],
    ['adv','El diagnóstico social cubre dos tercios de los bloques',`<p>Hay ${n0(TOT.fds)} fichas F-DS vigentes con ficha propia en <strong>${TOT.fdsB} de ${TOT.n} bloques</strong>; ${TOT.actF} de ${TOT.actT} actores consultados están a favor. En ninguno de los tres volúmenes se aplicó aún la ficha de percepción de peligros y cambio climático (F-DS-06).</p>`]];
  $('pFinds').innerHTML=F.map(([c,h,b])=>`<div class="find ${c}"><i></i><div class="body"><h3>${h}</h3>${b}</div></div>`).join('');
})();

(function volCards(){
  $('pVols').innerHTML=VK.map(k=>{const a=AG[k], kp=V[k].kpi; const g=l=>(kp.find(x=>x[0]===l)||[])[1]||'—';
    return `<article class="volcard"><div class="vh"><span>Volumen ${VM[k].vol}</span><h3><i class="pvd" style="background:${VM[k].col}"></i>Provincia de ${VM[k].name}</h3></div>
    <div class="vb"><dl class="kv"><dt>Bloques</dt><dd>${a.n} · ${a.nd} distrito${a.nd>1?'s':''} · ${a.nmc} microcuencas</dd><dt>Superficie</dt><dd>${n2(a.ha)} ha · ${pct(a.ha/TOT.ha)}</dd>
    <dt>PI Alto + Muy alto</dt><dd>${n2(a.piA)} ha · ${pct(a.piA/a.ha)}</dd><dt>Índice PI</dt><dd>${n3(a.ipp)}</dd><dt>Brecha MSAVI</dt><dd>${g('Brecha MSAVI')} ha</dd>
    <dt>Población INEI</dt><dd>${n0(a.pob)} hab.</dd><dt>Fichas F-DS</dt><dd>${a.fds} · ${a.fdsB} de ${a.n} bloques</dd><dt>Situación predial</dt><dd>${g('Situación predial')} en comunidad</dd></dl></div>
    <div class="vf"><a href="${VM[k].href}">Abrir el Volumen ${VM[k].vol} · ${VM[k].name} →</a></div></article>`;}).join('');
})();

(function distChart(){
  const m=new Map(); ALL.forEach(b=>{const key=b.v+'|'+b.d; const o=m.get(key)||{v:b.v,d:b.d,ha:0,n:0}; o.ha+=b.ha; o.n++; m.set(key,o);});
  const R=[...m.values()].sort((a,b)=>b.ha-a.ha), W=820, L=210, rh=26, H=R.length*rh+28, mx=Math.max(...R.map(r=>r.ha)), bw=W-L-150;
  const ticks=[0,500,1000,1500,2000,2500].filter(t=>t<=mx*1.05);
  const g=ticks.map(t=>`<line x1="${L+t/mx*bw}" x2="${L+t/mx*bw}" y1="0" y2="${R.length*rh}" style="stroke:var(--line);stroke-width:1"></line><text x="${L+t/mx*bw}" y="${R.length*rh+16}" text-anchor="middle" style="fill:var(--ink-3);font:500 10.5px 'IBM Plex Mono',monospace">${n0(t)}</text>`).join('');
  const bars=R.map((r,i)=>`<g data-i="${i}"><text x="${L-10}" y="${i*rh+17}" text-anchor="end" style="fill:var(--ink);font:600 12.5px 'Source Sans 3',Arial,sans-serif">${esc(r.d)}</text>
    <rect x="${L}" y="${i*rh+5}" width="${Math.max(1,r.ha/mx*bw).toFixed(1)}" height="${rh-10}" rx="2" style="fill:${VM[r.v].col}"></rect>
    <text x="${L+r.ha/mx*bw+7}" y="${i*rh+17}" style="fill:var(--ink-2);font:500 11.5px 'IBM Plex Mono',monospace">${n2(r.ha)} · ${r.n} bl.</text></g>`).join('');
  $('cDistI').innerHTML=`<div style="overflow-x:auto"><svg viewBox="0 0 ${W} ${H}" style="min-width:620px" role="img" aria-label="Superficie de bloques por distrito">${g}${bars}</svg></div>`;
  $('cDistI').querySelectorAll('g[data-i]').forEach(el=>{const r=R[+el.dataset.i]; el.addEventListener('mousemove',e=>showTip(e,`<b>${esc(r.d)}</b>${tr('Provincia',VM[r.v].name)}${tr('Bloques',r.n)}${tr('Superficie',n2(r.ha)+' ha')}${tr('Del ámbito',pct(r.ha/TOT.ha))}`)); el.addEventListener('mouseleave',hideTip);});
  $('capDistI').innerHTML=VK.map(k=>`<span style="margin-right:14px"><i class="pvd" style="background:${VM[k].col}"></i>${VM[k].name}</span>`).join('')+' Superficie de bloques V6 según los volúmenes; etiqueta: hectáreas y número de bloques.';
})();

(function piChart(){
  const W=820,L=150,rh=44,H=VK.length*rh+10,bw=W-L-20;
  const rows=VK.map((k,i)=>{const B=V[k].B, ha=sum(B,b=>b.ha); const c=[0,0,0,0,0]; B.forEach(b=>{const q=piCls(b.ipp); if(q!=null) c[q]+=b.ha;});
    let x=L; const seg=c.map((v,q)=>{const w=v/ha*bw; const s=w>0?`<rect x="${x.toFixed(1)}" y="${i*rh+8}" width="${w.toFixed(1)}" height="${rh-18}" data-k="${k}" data-q="${q}" style="fill:${PIC[q]}"></rect>${w>46?`<text x="${(x+w/2).toFixed(1)}" y="${i*rh+rh/2+3}" text-anchor="middle" style="fill:${q>=3||q===0?'#fff':'#1d2a10'};font:600 11px 'IBM Plex Mono',monospace;pointer-events:none">${pct(v/ha,0)}</text>`:''}`:''; x+=w; return s;}).join('');
    return `<text x="${L-10}" y="${i*rh+rh/2+3}" text-anchor="end" style="fill:var(--ink);font:600 13px 'Source Sans 3',Arial,sans-serif">${VM[k].name}</text>${seg}`;});
  $('cPiI').innerHTML=`<div style="overflow-x:auto"><svg viewBox="0 0 ${W} ${H}" style="min-width:560px" role="img" aria-label="Superficie por clase de Peligro Integrado y provincia">${rows.join('')}</svg></div>`;
  $('cPiI').querySelectorAll('rect').forEach(r=>{const k=r.dataset.k,q=+r.dataset.q,B=V[k].B.filter(b=>piCls(b.ipp)===q);
    r.addEventListener('mousemove',e=>showTip(e,`<b>${VM[k].name} · ${PIN[q]}</b>${tr('Bloques',B.length)}${tr('Superficie',n2(sum(B,b=>b.ha))+' ha')}`)); r.addEventListener('mouseleave',hideTip);});
  $('legPiI').innerHTML=PIN.map((p,q)=>`<span><i style="background:${PIC[q]}"></i>${p}</span>`).join('');
})();

/* =====================================================================
   COMPARATIVO PROVINCIAL
   ===================================================================== */
(function kpiTable(){
  const labels=V.mor.kpi.map(x=>x[0]);
  const integ={
    'Bloques':[n0(TOT.n),`${TOT.nd} distritos · ${TOT.nmc} microcuencas`],
    'Superficie':[n2(TOT.ha),'ha · suma de los volúmenes'],
    'Brecha MSAVI':['No consolidable','criterio distinto (areal en Vol. III, media del bloque en I y II)'],
    'PI Alto + Muy alto':[n2(TOT.piA),`ha · ${pct(TOT.piA/TOT.ha)} del ámbito`],
    'Índice PI':[n3(TOT.ipp),'media ponderada por superficie'],
    'Clima':[n0(TOT.P),`mm/año · IA ${n2(TOT.AI)} · media ponderada`],
    'Población INEI':[n0(TOT.pob),'hab. · suma de los volúmenes'],
    'Diagnóstico social':[n0(TOT.fds),`fichas F-DS vigentes · ${TOT.actF} de ${TOT.actT} actores a favor`],
    'Suelos E5':[n0(TOT.cal),`calicatas en ${TOT.calB} bloques · campo`],
    'Ecosistemas E6':[n2(TOT.e6ha)+' ha',`bajo umbral (criterio areal común) · ${TOT.e6b} bloques moderado o severo`],
    'Situación predial':['No consolidable','porcentaje sobre bases distintas por volumen']};
  $('tKpi').innerHTML=`<thead><tr><th class="l">Indicador</th>${VK.map(k=>`<th>Vol. ${VM[k].vol} · ${VM[k].name}</th>`).join('')}<th>Integral</th></tr></thead><tbody>`+
    labels.map(l=>`<tr><td><b>${esc(l)}</b></td>${VK.map(k=>{const x=V[k].kpi.find(y=>y[0]===l)||['','—',''];return `<td>${esc(x[1])}<small>${esc(x[2])}</small></td>`;}).join('')}<td class="int">${esc((integ[l]||['—'])[0])}<small>${esc((integ[l]||['',''])[1])}</small></td></tr>`).join('')+'</tbody>';
  $('kpiFoot').textContent='Valores tal como los publica cada volumen en su cinta de indicadores. Integral: suma o media ponderada por superficie cuando el criterio es el mismo en los tres volúmenes.';
})();

(function aggTable(){
  const cols=[['Bloques',B=>n0(B.length)],['Superficie (ha)',B=>n2(sum(B,b=>b.ha))],['Altitud (m)',B=>n0(wmean(B,b=>b.alt,b=>b.ha))],['Pendiente (%)',B=>n2(wmean(B,b=>b.pend,b=>b.ha))],
    ['MSAVI 2024',B=>fmt(wmean(B,b=>b.ms,b=>b.ha),4)],['Índice PI',B=>n3(wmean(B,b=>b.ipp,b=>b.ha))],['ISL-MM',B=>n2(wmean(B,b=>b.mm,b=>b.ha))],['ISL-EH',B=>n2(wmean(B,b=>b.eh,b=>b.ha))],
    ['P (mm/año)',B=>n0(wmean(B,b=>b.P,b=>b.ha))],['IA',B=>n3(wmean(B,b=>b.AI,b=>b.ha))],['Cárcavas (n)',B=>n0(sum(B,b=>b.ncv))],['Cárcavas (km)',B=>n2(sum(B,b=>b.lcv)/1000)],['Población',B=>n0(sum(B,b=>b.pob))]];
  $('tAgg').innerHTML=`<thead><tr><th class="l">Provincia</th>${cols.map(c=>`<th>${c[0]}</th>`).join('')}</tr></thead><tbody>${VK.map(k=>`<tr><td class="l" style="white-space:nowrap"><i class="pvd" style="background:${VM[k].col}"></i>${VM[k].name}</td>${cols.map(c=>`<td>${c[1](V[k].B)}</td>`).join('')}</tr>`).join('')}</tbody><tfoot><tr><td class="l">Integral</td>${cols.map(c=>`<td>${c[1](ALL)}</td>`).join('')}</tr></tfoot>`;
})();

(function regTable(){
  $('tReg').innerHTML=`<thead><tr><th class="l">Volumen</th><th class="l">Regímenes territoriales (bloques, superficie e índice PI)</th></tr></thead><tbody>`+
    VK.map(k=>`<tr><td><b>Vol. ${VM[k].vol} · ${VM[k].name}</b></td><td>${V[k].regTxt.split('; ').map(esc).join('<br>')}</td></tr>`).join('')+'</tbody>';
})();

(function gapTable(){
  const themes=[...new Set(VK.flatMap(k=>V[k].gaps.map(g=>g[0])))];
  $('tGapI').innerHTML=`<thead><tr><th class="l">Tema</th>${VK.map(k=>`<th class="l">Vol. ${VM[k].vol} · ${VM[k].name}</th>`).join('')}<th class="l">Implicancia</th><th class="l">Acción</th><th>Prioridad</th></tr></thead><tbody>`+
    themes.map(t=>{const gs=VK.map(k=>V[k].gaps.find(g=>g[0]===t)); const any=gs.find(Boolean);
      const pr=any[4]; return `<tr><td><b>${esc(t)}</b></td>${gs.map(g=>`<td>${g?esc(g[1]):'<span class="muted">Sin vacío registrado</span>'}</td>`).join('')}<td>${esc(any[2])}</td><td>${esc(any[3])}</td><td style="text-align:center"><span class="pill ${pr==='Alta'?'p-cri':pr==='Media'?'p-adv':'p-neu'}">${esc(pr)}</span></td></tr>`;}).join('')+'</tbody>';
})();

/* =====================================================================
   INVENTARIO DE LOS 117 BLOQUES
   ===================================================================== */
function sortable(tbl,cols,getRows,render){
  let sk=null,sd=1;
  tbl.querySelector('thead').innerHTML='<tr>'+cols.map((c,i)=>`<th class="s${c.l?' l':''}" data-i="${i}">${c.h}<span class="car">▲▼</span></th>`).join('')+'</tr>';
  tbl.querySelector('thead').addEventListener('click',e=>{const th=e.target.closest('th'); if(!th) return; const i=+th.dataset.i;
    sd=sk===i?-sd:1; sk=i; tbl.querySelectorAll('thead th').forEach(x=>{if(+x.dataset.i===i) x.dataset.dir=sd; else delete x.dataset.dir;}); draw();});
  function draw(){let R=getRows(); if(sk!=null){const f=cols[sk].v; R=[...R].sort((a,b)=>{const x=f(a),y=f(b); return (typeof x==='string'?x.localeCompare(y,'es',{numeric:true}):(x??-1e18)-(y??-1e18))*sd;});} render(R);}
  return draw;
}
(function inventory(){
  const sp=$('iProv'), sd=$('iDist'), si=$('iPi'), q=$('iQ');
  sp.innerHTML='<option value="">Todas</option>'+VK.map(k=>`<option value="${k}">${VM[k].name}</option>`).join('');
  si.innerHTML='<option value="">Todas</option>'+PIN.map((p,i)=>`<option value="${i}">${p}</option>`).join('');
  const fillD=()=>{const ds=[...new Set(ALL.filter(b=>!sp.value||b.v===sp.value).map(b=>b.d))].sort((a,b)=>a.localeCompare(b,'es')); const cur=sd.value; sd.innerHTML='<option value="">Todos</option>'+ds.map(d=>`<option${d===cur?' selected':''}>${esc(d)}</option>`).join('');};
  fillD();
  const cols=[{h:'Bloque',l:1,v:b=>b.b,r:b=>`<td class="cod">${esc(b.b)}</td>`},{h:'Provincia',l:1,v:b=>VM[b.v].name,r:b=>`<td class="l"><i class="pvd" style="background:${VM[b.v].col}"></i>${VM[b.v].name}</td>`},
    {h:'Distrito',l:1,v:b=>b.d,r:b=>`<td class="l">${esc(b.d)}</td>`},{h:'Superficie (ha)',v:b=>b.ha,r:b=>`<td class="num">${n2(b.ha)}</td>`},
    {h:'Este (m)',v:b=>b.e,r:b=>`<td class="num">${n0(b.e)}</td>`},{h:'Norte (m)',v:b=>b.n,r:b=>`<td class="num">${n0(b.n)}</td>`},
    {h:'Altitud (m)',v:b=>b.alt,r:b=>`<td class="num">${n0(b.alt)}</td>`},{h:'Pendiente (%)',v:b=>b.pend,r:b=>`<td class="num">${n2(b.pend)}</td>`},
    {h:'MSAVI',v:b=>b.ms,r:b=>`<td class="num">${fmt(b.ms,4)}</td>`},{h:'Índice PI',v:b=>b.ipp,r:b=>`<td>${piPill(b.ipp)}</td>`},
    {h:'ISL-MM',v:b=>b.mm,r:b=>`<td>${n2(b.mm)}</td>`},{h:'ISL-EH',v:b=>b.eh,r:b=>`<td>${n2(b.eh)}</td>`},
    {h:'Cárcavas',v:b=>b.ncv,r:b=>`<td>${b.ncv}</td>`},{h:'Long. cárcavas (m)',v:b=>b.lcv,r:b=>`<td class="num">${n2(b.lcv)}</td>`},
    {h:'P (mm)',v:b=>b.P,r:b=>`<td class="num">${n0(b.P)}</td>`},{h:'IA',v:b=>b.AI,r:b=>`<td>${n3(b.AI)}</td>`},
    {h:'Población',v:b=>b.pob,r:b=>`<td class="num">${n0(b.pob)}</td>`},{h:'Riesgo social',v:b=>b.rk||'',r:b=>`<td>${b.rk?`<span class="pill ${b.rk==='Alto'?'p-cri':b.rk==='Medio'?'p-adv':'p-ok'}">${esc(b.rk)}</span>`:'—'}</td>`},
    {h:'Volumen',v:b=>VM[b.v].vol,r:b=>`<td><a class="golink" href="${VM[b.v].href}#fic">Vol. ${VM[b.v].vol} →</a></td>`}];
  const rows=()=>ALL.filter(b=>(!sp.value||b.v===sp.value)&&(!sd.value||b.d===sd.value)&&(si.value===''||piCls(b.ipp)===+si.value)&&(!q.value||b.b.toUpperCase().includes(q.value.trim().toUpperCase())));
  const tbl=$('iTbl');
  const draw=sortable(tbl,cols,rows,R=>{tbl.querySelector('tbody').innerHTML=R.map(b=>`<tr>${cols.map(c=>c.r(b)).join('')}</tr>`).join('');
    tbl.querySelector('tfoot').innerHTML=`<tr><td class="l">${R.length} bloques</td><td></td><td></td><td class="num">${n2(sum(R,b=>b.ha))}</td><td colspan="9"></td><td>${n0(sum(R,b=>b.ncv))}</td><td class="num">${n2(sum(R,b=>b.lcv))}</td><td colspan="2"></td><td class="num">${n0(sum(R,b=>b.pob))}</td><td colspan="2"></td></tr>`;
    $('iCnt').textContent=`${R.length} de ${ALL.length} bloques · ${n2(sum(R,b=>b.ha))} ha`;});
  sp.addEventListener('change',()=>{fillD(); draw();}); [sd,si].forEach(x=>x.addEventListener('change',draw)); q.addEventListener('input',draw);
  draw();
})();

/* =====================================================================
   METRADOS Y COSTOS — simulador
   ===================================================================== */
const P=C.par;
const PROVN={MORROPON:'Morropón',HUANCABAMBA:'Huancabamba',AYABACA:'Ayabaca'};
const TIPD=Object.fromEntries(C.paq.map(q=>[q.t,`${q.grp.split(' (')[0]} · ${q.deg}`]));
const LEV=[['cartera','Cartera completa'],['piloto','Piloto (5 bloques)'],['prov','Provincia'],['dist','Distrito'],['mc','Microcuenca'],['tip','Tipología'],['blk','Bloque']];
function selOptions(lv){
  const cnt=f=>{const m=new Map(); CB.forEach(b=>{const k=f(b); m.set(k,(m.get(k)||0)+1);}); return m;};
  if(lv==='cartera') return [['all',`Los ${CB.length} bloques`]];
  if(lv==='piloto') return [['pil',CB.filter(b=>b.pil==='Sí').map(b=>b.id).join(' · ')]];
  if(lv==='prov'){const m=cnt(b=>b.prov); return ['MORROPON','HUANCABAMBA','AYABACA'].map(p=>[p,`${PROVN[p]} (${m.get(p)} bloques)`]);}
  if(lv==='dist'){const m=cnt(b=>b.prov+'|'+b.distN); return [...m.keys()].sort((a,b)=>a.localeCompare(b,'es')).map(k=>{const [p,d]=k.split('|'); return [k,`${d} · ${PROVN[p]} (${m.get(k)})`];});}
  if(lv==='mc'){const m=cnt(b=>b.mc); return [...m.keys()].sort().map(k=>{const b=CB.find(x=>x.mc===k); return [k,`${k} · ${PROVN[b.prov]} (${m.get(k)})`];});}
  if(lv==='tip'){const m=cnt(b=>b.tip); return C.paq.map(q=>[q.t,`${q.t} · ${TIPD[q.t]} (${m.get(q.t)||0})`]);}
  return [...CB].sort((a,b)=>a.id.localeCompare(b.id,'es',{numeric:true})).map(b=>[b.id,`${b.id} · ${b.distN} · ${b.tip}`]);
}
function selIds(lv,v){
  const f={cartera:()=>true,piloto:b=>b.pil==='Sí',prov:b=>b.prov===v,dist:b=>b.prov+'|'+b.distN===v,mc:b=>b.mc===v,tip:b=>b.tip===v,blk:b=>b.id===v}[lv];
  return new Set(CB.filter(f).map(b=>b.id));
}
function selTitle(lv,v){
  if(lv==='cartera') return 'Cartera de 117 bloques';
  if(lv==='piloto') return 'Piloto de calibración';
  if(lv==='prov') return 'Provincia de '+PROVN[v];
  if(lv==='dist'){const [p,d]=v.split('|'); return `Distrito de ${d} · ${PROVN[p]}`;}
  if(lv==='mc') return 'Microcuenca '+v;
  if(lv==='tip') return `Tipología ${v} · ${TIPD[v]}`;
  const b=CB.find(x=>x.id===v); return `Bloque ${v} · ${b.distN}, ${PROVN[b.prov]}`;
}
const CRON={x:C.flu.slice(),f:[.30,.25,.20,.15,.10],b:[.10,.15,.20,.25,.30]};
const carcKm=sum(Object.values(C.carcML||{}),x=>x)/1000;
const ESC=[
  {id:'jornal',l:'Jornal rural (peón)',o:[[80,'S/ 80 por día'],[90,'S/ 90 por día'],[100,'S/ 100 por día (Excel)'],[110,'S/ 110 por día'],[120,'S/ 120 por día']],d:100},
  {id:'r1',l:'Densidad de plantación R1',o:[['paq','Según paquete: 650 / 722 pl/ha (Excel)'],['400','400 pl/ha (5 × 5 m)'],['625','625 pl/ha (4 × 4 m)'],['833','833 pl/ha (3 × 4 m)'],['1111','1,111 pl/ha (3 × 3 m)']],d:'paq'},
  {id:'r2f',l:'Enriquecimiento R2',o:[['paq','50 % de la densidad R1 (Excel)'],['0.3','30 % de la densidad R1'],['0.7','70 % de la densidad R1'],['1','100 % de la densidad R1']],d:'paq'},
  {id:'ez',l:'Espaciamiento de zanjas',o:[['paq','2.78 m · 3,600 m/ha (Excel, ANIN)'],['4','4.00 m · 2,500 m/ha'],['5','5.00 m · 2,000 m/ha'],['6.6667','6.67 m · 1,500 m/ha'],['10','10.00 m · 1,000 m/ha']],d:'paq'},
  {id:'et',l:'Espaciamiento de terrazas',o:[['paq','8.33 m · 1,200 m/ha (Excel, ANIN)'],['10','10.00 m · 1,000 m/ha'],['12.5','12.50 m · 800 m/ha'],['16.6667','16.67 m · 600 m/ha']],d:'paq'},
  {id:'cer',l:'Cerco perimétrico',o:[['paq','100 % del perímetro (Excel)'],['0.75','75 % del perímetro'],['0.5','50 % del perímetro'],['0.25','25 % del perímetro'],['0','Sin cerco']],d:'paq'},
  {id:'carc',l:'Control de cárcavas',o:[['no','Sin metrado (Excel: columna AD vacía)'],['inv',`Capa de cárcavas de los volúmenes (${n2(carcKm)} km)`]],d:'no'},
  {id:'gg',l:'Gastos generales',o:[[P.P_GG,'15.08 % (referencia ANIN, Excel)'],[0.10,'10 % del CD'],[0.12,'12 % del CD'],[0.18,'18 % del CD'],[0.20,'20 % del CD']],d:P.P_GG},
  {id:'sup',l:'Supervisión',o:[[0,'0 % · por definir (Excel)'],[0.03,'3 % del CD'],[0.05,'5 % del CD'],[0.08,'8 % del CD']],d:0},
  {id:'et_pct',l:'Expediente técnico',o:[[0,'0 % · por definir (Excel)'],[0.02,'2 % del CD'],[0.03,'3 % del CD'],[0.05,'5 % del CD']],d:0},
  {id:'cron',l:'Cronograma de inversión',o:[['x','20 % anual en 5 años (Excel)'],['f','30 · 25 · 20 · 15 · 10 %'],['b','10 · 15 · 20 · 25 · 30 %']],d:'x'},
  {id:'tsd',l:'Tasa social de descuento',o:[[0.08,'8 % (MEF, Anexo 11)'],[0.06,'6 % · sensibilidad'],[0.10,'10 % · sensibilidad'],[0.12,'12 % · sensibilidad']],d:0.08}];
const store={get(k){try{return JSON.parse(localStorage.getItem('inp-cost-'+k));}catch(e){return null;}},set(k,v){try{localStorage.setItem('inp-cost-'+k,JSON.stringify(v));}catch(e){}}};
(function buildCtrls(){
  const nv=$('cNivel'); nv.innerHTML=LEV.map(([v,l])=>`<option value="${v}">${l}</option>`).join('');
  const box=$('cEsc');
  box.insertAdjacentHTML('beforeend',ESC.map(e=>`<div class="f"><label for="e_${e.id}">${e.l}</label><select id="e_${e.id}">${e.o.map(([v,l])=>`<option value="${v}"${String(v)===String(e.d)?' selected':''}>${esc(l)}</option>`).join('')}</select></div>`).join(''));
  const saved=store.get('esc')||{};
  ESC.forEach(e=>{const s=$('e_'+e.id); if(saved[e.id]!=null&&[...s.options].some(o=>o.value===String(saved[e.id]))) s.value=String(saved[e.id]); s.addEventListener('change',()=>{saveEsc(); renderCost();});});
  const amb=store.get('amb');
  nv.value=amb&&LEV.some(l=>l[0]===amb.lv)?amb.lv:'cartera'; fillSel(amb&&amb.v);
  nv.addEventListener('change',()=>{fillSel(); saveAmb(); renderCost();});
  $('cSel').addEventListener('change',()=>{saveAmb(); renderCost();});
  $('cReset').addEventListener('click',()=>{ESC.forEach(e=>{$('e_'+e.id).value=String(e.d);}); saveEsc(); renderCost();});
})();
function fillSel(v){const s=$('cSel'), o=selOptions($('cNivel').value); s.innerHTML=o.map(([k,l])=>`<option value="${esc(k)}">${esc(l)}</option>`).join(''); if(v&&o.some(x=>x[0]===v)) s.value=v; s.disabled=o.length<2;}
function saveAmb(){store.set('amb',{lv:$('cNivel').value,v:$('cSel').value});}
function saveEsc(){const o={}; ESC.forEach(e=>o[e.id]=$('e_'+e.id).value); store.set('esc',o);}
function scenario(){
  const S=ENG.defaults(), g=id=>$('e_'+id).value;
  S.jornal=+g('jornal'); S.r1=g('r1'); S.r2f=g('r2f'); S.ez=g('ez'); S.et=g('et'); S.cer=g('cer'); S.carc=g('carc');
  S.gg=+g('gg'); S.sup=+g('sup'); S.et_pct=+g('et_pct'); S.cron=CRON[g('cron')]; S.tsd=+g('tsd');
  let mod=0; ESC.forEach(e=>{const s=$('e_'+e.id), m=s.value!==String(e.d); s.classList.toggle('mod',m); if(m) mod++;});
  const st=$('cState'); st.className='c-state '+(mod?'m':'x'); st.textContent=mod?`Escenario modificado · ${mod} supuesto${mod>1?'s':''} distinto${mod>1?'s':''} del Excel`:'Valores del Excel v1.3';
  return S;
}
function renderCost(){
  const lv=$('cNivel').value, v=$('cSel').value, ids=selIds(lv,v), S=scenario();
  const R=ENG.run(ids,S), T=R.t, cart=ENG.run(IDS_ALL,S), techo=P.P_TECHO;
  const nT=new Map(); R.rows.forEach(r=>nT.set(r.tip,(nT.get(r.tip)||0)+1));
  $('cTitle').textContent=selTitle(lv,v);
  $('cSub').textContent=`${R.rows.length} bloque${R.rows.length>1?'s':''} · ${n3(T.ha)} ha en el motor · ${[...nT].sort().map(([t,n])=>`${t} (${n})`).join(', ')}`;
  const k=[
    ['Superficie',n3(T.ha),'ha de bloque (motor)'],
    ['R1 + R2',n2(T.E+T.F),`ha · R1 ${n2(T.E)} · R2 ${n2(T.F)}`],
    ['Plantas',n0(T.J),`und · ${n0(T.plantones)} plantones con recalce`],
    ['Zanjas · terrazas',`${n2(T.K)} · ${n2(T.L)}`,'km de zanjas · km de terrazas'],
    ['Cerco',n0(T.M),'m de cerco perimétrico'],
    ['Inversión total',solM(T.AA),`CD ${solM(T.Y)} + indirectos ${solM(T.Z)}`,1],
    ['Por hectárea',sol(T.ha?T.AA/T.ha:0),'S/ de inversión por ha de bloque'],
    ['O&M 15 años',solM(R.ind.om15),`primer año ${solM(R.om1)}`],
    ['VAC social',solM(R.ind.vacS),`TSD ${pct(S.tsd,0)} · 20 años`],
    ['Costo-efectividad',sol(R.ind.ceha),'VAC social por ha plantada']];
  $('cKpi').innerHTML=k.map(([a,b,c,h])=>`<div${h?' class="hl"':''}><dt>${a}</dt><dd>${b}<small>${c}</small></dd></div>`).join('');
  const haI=sum([...ids],id=>intHa[id]||0);
  $('cSup').innerHTML=`<div><dt>Superficie en el motor</dt><dd>${n3(T.ha)} <small>ha</small></dd></div><div><dt>Superficie en los volúmenes</dt><dd>${n2(haI)} <small>ha · diferencia ${fmt(T.ha-haI,3)} ha</small></dd></div>
    <div><dt>Peso en la cartera</dt><dd>${pct(cart.t.AA?T.AA/cart.t.AA:0)} <small>de la inversión de los 117 bloques</small></dd></div>
    <div><dt>Inversión + O&amp;M 15 años</dt><dd>${pct((T.AA+R.ind.om15)/techo)} <small>del presupuesto total S/ ${n0(techo)}</small></dd></div>`;
  // presupuesto
  const CU=R.CU, ind=S.gg+S.sup+S.et_pct;
  const flR=T.pl_kg?T.S/T.pl_kg:0;
  const lines=[['sub','Infraestructura natural (verde)'],
    ['v','Producción de plantones en vivero propio (incluye recalce de 10 %)','und',T.plantones,CU.PLANTON.h,T.P],
    ['v','Plantación y labores por planta (hoyado, transporte, plantación, riegos, recalce)','und',T.J,CU.PLANTA.h,T.Q],
    ['v','Labores por hectárea (trazo, limpieza, control sanitario, deshierbes)','ha',T.E+T.F,CU.HA.h,T.R],
    ['v','Flete rural de plantones e insumos (a pie o acémila)','kg',T.pl_kg,flR,T.S],
    ['v','Flete terrestre de insumos','kg',T.ins_kg,P.P_FT_KG,T.T],
    ['v','Cerco perimétrico de alambre de púas','m',T.M,CU.CERCO.h,T.W],
    ['sub','Conservación de suelos (marrón)'],
    ['m','Zanjas de infiltración 0.40 × 0.30 m','km',T.K,CU.ZANJA.h,T.U],
    ['m','Terrazas de formación lenta','km',T.L,CU.TERRAZA.h,T.V],
    ['m','Control de cárcavas con diques de piedra canteada','km',T.N,CU.CARCAVA.h,T.X]];
  const pc=x=>T.Y?pct(x/T.Y):'—';
  $('cBud').innerHTML=`<thead><tr><th class="l">Partida</th><th>Und</th><th class="num">Metrado</th><th class="num">CU (S/)</th><th class="num">Parcial (S/)</th><th class="num">% CD</th></tr></thead><tbody>`+
    lines.map(l=>l[0]==='sub'?`<tr class="sub"><td colspan="6">${l[1]}</td></tr>`:`<tr><td class="l"><span class="sw" style="background:${l[0]==='v'?'var(--c-ver)':'var(--c-mar)'}"></span>${l[1]}</td><td>${l[2]}</td><td class="num">${fmt(l[3],l[2]==='km'?3:0)}</td><td class="num">${fmt(l[4],l[4]<10?4:2)}</td><td class="num">${n0(l[5])}</td><td class="num">${pc(l[5])}</td></tr>`).join('')+
    `<tr class="tot"><td class="l">Costo directo</td><td></td><td></td><td></td><td class="num">${n0(T.Y)}</td><td class="num">100.0 %</td></tr>
     <tr><td class="l">Gastos generales y obras provisionales (${pct(S.gg,2)})</td><td></td><td></td><td></td><td class="num">${n0(T.Y*S.gg)}</td><td class="num">${pc(T.Y*S.gg)}</td></tr>
     <tr><td class="l">Supervisión (${pct(S.sup,0)})${S.sup?'':' · por definir'}</td><td></td><td></td><td></td><td class="num">${n0(T.Y*S.sup)}</td><td class="num">${pc(T.Y*S.sup)}</td></tr>
     <tr><td class="l">Expediente técnico (${pct(S.et_pct,0)})${S.et_pct?'':' · por definir'}</td><td></td><td></td><td></td><td class="num">${n0(T.Y*S.et_pct)}</td><td class="num">${pc(T.Y*S.et_pct)}</td></tr>
     <tr class="big"><td class="l">Inversión total a precios de mercado</td><td></td><td></td><td></td><td class="num">${n0(T.AA)}</td><td class="num">${pc(T.AA)}</td></tr>
     <tr><td class="l">· infraestructura natural (verde) con indirectos</td><td></td><td></td><td></td><td class="num">${n0(T.AG)}</td><td class="num">${pct(T.AA?T.AG/T.AA:0)} inv.</td></tr>
     <tr><td class="l">· conservación de suelos (marrón) con indirectos</td><td></td><td></td><td></td><td class="num">${n0(T.AH)}</td><td class="num">${pct(T.AA?T.AH/T.AA:0)} inv.</td></tr>
     <tr class="sub"><td colspan="6">Inversión a precios sociales</td></tr>
     <tr><td class="l">Mano de obra no calificada (CD) × ${P.P_FC_MONC}</td><td></td><td></td><td></td><td class="num">${n0(T.AC*P.P_FC_MONC)}</td><td class="num">${pc(T.AC)} CD</td></tr>
     <tr><td class="l">Mano de obra calificada (CD) × ${P.P_FC_MOC}</td><td></td><td></td><td></td><td class="num">${n0(T.AD*P.P_FC_MOC)}</td><td class="num">${pc(T.AD)} CD</td></tr>
     <tr><td class="l">Bienes, servicios e indirectos ÷ 1.18</td><td></td><td></td><td></td><td class="num">${n0((T.AE+T.Z)*P.P_FC_BIEN)}</td><td class="num"></td></tr>
     <tr class="big"><td class="l">Inversión total a precios sociales</td><td></td><td></td><td></td><td class="num">${n0(T.AF)}</td><td class="num">${pct(T.AA?T.AF/T.AA:0)} inv.</td></tr></tbody>`;
  // composición
  const comp=lines.filter(l=>l[0]!=='sub').map(l=>({n:l[1].split(' (')[0],c:l[0],v:l[5]}));
  const W=520,L=0,rh=34,H=comp.length*rh+4,mx=Math.max(1,...comp.map(c=>c.v));
  $('cCompo').innerHTML=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Composición del costo directo">${comp.map((c,i)=>`<text x="0" y="${i*rh+12}" style="fill:var(--ink-2);font:500 11.5px 'Source Sans 3',Arial,sans-serif">${esc(c.n)}</text>
    <rect x="0" y="${i*rh+17}" width="${Math.max(c.v>0?2:0,c.v/mx*(W-120)).toFixed(1)}" height="11" rx="2" style="fill:${c.c==='v'?'var(--c-ver)':'var(--c-mar)'}"></rect>
    <text x="${(c.v/mx*(W-120)+8).toFixed(1)}" y="${i*rh+27}" style="fill:var(--ink);font:600 11px 'IBM Plex Mono',monospace">${T.Y?pct(c.v/T.Y):'—'} · ${solM(c.v)}</text>`).join('')}</svg>`;
  // flujo
  const F=R.flow, fw=1100, fh=270, bl=64, bb=26, cw=(fw-bl-10)/F.length, fm=Math.max(1,...F.map(f=>Math.max(f.tot,f.totS)));
  const step=[1,2,2.5,5,10].map(m=>m*Math.pow(10,Math.floor(Math.log10(fm/4)))).find(s=>fm/s<=5)||fm/4;
  const ty=v=>fh-bb-v/fm*(fh-bb-12);
  let grid=''; for(let t=0;t<=fm+1e-6;t+=step) grid+=`<line x1="${bl}" x2="${fw-6}" y1="${ty(t)}" y2="${ty(t)}" style="stroke:var(--line);stroke-width:1"></line><text x="${bl-8}" y="${ty(t)+4}" text-anchor="end" style="fill:var(--ink-3);font:500 10.5px 'IBM Plex Mono',monospace">${fmt(t/1e6,t/1e6<10&&step<1e6?1:0)} M</text>`;
  const cols=F.map((f,i)=>{const x=bl+i*cw+cw*.18, w=cw*.64; return `<g data-y="${i}"><rect x="${x}" y="${ty(f.inv)}" width="${w}" height="${fh-bb-ty(f.inv)}" style="fill:var(--anin)"></rect><rect x="${x}" y="${ty(f.inv+f.om)}" width="${w}" height="${ty(f.inv)-ty(f.inv+f.om)}" style="fill:var(--c-om)"></rect><rect x="${bl+i*cw}" y="12" width="${cw}" height="${fh-bb-12}" style="fill:transparent"></rect><text x="${x+w/2}" y="${fh-8}" text-anchor="middle" style="fill:var(--ink-3);font:500 10px 'IBM Plex Mono',monospace">${f.y}</text></g>`;}).join('');
  const line=F.map((f,i)=>`${i?'L':'M'}${(bl+i*cw+cw/2).toFixed(1)} ${ty(f.totS).toFixed(1)}`).join(' ');
  $('cFlow').innerHTML=`<div style="overflow-x:auto"><svg viewBox="0 0 ${fw} ${fh}" style="min-width:600px" role="img" aria-label="Flujo anual de costos">${grid}${cols}<path d="${line}" style="fill:none;stroke:var(--brick);stroke-width:2"></path>${F.map((f,i)=>`<circle cx="${(bl+i*cw+cw/2).toFixed(1)}" cy="${ty(f.totS).toFixed(1)}" r="2.6" style="fill:var(--brick)"></circle>`).join('')}<line x1="${bl+5*cw}" x2="${bl+5*cw}" y1="8" y2="${fh-bb}" style="stroke:var(--ink-3);stroke-dasharray:3 3"></line><text x="${bl+5*cw+6}" y="20" style="fill:var(--ink-3);font:600 10px 'IBM Plex Mono',monospace">INICIO DE O&amp;M</text></svg></div>`;
  $('cFlow').querySelectorAll('g[data-y]').forEach(g=>{const f=F[+g.dataset.y]; g.addEventListener('mousemove',e=>showTip(e,`<b>Año ${f.y}</b>${tr('Inversión',sol(f.inv))}${tr('O&M',sol(f.om))}${tr('Total mercado',sol(f.tot))}${tr('Total social',sol(f.totS))}`)); g.addEventListener('mouseleave',hideTip);});
  $('cInd').innerHTML=`<thead><tr><th class="l">Indicador (horizonte de 20 años, TSD ${pct(S.tsd,0)})</th><th class="num">Valor</th></tr></thead><tbody>
    <tr><td class="l">Valor actual de la inversión (precios de mercado)</td><td class="num">${sol(R.ind.vaInv)}</td></tr>
    <tr><td class="l">Valor actual de la O&amp;M (precios de mercado)</td><td class="num">${sol(R.ind.vaOM)}</td></tr>
    <tr class="tot"><td class="l">VAC a precios de mercado</td><td class="num">${sol(R.ind.vacP)}</td></tr>
    <tr class="big"><td class="l">VAC a precios sociales</td><td class="num">${sol(R.ind.vacS)}</td></tr>
    <tr><td class="l">Costo-efectividad: VAC social por ha plantada (R1 + R2)</td><td class="num">${sol(R.ind.ceha)}</td></tr>
    <tr><td class="l">Costo-efectividad: VAC social por ha de bloque</td><td class="num">${sol(R.ind.ceblk)}</td></tr>
    <tr><td class="l">VANS y TIR social</td><td class="num"><span class="pill p-adv">Pendiente de beneficios</span></td></tr></tbody>`;
  // O&M
  const U={ha:'ha plantada',pl:'planta',kz:'km de zanja',kt:'km de terraza',kc:'km de cárcava'};
  $('cOm').innerHTML=`<thead><tr><th class="l">Componente</th><th class="l">Actividad (ANIN, Anexo 11)</th><th>Inductor</th><th class="num">Cantidad</th><th class="num">Tarifa anual (S/)</th><th>Años de aplicación</th><th class="num">Costo año 1 (S/)</th><th class="num">Costo 15 años (S/)</th></tr></thead><tbody>`+
    R.om.map(o=>`<tr><td class="l">${esc(o.comp)}</td><td class="l" style="white-space:normal;min-width:240px">${esc(o.act)}</td><td>${U[o.ind]}</td><td class="num">${fmt(o.q,o.ind==='pl'?0:2)}</td><td class="num">${fmt(o.f,4)}</td><td>${o.anios} de 15</td><td class="num">${n0(o.anual)}</td><td class="num">${n0(o.anual*o.anios)}</td></tr>`).join('')+
    `</tbody><tfoot><tr><td class="l" colspan="6">Total</td><td class="num">${n0(R.om1)}</td><td class="num">${n0(R.ind.om15)}</td></tr></tfoot>`;
  // bloques
  drawBlk(R.rows);
  // flujo anual
  $('cFlowT').innerHTML=`<thead><tr><th>Año</th><th>Fase</th><th class="num">Inversión</th><th class="num">O&amp;M</th><th class="num">Total mercado</th><th class="num">Inversión social</th><th class="num">O&amp;M social</th><th class="num">Total social</th></tr></thead><tbody>`+
    F.map(f=>`<tr><td>${f.y}</td><td>${f.y<=P.P_ANIOS_EJ?'Ejecución':'O&amp;M año '+(f.y-P.P_ANIOS_EJ)}</td><td class="num">${n0(f.inv)}</td><td class="num">${n0(f.om)}</td><td class="num">${n0(f.tot)}</td><td class="num">${n0(f.invS)}</td><td class="num">${n0(f.omS)}</td><td class="num">${n0(f.totS)}</td></tr>`).join('')+
    `</tbody><tfoot><tr><td colspan="2">Total</td>${['inv','om','tot','invS','omS','totS'].map(k=>`<td class="num">${n0(sum(F,f=>f[k]))}</td>`).join('')}</tr></tfoot>`;
  renderSup(S,R);
}
let drawBlk=()=>{};
(function blkTable(){
  const by=Object.fromEntries(CB.map(b=>[b.id,b]));
  const cols=[{h:'Bloque',l:1,v:r=>r.id,r:r=>`<td class="cod">${esc(r.id)}</td>`},{h:'Provincia',l:1,v:r=>PROVN[by[r.id].prov],r:r=>`<td class="l"><i class="pvd" style="background:${VM[XV[by[r.id].prov]].col}"></i>${PROVN[by[r.id].prov]}</td>`},
    {h:'Distrito',l:1,v:r=>by[r.id].distN,r:r=>`<td class="l">${esc(by[r.id].distN)}</td>`},{h:'Microcuenca',l:1,v:r=>by[r.id].mc,r:r=>`<td class="l mono" style="font-size:12px">${esc(by[r.id].mc)}</td>`},
    {h:'Tipología',v:r=>r.tip,r:r=>`<td title="${esc(TIPD[r.tip])}">${r.tip}</td>`},{h:'Piloto',v:r=>by[r.id].pil,r:r=>`<td>${by[r.id].pil==='Sí'?'<span class="pill p-neu">Sí</span>':''}</td>`},
    {h:'Área (ha)',v:r=>r.ha,r:r=>`<td class="num">${n3(r.ha)}</td>`},{h:'R1 (ha)',v:r=>r.E,r:r=>`<td class="num">${n2(r.E)}</td>`},{h:'R2 (ha)',v:r=>r.F,r:r=>`<td class="num">${n2(r.F)}</td>`},
    {h:'Plantas',v:r=>r.J,r:r=>`<td class="num">${n0(r.J)}</td>`},{h:'Zanjas (km)',v:r=>r.K,r:r=>`<td class="num">${n2(r.K)}</td>`},{h:'Terrazas (km)',v:r=>r.L,r:r=>`<td class="num">${n2(r.L)}</td>`},
    {h:'Cerco (m)',v:r=>r.M,r:r=>`<td class="num">${n0(r.M)}</td>`},{h:'Cárcavas (km)',v:r=>r.N,r:r=>`<td class="num">${n3(r.N)}</td>`},
    {h:'Acarreo (m)',v:r=>by[r.id].dac,r:r=>`<td class="num">${n0(by[r.id].dac)}</td>`},
    {h:'Inversión (S/)',v:r=>r.AA,r:r=>`<td class="num">${n0(r.AA)}</td>`},{h:'S/ por ha',v:r=>r.AB,r:r=>`<td class="num">${n0(r.AB)}</td>`},
    {h:'Verde / marrón',v:r=>r.AA?r.AG/r.AA:0,r:r=>`<td><div class="cbar" title="Verde ${pct(r.AA?r.AG/r.AA:0)}"><i style="width:${r.AA?r.AG/r.AA*100:0}%;background:var(--c-ver)"></i><i style="width:${r.AA?r.AH/r.AA*100:0}%;background:var(--c-mar)"></i></div></td>`}];
  let rows=[]; const tbl=$('cBlk');
  const draw=sortable(tbl,cols,()=>rows,R=>{tbl.querySelector('tbody').innerHTML=R.map(r=>`<tr data-id="${esc(r.id)}">${cols.map(c=>c.r(r)).join('')}</tr>`).join('');
    tbl.querySelector('tfoot').innerHTML=`<tr><td class="l" colspan="6">${R.length} bloques</td><td class="num">${n3(sum(R,r=>r.ha))}</td><td class="num">${n2(sum(R,r=>r.E))}</td><td class="num">${n2(sum(R,r=>r.F))}</td><td class="num">${n0(sum(R,r=>r.J))}</td><td class="num">${n2(sum(R,r=>r.K))}</td><td class="num">${n2(sum(R,r=>r.L))}</td><td class="num">${n0(sum(R,r=>r.M))}</td><td class="num">${n3(sum(R,r=>r.N))}</td><td></td><td class="num">${n0(sum(R,r=>r.AA))}</td><td class="num">${n0(sum(R,r=>r.ha)?sum(R,r=>r.AA)/sum(R,r=>r.ha):0)}</td><td></td></tr>`;});
  tbl.querySelector('tbody').addEventListener('click',e=>{const tr=e.target.closest('tr[data-id]'); if(!tr) return; $('cNivel').value='blk'; fillSel(tr.dataset.id); saveAmb(); renderCost();
    const t=$('cTitle'); const top=t.getBoundingClientRect().top+scrollY-110; scrollTo({top,behavior:'smooth'});});
  drawBlk=R=>{rows=R; draw();};
})();

/* ---------- supuestos ---------- */
function renderSup(S,R){
  if(!renderSup.once){renderSup.once=1;
    const L=C.leeme.filter(x=>x[1]);
    $('sLeeme').innerHTML=L.map(([a,b])=>`<div class="find ${/LIMITACIONES/.test(a)?'adv':/VALIDACI/.test(a)?'ok':''}"><i></i><div class="body"><h3>${esc(a.charAt(0)+a.slice(1).toLowerCase())}</h3><p>${esc(b)}</p></div></div>`).join('')+
      `<div class="find adv"><i></i><div class="body"><h3>Lectura rápida del motor</h3><ul class="fb-list" style="margin:0">${C.lect.map(t=>`<li>${esc(String(t).replace(/^\d\.\s*/,''))}</li>`).join('')}</ul></div></div>`;
    const base=ENG.run(IDS_ALL,ENG.defaults());
    $('sPaq').innerHTML=`<thead><tr><th>Tipología</th><th class="l">Grupo de ecosistema</th><th>Degradación</th><th>Densidad R1 (pl/ha)</th><th>R2 (pl/ha)</th><th>Zanjas (m/ha)</th><th>Terrazas (m/ha)</th><th>Cerco</th><th>Bloques</th><th>Área (ha)</th><th>Inversión (S/)</th><th>S/ por ha</th><th class="l">Fuente de los coeficientes</th></tr></thead><tbody>`+
      C.paq.map(q=>{const rr=base.rows.filter(r=>r.tip===q.t), ha=sum(rr,r=>r.ha), inv=sum(rr,r=>r.AA);
        return `<tr><td style="text-align:center"><b class="mono">${q.t}</b></td><td>${esc(q.grp)}</td><td style="text-align:center">${q.deg}</td><td style="text-align:center">${n0(q.r1)}</td><td style="text-align:center">${n0(q.r2)}</td><td style="text-align:center">${q.ez?n0(10000/q.ez):'—'}</td><td style="text-align:center">${q.et?n0(10000/q.et):'—'}</td><td style="text-align:center">${pct(q.cer,0)}</td><td style="text-align:center">${rr.length}</td><td style="text-align:right">${n3(ha)}</td><td style="text-align:right">${n0(inv)}</td><td style="text-align:right">${n0(ha?inv/ha:0)}</td><td>${esc(q.src)}</td></tr>`;}).join('')+'</tbody>';
    $('sPar').innerHTML=`<thead><tr><th class="l">Nombre</th><th class="l">Parámetro</th><th>Valor</th><th>Unidad</th><th class="l">Fuente o nota</th></tr></thead><tbody>`+
      C.parT.map(p=>p.sec?`<tr><td colspan="5" style="background:var(--sunken);font:600 11px/1.3 'IBM Plex Mono',monospace;letter-spacing:.09em;color:var(--ink-2)">${esc(p.sec)}</td></tr>`:
        `<tr><td class="mono" style="font-size:12px">${esc(p.n)}</td><td>${esc(p.p)}</td><td style="text-align:right;white-space:nowrap">${typeof p.v==='number'?(Math.abs(p.v)<1&&p.v!==0?fmt(p.v,4):fmt(p.v,p.v%1?4:0)):esc(p.v)}</td><td style="white-space:nowrap">${esc(p.u)}</td><td>${esc(p.f)}</td></tr>`).join('')+'</tbody>';
  }
  const U=[['PLANTON','Plantón producido en vivero propio','und'],['PLANTA','Planta instalada (sin plantón)','und'],['HA','Hectárea plantada R1 + R2','ha'],['CERCO','Cerco perimétrico','m'],['ZANJA','Zanja de infiltración 0.40 × 0.30 m','km'],['TERRAZA','Terraza de formación lenta','km'],['CARCAVA','Cárcava con diques de piedra canteada','km']];
  $('sCuTag').textContent=`Hoja CU_COMP · jornal del escenario: S/ ${S.jornal} por día`;
  $('sCu').innerHTML=`<thead><tr><th class="l">Costo unitario</th><th>Und</th><th class="num">CU total (S/)</th><th class="num">MO no calificada</th><th class="num">MO calificada</th><th class="num">Bienes y servicios</th><th class="l">Composición · MO no calificada / MO calificada / bienes</th></tr></thead><tbody>`+
    U.map(([k,n,u])=>{const c=R.CU[k], b=c.h-c.nc-c.c; return `<tr><td class="l">${n}</td><td>${u}</td><td class="num">${fmt(c.h,c.h<10?4:2)}</td><td class="num">${fmt(c.nc,c.h<10?4:2)}</td><td class="num">${fmt(c.c,c.h<10?4:2)}</td><td class="num">${fmt(b,c.h<10?4:2)}</td><td><div class="cbar" style="min-width:140px"><i style="width:${c.nc/c.h*100}%;background:var(--moss-2)"></i><i style="width:${c.c/c.h*100}%;background:var(--anin-2)"></i><i style="width:${b/c.h*100}%;background:var(--c-ind)"></i></div></td></tr>`;}).join('')+
    `</tbody>`;
}

/* ---------- arranque ---------- */
(function boot(){
  const h=(location.hash||'').slice(1); let p=tabs.some(t=>t.dataset.p===h)?h:null;
  if(!p){try{const s=localStorage.getItem('inp-int-tab'); if(tabs.some(t=>t.dataset.p===s)) p=s;}catch(e){}}
  show(p||'pano',false);
})();
