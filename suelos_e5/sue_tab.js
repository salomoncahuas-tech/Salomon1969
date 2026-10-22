
/* =====================================================================
   7. PESTAÑA SUELOS (E5) — Estudios básicos
   ===================================================================== */
const ORD=t=>t?t.split(' /')[0]:'—';
const PDM={C:6,D:11.5,E:20,F:37.5,G:62.5,H:80};
const pendN=c=>c.pn!=null?c.pn:(PDM[c.pc]??null);
const TIPO={B:'Bloque',R:'Referencia',S:'Lote SUS'};
const GRP={1:'Grupo 1 · ócrico sin horizonte de diagnóstico',2:'Grupo 2 · ócrico con horizonte cámbico',3:'Grupo 3 · epipedón mólico'};
const GRPs={1:'G1 Entisoles',2:'G2 Inceptisoles',3:'G3 Mollisoles'};
const peCls=c=>c.pcm==null?null:c.pcm<25?0:c.pcm<50?1:c.pcm<=100?2:3;
const PEN=['Muy superficial (< 25 cm)','Superficial (25–50 cm)','Moderadamente profundo (50–100 cm)','Profundo (> 100 cm)'];
const PEC=['var(--pi5)','var(--pi4)','var(--pi2)','var(--pi1)'];
const ERC={1:'var(--pi1)',2:'var(--pi2)',3:'var(--pi3)',4:'var(--pi5)'};
const ERN={1:'Muy ligera',2:'Ligera',3:'Moderada',4:'Severa'};
const GC={1:'var(--sg1)',2:'var(--sg2)',3:'var(--sg3)'};
const cumG=s=>{const m=String(s||'').match(/^[A-Z]/);return m?m[0]:'—';};
const CUMC={X:'var(--cu-x)',F:'var(--cu-f)',P:'var(--cu-p)',C:'var(--cu-c)',A:'var(--cu-a)'};
const CUMN={X:'Protección (X)',F:'Producción forestal (F)',P:'Pastoreo (P)',C:'Cultivo permanente (C)',A:'Cultivo en limpio (A)'};
const ICLS=[[0,.1,'Muy lenta'],[.1,.5,'Lenta'],[.5,2,'Mod. lenta'],[2,6.3,'Moderada'],[6.3,12.7,'Mod. rápida'],[12.7,25.4,'Rápida'],[25.4,200,'Muy rápida']];
const infOf=c=>SU.I.find(s=>s.cal===c.c)||null;
const infUnit=c=>SU.I.find(s=>s.u===c.u)||null;
const siteDist=s=>{const c=CAL.find(x=>x.c===s.cal);return c?c.dist:'—';};
const siteProv=s=>{const c=CAL.find(x=>x.c===s.cal);return c?c.prov:'—';};

/* ---------- aptitud de las MRR-CCC según la matriz suelo–intervención (Cuadro 21) ---------- */
function apt(c){
  const p=pendN(c), shallow=(c.ccm!=null&&c.ccm<50)||(c.pcm!=null&&c.pcm<50), dens=/Dénsico/.test(c.con)&&c.ccm!=null&&c.ccm<=60;
  let z,zc;
  if(c.g===3){z='No prioritaria · infiltra rápido';zc='p-neu';}
  else if(p!=null&&p>50){z='No · pendiente > 50 %';zc='p-cri';}
  else if(p!=null&&p<15){z='Fuera de rango · < 15 %';zc='p-neu';}
  else if(shallow){z='Restringida · contacto < 50 cm';zc='p-adv';}
  else if(dens){z='Restringida · dénsico ≤ 60 cm';zc='p-adv';}
  else {z='Compatible';zc='p-ok';}
  const t=(p!=null&&p>=15&&p<=40&&!shallow)?['Compatible','p-ok']:['—',''];
  const k=p!=null&&p>50?['Prioritaria','p-cri']:(c.eroN>=3?['Recomendable','p-adv']:['—','']);
  return {p,z,zc,t,k};
}

/* ---------- utilidades de color y perfiles ---------- */
const h2r=h=>{const n=parseInt(h.slice(1),16);return [n>>16&255,n>>8&255,n&255];};
const mixc=(a,b,t)=>{const x=h2r(a),y=h2r(b);return '#'+x.map((v,i)=>Math.round(v+(y[i]-v)*t).toString(16).padStart(2,'0')).join('');};
function hzList(c){
  const hz=c.hz, sam=c.sam||[], k=Math.min(hz.length,sam.length);
  let bottom=Math.max(c.pcm||0,c.ccm||0,sam.length?sam[sam.length-1][1]:0); const unk=!bottom; if(unk) bottom=60;
  const L=[]; for(let i=0;i<k;i++) L.push({n:hz[i],t:sam[i][0],b:sam[i][1],est:false});
  const st=k?sam[k-1][1]:0, r=hz.length-k;
  if(r>0){const end=Math.max(bottom,st+12*r), step=(end-st)/r; for(let i=0;i<r;i++) L.push({n:hz[k+i],t:Math.round(st+i*step),b:Math.round(st+(i+1)*step),est:true});}
  return {L,bottom:Math.max(bottom,L.length?L[L.length-1].b:0),unk};
}
function hzCol(c,n,first){
  const base=c.col||'#8a6a4a', s=n.replace(/\d/g,'');
  if(/^O/.test(s)) return '#2b231d';
  if(first) return base;
  if(s==='Ab') return mixc(base,'#000000',.12);
  if(/^A/.test(s)&&!/^A[BC]/.test(s)) return mixc(base,'#c9a878',.12);
  if(/^(AB|BA|AC|CA)/.test(s)) return mixc(base,'#c9a878',.3);
  if(/^B/.test(s)) return mixc(base,'#b98552',.45);
  if(/^Cr/.test(s)) return mixc(base,'#b8ab98',.72);
  if(/^C/.test(s)) return mixc(base,'#d6c3a0',.6);
  return mixc(base,'#c9a878',.4);
}
const lumi=h=>{const [r,g,b]=h2r(h);return (0.299*r+0.587*g+0.114*b)/255;};
function profSVG(c,o){
  o=Object.assign({W:150,H:300,max:150,ax:true,lab:true,top:12},o||{});
  const {L,unk}=hzList(c), ax=o.ax?30:3, lw=o.lab?78:3, x0=ax, w=o.W-ax-lw, ph=o.H-o.top-8;
  const y=d=>o.top+ph*Math.min(d,o.max)/o.max;
  let s=`<svg viewBox="0 0 ${o.W} ${o.H}" role="img" aria-label="Perfil de ${esc(c.c)}: ${esc(c.hz.join(', '))}">`;
  if(o.ax){for(let d=0;d<=o.max;d+=25){s+=`<line x1="${ax-4}" x2="${ax}" y1="${y(d)}" y2="${y(d)}" style="stroke:var(--ink-3)"/><text x="${ax-6}" y="${y(d)+3.5}" text-anchor="end" font-size="9" font-family="IBM Plex Mono,monospace" style="fill:var(--ink-3)">${d}</text>`;}
    s+=`<text x="2" y="${o.top-3}" font-size="8.5" font-family="IBM Plex Mono,monospace" style="fill:var(--ink-3)">cm</text>`;}
  let firstMin=true;
  L.forEach((h,i)=>{const isO=/^O/.test(h.n); const f=hzCol(c,h.n,!isO&&firstMin); if(!isO) firstMin=false;
    const y1=y(h.t), y2=y(h.b);
    s+=`<rect x="${x0}" y="${y1}" width="${w}" height="${Math.max(1,y2-y1)}" fill="${f}"/>`;
    if(/^Cr/.test(h.n.replace(/\d/g,''))) s+=`<rect x="${x0}" y="${y1}" width="${w}" height="${Math.max(1,y2-y1)}" fill="url(#sueCr)"/>`;
    if(i) s+=`<line x1="${x0}" x2="${x0+w}" y1="${y1}" y2="${y1}" stroke="rgba(20,14,8,.55)" stroke-width="1" ${h.est?'stroke-dasharray="3 3"':''}/>`;
    if(o.lab){const ty=(y1+y2)/2+3.5; if(y2-y1>=9) s+=`<text x="${x0+w+6}" y="${ty}" font-size="10" font-family="IBM Plex Mono,monospace" font-weight="600" style="fill:var(--ink)">${esc(h.n)}</text><text x="${x0+w+34}" y="${ty}" font-size="9.5" font-family="IBM Plex Mono,monospace" style="fill:var(--ink-3)">${h.t}–${h.b}${h.est?'*':''}</text>`;}
    else if(y2-y1>=11&&w>=26) s+=`<text x="${x0+w/2}" y="${(y1+y2)/2+3.5}" text-anchor="middle" font-size="8.5" font-family="IBM Plex Mono,monospace" font-weight="600" fill="${lumi(f)<.45?'#f4ede3':'#2a1d12'}">${esc(h.n)}</text>`;
  });
  const last=L.length?L[L.length-1].b:0;
  if(c.ccm){const yc=y(c.ccm), yb=y(o.max);
    if(yb>yc+1) s+=`<rect x="${x0}" y="${yc}" width="${w}" height="${yb-yc}" fill="url(#sueRock)"/>`;
    const lit=/Lítico/.test(c.con), den=/Dénsico/.test(c.con);
    if(den) s+=`<line x1="${x0-3}" x2="${x0+w+3}" y1="${yc}" y2="${yc}" style="stroke:var(--ink)" stroke-width="2" stroke-dasharray="5 3"/>`;
    else if(lit) s+=`<line x1="${x0-3}" x2="${x0+w+3}" y1="${yc}" y2="${yc}" style="stroke:var(--ink)" stroke-width="2.6"/>`;
    else {let p=`M${x0-3} ${yc}`; for(let x=x0-3;x<x0+w+3;x+=6) p+=` L${x+3} ${yc-2.5} L${x+6} ${yc}`; s+=`<path d="${p}" fill="none" style="stroke:var(--ink)" stroke-width="1.6"/>`;}
    if(o.lab&&yb-yc>12) s+=`<text x="${x0+w+6}" y="${Math.min(yc+13,o.H-4)}" font-size="9" font-family="IBM Plex Mono,monospace" style="fill:var(--ink-2)">${esc(c.con.replace(' cm',''))}</text>`;
  } else if(last<o.max&&!unk){ s+=`<line x1="${x0}" x2="${x0+w}" y1="${y(last)}" y2="${y(last)}" style="stroke:var(--ink-3)" stroke-dasharray="2 3"/>`;
    if(o.lab) s+=`<text x="${x0+w+6}" y="${Math.min(y(last)+13,o.H-4)}" font-size="9" font-family="IBM Plex Mono,monospace" style="fill:var(--ink-3)">sin contacto</text>`;}
  s+=`<rect x="${x0}" y="${o.top}" width="${w}" height="${Math.max(1,y(last)-o.top)}" fill="none" style="stroke:var(--ink-2)" stroke-width="1"/>`;
  if(o.ax&&c.pcm){const yp=y(c.pcm); s+=`<path d="M${x0-9} ${o.top} L${x0-9} ${yp} M${x0-12} ${yp} L${x0-6} ${yp}" style="stroke:var(--brick)" stroke-width="1.6" fill="none"/>`;}
  return s+'</svg>';
}
const kvS=(pairs)=>'<dl class="kv">'+pairs.map(([k,v,w])=>w?`<dt class="w">${k}</dt><dd class="w">${v}</dd>`:`<dt>${k}</dt><dd>${v}</dd>`).join('')+'</dl>';
const sw=h=>`<span class="sue-sw" style="background:${h}"></span>`;
const pill=(t,c)=>`<span class="pill ${c}">${esc(t)}</span>`;

/* ---------- estado y filtros ---------- */
let SF={d:'',t:''}, SEL=null;
const filt=()=>CV.filter(c=>(!SF.d||c.dist===SF.d)&&(!SF.t||c.t===SF.t));

function sueTab(){
  const root=$('p-sue'); if(!root) return;
  const ds=[...new Set(CV.map(c=>c.dist))].sort((a,b)=>a.localeCompare(b,'es'));
  $('sueFd').innerHTML='<option value="">Toda la provincia</option>'+ds.map(d=>`<option>${esc(d)}</option>`).join('');
  $('sueFt').innerHTML='<option value="">Todas</option><option value="B">Bloque</option><option value="R">Referencia</option><option value="S">Lote SUS</option>';
  $('sueFd').addEventListener('change',()=>{SF.d=$('sueFd').value;sueDraw();});
  $('sueFt').addEventListener('change',()=>{SF.t=$('sueFt').value;sueDraw();});
  $('sueMapVar').addEventListener('change',sueMap);
  SEL=(CV.find(c=>c.t==='B')||CV[0]||{}).c;
  sueProv(); sueInf(); sueQC();
  sueDraw();
}
function sueDraw(){ sueKpi(); sueMap(); suePick(); sueGal(); suePairs(); sueTax(); sueApt(); sueBlk();
  const n=filt().length; $('sueCnt').textContent=`${n} de ${CV.length} calicatas`; }

/* ---------- 7.1 KPIs ---------- */
function sueKpi(){
  const F=filt(), dsel=SF.d, BB_=B.filter(b=>!dsel||b.d===dsel), ha=sum(BB_,b=>b.ha);
  const own=BB_.filter(b=>b.sAll.length), ownHa=sum(own,b=>b.ha);
  const ana=BB_.filter(b=>{const bl=SU.BL[b.b];return bl&&bl.us.some(u=>SU.U[u]);});
  const fb=F.filter(c=>c.t==='B'), sh=F.filter(c=>c.pcm!=null&&c.pcm<50);
  const nInf=IV.filter(s=>!dsel||siteDist(s)===dsel);
  const chk=KV.filter(k=>F.some(c=>c.c===k.cal));
  const nCal=CV.filter(c=>!dsel||c.dist===dsel).length;
  $('sueKpi').innerHTML=[
    ['Calicatas descritas',n0(F.length),`${F.filter(c=>c.t==='B').length} de bloque · ${F.filter(c=>c.t==='R').length} de referencia · ${F.filter(c=>c.t==='S').length} en Lote SUS`,'hl'],
    ['Bloques con calicata propia',`${own.length} <small style="display:inline">de ${BB_.length}</small>`,`${n2(ownHa)} ha · ${n1(100*ownHa/(ha||1))} % de la superficie`,''],
    ['Bloques con unidad de suelo muestreada',`${ana.length} <small style="display:inline">de ${BB_.length}</small>`,'caracterización por analogía de unidad (plan GeoSIG)',''],
    ['Intensidad de muestreo',nCal?n0(ha/nCal):'—',`ha por calicata · mínimo TDR 1 cada 200 ha`,''],
    ['Profundidad efectiva',fb.length?n0(median(fb.map(c=>c.pcm)))+' cm':'—',`mediana de ${fb.length} calicatas de bloque`,''],
    ['Contacto antes de 50 cm',n0(sh.length),`calicatas con profundidad efectiva < 50 cm`,sh.length?'':''],
    ['Pruebas de infiltración',n0(nInf.length),nInf.length?`${sum(nInf,s=>s.nr)} repeticiones · IB ${nInf.map(s=>fmt(s.ib,2)).join(' / ')} cm/h`:'sin prueba en el filtro',''],
    ['Chequeos y muestras',`${n0(chk.length)} · ${n0(sum(F,c=>c.nm))}`,'barrenaciones/cortes · horizontes muestreados','']
  ].map(([k,v,s,c])=>`<div class="${c}"><dt>${k}</dt><dd>${v}<small>${s}</small></dd></div>`).join('');
}

/* ---------- 7.2 Comparación entre provincias ---------- */
function sueProv(){
  const P=['Morropón','Huancabamba','Ayabaca'];
  const row=p=>{const cs=CAL.filter(c=>c.prov===p), bl=Object.entries(SU.BL).filter(([k,v])=>v.p===p), ha=bl.reduce((s,[k,v])=>s+v.ha,0);
    const withC=bl.filter(([k,v])=>cs.some(c=>c.b===k)).length, cb=cs.filter(c=>c.t==='B');
    const ins=SU.I.filter(s=>siteProv(s)===p);
    const o=t=>cs.filter(c=>ORD(c.tax)===t).length;
    return {p,n:cs.length,nb:bl.length,withC,ha,hc:ha/cs.length,ent:o('Entisol'),inc:o('Inceptisol'),mol:o('Mollisol'),sh:cs.filter(c=>c.pcm!=null&&c.pcm<50).length,md:median(cb.map(c=>c.pcm)),ni:ins.length,ib:median(ins.map(s=>s.ib)),x:cs.filter(c=>cumG(c.cum)==='X').length};};
  const R=P.map(row), T={p:'Ámbito del estudio',n:CAL.length,nb:Object.keys(SU.BL).length,withC:Object.keys(SU.BL).filter(k=>CAL.some(c=>c.b===k)).length,ha:Object.values(SU.BL).reduce((s,v)=>s+v.ha,0)};
  T.hc=T.ha/T.n; ['ent','inc','mol','sh','ni','x'].forEach(k=>T[k]=sum(R,r=>r[k])); T.md=median(CAL.filter(c=>c.t==='B').map(c=>c.pcm)); T.ib=median(SU.I.map(s=>s.ib));
  $('sueProvT').innerHTML='<thead><tr><th class="l">Provincia</th><th>Calicatas</th><th>Bloques con calicata</th><th>ha por calicata</th><th>Entisoles</th><th>Inceptisoles</th><th>Mollisoles</th><th>Prof. efectiva < 50 cm</th><th>Mediana prof. (bloque)</th><th>CUM de campo X</th><th>Pruebas inf.</th><th>IB mediana cm/h</th></tr></thead><tbody>'+
    R.map(r=>`<tr class="${r.p===I.prov?'on':''}"><td class="l"><b>${esc(r.p)}</b>${r.p===I.prov?' <span class="pill p-neu">este volumen</span>':''}</td><td>${r.n}</td><td>${r.withC} de ${r.nb}</td><td>${n0(r.hc)}</td><td>${r.ent}</td><td>${r.inc}</td><td>${r.mol}</td><td>${r.sh} · ${n0(100*r.sh/r.n)} %</td><td>${r.md!=null?n0(r.md)+' cm':'—'}</td><td>${r.x} · ${n0(100*r.x/r.n)} %</td><td>${r.ni}</td><td>${fmt(r.ib,2)}</td></tr>`).join('')+
    `</tbody><tfoot><tr><td class="l">${T.p}</td><td>${T.n}</td><td>${T.withC} de ${T.nb}</td><td>${n0(T.hc)}</td><td>${T.ent}</td><td>${T.inc}</td><td>${T.mol}</td><td>${T.sh} · ${n0(100*T.sh/T.n)} %</td><td>${n0(T.md)} cm</td><td>${T.x} · ${n0(100*T.x/T.n)} %</td><td>${T.ni}</td><td>${fmt(T.ib,2)}</td></tr></tfoot>`;
}

/* ---------- 7.3 Mapa de puntos ---------- */
function sueMap(){
  const host=$('sueMap'); if(!host) return;
  const v=$('sueMapVar').value, F=filt(), FS=new Set(F.map(c=>c.c));
  const pts=[...B.filter(b=>b.e&&b.n).map(b=>[b.e,b.n]),...CV.map(c=>[c.e,c.n]),...KV.map(k=>[k.e,k.n])];
  const W=900,H=560,P=40;
  const x0=Math.min(...pts.map(p=>p[0])),x1=Math.max(...pts.map(p=>p[0])),y0=Math.min(...pts.map(p=>p[1])),y1=Math.max(...pts.map(p=>p[1]));
  const cx=(x0+x1)/2,cy=(y0+y1)/2, s=Math.min((W-2*P)/((x1-x0)||4000),(H-2*P)/((y1-y0)||4000))*0.94;
  const X=e=>W/2+(e-cx)*s, Y=n=>H/2-(n-cy)*s;
  const span=Math.max(x1-x0,y1-y0), step=span>40000?10000:span>15000?5000:span>6000?2000:1000;
  let g='';
  for(let e=Math.ceil((cx-W/2/s)/step)*step;e<=cx+W/2/s;e+=step){const x=X(e); if(x<P||x>W-20) continue; g+=`<line x1="${x}" x2="${x}" y1="14" y2="${H-24}" style="stroke:var(--line)"/><text x="${x}" y="${H-10}" text-anchor="middle" font-size="10" font-family="IBM Plex Mono,monospace" style="fill:var(--ink-3)">${(e/1000).toFixed(0)}</text>`;}
  for(let n=Math.ceil((cy-H/2/s)/step)*step;n<=cy+H/2/s;n+=step){const y=Y(n); if(y<16||y>H-30) continue; g+=`<line x1="${P-6}" x2="${W-8}" y1="${y}" y2="${y}" style="stroke:var(--line)"/><text x="4" y="${y+3.5}" font-size="10" font-family="IBM Plex Mono,monospace" style="fill:var(--ink-3)">${(n/1000).toFixed(0)}</text>`;}
  g+=`<text x="${W-8}" y="${H-10}" text-anchor="end" font-size="10" font-family="IBM Plex Mono,monospace" style="fill:var(--ink-3)">km E</text><text x="4" y="11" font-size="10" font-family="IBM Plex Mono,monospace" style="fill:var(--ink-3)">km N · UTM 17S</text>`;
  const maxha=Math.max(...B.map(b=>b.ha));
  const bl=B.filter(b=>b.e&&b.n).map(b=>{const on=!SF.d||b.d===SF.d, has=b.sAll.length;return `<circle class="sb" data-b="${esc(b.b)}" cx="${X(b.e).toFixed(1)}" cy="${Y(b.n).toFixed(1)}" r="${(4+Math.sqrt(b.ha/maxha)*22).toFixed(1)}" style="fill:${has?'var(--st-sue)':'var(--ink-3)'};fill-opacity:${on?(has?.16:.08):.03};stroke:${has?'var(--st-sue)':'var(--line-2)'};stroke-opacity:${on?.7:.2}" stroke-dasharray="${has?'':'3 3'}"/>`;}).join('');
  const colOf=c=>v==='mun'?c.col:v==='pe'?(peCls(c)==null?'var(--ink-3)':PEC[peCls(c)]):v==='ero'?(ERC[c.eroN]||'var(--ink-3)'):v==='cum'?(CUMC[cumG(c.cum)]||'var(--ink-3)'):GC[c.g];
  const lines=CV.filter(c=>c.t==='B'&&c.par).map(c=>{const r=CV.find(x=>x.c===c.par); if(!r) return ''; const on=FS.has(c.c)||FS.has(r.c);
    return `<line x1="${X(c.e)}" y1="${Y(c.n)}" x2="${X(r.e)}" y2="${Y(r.n)}" style="stroke:var(--ink-2)" stroke-opacity="${on?.55:.12}" stroke-dasharray="4 3"/>`;}).join('');
  const kk=KV.map(k=>{const on=FS.has(k.cal);return `<circle cx="${X(k.e).toFixed(1)}" cy="${Y(k.n).toFixed(1)}" r="2.2" style="fill:var(--ink-2)" fill-opacity="${on?.55:.12}"/>`;}).join('');
  const inf=IV.map(s=>{const on=FS.has(s.cal),x=X(s.e),y=Y(s.n)-15;return `<path class="si" data-i="${s.id}" d="M${x} ${y-6} L${x+6} ${y+5} L${x-6} ${y+5} Z" style="fill:var(--st-hid);stroke:var(--surface)" stroke-width="1.2" fill-opacity="${on?1:.25}"/>`;}).join('');
  const cm=[...CV].sort((a,b)=>(FS.has(a.c)?1:0)-(FS.has(b.c)?1:0)).map(c=>{const on=FS.has(c.c),x=X(c.e),y=Y(c.n),f=colOf(c),sel=c.c===SEL;
    const st=`fill:${f};stroke:${sel?'var(--ink)':'var(--surface)'};opacity:${on?1:.18}`, sw_=sel?2.6:1.4;
    const sh=c.t==='B'?`<circle cx="${x}" cy="${y}" r="7.5" style="${st}" stroke-width="${sw_}"/>`:c.t==='R'?`<path d="M${x} ${y-8.5} L${x+8.5} ${y} L${x} ${y+8.5} L${x-8.5} ${y} Z" style="${st}" stroke-width="${sw_}"/>`:`<rect x="${x-6.5}" y="${y-6.5}" width="13" height="13" style="${st}" stroke-width="${sw_}"/>`;
    return `<g class="sc" data-c="${c.c}" tabindex="0" role="button" aria-label="${esc(c.c)} ${esc(TIPO[c.t])} bloque ${esc(c.b)}">${sh}${on&&(CV.length<=12||sel)?`<text x="${x+11}" y="${y+4}" font-size="10.5" font-family="IBM Plex Mono,monospace" font-weight="600" style="fill:var(--ink)">${c.c.slice(4)}</text>`:''}</g>`;}).join('');
  const sb=step*s;
  host.innerHTML=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Calicatas, chequeos y pruebas de infiltración en UTM WGS 84 Zona 17S">${g}${bl}${lines}${kk}${inf}${cm}<g transform="translate(${P},${H-44})"><rect width="${sb}" height="4" style="fill:var(--ink-2)"/><text y="-5" font-size="10" font-family="IBM Plex Mono,monospace" style="fill:var(--ink-2)">${step/1000} km</text></g></svg>`;
  host.querySelectorAll('g.sc').forEach(el=>{const c=CV.find(x=>x.c===el.dataset.c);
    el.addEventListener('mousemove',e=>showTip(e,`<b>${esc(c.c)} · ${TIPO[c.t]}</b>${tr('Bloque',esc(c.b)+' · '+esc(c.dist))}${tr('Lugar',esc(c.lug))}${tr('Unidad de suelo',esc(c.u))}${tr('Prof. efectiva',c.pcm!=null?c.pcm+' cm':'—')}${tr('Contacto',esc(c.con))}${tr('Clasificación',esc(c.tax.replace('Entisol / ','').replace('Inceptisol / ','').replace('Mollisol / ','')))}${tr('Erosión',esc(c.ero))}${tr('CUM de campo',esc(c.cum))}${tr('UTM 17S',n0(c.e)+' E · '+n0(c.n)+' N · '+n0(c.z)+' m')}`));
    el.addEventListener('mouseleave',hideTip);
    const go=()=>{SEL=c.c; hideTip(); sueMap(); suePick(); $('suePairH').scrollIntoView({behavior:'smooth',block:'start'});};
    el.addEventListener('click',go); el.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();go();}});});
  host.querySelectorAll('circle.sb').forEach(el=>{const b=B.find(x=>x.b===el.dataset.b);
    el.addEventListener('mousemove',e=>showTip(e,`<b>Bloque ${esc(b.b)}</b>${tr('Distrito',esc(b.d))}${tr('Superficie',n2(b.ha)+' ha')}${tr('Calicatas',b.sAll.length?b.sAll.join(', '):'sin calicata propia')}`)); el.addEventListener('mouseleave',hideTip);
    el.style.cursor='pointer'; el.addEventListener('click',()=>goFicha(b.b));});
  host.querySelectorAll('path.si').forEach(el=>{const s_=SU.I.find(x=>x.id===el.dataset.i);
    el.addEventListener('mousemove',e=>showTip(e,`<b>${esc(s_.id)} · ${esc(s_.sit)}</b>${tr('Calicata',esc(s_.cal))}${tr('Repeticiones',s_.nr)}${tr('IB representativa',fmt(s_.ib,2)+' cm/h')}${tr('Clase',esc(s_.cl))}`)); el.addEventListener('mouseleave',hideTip);});
  const leg={g:[1,2,3].map(k=>[GC[k],GRPs[k]]),pe:PEN.map((t,i)=>[PEC[i],t]),ero:[1,2,3,4].map(k=>[ERC[k],ERN[k]]),cum:['X','F','P','C','A'].map(k=>[CUMC[k],CUMN[k]]),mun:[]}[v];
  $('sueMapLeg').innerHTML=(v==='mun'?'<span>Color Munsell en húmedo del horizonte superficial (Cuadro 13), convertido a sRGB</span>':leg.map(([c,t])=>`<span><i style="background:${c}"></i>${esc(t)}</span>`).join(''))+
    '<span class="sue-shp"><svg width="14" height="14" viewBox="0 0 14 14"><circle cx="7" cy="7" r="5.5" style="fill:var(--ink-3)"/></svg>Bloque</span><span class="sue-shp"><svg width="14" height="14" viewBox="0 0 14 14"><path d="M7 1 L13 7 L7 13 L1 7 Z" style="fill:var(--ink-3)"/></svg>Referencia</span><span class="sue-shp"><svg width="14" height="14" viewBox="0 0 14 14"><rect x="2" y="2" width="10" height="10" style="fill:var(--ink-3)"/></svg>Lote SUS</span><span class="sue-shp"><svg width="14" height="14" viewBox="0 0 14 14"><path d="M7 1 L13 12 L1 12 Z" style="fill:var(--st-hid)"/></svg>Prueba de infiltración</span><span class="sue-shp"><svg width="14" height="14" viewBox="0 0 14 14"><circle cx="7" cy="7" r="2.5" style="fill:var(--ink-2)"/></svg>Chequeo</span>';
}

/* ---------- 7.4 Visor de perfiles y pares ---------- */
function card(c){
  const inf=infOf(c), iu=!inf&&infUnit(c), a=apt(c);
  return `<div class="sue-side"><div class="sue-sh"><span class="pill ${c.t==='B'?'p-cri':c.t==='R'?'p-ok':'p-neu'}">${TIPO[c.t]}</span><h3 class="mono">${esc(c.c)}</h3><span class="muted sm">${esc(c.lug)} · bloque ${esc(c.b)} · ${esc(c.dist)}</span></div>
   <div class="sue-pc"><div class="sue-prof">${profSVG(c,{W:170,H:330})}</div><div>${kvS([
    ['Horizonte A',`${sw(c.col)} ${esc(c.mun)} · ${esc(c.tex)}`],['Estructura',esc(c.est)],['Raíces · fragmentos',esc(c.rai)+' · '+esc(c.frag)],
    ['Prof. efectiva',c.pcm!=null?`<b>${c.pcm} cm</b>`:'—'],['Contacto',esc(c.con)],['Epipedón · subsup.',esc(c.epi)+' · '+esc(c.sub)],
    ['Soil Taxonomy',esc(c.tax),true],['WRB · grupo',esc(c.wrb)+' · '+esc(GRPs[c.g]),true],
    ['Pendiente',esc(c.pend)+' · '+esc(c.pc)],['Posición · uso',esc(c.pos)+' · '+esc(c.uso)],['Erosión · drenaje',esc(c.ero)+' · '+esc(c.dren)],
    ['CUM campo · gabinete',esc(c.cum)+' · '+esc(c.cumu)],
    ['Infiltración',inf?`<b>${fmt(inf.ib,2)} cm/h</b> · ${esc(inf.cl)} (${esc(inf.id)})`:iu?`${fmt(iu.ib,2)} cm/h por analogía (${esc(iu.id)})`:'sin prueba en la unidad'],
    ...(c.t!=='R'?[['Zanjas de infiltración',pill(a.z,a.zc)]]:[]),
    ['Vegetación',esc(c.veg),true],['Unidad de suelo',esc(c.u)+' · '+esc(c.zh)+' · '+esc(c.eco),true],
    ['Muestras',esc(c.mu),true],['UTM 17S · desplaz.',`${n0(c.e)} E · ${n0(c.n)} N · ${n0(c.z)} m · ${n0(c.dz)} m del plan`]
   ])}${c.mot?`<p class="sm muted" style="margin:8px 0 0">Reubicación: ${esc(c.mot)}</p>`:''}${c.obs?`<p class="sm muted" style="margin:6px 0 0">${esc(c.obs)}</p>`:''}</div></div></div>`;
}
function suePick(){
  const sel=$('suePick'); const F=CV;
  const opts=['B','S','R'].map(t=>{const a=F.filter(c=>c.t===t); if(!a.length) return ''; return `<optgroup label="${TIPO[t]}">`+a.map(c=>`<option value="${c.c}">${c.c} · bloque ${esc(c.b)} · ${esc(c.dist)}${c.par?' · par '+c.par:''}</option>`).join('')+'</optgroup>';}).join('');
  if(sel.dataset.v!==String(CV.length)){sel.innerHTML=opts; sel.dataset.v=String(CV.length); sel.addEventListener('change',()=>{SEL=sel.value;sueMap();suePick();});}
  const c=CV.find(x=>x.c===SEL)||CV[0]; if(!c){$('suePair').innerHTML='<p class="muted">Sin calicatas en este volumen.</p>';return;}
  sel.value=c.c;
  const par=c.par?CAL.find(x=>x.c===c.par):null;
  const bl=c.t==='R'?par:c, rf=c.t==='R'?c:par;
  let cmp='';
  if(bl&&rf){
    const dP=(rf.pcm||0)-(bl.pcm||0);
    const dv=parseFloat((bl.mun.split('/')[0].split(' ')[1]))-parseFloat((rf.mun.split('/')[0].split(' ')[1]));
    const rk={Pocas:1,Comunes:2,Muchas:3};
    cmp=`<div class="sue-cmp"><div><span class="k">Profundidad efectiva</span><b>${dP>0?'+':''}${dP} cm</b><span>${dP>0?'la referencia es más profunda':dP<0?'el bloque es más profundo (posición en la ladera)':'igual profundidad'}</span></div>
      <div><span class="k">Color del horizonte A</span><b>${sw(bl.col)} → ${sw(rf.col)}</b><span>valor Munsell ${esc(bl.mun)} frente a ${esc(rf.mun)}${dv>0?' · la referencia es más oscura':dv<0?' · el bloque es más oscuro':''}</span></div>
      <div><span class="k">Horizontes de diagnóstico</span><b>${esc(bl.epi.slice(0,3))}+${esc(bl.sub==='No hay'?'No':bl.sub.slice(0,3))} / ${esc(rf.epi.slice(0,3))}+${esc(rf.sub==='No hay'?'No':rf.sub.slice(0,3))}</b><span>${(rf.sub!=='No hay'&&bl.sub==='No hay')||(rf.epi==='Mólico'&&bl.epi!=='Mólico')?'la referencia conserva un horizonte que el bloque perdió o no desarrolló':'mismo nivel de desarrollo en campo'}</span></div>
      <div><span class="k">Raíces · erosión</span><b>${esc(bl.rai)} / ${esc(rf.rai)}</b><span>erosión ${esc(bl.ero.toLowerCase())} en el bloque, ${esc(rf.ero.toLowerCase())} en la referencia${(rk[rf.rai]||0)>(rk[bl.rai]||0)?' · más raíces en la referencia':''}</span></div></div>`;
  }
  $('suePair').innerHTML=(bl&&rf?`<div class="sue-two">${card(bl)}${card(rf)}</div>${cmp}`:`<div class="sue-two one">${card(c)}</div>`)+
    `<p class="foot">Perfil a escala (0–150 cm). El color del horizonte superficial es el Munsell en húmedo descrito en campo; los demás horizontes se dibujan en tonos esquemáticos derivados de él. Límites con asterisco: horizonte sin muestra, profundidad repartida entre la última muestra y la profundidad del perfil. Trazo rojo a la izquierda: profundidad efectiva. Bajo el contacto: roca (lítico, línea continua gruesa), roca meteorizada (paralítico, línea quebrada) o material denso (dénsico, línea discontinua).</p>`;
  $('sueGal').querySelectorAll('button').forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.c===c.c)));
}
function sueGal(){
  const F=[...filt()].sort((a,b)=>(a.pcm??999)-(b.pcm??999)||a.c.localeCompare(b.c));
  $('sueGal').innerHTML=F.map(c=>`<button type="button" data-c="${c.c}" aria-pressed="${c.c===SEL}" title="${esc(c.c)} · ${esc(c.hz.join('-'))} · ${c.pcm!=null?c.pcm+' cm':'s/d'}">${profSVG(c,{W:46,H:170,ax:false,lab:false,top:4})}<span>${c.c.slice(4)}</span><i class="t-${c.t}"></i></button>`).join('')||'<p class="muted">Sin calicatas en el filtro.</p>';
  $('sueGal').querySelectorAll('button').forEach(b=>b.addEventListener('click',()=>{SEL=b.dataset.c;sueMap();suePick();$('suePairH').scrollIntoView({behavior:'smooth',block:'start'});}));
}

/* ---------- 7.5 Pares bloque–referencia ---------- */
function suePairs(){
  const host=$('suePairs'); const F=filt();
  const P=CV.filter(c=>c.t==='B'&&c.par).map(c=>[c,CAL.find(x=>x.c===c.par)]).filter(([b,r])=>r&&(F.some(x=>x.c===b.c)||F.some(x=>x.c===r.c)));
  if(!P.length){host.innerHTML='<p class="muted" style="padding:10px">Sin pares en el filtro.</p>';$('suePairsTxt').textContent='';return;}
  P.sort((a,b)=>((b[1].pcm||0)-(b[0].pcm||0))-((a[1].pcm||0)-(a[0].pcm||0)));
  const W=880,rowH=30,T=30,L=170,R=210,H=T+P.length*rowH+30,mx=150,X=v=>L+(W-L-R)*Math.min(v,mx)/mx;
  let s=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Profundidad efectiva de la calicata de bloque frente a su referencia">`;
  for(let v=0;v<=mx;v+=25) s+=`<line x1="${X(v)}" x2="${X(v)}" y1="${T-8}" y2="${H-26}" style="stroke:var(--line)"/><text x="${X(v)}" y="${H-10}" text-anchor="middle" font-size="10.5" font-family="IBM Plex Mono,monospace" style="fill:var(--ink-3)">${v}</text>`;
  s+=`<line x1="${X(50)}" x2="${X(50)}" y1="${T-8}" y2="${H-26}" style="stroke:var(--brick)" stroke-dasharray="4 3"/><text x="${X(50)+4}" y="${T-12}" font-size="10" style="fill:var(--brick)">50 cm · límite para zanjas</text>`;
  s+=`<text x="${W-R+14}" y="${T-12}" font-size="10" font-family="IBM Plex Mono,monospace" style="fill:var(--ink-3)">DIAGNÓSTICO BLOQUE / REF.</text>`;
  P.forEach(([b,r],i)=>{const y=T+i*rowH+rowH/2, xb=X(b.pcm||0), xr=X(r.pcm||0), better=(r.pcm||0)>(b.pcm||0)||(r.sub!=='No hay'&&b.sub==='No hay')||(r.epi==='Mólico'&&b.epi!=='Mólico');
    s+=`<g class="pr" data-c="${b.c}"><rect x="0" y="${y-rowH/2}" width="${W}" height="${rowH}" style="fill:${i%2?'var(--surface-2)':'transparent'}"/>`+
      `<text x="8" y="${y+4}" font-size="11.5" font-family="IBM Plex Mono,monospace" font-weight="600" style="fill:var(--anin)">${b.c.slice(4)}/${r.c.slice(4)}</text><text x="62" y="${y+4}" font-size="11" style="fill:var(--ink-2)">bl. ${esc(b.b)}</text>`+
      `<line x1="${xb}" x2="${xr}" y1="${y}" y2="${y}" style="stroke:var(--ink-3)" stroke-width="2"/>`+
      `<circle cx="${xb}" cy="${y}" r="6" style="fill:var(--brick)"/><path d="M${xr} ${y-7} L${xr+7} ${y} L${xr} ${y+7} L${xr-7} ${y} Z" style="fill:var(--moss-2)"/>`+
      `<text x="${W-R+14}" y="${y+4}" font-size="11" font-family="IBM Plex Mono,monospace" style="fill:${better?'var(--moss)':'var(--ink-2)'}">${esc(b.epi.slice(0,3))}+${b.sub==='No hay'?'No':esc(b.sub.slice(0,3))} / ${esc(r.epi.slice(0,3))}+${r.sub==='No hay'?'No':esc(r.sub.slice(0,3))}${better?' ◂':''}</text></g>`;});
  s+=`<text x="8" y="${H-10}" font-size="10.5" style="fill:var(--ink-3)">Profundidad efectiva (cm) →</text></svg>`;
  host.innerHTML=s;
  host.querySelectorAll('g.pr').forEach(g=>{const [b,r]=P.find(p=>p[0].c===g.dataset.c);
    g.addEventListener('mousemove',e=>showTip(e,`<b>Par ${esc(b.c)} / ${esc(r.c)}</b>${tr('Unidad de suelo',esc(b.u))}${tr('Prof. bloque',b.pcm!=null?b.pcm+' cm':'—')}${tr('Prof. referencia',r.pcm!=null?r.pcm+' cm':'—')}${tr('Uso bloque / ref.',esc(b.uso)+' / '+esc(r.uso))}${tr('Horizonte A',esc(b.mun)+' / '+esc(r.mun))}`));
    g.addEventListener('mouseleave',hideTip); g.style.cursor='pointer'; g.addEventListener('click',()=>{SEL=b.c;sueMap();suePick();$('suePairH').scrollIntoView({behavior:'smooth',block:'start'});});});
  const deeper=P.filter(([b,r])=>(r.pcm||0)>(b.pcm||0)).length, dev=P.filter(([b,r])=>(r.sub!=='No hay'&&b.sub==='No hay')||(r.epi==='Mólico'&&b.epi!=='Mólico')).length, dark=P.filter(([b,r])=>parseFloat(r.mun.split(' ')[1])<parseFloat(b.mun.split(' ')[1])).length;
  $('suePairsTxt').innerHTML=`En <b>${deeper} de ${P.length}</b> pares la referencia conservada es más profunda que el bloque degradado; en <b>${dev}</b> la referencia tiene un horizonte de diagnóstico (cámbico o mólico) que el bloque no tiene, y en <b>${dark}</b> su horizonte A es más oscuro (valor Munsell menor, indicio de más materia orgánica). Donde el bloque resulta más profundo, la diferencia responde a la posición en la ladera. El contraste se cuantificará con materia orgánica, densidad aparente y saturación de bases del laboratorio.`;
}

/* ---------- 7.6 Infiltración ---------- */
let INFSEL=null;
function sueInf(){
  const host=$('sueInf'); const S=[...SU.I].sort((a,b)=>b.ib-a.ib), mine=new Set(IV.map(s=>s.id));
  INFSEL=INFSEL||(IV[0]||S[0]).id;
  const W=880,rowH=30,T=34,L=300,R=70,H=T+S.length*rowH+34, lx=v=>Math.log10(v), a=lx(0.05), b=lx(150), X=v=>L+(W-L-R)*(lx(Math.max(v,.05))-a)/(b-a);
  let s=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Velocidad de infiltración básica por sitio, escala logarítmica, con las clases del USDA">`;
  ICLS.forEach(([lo,hi,t],i)=>{const x1=X(Math.max(lo,.05)),x2=X(Math.min(hi,150)); s+=`<rect x="${x1}" y="${T-6}" width="${x2-x1}" height="${H-T-28}" style="fill:${i%2?'var(--surface-2)':'var(--surface)'}"/><text x="${(x1+x2)/2}" y="${T-12}" text-anchor="middle" font-size="9.5" style="fill:var(--ink-3)">${t}</text>`;});
  [.1,.5,2,6.3,12.7,25.4,100].forEach(v=>s+=`<line x1="${X(v)}" x2="${X(v)}" y1="${T-6}" y2="${H-22}" style="stroke:var(--line-2)"/><text x="${X(v)}" y="${H-8}" text-anchor="middle" font-size="10" font-family="IBM Plex Mono,monospace" style="fill:var(--ink-3)">${v}</text>`);
  S.forEach((si,i)=>{const y=T+i*rowH+rowH/2, on=mine.has(si.id), sel=si.id===INFSEL, rp=SU.R.filter(r=>r.i===si.id);
    s+=`<g class="ir" data-i="${si.id}" tabindex="0" role="button" aria-label="${esc(si.id)} ${esc(si.sit)}">${sel?`<rect x="0" y="${y-rowH/2}" width="${W}" height="${rowH}" style="fill:var(--anin-wash)"/>`:''}<text x="6" y="${y+4}" font-size="11" font-family="IBM Plex Mono,monospace" font-weight="600" style="fill:${on?'var(--anin)':'var(--ink-3)'}">${esc(si.id)}</text><text x="66" y="${y+4}" font-size="11.5" style="fill:${on?'var(--ink)':'var(--ink-3)'}">${esc((si.sit+' · '+siteDist(si)).slice(0,36))}</text>`+
      `<rect x="${L}" y="${y-6}" width="${Math.max(1,X(si.ib)-L)}" height="12" rx="2" style="fill:${on?'var(--st-hid)':'var(--line-2)'}"/>`+
      rp.map(r=>`<circle cx="${X(r.ib)}" cy="${y}" r="3" style="fill:var(--surface);stroke:${on?'var(--ink)':'var(--ink-3)'}" stroke-width="1.2"/>`).join('')+
      `<text x="${W-R+8}" y="${y+4}" font-size="11.5" font-family="IBM Plex Mono,monospace" font-weight="600" style="fill:${on?'var(--ink)':'var(--ink-3)'}">${fmt(si.ib,2)}</text></g>`;});
  s+=`<text x="6" y="${H-8}" font-size="10" style="fill:var(--ink-3)">Velocidad básica (cm/h, escala logarítmica) →</text></svg>`;
  host.innerHTML=s;
  host.querySelectorAll('g.ir').forEach(g=>{const si=SU.I.find(x=>x.id===g.dataset.i);
    g.addEventListener('mousemove',e=>showTip(e,`<b>${esc(si.id)} · ${esc(si.sit)}</b>${tr('Provincia · distrito',esc(siteProv(si))+' · '+esc(siteDist(si)))}${tr('Calicata · unidad',esc(si.cal)+' · '+esc(si.u))}${tr('Repeticiones medidas',si.nr)}${tr('IB, rango',esc(si.rib)+' cm/h')}${tr('IB representativa',fmt(si.ib,2)+' cm/h')}${tr('Clase USDA',esc(si.cl))}`));
    g.addEventListener('mouseleave',hideTip); g.style.cursor='pointer';
    const go=()=>{INFSEL=si.id;sueInf();}; g.addEventListener('click',go); g.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();go();}});});
  sueCurve();
  // tabla del volumen
  const des=v=>v<2?'Zanjas con mayor capacidad de almacenamiento y menor espaciamiento; dimensionar con la IB del sitio':v<6.3?'Diseño con la IB del sitio; espaciamiento estándar de la guía':'Zanjas no prioritarias; conservar cobertura y horizonte orgánico';
  $('sueInfT').innerHTML='<thead><tr><th class="l">Prueba</th><th class="l">Sitio</th><th>Calicata</th><th class="l">Unidad de suelo</th><th>Rep.</th><th>IB rango (cm/h)</th><th>IB repres. (cm/h)</th><th>Clase</th><th class="l">Implicancia de diseño</th><th>Calicatas por analogía</th></tr></thead><tbody>'+
    (IV.length?IV.map(si=>{const an=CV.filter(c=>c.u===si.u&&c.c!==si.cal).map(c=>c.c);return `<tr><td class="cod">${esc(si.id)}</td><td class="l">${esc(si.sit)}</td><td class="mono">${esc(si.cal)}</td><td class="l mono sm">${esc(si.u)}</td><td>${si.nr}</td><td>${esc(si.rib)}</td><td><b>${fmt(si.ib,2)}</b></td><td>${esc(si.cl)}</td><td class="wrap">${des(si.ib)}</td><td class="mono sm">${an.join(', ')||'—'}</td></tr>`;}).join(''):`<tr><td colspan="10" class="l muted">Sin prueba de infiltración en los bloques de este volumen.</td></tr>`)+'</tbody>';
}
function sueCurve(){
  const si=SU.I.find(x=>x.id===INFSEL), rp=SU.R.filter(r=>r.i===si.id);
  const W=420,H=230,L=46,R=12,T=14,Bm=34, tm=240;
  const Imax=Math.max(...rp.map(r=>r.k*Math.pow(tm,r.a)))*1.08, vmax=Math.max(...rp.map(r=>60*r.k*r.a*Math.pow(5,r.a-1)))*1.05;
  const X=t=>L+(W-L-R)*t/tm, Y1=v=>T+(H-T-Bm)*(1-v/Imax), Y2=v=>T+(H-T-Bm)*(1-Math.min(v,vmax)/vmax);
  const nice=m=>{const raw=m/4,p=Math.pow(10,Math.floor(Math.log10(raw))),q=raw/p;return (q<1.5?1:q<3?2:q<7?5:10)*p;};
  const axes=(Y,mx,lab)=>{let s='';const st=nice(mx);for(let v=0;v<=mx;v+=st){s+=`<line x1="${L}" x2="${W-R}" y1="${Y(v)}" y2="${Y(v)}" style="stroke:var(--line)"/><text x="${L-6}" y="${Y(v)+3.5}" text-anchor="end" font-size="10" font-family="IBM Plex Mono,monospace" style="fill:var(--ink-3)">${+v.toFixed(2)}</text>`;}
    [0,60,120,180,240].forEach(t=>s+=`<text x="${X(t)}" y="${H-Bm+15}" text-anchor="middle" font-size="10" font-family="IBM Plex Mono,monospace" style="fill:var(--ink-3)">${t}</text>`);
    return s+`<text x="${L}" y="${H-4}" font-size="10" style="fill:var(--ink-3)">${lab}</text>`;};
  const cols=['var(--st-hid)','var(--anin)','var(--moss-2)','var(--ochre)','var(--brick)'];
  const line=(f,Y)=>rp.map((r,i)=>{let p='';for(let t=1;t<=Math.min(tm,r.d);t+=2){p+=(p?' L':'M')+X(t).toFixed(1)+' '+Y(f(r,t)).toFixed(1);}return `<path d="${p}" fill="none" style="stroke:${cols[i%5]}" stroke-width="1.8"/>`;}).join('');
  const s1=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Lámina infiltrada acumulada por repetición">${axes(Y1,Imax,'Lámina acumulada I = k·tᵃ (cm) · tiempo (min)')}${line((r,t)=>r.k*Math.pow(t,r.a),Y1)}</svg>`;
  const s2=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Velocidad de infiltración por repetición">${axes(Y2,vmax,'Velocidad i = 60·k·a·tᵃ⁻¹ (cm/h) · tiempo (min)')}<line x1="${L}" x2="${W-R}" y1="${Y2(si.ib)}" y2="${Y2(si.ib)}" style="stroke:var(--brick)" stroke-dasharray="5 4"/><text x="${W-R}" y="${Y2(si.ib)-5}" text-anchor="end" font-size="10.5" style="fill:var(--brick)">IB repres. ${fmt(si.ib,2)}</text>${line((r,t)=>60*r.k*r.a*Math.pow(t,r.a-1),Y2)}</svg>`;
  $('sueCurve').innerHTML=`<div class="sue-cv-h"><b class="mono">${esc(si.id)}</b> ${esc(si.sitio||si.sit)} · ${esc(si.u)} · ${si.nr} repeticiones medidas</div><div class="two"><figure class="chart">${s1}</figure><figure class="chart">${s2}</figure></div>`+
    `<div class="legend">${rp.map((r,i)=>`<span><i style="background:${cols[i%5]}"></i>${esc(r.r.replace(si.id+'-','').replace('INF-02-',''))}: k ${r.k} · a ${r.a} · R² ${r.r2} · IB ${fmt(r.ib,2)} (${esc(r.cl)})</span>`).join('')}</div>`;
}

/* ---------- 7.7 Clasificación y capacidad de uso ---------- */
function sueTax(){
  const F=filt(), ds=[...new Set(F.map(c=>c.dist))].sort((a,b)=>a.localeCompare(b,'es'));
  const O=[['Entisol','var(--sg1)'],['Inceptisol','var(--sg2)'],['Mollisol','var(--sg3)']];
  const rows=[...ds.map(d=>({l:d,cs:F.filter(c=>c.dist===d),k:'d'})),{l:SF.d||SF.t?'Total del filtro':'Total '+I.prov,cs:F,k:'t'},...['Morropón','Huancabamba','Ayabaca'].filter(p=>p!==I.prov).map(p=>({l:p+' (otro volumen)',cs:CAL.filter(c=>c.prov===p),k:'o'}))];
  const W=880,rowH=28,L=210,R=60,T=10,H=T+rows.length*rowH+8, X=v=>L+(W-L-R)*v;
  let s=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Orden taxonómico tentativo por distrito y provincia">`;
  rows.forEach((r,i)=>{const y=T+i*rowH, n=r.cs.length; let acc=0;
    s+=`<text x="${L-10}" y="${y+rowH/2+4}" text-anchor="end" font-size="12" ${r.k==='t'?'font-weight="700"':''} style="fill:${r.k==='o'?'var(--ink-3)':'var(--ink)'}">${esc(r.l)}</text>`;
    O.forEach(([o,c])=>{const k=r.cs.filter(x=>ORD(x.tax)===o).length; if(!k) return; const w=(W-L-R)*k/n;
      s+=`<rect x="${X(acc/n)}" y="${y+5}" width="${w}" height="${rowH-10}" style="fill:${c};opacity:${r.k==='o'?.45:1}"/>${w>22?`<text x="${X(acc/n)+w/2}" y="${y+rowH/2+4}" text-anchor="middle" font-size="11" font-weight="600" fill="#fff">${k}</text>`:''}`; acc+=k;});
    s+=`<text x="${W-R+8}" y="${y+rowH/2+4}" font-size="11" font-family="IBM Plex Mono,monospace" style="fill:var(--ink-3)">n = ${n}</text>`;});
  $('sueTaxC').innerHTML=s+'</svg>';
  $('sueTaxL').innerHTML=O.map(([o,c])=>`<span><i style="background:${c}"></i>${o}${o==='Entisol'?' (Regosoles, Leptosoles)':o==='Inceptisol'?' (Cambisoles)':' (Phaeozems)'}</span>`).join('');
  // CUM campo × gabinete
  const cr=['X','F','P','C','A'], cc=['X','P','C','A'];
  const cnt=(a,b)=>F.filter(c=>cumG(c.cum)===a&&cumG(c.cumu)===b).length;
  $('sueCum').innerHTML='<thead><tr><th class="l">CUM de campo (pre-clasificación) ↓ · gabinete de la unidad →</th>'+cc.map(k=>`<th>${k}</th>`).join('')+'<th>Total</th></tr></thead><tbody>'+
    cr.map(a=>{const tot=F.filter(c=>cumG(c.cum)===a).length; if(!tot) return ''; return `<tr><td class="l"><i class="sue-dot" style="background:${CUMC[a]}"></i>${CUMN[a]}</td>`+cc.map(b=>{const v=cnt(a,b);return `<td class="${v&&(a===b||(a==='F'&&b==='X'))?'sue-agree':''}">${v||'·'}</td>`;}).join('')+`<td><b>${tot}</b></td></tr>`;}).join('')+'</tbody>';
  const agree=F.filter(c=>{const a=cumG(c.cum),b=cumG(c.cumu);return a===b||(a==='F'&&b==='X');}).length;
  $('sueCumTxt').innerHTML=F.length?`<b>${agree} de ${F.length}</b> calicatas (${n0(100*agree/F.length)} %) coinciden con la CUM de gabinete de su unidad o quedan en forestal con alternativa de protección. La CUM definitiva por unidad de suelo requiere pH, salinidad y fertilidad del laboratorio (Entregable 4).`:'';
  // pendiente × profundidad
  const pcs=['C','D','E','F','G','H'];
  const res=(p,k)=>p==='G'||p==='H'?'X':p==='F'?(k>=2?'F':'X'):p==='E'?(k>=2?'F':'P'):p==='D'?'C':'A';
  $('sueHeat').innerHTML='<thead><tr><th class="l">Profundidad efectiva ↓ · clase de pendiente →</th>'+pcs.map(p=>`<th>${p}<br><span class="sm" style="font-weight:500">${{C:'4–8',D:'8–15',E:'15–25',F:'25–50',G:'50–75',H:'> 75'}[p]} %</span></th>`).join('')+'</tr></thead><tbody>'+
    [3,2,1,0].map(k=>`<tr><td class="l"><i class="sue-dot" style="background:${PEC[k]}"></i>${PEN[k]}</td>`+pcs.map(p=>{const v=F.filter(c=>c.pc===p&&peCls(c)===k).length, g=res(p,k);
      return `<td style="background:color-mix(in srgb, ${CUMC[g]} ${v?Math.min(70,18+v*9):6}%, var(--surface));${v>4?'color:#fff;font-weight:700':''}" title="${esc(CUMN[g])}">${v||''}</td>`;}).join('')+'</tr>').join('')+'</tbody>';
}

/* ---------- 7.8 Aptitud para las MRR-CCC ---------- */
function sueApt(){
  const F=filt().filter(c=>c.t!=='R');
  const rows=F.map(c=>({c,a:apt(c),inf:infOf(c),iu:infUnit(c)}));
  const cnt=k=>rows.filter(r=>r.a.z.startsWith(k)).length;
  $('sueAptK').innerHTML=[['Compatible','p-ok','zanjas de infiltración en 15–50 % con ≥ 50 cm de suelo'],['Restringida','p-adv','contacto lítico, paralítico o dénsico somero'],['No','p-cri','pendiente > 50 %: revegetación, exclusión y control de cárcavas'],['No prioritaria','p-neu','suelos mólicos que ya infiltran rápido'],['Fuera de rango','p-neu','pendiente < 15 %']].map(([k,c,d])=>`<div><dt><span class="pill ${c}">${k}</span></dt><dd>${cnt(k)}<small>${d}</small></dd></div>`).join('');
  $('sueApt').innerHTML='<thead><tr><th class="l">Calicata</th><th>Bloque</th><th class="l">Distrito</th><th>Pendiente</th><th>Prof. efectiva</th><th class="l">Contacto</th><th>Grupo</th><th>IB (cm/h)</th><th class="l">Zanjas de infiltración</th><th>Terrazas 15–40 %</th><th>Cárcavas y exclusión</th></tr></thead><tbody>'+
    rows.map(({c,a,inf,iu})=>`<tr data-cod="${esc(c.b)}"><td class="cod">${esc(c.c)}${c.t==='S'?' <span class="pill p-neu">SUS</span>':''}</td><td class="mono">${esc(c.b)}</td><td class="l">${esc(c.dist)}</td><td>${a.p!=null?n0(a.p)+' %':'—'} · ${esc(c.pc)}</td><td>${c.pcm!=null?c.pcm+' cm':'—'}</td><td class="l">${esc(c.con)}</td><td><i class="sue-dot" style="background:${GC[c.g]}"></i>${c.g}</td><td>${inf?`<b>${fmt(inf.ib,2)}</b>`:iu?`<span class="muted">${fmt(iu.ib,2)}ᵃ</span>`:'—'}</td><td class="l">${pill(a.z,a.zc)}</td><td>${a.t[1]?pill(a.t[0],a.t[1]):'—'}</td><td>${a.k[1]?pill(a.k[0],a.k[1]):'—'}</td></tr>`).join('')+'</tbody>';
  $('sueApt').querySelectorAll('tbody tr').forEach(r=>r.addEventListener('click',()=>goFicha(r.dataset.cod)));
}

/* ---------- 7.9 Cobertura por bloque ---------- */
let BK='cal', BD=-1;
function sueBlk(){
  const rs=B.filter(b=>!SF.d||b.d===SF.d).map(b=>{const bl=SU.BL[b.b]||{us:[]}; const sm=bl.us.filter(u=>SU.U[u]);
    return {b:b.b,d:b.d,ha:b.ha,ti:bl.ti||'—',eco:bl.eco||'—',zv:bl.zv||'—',nu:bl.us.length,ns:sm.length,sm,cal:b.sAll.length,cals:b.sAll,inf:bl.inf,prof:b.prof,ms:b.ms,ipp:b.ipp};});
  const key=r=>BK==='cov'?(r.nu?r.ns/r.nu:0):r[BK];
  rs.sort((a,b)=>{const x=key(a),y=key(b);if(x==null)return 1;if(y==null)return -1;return (typeof x==='string'?x.localeCompare(y,'es',{numeric:true}):x-y)*BD||b.ha-a.ha;});
  const cols=[['b','Bloque'],['d','Distrito'],['ha','ha'],['ti','Tipo (plan de suelos)'],['eco','Ecosistema'],['cov','Unidades muestreadas'],['cal','Calicatas propias'],['prof','Prof. efectiva'],['ipp','Índice PI']];
  const t=$('sueBlk');
  t.innerHTML='<thead><tr>'+cols.map(([k,l])=>`<th class="s${['b','d','ti','eco'].includes(k)?' l':''}" data-k="${k}" ${k===BK?`data-dir="${BD}"`:''}>${esc(l)}<span class="car">${k===BK?(BD<0?'▼':'▲'):'▲'}</span></th>`).join('')+'</tr></thead><tbody>'+
    rs.map(r=>`<tr data-cod="${esc(r.b)}"><td class="cod">${esc(r.b)}</td><td class="l">${esc(r.d)}</td><td>${n2(r.ha)}</td><td class="l sm">${esc(r.ti)}</td><td class="l sm">${esc(r.eco)}</td><td><span class="sue-cov"><span style="width:${r.nu?100*r.ns/r.nu:0}%"></span></span> ${r.ns} de ${r.nu}${r.sm.length?`<br><span class="sm muted mono">${r.sm.map(u=>SU.U[u].c.slice(0,2).join('/')).join(' · ')}</span>`:''}</td><td class="mono">${r.cals.join(', ')||'<span class="muted">—</span>'}</td><td>${r.prof!=null?r.prof+' cm':'—'}</td><td class="mono">${fmt(r.ipp,3)}</td></tr>`).join('')+'</tbody>'+
    `<tfoot><tr><td class="l" colspan="2">${rs.length} bloques</td><td>${n2(sum(rs,r=>r.ha))}</td><td colspan="2"></td><td>${rs.filter(r=>r.ns).length} con alguna unidad muestreada</td><td>${rs.filter(r=>r.cal).length} con calicata</td><td colspan="2"></td></tr></tfoot>`;
  t.querySelectorAll('th.s').forEach(th=>th.addEventListener('click',()=>{const k=th.dataset.k;if(BK===k)BD=-BD;else{BK=k;BD=['b','d','ti','eco'].includes(k)?1:-1;}sueBlk();}));
  t.querySelectorAll('tbody tr').forEach(r=>r.addEventListener('click',()=>goFicha(r.dataset.cod)));
}

/* ---------- 7.10 Consistencia de la fuente ---------- */
function sueQC(){
  const rel=CV.filter(c=>c.dz>300), ns=CV.filter(c=>!c.sam.length), hcl=CV.filter(c=>c.hcl&&c.hcl!=='Ninguna'), nocc=CV.filter(c=>c.pcm==null);
  const Q=[
    ['Alta','Laboratorio pendiente','La taxonomía (Gran Grupo), la Capacidad de Uso Mayor y el Factor K son tentativos: dependen de textura, materia orgánica, pH, carbonatos y saturación de bases (UNALM, Entregable 3 al 11/10/2026). Las muestras de la Brigada 1 aún no figuran en la cadena de custodia consolidada.'],
    ['Media','118 bloques en el estudio de suelos, 117 en el proyecto','La capa BLOQUES_INTERVENCION del estudio incluye el bloque 74 (Huarmaca, 29.23 ha, sin calicata), que no forma parte de los 117 bloques de los tres volúmenes. 12,299.45 − 29.23 = 12,270.22 ha, la superficie del proyecto.'],
    ['Media','Velocidad de infiltración representativa','El Cuadro 19 del informe usa la curva promedio de las repeticiones; la capa PRUEBAS_INFILTRACION_EJECUTADAS registra la mediana de las velocidades por repetición (p. ej. INF-01: 0.59 frente a 0.41 cm/h). Esta pestaña usa el Cuadro 19.'],
    ['Media','Calicatas reubicadas a más de 300 m',rel.length?`${rel.map(c=>`${c.c} (${n0(c.dz)} m)`).join(', ')}. Todas quedan dentro de su unidad de suelo y zona homogénea; los chequeos confirman la unidad.`:'Ninguna en este volumen.'],
    ['Baja','Calicatas sin muestra registrada o sin profundidad de contacto',[ns.length?`Sin muestras en la ficha: ${ns.map(c=>c.c).join(', ')}.`:'',nocc.length?`Sin profundidad efectiva numérica: ${nocc.map(c=>c.c).join(', ')}.`:''].filter(Boolean).join(' ')||'Ninguna en este volumen.'],
    ['Baja','Reacción al HCl',hcl.length?`Reacción ${hcl.map(c=>`${c.hcl.toLowerCase()} en ${c.c}`).join(', ')}; se confirmará con carbonatos totales.`:'Sin reacción al HCl en las calicatas de este volumen.'],
    ['Baja','Régimen de humedad','Los Torriorthents del matorral desértico suponen régimen arídico; si resulta ústico, pasan a Ustorthents. Los epipedones mólicos se confirman con saturación de bases y carbono orgánico.']
  ];
  if(CV.some(c=>c.c==='CAL-43')) Q.push(['Baja','CAL-43 · Norte corregido','El Norte de campo estaba incompleto; se corrigió en gabinete con el track del GPS a 9 377 600 m.']);
  if(CV.some(c=>c.c==='CAL-11')) Q.push(['Baja','CAL-11 · horizonte Bt por confirmar','Posible Alfisol (Haplustalfs) si el Bt resulta argílico; confirmar antes de diseñar.']);
  if(CV.some(c=>c.c==='CAL-46')) Q.push(['Baja','CAL-46 · epipedón no determinado','El horizonte superficial bajo el Oi no se clasificó en campo.']);
  $('sueQC').innerHTML='<thead><tr><th>Efecto</th><th class="l">Tema</th><th class="l">Observación y tratamiento en esta pestaña</th></tr></thead><tbody>'+Q.map(([p,t,d])=>`<tr><td>${pill(p,p==='Alta'?'p-cri':p==='Media'?'p-adv':'p-neu')}</td><td class="l"><b>${esc(t)}</b></td><td class="wrap" style="max-width:none">${esc(d)}</td></tr>`).join('')+'</tbody>';
}

/* ---------- 7.11 Ficha por bloque: panel de suelos ---------- */
function fichaSue(){
  const f=$('ficha'); if(!f) return;
  let busy=false;
  const add=()=>{ if(busy) return; const cod=I.engine==='a'?(document.querySelector('#picker button[aria-pressed="true"]')||{}).dataset?.cod:($('selB')||{}).value; if(!cod) return;
    const old=f.querySelector('.sue-fb'); if(old&&old.dataset.b===cod) return; busy=true; if(old) old.remove();
    const cs=CV.filter(c=>c.b===cod), bl=SU.BL[cod]||{us:[]}, sm=bl.us.filter(u=>SU.U[u]);
    const d=document.createElement('div'); d.className='sue-fb'; d.dataset.b=cod;
    d.innerHTML=`<h4>Suelos · E5 <small>GeoSIG, Entregable 2 · campo al 20/09/2026 · laboratorio pendiente</small></h4>`+
      `<p class="fb-txt">${esc(bl.ti||'—')} · ${esc(bl.eco||'—')} · ${esc(bl.zv||'—')} · ${bl.us.length} unidades de suelo en el bloque, ${sm.length} muestreadas${sm.length?': '+sm.map(u=>`<span class="mono">${esc(SU.U[u].s)}</span> (${SU.U[u].c.join(', ')})`).join('; '):''}.</p>`+
      (cs.length?`<div class="sue-fbg">${cs.map(c=>{const a=apt(c),inf=infOf(c)||infUnit(c);return `<div class="sue-fbc"><div class="sue-prof">${profSVG(c,{W:150,H:250})}</div><div><b class="mono">${esc(c.c)}</b> <span class="pill ${c.t==='B'?'p-cri':c.t==='R'?'p-ok':'p-neu'}">${TIPO[c.t]}</span>${kvS([['Prof. efectiva',c.pcm!=null?c.pcm+' cm':'—'],['Contacto',esc(c.con)],['Clasificación',esc(c.tax),true],['WRB · grupo',esc(c.wrb)+' · G'+c.g],['Horizonte A',`${sw(c.col)} ${esc(c.mun)} · ${esc(c.tex)}`],['Erosión',esc(c.ero)],['CUM campo · gabinete',esc(c.cum)+' · '+esc(c.cumu)],['Infiltración',inf?`${fmt(inf.ib,2)} cm/h · ${esc(inf.cl)}${inf.cal!==c.c?' (analogía)':''}`:'—'],['Zanjas',pill(a.z,a.zc)]])}</div></div>`;}).join('')}</div>`:
       `<p class="fb-txt muted" style="font-style:italic">Sin calicata propia en el bloque. ${sm.length?'Sus unidades de suelo muestreadas se caracterizan por analogía con las calicatas indicadas.':'Ninguna de sus unidades de suelo tiene calicata: caracterizar por analogía de litología y clima en el Entregable 4.'}</p>`)+
      `<p class="fb-txt"><button type="button" class="t8go go" style="background:none;border:0;padding:0;color:var(--anin-2);font-weight:600;cursor:pointer">Ver la pestaña Suelos →</button></p>`;
    f.appendChild(d); d.querySelector('.t8go').addEventListener('click',()=>{ if(cs[0]){SEL=cs[0].c; suePick(); sueMap();} goTab('sue');});
    busy=false; };
  new MutationObserver(()=>setTimeout(add,0)).observe(f,{childList:true});
  add();
}

/* ---------- 7.12 Panorama: hallazgo de suelos ---------- */
function panoSue(){
  const host=$('tFinds')||$('fDT'); if(!host||!CV.length) return;
  const sh=CB.filter(c=>c.pcm!=null&&c.pcm<50), ent=CV.filter(c=>ORD(c.tax)==='Entisol').length, x=CV.filter(c=>cumG(c.cum)==='X').length;
  const html=`<i></i><div class="body"><h3>El suelo de los bloques es joven, somero y de protección</h3><p>El Estudio de Suelos (E5, fase de campo) describió <strong>${CV.length} calicatas</strong> en ${new Set(CV.map(c=>c.b)).size} bloques de ${I.prov}: ${ent} son Entisoles tentativos y ${x} se pre-clasifican en protección (X). La profundidad efectiva mediana de las calicatas de bloque es de ${n0(median(CB.map(c=>c.pcm)))} cm y ${sh.length} de ${CB.length} tienen contacto antes de 50 cm, donde no conviene excavar zanjas. ${IV.length?`Las ${IV.length} pruebas de infiltración del volumen dan ${IV.map(s=>fmt(s.ib,2)).join(', ')} cm/h.`:''}</p><p class="src"><a href="#sue" class="golink" data-sue="1">Ver la pestaña Suelos →</a></p></div>`;
  const ensure=()=>{ if(!host.children.length||host.querySelector('.sue-find')) return;
    const a=document.createElement('article'); a.className='find adv sue-find'; a.innerHTML=html; host.appendChild(a);
    a.querySelector('[data-sue]').addEventListener('click',e=>{e.preventDefault();goTab('sue');}); };
  new MutationObserver(()=>setTimeout(ensure,0)).observe(host,{childList:true}); ensure();
}
