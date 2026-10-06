<script>
/* =====================================================================
   CAPA DE ACTUALIZACIÓN — corte 06/10/2026 (común a los Volúmenes I, II y III)
   Lee window.UPD: libro «Gráficos DS Consolidado» del aplicativo IN Piura
   (06/10/2026, valores y %), sin nombres ni datos personales; ecosistemas
   V6 por bloque (MINAM 2018), área de influencia aprobada y matriz de
   pendientes V6 (Copernicus GLO-30). Inserta sus bloques al final de las
   pestañas existentes; no crea pestañas nuevas.
   ===================================================================== */
(function(){
'use strict';
const U=window.UPD; if(!U) return;
const $=id=>document.getElementById(id);
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const fmt=(v,d)=>(v==null||v===''||Number.isNaN(+v))?'—':Number(v).toLocaleString('en-US',{minimumFractionDigits:d,maximumFractionDigits:d});
const n0=v=>fmt(v,0),n1=v=>fmt(v,1),n2=v=>fmt(v,2),n3=v=>fmt(v,3);
const num=v=>(v==null||v===''||!isFinite(+v))?null:+v;
const nv=v=>num(v)??0;
const fdate=s=>{const m=String(s||'').match(/^(\d{4})-(\d{2})-(\d{2})/);return m?`${m[3]}/${m[2]}/${m[1]}`:(s==null||s===''?'—':String(s));};
const CORTE=U.corte;

/* ---------- utilidades de tabla consolidada ---------- */
function model(t){
  const h=t.h, calc=new Set(t.calc||[]);
  const vi=h.map((_,i)=>i).filter(i=>i>0&&!calc.has(i));
  const rows=t.r.filter(r=>r&&r[0]!=null&&!String(r[0]).startsWith('PROMEDIO'));
  const prom=t.r.filter(r=>r&&String(r[0]).startsWith('PROMEDIO'));
  const V=rows.map(r=>vi.map(i=>num(r[i])));
  const rowT=V.map(v=>v.reduce((a,b)=>a+(b||0),0));
  const colT=vi.map((_,j)=>V.reduce((a,v)=>a+(v[j]||0),0));
  const all=colT.reduce((a,b)=>a+b,0);
  const pc=(i,j)=>{const v=V[i][j]; if(v==null) return null;
    if(t.m==='col') return colT[j]?v/colT[j]*100:null;
    if(t.m==='row') return rowT[i]?v/rowT[i]*100:null;
    if(t.m==='all') return all?v/all*100:null;
    if(t.m==='base') return t.base?v/t.base*100:null;
    if(t.m==='val') return null;
    return v;};
  return {h,vi,rows,prom,V,rowT,colT,all,pc,lab:h[0]};
}
function palette(names){
  const L=names.map(s=>String(s).toLowerCase());
  if(L.some(s=>s==='sí'||s==='si')) return L.map(s=>(s==='sí'||s==='si')?'--ux-yes':'--ux-no');
  if(L.includes('alto')||L.includes('medio')) return L.map(s=>({alto:'--ux-q5',medio:'--ux-q3',bajo:'--ux-q1'}[s]||'--ux-na'));
  if(L.some(s=>s.includes('autoconsumo')||s.includes('mercado'))) return L.map(s=>s.startsWith('100% auto')?'--ux-q1':s.startsWith('mayormente auto')?'--ux-q2':s.startsWith('mixto')?'--ux-q3':s.startsWith('mayormente merc')?'--ux-q4':s.startsWith('100% merc')?'--ux-q5':'--ux-na');
  if(L.some(s=>s.startsWith('hombre'))) return L.map(s=>s.startsWith('hombre')?'--ux-sex-h':s.startsWith('mujer')?'--ux-sex-m':'--ux-na');
  if(L.some(s=>s.startsWith('menores'))) return L.map(s=>s.startsWith('menores')?'--ux-c1':s.startsWith('de 18')?'--ux-c2':'--ux-c4');
  if(L.some(s=>s.startsWith('f-ds-0'))) return L.map(s=>s.startsWith('f-ds-01')?'--f01':s.startsWith('f-ds-02')?'--f02':'--f03');
  return L.map((_,i)=>['--ux-c1','--ux-c2','--ux-c3','--ux-c4','--ux-c5','--ux-c6'][i%6]);
}
const pctMode={col:'% de la columna',row:'% de la fila',all:'% del total de la tabla',base:'% de los CP con dato',none:'%',val:'Valores absolutos'};

/* gráfico de barras horizontales (apiladas si hay varias columnas) */
function chartSVG(t,M){
  const names=M.vi.map(i=>M.h[i]), cols=palette(names), multi=M.vi.length>1;
  const labs=M.rows.map(r=>String(r[0]));
  const LW=Math.min(250,Math.max(90,Math.max(...labs.map(s=>s.length))*6.3+12)), RW=multi?78:118, W=660, BH=18, G=8;
  const H=8+M.rows.length*(BH+G)+22;
  const isPct=t.m==='none';
  const tot=M.rows.map((_,i)=>isPct?Math.max(...M.V[i].map(v=>v||0)):M.rowT[i]);
  const mx=isPct?100:Math.max(1,...tot);
  const sc=v=>(W-LW-RW)*v/mx;
  let g='';
  const step=isPct?25:niceStep(mx);
  for(let v=0;v<=mx+1e-9;v+=step){const x=LW+sc(v);g+=`<line class="g" x1="${x}" x2="${x}" y1="2" y2="${H-18}"/><text class="t" x="${x}" y="${H-5}" text-anchor="middle">${isPct?v+' %':n0(v)}</text>`;}
  M.rows.forEach((r,i)=>{const y=6+i*(BH+G); let x=LW;
    g+=`<text x="${LW-8}" y="${y+BH/2+4}" text-anchor="end">${esc(trunc(labs[i],38))}<title>${esc(labs[i])}</title></text>`;
    if(isPct){ // valores ya en %: barras agrupadas finas
      const bh=BH/M.vi.length;
      M.vi.forEach((_,j)=>{const v=M.V[i][j]; if(v==null) return; g+=`<rect x="${LW}" y="${y+j*bh}" width="${Math.max(1.5,sc(v))}" height="${bh-1}" rx="1" style="fill:var(${cols[j]})"><title>${esc(labs[i])} · ${esc(names[j])}: ${n1(v)} %</title></rect>`;});
      g+=`<text class="v" x="${LW+sc(Math.max(...M.V[i].map(v=>v||0)))+6}" y="${y+BH/2+4}">${M.V[i].map(v=>v==null?'—':n0(v)+'%').join(' · ')}</text>`;
      return;}
    M.vi.forEach((_,j)=>{const v=M.V[i][j]; if(!v) return; const w=sc(v), p=M.pc(i,j);
      g+=`<rect x="${x}" y="${y}" width="${Math.max(1.5,w-(multi?1:0))}" height="${BH}" rx="1.5" style="fill:var(${cols[j]})"><title>${esc(labs[i])} · ${esc(names[j])}: ${n0(v)}${p!=null?' ('+n1(p)+' %)':''}</title></rect>`; x+=w;});
    const p0=multi?null:M.pc(i,0);
    g+=`<text class="v" x="${x+6}" y="${y+BH/2+4}">${multi?n0(tot[i]):n0(M.V[i][0])+(p0!=null?' · '+n1(p0)+' %':'')}</text>`;});
  const leg=(multi||isPct)?`<div class="ux-leg">${names.map((s,j)=>`<span><i style="background:var(${cols[j]})"></i>${esc(s)}</span>`).join('')}</div>`:'';
  return `<svg class="ux-svg" viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(t.t)}">${g}</svg>${leg}`;
}
function niceStep(mx){const raw=mx/5, p=Math.pow(10,Math.floor(Math.log10(raw))), f=raw/p; return (f<=1?1:f<=2?2:f<=5?5:10)*p;}
function trunc(s,n){return s.length>n?s.slice(0,n-1)+'…':s;}

/* tabla de valores y porcentajes */
function tableHTML(t,M){
  const names=M.vi.map(i=>M.h[i]); const isPct=t.m==='none'; const multi=M.vi.length>1;
  let th=`<th class="l" scope="col">${esc(M.lab)}</th>`;
  const isVal=t.m==='val';
  names.forEach(s=>{th+=isPct?`<th scope="col">${esc(s)} (%)</th>`:isVal?`<th scope="col">${esc(s)}</th>`:`<th scope="col">${esc(s)}</th><th scope="col">%</th>`;});
  if(multi&&!isPct) th+='<th scope="col">Total</th>';
  const mxp=100;
  const body=M.rows.map((r,i)=>{let tds=`<td class="l">${esc(r[0])}</td>`;
    M.vi.forEach((_,j)=>{const v=M.V[i][j], p=M.pc(i,j);
      if(isPct){tds+=`<td>${v==null?'<span class="na">—</span>':mbar(v,mxp)}</td>`;return;}
      tds+=`<td${v==null?' class="na"':''}>${v==null?'—':n0(v)}</td>`+(isVal?'':`<td>${p==null?'<span class="na">—</span>':mbar(p,mxp)}</td>`);});
    if(multi&&!isPct) tds+=`<td><b>${n0(M.rowT[i])}</b></td>`;
    return `<tr>${tds}</tr>`;}).join('');
  let foot='';
  if(isVal) foot='';
  else if(!isPct&&t.m!=='base'){foot=`<tfoot><tr><td class="l">Total</td>${M.vi.map((_,j)=>`<td>${n0(M.colT[j])}</td><td class="p">${t.m==='col'?'100.0 %':t.m==='all'&&M.all?n1(M.colT[j]/M.all*100)+' %':''}</td>`).join('')}${multi?`<td>${n0(M.all)}</td>`:''}</tr></tfoot>`;}
  if(t.m==='base'&&t.base) foot=`<tfoot><tr><td class="l">Base: centros poblados con dato</td><td>${n0(t.base)}</td><td class="p">100 % de la base</td></tr></tfoot>`;
  const prom=M.prom.length?`<tr><td class="l"><i>Promedio de los CP con dato</i></td>${M.vi.map((i,j)=>{const vals=M.V.map(v=>v[j]).filter(v=>v!=null);return `<td colspan="${isPct?1:2}"><i>${vals.length?n1(vals.reduce((a,b)=>a+b,0)/vals.length)+(isPct?' %':''):'—'}</i></td>`;}).join('')}</tr>`:'';
  return `<div class="ux-tw"><table><thead><tr>${th}</tr></thead><tbody>${body}${prom}</tbody>${foot}</table></div>`;
}
function mbar(p,mx){return `<div class="ux-mb"><span><i style="width:${Math.max(0,Math.min(100,p/mx*100))}%"></i></span><em>${n1(p)} %</em></div>`;}

function card(t,opt={}){
  const M=model(t); if(!M.rows.length) return '';
  const m=t.t.match(/^([A-Z])\. (.*)$/); const k=m?m[1]:'', ttl=m?m[2]:t.t;
  const big=M.rows.length>16||t.m==='none'&&M.vi.length>1||window.innerWidth<640;
  const notes=(t.n||[]).map(s=>s.replace(/^Fuente:\s*/,'Fuente: ')).join(' ');
  const pm=t.m!=='none'&&t.m!=='val'?`<span class="ux-stamp" title="Base del porcentaje">${esc(pctMode[t.m])}</span>`:'';
  return `<figure class="ux-card${big||opt.wide?' wide':''}"><figcaption class="ux-h">${k?`<span class="ux-k">${opt.pre||''}${k}</span>`:''}<h3>${esc(ttl)}</h3></figcaption>
    ${t.d?`<p class="ux-d">${esc(t.d)}</p>`:''}
    ${big?tableHTML(t,M):chartSVG(t,M)+`<details class="ux-tbl"><summary>Valores y porcentajes</summary>${tableHTML(t,M)}</details>`}
    <div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center">${pm}${notes?`<p class="ux-n">${esc(notes)}</p>`:''}</div></figure>`;
}

/* tabla de registros (T …) ordenable y filtrable */
const TX={};
function recTable(id,T,{wrapCols=[],dateCols=[],filters=[],search=true,caption=''}={}){
  if(!T||!T.h) return '';
  TX[id]={T,wrapCols,dateCols,filters,k:-1,dir:1};
  const fl=filters.map(c=>{const i=T.h.indexOf(c); if(i<0) return ''; const vals=[...new Set(T.r.map(r=>r[i]).filter(v=>v!=null&&v!==''))].map(String).sort((a,b)=>a.localeCompare(b,'es',{numeric:true}));
    return `<label for="${id}_f${i}">${esc(c)}</label><select id="${id}_f${i}" data-i="${i}"><option value="">Todos</option>${vals.map(v=>`<option>${esc(v)}</option>`).join('')}</select>`;}).join('');
  return `<div class="ux-ctrl" data-for="${id}">${fl}${search?`<label for="${id}_q">Buscar</label><input id="${id}_q" type="search" placeholder="Centro poblado, bloque o texto">`:''}<span class="cnt" id="${id}_n"></span></div>
    <div class="ux-tw"><table id="${id}"><thead></thead><tbody></tbody></table></div>${caption?`<p class="foot">${caption}</p>`:''}`;
}
function paintRec(id){
  const X=TX[id]; if(!X) return; const {T}=X; const tb=$(id); if(!tb) return;
  const sel=[...document.querySelectorAll(`.ux-ctrl[data-for="${id}"] select`)].map(s=>[+s.dataset.i,s.value]).filter(x=>x[1]);
  const q=(($(id+'_q')||{}).value||'').toLowerCase();
  let rs=T.r.filter(r=>sel.every(([i,v])=>String(r[i]??'')===v)&&(!q||r.join(' ').toLowerCase().includes(q)));
  if(X.k>=0){const i=X.k; rs=[...rs].sort((a,b)=>{const x=a[i],y=b[i]; const nx=num(x),ny=num(y); if(nx!=null&&ny!=null) return (nx-ny)*X.dir; return String(x??'').localeCompare(String(y??''),'es',{numeric:true})*X.dir;});}
  tb.querySelector('thead').innerHTML='<tr>'+T.h.map((h,i)=>`<th class="s${X.wrapCols.includes(h)||i===0?' l':''}" data-i="${i}" scope="col"${X.k===i?' data-dir="1"':''}>${esc(h)}<span class="car">${X.k===i?(X.dir>0?'▲':'▼'):'▲▼'}</span></th>`).join('')+'</tr>';
  tb.querySelectorAll('thead th').forEach(th=>th.onclick=()=>{const i=+th.dataset.i; if(X.k===i) X.dir=-X.dir; else {X.k=i; X.dir=1;} paintRec(id);});
  tb.querySelector('tbody').innerHTML=rs.map(r=>'<tr>'+T.h.map((h,i)=>{const v=r[i]; const w=X.wrapCols.includes(h);
    const txt=X.dateCols.includes(h)?fdate(v):(typeof v==='number'?(Number.isInteger(v)?n0(v):n2(v)):(v==null||v===''?'—':String(v)));
    return `<td class="${w?'w':(i===0?'l':'')}${(v==null||v==='')?' na':''}">${esc(txt)}</td>`;}).join('')+'</tr>').join('');
  const n=$(id+'_n'); if(n) n.textContent=`${rs.length} de ${T.r.length} registros`;
}
function wireRec(id){document.querySelectorAll(`.ux-ctrl[data-for="${id}"] select`).forEach(s=>s.addEventListener('change',()=>paintRec(id)));
  const q=$(id+'_q'); if(q) q.addEventListener('input',()=>paintRec(id)); paintRec(id);}

/* inserción en una pestaña existente */
function addBlock(sec,html,{first=false,after=null}={}){
  const s=$(sec); if(!s) return null;
  const d=document.createElement('div'); d.className='ux-block'; d.innerHTML=html;
  if(after){const a=s.querySelector(after); if(a){a.after(d); return d;}}
  if(first){const h=s.querySelector('.lede')||s.querySelector('.sec-h'); if(h){h.after(d); return d;}}
  s.appendChild(d); return d;
}
const secH=(t,tag)=>`<div class="sec-h"><h2>${esc(t)}</h2><span class="tag">${esc(tag)}</span></div>`;
const sep='<div class="block-sep"></div>';
const markTab=p=>{const b=document.querySelector(`nav.tabs button[data-p="${p}"]`); if(b) b.dataset.upd='1';};
const DS=U.ds||{};
const F01=DS.f01||[];
const isLoc=t=>/Población|etaria/.test(t.t), isTen=t=>/tenencia|tituladas|Superposición/i.test(t.t);

/* ---------- 1. Hallazgos sociales: cifras de cabecera y novedades ---------- */
function hs(){
  const R=U.res; if(!R) return;
  const kp=R.cab.map(([k,v,d])=>`<div${/Fichas|Bloques con/.test(k)?' class="hl"':''}><dt>${esc(k)}</dt><dd>${typeof v==='number'?n0(v):esc(v)} <small>${esc(d||'')}</small></dd></div>`).join('');
  const fd=(U.nov||[]).map(([c,ttl,ps,src])=>`<div class="find ${c}"><i></i><div class="body"><h3>${esc(ttl)}</h3>${ps.map(p=>`<p>${p}</p>`).join('')}${src?`<p class="src">${esc(src)}</p>`:''}</div></div>`).join('');
  addBlock('p-hs',`<div style="margin:0 0 14px"><span class="ux-stamp">Consolidado DS · aplicativo IN Piura · emitido el ${esc(R.id['Fecha de emisión']||CORTE)}</span></div>
    <dl class="sx-kpis">${kp}</dl><div class="finds">${fd}</div>${sep}
    <div class="sec-h"><h2>Hallazgos de la revisión anterior</h2><span class="tag">Corte 03/10/2026 · vigentes salvo en las cifras actualizadas arriba</span></div>`,{first:true});
  markTab('hs');
}

/* ---------- 2. Cobertura: consolidado A–F ---------- */
function cob(){
  const C=DS.cob||[]; if(!C.length) return;
  const R=U.res, cont=(R.cont||[]).map(([s,n,a])=>`<li><b>${esc(s)}</b> · ${n} gráficos · ${esc(a)}</li>`).join('');
  addBlock('p-cob',`${sep}${secH('Diagnóstico social consolidado',`Aplicativo IN Piura · corte ${CORTE} · valores y %`)}
   <p class="lede">Libro «Gráficos DS Consolidado» del aplicativo IN Piura, emitido el ${esc(R.id['Fecha de emisión']||CORTE)} para el ámbito ${esc(R.id['Ámbito']||U.prov)}. Se cuentan las fichas vigentes, tras descartar reediciones y copias idénticas; un centro poblado con varias fichas F-DS-01 se cuenta una sola vez. Solo se grafica lo declarado en campo: los campos sin respuesta no se estiman ni se completan por analogía.</p>
   <div class="ux-grid">${C.map(t=>card(t)).join('')}</div>
   ${cont?`<ul class="sx-pert" style="margin-top:16px">${cont}</ul>`:''}
   ${(R.av||[]).filter(s=>s.startsWith('Aviso')).map(s=>`<div class="warnbox" style="margin-top:14px">${esc(s)}</div>`).join('')}`);
  markTab('cob');
}

/* ---------- 3. Localidades: demografía declarada ---------- */
function loc(){
  const L=F01.filter(isLoc), T=U.T.demo; if(!L.length&&!T) return;
  addBlock('p-loc',`${sep}${secH('Demografía declarada por centro poblado',`F-DS-01 consolidada · ${CORTE} · cada CP una sola vez`)}
   <p class="lede">Población, familias, mano de obra disponible y nivel educativo predominante que declara la ficha de referencia de cada centro poblado, frente a la población del catálogo INEI. Las diferencias se informan, no se corrigen.</p>
   <div class="ux-grid">${L.map(t=>card(t,{pre:'F01·'})).join('')}</div>
   ${T?`<div style="margin-top:18px">${recTable('uxDemo',T,{dateCols:['Fecha de la ficha de referencia'],filters:['Distrito'],caption:'Fuente: aplicativo IN Piura, consolidado DS del '+CORTE+'. «—»: la ficha no consigna el dato. La columna de referencia indica la fecha de la ficha cuyos datos representan al CP; se omiten los nombres de responsables e informantes (Ley N.° 29733).'})}</div>`:''}`);
  if(T) wireRec('uxDemo'); markTab('loc');
}

/* ---------- 4. Servicios y economía ---------- */
function soc(){
  const S=F01.filter(t=>!isLoc(t)&&!isTen(t)), T=U.T.act; if(!S.length) return;
  addBlock('p-soc',`${sep}${secH('Servicios, medios de vida y gobernanza',`F-DS-01 consolidada · ${CORTE} · valores y %`)}
   <p class="lede">Servicios básicos, actividades económicas con su destino de producción, programas sociales, capacidades organizativas y percepciones, con cada centro poblado contado una sola vez. Los porcentajes indican su base en cada gráfico.</p>
   <div class="ux-grid">${S.map(t=>card(t,{pre:'F01·'})).join('')}</div>
   ${T?`${sep}${secH('Actividades económicas por centro poblado','Registro de la F-DS-01 · ordenable y filtrable')}${recTable('uxAct',T,{filters:['Actividad / Rubro','Destino de la producción'],wrapCols:['Productos principales'],caption:'Ingresos tal como se declaran en la ficha: mezclan montos mensuales, anuales y por campaña, y no se promedian.'})}`:''}`);
  if(T) wireRec('uxAct'); markTab('soc');
}

/* ---------- 5. Actores y entrevistas ---------- */
function act(){
  const A=DS.f02||[], E=DS.f03||[], TA=U.T.actores, TE=U.T.entrev; if(!A.length&&!E.length) return;
  addBlock('p-act',`${sep}${secH('Mapeo de actores consolidado',`F-DS-02 · ${CORTE} · cada actor una sola vez`)}
   <div class="ux-grid">${A.map(t=>card(t,{pre:'F02·'})).join('')}</div>
   ${TA?`<div style="margin-top:18px">${recTable('uxActT',TA,{filters:['Distrito','Tipo','Influencia'],wrapCols:['Rol frente al proyecto'],caption:'Cargo tomado de la ficha; se omiten nombres, DNI y teléfonos (Ley N.° 29733).'})}</div>`:''}
   ${E.length?`${sep}${secH('Entrevistas a autoridades y líderes',`F-DS-03 · ${CORTE}`)}<div class="ux-grid">${E.map(t=>card(t,{pre:'F03·'})).join('')}</div>
     ${TE?`<div style="margin-top:18px">${recTable('uxEnt',TE,{dateCols:['Fecha'],filters:['Distrito','Género'],caption:'Una fila por entrevistado. Se omite el nombre aunque el entrevistado consienta su uso: el volumen se difunde fuera del equipo.'})}</div>`:''}`
     :`${sep}<div class="warnbox"><b>F-DS-03 sin aplicar.</b> Al ${CORTE} no hay entrevistas a autoridades y líderes registradas en el ámbito de ${esc(U.prov)}.</div>`}`);
  if(TA) wireRec('uxActT'); if(TE) wireRec('uxEnt'); markTab('act');
}

/* ---------- 6. Riesgos y oportunidades: tenencia de la tierra ---------- */
function ten(){
  const L=F01.filter(isTen), T=U.T.ten; if(!L.length&&!T) return;
  addBlock('p-rie',`${sep}${secH('Tenencia de la tierra por bloque',`F-DS-01, numeral 4 · ${CORTE} · insumo del tamizaje predial`)}
   <p class="lede">Régimen predominante, titulación, conflictos de linderos y superposición con tierras comunales, con un valor por bloque (el más frecuente entre sus fichas F-DS-01). Anticipan con quién se suscriben las actas de libre disponibilidad y dónde habrá observaciones prediales.</p>
   <div class="ux-grid">${L.map(t=>card(t,{pre:'F01·'})).join('')}</div>
   ${T?`<div style="margin-top:18px">${recTable('uxTen',T,{filters:['Distrito','Régimen predominante de tenencia'],caption:'«—»: la ficha no consigna el dato; no equivale a «No».'})}</div>`:''}`);
  if(T) wireRec('uxTen'); markTab('rie');
}

/* ---------- 7. Discrepancias: control de calidad y fichas por CP ---------- */
function disc(){
  const C=U.T.cal, F=U.T.fichas; if(!C&&!F) return;
  addBlock('p-disc',`${sep}${secH('Control de calidad del consolidado',`Aplicativo IN Piura · ${CORTE}`)}
   <p class="lede">Observaciones que el propio consolidado registra al reunir las fichas: fichas de un mismo centro poblado que no coinciden (se usa el valor más frecuente), ámbitos repetidos o agrupados, poblaciones que difieren del INEI en más del doble y familias por actividad que superan las del centro poblado. Ninguna ficha se modificó.</p>
   ${C?recTable('uxCalT',C,{dateCols:['Fecha de la ficha de referencia'],filters:['Tema','Distrito'],wrapCols:['Detalle']}):''}
   ${F?`${sep}${secH('Fichas F-DS-01 por centro poblado','Cuál representa a cada CP y cuáles se suman')}${recTable('uxFic',F,{dateCols:['Fecha'],filters:['Distrito','Ficha de referencia del CP','Se suma en el análisis'],wrapCols:['Se suma en el análisis'],caption:'Se omiten responsables e informantes (Ley N.° 29733). «Ficha de referencia del CP: Sí» marca la ficha cuyos datos representan al centro poblado.'})}`:''}`);
  if(C) wireRec('uxCalT'); if(F) wireRec('uxFic'); markTab('disc');
}

/* ---------- 8. Condicionantes: matriz de pendientes V6 ---------- */
const SL=[['A','0–2 %','--sl-a'],['B','2–4 %','--sl-b'],['C','4–8 %','--sl-c'],['D','8–15 %','--sl-d'],['E','15–25 %','--sl-e'],['F','25–50 %','--sl-f'],['G','50–75 %','--sl-g'],['H','> 75 %','--sl-h']];
// pend[b] = [dist, mc, area, este, norte, media, mediana, p90, max, A..H(8), ha75]
const P=U.pend||{};
const pIdx={area:2,media:5,med:6,p90:7,max:8,cls:9,ha75:17};
function stack(p){return `<div class="ux-stack" role="img" aria-label="Clases de pendiente">${SL.map(([k,,c],j)=>{const v=p[pIdx.cls+j]; return v>0?`<span style="width:${v}%;background:var(${c})" title="${k} ${SL[j][1]}: ${n2(v)} %"></span>`:'';}).join('')}</div>`;}
function flag75(h){return h>=25?'<span class="ux-flag h">Revisar idoneidad</span>':h>=10?'<span class="ux-flag m">Sectorizar</span>':'<span class="ux-flag l">Conforme</span>';}
function pend(){
  const B=Object.keys(P); if(!B.length) return;
  const rows=B.map(b=>[b,P[b]]).sort((a,b)=>b[1][9+7]-a[1][9+7]);
  const A=B.reduce((s,b)=>s+P[b][2],0), A75=B.reduce((s,b)=>s+P[b][17],0), med=B.reduce((s,b)=>s+P[b][2]*P[b][5],0)/A;
  const nH=rows.filter(([,p])=>p[16]>=25).length;
  const dist={}; B.forEach(b=>{const d=P[b][0]; dist[d]=dist[d]||[0,0,0]; dist[d][0]++; dist[d][1]+=P[b][2]; dist[d][2]+=P[b][17];});
  addBlock('p-cond',`${sep}${secH('Pendientes por bloque · Matriz V6','Copernicus GLO-30 · método de Horn · clases A–H · 05/10/2026')}
   <p class="lede">Pendiente en porcentaje calculada sobre el polígono de cada bloque con el DEM Copernicus GLO-30 (30 m, reproyectado a UTM WGS 84 Zona 17S). La clase H (> 75 %) coincide con el límite del criterio de idoneidad del Paso 4 de la selección de bloques: la superficie en esa clase no admite obras de conservación de suelos y se trata con revegetación o se redelimita.</p>
   <dl class="sx-kpis"><div class="hl"><dt>Pendiente media ponderada</dt><dd>${n1(med)} % <small>${n0(B.length)} bloques · ${n2(A)} ha</small></dd></div>
     <div><dt>Superficie con pendiente &gt; 75 %</dt><dd>${n2(A75)} ha <small>${n1(A75/A*100)} % del ámbito de la provincia · estimada con el área del catálogo</small></dd></div>
     <div><dt>Bloques con ≥ 25 % en clase H</dt><dd>${nH} <small>la idoneidad por pendiente debe verificarse en campo</small></dd></div></dl>
   <div class="ux-leg" style="margin:0 0 10px">${SL.map(([k,r,c])=>`<span><i style="background:var(${c})"></i>${k} · ${r}</span>`).join('')}</div>
   <div class="ux-tw"><table><thead><tr><th class="l" scope="col">Bloque</th><th class="l" scope="col">Distrito</th><th scope="col">Área (ha)</th><th scope="col">Media %</th><th scope="col">Mediana %</th><th scope="col">P90 %</th><th scope="col">Máx. %</th><th class="l" scope="col">Clases A–H (% del bloque)</th><th scope="col">H &gt; 75 %</th><th scope="col">ha &gt; 75 %</th><th scope="col">Lectura</th></tr></thead>
   <tbody>${rows.map(([b,p])=>`<tr><td class="cod">${esc(b)}</td><td class="l">${esc(p[0])}</td><td>${n2(p[2])}</td><td>${n1(p[5])}</td><td>${n1(p[6])}</td><td>${n1(p[7])}</td><td>${n1(p[8])}</td><td>${stack(p)}</td><td>${n1(p[16])} %</td><td>${n2(p[17])}</td><td>${flag75(p[16])}</td></tr>`).join('')}</tbody>
   <tfoot><tr><td class="l">Total</td><td></td><td>${n2(A)}</td><td>${n1(med)}</td><td colspan="4"></td><td>${n1(A75/A*100)} %</td><td>${n2(A75)}</td><td></td></tr></tfoot></table></div>
   <div class="ux-tw" style="margin-top:14px;max-height:none"><table><thead><tr><th class="l" scope="col">Distrito</th><th scope="col">Bloques</th><th scope="col">Área (ha)</th><th scope="col">ha &gt; 75 %</th><th scope="col">% &gt; 75 %</th></tr></thead><tbody>${Object.entries(dist).sort((a,b)=>b[1][2]-a[1][2]).map(([d,v])=>`<tr><td class="l">${esc(d)}</td><td>${v[0]}</td><td>${n2(v[1])}</td><td>${n2(v[2])}</td><td>${mbar(v[2]/v[1]*100,100)}</td></tr>`).join('')}</tbody></table></div>
   <p class="foot">Fuente: Matriz de pendientes por bloque — catálogo V6 (ANIN-DIME-SESDI, 05/10/2026). El DEM es un modelo de superficie: bajo dosel alto la pendiente puede diferir de la del terreno. «ha &gt; 75 %» = área del catálogo × % de la clase H. Lectura: Revisar idoneidad (H ≥ 25 %), Sectorizar (10–25 %), Conforme (&lt; 10 %); es una regla de lectura de este volumen, no un criterio normativo. La media difiere de la pendiente zonal en grados de la pestaña «Los bloques» por fuente y método (D-P01). Clases según el Reglamento de Clasificación de Tierras por su Capacidad de Uso Mayor: verificar la versión vigente antes de citar.</p>`);
  markTab('cond');
}

/* ---------- 9. Geoespacial: ecosistemas V6 y área de influencia ---------- */
const ECOC={'Bes-cm':'--eco-bes','Agri':'--eco-agri','Ma':'--eco-ma','Br-mvoc':'--eco-br','Pa':'--eco-pa','':'--eco-sin'};
const E=U.eco||{};
const ecoBar=e=>`<div class="ux-stack" role="img" aria-label="Ecosistemas del bloque">${e.map(([n,s,,p])=>`<span style="width:${p}%;background:var(${ECOC[s]||'--eco-sin'})" title="${esc(n)}: ${n1(p)} %"></span>`).join('')}</div>`;
const imgOf=b=>`eco/Ecosistemas_Bloque_${encodeURIComponent(b)}.png`;
function eco(){
  const B=Object.keys(E); if(!B.length) return;
  const tot={}; let A=0, AI=0;
  B.forEach(b=>{E[b].e.forEach(([n,s,ha])=>{tot[n]=tot[n]||[s,0,0]; tot[n][1]+=ha; tot[n][2]++;}); A+=E[b].e.reduce((s,x)=>s+x[2],0); AI+=E[b].ai;});
  const T=Object.entries(tot).sort((a,b)=>b[1][1]-a[1][1]);
  const order=B.slice().sort((a,b)=>(P[a]?P[a][0]:'').localeCompare(P[b]?P[b][0]:'','es')||a.localeCompare(b,'es',{numeric:true}));
  addBlock('p-geo',`${sep}${secH('Ecosistemas de los bloques V6','Mapa Nacional de Ecosistemas (MINAM 2018) · atlas de 117 láminas · 05/10/2026')}
   <p class="lede">La unidad productora del proyecto es el ecosistema. Cada bloque V6 se intersectó con el Mapa Nacional de Ecosistemas del Perú (R.M. N.° 440-2018-MINAM), recortado a microcuencas y completado en M27B1 y M9B1; el área se prorratea al catálogo para que el total coincida. La escala de la capa es regional: el ecosistema de cada bloque se confirma en campo (F-DT-03).</p>
   <div class="ux-grid">
    <figure class="ux-card"><figcaption class="ux-h"><h3>Superficie por ecosistema en ${esc(U.prov)}</h3></figcaption>
     <div class="ux-tw" style="max-height:none"><table><thead><tr><th class="l" scope="col">Ecosistema</th><th scope="col">Símbolo</th><th scope="col">Bloques</th><th scope="col">Área (ha)</th><th scope="col">% del ámbito</th></tr></thead>
     <tbody>${T.map(([n,v])=>`<tr><td class="l"><i style="display:inline-block;width:11px;height:11px;border-radius:2px;margin-right:6px;vertical-align:-1px;background:var(${ECOC[v[0]]||'--eco-sin'})"></i>${esc(n)}</td><td class="mono">${esc(v[0])}</td><td>${v[2]}</td><td>${n2(v[1])}</td><td>${mbar(v[1]/A*100,100)}</td></tr>`).join('')}</tbody>
     <tfoot><tr><td class="l">Total</td><td></td><td>${B.length}</td><td>${n2(A)}</td><td class="p">100.0 %</td></tr></tfoot></table></div>
     <p class="ux-n">Área de influencia aprobada (AI_aprobado_2) de los mismos bloques: ${n2(AI)} ha, adicional al área de los bloques.</p></figure>
    <figure class="ux-card"><figcaption class="ux-h"><h3>Lámina por bloque</h3></figcaption>
     <div class="ux-ctrl"><label for="uxEcoSel">Bloque</label><select id="uxEcoSel">${order.map(b=>`<option value="${esc(b)}">${esc(b)} · ${esc(P[b]?P[b][0]:'')}</option>`).join('')}</select><a class="golink" id="uxEcoA" href="${imgOf(order[0])}" target="_blank" rel="noopener">Abrir en tamaño completo</a></div>
     <div class="ux-map"><img id="uxEcoImg" src="${imgOf(order[0])}" alt="Lámina de ecosistemas del bloque ${esc(order[0])}" loading="lazy"></div>
     <p class="ux-n">Atlas de ecosistemas — bloques de intervención V6, formato A4 (ANIN-DIME-SESDI, octubre de 2026). Coordenadas UTM WGS 84 Zona 17S.</p></figure>
   </div>
   <div class="ux-tw" style="margin-top:16px"><table><thead><tr><th class="l" scope="col">Bloque</th><th class="l" scope="col">Distrito</th><th scope="col">Área (ha)</th><th class="l" scope="col">Ecosistema dominante</th><th scope="col">%</th><th class="l" scope="col">Composición</th><th scope="col">Área de influencia (ha)</th></tr></thead>
   <tbody>${order.map(b=>{const e=E[b].e, d=e[0]||['—','',0,0]; return `<tr data-eco="${esc(b)}" style="cursor:pointer"><td class="cod">${esc(b)}</td><td class="l">${esc(P[b]?P[b][0]:'')}</td><td>${n3(e.reduce((s,x)=>s+x[2],0))}</td><td class="l">${esc(d[0])}</td><td>${n1(d[3])}</td><td>${ecoBar(e)}</td><td>${E[b].ai?n3(E[b].ai):'—'}</td></tr>`;}).join('')}</tbody></table></div>
   <div class="ux-leg" style="margin-top:10px">${Object.entries({'Bes-cm':'Bosque estacionalmente seco de colina y montaña','Agri':'Zona agrícola','Ma':'Matorral andino','Br-mvoc':'Bosque relicto montano de vertiente occidental','Pa':'Páramo'}).filter(([s])=>T.some(([,v])=>v[0]===s)).map(([s,n])=>`<span><i style="background:var(${ECOC[s]})"></i>${esc(n)} (${s})</span>`).join('')}</div>
   <p class="foot">Fuente: IN_Piura_Ecosistemas_por_Bloque_V6 y Bloques V6 con área de influencia (ANIN-DIME-SESDI, 2026), sobre el Mapa Nacional de Ecosistemas del Perú (MINAM 2018). Clic en una fila para ver su lámina.</p>`);
  const sel=$('uxEcoSel'), img=$('uxEcoImg'), a=$('uxEcoA');
  const show=b=>{sel.value=b; img.src=imgOf(b); img.alt='Lámina de ecosistemas del bloque '+b; a.href=imgOf(b);};
  sel.addEventListener('change',()=>show(sel.value));
  document.querySelectorAll('#p-geo tr[data-eco]').forEach(tr=>tr.addEventListener('click',()=>{show(tr.dataset.eco); sel.scrollIntoView({behavior:'smooth',block:'center'});}));
  markTab('geo');
}

/* ---------- 10. Ficha por bloque: ecosistemas, AI y pendientes ---------- */
function ficha(){
  const f=$('ficha'); if(!f) return;
  const add=()=>{const h=f.querySelector('.top h3'); if(!h) return; const b=h.textContent.trim(); if(!E[b]&&!P[b]) return;
    const old=f.querySelector('.ux-fb'); if(old&&old.dataset.b===b) return; if(old) old.remove();
    const e=(E[b]||{e:[]}).e, p=P[b];
    const d=document.createElement('div'); d.className='ux-fb'; d.dataset.b=b;
    d.innerHTML=`<h4>Ecosistemas V6 y pendientes del bloque <small>MINAM 2018 · área de influencia aprobada · Copernicus GLO-30 · ${CORTE}</small></h4>
     <div class="ux-fbg"><a href="${imgOf(b)}" target="_blank" rel="noopener" title="Abrir la lámina en tamaño completo"><div class="ux-map"><img src="${imgOf(b)}" alt="Lámina de ecosistemas del bloque ${esc(b)}" loading="lazy"></div></a>
      <div style="display:grid;gap:12px;min-width:0">
       <div class="ux-tw" style="max-height:none"><table><thead><tr><th class="l" scope="col">Ecosistema</th><th scope="col">ha</th><th scope="col">%</th></tr></thead><tbody>${e.map(([n,s,ha,pc])=>`<tr><td class="l"><i style="display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:6px;background:var(${ECOC[s]||'--eco-sin'})"></i>${esc(n)}</td><td>${n3(ha)}</td><td>${n1(pc)}</td></tr>`).join('')}</tbody></table></div>
       <dl class="ux-kv"><dt>Área de influencia aprobada</dt><dd>${E[b]&&E[b].ai?n3(E[b].ai)+' ha':'—'}</dd>
        ${p?`<dt>Pendiente media / mediana</dt><dd>${n1(p[5])} % / ${n1(p[6])} %</dd><dt>P90 / máxima</dt><dd>${n1(p[7])} % / ${n1(p[8])} %</dd><dt>Superficie &gt; 75 % (clase H)</dt><dd>${n1(p[16])} % · ${n2(p[17])} ha</dd>`:''}</dl>
       ${p?`${stack(p)}<div class="ux-leg">${SL.map(([k,r,c],j)=>p[9+j]>=1?`<span><i style="background:var(${c})"></i>${k} ${r}: ${n1(p[9+j])} %</span>`:'').join('')}</div>${flag75(p[16])}`:''}
      </div></div>`;
    const su=f.querySelector('.sue-fb'); if(su) su.before(d); else f.appendChild(d);
  };
  new MutationObserver(()=>setTimeout(add,0)).observe(f,{childList:true}); add();
}

function run(){
  [hs,cob,loc,soc,act,ten,disc,pend,eco,ficha].forEach(fn=>{try{fn();}catch(e){console.error('UPD '+fn.name,e);}});
}
if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',run); else run();
})();
</script>
