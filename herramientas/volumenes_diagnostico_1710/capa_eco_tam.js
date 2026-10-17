<script>
/* =====================================================================
   CAPA E6 · ECOSISTEMAS Y TAMIZAJE PREDIAL — corte 17/10/2026
   (común a los Volúmenes I, II y III)
   Lee window.ECO6 (Tablero E6: estado de los ecosistemas por bloque) y
   window.SPX (índice de mapas de situación predial v2 cruzado con el PI,
   el E6 y la F-DS-01 del consolidado DS del 17/10/2026). Llena las
   pestañas «Ecosistemas» y «Tamizaje predial» y agrega la lectura
   integrada en «Ficha por bloque», «Integralidad», «Consistencia»,
   «Geoespacial» y «Riesgos y oportunidades».
   ===================================================================== */
(function(){
'use strict';
const E=window.ECO6, S=window.SPX; if(!E||!S) return;
const $=id=>document.getElementById(id);
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const fmt=(v,d)=>(v==null||v===''||Number.isNaN(+v))?'—':Number(v).toLocaleString('en-US',{minimumFractionDigits:d,maximumFractionDigits:d});
const n0=v=>fmt(v,0),n1=v=>fmt(v,1),n2=v=>fmt(v,2),n3=v=>fmt(v,3);
const sum=(a,f)=>a.reduce((s,x)=>s+(+f(x)||0),0);
const B=E.b, X=S.x, XB=Object.fromEntries(X.map(r=>[r.b,r])), EB=Object.fromEntries(B.map(b=>[b.bloque,b]));
const PROV=E.prov;
const css=v=>getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const sortCod=(a,b)=>String(a).localeCompare(String(b),'es',{numeric:true});

/* ---------- paletas ---------- */
const ICN=['No degradado','Leve','Moderado','Severo'];
const ICV={'No degradado':'--e6-ok','Leve':'--e6-leve','Moderado':'--e6-mod','Severo':'--e6-sev'};
const TIPN={A:'A · Degradación convergente',B:'B · Enmascarada por verdor',C:'C · Núcleos degradados',D:'D · Conservada convergente'};
const TIPV={A:'--e6-ta',B:'--e6-tb',C:'--e6-tc',D:'--e6-td'};
const ECOV=e=>/Bes-cm/.test(e)?'--eco-bes':/agr[ií]cola/i.test(e)?'--eco-agri':/\(Ma\)/.test(e)?'--eco-ma':/Br-mvoc/.test(e)?'--eco-br':/\(Pa\)/.test(e)?'--eco-pa':'--eco-sin';
const ECOS=e=>{const m=String(e).match(/\(([^)]+)\)\s*$/); return m?m[1]:'Agri';};
const ECOT=e=>String(e).replace(/\s*\([^)]*\)\s*$/,'');
const PISON={A:'A < 1,000 m',B:'B 1,000–2,000 m',C:'C 2,000–2,900 m',D:'D > 2,900 m'};
const COB=[['cob1_pct','Dosel siempreverde denso','#1E6B3A'],['cob2_pct','Verde continua de porte bajo','#8CC63F'],['cob3_pct','Bosque ralo / matorral abierto','#9BA84E'],['cob4_pct','Caducifolio defoliado / matorral seco','#9C6B3F'],['cob5_pct','Herbazal seco','#E3C77E'],['cob6_pct','Suelo desnudo, roca, vías','#C9B8A6'],['cob8_pct','Sombra topográfica','#4A4A4A'],['cob9_pct','Nubes / neblina','#DCE4EA']];
const DN=[['DN1_ha','DN1 ≤ 0.2650','#8E1E14'],['DN2_ha','DN2 0.2650–0.3813','#B83A2E'],['DN3_ha','DN3 0.3813–0.4976','#E0A33A'],['DN4_ha','DN4 0.4976–0.6139','#7FB069'],['DN5_ha','DN5 > 0.6139','#2E7D4F']];
const ATT=[['ic_amenazas','Ausencia de amenazas'],['ic_cond_fisicas','Condiciones físicas'],['ic_div_estructural','Diversidad estructural'],['ic_funciones','Funciones (MSAVI)'],['ic_intercambios','Intercambios externos']];
const SITN=['Comunal','Parcialmente comunal','Mixta','Parcialmente privada','Privada','Tierras del Estado y/o posesionarios'];
const SITV={'Comunal':'--sp-cc','Parcialmente comunal':'--sp-pcc','Mixta':'--sp-mix','Parcialmente privada':'--sp-ppr','Privada':'--sp-pr','Tierras del Estado y/o posesionarios':'--sp-est'};
const PIV={'Muy alto':'--pi5','Alto':'--pi4','Medio':'--pi3','Bajo':'--pi2','Muy bajo':'--pi1'};
const pill=(t,v,dark)=>`<span class="e6-pill${dark?' dk':''}" style="background:var(${v})">${esc(t)}</span>`;
const icPill=c=>pill(c,ICV[c]||'--ux-na',c==='Leve');
const piPill=c=>c?pill('PI '+c,PIV[c]||'--ux-na',c==='Medio'||c==='Bajo'||c==='Muy bajo'):'—';
const sitPill=s=>`<span class="e6-sit"><i style="background:var(${SITV[s]||'--ux-na'})"></i>${esc(s)}</span>`;
const stack=(parts,label)=>`<div class="ux-stack" role="img" aria-label="${esc(label)}">${parts.filter(p=>p[1]>0).map(([t,v,c])=>`<span style="width:${v}%;background:${c.startsWith('--')?`var(${c})`:c}" title="${esc(t)}: ${n1(v)} %"></span>`).join('')}</div>`;
const mbar=(p,mx=100)=>`<div class="ux-mb"><span><i style="width:${Math.max(0,Math.min(100,p/mx*100))}%"></i></span><em>${n1(p)} %</em></div>`;
const secH=(t,tag)=>`<div class="sec-h"><h2>${esc(t)}</h2><span class="tag">${esc(tag)}</span></div>`;
const sep='<div class="block-sep"></div>';
const pl=(n,a,b)=>`${n} ${n===1?a:b}`;
const goTab=p=>{const b=document.querySelector(`nav.tabs button[data-p="${p}"]`); if(b) b.click();};
const goFicha=cod=>{goTab('fic'); setTimeout(()=>{const pk=[...document.querySelectorAll('#picker button')].find(b=>b.dataset.cod===cod);
  if(pk){pk.click();return;} const s=$('selB'); if(s){s.value=cod; s.dispatchEvent(new Event('change'));} },20);};
const imgOf=r=>r&&r.img&&!r.noimg?'sp/'+encodeURIComponent(r.img):null;

/* ---------- utilidades de mapa (polígonos E6 en el marco del tablero) ---------- */
function polyMap(id,colorOf,{sel=null,title=''}={}){
  const [x,y,w,h]=E.vb;
  const ps=B.map(b=>`<path class="e6-b${b.bloque===sel?' on':''}" d="${b.path}" data-b="${esc(b.bloque)}" tabindex="0" role="button" aria-label="Bloque ${esc(b.bloque)}" style="fill:${colorOf(b)}"><title>Bloque ${esc(b.bloque)} · ${esc(b.distrito)}</title></path>`).join('');
  return `<svg id="${id}" class="e6-map" viewBox="${x} ${y} ${w} ${h}" role="img" aria-label="${esc(title)}" preserveAspectRatio="xMidYMid meet">${ps}</svg>`;
}
function wireMap(svgId,onSel){
  const s=$(svgId); if(!s) return;
  const pick=t=>{const b=t.closest&&t.closest('path[data-b]'); if(b) onSel(b.dataset.b);};
  s.addEventListener('click',e=>pick(e.target));
  s.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault(); pick(e.target);}});
}
function markMap(svgId,cod){const s=$(svgId); if(!s) return; s.querySelectorAll('path[data-b]').forEach(p=>{const on=p.dataset.b===cod; p.classList.toggle('on',on); if(on) s.appendChild(p);});}
function sortable(tblId,cols,rows,onRow){
  const T=$(tblId); if(!T) return; let k=-1,dir=1;
  const paint=()=>{let rs=rows.slice(); if(k>=0){const f=cols[k].v; rs.sort((a,b)=>{const x=f(a),y=f(b); return (typeof x==='number'&&typeof y==='number'?x-y:String(x??'').localeCompare(String(y??''),'es',{numeric:true}))*dir;});}
    T.querySelector('thead').innerHTML='<tr>'+cols.map((c,i)=>`<th class="s${c.l?' l':''}" scope="col" data-i="${i}"${k===i?' data-dir="1"':''}>${c.h}<span class="car">${k===i?(dir>0?'▲':'▼'):'▲▼'}</span></th>`).join('')+'</tr>';
    T.querySelector('tbody').innerHTML=rs.map(r=>`<tr data-cod="${esc(r.b||r.bloque)}">${cols.map(c=>`<td class="${c.cls||''}">${c.f(r)}</td>`).join('')}</tr>`).join('');
    T.querySelectorAll('thead th').forEach(th=>th.onclick=()=>{const i=+th.dataset.i; if(k===i) dir=-dir; else {k=i; dir=1;} paint();});
    if(onRow) T.querySelectorAll('tbody tr').forEach(tr=>tr.onclick=()=>onRow(tr.dataset.cod));};
  paint();
}

/* =====================================================================
   A. PESTAÑA «ECOSISTEMAS» (E6)
   ===================================================================== */
const A=sum(B,b=>b.area_ha), BU=sum(B,b=>b.msavi_bajo_umbral_ha), RA=sum(B,b=>b.msavi_rest_activa_ha);
const ICm=sum(B,b=>b.ic_bloque*b.area_ha)/A, MSm=sum(B,b=>b.msavi_2024*b.area_ha)/A;
const nCls=c=>B.filter(b=>b.clase_estado===c);
const MS=B.filter(b=>b.clase_estado==='Moderado'||b.clase_estado==='Severo');
const nEvA=B.filter(b=>b.clase_evidencia==='A').length;
const spConf=new Set(), spEsp=new Set();
B.forEach(b=>(b.esp||[]).forEach(([n,,r])=>{ if(/^Conf/.test(r)) spConf.add(n); if(/^Esp/.test(r)) spEsp.add(n); }));
const DL={ra:sum(B,b=>b.delim_rest_activa_ha),en:sum(B,b=>b.delim_enriquecimiento_ha),co:sum(B,b=>b.delim_conservacion_ha),ag:sum(B,b=>b.delim_tamizaje_agricola_ha)};
const colorE6={
  bu:b=>{const p=b.msavi_bajo_umbral_pct||0; const s=[[0,[46,125,79]],[40,[201,162,39]],[70,[217,120,45]],[100,[184,58,46]]];
    for(let i=1;i<s.length;i++) if(p<=s[i][0]){const [a,ca]=s[i-1],[c,cb]=s[i];const t=(p-a)/(c-a);return `rgb(${ca.map((v,k)=>Math.round(v+(cb[k]-v)*t)).join(',')})`;} return 'rgb(184,58,46)';},
  ic:b=>`var(${ICV[b.clase_estado]})`, tip:b=>`var(${TIPV[b.tipologia_S2DR4]||'--ux-na'})`, eco:b=>`var(${ECOV(b.ecosistema)})`,
  pi:b=>`var(${PIV[(XB[b.bloque]||{}).pi]||'--ux-na'})`
};
function legE6(m){
  if(m==='bu') return `<span><i style="background:${colorE6.bu({msavi_bajo_umbral_pct:0})}"></i>0 %</span><span><i style="background:${colorE6.bu({msavi_bajo_umbral_pct:40})}"></i>40 %</span><span><i style="background:${colorE6.bu({msavi_bajo_umbral_pct:70})}"></i>70 %</span><span><i style="background:${colorE6.bu({msavi_bajo_umbral_pct:100})}"></i>100 % de la superficie bajo el umbral</span>`;
  if(m==='ic') return ICN.map(c=>`<span><i style="background:var(${ICV[c]})"></i>${c}</span>`).join('');
  if(m==='tip') return Object.entries(TIPN).map(([k,t])=>`<span><i style="background:var(${TIPV[k]})"></i>${t}</span>`).join('');
  if(m==='pi') return PIK.filter(c=>X.some(r=>r.pi===c)).map(c=>`<span><i style="background:var(${PIV[c]})"></i>PI ${c}</span>`).join('');
  return [...new Set(B.map(b=>b.ecosistema))].map(e=>`<span><i style="background:var(${ECOV(e)})"></i>${esc(ECOT(e))}</span>`).join('');
}
function fichaE6(b){
  if(!b) return '';
  const dn=DN.map(([k,t,c])=>[t,b[k],c]), dt=sum(dn,x=>x[1])||1, x=XB[b.bloque]||{};
  const esp=[...(b.esp||[])].sort((p,q)=>(/^Conf/.test(p[2])?0:/^Reg/.test(p[2])?1:2)-(/^Conf/.test(q[2])?0:/^Reg/.test(q[2])?1:2));
  return `<div class="e6-fh"><b class="mono">Bloque ${esc(b.bloque)}</b>${icPill(b.clase_estado)}<span class="mono sm">IC ${n2(b.ic_bloque)}</span>${piPill(x.pi)}</div>
   <p class="ux-d">${esc(b.distrito)} · ${esc(b.microcuenca)} · ${esc(b.ecosistema)}</p>
   <dl class="ux-kv"><dt>Área</dt><dd>${n2(b.area_ha)} ha</dd><dt>Centroide UTM 17S</dt><dd>${n0(b.utm[0])} E · ${n0(b.utm[1])} N</dd>
    <dt>Altitud media · pendiente</dt><dd>${n0(b.alt_media)} m · ${n1(b.pendiente_media_pct)} % · piso ${esc(PISON[b.piso_dom]||b.piso_dom)}</dd>
    <dt>MSAVI 2024</dt><dd>${n3(b.msavi_2024)} · <b>${n2(b.msavi_bajo_umbral_ha)} ha bajo umbral (${n1(b.msavi_bajo_umbral_pct)} %)</b></dd>
    <dt>Tipología S2DR4</dt><dd>${esc(TIPN[b.tipologia_S2DR4]||'—')}${b.tipologia_NDVI2025?` · NDVI 2025: ${esc(b.tipologia_NDVI2025)}`:''}</dd>
    <dt>Parcelas virtuales · evidencia</dt><dd>${n0(b.n_parcelas)} de 0.5 ha · clase ${esc(b.clase_evidencia)}</dd>
    <dt>Cobertura en campo (F-DT)</dt><dd>${b.dt_cobveg_campo!=null?n0(b.dt_cobveg_campo)+' %':'no registrada'}${b.fecha_verificacion?' · '+esc(b.fecha_verificacion):''}</dd></dl>
   <h4 class="e6-sh">Clases MSAVI 2024</h4>${stack(dn.map(([t,v,c])=>[t,v/dt*100,c]),'Clases MSAVI')}
   <div class="ux-leg">${dn.map(([t,v,c])=>`<span><i style="background:${c}"></i>${t.split(' ')[0]} ${n1(v)} ha</span>`).join('')}</div>
   <h4 class="e6-sh">Cobertura S2DR4 de 1 m</h4>${stack(COB.map(([k,t,c])=>[t,b[k]||0,c]),'Cobertura S2DR4')}
   <div class="ux-leg">${COB.filter(([k])=>b[k]>=1).map(([k,t,c])=>`<span><i style="background:${c}"></i>${t} ${n1(b[k])} %</span>`).join('')}</div>
   <h4 class="e6-sh">Índice de condición por atributo</h4>
   <div class="e6-att">${ATT.map(([k,t])=>`<span>${t}</span><span class="ux-mb"><span><i style="width:${(b[k]||0)*100}%"></i></span><em>${n2(b[k])}</em></span>`).join('')}</div>
   <h4 class="e6-sh">Delimitación interna (NDVI / coberturas)</h4>
   <dl class="ux-kv"><dt>Restauración activa</dt><dd>${n2(b.delim_rest_activa_ha)} ha</dd><dt>Enriquecimiento</dt><dd>${n2(b.delim_enriquecimiento_ha)} ha</dd><dt>Conservación</dt><dd>${n2(b.delim_conservacion_ha)} ha</dd><dt>Área agrícola a tamizar (Paso 5)</dt><dd>${n2(b.delim_tamizaje_agricola_ha)} ha</dd></dl>
   <h4 class="e6-sh">Especies · ${n0(b.n_especies_confirmadas)} confirmadas de ${n0(b.n_especies_esperadas)} esperadas</h4>
   <div class="e6-esp">${esp.map(([n,c,r])=>`<div><span><em>${esc(n)}</em>${c?` <small>${esc(c)}</small>`:''}</span><span class="e6-tag ${/^Conf/.test(r)?'c':/^Reg/.test(r)?'r':'e'}">${/^Conf/.test(r)?'Confirmada 2026':/^Reg/.test(r)?'Solo en campo':'Esperada'}</span></div>`).join('')||'<div>Sin especies asignadas</div>'}</div>
   <div class="e6-go"><button type="button" class="btn" data-go="tam" data-b="${esc(b.bloque)}">Situación predial del bloque</button><button type="button" class="btn" data-go="fic" data-b="${esc(b.bloque)}">Ficha por bloque</button></div>`;
}
function ecoTab(){
  const root=$('e6Root'); if(!root) return;
  const worst=[...B].sort((a,b)=>a.ic_bloque-b.ic_bloque)[0];
  const byEco={}; B.forEach(b=>{const k=b.ecosistema; byEco[k]=byEco[k]||{n:0,ha:0,bu:0,ic:0,cls:{}}; const o=byEco[k]; o.n++; o.ha+=b.area_ha; o.bu+=b.msavi_bajo_umbral_ha; o.ic+=b.ic_bloque*b.area_ha; o.cls[b.clase_estado]=(o.cls[b.clase_estado]||0)+1;});
  const byDist={}; B.forEach(b=>{const k=b.distrito; byDist[k]=byDist[k]||{n:0,ha:0,r:0,e:0,s:0}; const o=byDist[k]; o.n++; o.ha+=b.area_ha; o.r+=b.DN1_ha+b.DN2_ha; o.e+=b.DN3_ha; o.s+=b.DN4_ha+b.DN5_ha;});
  root.innerHTML=`
   <dl class="sx-kpis">
    <div class="hl"><dt>Bajo el umbral MSAVI 0.4976</dt><dd>${n2(BU)} ha <small>${n1(BU/A*100)} % de ${n2(A)} ha · indicador de brecha (distribución areal)</small></dd></div>
    <div><dt>Restauración activa</dt><dd>${n2(RA)} ha <small>MSAVI ≤ 0.3813 (clases DN1–DN2)</small></dd></div>
    <div><dt>Índice de condición medio</dt><dd>${n2(ICm)} <small>ponderado por área · MSAVI medio ${n3(MSm)}</small></dd></div>
    <div><dt>Bloques Moderado o Severo</dt><dd>${MS.length} <small>${n2(sum(MS,b=>b.area_ha))} ha · ${ICN.map(c=>`${nCls(c).length} ${c.toLowerCase()}`).join(' · ')}</small></dd></div>
    <div><dt>Evidencia de campo clase A</dt><dd>${nEvA} <small>de ${B.length} bloques con F-DT y parcelas virtuales</small></dd></div>
    <div><dt>Especies confirmadas en campo</dt><dd>${spConf.size} <small>taxones distintos · ${spEsp.size} solo esperados por firma espectral</small></dd></div></dl>
   <div class="e6-two">
    <figure class="ux-card"><figcaption class="ux-h"><h3>Mapa de los ${B.length} bloques</h3></figcaption>
     <div class="ux-ctrl"><label for="e6Col">Colorear por</label><select id="e6Col"><option value="bu">% de superficie bajo el umbral MSAVI</option><option value="ic">Clase de estado (índice de condición)</option><option value="tip">Tipología MSAVI–NDVI (S2DR4)</option><option value="eco">Ecosistema (ficha F-DT)</option><option value="pi">Peligro Integrado del bloque</option></select>
      <label for="e6Sel">Bloque</label><select id="e6Sel">${[...B].sort((a,b)=>sortCod(a.bloque,b.bloque)).map(b=>`<option value="${esc(b.bloque)}">${esc(b.bloque)} · ${esc(b.distrito)}</option>`).join('')}</select></div>
     <div class="e6-mapbox" id="e6MapBox">${polyMap('e6Map',colorE6.bu,{sel:worst.bloque,title:'Bloques de intervención coloreados por estado del ecosistema'})}</div>
     <div class="ux-leg" id="e6Leg">${legE6('bu')}</div>
     <p class="ux-n">Polígonos del catálogo V6 simplificados a 15 m (UTM WGS 84 Zona 17S), sin escala gráfica: la posición relativa es exacta, el tamaño de los bloques pequeños se ve reducido. Clic o Enter en un bloque para ver su ficha.</p></figure>
    <figure class="ux-card e6-ficha" id="e6Ficha" aria-live="polite">${fichaE6(worst)}</figure>
   </div>
   ${sep}${secH('Estado por ecosistema y por distrito','Superficie, brecha y clase de estado · E6')}
   <div class="ux-grid">
    <figure class="ux-card wide"><figcaption class="ux-h"><h3>Por ecosistema (ficha F-DT)</h3></figcaption>
     <div class="ux-tw" style="max-height:none"><table><thead><tr><th class="l" scope="col">Ecosistema</th><th scope="col">Bloques</th><th scope="col">Área (ha)</th><th scope="col">Bajo umbral (ha)</th><th scope="col">% bajo umbral</th><th scope="col">IC medio</th><th class="l" scope="col">Clases de estado</th></tr></thead>
     <tbody>${Object.entries(byEco).sort((a,b)=>b[1].ha-a[1].ha).map(([k,o])=>`<tr><td class="l"><i class="e6-sw" style="background:var(${ECOV(k)})"></i>${esc(ECOT(k))} <span class="mono sm">${esc(ECOS(k))}</span></td><td>${o.n}</td><td>${n2(o.ha)}</td><td>${n2(o.bu)}</td><td>${mbar(o.bu/o.ha*100)}</td><td>${n2(o.ic/o.ha)}</td><td class="l">${ICN.filter(c=>o.cls[c]).map(c=>`${icPill(c)} ${o.cls[c]}`).join(' ')}</td></tr>`).join('')}</tbody>
     <tfoot><tr><td class="l">Total</td><td>${B.length}</td><td>${n2(A)}</td><td>${n2(BU)}</td><td>${n1(BU/A*100)} %</td><td>${n2(ICm)}</td><td></td></tr></tfoot></table></div>
     <p class="ux-n">El ecosistema es el de la ficha F-DT (campo); difiere del Mapa Nacional de Ecosistemas (MINAM 2018) de la pestaña «Geoespacial» en los bloques listados en «Lectura integrada».</p></figure>
    <figure class="ux-card wide"><figcaption class="ux-h"><h3>Superficie bajo y sobre el umbral, por distrito</h3></figcaption>
     <div class="e6-hbar">${Object.entries(byDist).sort((a,b)=>b[1].ha-a[1].ha).map(([k,o])=>`<span>${esc(k)} <small>(${o.n})</small></span><span class="e6-hb">${[[o.r,'#B83A2E','DN1–DN2'],[o.e,'#E0A33A','DN3'],[o.s,'#4E9A6B','DN4–DN5']].map(([v,c,t])=>v>0?`<i style="width:${v/o.ha*100}%;background:${c}" title="${t}: ${n2(v)} ha"></i>`:'').join('')}</span><em>${n1((o.r+o.e)/o.ha*100)} % de ${n2(o.ha)} ha</em>`).join('')}</div>
     <div class="ux-leg"><span><i style="background:#B83A2E"></i>DN1–DN2 (≤ 0.3813, restauración activa)</span><span><i style="background:#E0A33A"></i>DN3 (0.3813–0.4976, enriquecimiento)</span><span><i style="background:#4E9A6B"></i>DN4–DN5 (sobre el umbral)</span></div>
     <h4 class="e6-sh">Delimitación interna de la provincia</h4>
     ${stack([['Restauración activa',DL.ra/A*100,'#B83A2E'],['Enriquecimiento',DL.en/A*100,'#E0A33A'],['Conservación',DL.co/A*100,'#2E7D4F'],['Área agrícola a tamizar',DL.ag/A*100,'#C9B8A6']],'Delimitación interna')}
     <div class="ux-leg"><span><i style="background:#B83A2E"></i>Restauración activa ${n2(DL.ra)} ha</span><span><i style="background:#E0A33A"></i>Enriquecimiento ${n2(DL.en)} ha</span><span><i style="background:#2E7D4F"></i>Conservación ${n2(DL.co)} ha</span><span><i style="background:#C9B8A6"></i>Área agrícola a tamizar ${n2(DL.ag)} ha</span></div>
     <p class="ux-n">La delimitación interna (NDVI y coberturas S2DR4) asigna cada sector del bloque a un tratamiento: es el insumo directo de las metas físicas de infraestructura verde (reforestación, revegetación, enriquecimiento) y de la exclusión de áreas agrícolas del Paso 5.</p></figure>
   </div>
   ${sep}${secH('Tabla de bloques','Ordenable · clic en una fila para ver la ficha')}
   <div class="ux-tw"><table id="e6Tbl"><thead></thead><tbody></tbody></table></div>
   ${sep}${secH('Lectura integrada','Peligro × estado del ecosistema · fuentes de ecosistema · brecha')}
   <div id="e6Int"></div>
   <p class="foot">Fuente: Tablero E6 «Estado de los ecosistemas por bloque de intervención» (ANIN-DIME-SESDI, octubre de 2026): clasificación K-means de 92 tomas S2DR4 de 1 m (Nota Técnica N.° 002-2026-ANIN/DIME/SESDI); MSAVI 2024 por clases (catálogo V6); NDVI mediana 2025 (59 bloques de Morropón); fichas F-DT 2026; calicatas ER (sep-2026); DEM 12.5 m. IC: media de cinco atributos (Gann et al., 2019) frente a referencias del mismo piso. Brecha según R.M. N.° 00213-2024-MINAM: superficie bajo el umbral MSAVI 0.4976. Las especies «esperadas» provienen de la firma espectral por piso y cobertura; no son inventario.</p>`;
  const show=cod=>{const b=EB[cod]; if(!b) return; $('e6Ficha').innerHTML=fichaE6(b); markMap('e6Map',cod); $('e6Sel').value=cod; wireGo($('e6Ficha'));};
  $('e6Sel').value=worst.bloque; wireGo($('e6Ficha'));
  $('e6Sel').addEventListener('change',e=>show(e.target.value));
  $('e6Col').addEventListener('change',e=>{const m=e.target.value; $('e6Map').querySelectorAll('path[data-b]').forEach(p=>p.style.fill=colorE6[m](EB[p.dataset.b])); $('e6Leg').innerHTML=legE6(m);});
  wireMap('e6Map',show);
  sortable('e6Tbl',[
    {h:'Bloque',f:b=>`<b class="mono">${esc(b.bloque)}</b>`,v:b=>b.bloque,l:1,cls:'cod'},
    {h:'Distrito',f:b=>esc(b.distrito),v:b=>b.distrito,l:1,cls:'l'},
    {h:'Ecosistema',f:b=>`<i class="e6-sw" style="background:var(${ECOV(b.ecosistema)})"></i>${esc(ECOS(b.ecosistema))}`,v:b=>b.ecosistema,l:1,cls:'l'},
    {h:'Piso',f:b=>esc(b.piso_dom),v:b=>b.piso_dom},
    {h:'Área (ha)',f:b=>n2(b.area_ha),v:b=>b.area_ha},
    {h:'MSAVI',f:b=>n3(b.msavi_2024),v:b=>b.msavi_2024},
    {h:'Bajo umbral (ha)',f:b=>n2(b.msavi_bajo_umbral_ha),v:b=>b.msavi_bajo_umbral_ha},
    {h:'% bajo umbral',f:b=>mbar(b.msavi_bajo_umbral_pct),v:b=>b.msavi_bajo_umbral_pct},
    {h:'IC',f:b=>n2(b.ic_bloque),v:b=>b.ic_bloque},
    {h:'Clase',f:b=>icPill(b.clase_estado),v:b=>ICN.indexOf(b.clase_estado)},
    {h:'Tip.',f:b=>`<span title="${esc(TIPN[b.tipologia_S2DR4]||'')}">${esc(b.tipologia_S2DR4)}</span>`,v:b=>b.tipologia_S2DR4},
    {h:'PI',f:b=>piPill((XB[b.bloque]||{}).pi),v:b=>['Muy bajo','Bajo','Medio','Alto','Muy alto'].indexOf((XB[b.bloque]||{}).pi)},
    {h:'Rest. activa (ha)',f:b=>n1(b.delim_rest_activa_ha),v:b=>b.delim_rest_activa_ha},
    {h:'Enriq. (ha)',f:b=>n1(b.delim_enriquecimiento_ha),v:b=>b.delim_enriquecimiento_ha},
    {h:'Conserv. (ha)',f:b=>n1(b.delim_conservacion_ha),v:b=>b.delim_conservacion_ha},
    {h:'Agrícola (ha)',f:b=>n1(b.delim_tamizaje_agricola_ha),v:b=>b.delim_tamizaje_agricola_ha},
    {h:'Esp. conf.',f:b=>n0(b.n_especies_confirmadas),v:b=>b.n_especies_confirmadas},
    {h:'Evid.',f:b=>esc(b.clase_evidencia),v:b=>b.clase_evidencia}
  ],B,cod=>{show(cod); $('e6Map').scrollIntoView({behavior:'smooth',block:'center'});});
  ecoInt();
}
function wireGo(el){ if(!el) return; el.querySelectorAll('[data-go]').forEach(bt=>bt.onclick=()=>{const g=bt.dataset.go, b=bt.dataset.b; if(g==='fic') goFicha(b); else if(g==='tam'){goTab('tam'); spShow(b);} else if(g==='eco'){goTab('eco'); const s=$('e6Sel'); if(s){s.value=b; s.dispatchEvent(new Event('change'));}}});}

/* lectura integrada del E6: PI × estado, fuentes de ecosistema y brecha */
const PIK=['Muy alto','Alto','Medio','Bajo','Muy bajo'];
function quad(r){const hi=r.pi==='Alto'||r.pi==='Muy alto', deg=r.cls==='Moderado'||r.cls==='Severo';
  return hi&&deg?['c3','Convergente: peligro alto y ecosistema degradado','Gestión correctiva: restauración activa y obras de conservación de suelos donde la pendiente lo permita.']
   :hi?['c2','Peligro alto con ecosistema leve o no degradado','Gestión prospectiva: conservación, clausura temporal y enriquecimiento para sostener la regulación que el ecosistema todavía presta.']
   :deg?['c2','Ecosistema degradado con peligro medio o bajo','Recuperación del servicio ecosistémico; prioridad por brecha antes que por riesgo.']
   :['c1','Peligro medio o bajo y ecosistema leve o no degradado','Monitoreo y conservación; candidato a reasignar superficie si el tamizaje lo exige.'];}
function ecoInt(){
  const h=$('e6Int'); if(!h) return;
  const piL=PIK.filter(k=>X.some(r=>r.pi===k));
  const cell=(p,c)=>{const L=X.filter(r=>r.pi===p&&r.cls===c); return L.length?`<td class="e6-q ${quad({pi:p,cls:c})[0]}" title="${esc(L.map(r=>r.b).join(', '))}"><b>${L.length}</b><small>${n2(sum(L,r=>r.ha))} ha</small></td>`:'<td class="na">—</td>';};
  const groups={}; X.forEach(r=>{const q=quad(r); (groups[q[1]]=groups[q[1]]||{q,L:[]}).L.push(r);});
  const dif=X.filter(r=>r.v6&&ECOT(r.eco).split(' ')[0].toLowerCase().slice(0,6)!==String(r.v6).split(' ')[0].toLowerCase().slice(0,6));
  const kb=(window.INT&&window.INT.kpi||[]).find(k=>/Brecha/.test(k[0]));
  h.innerHTML=`<div class="ux-grid">
   <figure class="ux-card"><figcaption class="ux-h"><h3>Peligro Integrado × clase de estado del ecosistema</h3></figcaption>
    <p class="ux-d">Número de bloques y superficie en cada cruce. El análisis de riesgo de la UP (el ecosistema) distingue el peligro (PI del bloque, MCA–AHP) de la fragilidad que expresa su estado (IC del E6).</p>
    <div class="ux-tw" style="max-height:none"><table class="e6-qt"><thead><tr><th class="l" scope="col">PI del bloque</th>${ICN.map(c=>`<th scope="col">${c}</th>`).join('')}</tr></thead>
    <tbody>${piL.map(p=>`<tr><td class="l">${piPill(p)}</td>${ICN.map(c=>cell(p,c)).join('')}</tr>`).join('')}</tbody></table></div>
    <div class="e6-ql">${Object.values(groups).sort((a,b)=>b.q[0].localeCompare(a.q[0])).map(({q,L})=>`<div class="find ${q[0]==='c3'?'cri':q[0]==='c2'?'adv':'ok'}"><i></i><div class="body"><h3>${esc(q[1])} · ${L.length} bloques · ${n2(sum(L,r=>r.ha))} ha</h3><p>${esc(q[2])}</p><p class="src">${L.sort((a,b)=>sortCod(a.b,b.b)).map(r=>`<button type="button" class="e6-chip" data-eb="${esc(r.b)}">${esc(r.b)}</button>`).join(' ')}</p></div></div>`).join('')}</div>
    <p class="ux-n">Lectura de este volumen para ordenar las MRR-CCC (prospectiva, correctiva); no es un criterio normativo.</p></figure>
   <figure class="ux-card"><figcaption class="ux-h"><h3>Tres fuentes de ecosistema y estado</h3></figcaption>
    <p class="ux-d">El E6 usa el ecosistema de la ficha F-DT (campo); «Geoespacial» usa el Mapa Nacional de Ecosistemas (MINAM 2018), de escala regional. ${dif.length?`Difieren en <b>${dif.length}</b> de ${X.length} bloques:`:'Coinciden en todos los bloques.'}</p>
    ${dif.length?`<div class="ux-tw" style="max-height:300px"><table><thead><tr><th class="l" scope="col">Bloque</th><th class="l" scope="col">E6 (F-DT)</th><th class="l" scope="col">MINAM 2018 dominante</th><th scope="col">%</th></tr></thead><tbody>${dif.sort((a,b)=>sortCod(a.b,b.b)).map(r=>`<tr data-cod="${esc(r.b)}"><td class="cod">${esc(r.b)}</td><td class="l">${esc(ECOT(r.eco))}</td><td class="l">${esc(r.v6)}</td><td>${n1(r.v6p)}</td></tr>`).join('')}</tbody></table></div>`:''}
    <p class="ux-d" style="margin-top:6px">La «zona agrícola» del MINAM corresponde, en la mayoría de esos bloques, a mosaicos con matorral o bosque seco en ladera que la F-DT y el S2DR4 de 1 m sí resuelven. La superficie de cultivo real del bloque es el «área agrícola a tamizar» del E6 (${n2(DL.ag)} ha en la provincia), que se excluye en el Paso 5.</p>
    <h4 class="e6-sh">Brecha: un solo criterio para los 117 bloques</h4>
    <p class="ux-d">El E6 calcula la superficie bajo el umbral MSAVI 0.4976 por distribución areal (clases DN) en los 117 bloques: <b>${n2(BU)} ha</b> en ${esc(PROV)} (${n1(BU/A*100)} %). ${kb?`La cinta de indicadores de este volumen reporta ${esc(kb[1])} ${esc(kb[2])}.`:''} ${S.brTxt||''}</p></figure>
  </div>`;
  h.querySelectorAll('[data-eb]').forEach(bt=>bt.onclick=()=>{const s=$('e6Sel'); s.value=bt.dataset.eb; s.dispatchEvent(new Event('change')); $('e6Map').scrollIntoView({behavior:'smooth',block:'center'});});
  h.querySelectorAll('tr[data-cod]').forEach(tr=>{tr.style.cursor='pointer'; tr.onclick=()=>{const s=$('e6Sel'); s.value=tr.dataset.cod; s.dispatchEvent(new Event('change')); $('e6Map').scrollIntoView({behavior:'smooth',block:'center'});};});
}

/* =====================================================================
   B. PESTAÑA «TAMIZAJE PREDIAL»
   ===================================================================== */
const TA=sum(X,r=>r.ha), TCC=sum(X,r=>r.cc), TCOF=sum(X,r=>r.cof), TEST=sum(X,r=>r.est), NPC=sum(X,r=>r.npc);
const CCN=new Map(); X.forEach(r=>String(r.ccn||'').split(';').map(s=>s.trim()).filter(Boolean).forEach(n=>{const o=CCN.get(n)||{b:[],ha:0}; o.b.push(r.b); o.ha+=r.cc; CCN.set(n,o);}));
const FL=r=>r.f||[];
const P19=X.filter(r=>(r.npc||0)>19), TEN=X.filter(r=>FL(r).some(f=>f[0]==='ten')), P5D=X.filter(r=>FL(r).some(f=>f[0]==='p5'&&/F-DS-01/.test(f[1]))), NOS=X.filter(r=>!r.nf);
const colorSP={sit:r=>`var(${SITV[r.sit]||'--ux-na'})`, npc:r=>{const n=r.npc||0; return n>19?'var(--e6-sev)':n>=10?'var(--e6-mod)':n>0?'var(--e6-leve)':'var(--ux-no)';},
  fl:r=>{const k=FL(r).filter(f=>f[0]!=='soc').length; return k>=2?'var(--e6-sev)':k===1?'var(--e6-mod)':'var(--e6-ok)';}, nf:r=>r.nf?'var(--ux-yes)':'var(--ux-no)'};
function legSP(m){
  if(m==='sit') return SITN.filter(s=>X.some(r=>r.sit===s)).map(s=>`<span><i style="background:var(${SITV[s]})"></i>${esc(s)}</span>`).join('');
  if(m==='npc') return `<span><i style="background:var(--ux-no)"></i>Sin predios COFOPRI</span><span><i style="background:var(--e6-leve)"></i>1 a 9</span><span><i style="background:var(--e6-mod)"></i>10 a 19</span><span><i style="background:var(--e6-sev)"></i>Más de 19 (supera el Paso 5)</span>`;
  if(m==='fl') return `<span><i style="background:var(--e6-ok)"></i>Sin alertas de gabinete</span><span><i style="background:var(--e6-mod)"></i>1 alerta</span><span><i style="background:var(--e6-sev)"></i>2 o más alertas</span>`;
  return `<span><i style="background:var(--ux-yes)"></i>Con ficha social (17/10)</span><span><i style="background:var(--ux-no)"></i>Sin ficha social</span>`;
}
function spCard(cod){
  const r=XB[cod]; if(!r) return '';
  const im=imgOf(r);
  const al=FL(r);
  return `<div class="e6-fh"><b class="mono">${esc(r.sp)} · Bloque ${esc(r.b)}</b>${sitPill(r.sit)}<span class="mono sm">${esc(r.asis)} · ${esc(r.esc)}</span></div>
   ${im?`<a href="${im}" target="_blank" rel="noopener" title="Abrir la lámina en tamaño completo"><div class="ux-map"><img src="${im}" alt="Mapa de situación predial ${esc(r.sp)} del bloque ${esc(r.b)}" loading="lazy"></div></a>`
     :`<div class="warnbox"><b>Lámina ${esc(r.sp)} no disponible.</b> No figura entre los archivos extraídos de la carpeta compartida (Parte 2, Morropón); solo está dentro del ZIP «Mapas_SP_v2_Parte2_Morropon_SP059-117». Los datos del índice sí se muestran.</div>`}
   <dl class="ux-kv"><dt>Superficie del bloque</dt><dd>${n3(r.ha)} ha · ${esc(r.d)}</dd>
    <dt>Comunidad campesina</dt><dd>${n3(r.cc)} ha · ${n1(r.pcc)} %</dd>
    <dt>Predios COFOPRI fuera de C. C.</dt><dd>${n3(r.cof)} ha · ${n1(r.pcof)} % · ${n0(r.npc)} predios</dd>
    <dt>Estado y/o posesionarios</dt><dd>${n3(r.est)} ha · ${n1(r.pest)} %</dd>
    <dt class="w">Comunidad(es) campesina(s)</dt><dd class="w">${esc(r.ccn||'—')}</dd>
    <dt class="w">Centros poblados asociados (${n0(r.nccpp)} · ${n0(r.pob)} hab.)</dt><dd class="w">${esc(r.ccpp||'—')}</dd></dl>
   ${stack([['C. C.',r.pcc,'--sp-cc'],['Predios COFOPRI',r.pcof,'--sp-pr'],['Estado y/o posesionarios',r.pest,'--sp-est']],'Composición predial del bloque')}
   <h4 class="e6-sh">Lo que declara la F-DS-01 (consolidado 17/10)</h4>
   <dl class="ux-kv"><dt>Fichas sociales del bloque</dt><dd>${n0(r.nf)}</dd><dt>Régimen predominante</dt><dd>${esc(r.regd||'sin dato')}</dd><dt>Se superpone a tierras comunales</dt><dd>${esc(r.supcc||'sin dato')}</dd><dt>Predios individuales (aprox.)</dt><dd>${r.npd!=null?n0(r.npd):'sin dato'}</dd><dt>Tierras tituladas</dt><dd>${r.tit!=null?n0(r.tit)+' %':'sin dato'}</dd><dt>Conflictos de linderos</dt><dd>${esc(r.lind||'sin dato')}</dd></dl>
   <h4 class="e6-sh">Alertas del cruce</h4>
   ${al.length?`<ul class="e6-al">${al.map(([k,t])=>`<li class="${k}">${esc(t)}</li>`).join('')}</ul>`:'<p class="ux-d">Sin alertas: la tenencia declarada coincide con el gabinete y el bloque no supera los 19 predios.</p>'}
   <div class="e6-go"><button type="button" class="btn" data-go="eco" data-b="${esc(r.b)}">Estado del ecosistema</button><button type="button" class="btn" data-go="fic" data-b="${esc(r.b)}">Ficha por bloque</button></div>`;
}
let spShow=()=>{};
function tamTab(){
  const root=$('spRoot'); if(!root) return;
  const bySit={}; X.forEach(r=>{const o=bySit[r.sit]=bySit[r.sit]||{n:0,ha:0,cc:0,cof:0,est:0,npc:0}; o.n++; o.ha+=r.ha; o.cc+=r.cc; o.cof+=r.cof; o.est+=r.est; o.npc+=r.npc||0;});
  const byAp={}; X.forEach(r=>{const o=byAp[r.asis]=byAp[r.asis]||{n:0,ha:0,d:new Set(),cc:0}; o.n++; o.ha+=r.ha; o.d.add(r.d); o.cc+=r.cc;});
  const first=[...X].sort((a,b)=>FL(b).filter(f=>f[0]!=='soc').length-FL(a).filter(f=>f[0]!=='soc').length||sortCod(a.b,b.b))[0];
  root.innerHTML=`
   <dl class="sx-kpis">
    <div class="hl"><dt>En comunidad campesina</dt><dd>${n2(TCC)} ha <small>${n1(TCC/TA*100)} % de ${n2(TA)} ha · ${CCN.size} comunidades distintas</small></dd></div>
    <div><dt>Tierras del Estado y/o posesionarios</dt><dd>${n2(TEST)} ha <small>${n1(TEST/TA*100)} % · sin registro en GEORURAL ni COFOPRI</small></dd></div>
    <div><dt>Predios COFOPRI fuera de C. C.</dt><dd>${n2(TCOF)} ha <small>${n0(NPC)} predios interceptados</small></dd></div>
    <div><dt>Bloques con más de 19 predios</dt><dd>${P19.length} <small>${P19.length?P19.map(r=>r.b).join(', ')+' · ':''}máximo del Paso 5</small></dd></div>
    <div><dt>Tenencia declarada ≠ gabinete</dt><dd>${TEN.length} <small>bloques con discrepancia F-DS-01 / mapa predial</small></dd></div>
    <div><dt>Sin ficha social</dt><dd>${NOS.length} <small>bloques sin titulares verificados con la población</small></dd></div></dl>
   <div class="e6-two">
    <figure class="ux-card"><figcaption class="ux-h"><h3>Mapa de situación predial</h3></figcaption>
     <div class="ux-ctrl"><label for="spCol">Colorear por</label><select id="spCol"><option value="sit">Situación predial (gabinete)</option><option value="npc">Predios COFOPRI interceptados</option><option value="fl">Alertas del cruce con la F-DS-01</option><option value="nf">Ficha social al 17/10</option></select>
      <label for="spSel">Lámina</label><select id="spSel">${[...X].sort((a,b)=>a.sp.localeCompare(b.sp)).map(r=>`<option value="${esc(r.b)}">${esc(r.sp)} · bloque ${esc(r.b)}</option>`).join('')}</select></div>
     <div class="e6-mapbox">${polyMap('spMap',b=>colorSP.sit(XB[b.bloque]),{sel:first.b,title:'Bloques coloreados por situación predial'})}</div>
     <div class="ux-leg" id="spLeg">${legSP('sit')}</div>
     <p class="ux-n">Clic o Enter en un bloque para abrir su lámina A3. Las láminas son de gabinete: «Tierras del Estado y/o posesionarios» es la superficie sin registro en las capas de comunidades campesinas (GEORURAL–MIDAGRI) y de predios rurales (COFOPRI), con tolerancia cartográfica de 5 %.</p></figure>
    <figure class="ux-card e6-ficha" id="spCard" aria-live="polite">${spCard(first.b)}</figure>
   </div>
   ${sep}${secH('Composición predial','Situación predial, asistentes prediales y comunidades campesinas')}
   <div class="ux-grid">
    <figure class="ux-card wide"><figcaption class="ux-h"><h3>Por situación predial</h3></figcaption>
     <div class="ux-tw" style="max-height:none"><table><thead><tr><th class="l" scope="col">Situación</th><th scope="col">Bloques</th><th scope="col">Superficie (ha)</th><th scope="col">% del ámbito</th><th scope="col">C. C. (ha)</th><th scope="col">Estado y/o pos. (ha)</th><th scope="col">Predios COFOPRI</th></tr></thead>
     <tbody>${SITN.filter(s=>bySit[s]).map(s=>{const o=bySit[s]; return `<tr><td class="l">${sitPill(s)}</td><td>${o.n}</td><td>${n2(o.ha)}</td><td>${mbar(o.ha/TA*100)}</td><td>${n2(o.cc)}</td><td>${n2(o.est)}</td><td>${n0(o.npc)}</td></tr>`;}).join('')}</tbody>
     <tfoot><tr><td class="l">Total</td><td>${X.length}</td><td>${n2(TA)}</td><td>100.0 %</td><td>${n2(TCC)}</td><td>${n2(TEST)}</td><td>${n0(NPC)}</td></tr></tfoot></table></div>
     <p class="ux-n">Define el tipo de acta: comunal (asamblea de la C. C.), privada o en posesión (titular y constancia) y estatal. «Parcialmente» indica que el bloque combina ese régimen con Tierras del Estado y/o posesionarios.</p></figure>
    <figure class="ux-card wide"><figcaption class="ux-h"><h3>Por asistente predial y social</h3></figcaption>
     <div class="ux-tw" style="max-height:none"><table><thead><tr><th class="l" scope="col">Asistente</th><th class="l" scope="col">Distritos</th><th scope="col">Bloques</th><th scope="col">Superficie (ha)</th><th scope="col">% en C. C.</th></tr></thead>
     <tbody>${Object.entries(byAp).sort((a,b)=>a[0].localeCompare(b[0])).map(([k,o])=>`<tr><td class="l mono">${esc(k)}</td><td class="l">${esc([...o.d].sort().join(', '))}</td><td>${o.n}</td><td>${n2(o.ha)}</td><td>${mbar(o.cc/o.ha*100)}</td></tr>`).join('')}</tbody></table></div>
     <h4 class="e6-sh">Comunidades campesinas en los bloques</h4>
     <div class="ux-tw" style="max-height:260px"><table><thead><tr><th class="l" scope="col">Comunidad campesina</th><th class="l" scope="col">Bloques</th><th scope="col">Superficie comunal de esos bloques (ha)</th></tr></thead>
     <tbody>${[...CCN.entries()].sort((a,b)=>b[1].ha-a[1].ha).map(([n,o])=>`<tr><td class="l">${esc(n)}</td><td class="l mono">${esc(o.b.sort(sortCod).join(', '))}</td><td>${n2(o.ha)}</td></tr>`).join('')||'<tr><td colspan="3">Sin comunidades campesinas registradas</td></tr>'}</tbody></table></div>
     <p class="ux-n">Si un bloque toca dos comunidades, su superficie comunal se cuenta en ambas: la columna no es sumable. La asignación de asistentes es la del catálogo de liberación de áreas («PENDIENTE»: aún sin asistente).</p></figure>
   </div>
   ${sep}${secH('Cruce con el diagnóstico social y con el Paso 5','F-DS-01 del 17/10 frente al gabinete predial')}
   <p class="lede">El Paso 5 de la selección de bloques admite como máximo 19 predios rurales interceptados por bloque y excluye las áreas agrícolas. El gabinete cuenta los predios COFOPRI; la F-DS-01 recoge el régimen, la superposición con tierras comunales y el número aproximado de predios que declara la población. Ninguna fuente reemplaza la verificación de campo: donde difieren, el acta no debe suscribirse hasta aclarar quién es el titular.</p>
   <div class="finds" id="spFinds"></div>
   ${sep}${secH('Índice de láminas','Ordenable · clic en una fila para ver la lámina')}
   <div class="ux-tw"><table id="spTbl"><thead></thead><tbody></tbody></table></div>
   <p class="foot">Fuente: Índice de mapas de situación predial de los 117 bloques V6 v2 (A3, UTM WGS 84 Zona 17S), Proyecto IN Piura, Entregable 2 · Subentregable 2.2 «Expediente preliminar de liberación de áreas» (ANIN-DIME-SESDI, 11/10/2026), sobre las capas de comunidades campesinas GEORURAL–MIDAGRI y de predios rurales COFOPRI, con centros poblados (matriz CCPP-INEI v5) y red vial departamental y vecinal; información de gabinete. F-DS-01: consolidado DS del aplicativo IN Piura, 17/10/2026. Delimitación agrícola: Tablero E6.</p>`;
  spShow=cod=>{if(!XB[cod]) return; $('spCard').innerHTML=spCard(cod); markMap('spMap',cod); $('spSel').value=cod; wireGo($('spCard'));};
  $('spSel').value=first.b; wireGo($('spCard'));
  $('spSel').addEventListener('change',e=>spShow(e.target.value));
  $('spCol').addEventListener('change',e=>{const m=e.target.value; $('spMap').querySelectorAll('path[data-b]').forEach(p=>p.style.fill=colorSP[m](XB[p.dataset.b])); $('spLeg').innerHTML=legSP(m);});
  wireMap('spMap',spShow);
  // hallazgos del cruce
  const supNo=X.filter(r=>r.supcc==='No'&&r.pcc>=5), regC=X.filter(r=>/^Comunal/.test(r.regd||'')&&r.pcc<5), regE=X.filter(r=>/^Estatal/.test(r.regd||'')&&r.pcc>=50), regI=X.filter(r=>/Individual/.test(r.regd||'')&&r.pcc>=50), supSi=X.filter(r=>r.supcc==='Sí'&&r.pcc<5);
  const L=a=>a.sort((x,y)=>sortCod(x.b,y.b)).map(r=>`<button type="button" class="e6-chip" data-spb="${esc(r.b)}">${esc(r.b)}</button>`).join(' ');
  const F=[];
  if(P19.length) F.push(['cri',`${pl(P19.length,'bloque supera','bloques superan')} los 19 predios COFOPRI interceptados`,`El gabinete cuenta ${P19.map(r=>`${r.b} (${r.npc})`).join(', ')}. Deben sectorizarse o redelimitarse antes de la convocatoria, o justificarse la excepción en el expediente.`,L(P19)]);
  if(TEN.length) F.push(['cri',`La tenencia que declara la F-DS-01 no coincide con el gabinete en ${pl(TEN.length,'bloque','bloques')}`,[
     supNo.length?`${pl(supNo.length,'bloque declara','bloques declaran')} que <b>no</b> se superpone${supNo.length>1?'n':''} a tierras comunales y el gabinete ${supNo.length>1?'los':'lo'} ubica en comunidad campesina (${supNo.map(r=>r.b).join(', ')}).`:'',
     regC.length?`${pl(regC.length,'se declara comunal','se declaran comunales')} y el gabinete no registra comunidad (${regC.map(r=>r.b).join(', ')}).`:'',
     regE.length?`${pl(regE.length,'se declara estatal y está','se declaran estatales y están')} en C. C. (${regE.map(r=>r.b).join(', ')}).`:'',
     regI.length?`${pl(regI.length,'se declara de régimen individual y está','se declaran de régimen individual y están')} en C. C. (${regI.map(r=>r.b).join(', ')}).`:'',
     supSi.length?`${pl(supSi.length,'declara','declaran')} superposición con tierras comunales donde el gabinete no la registra (${supSi.map(r=>r.b).join(', ')}).`:''].filter(Boolean).join(' ')+' Puede tratarse de comunidades sin georreferenciar, de posesionarios dentro de la comunidad o de un error de la ficha: se resuelve en el tamizaje de campo con la junta directiva comunal.',L(TEN)]);
  if(P5D.length) F.push(['adv',`${pl(P5D.length,'F-DS-01 declara','F-DS-01 declaran')} más de 19 predios individuales`,`${P5D.map(r=>`${r.b} (≈ ${n0(r.npd)})`).join(', ')}. Son estimaciones de la población, no un conteo catastral: indican posesionarios sin título que el gabinete no ve (el gabinete cuenta ${n0(NPC)} predios COFOPRI en toda la provincia).`,L(P5D)]);
  if(NOS.length) F.push(['adv',`${pl(NOS.length,'bloque','bloques')} sin ficha social`,`En ${NOS.map(r=>r.b).join(', ')} la situación predial solo se conoce por gabinete; el tamizaje de campo será la primera fuente de titularidad.`,L(NOS)]);
  const ag=sum(B,b=>b.delim_tamizaje_agricola_ha);
  F.push(['',`Áreas agrícolas a excluir: ${n2(ag)} ha según el E6`,`La delimitación interna del E6 identifica ${n2(ag)} ha agrícolas dentro de los bloques (${n1(ag/TA*100)} % del ámbito). Su exclusión (Paso 5) reduce la superficie a liberar y el número de titulares a convocar; la cifra debe contrastarse con la lámina de cada bloque.`,'']);
  $('spFinds').innerHTML=F.map(([c,t,p,l])=>`<div class="find ${c}"><i></i><div class="body"><h3>${esc(t)}</h3><p>${p}</p>${l?`<p class="src">${l}</p>`:''}</div></div>`).join('');
  $('spFinds').querySelectorAll('[data-spb]').forEach(bt=>bt.onclick=()=>{spShow(bt.dataset.spb); $('spMap').scrollIntoView({behavior:'smooth',block:'center'});});
  sortable('spTbl',[
    {h:'Lámina',f:r=>`<span class="mono">${esc(r.sp)}</span>${r.noimg?' <span class="ux-flag m">sin lámina</span>':''}`,v:r=>r.sp,l:1,cls:'l'},
    {h:'Bloque',f:r=>`<b class="mono">${esc(r.b)}</b>`,v:r=>r.b,l:1,cls:'cod'},
    {h:'Distrito',f:r=>esc(r.d),v:r=>r.d,l:1,cls:'l'},
    {h:'Asist.',f:r=>esc(r.asis),v:r=>r.asis},
    {h:'Sup. (ha)',f:r=>n2(r.ha),v:r=>r.ha},
    {h:'Composición',f:r=>stack([['C. C.',r.pcc,'--sp-cc'],['Predios COFOPRI',r.pcof,'--sp-pr'],['Estado y/o posesionarios',r.pest,'--sp-est']],'Composición predial'),v:r=>r.pcc,l:1,cls:'l'},
    {h:'Situación',f:r=>sitPill(r.sit),v:r=>SITN.indexOf(r.sit),l:1,cls:'l'},
    {h:'Predios COFOPRI',f:r=>r.npc>19?`<b>${n0(r.npc)}</b> <span class="ux-flag h">&gt; 19</span>`:n0(r.npc),v:r=>r.npc||0},
    {h:'F-DS-01: régimen',f:r=>esc(r.regd||'—'),v:r=>r.regd||'',l:1,cls:'l'},
    {h:'Fichas DS',f:r=>n0(r.nf),v:r=>r.nf},
    {h:'Alertas',f:r=>{const k=FL(r).filter(f=>f[0]!=='soc').length; return k?`<span class="ux-flag ${k>1?'h':'m'}">${k}</span>`:'<span class="ux-flag l">0</span>';},v:r=>FL(r).filter(f=>f[0]!=='soc').length},
    {h:'CCPP',f:r=>n0(r.nccpp),v:r=>r.nccpp},
    {h:'Pob. CCPP',f:r=>n0(r.pob),v:r=>r.pob}
  ],X,cod=>{spShow(cod); $('spMap').scrollIntoView({behavior:'smooth',block:'center'});});
}

/* =====================================================================
   C. INTEGRACIÓN EN LAS PESTAÑAS EXISTENTES
   ===================================================================== */
function addBlock(sec,html,{first=false}={}){
  const s=$(sec); if(!s) return null;
  const d=document.createElement('div'); d.className='e6-block'; d.innerHTML=html;
  if(first){const h=s.querySelector('.lede')||s.querySelector('.sec-h'); if(h){h.after(d); return d;}}
  s.appendChild(d); return d;
}
/* C1. Ficha por bloque: ecosistema E6 y situación predial */
function ficha(){
  const f=$('ficha'); if(!f) return;
  const add=()=>{const h=f.querySelector('.top h3'); const ux=f.querySelector('.ux-fb');
    const b=(ux&&ux.dataset.b)||(h&&h.textContent.trim().replace(/^Bloque\s+/,'')); if(!b||!EB[b]) return;
    const old=f.querySelector('.e6-fb'); if(old&&old.dataset.b===b) return; if(old) old.remove();
    const e=EB[b], r=XB[b]||{}, im=imgOf(r), al=FL(r).filter(x=>x[0]!=='soc');
    const d=document.createElement('div'); d.className='ux-fb e6-fb'; d.dataset.b=b;
    d.innerHTML=`<h4>Estado del ecosistema (E6) y situación predial <small>Tablero E6 · mapas de situación predial v2 (11/10/2026) · F-DS-01 al 17/10/2026</small></h4>
     <div class="ux-fbg">
      <div style="display:grid;gap:10px;min-width:0">
       <div class="e6-fh">${icPill(e.clase_estado)}<span class="mono sm">IC ${n2(e.ic_bloque)}</span>${piPill(r.pi)}<span class="mono sm">tipología ${esc(e.tipologia_S2DR4)}</span></div>
       <dl class="ux-kv"><dt>Ecosistema (F-DT)</dt><dd>${esc(ECOT(e.ecosistema))}</dd><dt>Bajo el umbral MSAVI</dt><dd>${n2(e.msavi_bajo_umbral_ha)} ha · ${n1(e.msavi_bajo_umbral_pct)} %</dd>
        <dt>Restauración activa · enriquecimiento</dt><dd>${n1(e.delim_rest_activa_ha)} · ${n1(e.delim_enriquecimiento_ha)} ha</dd><dt>Conservación · agrícola a tamizar</dt><dd>${n1(e.delim_conservacion_ha)} · ${n1(e.delim_tamizaje_agricola_ha)} ha</dd>
        <dt>Especies confirmadas</dt><dd>${n0(e.n_especies_confirmadas)} de ${n0(e.n_especies_esperadas)} esperadas</dd></dl>
       <p class="ux-d"><b>${esc(quad({pi:r.pi,cls:e.clase_estado})[1])}.</b> ${esc(quad({pi:r.pi,cls:e.clase_estado})[2])}</p>
       <div class="e6-go"><button type="button" class="btn" data-go="eco" data-b="${esc(b)}">Ecosistemas</button><button type="button" class="btn" data-go="tam" data-b="${esc(b)}">Tamizaje predial</button></div>
      </div>
      <div style="display:grid;gap:10px;min-width:0">
       ${im?`<a href="${im}" target="_blank" rel="noopener" title="Abrir la lámina ${esc(r.sp)}"><div class="ux-map"><img src="${im}" alt="Mapa de situación predial del bloque ${esc(b)}" loading="lazy"></div></a>`:`<div class="warnbox"><b>Lámina ${esc(r.sp||'')} no disponible</b> en la carpeta compartida.</div>`}
       <div class="e6-fh">${sitPill(r.sit)}<span class="mono sm">${esc(r.sp)} · ${esc(r.asis)}</span></div>
       ${stack([['C. C.',r.pcc,'--sp-cc'],['Predios COFOPRI',r.pcof,'--sp-pr'],['Estado y/o posesionarios',r.pest,'--sp-est']],'Composición predial')}
       <dl class="ux-kv"><dt>C. C. · COFOPRI · Estado</dt><dd>${n1(r.pcc)} · ${n1(r.pcof)} · ${n1(r.pest)} %</dd><dt>Predios COFOPRI</dt><dd>${n0(r.npc)}</dd><dt>F-DS-01: régimen</dt><dd>${esc(r.regd||'sin ficha o sin dato')}</dd></dl>
       ${al.length?`<ul class="e6-al">${al.map(([k,t])=>`<li class="${k}">${esc(t)}</li>`).join('')}</ul>`:''}
      </div></div>`;
    const su=f.querySelector('.sue-fb'); if(su) su.before(d); else f.appendChild(d);
    wireGo(d);
  };
  new MutationObserver(()=>setTimeout(add,0)).observe(f,{childList:true,subtree:false}); setTimeout(add,0);
}
/* C2. Integralidad: matriz territorial–social–predial por bloque */
function integ(){
  const rows=[...X].sort((a,b)=>sortCod(a.b,b.b));
  const nq={}; rows.forEach(r=>{const q=quad(r)[0]; nq[q]=(nq[q]||0)+1;});
  const ready=rows.filter(r=>!FL(r).length), verif=rows.filter(r=>FL(r).length);
  addBlock('p-int',`${sep}${secH('Ecosistema, peligro y tierra: la integralidad por bloque','E6 · PI · F-DS-01 (17/10) · situación predial v2')}
   <p class="lede">La unidad productora es el ecosistema; el peligro se origina en él y la intervención solo es viable donde la tierra está disponible. Esta matriz reúne, por bloque, el Peligro Integrado (MCA–AHP), el estado del ecosistema del E6, el tratamiento que propone su delimitación interna, la situación predial de gabinete y lo que declara la población en la F-DS-01.</p>
   <dl class="sx-kpis"><div class="hl"><dt>Convergentes</dt><dd>${nq.c3||0} <small>PI alto y ecosistema moderado o severo</small></dd></div>
    <div><dt>Gestión prospectiva o recuperación</dt><dd>${nq.c2||0} <small>peligro alto sin degradación marcada, o degradación con peligro menor</small></dd></div>
    <div><dt>Sin alertas de tenencia ni de predios</dt><dd>${ready.length} <small>de ${rows.length} bloques; ${verif.length} requieren verificación</small></dd></div>
    <div><dt>Brecha E6 (areal)</dt><dd>${n2(BU)} ha <small>${n1(BU/A*100)} % · criterio único de los 117 bloques</small></dd></div></dl>
   <div class="ux-tw"><table id="e6IntT"><thead></thead><tbody></tbody></table></div>
   <p class="foot">Tratamiento dominante: el mayor de los cuatro sectores de la delimitación interna del E6 (RA restauración activa, EN enriquecimiento, CO conservación, AG agrícola a tamizar). Alertas: discrepancias de tenencia entre la F-DS-01 y el gabinete y bloques sobre 19 predios (Paso 5). Lectura de este volumen, no un criterio normativo.</p>`);
  const TR=r=>{const e=EB[r.b]; const v=[['RA',e.delim_rest_activa_ha],['EN',e.delim_enriquecimiento_ha],['CO',e.delim_conservacion_ha],['AG',e.delim_tamizaje_agricola_ha]].sort((a,b)=>b[1]-a[1])[0]; return v;};
  sortable('e6IntT',[
    {h:'Bloque',f:r=>`<b class="mono">${esc(r.b)}</b>`,v:r=>r.b,l:1,cls:'cod'},
    {h:'Distrito',f:r=>esc(r.d),v:r=>r.d,l:1,cls:'l'},
    {h:'Sup. (ha)',f:r=>n2(r.ha),v:r=>r.ha},
    {h:'PI',f:r=>piPill(r.pi),v:r=>['Muy bajo','Bajo','Medio','Alto','Muy alto'].indexOf(r.pi)},
    {h:'Estado E6',f:r=>icPill(r.cls),v:r=>ICN.indexOf(r.cls)},
    {h:'% bajo umbral',f:r=>mbar(r.bu),v:r=>r.bu},
    {h:'Lectura',f:r=>{const q=quad(r); return `<span class="convcls ${q[0]}">${esc(q[1].split(':')[0])}</span>`;},v:r=>quad(r)[0],l:1,cls:'l'},
    {h:'Tratamiento dominante',f:r=>{const t=TR(r); return `<span class="mono">${t[0]}</span> ${n1(t[1])} ha`;},v:r=>TR(r)[0]},
    {h:'Situación predial',f:r=>sitPill(r.sit),v:r=>SITN.indexOf(r.sit),l:1,cls:'l'},
    {h:'F-DS-01',f:r=>r.nf?`${n0(r.nf)} · ${esc((r.regd||'sin régimen').replace(/\s*\(.*\)/,''))}`:'<span class="ux-flag m">sin ficha</span>',v:r=>r.nf,l:1,cls:'l'},
    {h:'Alertas',f:r=>{const k=FL(r).filter(f=>f[0]!=='soc'); return k.length?`<span class="ux-flag ${k.length>1?'h':'m'}" title="${esc(k.map(x=>x[1]).join(' · '))}">${k.length}</span>`:'<span class="ux-flag l">0</span>';},v:r=>FL(r).filter(f=>f[0]!=='soc').length}
  ],rows,cod=>goFicha(cod));
}
/* C3. Consistencia: fuentes de ecosistema y de tenencia */
function consist(){
  const dif=X.filter(r=>r.v6&&ECOT(r.eco).split(' ')[0].toLowerCase().slice(0,6)!==String(r.v6).split(' ')[0].toLowerCase().slice(0,6));
  addBlock('p-con',`${sep}${secH('Consistencia entre el E6, el MINAM 2018, la F-DS-01 y el gabinete predial','Corte 17/10/2026')}
   <div class="finds">
    <div class="find adv"><i></i><div class="body"><h3>Ecosistema: E6 (F-DT) frente al Mapa Nacional de Ecosistemas en ${dif.length} de ${X.length} bloques</h3><p>${dif.length?dif.sort((a,b)=>sortCod(a.b,b.b)).map(r=>`<b>${esc(r.b)}</b>: ${esc(ECOT(r.eco))} / ${esc(r.v6)}`).join(' · '):'Sin diferencias.'}</p><p class="src">Se adopta el ecosistema de la F-DT (escala de bloque); el MINAM 2018 se mantiene como referencia regional. Ver «Ecosistemas» → «Lectura integrada».</p></div></div>
    <div class="find ${TEN.length?'cri':'ok'}"><i></i><div class="body"><h3>Tenencia: F-DS-01 frente al mapa de situación predial en ${TEN.length} de ${X.length} bloques</h3><p>${TEN.sort((a,b)=>sortCod(a.b,b.b)).map(r=>`<b>${esc(r.b)}</b>: ${esc((r.regd||'').replace(/\s*\(.*\)/,''))} / ${esc(r.sit)}`).join(' · ')||'Sin diferencias.'}</p><p class="src">Ninguna fuente se corrige: se resuelven en el tamizaje de campo. Ver «Tamizaje predial».</p></div></div>
    <div class="find adv"><i></i><div class="body"><h3>Brecha MSAVI: criterio areal del E6</h3><p>El E6 calcula ${n2(BU)} ha bajo el umbral (${n1(BU/A*100)} %) por distribución areal en los ${X.length} bloques, con el mismo criterio en las tres provincias. ${S.brTxt||''}</p></div></div>
   </div>`);
}
/* C4. Geoespacial y Riesgos y oportunidades: enlaces a las nuevas pestañas */
function links(){
  addBlock('p-geo',`<div class="warnbox" style="margin-top:16px"><b>Estado del ecosistema por bloque.</b> El Mapa Nacional de Ecosistemas es de escala regional; el ecosistema de campo (F-DT), el índice de condición y la brecha areal de cada bloque están en la pestaña <button type="button" class="e6-chip" data-gt="eco">Ecosistemas</button>, y la situación predial (comunidades campesinas, predios COFOPRI y Tierras del Estado) en <button type="button" class="e6-chip" data-gt="tam">Tamizaje predial</button>.</div>`);
  addBlock('p-rie',`${sep}<div class="warnbox"><b>Tenencia declarada frente al gabinete predial.</b> ${pl(TEN.length,'bloque','bloques')} con discrepancia entre la F-DS-01 y el mapa de situación predial v2, y ${P19.length} con más de 19 predios COFOPRI (Paso 5). Son el principal riesgo social para la suscripción de actas de libre disponibilidad: <button type="button" class="e6-chip" data-gt="tam">ver Tamizaje predial</button>.</div>`);
  document.querySelectorAll('[data-gt]').forEach(b=>b.onclick=()=>goTab(b.dataset.gt));
}

function run(){[ecoTab,tamTab,ficha,integ,consist,links].forEach(fn=>{try{fn();}catch(e){console.error('E6/SP '+fn.name,e);}});}
if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',run); else run();
})();
</script>
