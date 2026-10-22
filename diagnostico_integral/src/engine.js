/* Motor de costos paramétrico IN Piura — réplica de Costos_Parametros_IN_Piura_v1.3.xlsx
   (hojas BLOQUES → PAQUETES → CALCULO ← CU_COMP → OYM → FLUJO). D = datos extraídos del libro. */
function cpEngine(D){
  const P=D.par;
  const lerp=(a,j)=>a[0]+(a[1]-a[0])*(j-100)/20;          // CU lineal en el jornal (validado 80/100/120)
  function cuAt(j){const o={};for(const k in D.cu){const c=D.cu[k];o[k]={h:lerp(c.h,j),nc:lerp(c.nc,j),c:lerp(c.c,j)};}return o;}
  function defaults(){return {jornal:P.P_JORNAL, r1:'paq', r2f:'paq', ez:'paq', et:'paq', cer:'paq', gg:P.P_GG, sup:P.P_SUP, et_pct:P.P_ET,
    tsd:P.P_TSD, cron:D.flu.slice(), carc:'no'};}
  function paqOf(t,S){const q=D.paq.find(x=>x.t===t);
    const r1=S.r1==='paq'?q.r1:+S.r1;
    const r2=S.r2f==='paq'?q.r2:(S.r1==='paq'?q.r1:+S.r1)*(+S.r2f);
    const ez=S.ez==='paq'?q.ez:(q.ez>0?+S.ez:0);
    const et=S.et==='paq'?q.et:(q.et>0?+S.et:0);
    const cer=S.cer==='paq'?q.cer:+S.cer;
    return {r1,r2,mz:ez>0?10000/ez:0,mt:et>0?10000/et:0,cer};}
  function block(b,S,CU,carcML){
    const q=paqOf(b.tip,S), j=S.jornal;
    const ind=S.gg+S.sup+S.et_pct;
    const E=b.r1a+b.r1b, F=b.r2a+b.r2b, G=b.ca+b.cb;
    const H=E*q.r1, I=F*q.r2, J=H+I;
    const K=(b.r1a+b.r2a)*q.mz/1000, L=(b.r1a+b.r2a)*q.mt/1000, M=b.per*q.cer;
    const ml=S.carc==='inv'?(carcML[b.id]||0):(typeof b.cml==='number'?b.cml:0);
    const N=ml/1000;
    const dk=b.dac/1000*P.P_SINUOS;
    const O=j/(P.P_CARGA*(P.P_HORAS*P.P_EFIC/(dk/P.P_VIDA+dk/P.P_VVUELTA+P.P_TCD/60)));
    const rec=1+P.P_RECALCE;
    const Pp=J*rec*CU.PLANTON.h, Q=J*CU.PLANTA.h, R=(E+F)*CU.HA.h;
    const Sf=J*rec*(P.P_PESO_PLANTON+P.P_PESO_INS)*O, T=J*rec*P.P_PESO_INS*P.P_FT_KG;
    const U=K*CU.ZANJA.h, V=L*CU.TERRAZA.h, W=M*CU.CERCO.h, X=N*CU.CARCAVA.h;
    const Y=Pp+Q+R+Sf+T+U+V+W+X, Z=Y*ind, AA=Y+Z;
    const AC=J*rec*CU.PLANTON.nc+J*CU.PLANTA.nc+(E+F)*CU.HA.nc+Sf+K*CU.ZANJA.nc+L*CU.TERRAZA.nc+M*CU.CERCO.nc+N*CU.CARCAVA.nc;
    const AD=J*rec*CU.PLANTON.c+J*CU.PLANTA.c+(E+F)*CU.HA.c+K*CU.ZANJA.c+L*CU.TERRAZA.c+M*CU.CERCO.c+N*CU.CARCAVA.c;
    const AE=Y-AC-AD;
    const AF=AC*P.P_FC_MONC+AD*P.P_FC_MOC+(AE+Z)*P.P_FC_BIEN;
    const AG=(Pp+Q+R+Sf+T+W)*(1+ind), AH=(U+V+X)*(1+ind);
    return {id:b.id,tip:b.tip,ha:b.ha,E,F,G,H,I,J,K,L,M,N,O,P:Pp,Q,R,S:Sf,T,U,V,W,X,Y,Z,AA,AB:b.ha?AA/b.ha:0,AC,AD,AE,AF,AG,AH,
      pl_kg:J*rec*(P.P_PESO_PLANTON+P.P_PESO_INS), ins_kg:J*rec*P.P_PESO_INS, plantones:J*rec};
  }
  function omRate(j){const H=lerp(D.omH,j), I=lerp(D.omI,j); return I/H;}
  function run(ids,S){
    const CU=cuAt(S.jornal);
    const carcML=D.carcML||{};
    const rows=D.blocks.filter(b=>ids.has(b.id)).map(b=>block(b,S,CU,carcML));
    const t={}; const sumK='ha E F G H I J K L M N P Q R S T U V W X Y Z AA AC AD AE AF AG AH pl_kg ins_kg plantones'.split(' ');
    sumK.forEach(k=>t[k]=rows.reduce((s,r)=>s+r[k],0));
    // O&M
    const r=omRate(S.jornal), hh=S.jornal/P.P_HORAS;
    const indv={ha:t.E+t.F, pl:t.J, kz:t.K, kt:t.L, kc:t.N};
    const om=D.oym.map(o=>{const f=o.c/o.d*(r*hh/P.P_HH_PEON_2020+(1-r)*P.P_FACT_IPC); return {...o,f,q:indv[o.ind],anual:f*indv[o.ind]};});
    const om1=om.reduce((s,o)=>s+o.anual,0);
    const nEj=P.P_ANIOS_EJ, nOM=P.P_ANIOS_OM, years=nEj+nOM;
    const fsoc=r*P.P_FC_MONC+(1-r)*P.P_FC_BIEN;
    const flow=[];
    for(let y=1;y<=years;y++){
      const tt=Math.max(0,y-nEj);
      const inv=t.AA*(S.cron[y-1]||0), invS=t.AF*(S.cron[y-1]||0);
      const omy=tt>0?om.reduce((s,o)=>s+(o.anios>=tt?o.anual:0),0):0;
      flow.push({y,inv,om:omy,tot:inv+omy,invS,omS:omy*fsoc,totS:invS+omy*fsoc});
    }
    const npv=(k)=>flow.reduce((s,f)=>s+f[k]/Math.pow(1+S.tsd,f.y),0);
    const ind={vacP:npv('tot'),vacS:npv('totS'),vaInv:npv('inv'),vaOM:npv('om'),om15:flow.reduce((s,f)=>s+f.om,0)};
    ind.ceha=(t.E+t.F)?ind.vacS/(t.E+t.F):0; ind.ceblk=t.ha?ind.vacS/t.ha:0;
    return {rows,t,om,om1,flow,ind,CU,omr:r};
  }
  return {run,defaults,cuAt,paqOf};
}
if(typeof module!=='undefined') module.exports=cpEngine;
