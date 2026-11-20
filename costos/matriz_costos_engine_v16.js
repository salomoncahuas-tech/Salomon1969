/* Matriz de Costos Paramétricos IN Piura v1.6 — réplica de Costos_Parametros_IN_Piura_v1.6.xlsx
   (hojas BLOQUES + METAS_MARRON → CALCULO ← CU_COMP ← PARTIDAS → OYM → FLUJO). D = datos extraídos del libro.
   Metas marrón: plano de obras de conservación de suelos (06/10/2026), por bloque. */
function cpEngine(D){
  const P=D.par;
  const lerp=(a,j)=>a[0]+(a[1]-a[0])*(j-100)/20;          // CU lineal en el jornal (validado 100/120)
  const BASIC=['PLANTON','PLANTA','HA','CERCO','CARCAVA','BVIVA_T','BVIVA_Z'];
  function defaults(){return {jornal:P.P_JORNAL, r1:'paq', r2f:'paq', cer:'paq', gg:P.P_GG, sup:P.P_SUP, et_pct:P.P_ET,
    tsd:P.P_TSD, cron:D.flu.slice(), carc:'no', v8:'v7', zsec:'plano', deriv:'si', bviva:'si', arido:'1',
    pirc:'no', rell:'no', elimT:'no', elimZ:'no', podas:'no', fert:'no'};}
  const inc=(x,S)=>x.st==='inc'||S[x.id]==='si';
  function cuAt(j,S){
    const o={};
    BASIC.forEach(k=>{const c=D.cu[k]; o[k]={h:lerp(c.h,j),nc:lerp(c.nc,j),c:lerp(c.c,j)};});
    const fv=S.zsec==='plano'?P.P_ZI_ANCHO*P.P_ZI_PROF/(0.4*0.3):1, ft=S.zsec==='plano'?P.P_ZI_PROF/0.3:1;
    const fac=f=>f==='v'?fv:f==='t'?ft:1;
    const comp=L=>{const c={h:0,nc:0,c:0,parts:[]};
      L.forEach(x=>{const m=x.m*fac(x.f), on=inc(x,S), h=lerp(x.h,j), nc=lerp(x.nc,j), cc=lerp(x.c,j);
        c.parts.push({k:x.k,d:x.d,u:x.u,m,cu:h,on,st:x.st,id:x.id,par:m*h});
        if(on){c.h+=m*h; c.nc+=m*nc; c.c+=m*cc;}});
      return c;};
    o.ZANJA=comp(D.zp); o.TERRAZA=comp(D.tp); o.fv=fv; o.ft=ft;
    return o;
  }
  function paqOf(t,S){const q=D.paq.find(x=>x.t===t);
    const r1=S.r1==='paq'?q.r1:+S.r1;
    const r2=S.r2f==='paq'?q.r2:(S.r1==='paq'?q.r1:+S.r1)*(+S.r2f);
    const cer=S.cer==='paq'?q.cer:+S.cer;
    return {r1,r2,cer};}
  function block(b,S,CU,carcML){
    const q=paqOf(b.tip,S), j=S.jornal;
    const ind=S.gg+S.sup+S.et_pct;
    const E=b.r1a+b.r1b, F=b.r2a+b.r2b, G=b.ca+b.cb;
    const H=E*q.r1, I=F*q.r2, J=H+I;
    const aF=(b.av&&S.arido==='2')?0:1;
    const K=(b.zi+(S.deriv==='si'?b.der:0))*aF, L=b.tfl*aF, M=b.per*q.cer;
    const bvF=(S.bviva==='si'&&!(b.av&&S.arido!=='0'))?1:0;
    const BVT=L*bvF, BVZ=K*bvF;
    const ml=S.carc==='inv'?(carcML[b.id]||0):(typeof b.cml==='number'?b.cml:0);
    const N=ml/1000;
    const dk=b.dac/1000*P.P_SINUOS;
    const O=j/(P.P_CARGA*(P.P_HORAS*P.P_EFIC/(dk/P.P_VIDA+dk/P.P_VVUELTA+P.P_TCD/60)));
    const rec=1+P.P_RECALCE;
    const Pp=J*rec*CU.PLANTON.h, Q=J*CU.PLANTA.h, R=(E+F)*CU.HA.h;
    const Sf=J*rec*(P.P_PESO_PLANTON+P.P_PESO_INS)*O, T=J*rec*P.P_PESO_INS*P.P_FT_KG;
    const U=K*CU.ZANJA.h, V=L*CU.TERRAZA.h, W=M*CU.CERCO.h, X=N*CU.CARCAVA.h;
    const BVC=BVT*CU.BVIVA_T.h+BVZ*CU.BVIVA_Z.h;
    const Y=Pp+Q+R+Sf+T+U+V+W+X+BVC, Z=Y*ind, AA=Y+Z;
    const AC=J*rec*CU.PLANTON.nc+J*CU.PLANTA.nc+(E+F)*CU.HA.nc+Sf+K*CU.ZANJA.nc+L*CU.TERRAZA.nc+M*CU.CERCO.nc+N*CU.CARCAVA.nc+BVT*CU.BVIVA_T.nc+BVZ*CU.BVIVA_Z.nc;
    const AD=J*rec*CU.PLANTON.c+J*CU.PLANTA.c+(E+F)*CU.HA.c+K*CU.ZANJA.c+L*CU.TERRAZA.c+M*CU.CERCO.c+N*CU.CARCAVA.c+BVT*CU.BVIVA_T.c+BVZ*CU.BVIVA_Z.c;
    const AE=Y-AC-AD;
    const AF=AC*P.P_FC_MONC+AD*P.P_FC_MOC+(AE+Z)*P.P_FC_BIEN;
    const AG=(Pp+Q+R+Sf+T+W)*(1+ind), AH=(U+V+X+BVC)*(1+ind);
    return {id:b.id,tip:b.tip,ha:b.ha,av:b.av,E,F,G,H,I,J,K,L,M,N,O,P:Pp,Q,R,S:Sf,T,U,V,W,X,BVT,BVZ,BVC,Y,Z,AA,AB:b.ha?AA/b.ha:0,AC,AD,AE,AF,AG,AH,
      zz:b.zz*aF,tz:b.tz*aF,obra:b.obra,pl_kg:J*rec*(P.P_PESO_PLANTON+P.P_PESO_INS), ins_kg:J*rec*P.P_PESO_INS, plantones:J*rec};
  }
  function omRate(j){const H=lerp(D.omH,j), I=lerp(D.omI,j); return I/H;}
  function run(ids,S){
    const CU=cuAt(S.jornal,S);
    const carcML=D.carcML||{};
    const f8=S.v8==='v8'&&D.f8?D.f8:{};
    const sc=(b,f)=>({...b,ha:b.ha*f,r1a:b.r1a*f,r1b:b.r1b*f,r2a:b.r2a*f,r2b:b.r2b*f,ca:b.ca*f,cb:b.cb*f,zi:b.zi*f,tfl:b.tfl*f,der:b.der*f,zz:b.zz*f,tz:b.tz*f,obra:b.obra*f});
    const rows=D.blocks.filter(b=>ids.has(b.id)).map(b=>block(f8[b.id]?sc(b,f8[b.id]):b,S,CU,carcML));
    const t={}; const sumK='ha E F G H I J K L M N P Q R S T U V W X BVT BVZ BVC Y Z AA AC AD AE AF AG AH zz tz obra pl_kg ins_kg plantones'.split(' ');
    sumK.forEach(k=>t[k]=rows.reduce((s,r)=>s+r[k],0));
    t.nAv=rows.filter(r=>r.av).length;
    const r=omRate(S.jornal), hh=S.jornal/P.P_HORAS;
    const indv={ha:t.E+t.F, pl:t.J, kz:t.K, kt:t.L, kc:t.N};
    const om=D.oym.map(o=>{const f=o.c/o.d*(r*hh/P.P_HH_PEON_2020+(1-r)*P.P_FACT_IPC);
      const on=!o.st||S[o.id]==='si', anios=on?(o.anios||o.anios0):0;
      return {...o,f,on,anios,q:indv[o.ind],anual:on?f*indv[o.ind]:0};});
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
