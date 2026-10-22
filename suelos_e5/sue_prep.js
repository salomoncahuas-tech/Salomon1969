
/* =====================================================================
   E5 ESTUDIO DE SUELOS — preparación por bloque (lee window.SUE)
   Fuente: GeoSIG Ingenieros, Entregable 2 (fase de campo, datos al
   20/09/2026) y capa CALICATAS_EJECUTADAS del Anexo 5. Laboratorio
   pendiente: taxonomía, CUM y Factor K son tentativos.
   ===================================================================== */
const SU=window.SUE||{C:[],K:[],I:[],R:[],BL:{},U:{}};
const BSET=new Set(B.map(b=>b.b));
const CAL=SU.C; CAL.forEach(c=>{const bl=SU.BL[c.b]; c.prov=bl?bl.p:'—'; c.dist=bl?bl.d:'—';
  const vb=B.find(b=>b.b===c.b); if(vb) c.dist=vb.d;});
const CV=CAL.filter(c=>BSET.has(c.b));                 // calicatas de este volumen
const KV=SU.K.filter(k=>CV.some(c=>c.c===k.cal));      // chequeos de este volumen
const IV=SU.I.filter(s=>CV.some(c=>c.c===s.cal));      // pruebas de infiltración de este volumen
B.forEach(b=>{
  const cb=CV.find(c=>c.b===b.b&&c.t==='B');
  b.sCal=cb?cb.c:null; b.prof=cb&&cb.pcm!=null?cb.pcm:null; b.ero=cb?cb.eroN:null;
  b.sLim=cb?((cb.pcm!=null&&cb.pcm<50)||cb.eroN===4):null;
  b.sAll=CV.filter(c=>c.b===b.b).map(c=>c.c);
});
const SLIM=B.filter(b=>b.sLim===true);
const CB=CV.filter(c=>c.t==='B');
const median=a=>{const s=a.filter(v=>v!=null).sort((x,y)=>x-y);if(!s.length)return null;const m=s.length>>1;return s.length%2?s[m]:(s[m-1]+s[m])/2;};
