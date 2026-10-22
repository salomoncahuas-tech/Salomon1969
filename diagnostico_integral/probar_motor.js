// Comprueba que src/engine.js reproduce el Excel (datos/ref.json) en los 117 bloques,
// el piloto y la cartera, con jornal de S/ 80, 100 y 120.  Uso: node probar_motor.js
const path=require('path');
const D=require(path.join(__dirname,'datos','costos_data.json'));
const ref=require(path.join(__dirname,'datos','ref.json'));
const E=require(path.join(__dirname,'src','engine.js'))(D);
const all=new Set(D.blocks.map(b=>b.id)), pil=new Set(D.blocks.filter(b=>b.pil==='Sí').map(b=>b.id));
let worst=0, fallas=0;
for(const j of [100,120,80]){
  const S=E.defaults(); S.jornal=j;
  const T=E.run(all,S), Pp=E.run(pil,S), R=ref[j].res;
  const cmp=(name,a,b)=>{const e=Math.abs(a-b)/Math.max(1,Math.abs(b)); worst=Math.max(worst,e); if(e>1e-9){fallas++; console.log('DIFERENCIA',j,name,a,b,e);}};
  cmp('C9',T.t.ha,R.C9);cmp('C12',T.t.J,R.C12);cmp('C13',T.t.K,R.C13);cmp('C14',T.t.L,R.C14);cmp('C15',T.t.M,R.C15);
  cmp('C17',T.t.Y,R.C17);cmp('C18',T.t.AA,R.C18);cmp('C19',T.t.AG,R.C19);cmp('C20',T.t.AH,R.C20);
  cmp('C22',T.om1,R.C22);cmp('C23',T.ind.om15,R.C23);cmp('C27',T.ind.vacS,R.C27);cmp('C28',T.ind.ceha,R.C28);
  cmp('B18',Pp.t.AA,R.B18);cmp('B22',Pp.om1,R.B22);cmp('B23',Pp.ind.om15,R.B23);cmp('B27',Pp.ind.vacS,R.B27);cmp('B28',Pp.ind.ceha,R.B28);
  for(const r of T.rows){const x=ref[j].blk[r.id]; ['J','K','L','M','Y','AA','AF','AG','AH'].forEach((k,i)=>cmp(r.id+k,r[k],x[i]));}
  console.log(`jornal S/ ${j}: inversión cartera ${T.t.AA.toFixed(2)} (Excel ${R.C18.toFixed(2)}) · VAC social ${T.ind.vacS.toFixed(2)} (Excel ${R.C27.toFixed(2)})`);
}
console.log('error relativo máximo', worst);
process.exit(fallas?1:0);
