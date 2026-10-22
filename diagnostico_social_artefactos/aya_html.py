import json
s=open('aya_orig.html').read(); L=s.split('\n')
X=json.load(open('aya_X.json'))
X['find_s'].insert(0,["cri","Sin avance de campo entre los cortes del 23/09 y el 26/09",
 ["El respaldo del aplicativo al <strong>26/09/2026</strong> mantiene las mismas <strong>19 fichas</strong> de Ayabaca (11 F-DS-01 y 8 F-DS-02, en 6 de los 10 bloques); el último registro es del 14/09/2026. En el mismo periodo Huancabamba pasó de 29 a 49 fichas y Morropón de 150 a 207, ambas con entrevistas F-DS-03.",
  "Ayabaca es la única provincia sin ninguna entrevista a autoridades (F-DS-03): la percepción del riesgo, las zonas a conservar y la disposición a participar siguen sin sustento de campo en los 10 bloques."]])
X['disc_ds'].append({"c":"DS-13","m":"Consolidado del aplicativo al 26/09/2026","n":"Medio",
 "d":"El libro de gráficos del aplicativo (26/09/2026) reporta 19 fichas vigentes, 10 ámbitos, 788 habitantes y 694 familias en 6 bloques; persiste la suma de las tres F-DS-01 de Guabal y el ámbito agregado «Guabal / Nogal / Guanábano Alto» se cuenta como centro poblado aparte.",
 "e":"Las cifras de cabecera del aplicativo no coinciden con las depuradas de este volumen (508 familias en 7 localidades).",
 "t":"Se mantienen las cifras depuradas; el corte no aportó registros nuevos.",
 "a":"Depurar las reediciones de Guabal antes de la siguiente exportación."})
X['pend_ds'].insert(0,['P-DS-00','Reactivar el levantamiento social','Programar la segunda campaña en Ayabaca: sin fichas nuevas desde el 14/09/2026, mientras los otros dos volúmenes ya cuentan con F-DS-03.','Coordinación social SESDI','Inmediata'])
assert L[946].startswith('const X')
L[946]='const X = '+json.dumps(X,ensure_ascii=False)+';'
s='\n'.join(L)
for a,b in [('Volumen III · Revisión 3 · Corte 23/09/2026','Volumen III · Revisión 4 · Cortes: DT 23/09/2026 · DS 26/09/2026'),
 ('las 19 fichas F-DS del aplicativo IN Piura (corte 23/09/2026)','las 19 fichas F-DS del aplicativo IN Piura (respaldo y consolidado al 26/09/2026; sin registros nuevos desde el 14/09/2026)'),
 ("label:'Fichas vigentes (corte 23/09/2026)'","label:'Fichas vigentes (corte 26/09/2026)'"),
 ('aplicativo IN Piura, F-DS-01 (23/09/2026).','aplicativo IN Piura, F-DS-01 (corte 26/09/2026).'),
 ('fichas F-DS-01 y F-DS-02 (corte 23/09/2026; brigada social)','fichas F-DS-01 y F-DS-02 (corte 26/09/2026, sin variación respecto del 23/09; brigada social)')]:
    assert s.count(a)==1,a; s=s.replace(a,b)
open('aya_new.html','w').write(s); print('ok')
