import json
D = json.load(open('mor_D_new.json'))
C, LOC, B = D['C'], D['LOC'], D['B']
V6s = {b['b'] for b in B}
LV6 = [l for l in LOC if any(x in V6s for x in l['bl'])]
LFV6 = [l for l in LV6 if l['nreg']]
both = [l for l in LOC if l['pob'] and l['pdec'] and not l['v5only']]
C['RATIO'] = sum(l['pdec'] for l in both) / sum(l['pob'] for l in both)
CT = C['CTRL']
CT.update(locs_car=len(LFV6), pob_car=sum(l['pob'] or 0 for l in LFV6), nreg=207, nvig=189, ndep=sum(C['CDEP'].values()),
          bl_ficha=sum(1 for b in B if b['ds']['caract'] == 'Con ficha propia'),
          sincar=[b['b'] for b in B if b['ds']['caract'].startswith('Sin')],
          via=[b['b'] for b in B if b['ds']['caract'].startswith('Loc')],
          act=C['NLACT'], ent=C['N03'], corte='26/09/2026')
pc = 100 * CT['pob_car'] / CT['pob6']
print(CT, round(C['RATIO'], 2), len(both), round(pc, 1))

C['FIND_DS'] = [
 ["adv", "La base social cubre casi toda la provincia; el conglomerado 83–87 sigue siendo el vacío",
  [f"<strong>{CT['bl_ficha']} de los 59 bloques</strong> tienen ficha social propia y otros {len(CT['via'])} tienen su localidad caracterizada desde un bloque vecino. Las <strong>{CT['locs_car']} localidades con ficha</strong> de los bloques V6 reúnen {CT['pob_car']:,} de los 9,317 habitantes INEI 2017 asociados (" + f"{pc:.1f}".replace('.', '.') + " %), frente a 53 localidades y 73.9 % al 23/09.",
   "Sin caracterización social: <strong>" + ', '.join(CT['sincar']) + "</strong>. En el conglomerado de mayor peligro modelado (83–87: Dótor, Cardal, Nueva Esperanza y Miguel Pampa) solo hay dos fichas de actores: la ronda neutral de Miguel Pampa y, desde el 25/09, el teniente gobernador de Cardal (87), a favor del proyecto. Siguen sin F-DS-01 ni F-DS-03."],
  "Respaldo del aplicativo IN Piura, 26/09/2026."],
 ["cri", "El aplicativo multiplica los registros de una misma localidad",
  [f"El respaldo registra <strong>207 fichas</strong> (110 F-DS-01, 56 F-DS-02, 41 F-DS-03) —el consolidado reconoce 189 vigentes— frente a 150 en la exportación por bloque del 23/09. Depuradas quedan <strong>{CT['ndep']} fichas</strong> ({C['CDEP']['F-DS-01']} / {C['CDEP']['F-DS-02']} / {C['CDEP']['F-DS-03']}) en {sum(1 for l in LOC if l['nreg'])} localidades.",
   "La F-DS-01, instrumento de localidad, se aplica a varios informantes con los mismos datos: La Alberca tiene 14 registros (12 informantes), Bigote y Alan García 8, Quemazón 6 y La Pilca 5. El nombre de la responsable se sigue escribiendo con sufijos (puntos, «xxx», letras repetidas) que evaden la regla de duplicidad."],
  "Depuración propia (una F-DS-01 por centro poblado); ver Discrepancias."],
 ["adv", "San Juan de Bigote concentra el avance del corte",
  ["Entre el 24 y el 25/09 se registraron o reeditaron 29 F-DS-01 en San Juan de Bigote: se incorporan Santa Rosa (M3B8), Quemazón y La Pareja (M11B3) y San Juan Bautista (M3B5), y se completan Manzanares / Bado de Garzas (M3B6), Polluco / Sinaí (M3B7) y Alan García / Bigote (M3B3).",
   "Las fichas describen caseríos con tierras del Estado o de la C.C. de Andanjo, agua de manantiales (San Rafael, Piedra Blanca, El Gallo) y de pozos con paneles solares, crianza de cabras como principal ingreso (hasta 20 cabras vendidas por familia al año) y bloques «ocupados por la población con sus animales» (Santa Rosa)."],
  "F-DS-01, San Juan de Bigote, 24–25/09/2026."],
 ["adv", "La población declarada no es sumable",
  [f"Donde existen ambos datos ({len(both)} localidades de bloques V6), la población declarada en F-DS-01 supera a la censal en <strong>×{C['RATIO']:.2f}</strong> (×1.53 al 23/09). Buenos Aires capital (8,000), Serrán (3,000), La Alberca (1,500), Ingenio de Buenos Aires (1,400) y Quemazón (750) no tienen cifra INEI de contraste.",
   "El consolidado del aplicativo suma 63,156 habitantes y 16,396 familias sobre registros repetidos: la población de referencia del perfil debe salir del Censo 2017 proyectado al año de evaluación."],
  "F-DS-01 y catálogo INEI–bloques."],
 ["cri", "La sequía es el problema que la población nombra primero",
  [f"En las {C['NL03']} localidades con entrevista F-DS-03: sequía o déficit hídrico en <strong>{C['HZ'].get('Sequía / déficit hídrico',0)}</strong>, deforestación o tala en {C['HZ'].get('Deforestación / tala',0)}, huaicos en {C['HZ'].get('Huaicos / flujos',0)}, activación de quebradas en {C['HZ'].get('Activación de quebradas / inundación',0)}, filtración en cerros y aumento de temperatura en {C['HZ'].get('Filtración de agua en cerros',0)}, deslizamientos en {C['HZ'].get('Deslizamientos / movimientos en masa',0)}.",
   "La población nombra los procesos que el proyecto regula (PMM, EPH, PGI) a través de sus efectos: quebradas que aíslan caseríos, filtraciones que dañan viviendas y cultivos, suelos que se agrietan. En Quemazón y Manzanares la crecida de quebradas y del río afecta los sembríos."],
  "Codificación propia de F-DS-03, numerales 1.1 a 1.3."],
 ["cri", "El palo santo se vende",
  ["La tala comercial de palo santo aparece en Río Seco («se ha vuelto un negocio rentable»), La Pilca («se llevan en camiones y manifiestan que tienen permiso»), Tórtola y Huaro Quispampa.",
   f"Especies nativas mencionadas por uso o disminución: palo santo ({C['SPP'].get('Palo santo',0)} localidades), algarrobo ({C['SPP'].get('Algarrobo',0)}), hualtaco, faique, charán y overal ({C['SPP'].get('Hualtaco',0)} cada una). Hualtaco y palo santo están En Peligro Crítico (D.S. N.° 043-2006-AG)."],
  "F-DS-03."],
 ["ok", "El gas doméstico ya redujo la presión por leña",
  ["En varias entrevistas las familias cocinan con gas y afirman que por ello la leña se usa menos (Chacayo: «desde que usamos gas ya no talamos»). Varias localidades declaran la prohibición comunal de talar en verde.",
   "La presión dominante sobre el bosque seco es hoy el pastoreo —que en San Juan de Bigote ocupa los propios bloques— y la extracción comercial, no el consumo doméstico de leña."],
  "F-DS-03, numeral 2.3; F-DS-01, San Juan de Bigote."],
 ["adv", "Posición favorable, pero la aceptación de autoridades no es aceptación comunal",
  [f"{C['ACU']} de {C['N03']} entrevistas expresan acuerdo con el proyecto. De {C['NLACT']} actores valorados, {C['POS'].get('A favor del proyecto',0)} están a favor y 5 son neutrales o sin posición: las rondas de Miguel Pampa («siempre llegan a ofrecer proyectos y no se cumplen»), Palo Blanco–El Cerezo, Serrán y Santa Rosa (San Juan de Bigote), y el teniente gobernador de Juan Velasco. La ronda de Santa Rosa condiciona su apoyo a la socialización con la población.",
   f"Solo {C['GEN'].get('Mujer',0)} de {C['N03']} entrevistas corresponden a mujeres. Ninguna ficha F-DS-04 a F-DS-07 (talleres, conflictos, percepción de peligros, consentimiento) fue aplicada."],
  "Respaldo del aplicativo, 26/09/2026."],
 ["cri", "La tenencia es mayoritariamente estatal o comunal",
  ["De las 43 localidades con F-DS-01, 26 declaran tierras del Estado, 12 tierras comunales (sobre todo la C.C. de Andanjo en San Juan de Bigote), 4 propiedad individual titulada y 1 posesión. El acuerdo de intervención debe suscribirse con el GORE o la SBN y con las asambleas comunales, además de los posesionarios.",
   "En San Juan Bautista la asociación de ganaderos paga a la C.C. de Andanjo S/ 1,500 al año por el uso de 250 ha de pastos: la clausura temporal afecta un derecho de uso ya pactado. La Pareja declara un conflicto de linderos con Alto San José."],
  "F-DS-01, numeral 4, y observaciones."],
 ["cri", "La resiliencia descansa en juntas, JASS y rondas; no existe alerta temprana",
  [f"En las {C['NF01']} localidades con F-DS-01 hay junta directiva vigente y, en casi todas, JASS. Ninguna ficha registra sistema de alerta temprana, comité de gestión del riesgo ni mecanismo de retribución por servicios ecosistémicos.",
   f"Fragilidad social preliminar: {C['FRAG'].get('Alta',0)} localidades Alta, {C['FRAG'].get('Media',0)} Media y {C['FRAG'].get('Baja',0)} Baja; Manzanares / Bado de Garzas y Polluco / Sinaí se suman a la clase Alta (sin establecimiento de salud, letrina seca, migración alta). Resiliencia: {C['RES'].get('Alta',0)} Alta y {C['RES'].get('Media',0)} Media."],
  "F-DS-01; regla de calificación en la pestaña Servicios y economía."],
 ["ok", "Hay con quién construir la línea de gobernanza",
  ["Comités del ACR Bosques Secos de Salitral–Huarmaca (11 localidades con vínculo declarado), FONCODES (cocinas mejoradas, biohuertos, paneles solares), CIPCA (ovinos mejorados en Santa Rosa, granjas en Bigote, asistencia técnica en La Pareja), ANCEP, asociaciones de ganaderos (Nueva Juventud San Juan Bautista, 45 socios) y un ofrecimiento de terreno para vivero en Serrán.",
   "Las autoridades identifican zonas a conservar con nombre propio —manantiales y cerros de recarga— que son la base de un MERESE hídrico; en San Juan de Bigote se declara mano de obra disponible de 20 a 150 personas por caserío."],
  "F-DS-01, F-DS-02 y F-DS-03."],
]

DI = {d['c']: d for d in C['DISC']}
DI['DS-01'].update(m="Registros repetidos de una misma localidad",
    d="207 registros se reducen a %d fichas depuradas. La Alberca concentra 14 F-DS-01 (12 informantes) con los mismos datos; Bigote y Alan García 8; Quemazón 6. El nombre de la responsable se escribe con sufijos («xxxx», puntos, letras repetidas) en al menos 30 registros." % CT['ndep'],
    t="Se consolida una F-DS-01 por centro poblado declarado (valores modales; población y familias por mediana); las F-DS-02 y F-DS-03 se depuran por actor y entrevistado. La regla se aplica igual en los tres volúmenes.",
    a="Fijar el responsable desde el usuario autenticado; si la F-DS-01 se aplica como encuesta a hogares, rediseñar el instrumento.")
DI['DS-02'].update(m="Fuente de conteo del diagnóstico social", n="Bajo",
    d="Hasta el 23/09 se usó la exportación por bloque (150 registros) y el libro consolidado (175). Al 26/09 el respaldo completo del aplicativo contiene 207 registros, todos vinculados a un bloque; el consolidado reconoce 189 vigentes tras su regla de duplicidad.",
    e="Resuelto: los registros de El Ala, Bado de Garzas, Sinaí, Taspa, La Laja y Mambluque quedan incorporados.",
    t="Se usa el respaldo del aplicativo como fuente única de conteo, actividades, actores y entrevistas.",
    a="Mantener el respaldo como fuente de verdad en las siguientes actualizaciones.")
DI['DS-06'].update(d="Buenos Aires capital 8,000 hab.; Serrán 3,000; La Alberca 1,500 con 500 familias; Ingenio de Buenos Aires 1,400; Quemazón entre 350 y 1,000; Alan García / Bigote con cuatro cifras distintas (120, 200, 1,000 y 3,000). Donde hay ambos datos (%d localidades de bloques V6), lo declarado supera a lo censal ×%.2f." % (len(both), C['RATIO']))
DI['DS-08']['d'] = "Las fichas de Bigote aparecen vinculadas también a M17B1 (Chulucanas), cuyo centro poblado INEI es Papelillo."
DI['DS-09'].update(d="El tablero de la revisión 2 (base 11/09/2026) registraba al presidente de la ronda de Taylín de Tuñalí «en contra»; en la base del 26/09/2026 no figura ningún actor en contra y la ronda de Santa Rosa (San Juan de Bigote) se registra como neutral.")
DI['DS-12'].update(m="Tenencia estatal y comunal", d="De las 43 localidades con F-DS-01, 26 declaran tierras del Estado y 12 comunales; solo 4 declaran propiedad individual titulada. El consolidado del aplicativo cuenta 59 fichas con tenencia estatal.",
    e="El acuerdo de intervención requiere al titular estatal (GORE/SBN) y a las comunidades, además de los posesionarios.",
    t="Se registra el régimen declarado; el tamizaje predial del DT decide.")
DI['DS-14'].update(d="Las fichas F-DS-03 registran la duración como texto («1 hora», «50 minutos»); el campo numérico que exporta el aplicativo la reduce a 1 y el consolidado reporta una media de 25 min.",
    e="Resuelto con el texto de la ficha.", t="Se adopta la duración escrita en la ficha, en minutos.")
DI['DS-15'] = {"c": "DS-15", "m": "Conflicto de linderos y derechos de pastoreo", "n": "Alto",
    "d": "La Pareja (M11B3) declara un conflicto de linderos con Alto San José; en San Juan Bautista la asociación de ganaderos paga S/ 1,500 al año a la C.C. de Andanjo por 250 ha de pastos; en Santa Rosa el bloque está ocupado por la población y sus animales.",
    "e": "Son los primeros conflictos y derechos de uso documentados; ninguno figura en un registro F-DS-05.",
    "t": "Se registran como riesgo social (RS-11).",
    "a": "Aplicar F-DS-05 en San Juan de Bigote y cruzar con el tamizaje predial de M11B3, M3B5 y M3B8."}
C['DISC'] = list(DI.values())

R = {r['c']: r for r in C['RIESGOS']}
R['RS-02'].update(ev="Dótor y Nueva Esperanza (SJB) sin ficha; en el conglomerado solo hay dos F-DS-02: la ronda neutral de Miguel Pampa (83–84) y el teniente gobernador de Cardal (87), a favor. Es el conjunto de mayor peligro y con ocupación interna.")
R['RS-03'].update(ev="0 de 59 bloques con tamizaje predial; 26 de 43 localidades con F-DS-01 declaran tierras del Estado y 12 comunales; decisiones individuales sobre parcelas en comunidades de Santo Domingo.")
R['RS-04'].update(ev="Ganadería caprina u ovina en %d localidades; bloques ocupados por la población con sus animales en Santa Rosa (M3B8); venta de hasta 20 cabras por familia al año en Polluco / Sinaí." % C['ACTS'].get('Ganadería caprina/ovina', 0))
R['RS-05'].update(ev="Miguel Pampa: «siempre llegan a ofrecer la ejecución de proyectos y no se cumplen»; Polluco: «totalmente olvidados por las autoridades», proyecto de paneles solares inconcluso y paneles robados.", amb="SJB; Salitral")
R['RS-06'].update(ev="Rondas de Miguel Pampa (candidatura a regidor), Palo Blanco–El Cerezo, Serrán y Santa Rosa (SJB), y teniente gobernador de Juan Velasco.")
R['RS-09'].update(ev="%d de %d entrevistas son a mujeres; una sola organización de mujeres (ANCEP)." % (C['GEN'].get('Mujer', 0), C['N03']))
R['RS-10'].update(ev="Pozos que se secan (San Pedro); bombeo con motores (La Alberca); pozos con paneles solares en Polluco y San Juan Bautista; captación de San Rafael con tubos sucios; riego hasta septiembre (Ingenio).")
R['RS-11'] = {"c": "RS-11", "r": "Conflicto de linderos y derechos de pastoreo pactados", "ev": "La Pareja: conflicto de linderos con Alto San José; San Juan Bautista: pago anual de S/ 1,500 a la C.C. de Andanjo por 250 ha de pastos.",
    "amb": "San Juan de Bigote (M11B3, M3B5)", "p": "Media", "i": "Alto",
    "med": "Aplicar F-DS-05; acuerdos de clausura con la asociación de ganaderos y la C.C. de Andanjo, con pastos alternativos; excluir áreas en litigio.", "n": "Alto"}
C['RIESGOS'] = list(R.values())

O = C['OPP']
O[5] = ["FONCODES, CIPCA y Gobierno Regional", "Programas con presencia (cocinas mejoradas, biohuertos, paneles solares, ovinos mejorados, granjas, promotores agrarios) que complementan las medidas productivas."]
O[7] = ["Oferta de terreno para vivero y mano de obra local", "Serrán ofrece terreno para vivero; las localidades declaran entre 20 y 800 personas disponibles como mano de obra (150 en Manzanares, 100 en San Juan Bautista)."]
O.append(["Asociaciones de ganaderos", "Nueva Juventud San Juan Bautista (45 socios; incluye Bado de Garzas y Manzanares) ya paga por el uso de pastos comunales: interlocutor para la clausura temporal y el manejo de pastizales."])

I = C['INST']
for r in I:
    if r[0].startswith('CIPCA'):
        r[2] = "Alan García / Bigote; Polluco / Sinaí; Santa Rosa y La Pareja (SJB)"
        r[3] = "Promotores agrarios, granjas, ovinos mejorados (20 familias en Santa Rosa), asistencia técnica."
    if r[0].startswith('FONCODES'):
        r[2] = "Salitral; SJB (Manzanares, Polluco, La Pareja, San Juan Bautista)"
        r[3] = "Cocinas mejoradas, biohuertos, riego tecnificado, paneles solares para agua."
    if r[0].startswith('Rondas'):
        r[3] = "Regulan el acceso al territorio; cuatro con posición neutral (Miguel Pampa, Palo Blanco–El Cerezo, Serrán, Santa Rosa de SJB)."
I.insert(9, ["Asociación de ganaderos Nueva Juventud San Juan Bautista", "Organización productiva", "M3B5, M3B6 (SJB)", "45 socios; paga a la C.C. de Andanjo por 250 ha de pastos.", "Gestionar de cerca"])

C['HIP'] = [
 ["Movimientos en masa (PMM)", "Filtración de agua en cerros, deslizamientos que dañan viviendas y chacras", "%d + %d localidades" % (C['HZ'].get('Filtración de agua en cerros', 0), C['HZ'].get('Deslizamientos / movimientos en masa', 0)), "Confirmar con F-DS-06 en la interfaz (58, M19B2, M32B3, M6B10, M3B3)."],
 ["Erosión hídrica (EPH)", "Suelo agrietado, tierras áridas, pérdida de áreas verdes", "%d localidades" % C['HZ'].get('Erosión / degradación del suelo', 0), "Contrastar con el conglomerado 83–87, aún sin F-DS-03."],
 ["Generación de inundaciones (PGI)", "Quebradas que se activan, desbordes, crecida del río sobre sembríos, aislamiento", "%d + %d localidades" % (C['HZ'].get('Activación de quebradas / inundación', 0), C['HZ'].get('Huaicos / flujos', 0)), "Delimitar la zona de impacto y la población aguas abajo."],
 ["Cambio climático", "Menos lluvia, lluvias que se adelantan, más calor, sequía agosto–diciembre", "%d + %d localidades" % (C['HZ'].get('Sequía / déficit hídrico', 0), C['HZ'].get('Aumento de temperatura / cambio climático', 0)), "Cruzar con el escenario MPI-ESM1-2HR SSP585."],
]

M = C['MEDIDAS']
M[1][2] = "Infraestructura natural marrón: clausura temporal, zanjas de infiltración, terrazas de formación lenta y diques para el control de cárcavas (39 registradas). Infraestructura natural verde: revegetación en macizo y enriquecimiento."
M[2][2] = "Infraestructura natural verde: recuperación de cobertura en cabecera de quebrada. Infraestructura complementaria: diques de contención de baja altura en quebradas activas."

P = {p[0]: p for p in C['PEND']}
P['P-01'][2] = "Levantar el diagnóstico social del conglomerado 83–87 (Dótor, Nueva Esperanza, Miguel Pampa; Cardal solo tiene F-DS-02) y de M17B4, M8B2, M6B2-1, M6B2-3."
P['P-04'][2] = "Depurar el aplicativo: registros repetidos por localidad, variantes del nombre de la responsable, catálogo V6, campo microcuenca, validación por sexo y duración de entrevistas."
C['PEND'].append(["P-17", "Alta", "Aplicar F-DS-05 en San Juan de Bigote (linderos de La Pareja; pastos de la C.C. de Andanjo en San Juan Bautista) antes de socializar metas de clausura.", "Especialista social", "Registro de conflictos y acuerdos de pastoreo"])

json.dump(D, open('mor_D_new.json', 'w'), ensure_ascii=False)
print('ok')
