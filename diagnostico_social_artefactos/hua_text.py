import json
D = json.load(open('hua_D_new.json'))
C = D['C']
LOC, B = D['LOC'], D['B']
LF = [l for l in LOC if l['nreg']]

# ---------------- CTRL ----------------
CT = C['CTRL']
CT.update({
    'nreg': 49, 'nzip': 49, 'nvig': 48, 'ndep': 33, 'nrep': 16,
    'bl_ficha': 14, 'bl_via': 1, 'locs_ficha': len(LF), 'cp_ficha': sum(len(l['inei']) for l in LF),
    'pob_ficha': sum(l['pob'] for l in LF),
    'sincar': [b['b'] for b in B if b['ds']['caract'].startswith('Sin')],
    'act': C['NLACT'], 'ent': C['N03'], 'n01dep': C['CDEP']['F-DS-01'], 'n02dep': C['CDEP']['F-DS-02'],
    'n03dep': C['CDEP']['F-DS-03'], 'ambitos': 17, 'corte': '26/09/2026'})
print(CT)

# ---------------- Hallazgos sociales ----------------
C['FIND_DS'] = [
 ["cri", "La base social crece, pero sigue siendo una muestra",
  ["El aplicativo registra <strong>49 fichas</strong> (25 F-DS-01, 16 F-DS-02, 8 F-DS-03) en <strong>14 de los 48 bloques</strong>, frente a 29 fichas en 13 bloques del corte del 23/09. El consolidado del aplicativo reconoce 48 vigentes; depuradas las reediciones quedan <strong>33 fichas</strong> (11 / 14 / 8) en 17 localidades, que cubren 18 de los 58 centros poblados INEI y 1,880 de los 6,573 habitantes censados (28.6 %).",
   "Se incorporan Pedregal (La Tuna, bloque 61) y Maray Grande (34), y se completan con F-DS-01 y F-DS-03 La Laguna (26), Hualtacal (81) y Almirante Miguel Grau (79). Siguen sin aplicarse 4 de los 7 instrumentos: talleres (F-DS-04), conflictos (F-DS-05), percepción de peligros (F-DS-06) y consentimiento previo (F-DS-07)."],
  "Consolidado DS del aplicativo y exportación por bloque, 26/09/2026."],
 ["cri", "El vacío social sigue coincidiendo con el mayor peligro",
  ["Los cinco bloques del núcleo de brecha (M4B4, M12B1, M4B3, M20B1, M2B8), todos Muy alto, continúan sin ficha social en el aplicativo. La hoja de control de campo reporta F-DS-01 en sus centros poblados (Hualcas I y II, Chignia Baja, Nueva Esperanza, Las Huacas, Huacas Baja): la información existe en papel pero no está digitalizada.",
   "El distrito de Huancabamba (bloques 64 y M30B5) sigue con 0 fichas registradas, aunque el control reporta 4 F-DS-01 en Pariamarca Centro. El avance del corte se concentró en Lalaquiz (30 de las 49 fichas) y Canchaque (11)."],
  "Hoja de control de campo, 17/09/2026 · aplicativo, 26/09/2026."],
 ["adv", "La brecha de digitalización se reduce, pero sigue siendo de uno a cinco",
  ["La hoja de control reporta <strong>126 F-DS-01 y 55 F-DS-03</strong> aplicadas sobre una muestra calculada de 113 (población finita, e = 5 %); el aplicativo tiene ahora 25 y 8 (antes 13 y 4). El instrumento de saberes ancestrales sigue en 0 de 55 aplicaciones.",
   "Las cifras sociales de este volumen describen 17 localidades de 4 distritos, no la provincia."],
  "D-HS01."],
 ["adv", "Las cifras declaradas no se pueden sumar tal como vienen",
  ["La cabecera del consolidado (6,370 hab.; 2,215 familias) suma varios informantes de una misma localidad: La Laguna tiene 5 F-DS-01 con los mismos datos, El Papayo 4, y Maray Chico y Ullma 3. Depurada, la población declarada es de <strong>2,610 hab. y 785 familias en 10 localidades</strong>.",
   "Frente al censo, la población declarada es en conjunto ×1.56 la de INEI 2017. Las mayores diferencias están en Almirante Miguel Grau (230 frente a 58, ×3.97), Pedregal (50 frente a 14, ×3.57: la población reside en Tunal) y Pirga (600 frente a 322, ×1.86). Se presentan ambas cifras; la validación corresponde al padrón comunal."],
  "D-HS02 a D-HS04."],
 ["ok", "Todos los actores consultados están a favor, con una condición",
  ["Los 14 actores mapeados (agentes municipales, tenientes gobernadores, rondas y JASS) declaran posición a favor, y las 8 autoridades entrevistadas expresan acuerdo con el proyecto por la recuperación del bosque, la prevención de movimientos en masa y el empleo local.",
   "El teniente gobernador de Almirante Miguel Grau condiciona toda decisión a la consulta con la población: el acuerdo de la autoridad no sustituye el de la asamblea. La valoración sigue siendo homogénea (14 de 14 con influencia e interés «Alto») y debe recalibrarse con criterios observables (D-HS05)."],
  "F-DS-02, F-DS-03 · D-HS05."],
 ["adv", "La tenencia es de posesión, salvo en la C.C. de Andanjo",
  ["Nueve de las 11 localidades con F-DS-01 declaran tenencia individual sin título o posesión (La Laguna 10 % titulado, Maray Grande 50 %, Túpac Amaru 60 %); Hualtacal y Almirante Miguel Grau, en la C.C. de Andanjo, declaran tierras comunales. El acuerdo de intervención deberá suscribirse predio por predio en Lalaquiz y Huarmaca, y con la asamblea comunal en Canchaque.",
   "En Ullma la autoridad teme que las zanjas de infiltración perjudiquen al ganado que pastorea en el cerro; en Hualtacal la leña solo se extrae con permiso de la comunidad. Ambas señales anticipan acuerdos de pastoreo y de uso del bosque."],
  "F-DS-01, F-DS-03."],
 ["", "La población nombra los mismos procesos que el DT",
  ["En las 8 entrevistas: disminución del agua en las fuentes en <strong>7</strong>; deforestación, variabilidad del clima y derrumbes o deslizamientos en 6; activación de quebradas en 5; interrupción de vías y plagas en 4; erosión del suelo en 2; sequía (agosto a diciembre, Almirante Miguel Grau) y huaicos en 1.",
   "Las zonas que la población pide conservar son fuentes de agua de consumo —quebradas El Mango, Sambe, Limonal y El Nogal; manantiales El Guayaquil, Palo Espanto, El Guineal, Piedra Colorada y El Higuerón—: base concreta para un MERESE hídrico con las JASS."],
  "F-DS-03, numerales 1.2 a 1.4."],
 ["", "Hay organización y hay proyectos de restauración en el mismo territorio",
  ["Juntas directivas vigentes, JASS que operan reservorios y tratan el agua, rondas campesinas, vaso de leche, club de madres, asociación de ganaderos y comités de regantes; decisiones en asamblea comunal.",
   "El GORE Piura ejecuta el proyecto de regulación hídrica de la microcuenca Pusmalca (CUI 2335868; 744 ha) con viveros en Hualtacal (21,000 plantones; 60 % de mujeres en el personal) y Maraypampa; FONCODES (Haku Wiñay) opera en Almirante Miguel Grau y Hualtacal; la Municipalidad de Lalaquiz ejecuta agua y alcantarillado en Maray Grande y Maray Chico (CUI 2338273). El área de reforestación del GORE en Hualtacal incluye el bloque 81 (D-HS17)."],
  "F-DS-01, F-DS-02 y F-DS-03."],
 ["adv", "Las coordenadas y la microcuenca de las fichas no son de la localidad",
  ["Las 49 fichas repiten el centroide del bloque como coordenada del centro poblado (a menos de 5 m) y todas consignan la microcuenca C1076-Q9584, valor por defecto de la plantilla, aunque los bloques pertenecen a seis microcuencas distintas. La ubicación de los centros poblados queda por determinar.",
   "La duración media de las entrevistas figura como «6 min» en el consolidado; las fichas F-DS-03 dicen «1 hora» (7) y «40 minutos» (Hualtacal): es un error de unidad del aplicativo, no de campo (D-HS06)."],
  "Exportación por bloque, 26/09/2026."],
]

# ---------------- Discrepancias ----------------
DISC = {d['c']: d for d in C['DISC']}
DISC['D-HS01'].update(d="La hoja de control (17/09) reporta 126 F-DS-01 y 55 F-DS-03 aplicadas; el aplicativo (26/09) registra 25 y 8 (13 y 4 al 23/09).",
    e="Las cifras sociales describen 17 localidades y no la provincia; el núcleo de brecha y el distrito de Huancabamba siguen sin caracterizar.")
DISC['D-HS02'].update(m="Varios registros F-DS-01 con los mismos datos de la localidad",
    d="La Laguna: 5 registros de 5 informantes con datos idénticos (difieren solo el destino migratorio y la tarifa de energía); El Papayo: 4; Maray Chico y Ullma: 3; Hualtacal, Almirante Miguel Grau y Pedregal: 2. El nombre de la responsable se escribe con variantes (letras duplicadas, puntos finales) que evaden la regla de duplicidad del aplicativo.",
    e="Inflan la cabecera del consolidado (6,370 hab.; 2,215 familias; Pensión 65 588 personas) y producen coberturas de agua y energía superiores a 100 %.",
    t="Se conserva una ficha por localidad con el valor modal de los campos: 11 F-DS-01 depuradas; el número de informantes se muestra en la ficha de la localidad.",
    a="Regla de unicidad por localidad; si la F-DS-01 se aplica como encuesta a hogares, rediseñar el instrumento.")
DISC['D-HS03'].update(d="El Papayo: hombres + mujeres = 250 frente a 350 declarados; mayores de 65 = 220 (63 %); un registro consigna 1,945 como «población originaria». Almirante Miguel Grau: 110 hombres + 220 mujeres = 330 frente a 230. Pirga: 550 frente a 600. Túpac Amaru: 148 frente a 150.")
DISC['D-HS04'].update(d="Almirante Miguel Grau 230 vs 58 (×3.97); Pedregal 50 vs 14 (×3.57; la población reside en Tunal); Pirga 600 vs 322 (×1.86); Maray Chico 250 vs 142 (×1.76); El Papayo 350 vs 212 (×1.65, cifra de control INEI+SIGRID); Hualtacal 130 vs 80 (×1.62); Ullma 200 vs 157 (×1.27); La Laguna 350 vs 343; Túpac Amaru 150 vs 149. Maray Grande (300 hab.) no tiene cifra INEI.")
DISC['D-HS05'].update(d="Los 14 actores tienen influencia e interés «Alto», posición a favor y nivel comunal.")
DISC['D-HS06'].update(d="El consolidado reporta una duración media de 6 min; las fichas F-DS-03 dicen «1 hora» (7) y «40 minutos» (Hualtacal).",
    t="Se adopta la duración escrita en la ficha.")
DISC['D-HS07'].update(d="Pirca (ficha) / Pirga (INEI); Pedregal (INEI) = anexo La Tuna del C.P. Tunal (ficha); «Yumbe / Cruz Roja» frente a C.P. INEI individuales; Cambrurán / «Camaruran»; Chacchacal = Chorro Blanco; Sapce = Chamelico (Canchaque), según la hoja de control.",
    t="Yumbe y Cruz Roja se tratan como una localidad; las fichas de La Tuna se asignan a Pedregal.")
DISC['D-HS09'].update(m="Fichas duplicadas entre brigadas", n="Bajo",
    d="Hualtacal tiene dos F-DS-02 del mismo agente municipal registradas por responsables distintos (07/08 y 26/09); Almirante Miguel Grau, dos del mismo teniente gobernador. El consolidado cuenta 16 F-DS-02 como 15 vigentes.",
    e="Duplica actores en la cabecera del consolidado.", t="Se cuenta cada actor una vez (14 actores).",
    a="Regla de unicidad por actor y localidad.")
DISC['D-HS10'].update(n="Bajo", d="La brigada constató en campo que Maray Grande está más cerca del bloque 34 que Maray Chico y es población directa; al 26/09 Maray Grande ya tiene F-DS-01, F-DS-02 y F-DS-03. Maray Grande (urbano) no tiene población en el catálogo INEI.",
    e="Resuelto en campo; queda pendiente la cifra de población de referencia.",
    t="Maray Grande se incorpora como localidad caracterizada del bloque 34.",
    a="Obtener la población de Maray Grande (INEI + SIGRID) y revisar el vínculo INEI.")
DISC['D-HS11'].update(d="Las 49 fichas repiten el centroide del bloque como coordenada del centro poblado (a menos de 5 m); todas consignan la microcuenca C1076-Q9584 (valor por defecto) aunque los bloques pertenecen a seis microcuencas (C1076-Q9585, Q9586, Q9587 y Q9593; C1081-Q9590 y Q9591); en Ullma la altitud figura como 126.5 y 1,265 msnm.")
DISC['D-HS12'].update(d="Almirante Miguel Grau declara en F-DS-01 que no tiene ronda y en F-DS-02 que «no hay presidente de la ronda campesina»; Hualtacal, Maray Grande, Pedregal y La Laguna identifican presidente de ronda.",
    t="Se registra la ronda de Almirante Miguel Grau como inexistente o desarticulada.")
DISC['D-HS14'].update(d="El consolidado suma Pensión 65 (588 personas), JUNTOS (419 familias) y Qali Warma (61 II.EE.) sobre los 25 registros F-DS-01, con duplicados; sobre las 11 fichas depuradas los valores son 200, 149 y 36.",
    t="Se usan las cifras depuradas como referencia, no como línea de base.")
DISC['D-HS17'] = {"c": "D-HS17", "m": "Superposición con el proyecto del GORE en Hualtacal", "n": "Alto",
    "d": "La autoridad de Hualtacal declara que el área de reforestación del proyecto del GORE Piura (hualtaco y algarrobo; vivero de 21,000 plantones) aún no se inicia e incluye el bloque 81; Almirante Miguel Grau (79) está en las 744 ha del mismo proyecto de la microcuenca Pusmalca.",
    "e": "Riesgo de doble intervención y de doble conteo del indicador de brecha en los bloques 79 y 81.",
    "t": "Se registra como condicionante social y de articulación interinstitucional.",
    "a": "Solicitar al GORE el polígono de intervención del CUI 2335868 y excluir o concertar la superficie superpuesta."}
C['DISC'] = list(DISC.values())

# ---------------- Riesgos ----------------
R = {r['c']: r for r in C['RIESGOS']}
R['RS-02']['ev'] = "9 de 11 localidades con posesión sin título (La Laguna 10 % titulado); El Papayo y Ullma con constancias de posesión y compraventa; herederos en El Papayo; tierras comunales en la C.C. de Andanjo (Hualtacal, Almirante Miguel Grau)."
R['RS-02']['med'] = "Tamizaje predial y actas de libre disponibilidad con posesionarios y herederos, avaladas por teniente gobernador y ronda; acuerdo de asamblea en la C.C. de Andanjo."
R['RS-04']['ev'] = "Túpac Amaru ofrece 30–40 jornaleros; FONCODES paga S/ 40 por día en Almirante Miguel Grau; el vivero del GORE paga jornales en Hualtacal; La Laguna, Maray Grande y Hualtacal esperan oportunidades laborales."
R['RS-07']['ev'] = "Maray Chico queda aislado en lluvias fuertes (abril); Ullma no pudo trasladarse por huaicos; derrumbes en la carretera de Maray Grande (enero–febrero); El Papayo sufre interrupción de la carretera."
R['RS-08']['ev'] = "Venta de roble y laurel (El Papayo); faique y huayacán para leña (Maray Chico); cedro y huabo (Ullma); chamelico para madera (Almirante Miguel Grau); hualtaco y algarrobo para madera (Hualtacal)."
R['RS-08']['amb'] = "Lalaquiz; Canchaque"
R['RS-09'].update(ev="Proyecto del GORE en la microcuenca Pusmalca (CUI 2335868, 744 ha) con área de reforestación que incluye el bloque 81 y alcanza a Almirante Miguel Grau (79); viveros en Hualtacal y Maraypampa; reforestación en Cruz de Piedra; vivero del GORE junto a M22B1 con pino radiata.",
    p="Alta", i="Alto", n="Alto", med="Convenio de coordinación con el GORE; delimitar la superficie ya intervenida o comprometida; excluir especies exóticas; evitar doble pago de jornales y doble conteo del indicador de brecha.")
R['RS-10']['ev'] = "Almirante Miguel Grau sin presidente de ronda; decisiones condicionadas a la consulta con toda la población."
R['RS-12']['ev'] = "Cruz de Piedra sin cobertura celular; Túpac Amaru solo con celular. Las localidades de Lalaquiz y Canchaque tienen celular e internet."
R['RS-13'] = {"c": "RS-13", "r": "Pérdida de productividad del café por escasez de agua", "ev": "Pedregal: la falta de agua perjudica la cosecha del café y la zona de chacras «ya se seca»; siete de ocho entrevistas reportan menos agua en las fuentes.",
    "amb": "Lalaquiz; Canchaque", "p": "Alta", "i": "Medio",
    "med": "Priorizar la infraestructura natural verde y marrón en las cabeceras de las fuentes de consumo y riego; agroforestería con café bajo sombra en los bordes agrícolas.", "n": "Alto"}
C['RIESGOS'] = list(R.values())

# ---------------- Oportunidades ----------------
C['OPP'] = [
 ["JASS que operan y tratan el agua", "El Papayo, Ullma, La Laguna, Maray Chico, Maray Grande, Pedregal, Almirante Miguel Grau, Cruz de Piedra y Santa Rosa captan de quebradas y manantiales y tratan el agua en reservorios: son los usuarios naturales de un MERESE de regulación hídrica."],
 ["Cultura de pago por el agua", "Los usuarios de canal pagan a la ANA (S/ 20 al año en Almirante Miguel Grau), a la Junta de Regantes de Canchaque (Hualtacal) o al comité de regantes (El Papayo, Maray Chico); las JASS cobran entre S/ 1.70 y S/ 12 al mes: antecedente para la retribución por servicios ecosistémicos."],
 ["Organizaciones productivas", "Cooperativa Agraria Norandino (El Papayo, La Laguna, Maray Grande), Comité de pequeños productores cafetaleros de Ullma (35 años), Asociación de ganaderos comunales y comité de cacaoteros de Hualtacal: vehículo para agroforestería con café y cacao y para acuerdos de manejo del pastoreo."],
 ["Viveros y mano de obra capacitada", "Viveros del GORE en Hualtacal (21,000 plantones; 60 % de mujeres) y Maraypampa, vivero junto a M22B1 y reforestación de FONCODES en Almirante Miguel Grau: capacidades locales instaladas y material vegetal nativo (hualtaco, algarrobo, roble, laurel, aliso)."],
 ["Fuentes que la población quiere conservar", "Quebradas El Mango, Sambe, Limonal y El Nogal; manantiales El Guayaquil, Palo Espanto, El Guineal, Piedra Colorada y El Higuerón: puntos de partida para zonas de protección de fuente."],
 ["Interlocutores colectivos", "C.C. de Andanjo (Canchaque), C.C. San José de Hualcas (Huarmaca), Municipalidad Delegada de El Higuerón (7 caseríos y 2 anexos) y asambleas comunales como espacio de decisión."],
 ["Disposición favorable", "Las 8 autoridades entrevistadas y los 14 actores mapeados están a favor; en Cruz de Piedra la población ya conoce las zanjas de infiltración y en Maray Grande se asocia el proyecto a la prevención de movimientos en masa."],
 ["Horarios de participación definidos", "Las autoridades indican reuniones entre semana desde las 15:00–19:00 h y los fines de semana (sábados por la tarde, domingos por la mañana o al mediodía): base del plan de participación."],
]

# ---------------- Actores institucionales ----------------
I = C['INST']
for r in I:
    if r[0] == 'Gobierno Regional de Piura':
        r[2] = "Hualtacal (81), Almirante Miguel Grau (79), Cruz de Piedra (44), M22B1, El Papayo"
        r[3] = "Proyecto de regulación hídrica de la microcuenca Pusmalca (CUI 2335868; 744 ha) con viveros en Hualtacal y Maraypampa; su área incluye el bloque 81. Vivero junto a M22B1 y complejo educativo en El Papayo."
    if r[0] == 'FONCODES':
        r[2] = "Almirante Miguel Grau (79), Hualtacal (81), Maray Chico (34)"
        r[3] = "Haku Wiñay (Núcleo Ejecutor Canchaque 02): galpones de aves, cocinas mejoradas y agua segura; reforestación con roble y laurel (jornal S/ 40)."
    if r[0] == 'Cooperativa Agraria Norandino':
        r[2] = "El Papayo, La Laguna, Maray Grande"
        r[3] = "Cadena del café; posible socio para agroforestería. En Pedregal la autoridad cuestiona que no brinde asistencia técnica."
    if r[0] == 'ANA y comités de regantes':
        r[2] = "Lalaquiz (canales Santa Ana, Chasqueros, Bentarrona, El Pate y Lanque); Canchaque (Junta de Regantes)"
        r[3] = "Pago por el uso del agua de riego (S/ 20 al año en Almirante Miguel Grau)."
    if r[0] == 'CIPCA':
        r[2] = "El Papayo (37), Pedregal (61)"
        r[3] = "Inducción técnica en educación en 2026 (El Papayo); mejoramiento del estanque de riego hace cinco años (Pedregal)."
    if r[0] == 'PRODERN / AGRORURAL':
        r[2] = "Maray Grande, Hualtacal, Almirante Miguel Grau (PRODERN)"
        r[3] = "Tres F-DS-01 declaran presencia de PRODERN o FONCODES; ninguna, de AGRORURAL."
        r[4] = "Coordinar"
I.insert(3, ["Municipalidad Distrital de Lalaquiz (proyecto de saneamiento)", "Gobierno local", "Maray Grande y Maray Chico (34)",
             "Ejecuta la ampliación del agua potable y la creación del alcantarillado (CUI 2338273; más de S/ 6.5 millones; más de 700 beneficiarios) desde fines de 2024.", "Coordinar"])
I.insert(9, ["Asociación de ganaderos comunales y comité de cacaoteros", "Organización productiva", "Hualtacal (81)",
             "Interlocutores para acuerdos de pastoreo y agroforestería con cacao.", "Involucrar"])

# ---------------- Hipótesis ----------------
C['HIP'] = [
 ["Movimientos en masa (PMM)", "Derrumbes en la carretera y el caserío, pequeños deslizamientos que perjudican chacras y ganado", "6 de 8 (El Papayo, Ullma, La Laguna, Maray Grande, Pedregal, Almirante Miguel Grau)", "Contrastar con F-DS-06 en la interfaz de 19, 37 y M4B4 y en el núcleo de brecha."],
 ["Erosión hídrica (EPH)", "La erosión del suelo va aumentando; el suelo «quema más»", "2 de 8 (Maray Chico, Ullma)", "Relacionar con el inventario de cárcavas pendiente."],
 ["Generación de escorrentía (PGI)", "Crecida y activación de quebradas que se llevan terrenos agrícolas; huaicos que aíslan al caserío", "5 de 8 (El Papayo, Maray Chico, Pedregal, Hualtacal, Almirante Miguel Grau)", "Ubicar los C.P. aguas abajo con coordenadas INEI."],
 ["Regulación hídrica (servicio)", "Menos agua en manantiales y quebradas por la deforestación; sequía de agosto a diciembre", "7 de 8", "Base del balance oferta–demanda y de un MERESE hídrico."],
]

# ---------------- Medidas (fuentes nuevas) ----------------
C['MEDIDAS'][2][1] = "Protección de fuentes de agua de consumo (El Mango, Sambe, El Guayaquil, Palo Espanto, El Guineal, Piedra Colorada, El Higuerón, Limonal, El Nogal) mediante MERESE con las JASS; regulación del cambio de uso en cabecera."
C['MEDIDAS'][1][2] = "Infraestructura natural marrón: control de cárcavas sobre gruss con barreras vivas y diques de baja altura; zanjas de infiltración donde no interfieran con el ganado. Infraestructura natural verde: revegetación con nativas de raíz densa."
C['MEDIDAS'][2][2] = "Infraestructura natural verde: revegetación y enriquecimiento en cabeceras y laderas del régimen A. Infraestructura complementaria: drenaje de trochas y disipadores en quebradas efímeras."

# ---------------- Agenda ----------------
P = {p[0]: p for p in C['PEND']}
P['P-01'][2] = "Digitalizar y reexportar las F-DS-01 y F-DS-03 reportadas en campo (126 y 55 en la hoja de control; 25 y 8 en el aplicativo al 26/09); depurar los registros repetidos por localidad."
P['P-02'][2] = "Levantar el diagnóstico social de los 5 bloques del núcleo de brecha y del distrito de Huancabamba (64, M30B5), que siguen sin ficha al 26/09."
P['P-05'][2] = "Aplicar F-DS-04 a F-DS-07 y el instrumento de saberes ancestrales (0 de 55); en Almirante Miguel Grau, la asamblea es condición previa a cualquier acuerdo."
P['P-12'][2] = "Mover datos personales a un anexo reservado; corregir unidades (duración), valores por defecto (microcuenca, coordenada = centroide) y la regla de duplicidad (variantes del nombre de la responsable) en el aplicativo."
C['PEND'].append(["P-14", "Alta", "Obtener del GORE Piura el polígono del proyecto de la microcuenca Pusmalca (CUI 2335868) y concertar la superficie superpuesta con los bloques 79 y 81 (D-HS17).", "SESDI / especialista social", "Acta de coordinación y polígono conciliado"])

json.dump(D, open('hua_D_new.json', 'w'), ensure_ascii=False)
print('ok')
