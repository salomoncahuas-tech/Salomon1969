import json
exec(open('dsbuild.py').read())

D = json.load(open('hua_D.json'))
LOC = D['LOC']
LL = {l['loc']: l for l in LOC}
ALIAS = {'Pirca': 'Pirga', 'Tupac Amaru': 'Túpac Amaru', 'Coyona': 'Coyona (S. M. de El Faique)',
         'El Higueron': 'El Higuerón'}
CAMBIADAS = ['La Laguna', 'Pedregal', 'Maray Chico', 'Maray Grande', 'Hualtacal', 'Almirante Miguel Grau']

h = ds[ds.provincia == 'Huancabamba']
recs_por = {}
for _, r in h.iterrows():
    loc = ALIAS.get(r.centro_poblado, r.centro_poblado)
    assert loc in LL, loc
    recs_por.setdefault(loc, []).append(r.to_dict())

AM, TG = 'Agente municipal', 'Teniente gobernador'
CUR = {
 'Pedregal': dict(
   cc='Ninguna', programa=None,
   cad='Ninguna (venta independiente de café; la autoridad cuestiona a la Cooperativa Norandino)',
   orgs=['JASS', 'Ronda campesina', 'Vaso de leche', 'Comité de riego La Tuna'],
   decide='Asamblea comunal', cond='Lunes a viernes desde las 19:00 h',
   meses='Desde febrero (derrumbes y activación de quebradas)',
   hz=['Disminución del agua en fuentes', 'Variabilidad del clima', 'Derrumbes y deslizamientos',
       'Crecida / activación de quebradas', 'Plagas en cultivos'],
   spp=['Pajul', 'Higuerón', 'Nogal', 'Checche', 'Huayacán'], sppdis=['Checche'],
   ben=['Beneficio para la población'],
   quote='Sí, es importante que prioricen la zona agraria, porque producen mejor la tierra. Sería beneficioso para todos los pobladores.',
   cons=['Quebrada Sambe (captación del sistema de agua potable)', 'Vertiente y estanque La Huaca (riego por gravedad)'],
   otros='Antes caserío Pedregal, hoy anexo La Tuna del C.P. Tunal: la población reside en Tunal y el anexo funciona como zona de chacras, por lo que la población presente es muy reducida. La leña se ha reemplazado por gas.',
   obs01='La falta de agua perjudica la cosecha del café. Los productores venden el café de forma independiente y consideran que la Cooperativa Norandino no brinda asistencia técnica. Los pobladores atribuyen al higuerón y al pajul la capacidad de «dar agua».',
   obs02='El agente municipal concentra la decisión; ronda campesina identificada. No se ha desarrollado ninguna plataforma de concertación.',
   _cargos={259: AM}, _cargos3={254: 'Agente municipal / agricultor'}),
 'Maray Grande': dict(
   cc='Ninguna', programa='PRODERN; Municipalidad de Lalaquiz (agua y alcantarillado, CUI 2338273)',
   cad='Cooperativa Norandino',
   orgs=['Vaso de leche', 'Ronda campesina', 'JASS', 'Club de madres'],
   decide='JASS (asuntos del agua)', cond='Lunes a viernes desde las 16:00 h',
   meses='Enero y febrero (derrumbes en el caserío y la carretera)',
   hz=['Disminución del agua en fuentes', 'Variabilidad del clima', 'Deforestación / venta de madera',
       'Derrumbes y deslizamientos', 'Interrupción de vías y aislamiento'],
   spp=['Higuerón', 'Huabo', 'Chirimoyo', 'Lúcumo'], sppdis=['Lúcumo', 'Higuerón'],
   ben=['Prevención de desastres', 'Empleo / mano de obra local'],
   quote='Sí, porque beneficiaría a la población y la prevendría de algún movimiento de masas por las fuertes lluvias.',
   cons=['Manantiales El Guineal y Piedra Colorada (captación para consumo)', 'Sector Cerro'],
   otros='Riego por los canales El Pate y Lanque hacia un reservorio exclusivo. Leña de palo seco. La brigada verificó en campo que Maray Grande está más cerca del bloque 34 que Maray Chico: es población directa.',
   obs01='Proyecto de agua potable y alcantarillado de Maray Grande y Maray Chico en ejecución por la Municipalidad de Lalaquiz (CUI 2338273; más de S/ 6.5 millones; más de 700 beneficiarios). Café y maíz una vez al año; guineo para consumo; ganado como «caja chica».',
   obs02='El agente municipal concentra la decisión; presidente de ronda campesina identificado. Aún no se ha realizado una plataforma de concertación.',
   _cargos={276: AM}, _cargos3={275: 'Agente municipal / agricultor'}),
 'Hualtacal': dict(
   cc='C.C. Andanjo',
   programa='GORE Piura (vivero satélite y reforestación); FONCODES (cocinas mejoradas, biohuertos, gallinas ponedoras)',
   cad='Asociación de ganaderos comunales; comité de cacaoteros',
   orgs=['Vaso de leche', 'Ronda campesina', 'JASS', 'Asociación de ganaderos', 'Comité de cacaoteros'],
   decide='Reuniones con la comunidad',
   cond='Lunes a viernes desde las 15:00 h; sábados desde las 14:00 h; domingos a cualquier hora',
   meses='Temporada de lluvias fuertes (activación de quebradas)',
   hz=['Crecida / activación de quebradas', 'Variabilidad del clima'],
   spp=['Palo santo', 'Hualtaco', 'Charán', 'Algarrobo', 'Parnaso', 'Angolo', 'Faique', 'Ceibo'], sppdis=[],
   ben=['Recuperación de áreas verdes', 'Empleo / mano de obra local'],
   quote='Sí me encuentro de acuerdo, para que el bosque se encuentre conservado y haya oportunidad de trabajo a la población.',
   cons=['Manantial El Higuerón (captación diaria para consumo)', 'Bosques del caserío, para reforestar las áreas perdidas'],
   otros='Agua de consumo sin clorar (la población reporta que el cloro afectó a los niños). Riego a cargo de la Junta de Regantes de Canchaque. Puente de 60 m (ARCC, 2018–2023) que asegura la salida de la producción. Potencial turístico («los peroles»).',
   obs01='Vivero forestal satélite del GORE Piura con meta de 21,000 plantones; el 60 % del personal son mujeres del caserío. Hualtacal integra las 744 ha priorizadas del proyecto de regulación hídrica de Pusmalca; la reforestación con hualtaco y algarrobo aún no inicia y su área incluye el bloque 81.',
   obs02='Se identificaron agente municipal, teniente gobernador y presidente de ronda. Aún no se han realizado plataformas de concertación. C.C. Andanjo.',
   _cargos={286: AM}, _cargos3={283: 'Agente municipal / agricultor'}),
 'Almirante Miguel Grau': dict(
   cc='C.C. de Andanjo',
   programa='FONCODES – Haku Wiñay (galpones, cocinas mejoradas, agua segura); GORE Piura (proyecto Pusmalca, CUI 2335868)',
   cad='Ninguna', orgs=['JASS'], decide='Asamblea con toda la comunidad',
   cond='Sábados y domingos desde el mediodía',
   meses='Agosto a diciembre (sequía); lluvias fuertes (activación de quebradas)',
   hz=['Disminución del agua en fuentes', 'Deforestación / venta de madera', 'Crecida / activación de quebradas',
       'Derrumbes y deslizamientos', 'Sequía', 'Variabilidad del clima'],
   spp=['Faique', 'Ceibo', 'Pasayo', 'Chamelico', 'Chirimoyo', 'Lúcumo', 'Higuerón'], sppdis=['Chamelico', 'Higuerón'],
   ben=['Empleo / mano de obra local'],
   quote='Como autoridad me encuentro de acuerdo, pero antes de cualquier decisión se tiene que consultar a toda la población.',
   cons=['Manantial El Guayaquil (consumo humano)'],
   otros='Riego desde quebradas locales con pago anual de S/ 20 a la ANA. Leña de bosque seco. El teniente gobernador atribuye la disminución del manantial a la pérdida del higuerón.',
   obs01='Café, yuca, plátano, maíz y menestra para venta; ganado como reserva ante emergencias. FONCODES (Haku Wiñay, Núcleo Ejecutor Canchaque 02) y proyecto del GORE en la microcuenca Pusmalca (744 ha con roble, laurel, aliso y saúco; vivero en Maraypampa con jornales para mujeres).',
   obs02='Ronda campesina desarticulada (sin presidente). El teniente gobernador, que preside la JASS, condiciona toda decisión a la consulta con la población. C.C. de Andanjo.',
   _cargos={290: TG}, _cargos3={287: 'Teniente gobernador / presidente de la JASS'}),
 'La Laguna': dict(
   programa=None,
   obs01='Café para venta (una cosecha al año; trabajan hombres y mujeres); plátano, yuca, cebolla y frutas para consumo. Agua de consumo de la quebrada El Nogal por gravedad (sistema de 2017); riego por canales tradicionales. Cinco informantes F-DS-01 con los mismos datos de la localidad.',
   obs02='Presidente de la ronda campesina entrevistado y mapeado; teniente gobernador y agente municipal identificados en F-DS-01.'),
 'Maray Chico': dict(),
}

for loc in CAMBIADAS:
    l = LL[loc]
    prev = (l['n01'], l['n02'], l['n03'], l['nreg'])
    reconstruir(l, recs_por[loc], dict(CUR[loc]))
    print(loc, prev, '->', (l['n01'], l['n02'], l['n03'], l['nreg'], l['nrep']), l.get('frag'), l.get('res'), l['frag_ind'], l['res_ind'])

# Campos adicionales (informantes, programas sociales, tenencia) para todas las localidades con registros
for loc, rs in recs_por.items():
    l = LL[loc]
    if loc in CAMBIADAS:
        continue
    tmp = json.loads(json.dumps(l))
    reconstruir(tmp, rs, {})
    for k in ('n01', 'n02', 'n03', 'nreg'):
        if tmp[k] != l[k]:
            print('  AVISO conteo', loc, k, l[k], '->', tmp[k])
    for k in ('ninf01', 'progsoc', 'pestatal', 'tenencia', 'ptit'):
        if k in tmp:
            l[k] = tmp[k]

json.dump(D, open('hua_D_new.json', 'w'), ensure_ascii=False)
