import json, statistics, copy
exec(open('mor_map.py').read())
OLD = copy.deepcopy(D)
V6 = {b['b'] for b in D['B']}

def plain(s):
    return unicodedata.normalize('NFD', str(s or '')).encode('ascii', 'ignore').decode().lower()

HZ = [('Sequía / déficit hídrico', r'sequ|escasez de(l)? agua|falta de agua|menos agua|secando|se seca|deficit|poca agua|no llueve|disminu\w* (del|de) agua|agua ha disminuido|escasez hidric'),
      ('Deforestación / tala', r'deforest|\btala\b|talan|talado|lucra con la venta|venta de (madera|lena|arbol)|extraccion de (madera|lena)'),
      ('Aumento de temperatura / cambio climático', r'calor|temperatura|cambio climatic|clima'),
      ('Erosión / degradación del suelo', r'erosion|agriet|suelo (se )?(degrad|empobrec|pobre)|degradacion del suelo'),
      ('Deslizamientos / movimientos en masa', r'desliz|derrumb|movimiento(s)? de masa|desprend'),
      ('Filtración de agua en cerros', r'filtr'),
      ('Activación de quebradas / inundación', r'quebrada\w*[^.]{0,40}(activ|crec|desbord)|activ\w*[^.]{0,30}quebrada|inund|desbord|crecida'),
      ('Huaicos / flujos', r'huaic|huayc|lodo'),
      ('Plagas', r'plaga')]
SPP = [('Palo santo', r'palo ?santo'), ('Hualtaco', r'hualtaco'), ('Algarrobo', r'algarrob'), ('Sapote', r'[sz]apote'),
       ('Faique', r'faique'), ('Overal', r'overal|overo'), ('Ceibo', r'ceibo'), ('Charán', r'charan'), ('Guayacán', r'guayacan'),
       ('Pasallo', r'pasallo|pasayo'), ('Higuerón', r'higueron'), ('Nogal', r'nogal'), ('Cedro', r'cedro')]
BEN = [('Aire / ambiente', r'ambiente|\baire\b|oxigen'), ('Empleo / mano de obra local', r'trabajo|empleo|ingreso|jornal|laboral'),
       ('Recuperación del bosque y áreas verdes', r'bosque|reforest|arbol|verde|recuper|forestal'),
       ('Prevención de desastres', r'preven|desastre|huaic|riesgo|barrera|desborde'), ('Agua', r'\bagua\b')]

def code(text, rules):
    t = plain(text)
    return [k for k, p in rules if re.search(p, t)]

def dur_min(f, r):
    t = plain(f.get('f3_dur'))
    x = num(t)
    if x is not None:
        if 'hora' in t:
            return int(x * 60)
        return int(x)
    return num(r.get('ds03_duracion'))

def aslist(v):
    return v if isinstance(v, list) else ([x.strip() for x in str(v).split(';') if x.strip()] if v else [])

def cp_key(r):
    return norm(r['centro_poblado']) if isinstance(r['centro_poblado'], str) else norm(locname(r))

def rebuild(l, rs):
    r1, r2, r3 = reg01(rs), reg02(rs), reg03(rs)
    # F-DS-01: una depurada por centro poblado declarado (los nombres agrupados «A / B» se omiten si hay nombres individuales)
    cps = {}
    for r in r1:
        cps.setdefault(cp_key(r), []).append(r)
    indiv = [k for k in cps if '/' not in str(cps[k][0]['centro_poblado'])]
    if indiv and len(cps) > len(indiv):
        for k in list(cps):
            if k not in indiv:
                cps[indiv[0]] += cps.pop(k)
    n01 = len(cps)
    # F-DS-02: fichas únicas (conjunto de actores); actores únicos por nombre
    sets = {tuple(sorted(nombre_actor(a) for a in form(r).get('f2_actores') or [])) or (r['id'],) for r in r2}
    n02 = len(sets)
    act = actores(rs)
    ent_recs = {}
    for r in sorted(r3, key=lambda r: r['fecha_registro'], reverse=True):
        k = norm(form(r).get('f3_nombre') or r.get('nombre_entrevistado'))
        ent_recs.setdefault(k, r)
    n03 = len(ent_recs)
    l.update(n01=n01, n02=n02, n03=n03, nreg=len(rs), nrep=len(rs) - (n01 + n02 + n03))
    l['bl_f'] = sorted({str(r['bloque']) for r in rs})
    l['ninf01'] = len({norm(r.get('nombre_entrevistado')) or r['id'] for r in r1})
    if r1:
        F = [form(r) for r in r1]
        g = lambda k: moda([txt(f.get(k)) for f in F])
        med = []
        pv = []
        for k, rr in cps.items():
            vals = [num(form(r).get('f1_pob_t')) for r in rr if num(form(r).get('f1_pob_t')) is not None]
            pv += vals
            if vals:
                med.append(statistics.median(vals))
        l['pvals'] = sorted(set(pv))
        l['pdec'] = float(statistics.median(l['pvals'])) if l['pvals'] else None
        fams = []
        for k, rr in cps.items():
            v = [num(form(r).get('f1_nfam')) for r in rr if num(form(r).get('f1_nfam')) is not None]
            if v: fams.append(statistics.median(v))
        fv = sorted({num(f.get('f1_nfam')) for f in F if num(f.get('f1_nfam')) is not None})
        l['fam'] = float(statistics.median(fv)) if fv else None
        l['org'] = g('f1_org_terr'); l['junta'] = g('f1_junta_vig'); l['idioma'] = g('f1_idioma')
        l['edu'] = g('f1_nivel_edu'); l['migra'] = g('f1_migracion'); l['dest'] = g('f1_destino_mig')
        l['agua'] = ' / '.join(moda([sorted(aslist(f.get('f1_agua'))) for f in F]) or []) or None
        l['aguac'] = None if pct(g('f1_agua_cob')) is None else float(pct(g('f1_agua_cob')))
        l['san'] = ' / '.join(moda([aslist(f.get('f1_sanea')) for f in F]) or []) or None
        l['ener'] = ' / '.join(moda([aslist(f.get('f1_energia')) for f in F]) or []) or None
        l['enerc'] = None if pct(g('f1_energia_cob')) is None else float(pct(g('f1_energia_cob')))
        l['tel'] = g('f1_telecom'); l['eess'] = g('f1_eess'); l['ie'] = g('f1_ie_niveles'); l['dsalud'] = g('f1_eess_dist')
        l['men18'] = num(g('f1_pob_men18')); l['may65'] = num(g('f1_pob_may65'))
        filas = {}
        for f in F:
            for a in f.get('f1_activ') or []:
                k = txt(a.get('Actividad / Rubro'))
                if k: filas.setdefault(k, []).append(a)
        l['acts'] = {k: num(moda([txt(a.get('N fam.')) for a in v])) for k, v in filas.items()}
        l['destino'] = {k: moda([txt(a.get('Destino')) for a in v]) for k, v in filas.items()}
        l['prods'] = sorted({(txt(a.get('Productos principales')) or '').lower().strip() for v in filas.values() for a in v} - {''})
        l['f1'] = min(str(r['fecha_evaluacion'])[:10] for r in r1)
        prog = {}
        for k, lab in (('f1_juntos', 'JUNTOS'), ('f1_pension65', 'Pensión 65'), ('f1_qaliwarma', 'Qali Warma'), ('f1_beca18', 'Beca 18')):
            v = num(g(k))
            if v: prog[lab] = v
        l['progsoc'] = prog
        l['tenencia'] = g('f1_tenencia'); l['ptit'] = pct(g('f1_pct_tituladas'))
    cc = moda([txt(r.get('comunidad_campesina')) for r in rs])
    if cc is not None or r1:
        l['cc'] = cc
    l['actors'] = act
    l['entrev'] = []
    for k, r in ent_recs.items():
        f = form(r); gg = (txt(f.get('f3_genero')) or '').upper()
        l['entrev'].append({'gen': 'Hombre' if gg[:1] in ('M', 'H') else ('Mujer' if gg[:1] == 'F' else 'Sin dato'),
                            'edad': num(f.get('f3_edad')), 'cargo': txt(f.get('f3_cargo')), 'inst': txt(f.get('f3_inst')),
                            'min': dur_min(f, r)})
    if r3:
        F3 = [form(r) for r in ent_recs.values()]
        l['hz'] = sorted({h for f in F3 for h in code((f.get('f3_r1') or '') + ' ' + (f.get('f3_r2') or '') + ' ' + (f.get('f3_r3') or ''), HZ)}, key=[k for k, _ in HZ].index)
        l['spp'] = sorted({h for f in F3 for h in code((f.get('f3_r1') or '') + ' ' + (f.get('f3_r2') or ''), SPP)}, key=[k for k, _ in SPP].index)
        l['ben'] = sorted({h for f in F3 for h in code((f.get('f3_r_acuerdo') or '') + ' ' + (f.get('f3_cierre') or ''), BEN)}, key=[k for k, _ in BEN].index)
        acs = [plain(f.get('f3_r_acuerdo')) for f in F3]
        l['acuerdo'] = sum(1 for a in acs if a.startswith('si') or 'de acuerdo' in a or 'bien' in a[:30])
        l['sinresp'] = sum(1 for a in acs if not a.strip())
        q = [txt(f.get('f3_r_acuerdo')) for f in F3 if txt(f.get('f3_r_acuerdo'))]
        l['quote'] = max(q, key=len) if q else None
        cons = []
        for f in F3:
            c = txt(f.get('f3_r4'))
            if c and c[:200] not in cons:
                cons.append(c[:200])
        l['cons'] = cons
        dec = ' '.join(plain(f.get('f3_r10')) for f in F3)
        l['decide'] = 'Asamblea comunal' if 'asamblea' in dec else ('Reuniones con la comunidad' if 'reunion' in dec else (txt(F3[0].get('f3_r10')) or '')[:80])
        l['f2'] = max(str(r['fecha_evaluacion'])[:10] for r in r3)
    alltext = plain(' '.join(json.dumps(form(r), ensure_ascii=False) + ' ' + str(r.get('observaciones_generales') or '') for r in rs))
    l['foncodes'] = bool(re.search(r'foncodes|haku', alltext)) or l.get('foncodes', False)
    l['cipca'] = 'cipca' in alltext or l.get('cipca', False)
    l['serfor'] = 'serfor' in alltext or l.get('serfor', False)
    l['acr'] = bool(re.search(r'\bacr\b|area de conservacion', alltext)) or l.get('acr', False)
    ronda = any((txt(form(r).get('f1_ronda')) or '') == 'Sí' for r in r1)
    l['ronda'] = bool(ronda)
    l['jass'] = 'JASS' in (l.get('agua') or '') or 'jass' in alltext
    progsi = any((txt(form(r).get(k)) or '') == 'Sí' for r in r1 for k in PROG_KEYS)
    if r1:
        fi = []
        if l['eess'] == 'No hay': fi.append('Sin establecimiento de salud')
        if l['san'] and ('Letrina seca' in l['san'] or 'Sin saneamiento' in l['san']): fi.append('Letrina seca / sin saneamiento')
        if (l['aguac'] is not None and l['aguac'] < 80) or (l['agua'] and 'JASS' not in l['agua'] and 'Manantial' in l['agua']): fi.append('Agua < 80 % o fuente directa')
        if l['migra'] == 'Alto': fi.append('Migración alta')
        if l['edu'] and ('Primaria' in l['edu'] or 'Sin nivel' in l['edu']): fi.append('Nivel educativo primaria o menor')
        if l['ie'] and l['ie'].startswith('No hay'): fi.append('Sin institución educativa')
        l['frag_ind'] = fi; l['frag'] = 'Alta' if len(fi) >= 4 else ('Media' if len(fi) >= 2 else 'Baja')
        ri = []
        if l['junta'] == 'Sí': ri.append('Junta directiva vigente')
        if l['jass']: ri.append('JASS')
        if l['ronda']: ri.append('Ronda campesina mencionada')
        if l.get('cc') and not re.match(r'^(ningun|no |-|nan)', plain(l['cc'])): ri.append('Comunidad campesina')
        if l['foncodes'] or l['cipca'] or progsi or re.search(r'\bong\b', alltext): ri.append('Programa o ONG presente')
        if l['acr']: ri.append('Vínculo con ACR')
        l['res_ind'] = ri; l['res'] = 'Alta' if len(ri) >= 4 else ('Media' if len(ri) >= 2 else 'Baja')
    l.setdefault('v5only', False)
    return l

for n, rs in recs_por.items():
    rebuild(LL[n], rs)

if __name__ == '__main__':
    # Validación contra el corte anterior en localidades sin registros nuevos ni adicionales
    OL = {l['loc']: l for l in OLD['LOC']}
    same = diff = 0
    for n, rs in recs_por.items():
        o, l = OL[n], LL[n]
        if o['nreg'] != len(rs):
            continue
        for k in ('hz', 'spp', 'ben', 'frag', 'res', 'pdec', 'fam', 'aguac', 'san', 'eess'):
            if (sorted(o[k]) if isinstance(o[k], list) else o[k]) == (sorted(l[k]) if isinstance(l[k], list) else l[k]):
                same += 1
            else:
                diff += 1
                print('DIF', n, k, o[k], '->', l[k])
    print('iguales', same, 'distintos', diff)
    json.dump(D, open('mor_D_new.json', 'w'), ensure_ascii=False)
