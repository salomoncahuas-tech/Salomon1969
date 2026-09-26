"""Reconstrucción de las localidades (LOC) del diagnóstico social a partir del
respaldo del aplicativo IN Piura (corte 26/09/2026).

Solo se recalculan las localidades cuyo conjunto de registros cambió respecto
del corte anterior; los campos de texto curados se conservan salvo que se
entreguen en `curado`.
"""
import json, re, math, unicodedata
from collections import Counter
import pandas as pd

ds = pd.read_pickle('ds.pkl')


def norm(s):
    s = unicodedata.normalize('NFD', str(s or '')).encode('ascii', 'ignore').decode().lower()
    return re.sub(r'[^a-z0-9]', '', s)


def form(r):
    n = r['ficha'].split('-')[-1]
    raw = r.get(f'ds{n}_data_v3')
    try:
        return json.loads(raw) if isinstance(raw, str) and raw.strip() else {}
    except Exception:
        return {}


def txt(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    t = str(v).strip()
    if t.lower() in ('', 'nan', 'none', '-', 's/d', 'sin dato', '.'):
        return None
    return t


def num(v):
    t = txt(v)
    if t is None:
        return None
    t = t.replace(' ', '')
    if re.search(r',\d{3}(\D|$)', t):
        t = t.replace(',', '')
    else:
        t = t.replace(',', '.')
    m = re.search(r'-?\d+(?:\.\d+)?', t)
    if not m:
        return None
    x = float(m.group(0))
    return int(x) if x == int(x) else x


def moda(vals):
    vals = [v for v in vals if v not in (None, '', [], {})]
    if not vals:
        return None
    c = Counter(json.dumps(v, ensure_ascii=False, sort_keys=True) for v in vals)
    return json.loads(c.most_common(1)[0][0])


def nombre_actor(a):
    n = a.get('Nombre del actor / Organizacion') or a.get('Nombre del actor / Organización') or ''
    return norm(n.split('/')[0])


def reg01(recs):
    return [r for r in recs if r['ficha'] == 'F-DS-01']


def reg02(recs):
    return [r for r in recs if r['ficha'] == 'F-DS-02']


def reg03(recs):
    return [r for r in recs if r['ficha'] == 'F-DS-03']


def actores(recs, cargos=None):
    """Actores únicos por nombre (se conserva el registro más reciente)."""
    cargos = cargos or {}
    vistos, out = {}, []
    for r in sorted(reg02(recs), key=lambda r: r['fecha_registro'], reverse=True):
        for a in form(r).get('f2_actores', []) or []:
            k = nombre_actor(a)
            if not k or k in vistos:
                continue
            vistos[k] = 1
            out.append({
                'tipo': txt(a.get('Tipo')),
                'cargo': cargos.get(k) or cargos.get(r['id']) or None,
                'rol': txt(a.get('Rol / Funcion frente al proyecto') or a.get('Rol / Función frente al proyecto')),
                'inf': txt(a.get('Influencia')), 'int': txt(a.get('Interes') or a.get('Interés')),
                'pos': txt(a.get('Posicion') or a.get('Posición')) or 'Sin posición registrada',
                'niv': txt(a.get('Nivel territorial'))})
    return out


def entrevistas(recs, cargo_fmt=None):
    vistos, out = set(), []
    for r in sorted(reg03(recs), key=lambda r: r['fecha_registro'], reverse=True):
        f = form(r)
        k = norm(f.get('f3_nombre') or r.get('nombre_entrevistado'))
        if k in vistos:
            continue
        vistos.add(k)
        g = (txt(f.get('f3_genero')) or '').upper()
        out.append({'gen': 'Hombre' if g.startswith('M') or g.startswith('H') else ('Mujer' if g.startswith('F') else 'Sin dato'),
                    'edad': num(f.get('f3_edad')),
                    'cargo': (cargo_fmt or {}).get(k) or (cargo_fmt or {}).get(r['id']) or txt(f.get('f3_cargo')),
                    'inst': txt(f.get('f3_inst')),
                    'dur': txt(f.get('f3_dur')) or 'Sin dato'})
    return out


PROG_KEYS = ('f1_prodern', 'f1_agrorural', 'f1_otros_proy', 'f1_ongs')


def fragilidad(l):
    ind, sd = [], []
    if l.get('eess') is None: sd.append('eess')
    elif l['eess'] == 'No hay': ind.append('Sin establecimiento de salud')
    san = l.get('san') or ''
    if not san: sd.append('san')
    elif 'Letrina seca' in san or 'Sin saneamiento' in san: ind.append('Letrina seca o sin saneamiento')
    ag = l.get('agua') or ''
    if l.get('aguac') is None and not ag: sd.append('agua')
    elif (l.get('aguac') is not None and l['aguac'] < 80) or (ag and 'JASS' not in ag and 'Manantial' in ag):
        ind.append('Agua < 80 % o de fuente directa')
    if l.get('migra') is None: sd.append('migra')
    elif l['migra'] == 'Alto': ind.append('Migración alta')
    edu = l.get('edu') or ''
    if not edu: sd.append('edu')
    elif 'Primaria' in edu or 'Sin nivel' in edu: ind.append('Educación primaria o menor')
    ie = l.get('ie') or ''
    if not ie: sd.append('ie')
    elif ie.startswith('No hay'): ind.append('Sin institución educativa')
    n = len(ind)
    return ind, ('Alta' if n >= 4 else ('Media' if n >= 2 else 'Baja')), sd


def resiliencia(l):
    ind = []
    if l.get('junta') == 'Sí': ind.append('Junta directiva vigente')
    if l.get('jass') is True: ind.append('JASS operativa')
    if l.get('ronda') is True: ind.append('Ronda campesina')
    if l.get('cc') and not str(l['cc']).startswith('Ninguna'): ind.append('Comunidad campesina')
    if l.get('programa'): ind.append('Programa o institución presente')
    if l.get('cad') and not str(l['cad']).startswith('Ninguna'): ind.append('Organización productiva')
    n = len(ind)
    return ind, ('Alta' if n >= 4 else ('Media' if n >= 2 else 'Baja'))


def pct(v):
    x = num(v)
    return None if x is None else int(round(x))


def reconstruir(l, recs, curado=None):
    """Actualiza la entrada LOC `l` con los registros `recs` (lista de dict)."""
    curado = curado or {}
    r1, r2, r3 = reg01(recs), reg02(recs), reg03(recs)
    act = actores(recs, curado.pop('_cargos', None))
    ent = entrevistas(recs, curado.pop('_cargos3', None))
    n01 = 1 if r1 else 0
    n02 = len(act)
    n03 = len(ent)
    l['n01'], l['n02'], l['n03'] = n01, n02, n03
    l['nreg'] = len(recs)
    l['nrep'] = len(recs) - (n01 + n02 + n03)
    l['ninf01'] = len({norm(r.get('nombre_entrevistado')) or r['id'] for r in r1})
    l['bl_f'] = sorted({str(r['bloque']) for r in recs})
    if r1:
        F = [form(r) for r in r1]
        g = lambda k: moda([txt(f.get(k)) for f in F])
        l['fam'] = num(g('f1_nfam'))
        l['pvals'] = [num(f.get('f1_pob_t')) for f in F if num(f.get('f1_pob_t')) is not None]
        l['pdec'] = moda(l['pvals'])
        l['hdec'] = num(g('f1_pob_h')); l['mdec'] = num(g('f1_pob_m'))
        l['men18'] = num(g('f1_pob_men18')); l['may65'] = num(g('f1_pob_may65'))
        l['idioma'] = g('f1_idioma'); l['edu'] = g('f1_nivel_edu')
        l['migra'] = g('f1_migracion'); l['dest'] = g('f1_destino_mig')
        l['org'] = g('f1_org_terr'); l['junta'] = g('f1_junta_vig')
        ag = moda([sorted(f.get('f1_agua') or []) for f in F]) or []
        l['agua'] = ' / '.join(ag) or None
        l['aguac'] = pct(g('f1_agua_cob'))
        l['san'] = ' / '.join(moda([f.get('f1_sanea') or [] for f in F]) or []) or None
        l['ener'] = ' / '.join(moda([f.get('f1_energia') or [] for f in F]) or []) or None
        l['enerc'] = pct(g('f1_energia_cob'))
        l['tel'] = g('f1_telecom'); l['eess'] = g('f1_eess'); l['ie'] = g('f1_ie_niveles')
        l['dsalud'] = curado.pop('dsalud', l.get('dsalud'))
        l['f1'] = moda([str(r['fecha_evaluacion'])[:10] for r in r1])
        l['alt'] = num(moda([r.get('altitud') for r in r1]))
        # Actividades económicas
        filas = {}
        for f in F:
            for a in f.get('f1_activ') or []:
                k = txt(a.get('Actividad / Rubro'))
                if k:
                    filas.setdefault(k, []).append(a)
        l['acts'] = {k: num(moda([txt(a.get('N fam.')) for a in v])) for k, v in filas.items()}
        l['destino'] = {k: moda([txt(a.get('Destino')) for a in v]) for k, v in filas.items()}
        l['ingr'] = {k: num(moda([txt(a.get('Ingreso (S/./mes)')) for a in v])) for k, v in filas.items()}
        l['prods'] = sorted({(txt(a.get('Productos principales')) or '').lower().strip()
                             for v in filas.values() for a in v} - {''})
        ronda = g('f1_ronda')
        l['ronda'] = True if ronda == 'Sí' else (l.get('ronda') if ronda is None else False)
        pres = ' '.join(txt(f.get('f1_pres_junta')) or '' for f in F)
        l['jass'] = ('JASS' in (l['agua'] or '')) or ('JASS' in pres.upper()) or bool(l.get('jass'))
        l['programa_si'] = any((txt(f.get(k)) or '') == 'Sí' for f in F for k in PROG_KEYS)
        prog = {}
        for k, lab in (('f1_juntos', 'JUNTOS'), ('f1_pension65', 'Pensión 65'), ('f1_qaliwarma', 'Qali Warma'), ('f1_beca18', 'Beca 18')):
            v = num(g(k))
            if v: prog[lab] = v
        l['progsoc'] = prog
        l['pestatal'] = g('f1_presencia_estatal')
        l['tenencia'] = g('f1_tenencia')
        l['ptit'] = pct(g('f1_pct_tituladas'))
    if r3:
        l['f2'] = moda([str(r['fecha_evaluacion'])[:10] for r in r3])
        acs = [txt(form(r).get('f3_r_acuerdo')) or '' for r in r3]
        l['acuerdo'] = sum(1 for a in acs if a.lower().startswith(('sí', 'si', 'el como autoridad se encuentra de acuerdo')) or 'de acuerdo' in a.lower())
        l['sinresp'] = sum(1 for a in acs if not a)
    l['actors'] = act
    l['entrev'] = ent
    for k, v in curado.items():
        l[k] = v
    if r1:
        l['frag_ind'], l['frag'], l['sd_ind'] = fragilidad(l)
        l['res_ind'], l['res'] = resiliencia(l)
    for k in ('programa_si',):
        l.pop(k, None)
    return l
