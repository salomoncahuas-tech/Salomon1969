"""Conjunto integrado por bloque: E6 (ecosistemas), SP (situación predial), PI/INT, F-DT (DATA), social (UPD 17/10)."""
import json
from collections import Counter, defaultdict
PM = {'A': 'Ayabaca', 'H': 'Huancabamba', 'M': 'Morropón'}
E6 = json.load(open('work/E6_D.json'))['b']
SP = json.load(open('work/SP_index.json'))

def grab(s, var):
    i = s.index(var) + len(var)
    return json.JSONDecoder().raw_decode(s[i:])[0]

def tab(U, k, pref):
    for t in U['ds'].get(k, []):
        if t['t'].split('. ', 1)[1].startswith(pref): return t

def norm_reg(r):
    if not r: return None
    r = r.lower()
    if r.startswith('comunal'): return 'Comunal'
    if r.startswith('estatal'): return 'Estatal'
    if 'titulada' in r: return 'Privada'
    if 'sin titular' in r or 'posesi' in r or 'trámite' in r: return 'Posesión / sin título'
    return 'Otro'

def build(p):
    prov = PM[p]
    s = open(f'src/{p}.html').read()
    DATA = grab(s, 'const DATA = ') if p == 'A' else grab(s, 'const D=')
    INT = json.load(open(f'work/{p}_INT.json'))
    U = json.load(open(f'work/{p}_new.json'))
    UO = json.load(open(f'work/{p}_UPD_old.json'))
    dB = {(b.get('cod') or b.get('b')): b for b in DATA['B']}
    iB = {b['b']: b for b in INT['B']}
    eB = {b['bloque']: b for b in E6 if b['provincia'] == prov}
    sB = {b['b']: b for b in SP if b['prov'] == prov}
    D = tab(U, 'cob', 'Fichas sociales por bloque'); nf = {str(r[0]): r[1] for r in D['r']}
    ten = U['T']['ten']; th = ten['h']; tB = {str(r[0]): dict(zip(th, r)) for r in ten['r']}
    eco_v6 = UO['eco']
    X = []
    for b in iB:
        e, sp, it, d = eB[b], sB[b], iB[b], dB.get(b, {})
        t = tB.get(b)
        v6 = eco_v6.get(b, {}).get('e', [])
        v = it.get('ipp')
        picls = None if v is None else ('Muy bajo' if v < 1.8 else 'Bajo' if v < 2.6 else 'Medio' if v < 3.4 else 'Alto' if v < 4.2 else 'Muy alto')
        rec = dict(b=b, d=e['distrito'], ha=e['area_ha'], pi=picls, ipp=it.get('ipp'), reg=it.get('reg'),
                   eco=e['ecosistema'], v6=(v6[0][0] if v6 else None), v6p=(v6[0][3] if v6 else None),
                   ic=e['ic_bloque'], cls=e['clase_estado'], bu=e['msavi_bajo_umbral_pct'], buha=e['msavi_bajo_umbral_ha'],
                   ra=e['msavi_rest_activa_ha'], tip=e['tipologia_S2DR4'], ev=e['clase_evidencia'],
                   dra=e['delim_rest_activa_ha'], den=e['delim_enriquecimiento_ha'], dco=e['delim_conservacion_ha'], dag=e['delim_tamizaje_agricola_ha'],
                   cons=d.get('cons') or d.get('s_deg'), dteco=d.get('eco') or d.get('s_eco'),
                   sit=sp['sit'], pcc=sp['pcc'], pcof=sp['pcof'], pest=sp['pest'], cc=sp['cc'], cof=sp['cof'], est=sp['est'],
                   ccn=sp['ccn'], npc=sp['npc'], nccpp=sp['nccpp'], ccpp=sp['ccpp'], pob=sp['pob'], asis=sp['asis'], sp=sp['sp'], img=sp['img'], esc=sp['esc'],
                   nf=nf.get(b, 0), regd=(t or {}).get('Régimen predominante de tenencia'), npd=(t or {}).get('N.° aprox. de predios individuales'),
                   supcc=(t or {}).get('Se superpone a tierras comunales'), lind=(t or {}).get('Conflictos de linderos'), tit=(t or {}).get('Tierras tituladas (%)'))
        # cruces
        f = []
        rn = norm_reg(rec['regd'])
        if rn == 'Comunal' and rec['pcc'] < 5:
            f.append(['ten', 'F-DS-01 declara régimen comunal; el gabinete no registra comunidad campesina en el bloque'])
        if rn == 'Estatal' and rec['pcc'] >= 50:
            f.append(['ten', f"F-DS-01 declara régimen estatal; el gabinete ubica {rec['pcc']:.0f} % del bloque en comunidad campesina"])
        if rn in ('Privada', 'Posesión / sin título') and rec['pcc'] >= 50:
            f.append(['ten', f"F-DS-01 declara régimen individual; el gabinete ubica {rec['pcc']:.0f} % del bloque en comunidad campesina"])
        if rec['supcc'] == 'No' and rec['pcc'] >= 5:
            f.append(['ten', f"F-DS-01 responde que no se superpone a tierras comunales; el gabinete registra {rec['pcc']:.0f} % en C. C."])
        if rec['supcc'] == 'Sí' and rec['pcc'] < 5:
            f.append(['ten', 'F-DS-01 responde que se superpone a tierras comunales; el gabinete no registra C. C.'])
        if (rec['npc'] or 0) > 19:
            f.append(['p5', f"{rec['npc']} predios COFOPRI interceptados: supera el máximo de 19 del Paso 5"])
        if (rec['npd'] or 0) > 19:
            f.append(['p5', f"F-DS-01 declara ≈ {rec['npd']} predios individuales: supera el máximo de 19 del Paso 5"])
        if rec['nf'] == 0:
            f.append(['soc', 'Sin ficha social: régimen y titulares no verificados con la población'])
        rec['f'] = f
        X.append(rec)
    return X

if __name__ == '__main__':
    for p in PM:
        X = build(p)
        json.dump(X, open(f'work/{p}_X.json', 'w'), ensure_ascii=False)
        print('#####', p, len(X))
        c = Counter(); ha = defaultdict(float)
        for r in X: c[r['cls']] += 1; ha[r['cls']] += r['ha']
        print(' clases IC', dict(c), {k: round(v, 1) for k, v in ha.items()})
        A = sum(r['ha'] for r in X); bu = sum(r['buha'] for r in X); ra = sum(r['ra'] for r in X)
        print(' area', round(A, 2), 'bajo umbral', round(bu, 2), round(bu / A * 100, 1), 'rest act', round(ra, 2), 'IC pond', round(sum(r['ic'] * r['ha'] for r in X) / A, 3))
        print(' delim', {k: round(sum(r[k] for r in X), 1) for k in ['dra', 'den', 'dco', 'dag']})
        print(' eco E6', Counter(r['eco'] for r in X))
        print(' eco mismatch E6 vs V6', [(r['b'], r['eco'][:25], r['v6']) for r in X if r['v6'] and r['v6'].split()[0].lower()[:6] not in r['eco'].lower()][:30])
        print(' PI x clase', Counter((r['pi'], r['cls']) for r in X))
        print(' sit', Counter(r['sit'] for r in X))
        fl = Counter(k for r in X for k, _ in r['f']); print(' flags', fl)
        for r in X:
            for k, t in r['f']:
                if k != 'soc': print('   ', r['b'], k, t, '| regd=', r['regd'], 'sit=', r['sit'])
        print(' tip', Counter(r['tip'] for r in X), 'ev', Counter(r['ev'] for r in X))
        print(' cons F-DT x clase', Counter((r['cons'], r['cls']) for r in X).most_common(8))
