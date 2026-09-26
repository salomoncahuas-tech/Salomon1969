"""Agregados sociales para los volúmenes con motor «b» (Huancabamba, Morropón)."""
from collections import Counter

FICHAS = ['F-DS-01', 'F-DS-02', 'F-DS-03']


def recalcular(D, regs_por_bloque, regs_por_ficha, ord_worst=('Alta', 'Media', 'Baja'), via='Localidad con ficha en bloque vecino', sin='Sin ficha en el aplicativo', extra=()):
    LOC, B, C = D['LOC'], D['B'], D['C']
    LL = {l['loc']: l for l in LOC}
    # --- bloques ---
    for b in list(B) + list(extra):
        ds = b['ds'] if 'ds' in b else b
        rb = regs_por_bloque.get(b['b'], [])
        for i, f in enumerate(FICHAS, 1):
            ds[f'n0{i}'] = sum(1 for r in rb if r['ficha'] == f)
        locsf = [LL[n] for n in ds.get('locs', []) if n in LL]
        own = [l for l in LOC if b['b'] in (l.get('bl_f') or [])]
        ds['d01'] = sum(l['n01'] for l in own)
        ds['d02'] = sum(l['n02'] for l in own)
        ds['d03'] = sum(l['n03'] for l in own)
        conf = [l for l in locsf + own if l.get('frag')]
        def worst(key, order):
            vals = [l[key] for l in conf if l.get(key)]
            for o in order:
                if o in vals:
                    return o
            return None
        ds['frag'] = worst('frag', ('Alta', 'Media', 'Baja'))
        ds['res'] = worst('res', ('Baja', 'Media', 'Alta'))
        if rb:
            ds['caract'] = 'Con ficha propia'
        elif any(l['nreg'] for l in locsf):
            ds['caract'] = via
        else:
            ds['caract'] = sin
    # --- agregados ---
    C['CREG'] = {f: len(regs_por_ficha.get(f, [])) for f in FICHAS if regs_por_ficha.get(f)}
    C['CDEP'] = {f: sum(l[f'n0{i}'] for l in LOC) for i, f in enumerate(FICHAS, 1) if regs_por_ficha.get(f)}
    L1 = [l for l in LOC if l['n01']]
    C['NF01'] = len(L1)
    acts, fam = Counter(), Counter()
    for l in L1:
        for k, v in (l.get('acts') or {}).items():
            acts[k] += 1
            fam[k] += v or 0
    C['ACTS'] = dict(acts)
    C['FAM'] = {k: float(v) for k, v in fam.items()}
    A = [a for l in LOC for a in l['actors']]
    C['POS'] = dict(Counter(a.get('pos') or 'Sin posición registrada' for a in A))
    C['NLACT'] = len(A)
    E = [e for l in LOC for e in l['entrev']]
    C['GEN'] = dict(Counter(e['gen'] for e in E))
    C['N03'] = len(E)
    L3 = [l for l in LOC if l['n03']]
    C['NL03'] = len(L3)
    C['ACU'] = sum(l.get('acuerdo') or 0 for l in L3)
    C['FRAG'] = dict(Counter(l['frag'] for l in L1 if l.get('frag')))
    C['RES'] = dict(Counter(l['res'] for l in L1 if l.get('res')))
    for key, fld in (('HZ', 'hz'), ('SPP', 'spp'), ('BEN', 'ben')):
        c = Counter()
        for l in L3:
            c.update(sorted(set(l.get(fld) or [])))
        C[key] = dict(c)
    refp = lambda l: l['pob'] or l.get('pctrl') or None
    both = [l for l in LOC if refp(l) and l.get('pdec')]
    C['RATIO'] = sum(l['pdec'] for l in both) / sum(refp(l) for l in both) if both else None
    return C
