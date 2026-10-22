"""Construye UPD (res, ds, T) del corte 17/10 a partir del libro VAL_PCT, con las convenciones del UPD 07/10."""
import json, re, sys
sys.path.insert(0, 'bin')
from parse_ds import parse_book

PRIV = {'Nombre del actor', 'Responsable', 'Entrevistado', 'Entrevistado/a'}

def old_modes():
    M = {}
    for p in 'AHM':
        O = json.load(open(f'work/{p}_UPD_old.json'))
        for k, L in O['ds'].items():
            for t in L:
                M[(k, re.sub(r'^[A-Z]\. ', '', t['t']))] = (t['m'], t['h'])
    return M

def fix_tables(ds, OM):
    out = {}
    for k, L in ds.items():
        titles = {t['t'] for t in L}
        res = []
        for t in L:
            # descartar gemelas en %: título X (%) cuya versión X existe
            if t['t'].endswith(' (%)') and t['t'][:-4] in titles:
                continue
            key = (k, re.sub(r'^[A-Z]\. ', '', t['t']))
            rows = [r for r in t['r']]
            # columnas de fórmulas detectadas sólo por filas que no son PROMEDIO
            h = list(t['h'])
            if h and h[-1] == '% de la base':
                h = h[:-1]
            calc = [i for i in t['calc'] if i < len(h)]
            if key in OM:
                m, oh = OM[key]
                if oh == h:
                    pass
            else:
                m = None
            if m is None:
                m = t.get('m')
            if m is None:
                if 'actividad económica' in t['t']: m = 'pct'
                elif t['t'].endswith('(%)'): m = 'none'
                else: m = 'val'
            if m in ('none', 'val'):
                calc = []
            if m == 'pct':
                calc = [len(h) - 1] if h[-1].startswith('Familias en la actividad') else calc
            if t['base'] is not None and m != 'base':
                m = 'base'
            res.append(dict(t=t['t'], d=t['d'], h=h, r=rows, n=t['n'], tot=t['tot'], m=m, base=t['base'], calc=calc))
        out[k] = res
    return out

def fix_T(T):
    out = {}
    for k, v in T.items():
        h = v['h']; keep = [i for i, c in enumerate(h) if c not in PRIV]
        nh = [('Fecha de la ficha de referencia' if h[i] == 'Ficha de referencia' else h[i]) for i in keep]
        rows = []
        for r in v['r']:
            if r and str(r[0]).strip().upper() == 'TOTAL':
                continue
            nr = []
            for i in keep:
                x = r[i] if i < len(r) else None
                if h[i] == 'Ficha de referencia' and isinstance(x, str):
                    x = x.split(' · ')[0].strip()
                nr.append(x)
            rows.append(nr)
        if k == 'demo':
            ix = nh.index
            for r in rows:
                T, H, M, I = r[ix('Población total (hab.)')], r[ix('Hombres')], r[ix('Mujeres')], r[ix('Población INEI (catálogo)')]
                hm = H + M if (H is not None and M is not None) else None
                r[-3:] = [hm, T - hm if (hm is not None and T is not None) else None, T - I if (T is not None and I is not None) else None]
        out[k] = dict(h=nh, r=rows)
    return out

if __name__ == '__main__':
    OM = old_modes()
    for p, f in [('A', 'Ayabaca'), ('H', 'Huancabamba'), ('M', 'Morropon')]:
        import glob
        B = parse_book(glob.glob(f'xl/*{f}.xlsx')[0])
        U = dict(res=B['res'], ds=fix_tables(B['ds'], OM), T=fix_T(B['T']))
        json.dump(U, open(f'work/{p}_new.json', 'w'), ensure_ascii=False, default=str)
        print(p, 'ok')
