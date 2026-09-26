"""Paso 0: prepara las entradas del pipeline.
- respaldo.xlsx: respaldo del aplicativo IN Piura (Respaldo_IN_Piura_AAAAMMDD_HHMMSS.xlsx).
- aya_orig.html, hua_orig.html, mor_orig.html: versión vigente de cada artefacto.
Genera ds.pkl (fichas F-DS con el código de bloque) y los JSON de datos de cada artefacto."""
import json
import pandas as pd
x = pd.read_excel('respaldo.xlsx', sheet_name=None)
ds, bl = x['Diagnostico social'], x['Bloques']
ds['bloque'] = ds.bloque_id.map(dict(zip(bl.id, bl.codigo)))
ds['fecha_registro'] = pd.to_datetime(ds.fecha_registro)
ds.to_pickle('ds.pkl')
for f, n, out in (('hua_orig.html', 1018, 'hua_D.json'), ('mor_orig.html', 958, 'mor_D.json'), ('aya_orig.html', 947, 'aya_X.json')):
    s = open(f).read().split('\n')[n - 1]
    json.dump(json.loads(s[s.index('=') + 1:].strip().rstrip(';')), open(out, 'w'), ensure_ascii=False)
print('ok')
