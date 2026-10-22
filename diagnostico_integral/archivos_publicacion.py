"""Genera las listas de archivos para publicar el artefacto con la herramienta Artifact.

  publicar/archivos_parte1.json   páginas de los tres volúmenes + láminas de los Vol. II y III
  publicar/archivos_parte2.json   láminas del Vol. I

Las láminas se copian en el servidor desde los artefactos originales de cada
volumen (máximo 255 archivos por publicación, por eso van en dos partes).
Uso:  python archivos_publicacion.py
"""
import json
import os

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, 'publicar')
A = 'https://claude.ai/artifact/4kF1tRzj2FkuDgbcuTcBye'   # Vol. III Ayabaca
H = 'https://claude.ai/artifact/9828ogNDewfGQLgQDCq7kn'   # Vol. II Huancabamba
M = 'https://claude.ai/artifact/Y5QRXCQRsXFvwFgYL1oZ64'   # Vol. I Morropón

aya_b = '27 3 36 39 56 6 M17B10 M17B5 M17B6 M17B7'.split()
aya = [f'eco/Ecosistemas_Bloque_{b}.png' for b in aya_b] + [f'maps/PI_{b}.png' for b in aya_b + ['Ayabaca_conjunto']]
aya += [f'sp/SP-{i:03d}_Bloque_{b}.png' for i, b in enumerate('3 6 27 36 39 56 M17B5 M17B6 M17B7 M17B10'.split(), 1)]

hua_b = ('10 16 18 19 20 21 24 26 28 30 31 34 35 37 38 4 40 41 43 44 47 50 51 53 60 61 63 64 66 67 69 70 71 72 73 '
         '76 79 8 80 81 9 M12B1 M20B1 M22B1 M2B8 M30B5 M4B3 M4B4').split()
hua = [f'eco/Ecosistemas_Bloque_{b}.png' for b in hua_b] + [f'mapas/PI_Bloque_{b}.png' for b in hua_b]
hua += [f'mapas/PI_Distrito_{d}.png' for d in 'Canchaque Huancabamba Huarmaca Lalaquiz San_Miguel_de_El_Faique'.split()]
hua += ['mapas/PI_Provincial_Huancabamba.png']
hua += [f'sp/SP-{i:03d}_Bloque_{b}.png' for i, b in enumerate(
    ('21 24 28 30 40 60 63 66 79 81 64 M30B5 4 8 10 16 18 19 20 31 41 51 53 69 70 71 72 73 76 80 M4B3 M4B4 M12B1 '
     '26 34 37 38 50 61 9 35 43 44 47 67 M2B8 M20B1 M22B1').split(), 11)]

mor_b = ('11 12 13 14 15 17 2 22 23 42 49 5 52 55 57 58 59 77 78 82 83 84 85 86 87 M10B4 M11B3 M17B1 M17B4 M18B1 '
         'M18B3 M19B2 M19B7 M1B1 M27B1 M28B2 M28B3 M28B4 M2B1 M2B5 M32B3 M36B2 M3B1 M3B3 M3B5 M3B6 M3B7 M3B8 M3B9 '
         'M6B10 M6B2-1 M6B2-2 M6B2-3 M7B1 M7B2 M7B3 M7B6 M8B2 M9B1').split()
mor = [f'eco/Ecosistemas_Bloque_{b}.png' for b in mor_b] + [f'mapas/PI_{b}.jpg' for b in mor_b]
mor += [f'sp/SP-{i:03d}_Bloque_{b}.png' for i, b in enumerate(
    ('M1B1 M6B2-1 M6B2-2 M6B2-3 M6B10 M18B1 57 59 2 22 M9B1 M10B4 M17B1 M17B4 M27B1 M19B2 M32B3 M2B1 M2B5 M3B1 '
     'M3B9 M7B1 M7B2 M7B3 M7B6 M18B3 77 83 84 85 86 87 M3B3 M3B5 M3B6 M3B7 M3B8 M11B3 14 42 58 M36B2 11 12 13 15 '
     '17 23 78 82 M8B2 M19B7 49 52 55 M28B2 M28B3 M28B4').split(), 60)]

os.makedirs(OUT, exist_ok=True)
p1 = {f'{d}/index.html': f'publicar/{d}/index.html' for d in ['vol1-morropon', 'vol2-huancabamba', 'vol3-ayabaca']}
p1.update({'vol3-ayabaca/' + p: {'artifact': A, 'path': p} for p in aya})
p1.update({'vol2-huancabamba/' + p: {'artifact': H, 'path': p} for p in hua})
p2 = {'vol1-morropon/' + p: {'artifact': M, 'path': p} for p in mor}
json.dump(p1, open(os.path.join(OUT, 'archivos_parte1.json'), 'w'), indent=0)
json.dump(p2, open(os.path.join(OUT, 'archivos_parte2.json'), 'w'), indent=0)
print(f'Vol. III {len(aya)} · Vol. II {len(hua)} · Vol. I {len(mor)} láminas · parte 1: {len(p1)} archivos · parte 2: {len(p2)}')
