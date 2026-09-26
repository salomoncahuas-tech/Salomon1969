"""Terminología de MRR-CCC: «infraestructura gris» -> «infraestructura complementaria»
y separación de la infraestructura natural en verde y marrón."""
import re, unicodedata

VERDE = 'Infraestructura natural verde'
MARRON = 'Infraestructura natural marrón'
COMP = 'Infraestructura complementaria'
GOB = 'Gobernanza y gestión comunitaria'
OTRAS = 'Otras medidas'
ORDEN = [VERDE, MARRON, COMP, GOB, OTRAS]


def plain(s):
    return unicodedata.normalize('NFD', str(s or '')).encode('ascii', 'ignore').decode().lower()

R_CARC = r'carcava|diques? de (piedra|mamposteria|palos|ramas|material)|fajina'
R_MARRON = r'zanja|terraza|carcava|clausura|cierre temporal|pastoreo|pastizal|carga (ganadera|caprina)|exclusion (del|de) ganado|barreras? viva|cercos? vivo|cerco perimetral|acequia|fajina|surcos? en contorno|curvas? (a|de) nivel'
R_COMP = r'talud|muro|gavion|drenaje|cuneta|alcantarill|disipador|dique|baden|banqueta|perfilado|estabilizacion|escorrentia|obras? (de|puntual|complementari|gris)|defensa ribere|enrocad|trocha|via |carretera|puente'
R_VERDE = r'revegeta|reforest|plantacion|plantar|plantones|enriquec|siembra|regeneracion|restauracion|conservacion|nucleacion|vivero|agroforest|silvopast|especies|nativas|sustitucion|reconversion|erradicacion|cobertura|relicto|rodal|semiller|arbol|arbust|forestal|recuperacion'
R_GOB = r'gobernanza|acuerdo|plan(es)? de manejo|capacit|fortalec|alerta|\bsat\b|merese|retribuc|vigilancia|tamizaje|socializ|concerta|coordina|permiso|convenio|comunitari|organizaci|monitoreo|sensibiliz|educacion ambiental|gestion'


def partir(seg):
    out, buf, d = [], '', 0
    i = 0
    while i < len(seg):
        ch = seg[i]
        if ch in '([': d += 1
        elif ch in ')]': d = max(0, d - 1)
        if ch == ';' and d == 0:
            out.append(buf); buf = ''
            while i + 1 < len(seg) and seg[i + 1] == ' ': i += 1
        else:
            buf += ch
        i += 1
    out.append(buf)
    return out


def clase_por_prefijo(pref):
    p = plain(pref)
    if 'gobernanza' in p or 'gestion comunitaria' in p:
        return GOB
    if 'gris' in p or 'complementari' in p:
        return COMP
    if 'marron' in p:
        return MARRON
    if 'natural' in p or 'verde' in p:
        return 'NAT'
    return None


def clasificar(item, pista=None):
    t = plain(item)
    if pista == GOB:
        return GOB
    if pista == COMP:
        return MARRON if re.search(R_CARC, t) or re.search(r'zanja|terraza', t) else COMP
    if re.search(R_CARC, t):
        return MARRON
    if re.search(R_MARRON, t):
        return MARRON
    if pista == 'NAT':
        return COMP if re.search(r'talud|muro|gavion|drenaje|cuneta|alcantarill|disipador', t) and not re.search(R_VERDE, t) else VERDE
    if re.search(R_COMP, t) and not re.search(r'revegeta|reforest|plantacion|enriquec', t):
        return COMP
    if re.search(R_VERDE, t):
        return VERDE
    if re.search(R_COMP, t):
        return COMP
    if re.search(R_GOB, t):
        return GOB
    return OTRAS


PREF = re.compile(r'^\s*((?:Infraestructura|Gobernanza|Obras?)[^:;]{0,70}):\s*', re.I)


def reagrupar(texto):
    """Devuelve (grupos ordenados, texto «A: x; y | B: z»)."""
    if not texto:
        return {}, texto
    t = texto.replace('…', '').strip()
    partes = [p for seg in t.split(' | ') for p in partir(seg)]
    grupos, pista = {}, None
    for p in partes:
        p = p.strip().rstrip('.').strip()
        if not p:
            continue
        m = PREF.match(p)
        if m:
            c = clase_por_prefijo(m.group(1))
            if c:
                pista = c
                p = p[m.end():]
        elif re.match(r'^\s*(Unidad|Zona|Sector)\b', p):
            pista = None
        if not p:
            continue
        p = p[0].upper() + p[1:]
        grupos.setdefault(clasificar(p, pista), []).append(p)
    txt = ' | '.join(f"{k}: {'; '.join(grupos[k])}" for k in ORDEN if k in grupos)
    return grupos, txt


def gris_a_complementaria(s):
    rep = [(r'Infraestructura gris \(complementaria\)', COMP), (r'Infraestructura gris complementaria', COMP),
           (r'infraestructura gris complementaria', 'infraestructura complementaria'),
           (r'Infraestructura gris', COMP), (r'infraestructura gris', 'infraestructura complementaria'),
           (r'Obras grises complementarias', 'Obras complementarias'), (r'obras grises', 'obras complementarias'),
           (r'Obra gris', 'Obra complementaria'), (r'obra gris', 'obra complementaria'),
           (r'Infraestructura natural \(verde\)', 'Infraestructura natural verde'),
           (r'Infraestructura natural/verde', 'Infraestructura natural verde'),
           (r'núcleo de la infraestructura verde', 'núcleo de la infraestructura natural verde')]
    for a, b in rep:
        s = re.sub(a, b, s)
    return s


ETQ = {VERDE: 'Infraestructura natural verde', MARRON: 'Infraestructura natural marrón', COMP: 'Infraestructura complementaria', GOB: 'Gobernanza'}


def etiquetar(texto):
    """Conserva el orden del texto y asigna una clase a cada medida.
    Devuelve (lista [[clase, texto]], texto con los prefijos corregidos)."""
    if not texto:
        return [], texto
    trunc = texto.rstrip().endswith('…')
    partes = [p for seg in texto.replace('…', '').split(' | ') for p in partir(seg)]
    out, pista = [], None
    for i, p in enumerate(partes):
        p = p.strip()
        if not p:
            continue
        m = PREF.match(p)
        cuerpo = p
        if m:
            c = clase_por_prefijo(m.group(1))
            if c:
                pista = c
                cuerpo = p[m.end():]
        elif re.match(r'^\s*(Unidad|Zona|Sector)\b', p):
            pista = None
        if len(cuerpo) < 18 and out:
            cls = out[-1][0]
        else:
            cls = clasificar(cuerpo, pista)
        if cls == OTRAS:
            cls = GOB if re.search(r'ordena|articulaci|control de (quemas|la expansion)|proteccion estricta|decision|matriz', plain(cuerpo)) else (VERDE if re.search(r'plateo|proteccion|rebrote|semil|nodriza', plain(cuerpo)) else (COMP if re.search(r'sediment|cabecera', plain(cuerpo)) else OTRAS))
        if m and clase_por_prefijo(m.group(1)) in ('NAT', COMP, MARRON):
            p = ETQ.get(cls, m.group(1)) + ': ' + cuerpo
        out.append([cls, p + ('…' if trunc and i == len(partes) - 1 else '')])
    return out, '; '.join(x[1] for x in out)
