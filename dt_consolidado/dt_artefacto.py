# -*- coding: utf-8 -*-
"""
Proyecto IN Piura — CUI 2669244 · ANIN / DIME / SESDI
Generador del artefacto de presentacion de resultados del Diagnostico
Territorial, una pagina por provincia.

Paletas de graficos validadas con el validador de contraste y CVD:
  par brecha   claro #256ABF / #C2621B      oscuro #3987e5 / #d95926
  rampa ISL    claro #E2A05E…#552A05        oscuro #8A470C…#F5C88F
  rampa NDVI   claro #83C09A…#10502B        oscuro #1C6339…#8FC7A4
"""
import json
import math
import os
from collections import defaultdict

import dt_data as D

BASE = os.path.dirname(os.path.abspath(__file__))


def corr(xs, ys):
    n = len(xs)
    if n < 3:
        return None
    mx, my = sum(xs) / n, sum(ys) / n
    sxy = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    sxx = sum((x - mx) ** 2 for x in xs)
    syy = sum((y - my) ** 2 for y in ys)
    if sxx <= 0 or syy <= 0:
        return None
    return sxy / math.sqrt(sxx * syy)


def nf(v, d=2):
    """Formato numerico peruano: miles con espacio fino, decimales con coma."""
    if v is None:
        return "s/d"
    s = f"{v:,.{d}f}"
    return s.replace(",", " ").replace(".", ",")


def hallazgos(prov, bs, r, meta):
    """Hallazgos derivados de los datos. Ninguno es afirmacion no sustentada."""
    out = []
    ha = r["ha"]

    # 1. Brecha espectral: media de bloque frente a distribucion areal
    sobre = [b for b in bs if b["msavi_condicion"] == "Sobre umbral"]
    ha_sobre = sum(b["area_ha"] for b in sobre)
    out.append(("cri", "La media del bloque oculta la brecha real",
                f"<p>Por la <strong>media</strong> de su MSAVI 2024, {len(sobre)} de los "
                f"{r['n']} bloques ({nf(100 * ha_sobre / ha, 1)} % de la superficie) quedan "
                f"sobre el umbral 0,4976 y parecerían no requerir recuperación.</p>"
                f"<p>Al descender a la distribución areal por clase, sin embargo, "
                f"<strong>{nf(r['ha_brecha'])} ha</strong> —el "
                f"<strong>{nf(r['pct_brecha'], 2)} %</strong> del ámbito— se sitúan bajo el "
                f"umbral. Esa superficie, y no el conteo de bloques, es la base pertinente "
                f"para el indicador de brecha de la R.M. N° 00213-2024-MINAM.</p>"))

    # 2. ISL: concentracion de susceptibilidad
    alta = [b for b in bs if b["isl_clase_mm"] in ("Alta", "Muy alta")
            or b["isl_clase_eh"] in ("Alta", "Muy alta")]
    ha_alta = sum(b["area_ha"] for b in alta)
    clases_mm = {b["isl_clase_mm"] for b in bs}
    disc = ("no discrimina entre bloques: es la misma clase en todo el ámbito"
            if len(clases_mm) == 1 else
            f"se reparte en {len(clases_mm)} clases dentro del ámbito")
    out.append(("adv" if len(clases_mm) > 1 else "cri",
                "El sustrato marca el piso del riesgo; la cobertura lo diferencia",
                f"<p>El ISL sitúa <strong>{len(alta)} bloques</strong> "
                f"({nf(ha_alta)} ha, {nf(100 * ha_alta / ha, 1)} %) en susceptibilidad alta o "
                f"muy alta en al menos uno de sus dos índices. El ISL-MM {disc}.</p>"
                f"<p>Cuando el sustrato es homogéneo, el factor que diferencia el riesgo "
                f"entre bloques no es la geología sino la <strong>cobertura vegetal y la "
                f"pendiente local</strong>. Eso sitúa a la infraestructura natural, y no a la "
                f"obra estructural, como la medida de mayor rendimiento esperado.</p>"))

    # 3. Convergencia de los tres factores
    conv = [b for b in bs if (b["isl_clase_mm"] in ("Alta", "Muy alta")
                              or b["isl_clase_eh"] in ("Alta", "Muy alta"))
            and (b["pct_bajo_umbral"] or 0) >= 50 and (b["pend_pct"] or 0) >= 25]
    ha_conv = sum(b["area_ha"] for b in conv)
    if conv:
        cods = ", ".join(sorted(b["codigo"] for b in conv)[:14])
        out.append(("cri", "Dónde convergen sustrato, cobertura y pendiente",
                    f"<p><strong>{len(conv)} bloques</strong> ({nf(ha_conv)} ha, "
                    f"{nf(100 * ha_conv / ha, 1)} % del ámbito) reúnen a la vez sustrato de "
                    f"susceptibilidad alta o muy alta, más de la mitad de su superficie bajo el "
                    f"umbral MSAVI y pendiente promedio ≥ 25 %.</p>"
                    f"<p class='mono' style='font-size:13px'>{cods}"
                    f"{'…' if len(conv) > 14 else ''}</p>"
                    f"<p>Son la prioridad máxima del ámbito para control de erosión y "
                    f"estabilización de laderas.</p>"))
    else:
        out.append(("ok", "Ningún bloque reúne los tres factores críticos",
                    "<p>No hay en el ámbito bloques que combinen simultáneamente sustrato de "
                    "susceptibilidad alta o muy alta, más de la mitad de su superficie bajo el "
                    "umbral MSAVI y pendiente promedio ≥ 25 %. La prioridad debe construirse "
                    "sobre factores tomados de a dos.</p>"))

    # 4. Cobertura que protege
    prot = [b for b in bs if (b["isl_clase_mm"] in ("Alta", "Muy alta")
                              or b["isl_clase_eh"] in ("Alta", "Muy alta"))
            and (b["pct_bajo_umbral"] or 0) < 50 and (b["pend_pct"] or 0) >= 25]
    ha_prot = sum(b["area_ha"] for b in prot)
    if prot:
        out.append(("ok", "Donde la cobertura ya está prestando el servicio",
                    f"<p><strong>{len(prot)} bloques</strong> ({nf(ha_prot)} ha, "
                    f"{nf(100 * ha_prot / ha, 1)} %) combinan sustrato susceptible y pendiente "
                    f"≥ 25 % con cobertura mayoritariamente sobre el umbral.</p>"
                    f"<p>En ellos la vegetación está cumpliendo hoy la función de regulación "
                    f"que el proyecto busca recuperar. Dimensionar allí metas de restauración "
                    f"activa a partir del índice llevaría a <strong>remover suelo y abrir dosel "
                    f"sobre superficies que ya funcionan</strong>. La medida pertinente es la "
                    f"conservación.</p>"))

    # 5. Correlacion MSAVI - altitud
    cda = D.con_dato(bs, "alt_min")
    xs = [(b["alt_min"] + b["alt_max"]) / 2 for b in cda]
    ys = [b["msavi"] for b in cda]
    rr = corr(xs, ys)
    if rr is not None and abs(rr) >= 0.4:
        signo = "positiva" if rr > 0 else "negativa"
        out.append(("cri" if rr > 0.6 else "adv",
                    "El MSAVI está leyendo el gradiente altitudinal",
                    f"<p>La correlación entre la altitud media del bloque y su MSAVI 2024 es "
                    f"<strong>r = {nf(rr, 3)}</strong> ({signo}, R² = {nf(rr * rr, 3)}) sobre "
                    f"los {len(cda)} bloques del ámbito con altitud disponible.</p>"
                    f"<p>El índice está capturando en buena medida el gradiente de humedad por "
                    f"piso altitudinal, no el grado de degradación. Un bloque bajo y seco "
                    f"puntúa menos que uno alto y húmedo aunque su ecosistema esté mejor "
                    f"conservado. Es la razón por la que el MSAVI <strong>no puede usarse solo"
                    f"</strong> como indicador de estado de la Unidad Productora.</p>"))

    # 6. Peligro integrado ausente
    sin_pi = len([b for b in bs if "No disponible" in str(b["peligro_integrado"])])
    if sin_pi:
        out.append(("cri", "Falta el peligro integrado en todo el ámbito",
                    f"<p>El nivel de peligro integrado (MCA-AHP) no consta en "
                    f"<strong>{sin_pi} de los {r['n']} bloques</strong>. Sin él, la "
                    f"jerarquización de la cartera descansa en descriptores parciales —ISL del "
                    f"sustrato, pendiente zonal, brecha espectral— y no en el peligro "
                    f"propiamente dicho.</p>"
                    f"<p>Debe tomarse del modelamiento de mesolocalización (PMM + EPH + PGI "
                    f"ponderados por AHP) antes de fijar la prioridad definitiva de "
                    f"intervención. Es el vacío de mayor consecuencia del diagnóstico.</p>"))

    # 7. Carcavas: cobertura del inventario
    if r["n_carcavas"]:
        out.append(("adv", "El inventario de cárcavas cubre una fracción del ámbito",
                    f"<p>Se han digitalizado y codificado <strong>{r['n_carcavas']} cárcavas</strong> "
                    f"({nf(r['long_carcavas'])} m de longitud acumulada) en "
                    f"<strong>{r['bloques_con_carcavas']} de los {r['n']} bloques</strong>.</p>"
                    f"<p>La ausencia de registros en los {r['n'] - r['bloques_con_carcavas']} "
                    f"bloques restantes <strong>no acredita ausencia de cárcavas</strong>. "
                    f"Mientras el levantamiento instrumental no se complete, no es posible "
                    f"fijar metas físicas de infraestructura gris para el conjunto.</p>"))
    else:
        out.append(("adv", "Sin inventario de cárcavas en el ámbito",
                    f"<p>Ninguno de los {r['n']} bloques cuenta con cárcavas digitalizadas y "
                    f"codificadas en la capa vectorial disponible. La ausencia de registros no "
                    f"acredita ausencia de cárcavas: impide, eso sí, dimensionar la "
                    f"infraestructura gris.</p>"))

    # 8. Discrepancia pendiente gabinete-campo
    dis = [b for b in bs if D.norm(b["clase_pendiente"]) != D.norm(b["pendiente_campo"])]
    if dis:
        out.append(("adv", "Pendiente de gabinete frente a pendiente de campo",
                    f"<p>En <strong>{len(dis)} de los {r['n']} bloques</strong> la clase de "
                    f"pendiente derivada de la estadística zonal no coincide con el rango "
                    f"declarado por la brigada.</p>"
                    f"<p>La divergencia es metodológica, no un error de dato: el valor zonal se "
                    f"calcula sobre todo el polígono y el de campo describe la traza recorrida, "
                    f"habitualmente la de menor pendiente por ser la transitable. Debe "
                    f"declararse la unidad y el alcance de cada registro antes de usarlos para "
                    f"dimensionar movimiento de tierras.</p>"))

    # 9. Vacios de campo
    campos = [("Cobertura vegetal total", "cobertura_pct"),
              ("Suelo desnudo", "suelo_desnudo"),
              ("Regeneración natural", "regeneracion"),
              ("N.° de cárcavas en ficha", "carcavas_ficha")]
    faltantes = [(e, len([b for b in bs if D.es_sin_dato(b[c])])) for e, c in campos]
    peor = [f"{e} ({n}/{r['n']})" for e, n in faltantes if n]
    if peor:
        out.append(("cri", "Los campos que deciden el presupuesto están en blanco",
                    f"<p>Se consignan «Por determinar» en la mayor parte del ámbito: "
                    f"{'; '.join(peor)}.</p>"
                    f"<p>La <strong>regeneración natural</strong> es el indicador que decide "
                    f"entre clausura temporal —de costo bajo y alto rendimiento donde el banco "
                    f"de semillas persiste— y plantación. Esa decisión gobierna una fracción "
                    f"sustantiva del presupuesto de inversión, y hoy no tiene sustento de "
                    f"campo.</p>"))

    # 10. Conflicto de unidad de la pendiente (D-P01)
    pa = D.ponderada(bs, "pend_pct")
    pb = D.ponderada(bs, "pend_si_grados_pct")
    mx = max(b["pend_si_grados_pct"] for b in D.con_dato(bs, "pend_si_grados_pct"))
    out.append(("cri", "La unidad de la pendiente del catálogo no está resuelta",
                f"<p>Las 117 plantillas de resumen V6 rotulan el valor de pendiente del catálogo "
                f"como <strong>porcentaje</strong>. El Informe Técnico Consolidado de "
                f"Diagnóstico Territorial del distrito de Frías concluye que ese "
                f"<strong>mismo valor</strong> está expresado en <strong>grados</strong>, tras "
                f"contrastarlo con la capa «Pendientes vector»: error cuadrático medio de 1,14° "
                f"leyéndolo en grados frente a 12,70° leyéndolo como porcentaje (r = +0,956).</p>"
                f"<p>Se ha verificado que el valor numérico es idéntico en ambas fuentes: el "
                f"conflicto es de <strong>unidad</strong>, no de dato. Para este ámbito, la "
                f"pendiente ponderada es <strong>{nf(pa)} %</strong> bajo la primera lectura y "
                f"<strong>{nf(pb)} %</strong> bajo la segunda: casi el doble.</p>"
                f"<p>La elegibilidad no está en cuestión —bajo ambas lecturas ningún bloque "
                f"supera el criterio del 75 %, con un máximo de {nf(mx)} %—, pero "
                f"<strong>el dimensionamiento de toda partida de movimiento de tierras sí lo "
                f"está</strong>. Requiere pronunciamiento formal de SESDI.</p>"))

    # 11. Bloques sin cobertura del MDE (D-M01)
    if r["sin_mde"]:
        out.append(("cri", "Cinco bloques carecen de cobertura del modelo de elevación",
                    f"<p>Los bloques <span class='mono'>{', '.join(sorted(r['sin_mde']))}</span> "
                    f"({nf(r['ha_sin_mde'])} ha) presentan altitud mínima = máxima = 0 msnm, "
                    f"amplitud 0 m y pendiente 0 % en el reporte de estadística zonal.</p>"
                    f"<p>Un bloque no está a nivel del mar ni es perfectamente plano: el cero es "
                    f"el <strong>marcador de ausencia</strong> de cobertura del MDE, no una "
                    f"medición. Son los mismos cinco que no figuran en la base general del "
                    f"estudio de geología. En esta consolidación sus valores se anulan y quedan "
                    f"excluidos de todo promedio, mínimo y máximo; los bloques se conservan en el "
                    f"universo de superficie.</p>"))

    # 12. Reversibilidad
    revers = sum(v["ha"] for k, v in D.conteo(bs, "reversibilidad")
                 if "reversible" in k.lower() and not k.lower().startswith("no"))
    if revers:
        out.append(("ok", "La degradación es mayoritariamente reversible",
                    f"<p><strong>{nf(revers)} ha</strong> ({nf(100 * revers / ha, 1)} % del "
                    f"ámbito) están calificadas en campo como técnicamente reversibles, total o "
                    f"parcialmente.</p>"
                    f"<p>Sustenta la viabilidad de la restauración pasiva y del enriquecimiento "
                    f"bajo dosel frente a la restauración activa integral, de costo unitario "
                    f"muy superior. Es el hallazgo que más favorece la relación "
                    f"beneficio-costo del proyecto en el ámbito.</p>"))
    return out


def datos_js(prov, bs, r):
    bl = []
    for b in bs:
        nd = b["ndvi_clases"]
        bl.append({
            "cod": b["codigo"], "dist": b["distrito"], "mc": b["microcuenca"],
            "cp": b["centro_poblado"], "area": round(b["area_ha"], 3),
            "e": int(b["este"]), "n": int(b["norte"]),
            "amin": int(b["alt_min"]) if b["alt_min"] is not None else None,
            "amax": int(b["alt_max"]) if b["alt_max"] is not None else None,
            "amed": round((b["alt_min"] + b["alt_max"]) / 2, 1) if b["alt_min"] is not None else None,
            "amp": int(b["amplitud"]) if b["amplitud"] is not None else None,
            "piso": b["piso"],
            "pend": round(b["pend_pct"], 2) if b["pend_pct"] is not None else None,
            "pendg": round(b["pend_grados"], 2) if b["pend_grados"] is not None else None,
            "pendB": b["pend_si_grados_pct"], "sinmde": bool(b.get("mde_sin_dato")),
            "clp": b["clase_pendiente"], "clpc": b["pendiente_campo"],
            "forma": b["forma_terreno"], "pos": b["posicion"], "expo": b["exposicion"],
            "aflo": b["afloramientos"], "esc": b["escarpes"], "rem": b["remociones"],
            "msavi": round(b["msavi"], 4), "mscl": b["msavi_clase"],
            "cond": b["msavi_condicion"], "dn": b["dn_dominante"],
            "bajo": round(b["ha_bajo_umbral"] or 0, 3),
            "pbajo": round(b["pct_bajo_umbral"] or 0, 2),
            "ndvi": b["ndvi_modal"],
            "nd": [round(nd.get(k, (0,))[0] or 0, 3) for k in
                   ("Vegetación alta", "Vegetación mediana", "Vegetación ligera",
                    "Tierra desnuda")],
            "eco": b["ecosistema"], "cons": b["estado_conservacion"],
            "uso": b["uso_actual"], "cobt": b["cobertura_tipo"],
            "cob": b["cobertura_pct"], "desn": b["suelo_desnudo"],
            "regen": b["regeneracion"], "ero": b["erosion"],
            "tax": b["taxones"], "san": b["estado_sanitario"],
            "islmm": b["isl_mm"], "clmm": b["isl_clase_mm"],
            "isleh": b["isl_eh"], "cleh": b["isl_clase_eh"],
            "geo": b["geo_unidad"], "ngeo": b["geo_n_unidades"],
            "causa": b["causa"], "vel": b["velocidad"], "rev": b["reversibilidad"],
            "urg": b["urgencia"], "urge": b["urgencia_erosion"],
            "rec": b["recarga"], "pi": b["peligro_integrado"],
            "acc": b["acceso"], "fecha": b["fecha"], "est": b["estaciones"],
            "ncv": b["n_carcavas"], "lcv": round(b["long_carcavas"], 2),
            "cv": [{"c": c["codigo"], "l": c["longitud"], "p": c["pend_prom"],
                    "nv": c["ndvi"], "cm": c["clase_morf"], "cn": c["clase_ndvi"],
                    "e": c["este"], "n": c["norte"],
                    "amin": c["alt_min"], "amax": c["alt_max"]} for c in b["carcavas"]],
            "disc": [{"c": d["codigo"], "campo": d["campo"], "d": d["discrepancia"],
                      "cal": d["calificacion"], "t": d["tratamiento"]}
                     for d in b["consistencia"]],
        })
    return bl


def construir(prov, bloques):
    bs = D.por_provincia(bloques, prov)
    r = D.resumen(bs)
    meta = D.PROV_META[prov]
    bl = datos_js(prov, bs, r)
    hall = hallazgos(prov, bs, r, meta)

    dist = D.distritos_de(bs)
    distritos = []
    for nombre, a in dist:
        sub = [b for b in bs if b["distrito"] == nombre]
        nd = [0.0, 0.0, 0.0, 0.0]
        for b in sub:
            for i, k in enumerate(("Vegetación alta", "Vegetación mediana",
                                   "Vegetación ligera", "Tierra desnuda")):
                nd[i] += b["ndvi_clases"].get(k, (0,))[0] or 0
        distritos.append({
            "d": nombre, "n": a["n"], "ha": round(a["ha"], 2),
            "pct": round(100 * a["ha"] / r["ha"], 2),
            "amin": int(min(x["alt_min"] for x in D.con_dato(sub, "alt_min")))
                    if D.con_dato(sub, "alt_min") else None,
            "amax": int(max(x["alt_max"] for x in D.con_dato(sub, "alt_max")))
                    if D.con_dato(sub, "alt_max") else None,
            "pend": round(D.ponderada(sub, "pend_pct"), 2)
                    if D.ponderada(sub, "pend_pct") is not None else None,
            "msavi": round(D.ponderada(sub, "msavi"), 4),
            "bajo": round(sum(x["ha_bajo_umbral"] or 0 for x in sub), 2),
            "pbajo": round(100 * sum(x["ha_bajo_umbral"] or 0 for x in sub) / a["ha"], 2),
            "ncv": sum(x["n_carcavas"] for x in sub),
            "islmm": round(D.ponderada(sub, "isl_mm"), 2),
            "isleh": round(D.ponderada(sub, "isl_eh"), 2),
            "nd": [round(v, 2) for v in nd],
            "mc": sorted({x["microcuenca"] for x in sub}),
        })

    ha = r["ha"]
    orden = ["Muy baja", "Baja", "Media", "Alta", "Muy alta"]
    isl_agg = {}
    for campo, clave in (("isl_clase_mm", "mm"), ("isl_clase_eh", "eh")):
        agg = defaultdict(lambda: {"n": 0, "ha": 0.0})
        for b in bs:
            agg[b[campo]]["n"] += 1
            agg[b[campo]]["ha"] += b["area_ha"]
        isl_agg[clave] = [{"cl": k, "n": agg[k]["n"], "ha": round(agg[k]["ha"], 2),
                           "pct": round(100 * agg[k]["ha"] / ha, 2)}
                          for k in orden if k in agg]

    cal = defaultdict(int)
    for b in bs:
        for d in b["consistencia"]:
            cal[d["calificacion"]] += 1

    ctx = {
        "prov": meta["nombre"], "vol": meta["volumen_geologia"],
        "n": r["n"], "ha": round(ha, 2), "nd": r["distritos"], "nmc": r["microcuencas"],
        "amin": int(r["alt_min"]), "amax": int(r["alt_max"]),
        "pend": round(r["pend_media"], 2),
        "pendp": round(r["pend_ponderada"], 2),
        "pendB": round(D.ponderada(bs, "pend_si_grados_pct"), 2),
        "sinmde": r["sin_mde"], "haSinMde": round(r["ha_sin_mde"], 2),
        "nAlt": len(D.con_dato(bs, "alt_min")),
        "vol": meta["volumen_geologia"],
        "msavi": round(r["msavi_medio"], 4), "msavip": round(r["msavi_ponderado"], 4),
        "brecha": round(r["ha_brecha"], 2), "pbrecha": round(r["pct_brecha"], 2),
        "ncv": r["n_carcavas"], "lcv": round(r["long_carcavas"], 2),
        "bcv": r["bloques_con_carcavas"],
        "islmm": round(D.ponderada(bs, "isl_mm"), 2),
        "isleh": round(D.ponderada(bs, "isl_eh"), 2),
        "cal": dict(cal), "ncal": sum(cal.values()),
        "distritos": distritos, "isl": isl_agg,
        "corrMsAlt": round(corr([(b["alt_min"] + b["alt_max"]) / 2
                                 for b in D.con_dato(bs, "alt_min")],
                                [b["msavi"] for b in D.con_dato(bs, "alt_min")]) or 0, 3),
    }
    return bl, ctx, hall, r, meta


def render(prov, bloques):
    bl, ctx, hall, r, meta = construir(prov, bloques)
    finds = "\n".join(
        f'<article class="find {k}"><i></i><div class="body"><h3>{t}</h3>{c}</div></article>'
        for k, t, c in hall)
    js_data = json.dumps({"B": bl, "C": ctx}, ensure_ascii=False, separators=(",", ":"))

    return TEMPLATE.replace("__FINDS__", finds).replace("__DATA__", js_data) \
        .replace("__PROV__", meta["nombre"]).replace("__VOL__", meta["volumen_geologia"]) \
        .replace("__N__", str(ctx["n"])).replace("__HA__", nf(ctx["ha"])) \
        .replace("__ND__", str(ctx["nd"])).replace("__NMC__", str(ctx["nmc"])) \
        .replace("__AMIN__", nf(ctx["amin"], 0)).replace("__AMAX__", nf(ctx["amax"], 0)) \
        .replace("__PEND__", nf(ctx["pendp"])).replace("__PENDB__", nf(ctx["pendB"])) \
        .replace("__MSAVI__", nf(ctx["msavip"], 4)) \
        .replace("__PBRECHA__", nf(ctx["pbrecha"], 1)).replace("__NCV__", str(ctx["ncv"])) \
        .replace("__ISLMM__", nf(ctx["islmm"])).replace("__ISLEH__", nf(ctx["isleh"])) \
        .replace("__DISTPAL__", "distrito" if ctx["nd"] == 1 else "distritos")


TEMPLATE = r"""<title>Diagnóstico Territorial __PROV__</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Archivo:wght@500;600;700&family=Source+Sans+3:ital,wght@0,400;0,600;1,400&family=IBM+Plex+Mono:wght@400;500;600&display=swap">
<style>
:root{
  --ground:#F4F7F9; --surface:#FFFFFF; --surface-2:#EBF2F6; --sunken:#E1EBF2;
  --ink:#0F2431; --ink-2:#3D5766; --ink-3:#6C8494;
  --line:#D3E0E8; --line-2:#B9CCD8;
  --anin:#1B4F72; --anin-2:#2E86AB; --anin-wash:#E8F1F7;
  --moss:#1B4D2E; --moss-2:#357051; --moss-wash:#E6F0EA;
  --brick:#A63A2B; --brick-wash:#F7E9E6;
  --ochre:#8A5209; --ochre-wash:#FAF0DC;
  --s-blue:#256ABF; --s-ochre:#C2621B;
  --isl-1:#E2A05E; --isl-2:#C87C2E; --isl-3:#A65C12; --isl-4:#7F4109; --isl-5:#552A05;
  --nd-4:#83C09A; --nd-3:#4F9B72; --nd-2:#2A754F; --nd-1:#10502B;
  --shadow:0 1px 2px rgba(15,36,49,.06), 0 6px 18px -10px rgba(15,36,49,.18);
}
:root:not([data-theme="light"]){ @media (prefers-color-scheme: dark){
  --ground:#0A1720; --surface:#11232E; --surface-2:#172E3B; --sunken:#0D1D26;
  --ink:#E6EFF4; --ink-2:#A9C0CD; --ink-3:#7A94A3;
  --line:#22404F; --line-2:#2E5468;
  --anin:#8CC2E3; --anin-2:#5EAACE; --anin-wash:#16303F;
  --moss:#7FBE9C; --moss-2:#5FA57F; --moss-wash:#14291F;
  --brick:#E2907F; --brick-wash:#33211D;
  --ochre:#DDAE4F; --ochre-wash:#302711;
  --s-blue:#3987e5; --s-ochre:#d95926;
  --isl-1:#8A470C; --isl-2:#A65C12; --isl-3:#C87C2E; --isl-4:#E2A05E; --isl-5:#F5C88F;
  --nd-4:#1C6339; --nd-3:#357F58; --nd-2:#5AA37A; --nd-1:#8FC7A4;
  --shadow:0 1px 2px rgba(0,0,0,.35), 0 8px 22px -12px rgba(0,0,0,.6);
}}
:root[data-theme="dark"]{
  --ground:#0A1720; --surface:#11232E; --surface-2:#172E3B; --sunken:#0D1D26;
  --ink:#E6EFF4; --ink-2:#A9C0CD; --ink-3:#7A94A3;
  --line:#22404F; --line-2:#2E5468;
  --anin:#8CC2E3; --anin-2:#5EAACE; --anin-wash:#16303F;
  --moss:#7FBE9C; --moss-2:#5FA57F; --moss-wash:#14291F;
  --brick:#E2907F; --brick-wash:#33211D;
  --ochre:#DDAE4F; --ochre-wash:#302711;
  --s-blue:#3987e5; --s-ochre:#d95926;
  --isl-1:#8A470C; --isl-2:#A65C12; --isl-3:#C87C2E; --isl-4:#E2A05E; --isl-5:#F5C88F;
  --nd-4:#1C6339; --nd-3:#357F58; --nd-2:#5AA37A; --nd-1:#8FC7A4;
  --shadow:0 1px 2px rgba(0,0,0,.35), 0 8px 22px -12px rgba(0,0,0,.6);
}
*{box-sizing:border-box}
body{margin:0; background:var(--ground); color:var(--ink);
  font-family:"Source Sans 3","Segoe UI",Arial,sans-serif; font-size:15px; line-height:1.58;
  -webkit-font-smoothing:antialiased}
.wrap{max-width:1200px; margin:0 auto; padding-inline:20px}
h1,h2,h3,h4{font-family:Archivo,"Arial Narrow",Arial,sans-serif; text-wrap:balance; margin:0}
.mono{font-family:"IBM Plex Mono",ui-monospace,Consolas,monospace; font-variant-numeric:tabular-nums}
.num{font-variant-numeric:tabular-nums}

.gov{background:var(--anin)}
:root:not([data-theme="light"]) .gov{background:#0F3247}
:root[data-theme="dark"] .gov{background:#0F3247}
.gov .wrap{display:flex; flex-wrap:wrap; gap:3px 18px; align-items:baseline; padding-block:9px}
.gov b{font-family:Archivo,Arial,sans-serif; font-weight:700; font-size:12.5px;
  letter-spacing:.11em; text-transform:uppercase; color:#EAF3F9}
.gov span{font-size:11.5px; letter-spacing:.05em; color:#C4DCEB; text-transform:uppercase}

header.head{border-bottom:1px solid var(--line); background:var(--surface)}
.head .wrap{padding-block:30px 0}
.eyebrow{font-family:"IBM Plex Mono",monospace; font-size:11.5px; letter-spacing:.13em;
  text-transform:uppercase; color:var(--anin-2); font-weight:500; margin:0}
h1{font-size:clamp(29px,4.3vw,45px); font-weight:700; letter-spacing:-.022em;
  line-height:1.05; margin:10px 0 0}
h1 em{font-style:normal; color:var(--anin-2)}
.sub{max-width:68ch; color:var(--ink-2); margin:14px 0 0; font-size:16px}

.strip{display:grid; grid-template-columns:repeat(auto-fit,minmax(136px,1fr));
  border:1px solid var(--line); border-radius:3px; background:var(--surface-2);
  margin:26px 0 30px; overflow:hidden}
.strip div{padding:13px 16px; border-right:1px solid var(--line); border-bottom:1px solid var(--line)}
.strip dt{font-size:10.5px; letter-spacing:.09em; text-transform:uppercase; color:var(--ink-3);
  font-family:"IBM Plex Mono",monospace; margin:0 0 5px}
.strip dd{margin:0; font-family:Archivo,Arial,sans-serif; font-weight:600; font-size:24px;
  letter-spacing:-.02em; font-variant-numeric:tabular-nums; line-height:1.1}
.strip dd small{font-size:12.5px; font-weight:500; color:var(--ink-3); letter-spacing:0}

nav.tabs{position:sticky; top:0; z-index:20; background:var(--surface);
  border-bottom:1px solid var(--line)}
nav.tabs .wrap{display:flex; gap:2px; overflow-x:auto}
nav.tabs button{appearance:none; background:none; border:0; border-bottom:2.5px solid transparent;
  padding:12px 13px; font:600 13.5px/1 Archivo,Arial,sans-serif; color:var(--ink-3);
  cursor:pointer; white-space:nowrap; transition:color .15s}
nav.tabs button:hover{color:var(--ink)}
nav.tabs button[aria-selected="true"]{color:var(--anin); border-bottom-color:var(--anin-2)}
nav.tabs button:focus-visible{outline:2px solid var(--anin-2); outline-offset:-3px}

main{padding-block:34px 20px}
section[hidden]{display:none!important}
.sec-h{display:flex; align-items:baseline; gap:14px; flex-wrap:wrap; margin:0 0 6px}
h2{font-size:23px; font-weight:600; letter-spacing:-.015em}
.sec-h .tag{font-family:"IBM Plex Mono",monospace; font-size:11px; letter-spacing:.1em;
  text-transform:uppercase; color:var(--ink-3)}
.lede{max-width:76ch; color:var(--ink-2); margin:0 0 22px}
.block-sep{margin-top:40px; padding-top:34px; border-top:1px solid var(--line)}

.tw{overflow-x:auto; border:1px solid var(--line); border-radius:3px; background:var(--surface)}
table{border-collapse:collapse; width:100%; font-size:13.5px}
thead th{position:sticky; top:0; background:var(--anin); color:#fff; text-align:center;
  font:600 11.5px/1.3 Archivo,Arial,sans-serif; letter-spacing:.05em; text-transform:uppercase;
  padding:9px 8px; white-space:nowrap; border-right:1px solid rgba(255,255,255,.16)}
:root:not([data-theme="light"]) thead th{background:#15384C; color:#E6EFF4}
:root[data-theme="dark"] thead th{background:#15384C; color:#E6EFF4}
thead th.s{cursor:pointer; user-select:none}
thead th.s:hover{background:var(--anin-2)}
thead th .car{opacity:.4; font-size:9px; margin-left:3px}
thead th[data-dir] .car{opacity:1}
tbody td{padding:7px 9px; border-bottom:1px solid var(--line); text-align:center;
  vertical-align:middle; font-variant-numeric:tabular-nums}
tbody tr:nth-child(even){background:var(--surface-2)}
tbody tr.on{background:var(--anin-wash)}
tbody tr[data-cod]:hover{background:var(--sunken); cursor:pointer}
td.l,th.l{text-align:left}
tfoot td{padding:8px 9px; font-weight:700; background:var(--sunken); border-top:2px solid var(--line-2);
  text-align:center; font-variant-numeric:tabular-nums}
td.cod{font:600 13.5px/1 "IBM Plex Mono",monospace; color:var(--anin); text-align:left;
  white-space:nowrap}
.foot{font-size:12.5px; color:var(--ink-3); margin:10px 0 0; max-width:96ch}

.pill{display:inline-block; padding:1.5px 8px; border-radius:999px; font-size:11.5px;
  font-weight:600; white-space:nowrap; line-height:1.55}
.p-cri{background:var(--brick-wash); color:var(--brick)}
.p-adv{background:var(--ochre-wash); color:var(--ochre)}
.p-ok{background:var(--moss-wash); color:var(--moss)}
.p-neu{background:var(--anin-wash); color:var(--anin)}

figure{margin:0}
.chart{background:var(--surface); border:1px solid var(--line); border-radius:3px;
  padding:16px 16px 10px}
.chart svg{display:block; width:100%; height:auto; overflow:visible}
figcaption{font-size:12.5px; color:var(--ink-3); margin:10px 2px 0; max-width:96ch}
.legend{display:flex; flex-wrap:wrap; gap:6px 18px; margin:12px 2px 0; font-size:12.5px;
  color:var(--ink-2)}
.legend i{display:inline-block; width:11px; height:11px; border-radius:2px; margin-right:6px;
  vertical-align:-1px}
.tip{position:fixed; z-index:60; pointer-events:none; background:var(--surface);
  border:1px solid var(--line-2); border-radius:4px; box-shadow:var(--shadow);
  padding:8px 11px; font-size:12.5px; max-width:270px; color:var(--ink); line-height:1.45}
.tip b{font-family:"IBM Plex Mono",monospace; color:var(--anin); display:block; margin-bottom:3px}
.tip .r{display:flex; justify-content:space-between; gap:14px; font-variant-numeric:tabular-nums}
.tip .r span:last-child{font-weight:600}

.finds{display:grid; gap:14px}
.find{display:grid; grid-template-columns:4px 1fr; background:var(--surface);
  border:1px solid var(--line); border-left:0; border-radius:0 3px 3px 0; overflow:hidden}
.find i{background:var(--anin-2)}
.find.cri i{background:var(--brick)} .find.adv i{background:var(--ochre)}
.find.ok i{background:var(--moss-2)}
.find .body{padding:15px 20px}
.find h3{font-size:15.5px; font-weight:600; margin:0 0 6px}
.find p{margin:0 0 8px; font-size:14.5px; color:var(--ink-2)}
.find p:last-child{margin-bottom:0}
.find strong{color:var(--ink); font-weight:600}

.picker{display:flex; flex-wrap:wrap; gap:5px; margin:0 0 20px; max-height:150px; overflow-y:auto}
.picker button{font:600 12.5px/1 "IBM Plex Mono",monospace; padding:7px 11px; cursor:pointer;
  background:var(--surface); color:var(--ink-2); border:1px solid var(--line-2);
  border-radius:3px; transition:all .13s}
.picker button:hover{border-color:var(--anin-2); color:var(--anin)}
.picker button[aria-pressed="true"]{background:var(--anin); border-color:var(--anin); color:#fff}
:root:not([data-theme="light"]) .picker button[aria-pressed="true"]{background:var(--anin-2); color:#08202C}
:root[data-theme="dark"] .picker button[aria-pressed="true"]{background:var(--anin-2); color:#08202C}
.picker button:focus-visible{outline:2px solid var(--anin-2); outline-offset:2px}

.ficha{background:var(--surface); border:1px solid var(--line); border-radius:3px; overflow:hidden}
.ficha .top{padding:18px 22px; border-bottom:1px solid var(--line); background:var(--surface-2);
  display:flex; flex-wrap:wrap; gap:6px 18px; align-items:baseline}
.ficha .top h3{font-size:26px; font-weight:700; letter-spacing:-.02em;
  font-family:"IBM Plex Mono",monospace}
.ficha .top .mc{font-family:"IBM Plex Mono",monospace; font-size:13px; color:var(--anin-2)}
.ficha .top .cp{color:var(--ink-2); font-size:14px; flex-basis:100%}
.grid{display:grid; grid-template-columns:repeat(auto-fit,minmax(300px,1fr))}
.grid > div{padding:16px 22px; border-right:1px solid var(--line); border-bottom:1px solid var(--line)}
.grid h4{font:600 11px/1 "IBM Plex Mono",monospace; letter-spacing:.11em; text-transform:uppercase;
  color:var(--anin-2); margin:0 0 11px}
dl.kv{margin:0; display:grid; grid-template-columns:auto 1fr; gap:5px 14px; font-size:13.5px}
dl.kv dt{color:var(--ink-3)}
dl.kv dd{margin:0; text-align:right; font-variant-numeric:tabular-nums}
dl.kv dt.w{grid-column:1/-1; font-size:11.5px; margin-top:3px}
dl.kv dd.w{grid-column:1/-1; text-align:left; font-size:13.5px; margin-top:-3px}

.bar{display:flex; height:20px; border-radius:2px; overflow:hidden; margin:5px 0 8px; gap:2px}
.bar span{display:block}

.disc{border:1px solid var(--line); border-radius:3px; overflow:hidden; background:var(--surface)}
.disc details{border-bottom:1px solid var(--line)}
.disc details:last-child{border-bottom:0}
.disc summary{display:grid; grid-template-columns:66px 58px 1fr auto; gap:12px; align-items:center;
  padding:10px 15px; cursor:pointer; list-style:none; font-size:13.8px}
.disc summary::-webkit-details-marker{display:none}
.disc summary:hover, .disc details[open] summary{background:var(--surface-2)}
.disc summary:focus-visible{outline:2px solid var(--anin-2); outline-offset:-2px}
.disc .dc{font:600 12.5px/1 "IBM Plex Mono",monospace; color:var(--anin)}
.disc .db{padding:2px 15px 16px 15px; font-size:13.8px; color:var(--ink-2); display:grid; gap:9px}
.disc .db b{color:var(--ink); font-weight:600; font-family:"IBM Plex Mono",monospace;
  font-size:10.5px; letter-spacing:.09em; text-transform:uppercase; display:block; margin-bottom:2px}

.ctrls{display:flex; flex-wrap:wrap; gap:10px; align-items:center; margin:0 0 14px}
.ctrls input, .ctrls select{font:400 13.5px/1 "Source Sans 3",Arial,sans-serif; padding:7px 10px;
  border:1px solid var(--line-2); border-radius:3px; background:var(--surface); color:var(--ink)}
.ctrls input:focus-visible, .ctrls select:focus-visible{outline:2px solid var(--anin-2);
  outline-offset:1px}
.ctrls label{font-size:12px; color:var(--ink-3); font-family:"IBM Plex Mono",monospace;
  letter-spacing:.06em; text-transform:uppercase}
.cnt{font-size:12.5px; color:var(--ink-3); margin-left:auto; font-variant-numeric:tabular-nums}

footer{border-top:1px solid var(--line); background:var(--surface); margin-top:40px}
footer .wrap{padding-block:22px 30px; font-size:12.5px; color:var(--ink-3); display:grid; gap:5px}
footer b{color:var(--ink-2); font-weight:600}
@media (prefers-reduced-motion:reduce){*{transition:none!important; animation:none!important}}
@media (max-width:640px){
  .grid > div{border-right:0}
  .ficha .top h3{font-size:22px}
  .disc summary{grid-template-columns:60px 1fr; gap:8px}
  .disc summary .dcal{grid-column:1/-1}
}
</style>

<div class="gov"><div class="wrap">
  <b>ANIN</b><span>Dirección de Intervenciones Multisectoriales y de Emergencia — DIME</span>
  <span>Subdirección de Estudios de Inversión — SESDI</span>
</div></div>

<header class="head"><div class="wrap">
  <p class="eyebrow">Proyecto IN Piura · CUI 2669244 · Preinversión (Perfil)</p>
  <h1>Diagnóstico Territorial<br><em>Provincia de __PROV__</em></h1>
  <p class="sub">Los __N__ bloques preliminares de intervención de la provincia, sobre __ND__ __DISTPAL__ y __NMC__ microcuencas de la Cuenca Alta del Río Piura. Integra las fichas F-DT-01 a F-DT-05, la estadística zonal del MDE, los índices MSAVI 2024 y NDVI mediana 2025, el Índice de Susceptibilidad Litológica del estudio de geología y el inventario de cárcavas. Insumo del Capítulo de Identificación.</p>
  <dl class="strip">
    <div><dt>Bloques</dt><dd>__N__</dd></div>
    <div><dt>Superficie</dt><dd>__HA__ <small>ha</small></dd></div>
    <div><dt>Altitud</dt><dd>__AMIN__–__AMAX__ <small>msnm</small></dd></div>
    <div><dt>Pendiente pond.</dt><dd>__PEND__ <small>% · __PENDB__ % s/ lectura B</small></dd></div>
    <div><dt>MSAVI 2024</dt><dd>__MSAVI__ <small>pond.</small></dd></div>
    <div><dt>Brecha espectral</dt><dd>__PBRECHA__ <small>%</small></dd></div>
    <div><dt>ISL-MM · ISL-EH</dt><dd>__ISLMM__·__ISLEH__</dd></div>
    <div><dt>Cárcavas</dt><dd>__NCV__ <small>rasgos</small></dd></div>
  </dl>
</div></header>

<nav class="tabs"><div class="wrap" role="tablist">
  <button role="tab" data-p="hall" aria-selected="true">Hallazgos</button>
  <button role="tab" data-p="dist" aria-selected="false">Distritos</button>
  <button role="tab" data-p="geo" aria-selected="false">Geoespacial</button>
  <button role="tab" data-p="veg" aria-selected="false">Índices de vegetación</button>
  <button role="tab" data-p="isl" aria-selected="false">Geología · ISL</button>
  <button role="tab" data-p="cv" aria-selected="false">Cárcavas</button>
  <button role="tab" data-p="tab" aria-selected="false">Tabla maestra</button>
  <button role="tab" data-p="fic" aria-selected="false">Ficha por bloque</button>
  <button role="tab" data-p="con" aria-selected="false">Consistencia</button>
</div></nav>

<main><div class="wrap">

<section id="p-hall">
  <div class="sec-h"><h2>Lo que gobierna la formulación en este ámbito</h2><span class="tag">Síntesis</span></div>
  <p class="lede">Hallazgos derivados del contraste de las fichas DT de los __N__ bloques con la estadística zonal del MDE, los compuestos Sentinel-2, el __VOL__ y el inventario de cárcavas codificadas. Cada cifra procede de una fuente primaria; ningún valor ausente ha sido imputado.</p>
  <div class="finds">__FINDS__</div>
</section>

<section id="p-dist" hidden>
  <div class="sec-h"><h2>Desagregado por distrito</h2><span class="tag">Unidad político-administrativa</span></div>
  <p class="lede">La distribución de la superficie entre distritos determina la logística de la intervención: cada distrito supone una contraparte municipal, un padrón de titulares de predios y una cadena de suministro de plantones distintas.</p>
  <figure class="chart"><div id="cDist"></div></figure>
  <figcaption id="capDist"></figcaption>
  <div class="tw" style="margin-top:20px"><table id="tDist"><thead></thead><tbody></tbody><tfoot></tfoot></table></div>
  <p class="foot">Las medias se ponderan por la superficie de cada bloque. El ISL ponderado se calcula sobre la escala 1–5 del estudio de geología.</p>

  <div class="block-sep"></div>
  <div class="sec-h"><h2>Cobertura por clase NDVI en cada distrito</h2><span class="tag">NDVI mediana 2025 · Sentinel-2</span></div>
  <p class="lede">Distribución areal de las cuatro clases de densidad de cobertura. Es el reparto de biomasa, no de integridad ecosistémica: un mosaico agrícola vigoroso puntúa como vegetación alta.</p>
  <figure class="chart"><div id="cNdvi"></div></figure>
  <div class="legend" id="legNdvi"></div>
  <figcaption>Barras al 100 % de la superficie clasificada de cada distrito. La superficie absoluta figura en el rótulo de cada barra y en la tabla superior.</figcaption>
</section>

<section id="p-geo" hidden>
  <div class="sec-h"><h2>Perfil altitudinal de los bloques</h2><span class="tag">Estadística zonal del MDE</span></div>
  <p class="lede">Cada barra es el rango de cota que abarca el polígono del bloque. La amplitud interna importa tanto como la cota: un bloque que atraviesa más de un piso altitudinal no admite un elenco florístico ni una prescripción de manejo únicos.</p>
  <figure class="chart"><div id="cAlt"></div></figure>
  <figcaption>Barra = rango altitudinal del polígono (mínimo a máximo de la estadística zonal). El orden es descendente por cota máxima. Pase el cursor sobre una barra para el detalle del bloque.</figcaption>

  <div class="block-sep"></div>
  <div class="sec-h"><h2>Vigor de la cobertura frente a la altitud</h2><span class="tag">MSAVI 2024 · Sentinel-2</span></div>
  <p class="lede" id="ledeSca"></p>
  <figure class="chart"><div id="cSca"></div></figure>
  <figcaption>Área del círculo proporcional a la superficie del bloque. La línea es el ajuste por mínimos cuadrados sobre los __N__ bloques del ámbito.</figcaption>

  <div class="block-sep"></div>
  <div class="sec-h"><h2>Pendiente promedio por clase</h2><span class="tag">Criterio de idoneidad: máx. 75 %</span></div>
  <div class="finds" style="margin-bottom:24px">
    <article class="find cri"><i></i><div class="body">
      <h3>Las dos lecturas del mismo valor</h3>
      <p>La unidad del campo de pendiente del catálogo no está resuelta entre las fuentes del proyecto. Esta página publica la <strong>lectura A</strong> —el valor rotulado como porcentaje en las plantillas V6— y declara en paralelo la <strong>lectura B</strong>, en la que ese mismo número son grados.</p>
      <p>Ponderadas por superficie, para este ámbito: <strong id="pA"></strong> bajo la lectura A y <strong id="pB"></strong> bajo la lectura B. Bajo ninguna de las dos hay bloques por encima del criterio de idoneidad del 75 %.</p>
    </div></article>
  </div>
  <div class="tw"><table id="tPend"><thead></thead><tbody></tbody><tfoot></tfoot></table></div>
  <p class="foot">La clase corresponde a la lectura A. La pendiente promedio zonal se calcula sobre la totalidad del polígono. Los perfiles de transecto levantados en campo cubren la traza recorrida y no son comparables con el valor zonal: la discrepancia entre ambos registros es metodológica, no un error de dato.</p>
</section>

<section id="p-veg" hidden>
  <div class="sec-h"><h2>Brecha espectral por bloque</h2><span class="tag">Umbral del proyecto: MSAVI = 0,4976</span></div>
  <p class="lede">Reparto de la superficie de cada bloque entre las clases sobre y bajo el umbral MSAVI 0,4976. La fracción bajo umbral es la base para el indicador de brecha de la R.M. N° 00213-2024-MINAM, una vez depurada la fracción de agroecosistema que el desagregado de ecosistemas —aún pendiente— determine.</p>
  <figure class="chart"><div id="cBrecha"></div></figure>
  <div class="legend" id="legBrecha"></div>
  <figcaption>Orden descendente por porcentaje bajo umbral. Cada barra está dimensionada al 100 % de la superficie clasificada del bloque; la superficie absoluta aparece al pasar el cursor.</figcaption>

  <div class="block-sep"></div>
  <div class="sec-h"><h2>Advertencias de lectura</h2><span class="tag">Qué mide y qué no mide el índice</span></div>
  <div class="finds">
    <article class="find cri"><i></i><div class="body">
      <h3>El índice mide biomasa, no integridad ecosistémica</h3>
      <p>Un valor alto <strong>no</strong> equivale a ausencia de degradación. En bloques con mosaico agrícola, pastizal cultivado o plantaciones, la respuesta espectral puede ser alta sobre una Unidad Productora sustituida en su composición. Un maizal vigoroso puntúa por encima de un bosque seco en estiaje.</p>
    </div></article>
    <article class="find adv"><i></i><div class="body">
      <h3>MSAVI 2024 y NDVI 2025 no son comparables entre sí</h3>
      <p>La diferencia entre ambos productos <strong>no debe leerse como mejora</strong> entre 2024 y 2025. Son índices distintos —el NDVI satura ante biomasa densa, el MSAVI conserva sensibilidad al suelo de fondo— aplicados sobre compuestos temporales distintos.</p>
    </div></article>
    <article class="find adv"><i></i><div class="body">
      <h3>La caducifolia del bosque seco sesga el estiaje</h3>
      <p>En los bloques de Bosque Estacionalmente Seco, las tomas de estiaje registran defoliación mientras el campo observa dosel verde continuo en las mismas laderas. No es un error de dato: es el ciclo fenológico. Toda meta dimensionada sobre índices de estiaje debe tratarse como <strong>cota superior</strong>, y su cierre exige al menos una captura de temporada húmeda (enero–abril).</p>
    </div></article>
  </div>
</section>

<section id="p-isl" hidden>
  <div class="sec-h"><h2>Índice de Susceptibilidad Litológica</h2><span class="tag" id="tagIsl"></span></div>
  <p class="lede">El ISL jerarquiza los bloques en función de la competencia mecánica, la anisotropía estructural, el perfil de meteorización y el estado de consolidación del sustrato, sobre la Carta Geológica Nacional 1:50 000 del INGEMMET. Dos escalas independientes de 1 a 5: <strong>ISL-MM</strong> para movimientos en masa e <strong>ISL-EH</strong> para erosión hídrica.</p>
  <figure class="chart"><div id="cIsl"></div></figure>
  <div class="legend" id="legIsl"></div>
  <figcaption>Superficie del ámbito por clase de susceptibilidad, en las dos escalas. Clases: &lt; 1,50 Muy baja · 1,50–2,49 Baja · 2,50–3,49 Media · 3,50–4,49 Alta · ≥ 4,50 Muy alta.</figcaption>

  <div class="finds" style="margin-top:26px">
    <article class="find cri"><i></i><div class="body">
      <h3>El ISL no es un nivel de peligro</h3>
      <p>El índice describe una propiedad del <strong>sustrato</strong> y nada más. No integra pendiente, precipitación detonante, cobertura vegetal ni condiciones hidrogeológicas. Su conversión en peligro —la integración de PMM, EPH y PGI ponderados por AHP— corresponde al Entregable 8. La verificación de campo y los ensayos geotécnicos se ejecutarán en la fase de estudio definitivo.</p>
    </div></article>
  </div>

  <div class="block-sep"></div>
  <div class="sec-h"><h2>Lectura cruzada: sustrato, cobertura y pendiente</h2><span class="tag">Aproximación de gabinete</span></div>
  <p class="lede">La susceptibilidad del sustrato solo adquiere significado operativo al cruzarse con el estado de la cobertura y la pendiente. Los cuatro cuadrantes ordenan la prioridad preliminar; no sustituyen el modelamiento de peligro integrado.</p>
  <div class="tw"><table id="tCuad"><thead></thead><tbody></tbody><tfoot></tfoot></table></div>
  <p class="foot">Sustrato susceptible = clase Alta o Muy alta en ISL-MM o en ISL-EH. Cobertura deficitaria = 50 % o más de la superficie del bloque bajo el umbral MSAVI 0,4976. Pendiente exigente = promedio zonal ≥ 25 %.</p>

  <div class="block-sep"></div>
  <div class="sec-h"><h2>Unidad geológica dominante</h2><span class="tag">Mayor superficie de intersección</span></div>
  <div class="tw"><table id="tGeo"><thead></thead><tbody></tbody><tfoot></tfoot></table></div>
  <p class="foot">La unidad dominante es la de mayor superficie de intersección con el polígono del bloque; no agota su composición litológica. El número de unidades por bloque figura en la ficha individual.</p>
</section>

<section id="p-cv" hidden>
  <div class="sec-h"><h2>Cárcavas codificadas</h2><span class="tag">Capa vectorial · caracterización morfométrica</span></div>
  <p class="lede" id="ledeCv"></p>
  <div id="cvBody"></div>
  <div class="finds" style="margin-top:26px">
    <article class="find cri"><i></i><div class="body">
      <h3>La ausencia de registros no acredita ausencia de cárcavas</h3>
      <p>El inventario recoge únicamente los rasgos efectivamente digitalizados y codificados. En los bloques sin registro, el levantamiento instrumental georreferenciado <strong>está pendiente</strong>, y las fichas F-DT-02 consignan el campo «N.° de cárcavas registradas» como «Por determinar».</p>
      <p>Cerrar ese levantamiento es requisito previo a fijar metas físicas de infraestructura gris: diques de contención, estructuras transversales de retención y obras de estabilización.</p>
    </div></article>
  </div>
</section>

<section id="p-tab" hidden>
  <div class="sec-h"><h2>Base geoespacial consolidada</h2><span class="tag">Ordenable · clic en la fila abre la ficha</span></div>
  <p class="lede">Superficie, centroide y microcuenca del catálogo maestro Bloques V5/V6; altitud y pendiente de la estadística zonal del MDE; MSAVI 2024 y NDVI mediana 2025 de los compuestos Sentinel-2; ISL del estudio de geología; cárcavas de la capa vectorial codificada.</p>
  <div class="ctrls">
    <label for="q">Buscar</label>
    <input id="q" type="search" placeholder="bloque, distrito, microcuenca…" style="min-width:210px">
    <label for="fd">Distrito</label>
    <select id="fd"></select>
    <span class="cnt" id="cnt"></span>
  </div>
  <div class="tw"><table id="tM"><thead></thead><tbody></tbody><tfoot></tfoot></table></div>
  <p class="foot">Coordenadas en UTM WGS 84 Zona 17S (EPSG:32717); todas validan en el rango del proyecto (Este 450 000–750 000 m; Norte 9 300 000–9 600 000 m). El peligro integrado (MCA-AHP) no consta en los insumos de este entregable y debe tomarse del modelamiento de mesolocalización.</p>
</section>

<section id="p-fic" hidden>
  <div class="sec-h"><h2>Ficha consolidada por bloque</h2><span class="tag">F-DT-01 a F-DT-05</span></div>
  <p class="lede">Las casillas «Por determinar» y «Por verificar» indican parámetros que la ficha no consignó. <strong>No equivalen a «No»</strong>: son vacíos de información, y así se conservan.</p>
  <div class="picker" id="picker"></div>
  <div class="ficha" id="ficha"></div>
</section>

<section id="p-con" hidden>
  <div class="sec-h"><h2>Control de consistencia</h2><span class="tag" id="tagCon"></span></div>
  <p class="lede">Discrepancias detectadas entre el registro de campo, los productos de gabinete y el catálogo maestro. Conforme al principio de trazabilidad, ninguna fuente primaria ha sido modificada: se declara el valor adoptado y la acción requerida. Una calificación CONFORME acredita que la verificación se ejecutó y no arrojó discrepancia; no equivale a validación de campo del dato.</p>
  <div class="ctrls">
    <label for="fc">Calificación</label>
    <select id="fc"></select>
    <label for="fb">Bloque</label>
    <select id="fb"></select>
    <span class="cnt" id="cntc"></span>
  </div>
  <div class="disc" id="disc"></div>
</section>

</div></main>

<footer><div class="wrap">
  <div><b>Elaboración:</b> Ing. Héctor Salomón Cahuas Miller — Ingeniero Forestal · ANIN / DIME / SESDI</div>
  <div><b>Fuentes:</b> fichas de Diagnóstico Territorial F-DT-01 a F-DT-05 por bloque (Plantilla DT Campo Check Validada V5) · catálogo maestro Bloques V5/V6 · estadística zonal de altitud y pendiente sobre el MDE · MSAVI Piura 2024 y NDVI mediana 2025 (Sentinel-2) · __VOL__ (Carta Geológica Nacional 1:50 000 — INGEMMET) · capa «Cárcavas codificadas» y su caracterización morfométrica · cruce INEI — Bloques V5</div>
  <div><b>Principio de no invención de datos:</b> ningún valor ausente ha sido estimado o imputado. Los campos no sustentados en observación de campo, estadística zonal o catálogo oficial se conservan con su marca de origen.</div>
  <div>Proyecto «Recuperación del servicio de regulación de riesgos naturales y recuperación de ecosistemas degradados en la Cuenca Alta del Río Piura» · CUI 2669244 · Marco del indicador de brecha: R.M. N° 00213-2024-MINAM</div>
</div></footer>

<script>
const DATA = __DATA__;
const B = DATA.B, C = DATA.C;
const NDK = ["Vegetación alta","Vegetación mediana","Vegetación ligera","Tierra desnuda"];
const ISLK = ["Muy baja","Baja","Media","Alta","Muy alta"];

const cs = getComputedStyle(document.documentElement);
const T = k => cs.getPropertyValue(k).trim();
function pal(){ return {
  blue:T('--s-blue'), ochre:T('--s-ochre'),
  isl:[T('--isl-1'),T('--isl-2'),T('--isl-3'),T('--isl-4'),T('--isl-5')],
  nd:[T('--nd-1'),T('--nd-2'),T('--nd-3'),T('--nd-4')],
  ink:T('--ink'), ink2:T('--ink-2'), ink3:T('--ink-3'), line:T('--line'),
  anin:T('--anin'), anin2:T('--anin-2'), surface:T('--surface')
};}
/* Formato numerico: miles con espacio fino (U+202F), decimales con coma. */
function fmt(v,d){ if(v==null||Number.isNaN(v)) return 's/d';
  const s=Math.abs(v).toFixed(d), i=s.split('.')[0], f=s.split('.')[1];
  return (v<0?'\u2212':'')+i.replace(/\B(?=(\d{3})+(?!\d))/g,'\u202F')+(f?','+f:''); }
const n0=v=>fmt(v,0), n1=v=>fmt(v,1), n2=v=>fmt(v,2), n4=v=>fmt(v,4);
const esc = s => String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const SD = v => v==null || ['Por determinar','Por verificar','s/d'].includes(String(v).trim());
const sd = v => SD(v) ? '<span style="color:var(--ink-3);font-style:italic">'+esc(v||'Por determinar')+'</span>' : esc(v);

/* ---------- tooltip compartido ---------- */
let tipEl=null;
function tip(html, ev){
  if(!tipEl){ tipEl=document.createElement('div'); tipEl.className='tip'; document.body.appendChild(tipEl); }
  tipEl.innerHTML=html; tipEl.style.display='block';
  const w=tipEl.offsetWidth, h=tipEl.offsetHeight;
  let x=ev.clientX+14, y=ev.clientY+14;
  if(x+w>innerWidth-8) x=ev.clientX-w-14;
  if(y+h>innerHeight-8) y=ev.clientY-h-14;
  tipEl.style.left=x+'px'; tipEl.style.top=Math.max(8,y)+'px';
}
function untip(){ if(tipEl) tipEl.style.display='none'; }
function hookTip(el, html){
  el.addEventListener('mousemove', e=>tip(html,e));
  el.addEventListener('mouseleave', untip);
}
const SVGNS='http://www.w3.org/2000/svg';
function el(t,a){ const e=document.createElementNS(SVGNS,t); for(const k in a) e.setAttribute(k,a[k]); return e; }
function txt(x,y,s,o){ const t=el('text',Object.assign({x,y},o||{})); t.textContent=s; return t; }

/* ---------- 1. barras por distrito ---------- */
function chDist(){
  const host=document.getElementById('cDist'); if(!host) return;
  const P=pal(), D=C.distritos, uni = D.length>1 ? D : C.distritos;
  const rows = D.length>1 ? D.map(d=>({k:d.d,v:d.ha,n:d.n}))
    : [...new Set(B.map(b=>b.mc))].map(m=>({k:m, v:B.filter(b=>b.mc===m).reduce((s,b)=>s+b.area,0),
        n:B.filter(b=>b.mc===m).length})).sort((a,b)=>b.v-a.v);
  document.getElementById('capDist').textContent = D.length>1
    ? 'Superficie de bloques preliminares por distrito. El número sobre cada barra es la cantidad de bloques.'
    : 'El ámbito corresponde a un solo distrito; la barra desagrega por microcuenca. El número sobre cada barra es la cantidad de bloques.';
  const LW=178, RW=96, BH=25, GAP=7, TOP=8;
  const H = TOP + rows.length*(BH+GAP) + 26, W=980;
  const max = Math.max(...rows.map(r=>r.v))*1.02;
  const sc = v => (W-LW-RW) * v/max;
  const svg=el('svg',{viewBox:`0 0 ${W} ${H}`,role:'img','aria-label':'Superficie por unidad'});
  rows.forEach((r,i)=>{
    const y=TOP+i*(BH+GAP);
    svg.appendChild(txt(LW-10,y+BH/2+4,r.k,{'text-anchor':'end','font-size':13,fill:P.ink2,
      'font-family':'"Source Sans 3",Arial,sans-serif'}));
    const w=Math.max(2,sc(r.v));
    const bar=el('rect',{x:LW,y,width:w,height:BH,rx:3,fill:P.blue});
    svg.appendChild(bar);
    hookTip(bar,`<b>${esc(r.k)}</b><div class="r"><span>Superficie</span><span>${n2(r.v)} ha</span></div>`+
      `<div class="r"><span>Bloques</span><span>${r.n}</span></div>`+
      `<div class="r"><span>% del ámbito</span><span>${n1(100*r.v/C.ha)} %</span></div>`);
    svg.appendChild(txt(LW+w+9,y+BH/2+4,`${n2(r.v)} ha · ${r.n} bl.`,{'font-size':12.5,fill:P.ink3,
      'font-family':'"IBM Plex Mono",monospace'}));
  });
  const yA=TOP+rows.length*(BH+GAP)+6;
  svg.appendChild(el('line',{x1:LW,y1:yA,x2:W-RW,y2:yA,stroke:P.line,'stroke-width':1}));
  [0,.25,.5,.75,1].forEach(f=>{
    const x=LW+ (W-LW-RW)*f;
    svg.appendChild(txt(x,yA+16,n0(max*f),{'text-anchor':'middle','font-size':11,fill:P.ink3,
      'font-family':'"IBM Plex Mono",monospace'}));
  });
  svg.appendChild(txt(W-RW,yA+16,'ha',{'text-anchor':'start','font-size':11,fill:P.ink3,
    'font-family':'"IBM Plex Mono",monospace'}));
  host.replaceChildren(svg);
}

/* ---------- 2. NDVI apilado por distrito ---------- */
function chNdvi(){
  const host=document.getElementById('cNdvi'); if(!host) return;
  const P=pal();
  const rows = C.distritos.length>1 ? C.distritos.map(d=>({k:d.d,nd:d.nd}))
    : [{k:C.distritos[0].d, nd:C.distritos[0].nd}];
  const LW=178, RW=118, BH=26, GAP=8, TOP=6, W=980;
  const H=TOP+rows.length*(BH+GAP)+4;
  const svg=el('svg',{viewBox:`0 0 ${W} ${H}`,role:'img','aria-label':'Clases NDVI por distrito'});
  rows.forEach((r,i)=>{
    const y=TOP+i*(BH+GAP), tot=r.nd.reduce((a,b)=>a+b,0)||1;
    svg.appendChild(txt(LW-10,y+BH/2+4,r.k,{'text-anchor':'end','font-size':13,fill:P.ink2,
      'font-family':'"Source Sans 3",Arial,sans-serif'}));
    let x=LW; const avail=W-LW-RW;
    r.nd.forEach((v,j)=>{
      const w=avail*v/tot; if(w<=0) return;
      const seg=el('rect',{x,y,width:Math.max(0,w-2),height:BH,fill:P.nd[j]});
      svg.appendChild(seg);
      hookTip(seg,`<b>${esc(r.k)}</b><div class="r"><span>${NDK[j]}</span><span>${n2(v)} ha</span></div>`+
        `<div class="r"><span>% del distrito</span><span>${n1(100*v/tot)} %</span></div>`);
      if(w>52) svg.appendChild(txt(x+w/2-1,y+BH/2+4,n1(100*v/tot)+' %',{'text-anchor':'middle',
        'font-size':11.5,fill:j<2?'#fff':P.ink,'font-family':'"IBM Plex Mono",monospace'}));
      x+=w;
    });
    svg.appendChild(txt(W-RW+9,y+BH/2+4,n2(tot)+' ha',{'font-size':12,fill:P.ink3,
      'font-family':'"IBM Plex Mono",monospace'}));
  });
  host.replaceChildren(svg);
  document.getElementById('legNdvi').innerHTML = NDK.map((k,j)=>
    `<span><i style="background:${P.nd[j]}"></i>${k}</span>`).join('');
}

/* ---------- 3. rango altitudinal ---------- */
function chAlt(){
  const host=document.getElementById('cAlt'); if(!host) return;
  const P=pal(), rows=[...B].sort((a,b)=>b.amax-a.amax);
  const LW=92, RW=104, BH=Math.max(9, Math.min(20, Math.round(560/rows.length))), GAP=Math.max(2,Math.round(BH*0.28));
  const TOP=8, W=980, H=TOP+rows.length*(BH+GAP)+30;
  const lo=Math.min(...rows.map(r=>r.amin)), hi=Math.max(...rows.map(r=>r.amax));
  const sc=v=>LW+(W-LW-RW)*(v-lo)/(hi-lo||1);
  const svg=el('svg',{viewBox:`0 0 ${W} ${H}`,role:'img','aria-label':'Rango altitudinal por bloque'});
  const yA=TOP+rows.length*(BH+GAP)+6;
  const step = (hi-lo)>2400?500:((hi-lo)>1200?250:100);
  for(let v=Math.ceil(lo/step)*step; v<=hi; v+=step){
    svg.appendChild(el('line',{x1:sc(v),y1:TOP-4,x2:sc(v),y2:yA,stroke:P.line,'stroke-width':1}));
    svg.appendChild(txt(sc(v),yA+16,n0(v),{'text-anchor':'middle','font-size':11,fill:P.ink3,
      'font-family':'"IBM Plex Mono",monospace'}));
  }
  svg.appendChild(txt(W-RW+8,yA+16,'msnm',{'font-size':11,fill:P.ink3,'font-family':'"IBM Plex Mono",monospace'}));
  rows.forEach((r,i)=>{
    const y=TOP+i*(BH+GAP);
    if(BH>=12) svg.appendChild(txt(LW-8,y+BH/2+4,r.cod,{'text-anchor':'end','font-size':Math.min(12,BH-1),
      fill:P.ink2,'font-family':'"IBM Plex Mono",monospace'}));
    const x1=sc(r.amin), x2=sc(r.amax);
    const bar=el('rect',{x:x1,y,width:Math.max(2,x2-x1),height:BH,rx:2,fill:P.blue,opacity:.9});
    svg.appendChild(bar);
    hookTip(bar,`<b>${esc(r.cod)} · ${esc(r.dist)}</b>`+
      `<div class="r"><span>Rango</span><span>${n0(r.amin)}–${n0(r.amax)} msnm</span></div>`+
      `<div class="r"><span>Amplitud</span><span>${n0(r.amp)} m</span></div>`+
      `<div class="r"><span>Superficie</span><span>${n2(r.area)} ha</span></div>`+
      `<div class="r"><span>Pendiente (lectura A)</span><span>${n2(r.pend)} %</span></div>`+
      `<div class="r"><span>MSAVI 2024</span><span>${n4(r.msavi)}</span></div>`);
    if(BH>=12) svg.appendChild(txt(x2+8,y+BH/2+4,n0(r.amp)+' m',{'font-size':11,fill:P.ink3,
      'font-family':'"IBM Plex Mono",monospace'}));
  });
  host.replaceChildren(svg);
}

/* ---------- 4. dispersión MSAVI vs altitud ---------- */
function chSca(){
  const host=document.getElementById('cSca'); if(!host) return;
  const P=pal();
  const L=62,R=22,Tp=14,Bt=46, W=980,H=430;
  const CB=B.filter(b=>b.amed!=null);
  const xs=CB.map(b=>b.amed), ys=CB.map(b=>b.msavi);
  const x0=Math.min(...xs), x1=Math.max(...xs), y0=Math.min(...ys), y1=Math.max(...ys);
  const px=v=>L+(W-L-R)*(v-x0)/((x1-x0)||1), py=v=>H-Bt-(H-Tp-Bt)*(v-y0)/((y1-y0)||1);
  const svg=el('svg',{viewBox:`0 0 ${W} ${H}`,role:'img','aria-label':'MSAVI frente a altitud media'});
  const stepY=(y1-y0)>0.4?0.1:0.05;
  for(let v=Math.ceil(y0/stepY)*stepY; v<=y1+1e-9; v+=stepY){
    svg.appendChild(el('line',{x1:L,y1:py(v),x2:W-R,y2:py(v),stroke:P.line,'stroke-width':1}));
    svg.appendChild(txt(L-9,py(v)+4,n2(v),{'text-anchor':'end','font-size':11,fill:P.ink3,
      'font-family':'"IBM Plex Mono",monospace'}));
  }
  const stepX=(x1-x0)>2000?500:((x1-x0)>900?250:100);
  for(let v=Math.ceil(x0/stepX)*stepX; v<=x1; v+=stepX){
    svg.appendChild(txt(px(v),H-Bt+18,n0(v),{'text-anchor':'middle','font-size':11,fill:P.ink3,
      'font-family':'"IBM Plex Mono",monospace'}));
  }
  svg.appendChild(el('line',{x1:L,y1:H-Bt,x2:W-R,y2:H-Bt,stroke:P.line,'stroke-width':1}));
  svg.appendChild(txt((L+W-R)/2,H-8,'Altitud media del bloque (msnm)',{'text-anchor':'middle',
    'font-size':12,fill:P.ink3,'font-family':'"Source Sans 3",Arial,sans-serif'}));
  svg.appendChild(txt(14,Tp+8,'MSAVI 2024',{'font-size':12,fill:P.ink3,
    'font-family':'"Source Sans 3",Arial,sans-serif'}));
  // umbral
  if(0.4976>=y0 && 0.4976<=y1){
    svg.appendChild(el('line',{x1:L,y1:py(0.4976),x2:W-R,y2:py(0.4976),stroke:P.ochre,
      'stroke-width':2,'stroke-dasharray':'6 4'}));
    svg.appendChild(txt(W-R-4,py(0.4976)-7,'umbral 0,4976',{'text-anchor':'end','font-size':11.5,
      fill:P.ochre,'font-family':'"IBM Plex Mono",monospace'}));
  }
  // ajuste
  const n=xs.length, mx=xs.reduce((a,b)=>a+b,0)/n, my=ys.reduce((a,b)=>a+b,0)/n;
  let sxy=0,sxx=0; for(let i=0;i<n;i++){ sxy+=(xs[i]-mx)*(ys[i]-my); sxx+=(xs[i]-mx)**2; }
  if(sxx>0){ const m=sxy/sxx, b0=my-m*mx;
    svg.appendChild(el('line',{x1:px(x0),y1:py(m*x0+b0),x2:px(x1),y2:py(m*x1+b0),
      stroke:P.ink3,'stroke-width':2,'stroke-dasharray':'2 3',opacity:.75})); }
  const amax=Math.max(...CB.map(b=>b.area));
  CB.forEach(b=>{
    const rr=Math.max(4,Math.sqrt(b.area/amax)*17);
    const c=el('circle',{cx:px(b.amed),cy:py(b.msavi),r:rr,fill:P.blue,
      'fill-opacity':.55,stroke:P.surface,'stroke-width':2});
    svg.appendChild(c);
    hookTip(c,`<b>${esc(b.cod)} · ${esc(b.dist)}</b>`+
      `<div class="r"><span>Altitud media</span><span>${n0(b.amed)} msnm</span></div>`+
      `<div class="r"><span>MSAVI 2024</span><span>${n4(b.msavi)}</span></div>`+
      `<div class="r"><span>Superficie</span><span>${n2(b.area)} ha</span></div>`+
      `<div class="r"><span>Bajo umbral</span><span>${n2(b.pbajo)} %</span></div>`);
  });
  host.replaceChildren(svg);
  const r=C.corrMsAlt, fuerza = Math.abs(r)>=.7?'fuerte':(Math.abs(r)>=.4?'moderada':'débil');
  document.getElementById('ledeSca').innerHTML =
    `La correlación entre la altitud media del bloque y su MSAVI 2024 es <strong>r = ${fmt(r,3)}</strong> `+
    `(${fuerza}${r>0?', positiva':', negativa'}; R² = ${fmt(r*r,3)}) sobre los ${C.nAlt} bloques del ámbito con altitud disponible. `+
    `En la medida en que el índice sigue el gradiente de humedad por piso altitudinal, no puede leerse directamente como grado de degradación de la Unidad Productora.`;
}

/* ---------- 5. brecha espectral apilada ---------- */
function chBrecha(){
  const host=document.getElementById('cBrecha'); if(!host) return;
  const P=pal(), rows=[...B].sort((a,b)=>b.pbajo-a.pbajo);
  const LW=92, RW=112, BH=Math.max(9,Math.min(20,Math.round(560/rows.length))),
        GAP=Math.max(2,Math.round(BH*0.28)), TOP=6, W=980;
  const H=TOP+rows.length*(BH+GAP)+4;
  const svg=el('svg',{viewBox:`0 0 ${W} ${H}`,role:'img','aria-label':'Brecha espectral por bloque'});
  const avail=W-LW-RW;
  rows.forEach((r,i)=>{
    const y=TOP+i*(BH+GAP), pb=r.pbajo, ps=100-pb;
    if(BH>=12) svg.appendChild(txt(LW-8,y+BH/2+4,r.cod,{'text-anchor':'end','font-size':Math.min(12,BH-1),
      fill:P.ink2,'font-family':'"IBM Plex Mono",monospace'}));
    const wS=avail*ps/100, wB=avail*pb/100;
    const s1=el('rect',{x:LW,y,width:Math.max(0,wS-2),height:BH,fill:P.blue});
    const s2=el('rect',{x:LW+wS,y,width:Math.max(0,wB),height:BH,fill:P.ochre});
    svg.appendChild(s1); svg.appendChild(s2);
    const t=`<b>${esc(r.cod)} · ${esc(r.dist)}</b>`+
      `<div class="r"><span>Sobre umbral</span><span>${n2(r.area-r.bajo)} ha · ${n1(ps)} %</span></div>`+
      `<div class="r"><span>Bajo umbral</span><span>${n2(r.bajo)} ha · ${n1(pb)} %</span></div>`+
      `<div class="r"><span>MSAVI medio</span><span>${n4(r.msavi)}</span></div>`;
    hookTip(s1,t); hookTip(s2,t);
    if(BH>=12) svg.appendChild(txt(W-RW+8,y+BH/2+4,n1(pb)+' % · '+n2(r.bajo)+' ha',
      {'font-size':11,fill:P.ink3,'font-family':'"IBM Plex Mono",monospace'}));
  });
  host.replaceChildren(svg);
  document.getElementById('legBrecha').innerHTML =
    `<span><i style="background:${P.blue}"></i>Sobre el umbral MSAVI 0,4976</span>`+
    `<span><i style="background:${P.ochre}"></i>Bajo el umbral — brecha espectral</span>`;
}

/* ---------- 6. ISL apilado ---------- */
function chIsl(){
  const host=document.getElementById('cIsl'); if(!host) return;
  const P=pal();
  const series=[{k:'ISL-MM · movimientos en masa',d:C.isl.mm},{k:'ISL-EH · erosión hídrica',d:C.isl.eh}];
  const LW=210, RW=104, BH=38, GAP=18, TOP=8, W=980, H=TOP+series.length*(BH+GAP)+6;
  const svg=el('svg',{viewBox:`0 0 ${W} ${H}`,role:'img','aria-label':'Superficie por clase de ISL'});
  const avail=W-LW-RW;
  series.forEach((s,i)=>{
    const y=TOP+i*(BH+GAP), tot=s.d.reduce((a,b)=>a+b.ha,0)||1;
    svg.appendChild(txt(LW-10,y+BH/2+4,s.k,{'text-anchor':'end','font-size':13,fill:P.ink2,
      'font-family':'"Source Sans 3",Arial,sans-serif'}));
    let x=LW;
    s.d.forEach(seg=>{
      const j=ISLK.indexOf(seg.cl), w=avail*seg.ha/tot; if(w<=0) return;
      const rct=el('rect',{x,y,width:Math.max(0,w-2),height:BH,fill:P.isl[j<0?2:j]});
      svg.appendChild(rct);
      hookTip(rct,`<b>${esc(s.k)}</b><div class="r"><span>Clase</span><span>${esc(seg.cl)}</span></div>`+
        `<div class="r"><span>Superficie</span><span>${n2(seg.ha)} ha</span></div>`+
        `<div class="r"><span>Bloques</span><span>${seg.n}</span></div>`+
        `<div class="r"><span>% del ámbito</span><span>${n2(seg.pct)} %</span></div>`);
      if(w>96){
        svg.appendChild(txt(x+w/2-1,y+BH/2-2,seg.cl,{'text-anchor':'middle','font-size':12,
          fill:(j>=2?'#fff':P.ink),'font-family':'"Source Sans 3",Arial,sans-serif','font-weight':600}));
        svg.appendChild(txt(x+w/2-1,y+BH/2+14,n1(seg.pct)+' % · '+seg.n+' bl.',{'text-anchor':'middle',
          'font-size':11,fill:(j>=2?'#fff':P.ink),'font-family':'"IBM Plex Mono",monospace',opacity:.9}));
      }
      x+=w;
    });
    svg.appendChild(txt(W-RW+9,y+BH/2+4,n2(tot)+' ha',{'font-size':12,fill:P.ink3,
      'font-family':'"IBM Plex Mono",monospace'}));
  });
  host.replaceChildren(svg);
  const usados=[...new Set([...C.isl.mm,...C.isl.eh].map(s=>s.cl))];
  document.getElementById('legIsl').innerHTML = ISLK.filter(k=>usados.includes(k)).map(k=>
    `<span><i style="background:${P.isl[ISLK.indexOf(k)]}"></i>${k}</span>`).join('');
  document.getElementById('tagIsl').textContent = C.vol || '';
}

/* ---------- tablas ---------- */
function mkTable(id, cols, rows, foot, opts){
  const t=document.getElementById(id); if(!t) return;
  const o=opts||{};
  t.querySelector('thead').innerHTML='<tr>'+cols.map((c,i)=>
    `<th class="${c.l?'l':''}${o.sortable?' s':''}" data-i="${i}">${c.h}${o.sortable?'<span class="car">▲▼</span>':''}</th>`).join('')+'</tr>';
  const body=t.querySelector('tbody');
  const paint=rs=>{ body.innerHTML=rs.map(r=>
    `<tr${r.__cod?` data-cod="${esc(r.__cod)}"`:''}>`+cols.map((c,i)=>
      `<td class="${c.l?'l':''}${c.cod?' cod':''}">${r.c[i]}</td>`).join('')+'</tr>').join(''); };
  paint(rows);
  if(foot) t.querySelector('tfoot').innerHTML='<tr>'+foot.map((v,i)=>
    `<td class="${cols[i]&&cols[i].l?'l':''}">${v}</td>`).join('')+'</tr>';
  if(o.sortable){
    let cur=-1,dir=1;
    t.querySelectorAll('thead th').forEach(th=>th.addEventListener('click',()=>{
      const i=+th.dataset.i;
      dir = (i===cur) ? -dir : 1; cur=i;
      t.querySelectorAll('thead th').forEach(x=>x.removeAttribute('data-dir'));
      th.setAttribute('data-dir',dir>0?'asc':'desc');
      th.querySelector('.car').textContent = dir>0?'▲':'▼';
      rows.sort((a,b)=>{ const x=a.s[i],y=b.s[i];
        return (typeof x==='number'&&typeof y==='number') ? (x-y)*dir
             : String(x).localeCompare(String(y),'es')*dir; });
      paint(rows);
    }));
  }
  return paint;
}

function tablaDistritos(){
  const cols=[{h:'Distrito',l:true},{h:'Bloques'},{h:'Superficie (ha)'},{h:'% ámbito'},
    {h:'Altitud (msnm)'},{h:'Pendiente pond. (%)'},{h:'MSAVI pond.'},{h:'Bajo umbral (ha)'},
    {h:'% bajo umbral'},{h:'ISL-MM'},{h:'ISL-EH'},{h:'Cárcavas'},{h:'Microcuencas'}];
  const rows=C.distritos.map(d=>({c:[esc(d.d),d.n,n2(d.ha),n2(d.pct),(d.amin==null?'s/d':n0(d.amin)+'–'+n0(d.amax)),
    n2(d.pend),n4(d.msavi),n2(d.bajo),n2(d.pbajo),n2(d.islmm),n2(d.isleh),d.ncv,d.mc.length],
    s:[d.d,d.n,d.ha,d.pct,d.amax,d.pend,d.msavi,d.bajo,d.pbajo,d.islmm,d.isleh,d.ncv,d.mc.length]}));
  mkTable('tDist',cols,rows,['TOTAL',C.n,n2(C.ha),'100,00',n0(C.amin)+'–'+n0(C.amax),
    n2(C.pendp),n4(C.msavip),n2(C.brecha),n2(C.pbrecha),n2(C.islmm),n2(C.isleh),C.ncv,C.nmc],
    {sortable:true});
}

function tablaPend(){
  const pa=document.getElementById('pA'), pb=document.getElementById('pB');
  if(pa) pa.textContent=n2(C.pendp)+' %';
  if(pb) pb.textContent=n2(C.pendB)+' %';
  const agg={};
  B.forEach(b=>{ const k=b.clp; (agg[k]=agg[k]||{n:0,ha:0}); agg[k].n++; agg[k].ha+=b.area; });
  const rows=Object.entries(agg).sort((a,b)=>b[1].ha-a[1].ha).map(([k,v])=>({
    c:[esc(k),v.n,n2(v.ha),n2(100*v.ha/C.ha)], s:[k,v.n,v.ha,100*v.ha/C.ha]}));
  mkTable('tPend',[{h:'Clase de pendiente (estadística zonal)',l:true},{h:'Bloques'},
    {h:'Superficie (ha)'},{h:'% del ámbito'}],rows,['TOTAL',C.n,n2(C.ha),'100,00'],{sortable:true});
}

function tablaCuadrantes(){
  const sus=b=>['Alta','Muy alta'].includes(b.clmm)||['Alta','Muy alta'].includes(b.cleh);
  const def=b=>b.pbajo>=50, exi=b=>b.pend>=25;
  const defs=[
    ['Sustrato susceptible · cobertura deficitaria · pendiente exigente',
     b=>sus(b)&&def(b)&&exi(b),'Prioridad máxima: control de erosión y estabilización','p-cri'],
    ['Sustrato susceptible · cobertura deficitaria · pendiente moderada',
     b=>sus(b)&&def(b)&&!exi(b),'Revegetación y conservación de suelos','p-adv'],
    ['Sustrato susceptible · cobertura sobre umbral · pendiente exigente',
     b=>sus(b)&&!def(b)&&exi(b),'Conservación del dosel funcional; no restauración activa','p-ok'],
    ['Sustrato susceptible · cobertura sobre umbral · pendiente moderada',
     b=>sus(b)&&!def(b)&&!exi(b),'Conservación y enriquecimiento selectivo','p-ok'],
    ['Sustrato de susceptibilidad media o baja · cobertura deficitaria',
     b=>!sus(b)&&def(b),'La brecha es de cobertura, no de sustrato','p-adv'],
    ['Sustrato de susceptibilidad media o baja · cobertura sobre umbral',
     b=>!sus(b)&&!def(b),'Seguimiento; sin prioridad de intervención estructural','p-neu']];
  const rows=defs.map(([k,f,o,p])=>{
    const s=B.filter(f), ha=s.reduce((a,b)=>a+b.area,0);
    return {c:[esc(k),s.length,n2(ha),n2(100*ha/C.ha),
      `<span class="pill ${p}">${o}</span>`,
      `<span class="mono" style="font-size:12px">${s.length?esc(s.map(b=>b.cod).sort().slice(0,10).join(', '))+(s.length>10?'…':''):'—'}</span>`],
      s:[k,s.length,ha,100*ha/C.ha,o,'']};
  }).filter(r=>r.s[1]>0);
  mkTable('tCuad',[{h:'Cuadrante',l:true},{h:'Bloques'},{h:'Superficie (ha)'},{h:'% ámbito'},
    {h:'Orientación preliminar',l:true},{h:'Bloques',l:true}],rows,
    ['TOTAL',C.n,n2(C.ha),'100,00','','']);
}

function tablaGeo(){
  const agg={};
  B.forEach(b=>{ const k=b.geo; (agg[k]=agg[k]||{n:0,ha:0,mm:0,eh:0});
    agg[k].n++; agg[k].ha+=b.area; agg[k].mm+=b.islmm*b.area; agg[k].eh+=b.isleh*b.area; });
  const rows=Object.entries(agg).sort((a,b)=>b[1].ha-a[1].ha).map(([k,v])=>({
    c:[esc(k),v.n,n2(v.ha),n2(100*v.ha/C.ha),n2(v.mm/v.ha),n2(v.eh/v.ha)],
    s:[k,v.n,v.ha,100*v.ha/C.ha,v.mm/v.ha,v.eh/v.ha]}));
  mkTable('tGeo',[{h:'Unidad geológica dominante',l:true},{h:'Bloques'},{h:'Superficie (ha)'},
    {h:'% del ámbito'},{h:'ISL-MM pond.'},{h:'ISL-EH pond.'}],rows,
    ['TOTAL',C.n,n2(C.ha),'100,00',n2(C.islmm),n2(C.isleh)],{sortable:true});
}

/* ---------- cárcavas ---------- */
function seccionCv(){
  const host=document.getElementById('cvBody'), lede=document.getElementById('ledeCv');
  const cvs=[]; B.forEach(b=>b.cv.forEach(c=>cvs.push(Object.assign({bloque:b.cod,dist:b.dist,mc:b.mc},c))));
  if(!cvs.length){
    lede.textContent='No se registran cárcavas digitalizadas y codificadas en los bloques de esta provincia dentro de la capa vectorial disponible.';
    host.innerHTML=''; return;
  }
  cvs.sort((a,b)=>b.l-a.l);
  const tot=cvs.reduce((s,c)=>s+c.l,0);
  const desnudo=cvs.filter(c=>c.cn==='Tierra desnuda').length;
  lede.innerHTML=`Se han digitalizado y codificado <strong>${cvs.length} cárcavas</strong> en `+
    `<strong>${C.bcv} de los ${C.n} bloques</strong> del ámbito, con `+
    `<strong>${n2(tot)} m</strong> de longitud acumulada. La caracterización morfométrica `+
    `—longitud, cotas, pendiente promedio, índice de irregularidad y NDVI sobre el eje— permite `+
    `discriminar los rasgos activos de los estabilizados por vegetación: `+
    `<strong>${desnudo}</strong> presentan clase NDVI «Tierra desnuda» sobre su eje.`;
  host.innerHTML='<div class="tw"><table id="tCv"><thead></thead><tbody></tbody><tfoot></tfoot></table></div>'+
    '<p class="foot">Coordenadas en UTM WGS 84 Zona 17S. El NDVI consignado corresponde al valor sobre el eje de la cárcava. La clase morfológica procede de la caracterización de la capa vectorial.</p>';
  const cols=[{h:'Código',l:true,cod:true},{h:'Bloque'},{h:'Distrito',l:true},{h:'Longitud (m)'},
    {h:'Este (m)'},{h:'Norte (m)'},{h:'Altitud (msnm)'},{h:'Pendiente prom. (%)'},{h:'NDVI'},
    {h:'Clase morfológica',l:true},{h:'Clase NDVI',l:true}];
  const rows=cvs.map(c=>({c:[esc(c.c),esc(c.bloque),esc(c.dist),n2(c.l),n0(c.e),n0(c.n),
      n0(c.amin)+'–'+n0(c.amax),n2(c.p),c.nv.toLocaleString('es-PE',{minimumFractionDigits:2}),
      esc(c.cm),`<span class="pill ${c.cn==='Tierra desnuda'?'p-cri':(c.cn==='Vegetación ligera'?'p-adv':'p-ok')}">${esc(c.cn)}</span>`],
    s:[c.c,c.bloque,c.dist,c.l,c.e,c.n,c.amax,c.p,c.nv,c.cm,c.cn]}));
  mkTable('tCv',cols,rows,['TOTAL',C.bcv+' bloques','',n2(tot),'','','','','','',''],{sortable:true});
}

/* ---------- tabla maestra ---------- */
let repaintM=null;
function tablaMaestra(){
  const cols=[{h:'Bloque',l:true,cod:true},{h:'Distrito',l:true},{h:'Microcuenca',l:true},
    {h:'Sup. (ha)'},{h:'Este (m)'},{h:'Norte (m)'},{h:'Altitud (msnm)'},{h:'Pend. (%)'},
    {h:'Pend. B (%)'},{h:'MSAVI 2024'},{h:'% bajo umbral'},{h:'ISL-MM'},{h:'ISL-EH'},{h:'Cárcavas'},{h:'Urgencia',l:true}];
  const mk=b=>({__cod:b.cod, c:[esc(b.cod),esc(b.dist),esc(b.mc),n2(b.area),n0(b.e),n0(b.n),
      (b.amin==null?'s/d':n0(b.amin)+'–'+n0(b.amax)),n2(b.pend),n4(b.msavi),
      `<span class="pill ${b.pbajo>=50?'p-cri':(b.pbajo>=25?'p-adv':'p-ok')}">${n2(b.pbajo)} %</span>`,
      `<span class="pill ${['Alta','Muy alta'].includes(b.clmm)?'p-adv':'p-neu'}">${n2(b.islmm)}</span>`,
      `<span class="pill ${['Alta','Muy alta'].includes(b.cleh)?'p-adv':'p-neu'}">${n2(b.isleh)}</span>`,
      b.ncv||'—',esc(b.urg)],
    s:[b.cod,b.dist,b.mc,b.area,b.e,b.n,b.amax,b.pend,b.pendB,b.msavi,b.pbajo,b.islmm,b.isleh,b.ncv,b.urg]});
  const rows=B.map(mk);
  repaintM=mkTable('tM',cols,rows,['TOTAL','','',n2(C.ha),'','',n0(C.amin)+'–'+n0(C.amax),
    n2(C.pend),n4(C.msavi),n2(C.pbrecha)+' %',n2(C.islmm),n2(C.isleh),C.ncv,''],{sortable:true});
  const fd=document.getElementById('fd');
  fd.innerHTML='<option value="">Todos</option>'+C.distritos.map(d=>`<option>${esc(d.d)}</option>`).join('');
  const q=document.getElementById('q'), cnt=document.getElementById('cnt');
  const filtra=()=>{
    const s=q.value.trim().toLowerCase(), d=fd.value;
    const f=rows.filter(r=>(!d||r.s[1]===d)&&(!s||[r.s[0],r.s[1],r.s[2],r.s[14]].join(' ').toLowerCase().includes(s)));
    repaintM(f);
    document.querySelectorAll('#tM tbody tr').forEach(tr=>tr.addEventListener('click',()=>abrir(tr.dataset.cod)));
    const ha=f.reduce((a,r)=>a+r.s[3],0);
    cnt.textContent=`${f.length} de ${C.n} bloques · ${n2(ha)} ha`;
  };
  q.addEventListener('input',filtra); fd.addEventListener('change',filtra); filtra();
}

/* ---------- ficha ---------- */
function abrir(cod){
  const btn=[...document.querySelectorAll('#picker button')].find(b=>b.dataset.cod===cod);
  if(btn){ document.querySelector('nav.tabs button[data-p="fic"]').click(); btn.click();
    document.getElementById('ficha').scrollIntoView({behavior:'smooth',block:'start'}); }
}
function ficha(cod){
  const b=B.find(x=>x.cod===cod); if(!b) return;
  const P=pal(), tot=b.nd.reduce((a,c)=>a+c,0)||1;
  const barra = b.nd.map((v,j)=>v>0?`<span style="width:${100*v/tot}%;background:${P.nd[j]}" title="${NDK[j]}: ${n2(v)} ha"></span>`:'').join('');
  const kv=(o)=>`<dl class="kv">${o.map(([k,v,w])=>w
      ? `<dt class="w">${k}</dt><dd class="w">${v}</dd>`
      : `<dt>${k}</dt><dd>${v}</dd>`).join('')}</dl>`;
  const pillCl=c=>['Alta','Muy alta'].includes(c)?'p-adv':'p-neu';
  document.getElementById('ficha').innerHTML=`
  <div class="top">
    <h3>${esc(b.cod)}</h3><span class="mc">${esc(b.mc)}</span>
    <span class="pill p-neu">${esc(b.dist)}</span>
    <span class="pill ${b.pbajo>=50?'p-cri':(b.pbajo>=25?'p-adv':'p-ok')}">${n2(b.pbajo)} % bajo umbral</span>
    <span class="cp">${sd(b.cp)}</span>
  </div>
  <div class="grid">
    <div><h4>Identificación y localización</h4>${kv([
      ['Superficie de catálogo',n2(b.area)+' ha'],['Centroide Este',n0(b.e)+' m'],
      ['Centroide Norte',n0(b.n)+' m'],['Microcuenca',esc(b.mc)],
      ['Distrito',esc(b.dist)],['Fecha de evaluación',esc(b.fecha)],
      ['Centro poblado asociado',sd(b.cp),true]])}</div>

    <div><h4>Parámetros físicos · MDE</h4>${kv([
      ['Altitud',(b.amin==null?'s/d':n0(b.amin)+'–'+n0(b.amax))+' msnm'],['Amplitud interna',n0(b.amp)+' m'],
      ['Pendiente — lectura A',b.pend==null?'s/d':n2(b.pend)+' % · '+n2(b.pendg)+'°'],
      ['Pendiente — lectura B',b.pendB==null?'s/d':n2(b.pendB)+' %'],
      ['Forma del terreno',esc(b.forma)],['Posición fisiográfica',esc(b.pos)],
      ['Exposición',esc(b.expo)],['Afloramientos rocosos',esc(b.aflo)],
      ['Escarpes activos',esc(b.esc)],['Remociones activas',esc(b.rem)],
      ['Piso altitudinal',esc(b.piso),true],
      ['Clase de pendiente (gabinete / campo)',esc(b.clp)+' / '+esc(b.clpc),true]])}</div>

    <div><h4>Índices de vegetación</h4>${kv([
      ['MSAVI 2024 (media)',n4(b.msavi)],['Condición vs. umbral',esc(b.cond)],
      ['Superficie bajo umbral',n2(b.bajo)+' ha'],['% bajo umbral',n2(b.pbajo)+' %'],
      ['NDVI — clase modal',esc(b.ndvi)],
      ['Clase MSAVI de la media',esc(b.mscl),true],
      ['Clase DN dominante',esc(b.dn),true]])}
      <div class="bar">${barra}</div>
      <div style="font-size:12.5px;color:var(--ink-2)">${NDK.map((k,j)=>
        b.nd[j]>0?`<div><i style="display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:7px;vertical-align:-1px;background:${P.nd[j]}"></i>${k} — ${n2(b.nd[j])} ha (${n1(100*b.nd[j]/tot)} %)</div>`:'').join('')}</div>
    </div>

    <div><h4>Geología · ISL</h4>${kv([
      ['ISL-MM',`${n2(b.islmm)} <span class="pill ${pillCl(b.clmm)}">${esc(b.clmm)}</span>`],
      ['ISL-EH',`${n2(b.isleh)} <span class="pill ${pillCl(b.cleh)}">${esc(b.cleh)}</span>`],
      ['N.° de unidades geológicas',esc(b.ngeo)],
      ['Unidad dominante',esc(b.geo),true]])}</div>

    <div><h4>Ecosistema y conservación</h4>${kv([
      ['Estado de conservación',sd(b.cons)],['Uso actual dominante',sd(b.uso)],
      ['Tipo de cobertura',sd(b.cobt)],['Cobertura vegetal (campo)',sd(b.cob)],
      ['Suelo desnudo (campo)',sd(b.desn)],['Regeneración natural',sd(b.regen)],
      ['N.° de taxones',esc(b.tax)],['Estado sanitario',sd(b.san)],
      ['Tipo de ecosistema (UP)',sd(b.eco),true]])}</div>

    <div><h4>Erosión y cárcavas</h4>${kv([
      ['Nivel general de erosión',sd(b.ero)],['Urgencia de control',sd(b.urge)],
      ['Cárcavas codificadas',b.ncv||'0'],['Longitud acumulada',b.lcv?n2(b.lcv)+' m':'—'],
      ['Densidad',b.ncv?n2(b.lcv/b.area)+' m/ha':'—']])}
      ${b.cv.length?`<div style="margin-top:9px;font-size:12.5px;color:var(--ink-2)">${
        b.cv.sort((x,y)=>y.l-x.l).map(c=>`<div class="mono" style="font-size:12px">${esc(c.c)} — ${n2(c.l)} m · ${esc(c.cn)}</div>`).join('')}</div>`:''}
    </div>

    <div><h4>Riesgo y degradación · GdR-CCC</h4>${kv([
      ['Velocidad de degradación',sd(b.vel)],['Reversibilidad técnica',sd(b.rev)],
      ['Urgencia de intervención',sd(b.urg)],['Zona de recarga hídrica',sd(b.rec)],
      ['Causa subyacente principal',sd(b.causa),true],
      ['Peligro integrado (MCA-AHP)',sd(b.pi),true]])}</div>

    <div><h4>Verificación de campo</h4>${kv([
      ['Fecha de evaluación',esc(b.fecha)],['Estaciones fotográficas',esc(b.est)],
      ['Modalidad de acceso',sd(b.acc),true]])}
      ${b.disc.length?`<div style="margin-top:10px"><b style="font:600 11px/1 'IBM Plex Mono',monospace;letter-spacing:.11em;text-transform:uppercase;color:var(--anin-2)">Consistencia</b>
        <div style="margin-top:7px;display:grid;gap:4px;font-size:12.5px">${
        b.disc.map(d=>`<div><span class="pill ${d.cal==='SUSTANTIVA'?'p-cri':(d.cal==='CORREGIDO'?'p-adv':(d.cal==='CONFORME'?'p-ok':'p-neu'))}">${esc(d.cal)}</span> ${esc(d.campo)}</div>`).join('')}</div></div>`:''}
    </div>
  </div>`;
}
function pickers(){
  const p=document.getElementById('picker');
  p.innerHTML=[...B].sort((a,b)=>a.cod.localeCompare(b.cod,'es',{numeric:true}))
    .map((b,i)=>`<button data-cod="${esc(b.cod)}" aria-pressed="${i===0?'true':'false'}">${esc(b.cod)}</button>`).join('');
  p.addEventListener('click',e=>{ const b=e.target.closest('button'); if(!b) return;
    p.querySelectorAll('button').forEach(x=>x.setAttribute('aria-pressed','false'));
    b.setAttribute('aria-pressed','true'); ficha(b.dataset.cod); });
  ficha(p.querySelector('button').dataset.cod);
}

/* ---------- consistencia ---------- */
function consistencia(){
  const all=[]; B.forEach(b=>b.disc.forEach(d=>all.push(Object.assign({bloque:b.cod,dist:b.dist},d))));
  const cal={}; all.forEach(d=>cal[d.cal]=(cal[d.cal]||0)+1);
  document.getElementById('tagCon').textContent=
    `${all.length} verificaciones · ${cal['SUSTANTIVA']||0} sustantivas`;
  const fc=document.getElementById('fc'), fb=document.getElementById('fb'), cnt=document.getElementById('cntc');
  fc.innerHTML='<option value="">Todas</option>'+Object.keys(cal).sort().map(k=>`<option>${k}</option>`).join('');
  fb.innerHTML='<option value="">Todos</option>'+[...new Set(all.map(d=>d.bloque))]
    .sort((a,b)=>a.localeCompare(b,'es',{numeric:true})).map(k=>`<option>${esc(k)}</option>`).join('');
  const host=document.getElementById('disc');
  const cls=c=>c==='SUSTANTIVA'?'p-cri':(c==='CORREGIDO'?'p-adv':(c==='CONFORME'?'p-ok':'p-neu'));
  const paint=()=>{
    const f=all.filter(d=>(!fc.value||d.cal===fc.value)&&(!fb.value||d.bloque===fb.value));
    host.innerHTML=f.map(d=>`<details><summary>
        <span class="dc">${esc(d.bloque)}</span><span class="dc">${esc(d.c)}</span>
        <span>${esc(d.campo)}</span>
        <span class="dcal pill ${cls(d.cal)}">${esc(d.cal)}</span></summary>
      <div class="db"><div><b>Discrepancia observada</b>${esc(d.d)}</div>
      <div><b>Tratamiento adoptado</b>${esc(d.t)}</div></div></details>`).join('')
      || '<div style="padding:22px;color:var(--ink-3)">Sin registros para el filtro seleccionado.</div>';
    cnt.textContent=`${f.length} de ${all.length} verificaciones`;
  };
  fc.addEventListener('change',paint); fb.addEventListener('change',paint); paint();
}

/* ---------- navegación ---------- */
document.querySelectorAll('nav.tabs button').forEach(btn=>btn.addEventListener('click',()=>{
  document.querySelectorAll('nav.tabs button').forEach(b=>b.setAttribute('aria-selected','false'));
  btn.setAttribute('aria-selected','true');
  document.querySelectorAll('main section').forEach(s=>s.hidden=true);
  document.getElementById('p-'+btn.dataset.p).hidden=false;
  untip(); window.scrollTo({top:0,behavior:'instant'});
}));

function todo(){
  chDist(); chNdvi(); chAlt(); chSca(); chBrecha(); chIsl();
  tablaDistritos(); tablaPend(); tablaCuadrantes(); tablaGeo();
  seccionCv(); tablaMaestra(); pickers(); consistencia();
}
todo();
matchMedia('(prefers-color-scheme: dark)').addEventListener('change',()=>{
  chDist(); chNdvi(); chAlt(); chSca(); chBrecha(); chIsl(); pickers();
});
</script>
"""


if __name__ == "__main__":
    bl = D.cargar_bloques()
    out = os.path.join(BASE, "entregables")
    os.makedirs(out, exist_ok=True)
    for p in D.PROVINCIAS:
        html = render(p, bl)
        ruta = os.path.join(out, f"artefacto_dt_{D.PROV_META[p]['archivo'].lower()}.html")
        with open(ruta, "w", encoding="utf-8") as fh:
            fh.write(html)
        print(f"OK: {ruta}  ({len(html) / 1024:.0f} KB)")
