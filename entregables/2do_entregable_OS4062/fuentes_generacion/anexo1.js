// Anexo N° 1 — Actividad b): actualización, integración y consolidación geoespacial y del diagnóstico territorial
const L = require("./lib");
const { para, b, i, run, h1, h2, h3, literal, guion, bullet, num, fuente, tabla, salto, vacio, centro, W, WH, C, R, fmt, figura } = L;
const { seccionCaratula } = require("./caratula");
const DEF = require("./anexos_def");
const D = require("./datos.json");
const OUT = process.argv[2];
const FIG = "/tmp/claude-0/-home-user-Salomon1969/03b82ed8-c0e9-54ec-93c9-bff5f93b575b/scratchpad/build/fig/";
const H = D.prov["Huancabamba"], M = D.prov["Morropón"];
const N = L.numerador("A1.");
const PIE = "Anexo N° 1 · Informe N° 008-2026-HSCM · Segundo Entregable · OS N° 4062-2026";
const TOT = H.ha + M.ha;
const sum = (f) => f(H) + f(M);
const pc = (a, t) => fmt(100 * a / t, 1) + " %";

const x = [];
x.push(centro([b("ANEXO N° 1", { size: 26 })], { spacing: { after: 60 } }));
x.push(centro([b("ACTUALIZACIÓN, INTEGRACIÓN Y CONSOLIDACIÓN DE LA INFORMACIÓN GEOESPACIAL Y DEL DIAGNÓSTICO TERRITORIAL DE LAS PROVINCIAS DE HUANCABAMBA Y MORROPÓN", { size: 22 })], { spacing: { after: 60 } }));
x.push(centro([i("Diagnóstico Territorial y Social, Versión 5 — corte al 17 de octubre de 2026", { size: 19 })], { spacing: { after: 240 } }));

x.push(h1("1. OBJETIVO Y ALCANCE"));
x.push(para("El presente anexo desarrolla la actividad b) del numeral 5.1 de los Términos de Referencia durante el periodo del segundo entregable (26 de septiembre al 27 de octubre de 2026). Su objetivo es actualizar la información geoespacial y el diagnóstico territorial de los bloques preliminares de intervención de Huancabamba y Morropón, integrar en una estructura única por bloque los estudios especializados del Capítulo de Identificación disponibles al corte y consolidar los resultados que sustentan la delimitación del área geográfica del componente de infraestructura verde."));
x.push(para(`El alcance comprende 107 bloques del catálogo V6 —48 en Huancabamba (${fmt(H.ha)} ha) y 59 en Morropón (${fmt(M.ha)} ha)—, en 14 distritos y ${H.nmc + M.nmc} microcuencas, con un total de ${fmt(TOT)} ha (88.7 % de las 12,270.235 ha del catálogo). La provincia de Ayabaca (10 bloques del distrito de Frías) se consolida en el Volumen III y se cita solo como contraste.`));

x.push(h1("2. LÍNEA BASE Y CORTES DE LA ACTUALIZACIÓN"));
x.push(para("La línea base es la reportada en el primer entregable: los Volúmenes I (Morropón) y II (Huancabamba) del Diagnóstico Territorial, con corte al 21 de septiembre de 2026, 1,175 observaciones sustantivas de control de calidad y catorce matrices M-01 a M-14 por volumen. Sobre esa base, la actualización se ejecutó en tres hitos:"));
x.push(bullet([b("Versión 4 + suelos (03/10/2026): "), "integración del E3 Geología (ISL), el hidroclima del E4 (PISCOp y PISCOt), los datos de campo del E5 Suelos y la capa social del E2; publicación de los tres volúmenes provinciales y del árbol de problemas, medios, fines y alternativas preliminares (E7)."]));
x.push(bullet([b("Actualizaciones del 06 y 07/10/2026: "), "atlas de ecosistemas por bloque V6 (117 láminas, MINAM 2018), matriz de pendientes DEM Copernicus GLO-30, área de influencia aprobada (AI_aprobado_2) y consolidado DS con un centro poblado contado una sola vez."]));
x.push(bullet([b("Versión 5 (17/10/2026): "), "integración del tablero E6 «Estado de los ecosistemas por bloque», de los mapas de situación predial v2 de los 117 bloques (11/10/2026) y del consolidado DS emitido el 17/10/2026 en valores absolutos y porcentajes."]));

x.push(h1("3. FUENTES INTEGRADAS Y MÉTODO"));
x.push(para(`El bloque es la unidad común de integración. Cada fuente se incorpora con su fecha de corte, sin modificar el dato primario; los vacíos se declaran y las divergencias entre fuentes se registran en las pestañas «Consistencia» y «Discrepancias» de cada volumen. El Cuadro N° ${N.sigC()} detalla las fuentes y su uso.`));
x.push(N.cuadro("Fuentes integradas en la Versión 5 y su aporte al diagnóstico por bloque"));
x.push(tabla([{ w: 2300 }, { w: 3600 }, { w: 2338 }, { w: 1400, align: C }], ["FUENTE", "CONTENIDO", "APORTE AL DIAGNÓSTICO", "CORTE"], [
  ["Informes DT y fichas F-DT-01 a F-DT-05", "107 informes técnicos por bloque; plantilla F-DT del catálogo", "Parámetros biofísicos, peligro declarado, accesibilidad, MRR-CCC, pendientes", "20–21/09/2026"],
  ["Capa PI_Bloque (MCA–AHP)", "Peligro Integrado PMM + EPH + PGI en cinco niveles", "Superficie por nivel e índice ponderado por bloque", "Mesolocalización"],
  ["E3 Estudio de Geología", "21 unidades litoestratigráficas; ISL-MM e ISL-EH", "Susceptibilidad del sustrato por bloque", "Culminado"],
  ["E4 Hidroclima (SENAMHI)", "PISCOp v3.0 1981–2025 y PISCOt v1.2 1981–2020 por bloque", "P, ETP, índice de aridez, MFI, años El Niño", "23/09/2026"],
  ["E5 Suelos (GeoSIG)", "74 calicatas, 148 chequeos, 13 pruebas de infiltración; GDB de puntos", "Profundidad, taxonomía y CUM de campo; matriz suelo–intervención", "20/09/2026"],
  ["Atlas de ecosistemas y pendientes V6", "Mapa Nacional de Ecosistemas (R.M. N° 440-2018-MINAM) y DEM GLO-30", "Ecosistema por bloque y área de influencia; pendiente zonal", "05–06/10/2026"],
  ["E6 Estado de los ecosistemas", "S2DR4 1 m, MSAVI por clases, índice de condición, delimitación interna, especies", "Clase de estado, brecha areal, unidades de manejo", "10/2026"],
  ["Situación predial v2", "117 mapas A3 SP-001 a SP-117 (GEORURAL–MIDAGRI, COFOPRI)", "Tenencia por bloque, predios interceptados, CP asociados", "11/10/2026"],
  ["E2 Diagnóstico social", "Fichas F-DS-01 a F-DS-03 vigentes del aplicativo IN Piura", "Tenencia declarada, actores, servicios, fragilidad y resiliencia", "17/10/2026"],
]));
x.push(fuente("Fuente: pie de fuentes de los Volúmenes I y II, Versión 5."));

x.push(h1("4. RESULTADOS CONSOLIDADOS"));
x.push(h2("4.1 Ámbito por provincia"));
x.push(N.cuadro("Ámbito consolidado por provincia"));
x.push(tabla([{ w: 2338 }, { w: 1100, align: R }, { w: 1500, align: R }, { w: 1100, align: R }, { w: 1300, align: R }, { w: 2300 }], ["PROVINCIA", "BLOQUES", "SUPERFICIE (ha)", "DISTRITOS", "MICROCUENCAS", "REGÍMENES TERRITORIALES"], [
  ["Huancabamba (Vol. II)", "48", fmt(H.ha), `${H.nd}`, `${H.nmc}`, H.regTxt],
  ["Morropón (Vol. I)", "59", fmt(M.ha), `${M.nd}`, `${M.nmc}`, M.regTxt],
], { total: ["TOTAL", "107", fmt(TOT), "14", `${H.nmc + M.nmc}`, "—"] }));
x.push(fuente("Fuente: pestaña «Integralidad» de los Volúmenes I y II, Versión 5. Régimen A: bosque seco del piso bajo (< 1,000 msnm); régimen B: mosaico andino de ladera."));

x.push(h2("4.2 Consolidado por distrito"));
x.push(para(`El Cuadro N° ${N.sigC()} consolida por distrito las variables que deciden la delimitación y las metas: brecha, peligro, estado del ecosistema, suelos y tenencia. Se presenta en página horizontal.`));

// --- sección horizontal: cuadro por distrito ---
const xh = [];
xh.push(N.cuadro("Consolidado del diagnóstico territorial por distrito — Huancabamba y Morropón (Versión 5)"));
const cd = [{ w: 1300 }, { w: 1900 }, { w: 800, align: C }, { w: 1150, align: R }, { w: 950, align: C }, { w: 1150, align: R }, { w: 1150, align: R }, { w: 800, align: C }, { w: 900, align: C }, { w: 850, align: C }, { w: 1150, align: R }, { w: 1000, align: R }, { w: 1150, align: R }, { w: 320 + 600, align: C }];
{ const t0 = cd.reduce((s, c) => s + c.w, 0); cd.forEach((c) => (c.w = Math.floor(c.w * WH / t0))); cd[1].w += WH - cd.reduce((s, c) => s + c.w, 0); }
const tot = (k) => D.dist.reduce((s, r) => s + r[k], 0);
xh.push(tabla(cd, ["PROVINCIA", "DISTRITO", "BLOQ.", "SUP. (ha)", "BLOQ. BAJO UMBRAL", "E6 BAJO UMBRAL (ha)", "PI A + MA (ha)", "IC POND.", "BLOQ. MOD./SEV.", "CALIC. DE BLOQUE", "C. C. (ha)", "COFOPRI (ha)", "SIN REGISTRO (ha)", "PREDIOS COFOPRI"],
  D.dist.map((r) => [r.prov, r.d, `${r.n}`, fmt(r.ha), `${r.bajo_n}`, fmt(r.e6_bajo), fmt(r.pi45), fmt(r.ic, 3), `${r.modsev}`, `${r.calB}`, fmt(r.cc), fmt(r.cof), fmt(r.est), `${r.npred}`]),
  { total: ["TOTAL", "14 distritos", `${tot("n")}`, fmt(tot("ha")), `${tot("bajo_n")}`, fmt(tot("e6_bajo")), fmt(tot("pi45")), "—", `${tot("modsev")}`, `${tot("calB")}`, fmt(tot("cc")), fmt(tot("cof")), fmt(tot("est")), `${tot("npred")}`], size: 15 }));
xh.push(fuente("Fuente: matrices de los Volúmenes I y II, Versión 5; Tablero E6; índice de mapas de situación predial v2. «Bloques bajo umbral»: MSAVI medio del bloque < 0.4976. «E6 bajo umbral»: superficie areal bajo 0.4976. IC: índice de condición ponderado por superficie. Calicatas de bloque del Informe de Avance del E5."));
xh.push(para(`Cuatro distritos de Morropón —Buenos Aires, Chulucanas, Salitral y San Juan de Bigote— tienen todos sus bloques bajo el umbral y concentran ${fmt(D.dist.filter((r) => ["Buenos Aires", "Chulucanas", "Salitral", "San Juan de Bigote"].includes(r.d)).reduce((s, r) => s + r.e6_bajo, 0))} ha bajo umbral según el E6. En Huancabamba, Huarmaca reúne ${fmt(D.dist.find((r) => r.d === "Huarmaca").e6_bajo)} ha bajo umbral y ${D.dist.find((r) => r.d === "Huarmaca").npred} predios COFOPRI; los distritos de altura (Lalaquiz, Canchaque, Santo Domingo y Chalaco) presentan los índices de condición más altos y la menor superficie bajo umbral, por lo que su intervención se justifica por la regulación de riesgos más que por la brecha de cobertura.`));

// --- vuelve a vertical ---
const y = [];
y.push(h2("4.3 Peligro Integrado"));
y.push(para(`La tabulación zonal de la capa PI_Bloque ubica ${fmt(H.h45)} ha de Huancabamba (${fmt(H.p45, 1)} %) y ${fmt(M.h45)} ha de Morropón (${fmt(M.p45, 1)} %) en niveles Alto y Muy alto; el índice ponderado es ${fmt(H.pi_idx, 3)} en Huancabamba y ${fmt(M.pi_idx, 3)} en Morropón, ambos en clase Alto. La concordancia entre el peligro declarado en campo y el modelado es baja: en Huancabamba coinciden ${H.conc.Coincide} bloques, ${H.conc["Subestimado en campo"]} están subestimados y ${H.conc["Sobrestimado en campo"]} sobrestimados; en Morropón ${M.conc["Sin declaración en campo"]} informes no declaran nivel. La calificación formal es la del modelo MCA–AHP; la declaración de campo se conserva como trazabilidad.`));
y.push(figura(FIG + "fig_pi.png", 15.5));
y.push(N.figura("Superficie por nivel de Peligro Integrado y provincia"));
y.push(fuente("Fuente: tabulación zonal de la capa PI_Bloque (MCA–AHP). Porcentajes sobre la superficie modelada de cada provincia."));
y.push(N.cuadro("Superficie por nivel de Peligro Integrado (ha)"));
const lv = ["Muy bajo", "Bajo", "Medio", "Alto", "Muy alto"];
y.push(tabla([{ w: 2338 }, ...lv.map(() => ({ w: 1200, align: R })), { w: 1300, align: R }], ["PROVINCIA", ...lv.map((s) => s.toUpperCase()), "TOTAL MODELADO"], [
  ["Huancabamba", ...H.pi_q.map((v) => fmt(v)), fmt(H.pi_q.reduce((a, c) => a + c, 0))],
  ["Morropón", ...M.pi_q.map((v) => fmt(v)), fmt(M.pi_q.reduce((a, c) => a + c, 0))],
], { total: ["TOTAL", ...H.pi_q.map((v, k) => fmt(v + M.pi_q[k])), fmt([...H.pi_q, ...M.pi_q].reduce((a, c) => a + c, 0))] }));
y.push(fuente("Fuente: capa PI_Bloque. La superficie modelada difiere ligeramente de la del catálogo por la cobertura del modelo en bloques de páramo (p. ej., 44 y M22B1)."));

y.push(h2("4.4 Estado del ecosistema (integración del E6)"));
y.push(para("El E6 evalúa cada bloque con tres líneas de evidencia: la clasificación de coberturas de las tomas S2DR4 de 1 m, el MSAVI 2024 por clases frente al umbral 0.4976 y un índice de condición de cinco atributos (ausencia de amenazas, condiciones físicas, diversidad estructural, funciones e intercambios externos) contrastado con ecosistemas de referencia del mismo piso. Su delimitación interna asigna cada sector del bloque a restauración activa, enriquecimiento, conservación o tamizaje agrícola."));
y.push(N.cuadro("Clase de estado del ecosistema y delimitación interna por provincia"));
const cl = ["No degradado", "Leve", "Moderado", "Severo"];
const rowE = (P) => cl.map((k) => `${P.e6_cls[k] || 0} · ${fmt(P.e6_clsha[k] || 0)}`);
y.push(tabla([{ w: 2938 }, { w: 2200, align: R }, { w: 2200, align: R }, { w: 2300, align: R }], ["VARIABLE", "HUANCABAMBA", "MORROPÓN", "TOTAL"], [
  ...cl.map((k, j) => [`Bloques · ha en estado ${k}`, rowE(H)[j], rowE(M)[j], `${(H.e6_cls[k] || 0) + (M.e6_cls[k] || 0)} · ${fmt((H.e6_clsha[k] || 0) + (M.e6_clsha[k] || 0))}`]),
  ["Índice de condición medio ponderado", fmt(H.e6_ic_pond, 3), fmt(M.e6_ic_pond, 3), "—"],
  ["Superficie bajo umbral MSAVI (ha)", fmt(H.e6_bajo_ha), fmt(M.e6_bajo_ha), fmt(H.e6_bajo_ha + M.e6_bajo_ha)],
  ["Restauración activa (ha)", fmt(H.e6_delim.delim_rest_activa_ha), fmt(M.e6_delim.delim_rest_activa_ha), fmt(sum((P) => P.e6_delim.delim_rest_activa_ha))],
  ["Enriquecimiento (ha)", fmt(H.e6_delim.delim_enriquecimiento_ha), fmt(M.e6_delim.delim_enriquecimiento_ha), fmt(sum((P) => P.e6_delim.delim_enriquecimiento_ha))],
  ["Conservación (ha)", fmt(H.e6_delim.delim_conservacion_ha), fmt(M.e6_delim.delim_conservacion_ha), fmt(sum((P) => P.e6_delim.delim_conservacion_ha))],
  ["Tamizaje agrícola (ha)", fmt(H.e6_delim.delim_tamizaje_agricola_ha), fmt(M.e6_delim.delim_tamizaje_agricola_ha), fmt(sum((P) => P.e6_delim.delim_tamizaje_agricola_ha))],
  ["Parcelas virtuales interpretadas", `${H.e6_parcelas}`, `${M.e6_parcelas}`, `${H.e6_parcelas + M.e6_parcelas}`],
  ["Bloques con evidencia de campo A / B / C", `${H.e6_evid.A || 0} / ${H.e6_evid.B || 0} / ${H.e6_evid.C || 0}`, `${M.e6_evid.A || 0} / ${M.e6_evid.B || 0} / ${M.e6_evid.C || 0}`, `${(H.e6_evid.A || 0) + (M.e6_evid.A || 0)} / ${(H.e6_evid.B || 0) + (M.e6_evid.B || 0)} / ${(H.e6_evid.C || 0) + (M.e6_evid.C || 0)}`],
]));
y.push(fuente("Fuente: Tablero E6 integrado en la Versión 5 (corte 10/2026). La delimitación interna se obtiene de las coberturas S2DR4: restauración activa (coberturas 5 y 6), enriquecimiento (3 y 4), conservación (1) y tamizaje agrícola (2)."));
y.push(N.cuadro("Ecosistema dominante de los bloques según el E6"));
const ecos = [...new Set([...Object.keys(H.e6_eco), ...Object.keys(M.e6_eco)])];
y.push(tabla([{ w: 4238 }, { w: 1800, align: R }, { w: 1800, align: R }, { w: 1800, align: R }], ["ECOSISTEMA", "HUANCABAMBA (bloq. · ha)", "MORROPÓN (bloq. · ha)", "TOTAL (ha)"],
  ecos.map((e) => [e, H.e6_eco[e] ? `${H.e6_eco[e][0]} · ${fmt(H.e6_eco[e][1])}` : "—", M.e6_eco[e] ? `${M.e6_eco[e][0]} · ${fmt(M.e6_eco[e][1])}` : "—", fmt((H.e6_eco[e] ? H.e6_eco[e][1] : 0) + (M.e6_eco[e] ? M.e6_eco[e][1] : 0))])));
y.push(fuente("Fuente: Tablero E6, campo «ecosistema» por bloque."));
y.push(para(`La lectura conjunta muestra que la brecha se concentra en el bosque estacionalmente seco del piso bajo de Morropón (${M.e6_cls.Leve} bloques leves y ${(M.e6_cls.Moderado || 0) + (M.e6_cls.Severo || 0)} moderados o severos), mientras que Huancabamba tiene ${H.e6_cls["No degradado"]} de 48 bloques no degradados, mayormente de matorral andino y mosaico agroforestal. Este contraste valida los dos regímenes territoriales del primer entregable y la necesidad de paquetes de medidas diferenciados.`));

y.push(h2("4.5 Suelos integrados al bloque (E5)"));
y.push(para(`El Informe de Avance de la Fase de Campo del E5 se llevó al bloque: Huancabamba cuenta con ${H.cal_n} calicatas (${H.cal_t.B} de bloque, ${H.cal_t.R} de referencia y ${H.cal_t.S} en Lotes SUS) en ${H.cal_bloques} bloques, y 43 de sus 48 bloques tienen alguna unidad de suelo muestreada; Morropón cuenta con ${M.cal_n} calicatas (${M.cal_t.B} de bloque, ${M.cal_t.R} de referencia y ${M.cal_t.S} en Lotes SUS) en ${M.cal_bloques} bloques, y 50 de 59 bloques con unidad muestreada. Los bloques sin calicata propia se caracterizan por analogía con la unidad muestreada de igual litología y clima, sin imputar valores.`));
y.push(N.cuadro("Caracterización preliminar de campo de las calicatas por provincia"));
y.push(tabla([{ w: 3238 }, { w: 3200 }, { w: 3200 }], ["VARIABLE", "HUANCABAMBA (28 calicatas)", "MORROPÓN (39 calicatas)"], [
  ["Orden taxonómico tentativo", "Entisol 14 · Inceptisol 10 · Mollisol 4", "Entisol 31 · Inceptisol 7 · Mollisol 1"],
  ["CUM de campo (grupo)", "Protección 13 · Forestal 10 · Pastoreo 4 · Cultivo permanente 1", "Protección 17 · Forestal 14 · Pastoreo 4 · Cultivo permanente 3 · Cultivo en limpio 1"],
  ["Erosión hídrica en campo", "Moderada 16 · Ligera 5 · Severa 4 · Muy ligera 3", "Moderada 18 · Ligera 15 · Severa 5 · Muy ligera 1"],
  ["Pruebas de infiltración asociadas", "INF-03 (0.41), INF-08 (13.25), INF-10 (2.20) e INF-12 (0.83) cm/h", "INF-01 (0.59), INF-02 (0.79), INF-02b (6.88), INF-04 (0.78), INF-05 (0.33), INF-06 (1.37), INF-07 (0.93) cm/h"],
]));
y.push(fuente("Fuente: pestaña «Suelos» de los Volúmenes I y II, sobre el Informe de Avance – Fase de Campo de GeoSIG Ingenieros (datos al 20/09/2026). Resultados preliminares, sujetos al laboratorio (Entregables N° 3 y N° 4)."));

y.push(h2("4.6 Tamizaje predial de gabinete"));
y.push(para("El Paso 5 del proceso de selección exige un máximo de 19 predios rurales interceptados por bloque y la exclusión de áreas agrícolas. Los mapas de situación predial v2 resuelven en gabinete la tenencia de los 107 bloques: superficie en comunidades campesinas (GEORURAL–MIDAGRI), en predios COFOPRI fuera de ellas y sin registro, asumida como Tierras del Estado y/o posesionarios, con tolerancia cartográfica de 5 %."));
y.push(N.cuadro("Situación predial de gabinete por provincia"));
const sits = ["Comunal", "Parcialmente comunal", "Privada", "Parcialmente privada", "Tierras del Estado y/o posesionarios"];
const sp = (P, k) => (P.pred_sit[k] ? `${P.pred_sit[k][0]} · ${fmt(P.pred_sit[k][1])}` : "0 · 0.00");
y.push(tabla([{ w: 3438 }, { w: 2000, align: R }, { w: 2000, align: R }, { w: 2200, align: R }], ["SITUACIÓN PREDIAL (bloques · ha)", "HUANCABAMBA", "MORROPÓN", "TOTAL"], [
  ...sits.map((k) => [k, sp(H, k), sp(M, k), `${(H.pred_sit[k] ? H.pred_sit[k][0] : 0) + (M.pred_sit[k] ? M.pred_sit[k][0] : 0)} · ${fmt((H.pred_sit[k] ? H.pred_sit[k][1] : 0) + (M.pred_sit[k] ? M.pred_sit[k][1] : 0))}`]),
  ["Superficie en comunidad campesina (ha)", `${fmt(H.pred_cc)} (${pc(H.pred_cc, H.ha)})`, `${fmt(M.pred_cc)} (${pc(M.pred_cc, M.ha)})`, fmt(H.pred_cc + M.pred_cc)],
  ["Superficie en predios COFOPRI (ha)", fmt(H.pred_cof), fmt(M.pred_cof), fmt(H.pred_cof + M.pred_cof)],
  ["Superficie sin registro (ha)", fmt(H.pred_est), fmt(M.pred_est), fmt(H.pred_est + M.pred_est)],
  ["Predios COFOPRI interceptados", `${H.pred_npred}`, `${M.pred_npred}`, `${H.pred_npred + M.pred_npred}`],
  ["Bloques con más de 19 predios COFOPRI", H.pred_gt19.map(([q, n]) => `${q} (${n})`).join(", "), "Ninguno", `${H.pred_gt19.length}`],
  ["Bloques con tenencia declarada (F-DS-01) distinta del gabinete", "3", "17", "20"],
  ["Lotes SUS preliminares asociados", `${H.lotes.nl} lotes · ${fmt(H.lotes.hl)} ha`, `${M.lotes.nl} lotes · ${fmt(M.lotes.hl)} ha`, `${H.lotes.nl + M.lotes.nl} lotes · ${fmt(H.lotes.hl + M.lotes.hl)} ha`],
]));
y.push(fuente("Fuente: Índice de mapas de situación predial de los 117 bloques V6, v2 (11/10/2026; transcripción verificada contra los totales del índice: 12,270.235 ha, 328 predios); consolidado DS del 17/10/2026; catálogo de liberación de áreas con 60 lotes SUS preliminares."));
y.push(para("La discrepancia de tenencia es material. En Huancabamba, la F-DS-01 del bloque 63 declara posesión individual cuando el gabinete lo ubica 100 % en comunidad campesina, la del bloque 81 se declara comunal sin superposición y la del bloque 60 declara superposición comunal en un bloque registrado como privado. En Morropón, doce bloques (49, 77, 83 a 87, M3B1, M3B6, M28B2, M28B3 y M28B4) declaran no superponerse a tierras comunales cuando el gabinete los ubica 100 % en comunidad campesina, y once fichas declaran más de 19 predios individuales, frente a solo 21 predios COFOPRI en toda la provincia, diferencia atribuible a posesionarios sin título. El tamizaje de campo, que la ANIN ejecuta por Administración Directa con el equipo de liberación de áreas que ingresa en la segunda quincena de octubre, es la única vía para resolverla antes de las actas."));
y.push(figura(FIG + "SP-042_Bloque_M4B4.jpg", 16.5, "jpg"));
y.push(N.figura("Lámina SP-042 — Mapa de situación predial del bloque M4B4 (Huarmaca): 99.9 % en comunidad campesina (ID 2735)"));
y.push(fuente("Fuente: Mapas de situación predial v2, Subentregable 2.2 (ANIN – DIME – SESDI, octubre de 2026). Información de gabinete sujeta a verificación de campo."));
y.push(figura(FIG + "SP-091_Bloque_87.jpg", 16.5, "jpg"));
y.push(N.figura("Lámina SP-091 — Mapa de situación predial del bloque 87 (San Juan de Bigote), parte del conglomerado severo 83–87"));
y.push(fuente("Fuente: Mapas de situación predial v2, Subentregable 2.2 (ANIN – DIME – SESDI, octubre de 2026)."));

y.push(h2("4.7 Diagnóstico social integrado"));
y.push(N.cuadro("Cobertura del diagnóstico social por provincia (consolidado del 17/10/2026)"));
y.push(tabla([{ w: 3638 }, { w: 3000, align: R }, { w: 3000, align: R }], ["VARIABLE", "HUANCABAMBA", "MORROPÓN"], [
  ["Fichas vigentes (F-DS-01 / F-DS-02 / F-DS-03)", "93 (48 / 27 / 18)", "252 (135 / 63 / 54)"],
  ["Bloques con ficha", "24 de 48", "52 de 59"],
  ["Centros poblados del catálogo INEI cubiertos", "28 de 58", "68 de 78"],
  ["Bloques sin ficha", "24 (incluye M4B4, M12B1, M4B3, M20B1, M2B8 y M30B5)", "7 (55, 58, M2B1, M17B4, M36B2, M6B2-1, M6B2-3)"],
  ["Instrumentos aplicados de la Plantilla V4", "3 de 7 (F-DS-04 a F-DS-07 no aplicados)", "3 de 7 (F-DS-04 a F-DS-07 no aplicados)"],
  ["Observaciones de control de calidad del consolidado", "44", "177"],
]));
y.push(fuente("Fuente: pestaña «Cobertura» y novedades de la actualización del 17/10/2026, Volúmenes I y II."));

y.push(h2("4.8 Riesgo y exposición por bloque"));
y.push(para(`La síntesis preliminar de gabinete (peligro × exposición, Anexo 2 GdR-CCC de la Guía General) califica en Huancabamba ${H.riesgo.Alto} bloques con riesgo Alto, ${H.riesgo.Medio} Medio y ${H.riesgo.Bajo} Bajo, y en Morropón ${M.riesgo.Alto} Alto, ${M.riesgo.Medio} Medio y ${M.riesgo.Bajo} Bajo. La exposición es mayoritariamente «aguas abajo por delimitar» (${H.expo["Aguas abajo (por delimitar)"]} y ${M.expo["Aguas abajo (por delimitar)"]} bloques), con ${H.expo.Interfaz + M.expo.Interfaz} bloques de interfaz y ${H.expo.Interna + M.expo.Interna} con exposición interna documentada; las propias MRR-CCC son elementos expuestos y deben evaluarse en el E8.`));

y.push(h2("4.9 Cartografía de ecosistemas y bloques V6"));
y.push(figura(FIG + "IN_Piura_03_ecosistemas_huancabamba.jpg", 16.5, "jpg"));
y.push(N.figura("Ecosistemas y bloques V6 — provincia de Huancabamba (lámina 3)"));
y.push(fuente("Fuente: Mapa Nacional de Ecosistemas (MINAM 2018, R.M. N° 440-2018-MINAM) recortado a microcuencas; bloques V6 (ANIN – DIME – SESDI). UTM WGS 84 Zona 17S."));
y.push(figura(FIG + "IN_Piura_04_ecosistemas_morropon.jpg", 16.5, "jpg"));
y.push(N.figura("Ecosistemas y bloques V6 — provincia de Morropón (lámina 4)"));
y.push(fuente("Fuente: Mapa Nacional de Ecosistemas (MINAM 2018) recortado a microcuencas; bloques V6 (ANIN – DIME – SESDI). UTM WGS 84 Zona 17S."));

y.push(h1("5. VACÍOS DE INTEGRACIÓN Y ACCIONES"));
y.push(para(`Los vacíos que impiden cerrar la Fase de Identificación en el ámbito se registran en la pestaña «Integralidad» de cada volumen. El Cuadro N° ${N.sigC()} los consolida para ambas provincias.`));
y.push(N.cuadro("Vacíos de integración y acciones para cerrar la Fase de Identificación"));
const gap = {};
[...H.gaps.map((g) => ["Huancabamba", ...g]), ...M.gaps.map((g) => ["Morropón", ...g])].forEach(([p, k, sit, efe, acc, pr]) => {
  if (!gap[k]) gap[k] = { sit: [], acc, pr }; gap[k].sit.push(`${p}: ${sit}`);
});
y.push(tabla([{ w: 1700 }, { w: 4338 }, { w: 2700 }, { w: 900, align: C }], ["ESTUDIO", "SITUACIÓN AL CORTE", "ACCIÓN", "PRIORIDAD"],
  Object.entries(gap).map(([k, g]) => [k, g.sit.join("\n"), g.acc, g.pr]), { size: 15, cantSplit: false }));
y.push(fuente("Fuente: pestaña «Integralidad», sección «Vacíos de integración», Volúmenes I y II, Versión 5."));

y.push(h1("6. HERRAMIENTAS Y PRODUCTOS GEOESPACIALES DEL PERIODO"));
y.push(para("La consolidación se apoyó en el aplicativo IN Piura (repositorio del proyecto) y en productos cartográficos generados en el periodo, todos en UTM WGS 84 Zona 17S:"));
[
  ["Módulo «Liberación de áreas» (27/09–06/10/2026): ", "tamizaje predial integrado al aplicativo con el catálogo de 117 bloques y 60 lotes SUS, tres vías de registro, plantillas Excel, reportes y sincronización de envíos KoboToolbox/ODK (formularios F-LA)."],
  ["Catálogo V6 (27/09/2026): ", "retiro de los 12 bloques descartados y consolidación de los 117 bloques vigentes (12,270.235 ha)."],
  ["Mapas V6 (04–05/10/2026): ", "lámina general y tres provinciales de bloques y área de influencia aprobada (AI_aprobado_2); láminas A3 de ecosistemas por bloque; Excel ANIN de bloques por distrito y GeoPackage de ecosistemas por bloque."],
  ["Diagnóstico social (05–08/10/2026): ", "ajustes de la ficha F-DS-01 (actividades económicas en porcentaje, mano de obra, educación), conteo único de centros poblados y tenencia por bloque."],
  ["Mapas de situación predial v2 (11/10/2026): ", "117 láminas A3 SP-001 a SP-117 con índice en Excel."],
  ["Volúmenes de diagnóstico publicados: ", "Versión 5 de Huancabamba, Morropón y Ayabaca, con fichas por bloque, cartografía de PI y matrices descargables."],
].forEach(([t, d]) => y.push(bullet([b(t), d])));

y.push(h1("7. MEDIOS VERIFICADORES"));
y.push(N.cuadro("Medios verificadores de la actividad b)"));
y.push(tabla([{ w: 3200 }, { w: 4238 }, { w: 2200, align: C }], ["MEDIO VERIFICADOR", "UBICACIÓN / IDENTIFICADOR", "FECHA"], [
  ["Diagnóstico Huancabamba Versión 5 (Vol. II)", "https://claude.ai/artifact/9828ogNDewfGQLgQDCq7kn", "17/10/2026"],
  ["Diagnóstico Morropón Versión 5 (Vol. I, rev. 10)", "https://claude.ai/artifact/Y5QRXCQRsXFvwFgYL1oZ64", "17/10/2026"],
  ["Diagnóstico Ayabaca Versión 5 (Vol. III, contraste)", "https://claude.ai/artifact/4kF1tRzj2FkuDgbcuTcBye", "17/10/2026"],
  ["Índice de mapas de situación predial V6 v2 (Excel)", "Google Drive ID 1jkNZ0sikCKQXZyUu4jlaiekNU82eQXE2", "11/10/2026"],
  ["Láminas SP-001 a SP-117 (PNG, A3)", "Google Drive, carpetas de mapas de situación predial (p. ej., SP-042: 1g2ax3ZxRLfZvhH-OEcPyh97se6gHGIj-; SP-091: 1eVz0dnA5gcYFZDWue7DQS6tWtLlxhyvJ)", "11/10/2026"],
  ["Informes E2 Diagnóstico social, Vol. I y II (B1) y matrices", "Google Drive ID 1AnXV3tTvQphvFLGGW4AVWwhRGRp0aJbv; 1DY7r3RBZy5lYEX6DkTakAaO3Pxu9oKt2; 11-tKOKJYgPF-uGGF7nRkWuxR5kjfJ6Xg", "10/10/2026"],
  ["Informe E6 y matrices del diagnóstico del ecosistema", "Google Drive ID 1NfYjSV3QXI3uA4eHsMNQNTw9vdY5RdDA; 1Uie1il6FeHI4SD8QhKXsfDof7cm4NF8J", "15/10/2026"],
  ["Mapas y Excel de bloques V6 (ecosistemas, área de influencia)", "Repositorio del aplicativo IN Piura, carpeta mapas/salidas (commits 0163a3a a 77a0266)", "04–05/10/2026"],
  ["Módulo de liberación de áreas y catálogo V6", "Repositorio del aplicativo IN Piura (commits 042679b, 9a4d846, b3d9dcf, 33ae71e, ae1f754)", "27/09–06/10/2026"],
  ["Matrices nativas consolidadas del presente anexo", "Anexo 4 (Excel): hojas Bloques_107, Distritos y Situación_predial", "Octubre 2026"],
], { size: 15 }));
y.push(fuente("Fuente: elaboración propia."));

y.push(h1("8. CONCLUSIONES DEL ANEXO"));
[
  `La información geoespacial y el diagnóstico territorial de los 107 bloques de Huancabamba y Morropón (${fmt(TOT)} ha) quedaron actualizados, integrados y consolidados en la Versión 5, con nueve fuentes leídas por bloque y fecha de corte explícita.`,
  `El estado del ecosistema, la situación predial de gabinete y los suelos de campo —ausentes en la línea base del primer entregable— están disponibles por bloque; el tamizaje predial pasó de 0 a 107 bloques en gabinete, aunque sin actas ni verificación de campo.`,
  `La brecha y el peligro convergen en el bosque seco del piso bajo de Morropón y en el núcleo de Huarmaca y San Miguel de El Faique; los bloques de altura de Huancabamba se justifican por la regulación de riesgos más que por la brecha de cobertura.`,
  "Antes de fijar metas físicas deben resolverse cuatro condicionantes: el criterio único de brecha, la tenencia y los bloques con más de 19 predios, los filtros formales de exclusión y el inventario de cárcavas en campo.",
].forEach((t, k) => y.push(num(`8.${k + 1}.`, t)));

// --- Anexo 1-A: tabla por bloque (horizontal) ---
const z = [];
z.push(h1("ANEXO 1-A. CONSOLIDADO POR BLOQUE — HUANCABAMBA Y MORROPÓN"));
z.push(para("Una fila por bloque del catálogo V6. Coordenadas del centroide en UTM WGS 84 Zona 17S (validadas: Este entre 450,000 y 750,000 m; Norte entre 9,300,000 y 9,600,000 m). MSAVI: media del bloque 2024. PI: clase del índice ponderado. IC y clase: índice de condición del E6. «% bajo umbral»: superficie del bloque bajo MSAVI 0.4976. Calicata: «Sí» calicata de bloque; «Ref.» solo referencia o Lote SUS."));
z.push(N.cuadro("Consolidado de los 107 bloques preliminares de intervención (Versión 5)"));
const cb = [{ w: 1050 }, { w: 1450 }, { w: 800 }, { w: 1150 }, { w: 850, align: R }, { w: 850, align: R }, { w: 1000, align: R }, { w: 500, align: C }, { w: 650, align: C }, { w: 850, align: C }, { w: 600, align: C }, { w: 900, align: C }, { w: 650, align: C }, { w: 1550 }, { w: 620, align: C }, { w: 600, align: C }, { w: 500, align: C }];
const tw = cb.reduce((s, c) => s + c.w, 0); cb[13].w += WH - tw;
const bl = [...D.bloques].sort((a, c) => (a.prov + a.d).localeCompare(c.prov + c.d) || a.b.localeCompare(c.b, "es", { numeric: true }));
z.push(tabla(cb, ["PROV.", "DISTRITO", "BLOQUE", "MICRO-CUENCA", "SUP. (ha)", "ESTE (m)", "NORTE (m)", "RÉG.", "MSAVI", "PI", "IC", "CLASE E6", "% BAJO UMBRAL", "SITUACIÓN PREDIAL", "PRE-DIOS", "CALI-CATA", "F-DS"],
  bl.map((r) => [r.prov === "Huancabamba" ? "Huanc." : "Morrop.", r.d, r.b, r.mc, fmt(r.ha), fmt(r.e, 0), fmt(r.n, 0), r.reg, fmt(r.ms, 3), r.pi, fmt(r.ic, 2), r.est, fmt(r.bu, 1), r.sit, `${r.npred}`, r.cal, r.ds]),
  { total: ["TOTAL", "14 distritos", "107", "", fmt(TOT), "", "", "", "", "", "", "", "", "", `${H.pred_npred + M.pred_npred}`, "", ""], size: 13 }));
z.push(fuente("Fuente: matrices por bloque de los Volúmenes I y II (Versión 5); Tablero E6; índice de mapas de situación predial v2; Informe de Avance del E5; consolidado DS del 17/10/2026. Hoja «Bloques_107» del Anexo 4."));

L.guardar([
  seccionCaratula(DEF.anexo1),
  L.seccion(x, { pie: PIE }),
  L.seccion(xh, { pie: PIE, horizontal: true }),
  L.seccion(y, { pie: PIE }),
  L.seccion(z, { pie: PIE, horizontal: true }),
], OUT, "Anexo N° 1 – Actividad b)");
