"""Contenido de los árboles de problemas del Proyecto IN Piura (fuente única).

    python arboles_problemas/contenido_arboles.py   ->  arboles_problemas/arboles_IN_Piura.json

Cada nodo lleva su texto, la evidencia del diagnóstico que lo sustenta, la fuente, su lectura en el
análisis de riesgo (AdR-CCC) y la línea de intervención del proyecto que lo atiende. Las cifras se
transcriben de los informes del Entregable 2 (B2), del Informe EVAR y de los diagnósticos V5; donde
las fuentes discrepan se indica y se marca [POR DEFINIR].
"""
import json
from pathlib import Path

AQUI = Path(__file__).resolve().parent

# ------------------------------------------------------------------ fuentes
F = {
    "E2-I": "Informe E2 Vol. I Morropón, B2 Rev. hscm (28-oct-2026)",
    "E2-II": "Informe E2 Vol. II Huancabamba, B2 Rev. hscm (28-oct-2026)",
    "E2-III": "Informe E2 Vol. III Ayabaca, B2 Rev. hscm (30-oct-2026)",
    "DX-M": "Diagnóstico Morropón Versión 5 (artefacto, oct-2026): DT, PI, suelos E5, ecosistemas E6, hidroclima",
    "DX-H": "Diagnóstico Huancabamba Versión 5 (artefacto, oct-2026): DT, PI, suelos E5, ecosistemas E6, hidroclima",
    "DX-A": "Diagnóstico Ayabaca Versión 5 (artefacto, oct-2026): DT, PI, suelos E5, ecosistemas E6, hidroclima",
    "EVAR": "Informe EVAR Piura – evaluación del riesgo por movimientos en masa, CENEPRED (19-oct-2026)",
    "MAT": "Matrices E2 IN Piura B2 (24-oct-2026)",
    "ZZ05": "ANIN ZZ-000005-R00, contenidos mínimos E7, Fig. 2 (modelo de referencia del árbol de problemas IN)",
    "ZZ06": "ANIN ZZ-000006-R00, contenidos mínimos del capítulo de Identificación (E9), num. 10.1",
    "F04A": "MINAM, Formato N° 04-A, indicador de brecha (RM N° 00213-2024-MINAM)",
    "AVA": "Avance del Perfil IN Piura · SESDI, 28-oct-2026 (línea base preliminar, criterio areal E6)",
    "DI": "Diagnóstico Integral IN Piura V2 (artefacto, oct-2026)",
}

LINEA = {
    "verde": "Línea 1 · Infraestructura natural (verde): revegetación y reforestación con especies nativas; viveros permanentes y temporales.",
    "marron": "Línea 2 · Infraestructura marrón: zanjas de infiltración, terrazas de formación lenta, diques para control de cárcavas, clausura temporal y manejo de pastizales, barreras y cercos vivos.",
    "complementaria": "Línea 2.1 · Obras complementarias: estabilización de taludes, muros y gaviones, diques en quebradas, drenaje y control de escorrentía (coordinadas con el titular de la vía cuando corresponda).",
    "gob": "Línea 3 · Gobernanza y gestión comunitaria: planes de manejo comunitario, fortalecimiento de capacidades en GdR, alerta temprana comunitaria y MERESE.",
    "gob_predial": "Línea 3 · Gobernanza: acuerdos de uso y actas de libre disponibilidad, tamizaje predial y mesa de diálogo previa a fijar metas en áreas en disputa.",
    "efecto": "Se reduce al recuperar el servicio de regulación de riesgos (objetivo central); se mide con los indicadores de fin.",
    "exogeno": "No es modificable por el proyecto. Se considera en el AdR-CCC (E8) y en el diseño de las MRR-CCC.",
}

ADR = {
    "condicionante": "Peligro · factor condicionante modificable (cobertura, suelo, drenaje)",
    "fragilidad": "Vulnerabilidad · fragilidad (prácticas y condiciones que debilitan al ecosistema y a la población)",
    "resiliencia": "Vulnerabilidad · baja resiliencia (capacidad de organización, gestión y respuesta)",
    "exposicion": "Exposición · elementos y medios de vida expuestos",
    "peligro": "Peligro · manifestación del evento (movimientos en masa, erosión, flujos)",
}


def n(id_, texto, sustento=(), fuentes=(), adr=None, linea=None, prov=None, **extra):
    d = {"id": id_, "texto": texto, "sustento": list(sustento), "fuentes": [F[f] if f in F else f for f in fuentes]}
    if adr:
        d["adr"] = ADR.get(adr, adr)
    if linea:
        d["linea"] = LINEA.get(linea, linea)
    if prov:
        d["prov"] = prov
    d.update(extra)
    return d


# ================================================================== AYABACA
ayabaca = {
    "pestana": "Ayabaca",
    "sigla": "10 bloques",
    "titulo": "Árbol de problemas – Provincia de Ayabaca (distrito de Frías)",
    "subtitulo": "Cuenca alta del río Piura · 10 bloques · 1,391.70 ha · 4 microcuencas · 17 centros poblados (2,541 hab., Censo 2017)",
    "fuente": "Fuentes: Informe E2 Vol. III Ayabaca (30-oct-2026); Diagnóstico Ayabaca V5; Informe EVAR (19-oct-2026)",
    "ficha": [
        {"k": "Bloques", "v": "10 (todos en Frías)"},
        {"k": "Superficie", "v": "1,391.70 ha"},
        {"k": "Centros poblados", "v": "17 (2,541 hab.)"},
        {"k": "PI alto + muy alto", "v": "782.64 ha (56.2 %)"},
        {"k": "PMM alto + muy alto (EVAR)", "v": "99.4 % de los bloques"},
        {"k": "Bajo umbral MSAVI", "v": "525.21 ha (37.7 %)"},
    ],
    "lectura": (
        "En Frías el peligro se concentra en el piedemonte de bosque seco (régimen A, índice PI 3.968), donde el corredor "
        "Chulucanas–Frías y las trochas sin drenaje activan taludes, y en el mosaico agrícola de ladera (68 % del ámbito es "
        "zona agrícola). El PI no sigue a la pendiente (r = +0.126) sino a la pérdida de cobertura (r = −0.968 con el MSAVI): "
        "por eso la cobertura y el suelo son las causas que el proyecto puede revertir."
    ),
    "efecto_final": n(
        "EF",
        "Incremento de los costos de reconstrucción y disminución de la resiliencia de la población de Frías y sus medios de vida "
        "frente a movimientos en masa e inundaciones desencadenados por lluvias extremas, intensificadas por el cambio climático",
        ["2,541 hab. en 17 CP rurales (Censo 2017); 400 personas vulnerables (15.7 %); proyección 2026: 2,376 hab.",
         "Riesgo por movimientos en masa en las áreas de influencia: alto en 745.65 ha (90.9 %); las 10 áreas con riesgo alto dominante (EVAR, cálculo por provincia).",
         "Fragilidad alta en Rincón de Geraldo, Guabal y Guanábano Alto; vivienda de adobe y saneamiento precario (EVAR).",
         "Las pérdidas no están valoradas en ninguna fuente: [POR DEFINIR] con registro SIGRID/INDECI de eventos."],
        ["E2-III", "EVAR", "ZZ05"], linea="efecto"),
    "efectos_indirectos": [
        n("EI1", "Daños a taludes y vías, viviendas, líneas de conducción de agua y tendido eléctrico; cortes de servicio en temporada de lluvias",
          ["Taludes viales activos con desprendimientos en el corredor Chulucanas–Frías (M17B6, M17B10).",
           "Vivienda junto a camino erosionado (bloque 3); tubería de conducción en ladera (bloque 6, El Guabo); tendido eléctrico (M17B6).",
           "Cortes de energía en lluvias (Rincón de Geraldo)."],
          ["DX-A", "E2-III"], adr="exposicion", linea="efecto", padres=["ED1"]),
        n("EI2", "Mayor aporte de sedimentos y colmatación de quebradas que drenan al valle del Alto Piura",
          ["Bloques 3, 6 y 36 declarados «zona de origen que aporta escorrentía concentrada a las quebradas» (informes DT).",
           "Cauce torrencial con bolonería y socavamiento en M17B10.",
           "Hipótesis a cuantificar con el estudio hidrológico (E4): [POR DEFINIR]."],
          ["DX-A"], adr="exposicion", linea="efecto", padres=["ED3"]),
        n("EI3", "Pérdida de suelo productivo y de rendimiento de cultivos (maíz, frejol, café, palta, cacao) que sostienen a las familias",
          ["Agricultura de secano en el 100 % de las localidades; destino mixto autoconsumo/mercado.",
           "Calicata degradada CAL-13 (bloque 36): Entisol de 61 cm con erosión severa frente a referencia Inceptisol de 84 cm.",
           "En el bloque 39 la degradación eliminó el epipedón mólico (CAL-59 frente a CAL-60)."],
          ["E2-III", "DX-A"], adr="exposicion", linea="efecto", padres=["ED2"]),
        n("EI4", "Pérdida de otros servicios del ecosistema: regulación hídrica de fuentes, hábitat de especies en peligro crítico y carbono",
          ["7 bloques en zona de recarga; captación de La Cría a 419 m del bloque 39.",
           "Hualtaco, palo santo y sapote (En Peligro Crítico) en M17B5, M17B6, M17B7 y M17B10.",
           "Son efectos colaterales: la regulación hídrica es otra tipología y no es el fin de este proyecto."],
          ["DX-A", "ZZ05"], linea="efecto", padres=["ED2", "ED3"]),
    ],
    "efectos_directos": [
        n("ED1", "Mayor ocurrencia de movimientos en masa (desprendimientos, caídas de rocas y deslizamientos) en taludes, laderas y piedemonte",
          ["Peligro por movimientos en masa alto o muy alto en el 99.4 % de los bloques (muy alto 73.2 %) y en el 93.1 % de las áreas de influencia (EVAR).",
           "Taludes de 4–8 m en saprolito granítico con desprendimientos activos (M17B6).",
           "Antecedente de reubicación de un centro poblado por deslizamientos en lluvias (bloque 56; CP y fecha por precisar)."],
          ["EVAR", "DX-A"], adr="peligro", linea="efecto"),
        n("ED2", "Erosión hídrica acelerada y pérdida de suelo en laderas (surcos y cárcavas)",
          ["7 cárcavas inventariadas: 1,154.94 m en 6 bloques (M17B7: 570.9 m; 36: 258.4 m; 3: 217.7 m).",
           "Bloque 27: erosión hídrica «peligro dominante», cárcavas activas CAR-01 y CAR-02, labranza a favor de la pendiente, dosel ≈ 5 %.",
           "Surcos sobre huellas de ganado en la franja occidental de M17B5."],
          ["DX-A"], adr="peligro", linea="efecto"),
        n("ED3", "Mayor escorrentía superficial concentrada y carga de sedimentos desde las zonas de origen hacia las quebradas",
          ["Índice de Fournier 152.6 (agresividad alta); 85–91 % de la lluvia entre enero y abril.",
           "Peligro generador de inundaciones declarado Medio en 3, 6, 27, 36 y M17B7 (informes DT).",
           "No existe aún el modelo de PGI por bloque: [POR DEFINIR]."],
          ["DX-A"], adr="peligro", linea="efecto"),
    ],
    "problema": n(
        "PC",
        "La población de Frías (provincia de Ayabaca) no accede adecuadamente al servicio ecosistémico de regulación de riesgos naturales "
        "por la degradación del bosque estacionalmente seco y de los mosaicos agropecuarios de ladera en las zonas de origen de movimientos en masa e inundaciones",
        ["Peligro integrado alto o muy alto en 782.64 ha (56.2 % de 1,391.70 ha); índice provincial 3.580 (Alto).",
         "Superficie bajo el umbral MSAVI 0.4976: 525.21 ha (37.7 %) con el criterio areal E6 propuesto en el Avance; 181.90 ha si se toma la media de bloque.",
         "Correlación PI–MSAVI r = −0.968: donde se pierde cobertura, sube el peligro.",
         "94.4 % de la superficie es reversible total o parcialmente (1,314.03 ha)."],
        ["E2-III", "DX-A", "EVAR", "AVA", "F04A"],
        indicador="Magnitud: peligro integrado alto + muy alto en 782.64 ha (56.2 %) · línea base preliminar de brecha: 525.21 ha bajo el umbral MSAVI (37.7 %, criterio areal E6)"),
    "factores_exogenos": [
        n("EX1", "Pendiente media 57.5 % (64.6 % del área sobre 50 % y 204.37 ha sobre 75 %), la más fuerte del proyecto; susceptibilidad litológica a movimientos en masa alta en los 10 bloques; granito meteorizado (Súper Unidad Paltashaco).",
          ["Susceptibilidad litológica a erosión hídrica muy alta en 6 bloques (708.02 ha).", "El peligro integrado no sigue a la pendiente (r = +0.126) sino a la cobertura."],
          ["DX-A", "EVAR"], adr="Peligro · factores condicionantes no modificables", linea="exogeno",
          titulo="Factores condicionantes del peligro"),
        n("EX2", "Lluvias de El Niño 1983, 1998 y 2023 entre 2.7 y 4.9 veces la media; abril 2023: 765–1,082 mm; 84 % de probabilidad de un año ≥ 2 veces la media en 20 años; Tmáx +0.21 a +0.25 °C por década.",
          ["Precipitación media 557 mm/año, índice de aridez 0.39 (semiárido), coeficiente de variación 87 %.",
           "Déficit hídrico futuro en el escenario SSP5-8.5 (criterio de idoneidad del proyecto)."],
          ["DX-A"], adr="Peligro · factor desencadenante y cambio climático", linea="exogeno",
          titulo="Factor desencadenante y cambio climático"),
    ],
    "causas_directas": [
        n("CD1", "Degradación de la cobertura vegetal protectora en las zonas de origen",
          ["Bosque ralo o matorral abierto 36.5 % y herbazal seco 18.4 % del ámbito (S2DR4); dosel denso solo 17.9 %.",
           "Bajo umbral MSAVI: M17B6 81.8 %, M17B10 73.3 %, M17B7 48.4 % de su superficie.",
           "Bloque 27 «muy alterado/degradado»: 30 % de suelo desnudo y regeneración escasa."],
          ["DX-A", "E2-III"], adr="condicionante", linea="verde"),
        n("CD2", "Reducida capacidad de los suelos y laderas para controlar la erosión y la escorrentía",
          ["«La erosión activa documentada es mayoritariamente antrópico-vial: taludes, trochas y sendas».",
           "Pendiente media 57.5 %: la mayor parte de las laderas exige obras adaptadas a pendientes fuertes.",
           "Suelos sin análisis de laboratorio (materia orgánica, textura, factor K): [POR DEFINIR]."],
          ["DX-A"], adr="condicionante", linea="marron"),
        n("CD3", "Inadecuada gobernanza y gestión local del ecosistema y del riesgo",
          ["Ninguna localidad tiene comité de recursos naturales, espacio de concertación, sistema de alerta temprana ni plan de gestión del riesgo.",
           "Tenencia 83.9 % comunal, con posesionarios sin título y un polígono comunal sin nombre (1508)."],
          ["E2-III", "DX-A"], adr="resiliencia", linea="gob"),
    ],
    "causas_indirectas": [
        n("CI1.1", "Expansión de la agricultura de subsistencia de secano hacia laderas de fuerte pendiente",
          ["Causa principal en 7 de 10 informes DT (bloques 3, 6, 36, 27, M17B7, 39 y 56).",
           "Zona agrícola: 946.82 ha (68 % del ámbito); uso agrícola dentro de los bloques 3, 6, 27, 36 y M17B7.",
           "En el bloque 39 la frontera agropecuaria avanza desde La Cría hacia la cabecera."],
          ["E2-III", "DX-A"], adr="fragilidad", linea="verde", padre="CD1"),
        n("CI1.2", "Tala y extracción selectiva de especies del bosque seco, incluidas especies en peligro crítico",
          ["La Cría se separó de la C.C. San Andrés de Culcas por tala y linderos no respetados.",
           "Sectores con tala en M17B7; hualtaco, palo santo y sapote (En Peligro Crítico) en 4 bloques del piedemonte."],
          ["E2-III", "DX-A"], adr="fragilidad", linea="verde", padre="CD1"),
        n("CI1.3", "Pastoreo libre de vacunos y caprinos sin ordenamiento, que impide la regeneración natural",
          ["Ganadería vacuna en 100 % de las localidades; porcina 85.7 %; caprina/ovina 28.6 %.",
           "Causa principal en M17B5 (El Guabo); pastoreo extensivo en M17B10; compactación por ganado en M17B7.",
           "Carga ganadera: [POR DEFINIR]."],
          ["E2-III", "DX-A"], adr="fragilidad", linea="marron", padre="CD1"),
        n("CI1.4", "Baja capacidad de regeneración natural del bosque seco por aridez y alta variabilidad de las lluvias",
          ["Índice de aridez 0.39, déficit de 931 mm y coeficiente de variación de la lluvia de 87 % (90–99 % en ladera y piedemonte)."],
          ["DX-A"], adr="fragilidad", linea="verde", padre="CD1"),
        n("CI2.1", "Corredor vial y trochas con cortes de talud sin estabilización ni drenaje, que concentran la escorrentía",
          ["Corredor Chulucanas–Frías en mejoramiento con taludes sin tratamiento: causa principal en M17B6 y M17B10.",
           "Trochas con taludes inestables, cárcavas o badenes en M17B7, 6, 56 y 27; camino erosionado en 3 (4 de 10 informes DT)."],
          ["DX-A", "E2-III"], adr="condicionante", linea="complementaria", padre="CD2"),
        n("CI2.2", "Laderas cultivadas sin prácticas de conservación de suelos (labranza a favor de la pendiente)",
          ["Labranza a favor de la pendiente y cárcavas activas en el bloque 27.",
           "Uso agrícola dentro de 5 bloques; medidas compatibles propuestas por los agricultores: terrazas, barreras vivas, agroforestería."],
          ["DX-A", "E2-III"], adr="fragilidad", linea="marron", padre="CD2"),
        n("CI2.3", "Cárcavas activas sin control que progresan en cabecera",
          ["7 cárcavas con 1,154.94 m en 6 bloques; el bloque 27 tiene cárcavas en campo que no figuran en la capa (inventario parcial)."],
          ["DX-A"], adr="condicionante", linea="marron", padre="CD2"),
        n("CI3.1", "Limitado conocimiento y organización comunal para el manejo del ecosistema y la gestión del riesgo",
          ["Comité de recursos naturales: 0 % de las localidades; rondas desarticuladas en Putagas y La Cría.",
           "Ninguna localidad tiene sistema de alerta temprana ni plan de gestión del riesgo.",
           "La percepción del peligro de la población no se ha levantado (F-DS-03 y F-DS-06 sin aplicar): [POR DEFINIR]."],
          ["E2-III", "DX-A"], adr="resiliencia", linea="gob", padre="CD3"),
        n("CI3.2", "Limitada presencia y articulación institucional, sin espacios de concertación, y desconfianza hacia agentes externos",
          ["Ninguna localidad tiene espacio de concertación; AGRO RURAL presente en 0 localidades; presencia estatal baja en 3 y media en 4.",
           "Percepción de «abandono» en Guabal y Pampa de Ramada; desconfianza por una estafa previa en Huayabal."],
          ["E2-III"], adr="resiliencia", linea="gob", padre="CD3"),
        n("CI3.3", "Tenencia comunal con posesión informal y linderos en disputa",
          ["1,167.02 ha (83.9 %) en comunidades campesinas; 224.69 ha sin registro; 0 predios COFOPRI; polígono comunal 1508 sin nombre.",
           "150 predios en posesión sin título en Guanábano Alto; conflicto latente La Cría–San Andrés de Culcas.",
           "Las fichas declaran 36, 93 y 38 predios en los bloques 3, 6 y M17B7 (más que el máximo de 19)."],
          ["E2-III", "DX-A"], adr="fragilidad", linea="gob_predial", padre="CD3"),
    ],
    "notas": "Árbol provincial de Ayabaca (Frías). Problema central redactado según el modelo de referencia ANIN (ZZ-000005, Fig. 2), "
             "sin incluir la regulación hídrica en el enunciado. Las cifras provienen del Informe E2 Vol. III (30-oct-2026), del "
             "Diagnóstico Ayabaca V5 y del Informe EVAR. Pendientes: criterio único de brecha MSAVI, modelo PMM/EPH/PGI por bloque, "
             "registro de daños históricos, laboratorio de suelos y fichas F-DS-03 a F-DS-07.",
}

# ================================================================== HUANCABAMBA
huancabamba = {
    "pestana": "Huancabamba",
    "sigla": "48 bloques",
    "titulo": "Árbol de problemas – Provincia de Huancabamba",
    "subtitulo": "Canchaque, Huancabamba, Huarmaca, Lalaquiz y San Miguel de El Faique · 48 bloques · 3,452.21 ha netas (V8) · 58 centros poblados (6,573 hab., Censo 2017)",
    "fuente": "Fuentes: Informe E2 Vol. II Huancabamba (28-oct-2026); Diagnóstico Huancabamba V5; Informe EVAR (19-oct-2026)",
    "ficha": [
        {"k": "Bloques", "v": "48 en 5 distritos"},
        {"k": "Superficie", "v": "3,452.21 ha V8 (3,551.16 ha V7)"},
        {"k": "Centros poblados", "v": "58 (6,573 hab.)"},
        {"k": "PI alto + muy alto", "v": "2,353.73 ha (66.3 %)"},
        {"k": "PMM alto + muy alto (EVAR)", "v": "98.2 % de los bloques"},
        {"k": "Bajo umbral MSAVI", "v": "1,599.99 ha (45.1 %)"},
    ],
    "lectura": (
        "Huancabamba combina dos situaciones: un núcleo de bosque seco bajo 1,000 msnm en Huarmaca y San Miguel de El Faique "
        "(M4B4, M12B1, M4B3, M20B1, M2B8), con degradación difusa por pastoreo y leña y peligro muy alto, y un mosaico cafetalero "
        "de ladera con pendientes cercanas o superiores a 50 %, remociones activas, trochas y agricultura. La correlación "
        "MSAVI–PI (r = −0.785) muestra que la cobertura gobierna el peligro; la litología solo orienta el tipo de obra."
    ),
    "efecto_final": n(
        "EF",
        "Incremento de los costos de reconstrucción y disminución de la resiliencia de la población de Huancabamba y sus medios de vida "
        "frente a movimientos en masa e inundaciones desencadenados por lluvias extremas, intensificadas por el cambio climático",
        ["6,573 hab. en 58 CP (Censo 2017); 1,245 personas vulnerables (18.9 %); 27 CP (2,872 hab.) dentro del área de influencia aprobada.",
         "Riesgo por movimientos en masa: alto en 2,656.10 ha (78.3 %) y muy alto en 56.92 ha de las áreas de influencia; las dos únicas áreas en riesgo muy alto del proyecto están en Huarmaca (Chococa y Cambruran) (EVAR).",
         "Nueve de las diez áreas de influencia con mayor riesgo del proyecto son de Huancabamba (EVAR, Tabla 35).",
         "Pérdidas económicas no valoradas: [POR DEFINIR]."],
        ["E2-II", "EVAR", "ZZ05"], linea="efecto"),
    "efectos_indirectos": [
        n("EI1", "Interrupción de vías y aislamiento estacional de caseríos, con pérdida de cosechas de café que no llegan al mercado",
          ["Interrupción de la carretera en 11 localidades; caseríos aislados (7) y café sin vender (6).",
           "Shuturumbe aislado 8 días por un huaico (marzo 2024); la quebrada La Mala aisló al C.P. Ranchos en 2026 con pérdida de café."],
          ["E2-II", "DX-H"], adr="exposicion", linea="efecto", padres=["ED1", "ED3"]),
        n("EI2", "Daños a viviendas, parcelas, ganado, captaciones y líneas de conducción de agua",
          ["Daños a chacras, ganado y viviendas en 11 localidades.",
           "Exposición interna en 3 bloques (M4B4: vivienda y corral en lecho de quebrada al pie de talud inestable; 72; 67) y de interfaz en 12 (19, 37, 71, M22B1, 35, 8, 10, 51, 73, 76, M2B8)."],
          ["E2-II", "DX-H"], adr="exposicion", linea="efecto", padres=["ED1"]),
        n("EI3", "Pérdida de suelo agrícola y de productividad de la economía cafetalera y de subsistencia",
          ["Cárcavas en parcelas (6 localidades) y terrenos agrícolas arrastrados (5).",
           "AGRO RURAL entrega urea en Shuturumbe por suelos degradados; quintal de café ≈ S/ 1,000–1,200."],
          ["E2-II"], adr="exposicion", linea="efecto", padres=["ED2"]),
        n("EI4", "Aporte de sedimentos que deteriora quebradas y fuentes de agua, y pérdida de otros servicios del ecosistema",
          ["Menos manantiales percibidos en 16 localidades; la población prioriza el agua (11 de 13 entrevistas).",
           "Sedimentación aguas abajo por verificar con el E4: [POR DEFINIR].",
           "Efecto colateral: la regulación hídrica es otra tipología y no es el fin de este proyecto."],
          ["E2-II", "DX-H", "ZZ05"], linea="efecto", padres=["ED2", "ED3"]),
    ],
    "efectos_directos": [
        n("ED1", "Mayor ocurrencia de movimientos en masa (deslizamientos, derrumbes, reptación y flujos de detritos) en laderas y quebradas de cabecera",
          ["Peligro por movimientos en masa alto o muy alto en el 98.2 % de los bloques (muy alto 69.7 %) y en el 83.3 % de las áreas de influencia (EVAR).",
           "Remociones activas en 16 bloques, escarpes activos en 34, deslizamientos antiguos en 15 y reptación en 18; flujo de detritos en el bloque 21 (pendiente 63.9 %).",
           "Coyona: movimientos en masa desde 1983; deslizamiento en la quebrada Potreros; derrumbes en Maray Grande."],
          ["EVAR", "DX-H", "E2-II"], adr="peligro", linea="efecto"),
        n("ED2", "Erosión hídrica acelerada y pérdida de suelo en laderas (surcos y cárcavas)",
          ["Cárcavas activas severas en el bloque 40; incipientes en 31, 80, 76, 72, 63, 30, 61 y M30B5.",
           "Capa de cárcavas: 9 cárcavas y 3,903.15 m en 8 bloques (bloque 4: 949 m; M20B1: 758 m; M4B4: 749 m).",
           "Calicatas someras frente a su referencia: bloque 44, 18 cm frente a 150 cm; bloque 19, 30 cm frente a 80 cm."],
          ["DX-H"], adr="peligro", linea="efecto"),
        n("ED3", "Mayor escorrentía torrencial y transporte de sedimentos desde las zonas de origen hacia las quebradas",
          ["Peligro generador de inundaciones alto en 31, 69 y 35, medio-alto en 5 bloques y medio en 33 (informes DT).",
           "Agresividad de la lluvia 134 en promedio y 303 en años El Niño (339 en el régimen de bosque seco).",
           "Infiltración básica muy baja: 0.41 cm/h en Sahuate Hualanga (76) y 0.83 cm/h en Pampa Quemada (M12B1)."],
          ["DX-H"], adr="peligro", linea="efecto"),
    ],
    "problema": n(
        "PC",
        "La población de la provincia de Huancabamba no accede adecuadamente al servicio ecosistémico de regulación de riesgos naturales "
        "por la degradación del bosque estacionalmente seco, el matorral andino y los mosaicos agroforestales de ladera en las zonas de origen de movimientos en masa e inundaciones",
        ["Peligro integrado alto o muy alto en 2,353.73 ha (66.3 %); 9 bloques en muy alto y 20 en alto; índice provincial 3.898 (Alto).",
         "Huarmaca concentra el 61 % de la superficie crítica (1,424.30 ha; 72.4 %).",
         "Núcleo de brecha (M4B4, M12B1, M4B3, M20B1, M2B8): 1,188.15 ha, MSAVI 0.340, 96.4 % en clase crítica.",
         "Bajo el umbral MSAVI 0.4976: 1,599.99 ha (45.1 %) con el criterio areal E6 propuesto en el Avance; 1,365.04 ha (10 bloques) si se toma la media de bloque."],
        ["E2-II", "DX-H", "EVAR", "AVA", "F04A"],
        indicador="Magnitud: peligro integrado alto + muy alto en 2,353.73 ha (66.3 %) · línea base preliminar de brecha: 1,599.99 ha bajo el umbral MSAVI (45.1 %, criterio areal E6)"),
    "factores_exogenos": [
        n("EX1", "Pendiente media 49 % (424.62 ha sobre 75 %; Canchaque 59.4 %); susceptibilidad litológica a movimientos en masa muy alta en 18 bloques y alta en 18 (Complejo de Olmos).",
          ["Litología con correlación débil con el PI (r = −0.24): orienta el tipo de obra, no la prioridad."],
          ["DX-H", "EVAR"], adr="Peligro · factores condicionantes no modificables", linea="exogeno",
          titulo="Factores condicionantes del peligro"),
        n("EX2", "Lluvias de El Niño 1983, 1998, 2017 y 2023 entre 2.0 y 2.7 veces la media provincial (3.2–3.3 veces en el bosque seco); tendencia +32.6 mm y +0.15 °C por década.",
          ["541 mm/año, índice de aridez 0.38; 82 % de la lluvia entre enero y abril; un año seco de cada tres en el régimen de bosque seco.",
           "Déficit hídrico futuro en el escenario SSP5-8.5."],
          ["DX-H"], adr="Peligro · factor desencadenante y cambio climático", linea="exogeno",
          titulo="Factor desencadenante y cambio climático"),
    ],
    "causas_directas": [
        n("CD1", "Degradación de la cobertura vegetal protectora en las zonas de origen",
          ["Suelo desnudo, roca y vías 576.1 ha; herbazal seco 909.6 ha; dosel denso solo 563.9 ha (S2DR4).",
           "Estado E6: 18 bloques con degradación leve (1,555 ha) y 6 con degradación moderada (619.5 ha).",
           "Los bloques bajo el umbral MSAVI tienen el 92.8 % de su superficie en clase crítica; el resto, el 49.7 %."],
          ["DX-H"], adr="condicionante", linea="verde"),
        n("CD2", "Reducida capacidad de los suelos y laderas para controlar la erosión y la escorrentía",
          ["Suelos francos con horizonte superficial pobre (ócrico), Entisoles e Inceptisoles someros.",
           "Erosión natural severa señalada en 33 de 48 informes DT.",
           "Materia orgánica, pH y factor K pendientes de laboratorio (UNALM): [POR DEFINIR]."],
          ["DX-H"], adr="condicionante", linea="marron"),
        n("CD3", "Inadecuada gobernanza y gestión local del ecosistema y del riesgo",
          ["Ninguna localidad tiene sistema de alerta temprana ni comité de gestión del riesgo; comité de recursos naturales solo en 13 %.",
           "71 % de la superficie sin registro predial."],
          ["E2-II", "DX-H"], adr="resiliencia", linea="gob"),
    ],
    "causas_indirectas": [
        n("CI1.1", "Sobrepastoreo caprino y vacuno sin manejo, que elimina la regeneración del bosque seco y del matorral",
          ["Pastoreo o sobrepastoreo en 45 de 48 informes DT; intensidad fuerte en M4B4 y M30B5 (80 %), 31 y 70 (70 %), 63 (40 %).",
           "Ganadería vacuna en 90.9 % de las localidades y caprina u ovina en 27.3 %."],
          ["DX-H", "E2-II"], adr="fragilidad", linea="marron", padre="CD1"),
        n("CI1.2", "Cambio de uso hacia agricultura de ladera (café, maíz, pastos)",
          ["Agricultura en ladera en 38 de 48 informes DT; bloque 43 100 % agrícola, M30B5 82 %, 43 50 %, 51 40 %.",
           "Zona agrícola: 1,265.97 ha (35.6 %) según el mapa MINAM."],
          ["DX-H", "E2-II"], adr="fragilidad", linea="verde", padre="CD1"),
        n("CI1.3", "Extracción de leña y madera y quemas para renovar pastos",
          ["Tala para leña o madera en 25 de 48 bloques; deforestación y venta de madera en 12 localidades.",
           "Venta de roble y laurel; tala ilegal de nogal y quinahuiro; el 70 % usa leña en Huamala Alto.",
           "Quemas para renovar pastos en 6 bloques (31, 51, 70, M22B1, 44, 21)."],
          ["DX-H", "E2-II"], adr="fragilidad", linea="verde", padre="CD1"),
        n("CI1.4", "Baja regeneración de especies nativas por aridez y avance de especies exóticas",
          ["Especies exóticas (pino, eucalipto, grevilea, higuerilla) en 16 bloques; sequía o cambio climático señalados en 12.",
           "Régimen de bosque seco: 414 mm/año, índice de aridez 0.24."],
          ["DX-H"], adr="fragilidad", linea="verde", padre="CD1"),
        n("CI2.1", "Trochas, vías y cortes de talud sin drenaje ni estabilización",
          ["Vías y taludes de corte en 40 de 48 informes DT; intensidad fuerte en 10, 8, 18 (trocha departamental 108), 20, 80, 71, 76 y 43.",
           "Desprendimientos en la trocha que afectan viviendas de Nuevo Progreso (bloque 19)."],
          ["DX-H"], adr="condicionante", linea="complementaria", padre="CD2"),
        n("CI2.2", "Laderas cultivadas sin prácticas de conservación y cárcavas sin control",
          ["Cárcavas activas severas en 40 e incipientes en 8 bloques; la ficha F-DT-02 solo cubre 26 bloques.",
           "Agroquímicos y aguas mieles del café en las quebradas de Huabal."],
          ["DX-H", "E2-II"], adr="fragilidad", linea="marron", padre="CD2"),
        n("CI2.3", "Pisoteo y compactación de suelos someros por el ganado, con baja infiltración",
          ["Infiltración básica de 0.41 a 2.2 cm/h en 3 de 4 ensayos; calicatas de 18 a 45 cm en bloques degradados."],
          ["DX-H"], adr="condicionante", linea="marron", padre="CD2"),
        n("CI3.1", "Limitado conocimiento y organización comunal para el manejo del ecosistema y la gestión del riesgo",
          ["Comité de recursos naturales en 13 % de 23 localidades; ninguna con alerta temprana ni comité de gestión del riesgo.",
           "Rondas desarticuladas en Shuturumbe y Sapce–Chamelico; prohibición de tala en Ullma sin control efectivo; faenas comunales solo reactivas.",
           "La población prioriza el agua (11 de 13 entrevistas); Cruz de Piedra sin señal celular para avisos."],
          ["E2-II", "DX-H"], adr="resiliencia", linea="gob", padre="CD3"),
        n("CI3.2", "Limitada gobernanza y articulación de las entidades públicas y privadas, con desconfianza por obras inconclusas",
          ["Sin espacios de concertación en los 24 CP con ficha; presencia estatal percibida baja en 11 localidades.",
           "Obra de agua y alcantarillado de Sapce y Chacchacal paralizada al 80 %; obras inconclusas en Huamala Alto y Flor de Café.",
           "Superposición con el proyecto del GORE en Pusmalca (CUI 2335868) sobre los bloques 81 y 79."],
          ["E2-II"], adr="resiliencia", linea="gob", padre="CD3"),
        n("CI3.3", "Alta informalidad en la tenencia de la tierra y presión de derechos mineros",
          ["71.0 % sin registro (2,451.17 ha); posesión sin título en 17 de 20 bloques; 5 bloques superan los 19 predios.",
           "Petitorio minero LM30 (bloques 8 y 71) y oposición activa a la minería (Río Blanco); San José de Hualcas no firmó."],
          ["E2-II"], adr="fragilidad", linea="gob_predial", padre="CD3"),
    ],
    "notas": "Árbol provincial de Huancabamba. Superficie V8 neta 3,452.21 ha (se excluyeron 98.95 ha del petitorio minero LM30); "
             "el PI se calculó sobre el polígono V7 (3,551.16 ha). Pendientes: criterio único de brecha MSAVI, modelo PMM/EPH/PGI "
             "por bloque, inventario de daños, delimitación de exposición aguas abajo, laboratorio de suelos y fichas sociales en Huarmaca.",
}


# ================================================================== MORROPÓN
morropon = {
    "pestana": "Morropón",
    "sigla": "59 bloques",
    "titulo": "Árbol de problemas – Provincia de Morropón",
    "subtitulo": "Buenos Aires, Chalaco, Chulucanas, Morropón, Salitral, San Juan de Bigote, Santa Catalina de Mossa, Santo Domingo y Yamango · 59 bloques · 7,327.37 ha (7,313.55 ha netas V8) · 79 centros poblados (9,537 hab.)",
    "fuente": "Fuentes: Informe E2 Vol. I Morropón (28-oct-2026); Diagnóstico Morropón V5; Informe EVAR (19-oct-2026); Avance del Perfil (28-oct-2026)",
    "ficha": [
        {"k": "Bloques", "v": "59 en 9 distritos"},
        {"k": "Superficie", "v": "7,327.37 ha (V8 neta 7,313.55 ha)"},
        {"k": "Centros poblados", "v": "79 (9,537 hab.)"},
        {"k": "PI alto + muy alto", "v": "5,335.20 ha (72.8 %)"},
        {"k": "PMM alto + muy alto (EVAR)", "v": "90.7 % de los bloques"},
        {"k": "Bajo umbral MSAVI", "v": "6,216.63 ha (84.8 %)"},
    ],
    "lectura": (
        "Morropón es bosque seco de baja altitud (Bes-cm en el 88 % del ámbito; 38 bloques bajo 1,000 msnm) con aridez extrema en "
        "el valle (índice 0.26). Su peligro dominante es la escorrentía torrencial que activa quebradas y genera huaicos, más la "
        "erosión laminar y en cárcavas; los movimientos en masa son sobre todo latentes, salvo en taludes viales y frentes rocosos con "
        "bolones sobre viviendas. El peligro lo explica la cobertura (r = −0.839 con el MSAVI), no la litología (r = −0.075). Los "
        "motores son el pastoreo caprino y la tala comercial de palo santo."
    ),
    "efecto_final": n(
        "EF",
        "Incremento de los costos de reconstrucción y disminución de la resiliencia de la población de Morropón y sus medios de vida "
        "frente a huaicos, movimientos en masa e inundaciones desencadenados por lluvias extremas, intensificadas por el cambio climático",
        ["9,537 hab. (8,832 censados en 76 CP más 705 en 3 CP sin código INEI); 1,640 personas vulnerables (18.6 %; 30.2 % en Santo Domingo).",
         "49 CP (5,730 hab., 64.9 %) dentro del área de influencia aprobada; la población aguas abajo está por delimitar en 47 bloques.",
         "Riesgo por movimientos en masa alto en 3,730.65 ha (28.9 %) de las áreas de influencia; vulnerabilidad alta en Guabo, Buenos Aires (0.594) y Serrán (EVAR, cálculo por provincia).",
         "Pérdidas económicas no valoradas: [POR DEFINIR]."],
        ["E2-I", "EVAR", "ZZ05"], linea="efecto"),
    "efectos_indirectos": [
        n("EI1", "Inundación de parcelas, arrastre de ganado y aislamiento estacional de caseríos por la activación de quebradas",
          ["Activación de quebradas en 24 localidades (53.3 % de las fichas F-DS-03): parcelas inundadas, ganado arrastrado, caseríos aislados.",
           "Huaro Quispampa, Mangomanguía, Río Seco y La Alberca quedan aislados de enero a abril.",
           "Viviendas reforzadas con sacos de arena aguas abajo de M7B1 (Víctor Raúl, Nuevo Progreso, Salitral)."],
          ["E2-I", "DX-M"], adr="exposicion", linea="efecto", padres=["ED3"]),
        n("EI2", "Daños a viviendas, colegios, vías, badenes y captaciones de agua en los centros poblados de cabecera",
          ["Huaicos y deslizamientos en 9 y 8 localidades: carreteras tapadas, daños en viviendas y chacras.",
           "El flujo del bloque 85 descarga sobre el casco urbano de Dótor y su colegio; quebradas del bloque 2 hacia el colegio N.° 14891; muros derruidos en Serrán (M7B6).",
           "Badenes y parcelas de La Peña (M9B1); captación de agua a 50 m de M32B3; bolones sueltos sobre viviendas en M6B10, M19B2, M32B3 y 58."],
          ["E2-I", "DX-M"], adr="exposicion", linea="efecto", padres=["ED1", "ED3"]),
        n("EI3", "Pérdida de suelo agrícola y de ingresos agropecuarios",
          ["Erosión y cárcavas en 14 localidades con pérdida de suelo agrícola.",
           "«La quebrada Limonal destruye los ingresos de la agricultura» (F-DS-06, San Juan de Bigote)."],
          ["E2-I"], adr="exposicion", linea="efecto", padres=["ED2"]),
        n("EI4", "Mayor aporte de sedimentos al río Piura y pérdida de otros servicios del bosque seco (fuentes de agua, hábitat, carbono)",
          ["Sedimentos al río Piura: hipótesis a verificar con el E4 (hidrología): [POR DEFINIR].",
           "Deforestación asociada por la población a la pérdida de manantiales (30 localidades); agua turbia en Palo Blanco durante los huaicos.",
           "Efecto colateral: la regulación hídrica es otra tipología y no es el fin de este proyecto."],
          ["E2-I", "DX-M", "ZZ05"], linea="efecto", padres=["ED2", "ED3"]),
    ],
    "efectos_directos": [
        n("ED1", "Mayor ocurrencia de huaicos y movimientos en masa (caídas de bloques, derrumbes, deslizamientos) en quebradas, taludes y frentes rocosos",
          ["Peligro por movimientos en masa alto o muy alto en el 90.7 % de los bloques (EVAR); peligro integrado muy alto en los 21 bloques de San Juan de Bigote y Salitral.",
           "Huaicos o flujos de detritos en M6B2-2 (bloques de 1.5–2 m), M2B1 («zona generadora de huaicos confirmada» en el río Huarmaca), M6B10 y M28B4.",
           "Remoción activa en M32B3, 59, M3B6, 17 y M8B2; reptación en 6 bloques; deslizamientos antiguos en 7."],
          ["EVAR", "DX-M", "E2-I"], adr="peligro", linea="efecto"),
        n("ED2", "Erosión hídrica laminar y en cárcavas, con pérdida acelerada de suelo",
          ["39 cárcavas con 18,022.5 m en 18 bloques (34 en el bosque seco del valle); 15 cárcavas activas en M3B7 (10 por ha).",
           "Erosión laminar en el 85 % de los bloques 84 y 85 y en el 80 % del 87; suelo desnudo 68 % (87) y 90 % (77).",
           "Erosión severa en las calicatas de M3B8, M7B1 y 14; 38 de 39 calicatas con horizonte superficial pobre (ócrico)."],
          ["DX-M"], adr="peligro", linea="efecto"),
        n("ED3", "Mayor escorrentía torrencial y transporte de sedimentos que activan las quebradas desde las zonas de origen",
          ["93 % de la lluvia entre enero y abril; marzo de 2017: 1,165 mm en un solo mes (M6B10, M6B2-3).",
           "Infiltración básica lenta a moderadamente lenta (0.33–1.37 cm/h).",
           "Inundación o activación de quebradas es el peligro más percibido en la zona baja (53.3 % de localidades)."],
          ["DX-M", "E2-I"], adr="peligro", linea="efecto"),
    ],
    "problema": n(
        "PC",
        "La población de la provincia de Morropón no accede adecuadamente al servicio ecosistémico de regulación de riesgos naturales "
        "por la degradación del bosque estacionalmente seco de colina y montaña y del mosaico andino en las zonas de origen de huaicos, movimientos en masa e inundaciones",
        ["Peligro integrado alto o muy alto en 5,335.20 ha (72.8 %); 26 bloques en muy alto y 19 en alto; índice provincial 4.020 (Alto).",
         "Bajo el umbral MSAVI 0.4976: 6,216.63 ha (84.8 %) con el criterio areal E6; 6,258.29 ha (40 bloques) si se toma la media de bloque.",
         "Estado del ecosistema (E6): 8 bloques moderados y 6 severos (77, 83, 84, 85, 86, 87); índice de condición 0.706.",
         "Peligro frente a MSAVI r = −0.839; frente a pendiente r = −0.619; frente a litología r = −0.075."],
        ["E2-I", "DX-M", "AVA", "F04A"],
        indicador="Magnitud: peligro integrado alto + muy alto en 5,335.20 ha (72.8 %) · línea base preliminar de brecha: 6,216.63 ha bajo el umbral MSAVI (84.8 %, criterio areal E6)"),
    "factores_exogenos": [
        n("EX1", "Pendiente media 34.7 % (49.6 % en el mosaico andino); susceptibilidad litológica a movimientos en masa alta o muy alta en 42 bloques (Complejo de Olmos, Fm. Río Seco, Fm. La Bocana).",
          ["Susceptibilidad a erosión hídrica alta o muy alta en 45 bloques (3,990 ha).", "La litología casi no explica el peligro integrado (r = −0.075)."],
          ["DX-M", "EVAR"], adr="Peligro · factores condicionantes no modificables", linea="exogeno",
          titulo="Factores condicionantes del peligro"),
        n("EX2", "Lluvias de El Niño 1983 (×3.71), 1998 (×3.75), 2017 (×2.97) y 2023 (×3.36) la media, hasta ×6.1 en M27B1; 15 de 45 años con menos de la mitad de la media; Tmáx +0.19 °C por década.",
          ["543 mm/año e índice de aridez 0.31 (0.26 en el valle); evapotranspiración +14.8 mm por década.",
           "Déficit hídrico futuro en el escenario SSP5-8.5."],
          ["DX-M"], adr="Peligro · factor desencadenante y cambio climático", linea="exogeno",
          titulo="Factor desencadenante y cambio climático"),
    ],
    "causas_directas": [
        n("CD1", "Degradación de la cobertura vegetal protectora en las zonas de origen",
          ["Herbazal seco 34.5 %, matorral seco o caducifolio defoliado 27.3 % y suelo desnudo, roca y vías 13.3 %; dosel denso solo 4.0 % (S2DR4).",
           "Suelo desnudo 1,060 ha (14.5 %), vigor muy bajo 3,475 ha (47.4 %) y vigor bajo 1,681 ha (22.9 %).",
           "M2B1 con tendencia negativa significativa del vigor (2014–2024)."],
          ["DX-M"], adr="condicionante", linea="verde"),
        n("CD2", "Reducida capacidad de los suelos y laderas para controlar la erosión y la escorrentía",
          ["Entisoles dominantes con horizonte superficial pobre; suelos someros en M28B3 (15 cm), M11B3 y M3B5 (30 cm), M18B1 (34 cm).",
           "Materia orgánica, textura, pH y factor K pendientes de laboratorio (UNALM): [POR DEFINIR]."],
          ["DX-M"], adr="condicionante", linea="marron"),
        n("CD3", "Inadecuada gobernanza y gestión local del ecosistema y del riesgo",
          ["Comité de recursos naturales solo en 2 de 56 localidades; ningún espacio de concertación ni sistema formal de alerta temprana.",
           "71.8 % de la superficie sin registro predial."],
          ["E2-I", "DX-M"], adr="resiliencia", linea="gob"),
    ],
    "causas_indirectas": [
        n("CI1.1", "Sobrepastoreo caprino y vacuno sin manejo, que impide la regeneración del bosque seco",
          ["Pastoreo en 41 de 59 informes DT; ganadería caprina u ovina en 83.6 % de las localidades y vacuna en 92.7 %.",
           "Carga de 5 UA/ha en el bloque 77, 3 UA/ha en el 60 % de M9B1 y 2.5 UA/ha en el 52.",
           "Bloques 83–87, 77 y M3B5 con vigor 0.10–0.18 por debajo de lo esperado para su aridez, atribuido al pastoreo caprino."],
          ["DX-M", "E2-I"], adr="fragilidad", linea="marron", padre="CD1"),
        n("CI1.2", "Tala comercial de palo santo y extracción de madera y leña de especies del bosque seco",
          ["Deforestación y tala mencionadas en 30 localidades; tala comercial de palo santo (incluida madera verde) en La Pilca, Río Seco, Tórtola, Huaro Quispampa, Maray e Ingenio (riesgo social RS-01, alto).",
           "Hualtaco y palo santo están En Peligro Crítico; leña en 6 informes DT y ladrilleras en M27B1.",
           "El gas y las cocinas mejoradas ya reducen la presión por leña en 10 localidades."],
          ["E2-I", "DX-M"], adr="fragilidad", linea="verde", padre="CD1"),
        n("CI1.3", "Expansión agrícola y ocupación urbana de laderas y cauces",
          ["Uso agrícola en 17 de 59 informes DT; en el mosaico andino la zona agrícola domina los bloques 11 (80 %), 12 (94.9 %), 15 (82.4 %) y 17 (97.6 %).",
           "Lotización en M7B1; habilitación urbana en M2B1 y M7B6; viviendas y corrales dentro de los bloques 84–87."],
          ["DX-M"], adr="fragilidad", linea="verde", padre="CD1"),
        n("CI1.4", "Baja capacidad de regeneración natural del bosque seco por aridez extrema, sequías recurrentes y quemas",
          ["La aridez explica el 68 % de la variación del MSAVI; índice de aridez 0.26 en el valle.",
           "Sequía: primer problema citado por la población (37 localidades).",
           "Quemas en los bloques 77 (1 cada 5 años), 52 (3 cada 5 años), 82 y M17B1."],
          ["DX-M", "E2-I"], adr="fragilidad", linea="verde", padre="CD1"),
        n("CI2.1", "Senderos de ganado y pisoteo que compactan el suelo y concentran la escorrentía",
          ["Terracetas por pisoteo en M3B8; senderos de ganado que funcionan como microcauces en M3B9, M19B2, M32B3, 58 y M28B3."],
          ["DX-M"], adr="condicionante", linea="marron", padre="CD2"),
        n("CI2.2", "Trochas, vías y canteras con cortes de talud sin drenaje ni estabilización",
          ["En M17B1 la obra vial Chulucanas–Frías es «la causa de degradación de mayor intensidad»; caída de bloques en la trocha PI-799 (11, 17).",
           "Trochas o vías como factor en 7 de 59 informes DT; canteras de áridos en 59 y M6B2-1."],
          ["DX-M"], adr="condicionante", linea="complementaria", padre="CD2"),
        n("CI2.3", "Cárcavas activas y erosión laminar sin prácticas de conservación ni control",
          ["39 cárcavas registradas; 41 de 59 bloques sin registro de campo (subregistro).",
           "Ninguna práctica de conservación de suelos reportada en los bloques con cárcavas."],
          ["DX-M"], adr="condicionante", linea="marron", padre="CD2"),
        n("CI3.1", "Limitado conocimiento y organización comunal para el manejo del ecosistema y la gestión del riesgo",
          ["Comité de recursos naturales en 3.6 % de las localidades; solo avisos telefónicos y monitoreo visual como alerta (RS-12, alto).",
           "Movimientos en masa o derrumbes mencionados solo en 17.8 % de las localidades, aunque el 72.8 % del ámbito tiene peligro alto o muy alto.",
           "Mujeres: 13 % de los entrevistados."],
          ["E2-I", "MAT"], adr="resiliencia", linea="gob", padre="CD3"),
        n("CI3.2", "Limitada gobernanza de las entidades públicas y privadas: sin concertación, débil control forestal y sin mecanismos de retribución",
          ["Ningún espacio de concertación en 56 CP; presencia estatal baja en 26 localidades.",
           "La prohibición comunal de tala (8 localidades) no impide la extracción comercial; las rondas no pueden controlarla.",
           "Ningún MERESE; desconfianza por promesas incumplidas (RS-05)."],
          ["E2-I", "DX-M"], adr="resiliencia", linea="gob", padre="CD3"),
        n("CI3.3", "Tierras sin registro, de dominio estatal o comunal, y conflictos de linderos y de uso de pastos",
          ["71.8 % sin registro (5,253.11 ha); 28.0 % comunal; polígono comunal 2790 sin nombre (10 bloques).",
           "Conflictos de linderos La Pareja–Alto San José, Dótor–Nueva Esperanza y Cardal–Charanas; pago por pastos comunales en Andanjo (250 ha).",
           "Tamizaje predial aún no iniciado."],
          ["E2-I", "DX-M"], adr="fragilidad", linea="gob_predial", padre="CD3"),
    ],
    "notas": "Árbol provincial de Morropón. Superficie V8 neta 7,313.55 ha (se excluyeron 13.81 ha por sitio arqueológico y derechos mineros). "
             "Pendientes: niveles PMM/EPH/PGI por bloque, inventario de daños por evento, delimitación de la población aguas abajo, "
             "laboratorio de suelos, pobreza y NBI, y F-DS-04, 05 y 07.",
}

# ================================================================== UNIFICADO
P = lambda m, h, a: {"M": m, "H": h, "A": a}
unificado = {
    "pestana": "Unificado",
    "sigla": "117 bloques",
    "prov_leyenda": True,
    "titulo": "Árbol de problemas unificado – Cuenca alta del río Piura",
    "subtitulo": "Morropón, Huancabamba y Ayabaca · 15 distritos · 117 bloques · 12,270.24 ha (12,157.47 ha netas V8) · 154 centros poblados (18,651 hab.)",
    "fuente": "Fuentes: Informes E2 Vol. I–III (28 y 30-oct-2026); Matrices E2 B2 (24-oct-2026); Informe EVAR (19-oct-2026); Avance del Perfil (28-oct-2026); modelo ANIN ZZ-000005",
    "ficha": [
        {"k": "Bloques", "v": "117 en 15 distritos"},
        {"k": "Superficie", "v": "12,270.24 ha (V8 neta 12,157.47 ha)"},
        {"k": "Centros poblados", "v": "154 (18,651 hab.)"},
        {"k": "PI alto + muy alto", "v": "8,471.57 ha (69.0 %)"},
        {"k": "PMM alto + muy alto (EVAR)", "v": "93.9 % de los bloques"},
        {"k": "Bajo umbral MSAVI", "v": "8,341.83 ha (68.0 %)"},
    ],
    "lectura": (
        "El árbol unificado conserva las causas y efectos que se repiten con evidencia en las tres provincias y marca con M, H y A "
        "dónde aparecen y con qué intensidad. La estructura sigue el modelo de referencia ANIN: las dos primeras causas directas "
        "desagregan la «alteración de los factores condicionantes en zonas de origen» (cobertura y suelo) y la tercera es la gobernanza; "
        "cada una corresponde a un componente del proyecto (infraestructura verde, marrón y gobernanza)."
    ),
    "efecto_final": n(
        "EF",
        "Incremento de los costos de reconstrucción y disminución de la resiliencia de la población de la cuenca del río Piura y sus medios de vida "
        "frente a movimientos en masa e inundaciones desencadenados por eventos climáticos extremos, intensificados por el cambio climático",
        ["18,651 hab. en 154 centros poblados asociados (Censo 2017; incluye 705 hab. en 3 CP sin código INEI); 3,285 personas vulnerables; 6,447 viviendas, 89 % de adobe (EVAR).",
         "Riesgo por movimientos en masa en las áreas de influencia: alto en 7,132.39 ha (41.6 %) y muy alto en 56.92 ha (Chococa y Cambruran, Huarmaca).",
         "Redacción alineada al efecto final del modelo de referencia ANIN (ZZ-000005, Fig. 2).",
         "Pérdidas económicas no valoradas en ninguna fuente: [POR DEFINIR]."],
        ["MAT", "EVAR", "ZZ05"], linea="efecto", prov=P(3, 3, 3)),
    "efectos_indirectos": [
        n("EI1", "Interrupción de vías y aislamiento estacional de caseríos, con pérdida de cosechas que no llegan al mercado",
          ["Huancabamba: interrupción de la carretera en 11 localidades; Shuturumbe aislado 8 días (marzo 2024); café sin vender.",
           "Morropón: 4 caseríos aislados de enero a abril; carreteras tapadas por huaicos.",
           "Ayabaca: taludes activos del corredor Chulucanas–Frías y cortes de energía en lluvias."],
          ["E2-I", "E2-II", "DX-A"], adr="exposicion", linea="efecto", prov=P(2, 3, 2), padres=["ED1", "ED3"]),
        n("EI2", "Daños a viviendas, colegios, parcelas, ganado, captaciones y líneas de conducción de agua",
          ["Morropón: flujo sobre Dótor y su colegio; muros derruidos en Serrán; captación a 50 m de M32B3; bolones sobre viviendas.",
           "Huancabamba: daños a chacras, ganado y viviendas en 11 localidades; exposición interna en 3 bloques y de interfaz en 12.",
           "Ayabaca: tubería de conducción en ladera (bloque 6); antecedente de reubicación de un CP (bloque 56)."],
          ["E2-I", "E2-II", "DX-A"], adr="exposicion", linea="efecto", prov=P(3, 3, 2), padres=["ED1"]),
        n("EI3", "Pérdida de suelo agrícola y de ingresos de la agricultura y la ganadería familiar",
          ["Morropón: erosión y cárcavas en 14 localidades; Huancabamba: terrenos arrastrados en 5 y cárcavas en parcelas en 6.",
           "Ayabaca: pérdida del horizonte superficial en suelos de ladera cultivada (bloques 36 y 39)."],
          ["E2-I", "E2-II", "DX-A"], adr="exposicion", linea="efecto", prov=P(2, 3, 2), padres=["ED2"]),
        n("EI4", "Mayor aporte de sedimentos que colmata quebradas y cauces hacia el valle del Alto Piura (zona de impacto)",
          ["Hipótesis presente en los tres volúmenes; debe cuantificarse con el estudio hidrológico (E4): [POR DEFINIR].",
           "El EVAR advierte que el peligro alto en cabeceras justifica intervenir aunque el riesgo aguas abajo aún no esté medido."],
          ["DX-M", "DX-H", "DX-A", "EVAR"], adr="exposicion", linea="efecto", prov=P(2, 1, 1), padres=["ED3"]),
        n("EI5", "Pérdida de otros servicios del ecosistema: regulación hídrica de fuentes, hábitat de especies amenazadas y captura de carbono",
          ["Sequía o disminución de fuentes: problema más citado (82–89 % de localidades); deforestación asociada a la pérdida de manantiales.",
           "Hualtaco y palo santo (En Peligro Crítico) en Morropón y Ayabaca.",
           "Se registra como efecto colateral, como en el modelo ANIN: la regulación hídrica tiene su propia tipología y brecha."],
          ["MAT", "ZZ05"], linea="efecto", prov=P(3, 3, 2), padres=["ED2", "ED3"]),
    ],
    "efectos_directos": [
        n("ED1", "Mayor ocurrencia de movimientos en masa (deslizamientos, derrumbes, caídas de rocas, reptación) y huaicos en laderas, taludes y quebradas de cabecera",
          ["Peligro por movimientos en masa alto o muy alto en el 93.9 % de los bloques (5,977.80 ha muy alto) y en el 96.9 % de las áreas de aporte (EVAR).",
           "Remociones activas en 16 bloques de Huancabamba y 6 de Morropón; taludes viales activos en Ayabaca.",
           "Morropón: huaicos confirmados en M2B1, M6B2-2, M6B10 y M28B4."],
          ["EVAR", "DX-M", "DX-H", "DX-A"], adr="peligro", linea="efecto", prov=P(2, 3, 3)),
        n("ED2", "Erosión hídrica acelerada y pérdida de suelo (laminar, en surcos y en cárcavas)",
          ["55 cárcavas registradas con 23.08 km (Morropón 39 y 18.02 km; Huancabamba 9 y 3.90 km; Ayabaca 7 y 1.15 km); 85 bloques sin registro de campo.",
           "Erosión severa en calicatas de bloques degradados de las tres provincias."],
          ["MAT", "DX-M", "DX-H", "DX-A"], adr="peligro", linea="efecto", prov=P(3, 2, 2)),
        n("ED3", "Mayor escorrentía torrencial y transporte de sedimentos desde las zonas de origen, que activan las quebradas",
          ["82–93 % de la lluvia entre enero y abril; en años El Niño la lluvia llega a 2–6 veces la media.",
           "Activación de quebradas o inundación citada en 53.3 % de las localidades de Morropón y 27.8 % de Huancabamba.",
           "Infiltración básica lenta en suelos degradados (0.33–2.2 cm/h)."],
          ["MAT", "DX-M", "DX-H"], adr="peligro", linea="efecto", prov=P(3, 2, 2)),
    ],
    "problema": n(
        "PC",
        "La población de la cuenca alta del río Piura no accede adecuadamente al servicio ecosistémico de regulación de riesgos naturales "
        "por la degradación de los ecosistemas en las zonas de origen de movimientos en masa e inundaciones",
        ["Peligro integrado alto o muy alto en 8,471.57 ha (69.0 % de los 117 bloques): Morropón 5,335.20 ha, Huancabamba 2,353.73 ha, Ayabaca 782.64 ha.",
         "Línea base preliminar de brecha: 8,341.83 ha bajo el umbral MSAVI 0.4976 (68.0 %, criterio areal E6): restauración activa 5,503.42 ha y enriquecimiento 2,838.40 ha.",
         "En las tres provincias el peligro sigue a la pérdida de cobertura (r con MSAVI: −0.839, −0.785 y −0.968) y no a la litología.",
         "Servicio definido en el Formato 04-A (RM N° 00213-2024-MINAM): «capacidad de los ecosistemas de reducir las condiciones de vulnerabilidad … (huaycos, avalanchas, entre otros)»."],
        ["AVA", "MAT", "F04A", "ZZ05"],
        indicador="Indicador de brecha (RM N° 00213-2024-MINAM): % de superficie de ecosistemas degradados que requieren recuperación · línea base preliminar 8,341.83 ha (68.0 %)",
        prov=P(3, 3, 3)),
    "factores_exogenos": [
        n("EX1", "Pendientes fuertes (media 41 %; 822.9 ha sobre 75 %), relieve montañoso y vertientes coluviales (88.1 % de los bloques en geomorfología de nivel 4–5) y litologías meteorizadas o foliadas.",
          ["Ayabaca es la más empinada (64.6 % del área sobre 50 %), luego Huancabamba (45.9 %) y Morropón (16.7 %).",
           "Pesos del modelo de peligro por movimientos en masa: pendiente 0.633, geomorfología 0.260, geología 0.106 (EVAR)."],
          ["EVAR", "MAT"], adr="Peligro · factores condicionantes no modificables", linea="exogeno",
          titulo="Factores condicionantes del peligro"),
        n("EX2", "Lluvias de El Niño 1983, 1998, 2017 y 2023 de 2 a 6 veces la media; régimen semiárido (544 mm/año, aridez 0.34) con 82–93 % de la lluvia entre enero y abril; calentamiento de +0.15 a +0.25 °C por década.",
          ["Déficit hídrico futuro en el escenario SSP5-8.5 (MPI-ESM1-2HR, 2021–2040), criterio de idoneidad del proyecto.",
           "El EVAR usa un desencadenante fijo (lluvia 60–80 % sobre la normal) y no modela escenarios de cambio climático: [POR DEFINIR]."],
          ["MAT", "EVAR"], adr="Peligro · factor desencadenante y cambio climático", linea="exogeno",
          titulo="Factor desencadenante y cambio climático"),
    ],
    "causas_directas": [
        n("CD1", "Degradación de la cobertura vegetal protectora en las zonas de origen",
          ["MSAVI medio ponderado: Morropón 0.379, Huancabamba 0.508, Ayabaca 0.548 (umbral 0.4976).",
           "20 bloques en estado moderado o severo (E6): 14 en Morropón y 6 en Huancabamba.",
           "Corresponde a la CI 1.1 del modelo ANIN («Degradación de la cobertura vegetal en zonas que proveen el SERRN»)."],
          ["MAT", "ZZ05"], adr="condicionante", linea="verde", prov=P(3, 2, 2)),
        n("CD2", "Reducida capacidad de los suelos y laderas para controlar la erosión y la escorrentía",
          ["Suelos someros (15–45 cm en bloques degradados), Entisoles con horizonte superficial pobre e infiltración lenta.",
           "Corresponde a la CI 1.2 del modelo ANIN («Reducida capacidad de control de erosión y de estabilidad de suelos»)."],
          ["DX-M", "DX-H", "DX-A", "ZZ05"], adr="condicionante", linea="marron", prov=P(3, 3, 2)),
        n("CD3", "Inadecuada gobernanza y gestión local del ecosistema y del riesgo",
          ["Comité de recursos naturales en 5 de 85 localidades; ninguna con plataforma de concertación ni alerta temprana formal.",
           "Corresponde a la CD 2 del modelo ANIN («Inadecuada gobernanza y gestión del ecosistema»)."],
          ["MAT", "ZZ05"], adr="resiliencia", linea="gob", prov=P(3, 3, 3)),
    ],
    "causas_indirectas": [
        n("CI1.1", "Sobrepastoreo caprino y vacuno sin manejo, que impide la regeneración natural",
          ["Morropón: pastoreo en 41 de 59 informes DT; caprinos en 83.6 % de las localidades; hasta 5 UA/ha.",
           "Huancabamba: pastoreo en 45 de 48 informes DT; intensidad fuerte en 8 bloques.",
           "Ayabaca: causa principal en M17B5; vacunos en 100 % de las localidades."],
          ["DX-M", "DX-H", "DX-A", "MAT"], adr="fragilidad", linea="marron", prov=P(3, 3, 2), padre="CD1"),
        n("CI1.2", "Tala y extracción selectiva de madera y leña de especies nativas, incluidas especies en peligro crítico",
          ["Morropón: tala comercial de palo santo (RS-01, alto); deforestación citada en 30 localidades.",
           "Huancabamba: tala para leña o madera en 25 bloques; venta de roble, laurel, nogal y cedro.",
           "Ayabaca: tala en M17B7 y en La Cría."],
          ["E2-I", "E2-II", "E2-III"], adr="fragilidad", linea="verde", prov=P(3, 2, 2), padre="CD1"),
        n("CI1.3", "Expansión de la agricultura y de la ocupación hacia laderas de fuerte pendiente y cauces",
          ["Ayabaca: causa principal en 7 de 10 informes DT; 68 % del ámbito es zona agrícola.",
           "Huancabamba: agricultura en ladera en 38 de 48 informes DT (café, maíz, pastos).",
           "Morropón: uso agrícola en 17 de 59 informes DT y lotización u ocupación urbana en 4 bloques."],
          ["DX-A", "DX-H", "DX-M"], adr="fragilidad", linea="verde", prov=P(2, 3, 3), padre="CD1"),
        n("CI1.4", "Baja capacidad de regeneración natural por aridez, sequías recurrentes, quemas y especies exóticas",
          ["Índice de aridez 0.31–0.39 (0.24–0.26 en los bosques secos de valle); la aridez explica el 68 % de la variación del MSAVI en Morropón.",
           "Quemas puntuales en 4 bloques de Morropón y 6 de Huancabamba; especies exóticas en 16 bloques de Huancabamba."],
          ["DX-M", "DX-H", "DX-A"], adr="fragilidad", linea="verde", prov=P(3, 2, 2), padre="CD1"),
        n("CI2.1", "Trochas, vías y canteras con cortes de talud sin drenaje ni estabilización",
          ["Huancabamba: vías y taludes de corte en 40 de 48 informes DT.",
           "Ayabaca: corredor Chulucanas–Frías, causa principal en M17B6 y M17B10.",
           "Morropón: en M17B1 la obra vial es la causa de mayor intensidad; trocha PI-799."],
          ["DX-H", "DX-A", "DX-M"], adr="condicionante", linea="complementaria", prov=P(2, 3, 3), padre="CD2"),
        n("CI2.2", "Laderas cultivadas sin prácticas de conservación de suelos y cárcavas activas sin control",
          ["55 cárcavas y 23.08 km registrados; labranza a favor de la pendiente (Ayabaca, bloque 27).",
           "15 cárcavas activas en M3B7 (Morropón); cárcavas activas severas en el bloque 40 (Huancabamba)."],
          ["MAT", "DX-M", "DX-H", "DX-A"], adr="condicionante", linea="marron", prov=P(3, 2, 2), padre="CD2"),
        n("CI2.3", "Pisoteo del ganado y senderos que compactan suelos someros y concentran la escorrentía",
          ["Morropón: senderos de ganado como microcauces en 5 bloques; terracetas por pisoteo.",
           "Huancabamba: infiltración básica de 0.41 a 2.2 cm/h; Ayabaca: compactación por ganado en M17B7."],
          ["DX-M", "DX-H", "DX-A"], adr="condicionante", linea="marron", prov=P(3, 2, 1), padre="CD2"),
        n("CI3.1", "Limitado conocimiento y organización de las comunidades locales para el manejo del ecosistema, la adaptación al cambio climático y la gestión del riesgo",
          ["Comité de recursos naturales: Morropón 2 de 56, Huancabamba 3 de 22, Ayabaca 0 de 7 localidades.",
           "Ningún sistema formal de alerta temprana; prohibiciones comunales de tala sin control efectivo.",
           "Corresponde a la CI 2.2 del modelo ANIN."],
          ["MAT", "ZZ05"], adr="resiliencia", linea="gob", prov=P(3, 3, 3), padre="CD3"),
        n("CI3.2", "Limitada gobernanza de las entidades públicas y privadas: sin concertación, con escasa presencia institucional y desconfianza por intervenciones incumplidas",
          ["Ninguna plataforma de concertación en las tres provincias; presencia estatal «baja» en 27 de 56, 10 de 22 y 3 de 7 localidades.",
           "Obras inconclusas (Huancabamba), promesas incumplidas (Morropón) y estafa previa (Ayabaca); ningún MERESE.",
           "Corresponde a la CI 2.1 del modelo ANIN."],
          ["MAT", "ZZ05"], adr="resiliencia", linea="gob", prov=P(3, 3, 2), padre="CD3"),
        n("CI3.3", "Tenencia de la tierra sin sanear: tierras sin registro, posesión informal y conflictos de linderos",
          ["Sin registro: Morropón 71.8 %, Huancabamba 71.0 %; Ayabaca 83.9 % comunal con posesionarios sin título.",
           "Conflictos de linderos en Morropón (3) y Ayabaca (La Cría); 5 bloques de Huancabamba superan los 19 predios."],
          ["MAT"], adr="fragilidad", linea="gob_predial", prov=P(3, 3, 2), padre="CD3"),
    ],
    "notas": "Árbol unificado. Una causa o efecto entra al árbol unificado cuando tiene evidencia en al menos dos provincias, o cuando es "
             "generalizado en una de ellas. Las marcas M, H y A indican presencia (· puntual, ·· frecuente, ··· generalizada). "
             "Las causas propias de una sola provincia (especies exóticas y derechos mineros en Huancabamba; corredor vial en Ayabaca; "
             "ocupación urbana en Morropón) se detallan en los árboles provinciales.",
}

ARBOLES = {"unificado": unificado, "morropon": morropon, "huancabamba": huancabamba, "ayabaca": ayabaca}
ORDEN = ["unificado", "morropon", "huancabamba", "ayabaca"]


def _reordenar_efectos(arbol, orden):
    """Ordena los efectos indirectos según el efecto directo que los origina (evita cruces) y renumera EI1…EIn."""
    por_id = {e["id"]: e for e in arbol["efectos_indirectos"]}
    nuevos = [por_id[i] for i in orden]
    for k, e in enumerate(nuevos, 1):
        e["id"] = f"EI{k}"
    arbol["efectos_indirectos"] = nuevos


_reordenar_efectos(morropon, ["EI2", "EI3", "EI1", "EI4"])
_reordenar_efectos(ayabaca, ["EI1", "EI3", "EI2", "EI4"])
