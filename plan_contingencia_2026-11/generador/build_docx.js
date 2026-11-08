// Plan de Contingencia del Estudio de Perfil · Word con formato ANIN (docx-js + anin_docx.js)
// Uso: NODE_PATH=$(npm root -g) node build_docx.js plan_contingencia.json salida.docx
const fs = require("fs"), path = require("path");
const A = require(path.join(__dirname, "..", "..", ".claude", "skills", "anin-in-piura", "scripts", "anin_docx"));
const { Document, Packer, Paragraph, TextRun, AlignmentType, PageBreak } = require("docx");
const D = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const OUT = process.argv[3];
const WL = 14570; // ancho útil A4 horizontal con márgenes de 2 cm (DXA)
const f0 = (n) => n.toLocaleString("en-US", { maximumFractionDigits: 0 });
const f2 = (n) => n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const cols = (ws, total) => { const s = ws.reduce((a, b) => a + b, 0); return ws.map((w, i) => ({ w: Math.round(w / s * total), align: A.L })); };
const sp = () => new Paragraph({ children: [], spacing: { after: 60 } });
const salto = () => new Paragraph({ children: [new PageBreak()] });
const nivel = (p, i) => { const s = p * Math.abs(i); return s >= 20 ? "Muy alto" : s >= 12 ? "Alto" : s >= 6 ? "Medio" : "Bajo"; };
const ent = Object.fromEntries(D.entregables.map((e) => [e[0], e]));
const avance = D.avance.toFixed(1);

// ------------------------------------------------------------------ Portada y datos
const portada = [
  new Paragraph({ spacing: { before: 2400, after: 200 }, alignment: AlignmentType.CENTER, children: [A.run(D.meta.titulo, { bold: true, size: 32, color: A.VERDE })] }),
  new Paragraph({ spacing: { after: 400 }, alignment: AlignmentType.CENTER, children: [A.run(D.meta.subtitulo, { size: 23, color: A.GRIS })] }),
  A.meta([
    ["Proyecto", D.meta.proyecto], ["CUI", D.meta.cui], ["Entidad", "Autoridad Nacional de Infraestructura (ANIN) · Dirección de Intervenciones Multisectoriales y de Emergencia (DIME) · Subdirección de Estudios de Inversión (SESDI)"],
    ["Fase", D.meta.fase], ["Modalidad", D.meta.modalidad],
    ["Presupuesto total de la inversión", `S/ ${f0(D.meta.techo)} (≈ USD ${(D.meta.techo / D.meta.tc / 1e6).toFixed(1)} M a ${D.meta.tc} S/ por USD)`],
    ["Ámbito", D.meta.ambito], ["Sistema de coordenadas", D.meta.crs],
    ["Contexto", `${D.decreto.norma} (${D.decreto.publicacion}) · Comunicado Oficial ENFEN N° 19-2026 (30/10/2026)`],
    ["Versión y fecha", `${D.meta.version} · ${D.meta.fecha}`], ["Elaborado por", D.meta.elaborado],
  ]),
  sp(),
  A.fuente("Documento de trabajo del equipo formulador. Las cifras del Perfil provienen de la presentación resumida ante la SESDI del 03/11/2026 (Diagnóstico Integral IN Piura V3). Lo que no tiene fuente verificada se marca [POR DEFINIR], [POR VERIFICAR] o [SUPUESTO]."),
  A.h2("Contenido"),
  ...["1. Antecedentes y contexto", "2. Objetivo, alcance y principios del plan", "3. Calendario de actividades hasta la culminación del Perfil", "4. Presupuesto y recursos necesarios", "5. Cambios de alcance y contrataciones estrictamente necesarias", "6. Gestión de riesgos del plan", "7. Actividades pendientes por entregable", "8. Aspectos administrativos y de gestión", "9. Seguimiento y actualización del plan", "10. Conclusiones y recomendaciones", "Fuentes", "Anexo A. Distritos del ámbito y verificación del Anexo del decreto", "Anexo B. Archivos complementarios"].map((t) => A.bullet(t)),
];

// ------------------------------------------------------------------ 1. Antecedentes
const s1 = [
  A.h1("1. ANTECEDENTES Y CONTEXTO", { nuevaPagina: true }),
  A.h2("1.1 El Decreto Supremo N° 149-2026-PCM"),
  A.para(`El ${D.decreto.norma}, publicado en el ${D.decreto.publicacion}, ${D.decreto.objeto.charAt(0).toLowerCase()}${D.decreto.objeto.slice(1)}`),
  A.para([A.b("Plazo. "), D.decreto.plazo]),
  A.para([A.b("Alcance nacional. "), D.decreto.alcance_nacional]),
  A.para([A.b("Alcance en Piura. "), D.decreto.alcance_piura]),
  A.para([A.b("Acciones. "), D.decreto.acciones]),
  A.para([A.b("Financiamiento. "), D.decreto.financiamiento]),
  A.para([A.b("Sustento técnico. "), D.decreto.sustento]),
  A.para([A.b("Nota sobre la fuente. "), D.decreto.nota_acceso]),
  A.h2("1.2 Escenario climático: El Niño 2026-2027"),
  A.para(D.clima.comunicado + " " + D.clima.magnitud),
  A.para(D.clima.lluvias),
  A.para([A.b("Implicancia para el ámbito del proyecto. "), D.clima.implicancia]),
  A.h2("1.3 Implicancia para el Proyecto IN Piura"),
  A.para("El Estado de Emergencia y el escenario climático inciden sobre el Perfil en cuatro frentes:"),
  A.bullet([A.b("Trabajo de campo. "), "El tamizaje predial (ruta F-LA), las fichas sociales pendientes, las pruebas de infiltración, el inventario de cárcavas y la validación del diagnóstico de ecosistemas requieren acceso a 117 bloques situados en laderas con peligro alto y muy alto por movimientos en masa; la temporada de lluvias de diciembre a abril restringe ese acceso y eleva el riesgo para las brigadas."]),
  A.bullet([A.b("Autoridades e instituciones. "), "Los gobiernos regionales y locales asumen las acciones de reducción del riesgo, respuesta y rehabilitación durante la emergencia, y el 01/01/2027 cambian las autoridades regionales, provinciales y distritales; además se renuevan juntas directivas comunales. Las asambleas, actas y acuerdos del proyecto compiten con la agenda de la emergencia y deben ratificarse con las nuevas gestiones."]),
  A.bullet([A.b("Capacidad institucional de la ANIN. "), "La DIME es la dirección de intervenciones de emergencia de la ANIN; la declaratoria puede reasignar personal y prioridades, y el cierre del año fiscal 2026 condiciona la continuidad de contratos y servicios en enero de 2027."]),
  A.bullet([A.b("Oportunidad técnica. "), "El evento genera evidencia directa del peligro en la zona de origen (precipitación, activación de cárcavas, movimientos en masa, daños), útil para la hidrología (E4), el EVAR (E8), el problema (E7) y los beneficios por daños evitados (E13)."]),
  A.h2("1.4 Estado del Estudio de Perfil al 03/11/2026"),
  A.para(`El Perfil alcanza un avance ponderado de ${avance} % sobre los 15 entregables (45 % el 28/10/2026). El Capítulo de Identificación (E1 a E9, peso 64) está al 71.9 %, con Geología (E3), Suelos (E5) y Análisis de riesgos (E8) culminados; Formulación (E10 a E12, peso 18) al 13.9 % y Evaluación (E13 y E14, peso 13) sin iniciar. La programación original prevé la Identificación en diciembre de 2026, Formulación y Evaluación en febrero de 2027 y el Perfil completo en marzo de 2027, con viabilidad en abril de 2027.`),
  A.cap("Tabla 1. Avance por entregable al 03/11/2026"),
  A.tabla(cols([8, 40, 7, 9, 12, 60], A.W), ["Entregable", "Nombre", "Peso", "Avance", "Estado", "Sustento al corte"],
    D.entregables.map((e) => [e[0], e[2], String(e[3]), `${e[4]} %`, e[5], e[6]]),
    ["Total", "15 entregables", "100", `${avance} %`, "", "Avance ponderado = suma de peso × avance / 100"]),
  A.fuente("Fuente: Presentación resumida de avance del Perfil ante la SESDI (03/11/2026); Diagnóstico Integral IN Piura V3."),
  A.cap("Tabla 2. Programación original del Perfil (03/11/2026)"),
  A.tabla(cols([14, 14, 72], A.W), ["Hito", "Avance previsto", "Contenido"], D.programacion_original.map((p) => [p[0], p[1], p[2]])),
  A.fuente("Fuente: Presentación resumida ante la SESDI (03/11/2026), lámina «Proyección del avance del Perfil». Fechas referenciales, sujetas al campo y a la temporada de lluvias."),
];

// ------------------------------------------------------------------ 2. Objetivo
const s2 = [
  A.h1("2. OBJETIVO, ALCANCE Y PRINCIPIOS DEL PLAN", { nuevaPagina: true }),
  A.para([A.b("Objetivo. "), "Asegurar la culminación del Estudio de Perfil del Proyecto IN Piura (CUI 2669244) con la calidad exigida por la Guía General de la DGPMI-MEF y su Anexo 2 (GdR-CCC), en el contexto del Estado de Emergencia declarado por el D.S. N° 149-2026-PCM y de la temporada de lluvias del Fenómeno El Niño 2026-2027, protegiendo la seguridad del personal y la continuidad institucional."]),
  A.para([A.b("Alcance. "), "Comprende las actividades pendientes de los 15 entregables del Perfil, el trabajo de campo en los 117 bloques y 60 lotes SUS del catálogo V8, las coordinaciones con autoridades regionales, locales y comunales, los recursos y contrataciones estrictamente necesarios y la gestión de los riesgos que podrían impedir o dificultar la culminación del estudio. El horizonte del plan va del 04/11/2026 a la declaración de viabilidad (junio de 2027 en el escenario base)."]),
  A.para([A.b("Principios. ")]),
  A.bullet([A.b("Seguridad primero. "), "Ninguna salida de campo sin autorización del protocolo de seguridad; umbrales de suspensión por avisos del SENAMHI y reportes del COER Piura."]),
  A.bullet([A.b("Campo antes de las lluvias, gabinete durante las lluvias. "), "Noviembre de 2026 concentra el campo prioritario; de diciembre a marzo se avanza en gabinete (Identificación, Formulación y Evaluación); abril y mayo de 2027 cierran el campo residual."]),
  A.bullet([A.b("Priorización por riesgo y superficie. "), "Bloques piloto (21, 31, 66, 73, M6B10), bloques con riesgo muy alto (AI_73 Chococa y AI_41 Cambruran, Huarmaca), bloques comunales y los de mayor superficie se atienden primero."]),
  A.bullet([A.b("Trazabilidad. "), "Todo acuerdo con autoridades se documenta antes del 31/12/2026 y se ratifica con las gestiones electas; las decisiones técnicas se registran en actas."]),
  A.bullet([A.b("Sin datos inventados. "), "Lo que el campo no permita verificar se declara como vacío con su plan de cierre; las metas se fijan sobre superficie con libre disponibilidad verificada."]),
  A.bullet([A.b("El evento como evidencia. "), "Se registran los eventos del FEN 2026-2027 en el ámbito (fecha, ubicación UTM 17S, tipo, daños, fuente) para fortalecer E4, E7, E8 y E13."]),
];

// ------------------------------------------------------------------ 3. Cronograma (horizontal)
const s3 = [
  A.h1("3. CALENDARIO DE ACTIVIDADES HASTA LA CULMINACIÓN DEL PERFIL"),
  A.para("El cronograma del escenario base reorganiza las actividades pendientes en seis fases. Las fechas son referenciales y se actualizan al cierre de cada fase según el clima, la disponibilidad de autoridades y la capacidad del equipo. El diagrama semanal con fórmulas está en la hoja 2 del libro Excel adjunto."),
  A.cap("Tabla 3. Fases del plan"),
  A.tabla(cols([62, 38], WL), ["Fase", "Periodo"], Object.values(D.fases).map((f) => [f[0], f[1]])),
  A.cap("Tabla 4. Actividades del cronograma (escenario base)"),
  A.tabla(cols([5, 9, 40, 9, 9, 11, 7, 17, 17], WL), ["ID", "Fase", "Actividad", "Inicio", "Fin", "Entregables", "Campo", "Responsable", "Hito"],
    D.cronograma.map((a) => [a[0], a[1], a[2], a[3].split("-").reverse().join("/"), a[4].split("-").reverse().join("/"), a[5], a[6], a[7], a[8]])),
  A.fuente("Elaboración propia a partir de los pendientes del Diagnóstico Integral V3 («Vacíos para cerrar el Perfil») y la proyección presentada a la SESDI el 03/11/2026. Campo: Sí = depende del acceso a los bloques; Parcial = combina campo y gabinete; No = gabinete."),
  A.cap("Tabla 5. Escenarios de culminación"),
  A.tabla(cols([11, 32, 27, 10, 10, 10], WL), ["Escenario", "Supuestos", "Estrategia", "Perfil completo", "Viabilidad", "Variación"], D.escenarios.map((e) => [...e])),
  A.para("En el escenario base el Perfil se culmina el 31/05/2027 y la viabilidad se obtiene en junio de 2027, dos meses después de la programación original. La versión para revisión del Capítulo de Identificación mantiene la fecha de diciembre de 2026, con los vacíos de campo explícitos y su plan de cierre; su versión final se consolida en mayo de 2027 con el campo residual."),
];

// ------------------------------------------------------------------ 4. Presupuesto (horizontal)
const comps = [...new Set(D.presupuesto.map((p) => p[1]))];
const valorizado = D.presupuesto.reduce((s, p) => s + (p[6] ? p[4] * p[6] : 0), 0);
const s4 = [
  A.h1("4. PRESUPUESTO Y RECURSOS NECESARIOS"),
  A.para("El presupuesto del plan se estructura en siete componentes (A a G) con las partidas y cantidades que se derivan de las actividades pendientes. Las cantidades son supuestos de planificación explícitos y los costos unitarios están [POR DEFINIR] por la Administración de la DIME, con excepción del viático diario, para el que se consigna el valor normativo de referencia del D.S. N° 007-2013-EF (S/ 320 por día), sujeto a la escala vigente de la ANIN. No se han inventado precios: el libro Excel calcula los parciales y totales con fórmulas a medida que se completen las celdas de costo unitario."),
  A.para(`Referencia del techo: la Matriz de Costos Paramétricos v1.6 (03/11/2026) suma S/ 356.63 M de inversión y operación y mantenimiento frente al presupuesto total de la inversión de S/ ${f0(D.meta.techo)}, lo que deja S/ ${f0(D.margen_estudios)} para los estudios (preinversión y expediente técnico). El costo total del Perfil, incluido este plan, debe caber en ese margen junto con el expediente técnico. El presupuesto asignado al Perfil en el POI/PIA de la ANIN está [POR DEFINIR] en este documento.`),
  A.cap("Tabla 6. Resumen por componente"),
  A.tabla(cols([55, 15, 30], WL), ["Componente", "Partidas", "Estado de la valorización"], comps.map((c) => {
    const ps = D.presupuesto.filter((p) => p[1] === c); const sin = ps.filter((p) => p[6] == null).length;
    return [c, String(ps.length), sin === ps.length ? "Costos unitarios [POR DEFINIR]" : `${ps.length - sin} partida(s) con valor referencial; ${sin} [POR DEFINIR]`]; }),
    ["Total", String(D.presupuesto.length), `Valorizado parcial (solo viáticos referenciales): S/ ${f2(valorizado)} + reserva de contingencia ${Math.round(D.contingencia_pct * 100)} % [SUPUESTO]`]),
  A.cap("Tabla 7. Partidas y recursos (cantidades [SUPUESTO]; costos unitarios [POR DEFINIR] salvo referencia)"),
  A.tabla(cols([4, 14, 30, 7, 7, 22, 9, 15, 5], WL), ["Cód.", "Componente", "Partida / recurso", "Unidad", "Cantidad", "Base del cálculo", "CU (S/)", "Fuente del CU", "Fase"],
    D.presupuesto.map((p) => [p[0], p[1].split("·")[1].trim(), p[2], p[3], f0(p[4]), p[5], p[6] != null ? f2(p[6]) : "[POR DEFINIR]", p[7], p[8]])),
  A.fuente("Fuente: elaboración propia. Viático: D.S. N° 007-2013-EF. Las cantidades de brigadas, vehículos y días son supuestos para dimensionar el plan y se ajustan con la disponibilidad real de la DIME."),
  A.h2("4.1 Recursos humanos"),
  A.para("El plan requiere mantener el equipo núcleo del Perfil (coordinación, SIG, social, riesgos, ecosistemas y legal-predial) durante todo 2027 y sumar, por locación de servicios, un especialista en evaluación social (E13) de enero a abril de 2027 y un especialista en gestión de la implementación y operación y mantenimiento (E11 y E12) de noviembre de 2026 a febrero de 2027. Las brigadas de campo se dimensionan en tres equipos de dos integrantes para noviembre de 2026 y para abril y mayo de 2027 [SUPUESTO]. La planilla y los contratos vigentes están [POR VERIFICAR] con la Administración de la DIME."),
  A.h2("4.2 Recursos logísticos, de seguridad y tecnológicos"),
  A.para("Camionetas 4x4 con conductor, combustible, equipos de protección personal para lluvias, Seguro Complementario de Trabajo de Riesgo, comunicación satelital o radio, botiquines y prevención sanitaria; servicio de base de datos y aplicativo IN Piura (PostgreSQL, KoboToolbox) con copias de seguridad diarias; imágenes de alta resolución post-evento y vuelo de dron para cárcavas, drenaje y validación remota del E6."),
];

// ------------------------------------------------------------------ 5. Cambios de alcance
const s5 = [
  A.h1("5. CAMBIOS DE ALCANCE Y CONTRATACIONES ESTRICTAMENTE NECESARIAS"),
  A.para("Los cambios que se proponen no modifican el objetivo ni la tipología del proyecto; ajustan el método o la secuencia para obtener la información con el menor riesgo posible y mantener el Perfil dentro del marco de la Guía DGPMI. Cada cambio identifica la instancia que debe decidirlo."),
  A.cap("Tabla 8. Cambios de alcance, de método y contrataciones estrictamente necesarias"),
  A.tabla(cols([14, 22, 30, 16, 10, 8], WL), ["Tema", "Situación actual", "Cambio propuesto", "Efecto en el Perfil", "Decisión de", "Tipo"], D.cambios.map((c) => [...c])),
  A.fuente("Fuente: elaboración propia a partir del Diagnóstico Integral V3, el EVAR final (19/10/2026), el E5 final (01/11/2026) y la Matriz de Costos v1.6 (03/11/2026)."),
];

// ------------------------------------------------------------------ 6. Riesgos
const cnt = {}; D.riesgos.forEach((r) => { const n = nivel(r[5], r[6]); cnt[n] = (cnt[n] || 0) + 1; });
const s6 = [
  A.h1("6. GESTIÓN DE RIESGOS DEL PLAN"),
  A.para("Se entiende por riesgo todo factor que podría impedir o dificultar la culminación del Perfil. La matriz distingue la causa (origen del factor), el efecto sobre el estudio (entregables afectados), la probabilidad y el impacto en escala de 1 a 5, el nivel resultante (P × I: Bajo menor que 6, Medio de 6 a 11, Alto de 12 a 19, Muy alto de 20 a 25), las medidas prospectivas y correctivas de mitigación, las medidas reactivas de respuesta, el responsable y el indicador o disparador que activa la respuesta. El riesgo R20 se registra como oportunidad (impacto positivo). La matriz completa, con fórmulas y mapa de calor, está en la hoja 4 del libro Excel."),
  A.cap("Tabla 9. Resumen de la matriz de riesgos por nivel"),
  A.tabla(cols([30, 20, 50], WL), ["Nivel", "Cantidad", "Riesgos"], ["Muy alto", "Alto", "Medio", "Bajo"].map((n) => [n, String(cnt[n] || 0), D.riesgos.filter((r) => nivel(r[5], r[6]) === n).map((r) => r[0]).join(", ") || "—"])),
  A.cap("Tabla 10. Matriz de riesgos del Plan de Contingencia"),
  A.tabla(cols([4, 8, 16, 14, 14, 4, 4, 6, 18, 12], WL), ["ID", "Categoría", "Riesgo", "Causa", "Efecto en el Perfil", "P", "I", "Nivel", "Medidas de mitigación (prospectivas y correctivas)", "Medidas reactivas"],
    D.riesgos.map((r) => [r[0], r[1], r[2], r[3], r[4], String(r[5]), String(r[6]), nivel(r[5], r[6]), r[7], r[8]])),
  A.cap("Tabla 11. Responsables, disparadores y entregables afectados"),
  A.tabla(cols([5, 25, 20, 36, 14], WL), ["ID", "Riesgo", "Responsable", "Indicador / disparador", "Entregables"], D.riesgos.map((r) => [r[0], r[2], r[9], r[10], r[11]])),
  A.fuente("Fuente: elaboración propia. Distinción de peligro, exposición, fragilidad y resiliencia aplicada al propio estudio: el peligro es el evento climático y la emergencia; la exposición, las brigadas, los bloques y la agenda institucional; la fragilidad, la dependencia de campo y de autoridades; la resiliencia, la capacidad de reprogramar, sustituir métodos y mantener el equipo."),
  A.h2("6.1 Monitoreo de riesgos"),
  A.para("La matriz se revisa cada mes y ante cada evento (aviso meteorológico, emergencia registrada en el SINPAD, cambio de autoridad, reasignación de personal). El tablero semanal del plan registra los disparadores activados y las medidas aplicadas; el informe quincenal a la SESDI consolida los riesgos activos y las decisiones requeridas."),
];

// ------------------------------------------------------------------ 7. Pendientes por entregable
const s7 = [
  A.h1("7. ACTIVIDADES PENDIENTES PROGRAMADAS POR ENTREGABLE"),
  A.para("La tabla organiza las actividades pendientes de cada entregable con su dependencia de campo, la ventana propuesta en el escenario base, el producto esperado y la prioridad. Los entregables culminados (E3, E5 y E8) solo registran complementos que alimentan a otros entregables."),
  A.cap("Tabla 12. Actividades pendientes por entregable"),
  A.tabla(cols([5, 8, 40, 7, 14, 16, 6, 10], WL), ["Entr.", "Avance", "Actividad pendiente", "Campo", "Ventana propuesta", "Producto", "Prior.", "Dependencia"],
    D.pendientes.map((p) => [p[0], `${ent[p[0]][4]} %`, p[1], p[2], p[3], p[4], p[5], p[6]])),
  A.fuente("Fuente: Diagnóstico Integral V3 («Vacíos para cerrar el Perfil»), presentación a la SESDI del 03/11/2026 y elaboración propia."),
];

// ------------------------------------------------------------------ 8. Administrativo
const s8 = [
  A.h1("8. ASPECTOS ADMINISTRATIVOS Y DE GESTIÓN"),
  A.para("El periodo del plan coincide con tres transiciones: el Estado de Emergencia, el cambio de autoridades regionales y municipales el 01/01/2027 (con segunda elección regional en Piura en diciembre de 2026) y el cierre del ejercicio fiscal 2026. A ello se suman las renovaciones de juntas directivas comunales y los posibles cambios de titulares en la cadena institucional de la ANIN."),
  A.para(D.electoral.erm + " " + D.electoral.segunda_vuelta),
  A.para(D.electoral.municipal + " " + D.electoral.comunal),
  A.para(D.electoral.nacional),
  A.cap("Tabla 13. Aspectos administrativos y de gestión: situación, efectos y acciones"),
  A.tabla(cols([13, 20, 22, 29, 8, 8], WL), ["Tema", "Situación", "Efecto esperado sobre el Perfil", "Acciones del plan", "Plazo", "Responsable"], D.admin.map((a) => [...a])),
  A.h2("8.1 Cómo repercute el cambio de autoridades en el Perfil"),
  A.bullet([A.b("Diagnóstico social y liberación de áreas (E2). "), "Las asambleas F-LA-01, las actas F-LA-04 y las fichas F-DS-03 de aceptación se sustentan en autoridades que cambian; las actas firmadas antes del 31/12/2026 conservan validez documental, pero la cooperación operativa (convocatorias, locales, acompañamiento) debe reconstruirse con las nuevas gestiones."]),
  A.bullet([A.b("Plan de gestión y sostenibilidad (E11, E12, E13). "), "Los arreglos institucionales (comités de gestión, convenios con municipalidades, MERESE con las JASS) requieren compromisos de las autoridades que ejercerán entre 2027 y 2030; conviene suscribirlos con los electos y no con las gestiones salientes."]),
  A.bullet([A.b("Información territorial. "), "La DRA Piura, las gerencias regionales y las oficinas municipales de catastro y Defensa Civil cambian de responsables; las solicitudes de información se reiteran en enero con los nuevos funcionarios."]),
  A.bullet([A.b("Riesgo de reapertura de acuerdos. "), "Las comunidades con juntas renovadas pueden pedir volver a discutir la libre disponibilidad; se mitiga con la inscripción de los acuerdos en los libros de actas comunales y la ratificación en la primera asamblea ordinaria de 2027."]),
];

// ------------------------------------------------------------------ 9. Seguimiento, 10. Conclusiones, fuentes, anexos
const s9 = [
  A.h1("9. SEGUIMIENTO Y ACTUALIZACIÓN DEL PLAN"),
  A.cap("Tabla 14. Instrumentos de seguimiento"),
  A.tabla(cols([22, 44, 18, 16], A.W), ["Instrumento", "Contenido", "Frecuencia", "Responsable"], D.seguimiento.map((s) => [...s])),
  A.para("El avance ponderado del Perfil se reporta con los pesos vigentes de los 15 entregables. El plan se actualiza al cierre de cada fase y cuando se active un disparador de nivel Muy alto; cada actualización registra la versión, la fecha y el escenario vigente (optimista, base o pesimista)."),
  A.h1("10. CONCLUSIONES Y RECOMENDACIONES"),
  A.bullet("El D.S. N° 149-2026-PCM y el escenario del ENFEN confirman que la temporada de diciembre de 2026 a abril de 2027 no es compatible con el campo masivo que el Perfil todavía requiere (107 bloques y 60 lotes SUS sin tamizaje predial; 32 bloques sin ficha social). El plan concentra el campo prioritario en noviembre de 2026 y el residual en abril y mayo de 2027."),
  A.bullet("En el escenario base el Perfil se culmina el 31/05/2027 y la viabilidad se obtiene en junio de 2027 (dos meses más que la programación original). La Identificación mantiene su versión para revisión en diciembre de 2026 con los vacíos explícitos."),
  A.bullet("Los riesgos de nivel Muy alto son el acceso a los bloques por lluvias (R01), la indisponibilidad de autoridades por la emergencia (R03), el cambio de autoridades (R04), la reasignación del equipo de la DIME (R06) y los datos de campo incompletos para fijar metas (R09). Sus medidas requieren decisiones de la jefatura en noviembre de 2026."),
  A.bullet("Recomendación 1: aprobar el plan y el protocolo de seguridad antes del 10/11/2026 y asegurar transporte, SCTR, EPP y comunicaciones para la ventana de campo del 09 al 30 de noviembre."),
  A.bullet("Recomendación 2: formalizar la dedicación del equipo núcleo al Perfil durante la emergencia y programar en noviembre las certificaciones y contrataciones de 2027 (E4, E13, dron e imágenes, transporte)."),
  A.bullet("Recomendación 3: documentar todos los acuerdos con autoridades antes del 31/12/2026 y ejecutar la re-socialización con las gestiones electas en enero y febrero de 2027."),
  A.bullet("Recomendación 4: acordar con la SESDI y la OPMI el criterio de corte para fijar metas sobre superficie con libre disponibilidad verificada y presentar el resto como cartera condicionada, si en mayo de 2027 persisten bloques sin acta."),
  A.bullet("Recomendación 5: cotejar el texto oficial del decreto y su Anexo (inclusión de los 15 distritos, cómputo del plazo, artículos 2 y 3) y completar los costos unitarios del presupuesto antes de la aprobación del plan."),
  A.h1("FUENTES"),
  A.bullet("Presentación resumida de avance del Estudio de Perfil ante la SESDI (03/11/2026) y Diagnóstico Integral IN Piura V3; Matriz de Costos Paramétricos v1.6 (03/11/2026); informe final del Estudio de Suelos E5 (01/11/2026); informe final EVAR (19/10/2026); E2 B3 (03/11/2026); catálogo V8 (09/10/2026)."),
  ...D.decreto.fuentes.map((f) => A.bullet(`${f[0]} · ${f[1]}`)),
  ...D.clima.fuentes.map((f) => A.bullet(`${f[0]} · ${f[1]}`)),
  ...D.electoral.fuentes.map((f) => A.bullet(`${f[0]} · ${f[1]}`)),
  A.bullet("Guía General para la Identificación, Formulación y Evaluación de Proyectos de Inversión (DGPMI-MEF, 2022), Anexo 2 GdR-CCC; Ley N° 29664 y D.S. N° 048-2011-PCM; Ley N° 30754; D.S. N° 017-2018-MINAM; R.M. N° 178-2019-MINAM; R.M. N° 00213-2024-MINAM; D.S. N° 007-2013-EF; Ley N° 27867; Ley N° 27972; Ley N° 24656."),
  A.h1("ANEXO A. DISTRITOS DEL ÁMBITO Y VERIFICACIÓN DEL ANEXO DEL DECRETO", { nuevaPagina: true }),
  A.para("La prensa regional reporta 65 distritos del departamento de Piura en el Anexo del D.S. N° 149-2026-PCM, en las provincias de Piura, Ayabaca, Huancabamba, Morropón, Paita, Sullana, Talara y Sechura. La inclusión de cada distrito del proyecto debe cotejarse con el Anexo oficial publicado en El Peruano; la columna se deja en blanco para ese cotejo."),
  A.cap("Tabla A1. Distritos del ámbito del Proyecto IN Piura"),
  A.tabla(cols([8, 32, 24, 36], A.W), ["N.°", "Distrito", "Provincia", "Incluido en el Anexo del D.S. N° 149-2026-PCM"], D.distritos.map((d, i) => [String(i + 1), d[0], d[1], "[POR VERIFICAR]"])),
  A.h1("ANEXO B. ARCHIVOS COMPLEMENTARIOS"),
  A.bullet("Matrices_Plan_Contingencia_IN_Piura.xlsx: 9 hojas (resumen y avance, cronograma con diagrama semanal, presupuesto con fórmulas, matriz de riesgos con nivel y mapa de calor, pendientes por entregable, cambios de alcance y contrataciones, aspectos administrativos y calendario, seguimiento y verificación de distritos, notas y fuentes)."),
  A.bullet("Plan_Contingencia_IN_Piura.html: artefacto de consulta con el resumen ejecutivo, el cronograma, la matriz de riesgos, el presupuesto y los pendientes por entregable."),
  A.bullet("generador/datos_plan.py: datos maestros del plan (un solo origen para los tres archivos)."),
];

const doc = new Document({
  ...A.estilos(),
  sections: [
    { properties: A.pagina(), headers: { default: A.encabezado() }, footers: { default: A.pie("Plan de Contingencia del Estudio de Perfil · Proyecto IN Piura · CUI 2669244") }, children: [...portada, ...s1, ...s2] },
    { properties: A.pagina({ horizontal: true }), headers: { default: A.encabezado() }, footers: { default: A.pie("Plan de Contingencia del Estudio de Perfil · Proyecto IN Piura · CUI 2669244") }, children: [...s3, ...s4, ...s5, ...s6, ...s7, ...s8] },
    { properties: A.pagina(), headers: { default: A.encabezado() }, footers: { default: A.pie("Plan de Contingencia del Estudio de Perfil · Proyecto IN Piura · CUI 2669244") }, children: s9 },
  ],
});
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("ok", OUT); });
