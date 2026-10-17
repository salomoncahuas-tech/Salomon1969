// Definición de las carátulas de los anexos del Segundo Entregable (actividades b, c y d del TdR)
const INF = "Informe N° 008-2026-HSCM";
const FECHA = "Lima, octubre de 2026";
const PERIODO = "Segundo entregable — del 26 de septiembre al 27 de octubre de 2026 (días 31 a 60 de la Orden de Servicio)";
module.exports = {
  INF, FECHA, PERIODO,
  anexo1: { espIni: 24, espFinal: 20, numero: "1", actividad: "b)", informe: INF, fecha: FECHA, periodo: PERIODO,
    textoActividad: "Actualizar, integrar y consolidar la información geoespacial y el diagnóstico territorial de los bloques preliminares de intervención ubicados en los distritos de la provincia de Huancabamba y Morropón.",
    titulo: "Actualización, integración y consolidación de la información geoespacial y del diagnóstico territorial",
    subtitulo: "Bloques preliminares de intervención de las provincias de Huancabamba y Morropón — Diagnóstico Territorial Versión 5",
    ambito: "Provincias de Huancabamba y Morropón — 14 distritos; catálogo maestro de bloques V6",
    comprende: [
      ["1", "Informe de actualización a la Versión 5: fuentes integradas, consolidación por provincia y distrito, peligro integrado, estado del ecosistema (E6), suelos (E5), tamizaje predial de gabinete, diagnóstico social, riesgo y cartografía."],
      ["1-A", "Consolidado de los 107 bloques preliminares de intervención (coordenadas UTM, MSAVI, PI, índice de condición, situación predial, calicatas y fichas sociales)."],
      ["MV", "Vacíos de integración, herramientas geoespaciales del periodo y medios verificadores (volúmenes publicados, láminas SP, archivos nativos y commits)."],
    ] },
  anexo2: { espIni: 14, espFinal: 12, numero: "2", actividad: "c)", informe: INF, fecha: FECHA, periodo: PERIODO,
    textoActividad: "Realizar el seguimiento y reportar el avance en coordinación con el Especialista en Infraestructura Marrón, del Estudio de Reconocimiento de Suelos, incluyendo la revisión técnica, la formulación y el levantamiento de observaciones, así como la emisión de la conformidad —cuando corresponda— respecto del plan de trabajo, las actividades de campo y los informes presentados, verificando que la información generada cumpla con los requerimientos técnicos necesarios para su aplicación en el componente verde del proyecto IN Piura.",
    titulo: "Seguimiento técnico del Estudio de Suelos a Nivel de Reconocimiento",
    subtitulo: "Revisión de entregables, formulación y levantamiento de observaciones, conformidades y verificación de su aplicación en el componente verde",
    coordinacion: "Especialista en Infraestructura Marrón — Proyecto IN Piura",
    ambito: "117 bloques de intervención vigentes, 12,270.24 ha — provincias de Morropón, Huancabamba y Ayabaca",
    comprende: [
      ["2", "Informe de seguimiento técnico: servicio supervisado, coordinación con Infraestructura Marrón, estado de los entregables, seguimiento de la fase de campo y del laboratorio."],
      ["2-A", "Revisión técnica del Entregable N° 2 y matriz de observaciones OS-01 a OS-14 con su forma de levantamiento."],
      ["2-B", "Verificación de la aplicación al componente verde (infiltración, CUM, matriz suelo–intervención) y medios verificadores."],
    ] },
  anexo3: { espIni: 12, espFinal: 10, numero: "3", actividad: "d)", informe: INF, fecha: FECHA, periodo: PERIODO,
    textoActividad: "Realizar el seguimiento y reportar el avance, en coordinación con el Especialista en Intervenciones de Uso Sostenible, del Diagnóstico del Estado de los Ecosistemas y del Estudio de Oferta de Material Vegetal de Propagación, incluyendo la revisión y actualización de sus respectivos planes de trabajo, el seguimiento de las actividades de campo y la revisión técnica de los entregables, así como la formulación y el levantamiento de observaciones y la emisión de la conformidad, cuando corresponda.",
    titulo: "Seguimiento del Diagnóstico del Estado de los Ecosistemas y del Estudio de Oferta de Material Vegetal de Propagación",
    subtitulo: "Revisión y actualización de planes de trabajo, seguimiento de campo, revisión técnica de entregables y conformidades",
    coordinacion: "Especialista en Intervenciones de Uso Sostenible — Proyecto IN Piura",
    comprende: [
      ["3", "Contexto del cambio de modalidad y revisión y actualización de los planes de trabajo (E6 v1.1 a v2.0 «In house» y Plan de Trabajo Detallado del Perfil)."],
      ["3-A", "Diagnóstico del Estado de los Ecosistemas: seguimiento de actividades, revisión de insumos (OS), Nota Técnica N° 002-2026 e Informe E6; observaciones OE-01 a OE-13 y conformidad."],
      ["3-B", "Estudio de Oferta de Material Vegetal: plan del componente, revisión de insumos, viveros identificados, observaciones OM-01 a OM-08, conformidad y medios verificadores."],
    ] },
  anexo3A: { variante: "separador", numero: "3-A", actividad: "d)", informe: INF, fecha: FECHA, periodo: PERIODO,
    titulo: "Diagnóstico del Estado de los Ecosistemas",
    subtitulo: "Seguimiento, revisión técnica, observaciones y conformidad",
    coordinacion: "Especialista en Intervenciones de Uso Sostenible — Proyecto IN Piura" },
  anexo3B: { variante: "separador", numero: "3-B", actividad: "d)", informe: INF, fecha: FECHA, periodo: PERIODO,
    titulo: "Estudio de Oferta de Material Vegetal de Propagación",
    subtitulo: "Seguimiento, revisión técnica, observaciones y conformidad",
    coordinacion: "Especialista en Intervenciones de Uso Sostenible — Proyecto IN Piura" },
};
