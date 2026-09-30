# -*- coding: utf-8 -*-
"""
Diccionario único de campos de los formularios de Liberación de Áreas (F-LA-01 … F-LA-06).

Es la fuente común de las TRES vías de registro, para que todas produzcan el mismo envío:
  1. KoboToolbox       → XLSForm generado con estos nombres (la_plantillas.generar_xlsform).
  2. Digitación en app → formulario Streamlit construido con estos campos (pagina._tab_registro).
  3. Plantilla Excel   → libro ANIN con una hoja por formulario (la_plantillas.generar_plantilla_excel).
Las tres terminan en el mismo diccionario "aplanado" que produce Kobo (core.aplanar) y pasan por
core.validar_envio y la_db.importar, sin lógica paralela.

Las coordenadas se digitan en UTM WGS84 Zona 17S (Este/Norte) y se convierten al formato geopoint
de ODK ("lat lon alt precisión") para que la validación espacial sea idéntica a la de Kobo.
"""
from __future__ import annotations

import hashlib
import json
import math
import uuid as _uuid
from dataclasses import dataclass, field
from datetime import date, datetime

from pyproj import Transformer

from . import la_core as core

_INV = Transformer.from_crs(core.EPSG_UTM, 4326, always_xy=True)


# ------------------------------------------------------------------ listas de opciones (códigos Kobo)
def _op(*pares):
    return [(c, e) for c, e in pares]


OPCIONES: dict[str, list[tuple[str, str]]] = {
    "asistente": _op(("AP-1", "AP-1"), ("AP-2", "AP-2"), ("AP-3", "AP-3"), ("AP-4", "AP-4"), ("AP-5", "AP-5"),
                     ("AP-6", "AP-6"), ("ESP", "Especialista (ESP)")),
    "si_no": _op(("si", "Sí"), ("no", "No")),
    "tipo_evento": _op(("asamblea_comunal", "Asamblea comunal"), ("reunion_informativa", "Reunión informativa"),
                       ("taller_participativo", "Taller participativo"), ("visita_domiciliaria", "Visita domiciliaria"),
                       ("otro", "Otro")),
    "aceptacion": _op(("favorable", "Favorable"), ("condicionada", "Condicionada"), ("rechazo", "Rechazo"),
                      ("sin_definir", "Sin definir")),
    "tipo_titularidad": _op(("comunal", "Comunidad campesina"), ("privado", "Privado (con título)"),
                            ("posesionario", "Posesionario"), ("sucesion", "Sucesión"), ("estatal", "Estatal"),
                            ("sin_titular", "Sin titular / en conflicto")),
    "sucesion_estado": _op(("declarada", "Sucesión declarada"), ("en_tramite", "En trámite"),
                           ("sin_tramite", "Sin trámite")),
    "estado_civil": _op(("soltero", "Soltero(a)"), ("casado", "Casado(a)"), ("conviviente", "Conviviente"),
                        ("viudo", "Viudo(a)"), ("divorciado", "Divorciado(a)")),
    "docs_exhibidos": _op(("titulo", "Título de propiedad"), ("partida", "Partida registral"),
                          ("constancia_posesion", "Constancia de posesión"), ("certificado_posesion", "Certificado de posesión"),
                          ("acta_asamblea", "Acta de asamblea"), ("declaratoria_herederos", "Declaratoria de herederos"),
                          ("dni", "DNI"), ("ninguno", "Ninguno")),
    "uso_actual": _op(("bosque_seco", "Bosque seco"), ("matorral", "Matorral"), ("pastizal", "Pastizal / pastoreo"),
                      ("agricola", "Agrícola"), ("purma", "Purma / descanso"), ("eriazo", "Eriazo / suelo desnudo"),
                      ("otro", "Otro")),
    "ocupacion": _op(("sin_ocupacion", "Sin ocupación"), ("pastoreo", "Pastoreo"), ("cultivo", "Cultivo"),
                     ("vivienda", "Vivienda"), ("mixta", "Mixta")),
    "estado_la": [(e, e) for e in core.ESTADOS_LA],
    "tipo_unidad": _op(("bloque", "Bloque"), ("lote_sus", "Lote SUS"), ("vivero", "Vivero")),
    "pt_tipo": _op(("vertice", "Vértice"), ("hito", "Hito"), ("referencia", "Punto de referencia")),
    "posicion_sus": _op(("DENTRO", "Dentro del bloque"), ("CONTIGUO", "Contiguo (≤ 50 m)")),
    "interferencias": _op(("ninguna", "Ninguna"), ("bosque_natural", "Bosque natural"), ("cultivos", "Cultivos"),
                          ("vivienda", "Vivienda"), ("infraestructura", "Infraestructura"),
                          ("restos_arqueologicos", "Restos arqueológicos"), ("concesion_minera", "Concesión minera"),
                          ("linea_electrica", "Línea eléctrica"), ("otro", "Otro")),
    "conclusion_campo": _op(("sin_restriccion", "Sin restricción"), ("con_restriccion", "Con restricción"),
                            ("no_apto", "No apto")),
    "tipo_acta": _op(("A-03", "A-03 · Acta de asamblea comunal"), ("A-04", "A-04 · Acta del titular / posesionario"),
                     ("A-07", "A-07 · Acta de terreno para vivero")),
    "checklist": _op(("coordenadas", "Coordenadas"), ("croquis", "Croquis"), ("titularidad", "Titularidad"),
                     ("gratuidad", "Gratuidad"), ("plazo", "Plazo"), ("firmas", "Firmas"), ("fedatario", "Fedatario"),
                     ("dj", "Declaración jurada"), ("datos", "Datos completos"), ("legible", "Legible")),
    "fir_rol": _op(("titular", "Titular"), ("conyuge", "Cónyuge"), ("representante", "Representante"),
                   ("heredero", "Heredero"), ("presidente_cc", "Presidente comunal"), ("comunero", "Comunero"),
                   ("testigo", "Testigo"), ("fedatario", "Fedatario")),
    "fedatario_tipo": _op(("juez_paz", "Juez de paz"), ("notario", "Notario"),
                          ("teniente_gobernador", "Teniente gobernador"), ("ninguno", "Ninguno")),
    "estado_acta": _op(("completa", "Completa"), ("observada", "Observada"), ("pendiente", "Pendiente")),
    "modalidad": _op(("donacion", "Donación"), ("cesion_uso", "Cesión en uso"), ("convenio", "Convenio"),
                     ("alquiler", "Alquiler")),
    "pendiente": _op(("p0_5", "0–5 %"), ("p5_15", "5–15 %"), ("p15_30", "15–30 %"), ("p30", "> 30 %")),
    "agua_fuente": _op(("canal", "Canal"), ("rio", "Río"), ("quebrada", "Quebrada"), ("manantial", "Manantial"),
                       ("pozo", "Pozo"), ("red_publica", "Red pública"), ("ninguna", "Ninguna")),
    "agua_permanencia": _op(("permanente", "Permanente"), ("temporal", "Temporal")),
    "acceso_tipo": _op(("carretera", "Carretera"), ("trocha", "Trocha carrozable"), ("herradura", "Camino de herradura"),
                       ("sin_camino", "Sin camino")),
    "energia": _op(("red", "Red eléctrica"), ("cerca", "Red cercana"), ("no", "Sin energía")),
    "nivel": _op(("nulo", "Nulo"), ("bajo", "Bajo"), ("medio", "Medio"), ("alto", "Alto")),
    "disposicion": _op(("favorable", "Favorable"), ("condicionada", "Condicionada"), ("negativa", "Negativa")),
    # F-LA-05 (se registra en el aplicativo o en la plantilla; no tiene formulario Kobo)
    "doc_tipo": _op(("constancia_busqueda", "Constancia de búsqueda"), ("certificado_busqueda_catastral", "Certificado de búsqueda catastral"),
                    ("partida_registral", "Partida registral"), ("titulo", "Título"), ("vigencia_poder", "Vigencia de poder"),
                    ("acta_asamblea_certificada", "Acta de asamblea certificada"), ("constancia_posesion", "Constancia de posesión"),
                    ("sucesion_intestada", "Sucesión intestada"), ("consulta_sinabip", "Consulta SINABIP"), ("dni", "DNI"),
                    ("otro", "Otro")),
    "doc_entidad": _op(("SUNARP", "SUNARP"), ("SBN", "SBN"), ("GORE / DRA Piura", "GORE / DRA Piura"), ("COFOPRI", "COFOPRI"),
                       ("Municipalidad", "Municipalidad"), ("SERFOR / ATFFS", "SERFOR / ATFFS"), ("Min. Cultura", "Min. Cultura"),
                       ("INGEMMET", "INGEMMET"), ("Juzgado de paz", "Juzgado de paz"), ("Otra", "Otra")),
    "doc_resultado": _op(("positivo", "Positivo"), ("negativo", "Negativo"), ("en_tramite", "En trámite")),
}


# ------------------------------------------------------------------ especificación de campos
@dataclass
class Campo:
    nombre: str                       # nombre XML (el mismo de Kobo)
    etiqueta: str
    tipo: str                         # text, textarea, int, decimal, date, select_one, select_multiple, unidad,
    #                                   unidades, utm (punto), grupo (título)
    lista: str | None = None          # clave de OPCIONES
    requerido: bool = False
    ayuda: str = ""
    relevante: dict | None = None     # {"campo": [valores]} → se muestra solo si el campo tiene uno de los valores
    minimo: float | None = None
    maximo: float | None = None


@dataclass
class Repeat:
    nombre: str                       # r_puntos / r_firmantes / vertices del lote SUS
    etiqueta: str
    campos: list[Campo]
    hoja: str                         # hoja de la plantilla Excel
    relevante: dict | None = None


@dataclass
class Formulario:
    form_id: str
    codigo: str
    nombre: str
    campos: list[Campo]
    repeats: list[Repeat] = field(default_factory=list)
    con_predio: bool = False          # cod_predio = {unidad}-P{nn}
    kobo: bool = True                 # F-LA-05 no tiene formulario Kobo


C = Campo
_COMUNES = [
    C("hoy", "Fecha de registro", "date", requerido=True),
    C("asistente", "Asistente predial", "select_one", "asistente", requerido=True,
      ayuda="Debe coincidir con el asistente asignado a la unidad en el catálogo V6 (o ESP)."),
]
_UNIDAD = [C("unidad", "Unidad (bloque / lote SUS / vivero)", "unidad", requerido=True)]
_PREDIO = [C("n_predio", "N.° de predio dentro de la unidad", "int", requerido=True, minimo=0, maximo=99,
             ayuda="Genera el código {UNIDAD}-P{nn}. Use 0 solo en actas comunales A-03 que abarcan toda la unidad.")]

FORMULARIOS: dict[str, Formulario] = {}

FORMULARIOS["f_la_01_reunion"] = Formulario(
    "f_la_01_reunion", "F-LA-01", "Reunión / asamblea informativa",
    _COMUNES + [
        C("unidades", "Unidades socializadas", "unidades", requerido=True,
          ayuda="Códigos separados por espacio (p. ej.: 27 3 SUS-058)."),
        C("tipo_evento", "Tipo de evento", "select_one", "tipo_evento", requerido=True),
        C("fecha_evento", "Fecha del evento", "date", requerido=True),
        C("centro_poblado", "Centro poblado", "text"),
        C("comunidad", "Comunidad campesina", "text"),
        C("gps_evento", "Punto GPS del evento", "utm"),
        C("asist_hombres", "Asistentes hombres", "int", minimo=0),
        C("asist_mujeres", "Asistentes mujeres", "int", minimo=0),
        C("titulares_presentes", "Titulares / posesionarios presentes", "int", minimo=0),
        C("aceptacion", "Aceptación del proyecto", "select_one", "aceptacion", requerido=True),
        C("alertas", "Alertas sociales", "textarea"),
        C("acuerdos", "Acuerdos", "textarea"),
        C("a01_suscrita", "Acta A-01 suscrita", "select_one", "si_no"),
        C("a02_suscrita", "Acta A-02 (lista de asistencia) suscrita", "select_one", "si_no"),
    ])

FORMULARIOS["f_la_02_titular"] = Formulario(
    "f_la_02_titular", "F-LA-02", "Ficha predial y socio-territorial",
    _COMUNES + _UNIDAD + _PREDIO + [
        C("consentimiento", "Consentimiento para el uso de datos personales (Ley 29733)", "select_one", "si_no", requerido=True),
        C("tipo_titularidad", "Tipo de titularidad", "select_one", "tipo_titularidad", requerido=True),
        C("sucesion_estado", "Estado de la sucesión", "select_one", "sucesion_estado", relevante={"tipo_titularidad": ["sucesion"]}),
        C("tit_nombres", "Nombres del titular", "text", relevante={"tipo_titularidad": ["privado", "posesionario", "sucesion"]}),
        C("tit_apellidos", "Apellidos del titular", "text", relevante={"tipo_titularidad": ["privado", "posesionario", "sucesion"]}),
        C("tit_dni", "DNI del titular (8 dígitos)", "text", relevante={"tipo_titularidad": ["privado", "posesionario", "sucesion"]}),
        C("tit_celular", "Celular del titular", "text", relevante={"tipo_titularidad": ["privado", "posesionario", "sucesion"]}),
        C("tit_estado_civil", "Estado civil", "select_one", "estado_civil", relevante={"tipo_titularidad": ["privado", "posesionario", "sucesion"]}),
        C("cony_nombre", "Nombre del cónyuge / conviviente", "text", relevante={"tit_estado_civil": ["casado", "conviviente"]}),
        C("cony_dni", "DNI del cónyuge", "text", relevante={"tit_estado_civil": ["casado", "conviviente"]}),
        C("cc_nombre", "Comunidad campesina", "text", relevante={"tipo_titularidad": ["comunal"]}),
        C("cc_ruc", "RUC de la comunidad", "text", relevante={"tipo_titularidad": ["comunal"]}),
        C("cc_presidente", "Presidente comunal", "text", relevante={"tipo_titularidad": ["comunal"]}),
        C("cc_presidente_dni", "DNI del presidente comunal", "text", relevante={"tipo_titularidad": ["comunal"]}),
        C("cc_presidente_cel", "Celular del presidente comunal", "text", relevante={"tipo_titularidad": ["comunal"]}),
        C("cc_partida", "Partida de la comunidad", "text", relevante={"tipo_titularidad": ["comunal"]}),
        C("est_entidad", "Entidad estatal titular", "text", relevante={"tipo_titularidad": ["estatal"]}),
        C("est_contacto", "Contacto de la entidad", "text", relevante={"tipo_titularidad": ["estatal"]}),
        C("doc_partida", "N.° de partida / título", "text"),
        C("docs_exhibidos", "Documentos exhibidos", "select_multiple", "docs_exhibidos"),
        C("predio_nombre", "Nombre del predio", "text"),
        C("predio_area_decl_ha", "Área declarada del predio (ha)", "decimal", minimo=0),
        C("predio_area_unidad_ha", "Área del predio dentro de la unidad (ha)", "decimal", minimo=0),
        C("uso_actual", "Uso actual", "select_one", "uso_actual"),
        C("ocupacion", "Ocupación", "select_one", "ocupacion"),
        C("conflicto_linderos_det", "Conflictos de linderos / observaciones", "textarea"),
        C("aceptacion", "Aceptación del titular", "select_one", "aceptacion", requerido=True),
        C("estado_la_propuesto", "Estado LA propuesto", "select_one", "estado_la"),
        C("cod_titular_prev", "Código de titular ya registrado (T0000), si existe", "text"),
    ], con_predio=True)

FORMULARIOS["f_la_03_inspeccion"] = Formulario(
    "f_la_03_inspeccion", "F-LA-03", "Inspección in situ y croquis",
    _COMUNES + _UNIDAD + _PREDIO + [
        C("fecha_insp", "Fecha de inspección", "date", requerido=True),
        C("punto_interior", "Punto interior del predio", "utm", requerido=True),
        C("posicion_sus_campo", "Posición del lote SUS declarada en campo", "select_one", "posicion_sus",
          relevante={"tipo_unidad": ["lote_sus"]}),
        C("interferencias", "Interferencias", "select_multiple", "interferencias"),
        C("interf_detalle", "Detalle de interferencias", "textarea"),
        C("uso_observado", "Uso observado", "select_one", "uso_actual"),
        C("conclusion_campo", "Conclusión de campo", "select_one", "conclusion_campo", requerido=True),
    ],
    repeats=[
        Repeat("r_puntos", "Puntos de control / vértices del predio",
               [C("pt_gps", "Punto", "utm", requerido=True), C("pt_tipo", "Tipo de punto", "select_one", "pt_tipo")],
               "F-LA-03_puntos"),
        Repeat("poligono_sus", "Vértices del lote SUS (en orden, UTM 17S)",
               [C("orden", "Orden", "int", requerido=True), C("pt_gps", "Vértice", "utm", requerido=True)],
               "F-LA-03_vertices_SUS", relevante={"tipo_unidad": ["lote_sus"]}),
    ], con_predio=True)

FORMULARIOS["f_la_04_actas"] = Formulario(
    "f_la_04_actas", "F-LA-04", "Registro de suscripción de actas",
    _COMUNES + _UNIDAD + _PREDIO + [
        C("tipo_acta", "Tipo de acta", "select_one", "tipo_acta", requerido=True),
        C("cod_doc", "Código del acta", "text", ayuda="Si se deja vacío se genera como A04-{UNIDAD}-P{nn}."),
        C("fecha_acta", "Fecha de suscripción", "date", requerido=True),
        C("area_comprometida_ha", "Área comprometida (ha)", "decimal", minimo=0),
        C("plazo_consignado", "Plazo consignado", "text"),
        C("checklist", "Lista de verificación del acta", "select_multiple", "checklist"),
        C("fedatario_tipo", "Fedatario", "select_one", "fedatario_tipo"),
        C("quorum_pct", "Quórum de la asamblea (%)", "decimal", minimo=0, maximo=100, relevante={"tipo_acta": ["A-03"]}),
        C("libro_actas", "Transcrita al libro de actas", "select_one", "si_no", relevante={"tipo_acta": ["A-03"]}),
        C("estado_acta", "Estado del acta", "select_one", "estado_acta", requerido=True),
        C("fecha_entrega_cd", "Fecha de entrega a control documentario", "date"),
    ],
    repeats=[Repeat("r_firmantes", "Firmantes",
                    [C("fir_nombre", "Nombre del firmante", "text", requerido=True), C("fir_dni", "DNI", "text"),
                     C("fir_rol", "Rol", "select_one", "fir_rol", requerido=True),
                     C("fir_huella", "Huella digital", "select_one", "si_no")],
                    "F-LA-04_firmantes")],
    con_predio=True)

FORMULARIOS["f_la_05_documentos"] = Formulario(
    "f_la_05_documentos", "F-LA-05", "Constancia de búsqueda documental",
    [C("unidad", "Unidad", "unidad", requerido=True),
     C("n_predio", "N.° de predio (vacío = toda la unidad)", "int", minimo=0, maximo=99),
     C("tipo", "Tipo de documento", "select_one", "doc_tipo", requerido=True),
     C("entidad", "Entidad", "select_one", "doc_entidad"),
     C("fecha", "Fecha", "date", requerido=True),
     C("resultado", "Resultado", "select_one", "doc_resultado"),
     C("n_partida", "N.° de partida / expediente", "text"),
     C("descripcion", "Descripción / hallazgos", "textarea"),
     C("archivo_url", "Enlace al archivo (Storage / Drive)", "text"),
     C("registrado_por", "Registrado por", "text", requerido=True)],
    kobo=False)

FORMULARIOS["f_la_06_vivero"] = Formulario(
    "f_la_06_vivero", "F-LA-06", "Evaluación de terreno – Vivero Central",
    _COMUNES + [
        C("cod_vivero", "Código de la alternativa (VIV-nn)", "text", requerido=True),
        C("provincia", "Provincia", "text", requerido=True),
        C("distrito", "Distrito", "text", requerido=True),
        C("alt_nombre", "Nombre de la alternativa", "text", requerido=True),
        C("modalidad", "Modalidad de disponibilidad", "select_one", "modalidad"),
        C("tipo_titularidad", "Tipo de titularidad", "select_one", "tipo_titularidad"),
        C("titular_nombre", "Titular", "text"),
        C("gps_centro", "Punto central del terreno", "utm", requerido=True),
        C("area_ha", "Área disponible (ha)", "decimal", minimo=0),
        C("pendiente", "Pendiente", "select_one", "pendiente"),
        C("agua_fuente", "Fuente de agua", "select_one", "agua_fuente"),
        C("agua_caudal_ls", "Caudal (l/s)", "decimal", minimo=0),
        C("agua_permanencia", "Permanencia del agua", "select_one", "agua_permanencia"),
        C("acceso_tipo", "Tipo de acceso", "select_one", "acceso_tipo"),
        C("acceso_camion", "Acceso para camión", "select_one", "si_no"),
        C("energia", "Energía eléctrica", "select_one", "energia"),
        C("dist_bloques_km", "Distancia a los bloques (km)", "decimal", minimo=0),
        C("inundabilidad", "Peligro de inundación", "select_one", "nivel"),
        C("deslizamiento", "Peligro de deslizamiento", "select_one", "nivel"),
        C("disposicion", "Disposición del titular", "select_one", "disposicion"),
        C("firmaria_a07", "Firmaría el acta A-07", "select_one", "si_no"),
    ])

FORMS_ENVIO = [f for f in FORMULARIOS if f in core.FORMULARIOS]   # los que pasan por validar_envio / importar
UTM_SUFIJOS = ("este", "norte", "alt", "prec")


# ------------------------------------------------------------------ utilidades
def etiqueta_opcion(lista: str, codigo) -> str:
    return dict(OPCIONES.get(lista, [])).get(codigo, "" if codigo is None else str(codigo))


def codigo_opcion(lista: str, valor) -> str | None:
    """Acepta el código o la etiqueta (sin distinguir mayúsculas) → código."""
    if valor is None or (isinstance(valor, float) and math.isnan(valor)):
        return None
    v = str(valor).strip()
    if not v:
        return None
    for c, e in OPCIONES.get(lista, []):
        if v == c or v.lower() == c.lower() or v.lower() == e.lower():
            return c
    return v


def codigos_multiples(lista: str, valor) -> list[str]:
    if valor is None or (isinstance(valor, float) and math.isnan(valor)):
        return []
    if isinstance(valor, (list, tuple)):
        partes = list(valor)
    else:
        partes = str(valor).replace(",", " ").replace(";", " ").split()
    return [codigo_opcion(lista, p) for p in partes if str(p).strip()]


def es_relevante(campo: Campo | Repeat, valores: dict) -> bool:
    if not campo.relevante:
        return True
    return all(valores.get(k) in v for k, v in campo.relevante.items())


def utm_a_geopoint(este, norte, alt=None, prec=None) -> str | None:
    """UTM 17S → 'lat lon alt precisión' (formato geopoint de ODK/Kobo)."""
    e, n = core._num(este), core._num(norte)
    if e is None or n is None:
        return None
    lon, lat = _INV.transform(e, n)
    return f"{lat:.8f} {lon:.8f} {core._num(alt) or 0:g} {core._num(prec) if core._num(prec) is not None else 0:g}"


def geopoint_a_utm(txt) -> dict:
    p = core.parse_geopoint(txt) if isinstance(txt, str) else None
    if not p:
        return {}
    return {"este": p["este"], "norte": p["norte"], "alt": p["alt"], "prec": p["prec"]}


def utm_en_rango(este, norte) -> bool:
    e, n = core._num(este), core._num(norte)
    return e is not None and n is not None and core.E_MIN <= e <= core.E_MAX and core.N_MIN <= n <= core.N_MAX


def cod_predio(unidad: str | None, n_predio) -> str | None:
    n = core._num(n_predio)
    if not unidad or n is None:
        return None
    return f"{unidad}-P{int(n):02d}"


def _fecha_txt(v) -> str | None:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = str(v).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s[:19] if "%H" in fmt else s[:10], fmt).date().isoformat()
        except ValueError:
            continue
    return s


def _normalizar(c: Campo, v):
    if c.tipo in ("int",):
        n = core._num(v)
        return int(n) if n is not None else None
    if c.tipo == "decimal":
        return core._num(v)
    if c.tipo == "date":
        return _fecha_txt(v)
    if c.tipo == "select_one":
        return codigo_opcion(c.lista, v)
    if c.tipo == "select_multiple":
        return " ".join(codigos_multiples(c.lista, v)) or None
    if c.tipo == "unidades":
        return " ".join(str(v).replace(",", " ").replace(";", " ").split()) if v not in (None, "") else None
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    if isinstance(v, float) and v.is_integer() and c.nombre.endswith(("dni", "ruc", "celular", "cel")):
        v = int(v)
    s = str(v).strip()
    return s or None


def nuevo_uuid(prefijo: str = "app") -> str:
    return f"{prefijo}-{_uuid.uuid4()}"


def uuid_contenido(form_id: str, valores: dict, prefijo: str = "xls") -> str:
    """Identificador determinista de una fila de plantilla: re-importar el mismo archivo da DUPLICADO."""
    base = json.dumps({"f": form_id, "v": valores}, ensure_ascii=False, sort_keys=True, default=str)
    return f"{prefijo}-{hashlib.sha1(base.encode('utf-8')).hexdigest()[:24]}"


# ------------------------------------------------------------------ construcción del envío
def construir_registro(form_id: str, valores: dict, catalogo: dict[str, dict], repeats: dict[str, list[dict]] | None = None,
                       kobo_uuid: str | None = None, base: dict | None = None) -> tuple[dict, list[str]]:
    """Valores capturados (app / plantilla) → envío aplanado igual al de Kobo, listo para core.validar_envio.
    `valores` usa los nombres de campo; los puntos 'utm' vienen como {campo}_este, {campo}_norte, {campo}_alt,
    {campo}_prec. `base` es el payload previo (edición): se conservan sus campos no editados (adjuntos, metadatos).
    Devuelve (registro, avisos). Los avisos son de captura (campos obligatorios, rangos UTM); las reglas de
    negocio las aplica core.validar_envio."""
    f = FORMULARIOS[form_id]
    avisos: list[str] = []
    reg: dict = {k: v for k, v in (base or {}).items()}
    for c in f.campos:
        if c.tipo == "utm":
            e, n = valores.get(f"{c.nombre}_este"), valores.get(f"{c.nombre}_norte")
            if core._num(e) is None and core._num(n) is None:
                if valores.get(c.nombre):                      # geopoint ya en formato ODK
                    reg[c.nombre] = valores[c.nombre]
                elif f"{c.nombre}_este" in valores or f"{c.nombre}_norte" in valores:
                    reg.pop(c.nombre, None)                    # el usuario borró el punto
                if c.requerido and not reg.get(c.nombre):
                    avisos.append(f"Falta {c.etiqueta} (Este / Norte)")
                continue
            if not utm_en_rango(e, n):
                avisos.append(f"{c.etiqueta}: Este {e} / Norte {n} fuera del rango UTM 17S del ámbito "
                              f"(E {core.E_MIN:,}–{core.E_MAX:,}; N {core.N_MIN:,}–{core.N_MAX:,})")
            reg[c.nombre] = utm_a_geopoint(e, n, valores.get(f"{c.nombre}_alt"), valores.get(f"{c.nombre}_prec"))
            continue
        if c.nombre not in valores:
            continue
        v = _normalizar(c, valores.get(c.nombre))
        if v is None:
            reg.pop(c.nombre, None)
        else:
            reg[c.nombre] = v
    for c in f.campos:
        if c.requerido and c.tipo != "utm" and reg.get(c.nombre) in (None, "") and es_relevante(c, reg):
            avisos.append(f"Falta {c.etiqueta}")

    unidad = reg.get("unidad")
    if unidad:
        reg["cod_unidad"] = unidad
        cat = catalogo.get(unidad, {})
        reg["tipo_unidad"] = cat.get("tipo_unidad") or reg.get("tipo_unidad")
        for k in ("provincia", "distrito"):
            if cat.get(k):
                reg[k] = cat[k]
        if cat.get("tipo_unidad") == "lote_sus":
            reg["bloque_ref"] = cat.get("bloque_ref")
            if cat.get("area_bloque_ha") not in (None, ""):
                reg["area_bloque_ha"] = str(cat.get("area_bloque_ha"))
        if not cat:
            avisos.append(f"La unidad «{unidad}» no está en el catálogo vigente")
    if form_id == "f_la_01_reunion" and reg.get("unidades"):
        reg["cod_unidad"] = reg["unidades"]
        primera = catalogo.get(reg["unidades"].split()[0], {})
        for k in ("provincia", "distrito"):
            if primera.get(k) and not reg.get(k):
                reg[k] = primera[k]
        desconocidas = [u for u in reg["unidades"].split() if u not in catalogo]
        if desconocidas:
            avisos.append(f"Unidades fuera del catálogo: {', '.join(desconocidas)}")
    if form_id == "f_la_06_vivero":
        reg["cod_unidad"] = reg.get("cod_vivero")
        reg["tipo_unidad"] = "vivero"
    if f.con_predio:
        cp = cod_predio(unidad, reg.get("n_predio"))
        if cp and not (form_id == "f_la_04_actas" and int(core._num(reg.get("n_predio")) or 0) == 0):
            reg["cod_predio"] = cp
        elif form_id == "f_la_04_actas":
            reg.pop("cod_predio", None)
    if form_id == "f_la_04_actas" and not reg.get("cod_doc") and unidad and reg.get("tipo_acta"):
        sufijo = f"-P{int(core._num(reg.get('n_predio')) or 0):02d}" if core._num(reg.get("n_predio")) else ""
        reg["cod_doc"] = f"{reg['tipo_acta'].replace('-', '')}-{unidad}{sufijo}"

    # repeats
    for rp in f.repeats:
        if repeats is None or rp.nombre not in repeats:
            continue
        filas = [x for x in repeats[rp.nombre] if any(v not in (None, "") and not (isinstance(v, float) and math.isnan(v))
                                                      for v in x.values())]
        if rp.nombre == "poligono_sus":
            filas = sorted(filas, key=lambda x: core._num(x.get("orden")) or 0)
            pts = [utm_a_geopoint(x.get("pt_gps_este"), x.get("pt_gps_norte"), x.get("pt_gps_alt"), x.get("pt_gps_prec"))
                   for x in filas]
            pts = [p for p in pts if p]
            if pts:
                if pts[0] != pts[-1]:
                    pts.append(pts[0])
                reg["poligono_sus"] = ";".join(pts)
            else:
                reg.pop("poligono_sus", None)
            continue
        salida = []
        for x in filas:
            item = {}
            for c in rp.campos:
                if c.tipo == "utm":
                    gp = utm_a_geopoint(x.get(f"{c.nombre}_este"), x.get(f"{c.nombre}_norte"),
                                        x.get(f"{c.nombre}_alt"), x.get(f"{c.nombre}_prec"))
                    if gp:
                        item[c.nombre] = gp
                else:
                    v = _normalizar(c, x.get(c.nombre))
                    if v is not None:
                        item[c.nombre] = v
            if item:
                salida.append(item)
        reg[rp.nombre] = salida

    if kobo_uuid:
        reg["meta_instanceID"] = f"uuid:{kobo_uuid}"
    return reg, avisos


def registro_a_valores(form_id: str, payload: dict) -> tuple[dict, dict[str, list[dict]]]:
    """Inverso de construir_registro (para editar un envío ya guardado, sea de Kobo, app o plantilla)."""
    f = FORMULARIOS[form_id]
    val: dict = {}
    for c in f.campos:
        v = payload.get(c.nombre)
        if c.nombre == "unidad" and not v:
            v = payload.get("cod_unidad")
        if c.tipo == "utm":
            for k, x in geopoint_a_utm(v).items():
                val[f"{c.nombre}_{k}"] = x
        elif c.tipo == "select_multiple":
            val[c.nombre] = codigos_multiples(c.lista, v)
        else:
            val[c.nombre] = v
    reps: dict[str, list[dict]] = {}
    for rp in f.repeats:
        if rp.nombre == "poligono_sus":
            txt = payload.get("poligono_sus")
            filas = []
            if isinstance(txt, str) and txt.strip():
                tramos = [t for t in txt.strip().strip(";").split(";") if t.strip()]
                if len(tramos) > 1 and tramos[0].strip() == tramos[-1].strip():
                    tramos = tramos[:-1]
                for i, t in enumerate(tramos, 1):
                    u = geopoint_a_utm(t)
                    filas.append({"orden": i, "pt_gps_este": u.get("este"), "pt_gps_norte": u.get("norte")})
            reps[rp.nombre] = filas
            continue
        filas = []
        for x in payload.get(rp.nombre) or []:
            fila = {}
            for c in rp.campos:
                v = x.get(c.nombre, x.get(f"{rp.nombre}/{c.nombre}"))
                if c.tipo == "utm":
                    u = geopoint_a_utm(v)
                    fila[f"{c.nombre}_este"], fila[f"{c.nombre}_norte"] = u.get("este"), u.get("norte")
                else:
                    fila[c.nombre] = v
            filas.append(fila)
        reps[rp.nombre] = filas
    return val, reps
