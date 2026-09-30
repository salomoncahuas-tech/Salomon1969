# -*- coding: utf-8 -*-
"""Reportes del módulo de Liberación de Áreas: Excel consolidado ANIN y expediente preliminar PDF por unidad.
Los datos personales reservados (DNI, celular – Ley 29733) no se incluyen en ningún reporte."""
from __future__ import annotations

import io
import zipfile
from datetime import date, datetime

import pandas as pd
from fpdf import FPDF

from . import la_core as core
from . import la_db as db
from . import la_export as ex

SQL_MATRIZ = """
SELECT p.cod_predio, p.cod_unidad, u.tipo_unidad, u.provincia, u.distrito, u.asistente, p.nombre_predio,
       string_agg(DISTINCT t.cod_titular || ' ' || t.nombre, '; ') AS titulares,
       string_agg(DISTINCT t.tipo_titularidad, '; ') AS tipo_titularidad,
       p.area_decl_ha, p.area_unidad_ha, p.uso_actual, p.ocupacion, p.aceptacion, p.estado_la, p.clasificacion,
       p.nucleo, p.excepcion_manual, p.docs_completos, p.expediente_conforme, p.alertas
FROM la_predios p JOIN la_unidades u ON u.codigo = p.cod_unidad
LEFT JOIN la_predio_titular pt ON pt.cod_predio = p.cod_predio
LEFT JOIN la_titulares t ON t.cod_titular = pt.cod_titular
GROUP BY p.cod_predio, u.codigo ORDER BY u.provincia, u.distrito, p.cod_predio"""


def matriz(conn) -> pd.DataFrame:
    return db.df(conn, SQL_MATRIZ)


def semaforo(conn) -> tuple[pd.DataFrame, pd.DataFrame]:
    """(semáforo por unidad, resumen por asistente)."""
    u = db.df(conn, "SELECT codigo, tipo_unidad, provincia, distrito, asistente, area_ha FROM la_unidades WHERE activo")
    p = db.df(conn, "SELECT cod_predio, cod_unidad, estado_la, area_unidad_ha, nucleo FROM la_predios")
    if u.empty:
        return pd.DataFrame(), pd.DataFrame()
    filas = []
    for x in u.to_dict("records"):
        pp = p[p.cod_unidad == x["codigo"]] if not p.empty else p
        sem, pct = core.semaforo_unidad([{"estado": r["estado_la"], "area_ha": r["area_unidad_ha"], "nucleo": r["nucleo"]}
                                        for r in pp.to_dict("records")], float(x["area_ha"] or 0))
        filas.append({**x, "area_ha": float(x["area_ha"] or 0), "predios": len(pp), "pct_area_LA4": round(pct * 100, 1),
                      "semaforo": sem})
    s = pd.DataFrame(filas)
    por_asis = s.groupby("asistente", dropna=False).agg(
        unidades=("codigo", "count"), ha=("area_ha", "sum"), predios=("predios", "sum"),
        verdes=("semaforo", lambda z: int((z == "VERDE").sum())), ambar=("semaforo", lambda z: int((z == "AMBAR").sum())),
        rojas=("semaforo", lambda z: int((z == "ROJO").sum())),
        sin_datos=("semaforo", lambda z: int((z == "SIN_DATOS").sum()))).reset_index()
    return s, por_asis


def tabla_formulario(conn, form_id: str) -> pd.DataFrame:
    """Filas de un formulario con su origen, estado de importación y motivos."""
    if form_id == "f_la_05_documentos":
        return db.df(conn, "SELECT id, cod_doc, cod_unidad, cod_predio, tipo, entidad, fecha, resultado, n_partida, "
                           "descripcion, archivo_url, registrado_por, creado, editado, editado_por FROM la_documentos "
                           "ORDER BY id DESC")
    tabla = core.FORMULARIOS[form_id]["tabla"]
    d = db.df(conn, f"""SELECT t.*, COALESCE(r.origen, 'KOBO') AS origen, r.estado_import, r.motivos, r.editado, r.editado_por
                        FROM {tabla} t JOIN la_envios_raw r USING (kobo_uuid) ORDER BY r.creado DESC""")
    return d.drop(columns=[c for c in ("geom_wkt",) if c in d.columns])


def titulares_sin_reservados(conn) -> pd.DataFrame:
    return db.df(conn, """SELECT t.cod_titular, t.tipo_titularidad, t.nombre, t.representante, t.partida,
                                 string_agg(DISTINCT pt.cod_predio, '; ') AS predios
                          FROM la_titulares t LEFT JOIN la_predio_titular pt USING (cod_titular)
                          GROUP BY t.cod_titular ORDER BY t.cod_titular""")


def _filtrar(d: pd.DataFrame, asistente: str | None, distrito: str | None, unidades: set[str] | None) -> pd.DataFrame:
    if d.empty:
        return d
    if unidades is not None:
        col = next((c for c in ("cod_unidad", "codigo") if c in d.columns), None)
        if col:
            d = d[d[col].astype(str).apply(lambda v: any(u in unidades for u in str(v).split()))]
    if asistente and "asistente" in d.columns:
        d = d[d["asistente"] == asistente]
    if distrito and "distrito" in d.columns:
        d = d[d["distrito"] == distrito]
    return d


def excel_consolidado(conn, asistente: str | None = None, distrito: str | None = None) -> bytes:
    """Libro ANIN: resumen, matriz predial, semáforo, avance, cada formulario, documentos, titulares y observados."""
    s, por_asis = semaforo(conn)
    unidades = None
    if (asistente or distrito) and not s.empty:
        unidades = set(_filtrar(s, asistente, distrito, None)["codigo"].astype(str))
    m = _filtrar(matriz(conn), asistente, distrito, None)
    s_f = _filtrar(s, asistente, distrito, None)
    resumen = pd.DataFrame([
        {"indicador": "Unidades activas", "valor": len(s_f)},
        {"indicador": "Superficie de las unidades (ha)", "valor": round(float(s_f["area_ha"].sum()), 3) if not s_f.empty else 0},
        {"indicador": "Predios registrados", "valor": len(m)},
        {"indicador": "Predios con acta conforme o más (LA-4 a LA-6)",
         "valor": int(m["estado_la"].isin(["LA-4", "LA-5", "LA-6"]).sum()) if not m.empty else 0},
        {"indicador": "Predios observados (OBS)", "valor": int((m["estado_la"] == "OBS").sum()) if not m.empty else 0},
        {"indicador": "Predios con negativa (NEG)", "valor": int((m["estado_la"] == "NEG").sum()) if not m.empty else 0},
        {"indicador": "Unidades en VERDE", "valor": int((s_f["semaforo"] == "VERDE").sum()) if not s_f.empty else 0},
        {"indicador": "Unidades en ÁMBAR", "valor": int((s_f["semaforo"] == "AMBAR").sum()) if not s_f.empty else 0},
        {"indicador": "Unidades en ROJO", "valor": int((s_f["semaforo"] == "ROJO").sum()) if not s_f.empty else 0},
        {"indicador": "Filtro aplicado", "valor": " · ".join(x for x in (asistente, distrito) if x) or "Ninguno"},
    ])
    hojas = {"Resumen": ("RESUMEN – LIBERACIÓN DE ÁREAS", resumen, None),
             "Matriz_predial": ("MATRIZ PREDIAL – EXPEDIENTE PRELIMINAR DE LIBERACIÓN DE ÁREAS", m,
                                ["area_decl_ha", "area_unidad_ha"]),
             "Semaforo_unidad": ("SEMÁFORO POR UNIDAD", s_f, ["area_ha", "predios"])}
    if not asistente and not distrito:
        hojas["Avance_asistente"] = ("AVANCE POR ASISTENTE", por_asis, ["unidades", "ha", "predios"])
    for fid, f in core.FORMULARIOS.items():
        d = _filtrar(tabla_formulario(conn, fid), asistente, None, unidades)
        hojas[f["codigo"]] = (f"{f['codigo']} – {f['nombre'].upper()}", d, None)
    docs = _filtrar(tabla_formulario(conn, "f_la_05_documentos"), None, None, unidades)
    hojas["F-LA-05"] = ("F-LA-05 – CONSTANCIAS DE BÚSQUEDA Y DOCUMENTOS", docs, None)
    tit = titulares_sin_reservados(conn)
    if unidades is not None and not tit.empty and not m.empty:
        tit = tit[tit["predios"].fillna("").apply(lambda v: any(x.strip() in set(m["cod_predio"]) for x in v.split(";")))]
    hojas["Titulares"] = ("TITULARES (SIN DATOS RESERVADOS – LEY 29733)", tit, None)
    obs = _filtrar(db.df(conn, """SELECT kobo_uuid, form_id, cod_unidad, cod_predio, asistente, fecha_envio,
                                         COALESCE(origen,'KOBO') AS origen, motivos FROM la_envios_raw
                                  WHERE estado_import = 'OBSERVADO' ORDER BY cod_unidad, cod_predio"""), asistente, None, unidades)
    hojas["Observados"] = ("ENVÍOS OBSERVADOS POR RESOLVER", obs, None)
    return ex.exportar(hojas)


# ------------------------------------------------------------------ PDF por unidad
VERDE_RGB = (27, 77, 46)
_REEMPLAZOS = {"–": "-", "—": "-", "‘": "'", "’": "'", "“": '"', "”": '"', "…": "...", "•": "-", "≥": ">=",
               "≤": "<=", "→": "->", "·": "-", " ": " "}


def _t(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "-"
    if isinstance(v, bool):
        return "Sí" if v else "No"
    if isinstance(v, float):
        v = f"{v:,.3f}".rstrip("0").rstrip(".")
    s = str(v)
    for a, b in _REEMPLAZOS.items():
        s = s.replace(a, b)
    return s.encode("latin-1", "replace").decode("latin-1")


class _PDF(FPDF):
    titulo = ""

    def header(self):
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(*VERDE_RGB)
        self.cell(0, 4.5, _t("AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN"), 0, 1, "C")
        self.set_font("Helvetica", "", 8)
        self.cell(0, 4, _t("DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME"), 0, 1, "C")
        self.cell(0, 4, _t("SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN"), 0, 1, "C")
        self.set_font("Helvetica", "B", 11)
        self.cell(0, 6, _t(self.titulo), 0, 1, "C")
        self.set_draw_color(*VERDE_RGB)
        self.line(10, self.get_y() + 1, self.w - 10, self.get_y() + 1)
        self.set_text_color(0, 0, 0)
        self.ln(3)

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "I", 7)
        self.cell(0, 5, _t(f"Proyecto IN Piura - CUI 2669244 · Página {self.page_no()}/{{nb}} · "
                           f"Generado {datetime.now():%d/%m/%Y %H:%M} · Datos personales reservados omitidos (Ley 29733)"),
                  0, 0, "C")

    def seccion(self, t):
        self.ln(1)
        self.set_font("Helvetica", "B", 10)
        self.set_fill_color(*VERDE_RGB)
        self.set_text_color(255, 255, 255)
        self.cell(0, 6, _t(f"  {t}"), 0, 1, fill=True)
        self.set_text_color(0, 0, 0)
        self.ln(1)

    def pares(self, datos: list[tuple[str, object]], cols: int = 3):
        ancho = (self.w - 20) / cols
        for i, (k, v) in enumerate(datos):
            self.set_font("Helvetica", "B", 8)
            self.cell(ancho * 0.45, 5, _t(k + ":"))
            self.set_font("Helvetica", "", 8)
            self.cell(ancho * 0.55, 5, _t(v)[:40], 0, 1 if (i + 1) % cols == 0 else 0)
        if len(datos) % cols:
            self.ln(5)

    def tabla(self, d: pd.DataFrame, columnas: list[tuple[str, str, float]], vacio: str = "Sin registros."):
        if d is None or d.empty:
            self.set_font("Helvetica", "I", 8)
            self.cell(0, 5, _t(vacio), 0, 1)
            return
        total = sum(a for _, _, a in columnas)
        escala = (self.w - 20) / total
        self.set_font("Helvetica", "B", 7)
        self.set_fill_color(234, 242, 248)
        for _, h, a in columnas:
            self.cell(a * escala, 5, _t(h), 1, 0, "C", fill=True)
        self.ln(5)
        self.set_font("Helvetica", "", 7)
        for r in d.to_dict("records"):
            if self.get_y() > self.h - 20:
                self.add_page()
            for c, _, a in columnas:
                txt = _t(r.get(c))
                maxc = max(4, int(a * escala / 1.45))
                self.cell(a * escala, 4.5, txt if len(txt) <= maxc else txt[:maxc - 1] + ".", 1, 0)
            self.ln(4.5)


def pdf_expediente_unidad(conn, codigo: str) -> bytes:
    """Expediente preliminar de liberación de áreas de una unidad (bloque / lote SUS)."""
    u = db.df(conn, "SELECT * FROM la_unidades WHERE codigo = %s", (codigo,))
    if u.empty:
        raise ValueError(f"Unidad {codigo} inexistente en el catálogo.")
    u = u.to_dict("records")[0]
    s, _ = semaforo(conn)
    sx = s[s["codigo"] == codigo].to_dict("records")[0] if not s.empty and (s["codigo"] == codigo).any() else {}
    m = matriz(conn)
    m = m[m["cod_unidad"] == codigo] if not m.empty else m
    reun = db.df(conn, "SELECT * FROM la_reuniones ORDER BY fecha")
    if not reun.empty:
        reun = reun[reun["unidades"].fillna("").apply(lambda v: codigo in str(v).split())]
    insp = db.df(conn, "SELECT * FROM la_inspecciones WHERE cod_unidad = %s ORDER BY cod_predio, fecha", (codigo,))
    actas = db.df(conn, "SELECT * FROM la_actas WHERE cod_unidad = %s ORDER BY fecha", (codigo,))
    docs = db.df(conn, "SELECT * FROM la_documentos WHERE cod_unidad = %s ORDER BY fecha", (codigo,))
    obs = db.df(conn, """SELECT form_id, cod_predio, asistente, motivos FROM la_envios_raw
                         WHERE estado_import = 'OBSERVADO' AND cod_unidad = %s ORDER BY cod_predio""", (codigo,))
    if not obs.empty:
        obs["form_id"] = obs["form_id"].map(lambda f: core.FORMULARIOS.get(f, {}).get("codigo", f))

    pdf = _PDF(orientation="L", unit="mm", format="A4")
    pdf.titulo = f"EXPEDIENTE PRELIMINAR DE LIBERACIÓN DE ÁREAS – UNIDAD {codigo}"
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(True, margin=15)
    pdf.add_page()
    pdf.seccion("1. Datos de la unidad")
    pdf.pares([("Código", codigo), ("Tipo", u.get("tipo_unidad")), ("Provincia", str(u.get("provincia") or "").title()),
               ("Distrito", str(u.get("distrito") or "").replace("_", " ").title()), ("Área (ha)", u.get("area_ha")),
               ("Bloque de referencia", u.get("bloque_ref")), ("Asistente", f"{u.get('asistente') or ''} {u.get('asistente_nombre') or ''}"),
               ("Semáforo", sx.get("semaforo", "SIN_DATOS")), ("% área ≥ LA-4", sx.get("pct_area_LA4", 0)),
               ("Predios", len(m)), ("Posición SUS", u.get("posicion_sus")), ("Fecha", f"{date.today():%d/%m/%Y}")])
    pdf.seccion("2. Matriz predial")
    pdf.tabla(m, [("cod_predio", "Predio", 22), ("nombre_predio", "Nombre", 30), ("titulares", "Titular(es)", 55),
                  ("tipo_titularidad", "Titularidad", 22), ("area_unidad_ha", "Área (ha)", 16), ("aceptacion", "Aceptación", 20),
                  ("estado_la", "Estado", 13), ("clasificacion", "Clasificación", 45), ("nucleo", "Núcleo", 12),
                  ("alertas", "Alertas", 40)])
    pdf.seccion("3. Socialización (F-LA-01)")
    pdf.tabla(reun, [("fecha", "Fecha", 20), ("tipo_evento", "Evento", 30), ("centro_poblado", "Centro poblado", 35),
                     ("comunidad", "Comunidad", 35), ("asist_hombres", "H", 10), ("asist_mujeres", "M", 10),
                     ("aceptacion", "Aceptación", 22), ("a01_suscrita", "A-01", 12), ("acuerdos", "Acuerdos", 90)])
    pdf.seccion("4. Inspecciones in situ (F-LA-03) – UTM WGS84 17S")
    pdf.tabla(insp, [("cod_predio", "Predio", 22), ("fecha", "Fecha", 20), ("este", "Este (m)", 22), ("norte", "Norte (m)", 24),
                     ("precision_m", "Prec. (m)", 14), ("validacion_espacial", "Validación", 28), ("distancia_m", "Dist. (m)", 14),
                     ("interferencias", "Interferencias", 45), ("conclusion_campo", "Conclusión", 28),
                     ("area_sus_ha", "SUS (ha)", 14), ("posicion_sus_calc", "Posición SUS", 22)])
    pdf.seccion("5. Actas (F-LA-04)")
    pdf.tabla(actas, [("cod_doc", "Código", 28), ("tipo_acta", "Tipo", 12), ("fecha", "Fecha", 20), ("cod_predio", "Predio", 22),
                      ("area_comprometida_ha", "Área (ha)", 16), ("n_firmantes", "Firm.", 10), ("fedatario_tipo", "Fedatario", 24),
                      ("estado_acta", "Estado", 18), ("conforme", "Conforme", 15), ("fecha_entrega_cd", "Entrega CD", 20),
                      ("checklist", "Checklist", 90)])
    pdf.seccion("6. Búsqueda documental (F-LA-05)")
    pdf.tabla(docs, [("cod_doc", "Código", 32), ("cod_predio", "Predio", 22), ("tipo", "Tipo", 40), ("entidad", "Entidad", 28),
                     ("fecha", "Fecha", 20), ("resultado", "Resultado", 18), ("n_partida", "Partida / exp.", 30),
                     ("descripcion", "Descripción", 85)])
    pdf.seccion("7. Observaciones por resolver")
    pdf.tabla(obs, [("form_id", "Form.", 16), ("cod_predio", "Predio", 22), ("asistente", "Asist.", 14), ("motivos", "Motivos", 225)],
              vacio="Sin envíos observados.")
    out = pdf.output()
    return bytes(out)


def zip_expedientes(conn, codigos: list[str]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for c in codigos:
            z.writestr(f"Expediente_LA_{c}.pdf", pdf_expediente_unidad(conn, c))
    return buf.getvalue()
