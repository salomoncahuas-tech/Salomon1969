"""
IN Piura - Plan de Ingreso / Verificacion de Campo
Aplicacion web con Streamlit.
Restauracion de ecosistemas - Cuenca alta del rio Piura, Peru.
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import os
import uuid
import io
import csv
import tempfile

import database as db
import reports
from georeferenciacion import utm_a_latlon, latlon_a_utm
from odk_kobo import generar_xlsform, importar_csv_odk, importar_desde_kobo, KoBoClient

# ── Configuracion ─────────────────────────────────────────────────────────
st.set_page_config(
    page_title="IN Piura - Plan de Ingreso",
    page_icon="\U0001F331",
    layout="wide",
    initial_sidebar_state="expanded",
)

db.inicializar_bd()

# ── Constantes ────────────────────────────────────────────────────────────
TIPOS_INTERVENCION = [
    "Revegetacion", "Zanjas de infiltracion",
    "Terrazas de formacion lenta", "Diques de mamposteria", "Otras",
]
ESTADOS_BLOQUE = ["Pendiente", "En progreso", "Verificado"]
CONDICIONES_CLIMATICAS = [
    "Despejado", "Parcialmente nublado", "Nublado",
    "Lluvia ligera", "Lluvia moderada", "Lluvia intensa", "Neblina",
]
MICROCUENCAS = [
    "C1075-O9580","C1076-O9581","C1076-O9584","C1076-O9585","C1076-O9586",
    "C1076-O9587","C1076-O9588","C1076-O9589","C1076-O9592","C1077-O9566",
    "C1077-O9579","C1078-O9562","C1080-O9560","C1081-O9582","C1081-O9583",
    "C1081-O9591","C1086-O9569","C1086-O9570","C1086-O9575","C1086-O9576",
    "C1096-O9545","C1096-O9547","C1096-O9556","C1096-O9557","C1096-O9558",
    "C1096-O9564","C1107-O9539","C1107-O9541","C1108-O9552",
]
PROVINCIAS_DISTRITOS = {
    "Ayabaca": ["Frias"],
    "Huancabamba": ["Canchaque","Huancabamba","Huarmaca","San Miguel de El Faique"],
    "Morropon": ["Buenos Aires","Chalaco","Chulucanas","La Matanza","Morropon",
                 "Salitral","San Juan de Bigote","Santa Catalina de Mossa",
                 "Santo Domingo","Yamango"],
    "Piura": ["Las Lomas","Tambo Grande"],
    "Sullana": ["Sullana"],
}
PROVINCIAS = list(PROVINCIAS_DISTRITOS.keys())
DISTRITOS_PIURA = [d for ds in PROVINCIAS_DISTRITOS.values() for d in ds]
TIPOS_COBERTURA = ["Arborea","Arbustiva","Herbacea","Mixta"]
VIGOR_COBERTURA = ["Excelente","Bueno","Regular","Deficiente","Muy deficiente"]
CATEGORIAS_PRESUPUESTO = [
    "Mano de obra","Materiales e insumos","Equipos y herramientas",
    "Transporte y logistica","Plantones y semillas","Asistencia tecnica",
    "Supervision y monitoreo","Capacitacion","Gastos administrativos","Otros",
]
FUENTES_FINANCIAMIENTO = [
    "Presupuesto publico","Cooperacion internacional","Canon y sobrecanon",
    "Recursos propios","Donaciones","Otro",
]
ESTADOS_ACTIVIDAD = ["Programado","En ejecucion","Completado","Retrasado","Suspendido"]
ACTIVIDADES_TIPO = [
    "Preparacion de terreno","Produccion de plantones","Plantacion / Revegetacion",
    "Excavacion de zanjas de infiltracion","Construccion de terrazas",
    "Construccion de diques","Mantenimiento y riego","Monitoreo y evaluacion",
    "Capacitacion a comunidades","Supervision tecnica","Elaboracion de informes",
    "Otra actividad",
]
COLORES_ESTADO = {"Pendiente":[231,76,60],"En progreso":[243,156,18],"Verificado":[39,174,96]}

# ── CSS ───────────────────────────────────────────────────────────────────
st.markdown("""<style>
.main-header{background:#2C3E50;padding:1rem 2rem;border-radius:.5rem;margin-bottom:1rem}
.main-header h1{color:#fff!important;margin:0!important;font-size:1.8rem!important}
.main-header p{color:#BDC3C7!important;margin:0!important}
</style>""", unsafe_allow_html=True)

st.markdown("""<div class="main-header">
<h1>\U0001F331 IN Piura</h1>
<p>Plan de Ingreso | Verificacion de Campo | Cuenca Alta del Rio Piura</p>
</div>""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────
pagina = st.sidebar.selectbox("Navegacion", [
    "Panel de Control","Bloques de Intervencion","Inspeccion de Campo",
    "Indicadores de Calidad","Presupuesto","Cronograma",
    "Georreferenciacion","ODK / KoBoToolbox","Reportes",
])
st.sidebar.markdown("---")
st.sidebar.markdown("**IN Piura** v2.0 Web\n\nRestauracion de Ecosistemas\nCuenca Alta del Rio Piura")

# ── Helpers ───────────────────────────────────────────────────────────────
def _bloques_map():
    return {f"{b['codigo']} - {b['tipo_intervencion']}": b["id"] for b in db.obtener_bloques()}

def _distritos(prov):
    return PROVINCIAS_DISTRITOS.get(prov, DISTRITOS_PIURA) if prov else DISTRITOS_PIURA

# ══════════════════════════════════════════════════════════════════════════
# PANEL DE CONTROL
# ══════════════════════════════════════════════════════════════════════════
def pagina_dashboard():
    st.subheader("Panel de Control - Resumen Ejecutivo")
    stats = db.obtener_estadisticas_generales()
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Total Bloques", stats["total_bloques"])
    c2.metric("Area Total", f"{stats['area_total_ha']:.2f} ha")
    c3.metric("Inspecciones", stats["total_inspecciones"])
    c4.metric("Avance Promedio", f"{stats['avance_promedio']:.1f}%")
    c5.metric("Personal Activo", stats["personal_activo"])
    st.markdown("---")
    ci,cd = st.columns(2)
    with ci:
        st.markdown("**Distribucion por Estado**")
        tb = max(stats["total_bloques"],1)
        for e in ["Pendiente","En progreso","Verificado"]:
            n = stats["bloques_por_estado"].get(e,0)
            st.progress(n/tb, text=f"{e}: {n} ({n/tb*100:.1f}%)")
    with cd:
        st.markdown("**Distribucion por Tipo**")
        if stats["bloques_por_tipo"]:
            df = pd.DataFrame(list(stats["bloques_por_tipo"].items()), columns=["Tipo","Cantidad"])
            st.bar_chart(df.set_index("Tipo"))
        else:
            st.info("Sin bloques.")
    st.markdown("---")
    cp,cc = st.columns(2)
    with cp:
        st.markdown("**Resumen Presupuestal**")
        pl,ej = stats["presupuesto_planificado"],stats["presupuesto_ejecutado"]
        pe = (ej/pl*100) if pl>0 else 0
        st.metric("Planificado",f"S/ {pl:,.2f}"); st.metric("Ejecutado",f"S/ {ej:,.2f}")
        st.progress(min(pe/100,1.0), text=f"Ejecucion: {pe:.1f}%")
        st.caption(f"Saldo: S/ {pl-ej:,.2f}")
    with cc:
        st.markdown("**Cronograma**")
        ae = stats["actividades_por_estado"]; ta = sum(ae.values()) if ae else 0
        st.caption(f"Total actividades: {ta}")
        for en in ["Programado","En ejecucion","Completado","Retrasado"]:
            cn = ae.get(en,0); st.progress((cn/ta) if ta>0 else 0, text=f"{en}: {cn}")
    st.markdown("---")
    st.markdown("**Resumen de Bloques**")
    res = db.obtener_resumen_bloques()
    if res:
        st.dataframe(pd.DataFrame([{"Codigo":b["codigo"],"Tipo":b["tipo_intervencion"],
            "Distrito":b["distrito"],"Area (ha)":f"{b['area_hectareas']:.4f}",
            "Estado":b["estado"],"Avance %":f"{(b.get('ultimo_avance') or 0):.1f}",
            "Inspecciones":b.get("total_inspecciones",0)} for b in res]),
            use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════
# BLOQUES DE INTERVENCION
# ══════════════════════════════════════════════════════════════════════════
def pagina_bloques():
    st.subheader("Bloques de Intervencion")
    cf,ct = st.columns([1,2])
    with cf:
        st.markdown("**Registro de Bloque**")
        with st.form("form_bloque", clear_on_submit=False):
            codigo = st.text_input("Codigo de bloque")
            cuenca = st.text_input("Cuenca", value="Cuenca Alta del Rio Piura")
            microcuenca = st.selectbox("Microcuenca", [""]+MICROCUENCAS)
            provincia = st.selectbox("Provincia", [""]+PROVINCIAS)
            distrito = st.selectbox("Distrito", [""]+_distritos(provincia))
            tipo = st.selectbox("Tipo intervencion", TIPOS_INTERVENCION)
            a1,a2 = st.columns(2)
            ue = a1.text_input("UTM Este","0"); un = a2.text_input("UTM Norte","0")
            b1,b2 = st.columns(2)
            uz = b1.text_input("Zona UTM","17S"); alt = b2.text_input("Altitud","0")
            area = st.text_input("Area (ha)","0")
            resp = st.text_input("Responsable")
            estado = st.selectbox("Estado", ESTADOS_BLOQUE)
            guardar = st.form_submit_button("Guardar", type="primary")
        if guardar:
            if not codigo: st.warning("Codigo obligatorio.")
            elif not distrito: st.warning("Seleccione distrito.")
            else:
                try:
                    db.insertar_bloque(codigo=codigo,tipo_intervencion=tipo,cuenca=cuenca,
                        distrito=distrito,utm_este=float(ue),utm_norte=float(un),
                        utm_zona=uz,area_hectareas=float(area),estado=estado,
                        altitud=float(alt or 0),responsable=resp,
                        microcuenca=microcuenca,provincia=provincia)
                    st.success(f"Bloque {codigo} registrado."); st.rerun()
                except Exception as e: st.error(f"Error: {e}")
    with ct:
        st.markdown("**Bloques Registrados**")
        busq = st.text_input("Buscar","",key="busq_bl")
        bloques = db.buscar_bloques(busq) if busq else db.obtener_bloques()
        if bloques:
            st.dataframe(pd.DataFrame([{"ID":b["id"],"Codigo":b["codigo"],
                "Microcuenca":b.get("microcuenca","") or "","Tipo":b["tipo_intervencion"],
                "Provincia":b.get("provincia","") or "","Distrito":b["distrito"],
                "UTM Este":f"{b['utm_este']:.2f}","UTM Norte":f"{b['utm_norte']:.2f}",
                "Altitud":f"{(b.get('altitud',0) or 0):.0f}",
                "Area":f"{b['area_hectareas']:.4f}",
                "Responsable":b.get("responsable","") or "","Estado":b["estado"]
                } for b in bloques]), use_container_width=True, hide_index=True)
            st.markdown("---")
            bm = {f"{b['codigo']} - {b['tipo_intervencion']}":b["id"] for b in bloques}
            sel = st.selectbox("Seleccionar bloque para eliminar",[""]+list(bm.keys()),key="del_bl")
            if sel and sel in bm and st.button("Eliminar bloque"):
                db.eliminar_bloque(bm[sel]); st.success("Eliminado."); st.rerun()
        else: st.info("Sin bloques.")

# ══════════════════════════════════════════════════════════════════════════
# INSPECCION DE CAMPO
# ══════════════════════════════════════════════════════════════════════════
def pagina_inspeccion():
    st.subheader("Inspeccion de Campo")
    bm = _bloques_map()
    if not bm: st.warning("Registre un bloque primero."); return
    with st.form("form_insp", clear_on_submit=True):
        bl = st.selectbox("Bloque", list(bm.keys()))
        mc = st.selectbox("Microcuenca",[""]+MICROCUENCAS)
        fecha = st.date_input("Fecha de visita",value=datetime.now())
        inspector = st.text_input("Inspector")
        clima = st.selectbox("Condiciones climaticas",CONDICIONES_CLIMATICAS)
        avance = st.number_input("Avance fisico (%)",0.0,100.0,0.0)
        obs = st.text_area("Observaciones tecnicas")
        desv = st.text_area("Desviaciones al exp. tecnico")
        ver = st.text_input("Codigo de verificacion",
            value=f"VER-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}")
        guardar = st.form_submit_button("Guardar Inspeccion", type="primary")
    if guardar:
        if not inspector: st.warning("Inspector obligatorio.")
        else:
            try:
                db.insertar_inspeccion(bloque_id=bm[bl],fecha_visita=fecha.strftime("%Y-%m-%d"),
                    inspector=inspector,condiciones_climaticas=clima,avance_fisico=avance,
                    observaciones=obs,desviaciones=desv,registro_fotografico="",
                    codigo_verificacion=ver,microcuenca=mc)
                st.success("Inspeccion registrada."); st.rerun()
            except Exception as e: st.error(f"Error: {e}")
    st.markdown("---")
    st.markdown("**Historial de Inspecciones**")
    insp = db.obtener_todas_inspecciones()
    if insp:
        st.dataframe(pd.DataFrame([{"ID":i["id"],"Bloque":i["bloque_codigo"],
            "Microcuenca":i.get("microcuenca","") or "","Fecha":i["fecha_visita"],
            "Inspector":i["inspector"],"Avance %":f"{i['avance_fisico']:.1f}",
            "Verificacion":i["codigo_verificacion"]} for i in insp]),
            use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════
# INDICADORES DE CALIDAD
# ══════════════════════════════════════════════════════════════════════════
def pagina_indicadores():
    st.subheader("Indicadores de Calidad")
    bm = _bloques_map()
    if not bm: st.warning("Registre un bloque primero."); return
    bl = st.selectbox("Bloque", list(bm.keys()), key="ind_bl")
    bid = bm[bl]
    ins = db.obtener_inspecciones_por_bloque(bid)
    if not ins: st.warning("Sin inspecciones para este bloque."); return
    im = {f"{i['fecha_visita']} - {i['inspector']}":i["id"] for i in ins}
    isel = st.selectbox("Inspeccion", list(im.keys()))
    with st.form("form_ind", clear_on_submit=True):
        mc = st.selectbox("Microcuenca",[""]+MICROCUENCAS,key="ind_mc")
        pc = st.number_input("Cobertura vegetal (%)",0.0,100.0,0.0)
        tc = st.selectbox("Tipo cobertura",[""]+TIPOS_COBERTURA)
        vi = st.selectbox("Vigor cobertura",[""]+VIGOR_COBERTURA)
        so = st.number_input("Sobrevivencia especies (%)",0.0,100.0,0.0)
        lz = st.number_input("Longitud zanjas (ml)",0.0,value=0.0)
        vr = st.number_input("Vol. retencion sedimentos (m3)",0.0,value=0.0)
        guardar = st.form_submit_button("Guardar Indicadores", type="primary")
    if guardar:
        try:
            db.insertar_indicadores(bloque_id=bid,inspeccion_id=im[isel],
                cobertura_vegetal_planificada=0,cobertura_vegetal_lograda=0,
                sobrevivencia_especies=so,longitud_zanjas_ejecutada=lz,
                volumen_retencion_sedimentos=vr,porcentaje_cobertura_vegetal=pc,
                tipo_cobertura_vegetal=tc,vigor_cobertura_vegetal=vi,microcuenca=mc)
            st.success("Indicadores guardados."); st.rerun()
        except Exception as e: st.error(f"Error: {e}")
    st.markdown("---")
    st.markdown("**Indicadores Registrados**")
    ind = db.obtener_indicadores_por_bloque(bid)
    if ind:
        st.dataframe(pd.DataFrame([{"Fecha":x.get("fecha_visita",""),
            "Microcuenca":x.get("microcuenca","") or "",
            "Cobert.%":f"{x.get('porcentaje_cobertura_vegetal',0):.1f}",
            "Tipo":x.get("tipo_cobertura_vegetal","") or "",
            "Vigor":x.get("vigor_cobertura_vegetal","") or "",
            "Sobrev.%":f"{x['sobrevivencia_especies']:.1f}",
            "Zanjas":f"{x['longitud_zanjas_ejecutada']:.2f}",
            "Vol.Ret.":f"{x['volumen_retencion_sedimentos']:.2f}"} for x in ind]),
            use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════
# PRESUPUESTO
# ══════════════════════════════════════════════════════════════════════════
def pagina_presupuesto():
    st.subheader("Presupuesto y Recursos")
    bm = _bloques_map()
    if not bm: st.warning("Registre un bloque primero."); return
    bl = st.selectbox("Bloque", list(bm.keys()), key="pres_bl")
    bid = bm[bl]
    with st.form("form_pres", clear_on_submit=True):
        cat = st.selectbox("Categoria",CATEGORIAS_PRESUPUESTO)
        desc = st.text_input("Descripcion")
        x1,x2 = st.columns(2)
        mp = x1.number_input("Monto planificado (S/)",0.0,value=0.0,format="%.2f")
        me = x2.number_input("Monto ejecutado (S/)",0.0,value=0.0,format="%.2f")
        fu = st.selectbox("Fuente financiamiento",FUENTES_FINANCIAMIENTO)
        guardar = st.form_submit_button("Guardar Partida", type="primary")
    if guardar:
        try:
            db.insertar_presupuesto(bid,cat,desc,mp,me,fu)
            st.success("Partida registrada."); st.rerun()
        except Exception as e: st.error(f"Error: {e}")
    st.markdown("---")
    st.markdown("**Partidas del Bloque**")
    pa = db.obtener_presupuesto_por_bloque(bid)
    if pa:
        tp = sum(p["monto_planificado"] for p in pa)
        te = sum(p["monto_ejecutado"] for p in pa)
        st.dataframe(pd.DataFrame([{"ID":p["id"],"Categoria":p["categoria"],
            "Descripcion":p["descripcion"],
            "Planificado":f"S/ {p['monto_planificado']:,.2f}",
            "Ejecutado":f"S/ {p['monto_ejecutado']:,.2f}",
            "%Ejec":f"{(p['monto_ejecutado']/p['monto_planificado']*100) if p['monto_planificado']>0 else 0:.1f}%",
            "Fuente":p["fuente_financiamiento"]} for p in pa]),
            use_container_width=True, hide_index=True)
        st.info(f"**Subtotal:** Plan S/ {tp:,.2f} | Ejec S/ {te:,.2f} | {(te/tp*100) if tp>0 else 0:.1f}%")
        pm = {f"ID {p['id']} - {p['categoria']}":p["id"] for p in pa}
        sp = st.selectbox("Partida a eliminar",[""]+list(pm.keys()),key="del_pa")
        if sp and sp in pm and st.button("Eliminar partida"):
            db.eliminar_presupuesto(pm[sp]); st.success("Eliminada."); st.rerun()
    st.markdown("---")
    st.markdown("**Resumen General**")
    rp = db.obtener_resumen_presupuesto()
    if rp:
        st.dataframe(pd.DataFrame([{"Codigo":r["codigo"],"Tipo":r["tipo_intervencion"],
            "Distrito":r["distrito"],
            "Planificado":f"S/ {r['total_planificado']:,.2f}",
            "Ejecutado":f"S/ {r['total_ejecutado']:,.2f}",
            "%Ejec":f"{(r['total_ejecutado']/r['total_planificado']*100) if r['total_planificado']>0 else 0:.1f}%",
            "Partidas":r["num_partidas"]} for r in rp]),
            use_container_width=True, hide_index=True)
        t = db.obtener_presupuesto_total()
        pt,et = t["total_planificado"],t["total_ejecutado"]
        st.success(f"**TOTAL:** Plan S/ {pt:,.2f} | Ejec S/ {et:,.2f} ({(et/pt*100) if pt>0 else 0:.1f}%) | Saldo S/ {pt-et:,.2f}")

# ══════════════════════════════════════════════════════════════════════════
# CRONOGRAMA
# ══════════════════════════════════════════════════════════════════════════
def pagina_cronograma():
    st.subheader("Cronograma de Actividades")
    bm = _bloques_map()
    if not bm: st.warning("Registre un bloque primero."); return
    bl = st.selectbox("Bloque", list(bm.keys()), key="crono_bl")
    bid = bm[bl]
    with st.form("form_crono", clear_on_submit=True):
        act = st.selectbox("Actividad",ACTIVIDADES_TIPO)
        x1,x2 = st.columns(2)
        ip = x1.date_input("Inicio plan.",value=datetime.now())
        fp = x2.date_input("Fin plan.",value=datetime.now())
        x3,x4 = st.columns(2)
        ir = x3.text_input("Inicio real",""); fr = x4.text_input("Fin real","")
        x5,x6 = st.columns(2)
        av = x5.number_input("Avance %",0.0,100.0,0.0)
        ea = x6.selectbox("Estado",ESTADOS_ACTIVIDAD)
        re = st.text_input("Responsable"); ob = st.text_area("Observaciones")
        guardar = st.form_submit_button("Guardar Actividad", type="primary")
    if guardar:
        try:
            db.insertar_actividad(bloque_id=bid,actividad=act,
                fecha_inicio_plan=ip.strftime("%Y-%m-%d"),fecha_fin_plan=fp.strftime("%Y-%m-%d"),
                fecha_inicio_real=ir,fecha_fin_real=fr,porcentaje_avance=av,
                responsable=re,observaciones=ob,estado=ea)
            st.success("Actividad registrada."); st.rerun()
        except Exception as e: st.error(f"Error: {e}")
    st.markdown("---")
    st.markdown("**Actividades del Bloque**")
    acs = db.obtener_actividades_por_bloque(bid)
    if acs:
        st.dataframe(pd.DataFrame([{"ID":a["id"],"Actividad":a["actividad"],
            "Inicio":a["fecha_inicio_plan"],"Fin":a["fecha_fin_plan"],
            "Inicio Real":a["fecha_inicio_real"] or "-","Fin Real":a["fecha_fin_real"] or "-",
            "Avance":f"{a['porcentaje_avance']:.0f}%","Estado":a["estado"],
            "Responsable":a["responsable"]} for a in acs]),
            use_container_width=True, hide_index=True)
        am = {f"ID {a['id']} - {a['actividad']}":a["id"] for a in acs}
        sa = st.selectbox("Actividad a eliminar",[""]+list(am.keys()),key="del_ac")
        if sa and sa in am and st.button("Eliminar actividad"):
            db.eliminar_actividad(am[sa]); st.success("Eliminada."); st.rerun()
    st.markdown("---")
    st.markdown("**Cronograma General**")
    fe = st.selectbox("Filtrar estado",["Todos"]+ESTADOS_ACTIVIDAD,key="f_crono")
    ta = db.obtener_todas_actividades()
    if ta:
        if fe != "Todos": ta = [a for a in ta if a.get("estado")==fe]
        st.dataframe(pd.DataFrame([{"Bloque":a.get("bloque_codigo",""),
            "Actividad":a["actividad"],"Inicio":a["fecha_inicio_plan"],
            "Fin":a["fecha_fin_plan"],"Avance":f"{a['porcentaje_avance']:.0f}%",
            "Estado":a["estado"],"Responsable":a["responsable"]} for a in ta]),
            use_container_width=True, hide_index=True)
        rc = db.obtener_resumen_cronograma(); tt = sum(rc.values())
        co = rc.get("Completado",0)
        st.info(f"Total: {tt} | Programadas: {rc.get('Programado',0)} | En ejecucion: {rc.get('En ejecucion',0)} | Completadas: {co} ({(co/tt*100) if tt>0 else 0:.0f}%) | Retrasadas: {rc.get('Retrasado',0)}")

# ══════════════════════════════════════════════════════════════════════════
# GEORREFERENCIACION
# ══════════════════════════════════════════════════════════════════════════
def pagina_georreferenciacion():
    st.subheader("Georreferenciacion")
    bloques = db.obtener_bloques()
    f1,f2 = st.columns(2)
    fe = f1.selectbox("Filtrar estado",["Todos","Pendiente","En progreso","Verificado"],key="ge")
    ft = f2.selectbox("Filtrar tipo",["Todos"]+TIPOS_INTERVENCION,key="gt")
    bf = bloques
    if fe != "Todos": bf = [b for b in bf if b.get("estado")==fe]
    if ft != "Todos": bf = [b for b in bf if b.get("tipo_intervencion")==ft]
    cm,ci = st.columns([3,1])
    with cm:
        if bf:
            md = []
            for b in bf:
                try:
                    zn = int(b["utm_zona"].replace("S","").replace("N",""))
                    he = "S" if "S" in b["utm_zona"] else "N"
                    la,lo = utm_a_latlon(b["utm_este"],b["utm_norte"],zn,he)
                    co = COLORES_ESTADO.get(b.get("estado",""),[149,165,166])
                    md.append({"lat":la,"lon":lo,"codigo":b["codigo"],
                        "tipo":b["tipo_intervencion"],"estado":b["estado"],
                        "area":b["area_hectareas"],"r":co[0],"g":co[1],"b":co[2]})
                except (ValueError,KeyError): pass
            if md:
                import pydeck as pdk
                df = pd.DataFrame(md)
                layer = pdk.Layer("ScatterplotLayer",data=df,
                    get_position=["lon","lat"],get_color=["r","g","b",200],
                    get_radius=300,pickable=True)
                vs = pdk.ViewState(latitude=df["lat"].mean(),longitude=df["lon"].mean(),zoom=10)
                st.pydeck_chart(pdk.Deck(layers=[layer],initial_view_state=vs,
                    tooltip={"text":"Codigo: {codigo}\nTipo: {tipo}\nEstado: {estado}\nArea: {area} ha"}))
                st.caption(":red_circle: Pendiente | :orange_circle: En progreso | :green_circle: Verificado")
            else: st.info("No se pudieron convertir coordenadas.")
        else: st.info("Sin bloques para los filtros.")
    with ci:
        st.markdown("**Resumen**")
        st.metric("Bloques",len(bf))
        st.metric("Area",f"{sum(b.get('area_hectareas',0) for b in bf):.4f} ha")
        st.markdown("---")
        st.markdown("**Conversor UTM <-> Lat/Lon**")
        t1,t2 = st.tabs(["UTM->LatLon","LatLon->UTM"])
        with t1:
            with st.form("conv_u"):
                ce = st.text_input("Este",key="ce"); cn = st.text_input("Norte",key="cn")
                cz = st.text_input("Zona","17S",key="cz")
                if st.form_submit_button("Convertir") and ce and cn:
                    try:
                        zn = int(cz.replace("S","").replace("N",""))
                        he = "S" if "S" in cz.upper() else "N"
                        la,lo = utm_a_latlon(float(ce),float(cn),zn,he)
                        st.success(f"Lat: {la:.8f}\nLon: {lo:.8f}")
                    except: st.error("Valores invalidos.")
        with t2:
            with st.form("conv_l"):
                cl = st.text_input("Latitud",key="cl"); co = st.text_input("Longitud",key="co")
                if st.form_submit_button("Convertir") and cl and co:
                    try:
                        e,n,z = latlon_a_utm(float(cl),float(co))
                        st.success(f"Este: {e:.2f}\nNorte: {n:.2f}\nZona: {z}")
                    except: st.error("Valores invalidos.")

# ══════════════════════════════════════════════════════════════════════════
# ODK / KoBoToolbox
# ══════════════════════════════════════════════════════════════════════════
def pagina_odk():
    st.subheader("ODK / KoBoToolbox")
    st.markdown("### Generar Formulario XLSForm")
    if st.button("Generar Formulario XLSForm", type="primary"):
        try:
            ruta = generar_xlsform()
            with open(ruta,"rb") as f: data = f.read()
            st.download_button("Descargar XLSForm",data,os.path.basename(ruta),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.success("Formulario generado.")
        except Exception as e: st.error(f"Error: {e}")
    st.markdown("---")
    st.markdown("### Importar CSV")
    uf = st.file_uploader("Archivo CSV",type=["csv"])
    if uf:
        with tempfile.NamedTemporaryFile(mode="w",suffix=".csv",delete=False,encoding="utf-8") as tmp:
            tmp.write(uf.read().decode("utf-8-sig")); tp = tmp.name
        try:
            r = importar_csv_odk(tp)
            st.success(f"**{r['total_filas']} registros** | Nuevos: {r['bloques_nuevos']} | Actualizados: {r['bloques_actualizados']} | Inspecciones: {r['inspecciones_creadas']} | Indicadores: {r['indicadores_creados']}")
            if r["errores"]:
                with st.expander(f"{len(r['errores'])} errores"):
                    for e in r["errores"]: st.warning(e)
        except Exception as e: st.error(f"Error: {e}")
        finally: os.unlink(tp)
    enc = ["codigo_bloque","tipo_intervencion","cuenca","distrito","utm_este","utm_norte",
        "utm_zona","area_hectareas","estado","ubicacion_gps","fecha_visita","inspector",
        "condiciones_climaticas","avance_fisico","observaciones","desviaciones",
        "foto_1","foto_2","foto_3","cobertura_vegetal_planificada",
        "cobertura_vegetal_lograda","sobrevivencia_especies","longitud_zanjas",
        "volumen_retencion","codigo_verificacion"]
    buf = io.StringIO(); w = csv.writer(buf); w.writerow(enc)
    w.writerow(["BLQ-001","revegetacion","Cuenca Alta del Rio Piura","Canchaque",
        "622150.50","9436720.30","17S","2.5","pendiente","","2026-01-15","Juan Perez",
        "despejado","45","","","","","","1100","850","78.5","120.5","35.2",""])
    st.download_button("Descargar plantilla CSV",buf.getvalue(),"plantilla_odk.csv","text/csv")
    st.markdown("---")
    st.markdown("### API KoBoToolbox")
    sv = st.selectbox("Servidor",["https://kf.kobotoolbox.org","https://kobo.humanitarianresponse.info"])
    tk = st.text_input("Token API",type="password")
    c1,c2,c3 = st.columns(3)
    if c1.button("Probar Conexion"):
        if not tk: st.warning("Ingrese token.")
        else:
            cl = KoBoClient(sv,tk); ok,msg = cl.test_conexion()
            if ok: st.success("Conexion exitosa"); st.session_state["kobo"]=cl
            else: st.error(msg)
    if c2.button("Listar Formularios"):
        if "kobo" not in st.session_state: st.warning("Pruebe la conexion primero.")
        else:
            try:
                fs = st.session_state["kobo"].listar_formularios()
                if fs: st.dataframe(pd.DataFrame([{"UID":f["uid"],"Nombre":f["nombre"],
                    "Envios":f["envios"],"Estado":"Desplegado" if f["desplegado"] else "Borrador"} for f in fs]),
                    use_container_width=True, hide_index=True)
            except Exception as e: st.error(f"Error: {e}")
    uid = st.text_input("UID del formulario")
    if c3.button("Importar Envios"):
        if not uid or not tk: st.warning("Complete los campos.")
        else:
            try:
                r = importar_desde_kobo(sv,tk,uid)
                st.success(f"Importado: {r['total_filas']} envios | Nuevos: {r['bloques_nuevos']} | Inspecciones: {r['inspecciones_creadas']}")
            except Exception as e: st.error(f"Error: {e}")
    with st.expander("Guia Rapida"):
        st.markdown("""
1. **GENERAR** el formulario XLSForm
2. **SUBIR** a KoBoToolbox o ODK Central
3. **DESPLEGAR** el formulario
4. **RECOLECTAR** datos con KoBoCollect/ODK Collect (sin internet)
5. **SINCRONIZAR** al tener conexion
6. **IMPORTAR** CSV o usar API directa""")

# ══════════════════════════════════════════════════════════════════════════
# REPORTES
# ══════════════════════════════════════════════════════════════════════════
def pagina_reportes():
    st.subheader("Generacion de Reportes")
    bm = _bloques_map()
    st.markdown("### Ficha de Inspeccion (PDF)")
    if bm:
        bl = st.selectbox("Bloque",list(bm.keys()),key="rep_bl")
        if st.button("Generar Ficha PDF", type="primary"):
            try:
                ruta = reports.generar_ficha_pdf(bm[bl])
                with open(ruta,"rb") as f: data = f.read()
                st.download_button("Descargar PDF",data,os.path.basename(ruta),"application/pdf")
                st.success("PDF generado.")
            except Exception as e: st.error(f"Error: {e}")
    else: st.info("Registre bloques primero.")
    st.markdown("---")
    st.markdown("### Tabla Resumen (Excel)")
    if st.button("Generar Resumen Excel"):
        try:
            ruta = reports.generar_resumen_excel()
            with open(ruta,"rb") as f: data = f.read()
            st.download_button("Descargar Excel",data,os.path.basename(ruta),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.success("Excel generado.")
        except Exception as e: st.error(f"Error: {e}")
    st.markdown("---")
    st.markdown("### Tabla UTM para ArcGIS (Excel)")
    if st.button("Generar Excel ArcGIS"):
        try:
            ruta = reports.generar_excel_arcgis()
            with open(ruta,"rb") as f: data = f.read()
            st.download_button("Descargar Excel ArcGIS",data,os.path.basename(ruta),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.success("Archivo generado.")
        except Exception as e: st.error(f"Error: {e}")

# ══════════════════════════════════════════════════════════════════════════
# ROUTER
# ══════════════════════════════════════════════════════════════════════════
if pagina == "Panel de Control": pagina_dashboard()
elif pagina == "Bloques de Intervencion": pagina_bloques()
elif pagina == "Inspeccion de Campo": pagina_inspeccion()
elif pagina == "Indicadores de Calidad": pagina_indicadores()
elif pagina == "Presupuesto": pagina_presupuesto()
elif pagina == "Cronograma": pagina_cronograma()
elif pagina == "Georreferenciacion": pagina_georreferenciacion()
elif pagina == "ODK / KoBoToolbox": pagina_odk()
elif pagina == "Reportes": pagina_reportes()
