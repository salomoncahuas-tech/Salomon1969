"""
IN Piura - Panel de Control (Dashboard)
Vista general del proyecto con estadísticas, indicadores de avance
y resumen ejecutivo del Plan de Ingreso.
Cuenca Alta del Río Piura, Perú.
"""

import tkinter as tk
from tkinter import ttk

import database as db

COLOR_PRIMARIO = "#2C3E50"
COLOR_SECUNDARIO = "#3498DB"
COLOR_FONDO = "#ECF0F1"
COLOR_EXITO = "#27AE60"
COLOR_ALERTA = "#E74C3C"
COLOR_ADVERTENCIA = "#F39C12"


class BarraProgreso(tk.Canvas):
    """Barra de progreso personalizada con colores y etiqueta."""

    def __init__(self, parent, ancho=260, alto=22, **kwargs):
        super().__init__(parent, width=ancho, height=alto,
                         bg=COLOR_FONDO, highlightthickness=0, **kwargs)
        self.ancho = ancho
        self.alto = alto

    def set_valor(self, porcentaje, etiqueta="", color=None):
        self.delete("all")
        porcentaje = max(0, min(100, porcentaje))

        if color is None:
            if porcentaje >= 75:
                color = COLOR_EXITO
            elif porcentaje >= 40:
                color = COLOR_ADVERTENCIA
            else:
                color = COLOR_ALERTA

        # Fondo
        self.create_rectangle(0, 0, self.ancho, self.alto,
                              fill="#D5D8DC", outline="#BDC3C7")
        # Barra
        ancho_barra = (porcentaje / 100) * self.ancho
        if ancho_barra > 0:
            self.create_rectangle(0, 0, ancho_barra, self.alto,
                                  fill=color, outline="")
        # Texto
        texto = f"{etiqueta} {porcentaje:.1f}%" if etiqueta else f"{porcentaje:.1f}%"
        self.create_text(self.ancho / 2, self.alto / 2, text=texto,
                         font=("Segoe UI", 8, "bold"), fill="white" if porcentaje > 30 else COLOR_PRIMARIO)


class TarjetaMetrica(ttk.Frame):
    """Tarjeta con valor numérico grande y etiqueta descriptiva."""

    def __init__(self, parent, titulo, valor, subtitulo="", color=COLOR_SECUNDARIO):
        super().__init__(parent, padding=8)

        # Fondo coloreado via tk.Frame interno
        inner = tk.Frame(self, bg="white", padx=14, pady=10,
                         highlightbackground="#BDC3C7", highlightthickness=1)
        inner.pack(fill="both", expand=True)

        barra_color = tk.Frame(inner, bg=color, height=3)
        barra_color.pack(fill="x", pady=(0, 6))

        tk.Label(inner, text=titulo, font=("Segoe UI", 9),
                 bg="white", fg="#7F8C8D").pack(anchor="w")
        tk.Label(inner, text=str(valor), font=("Segoe UI", 20, "bold"),
                 bg="white", fg=COLOR_PRIMARIO).pack(anchor="w")
        if subtitulo:
            tk.Label(inner, text=subtitulo, font=("Segoe UI", 8),
                     bg="white", fg="#95A5A6").pack(anchor="w")


class TabDashboard(ttk.Frame):
    """Pestaña de Panel de Control con resumen ejecutivo del proyecto."""

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self._crear_widgets()
        self.actualizar_datos()

    def _crear_widgets(self):
        # Canvas con scroll
        canvas = tk.Canvas(self, bg=COLOR_FONDO, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas)
        self.scroll_frame.bind("<Configure>",
                               lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        frame = self.scroll_frame

        # ── Encabezado ──
        header = ttk.Frame(frame)
        header.pack(fill="x", padx=12, pady=(12, 4))
        ttk.Label(header, text="Panel de Control - Plan de Ingreso IN Piura",
                  style="Header.TLabel").pack(side="left")
        ttk.Button(header, text="Actualizar",
                   command=self.actualizar_datos).pack(side="right")

        ttk.Label(frame,
                  text="Restauración de Ecosistemas - Cuenca Alta del Río Piura",
                  foreground="#7F8C8D").pack(anchor="w", padx=12, pady=(0, 10))

        # ── Fila 1: Tarjetas de métricas principales ──
        self.frame_tarjetas = ttk.Frame(frame)
        self.frame_tarjetas.pack(fill="x", padx=8, pady=4)

        # ── Fila 2: Progreso por estado y tipo ──
        fila2 = ttk.Frame(frame)
        fila2.pack(fill="x", padx=8, pady=4)

        # Panel izquierdo: Distribución por estado
        self.frame_estados = ttk.LabelFrame(fila2, text=" Distribución por Estado ", padding=10)
        self.frame_estados.pack(side="left", fill="both", expand=True, padx=4)

        # Panel derecho: Distribución por tipo
        self.frame_tipos = ttk.LabelFrame(fila2, text=" Distribución por Tipo de Intervención ", padding=10)
        self.frame_tipos.pack(side="left", fill="both", expand=True, padx=4)

        # ── Fila 3: Presupuesto y Cronograma ──
        fila3 = ttk.Frame(frame)
        fila3.pack(fill="x", padx=8, pady=4)

        self.frame_presupuesto = ttk.LabelFrame(fila3, text=" Resumen Presupuestal ", padding=10)
        self.frame_presupuesto.pack(side="left", fill="both", expand=True, padx=4)

        self.frame_cronograma = ttk.LabelFrame(fila3, text=" Estado del Cronograma ", padding=10)
        self.frame_cronograma.pack(side="left", fill="both", expand=True, padx=4)

        # ── Fila 4: Tabla resumen de bloques ──
        self.frame_tabla = ttk.LabelFrame(frame, text=" Resumen de Bloques de Intervención ", padding=10)
        self.frame_tabla.pack(fill="both", expand=True, padx=12, pady=(4, 12))

        columnas = ("codigo", "tipo", "distrito", "area", "estado", "avance", "inspecciones")
        self.tree_resumen = ttk.Treeview(self.frame_tabla, columns=columnas,
                                         show="headings", height=10)
        self.tree_resumen.heading("codigo", text="Código")
        self.tree_resumen.heading("tipo", text="Tipo Intervención")
        self.tree_resumen.heading("distrito", text="Distrito")
        self.tree_resumen.heading("area", text="Área (ha)")
        self.tree_resumen.heading("estado", text="Estado")
        self.tree_resumen.heading("avance", text="Avance %")
        self.tree_resumen.heading("inspecciones", text="Inspecciones")

        self.tree_resumen.column("codigo", width=90)
        self.tree_resumen.column("tipo", width=180)
        self.tree_resumen.column("distrito", width=110)
        self.tree_resumen.column("area", width=80)
        self.tree_resumen.column("estado", width=100)
        self.tree_resumen.column("avance", width=80)
        self.tree_resumen.column("inspecciones", width=90)

        scroll_tabla = ttk.Scrollbar(self.frame_tabla, orient="vertical",
                                     command=self.tree_resumen.yview)
        self.tree_resumen.configure(yscrollcommand=scroll_tabla.set)
        self.tree_resumen.pack(side="left", fill="both", expand=True)
        scroll_tabla.pack(side="right", fill="y")

    def actualizar_datos(self):
        """Recarga todos los datos del dashboard."""
        stats = db.obtener_estadisticas_generales()

        # ── Limpiar y reconstruir tarjetas ──
        for w in self.frame_tarjetas.winfo_children():
            w.destroy()

        tarjetas_data = [
            ("Total Bloques", stats["total_bloques"], "bloques registrados", COLOR_SECUNDARIO),
            ("Área Total", f"{stats['area_total_ha']:.2f}", "hectáreas", COLOR_EXITO),
            ("Inspecciones", stats["total_inspecciones"], "visitas realizadas", COLOR_ADVERTENCIA),
            ("Avance Promedio", f"{stats['avance_promedio']:.1f}%", "progreso físico", COLOR_PRIMARIO),
            ("Personal Activo", stats["personal_activo"], "técnicos en campo", COLOR_SECUNDARIO),
        ]

        for titulo, valor, sub, color in tarjetas_data:
            tarjeta = TarjetaMetrica(self.frame_tarjetas, titulo, valor, sub, color)
            tarjeta.pack(side="left", fill="both", expand=True)

        # ── Distribución por estado ──
        for w in self.frame_estados.winfo_children():
            w.destroy()

        total_bloques = max(stats["total_bloques"], 1)
        colores_estado = {
            "Pendiente": COLOR_ALERTA,
            "En progreso": COLOR_ADVERTENCIA,
            "Verificado": COLOR_EXITO,
        }

        for estado in ["Pendiente", "En progreso", "Verificado"]:
            cantidad = stats["bloques_por_estado"].get(estado, 0)
            pct = (cantidad / total_bloques) * 100

            row = ttk.Frame(self.frame_estados)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=f"{estado} ({cantidad})",
                      width=22).pack(side="left")
            barra = BarraProgreso(row, ancho=220, alto=18)
            barra.pack(side="left", padx=4)
            barra.set_valor(pct, color=colores_estado.get(estado, COLOR_SECUNDARIO))

        # ── Distribución por tipo ──
        for w in self.frame_tipos.winfo_children():
            w.destroy()

        colores_tipo = {
            "Revegetación": "#27AE60",
            "Zanjas de infiltración": "#3498DB",
            "Terrazas de formación lenta": "#E67E22",
            "Diques de mampostería": "#9B59B6",
        }

        for tipo, cantidad in stats["bloques_por_tipo"].items():
            pct = (cantidad / total_bloques) * 100
            row = ttk.Frame(self.frame_tipos)
            row.pack(fill="x", pady=2)
            nombre_corto = tipo if len(tipo) <= 28 else tipo[:25] + "..."
            ttk.Label(row, text=f"{nombre_corto} ({cantidad})",
                      width=30).pack(side="left")
            barra = BarraProgreso(row, ancho=180, alto=18)
            barra.pack(side="left", padx=4)
            barra.set_valor(pct, color=colores_tipo.get(tipo, COLOR_SECUNDARIO))

        # ── Resumen presupuestal ──
        for w in self.frame_presupuesto.winfo_children():
            w.destroy()

        plan = stats["presupuesto_planificado"]
        ejec = stats["presupuesto_ejecutado"]
        pct_ejec = (ejec / plan * 100) if plan > 0 else 0

        tk.Label(self.frame_presupuesto, text="Presupuesto Planificado:",
                 font=("Segoe UI", 9), bg=COLOR_FONDO).pack(anchor="w")
        tk.Label(self.frame_presupuesto, text=f"S/ {plan:,.2f}",
                 font=("Segoe UI", 14, "bold"), bg=COLOR_FONDO,
                 fg=COLOR_PRIMARIO).pack(anchor="w")

        tk.Label(self.frame_presupuesto, text="Presupuesto Ejecutado:",
                 font=("Segoe UI", 9), bg=COLOR_FONDO).pack(anchor="w", pady=(6, 0))
        tk.Label(self.frame_presupuesto, text=f"S/ {ejec:,.2f}",
                 font=("Segoe UI", 14, "bold"), bg=COLOR_FONDO,
                 fg=COLOR_EXITO if pct_ejec <= 100 else COLOR_ALERTA).pack(anchor="w")

        ttk.Label(self.frame_presupuesto, text="Ejecución:").pack(anchor="w", pady=(6, 2))
        barra_pres = BarraProgreso(self.frame_presupuesto, ancho=280, alto=20)
        barra_pres.pack(anchor="w")
        barra_pres.set_valor(pct_ejec)

        saldo = plan - ejec
        tk.Label(self.frame_presupuesto, text=f"Saldo disponible: S/ {saldo:,.2f}",
                 font=("Segoe UI", 9), bg=COLOR_FONDO,
                 fg=COLOR_EXITO if saldo >= 0 else COLOR_ALERTA).pack(anchor="w", pady=(6, 0))

        # ── Estado del cronograma ──
        for w in self.frame_cronograma.winfo_children():
            w.destroy()

        act_estados = stats["actividades_por_estado"]
        total_act = sum(act_estados.values()) if act_estados else 0

        estados_crono = [
            ("Programado", "#3498DB"),
            ("En ejecución", COLOR_ADVERTENCIA),
            ("Completado", COLOR_EXITO),
            ("Retrasado", COLOR_ALERTA),
        ]

        tk.Label(self.frame_cronograma, text=f"Total actividades: {total_act}",
                 font=("Segoe UI", 10, "bold"), bg=COLOR_FONDO,
                 fg=COLOR_PRIMARIO).pack(anchor="w", pady=(0, 6))

        for est_nombre, est_color in estados_crono:
            cant = act_estados.get(est_nombre, 0)
            pct = (cant / total_act * 100) if total_act > 0 else 0

            row = ttk.Frame(self.frame_cronograma)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=f"{est_nombre} ({cant})", width=20).pack(side="left")
            barra = BarraProgreso(row, ancho=200, alto=16)
            barra.pack(side="left", padx=4)
            barra.set_valor(pct, color=est_color)

        # ── Tabla resumen de bloques ──
        for item in self.tree_resumen.get_children():
            self.tree_resumen.delete(item)

        resumen = db.obtener_resumen_bloques()
        for b in resumen:
            avance = b.get("ultimo_avance") or 0
            total_insp = b.get("total_inspecciones", 0)
            self.tree_resumen.insert("", "end", values=(
                b["codigo"],
                b["tipo_intervencion"],
                b["distrito"],
                f"{b['area_hectareas']:.4f}",
                b["estado"],
                f"{avance:.1f}",
                total_insp,
            ))
