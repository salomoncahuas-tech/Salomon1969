"""
IN Piura - Módulo de Presupuesto y Recursos
Gestión del presupuesto planificado vs ejecutado por bloque de intervención,
categorías de gasto y fuentes de financiamiento.
Cuenca Alta del Río Piura, Perú.
"""

import tkinter as tk
from tkinter import ttk, messagebox

import database as db

COLOR_PRIMARIO = "#2C3E50"
COLOR_SECUNDARIO = "#3498DB"
COLOR_FONDO = "#ECF0F1"
COLOR_EXITO = "#27AE60"
COLOR_ALERTA = "#E74C3C"

CATEGORIAS_PRESUPUESTO = [
    "Mano de obra",
    "Materiales e insumos",
    "Equipos y herramientas",
    "Transporte y logística",
    "Plantones y semillas",
    "Asistencia técnica",
    "Supervisión y monitoreo",
    "Capacitación",
    "Gastos administrativos",
    "Otros",
]

FUENTES_FINANCIAMIENTO = [
    "Presupuesto público",
    "Cooperación internacional",
    "Canon y sobrecanon",
    "Recursos propios",
    "Donaciones",
    "Otro",
]


class TabPresupuesto(ttk.Frame):
    """Pestaña de gestión presupuestal del proyecto."""

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.partida_seleccionada = None
        self._crear_widgets()
        self.cargar_combos()

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
        frame.columnconfigure(1, weight=1)

        # ── Encabezado ──
        ttk.Label(frame, text="Presupuesto y Recursos del Proyecto",
                  style="Header.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(12, 4))

        ttk.Label(frame,
                  text="Registro y seguimiento del presupuesto planificado vs. ejecutado por bloque.",
                  foreground="#7F8C8D").grid(
            row=1, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 10))

        # ── Formulario de partida presupuestal ──
        form_frame = ttk.LabelFrame(frame, text=" Registrar Partida Presupuestal ", padding=12)
        form_frame.grid(row=2, column=0, columnspan=3, sticky="ew", padx=12, pady=6)
        form_frame.columnconfigure(1, weight=1)

        # Bloque
        ttk.Label(form_frame, text="Bloque:").grid(row=0, column=0, sticky="w", pady=4)
        self.combo_bloque = ttk.Combobox(form_frame, state="readonly", width=40)
        self.combo_bloque.grid(row=0, column=1, sticky="w", padx=6, pady=4)
        self.combo_bloque.bind("<<ComboboxSelected>>", lambda e: self.cargar_partidas())

        # Categoría
        ttk.Label(form_frame, text="Categoría:").grid(row=1, column=0, sticky="w", pady=4)
        self.combo_categoria = ttk.Combobox(form_frame, values=CATEGORIAS_PRESUPUESTO,
                                            state="readonly", width=30)
        self.combo_categoria.grid(row=1, column=1, sticky="w", padx=6, pady=4)
        self.combo_categoria.set(CATEGORIAS_PRESUPUESTO[0])

        # Descripción
        ttk.Label(form_frame, text="Descripción:").grid(row=2, column=0, sticky="w", pady=4)
        self.entry_descripcion = ttk.Entry(form_frame, width=50)
        self.entry_descripcion.grid(row=2, column=1, sticky="w", padx=6, pady=4)

        # Monto planificado
        ttk.Label(form_frame, text="Monto planificado (S/):").grid(row=3, column=0, sticky="w", pady=4)
        self.entry_monto_plan = ttk.Entry(form_frame, width=18)
        self.entry_monto_plan.grid(row=3, column=1, sticky="w", padx=6, pady=4)
        self.entry_monto_plan.insert(0, "0.00")

        # Monto ejecutado
        ttk.Label(form_frame, text="Monto ejecutado (S/):").grid(row=4, column=0, sticky="w", pady=4)
        self.entry_monto_ejec = ttk.Entry(form_frame, width=18)
        self.entry_monto_ejec.grid(row=4, column=1, sticky="w", padx=6, pady=4)
        self.entry_monto_ejec.insert(0, "0.00")

        # Fuente de financiamiento
        ttk.Label(form_frame, text="Fuente financiamiento:").grid(row=5, column=0, sticky="w", pady=4)
        self.combo_fuente = ttk.Combobox(form_frame, values=FUENTES_FINANCIAMIENTO,
                                         state="readonly", width=30)
        self.combo_fuente.grid(row=5, column=1, sticky="w", padx=6, pady=4)
        self.combo_fuente.set(FUENTES_FINANCIAMIENTO[0])

        # Botones
        btn_frame = ttk.Frame(form_frame)
        btn_frame.grid(row=6, column=0, columnspan=2, pady=10)

        ttk.Button(btn_frame, text="Guardar Partida",
                   command=self.guardar_partida,
                   style="Accent.TButton").pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Actualizar",
                   command=self.actualizar_partida).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Eliminar",
                   command=self.eliminar_partida,
                   style="Danger.TButton").pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Limpiar",
                   command=self.limpiar_formulario).pack(side="left", padx=3)

        # ── Tabla de partidas ──
        tabla_frame = ttk.LabelFrame(frame, text=" Partidas Presupuestales por Bloque ", padding=10)
        tabla_frame.grid(row=3, column=0, columnspan=3, sticky="ew", padx=12, pady=6)

        columnas = ("id", "categoria", "descripcion", "planificado", "ejecutado", "pct", "fuente")
        self.tree_partidas = ttk.Treeview(tabla_frame, columns=columnas,
                                          show="headings", height=8)
        self.tree_partidas.heading("id", text="ID")
        self.tree_partidas.heading("categoria", text="Categoría")
        self.tree_partidas.heading("descripcion", text="Descripción")
        self.tree_partidas.heading("planificado", text="Planificado (S/)")
        self.tree_partidas.heading("ejecutado", text="Ejecutado (S/)")
        self.tree_partidas.heading("pct", text="% Ejec.")
        self.tree_partidas.heading("fuente", text="Fuente")

        self.tree_partidas.column("id", width=40)
        self.tree_partidas.column("categoria", width=160)
        self.tree_partidas.column("descripcion", width=200)
        self.tree_partidas.column("planificado", width=120)
        self.tree_partidas.column("ejecutado", width=120)
        self.tree_partidas.column("pct", width=70)
        self.tree_partidas.column("fuente", width=150)

        scroll_partidas = ttk.Scrollbar(tabla_frame, orient="vertical",
                                        command=self.tree_partidas.yview)
        self.tree_partidas.configure(yscrollcommand=scroll_partidas.set)
        self.tree_partidas.pack(side="left", fill="both", expand=True)
        scroll_partidas.pack(side="right", fill="y")

        self.tree_partidas.bind("<<TreeviewSelect>>", self.on_seleccionar_partida)

        # Subtotales del bloque
        self.label_subtotal = ttk.Label(tabla_frame, text="", font=("Segoe UI", 9, "bold"))
        self.label_subtotal.pack(anchor="w", pady=(6, 0))

        # ── Resumen general ──
        resumen_frame = ttk.LabelFrame(frame, text=" Resumen Presupuestal por Bloque ", padding=10)
        resumen_frame.grid(row=4, column=0, columnspan=3, sticky="ew", padx=12, pady=(6, 12))

        columnas_res = ("codigo", "tipo", "distrito", "planificado", "ejecutado", "pct", "partidas")
        self.tree_resumen = ttk.Treeview(resumen_frame, columns=columnas_res,
                                         show="headings", height=8)
        self.tree_resumen.heading("codigo", text="Código Bloque")
        self.tree_resumen.heading("tipo", text="Tipo Intervención")
        self.tree_resumen.heading("distrito", text="Distrito")
        self.tree_resumen.heading("planificado", text="Total Plan. (S/)")
        self.tree_resumen.heading("ejecutado", text="Total Ejec. (S/)")
        self.tree_resumen.heading("pct", text="% Ejec.")
        self.tree_resumen.heading("partidas", text="Partidas")

        self.tree_resumen.column("codigo", width=100)
        self.tree_resumen.column("tipo", width=180)
        self.tree_resumen.column("distrito", width=110)
        self.tree_resumen.column("planificado", width=120)
        self.tree_resumen.column("ejecutado", width=120)
        self.tree_resumen.column("pct", width=70)
        self.tree_resumen.column("partidas", width=70)

        self.tree_resumen.pack(fill="both", expand=True)

        # Totales globales
        self.label_total_global = ttk.Label(resumen_frame, text="",
                                            font=("Segoe UI", 10, "bold"))
        self.label_total_global.pack(anchor="w", pady=(6, 0))

    def cargar_combos(self):
        bloques = db.obtener_bloques()
        self.bloques_map = {f"{b['codigo']} - {b['tipo_intervencion']}": b["id"] for b in bloques}
        self.combo_bloque["values"] = list(self.bloques_map.keys())
        self.cargar_resumen_general()

    def cargar_partidas(self):
        """Carga partidas del bloque seleccionado."""
        sel = self.combo_bloque.get()
        if sel not in self.bloques_map:
            return
        bloque_id = self.bloques_map[sel]

        for item in self.tree_partidas.get_children():
            self.tree_partidas.delete(item)

        partidas = db.obtener_presupuesto_por_bloque(bloque_id)
        total_plan = 0
        total_ejec = 0

        for p in partidas:
            pct = (p["monto_ejecutado"] / p["monto_planificado"] * 100) if p["monto_planificado"] > 0 else 0
            total_plan += p["monto_planificado"]
            total_ejec += p["monto_ejecutado"]

            self.tree_partidas.insert("", "end", iid=p["id"], values=(
                p["id"],
                p["categoria"],
                p["descripcion"],
                f"{p['monto_planificado']:,.2f}",
                f"{p['monto_ejecutado']:,.2f}",
                f"{pct:.1f}%",
                p["fuente_financiamiento"],
            ))

        pct_total = (total_ejec / total_plan * 100) if total_plan > 0 else 0
        self.label_subtotal.config(
            text=f"Subtotal: Planificado S/ {total_plan:,.2f}  |  "
                 f"Ejecutado S/ {total_ejec:,.2f}  |  Ejecución: {pct_total:.1f}%"
        )

    def cargar_resumen_general(self):
        """Carga el resumen presupuestal de todos los bloques."""
        for item in self.tree_resumen.get_children():
            self.tree_resumen.delete(item)

        resumen = db.obtener_resumen_presupuesto()
        for r in resumen:
            pct = (r["total_ejecutado"] / r["total_planificado"] * 100) if r["total_planificado"] > 0 else 0
            self.tree_resumen.insert("", "end", values=(
                r["codigo"],
                r["tipo_intervencion"],
                r["distrito"],
                f"{r['total_planificado']:,.2f}",
                f"{r['total_ejecutado']:,.2f}",
                f"{pct:.1f}%",
                r["num_partidas"],
            ))

        totales = db.obtener_presupuesto_total()
        plan_t = totales["total_planificado"]
        ejec_t = totales["total_ejecutado"]
        pct_t = (ejec_t / plan_t * 100) if plan_t > 0 else 0
        saldo = plan_t - ejec_t

        self.label_total_global.config(
            text=f"TOTAL PROYECTO: Planificado S/ {plan_t:,.2f}  |  "
                 f"Ejecutado S/ {ejec_t:,.2f}  ({pct_t:.1f}%)  |  "
                 f"Saldo S/ {saldo:,.2f}"
        )

    def on_seleccionar_partida(self, event):
        sel = self.tree_partidas.selection()
        if not sel:
            return
        partida_id = int(sel[0])
        self.partida_seleccionada = partida_id

        # Buscar datos en la tabla
        values = self.tree_partidas.item(sel[0], "values")
        if values:
            self.combo_categoria.set(values[1])
            self.entry_descripcion.delete(0, tk.END)
            self.entry_descripcion.insert(0, values[2])
            self.entry_monto_plan.delete(0, tk.END)
            self.entry_monto_plan.insert(0, values[3].replace(",", ""))
            self.entry_monto_ejec.delete(0, tk.END)
            self.entry_monto_ejec.insert(0, values[4].replace(",", ""))
            self.combo_fuente.set(values[6])

    def _validar_campos(self):
        sel = self.combo_bloque.get()
        if not sel or sel not in self.bloques_map:
            messagebox.showwarning("Validación", "Seleccione un bloque.")
            return None
        try:
            monto_plan = float(self.entry_monto_plan.get().strip().replace(",", ""))
            monto_ejec = float(self.entry_monto_ejec.get().strip().replace(",", ""))
        except ValueError:
            messagebox.showwarning("Validación", "Los montos deben ser valores numéricos.")
            return None

        return {
            "bloque_id": self.bloques_map[sel],
            "categoria": self.combo_categoria.get(),
            "descripcion": self.entry_descripcion.get().strip(),
            "monto_planificado": monto_plan,
            "monto_ejecutado": monto_ejec,
            "fuente_financiamiento": self.combo_fuente.get(),
        }

    def guardar_partida(self):
        datos = self._validar_campos()
        if not datos:
            return
        try:
            db.insertar_presupuesto(**datos)
            messagebox.showinfo("Éxito", "Partida presupuestal registrada correctamente.")
            self.limpiar_formulario()
            self.cargar_partidas()
            self.cargar_resumen_general()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar:\n{e}")

    def actualizar_partida(self):
        if not self.partida_seleccionada:
            messagebox.showwarning("Selección", "Seleccione una partida de la tabla.")
            return
        datos = self._validar_campos()
        if not datos:
            return
        try:
            db.actualizar_presupuesto(
                self.partida_seleccionada,
                datos["categoria"], datos["descripcion"],
                datos["monto_planificado"], datos["monto_ejecutado"],
                datos["fuente_financiamiento"]
            )
            messagebox.showinfo("Éxito", "Partida actualizada correctamente.")
            self.limpiar_formulario()
            self.cargar_partidas()
            self.cargar_resumen_general()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo actualizar:\n{e}")

    def eliminar_partida(self):
        if not self.partida_seleccionada:
            messagebox.showwarning("Selección", "Seleccione una partida para eliminar.")
            return
        if messagebox.askyesno("Confirmar", "¿Eliminar esta partida presupuestal?"):
            db.eliminar_presupuesto(self.partida_seleccionada)
            messagebox.showinfo("Eliminado", "Partida eliminada.")
            self.limpiar_formulario()
            self.cargar_partidas()
            self.cargar_resumen_general()

    def limpiar_formulario(self):
        self.partida_seleccionada = None
        self.combo_categoria.set(CATEGORIAS_PRESUPUESTO[0])
        self.entry_descripcion.delete(0, tk.END)
        self.entry_monto_plan.delete(0, tk.END)
        self.entry_monto_plan.insert(0, "0.00")
        self.entry_monto_ejec.delete(0, tk.END)
        self.entry_monto_ejec.insert(0, "0.00")
        self.combo_fuente.set(FUENTES_FINANCIAMIENTO[0])
