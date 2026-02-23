"""
IN Piura - Módulo de Cronograma de Actividades
Gestión del cronograma de intervención: hitos, fechas planificadas vs reales,
seguimiento de avance por actividad y bloque.
Cuenca Alta del Río Piura, Perú.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

import database as db

COLOR_PRIMARIO = "#2C3E50"
COLOR_SECUNDARIO = "#3498DB"
COLOR_FONDO = "#ECF0F1"
COLOR_EXITO = "#27AE60"
COLOR_ALERTA = "#E74C3C"
COLOR_ADVERTENCIA = "#F39C12"

ESTADOS_ACTIVIDAD = [
    "Programado",
    "En ejecución",
    "Completado",
    "Retrasado",
    "Suspendido",
]

ACTIVIDADES_TIPO = [
    "Preparación de terreno",
    "Producción de plantones",
    "Plantación / Revegetación",
    "Excavación de zanjas de infiltración",
    "Construcción de terrazas",
    "Construcción de diques",
    "Mantenimiento y riego",
    "Monitoreo y evaluación",
    "Capacitación a comunidades",
    "Supervisión técnica",
    "Elaboración de informes",
    "Otra actividad",
]


class TabCronograma(ttk.Frame):
    """Pestaña de gestión del cronograma de actividades del proyecto."""

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.actividad_seleccionada = None
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
        ttk.Label(frame, text="Cronograma de Actividades",
                  style="Header.TLabel").grid(
            row=0, column=0, columnspan=4, sticky="w", padx=12, pady=(12, 4))

        ttk.Label(frame,
                  text="Programación y seguimiento de actividades por bloque de intervención.",
                  foreground="#7F8C8D").grid(
            row=1, column=0, columnspan=4, sticky="w", padx=12, pady=(0, 10))

        # ── Formulario ──
        form_frame = ttk.LabelFrame(frame, text=" Registrar Actividad ", padding=12)
        form_frame.grid(row=2, column=0, columnspan=4, sticky="ew", padx=12, pady=6)
        form_frame.columnconfigure(1, weight=1)
        form_frame.columnconfigure(3, weight=1)

        # Fila 1: Bloque y Actividad
        ttk.Label(form_frame, text="Bloque:").grid(row=0, column=0, sticky="w", pady=4)
        self.combo_bloque = ttk.Combobox(form_frame, state="readonly", width=36)
        self.combo_bloque.grid(row=0, column=1, sticky="w", padx=6, pady=4)
        self.combo_bloque.bind("<<ComboboxSelected>>", lambda e: self.cargar_actividades())

        ttk.Label(form_frame, text="Actividad:").grid(row=0, column=2, sticky="w", padx=(12, 0), pady=4)
        self.combo_actividad = ttk.Combobox(form_frame, values=ACTIVIDADES_TIPO, width=32)
        self.combo_actividad.grid(row=0, column=3, sticky="w", padx=6, pady=4)
        self.combo_actividad.set(ACTIVIDADES_TIPO[0])

        # Fila 2: Fechas planificadas
        ttk.Label(form_frame, text="Inicio planificado:").grid(row=1, column=0, sticky="w", pady=4)
        self.entry_inicio_plan = ttk.Entry(form_frame, width=16)
        self.entry_inicio_plan.grid(row=1, column=1, sticky="w", padx=6, pady=4)
        self.entry_inicio_plan.insert(0, datetime.now().strftime("%Y-%m-%d"))

        ttk.Label(form_frame, text="Fin planificado:").grid(row=1, column=2, sticky="w", padx=(12, 0), pady=4)
        self.entry_fin_plan = ttk.Entry(form_frame, width=16)
        self.entry_fin_plan.grid(row=1, column=3, sticky="w", padx=6, pady=4)
        self.entry_fin_plan.insert(0, datetime.now().strftime("%Y-%m-%d"))

        # Fila 3: Fechas reales
        ttk.Label(form_frame, text="Inicio real:").grid(row=2, column=0, sticky="w", pady=4)
        self.entry_inicio_real = ttk.Entry(form_frame, width=16)
        self.entry_inicio_real.grid(row=2, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(form_frame, text="Fin real:").grid(row=2, column=2, sticky="w", padx=(12, 0), pady=4)
        self.entry_fin_real = ttk.Entry(form_frame, width=16)
        self.entry_fin_real.grid(row=2, column=3, sticky="w", padx=6, pady=4)

        # Fila 4: Avance, responsable, estado
        ttk.Label(form_frame, text="Avance (%):").grid(row=3, column=0, sticky="w", pady=4)
        self.entry_avance = ttk.Entry(form_frame, width=10)
        self.entry_avance.grid(row=3, column=1, sticky="w", padx=6, pady=4)
        self.entry_avance.insert(0, "0")

        ttk.Label(form_frame, text="Estado:").grid(row=3, column=2, sticky="w", padx=(12, 0), pady=4)
        self.combo_estado = ttk.Combobox(form_frame, values=ESTADOS_ACTIVIDAD,
                                         state="readonly", width=16)
        self.combo_estado.grid(row=3, column=3, sticky="w", padx=6, pady=4)
        self.combo_estado.set(ESTADOS_ACTIVIDAD[0])

        # Fila 5: Responsable
        ttk.Label(form_frame, text="Responsable:").grid(row=4, column=0, sticky="w", pady=4)
        self.entry_responsable = ttk.Entry(form_frame, width=36)
        self.entry_responsable.grid(row=4, column=1, sticky="w", padx=6, pady=4)

        # Fila 6: Observaciones
        ttk.Label(form_frame, text="Observaciones:").grid(row=5, column=0, sticky="nw", pady=4)
        self.text_observaciones = tk.Text(form_frame, width=70, height=3, font=("Segoe UI", 9))
        self.text_observaciones.grid(row=5, column=1, columnspan=3, sticky="w", padx=6, pady=4)

        # Botones
        btn_frame = ttk.Frame(form_frame)
        btn_frame.grid(row=6, column=0, columnspan=4, pady=10)

        ttk.Button(btn_frame, text="Guardar Actividad",
                   command=self.guardar_actividad,
                   style="Accent.TButton").pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Actualizar",
                   command=self.actualizar_actividad).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Eliminar",
                   command=self.eliminar_actividad,
                   style="Danger.TButton").pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Limpiar",
                   command=self.limpiar_formulario).pack(side="left", padx=3)

        # ── Tabla de actividades del bloque ──
        tabla_frame = ttk.LabelFrame(frame, text=" Actividades del Bloque ", padding=10)
        tabla_frame.grid(row=3, column=0, columnspan=4, sticky="ew", padx=12, pady=6)

        columnas = ("id", "actividad", "inicio_plan", "fin_plan", "inicio_real",
                    "fin_real", "avance", "estado", "responsable")
        self.tree_actividades = ttk.Treeview(tabla_frame, columns=columnas,
                                             show="headings", height=8)
        self.tree_actividades.heading("id", text="ID")
        self.tree_actividades.heading("actividad", text="Actividad")
        self.tree_actividades.heading("inicio_plan", text="Inicio Plan.")
        self.tree_actividades.heading("fin_plan", text="Fin Plan.")
        self.tree_actividades.heading("inicio_real", text="Inicio Real")
        self.tree_actividades.heading("fin_real", text="Fin Real")
        self.tree_actividades.heading("avance", text="Avance %")
        self.tree_actividades.heading("estado", text="Estado")
        self.tree_actividades.heading("responsable", text="Responsable")

        self.tree_actividades.column("id", width=40)
        self.tree_actividades.column("actividad", width=200)
        self.tree_actividades.column("inicio_plan", width=90)
        self.tree_actividades.column("fin_plan", width=90)
        self.tree_actividades.column("inicio_real", width=90)
        self.tree_actividades.column("fin_real", width=90)
        self.tree_actividades.column("avance", width=70)
        self.tree_actividades.column("estado", width=100)
        self.tree_actividades.column("responsable", width=130)

        scroll_act = ttk.Scrollbar(tabla_frame, orient="vertical",
                                   command=self.tree_actividades.yview)
        self.tree_actividades.configure(yscrollcommand=scroll_act.set)
        self.tree_actividades.pack(side="left", fill="both", expand=True)
        scroll_act.pack(side="right", fill="y")

        self.tree_actividades.bind("<<TreeviewSelect>>", self.on_seleccionar_actividad)

        # ── Vista general de todas las actividades ──
        general_frame = ttk.LabelFrame(frame, text=" Cronograma General del Proyecto ", padding=10)
        general_frame.grid(row=4, column=0, columnspan=4, sticky="ew", padx=12, pady=(6, 4))

        # Filtros
        filtro_frame = ttk.Frame(general_frame)
        filtro_frame.pack(fill="x", pady=(0, 6))

        ttk.Label(filtro_frame, text="Filtrar estado:").pack(side="left", padx=(0, 4))
        self.combo_filtro_estado = ttk.Combobox(
            filtro_frame,
            values=["Todos"] + ESTADOS_ACTIVIDAD,
            state="readonly", width=16)
        self.combo_filtro_estado.pack(side="left", padx=2)
        self.combo_filtro_estado.set("Todos")
        self.combo_filtro_estado.bind("<<ComboboxSelected>>", lambda e: self.cargar_vista_general())

        columnas_gen = ("bloque", "actividad", "inicio_plan", "fin_plan",
                        "avance", "estado", "responsable")
        self.tree_general = ttk.Treeview(general_frame, columns=columnas_gen,
                                         show="headings", height=10)
        self.tree_general.heading("bloque", text="Bloque")
        self.tree_general.heading("actividad", text="Actividad")
        self.tree_general.heading("inicio_plan", text="Inicio Plan.")
        self.tree_general.heading("fin_plan", text="Fin Plan.")
        self.tree_general.heading("avance", text="Avance %")
        self.tree_general.heading("estado", text="Estado")
        self.tree_general.heading("responsable", text="Responsable")

        self.tree_general.column("bloque", width=90)
        self.tree_general.column("actividad", width=200)
        self.tree_general.column("inicio_plan", width=100)
        self.tree_general.column("fin_plan", width=100)
        self.tree_general.column("avance", width=70)
        self.tree_general.column("estado", width=100)
        self.tree_general.column("responsable", width=130)

        scroll_gen = ttk.Scrollbar(general_frame, orient="vertical",
                                   command=self.tree_general.yview)
        self.tree_general.configure(yscrollcommand=scroll_gen.set)
        self.tree_general.pack(side="left", fill="both", expand=True)
        scroll_gen.pack(side="right", fill="y")

        # Resumen estadístico
        self.label_resumen = ttk.Label(general_frame, text="",
                                       font=("Segoe UI", 9, "bold"))
        self.label_resumen.pack(anchor="w", pady=(6, 0))

    def cargar_combos(self):
        bloques = db.obtener_bloques()
        self.bloques_map = {f"{b['codigo']} - {b['tipo_intervencion']}": b["id"] for b in bloques}
        self.combo_bloque["values"] = list(self.bloques_map.keys())
        self.cargar_vista_general()

    def cargar_actividades(self):
        """Carga actividades del bloque seleccionado."""
        sel = self.combo_bloque.get()
        if sel not in self.bloques_map:
            return
        bloque_id = self.bloques_map[sel]

        for item in self.tree_actividades.get_children():
            self.tree_actividades.delete(item)

        actividades = db.obtener_actividades_por_bloque(bloque_id)
        for a in actividades:
            self.tree_actividades.insert("", "end", iid=a["id"], values=(
                a["id"],
                a["actividad"],
                a["fecha_inicio_plan"],
                a["fecha_fin_plan"],
                a["fecha_inicio_real"] or "-",
                a["fecha_fin_real"] or "-",
                f"{a['porcentaje_avance']:.0f}%",
                a["estado"],
                a["responsable"],
            ))

    def cargar_vista_general(self):
        """Carga la vista general de todas las actividades."""
        for item in self.tree_general.get_children():
            self.tree_general.delete(item)

        filtro = self.combo_filtro_estado.get()
        actividades = db.obtener_todas_actividades()

        conteo = {"Programado": 0, "En ejecución": 0, "Completado": 0,
                  "Retrasado": 0, "Suspendido": 0}

        for a in actividades:
            estado = a.get("estado", "Programado")
            conteo[estado] = conteo.get(estado, 0) + 1

            if filtro != "Todos" and estado != filtro:
                continue

            self.tree_general.insert("", "end", values=(
                a.get("bloque_codigo", ""),
                a["actividad"],
                a["fecha_inicio_plan"],
                a["fecha_fin_plan"],
                f"{a['porcentaje_avance']:.0f}%",
                estado,
                a["responsable"],
            ))

        total = sum(conteo.values())
        completadas = conteo.get("Completado", 0)
        pct = (completadas / total * 100) if total > 0 else 0

        self.label_resumen.config(
            text=f"Total: {total}  |  Programadas: {conteo.get('Programado', 0)}  |  "
                 f"En ejecución: {conteo.get('En ejecución', 0)}  |  "
                 f"Completadas: {completadas} ({pct:.0f}%)  |  "
                 f"Retrasadas: {conteo.get('Retrasado', 0)}"
        )

    def on_seleccionar_actividad(self, event):
        sel = self.tree_actividades.selection()
        if not sel:
            return
        self.actividad_seleccionada = int(sel[0])
        values = self.tree_actividades.item(sel[0], "values")

        if values:
            self.combo_actividad.set(values[1])
            self.entry_inicio_plan.delete(0, tk.END)
            self.entry_inicio_plan.insert(0, values[2])
            self.entry_fin_plan.delete(0, tk.END)
            self.entry_fin_plan.insert(0, values[3])
            self.entry_inicio_real.delete(0, tk.END)
            self.entry_inicio_real.insert(0, values[4] if values[4] != "-" else "")
            self.entry_fin_real.delete(0, tk.END)
            self.entry_fin_real.insert(0, values[5] if values[5] != "-" else "")
            self.entry_avance.delete(0, tk.END)
            self.entry_avance.insert(0, values[6].replace("%", ""))
            self.combo_estado.set(values[7])
            self.entry_responsable.delete(0, tk.END)
            self.entry_responsable.insert(0, values[8])

    def _validar_campos(self):
        sel = self.combo_bloque.get()
        if not sel or sel not in self.bloques_map:
            messagebox.showwarning("Validación", "Seleccione un bloque.")
            return None

        actividad = self.combo_actividad.get().strip()
        if not actividad:
            messagebox.showwarning("Validación", "Ingrese o seleccione una actividad.")
            return None

        inicio_plan = self.entry_inicio_plan.get().strip()
        fin_plan = self.entry_fin_plan.get().strip()
        if not inicio_plan or not fin_plan:
            messagebox.showwarning("Validación", "Las fechas planificadas son obligatorias.")
            return None

        try:
            avance = float(self.entry_avance.get().strip())
        except ValueError:
            messagebox.showwarning("Validación", "El avance debe ser un valor numérico.")
            return None

        return {
            "bloque_id": self.bloques_map[sel],
            "actividad": actividad,
            "fecha_inicio_plan": inicio_plan,
            "fecha_fin_plan": fin_plan,
            "fecha_inicio_real": self.entry_inicio_real.get().strip(),
            "fecha_fin_real": self.entry_fin_real.get().strip(),
            "porcentaje_avance": avance,
            "responsable": self.entry_responsable.get().strip(),
            "observaciones": self.text_observaciones.get("1.0", tk.END).strip(),
            "estado": self.combo_estado.get(),
        }

    def guardar_actividad(self):
        datos = self._validar_campos()
        if not datos:
            return
        try:
            db.insertar_actividad(**datos)
            messagebox.showinfo("Éxito", "Actividad registrada correctamente.")
            self.limpiar_formulario()
            self.cargar_actividades()
            self.cargar_vista_general()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar:\n{e}")

    def actualizar_actividad(self):
        if not self.actividad_seleccionada:
            messagebox.showwarning("Selección", "Seleccione una actividad de la tabla.")
            return
        datos = self._validar_campos()
        if not datos:
            return
        try:
            db.actualizar_actividad(
                self.actividad_seleccionada,
                datos["actividad"], datos["fecha_inicio_plan"],
                datos["fecha_fin_plan"], datos["fecha_inicio_real"],
                datos["fecha_fin_real"], datos["porcentaje_avance"],
                datos["responsable"], datos["observaciones"],
                datos["estado"]
            )
            messagebox.showinfo("Éxito", "Actividad actualizada correctamente.")
            self.limpiar_formulario()
            self.cargar_actividades()
            self.cargar_vista_general()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo actualizar:\n{e}")

    def eliminar_actividad(self):
        if not self.actividad_seleccionada:
            messagebox.showwarning("Selección", "Seleccione una actividad para eliminar.")
            return
        if messagebox.askyesno("Confirmar", "¿Eliminar esta actividad del cronograma?"):
            db.eliminar_actividad(self.actividad_seleccionada)
            messagebox.showinfo("Eliminado", "Actividad eliminada.")
            self.limpiar_formulario()
            self.cargar_actividades()
            self.cargar_vista_general()

    def limpiar_formulario(self):
        self.actividad_seleccionada = None
        self.combo_actividad.set(ACTIVIDADES_TIPO[0])
        self.entry_inicio_plan.delete(0, tk.END)
        self.entry_inicio_plan.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.entry_fin_plan.delete(0, tk.END)
        self.entry_fin_plan.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.entry_inicio_real.delete(0, tk.END)
        self.entry_fin_real.delete(0, tk.END)
        self.entry_avance.delete(0, tk.END)
        self.entry_avance.insert(0, "0")
        self.combo_estado.set(ESTADOS_ACTIVIDAD[0])
        self.entry_responsable.delete(0, tk.END)
        self.text_observaciones.delete("1.0", tk.END)
