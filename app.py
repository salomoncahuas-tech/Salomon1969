"""
IN Piura - Plan de Ingreso / Verificación de Campo
Aplicación de escritorio con interfaz gráfica (tkinter).
Restauración de ecosistemas - Cuenca alta del río Piura, Perú.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os
import uuid

import database as db
import reports
from georeferenciacion import TabGeorreferenciacion
from odk_kobo import TabODKKobo
from dashboard import TabDashboard
from presupuesto import TabPresupuesto
from cronograma import TabCronograma

# ── Constantes ─────────────────────────────────────────────────────────────

TIPOS_INTERVENCION = [
    "Revegetación",
    "Zanjas de infiltración",
    "Terrazas de formación lenta",
    "Diques de mampostería",
    "Otras",
]

ESTADOS_BLOQUE = ["Pendiente", "En progreso", "Verificado"]

CONDICIONES_CLIMATICAS = [
    "Despejado",
    "Parcialmente nublado",
    "Nublado",
    "Lluvia ligera",
    "Lluvia moderada",
    "Lluvia intensa",
    "Neblina",
]

# Microcuencas del Proyecto IN Piura - Cuenca Alta del Río Piura (29 códigos)
MICROCUENCAS = [
    "C1075-O9580",
    "C1076-O9581",
    "C1076-O9584",
    "C1076-O9585",
    "C1076-O9586",
    "C1076-O9587",
    "C1076-O9588",
    "C1076-O9589",
    "C1076-O9592",
    "C1077-O9566",
    "C1077-O9579",
    "C1078-O9562",
    "C1080-O9560",
    "C1081-O9582",
    "C1081-O9583",
    "C1081-O9591",
    "C1086-O9569",
    "C1086-O9570",
    "C1086-O9575",
    "C1086-O9576",
    "C1096-O9545",
    "C1096-O9547",
    "C1096-O9556",
    "C1096-O9557",
    "C1096-O9558",
    "C1096-O9564",
    "C1107-O9539",
    "C1107-O9541",
    "C1108-O9552",
]

# 5 Provincias y 18 Distritos del Proyecto IN Piura
PROVINCIAS_DISTRITOS = {
    "Ayabaca": [
        "Frías",
    ],
    "Huancabamba": [
        "Canchaque",
        "Huancabamba",
        "Huarmaca",
        "San Miguel de El Faique",
    ],
    "Morropón": [
        "Buenos Aires",
        "Chalaco",
        "Chulucanas",
        "La Matanza",
        "Morropón",
        "Salitral",
        "San Juan de Bigote",
        "Santa Catalina de Mossa",
        "Santo Domingo",
        "Yamango",
    ],
    "Piura": [
        "Las Lomas",
        "Tambo Grande",
    ],
    "Sullana": [
        "Sullana",
    ],
}

PROVINCIAS = list(PROVINCIAS_DISTRITOS.keys())

# Lista plana de todos los distritos (para validación)
DISTRITOS_PIURA = []
for _distritos in PROVINCIAS_DISTRITOS.values():
    DISTRITOS_PIURA.extend(_distritos)

# Tipos de cobertura vegetal
TIPOS_COBERTURA_VEGETAL = [
    "Arbórea",
    "Arbustiva",
    "Herbácea",
    "Mixta",
]

# Vigor / calidad de cobertura vegetal
VIGOR_COBERTURA_VEGETAL = [
    "Excelente",
    "Bueno",
    "Regular",
    "Deficiente",
    "Muy deficiente",
]

COLOR_PRIMARIO = "#2C3E50"
COLOR_SECUNDARIO = "#3498DB"
COLOR_FONDO = "#ECF0F1"
COLOR_EXITO = "#27AE60"
COLOR_ALERTA = "#E74C3C"


# ── Aplicación Principal ──────────────────────────────────────────────────

class AplicacionINPiura(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("IN Piura - Plan de Ingreso - Verificación de Campo")
        self.geometry("1280x800")
        self.minsize(1080, 720)
        self.configure(bg=COLOR_FONDO)

        # Inicializar base de datos
        db.inicializar_bd()

        self._crear_estilos()
        self._crear_interfaz()

    def _crear_estilos(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("TNotebook", background=COLOR_FONDO)
        style.configure("TNotebook.Tab", padding=[14, 6], font=("Segoe UI", 10))
        style.map("TNotebook.Tab",
                  background=[("selected", COLOR_PRIMARIO)],
                  foreground=[("selected", "white")])

        style.configure("TFrame", background=COLOR_FONDO)
        style.configure("TLabel", background=COLOR_FONDO, font=("Segoe UI", 9))
        style.configure("TButton", font=("Segoe UI", 9), padding=6)
        style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"),
                        foreground=COLOR_PRIMARIO)
        style.configure("Accent.TButton", background=COLOR_SECUNDARIO,
                        foreground="white", font=("Segoe UI", 9, "bold"))
        style.configure("Success.TButton", background=COLOR_EXITO,
                        foreground="white")
        style.configure("Danger.TButton", background=COLOR_ALERTA,
                        foreground="white")

        style.configure("Treeview", font=("Segoe UI", 9), rowheight=26)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

    def _crear_interfaz(self):
        # Cabecera
        header_frame = tk.Frame(self, bg=COLOR_PRIMARIO, height=56)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)

        tk.Label(header_frame, text="IN Piura", font=("Segoe UI", 18, "bold"),
                 bg=COLOR_PRIMARIO, fg="white").pack(side="left", padx=16)
        tk.Label(header_frame,
                 text="Plan de Ingreso  |  Verificación de Campo  |  Cuenca Alta del Río Piura",
                 font=("Segoe UI", 10), bg=COLOR_PRIMARIO, fg="#BDC3C7").pack(side="left", padx=8)

        # Notebook (pestañas)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)

        self.tab_dashboard = TabDashboard(self.notebook, self)
        self.tab_bloques = TabBloques(self.notebook, self)
        self.tab_inspeccion = TabInspeccion(self.notebook, self)
        self.tab_indicadores = TabIndicadores(self.notebook, self)
        self.tab_presupuesto = TabPresupuesto(self.notebook, self)
        self.tab_cronograma = TabCronograma(self.notebook, self)
        self.tab_georef = TabGeorreferenciacion(self.notebook, self)
        self.tab_odk = TabODKKobo(self.notebook, self)
        self.tab_reportes = TabReportes(self.notebook, self)

        self.notebook.add(self.tab_dashboard, text="  Panel de Control  ")
        self.notebook.add(self.tab_bloques, text="  Bloques de Intervención  ")
        self.notebook.add(self.tab_inspeccion, text="  Inspección de Campo  ")
        self.notebook.add(self.tab_indicadores, text="  Indicadores de Calidad  ")
        self.notebook.add(self.tab_presupuesto, text="  Presupuesto  ")
        self.notebook.add(self.tab_cronograma, text="  Cronograma  ")
        self.notebook.add(self.tab_georef, text="  Georreferenciación  ")
        self.notebook.add(self.tab_odk, text="  ODK / KoBoToolbox  ")
        self.notebook.add(self.tab_reportes, text="  Reportes  ")

    def refrescar_todo(self):
        self.tab_dashboard.actualizar_datos()
        self.tab_bloques.cargar_bloques()
        self.tab_inspeccion.cargar_combos()
        self.tab_indicadores.cargar_combos()
        self.tab_presupuesto.cargar_combos()
        self.tab_cronograma.cargar_combos()
        self.tab_georef.cargar_datos()
        self.tab_reportes.cargar_combos()


# ── Tab 1: Bloques de Intervención ────────────────────────────────────────

class TabBloques(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.bloque_seleccionado = None
        self._crear_widgets()
        self.cargar_bloques()

    def _crear_widgets(self):
        # Panel izquierdo - Formulario
        panel_izq = ttk.Frame(self, padding=10)
        panel_izq.pack(side="left", fill="y", padx=(0, 4))

        ttk.Label(panel_izq, text="Registro de Bloque", style="Header.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        campos = [
            ("Código de bloque:", "codigo"),
            ("Cuenca hidrográfica:", "cuenca"),
            ("UTM Este (m):", "utm_este"),
            ("UTM Norte (m):", "utm_norte"),
            ("Zona UTM:", "utm_zona"),
            ("Altitud (m.s.n.m.):", "altitud"),
            ("Área (ha):", "area"),
            ("Responsable:", "responsable"),
        ]

        self.entries = {}
        for i, (label, key) in enumerate(campos, 1):
            ttk.Label(panel_izq, text=label).grid(row=i, column=0, sticky="w", pady=3)
            entry = ttk.Entry(panel_izq, width=24)
            entry.grid(row=i, column=1, sticky="ew", pady=3, padx=(6, 0))
            self.entries[key] = entry

        # Valores por defecto
        self.entries["utm_zona"].insert(0, "17S")
        self.entries["cuenca"].insert(0, "Cuenca Alta del Río Piura")
        self.entries["altitud"].insert(0, "0")

        # Microcuenca (lista desplegable validada)
        fila = len(campos) + 1
        ttk.Label(panel_izq, text="Código microcuenca:").grid(row=fila, column=0, sticky="w", pady=3)
        self.combo_microcuenca = ttk.Combobox(panel_izq, values=MICROCUENCAS, state="readonly", width=22)
        self.combo_microcuenca.grid(row=fila, column=1, sticky="ew", pady=3, padx=(6, 0))

        # Provincia (lista desplegable de 5 provincias)
        fila += 1
        ttk.Label(panel_izq, text="Provincia:").grid(row=fila, column=0, sticky="w", pady=3)
        self.combo_provincia = ttk.Combobox(panel_izq, values=PROVINCIAS, state="readonly", width=22)
        self.combo_provincia.grid(row=fila, column=1, sticky="ew", pady=3, padx=(6, 0))
        self.combo_provincia.bind("<<ComboboxSelected>>", self._on_provincia_seleccionada)

        # Distrito (filtrado por provincia seleccionada)
        fila += 1
        ttk.Label(panel_izq, text="Distrito:").grid(row=fila, column=0, sticky="w", pady=3)
        self.combo_distrito = ttk.Combobox(panel_izq, values=DISTRITOS_PIURA, state="readonly", width=22)
        self.combo_distrito.grid(row=fila, column=1, sticky="ew", pady=3, padx=(6, 0))

        # Tipo de intervención (incluye "Otras")
        fila += 1
        ttk.Label(panel_izq, text="Tipo intervención:").grid(row=fila, column=0, sticky="w", pady=3)
        self.combo_tipo = ttk.Combobox(panel_izq, values=TIPOS_INTERVENCION, state="readonly", width=22)
        self.combo_tipo.grid(row=fila, column=1, sticky="ew", pady=3, padx=(6, 0))
        self.combo_tipo.set(TIPOS_INTERVENCION[0])

        # Estado
        fila += 1
        ttk.Label(panel_izq, text="Estado:").grid(row=fila, column=0, sticky="w", pady=3)
        self.combo_estado = ttk.Combobox(panel_izq, values=ESTADOS_BLOQUE, state="readonly", width=22)
        self.combo_estado.grid(row=fila, column=1, sticky="ew", pady=3, padx=(6, 0))
        self.combo_estado.set(ESTADOS_BLOQUE[0])

        # Botones
        fila += 1
        btn_frame = ttk.Frame(panel_izq)
        btn_frame.grid(row=fila, column=0, columnspan=2, pady=12)

        ttk.Button(btn_frame, text="Guardar", command=self.guardar_bloque,
                   style="Accent.TButton").pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Actualizar", command=self.actualizar_bloque).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Eliminar", command=self.eliminar_bloque,
                   style="Danger.TButton").pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Limpiar", command=self.limpiar_formulario).pack(side="left", padx=3)

        # Panel derecho - Tabla
        panel_der = ttk.Frame(self, padding=10)
        panel_der.pack(side="right", fill="both", expand=True)

        # Barra de búsqueda
        busqueda_frame = ttk.Frame(panel_der)
        busqueda_frame.pack(fill="x", pady=(0, 6))

        ttk.Label(busqueda_frame, text="Bloques Registrados",
                  style="Header.TLabel").pack(side="left")

        ttk.Button(busqueda_frame, text="Buscar",
                   command=self._buscar_bloques).pack(side="right", padx=(4, 0))
        ttk.Button(busqueda_frame, text="Mostrar todos",
                   command=self.cargar_bloques).pack(side="right", padx=4)
        self.entry_busqueda = ttk.Entry(busqueda_frame, width=20)
        self.entry_busqueda.pack(side="right", padx=4)
        self.entry_busqueda.bind("<Return>", lambda e: self._buscar_bloques())
        ttk.Label(busqueda_frame, text="Buscar:").pack(side="right")

        columnas = ("codigo", "microcuenca", "tipo", "provincia", "distrito",
                    "utm_este", "utm_norte", "altitud", "area", "responsable", "estado")
        self.tree = ttk.Treeview(panel_der, columns=columnas, show="headings", height=16)

        self.tree.heading("codigo", text="Código")
        self.tree.heading("microcuenca", text="Microcuenca")
        self.tree.heading("tipo", text="Tipo Intervención")
        self.tree.heading("provincia", text="Provincia")
        self.tree.heading("distrito", text="Distrito")
        self.tree.heading("utm_este", text="UTM Este")
        self.tree.heading("utm_norte", text="UTM Norte")
        self.tree.heading("altitud", text="Altitud")
        self.tree.heading("area", text="Área (ha)")
        self.tree.heading("responsable", text="Responsable")
        self.tree.heading("estado", text="Estado")

        self.tree.column("codigo", width=75)
        self.tree.column("microcuenca", width=110)
        self.tree.column("tipo", width=130)
        self.tree.column("provincia", width=85)
        self.tree.column("distrito", width=90)
        self.tree.column("utm_este", width=75)
        self.tree.column("utm_norte", width=75)
        self.tree.column("altitud", width=55)
        self.tree.column("area", width=55)
        self.tree.column("responsable", width=85)
        self.tree.column("estado", width=70)

        scrollbar = ttk.Scrollbar(panel_der, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.tree.bind("<<TreeviewSelect>>", self.on_seleccionar_bloque)

    def cargar_bloques(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        bloques = db.obtener_bloques()
        self._mostrar_bloques(bloques)

    def _buscar_bloques(self):
        texto = self.entry_busqueda.get().strip()
        if not texto:
            self.cargar_bloques()
            return
        for item in self.tree.get_children():
            self.tree.delete(item)
        bloques = db.buscar_bloques(texto)
        self._mostrar_bloques(bloques)

    def _mostrar_bloques(self, bloques):
        for b in bloques:
            altitud = b.get("altitud", 0) or 0
            responsable = b.get("responsable", "") or ""
            microcuenca = b.get("microcuenca", "") or ""
            provincia = b.get("provincia", "") or ""
            self.tree.insert("", "end", iid=b["id"], values=(
                b["codigo"], microcuenca, b["tipo_intervencion"],
                provincia, b["distrito"],
                f"{b['utm_este']:.2f}", f"{b['utm_norte']:.2f}",
                f"{altitud:.0f}", f"{b['area_hectareas']:.4f}",
                responsable, b["estado"]
            ))

    def _on_provincia_seleccionada(self, event=None):
        """Filtra los distritos según la provincia seleccionada."""
        provincia = self.combo_provincia.get()
        if provincia in PROVINCIAS_DISTRITOS:
            distritos = PROVINCIAS_DISTRITOS[provincia]
            self.combo_distrito["values"] = distritos
            self.combo_distrito.set("")
        else:
            self.combo_distrito["values"] = DISTRITOS_PIURA
            self.combo_distrito.set("")

    def on_seleccionar_bloque(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        bloque_id = int(sel[0])
        bloque = db.obtener_bloque_por_id(bloque_id)
        if not bloque:
            return

        self.bloque_seleccionado = bloque_id
        self.limpiar_formulario(reset_defaults=False)

        self.entries["codigo"].insert(0, bloque["codigo"])
        self.entries["cuenca"].insert(0, bloque["cuenca"])
        self.entries["utm_este"].insert(0, str(bloque["utm_este"]))
        self.entries["utm_norte"].insert(0, str(bloque["utm_norte"]))
        self.entries["utm_zona"].insert(0, bloque["utm_zona"])
        self.entries["altitud"].insert(0, str(bloque.get("altitud", 0) or 0))
        self.entries["area"].insert(0, str(bloque["area_hectareas"]))
        self.entries["responsable"].insert(0, bloque.get("responsable", "") or "")
        self.combo_tipo.set(bloque["tipo_intervencion"])
        self.combo_estado.set(bloque["estado"])
        # Restaurar provincia y distrito
        provincia_val = bloque.get("provincia", "") or ""
        if provincia_val in PROVINCIAS:
            self.combo_provincia.set(provincia_val)
            self.combo_distrito["values"] = PROVINCIAS_DISTRITOS.get(provincia_val, DISTRITOS_PIURA)
        else:
            self.combo_provincia.set("")
            self.combo_distrito["values"] = DISTRITOS_PIURA
        distrito_val = bloque.get("distrito", "")
        if distrito_val in DISTRITOS_PIURA:
            self.combo_distrito.set(distrito_val)
        else:
            self.combo_distrito.set("")
        # Restaurar microcuenca en combobox
        microcuenca_val = bloque.get("microcuenca", "") or ""
        if microcuenca_val in MICROCUENCAS:
            self.combo_microcuenca.set(microcuenca_val)
        else:
            self.combo_microcuenca.set("")

    def _validar_campos(self):
        codigo = self.entries["codigo"].get().strip()
        if not codigo:
            messagebox.showwarning("Validación", "El código de bloque es obligatorio.")
            return None

        distrito = self.combo_distrito.get()
        if not distrito:
            messagebox.showwarning("Validación", "Debe seleccionar un distrito de la lista.")
            return None

        try:
            utm_este = float(self.entries["utm_este"].get().strip())
            utm_norte = float(self.entries["utm_norte"].get().strip())
            area = float(self.entries["area"].get().strip())
        except ValueError:
            messagebox.showwarning("Validación",
                                   "UTM Este, UTM Norte y Área deben ser valores numéricos.")
            return None

        try:
            altitud = float(self.entries["altitud"].get().strip() or "0")
        except ValueError:
            altitud = 0

        return {
            "codigo": codigo,
            "tipo_intervencion": self.combo_tipo.get(),
            "cuenca": self.entries["cuenca"].get().strip() or "Cuenca Alta del Río Piura",
            "distrito": distrito,
            "utm_este": utm_este,
            "utm_norte": utm_norte,
            "utm_zona": self.entries["utm_zona"].get().strip() or "17S",
            "area_hectareas": area,
            "estado": self.combo_estado.get(),
            "altitud": altitud,
            "responsable": self.entries["responsable"].get().strip(),
            "microcuenca": self.combo_microcuenca.get(),
            "provincia": self.combo_provincia.get(),
        }

    def guardar_bloque(self):
        datos = self._validar_campos()
        if not datos:
            return
        try:
            db.insertar_bloque(**datos)
            messagebox.showinfo("Éxito", f"Bloque {datos['codigo']} registrado correctamente.")
            self.limpiar_formulario()
            self.cargar_bloques()
            self.app.refrescar_todo()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el bloque:\n{e}")

    def actualizar_bloque(self):
        if not self.bloque_seleccionado:
            messagebox.showwarning("Selección", "Seleccione un bloque de la tabla para actualizar.")
            return
        datos = self._validar_campos()
        if not datos:
            return
        try:
            db.actualizar_bloque(self.bloque_seleccionado, **datos)
            messagebox.showinfo("Éxito", f"Bloque {datos['codigo']} actualizado.")
            self.limpiar_formulario()
            self.cargar_bloques()
            self.app.refrescar_todo()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo actualizar:\n{e}")

    def eliminar_bloque(self):
        if not self.bloque_seleccionado:
            messagebox.showwarning("Selección", "Seleccione un bloque para eliminar.")
            return
        if messagebox.askyesno("Confirmar", "¿Eliminar este bloque y todas sus inspecciones asociadas?"):
            db.eliminar_bloque(self.bloque_seleccionado)
            messagebox.showinfo("Eliminado", "Bloque eliminado correctamente.")
            self.limpiar_formulario()
            self.cargar_bloques()
            self.app.refrescar_todo()

    def limpiar_formulario(self, reset_defaults=True):
        self.bloque_seleccionado = None
        for entry in self.entries.values():
            entry.delete(0, tk.END)
        self.combo_provincia.set("")
        self.combo_distrito.set("")
        self.combo_distrito["values"] = DISTRITOS_PIURA
        self.combo_microcuenca.set("")
        if reset_defaults:
            self.entries["utm_zona"].insert(0, "17S")
            self.entries["cuenca"].insert(0, "Cuenca Alta del Río Piura")
            self.entries["altitud"].insert(0, "0")
            self.combo_tipo.set(TIPOS_INTERVENCION[0])
            self.combo_estado.set(ESTADOS_BLOQUE[0])


# ── Tab 2: Inspección de Campo ────────────────────────────────────────────

class TabInspeccion(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.rutas_fotos = []
        self._crear_widgets()
        self.cargar_combos()

    def _crear_widgets(self):
        # Contenedor principal con scroll
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

        ttk.Label(frame, text="Formulario de Inspección de Campo",
                  style="Header.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 12))

        # Bloque
        ttk.Label(frame, text="Bloque:").grid(row=1, column=0, sticky="w", padx=10, pady=4)
        self.combo_bloque = ttk.Combobox(frame, state="readonly", width=40)
        self.combo_bloque.grid(row=1, column=1, sticky="w", padx=6, pady=4)

        # Código de microcuenca
        ttk.Label(frame, text="Código microcuenca:").grid(row=2, column=0, sticky="w", padx=10, pady=4)
        self.combo_microcuenca = ttk.Combobox(frame, values=MICROCUENCAS, state="readonly", width=40)
        self.combo_microcuenca.grid(row=2, column=1, sticky="w", padx=6, pady=4)

        # Fecha
        ttk.Label(frame, text="Fecha de visita:").grid(row=3, column=0, sticky="w", padx=10, pady=4)
        self.entry_fecha = ttk.Entry(frame, width=20)
        self.entry_fecha.grid(row=3, column=1, sticky="w", padx=6, pady=4)
        self.entry_fecha.insert(0, datetime.now().strftime("%Y-%m-%d"))

        # Inspector
        ttk.Label(frame, text="Nombre del inspector:").grid(row=4, column=0, sticky="w", padx=10, pady=4)
        self.entry_inspector = ttk.Entry(frame, width=40)
        self.entry_inspector.grid(row=4, column=1, sticky="w", padx=6, pady=4)

        # Condiciones climáticas
        ttk.Label(frame, text="Condiciones climáticas:").grid(row=5, column=0, sticky="w", padx=10, pady=4)
        self.combo_clima = ttk.Combobox(frame, values=CONDICIONES_CLIMATICAS, state="readonly", width=28)
        self.combo_clima.grid(row=5, column=1, sticky="w", padx=6, pady=4)
        self.combo_clima.set(CONDICIONES_CLIMATICAS[0])

        # Avance físico
        self.entry_avance = ttk.Entry(frame, width=12)
        self.entry_avance.grid(row=6, column=1, sticky="w", padx=6, pady=4)
        self.entry_avance.insert(0, "0")

        # Observaciones
        ttk.Label(frame, text="Observaciones técnicas:").grid(row=7, column=0, sticky="nw", padx=10, pady=4)
        self.text_observaciones = tk.Text(frame, width=55, height=4, font=("Segoe UI", 9))
        self.text_observaciones.grid(row=7, column=1, sticky="w", padx=6, pady=4)

        # Desviaciones
        ttk.Label(frame, text="Desviaciones al exp. técnico:").grid(row=8, column=0, sticky="nw", padx=10, pady=4)
        self.text_desviaciones = tk.Text(frame, width=55, height=4, font=("Segoe UI", 9))
        self.text_desviaciones.grid(row=8, column=1, sticky="w", padx=6, pady=4)

        # Registro fotográfico
        ttk.Label(frame, text="Registro fotográfico:").grid(row=9, column=0, sticky="nw", padx=10, pady=4)
        foto_frame = ttk.Frame(frame)
        foto_frame.grid(row=9, column=1, sticky="w", padx=6, pady=4)

        ttk.Button(foto_frame, text="Agregar imágenes...", command=self.agregar_fotos).pack(side="left")
        ttk.Button(foto_frame, text="Limpiar fotos", command=self.limpiar_fotos).pack(side="left", padx=6)

        self.label_fotos = ttk.Label(frame, text="Sin imágenes seleccionadas.", wraplength=450)
        self.label_fotos.grid(row=10, column=1, sticky="w", padx=6, pady=2)

        # Código de verificación
        ttk.Label(frame, text="Código de verificación:").grid(row=11, column=0, sticky="w", padx=10, pady=4)
        verif_frame = ttk.Frame(frame)
        verif_frame.grid(row=11, column=1, sticky="w", padx=6, pady=4)

        self.entry_verificacion = ttk.Entry(verif_frame, width=30)
        self.entry_verificacion.pack(side="left")
        ttk.Button(verif_frame, text="Generar", command=self.generar_codigo).pack(side="left", padx=6)

        # Botones de acción
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=12, column=0, columnspan=2, pady=16, padx=10)

        ttk.Button(btn_frame, text="Guardar Inspección", command=self.guardar_inspeccion,
                   style="Accent.TButton").pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Limpiar Formulario", command=self.limpiar_formulario).pack(side="left", padx=4)

        # Historial de inspecciones
        ttk.Label(frame, text="Historial de Inspecciones", style="Header.TLabel").grid(
            row=13, column=0, columnspan=3, sticky="w", padx=10, pady=(16, 6))

        columnas = ("id", "bloque", "microcuenca", "fecha", "inspector", "avance", "verificacion")
        self.tree_hist = ttk.Treeview(frame, columns=columnas, show="headings", height=8)
        self.tree_hist.heading("id", text="ID")
        self.tree_hist.heading("bloque", text="Bloque")
        self.tree_hist.heading("microcuenca", text="Microcuenca")
        self.tree_hist.heading("fecha", text="Fecha Visita")
        self.tree_hist.heading("inspector", text="Inspector")
        self.tree_hist.heading("avance", text="Avance %")
        self.tree_hist.heading("verificacion", text="Cód. Verificación")

        self.tree_hist.column("id", width=40)
        self.tree_hist.column("bloque", width=90)
        self.tree_hist.column("microcuenca", width=130)
        self.tree_hist.column("fecha", width=100)
        self.tree_hist.column("inspector", width=140)
        self.tree_hist.column("avance", width=70)
        self.tree_hist.column("verificacion", width=180)

        self.tree_hist.grid(row=14, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 10))
        self.cargar_historial()

    def cargar_combos(self):
        bloques = db.obtener_bloques()
        self.bloques_map = {f"{b['codigo']} - {b['tipo_intervencion']}": b["id"] for b in bloques}
        self.combo_bloque["values"] = list(self.bloques_map.keys())
        self.cargar_historial()

    def cargar_historial(self):
        for item in self.tree_hist.get_children():
            self.tree_hist.delete(item)
        inspecciones = db.obtener_todas_inspecciones()
        for insp in inspecciones:
            self.tree_hist.insert("", "end", values=(
                insp["id"], insp["bloque_codigo"],
                insp.get("microcuenca", "") or "",
                insp["fecha_visita"],
                insp["inspector"], f"{insp['avance_fisico']:.1f}",
                insp["codigo_verificacion"]
            ))

    def agregar_fotos(self):
        rutas = filedialog.askopenfilenames(
            title="Seleccionar imágenes",
            filetypes=[("Imágenes", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff"), ("Todos", "*.*")]
        )
        if rutas:
            self.rutas_fotos.extend(rutas)
            self.label_fotos.config(text=f"{len(self.rutas_fotos)} imagen(es) seleccionada(s):\n" +
                                         "\n".join(os.path.basename(r) for r in self.rutas_fotos[-5:]))

    def limpiar_fotos(self):
        self.rutas_fotos = []
        self.label_fotos.config(text="Sin imágenes seleccionadas.")

    def generar_codigo(self):
        codigo = f"VER-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}"
        self.entry_verificacion.delete(0, tk.END)
        self.entry_verificacion.insert(0, codigo)

    def guardar_inspeccion(self):
        sel = self.combo_bloque.get()
        if not sel or sel not in self.bloques_map:
            messagebox.showwarning("Validación", "Seleccione un bloque de intervención.")
            return

        fecha = self.entry_fecha.get().strip()
        inspector = self.entry_inspector.get().strip()
        if not fecha or not inspector:
            messagebox.showwarning("Validación", "La fecha de visita y el inspector son obligatorios.")
            return

        try:
            avance = float(self.entry_avance.get().strip())
        except ValueError:
            messagebox.showwarning("Validación", "El avance físico debe ser un valor numérico.")
            return

        bloque_id = self.bloques_map[sel]
        registro_foto = "; ".join(self.rutas_fotos) if self.rutas_fotos else ""

        try:
            db.insertar_inspeccion(
                bloque_id=bloque_id,
                fecha_visita=fecha,
                inspector=inspector,
                condiciones_climaticas=self.combo_clima.get(),
                avance_fisico=avance,
                observaciones=self.text_observaciones.get("1.0", tk.END).strip(),
                desviaciones=self.text_desviaciones.get("1.0", tk.END).strip(),
                registro_fotografico=registro_foto,
                codigo_verificacion=self.entry_verificacion.get().strip(),
                microcuenca=self.combo_microcuenca.get(),
            )
            messagebox.showinfo("Éxito", "Inspección registrada correctamente.")
            self.limpiar_formulario()
            self.cargar_historial()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la inspección:\n{e}")

    def limpiar_formulario(self):
        self.combo_bloque.set("")
        self.combo_microcuenca.set("")
        self.entry_fecha.delete(0, tk.END)
        self.entry_fecha.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.entry_inspector.delete(0, tk.END)
        self.combo_clima.set(CONDICIONES_CLIMATICAS[0])
        self.entry_avance.delete(0, tk.END)
        self.entry_avance.insert(0, "0")
        self.text_observaciones.delete("1.0", tk.END)
        self.text_desviaciones.delete("1.0", tk.END)
        self.limpiar_fotos()
        self.entry_verificacion.delete(0, tk.END)


# ── Tab 3: Indicadores de Calidad ─────────────────────────────────────────

class TabIndicadores(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self._crear_widgets()
        self.cargar_combos()

    def _crear_widgets(self):
        # Contenedor con scroll
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

        ttk.Label(frame, text="Indicadores de Calidad por Bloque",
                  style="Header.TLabel").grid(row=0, column=0, columnspan=4, sticky="w", padx=10, pady=(10, 12))

        # Selección de bloque
        ttk.Label(frame, text="Bloque:").grid(row=1, column=0, sticky="w", padx=10, pady=4)
        self.combo_bloque = ttk.Combobox(frame, state="readonly", width=40)
        self.combo_bloque.grid(row=1, column=1, columnspan=2, sticky="w", padx=6, pady=4)
        self.combo_bloque.bind("<<ComboboxSelected>>", self.on_bloque_seleccionado)

        # Código de microcuenca
        ttk.Label(frame, text="Código microcuenca:").grid(row=2, column=0, sticky="w", padx=10, pady=4)
        self.combo_microcuenca = ttk.Combobox(frame, values=MICROCUENCAS, state="readonly", width=40)
        self.combo_microcuenca.grid(row=2, column=1, columnspan=2, sticky="w", padx=6, pady=4)

        # Selección de inspección
        ttk.Label(frame, text="Inspección:").grid(row=3, column=0, sticky="w", padx=10, pady=4)
        self.combo_inspeccion = ttk.Combobox(frame, state="readonly", width=40)
        self.combo_inspeccion.grid(row=3, column=1, columnspan=2, sticky="w", padx=6, pady=4)

        # Indicadores
        ttk.Separator(frame, orient="horizontal").grid(row=4, column=0, columnspan=4, sticky="ew", pady=10)

        # Campos numéricos de indicadores
        indicadores_campos = [
            ("Porcentaje de cobertura vegetal (%):", "pct_cobertura"),
            ("Sobrevivencia de especies (%):", "sobrevivencia"),
            ("Longitud ejecutada de zanjas (ml):", "longitud_zanjas"),
            ("Volumen retención sedimentos estimado (m³):", "vol_retencion"),
        ]

        self.ind_entries = {}
        for i, (label, key) in enumerate(indicadores_campos, 5):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky="w", padx=10, pady=5)
            entry = ttk.Entry(frame, width=16)
            entry.grid(row=i, column=1, sticky="w", padx=6, pady=5)
            entry.insert(0, "0")
            self.ind_entries[key] = entry

        # Tipo de cobertura vegetal (combo)
        fila_tipo_cob = 5 + len(indicadores_campos)
        ttk.Label(frame, text="Tipo de cobertura vegetal:").grid(
            row=fila_tipo_cob, column=0, sticky="w", padx=10, pady=5)
        self.combo_tipo_cobertura = ttk.Combobox(
            frame, values=TIPOS_COBERTURA_VEGETAL, state="readonly", width=20)
        self.combo_tipo_cobertura.grid(row=fila_tipo_cob, column=1, sticky="w", padx=6, pady=5)

        # Vigor / calidad de cobertura vegetal (combo)
        fila_vigor = fila_tipo_cob + 1
        ttk.Label(frame, text="Calidad / vigor cobertura vegetal:").grid(
            row=fila_vigor, column=0, sticky="w", padx=10, pady=5)
        self.combo_vigor = ttk.Combobox(
            frame, values=VIGOR_COBERTURA_VEGETAL, state="readonly", width=20)
        self.combo_vigor.grid(row=fila_vigor, column=1, sticky="w", padx=6, pady=5)

        # Botones
        fila_btn = fila_vigor + 1
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=fila_btn, column=0, columnspan=4, pady=14, padx=10)
        ttk.Button(btn_frame, text="Guardar Indicadores",
                   command=self.guardar_indicadores, style="Accent.TButton").pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Limpiar", command=self.limpiar_formulario).pack(side="left", padx=4)

        # Tabla de indicadores registrados
        fila_tabla_label = fila_btn + 1
        ttk.Label(frame, text="Indicadores Registrados", style="Header.TLabel").grid(
            row=fila_tabla_label, column=0, columnspan=4, sticky="w", padx=10, pady=(14, 6))

        columnas = ("fecha", "microcuenca", "pct_cob", "tipo_cob", "vigor",
                    "sobrev", "zanjas", "vol_ret")
        self.tree_ind = ttk.Treeview(frame, columns=columnas, show="headings", height=8)
        self.tree_ind.heading("fecha", text="Fecha Visita")
        self.tree_ind.heading("microcuenca", text="Microcuenca")
        self.tree_ind.heading("pct_cob", text="Cobert. (%)")
        self.tree_ind.heading("tipo_cob", text="Tipo Cobert.")
        self.tree_ind.heading("vigor", text="Vigor")
        self.tree_ind.heading("sobrev", text="Sobreviv. (%)")
        self.tree_ind.heading("zanjas", text="Zanjas (ml)")
        self.tree_ind.heading("vol_ret", text="Vol. Ret. (m³)")

        self.tree_ind.column("fecha", width=95)
        self.tree_ind.column("microcuenca", width=120)
        self.tree_ind.column("pct_cob", width=80)
        self.tree_ind.column("tipo_cob", width=90)
        self.tree_ind.column("vigor", width=80)
        self.tree_ind.column("sobrev", width=80)
        self.tree_ind.column("zanjas", width=85)
        self.tree_ind.column("vol_ret", width=85)

        fila_tabla = fila_tabla_label + 1
        self.tree_ind.grid(row=fila_tabla, column=0, columnspan=4, sticky="ew", padx=10, pady=(0, 10))

    def cargar_combos(self):
        bloques = db.obtener_bloques()
        self.bloques_map = {f"{b['codigo']} - {b['tipo_intervencion']}": b["id"] for b in bloques}
        self.combo_bloque["values"] = list(self.bloques_map.keys())

    def on_bloque_seleccionado(self, event=None):
        sel = self.combo_bloque.get()
        if sel not in self.bloques_map:
            return
        bloque_id = self.bloques_map[sel]

        # Cargar inspecciones del bloque
        inspecciones = db.obtener_inspecciones_por_bloque(bloque_id)
        self.inspecciones_map = {
            f"{i['fecha_visita']} - {i['inspector']}": i["id"] for i in inspecciones
        }
        self.combo_inspeccion["values"] = list(self.inspecciones_map.keys())
        if self.inspecciones_map:
            self.combo_inspeccion.current(0)

        # Cargar indicadores existentes
        self.cargar_tabla_indicadores(bloque_id)

    def cargar_tabla_indicadores(self, bloque_id):
        for item in self.tree_ind.get_children():
            self.tree_ind.delete(item)
        indicadores = db.obtener_indicadores_por_bloque(bloque_id)
        for ind in indicadores:
            self.tree_ind.insert("", "end", values=(
                ind.get("fecha_visita", ""),
                ind.get("microcuenca", "") or "",
                f"{ind.get('porcentaje_cobertura_vegetal', 0):.1f}",
                ind.get("tipo_cobertura_vegetal", "") or "",
                ind.get("vigor_cobertura_vegetal", "") or "",
                f"{ind['sobrevivencia_especies']:.1f}",
                f"{ind['longitud_zanjas_ejecutada']:.2f}",
                f"{ind['volumen_retencion_sedimentos']:.2f}",
            ))

    def guardar_indicadores(self):
        sel_bloque = self.combo_bloque.get()
        sel_insp = self.combo_inspeccion.get()

        if not sel_bloque or sel_bloque not in self.bloques_map:
            messagebox.showwarning("Validación", "Seleccione un bloque.")
            return
        if not sel_insp or sel_insp not in self.inspecciones_map:
            messagebox.showwarning("Validación", "Seleccione una inspección.")
            return

        try:
            pct_cobertura = float(self.ind_entries["pct_cobertura"].get().strip())
            sobrevivencia = float(self.ind_entries["sobrevivencia"].get().strip())
            longitud_zanjas = float(self.ind_entries["longitud_zanjas"].get().strip())
            vol_retencion = float(self.ind_entries["vol_retencion"].get().strip())
        except ValueError:
            messagebox.showwarning("Validación", "Los indicadores numéricos deben ser valores válidos.")
            return

        tipo_cobertura = self.combo_tipo_cobertura.get()
        vigor_cobertura = self.combo_vigor.get()
        microcuenca = self.combo_microcuenca.get()

        bloque_id = self.bloques_map[sel_bloque]
        inspeccion_id = self.inspecciones_map[sel_insp]

        try:
            db.insertar_indicadores(
                bloque_id=bloque_id,
                inspeccion_id=inspeccion_id,
                cobertura_vegetal_planificada=0,
                cobertura_vegetal_lograda=0,
                sobrevivencia_especies=sobrevivencia,
                longitud_zanjas_ejecutada=longitud_zanjas,
                volumen_retencion_sedimentos=vol_retencion,
                porcentaje_cobertura_vegetal=pct_cobertura,
                tipo_cobertura_vegetal=tipo_cobertura,
                vigor_cobertura_vegetal=vigor_cobertura,
                microcuenca=microcuenca,
            )
            messagebox.showinfo("Éxito", "Indicadores guardados correctamente.")
            self.cargar_tabla_indicadores(bloque_id)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar:\n{e}")

    def limpiar_formulario(self):
        for entry in self.ind_entries.values():
            entry.delete(0, tk.END)
            entry.insert(0, "0")
        self.combo_tipo_cobertura.set("")
        self.combo_vigor.set("")
        self.combo_microcuenca.set("")


# ── Tab 4: Reportes ──────────────────────────────────────────────────────

class TabReportes(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self._crear_widgets()
        self.cargar_combos()

    def _crear_widgets(self):
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Generación de Reportes",
                  style="Header.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 16))

        # Sección 1: Ficha PDF por bloque
        ttk.Label(frame, text="Ficha de Inspección (PDF)", font=("Segoe UI", 10, "bold")).grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(8, 4))

        ttk.Label(frame, text="Seleccione bloque:").grid(row=2, column=0, sticky="w", pady=4)
        self.combo_bloque_pdf = ttk.Combobox(frame, state="readonly", width=42)
        self.combo_bloque_pdf.grid(row=2, column=1, sticky="w", padx=6, pady=4)

        ttk.Button(frame, text="Generar Ficha PDF",
                   command=self.generar_pdf, style="Accent.TButton").grid(row=2, column=2, padx=8, pady=4)

        ttk.Separator(frame, orient="horizontal").grid(row=3, column=0, columnspan=3, sticky="ew", pady=14)

        # Sección 2: Resumen Excel
        ttk.Label(frame, text="Tabla Resumen (Excel)", font=("Segoe UI", 10, "bold")).grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(8, 4))

        ttk.Label(frame, text="Exporta todos los bloques con inspecciones e indicadores.").grid(
            row=5, column=0, columnspan=2, sticky="w", pady=4)

        ttk.Button(frame, text="Generar Resumen Excel",
                   command=self.generar_excel, style="Accent.TButton").grid(row=5, column=2, padx=8, pady=4)

        ttk.Separator(frame, orient="horizontal").grid(row=6, column=0, columnspan=3, sticky="ew", pady=14)

        # Sección 3: Excel ArcGIS
        ttk.Label(frame, text="Tabla UTM para ArcGIS (Excel)", font=("Segoe UI", 10, "bold")).grid(
            row=7, column=0, columnspan=3, sticky="w", pady=(8, 4))

        ttk.Label(frame, text="Columnas POINT_X/POINT_Y y campos compatibles con ArcGIS.").grid(
            row=8, column=0, columnspan=2, sticky="w", pady=4)

        ttk.Button(frame, text="Generar Excel ArcGIS",
                   command=self.generar_arcgis, style="Accent.TButton").grid(row=8, column=2, padx=8, pady=4)

        ttk.Separator(frame, orient="horizontal").grid(row=9, column=0, columnspan=3, sticky="ew", pady=14)

        # Estado
        self.label_estado = ttk.Label(frame, text="", wraplength=600, foreground=COLOR_EXITO)
        self.label_estado.grid(row=10, column=0, columnspan=3, sticky="w", pady=8)

    def cargar_combos(self):
        bloques = db.obtener_bloques()
        self.bloques_map = {f"{b['codigo']} - {b['tipo_intervencion']}": b["id"] for b in bloques}
        self.combo_bloque_pdf["values"] = list(self.bloques_map.keys())

    def generar_pdf(self):
        sel = self.combo_bloque_pdf.get()
        if not sel or sel not in self.bloques_map:
            messagebox.showwarning("Selección", "Seleccione un bloque para generar la ficha PDF.")
            return
        try:
            ruta = reports.generar_ficha_pdf(self.bloques_map[sel])
            self.label_estado.config(text=f"PDF generado: {ruta}", foreground=COLOR_EXITO)
            messagebox.showinfo("Reporte PDF", f"Ficha generada correctamente:\n{ruta}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar el PDF:\n{e}")

    def generar_excel(self):
        try:
            ruta = reports.generar_resumen_excel()
            self.label_estado.config(text=f"Excel generado: {ruta}", foreground=COLOR_EXITO)
            messagebox.showinfo("Reporte Excel", f"Resumen generado:\n{ruta}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar el Excel:\n{e}")

    def generar_arcgis(self):
        try:
            ruta = reports.generar_excel_arcgis()
            self.label_estado.config(text=f"Excel ArcGIS generado: {ruta}", foreground=COLOR_EXITO)
            messagebox.showinfo("ArcGIS", f"Archivo generado:\n{ruta}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar el archivo:\n{e}")
