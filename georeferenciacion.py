"""
IN Piura - Módulo de Georreferenciación
Visualización de bloques de intervención en mapa interactivo.
Conversión UTM ↔ Lat/Lon y herramientas de navegación espacial.
Cuenca Alta del Río Piura, Perú.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import math

import database as db

# ── Conversión UTM ↔ Lat/Lon (sin dependencia externa) ────────────────────

def utm_a_latlon(este, norte, zona=17, hemisferio="S"):
    """Convierte coordenadas UTM a Latitud/Longitud (WGS84).
    Implementación basada en las fórmulas de Karney para proyección
    transversa de Mercator."""
    # Constantes WGS84
    a = 6378137.0
    f = 1 / 298.257223563
    e2 = 2 * f - f ** 2
    e_prime2 = e2 / (1 - e2)
    k0 = 0.9996

    x = este - 500000.0
    y = norte
    if hemisferio.upper() == "S":
        y = y - 10000000.0

    lon0 = math.radians((zona - 1) * 6 - 180 + 3)

    M = y / k0
    mu = M / (a * (1 - e2 / 4 - 3 * e2 ** 2 / 64 - 5 * e2 ** 3 / 256))

    e1 = (1 - math.sqrt(1 - e2)) / (1 + math.sqrt(1 - e2))

    phi1 = (mu + (3 * e1 / 2 - 27 * e1 ** 3 / 32) * math.sin(2 * mu)
            + (21 * e1 ** 2 / 16 - 55 * e1 ** 4 / 32) * math.sin(4 * mu)
            + (151 * e1 ** 3 / 96) * math.sin(6 * mu)
            + (1097 * e1 ** 4 / 512) * math.sin(8 * mu))

    N1 = a / math.sqrt(1 - e2 * math.sin(phi1) ** 2)
    T1 = math.tan(phi1) ** 2
    C1 = e_prime2 * math.cos(phi1) ** 2
    R1 = a * (1 - e2) / (1 - e2 * math.sin(phi1) ** 2) ** 1.5
    D = x / (N1 * k0)

    lat = phi1 - (N1 * math.tan(phi1) / R1) * (
        D ** 2 / 2
        - (5 + 3 * T1 + 10 * C1 - 4 * C1 ** 2 - 9 * e_prime2) * D ** 4 / 24
        + (61 + 90 * T1 + 298 * C1 + 45 * T1 ** 2
           - 252 * e_prime2 - 3 * C1 ** 2) * D ** 6 / 720
    )

    lon = (D - (1 + 2 * T1 + C1) * D ** 3 / 6
           + (5 - 2 * C1 + 28 * T1 - 3 * C1 ** 2
              + 8 * e_prime2 + 24 * T1 ** 2) * D ** 5 / 120) / math.cos(phi1)

    lat_deg = math.degrees(lat)
    lon_deg = math.degrees(lon) + math.degrees(lon0)

    return lat_deg, lon_deg


def latlon_a_utm(lat, lon, zona=17):
    """Convierte Lat/Lon (WGS84) a coordenadas UTM."""
    a = 6378137.0
    f = 1 / 298.257223563
    e2 = 2 * f - f ** 2
    e_prime2 = e2 / (1 - e2)
    k0 = 0.9996

    lat_rad = math.radians(lat)
    lon_rad = math.radians(lon)
    lon0 = math.radians((zona - 1) * 6 - 180 + 3)

    N = a / math.sqrt(1 - e2 * math.sin(lat_rad) ** 2)
    T = math.tan(lat_rad) ** 2
    C = e_prime2 * math.cos(lat_rad) ** 2
    A = math.cos(lat_rad) * (lon_rad - lon0)

    M = a * (
        (1 - e2 / 4 - 3 * e2 ** 2 / 64 - 5 * e2 ** 3 / 256) * lat_rad
        - (3 * e2 / 8 + 3 * e2 ** 2 / 32 + 45 * e2 ** 3 / 1024) * math.sin(2 * lat_rad)
        + (15 * e2 ** 2 / 256 + 45 * e2 ** 3 / 1024) * math.sin(4 * lat_rad)
        - (35 * e2 ** 3 / 3072) * math.sin(6 * lat_rad)
    )

    este = k0 * N * (
        A + (1 - T + C) * A ** 3 / 6
        + (5 - 18 * T + T ** 2 + 72 * C - 58 * e_prime2) * A ** 5 / 120
    ) + 500000.0

    norte = k0 * (
        M + N * math.tan(lat_rad) * (
            A ** 2 / 2
            + (5 - T + 9 * C + 4 * C ** 2) * A ** 4 / 24
            + (61 - 58 * T + T ** 2 + 600 * C - 330 * e_prime2) * A ** 6 / 720
        )
    )

    if lat < 0:
        norte += 10000000.0

    hemisferio = "N" if lat >= 0 else "S"
    return este, norte, f"{zona}{hemisferio}"


# ── Colores del mapa ──────────────────────────────────────────────────────

COLORES_ESTADO = {
    "Pendiente": "#E74C3C",
    "En progreso": "#F39C12",
    "Verificado": "#27AE60",
}

COLORES_TIPO = {
    "Revegetación": "#27AE60",
    "Zanjas de infiltración": "#3498DB",
    "Terrazas de formación lenta": "#E67E22",
    "Diques de mampostería": "#9B59B6",
}


# ── Widget de Mapa con Canvas de tkinter ──────────────────────────────────

class MapaCanvas(tk.Canvas):
    """Mapa interactivo basado en Canvas de tkinter.
    Renderiza bloques de intervención como marcadores sobre un fondo
    con grilla de coordenadas. Soporta zoom y desplazamiento."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg="#D6EAF8", highlightthickness=1,
                         highlightbackground="#BDC3C7", **kwargs)

        # Estado del mapa
        self.bloques = []
        self.marcadores = {}  # id -> item_ids
        self.bloque_seleccionado = None
        self.callback_seleccion = None

        # Centro del mapa (Piura, Perú - aproximado)
        self.centro_lat = -5.05
        self.centro_lon = -79.70
        self.zoom = 10  # nivel 1-18
        self.escala = 5000.0  # pixeles por grado

        # Drag
        self._drag_data = {"x": 0, "y": 0}

        # Bindings
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<B1-Motion>", self._on_drag)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<MouseWheel>", self._on_scroll)
        self.bind("<Button-4>", lambda e: self._zoom_in())
        self.bind("<Button-5>", lambda e: self._zoom_out())
        self.bind("<Configure>", lambda e: self.redibujar())

    def set_callback_seleccion(self, callback):
        self.callback_seleccion = callback

    def cargar_bloques(self, bloques):
        """Carga la lista de bloques y ajusta la vista."""
        self.bloques = bloques
        if bloques:
            lats, lons = [], []
            for b in bloques:
                try:
                    zona_num = int(b["utm_zona"].replace("S", "").replace("N", ""))
                    hemisferio = "S" if "S" in b["utm_zona"] else "N"
                    lat, lon = utm_a_latlon(b["utm_este"], b["utm_norte"],
                                            zona_num, hemisferio)
                    lats.append(lat)
                    lons.append(lon)
                except (ValueError, KeyError):
                    continue
            if lats:
                self.centro_lat = sum(lats) / len(lats)
                self.centro_lon = sum(lons) / len(lons)
                if len(lats) > 1:
                    rango_lat = max(lats) - min(lats)
                    rango_lon = max(lons) - min(lons)
                    rango = max(rango_lat, rango_lon, 0.01)
                    self.escala = min(self.winfo_width(), self.winfo_height()) / (rango * 1.5) if rango > 0 else 5000.0
                    self.escala = max(500, min(self.escala, 50000))
        self.redibujar()

    def _latlon_a_pixel(self, lat, lon):
        """Convierte lat/lon a coordenadas de pixel en el canvas."""
        w = self.winfo_width()
        h = self.winfo_height()
        cx, cy = w / 2, h / 2

        dx = (lon - self.centro_lon) * self.escala
        dy = -(lat - self.centro_lat) * self.escala  # Y invertido

        return cx + dx, cy + dy

    def _pixel_a_latlon(self, px, py):
        """Convierte coordenadas de pixel a lat/lon."""
        w = self.winfo_width()
        h = self.winfo_height()
        cx, cy = w / 2, h / 2

        lon = self.centro_lon + (px - cx) / self.escala
        lat = self.centro_lat - (py - cy) / self.escala

        return lat, lon

    def redibujar(self):
        """Redibuja todo el mapa."""
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()

        if w < 10 or h < 10:
            return

        # Fondo degradado
        for i in range(0, h, 4):
            ratio = i / max(h, 1)
            r = int(214 + (240 - 214) * ratio)
            g = int(234 + (248 - 234) * ratio)
            b_color = int(248 + (255 - 248) * ratio)
            color = f"#{r:02x}{g:02x}{b_color:02x}"
            self.create_rectangle(0, i, w, i + 4, fill=color, outline="")

        # Grilla de coordenadas
        self._dibujar_grilla(w, h)

        # Marcadores de bloques
        self.marcadores = {}
        for bloque in self.bloques:
            self._dibujar_marcador(bloque)

        # Leyenda
        self._dibujar_leyenda(w, h)

        # Escala
        self._dibujar_escala(w, h)

        # Info de coordenadas del centro
        lat_str = f"{abs(self.centro_lat):.4f}{'S' if self.centro_lat < 0 else 'N'}"
        lon_str = f"{abs(self.centro_lon):.4f}{'W' if self.centro_lon < 0 else 'E'}"
        self.create_text(w - 10, 12, text=f"{lat_str}, {lon_str}",
                         anchor="ne", font=("Consolas", 8), fill="#7F8C8D")

    def _dibujar_grilla(self, w, h):
        """Dibuja una grilla de referencia."""
        # Calcular intervalo de grilla apropiado
        grado_en_px = self.escala
        if grado_en_px > 2000:
            intervalo = 0.01
        elif grado_en_px > 500:
            intervalo = 0.05
        elif grado_en_px > 100:
            intervalo = 0.1
        else:
            intervalo = 0.5

        # Esquinas del viewport
        lat_top, lon_left = self._pixel_a_latlon(0, 0)
        lat_bottom, lon_right = self._pixel_a_latlon(w, h)

        # Líneas verticales (longitud)
        lon_start = math.floor(lon_left / intervalo) * intervalo
        lon = lon_start
        while lon <= lon_right:
            px, _ = self._latlon_a_pixel(0, lon)
            self.create_line(px, 0, px, h, fill="#BDC3C7", dash=(2, 4))
            self.create_text(px, h - 5, text=f"{lon:.3f}",
                             anchor="s", font=("Consolas", 7), fill="#95A5A6")
            lon += intervalo

        # Líneas horizontales (latitud)
        lat_start = math.floor(lat_bottom / intervalo) * intervalo
        lat = lat_start
        while lat <= lat_top:
            _, py = self._latlon_a_pixel(lat, 0)
            self.create_line(0, py, w, py, fill="#BDC3C7", dash=(2, 4))
            self.create_text(5, py, text=f"{lat:.3f}",
                             anchor="w", font=("Consolas", 7), fill="#95A5A6")
            lat += intervalo

    def _dibujar_marcador(self, bloque):
        """Dibuja un marcador para un bloque."""
        try:
            zona_num = int(bloque["utm_zona"].replace("S", "").replace("N", ""))
            hemisferio = "S" if "S" in bloque["utm_zona"] else "N"
            lat, lon = utm_a_latlon(bloque["utm_este"], bloque["utm_norte"],
                                    zona_num, hemisferio)
        except (ValueError, KeyError):
            return

        px, py = self._latlon_a_pixel(lat, lon)

        # Determinar color según estado
        color = COLORES_ESTADO.get(bloque.get("estado", ""), "#95A5A6")
        borde_color = COLORES_TIPO.get(bloque.get("tipo_intervencion", ""), "#2C3E50")

        radio = 8
        es_seleccionado = (bloque["id"] == self.bloque_seleccionado)
        if es_seleccionado:
            radio = 12
            # Halo de selección
            self.create_oval(px - radio - 4, py - radio - 4,
                             px + radio + 4, py + radio + 4,
                             fill="", outline="#3498DB", width=3)

        # Sombra
        items = []
        items.append(self.create_oval(px - radio + 1, py - radio + 1,
                                      px + radio + 1, py + radio + 1,
                                      fill="#7F8C8D", outline=""))

        # Marcador principal
        items.append(self.create_oval(px - radio, py - radio,
                                      px + radio, py + radio,
                                      fill=color, outline=borde_color, width=2))

        # Punto central
        items.append(self.create_oval(px - 2, py - 2, px + 2, py + 2,
                                      fill="white", outline=""))

        # Etiqueta
        items.append(self.create_text(px, py - radio - 8,
                                      text=bloque["codigo"],
                                      font=("Segoe UI", 8, "bold"),
                                      fill="#2C3E50"))

        # Guardar referencia
        self.marcadores[bloque["id"]] = {
            "items": items,
            "lat": lat,
            "lon": lon,
            "px": px,
            "py": py,
        }

        # Bind click en marcador
        for item in items:
            self.tag_bind(item, "<Button-1>",
                          lambda e, bid=bloque["id"]: self._on_click_marcador(bid))

    def _dibujar_leyenda(self, w, h):
        """Dibuja la leyenda del mapa."""
        x0, y0 = 10, 10
        padding = 6
        line_h = 16

        # Fondo de leyenda
        total_lines = len(COLORES_ESTADO) + len(COLORES_TIPO) + 2
        box_h = total_lines * line_h + padding * 2
        box_w = 180
        self.create_rectangle(x0, y0, x0 + box_w, y0 + box_h,
                              fill="white", outline="#BDC3C7", stipple="")

        y = y0 + padding
        self.create_text(x0 + padding, y, text="Estado:", anchor="w",
                         font=("Segoe UI", 8, "bold"), fill="#2C3E50")
        y += line_h

        for estado, color in COLORES_ESTADO.items():
            self.create_oval(x0 + padding, y, x0 + padding + 10, y + 10,
                             fill=color, outline="#2C3E50")
            self.create_text(x0 + padding + 16, y + 5, text=estado,
                             anchor="w", font=("Segoe UI", 7), fill="#2C3E50")
            y += line_h

        y += 4
        self.create_text(x0 + padding, y, text="Tipo (borde):", anchor="w",
                         font=("Segoe UI", 8, "bold"), fill="#2C3E50")
        y += line_h

        for tipo, color in COLORES_TIPO.items():
            self.create_rectangle(x0 + padding, y, x0 + padding + 10, y + 10,
                                  fill="", outline=color, width=2)
            self.create_text(x0 + padding + 16, y + 5, text=tipo,
                             anchor="w", font=("Segoe UI", 7), fill="#2C3E50")
            y += line_h

    def _dibujar_escala(self, w, h):
        """Dibuja barra de escala aproximada."""
        # 1 grado de latitud ≈ 111 km
        km_por_grado = 111.0
        px_por_km = self.escala / km_por_grado

        # Elegir escala legible
        escalas_posibles = [0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 50, 100]
        barra_km = 1
        for es in escalas_posibles:
            if es * px_por_km >= 40 and es * px_por_km <= 200:
                barra_km = es
                break

        barra_px = barra_km * px_por_km
        x0 = w - 20 - barra_px
        y0 = h - 25

        self.create_line(x0, y0, x0 + barra_px, y0, fill="#2C3E50", width=2)
        self.create_line(x0, y0 - 4, x0, y0 + 4, fill="#2C3E50", width=2)
        self.create_line(x0 + barra_px, y0 - 4, x0 + barra_px, y0 + 4,
                         fill="#2C3E50", width=2)

        if barra_km >= 1:
            texto_escala = f"{barra_km:.0f} km"
        else:
            texto_escala = f"{barra_km * 1000:.0f} m"
        self.create_text(x0 + barra_px / 2, y0 - 8, text=texto_escala,
                         font=("Segoe UI", 8), fill="#2C3E50")

    def _on_click_marcador(self, bloque_id):
        """Maneja click en un marcador."""
        self.bloque_seleccionado = bloque_id
        self.redibujar()
        if self.callback_seleccion:
            self.callback_seleccion(bloque_id)

    def _on_press(self, event):
        self._drag_data = {"x": event.x, "y": event.y, "moved": False}

    def _on_drag(self, event):
        dx = event.x - self._drag_data["x"]
        dy = event.y - self._drag_data["y"]
        self._drag_data["moved"] = True

        self.centro_lon -= dx / self.escala
        self.centro_lat += dy / self.escala

        self._drag_data["x"] = event.x
        self._drag_data["y"] = event.y
        self.redibujar()

    def _on_release(self, event):
        pass

    def _on_scroll(self, event):
        if event.delta > 0:
            self._zoom_in()
        else:
            self._zoom_out()

    def _zoom_in(self):
        self.escala *= 1.3
        self.escala = min(self.escala, 100000)
        self.redibujar()

    def _zoom_out(self):
        self.escala /= 1.3
        self.escala = max(self.escala, 100)
        self.redibujar()

    def centrar_en_bloque(self, bloque_id):
        """Centra el mapa en un bloque específico."""
        self.bloque_seleccionado = bloque_id
        if bloque_id in self.marcadores:
            m = self.marcadores[bloque_id]
            self.centro_lat = m["lat"]
            self.centro_lon = m["lon"]
        self.redibujar()


# ── Tab Georreferenciación ────────────────────────────────────────────────

class TabGeorreferenciacion(ttk.Frame):
    """Pestaña de georreferenciación con mapa interactivo y herramientas
    de conversión de coordenadas."""

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self._crear_widgets()
        self.cargar_datos()

    def _crear_widgets(self):
        # ── Panel superior: controles ──
        control_frame = ttk.Frame(self, padding=6)
        control_frame.pack(fill="x")

        ttk.Label(control_frame, text="Georreferenciación",
                  style="Header.TLabel").pack(side="left", padx=(4, 16))

        # Filtros
        ttk.Label(control_frame, text="Filtrar estado:").pack(side="left", padx=(10, 4))
        self.combo_filtro_estado = ttk.Combobox(
            control_frame,
            values=["Todos", "Pendiente", "En progreso", "Verificado"],
            state="readonly", width=14
        )
        self.combo_filtro_estado.pack(side="left", padx=2)
        self.combo_filtro_estado.set("Todos")
        self.combo_filtro_estado.bind("<<ComboboxSelected>>", lambda e: self.aplicar_filtro())

        ttk.Label(control_frame, text="Filtrar tipo:").pack(side="left", padx=(10, 4))
        self.combo_filtro_tipo = ttk.Combobox(
            control_frame,
            values=["Todos", "Revegetación", "Zanjas de infiltración",
                    "Terrazas de formación lenta", "Diques de mampostería"],
            state="readonly", width=22
        )
        self.combo_filtro_tipo.pack(side="left", padx=2)
        self.combo_filtro_tipo.set("Todos")
        self.combo_filtro_tipo.bind("<<ComboboxSelected>>", lambda e: self.aplicar_filtro())

        # Botones de zoom
        btn_zoom_frame = ttk.Frame(control_frame)
        btn_zoom_frame.pack(side="right", padx=4)
        ttk.Button(btn_zoom_frame, text="+", width=3,
                   command=self._zoom_in).pack(side="left", padx=1)
        ttk.Button(btn_zoom_frame, text="-", width=3,
                   command=self._zoom_out).pack(side="left", padx=1)
        ttk.Button(btn_zoom_frame, text="Ajustar vista", width=12,
                   command=self.ajustar_vista).pack(side="left", padx=4)

        # ── Panel principal dividido ──
        paned = ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=4, pady=4)

        # ── Panel izquierdo: Mapa ──
        mapa_frame = ttk.Frame(paned)
        paned.add(mapa_frame, weight=3)

        self.mapa = MapaCanvas(mapa_frame)
        self.mapa.pack(fill="both", expand=True)
        self.mapa.set_callback_seleccion(self._on_seleccionar_marcador)

        # ── Panel derecho: Info y herramientas ──
        info_frame = ttk.Frame(paned, padding=8)
        paned.add(info_frame, weight=1)

        # Sección: Info del bloque seleccionado
        ttk.Label(info_frame, text="Bloque Seleccionado",
                  style="Header.TLabel").pack(anchor="w", pady=(0, 6))

        self.info_text = tk.Text(info_frame, width=32, height=14,
                                 font=("Consolas", 9), state="disabled",
                                 bg="#FDFEFE", wrap="word")
        self.info_text.pack(fill="x", pady=(0, 8))

        ttk.Separator(info_frame).pack(fill="x", pady=8)

        # Sección: Conversor de coordenadas
        ttk.Label(info_frame, text="Conversor UTM ↔ Lat/Lon",
                  style="Header.TLabel").pack(anchor="w", pady=(0, 6))

        conv_frame = ttk.Frame(info_frame)
        conv_frame.pack(fill="x")

        ttk.Label(conv_frame, text="UTM Este:").grid(row=0, column=0, sticky="w", pady=2)
        self.entry_conv_este = ttk.Entry(conv_frame, width=16)
        self.entry_conv_este.grid(row=0, column=1, padx=4, pady=2)

        ttk.Label(conv_frame, text="UTM Norte:").grid(row=1, column=0, sticky="w", pady=2)
        self.entry_conv_norte = ttk.Entry(conv_frame, width=16)
        self.entry_conv_norte.grid(row=1, column=1, padx=4, pady=2)

        ttk.Label(conv_frame, text="Zona UTM:").grid(row=2, column=0, sticky="w", pady=2)
        self.entry_conv_zona = ttk.Entry(conv_frame, width=16)
        self.entry_conv_zona.grid(row=2, column=1, padx=4, pady=2)
        self.entry_conv_zona.insert(0, "17S")

        ttk.Button(conv_frame, text="UTM → Lat/Lon",
                   command=self._convertir_utm_a_ll).grid(
            row=3, column=0, columnspan=2, pady=6, sticky="ew")

        ttk.Separator(conv_frame).grid(row=4, column=0, columnspan=2,
                                        sticky="ew", pady=4)

        ttk.Label(conv_frame, text="Latitud:").grid(row=5, column=0, sticky="w", pady=2)
        self.entry_conv_lat = ttk.Entry(conv_frame, width=16)
        self.entry_conv_lat.grid(row=5, column=1, padx=4, pady=2)

        ttk.Label(conv_frame, text="Longitud:").grid(row=6, column=0, sticky="w", pady=2)
        self.entry_conv_lon = ttk.Entry(conv_frame, width=16)
        self.entry_conv_lon.grid(row=6, column=1, padx=4, pady=2)

        ttk.Button(conv_frame, text="Lat/Lon → UTM",
                   command=self._convertir_ll_a_utm).grid(
            row=7, column=0, columnspan=2, pady=6, sticky="ew")

        self.label_resultado_conv = ttk.Label(info_frame, text="",
                                              wraplength=250, foreground="#2C3E50")
        self.label_resultado_conv.pack(anchor="w", pady=4)

        ttk.Separator(info_frame).pack(fill="x", pady=8)

        # Estadísticas rápidas
        ttk.Label(info_frame, text="Resumen Espacial",
                  style="Header.TLabel").pack(anchor="w", pady=(0, 6))
        self.label_resumen = ttk.Label(info_frame, text="", wraplength=250)
        self.label_resumen.pack(anchor="w")

    def cargar_datos(self):
        """Carga los bloques desde la base de datos y los muestra en el mapa."""
        self.todos_bloques = db.obtener_bloques()
        self.bloques_filtrados = list(self.todos_bloques)
        self.mapa.cargar_bloques(self.bloques_filtrados)
        self._actualizar_resumen()

    def aplicar_filtro(self):
        """Aplica los filtros de estado y tipo."""
        estado = self.combo_filtro_estado.get()
        tipo = self.combo_filtro_tipo.get()

        self.bloques_filtrados = []
        for b in self.todos_bloques:
            if estado != "Todos" and b.get("estado") != estado:
                continue
            if tipo != "Todos" and b.get("tipo_intervencion") != tipo:
                continue
            self.bloques_filtrados.append(b)

        self.mapa.cargar_bloques(self.bloques_filtrados)
        self._actualizar_resumen()

    def ajustar_vista(self):
        """Re-centra el mapa para mostrar todos los bloques."""
        self.mapa.cargar_bloques(self.bloques_filtrados)

    def _zoom_in(self):
        self.mapa._zoom_in()

    def _zoom_out(self):
        self.mapa._zoom_out()

    def _on_seleccionar_marcador(self, bloque_id):
        """Muestra información del bloque seleccionado."""
        bloque = db.obtener_bloque_por_id(bloque_id)
        if not bloque:
            return

        try:
            zona_num = int(bloque["utm_zona"].replace("S", "").replace("N", ""))
            hemisferio = "S" if "S" in bloque["utm_zona"] else "N"
            lat, lon = utm_a_latlon(bloque["utm_este"], bloque["utm_norte"],
                                    zona_num, hemisferio)
            coord_str = f"Lat: {lat:.6f}\nLon: {lon:.6f}"
        except (ValueError, KeyError):
            coord_str = "Error en conversión"

        # Obtener inspecciones
        inspecciones = db.obtener_inspecciones_por_bloque(bloque_id)
        n_insp = len(inspecciones)
        ultimo_avance = inspecciones[0]["avance_fisico"] if inspecciones else 0

        info = (
            f"Código: {bloque['codigo']}\n"
            f"Tipo: {bloque['tipo_intervencion']}\n"
            f"Cuenca: {bloque['cuenca']}\n"
            f"Distrito: {bloque['distrito']}\n"
            f"Estado: {bloque['estado']}\n"
            f"Área: {bloque['area_hectareas']:.4f} ha\n"
            f"─────────────────\n"
            f"UTM Este: {bloque['utm_este']:.2f}\n"
            f"UTM Norte: {bloque['utm_norte']:.2f}\n"
            f"Zona UTM: {bloque['utm_zona']}\n"
            f"{coord_str}\n"
            f"─────────────────\n"
            f"Inspecciones: {n_insp}\n"
            f"Último avance: {ultimo_avance:.1f}%"
        )

        self.info_text.config(state="normal")
        self.info_text.delete("1.0", tk.END)
        self.info_text.insert("1.0", info)
        self.info_text.config(state="disabled")

    def _convertir_utm_a_ll(self):
        """Convierte UTM a Lat/Lon."""
        try:
            este = float(self.entry_conv_este.get().strip())
            norte = float(self.entry_conv_norte.get().strip())
            zona_str = self.entry_conv_zona.get().strip()
            zona_num = int(zona_str.replace("S", "").replace("N", ""))
            hemisferio = "S" if "S" in zona_str.upper() else "N"

            lat, lon = utm_a_latlon(este, norte, zona_num, hemisferio)

            self.entry_conv_lat.delete(0, tk.END)
            self.entry_conv_lat.insert(0, f"{lat:.8f}")
            self.entry_conv_lon.delete(0, tk.END)
            self.entry_conv_lon.insert(0, f"{lon:.8f}")

            self.label_resultado_conv.config(
                text=f"Lat: {lat:.6f}, Lon: {lon:.6f}",
                foreground="#27AE60")
        except ValueError:
            self.label_resultado_conv.config(
                text="Error: valores numéricos inválidos",
                foreground="#E74C3C")

    def _convertir_ll_a_utm(self):
        """Convierte Lat/Lon a UTM."""
        try:
            lat = float(self.entry_conv_lat.get().strip())
            lon = float(self.entry_conv_lon.get().strip())

            este, norte, zona = latlon_a_utm(lat, lon)

            self.entry_conv_este.delete(0, tk.END)
            self.entry_conv_este.insert(0, f"{este:.2f}")
            self.entry_conv_norte.delete(0, tk.END)
            self.entry_conv_norte.insert(0, f"{norte:.2f}")
            self.entry_conv_zona.delete(0, tk.END)
            self.entry_conv_zona.insert(0, zona)

            self.label_resultado_conv.config(
                text=f"UTM: {este:.2f} E, {norte:.2f} N ({zona})",
                foreground="#27AE60")
        except ValueError:
            self.label_resultado_conv.config(
                text="Error: valores numéricos inválidos",
                foreground="#E74C3C")

    def _actualizar_resumen(self):
        """Actualiza las estadísticas del panel."""
        total = len(self.bloques_filtrados)
        pendientes = sum(1 for b in self.bloques_filtrados if b.get("estado") == "Pendiente")
        en_prog = sum(1 for b in self.bloques_filtrados if b.get("estado") == "En progreso")
        verificados = sum(1 for b in self.bloques_filtrados if b.get("estado") == "Verificado")

        area_total = sum(b.get("area_hectareas", 0) for b in self.bloques_filtrados)

        # Calcular extensión geográfica
        lats, lons = [], []
        for b in self.bloques_filtrados:
            try:
                zona_num = int(b["utm_zona"].replace("S", "").replace("N", ""))
                hemisferio = "S" if "S" in b["utm_zona"] else "N"
                lat, lon = utm_a_latlon(b["utm_este"], b["utm_norte"],
                                        zona_num, hemisferio)
                lats.append(lat)
                lons.append(lon)
            except (ValueError, KeyError):
                continue

        if lats:
            extension = f"Extensión: {min(lats):.4f} a {max(lats):.4f} lat\n"
            extension += f"           {min(lons):.4f} a {max(lons):.4f} lon"
        else:
            extension = "Sin datos geográficos"

        texto = (
            f"Bloques mostrados: {total}\n"
            f"  Pendientes: {pendientes}\n"
            f"  En progreso: {en_prog}\n"
            f"  Verificados: {verificados}\n"
            f"Área total: {area_total:.4f} ha\n"
            f"{extension}"
        )
        self.label_resumen.config(text=texto)
