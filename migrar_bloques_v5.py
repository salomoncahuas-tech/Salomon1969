"""
Script de migracion: Reemplaza los bloques existentes en la BD por los 124 bloques V5.
ATENCION: Elimina TODOS los registros vinculados (inspecciones, diagnosticos, etc.)

Uso:
    python migrar_bloques_v5.py
"""

import sys
import os
import psycopg2
from datetime import datetime

# ── Leer DATABASE_URL desde secrets.toml ────────────────────────────────────
def leer_database_url():
    secrets_path = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")
    if not os.path.exists(secrets_path):
        print(f"ERROR: No se encontro {secrets_path}")
        print("Ingresa la DATABASE_URL manualmente:")
        return input("DATABASE_URL: ").strip()
    with open(secrets_path, "r") as f:
        for line in f:
            if line.strip().startswith("DATABASE_URL"):
                return line.split("=", 1)[1].strip().strip('"').strip("'")
    print("ERROR: DATABASE_URL no encontrada en secrets.toml")
    sys.exit(1)

# ── 124 Bloques V5 ──────────────────────────────────────────────────────────
# Fuente: Reporte_Bloques_V5_25abril.xlsx (centroides UTM Zona 17S WGS84)
# (codigo, microcuenca, area_ha, provincia, distrito, utm_este, utm_norte, msavi_2024)
BLOQUES_V5 = [
    ("56",       "C1096-Q9564",      77.666, "Ayabaca",       "Frias",                         621613, 9452884, 0.677658),
    ("57",       "C1086-Q9576",      25.668, "Morropon",      "Chalaco",                       631723, 9439702, 0.603055),
    ("59",       "C1086-Q9570",      13.554, "Morropon",      "Chalaco",                       631195, 9441777, 0.614739),
    ("58",       "C1086-Q9569",      11.475, "Morropon",      "Santa Catalina de Mossa",       619094, 9426406, 0.420325),
    ("61",       "C1081-Q9590",      26.826, "Huancabamba",   "Lalaquiz",                      646236, 9425052, 0.611186),
    ("60",       "C1081-Q9591",      40.132, "Huancabamba",   "Canchaque",                     653850, 9426236, 0.67447),
    ("64",       "C1081-Q9591",      35.774, "Huancabamba",   "Huancabamba",                   657819, 9428541, 0.525448),
    ("66",       "C1081-Q9591",     102.335, "Huancabamba",   "Canchaque",                     652566, 9417311, 0.640254),
    ("63",       "C1081-Q9583",      50.459, "Huancabamba",   "Canchaque",                     655059, 9413845, 0.603582),
    ("75",       "C1076-Q9587",      15.176, "Huancabamba",   "San Miguel de El Faique",       657722, 9401669, 0.630888),
    ("67",       "C1076-Q9586",      13.184, "Huancabamba",   "San Miguel de El Faique",       654207, 9394210, 0.605293),
    ("70",       "C1076-Q9592",      22.388, "Huancabamba",   "Huarmaca",                      666781, 9395573, 0.622295),
    ("68",       "C1076-Q9592",      21.291, "Huancabamba",   "Huarmaca",                      664225, 9396339, 0.482981),
    ("73",       "C1076-Q9592",       54.29, "Huancabamba",   "Huarmaca",                      658547, 9395594, 0.554974),
    ("69",       "C1076-Q9592",      19.425, "Huancabamba",   "Huarmaca",                      663291, 9392332, 0.448926),
    ("72",       "C1076-Q9592",      12.398, "Huancabamba",   "Huarmaca",                      662236, 9391211, 0.518477),
    ("71",       "C1076-Q9593",      10.673, "Huancabamba",   "Huarmaca",                      652104, 9373045, 0.557066),
    ("80",       "C1076-Q9593",       28.83, "Huancabamba",   "Huarmaca",                      657160, 9373050, 0.593069),
    ("74",       "C1076-Q9588",      29.226, "Huancabamba",   "Huarmaca",                      652449, 9389179, 0.356463),
    ("76",       "C1076-Q9593",      19.615, "Huancabamba",   "Huarmaca",                      657922, 9377477, 0.492462),
    ("79",       "C1076-Q9587",      94.812, "Huancabamba",   "Canchaque",                     650305, 9402785, 0.547879),
    ("77",       "C1081-Q9583",      17.821, "Morropon",      "San Juan de Bigote",            638648, 9411370, 0.272488),
    ("82",       "C1086-Q9570",      35.661, "Morropon",      "Santo Domingo",                 621452, 9443855, 0.62835),
    ("M17B1",    "C1096-Q9556",      42.449, "Morropon",      "Chulucanas",                    600644, 9440436, 0.332263),
    ("32",       "C1096-Q9556",       35.83, "Morropon",      "Chulucanas",                    599927, 9442604, 0.371619),
    ("M22B1",    "C1076-Q9586",       55.17, "Huancabamba",   "San Miguel de El Faique",       664082, 9404172, 0.594025),
    ("M6B2-1",   "C1077-Q9566",     514.359, "Morropon",      "Buenos Aires",                  612864, 9413345, 0.3381),
    ("M6B2-2",   "C1077-Q9566",     333.011, "Morropon",      "Buenos Aires",                  611196, 9416725, 0.302636),
    ("M6B10",    "C1077-Q9566",      249.89, "Morropon",      "Buenos Aires",                  618299, 9423119, 0.30989),
    ("M17B4",    "C1096-Q9556",     124.709, "Morropon",      "Chulucanas",                    604926, 9444327, 0.423543),
    ("M10B4",    "C1096-Q9547",      68.887, "Morropon",      "Chulucanas",                    596773, 9448273, 0.313301),
    ("M12B1",    "C1076-Q9588",     268.393, "Huancabamba",   "Huarmaca",                      644546, 9389106, 0.34753),
    ("M3B3",     "C1081-Q9582",      84.825, "Morropon",      "San Juan de Bigote",            633015, 9411483, 0.298483),
    ("M17B7",    "C1096-Q9556",     258.837, "Ayabaca",       "Frias",                         608472, 9448056, 0.508715),
    ("M18B5",    "C1077-Q9579",     197.352, "Morropon",      "Salitral",                      621551, 9412989, 0.356522),
    ("M28B4",    "C1086-Q9575",     283.428, "Morropon",      "Yamango",                       626078, 9427110, 0.42751),
    ("M8B2",     "C1096-Q9558",      35.347, "Morropon",      "Santo Domingo",                 617596, 9436227, 0.530453),
    ("M4B3",     "C1076-Q9589",     234.069, "Huancabamba",   "Huarmaca",                      643474, 9385620, 0.322153),
    ("M4B4",     "C1076-Q9589",     289.914, "Huancabamba",   "Huarmaca",                      642775, 9388574, 0.319788),
    ("M3B9",     "C1081-Q9582",     294.204, "Morropon",      "Salitral",                      629338, 9413305, 0.321386),
    ("M3B8",     "C1081-Q9582",     565.584, "Morropon",      "San Juan de Bigote",            632119, 9415141, 0.293351),
    ("M3B7",     "C1081-Q9582",     122.946, "Morropon",      "San Juan de Bigote",            638309, 9415217, 0.370678),
    ("M3B5",     "C1081-Q9582",      52.743, "Morropon",      "San Juan de Bigote",            635887, 9411738, 0.283016),
    ("M3B6",     "C1081-Q9582",      60.352, "Morropon",      "San Juan de Bigote",            637918, 9412134, 0.342535),
    ("M11B3",    "C1081-Q9583",     106.027, "Morropon",      "San Juan de Bigote",            642079, 9413823, 0.395609),
    ("M18B3",    "C1077-Q9579",     373.807, "Morropon",      "Salitral",                      625311, 9412195, 0.369424),
    ("M17B6",    "C1096-Q9556",     106.538, "Ayabaca",       "Frias",                         607036, 9446211, 0.446441),
    ("M28B2",    "C1086-Q9575",      90.197, "Morropon",      "Yamango",                       636249, 9425665, 0.475342),
    ("M2B5",     "C1076-Q9584",      30.909, "Morropon",      "Salitral",                      639987, 9394392, 0.304733),
    ("M17B5",    "C1096-Q9556",      74.028, "Ayabaca",       "Frias",                         608800, 9445585, 0.54817),
    ("M28B3",    "C1086-Q9575",      58.292, "Morropon",      "Yamango",                       634061, 9426778, 0.507529),
    ("33",       "C1081-Q9590",      29.681, "Huancabamba",   "Lalaquiz",                      646466, 9431126, 0.71923),
    ("50",       "C1081-Q9590",      19.266, "Huancabamba",   "Lalaquiz",                      647350, 9427252, 0.68307),
    ("38",       "C1081-Q9590",      44.201, "Huancabamba",   "Lalaquiz",                      647123, 9427602, 0.672487),
    ("34",       "C1081-Q9590",       28.45, "Huancabamba",   "Lalaquiz",                      646483, 9426305, 0.648975),
    ("26",       "C1081-Q9591",      46.962, "Huancabamba",   "Lalaquiz",                      649472, 9425705, 0.673303),
    ("37",       "C1081-Q9591",      35.491, "Huancabamba",   "Lalaquiz",                      649081, 9424506, 0.658846),
    ("28",       "C1081-Q9591",      52.966, "Huancabamba",   "Canchaque",                     654188, 9422849, 0.639425),
    ("40",       "C1081-Q9591",      28.523, "Huancabamba",   "Canchaque",                     654137, 9424425, 0.640099),
    ("21",       "C1081-Q9591",      84.227, "Huancabamba",   "Canchaque",                     654604, 9417736, 0.620078),
    ("24",       "C1081-Q9583",      90.159, "Huancabamba",   "Canchaque",                     652191, 9414358, 0.674839),
    ("47",       "C1076-Q9586",      13.551, "Huancabamba",   "San Miguel de El Faique",       659649, 9401436, 0.593628),
    ("43",       "C1076-Q9586",      23.171, "Huancabamba",   "San Miguel de El Faique",       656489, 9395403, 0.551525),
    ("35",       "C1076-Q9586",      30.804, "Huancabamba",   "San Miguel de El Faique",       654607, 9395038, 0.624477),
    ("44",       "C1076-Q9586",      28.234, "Huancabamba",   "San Miguel de El Faique",       663637, 9400618, 0.511324),
    ("53",       "C1076-Q9592",      13.065, "Huancabamba",   "Huarmaca",                      666537, 9395918, 0.596277),
    ("41",       "C1076-Q9592",      64.923, "Huancabamba",   "Huarmaca",                      660525, 9394327, 0.592669),
    ("19",       "C1076-Q9592",      82.088, "Huancabamba",   "Huarmaca",                      665942, 9391144, 0.537184),
    ("31",       "C1076-Q9592",      47.581, "Huancabamba",   "Huarmaca",                      663334, 9392820, 0.435805),
    ("20",       "C1076-Q9593",      81.334, "Huancabamba",   "Huarmaca",                      649847, 9376260, 0.48513),
    ("51",       "C1076-Q9593",      22.927, "Huancabamba",   "Huarmaca",                      656291, 9379096, 0.523788),
    ("29",       "C1076-Q9592",      40.104, "Huancabamba",   "Huarmaca",                      662123, 9397575, 0.548192),
    ("1",        "C1076-Q9584",    1370.958, "Morropon",      "Salitral",                      645575, 9393500, 0.454493),
    ("7",        "C1076-Q9584",     173.428, "Morropon",      "Salitral",                      639728, 9396836, 0.375288),
    ("5",        "C1077-Q9566",     295.726, "Morropon",      "Buenos Aires",                  617860, 9413182, 0.3394),
    ("2",        "C1096-Q9558",     330.599, "Morropon",      "Chulucanas",                    603653, 9439351, 0.358601),
    ("25",       "C1096-Q9547",      53.484, "Morropon",      "Chulucanas",                    596578, 9442453, 0.256543),
    ("9",        "C1076-Q9586",     109.861, "Huancabamba",   "San Miguel de El Faique",       656578, 9397762, 0.574527),
    ("4",        "C1076-Q9588",     230.098, "Huancabamba",   "Huarmaca",                      656549, 9382009, 0.660111),
    ("18",       "C1076-Q9592",     103.617, "Huancabamba",   "Huarmaca",                      656445, 9393180, 0.640219),
    ("10",       "C1076-Q9593",     150.214, "Huancabamba",   "Huarmaca",                      650830, 9382303, 0.546492),
    ("49",       "C1086-Q9575",      15.839, "Morropon",      "Yamango",                       642334, 9438263, 0.686036),
    ("55",       "C1086-Q9575",       0.859, "Morropon",      "Yamango",                       642690, 9438409, 0.630032),
    ("52",       "C1086-Q9575",      11.059, "Morropon",      "Yamango",                       641522, 9438064, 0.66855),
    ("46",       "C1086-Q9575",      18.866, "Morropon",      "Yamango",                       638751, 9433236, 0.724382),
    ("48",       "C1086-Q9575",       5.754, "Morropon",      "Yamango",                       639496, 9433679, 0.697809),
    ("3",        "C1096-Q9545",     459.295, "Ayabaca",       "Frias",                         606816, 9458469, 0.567707),
    ("6",        "C1096-Q9547",     168.411, "Ayabaca",       "Frias",                         605636, 9455269, 0.538144),
    ("36",       "C1096-Q9547",      24.843, "Ayabaca",       "Frias",                         606468, 9454689, 0.539458),
    ("27",       "C1096-Q9556",      78.154, "Ayabaca",       "Frias",                         613302, 9450231, 0.53228),
    ("12",       "C1086-Q9570",     119.538, "Morropon",      "Santo Domingo",                 625419, 9444932, 0.706689),
    ("13",       "C1096-Q9564",     100.514, "Morropon",      "Santo Domingo",                 620263, 9442343, 0.671094),
    ("11",       "C1086-Q9570",     103.948, "Morropon",      "Santo Domingo",                 620759, 9439988, 0.604743),
    ("17",       "C1086-Q9570",      96.079, "Morropon",      "Santo Domingo",                 618661, 9438267, 0.647895),
    ("23",       "C1086-Q9570",      81.776, "Morropon",      "Santo Domingo",                 619939, 9438544, 0.573141),
    ("15",       "C1086-Q9570",      86.147, "Morropon",      "Santo Domingo",                 620200, 9437862, 0.560518),
    ("14",       "C1086-Q9570",      92.905, "Morropon",      "Santa Catalina de Mossa",       624397, 9437085, 0.51974),
    ("42",       "C1086-Q9570",      80.616, "Morropon",      "Santa Catalina de Mossa",       625152, 9435011, 0.634841),
    ("30",       "C1081-Q9591",      34.874, "Huancabamba",   "Canchaque",                     652758, 9424380, 0.663804),
    ("8",        "C1076-Q9593",     126.952, "Huancabamba",   "Huarmaca",                      651673, 9373383, 0.554802),
    ("M17B10",   "C1096-Q9556",      75.363, "Ayabaca",       "Frias",                         606173, 9447361, 0.472466),
    ("39",       "C1096-Q9564",      68.569, "Ayabaca",       "Frias",                         618523, 9448918, 0.701601),
    ("M19B7",    "C1086-Q9570",      34.171, "Morropon",      "Santo Domingo",                 620770, 9436562, 0.591921),
    ("M36B2",    "C1086-Q9576",       28.19, "Morropon",      "Santa Catalina de Mossa",       626438, 9435060, 0.597433),
    ("M30B5",    "C1081-Q9591",      90.903, "Huancabamba",   "Huancabamba",                   655263, 9427340, 0.5564),
    ("16",       "C1076-Q9593",      83.327, "Huancabamba",   "Huarmaca",                      652939, 9382105, 0.628423),
    ("M18B1",    "C1077-Q9579",      160.36, "Morropon",      "Buenos Aires",                  621058, 9410427, 0.334299),
    ("M1B1",     "C1077-Q9580",     188.797, "Morropon",      "Buenos Aires",                  617695, 9409136, 0.319305),
    ("M6B2-3",   "C1077-Q9566",     160.978, "Morropon",      "Buenos Aires",                  611305, 9419274, 0.406259),
    ("M32B3",    "C1086-Q9569",      50.972, "Morropon",      "Morropon",                      618248, 9426493, 0.385213),
    ("M19B2",    "C1086-Q9570",      74.372, "Morropon",      "Morropon",                      617715, 9427284, 0.401185),
    ("22",       "C1096-Q9556",       79.33, "Morropon",      "Chulucanas",                    601122, 9443651, 0.42317),
    ("M2B8",     "C1076-Q9584",     116.099, "Huancabamba",   "San Miguel de El Faique",       639059, 9398295, 0.375171),
    ("M3B1",     "C1081-Q9582",      80.983, "Morropon",      "Salitral",                      630685, 9410800, 0.323882),
    ("M20B1",    "C1076-Q9585",      279.68, "Huancabamba",   "San Miguel de El Faique",       641196, 9398710, 0.355329),
    ("M7B1",     "C1076-Q9581",     130.425, "Morropon",      "Salitral",                      629006, 9407536, 0.313647),
    ("M7B6",     "C1076-Q9581",      91.285, "Morropon",      "Salitral",                      635110, 9399092, 0.287833),
    ("M7B2",     "C1076-Q9581",     115.379, "Morropon",      "Salitral",                      631438, 9403547, 0.287689),
    ("M2B1",     "C1076-Q9584",      81.795, "Morropon",      "Salitral",                      635876, 9397941, 0.29836),
    ("M7B3",     "C1076-Q9581",      65.745, "Morropon",      "Salitral",                      633232, 9401227, 0.298017),
    ("M27B1",    "C1096-Q9557",     448.804, "Morropon",      "Chulucanas",                    593246, 9430139, 0.363228),
    ("M9B1",     "C1096-Q9545",     355.364, "Morropon",      "Chulucanas",                    595552, 9451222, 0.314312),
    ("78",       "C1086-Q9570",        35.1, "Morropon",      "Santo Domingo",                 619630, 9440043, 0.707402),
    ("81",       "C1076-Q9585",       8.929, "Huancabamba",   "Canchaque",                     641538, 9401360, 0.274714),
]

def migrar(db_url):
    print(f"\nConectando a la base de datos...")
    conn = psycopg2.connect(db_url)
    cursor = conn.cursor()

    # 1. Eliminar todos los bloques (CASCADE borra registros vinculados)
    print("Eliminando bloques existentes (y registros vinculados)...")
    cursor.execute("DELETE FROM bloques")
    eliminados = cursor.rowcount
    print(f"  {eliminados} bloques eliminados.")

    # 2. Reiniciar secuencia de IDs
    cursor.execute("ALTER SEQUENCE IF EXISTS bloques_id_seq RESTART WITH 1")

    # 3. Insertar los 124 bloques V5
    print("Insertando 124 bloques V5...")
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    insertados = 0
    for (codigo, microcuenca, area_ha, provincia, distrito, utm_este, utm_norte, msavi) in BLOQUES_V5:
        cursor.execute("""
            INSERT INTO bloques (codigo, tipo_intervencion, cuenca, distrito,
                                 utm_este, utm_norte, utm_zona, altitud,
                                 area_hectareas, responsable, estado,
                                 microcuenca, provincia, fecha_registro)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            codigo,
            "Restauracion",   # tipo_intervencion por defecto
            microcuenca,      # cuenca
            distrito,
            float(utm_este),  # utm_este (Zona 17S WGS84)
            float(utm_norte), # utm_norte (Zona 17S WGS84)
            "17S",            # utm_zona
            0.0,              # altitud
            area_ha,
            "",               # responsable
            "Pendiente",      # estado
            microcuenca,
            provincia,
            fecha,
        ))
        insertados += 1

    conn.commit()
    cursor.close()
    conn.close()
    print(f"  {insertados} bloques V5 insertados exitosamente.")
    print("\nMigracion completada. Reinicia el aplicativo para ver los cambios.")

if __name__ == "__main__":
    db_url = leer_database_url()
    migrar(db_url)
