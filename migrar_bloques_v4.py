"""
Script de migracion: Reemplaza los bloques existentes en la BD por los 128 bloques V4.
ATENCION: Elimina TODOS los registros vinculados (inspecciones, diagnosticos, etc.)

Uso:
    python migrar_bloques_v4.py
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

# ── 128 Bloques V4 ──────────────────────────────────────────────────────────
# (codigo, microcuenca, area_ha, provincia, distrito, msavi_2024)
BLOQUES_128 = [
    ("1",      "C1076-Q9584", 1371.335, "Morropon",     "Salitral",                   0.454493),
    ("2",      "C1096-Q9558",  330.776, "Morropon",     "Chulucanas",                 0.358601),
    ("3",      "C1096-Q9545",  459.532, "Ayabaca",      "Frias",                      0.567707),
    ("4",      "C1076-Q9588",  230.143, "Huancabamba",  "Huarmaca",                   0.660111),
    ("5",      "C1077-Q9566",  295.861, "Morropon",     "Buenos Aires",               0.3394),
    ("6",      "C1096-Q9547",  168.499, "Ayabaca",      "Frias",                      0.538144),
    ("7",      "C1076-Q9584",  173.482, "Morropon",     "Salitral",                   0.375288),
    ("8",      "C1076-Q9593",  126.981, "Huancabamba",  "Huarmaca",                   0.554802),
    ("9",      "C1076-Q9586",  109.882, "Huancabamba",  "San Miguel de El Faique",    0.574527),
    ("10",     "C1076-Q9593",  150.25,  "Huancabamba",  "Huarmaca",                   0.546492),
    ("11",     "C1086-Q9570",  103.993, "Morropon",     "Santo Domingo",              0.604743),
    ("12",     "C1086-Q9570",  119.587, "Morropon",     "Santo Domingo",              0.706689),
    ("13",     "C1096-Q9564",  100.558, "Morropon",     "Santo Domingo",              0.671094),
    ("14",     "C1086-Q9570",   92.944, "Morropon",     "Santa Catalina de Mossa",    0.51974),
    ("15",     "C1086-Q9570",   86.185, "Morropon",     "Santo Domingo",              0.560518),
    ("16",     "C1076-Q9593",   83.346, "Huancabamba",  "Huarmaca",                   0.628423),
    ("17",     "C1086-Q9570",   96.123, "Morropon",     "Santo Domingo",              0.647895),
    ("18",     "C1076-Q9592",  103.637, "Huancabamba",  "Huarmaca",                   0.640219),
    ("19",     "C1076-Q9592",   82.098, "Huancabamba",  "Huarmaca",                   0.537184),
    ("20",     "C1076-Q9593",   81.354, "Huancabamba",  "Huarmaca",                   0.48513),
    ("21",     "C1081-Q9591",   84.244, "Huancabamba",  "Canchaque",                  0.620078),
    ("22",     "C1096-Q9556",   79.374, "Morropon",     "Chulucanas",                 0.42317),
    ("23",     "C1086-Q9570",   81.812, "Morropon",     "Santo Domingo",              0.573141),
    ("24",     "C1081-Q9583",   90.18,  "Huancabamba",  "Canchaque",                  0.674839),
    ("25",     "C1096-Q9547",   53.515, "Morropon",     "Chulucanas",                 0.256543),
    ("26",     "C1081-Q9591",   46.973, "Huancabamba",  "Lalaquiz",                   0.673303),
    ("27",     "C1096-Q9556",   78.192, "Ayabaca",      "Frias",                      0.53228),
    ("28",     "C1081-Q9591",   52.977, "Huancabamba",  "Canchaque",                  0.639425),
    ("29",     "C1076-Q9592",   40.11,  "Huancabamba",  "Huarmaca",                   0.548192),
    ("30",     "C1081-Q9591",   34.882, "Huancabamba",  "Canchaque",                  0.663804),
    ("31",     "C1076-Q9592",   47.588, "Huancabamba",  "Huarmaca",                   0.435805),
    ("32",     "C1096-Q9556",   35.85,  "Morropon",     "Chulucanas",                 0.371619),
    ("33",     "C1081-Q9590",   29.689, "Huancabamba",  "Lalaquiz",                   0.71923),
    ("34",     "C1081-Q9590",   28.458, "Huancabamba",  "Lalaquiz",                   0.648975),
    ("35",     "C1076-Q9586",   30.811, "Huancabamba",  "San Miguel de El Faique",    0.624477),
    ("36",     "C1096-Q9547",   24.856, "Ayabaca",      "Frias",                      0.539458),
    ("37",     "C1081-Q9591",   35.5,   "Huancabamba",  "Lalaquiz",                   0.658846),
    ("38",     "C1081-Q9590",   44.212, "Huancabamba",  "Lalaquiz",                   0.672487),
    ("39",     "C1096-Q9564",   68.6,   "Ayabaca",      "Frias",                      0.701601),
    ("40",     "C1081-Q9591",   28.529, "Huancabamba",  "Canchaque",                  0.640099),
    ("41",     "C1076-Q9592",   64.933, "Huancabamba",  "Huarmaca",                   0.592669),
    ("42",     "C1086-Q9570",   80.649, "Morropon",     "Santa Catalina de Mossa",    0.634841),
    ("43",     "C1076-Q9586",   23.176, "Huancabamba",  "San Miguel de El Faique",    0.551525),
    ("44",     "C1076-Q9586",   28.238, "Huancabamba",  "San Miguel de El Faique",    0.511324),
    ("45",     "C1076-Q9592",   17.626, "Huancabamba",  "Huarmaca",                   0.445813),
    ("46",     "C1086-Q9575",   18.872, "Morropon",     "Yamango",                    0.724382),
    ("47",     "C1076-Q9586",   13.554, "Huancabamba",  "San Miguel de El Faique",    0.593628),
    ("48",     "C1086-Q9575",    5.755, "Morropon",     "Yamango",                    0.697809),
    ("49",     "C1086-Q9575",   15.844, "Morropon",     "Yamango",                    0.686036),
    ("50",     "C1081-Q9590",   19.272, "Huancabamba",  "Lalaquiz",                   0.68307),
    ("51",     "C1076-Q9593",   22.932, "Huancabamba",  "Huarmaca",                   0.523788),
    ("52",     "C1086-Q9575",   11.063, "Morropon",     "Yamango",                    0.66855),
    ("53",     "C1076-Q9592",   13.066, "Huancabamba",  "Huarmaca",                   0.596277),
    ("54",     "C1081-Q9591",   40.27,  "Huancabamba",  "Huancabamba",                0.600447),
    ("55",     "C1086-Q9575",    0.859, "Morropon",     "Yamango",                    0.630032),
    ("56",     "C1096-Q9564",   77.699, "Ayabaca",      "Frias",                      0.677658),
    ("57",     "C1086-Q9576",   25.677, "Morropon",     "Chalaco",                    0.603055),
    ("58",     "C1086-Q9569",   11.48,  "Morropon",     "Santa Catalina de Mossa",    0.420325),
    ("59",     "C1086-Q9570",   13.559, "Morropon",     "Chalaco",                    0.614739),
    ("60",     "C1081-Q9591",   40.14,  "Huancabamba",  "Canchaque",                  0.67447),
    ("61",     "C1081-Q9590",   26.833, "Huancabamba",  "Lalaquiz",                   0.611186),
    ("62",     "C1081-Q9591",   73.837, "Huancabamba",  "Canchaque",                  0.681648),
    ("63",     "C1081-Q9583",   50.469, "Huancabamba",  "Canchaque",                  0.603582),
    ("64",     "C1081-Q9591",   35.78,  "Huancabamba",  "Huancabamba",                0.525448),
    ("65",     "C1081-Q9591",   53.944, "Huancabamba",  "Canchaque",                  0.649297),
    ("66",     "C1081-Q9591",  102.358, "Huancabamba",  "Canchaque",                  0.640254),
    ("67",     "C1076-Q9586",   13.187, "Huancabamba",  "San Miguel de El Faique",    0.605293),
    ("68",     "C1076-Q9592",   21.294, "Huancabamba",  "Huarmaca",                   0.482981),
    ("69",     "C1076-Q9592",   19.428, "Huancabamba",  "Huarmaca",                   0.448926),
    ("70",     "C1076-Q9592",   22.39,  "Huancabamba",  "Huarmaca",                   0.622295),
    ("71",     "C1076-Q9593",   10.676, "Huancabamba",  "Huarmaca",                   0.557066),
    ("72",     "C1076-Q9592",   12.4,   "Huancabamba",  "Huarmaca",                   0.518477),
    ("73",     "C1076-Q9592",   54.3,   "Huancabamba",  "Huarmaca",                   0.554974),
    ("74",     "C1076-Q9588",   29.233, "Huancabamba",  "Huarmaca",                   0.356463),
    ("75",     "C1076-Q9587",   15.179, "Huancabamba",  "San Miguel de El Faique",    0.630888),
    ("76",     "C1076-Q9593",   19.619, "Huancabamba",  "Huarmaca",                   0.492462),
    ("77",     "C1081-Q9583",   17.827, "Morropon",     "San Juan de Bigote",         0.272488),
    ("78",     "C1086-Q9570",   35.116, "Morropon",     "Santo Domingo",              0.707402),
    ("79",     "C1076-Q9587",   94.834, "Huancabamba",  "Canchaque",                  0.547879),
    ("80",     "C1076-Q9593",   28.835, "Huancabamba",  "Huarmaca",                   0.593069),
    ("81",     "C1076-Q9585",    8.931, "Huancabamba",  "Canchaque",                  0.274714),
    ("82",     "C1086-Q9570",   35.677, "Morropon",     "Santo Domingo",              0.62835),
    ("M1B1",   "C1077-Q9580",  188.883, "Morropon",     "Buenos Aires",               0.319305),
    ("M2B1",   "C1076-Q9584",   81.823, "Morropon",     "Salitral",                   0.29836),
    ("M2B5",   "C1076-Q9584",   30.918, "Morropon",     "Salitral",                   0.304733),
    ("M2B8",   "C1076-Q9584",  116.136, "Huancabamba",  "San Miguel de El Faique",    0.375171),
    ("M3B1",   "C1081-Q9582",   81.014, "Morropon",     "Salitral",                   0.323882),
    ("M3B3",   "C1081-Q9582",   84.855, "Morropon",     "San Juan de Bigote",         0.298483),
    ("M3B5",   "C1081-Q9582",   52.761, "Morropon",     "San Juan de Bigote",         0.283016),
    ("M3B6",   "C1081-Q9582",   60.372, "Morropon",     "San Juan de Bigote",         0.342535),
    ("M3B7",   "C1081-Q9582",  122.987, "Morropon",     "San Juan de Bigote",         0.370678),
    ("M3B8",   "C1081-Q9582",  565.792, "Morropon",     "San Juan de Bigote",         0.293351),
    ("M3B9",   "C1081-Q9582",  294.318, "Morropon",     "Salitral",                   0.321386),
    ("M4B3",   "C1076-Q9589",  234.137, "Huancabamba",  "Huarmaca",                   0.322153),
    ("M4B4",   "C1076-Q9589",  290.0,   "Huancabamba",  "Huarmaca",                   0.319788),
    ("M6B2-1", "C1077-Q9566",  514.608, "Morropon",     "Buenos Aires",               0.3381),
    ("M6B2-2", "C1077-Q9566",  333.176, "Morropon",     "Buenos Aires",               0.302636),
    ("M6B2-3", "C1077-Q9566",  161.058, "Morropon",     "Buenos Aires",               0.406259),
    ("M6B10",  "C1077-Q9566",  250.004, "Morropon",     "Buenos Aires",               0.30989),
    ("M7B1",   "C1076-Q9581",  130.475, "Morropon",     "Salitral",                   0.313647),
    ("M7B2",   "C1076-Q9581",  115.422, "Morropon",     "Salitral",                   0.287689),
    ("M7B3",   "C1076-Q9581",   65.768, "Morropon",     "Salitral",                   0.298017),
    ("M7B6",   "C1076-Q9581",   91.317, "Morropon",     "Salitral",                   0.287833),
    ("M8B2",   "C1096-Q9558",   35.363, "Morropon",     "Santo Domingo",              0.530453),
    ("M9B1",   "C1096-Q9545",  355.568, "Morropon",     "Chulucanas",                 0.314312),
    ("M10B4",  "C1096-Q9547",   68.926, "Morropon",     "Chulucanas",                 0.313301),
    ("M11B3",  "C1081-Q9583",  106.059, "Morropon",     "San Juan de Bigote",         0.395609),
    ("M12B1",  "C1076-Q9588",  268.469, "Huancabamba",  "Huarmaca",                   0.34753),
    ("M17B1",  "C1096-Q9556",   42.472, "Morropon",     "Chulucanas",                 0.332263),
    ("M17B4",  "C1096-Q9556",  124.775, "Morropon",     "Chulucanas",                 0.423543),
    ("M17B5",  "C1096-Q9556",   74.065, "Ayabaca",      "Frias",                      0.54817),
    ("M17B6",  "C1096-Q9556",  106.593, "Ayabaca",      "Frias",                      0.446441),
    ("M17B7",  "C1096-Q9556",  258.969, "Ayabaca",      "Frias",                      0.508715),
    ("M17B10", "C1096-Q9556",   75.403, "Ayabaca",      "Frias",                      0.472466),
    ("M18B1",  "C1077-Q9579",  160.43,  "Morropon",     "Buenos Aires",               0.334299),
    ("M18B3",  "C1077-Q9579",  373.96,  "Morropon",     "Salitral",                   0.369424),
    ("M18B5",  "C1077-Q9579",  197.437, "Morropon",     "Salitral",                   0.356522),
    ("M19B2",  "C1086-Q9570",   74.406, "Morropon",     "Morropon",                   0.401185),
    ("M19B7",  "C1086-Q9570",   34.186, "Morropon",     "Santo Domingo",              0.591921),
    ("M20B1",  "C1076-Q9585",  279.766, "Huancabamba",  "San Miguel de El Faique",    0.355329),
    ("M22B1",  "C1076-Q9586",   55.177, "Huancabamba",  "San Miguel de El Faique",    0.594025),
    ("M27B1",  "C1096-Q9557",  449.066, "Morropon",     "Chulucanas",                 0.363228),
    ("M28B2",  "C1086-Q9575",   90.228, "Morropon",     "Yamango",                    0.475342),
    ("M28B3",  "C1086-Q9575",   58.312, "Morropon",     "Yamango",                    0.507529),
    ("M28B4",  "C1086-Q9575",  283.543, "Morropon",     "Yamango",                    0.42751),
    ("M30B5",  "C1081-Q9591",   90.922, "Huancabamba",  "Huancabamba",                0.5564),
    ("M32B3",  "C1086-Q9569",   50.996, "Morropon",     "Morropon",                   0.385213),
    ("M36B2",  "C1086-Q9576",   28.202, "Morropon",     "Santa Catalina de Mossa",    0.597433),
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

    # 3. Insertar los 128 bloques V4
    print("Insertando 128 bloques V4...")
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    insertados = 0
    for (codigo, microcuenca, area_ha, provincia, distrito, msavi) in BLOQUES_128:
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
            0.0,              # utm_este (sin datos en V4)
            0.0,              # utm_norte
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
    print(f"  {insertados} bloques insertados exitosamente.")
    print("\nMigracion completada. Reinicia el aplicativo para ver los cambios.")

if __name__ == "__main__":
    db_url = leer_database_url()
    migrar(db_url)
