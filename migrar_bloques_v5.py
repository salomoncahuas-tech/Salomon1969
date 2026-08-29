"""
Script de migracion: Reemplaza los bloques existentes en la BD por los
132 bloques V5 (con zonas Z01..Z14 y coordenadas UTM).
ATENCION: Elimina TODOS los registros vinculados (inspecciones, diagnosticos, etc.)

Para SOLO dar de alta bloques faltantes sin borrar nada, use en su lugar
`database.sincronizar_bloques_catalogo()` o la seccion "Sincronizar catalogo
de bloques" de la pagina "Bloques de Intervencion" del aplicativo.

Uso:
    python migrar_bloques_v5.py
"""

import sys
import os
import psycopg2
from datetime import datetime

# Lectura de DATABASE_URL desde secrets.toml
def leer_database_url():
    secrets_path = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")
    if not os.path.exists(secrets_path):
        print(f"ERROR: No se encontro {secrets_path}")
        return input("DATABASE_URL: ").strip()
    with open(secrets_path, "r") as fh:
        for line in fh:
            if line.strip().startswith("DATABASE_URL"):
                return line.split("=", 1)[1].strip().strip('"').strip("'")
    print("ERROR: DATABASE_URL no encontrada en secrets.toml")
    sys.exit(1)

# Bloques V5 (132 bloques) - Plantilla_DT_Campo_Check_Validada_V5.xlsx
# + ampliacion "Bloques Adicionalea.xlsx" (bloques 83..87)
# (codigo, microcuenca, area_ha, provincia, distrito, zona, utm_este, utm_norte, msavi_2024)
BLOQUES_V5 = [
    ("M9B1", "C1096-Q9545", 355.36, "Morropon", "Chulucanas", "Z01", 595552, 9451222, 0.314312),
    ("M10B4", "C1096-Q9547", 68.89, "Morropon", "Chulucanas", "Z01", 596773, 9448273, 0.313301),
    ("25", "C1096-Q9547", 53.48, "Morropon", "Chulucanas", "Z01", 596578, 9442453, 0.256543),
    ("32", "C1096-Q9556", 35.83, "Morropon", "Chulucanas", "Z01", 599927, 9442604, 0.371619),
    ("22", "C1096-Q9556", 79.33, "Morropon", "Chulucanas", "Z01", 601122, 9443651, 0.42317),
    ("M17B1", "C1096-Q9556", 42.45, "Morropon", "Chulucanas", "Z01", 600644, 9440436, 0.332263),
    ("2", "C1096-Q9558", 330.6, "Morropon", "Chulucanas", "Z01", 603653, 9439351, 0.358601),
    ("M17B4", "C1096-Q9556", 124.71, "Morropon", "Chulucanas", "Z01", 604926, 9444327, 0.423543),
    ("M27B1", "C1096-Q9557", 448.8, "Morropon", "Chulucanas", "Z01", 593246, 9430139, 0.363228),
    ("3", "C1096-Q9545", 459.29, "Ayabaca", "Frias", "Z02", 606816, 9458469, 0.567707),
    ("6", "C1096-Q9547", 168.41, "Ayabaca", "Frias", "Z02", 605636, 9455269, 0.538144),
    ("36", "C1096-Q9547", 24.84, "Ayabaca", "Frias", "Z02", 606468, 9454689, 0.539458),
    ("M17B7", "C1096-Q9556", 258.84, "Ayabaca", "Frias", "Z02", 608472, 9448056, 0.508715),
    ("M17B6", "C1096-Q9556", 106.54, "Ayabaca", "Frias", "Z02", 607036, 9446211, 0.446441),
    ("M17B10", "C1096-Q9556", 75.36, "Ayabaca", "Frias", "Z02", 606173, 9447361, 0.472466),
    ("M17B5", "C1096-Q9556", 74.03, "Ayabaca", "Frias", "Z02", 608800, 9445585, 0.54817),
    ("27", "C1096-Q9556", 78.15, "Ayabaca", "Frias", "Z02", 613302, 9450231, 0.53228),
    ("39", "C1096-Q9564", 68.57, "Ayabaca", "Frias", "Z02", 618523, 9448918, 0.701601),
    ("56", "C1096-Q9564", 77.67, "Ayabaca", "Frias", "Z02", 631723, 9439702, 0.677658),
    ("12", "C1086-Q9570", 119.54, "Morropon", "Santo Domingo", "Z03", 625419, 9444932, 0.706689),
    ("82", "C1086-Q9570", 35.66, "Morropon", "Santo Domingo", "Z03", 621452, 9443855, 0.62835),
    ("13", "C1096-Q9564", 100.51, "Morropon", "Santo Domingo", "Z03", 620263, 9442343, 0.671094),
    ("78", "C1086-Q9570", 35.1, "Morropon", "Santo Domingo", "Z03", 619630, 9440043, 0.707402),
    ("11", "C1086-Q9570", 103.95, "Morropon", "Santo Domingo", "Z03", 620759, 9439988, 0.604743),
    ("23", "C1086-Q9570", 81.78, "Morropon", "Santo Domingo", "Z03", 619939, 9438544, 0.573141),
    ("15", "C1086-Q9570", 86.15, "Morropon", "Santo Domingo", "Z03", 620200, 9437862, 0.560518),
    ("M19B7", "C1086-Q9570", 34.17, "Morropon", "Santo Domingo", "Z03", 620770, 9436562, 0.591921),
    ("17", "C1086-Q9570", 96.08, "Morropon", "Santo Domingo", "Z03", 618661, 9438267, 0.647895),
    ("M8B2", "C1096-Q9558", 35.35, "Morropon", "Santo Domingo", "Z03", 617596, 9436227, 0.530453),
    ("57", "C1086-Q9576", 25.67, "Morropon", "Chalaco", "Z04", 631195, 9441777, 0.603055),
    ("59", "C1086-Q9570", 13.55, "Morropon", "Chalaco", "Z04", 631195, 9441777, 0.614739),
    ("M36B2", "C1086-Q9576", 28.19, "Morropon", "Santa Catalina de Mossa", "Z04", 626438, 9435060, 0.597433),
    ("42", "C1086-Q9570", 80.62, "Morropon", "Santa Catalina de Mossa", "Z04", 625152, 9435011, 0.634841),
    ("14", "C1086-Q9570", 92.9, "Morropon", "Santa Catalina de Mossa", "Z04", 624397, 9437085, 0.51974),
    ("58", "C1086-Q9569", 11.47, "Morropon", "Santa Catalina de Mossa", "Z04", 619094, 9426406, 0.420325),
    ("M19B2", "C1086-Q9570", 74.37, "Morropon", "Morropon", "Z05", 617715, 9427284, 0.401185),
    ("M32B3", "C1086-Q9569", 50.97, "Morropon", "Morropon", "Z05", 618248, 9426493, 0.385213),
    ("M6B10", "C1077-Q9566", 249.89, "Morropon", "Buenos Aires", "Z06", 618299, 9423119, 0.30989),
    ("M6B2-3", "C1077-Q9566", 160.98, "Morropon", "Buenos Aires", "Z06", 611305, 9419274, 0.406259),
    ("M6B2-2", "C1077-Q9566", 333.01, "Morropon", "Buenos Aires", "Z06", 611196, 9416725, 0.302636),
    ("M6B2-1", "C1077-Q9566", 514.36, "Morropon", "Buenos Aires", "Z06", 612864, 9413345, 0.3381),
    ("5", "C1077-Q9566", 295.73, "Morropon", "Buenos Aires", "Z06", 617860, 9413182, 0.3394),
    ("M1B1", "C1077-Q9580", 188.8, "Morropon", "Buenos Aires", "Z06", 617695, 9409136, 0.319305),
    ("M18B1", "C1077-Q9579", 160.36, "Morropon", "Buenos Aires", "Z06", 621058, 9410427, 0.334299),
    ("M3B9", "C1081-Q9582", 294.2, "Morropon", "Salitral", "Z07", 629338, 9413305, 0.321386),
    ("M3B1", "C1081-Q9582", 80.98, "Morropon", "Salitral", "Z07", 630685, 9410800, 0.323882),
    ("M7B1", "C1076-Q9581", 130.43, "Morropon", "Salitral", "Z07", 629006, 9407536, 0.313647),
    ("M7B2", "C1076-Q9581", 115.38, "Morropon", "Salitral", "Z07", 631438, 9403547, 0.287689),
    ("M7B3", "C1076-Q9581", 65.74, "Morropon", "Salitral", "Z07", 633232, 9401227, 0.298017),
    ("M7B6", "C1076-Q9581", 91.29, "Morropon", "Salitral", "Z07", 635110, 9399092, 0.287833),
    ("M2B1", "C1076-Q9584", 81.79, "Morropon", "Salitral", "Z07", 635876, 9397941, 0.29836),
    ("7", "C1076-Q9584", 173.43, "Morropon", "Salitral", "Z07", 639728, 9396836, 0.375288),
    ("M2B5", "C1076-Q9584", 30.91, "Morropon", "Salitral", "Z07", 639987, 9394392, 0.304733),
    ("1", "C1076-Q9584", 1370.96, "Morropon", "Salitral", "Z07", 645575, 9393500, 0.454493),
    ("M18B3", "C1077-Q9579", 373.81, "Morropon", "Salitral", "Z07", 625311, 9412195, 0.369424),
    ("M18B5", "C1077-Q9579", 197.35, "Morropon", "Salitral", "Z07", 621551, 9412989, 0.356522),
    ("M3B7", "C1081-Q9582", 122.95, "Morropon", "San Juan de Bigote", "Z08", 638309, 9415217, 0.370678),
    ("M3B6", "C1081-Q9582", 60.35, "Morropon", "San Juan de Bigote", "Z08", 637918, 9412134, 0.342535),
    ("77", "C1081-Q9583", 17.82, "Morropon", "San Juan de Bigote", "Z08", 638648, 9411370, 0.272488),
    ("M3B5", "C1081-Q9582", 52.74, "Morropon", "San Juan de Bigote", "Z08", 635887, 9411738, 0.283016),
    ("M3B3", "C1081-Q9582", 84.82, "Morropon", "San Juan de Bigote", "Z08", 633015, 9411483, 0.298483),
    ("M3B8", "C1081-Q9582", 565.58, "Morropon", "San Juan de Bigote", "Z08", 632119, 9415141, 0.293351),
    ("M11B3", "C1081-Q9583", 106.03, "Morropon", "San Juan de Bigote", "Z08", 642079, 9413823, 0.395609),
    ("55", "C1086-Q9575", 0.86, "Morropon", "Yamango", "Z09", 642690, 9438409, 0.630032),
    ("49", "C1086-Q9575", 15.84, "Morropon", "Yamango", "Z09", 642334, 9438263, 0.686036),
    ("52", "C1086-Q9575", 11.06, "Morropon", "Yamango", "Z09", 641522, 9438064, 0.66855),
    ("48", "C1086-Q9575", 5.75, "Morropon", "Yamango", "Z09", 639496, 9433679, 0.697809),
    ("46", "C1086-Q9575", 18.87, "Morropon", "Yamango", "Z09", 638751, 9433236, 0.724382),
    ("M28B2", "C1086-Q9575", 90.2, "Morropon", "Yamango", "Z09", 636249, 9425665, 0.475342),
    ("M28B3", "C1086-Q9575", 58.29, "Morropon", "Yamango", "Z09", 634061, 9426778, 0.507529),
    ("M28B4", "C1086-Q9575", 283.43, "Morropon", "Yamango", "Z09", 626078, 9427110, 0.42751),
    ("33", "C1081-Q9590", 29.68, "Huancabamba", "Lalaquiz", "Z10", 646466, 9431126, 0.71923),
    ("38", "C1081-Q9590", 44.2, "Huancabamba", "Lalaquiz", "Z10", 647123, 9427602, 0.672487),
    ("50", "C1081-Q9590", 19.27, "Huancabamba", "Lalaquiz", "Z10", 647350, 9427252, 0.68307),
    ("34", "C1081-Q9590", 28.45, "Huancabamba", "Lalaquiz", "Z10", 646483, 9426305, 0.648975),
    ("61", "C1081-Q9590", 26.83, "Huancabamba", "Lalaquiz", "Z10", 646236, 9425052, 0.611186),
    ("37", "C1081-Q9591", 35.49, "Huancabamba", "Lalaquiz", "Z10", 649081, 9424506, 0.658846),
    ("26", "C1081-Q9591", 46.96, "Huancabamba", "Lalaquiz", "Z10", 649472, 9425705, 0.673303),
    ("64", "C1081-Q9591", 35.77, "Huancabamba", "Huancabamba", "Z11", 657819, 9428541, 0.525448),
    ("54", "C1081-Q9591 (inferida)", 40.26, "Huancabamba", "Huancabamba", "Z11", 656897, 9426033, 0.600447),
    ("M30B5", "C1081-Q9591", 90.9, "Huancabamba", "Huancabamba", "Z11", 655263, 9427340, 0.5564),
    ("60", "C1081-Q9591", 40.13, "Huancabamba", "Canchaque", "Z11", 653850, 9426236, 0.67447),
    ("40", "C1081-Q9591", 28.52, "Huancabamba", "Canchaque", "Z11", 654137, 9424425, 0.640099),
    ("30", "C1081-Q9591", 34.87, "Huancabamba", "Canchaque", "Z11", 652758, 9424380, 0.663804),
    ("28", "C1081-Q9591", 52.97, "Huancabamba", "Canchaque", "Z11", 654188, 9422849, 0.639425),
    ("62", "C1081-Q9591", 73.82, "Huancabamba", "Canchaque", "Z11", 655046, 9420450, 0.681648),
    ("65", "C1081-Q9591 (inferida)", 53.93, "Huancabamba", "Canchaque", "Z11", 657541, 9421083, 0.649297),
    ("21", "C1081-Q9591", 84.23, "Huancabamba", "Canchaque", "Z11", 654604, 9417736, 0.620078),
    ("66", "C1081-Q9591", 102.34, "Huancabamba", "Canchaque", "Z11", 652566, 9417311, 0.640254),
    ("24", "C1081-Q9583", 90.16, "Huancabamba", "Canchaque", "Z11", 652191, 9414358, 0.674839),
    ("63", "C1081-Q9583", 50.46, "Huancabamba", "Canchaque", "Z11", 655059, 9413845, 0.603582),
    ("79", "C1076-Q9587", 94.81, "Huancabamba", "Canchaque", "Z11", 650305, 9402785, 0.547879),
    ("81", "C1076-Q9585", 8.93, "Huancabamba", "Canchaque", "Z11", 641538, 9401360, 0.274714),
    ("M22B1", "C1076-Q9586", 55.17, "Huancabamba", "San Miguel de El Faique", "Z12", 664082, 9404172, 0.594025),
    ("44", "C1076-Q9586", 28.23, "Huancabamba", "San Miguel de El Faique", "Z12", 663637, 9400618, 0.511324),
    ("47", "C1076-Q9586", 13.55, "Huancabamba", "San Miguel de El Faique", "Z12", 659649, 9401436, 0.593628),
    ("75", "C1076-Q9587", 15.18, "Huancabamba", "San Miguel de El Faique", "Z12", 657722, 9401669, 0.630888),
    ("9", "C1076-Q9586", 109.86, "Huancabamba", "San Miguel de El Faique", "Z12", 656578, 9397762, 0.574527),
    ("43", "C1076-Q9586", 23.17, "Huancabamba", "San Miguel de El Faique", "Z12", 656489, 9395403, 0.551525),
    ("35", "C1076-Q9586", 30.8, "Huancabamba", "San Miguel de El Faique", "Z12", 654607, 9395038, 0.624477),
    ("67", "C1076-Q9586", 13.18, "Huancabamba", "San Miguel de El Faique", "Z12", 654207, 9394210, 0.605293),
    ("M20B1", "C1076-Q9585", 279.68, "Huancabamba", "San Miguel de El Faique", "Z12", 641196, 9398710, 0.355329),
    ("M2B8", "C1076-Q9584", 116.1, "Huancabamba", "San Miguel de El Faique", "Z12", 639059, 9398295, 0.375171),
    ("29", "C1076-Q9592", 40.1, "Huancabamba", "Huarmaca", "Z13", 662123, 9397575, 0.548192),
    ("68", "C1076-Q9592", 21.29, "Huancabamba", "Huarmaca", "Z13", 664225, 9396339, 0.482981),
    ("53", "C1076-Q9592", 13.06, "Huancabamba", "Huarmaca", "Z13", 666537, 9395918, 0.596277),
    ("70", "C1076-Q9592", 22.39, "Huancabamba", "Huarmaca", "Z13", 666781, 9395573, 0.622295),
    ("31", "C1076-Q9592", 47.58, "Huancabamba", "Huarmaca", "Z13", 663334, 9392820, 0.435805),
    ("69", "C1076-Q9592", 19.43, "Huancabamba", "Huarmaca", "Z13", 663291, 9392332, 0.448926),
    ("72", "C1076-Q9592", 12.4, "Huancabamba", "Huarmaca", "Z13", 662236, 9391211, 0.518477),
    ("41", "C1076-Q9592", 64.92, "Huancabamba", "Huarmaca", "Z13", 660525, 9394327, 0.592669),
    ("73", "C1076-Q9592", 54.29, "Huancabamba", "Huarmaca", "Z13", 658547, 9395594, 0.554974),
    ("18", "C1076-Q9592", 103.62, "Huancabamba", "Huarmaca", "Z13", 656445, 9393180, 0.640219),
    ("74", "C1076-Q9588", 29.23, "Huancabamba", "Huarmaca", "Z13", 652449, 9389179, 0.356463),
    ("M12B1", "C1076-Q9588", 268.39, "Huancabamba", "Huarmaca", "Z13", 644546, 9389106, 0.34753),
    ("19", "C1076-Q9592", 82.09, "Huancabamba", "Huarmaca", "Z13", 665942, 9391144, 0.537184),
    ("M4B4", "C1076-Q9589", 289.91, "Huancabamba", "Huarmaca", "Z14", 642775, 9388574, 0.319788),
    ("M4B3", "C1076-Q9589", 234.07, "Huancabamba", "Huarmaca", "Z14", 643474, 9385620, 0.322153),
    ("10", "C1076-Q9593", 150.21, "Huancabamba", "Huarmaca", "Z14", 650830, 9382303, 0.546492),
    ("16", "C1076-Q9593", 83.33, "Huancabamba", "Huarmaca", "Z14", 652939, 9382105, 0.628423),
    ("4", "C1076-Q9588", 230.1, "Huancabamba", "Huarmaca", "Z14", 656549, 9382009, 0.660111),
    ("51", "C1076-Q9593", 22.93, "Huancabamba", "Huarmaca", "Z14", 656291, 9379096, 0.523788),
    ("76", "C1076-Q9593", 19.62, "Huancabamba", "Huarmaca", "Z14", 657922, 9377477, 0.492462),
    ("80", "C1076-Q9593", 28.83, "Huancabamba", "Huarmaca", "Z14", 657160, 9373050, 0.593069),
    ("71", "C1076-Q9593", 10.67, "Huancabamba", "Huarmaca", "Z14", 652104, 9373045, 0.557066),
    ("8", "C1076-Q9593", 126.95, "Huancabamba", "Huarmaca", "Z14", 651673, 9373383, 0.554802),
    ("20", "C1076-Q9593", 81.33, "Huancabamba", "Huarmaca", "Z14", 649847, 9376260, 0.48513),
    # Ampliacion: bloques adicionales (centroides de "CENTROIDES BLOQUES ADICIONALEES.xlsx").
    ("83", "C1081-Q9583", 24.344, "Morropon", "San Juan de Bigote", "Z08", 639720, 9410622, 0.30545594),
    ("84", "C1081-Q9583", 5.097, "Morropon", "San Juan de Bigote", "Z08", 640237, 9410714, 0.251595558),
    ("85", "C1081-Q9583", 11.906, "Morropon", "San Juan de Bigote", "Z08", 642200, 9411072, 0.228928157),
    ("86", "C1081-Q9583", 30.001, "Morropon", "San Juan de Bigote", "Z08", 642415, 9410216, 0.273471802),
    ("87", "C1081-Q9583", 62.219, "Morropon", "San Juan de Bigote", "Z08", 642882, 9411123, 0.248717766),
]

def migrar(db_url):
    print("\nConectando a la base de datos...")
    conn = psycopg2.connect(db_url)
    cursor = conn.cursor()

    print("Eliminando bloques existentes (y registros vinculados)...")
    cursor.execute("DELETE FROM bloques")
    eliminados = cursor.rowcount
    print(f"  {eliminados} bloques eliminados.")

    cursor.execute("ALTER SEQUENCE IF EXISTS bloques_id_seq RESTART WITH 1")

    print(f"Insertando {len(BLOQUES_V5)} bloques V5...")
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    insertados = 0
    for (codigo, microcuenca, area_ha, provincia, distrito,
         zona, utm_este, utm_norte, msavi) in BLOQUES_V5:
        cursor.execute("""
            INSERT INTO bloques (codigo, tipo_intervencion, cuenca, distrito,
                                 utm_este, utm_norte, utm_zona, altitud,
                                 area_hectareas, responsable, estado,
                                 microcuenca, provincia, fecha_registro)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            codigo,
            "Restauracion",
            microcuenca,
            distrito,
            float(utm_este),
            float(utm_norte),
            "17S",
            0.0,
            area_ha,
            "",
            "Pendiente",
            microcuenca,
            provincia,
            fecha,
        ))
        insertados += 1

    conn.commit()
    cursor.close()
    conn.close()
    print(f"  {insertados} bloques insertados exitosamente.")
    print("\nMigracion V5 completada. Reinicia el aplicativo para ver los cambios.")

if __name__ == "__main__":
    db_url = leer_database_url()
    migrar(db_url)
