"""
Crea los ZIP de entrega y el indice de informacion remitida.

Uso: python empaquetar.py --paquetes <carpeta de paquetes> --salida <carpeta de entrega>
El paquete 06 se divide en 06A (capas) y 06B (mapas) para que cada ZIP sea manejable por correo.
"""
import argparse
import shutil
import zipfile
from pathlib import Path

from openpyxl import Workbook

import anin_excel_fmt as X

DESC = {
    "01_Delimitacion_Ambito_Estudio": "Delimitación del ámbito de estudio, área de intervención y área de influencia",
    "02_Criterios_Seleccion": "Criterios de selección (territoriales y técnicos) de los ámbitos de intervención",
    "03_Diagnostico_Poblacion_Afectada": "Avance del diagnóstico de la población afectada (E2 B4)",
    "04_Diagnostico_Unidades_Productoras": "Avance del diagnóstico de la Unidad Productora: ecosistemas (E5, E6, E8)",
    "05_Diagnostico_Involucrados": "Avance del diagnóstico de involucrados (actores, matriz T1.7, riesgos sociales)",
    "06A_Informacion_Georreferenciada_Capas": "Capas SHP, KMZ y GPKG, procesos e inventario",
    "06B_Informacion_Georreferenciada_Mapas": "Mapas temáticos y distritales (PDF, PNG, JPG)",
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--paquetes", required=True)
    ap.add_argument("--salida", required=True)
    a = ap.parse_args()
    P, S = Path(a.paquetes), Path(a.salida)
    if S.exists():
        shutil.rmtree(S)
    S.mkdir(parents=True)
    filas_idx = []
    planes = {}
    for k in ("01_Delimitacion_Ambito_Estudio", "02_Criterios_Seleccion", "03_Diagnostico_Poblacion_Afectada",
              "04_Diagnostico_Unidades_Productoras", "05_Diagnostico_Involucrados"):
        planes[k] = [(P / k, ".")]
    g = P / "06_Informacion_Georreferenciada_Cartografica"
    planes["06A_Informacion_Georreferenciada_Capas"] = [
        (g / "06_Nota_Tecnica_Informacion_Espacial.docx", "06_Nota_Tecnica_Informacion_Espacial.docx"),
        (g / "06_Inventario_Capas_IN_Piura.xlsx", "06_Inventario_Capas_IN_Piura.xlsx"),
        (g / "01_SHP", "01_SHP"), (g / "02_KMZ", "02_KMZ"), (g / "03_GPKG", "03_GPKG"), (g / "05_Procesos", "05_Procesos")]
    planes["06B_Informacion_Georreferenciada_Mapas"] = [
        (g / "06_Nota_Tecnica_Informacion_Espacial.docx", "06_Nota_Tecnica_Informacion_Espacial.docx"),
        (g / "04_Mapas", "04_Mapas")]
    for nombre, items in planes.items():
        items = [(s, arc.replace("./", "")) for s, arc in items]
        z = S / f"{nombre}.zip"
        # carpeta raiz del zip = nombre del paquete
        with zipfile.ZipFile(z, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
            n = 0
            for src, arc in items:
                src = Path(src)
                if src.is_dir():
                    for f in sorted(src.rglob("*")):
                        if f.is_file():
                            rel = f.relative_to(src)
                            dest = f"{nombre}/{rel}" if arc == "." else f"{nombre}/{arc}/{rel}"
                            zf.write(f, dest); n += 1
                            filas_idx.append([nombre, dest.split("/", 1)[1], round(f.stat().st_size / 1048576, 3)])
                else:
                    zf.write(src, f"{nombre}/{arc}"); n += 1
                    filas_idx.append([nombre, arc, round(src.stat().st_size / 1048576, 3)])
        print(f"{z.name}: {n} archivos, {z.stat().st_size / 1048576:.1f} MB")
    shutil.copy2(P / "00_Informe_Remision_Avance_Perfil_IN_Piura.docx", S)

    wb = Workbook()
    ws = wb.active; ws.title = "Resumen_Paquetes"
    f = X.encabezado(ws, "ÍNDICE DE INFORMACIÓN REMITIDA — PERFIL «INFRAESTRUCTURA NATURAL» (CUI 2669244)",
                     "Respuesta a la solicitud del Programa de Inversión del Río Piura. Corte de información: 12/11/2026.", 5)
    resumen = []
    for k, nombre in enumerate(planes):
        fila = f + 1 + k
        resumen.append([nombre + ".zip", DESC[nombre],
                        f'=COUNTIF(Detalle_Archivos!$A$9:$A${8 + len(filas_idx)},"{nombre}")',
                        f'=SUMIF(Detalle_Archivos!$A$9:$A${8 + len(filas_idx)},"{nombre}",Detalle_Archivos!$C$9:$C${8 + len(filas_idx)})',
                        round((S / f"{nombre}.zip").stat().st_size / 1048576, 2)])
    X.tabla(ws, f, ["Archivo ZIP", "Contenido", "N.° de archivos", "Tamaño sin comprimir (MB)", "Tamaño ZIP (MB)"],
            resumen, anchos=[46, 80, 14, 22, 16], total_cols=[3, 4, 5], fmt={4: "#,##0.00", 5: "#,##0.00"})
    ws2 = wb.create_sheet("Detalle_Archivos")
    f2 = X.encabezado(ws2, "DETALLE DE ARCHIVOS POR PAQUETE", "Rutas relativas a la carpeta raíz de cada ZIP.", 3)
    X.tabla(ws2, f2, ["Paquete", "Archivo", "Tamaño (MB)"], filas_idx, anchos=[42, 110, 14], total_cols=[3],
            fmt={3: "#,##0.000"})
    wb.save(S / "00_Indice_Informacion_Remitida_IN_Piura.xlsx")
    print("índice:", len(filas_idx), "archivos")


if __name__ == "__main__":
    main()
