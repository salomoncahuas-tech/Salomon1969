"""Integra las 117 plantillas DT de campo con las 117 fichas de resumen V6.

Proyecto IN Piura | CUI 2669244 | ANIN - DIME - SESDI.

Genera, a partir de las dos carpetas que el repositorio ya trae:

  plantillas_dt_campo/         fichas F-DT-01 a F-DT-05 levantadas en campo
  plantillas_117_msavi_v6/     fichas de resumen de gabinete (V6)

los productos de la integracion:

  plantillas_117_integradas_v7/            las 117 fichas de resumen actualizadas
  datos/manifiesto_dt_campo_117.json       catalogo de las plantillas de campo
  datos/integracion_dt_campo_117.json      indice de la integracion por bloque
  REPORTE_INTEGRACION_DT_CAMPO_117.xlsx    consolidado con graficos
  REPORTE_INTEGRACION_DT_CAMPO_117.pdf     consolidado con graficos

Uso:
    python integrar_dt_campo.py [--sin-reportes] [--agrupacion provincia]

No requiere base de datos ni red: trabaja sobre los archivos del repositorio.
"""

import argparse
import json
import os
import sys
import time

import dt_campo as dtc
import resumenes_bloques as rbq

RAIZ = os.path.dirname(os.path.abspath(__file__))
CARPETA_SALIDA = os.path.join(RAIZ, dtc.CARPETA_INTEGRADA)
MANIFIESTO_CAMPO = os.path.join(RAIZ, "datos", "manifiesto_dt_campo_117.json")
INDICE_INTEGRACION = os.path.join(RAIZ, "datos", "integracion_dt_campo_117.json")
REPORTE_XLSX = os.path.join(RAIZ, "REPORTE_INTEGRACION_DT_CAMPO_117.xlsx")
REPORTE_PDF = os.path.join(RAIZ, "REPORTE_INTEGRACION_DT_CAMPO_117.pdf")


def _drive_por_archivo():
    """Enlaces de Drive del manifiesto anterior, si lo hay, para conservarlos."""
    try:
        with open(MANIFIESTO_CAMPO, encoding="utf-8") as fh:
            previo = json.load(fh)
    except (OSError, ValueError):
        return {}
    return {b.get("archivo", ""): b for b in previo.get("bloques", [])}


def leer_campo():
    """Parsea las plantillas de campo del repositorio."""
    fichas = dtc.fichas_del_repositorio()
    if not fichas:
        sys.exit("No se encontro la carpeta %s." % dtc.CARPETA_CAMPO)
    registros, fallidos = dtc.parsear_lote_campo(fichas)
    por_codigo, duplicados = {}, []
    for registro in registros:
        codigo = registro.get("codigo_bloque")
        if not codigo:
            fallidos.append((registro.get("nombre_archivo", ""),
                             "No declara codigo de bloque."))
        elif codigo in por_codigo:
            duplicados.append((registro.get("nombre_archivo", ""), codigo))
        else:
            por_codigo[codigo] = registro
    return por_codigo, fallidos, duplicados


def escribir_manifiesto_campo(por_codigo):
    """Catalogo de las plantillas de campo, con su enlace de Drive."""
    previo = _drive_por_archivo()
    bloques = []
    for codigo, registro in por_codigo.items():
        nombre = registro.get("nombre_archivo", "")
        ruta = os.path.join(RAIZ, dtc.CARPETA_CAMPO, nombre)
        anterior = previo.get(nombre, {})
        bloques.append({
            "codigo": codigo,
            "archivo": nombre,
            "drive_id": anterior.get("drive_id", ""),
            "drive_url": anterior.get("drive_url", ""),
            "tamano_bytes": os.path.getsize(ruta) if os.path.exists(ruta) else 0,
            "fichas": registro.get("fichas_leidas", []),
            "completitud_pct": registro.get("completitud_pct", 0.0),
        })
    bloques.sort(key=lambda b: dtc._orden_codigo(b["codigo"]))
    manifiesto = {
        "proyecto": "IN Piura - Recuperacion del servicio de regulacion de "
                    "riesgos naturales y de ecosistemas degradados en la "
                    "Cuenca Alta del Rio Piura",
        "cui": "2669244",
        "entidad": "ANIN - DIME - SESDI",
        "documento": "Plantilla DT Campo Check Validada V5 - fichas F-DT-01 "
                     "a F-DT-05",
        "carpeta_drive": "Plantillas Excel FDT actualizadas",
        "carpeta_drive_id": "1ihFqVxhgWIdT8uk88nAKWdJRnsWk_6xD",
        "carpeta_drive_url": "https://drive.google.com/drive/folders/"
                             "1ihFqVxhgWIdT8uk88nAKWdJRnsWk_6xD",
        "carpeta_repositorio": dtc.CARPETA_CAMPO,
        "total_bloques": len(bloques),
        "version": "F-DT actualizadas - verificacion de campo vigente",
        "bloques": bloques,
    }
    _volcar(MANIFIESTO_CAMPO, manifiesto)
    return manifiesto


def escribir_indice(integrados):
    """Indice liviano de la integracion, para consultarla sin abrir libros."""
    indice = {
        "version": dtc.VERSION_INTEGRACION,
        "carpeta_repositorio": dtc.CARPETA_INTEGRADA,
        "totales": dtc.totales(integrados),
        "por_provincia": dtc.agrupar(integrados, "provincia"),
        "por_distrito": dtc.agrupar(integrados, "distrito"),
        "por_microcuenca": dtc.agrupar(integrados, "microcuenca"),
        "bloques": [{
            "codigo": b.get("codigo_bloque", ""),
            "provincia": b.get("provincia", ""),
            "distrito": b.get("distrito", ""),
            "microcuenca": b.get("microcuenca", ""),
            "archivo_campo": b.get("nombre_archivo_campo", ""),
            "archivo_resumen": b.get("nombre_archivo_resumen", ""),
            "cobertura_pct": b.get("cobertura_pct"),
            "conteo_fuente": b.get("conteo_fuente"),
            "conteo_estado": b.get("conteo_estado"),
            "consistencia": b.get("consistencia_resumen"),
            "campos_reescritos": b.get("campos_reescritos", []),
        } for b in integrados],
    }
    _volcar(INDICE_INTEGRACION, indice)
    return indice


def _volcar(ruta, contenido):
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    with open(ruta, "w", encoding="utf-8") as fh:
        json.dump(contenido, fh, ensure_ascii=False, indent=1)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--sin-reportes", action="store_true",
                        help="Solo actualiza los libros y los indices.")
    parser.add_argument("--agrupacion", default="provincia",
                        choices=sorted(dtc.AGRUPACIONES),
                        help="Nivel de agregacion del reporte consolidado.")
    args = parser.parse_args()

    inicio = time.time()
    print("Leyendo plantillas DT de campo...")
    campos, fallidos_campo, duplicados = leer_campo()
    print("  %d fichas de campo leidas." % len(campos))
    for nombre, motivo in fallidos_campo:
        print("  ! %s: %s" % (nombre, motivo))
    for nombre, codigo in duplicados:
        print("  ! %s: el bloque %s ya tenia ficha; se conserva la primera."
              % (nombre, codigo))

    print("Leyendo fichas de resumen V6...")
    # Siempre se parte de los libros de gabinete V6, no de una integracion
    # anterior: asi el registro de lo que la verificacion de campo cambia se
    # calcula contra la misma linea base en cada corrida.
    libros = rbq.libros_del_repositorio(
        os.path.join(RAIZ, rbq.CARPETA_LIBROS_V6))
    if not libros:
        sys.exit("No se encontro la carpeta %s." % rbq.CARPETA_LIBROS_V6)
    print("  %d libros de resumen leidos." % len(libros))

    print("Integrando y actualizando los libros...")
    actualizados, integrados, fallidos = dtc.actualizar_libros(libros, campos)
    for nombre, motivo in fallidos:
        print("  ! %s: %s" % (nombre, motivo))

    os.makedirs(CARPETA_SALIDA, exist_ok=True)
    for nombre, contenido in actualizados:
        with open(os.path.join(CARPETA_SALIDA, nombre), "wb") as fh:
            fh.write(contenido)
    print("  %d libros escritos en %s/" % (len(actualizados),
                                           dtc.CARPETA_INTEGRADA))

    escribir_manifiesto_campo(campos)
    indice = escribir_indice(integrados)
    print("  Indices escritos en datos/.")

    if not args.sin_reportes:
        print("Generando reportes consolidados...")
        with open(REPORTE_XLSX, "wb") as fh:
            fh.write(dtc.generar_excel_consolidado(integrados, args.agrupacion))
        with open(REPORTE_PDF, "wb") as fh:
            fh.write(dtc.generar_pdf_consolidado(integrados, args.agrupacion))
        print("  %s" % os.path.basename(REPORTE_XLSX))
        print("  %s" % os.path.basename(REPORTE_PDF))

    totales = indice["totales"]
    print("\nResumen de la integracion")
    print("  Bloques integrados .............. %d" % totales["n_bloques"])
    print("  Con ficha de campo .............. %d" % totales["bloques_con_campo"])
    print("  Hechos desde campo .............. %d" % totales["hechos_campo"])
    print("  Hechos desde gabinete ........... %d" % totales["hechos_gabinete"])
    print("  Hechos desde fuente oficial ..... %d" % totales["hechos_oficial"])
    print("  Hechos actualizados por campo ... %d" % totales["n_actualizados"])
    print("  Verificaciones de consistencia .. %d" % totales["n_verificaciones"])
    print("  Discrepancias sustantivas ....... %d" % totales["n_sustantivas"])
    print("  Tiempo .......................... %.1f s" % (time.time() - inicio))


if __name__ == "__main__":
    main()
