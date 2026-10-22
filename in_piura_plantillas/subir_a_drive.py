# -*- coding: utf-8 -*-
"""
AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN
DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME
SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN - SESDI
PROYECTO IN PIURA | CUI 2669244

Sube cada plantilla Excel generada en «salida/» a la subcarpeta de Google
Drive del bloque al que corresponde, según la correspondencia registrada en
«datos/bloques_dt.json».

Requisitos
----------
    pip install google-api-python-client google-auth-oauthlib

Credenciales (una de las dos vías)
----------------------------------
1. OAuth de usuario (recomendado para la cuenta personal del formulador):
   descargar el «client_secret.json» de un proyecto de Google Cloud con la
   API de Drive habilitada y ejecutar:

       python subir_a_drive.py --credenciales client_secret.json

   La primera ejecución abre el navegador para autorizar; el token queda
   guardado en «token_drive.json» para las siguientes.

2. Cuenta de servicio: exportar la clave JSON y compartir con su correo
   («...iam.gserviceaccount.com») la carpeta raíz de los bloques. Luego:

       python subir_a_drive.py --cuenta-servicio clave_sa.json

Comportamiento
--------------
* Verifica que el tamaño del archivo subido coincida con el del archivo local.
* Es idempotente: si en la subcarpeta ya existe un archivo con el mismo
  nombre, lo ACTUALIZA en lugar de crear un duplicado (salvo --no-actualizar).
* Reintenta hasta 4 veces con espera exponencial ante errores de red.
* Registra el resultado en «subida_drive.log» y en «subida_drive.json».

Uso
---
    python subir_a_drive.py --credenciales client_secret.json
    python subir_a_drive.py --credenciales client_secret.json --solo 38 M22B1
    python subir_a_drive.py --credenciales client_secret.json --simular
"""

import argparse
import json
import os
import sys
import time

BASE = os.path.dirname(os.path.abspath(__file__))
DATOS = os.path.join(BASE, "datos")
SALIDA = os.path.join(BASE, "salida")
MIME_XLSX = ("application/vnd.openxmlformats-officedocument"
             ".spreadsheetml.sheet")
ALCANCE = ["https://www.googleapis.com/auth/drive"]
PLANTILLA_NOMBRE = "Plantilla_Excel_Bloque_{}_IN_Piura.xlsx"


def registrar(fh, texto):
    print(texto)
    fh.write(texto + "\n")
    fh.flush()


def servicio(args):
    from googleapiclient.discovery import build

    if args.cuenta_servicio:
        from google.oauth2 import service_account
        cred = service_account.Credentials.from_service_account_file(
            args.cuenta_servicio, scopes=ALCANCE)
    else:
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        cred = None
        if os.path.exists(args.token):
            cred = Credentials.from_authorized_user_file(args.token, ALCANCE)
        if not cred or not cred.valid:
            if cred and cred.expired and cred.refresh_token:
                cred.refresh(Request())
            else:
                flujo = InstalledAppFlow.from_client_secrets_file(
                    args.credenciales, ALCANCE)
                cred = flujo.run_local_server(port=0)
            with open(args.token, "w", encoding="utf-8") as fh:
                fh.write(cred.to_json())
    return build("drive", "v3", credentials=cred, cache_discovery=False)


def con_reintentos(fn, intentos=4):
    espera = 2
    for i in range(intentos):
        try:
            return fn()
        except Exception as exc:            # noqa: BLE001
            if i == intentos - 1:
                raise
            print("  reintento %d tras error: %s" % (i + 1, exc))
            time.sleep(espera)
            espera *= 2


def existente(srv, carpeta, nombre):
    consulta = ("name = %s and '%s' in parents and trashed = false"
                % (json.dumps(nombre), carpeta))
    res = con_reintentos(lambda: srv.files().list(
        q=consulta, fields="files(id,name,size)", pageSize=10,
        supportsAllDrives=True, includeItemsFromAllDrives=True).execute())
    archivos = res.get("files", [])
    return archivos[0] if archivos else None


def subir(srv, ruta, carpeta, nombre, actualizar=True):
    from googleapiclient.http import MediaFileUpload
    medio = MediaFileUpload(ruta, mimetype=MIME_XLSX, resumable=True)
    previo = existente(srv, carpeta, nombre) if actualizar else None
    if previo:
        return con_reintentos(lambda: srv.files().update(
            fileId=previo["id"], media_body=medio,
            fields="id,name,size,webViewLink",
            supportsAllDrives=True).execute()), "actualizado"
    cuerpo = {"name": nombre, "parents": [carpeta], "mimeType": MIME_XLSX}
    return con_reintentos(lambda: srv.files().create(
        body=cuerpo, media_body=medio, fields="id,name,size,webViewLink",
        supportsAllDrives=True).execute()), "creado"


def main():
    p = argparse.ArgumentParser(
        description="Sube las plantillas Excel de IN Piura a Google Drive.")
    p.add_argument("--credenciales", help="client_secret.json (OAuth)")
    p.add_argument("--cuenta-servicio", help="clave JSON de cuenta de servicio")
    p.add_argument("--token", default=os.path.join(BASE, "token_drive.json"),
                   help="ruta del token OAuth persistido")
    p.add_argument("--solo", nargs="*", metavar="BLOQUE",
                   help="subir únicamente estos bloques")
    p.add_argument("--simular", action="store_true",
                   help="no sube nada; solo muestra el plan")
    p.add_argument("--no-actualizar", action="store_true",
                   help="crear siempre un archivo nuevo aunque ya exista")
    args = p.parse_args()

    if not args.simular and not (args.credenciales or args.cuenta_servicio):
        p.error("indique --credenciales o --cuenta-servicio (o use --simular)")

    with open(os.path.join(DATOS, "bloques_dt.json"), encoding="utf-8") as fh:
        bloques = json.load(fh)
    if args.solo:
        pedidos = set(args.solo)
        bloques = [b for b in bloques if b[0] in pedidos]
        faltan = pedidos - {b[0] for b in bloques}
        if faltan:
            sys.exit("bloques sin carpeta registrada: %s"
                     % ", ".join(sorted(faltan)))

    srv = None if args.simular else servicio(args)
    resultados, errores = [], 0
    with open(os.path.join(BASE, "subida_drive.log"), "w",
              encoding="utf-8") as log:
        registrar(log, "PROYECTO IN PIURA | CUI 2669244 | ANIN-DIME-SESDI")
        registrar(log, "Subida de %d plantillas Excel a Google Drive"
                  % len(bloques))
        registrar(log, "-" * 72)
        for i, (bloque, carpeta, titulo) in enumerate(bloques, 1):
            nombre = PLANTILLA_NOMBRE.format(bloque)
            ruta = os.path.join(SALIDA, nombre)
            if not os.path.exists(ruta):
                registrar(log, "[%3d/%3d] %-8s FALTA el archivo %s"
                          % (i, len(bloques), bloque, nombre))
                errores += 1
                continue
            tam = os.path.getsize(ruta)
            if args.simular:
                registrar(log, "[%3d/%3d] %-8s %s -> %s (%s, %d B)"
                          % (i, len(bloques), bloque, nombre, carpeta,
                             titulo, tam))
                continue
            try:
                arch, accion = subir(srv, ruta, carpeta, nombre,
                                     actualizar=not args.no_actualizar)
            except Exception as exc:        # noqa: BLE001
                registrar(log, "[%3d/%3d] %-8s ERROR: %s"
                          % (i, len(bloques), bloque, exc))
                errores += 1
                continue
            subido = int(arch.get("size") or 0)
            ok = subido == tam
            if not ok:
                errores += 1
            registrar(log, "[%3d/%3d] %-8s %-11s %d B %s  %s"
                      % (i, len(bloques), bloque, accion, subido,
                         "OK" if ok else "TAMAÑO NO COINCIDE (local %d)" % tam,
                         arch.get("webViewLink", "")))
            resultados.append({
                "bloque": bloque, "carpeta": carpeta, "carpeta_titulo": titulo,
                "archivo": nombre, "id_drive": arch.get("id"),
                "bytes_local": tam, "bytes_drive": subido,
                "accion": accion, "verificado": ok,
                "enlace": arch.get("webViewLink"),
            })
        registrar(log, "-" * 72)
        registrar(log, "Subidas correctas: %d | incidencias: %d"
                  % (len(resultados) - errores if resultados else 0, errores))

    if resultados:
        with open(os.path.join(BASE, "subida_drive.json"), "w",
                  encoding="utf-8") as fh:
            json.dump(resultados, fh, ensure_ascii=False, indent=1)
    return 1 if errores else 0


if __name__ == "__main__":
    sys.exit(main())
