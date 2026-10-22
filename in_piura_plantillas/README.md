# Plantillas Excel de Diagnóstico Territorial por bloque — Proyecto IN Piura

**AUTORIDAD NACIONAL DE INFRAESTRUCTURA - ANIN**
**DIRECCIÓN DE INTERVENCIONES MULTISECTORIALES Y DE EMERGENCIA - DIME**
**SUBDIRECCIÓN DE ESTUDIOS DE INVERSIÓN - SESDI**

Proyecto IN Piura | CUI 2669244 | UTM WGS 84 Zona 17S

---

## 1. Qué contiene

Un libro Excel por cada uno de los **117 bloques de intervención que cuentan
con archivo de Diagnóstico Territorial**, con cinco hojas de cálculo:

| Hoja | Contenido | Modelo |
|---|---|---|
| 1. Resumen | Identificación, ubicación política, parámetros geoespaciales, ecosistema, erosión, causas de degradación, recursos hídricos y accesibilidad | Bloque 38 |
| 2. Cobertura MSAVI-NDVI | Clases MSAVI 2024 frente al umbral de brecha 0.4976 (R.M. N.° 00213-2024-MINAM) y distribución de clases NDVI 2025 con metrado en hectáreas | Bloque 38 |
| 3. Estaciones fotográficas | Estaciones georreferenciadas, distancia al centroide y verificación de pertenencia al polígono | Bloque 38 |
| 4. Microcuenca | Indicadores agregados a nivel de microcuenca y posición relativa del bloque | Bloque 38 |
| 5. Control de consistencia | Verificaciones y discrepancias calificadas como SUSTANTIVA / NO SUSTANTIVA / CORREGIDO / CONFORME | Bloque 51 |

## 2. Estructura del directorio

```
in_piura_plantillas/
├── generar_plantillas.py                 generador (openpyxl)
├── subir_a_drive.py                      carga automática a Google Drive
├── correspondencia_bloque_carpeta.csv    bloque → carpeta de Drive
├── datos/                                insumos de gabinete
│   ├── catalogo_v5.json                  catálogo Bloques_V5 (área, UTM, MSAVI)
│   ├── alt_pend.json                     altitud y pendiente por MDE
│   ├── ndvi.json                         clases NDVI 2025 y metrado
│   ├── centros_poblados.json             cruce INEI-Bloques V5
│   └── bloques_dt.json                   bloque → carpeta de Drive
├── dt/<bloque>.json                      fichas F-DT-01 a F-DT-05 extraídas
└── salida/                               los 117 libros Excel generados
```

## 3. Regenerar los libros

```bash
pip install openpyxl
python generar_plantillas.py            # los 117
python generar_plantillas.py 38 M22B1   # solo algunos
```

## 4. Subir los libros a Google Drive

Cada libro debe alojarse en la subcarpeta de Drive del bloque al que
corresponde. La correspondencia está en `datos/bloques_dt.json` y, en formato
legible, en `correspondencia_bloque_carpeta.csv`.

```bash
pip install google-api-python-client google-auth-oauthlib
python subir_a_drive.py --simular                              # plan, sin subir
python subir_a_drive.py --credenciales client_secret.json      # subida completa
python subir_a_drive.py --credenciales client_secret.json --solo 38 M22B1
```

El script verifica que el tamaño del archivo en Drive coincida con el local,
actualiza en lugar de duplicar si el archivo ya existe, reintenta ante errores
de red y deja constancia en `subida_drive.log` y `subida_drive.json`.

## 5. Criterio de integridad declarativa

**No se estima ni se infiere ningún valor ausente.** Los vacíos de las fichas
se consignan como «Por determinar», «Por verificar» o «Sin registro en ficha».
Las inconsistencias entre formatos se conservan sin promediar y quedan
registradas en la hoja 5 de cada libro, junto con las verificaciones
automáticas de coordenadas, superficie, altitud y centro poblado.

## 6. Umbrales aplicados

* **Brecha de degradación (MSAVI 2024):** 0.4976 — R.M. N.° 00213-2024-MINAM.
* **Clases MSAVI:** > 0.6139 vigor alto; 0.4976–0.6139 moderado;
  0.3813–0.4976 bajo; 0.2650–0.3813 muy bajo; ≤ 0.2650 suelo desnudo.
* **Validación UTM 17S:** ESTE 450 000–750 000 m; NORTE 9 300 000–9 600 000 m.
