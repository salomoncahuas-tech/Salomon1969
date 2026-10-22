# Enviador de correos institucionales ANIN

**Proyecto IN Piura** — CUI 2669244
Autoridad Nacional de Infraestructura — ANIN
Dirección de Intervenciones Multisectoriales y de Emergencia — DIME
Subdirección de Estudios de Inversión — SESDI

Envía correos programados **desde la cuenta institucional** (`anin_690@dime.gob.pe`),
con plantilla HTML de identidad ANIN, ejecutándose en su PC mediante el
Programador de tareas de Windows.

---

## Por qué este script y no Claude

Claude no puede enviar correos desde su cuenta institucional: el conector de
Gmail está autenticado únicamente como `salomoncahuas@gmail.com` y no permite
cambiar el remitente. Este script corre en **su** equipo, con **sus**
credenciales institucionales, y por lo tanto el correo sale siempre desde
`@dime.gob.pe`.

**Su contraseña nunca se escribe en ningún archivo de este repositorio.**

---

## Requisitos

| Requisito | Detalle |
|---|---|
| Python 3.8 o superior | https://www.python.org/downloads/ — marque **"Add Python to PATH"** al instalar |
| Bibliotecas externas | **Ninguna.** El script usa solo la biblioteca estándar |
| Windows | 10 u 11, con PowerShell (viene instalado) |
| Cuenta institucional | Con SMTP habilitado (ver Paso 1) |

---

## Instalación en 5 pasos

### Paso 1 — Identificar la plataforma del correo institucional

```bat
cd ruta\al\repositorio\enviador_correos_anin
python enviar_correo.py --detectar dime.gob.pe
```

El resultado le dirá si debe usar `perfil = google` o `perfil = microsoft`.

Si el comando no puede consultar el DNS, use el método manual: abra su correo
institucional en el navegador y mire la barra de direcciones.

| Lo que ve | Perfil | Servidor |
|---|---|---|
| `mail.google.com` | `google` | smtp.gmail.com:587 |
| `outlook.office.com` | `microsoft` | smtp.office365.com:587 |
| Otro (webmail propio) | `manual` | Pida host y puerto a TI |

### Paso 2 — Crear el archivo de configuración

```bat
copy config.ini.ejemplo config.ini
notepad config.ini
```

Complete el `perfil` del Paso 1 y los datos de su firma.
**Guarde el archivo en codificación UTF-8** (Bloc de notas: *Archivo > Guardar como > Codificación: UTF-8*).

### Paso 3 — Guardar la contraseña de forma segura

Elija **una** de las dos opciones.

#### Opción A — Archivo cifrado (recomendada)

```powershell
.\guardar_clave.ps1
```

Escriba la contraseña cuando se la pida (no se muestra en pantalla). Se guarda
cifrada con **DPAPI de Windows**: queda atada a su usuario y a ese equipo, de
modo que nadie más puede descifrarla aunque copie el archivo.

Luego, en `config.ini`, deje activa la línea:
```ini
clave_archivo = clave_smtp.txt
```

#### Opción B — Variable de entorno

```bat
setx ANIN_SMTP_PASSWORD "su-contraseña-de-aplicación"
```
Cierre y vuelva a abrir la consola para que tome efecto.

> **Si su correo está en Google Workspace**, la contraseña NO es la de su
> cuenta: es una **contraseña de aplicación** de 16 caracteres que se genera en
> https://myaccount.google.com/apppasswords (requiere verificación en 2 pasos
> activada). Si esa página no aparece, el administrador de ANIN la tiene
> restringida y debe solicitarla a TI.
>
> **Si su correo está en Microsoft 365**, es probable que la autenticación SMTP
> básica esté desactivada por política. En ese caso TI debe habilitar
> *Authenticated SMTP* para su buzón. Es un cambio de un minuto en el Centro de
> administración.

### Paso 4 — Verificar la conexión

```bat
python enviar_correo.py --probar-conexion
```

No envía ningún correo: solo comprueba servidor, usuario y contraseña.
Si algo falla, el mensaje de error le indica exactamente qué corregir.

### Paso 5 — Preparar y probar su primer correo

```bat
copy correos\reporte_semanal.ini.ejemplo correos\reporte_semanal.ini
notepad correos\reporte_semanal.ini
python enviar_correo.py correos\reporte_semanal.ini --prueba
```

`--prueba` **no envía nada**: genera una vista previa HTML y la abre en el
navegador para que revise cómo se verá. Cuando esté conforme, quite `--prueba`.

---

## Cómo redactar un correo

Cada correo es un archivo `.ini` dentro de la carpeta `correos\`. No necesita
saber HTML.

```ini
[mensaje]
para = destinatario@anin.gob.pe, otro@anin.gob.pe
copia = anin_690@dime.gob.pe
asunto = Reporte semanal N° {{SEMANA}} — Proyecto IN Piura

cuerpo =
    Estimados señores:

    Remito el avance de la **semana {{SEMANA}}**.

    ## 1. Avance de campo

    - Bloques verificados: **8**
    - Superficie: 7,104 ha

    Atentamente,
```

### Reglas de formato

| Escriba | Resultado |
|---|---|
| `## Título` | Subtítulo en verde ANIN con línea inferior |
| `- texto` | Lista con viñetas |
| `1. texto` | Lista numerada |
| `**texto**` | **Negrita** |
| `*texto*` | *Cursiva* |
| Línea en blanco | Separa párrafos |

**Importante:**

- Todas las líneas del `cuerpo` deben ir **indentadas** con espacios al inicio.
- Un párrafo o un ítem de lista puede ocupar varias líneas: se unen solos.
- Los comentarios del archivo `.ini` empiezan con **punto y coma** `;`
  (no con `#`, porque `##` se usa para los subtítulos).

### Variables

Se reemplazan escribiéndolas entre llaves dobles: `{{SEMANA}}`.

**Automáticas** (no hay que definirlas):

`{{FECHA}}` `{{FECHA_LARGA}}` `{{FECHA_ISO}}` `{{HORA}}` `{{DIA}}`
`{{DIA_SEMANA}}` `{{MES}}` `{{MES_NUM}}` `{{ANIO}}` `{{SEMANA}}` `{{TRIMESTRE}}`

**Propias**, en la sección `[variables]` del archivo del correo:

```ini
[variables]
BLOQUES_ACUMULADO = 62
MICROCUENCA = M5
```

O al momento de ejecutar, lo que tiene prioridad:

```bat
python enviar_correo.py correos\reporte_semanal.ini --variable BLOQUES_ACUMULADO=70
```

### Adjuntos

```ini
[adjuntos]
archivo1 = C:\Users\HP\Documents\IN_Piura\Reporte.xlsx
archivo2 = ..\reportes\Fichas_FDT.pdf
```

Rutas absolutas o relativas al archivo `.ini`. **Límite total: 25 MB.**
Para archivos más pesados, suba a Drive/SharePoint y pegue el enlace en el cuerpo.

---

## Programar el envío automático

```powershell
# Todos los lunes a las 08:30
.\programar_tarea.ps1 -Correo correos\reporte_semanal.ini -Frecuencia Semanal -DiaSemana Lunes -Hora 08:30 -NoRepetir

# Todos los días a las 07:45
.\programar_tarea.ps1 -Correo correos\aviso_diario.ini -Frecuencia Diaria -Hora 07:45 -NoRepetir

# El día 1 de cada mes a las 09:00
.\programar_tarea.ps1 -Correo correos\reporte_mensual.ini -Frecuencia Mensual -Dia 1 -Hora 09:00 -NoRepetir

# Una sola vez, en fecha y hora exactas
.\programar_tarea.ps1 -Correo correos\convocatoria.ini -Frecuencia UnaVez -Fecha 15/10/2026 -Hora 09:00
```

### Administrar las tareas creadas

```powershell
# Ver las tareas del proyecto
Get-ScheduledTask | Where-Object { $_.TaskName -like "IN Piura*" }

# Ejecutar ahora, sin esperar la hora programada
Start-ScheduledTask -TaskName "IN Piura - reporte_semanal"

# Eliminar una tarea
.\programar_tarea.ps1 -Nombre "IN Piura - reporte_semanal" -Eliminar
```

### Advertencias sobre las tareas programadas

> ⚠️ **La PC debe estar encendida y con la sesión iniciada** a la hora prevista.
> Si estaba apagada, la tarea se recupera al encenderla (opción
> *StartWhenAvailable*). Por eso conviene usar siempre **`-NoRepetir`** en
> envíos recurrentes: evita que un correo salga dos veces el mismo día.
>
> Si necesita que el envío salga **sí o sí a una hora exacta con la PC
> apagada**, este método no sirve; ahí corresponde el "Programar envío" de
> Gmail/Outlook o un servidor de la entidad.

---

## Solución de problemas

| Mensaje de error | Qué significa y cómo se resuelve |
|---|---|
| `El servidor rechazo el usuario o la clave` | En Google: use contraseña de aplicación, no la normal. En Microsoft: pida a TI habilitar *Authenticated SMTP* |
| `5.7.139 basic authentication is disabled` | Microsoft 365 bloquea SMTP básico. Requiere gestión de TI |
| `No hay clave disponible` | No ejecutó `guardar_clave.ps1` ni configuró la variable de entorno (Paso 3) |
| `No se pudo descifrar el archivo de clave` | DPAPI ata el archivo al usuario y equipo. Si cambió de PC o de usuario, ejecute `guardar_clave.ps1` de nuevo |
| `El archivo no esta guardado en UTF-8` | Vuelva a guardarlo en el Bloc de notas con codificación UTF-8. Los acentos y la Ñ lo requieren |
| `Fallo la conexion` | Sin internet, o el cortafuegos de la entidad bloquea el puerto 587 de salida |
| `Los adjuntos superan el limite de 25 MB` | Use un enlace de Drive/SharePoint en el cuerpo |
| `El servidor rechazo el remitente` | El campo `remitente` debe ser igual a `usuario`, o un alias autorizado |
| La tarea programada no envía nada | Revise `registro\salida_tareas.log`. Pruebe el lanzador a mano: doble clic en `tareas\<nombre>.cmd` |

### Registros

| Archivo | Contenido |
|---|---|
| `registro\envios_AAAA-MM.log` | Un renglón por correo enviado: fecha, archivo, asunto y destinatarios |
| `registro\salida_tareas.log` | Salida completa de las ejecuciones automáticas, con los errores |

---

## Estructura de archivos

```
enviador_correos_anin\
├── enviar_correo.py            Script principal (no requiere edición)
├── plantilla_anin.html         Plantilla HTML con identidad ANIN
├── config.ini.ejemplo          Modelo de configuración → copiar a config.ini
├── guardar_clave.ps1           Guarda la contraseña cifrada con DPAPI
├── programar_tarea.ps1         Registra la tarea en el Programador de Windows
├── LEEME.md                    Este documento
├── correos\
│   └── reporte_semanal.ini.ejemplo   Modelo de correo → copiar y editar
├── registro\                   (se crea solo) Bitácora de envíos
├── marcadores\                 (se crea solo) Control anti-duplicado
├── pruebas\                    (se crea solo) Vistas previas HTML
└── tareas\                     (se crea solo) Lanzadores .cmd
```

Los archivos `config.ini`, `clave_smtp.txt` y `correos\*.ini` están excluidos
del repositorio mediante `.gitignore`: contienen credenciales y direcciones de
destinatarios, y no deben subirse a GitHub.

---

## Seguridad

- La contraseña no aparece en ningún archivo versionado, ni en el código, ni
  en la línea de comandos, ni en los registros.
- Con la Opción A, la contraseña se almacena cifrada con DPAPI de Windows y
  solo puede descifrarla su usuario en ese equipo.
- Toda la comunicación con el servidor usa **TLS** con verificación de
  certificado (`ssl.create_default_context()`).
- Si alguna vez expone la contraseña, revóquela de inmediato:
  Google → https://myaccount.google.com/apppasswords (eliminar la contraseña de
  aplicación) · Microsoft → cambiar la contraseña de la cuenta.
