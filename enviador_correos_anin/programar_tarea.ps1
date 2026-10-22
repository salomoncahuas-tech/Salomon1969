<#
.SYNOPSIS
    Registra el envío de un correo en el Programador de tareas de Windows.

.DESCRIPTION
    Proyecto IN Piura — ANIN / DIME / SESDI

    Crea un lanzador .cmd y registra una tarea programada que lo ejecuta.
    El .cmd queda en la carpeta 'tareas' y puede ejecutarse a mano para
    probar el envío sin esperar la hora programada.

.PARAMETER Correo
    Archivo .ini del correo a enviar. Ej: correos\reporte_semanal.ini

.PARAMETER Nombre
    Nombre de la tarea. Por defecto se deriva del nombre del archivo .ini.

.PARAMETER Frecuencia
    Diaria | Semanal | Mensual | UnaVez

.PARAMETER Hora
    Hora de envío en formato 24 h. Ej: 08:30

.PARAMETER DiaSemana
    Solo para Frecuencia = Semanal. Lunes..Domingo

.PARAMETER Dia
    Solo para Frecuencia = Mensual. Día del mes, 1 a 28.

.PARAMETER Fecha
    Solo para Frecuencia = UnaVez. Formato dd/MM/yyyy. Ej: 15/10/2026

.PARAMETER Eliminar
    Elimina la tarea con el nombre indicado en lugar de crearla.

.EXAMPLE
    .\programar_tarea.ps1 -Correo correos\reporte_semanal.ini -Frecuencia Semanal -DiaSemana Lunes -Hora 08:30

.EXAMPLE
    .\programar_tarea.ps1 -Correo correos\aviso.ini -Frecuencia UnaVez -Fecha 15/10/2026 -Hora 09:00

.EXAMPLE
    .\programar_tarea.ps1 -Nombre "IN Piura - reporte_semanal" -Eliminar
#>

[CmdletBinding()]
param(
    [string]$Correo,
    [string]$Nombre,
    [ValidateSet("Diaria", "Semanal", "Mensual", "UnaVez")]
    [string]$Frecuencia = "Semanal",
    [string]$Hora = "08:30",
    [ValidateSet("Lunes","Martes","Miercoles","Jueves","Viernes","Sabado","Domingo")]
    [string]$DiaSemana = "Lunes",
    [ValidateRange(1, 28)]
    [int]$Dia = 1,
    [string]$Fecha,
    [string]$Python,
    [switch]$NoRepetir,
    [switch]$Eliminar
)

$ErrorActionPreference = "Stop"
$base = Split-Path -Parent $MyInvocation.MyCommand.Path

function Escribir-Titulo {
    Write-Host ""
    Write-Host "===============================================================" -ForegroundColor DarkGreen
    Write-Host " ANIN - DIME - SESDI  |  Proyecto IN Piura"                      -ForegroundColor DarkGreen
    Write-Host " Programador de envios de correo"                                -ForegroundColor DarkGreen
    Write-Host "===============================================================" -ForegroundColor DarkGreen
    Write-Host ""
}

Escribir-Titulo

# --- Eliminar una tarea existente -----------------------------------------
if ($Eliminar) {
    if (-not $Nombre) {
        Write-Host " [ERROR] Indique -Nombre de la tarea a eliminar." -ForegroundColor Red
        Write-Host "         Para ver las tareas creadas:" -ForegroundColor Yellow
        Write-Host '         Get-ScheduledTask | Where-Object { $_.TaskName -like "IN Piura*" }'
        Write-Host ""
        exit 1
    }
    $existente = Get-ScheduledTask -TaskName $Nombre -ErrorAction SilentlyContinue
    if (-not $existente) {
        Write-Host " [AVISO] No existe una tarea llamada '$Nombre'." -ForegroundColor Yellow
        Write-Host ""
        exit 0
    }
    Unregister-ScheduledTask -TaskName $Nombre -Confirm:$false
    Write-Host " [OK] Tarea eliminada: $Nombre" -ForegroundColor Green
    Write-Host ""
    exit 0
}

# --- Validaciones ----------------------------------------------------------
if (-not $Correo) {
    Write-Host " [ERROR] Indique -Correo con la ruta del archivo .ini." -ForegroundColor Red
    Write-Host "         Ejemplo: .\programar_tarea.ps1 -Correo correos\reporte_semanal.ini" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

$rutaCorreo = if ([System.IO.Path]::IsPathRooted($Correo)) { $Correo } else { Join-Path $base $Correo }
if (-not (Test-Path -LiteralPath $rutaCorreo)) {
    Write-Host " [ERROR] No se encuentra el archivo del correo:" -ForegroundColor Red
    Write-Host "         $rutaCorreo"
    Write-Host ""
    exit 1
}
$rutaCorreo = (Resolve-Path -LiteralPath $rutaCorreo).Path

$rutaConfig = Join-Path $base "config.ini"
if (-not (Test-Path -LiteralPath $rutaConfig)) {
    Write-Host " [ERROR] Falta config.ini." -ForegroundColor Red
    Write-Host "         Copie config.ini.ejemplo como config.ini y complete los datos." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

$rutaScript = Join-Path $base "enviar_correo.py"
if (-not (Test-Path -LiteralPath $rutaScript)) {
    Write-Host " [ERROR] No se encuentra enviar_correo.py junto a este script." -ForegroundColor Red
    Write-Host ""
    exit 1
}

if ($Hora -notmatch '^([01]\d|2[0-3]):([0-5]\d)$') {
    Write-Host " [ERROR] -Hora debe tener formato 24 h, por ejemplo 08:30 o 16:45." -ForegroundColor Red
    Write-Host ""
    exit 1
}

# --- Ubicar Python ---------------------------------------------------------
if (-not $Python) {
    $cmd = Get-Command python.exe -ErrorAction SilentlyContinue
    if (-not $cmd) { $cmd = Get-Command py.exe -ErrorAction SilentlyContinue }
    if (-not $cmd) {
        Write-Host " [ERROR] No se encontro Python en el PATH." -ForegroundColor Red
        Write-Host "         Instalelo desde https://www.python.org/downloads/ marcando" -ForegroundColor Yellow
        Write-Host "         la casilla 'Add Python to PATH', o indique la ruta con:"     -ForegroundColor Yellow
        Write-Host '         -Python "C:\Users\HP\AppData\Local\Programs\Python\Python312\python.exe"'
        Write-Host ""
        exit 1
    }
    $Python = $cmd.Source
}
if (-not (Test-Path -LiteralPath $Python)) {
    Write-Host " [ERROR] No existe el ejecutable de Python indicado: $Python" -ForegroundColor Red
    Write-Host ""
    exit 1
}

# --- Nombre de la tarea ----------------------------------------------------
if (-not $Nombre) {
    $Nombre = "IN Piura - " + [System.IO.Path]::GetFileNameWithoutExtension($rutaCorreo)
}

# --- Lanzador .cmd ---------------------------------------------------------
# Se usa un .cmd intermedio a proposito: evita los problemas de comillas al
# pasar rutas con espacios al Programador de tareas, y permite probar el
# envio a mano con doble clic.
$dirTareas = Join-Path $base "tareas"
if (-not (Test-Path -LiteralPath $dirTareas)) {
    New-Item -ItemType Directory -Path $dirTareas | Out-Null
}
$nombreArchivo = ($Nombre -replace '[^A-Za-z0-9_\-]', '_')
$rutaCmd = Join-Path $dirTareas "$nombreArchivo.cmd"

$argNoRepetir = if ($NoRepetir) { " --no-repetir" } else { "" }
$contenidoCmd = @"
@echo off
rem ---------------------------------------------------------------------
rem  ANIN - DIME - SESDI  |  Proyecto IN Piura
rem  Lanzador generado por programar_tarea.ps1
rem  Tarea: $Nombre
rem  Puede ejecutar este archivo a mano para probar el envio.
rem ---------------------------------------------------------------------
cd /d "$base"
"$Python" "$rutaScript" "$rutaCorreo"$argNoRepetir >> "$base\registro\salida_tareas.log" 2>&1
exit /b %errorlevel%
"@

if (-not (Test-Path -LiteralPath (Join-Path $base "registro"))) {
    New-Item -ItemType Directory -Path (Join-Path $base "registro") | Out-Null
}
Set-Content -LiteralPath $rutaCmd -Value $contenidoCmd -Encoding OEM

# --- Disparador ------------------------------------------------------------
$partesHora = $Hora.Split(":")
$h = [int]$partesHora[0]
$m = [int]$partesHora[1]

$mapaDias = @{
    "Lunes" = "Monday"; "Martes" = "Tuesday"; "Miercoles" = "Wednesday"
    "Jueves" = "Thursday"; "Viernes" = "Friday"; "Sabado" = "Saturday"
    "Domingo" = "Sunday"
}

$descripcion = "Proyecto IN Piura (CUI 2669244) - envio automatico de correo institucional. Origen: $([System.IO.Path]::GetFileName($rutaCorreo))"

switch ($Frecuencia) {
    "Diaria" {
        $inicio = (Get-Date -Hour $h -Minute $m -Second 0)
        if ($inicio -lt (Get-Date)) { $inicio = $inicio.AddDays(1) }
        $trigger = New-ScheduledTaskTrigger -Daily -At $inicio
        $cuando = "todos los dias a las $Hora"
    }
    "Semanal" {
        $inicio = (Get-Date -Hour $h -Minute $m -Second 0)
        $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $mapaDias[$DiaSemana] -At $inicio
        $cuando = "cada $DiaSemana a las $Hora"
    }
    "UnaVez" {
        if (-not $Fecha) {
            Write-Host " [ERROR] Con -Frecuencia UnaVez debe indicar -Fecha dd/MM/yyyy." -ForegroundColor Red
            Write-Host ""
            exit 1
        }
        try {
            $fechaBase = [datetime]::ParseExact($Fecha, "dd/MM/yyyy", $null)
        } catch {
            Write-Host " [ERROR] -Fecha invalida: '$Fecha'. Use el formato dd/MM/yyyy." -ForegroundColor Red
            Write-Host ""
            exit 1
        }
        $inicio = $fechaBase.AddHours($h).AddMinutes($m)
        if ($inicio -lt (Get-Date)) {
            Write-Host " [ERROR] La fecha y hora indicadas ya pasaron: $($inicio.ToString('dd/MM/yyyy HH:mm'))" -ForegroundColor Red
            Write-Host ""
            exit 1
        }
        $trigger = New-ScheduledTaskTrigger -Once -At $inicio
        $cuando = "una sola vez el $($inicio.ToString('dd/MM/yyyy')) a las $Hora"
    }
    "Mensual" {
        # New-ScheduledTaskTrigger no soporta disparadores mensuales, por eso
        # se construye la instancia CIM directamente. DaysOfMonth es un mapa
        # de bits: el dia N corresponde al bit (N-1).
        $inicio = (Get-Date -Hour $h -Minute $m -Second 0)
        try {
            $claseCim = Get-CimClass -ClassName MSFT_TaskMonthlyTrigger `
                -Namespace Root/Microsoft/Windows/TaskScheduler
            $trigger = New-CimInstance -CimClass $claseCim -ClientOnly
            $trigger.DaysOfMonth  = [uint32](1 -shl ($Dia - 1))
            $trigger.MonthsOfYear = [uint16]4095      # los 12 meses
            $trigger.StartBoundary = $inicio.ToString("yyyy-MM-ddTHH:mm:ss")
            $trigger.Enabled = $true
        } catch {
            Write-Host " [ERROR] No se pudo crear el disparador mensual en este equipo." -ForegroundColor Red
            Write-Host "         Alternativa manual, igual de valida:" -ForegroundColor Yellow
            Write-Host "         1. Abra 'Programador de tareas' (taskschd.msc)"
            Write-Host "         2. Crear tarea basica > Mensual"
            Write-Host "         3. Como accion, apunte a este archivo ya generado:"
            Write-Host "            $rutaCmd"
            Write-Host ""
            exit 1
        }
        $cuando = "el dia $Dia de cada mes a las $Hora"
    }
}

# --- Registro de la tarea --------------------------------------------------
$accion = New-ScheduledTaskAction -Execute $rutaCmd -WorkingDirectory $base

# StartWhenAvailable recupera el envio si la PC estaba apagada a la hora
# prevista. Por eso conviene usar -NoRepetir en envios recurrentes.
$opciones = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -DontStopIfGoingOnBatteries `
    -AllowStartIfOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
    -MultipleInstances IgnoreNew

$yaExiste = Get-ScheduledTask -TaskName $Nombre -ErrorAction SilentlyContinue
if ($yaExiste) {
    Write-Host " [AVISO] Ya existia una tarea '$Nombre'. Se reemplaza." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $Nombre -Confirm:$false
}

try {
    Register-ScheduledTask -TaskName $Nombre -Action $accion -Trigger $trigger `
        -Settings $opciones -Description $descripcion -RunLevel Limited | Out-Null
} catch {
    Write-Host " [ERROR] No se pudo registrar la tarea:" -ForegroundColor Red
    Write-Host "         $($_.Exception.Message)"
    Write-Host ""
    Write-Host "         Alternativa manual: abra 'Programador de tareas' (taskschd.msc)," -ForegroundColor Yellow
    Write-Host "         cree una tarea basica y como accion apunte a:"                    -ForegroundColor Yellow
    Write-Host "         $rutaCmd"
    Write-Host ""
    exit 1
}

Write-Host " [OK] Tarea programada correctamente." -ForegroundColor Green
Write-Host ""
Write-Host "   Tarea      : $Nombre"
Write-Host "   Se ejecuta : $cuando"
Write-Host "   Correo     : $([System.IO.Path]::GetFileName($rutaCorreo))"
Write-Host "   Lanzador   : $rutaCmd"
Write-Host "   Registro   : $base\registro\"
Write-Host ""
Write-Host " IMPORTANTE:" -ForegroundColor Cyan
Write-Host "   - La PC debe estar encendida y con sesion iniciada a esa hora."
Write-Host "   - Si estaba apagada, el envio se recupera al encenderla."
if (-not $NoRepetir) {
    Write-Host "   - Use -NoRepetir para evitar envios duplicados si la tarea" -ForegroundColor Yellow
    Write-Host "     se dispara dos veces el mismo dia."                       -ForegroundColor Yellow
}
Write-Host ""
Write-Host " Para probar AHORA sin esperar la hora:" -ForegroundColor Cyan
Write-Host "   Start-ScheduledTask -TaskName `"$Nombre`""
Write-Host ""
Write-Host " Para eliminarla:" -ForegroundColor Cyan
Write-Host "   .\programar_tarea.ps1 -Nombre `"$Nombre`" -Eliminar"
Write-Host ""
