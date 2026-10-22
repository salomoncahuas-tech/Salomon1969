<#
.SYNOPSIS
    Guarda la contraseña SMTP cifrada con DPAPI de Windows.

.DESCRIPTION
    Proyecto IN Piura — ANIN / DIME / SESDI

    Pide la contraseña de aplicación de la cuenta institucional y la guarda
    cifrada en un archivo. El cifrado DPAPI ata el archivo a SU usuario de
    Windows y a ESTE equipo: ninguna otra persona ni otra PC puede descifrarlo,
    aunque copie el archivo.

    Después de ejecutarlo, en config.ini descomente la línea:
        clave_archivo = clave_smtp.txt

.EXAMPLE
    .\guardar_clave.ps1
#>

[CmdletBinding()]
param(
    [string]$Archivo = "clave_smtp.txt"
)

$ErrorActionPreference = "Stop"
$base = Split-Path -Parent $MyInvocation.MyCommand.Path
$destino = if ([System.IO.Path]::IsPathRooted($Archivo)) { $Archivo } else { Join-Path $base $Archivo }

Write-Host ""
Write-Host "===============================================================" -ForegroundColor DarkGreen
Write-Host " ANIN - DIME - SESDI  |  Proyecto IN Piura"                      -ForegroundColor DarkGreen
Write-Host " Guardar contrasena SMTP cifrada (DPAPI)"                        -ForegroundColor DarkGreen
Write-Host "===============================================================" -ForegroundColor DarkGreen
Write-Host ""
Write-Host " Si su correo esta en Google Workspace, aqui va la CONTRASENA"
Write-Host " DE APLICACION de 16 caracteres, no su contrasena normal."
Write-Host " Generela en: https://myaccount.google.com/apppasswords"
Write-Host ""
Write-Host " Lo que escriba no se muestra en pantalla." -ForegroundColor Yellow
Write-Host ""

$clave = Read-Host -Prompt " Contrasena" -AsSecureString
if ($clave.Length -eq 0) {
    Write-Host ""
    Write-Host " [ERROR] No escribio nada. No se guardo el archivo." -ForegroundColor Red
    Write-Host ""
    exit 1
}

$confirmacion = Read-Host -Prompt " Repita la contrasena" -AsSecureString

# Comparacion segura: se descifra en memoria solo para verificar y se limpia.
$p1 = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($clave)
$p2 = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($confirmacion)
try {
    $texto1 = [Runtime.InteropServices.Marshal]::PtrToStringAuto($p1)
    $texto2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto($p2)
    $iguales = ($texto1 -ceq $texto2)
    $texto1 = $null
    $texto2 = $null
} finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($p1)
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($p2)
}

if (-not $iguales) {
    Write-Host ""
    Write-Host " [ERROR] Las contrasenas no coinciden. No se guardo nada." -ForegroundColor Red
    Write-Host ""
    exit 1
}

$clave | ConvertFrom-SecureString | Set-Content -LiteralPath $destino -Encoding ASCII -NoNewline

Write-Host ""
Write-Host " [OK] Clave cifrada guardada en:" -ForegroundColor Green
Write-Host "      $destino"
Write-Host ""
Write-Host " Ahora edite config.ini y deje esta linea activa:" -ForegroundColor Cyan
Write-Host "      clave_archivo = $Archivo"
Write-Host ""
Write-Host " Verifique con:" -ForegroundColor Cyan
Write-Host "      python enviar_correo.py --probar-conexion"
Write-Host ""
