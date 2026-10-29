<#
.SYNOPSIS
    Instala Outlook AI: copia el script a tu carpeta de usuario y crea los atajos de teclado.

.DESCRIPTION
    Crea %LOCALAPPDATA%\OutlookAI, copia outlook-ai.ps1 y registra accesos directos con
    atajos Ctrl+Alt+1..4 en el menú Inicio. Los atajos abren el panel de resultado sobre
    Outlook con el correo que tengas seleccionado.

.PARAMETER Uninstall
    Elimina los accesos directos y la carpeta de instalación.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File install.ps1
    powershell -ExecutionPolicy Bypass -File install.ps1 -Uninstall
#>
[CmdletBinding()]
param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'

$targetDir = Join-Path $env:LOCALAPPDATA 'OutlookAI'
$startMenu = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Outlook AI'
$scriptSource = Join-Path $PSScriptRoot 'outlook-ai.ps1'
$scriptTarget = Join-Path $targetDir 'outlook-ai.ps1'

function New-Shortcut {
    param(
        [string]$Path,
        [string]$Arguments,
        [string]$Description,
        [string]$Hotkey
    )
    $shell = New-Object -ComObject WScript.Shell
    $sc = $shell.CreateShortcut($Path)
    $sc.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $sc.Arguments = $Arguments
    $sc.WorkingDirectory = $targetDir
    $sc.Description = $Description
    $sc.WindowStyle = 7   # minimizada: evita el parpadeo de consola
    if ($Hotkey) { $sc.Hotkey = $Hotkey }
    $sc.Save()
}

if ($Uninstall) {
    if (Test-Path -LiteralPath $startMenu) {
        Remove-Item -LiteralPath $startMenu -Recurse -Force
        Write-Host "Accesos directos eliminados: $startMenu"
    }
    if (Test-Path -LiteralPath $targetDir) {
        Remove-Item -LiteralPath $targetDir -Recurse -Force
        Write-Host "Carpeta eliminada: $targetDir"
    }
    Write-Host 'Listo. La clave de API se conserva en %APPDATA%\outlook-ai si la quieres borrar también.'
    exit 0
}

if (-not (Test-Path -LiteralPath $scriptSource)) {
    throw "No encuentro outlook-ai.ps1 junto a este instalador ($PSScriptRoot)."
}

New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
Copy-Item -LiteralPath $scriptSource -Destination $scriptTarget -Force
Write-Host "Script instalado en: $scriptTarget"

if (-not (Test-Path -LiteralPath $startMenu)) { New-Item -ItemType Directory -Force -Path $startMenu | Out-Null }

$base = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptTarget`""

$shortcuts = @(
    @{ Name = '1 - Redactar respuesta';  Args = "$base -Mode reply -Tone formal"; Hotkey = 'Ctrl+Alt+1'; Desc = 'Redacta una respuesta formal al correo seleccionado' },
    @{ Name = '2 - Resumir correo';      Args = "$base -Mode summary";            Hotkey = 'Ctrl+Alt+2'; Desc = 'Resume el correo seleccionado' },
    @{ Name = '3 - Traducir al espanol'; Args = "$base -Mode translate -TargetLang 'espanol'"; Hotkey = 'Ctrl+Alt+3'; Desc = 'Traduce el correo seleccionado al espanol' },
    @{ Name = '4 - Clasificar correo';   Args = "$base -Mode classify";           Hotkey = 'Ctrl+Alt+4'; Desc = 'Clasifica y prioriza el correo seleccionado' },
    @{ Name = '5 - Responder cercano';   Args = "$base -Mode reply -Tone cercano"; Hotkey = 'Ctrl+Alt+5'; Desc = 'Respuesta en tono cercano' },
    @{ Name = '6 - Responder breve';     Args = "$base -Mode reply -Tone breve";   Hotkey = 'Ctrl+Alt+6'; Desc = 'Respuesta muy breve' },
    @{ Name = '7 - Traducir al ingles';  Args = "$base -Mode translate -TargetLang 'ingles'"; Hotkey = 'Ctrl+Alt+7'; Desc = 'Traduce el correo seleccionado al ingles' },
    @{ Name = 'Configurar motor y clave'; Args = "$base -Setup";                  Hotkey = '';           Desc = 'Elige motor (Ollama, Groq, Gemini, DeepSeek) y guarda la clave' }
)

foreach ($sc in $shortcuts) {
    $path = Join-Path $startMenu ($sc.Name + '.lnk')
    New-Shortcut -Path $path -Arguments $sc.Args -Description $sc.Desc -Hotkey $sc.Hotkey
    $hk = if ($sc.Hotkey) { $sc.Hotkey } else { '(sin atajo)' }
    Write-Host ("  {0,-28} {1}" -f $sc.Name, $hk)
}

# eliminar accesos directos de versiones anteriores
$valid = @()
foreach ($sc in $shortcuts) { $valid += ($sc.Name + '.lnk') }
Get-ChildItem -LiteralPath $startMenu -Filter *.lnk -ErrorAction SilentlyContinue | Where-Object {
    $valid -notcontains $_.Name
} | ForEach-Object {
    Remove-Item -LiteralPath $_.FullName -Force
    Write-Host ("  eliminado obsoleto: {0}" -f $_.Name)
}

Write-Host ''
Write-Host 'Instalado.' -ForegroundColor Green
Write-Host '  1. Inicio -> "Outlook AI" -> "Configurar motor y clave" (elige Ollama si no quieres pagar).'
Write-Host '  2. En Outlook, selecciona un correo y pulsa Ctrl+Alt+1..7.'
Write-Host ''
Write-Host 'Nota: si el atajo no responde, abre el menu Inicio y busca "Outlook AI";'
Write-Host 'Windows registra los atajos de los accesos directos del menu Inicio al iniciar sesion.'
