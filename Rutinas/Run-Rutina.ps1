<#
.SYNOPSIS
    Lanza una skill programada con "claude -p", sin supervision.

.DESCRIPTION
    Reemplaza a run-*.sh + cron, que era el mecanismo en Ubuntu. Un solo
    script parametrizado cubre las dos rutinas, en vez de un .sh por cada
    una: la unica diferencia entre ellos era el nombre de la skill.

    Igual que el cron anterior, la Tarea Programada llama a este script
    cada 30 minutos y no solo a la hora objetivo, para cubrir el caso de
    que la PC este apagada a esa hora. El marcador .last-success-<skill>
    evita repetir la corrida el mismo dia una vez que ya salio bien.

.PARAMETER Skill
    Nombre de la carpeta de la skill dentro de Claude\Scheduled.

.PARAMETER SoloDiasHabiles
    Si se indica, no hace nada sabados ni domingos (igual que el cron).

.PARAMETER DryRun
    Verifica todo el cableado (skill, CLI, prompt, marcador) e informa que
    haria, sin lanzar Claude ni escribir el marcador ni el log. Sirve para
    comprobar la rutina sin disparar una corrida autonoma de verdad.

.PARAMETER Force
    Ignora el marcador diario y corre aunque la rutina ya haya salido bien
    hoy. Lo usan los botones de "Rutinas automaticas" del panel: ahi la
    corrida la pide una persona a proposito, y sin esto el script saldria en
    silencio con exit 0 y desde el boton pareceria que no hizo nada
    (9-sep-2026). La Tarea Programada NO lo pasa: para ella el marcador es
    justamente lo que evita repetir el trabajo el mismo dia.

.EXAMPLE
    .\Run-Rutina.ps1 -Skill daily-client-presentation-prep -SoloDiasHabiles

.EXAMPLE
    .\Run-Rutina.ps1 -Skill daily-client-presentation-prep -DryRun
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Skill,

    [switch]$SoloDiasHabiles,

    [switch]$DryRun,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# La raiz del proyecto se deduce de la ubicacion de este script (esta en
# Rutinas\, o sea un nivel bajo la raiz), no de una ruta escrita a mano.
$RutinasDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$GBClaude   = Split-Path -Parent $RutinasDir
$LogsDir    = Join-Path $RutinasDir 'logs'
$SkillFile  = Join-Path $GBClaude "Claude\Scheduled\$Skill\SKILL.md"
$MarkerFile = Join-Path $LogsDir ".last-success-$Skill"
$LogFile    = Join-Path $LogsDir "$Skill.log"

if (-not (Test-Path $LogsDir)) { New-Item -ItemType Directory -Force -Path $LogsDir | Out-Null }

function Write-Log([string]$Mensaje) {
    if ($DryRun) { Write-Host $Mensaje; return }
    # UTF-8 explicito: Add-Content por defecto escribe en la codificacion
    # ANSI del sistema y destroza los acentos de los mensajes.
    Add-Content -Path $LogFile -Value $Mensaje -Encoding utf8
}

$Hoy = Get-Date -Format 'yyyy-MM-dd'

if ($SoloDiasHabiles) {
    $dow = (Get-Date).DayOfWeek
    if ($dow -eq 'Saturday' -or $dow -eq 'Sunday') {
        if ($DryRun) { Write-Host "Fin de semana ($dow): la rutina no correria." }
        exit 0
    }
}

# Ya corrio bien hoy: no repetir, salvo que se pida a proposito con -Force.
if (-not $Force -and (Test-Path $MarkerFile)) {
    $marca = (Get-Content $MarkerFile -Raw -Encoding utf8).Trim()
    if ($marca -eq $Hoy) {
        if ($DryRun) { Write-Host "Ya corrio bien hoy ($Hoy): la rutina no correria de nuevo." }
        exit 0
    }
}
if ($Force -and $DryRun) {
    Write-Host "-Force: se ignora el marcador, correria aunque ya hubiera corrido hoy."
}

$inicio = Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz'
Write-Log "=== $inicio - inicio ($Skill) ==="

if (-not (Test-Path $SkillFile)) {
    Write-Log "ERROR: no se encontro $SkillFile (se movio la carpeta?). Abortando sin ejecutar nada."
    exit 1
}

# El CLI puede estar en el PATH o en las rutas habituales del instalador.
# No basta con Get-Command: si este proceso heredo un PATH anterior a la
# instalacion del CLI, no lo ve aunque este bien instalado.
$claude = (Get-Command claude -ErrorAction SilentlyContinue).Source
if (-not $claude) {
    $candidatos = @(
        "$env:LOCALAPPDATA\Microsoft\WinGet\Links\claude.exe",
        "$env:USERPROFILE\.local\bin\claude.exe",
        "$env:APPDATA\npm\claude.cmd",
        "$env:LOCALAPPDATA\Programs\claude\claude.exe"
    )
    # Instalacion por winget: la carpeta del paquete lleva un sufijo de hash.
    $candidatos += (Get-ChildItem "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" `
        -Filter 'Anthropic.ClaudeCode*' -Directory -ErrorAction SilentlyContinue |
        ForEach-Object { Join-Path $_.FullName 'claude.exe' })

    foreach ($c in $candidatos) { if ($c -and (Test-Path $c)) { $claude = $c; break } }
}
if (-not $claude) {
    Write-Log "ERROR: no se encontro el CLI de Claude Code. Instalalo o agregalo al PATH."
    exit 1
}

# Salta el frontmatter YAML (4 lineas: ---, name, description, ---), igual
# que el "tail -n +5" del script de bash. Lectura UTF-8 explicita porque los
# SKILL.md tienen acentos.
$lineas = Get-Content -Path $SkillFile -Encoding utf8
if ($lineas.Count -le 4) {
    Write-Log "ERROR: $SkillFile no tiene contenido despues del frontmatter."
    exit 1
}
$prompt = ($lineas[4..($lineas.Count - 1)] -join "`n")

if ($DryRun) {
    Write-Host "Skill      : $SkillFile"
    Write-Host "CLI        : $claude"
    Write-Host "Directorio : $GBClaude"
    Write-Host "Log        : $LogFile"
    Write-Host "Marcador   : $MarkerFile"
    Write-Host "Prompt     : $($prompt.Length) caracteres, $($lineas.Count - 4) lineas"
    Write-Host ''
    Write-Host 'Primeras 3 lineas del prompt:'
    ($prompt -split "`n" | Select-Object -First 3) | ForEach-Object { Write-Host "  | $_" }
    Write-Host ''
    Write-Host 'DRY-RUN: no se lanzo Claude, no se escribio el log ni el marcador.'
    exit 0
}

# Para que el CLI y la skill escriban/lean acentos sin romperse.
$env:PYTHONIOENCODING = 'utf-8'

# $OutputEncoding es lo que PowerShell usa para codificar lo que ENTUBA hacia
# un ejecutable nativo, y [Console]::OutputEncoding para decodificar lo que
# ese ejecutable devuelve. Sin las dos en UTF-8, los acentos del prompt salen
# rotos y el log queda con "DEFINICI├ôN" en vez de "DEFINICIÓN" (9-sep-2026).
$OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$EncodingPrevia = [Console]::OutputEncoding
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

Push-Location $GBClaude
try {
    # El prompt va por STDIN, no como argumento de -p. PowerShell 5.1 rompe el
    # paso de argumentos a ejecutables nativos cuando el texto trae comillas
    # dobles: el 9-sep-2026 esta rutina corrio y de los 13.942 caracteres del
    # SKILL.md solo le llegaron 671 al CLI, cortados justo en la primera
    # comilla (offset 664). Claude respondio pidiendo el mensaje completo, y
    # como salio con codigo 0 la rutina se dio por exitosa y escribio el
    # marcador sin haber hecho nada. Medido: por argumento llegan 671 chars,
    # por stdin llegan los 13.942.
    $prompt | & $claude -p --dangerously-skip-permissions 2>&1 |
        ForEach-Object { Write-Log $_ }
    $code = $LASTEXITCODE
}
finally {
    Pop-Location
    [Console]::OutputEncoding = $EncodingPrevia
}

$fin = Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz'
Write-Log "=== $fin - fin (exit $code) ==="

if ($code -eq 0) {
    # WriteAllText con UTF8Encoding($false) en vez de Set-Content -Encoding
    # utf8: en PowerShell 5.1 ese parametro escribe BOM, y el panel lee el
    # marcador con Python, que NO lo quita -- comparaba "﻿2026-09-09"
    # contra "2026-09-09" y el semaforo se quedaba en amarillo aunque la
    # rutina hubiera corrido bien (9-sep-2026).
    [System.IO.File]::WriteAllText($MarkerFile, $Hoy,
        (New-Object System.Text.UTF8Encoding($false)))
}

exit $code
