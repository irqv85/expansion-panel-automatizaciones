<#
.SYNOPSIS
    Registra en el Programador de tareas de Windows las rutinas diarias que
    en Ubuntu corrian por cron.

.DESCRIPTION
    Crea (o actualiza) dos tareas que llaman a Run-Rutina.ps1. Se registran
    para el usuario actual, asi que NO requieren permisos de administrador.

    Equivalencia con el cron anterior:
      - Disparador diario a las 8:00, repitiendo cada 30 minutos durante 12
        horas. La repeticion cubre el caso de que la PC este apagada a la
        hora objetivo, igual que el cron cada 30 min; el marcador
        .last-success-<skill> evita que se repita el trabajo el mismo dia.
      - StartWhenAvailable: si la PC estaba apagada a las 8:00, Windows
        dispara la tarea en cuanto arranca, sin esperar el siguiente ciclo.
      - Solo dias habiles: lo decide Run-Rutina.ps1 con -SoloDiasHabiles,
        no el disparador, para conservar la misma logica que tenia el .sh.

    Para revisar despues:  Get-ScheduledTask -TaskPath '\GB Claude\'
    Para correr una a mano: Start-ScheduledTask -TaskPath '\GB Claude\' -TaskName '<nombre>'
    Para quitarlas:        Unregister-ScheduledTask -TaskPath '\GB Claude\' -TaskName '<nombre>'

.EXAMPLE
    .\Registrar-Tareas.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$RutinasDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Runner     = Join-Path $RutinasDir 'Run-Rutina.ps1'
$TaskPath   = '\GB Claude\'

if (-not (Test-Path $Runner)) {
    throw "No se encontro $Runner"
}

$rutinas = @(
    @{ Nombre = 'GB - Presentaciones diarias (Sofia)'; Skill = 'daily-client-presentation-prep'; Hora = '08:00' },
    @{ Nombre = 'GB - Reasignacion orgs (Sofia)';      Skill = 'deteccion-reasignacion-orgs-sofia'; Hora = '08:30' }
)

foreach ($r in $rutinas) {
    $accion = New-ScheduledTaskAction `
        -Execute 'powershell.exe' `
        -Argument ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -Skill {1} -SoloDiasHabiles' -f $Runner, $r.Skill) `
        -WorkingDirectory $RutinasDir

    $disparador = New-ScheduledTaskTrigger -Daily -At $r.Hora
    $disparador.Repetition = (New-ScheduledTaskTrigger -Once -At $r.Hora `
        -RepetitionInterval (New-TimeSpan -Minutes 30) `
        -RepetitionDuration (New-TimeSpan -Hours 12)).Repetition

    $ajustes = New-ScheduledTaskSettingsSet `
        -StartWhenAvailable `
        -DontStopIfGoingOnBatteries `
        -AllowStartIfOnBatteries `
        -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Hours 1)

    $principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive

    Register-ScheduledTask `
        -TaskName $r.Nombre `
        -TaskPath $TaskPath `
        -Action $accion `
        -Trigger $disparador `
        -Settings $ajustes `
        -Principal $principal `
        -Description ("Corre la skill {0} con claude -p. Reemplaza al cron de Ubuntu. Log en Rutinas\logs\{0}.log" -f $r.Skill) `
        -Force | Out-Null

    Write-Host ("Registrada: {0}  ->  skill {1}, diaria {2} (repite cada 30 min, 12 h)" -f $r.Nombre, $r.Skill, $r.Hora)
}

Write-Host ''
Write-Host 'Listo. Revisa con: Get-ScheduledTask -TaskPath ''\GB Claude\'''
Write-Host 'Nota: estas tareas corren Claude Code con --dangerously-skip-permissions,'
Write-Host 'igual que el cron anterior. Solo se disparan si la sesion del usuario existe.'
