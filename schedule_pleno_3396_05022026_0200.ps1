$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$taskName = 'PautaEletronica_Pleno_3396_20260205_0200'
$runScript = Join-Path $scriptDir 'run_pleno_3396_task.ps1'

if (-not (Test-Path $runScript)) {
    throw "Script de execucao nao encontrado: $runScript"
}

$pwshExe = (Get-Command powershell -ErrorAction Stop).Source
$at = [datetime]'2026-02-05T02:00:00'
$userId = "$env:USERDOMAIN\$env:USERNAME"

$action = New-ScheduledTaskAction -Execute $pwshExe -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$runScript`""
$trigger = New-ScheduledTaskTrigger -Once -At $at
$principal = New-ScheduledTaskPrincipal -UserId $userId -LogonType Interactive -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description 'Gera a pauta 3396, rebaixa planilhas e envia por e-mail automaticamente.' `
    -Force | Out-Null

$info = Get-ScheduledTaskInfo -TaskName $taskName
Write-Host "Tarefa agendada: $taskName"
Write-Host "Proxima execucao: $($info.NextRunTime)"
Write-Host "Ultimo resultado: $($info.LastTaskResult)"
