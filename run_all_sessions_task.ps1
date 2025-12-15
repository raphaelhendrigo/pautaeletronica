$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

$sessionScripts = @(
    @{ Name = '1a Camara 361'; File = 'run_1c_361_task.ps1' },
    @{ Name = '2a Camara 366'; File = 'run_2c_366_task.ps1' },
    @{ Name = 'Pleno 3391'; File = 'run_pleno_3391_task.ps1' },
    @{ Name = 'Pleno 3392'; File = 'run_pleno_3392_task.ps1' },
    @{ Name = 'SONP 73'; File = 'run_sonp_73_task.ps1' }
)

$pwshExe = (Get-Command pwsh -ErrorAction SilentlyContinue)?.Source
if (-not $pwshExe) {
    $pwshExe = (Get-Command powershell -ErrorAction Stop).Source
}

foreach ($session in $sessionScripts) {
    $scriptPath = Join-Path $scriptDir $session.File
    if (-not (Test-Path $scriptPath)) {
        throw "Script não encontrado: $($session.File)"
    }

    Write-Host "===== Iniciando $($session.Name) =====" -ForegroundColor Cyan
    $arguments = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', "`"$scriptPath`""
    )
    $proc = Start-Process -FilePath $pwshExe -ArgumentList $arguments -WorkingDirectory $scriptDir -Wait -PassThru
    if ($proc.ExitCode -ne 0) {
        throw "Falha ao executar $($session.File) (código $($proc.ExitCode))."
    }
    Write-Host "===== Finalizado $($session.Name) =====" -ForegroundColor Green
}

Write-Host "Todas as sessões foram geradas com sucesso." -ForegroundColor Green
