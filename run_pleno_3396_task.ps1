param(
    [string]$To = 'raphael.goncalves@tcmsp.tc.br; glaucia.calvet@tcmsp.tc.br; ramon.ramos@tcmsp.tc.br; vera.tucunduva@tcmsp.tc.br; yelmo.junior@tcmsp.tc.br; sonia.santos@tcmsp.tc.br',
    [string]$EmailAccount = 'raphael.goncalves@tcmsp.tc.br'
)

$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

$logDir = Join-Path $scriptDir 'logs'
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
$logPath = Join-Path $logDir ("run_pleno_3396_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
Start-Transcript -Path $logPath -Append | Out-Null

try {
    $to = $To

    # Mantem o cabecalho exatamente como solicitado para a 3.396.
    $env:TCM_META_ABERTURA_FINAL = '04/02/2026'

    $downloadDir = Join-Path $scriptDir 'planilhas_PLENO_3396_2026'
    if (Test-Path $downloadDir) {
        Remove-Item -Recurse -Force $downloadDir
    }
    New-Item -ItemType Directory -Path $downloadDir -Force | Out-Null

    $outputDir = Join-Path $scriptDir 'output'
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

    $outputDocName = 'PAUTA_PLENO_3396_2026.docx'
    $outputDoc = Join-Path $outputDir $outputDocName
    if (Test-Path $outputDoc) {
        try {
            Remove-Item -Force -ErrorAction Stop $outputDoc
        } catch {
            $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $outputDocName = "PAUTA_PLENO_3396_2026_REENVIO_$stamp.docx"
            $outputDoc = Join-Path $outputDir $outputDocName
            Write-Host "Aviso: nao foi possivel remover o DOCX anterior. Gerando novo arquivo: $outputDocName"
        }
    }

    $python = Join-Path $scriptDir '.venv\Scripts\python.exe'
    if (-not (Test-Path $python)) {
        throw "Python nao encontrado em: $python"
    }

    $args = @(
        '.\app.py',
        '--headless', 'true',
        '--sessao', '3396',
        '--de', '01/02/2026',
        '--ate', '03/03/2026',
        '--download-dir', '.\planilhas_PLENO_3396_2026',
        '--output-dir', '.\output',
        '--nome-docx', $outputDocName,
        '--meta-tipo', 'ordinaria',
        '--meta-formato', 'presencial',
        '--meta-competencia', 'pleno',
        '--meta-numero', '3.396',
        '--meta-data-abertura', '04/02/2026',
        '--meta-horario', '9h30min.',
        '--competencias-download', 'pleno',
        '--send-email',
        '--email-to', $to,
        '--email-subject', 'Pauta Pleno 3.396 - Automatica',
        '--email-account', $EmailAccount,
        '--email-force-sync',
        '--email-verbose'
    )

    $maxAttempts = 6
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        Write-Host "=== Sessao 3396: tentativa $attempt/$maxAttempts ==="
        & $python @args
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Execucao concluida com sucesso."
            exit 0
        }
        if ($attempt -lt $maxAttempts) {
            Write-Host "Falha na tentativa $attempt. Aguardando para nova tentativa..."
            Start-Sleep -Seconds 45
        }
    }

    throw "Falha ao executar a sessao 3396 apos $maxAttempts tentativas."
}
finally {
    Stop-Transcript | Out-Null
}
