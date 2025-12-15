$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

$to = 'raphael.goncalves@tcmsp.tc.br; ramon.ramos@tcmsp.tc.br; sonia.santos@tcmsp.tc.br; yelmo.junior@tcmsp.tc.br; vera.tucunduva@tcmsp.tc.br; glaucia.calvet@tcmsp.tc.br'
$env:TCM_ASSINATURA_DATA = '18/11/2025'
$env:TCM_META_ABERTURA_FINAL = '26/11/2025'

$downloadDir = Join-Path $scriptDir 'planilhas_SONP_73_2025'
if (Test-Path $downloadDir) {
    Remove-Item -Recurse -Force $downloadDir
}

$python = Join-Path $scriptDir '.venv\Scripts\python.exe'
$args = @(
    '.\app.py',
    '--headless', 'true',
    '--sessao', '73',
    '--de', '01/11/2025',
    '--ate', '30/11/2025',
    '--download-dir', '.\planilhas_SONP_73_2025',
    '--output-dir', '.\output',
    '--meta-tipo', 'ordinaria',
    '--meta-formato', 'nao-presencial',
    '--meta-competencia', 'pleno',
    '--meta-numero', '73',
    '--meta-data-abertura', '26/11/2025',
    '--send-email',
    '--email-to', $to,
    '--email-subject', 'Pauta SONP 73 - Automatica',
    '--email-force-sync',
    '--email-verbose'
)

& $python @args
exit $LASTEXITCODE
