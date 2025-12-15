$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

$to = 'raphael.goncalves@tcmsp.tc.br; ramon.ramos@tcmsp.tc.br; sonia.santos@tcmsp.tc.br; yelmo.junior@tcmsp.tc.br; vera.tucunduva@tcmsp.tc.br; glaucia.calvet@tcmsp.tc.br'
$env:TCM_ASSINATURA_DATA = '18/11/2025'
$env:TCM_META_ABERTURA_FINAL = '26/11/2025'

$downloadDir = Join-Path $scriptDir 'planilhas_2C_366_2025'
if (Test-Path $downloadDir) {
    Remove-Item -Recurse -Force $downloadDir
}

$python = Join-Path $scriptDir '.venv\Scripts\python.exe'
$args = @(
    '.\app.py',
    '--headless', 'true',
    '--sessao', '366',
    '--de', '01/09/2025',
    '--ate', '30/09/2025',
    '--download-dir', '.\planilhas_2C_366_2025',
    '--output-dir', '.\output',
    '--nome-docx', 'PAUTA_2C_366_2025.docx',
    '--meta-tipo', 'ordinaria',
    '--meta-formato', 'presencial',
    '--meta-competencia', '2c',
    '--meta-numero', '366',
    '--meta-data-abertura', '26/11/2025',
    '--meta-horario', '9h30min.',
    '--send-email',
    '--email-to', $to,
    '--email-subject', 'Pauta 2a Camara 366 - Automatica',
    '--email-force-sync',
    '--email-verbose'
)

& $python @args
exit $LASTEXITCODE
