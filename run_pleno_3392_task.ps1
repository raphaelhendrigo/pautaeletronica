$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

$to = 'raphael.goncalves@tcmsp.tc.br; sonia.santos@tcmsp.tc.br; ramon.ramos@tcmsp.tc.br; yelmo.junior@tcmsp.tc.br; vera.tucunduva@tcmsp.tc.br; glaucia.calvet@tcmsp.tc.br'
$body = 'Envio automatico da pauta Pleno 3392. Se nao receber, por favor avise.'
$env:TCM_ASSINATURA_NOME = 'ROSELI DE MORAIS CHAVES'
$env:TCM_ASSINATURA_CARGO = 'SUBSECRETÁRIA-GERAL'
$env:TCM_ASSINATURA_DATA = '04/12/2025'
$env:TCM_META_ABERTURA_FINAL = '10/12/2025'
$env:TCM_META_ENCERRAMENTO_FINAL = '04/12/2025'

$downloadDir = Join-Path $scriptDir 'planilhas_PLENO_3392_2025'
if (Test-Path $downloadDir) {
    Remove-Item -Recurse -Force $downloadDir
}

$outputDoc = Join-Path $scriptDir 'output\PAUTA_PLENO_3392_2025.docx'
if (Test-Path $outputDoc) {
    Remove-Item -Force $outputDoc
}

$python = Join-Path $scriptDir '.venv\Scripts\python.exe'
$args = @(
    '.\app.py',
    '--headless', 'true',
    '--sessao', '3392',
    '--de', '01/12/2025',
    '--ate', '31/12/2025',
    '--download-dir', '.\planilhas_PLENO_3392_2025',
    '--output-dir', '.\output',
    '--nome-docx', 'PAUTA_PLENO_3392_2025.docx',
    '--meta-tipo', 'ordinaria',
    '--meta-formato', 'presencial',
    '--meta-competencia', 'pleno',
    '--meta-numero', '3.392',
    '--meta-data-abertura', '10/12/2025',
    '--meta-horario', '9h30min.',
    '--send-email',
    '--email-to', $to,
    '--email-subject', 'Pauta Pleno 3392 - Automatica',
    '--email-body', $body,
    '--email-force-sync',
    '--email-verbose'
)

& $python @args
exit $LASTEXITCODE
