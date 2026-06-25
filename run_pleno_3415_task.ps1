$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

$to = 'raphael.goncalves@tcmsp.tc.br; ramon.ramos@tcmsp.tc.br; sonia.santos@tcmsp.tc.br'
$body = 'Envio automatico da pauta Pleno 3.415 (Sessao Extraordinaria). Se nao receber, por favor avise.'
$env:TCM_ASSINATURA_DATA = '19/06/2026'
$env:TCM_META_ABERTURA_FINAL = '24/06/2026'

$downloadDir = Join-Path $scriptDir 'planilhas_PLENO_3415_2026'
if (Test-Path $downloadDir) {
    Remove-Item -Recurse -Force $downloadDir
}

$outputDoc = Join-Path $scriptDir 'output\PAUTA_PLENO_3415_2026.docx'
if (Test-Path $outputDoc) {
    Remove-Item -Force $outputDoc
}

$python = Join-Path $scriptDir '.venv\Scripts\python.exe'
$args = @(
    '.\app.py',
    '--headless', 'true',
    '--sessao', '3415',
    '--de', '01/06/2026',
    '--ate', '30/06/2026',
    '--download-dir', '.\planilhas_PLENO_3415_2026',
    '--output-dir', '.\output',
    '--nome-docx', 'PAUTA_PLENO_3415_2026.docx',
    '--meta-tipo', 'extraordinaria',
    '--meta-formato', 'presencial',
    '--meta-competencia', 'pleno',
    '--meta-numero', '3.415',
    '--meta-data-abertura', '24/06/2026',
    '--meta-horario', '9h30min.',
    '--send-email',
    '--email-to', $to,
    '--email-subject', 'Pauta Pleno 3415 - Automatica',
    '--email-body', $body,
    '--email-force-sync',
    '--email-verbose'
)

& $python @args
exit $LASTEXITCODE
