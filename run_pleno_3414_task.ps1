$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

$to = 'raphael.goncalves@tcmsp.tc.br; ramon.ramos@tcmsp.tc.br; sonia.santos@tcmsp.tc.br'
$body = 'Segue a pauta da Sessao Ordinaria 3.414, de 24/06/2026. Esta sessao nao possui processos a relatar (pauta sem itens). Esta e a versao correta - favor desconsiderar o e-mail anterior de mesmo assunto.'
$env:TCM_ASSINATURA_DATA = '19/06/2026'
$env:TCM_META_ABERTURA_FINAL = '24/06/2026'

$downloadDir = Join-Path $scriptDir 'planilhas_PLENO_3414_2026'
if (Test-Path $downloadDir) {
    Remove-Item -Recurse -Force $downloadDir
}

$outputDoc = Join-Path $scriptDir 'output\PAUTA_PLENO_3414_2026.docx'
if (Test-Path $outputDoc) {
    Remove-Item -Force $outputDoc
}

$python = Join-Path $scriptDir '.venv\Scripts\python.exe'
$args = @(
    '.\app.py',
    '--headless', 'true',
    '--sessao', '3414',
    '--de', '01/06/2026',
    '--ate', '30/06/2026',
    '--download-dir', '.\planilhas_PLENO_3414_2026',
    '--output-dir', '.\output',
    '--nome-docx', 'PAUTA_PLENO_3414_2026.docx',
    '--meta-tipo', 'ordinaria',
    '--meta-formato', 'presencial',
    '--meta-competencia', 'pleno',
    '--meta-numero', '3.414',
    '--meta-data-abertura', '24/06/2026',
    '--meta-horario', '9h30min.',
    '--send-email',
    '--email-to', $to,
    '--email-subject', 'Pauta Pleno 3414 - Automatica',
    '--email-body', $body,
    '--email-force-sync',
    '--email-verbose'
)

& $python @args
exit $LASTEXITCODE
