$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Resolve-Path (Join-Path $scriptDir '..')
Set-Location $projectDir

$downloadDir = Join-Path $projectDir 'planilhas_74_2026'
if (Test-Path $downloadDir) {
    Remove-Item -Recurse -Force $downloadDir
}

$outputDoc = Join-Path $projectDir 'output\PAUTA_UNIFICADA_74_2026.docx'
if (Test-Path $outputDoc) {
    Remove-Item -Force $outputDoc
}

$env:TCM_META_ABERTURA_FINAL = '03/02/2026'
$env:TCM_META_ENCERRAMENTO_FINAL = '19/02/2026'
$env:TCM_ASSINATURA_DATA = '22/01/2026'

$to = 'ramon.ramos@tcmsp.tc.br; glaucia.calvet@tcmsp.tc.br; yelmo.junior@tcmsp.tc.br; raphael.goncalves@tcmsp.tc.br; vera.tucunduva@tcmsp.tc.br; sonia.santos@tcmsp.tc.br'

$python = Join-Path $projectDir '.venv\Scripts\python.exe'
$args = @(
    '.\app.py',
    '--headless', 'true',
    '--sessao', '74',
    '--de', '01/01/2026',
    '--ate', '31/12/2026',
    '--download-dir', '.\planilhas_74_2026',
    '--output-dir', '.\output',
    '--meta-tipo', 'ordinaria',
    '--meta-formato', 'nao-presencial',
    '--meta-competencia', 'pleno',
    '--meta-numero', '74',
    '--meta-data-abertura', '03/02/2026',
    '--send-email',
    '--email-to', $to,
    '--email-subject', 'Pauta SONP 74 - Automatica',
    '--email-account', 'pautaeletronica@tcmsp.tc.br',
    '--email-force-sync',
    '--email-verbose'
)

& $python @args
exit $LASTEXITCODE
