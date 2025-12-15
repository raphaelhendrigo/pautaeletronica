@echo off
pushd "%~dp0.."
set "TO=
sonia.santos@tcmsp.tc.br; ramon.ramos@tcmsp.tc.br; raphael.goncalves@tcmsp.tc.br; yelmo.junior@tcmsp.tc.br; vera.tucunduva@tcmsp.tc.br
"
del /q /s ".\planilhas_PLENO_3389_2025*" 2>nul
"..venv\Scripts\python.exe" ".\app.py" --headless true --sessao 3389 --de 27/10/2025 --ate 26/11/2025 --download-dir ".\planilhas_PLENO_3389_2025" --output-dir ".\output" --nome-docx "PAUTA_PLENO_3389.docx" --meta-tipo ordinaria --meta-formato presencial --meta-competencia pleno --meta-numero 3.389 --meta-data-abertura 12/11/2025 --send-email --email-to "%TO%" --email-subject "Pauta Pleno 3.389 - Automatica" --email-account "pautaeletronica@tcmsp.tc.br" --email-verbose --email-force-sync
popd
