@echo off
pushd "%~dp0.."
set "TO=
sonia.santos@tcmsp.tc.br; ramon.ramos@tcmsp.tc.br; raphael.goncalves@tcmsp.tc.br; yelmo.junior@tcmsp.tc.br; vera.tucunduva@tcmsp.tc.br
"
".\.venv\Scripts\python.exe" ".\app.py" --headless true --sessao 72 --de 01/10/2025 --ate 31/10/2025 --download-dir ".\planilhas_SONP_72_2025" --output-dir ".\output" --meta-tipo ordinaria --meta-formato nao-presencial --meta-competencia pleno --meta-numero 72 --meta-data-abertura 23/10/2025 --send-email --email-to "%TO%" --email-subject "Pauta SONP 72 - Automatica"
popd
