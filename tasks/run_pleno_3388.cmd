@echo off
pushd "%~dp0.."
set "TO=
sonia.santos@tcmsp.tc.br; ramon.ramos@tcmsp.tc.br; raphael.goncalves@tcmsp.tc.br; yelmo.junior@tcmsp.tc.br; vera.tucunduva@tcmsp.tc.br
"
".\.venv\Scripts\python.exe" ".\app.py" --headless true --sessao 3388 --de 01/10/2025 --ate 31/10/2025 --download-dir ".\planilhas_PLENO_3388_2025" --output-dir ".\output" --meta-tipo ordinaria --meta-formato presencial --meta-competencia pleno --meta-numero "3.388" --meta-data-abertura 23/10/2025 --meta-horario "9h30min." --send-email --email-to "%TO%" --email-subject "Pauta Pleno 3388 - Automatica"
popd
