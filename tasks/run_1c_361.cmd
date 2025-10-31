@echo off
pushd "%~dp0.."
set "TO=
sonia.santos@tcmsp.tc.br; ramon.ramos@tcmsp.tc.br; raphael.goncalves@tcmsp.tc.br; yelmo.junior@tcmsp.tc.br; vera.tucunduva@tcmsp.tc.br
"
".\.venv\Scripts\python.exe" ".\app.py" --headless true --sessao 361 --de 01/04/2025 --ate 30/04/2025 --download-dir ".\planilhas_1C_361_2025" --output-dir ".\output" --meta-tipo ordinaria --meta-formato presencial --meta-competencia 1c --meta-numero 361 --meta-data-abertura 30/04/2025 --meta-horario "9h30min." --send-email --email-to "%TO%" --email-subject "Pauta 1a Camara 361 - Automatica"
popd
