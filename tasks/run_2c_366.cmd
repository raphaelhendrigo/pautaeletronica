@echo off
pushd "%~dp0.."
set "TO=
sonia.santos@tcmsp.tc.br; ramon.ramos@tcmsp.tc.br; raphael.goncalves@tcmsp.tc.br; yelmo.junior@tcmsp.tc.br; vera.tucunduva@tcmsp.tc.br
"
".\.venv\Scripts\python.exe" ".\app.py" --headless true --sessao 366 --de 01/09/2025 --ate 30/09/2025 --download-dir ".\planilhas_2C_366_2025" --output-dir ".\output" --meta-tipo ordinaria --meta-formato presencial --meta-competencia 2c --meta-numero 366 --meta-data-abertura 24/09/2025 --meta-horario "9h30min." --send-email --email-to "%TO%" --email-subject "Pauta 2a Camara 366 - Automatica"
popd
