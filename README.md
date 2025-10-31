# Automação e-TCM (Sessão 71/2025)

Fluxo:
1) Login no e-TCM
2) Filtro da Sessão 71/2025 (29/09/2025 a 29/10/2025)
3) Download das planilhas por conselheiro
4) Consolidação em `output/PAUTA_UNIFICADA_71_2025.docx`

## Requisitos

- Python 3.10+
- Chromium do Playwright

## Instalação

```powershell
py -m venv .venv; .\.venv\Scripts\activate; `
pip install -U pip -r requirements.txt; `
python -m playwright install chromium
