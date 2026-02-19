# Automacao e-TCM (SONP)

Fluxo:
1) Login no e-TCM
2) Filtro da sessao e periodo
3) Download das planilhas por conselheiro
4) Consolidacao em DOCX (pauta unificada)

## Requisitos

- Python 3.10+
- Chromium do Playwright

## Instalacao

```powershell
py -m venv .venv; .\.venv\Scripts\activate; `
pip install -U pip -r requirements.txt; `
python -m playwright install chromium
```

## Configuracao

Copie `.env.example` para `.env` e preencha os valores (nao versione `.env`).

Variaveis obrigatorias:
- e-TCM: `ETCM_USERNAME`, `ETCM_PASSWORD`

Variaveis opcionais:
- e-TCM: `ETCM_BASE_URL`
- Pipeline: `SESSAO`, `DATA_DE`, `DATA_ATE`, `HEADLESS`, `HEADER_TEMPLATE`, `DOWNLOAD_DIR`, `OUTPUT_DIR`
- Outlook (app.py): `TCM_EMAIL_ACCOUNT`, `TCM_EMAIL_TO`, `TCM_EMAIL_CC`, `TCM_EMAIL_BCC`, `TCM_EMAIL_SUBJECT`, `TCM_EMAIL_BODY`
- SMTP (server.py): `SMTP_HOST`, `SMTP_PORT`, `SMTP_USERNAME`, `SMTP_PASSWORD`, `SMTP_FROM`, `SMTP_TO`, `EMAIL_CC`, `EMAIL_BCC`, `EMAIL_SUBJECT`, `EMAIL_BODY`

## Execucao local (gera DOCX)

```powershell
.\.venv\Scripts\python.exe .\app.py `
  --headless true `
  --sessao 74 `
  --de 01/01/2026 `
  --ate 31/12/2026 `
  --download-dir .\planilhas_74_2026 `
  --output-dir .\output `
  --meta-tipo ordinaria `
  --meta-formato nao-presencial `
  --meta-competencia pleno `
  --meta-numero 74 `
  --meta-data-abertura 03/02/2026
```

## Execucao local (gera Excel consolidado - Consulta de Pautas)

```powershell
.\.venv\Scripts\python.exe .\app.py `
  --modo consulta-pautas `
  --headless true `
  --de 01/01/2026 `
  --ate 31/01/2026 `
  --download-dir .\planilhas_pautas_2026 `
  --output-dir .\output `
  --competencias pleno,1c,2c `
  --situacao Aberta `
  --consolidado-nome pautas_CONSOLIDADO_01_2026.xlsx
```

## Envio via Outlook

```powershell
.\.venv\Scripts\python.exe .\app.py `
  --send-email `
  --email-to "destinatario1@exemplo; destinatario2@exemplo" `
  --email-account "sua_conta_outlook"
```

## Execucao via API (SMTP)

```powershell
.\.venv\Scripts\python.exe .\server.py
```

## Testes (smoke)

```powershell
.\.venv\Scripts\python.exe .\scripts\smoke_test.py
```
