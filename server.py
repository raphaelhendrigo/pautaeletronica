# server.py
from __future__ import annotations

import os
from pathlib import Path
from flask import Flask, jsonify, request
from dotenv import load_dotenv

from main import run_pipeline
from email_smtp import send_email_smtp

app = Flask(__name__)
load_dotenv()

def env(name: str, default: str | None = None) -> str | None:
    val = os.getenv(name)
    return val if val is not None and val != "" else default

def _bool_env(name: str, default: bool = False) -> bool:
    v = (env(name) or "").strip().lower()
    if v in {"1", "true", "t", "yes", "y", "on"}:
        return True
    if v in {"0", "false", "f", "no", "n", "off"}:
        return False
    return default

def _run_once() -> dict:
    # 1) parâmetros do pipeline
    base_url = env("BASE_URL", "https://etcm.tcm.sp.gov.br")
    etcm_user = env("ETCM_USER")
    etcm_pass = env("ETCM_PASS")
    num_sessao = env("SESSAO", "71")
    data_de = env("DATA_DE", "29/09/2025")
    data_ate = env("DATA_ATE", "29/10/2025")
    download_dir = env("DOWNLOAD_DIR", "/workspace/downloads")
    output_dir = env("OUTPUT_DIR", "/workspace/output")
    titulo_docx = env("TITULO_DOCX", f"Pauta Unificada - Sessão {num_sessao}/2025")
    nome_docx = env("NOME_DOCX", f"PAUTA_UNIFICADA_{num_sessao}_2025.docx")
    header_template = env("HEADER_TEMPLATE", "papel_timbrado_tcm.docx")
    headless = _bool_env("HEADLESS", True)

    if not etcm_user or not etcm_pass:
        raise RuntimeError("Defina ETCM_USER e ETCM_PASS nas variáveis de ambiente.")

    # 2) executa pipeline (gera DOCX)
    out_path = run_pipeline(
        base_url=base_url,
        usuario=etcm_user,
        senha=etcm_pass,
        num_sessao=num_sessao,
        data_de=data_de,
        data_ate=data_ate,
        download_dir=download_dir,
        output_dir=output_dir,
        headless=headless,
        titulo_docx=titulo_docx,
        header_template=header_template,
        nome_docx=nome_docx,
    )

    # 3) email via SMTP (para Cloud Run)
    smtp_host = env("SMTP_HOST", "smtp.office365.com")
    smtp_port = int(env("SMTP_PORT", "587"))
    smtp_user = env("SMTP_USER") or env("SMTP_USERNAME") or env("EMAIL_USER") or ""
    smtp_pass = env("SMTP_PASS") or env("SMTP_PASSWORD") or ""
    sender = env("EMAIL_SENDER") or smtp_user
    to = env("EMAIL_TO", sender)
    cc = env("EMAIL_CC", "")
    bcc = env("EMAIL_BCC", "")
    subject = env("EMAIL_SUBJECT", f"TESTE – Pauta SONP {num_sessao} gerada automaticamente (GCP)")
    body = env("EMAIL_BODY", f"<p>Segue em anexo a pauta gerada automaticamente para <b>SONP {num_sessao}</b>.<br/><i>Este é um envio de <b>teste</b> pelo GCP.</i></p>")

    if not smtp_user or not smtp_pass or not sender:
        raise RuntimeError("SMTP_USER/SMTP_PASS/EMAIL_SENDER não configurados para envio SMTP.")

    send_email_smtp(
        smtp_host=smtp_host,
        smtp_port=smtp_port,
        smtp_user=smtp_user,
        smtp_pass=smtp_pass,
        sender=sender,
        to=to,
        cc=cc,
        bcc=bcc,
        subject=subject,
        html_body=body,
        attachment=Path(out_path),
        sessao=num_sessao,
    )
    return {"ok": True, "docx": out_path}

@app.get("/healthz")
def healthz():
    return jsonify({"ok": True})

@app.post("/run")
def run_now():
    try:
        result = _run_once()
        return jsonify(result), 200
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.getenv("PORT", "8080"))
    app.run(host="0.0.0.0", port=port)
