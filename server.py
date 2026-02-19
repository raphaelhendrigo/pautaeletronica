# server.py
from __future__ import annotations

import os
from pathlib import Path
from flask import Flask, jsonify, request
from main import run_pipeline
from email_smtp import send_email_smtp
from settings import ConfigError, env, get_etcm_config, get_smtp_config, load_env

app = Flask(__name__)
load_env()

def _bool_env(name: str, default: bool = False) -> bool:
    v = (env(name) or "").strip().lower()
    if v in {"1", "true", "t", "yes", "y", "on"}:
        return True
    if v in {"0", "false", "f", "no", "n", "off"}:
        return False
    return default

def _run_once() -> dict:
    # 1) parametros do pipeline
    try:
        etcm = get_etcm_config()
    except ConfigError as e:
        raise RuntimeError(str(e))
    base_url = env("ETCM_BASE_URL", env("BASE_URL", etcm.base_url))
    num_sessao = env("SESSAO", "71")
    data_de = env("DATA_DE", "29/09/2025")
    data_ate = env("DATA_ATE", "29/10/2025")
    download_dir = env("DOWNLOAD_DIR", "/workspace/downloads")
    output_dir = env("OUTPUT_DIR", "/workspace/output")
    titulo_docx = env("TITULO_DOCX", f"Pauta Unificada - Sessão {num_sessao}/2025")
    nome_docx = env("NOME_DOCX", f"PAUTA_UNIFICADA_{num_sessao}_2025.docx")
    header_template = env("HEADER_TEMPLATE", "papel_timbrado_tcm.docx")
    headless = _bool_env("HEADLESS", True)
    # 2) executa pipeline (gera DOCX)
    out_path = run_pipeline(
        base_url=base_url,
        usuario=etcm.username,
        senha=etcm.password,
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
    default_subject = f"TESTE - Pauta SONP {num_sessao} gerada automaticamente (GCP)"
    default_body = (
        f"<p>Segue em anexo a pauta gerada automaticamente para <b>SONP {num_sessao}</b>."
        f"<br/><i>Este e um envio de <b>teste</b> pelo GCP.</i></p>"
    )
    try:
        smtp = get_smtp_config(default_subject=default_subject, default_body=default_body)
    except ConfigError as e:
        raise RuntimeError(str(e))

    send_email_smtp(

        smtp_host=smtp.host,
        smtp_port=smtp.port,
        smtp_user=smtp.username,
        smtp_pass=smtp.password,
        sender=smtp.sender,
        to=smtp.to,
        cc=smtp.cc,
        bcc=smtp.bcc,
        subject=smtp.subject,
        html_body=smtp.body,
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
