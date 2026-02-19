# local_agent.py
from __future__ import annotations
import os, subprocess
from pathlib import Path
from flask import Flask, request, jsonify
from settings import env, load_env

PROJECT_DIR = Path(__file__).parent.resolve()
VENV_PY = PROJECT_DIR / ".venv" / "Scripts" / "python.exe"
APP_PY  = PROJECT_DIR / "app.py"

app = Flask(__name__)
load_env()

def _defaults() -> dict:
    return {
        "headless": True,
        "sessao": env("SESSAO", "71"),
        "de": env("DATA_DE", "29/09/2025"),
        "ate": env("DATA_ATE", "29/10/2025"),
        "download_dir": str(PROJECT_DIR / "downloads"),
        "output_dir": str(PROJECT_DIR / "output"),
        "header_template": str(PROJECT_DIR / "papel_timbrado_tcm.docx"),
        "nome_docx": "SONP_71_2025.docx",
        # email (usa a própria funcionalidade do app.py)
        "send_email": True,
        "email_account": env("TCM_EMAIL_ACCOUNT", ""),
        "email_to": env("TCM_EMAIL_TO", ""),
        "email_subject": os.getenv("EMAIL_SUBJECT", "TESTE – Pauta SONP gerada automaticamente"),
        "email_body": os.getenv("EMAIL_BODY", "<p>Envio automático <b>teste</b>.</p>"),
        "email_verbose": True,
        "email_force_sync": False,
    }

@app.get("/healthz")
def healthz():
    return jsonify({"ok": True})

@app.post("/run")
def run_now():
    data = request.get_json(silent=True) or {}
    args = _defaults()
    args.update(data)

    cmd = [
        str(VENV_PY), str(APP_PY),
        "--headless", str(args["headless"]).lower(),
        "--sessao", str(args["sessao"]),
        "--de", args["de"],
        "--ate", args["ate"],
        "--download-dir", args["download_dir"],
        "--output-dir", args["output_dir"],
        "--header-template", args["header_template"],
        "--nome-docx", args["nome_docx"],
    ]
    if args.get("send_email"): cmd.append("--send-email")
    if args.get("email_account"): cmd += ["--email-account", args["email_account"]]
    if args.get("email_to"): cmd += ["--email-to", args["email_to"]]
    if args.get("email_subject"): cmd += ["--email-subject", args["email_subject"]]
    if args.get("email_body"): cmd += ["--email-body", args["email_body"]]
    if args.get("email_verbose"): cmd.append("--email-verbose")
    if args.get("email_force_sync"): cmd.append("--email-force-sync")

    env = os.environ.copy()
    env.setdefault("ETCM_USERNAME", os.getenv("ETCM_USERNAME", os.getenv("ETCM_USER", "")))
    env.setdefault("ETCM_PASSWORD", os.getenv("ETCM_PASSWORD", os.getenv("ETCM_PASS", "")))

    try:
        p = subprocess.run(
            cmd, cwd=str(PROJECT_DIR),
            capture_output=True, text=True, shell=False, timeout=3600, env=env
        )
        return jsonify({
            "ok": p.returncode == 0,
            "code": p.returncode,
            "cmd": cmd,
            "stdout": p.stdout[-4000:],
            "stderr": p.stderr[-4000:]
        }), (200 if p.returncode == 0 else 500)
    except subprocess.TimeoutExpired as e:
        return jsonify({"ok": False, "error": "timeout"}), 504

if __name__ == "__main__":
    # roda em 127.0.0.1:5000
    app.run(host="127.0.0.1", port=5000)
