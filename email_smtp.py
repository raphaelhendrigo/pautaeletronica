# email_smtp.py
from __future__ import annotations

import mimetypes
import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path
from typing import Optional
from datetime import datetime

DEFAULT_SMTP_HOST = "smtp.office365.com"
DEFAULT_SMTP_PORT = 587

def _now() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M")

def _split_addrs(s: Optional[str]) -> list[str]:
    if not s:
        return []
    out = []
    for part in s.replace(";", ",").split(","):
        p = part.strip()
        if p:
            out.append(p)
    return out

def _default_subject(sessao: Optional[str]) -> str:
    sfx = f"SONP {sessao}" if sessao else "SONP"
    return f"TESTE – Pauta {sfx} gerada automaticamente (GCP)"

def _default_body_html(sessao: Optional[str]) -> str:
    sfx = f"SONP {sessao}" if sessao else "SONP"
    return f"""<p>Prezados(as),</p>
<p>Segue em anexo a pauta gerada automaticamente para <b>{sfx}</b>.</p>
<p><i>Este é um envio de <b>teste</b> a partir do GCP.</i></p>
<hr/>
<p><small>Gerado em {_now()} – Automação Pauta SONP.</small></p>"""

def _attach(msg: EmailMessage, path: Path):
    ctype, encoding = mimetypes.guess_type(str(path))
    if ctype is None or encoding is not None:
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)
    with open(path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype=maintype,
            subtype=subtype,
            filename=path.name,
        )

def send_email_smtp(
    smtp_host: str,
    smtp_port: int,
    smtp_user: str,
    smtp_pass: str,
    sender: str,
    to: str,
    subject: Optional[str],
    html_body: Optional[str],
    attachment: Path,
    cc: Optional[str] = None,
    bcc: Optional[str] = None,
    sessao: Optional[str] = None,
):
    if not attachment.exists():
        raise FileNotFoundError(f"Anexo não encontrado: {attachment}")

    to_list = _split_addrs(to)
    cc_list = _split_addrs(cc)
    bcc_list = _split_addrs(bcc)

    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = ", ".join(to_list) if to_list else sender
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = subject or _default_subject(sessao)

    body = html_body or _default_body_html(sessao)
    msg.set_content("Seu cliente não suporta HTML.")
    msg.add_alternative(body, subtype="html")

    _attach(msg, attachment)

    all_rcpts = to_list + cc_list + bcc_list
    if not all_rcpts:
        all_rcpts = [sender]

    context = ssl.create_default_context()
    with smtplib.SMTP(smtp_host or DEFAULT_SMTP_HOST, smtp_port or DEFAULT_SMTP_PORT) as server:
        server.starttls(context=context)
        server.login(smtp_user, smtp_pass)
        server.send_message(msg, from_addr=sender, to_addrs=all_rcpts)
