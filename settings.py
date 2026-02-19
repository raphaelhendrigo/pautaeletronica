from __future__ import annotations

from dataclasses import dataclass
import os
from typing import Iterable

from dotenv import load_dotenv


class ConfigError(RuntimeError):
    pass


OUTPUT_DIR_PLANILHAS = r"C:\Users\20386\pautaeletronica\planilhas_74_2026"


def load_env() -> None:
    load_dotenv()


def env(name: str, default: str | None = None) -> str | None:
    val = os.getenv(name)
    if val is None or val == "":
        return default
    return val


def require_env(name: str) -> str:
    val = env(name)
    if not val:
        raise ConfigError(f"Missing required environment variable: {name}")
    return val


def require_any(names: Iterable[str]) -> str:
    names_list = list(names)
    for name in names_list:
        val = env(name)
        if val:
            return val
    joined = ", ".join(names_list)
    raise ConfigError(f"Missing required environment variable (one of): {joined}")


@dataclass
class EtcmConfig:
    base_url: str
    username: str
    password: str


def get_etcm_config() -> EtcmConfig:
    base_url = env("ETCM_BASE_URL", env("BASE_URL", "https://etcm.tcm.sp.gov.br"))
    username = require_any(["ETCM_USERNAME", "ETCM_USER"])
    password = require_any(["ETCM_PASSWORD", "ETCM_PASS"])
    return EtcmConfig(base_url=base_url or "https://etcm.tcm.sp.gov.br", username=username, password=password)


@dataclass
class SmtpConfig:
    host: str
    port: int
    username: str
    password: str
    sender: str
    to: str
    cc: str
    bcc: str
    subject: str
    body: str


def get_smtp_config(*, default_subject: str, default_body: str) -> SmtpConfig:
    host = env("SMTP_HOST", "smtp.office365.com") or "smtp.office365.com"
    port = int(env("SMTP_PORT", "587") or "587")
    username = require_any(["SMTP_USERNAME", "SMTP_USER", "EMAIL_USER"])
    password = require_any(["SMTP_PASSWORD", "SMTP_PASS"])
    sender = env("SMTP_FROM", env("EMAIL_SENDER", username))
    if not sender:
        raise ConfigError("Missing required environment variable: SMTP_FROM (or EMAIL_SENDER)")
    to = env("SMTP_TO", env("EMAIL_TO", sender)) or sender
    cc = env("EMAIL_CC", "") or ""
    bcc = env("EMAIL_BCC", "") or ""
    subject = env("EMAIL_SUBJECT", default_subject) or default_subject
    body = env("EMAIL_BODY", default_body) or default_body
    return SmtpConfig(
        host=host,
        port=port,
        username=username,
        password=password,
        sender=sender,
        to=to,
        cc=cc,
        bcc=bcc,
        subject=subject,
        body=body,
    )
