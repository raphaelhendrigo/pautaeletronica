# email_outlook.py
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional, List, Tuple
from datetime import datetime
import time
import os

try:
    import win32com.client as win32
except Exception as e:
    raise RuntimeError(
        "pywin32 não instalado ou Outlook não disponível. "
        "Instale com: pip install pywin32"
    ) from e


@dataclass
class SendResult:
    status: str
    account: Optional[str]
    recipients_resolved: bool
    entry_id: Optional[str]
    outbox_before: int
    outbox_after: int
    sent_before: int
    sent_after: int
    online_before: Optional[str]
    online_after: Optional[str]
    attachment: Optional[str]


# ---------- util ----------
def _now() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M")


def _split_addrs(s: Optional[str]) -> List[str]:
    if not s:
        return []
    out = []
    for part in s.replace(";", ",").split(","):
        p = part.strip()
        if p:
            out.append(p)
    return out


def _latest_docx(output_dir: str | Path) -> Path:
    p = Path(output_dir)
    if not p.exists():
        raise FileNotFoundError(f"Pasta de saída não existe: {p}")
    files = sorted(p.glob("*.docx"), key=lambda f: f.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(f"Nenhum .docx encontrado em: {p}")
    return files[0]


def _default_subject(sessao: Optional[str]) -> str:
    sfx = f"SONP {sessao}" if sessao else "SONP"
    return f"Pauta {sfx} – gerada automaticamente"


def _default_body(sessao: Optional[str]) -> str:
    sfx = f"SONP {sessao}" if sessao else "SONP"
    return (
        f"<p>Prezados(as),</p>"
        f"<p>Segue em anexo a pauta gerada automaticamente para <b>{sfx}</b>.</p>"
        f"<hr/><p><small>Gerado em {_now()} – Automação Pauta SONP.</small></p>"
    )


# ---------- Outlook helpers ----------
def _get_app_ns():
    app = win32.gencache.EnsureDispatch("Outlook.Application")
    ns = app.GetNamespace("MAPI")
    return app, ns


def _account_hint_match(acc, hint: str) -> bool:
    if not hint:
        return False
    hint_l = hint.lower()
    name = str(getattr(acc, "DisplayName", "") or "").lower()
    smtp = str(getattr(acc, "SmtpAddress", "") or getattr(acc, "Address", "") or "").lower()
    return hint_l in name or hint_l in smtp


def _resolve_account(ns, hint: Optional[str]):
    if not hint:
        return None
    try:
        for acc in ns.Session.Accounts:
            if _account_hint_match(acc, hint):
                return acc
    except Exception:
        pass
    return None


def _folder_counts(ns) -> Tuple[int, int]:
    try:
        outbox = ns.GetDefaultFolder(4)  # Outbox
        sent = ns.GetDefaultFolder(5)    # Sent Items
        return int(outbox.Items.Count), int(sent.Items.Count)
    except Exception:
        return 0, 0


def _online_state(ns) -> str | None:
    try:
        mode = int(getattr(ns, "ExchangeConnectionMode", 0))
        return {0: "Desconhecido", 100: "Online", 200: "Offline"}.get(mode, str(mode))
    except Exception:
        return None


def _force_sync(ns, verbose=False):
    try:
        if verbose:
            print("[email][sync] SendAndReceive(False)")
        ns.SendAndReceive(False)
    except Exception as e:
        if verbose:
            print(f"[email][sync] SendAndReceive falhou: {e}")
    try:
        so = ns.SyncObjects
        for i in range(so.Count):
            if verbose:
                print(f"[email][sync] SyncObjects.Item({i+1}).Start()")
            so.Item(i + 1).Start()
    except Exception as e:
        if verbose:
            print(f"[email][sync] SyncObjects falhou: {e}")


# ---------- resolução robusta de caminho do anexo ----------
def _resolve_attachment_path(p: Path, verbose=False) -> Path:
    candidates: list[Path] = []
    if p.is_absolute():
        candidates.append(p)
    candidates.append(Path.cwd() / p)
    here = Path(__file__).resolve().parent
    candidates.append(here / p)
    try:
        candidates.append(p.resolve(strict=False))
    except Exception:
        pass

    seen = set()
    uniq: list[Path] = []
    for c in candidates:
        if str(c) not in seen:
            uniq.append(c)
            seen.add(str(c))

    if verbose:
        print("[email][path] Tentativas de resolução para anexo:")
        for c in uniq:
            print(f"  - {c}")

    for c in uniq:
        try:
            if c.exists():
                return c
        except Exception:
            pass
    return uniq[0]


def _wait_file_ready(p: Path, timeout_s=6.0, verbose=False) -> bool:
    end = time.time() + timeout_s
    last_err = None
    while time.time() < end:
        try:
            if p.exists() and p.is_file():
                with open(p, "rb") as fh:
                    fh.read(64)
                if verbose:
                    try:
                        size = os.path.getsize(p)
                        print(f"[email][path] OK '{p}' (size={size})")
                    except Exception:
                        print(f"[email][path] OK '{p}'")
                return True
        except Exception as e:
            last_err = e
        time.sleep(0.15)
    if verbose:
        print(f"[email][path] Arquivo não ficou pronto a tempo: {p} ({last_err})")
    return False


# ---------- API principal ----------
def send_pauta_unificada(
    *,
    docx_path: str | Path | None,
    output_dir: str | Path,
    sessao: Optional[str],
    to: Optional[str],
    cc: Optional[str],
    bcc: Optional[str],
    subject: Optional[str],
    body: Optional[str],
    preview: bool = False,
    save_to_drafts: bool = False,
    account_hint: Optional[str] = None,
    verbose: bool = False,
    force_sync: bool = False,
) -> SendResult:
    """
    Cria o e-mail no Outlook com o DOCX anexado e:
      - preview=True  -> abre janela (você clica Enviar)
      - save_to_drafts=True -> salva em Rascunhos
      - padrão -> envia imediatamente via Outlook
    """
    app, ns = _get_app_ns()

    attach_in = Path(docx_path) if docx_path else _latest_docx(output_dir)
    attach_abs = _resolve_attachment_path(attach_in, verbose=verbose)

    if verbose:
        print(f"[email] CWD: {Path.cwd()}")
        print(f"[email] __file__ dir: {Path(__file__).resolve().parent}")
        print(f"[email] Anexo (solicitado): {attach_in}")
        print(f"[email] Anexo (resolvido):  {attach_abs}")

    if not _wait_file_ready(attach_abs, timeout_s=6.0, verbose=verbose):
        raise FileNotFoundError(
            f"Não foi possível preparar o anexo para envio.\n"
            f"Solicitado: {attach_in}\nResolvido:  {attach_abs}\n"
            f"Dica: verifique se o arquivo existe e está sincronizado (OneDrive)."
        )

    # Contadores antes
    outbox_before, sent_before = _folder_counts(ns)
    online_before = _online_state(ns)

    # Resolve destinatários
    to_list = _split_addrs(to or "")
    cc_list = _split_addrs(cc or "")
    bcc_list = _split_addrs(bcc or "")

    m = app.CreateItem(0)  # olMailItem

    # Conta a usar (e capturar nome ANTES do Send)
    account_display = None
    if account_hint:
        acc = _resolve_account(ns, account_hint)
        if acc is not None:
            try:
                m.SendUsingAccount = acc
                if verbose:
                    try:
                        print(f"[email] SendUsingAccount set to: {getattr(acc,'DisplayName',None) or getattr(acc,'SmtpAddress',None)}")
                    except Exception:
                        pass
            except Exception as e:
                if verbose:
                    print(f"[email] SendUsingAccount failed: {e}; trying COM invoke")
                try:
                    m._oleobj_.Invoke(64209, 0, 8, 0, acc)  # fallback
                    if verbose:
                        print("[email] COM invoke to set SendUsingAccount succeeded")
                except Exception as e2:
                    if verbose:
                        print(f"[email] COM invoke failed: {e2}")
            try:
                account_display = str(getattr(acc, "DisplayName", None) or getattr(acc, "SmtpAddress", None))
            except Exception:
                account_display = None
        # 'On behalf of' como reforço (requer permissão no Exchange)
        try:
            m.SentOnBehalfOfName = account_hint
            if verbose:
                print(f"[email] SentOnBehalfOfName set to: {account_hint}")
        except Exception as e:
            if verbose:
                print(f"[email] SentOnBehalfOfName failed: {e}")

    # se não informaram account_hint, tenta ler a padrão ANTES do send
    if account_display is None:
        try:
            su = getattr(m, "SendUsingAccount", None)
            if su is not None:
                account_display = str(getattr(su, "DisplayName", None) or getattr(su, "SmtpAddress", None))
        except Exception:
            account_display = None

    m.Subject = subject or _default_subject(sessao)
    html_body = body or _default_body(sessao)
    m.HTMLBody = html_body

    if to_list:
        m.To = "; ".join(to_list)
    if cc_list:
        m.CC = "; ".join(cc_list)
    if bcc_list:
        m.BCC = "; ".join(bcc_list)

    m.Attachments.Add(str(attach_abs))

    # resolve recipientes ANTES do send
    recipients_ok = True
    try:
        recipients_ok = bool(m.Recipients.ResolveAll())
        if verbose:
            print(f"[email] Recipients.ResolveAll() => {recipients_ok}")
    except Exception:
        pass

    entry_id = None  # só teremos em rascunho/salvo

    if preview:
        if verbose:
            print("[email] Abrindo pré-visualização (não envia).")
        m.Display(False)
        try:
            entry_id = getattr(m, "EntryID", None)
        except Exception:
            entry_id = None
        status = "preview"
    elif save_to_drafts:
        if verbose:
            print("[email] Salvando em Rascunhos.")
        m.Save()
        try:
            entry_id = getattr(m, "EntryID", None)
        except Exception:
            entry_id = None
        status = "drafts"
    else:
        if verbose:
            print("[email] Enviando…")
        m.Send()
        status = "sent"

    if force_sync:
        _force_sync(ns, verbose=verbose)
        time.sleep(2)

    outbox_after, sent_after = _folder_counts(ns)
    online_after = _online_state(ns)

    # IMPORTANTE: não tocar mais no objeto m após Send() para evitar "item movido/excluído".
    return SendResult(
        status=status,
        account=account_display,
        recipients_resolved=bool(recipients_ok),
        entry_id=entry_id,
        outbox_before=outbox_before,
        outbox_after=outbox_after,
        sent_before=sent_before,
        sent_after=sent_after,
        online_before=online_before,
        online_after=online_after,
        attachment=str(attach_abs),
    )
