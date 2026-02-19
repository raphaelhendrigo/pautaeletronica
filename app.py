# app.py
from __future__ import annotations
import argparse
import os
import sys
from pathlib import Path
from main import run_pipeline
from pautas_consulta import run_consulta_pautas_pipeline
from settings import ConfigError, env, get_etcm_config, load_env

# Envio via Outlook (pywin32)
try:
    from email_outlook import send_pauta_unificada, SendResult
except Exception:
    send_pauta_unificada = None  # type: ignore
    SendResult = None            # type: ignore


def str2bool(v):
    if isinstance(v, bool):
        return v
    if v is None:
        return False
    return str(v).strip().lower() in ("1", "true", "t", "yes", "y", "on")


def _find_first_existing(paths: list[Path]) -> Path | None:
    for p in paths:
        try:
            if p.exists():
                return p
        except Exception:
            pass
    return None


def _auto_header_default() -> str:
    env_val = os.getenv("TCM_HEADER_TEMPLATE", "").strip()
    script_dir = Path(__file__).resolve().parent
    cwd = Path.cwd()

    names = [
        "papel_timbrado_tcm.docx",
        "papel_timbrado_tcm.docx.docx",
        "PAPEL TIMBRADO.docx",
        "PAPEL TIMBRADO.DOCX",
    ]

    candidates: list[Path] = []
    if env_val:
        candidates.append(Path(env_val))

    for n in names:
        candidates.append(cwd / n)
        candidates.append(script_dir / n)

    found = _find_first_existing(candidates)
    if found:
        return str(found)
    return "papel_timbrado_tcm.docx"


def parse_args():
    p = argparse.ArgumentParser(
        description="Automação e-TCM: baixar planilhas, gerar PAUTA_UNIFICADA e (opcionalmente) enviar por Outlook."
    )
    # Execução/pipeline
    p.add_argument(
        "--modo",
        type=str,
        default="sonp",
        choices=["sonp", "consulta-pautas"],
        help="sonp (padrao) | consulta-pautas (exporta Excel consolidado)",
    )
    p.add_argument("--headless", type=str, default="true", help="true|false (padrão: true)")
    p.add_argument(
        "--base-url",
        type=str,
        default=env("ETCM_BASE_URL", env("BASE_URL", "https://etcm.tcm.sp.gov.br")),
        help="Base URL (padrão produção).",
    )
    p.add_argument("--sessao", type=str, default="71", help="Número da sessão (ex.: 71).")
    p.add_argument("--de", type=str, default="29/09/2025", help="Data inicial DD/MM/AAAA.")
    p.add_argument("--ate", type=str, default="29/10/2025", help="Data final DD/MM/AAAA.")
    p.add_argument("--download-dir", type=str, default="planilhas_71_2025", help="Pasta de downloads.")
    p.add_argument("--output-dir", type=str, default="output", help="Pasta de saida (DOCX ou XLSX).")
    p.add_argument(
        "--competencias-download",
        type=str,
        default="pleno,1c,2c",
        help="SONP: competencias a baixar (pleno,1c,2c).",
    )
    p.add_argument(
        "--competencias",
        type=str,
        default="pleno,1c,2c",
        help="Consulta de pautas: competencias separadas por virgula (pleno,1c,2c).",
    )
    p.add_argument("--situacao", type=str, default="", help="Consulta de pautas: filtrar situacao (opcional).")
    p.add_argument("--consolidado-nome", type=str, default="", help="Nome do XLSX consolidado (opcional).")
    p.add_argument("--sem-abas-competencia", action="store_true", help="Nao gerar abas por competencia.")
    p.add_argument("--resumo", action="store_true", help="Gerar aba Resumo no XLSX consolidado.")
    p.add_argument("--dedupe", action="store_true", help="Tentar deduplicar linhas no XLSX consolidado.")

    # Documento
    p.add_argument("--header-template", dest="header_template", default=_auto_header_default())
    p.add_argument("--titulo-docx", dest="titulo_docx", default="Pauta Classificada")
    p.add_argument("--nome-docx", dest="nome_docx", default=None)

    # >>> METADADOS DA SESSÃO (para cabeçalho específico no DOCX) <<<
    # Preencha para SONP/SENP/Pleno/1ª/2ª. Se você passar qualquer --meta-*, os campos obrigatórios são:
    # --meta-tipo, --meta-formato, --meta-competencia, --meta-numero, --meta-data-abertura
    p.add_argument("--meta-numero", type=str, default="", help='Ex.: "71" ou "3.385" (o sufixo "ª" será adicionado se faltar)')
    p.add_argument("--meta-tipo", type=str, default="", choices=["ordinaria", "extraordinaria", ""],
                   help="Tipo: 'ordinaria' | 'extraordinaria'")
    p.add_argument("--meta-formato", type=str, default="", choices=["nao-presencial", "presencial", ""],
                   help="Formato: 'nao-presencial' | 'presencial'")
    p.add_argument("--meta-competencia", type=str, default="", choices=["pleno", "1c", "2c", ""],
                   help="Competência: 'pleno' | '1c' | '2c'")
    p.add_argument("--meta-data-abertura", type=str, default="", help='DD/MM/AAAA (obrigatório quando usar meta)')
    p.add_argument("--meta-data-encerramento", type=str, default="", help='DD/MM/AAAA (NP); se vazio, calcula +15 dias')
    p.add_argument("--meta-horario", type=str, default="", help='Presencial: horário (padrão "9h30min.")')

    # E-mail (parametrizável por CLI/.env)
    default_to = (env("TCM_EMAIL_TO", "") or "").strip()
    p.add_argument("--send-email", action="store_true")
    p.add_argument("--email-preview", action="store_true")
    p.add_argument("--email-drafts", action="store_true")
    p.add_argument("--email-force-sync", action="store_true")
    p.add_argument("--email-to", type=str, default=default_to)
    p.add_argument("--email-cc", type=str, default=env("TCM_EMAIL_CC", "") or "")
    p.add_argument("--email-bcc", type=str, default=env("TCM_EMAIL_BCC", "") or "")
    p.add_argument("--email-subject", type=str, default=env("TCM_EMAIL_SUBJECT", "") or "")
    p.add_argument("--email-body", type=str, default=env("TCM_EMAIL_BODY", "") or "")
    p.add_argument("--email-body-file", type=str, default="")
    p.add_argument("--email-account", type=str, default=env("TCM_EMAIL_ACCOUNT", "") or "")
    p.add_argument("--email-verbose", action="store_true")
    return p.parse_args()


def _export_meta_to_env(args):
    """
    Se o usuário passar metadados da sessão (--meta-*), exporta para variáveis de ambiente
    que o docx_maker.py consumirá automaticamente e gerará o cabeçalho correto.
    Obrigatórios quando usar meta: tipo, formato, competencia, numero, data_abertura.
    """
    must = [args.meta_tipo, args.meta_formato, args.meta_competencia, args.meta_numero, args.meta_data_abertura]
    any_meta = any(bool(x) for x in must + [args.meta_data_encerramento, args.meta_horario])
    if not any_meta:
        return  # nada informado -> uso do cabeçalho padrão existente

    if not all(must):
        raise SystemExit(
            "Para usar o cabeçalho específico, informe TODOS: "
            "--meta-tipo --meta-formato --meta-competencia --meta-numero --meta-data-abertura "
            "(--meta-data-encerramento é opcional para Não Presencial; --meta-horario opcional para Presencial)."
        )

    os.environ["TCM_META_TIPO"] = args.meta_tipo
    os.environ["TCM_META_FORMATO"] = args.meta_formato
    os.environ["TCM_META_COMPETENCIA"] = args.meta_competencia
    os.environ["TCM_META_NUMERO"] = args.meta_numero
    os.environ["TCM_META_DATA_ABERTURA"] = args.meta_data_abertura
    if args.meta_data_encerramento:
        os.environ["TCM_META_DATA_ENCERRAMENTO"] = args.meta_data_encerramento
    os.environ["TCM_META_HORARIO"] = args.meta_horario or "9h30min."

    # Forca abertura/encerramento exatamente como informado no CLI
    os.environ["TCM_META_ABERTURA_FINAL"] = args.meta_data_abertura
    if args.meta_data_encerramento:
        os.environ["TCM_META_ENCERRAMENTO_FINAL"] = args.meta_data_encerramento
    else:
        try:
            from datetime import datetime, timedelta
            d = datetime.strptime(args.meta_data_abertura, "%d/%m/%Y")
            # 15 dias corridos, contando o dia inicial -> +16 dias no calendario
            os.environ["TCM_META_ENCERRAMENTO_FINAL"] = (d + timedelta(days=16)).strftime("%d/%m/%Y")
        except Exception:
            pass


def main():
    load_env()
    args = parse_args()
    headless = str2bool(args.headless)

    try:
        etcm = get_etcm_config()
    except ConfigError as e:
        raise SystemExit(str(e))

    if args.modo == "consulta-pautas":
        if args.send_email:
            raise SystemExit("Envio por Outlook disponivel apenas no modo sonp.")

        comp_list = [c.strip() for c in (args.competencias or "").split(",") if c.strip()]
        nome_consolidado = args.consolidado_nome or None
        sessao_input = None
        if any(a.startswith("--sessao") for a in sys.argv[1:]):
            sessao_input = args.sessao

        run_consulta_pautas_pipeline(
            base_url=(args.base_url or etcm.base_url).rstrip("/"),
            usuario=etcm.username,
            senha=etcm.password,
            data_de=args.de,
            data_ate=args.ate,
            download_dir=args.download_dir,
            output_dir=args.output_dir,
            headless=headless,
            competencias=comp_list,
            num_sessao=sessao_input,
            situacao=args.situacao or None,
            nome_consolidado=nome_consolidado,
            include_competencia_sheets=not args.sem_abas_competencia,
            include_resumo=bool(args.resumo),
            dedupe=bool(args.dedupe),
        )
        return

    # Injeta metadados (se informados) para o docx_maker
    _export_meta_to_env(args)

    pipeline_kwargs = dict(
        base_url=(args.base_url or etcm.base_url).rstrip("/"),
        usuario=etcm.username,
        senha=etcm.password,
        num_sessao=args.sessao,
        data_de=args.de,
        data_ate=args.ate,
        download_dir=args.download_dir,
        output_dir=args.output_dir,
        headless=headless,
        competencia=(args.meta_competencia or None),
    )
    comp_download = [c.strip() for c in (args.competencias_download or "").split(",") if c.strip()]
    pipeline_kwargs["competencias_download"] = comp_download or None

    optional_kwargs = {}
    if args.header_template:
        optional_kwargs["header_template"] = args.header_template
    if args.titulo_docx:
        optional_kwargs["titulo_docx"] = args.titulo_docx
    if args.nome_docx:
        optional_kwargs["nome_docx"] = args.nome_docx

    # Executa pipeline (login -> download planilhas -> gera DOCX)
    try:
        run_pipeline(**pipeline_kwargs, **optional_kwargs)
    except TypeError:
        # Compat com versões antigas de run_pipeline
        run_pipeline(**pipeline_kwargs)

    # Envio por Outlook (opcional)
    if args.send_email:
        if not args.email_to:
            raise SystemExit("Defina --email-to ou TCM_EMAIL_TO para envio.")
        if send_pauta_unificada is None:
            raise SystemExit("Envio indisponível. Confirme email_outlook.py e pywin32 instalado.")

        body_text = args.email_body
        if not body_text and args.email_body_file:
            try:
                with open(args.email_body_file, "r", encoding="utf-8") as f:
                    body_text = f.read()
            except Exception as e:
                print(f"[email] Falha ao ler --email-body-file: {e}", file=sys.stderr)

        result = send_pauta_unificada(
            docx_path=None,
            output_dir=args.output_dir,
            sessao=str(args.sessao) if args.sessao else None,
            to=args.email_to or None,
            cc=args.email_cc or None,
            bcc=args.email_bcc or None,
            subject=args.email_subject or None,
            body=body_text or None,
            preview=bool(args.email_preview),
            save_to_drafts=bool(args.email_drafts),
            account_hint=args.email_account or None,
            verbose=bool(args.email_verbose),
            force_sync=bool(args.email_force_sync),
        )
        print(f"[email] status={result.status} account={result.account or '-'}")
        print(f"[email] recipients_resolved={result.recipients_resolved} entry_id={result.entry_id or '-'}")
        print(f"[email] outbox_before={result.outbox_before} -> outbox_after={result.outbox_after}")
        print(f"[email] sent_before={result.sent_before} -> sent_after={result.sent_after}")
        print(f"[email] online_before={result.online_before} -> online_after={result.online_after}")
        print(f"[email] anexado: {result.attachment}")
        if result.log_path:
            print(f"[email] log registrado em: {result.log_path}")
        unresolved = [r for r in result.recipient_status if not r.resolved]
        if unresolved:
            print("[email][aviso] Destinatários não resolvidos:")
            for item in unresolved:
                who = item.original or item.display or item.address or "(desconhecido)"
                reason = item.reason or "Motivo não informado."
                print(f"  - {who}: {reason}")


if __name__ == "__main__":
    main()
