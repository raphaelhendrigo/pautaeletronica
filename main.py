from __future__ import annotations

import re
from pathlib import Path
from playwright.sync_api import sync_playwright

from login import efetuar_login
from downloader import baixar_planilhas_sessao
from docx_maker import gerar_docx_unificado, gerar_docx_vazio


def _guess_year(*dates: str) -> str:
    """Extrai um ano (AAAA) do primeiro argumento que contiver 4 dígitos; fallback 2025."""
    for d in dates:
        if not d:
            continue
        m = re.search(r"(\d{4})", str(d))
        if m:
            return m.group(1)
    return "2025"


def _candidate_paths(name: str) -> list[Path]:
    """
    Monta uma lista de caminhos candidatos para o timbrado, considerando:
    - nome passado (relativo/absoluto)
    - variação com '.docx.docx'
    - nomes antigos
    - procura em CWD e na pasta deste arquivo
    """
    cwd = Path.cwd()
    here = Path(__file__).resolve().parent

    given = Path(name)

    if given.name.lower() == "papel_timbrado_tcm.docx":
        base_names = [
            "papel_timbrado_tcm.docx",
            "papel_timbrado_tcm.docx.docx",
        ]
    else:
        base_names = [given.name, given.name + ".docx"]

    base_names += ["PAPEL TIMBRADO.docx", "PAPEL TIMBRADO.DOCX"]

    cands: list[Path] = []
    cands.append(given)
    if not given.is_absolute():
        cands.append(cwd / given)
        cands.append(here / given)

    for n in base_names:
        cands.append(cwd / n)
        cands.append(here / n)

    uniq: list[Path] = []
    seen = set()
    for p in cands:
        key = str(p.resolve()) if p.exists() else str(p)
        if key not in seen:
            uniq.append(p)
            seen.add(key)
    return uniq


def _resolve_header_template(header_template: str | None) -> str | None:
    """
    Resolve um caminho existente para o timbrado.
    - Se `header_template` vier vazio, tenta automaticamente os nomes padrão.
    - Se vier um caminho inválido, tentamos as variações automaticamente.
    Retorna caminho str existente ou None.
    """
    if header_template:
        cands = _candidate_paths(header_template)
    else:
        cands = _candidate_paths("papel_timbrado_tcm.docx")

    for p in cands:
        try:
            if p.exists() and p.is_file():
                print(f"[docx] Usando papel timbrado: {p}")
                return str(p)
        except Exception:
            pass

    print("[docx] Aviso: papel timbrado não encontrado. O DOCX será gerado sem cabeçalho do template.")
    return None


def run_pipeline(
    base_url: str,
    usuario: str,
    senha: str,
    num_sessao: str,
    data_de: str,
    data_ate: str,
    download_dir: str,
    output_dir: str,
    headless: bool = True,
    titulo_docx: str | None = None,
    header_template: str | None = None,
    nome_docx: str | None = None,
) -> str:
    """
    1) Abre navegador, faz login e baixa as planilhas da sessão.
    2) Gera o DOCX unificado (ou vazio, se não houver itens) com cabeçalho opcional.
    Retorna o caminho absoluto do DOCX gerado.
    """
    download_path = Path(download_dir)
    output_path = Path(output_dir)
    download_path.mkdir(parents=True, exist_ok=True)
    output_path.mkdir(parents=True, exist_ok=True)

    print(f"[1/3] Abrindo navegador (headless={headless}) e fazendo login.")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(
            accept_downloads=True,
            locale="pt-BR",
            timezone_id="America/Sao_Paulo",
        )
        page = context.new_page()
        try:
            efetuar_login(page, base_url, usuario, senha)

            print(f"[2/3] Baixando planilhas da sessão {num_sessao}/2025.")
            baixar_planilhas_sessao(
                page=page,
                base_url=base_url,
                num_sessao=num_sessao,
                data_de=data_de,
                data_ate=data_ate,
                download_dir=str(download_path),
            )
        finally:
            context.close()
            browser.close()

    # Monta nome do arquivo de saída
    ano = _guess_year(data_ate, data_de)
    if not titulo_docx:
        titulo_docx = f"Pauta Unificada - Sessão {num_sessao}/{ano}"
    nome_arquivo = nome_docx or f"PAUTA_UNIFICADA_{num_sessao}_{ano}.docx"
    saida_docx = output_path / nome_arquivo

    # Resolve o timbrado
    header_resolvido = _resolve_header_template(header_template)

    # Decide o modo de geração do DOCX
    xls = list(download_path.glob("*.xls*"))
    if not xls:
        print("[3/3] Nenhuma planilha encontrada. Gerando DOCX sem itens (apenas cabeçalho).")
        out_path = gerar_docx_vazio(
            saida_docx=str(saida_docx),
            titulo=titulo_docx,
            header_template=header_resolvido,
        )
    else:
        print("[3/3] Gerando documento unificado DOCX.")
        out_path = gerar_docx_unificado(
            pasta_planilhas=str(download_path),
            saida_docx=str(saida_docx),
            titulo=titulo_docx,
            header_template=header_resolvido,
        )

    # Evita caractere não-ASCII (✓) que quebra em consoles CP-1252/437
    print(f"Concluido: {out_path}")
    return out_path
