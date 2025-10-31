# downloader.py
from __future__ import annotations
import re
import time
import unicodedata
from pathlib import Path
from typing import Iterable, Optional, List

from playwright.sync_api import Page, Download, Frame, ElementHandle
from playwright.sync_api import TimeoutError as PlayTimeout, Error as PWError


# -----------------------------
# Utilidades
# -----------------------------
def _slug(s: str) -> str:
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")
    return s.upper() or "DESCONHECIDO"

def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s).strip().upper()
    s = re.sub(r"\s+\d+$", "", s)  # remove contador no final ("NOME 12")
    return s

def _goto_pagina_pauta(page: Page, base_url: str) -> None:
    url = f"{base_url}/paginas/resultado/consultarSessoesParaGabinete.aspx"
    page.goto(url, wait_until="domcontentloaded")
    page.wait_for_load_state("networkidle")

def _fill_masked_input(page_or_frame, selectors: Iterable[str], value: str) -> bool:
    for sel in selectors:
        try:
            loc = page_or_frame.locator(sel)
            if loc.count() == 0:
                continue
            el = loc.first
            el.click(timeout=2500)
            try:
                el.press("Control+A"); el.press("Delete")
            except PWError:
                pass
            el.fill(value, timeout=2500)
            page_or_frame.keyboard.press("Tab")
            return True
        except PWError:
            continue
    return False

def _clicar_pesquisar_robusto(scope) -> bool:
    candidatos = [
        "#btnPesquisar", "[id='btnPesquisar']",
        "li[title*='Pesquisar']", "li:has-text('Pesquisar')",
        "span:has-text('Pesquisar')", "a:has-text('Pesquisar')",
        "button:has-text('Pesquisar')",
        "[id*='btnPesquisar']",
        "input[type='submit'][name='btnPesquisar']",
    ]
    for sel in candidatos:
        loc = scope.locator(sel)
        if loc.count() == 0:
            continue
        try:
            loc.first.click(timeout=2000); return True
        except PWError:
            try:
                loc.first.click(timeout=2000, force=True); return True
            except PWError:
                pass
    try:
        scope.keyboard.press("Enter"); return True
    except PWError:
        return False

def _preencher_filtros(page: Page, num_sessao: str, data_de: str, data_ate: str) -> None:
    ok1 = _fill_masked_input(page, ["#spnNumSessao_I", "input[id*='spnNumSessao']", "input[name*='NumSessao']"], num_sessao)
    ok2 = _fill_masked_input(page, ["#dteInicial_I", "input[id*='dteInicial']"], data_de)
    ok3 = _fill_masked_input(page, ["#dteFinal_I", "input[id*='dteFinal']"], data_ate)
    if not (ok1 and ok2 and ok3):
        raise RuntimeError("Não foi possível preencher todos os filtros (sessão e datas).")

# -------- Abertura da sessão + captura do iframe do popup --------
def _clicar_botao_consulta_da_pauta(page: Page, num_sessao: str, ano: str) -> None:
    alvo = f"{num_sessao}/{ano}"
    # encontra a TR com o texto da sessão
    row = page.locator(f"tr.dxgvDataRow:has-text('{alvo}')").first
    if row.count() == 0:
        row = page.locator(f"tr:has(td:has-text('{alvo}'))").first
    if row.count() == 0:
        raise RuntimeError(f"Não encontrei a sessão {alvo} na grid.")

    # dentro da linha, clique no botão da coluna Ação (id gvConsulta_DXCBtn*)
    btn = row.locator("a[id^='gvConsulta_DXCBtn']").first
    if btn.count() == 0:
        # fallback: clicar no <img> com alt/title "ConsultaDaPauta"
        btn = row.locator("img[alt='ConsultaDaPauta'], img[title*='Consulta de Processos']").first
        if btn.count() == 0:
            raise RuntimeError("Não achei o botão 'ConsultaDaPauta' na linha da sessão.")
    try:
        btn.click(timeout=3000)
    except PWError:
        btn.click(timeout=3000, force=True)

def _achar_iframe_element(page: Page) -> Optional[ElementHandle]:
    """Procura o elemento <iframe> do popup da SONP."""
    seletores_iframe = [
        "div[id^='popProtocolosDaSessao'] iframe",
        "iframe[id*='popProtocolosDaSessao' i]",
        "iframe[name*='popProtocolosDaSessao' i]",
        "iframe[src*='processosDaPautaPorGabinete.aspx' i]",
    ]
    for sel in seletores_iframe:
        try:
            loc = page.locator(sel)
            if loc.count() > 0:
                # retorna o primeiro correspondente
                return loc.first.element_handle()
        except PWError:
            continue
    return None

def _esperar_iframe_sonp(page: Page, timeout_ms: int = 30000) -> Frame:
    """
    Após clicar em 'ConsultaDaPauta', o DevExpress abre um popup (popProtocolosDaSessao)
    e carrega seu conteúdo via <iframe>. Aqui aguardamos o <iframe> aparecer e ficar utilizável.
    Procuramos tanto por frame com URL contendo 'processosDaPautaPorGabinete.aspx'
    quanto pelo elemento <iframe> dentro do popup (que pode iniciar como about:blank).
    """
    alvo_substr = "processosDaPautaPorGabinete.aspx"
    t0 = time.time()
    ultimo_handle: Optional[ElementHandle] = None

    while (time.time() - t0) * 1000 < timeout_ms:
        # 1) procura por frames já resolvidos na árvore
        for fr in page.frames:
            try:
                if alvo_substr in (fr.url or ""):
                    try:
                        fr.wait_for_selector("body", timeout=1500)
                    except PWError:
                        pass
                    return fr
            except PWError:
                continue

        # 2) procura explicitamente pelo elemento <iframe> do popup
        eh = _achar_iframe_element(page)
        if eh:
            ultimo_handle = eh
            try:
                fr = eh.content_frame()
            except PWError:
                fr = None
            if fr:
                # pode começar como about:blank; esperamos o conteúdo real
                try:
                    fr.wait_for_selector("body", timeout=1500)
                except PWError:
                    pass
                # Se a URL ainda for about:blank, damos mais um tempo para trocar o src
                if alvo_substr in (fr.url or ""):
                    return fr

        time.sleep(0.25)

    # Se chegamos aqui, tente ao menos devolver um frame válido do último handle visto
    if ultimo_handle:
        fr = ultimo_handle.content_frame()
        if fr:
            return fr
    raise RuntimeError("Iframe da SONP não apareceu (popup 'ConsultaDaPauta').")

# ---------- Operações dentro da SONP (no frame) ----------
def _listar_conselheiros(scope) -> List[str]:
    """
    Lista os nomes das abas de conselheiros dentro do escopo (Frame).
    Aceita várias estruturas (DevExpress, tabs com role, nav-tabs etc.).
    """
    candidatos_sel = [
        "li .titulo-tab",
        "ul[role='tablist'] li",
        "li[role='tab']",
        ".dxpc-tab, .dxTab, .nav-tabs li",
        "li:has(a)",
    ]
    # aguarda qualquer estrutura aparecer
    try:
        scope.wait_for_selector(", ".join(candidatos_sel), timeout=20000)
    except PWError:
        pass

    nomes: List[str] = []
    vistos = set()
    for sel in candidatos_sel:
        try:
            items = scope.locator(sel)
            n = items.count()
        except PWError:
            n = 0
        for i in range(n):
            try:
                txt = items.nth(i).inner_text(timeout=1200).strip()
            except PWError:
                continue
            base = _norm(txt)
            if not base:
                continue
            if base in {"PLENO", "PLENARIO", "PLENÁRIO", "TODOS"}:
                continue
            if len(base) < 3:
                continue
            if base not in vistos:
                vistos.add(base)
                nomes.append(base)

    if not nomes:
        raise RuntimeError("Não encontrei títulos das abas de conselheiro dentro da SONP.")
    return nomes

def _ativar_aba_conselheiro(scope, nome_upper: str) -> None:
    candidatos_sel = [
        "li .titulo-tab", "ul[role='tablist'] li", "li[role='tab']",
        ".dxpc-tab, .dxTab, .nav-tabs li", "li:has(a)"
    ]
    # 1) clique direto por comparação de texto
    for sel in candidatos_sel:
        items = scope.locator(sel)
        try:
            n = items.count()
        except PWError:
            n = 0
        for i in range(n):
            try:
                li = items.nth(i)
                txt = li.inner_text(timeout=1200)
            except PWError:
                continue
            if _norm(txt) == nome_upper:
                try:
                    li.click(timeout=1500); time.sleep(0.3); return
                except PWError:
                    pass

    # 2) fallback via DevExpress SetActiveTabIndex (se existir)
    try:
        # mapeia textos para índice
        arr = []
        for sel in candidatos_sel:
            items = scope.locator(sel)
            try:
                n = items.count()
            except PWError:
                n = 0
            for i in range(n):
                try:
                    t = items.nth(i).inner_text(timeout=800)
                except PWError:
                    continue
                arr.append(_norm(t))
        if nome_upper in arr:
            idx = arr.index(nome_upper)
            scope.evaluate(
                """(i)=>{
                    try {
                      if (window.cbp_pcConselheiros?.SetActiveTabIndex) {
                        cbp_pcConselheiros.SetActiveTabIndex(i); return;
                      }
                      if (window.pcConselheiros?.SetActiveTabIndex) {
                        pcConselheiros.SetActiveTabIndex(i); return;
                      }
                      const tl = document.querySelector('ul[role=tablist]');
                      if (tl && tl.children[i]) tl.children[i].click();
                    } catch(e) {}
                }""",
                idx
            )
            time.sleep(0.3); return
    except PWError:
        pass
    raise RuntimeError(f"Não consegui ativar a aba do conselheiro: {nome_upper}")

def _aba_tem_processos(scope) -> bool:
    rows = scope.locator("tr[id*='DXDataRow'], tr.dxgvDataRow, .dxgvDataRow")
    try:
        return rows.count() > 0
    except PWError:
        return False

def _clicar_exportar_excel(scope) -> Optional[Download]:
    candidatos = [
        "a[title*='Excel' i]",
        "a:has-text('Exportar Excel')",
        "button:has-text('Excel')",
        "li:has-text('Exportar Excel')",
        "img[alt*='Excel' i]",
        "[id*='btnExport'][id*='Xls']",
        "[id*='btnExport'][id*='Xlsx']",
        "[id*='btnExportar'][id*='Excel']",
        "a:has(svg[aria-label*='Excel'])",
    ]
    for sel in candidatos:
        loc = scope.locator(sel)
        if loc.count() == 0:
            continue
        try:
            owner_page: Page = getattr(scope, "page", None) or scope
            with owner_page.expect_download(timeout=20000) as ev:
                loc.first.click(timeout=3000)
            return ev.value
        except (PWError, PlayTimeout):
            continue
    return None

def _salvar_download(d: Download, destino: Path) -> None:
    try:
        tmp = d.path()
    except PlayTimeout:
        tmp = None
    if tmp:
        destino.write_bytes(Path(tmp).read_bytes())
    else:
        d.save_as(str(destino))


# -----------------------------
# Fluxo público
# -----------------------------
def baixar_planilhas_sessao(
    page: Page,
    base_url: str,
    num_sessao: str,
    data_de: str,
    data_ate: str,
    download_dir: str,
) -> List[Path]:
    """
    1) Abrir pesquisa e filtrar (consultarSessoesParaGabinete.aspx).
    2) Clicar no botão 'ConsultaDaPauta' da linha {num_sessao}/2025.
    3) Capturar o iframe do popup (processosDaPautaPorGabinete.aspx).
    4) Dentro do iframe, iterar abas de conselheiros e exportar Excel.
    """
    Path(download_dir).mkdir(parents=True, exist_ok=True)

    # Página de pesquisa
    _goto_pagina_pauta(page, base_url)
    _preencher_filtros(page, num_sessao, data_de, data_ate)
    if not _clicar_pesquisar_robusto(page):
        raise RuntimeError("Não consegui acionar a pesquisa (Pesquisar).")
    page.wait_for_load_state("networkidle")
    time.sleep(0.5)

    # Abre popup da sessão e obtém o iframe da SONP
    _clicar_botao_consulta_da_pauta(page, num_sessao, "2025")
    sonp_frame = _esperar_iframe_sonp(page, timeout_ms=35000)

    # Lista dinâmica de conselheiros
    nomes = _listar_conselheiros(sonp_frame)
    baixados: List[Path] = []
    vistos: set[str] = set()

    for nome in nomes:
        if nome in vistos:
            continue
        vistos.add(nome)

        _ativar_aba_conselheiro(sonp_frame, nome)
        time.sleep(0.3)

        if not _aba_tem_processos(sonp_frame):
            # Sem processos para este conselheiro; segue
            continue

        dest = Path(download_dir) / f"PLENARIO_{_slug(nome)}.xlsx"
        if dest.exists() and dest.stat().st_size > 0:
            baixados.append(dest)
            continue

        d = _clicar_exportar_excel(sonp_frame)
        if not d:
            time.sleep(0.7)
            d = _clicar_exportar_excel(sonp_frame)
        if not d:
            print(f"[warn] Não encontrei botão de exportar Excel para {nome}. Pulando…")
            continue

        try:
            _salvar_download(d, dest)
            baixados.append(dest)
            print(f"✔ Baixado: {dest.name}")
        except Exception as e:
            print(f"[warn] Falha ao salvar download de {nome}: {e}")

    if not baixados:
        print("[info] Nenhuma planilha foi baixada (talvez não haja processos).")

    return baixados
