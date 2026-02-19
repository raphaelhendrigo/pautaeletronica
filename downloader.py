# downloader.py
from __future__ import annotations
import re
import time
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional, List, Callable

import pandas as pd
from playwright.sync_api import Page, Download, Frame, ElementHandle
from playwright.sync_api import TimeoutError as PlayTimeout, Error as PWError

# -----------------------------
# Utilidades
# -----------------------------
def _slug(s: str) -> str:
    return normalize_text(s)

def _norm(s: str) -> str:
    return normalize_text(s)

def normalize_text(s: str) -> str:
    txt = unicodedata.normalize("NFKD", str(s))
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    txt = txt.upper()
    txt = re.sub(r"\s+", " ", txt).strip()
    txt = re.sub(r"[^A-Z0-9 ]+", "", txt)
    txt = txt.replace(" ", "_")
    txt = re.sub(r"_+", "_", txt).strip("_")
    return txt or "DESCONHECIDO"

def _ws(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()

_NAME_MAP = {
    "ET": "EDUARDO TUMA",
    "DD": "DOMINGOS DISSEI",
    "JA": "JOAO ANTONIO",
    "RT": "RICARDO TORRES",
    "RB": "ROBERTO BRAGUIM",
}

def _expand_relator_name(value: str) -> str:
    s = _ws(value)
    if not s:
        return ""
    code = re.sub(r"[^A-Za-z]", "", s).upper()
    if 1 <= len(code) <= 3 and code in _NAME_MAP:
        return _NAME_MAP[code]
    return s.upper()

def _clean_tab_label(text: str) -> str:
    if not text:
        return ""
    cleaned = re.sub(r"\s+", " ", text).strip()
    cleaned = re.sub(r"\s+\d+\s*$", "", cleaned).strip()
    return cleaned

def _looks_generic_tab_key(key: str) -> bool:
    if not key:
        return True
    if re.fullmatch(r"C\d+", key):
        return True
    if key.startswith("TAB_"):
        return True
    return False

def _guess_year(*dates: str) -> str:
    """Extrai um ano (AAAA) do primeiro argumento que contiver 4 digitos; fallback 2025."""
    for d in dates:
        if not d:
            continue
        m = re.search(r"(\d{4})", str(d))
        if m:
            return m.group(1)
    return "2025"

def _normalize_competencia_key(value: str | None) -> str | None:
    if not value:
        return None
    key = normalize_text(value)
    if key in {"PLENO", "PLENARIO"}:
        return "PLENO"
    if key in {"1C", "1A", "1_CAMARA", "1A_CAMARA"}:
        return "1_CAMARA"
    if key in {"2C", "2A", "2_CAMARA", "2A_CAMARA"}:
        return "2_CAMARA"
    if "1" in key and "CAMARA" in key:
        return "1_CAMARA"
    if "2" in key and "CAMARA" in key:
        return "2_CAMARA"
    return key

def _match_row_by_competencia(row_text: str, alvo_norm: str, comp_key: str | None) -> bool:
    text_norm = normalize_text(row_text)
    if alvo_norm not in text_norm:
        return False
    if not comp_key:
        return True
    if comp_key == "PLENO" and "PLENO" in text_norm:
        return True
    if comp_key == "1_CAMARA" and ("1_CAMARA" in text_norm or "1A_CAMARA" in text_norm):
        return True
    if comp_key == "2_CAMARA" and ("2_CAMARA" in text_norm or "2A_CAMARA" in text_norm):
        return True
    return comp_key in text_norm

def _prefix_from_competencia(competencia: str | None) -> str:
    key = _normalize_competencia_key(competencia)
    if key == "1_CAMARA":
        return "1CAMARA"
    if key == "2_CAMARA":
        return "2CAMARA"
    return "PLENARIO"


@dataclass
class PlanilhaStats:
    conselheiro: str
    conselheiro_norm: str
    path: Path
    itens: int
    tamanho_bytes: int
    evidencias: list[str]
    prefixo: str = ""


@dataclass
class TabRef:
    key: str
    label: str
    handle: Optional[ElementHandle]
    index: Optional[int]


@dataclass
class SessaoProcessosEsperados:
    total: Optional[int]
    por_sigla: dict[str, int]
    por_conselheiro_norm: dict[str, int]


def _parse_int(value: str) -> Optional[int]:
    raw = re.sub(r"[^0-9]", "", _ws(value))
    if not raw:
        return None
    try:
        return int(raw)
    except Exception:
        return None


def _format_esperados(esperado: SessaoProcessosEsperados) -> str:
    partes: list[str] = []
    if esperado.total is not None:
        partes.append(f"TOTAL={esperado.total}")
    for sigla in ("DD", "JA", "RB", "ET", "RT"):
        if sigla in esperado.por_sigla:
            partes.append(f"{sigla}={esperado.por_sigla[sigla]}")
    return " ".join(partes)


def _extrair_processos_esperados_da_linha(page: Page, row) -> SessaoProcessosEsperados:
    keys = ("TOTAL", "DD", "JA", "RB", "ET", "RT")
    por_sigla: dict[str, int] = {}

    header_idx: dict[str, int] = {}
    headers = page.locator("tr.dxgvHeaderRow td, tr.dxgvHeaderRow th")
    try:
        hcount = headers.count()
    except PWError:
        hcount = 0

    for i in range(hcount):
        try:
            txt = headers.nth(i).inner_text(timeout=900).strip()
        except PWError:
            continue
        norm = normalize_text(txt)
        if norm in keys:
            header_idx[norm] = i

    cells = row.locator("td")
    try:
        ccount = cells.count()
    except PWError:
        ccount = 0

    if header_idx:
        for key in keys:
            idx = header_idx.get(key)
            if idx is None or idx >= ccount:
                continue
            try:
                txt = cells.nth(idx).inner_text(timeout=900).strip()
            except PWError:
                continue
            num = _parse_int(txt)
            if num is not None:
                por_sigla[key] = num

    if not por_sigla:
        row_text = ""
        try:
            row_text = _ws(row.inner_text(timeout=1200))
        except PWError:
            row_text = ""

        m = re.search(
            r"(Aberta|Fechada)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)",
            row_text,
            flags=re.IGNORECASE,
        )
        if m:
            nums = [int(m.group(i)) for i in range(2, 8)]
            por_sigla = dict(zip(keys, nums))
        else:
            nums = [int(n) for n in re.findall(r"\b\d+\b", row_text)]
            if len(nums) >= 6:
                por_sigla = dict(zip(keys, nums[-6:]))

    por_conselheiro_norm: dict[str, int] = {}
    for sigla, nome in _NAME_MAP.items():
        if sigla in por_sigla:
            por_conselheiro_norm[normalize_text(nome)] = por_sigla[sigla]

    return SessaoProcessosEsperados(
        total=por_sigla.get("TOTAL"),
        por_sigla=por_sigla,
        por_conselheiro_norm=por_conselheiro_norm,
    )


def _limpar_planilhas_por_prefixo(download_dir: Path, prefixo: str) -> None:
    for item in download_dir.glob(f"{prefixo}_*.xlsx"):
        try:
            item.unlink()
        except Exception:
            pass


def _contar_processos_por_conselheiro_arquivo(download_dir: Path, prefixo: str) -> dict[str, int]:
    out: dict[str, int] = {}
    for path in download_dir.glob(f"{prefixo}_*.xlsx"):
        stem = path.stem
        marker = f"{prefixo}_"
        if not stem.startswith(marker):
            continue
        conselheiro_norm = stem[len(marker):]
        out[conselheiro_norm] = _contar_itens_planilha(path)
    return out


def _validar_qtd_por_conselheiro(
    download_dir: Path,
    prefixo: str,
    esperado: SessaoProcessosEsperados,
) -> None:
    if not esperado.por_conselheiro_norm:
        raise RuntimeError("Nao foi possivel ler a quantidade esperada de processos por conselheiro na grid.")

    obtido = _contar_processos_por_conselheiro_arquivo(download_dir, prefixo)
    inconsistencias: list[str] = []

    for conselheiro_norm, qtd_esperada in esperado.por_conselheiro_norm.items():
        qtd_obtida = obtido.get(conselheiro_norm)
        if qtd_obtida is None:
            inconsistencias.append(f"{conselheiro_norm}: esperado {qtd_esperada}, arquivo nao encontrado")
        elif qtd_obtida != qtd_esperada:
            inconsistencias.append(f"{conselheiro_norm}: esperado {qtd_esperada}, obtido {qtd_obtida}")

    if esperado.total is not None:
        soma = sum(obtido.get(nome, 0) for nome in esperado.por_conselheiro_norm.keys())
        if soma != esperado.total:
            inconsistencias.append(f"TOTAL: esperado {esperado.total}, obtido {soma}")

    if inconsistencias:
        raise RuntimeError("Divergencia nas quantidades por conselheiro: " + " | ".join(inconsistencias))

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
def _clicar_botao_consulta_da_pauta(
    page: Page,
    num_sessao: str,
    ano: str,
    competencia: str | None = None,
) -> SessaoProcessosEsperados:
    alvo = f"{num_sessao}/{ano}"
    alvo_norm = normalize_text(alvo)
    comp_key = _normalize_competencia_key(competencia)

    row = None
    rows = page.locator("tr.dxgvDataRow")
    try:
        rcount = rows.count()
    except PWError:
        rcount = 0

    for i in range(rcount):
        loc = rows.nth(i)
        try:
            txt = loc.inner_text(timeout=1200)
        except PWError:
            continue
        if _match_row_by_competencia(txt, alvo_norm, comp_key):
            row = loc
            break

    if row is None:
        if comp_key:
            raise RuntimeError(f"Nao encontrei a sessao {alvo} para a competencia {competencia}.")
        row = page.locator(f"tr.dxgvDataRow:has-text(\"{alvo}\")").first
        if row.count() == 0:
            row = page.locator(f"tr:has(td:has-text(\"{alvo}\"))").first
        if row.count() == 0:
            raise RuntimeError(f"Nao encontrei a sessao {alvo} na grid.")

    esperado = _extrair_processos_esperados_da_linha(page, row)

    try:
        row.scroll_into_view_if_needed(timeout=2000)
    except PWError:
        pass

    candidatos = [
        "a[id^='gvConsulta_DXCBtn']",
        "img[alt='ConsultaDaPauta']",
        "img[title*='Consulta de Processos']",
        "td:has(a) a",
    ]

    for sel in candidatos:
        loc = row.locator(sel).first
        if loc.count() == 0:
            continue
        try:
            loc.click(timeout=5000)
            return esperado
        except PWError:
            try:
                loc.click(timeout=5000, force=True)
                return esperado
            except PWError:
                continue

    raise RuntimeError("Nao achei o botao 'ConsultaDaPauta' na linha da sessao.")
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

def _fechar_popup_sonp(page: Page) -> None:
    seletores = [
        "div[id^='popProtocolosDaSessao'] .dxpc-closeBtn",
        "div[id^='popProtocolosDaSessao'] a.dxpc-closeBtn",
        "div[id^='popProtocolosDaSessao'] .dxpc-closeButton",
        "div[id^='popProtocolosDaSessao'] img[alt='Close']",
        "div[id^='popProtocolosDaSessao'] img[alt='Fechar']",
        "div[id^='popProtocolosDaSessao'] a[title*='Fechar']",
        "div[id^='popProtocolosDaSessao'] a[title*='Close']",
    ]
    closed = False
    for sel in seletores:
        try:
            loc = page.locator(sel)
            if loc.count() == 0:
                continue
            loc.first.click(timeout=2000)
            closed = True
            break
        except PWError:
            continue
    if not closed:
        try:
            page.keyboard.press("Escape")
        except PWError:
            pass
    try:
        page.locator("div[id^='popProtocolosDaSessao']").wait_for(state="hidden", timeout=5000)
    except PWError:
        pass

# ---------- Operações dentro da SONP (no frame) ----------
def _listar_conselheiros_js(scope) -> List[str]:
    try:
        items = scope.evaluate(
            """() => {
                const out = [];
                const norm = (s) => {
                  try {
                    return (s || "").toString().normalize('NFD').replace(/[\\u0300-\\u036f]/g, '')
                      .replace(/\\s+/g, ' ').trim();
                  } catch(e) {
                    return (s || "").toString().replace(/\\s+/g, ' ').trim();
                  }
                };
                const add = (txt) => {
                  const v = norm(txt);
                  if (v) out.push(v);
                };

                const ctrl = window.cbp_pcConselheiros || window.pcConselheiros;
                if (ctrl && ctrl.GetTabCount && ctrl.GetTab) {
                  const n = ctrl.GetTabCount();
                  for (let i = 0; i < n; i++) {
                    const tab = ctrl.GetTab(i);
                    if (!tab) continue;
                    let text = '';
                    try { text = tab.GetText ? tab.GetText() : (tab.text || ''); } catch(e) {}
                    if (text) add(text);
                  }
                }

                const nodes = document.querySelectorAll(
                  'li[role=tab], .dxpc-tab, .dxTab, .nav-tabs li, li .titulo-tab'
                );
                nodes.forEach((el) => add(el.innerText || el.textContent || ''));
                return out;
            }"""
        )
        if isinstance(items, list):
            return [str(x) for x in items if str(x).strip()]
    except PWError:
        pass
    return []

def _listar_conselheiros(scope) -> List[str]:
    """
    Lista os nomes das abas de conselheiros dentro do escopo (Frame).
    Aceita v?rias estruturas (DevExpress, tabs com role, nav-tabs etc.).
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
            if base in {"PLENO", "PLENARIO", "PLEN?RIO", "TODOS"}:
                continue
            if len(base) < 3:
                continue
            if base not in vistos:
                vistos.add(base)
                nomes.append(base)

    js_names = _listar_conselheiros_js(scope)
    for name in js_names:
        base = _norm(name)
        if not base:
            continue
        if base in {"PLENO", "PLENARIO", "PLEN?RIO", "TODOS"}:
            continue
        if len(base) < 3:
            continue
        if base not in vistos:
            vistos.add(base)
            nomes.append(base)

    if not nomes:
        raise RuntimeError("N?o encontrei t?tulos das abas de conselheiro dentro da SONP.")
    return nomes

def _map_conselheiro_tabs(scope, *, allow_empty: bool = False) -> list[TabRef]:
    refs: dict[str, TabRef] = {}
    order: list[str] = []

    # Preferir labels visiveis no DOM (titulo-tab)
    try:
        tab_items = scope.locator("#cbp_pcConselheiros_TC li")
        tab_count = tab_items.count()
    except PWError:
        tab_count = 0

    for i in range(tab_count):
        loc = tab_items.nth(i)
        try:
            label = loc.locator(".titulo-tab").inner_text(timeout=1200).strip()
        except PWError:
            label = ""
        if not label:
            try:
                label = loc.inner_text(timeout=800).strip()
            except PWError:
                label = ""
        label = _clean_tab_label(label)
        key = normalize_text(label)
        if not key or key == "DESCONHECIDO":
            continue
        if key in {"PLENO", "PLENARIO", "PLENARIO", "TODOS"}:
            continue
        if _looks_generic_tab_key(key):
            continue
        idx = None
        try:
            idx = loc.evaluate(
                "el => { const m = el.id && el.id.match(/_T(\\d+)$/); return m ? parseInt(m[1], 10) : null; }"
            )
        except PWError:
            idx = None
        handle = None
        try:
            handle = loc.element_handle()
        except PWError:
            handle = None
        if key not in refs:
            refs[key] = TabRef(key=key, label=label, handle=handle, index=idx)
            order.append(key)
        elif refs[key].index is None and idx is not None:
            refs[key] = TabRef(key=refs[key].key, label=refs[key].label, handle=refs[key].handle, index=idx)

    if refs and any(not _looks_generic_tab_key(k) for k in refs):
        return [refs[k] for k in order]

    # Fallback: seletores genÃ©ricos
    candidatos_sel = [
        "li .titulo-tab",
        "ul[role='tablist'] li",
        "li[role='tab']",
        ".dxpc-tab, .dxTab, .nav-tabs li",
        "li:has(a)",
    ]
    for sel in candidatos_sel:
        try:
            items = scope.locator(sel)
            n = items.count()
        except PWError:
            n = 0
        for i in range(n):
            try:
                loc = items.nth(i)
                txt = loc.inner_text(timeout=1200).strip()
            except PWError:
                continue
            txt = _clean_tab_label(txt)
            key = normalize_text(txt)
            if not key or key == "DESCONHECIDO":
                continue
            if key in {"PLENO", "PLENARIO", "PLENARIO", "TODOS"}:
                continue
            if _looks_generic_tab_key(key):
                continue
            if key not in refs:
                handle = None
                try:
                    handle = loc.element_handle()
                except PWError:
                    handle = None
                refs[key] = TabRef(key=key, label=txt, handle=handle, index=None)
                order.append(key)

    if refs:
        return [refs[k] for k in order]

    # Ultimo recurso: JS (podem ser c1/c2/c3)
    js_refs = _map_conselheiro_tabs_js(scope)
    if js_refs:
        return js_refs

    if not refs and not allow_empty:
        raise RuntimeError("Nao encontrei titulos das abas de conselheiro dentro da SONP.")
    return [refs[k] for k in order]


def _match_joao_tab(refs: list[TabRef]) -> Optional[TabRef]:
    for ref in refs:
        key = ref.key.replace("_", "")
        if "JOAO" in key and "ANTONIO" in key:
            return ref
    return None


def _map_conselheiro_tabs_js(scope) -> list[TabRef]:
    ordem_padrao = [
        "DOMINGOS DISSEI",
        "JOAO ANTONIO",
        "ROBERTO BRAGUIM",
        "EDUARDO TUMA",
        "RICARDO TORRES",
    ]
    try:
        info = scope.evaluate(
            """() => {
                const ctrl = window.cbp_pcConselheiros || window.pcConselheiros;
                if (!ctrl || !ctrl.GetTabCount) return {count: 0, tabs: []};
                const n = ctrl.GetTabCount();
                const tabs = [];
                for (let i=0;i<n;i++) {
                    const t = ctrl.GetTab(i);
                    let name = '';
                    let text = '';
                    try { name = t?.name || t?.GetName?.() || ''; } catch(e) {}
                    try { text = t?.GetText?.() || ''; } catch(e) {}
                    tabs.push({i, name, text});
                }
                return {count: n, tabs};
            }"""
        )
    except PWError:
        return []
    out: list[TabRef] = []
    if not info:
        return out
    tabs = info.get("tabs") if isinstance(info, dict) else None
    if not tabs:
        return out
    for item in tabs:
        idx = item.get("i")
        raw_name = _clean_tab_label(item.get("name") or "")
        raw_text = _clean_tab_label(item.get("text") or "")
        label = raw_text or raw_name or f"TAB_{idx}"
        key = normalize_text(label)
        if _looks_generic_tab_key(key) and isinstance(idx, int) and 0 <= idx < len(ordem_padrao):
            label = ordem_padrao[idx]
            key = normalize_text(label)
        if _looks_generic_tab_key(key):
            continue
        out.append(TabRef(key=key or f"TAB_{idx}", label=label, handle=None, index=idx))
    return out


def _collect_relatores(scope, limit: int = 50) -> list[str]:
    headers = scope.locator("tr.dxgvHeaderRow td, tr.dxgvHeaderRow th")
    hcount = 0
    try:
        hcount = headers.count()
    except PWError:
        hcount = 0
    rel_idx = None
    for h in range(hcount):
        try:
            text = headers.nth(h).inner_text(timeout=1000).strip().lower()
        except PWError:
            continue
        if "relator" in text:
            rel_idx = h
            break
    if rel_idx is None:
        return []

    rows = scope.locator("tr.dxgvDataRow")
    try:
        rcount = rows.count()
    except PWError:
        rcount = 0
    vistos: list[str] = []
    for i in range(min(rcount, limit)):
        try:
            cell = rows.nth(i).locator("td").nth(rel_idx)
            val = cell.inner_text(timeout=1000).strip()
        except PWError:
            continue
        expanded = _expand_relator_name(val)
        if expanded and expanded not in vistos:
            vistos.append(expanded)
    return vistos

def _ativar_aba_por_ref(scope, ref: TabRef) -> None:
    if ref.handle:
        try:
            ref.handle.click(timeout=1500)
            time.sleep(0.3)
            return
        except PWError:
            pass
    if ref.index is not None:
        try:
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
                ref.index,
            )
            time.sleep(0.3)
            return
        except PWError:
            pass
    _ativar_aba_conselheiro(scope, ref.key)


def _ativar_aba_conselheiro(scope, nome_upper: str) -> None:
    target = normalize_text(nome_upper)
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
                txt = _clean_tab_label(txt)
            except PWError:
                continue
            if normalize_text(txt) == target:
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
                    t = _clean_tab_label(t)
                except PWError:
                    continue
                arr.append(normalize_text(t))
        if target in arr:
            idx = arr.index(target)
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
    if destino.exists():
        try:
            destino.unlink()
        except Exception:
            pass
    try:
        tmp = d.path()
    except PlayTimeout:
        tmp = None
    if tmp:
        destino.write_bytes(Path(tmp).read_bytes())
    else:
        d.save_as(str(destino))


def _download_valido(path: Path) -> bool:
    try:
        return path.exists() and path.is_file() and path.stat().st_size > 0
    except Exception:
        return False


def _log_dir(base_dir: Path | None = None) -> Path:
    base = base_dir if base_dir is not None else (Path.cwd() / "logs")
    p = base / "_logs"
    p.mkdir(parents=True, exist_ok=True)
    return p


def _save_error_screenshot(scope, nome_norm: str, suffix: str, download_dir: Path | None = None) -> None:
    try:
        owner_page: Page = getattr(scope, "page", None) or scope
        path = _log_dir(download_dir) / f"erro_{nome_norm}_{suffix}.png"
        owner_page.screenshot(path=str(path), full_page=True)
    except Exception:
        pass


def _wait_grid_ready(scope, timeout_ms: int = 20000) -> None:
    selectors = [
        "table[id*='gv']",
        "table[id*='gvProtocolos']",
        "table[id*='gvConsulta']",
        ".dxgvControl",
    ]
    for sel in selectors:
        try:
            scope.wait_for_selector(sel, timeout=timeout_ms)
            break
        except PWError:
            continue

    for sel in [
        ".dxgvLoadingPanel",
        ".dxlpLoadingPanel",
        ".dxgvLoadingPanel_Material",
        ".dxlpLoadingDiv",
    ]:
        loc = scope.locator(sel)
        if loc.count() == 0:
            continue
        try:
            loc.first.wait_for(state="hidden", timeout=timeout_ms)
        except PWError:
            pass
    try:
        scope.wait_for_load_state("networkidle")
    except Exception:
        pass
    time.sleep(0.3)


def _extrair_evidencias_tc(scope, limit: int = 5) -> list[str]:
    try:
        text = scope.inner_text("body", timeout=3000)
    except Exception:
        return []
    encontrados = re.findall(r"TC/\\d{6}/\\d{4}", text)
    vistos = []
    for tc in encontrados:
        if tc not in vistos:
            vistos.append(tc)
        if len(vistos) >= limit:
            break
    return vistos


def _limpar_df_processos(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    out = df.copy()
    proc_col = None
    for col in out.columns:
        if "PROCESSO" in normalize_text(col):
            proc_col = col
            break

    if proc_col is not None:
        serie = out[proc_col].fillna("").astype(str).map(_ws)
        mask_tc = serie.str.contains(r"TC/\d+/\d{4}", case=False, regex=True, na=False)
        if bool(mask_tc.any()):
            out = out[mask_tc].copy()
        else:
            out = out[serie != ""].copy()
    else:
        out = out.dropna(how="all").copy()

    cols_drop = [c for c in out.columns if str(c).strip().upper().startswith("UNNAMED:")]
    if cols_drop:
        out = out.drop(columns=cols_drop, errors="ignore")
    return out


def _contar_itens_planilha(path: Path) -> int:
    try:
        df = pd.read_excel(path, dtype=str)
    except Exception:
        return 0
    df = _limpar_df_processos(df)
    return int(len(df.index))


def _split_planilha_por_relator(
    temp_path: Path,
    output_dir: Path,
    fallback_label: str,
    evidencias: list[str],
    prefixo: str = "PLENARIO",
) -> list[PlanilhaStats]:
    stats: list[PlanilhaStats] = []
    try:
        df = pd.read_excel(temp_path, dtype=str)
    except Exception:
        df = pd.DataFrame()
    df = _limpar_df_processos(df)

    if df.empty:
        fallback_name = _expand_relator_name(fallback_label) or _ws(fallback_label) or "DESCONHECIDO"
        final_name = f"{prefixo}_{normalize_text(fallback_name)}.xlsx"
        dest = output_dir / final_name
        if dest.exists():
            try:
                dest.unlink()
            except Exception:
                pass
        df.to_excel(dest, index=False)
        try:
            temp_path.unlink()
        except Exception:
            pass
        stats.append(
            PlanilhaStats(
                conselheiro=fallback_name,
                conselheiro_norm=normalize_text(fallback_name),
                path=dest,
                itens=0,
                tamanho_bytes=dest.stat().st_size if dest.exists() else 0,
                evidencias=evidencias,
                prefixo=prefixo,
            )
        )
        return stats

    # tenta localizar coluna Relator
    rel_col = None
    for col in df.columns:
        if "RELATOR" in normalize_text(col):
            rel_col = col
            break

    if not rel_col:
        fallback_name = _expand_relator_name(fallback_label) or _ws(fallback_label) or "DESCONHECIDO"
        final_name = f"{prefixo}_{normalize_text(fallback_name)}.xlsx"
        dest = output_dir / final_name
        if dest.exists():
            try:
                dest.unlink()
            except Exception:
                pass
        df.to_excel(dest, index=False)
        try:
            temp_path.unlink()
        except Exception:
            pass
        stats.append(
            PlanilhaStats(
                conselheiro=fallback_name,
                conselheiro_norm=normalize_text(fallback_name),
                path=dest,
                itens=len(df.index),
                tamanho_bytes=dest.stat().st_size if dest.exists() else 0,
                evidencias=evidencias,
                prefixo=prefixo,
            )
        )
        return stats

    rel_series = df[rel_col].dropna().astype(str).map(_expand_relator_name)
    df["_RelatorExpanded"] = rel_series
    valores = [v for v in rel_series.tolist() if v.strip()]
    relatores: list[str] = []
    for v in valores:
        if v not in relatores:
            relatores.append(v)

    if len(relatores) <= 1:
        relator_nome = relatores[0] if relatores else (_expand_relator_name(fallback_label) or _ws(fallback_label) or "DESCONHECIDO")
        nome_norm = normalize_text(relator_nome)
        final_name = f"{prefixo}_{nome_norm}.xlsx"
        dest = output_dir / final_name
        if dest.exists():
            try:
                dest.unlink()
            except Exception:
                pass
        df.drop(columns=["_RelatorExpanded"], errors="ignore").to_excel(dest, index=False)
        try:
            temp_path.unlink()
        except Exception:
            pass
        stats.append(
            PlanilhaStats(
                conselheiro=relator_nome,
                conselheiro_norm=nome_norm,
                path=dest,
                itens=len(df.index),
                tamanho_bytes=dest.stat().st_size if dest.exists() else 0,
                evidencias=evidencias,
                prefixo=prefixo,
            )
        )
        return stats

    # separa por relator
    for relator_nome in relatores:
        nome_norm = normalize_text(relator_nome)
        subset = df[df["_RelatorExpanded"] == relator_nome].drop(columns=["_RelatorExpanded"], errors="ignore")
        dest = output_dir / f"{prefixo}_{nome_norm}.xlsx"
        if dest.exists():
            try:
                dest.unlink()
            except Exception:
                pass
        subset.to_excel(dest, index=False)
        stats.append(
            PlanilhaStats(
                conselheiro=relator_nome,
                conselheiro_norm=nome_norm,
                path=dest,
                itens=len(subset.index),
                tamanho_bytes=dest.stat().st_size if dest.exists() else 0,
                evidencias=evidencias,
                prefixo=prefixo,
            )
        )

    try:
        temp_path.unlink()
    except Exception:
        pass
    return stats


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
    ano: str | None = None,
    max_retries: int = 3,
    competencia: str | None = None,
    competencias: list[str] | None = None,
    on_after_download: Callable[[PlanilhaStats], None] | None = None,
) -> List[PlanilhaStats]:
    """
    1) Abrir pesquisa e filtrar (consultarSessoesParaGabinete.aspx).
    2) Clicar no botao 'ConsultaDaPauta' da linha {num_sessao}/2025.
    3) Capturar o iframe do popup (processosDaPautaPorGabinete.aspx).
    4) Dentro do iframe, iterar abas de conselheiros e exportar Excel.
    """
    download_path = Path(download_dir)
    download_path.mkdir(parents=True, exist_ok=True)
    ano_final = ano or _guess_year(data_ate, data_de)

    # Pagina de pesquisa
    _goto_pagina_pauta(page, base_url)
    _preencher_filtros(page, num_sessao, data_de, data_ate)
    if not _clicar_pesquisar_robusto(page):
        raise RuntimeError("Nao consegui acionar a pesquisa (Pesquisar).")
    page.wait_for_load_state("networkidle")
    time.sleep(0.5)
    _wait_grid_ready(page)

    baixados: List[PlanilhaStats] = []
    relatores_encontrados: list[str] = []
    joao_presente = False

    competencias_exec: list[str | None] = []
    if competencias:
        competencias_exec = list(competencias)
    elif competencia:
        competencias_exec = [competencia]
    else:
        competencias_exec = [None]

    for comp in competencias_exec:
        comp_label = comp or "AUTO"
        prefixo = _prefix_from_competencia(comp)

        try:
            esperado = _clicar_botao_consulta_da_pauta(page, num_sessao, ano_final, comp)
        except RuntimeError as e:
            if len(competencias_exec) > 1:
                print(f"[WARN] Linha da competencia {comp_label} nao encontrada: {e}")
                continue
            raise

        if esperado.por_sigla:
            print(f"[CHECK] Esperado em tela ({comp_label}): {_format_esperados(esperado)}")
        else:
            print(f"[WARN] Nao foi possivel capturar contagens esperadas da grid ({comp_label}).")

        # Remove planilhas antigas desta competencia para garantir sobrescrita total.
        _limpar_planilhas_por_prefixo(download_path, prefixo)

        try:
            sonp_frame = _esperar_iframe_sonp(page, timeout_ms=35000)
        except RuntimeError:
            # Fallback: operate on current page (no iframe)
            print("[info] Iframe da SONP nao detectado; operando na pagina atual.")
            sonp_frame = page

        # Lista dinamica de conselheiros
        tab_refs: list[TabRef] = []
        t0 = time.time()
        while (time.time() - t0) < 20.0:
            try:
                tab_refs = _map_conselheiro_tabs(sonp_frame, allow_empty=True)
            except Exception:
                tab_refs = []
            if tab_refs:
                break
            time.sleep(0.5)

        if not tab_refs:
            _save_error_screenshot(sonp_frame, "tabs", "nao_carregadas", download_path)
            raise RuntimeError("Abas de conselheiro nao carregaram.")

        for tab_idx, ref in enumerate(tab_refs):
            nome = _clean_tab_label(ref.label or ref.key)
            nome_norm = normalize_text(nome)

            success = False
            for attempt in range(1, max_retries + 1):
                try:
                    print(f"[INICIO] {nome} ({comp_label}) (tentativa {attempt}/{max_retries})")
                    try:
                        _ativar_aba_por_ref(sonp_frame, ref)
                    except RuntimeError:
                        if tab_idx == 0:
                            print(f"[WARN] Nao foi possivel clicar na primeira aba ({nome}); seguindo com a aba atual.")
                        else:
                            raise
                    print(f"[ABA OK] {nome}")
                    _wait_grid_ready(sonp_frame)
                    print(f"[GRID OK] {nome}")

                    relatores = _collect_relatores(sonp_frame)
                    for r in relatores:
                        if r not in relatores_encontrados:
                            relatores_encontrados.append(r)
                    if any(("JOAO" in normalize_text(r) and "ANTONIO" in normalize_text(r)) for r in relatores):
                        joao_presente = True

                    evidencias = _extrair_evidencias_tc(sonp_frame)

                    temp_name = f"_TMP_{ref.key}_{attempt}.xlsx"
                    temp_path = download_path / temp_name
                    if temp_path.exists():
                        try:
                            temp_path.unlink()
                        except Exception:
                            pass

                    d = _clicar_exportar_excel(sonp_frame)
                    if not d:
                        time.sleep(0.8)
                        d = _clicar_exportar_excel(sonp_frame)
                    if not d:
                        raise RuntimeError("Nao encontrei botao de exportar Excel.")
                    print(f"[EXPORT OK] {nome}")

                    _salvar_download(d, temp_path)
                    if not _download_valido(temp_path):
                        raise RuntimeError("Arquivo baixado esta vazio ou inexistente.")

                    fallback_label = ref.label or ref.key
                    stats_list = _split_planilha_por_relator(
                        temp_path,
                        download_path,
                        fallback_label,
                        evidencias,
                        prefixo=prefixo,
                    )

                    for stats in stats_list:
                        if stats.conselheiro not in relatores_encontrados:
                            relatores_encontrados.append(stats.conselheiro)
                        if "JOAO" in stats.conselheiro_norm and "ANTONIO" in stats.conselheiro_norm:
                            joao_presente = True
                        print(f"[SALVO OK] {stats.conselheiro} -> QTD_ITENS={stats.itens} BYTES={stats.tamanho_bytes}")
                        if stats.itens == 0 and ("JOAO" in stats.conselheiro_norm and "ANTONIO" in stats.conselheiro_norm) and evidencias:
                            _save_error_screenshot(sonp_frame, stats.conselheiro_norm, "planilha_vazia", download_path)
                            raise RuntimeError(
                                "Planilha do Joao Antonio vazia, mas ha evidencia de processos: "
                                + ", ".join(evidencias)
                            )
                        elif stats.itens == 0:
                            print(f"[WARN] Planilha vazia para {stats.conselheiro}")

                        baixados.append(stats)

                        if on_after_download:
                            try:
                                on_after_download(stats)
                            except Exception as e:
                                _save_error_screenshot(sonp_frame, stats.conselheiro_norm, "docx_erro", download_path)
                                raise RuntimeError(f"Falha ao gerar DOCX apos download de {stats.conselheiro}: {e}")
                    success = True
                    break
                except Exception as e:
                    print(f"[WARN] Falha ao baixar {nome}: {e}")
                    time.sleep(0.8)

            if not success:
                _save_error_screenshot(sonp_frame, nome_norm, "falha_download", download_path)
                raise RuntimeError(f"Falha permanente ao baixar {nome} apos {max_retries} tentativas.")

        try:
            _validar_qtd_por_conselheiro(download_path, prefixo, esperado)
            print(f"[CHECK OK] Quantidade por conselheiro validada ({comp_label}).")
        except Exception as e:
            _save_error_screenshot(sonp_frame, normalize_text(comp_label), "divergencia_qtd", download_path)
            raise RuntimeError(f"Falha na validacao de quantidade por conselheiro ({comp_label}): {e}")

        _fechar_popup_sonp(page)
        try:
            _wait_grid_ready(page)
        except Exception:
            pass

    joao_baixado = any(("JOAO" in item.conselheiro_norm and "ANTONIO" in item.conselheiro_norm) for item in baixados)
    if joao_presente and not joao_baixado:
        raise RuntimeError("Joao Antonio aparece nos relatores, mas nao foi baixado.")

    if not baixados:
        print("[info] Nenhuma planilha foi baixada (talvez nao haja processos).")
    else:
        print("[RESUMO] Planilhas baixadas:")
        for item in baixados:
            print(
                f"  - {item.conselheiro}: QTD_ITENS={item.itens} BYTES={item.tamanho_bytes} ARQ={item.path.name}"
            )

    return baixados
