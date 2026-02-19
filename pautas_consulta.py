from __future__ import annotations

import re
import time
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence

import pandas as pd
from playwright.sync_api import Page, Download, Error as PWError, TimeoutError as PlayTimeout, sync_playwright

from login import efetuar_login


@dataclass(frozen=True)
class PlanilhaCompetencia:
    key: str
    label: str
    slug: str
    sheet: str
    path: Path


_COMPETENCIA_INFO = {
    "pleno": {
        "label": "Pleno",
        "slug": "PLENO",
        "sheet": "Pleno",
    },
    "1c": {
        "label": "1\u00aa C\u00e2mara",
        "slug": "1CAMARA",
        "sheet": "1a_Camara",
    },
    "2c": {
        "label": "2\u00aa C\u00e2mara",
        "slug": "2CAMARA",
        "sheet": "2a_Camara",
    },
}


def _strip_accents(text: str) -> str:
    return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")


def _norm_text(text: str) -> str:
    text = _strip_accents(str(text))
    text = re.sub(r"\s+", " ", text).strip().upper()
    return text


def _slug_date(value: str) -> str:
    s = re.sub(r"[^0-9]+", "_", str(value)).strip("_")
    return s or "SEM_DATA"


def _resolve_competencias(raw: Sequence[str] | None) -> list[str]:
    if not raw:
        return ["pleno", "1c", "2c"]
    out: list[str] = []
    for item in raw:
        key = _normalize_competencia_key(item)
        if key not in out:
            out.append(key)
    return out


def _normalize_competencia_key(value: str) -> str:
    v = _norm_text(value)
    if v in {"PLENO", "PLENARIO", "PLENARIO"}:
        return "pleno"
    if "CAMARA" in v and "1" in v:
        return "1c"
    if "CAMARA" in v and "2" in v:
        return "2c"
    if v in {"1C", "1A", "1"}:
        return "1c"
    if v in {"2C", "2A", "2"}:
        return "2c"
    raise ValueError(f"Competencia invalida: {value}")


def _goto_consulta_pautas(page: Page, base_url: str) -> None:
    url = f"{base_url}/paginas/resultado/consultarSessoesParaGabinete.aspx"
    page.goto(url, wait_until="domcontentloaded")
    page.wait_for_load_state("networkidle")


def _fill_masked_input(page_or_frame, selectors: Iterable[str], value: str | None) -> bool:
    if value is None:
        return False
    for sel in selectors:
        try:
            loc = page_or_frame.locator(sel)
            if loc.count() == 0:
                continue
            el = loc.first
            el.click(timeout=2500)
            try:
                el.press("Control+A")
                el.press("Delete")
            except PWError:
                pass
            el.fill(value, timeout=2500)
            page_or_frame.keyboard.press("Tab")
            return True
        except PWError:
            continue
    return False


def _clear_input_if_present(page_or_frame, selectors: Iterable[str]) -> None:
    for sel in selectors:
        try:
            loc = page_or_frame.locator(sel)
            if loc.count() == 0:
                continue
            el = loc.first
            el.click(timeout=2000)
            try:
                el.press("Control+A")
                el.press("Delete")
            except PWError:
                pass
            el.fill("", timeout=2000)
            page_or_frame.keyboard.press("Tab")
            return
        except PWError:
            continue


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
            loc.first.click(timeout=2000)
            return True
        except PWError:
            try:
                loc.first.click(timeout=2000, force=True)
                return True
            except PWError:
                pass
    try:
        scope.keyboard.press("Enter")
        return True
    except PWError:
        return False


def _wait_grid_ready(page: Page, timeout_ms: int = 20000) -> None:
    try:
        page.wait_for_selector("table[id*='gvConsulta'], div[id*='gvConsulta']", timeout=timeout_ms)
    except PWError:
        pass
    for sel in [
        "div[id*='gvConsulta_LPV']",
        ".dxgvLoadingPanel",
        ".dxlpLoadingPanel",
        ".dxgvLoadingPanel_Material",
    ]:
        loc = page.locator(sel)
        if loc.count() == 0:
            continue
        try:
            loc.first.wait_for(state="hidden", timeout=timeout_ms)
        except PWError:
            pass
    page.wait_for_load_state("networkidle")
    time.sleep(0.4)


def _get_grid_headers(page: Page) -> list[str]:
    selectors = [
        "table[id*='gvConsulta'] th",
        "table[id*='gvConsulta'] td.dxgvHeader",
        "tr.dxgvHeaderRow td",
        "tr.dxgvHeaderRow th",
    ]
    headers: list[str] = []
    for sel in selectors:
        loc = page.locator(sel)
        try:
            n = loc.count()
        except PWError:
            n = 0
        if n == 0:
            continue
        for i in range(n):
            try:
                txt = loc.nth(i).inner_text(timeout=1000)
            except PWError:
                continue
            txt = re.sub(r"\s+", " ", txt or "").strip()
            if txt and txt not in headers:
                headers.append(txt)
        if headers:
            break
    return headers


def _open_competencia_dropdown(page: Page):
    input_selectors = [
        "#cbCompetencia_I",
        "#ddlCompetencia_I",
        "input[id*='Competencia'][id$='_I']",
        "input[name*='Competencia']",
        "input[aria-label*='Compet']",
        "input[placeholder*='Compet']",
    ]
    input_loc = None
    for sel in input_selectors:
        loc = page.locator(sel)
        if loc.count() > 0:
            input_loc = loc.first
            break
    if input_loc is None:
        try:
            label = page.locator("text=Compet\u00eancia").first
            if label.count() > 0:
                input_loc = label.locator("xpath=following::input[1]").first
        except PWError:
            input_loc = None
    if input_loc is None or input_loc.count() == 0:
        raise RuntimeError("Nao encontrei o campo de Competencia.")

    try:
        input_loc.click(timeout=2000)
    except PWError:
        input_loc.click(timeout=2000, force=True)

    try:
        input_id = input_loc.get_attribute("id")
    except PWError:
        input_id = None
    if input_id:
        btn_sel = f"#{input_id[:-2]}_B" if input_id.endswith("_I") else f"#{input_id}_B"
        btn = page.locator(btn_sel)
        if btn.count() > 0:
            try:
                btn.first.click(timeout=1500)
            except PWError:
                pass
    return input_loc


def _clear_competencia_selections(page: Page) -> None:
    for sel in [
        "a:has-text('Limpar')",
        "a:has-text('Nenhum')",
        "a:has-text('Desmarcar')",
        "button:has-text('Limpar')",
    ]:
        loc = page.locator(sel)
        if loc.count() == 0:
            continue
        try:
            loc.first.click(timeout=1000)
            time.sleep(0.2)
            return
        except PWError:
            pass

    for sel in [
        ".dxeListBoxItemSelected",
        ".dxeListBoxItemChecked",
        ".dxlbSelected",
    ]:
        loc = page.locator(sel)
        try:
            n = loc.count()
        except PWError:
            n = 0
        for i in range(n):
            try:
                loc.nth(i).click(timeout=800)
            except PWError:
                pass


def _click_option_by_text(page: Page, label: str) -> bool:
    label_norm = _norm_text(label)
    scope = page
    dropdown_root = _find_visible_dropdown_root(page)
    if dropdown_root is not None:
        scope = dropdown_root
    selectors = [
        "div[role='listbox'] [role='option']",
        "ul[role='listbox'] li",
        "div[id*='Competencia_DDD'] tr td",
        "table[id*='Competencia_DDD'] tr td",
        ".dxeListBoxItem",
        ".dxmMenuItem",
    ]
    for sel in selectors:
        loc = scope.locator(sel)
        try:
            n = loc.count()
        except PWError:
            n = 0
        for i in range(n):
            try:
                txt = loc.nth(i).inner_text(timeout=800).strip()
            except PWError:
                continue
            if not txt:
                continue
            txt_norm = _norm_text(txt)
            if txt_norm == label_norm or label_norm in txt_norm:
                try:
                    loc.nth(i).click(timeout=1200)
                    return True
                except PWError:
                    pass
    if _try_click_checkbox_option(scope, label):
        return True
    return False


def _set_competencia(page: Page, label: str) -> None:
    if _try_select_native_competencia(page, label):
        return
    input_loc = _open_competencia_dropdown(page)
    _clear_competencia_selections(page)
    if not _click_option_by_text(page, label):
        if _try_set_competencia_js(page, label):
            return
        raise RuntimeError(f"Nao consegui selecionar a competencia: {label}")
    try:
        page.keyboard.press("Enter")
    except PWError:
        pass
    time.sleep(0.2)

    try:
        current = input_loc.input_value(timeout=1500)
    except PWError:
        current = ""
    if current and _norm_text(label) not in _norm_text(current):
        print(f"[warn] Campo Competencia nao refletiu '{label}': '{current}'.")


def _find_visible_dropdown_root(page: Page):
    selectors = [
        "div[id$='_DDD']:visible",
        "div[id*='Competencia_DDD']:visible",
        "div.dxeDropDownWindow:visible",
        "div.dxpcDropDown:visible",
        "div.dxpc-content:visible",
        "div.dxeListBox:visible",
    ]
    for sel in selectors:
        loc = page.locator(sel)
        try:
            if loc.count() > 0:
                return loc.first
        except PWError:
            continue
    return None


def _try_select_native_competencia(page: Page, label: str) -> bool:
    selectors = [
        "select[id*='Competencia']",
        "select[name*='Competencia']",
    ]
    for sel in selectors:
        loc = page.locator(sel)
        if loc.count() == 0:
            continue
        try:
            loc.first.select_option(label=label)
            return True
        except PWError:
            try:
                loc.first.select_option(value=label)
                return True
            except PWError:
                pass
    return False


def _try_click_checkbox_option(scope, label: str) -> bool:
    label_norm = _norm_text(label)
    text_loc = scope.locator(f"text={label}")
    try:
        n = text_loc.count()
    except PWError:
        n = 0
    for i in range(n):
        try:
            el = text_loc.nth(i)
        except PWError:
            continue
        try:
            txt = el.inner_text(timeout=500).strip()
        except PWError:
            txt = ""
        if txt and label_norm not in _norm_text(txt):
            continue
        for xpath in [
            "xpath=ancestor::tr[1]//input[@type='checkbox' or @type='radio']",
            "xpath=ancestor::li[1]//input[@type='checkbox' or @type='radio']",
            "xpath=ancestor::div[1]//input[@type='checkbox' or @type='radio']",
        ]:
            inp = el.locator(xpath)
            if inp.count() == 0:
                continue
            try:
                inp.first.click(timeout=800)
                return True
            except PWError:
                pass
    return False


def _try_set_competencia_js(page: Page, label: str) -> bool:
    try:
        return bool(page.evaluate(
            """(label) => {
                const norm = (s) => {
                  try {
                    return (s || "").toString().normalize('NFD').replace(/[\\u0300-\\u036f]/g, '')
                      .replace(/\\s+/g, ' ').trim().toUpperCase();
                  } catch(e) {
                    return (s || "").toString().toUpperCase();
                  }
                };
                const target = norm(label);

                const tryListBox = (lb, cb) => {
                  try {
                    if (!lb) return false;
                    if (lb.UnselectAll) lb.UnselectAll();
                    const n = lb.GetItemCount ? lb.GetItemCount() : 0;
                    for (let j = 0; j < n; j++) {
                      let text = "";
                      try {
                        text = lb.GetItem ? (lb.GetItem(j)?.text || "") : (lb.GetItemText ? lb.GetItemText(j) : "");
                      } catch(e) {}
                      if (!text) continue;
                      if (norm(text).includes(target)) {
                        if (lb.SelectItem) lb.SelectItem(j);
                        else if (lb.SelectIndices) lb.SelectIndices([j]);
                        else if (lb.SetSelectedIndex) lb.SetSelectedIndex(j);
                        if (cb?.SetText && window.getSelectedItemsText && lb.GetSelectedItems) {
                          cb.SetText(window.getSelectedItemsText(lb.GetSelectedItems()));
                        }
                        return true;
                      }
                    }
                  } catch(e) {}
                  return false;
                };

                const lbGlobal = window.lbCompetencia;
                const cbGlobal = window.cbCompetencia;
                if (tryListBox(lbGlobal, cbGlobal)) return true;

                const col = window.ASPxClientControl?.GetControlCollection?.();
                if (!col) return false;
                const count = col.GetCount ? col.GetCount() : 0;
                for (let i = 0; i < count; i++) {
                  const c = col.Get(i);
                  const name = c?.name || (c?.GetName ? c.GetName() : "");
                  if (name === "lbCompetencia") {
                    if (tryListBox(c, cbGlobal)) return true;
                  }
                  if (!name || !name.toLowerCase().includes('compet')) continue;
                  try {
                    if (c.UnselectAll) c.UnselectAll();
                    if (c.SelectAll && c.UnselectAll) c.UnselectAll();
                  } catch(e) {}
                  try {
                    if (c.GetItemCount) {
                      const n = c.GetItemCount();
                      for (let j = 0; j < n; j++) {
                        let text = "";
                        try {
                          text = c.GetItem ? (c.GetItem(j)?.text || "") : (c.GetItemText ? c.GetItemText(j) : "");
                        } catch(e) {}
                        if (!text) continue;
                        if (norm(text).includes(target)) {
                          if (c.SelectItem) c.SelectItem(j);
                          else if (c.SetSelectedIndex) c.SetSelectedIndex(j);
                          else if (c.SetValue) c.SetValue(c.GetItem(j)?.value ?? text);
                          if (c.HideDropDown) c.HideDropDown();
                          return true;
                        }
                      }
                    }
                  } catch(e) {}
                }
                return false;
            }""",
            label,
        ))
    except PWError:
        return False


def _selecionar_situacao(page: Page, situacao: str) -> None:
    if not situacao:
        return
    situacao_norm = _norm_text(situacao)
    select_selectors = [
        "select[id*='Situacao']",
        "select[name*='Situacao']",
    ]
    for sel in select_selectors:
        loc = page.locator(sel)
        if loc.count() == 0:
            continue
        try:
            loc.first.select_option(label=situacao)
            return
        except PWError:
            pass

    input_selectors = [
        "input[id*='Situacao'][id$='_I']",
        "input[name*='Situacao']",
        "input[aria-label*='Situacao']",
        "input[placeholder*='Situacao']",
    ]
    input_loc = None
    for sel in input_selectors:
        loc = page.locator(sel)
        if loc.count() > 0:
            input_loc = loc.first
            break
    if input_loc is None:
        return
    try:
        input_loc.click(timeout=1500)
    except PWError:
        input_loc.click(timeout=1500, force=True)
    if not _click_option_by_text(page, situacao):
        raise RuntimeError(f"Nao consegui selecionar a situacao: {situacao}")
    try:
        page.keyboard.press("Enter")
    except PWError:
        pass
    time.sleep(0.2)


def _preencher_filtros_consulta(
    page: Page,
    data_de: str,
    data_ate: str,
    num_sessao: str | None,
    situacao: str | None,
) -> None:
    if num_sessao:
        _fill_masked_input(page, ["#spnNumSessao_I", "input[id*='spnNumSessao']", "input[name*='NumSessao']"], num_sessao)
    else:
        _clear_input_if_present(page, ["#spnNumSessao_I", "input[id*='spnNumSessao']", "input[name*='NumSessao']"])

    ok2 = _fill_masked_input(page, ["#dteInicial_I", "input[id*='dteInicial']"], data_de)
    ok3 = _fill_masked_input(page, ["#dteFinal_I", "input[id*='dteFinal']"], data_ate)
    if not (ok2 and ok3):
        raise RuntimeError("Nao foi possivel preencher datas (De/Ate).")

    if situacao:
        _selecionar_situacao(page, situacao)


def _clicar_exportar_excel(page: Page) -> Download | None:
    direct_selectors = [
        "a[title*='Excel' i]",
        "a:has-text('Exportar para Excel')",
        "button:has-text('Excel')",
        "li:has-text('Exportar para Excel')",
        "img[alt*='Excel' i]",
        "[id*='btnExport'][id*='Xls']",
        "[id*='btnExport'][id*='Xlsx']",
        "[id*='btnExportar'][id*='Excel']",
    ]
    for sel in direct_selectors:
        loc = page.locator(sel)
        if loc.count() == 0:
            continue
        try:
            with page.expect_download(timeout=25000) as ev:
                loc.first.click(timeout=3000)
            return ev.value
        except (PWError, PlayTimeout):
            continue

    menu_selectors = [
        "a:has-text('Exportar')",
        "button:has-text('Exportar')",
        "span:has-text('Exportar')",
        "[title*='Exportar']",
        "[id*='btnExportar']",
    ]
    for sel in menu_selectors:
        loc = page.locator(sel)
        if loc.count() == 0:
            continue
        try:
            loc.first.click(timeout=2000)
        except PWError:
            try:
                loc.first.click(timeout=2000, force=True)
            except PWError:
                continue

        for excel_sel in [
            "li:has-text('Exportar para Excel')",
            "a:has-text('Exportar para Excel')",
            "li:has-text('Excel')",
            "a:has-text('Excel')",
            "span:has-text('Excel')",
        ]:
            ex = page.locator(excel_sel)
            if ex.count() == 0:
                continue
            try:
                with page.expect_download(timeout=25000) as ev:
                    ex.first.click(timeout=3000)
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


def _is_probably_xlsx(path: Path) -> bool:
    try:
        with open(path, "rb") as fh:
            sig = fh.read(4)
        return sig[:2] == b"PK" or sig.startswith(b"\xD0\xCF\x11\xE0")
    except Exception:
        return False


def _is_html_file(path: Path) -> bool:
    try:
        with open(path, "rb") as fh:
            head = fh.read(256).lstrip()
        return head.startswith(b"<")
    except Exception:
        return False


def _write_empty_excel(path: Path, headers: Sequence[str]) -> None:
    df = pd.DataFrame(columns=list(headers))
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(path, index=False)


def baixar_planilhas_consulta_pautas(
    page: Page,
    base_url: str,
    data_de: str,
    data_ate: str,
    download_dir: str,
    competencias: Sequence[str] | None = None,
    num_sessao: str | None = None,
    situacao: str | None = None,
) -> list[PlanilhaCompetencia]:
    Path(download_dir).mkdir(parents=True, exist_ok=True)
    comp_keys = _resolve_competencias(competencias)
    planilhas: list[PlanilhaCompetencia] = []

    for key in comp_keys:
        info = _COMPETENCIA_INFO[key]
        _goto_consulta_pautas(page, base_url)
        _preencher_filtros_consulta(page, data_de, data_ate, num_sessao, situacao)
        _set_competencia(page, info["label"])

        if not _clicar_pesquisar_robusto(page):
            raise RuntimeError("Nao consegui acionar a pesquisa (Pesquisar).")
        _wait_grid_ready(page)

        fname = f"pautas_{info['slug']}_{_slug_date(data_de)}_{_slug_date(data_ate)}.xlsx"
        dest = Path(download_dir) / fname

        download = _clicar_exportar_excel(page)
        if not download:
            headers = _get_grid_headers(page)
            _write_empty_excel(dest, headers)
            print(f"[warn] Exportacao nao encontrada para {info['label']}. Arquivo vazio gerado.")
        else:
            try:
                _salvar_download(download, dest)
                if not _is_probably_xlsx(dest):
                    print(f"[warn] Arquivo baixado nao parece XLSX: {dest.name}")
            except Exception as e:
                headers = _get_grid_headers(page)
                _write_empty_excel(dest, headers)
                print(f"[warn] Falha ao salvar download de {info['label']}: {e}")

        planilhas.append(
            PlanilhaCompetencia(
                key=key,
                label=info["label"],
                slug=info["slug"],
                sheet=info["sheet"],
                path=dest,
            )
        )

    return planilhas


def _fix_mojibake(value: str) -> str:
    if not value:
        return value
    if "\u00c3" not in value and "\u00c2" not in value:
        return value
    try:
        fixed = value.encode("latin1").decode("utf-8")
    except Exception:
        return value
    return fixed


def _normalize_column_names(columns: Sequence[str]) -> list[str]:
    out: list[str] = []
    seen: dict[str, int] = {}
    for col in columns:
        txt = _fix_mojibake(str(col))
        txt = txt.replace("\r", " ").replace("\n", " ")
        txt = re.sub(r"\s+", " ", txt).strip()
        if not txt:
            txt = "Coluna"
        base = txt
        idx = seen.get(base, 0)
        if idx:
            txt = f"{base}_{idx + 1}"
        seen[base] = idx + 1
        out.append(txt)
    return out


def _col_key(name: str) -> str:
    name = _strip_accents(str(name))
    name = re.sub(r"[^A-Za-z0-9]+", "", name).upper()
    return name


def _coerce_types(df: pd.DataFrame) -> pd.DataFrame:
    numeric_keys = {"TOTAL", "DD", "JA", "RB", "ET", "RT"}
    date_keys = {"DATA", "DATASESSAO", "DATASESSAOPLENARIA"}
    for col in df.columns:
        key = _col_key(col)
        if key in numeric_keys:
            series = df[col].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
            df[col] = pd.to_numeric(series, errors="coerce")
        if key in date_keys:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
    return df


def _ensure_competencia_col(df: pd.DataFrame, label: str) -> pd.DataFrame:
    comp_keys = {"COMPETENCIA", "COMPETENCIAS"}
    comp_cols = [c for c in df.columns if _col_key(c) in comp_keys]
    if comp_cols:
        df = df.rename(columns={comp_cols[0]: "Compet\u00eancia"})
        if len(comp_cols) > 1:
            df = df.drop(columns=comp_cols[1:])
    df["Compet\u00eancia"] = label
    return df


def _read_planilha(path: Path) -> pd.DataFrame:
    if not path.exists() or path.stat().st_size == 0:
        return pd.DataFrame()
    if _is_probably_xlsx(path):
        try:
            df = pd.read_excel(path, dtype=str)
        except Exception as e:
            print(f"[warn] Falha ao ler XLSX {path.name}: {e}")
            return pd.DataFrame()
    elif _is_html_file(path):
        try:
            tables = pd.read_html(path)
            df = tables[0] if tables else pd.DataFrame()
        except Exception as e:
            print(f"[warn] Falha ao ler HTML {path.name}: {e}")
            return pd.DataFrame()
    else:
        print(f"[warn] Formato desconhecido: {path.name}")
        return pd.DataFrame()

    df.columns = _normalize_column_names(df.columns)
    return df


def _dedupe_frame(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    id_keys = {"CODSESSAO", "IDSESSAO", "CODIGOSESSAO", "ID"}
    id_col = None
    for col in df.columns:
        if _col_key(col) in id_keys:
            id_col = col
            break
    if id_col:
        return df.drop_duplicates(subset=[id_col, "Compet\u00eancia"], keep="first")

    col_map = { _col_key(c): c for c in df.columns }
    required_keys = {
        "NSESSAOPLENARIA",
        "NUMEROSESSAOPLENARIA",
        "NODASESSAOPLENARIA",
        "NSESSAO",
        "NUMEROSESSAO",
        "NODASESSAO",
    }
    sessao_col = None
    for k in required_keys:
        if k in col_map:
            sessao_col = col_map[k]
            break
    if not sessao_col:
        for col in df.columns:
            key = _col_key(col)
            if "SESSAO" in key and ("NUMERO" in key or key.startswith("N")):
                sessao_col = col
                break
    data_col = col_map.get("DATA")
    tipo_col = col_map.get("TIPODESESSAO")
    formato_col = col_map.get("FORMATO")

    if sessao_col and data_col and tipo_col and formato_col:
        subset = [sessao_col, data_col, "Compet\u00eancia", tipo_col, formato_col]
        return df.drop_duplicates(subset=subset, keep="first")
    return df


def consolidar_planilhas_competencias(
    planilhas: Sequence[PlanilhaCompetencia],
    output_path: str | Path,
    include_competencia_sheets: bool = True,
    include_resumo: bool = False,
    dedupe: bool = False,
    periodo: str | None = None,
) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    frames: list[pd.DataFrame] = []
    por_comp: dict[str, pd.DataFrame] = {}
    counts: dict[str, int] = {}

    for item in planilhas:
        df = _read_planilha(item.path)
        if not df.empty:
            df = _coerce_types(df)
        df = _ensure_competencia_col(df, item.label)
        counts[item.key] = len(df)
        por_comp[item.key] = df
        frames.append(df)

    if frames:
        full = pd.concat(frames, ignore_index=True)
    else:
        full = pd.DataFrame()

    if dedupe:
        full = _dedupe_frame(full)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        full.to_excel(writer, sheet_name="Consolidado", index=False)
        if include_competencia_sheets:
            for item in planilhas:
                df_comp = por_comp.get(item.key, pd.DataFrame())
                df_comp.to_excel(writer, sheet_name=item.sheet, index=False)
        if include_resumo:
            resumo_rows = [
                {"Chave": "Periodo", "Valor": periodo or ""},
                {"Chave": "Gerado_em", "Valor": pd.Timestamp.now()},
                {"Chave": "Linhas_total", "Valor": len(full)},
            ]
            for item in planilhas:
                resumo_rows.append({"Chave": f"Linhas_{item.slug}", "Valor": counts.get(item.key, 0)})
            resumo = pd.DataFrame(resumo_rows)
            resumo.to_excel(writer, sheet_name="Resumo", index=False)

    return output_path


def run_consulta_pautas_pipeline(
    *,
    base_url: str,
    usuario: str,
    senha: str,
    data_de: str,
    data_ate: str,
    download_dir: str,
    output_dir: str,
    headless: bool = True,
    competencias: Sequence[str] | None = None,
    num_sessao: str | None = None,
    situacao: str | None = None,
    nome_consolidado: str | None = None,
    include_competencia_sheets: bool = True,
    include_resumo: bool = False,
    dedupe: bool = False,
) -> str:
    download_path = Path(download_dir)
    output_path = Path(output_dir)
    download_path.mkdir(parents=True, exist_ok=True)
    output_path.mkdir(parents=True, exist_ok=True)

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
            planilhas = baixar_planilhas_consulta_pautas(
                page=page,
                base_url=base_url,
                data_de=data_de,
                data_ate=data_ate,
                download_dir=str(download_path),
                competencias=competencias,
                num_sessao=num_sessao,
                situacao=situacao,
            )
        finally:
            context.close()
            browser.close()

    nome_final = nome_consolidado or f"pautas_CONSOLIDADO_{_slug_date(data_de)}_{_slug_date(data_ate)}.xlsx"
    saida = output_path / nome_final
    consolidar_planilhas_competencias(
        planilhas=planilhas,
        output_path=saida,
        include_competencia_sheets=include_competencia_sheets,
        include_resumo=include_resumo,
        dedupe=dedupe,
        periodo=f"{data_de} - {data_ate}",
    )
    print(f"Concluido: {saida}")
    return str(saida)
