from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime, timedelta, date
from pathlib import Path
from typing import Optional, Tuple, List

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.opc.exceptions import PackageNotFoundError


# =========================
# Utilidades básicas
# =========================

def _ws(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()


def _strip_accents_lower(s: str) -> str:
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.lower()


def _parse_date_br(s: str) -> datetime | None:
    try:
        return datetime.strptime(s.strip(), "%d/%m/%Y")
    except Exception:
        return None


def _fmt_date_br(d: datetime | date) -> str:
    return (d if isinstance(d, datetime) else datetime.combine(d, datetime.min.time())).strftime("%d/%m/%Y")


def _first_weekday_of_next_month(d: datetime, weekday: int) -> date:
    """weekday: Monday=0 .. Sunday=6. Returns date in next month."""
    year = d.year + (1 if d.month == 12 else 0)
    month = 1 if d.month == 12 else d.month + 1
    first = date(year, month, 1)
    delta = (weekday - first.weekday()) % 7
    return first + timedelta(days=delta)


def _nth_weekday_of_next_month(d: datetime, weekday: int, n: int) -> date:
    first = _first_weekday_of_next_month(d, weekday)
    return first + timedelta(days=7 * (n - 1))


def _next_weekday_strict(d: datetime, weekday: int) -> date:
    """Next weekday strictly after given date."""
    delta = (weekday - d.weekday()) % 7
    if delta == 0:
        delta = 7
    return (d + timedelta(days=delta)).date()


def _weekday_of_next_week(today: datetime, weekday: int) -> date:
    """Return the date of `weekday` (Mon=0..Sun=6) in the next ISO week from `today`."""
    # Monday of next week = today - weekday + 7
    monday_next = today.date() + timedelta(days=(7 - today.weekday()))
    return monday_next + timedelta(days=weekday)


def _cargo_conselheiro(nome: str) -> str:
    k = _strip_accents_lower(_ws(nome))
    if k == "domingos dissei":
        return "CONSELHEIRO PRESIDENTE"
    if k == "ricardo torres":
        return "CONSELHEIRO VICE-PRESIDENTE"
    if k == "roberto braguim":
        return "CONSELHEIRO CORREGEDOR"
    return "CONSELHEIRO"


def _detect_cols_basic(cols: List[str]) -> Tuple[int, int, Optional[int], Optional[int], Optional[int]]:
    """
    Retorna índices (0-based): (proc_idx, obj_idx, relator_idx, revisor_idx, motivo_idx)
    - Heurística por nome de coluna (processo/objeto/relator/revisor/motivo).
    - Fallbacks:
        * Relator  -> coluna 7 (index 6), se existir
        * Revisor  -> coluna 8 (index 7), se existir
        * Motivo   -> coluna 10 (index 9), se existir
    """
    proc_idx, obj_idx = 1, 3
    relator_idx, revisor_idx, motivo_idx = None, None, None

    for i, c in enumerate(cols):
        cl = c.lower()
        if any(k in cl for k in ["processo", "nº do processo", "numero do processo", "n. do processo", "nº do proc."]):
            proc_idx = i
        elif any(k in cl for k in ["objeto", "objeto de julgamento"]):
            obj_idx = i
        elif any(k in cl for k in ["relator", "conselheiro", "relator(a)"]):
            relator_idx = i
        elif "revisor" in cl or "revisor(a)" in cl:
            revisor_idx = i
        elif "motivo" in cl:
            motivo_idx = i

    n = len(cols)
    proc_idx = proc_idx if proc_idx < n else min(1, n - 1)
    obj_idx = obj_idx if obj_idx < n else min(3, n - 1)
    if relator_idx is None and n >= 7:
        relator_idx = 6
    if revisor_idx is None and n >= 8:
        revisor_idx = 7
    if motivo_idx is None and n >= 10:
        motivo_idx = 9
    return proc_idx, obj_idx, relator_idx, revisor_idx, motivo_idx


# --------- Mapeamento (iniciais → nome por extenso) ---------
_NAME_MAP = {
    "ET": "EDUARDO TUMA",
    "DD": "DOMINGOS DISSEI",
    "JA": "JOÃO ANTÔNIO",
    "RT": "RICARDO TORRES",
    "RB": "ROBERTO BRAGUIM",
}

# Duplas padrão Relator → Revisor
_DUO_RELATOR_REVISOR = {
    "ricardo torres": "ROBERTO BRAGUIM",
    "roberto braguim": "JOÃO ANTÔNIO",
    "joao antonio": "EDUARDO TUMA",
    "eduardo tuma": "RICARDO TORRES",
    "domingos dissei": "ROBERTO BRAGUIM",
}

# Palavras/expressões-chave a destacar (em negrito) no campo "Objeto".
_OBJ_HL_TERMS: list[str] = [
    # A) Recurso; B) Diversos; C) Contratos
    "Recurso",
    "Diversos",
    # Contratos - subtipos
    "Nota de Empenho com Termo Aditivo",
    "Nota de Empenho",
    "Análise de Licitação",
    "Lei 8.666",
    "Lei 14.133",
    "Ordem de Execução de Serviço",
    "Execução Contratual",
    "Carta-Contrato",
    "Contrato com Termo Aditivo",
    "Carta-Contrato com Termo Aditivo",
    "Consórcio",
    "Termo de Copatrocínio",
    "Convênio com Termo Aditivo",
    "Convênio com Termo de Rescisão",
    "Convênio",
    "Termo de Quitação",
    "Termo de Recebimento Definitivo",
    "Carta Aditiva",
    "Termo Aditivo com Nota de Empenho",
    "Termo Aditivo",
    "Termo de Prorrogação",
    "Termo de Recebimento Provisório",
    "Termo de Retirratificação com valor",
    "Termo de Retirratificação sem valor",
    "Termo de Compromisso e Autorização de Fornecimento",
    "Termo de Quitação de Obrigações",
    "Termo de Rescisão",
    "Termo Aditivo com Execução Contratual",
    "Autorização de Fornecimento",
    "Pedido de Compra",
    "Acompanhamentos",
    "Acompanhamento de edital",
    "Acompanhamento de procedimento licitatório",
    "Acompanhamento - Execução Contratual",
    "Acompanhamento - Execução Contábil e Financeira",
    "- Acompanhamento - Execução Contábil e Financeira",
    # D) E) F)
    "Contratos de Emergência",
    "Lei Municipal 11.100/91",
    "Subvenções",
    "Auxílios",
    "Reinclusões",
]


def _norm_term_label(term: str) -> str:
    t = _strip_accents_lower(_ws(term))
    t = t.replace("-", "-")
    t = re.sub(r"^[\s\-]+", "", t)
    t = re.sub(r"\s+", " ", t)
    return t


def _compile_keyword_patterns() -> list[tuple[str, re.Pattern]]:
    patterns: list[tuple[str, re.Pattern]] = []
    # Classe de hífens/traços Unicode comum: ASCII '-' e U+2010..U+2015
    hy = r"[-\u2010-\u2015]"
    for term in sorted(_OBJ_HL_TERMS, key=len, reverse=True):  # prioriza termos mais longos
        # Constrói regex tolerante a espaços múltiplos e variação de hífens
        esc = re.escape(term)
        esc = esc.replace(r"\ ", r"\s+")
        # Normaliza hífens/traços para classe comum
        esc = esc.replace(r"\-", hy)  # hífen ASCII escapado
        esc = esc.replace("\u2010", hy)  # HYPHEN (se escapar como literal)
        esc = esc.replace("\u2011", hy)  # NON-BREAKING HYPHEN
        esc = esc.replace("\u2012", hy)  # FIGURE DASH
        esc = esc.replace("\u2013", hy)  # EN DASH
        esc = esc.replace("\u2014", hy)  # EM DASH
        esc = esc.replace("\u2015", hy)  # HORIZONTAL BAR
        # aceita variações de ':' com ou sem espaço
        esc = esc.replace(r"\:\ ", r":?\s*")
        try:
            pat = re.compile(esc, flags=re.IGNORECASE)
            patterns.append((_norm_term_label(term), pat))
        except re.error:
            patterns.append((_norm_term_label(term), re.compile(re.escape(term), flags=re.IGNORECASE)))
    return patterns


_OBJ_HL_PATTERNS = _compile_keyword_patterns()


def _expand_initials(value: str) -> str:
    """
    Converte iniciais para nome por extenso:
    - 'ET', 'E.T.', ' et ' → 'EDUARDO TUMA'
    - Se já vier por extenso, normaliza (UPPER, trim).
    """
    s = _ws(value)
    if not s:
        return ""
    code = re.sub(r"[^A-Za-z]", "", s).upper()
    if 1 <= len(code) <= 3 and code in _NAME_MAP:
        return _NAME_MAP[code]
    return s.upper()


def _relator_from_filename(path: Path) -> str:
    """Extrai o relator do nome do arquivo exportado pelo e-TCM.
    Ex.: PLENARIO_DOMINGOS_DISSEI.xlsx -> DOMINGOS DISSEI
    """
    stem = path.stem  # ex.: PLENARIO_DOMINGOS_DISSEI
    stem = re.sub(r"(?i)^PLENARIO_", "", stem)
    stem = re.sub(r"[_\s]+", " ", stem).strip()
    return _expand_initials(stem)


def _is_reinclusao_text(motivo: str) -> bool:
    """Detecta 'reinclusão' de forma robusta (ignora acento, hífen, caixa)."""
    if not motivo:
        return False
    t = _strip_accents_lower(motivo)
    t = t.replace("-", "").replace(" ", "")
    return ("reinclus" in t) or ("reinc" in t)


def _alpha(n: int) -> str:
    """1->A, 2->B, ..., 26->Z, 27->AA ..."""
    s = ""
    while n > 0:
        n -= 1
        s = chr(65 + (n % 26)) + s
        n //= 26
    return s or "A"


def _ler_planilha(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=str)
    df.columns = [str(c) for c in df.columns]
    # Garante no mínimo 10 colunas para atender aos fallbacks (inclui Motivo)
    if len(df.columns) < 10:
        for k in range(len(df.columns), 10):
            df[f"_X{k+1}"] = ""
        df = df[[*df.columns]]

    proc_idx, obj_idx, relator_idx, revisor_idx, motivo_idx = _detect_cols_basic(df.columns.tolist())

    processos = df.iloc[:, proc_idx].apply(_ws)
    objetos = df.iloc[:, obj_idx].apply(_ws)

    # RELATOR (coluna 7; se vazio, extrai do nome do arquivo)
    if relator_idx is not None and relator_idx < len(df.columns):
        relatores = df.iloc[:, relator_idx].apply(_expand_initials)
        if relatores.fillna("").eq("").all():
            relator_nome_arquivo = _relator_from_filename(path)
            relatores = pd.Series([relator_nome_arquivo] * len(df))
    else:
        relator_nome_arquivo = _relator_from_filename(path)
        relatores = pd.Series([relator_nome_arquivo] * len(df))

    # REVISOR (coluna 8; se ausente, deixa '-')
    if revisor_idx is not None and revisor_idx < len(df.columns):
        revisores = df.iloc[:, revisor_idx].apply(_expand_initials)
    else:
        revisores = pd.Series([""] * len(df))

    # Completa revisor ausente com base nas duplas definidas
    def _rev_fallback(rel: str, rev: str) -> str:
        r = _ws(rev)
        if r and r != "-":
            return r
        key = _strip_accents_lower(_ws(rel))
        return _DUO_RELATOR_REVISOR.get(key, r)

    if not relatores.empty:
        revisores = pd.Series([
            _rev_fallback(rel, rev) for rel, rev in zip(relatores.tolist(), revisores.tolist())
        ])

    # MOTIVO (coluna 10) → flag de reinclusão
    if motivo_idx is not None and motivo_idx < len(df.columns):
        motivos = df.iloc[:, motivo_idx].apply(_ws)
    else:
        motivos = pd.Series([""] * len(df))
    is_reinc = motivos.apply(_is_reinclusao_text)

    out = pd.DataFrame(
        {
            "Relator": relatores.apply(_ws),
            "Revisor": revisores.apply(lambda s: _ws(s) if _ws(s) else "-"),
            "Processo": processos.apply(_ws),
            "Objeto": objetos.apply(_ws),
            "Motivo": motivos.apply(_ws),
            "IsReinc": is_reinc.astype(bool),
        }
    )

    # filtra linhas válidas
    out = out[(out["Processo"] != "") & (out["Objeto"] != "")]
    out["Fonte"] = path.name
    return out


def _coletar_planilhas(pasta_planilhas: str | Path) -> pd.DataFrame:
    pasta = Path(pasta_planilhas)
    arquivos = sorted([*pasta.glob("*.xlsx"), *pasta.glob("*.xls")])
    frames = []
    for arq in arquivos:
        try:
            frames.append(_ler_planilha(arq))
        except Exception as e:
            print(f"[docx] Aviso: falha ao ler {arq.name}: {e}")
    if not frames:
        return pd.DataFrame(columns=["Relator", "Revisor", "Processo", "Objeto", "Motivo", "IsReinc", "Fonte"])
    full = pd.concat(frames, ignore_index=True)

    # Ordena pela sequência padrão de relatores e, em seguida, por revisor e processo
    def _norm_name(n: str) -> str:
        return _strip_accents_lower(_ws(n))

    ordem_relatores = {
        # Ordem padrão (será substituída por composição específica em gerar_docx_unificado)
        "domingos dissei": 1,
        "ricardo torres": 2,
        "roberto braguim": 3,
        "joao antonio": 4,
        "eduardo tuma": 5,
    }

    full["_RelatorOrder"] = full["Relator"].map(lambda n: ordem_relatores.get(_norm_name(n), 999))
    full = (
        full
        .sort_values(by=["_RelatorOrder", "Relator", "Revisor", "Processo"], kind="stable")
        .reset_index(drop=True)
        .drop(columns=["_RelatorOrder"])
    )
    return full


def _open_document_from_template(header_template: str | Path | None) -> Document:
    here = Path(__file__).resolve().parent
    cwd = Path.cwd()
    candidates: list[Path] = []

    def push(p: Path):
        if p not in candidates:
            candidates.append(p)

    if header_template:
        hp = Path(header_template)
        push(hp)
        if not hp.is_absolute():
            push(cwd / hp)
            push(here / hp)
        if hp.suffix.lower() != ".docx":
            push(hp.with_suffix(".docx"))
        if hp.name.lower() == "papel_timbrado_tcm.docx":
            push(cwd / "papel_timbrado_tcm.docx.docx")
            push(here / "papel_timbrado_tcm.docx.docx")

    for n in [
        "papel_timbrado_tcm.docx",
        "papel_timbrado_tcm.docx.docx",
        "PAPEL TIMBRADO.docx",
        "PAPEL TIMBRADO.DOCX",
    ]:
        push(cwd / n)
        push(here / n)

    for p in candidates:
        try:
            if p.exists() and p.is_file():
                print(f"[docx] Template candidato: {p}")
                return Document(str(p))
        except PackageNotFoundError:
            print(f"[docx] Aviso: '{p}' não é um DOCX válido. Tentando próximo.")
        except Exception as e:
            print(f"[docx] Aviso: falha ao abrir '{p}': {e}. Tentando próximo.")

    print("[docx] Aviso: nenhum template válido encontrado. Gerando sem papel timbrado.")
    return Document()


def _fontify(run, size=12, small_caps=False, bold=False):
    f = run.font
    f.name = "Arial"
    f.size = Pt(size)
    f.small_caps = bool(small_caps)
    f.bold = bool(bold)


def _para_fmt(paragraph, align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=6, after=6, line=1.15):
    pf = paragraph.paragraph_format
    paragraph.alignment = align
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line


def _add_centered(doc: Document, texto: str, bold=False, size=12) -> None:
    p = doc.add_paragraph()
    _para_fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER, before=6, after=6, line=1.15)
    run = p.add_run(texto)
    _fontify(run, size=size, small_caps=False, bold=bold)


def _add_assinatura_final(doc: Document) -> None:
    """Adiciona bloco de assinatura.
    Se TCM_ASSINATURA_NOME/TCM_ASSINATURA_CARGO estiverem definidos, usa assinatura customizada;
    caso contrário, usa o bloco padrão (Presidente, Vice, Corregedor).
    """
    import os
    nome = os.getenv("TCM_ASSINATURA_NOME", "").strip()
    cargo = os.getenv("TCM_ASSINATURA_CARGO", "").strip()
    data_linha = os.getenv("TCM_ASSINATURA_DATA", "").strip()
    doc.add_paragraph("")
    if nome and cargo:
        _add_centered(doc, nome, bold=False, size=11)
        _add_centered(doc, cargo, bold=False, size=11)
        if data_linha:
            doc.add_paragraph("")
            _add_centered(doc, data_linha, bold=False, size=11)
    else:
        # Assinatura padrão solicitada
        _add_centered(doc, "ROSELI DE MORAIS CHAVES", bold=False, size=11)
        _add_centered(doc, "SUBSECRETÁRIA-GERAL", bold=False, size=11)
        doc.add_paragraph("")
        _add_centered(doc, "30 de outubro de 2025", bold=False, size=11)


def _add_item_paragraph(doc: Document, processo: str, objeto: str, idx: int | None = None) -> None:
    p = doc.add_paragraph()
    _para_fmt(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=0, after=6, line=1.15)

    if idx is not None:
        r_idx = p.add_run(f"{idx}) ")
        _fontify(r_idx, size=12, bold=True)

    r_proc = p.add_run(_ws(processo))
    _fontify(r_proc, size=12, bold=True)

    r_sep = p.add_run(" - ")
    _fontify(r_sep, size=12)

    _add_obj_with_highlights(p, _ws(objeto))


def _add_obj_with_highlights(paragraph, texto: str) -> None:
    """Renderiza o texto do objeto dividindo em runs e deixando termos-chave em negrito."""
    if not texto:
        r = paragraph.add_run("")
        _fontify(r, size=12)
        return

    # Destaca apenas a primeira ocorrência de cada termo (por label), sem sobreposição
    selected: list[tuple[int, int, str]] = []  # (start, end, label)
    used_labels: set[str] = set()

    def _overlaps(s: int, e: int) -> bool:
        for ss, ee, _ in selected:
            if not (e <= ss or s >= ee):
                return True
        return False

    for label, pat in _OBJ_HL_PATTERNS:
        if label in used_labels:
            continue
        m = pat.search(texto)
        if not m:
            continue
        s, e = m.start(), m.end()
        if s >= 0 and e > s and not _overlaps(s, e):
            selected.append((s, e, label))
            used_labels.add(label)

    if not selected:
        r = paragraph.add_run(texto)
        _fontify(r, size=12)
        return

    # Ordena pelo índice de início
    selected.sort(key=lambda t: t[0])

    pos = 0
    for s, e, _ in selected:
        if pos < s:
            r = paragraph.add_run(texto[pos:s])
            _fontify(r, size=12)
        r = paragraph.add_run(texto[s:e])
        _fontify(r, size=12, bold=True)
        pos = e
    if pos < len(texto):
        r = paragraph.add_run(texto[pos:])
        _fontify(r, size=12)


_ROMANS = [
    (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
    (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
    (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
]


def _roman(n: int) -> str:
    if n <= 0:
        return str(n)
    out = []
    for val, sym in _ROMANS:
        while n >= val:
            out.append(sym)
            n -= val
    return "".join(out)


# =========================
# Cabeçalhos por tipo de sessão
# =========================

@dataclass
class SessionMeta:
    numero: str                 # ex: "71" ou "3.385"
    tipo: str                   # 'ordinaria' | 'extraordinaria'
    formato: str                # 'nao-presencial' | 'presencial'
    competencia: str            # 'pleno' | '1c' | '2c'
    data_abertura: str          # "DD/MM/AAAA" (NP) OU data da realização (presencial)
    data_encerramento: str = "" # só NP (se vazio, calcula +15 dias)
    horario: str = "9h30min."   # só presencial
    local: str = (
        "NO PLENÁRIO DO EDIFÍCIO PREFEITO FARIA LIMA E COM TRANSMISSÃO AO VIVO "
        "PELO CANAL TV TCMSP NO YOUTUBE."
    )

    def normalizar(self):
        self.tipo = self.tipo.strip().lower()
        self.formato = self.formato.strip().lower()
        self.formato = {
            "nao presencial": "nao-presencial",
            "nao-presencial": "nao-presencial",
            "presencial": "presencial",
        }.get(self.formato, self.formato)
        self.competencia = self.competencia.strip().lower()
        # garante 'ª' ao final
        if "ª" not in self.numero:
            try:
                int(self.numero.replace(".", ""))
                self.numero = f"{self.numero}ª"
            except Exception:
                pass
        # A data informada é considerada como data de publicação.
        pub = _parse_date_br(self.data_abertura) if self.data_abertura else None

        # Regras de abertura por tipo/formato/competência
        abertura_calc: date | None = None
        # Pleno presencial: quarta-feira da semana seguinte ao disparo
        if self.formato == "presencial" and self.competencia == "pleno":
            now = datetime.now()
            abertura_calc = _weekday_of_next_week(now, weekday=2)  # Wednesday of next week
        elif pub is not None:
            # SONP: ordinária não-presencial → 1ª terça-feira do mês subsequente
            if self.formato == "nao-presencial" and self.tipo.startswith("ordin"):
                abertura_calc = _first_weekday_of_next_month(pub, weekday=1)  # Tuesday
            # SENP: extraordinária não-presencial → 2ª terça-feira do mês subsequente
            elif self.formato == "nao-presencial" and self.tipo.startswith("extra"):
                abertura_calc = _nth_weekday_of_next_month(pub, weekday=1, n=2)  # Tuesday, 2ª

        if abertura_calc is not None:
            self.data_abertura = _fmt_date_br(abertura_calc)
        # Força de datas finais via env (mantém compatibilidade)
        import os
        ab_forcada = os.getenv("TCM_META_ABERTURA_FINAL", "").strip()
        en_forcada = os.getenv("TCM_META_ENCERRAMENTO_FINAL", "").strip()
        if ab_forcada:
            d = _parse_date_br(ab_forcada)
            if d is not None:
                self.data_abertura = _fmt_date_br(d)

        if en_forcada:
            d = _parse_date_br(en_forcada)
            if d is not None:
                self.data_encerramento = _fmt_date_br(d)
        else:
            # Todas as sessões: 15 dias corridos a partir da Abertura
            base = _parse_date_br(self.data_abertura)
            if base is not None:
                self.data_encerramento = _fmt_date_br(base + timedelta(days=15))
            else:
                self.data_encerramento = ""


def _meta_from_env() -> Optional[SessionMeta]:
    import os
    tipo = os.getenv("TCM_META_TIPO", "").strip()
    formato = os.getenv("TCM_META_FORMATO", "").strip()
    comp = os.getenv("TCM_META_COMPETENCIA", "").strip()
    num = os.getenv("TCM_META_NUMERO", "").strip()
    d_ab = os.getenv("TCM_META_DATA_ABERTURA", "").strip()
    d_en = os.getenv("TCM_META_DATA_ENCERRAMENTO", "").strip()
    hr = os.getenv("TCM_META_HORARIO", "").strip() or "9h30min."
    if not (tipo and formato and comp and num and d_ab):
        return None
    meta = SessionMeta(numero=num, tipo=tipo, formato=formato, competencia=comp,
                       data_abertura=d_ab, data_encerramento=d_en, horario=hr)
    meta.normalizar()
    return meta


def _texto_competencia(meta: SessionMeta) -> str:
    if meta.competencia in ("1c", "1ª", "1a", "primeira", "1ª camara", "1a camara"):
        return "1ª CÂMARA"
    if meta.competencia in ("2c", "2ª", "2a", "segunda", "2ª camara", "2a camara"):
        return "2ª CÂMARA"
    return "PLENO"


def _montar_intro(meta: SessionMeta) -> str:
    tipo_up = "ORDINÁRIA" if meta.tipo.startswith("ordin") else "EXTRAORDINÁRIA"
    if meta.formato == "nao-presencial":
        artigo = "art.153-A" if meta.tipo.startswith("ordin") else "art.153-b c/c art. 153 § 5º"
        enc = (
            f" e o encerramento previsto para 15 dias corridos ({meta.data_encerramento})."
            if meta.data_encerramento else "."
        )
        return (
            f"PAUTA DA {meta.numero} SESSÃO {tipo_up} NÃO PRESENCIAL DO TRIBUNAL DE CONTAS "
            f"DO MUNICÍPIO DE SÃO PAULO, nos termos do {artigo} do Regimento Interno do "
            f"TCMSP, cuja abertura está designada para o dia {meta.data_abertura}{enc} "
            f"Aplicam-se, no que couber, as disposições da Resolução n.º 07/2019 e da "
            f"Instrução n.º 01/2019."
        )
    else:
        comp_txt = _texto_competencia(meta)
        if comp_txt == "PLENO":
            return (
                f"PAUTA DA {meta.numero} SESSÃO {tipo_up} DO TRIBUNAL DE CONTAS DO MUNICÍPIO DE "
                f"SÃO PAULO, A REALIZAR-SE NO DIA {meta.data_abertura}, ÀS {meta.horario}, "
                f"{meta.local}"
            )
        else:
            return (
                f"PAUTA DA {meta.numero} SESSÃO {tipo_up} DA {comp_txt} DO TRIBUNAL DE CONTAS "
                f"DO MUNICÍPIO DE SÃO PAULO, A REALIZAR-SE NO DIA {meta.data_abertura}, ÀS "
                f"{meta.horario}, {meta.local}"
            )


def _add_intro_from_meta(doc: Document, meta: SessionMeta) -> None:
    # Centraliza "PAUTA" na primeira linha (em negrito), e o restante justificado abaixo.
    intro = _montar_intro(meta)
    head = "PAUTA"
    rest = intro
    if intro.upper().startswith("PAUTA "):
        rest = intro[len("PAUTA "):].lstrip()

    _add_centered(doc, head, bold=True, size=14)

    p = doc.add_paragraph()
    _para_fmt(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=0, after=10, line=1.15)
    run = p.add_run(rest)
    _fontify(run, size=12)
    _add_centered(doc, "- I -", bold=True, size=12)
    _add_centered(doc, "ORDEM DO DIA", bold=True, size=12)
    doc.add_paragraph("")
    _add_centered(doc, "- II -", bold=True, size=12)
    _add_centered(doc, "JULGAMENTOS", bold=True, size=12)


def _add_intro_padrao(doc: Document, titulo: Optional[str]) -> None:
    _add_centered(doc, "PAUTA", bold=True, size=14)
    intro = (
        "DA SESSÃO ORDINÁRIA NÃO PRESENCIAL DO TRIBUNAL DE CONTAS DO MUNICÍPIO DE SÃO PAULO, "
        "nos termos das disposições da Resolução n.º 07/2019 e da Instrução n.º 01/2019."
    )
    p = doc.add_paragraph(intro)
    _para_fmt(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=0, after=8, line=1.15)
    _fontify(p.runs[0] if p.runs else p.add_run(""), size=12)
    _add_centered(doc, "- I -", bold=True, size=12)
    _add_centered(doc, "ORDEM DO DIA", bold=True, size=12)
    doc.add_paragraph("")
    _add_centered(doc, "- II -", bold=True, size=12)
    _add_centered(doc, "JULGAMENTOS", bold=True, size=12)


# =========================
# Geração do DOCX
# =========================
def gerar_docx_unificado(
    pasta_planilhas: str | Path,
    saida_docx: str | Path,
    titulo: str | None = None,
    header_template: str | Path | None = None,
    meta_sessao: SessionMeta | None = None,
) -> str:
    df = _coletar_planilhas(pasta_planilhas)
    if df.empty:
        raise RuntimeError("Nenhuma planilha válida encontrada para unificação.")

    doc = _open_document_from_template(header_template)

    # Cabeçalho contextual (se houver meta) ou padrão
    if meta_sessao is None:
        env_meta = _meta_from_env()
        meta_sessao = env_meta
    if meta_sessao:
        meta_sessao.normalizar()
        _add_intro_from_meta(doc, meta_sessao)
    else:
        _add_intro_padrao(doc, titulo)

    # Reordena os relatores conforme a composição da sessão
    def _ordem_por_competencia(comp: Optional[str]) -> dict:
        comp = (comp or "").strip().lower()
        if comp == "1c":
            base = [
                "domingos dissei",
                "ricardo torres",
                "roberto braguim",
            ]
        elif comp == "2c":
            base = [
                "ricardo torres",
                "joao antonio",
                "eduardo tuma",
            ]
        else:  # 'pleno' ou qualquer outro
            base = [
                "domingos dissei",
                "ricardo torres",
                "roberto braguim",
                "joao antonio",
                "eduardo tuma",
            ]
        return {name: i + 1 for i, name in enumerate(base)}

    ordem_map = _ordem_por_competencia(getattr(meta_sessao, "competencia", None))
    df["__RelatorOrder"] = df["Relator"].map(lambda n: ordem_map.get(_strip_accents_lower(_ws(n)), 999))

    # Prioridade fixa de revisores dentro de cada relator
    # 1) Vice-Presidente Ricardo Torres; 2) Corregedor Roberto Braguim; 3) João Antonio; 4) Eduardo Tuma; demais depois.
    rev_order = {
        "ricardo torres": 1,
        "roberto braguim": 2,
        "joao antonio": 3,
        "eduardo tuma": 4,
    }
    df["__RevisorOrder"] = df["Revisor"].map(lambda n: rev_order.get(_strip_accents_lower(_ws(n)), 999))
    df = (
        df.sort_values(by=["__RelatorOrder", "Relator", "__RevisorOrder", "Revisor", "Processo"], kind="stable")
          .reset_index(drop=True)
          .drop(columns=["__RelatorOrder", "__RevisorOrder"])
    )

    # Função para rotular o cargo do revisor no título
    def _cargo_revisor(nome: str) -> str:
        k = _strip_accents_lower(_ws(nome))
        if k == "domingos dissei":
            return "CONSELHEIRO PRESIDENTE"
        if k == "ricardo torres":
            return "CONSELHEIRO VICE-PRESIDENTE"
        if k == "roberto braguim":
            return "CONSELHEIRO CORREGEDOR"
        return "CONSELHEIRO"

    # ==== Passo 1: itens NÃO reinclusão (IsReinc == False) ====
    df_main = df[df["IsReinc"] == False]
    roman_counter = 1
    for relator, bloco_relator in df_main.groupby("Relator", sort=False):
        # Título do RELATOR (ALINHADO À ESQUERDA)
        cargo = _cargo_conselheiro(relator)
        rotulo_relator = f"{_roman(roman_counter)} - RELATOR {cargo} {relator}".upper()
        p_rel = doc.add_paragraph()
        _para_fmt(p_rel, align=WD_ALIGN_PARAGRAPH.LEFT, before=8, after=6, line=1.0)
        run_rel = p_rel.add_run(rotulo_relator)
        _fontify(run_rel, size=12, bold=True)
        roman_counter += 1

        # Dentro do relator: agrupar por REVISOR (com letras APENAS se houver >1 revisor)
        groups = list(bloco_relator.groupby("Revisor", sort=False))
        multi = len(groups) > 1
        for idx, (revisor, bloco_revisor) in enumerate(groups, start=1):
            prefix = f"{_alpha(idx)} - " if multi else ""
            subt = doc.add_paragraph()
            _para_fmt(subt, align=WD_ALIGN_PARAGRAPH.LEFT, before=4, after=2, line=1.0)
            cargo = _cargo_revisor(revisor)
            run_sub = subt.add_run(f"{prefix}REVISOR {cargo} {revisor}")
            _fontify(run_sub, size=12, bold=True)

            for i, row in enumerate(bloco_revisor.itertuples(index=False), start=1):
                _add_item_paragraph(doc, row.Processo, row.Objeto, idx=i)

            doc.add_paragraph("")  # espaçamento entre revisores

        doc.add_paragraph("")  # espaço entre relatores

    # ==== Passo 2: REINCLUSÕES ao final (IsReinc == True) ====
    df_reinc = df[df["IsReinc"] == True]
    if not df_reinc.empty:
        _add_centered(doc, "- REINCLUSÕES -", bold=True, size=12)
        doc.add_paragraph("")

        for relator, bloco_relator in df_reinc.groupby("Relator", sort=False):
            cargo = _cargo_conselheiro(relator)
            rotulo_relator = f"{_roman(roman_counter)} - RELATOR {cargo} {relator}".upper()
            p_rel = doc.add_paragraph()
            _para_fmt(p_rel, align=WD_ALIGN_PARAGRAPH.LEFT, before=8, after=6, line=1.0)
            run_rel = p_rel.add_run(rotulo_relator)
            _fontify(run_rel, size=12, bold=True)
            roman_counter += 1

            groups = list(bloco_relator.groupby("Revisor", sort=False))
            multi = len(groups) > 1
            for idx, (revisor, bloco_revisor) in enumerate(groups, start=1):
                prefix = f"{_alpha(idx)} - " if multi else ""
                subt = doc.add_paragraph()
                _para_fmt(subt, align=WD_ALIGN_PARAGRAPH.LEFT, before=4, after=2, line=1.0)
                cargo = _cargo_revisor(revisor)
                run_sub = subt.add_run(f"{prefix}REVISOR {cargo} {revisor}")
                _fontify(run_sub, size=12, bold=True)

                for i, row in enumerate(bloco_revisor.itertuples(index=False), start=1):
                    _add_item_paragraph(doc, row.Processo, row.Objeto, idx=i)

                doc.add_paragraph("")

            doc.add_paragraph("")

    # Assinatura ao final (aplicada a todas as sessões)
    _add_assinatura_final(doc)

    out_path = Path(saida_docx)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return str(out_path)


def gerar_docx_vazio(
    *,
    saida_docx: str | Path,
    titulo: str | None = None,
    header_template: str | Path | None = None,
    meta_sessao: SessionMeta | None = None,
) -> str:
    """Gera um DOCX apenas com cabeçalho (meta ou padrão), sem itens."""
    doc = _open_document_from_template(header_template)
    if meta_sessao is None:
        env_meta = _meta_from_env()
        meta_sessao = env_meta
    if meta_sessao:
        meta_sessao.normalizar()
        _add_intro_from_meta(doc, meta_sessao)
    else:
        _add_intro_padrao(doc, titulo)
    # Indicação opcional de ausência de itens
    _add_centered(doc, "(Sem itens)", bold=False, size=11)
    # Assinatura ao final (aplicada a todas as sessões)
    _add_assinatura_final(doc)

    out_path = Path(saida_docx)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return str(out_path)
