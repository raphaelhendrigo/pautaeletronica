# utils.py
from __future__ import annotations
import unicodedata
import re
import pandas as pd

# Normaliza string para comparação (sem acento, minúscula e sem pontuação)
def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s

# Alias robusto para cabeçalhos vindos das planilhas do e-TCM
# mapeia "apelidos" -> nome padronizado
_ALIAS = {
    # Processo
    "processo": "Processo",
    "nº processo": "Processo",
    "nº do processo": "Processo",
    "numero processo": "Processo",
    "número do processo": "Processo",
    "proc.": "Processo",
    "n processo": "Processo",

    # Objeto de Julgamento
    "objeto de julgamento": "Objeto de Julgamento",
    "objeto": "Objeto de Julgamento",
    "assunto": "Objeto de Julgamento",
    "descricao": "Objeto de Julgamento",
    "descrição": "Objeto de Julgamento",
    "texto": "Objeto de Julgamento",

    # Revisor / Relator
    "revisor": "Revisor",
    "relator": "Revisor",

    # Extras comuns (mantidos se aparecerem)
    "orgao": "Órgão",
    "órgão": "Órgão",
    "unidade gestora": "Órgão",
    "ug": "Órgão",
}

def _slug_col(c: str) -> str:
    c = str(c).replace("\n", " ").replace("\r", " ")
    c = re.sub(r"\s+", " ", c).strip()
    return _norm(c)

def normalizar_colunas_padrao(df: pd.DataFrame) -> pd.DataFrame:
    """
    Renomeia colunas variadas para um conjunto padrão:
    - 'Processo'
    - 'Objeto de Julgamento'
    - 'Revisor'
    Mantém o resto como está.
    """
    rename_map = {}
    for col in df.columns:
        key = _slug_col(col)
        alvo = _ALIAS.get(key)
        rename_map[col] = alvo if alvo else col
    return df.rename(columns=rename_map)

# (Opcional) útil para gerar nomes de arquivos
def slugify_nome(s: str) -> str:
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")
    return s.upper() or "DESCONHECIDO"
