"""Regera os 3 DOCX da SONP 80/2026 (Pleno, 1ª Câmara, 2ª Câmara) usando
as planilhas já baixadas em planilhas_SONP_80_*_2026.
"""
from __future__ import annotations

import os
from pathlib import Path

from docx_maker import gerar_docx_unificado, SessionMeta

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

# Data de abertura (1ª terça do mês subsequente à publicação): 02/06/2026.
DATA_ABERTURA = "02/06/2026"

# Forçar a data de abertura para não depender da data atual.
os.environ["TCM_META_ABERTURA_FINAL"] = DATA_ABERTURA

SESSOES = [
    {
        "pasta": ROOT / "planilhas_SONP_80_PLENO_2026",
        "saida": OUTPUT / "PAUTA_SONP_80_PLENO_2026.docx",
        "competencia": "pleno",
    },
    {
        "pasta": ROOT / "planilhas_SONP_80_1CAMARA_2026",
        "saida": OUTPUT / "PAUTA_SONP_80_1CAMARA_2026.docx",
        "competencia": "1c",
    },
    {
        "pasta": ROOT / "planilhas_SONP_80_2CAMARA_2026",
        "saida": OUTPUT / "PAUTA_SONP_80_2CAMARA_2026.docx",
        "competencia": "2c",
    },
]


def main() -> list[Path]:
    gerados: list[Path] = []
    for cfg in SESSOES:
        meta = SessionMeta(
            numero="80",
            tipo="ordinaria",
            formato="nao-presencial",
            competencia=cfg["competencia"],
            data_abertura=DATA_ABERTURA,
        )
        out = gerar_docx_unificado(
            pasta_planilhas=str(cfg["pasta"]),
            saida_docx=str(cfg["saida"]),
            header_template="papel_timbrado_tcm.docx",
            meta_sessao=meta,
        )
        print(f"[ok] gerado: {out}")
        gerados.append(Path(out))
    return gerados


if __name__ == "__main__":
    main()
