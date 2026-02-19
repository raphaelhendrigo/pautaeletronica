import os
import re
from pathlib import Path

from docx import Document

from docx_maker import SessionMeta, gerar_docx_unificado


def _find_item(paragraphs: list[str], tc: str) -> str:
    pattern = re.compile(rf"^\d+\)\s+{re.escape(tc)}\b")
    for p in paragraphs:
        if pattern.search(p):
            return p
    raise AssertionError(f"Nao encontrou item {tc}")


def test_sonp74_published_compliance() -> None:
    base = Path("planilhas_74_2026")
    out_dir = Path("output")
    out_dir.mkdir(exist_ok=True)
    out_path = out_dir / "TESTE_SONP_74_PYTEST.docx"

    os.environ["TCM_ASSINATURA_DATA"] = "22/01/2026"
    meta = SessionMeta(
        numero="74",
        tipo="ordinaria",
        formato="nao-presencial",
        competencia="pleno",
        data_abertura="03/02/2026",
        data_encerramento="19/02/2026",
    )

    gerar_docx_unificado(
        pasta_planilhas=str(base),
        saida_docx=str(out_path),
        titulo="Pauta Unificada - Sess\u00e3o 74/2026",
        header_template=None,
        meta_sessao=meta,
    )

    doc = Document(str(out_path))
    paragraphs = [p.text for p in doc.paragraphs]
    text = "\n".join(paragraphs)

    assert "03/02/2026" in text
    assert "(19/02/2026)" in text
    assert "PROCESSOS DA 1\u00aa C\u00c2MARA" in text
    assert "PROCESSOS DO PLENO" in text

    sec_1c = text.split("PROCESSOS DA 1\u00aa C\u00c2MARA", 1)[1].split("PROCESSOS DA 2\u00aa C\u00c2MARA", 1)[0]
    assert "REVISOR" not in sec_1c

    item_012129 = _find_item(paragraphs, "TC/012129/2023")
    assert "Retorno \u00e0 pauta" in item_012129
    assert "Retirado de Pauta" not in item_012129
    assert "na 70\u00aa Sonp" in item_012129
    assert "na 68\u00aa Sonp" not in item_012129

    item_003428 = _find_item(paragraphs, "TC/003428/2016")
    assert "Retorno \u00e0 pauta" in item_003428
    assert "Retirado de Pauta" not in item_003428

    assert "valor do instrumento" not in text.lower()
    assert "()" not in text
    assert " ." not in text

    item_005107 = _find_item(paragraphs, "TC/005107/2016")
    assert "(R$ 4.498.938,66 est.)." in item_005107

    item_005116 = _find_item(paragraphs, "TC/005116/2016")
    assert "(R$ 537.000,00), cujo" in item_005116
    assert "(R$ 537.000,00). cujo" not in item_005116

    item_007543 = _find_item(paragraphs, "TC/007543/1999")
    assert "no valor de R$ 5.997.776,30 em 09/12/2015" in item_007543
    assert "no valor de em 09/12/2015" not in item_007543

    item_009301 = _find_item(paragraphs, "TC/009301/2022")
    assert "R$" not in item_009301

    assert re.search(r"\\bRT\\b", text) is None
    assert re.search(r"\\bRB\\b", text) is None

    assert "(Itens englobados: 1 e 2)" in text
    assert "(Tramitam em conjunto os TCs: TC/005107/2016 e TC/005116/2016)" in text
    assert "(Tramitam em conjunto os TCs: TC/003428/2016 e TC/003429/2016)" in text
    assert "(Itens englobados - 4 e 5)" in text
