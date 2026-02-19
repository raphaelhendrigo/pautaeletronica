import os
import re
import unittest
from pathlib import Path

from docx import Document

from docx_maker import gerar_docx_unificado, SessionMeta


class TestSonp74Regression(unittest.TestCase):
    def test_sonp74_docx_content(self) -> None:
        base = Path("planilhas_74_2026")
        out_dir = Path("output")
        out_dir.mkdir(exist_ok=True)
        out_path = out_dir / "TESTE_SONP_74.docx"

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
            titulo="Pauta Unificada - Sessão 74/2026",
            header_template=None,
            meta_sessao=meta,
        )

        doc = Document(str(out_path))
        text = "\n".join(p.text for p in doc.paragraphs)

        self.assertIn("03/02/2026", text)
        self.assertIn("(19/02/2026)", text)
        self.assertNotIn("03/03/2026", text)
        self.assertNotIn("18/03/2026", text)

        self.assertIn("PROCESSOS DA 1ª CÂMARA", text)
        self.assertIn("PROCESSOS DA 2ª CÂMARA", text)
        self.assertIn("PROCESSOS DO PLENO", text)
        self.assertIn("PROCESSOS DE REINCLUSÃO", text)

        sec_1c = text.split("PROCESSOS DA 1ª CÂMARA", 1)[1].split("PROCESSOS DA 2ª CÂMARA", 1)[0]
        self.assertLess(sec_1c.find("TC/005107/2016"), sec_1c.find("TC/005116/2016"))

        sec_pleno = text.split("PROCESSOS DO PLENO", 1)[1]
        self.assertIn("TC/003496/2014", sec_pleno)
        first_tc = re.search(r"TC/\d{6}/\d{4}", sec_pleno)
        self.assertIsNotNone(first_tc)
        self.assertEqual(first_tc.group(0), "TC/003496/2014")

        self.assertIn("(Itens englobados: 1 e 2)", text)
        self.assertIn(
            "(Tramitam em conjunto os TCs: TC/005107/2016 e TC/005116/2016)",
            text,
        )
        self.assertIn(
            "(Tramitam em conjunto os TCs: TC/003428/2016 e TC/003429/2016)",
            text,
        )
        self.assertIn("(Itens englobados - 4 e 5)", text)

        self.assertIn(
            "Retorno à pauta, após determinação do Conselheiro Presidente Domingos Dissei, na 63ª Sonp, "
            "para que os autos lhe fossem conclusos, para proferir voto de desempate, tendo como Relator o "
            "Conselheiro Eduardo Tuma.",
            text,
        )
        self.assertIn(
            "Retorno à pauta, após determinação do Conselheiro Presidente Domingos Dissei, na 3.369ª SO, "
            "para que os autos lhe fossem conclusos, para proferir voto de desempate, tendo como Relator o "
            "Conselheiro Vice-Presidente Ricardo Torres.",
            text,
        )

        self.assertIn("Retirado de Pauta na 63ª Sonp", text)
        self.assertNotIn("50ª Sonp", text)

        self.assertNotIn("RT/RB", text)
        self.assertNotIn("pesquisado em", text.lower())
        self.assertNotIn("verificado até peç", text.lower())
        self.assertNotIn("R$ 95.445.376,00", text)


if __name__ == "__main__":
    unittest.main()
