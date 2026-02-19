import unittest

import pandas as pd

from docx_maker import (
    SEM_CATEGORIA_RANK,
    compute_primary_keyword,
    sort_items_for_segment,
    _split_objeto_runs,
)


class TestPrimaryKeyword(unittest.TestCase):
    def test_primary_keyword_specificity(self) -> None:
        text = "Acompanhamento - Execução Contratual referente a contrato."
        keyword, rank, span = compute_primary_keyword(text)
        self.assertEqual(keyword, "Acompanhamento - Execução Contratual")
        self.assertEqual(rank, 4)
        self.assertEqual(text[span[0]:span[1]], "Acompanhamento - Execução Contratual")

    def test_primary_keyword_groups(self) -> None:
        _, rank, _ = compute_primary_keyword("Embargos de Declaração em face do acórdão")
        self.assertEqual(rank, 1)
        _, rank, _ = compute_primary_keyword("Recurso ordinário")
        self.assertEqual(rank, 2)
        _, rank, _ = compute_primary_keyword("Pedido de Revisão do processo")
        self.assertEqual(rank, 3)
        _, rank, _ = compute_primary_keyword("Contrato emergencial")
        self.assertEqual(rank, 4)

    def test_primary_keyword_none(self) -> None:
        keyword, rank, span = compute_primary_keyword("Texto sem tipo conhecido")
        self.assertIsNone(keyword)
        self.assertEqual(rank, SEM_CATEGORIA_RANK)
        self.assertIsNone(span)

    def test_split_objeto_runs_bold(self) -> None:
        text = (
            "Representação sobre Pregão Eletrônico. "
            "Para proferir voto de desempate. "
            "(Itens englobados - 4 e 5)."
        )
        runs = _split_objeto_runs(text)
        bold_chunks = "".join(chunk for chunk, bold in runs if bold)
        self.assertIn("Representação", bold_chunks)
        self.assertIn("Para proferir voto de desempate", bold_chunks)
        self.assertIn("Itens englobados - 4 e 5", bold_chunks)
        self.assertNotIn("Pregão Eletrônico", bold_chunks)


class TestSorting(unittest.TestCase):
    def test_sort_items_for_segment_groups(self) -> None:
        items = pd.DataFrame(
            [
                {"Processo": "1", "Objeto": "Contrato emergencial para obra"},
                {"Processo": "2", "Objeto": "Recurso ordinário"},
                {"Processo": "3", "Objeto": "Embargos de Declaração do acórdão"},
                {"Processo": "4", "Objeto": "Pedido de Revisão do processo"},
                {"Processo": "5", "Objeto": "Texto sem categoria"},
                {"Processo": "6", "Objeto": "Embargo de Declaração do voto"},
            ]
        )
        sorted_df = sort_items_for_segment(items)
        self.assertEqual(sorted_df["Processo"].tolist(), ["3", "6", "2", "4", "1", "5"])


if __name__ == "__main__":
    unittest.main()
