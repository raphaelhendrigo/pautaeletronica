from pathlib import Path
import pandas as pd

from docx_maker import gerar_docx_unificado, SessionMeta


def main():
    base = Path('test_planilhas')
    base.mkdir(exist_ok=True)

    # Cria 3 planilhas simples, uma por relator, com 1 item cada
    dados = [
        {
            'fname': 'PLENARIO_DOMINGOS_DISSEI.xlsx',
            'Relator': 'DOMINGOS DISSEI',
            'Revisor': 'RICARDO TORRES',  # testa cargo de revisor
            'Processo': '0001/2025',
            'Objeto': 'Diversos - Teste de cargos',
            'Motivo': '',
        },
        {
            'fname': 'PLENARIO_RICARDO_TORRES.xlsx',
            'Relator': 'RICARDO TORRES',
            'Revisor': 'ROBERTO BRAGUIM',
            'Processo': '0002/2025',
            'Objeto': 'Recurso - Teste de cargos',
            'Motivo': '',
        },
        {
            'fname': 'PLENARIO_ROBERTO_BRAGUIM.xlsx',
            'Relator': 'ROBERTO BRAGUIM',
            'Revisor': 'DOMINGOS DISSEI',
            'Processo': '0003/2025',
            'Objeto': 'Contrato com Termo Aditivo - Teste de cargos',
            'Motivo': '',
        },
    ]

    cols = ['Processo', 'Objeto', 'Relator', 'Revisor', 'Motivo']
    for d in dados:
        df = pd.DataFrame([{k: d[k] for k in cols}])
        df.to_excel(base / d['fname'], index=False)

    # Gera o DOCX unificado a partir das planilhas de teste
    out_dir = Path('output'); out_dir.mkdir(exist_ok=True)
    out_path = out_dir / 'TESTE_CARGOS.docx'

    # Meta simples apenas para cabeçalho do documento (opcional)
    meta = SessionMeta(
        numero='TESTE', tipo='ordinaria', formato='presencial', competencia='pleno',
        data_abertura='30/04/2025', data_encerramento='', horario='9h30min.'
    )

    docx = gerar_docx_unificado(
        pasta_planilhas=str(base),
        saida_docx=str(out_path),
        titulo='Pauta Classificada - Teste Cargos',
        header_template=None,
        meta_sessao=meta,
    )
    print('Gerado:', docx)


if __name__ == '__main__':
    main()

