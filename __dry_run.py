from main import run_pipeline
from pathlib import Path

# Create a dummy output dir
Path('output').mkdir(exist_ok=True)

# Call pipeline with a download folder that is empty to force empty docx path
print('Running pipeline with empty download dir...')
out = run_pipeline(
    base_url='https://etcm.tcm.sp.gov.br',
    usuario='user',
    senha='pass',
    num_sessao='361',
    data_de='01/04/2025',
    data_ate='30/04/2025',
    download_dir='__empty_dl_for_test',
    output_dir='output',
    headless=True,
    titulo_docx='Pauta Classificada',
    header_template=None,
    nome_docx='TEST_DOC.docx',
)
print('Pipeline returned:', out)
