from docx_maker import SessionMeta

def show(pub):
    meta = SessionMeta(numero='15', tipo='extraordinaria', formato='nao-presencial', competencia='pleno', data_abertura=pub)
    meta.normalizar()
    print(pub, '=> abertura=', meta.data_abertura, 'encerramento=', meta.data_encerramento)

for pub in ['30/10/2025', '29/10/2025', '31/10/2025']:
    show(pub)
