from docx_maker import SessionMeta

def show(meta):
    before = (meta.data_abertura, meta.data_encerramento)
    meta.normalizar()
    after = (meta.data_abertura, meta.data_encerramento)
    print(meta.tipo, meta.formato, meta.competencia, '::', before, '->', after)

# Publicação 30/04/2025
print('Regras com publicação 30/04/2025:')
sonp = SessionMeta(numero='361', tipo='ordinaria', formato='nao-presencial', competencia='1c', data_abertura='30/04/2025')
show(sonp)
senp = SessionMeta(numero='999', tipo='extraordinaria', formato='nao-presencial', competencia='2c', data_abertura='30/04/2025')
show(senp)
pleno_pres = SessionMeta(numero='3388', tipo='ordinaria', formato='presencial', competencia='pleno', data_abertura='30/04/2025')
show(pleno_pres)
