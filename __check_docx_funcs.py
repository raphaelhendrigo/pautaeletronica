import importlib
m = importlib.import_module('docx_maker')
print('has gerar_docx_unificado:', hasattr(m,'gerar_docx_unificado'))
print('has gerar_docx_vazio:', hasattr(m,'gerar_docx_vazio'))
