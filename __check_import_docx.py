import importlib
m = importlib.import_module('docx_maker')
print('OK', hasattr(m, 'gerar_docx_unificado'))
