from docx import Document
from docx.enum.style import WD_STYLE_TYPE

# Carregue o documento que contém o estilo de tabela desejado
doc_com_estilo = Document('Modelo para crítica - (NÃO APAGAR).docx')

print("Estilos de Tabela encontrados neste documento:")
for style in doc_com_estilo.styles:
    if style.type == WD_STYLE_TYPE.TABLE:
        print(f"- {style.name}")
# Anote o nome exato que você quer usar, por exemplo: 'Table Grid'