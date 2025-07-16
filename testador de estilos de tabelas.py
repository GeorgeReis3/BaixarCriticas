from docx import Document
from docx.enum.style import WD_STYLE_TYPE

# Nome do arquivo DOCX que você criou com vários estilos de tabela aplicados
nome_do_arquivo_docx = 'Modelo para crítica - (NÃO APAGAR).docx' # Altere para o nome do seu arquivo

try:
    # Tenta carregar o documento existente
    documento = Document(nome_do_arquivo_docx)
    print(f"Documento '{nome_do_arquivo_docx}' carregado com sucesso.")

    print("\nEstilos de Tabela encontrados no documento:")
    encontrou_estilos_tabela = False
    for style in documento.styles:
        # Verifica se o estilo é do tipo TABELA
        if style.type == WD_STYLE_TYPE.TABLE:
            print(f"- {style.name}")
            encontrou_estilos_tabela = True

    if not encontrou_estilos_tabela:
        print(f"Nenhum estilo de tabela explícito do tipo WD_STYLE_TYPE.TABLE foi encontrado em '{nome_do_arquivo_docx}'.")
        print("Certifique-se de ter aplicado estilos de tabela diferentes no Word e salvo o arquivo.")

except FileNotFoundError:
    print(f"ERRO: O arquivo '{nome_do_arquivo_docx}' não foi encontrado.")
    print("Por favor, crie um arquivo Word, aplique alguns estilos de tabela da galeria e salve-o com este nome.")
    print("Exemplo: Abra o Word, insira uma tabela, vá em 'Design da Tabela', aplique 3-4 estilos diferentes e salve como 'todos_estilos_tabela.docx'.")
except Exception as e:
    print(f"Ocorreu um erro ao processar o documento: {e}")