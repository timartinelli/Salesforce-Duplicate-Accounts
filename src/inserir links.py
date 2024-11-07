import openpyxl

# Arquivo de entrada e saída
input_file = 'data/conta_duplicadas_no_info.xlsx'
output_file_links = 'data/conta_duplicadas_no_info_arquivo_com_links.xlsx'

# Passo 1: Abrir o arquivo e converter valores em links clicáveis na coluna AJ
workbook = openpyxl.load_workbook(input_file)
sheet = workbook['Contas']

# Coluna AJ (índice 36) para converter valores em links
for row in sheet.iter_rows(min_row=2, max_row=954, min_col=36, max_col=36):  # Coluna AJ é a coluna 36
    for cell in row:
        if cell.value and isinstance(cell.value, str):  # Verifica se há valor e é uma string
            link = cell.value
            # Definir o hyperlink na célula
            cell.hyperlink = link
            cell.value = link  # Define o texto visível da célula como o valor do link
            cell.style = "Hyperlink"  # Aplica o estilo de link (opcional)

# Salva o arquivo com links clicáveis
workbook.save(output_file_links)

print(f'Links na coluna AJ convertidos com sucesso para o arquivo {output_file_links}')
