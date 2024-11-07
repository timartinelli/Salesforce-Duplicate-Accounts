import openpyxl

# Arquivos de entrada e saída
input_file = 'data/conta_duplicadas_no_info_arquivo_com_links.xlsx'
output_file = 'data/conta_duplicadas_no_info_arquivo_com_links_clicaveis.xlsx'

# Carregar o workbook e a planilha
workbook = openpyxl.load_workbook(input_file)
sheet = workbook['Contas']

# Iterar sobre as células da coluna AI (coluna 35) e transformar em links clicáveis
for row in range(2, sheet.max_row + 1):  # Começa na linha 2 até a última linha
    cell = sheet[f'AI{row}']
    if cell.value:  # Apenas se houver um valor
        # Aplicar a fórmula de link clicável
        cell.value = f'=HYPERLINK("{cell.value}", "{cell.value}")'

# Salvar o arquivo de saída
workbook.save(output_file)

print(f'Links clicáveis na coluna AI foram salvos em: {output_file}')
