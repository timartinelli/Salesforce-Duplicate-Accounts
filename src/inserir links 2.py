import openpyxl

# Input and output files
input_file = 'data/conta_duplicadas_no_info_arquivo_com_links.xlsx'
output_file = 'data/conta_duplicadas_no_info_arquivo_com_links_clickable.xlsx'

# Load the workbook and worksheet
workbook = openpyxl.load_workbook(input_file)
sheet = workbook['Contas']

# Iterate over the cells in column AI (column 35) and convert them into clickable links
for row in range(2, sheet.max_row + 1):  # Starts from row 2 to the last row
    cell = sheet[f'AI{row}']
    if cell.value:  # Only if there is a value
        # Apply the clickable hyperlink formula
        cell.value = f'=HYPERLINK("{cell.value}", "{cell.value}")'

# Save the output file
workbook.save(output_file)

print(f'Clickable links in column AI have been saved to: {output_file}')
