import openpyxl

# Input and output files
input_file = 'data/conta_duplicadas_no_info.xlsx'
output_file_links = 'data/conta_duplicadas_no_info_file_with_links.xlsx'

# Step 1: Open the file and convert values into clickable links in column AJ
workbook = openpyxl.load_workbook(input_file)
sheet = workbook['Contas']

# Column AJ (index 36) to convert values into links
for row in sheet.iter_rows(min_row=2, max_row=954, min_col=36, max_col=36):  # Column AJ is column 36
    for cell in row:
        if cell.value and isinstance(cell.value, str):  # Check if there is a value and it's a string
            link = cell.value
            # Set the hyperlink in the cell
            cell.hyperlink = link
            cell.value = link  # Set the visible text of the cell as the link value
            cell.style = "Hyperlink"  # Apply the hyperlink style (optional)

# Save the file with clickable links
workbook.save(output_file_links)

print(f'Links in column AJ successfully converted to the file {output_file_links}')
