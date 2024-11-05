import pandas as pd
import os

def copy_and_mark_duplicates(duplicates):
    input_file_path = os.path.join('data_source', 'conta_duplicadas.xlsx')
    
    if not os.path.exists(input_file_path):
        print(f"Input file {input_file_path} does not exist.")
        return

    df = pd.read_excel(input_file_path)

    # Nova coluna para marcar IDs duplicados
    df['Duplicated Account IDs'] = ''

    for cpf_cnpj, ids in duplicates.items():
        for idx in df.index:
            if df.at[idx, 'NEO_Cpfcnpj__c'] == cpf_cnpj:
                df.at[idx, 'Duplicated Account IDs'] = ', '.join(ids)

    output_file_path = os.path.join('data', 'conta_duplicadas_marked.xlsx')
    df.to_excel(output_file_path, index=False)
    print(f"Marked duplicates saved to {output_file_path}")
