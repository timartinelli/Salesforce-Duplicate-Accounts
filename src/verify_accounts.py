import pandas as pd
from simple_salesforce import Salesforce

def get_salesforce_ids(sf):
    query = "SELECT Id, Name, NEO_Cpfcnpj__c FROM Account"
    accounts = sf.query_all(query)
    return accounts['records']  # Retorna todos os registros

def verify_accounts(sf):
    try:
        print("Verifying accounts...")

        # Passo 1: Ler os IDs do arquivo Excel
        excel_file_path = 'data_source/conta_duplicadas.xlsx'
        df = pd.read_excel(excel_file_path)

        # Verifique se a coluna correta está presente
        if 'Account_18_Digit_ID__c' not in df.columns:
            print("A coluna 'Account_18_Digit_ID__c' não foi encontrada no arquivo.")
            return

        # Passo 2: Verificar se as contas estão no Salesforce
        salesforce_accounts = get_salesforce_ids(sf)
        salesforce_ids = {account['Id']: account for account in salesforce_accounts}

        # Passo 3: Criar uma nova coluna para verificar a existência das contas
        df['Account Exists'] = df['Account_18_Digit_ID__c'].apply(lambda x: 'Exists' if x in salesforce_ids else 'Does Not Exist')

        # Passo 4: Salvar o resultado em um novo arquivo Excel
        output_file_path = 'data/conta_duplicadas_verificacao.xlsx'
        df.to_excel(output_file_path, index=False)
        print(f"Results saved to {output_file_path}")

    except Exception as e:
        print(f"An error occurred: {e}")
