import os
import pandas as pd
from simple_salesforce import Salesforce, SalesforceAuthenticationFailed

# Funções para a consulta no Salesforce
def get_salesforce_ids(sf):
    query = "SELECT Id, Name FROM Account"
    accounts = sf.query_all(query)
    return {account['Id']: account['Name'] for account in accounts['records']}  # Retorna um dicionário com ID e Nome

def main():
    # Configurações do Salesforce
    from config import USERNAME, PASSWORD, SECURITY_TOKEN, DOMAIN

    try:
        sf = Salesforce(username=USERNAME, password=PASSWORD, security_token=SECURITY_TOKEN, domain=DOMAIN)
        print("Connected to Salesforce!")

        # Passo 1: Exportar todos os IDs e nomes de contas do Salesforce
        salesforce_accounts = get_salesforce_ids(sf)
        salesforce_ids = set(salesforce_accounts.keys())  # Apenas os IDs

        # Passo 2: Ler os IDs do arquivo Excel
        excel_file_path = 'data_source/conta_duplicadas.xlsx'
        df = pd.read_excel(excel_file_path)

        # Verifique se a coluna correta está presente
        if 'Account_18_Digit_ID__c' not in df.columns:
            print("A coluna 'Account_18_Digit_ID__c' não foi encontrada no arquivo.")
            return

        # Passo 3: Criar uma nova coluna para verificar a existência das contas
        df['Account Exists'] = df['Account_18_Digit_ID__c'].apply(lambda x: 'Exists' if x in salesforce_ids else 'Does Not Exist')

        # Passo 4: Salvar o resultado em um novo arquivo Excel
        output_file_path = 'data/conta_duplicadas_verificacao.xlsx'
        df.to_excel(output_file_path, index=False)
        print(f"Results saved to {output_file_path}")

        # Passo 5: Identificar novas contas que estão no Salesforce mas não na planilha
        new_accounts = {id_: name for id_, name in salesforce_accounts.items() if id_ not in df['Account_18_Digit_ID__c'].values}

        # Passo 6: Criar um DataFrame para novas contas
        new_accounts_df = pd.DataFrame(new_accounts.items(), columns=['Account ID', 'Account Name'])

        # Salvar novas contas em um novo arquivo Excel
        new_accounts_file_path = 'data/novas_contas.xlsx'
        new_accounts_df.to_excel(new_accounts_file_path, index=False)
        print(f"New accounts saved to {new_accounts_file_path}")

    except SalesforceAuthenticationFailed as e:
        print(f"Authentication failed: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
