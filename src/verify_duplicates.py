import os
import pandas as pd
from simple_salesforce import Salesforce, SalesforceAuthenticationFailed
import logging

# Configurando o logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_salesforce_ids(sf):
    query = "SELECT Id, Name FROM Account"
    accounts = sf.query_all(query)
    return {account['Id']: account['Name'] for account in accounts['records']}  # Retorna um dicionário com ID e Nome

def verify_accounts(sf):
    try:
        logging.info("Verifying accounts...")

        # Passo 1: Exportar todos os IDs e nomes de contas do Salesforce
        salesforce_accounts = get_salesforce_ids(sf)
        salesforce_ids = set(salesforce_accounts.keys())  # Apenas os IDs
        logging.info(f"Retrieved {len(salesforce_ids)} accounts from Salesforce.")

        # Passo 2: Ler os IDs do arquivo Excel
        excel_file_path = 'data_source/conta_duplicadas.xlsx'
        df = pd.read_excel(excel_file_path)

        # Verifique se a coluna correta está presente
        if 'Account_18_Digit_ID__c' not in df.columns:
            logging.error("A coluna 'Account_18_Digit_ID__c' não foi encontrada no arquivo.")
            return

        # Passo 3: Criar uma nova coluna para verificar a existência das contas
        df['Account Exists'] = df['Account_18_Digit_ID__c'].apply(lambda x: 'Exists' if x in salesforce_ids else 'Does Not Exist')

        # Passo 4: Salvar o resultado em um novo arquivo Excel
        output_file_path = 'data/conta_duplicadas_verificacao.xlsx'
        df.to_excel(output_file_path, index=False)
        logging.info(f"Results saved to {output_file_path}")

        # Passo 5: Identificar novas contas que estão no Salesforce mas não na planilha
        new_accounts = {id_: name for id_, name in salesforce_accounts.items() if id_ not in df['Account_18_Digit_ID__c'].values}

        # Passo 6: Criar um DataFrame para novas contas
        new_accounts_df = pd.DataFrame(new_accounts.items(), columns=['Account ID', 'Account Name'])

        # Salvar novas contas em um novo arquivo Excel
        new_accounts_file_path = 'data/novas_contas.xlsx'
        new_accounts_df.to_excel(new_accounts_file_path, index=False)
        logging.info(f"New accounts saved to {new_accounts_file_path}")

        return {'existing_ids': salesforce_ids}

    except SalesforceAuthenticationFailed as e:
        logging.error(f"Authentication failed: {e}")
    except Exception as e:
        logging.error(f"An error occurred: {e}")

if __name__ == "__main__":
    pass  # O código não deve ser executado diretamente
