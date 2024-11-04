import sys
import os
import pandas as pd
from simple_salesforce import Salesforce, SalesforceAuthenticationFailed
from verify_duplicates import verify_accounts
import logging

# Adiciona o diretório atual ao path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from config import USERNAME, PASSWORD, SECURITY_TOKEN, DOMAIN

# Configurando logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

logging.info(f"Salesforce domain being used: {DOMAIN}")

def fetch_related_counts(sf, account_id):
    query = f"""
    SELECT 
        Id,
        Name,
        (SELECT Id FROM Cases),
        (SELECT Id FROM Opportunities),
        (SELECT Id FROM Contacts),
        (SELECT Id FROM Assets),
        (SELECT Id FROM Contracts),
        (SELECT Id FROM Orders),
        (SELECT Id FROM Tasks),
        (SELECT Id FROM Events)
    FROM Account WHERE Id = '{account_id}'
    """

    try:
        accounts = sf.query(query)
        
        if accounts and accounts['records']:
            return accounts['records'][0]  # Retorna o primeiro registro
        else:
            logging.warning(f"No records found for account ID: {account_id}")
            return {}
    except Exception as e:
        logging.error(f"Error querying Salesforce for account ID {account_id}: {e}")
        return {}

def check_duplicate_cpf(sf, cpf_cnpj, current_account_id):
    # Consulta para encontrar contas com o mesmo CPF/CNPJ, excluindo a conta atual
    query = f"""
    SELECT Id FROM Account WHERE NEO_Cpfcnpj__c = '{cpf_cnpj}' AND Id != '{current_account_id}'
    """
    try:
        logging.info(f"Executing query for CPF/CNPJ: {cpf_cnpj}")
        accounts = sf.query(query)
        
        if accounts and accounts['records']:
            duplicate_ids = [record['Id'] for record in accounts['records']]
            logging.info(f"CPF/CNPJ {cpf_cnpj} found in accounts: {duplicate_ids}")
            return duplicate_ids
        else:
            logging.info(f"No duplicates found for CPF/CNPJ: {cpf_cnpj}")
            return []
    except Exception as e:
        logging.error(f"Error querying Salesforce for CPF/CNPJ {cpf_cnpj}: {e}")
        return []

def main(input_file_path):
    try:
        # Verificação de existência do arquivo
        if not os.path.isfile(input_file_path):
            logging.error(f"Arquivo não encontrado: {input_file_path}")
            return
        
        logging.info("Arquivos na pasta 'data': %s", os.listdir('data'))

        sf = Salesforce(username=USERNAME, password=PASSWORD, security_token=SECURITY_TOKEN, domain=DOMAIN)
        logging.info("Connected to Salesforce!")

        # Chamando a função de verificação de contas duplicadas
        logging.info("Verifying accounts...")
        verification_result = verify_accounts(sf)

        # Certifique-se de que a verificação retorna os dados esperados
        if not verification_result or 'existing_ids' not in verification_result:
            logging.error("Verification result is empty or missing 'existing_ids'")
            return

        df = pd.read_excel(input_file_path)

        # Adicionar novas colunas para as contagens
        df['Case Count'] = 0
        df['Opportunity Count'] = 0
        df['Contact Count'] = 0
        df['Asset Count'] = 0
        df['Contract Count'] = 0
        df['Order Count'] = 0
        df['Task Count'] = 0
        df['Event Count'] = 0
        df['Duplicate CPF/CNPJ Check'] = ''  # Nova coluna para verificar duplicidade de CPF/CNPJ

        # Loop através das contas
        for index, row in df.iterrows():
            account_id = row['Account_18_Digit_ID__c']
            cpf_cnpj = row.get('NEO_Cpfcnpj__c')

            if row['STATUS'] == 'NO INFO' and account_id in verification_result['existing_ids']:
                related_counts = fetch_related_counts(sf, account_id)

                # Checa se o related_counts está vazio ou se cada relacionamento é um dicionário antes de acessar `.get()`
                case_records = related_counts.get('Cases', {}).get('records', []) if isinstance(related_counts.get('Cases'), dict) else []
                opportunity_records = related_counts.get('Opportunities', {}).get('records', []) if isinstance(related_counts.get('Opportunities'), dict) else []
                contact_records = related_counts.get('Contacts', {}).get('records', []) if isinstance(related_counts.get('Contacts'), dict) else []
                asset_records = related_counts.get('Assets', {}).get('records', []) if isinstance(related_counts.get('Assets'), dict) else []
                contract_records = related_counts.get('Contracts', {}).get('records', []) if isinstance(related_counts.get('Contracts'), dict) else []
                order_records = related_counts.get('Orders', {}).get('records', []) if isinstance(related_counts.get('Orders'), dict) else []
                task_records = related_counts.get('Tasks', {}).get('records', []) if isinstance(related_counts.get('Tasks'), dict) else []
                event_records = related_counts.get('Events', {}).get('records', []) if isinstance(related_counts.get('Events'), dict) else []

                # Preencher contagens nas colunas
                df.at[index, 'Case Count'] = len(case_records)
                df.at[index, 'Opportunity Count'] = len(opportunity_records)
                df.at[index, 'Contact Count'] = len(contact_records)
                df.at[index, 'Asset Count'] = len(asset_records)
                df.at[index, 'Contract Count'] = len(contract_records)
                df.at[index, 'Order Count'] = len(order_records)
                df.at[index, 'Task Count'] = len(task_records)
                df.at[index, 'Event Count'] = len(event_records)
                
                logging.info(f"Processed account {account_id} with related counts: Cases={len(case_records)}, Opportunities={len(opportunity_records)}, Contacts={len(contact_records)}, Assets={len(asset_records)}, Contracts={len(contract_records)}, Orders={len(order_records)}, Tasks={len(task_records)}, Events={len(event_records)}")

            elif row['STATUS'] in ['DUPLICATE', 'DUPLICATE - APPEARS 3X', 'DUPLICATE 5X'] and cpf_cnpj:
                duplicate_ids = check_duplicate_cpf(sf, cpf_cnpj, account_id)
                
                if duplicate_ids:
                    df.at[index, 'Duplicate CPF/CNPJ Check'] = 'Duplicate found: ' + ', '.join(duplicate_ids)
                    logging.info(f"CPF/CNPJ {cpf_cnpj} found in other accounts for account ID: {account_id}. Duplicate IDs: {duplicate_ids}")
                else:
                    df.at[index, 'Duplicate CPF/CNPJ Check'] = 'No duplicates found'
                    logging.info(f"CPF/CNPJ {cpf_cnpj} not found in other accounts for account ID: {account_id}")

        # Salvar o DataFrame atualizado em um novo arquivo na pasta 'data'
        output_file_path = os.path.join('data', 'conta_duplicadas_verificacao_atualizada.xlsx')
        df.to_excel(output_file_path, index=False)
        logging.info(f"Updated results saved to {output_file_path}")

    except SalesforceAuthenticationFailed as e:
        logging.error(f"Authentication failed: {e}")
    except Exception as e:
        logging.error(f"An error occurred: {e}")

if __name__ == "__main__":
    input_file_path = os.path.join('data_source', 'conta_duplicadas.xlsx')
    main(input_file_path)
