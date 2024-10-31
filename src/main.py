import sys
import os
import pandas as pd
import logging
from simple_salesforce import Salesforce, SalesforceAuthenticationFailed
from verify_duplicates import verify_accounts
from config import USERNAME, PASSWORD, SECURITY_TOKEN, DOMAIN

# Configurações iniciais
DATA_DIR = 'data'
INPUT_FILE = os.path.join(DATA_DIR, 'conta_duplicadas_verificacao.xlsx')
LOG_FILE = os.path.join(DATA_DIR, 'script_log.log')

# Configuração do logging
logging.basicConfig(filename=LOG_FILE, level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Função para autenticação no Salesforce
def authenticate_salesforce():
    try:
        sf = Salesforce(username=USERNAME, password=PASSWORD, security_token=SECURITY_TOKEN, domain=DOMAIN)
        logging.info("Connected to Salesforce successfully!")
        return sf
    except SalesforceAuthenticationFailed as e:
        logging.error(f"Salesforce authentication failed: {e}")
        return None

# Função para consultar contagens de objetos relacionados
def fetch_related_counts(sf, account_id: str):
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
            return accounts['records'][0]
        else:
            logging.warning(f"No related records found for account ID: {account_id}")
            return {}
    except Exception as e:
        logging.error(f"Error querying related counts for account {account_id}: {e}")
        return {}

# Função para processar e atualizar o arquivo Excel
def process_excel(sf, input_file: str):
    df = pd.read_excel(input_file)

    # Adiciona as colunas para contagens, se ainda não existirem
    columns = ['Case Count', 'Opportunity Count', 'Contact Count', 'Asset Count', 
               'Contract Count', 'Order Count', 'Task Count', 'Event Count']
    for column in columns:
        if column not in df.columns:
            df[column] = 0

    # Verificação e preenchimento das contagens
    verification_result = verify_accounts(sf)
    for index, row in df.iterrows():
        account_id = row['Account_18_Digit_ID__c']
        if row['STATUS'] == 'NO INFO' and account_id in verification_result['existing_ids']:
            related_counts = fetch_related_counts(sf, account_id)
            
            # Obtenção das contagens com verificação de existência
            case_records = related_counts.get('Cases', {}).get('records', []) if isinstance(related_counts.get('Cases'), dict) else []
            opportunity_records = related_counts.get('Opportunities', {}).get('records', []) if isinstance(related_counts.get('Opportunities'), dict) else []
            contact_records = related_counts.get('Contacts', {}).get('records', []) if isinstance(related_counts.get('Contacts'), dict) else []
            asset_records = related_counts.get('Assets', {}).get('records', []) if isinstance(related_counts.get('Assets'), dict) else []
            contract_records = related_counts.get('Contracts', {}).get('records', []) if isinstance(related_counts.get('Contracts'), dict) else []
            order_records = related_counts.get('Orders', {}).get('records', []) if isinstance(related_counts.get('Orders'), dict) else []
            task_records = related_counts.get('Tasks', {}).get('records', []) if isinstance(related_counts.get('Tasks'), dict) else []
            event_records = related_counts.get('Events', {}).get('records', []) if isinstance(related_counts.get('Events'), dict) else []

            # Atualização das colunas no DataFrame
            df.at[index, 'Case Count'] = len(case_records)
            df.at[index, 'Opportunity Count'] = len(opportunity_records)
            df.at[index, 'Contact Count'] = len(contact_records)
            df.at[index, 'Asset Count'] = len(asset_records)
            df.at[index, 'Contract Count'] = len(contract_records)
            df.at[index, 'Order Count'] = len(order_records)
            df.at[index, 'Task Count'] = len(task_records)
            df.at[index, 'Event Count'] = len(event_records)
            
            logging.info(f"Processed account {account_id} with counts: "
                         f"Cases={len(case_records)}, Opportunities={len(opportunity_records)}, "
                         f"Contacts={len(contact_records)}, Assets={len(asset_records)}, "
                         f"Contracts={len(contract_records)}, Orders={len(order_records)}, "
                         f"Tasks={len(task_records)}, Events={len(event_records)}")

    # Salvar o DataFrame atualizado
    df.to_excel(input_file, index=False)
    logging.info(f"Updated results saved to {input_file}")

# Função principal para execução do script
def main():
    sf = authenticate_salesforce()
    if sf:
        process_excel(sf, INPUT_FILE)
    else:
        logging.error("Failed to connect to Salesforce. Exiting script.")

if __name__ == "__main__":
    main()
