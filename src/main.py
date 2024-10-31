import sys
import os
import pandas as pd
from simple_salesforce import Salesforce, SalesforceAuthenticationFailed
from verify_duplicates import verify_accounts

# Adiciona o diretório atual ao path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from config import USERNAME, PASSWORD, SECURITY_TOKEN, DOMAIN

def fetch_related_counts(sf, account_id):
    # Consulta para contar objetos relacionados
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
        
        # Verifica se a consulta retornou registros
        if accounts and accounts['records']:
            return accounts['records'][0]  # Retorna o primeiro registro
        else:
            print(f"No records found for account ID: {account_id}")  # Mensagem de depuração
            return {}  # Retorna um dicionário vazio para evitar erros
    except Exception as e:
        print(f"Error querying Salesforce for account ID {account_id}: {e}")
        return {}

def main(input_file_path):
    try:
        # Listar arquivos na pasta 'data'
        print("Arquivos na pasta 'data':", os.listdir('data'))

        sf = Salesforce(username=USERNAME, password=PASSWORD, security_token=SECURITY_TOKEN, domain=DOMAIN)
        print("Connected to Salesforce!")

        # Verifica as contas e salva o resultado
        verification_result = verify_accounts(sf)

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

        # Loop através das contas
        for index, row in df.iterrows():
            account_id = row['Account_18_Digit_ID__c']
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
                    
                print(f"Processed account {account_id} with related counts: Cases={len(case_records)}, Opportunities={len(opportunity_records)}, Contacts={len(contact_records)}, Assets={len(asset_records)}, Contracts={len(contract_records)}, Orders={len(order_records)}, Tasks={len(task_records)}, Events={len(event_records)}")
            else:
                print(f"No related counts found or status was not 'NO INFO' for account ID: {account_id}")

        # Salvar o DataFrame atualizado de volta ao Excel
        df.to_excel(input_file_path, index=False)
        print(f"Updated results saved to {input_file_path}")

    except SalesforceAuthenticationFailed as e:
        print(f"Authentication failed: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    input_file_path = os.path.join('data', 'conta_duplicadas_verificacao.xlsx')
    main(input_file_path)
