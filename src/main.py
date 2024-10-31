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
        if accounts and accounts['records']:
            return accounts['records'][0]
        else:
            print(f"[DEBUG] No records found for account ID: {account_id}")
            return {}  # Retorna um dicionário vazio
    except Exception as e:
        print(f"[ERROR] Error querying Salesforce for account ID {account_id}: {e}")
        return {}

def process_excel(sf, input_file_path):
    try:
        print(f"[INFO] Processing file: {input_file_path}")
        
        # Carregar dados do Excel
        df = pd.read_excel(input_file_path)

        # Adicionar colunas para contagens de objetos relacionados
        relation_columns = [
            'Case Count', 'Opportunity Count', 'Contact Count', 'Asset Count',
            'Contract Count', 'Order Count', 'Task Count', 'Event Count'
        ]
        for col in relation_columns:
            df[col] = 0

        verification_result = verify_accounts(sf)
        for index, row in df.iterrows():
            account_id = row.get('Account_18_Digit_ID__c')
            status = row.get('STATUS', '')

            if status == 'NO INFO' and account_id in verification_result['existing_ids']:
                related_counts = fetch_related_counts(sf, account_id)

                if not isinstance(related_counts, dict) or not related_counts:
                    print(f"[WARNING] No related data for account ID {account_id}. Related counts: {related_counts}")
                    continue

                try:
                    # Obter contagens de registros de cada relacionamento
                    case_records = related_counts.get('Cases', {}).get('records', [])
                    opportunity_records = related_counts.get('Opportunities', {}).get('records', [])
                    contact_records = related_counts.get('Contacts', {}).get('records', [])
                    asset_records = related_counts.get('Assets', {}).get('records', [])
                    contract_records = related_counts.get('Contracts', {}).get('records', [])
                    order_records = related_counts.get('Orders', {}).get('records', [])
                    task_records = related_counts.get('Tasks', {}).get('records', [])
                    event_records = related_counts.get('Events', {}).get('records', [])

                    # Preencher contagens nas colunas
                    df.at[index, 'Case Count'] = len(case_records)
                    df.at[index, 'Opportunity Count'] = len(opportunity_records)
                    df.at[index, 'Contact Count'] = len(contact_records)
                    df.at[index, 'Asset Count'] = len(asset_records)
                    df.at[index, 'Contract Count'] = len(contract_records)
                    df.at[index, 'Order Count'] = len(order_records)
                    df.at[index, 'Task Count'] = len(task_records)
                    df.at[index, 'Event Count'] = len(event_records)

                    print(f"[DEBUG] Processed account {account_id} with counts: Cases={len(case_records)}, Opportunities={len(opportunity_records)}, Contacts={len(contact_records)}, Assets={len(asset_records)}, Contracts={len(contract_records)}, Orders={len(order_records)}, Tasks={len(task_records)}, Events={len(event_records)}")
                except Exception as inner_e:
                    print(f"[ERROR] Error processing related counts for account ID {account_id}: {inner_e}")
                    print(f"[DEBUG] Related counts data: {related_counts}")  # Exibe o conteúdo completo em caso de erro
            else:
                print(f"[INFO] Account {account_id} skipped (STATUS: {status})")

        # Salvar o DataFrame atualizado de volta ao Excel
        df.to_excel(input_file_path, index=False)
        print(f"[INFO] Updated results saved to {input_file_path}")

    except FileNotFoundError as fnf_error:
        print(f"[ERROR] File not found: {fnf_error}")
    except Exception as e:
        print(f"[ERROR] An error occurred during Excel processing: {e}")

def main():
    input_file_path = os.path.join('data', 'conta_duplicadas_verificacao.xlsx')
    try:
        print(f"[INFO] Listing files in 'data' directory: {os.listdir('data')}")
        
        sf = Salesforce(username=USERNAME, password=PASSWORD, security_token=SECURITY_TOKEN, domain=DOMAIN)
        print("[INFO] Connected to Salesforce successfully!")
        
        # Processa o arquivo Excel
        process_excel(sf, input_file_path)
    except SalesforceAuthenticationFailed as auth_error:
        print(f"[ERROR] Salesforce authentication failed: {auth_error}")
    except Exception as e:
        print(f"[ERROR] Unexpected error in main: {e}")

if __name__ == "__main__":
    main()
