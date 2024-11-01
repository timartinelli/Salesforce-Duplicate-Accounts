import sys
import os
import pandas as pd
from simple_salesforce import Salesforce, SalesforceAuthenticationFailed
from verify_duplicates import verify_accounts
import logging

# Adiciona o diretório atual ao path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from config import USERNAME, PASSWORD, SECURITY_TOKEN, DOMAIN

# Configuração de logging com nível de informação para rastrear a execução
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def fetch_related_counts(sf, account_id):
    # Realiza a consulta de relacionamentos de uma conta no Salesforce
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
            return accounts['records'][0]  # Retorna o primeiro registro encontrado
        else:
            logging.warning(f"Nenhum registro encontrado para o ID da conta: {account_id}")
            return {}
    except Exception as e:
        logging.error(f"Erro ao consultar Salesforce para o ID da conta {account_id}: {e}")
        return {}

def check_duplicate_cpf(sf, cpf_cnpj, current_account_id):
    """
    Verifica se o CPF/CNPJ especificado está presente em outras contas no Salesforce.
    Exclui a conta atual da verificação para evitar falsos positivos.
    """
    query = f"""
    SELECT Id FROM Account WHERE NEO_Cpfcnpj__c = '{cpf_cnpj}' AND Id != '{current_account_id}'
    """
    try:
        logging.info(f"Executando consulta para CPF/CNPJ: {cpf_cnpj}")
        accounts = sf.query(query)
        
        if accounts and accounts['records']:
            duplicate_ids = [record['Id'] for record in accounts['records']]
            logging.info(f"CPF/CNPJ {cpf_cnpj} encontrado nas contas: {duplicate_ids}")
            return duplicate_ids
        else:
            logging.info(f"Nenhuma duplicata encontrada para CPF/CNPJ: {cpf_cnpj}")
            return []
    except Exception as e:
        logging.error(f"Erro ao consultar Salesforce para CPF/CNPJ {cpf_cnpj}: {e}")
        return []

def main(input_file_path):
    try:
        logging.info("Arquivos na pasta 'data': %s", os.listdir('data'))

        # Conexão com o Salesforce
        sf = Salesforce(username=USERNAME, password=PASSWORD, security_token=SECURITY_TOKEN, domain=DOMAIN)
        logging.info("Conectado ao Salesforce com sucesso!")

        # Verificação inicial de contas duplicadas
        verification_result = verify_accounts(sf)

        # Carregando o arquivo Excel
        df = pd.read_excel(input_file_path)

        # Adicionando novas colunas para as contagens e para a verificação de duplicidade de CPF/CNPJ
        df['Case Count'] = 0
        df['Opportunity Count'] = 0
        df['Contact Count'] = 0
        df['Asset Count'] = 0
        df['Contract Count'] = 0
        df['Order Count'] = 0
        df['Task Count'] = 0
        df['Event Count'] = 0
        df['Duplicate CPF/CNPJ Check'] = ''  # Coluna para indicar duplicidade de CPF/CNPJ

        # Loop sobre as contas no DataFrame
        for index, row in df.iterrows():
            account_id = row['Account_18_Digit_ID__c']
            cpf_cnpj = row.get('NEO_Cpfcnpj__c')

            # Verifica contas com status 'NO INFO' e checa a existência no Salesforce
            if row['STATUS'] == 'NO INFO' and account_id in verification_result['existing_ids']:
                related_counts = fetch_related_counts(sf, account_id)

                # Validação e extração de contagens para cada relacionamento
                case_records = related_counts.get('Cases', {}).get('records', []) if isinstance(related_counts.get('Cases'), dict) else []
                opportunity_records = related_counts.get('Opportunities', {}).get('records', []) if isinstance(related_counts.get('Opportunities'), dict) else []
                contact_records = related_counts.get('Contacts', {}).get('records', []) if isinstance(related_counts.get('Contacts'), dict) else []
                asset_records = related_counts.get('Assets', {}).get('records', []) if isinstance(related_counts.get('Assets'), dict) else []
                contract_records = related_counts.get('Contracts', {}).get('records', []) if isinstance(related_counts.get('Contracts'), dict) else []
                order_records = related_counts.get('Orders', {}).get('records', []) if isinstance(related_counts.get('Orders'), dict) else []
                task_records = related_counts.get('Tasks', {}).get('records', []) if isinstance(related_counts.get('Tasks'), dict) else []
                event_records = related_counts.get('Events', {}).get('records', []) if isinstance(related_counts.get('Events'), dict) else []

                # Preenchendo contagens nas colunas correspondentes
                df.at[index, 'Case Count'] = len(case_records)
                df.at[index, 'Opportunity Count'] = len(opportunity_records)
                df.at[index, 'Contact Count'] = len(contact_records)
                df.at[index, 'Asset Count'] = len(asset_records)
                df.at[index, 'Contract Count'] = len(contract_records)
                df.at[index, 'Order Count'] = len(order_records)
                df.at[index, 'Task Count'] = len(task_records)
                df.at[index, 'Event Count'] = len(event_records)
                
                logging.info(f"Conta {account_id} processada com contagens: Cases={len(case_records)}, Opportunities={len(opportunity_records)}, Contacts={len(contact_records)}, Assets={len(asset_records)}, Contracts={len(contract_records)}, Orders={len(order_records)}, Tasks={len(task_records)}, Events={len(event_records)}")

            # Verifica duplicatas para contas com status 'DUPLICATE'
            elif row['STATUS'] in ['DUPLICATE', 'DUPLICATE - APPEARS 3X', 'DUPLICATE 5X']:
                # Verifica se o CPF/CNPJ está preenchido
                if pd.notna(cpf_cnpj) and isinstance(cpf_cnpj, str):
                    duplicate_ids = check_duplicate_cpf(sf, cpf_cnpj, account_id)
                    
                    # Preenche a coluna 'Duplicate CPF/CNPJ Check' com o resultado da verificação de duplicidade
                    if duplicate_ids:
                        df.at[index, 'Duplicate CPF/CNPJ Check'] = 'Duplicate found: ' + ', '.join(duplicate_ids)
                        logging.info(f"CPF/CNPJ {cpf_cnpj} encontrado em outras contas para o ID da conta: {account_id}. IDs duplicados: {duplicate_ids}")
                    else:
                        df.at[index, 'Duplicate CPF/CNPJ Check'] = 'No duplicates found'
                        logging.info(f"CPF/CNPJ {cpf_cnpj} não encontrado em outras contas para o ID da conta: {account_id}")
                else:
                    # Marca como 'No CPF/CNPJ' caso o valor esteja ausente ou incorreto
                    df.at[index, 'Duplicate CPF/CNPJ Check'] = 'No CPF/CNPJ'
                    logging.info(f"CPF/CNPJ ausente ou inválido para a conta ID: {account_id}")

        # Salva as atualizações no arquivo Excel
        df.to_excel(input_file_path, index=False)
        logging.info(f"Resultados atualizados e salvos em {input_file_path}")

    except SalesforceAuthenticationFailed as e:
        logging.error(f"Falha na autenticação: {e}")
    except Exception as e:
        logging.error(f"Ocorreu um erro: {e}")

if __name__ == "__main__":
    input_file_path = os.path.join('data', 'conta_duplicadas_verificacao.xlsx')
    main(input_file_path)
