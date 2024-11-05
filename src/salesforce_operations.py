# salesforce_operations.py
from simple_salesforce import Salesforce

def get_salesforce_ids(sf):
    """Retorna um dicionário com IDs, Nomes e NEO_Cpfcnpj__c das contas no Salesforce."""
    try:
        # Consulta para obter ID, Nome e NEO_Cpfcnpj__c
        query = "SELECT Id, Name, NEO_Cpfcnpj__c FROM Account"
        accounts = sf.query_all(query)

        # Cria um dicionário com ID como chave e um dicionário como valor
        return {
            account['Id']: {
                'Name': account['Name'],
                'NEO_Cpfcnpj__c': account.get('NEO_Cpfcnpj__c')  # Usando get() para evitar KeyError se o campo não existir
            }
            for account in accounts['records']
        }

    except Exception as e:
        print(f"An error occurred while fetching accounts from Salesforce: {e}")
        return {}
