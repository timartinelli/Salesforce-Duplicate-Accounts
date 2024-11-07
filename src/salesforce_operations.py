# operacoes_salesforce.py
from simple_salesforce import Salesforce

def obter_ids_salesforce(sf):
    """Retorna um dicionário com IDs, Nomes e NEO_Cpfcnpj__c das contas no Salesforce."""
    try:
        # Consulta para obter ID, Nome e NEO_Cpfcnpj__c
        consulta = "SELECT Id, Name, NEO_Cpfcnpj__c FROM Account"
        contas = sf.query_all(consulta)

        # Cria um dicionário com ID como chave e um dicionário como valor
        return {
            conta['Id']: {
                'Nome': conta['Name'],
                'NEO_Cpfcnpj__c': conta.get('NEO_Cpfcnpj__c')  # Usando get() para evitar KeyError se o campo não existir
            }
            for conta in contas['records']
        }

    except Exception as e:
        print(f"Ocorreu um erro ao buscar contas do Salesforce: {e}")
        return {}
