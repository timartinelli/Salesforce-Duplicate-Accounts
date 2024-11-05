import pandas as pd
import logging

def list_salesforce_accounts(sf):
    """Lists all accounts from Salesforce and saves them in an Excel file with ID, Name, and NEO_Cpfcnpj__c."""
    logging.debug("Fetching all accounts from Salesforce.")
    try:
        # Consultando todas as contas do Salesforce com o campo NEO_Cpfcnpj__c
        query = "SELECT Id, Name, NEO_Cpfcnpj__c FROM Account"
        result = sf.query_all(query)
        
        # Extraindo os dados para uma lista
        accounts_data = [
            {
                "Id": account["Id"],
                "Name": account["Name"],
                "NEO_Cpfcnpj__c": account.get("NEO_Cpfcnpj__c", None)  # Usando get para evitar erro se o campo estiver vazio
            }
            for account in result["records"]
        ]
        
        # Salvando os dados em um DataFrame e exportando para Excel
        df = pd.DataFrame(accounts_data)
        df.to_excel('data/salesforce_accounts.xlsx', index=False)
        
        logging.info("Salesforce accounts saved to 'data/salesforce_accounts.xlsx' with NEO_Cpfcnpj__c.")
    except Exception as e:
        logging.error(f"Failed to list Salesforce accounts: {e}")
        raise
