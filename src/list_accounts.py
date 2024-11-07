import pandas as pd
import logging

def list_salesforce_accounts(sf):
    """Lists all accounts from Salesforce and saves them in an Excel file with ID, Name, and NEO_Cpfcnpj__c."""
    logging.debug("Fetching all accounts from Salesforce.")
    try:
        # Querying all accounts from Salesforce with the NEO_Cpfcnpj__c field
        query = "SELECT Id, Name, NEO_Cpfcnpj__c FROM Account"
        result = sf.query_all(query)
        
        # Extracting data into a list
        accounts_data = [
            {
                "Id": account["Id"],
                "Name": account["Name"],
                "NEO_Cpfcnpj__c": account.get("NEO_Cpfcnpj__c", None)  # Using get to avoid error if the field is empty
            }
            for account in result["records"]
        ]
        
        # Saving data to a DataFrame and exporting to Excel
        df = pd.DataFrame(accounts_data)
        df.to_excel('data/salesforce_accounts.xlsx', index=False)
        
        logging.info("Salesforce accounts saved to 'data/salesforce_accounts.xlsx' with NEO_Cpfcnpj__c.")
    except Exception as e:
        logging.error(f"Failed to list Salesforce accounts: {e}")
        raise
