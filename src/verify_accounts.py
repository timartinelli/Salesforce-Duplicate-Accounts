import pandas as pd
from simple_salesforce import Salesforce

def get_salesforce_ids(sf):
    query = "SELECT Id, Name, NEO_Cpfcnpj__c FROM Account"
    accounts = sf.query_all(query)
    return accounts['records']  # Returns all records

def verify_accounts(sf):
    try:
        print("Verifying accounts...")

        # Step 1: Read the IDs from the Excel file
        excel_file_path = 'data_source/conta_duplicadas.xlsx'
        df = pd.read_excel(excel_file_path)

        # Check if the correct column is present
        if 'Account_18_Digit_ID__c' not in df.columns:
            print("The 'Account_18_Digit_ID__c' column was not found in the file.")
            return

        # Step 2: Check if the accounts exist in Salesforce
        salesforce_accounts = get_salesforce_ids(sf)
        salesforce_ids = {account['Id']: account for account in salesforce_accounts}

        # Step 3: Create a new column to check the existence of the accounts
        df['Account Exists'] = df['Account_18_Digit_ID__c'].apply(lambda x: 'Exists' if x in salesforce_ids else 'Does Not Exist')

        # Step 4: Save the result to a new Excel file
        output_file_path = 'data/conta_duplicadas_verificacao.xlsx'
        df.to_excel(output_file_path, index=False)
        print(f"Results saved to {output_file_path}")

    except Exception as e:
        print(f"An error occurred: {e}")
