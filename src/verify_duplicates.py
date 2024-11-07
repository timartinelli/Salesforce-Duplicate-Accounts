import os
import pandas as pd
from simple_salesforce import Salesforce, SalesforceAuthenticationFailed

def get_salesforce_ids(sf):
    query = "SELECT Id, Name FROM Account"
    accounts = sf.query_all(query)
    return {account['Id']: account['Name'] for account in accounts['records']}  # Returns a dictionary with ID and Name

def verify_accounts(sf):
    try:
        print("Verifying accounts...")

        # Step 1: Export all IDs and account names from Salesforce
        salesforce_accounts = get_salesforce_ids(sf)
        salesforce_ids = set(salesforce_accounts.keys())  # Only the IDs

        # Step 2: Read the IDs from the Excel file
        excel_file_path = 'data_source/conta_duplicadas.xlsx'
        df = pd.read_excel(excel_file_path)

        # Check if the correct column is present
        if 'Account_18_Digit_ID__c' not in df.columns:
            print("The 'Account_18_Digit_ID__c' column was not found in the file.")
            return

        # Step 3: Create a new column to check the existence of the accounts
        df['Account Exists'] = df['Account_18_Digit_ID__c'].apply(lambda x: 'Exists' if x in salesforce_ids else 'Does Not Exist')

        # Step 4: Save the results in a new Excel file
        output_file_path = 'data/conta_duplicadas_verificacao.xlsx'
        df.to_excel(output_file_path, index=False)
        print(f"Results saved to {output_file_path}")

        # Step 5: Identify new accounts that are in Salesforce but not in the spreadsheet
        new_accounts = {id_: name for id_, name in salesforce_accounts.items() if id_ not in df['Account_18_Digit_ID__c'].values}

        # Step 6: Create a DataFrame for new accounts
        new_accounts_df = pd.DataFrame(new_accounts.items(), columns=['Account ID', 'Account Name'])

        # Save new accounts in a new Excel file
        new_accounts_file_path = 'data/new_accounts.xlsx'
        new_accounts_df.to_excel(new_accounts_file_path, index=False)
        print(f"New accounts saved to {new_accounts_file_path}")

        return {'existing_ids': salesforce_ids}

    except SalesforceAuthenticationFailed as e:
        print(f"Authentication failed: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    pass  # The code should not be executed directly
