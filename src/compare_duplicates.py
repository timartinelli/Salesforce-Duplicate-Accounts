import pandas as pd
import logging
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from main import connect_to_salesforce  # Function to connect to Salesforce imported from main.py

# Logging setup
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Input and output files
input_file = 'data_source/conta_Somente_duplicadas.xlsx'  # Source file in data_source folder
output_file = 'data/duplicated_accounts_consolidated_with_colors.xlsx'  # Output file with colors

def load_duplicate_ids(filename):
    """Loads the duplicate account IDs from an Excel file."""
    try:
        df_ids = pd.read_excel(filename)
        return df_ids['Id'].dropna().tolist()  # Correctly using the 'Id' field
    except Exception as e:
        logger.error(f"Error loading duplicate IDs: {e}")
        raise

def fetch_accounts_by_ids(sf, ids_list):
    """Queries Salesforce for the specified account IDs and returns the records."""
    try:
        fields = ', '.join([f['name'] for f in sf.Account.describe()['fields']])
        query = f"SELECT {fields} FROM Account WHERE Id IN ({', '.join([f"'{id}'" for id in ids_list])})"
        logger.info("Fetching duplicate accounts from Salesforce...")
        return sf.query_all(query)['records']
    except Exception as e:
        logger.error(f"Error fetching accounts: {e}")
        raise

def count_non_empty_fields(account):
    """Counts how many fields in an account are not empty."""
    return sum(1 for value in account.values() if value not in [None, '', 'NA'])

def choose_base_account(accounts):
    """Chooses the base account (the one to keep) among several duplicate accounts."""
    base_account = accounts[0]
    discarded_account_ids = []  # Discarded account IDs

    # Sort accounts by last modified date (most recent first)
    accounts = sorted(accounts, key=lambda x: x.get('LastModifiedDate', ''), reverse=True)

    for account in accounts[1:]:
        last_modified_base = base_account.get('LastModifiedDate')
        last_modified_account = account.get('LastModifiedDate')

        # If the current account's last modified date is more recent than the base account's, replace
        if last_modified_account > last_modified_base:
            discarded_account_ids.append(base_account['Id'])
            base_account = account  # Update the base account
        elif last_modified_account == last_modified_base:
            # If the dates are the same, choose the account with more fields filled
            count_base = count_non_empty_fields(base_account)
            count_account = count_non_empty_fields(account)
            
            if count_account > count_base:
                discarded_account_ids.append(base_account['Id'])
                base_account = account  # Update the base account

        # Now, fill empty fields in the base account with data from the secondary account
        for field in base_account.keys():
            value_base = base_account.get(field)
            value_account = account.get(field)
            
            # If the field in the base account is empty or 'NA', use the value from the secondary account
            if value_base in [None, '', 'NA']:
                base_account[field] = value_account  # Fill with the value from the duplicate account

    base_account['Discarded_Account_Ids'] = discarded_account_ids  # Add discarded account IDs
    return base_account, discarded_account_ids  # Return the base account and the discarded IDs

def consolidate_accounts(df):
    """Consolidates multiple duplicate accounts into one base account and applies color formatting."""
    consolidated_accounts = []
    discarded_account_ids = []

    # Group by account ID
    grouped_accounts = df.groupby('Id')

    for account_id, group in grouped_accounts:
        accounts = group.to_dict('records')
        base_account, discarded_ids = choose_base_account(accounts)

        # Add the base account to the consolidated accounts list
        consolidated_accounts.append(base_account)

        # Add the discarded account IDs to the general list
        discarded_account_ids.extend(discarded_ids)

    # Convert to final DataFrame
    df_consolidated = pd.DataFrame(consolidated_accounts)

    return df_consolidated

def apply_formatting_to_excel(df):
    """Applies color formatting to the Excel file, marking consolidated accounts in green and duplicates in yellow."""
    # Load the Excel file for editing
    wb = load_workbook(output_file)
    ws = wb.active

    # Define color styles
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Loop through the accounts and apply colors
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        account_id_cell = row[0]  # Assuming the account ID is in the first column
        discarded_id_cell = row[6]  # The 'Discarded_Account_Ids' column is the 7th (index 6)

        if discarded_id_cell.value == "":  # Account was not discarded
            account_id_cell.fill = green_fill  # Base account in green
        else:
            account_id_cell.fill = yellow_fill  # Duplicate account in yellow

    # Save the file with the formatting applied
    wb.save(output_file)
    logger.info(f"Consolidated accounts with colors saved to: {output_file}")

def main():
    """Main function that connects to Salesforce, loads duplicates, and performs automatic consolidation."""
    try:
        # Connect to Salesforce
        sf = connect_to_salesforce()
        
        # Load duplicate account IDs
        ids_list = load_duplicate_ids(input_file)
        logger.info(f"{len(ids_list)} duplicate IDs loaded for comparison.")
        
        # Fetch duplicate accounts from Salesforce
        accounts = fetch_accounts_by_ids(sf, ids_list)
        df_accounts = pd.DataFrame(accounts)

        # Consolidate duplicate accounts
        df_consolidated = consolidate_accounts(df_accounts)
        
        # Save the consolidated result to an Excel file
        df_consolidated.to_excel(output_file, index=False)
        logger.info(f'Duplicated accounts consolidated and saved to: {output_file}')
        
        # Apply color formatting to the Excel file
        apply_formatting_to_excel(df_consolidated)
        
    except Exception as e:
        logger.error(f"An error occurred during execution: {e}")

if __name__ == "__main__":
    main()
