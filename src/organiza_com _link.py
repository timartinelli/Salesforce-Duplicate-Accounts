import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import logging
import os

from main import connect_to_salesforce  # Ensure this connection function is defined correctly

# Logging configuration
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Input and output files
input_file = 'data_source/conta_Somente_duplicadas.xlsx'  # Source file in the data_source folder
output_file = 'data/contas_duplicadas_com_cores.xlsx'  # Output file with colors
debug_folder = 'data/debug/'  # Folder to store debug files

# Create the debug folder if it doesn't exist
if not os.path.exists(debug_folder):
    os.makedirs(debug_folder)

def load_duplicate_ids(filename):
    """Loads the duplicate account IDs from an Excel file."""
    try:
        df_ids = pd.read_excel(filename)
        logger.debug(f"Loaded IDs: {df_ids['Id'].dropna().tolist()}")  # Adding log for IDs
        return df_ids['Id'].dropna().tolist()  # Now using the 'Id' field correctly
    except Exception as e:
        logger.error(f"Error loading duplicate IDs: {e}")
        raise

def fetch_accounts_by_ids(sf, ids_list):
    """Queries all accounts in Salesforce for the specified IDs and returns the records."""
    try:
        fields = ', '.join([f['name'] for f in sf.Account.describe()['fields']])
        query = f"SELECT {fields} FROM Account WHERE Id IN ({', '.join([f"'{id}'" for id in ids_list])})"
        logger.info("Fetching duplicate accounts from Salesforce...")
        result = sf.query_all(query)['records']
        
        # Limit the number of accounts for logging (showing the first 5)
        logger.debug(f"Accounts returned from Salesforce (showing the first 5): {result[:5]}")
        
        return result
    except Exception as e:
        logger.error(f"Error fetching accounts: {e}")
        raise

def count_non_empty_fields(account):
    """Counts how many fields are not empty in an account."""
    return sum(1 for value in account.values() if value not in [None, '', 'NA'])

def choose_base_account(accounts, group_key):
    """Chooses the base account (the one that will be kept) among several duplicate accounts."""
    logger.debug(f"Accounts to be compared: {accounts}")  # Log to see the passed accounts

    if not accounts:
        logger.warning(f"No account to compare in group {group_key}!")
        return None, []

    base_account = accounts[0]
    discarded_account_ids = []  # IDs of discarded accounts

    # Sort accounts by modification date (most recent first)
    accounts = sorted(accounts, key=lambda x: x.get('LastModifiedDate', ''), reverse=True)

    logger.debug(f"Sorted accounts for group {group_key}: {accounts}")  # Log to check sorted accounts

    for account in accounts[1:]:
        last_modified_base = base_account.get('LastModifiedDate')
        last_modified_account = account.get('LastModifiedDate')

        logger.debug(f"Comparing base account (ID {base_account['Id']}) and duplicate account (ID {account['Id']})")

        # Pause to review the information before proceeding
        input(f"Comparing base account (ID {base_account['Id']}) with duplicate account (ID {account['Id']}). Press Enter to continue...")

        # If the current account's modification date is more recent than the base account's, replace it
        if last_modified_account > last_modified_base:
            discarded_account_ids.append(base_account['Id'])
            base_account = account  # Update the base account
            logger.debug(f"Base account updated to: {base_account['Id']}")

        elif last_modified_account == last_modified_base:
            # If the dates are the same, choose the account with more filled fields
            count_base = count_non_empty_fields(base_account)
            count_account = count_non_empty_fields(account)
            
            logger.debug(f"Filled field counts for base account: {count_base}, duplicate account: {count_account}")

            # Pause to review the information before proceeding
            input(f"Filled field counts for base account ({count_base}) and duplicate account ({count_account}). Press Enter to continue...")

            if count_account > count_base:
                discarded_account_ids.append(base_account['Id'])
                base_account = account  # Update the base account
                logger.debug(f"Base account updated to: {base_account['Id']}")

        # Now, fill empty fields in the base account with data from the duplicate account
        for field in base_account.keys():
            value_base = base_account.get(field)
            value_account = account.get(field)
            
            # If the field in the base account is empty or 'NA', use the value from the duplicate account
            if value_base in [None, '', 'NA']:
                base_account[field] = value_account  # Fill with the duplicate account's value

    base_account['Discarded_Account_Ids'] = discarded_account_ids  # Add the discarded account IDs
    
    # Save the group of accounts to a debug text file
    debug_filename = os.path.join(debug_folder, f"group_{group_key}_debug.txt")
    with open(debug_filename, 'w') as f:
        f.write(f"Group: {group_key}\n")
        f.write(f"Compared accounts:\n")
        for account in accounts:
            f.write(f"Account ID {account['Id']} - Last Modified: {account.get('LastModifiedDate')}\n")
        f.write(f"Base account chosen: {base_account['Id']}\n")
        f.write(f"Discarded accounts: {discarded_account_ids}\n")

    return base_account, discarded_account_ids  # Return the base account and discarded account IDs

def consolidate_accounts(df):
    """Consolidates multiple duplicate accounts into a single base account and applies color formatting."""
    consolidated_accounts = []

    # Clean values in the grouping columns
    logger.debug("Cleaning values in NEO_Cpfcnpj__c and NEO_Clave_Cliente__c columns before grouping...")
    df['NEO_Cpfcnpj__c'] = df['NEO_Cpfcnpj__c'].str.strip()  # Remove extra spaces
    df['NEO_Clave_Cliente__c'] = df['NEO_Clave_Cliente__c'].str.strip()  # Remove extra spaces

    # Check unique values in the grouping columns
    logger.debug(f"Unique values in NEO_Cpfcnpj__c after cleaning: {df['NEO_Cpfcnpj__c'].unique()}")
    logger.debug(f"Unique values in NEO_Clave_Cliente__c after cleaning: {df['NEO_Clave_Cliente__c'].unique()}")

    # Group accounts by NEO_Cpfcnpj__c and NEO_Clave_Cliente__c
    grouped_accounts = df.groupby(['NEO_Cpfcnpj__c', 'NEO_Clave_Cliente__c'])

    logger.debug(f"Account groups to be consolidated: {grouped_accounts}")  # Log to check the grouping

    for group_key, group in grouped_accounts:
        accounts = group.to_dict('records')
        logger.debug(f"Account group: {accounts}")  # Log to check the groups
        base_account, discarded_ids = choose_base_account(accounts, group_key)

        if base_account is None:
            logger.warning(f"No base account for group: {group_key}")
            continue

        # Add the base account to the consolidated list
        consolidated_accounts.append(base_account)

    # Convert to final DataFrame
    df_consolidated = pd.DataFrame(consolidated_accounts)

    return df_consolidated

def save_to_excel(df):
    """Applies color formatting to the Excel file."""
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Consolidated Accounts')

        wb = load_workbook(output_file)
        ws = wb['Consolidated Accounts']

        green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # Apply color formatting
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            account_id = row[0].value  # Account ID
            discarded = row[0].offset(0, 16).value  # Discarded accounts column (based on the index)
            if discarded:  # If the account was discarded
                row[0].fill = yellow_fill
            else:  # If it is the base account
                row[0].fill = green_fill

        wb.save(output_file)
        logger.info(f"File with consolidated accounts saved at {output_file}")

    except Exception as e:
        logger.error(f"Error saving Excel file: {e}")
        raise

if __name__ == "__main__":
    try:
        # Load the duplicate account IDs
        ids_list = load_duplicate_ids(input_file)
        
        # Connect to Salesforce
        sf = connect_to_salesforce()

        # Fetch duplicate accounts
        accounts = fetch_accounts_by_ids(sf, ids_list)
        
        # Consolidate duplicate accounts
        consolidated_df = consolidate_accounts(pd.DataFrame(accounts))
        
        # Save the result with colors
        save_to_excel(consolidated_df)
    except Exception as e:
        logger.error(f"An error occurred: {e}")
