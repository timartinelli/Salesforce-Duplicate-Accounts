import pandas as pd

def check_duplicates(sf, salesforce_accounts):
    print("Checking for duplicate NEO_Cpfcnpj__c...")

    # Step 1: Collect NEO_Cpfcnpj__c from accounts
    cpf_data = {}
    for account in salesforce_accounts:
        cpf_value = account.get('NEO_Cpfcnpj__c')
        if cpf_value:
            if cpf_value in cpf_data:
                cpf_data[cpf_value].append(account['Id'])
            else:
                cpf_data[cpf_value] = [account['Id']]
    
    # Step 2: Filter duplicates
    duplicates = {cpf: ids for cpf, ids in cpf_data.items() if len(ids) > 1}
    
    # Step 3: Create file with duplicated accounts
    if duplicates:
        df_duplicates = pd.DataFrame([(cpf, id) for cpf, ids in duplicates.items() for id in ids], columns=['NEO_Cpfcnpj__c', 'Account ID'])
        output_file_path = 'data/duplicated_accounts.xlsx'
        df_duplicates.to_excel(output_file_path, index=False)
        print(f"Duplicated accounts saved to {output_file_path}")
    else:
        print("No duplicated accounts found.")
