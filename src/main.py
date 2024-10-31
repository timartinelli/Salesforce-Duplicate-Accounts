# src/main.py
import sys
import os
import pandas as pd
from simple_salesforce import Salesforce, SalesforceAuthenticationFailed

# Adiciona o diretório atual ao path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from config import USERNAME, PASSWORD, SECURITY_TOKEN, DOMAIN

try:
    sf = Salesforce(username=USERNAME, password=PASSWORD, security_token=SECURITY_TOKEN, domain=DOMAIN)
    print("Connected to Salesforce!")

    # Consulta para listar contas e contar objetos relacionados
    query = """
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
    FROM Account
    """
    
    accounts = sf.query_all(query)

    # Depuração: imprime a resposta da consulta
    print(f"Query Response: {accounts}")

    if 'records' not in accounts or accounts['records'] is None:
        print("No records found in the Salesforce query.")
    else:
        accounts = accounts['records']
        
        # Processamento e contagem dos objetos relacionados
        data = []
        for account in accounts:
            data.append({
                'Account Id': account['Id'],
                'Account Name': account['Name'],
                'Cases Count': len(account.get('Cases', {}).get('records', [])) if account.get('Cases') else 0,
                'Opportunities Count': len(account.get('Opportunities', {}).get('records', [])) if account.get('Opportunities') else 0,
                'Contacts Count': len(account.get('Contacts', {}).get('records', [])) if account.get('Contacts') else 0,
                'Assets Count': len(account.get('Assets', {}).get('records', [])) if account.get('Assets') else 0,
                'Contracts Count': len(account.get('Contracts', {}).get('records', [])) if account.get('Contracts') else 0,
                'Orders Count': len(account.get('Orders', {}).get('records', [])) if account.get('Orders') else 0,
                'Tasks Count': len(account.get('Tasks', {}).get('records', [])) if account.get('Tasks') else 0,
                'Events Count': len(account.get('Events', {}).get('records', [])) if account.get('Events') else 0
            })

        df = pd.DataFrame(data)
        output_folder = 'data'
        os.makedirs(output_folder, exist_ok=True)
        output_file_path = os.path.join(output_folder, 'Salesforce_Accounts_Related_Objects.xlsx')
        df.to_excel(output_file_path, index=False)
        print(f"Data saved to {output_file_path}")

except SalesforceAuthenticationFailed as e:
    print(f"Authentication failed: {e}")
except Exception as e:
    print(f"An error occurred: {e}")
