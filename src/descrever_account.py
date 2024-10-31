# src/descrever_account.py
from simple_salesforce import Salesforce
import os

from config import USERNAME, PASSWORD, SECURITY_TOKEN, DOMAIN

# Conectar ao Salesforce
sf = Salesforce(username=USERNAME, password=PASSWORD, security_token=SECURITY_TOKEN, domain=DOMAIN)

# Descrever o objeto Account
account_description = sf.Account.describe()

# Salvar a saída em um arquivo de texto
output_folder = 'data'
os.makedirs(output_folder, exist_ok=True)
output_file_path = os.path.join(output_folder, 'Account_Description.txt')

with open(output_file_path, 'w') as file:
    file.write(str(account_description))

print(f"Account description saved to {output_file_path}")
