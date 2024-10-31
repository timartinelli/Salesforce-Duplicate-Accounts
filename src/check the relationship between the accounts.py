# Exemplo de como descrever o objeto Account
from simple_salesforce import Salesforce
import os

from config import USERNAME, PASSWORD, SECURITY_TOKEN, DOMAIN

# Conexão com o Salesforce
sf = Salesforce(username=USERNAME, password=PASSWORD, security_token=SECURITY_TOKEN, domain=DOMAIN)

# Descrevendo o objeto Account
account_description = sf.Account.describe()

# Criar a pasta "data" se não existir
output_folder = 'data'
os.makedirs(output_folder, exist_ok=True)

# Caminho do arquivo de saída
output_file_path = os.path.join(output_folder, 'account_description.txt')

# Salvar a descrição em um arquivo de texto
with open(output_file_path, 'w') as file:
    file.write(str(account_description))

print(f"Account description saved to {output_file_path}")
