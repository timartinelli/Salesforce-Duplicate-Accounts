# src/describe_account.py
from simple_salesforce import Salesforce
import os

from config import USERNAME, PASSWORD, SECURITY_TOKEN, DOMAIN

# Connect to Salesforce
sf = Salesforce(username=USERNAME, password=PASSWORD, security_token=SECURITY_TOKEN, domain=DOMAIN)

# Describe the Account object
account_description = sf.Account.describe()

# Save the output to a text file
output_folder = 'data'
os.makedirs(output_folder, exist_ok=True)
output_file_path = os.path.join(output_folder, 'Account_Description.txt')

with open(output_file_path, 'w') as file:
    file.write(str(account_description))

print(f"Account description saved to {output_file_path}")
