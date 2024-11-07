# Example of how to describe the Account object
from simple_salesforce import Salesforce
import os

from config import USERNAME, PASSWORD, SECURITY_TOKEN, DOMAIN

# Connect to Salesforce
sf = Salesforce(username=USERNAME, password=PASSWORD, security_token=SECURITY_TOKEN, domain=DOMAIN)

# Describe the Account object
account_description = sf.Account.describe()

# Create the "data" folder if it doesn't exist
output_folder = 'data'
os.makedirs(output_folder, exist_ok=True)

# Output file path
output_file_path = os.path.join(output_folder, 'account_description.txt')

# Save the description to a text file
with open(output_file_path, 'w') as file:
    file.write(str(account_description))

print(f"Account description saved to {output_file_path}")
