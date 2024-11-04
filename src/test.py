import os
from dotenv import load_dotenv
import logging

# Configurando logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Carrega variáveis de ambiente do arquivo .env
load_dotenv()

# Listando as variáveis de ambiente carregadas
username = os.getenv('SALESFORCE_USERNAME')
password = os.getenv('SALESFORCE_PASSWORD')
security_token = os.getenv('SALESFORCE_SECURITY_TOKEN')
domain = os.getenv('SALESFORCE_DOMAIN')

# Logando os valores
logging.info(f"USERNAME: {username}")
logging.info(f"PASSWORD: {password}")
logging.info(f"SECURITY_TOKEN: {security_token}")
logging.info(f"DOMAIN: '{domain}'")  # Use aspas para ver espaços
