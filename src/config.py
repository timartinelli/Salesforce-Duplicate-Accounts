# src/config.py
from dotenv import load_dotenv
import os

# Carrega variáveis de ambiente do arquivo .env
load_dotenv()

USERNAME = os.getenv('SALESFORCE_USERNAME')
PASSWORD = os.getenv('SALESFORCE_PASSWORD')
SECURITY_TOKEN = os.getenv('SALESFORCE_SECURITY_TOKEN')
DOMAIN = os.getenv('SALESFORCE_DOMAIN')

# Adicione o nível de log aqui se necessário
import logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


