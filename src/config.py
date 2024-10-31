# src/config.py
from dotenv import load_dotenv
import os

# Carrega variáveis de ambiente do arquivo .env
load_dotenv()

USERNAME = os.getenv('SALESFORCE_USERNAME')
PASSWORD = os.getenv('SALESFORCE_PASSWORD')
SECURITY_TOKEN = os.getenv('SALESFORCE_SECURITY_TOKEN')
DOMAIN = os.getenv('SALESFORCE_DOMAIN')
