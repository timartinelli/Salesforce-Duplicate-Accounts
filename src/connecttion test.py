import logging
import os
from simple_salesforce import Salesforce
from dotenv import load_dotenv

# Configuração do logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

def connect_to_salesforce():
    # Carrega as variáveis de ambiente do arquivo .env
    load_dotenv()
    
    username = os.getenv('SALESFORCE_USERNAME')
    password = os.getenv('SALESFORCE_PASSWORD')
    security_token = os.getenv('SALESFORCE_SECURITY_TOKEN')
    domain = os.getenv('SALESFORCE_DOMAIN', 'login')  # 'login' para produção, 'test' para sandbox

    # Logando as informações de conexão
    logger.debug("Tentando conectar ao Salesforce...")
    logger.debug(f"Usuário: `{username}`")
    logger.debug(f"Domínio: `{domain}`")
    logger.debug(f"security_token: `{security_token}`")
    logger.debug(f"password: `{password}`")

    # Adiciona informações de segurança (não logar a senha ou token)
    logger.debug("Tentando autenticar...")

    try:
        # Conecta ao Salesforce
        sf = Salesforce(username=username, password=password, security_token=security_token, domain=domain)
        logger.info("Conexão bem-sucedida ao Salesforce!")
        return sf
    except Exception as e:
        logger.error(f"Falha na autenticação: {e}")
        raise

if __name__ == "__main__":
    try:
        sf = connect_to_salesforce()
    except Exception as e:
        logger.error(f"Ocorreu um erro: {e}")
