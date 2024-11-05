import os  
import logging
from simple_salesforce import Salesforce
from dotenv import load_dotenv
from list_accounts import list_salesforce_accounts
from process_accounts import load_no_info_accounts, process_all_accounts

# Configuração de logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

def load_credentials():
    """Carrega as credenciais do arquivo .env."""
    load_dotenv()
    username = os.getenv('SALESFORCE_USERNAME')
    password = os.getenv('SALESFORCE_PASSWORD')
    security_token = os.getenv('SALESFORCE_SECURITY_TOKEN')
    domain = os.getenv('SALESFORCE_DOMAIN', 'login')  # 'login' ou 'test'
    
    logger.debug(f"Loaded credentials: Username: {username}, Domain: {domain}")
    return username, password, security_token, domain

def connect_to_salesforce():
    """Conecta ao Salesforce e retorna a sessão.""" 
    username, password, security_token, domain = load_credentials()
    try:
        logger.debug("Tentando conectar ao Salesforce...")
        logger.debug(f"Connecting with Username: {username}, Domain: {domain}")
        
        sf = Salesforce(username=username, password=password, security_token=security_token, domain=domain)
        logger.info("Conexão com Salesforce estabelecida com sucesso!")
        return sf
    except Exception as e:
        logger.error(f"Falha na autenticação: {e}")
        raise

def main():
    """Função principal do script.""" 
    try:
        # Conectar ao Salesforce
        sf = connect_to_salesforce()
        logger.info("Conexão estabelecida com sucesso.")

        # Listar contas e salvar no Excel
        list_salesforce_accounts(sf)

        # Carregar contas com 'NO INFO'
        no_info_accounts = load_no_info_accounts()
        if no_info_accounts is not None and not no_info_accounts.empty:
            # Processar contas com 'NO INFO'
            result_file_no_info = process_all_accounts(sf, only_no_info=True)  # Corrigido aqui
            logger.info("Processamento de contas 'NO INFO' concluído.")
            logger.info(f"Contagens de objetos relacionados salvos em: {result_file_no_info}")
        
        # Processar todas as demais contas
        result_file_all_accounts = process_all_accounts(sf, only_no_info=False)  # Corrigido aqui
        logger.info("Processamento das demais contas concluído.")
        logger.info(f"Contagens de objetos relacionados salvos em: {result_file_all_accounts}")

    except Exception as e:
        logger.error(f"Ocorreu um erro na execução do script: {e}")

if __name__ == "__main__":
    main()
