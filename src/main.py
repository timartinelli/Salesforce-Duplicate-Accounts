import os  
import logging
from simple_salesforce import Salesforce
from dotenv import load_dotenv
from list_accounts import list_salesforce_accounts
from process_accounts import load_no_info_accounts, process_all_accounts

# Logging configuration
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

def load_credentials():
    """Load credentials from the .env file."""
    load_dotenv()
    username = os.getenv('SALESFORCE_USERNAME')
    password = os.getenv('SALESFORCE_PASSWORD')
    security_token = os.getenv('SALESFORCE_SECURITY_TOKEN')
    domain = os.getenv('SALESFORCE_DOMAIN', 'login')  # 'login' or 'test'
    
    logger.debug(f"Loaded credentials: Username: {username}, Domain: {domain}")
    return username, password, security_token, domain

def connect_to_salesforce():
    """Connect to Salesforce and return the session.""" 
    username, password, security_token, domain = load_credentials()
    try:
        logger.debug("Attempting to connect to Salesforce...")
        logger.debug(f"Connecting with Username: {username}, Domain: {domain}")
        
        sf = Salesforce(username=username, password=password, security_token=security_token, domain=domain)
        logger.info("Connection to Salesforce established successfully!")
        return sf
    except Exception as e:
        logger.error(f"Authentication failed: {e}")
        raise

def main():
    """Main function of the script.""" 
    try:
        # Connect to Salesforce
        sf = connect_to_salesforce()
        logger.info("Connection established successfully.")

        # List accounts and save to Excel
        list_salesforce_accounts(sf)

        # Load accounts with 'NO INFO'
        no_info_accounts = load_no_info_accounts()
        if no_info_accounts is not None and not no_info_accounts.empty:
            # Process accounts with 'NO INFO'
            result_file_no_info = process_all_accounts(sf, only_no_info=True)  # Fixed here
            logger.info("Processing 'NO INFO' accounts completed.")
            logger.info(f"Related object counts saved to: {result_file_no_info}")
        
        # Process all other accounts
        result_file_all_accounts = process_all_accounts(sf, only_no_info=False)  # Fixed here
        logger.info("Processing other accounts completed.")
        logger.info(f"Related object counts saved to: {result_file_all_accounts}")

    except Exception as e:
        logger.error(f"An error occurred during script execution: {e}")

if __name__ == "__main__":
    main()
