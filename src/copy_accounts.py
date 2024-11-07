import logging
import os
from simple_salesforce import Salesforce
from dotenv import load_dotenv

# Logging configuration
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

def connect_to_salesforce():
    # Load environment variables from the .env file
    load_dotenv()
    
    username = os.getenv('SALESFORCE_USERNAME')
    password = os.getenv('SALESFORCE_PASSWORD')
    security_token = os.getenv('SALESFORCE_SECURITY_TOKEN')
    domain = os.getenv('SALESFORCE_DOMAIN', 'login')  # 'login' for production, 'test' for sandbox

    # Log the connection details
    logger.debug("Attempting to connect to Salesforce...")
    logger.debug(f"Username: `{username}`")
    logger.debug(f"Domain: `{domain}`")
    logger.debug(f"Security Token: `{security_token}`")
    logger.debug(f"Password: `{password}`")

    # Log sensitive information (do not log password or token)
    logger.debug("Attempting to authenticate...")

    try:
        # Connect to Salesforce
        sf = Salesforce(username=username, password=password, security_token=security_token, domain=domain)
        logger.info("Successfully connected to Salesforce!")
        return sf
    except Exception as e:
        logger.error(f"Authentication failed: {e}")
        raise

if __name__ == "__main__":
    try:
        sf = connect_to_salesforce()
    except Exception as e:
        logger.error(f"An error occurred: {e}")

