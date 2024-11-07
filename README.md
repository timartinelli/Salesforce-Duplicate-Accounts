# Salesforce Duplicate Accounts Project

This project aims to automate the detection and organization of duplicate accounts in Salesforce, leveraging Python scripts to interact with Excel files and Salesforce data. By identifying and grouping duplicates based on key identifiers like `NEO_Cpfcnpj__c`, we facilitate a streamlined process for account cleanup, enhancing data quality and supporting better business decisions.

## Project Structure

### `src/`
Contains all Python scripts:

- **check_duplicates.py**: Identifies duplicate accounts based on specific fields (e.g., `NEO_Cpfcnpj__c`).
- **compare_duplicates.py**: Compares similarities and differences between duplicate records in Salesforce.
- **config.py**: Stores configuration parameters, including Salesforce credentials.
- **list_accounts.py**: Retrieves a list of all accounts from Salesforce and saves it to an Excel file.
- **process_accounts.py**: Loads and processes accounts, checking related objects in Salesforce.
- **salesforce_connection.py**: Manages connection to the Salesforce API.
- **verify_duplicates.py**: Verifies the existence of duplicate accounts in the Excel file and provides details.

### `data/`
Contains generated files from data processing, such as:

- Grouped duplicates and analysis results.

### `data_source/`
Stores source files, including original account data for processing.

### `.env`
Stores environment variables for sensitive data like Salesforce API credentials (excluded from Git).

## Technologies Used

- **Python**: Used for automation and data processing.
- **Salesforce**: Cloud-based CRM platform for account management and related data.
- **Excel**: Input and output files for processing and tracking account data.
- **SQL**: Used for querying related Salesforce objects such as Opportunities, Contacts, and Activities.

### Directory Structure

```plaintext
Salesforce-Duplicate-Accounts/
│
├── src/                         # Contains all Python scripts
│   ├── check_duplicates.py       # Identifies duplicate accounts
│   ├── compare_duplicates.py     # Compares duplicate records in Salesforce
│   ├── config.py                # Stores configuration parameters (Salesforce credentials)
│   ├── list_accounts.py         # Retrieves a list of accounts from Salesforce and saves it to an Excel file
│   ├── process_accounts.py      # Processes accounts and checks related objects in Salesforce
│   ├── salesforce_connection.py # Manages connection to the Salesforce API
│   └── verify_duplicates.py     # Verifies the existence of duplicate accounts in the Excel file and provides details
│
├── data/                        # Contains generated files from data processing
│   ├── conta_duplicadas_verificacao.xlsx   # Result of verified accounts
│   ├── novas_contas.xlsx        # File with new accounts found in Salesforce
│   └── ...                      # Other analysis and result files
│
├── data_source/                 # Stores source files, such as original account data for processing
│   └── conta_duplicadas.xlsx    # Input file with accounts for processing
│
├── .env                         # Environment file with sensitive data like Salesforce API credentials (excluded from Git)
├── requirements.txt             # File with the project's dependencies
├── LICENSE                      # Project license file
└── README.md                    # Main file with project description
```

## Project Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/timartinelli/Salesforce-Duplicate-Accounts.git



