import pandas as pd
import logging
from simple_salesforce import Salesforce

logger = logging.getLogger(__name__)

# Definindo as consultas SQL para cada objeto relacionado
RELATED_OBJECT_QUERIES = {
    'Opportunities': "SELECT count() FROM Opportunity WHERE AccountId = '{account_id}'",
    'Contacts': "SELECT count() FROM Contact WHERE AccountId = '{account_id}'",
    'Cases': "SELECT count() FROM Case WHERE AccountId = '{account_id}'",
    'Tasks': "SELECT count() FROM Task WHERE WhatId = '{account_id}'",
    'Events': "SELECT count() FROM Event WHERE WhatId = '{account_id}'",
    'Campaigns': "SELECT count() FROM Campaign WHERE AccountId = '{account_id}'",
    'Notes': "SELECT count() FROM Note WHERE ParentId = '{account_id}'",
    'Contracts': "SELECT count() FROM Contract WHERE AccountId = '{account_id}'",
    'Invoices': "SELECT count() FROM Invoice WHERE AccountId = '{account_id}'",
    'Orders': "SELECT count() FROM Order WHERE AccountId = '{account_id}'",
    'Products': "SELECT count() FROM Product2 WHERE AccountId = '{account_id}'",
    'PaymentInvoices': "SELECT count() FROM PaymentInvoice WHERE AccountId = '{account_id}'",
    'BusinessReports': "SELECT count() FROM BusinessReport WHERE AccountId = '{account_id}'",
    'Documents': "SELECT count() FROM Document WHERE AccountId = '{account_id}'"
}

def load_no_info_accounts():
    """Carrega contas do arquivo conta_duplicadas.xlsx com Status 'NO INFO'."""
    try:
        file_path = 'data_source/conta_duplicadas.xlsx'
        df = pd.read_excel(file_path)
        no_info_accounts = df[df['STATUS'] == 'NO INFO']
        logger.debug(f"Contas carregadas com 'NO INFO': {len(no_info_accounts)}")
        return no_info_accounts
    except Exception as e:
        logger.error(f"Erro ao carregar o arquivo conta_duplicadas.xlsx: {e}")
        return None

def load_other_accounts():
    """Carrega contas do arquivo conta_duplicadas.xlsx com Status diferente de 'NO INFO'."""
    try:
        file_path = 'data_source/conta_duplicadas.xlsx'
        df = pd.read_excel(file_path)
        other_accounts = df[df['STATUS'] != 'NO INFO']
        logger.debug(f"Contas carregadas com Status diferente de 'NO INFO': {len(other_accounts)}")
        return other_accounts
    except Exception as e:
        logger.error(f"Erro ao carregar o arquivo conta_duplicadas.xlsx: {e}")
        return None

def get_related_counts(sf, account_id):
    """Retorna a contagem de objetos relacionados para uma conta usando o ID da conta."""
    counts = {}
    for key, query in RELATED_OBJECT_QUERIES.items():
        try:
            counts[key] = sf.query(query.format(account_id=account_id))['totalSize']
        except Exception as e:
            logger.error(f"Erro ao obter contagem para {key} na conta {account_id}: {e}")
            counts[key] = 0
    logger.debug(f"Contagens para a conta {account_id}: {counts}")
    return counts

def process_accounts(sf, accounts_df, output_path):
    """Processa contas, adiciona contagens de objetos relacionados e salva em um novo arquivo Excel."""
    try:
        accounts_copy = accounts_df.copy()
        counts_data = {key: [] for key in RELATED_OBJECT_QUERIES.keys()}

        for _, row in accounts_copy.iterrows():
            account_id = row['Id']
            counts = get_related_counts(sf, account_id)
            for key, count in counts.items():
                counts_data[key].append(count)

        # Adiciona as contagens como novas colunas no DataFrame
        for key, values in counts_data.items():
            accounts_copy[key] = values

        accounts_copy.to_excel(output_path, index=False)
        logger.info(f"Arquivo salvo com contagens de objetos relacionados em: {output_path}")
        return output_path
    except Exception as e:
        logger.error(f"Erro ao processar contas: {e}")
        return None

def process_all_accounts(sf, only_no_info=False):
    """Processa contas 'NO INFO' e demais contas conforme especificado."""
    if only_no_info:
        no_info_accounts = load_no_info_accounts()
        if no_info_accounts is not None and not no_info_accounts.empty:
            return process_accounts(sf, no_info_accounts, 'data/conta_duplicadas_no_info.xlsx')
    else:
        other_accounts = load_other_accounts()
        if other_accounts is not None and not other_accounts.empty:
            return process_accounts(sf, other_accounts, 'data/conta_duplicadas_demais_contas.xlsx')
