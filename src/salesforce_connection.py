from simple_salesforce import Salesforce, SalesforceAuthenticationFailed

def connect_to_salesforce(username, password, security_token, domain):
    try:
        sf = Salesforce(username=username, password=password, security_token=security_token, domain=domain)
        print("Connected to Salesforce!")
        return sf
    except SalesforceAuthenticationFailed as e:
        print(f"Authentication failed: {e}")
        return None
