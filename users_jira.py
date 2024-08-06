import requests
import pandas as pd

# JIRA configurations
JIRA_BASE_URL = 'https://newtglobal.atlassian.net'
JIRA_USERNAME = 'dhanushv@newtglobalcorp.com'  # Your JIRA username (email)
JIRA_API_TOKEN = ''  # Your JIRA API token
JIRA_PROJECT_KEY = 'CAMPP'  # The project key in JIRA

# Function to fetch all roles for a project
def fetch_all_roles(project_key, username, api_token):
    url = f"{JIRA_BASE_URL}/rest/api/2/project/{project_key}/role"
    headers = {
        "Accept": "application/json"
    }
    auth = (username, api_token)
    response = requests.get(url, headers=headers, auth=auth)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error fetching roles for project {project_key}:", response.text)
        return None

# Function to fetch users assigned to a specific role
def fetch_users_by_role(role_url, username, api_token):
    headers = {
        "Accept": "application/json"
    }
    auth = (username, api_token)
    response = requests.get(role_url, headers=headers, auth=auth)
    if response.status_code == 200:
        return response.json().get('actors', [])  # Extract 'actors' list from the response
    else:
        print(f"Error fetching users for URL {role_url}:", response.text)
        return None

# Function to fetch user details by account ID
def fetch_user_details(account_id, username, api_token):
    url = f"{JIRA_BASE_URL}/rest/api/2/user?accountId={account_id}"
    headers = {
        "Accept": "application/json"
    }
    auth = (username, api_token)
    response = requests.get(url, headers=headers, auth=auth)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error fetching user details for account ID {account_id}:", response.text)
        return None

# Fetch all roles
roles_data = fetch_all_roles(JIRA_PROJECT_KEY, JIRA_USERNAME, JIRA_API_TOKEN)

# Define roles to include (excluding 'atlassian-addons-project-access')
roles_to_include = ['Member', 'Developers', 'Administrator']

# Initialize an empty list to store user details
user_details = []

# Fetch users for each role
if roles_data:
    for role_name, role_url in roles_data.items():
        if role_name in roles_to_include:
            users_data = fetch_users_by_role(role_url, JIRA_USERNAME, JIRA_API_TOKEN)
            if users_data:
                for user in users_data:
                    account_id = user.get('actorUser', {}).get('accountId', '')
                    if account_id:
                        user_detail = fetch_user_details(account_id, JIRA_USERNAME, JIRA_API_TOKEN)
                        email = user_detail.get('emailAddress', 'NA') if user_detail else 'NA'
                    else:
                        email = 'NA'
                    
                    user_details.append({
                        'Project Name': JIRA_PROJECT_KEY,
                        'Role': role_name,
                        'Display Name': user.get('displayName', ''),  # Provide default value if 'displayName' is missing
                        'Account ID': account_id,  # Provide default value if 'actorUser' or 'accountId' is missing
                        'Email': email  # Include email in the output
                    })
            else:
                user_details.append({
                    'Project Name': JIRA_PROJECT_KEY,
                    'Role': role_name,
                    'Display Name': '',  # Default value if users data is not available
                    'Account ID': '',  # Default value if users data is not available
                    'Email': 'NA'  # Default value if users data is not available
                })

# Convert user details list to a DataFrame
df = pd.DataFrame(user_details)

# Write DataFrame to an Excel file
excel_file = f"users/{JIRA_PROJECT_KEY}_users.xlsx"
df.to_excel(excel_file, index=False)
print("User details written to", excel_file)
