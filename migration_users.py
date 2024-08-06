import requests
import json
from requests.auth import HTTPBasicAuth

# Azure DevOps organization and project details
organization = "dhanushv"
project_name = "Newt CAMPP1"  # Replace with your project name
pat = "koergsg5rol67lg7ulan63qxmu2fv3tz46fpqsme5cumqf2pkiha"  # Personal Access Token

# Users to be added
users = [
    {
        "principalName": "hemag@newtglobalcorp.com",
        "accessLevel": "express",
        "groupType": "projectAdministrator"
    }
]

# Base64-encode the Personal Access Token (PAT)
auth = HTTPBasicAuth('', pat)

# Function to get project ID by project name
def get_project_id(project_name):
    url = f"https://dev.azure.com/{organization}/_apis/projects?api-version=6.0"
    response = requests.get(url, auth=auth)
    if response.status_code == 200:
        projects = response.json()['value']
        for project in projects:
            if project['name'].lower() == project_name.lower():
                print(project['id'])
                return project['id']
        raise Exception(f"Project '{project_name}' not found.")
    else:
        raise Exception(f"Failed to fetch projects. Status Code: {response.status_code}, Response: {response.text}")

# Function to add user to project
def add_user_to_project(user, project_id):
    url = f"https://vsaex.dev.azure.com/{organization}/_apis/userentitlements?api-version=7.1-preview.4"
    body = {
        "accessLevel": {
            "licensingSource": "account",
            "accountLicenseType": user["accessLevel"]
        },
        "user": {
            "principalName": user["principalName"],
            "subjectKind": "user"
        },
        "projectEntitlements": [
            {
                "group": {
                    "groupType": user["groupType"]
                },
                "projectRef": {
                    "id": project_id
                }
            }
        ]
    }
    response = requests.post(url, auth=auth, headers={'Content-Type': 'application/json'}, data=json.dumps(body))
    
    if response.status_code in [200, 201]:
        print(f"Successfully added {user['principalName']} to the project with {user['groupType']} role.")
    else:
        print(f"Failed to add {user['principalName']} to the project. Status Code: {response.status_code}, Response: {response.text}")

# Get project ID
project_id = get_project_id(project_name)

# Add each user to the project
for user in users:
    add_user_to_project(user, project_id)
