import requests
import psycopg2
import pandas as pd
from requests.auth import HTTPBasicAuth
from bs4 import BeautifulSoup
from collections import Counter
from requests.exceptions import HTTPError, ConnectionError, Timeout
import time

# Function to fetch all work item IDs for a project
def fetch_work_item_ids(ado_org_url, project_name, pat):
    wiql_url = f"{ado_org_url}/{project_name}/_apis/wit/wiql?api-version=6.0"
    wiql_query = {"query": f"SELECT [System.Id] FROM workitems WHERE [System.TeamProject] = '{project_name}'"}
    response = requests.post(wiql_url, json=wiql_query, auth=HTTPBasicAuth('', pat))
    response.raise_for_status()
    work_item_ids = [item['id'] for item in response.json()['workItems']]
    return work_item_ids

# Function to fetch work item details in batches with retries and rate limiting
def fetch_work_item_details(ado_org_url, pat, work_item_ids):
    work_items_url = f"{ado_org_url}/_apis/wit/workitemsbatch?api-version=6.0"
    chunk_size = 200  # Azure DevOps allows fetching up to 200 work items per batch
    work_item_details = []

    for i in range(0, len(work_item_ids), chunk_size):
        ids_chunk = work_item_ids[i:i + chunk_size]
        success = False
        retries = 3
        while not success and retries > 0:
            try:
                response = requests.post(
                    work_items_url,
                    json={"ids": ids_chunk},
                    headers={"Content-Type": "application/json"},
                    auth=HTTPBasicAuth('', pat),
                    timeout=30  # Increase timeout as needed
                )
                response.raise_for_status()
                work_item_details.extend(response.json().get('value', []))
                print(response.json())
                success = True
            except (HTTPError, ConnectionError, Timeout) as e:
                print(f"Error occurred: {e}")
                retries -= 1
                time.sleep(5)  # Wait before retrying
                if retries == 0:
                    print("Max retries reached. Moving to the next chunk.")
        time.sleep(1)  # Add delay to respect rate limits

    return work_item_details

# Function to extract relevant details and clean HTML tags from descriptions
def extract_details(work_item_details, project_name):
    data = []
    for item in work_item_details:
        fields = item['fields']
        description_html = fields.get('System.Description', '')
        description_text = BeautifulSoup(
            description_html, 'html.parser').get_text() if description_html else ''
        
        due_date = fields.get('Microsoft.VSTS.Scheduling.DueDate', '')
        created_date = fields.get('System.CreatedDate', '')

        # Handle empty date fields
        if not due_date:
            due_date = None
        if not created_date:
            created_date = None

        data.append({
            'ID': item['id'],
            'Project Name': project_name,
            'Key': fields.get('System.WorkItemType', '') + '-' + str(item['id']),
            'Summary': fields.get('System.Title', ''),
            'Description': description_text,
            'Assignee': fields.get('System.AssignedTo', {}).get('displayName', ''),
            'Reporter': fields.get('System.CreatedBy', {}).get('displayName', ''),
            'Issue Type': fields.get('System.WorkItemType', ''),
            'Time Estimate': fields.get('Microsoft.VSTS.Scheduling.OriginalEstimate', 0),
            'Time Spent': fields.get('Microsoft.VSTS.Scheduling.CompletedWork', 0),
            'Due Date': due_date,
            'Created Date': created_date
        })
    return data

# Function to calculate statistics using Counter
def calculate_statistics(work_item_details):
    work_item_types = [item['fields'].get('System.WorkItemType', '') for item in work_item_details]
    return Counter(work_item_types)

# Function to process a single project
def process_project(project_row, connection):
    project_name = project_row['Project Name']
    ado_org_url = project_row['Organization URL']
    pat = project_row['PAT']

    work_item_ids = fetch_work_item_ids(ado_org_url, project_name, pat)
    work_item_details = fetch_work_item_details(ado_org_url, pat, work_item_ids)
    print(work_item_details)
    work_item_data = extract_details(work_item_details, project_name)
    statistics = calculate_statistics(work_item_details)

    # Insert work item details into the database
    insert_work_items(connection, work_item_data)

    # Insert statistics into the database
    insert_statistics(connection, statistics, project_name)

def insert_work_items(connection, work_item_data):
    cursor = connection.cursor()
    for item in work_item_data:
        cursor.execute("""
            INSERT INTO ado_issue_details (id, project_name, key, summary, description, assignee, reporter, issue_type, time_estimate, time_spent, due_date, created_date)
            VALUES (%(ID)s, %(Project Name)s, %(Key)s, %(Summary)s, %(Description)s, %(Assignee)s, %(Reporter)s, %(Issue Type)s, %(Time Estimate)s, %(Time Spent)s, %(Due Date)s, %(Created Date)s)
        """, item)
        print(f"{item['Key']} was added.")
    connection.commit()
    cursor.close()

def insert_statistics(connection, statistics, project_name):
    cursor = connection.cursor()
    for work_item_type, count in statistics.items():
        cursor.execute("""
            INSERT INTO ado_issue_stats (project_name, issue_type, count)
            VALUES (%s, %s, %s)
        """, (project_name, work_item_type, count))
        print(f"{work_item_type} for project {project_name} was added.")
    connection.commit()
    cursor.close()

# Establish connection to the PostgreSQL database
def connect_to_database():
    conn = psycopg2.connect(
        dbname="jiratoadodb",
        user="postgres",
        password="0000",
        host="40.117.145.14",
        port="5432"
    )
    print("Connected to the database")
    return conn

# Read project details from the Excel file
projects_df = pd.read_excel('ado_project_details_db.xlsx')

# Process each project separately
for index, project_row in projects_df.iterrows():
    connection = connect_to_database()
    process_project(project_row, connection)
    connection.close()

print("Work item details and statistics saved to PostgreSQL database.")
