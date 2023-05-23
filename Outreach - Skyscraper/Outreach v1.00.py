import pandas as pd
import requests
import tldextract
import os
from openpyxl import Workbook
import time


# API Key
API_KEY = '346266b96ff7d733bf03f4d4bbbe5bfdad6520c1'

error_log = []

# Get directory from the user
directory = r"C:\Users\Hooman Deghani\OneDrive\PC\Desktop\Input"

# List all CSV files in the directory
urls = [os.path.join(directory, filename) for filename in os.listdir(directory) if filename.endswith('.csv')]

df = pd.concat([pd.read_csv(url) for url in urls], ignore_index=True)

# Phase 2: Clean up dataframe
df = df[['Referring page URL', 'Domain rating', 'Target URL', 'Anchor']]
df['Root URL'] = df['Referring page URL'].apply(lambda x: f"https://{tldextract.extract(x).domain}.{tldextract.extract(x).suffix}")
df[['First Name', 'Last Name', 'Recipient', 'Status', 'Replied', 'Converted']] = ""


# Phase 3: The conditions (outlined in the instructions)
departments = ['marketing', 'communication', 'it', 'hr']
seniority_levels = ['executive', 'senior', 'junior']

def get_email_from_api(url, department, seniority):
    params = {
        'domain': url,
        'api_key': API_KEY,
        'department': department,
        'seniority': seniority
    }
    try:
        response = requests.get('https://api.hunter.io/v2/domain-search', params=params)
        response.raise_for_status()  # Raises a HTTPError if the status is 4xx, 5xx
        data = response.json()
        if 'data' in data and 'emails' in data['data']:
            if data['data']['emails']:
                return data['data']['emails'][0]
    except Exception as e:
        error_log.append(f"Error fetching email for '{url}': {str(e)}")
    return None

def verify_email(email):
    params = {
        'email': email,
        'api_key': API_KEY
    }
    try:
        response = requests.get('https://api.hunter.io/v2/email-verifier', params=params)
        response.raise_for_status()
        data = response.json()
        if 'data' in data and 'status' in data['data']:
            return data['data']['status']
    except Exception as e:
        error_log.append(f"Error verifying email '{email}': {str(e)}")
    return None

# Phase 4: Connect to API
for index, row in df.iterrows():
    time.sleep(0.2)  # Add a sleep duration of 200 ms
    email_data = None
    root_url = row['Root URL']

    for department in departments:
        for seniority in seniority_levels:
            email_data = get_email_from_api(root_url, department, seniority)
            if email_data:
                break
        if email_data:
            break

    if not email_data:
        email_data = get_email_from_api(root_url, None, None)

    if email_data:
        email = email_data['value']
        status = verify_email(email)
        if status in ['valid', 'accept_all']:
            df.loc[index, 'First Name'] = email_data['first_name']
            df.loc[index, 'Last Name'] = email_data['last_name']
            df.loc[index, 'Recipient'] = email

# Fill missing first names with "Sir/Madam"
df.loc[df['First Name'].isnull() & df['Recipient'].notnull(), 'First Name'] = 'Sir/Madam'

# Phase 5: Save the file
project = input("What is the project name? ")
df.to_excel(f"{project}.xlsx", index=False)
print("Work is done!")

# After saving the file
print("\n".join(error_log))  # Print all error messages

# Save error messages to a log file
with open('error_log.txt', 'w') as f:
    f.write("\n".join(error_log))