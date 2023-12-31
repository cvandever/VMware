#!/usr/bin/env python3

import requests
import pandas as pd
from datetime import datetime, timedelta
import os, base64, json
from jinja2 import Template

base_url = "https://apis-us0.druva.com"

def get_valid_integer_input(prompt, default_value, max_attempts):
    for attempt in range(1, max_attempts + 1):
        user_input = input(f"{prompt} (default is {default_value}). Attempt {attempt}: ")
        
        try:
            value = int(user_input)
            return value
        except ValueError:
            print("Invalid input. Only integer values are allowed.")
    
    print(f"Exceeded the maximum number of attempts. Using the default value of {default_value}.")
    return default_value


def generate_base64_credentials(client_id, client_secret):
    # Concatenate client ID and client secret with a colon separator
    credentials_str = f"{client_id}:{client_secret}"
    
    # Encode the string using Base64
    credentials_bytes = credentials_str.encode('ascii')
    base64_credentials = base64.b64encode(credentials_bytes).decode('ascii')
    
    return base64_credentials

def get_token():
    token_url = f"{base_url}/token"
    client_id = os.environ.get("DRUVA_CLIENT_ID")
    client_secret = os.environ.get("DRUVA_CLIENT_SECRET")
    payload = {
        "grant_type": "client_credentials"
    }
    token_headers = {
    "Accept": "application/json",
    "Content-Type": "application/x-www-form-urlencoded",
    "Authorization": f"Basic {generate_base64_credentials(client_id, client_secret)}"
    }
    # Send the token request
    response = requests.post(token_url, data=payload, headers=token_headers)
    if response.status_code == 200:
        access_token = response.json()["access_token"]
        token_headers.update({"Content-Type": "application/json"})
        # Creating a new header with the retrieved access token to be used for subsequent requests
        token_headers.update({"Authorization": "Bearer " + access_token})
    else:
        print("Token generation failed. Status code:", response.status_code)

    return token_headers

headers = get_token()

def get_backup_activity(date_str: str, page_token: str = None):
    url = f"{base_url}/platform/reportsvc/v1/reports/ewBackupActivity"
    payload = { "filters": {
        "filterBy": [
            {
                "fieldName": "lastUpdatedTime",
                "operator": "GTE",
                "value": date_str
            }
        ],
        "pageSize": 500
        },
        "pageToken": page_token
    }
    response = requests.post(url, json=payload,headers=headers)
    if response.status_code == 200:
        backup_info = response.json()
        if backup_info["nextPageToken"] != "":
            print("Getting next page of results")
            payload.update({"pageToken": backup_info["nextPageToken"]})
            backup_info["data"].extend(get_backup_activity(date_str, backup_info["nextPageToken"])["data"])
    else:
        print("Backup info request failed. Status code:", response.status_code)
        backup_info = []
    return backup_info

def readable_datetime(date_str: str):
    try:
        date_obj = datetime.strptime(date_str, '%Y-%m-%dT%H:%M:%S.%fZ')
    except ValueError:
        date_obj = datetime.strptime(date_str, '%Y-%m-%dT%H:%M:%SZ')
    readable_date = date_obj.strftime("%B %d, %Y")
    return readable_date


def get_seven_days_prior(delta: int = 7):
    current_datetime = datetime.utcnow()
    # Subtract 7 days from the current date and time to get the date 7 days ago in UTC
    prior_datetime = current_datetime - timedelta(days=delta)
    formatted_datetime = prior_datetime.isoformat() + "Z"
    return formatted_datetime


def get_backup_activity_report(date_str: str = get_seven_days_prior()):
    backup_list = []
    backup_data = get_backup_activity(date_str)
    for backup in backup_data["data"]:
        backup["started"] = readable_datetime(backup["started"])
        backup["ended"] = readable_datetime(backup["ended"])
        backup = {key: backup[key] for key in backup if key not in ["backupContent", "backupMethod", "backupMountName", "backupSet", "deviceName", "resourceType", "scanType", "scheduled", "lastUpdatedTime"]}
        backup_list.append(backup)
    return backup_list

def generate_report(data):
    columns = ["jobID","vm","status","started","ended","backupType","backupPolicy","workloads","organization","administrativeGroup","details"]
    # Convert the data into a pandas DataFrame with the specified columns
    df = pd.DataFrame(data, columns=columns)
    df = df.sort_values(by="jobID", ascending=False)

    def apply_status_style(val):
        if val == 'Successful':
            return 'background-color: green;'
        elif val == 'Failed':
            return 'background-color: red;'
        else:
            return 'background-color: yellow;'
        
    # Apply the style to the DataFrame
    df = df.style.applymap(apply_status_style, subset=pd.IndexSlice[:, ['status']])
        
    html_table = df.hide_index().set_table_styles([{'selector': 'td', 'props': [('border', '1px solid black')]}, {'selector': 'th', 'props': [('border', '1px solid black')]}]).render()

    # Define the HTML template for the report
    template_str = """
    <html>
    <head>
        <title>Backup Activity Report</title>
        <style>
        table {
            border-collapse: collapse;
        }
        </style>
    </head>
    <body>
        <h1>Backup Activity Report</h1>
        {{ table }}
    </body>
    </html>
    """

    # Create a Jinja2 template from the template string
    template = Template(template_str)

    # Render the template with the HTML table
    html_report = template.render(table=html_table)

    return html_report


def main():
    delta = get_valid_integer_input("Enter the number of days to include in the report", 7, 3)
    date_str = get_seven_days_prior(delta)
    # Get the backup activity report
    backup_activity = get_backup_activity_report(date_str)

    # Generate the report
    report = generate_report(backup_activity)

    # Write the report to a file
    with open("backup_activity_report.html", "w") as f:
        f.write(report)

if __name__ == "__main__":
    main()