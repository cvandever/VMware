#!/usr/bin/env python3

import requests
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox
import os, csv

base_url = "https://apigtwb2c.us.dell.com"

def get_token():
    token_url = f"{base_url}/auth/oauth/v2/token"
    client_id = os.environ.get("DELL_API_KEY")
    client_secret = os.environ.get("DELL_API_SECRET")
    payload = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret
    }
    # Send the token request
    response = requests.post(token_url, data=payload)
    # Get the access token from the response
    if response.status_code == 200:
        access_token = response.json()["access_token"]
    else:
        print("Token generation failed. Status code:", response.status_code)

    return access_token

access_token = get_token()

headers = {
    "Accept": "application/json",
    "Authorization": "Bearer " + access_token
}

# Convert the date string to a readable format
def convert_datetime(date_str: str):
    try:
        date_obj = datetime.strptime(date_str, '%Y-%m-%dT%H:%M:%S.%fZ')
    except ValueError:
        date_obj = datetime.strptime(date_str, '%Y-%m-%dT%H:%M:%SZ')
    readable_date = date_obj.strftime("%B %d, %Y")
    return readable_date

# Get the device details
def get_device_details(service_tag: str):
    device_url = f"{base_url}/PROD/sbil/eapi/v5/asset-entitlement-components?servicetag={service_tag}"
    response = requests.get(device_url, headers=headers)
    if response.status_code == 200:
        device_info = response.json()
    else:
        print("Device info request failed. Status code:", response.status_code)
        device_info = []
    return device_info


# Process the info to create Custom Json
def process_warranty_info(service_tags: list, customer_name: str):
    details_info =[get_device_details(service_tag) for service_tag in service_tags]
    devices = []
    for device in details_info:
        # List Comprehension to create a list of dictionaries for each entitlement
        device_entitlements = [{"Item Number": entitlement["itemNumber"],
                "Start Date": convert_datetime(entitlement["startDate"]),
                "End Date": convert_datetime(entitlement["endDate"]),
                "Entitlement Type": entitlement["entitlementType"],
                "Service Level Description": entitlement["serviceLevelDescription"]} for entitlement in device["entitlements"]]
        # List Comprehension to create a list of dictionaries for each component if the quantity is greater than 0
        device_components = [{"Item Number": component["itemNumber"],
                "Part Number": component["partNumber"],
                "Part Description": component["partDescription"],
                "Item Description": component["itemDescription"],
                "Quantity": component["partQuantity"]} for component in device["components"] if component["partQuantity"] > 0]
        # Append the device details to the devices list
        devices.append({
            "Service Tag": device["serviceTag"],
            "Product ID": device["productId"],
            "System Description": device["systemDescription"],
            "Ship Date": convert_datetime(device["shipDate"]),
            "Entitlements": device_entitlements,
            "Components": device_components
        })
    return export_to_excel(devices, customer_name)


def export_to_excel(devices: list,customer_name: str):
    datetime_now = datetime.now().strftime("%Y-%m-%d")
    # Create Entitlements dataframe
    warranty_df = pd.json_normalize(devices).explode('Entitlements').drop('Components', axis=1)

    # Expand the nested dictionary into separate columns
    entitlement_cols = warranty_df['Entitlements'].apply(pd.Series)
    warranty_df = pd.concat([warranty_df, entitlement_cols], axis=1)
    warranty_df.drop('Entitlements', axis=1, inplace=True)

    # Mark duplicated rows and remove duplicate values
    warranty_df['Duplicated'] = warranty_df.duplicated(subset=['Service Tag', 'Product ID', 'System Description', 'Ship Date'], keep='first')
    warranty_df.loc[warranty_df['Duplicated'], ['Service Tag', 'Product ID', 'System Description', 'Ship Date']] = ""
    warranty_df.drop('Duplicated', axis=1, inplace=True)

    # Create Components dataframe
    component_df = pd.json_normalize(devices).explode('Components').drop('Entitlements', axis=1)

    # Expand the nested dictionary into separate columns
    component_cols = component_df['Components'].apply(pd.Series)
    component_df = pd.concat([component_df, component_cols], axis=1)
    component_df.drop('Components', axis=1, inplace=True)

    # Mark duplicated rows and remove duplicate values
    component_df['Duplicated'] = component_df.duplicated(subset=['Service Tag', 'Product ID', 'System Description', 'Ship Date'], keep='first')
    component_df.loc[component_df['Duplicated'], ['Service Tag', 'Product ID', 'System Description', 'Ship Date']] = ""
    component_df.drop('Duplicated', axis=1, inplace=True)


    # Create the Excel writer
    # prompt the user for the customer name and create the output file name
    output_file = f'{customer_name}_ServiceTags_{datetime_now}.xlsx'
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    warranty_df.to_excel(writer, index=False, sheet_name='Warranty Info')
    component_df.to_excel(writer, index=False, sheet_name='Component Info')

    # Access the workbook and the worksheet
    workbook = writer.book
    warranty_worksheet = writer.sheets['Warranty Info']
    component_worksheet = writer.sheets['Component Info']

    # Add borders to the cells
    border_format = workbook.add_format({'border': 1})
    num_rows, num_cols = warranty_df.shape
    warranty_worksheet.conditional_format(1, 0, num_rows, num_cols-1, {'type': 'no_blanks', 'format': border_format})
    num_rows, num_cols = component_df.shape
    component_worksheet.conditional_format(1, 0, num_rows, num_cols-1, {'type': 'no_blanks', 'format': border_format})

    
    # Save the Excel file
    writer.save()
    return output_file


def select_csv_file():
    app = QApplication([])
    file = QFileDialog.getOpenFileName(None, "Import CSV", "", "CSV Files (*.csv)")

    if file:
        app.quit()
        return file[0]
    else:
        QMessageBox.warning(None, "Warning", "No CSV file selected.")
        return None

def get_csv_values(csv_file):
    column_values = []
    with open(csv_file, 'r', newline='') as file:
        reader = csv.reader(file)
        next(reader)  # Skip the header row
        for row in reader:
            column_values.append(row[0])
    return column_values


def main():
    csv_file = select_csv_file()
    if csv_file:
        customer_name = input("Enter the customer name: ")
        service_tags = get_csv_values(csv_file)
        base_path = os.path.dirname(csv_file)
        outfile = process_warranty_info(service_tags, customer_name)
        print(f"Writing Data to {base_path}/{outfile}")
    

if __name__ == "__main__":
    main()


# Unused Endpoints

'''def get_warranty_info(service_tag):
    warranty_url = f"{base_url}/PROD/sbil/eapi/v5/asset-entitlements?servicetags={service_tag}"
    response = requests.get(warranty_url, headers=headers)
    if response.status_code == 200:
        warranty_info = response.json()
    else:
        print("Warranty info request failed. Status code:", response.status_code)
    return warranty_info

def get_device_summary(service_tag):
    summary_url = f"{base_url}/PROD/sbil/eapi/v5/asset-components?servicetag={service_tag}"
    response = requests.get(summary_url, headers=headers)
    if response.status_code == 200:
        summary_info = response.json()
    else:
        print("Device summary request failed. Status code:", response.status_code)
        summary_info = []
    return summary_info'''