import os
import json
import re
from glob import glob
import openpyxl

# Paths to the folder containing JSON files and the Excel file
env = 'PRD'
folder_path = env
excel_file_path = 'ssm_parameter_store.xlsx'

# Regular expression pattern to find values after "parameter/"
pattern = r'parameter/([^"]+)'

# Step 1: Extract values after "parameter/" from JSON files
extracted_values = []
for file_path in glob(os.path.join(folder_path, '*.json')):
    with open(file_path, 'r') as file:
        try:
            data = json.load(file)  # Load JSON data
            json_str = json.dumps(data)  # Convert JSON data to a string
            matches = re.findall(pattern, json_str)  # Find all matches
            extracted_values.extend(matches)  # Add matches to the list
        except json.JSONDecodeError:
            print(f"Error decoding JSON in file: {file_path}")

# Step 2: Load the existing Excel file and select the sheet
workbook = openpyxl.load_workbook(excel_file_path)
sheet = workbook[env]  # Use the active sheet or specify by name, e.g., workbook["Sheet1"]

# Step 3: Collect existing values in column B to avoid duplicates
existing_values = set()
for cell in sheet['B']:
    if cell.value:
        existing_values.add(cell.value)

# Step 4: Append new, unique values to column B
new_values_added = False
for value in set(extracted_values):  # Use a set to avoid duplicates in the extracted values
    if value not in existing_values:
        sheet.append([None, value])  # Append to column B (column A is left empty)
        existing_values.add(value)  # Add to the set to prevent future duplicates
        new_values_added = True
    else:
        print("duplicate: ", value)

# Step 5: Save changes if new values were added
if new_values_added:
    workbook.save(excel_file_path)
    print("New values added and file saved.")
else:
    print("No new values to add.")

