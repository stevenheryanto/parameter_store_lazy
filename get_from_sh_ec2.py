import re
import openpyxl

# Path to your text file and Excel file
env = 'PRD'
text_file_path = env+'/pre'+env+'.sh'
excel_file_path = 'ssm_parameter_store.xlsx'

# Regular expression pattern to find values after --name "VALUE"
pattern = r'--name\s+"([^"]+)"'

# Step 1: Extract values from the text file
with open(text_file_path, 'r') as file:
    content = file.read()
    matches = re.findall(pattern, content)

# Step 2: Load the existing Excel file and sheet
workbook = openpyxl.load_workbook(excel_file_path)
sheet = workbook[env]  # Use the active sheet, or specify by name, e.g., workbook["Sheet1"]

# Step 3: Collect existing values in column B to avoid duplicates
existing_values = set()
for cell in sheet['B']:
    if cell.value:
        existing_values.add(cell.value)

# Step 4: Append new values to column B, avoiding duplicates
new_values_added = False
for match in matches:
    if match not in existing_values:
        print(match)
        sheet.append([None, match])  # Append to column B (column A is left empty)
        existing_values.add(match)  # Add to the set to prevent future duplicates
        new_values_added = True
    else:
        print("duplicate: ", match)

# Step 5: Save changes if new values were added
if new_values_added:
    workbook.save(excel_file_path)
    print("New values added and file saved.")
else:
    print("No new values to add.")

