import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Read the text data
with open('data/reports 10_7-8-2025.txt', 'r', encoding='utf-8') as file:
    text = file.read()

# Split into station blocks
station_blocks = re.split(r'Voice System name: SuperNap - STATION', text)[1:]

records = []

for block in station_blocks:
    row = {}

    # Extract simple fields
    row['Extension'] = re.search(r'Extension:\s*(\d+)', block)
    row['Type'] = re.search(r'Type:\s*(\S+)', block)
    
    # Improved Name extraction
    name_match = re.search(r'Name:\s*(.*?)(?=\s+Coverage Path 2:|\s{2,}|\Z)', block)
    row['Name'] = name_match.group(1).strip() if name_match else ''
    
    row['Coverage Path 1'] = re.search(r'Coverage Path 1:\s*(\d*)', block)
    row['Coverage Path 2'] = re.search(r'Coverage Path 2:\s*(\d*)', block)

    # Error handling: Preventing Hunt-to from extracting extra text
    def safe_extract_hunt_to(block):
        # Look for 'Hunt-to Station:'
        match = re.search(r'Hunt-to Station:\s*(\S*)', block)
        if match:
            val = match.group(1)
            # If it's more than 3 spaces away (empty or very delayed), ignore it
            if len(match.group(0).split(':')[1].lstrip()) > 3:
                return ''
            # Also make sure it’s not accidentally picking up "Tests?"
            if val and not val.lower().startswith("tests"):
                return val.strip()
        return ''
    
    row['Hunt-to Station'] = safe_extract_hunt_to(block)
    row['Message Lamp Ext'] = re.search(r'Message Lamp Ext:\s*(\S*)', block)
    row['EC500 State'] = re.search(r'EC500 State:\s*(\S*)', block)

    def extract_cfwd_line(label):
        match = re.search(rf'{re.escape(label)}:\s*(.*?)\s+([yn])\s*$', block, re.MULTILINE)
        if match:
            destination = match.group(1).strip()
            active = match.group(2).strip()
            return destination, active
        else:
            return '', ''

    row['CFWD_Uncond_Internal'], row['CFWD_Uncond_Internal_Active'] = extract_cfwd_line('Unconditional For Internal Calls To')
    row['CFWD_External'], row['CFWD_External_Active'] = extract_cfwd_line('External Calls To')
    row['CFWD_Busy_Internal'], row['CFWD_Busy_Internal_Active'] = extract_cfwd_line('Busy For Internal Calls To')
    row['CFWD_External_2'], row['CFWD_External_2_Active'] = extract_cfwd_line('External Calls To')  # second appearance
    row['CFWD_NoReply_Internal'], row['CFWD_NoReply_Internal_Active'] = extract_cfwd_line('No Reply For Internal Calls To')
    row['CFWD_External_3'], row['CFWD_External_3_Active'] = extract_cfwd_line('External Calls To')  # third appearance

    # Extract Button Assignments 1–24
    buttons = [''] * 24
    collecting = False

    for line in block.splitlines():
        if 'BUTTON ASSIGNMENTS' in line:
            collecting = True
            continue

        if collecting and (not line.strip() or 'STATION' in line or 'SITE DATA' in line):
            collecting = False

        if collecting:
            matches = re.findall(r'(\d{1,2}):\s*(\S.*?)(?=\s+\d+:|$)', line)
            for idx_str, val in matches:
                idx = int(idx_str)
                if 1 <= idx <= 24:
                    buttons[idx - 1] = val.strip()

    for i in range(24):
        row[f'Button_{i+1}'] = buttons[i]



    # Add cleaned values to list
    records.append({k: (v.group(1).strip() if isinstance(v, re.Match) else v or '') for k, v in row.items()})

# Convert to DataFrame
df = pd.DataFrame(records)

# Create Excel file starting at D3
wb = Workbook()
ws = wb.active

ws.append([])

# Write to excel
for row in dataframe_to_rows(df, index=False, header=False):
    ws.append(row)

from openpyxl.styles import Protection

for row in ws.iter_rows():
    for cell in row:
        cell.protection = Protection(locked=False)

# Save to file
wb.save("outputs 10_7-8-2025.xlsx")
