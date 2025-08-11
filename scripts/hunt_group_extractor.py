import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Initialize storage
records = []

with open("data/Display Hunt Group_7-9-2025.txt", "r", encoding="utf-8") as file:
    lines = file.readlines()

group_number = group_extension = group_type = acd = voicemail = caller_display = first_member_ext = None

for i, line in enumerate(lines):
    # Extract core group info
    if "Group Number:" in line and "ACD?" in line:
        group_number = re.search(r"Group Number:\s*(\d+)", line)
        acd = re.search(r"ACD\?\s*(\w)", line)
        group_number = group_number.group(1) if group_number else None
        acd = acd.group(1) if acd else None

    if "Group Extension:" in line and "Vector?" in line:
        group_extension = re.search(r"Group Extension:\s*(\d+)", line)
        group_extension = group_extension.group(1) if group_extension else None

    if "Group Type:" in line:
        group_type = re.search(r"Group Type:\s*([\w-]+)", line)
        group_type = group_type.group(1) if group_type else None

    if "ISDN/SIP Caller Display:" in line:
        caller_display = line.split("ISDN/SIP Caller Display:")[-1].strip()

    if "Message Center:" in line:
        voicemail = line.split("Message Center:")[-1].strip()

    if first_member_ext is None and line.strip().startswith('1:'):
        # Take only the first five digits after '1:'
        match = re.search(r'1:\s*(\d{5})', line)
        if match:
            first_member_ext = match.group(1)
  
    # Once the "At End of Member List" is hit, record the block
    if "At End of Member List" in line:
        records.append({
            "Group Number": group_number,
            "Group Extension": group_extension,
            "Group Type": group_type,
            "ACD": acd,
            "Voicemail": voicemail,
            "External Caller Display": caller_display,
            "First Member Ext": first_member_ext
        })
        # Reset group info for next block
        group_number = group_extension = group_type = acd = voicemail = caller_display = first_member_ext = None

# Create DataFrame
df = pd.DataFrame(records)

# Keep only rows that have a Group Number
df = df[df['Group Number'].notna() & (df['Group Number'] != '')]

# Write to Excel starting at A3
output_path = "hunt_groups_output.xlsx"

# Create workbook with openpyxl to start at A3
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

# Leave rows 1 and 2 blank
ws.append([])
ws.append([])

# Add header and data starting from A3
for row in dataframe_to_rows(df, index=False, header=False):
    ws.append(row)

wb.save(output_path)
