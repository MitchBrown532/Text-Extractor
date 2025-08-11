import re
import pandas as pd
import os

path = r"data\status data\Status Station Report 4_7-23-2025.txt"
if not os.path.exists(path):
    raise FileNotFoundError(f"File not found: {path}")

# Load your file
with open(path, "r", encoding="utf-8") as f:
    text = f.read()

# Split each record based on a repeating identifier
records = re.split(r"\n\s+Voice System name:", text)
parsed_data = []

# Define regex patterns for your desired fields
patterns = {
    "Extension": r"Extension:\s+(\d+)",
    "Service State": r"Service State:\s+([\w\-\/]+)",
    "EC500": r"EC500 Status:\s+([\w\/\-]+)",
    "Off-PBX Service State": r"Off-PBX Service State:\s+([\w\-\/]+)",
    "TCP Signal Status": r"TCP Signal Status:\s+([\w\-\/]+)",
    "Registration Status": r"Registration Status:\s+([\w\-]+)",
    "MAC Address": r"MAC Address:\s+([\w:]+)",
}

# Extract data for each block
for block in records:
    if "Extension:" not in block:
        continue
    entry = {}
    for field, pattern in patterns.items():
        match = re.search(pattern, block)
        entry[field] = match.group(1) if match else ""
    parsed_data.append(entry)

# Convert to DataFrame
df = pd.DataFrame(parsed_data)

# View
# print(df.head())  

# Save
df.to_csv("status_output_4.csv", index=False)
