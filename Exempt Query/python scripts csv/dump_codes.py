import pandas as pd
import json

excel_path = 'local code - plaza wise.xlsx'
df = pd.read_excel(excel_path)

# Build a dictionary
plaza_dict = {}
for i, row in df.iterrows():
    plaza = str(row[df.columns[0]]).strip()
    code = str(row[df.columns[1]]).strip()
    if plaza and plaza != 'nan':
        plaza_dict[plaza] = code

with open('codes_dump.json', 'w') as f:
    json.dump(plaza_dict, f, indent=4)
