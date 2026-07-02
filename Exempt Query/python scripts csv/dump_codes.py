import json

import pandas as pd

from config import input_path, output_path

excel_path = input_path("local code - plaza wise.xlsx")
df = pd.read_excel(excel_path)

plaza_dict = {}
for i, row in df.iterrows():
    plaza = str(row[df.columns[0]]).strip()
    code = str(row[df.columns[1]]).strip()
    if plaza and plaza != "nan":
        plaza_dict[plaza] = code

with open(input_path("codes_dump.json"), "w") as f:
    json.dump(plaza_dict, f, indent=4)
