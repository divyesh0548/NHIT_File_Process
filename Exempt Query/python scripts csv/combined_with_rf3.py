import pandas as pd
import numpy as np
from config import get_file_path, get_base_output_path, get_local_code
from data_processing import load_and_preprocess_data
# from main_script import main
# from RF3_condition import main as rf3_main

base_output_path = get_base_output_path()

# Load the necessary Excel files into DataFrames
# final_result = main()
# combine_rf = final_result


# rf3_1, rf3_2, result_df = rf3_main() # to capture both returned DataFrames

# rf_3_1 = rf3_1
# rf_3_1['TRI 2 & 3'] = rf_3_1['TRI 2 & 3'].astype(str)

# rf_3_2 = rf3_2
# rf_3_2['TRI 2 & 3'] = rf_3_2['TRI 2'].astype(str)
combine_rf = pd.read_excel(base_output_path + "/updated_new_Combined_Side_Results.xlsx")
rf_3_1 = pd.read_excel(base_output_path + "/rf3_1.xlsx", dtype={'TRI 2 & 3': str})
rf_3_2 = pd.read_excel(base_output_path + "/rf3_2.xlsx", dtype={'TRI 2': str})

#combine_rf["Txn ID"] = combine_rf["Txn ID"].astype(str)

# Load and preprocess data from data_processing.py (including MOP = "Violation")
file_path = get_file_path()  # Get the file path from config
df, exempt_vehicle_list = load_and_preprocess_data(file_path)

# Perform the first left join with RF 3 - 1
df_merged = combine_rf.merge(rf_3_1[['Veh Reg No.', 'TRI 2 & 3']], on='Veh Reg No.', how='left')
print(df_merged.info())

# Perform the second left join with RF 3 - 2
df_merged = df_merged.merge(rf_3_2[['Veh Reg No.', 'TRI 1', 'TRI 2', 'TRI 3']], on='Veh Reg No.', how='left')

# Replace "1" with NaN in the 'TRI 2' column
df_merged['TRI 2'] = df_merged['TRI 2'].replace('1', np.nan)

# Add the 'Local/Non Local' column dynamically from codes_dump.json
local_codes = get_local_code()
df_merged['Local/Non Local'] = df_merged['Veh Reg No.'].apply(lambda x: 'Local' if str(x).startswith(local_codes) else 'Non Local')
#df_merged['Local/Non Local'] = 'Non Local'
# df_merged['Local/Non Local'] = df_merged['Veh Reg No.'].apply(
#     lambda x: 'Non Local'
#     if len(str(x)) == 4
#     else (
#         'Local'
#         if str(x).startswith(('UP31', 'UP30', 'UP34', 'UP27'))
#         else 'Non Local'
#     )
# )


# Add the 'C/NC' column based on 'TC Class'
df_merged['C/NC'] = df_merged['TC Class'].apply(lambda x: 'Non Commercial' if x in ['Car', 'CAR']  else 'Commercial')

# Combine 'Local/Non Local' and 'C/NC' into a new column 'LNC/NLNC/LC/NLC'
df_merged['LNC/NLNC/LC/NLC'] = df_merged['Local/Non Local'] + ' ' + df_merged['C/NC']


# Re-load the data and include 'Violation' rows for 'Trip Rules' calculation
original_df = pd.read_csv(file_path, skiprows=0, skipfooter=0)  # Load the original DataFrame
#df_violation = original_df[(original_df['MOP'] == "RUNTHROUGH") & (original_df['MOP'] == "RUNTHROUGHRUNTHROUGH")]
#df_violation = original_df[(original_df['MOP'] == "VIOLATION") & (original_df['MOP'] == "FLEET") & (original_df['MOP'] == "CONVEY") & (original_df['MOP'] == "RUNTHROUGH") & (original_df['MOP'] == "RUNTHROUGHRUNTHROUGH")]
dup_cols = df.columns[df.columns.duplicated()]
print(dup_cols)
df_violation = original_df[
    original_df['MOP'].isin([
        "VIOLATION",
        "FLEET",
        "CONVEY",
        "RUNTHROUGH",
        "RUNTHROUGHRUNTHROUGH"
    ])
]
# Combine df (processed) with violation rows for trip count calculation
df_combined_for_trips = pd.concat([df, df_violation], ignore_index=True)

# Group by 'Veh Reg No.' to count the number of trips, including violations
trip_df = df_combined_for_trips.groupby('Veh Reg No.').size().reset_index(name='Trip Rules')

# Merge with the 'Trip Rules' DataFrame
df_merged = df_merged.merge(trip_df, on='Veh Reg No.', how='left')

# Add the "Exempt txn" column (Trip Rules - TRI 3)
df_merged['Exempt txn'] = df_merged['Trip Rules'] - df_merged['TRI 3'].fillna(0).astype(int)

# Add the "Trip Group" column based on the "Exempt txn" logic
df_merged['Trip Group'] = df_merged['Exempt txn'].apply(lambda x: "0-5" if 0 <= x <= 5 else ">5")

# Get base output path and save the final merged DataFrame to an Excel file
base_output_path = get_base_output_path()
output_path = f"{base_output_path}/combined_with_rf3.xlsx"
#df_merged.to_excel(output_path, index=False)
df_merged.to_excel(output_path, index=False, sheet_name="Combined with RF 3", engine="openpyxl")
print(f"Data saved to {output_path}")