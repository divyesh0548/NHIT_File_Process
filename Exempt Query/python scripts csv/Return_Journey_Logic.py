import pandas as pd
from config import get_base_output_path, put_file_name
from data_processing import SIDE_1_LANES, SIDE_2_LANES
# import re
# from Date_validity_check_LT import final_df
# from RF3_condition import main 

base_output_path = get_base_output_path()

# Replace this with the path to your Excel file
file_path = base_output_path + '/'  "MP Output.xlsx"

# Load the Excel file
#df = final_df
df = pd.read_excel(file_path, sheet_name='Combined with RF 3', header=0)
#df["Txn ID"] = df["Txn ID"].astype(str)


df['Date & Time'] = pd.to_datetime(df['Date & Time'], errors='coerce')

# Sort the data by Veh Reg No. and Date & Time to ensure chronological order
df = df.sort_values(by=['Veh Reg No.', 'Date & Time'])

# Define a function to determine the side based on the lane number
def determine_side(lane_no):
    if lane_no in SIDE_1_LANES:
        return 'Side 1'
    elif lane_no in SIDE_2_LANES:     
        return 'Side 2'
    else:
        return 'Unknown'  # Handle any unexpected lane numbers

# Determine sides
df['Side'] = df['Lane No'].apply(determine_side)

# Initialize a list to store the journey type (First Journey / Return Journey)
journey_type = []

# Iterate over the rows to classify each transaction
for i in range(len(df)):
    if i == 0 or df.iloc[i]['Veh Reg No.'] != df.iloc[i-1]['Veh Reg No.']:
        # First row or first transaction of a different vehicle
        journey_type.append('First Journey')
    else:
        # Check if the current transaction is for the same vehicle
        time_difference = df.iloc[i]['Date & Time'] - df.iloc[i-1]['Date & Time']
        current_side = df.iloc[i]['Side']
        previous_side = df.iloc[i-1]['Side']

        # Determine if it's a Return Journey based on side and time interval
        if time_difference.total_seconds() <= 86400:  # within 24 hours
            if journey_type[-1] == 'Return Journey':
                journey_type.append('First Journey')
            elif current_side != previous_side:
                journey_type.append('Return Journey')
            else:
                journey_type.append('First Journey')
        else:
            journey_type.append('First Journey')

# Add the journey type to the dataframe
df['Journey Type'] = journey_type

# Display the resulting dataframe
print(df)

base_output_path = get_base_output_path()

file_name = put_file_name()

# Construct the output path for the processed data
output_path = f"{base_output_path}/{file_name}"
df.to_excel(output_path, index=False, sheet_name= "Combined with RF 3")

rf3_pivot_file_path = f"{base_output_path}/RF3 Pivot.xlsx"
result_df = pd.read_excel(rf3_pivot_file_path)
#rf3_1, rf3_2, result_df = main()

# Write both DataFrames to the same Excel file in different sheets
with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
    result_df.to_excel(writer, sheet_name='RF3 Pivot', index=False, engine="openpyxl")

