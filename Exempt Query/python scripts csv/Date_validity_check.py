import pandas as pd
from config import get_file_path, get_base_output_path  # Import get_base_output_path
from datetime import datetime

base_output_path = get_base_output_path()


# Load the Excel files (you need to specify the correct file paths)
combined_rf_df = pd.read_excel(base_output_path + "/combined_with_rf3.xlsx")
working_file_df = pd.read_excel(base_output_path + "/Pass_MP.xlsx")

#combined_rf_df["Txn ID"] = combined_rf_df["Txn ID"].astype(str)

# Step 1: Create a copy of 'Date & Time' called 'Date & Time Copy'
combined_rf_df['Date & Time Copy'] = combined_rf_df['Date & Time']

# Convert 'Date & Time' in combined_rf_df to just the date
combined_rf_df['Date & Time'] = pd.to_datetime(combined_rf_df['Date & Time']).dt.date

combined_rf_df.head(10)

# Perform the left join on 'Veh Reg No.' from combined_rf_df and 'Chassis/ Vehicle No' from working_file_df
merged_df = pd.merge(combined_rf_df, working_file_df, left_on='Veh Reg No.', right_on='Chassis/ Vehicle No', how='left')

# Expand the "Working File" columns (Start Date, End Date)
# Assuming these columns already exist after the merge
merged_df = merged_df[['Veh Reg No.', 'Date & Time', 'Start Date', 'End Date']]

# Filter rows where Start Date is not null
filtered_df = merged_df[merged_df['Start Date'].notna()]

# Convert 'Start Date' and 'End Date' to date format if they are not already
filtered_df['Start Date'] = pd.to_datetime(filtered_df['Start Date']).dt.date
filtered_df['End Date'] = pd.to_datetime(filtered_df['End Date']).dt.date

# Sort the rows based on 'Veh Reg No.' and 'Date & Time'
sorted_df = filtered_df.sort_values(by=['Veh Reg No.', 'Date & Time'])

# Perform the date validity check
def check_validity(row, df):
    current_date = row['Date & Time']
    current_veh_reg_no = row['Veh Reg No.']

    # Get all rows for the same Veh Reg No.
    relevant_rows = df[df['Veh Reg No.'] == current_veh_reg_no]

    # Extract Start Date and End Date columns as lists of dates
    start_dates = relevant_rows['Start Date'].tolist()
    end_dates = relevant_rows['End Date'].tolist()

    # Check if current Date falls within any valid periods
    is_valid = any(start <= current_date <= end for start, end in zip(start_dates, end_dates))

    return "Valid" if is_valid else "Not Valid"

# Apply the validity check for each row
sorted_df['Date Validity Check MP'] = sorted_df.apply(check_validity, df=sorted_df, axis=1)

# Group by 'Date & Time', 'Veh Reg No.', and 'Date Validity Check' and count occurrences
grouped_df = sorted_df.groupby(['Date & Time', 'Veh Reg No.', 'Date Validity Check MP']).size().reset_index(name='Count')

# Merge grouped_df back with combined_rf_df on 'Veh Reg No.' and 'Date & Time' (Left join to mimic Power Query behavior)
final_df = pd.merge(combined_rf_df, grouped_df[['Veh Reg No.', 'Date & Time', 'Date Validity Check MP']], 
                    on=['Veh Reg No.', 'Date & Time'], how='left')

final_df = final_df.drop(columns= ['Date & Time'])

# Rename 'Date & Time Copy' to 'Date & Time'
final_df = final_df.rename(columns={'Date & Time Copy': 'Date & Time'})

# Create a list of columns in the desired order
columns_order = final_df.columns.tolist()  # Get the current column order
columns_order.remove("Date & Time")  # Remove 'Date & Time' from its current position
columns_order.insert(2, "Date & Time")  # Insert 'Date & Time' at the 3rd position

# Reorder the DataFrame
final_df = final_df[columns_order]

base_output_path = get_base_output_path()
    
# Save the final_df to Excel
final_output_path = f"{base_output_path}/MP Output.xlsx"  # Adjust file name as necessary
final_df.to_excel(final_output_path, index=False, sheet_name="Combined with RF 3")
