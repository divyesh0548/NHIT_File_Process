import pandas as pd
import sys
from config import get_file_path, get_base_output_path  # Import the new function
from data_processing import load_and_preprocess_data

def process_side(df, exempt_vehicle_list, side):
    # Step 1: Join `df` with `exempt_vehicle_list`
    merged_df = pd.merge(df, exempt_vehicle_list, on="Veh Reg No.", how="inner")
    merged_df = merged_df[merged_df['Side'] == side]

    # Print shape for validation
    print(f"Shape of {side}: {merged_df.shape}")

    # Step 4: Sort rows
    merged_df = merged_df.sort_values(by=["Veh Reg No.", "Date & Time"])

    # Get previous record data by shifting rows
    merged_df['Date & Time_Prev'] = merged_df.groupby('Veh Reg No.')['Date & Time'].shift(1)
    merged_df['MOP Final_Prev'] = merged_df.groupby('Veh Reg No.')['MOP Final'].shift(1)
    #merged_df['Txn No._Prev'] = merged_df.groupby('Veh Reg No.')['Trans No.'].shift(1)


    # Get next record data by shifting rows in the opposite direction
    merged_df['Date & Time_Next'] = merged_df.groupby('Veh Reg No.')['Date & Time'].shift(-1)
    merged_df['MOP Final_Next'] = merged_df.groupby('Veh Reg No.')['MOP Final'].shift(-1)
    #merged_df['Txn No._Next'] = merged_df.groupby('Veh Reg No.')['Trans No.'].shift(-1)


    # Convert column types
    merged_df['Date & Time'] = pd.to_datetime(merged_df['Date & Time'])
    merged_df['Date & Time_Prev'] = pd.to_datetime(merged_df['Date & Time_Prev'])
    merged_df['Date & Time_Next'] = pd.to_datetime(merged_df['Date & Time_Next'])

    # Step 5: Filter rows where 'MOP Final' is "EXEMPT"
    filtered_df = merged_df[merged_df['MOP Final'] == 'EXEMPT']

    # Step 6: Apply custom conditions to the filtered DataFrame
    filtered_df['Custom'] = filtered_df.apply(apply_custom_conditions, axis=1)

    # Return the processed DataFrame
    return filtered_df

def apply_custom_conditions(row):
    if (row['MOP Final_Prev'] and
        (row['Date & Time'] - row['Date & Time_Prev']).total_seconds() / 3600 <= 0.25):
        return "Non Exception 1"
    elif (row['MOP Final_Next'] in ['ETC', 'Cash'] and
          (row['Date & Time_Next'] - row['Date & Time']).total_seconds() / 3600 <= 16):
        return "Non Exception 2"
    elif (row['MOP Final_Next'] == 'Exempt' and
          (row['Date & Time_Next'] - row['Date & Time']).total_seconds() / 3600 <= 0.25):
        return "Non Exception 3"
    else:
        return "Exception"

def main():
    # Get file path from the user
    file_path = get_file_path()

    # Load and preprocess the data
    df, exempt_vehicle_list = load_and_preprocess_data(file_path)

    # Display the first 15 rows
    print(df.head(15))

    # Example of calling `process_side` function
    result_side_1 = process_side(df, exempt_vehicle_list, 'Side 1')
    result_side_2 = process_side(df, exempt_vehicle_list, 'Side 2')

    # Combine and save results
    final_result = pd.concat([result_side_1, result_side_2])

    # Get the base output path
    base_output_path = get_base_output_path()  # Get the base output directory

    # Construct the full output path
    output_file_path = f"{base_output_path}/updated_new_Combined_Side_Results.xlsx"  # Change as necessary for your structure

    # Save the final results to the specified output path
    final_result.to_excel(output_file_path, index=False)

    print(f"Filtered DataFrame with updated column names has been saved to '{output_file_path}'.")

if __name__ == "__main__":
    main()