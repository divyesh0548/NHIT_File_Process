import pandas as pd
from config import get_file_path, output_path
from data_processing import load_and_preprocess_data

def main():
    # Get file path from the user
    file_path = get_file_path()

    # Load and preprocess data
    df, exempt_vehicle_list = load_and_preprocess_data(file_path)

    # Replace NaN with blank in 'Description'
    df['Description'] = df['Description'].fillna('')

    # =========================
    # FILTER EXEMPT
    # =========================
    df_filtered = df[df['MOP Final'] == 'EXEMPT']

    # Select needed columns
    df_selected = df_filtered[['Veh Reg No.', 'Description']]

    # =========================
    # GROUP & KEEP DESCRIPTIONS
    # =========================
    df_grouped = (
        df_selected.groupby('Veh Reg No.')
        .agg(
            Ground_Total=('Veh Reg No.', 'size'),
            All_Rows=('Description', list)
        )
        .reset_index()
    )

    # =========================
    # EXPLODE DESCRIPTIONS
    # =========================
    df_expanded = df_grouped.explode('All_Rows').reset_index(drop=True)
    df_expanded = df_expanded.rename(columns={'All_Rows': 'Description'})

    # =========================
    # COUNT EACH DESCRIPTION
    # =========================
    df_grouped_1 = (
        df_expanded.groupby(['Veh Reg No.', 'Ground_Total', 'Description'])
        .size()
        .reset_index(name='Count')
    )

    # Replace empty Description with 'Others'
    df_grouped_1['Description'] = df_grouped_1['Description'].replace('', 'Others')

    # =========================
    # TRI LOGIC
    # =========================
    df_grouped_1['TRI 1'] = df_grouped_1['Ground_Total'].apply(lambda x: 1 if x == 1 else 0)
    df_grouped_1['TRI 2'] = df_grouped_1.apply(
        lambda row: 1 if row['Count'] == row['Ground_Total'] else 0,
        axis=1
    )

    # =========================
    # PIVOT DESCRIPTION
    # =========================
    df_pivoted = df_grouped_1.pivot_table(
        index=['Veh Reg No.', 'Ground_Total', 'TRI 1', 'TRI 2'],
        columns='Description',
        values='Count',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    # =========================
    # TRI 3 (ETC)
    # =========================
    df_etc = df[df['MOP Final'] == 'ETC']
    df_tri_3 = (
        df_etc.groupby('Veh Reg No.')
        .size()
        .reset_index(name='TRI 3')
    )

    # =========================
    # MERGE TRI 3
    # =========================
    df_merged = pd.merge(df_pivoted, df_tri_3, how='left', on='Veh Reg No.')

    df_merged['TRI 3'] = df_merged['TRI 3'].fillna(0).astype(int)

    # Create TRI 2 & 3 code
    df_merged['TRI 2 & 3'] = df_merged['TRI 2'].astype(str) + df_merged['TRI 3'].astype(str)

    # Final output
    result_df = df_merged.drop(columns=['TRI 2 & 3'])

    # =========================
    # SAVE FILES
    # =========================
    result_df.to_excel(output_path("RF3 Pivot.xlsx"), index=False)

    rf3_1 = df_merged[df_merged['TRI 2 & 3'] == '00'][['Veh Reg No.', 'TRI 2 & 3']].drop_duplicates()
    rf3_1.to_excel(output_path("rf3_1.xlsx"), index=False)

    rf3_2 = df_merged[['Veh Reg No.', 'TRI 1', 'TRI 2', 'TRI 3']].drop_duplicates()
    rf3_2.to_excel(output_path("rf3_2.xlsx"), index=False)

    print("RF3 files generated successfully.")

if __name__ == "__main__":
    main()