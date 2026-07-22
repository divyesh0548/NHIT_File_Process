import pandas as pd
from config import (
    input_path,
    output_path,
    parse_pass_date_series,
    standardize_pass_file_columns,
)

combined_rf_df = pd.read_excel(output_path("MP Output.xlsx"))
working_file_df = pd.read_excel(input_path("LT_Pass.xlsx"))
working_file_df = standardize_pass_file_columns(working_file_df)

combined_rf_df["Date & Time Copy"] = combined_rf_df["Date & Time"]
combined_rf_df["Date & Time"] = pd.to_datetime(combined_rf_df["Date & Time"]).dt.date

merged_df = pd.merge(
    combined_rf_df,
    working_file_df,
    left_on="Veh Reg No.",
    right_on="Chassis/ Vehicle No",
    how="left",
)
merged_df = merged_df[["Veh Reg No.", "Date & Time", "Start Date", "End Date"]]

filtered_df = merged_df[merged_df["Start Date"].notna()].copy()
filtered_df["Start Date"] = parse_pass_date_series(
    filtered_df["Start Date"], "LT start date"
)
filtered_df["End Date"] = parse_pass_date_series(
    filtered_df["End Date"], "LT end date"
)
filtered_df = filtered_df.dropna(subset=["Start Date", "End Date"])

sorted_df = filtered_df.sort_values(by=["Veh Reg No.", "Date & Time"])


def check_validity(row, df):
    current_date = row["Date & Time"]
    current_veh_reg_no = row["Veh Reg No."]
    relevant_rows = df[df["Veh Reg No."] == current_veh_reg_no]
    start_dates = relevant_rows["Start Date"].tolist()
    end_dates = relevant_rows["End Date"].tolist()
    is_valid = any(start <= current_date <= end for start, end in zip(start_dates, end_dates))
    return "Valid" if is_valid else "Not Valid"


sorted_df["Date Validity Check LT"] = sorted_df.apply(check_validity, df=sorted_df, axis=1)

grouped_df = (
    sorted_df.groupby(["Date & Time", "Veh Reg No.", "Date Validity Check LT"])
    .size()
    .reset_index(name="Count")
)

final_df = pd.merge(
    combined_rf_df,
    grouped_df[["Veh Reg No.", "Date & Time", "Date Validity Check LT"]],
    on=["Veh Reg No.", "Date & Time"],
    how="left",
)

final_df = final_df.drop(columns=["Date & Time"])
final_df = final_df.rename(columns={"Date & Time Copy": "Date & Time"})

columns_order = final_df.columns.tolist()
columns_order.remove("Date & Time")
columns_order.insert(2, "Date & Time")
final_df = final_df[columns_order]

final_df.to_excel(output_path("LT_output.xlsx"), index=False, sheet_name="Combined with RF 3")
