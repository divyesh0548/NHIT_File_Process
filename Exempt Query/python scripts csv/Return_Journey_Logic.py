import pandas as pd
from config import get_return_journey_input_path, output_path, put_file_name
from data_processing import SIDE_1_LANES, SIDE_2_LANES

input_path = get_return_journey_input_path()
df = pd.read_excel(input_path, sheet_name="Combined with RF 3", header=0)

df["Date & Time"] = pd.to_datetime(df["Date & Time"], errors="coerce")
df = df.sort_values(by=["Veh Reg No.", "Date & Time"])


def determine_side(lane_no):
    if lane_no in SIDE_1_LANES:
        return "Side 1"
    if lane_no in SIDE_2_LANES:
        return "Side 2"
    return "Unknown"


df["Side"] = df["Lane No"].apply(determine_side)

journey_type = []
for i in range(len(df)):
    if i == 0 or df.iloc[i]["Veh Reg No."] != df.iloc[i - 1]["Veh Reg No."]:
        journey_type.append("First Journey")
    else:
        time_difference = df.iloc[i]["Date & Time"] - df.iloc[i - 1]["Date & Time"]
        current_side = df.iloc[i]["Side"]
        previous_side = df.iloc[i - 1]["Side"]

        if time_difference.total_seconds() <= 86400:
            if journey_type[-1] == "Return Journey":
                journey_type.append("First Journey")
            elif current_side != previous_side:
                journey_type.append("Return Journey")
            else:
                journey_type.append("First Journey")
        else:
            journey_type.append("First Journey")

df["Journey Type"] = journey_type
print(df)

out_file = output_path(put_file_name())
df.to_excel(out_file, index=False, sheet_name="Combined with RF 3")

result_df = pd.read_excel(output_path("RF3 Pivot.xlsx"))
with pd.ExcelWriter(out_file, engine="openpyxl", mode="a") as writer:
    result_df.to_excel(writer, sheet_name="RF3 Pivot", index=False, engine="openpyxl")
