from input_valid_invalid import df_combined, rates_df, output_invalid_path, output_valid_path
import pandas as pd

print("Valid/Invalid Lookup: loaded input data and rates.")

# ============================
# BUILD SINGLE JOURNEY RATE LOOKUP
# ============================
single_rates = rates_df[rates_df["Attribute"] == "Single Journey"]
rate_lookup = dict(zip(single_rates["TC Class"], single_rates["value"]))

# ============================
# MAP NPCI → TC CLASS
# ============================
def map_npci_to_tc(npci):
    if npci in [
        "Car / Jeep / Van","Car/Jeep/Van","VC4 - Car / Jeep / Van","VC4",
        "Tata Ace and Similar mini Light Commercial Vehicle",
        "Tata Ace or Similar Mini LCV","VC20","VC20 - Tata Ace or similar mini LCV"
    ]:
        return "Car"

    elif npci in ["Mini-Bus","LCV","Light Commercial vehicle 2-axle","VC5","VC9"]:
        return "LCV"

    elif npci in ["Truck 2 - axle","Bus 2-axle","VC10","VC7","VC10 - Truck 2 Axle", "VC7 - Bus 2 Axle"]:
        return "Trk 2 Axle"

    elif npci in ["Truck 3 - axle","Truck 3-Axle","VC8","VC11", "Bus 3-axle", "VC11 - Truck 3 Axle", "VC8 - Bus 3 Axle"]:
        return "Truck 3 axle"

    elif npci in [
        "Truck 4 - axle","Truck 4-Axle","Truck 5 - axle","Truck 6 - axle",
        "VC12","VC13","VC14"
    ]:
        return "MAV"
    else:
        return None


# ============================
# VEHICLE CLASS CORRECTION
# ============================
def process_vehicle_classification(df):
    df = df.copy()
    df["weight"] = df["weight"].replace(0, None)

    def correct_vehicle_class(weight):
        if pd.isna(weight) or weight == "":
            return "MAV"
        elif weight <= 7500:
            return "LMV"
        elif weight <= 12000:
            return "LCV"
        elif weight <= 18500:
            return "Truck 2 axle"
        elif weight <= 28000:
            return "Truck 3 axle"
        elif 28000 < weight <= 60000:
            return "MAV"
        else:
            return "Truck 7 axle"

    df["Correct Vehicle Class"] = df["weight"].apply(correct_vehicle_class)

    def classify_custom_2(row):
        if row["Custom.1"] == "Less than equal to 12 seating capacity":
            return "LMV"
        elif row["Custom.1"] == "More than 32 seating capacity":
            return "BUS"
        elif row["Custom.1"] == "exceeds 12 but less than equal to 32":
            return "LCV"
        else:
            return row["Correct Vehicle Class"]

    df["Custom.2"] = df.apply(classify_custom_2, axis=1)
    df.drop(columns=['Correct Vehicle Class'], inplace=True)
    df.rename(columns={"Custom.2": "Correct Vehicle Class"}, inplace=True)
    return df


# ============================
# SPLIT VALID / INVALID
# ============================
df_invalid = df_combined[df_combined['Custom.4'] == "Invalid"]
df_invalid = process_vehicle_classification(df_invalid)

df_valid = df_combined[df_combined['Custom.4'] == "Valid"]
df_valid = process_vehicle_classification(df_valid)

print(f"Valid/Invalid Lookup: split data -> invalid={len(df_invalid)}, valid={len(df_valid)}")

# ============================
# REMOVE ZERO AMOUNT ROWS
# ============================
df_invalid["Custom.2"] = df_invalid.apply(
    lambda row: "Remove" if row["Net Settlement Amt"] == 0 and 
                row["Journey Type"] in [
                    "Cont. Journey","Single Journey","Local Conti/Single",
                    "Local Cont.","Local Single","Return Journey","SINGLE",
                    "RETURN","LocalSingleC","FULL","LOCAL TRIP",
                    "SINGLE JOURNEY","RETURN PASS","Return","Single", "LocalSingleC", "Local Vehicle(Commercial)", "Local Vehicle Commercial"
                ] 
                else "Keep", 
    axis=1
)

df_invalid = df_invalid[df_invalid['Custom.2'] != "Remove"]
df_invalid.drop(columns=['Custom.2'], inplace=True)


# ============================
# UPDATE NET SETTLEMENT (DYNAMIC RATES)
# ============================
def update_net_settlement(row):

    if row["Journey Type"] in [
        "DISCOUNTED","DISCOUNTMP","Monthly Pass-50 trips",
        "Monthly Pass Comm","Monthly Pass Non Comm","Local Non Commercial",
        "Local20KM",
        "MonthlyExempted","Discount Pass", "Monthly Pass", "Local Vehicle(Commercial)", "Local Pass Non Comm", "NON_FIN", "Monthly Pass 20 Km"
    ]:

        tc_class = map_npci_to_tc(row["NPCI Class Desc"])

        if tc_class and tc_class in rate_lookup:
            return rate_lookup[tc_class]   # 🔥 from rates file
        else:
            return row["Net Settlement Amt"]

    else:
        return row["Net Settlement Amt"]


df_invalid["updated net settlement amount"] = df_invalid.apply(update_net_settlement, axis=1)
df_invalid.drop_duplicates(inplace=True)

df_invalid['Impact'] = df_invalid['Rate'] - df_invalid['updated net settlement amount']

# remove rows where Impact = 0
df_invalid = df_invalid[(df_invalid['Impact'] != 0) & (df_invalid['Impact'] > 0)]

# Step 2: remove unwanted vehicle classes
exclude_classes = [
    "BUS",
    "CONSTRUCTION EQUIPMENT VEHICLE",
    "OMNI BUS",
    "TOWER WAGONS",
    "VEHICLE FITTED WITH AIR GENERATOR",
    "VEHICLE FITTED WITH COMPRESSOR",
    "VEHICLE FITTED WITH RIG"
]

df_invalid = df_invalid[
    ~df_invalid['vehicle class']
    .fillna('') 
    .str.strip()
    .str.upper()
    .isin(exclude_classes)
]

print(f"Valid/Invalid Lookup: final invalid rows={len(df_invalid)}, final valid rows={len(df_valid)}")

df_valid = df_valid[df_valid['Rate'] != 0]

# ============================
# EXPORT FILES
# ============================
df_invalid.to_csv(output_invalid_path, index=False)
df_valid.to_csv(output_valid_path, index=False)
print(f"Valid/Invalid Lookup: exported invalid -> {output_invalid_path}")
print(f"Valid/Invalid Lookup: exported valid -> {output_valid_path}")
