import pandas as pd
import numpy as np
import os
from pathlib import Path
from dotenv import load_dotenv
from sqlalchemy import create_engine

REPO_ROOT = Path(__file__).resolve().parents[3]
load_dotenv(REPO_ROOT / ".env")

# -----------------------------
# Input/Output Paths
# -----------------------------
OUTPUT_DIR = os.environ.get("VALID_INVALID_OUTPUT_DIR")
output_dir = Path(OUTPUT_DIR) if OUTPUT_DIR else Path(__file__).resolve().parent / "veeravalli valid invalid output"
output_dir.mkdir(parents=True, exist_ok=True)
output_invalid_path = output_dir / "invalid_table.csv"
output_valid_path = output_dir / "valid_table.csv"

# -----------------------------
# RDS Connection Details
# -----------------------------
DB_HOST = os.environ.get("RDS_HOST")
DB_PORT = os.environ.get("RDS_PORT", "5432")
DB_NAME = os.environ.get("RDS_DB_NAME")
DB_USER = os.environ.get("RDS_USER")
DB_PASSWORD = os.environ.get("RDS_PASSWORD")
TABLE_NAME = os.environ.get("RDS_TABLE_NAME", "checkpostmaster")

# -----------------------------
# Load Rates File
# -----------------------------
RATES_INPUT_PATH = os.environ.get("VALID_INVALID_RATES_PATH", r"Veravalli Rates.xlsx")

df = pd.read_excel(
    RATES_INPUT_PATH,
    skiprows=1
)

# -----------------------------
# Process Vehicle Lifecycle
# -----------------------------
LIFECYCLE_INPUT_PATH = os.environ.get("VALID_INVALID_LIFECYCLE_PATH", r"merged_output.csv")


def process_vehicle(lifecycle_input_path):
    missing_db_config = [
        key
        for key, value in {
            "RDS_HOST": DB_HOST,
            "RDS_DB_NAME": DB_NAME,
            "RDS_USER": DB_USER,
            "RDS_PASSWORD": DB_PASSWORD,
        }.items()
        if not value
    ]
    if missing_db_config:
        raise EnvironmentError(
            "Missing required database configuration in .env: "
            + ", ".join(missing_db_config)
        )

    df_l = pd.read_csv(lifecycle_input_path, low_memory=False)

    # Read checkpostmaster from RDS in chunks to support million+ row tables.
    engine = create_engine(
        f"postgresql+psycopg2://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}"
    )
    query = f'SELECT * FROM "{TABLE_NAME}"'

    chunk_frames = []
    for chunk in pd.read_sql_query(query, con=engine, chunksize=100000):
        chunk_frames.append(chunk)

    if not chunk_frames:
        raise ValueError(f"No data found in table: {TABLE_NAME}")

    df_c = pd.concat(chunk_frames, ignore_index=True)
    engine.dispose()

    # Normalize DB column names to expected names used in the processing logic.
    def normalize_column_name(name):
        return "".join(ch for ch in str(name).lower() if ch.isalnum())

    normalized_to_original = {normalize_column_name(col): col for col in df_c.columns}
    required_map = {
        "Unique Vehicle Number": ["uniquevehiclenumber", "vehicleregno", "vrn"],
        "vehicle class": ["vehicleclass", "npciclassdesc", "class"],
        "weight": ["weight", "seatingcapacity", "capacity"],
    }

    rename_columns = {}
    for expected_col, candidates in required_map.items():
        for candidate in candidates:
            if candidate in normalized_to_original:
                rename_columns[normalized_to_original[candidate]] = expected_col
                break

    df_c = df_c.rename(columns=rename_columns)

    missing_required = [
        col for col in ["Unique Vehicle Number", "vehicle class", "weight"] if col not in df_c.columns
    ]
    if missing_required:
        raise KeyError(
            f"Missing required checkpostmaster columns: {missing_required}. "
            f"Available columns: {list(df_c.columns)}"
        )

    # Apply logic from checkpostmaster.py
    df_c['Unique Vehicle Number'] = df_c['Unique Vehicle Number'].astype(str).str.replace('.', '', regex=False)
    df_c['weight'] = df_c['weight'].apply(lambda x: str(x) if pd.notnull(x) else None)

    df_c = df_c[
        ~df_c['vehicle class'].isin([
            "---Select Vehicle Class---", 
            "THREE WHEE! ER{PASSENGER)", 
            "THREE WHEELER(GOODS)", 
            "THREE WHEELER(PASSENGER)", 
            "THREE WHEELER{GOODS)", 
            "—Select Vehicle Class—",
            "--Select Vehicle Class--",
            "----Select Vehicle Class----",
            "---Select Vehicle Class---",
            "----Select Vehicle Class----"
        ]) & 
        (df_c['Unique Vehicle Number'] != "PB11U4131") & 
        (df_c['Unique Vehicle Number'] != "MH04DT4250") & 
        (df_c['Unique Vehicle Number'] != "MP13WA9975") & 
        (df_c['Unique Vehicle Number'] != "MP13WA9928")
    ]

    df_c = df_c[~((df_c['vehicle class'].isnull()) & (df_c['weight'].isnull()))]
    df_c = df_c.drop_duplicates()
    df_c = df_c.drop_duplicates(subset='Unique Vehicle Number')

    df_l.columns = df_l.columns.str.strip()

    # Dynamic Column Renaming
    rename_map = {
        'Journey Type': ['MOP', 'Fare Type', 'TRIPTYPEDISCRIPTION', 'Trip Type Description', 'fare type'],
        'Net Settlement Amt': ['SETTLED Amount (Rs.)', 'Settled Amount (Rs.)', 'Net Settlement Amount', 'accepted amount'],
        'NPCI Class Desc': ['Mapper Vehicle Class', 'Vehicle Class', 'mapper vehicle class'],
        'Veh Reg No.': [
            'NPCI VRN',
            'vehicle registration number',
            'Vehicle Registration Number',
            'Vehicle Reg No',
            'Vehicle Reg No.',
            'Veh Reg No',
            'Registration Number'
        ]
    }

    # Case/space tolerant rename to handle source files with inconsistent headers.
    normalized_columns = {" ".join(col.strip().lower().split()): col for col in df_l.columns}

    for target, alternatives in rename_map.items():
        if target not in df_l.columns:
            for alt in alternatives:
                normalized_alt = " ".join(alt.strip().lower().split())
                if normalized_alt in normalized_columns:
                    df_l = df_l.rename(columns={normalized_columns[normalized_alt]: target})
                    break

    # Filtering Journey Types
    if 'Journey Type' in df_l.columns:
        df_l = df_l[~df_l['Journey Type'].isin([
            'ANNUALPASS','ANNUALPASS:LocalSingleC','ANNUALPASS:RETURN',
            'ANNUALPASS:SINGLE','GlobalExempt','ANNUALPASS_RETURN',
            'ANNUALPASS_FULL','Annual Pass Single','GlobalExemption', 'Global Exemption', 'Annual Pass Trip', '0',
            '0.0000',
            'Rejected Trip',
            'Rejected Trip Commercial', 'Annual Pass Trip Personal', 'Rejected Trip Personal', 'EXEMPTED'
        ])]

    status_col = next((col for col in ["Settlement Type", "Txn Status", "Status", "Transaction Status", "transaction status"] if col in df_l.columns), None)
    if status_col:
        df_l = df_l[df_l[status_col].isin(["ACCEPTED", "Accepted", "SETTLED"])]

    df_l["Veh Reg No."] = df_l["Veh Reg No."].str.upper().str.strip()
    df_c["Unique Vehicle Number"] = df_c["Unique Vehicle Number"].str.upper().str.strip()

    df_merged = df_l.merge(
        df_c,
        left_on="Veh Reg No.",
        right_on="Unique Vehicle Number",
        how="left"
    ).drop(columns=['Unique Vehicle Number'])

    df_merged["updated journey type"] = df_merged["Journey Type"].apply(lambda x:
        "Local Conti/Single" if x in ["Local Cont.", "Local Single", "Local Vehicle(Commercial)", "LocalSingleC", "Local Single Journey Commercial", "Local Vehicle Commercial"]
        else "Cont. Journey" if x in ["Return Journey", "RETURN", "RETURN PASS", "Return", "Cont. Journey", "Return Journey Personal"]
        else "Single Journey"
    )

    df_final = df_merged[
       # (df_merged["Settlement Type"] == "Accepted") &
        (df_merged["Veh Reg No."].str.len() <= 12) &
        ((df_merged["vehicle class"].notna()) |
         (df_merged["weight"].notna() & (df_merged["weight"] != "nan")))
    ]

    df_final["weight"] = pd.to_numeric(df_final["weight"], errors='coerce').astype("Int64")
    return df_final


df_result = process_vehicle(LIFECYCLE_INPUT_PATH)

# -----------------------------
# Rates Transform
# -----------------------------
def rates(df):
    df = df.rename(columns={
        "Return Journey": "Cont. Journey",
        "Local Conti/Local Single": "Local Conti/Single"
    })
    df = df.melt(
        id_vars=['TC Class', 'Weight/Capacity', 'Vehicle Class'],
        var_name='Attribute'
    )
    return df

rates_df = rates(df)

# -----------------------------
# Helper: Get rate dynamically
# -----------------------------
def get_rate_from_rates(tc_class, journey_type):
    row = rates_df.loc[
        (rates_df["TC Class"] == tc_class) &
        (rates_df["Attribute"] == journey_type),
        "value"
    ]
    return row.values[0] if not row.empty else 0


# -----------------------------
# Vehicle lifecycle (NON BUS)
# -----------------------------
def vehicle_lifecycle(df_result):

    df_result = df_result[~df_result['vehicle class'].isin([
        "BUS"," BUS","CAMPER VAN / TRAILER","CASH VAN","Error",
        "MAXI CAB","MOTOR CAB","OMNI BUS","OMNIBUS","Invalid"
    ])]

    vehicle_classes = {
        "CONSTRUCTION EQUIPMENT VEHICLE","EARTH MOVING EQUIPMENT",
        "VEHICLE FITTED WITH AIR GENERATOR","VEHICLE FITTED WITH COMPRESSOR",
        "VEHICLE FITTED WITH RIG","EXCAVATOR","TREE TRIMMING VEHICLE","TOWER WAGONS"
    }

    conditions = [
        df_result["vehicle class"].isin(vehicle_classes) & (df_result["updated journey type"] == "Single Journey"),
        df_result["vehicle class"].isin(vehicle_classes) & (df_result["updated journey type"] == "Cont. Journey"),
        df_result["vehicle class"].isin(vehicle_classes) & (df_result["updated journey type"] == "Local Conti/Single"),
        (df_result["vehicle class"] == "CRANE MOUNTED VEHICLE") & (df_result["updated journey type"] == "Single Journey"),
        (df_result["vehicle class"] == "CRANE MOUNTED VEHICLE") & (df_result["updated journey type"] == "Cont. Journey"),
        (df_result["vehicle class"] == "CRANE MOUNTED VEHICLE") & (df_result["updated journey type"] == "Local Conti/Single")
    ]

    values = [
        get_rate_from_rates("MAV", "Single Journey"),
        get_rate_from_rates("MAV", "Cont. Journey"),
        get_rate_from_rates("MAV", "Local Conti/Single"),
        get_rate_from_rates("LCV", "Single Journey"),
        get_rate_from_rates("LCV", "Cont. Journey"),
        get_rate_from_rates("LCV", "Local Conti/Single"),
    ]

    df_result["Rate 1"] = np.select(conditions, values, default=None)

    def classify_weight(weight):
        if pd.isna(weight):
            return None
        elif weight <= 7500:
            return "<=7500 Kgs"
        elif weight <= 12000:
            return ">7500 Kgs but <=12,000 Kgs"
        elif weight <= 18500:
            return "> 12,000 Kgs but <= 18,500 Kgs"
        elif weight <= 28000:
            return "> 18,500 Kgs but <= 28,000 Kgs"
        elif weight <= 60000:
            return "> 28,000 Kgs but <= 60,000 Kgs"
        else:
            return ">60,000 Kgs"

    df_result["Custom.1"] = df_result["weight"].apply(classify_weight)
    return df_result


df_final_1 = vehicle_lifecycle(df_result)

# -----------------------------
# BUS logic
# -----------------------------
def vehicle_bus(df_result):

    df_result = df_result[df_result['vehicle class'].isin([
        " BUS","BUS","OMNI BUS","OMNIBUS"
    ])]

    df_result = df_result.dropna(subset=["weight"])

    def classify_weight(row):
        weight = row["weight"]
        jt = row["updated journey type"]

        if weight > 32:
            return "More than 32 seating capacity"
        elif 12 < weight <= 32:
            return "exceeds 12 but less than equal to 32"
        else:
            return "Less than equal to 12 seating capacity"

    df_result["Custom.1"] = df_result.apply(classify_weight, axis=1)
    return df_result


df_bus = vehicle_bus(df_result)

# -----------------------------
# Get final rate
# -----------------------------
def get_rate(df, rates_df, is_bus=False):

    df = df.merge(
        rates_df[["Attribute", "Weight/Capacity", "value"]],
        left_on=["updated journey type", "Custom.1"],
        right_on=["Attribute", "Weight/Capacity"],
        how="left"
    ).drop(columns=["Attribute", "Weight/Capacity"])

    df["value"] = df["value"].fillna(0).astype(int)

    if is_bus:
        df.rename(columns={"value": "Rate"}, inplace=True)
    else:
        df["Rate 1"] = df["Rate 1"].fillna(0).astype(int)
        df["Rate"] = df["Rate 1"] + df["value"]

    # Resolve settlement amount column using case/space tolerant matching.
    settlement_candidates = [
        "Net Settlement Amt",
        "Net Settlement Amount",
        "SETTLED Amount (Rs.)",
        "Settled Amount (Rs.)",
        "settled amount",
    ]
    normalized_df_cols = {" ".join(str(col).strip().lower().split()): col for col in df.columns}
    settlement_col = None
    for candidate in settlement_candidates:
        normalized_candidate = " ".join(candidate.strip().lower().split())
        if normalized_candidate in normalized_df_cols:
            settlement_col = normalized_df_cols[normalized_candidate]
            break

    if not settlement_col:
        raise KeyError(
            "Settlement amount column not found. Expected one of: "
            f"{settlement_candidates}. Available columns: {list(df.columns)}"
        )

    if settlement_col != "Net Settlement Amt":
        df.rename(columns={settlement_col: "Net Settlement Amt"}, inplace=True)
    df["Net Settlement Amt"] = pd.to_numeric(df["Net Settlement Amt"], errors="coerce").fillna(0)
    df["Custom.4"] = np.where(df["Net Settlement Amt"] < df["Rate"], "Invalid", "Valid")

    return df


df_final_1 = get_rate(df_final_1, rates_df, is_bus=False)
df_bus = get_rate(df_bus, rates_df, is_bus=True)

df_combined = pd.concat([df_final_1, df_bus], ignore_index=True).drop(columns=["Rate 1", "value"], errors="ignore")

# -----------------------------
# Save valid/invalid output files
# -----------------------------
invalid_df = df_combined[df_combined["Custom.4"] == "Invalid"]
valid_df = df_combined[df_combined["Custom.4"] == "Valid"]

invalid_df.to_csv(output_invalid_path, index=False)
valid_df.to_csv(output_valid_path, index=False)
