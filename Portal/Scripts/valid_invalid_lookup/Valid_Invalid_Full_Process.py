import os
from pathlib import Path
import json
import sys

import numpy as np
import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine

PORTAL_ROOT = Path(__file__).resolve().parents[2]
REPO_ROOT = Path(__file__).resolve().parents[3]
if str(PORTAL_ROOT) not in sys.path:
    sys.path.insert(0, str(PORTAL_ROOT))

load_dotenv(REPO_ROOT / ".env")

from Header_Mapping.header_mapping import VALID_INVALID_LOOKUP_HEADER_MAPPING
from header_matching import normalize_header_match

from valid_invalid_config import (
    ACCEPTED_STATUS_VALUES,
    BUS_VEHICLE_CLASSES,
    CRANE_MOUNTED_VEHICLE_CLASS,
    DISCOUNT_PASS_JOURNEY_TYPES,
    HEAVY_SPECIAL_VEHICLE_CLASSES,
    INVALID_EXCLUDED_VEHICLE_CLASSES,
    JOURNEY_TYPES_TO_DROP,
    LIFECYCLE_EXCLUDED_VEHICLE_CLASSES,
    RATE_SHEET_JOURNEY_COLUMN_RENAMES,
    STATUS_COLUMN_CANDIDATES,
    UPDATED_JOURNEY_TYPE_CONT,
    UPDATED_JOURNEY_TYPE_LOCAL,
    UPDATED_JOURNEY_TYPE_SINGLE,
    ZERO_SETTLEMENT_REMOVE_JOURNEY_TYPES,
    map_npci_to_tc_class,
    normalize_journey_type,
)


def normalize_vehicle_class(value) -> str:
    """Strip + casefold so ' BUS' / 'BUS' / 'bus' compare equal."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip().casefold()


def normalized_vehicle_class_set(values) -> set:
    return {normalize_vehicle_class(v) for v in values if normalize_vehicle_class(v)}


def vehicle_class_isin(series: pd.Series, values) -> pd.Series:
    targets = normalized_vehicle_class_set(values)
    return series.map(normalize_vehicle_class).isin(targets)


def vehicle_class_eq(series: pd.Series, value) -> pd.Series:
    target = normalize_vehicle_class(value)
    if not target:
        return pd.Series(False, index=series.index)
    return series.map(normalize_vehicle_class) == target


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
RATES_INPUT_PATH = os.environ.get("VALID_INVALID_RATES_PATH")
LIFECYCLE_INPUT_PATH = os.environ.get("VALID_INVALID_LIFECYCLE_PATH", r"merged_output.csv")
USER_HEADER_MAPPING = json.loads(os.environ.get("VALID_INVALID_HEADER_MAPPING", "{}"))


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

    def normalize_column_name(name):
        return normalize_header_match(name)

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

    df_c["Unique Vehicle Number"] = df_c["Unique Vehicle Number"].astype(str).str.replace(".", "", regex=False)
    df_c["weight"] = df_c["weight"].apply(lambda x: str(x) if pd.notnull(x) else None)

    df_c = df_c[
        ~df_c["vehicle class"].isin(
            [
                "---Select Vehicle Class---",
                "THREE WHEE! ER{PASSENGER)",
                "THREE WHEELER(GOODS)",
                "THREE WHEELER(PASSENGER)",
                "THREE WHEELER{GOODS)",
                "—Select Vehicle Class—",
                "--Select Vehicle Class--",
                "----Select Vehicle Class----",
                "---Select Vehicle Class---",
                "----Select Vehicle Class----",
            ]
        )
        & (df_c["Unique Vehicle Number"] != "PB11U4131")
        & (df_c["Unique Vehicle Number"] != "MH04DT4250")
        & (df_c["Unique Vehicle Number"] != "MP13WA9975")
        & (df_c["Unique Vehicle Number"] != "MP13WA9928")
    ]

    df_c = df_c[~((df_c["vehicle class"].isnull()) & (df_c["weight"].isnull()))]
    df_c = df_c.drop_duplicates()
    df_c = df_c.drop_duplicates(subset="Unique Vehicle Number")

    df_l.columns = df_l.columns.str.strip()

    rename_map = {
        canonical: [candidate for candidate in candidates if candidate != canonical]
        for canonical, candidates in VALID_INVALID_LOOKUP_HEADER_MAPPING.items()
    }

    user_renames = {}
    for target, selected_source in USER_HEADER_MAPPING.items():
        if (
            target in VALID_INVALID_LOOKUP_HEADER_MAPPING
            and selected_source
            and selected_source in df_l.columns
            and selected_source != target
        ):
            user_renames[selected_source] = target

    if user_renames:
        df_l = df_l.rename(columns=user_renames)

    normalized_columns = {normalize_header_match(col): col for col in df_l.columns}

    for target, alternatives in rename_map.items():
        if target not in df_l.columns:
            for alt in alternatives:
                normalized_alt = normalize_header_match(alt)
                if normalized_alt in normalized_columns:
                    df_l = df_l.rename(columns={normalized_columns[normalized_alt]: target})
                    break

    if "Journey Type" in df_l.columns:
        df_l = df_l[~df_l["Journey Type"].isin(JOURNEY_TYPES_TO_DROP)]

    status_col = next(
        (col for col in STATUS_COLUMN_CANDIDATES if col in df_l.columns),
        None,
    )
    if status_col:
        df_l = df_l[df_l[status_col].isin(ACCEPTED_STATUS_VALUES)]

    df_l["Veh Reg No."] = df_l["Veh Reg No."].str.upper().str.strip()
    df_c["Unique Vehicle Number"] = df_c["Unique Vehicle Number"].str.upper().str.strip()

    df_merged = df_l.merge(
        df_c,
        left_on="Veh Reg No.",
        right_on="Unique Vehicle Number",
        how="left",
    ).drop(columns=["Unique Vehicle Number"])

    df_merged["updated journey type"] = df_merged["Journey Type"].apply(normalize_journey_type)

    df_final = df_merged[
        (df_merged["Veh Reg No."].str.len() <= 12)
        & ((df_merged["vehicle class"].notna()) | (df_merged["weight"].notna() & (df_merged["weight"] != "nan")))
    ]

    df_final["weight"] = pd.to_numeric(df_final["weight"], errors="coerce").astype("Int64")
    return df_final


def rates(df):
    df = df.rename(columns=RATE_SHEET_JOURNEY_COLUMN_RENAMES)
    return df.melt(id_vars=["TC Class", "Weight/Capacity", "Vehicle Class"], var_name="Attribute")


def get_rate_from_rates(tc_class, journey_type, rates_df):
    row = rates_df.loc[
        (rates_df["TC Class"] == tc_class) & (rates_df["Attribute"] == journey_type),
        "value",
    ]
    return row.values[0] if not row.empty else 0


def vehicle_lifecycle(df_result, rates_df):
    df_result = df_result[~vehicle_class_isin(df_result["vehicle class"], LIFECYCLE_EXCLUDED_VEHICLE_CLASSES)]

    conditions = [
        vehicle_class_isin(df_result["vehicle class"], HEAVY_SPECIAL_VEHICLE_CLASSES)
        & (df_result["updated journey type"] == UPDATED_JOURNEY_TYPE_SINGLE),
        vehicle_class_isin(df_result["vehicle class"], HEAVY_SPECIAL_VEHICLE_CLASSES)
        & (df_result["updated journey type"] == UPDATED_JOURNEY_TYPE_CONT),
        vehicle_class_isin(df_result["vehicle class"], HEAVY_SPECIAL_VEHICLE_CLASSES)
        & (df_result["updated journey type"] == UPDATED_JOURNEY_TYPE_LOCAL),
        vehicle_class_eq(df_result["vehicle class"], CRANE_MOUNTED_VEHICLE_CLASS)
        & (df_result["updated journey type"] == UPDATED_JOURNEY_TYPE_SINGLE),
        vehicle_class_eq(df_result["vehicle class"], CRANE_MOUNTED_VEHICLE_CLASS)
        & (df_result["updated journey type"] == UPDATED_JOURNEY_TYPE_CONT),
        vehicle_class_eq(df_result["vehicle class"], CRANE_MOUNTED_VEHICLE_CLASS)
        & (df_result["updated journey type"] == UPDATED_JOURNEY_TYPE_LOCAL),
    ]

    values = [
        get_rate_from_rates("MAV", UPDATED_JOURNEY_TYPE_SINGLE, rates_df),
        get_rate_from_rates("MAV", UPDATED_JOURNEY_TYPE_CONT, rates_df),
        get_rate_from_rates("MAV", UPDATED_JOURNEY_TYPE_LOCAL, rates_df),
        get_rate_from_rates("LCV", UPDATED_JOURNEY_TYPE_SINGLE, rates_df),
        get_rate_from_rates("LCV", UPDATED_JOURNEY_TYPE_CONT, rates_df),
        get_rate_from_rates("LCV", UPDATED_JOURNEY_TYPE_LOCAL, rates_df),
    ]

    df_result["Rate 1"] = np.select(conditions, values, default=None)

    def classify_weight(weight):
        if pd.isna(weight):
            return None
        if weight <= 7500:
            return "<=7500 Kgs"
        if weight <= 12000:
            return ">7500 Kgs but <=12,000 Kgs"
        if weight <= 18500:
            return "> 12,000 Kgs but <= 18,500 Kgs"
        if weight <= 28000:
            return "> 18,500 Kgs but <= 28,000 Kgs"
        if weight <= 60000:
            return "> 28,000 Kgs but <= 60,000 Kgs"
        return ">60,000 Kgs"

    df_result["Custom.1"] = df_result["weight"].apply(classify_weight)
    return df_result


def vehicle_bus(df_result):
    df_result = df_result[vehicle_class_isin(df_result["vehicle class"], BUS_VEHICLE_CLASSES)]
    df_result = df_result.dropna(subset=["weight"])

    def classify_weight(row):
        weight = row["weight"]
        if weight > 32:
            return "More than 32 seating capacity"
        if 12 < weight <= 32:
            return "exceeds 12 but less than equal to 32"
        return "Less than equal to 12 seating capacity"

    df_result["Custom.1"] = df_result.apply(classify_weight, axis=1)
    return df_result


def get_rate(df, rates_df, is_bus=False):
    df = df.merge(
        rates_df[["Attribute", "Weight/Capacity", "value"]],
        left_on=["updated journey type", "Custom.1"],
        right_on=["Attribute", "Weight/Capacity"],
        how="left",
    ).drop(columns=["Attribute", "Weight/Capacity"])

    df["value"] = df["value"].fillna(0).astype(int)

    if is_bus:
        df.rename(columns={"value": "Rate"}, inplace=True)
    else:
        df["Rate 1"] = df["Rate 1"].fillna(0).astype(int)
        df["Rate"] = df["Rate 1"] + df["value"]

    settlement_candidates = [
        "Net Settlement Amt",
        "Net Settlement Amount",
        "SETTLED Amount (Rs.)",
        "Settled Amount (Rs.)",
        "settled amount",
    ]
    normalized_df_cols = {normalize_header_match(col): col for col in df.columns}
    settlement_col = None
    for candidate in settlement_candidates:
        normalized_candidate = normalize_header_match(candidate)
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


def map_npci_to_tc(npci):
    return map_npci_to_tc_class(npci)

def process_vehicle_classification(df):
    df = df.copy()
    df["weight"] = df["weight"].replace(0, None)

    def correct_vehicle_class(weight):
        if pd.isna(weight) or weight == "":
            return "MAV"
        if weight <= 7500:
            return "LMV"
        if weight <= 12000:
            return "LCV"
        if weight <= 18500:
            return "Truck 2 axle"
        if weight <= 28000:
            return "Truck 3 axle"
        if 28000 < weight <= 60000:
            return "MAV"
        return "Truck 7 axle"

    df["Correct Vehicle Class"] = df["weight"].apply(correct_vehicle_class)

    def classify_custom_2(row):
        if row["Custom.1"] == "Less than equal to 12 seating capacity":
            return "LMV"
        if row["Custom.1"] == "More than 32 seating capacity":
            return "BUS"
        if row["Custom.1"] == "exceeds 12 but less than equal to 32":
            return "LCV"
        return row["Correct Vehicle Class"]

    df["Custom.2"] = df.apply(classify_custom_2, axis=1)
    df.drop(columns=["Correct Vehicle Class"], inplace=True)
    df.rename(columns={"Custom.2": "Correct Vehicle Class"}, inplace=True)
    return df


def main():
    print("Valid/Invalid Lookup: loading input data and rates.")
    rates_source = pd.read_excel(RATES_INPUT_PATH, skiprows=1)
    rates_df = rates(rates_source)

    df_result = process_vehicle(LIFECYCLE_INPUT_PATH)
    df_final_1 = vehicle_lifecycle(df_result, rates_df)
    df_bus = vehicle_bus(df_result)

    df_final_1 = get_rate(df_final_1, rates_df, is_bus=False)
    df_bus = get_rate(df_bus, rates_df, is_bus=True)
    df_combined = pd.concat([df_final_1, df_bus], ignore_index=True).drop(columns=["Rate 1", "value"], errors="ignore")

    single_rates = rates_df[rates_df["Attribute"] == "Single Journey"]
    rate_lookup = dict(zip(single_rates["TC Class"], single_rates["value"]))

    df_invalid = df_combined[df_combined["Custom.4"] == "Invalid"]
    df_invalid = process_vehicle_classification(df_invalid)

    df_valid = df_combined[df_combined["Custom.4"] == "Valid"]
    df_valid = process_vehicle_classification(df_valid)
    print(f"Valid/Invalid Lookup: split data -> invalid={len(df_invalid)}, valid={len(df_valid)}")

    df_invalid["Custom.2"] = df_invalid.apply(
        lambda row: "Remove"
        if row["Net Settlement Amt"] == 0
        and row["Journey Type"] in ZERO_SETTLEMENT_REMOVE_JOURNEY_TYPES
        else "Keep",
        axis=1,
    )
    df_invalid = df_invalid[df_invalid["Custom.2"] != "Remove"]
    df_invalid.drop(columns=["Custom.2"], inplace=True)

    def update_net_settlement(row):
        if row["Journey Type"] in DISCOUNT_PASS_JOURNEY_TYPES:
            tc_class = map_npci_to_tc(row["NPCI Class Desc"])
            if tc_class and tc_class in rate_lookup:
                return rate_lookup[tc_class]
        return row["Net Settlement Amt"]

    df_invalid["updated net settlement amount"] = df_invalid.apply(update_net_settlement, axis=1)
    df_invalid.drop_duplicates(inplace=True)
    df_invalid["Impact"] = df_invalid["Rate"] - df_invalid["updated net settlement amount"]
    df_invalid = df_invalid[(df_invalid["Impact"] != 0) & (df_invalid["Impact"] > 0)]

    df_invalid = df_invalid[
        ~vehicle_class_isin(df_invalid["vehicle class"], INVALID_EXCLUDED_VEHICLE_CLASSES)
    ]

    df_valid = df_valid[df_valid["Rate"] != 0]
    print(f"Valid/Invalid Lookup: final invalid rows={len(df_invalid)}, final valid rows={len(df_valid)}")

    df_invalid.to_csv(output_invalid_path, index=False)
    df_valid.to_csv(output_valid_path, index=False)
    print(f"Valid/Invalid Lookup: exported invalid -> {output_invalid_path}")
    print(f"Valid/Invalid Lookup: exported valid -> {output_valid_path}")


if __name__ == "__main__":
    main()
