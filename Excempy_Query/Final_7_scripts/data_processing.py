import re

import pandas as pd

# Populated from distinct "Lane No" values in the semi-final file (first half / second half).
SIDE_1_LANES = []
SIDE_2_LANES = []


def _lane_sort_key(lane):
    """Sort L1, L2, ... L10 numerically (not lexicographically)."""
    text = str(lane).strip().upper()
    match = re.search(r"(\d+)", text)
    num = int(match.group(1)) if match else -1
    return (num, text)


def _normalize_lane_label(lane) -> str:
    if lane is None:
        return ""
    try:
        if pd.isna(lane):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(lane).strip()
    if not text or text.lower() == "nan":
        return ""
    return text


def allocate_side_lanes(lane_values):
    """
    Build Side 1 / Side 2 lane lists from distinct Lane No values.

    Distinct lanes are sorted numerically (L1..Ln). First half → Side 1,
    second half → Side 2. If the count is odd, append the next lane label
    (e.g. L15 → add L16) so the list is even before splitting.
    """
    global SIDE_1_LANES, SIDE_2_LANES

    distinct = []
    seen = set()
    for raw in lane_values:
        label = _normalize_lane_label(raw)
        if not label:
            continue
        key = label.casefold()
        if key in seen:
            continue
        seen.add(key)
        distinct.append(label)

    lanes = sorted(distinct, key=_lane_sort_key)

    if not lanes:
        SIDE_1_LANES = []
        SIDE_2_LANES = []
        print("[LANES] No distinct Lane No values found; Side mapping will be Unknown.")
        return SIDE_1_LANES, SIDE_2_LANES

    if len(lanes) % 2 == 1:
        last_num = _lane_sort_key(lanes[-1])[0]
        pad_lane = f"L{last_num + 1}" if last_num >= 0 else f"{lanes[-1]}_PAD"
        lanes.append(pad_lane)
        print(
            f"[LANES] Odd distinct count; padded with {pad_lane} to make even "
            f"({len(lanes)} lanes)."
        )

    mid = len(lanes) // 2
    SIDE_1_LANES = lanes[:mid]
    SIDE_2_LANES = lanes[mid:]
    print(
        f"[LANES] Side 1 ({len(SIDE_1_LANES)}): {SIDE_1_LANES}\n"
        f"[LANES] Side 2 ({len(SIDE_2_LANES)}): {SIDE_2_LANES}"
    )
    return SIDE_1_LANES, SIDE_2_LANES


def map_lane_to_side(lane, side_1_lanes=None, side_2_lanes=None):
    side_1 = side_1_lanes if side_1_lanes is not None else SIDE_1_LANES
    side_2 = side_2_lanes if side_2_lanes is not None else SIDE_2_LANES
    label = _normalize_lane_label(lane)
    if not label:
        return "Unknown Side"
    side_1_keys = {str(x).strip().casefold() for x in side_1}
    side_2_keys = {str(x).strip().casefold() for x in side_2}
    key = label.casefold()
    if key in side_1_keys:
        return "Side 1"
    if key in side_2_keys:
        return "Side 2"
    return "Unknown Side"


def load_and_preprocess_data(file_path):
    # Load CSV file directly
    df = pd.read_csv(file_path)

    # Strip extra spaces from column headers
    df.columns = df.columns.str.strip()

    # Define column name mapping
    column_mapping = {
        # 'TC_VEH_REG_NO': 'Veh Reg No.',
        # 'MVC_TLC_CLASS': 'TC Class',
        # 'MOP_DESCRIPTION': 'Description',
        # 'LANE_NO': 'Lane No',
        # 'TXN_DATE' : 'Date & Time',
        # 'PAYMENT_TYPE' : 'MOP'
    }

    # Rename columns
    df.rename(columns=column_mapping, inplace=True)

    # Ensure 'Veh Reg No.' is string and remove blank/nan-like entries
    df['Veh Reg No.'] = df['Veh Reg No.'].astype(str)
    #df['Txn ID'] = df['Txn ID'].astype(str)
    df = df[df['Veh Reg No.'].str.strip().str.lower() != 'nan']

    # Filter for MOP not in specific unwanted values
    #df = df[(df['MOP'] != "Violation")]
    df = df[(df['MOP'] != "VIOLATION") & (df['MOP'] != "FLEET") & (df['MOP'] != "CONVEY") & (df['MOP'] != "RUNTHROUGH") & (df['MOP'] != "RUNTHROUGHRUNTHROUGH")]

    #df = df[(df['MOP'] != "RUNTHROUGHRUNTHROUGH") & (df['MOP'] != "RUNTHROUGH") & (df['MOP'] != "VIOLATION") & (df['MOP'] != "CONVEY")]

    # Keep only allowed MOP values
    #df = df[df['MOP'].isin(["ETC", "Cash", "Exempt"])]
    df = df[df['MOP'].isin(["BARCODE", "CASH", "CONVEY", "ETC", "FASTAG", "EXEMPT", "RFID TAG", "TAG", "DEMAND DRAFT", "CCH", "E-WALLET", "UPI", "CC DC", "CARD PAYMENT", "TAGCASH", "PENALTY", "EWALLET"])]
    #df = df[df['MOP'].isin(["CASH", "CASHSINGLE", "ETC", "TAG", "EXEMPT", "TAGFAST-TAG", "RFID TAG", "TAGFAST-TAG"])]

    # Derive MOP Final
    df['MOP Final'] = df['MOP'].apply(lambda x: 'ETC' if x in ["BARCODE", "CASH", "CONVEY", "ETC", "FASTAG", "RFID TAG", "TAG", "DEMAND DRAFT",  "CCH", "E-WALLET", "UPI", "CC DC", "CARD PAYMENT", "TAGCASH", "PENALTY", "EWALLET"] else 'EXEMPT')
    #df['MOP Final'] = df['MOP'].apply(lambda x: 'ETC' if x in ["CASH", "CASHSINGLE", "ETC", "TAG", "TAGFAST-TAG"] else 'EXEMPT')

    # Parse Date & Time
    df['Date & Time'] = pd.to_datetime(df['Date & Time'], errors='coerce')
    # Check parsed 'Date & Time'
    print("\nSample of 'Date & Time' values:")
    print(df['Date & Time'].head(10))

    print("\nNumber of rows with invalid 'Date & Time' (NaT):", df['Date & Time'].isna().sum())

    #df['Date & Time'] = pd.to_datetime(df['Date & Time'], format='%d-%m-%Y %H:%M:%S', errors='coerce')

    # Extract Exempt vehicle list
    exempt_vehicle_list = df[df['MOP Final'] == 'EXEMPT'][['Veh Reg No.']].drop_duplicates()

    # Lane to Side Mapping — derived from distinct Lane No in this semi-final file
    allocate_side_lanes(df["Lane No"])
    df['Side'] = df['Lane No'].apply(map_lane_to_side)

    return df, exempt_vehicle_list
