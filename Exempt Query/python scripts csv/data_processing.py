import pandas as pd

SIDE_1_LANES = ['L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'L7', 'L8']
SIDE_2_LANES = ['L9', 'L10', 'L11', 'L12', 'L13', 'L14', 'L15', 'L16']


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

    # Lane to Side Mapping
    def map_lane_to_side(lane):
        if lane in SIDE_1_LANES:
            return 'Side 1'
        elif lane in SIDE_2_LANES:
            return 'Side 2'
        else:
            return 'Unknown Side'

    df['Side'] = df['Lane No'].apply(map_lane_to_side)

    return df, exempt_vehicle_list