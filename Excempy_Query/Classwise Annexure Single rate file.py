import pandas as pd
import os
import glob
import re
import warnings
import concurrent.futures
from openpyxl.styles import Font, Border, Side

# Ignore performance warnings
warnings.filterwarnings("ignore")

def safe_to_excel(df, path, **kwargs):
    """Saves DataFrame to excel, prompting the user if a PermissionError occurs."""
    while True:
        try:
            df.to_excel(path, **kwargs)
            break
        except PermissionError:
            input(f"\n[ERROR] Permission denied: Cannot write to '{path}'.\n"
                  f"Please CLOSE this file if it is open in Excel, then press Enter to retry...")

# ==========================================
# CENTRAL CONFIGURATION
# ==========================================
CONFIG = {
    # Step 1: Tool Processing Paths (Updated based on latest local settings)
    "SOURCE_DIR": r"F:\NHIT Q2 26-27\Chappar\august\Exempt file output",
    "EXEMPTION_FILE":  r"F:\NHIT Q2 26-27\Chappar\august\Concessioneir\Project Vehicle Details (1).xlsx",
    "RATES_FILE": r"F:\NHIT Q2 26-27\Rates\Bhojpuri.xlsx",
    "OUTPUT_FOLDER": r"F:\NHIT Q2 26-27\Chappar\august\Classwise annexure",
    "TARGET_SHEET": "Combined with RF 3",
    
    # Patterns & Filters
    "EXCLUDE_TC_PATTERN": r"AUTO|Tractor|Trailer",
    "GOV_VIP_PATTERN": r"GOV|GOT|VIP|CENTRAL|STATE",
    "OTHERS_PATTERN": r"Ambulance|Army|Police|Project Vehicle|Concessionaire|Fire Brigade|Defence|Funeral|Physically handicapped|Handicap",
    "EXCLUDE_VRNS": ["100", "108", "1033"],
    
    # Description Categories for General Reports
    "KEEP_CATEGORIES": [
        'Others', 'ETC Exempt', 'Non Tollable', 'Local Vehicle', 
        'Dispute', 'Dispute clear', 'Exempt Other', 'Force exemption', 
        'Local', 'CASH SINGLE', 'FORCED ION', 'CCH', 'NHAI EXEMPT', 'null', 
        'FASTAG', 'Cash', 'Exception', 'Exempt', 'Exemption'
    ],
    
    # Column Names
    "DATE_COLUMN": "Date & Time",
    "AMOUNT_COLUMN": "Applicable Rate",
    "VRN_COLUMN": "Veh Reg No."
}

# ==============================================================================
# STAGE 1: DATA PREPARATION & REPORT GENERATION (From unified_toll_pipeline.py)
# ==============================================================================

def load_data():
    """Reads all Excel files from the source directory and returns a concatenated DataFrame."""
    excel_files = glob.glob(os.path.join(CONFIG["SOURCE_DIR"], "*.xlsx"))
    if not excel_files:
        print(f"No Excel files found in {CONFIG['SOURCE_DIR']}")
        return None

    all_data = []
    print(f"Loading {len(excel_files)} files...")
    for file_path in excel_files:
        file_name = os.path.basename(file_path)
        try:
            df = pd.read_excel(file_path, sheet_name=CONFIG["TARGET_SHEET"])
            if not df.empty:
                df['Source_File'] = file_name
                all_data.append(df)
                print(f"  Successfully read: {file_name} ({len(df)} rows)")
        except Exception as e:
            print(f"  Error reading {file_name}: {e}")

    if not all_data:
        return None
    return pd.concat(all_data, ignore_index=True)

def clean_base_data(df):
    """Applies common cleaning steps across all processing flows."""
    print("Applying base cleaning...")
    
    # 1. TC Class Exclusion
    df = df[~df['TC Class'].str.contains(CONFIG["EXCLUDE_TC_PATTERN"], case=False, na=False)].copy()
    
    # 2. TC Class Standardization
    df['TC Class'] = df['TC Class'].astype(str).str.strip().str.upper()
    df['TC Class'] = df['TC Class'].str.replace(r'(TRUCK|TRK)\s*2\s*AXLE', 'TRK 2 AXLE', case=False, regex=True)
    df.loc[df['TC Class'] == 'TRUCK', 'TC Class'] = 'TRK 2 AXLE'
    
    # 3. Rename Custom to Custom.1
    if 'Custom' in df.columns:
        df.rename(columns={'Custom': 'Custom.1'}, inplace=True)
    
    # 4. Filter Specific VRNs
    if CONFIG["VRN_COLUMN"] in df.columns:
        df[CONFIG["VRN_COLUMN"]] = df[CONFIG["VRN_COLUMN"]].astype(str).str.strip().str.upper().str.replace(r'[\s\-\/_]', '', regex=True)
        df = df[~df[CONFIG["VRN_COLUMN"]].isin(CONFIG["EXCLUDE_VRNS"])]
        
    # 5. Remove Duplicate Transactions by Vehicle No + Date & Time
    if CONFIG["VRN_COLUMN"] in df.columns and CONFIG["DATE_COLUMN"] in df.columns:
        # Convert Date & Time to datetime for robust comparison
        df[CONFIG["DATE_COLUMN"]] = pd.to_datetime(df[CONFIG["DATE_COLUMN"]], errors='coerce')
        initial_len = len(df)
        df = df.drop_duplicates(subset=[CONFIG["VRN_COLUMN"], CONFIG["DATE_COLUMN"]]).copy()
        print(f"  Removed {initial_len - len(df)} duplicate transactions by Vehicle No + Date & Time.")

    return df

def apply_enrichment(df):
    """Applies Journey Type updates and Rates integration."""
    # 1. Updated Journey Type Logic
    if 'Journey Type' in df.columns:
        df['updated journey type'] = df['Journey Type']
    else:
        df['updated journey type'] = ""
        
    cond1 = (df['LNC/NLNC/LC/NLC'] == "Local Commercial") if 'LNC/NLNC/LC/NLC' in df.columns else pd.Series([False] * len(df))
    cond2 = (df['Date Validity Check LT'] == "Valid") if 'Date Validity Check LT' in df.columns else pd.Series([False] * len(df))
    
    df.loc[cond1 | cond2, 'updated journey type'] = "Local Conti/Single"
    df['updated journey type'] = df['updated journey type'].replace("Single Journey", "First Journey")

    # 2. Rates Integration
    try:
        rates_temp = pd.read_excel(CONFIG["RATES_FILE"], header=None)
        header_row_idx = -1
        for i, row in rates_temp.iterrows():
            if 'TC Class' in row.values:
                header_row_idx = i
                break
        
        if header_row_idx != -1:
            rates_df_raw = pd.read_excel(CONFIG["RATES_FILE"], header=header_row_idx)
            rates_df_raw = rates_df_raw.rename(columns={"Single Journey": "First Journey", "Local Conti/Local Single": "Local Conti/Single"})
            rates_melted = rates_df_raw.melt(id_vars=['TC Class', 'Vehicle Class', 'Weight/Capacity'], var_name='Attribute', value_name='Value')
            rates_melted['TC Class'] = rates_melted['TC Class'].astype(str).str.strip().str.upper()
            rates_melted['Attribute'] = rates_melted['Attribute'].astype(str).str.strip()
            
            rates_unique = rates_melted[['TC Class', 'Attribute', 'Value']].drop_duplicates(subset=['TC Class', 'Attribute'])
            df = df.merge(
                rates_unique,
                left_on=['TC Class', 'updated journey type'],
                right_on=['TC Class', 'Attribute'],
                how='left'
            )
            if 'Attribute' in df.columns:
                df.drop(columns=['Attribute'], inplace=True)
            if 'Value' in df.columns:
                df.rename(columns={'Value': CONFIG["AMOUNT_COLUMN"]}, inplace=True)
                df[CONFIG["AMOUNT_COLUMN"]] = df[CONFIG["AMOUNT_COLUMN"]].fillna(0).astype(int)
    except Exception as e:
        print(f"  Error integrating rates: {e}")
        
    return df

def apply_exemption_filter(df):
    """Filters data based on the external Concessionaire-Project Vehicles list."""
    try:
        exemption_df = pd.read_excel(CONFIG["EXEMPTION_FILE"])
        vrn_col = next((col for col in exemption_df.columns if col.upper() == 'VRN'), None)
        if vrn_col and CONFIG["VRN_COLUMN"] in df.columns:
            exemption_df[vrn_col] = exemption_df[vrn_col].astype(str).str.strip().str.upper().str.replace(r'[\s\-\/_]', '', regex=True)
            exempted_vrns = set(exemption_df[vrn_col].unique())
            df = df[~df[CONFIG["VRN_COLUMN"]].isin(exempted_vrns)]
    except Exception as e:
        print(f"  Error in exemption filter: {e}")
    return df

def generate_reports(base_df):
    """Generates all categorized reports in the output folder."""
    print("\nStarting Report Generation...")
    
    # PRE-CLEANING TRI COLUMNS
    for col in ['TRI 3', 'TRI 2']:
        if col in base_df.columns:
            base_df[col] = base_df[col].astype(str).str.replace(r'\.0$', '', regex=True)
    
    base_df['Merged_TRI'] = base_df[['TRI 3', 'TRI 2']].apply(
        lambda x: '|'.join([str(val) for val in x if str(val) not in ['nan', 'None', '']]), axis=1
    ).str.strip('|').fillna('')

    # --- GROUP 1: EXEMPT REPORTS ---
    exempt_configs = [
        {"name": "Gov/VIP", "pattern": CONFIG["GOV_VIP_PATTERN"], "file": "govt etc + exmpt.xlsx", "logic": "standard"},
        {"name": "Others", "pattern": CONFIG["OTHERS_PATTERN"], "file": "multiple cat others + exmpt.xlsx", "logic": "standard"},
        {"name": "Multiple Class", "pattern": f"{CONFIG['GOV_VIP_PATTERN']}|{CONFIG['OTHERS_PATTERN']}", "file": "multipe class.xlsx", "logic": "both_zero"}
    ]

    for config in exempt_configs:
        temp_df = base_df[base_df['Description'].str.contains(config['pattern'], case=False, na=False)].copy()
        if config['logic'] == "standard":
            temp_df = temp_df[(temp_df['TRI 3'] != "0") & (temp_df['Merged_TRI'] != "0")]
        else:
            temp_df = temp_df[(temp_df['TRI 2'] == "0") & (temp_df['TRI 3'] == "0")]
        
        final_df = temp_df[(temp_df['Date Validity Check MP'] != "Valid") & (temp_df['Custom.1'] == "Exception")].copy()
        if not final_df.empty:
            final_df = apply_exemption_filter(final_df)
            final_df = apply_enrichment(final_df)
            if CONFIG["VRN_COLUMN"] in final_df.columns and CONFIG["DATE_COLUMN"] in final_df.columns:
                final_df[CONFIG["DATE_COLUMN"]] = pd.to_datetime(final_df[CONFIG["DATE_COLUMN"]], errors='coerce')
                initial_len = len(final_df)
                final_df = final_df.drop_duplicates(subset=[CONFIG["VRN_COLUMN"], CONFIG["DATE_COLUMN"]])
                print(f"  Removed {initial_len - len(final_df)} duplicate rows from {config['file']}.")
            output_path = os.path.join(CONFIG["OUTPUT_FOLDER"], config['file'])
            safe_to_excel(final_df, output_path, index=False)
            print(f"  Generated: {config['file']} ({len(final_df)} rows)")

    # --- GROUP 2: CATEGORY REPORTS ---
    desc_pattern = '|'.join(CONFIG["KEEP_CATEGORIES"])
    cat_df = base_df[
        base_df['Description'].str.contains(desc_pattern, case=False, na=True) | 
        (base_df['Description'].astype(str).str.strip() == '') |
        (base_df['Description'].astype(str).str.lower() == 'null')
    ].copy()

    cat_df['Custom_Helper'] = cat_df['Custom.1'].apply(lambda x: x if str(x).strip() == 'Exception' else None)
    mask_custom2 = (cat_df['Custom.1'].isin(['Non Exception 1', 'Non Exception 2'])) & \
                   (cat_df[CONFIG["VRN_COLUMN"]].astype(str).str.startswith('NEW', na=False))
    cat_df['Custom.2'] = None
    cat_df.loc[mask_custom2, 'Custom.2'] = cat_df['Custom.1']
    cat_df['Merged_Custom'] = cat_df['Custom_Helper'].fillna('').astype(str) + cat_df['Custom.2'].fillna('').astype(str)
    
    cat_df = cat_df[cat_df['Merged_Custom'].str.strip() != '']
    cat_df = cat_df[(cat_df['Date Validity Check MP'] != "Valid") & (cat_df['Custom.1'] == "Exception")]
    
    if not cat_df.empty:
        cat_df = apply_exemption_filter(cat_df)
        cat_df = apply_enrichment(cat_df)
        cat_df = cat_df.drop_duplicates()

        cat_configs = [
            {"file": "LNC FT.xlsx", "cond": lambda d: (d['LNC/NLNC/LC/NLC'] == "Local Non Commercial") & (d['Trip Group'] == ">5")},
            {"file": "LNC CT.xlsx", "cond": lambda d: (d['LNC/NLNC/LC/NLC'] == "Local Non Commercial") & (d['Trip Group'] == "0-5")},
            {"file": "NLNC.xlsx", "cond": lambda d: (d['LNC/NLNC/LC/NLC'] == "Non Local Non Commercial")},
            {"file": "commercial.xlsx", "cond": lambda d: (d['LNC/NLNC/LC/NLC'].isin(["Local Commercial", "Non Local Commercial"]))}
        ]

        for config in cat_configs:
            subset = cat_df[config['cond'](cat_df)].copy()
            if not subset.empty:
                cols_to_drop = ['Custom_Helper', 'Custom.2', 'Merged_Custom', 'updated journey type', 'Merged_TRI']
                subset.drop(columns=[c for c in cols_to_drop if c in subset.columns], inplace=True)
                if CONFIG["VRN_COLUMN"] in subset.columns and CONFIG["DATE_COLUMN"] in subset.columns:
                    subset[CONFIG["DATE_COLUMN"]] = pd.to_datetime(subset[CONFIG["DATE_COLUMN"]], errors='coerce')
                    initial_len = len(subset)
                    subset = subset.drop_duplicates(subset=[CONFIG["VRN_COLUMN"], CONFIG["DATE_COLUMN"]])
                    print(f"  Removed {initial_len - len(subset)} duplicate rows from {config['file']}.")
                output_path = os.path.join(CONFIG["OUTPUT_FOLDER"], config['file'])
                safe_to_excel(subset, output_path, index=False)
                print(f"  Generated: {config['file']} ({len(subset)} rows)")

# ==============================================================================
# STAGE 2: ANNEXURE SUMaprY & FORMATTING (From process_all_annexures.py)
# ==============================================================================

def add_sumapry_and_formatting(file_path):
    """Processes a single file to add month-based sheets and a formatted sumapry."""
    filename = os.path.basename(file_path)
    is_lnc_ft = (filename == "LNC FT.xlsx")
    
    try:
        # 1. Load data
        df = pd.read_excel(file_path, parse_dates=[CONFIG["DATE_COLUMN"]])
        df = df.dropna(subset=[CONFIG["DATE_COLUMN"]])
        df["Month"] = df[CONFIG["DATE_COLUMN"]].dt.strftime("%b-%y")

        # 2. Aggregation
        if is_lnc_ft:
            sumapry_df = df.groupby("Month").agg(
                Total_Amount=(CONFIG["AMOUNT_COLUMN"], "sum"),
                Unique_Transaction_Count=(CONFIG["VRN_COLUMN"], "nunique")
            ).reset_index()
             # Multiply Total Amount by 360 for LNC FT file
            sumapry_df["Total_Amount"] = sumapry_df["Unique_Transaction_Count"] * 360   
        
            count_col = "Unique_Transaction_Count"
        else:
            sumapry_df = df.groupby("Month").agg(
                Total_Amount=(CONFIG["AMOUNT_COLUMN"], "sum"),
                Transaction_Count=(CONFIG["AMOUNT_COLUMN"], "count")
            ).reset_index()
            count_col = "Transaction_Count"

        # 3. Sort Chronologically
        sumapry_df["SortMonth"] = pd.to_datetime(sumapry_df["Month"], format="%b-%y")
        sumapry_df = sumapry_df.sort_values("SortMonth").drop(columns="SortMonth")

        # 4. Add Total Row
        total_row = pd.DataFrame([{
            "Month": "Total",
            "Total_Amount": sumapry_df["Total_Amount"].sum(),
            count_col: sumapry_df[count_col].sum()
        }])
        sumapry_df = pd.concat([sumapry_df, total_row], ignore_index=True)

        # 5. Write Sheets
        while True:
            try:
                with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    for month, group in df.groupby("Month"):
                        group.drop(columns="Month").to_excel(writer, sheet_name=month, index=False)
                    sumapry_df.to_excel(writer, sheet_name="Sumapry", index=False)

                    # Formatting
                    ws = writer.sheets["Sumapry"]
                    bold_font = Font(bold=True)
                    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

                    for row in ws.iter_rows():
                        for cell in row:
                            cell.border = thin_border
                            if cell.row == 1 or cell.value == "Total":
                                cell.font = bold_font
                break
            except PermissionError:
                input(f"\n[ERROR] Permission denied: Cannot write to '{file_path}'.\n"
                      f"Please CLOSE this file if it is open in Excel, then press Enter to retry...")
                        
        print(f"  Formatted sumapry added for: {filename}")
    except Exception as e:
        print(f"  Error processing sumapry for {filename}: {e}")

def run_annexure_sumapries():
    """Finds all generated files and processes them in parallel."""
    print("\nStarting Sumapry Generation and Formatting...")
    xlsx_files = glob.glob(os.path.join(CONFIG["OUTPUT_FOLDER"], "*.xlsx"))
    valid_files = [f for f in xlsx_files if not os.path.basename(f).startswith("~$")]
    
    if not valid_files:
        print("  No files found for formatting.")
        return

    print(f"  Formatting {len(valid_files)} files in parallel...")
    with concurrent.futures.ThreadPoolExecutor() as executor:
        executor.map(add_sumapry_and_formatting, valid_files)

# ==========================================
# MAIN WORKFLOW
# ==========================================

def main():
    print("=== FINAL TOLL & ANNEXURE PIPELINE STARTED ===\n")
    
    # Check/Create output dir
    os.makedirs(CONFIG["OUTPUT_FOLDER"], exist_ok=True)
    
    # STAGE 1: Data Preparation
    raw_df = load_data()
    if raw_df is None:
        print("Pipeline aborted: No data loaded.")
        return
        
    base_df = clean_base_data(raw_df)
    generate_reports(base_df)
    
    # STAGE 2: Sumapry and Formatting
    run_annexure_sumapries()
    
    print("\n=== PIPELINE COMPLETED SUCCESSFULLY ===")

if __name__ == "__main__":
    main()
