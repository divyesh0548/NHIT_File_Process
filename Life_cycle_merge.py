import pandas as pd
import os
from multiprocessing import Pool, cpu_count
from pathlib import Path
import shutil
import tempfile
from time import perf_counter

def _normalize_header_value(value):
    if pd.isna(value):
        return ""
    return str(value).strip().lower()


def _detect_header_row_index(df_raw, header_keywords, min_matches=3):
    """
    Detect header row index by scanning row values and matching header keywords.
    Returns 0 when no confident match is found.
    """
    target = {_normalize_header_value(k) for k in header_keywords}
    best_idx = 0
    best_matches = 0

    for idx in range(len(df_raw)):
        row_values = {_normalize_header_value(v) for v in df_raw.iloc[idx].tolist()}
        match_count = len(target.intersection(row_values))
        if match_count > best_matches:
            best_matches = match_count
            best_idx = idx
        if match_count >= min_matches:
            return idx

    return best_idx if best_matches >= min_matches else 0


def find_header_csv(df, header_keywords, file_name):
    """
    Finds the header row in a CSV DataFrame based on the presence of given header keywords.
    At least 3 keywords must be detected in the columns.

    Arguments:
    df : DataFrame : The loaded DataFrame of the CSV file
    header_keywords : list : List of header keywords to search for
    file_name : str : The name of the file being processed
    
    Returns:
    DataFrame : The DataFrame starting from the header row
    """
    # Count how many of the header keywords are found in the columns
    matched_keywords = [keyword for keyword in header_keywords if keyword in df.columns]
    
    # If at least 3 keywords are found, find the starting index and return the subset
    if len(matched_keywords) >= 3:
        # Get the index of the first matched keyword
        start_idx = df.columns.get_loc(matched_keywords[0])
        # Print the file and header row details (row index)
        print(f"Header detected in {file_name} at row index {start_idx} with keywords {matched_keywords}")
        # Return the DataFrame starting from that index
        df = df.iloc[:, start_idx:]
        
    return df

def find_header_excel(excel_data, sheet_name, header_keywords, file_name):
    """
    Finds the header row in an Excel sheet based on the presence of given header keywords.
    At least 3 keywords must be detected in the header row (not column).

    Arguments:
    excel_data : pandas.ExcelFile : Loaded Excel workbook
    sheet_name : str : The sheet name
    header_keywords : list : List of header keywords to search for
    
    Returns:
    DataFrame : The DataFrame starting from the header row
    """
    # Read sheet without assuming header position, detect header row first.
    df_raw = pd.read_excel(excel_data, sheet_name=sheet_name, header=None, dtype=str)
    if df_raw.empty:
        return pd.DataFrame()

    header_idx = _detect_header_row_index(df_raw, header_keywords)
    print(f"Header detected in {file_name} [{sheet_name}] at row index {header_idx}")

    # Re-read using the detected header row so columns are correctly parsed.
    df = pd.read_excel(excel_data, sheet_name=sheet_name, skiprows=header_idx, header=0)
    
    return df

def process_file(file, header_keywords, columns_in_first_file):
    """
    Processes a single file and returns the processed data as a DataFrame.
    Ensures that the column order matches the first file and new columns are added.

    Arguments:
    file : str : The file path
    header_keywords : list : List of header keywords to detect the header row
    columns_in_first_file : list : Columns from the first file to align subsequent files with
    
    Returns:
    DataFrame : Merged DataFrame of the processed file
    """
    if file.endswith('.csv'):
        # For CSV files, try multi-threaded parser first (when available).
        try:
            df = pd.read_csv(file, engine="pyarrow")
        except Exception:
            df = pd.read_csv(file, low_memory=False)
        # Apply header detection logic
        df = find_header_csv(df, header_keywords, file)
    elif file.endswith('.xls') or file.endswith('.xlsx'):
        excel_data = pd.ExcelFile(file)
        merged_data = []
        for sheet_name in excel_data.sheet_names:
            # Apply header detection logic for Excel sheets
            sheet_data = find_header_excel(excel_data, sheet_name, header_keywords, file)
            merged_data.append(sheet_data)
        df = pd.concat(merged_data, ignore_index=True)
    else:
        print(f"Unsupported file format: {file}")
        return None
    
    # Normalize columns (convert to lowercase)
    df.columns = df.columns.str.lower()
    
    # Align columns with the first file's columns
    for col in columns_in_first_file:
        if col not in df.columns:
            # If column is missing, add it as the last column with empty values
            df[col] = pd.NA
    
    # Reorder columns to match the first file's column order
    df = df[columns_in_first_file]
    
    return df


def process_file_to_temp(file, header_keywords, columns_in_first_file, temp_dir):
    """
    Process one file and spill to a temporary CSV to avoid large inter-process
    DataFrame transfers (major speed + memory win on big files).
    """
    df = process_file(file, header_keywords, columns_in_first_file)
    if df is None or df.empty:
        return None

    tmp_name = f"{Path(file).stem}_{os.getpid()}.csv"
    tmp_path = os.path.join(temp_dir, tmp_name)
    df.to_csv(tmp_path, index=False, header=False)
    return tmp_path

def merge_files_in_folder(folder_path, output_file, header_keywords):
    """
    Merges all valid files (.csv, .xls, .xlsx) in the given folder into one CSV file.
    Detects the header based on provided header keywords and skips irrelevant rows before the header.
    Ensures columns are aligned.

    Arguments:
    folder_path : str : Path to the folder containing input files
    output_file : str : Path for the final merged output CSV file
    header_keywords : list : List of header keywords to detect the header row
    """
    merge_start = perf_counter()

    # List valid files in the folder
    files_to_process = []
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path) and file_name.lower().endswith(('.csv', '.xls', '.xlsx')):
            files_to_process.append(file_path)
    files_to_process.sort()

    if not files_to_process:
        print("No valid files found in the folder!")
        return
    
    print(f"Processing the following files: {files_to_process}")
    
    # Get columns from the first file to align all subsequent files
    first_file = files_to_process[0]
    first_df = pd.read_csv(first_file) if first_file.lower().endswith('.csv') else pd.read_excel(first_file)
    first_df = find_header_csv(first_df, header_keywords, first_file)
    columns_in_first_file = [col.lower() for col in first_df.columns]

    # Process files in parallel and spill each result into temporary CSVs.
    temp_dir = tempfile.mkdtemp(prefix="life_cycle_merge_", dir=folder_path)
    try:
        # Use full CPU count to maximize parallelism on dedicated runs.
        worker_count = max(1, cpu_count())
        with Pool(processes=worker_count) as pool:
            temp_csv_paths = pool.starmap(
                process_file_to_temp,
                [(file, header_keywords, columns_in_first_file, temp_dir) for file in files_to_process],
                chunksize=1,
            )

        # Remove None values (failed files)
        temp_csv_paths = [path for path in temp_csv_paths if path is not None and os.path.exists(path)]
        if not temp_csv_paths:
            print("No processable data found in input files.")
            return

        # Create output with header first.
        pd.DataFrame(columns=columns_in_first_file).to_csv(output_file, index=False)

        # Append pre-normalized temp CSVs without reparsing.
        with open(output_file, "a", encoding="utf-8", newline="") as out_f:
            for temp_csv in temp_csv_paths:
                with open(temp_csv, "r", encoding="utf-8", newline="") as in_f:
                    shutil.copyfileobj(in_f, out_f)

        total_elapsed = perf_counter() - merge_start
        print(f"Files merged successfully into {output_file} using {worker_count} workers")
        print(f"Total merge duration: {total_elapsed:.2f} seconds")
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == '__main__':
    # Example usage
    folder_path = 'Life Cycle Report'  # Replace this with the path to your folder
    output_file = 'merged_output_life_cycle_report.csv'
    header_keywords = ['Agency Txn Id', 'Settlement Amount', 'Plaza ID', 'Violation Amts']
    merge_files_in_folder(folder_path, output_file, header_keywords)