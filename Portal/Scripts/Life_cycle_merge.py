import os
import shutil
import tempfile
from multiprocessing import Pool, cpu_count
from pathlib import Path
from time import perf_counter

import pandas as pd

HEADER_SCAN_ROWS = 50
FINAL_MERGE_CHUNK_ROWS = 100_000


def _normalize_header_value(value):
    if pd.isna(value):
        return ""
    return str(value).strip().lower()


def _normalize_col_name(col):
    s = str(col).strip().lower()
    return "".join(s.split())


def _detect_header_row_index(df_raw, header_keywords, min_matches=3):
    """
    Detect the header row by scanning only a small sample from the top of the sheet.
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


def _read_csv_fast(file_path):
    try:
        return pd.read_csv(file_path, engine="pyarrow")
    except Exception:
        return pd.read_csv(file_path, low_memory=False)


def find_header_csv(df, header_keywords, file_name):
    """
    Finds the header row in a CSV DataFrame based on the presence of given header keywords.
    At least 3 keywords must be detected in the columns (matched with normalized column names).
    """
    norm_to_original = {_normalize_col_name(c): c for c in df.columns}
    matched_keywords = []
    for keyword in header_keywords:
        nk = _normalize_col_name(keyword)
        if nk in norm_to_original:
            matched_keywords.append(norm_to_original[nk])

    if len(matched_keywords) >= 3:
        start_idx = df.columns.get_loc(matched_keywords[0])
        print(f"Header detected in {file_name} at column index {start_idx} with keywords {matched_keywords}")
        df = df.iloc[:, start_idx:]

    return df


def find_header_excel(excel_data, sheet_name, header_keywords, file_name, scan_rows=HEADER_SCAN_ROWS):
    """
    Detect the header row using only the first few rows, then read the sheet once fully.
    This avoids a full-sheet read just for header detection.
    """
    df_sample = pd.read_excel(
        excel_data,
        sheet_name=sheet_name,
        header=None,
        dtype=str,
        nrows=scan_rows,
    )
    if df_sample.empty:
        return pd.DataFrame()

    header_idx = _detect_header_row_index(df_sample, header_keywords)
    print(f"Header detected in {file_name} [{sheet_name}] at row index {header_idx}")

    return pd.read_excel(excel_data, sheet_name=sheet_name, skiprows=header_idx, header=0)


def _load_dataframe_for_file(file, header_keywords):
    """Load one file once, including all Excel sheets when present."""
    if file.lower().endswith(".csv"):
        df = _read_csv_fast(file)
        return find_header_csv(df, header_keywords, file)

    if file.lower().endswith((".xls", ".xlsx")):
        excel_data = pd.ExcelFile(file)
        merged_data = []
        for sheet_name in excel_data.sheet_names:
            sheet_data = find_header_excel(excel_data, sheet_name, header_keywords, file)
            if not sheet_data.empty:
                merged_data.append(sheet_data)
        if not merged_data:
            return pd.DataFrame()
        return pd.concat(merged_data, ignore_index=True)

    return None


def _normalize_dataframe_columns(df):
    """
    Normalize column names once, collapse duplicates, and preserve first-seen display names.
    Returns the normalized DataFrame and column metadata for the final merge.
    """
    rename_map = {}
    normalized_order = []
    normalized_display_names = {}
    seen_norm = set()

    for raw_col in df.columns:
        normalized = _normalize_col_name(raw_col)
        display_name = str(raw_col).strip()
        rename_map[raw_col] = normalized
        if normalized not in seen_norm:
            seen_norm.add(normalized)
            normalized_order.append(normalized)
            normalized_display_names[normalized] = display_name

    df = df.rename(columns=rename_map)

    if df.columns.duplicated().any():
        df = df.loc[:, ~df.columns.duplicated()].copy()

    return df, normalized_order, normalized_display_names


def process_file_to_temp(file, header_keywords, temp_dir):
    """
    Parse each input file once, normalize its headers, and spill to a staging CSV.
    The final merge then works from staging CSVs instead of reopening Excel files.
    """
    df = _load_dataframe_for_file(file, header_keywords)
    if df is None:
        print(f"Unsupported file format: {file}")
        return None
    if df.empty:
        return None

    df, normalized_order, normalized_display_names = _normalize_dataframe_columns(df)

    tmp_name = f"{Path(file).stem}_{os.getpid()}.csv"
    tmp_path = os.path.join(temp_dir, tmp_name)
    df.to_csv(tmp_path, index=False)

    return {
        "temp_path": tmp_path,
        "normalized_order": normalized_order,
        "normalized_display_names": normalized_display_names,
    }


def _build_canonical_columns(staged_files):
    canonical_order = []
    norm_to_canonical = {}
    seen_norm = set()

    for staged_file in staged_files:
        for normalized in staged_file["normalized_order"]:
            if normalized not in seen_norm:
                seen_norm.add(normalized)
                canonical_order.append(normalized)
                norm_to_canonical[normalized] = staged_file["normalized_display_names"][normalized]

    return canonical_order, norm_to_canonical


def _append_staged_csvs(output_file, staged_files, canonical_order, norm_to_canonical):
    output_headers = [norm_to_canonical[col] for col in canonical_order]
    pd.DataFrame(columns=output_headers).to_csv(output_file, index=False)

    for staged_file in staged_files:
        staged_path = staged_file["temp_path"]
        for chunk in pd.read_csv(
            staged_path,
            chunksize=FINAL_MERGE_CHUNK_ROWS,
            low_memory=False,
            dtype=str,
        ):
            chunk = chunk.reindex(columns=canonical_order)
            chunk.columns = output_headers
            chunk.to_csv(output_file, mode="a", header=False, index=False)


def merge_files_in_folder(folder_path, output_file, header_keywords):
    """
    Merge all valid files in the folder into one CSV.
    Expensive parsing is done once per source file, then final alignment happens from staged CSVs.
    """
    merge_start = perf_counter()

    files_to_process = []
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path) and file_name.lower().endswith((".csv", ".xls", ".xlsx")):
            files_to_process.append(file_path)
    files_to_process.sort()

    if not files_to_process:
        print("No valid files found in the folder!")
        return

    print(f"Processing the following files: {files_to_process}")

    temp_dir = tempfile.mkdtemp(prefix="life_cycle_merge_", dir=folder_path)
    try:
        worker_count = max(1, cpu_count())
        with Pool(processes=worker_count) as pool:
            staged_results = pool.starmap(
                process_file_to_temp,
                [(file, header_keywords, temp_dir) for file in files_to_process],
                chunksize=1,
            )

        staged_files = [
            staged_result
            for staged_result in staged_results
            if staged_result is not None and os.path.exists(staged_result["temp_path"])
        ]
        if not staged_files:
            print("No processable data found in input files.")
            return

        canonical_order, norm_to_canonical = _build_canonical_columns(staged_files)
        print(f"Union column count: {len(canonical_order)}")

        _append_staged_csvs(output_file, staged_files, canonical_order, norm_to_canonical)

        total_elapsed = perf_counter() - merge_start
        print(f"Files merged successfully into {output_file} using {worker_count} workers")
        print(f"Total merge duration: {total_elapsed:.2f} seconds")
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    import sys

    portal_root = Path(__file__).resolve().parent.parent
    if str(portal_root) not in sys.path:
        sys.path.insert(0, str(portal_root))

    from db.nhit_file_process import get_lc_etc_header_keyword_strings

    folder_path = "Life Cycle Report"
    output_file = "merged_output_life_cycle_report_with_all_columns.csv"
    header_keywords = get_lc_etc_header_keyword_strings()
    merge_files_in_folder(folder_path, output_file, header_keywords)
