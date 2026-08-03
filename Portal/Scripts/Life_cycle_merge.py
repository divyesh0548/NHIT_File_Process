import os
import shutil
import tempfile
from multiprocessing import Pool, cpu_count
from pathlib import Path
import sys
from time import perf_counter

import pandas as pd

_PORTAL_ROOT = Path(__file__).resolve().parent.parent
if str(_PORTAL_ROOT) not in sys.path:
    sys.path.insert(0, str(_PORTAL_ROOT))

from header_matching import normalize_header_match

HEADER_SCAN_ROWS = 50
FINAL_MERGE_CHUNK_ROWS = 100_000
MIN_HEADER_MATCHES = 3


class HeaderNotDetectedError(ValueError):
    """Raised when a file/sheet does not contain enough configured header keywords."""


def _normalize_col_name(col):
    return normalize_header_match(col)


def _detect_header_row_index(df_raw, header_keywords, min_matches=MIN_HEADER_MATCHES):
    """
    Detect the header row by scanning a sample from the top of the sheet.

    Returns the row index when at least min_matches keywords are found.
    Returns None when no confident match is found (does not fall back to row 0).
    """
    target = {normalize_header_match(k) for k in header_keywords if normalize_header_match(k)}
    best_idx = None
    best_matches = 0

    for idx in range(len(df_raw)):
        row_values = {
            normalize_header_match(v)
            for v in df_raw.iloc[idx].tolist()
            if normalize_header_match(v)
        }
        match_count = len(target.intersection(row_values))
        if match_count > best_matches:
            best_matches = match_count
            best_idx = idx
        if match_count >= min_matches:
            return idx

    if best_idx is not None and best_matches >= min_matches:
        return best_idx
    return None


def _read_csv_fast(file_path):
    try:
        return pd.read_csv(file_path, engine="pyarrow")
    except Exception:
        return pd.read_csv(file_path, low_memory=False)


def find_header_csv(df, header_keywords, file_name, scan_rows=HEADER_SCAN_ROWS, file_path=None):
    """
    Finds the header row in a CSV based on configured header keywords.
    Requires at least MIN_HEADER_MATCHES matches; otherwise raises HeaderNotDetectedError.
    """
    norm_to_original = {_normalize_col_name(c): c for c in df.columns}
    matched_keywords = []
    for keyword in header_keywords:
        nk = _normalize_col_name(keyword)
        if nk in norm_to_original:
            matched_keywords.append(norm_to_original[nk])

    if len(matched_keywords) >= MIN_HEADER_MATCHES:
        start_idx = df.columns.get_loc(matched_keywords[0])
        print(
            f"Header detected in {file_name} at column index {start_idx} "
            f"with keywords {matched_keywords}"
        )
        return df.iloc[:, start_idx:]

    # Row 0 columns did not match — scan first N rows for a header row.
    if file_path is None:
        raise HeaderNotDetectedError(
            f"No header detected in {Path(file_name).name}. "
            f"Need at least {MIN_HEADER_MATCHES} header keywords. Check the header keywords."
        )

    try:
        df_sample = pd.read_csv(
            file_path, header=None, dtype=str, nrows=scan_rows, low_memory=False
        )
    except Exception:
        df_sample = pd.read_csv(
            file_path,
            header=None,
            dtype=str,
            nrows=scan_rows,
            low_memory=False,
            encoding="cp1252",
        )

    if df_sample.empty:
        raise HeaderNotDetectedError(
            f"No header detected in {Path(file_name).name}. "
            f"Need at least {MIN_HEADER_MATCHES} header keywords. Check the header keywords."
        )

    header_idx = _detect_header_row_index(df_sample, header_keywords)
    if header_idx is None:
        raise HeaderNotDetectedError(
            f"No header detected in {Path(file_name).name}. "
            f"Need at least {MIN_HEADER_MATCHES} header keywords. Check the header keywords."
        )

    print(f"Header detected in {file_name} at row index {header_idx}")
    try:
        return pd.read_csv(file_path, skiprows=header_idx, header=0, engine="pyarrow")
    except Exception:
        return pd.read_csv(file_path, skiprows=header_idx, header=0, low_memory=False)


def find_header_excel(excel_data, sheet_name, header_keywords, file_name, scan_rows=HEADER_SCAN_ROWS):
    """
    Detect the header row using only the first few rows, then read the sheet once fully.

    Returns:
      - DataFrame when header is detected
      - None when the sheet is empty
    Raises HeaderNotDetectedError when the sheet has data but < MIN_HEADER_MATCHES keywords.
    """
    df_sample = pd.read_excel(
        excel_data,
        sheet_name=sheet_name,
        header=None,
        dtype=str,
        nrows=scan_rows,
    )
    if df_sample.empty:
        return None

    header_idx = _detect_header_row_index(df_sample, header_keywords)
    if header_idx is None:
        raise HeaderNotDetectedError(
            f"No header detected in {Path(file_name).name} [{sheet_name}]. "
            f"Need at least {MIN_HEADER_MATCHES} header keywords. Check the header keywords."
        )

    print(f"Header detected in {file_name} [{sheet_name}] at row index {header_idx}")
    return pd.read_excel(excel_data, sheet_name=sheet_name, skiprows=header_idx, header=0)


def _load_dataframe_for_file(file, header_keywords):
    """Load one file once, including all Excel sheets when present."""
    if file.lower().endswith(".csv"):
        df = _read_csv_fast(file)
        return find_header_csv(df, header_keywords, file, file_path=file)

    if file.lower().endswith((".xls", ".xlsx")):
        excel_data = pd.ExcelFile(file)
        merged_data = []
        sheet_errors = []
        for sheet_name in excel_data.sheet_names:
            try:
                sheet_data = find_header_excel(excel_data, sheet_name, header_keywords, file)
            except HeaderNotDetectedError as exc:
                sheet_errors.append(str(exc))
                print(f"Header not detected — skipping sheet: {exc}")
                continue
            if sheet_data is not None and not sheet_data.empty:
                merged_data.append(sheet_data)

        if merged_data:
            return pd.concat(merged_data, ignore_index=True)

        # No usable sheet: block merge with a clear message.
        detail = sheet_errors[0] if sheet_errors else (
            f"No header detected in {Path(file).name}. "
            f"Need at least {MIN_HEADER_MATCHES} header keywords. Check the header keywords."
        )
        raise HeaderNotDetectedError(detail)

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

    Raises HeaderNotDetectedError when any input file has no sheet/row with enough
    header keywords — merge is blocked in that case.
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
        # Process sequentially first for clearer header errors; use pool only after
        # validating would lose exception clarity with multiprocessing wrappers.
        # Keep parallel processing but unwrap HeaderNotDetectedError for the portal.
        worker_count = max(1, cpu_count())
        with Pool(processes=worker_count) as pool:
            async_results = [
                pool.apply_async(process_file_to_temp, (file, header_keywords, temp_dir))
                for file in files_to_process
            ]
            staged_results = []
            for file, async_result in zip(files_to_process, async_results):
                try:
                    staged_results.append(async_result.get())
                except HeaderNotDetectedError:
                    raise
                except Exception as exc:
                    # multiprocessing may wrap remote exceptions
                    message = str(exc)
                    if "No header detected" in message or "HeaderNotDetectedError" in message:
                        raise HeaderNotDetectedError(
                            f"No header detected in {Path(file).name}. "
                            f"Need at least {MIN_HEADER_MATCHES} header keywords. "
                            "Check the header keywords."
                        ) from exc
                    raise

        staged_files = [
            staged_result
            for staged_result in staged_results
            if staged_result is not None and os.path.exists(staged_result["temp_path"])
        ]
        if not staged_files:
            raise HeaderNotDetectedError(
                f"No header detected in the uploaded files. "
                f"Need at least {MIN_HEADER_MATCHES} header keywords. Check the header keywords."
            )

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
