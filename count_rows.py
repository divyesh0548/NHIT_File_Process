from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed
import os
import dask.dataframe as dd
import pandas as pd


def _resolve_excel_engine(path: Path) -> str:
    """
    Pick openpyxl vs xlrd. Many files are named .xls but are actually OOXML (.xlsx);
    xlrd then fails with: Excel xlsx file; not supported
    """
    suffix = path.suffix.lower()
    try:
        with open(path, "rb") as f:
            header = f.read(8)
    except OSError:
        header = b""

    # .xlsx / OOXML is a ZIP container → starts with PK
    if len(header) >= 2 and header[:2] == b"PK":
        return "openpyxl"

    # Classic binary .xls (OLE compound document)
    if len(header) >= 8 and header[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "xlrd"

    if suffix == ".xlsx":
        return "openpyxl"
    if suffix == ".xls":
        return "xlrd"

    return "openpyxl"


def _print_dataframe_details(df, head_rows=5, tail_rows=5):
    """Print standard profiling output for an already loaded DataFrame."""
    print(f"Shape: {df.shape[0]} rows × {df.shape[1]} columns")

    print("\nColumn names:")
    print(df.columns.tolist())

    print("\nData types:")
    print(df.dtypes)

    numeric_cols = df.select_dtypes(include="number").columns
    if len(numeric_cols) > 0:
        print("\nSummary statistics (numeric columns):")
        print(df.describe())
    else:
        print("\nSummary statistics (numeric columns): none")

    print("\nMissing values per column:")
    print(df.isnull().sum())

    print(f"\nFirst {head_rows} rows (head):")
    print(df.head(head_rows))

    print(f"\nLast {tail_rows} rows (tail):")
    print(df.tail(tail_rows))


def _build_dataframe_profile(df, head_rows=5, tail_rows=5):
    """Build a lightweight profile payload for a DataFrame."""
    numeric_cols = df.select_dtypes(include="number").columns
    if len(numeric_cols) > 0:
        numeric_summary = df.describe()
    else:
        numeric_summary = None

    return {
        "shape": df.shape,
        "columns": df.columns.tolist(),
        "dtypes": df.dtypes,
        "numeric_summary": numeric_summary,
        "missing": df.isnull().sum(),
        "head": df.head(head_rows),
        "tail": df.tail(tail_rows),
    }


def _print_dataframe_profile(profile, head_rows=5, tail_rows=5):
    """Print profile payload produced by _build_dataframe_profile."""
    rows, cols = profile["shape"]
    print(f"Shape: {rows} rows × {cols} columns")

    print("\nColumn names:")
    print(profile["columns"])

    print("\nData types:")
    print(profile["dtypes"])

    if profile["numeric_summary"] is not None:
        print("\nSummary statistics (numeric columns):")
        print(profile["numeric_summary"])
    else:
        print("\nSummary statistics (numeric columns): none")

    print("\nMissing values per column:")
    print(profile["missing"])

    print(f"\nFirst {head_rows} rows (head):")
    print(profile["head"])

    print(f"\nLast {tail_rows} rows (tail):")
    print(profile["tail"])


def print_csv_details(
    file_path,
    chunksize=200000,
    head_rows=5,
    tail_rows=5,
    fast_mode=True,
    high_cpu_mode=True,
 ):
    """
    Print CSV details.

    - fast_mode=True: uses chunked processing (better for very large files)
    - fast_mode=False: loads full DataFrame and runs full describe()
    - high_cpu_mode=True: tries pyarrow engine first (multi-threaded parser)

    Arguments:
    file_path : str : path to CSV
    chunksize : int : rows per chunk in fast mode
    head_rows : int : rows to print for head
    tail_rows : int : rows to print for tail
    fast_mode : bool : enable chunked mode
    high_cpu_mode : bool : prefer high-CPU parser if available
    """
    try:
        if high_cpu_mode:
            try:
                # PyArrow parser is multi-threaded and often uses much more CPU.
                df = pd.read_csv(file_path, engine="pyarrow")
                print("CSV parser mode: pyarrow (high CPU mode)")
                _print_dataframe_details(df, head_rows=head_rows, tail_rows=tail_rows)
                return
            except Exception:
                # Fallback to pandas C/chunk parser.
                pass

        if not fast_mode:
            df = pd.read_csv(file_path, low_memory=False)
            print("CSV parser mode: pandas full-read")
            _print_dataframe_details(df, head_rows=head_rows, tail_rows=tail_rows)
            return

        total_rows = 0
        columns = None
        dtypes = None
        missing_counts = None
        head_df = None
        tail_df = pd.DataFrame()
        numeric_stats = {}

        chunk_iter = pd.read_csv(
            file_path,
            chunksize=chunksize,
            low_memory=False,
            memory_map=True,
        )

        for chunk in chunk_iter:
            if chunk.empty:
                continue

            if columns is None:
                columns = chunk.columns.tolist()
                missing_counts = pd.Series(0, index=columns, dtype="int64")

            if dtypes is None:
                dtypes = chunk.dtypes

            total_rows += len(chunk)
            missing_counts = missing_counts.add(chunk.isnull().sum(), fill_value=0).astype("int64")

            if head_df is None:
                head_df = chunk.head(head_rows)

            tail_df = pd.concat([tail_df, chunk.tail(tail_rows)], ignore_index=True).tail(tail_rows)

            numeric_chunk = chunk.select_dtypes(include="number")
            for col in numeric_chunk.columns:
                series = numeric_chunk[col].dropna()
                if series.empty:
                    continue

                stat = numeric_stats.setdefault(
                    col,
                    {"count": 0, "sum": 0.0, "sum_sq": 0.0, "min": None, "max": None},
                )
                col_count = int(series.count())
                col_sum = float(series.sum())
                col_sum_sq = float((series.astype("float64") ** 2).sum())
                col_min = float(series.min())
                col_max = float(series.max())

                stat["count"] += col_count
                stat["sum"] += col_sum
                stat["sum_sq"] += col_sum_sq
                stat["min"] = col_min if stat["min"] is None else min(stat["min"], col_min)
                stat["max"] = col_max if stat["max"] is None else max(stat["max"], col_max)

        if columns is None:
            print("CSV is empty or unreadable.")
            return

        print("CSV parser mode: pandas chunked")
        print(f"Shape: {total_rows} rows × {len(columns)} columns")

        print("\nColumn names:")
        print(columns)

        print("\nData types (inferred from first chunk):")
        print(dtypes)

        if numeric_stats:
            summary_rows = {}
            for col, stat in numeric_stats.items():
                cnt = stat["count"]
                mean = stat["sum"] / cnt if cnt else float("nan")
                variance = (stat["sum_sq"] / cnt) - (mean * mean) if cnt else float("nan")
                variance = max(variance, 0.0) if cnt else variance
                std = variance ** 0.5 if cnt else float("nan")
                summary_rows[col] = {
                    "count": cnt,
                    "mean": mean,
                    "std": std,
                    "min": stat["min"],
                    "max": stat["max"],
                }
            summary_df = pd.DataFrame(summary_rows)
            print("\nFast numeric summary (count/mean/std/min/max):")
            print(summary_df)
        else:
            print("\nFast numeric summary: none")

        print("\nMissing values per column:")
        print(missing_counts)

        print(f"\nFirst {head_rows} rows (head):")
        print(head_df.head(head_rows) if head_df is not None else pd.DataFrame())

        print(f"\nLast {tail_rows} rows (tail):")
        print(tail_df.tail(tail_rows))

    except Exception as e:
        print(f"Error: {e}")


def _read_excel_sheet(path, engine, name):
    """Read one Excel sheet (helper for threaded execution)."""
    return name, pd.read_excel(path, sheet_name=name, engine=engine)


def _read_and_profile_excel_sheet(path, engine, name, head_rows, tail_rows):
    """Worker helper for multiprocessing: read one sheet and return compact profile."""
    df = pd.read_excel(path, sheet_name=name, engine=engine)
    return name, _build_dataframe_profile(df, head_rows=head_rows, tail_rows=tail_rows)


def print_excel_details(
    file_path,
    sheet_name=None,
    head_rows=5,
    tail_rows=5,
    use_threading=True,
    use_multiprocessing=True,
    max_workers=10,
  ):
    """
    Print details for an Excel workbook (.xlsx or .xls): row/column counts, dtypes,
    summary stats, missing values, head and tail.

    If sheet_name is None, every sheet in the workbook is summarized in order.

    Arguments:
        file_path : str or pathlib.Path : Path to the .xlsx or .xls file
        sheet_name : str or int or None : Sheet to read; None reads all sheets
        head_rows : int : Rows to show in head()
        tail_rows : int : Rows to show in tail()
        use_threading : bool : parallelize reading multiple sheets
        use_multiprocessing : bool : use process pool for multi-sheet CPU-heavy parsing
        max_workers : int : worker count for parallel mode
    """
    path = Path(file_path)
    suffix = path.suffix.lower()
    if suffix not in (".xlsx", ".xls"):
        print(f"Error: expected .xlsx or .xls, got {suffix or '(no extension)'}")
        return

    engine = _resolve_excel_engine(path)
    try:
        xl = pd.ExcelFile(path, engine=engine)
    except ImportError as e:
        if engine == "openpyxl":
            print(
                "Error: reading .xlsx / OOXML workbooks requires openpyxl. "
                "Install with: pip install openpyxl"
            )
        else:
            print(
                "Error: reading legacy .xls requires xlrd. "
                "Install with: pip install xlrd"
            )
        print(f"Details: {e}")
        return
    except Exception as e:
        # Extension said .xls but sniff missed; retry as xlsx if xlrd complains
        err_msg = str(e).lower()
        if engine == "xlrd" and ("xlsx" in err_msg or "not supported" in err_msg):
            try:
                xl = pd.ExcelFile(path, engine="openpyxl")
            except ImportError:
                print(
                    "Error: this file is actually an .xlsx (OOXML) workbook. "
                    "Install openpyxl: pip install openpyxl"
                )
                return
            except Exception as e2:
                print(f"Error opening workbook: {e2}")
                return
        else:
            print(f"Error opening workbook: {e}")
            return

    sheets = [sheet_name] if sheet_name is not None else xl.sheet_names

    loaded_sheets = {}
    detected_workers = max(1, int(max_workers))

    if use_multiprocessing and len(sheets) > 1:
        worker_count = min(detected_workers, len(sheets))
        print(f"Excel parser mode: multiprocessing ({worker_count} workers)")
        try:
            with ProcessPoolExecutor(max_workers=worker_count) as executor:
                futures = {
                    executor.submit(
                        _read_and_profile_excel_sheet,
                        str(path),
                        engine,
                        name,
                        head_rows,
                        tail_rows,
                    ): name
                    for name in sheets
                }
                for future in as_completed(futures):
                    name = futures[future]
                    try:
                        sheet_name_loaded, profile = future.result()
                        loaded_sheets[sheet_name_loaded] = profile
                    except Exception as e:
                        print(f"\n--- Sheet: {name!r} ---\nError reading sheet: {e}")
        except Exception as e:
            print(f"Multiprocessing mode failed, falling back to threading/sequential. Details: {e}")

    if not loaded_sheets and use_threading and len(sheets) > 1:
        worker_count = min(4, len(sheets))
        print(f"Excel parser mode: threading ({worker_count} workers)")
        with ThreadPoolExecutor(max_workers=worker_count) as executor:
            futures = {
                executor.submit(_read_excel_sheet, path, engine, name): name for name in sheets
            }
            for future in as_completed(futures):
                name = futures[future]
                try:
                    sheet_name_loaded, df = future.result()
                    loaded_sheets[sheet_name_loaded] = _build_dataframe_profile(
                        df, head_rows=head_rows, tail_rows=tail_rows
                    )
                except Exception as e:
                    print(f"\n--- Sheet: {name!r} ---\nError reading sheet: {e}")

    if not loaded_sheets:
        print("Excel parser mode: sequential")
        for name in sheets:
            try:
                df = pd.read_excel(xl, sheet_name=name)
                loaded_sheets[name] = _build_dataframe_profile(
                    df, head_rows=head_rows, tail_rows=tail_rows
                )
            except Exception as e:
                print(f"\n--- Sheet: {name!r} ---\nError reading sheet: {e}")

    for name in sheets:
        if name not in loaded_sheets:
            continue
        print(f"\n{'=' * 60}\nSheet: {name!r}\n{'=' * 60}")
        _print_dataframe_profile(loaded_sheets[name], head_rows=head_rows, tail_rows=tail_rows)


def _normalize_header_value(value):
    """Normalize raw header cell values for matching."""
    if pd.isna(value):
        return ""
    return str(value).strip().lower()

def _detect_header_row_index(df_raw, header_keywords, min_matches=3):
    """
    Detect header row index by scanning row values and matching keywords.
    Returns 0 if not confidently detected.
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

def _count_csv_with_header_detection(file_path, header_keywords):
    """
    Return (row_count_excluding_header, column_count, detected_header_row_index)
    for a CSV file.
    """
    df_raw = pd.read_csv(file_path, header=None, dtype=str, low_memory=False)
    if df_raw.empty:
        return 0, 0, 0

    header_idx = _detect_header_row_index(df_raw, header_keywords)
    df = pd.read_csv(file_path, skiprows=header_idx, header=0, low_memory=False)
    return len(df), len(df.columns), header_idx


def _count_excel_sheet_with_header_detection(file_path, sheet_name, header_keywords):
    """
    Return (row_count_excluding_header, column_count, detected_header_row_index)
    for one Excel sheet with header detection.
    """
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=str)
    if df_raw.empty:
        return 0, 0, 0

    header_idx = _detect_header_row_index(df_raw, header_keywords)
    df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=header_idx, header=0)
    return len(df), len(df.columns), header_idx


def _count_single_file(file_path, header_keywords, show_plus_header_row):
    """Count one file and return printable lines plus raw row total (header excluded)."""
    lines = []
    raw_rows_total = 0
    suffix = file_path.suffix.lower()

    try:
        if suffix == ".csv":
            # Use Dask for better performance on large CSV files
            df = dd.read_csv(file_path, dtype=str)  # Use Dask to load CSV files
            lines.append(f"{file_path.name} | type=CSV | rows={df.shape[0].compute()} | cols={len(df.columns)}")
            raw_rows_total += df.shape[0].compute()
        else:
            engine = _resolve_excel_engine(file_path)
            xl = pd.ExcelFile(file_path, engine=engine)
            for sheet in xl.sheet_names:
                rows, cols, header_idx = _count_excel_sheet_with_header_detection(file_path, sheet, header_keywords)
                lines.append(f"{file_path.name} | sheet={sheet} | header_row={header_idx} | rows={rows} | cols={cols}")
                raw_rows_total += rows

    except Exception as exc:
        lines.append(f"{file_path.name} | ERROR: {exc}")

    return file_path.name, lines, raw_rows_total



def count_rows_in_folder(folder_path, header_keywords, show_plus_header_row=False):
    """
    Count rows and columns for each .csv/.xls/.xlsx file in a folder.
    Optimized to handle large files efficiently.
    """
    folder = Path(folder_path)
    if not folder.exists() or not folder.is_dir():
        print(f"Invalid folder path: {folder_path}")
        return

    files = sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in {".csv", ".xls", ".xlsx"}])
    if not files:
        print("No .csv/.xls/.xlsx files found in the folder.")
        return

    total_rows = 0
    print(f"Scanning folder: {folder.resolve()}")

    max_workers = min(10, len(files))
    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        futures = {
            executor.submit(_count_single_file, file_path, header_keywords, show_plus_header_row): file_path
            for file_path in files
        }
        for future in as_completed(futures):
            _file_name, lines, file_rows = future.result()
            for line in lines:
                print(line)
            total_rows += file_rows

    print(f"TOTAL ROWS (combined, header excluded): {total_rows}")

# Example usage
if __name__ == "__main__":
    file_path = "Life Cycle Report/Life Cycle Report Feb-2026.xlsx"  # Replace with your CSV file path
    # print_csv_details(file_path)
    # print_excel_details(file_path)
    header_keywords = ['Agency Txn Id', 'Settlement Amount', 'Plaza ID', 'Violation Amts']
    count_rows_in_folder("Life Cycle Report", header_keywords)