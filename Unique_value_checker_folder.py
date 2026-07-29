import csv
import json
from pathlib import Path

import pandas as pd

# ──────────────── Configuration ────────────────
FOLDER_PATH = r"C:\Divyesh\NHIT_File_process\Exempt Query\Madai base files\pass"
COLUMN_NAME = "Pass Type"
# ───────────────────────────────────────────────

SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".csv"}
HEADER_SCAN_ROWS = 25


def _read_csv_flexible(file_path: Path) -> pd.DataFrame:
    """Read CSV allowing uneven field counts (pads shorter rows / keeps extra fields)."""
    rows = []
    with open(file_path, newline="", encoding="utf-8", errors="replace") as handle:
        for row in csv.reader(handle):
            rows.append(row)
    if not rows:
        return pd.DataFrame()
    width = max(len(row) for row in rows)
    padded = [row + [""] * (width - len(row)) for row in rows]
    return pd.DataFrame(padded, dtype=str)


def _apply_header(raw: pd.DataFrame, column: str) -> pd.DataFrame:
    """Scan first HEADER_SCAN_ROWS for the target column and promote that row to header."""
    if raw.empty:
        return raw

    target = column.strip().lower()
    scan_limit = min(HEADER_SCAN_ROWS, len(raw))
    header_row = None
    for i in range(scan_limit):
        row_vals = [
            str(v).strip().lower()
            for v in raw.iloc[i].tolist()
            if pd.notna(v) and str(v).strip()
        ]
        if target in row_vals:
            header_row = i
            break

    if header_row is None:
        header_row = 0

    headers = [
        str(v).strip() if pd.notna(v) and str(v).strip() else f"Unnamed_{idx}"
        for idx, v in enumerate(raw.iloc[header_row].tolist())
    ]
    # Deduplicate blank / repeated headers
    seen = {}
    unique_headers = []
    for name in headers:
        if name in seen:
            seen[name] += 1
            unique_headers.append(f"{name}_{seen[name]}")
        else:
            seen[name] = 0
            unique_headers.append(name)

    body = raw.iloc[header_row + 1 :].copy()
    body.columns = unique_headers
    body = body.reset_index(drop=True)
    return body


def read_file(file_path: Path, column: str) -> pd.DataFrame:
    """Read file, scanning up to HEADER_SCAN_ROWS to find the header row."""
    ext = file_path.suffix.lower()
    if ext == ".csv":
        raw = _read_csv_flexible(file_path)
    else:
        raw = pd.read_excel(file_path, header=None, dtype=str)
    return _apply_header(raw, column)


def get_distinct_values(folder: str, column: str) -> pd.DataFrame:
    folder_path = Path(folder).resolve()
    if not folder_path.is_dir():
        raise FileNotFoundError(f"Folder not found: {folder_path}")

    all_values = []
    files_processed = 0

    for file in sorted(folder_path.iterdir()):
        if not file.is_file() or file.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
        try:
            df = read_file(file, column)
        except Exception as e:
            print(f"  [SKIP] {file.name}: {e}")
            continue

        matched_col = None
        target = column.strip().lower()
        for col in df.columns:
            if str(col).strip().lower() == target:
                matched_col = col
                break

        if matched_col is None:
            print(f"  [SKIP] {file.name}: column '{column}' not found")
            continue

        files_processed += 1
        values = df[matched_col].dropna().unique().tolist()
        for v in values:
            all_values.append({"File": file.name, "Value": str(v).strip()})
        print(f"  [OK]   {file.name}: {len(values)} unique value(s)")

    if not all_values:
        print(f"\nNo values found for column '{column}' in any file.")
        return pd.DataFrame(columns=["Value", "File Count", "Files"])

    result_df = pd.DataFrame(all_values)
    summary = (
        result_df.groupby("Value", sort=False)
        .agg(**{"File Count": ("File", "nunique"), "Files": ("File", lambda x: ", ".join(sorted(x.unique())))})
        .reset_index()
        .sort_values("Value", key=lambda s: s.str.lower())
        .reset_index(drop=True)
    )

    print(f"\nFiles processed: {files_processed}")
    print(f"Distinct values: {len(summary)}\n")
    return summary


if __name__ == "__main__":
    print(f"Folder : {FOLDER_PATH}")
    print(f"Column : {COLUMN_NAME}\n")
    result = get_distinct_values(FOLDER_PATH, COLUMN_NAME)
    if not result.empty:
        print(result.to_string(index=False))

    output_path = Path(__file__).resolve().parent / "Unique_values.json"
    output_data = {
        "folder": FOLDER_PATH,
        "column": COLUMN_NAME,
        "distinct_count": len(result),
        "values": result.to_dict(orient="records"),
    }
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(output_data, f, indent=2, ensure_ascii=False)
    print(f"\nJSON saved to: {output_path}")
