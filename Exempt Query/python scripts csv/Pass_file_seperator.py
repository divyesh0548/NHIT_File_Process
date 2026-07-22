"""
Split pass working Excel files by "Pass Type" into MP / LT / Other outputs.

Reads every .xls / .xlsx in INPUT_FOLDER, detects the header row using
HEADER_KEYWORDS, then routes each row by the Pass Type column into:
  - MP_Pass.xlsx
  - LT_Pass.xlsx
  - Other_Pass.xlsx (only written when non-MP/LT rows exist)

Columns are matched across files by normalized header name. Known columns align;
new columns are appended at the end with blanks for earlier rows.
"""

from __future__ import annotations

import sys
from pathlib import Path
from time import perf_counter

import pandas as pd

from merge_files import (
    HEADER_SCAN_ROWS,
    _normalize_col_name,
    _normalize_dataframe_columns,
    find_header_excel,
)

BASE_DIR = Path(__file__).resolve().parent

# --- Configuration (edit before running) ---
INPUT_FOLDER = BASE_DIR / "Daroda_lc_vrn_files/pass"
OUTPUT_FOLDER = BASE_DIR / "pass_output"

# At least 3 of these must appear in the detected header row.
HEADER_KEYWORDS = [
    "Chassis/ Vehicle No",
    "NPCI Vehicle Class",
    "Start Date",
    "Plaza Name",
    "End Date",
    "Mobile No",
    "Pass Type",
    "Payment Mode"
]

# Header aliases for the pass-type column (normalized match).
PASS_TYPE_COLUMN_NAMES = [
    "Pass Type",
    "PASS TYPE",
    "PassType",
    "PASS_TYPE",
]

MP_VALUE = "MP"
LT_VALUE = "LT"

OUTPUT_MP = "MP_Pass.xlsx"
OUTPUT_LT = "LT_Pass.xlsx"
OUTPUT_OTHER = "Other_Pass.xlsx"

SUPPORTED_EXTENSIONS = {".xls", ".xlsx"}
MIN_HEADER_MATCHES = 3


def _pass_type_targets() -> set[str]:
    return {_normalize_col_name(name) for name in PASS_TYPE_COLUMN_NAMES}


def find_pass_type_column(columns) -> str | None:
    targets = _pass_type_targets()
    for col in columns:
        if _normalize_col_name(col) in targets:
            return col
    return None


def normalize_pass_type_value(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


def load_excel_sheets(file_path: Path, header_keywords: list[str]) -> list[pd.DataFrame]:
    frames: list[pd.DataFrame] = []
    excel_data = pd.ExcelFile(file_path)
    for sheet_name in excel_data.sheet_names:
        df = find_header_excel(
            excel_data,
            sheet_name,
            header_keywords,
            str(file_path),
            scan_rows=HEADER_SCAN_ROWS,
        )
        if not df.empty:
            frames.append(df)
    return frames


def split_by_pass_type(df: pd.DataFrame, source_label: str):
    pass_col = find_pass_type_column(df.columns)
    if pass_col is None:
        print(f"  [WARN] No Pass Type column in {source_label}; skipping.", flush=True)
        return None, None, None

    mp_rows = []
    lt_rows = []
    other_rows = []

    for _, row in df.iterrows():
        bucket = normalize_pass_type_value(row[pass_col])
        if bucket == MP_VALUE:
            mp_rows.append(row)
        elif bucket == LT_VALUE:
            lt_rows.append(row)
        elif bucket:
            other_rows.append(row)

    def _to_frame(rows):
        return pd.DataFrame(rows).reset_index(drop=True) if rows else pd.DataFrame()

    return _to_frame(mp_rows), _to_frame(lt_rows), _to_frame(other_rows)


def merge_frames(frames: list[pd.DataFrame]) -> pd.DataFrame:
    """Union columns by normalized name; append new columns at the end."""
    if not frames:
        return pd.DataFrame()

    canonical_order: list[str] = []
    norm_to_display: dict[str, str] = {}
    aligned: list[pd.DataFrame] = []

    for df in frames:
        if df.empty:
            continue

        normalized_df, order, display_names = _normalize_dataframe_columns(df.copy())
        for norm in order:
            if norm not in norm_to_display:
                canonical_order.append(norm)
                norm_to_display[norm] = display_names[norm]

        aligned.append(normalized_df.reindex(columns=canonical_order))

    if not aligned:
        return pd.DataFrame()

    merged = pd.concat(aligned, ignore_index=True)
    merged.columns = [norm_to_display[norm] for norm in canonical_order]
    return merged


def write_output(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(path, index=False, engine="openpyxl")


def separate_pass_files(
    input_folder: Path,
    output_folder: Path,
    header_keywords: list[str],
) -> None:
    if len(header_keywords) < MIN_HEADER_MATCHES:
        raise ValueError(
            f"HEADER_KEYWORDS must contain at least {MIN_HEADER_MATCHES} entries."
        )

    input_folder = Path(input_folder).resolve()
    output_folder = Path(output_folder).resolve()

    if not input_folder.is_dir():
        raise FileNotFoundError(f"Input folder not found: {input_folder}")

    files = sorted(
        p
        for p in input_folder.iterdir()
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS
    )

    if not files:
        print(f"No Excel files found in {input_folder}")
        return

    print(f"Input folder:  {input_folder}")
    print(f"Output folder: {output_folder}")
    print(f"Files to process ({len(files)}):")
    for file_path in files:
        print(f"  - {file_path.name}")

    mp_frames: list[pd.DataFrame] = []
    lt_frames: list[pd.DataFrame] = []
    other_frames: list[pd.DataFrame] = []

    for file_path in files:
        print(f"\nProcessing: {file_path.name}", flush=True)
        try:
            sheets = load_excel_sheets(file_path, header_keywords)
        except Exception as exc:
            print(f"  [ERROR] Could not read {file_path.name}: {exc}", flush=True)
            continue

        if not sheets:
            print("  [WARN] No data rows found.", flush=True)
            continue

        for sheet_idx, df in enumerate(sheets, start=1):
            label = f"{file_path.name} (sheet {sheet_idx})"
            mp_part, lt_part, other_part = split_by_pass_type(df, label)
            if mp_part is None:
                continue
            if not mp_part.empty:
                mp_frames.append(mp_part)
            if not lt_part.empty:
                lt_frames.append(lt_part)
            if not other_part.empty:
                other_frames.append(other_part)
            print(
                f"  {label}: MP={len(mp_part)}, LT={len(lt_part)}, Other={len(other_part)}",
                flush=True,
            )

    mp_merged = merge_frames(mp_frames)
    lt_merged = merge_frames(lt_frames)
    other_merged = merge_frames(other_frames)

    mp_path = output_folder / OUTPUT_MP
    lt_path = output_folder / OUTPUT_LT
    other_path = output_folder / OUTPUT_OTHER

    write_output(mp_merged, mp_path)
    write_output(lt_merged, lt_path)
    print(f"\n[OUT] {mp_path} ({len(mp_merged)} rows, {len(mp_merged.columns)} columns)")
    print(f"[OUT] {lt_path} ({len(lt_merged)} rows, {len(lt_merged.columns)} columns)")

    if other_frames:
        write_output(other_merged, other_path)
        print(
            f"[OUT] {other_path} ({len(other_merged)} rows, {len(other_merged.columns)} columns)"
        )
    else:
        print("[OUT] No Other_Pass file (no rows outside MP / LT).")


def main() -> int:
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        print("Requires openpyxl: pip install openpyxl", file=sys.stderr)
        return 1

    start = perf_counter()
    try:
        separate_pass_files(INPUT_FOLDER, OUTPUT_FOLDER, HEADER_KEYWORDS)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"\nDone in {perf_counter() - start:.2f} seconds")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
