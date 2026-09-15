"""
Merge Concessionaire / Project Vehicle files into one VRN list.

Standalone helper for `Aprooved Exemption (1).py` EXEMPTION_FILE input.

- Takes one INPUT_FOLDER and merges all .xlsx / .xls / .csv files in it
- Scans the first HEADER_SCAN_ROWS rows to locate the VRN header cell
- Detects the vehicle-number column using VRN_COLUMN_ALIASES
- Normalizes that column name to OUTPUT_VRN_COLUMN ("Veh Reg No.")
- Cleans VRN values, drops blanks, de-duplicates, writes one Excel file
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

_PORTAL_ROOT = Path(__file__).resolve().parents[2]
if str(_PORTAL_ROOT) not in sys.path:
    sys.path.insert(0, str(_PORTAL_ROOT))

from header_matching import normalize_header_match

# ==========================================
# CONFIG — edit these paths / aliases
# ==========================================
CONFIG = {
    # Folder containing Concessionaire / Project Vehicle source files
    "INPUT_FOLDER": r"C:\Divyesh\NHIT_File_process\Portal\Scripts\Excempy_Query\Madai Concessionier",
    # Merged output (point Aprooved Exemption EXEMPTION_FILE here)
    "OUTPUT_FILE": r"C:\Divyesh\NHIT_File_process\Portal\Scripts\Excempy_Query\Madai Concessionier\madai_concessionaire_merged.xlsx",
    # Canonical column name expected by Aprooved Exemption CONFIG["VRN_COLUMN"]
    "OUTPUT_VRN_COLUMN": "Veh Reg No.",
    # Scan this many leading rows to find the VRN header cell
    "HEADER_SCAN_ROWS": 10,
    # Possible header names in source files (matched after strip/casefold/symbol cleanup)
    "VRN_COLUMN_ALIASES": [
        "VRN",
        "Veh Reg No.",
        "Veh Reg No",
        "Veh.Reg No.",
        "Veh.Reg No",
        "Vehicle Reg No",
        "Vehicle Reg. No.",
        "Vehicle Registration Number",
        "Vehicle No",
        "Vehicle No.",
        "Vehicle Number",
        "VehicleNumber",
        "Reg No",
        "Reg. No.",
        "Registration No",
        "Registration Number",
        "Chassis/ Vehicle No",
        "Chassis/Vehicle No",
        "Chassis Vehicle No",
        "TC_VEH_REG_NO",
        "VEH_REG_NO",
        "veh_reg_no",
        "Plate No",
        "Platenumber",
        "Licence Plate No",
        "Licence Plate No.",
    ],
    "SUPPORTED_EXTENSIONS": {".xlsx", ".xls", ".csv"},
}


def _clean_vrn(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip().upper()
    if not text or text in {"NAN", "NONE", "NULL", "NAT"}:
        return ""
    return "".join(ch for ch in text if ch.isalnum())


def _alias_key_set(aliases) -> set[str]:
    return {normalize_header_match(a) for a in aliases if normalize_header_match(a)}


def _find_vrn_header(raw: pd.DataFrame, aliases, scan_rows: int):
    """
    Scan the first scan_rows rows for a cell matching a VRN alias.

    Returns (header_row_index, vrn_column_index, header_label) or None.
    """
    keys = _alias_key_set(aliases)
    limit = min(scan_rows, len(raw))
    for row_idx in range(limit):
        row = raw.iloc[row_idx].tolist()
        for col_idx, cell in enumerate(row):
            if normalize_header_match(cell) in keys:
                return row_idx, col_idx, str(cell).strip()
    return None


def _list_input_files() -> list[Path]:
    """Return all supported files in INPUT_FOLDER (excludes the output file)."""
    folder = Path(CONFIG["INPUT_FOLDER"]).expanduser().resolve()
    if not folder.is_dir():
        raise FileNotFoundError(f"INPUT_FOLDER not found: {folder}")

    output_resolved = Path(CONFIG["OUTPUT_FILE"]).expanduser().resolve()
    files = []
    for path in sorted(folder.iterdir()):
        if not path.is_file():
            continue
        if path.suffix.lower() not in CONFIG["SUPPORTED_EXTENSIONS"]:
            continue
        if path.name.startswith("~$"):
            continue
        # Do not re-read the merged output file if it already sits in the folder
        if path.resolve() == output_resolved:
            continue
        files.append(path)
    return files


def _read_raw_table(path: Path) -> pd.DataFrame:
    """Read file with no header assumption so title rows can be scanned."""
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path, header=None, dtype=str)
    return pd.read_excel(path, header=None, dtype=str)


def extract_vrns_from_file(path: Path, aliases) -> list[str]:
    raw = _read_raw_table(path)
    if raw.empty:
        print(f"  Skip (empty): {path.name}")
        return []

    scan_rows = int(CONFIG.get("HEADER_SCAN_ROWS", 10))
    found = _find_vrn_header(raw, aliases, scan_rows)
    if not found:
        print(
            f"  Skip (no VRN header in first {scan_rows} rows): {path.name}"
        )
        return []

    header_row, vrn_col_idx, header_label = found
    series = raw.iloc[header_row + 1 :, vrn_col_idx]
    values = [_clean_vrn(v) for v in series.tolist()]
    values = [v for v in values if v]
    print(
        f"  {path.name}: header '{header_label}' at row {header_row + 1}, "
        f"col {vrn_col_idx + 1} → {len(values)} VRN value(s)"
    )
    return values


def merge_concessionaire_files() -> Path:
    aliases = CONFIG["VRN_COLUMN_ALIASES"]
    out_col = CONFIG["OUTPUT_VRN_COLUMN"]
    output_path = Path(CONFIG["OUTPUT_FILE"]).expanduser().resolve()
    folder = Path(CONFIG["INPUT_FOLDER"]).expanduser().resolve()

    files = _list_input_files()
    if not files:
        raise FileNotFoundError(
            f"No .xlsx / .xls / .csv files found in INPUT_FOLDER:\n  {folder}"
        )

    print(f"Input folder: {folder}")
    print(f"Found {len(files)} file(s) to merge:")
    all_vrns: list[str] = []
    for path in files:
        try:
            all_vrns.extend(extract_vrns_from_file(path, aliases))
        except Exception as exc:
            print(f"  Error reading {path.name}: {exc}")

    if not all_vrns:
        raise ValueError("No VRN values extracted from any input file.")

    # Preserve first-seen order while de-duplicating
    unique_vrns = list(dict.fromkeys(all_vrns))
    merged = pd.DataFrame({out_col: unique_vrns})

    output_path.parent.mkdir(parents=True, exist_ok=True)
    merged.to_excel(output_path, index=False)

    print(
        f"\nMerged {len(unique_vrns)} unique VRN(s) "
        f"(from {len(all_vrns)} raw) into:\n  {output_path}"
    )
    print(f"Output column: '{out_col}'")
    return output_path


def main():
    print("=== Concessionaire / Project Vehicles VRN Merge ===\n")
    merge_concessionaire_files()
    print("\n=== Done ===")


if __name__ == "__main__":
    main()
