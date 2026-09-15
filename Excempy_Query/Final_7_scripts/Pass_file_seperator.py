"""Split dynamically supplied pass files into MP, LT, and Other workbooks.

The portal supplies ``EXEMPT_FINAL_PASS_INPUT`` and
``EXEMPT_FINAL_PASS_OUTPUT`` for each run.  No user-specific location is
embedded in this script, so concurrent process folders never share inputs.

Flow per sheet / CSV:
1. Scan the first HEADER_SCAN_ROWS rows to detect the header.
2. Only after a header is found, look for a Pass Type column in it.
"""

from __future__ import annotations

import csv
import json
import os
import sys
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd

from config import (
    lt_pass_type_normalized_set,
    normalize_pass_type_value,
)
from pass_config import load_config_values as _load_pass_config

_PORTAL_ROOT = Path(__file__).resolve().parents[3]
if str(_PORTAL_ROOT) not in sys.path:
    sys.path.insert(0, str(_PORTAL_ROOT))

from header_matching import normalize_header_match

SUPPORTED_EXTENSIONS = {".xls", ".xlsx", ".csv"}
SUMMARY_FILENAME = "pass_separator_summary.json"


def _load_runtime_config():
    """Load pass-separator settings from JSON (with Python defaults fallback)."""
    cfg, _from_file = _load_pass_config()
    return cfg


def _normalise(value) -> str:
    return normalize_header_match(value)


def _find_header_row(
    sample: pd.DataFrame,
    header_keywords: list,
    min_header_matches: int,
) -> Optional[Tuple[int, int]]:
    """
    Find the best header row in the scanned sample.

    Returns (row_index, keyword_matches) or None when no row reaches
    min_header_matches. Never falls back to row 0 on failure.
    """
    expected = {_normalise(value) for value in header_keywords}
    best_index: Optional[int] = None
    best_matches = 0
    for index in range(len(sample)):
        values = {
            _normalise(value)
            for value in sample.iloc[index].tolist()
            if pd.notna(value) and str(value).strip()
        }
        matches = len(values & expected)
        if matches > best_matches:
            best_index, best_matches = index, matches
        if matches >= min_header_matches:
            return index, matches
    if best_index is not None and best_matches >= min_header_matches:
        return best_index, best_matches
    return None


def _find_pass_type_column(columns, pass_type_column_names: list) -> str | None:
    candidates = {_normalise(value) for value in pass_type_column_names}
    for column in columns:
        if _normalise(column) in candidates:
            return column
    return None


def _frame_from_raw(
    raw: pd.DataFrame,
    header_keywords: list,
    pass_type_column_names: list,
    header_scan_rows: int,
    min_header_matches: int,
) -> Tuple[str, Optional[pd.DataFrame], Optional[str], Optional[int]]:
    """
    Detect header in a header=None raw frame, then return typed status.

    Returns (status, frame, detail, header_row):
      - ("empty", None, reason, None)
      - ("no_header", None, reason, None)
      - ("no_pass_type", frame, reason, header_row)
      - ("ok", frame, pass_column_name, header_row)
    """
    if raw is None or raw.empty:
        return "empty", None, "sheet is empty", None

    sample = raw.iloc[:header_scan_rows]
    found = _find_header_row(sample, header_keywords, min_header_matches)
    if found is None:
        return (
            "no_header",
            None,
            (
                f"header not detected in first {header_scan_rows} rows "
                f"(need at least {min_header_matches} of: "
                f"{', '.join(header_keywords)})"
            ),
            None,
        )

    header_row, _score = found
    headers = [
        str(v).strip() if pd.notna(v) and str(v).strip() else f"Unnamed_{idx}"
        for idx, v in enumerate(raw.iloc[header_row].tolist())
    ]
    frame = raw.iloc[header_row + 1 :].copy()
    frame.columns = headers
    frame = frame.reset_index(drop=True)
    if frame.empty:
        return "empty", None, "no data rows under detected header", header_row

    pass_column = _find_pass_type_column(frame.columns, pass_type_column_names)
    if pass_column is None:
        available = ", ".join(str(c).strip() for c in frame.columns if str(c).strip())
        return (
            "no_pass_type",
            frame,
            (
                f"header found at row {header_row + 1}, but no Pass Type column "
                f"(tried: {', '.join(pass_type_column_names)}; "
                f"columns: {available or '(none)'})"
            ),
            header_row,
        )

    return "ok", frame, pass_column, header_row


def _read_csv_flexible(file_path: Path) -> pd.DataFrame:
    """Read CSV allowing uneven field counts (common in bank portal exports)."""
    rows = []
    with open(file_path, newline="", encoding="utf-8", errors="replace") as handle:
        for row in csv.reader(handle):
            rows.append(row)
    if not rows:
        return pd.DataFrame()
    width = max(len(row) for row in rows)
    padded = [row + [""] * (width - len(row)) for row in rows]
    return pd.DataFrame(padded, dtype=str)


def _read_sheet_with_header(
    excel_file: pd.ExcelFile,
    sheet_name: str,
    header_keywords: list,
    pass_type_column_names: list,
    header_scan_rows: int,
    min_header_matches: int,
) -> Tuple[str, Optional[pd.DataFrame], Optional[str], Optional[int]]:
    raw = pd.read_excel(
        excel_file,
        sheet_name=sheet_name,
        header=None,
        dtype=str,
    )
    return _frame_from_raw(
        raw,
        header_keywords,
        pass_type_column_names,
        header_scan_rows,
        min_header_matches,
    )


def _read_csv_with_header(
    file_path: Path,
    header_keywords: list,
    pass_type_column_names: list,
    header_scan_rows: int,
    min_header_matches: int,
) -> Tuple[str, Optional[pd.DataFrame], Optional[str], Optional[int]]:
    raw = _read_csv_flexible(file_path)
    return _frame_from_raw(
        raw,
        header_keywords,
        pass_type_column_names,
        header_scan_rows,
        min_header_matches,
    )


def _normalise_frame_columns(frame: pd.DataFrame) -> tuple[pd.DataFrame, list[str], dict[str, str]]:
    order: list[str] = []
    display: dict[str, str] = {}
    rename: dict[object, str] = {}
    for column in frame.columns:
        normalized = _normalise(column)
        if not normalized:
            continue
        rename[column] = normalized
        if normalized not in display:
            order.append(normalized)
            display[normalized] = str(column).strip()
    normalized_frame = frame.rename(columns=rename)
    normalized_frame = normalized_frame.loc[:, ~normalized_frame.columns.duplicated()].copy()
    return normalized_frame, order, display


def _merge_frames(frames: list[pd.DataFrame]) -> pd.DataFrame:
    if not frames:
        return pd.DataFrame()
    canonical_order: list[str] = []
    display_names: dict[str, str] = {}
    normalized_frames: list[pd.DataFrame] = []
    for frame in frames:
        normalized, order, display = _normalise_frame_columns(frame)
        for column in order:
            if column not in display_names:
                canonical_order.append(column)
                display_names[column] = display[column]
        normalized_frames.append(normalized)
    merged = pd.concat(
        [frame.reindex(columns=canonical_order) for frame in normalized_frames],
        ignore_index=True,
        sort=False,
    )
    merged.columns = [display_names[column] for column in canonical_order]
    return merged


def _write(frame: pd.DataFrame, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    frame.to_excel(output_path, index=False, engine="openpyxl")


def _split_by_pass_type(
    frame: pd.DataFrame,
    pass_column: str,
    lt_aliases: set,
    mp_frames: list,
    lt_frames: list,
) -> None:
    pass_values = frame[pass_column].map(normalize_pass_type_value)
    lt_mask = pass_values.isin(lt_aliases)
    mp_mask = (pass_values != "") & ~lt_mask
    mp_part = frame.loc[mp_mask].copy()
    lt_part = frame.loc[lt_mask].copy()
    if not mp_part.empty:
        mp_frames.append(mp_part)
    if not lt_part.empty:
        lt_frames.append(lt_part)


def separate_pass_files(input_folder: Path, output_folder: Path) -> dict:
    input_folder = Path(input_folder).resolve()
    output_folder = Path(output_folder).resolve()
    if not input_folder.is_dir():
        raise FileNotFoundError(f"Pass input folder does not exist: {input_folder}")

    cfg = _load_runtime_config()
    header_keywords = cfg["HEADER_KEYWORDS"]
    pass_type_column_names = cfg["PASS_TYPE_COLUMN_NAMES"]
    header_scan_rows = cfg["HEADER_SCAN_ROWS"]
    min_header_matches = cfg["MIN_HEADER_MATCHES"]

    source_files = sorted(
        file for file in input_folder.iterdir()
        if file.is_file() and file.suffix.lower() in SUPPORTED_EXTENSIONS
    )
    if not source_files:
        raise ValueError("No .xls, .xlsx, or .csv pass files were supplied.")

    mp_frames: list[pd.DataFrame] = []
    lt_frames: list[pd.DataFrame] = []
    warnings: list[str] = []
    header_detected = False
    pass_type_detected = False
    processed_sheets = 0
    sheets_no_header = 0
    sheets_no_pass_type = 0
    lt_aliases = lt_pass_type_normalized_set()

    for source_file in source_files:
        suffix = source_file.suffix.lower()

        if suffix == ".csv":
            try:
                status, frame, detail, _header_row = _read_csv_with_header(
                    source_file,
                    header_keywords,
                    pass_type_column_names,
                    header_scan_rows,
                    min_header_matches,
                )
            except Exception as exc:
                warnings.append(f"Could not read {source_file.name}: {exc}")
                continue

            label = source_file.name
            if status == "empty":
                warnings.append(f"Skipped empty file {label}: {detail}")
                continue
            if status == "no_header":
                sheets_no_header += 1
                warnings.append(f"Header not detected in {label}: {detail}")
                continue

            header_detected = True
            processed_sheets += 1

            if status == "no_pass_type":
                sheets_no_pass_type += 1
                warnings.append(f"No Pass Type column in {label}: {detail}")
                continue

            pass_column = detail
            assert frame is not None and pass_column is not None
            pass_type_detected = True
            _split_by_pass_type(frame, pass_column, lt_aliases, mp_frames, lt_frames)
            continue

        try:
            excel_file = pd.ExcelFile(source_file)
        except Exception as exc:
            warnings.append(f"Could not open {source_file.name}: {exc}")
            continue
        for sheet_name in excel_file.sheet_names:
            try:
                status, frame, detail, _header_row = _read_sheet_with_header(
                    excel_file, sheet_name,
                    header_keywords, pass_type_column_names,
                    header_scan_rows, min_header_matches,
                )
            except Exception as exc:
                warnings.append(f"Could not read {source_file.name} ({sheet_name}): {exc}")
                continue

            label = f"{source_file.name} ({sheet_name})"
            if status == "empty":
                warnings.append(f"Skipped empty sheet {label}: {detail}")
                continue
            if status == "no_header":
                sheets_no_header += 1
                warnings.append(f"Header not detected in {label}: {detail}")
                continue

            header_detected = True
            processed_sheets += 1

            if status == "no_pass_type":
                sheets_no_pass_type += 1
                warnings.append(f"No Pass Type column in {label}: {detail}")
                continue

            pass_column = detail
            assert frame is not None and pass_column is not None
            pass_type_detected = True
            _split_by_pass_type(frame, pass_column, lt_aliases, mp_frames, lt_frames)

    mp_output = _merge_frames(mp_frames)
    lt_output = _merge_frames(lt_frames)
    _write(mp_output, output_folder / "MP_Pass.xlsx")
    _write(lt_output, output_folder / "LT_Pass.xlsx")

    summary = {
        "source_file_count": len(source_files),
        "processed_sheet_count": processed_sheets,
        "header_detected": header_detected,
        "pass_type_detected": pass_type_detected,
        "sheets_no_header": sheets_no_header,
        "sheets_no_pass_type": sheets_no_pass_type,
        "header_scan_rows": header_scan_rows,
        "min_header_matches": min_header_matches,
        "header_keywords": list(header_keywords),
        "mp_aliases": cfg["MP_PASS_TYPE_VALUES"],
        "lt_aliases": cfg["LT_PASS_TYPE_VALUES"],
        "mp_rows": int(len(mp_output)),
        "lt_rows": int(len(lt_output)),
        "warnings": warnings,
    }
    (output_folder / SUMMARY_FILENAME).write_text(json.dumps(summary, indent=2), encoding="utf-8")
    return summary


def main() -> int:
    input_folder = os.environ.get("EXEMPT_FINAL_PASS_INPUT", "").strip()
    output_folder = os.environ.get("EXEMPT_FINAL_PASS_OUTPUT", "").strip()
    if not input_folder or not output_folder:
        print(
            "EXEMPT_FINAL_PASS_INPUT and EXEMPT_FINAL_PASS_OUTPUT are required.",
            file=sys.stderr,
        )
        return 2
    try:
        summary = separate_pass_files(Path(input_folder), Path(output_folder))
    except Exception as exc:
        print(f"Pass separation failed: {exc}", file=sys.stderr)
        return 1
    print(f"PASS_SEPARATOR_SUMMARY={json.dumps(summary)}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
