"""Split dynamically supplied pass files into MP, LT, and Other workbooks.

The portal supplies ``EXEMPT_FINAL_PASS_INPUT`` and
``EXEMPT_FINAL_PASS_OUTPUT`` for each run.  No user-specific location is
embedded in this script, so concurrent process folders never share inputs.
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import pandas as pd


HEADER_SCAN_ROWS = 50
MIN_HEADER_MATCHES = 3
HEADER_KEYWORDS = [
    "Chassis/ Vehicle No",
    "NPCI Vehicle Class",
    "Start Date",
    "Plaza Name",
    "End Date",
    "Mobile No",
    "Pass Type",
    "Payment Mode",
]
PASS_TYPE_COLUMN_NAMES = ("Pass Type", "PassType", "PASS_TYPE")
SUPPORTED_EXTENSIONS = {".xls", ".xlsx"}
SUMMARY_FILENAME = "pass_separator_summary.json"


def _normalise(value) -> str:
    return "".join(str(value).strip().casefold().split())


def _find_header_row(sample: pd.DataFrame) -> int:
    expected = {_normalise(value) for value in HEADER_KEYWORDS}
    best_index = 0
    best_matches = 0
    for index in range(len(sample)):
        values = {_normalise(value) for value in sample.iloc[index].tolist() if pd.notna(value)}
        matches = len(values & expected)
        if matches > best_matches:
            best_index, best_matches = index, matches
        if matches >= MIN_HEADER_MATCHES:
            return index
    return best_index if best_matches >= MIN_HEADER_MATCHES else 0


def _read_sheet(excel_file: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    sample = pd.read_excel(
        excel_file,
        sheet_name=sheet_name,
        header=None,
        dtype=str,
        nrows=HEADER_SCAN_ROWS,
    )
    if sample.empty:
        return pd.DataFrame()
    header_row = _find_header_row(sample)
    return pd.read_excel(excel_file, sheet_name=sheet_name, skiprows=header_row, header=0)


def _find_pass_type_column(columns) -> str | None:
    candidates = {_normalise(value) for value in PASS_TYPE_COLUMN_NAMES}
    for column in columns:
        if _normalise(column) in candidates:
            return column
    return None


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


def separate_pass_files(input_folder: Path, output_folder: Path) -> dict:
    input_folder = Path(input_folder).resolve()
    output_folder = Path(output_folder).resolve()
    if not input_folder.is_dir():
        raise FileNotFoundError(f"Pass input folder does not exist: {input_folder}")

    source_files = sorted(
        file for file in input_folder.iterdir()
        if file.is_file() and file.suffix.lower() in SUPPORTED_EXTENSIONS
    )
    if not source_files:
        raise ValueError("No .xls or .xlsx pass files were supplied.")

    mp_frames: list[pd.DataFrame] = []
    lt_frames: list[pd.DataFrame] = []
    other_frames: list[pd.DataFrame] = []
    warnings: list[str] = []
    pass_type_detected = False
    processed_sheets = 0

    for source_file in source_files:
        try:
            excel_file = pd.ExcelFile(source_file)
        except Exception as exc:
            warnings.append(f"Could not open {source_file.name}: {exc}")
            continue
        for sheet_name in excel_file.sheet_names:
            try:
                frame = _read_sheet(excel_file, sheet_name)
            except Exception as exc:
                warnings.append(f"Could not read {source_file.name} ({sheet_name}): {exc}")
                continue
            if frame.empty:
                continue
            processed_sheets += 1
            pass_column = _find_pass_type_column(frame.columns)
            if pass_column is None:
                warnings.append(
                    f"No Pass Type column in {source_file.name} ({sheet_name}); sheet skipped."
                )
                continue
            pass_type_detected = True
            pass_values = frame[pass_column].fillna("").astype(str).str.strip().str.upper()
            mp_part = frame.loc[pass_values == "MP"].copy()
            lt_part = frame.loc[pass_values == "LT"].copy()
            other_part = frame.loc[(pass_values != "") & ~pass_values.isin(["MP", "LT"])].copy()
            if not mp_part.empty:
                mp_frames.append(mp_part)
            if not lt_part.empty:
                lt_frames.append(lt_part)
            if not other_part.empty:
                other_frames.append(other_part)

    mp_output = _merge_frames(mp_frames)
    lt_output = _merge_frames(lt_frames)
    other_output = _merge_frames(other_frames)
    _write(mp_output, output_folder / "MP_Pass.xlsx")
    _write(lt_output, output_folder / "LT_Pass.xlsx")
    if not other_output.empty:
        _write(other_output, output_folder / "Other_Pass.xlsx")

    summary = {
        "source_file_count": len(source_files),
        "processed_sheet_count": processed_sheets,
        "pass_type_detected": pass_type_detected,
        "mp_rows": int(len(mp_output)),
        "lt_rows": int(len(lt_output)),
        "other_rows": int(len(other_output)),
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
