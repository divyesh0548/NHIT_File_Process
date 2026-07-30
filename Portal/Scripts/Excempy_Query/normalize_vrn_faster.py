"""
Phase 1 now converts .xlsx / .xls directly into a normalized .csv in the same pass.
Phase 2 runs concurrently on all .csv files in the folder.

after that, merge_normalized_files() is the same as before.
"""

import csv
import json
import os
import sys
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, time as dt_time, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Set, Tuple

import pythoncom
from openpyxl import load_workbook
from win32com.client import DispatchEx

from merge_files import merge_files_in_folder
from vrn_normalization_config import get_normalization_groups

BASE_DIR = Path(__file__).resolve().parent
PORTAL_ROOT = BASE_DIR.parents[1]
_DEFAULT_INPUT_FOLDER = BASE_DIR / "Daroda_lc_vrn_files/vrn"
_DEFAULT_MERGE_OUTPUT_FILE = BASE_DIR / "normalized_and_merged_vrn_daroda.csv"

# Portal sets these env vars when running from Exempt Query → Merge + Normalize.
INPUT_FOLDER = Path(
    os.environ.get("MERGE_NORMALIZE_INPUT_FOLDER", str(_DEFAULT_INPUT_FOLDER))
).resolve()
MERGE_OUTPUT_FILE = Path(
    os.environ.get("MERGE_NORMALIZE_OUTPUT_FILE", str(_DEFAULT_MERGE_OUTPUT_FILE))
).resolve()

MERGE_AFTER_NORMALIZATION = True
MAX_WORKER_THREADS = 5
MIN_MERGE_HEADER_KEYWORDS = 3
XLSX_STREAM_MAX_ROWS = 1_048_576
XLSX_STREAM_TAIL_EMPTY_ROWS = 200

SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".csv"}
TRACKED_HEADER_COLUMNS = {
    "TC Class",
    "Veh Reg No.",
    "MOP",
    "Lane No",
    "Description",
    "File Name",
    "Date & Time",
}

# Resolved from Normalization_JSON_Data.json with Python defaults as fallback.
NORMALIZATION_GROUPS = get_normalization_groups()


def normalize_text(value):
    if value is None:
        return ""
    text = str(value).strip().casefold()
    for char in ("_", "-", "/", "\\", ".", ",", "(", ")", "[", "]"):
        text = text.replace(char, " ")
    return " ".join(text.split())


def build_lookup(groups):
    lookup = {}
    for common_name, aliases in groups.items():
        lookup[normalize_text(common_name)] = common_name
        for alias in aliases:
            lookup[normalize_text(alias)] = common_name
    return lookup


def update_value(value, lookup):
    replacement = lookup.get(normalize_text(value))
    if replacement and value != replacement:
        return replacement
    return None


def normalize_vrn_value(value):
    if value is None:
        return None

    text = str(value).strip().upper()
    if not text:
        return None

    normalized = "".join(ch for ch in text if ch.isalnum())
    if not normalized or normalized == text:
        return None

    return normalized


def clean_lane_value(value):
    if value is None:
        return None

    text = str(value)
    if not text:
        return None

    transformed = text
    transformed = transformed.replace("Lane ", "L0")
    transformed = transformed.replace("L010", "L10")
    transformed = transformed.replace("LN00", "L")
    transformed = transformed.replace("LN0", "L")
    transformed = transformed.replace("LN", "L")
    transformed = transformed.replace("L011", "L11")
    transformed = transformed.replace("L012", "L12")
    transformed = transformed.split("-", 1)[0]
    transformed = transformed.strip()
    transformed = transformed.replace("L0", "L")

    if transformed != text:
        return transformed
    return None


def clean_description_value(value):
    if value is None:
        return None

    text = str(value)
    if not text:
        return None

    transformed = text
    transformed = transformed.replace("()", "")
    transformed = transformed.replace("(CASH)", "")
    transformed = transformed.replace(" - DOWN", "")
    transformed = transformed.replace(" - UP", "")
    transformed = transformed.replace("ED - ", "")
    transformed = transformed.replace("EXEMPT", "")

    if transformed != text:
        return transformed
    return None


def clean_mop_value(value):
    if value is None:
        return None

    text = str(value)
    if not text:
        return None

    transformed = text.upper().split("-", 1)[0]
    if transformed != text:
        return transformed
    return None


def clean_tc_class_value(value):
    if value is None:
        return None

    text = str(value)
    if not text:
        return None

    if "MAV" in text.upper() and text != "MAV":
        return "MAV"
    return None


def _excel_serial_to_datetime(serial: float) -> datetime:
    base = datetime(1899, 12, 30)
    return base + timedelta(days=float(serial))


def parse_datetime_cell(value) -> Optional[datetime]:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date) and not isinstance(value, datetime):
        return datetime.combine(value, dt_time.min)
    if isinstance(value, (int, float)):
        try:
            return _excel_serial_to_datetime(float(value))
        except (ValueError, OSError, OverflowError):
            return None
    text = str(value).strip()
    if not text:
        return None
    fmts = (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%m/%d/%Y",
    )
    for fmt in fmts:
        try:
            return datetime.strptime(text[:26], fmt)
        except ValueError:
            continue
    return None


def coerce_datetime_value(value) -> Optional[datetime]:
    if isinstance(value, datetime):
        return None
    return parse_datetime_cell(value)


def resolve_header_name(value, normalization_lookup):
    canonical = normalization_lookup.get(normalize_text(value))
    if canonical in TRACKED_HEADER_COLUMNS:
        return canonical
    return None


def should_drop_row(row_values, header_columns):
    lane_col_idx = next((idx for idx, name in header_columns.items() if name == "Lane No"), None)
    if lane_col_idx is not None and lane_col_idx <= len(row_values):
        lane_value = row_values[lane_col_idx - 1]
        if lane_value is None or str(lane_value).strip() == "":
            return True
        if str(lane_value).strip().upper() == "EXEMPTED":
            return True

    # Keep only rows with a non-empty vehicle registration number (no format validation).
    vrn_col_idx = next((idx for idx, name in header_columns.items() if name == "Veh Reg No."), None)
    if vrn_col_idx is not None and vrn_col_idx <= len(row_values):
        vrn_value = row_values[vrn_col_idx - 1]
        if vrn_value is None or str(vrn_value).strip() == "":
            return True

    return False


def process_row_values(row_values, normalization_lookup, header_columns: Dict[int, str]):
    updates = []
    row_values_out = list(row_values)
    header_cell_indexes: Set[int] = set()

    for col_idx, value in enumerate(row_values, start=1):
        header_name = resolve_header_name(value, normalization_lookup)
        if header_name is not None:
            header_columns[col_idx] = header_name
            header_cell_indexes.add(col_idx)

        replacement = update_value(value, normalization_lookup)
        if replacement is not None:
            updates.append((col_idx, replacement))
            row_values_out[col_idx - 1] = replacement

    for col_idx, value in enumerate(row_values_out, start=1):
        if col_idx in header_cell_indexes:
            continue

        column_name = header_columns.get(col_idx)
        replacement = None

        if column_name == "Date & Time":
            replacement = coerce_datetime_value(value)
        elif column_name == "Veh Reg No.":
            replacement = normalize_vrn_value(value)
        elif column_name == "MOP":
            replacement = clean_mop_value(value)
        elif column_name == "Lane No":
            replacement = clean_lane_value(value)
        elif column_name == "Description":
            replacement = clean_description_value(value)
        elif column_name == "TC Class":
            replacement = clean_tc_class_value(value)

        if replacement is not None and replacement != value:
            updates.append((col_idx, replacement))
            row_values_out[col_idx - 1] = replacement

    drop_row = should_drop_row(row_values_out, header_columns)
    return updates, drop_row


def get_last_used_row_col(ws):
    xlFormulas = -4123
    xlByRows = 1
    xlByColumns = 2
    xlPrevious = 2

    last_row_cell = ws.Cells.Find(
        What="*", LookIn=xlFormulas, SearchOrder=xlByRows, SearchDirection=xlPrevious
    )

    last_col_cell = ws.Cells.Find(
        What="*", LookIn=xlFormulas, SearchOrder=xlByColumns, SearchDirection=xlPrevious
    )

    if last_row_cell is None or last_col_cell is None:
        return 0, 0

    return last_row_cell.Row, last_col_cell.Column


def normalize_xlsx_file(file_path: Path, normalization_lookup) -> int:
    print(f"[START][XLSX] {file_path.name}")
    workbook = load_workbook(file_path, keep_links=False)
    total_changes = 0

    try:
        for ws in workbook.worksheets:
            print(
                f"[SCAN] {file_path.name} -> {ws.title}: single-pass worksheet traversal"
            )
            sheet_changes = 0
            header_columns: Dict[int, str] = {}
            rows_to_delete: List[int] = []

            for row_idx, row in enumerate(
                ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column),
                start=1,
            ):
                row_values = [cell.value for cell in row]
                updates, drop_row = process_row_values(
                    row_values,
                    normalization_lookup,
                    header_columns,
                )

                if drop_row:
                    rows_to_delete.append(row_idx)
                    continue

                for col_idx, replacement in updates:
                    row[col_idx - 1].value = replacement
                    total_changes += 1
                    sheet_changes += 1

            for row_idx in reversed(rows_to_delete):
                ws.delete_rows(row_idx, 1)
                total_changes += 1
                sheet_changes += 1

            file_name_col = next(
                (idx for idx, name in header_columns.items() if name == "File Name"),
                None,
            )
            if file_name_col is not None:
                ws.delete_cols(file_name_col, 1)
                total_changes += 1
                sheet_changes += 1

            if sheet_changes:
                print(
                    f"[MATCH] {file_path.name} -> {ws.title}: "
                    f"{sheet_changes} cell(s) normalized"
                )
            else:
                print(f"[NO-MATCH] {file_path.name} -> {ws.title}")

        if total_changes:
            workbook.save(file_path)
            print(f"[SAVE][XLSX] {file_path.name}")
        else:
            print(f"[SKIP][XLSX] No changes: {file_path.name}")
    finally:
        workbook.close()

    return total_changes


def normalize_xls_file(file_path: Path, normalization_lookup) -> int:
    pythoncom.CoInitialize()

    excel = None
    workbook = None

    try:
        print(f"[START][XLS] {file_path.name}")

        excel = DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            excel.ScreenUpdating = False
        except Exception:
            pass

        try:
            excel.EnableEvents = False
        except Exception:
            pass

        try:
            excel.Calculation = -4135
        except Exception:
            pass

        workbook = excel.Workbooks.Open(str(file_path))

        total_changes = 0

        for ws in workbook.Worksheets:
            last_row, last_col = get_last_used_row_col(ws)

            if last_row == 0 or last_col == 0:
                print(f"[SKIP-SHEET] {file_path.name} -> {ws.Name}: empty sheet")
                continue

            print(
                f"[SCAN] {file_path.name} -> {ws.Name}: "
                f"single-pass rows 1-{last_row}, columns 1-{last_col}"
            )

            sheet_changes = 0
            header_columns: Dict[int, str] = {}
            rows_to_delete: List[int] = []

            for r_idx in range(1, last_row + 1):
                row_values = [
                    ws.Cells(r_idx, c_idx).Text for c_idx in range(1, last_col + 1)
                ]
                updates, drop_row = process_row_values(
                    row_values,
                    normalization_lookup,
                    header_columns,
                )

                if drop_row:
                    rows_to_delete.append(r_idx)
                    continue

                for c_idx, replacement in updates:
                    cell = ws.Cells(r_idx, c_idx)
                    cell.Value = replacement
                    total_changes += 1
                    sheet_changes += 1

            for r_idx in reversed(rows_to_delete):
                ws.Rows(r_idx).Delete()
                total_changes += 1
                sheet_changes += 1

            file_name_col = next(
                (idx for idx, name in header_columns.items() if name == "File Name"),
                None,
            )
            if file_name_col is not None:
                ws.Columns(file_name_col).Delete()
                total_changes += 1
                sheet_changes += 1

            if sheet_changes:
                print(
                    f"[MATCH] {file_path.name} -> {ws.Name}: "
                    f"{sheet_changes} cell(s) normalized"
                )
            else:
                print(f"[NO-MATCH] {file_path.name} -> {ws.Name}")

        if total_changes:
            workbook.Save()
            print(f"[SAVE] {file_path.name}")
        else:
            print(f"[SKIP][XLS] No changes: {file_path.name}")

        workbook.Close(SaveChanges=False)
        excel.Quit()

        return total_changes

    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        except Exception:
            pass

        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass

        pythoncom.CoUninitialize()


def detect_csv_encoding(file_path: Path):
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            with file_path.open("r", encoding=enc, newline="") as f:
                f.read(4096)
            return enc
        except UnicodeDecodeError:
            continue
    return "latin-1"


def normalize_csv_file_fast(file_path: Path, normalization_lookup):
    print(
        f"[START][CSV] {file_path.name}: detect encoding, stream normalize, write temp",
        flush=True,
    )
    encoding = detect_csv_encoding(file_path)
    print(
        f"[OPER][CSV] {file_path.name}: encoding={encoding}; "
        "normalize cells / drop rows / strip File Name column",
        flush=True,
    )
    total_changes = 0

    temp_fd, temp_path = tempfile.mkstemp(suffix=".csv")
    os.close(temp_fd)

    with (
        file_path.open("r", encoding=encoding, newline="") as src,
        open(temp_path, "w", encoding=encoding, newline="") as dst,
    ):
        reader = csv.reader(src)
        writer = csv.writer(dst)
        header_columns: Dict[int, str] = {}
        file_name_col_idx = None

        for row in reader:
            updates, drop_row = process_row_values(
                row,
                normalization_lookup,
                header_columns,
            )

            for col_index, replacement in updates:
                row[col_index - 1] = replacement
                total_changes += 1

            if drop_row:
                total_changes += 1
                continue

            if file_name_col_idx is None:
                file_name_col_idx = next(
                    (idx for idx, name in header_columns.items() if name == "File Name"),
                    None,
                )
            if file_name_col_idx is not None and file_name_col_idx <= len(row):
                del row[file_name_col_idx - 1]
                total_changes += 1

            writer.writerow(row)

    if total_changes:
        os.replace(temp_path, file_path)
        print(
            f"[SAVE][CSV] {file_path.name}: replace file in place (edits applied)",
            flush=True,
        )
    else:
        os.remove(temp_path)
        print(
            f"[SKIP][CSV] {file_path.name}: discard temp (no edits)",
            flush=True,
        )

    return total_changes


def get_merge_header_keywords() -> List[str]:
    raw = os.environ.get("MERGE_NORMALIZE_HEADER_KEYWORDS", "").strip()
    if raw:
        try:
            parsed = json.loads(raw)
            if isinstance(parsed, list):
                keywords = [str(k).strip() for k in parsed if str(k).strip()]
                if len(keywords) >= MIN_MERGE_HEADER_KEYWORDS:
                    return keywords
        except json.JSONDecodeError:
            pass

    portal_root = str(PORTAL_ROOT)
    if portal_root not in sys.path:
        sys.path.insert(0, portal_root)

    from db.nhit_file_process import get_lc_etc_header_keyword_strings

    keywords = get_lc_etc_header_keyword_strings()
    if len(keywords) < MIN_MERGE_HEADER_KEYWORDS:
        raise RuntimeError(
            f"Need at least {MIN_MERGE_HEADER_KEYWORDS} LC/ETC/VRN header keywords in "
            f"nhit_file_process; found {len(keywords)}. "
            "Configure them from Merge + Normalize -> Header Keywords."
        )
    return keywords


def find_input_files(folder_path: Path):
    return sorted(
        p
        for p in folder_path.iterdir()
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS
    )


def _normalize_and_filter_row_for_csv(
    row_values,
    normalization_lookup,
    header_columns: Dict[int, str],
     ):
    updates, drop_row = process_row_values(
        row_values,
        normalization_lookup,
        header_columns,
    )
    row_out = list(row_values)
    changes = 0

    for col_idx, replacement in updates:
        if col_idx <= len(row_out):
            row_out[col_idx - 1] = replacement
            changes += 1

    if drop_row:
        return None, changes + 1

    file_name_col_idx = next(
        (idx for idx, name in header_columns.items() if name == "File Name"),
        None,
    )
    if file_name_col_idx is not None and file_name_col_idx <= len(row_out):
        del row_out[file_name_col_idx - 1]
        changes += 1

    row_out = ["" if value is None else value for value in row_out]
    return row_out, changes


def _append_normalized_xlsx_rows(
    writer,
    ws,
    first_row_inclusive: int,
    out_cols: int,
    normalization_lookup,
    header_columns: Dict[int, str],
    sheet_label: str = "",
    ) -> Tuple[int, int]:
    rows_written = 0
    total_changes = 0
    tail_empty = 0
    scan_last = min(first_row_inclusive + XLSX_STREAM_MAX_ROWS - 1, 1_048_576)

    for r_idx, row in enumerate(
        ws.iter_rows(
            min_row=1,
            max_row=scan_last,
            min_col=1,
            max_col=out_cols,
            values_only=True,
        ),
        start=1,
    ):
        if r_idx < first_row_inclusive:
            continue

        row_values = list(row)[:out_cols]
        while len(row_values) < out_cols:
            row_values.append(None)

        has_data = any(normalize_text(c) != "" for c in row_values)
        if not has_data:
            tail_empty += 1
            if tail_empty >= XLSX_STREAM_TAIL_EMPTY_ROWS:
                break
        else:
            tail_empty = 0

        row_out, row_changes = _normalize_and_filter_row_for_csv(
            row_values,
            normalization_lookup,
            header_columns,
        )
        total_changes += row_changes

        if row_out is None:
            continue

        writer.writerow(row_out)
        rows_written += 1

    if sheet_label:
        print(
            f"[XLSX][WRITE] {sheet_label}: wrote {rows_written} normalized row(s) "
            f"from row {first_row_inclusive} onward",
            flush=True,
        )

    return rows_written, total_changes


def _detect_xls_header(
    path: Path,
    keywords: Sequence[str],
    min_keyword_matches: Optional[int] = None,
     ) -> Tuple[int, int]:
    from csv_converter import detect_header_xls

    return detect_header_xls(path, keywords, min_keyword_matches)


def convert_xlsx_to_csv_with_normalization(
    path: Path,
    out_path: Path,
    keywords: Sequence[str],
    normalization_lookup,
    min_keyword_matches: Optional[int] = None,
    ) -> Tuple[int, int]:
    from csv_converter import HEADER_SCAN_MAX_ROW, detect_header_row_on_worksheet

    path = path.resolve()
    out_path = Path(out_path).resolve()
    min_matches = (
        min_keyword_matches
        if min_keyword_matches is not None
        else MIN_MERGE_HEADER_KEYWORDS
    )

    from csv_converter import WorkbookHeaderNotFoundError

    sheet_specs: List[Tuple[int, int, int]] = []
    skipped_sheets: List[str] = []
    wb1 = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        for si, ws in enumerate(wb1.worksheets):
            sheet_name = getattr(ws, "title", f"Sheet{si}")
            print(
                f"[XLSX][SCAN] sheet {si} ({sheet_name}): "
                f"checking rows 1-{HEADER_SCAN_MAX_ROW} for header..."
            )
            det = detect_header_row_on_worksheet(ws, keywords, min_matches)
            if det is None:
                skipped_sheets.append(sheet_name)
                print(
                    f"[XLSX][SKIP] sheet {si} ({sheet_name}): no row in 1-"
                    f"{HEADER_SCAN_MAX_ROW} matched at least {min_matches} "
                    "header keyword(s); skipping sheet.",
                    flush=True,
                )
                continue
            header_row, header_cols, score = det
            print(
                f"[XLSX][HEADER] sheet {si} ({sheet_name}): "
                f"row {header_row}, score={score}, cols={header_cols}"
            )
            sheet_specs.append((si, header_row, header_cols))
    finally:
        wb1.close()

    if not sheet_specs:
        skipped = ", ".join(repr(s) for s in skipped_sheets) or "(none)"
        raise WorkbookHeaderNotFoundError(
            f"{path.name}: no sheet matched at least {min_matches} header "
            f"keyword(s) in rows 1-{HEADER_SCAN_MAX_ROW}. "
            f"Skipped sheet(s): {skipped}. "
            "Adjust header keywords or sheet layout."
        )

    out_cols = max(p[2] for p in sheet_specs)
    rows_written = 0
    total_changes = 0
    header_columns: Dict[int, str] = {}

    wb2 = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        with out_path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            first_sheet = True
            for si, header_row, _header_cols in sheet_specs:
                ws = wb2.worksheets[si]
                sheet_name = getattr(ws, "title", f"Sheet{si}")
                start = header_row if first_sheet else header_row + 1
                mode = "header+data" if first_sheet else "data-only"
                first_sheet = False
                print(
                    f"[XLSX][START] sheet {si} ({sheet_name}): "
                    f"{mode}, starting row {start}"
                )
                written, changes = _append_normalized_xlsx_rows(
                    writer,
                    ws,
                    start,
                    out_cols,
                    normalization_lookup,
                    header_columns,
                    f"sheet {si} ({sheet_name})",
                )
                rows_written += written
                total_changes += changes
    finally:
        wb2.close()

    print(
        f"[XLSX][DONE] {path.name}: total normalized rows written={rows_written}",
        flush=True,
    )
    return rows_written, total_changes


def convert_xls_to_csv_with_normalization(
    path: Path,
    out_path: Path,
    keywords: Sequence[str],
    normalization_lookup,
    min_keyword_matches: Optional[int] = None,
    ) -> Tuple[int, int]:
    header_row, header_cols = _detect_xls_header(path, keywords, min_keyword_matches)

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    rows_written = 0
    total_changes = 0
    try:
        excel = DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(path))
        ws = workbook.Worksheets(1)
        last_row, last_col = get_last_used_row_col(ws)
        end_row = last_row if last_row else header_row
        header_columns: Dict[int, str] = {}

        with out_path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            for r_idx in range(header_row, end_row + 1):
                raw = [
                    ws.Cells(r_idx, c_idx).Text
                    for c_idx in range(1, max(header_cols, last_col, 1) + 1)
                ]
                raw = raw[:header_cols]
                while len(raw) < header_cols:
                    raw.append("")

                row_out, row_changes = _normalize_and_filter_row_for_csv(
                    raw,
                    normalization_lookup,
                    header_columns,
                )
                total_changes += row_changes

                if row_out is None:
                    continue

                writer.writerow(row_out)
                rows_written += 1
    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    return rows_written, total_changes


def convert_workbook_to_normalized_csv(
    path: Path,
    keywords: Sequence[str],
    normalization_lookup,
    out_path: Optional[Path] = None,
    min_keyword_matches: Optional[int] = None,
    ) -> Tuple[Path, int, int]:
    path = path.resolve()
    if not path.is_file():
        raise FileNotFoundError(path)

    suffix = path.suffix.lower()
    if suffix not in (".xlsx", ".xls"):
        raise ValueError(f"Expected .xlsx or .xls, got {suffix}")

    if out_path is None:
        out_path = path.with_suffix(".csv")
    else:
        out_path = Path(out_path).resolve()

    if suffix == ".xlsx":
        rows_written, total_changes = convert_xlsx_to_csv_with_normalization(
            path,
            out_path,
            keywords,
            normalization_lookup,
            min_keyword_matches,
        )
    else:
        rows_written, total_changes = convert_xls_to_csv_with_normalization(
            path,
            out_path,
            keywords,
            normalization_lookup,
            min_keyword_matches,
        )

    return out_path, rows_written, total_changes


def phase_convert_workbooks_to_csv_and_delete(
    folder_path: Path,
    normalization_lookup,
     ) -> Set[Path]:
    converted_csvs: Set[Path] = set()

    try:
        from csv_converter import (
            HEADER_SCAN_MAX_ROW as _csv_scan_max_row,
            WorkbookHeaderNotFoundError,
        )
    except ImportError as exc:
        print(
            f"[WARN] Could not import csv_converter ({exc}); "
            "skipping pre-conversion to CSV.",
            flush=True,
        )
        return converted_csvs

    try:
        keywords = get_merge_header_keywords()
    except Exception as exc:
        print(
            f"[WARN] Could not load LC/ETC/VRN header keywords ({exc}); "
            "skipping pre-conversion to CSV.",
            flush=True,
        )
        return converted_csvs

    conv_min_matches = min(MIN_MERGE_HEADER_KEYWORDS, len(keywords))

    folder_path = folder_path.resolve()
    if not folder_path.is_dir():
        return converted_csvs

    targets = [
        p
        for p in sorted(folder_path.iterdir(), key=lambda x: x.name.lower())
        if p.is_file()
        and p.suffix.lower() in (".xlsx", ".xls")
        and not p.name.startswith("~$")
    ]
    if not targets:
        print("[PHASE 1] No .xlsx/.xls files to pre-convert.", flush=True)
        return converted_csvs

    print(
        f"\n[PHASE 1] workbook(s) -> normalized .csv in one pass "
        f"(header scan rows 1-{_csv_scan_max_row}, min matches={conv_min_matches}; "
        f"delete source on success)...",
        flush=True,
    )
    for path in targets:
        out_csv = path.with_suffix(".csv")
        try:
            print(
                f"[PHASE 1][START] {path.name}: detecting header + converting + normalizing...",
                flush=True,
            )
            _, rows_written, total_changes = convert_workbook_to_normalized_csv(
                path,
                keywords,
                normalization_lookup,
                out_path=out_csv,
                min_keyword_matches=conv_min_matches,
            )
            path.unlink()
            converted_csvs.add(out_csv.resolve())
            print(
                f"[PHASE 1][DONE] {path.name}: converted to {out_csv.name} "
                f"({rows_written} row(s), {total_changes} change(s))",
                flush=True,
            )
            print(f"[CONVERT+DEL] {path.name} -> {out_csv.name}", flush=True)
        except WorkbookHeaderNotFoundError as exc:
            print(f"[CONVERT-FAIL] {path.name}: {exc}", flush=True)
            raise SystemExit(f"Header detection failed for {path.name}: {exc}") from exc
        except Exception as exc:
            print(
                f"[CONVERT-SKIP] {path.name}: {exc} "
                f"(will use default .xlsx/.xls normalization)",
                flush=True,
            )

    return converted_csvs


def merge_normalized_files():
    print(f"\n[MERGE] Starting merge from {INPUT_FOLDER}")
    header_keywords = get_merge_header_keywords()
    print(
        f"[MERGE] Using {len(header_keywords)} LC/ETC/VRN header keyword(s) from database.",
        flush=True,
    )
    MERGE_OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    merge_files_in_folder(
        str(INPUT_FOLDER),
        str(MERGE_OUTPUT_FILE),
        header_keywords,
    )


def process_file(file_path: Path, normalization_lookup):
    suffix = file_path.suffix.lower()

    if suffix == ".xlsx":
        total = normalize_xlsx_file(file_path, normalization_lookup)
    elif suffix == ".xls":
        total = normalize_xls_file(file_path, normalization_lookup)
    elif suffix == ".csv":
        total = normalize_csv_file_fast(file_path, normalization_lookup)
    else:
        return None

    return file_path.name, total


def main():
    start_time = time.perf_counter()
    # Reload JSON each run so portal edits apply without restarting the process script.
    normalization_lookup = build_lookup(get_normalization_groups())

    INPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    MERGE_OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    print(f"Input folder: {INPUT_FOLDER}", flush=True)
    print(f"Merge output: {MERGE_OUTPUT_FILE}", flush=True)

    converted_csvs = phase_convert_workbooks_to_csv_and_delete(
        INPUT_FOLDER,
        normalization_lookup,
    )

    files = [
        p for p in find_input_files(INPUT_FOLDER)
        if p.resolve() not in converted_csvs
    ]

    print(f"\n[PHASE 2] Found {len(files)} file(s) to normalize")

    worker_count = min(MAX_WORKER_THREADS, len(files)) if files else 0
    print(f"Using up to {worker_count} concurrent worker(s)")

    with ThreadPoolExecutor(max_workers=MAX_WORKER_THREADS) as executor:
        future_to_file = {
            executor.submit(process_file, file_path, normalization_lookup): file_path
            for file_path in files
        }

        for future in as_completed(future_to_file):
            file_path = future_to_file[future]

            try:
                result = future.result()
                if result is None:
                    continue

                file_name, total = result
                print(f"\nFile: {file_name}")
                print(f"Total cells changed: {total}")

            except Exception as e:
                print(f"\nError in {file_path.name}: {e}")

    if MERGE_AFTER_NORMALIZATION:
        merge_normalized_files()
    else:
        print("\n[MERGE] Skipped because MERGE_AFTER_NORMALIZATION is False")

    elapsed_seconds = time.perf_counter() - start_time
    print(f"[TIME] Total time taken: {elapsed_seconds:.2f} seconds")


if __name__ == "__main__":
    main()
