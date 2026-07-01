"""
Convert .xlsx / .xls to .csv with header detection and row trimming.

- For .xlsx export, every worksheet is scanned: rows 1–50 on each sheet (read-only
  workbooks often report max_column=1, so each row uses HEADER_SCAN_MAX_COL). On
  each sheet the best-scoring header row wins (ties: earliest row on that sheet).
  All sheets are appended into one CSV: the header row is written once (from the
  first sheet), then that sheet's data rows, then each following sheet's data
  only (its header row skipped). Output column width is the max width found across
  sheets. Cells are matched to HEADER_KEYWORDS using norm_key (strip, casefold,
  all whitespace removed).
- All rows above each sheet's chosen header are omitted for that sheet.
- Every output row is truncated (or padded) to the header row's column count;
  any values beyond that width are dropped.

Original workbooks are never modified or deleted. For each .xlsx / .xls in the
input folder, a new .csv is written alongside it (same stem). Files that are
already .csv are skipped. Processing runs one file at a time in sorted order.
"""

from __future__ import annotations

import csv
import sys
import time
from datetime import date, datetime
from pathlib import Path
from typing import Any, List, Optional, Sequence, Tuple

from openpyxl import load_workbook

try:
    import pythoncom
    from win32com.client import DispatchEx

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# Folder containing workbooks to convert (edit path). No CLI.
INPUT_FOLDER = Path(__file__).resolve().parent / "vrn"

CONVERT_EXTENSIONS = {".xlsx", ".xls"}

# Header keywords: each string is normalized like cells (strip, casefold, no whitespace).
# A row qualifies only if it matches at least MIN_HEADER_KEYWORD_MATCHES keywords.
HEADER_KEYWORDS: List[str] = [
    "VEH REG NO",
    "CCH TXN NO",
    "AVC",
    "LANE",
    "MVC MOP",
    "Date & Time",
    "Veh Reg No.",
    "MOP",
    "Lane No",
    "TC Class",
    "Settlement Type",
    "Agency Txn Id",
    "Plaza ID",
    "Violation Flag",
    "Txn ID",
    "Settlement Amount",
    "Violation Amt",
    "Journey Type"
]

HEADER_SCAN_MAX_ROW = 50
# read_only xlsx often has max_column=1; always scan at least this many columns.
HEADER_SCAN_MAX_COL = 512
MIN_HEADER_KEYWORD_MATCHES = 5
# read_only: random high min_row re-scans from row 1 each time (quadratic). Stream in one
# forward pass from row 1 with a high max_row cap; skip rows before the detected header.
XLSX_STREAM_MAX_ROWS = 10_000_000
XLSX_STREAM_TAIL_EMPTY_ROWS = 100


def norm_key(value) -> str:
    """
    Normalize for header matching: strip leading/trailing whitespace, casefold,
    then remove every Unicode whitespace character (including spaces between
    words and line breaks from pasted Excel text). Same rule applies to
    HEADER_KEYWORDS when scored.
    """
    if value is None:
        return ""
    s = str(value).strip().casefold()
    return "".join(ch for ch in s if not ch.isspace())


def score_header_row(cell_values: Sequence, keywords: Sequence[str]) -> int:
    if not keywords:
        return 0
    cell_norms = {norm_key(c) for c in cell_values if c is not None and norm_key(c)}
    return sum(1 for kw in keywords if norm_key(kw) in cell_norms)


def trim_trailing_empty(values: Sequence) -> List:
    lst = [v for v in values]
    while lst and (lst[-1] is None or norm_key(lst[-1]) == ""):
        lst.pop()
    return lst if lst else [""]


def cell_to_csv(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


def _header_row_better(
    sheet_idx: int,
    row_idx: int,
    score: int,
    best_sheet: Optional[int],
    best_row: Optional[int],
    best_score: int,
) -> bool:
    """
    Prefer higher score. On tie: prefer an earlier header row (row 1 beats row 8
    when a summary sheet repeats headers lower down), then an earlier sheet index.
    """
    if best_row is None:
        return True
    if score > best_score:
        return True
    if score < best_score:
        return False
    if row_idx < best_row:
        return True
    if row_idx > best_row:
        return False
    if best_sheet is None:
        return True
    return sheet_idx < best_sheet


def detect_header_row_on_worksheet(
    ws, keywords: Sequence[str]
) -> Optional[Tuple[int, int, int]]:
    """
    Find the best header row on one worksheet (rows 1–HEADER_SCAN_MAX_ROW only).

    Returns (header_row_1based, header_col_count, keyword_score) or None if no row qualifies.
    """
    best_row: Optional[int] = None
    best_score = -1
    best_width = 1

    for r_idx, row in enumerate(
        ws.iter_rows(
            min_row=1,
            max_row=HEADER_SCAN_MAX_ROW,
            min_col=1,
            max_col=HEADER_SCAN_MAX_COL,
            values_only=True,
        ),
        start=1,
    ):
        vals = list(row)
        score = score_header_row(vals, keywords)
        if score < MIN_HEADER_KEYWORD_MATCHES:
            continue
        width = len(trim_trailing_empty(vals))
        if best_row is None or score > best_score or (
            score == best_score and r_idx < best_row
        ):
            best_score = score
            best_row = r_idx
            best_width = max(width, 1)

    if best_row is None:
        return None
    return best_row, best_width, best_score


def detect_header_xlsx(path: Path, keywords: Sequence[str]) -> Tuple[int, int, int]:
    """
    Scan first HEADER_SCAN_MAX_ROW rows on every worksheet (max_col capped).

    Returns (sheet_index_0based, header_row_1based, header_col_count).
    Only rows with score >= MIN_HEADER_KEYWORD_MATCHES are candidates.
    """
    wb = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        best_sheet: Optional[int] = None
        best_row: Optional[int] = None
        best_score = -1
        best_width = 1

        for sheet_idx, ws in enumerate(wb.worksheets):
            det = detect_header_row_on_worksheet(ws, keywords)
            if det is None:
                continue
            r_idx, width, score = det
            if _header_row_better(sheet_idx, r_idx, score, best_sheet, best_row, best_score):
                best_score = score
                best_sheet = sheet_idx
                best_row = r_idx
                best_width = max(width, 1)

        if best_row is None or best_sheet is None:
            raise RuntimeError(
                f"{path.name}: no row in 1–{HEADER_SCAN_MAX_ROW} on any sheet matched "
                f"at least {MIN_HEADER_KEYWORD_MATCHES} header keyword(s). "
                "Adjust HEADER_KEYWORDS or INPUT_FILE."
            )

        return best_sheet, best_row, best_width
    finally:
        wb.close()


def _append_xlsx_worksheet_rows(
    writer: Any,
    ws,
    first_row_inclusive: int,
    out_cols: int,
    sheet_label: str = "",
) -> int:
    """
    Append rows from ws starting at first_row_inclusive (1-based) through tail-empty
    stop. One forward read_only pass from row 1 (required by openpyxl read_only).
    """
    rows_written = 0
    tail_empty = 0
    scan_last = min(
        first_row_inclusive + XLSX_STREAM_MAX_ROWS - 1,
        1_048_576,
    )
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
        cells = list(row)[:out_cols]
        while len(cells) < out_cols:
            cells.append(None)
        has_data = any(norm_key(c) != "" for c in cells)
        if not has_data:
            tail_empty += 1
            if tail_empty >= XLSX_STREAM_TAIL_EMPTY_ROWS:
                break
        else:
            tail_empty = 0
        writer.writerow([cell_to_csv(c) for c in cells])
        rows_written += 1
    if sheet_label:
        print(
            f"[XLSX][WRITE] {sheet_label}: wrote {rows_written} row(s) "
            f"from row {first_row_inclusive} onward (out_cols={out_cols})"
        )
    return rows_written


def stream_xlsx_all_sheets_to_csv(
    path: Path,
    out_path: Path,
    keywords: Sequence[str],
) -> int:
    """
    One CSV from all worksheets: detect header on each sheet (rows 1–50), write
    header line once, then all data rows from every sheet in order.

    Every sheet must have at least one qualifying header row. Output width is the
    maximum header width found across sheets.

    Returns total CSV rows written (including the single header row).
    """
    path = path.resolve()
    out_path = Path(out_path).resolve()

    # read_only: iterating for header detection advances each sheet's stream; open
    # again before streaming so every sheet is read from row 1 once.
    sheet_specs: List[Tuple[int, int, int]] = []
    wb1 = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        for si, ws in enumerate(wb1.worksheets):
            sheet_name = getattr(ws, "title", f"Sheet{si}")
            print(
                f"[XLSX][SCAN] sheet {si} ({sheet_name}): "
                f"checking rows 1-{HEADER_SCAN_MAX_ROW} for header..."
            )
            det = detect_header_row_on_worksheet(ws, keywords)
            if det is None:
                raise RuntimeError(
                    f"{path.name}: sheet {si} ({getattr(ws, 'title', '?')!r}): "
                    f"no row in 1–{HEADER_SCAN_MAX_ROW} matched at least "
                    f"{MIN_HEADER_KEYWORD_MATCHES} header keyword(s). "
                    "Adjust HEADER_KEYWORDS or sheet layout."
                )
            header_row, header_cols, score = det
            print(
                f"[XLSX][HEADER] sheet {si} ({sheet_name}): "
                f"row {header_row}, score={score}, cols={header_cols}"
            )
            sheet_specs.append((si, header_row, header_cols))
    finally:
        wb1.close()

    out_cols = max(p[2] for p in sheet_specs)
    print(
        f"[XLSX][PLAN] {path.name}: {len(sheet_specs)} sheet(s), "
        f"single output width={out_cols} columns"
    )
    rows_written = 0
    wb2 = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        with out_path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            first_sheet = True
            for si, header_row, header_cols in sheet_specs:
                ws = wb2.worksheets[si]
                sheet_name = getattr(ws, "title", f"Sheet{si}")
                start = header_row if first_sheet else header_row + 1
                mode = "header+data" if first_sheet else "data-only"
                first_sheet = False
                print(
                    f"[XLSX][START] sheet {si} ({sheet_name}): "
                    f"{mode}, starting row {start}"
                )
                rows_written += _append_xlsx_worksheet_rows(
                    writer,
                    ws,
                    start,
                    out_cols,
                    f"sheet {si} ({sheet_name})",
                )
    finally:
        wb2.close()
    print(f"[XLSX][DONE] {path.name}: total rows written={rows_written}")
    return rows_written


def stream_xlsx_to_csv(
    path: Path,
    out_path: Path,
    header_row: int,
    header_cols: int,
    sheet_index: int = 0,
) -> int:
    """Write one worksheet to CSV (single sheet). Prefer stream_xlsx_all_sheets_to_csv for workbooks."""
    wb = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        if sheet_index < 0 or sheet_index >= len(wb.worksheets):
            raise IndexError(f"sheet_index {sheet_index} out of range for workbook")
        ws = wb.worksheets[sheet_index]
        with out_path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            return _append_xlsx_worksheet_rows(writer, ws, header_row, header_cols)
    finally:
        wb.close()


def get_last_used_row_col_xls(ws):
    xl_formulas = -4123
    xl_by_rows = 1
    xl_by_columns = 2
    xl_previous = 2
    lr = ws.Cells.Find(
        What="*",
        LookIn=xl_formulas,
        SearchOrder=xl_by_rows,
        SearchDirection=xl_previous,
    )
    lc = ws.Cells.Find(
        What="*",
        LookIn=xl_formulas,
        SearchOrder=xl_by_columns,
        SearchDirection=xl_previous,
    )
    if lr is None or lc is None:
        return 0, 0
    return lr.Row, lc.Column


def read_xls_row_text(ws, r: int, max_col: int) -> List[str]:
    return [ws.Cells(r, c).Text for c in range(1, max_col + 1)]


def detect_header_xls(path: Path, keywords: Sequence[str]) -> Tuple[int, int]:
    if not HAS_WIN32:
        raise RuntimeError("pywin32 is required for .xls files. Install with: pip install pywin32")

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(path))
        ws = workbook.Worksheets(1)
        last_row, last_col = get_last_used_row_col_xls(ws)
        scan_last = min(HEADER_SCAN_MAX_ROW, last_row) if last_row else 0
        if scan_last == 0:
            return 1, 1

        best_row: Optional[int] = None
        best_score = -1
        best_width = 1

        for r in range(1, scan_last + 1):
            vals = read_xls_row_text(ws, r, max(last_col, 1))
            while vals and (vals[-1] is None or str(vals[-1]).strip() == ""):
                vals.pop()
            score = score_header_row(vals, keywords)
            if score < MIN_HEADER_KEYWORD_MATCHES:
                continue
            width = max(len(vals), 1)
            if best_row is None or score > best_score or (
                score == best_score and r < best_row
            ):
                best_score = score
                best_row = r
                best_width = width

        if best_row is None:
            raise RuntimeError(
                f"{path.name}: no row in 1–{scan_last} matched at least "
                f"{MIN_HEADER_KEYWORD_MATCHES} header keyword(s). "
                "Adjust HEADER_KEYWORDS or INPUT_FILE."
            )

        return best_row, best_width
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


def stream_xls_to_csv(
    path: Path,
    out_path: Path,
    header_row: int,
    header_cols: int,
) -> int:
    if not HAS_WIN32:
        raise RuntimeError("pywin32 is required for .xls files.")

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    rows_written = 0
    try:
        excel = DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(path))
        ws = workbook.Worksheets(1)
        last_row, last_col = get_last_used_row_col_xls(ws)
        end_row = last_row if last_row else header_row

        with out_path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            for r in range(header_row, end_row + 1):
                raw = read_xls_row_text(ws, r, max(header_cols, last_col, 1))
                raw = raw[:header_cols]
                while len(raw) < header_cols:
                    raw.append("")
                writer.writerow([cell_to_csv(c) for c in raw])
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

    return rows_written


def convert_workbook(
    path: Path,
    keywords: Sequence[str],
    out_path: Optional[Path] = None,
) -> Path:
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
        n = stream_xlsx_all_sheets_to_csv(path, out_path, keywords)
    else:
        header_row, header_cols = detect_header_xls(path, keywords)
        n = stream_xls_to_csv(path, out_path, header_row, header_cols)

    if suffix == ".xlsx":
        print(
            f"[OK] {path.name} → {out_path.name}  "
            f"(all sheets → one CSV, {n} row(s))"
        )
    else:
        print(
            f"[OK] {path.name} → {out_path.name}  (header row {header_row}, "
            f"{header_cols} cols, {n} CSV row(s))"
        )
    return out_path


def find_workbook_files(folder: Path) -> List[Path]:
    if not folder.is_dir():
        return []
    return sorted(
        p
        for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in CONVERT_EXTENSIONS
        and not p.name.startswith("~$")
    )


def main() -> int:
    start_time = time.perf_counter()
    keywords = list(HEADER_KEYWORDS)
    if not keywords:
        print("[ERROR] HEADER_KEYWORDS is empty.", file=sys.stderr)
        return 1
    if len(keywords) < MIN_HEADER_KEYWORD_MATCHES:
        print(
            f"[ERROR] HEADER_KEYWORDS must list at least {MIN_HEADER_KEYWORD_MATCHES} "
            "strings (so a row can match that many).",
            file=sys.stderr,
        )
        return 1

    folder = Path(INPUT_FOLDER).resolve()
    if not folder.is_dir():
        print(f"[ERROR] Input folder does not exist: {folder}", file=sys.stderr)
        return 1

    all_files = sorted(p for p in folder.iterdir() if p.is_file())
    skipped_csv = sum(1 for p in all_files if p.suffix.lower() == ".csv")
    targets = find_workbook_files(folder)

    print(f"Input folder: {folder}")
    print(
        f"Found {len(all_files)} file(s) in folder "
        f"({skipped_csv} .csv skipped, {len(targets)} .xlsx/.xls to convert)."
    )

    code = 0
    for path in targets:
        try:
            convert_workbook(path, keywords, None)
        except Exception as e:
            print(f"[ERROR] {path.name}: {e}", file=sys.stderr)
            code = 1
    elapsed_seconds = time.perf_counter() - start_time
    print(f"[TIME] Total time taken: {elapsed_seconds:.2f} seconds")

    return code


if __name__ == "__main__":
    raise SystemExit(main())
