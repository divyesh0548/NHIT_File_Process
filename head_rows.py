from pathlib import Path
from typing import Optional

import pandas as pd

# Set the value to look for in the top N rows of each file/sheet.
SEARCH_VALUE = "Veh Reg No."

# When True, scan top N rows for SEARCH_VALUE and print match results.
SEARCH_IN_TOP_ROWS = False

# When True, save the first N rows of each Excel file's first sheet as a new .xlsx.
SAVE_TOP_ROWS_TO_EXCEL = True

TOP_N_ROWS = 30


def _read_top_rows(file_path: Path, n_rows: int, suffix: str, sheet_name: Optional[str] = None):
    if suffix == ".csv":
        return pd.read_csv(file_path, nrows=n_rows)
    if sheet_name is None:
        raise ValueError("sheet_name is required for Excel files")
    return pd.read_excel(file_path, sheet_name=sheet_name, nrows=n_rows)


def _find_value_in_top_rows(df: pd.DataFrame, search_value: str, n_rows: int):
    """
    Search the first n_rows of df for search_value (case-insensitive substring match).
    Returns a list of (row_index, column_name, cell_value) matches.
    """
    if not search_value:
        return []

    needle = search_value.strip().casefold()
    if not needle:
        return []

    sample = df.head(n_rows)
    matches = []

    for row_idx, row in sample.iterrows():
        for col_name, cell in row.items():
            if pd.isna(cell):
                continue
            haystack = str(cell).strip().casefold()
            if needle in haystack:
                matches.append((row_idx, col_name, cell))

    return matches


def _print_search_results(label: str, search_value: str, n_rows: int, matches):
    print(f"\n[SEARCH] {label}")
    if not matches:
        print(f'  Value "{search_value}" not found in top {n_rows} row(s).')
        return

    print(f'  Value "{search_value}" found {len(matches)} time(s) in top {n_rows} row(s):')
    for row_idx, col_name, cell in matches:
        print(f"    row={row_idx}, column={col_name!r}, cell={cell!r}")


def _save_top_rows_excel(df: pd.DataFrame, source_path: Path, n_rows: int, out_dir: Path):
    """Write the first n_rows of df to out_dir/<stem>_head_<n>.xlsx."""
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{source_path.stem}_head_{n_rows}.xlsx"
    df.head(n_rows).to_excel(out_path, index=False, sheet_name="Sheet1")
    print(f"[SAVED] {out_path}")
    return out_path


def print_top_rows_in_folder(
    folder_path,
    n_rows=TOP_N_ROWS,
    search_value=SEARCH_VALUE,
    search_in_top_rows=SEARCH_IN_TOP_ROWS,
    save_top_rows_to_excel=SAVE_TOP_ROWS_TO_EXCEL,
    ):
    """
    Print top N rows for every .csv/.xls/.xlsx file in a folder.
    For Excel files, prints top rows for each sheet.

    When search_in_top_rows is True, also scans the top N rows for search_value.
    When save_top_rows_to_excel is True, saves the first N rows of each Excel
    file's first sheet only to <folder>/head_rows/<stem>_head_<n>.xlsx.
    """
    folder = Path(folder_path)
    if not folder.exists() or not folder.is_dir():
        print(f"Invalid folder path: {folder_path}")
        return

    files = sorted(
        [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in {".csv", ".xls", ".xlsx"}]
    )

    if not files:
        print("No .csv/.xls/.xlsx files found in the folder.")
        return

    out_dir = folder / "head_rows"

    for file_path in files:
        suffix = file_path.suffix.lower()
        print("\n" + "=" * 90)
        print(f"FILE: {file_path.name}")
        print("=" * 90)

        try:
            if suffix == ".csv":
                df = _read_top_rows(file_path, n_rows, suffix)
                print(df)
                if search_in_top_rows:
                    matches = _find_value_in_top_rows(df, search_value, n_rows)
                    _print_search_results(file_path.name, search_value, n_rows, matches)
            else:
                excel_data = pd.ExcelFile(file_path)
                for sheet_name in excel_data.sheet_names:
                    print(f"\n--- SHEET: {sheet_name} ---")
                    df = _read_top_rows(file_path, n_rows, suffix, sheet_name=sheet_name)
                    print(df)
                    if search_in_top_rows:
                        matches = _find_value_in_top_rows(df, search_value, n_rows)
                        _print_search_results(
                            f"{file_path.name} [{sheet_name}]",
                            search_value,
                            n_rows,
                            matches,
                        )

                if save_top_rows_to_excel and excel_data.sheet_names:
                    first_sheet = excel_data.sheet_names[0]
                    head_df = _read_top_rows(
                        file_path, n_rows, suffix, sheet_name=first_sheet
                    )
                    _save_top_rows_excel(head_df, file_path, n_rows, out_dir)
        except Exception as exc:
            print(f"Error reading {file_path.name}: {exc}")


if __name__ == "__main__":
    folder_path = "Exempt Query/Daroda Exempt Query Base Files/vrn"  # Replace with your folder
    print_top_rows_in_folder(
        folder_path,
        n_rows=TOP_N_ROWS,
        search_value=SEARCH_VALUE,
        search_in_top_rows=SEARCH_IN_TOP_ROWS,
        save_top_rows_to_excel=SAVE_TOP_ROWS_TO_EXCEL,
    )
