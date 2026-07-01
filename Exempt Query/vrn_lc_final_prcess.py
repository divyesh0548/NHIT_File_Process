"""
VRN + LC final combine.

Stage 1 — Append LC into VRN
Stage 2 — Merged = Veh Reg No. + Text.From(Date & Time, en-IN) + MOP (no separator)
Stage 3 — Dedupe by Merged (keep first)
Stage 4 — Drop Merged column
Stage 5 — Uppercase Description, then map AMBUL→AMBULANCE, GOV→GOVT, LOCAL→LOCAL
Stage 6 — Uppercase MOP
"""

from __future__ import annotations

import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

# --- Input / output paths (edit here) ---
VRN_FILE = BASE_DIR / "normalized_and_merged_vrn_daroda.csv"
LC_FILE = BASE_DIR / "normalized_and_merged_etc_daroda.csv"
OUTPUT_FILE = BASE_DIR / "daroda_semi_final_output.csv"

# Debug: rows removed as duplicates (second+ occurrence per Merged) → Excel before dedupe
DEBUG_EXPORT_DUPLICATES = True
DEBUG_DUPLICATES_XLSX = BASE_DIR / "daroda_semi_final_duplicates_debug.xlsx"

REQUIRED_MERGE_COLS = ("Veh Reg No.", "Date & Time", "MOP")
MERGED_COL = "Merged"
DESC_COL = "Description"
MOP_COL = "MOP" 


def _read_csv(path: Path):
    import pandas as pd

    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return pd.read_csv(path, encoding=enc, low_memory=False)
        except UnicodeDecodeError:
            continue
    return pd.read_csv(path, encoding="latin-1", low_memory=False)


def stage1_append_lc_to_vrn(vrn_df, lc_df):
    """Append LC rows under VRN; LC-only columns must not exist (VRN is superset)."""
    import pandas as pd

    vrn_cols = list(vrn_df.columns)
    lc_extra = [c for c in lc_df.columns if c not in vrn_cols]
    if lc_extra:
        raise ValueError(
            "LC file has columns not present in VRN header: "
            + ", ".join(repr(c) for c in lc_extra)
        )

    lc_aligned = lc_df.reindex(columns=vrn_cols, fill_value="")
    out = pd.concat([vrn_df, lc_aligned], ignore_index=True)
    return out


def _date_time_to_en_in_text(val) -> str:
    import pandas as pd

    if pd.isna(val):
        return ""
    if isinstance(val, str):
        s = val.strip()
        if not s:
            return ""
        ts = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.isna(ts):
            return s
        return ts.strftime("%d/%m/%Y %H:%M:%S")
    ts = pd.to_datetime(val, errors="coerce")
    if pd.isna(ts):
        return str(val).strip()
    return ts.strftime("%d/%m/%Y %H:%M:%S")


def stage2_insert_merged_column(df):
    import pandas as pd

    missing = [c for c in REQUIRED_MERGE_COLS if c not in df.columns]
    if missing:
        raise ValueError(
            "After append, these columns are required for merge: "
            + ", ".join(missing)
        )

    parts_vrn = df["Veh Reg No."].map(lambda x: "" if pd.isna(x) else str(x).strip())
    parts_dt = df["Date & Time"].map(_date_time_to_en_in_text)
    parts_mop = df["MOP"].map(lambda x: "" if pd.isna(x) else str(x).strip())
    df = df.copy()
    df[MERGED_COL] = parts_vrn.astype(str) + parts_dt.astype(str) + parts_mop.astype(str)
    return df


def stage3_drop_duplicates_by_merged(
    df,
    debug_duplicates_path: Path | None = None,
    ):
    """
    Keep first row per Merged; optionally write rows that will be dropped to Excel.
    """
    if MERGED_COL not in df.columns:
        raise ValueError(f"Expected column {MERGED_COL!r} before dedupe.")

    dup_mask = df.duplicated(subset=[MERGED_COL], keep="first")
    dup_rows = df.loc[dup_mask].copy()

    if debug_duplicates_path is not None:
        debug_duplicates_path = Path(debug_duplicates_path).resolve()
        debug_duplicates_path.parent.mkdir(parents=True, exist_ok=True)
        if len(dup_rows) == 0:
            print("[DEBUG] No duplicate rows to export.", flush=True)
        else:
            dup_rows.to_excel(debug_duplicates_path, index=False, engine="openpyxl")
            print(
                f"[DEBUG] Wrote {len(dup_rows)} duplicate row(s) (removed in stage 3) "
                f"to {debug_duplicates_path}",
                flush=True,
            )

    before = len(df)
    out = df.loc[~dup_mask].reset_index(drop=True)
    removed = before - len(out)
    return out, removed


def stage4_remove_merged_column(df):
    """Drop Merged (dedupe key only). Also drops Month Name if present (PBI cleanup)."""
    drop = [c for c in (MERGED_COL, "Month Name") if c in df.columns]
    if not drop:
        return df
    return df.drop(columns=drop)


def _transform_description_cell(val):
    """
    Power BI: Text.Upper then substring rules (order: AMBUL, GOV, LOCAL).
    null stays null; whole value replaced when rule matches.
    """
    import pandas as pd

    if val is None or (isinstance(val, float) and pd.isna(val)):
        return val
    if pd.isna(val):
        return val

    text = str(val).strip()
    if not text:
        return val

    upper = text.upper()
    if "AMBUL" in upper:
        return "AMBULANCE"
    if "GOV" in upper:
        return "GOVT"
    if "LOCAL" in upper:
        return "LOCAL"
    return upper


def stage5_transform_description(df):
    if DESC_COL not in df.columns:
        print(f"      [WARN] No {DESC_COL!r} column; skipping description transforms.", flush=True)
        return df
    out = df.copy()
    out[DESC_COL] = out[DESC_COL].map(_transform_description_cell)
    return out


def stage6_uppercase_mop(df):
    import pandas as pd

    if MOP_COL not in df.columns:
        print(f"      [WARN] No {MOP_COL!r} column; skipping MOP uppercase.", flush=True)
        return df
    out = df.copy()

    def _upper_mop(x):
        if x is None or pd.isna(x):
            return x
        return str(x).strip().upper()

    out[MOP_COL] = out[MOP_COL].map(_upper_mop)
    return out


def stage4_to_6_post_dedupe(df):
    """After dedupe: remove Merged, normalize Description, uppercase MOP."""
    to_drop = [c for c in (MERGED_COL, "Month Name") if c in df.columns]
    if to_drop:
        print(f"      Dropping column(s): {', '.join(to_drop)}", flush=True)
    out = stage4_remove_merged_column(df)
    out = stage5_transform_description(out)
    out = stage6_uppercase_mop(out)
    return out


def run_pipeline(vrn_path: Path, lc_path: Path, output_path: Path) -> None:
    debug_xlsx = DEBUG_DUPLICATES_XLSX if DEBUG_EXPORT_DUPLICATES else None

    print(f"[1/6] Reading VRN: {vrn_path}", flush=True)
    vrn_df = _read_csv(vrn_path)
    print(f"      VRN rows={len(vrn_df)}, cols={len(vrn_df.columns)}", flush=True)

    print(f"[1/6] Reading LC:  {lc_path}", flush=True)
    lc_df = _read_csv(lc_path)
    print(f"      LC rows={len(lc_df)}, cols={len(lc_df.columns)}", flush=True)

    print("[1/6] Appending LC rows under VRN (align columns, pad missing)...", flush=True)
    combined = stage1_append_lc_to_vrn(vrn_df, lc_df)
    print(f"      Combined rows={len(combined)}", flush=True)

    print("[2/6] Building Merged column...", flush=True)
    merged_df = stage2_insert_merged_column(combined)
    print(f"      Added column {MERGED_COL!r}", flush=True)

    print("[3/6] Dedupe by Merged (keep first)...", flush=True)
    deduped_df, dup_removed = stage3_drop_duplicates_by_merged(
        merged_df, debug_duplicates_path=debug_xlsx
    )
    print(f"      Removed {dup_removed} duplicate row(s); rows={len(deduped_df)}", flush=True)

    print("[4/6] Remove Merged column...", flush=True)
    print("[5/6] Description: uppercase + AMBUL/GOV/LOCAL mapping...", flush=True)
    print("[6/6] MOP: uppercase...", flush=True)
    final_df = stage4_to_6_post_dedupe(deduped_df)
    print(f"      Final rows={len(final_df)}, cols={len(final_df.columns)}", flush=True)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    final_df.to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"[OUT] Wrote {output_path}", flush=True)


def main() -> int:
    try:
        import pandas as pd  # noqa: F401
    except ImportError:
        print("Requires pandas: pip install pandas", file=sys.stderr)
        return 1

    if DEBUG_EXPORT_DUPLICATES:
        try:
            import openpyxl  # noqa: F401
        except ImportError:
            print(
                "DEBUG_EXPORT_DUPLICATES needs openpyxl: pip install openpyxl",
                file=sys.stderr,
            )
            return 1

    vrn_path = Path(VRN_FILE).resolve()
    lc_path = Path(LC_FILE).resolve()
    output_path = Path(OUTPUT_FILE).resolve()

    if not vrn_path.is_file():
        print(f"VRN file not found: {vrn_path}", file=sys.stderr)
        return 1
    if not lc_path.is_file():
        print(f"LC file not found: {lc_path}", file=sys.stderr)
        return 1

    try:
        run_pipeline(vrn_path, lc_path, output_path)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
