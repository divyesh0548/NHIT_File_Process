import os
import sys
from pathlib import Path

import pandas as pd

_PORTAL_ROOT = Path(__file__).resolve().parents[3]
if str(_PORTAL_ROOT) not in sys.path:
    sys.path.insert(0, str(_PORTAL_ROOT))

from header_matching import normalize_header_match


def _env_path(name, default):
    """Resolve a runtime path supplied by the portal, with a local fallback."""
    return Path(os.environ.get(name, str(default))).expanduser().resolve()


# Every final-stage run gets its own working tree from the portal.  The defaults
# retain standalone-script usability without requiring edits to this file.
WORK_DIR = _env_path("EXEMPT_FINAL_WORK_DIR", Path(__file__).resolve().parent)
INPUT_DIR = _env_path("EXEMPT_FINAL_INPUT_DIR", WORK_DIR / "input")
OUTPUT_DIR = _env_path("EXEMPT_FINAL_OUTPUT_DIR", WORK_DIR / "output")

# File names are deliberately names, not absolute paths.  config resolves them
# inside the per-run folders above.
input_file_name = os.environ.get("EXEMPT_FINAL_SEMI_CSV_NAME", "daroda_semi_final_output.csv")
output_file_name = os.environ.get("EXEMPT_FINAL_OUTPUT_FILE_NAME", "final_output.xlsx")
return_journey_input_file = os.environ.get(
    "EXEMPT_FINAL_RETURN_INPUT", "MP Output.xlsx"
)
codes_dump_file = os.environ.get("EXEMPT_FINAL_CODES_FILE", "")
plaza_name = os.environ.get("EXEMPT_FINAL_PLAZA", "").strip()


def get_work_dir():
    return str(WORK_DIR)


def get_input_dir():
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    return str(INPUT_DIR)


def get_output_dir():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    return str(OUTPUT_DIR)


def input_path(*parts: str) -> str:
    """Build an absolute path under the input folder."""
    return os.path.join(get_input_dir(), *parts)


def output_path(*parts: str) -> str:
    """Build an absolute path under the output folder."""
    return os.path.join(get_output_dir(), *parts)


def get_file_path():
    return input_path(input_file_name)


def put_file_name():
    return output_file_name


def get_return_journey_input_path():
    return output_path(return_journey_input_file)


def get_plaza_name():
    return plaza_name


def get_local_code():
    import json

    json_path = codes_dump_file or input_path("codes_dump.json")
    try:
        with open(json_path, "r") as f:
            codes = json.load(f)
            codes_upper = {k.upper().strip(): str(v) for k, v in codes.items()}
            raw_codes = codes_upper.get(plaza_name.upper().strip())

            if not raw_codes:
                return ()

            return tuple(
                code.strip()
                for code in raw_codes.replace("and", ",").split(",")
                if code.strip()
            )

    except Exception as e:
        print(
            f"Warning: Could not read local code from codes_dump.json. "
            f"Defaulting to null. Error: {e}"
        )
        return ()


# ====== Pass file column aliases & type values (loaded from JSON with defaults) ======
from pass_config import load_config_values as _load_pass_config


def _pass_cfg():
    cfg, _from_file = _load_pass_config()
    return cfg


_PASS_CFG = _pass_cfg()

PASS_CHASSIS_COLUMN_NAMES = _PASS_CFG["PASS_CHASSIS_COLUMN_NAMES"]
PASS_START_DATE_COLUMN_NAMES = _PASS_CFG["PASS_START_DATE_COLUMN_NAMES"]
PASS_END_DATE_COLUMN_NAMES = _PASS_CFG["PASS_END_DATE_COLUMN_NAMES"]
MP_PASS_TYPE_VALUES = _PASS_CFG["MP_PASS_TYPE_VALUES"]
LT_PASS_TYPE_VALUES = _PASS_CFG["LT_PASS_TYPE_VALUES"]


def _normalize_pass_col_name(col) -> str:
    return normalize_header_match(col)


def normalize_pass_type_value(value) -> str:
    """Normalize a Pass Type cell for alias matching (same rules as header matching)."""
    return normalize_header_match(value)


def mp_pass_type_normalized_set():
    return {normalize_pass_type_value(v) for v in MP_PASS_TYPE_VALUES if normalize_pass_type_value(v)}


def lt_pass_type_normalized_set():
    return {normalize_pass_type_value(v) for v in LT_PASS_TYPE_VALUES if normalize_pass_type_value(v)}


def find_pass_file_column(df, candidates, label: str) -> str:
    """Return the actual column name in df matching one of candidates."""
    norm_to_col = {_normalize_pass_col_name(c): c for c in df.columns}
    for name in candidates:
        key = _normalize_pass_col_name(name)
        if key in norm_to_col:
            return norm_to_col[key]
    raise KeyError(
        f"Could not find {label} column. Tried {candidates}. "
        f"Available columns: {list(df.columns)}"
    )


def standardize_pass_file_columns(df):
    """Rename pass-file columns to canonical names used by date validity scripts."""
    rename = {
        find_pass_file_column(df, PASS_CHASSIS_COLUMN_NAMES, "Chassis/ Vehicle No"): "Chassis/ Vehicle No",
        find_pass_file_column(df, PASS_START_DATE_COLUMN_NAMES, "Start Date"): "Start Date",
        find_pass_file_column(df, PASS_END_DATE_COLUMN_NAMES, "End Date"): "End Date",
    }
    return df.rename(columns=rename)


def parse_pass_date_series(values, label="pass date"):
    """Parse Excel, ISO, and Indian day-first pass dates in one column.

    Different monthly pass exports mix real Excel dates with strings such as
    ``23/02/2026 00:00:00``.  pandas' default parser can lock onto a month-first
    format from an earlier value and then fail for the Indian day-first values.
    ``format='mixed'`` parses each value independently while ``dayfirst=True``
    preserves the convention used in these pass reports.
    """
    parsed = pd.to_datetime(values, format="mixed", dayfirst=True, errors="coerce")
    invalid_count = int(values.notna().sum() - parsed.notna().sum())
    if invalid_count:
        print(
            f"Warning: {invalid_count} {label} value(s) could not be parsed and will not match a pass window.",
            flush=True,
        )
    return parsed.dt.date
