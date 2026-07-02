import os

# Script folder ("python scripts csv")
WORK_DIR = os.path.abspath(os.path.dirname(__file__))

# Subfolders under WORK_DIR — edit names here if needed
INPUT_DIR_NAME = "input"
OUTPUT_DIR_NAME = "output"

# Input VRN/LC merged CSV (inside input/)
input_file_name = "daroda_semi_final_output.csv"

# Final Excel output filename (inside output/)
output_file_name = "Daroda_final_output.xlsx"

# Return journey input (inside output/; use "MP Output.xlsx" if LT step is skipped)
return_journey_input_file = "LT_output.xlsx"


def get_work_dir():
    return WORK_DIR


def get_input_dir():
    path = os.path.join(WORK_DIR, INPUT_DIR_NAME)
    os.makedirs(path, exist_ok=True)
    return path


def get_output_dir():
    path = os.path.join(WORK_DIR, OUTPUT_DIR_NAME)
    os.makedirs(path, exist_ok=True)
    return path


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


# ====== Plaza Configuration ======
plaza_name = "DARODA"


def get_plaza_name():
    return plaza_name


def get_local_code():
    import json

    json_path = input_path("codes_dump.json")
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


# ====== Pass file column aliases (MP / LT working files) ======
PASS_CHASSIS_COLUMN_NAMES = [
    "Chassis/ Vehicle No",
    "Chassis/Vehicle No",
    "Chassis Vehicle No",
    "Vehicle No",
]

PASS_START_DATE_COLUMN_NAMES = [
    "Start Date",
    "Start Effective Date",
]

PASS_END_DATE_COLUMN_NAMES = [
    "End Date",
    "End Effective Date",
]


def _normalize_pass_col_name(col) -> str:
    s = str(col).strip().lower()
    return "".join(s.split())


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
