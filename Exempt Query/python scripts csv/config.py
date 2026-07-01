import os

# Project root = this folder ("python scripts csv")
PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))

# Paths relative to PROJECT_ROOT — edit these to point at files inside the project
file_path = "vrn_lc_merged_daroda.csv"
base_output_path = "output_daroda"

def get_file_path():
    return os.path.join(PROJECT_ROOT, file_path)

def get_base_output_path():
    path = os.path.join(PROJECT_ROOT, base_output_path)
    os.makedirs(path, exist_ok=True)
    return path

file_name = "Daroda_final_output.xlsx"
def put_file_name():
    return file_name

# ====== Plaza Configuration ======
plaza_name = "DARODA"

def get_plaza_name():
    return plaza_name

def get_local_code():
    import json

    json_path = os.path.join(os.path.dirname(__file__), 'codes_dump.json')
    try:
        with open(json_path, 'r') as f:
            codes = json.load(f)
            # Make case-insensitive lookup
            codes_upper = {k.upper().strip(): str(v) for k, v in codes.items()}
            raw_codes = codes_upper.get(plaza_name.upper().strip())

            if not raw_codes:
                return ()

            # Allow multiple codes separated by commas (e.g. "WB50, WB49")
            return tuple(code.strip() for code in raw_codes.replace('and', ',').split(',') if code.strip())

    except Exception as e:
        print(f"Warning: Could not read local code from codes_dump.json. Defaulting to null. Error: {e}")
        return ()
