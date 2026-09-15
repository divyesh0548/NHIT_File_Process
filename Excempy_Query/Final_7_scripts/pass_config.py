"""
Pass-file keyword / alias configuration.

Defaults live here.  Editable values are stored in pass_separator_config.json
(same folder).  Used by Pass_file_seperator.py and config.py.
"""

from __future__ import annotations

import json
import os
import tempfile
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

CONFIG_DIR = Path(__file__).resolve().parent
CONFIG_JSON_PATH = Path(
    os.environ.get(
        "PASS_SEPARATOR_CONFIG_JSON",
        str(CONFIG_DIR / "pass_separator_config.json"),
    )
).resolve()

_DEFAULT_VALUES: Dict[str, Any] = {
    "HEADER_KEYWORDS": [
        "Chassis/ Vehicle No",
        "NPCI Vehicle Class",
        "Start Date",
        "Plaza Name",
        "End Date",
        "Mobile No",
        "Pass Type",
        "Payment Mode",
        "Agency Code",
        "Vehicle Reg No",
        "Issued Date",
    ],
    "PASS_TYPE_COLUMN_NAMES": ["Pass Type", "PassType", "PASS_TYPE"],
    "MP_PASS_TYPE_VALUES": [
        "MP",
        "LP",
        "MLP",
        "Monthly pass",
        "Local 20 Km VC4 (350)",
        "Local20KM",
    ],
    "LT_PASS_TYPE_VALUES": [
        "LT",
        "Local Commercial Tariff 50%",
        "50 % discount for local Comm",
    ],
    "PASS_CHASSIS_COLUMN_NAMES": [
        "Chassis/ Vehicle No",
        "Chassis/Vehicle No",
        "Chassis Vehicle No",
        "Vehicle No",
        "Vehicle Reg No",
    ],
    "PASS_START_DATE_COLUMN_NAMES": [
        "Start Date",
        "Start Effective Date",
        "Validity Start Date",
    ],
    "PASS_END_DATE_COLUMN_NAMES": [
        "End Date",
        "End Effective Date",
        "Validity End Date",
    ],
    "HEADER_SCAN_ROWS": 25,
    "MIN_HEADER_MATCHES": 3,
}

CONFIG_SCHEMA = [
    {
        "key": "HEADER_KEYWORDS",
        "type": "list",
        "label": "Header Detection Keywords",
        "description": "Column names scanned in the first N rows to detect the header row.",
        "order": 1,
    },
    {
        "key": "PASS_TYPE_COLUMN_NAMES",
        "type": "list",
        "label": "Pass Type Column Names",
        "description": "Possible column names for the Pass Type field.",
        "order": 2,
    },
    {
        "key": "MP_PASS_TYPE_VALUES",
        "type": "list",
        "label": "MP (Monthly Pass) Type Values",
        "description": "Pass Type cell values that map to the MP workbook.",
        "order": 3,
    },
    {
        "key": "LT_PASS_TYPE_VALUES",
        "type": "list",
        "label": "LT (Local Tariff) Type Values",
        "description": "Pass Type cell values that map to the LT workbook.",
        "order": 4,
    },
    {
        "key": "PASS_CHASSIS_COLUMN_NAMES",
        "type": "list",
        "label": "Chassis / Vehicle No Column Aliases",
        "description": "Column name aliases for the chassis/vehicle number field in pass files.",
        "order": 5,
    },
    {
        "key": "PASS_START_DATE_COLUMN_NAMES",
        "type": "list",
        "label": "Start Date Column Aliases",
        "description": "Column name aliases for the pass start-date field.",
        "order": 6,
    },
    {
        "key": "PASS_END_DATE_COLUMN_NAMES",
        "type": "list",
        "label": "End Date Column Aliases",
        "description": "Column name aliases for the pass end-date field.",
        "order": 7,
    },
    {
        "key": "HEADER_SCAN_ROWS",
        "type": "integer",
        "label": "Header Scan Rows",
        "description": "Number of rows to scan from the top when detecting the header.",
        "order": 8,
    },
    {
        "key": "MIN_HEADER_MATCHES",
        "type": "integer",
        "label": "Minimum Header Matches",
        "description": "Minimum keyword matches required to accept a row as the header.",
        "order": 9,
    },
]


def get_default_values() -> Dict[str, Any]:
    return deepcopy(_DEFAULT_VALUES)


def get_config_schema_for_api() -> List[Dict[str, Any]]:
    return sorted(deepcopy(CONFIG_SCHEMA), key=lambda s: s.get("order", 999))


def _normalize_list(value: Any, field_name: str) -> List[str]:
    if not isinstance(value, list):
        raise ValueError(f"{field_name} must be a list of strings.")
    out: List[str] = []
    for item in value:
        if item is None:
            continue
        text = str(item)
        if text.strip() == "":
            continue
        out.append(text)
    return out


def _normalize_integer(value: Any, field_name: str) -> int:
    try:
        v = int(value)
    except (TypeError, ValueError):
        raise ValueError(f"{field_name} must be an integer.")
    if v < 1:
        raise ValueError(f"{field_name} must be at least 1.")
    return v


def validate_and_normalize_values(payload: Dict[str, Any]) -> Dict[str, Any]:
    if not isinstance(payload, dict):
        raise ValueError("Config payload must be a JSON object.")

    schema_map = {s["key"]: s for s in CONFIG_SCHEMA}
    unknown = sorted(set(payload) - set(schema_map))
    if unknown:
        raise ValueError("Unknown key(s): " + ", ".join(unknown))
    missing = sorted(set(schema_map) - set(payload))
    if missing:
        raise ValueError("Missing required key(s): " + ", ".join(missing))

    normalized: Dict[str, Any] = {}
    for key, meta in schema_map.items():
        if meta["type"] == "list":
            normalized[key] = _normalize_list(payload[key], key)
        elif meta["type"] == "integer":
            normalized[key] = _normalize_integer(payload[key], key)
        else:
            normalized[key] = payload[key]
    return normalized


def _read_json_file(path: Path) -> Optional[Dict[str, Any]]:
    if not path.is_file():
        return None
    raw = path.read_text(encoding="utf-8").strip()
    if not raw:
        return None
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValueError(f"{path.name} must contain a JSON object.")
    return data


def load_config_values() -> Tuple[Dict[str, Any], bool]:
    defaults = get_default_values()
    try:
        stored = _read_json_file(CONFIG_JSON_PATH)
    except (OSError, json.JSONDecodeError, ValueError):
        return defaults, False
    if stored is None:
        return defaults, False
    merged = deepcopy(defaults)
    for key in _DEFAULT_VALUES:
        if key in stored:
            merged[key] = stored[key]
    try:
        return validate_and_normalize_values(merged), True
    except ValueError:
        return defaults, False


def save_config_values(payload: Dict[str, Any]) -> Dict[str, Any]:
    normalized = validate_and_normalize_values(payload)
    CONFIG_JSON_PATH.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(
        prefix="pass_sep_cfg_", suffix=".json", dir=str(CONFIG_JSON_PATH.parent)
    )
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(normalized, handle, indent=2, ensure_ascii=False)
            handle.write("\n")
        os.replace(tmp_name, CONFIG_JSON_PATH)
    except Exception:
        try:
            os.unlink(tmp_name)
        except OSError:
            pass
        raise
    return normalized


def reset_config_to_defaults() -> Dict[str, Any]:
    return save_config_values(get_default_values())


def ensure_config_json_exists() -> Dict[str, Any]:
    values, from_file = load_config_values()
    if not from_file:
        return save_config_values(values)
    return values
