"""
LC/ETC (Lifecycle) normalization header groups config.

Defaults live here. Editable values are stored in LC_Normalization_Data.json
(same folder). Used only by LC/ETC Merge + Normalize (normalize_lcy_faster.py).
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
        "LC_NORMALIZATION_CONFIG_JSON",
        str(CONFIG_DIR / "LC_Normalization_Data.json"),
    )
).resolve()

# Canonical group name → alias list (defaults).
_DEFAULT_NORMALIZATION_GROUPS: Dict[str, List[str]] = {
    "Date & Time": [
        "DATE",
        "Date Time",
        "Reader Read Time",
        "DATETIME",
        "DATE TIME",
        "Date/Time",
        "Transaction Date",
        "Txn Date",
        "TXN DATE",
    ],
    "Settlement Type": [
        "SettlementType",
        "SETTLEMENT TYPE",
        "Settlement",
    ],
    "Transaction Status": [
        "Transaction Status",
        "TRANSACTION STATUS",
        "TransactionStatus",
        "Txn Status",
    ],
    "Veh Reg No.": [
        "TC_VEH_REG_NO",
        "Vehicle Reg. No.",
        "Vehicle Reg. No",
        "VEHICLE_REG_NO",
        "veh_reg_no_",
        "Veh Reg No.",
        "VEH REG NO",
        "TC VEH REG NO",
        "VEH. REG. NO.",
        "Licence Plate No",
        "VRN",
        "Licence Plate No.",
        "Plate No",
        "NPCI VRN",
        "Veh Reg Num",
        "VehicleNumber",
        "VEH_REG_NO",
        "Platenumber",
        "Vehicle Registration Number",
        "vehicle_reg_no",
        "Vehicle No",
    ],
    "MOP": [
        "PAYMENT_TYPE",
        "Payment Method",
        "MVC MOP",
        "PAYMENT TYPE",
        "Payment",
        "Mode",
        "Ticket_Type",
        "MVC_TLC_MOP",
        "TransactionTypeTC",
        "PaymentMeans",
        "PAYMENT METHOD",
        "MVC (TLC MOP)",
        "MVC TLC MOP",
    ],
    "Lane No": [
        "LANE NO",
        "LaneNo",
        "Lane ID",
        "Lane Number",
        "LANE_NUMBER",
        "Lane",
    ],
}

# Fixed group order / labels for the portal UI. Group keys cannot be removed.
_GROUP_META: Dict[str, Dict[str, Any]] = {
    "Date & Time": {"label": "Date & Time aliases", "order": 1},
    "Settlement Type": {"label": "Settlement Type aliases", "order": 2},
    "Transaction Status": {"label": "Transaction Status aliases", "order": 3},
    "Veh Reg No.": {"label": "Veh Reg No. aliases", "order": 4},
    "MOP": {"label": "MOP aliases", "order": 5},
    "Lane No": {"label": "Lane No aliases", "order": 6},
}


def get_default_values() -> Dict[str, List[str]]:
    return deepcopy(_DEFAULT_NORMALIZATION_GROUPS)


def get_config_schema_for_api() -> List[Dict[str, Any]]:
    items = []
    for key, meta in _GROUP_META.items():
        items.append(
            {
                "key": key,
                "type": "list",
                "label": meta["label"],
                "description": f'Aliases rewritten to canonical "{key}".',
                "order": meta.get("order", 999),
                "editable_keys": False,
            }
        )
    items.sort(key=lambda row: (row["order"], row["key"]))
    return items


def _normalize_alias_list(value: Any, field_name: str) -> List[str]:
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


def validate_and_normalize_values(payload: Dict[str, Any]) -> Dict[str, List[str]]:
    if not isinstance(payload, dict):
        raise ValueError("Config payload must be a JSON object.")

    required_keys = list(_DEFAULT_NORMALIZATION_GROUPS.keys())
    unknown = sorted(set(payload) - set(required_keys))
    if unknown:
        raise ValueError(
            "Unknown normalization group(s) are not allowed: " + ", ".join(unknown)
        )

    missing = sorted(set(required_keys) - set(payload))
    if missing:
        raise ValueError(
            "Missing required normalization group(s): " + ", ".join(missing)
        )

    normalized: Dict[str, List[str]] = {}
    for key in required_keys:
        normalized[key] = _normalize_alias_list(payload[key], key)
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


def load_config_values() -> Tuple[Dict[str, List[str]], bool]:
    """
    Return (groups, from_file).

    Falls back to defaults if JSON is missing, empty, unreadable, or invalid.
    """
    defaults = get_default_values()
    try:
        stored = _read_json_file(CONFIG_JSON_PATH)
    except (OSError, json.JSONDecodeError, ValueError):
        return defaults, False

    if stored is None:
        return defaults, False

    merged = deepcopy(defaults)
    for key in _DEFAULT_NORMALIZATION_GROUPS:
        if key in stored:
            merged[key] = stored[key]

    try:
        return validate_and_normalize_values(merged), True
    except ValueError:
        return defaults, False


def get_normalization_groups() -> Dict[str, List[str]]:
    """Resolved groups for LC/ETC normalize scripts."""
    values, _from_file = load_config_values()
    return values


def save_config_values(payload: Dict[str, Any]) -> Dict[str, List[str]]:
    normalized = validate_and_normalize_values(payload)
    CONFIG_JSON_PATH.parent.mkdir(parents=True, exist_ok=True)

    fd, tmp_name = tempfile.mkstemp(
        prefix="lc_normalization_",
        suffix=".json",
        dir=str(CONFIG_JSON_PATH.parent),
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


def reset_config_to_defaults() -> Dict[str, List[str]]:
    return save_config_values(get_default_values())


def ensure_config_json_exists() -> Dict[str, List[str]]:
    values, from_file = load_config_values()
    if not from_file:
        return save_config_values(values)
    return values
