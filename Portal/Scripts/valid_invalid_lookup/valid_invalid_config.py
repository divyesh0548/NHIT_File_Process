"""
Configurable keyword lists and mappings for Valid/Invalid Lookup.

Defaults live in this module. Editable values are stored in
valid_invalid_config.json (same folder). The portal reads/writes that JSON;
process scripts import the resolved module-level names below.
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
        "VALID_INVALID_CONFIG_JSON",
        str(CONFIG_DIR / "valid_invalid_config.json"),
    )
).resolve()

# ---------------------------------------------------------------------------
# Defaults (also used when JSON is missing or for Reset)
# ---------------------------------------------------------------------------

_DEFAULT_VALUES: Dict[str, Any] = {
    "JOURNEY_TYPES_TO_DROP": [
        "ANNUALPASS",
        "ANNUALPASS:LocalSingleC",
        "ANNUALPASS:RETURN",
        "ANNUALPASS:SINGLE",
        "GlobalExempt",
        "ANNUALPASS_RETURN",
        "ANNUALPASS_FULL",
        "Annual Pass Single",
        "GlobalExemption",
        "Global Exemption",
        "Annual Pass Trip",
        "0",
        "0.0000",
        "Rejected Trip",
        "Rejected Trip Commercial",
        "Annual Pass Trip Personal",
        "Rejected Trip Personal",
        "EXEMPTED",
    ],
    "STATUS_COLUMN_CANDIDATES": [
        "Settlement Type",
        "Txn Status",
        "Status",
        "Transaction Status",
        "transaction status",
    ],
    "ACCEPTED_STATUS_VALUES": [
        "ACCEPTED",
        "Accepted",
        "SETTLED",
    ],
    "LOCAL_CONTI_SINGLE_JOURNEY_TYPES": [
        "Local Cont.",
        "Local Single",
        "Local Vehicle(Commercial)",
        "LocalSingleC",
        "Local Single Journey Commercial",
        "Local Vehicle Commercial",
    ],
    "CONT_JOURNEY_TYPES": [
        "Return Journey",
        "RETURN",
        "RETURN PASS",
        "Return",
        "Cont. Journey",
        "Return Journey Personal",
    ],
    "RATE_SHEET_JOURNEY_COLUMN_RENAMES": {
        "Return Journey": "Cont. Journey",
        "Local Conti/Local Single": "Local Conti/Single",
        # Sheet column aliases → canonical Attribute used after melt / rate lookup.
        "Single journey": "Single Journey",
    },
    # Aliases for the three fixed identity columns on the rates sheet.
    # Canonical keys cannot be removed; users add alternate header spellings under each.
    "RATE_SHEET_ID_COLUMN_ALIASES": {
        "TC Class": [
            "TCClass",
            "TC_Class",
            "T C Class",
            "Vehicle TC Class",
        ],
        "Weight/Capacity": [
            "Weight Capacity",
            "Weight/Capacity (Kgs)",
            "Weight Capacity (Kgs)",
            "Weight",
            "Capacity",
        ],
        "Vehicle Class": [
            "VehicleClass",
            "Veh Class",
            "Vehicle Class Desc",
            "Veh. Class",
        ],
    },
    "LIFECYCLE_EXCLUDED_VEHICLE_CLASSES": [
        "BUS",
        " BUS",
        "CAMPER VAN / TRAILER",
        "CASH VAN",
        "Error",
        "MAXI CAB",
        "MOTOR CAB",
        "OMNI BUS",
        "OMNIBUS",
        "Invalid",
    ],
    "BUS_VEHICLE_CLASSES": [
        " BUS",
        "BUS",
        "OMNI BUS",
        "OMNIBUS",
    ],
    "HEAVY_SPECIAL_VEHICLE_CLASSES": [
        "CONSTRUCTION EQUIPMENT VEHICLE",
        "EARTH MOVING EQUIPMENT",
        "VEHICLE FITTED WITH AIR GENERATOR",
        "VEHICLE FITTED WITH COMPRESSOR",
        "VEHICLE FITTED WITH RIG",
        "EXCAVATOR",
        "TREE TRIMMING VEHICLE",
        "TOWER WAGONS",
    ],
    "CRANE_MOUNTED_VEHICLE_CLASS": "CRANE MOUNTED VEHICLE",
    "NPCI_TO_TC_CLASS": {
        "Car": [
            "Car / Jeep / Van",
            "Car/Jeep/Van",
            "VC4 - Car / Jeep / Van",
            "VC4",
            "Tata Ace and Similar mini Light Commercial Vehicle",
            "Tata Ace or Similar Mini LCV",
            "VC20",
            "VC20 - Tata Ace or similar mini LCV",
        ],
        "LCV": [
            "Mini-Bus",
            "LCV",
            "Light Commercial vehicle 2-axle",
            "VC5",
            "VC9",
        ],
        "Trk 2 Axle": [
            "Truck 2 - axle",
            "Bus 2-axle",
            "VC10",
            "VC7",
            "VC10 - Truck 2 Axle",
            "VC7 - Bus 2 Axle",
        ],
        "Truck 3 axle": [
            "Truck 3 - axle",
            "Truck 3-Axle",
            "VC8",
            "VC11",
            "Bus 3-axle",
            "VC11 - Truck 3 Axle",
            "VC8 - Bus 3 Axle",
        ],
        "MAV": [
            "Truck 4 - axle",
            "Truck 4-Axle",
            "Truck 5 - axle",
            "Truck 6 - axle",
            "VC12",
            "VC13",
            "VC14",
        ],
    },
    "ZERO_SETTLEMENT_REMOVE_JOURNEY_TYPES": [
        "Cont. Journey",
        "Single Journey",
        "Local Conti/Single",
        "Local Cont.",
        "Local Single",
        "Return Journey",
        "SINGLE",
        "RETURN",
        "LocalSingleC",
        "FULL",
        "LOCAL TRIP",
        "SINGLE JOURNEY",
        "RETURN PASS",
        "Return",
        "Single",
        "LocalSingleC",
        "Local Vehicle(Commercial)",
        "Local Vehicle Commercial",
    ],
    "DISCOUNT_PASS_JOURNEY_TYPES": [
        "DISCOUNTED",
        "DISCOUNTMP",
        "Monthly Pass-50 trips",
        "Monthly Pass Comm",
        "Monthly Pass Non Comm",
        "Local Non Commercial",
        "Local20KM",
        "MonthlyExempted",
        "Discount Pass",
        "Monthly Pass",
        "Local Vehicle(Commercial)",
        "Local Pass Non Comm",
        "NON_FIN",
        "Monthly Pass 20 Km",
    ],
    "INVALID_EXCLUDED_VEHICLE_CLASSES": [
        "BUS",
        "CONSTRUCTION EQUIPMENT VEHICLE",
        "OMNI BUS",
        "TOWER WAGONS",
        "VEHICLE FITTED WITH AIR GENERATOR",
        "VEHICLE FITTED WITH COMPRESSOR",
        "VEHICLE FITTED WITH RIG",
    ],
}

# Fixed output labels used by normalize_journey_type() — not editable in the portal.
UPDATED_JOURNEY_TYPE_LOCAL = "Local Conti/Single"
UPDATED_JOURNEY_TYPE_CONT = "Cont. Journey"
UPDATED_JOURNEY_TYPE_SINGLE = "Single Journey"

_FIXED_CONFIG_KEYS = frozenset(
    {
        "UPDATED_JOURNEY_TYPE_LOCAL",
        "UPDATED_JOURNEY_TYPE_CONT",
        "UPDATED_JOURNEY_TYPE_SINGLE",
    }
)

# UI / API schema: top-level keys cannot be removed.
CONFIG_SCHEMA: Dict[str, Dict[str, Any]] = {
    "JOURNEY_TYPES_TO_DROP": {
        "type": "list",
        "label": "Journey types to drop",
        "description": "Dropped before merge / rate logic.",
        "order": 1,
    },
    "STATUS_COLUMN_CANDIDATES": {
        "type": "list",
        "label": "Status column candidates",
        "description": "Possible status / settlement columns on lifecycle file.",
        "order": 2,
    },
    "ACCEPTED_STATUS_VALUES": {
        "type": "list",
        "label": "Accepted status values",
        "description": "Keep only rows with these status values.",
        "order": 3,
    },
    "LOCAL_CONTI_SINGLE_JOURNEY_TYPES": {
        "type": "list",
        "label": "Local Conti/Single journey types",
        "description": 'Mapped to fixed label "Local Conti/Single".',
        "order": 4,
    },
    "CONT_JOURNEY_TYPES": {
        "type": "list",
        "label": "Cont. Journey types",
        "description": 'Mapped to fixed label "Cont. Journey".',
        "order": 5,
    },
    "RATE_SHEET_JOURNEY_COLUMN_RENAMES": {
        "type": "dict",
        "label": "Rate sheet journey column renames",
        "description": (
            "Map any rates-sheet journey column name → canonical Attribute: "
            '"Single Journey", "Cont. Journey", or "Local Conti/Single". '
            "Add a row when a plaza uses a different header "
            '(e.g. "Single" → "Single Journey", '
            '"Local Conti/Local Single" → "Local Conti/Single").'
        ),
        "order": 6,
        "editable_keys": True,
    },
    "RATE_SHEET_ID_COLUMN_ALIASES": {
        "type": "dict_of_lists",
        "label": "Rate sheet ID column aliases",
        "description": (
            "Aliases for the three rates identity columns. Canonical keys are fixed; "
            "add alternate spellings under TC Class, Weight/Capacity, and Vehicle Class."
        ),
        "order": 7,
        "editable_keys": False,
    },
    "LIFECYCLE_EXCLUDED_VEHICLE_CLASSES": {
        "type": "list",
        "label": "Lifecycle excluded vehicle classes",
        "description": "Excluded from non-bus lifecycle processing.",
        "order": 8,
    },
    "BUS_VEHICLE_CLASSES": {
        "type": "list",
        "label": "Bus vehicle classes",
        "order": 9,
    },
    "HEAVY_SPECIAL_VEHICLE_CLASSES": {
        "type": "list",
        "label": "Heavy / special vehicle classes",
        "description": "Use MAV Rate 1 by journey type.",
        "order": 10,
    },
    "CRANE_MOUNTED_VEHICLE_CLASS": {
        "type": "string",
        "label": "Crane mounted vehicle class",
        "description": "Uses LCV Rate 1 instead of MAV.",
        "order": 11,
    },
    "NPCI_TO_TC_CLASS": {
        "type": "dict_of_lists",
        "label": "NPCI → TC class mapping",
        "description": "TC class keys are fixed; edit NPCI labels under each.",
        "order": 12,
        "editable_keys": False,
    },
    "ZERO_SETTLEMENT_REMOVE_JOURNEY_TYPES": {
        "type": "list",
        "label": "Zero-settlement journey types to remove",
        "order": 13,
    },
    "DISCOUNT_PASS_JOURNEY_TYPES": {
        "type": "list",
        "label": "Discount / pass journey types",
        "description": "Override settlement with Single Journey rate.",
        "order": 14,
    },
    "INVALID_EXCLUDED_VEHICLE_CLASSES": {
        "type": "list",
        "label": "Invalid excluded vehicle classes",
        "description": "Dropped from final invalid export.",
        "order": 15,
    },
}


def get_default_values() -> Dict[str, Any]:
    return deepcopy(_DEFAULT_VALUES)


def get_config_schema_for_api() -> List[Dict[str, Any]]:
    items = []
    for key, meta in CONFIG_SCHEMA.items():
        items.append(
            {
                "key": key,
                "type": meta["type"],
                "label": meta["label"],
                "description": meta.get("description", ""),
                "order": meta.get("order", 999),
                "editable_keys": bool(meta.get("editable_keys", False)),
            }
        )
    items.sort(key=lambda row: (row["order"], row["key"]))
    return items


def _normalize_list(value: Any, field_name: str) -> List[str]:
    if not isinstance(value, list):
        raise ValueError(f"{field_name} must be a list of strings.")
    out: List[str] = []
    for item in value:
        if item is None:
            continue
        # Keep leading/trailing spaces — some source files use values like " BUS".
        text = str(item)
        if text.strip() == "":
            continue
        out.append(text)
    return out


def _normalize_string(value: Any, field_name: str) -> str:
    if value is None:
        raise ValueError(f"{field_name} must be a non-empty string.")
    text = str(value)
    if text.strip() == "":
        raise ValueError(f"{field_name} must be a non-empty string.")
    return text


def _normalize_dict(value: Any, field_name: str) -> Dict[str, str]:
    if not isinstance(value, dict):
        raise ValueError(f"{field_name} must be an object of string keys/values.")
    out: Dict[str, str] = {}
    for raw_key, raw_val in value.items():
        key = str(raw_key)
        val = "" if raw_val is None else str(raw_val)
        if key.strip() == "" or val.strip() == "":
            continue
        out[key] = val
    return out


def _normalize_dict_of_lists(
    value: Any,
    field_name: str,
    required_keys: Optional[List[str]] = None,
) -> Dict[str, List[str]]:
    if not isinstance(value, dict):
        raise ValueError(f"{field_name} must be an object of string → list.")
    out: Dict[str, List[str]] = {}
    for raw_key, raw_list in value.items():
        key = str(raw_key).strip()
        if not key:
            continue
        out[key] = _normalize_list(raw_list, f"{field_name}.{key}")
    if required_keys is not None:
        missing = [k for k in required_keys if k not in out]
        if missing:
            raise ValueError(
                f"{field_name} is missing required key(s): {', '.join(missing)}"
            )
        # Drop unknown keys when keys are fixed.
        out = {k: out[k] for k in required_keys}
    return out


def validate_and_normalize_values(payload: Dict[str, Any]) -> Dict[str, Any]:
    if not isinstance(payload, dict):
        raise ValueError("Config payload must be a JSON object.")

    # Ignore fixed Python-only keys if an older JSON/payload still includes them.
    payload = {k: v for k, v in payload.items() if k not in _FIXED_CONFIG_KEYS}

    unknown = sorted(set(payload) - set(CONFIG_SCHEMA))
    if unknown:
        raise ValueError(
            "Unknown config key(s) are not allowed: " + ", ".join(unknown)
        )

    missing = sorted(set(CONFIG_SCHEMA) - set(payload))
    if missing:
        raise ValueError(
            "Missing required config key(s): " + ", ".join(missing)
        )

    normalized: Dict[str, Any] = {}
    for key, meta in CONFIG_SCHEMA.items():
        field_type = meta["type"]
        raw = payload[key]
        if field_type == "list":
            normalized[key] = _normalize_list(raw, key)
        elif field_type == "string":
            normalized[key] = _normalize_string(raw, key)
        elif field_type == "dict":
            normalized[key] = _normalize_dict(raw, key)
        elif field_type == "dict_of_lists":
            required_keys = None
            if not meta.get("editable_keys", False):
                required_keys = list(_DEFAULT_VALUES[key].keys())
            normalized[key] = _normalize_dict_of_lists(raw, key, required_keys)
        else:
            raise ValueError(f"Unsupported schema type for {key}: {field_type}")
    return normalized


def _read_json_file(path: Path) -> Optional[Dict[str, Any]]:
    if not path.is_file():
        return None
    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError(f"{path.name} must contain a JSON object.")
    return data


def load_config_values() -> Tuple[Dict[str, Any], bool]:
    """
    Return (values, from_file).

    Merges JSON over defaults for known keys only. Invalid/partial JSON falls
    back to defaults for bad keys where possible; fully invalid files raise.
    """
    defaults = get_default_values()
    try:
        stored = _read_json_file(CONFIG_JSON_PATH)
    except (OSError, json.JSONDecodeError, ValueError) as exc:
        raise ValueError(f"Could not read {CONFIG_JSON_PATH.name}: {exc}") from exc

    if stored is None:
        return defaults, False

    merged = deepcopy(defaults)
    for key in CONFIG_SCHEMA:
        if key in stored:
            merged[key] = stored[key]

    try:
        return validate_and_normalize_values(merged), True
    except ValueError:
        # Keep process runnable: fall back to defaults if overlay is corrupt.
        return defaults, False


def save_config_values(payload: Dict[str, Any]) -> Dict[str, Any]:
    normalized = validate_and_normalize_values(payload)
    CONFIG_JSON_PATH.parent.mkdir(parents=True, exist_ok=True)

    fd, tmp_name = tempfile.mkstemp(
        prefix="valid_invalid_config_",
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

    _apply_values_to_module(normalized)
    return normalized


def reset_config_to_defaults() -> Dict[str, Any]:
    defaults = get_default_values()
    return save_config_values(defaults)


def ensure_config_json_exists() -> Dict[str, Any]:
    """Create JSON from defaults when missing; return current values."""
    values, from_file = load_config_values()
    if not from_file:
        save_config_values(values)
        return values
    _apply_values_to_module(values)
    return values


def _apply_values_to_module(values: Dict[str, Any]) -> None:
    globals().update(
        {
            "JOURNEY_TYPES_TO_DROP": values["JOURNEY_TYPES_TO_DROP"],
            "STATUS_COLUMN_CANDIDATES": values["STATUS_COLUMN_CANDIDATES"],
            "ACCEPTED_STATUS_VALUES": values["ACCEPTED_STATUS_VALUES"],
            "LOCAL_CONTI_SINGLE_JOURNEY_TYPES": values["LOCAL_CONTI_SINGLE_JOURNEY_TYPES"],
            "CONT_JOURNEY_TYPES": values["CONT_JOURNEY_TYPES"],
            # Always keep these as fixed Python constants (not portal-editable).
            "UPDATED_JOURNEY_TYPE_LOCAL": UPDATED_JOURNEY_TYPE_LOCAL,
            "UPDATED_JOURNEY_TYPE_CONT": UPDATED_JOURNEY_TYPE_CONT,
            "UPDATED_JOURNEY_TYPE_SINGLE": UPDATED_JOURNEY_TYPE_SINGLE,
            "RATE_SHEET_JOURNEY_COLUMN_RENAMES": values["RATE_SHEET_JOURNEY_COLUMN_RENAMES"],
            "RATE_SHEET_ID_COLUMN_ALIASES": values["RATE_SHEET_ID_COLUMN_ALIASES"],
            "LIFECYCLE_EXCLUDED_VEHICLE_CLASSES": values["LIFECYCLE_EXCLUDED_VEHICLE_CLASSES"],
            "BUS_VEHICLE_CLASSES": values["BUS_VEHICLE_CLASSES"],
            "HEAVY_SPECIAL_VEHICLE_CLASSES": values["HEAVY_SPECIAL_VEHICLE_CLASSES"],
            "CRANE_MOUNTED_VEHICLE_CLASS": values["CRANE_MOUNTED_VEHICLE_CLASS"],
            "NPCI_TO_TC_CLASS": values["NPCI_TO_TC_CLASS"],
            "ZERO_SETTLEMENT_REMOVE_JOURNEY_TYPES": values["ZERO_SETTLEMENT_REMOVE_JOURNEY_TYPES"],
            "DISCOUNT_PASS_JOURNEY_TYPES": values["DISCOUNT_PASS_JOURNEY_TYPES"],
            "INVALID_EXCLUDED_VEHICLE_CLASSES": values["INVALID_EXCLUDED_VEHICLE_CLASSES"],
        }
    )


def map_npci_to_tc_class(npci):
    """Map NPCI Class Desc to rates TC Class using NPCI_TO_TC_CLASS."""
    for tc_class, npci_labels in NPCI_TO_TC_CLASS.items():
        if npci in npci_labels:
            return tc_class
    return None


def normalize_journey_type(journey_type):
    """Map raw Journey Type → updated journey type used in rates lookups."""
    if journey_type in LOCAL_CONTI_SINGLE_JOURNEY_TYPES:
        return UPDATED_JOURNEY_TYPE_LOCAL
    if journey_type in CONT_JOURNEY_TYPES:
        return UPDATED_JOURNEY_TYPE_CONT
    return UPDATED_JOURNEY_TYPE_SINGLE


# Resolve module-level exports for Valid_Invalid_Full_Process imports.
_apply_values_to_module(load_config_values()[0])
