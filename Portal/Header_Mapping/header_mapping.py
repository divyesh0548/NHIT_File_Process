"""
Central place for header detection keywords and column/header mappings used by Portal jobs.

Add new lists or dicts here as workflows grow; import them from `app` or scripts as needed.
"""

# Life cycle merge (`Scripts/Life_cycle_merge.merge_files_in_folder`): keywords used to locate
# the header row (at least 3 must match in the script logic).
LIFE_CYCLE_MERGE_HEADER_KEYWORDS = [
    "Agency Txn Id",
    "Settlement Amount",
    "Plaza ID",
    "Violation Amts",
]

VALID_INVALID_LOOKUP_HEADER_MAPPING = {
    "Journey Type": [
        "Journey Type",
        "MOP",
        "Fare Type",
        "TRIPTYPEDISCRIPTION",
        "Trip Type Description",
        "fare type",
    ],
    "Net Settlement Amt": [
        "Net Settlement Amt",
        "SETTLED Amount (Rs.)",
        "Settled Amount (Rs.)",
        "Net Settlement Amount",
        "accepted amount",
    ],
    "NPCI Class Desc": [
        "NPCI Class Desc",
        "Mapper Vehicle Class",
        "Vehicle Class",
        "mapper vehicle class",
    ],
    "Veh Reg No.": [
        "Veh Reg No.",
        "NPCI VRN",
        "vehicle registration number",
        "Vehicle Registration Number",
        "Vehicle Reg No",
        "Vehicle Reg No.",
        "Veh Reg No",
        "Registration Number",
    ],
}

VALID_INVALID_LOOKUP_REQUIRED_COLUMNS = list(VALID_INVALID_LOOKUP_HEADER_MAPPING.keys())

HEADER_MAPPINGS = {
    "valid_invalid_lookup": VALID_INVALID_LOOKUP_HEADER_MAPPING,
}
