"""
Central place for header detection keywords and column/header mappings used by Portal jobs.

LC/ETC merge header keywords are stored in the `nhit_file_process` database table.
See `db.nhit_file_process` for load/save helpers.
"""

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
