"""
Configurable keyword lists and mappings for Valid/Invalid Lookup.

Edit values here — Valid_Invalid_Full_Process.py imports from this module.
Items correspond to Config_cadidates.txt.
"""

# 1. Journey types dropped before merge / rate logic
JOURNEY_TYPES_TO_DROP = [
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
]

# 2. Possible status / settlement columns on lifecycle file
STATUS_COLUMN_CANDIDATES = [
    "Settlement Type",
    "Txn Status",
    "Status",
    "Transaction Status",
    "transaction status",
]

# 3. Accepted status values (keep only these rows)
ACCEPTED_STATUS_VALUES = [
    "ACCEPTED",
    "Accepted",
    "SETTLED",
]

# 4. Journey Type value → updated journey type (used for rates lookup)
# Values not in these lists default to "Single Journey"
LOCAL_CONTI_SINGLE_JOURNEY_TYPES = [
    "Local Cont.",
    "Local Single",
    "Local Vehicle(Commercial)",
    "LocalSingleC",
    "Local Single Journey Commercial",
    "Local Vehicle Commercial",
]

CONT_JOURNEY_TYPES = [
    "Return Journey",
    "RETURN",
    "RETURN PASS",
    "Return",
    "Cont. Journey",
    "Return Journey Personal",
]

UPDATED_JOURNEY_TYPE_LOCAL = "Local Conti/Single"
UPDATED_JOURNEY_TYPE_CONT = "Cont. Journey"
UPDATED_JOURNEY_TYPE_SINGLE = "Single Journey"

# 5. Rates sheet column renames before melt (wide → Attribute names)
RATE_SHEET_JOURNEY_COLUMN_RENAMES = {
    "Return Journey": "Cont. Journey",
    "Local Conti/Local Single": "Local Conti/Single",
}

# 6. Vehicle classes excluded from non-bus lifecycle processing
#    (buses are handled separately by vehicle_bus)
LIFECYCLE_EXCLUDED_VEHICLE_CLASSES = [
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
]

# 7. Bus vehicle classes (processed by vehicle_bus)
BUS_VEHICLE_CLASSES = [
    " BUS",
    "BUS",
    "OMNI BUS",
    "OMNIBUS",
]

# 8. Heavy / special classes → Rate 1 as MAV by journey type
HEAVY_SPECIAL_VEHICLE_CLASSES = [
    "CONSTRUCTION EQUIPMENT VEHICLE",
    "EARTH MOVING EQUIPMENT",
    "VEHICLE FITTED WITH AIR GENERATOR",
    "VEHICLE FITTED WITH COMPRESSOR",
    "VEHICLE FITTED WITH RIG",
    "EXCAVATOR",
    "TREE TRIMMING VEHICLE",
    "TOWER WAGONS",
]

# Crane uses LCV Rate 1 instead of MAV
CRANE_MOUNTED_VEHICLE_CLASS = "CRANE MOUNTED VEHICLE"

# 9. NPCI Class Desc → TC Class (for Single Journey rate override)
# Keys are TC Class names used in rates sheet; values are NPCI labels
NPCI_TO_TC_CLASS = {
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
}

# 10. Journey types where zero Net Settlement Amt rows are removed from invalid
ZERO_SETTLEMENT_REMOVE_JOURNEY_TYPES = [
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
]

# 11. Discount / pass journey types → override settlement with Single Journey rate
DISCOUNT_PASS_JOURNEY_TYPES = [
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
]

# 12. Vehicle classes dropped from final invalid export
INVALID_EXCLUDED_VEHICLE_CLASSES = [
    "BUS",
    "CONSTRUCTION EQUIPMENT VEHICLE",
    "OMNI BUS",
    "TOWER WAGONS",
    "VEHICLE FITTED WITH AIR GENERATOR",
    "VEHICLE FITTED WITH COMPRESSOR",
    "VEHICLE FITTED WITH RIG",
]


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
