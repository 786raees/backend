"""Application constants for GLC Dashboard."""

# Truck capacity in SQM
TRUCK1_CAPACITY = 500
TRUCK2_CAPACITY = 600

TRUCK_CAPACITIES = {
    1: TRUCK1_CAPACITY,
    2: TRUCK2_CAPACITY,
}

# Valid turf varieties
VALID_TURF_VARIETIES = [
    "Empire Zoysia",
    "Sir Walter",
    "Wintergreen Couch",
    "AussiBlue Couch",
    "Summerland Buffalo",
]

# Service type color mapping (for frontend reference)
SERVICE_TYPE_COLORS = {
    "SL": "#3B82F6",  # Blue - Supply and Lay
    "SD": "#57924b",  # Green - Supply and Delivery
    "P": "#EF4444",   # Red - Project
}

# Day name mapping for Excel to date matching
DAY_NAME_MAP = {
    "monday": 0,
    "tuesday": 1,
    "wednesday": 2,
    "thursday": 3,
    "friday": 4,
    "saturday": 5,
    "sunday": 6,
}
