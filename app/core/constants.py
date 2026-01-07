"""Application constants for GLC Dashboard."""
from datetime import datetime, date, timezone, timedelta
from zoneinfo import ZoneInfo

# Australian timezone (Brisbane - no DST)
AUSTRALIA_TZ = ZoneInfo("Australia/Brisbane")


def get_australia_now() -> datetime:
    """Get current datetime in Australian timezone."""
    return datetime.now(AUSTRALIA_TZ)


def get_australia_today() -> date:
    """Get current date in Australian timezone.

    This should be used instead of date.today() to ensure the dashboard
    displays the correct day for Australian business hours, regardless
    of the server's timezone (e.g., UTC on Render.com).
    """
    return get_australia_now().date()


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
