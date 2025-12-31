"""Mock data generator for development."""
from datetime import date, datetime, timedelta, timezone
from typing import List
import random

from app.models.schedule import Delivery, TruckData, DaySchedule, ScheduleResponse

# Constants
TRUCK1_CAPACITY = 500
TRUCK2_CAPACITY = 600

TURF_VARIETIES = [
    "Empire Zoysia",
    "Sir Walter",
    "Wintergreen Couch",
    "AussiBlue Couch",
    "Summerland Buffalo"
]

SUBURBS = [
    "Pimpama",
    "Coomera",
    "Ipswich",
    "Yatala",
    "Gold Coast",
    "Logan",
    "Brisbane"
]

SERVICE_TYPES = ["SL", "SD", "P"]


def generate_delivery() -> Delivery:
    """Generate a random delivery."""
    sqm = random.choice([50, 100, 150, 200, 250, 300])
    # Approximately 50 SQM per pallet
    pallets = round(sqm / 50, 1)

    return Delivery(
        sqm=sqm,
        pallets=pallets,
        variety=random.choice(TURF_VARIETIES),
        suburb=random.choice(SUBURBS),
        service_type=random.choice(SERVICE_TYPES)
    )


def generate_truck_data(capacity: int, num_deliveries: int) -> TruckData:
    """Generate truck data with random deliveries."""
    deliveries = [generate_delivery() for _ in range(num_deliveries)]

    sqm_total = sum(d.sqm for d in deliveries)
    pallet_total = sum(d.pallets for d in deliveries)
    available_sqm = max(0, capacity - sqm_total)

    return TruckData(
        deliveries=deliveries,
        sqm_total=sqm_total,
        pallet_total=pallet_total,
        capacity=capacity,
        available_sqm=available_sqm
    )


def get_ordinal_suffix(day: int) -> str:
    """Get ordinal suffix for a day number."""
    if 11 <= day <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")


def format_day_name(d: date) -> str:
    """Format date as 'Mon 22nd Dec'."""
    day = d.day
    suffix = get_ordinal_suffix(day)
    return d.strftime(f"%a {day}{suffix} %b")


def get_next_business_days(start_date: date, count: int) -> List[date]:
    """Get the next N business days (Mon-Fri) starting from start_date."""
    business_days = []
    current = start_date

    while len(business_days) < count:
        # Monday = 0, Sunday = 6
        if current.weekday() < 5:  # Mon-Fri
            business_days.append(current)
        current += timedelta(days=1)

    return business_days


def generate_mock_schedule() -> ScheduleResponse:
    """Generate a complete mock schedule response."""
    today = date.today()
    business_days = get_next_business_days(today, 10)

    days = []
    for i, day_date in enumerate(business_days):
        # Week 2 starts at index 5 (6th day)
        is_week_two = i >= 5

        # Random number of deliveries (0-6 per truck)
        # Make some days busier, some lighter
        t1_deliveries = random.randint(0, 5)
        t2_deliveries = random.randint(0, 6)

        day_schedule = DaySchedule(
            date=day_date,
            day_name=format_day_name(day_date),
            is_week_two=is_week_two,
            truck1=generate_truck_data(TRUCK1_CAPACITY, t1_deliveries),
            truck2=generate_truck_data(TRUCK2_CAPACITY, t2_deliveries)
        )
        days.append(day_schedule)

    return ScheduleResponse(
        success=True,
        generated_at=datetime.now(timezone.utc).isoformat(),
        source="mock",
        days=days
    )
