"""Schedule builder service for assembling schedule data from Excel."""
import logging
from datetime import date, datetime, timedelta, timezone
from typing import Dict, List, Optional, Tuple

from app.models.schedule import Delivery, TruckData, DaySchedule, ScheduleResponse
from app.core.constants import TRUCK1_CAPACITY, TRUCK2_CAPACITY, get_australia_today
from app.services.excel.excel_parser import ExcelDeliveryRow

logger = logging.getLogger(__name__)


def get_ordinal_suffix(day: int) -> str:
    """Get ordinal suffix for a day number (st, nd, rd, th)."""
    if 11 <= day <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")


def format_display_date(d: date) -> str:
    """Format date as 'Mon 22nd Dec'."""
    day = d.day
    suffix = get_ordinal_suffix(day)
    return d.strftime(f"%a {day}{suffix} %b")


def get_next_business_days(start_date: date, count: int = 10) -> List[date]:
    """Get the next N business days (Mon-Fri) starting from start_date.

    Args:
        start_date: The date to start from.
        count: Number of business days to generate.

    Returns:
        List of business day dates.
    """
    business_days = []
    current = start_date

    while len(business_days) < count:
        # Monday = 0, Sunday = 6
        if current.weekday() < 5:  # Mon-Fri
            business_days.append(current)
        current += timedelta(days=1)

    return business_days


def normalize_day_name(day_name: str) -> Tuple[str, int]:
    """Normalize day name and extract week number.

    Args:
        day_name: Day name from Excel (e.g., "Monday", "Monday W2", "Mon W2")

    Returns:
        Tuple of (normalized day name, week number 1 or 2)
    """
    if not day_name:
        return ("", 1)

    day_map = {
        "mon": "Monday",
        "tue": "Tuesday",
        "wed": "Wednesday",
        "thu": "Thursday",
        "fri": "Friday",
        "sat": "Saturday",
        "sun": "Sunday",
    }

    # Check for week 2 indicator
    week = 1
    day_lower = day_name.lower().strip()
    if "w2" in day_lower or "week 2" in day_lower or "week2" in day_lower:
        week = 2
        # Remove week indicator
        day_lower = day_lower.replace("week 2", "").replace("week2", "").replace("w2", "").strip()

    # Handle short forms
    short = day_lower[:3]
    if short in day_map:
        return (day_map[short], week)

    # Handle full names
    return (day_name.split()[0].title(), week)


class ScheduleBuilder:
    """Builds schedule response from Excel delivery data."""

    def __init__(self):
        """Initialize the schedule builder."""
        self.truck1_capacity = TRUCK1_CAPACITY
        self.truck2_capacity = TRUCK2_CAPACITY

    def _create_date_to_day_map(self, business_days: List[date]) -> Dict[str, date]:
        """Create a mapping from day names to dates.

        For each business day, maps its weekday name to its date.
        If there are multiple dates with the same weekday (e.g., two Mondays),
        the first one in Week 1 takes priority.
        """
        day_map = {}

        for d in business_days:
            day_name = d.strftime("%A")
            if day_name not in day_map:
                day_map[day_name] = d

        return day_map

    def _group_deliveries_by_date(
        self,
        deliveries: List[ExcelDeliveryRow],
        business_days: List[date],
    ) -> Dict[date, List[ExcelDeliveryRow]]:
        """Group deliveries by their matching business day date.

        Uses week_start_date from each delivery if available for accurate
        date mapping. Falls back to day name matching if week_start_date
        is not set.

        Args:
            deliveries: List of parsed delivery rows.
            business_days: List of business day dates.

        Returns:
            Dictionary mapping dates to lists of deliveries.
        """
        # Initialize all dates with empty lists
        grouped: Dict[date, List[ExcelDeliveryRow]] = {d: [] for d in business_days}
        business_days_set = set(business_days)

        # Create day name to date mapping for fallback (supports Week 1 and Week 2)
        day_to_dates = {}
        for d in business_days:
            day_name = d.strftime("%A")
            if day_name not in day_to_dates:
                day_to_dates[day_name] = []
            day_to_dates[day_name].append(d)

        for delivery in deliveries:
            # Try to get actual date from week_start_date
            actual_date = delivery.get_actual_date()

            if actual_date:
                # Week start date is set - use strict date matching only
                if actual_date in business_days_set:
                    grouped[actual_date].append(delivery)
                    logger.debug(f"Mapped '{delivery.day}' to {actual_date} using week_start_date")
                else:
                    # Delivery is outside the business days window - skip it
                    logger.debug(f"Skipping delivery '{delivery.suburb}' on {actual_date} - outside 10-day window")
            else:
                # No week_start_date - use fallback day name matching (legacy behavior)
                normalized_day, week = normalize_day_name(delivery.day)
                dates_for_day = day_to_dates.get(normalized_day, [])

                if dates_for_day:
                    # Assign to Week 1 (index 0) or Week 2 (index 1)
                    week_index = week - 1  # Convert to 0-based index
                    if week_index < len(dates_for_day):
                        target_date = dates_for_day[week_index]
                        grouped[target_date].append(delivery)
                    else:
                        # Week 2 requested but only one occurrence - use first
                        target_date = dates_for_day[0]
                        grouped[target_date].append(delivery)
                        logger.warning(f"Week {week} requested for '{delivery.day}' but only {len(dates_for_day)} occurrence(s) available")
                else:
                    logger.warning(f"No matching date for day '{delivery.day}'")

        return grouped

    def _build_truck_data(
        self,
        deliveries: List[ExcelDeliveryRow],
        capacity: int,
        truck_name: str,
    ) -> TruckData:
        """Build TruckData from a list of deliveries.

        Args:
            deliveries: List of delivery rows for this truck.
            capacity: Truck capacity in SQM.
            truck_name: Truck identifier ("TRUCK 1" or "TRUCK 2").

        Returns:
            TruckData with calculated totals.
        """
        # Convert ExcelDeliveryRow to Delivery models WITH truck context
        delivery_models = [d.to_delivery(truck=truck_name) for d in deliveries]

        sqm_total = sum(d.sqm for d in delivery_models)
        pallet_total = sum(d.pallets for d in delivery_models)
        laying_cost_total = sum(d.laying_cost for d in delivery_models)
        delivery_fee_total = sum(d.delivery_fee for d in delivery_models)
        laying_fee_total = sum(d.laying_fee for d in delivery_models)
        available_sqm = capacity - sqm_total  # Can be negative for overflow

        return TruckData(
            deliveries=delivery_models,
            sqm_total=round(sqm_total, 2),
            pallet_total=round(pallet_total, 2),
            laying_cost_total=round(laying_cost_total, 2),
            delivery_fee_total=round(delivery_fee_total, 2),
            laying_fee_total=round(laying_fee_total, 2),
            capacity=capacity,
            available_sqm=round(available_sqm, 2),
        )

    def build_schedule(
        self,
        truck1_deliveries: List[ExcelDeliveryRow],
        truck2_deliveries: List[ExcelDeliveryRow],
        start_date: Optional[date] = None,
        source: str = "google_sheets",
    ) -> ScheduleResponse:
        """Build a complete schedule response from truck delivery data.

        Args:
            truck1_deliveries: Deliveries for Truck 1.
            truck2_deliveries: Deliveries for Truck 2.
            start_date: Optional start date (defaults to today).
            source: Data source identifier ("google_sheets", "onedrive", "mock").

        Returns:
            Complete ScheduleResponse with all days populated.
        """
        if start_date is None:
            start_date = get_australia_today()

        # Generate 10 business days
        business_days = get_next_business_days(start_date, 10)

        # Group deliveries by date for each truck
        t1_by_date = self._group_deliveries_by_date(truck1_deliveries, business_days)
        t2_by_date = self._group_deliveries_by_date(truck2_deliveries, business_days)

        # Build day schedules
        days = []
        for i, day_date in enumerate(business_days):
            is_week_two = i >= 5  # Days 6-10 are week 2

            t1_data = self._build_truck_data(
                t1_by_date.get(day_date, []),
                self.truck1_capacity,
                "TRUCK 1",
            )
            t2_data = self._build_truck_data(
                t2_by_date.get(day_date, []),
                self.truck2_capacity,
                "TRUCK 2",
            )

            day_schedule = DaySchedule(
                date=day_date,
                day_name=format_display_date(day_date),
                day_of_week=day_date.strftime('%A'),
                is_week_two=is_week_two,
                truck1=t1_data,
                truck2=t2_data,
            )
            days.append(day_schedule)

        return ScheduleResponse(
            success=True,
            generated_at=datetime.now(timezone.utc).isoformat(),
            source=source,
            days=days,
        )

    def build_empty_schedule(
        self,
        start_date: Optional[date] = None,
    ) -> ScheduleResponse:
        """Build a schedule with no deliveries (all zeros).

        Useful when OneDrive data is unavailable or empty.
        """
        return self.build_schedule([], [], start_date)


# Global instance
schedule_builder = ScheduleBuilder()
