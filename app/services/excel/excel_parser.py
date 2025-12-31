"""Excel data parser for transforming worksheet data into structured models."""
import logging
from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Any, List, Optional

from app.models.schedule import Delivery

logger = logging.getLogger(__name__)

# Column mapping for truck worksheets (0-indexed)
# Based on the Excel structure defined in the project plan
TRUCK_COLUMN_MAP = {
    0: "day",           # Column A: Day (Monday, Tuesday, etc.)
    1: "slot",          # Column B: Slot (1-6)
    2: "variety",       # Column C: Turf variety
    3: "suburb",        # Column D: Delivery suburb
    4: "service_type",  # Column E: SL, SD, or P
    5: "sqm_sold",      # Column F: Square meters
    6: "pallets",       # Column G: Pallet count
    # Columns H, I, J are financial data not used in dashboard
}

# Laying cost per square meter (only applies to SL - Supply and Lay)
LAYING_COST_PER_SQM = 2.20

# Map day names to weekday numbers (Monday=0, Friday=4)
DAY_TO_WEEKDAY = {
    "Monday": 0,
    "Tuesday": 1,
    "Wednesday": 2,
    "Thursday": 3,
    "Friday": 4,
}


@dataclass
class ExcelDeliveryRow:
    """Represents a parsed delivery row from Excel."""

    day: str                    # "Monday", "Tuesday", etc.
    slot: int                   # 1-6
    variety: str                # "Sir Walter", "Empire Zoysia", etc.
    suburb: str                 # "Pimpama", "Coomera", etc.
    service_type: str           # "SL", "SD", "P"
    sqm_sold: float             # e.g., 100.0
    pallets: float              # e.g., 2.0
    week_start_date: Optional[date] = None  # Monday of the week this delivery belongs to

    def get_actual_date(self) -> Optional[date]:
        """Calculate the actual date of this delivery based on week_start_date and day name.

        Returns:
            The actual date, or None if week_start_date is not set.
        """
        if self.week_start_date is None:
            return None

        weekday = DAY_TO_WEEKDAY.get(self.day)
        if weekday is None:
            return None

        return self.week_start_date + timedelta(days=weekday)

    @property
    def laying_cost(self) -> float:
        """Calculate laying cost for this delivery.

        Laying cost is $2.20/SQM and only applies to SL (Supply and Lay) service type.
        SD (Supply and Delivery) and P (Project) have no laying cost.
        """
        if self.service_type == "SL":
            return round(self.sqm_sold * LAYING_COST_PER_SQM, 2)
        return 0.0

    def to_delivery(self) -> Delivery:
        """Convert to Pydantic Delivery model."""
        return Delivery(
            sqm=self.sqm_sold,
            pallets=self.pallets,
            variety=self.variety,
            suburb=self.suburb,
            service_type=self.service_type,
            laying_cost=self.laying_cost,
        )


class ExcelParser:
    """Parses Excel worksheet data into structured delivery records."""

    @staticmethod
    def _safe_float(value: Any, default: float = 0.0) -> float:
        """Safely convert a value to float."""
        if value is None or value == "":
            return default
        try:
            return float(value)
        except (ValueError, TypeError):
            logger.warning(f"Could not convert '{value}' to float, using default {default}")
            return default

    @staticmethod
    def _safe_int(value: Any, default: int = 0) -> int:
        """Safely convert a value to int."""
        if value is None or value == "":
            return default
        try:
            return int(float(value))
        except (ValueError, TypeError):
            logger.warning(f"Could not convert '{value}' to int, using default {default}")
            return default

    @staticmethod
    def _safe_str(value: Any, default: str = "") -> str:
        """Safely convert a value to string."""
        if value is None:
            return default
        return str(value).strip()

    @staticmethod
    def _is_valid_service_type(value: str) -> bool:
        """Check if a service type is valid."""
        return value.upper() in ("SL", "SD", "P")

    @staticmethod
    def _is_header_or_summary_row(row: List[Any]) -> bool:
        """Check if a row is a header or summary row that should be skipped."""
        if not row:
            return False
        first_cell = str(row[0]).lower().strip()
        # Common header indicators
        if first_cell in ("day", "date", "slot", "variety", "suburb", "#", ""):
            return True
        # Summary/total row indicators
        if "total" in first_cell or "weekly" in first_cell or "daily" in first_cell:
            return True
        return False

    @staticmethod
    def _is_empty_row(row: List[Any]) -> bool:
        """Check if a row is empty or has no meaningful data."""
        if not row:
            return True
        # Check if all cells are empty or whitespace
        return all(cell is None or str(cell).strip() == "" for cell in row)

    @classmethod
    def parse_rows(cls, values: List[List[Any]], skip_header: bool = True) -> List[ExcelDeliveryRow]:
        """Parse worksheet values into ExcelDeliveryRow objects.

        Args:
            values: 2D array of cell values from Graph API.
            skip_header: Whether to skip the first row (header).

        Returns:
            List of parsed ExcelDeliveryRow objects.
        """
        if not values:
            logger.warning("No values to parse")
            return []

        parsed_rows: List[ExcelDeliveryRow] = []
        start_index = 1 if skip_header else 0

        for row_index, row in enumerate(values[start_index:], start=start_index):
            # Skip empty rows
            if cls._is_empty_row(row):
                continue

            # Skip if it looks like a header or summary row
            if cls._is_header_or_summary_row(row):
                continue

            try:
                # Extract values with safe defaults
                day = cls._safe_str(row[0] if len(row) > 0 else None)
                slot = cls._safe_int(row[1] if len(row) > 1 else None)
                variety = cls._safe_str(row[2] if len(row) > 2 else None)
                suburb = cls._safe_str(row[3] if len(row) > 3 else None)
                service_type = cls._safe_str(row[4] if len(row) > 4 else None).upper()
                sqm_sold = cls._safe_float(row[5] if len(row) > 5 else None)
                pallets = cls._safe_float(row[6] if len(row) > 6 else None)

                # Skip rows without essential data (must have SQM > 0 to be a valid delivery)
                if sqm_sold <= 0:
                    logger.debug(f"Skipping row {row_index}: SQM is 0 or invalid")
                    continue
                if not variety and not suburb:
                    logger.debug(f"Skipping row {row_index}: missing variety and suburb")
                    continue

                # Validate service type
                if service_type and not cls._is_valid_service_type(service_type):
                    logger.warning(f"Row {row_index}: Invalid service type '{service_type}', defaulting to 'SD'")
                    service_type = "SD"
                elif not service_type:
                    service_type = "SD"

                delivery_row = ExcelDeliveryRow(
                    day=day,
                    slot=slot,
                    variety=variety,
                    suburb=suburb,
                    service_type=service_type,
                    sqm_sold=sqm_sold,
                    pallets=pallets,
                )
                parsed_rows.append(delivery_row)

            except Exception as e:
                logger.error(f"Error parsing row {row_index}: {e}")
                continue

        logger.info(f"Parsed {len(parsed_rows)} delivery rows from {len(values)} total rows")
        return parsed_rows

    @classmethod
    def group_by_day(cls, rows: List[ExcelDeliveryRow]) -> dict[str, List[ExcelDeliveryRow]]:
        """Group parsed rows by day name.

        Args:
            rows: List of ExcelDeliveryRow objects.

        Returns:
            Dictionary mapping day names to lists of deliveries.
        """
        grouped: dict[str, List[ExcelDeliveryRow]] = {}

        for row in rows:
            day = row.day
            if day not in grouped:
                grouped[day] = []
            grouped[day].append(row)

        return grouped
