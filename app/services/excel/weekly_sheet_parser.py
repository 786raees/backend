"""Parser for the new weekly sheet structure with truck sections."""
import logging
from dataclasses import dataclass
from typing import Any, Dict, List, Tuple

from app.services.excel.excel_parser import ExcelDeliveryRow

logger = logging.getLogger(__name__)

# Days that we track in the weekly sheet
VALID_DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']


class WeeklySheetParser:
    """Parses weekly sheet data with truck-based sections.

    Supports TWO column structures:

    STRUCTURE A (Slot-first - Original):
        Row: "Monday - Daily Turf Deliveries..."
        Row: (empty)
        Row: "TRUCK 1"
        Row: Header (Slot, Variety, Suburb, Service Type, SQM Sold, Pallets)
        Rows: Slots 1-6 (Column A has slot number)
        Row: Totals
        (repeat for TRUCK 2 and each day)

        Columns: A=Slot, B=Variety, C=Suburb, D=Service Type, E=SQM, F=Pallets

    STRUCTURE B (Day-first - New):
        Row: Week header
        Row: (empty)
        Row: Column headers
        Row: "TRUCK 1"
        Rows: Monday slots 1-6 (Column A has "Monday")
        Rows: Tuesday slots 1-6 (Column A has "Tuesday")
        ... (all days under TRUCK 1)
        Row: Totals
        (repeat for TRUCK 2)

        Columns: A=Day, B=Variety, C=Suburb, D=Service Type, E=SQM, F=Pallets
    """

    @staticmethod
    def _safe_float(value: Any, default: float = 0.0) -> float:
        """Safely convert a value to float."""
        if value is None or value == "":
            return default
        try:
            # Handle formatted numbers like "100.00"
            cleaned = str(value).replace(',', '').replace('$', '').strip()
            return float(cleaned)
        except (ValueError, TypeError):
            logger.warning(f"Could not convert '{value}' to float, using default {default}")
            return default

    @staticmethod
    def _safe_str(value: Any, default: str = "") -> str:
        """Safely convert a value to string."""
        if value is None:
            return default
        return str(value).strip()

    @staticmethod
    def _safe_int(value: Any, default: int = 0) -> int:
        """Safely convert a value to int."""
        if value is None or value == "":
            return default
        try:
            return int(float(value))
        except (ValueError, TypeError):
            return default

    @staticmethod
    def _is_day_header(row: List[Any]) -> Tuple[bool, str]:
        """Check if a row is a day header and extract the day name.

        Returns:
            Tuple of (is_day_header, day_name)
        """
        if not row or not row[0]:
            return False, ""

        cell = str(row[0]).strip()
        for day in VALID_DAYS:
            if cell.startswith(day):
                return True, day
        return False, ""

    @staticmethod
    def _is_day_value(value: Any) -> Tuple[bool, str]:
        """Check if a value is a day name (for Day-first structure).

        Returns:
            Tuple of (is_day, day_name)
        """
        if value is None:
            return False, ""

        cell = str(value).strip()
        for day in VALID_DAYS:
            if cell == day or cell.startswith(day):
                return True, day
        return False, ""

    @staticmethod
    def _is_truck_header(row: List[Any]) -> Tuple[bool, str]:
        """Check if a row is a truck header.

        Returns:
            Tuple of (is_truck_header, truck_name: "Truck 1" or "Truck 2")
        """
        if not row or not row[0]:
            return False, ""

        cell = str(row[0]).strip().upper()
        if cell == "TRUCK 1":
            return True, "Truck 1"
        elif cell == "TRUCK 2":
            return True, "Truck 2"
        return False, ""

    @staticmethod
    def _is_slot_row(row: List[Any]) -> bool:
        """Check if a row is a slot data row (slot number in first column)."""
        if not row or not row[0]:
            return False
        try:
            slot = int(row[0])
            return 1 <= slot <= 6
        except (ValueError, TypeError):
            return False

    @staticmethod
    def _is_day_data_row(row: List[Any]) -> Tuple[bool, str]:
        """Check if a row is a data row with day name in first column (Day-first structure).

        Returns:
            Tuple of (is_data_row, day_name)
        """
        if not row or len(row) < 5:
            return False, ""

        # Column A should be a day name
        is_day, day_name = WeeklySheetParser._is_day_value(row[0])
        if not is_day:
            return False, ""

        # Should have some data in other columns (variety or SQM)
        variety = str(row[1]).strip() if len(row) > 1 and row[1] else ""
        sqm_str = str(row[4]).strip() if len(row) > 4 and row[4] else ""

        # Must have either variety or SQM to be a valid data row
        if variety or sqm_str:
            return True, day_name

        return False, ""

    @staticmethod
    def _is_valid_service_type(value: str) -> bool:
        """Check if a service type is valid."""
        return value.upper() in ("SL", "SD", "P")

    @classmethod
    def _detect_structure(cls, values: List[List[Any]]) -> str:
        """Detect which column structure the sheet uses.

        Returns:
            "slot_first" if column A contains slot numbers (1-6)
            "day_first" if column A contains day names (Monday, Tuesday, etc.)
        """
        for row in values:
            if not row or not row[0]:
                continue

            cell = str(row[0]).strip()

            # Skip obvious non-data rows
            if cell.upper() in ("TRUCK 1", "TRUCK 2", "SLOT", "DAY", "TOTALS", ""):
                continue
            if cell.lower().startswith(("monday -", "tuesday -", "wednesday -", "thursday -", "friday -")):
                # This is a day header row like "Monday - Daily Turf Deliveries"
                continue

            # Check if it's a slot number
            try:
                slot = int(cell)
                if 1 <= slot <= 6:
                    logger.info("Detected SLOT-FIRST column structure")
                    return "slot_first"
            except ValueError:
                pass

            # Check if it's a day name
            is_day, _ = cls._is_day_value(cell)
            if is_day:
                logger.info("Detected DAY-FIRST column structure")
                return "day_first"

        # Default to slot-first (original structure)
        logger.info("Could not detect structure, defaulting to SLOT-FIRST")
        return "slot_first"

    @classmethod
    def parse_weekly_sheet(cls, values: List[List[Any]]) -> Dict[str, List[ExcelDeliveryRow]]:
        """Parse a weekly sheet into deliveries grouped by truck.

        Automatically detects and handles both column structures:
        - Slot-first: Column A = Slot number (1-6)
        - Day-first: Column A = Day name (Monday, Tuesday, etc.)

        Args:
            values: 2D array of cell values from Google Sheets API.

        Returns:
            Dictionary with keys "Truck 1" and "Truck 2", each containing
            a list of ExcelDeliveryRow objects.
        """
        result = {
            "Truck 1": [],
            "Truck 2": [],
        }

        if not values:
            logger.warning("No values to parse in weekly sheet")
            return result

        # Detect column structure
        structure = cls._detect_structure(values)

        if structure == "day_first":
            return cls._parse_day_first_structure(values)
        else:
            return cls._parse_slot_first_structure(values)

    @classmethod
    def _parse_slot_first_structure(cls, values: List[List[Any]]) -> Dict[str, List[ExcelDeliveryRow]]:
        """Parse sheet with Slot-first column structure (original).

        Column mapping: A=Slot, B=Variety, C=Suburb, D=Service Type, E=SQM, F=Pallets
        Day comes from day header rows like "Monday - Daily Turf Deliveries"
        """
        result = {
            "Truck 1": [],
            "Truck 2": [],
        }

        current_day = None
        current_truck = None
        in_data_section = False

        for row_idx, row in enumerate(values):
            if not row:
                continue

            # Check for day header
            is_day, day_name = cls._is_day_header(row)
            if is_day:
                current_day = day_name
                current_truck = None
                in_data_section = False
                continue

            # Check for truck header
            is_truck, truck_name = cls._is_truck_header(row)
            if is_truck:
                current_truck = truck_name
                in_data_section = False
                continue

            # Check for column header row (after truck header)
            if current_truck and len(row) >= 2 and cls._safe_str(row[0]).lower() == "slot":
                in_data_section = True
                continue

            # Check for totals row (end of data section)
            if cls._safe_str(row[0]).lower() == "totals":
                in_data_section = False
                continue

            # Parse slot data row
            if in_data_section and current_day and current_truck and cls._is_slot_row(row):
                try:
                    slot = cls._safe_int(row[0])
                    variety = cls._safe_str(row[1] if len(row) > 1 else None)
                    suburb = cls._safe_str(row[2] if len(row) > 2 else None)
                    service_type = cls._safe_str(row[3] if len(row) > 3 else None).upper()
                    sqm_sold = cls._safe_float(row[4] if len(row) > 4 else None)
                    pallets = cls._safe_float(row[5] if len(row) > 5 else None)
                    # Column I (index 8) = Delivery Fee
                    delivery_fee = cls._safe_float(row[8] if len(row) > 8 else None, default=0)
                    # Column J (index 9) = Laying Fee
                    laying_fee = cls._safe_float(row[9] if len(row) > 9 else None, default=0)
                    # Column P (index 15) = Payment Status
                    payment_status = cls._safe_str(row[15] if len(row) > 15 else None) or None

                    # Skip empty slots (no SQM)
                    if sqm_sold <= 0:
                        continue

                    # Validate/default service type
                    if not service_type or not cls._is_valid_service_type(service_type):
                        service_type = "SD"

                    delivery = ExcelDeliveryRow(
                        day=current_day,
                        slot=slot,
                        variety=variety if variety else "Unknown",
                        suburb=suburb if suburb else "Unknown",
                        service_type=service_type,
                        sqm_sold=sqm_sold,
                        pallets=pallets,
                        delivery_fee=delivery_fee,
                        laying_fee=laying_fee,
                        payment_status=payment_status,
                    )
                    result[current_truck].append(delivery)

                except Exception as e:
                    logger.error(f"Error parsing row {row_idx}: {e}")
                    continue

        logger.info(f"Parsed {len(result['Truck 1'])} Truck 1 and {len(result['Truck 2'])} Truck 2 deliveries (slot-first)")
        return result

    @classmethod
    def _parse_day_first_structure(cls, values: List[List[Any]]) -> Dict[str, List[ExcelDeliveryRow]]:
        """Parse sheet with Day-first column structure (new).

        Column mapping: A=Day, B=Variety, C=Suburb, D=Service Type, E=SQM, F=Pallets
        Day comes from column A of each data row.
        """
        result = {
            "Truck 1": [],
            "Truck 2": [],
        }

        current_truck = None
        in_data_section = False
        slot_counter = {}  # Track slot numbers per day per truck

        for row_idx, row in enumerate(values):
            if not row:
                continue

            # Check for truck header
            is_truck, truck_name = cls._is_truck_header(row)
            if is_truck:
                current_truck = truck_name
                in_data_section = True
                slot_counter = {}  # Reset slot counter for new truck
                continue

            # Check for column header row (skip it)
            first_cell = cls._safe_str(row[0]).lower()
            if first_cell in ("day", "slot", ""):
                continue

            # Check for totals row (end of data section)
            if first_cell == "totals":
                in_data_section = False
                continue

            # Check if this is a day-first data row
            is_data, day_name = cls._is_day_data_row(row)
            if is_data and current_truck and in_data_section:
                try:
                    # Day-first structure: A=Day, B=Variety, C=Suburb, D=Service Type, E=SQM, F=Pallets
                    variety = cls._safe_str(row[1] if len(row) > 1 else None)
                    suburb = cls._safe_str(row[2] if len(row) > 2 else None)
                    service_type = cls._safe_str(row[3] if len(row) > 3 else None).upper()
                    sqm_sold = cls._safe_float(row[4] if len(row) > 4 else None)
                    pallets = cls._safe_float(row[5] if len(row) > 5 else None)
                    # Column I (index 8) = Delivery Fee
                    delivery_fee = cls._safe_float(row[8] if len(row) > 8 else None, default=0)
                    # Column J (index 9) = Laying Fee
                    laying_fee = cls._safe_float(row[9] if len(row) > 9 else None, default=0)
                    # Column P (index 15) = Payment Status
                    payment_status = cls._safe_str(row[15] if len(row) > 15 else None) or None

                    # Skip empty rows (no SQM)
                    if sqm_sold <= 0:
                        continue

                    # Auto-increment slot number for each day
                    if day_name not in slot_counter:
                        slot_counter[day_name] = 0
                    slot_counter[day_name] += 1
                    slot = slot_counter[day_name]

                    # Validate/default service type
                    if not service_type or not cls._is_valid_service_type(service_type):
                        service_type = "SD"

                    delivery = ExcelDeliveryRow(
                        day=day_name,
                        slot=slot,
                        variety=variety if variety else "Unknown",
                        suburb=suburb if suburb else "Unknown",
                        service_type=service_type,
                        sqm_sold=sqm_sold,
                        pallets=pallets,
                        delivery_fee=delivery_fee,
                        laying_fee=laying_fee,
                        payment_status=payment_status,
                    )
                    result[current_truck].append(delivery)

                except Exception as e:
                    logger.error(f"Error parsing row {row_idx}: {e}")
                    continue

        logger.info(f"Parsed {len(result['Truck 1'])} Truck 1 and {len(result['Truck 2'])} Truck 2 deliveries (day-first)")
        return result
