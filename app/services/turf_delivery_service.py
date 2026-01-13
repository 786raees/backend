"""
Turf Delivery Service - Business logic for turf delivery dashboard.

This service handles:
1. Reading delivery data from Google Sheets
2. Creating new deliveries
3. Updating existing deliveries
4. Deleting deliveries
"""

import logging
import re
from datetime import date, timedelta
from typing import Dict, List, Optional, Any

from app.services.google.sheets_client import GoogleSheetsClient
from app.config import settings

logger = logging.getLogger(__name__)


class TurfDeliveryService:
    """Service for turf delivery data operations."""

    # Configuration constants
    TRUCKS = ["TRUCK 1", "TRUCK 2"]
    DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    VARIETIES = [
        "Empire Zoysia",
        "Sir Walter",
        "Wintergreen Couch",
        "AussiBlue Couch",
        "Summerland Buffalo"
    ]
    SERVICE_TYPES = ["SL", "SD", "P"]  # Supply & Lay, Supply & Delivery, Project
    PAYMENT_STATUSES = ["Unpaid", "Paid", "Partial"]

    # Row structure constants for slot-first weekly sheets
    # Each day section = 23 rows:
    #   Day header (1) + empty (1) + TRUCK 1 header (1) + column headers (1) +
    #   6 slots + totals (1) + empty (1) + TRUCK 2 header (1) + column headers (1) +
    #   6 slots + totals (1) + separator (1) = 23 rows
    DAY_HEADER_ROWS = {
        "Monday": 1,
        "Tuesday": 24,      # 1 + 23
        "Wednesday": 47,    # 24 + 23
        "Thursday": 70,     # 47 + 23
        "Friday": 93        # 70 + 23
    }

    # Truck offsets within each day
    TRUCK_SLOT_OFFSETS = {
        "TRUCK 1": 4,       # First truck starts at day_row + 4 (header + empty + truck header + column headers)
        "TRUCK 2": 14       # Second truck starts at day_row + 14 (after truck1 slots + totals + empty + truck2 header + column headers)
    }

    SLOTS_PER_TRUCK = 6

    # Column mapping (0-indexed for API)
    COLUMNS = {
        "slot": 0,           # A
        "variety": 1,        # B
        "suburb": 2,         # C
        "service_type": 3,   # D
        "sqm_sold": 4,       # E
        "pallets": 5,        # F (formula)
        "sell_price_sqm": 6,   # G (formula)
        "cost_price_sqm": 7,   # H (formula)
        "delivery_fee": 8,     # I (formula)
        "laying_fee": 9,       # J (formula)
        "sell_price": 10,      # K (formula)
        "turf_cost": 11,       # L (formula)
        "total_revenue": 12,   # M (formula)
        "gross_profit": 13,    # N (formula)
        "margin_pct": 14,      # O (formula)
        "payment_status": 15   # P
    }

    # Editable columns (letter format)
    EDITABLE_COLUMNS = ["B", "C", "D", "E", "I", "J", "P"]  # Variety, Suburb, Service Type, SQM, Delivery Fee, Laying Fee, Payment Status

    def __init__(self):
        self._client: Optional[GoogleSheetsClient] = None
        self._service = None
        self._spreadsheet_id = settings.google_spreadsheet_id if hasattr(settings, 'google_spreadsheet_id') else None

    def _get_service(self):
        """Lazy initialization of sheets service."""
        if self._service is None:
            if self._client is None:
                print(f"DEBUG: Creating new GoogleSheetsClient")
                self._client = GoogleSheetsClient()
                print(f"DEBUG: Client created, type={type(self._client)}, methods={dir(self._client)}")
            print(f"DEBUG: Calling _ensure_service on client")
            self._service = self._client._ensure_service()
            print(f"DEBUG: Service created successfully")
        return self._service

    def _get_spreadsheet_id(self):
        """Get the turf supply spreadsheet ID."""
        if self._spreadsheet_id is None:
            self._spreadsheet_id = settings.google_spreadsheet_id
        return self._spreadsheet_id

    def get_week_tab_name(self, target_date: date) -> str:
        """Get the sheet name for the week containing target_date.

        Sheet names are in format "Mon-DD" where Mon is 3-letter month
        and DD is the day of the Monday of that week.
        """
        days_since_monday = target_date.weekday()
        monday = target_date - timedelta(days=days_since_monday)
        return monday.strftime('%b-%d')

    def calculate_row_number(self, day: str, truck: str, slot: int) -> int:
        """Calculate the row number for a delivery.

        Args:
            day: Day name (Monday, Tuesday, etc.)
            truck: Truck name (TRUCK 1, TRUCK 2)
            slot: Slot number (1-6)

        Returns:
            Row number (1-indexed)
        """
        day_base = self.DAY_HEADER_ROWS[day]
        truck_offset = self.TRUCK_SLOT_OFFSETS[truck]
        return day_base + truck_offset + (slot - 1)

    async def get_available_weeks(self) -> List[str]:
        """Get list of available week tabs."""
        service = self._get_service()
        spreadsheet_id = self._get_spreadsheet_id()

        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])

        week_pattern = re.compile(r'^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{2}$')
        week_tabs = []

        for sheet in sheets:
            title = sheet['properties']['title']
            if week_pattern.match(title):
                week_tabs.append(title)

        return sorted(week_tabs)

    def parse_currency(self, value: str) -> float:
        """Parse currency string to float."""
        if not value:
            return 0.0
        try:
            # Remove currency symbols and commas
            clean_value = str(value).replace('$', '').replace(',', '').strip()
            return float(clean_value) if clean_value else 0.0
        except (ValueError, AttributeError):
            return 0.0

    def _parse_row_to_delivery(self, row: List[Any], slot: int, row_number: int) -> Dict:
        """Parse a single row into a delivery dict."""
        def safe_get(idx: int, default: str = "") -> str:
            if idx < len(row):
                val = row[idx]
                return str(val) if val is not None else default
            return default

        return {
            "slot": slot,
            "row_number": row_number,
            "variety": safe_get(self.COLUMNS["variety"]),
            "suburb": safe_get(self.COLUMNS["suburb"]),
            "service_type": safe_get(self.COLUMNS["service_type"]),
            "sqm_sold": self.parse_currency(safe_get(self.COLUMNS["sqm_sold"])),
            "pallets": self.parse_currency(safe_get(self.COLUMNS["pallets"])),
            "payment_status": safe_get(self.COLUMNS["payment_status"])
        }

    async def update_delivery_field(
        self,
        week_tab: str,
        row_number: int,
        column: str,
        value: str
    ) -> Dict:
        """Update a single field in a delivery row."""
        # Validate column is editable
        if column not in self.EDITABLE_COLUMNS:
            return {
                "success": False,
                "error": f"Column {column} is not editable. Editable columns: {', '.join(self.EDITABLE_COLUMNS)}",
                "error_code": 400
            }

        try:
            service = self._get_service()
            spreadsheet_id = self._get_spreadsheet_id()

            print(f"DEBUG UPDATE: week_tab={week_tab}, row={row_number}, column={column}, value={value}")
            print(f"DEBUG UPDATE: spreadsheet_id={spreadsheet_id}")

            # Check if week tab exists
            available_weeks = await self.get_available_weeks()
            if week_tab not in available_weeks:
                return {
                    "success": False,
                    "error": f"Week tab '{week_tab}' not found",
                    "error_code": 404
                }

            # Update the cell
            range_notation = f"'{week_tab}'!{column}{row_number}"
            print(f"DEBUG UPDATE: range_notation={range_notation}")

            result = service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_notation,
                valueInputOption="USER_ENTERED",
                body={"values": [[value]]}
            ).execute()

            print(f"DEBUG UPDATE: Google Sheets API response={result}")
            logger.info(f"Updated {week_tab}!{column}{row_number} = {value}")

            # Clear cache so next read gets fresh data
            print(f"DEBUG UPDATE: Clearing cache after successful update")
            try:
                import app.services.google_sheets_service as gss_module
                gss_module.google_sheets_service.clear_cache()
                print(f"DEBUG UPDATE: Cache cleared successfully")
            except Exception as cache_error:
                print(f"DEBUG UPDATE: Error clearing cache: {cache_error}")

            return {
                "success": True,
                "updated": {
                    "week_tab": week_tab,
                    "row_number": row_number,
                    "column": column,
                    "value": value
                }
            }

        except Exception as e:
            logger.error(f"Error updating delivery field: {e}")
            return {
                "success": False,
                "error": str(e),
                "error_code": 500
            }

    async def create_delivery(
        self,
        week_tab: str,
        day: str,
        truck: str,
        slot: int,
        variety: str = "",
        suburb: str = "",
        service_type: str = "",
        sqm_sold: str = "",
        delivery_fee: str = "",
        laying_fee: str = ""
    ) -> Dict:
        """Create a new delivery in an empty slot."""
        # Validate inputs
        if day not in self.DAYS:
            return {
                "success": False,
                "error": f"Invalid day '{day}'. Must be one of: {', '.join(self.DAYS)}",
                "error_code": 400
            }

        if truck not in self.TRUCKS:
            return {
                "success": False,
                "error": f"Invalid truck '{truck}'. Must be one of: {', '.join(self.TRUCKS)}",
                "error_code": 400
            }

        if slot < 1 or slot > self.SLOTS_PER_TRUCK:
            return {
                "success": False,
                "error": f"Invalid slot {slot}. Must be 1-{self.SLOTS_PER_TRUCK}",
                "error_code": 400
            }

        try:
            service = self._get_service()
            spreadsheet_id = self._get_spreadsheet_id()

            # Calculate row number
            row_number = self.calculate_row_number(day, truck, slot)

            # Check if week tab exists
            available_weeks = await self.get_available_weeks()
            if week_tab not in available_weeks:
                return {
                    "success": False,
                    "error": f"Week tab '{week_tab}' not found",
                    "error_code": 404
                }

            # Check if slot is already occupied
            check_range = f"'{week_tab}'!B{row_number}"
            check_result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=check_range
            ).execute()

            existing_values = check_result.get('values', [])
            if existing_values and len(existing_values) > 0 and len(existing_values[0]) > 0:
                existing_variety = existing_values[0][0]
                if existing_variety and existing_variety.strip():
                    return {
                        "success": False,
                        "error": f"Slot is already occupied",
                        "error_code": 409
                    }

            # Prepare data to write
            # Columns: A=slot, B=variety, C=suburb, D=service_type, E=sqm_sold, F-H=formulas, I=delivery_fee, J=laying_fee, K-O=formulas, P=payment_status
            row_values = [
                str(slot),          # A: Slot
                variety,            # B: Variety
                suburb,             # C: Suburb
                service_type,       # D: Service Type
                sqm_sold,           # E: SQM Sold
                # F-H are formulas (pallets, pricing) - don't overwrite
            ]

            # Write columns A-E first
            range_notation = f"'{week_tab}'!A{row_number}:E{row_number}"

            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_notation,
                valueInputOption="USER_ENTERED",
                body={"values": [row_values]}
            ).execute()

            # Write delivery_fee and laying_fee (columns I and J) if provided
            if delivery_fee or laying_fee:
                fee_range = f"'{week_tab}'!I{row_number}:J{row_number}"
                fee_values = [[delivery_fee or "", laying_fee or ""]]
                service.spreadsheets().values().update(
                    spreadsheetId=spreadsheet_id,
                    range=fee_range,
                    valueInputOption="USER_ENTERED",
                    body={"values": fee_values}
                ).execute()

            logger.info(f"Created delivery: {week_tab} {day} {truck} Slot {slot}")

            # Clear cache after successful create
            try:
                logger.debug("DEBUG CREATE: Clearing cache after successful create")
                import app.services.google_sheets_service as gss_module
                gss_module.google_sheets_service.clear_cache()
                logger.debug("DEBUG CREATE: Cache cleared successfully")
            except Exception as cache_error:
                logger.warning(f"Failed to clear cache after create: {cache_error}")

            # Return the created delivery
            delivery = {
                "slot": slot,
                "row_number": row_number,
                "variety": variety,
                "suburb": suburb,
                "service_type": service_type,
                "sqm_sold": sqm_sold,
                "delivery_fee": delivery_fee,
                "laying_fee": laying_fee,
                "payment_status": ""
            }

            return {
                "success": True,
                "message": "Delivery created successfully",
                "row_number": row_number,
                "delivery": delivery
            }

        except Exception as e:
            logger.error(f"Error creating delivery: {e}")
            return {
                "success": False,
                "error": str(e),
                "error_code": 500
            }

    async def delete_delivery(
        self,
        week_tab: str,
        row_number: Optional[int] = None,
        day: Optional[str] = None,
        truck: Optional[str] = None,
        slot: Optional[int] = None
    ) -> Dict:
        """Delete a delivery by clearing all data in the row."""
        # Validate inputs
        if row_number is None:
            if not day or not truck or slot is None:
                return {
                    "success": False,
                    "error": "Must provide either row_number or (day + truck + slot)",
                    "error_code": 400
                }

            if day not in self.DAYS:
                return {
                    "success": False,
                    "error": f"Invalid day '{day}'. Must be one of: {', '.join(self.DAYS)}",
                    "error_code": 400
                }

            if truck not in self.TRUCKS:
                return {
                    "success": False,
                    "error": f"Invalid truck '{truck}'. Must be one of: {', '.join(self.TRUCKS)}",
                    "error_code": 400
                }

            if slot < 1 or slot > self.SLOTS_PER_TRUCK:
                return {
                    "success": False,
                    "error": f"Invalid slot {slot}. Must be 1-{self.SLOTS_PER_TRUCK}",
                    "error_code": 400
                }

            # Calculate row number
            row_number = self.calculate_row_number(day, truck, slot)

        try:
            service = self._get_service()
            spreadsheet_id = self._get_spreadsheet_id()

            # Check if week tab exists
            available_weeks = await self.get_available_weeks()
            if week_tab not in available_weeks:
                return {
                    "success": False,
                    "error": f"Week tab '{week_tab}' not found",
                    "error_code": 404
                }

            # Clear data columns (B through P)
            # Column A (Slot) should remain to preserve structure
            empty_row = [""] * 15  # 15 empty values for columns B-P

            range_notation = f"'{week_tab}'!B{row_number}:P{row_number}"

            logger.debug(f"DEBUG DELETE: week_tab={week_tab}, row={row_number}")
            logger.debug(f"DEBUG DELETE: range_notation={range_notation}")

            result = service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_notation,
                valueInputOption="USER_ENTERED",
                body={"values": [empty_row]}
            ).execute()

            logger.debug(f"DEBUG DELETE: Google Sheets API response={result}")
            logger.info(f"Deleted delivery at row {row_number} in {week_tab}")

            # Clear cache after successful delete
            try:
                logger.debug("DEBUG DELETE: Clearing cache after successful delete")
                import app.services.google_sheets_service as gss_module
                gss_module.google_sheets_service.clear_cache()
                logger.debug("DEBUG DELETE: Cache cleared successfully")
            except Exception as cache_error:
                logger.warning(f"Failed to clear cache after delete: {cache_error}")

            return {
                "success": True,
                "message": "Delivery deleted successfully",
                "row_number": row_number
            }

        except Exception as e:
            logger.error(f"Error deleting delivery: {e}")
            return {
                "success": False,
                "error": str(e),
                "error_code": 500
            }


# Singleton instance
turf_delivery_service = TurfDeliveryService()
