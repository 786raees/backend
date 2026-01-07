"""Google Sheets service for fetching delivery data from weekly sheets."""
import logging
import re
from datetime import datetime, timezone, date, timedelta
from typing import Dict, List, Optional, Tuple

from app.config import settings
from app.core.constants import get_australia_today
from app.services.google.sheets_client import GoogleSheetsClient
from app.services.google.exceptions import (
    GoogleSheetsError,
    GoogleSheetsNotFoundError,
    GoogleSheetsAPIError,
)
from app.services.excel.excel_parser import ExcelDeliveryRow
from app.services.excel.weekly_sheet_parser import WeeklySheetParser

logger = logging.getLogger(__name__)


def parse_sheet_name_to_date(sheet_name: str) -> Optional[date]:
    """Parse a sheet name like 'Dec-29' into a date (the Monday of that week).

    Args:
        sheet_name: Sheet name in format 'Mon-DD' (e.g., 'Dec-29', 'Jan-05')

    Returns:
        The date corresponding to the sheet name, or None if parsing fails.
    """
    try:
        # Parse the month and day from sheet name
        parsed = datetime.strptime(sheet_name, '%b-%d')

        # Determine the year - use current year, but handle year boundary
        today = get_australia_today()
        year = today.year

        # If the parsed month is January and we're in December, it's next year
        if parsed.month == 1 and today.month == 12:
            year = today.year + 1
        # If the parsed month is December and we're in January, it's last year
        elif parsed.month == 12 and today.month == 1:
            year = today.year - 1

        return date(year, parsed.month, parsed.day)
    except ValueError:
        logger.warning(f"Could not parse sheet name '{sheet_name}' to date")
        return None


class CacheEntry:
    """Simple cache entry with TTL support."""

    def __init__(self, data: any, ttl_seconds: int):
        self.data = data
        self.created_at = datetime.now(timezone.utc)
        self.ttl_seconds = ttl_seconds

    def is_expired(self) -> bool:
        """Check if the cache entry has expired."""
        elapsed = (datetime.now(timezone.utc) - self.created_at).total_seconds()
        return elapsed >= self.ttl_seconds


def get_week_sheet_name(target_date: date) -> str:
    """Get the sheet name for the week containing target_date.

    Sheet names are in format "Mon-DD" where Mon is 3-letter month
    and DD is the day of the Monday of that week.

    Example: If target_date is 2025-12-30 (Tuesday), returns "Dec-29"
    (the Monday of that week).
    """
    # Find the Monday of the week containing target_date
    days_since_monday = target_date.weekday()
    monday = target_date - timedelta(days=days_since_monday)
    return monday.strftime('%b-%d')


def find_best_matching_sheet(target_sheet: str, available_sheets: List[str]) -> Optional[str]:
    """Find the best matching sheet from available sheets.

    Args:
        target_sheet: The ideal sheet name (e.g., "Dec-29")
        available_sheets: List of all available sheet names

    Returns:
        The best matching sheet name, or None if no match found.
    """
    import re

    # First, check if exact match exists
    if target_sheet in available_sheets:
        return target_sheet

    # Filter to only weekly sheets (format like "Dec-28", "Jan-05")
    week_pattern = re.compile(r'^[A-Z][a-z]{2}-\d{2}$')
    weekly_sheets = [s for s in available_sheets if week_pattern.match(s)]

    if not weekly_sheets:
        return None

    # Parse target date
    try:
        target_month_day = datetime.strptime(target_sheet, '%b-%d')
        today = get_australia_today()
        current_year = today.year
        if target_month_day.month == 1 and today.month == 12:
            target_full_date = date(current_year + 1, target_month_day.month, target_month_day.day)
        else:
            target_full_date = date(current_year, target_month_day.month, target_month_day.day)
    except ValueError:
        return weekly_sheets[0] if weekly_sheets else None

    # Find the closest sheet by date
    best_match = None
    best_diff = timedelta(days=365)

    for sheet_name in weekly_sheets:
        try:
            sheet_month_day = datetime.strptime(sheet_name, '%b-%d')
            sheet_year = current_year
            if sheet_month_day.month == 1 and target_full_date.month == 12:
                sheet_year = current_year + 1
            elif sheet_month_day.month == 12 and target_full_date.month == 1:
                sheet_year = current_year - 1

            sheet_date = date(sheet_year, sheet_month_day.month, sheet_month_day.day)
            diff = abs(target_full_date - sheet_date)

            if diff < best_diff:
                best_diff = diff
                best_match = sheet_name
        except ValueError:
            continue

    return best_match


def get_required_week_sheets(start_date: date, num_business_days: int = 10) -> List[str]:
    """Get the list of weekly sheet names needed to cover num_business_days.

    Args:
        start_date: The starting date.
        num_business_days: Number of business days to cover.

    Returns:
        List of unique sheet names (e.g., ["Dec-29", "Jan-05"])
    """
    sheets_needed = set()
    current = start_date
    days_counted = 0

    while days_counted < num_business_days:
        if current.weekday() < 5:  # Monday-Friday
            sheet_name = get_week_sheet_name(current)
            sheets_needed.add(sheet_name)
            days_counted += 1
        current += timedelta(days=1)

    return sorted(list(sheets_needed))


class GoogleSheetsService:
    """Service for fetching delivery data from Google Sheets weekly tabs.

    Reads from weekly sheets with structure:
    - Day headers (Monday - Daily Turf Deliveries)
    - TRUCK 1 section with slots 1-6
    - TRUCK 2 section with slots 1-6
    - Repeat for each day
    """

    _instance: Optional["GoogleSheetsService"] = None
    _cache: Dict[str, CacheEntry] = {}

    def __new__(cls) -> "GoogleSheetsService":
        """Singleton pattern for service reuse."""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._client = None
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        """Initialize the Google Sheets service (lazy - client created on first use)."""
        pass

    def _ensure_client(self) -> GoogleSheetsClient:
        """Lazy initialization of Google Sheets client."""
        if self._client is None:
            if not settings.has_google_credentials:
                raise GoogleSheetsAPIError("Google Sheets credentials not configured")
            self._client = GoogleSheetsClient()
        return self._client

    def _get_cache_key(self, sheet_name: str) -> str:
        """Generate a cache key for a sheet."""
        return f"gsheets:weekly:{sheet_name}"

    def _get_from_cache(self, key: str) -> Optional[any]:
        """Get data from cache if not expired."""
        entry = self._cache.get(key)
        if entry and not entry.is_expired():
            logger.debug(f"Cache hit for {key}")
            return entry.data
        if entry:
            logger.debug(f"Cache expired for {key}")
            del self._cache[key]
        return None

    def _set_cache(self, key: str, data: any) -> None:
        """Store data in cache with TTL."""
        self._cache[key] = CacheEntry(
            data=data,
            ttl_seconds=settings.graph_cache_ttl_seconds,
        )
        logger.debug(f"Cached data for {key} with TTL {settings.graph_cache_ttl_seconds}s")

    def clear_cache(self) -> None:
        """Clear all cached data."""
        self._cache.clear()
        logger.info("Google Sheets service cache cleared")

    def get_available_sheets(self) -> List[str]:
        """Get list of all available sheet names in the spreadsheet."""
        client = self._ensure_client()
        return client.get_available_sheets()

    def get_weekly_sheet_data(self, sheet_name: str) -> Dict[str, List[ExcelDeliveryRow]]:
        """Fetch and parse a weekly sheet.

        Args:
            sheet_name: Name of the weekly sheet (e.g., "Dec-28").

        Returns:
            Dictionary with "Truck 1" and "Truck 2" keys containing delivery lists.

        Raises:
            GoogleSheetsNotFoundError: If sheet doesn't exist.
            GoogleSheetsAPIError: If API call fails.
        """
        cache_key = self._get_cache_key(sheet_name)

        # Check cache first
        cached = self._get_from_cache(cache_key)
        if cached is not None:
            return cached

        # Fetch from Google Sheets
        logger.info(f"Fetching weekly sheet data for '{sheet_name}'")
        client = self._ensure_client()

        try:
            values = client.get_worksheet_data(sheet_name)
        except GoogleSheetsNotFoundError:
            logger.warning(f"Weekly sheet '{sheet_name}' not found, trying MASTER")
            # Fall back to MASTER sheet if specific week not found
            values = client.get_worksheet_data("MASTER")

        # Parse with the weekly sheet parser
        parsed_data = WeeklySheetParser.parse_weekly_sheet(values)

        # Cache the result
        self._set_cache(cache_key, parsed_data)

        return parsed_data

    def get_all_truck_deliveries(self) -> Tuple[List[ExcelDeliveryRow], List[ExcelDeliveryRow]]:
        """Get deliveries for both trucks from the current week's sheet.

        Determines which weekly sheet(s) to read based on today's date,
        then parses the truck-based sections.

        Returns:
            Tuple of (truck1_deliveries, truck2_deliveries).
        """
        today = get_australia_today()

        # Get the sheet names we need to cover 10 business days
        required_sheets = get_required_week_sheets(today, 10)
        logger.info(f"Required weekly sheets for schedule: {required_sheets}")

        # Get available sheets to find best matches
        try:
            available_sheets = self.get_available_sheets()
        except GoogleSheetsError as e:
            logger.error(f"Failed to get available sheets: {e}")
            available_sheets = []

        # Collect all deliveries
        all_truck1: List[ExcelDeliveryRow] = []
        all_truck2: List[ExcelDeliveryRow] = []
        sheets_read: set = set()

        for sheet_name in required_sheets:
            # Find the best matching sheet from available sheets
            actual_sheet = find_best_matching_sheet(sheet_name, available_sheets)

            if not actual_sheet:
                logger.warning(f"No matching sheet for '{sheet_name}', using MASTER")
                actual_sheet = "MASTER"

            if actual_sheet != sheet_name:
                logger.info(f"Mapped '{sheet_name}' -> '{actual_sheet}'")

            # Avoid reading the same sheet multiple times
            if actual_sheet in sheets_read:
                continue
            sheets_read.add(actual_sheet)

            try:
                sheet_data = self.get_weekly_sheet_data(actual_sheet)

                # Parse sheet name to get week start date
                week_start = parse_sheet_name_to_date(actual_sheet)
                if week_start:
                    logger.info(f"Sheet '{actual_sheet}' -> week_start_date: {week_start}")

                # Set week_start_date on each delivery
                for delivery in sheet_data.get("Truck 1", []):
                    delivery.week_start_date = week_start
                    all_truck1.append(delivery)

                for delivery in sheet_data.get("Truck 2", []):
                    delivery.week_start_date = week_start
                    all_truck2.append(delivery)

            except GoogleSheetsNotFoundError as e:
                logger.warning(f"Sheet '{actual_sheet}' not found: {e}")
                continue
            except GoogleSheetsError as e:
                logger.error(f"Error reading sheet '{actual_sheet}': {e}")
                continue

        logger.info(f"Total: {len(all_truck1)} Truck 1 and {len(all_truck2)} Truck 2 deliveries")
        return all_truck1, all_truck2

    def get_worksheet_deliveries(self, worksheet_name: str) -> List[ExcelDeliveryRow]:
        """Legacy method for backwards compatibility - reads from Truck 1/Truck 2 sheets.

        For the new weekly structure, use get_all_truck_deliveries() instead.
        """
        logger.warning(f"get_worksheet_deliveries called with '{worksheet_name}' - using legacy mode")
        from app.services.excel.excel_parser import ExcelParser

        client = self._ensure_client()
        values = client.get_worksheet_data(worksheet_name)
        return ExcelParser.parse_rows(values)

    def is_available(self) -> bool:
        """Check if the Google Sheets service is available.

        Returns True if Google credentials are configured.
        """
        return settings.has_google_credentials


# Global service instance
google_sheets_service = GoogleSheetsService()
