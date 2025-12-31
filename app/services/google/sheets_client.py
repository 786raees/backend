"""Google Sheets API client for fetching worksheet data."""
import json
import logging
from typing import Any, List, Optional

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from app.config import settings
from app.services.google.exceptions import (
    GoogleSheetsAuthError,
    GoogleSheetsNotFoundError,
    GoogleSheetsAPIError,
)

logger = logging.getLogger(__name__)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


class GoogleSheetsClient:
    """Client for Google Sheets API operations.

    Uses service account authentication for server-to-server access.
    The Google Sheet must be shared with the service account email.
    """

    def __init__(self):
        """Initialize the client (lazy initialization)."""
        self._service = None
        self._credentials = None

    def _ensure_service(self):
        """Lazy initialization of Sheets service.

        Creates credentials from service account JSON and builds
        the Sheets API service.

        Raises:
            GoogleSheetsAuthError: If authentication fails
        """
        if self._service is None:
            try:
                # Parse service account JSON from settings
                service_account_info = settings.google_service_account_info
                if not service_account_info:
                    raise GoogleSheetsAuthError(
                        "Google service account credentials not configured"
                    )

                self._credentials = service_account.Credentials.from_service_account_info(
                    service_account_info,
                    scopes=SCOPES,
                )
                self._service = build(
                    "sheets", "v4", credentials=self._credentials, cache_discovery=False
                )
                logger.info("Google Sheets service initialized successfully")

            except json.JSONDecodeError as e:
                logger.error(f"Invalid service account JSON: {e}")
                raise GoogleSheetsAuthError(
                    "Invalid service account JSON format",
                    details={"error": str(e)},
                )
            except Exception as e:
                logger.error(f"Failed to initialize Google Sheets service: {e}")
                raise GoogleSheetsAuthError(
                    f"Authentication failed: {e}",
                    details={"error": str(e)},
                )

        return self._service

    def get_worksheet_data(
        self,
        sheet_name: str,
        range_cols: str = "A:G",
    ) -> List[List[Any]]:
        """Fetch data from a worksheet.

        Args:
            sheet_name: Name of the worksheet tab (e.g., "Jan 5 - T1")
            range_cols: Column range to fetch (default: A:G for delivery data)

        Returns:
            2D list of cell values, similar to Excel format.
            Each inner list is a row, each element is a cell value.

        Raises:
            GoogleSheetsNotFoundError: If worksheet doesn't exist
            GoogleSheetsAPIError: If API call fails
        """
        service = self._ensure_service()
        spreadsheet_id = settings.google_spreadsheet_id

        if not spreadsheet_id:
            raise GoogleSheetsAPIError("Google Spreadsheet ID not configured")

        try:
            # Format range as 'Sheet Name'!A:G
            range_name = f"'{sheet_name}'!{range_cols}"

            result = (
                service.spreadsheets()
                .values()
                .get(spreadsheetId=spreadsheet_id, range=range_name)
                .execute()
            )

            values = result.get("values", [])
            logger.info(f"Retrieved {len(values)} rows from '{sheet_name}'")
            return values

        except HttpError as e:
            if e.resp.status == 404:
                logger.warning(f"Worksheet '{sheet_name}' not found")
                raise GoogleSheetsNotFoundError(
                    f"Worksheet '{sheet_name}' not found",
                    details={"sheet_name": sheet_name},
                )
            elif e.resp.status == 403:
                logger.error(f"Permission denied for spreadsheet: {e}")
                raise GoogleSheetsAuthError(
                    "Permission denied. Ensure the spreadsheet is shared with the service account.",
                    details={"status": 403},
                )
            else:
                logger.error(f"Google Sheets API error: {e}")
                raise GoogleSheetsAPIError(
                    f"API error: {e.reason}",
                    status_code=e.resp.status,
                    details={"error": str(e)},
                )

    def get_available_sheets(self) -> List[str]:
        """Get list of all worksheet names in the spreadsheet.

        Returns:
            List of worksheet tab names

        Raises:
            GoogleSheetsAPIError: If API call fails
        """
        service = self._ensure_service()
        spreadsheet_id = settings.google_spreadsheet_id

        if not spreadsheet_id:
            raise GoogleSheetsAPIError("Google Spreadsheet ID not configured")

        try:
            metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

            sheets = metadata.get("sheets", [])
            sheet_names = [sheet["properties"]["title"] for sheet in sheets]

            logger.info(f"Found {len(sheet_names)} worksheets: {sheet_names}")
            return sheet_names

        except HttpError as e:
            logger.error(f"Failed to get sheet list: {e}")
            raise GoogleSheetsAPIError(
                f"Failed to retrieve spreadsheet metadata: {e.reason}",
                status_code=e.resp.status,
                details={"error": str(e)},
            )

    def check_sheet_exists(self, sheet_name: str) -> bool:
        """Check if a specific worksheet exists.

        Args:
            sheet_name: Name of the worksheet to check

        Returns:
            True if sheet exists, False otherwise
        """
        try:
            available = self.get_available_sheets()
            return sheet_name in available
        except GoogleSheetsAPIError:
            return False

    def update_worksheet_data(
        self,
        sheet_name: str,
        values: List[List[Any]],
        range_cols: str = "A:G",
    ) -> int:
        """Update data in a worksheet.

        Args:
            sheet_name: Name of the worksheet tab (e.g., "Truck 1")
            values: 2D list of cell values to write
            range_cols: Column range to write to (default: A:G)

        Returns:
            Number of cells updated

        Raises:
            GoogleSheetsAPIError: If API call fails
        """
        service = self._ensure_service()
        spreadsheet_id = settings.google_spreadsheet_id

        if not spreadsheet_id:
            raise GoogleSheetsAPIError("Google Spreadsheet ID not configured")

        try:
            # Format range as 'Sheet Name'!A:G
            range_name = f"'{sheet_name}'!{range_cols}"

            # Clear existing data first
            service.spreadsheets().values().clear(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                body={}
            ).execute()

            # Write new data
            body = {
                "values": values
            }
            result = (
                service.spreadsheets()
                .values()
                .update(
                    spreadsheetId=spreadsheet_id,
                    range=range_name,
                    valueInputOption="RAW",
                    body=body
                )
                .execute()
            )

            updated_cells = result.get("updatedCells", 0)
            logger.info(f"Updated {updated_cells} cells in '{sheet_name}'")
            return updated_cells

        except HttpError as e:
            if e.resp.status == 403:
                logger.error(f"Permission denied for writing: {e}")
                raise GoogleSheetsAuthError(
                    "Permission denied. Ensure the service account has Editor access.",
                    details={"status": 403},
                )
            else:
                logger.error(f"Google Sheets API error: {e}")
                raise GoogleSheetsAPIError(
                    f"API error: {e.reason}",
                    status_code=e.resp.status,
                    details={"error": str(e)},
                )
