"""
Add Payment Status Column to Google Sheets
==========================================

This script adds a new "Payment Status" column (K) with dropdown validation
to all weekly tabs in the Turf Supply Google Sheet.

Payment Status Options:
- Paid (Green indicator)
- Payment Pending (Red indicator)
- Cash (Yellow indicator)

Run from the backend directory:
    python scripts/add_payment_status_column.py
"""

import os
import sys
import json
import logging
import time
from typing import List, Dict

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Rate limiting - wait between API calls
API_DELAY = 1.5  # seconds between calls

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Configuration
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "1aygUdgHPMuI14uiGfqZdDFIvBHMmarPwbdfdONml_UI"

# Payment Status dropdown values
PAYMENT_STATUS_OPTIONS = ["Paid", "Payment Pending", "Cash"]


def get_credentials():
    """Get Google API credentials from environment or .env file."""
    creds_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")

    if not creds_json:
        env_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), ".env")
        if os.path.exists(env_path):
            with open(env_path, "r") as f:
                for line in f:
                    if line.startswith("GOOGLE_SERVICE_ACCOUNT_JSON="):
                        creds_json = line.split("=", 1)[1].strip()
                        break

    if not creds_json:
        raise ValueError("GOOGLE_SERVICE_ACCOUNT_JSON not found in environment or .env file")

    creds_info = json.loads(creds_json)
    credentials = service_account.Credentials.from_service_account_info(
        creds_info, scopes=SCOPES
    )
    return credentials


def get_sheets_service():
    """Create Google Sheets API service."""
    credentials = get_credentials()
    service = build("sheets", "v4", credentials=credentials)
    return service


def get_sheet_id_by_name(service, sheet_name: str) -> int:
    """Get sheet ID by name."""
    result = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for sheet in result.get("sheets", []):
        if sheet["properties"]["title"] == sheet_name:
            return sheet["properties"]["sheetId"]
    return None


def get_weekly_tabs(service) -> List[str]:
    """Get list of weekly tab names (format: Mon-DD like Dec-29, Jan-05)."""
    import re
    result = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = result.get("sheets", [])
    existing = [sheet["properties"]["title"] for sheet in sheets]
    pattern = r'^[A-Z][a-z]{2}-\d{2}$'
    weekly_tabs = [name for name in existing if re.match(pattern, name)]
    logger.info(f"Found {len(weekly_tabs)} weekly tabs: {weekly_tabs}")
    return weekly_tabs


def find_slot_rows(service, sheet_name: str) -> List[Dict]:
    """Find all slot rows in a weekly sheet."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A:F"
    ).execute()

    values = result.get("values", [])
    slot_rows = []
    current_day = None
    current_truck = None

    for idx, row in enumerate(values):
        if not row:
            continue

        cell = str(row[0]).strip() if row else ""

        for day in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]:
            if cell.startswith(day):
                current_day = day
                current_truck = None
                break

        if cell.upper() == "TRUCK 1":
            current_truck = "Truck 1"
        elif cell.upper() == "TRUCK 2":
            current_truck = "Truck 2"

        try:
            slot_num = int(cell)
            if 1 <= slot_num <= 6 and current_day and current_truck:
                slot_rows.append({
                    "row_index": idx,
                    "row_number": idx + 1,
                    "day": current_day,
                    "truck": current_truck,
                    "slot": slot_num
                })
        except (ValueError, TypeError):
            pass

    return slot_rows


def find_header_rows(service, sheet_name: str) -> List[int]:
    """Find all column header rows (rows containing 'Slot' in column A)."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A:A"
    ).execute()

    values = result.get("values", [])
    header_rows = []

    for idx, row in enumerate(values):
        if row and str(row[0]).strip().lower() == "slot":
            header_rows.append(idx + 1)

    return header_rows


def add_payment_status_header(service, sheet_name: str, header_rows: List[int]):
    """Add 'Payment Status' header to column K for all header rows."""
    if not header_rows:
        return

    data_updates = []
    for row_num in header_rows:
        data_updates.append({
            "range": f"'{sheet_name}'!K{row_num}",
            "values": [["Payment Status"]]
        })

    if data_updates:
        body = {
            "valueInputOption": "USER_ENTERED",
            "data": data_updates
        }
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=body
        ).execute()
        logger.info(f"Added 'Payment Status' header to {len(header_rows)} rows in '{sheet_name}'")


def add_payment_status_dropdown(service, sheet_name: str):
    """Add Payment Status dropdown to column K for all slot rows."""
    sheet_id = get_sheet_id_by_name(service, sheet_name)
    if sheet_id is None:
        logger.error(f"Could not find sheet ID for '{sheet_name}'")
        return

    slot_rows = find_slot_rows(service, sheet_name)
    if not slot_rows:
        logger.warning(f"No slot rows found in '{sheet_name}'")
        return

    min_row = min(r["row_index"] for r in slot_rows)
    max_row = max(r["row_index"] for r in slot_rows) + 1

    request = {
        "requests": [{
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": min_row,
                    "endRowIndex": max_row,
                    "startColumnIndex": 10,
                    "endColumnIndex": 11
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [
                            {"userEnteredValue": option}
                            for option in PAYMENT_STATUS_OPTIONS
                        ]
                    },
                    "showCustomUi": True,
                    "strict": False
                }
            }
        }]
    }

    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=request
        ).execute()
        logger.info(f"Added Payment Status dropdown to column K in '{sheet_name}'")
    except HttpError as e:
        logger.error(f"Error adding dropdown to '{sheet_name}': {e}")


def main():
    """Main function to add Payment Status column to all weekly tabs."""
    logger.info("=" * 60)
    logger.info("GLC Dashboard - Add Payment Status Column Script")
    logger.info("=" * 60)

    try:
        service = get_sheets_service()
        logger.info("Connected to Google Sheets API")

        weekly_tabs = get_weekly_tabs(service)

        if not weekly_tabs:
            logger.warning("No weekly tabs found.")
            return

        # Skip already processed tabs (up to Jun-29)
        processed_tabs = ['Dec-29', 'Jan-05', 'Jan-12', 'Jan-19', 'Jan-27', 'Jan-26',
                          'Feb-02', 'Feb-09', 'Feb-16', 'Feb-23', 'Mar-02', 'Mar-09',
                          'Mar-16', 'Mar-23', 'Mar-30', 'Apr-06', 'Apr-13', 'Apr-20',
                          'Apr-27', 'May-04', 'May-11', 'May-18', 'May-25', 'Jun-01',
                          'Jun-08', 'Jun-15', 'Jun-22', 'Jun-29']

        for tab_name in weekly_tabs:
            if tab_name in processed_tabs:
                logger.info(f"Skipping already processed '{tab_name}'")
                continue

            logger.info(f"\nProcessing '{tab_name}'...")
            time.sleep(API_DELAY)
            header_rows = find_header_rows(service, tab_name)
            time.sleep(API_DELAY)
            add_payment_status_header(service, tab_name, header_rows)
            time.sleep(API_DELAY)
            add_payment_status_dropdown(service, tab_name)
            time.sleep(API_DELAY)

        logger.info("\n" + "=" * 60)
        logger.info("Payment Status column added successfully!")
        logger.info("=" * 60)

    except Exception as e:
        logger.error(f"Error: {e}")
        raise


if __name__ == "__main__":
    main()
