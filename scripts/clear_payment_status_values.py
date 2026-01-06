"""
Clear Payment Status Column Values Script
==========================================

This script clears any existing values/formulas in column K (Payment Status)
for all slot rows in weekly tabs, allowing users to use the dropdown.

Run from the backend directory:
    python scripts/clear_payment_status_values.py
"""

import os
import sys
import json
import logging
import time
from typing import List, Dict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Rate limiting
API_DELAY = 1.5  # seconds between calls

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "1aygUdgHPMuI14uiGfqZdDFIvBHMmarPwbdfdONml_UI"


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
        raise ValueError("GOOGLE_SERVICE_ACCOUNT_JSON not found")

    creds_info = json.loads(creds_json)
    return service_account.Credentials.from_service_account_info(creds_info, scopes=SCOPES)


def get_sheets_service():
    """Create Google Sheets API service."""
    return build("sheets", "v4", credentials=get_credentials())


def get_weekly_tabs(service) -> List[str]:
    """Get list of weekly tab names."""
    import re
    result = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = result.get("sheets", [])
    existing = [sheet["properties"]["title"] for sheet in sheets]
    pattern = r'^[A-Z][a-z]{2}-\d{2}$'
    return [name for name in existing if re.match(pattern, name)]


def find_slot_rows(service, sheet_name: str) -> List[Dict]:
    """Find all slot rows (rows where column A is 1-6)."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A:A"
    ).execute()

    values = result.get("values", [])
    slot_rows = []

    for idx, row in enumerate(values):
        if not row:
            continue
        try:
            slot_num = int(row[0])
            if 1 <= slot_num <= 6:
                slot_rows.append({
                    "row_index": idx,
                    "row_number": idx + 1
                })
        except (ValueError, TypeError):
            pass

    return slot_rows


def clear_payment_status_values(service, sheet_name: str):
    """Clear column K values for all slot rows that have formula data."""
    slot_rows = find_slot_rows(service, sheet_name)

    if not slot_rows:
        logger.warning(f"No slot rows found in '{sheet_name}'")
        return 0

    # Get current values in column K for slot rows
    ranges = [f"'{sheet_name}'!K{r['row_number']}" for r in slot_rows]

    # Batch get values
    result = service.spreadsheets().values().batchGet(
        spreadsheetId=SPREADSHEET_ID,
        ranges=ranges
    ).execute()

    value_ranges = result.get("valueRanges", [])

    # Find rows that need clearing (have formula values like $xxx.xx)
    rows_to_clear = []
    for i, vr in enumerate(value_ranges):
        values = vr.get("values", [[]])
        if values and values[0]:
            cell_value = str(values[0][0])
            # Clear if it looks like a price or formula result (starts with $)
            # or is not a valid payment status
            valid_statuses = ["Paid", "Payment Pending", "Cash", ""]
            if cell_value not in valid_statuses:
                rows_to_clear.append(slot_rows[i]["row_number"])

    if not rows_to_clear:
        logger.info(f"No values to clear in '{sheet_name}'")
        return 0

    # Clear the values
    data_updates = []
    for row_num in rows_to_clear:
        data_updates.append({
            "range": f"'{sheet_name}'!K{row_num}",
            "values": [[""]]
        })

    body = {
        "valueInputOption": "USER_ENTERED",
        "data": data_updates
    }

    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body=body
    ).execute()

    cleared = len(rows_to_clear)
    logger.info(f"Cleared {cleared} payment status cells in '{sheet_name}'")
    return cleared


def main():
    """Main function to clear payment status values in all weekly tabs."""
    logger.info("=" * 60)
    logger.info("GLC Dashboard - Clear Payment Status Values Script")
    logger.info("=" * 60)

    try:
        service = get_sheets_service()
        logger.info("Connected to Google Sheets API")

        weekly_tabs = get_weekly_tabs(service)
        logger.info(f"Found {len(weekly_tabs)} weekly tabs")

        if not weekly_tabs:
            logger.warning("No weekly tabs found.")
            return

        total_cleared = 0
        for tab_name in weekly_tabs:
            logger.info(f"\nProcessing '{tab_name}'...")
            time.sleep(API_DELAY)
            cleared = clear_payment_status_values(service, tab_name)
            total_cleared += cleared
            time.sleep(API_DELAY)

        logger.info("\n" + "=" * 60)
        logger.info(f"Cleared {total_cleared} cells across {len(weekly_tabs)} tabs")
        logger.info("=" * 60)

    except Exception as e:
        logger.error(f"Error: {e}")
        raise


if __name__ == "__main__":
    main()
