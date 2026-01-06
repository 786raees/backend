"""
Fix Column K (Sell Price) and Add Payment Status Column
========================================================

This script:
1. Restores "Sell Price" header to column K (was incorrectly changed to Payment Status)
2. Restores the Sell Price formula (=E*G) to column K for all slot rows
3. Adds "Payment Status" header to column P
4. Adds Payment Status dropdown to column P for all slot rows

Run from the backend directory:
    python scripts/fix_column_k_and_add_payment_status.py
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


def get_sheet_id_by_name(service, sheet_name: str) -> int:
    """Get sheet ID by name."""
    result = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for sheet in result.get("sheets", []):
        if sheet["properties"]["title"] == sheet_name:
            return sheet["properties"]["sheetId"]
    return None


def find_header_rows(service, sheet_name: str) -> List[int]:
    """Find all header rows (rows where column A is 'Slot')."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A:A"
    ).execute()

    values = result.get("values", [])
    header_rows = []

    for idx, row in enumerate(values):
        if row and str(row[0]).strip().lower() == "slot":
            header_rows.append(idx + 1)  # 1-indexed

    return header_rows


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


def fix_column_k_headers(service, sheet_name: str, header_rows: List[int]):
    """Restore 'Sell Price' header to column K."""
    if not header_rows:
        return

    data_updates = []
    for row_num in header_rows:
        data_updates.append({
            "range": f"'{sheet_name}'!K{row_num}",
            "values": [["Sell Price"]]
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
        logger.info(f"Restored 'Sell Price' header to column K in '{sheet_name}'")


def fix_column_k_formulas(service, sheet_name: str, slot_rows: List[Dict]):
    """Restore Sell Price formula (=E*G) to column K for all slot rows."""
    if not slot_rows:
        return

    data_updates = []
    for slot_info in slot_rows:
        row_num = slot_info["row_number"]
        # Sell Price formula: =IF(OR(E{r}="",G{r}=""),"",E{r}*G{r})
        formula = f'=IF(OR(E{row_num}="",G{row_num}=""),"",E{row_num}*G{row_num})'
        data_updates.append({
            "range": f"'{sheet_name}'!K{row_num}",
            "values": [[formula]]
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
        logger.info(f"Restored {len(data_updates)} Sell Price formulas in column K for '{sheet_name}'")


def add_payment_status_header(service, sheet_name: str, header_rows: List[int]):
    """Add 'Payment Status' header to column P."""
    if not header_rows:
        return

    data_updates = []
    for row_num in header_rows:
        data_updates.append({
            "range": f"'{sheet_name}'!P{row_num}",
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
        logger.info(f"Added 'Payment Status' header to column P in '{sheet_name}'")


def add_payment_status_dropdown(service, sheet_name: str, slot_rows: List[Dict]):
    """Add Payment Status dropdown to column P for all slot rows."""
    sheet_id = get_sheet_id_by_name(service, sheet_name)
    if sheet_id is None:
        logger.error(f"Could not find sheet ID for '{sheet_name}'")
        return

    if not slot_rows:
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
                    "startColumnIndex": 15,  # Column P (0-indexed)
                    "endColumnIndex": 16
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
        logger.info(f"Added Payment Status dropdown to column P in '{sheet_name}'")
    except HttpError as e:
        logger.error(f"Failed to add dropdown to '{sheet_name}': {e}")


def clear_old_payment_status_column_k(service, sheet_name: str, header_rows: List[int]):
    """Remove 'Payment Status' text from column K if it exists (from our previous error)."""
    # This is handled by fix_column_k_headers which overwrites with 'Sell Price'
    pass


def main():
    """Main function to fix column K and add Payment Status to column P."""
    logger.info("=" * 60)
    logger.info("GLC Dashboard - Fix Column K & Add Payment Status to Column P")
    logger.info("=" * 60)

    try:
        service = get_sheets_service()
        logger.info("Connected to Google Sheets API")

        weekly_tabs = get_weekly_tabs(service)
        logger.info(f"Found {len(weekly_tabs)} weekly tabs")

        if not weekly_tabs:
            logger.warning("No weekly tabs found.")
            return

        for tab_name in weekly_tabs:
            logger.info(f"\nProcessing '{tab_name}'...")

            time.sleep(API_DELAY)
            header_rows = find_header_rows(service, tab_name)

            time.sleep(API_DELAY)
            slot_rows = find_slot_rows(service, tab_name)

            # Step 1: Restore 'Sell Price' header to column K
            time.sleep(API_DELAY)
            fix_column_k_headers(service, tab_name, header_rows)

            # Step 2: Restore Sell Price formula to column K
            time.sleep(API_DELAY)
            fix_column_k_formulas(service, tab_name, slot_rows)

            # Step 3: Add 'Payment Status' header to column P
            time.sleep(API_DELAY)
            add_payment_status_header(service, tab_name, header_rows)

            # Step 4: Add Payment Status dropdown to column P
            time.sleep(API_DELAY)
            add_payment_status_dropdown(service, tab_name, slot_rows)

        logger.info("\n" + "=" * 60)
        logger.info("Column K restored and Payment Status added to column P!")
        logger.info("=" * 60)

    except Exception as e:
        logger.error(f"Error: {e}")
        raise


if __name__ == "__main__":
    main()
