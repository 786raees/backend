"""
Fix Pallet Formula Script
==========================

This script fixes the pallet calculation formula for Row 3 on Truck 1
in all weekly tabs.

Pallet Formula:
- Empire Zoysia, Sir Walter, Summerland Buffalo: SQM / 50
- Wintergreen Couch, AussiBlue Couch: SQM / 60

Run from the backend directory:
    python scripts/fix_pallet_formula.py
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


def find_all_slot_rows(service, sheet_name: str) -> List[Dict]:
    """Find all slot rows with their row numbers."""
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


def generate_pallet_formula(row_number: int) -> str:
    """Generate the pallet calculation formula for a given row."""
    r = row_number
    return (
        f'=IF(E{r}="","",'
        f'IF(OR(B{r}="Empire Zoysia",B{r}="Sir Walter",B{r}="Summerland Buffalo"),E{r}/50,'
        f'IF(OR(B{r}="Wintergreen Couch",B{r}="AussiBlue Couch"),E{r}/60,"")))'
    )


def fix_pallet_formulas(service, sheet_name: str):
    """Fix pallet formulas for all slot rows in a sheet."""
    slot_rows = find_all_slot_rows(service, sheet_name)

    if not slot_rows:
        logger.warning(f"No slot rows found in '{sheet_name}'")
        return 0

    data_updates = []
    for slot_info in slot_rows:
        row_num = slot_info["row_number"]
        formula = generate_pallet_formula(row_num)
        data_updates.append({
            "range": f"'{sheet_name}'!F{row_num}",
            "values": [[formula]]
        })

    if data_updates:
        body = {
            "valueInputOption": "USER_ENTERED",
            "data": data_updates
        }
        result = service.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=body
        ).execute()

        updated = result.get('totalUpdatedCells', 0)
        logger.info(f"Fixed {updated} pallet formulas in '{sheet_name}'")
        return updated

    return 0


def main():
    """Main function to fix pallet formulas in all weekly tabs."""
    logger.info("=" * 60)
    logger.info("GLC Dashboard - Fix Pallet Formula Script")
    logger.info("=" * 60)

    try:
        service = get_sheets_service()
        logger.info("Connected to Google Sheets API")

        weekly_tabs = get_weekly_tabs(service)
        logger.info(f"Found {len(weekly_tabs)} weekly tabs: {weekly_tabs}")

        if not weekly_tabs:
            logger.warning("No weekly tabs found.")
            return

        total_fixed = 0
        for tab_name in weekly_tabs:
            logger.info(f"\nProcessing '{tab_name}'...")
            time.sleep(API_DELAY)
            fixed = fix_pallet_formulas(service, tab_name)
            total_fixed += fixed
            time.sleep(API_DELAY)

        logger.info("\n" + "=" * 60)
        logger.info(f"Fixed {total_fixed} pallet formulas across {len(weekly_tabs)} tabs")
        logger.info("=" * 60)

    except Exception as e:
        logger.error(f"Error: {e}")
        raise


if __name__ == "__main__":
    main()
