"""
Add Laying Cost Formula to Google Sheet
========================================

This script adds the Laying Cost formula to column P (after existing columns).
Laying Cost = SQM × $2.20

Run from the backend directory:
    python scripts/add_laying_cost.py
"""

import os
import sys
import json
import logging
import time
from typing import List, Dict, Any

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Configuration
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "1aygUdgHPMuI14uiGfqZdDFIvBHMmarPwbdfdONml_UI"

# Laying cost per SQM
LAYING_COST_PER_SQM = 2.20


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
    credentials = service_account.Credentials.from_service_account_info(
        creds_info, scopes=SCOPES
    )
    return credentials


def get_sheets_service():
    """Create Google Sheets API service."""
    credentials = get_credentials()
    service = build("sheets", "v4", credentials=credentials)
    return service


def get_weekly_tabs(service) -> List[str]:
    """Get list of weekly tab names."""
    import re
    result = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = result.get("sheets", [])
    names = [sheet["properties"]["title"] for sheet in sheets]
    pattern = r'^[A-Z][a-z]{2}-\d{2}$'
    weekly_tabs = [name for name in names if re.match(pattern, name)]
    return weekly_tabs


def find_slot_rows(service, sheet_name: str) -> List[Dict]:
    """Find all slot rows in a weekly sheet."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A:A"
    ).execute()

    values = result.get("values", [])
    slot_rows = []
    current_day = None
    current_truck = None

    for idx, row in enumerate(values):
        if not row:
            continue

        cell = str(row[0]).strip() if row else ""

        # Check for day header
        for day in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]:
            if cell.startswith(day):
                current_day = day
                current_truck = None
                break

        # Check for truck header
        if cell.upper() == "TRUCK 1":
            current_truck = "Truck 1"
        elif cell.upper() == "TRUCK 2":
            current_truck = "Truck 2"

        # Check for slot row (1-6) or day name (for day-first structure)
        try:
            slot_num = int(cell)
            if 1 <= slot_num <= 6 and current_day and current_truck:
                slot_rows.append({
                    "row_number": idx + 1,
                    "day": current_day,
                    "truck": current_truck,
                })
        except (ValueError, TypeError):
            # Check if it's a day name (day-first structure)
            if current_truck and cell in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]:
                slot_rows.append({
                    "row_number": idx + 1,
                    "day": cell,
                    "truck": current_truck,
                })

    return slot_rows


def add_laying_cost_header(service, sheet_name: str):
    """Add 'Laying Cost' header to column P in header rows."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A:A"
    ).execute()

    values = result.get("values", [])
    header_rows = []

    for idx, row in enumerate(values):
        if not row:
            continue
        cell = str(row[0]).strip().lower() if row else ""
        if cell in ("slot", "day"):
            header_rows.append(idx + 1)

    if not header_rows:
        return

    data_updates = []
    for row_num in header_rows:
        data_updates.append({
            "range": f"'{sheet_name}'!P{row_num}",
            "values": [["Laying Cost"]]
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
        logger.info(f"Added 'Laying Cost' header to {len(header_rows)} rows in '{sheet_name}'")


def add_laying_cost_formulas(service, sheet_name: str):
    """Add Laying Cost formula to column P for all slot rows."""
    slot_rows = find_slot_rows(service, sheet_name)

    if not slot_rows:
        logger.warning(f"No slot rows found in '{sheet_name}'")
        return

    logger.info(f"Adding Laying Cost formula to {len(slot_rows)} rows in '{sheet_name}'...")

    data_updates = []
    for slot_info in slot_rows:
        row_num = slot_info["row_number"]
        # Laying Cost = SQM (column E) × $2.20
        formula = f'=IF(E{row_num}="","",E{row_num}*{LAYING_COST_PER_SQM})'
        data_updates.append({
            "range": f"'{sheet_name}'!P{row_num}",
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
        logger.info(f"Added Laying Cost formula to {len(slot_rows)} rows in '{sheet_name}'")


def main():
    """Main function to add Laying Cost formula."""
    logger.info("=" * 60)
    logger.info("GLC Dashboard - Add Laying Cost Formula")
    logger.info("=" * 60)
    logger.info(f"Laying Cost Rate: ${LAYING_COST_PER_SQM:.2f}/SQM")

    try:
        service = get_sheets_service()
        logger.info("Connected to Google Sheets API")

        # Get all weekly tabs
        weekly_tabs = get_weekly_tabs(service)
        logger.info(f"Found {len(weekly_tabs)} weekly tabs")

        # Process each weekly tab
        for tab_name in weekly_tabs:
            logger.info(f"\nProcessing '{tab_name}'...")

            # Add header
            add_laying_cost_header(service, tab_name)

            # Add formulas
            add_laying_cost_formulas(service, tab_name)

            # Rate limiting
            time.sleep(1)

        logger.info("\n" + "=" * 60)
        logger.info("Laying Cost formula added to all weekly tabs!")
        logger.info("=" * 60)
        logger.info("\nColumn P now contains: Laying Cost = SQM × $2.20")

    except Exception as e:
        logger.error(f"Error: {e}")
        raise


if __name__ == "__main__":
    main()
