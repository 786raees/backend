"""
Update Payment Status Dropdown to Include 'Project' Option
==========================================================

Adds "Project" option to Payment Status dropdown (Column P) across all weekly tabs.

New dropdown options:
- Paid
- Payment Pending
- Cash
- Project (NEW)

Run: python scripts/update_payment_status_dropdown.py
"""

import os
import sys
import json
import logging
import re
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from google.oauth2 import service_account
from googleapiclient.discovery import build

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "1aygUdgHPMuI14uiGfqZdDFIvBHMmarPwbdfdONml_UI"

# Column P = index 15 (0-indexed)
PAYMENT_STATUS_COLUMN = 15


def get_credentials():
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
    credentials = get_credentials()
    return build("sheets", "v4", credentials=credentials)


def get_all_weekly_tabs_with_ids(service) -> list:
    """Get all weekly tab names with sheet IDs."""
    result = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = result.get("sheets", [])
    pattern = r'^[A-Z][a-z]{2}-\d{2}$'
    weekly_tabs = []
    for sheet in sheets:
        title = sheet["properties"]["title"]
        sheet_id = sheet["properties"]["sheetId"]
        row_count = sheet["properties"]["gridProperties"]["rowCount"]
        if re.match(pattern, title):
            weekly_tabs.append({"name": title, "sheet_id": sheet_id, "row_count": row_count})
    return weekly_tabs


def find_slot_rows(service, sheet_name: str) -> list:
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
                slot_rows.append(idx)  # 0-indexed row
        except (ValueError, TypeError):
            pass

    return slot_rows


def update_payment_status_dropdown(service, sheet_name: str, sheet_id: int, slot_rows: list):
    """Update Payment Status dropdown to include 'Project' option for all slot rows."""

    if not slot_rows:
        logger.warning(f"No slot rows found in '{sheet_name}'")
        return

    requests = []

    for row_idx in slot_rows:
        requests.append({
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row_idx,
                    "endRowIndex": row_idx + 1,
                    "startColumnIndex": PAYMENT_STATUS_COLUMN,
                    "endColumnIndex": PAYMENT_STATUS_COLUMN + 1
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [
                            {"userEnteredValue": "Paid"},
                            {"userEnteredValue": "Payment Pending"},
                            {"userEnteredValue": "Cash"},
                            {"userEnteredValue": "Project"}
                        ]
                    },
                    "showCustomUi": True,
                    "strict": False
                }
            }
        })

    # Batch requests in chunks to avoid API limits
    chunk_size = 100
    for i in range(0, len(requests), chunk_size):
        chunk = requests[i:i + chunk_size]
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": chunk}
        ).execute()

    logger.info(f"Updated {len(slot_rows)} dropdowns in '{sheet_name}'")


def main():
    logger.info("=" * 60)
    logger.info("Adding 'Project' to Payment Status Dropdown")
    logger.info("=" * 60)
    logger.info("New options: Paid, Payment Pending, Cash, Project")
    logger.info("")

    service = get_sheets_service()
    weekly_tabs = get_all_weekly_tabs_with_ids(service)

    logger.info(f"Found {len(weekly_tabs)} weekly tabs")

    total_updated = 0
    for tab in weekly_tabs:
        logger.info(f"Processing '{tab['name']}'...")
        time.sleep(0.5)
        slot_rows = find_slot_rows(service, tab["name"])
        time.sleep(0.5)
        update_payment_status_dropdown(service, tab["name"], tab["sheet_id"], slot_rows)
        total_updated += len(slot_rows)
        time.sleep(0.5)

    logger.info("\n" + "=" * 60)
    logger.info(f"Updated {total_updated} dropdown cells across {len(weekly_tabs)} tabs")
    logger.info("=" * 60)


if __name__ == "__main__":
    main()
