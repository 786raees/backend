"""
Unhide Staff Columns (Delivery Fee and Laying Fee)
==================================================

Unhides columns I (Delivery Fee) and J (Laying Fee) across all weekly tabs.

Run: python scripts/unhide_staff_columns.py
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

# Columns to unhide (0-indexed)
# I = 8, J = 9
COLUMNS_TO_UNHIDE = [8, 9]


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
        if re.match(pattern, title):
            weekly_tabs.append({"name": title, "sheet_id": sheet_id})
    return weekly_tabs


def unhide_columns(service, sheet_name: str, sheet_id: int):
    """Unhide Delivery Fee (I) and Laying Fee (J) columns."""
    requests = []

    for col_index in COLUMNS_TO_UNHIDE:
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": col_index,
                    "endIndex": col_index + 1
                },
                "properties": {
                    "hiddenByUser": False
                },
                "fields": "hiddenByUser"
            }
        })

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": requests}
    ).execute()
    logger.info(f"Unhidden columns I (Delivery Fee) and J (Laying Fee) in '{sheet_name}'")


def main():
    logger.info("=" * 60)
    logger.info("Unhiding Delivery Fee (I) and Laying Fee (J) Columns")
    logger.info("=" * 60)

    service = get_sheets_service()
    weekly_tabs = get_all_weekly_tabs_with_ids(service)

    logger.info(f"Found {len(weekly_tabs)} weekly tabs")

    for tab in weekly_tabs:
        unhide_columns(service, tab["name"], tab["sheet_id"])
        time.sleep(0.3)

    logger.info("\n" + "=" * 60)
    logger.info("Done! Delivery Fee and Laying Fee columns now visible.")
    logger.info("=" * 60)


if __name__ == "__main__":
    main()
