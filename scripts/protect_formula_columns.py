"""
Protect Formula Columns Script
==============================

This script protects formula columns (G, H, K, L, M, N, O) from editing
while keeping data entry columns (A-F, I, J, P) editable.

This prevents users from accidentally breaking formulas.

Run from the backend directory:
    python scripts/protect_formula_columns.py
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

API_DELAY = 1.0
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "1aygUdgHPMuI14uiGfqZdDFIvBHMmarPwbdfdONml_UI"

# Column indices (0-based): G=6, H=7, K=10, L=11, M=12, N=13, O=14
PROTECTED_COLUMNS = [
    {"start": 6, "end": 8},   # G-H (columns 6-7)
    {"start": 10, "end": 15}, # K-O (columns 10-14)
]


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
    return build("sheets", "v4", credentials=get_credentials())


def get_weekly_tabs_with_ids(service) -> List[Dict]:
    """Get list of weekly tab names with their sheet IDs."""
    import re
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


def protect_columns(service, sheet_name: str, sheet_id: int) -> bool:
    """Add protected ranges for formula columns on a sheet."""

    requests = []

    for col_range in PROTECTED_COLUMNS:
        requests.append({
            "addProtectedRange": {
                "protectedRange": {
                    "range": {
                        "sheetId": sheet_id,
                        "startColumnIndex": col_range["start"],
                        "endColumnIndex": col_range["end"],
                    },
                    "description": f"Formula columns - Do not edit",
                    "warningOnly": True,  # Shows warning but allows owner to edit
                }
            }
        })

    if requests:
        body = {"requests": requests}
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=body
        ).execute()
        logger.info(f"Protected formula columns in '{sheet_name}'")
        return True

    return False


def main():
    logger.info("=" * 60)
    logger.info("GLC Dashboard - Protect Formula Columns Script")
    logger.info("=" * 60)
    logger.info("Protecting columns: G, H, K, L, M, N, O")
    logger.info("Users will see a warning when trying to edit these columns")
    logger.info("")

    try:
        service = get_sheets_service()
        logger.info("Connected to Google Sheets API")

        weekly_tabs = get_weekly_tabs_with_ids(service)
        logger.info(f"Found {len(weekly_tabs)} weekly tabs")

        if not weekly_tabs:
            logger.warning("No weekly tabs found.")
            return

        protected_count = 0
        for tab in weekly_tabs:
            logger.info(f"Processing '{tab['name']}'...")
            time.sleep(API_DELAY)
            if protect_columns(service, tab["name"], tab["sheet_id"]):
                protected_count += 1

        logger.info("\n" + "=" * 60)
        logger.info(f"Protected formula columns in {protected_count} tabs")
        logger.info("=" * 60)

    except Exception as e:
        logger.error(f"Error: {e}")
        raise


if __name__ == "__main__":
    main()
