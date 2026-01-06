"""
Fix All Calculation Formulas Script
====================================

This script fixes ALL formulas in columns G through O for all slot rows.
It regenerates correct formulas with proper row references to eliminate #REF! errors.

Column formulas:
- G (Sell $/SQM): =IF(B{r}="","",VLOOKUP(B{r},Setup!$A$2:$C$6,2,FALSE))
- H (Cost $/SQM): =IF(B{r}="","",VLOOKUP(B{r},Setup!$A$2:$C$6,3,FALSE))
- I (Delivery Fee): Manual entry (no formula)
- J (Laying Fee): Manual entry (no formula)
- K (Sell Price): =IF(OR(E{r}="",G{r}=""),"",E{r}*G{r})
- L (Turf Cost): =IF(OR(E{r}="",H{r}=""),"",E{r}*H{r})
- M (Total Revenue): =IF(OR(K{r}="",I{r}="",J{r}=""),"",K{r}+I{r}+J{r})
- N (Gross Profit): =IF(OR(M{r}="",L{r}=""),"",M{r}-L{r})
- O (Margin %): =IF(OR(N{r}="",M{r}="",M{r}=0),"",N{r}/M{r})

Run from the backend directory:
    python scripts/fix_all_formulas.py
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

API_DELAY = 1.5
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
    return build("sheets", "v4", credentials=get_credentials())


def get_weekly_tabs(service) -> List[str]:
    import re
    result = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = result.get("sheets", [])
    existing = [sheet["properties"]["title"] for sheet in sheets]
    pattern = r'^[A-Z][a-z]{2}-\d{2}$'
    return [name for name in existing if re.match(pattern, name)]


def find_all_slot_rows(service, sheet_name: str) -> List[Dict]:
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
                    "row_number": idx + 1,
                    "slot": slot_num
                })
        except (ValueError, TypeError):
            pass
    return slot_rows


def generate_formulas(row: int) -> Dict[str, str]:
    """Generate all formulas for a given row number."""
    r = row
    return {
        "G": f'=IF(B{r}="","",VLOOKUP(B{r},Setup!$A$2:$C$6,2,FALSE))',
        "H": f'=IF(B{r}="","",VLOOKUP(B{r},Setup!$A$2:$C$6,3,FALSE))',
        "K": f'=IF(OR(E{r}="",G{r}=""),"",E{r}*G{r})',
        "L": f'=IF(OR(E{r}="",H{r}=""),"",E{r}*H{r})',
        "M": f'=IF(OR(K{r}="",I{r}="",J{r}=""),"",K{r}+I{r}+J{r})',
        "N": f'=IF(OR(M{r}="",L{r}=""),"",M{r}-L{r})',
        "O": f'=IF(OR(N{r}="",M{r}="",M{r}=0),"",N{r}/M{r})',
    }


def fix_all_formulas(service, sheet_name: str) -> int:
    """Fix all formulas in columns G, H, K-O for all slot rows."""
    slot_rows = find_all_slot_rows(service, sheet_name)
    if not slot_rows:
        logger.warning(f"No slot rows found in '{sheet_name}'")
        return 0

    data_updates = []
    columns_to_fix = ["G", "H", "K", "L", "M", "N", "O"]

    for slot_info in slot_rows:
        row_num = slot_info["row_number"]
        formulas = generate_formulas(row_num)

        for col in columns_to_fix:
            data_updates.append({
                "range": f"'{sheet_name}'!{col}{row_num}",
                "values": [[formulas[col]]]
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
        logger.info(f"Fixed {updated} formulas in '{sheet_name}' ({len(slot_rows)} rows × {len(columns_to_fix)} columns)")
        return updated
    return 0


def main():
    logger.info("=" * 60)
    logger.info("GLC Dashboard - Fix All Calculation Formulas Script")
    logger.info("=" * 60)
    logger.info("Fixing columns G, H, K, L, M, N, O")
    logger.info("")

    try:
        service = get_sheets_service()
        logger.info("Connected to Google Sheets API")

        weekly_tabs = get_weekly_tabs(service)
        logger.info(f"Found {len(weekly_tabs)} weekly tabs")

        if not weekly_tabs:
            logger.warning("No weekly tabs found.")
            return

        total_fixed = 0
        for tab_name in weekly_tabs:
            logger.info(f"\nProcessing '{tab_name}'...")
            time.sleep(API_DELAY)
            fixed = fix_all_formulas(service, tab_name)
            total_fixed += fixed
            time.sleep(API_DELAY)

        logger.info("\n" + "=" * 60)
        logger.info(f"Fixed {total_fixed} formulas across {len(weekly_tabs)} tabs")
        logger.info("=" * 60)

    except Exception as e:
        logger.error(f"Error: {e}")
        raise


if __name__ == "__main__":
    main()
