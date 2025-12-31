"""
Google Sheet Setup Script for GLC Dashboard
============================================

This script updates the Google Sheet with:
1. Setup sheet - Pricing reference table
2. Lists sheet - Dropdown values
3. Weekly tabs - Formulas and data validation

Run from the backend directory:
    python scripts/setup_google_sheet.py
"""

import os
import sys
import json
import logging
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

# Pricing data for Setup sheet
PRICING_DATA = [
    ["Turf Type", "Sell $/SQM", "Cost $/SQM"],
    ["Empire Zoysia", 12.00, 8.25],
    ["Sir Walter", 13.00, 8.25],
    ["Wintergreen Couch", 6.60, 3.19],
    ["AussiBlue Couch", 8.80, 5.50],
    ["Summerland Buffalo", 13.00, 8.25],
]

# Dropdown values for Lists sheet
LISTS_DATA = [
    ["Varieties", "Service Types"],
    ["Empire Zoysia", "SL"],
    ["Sir Walter", "SD"],
    ["Wintergreen Couch", "P"],
    ["AussiBlue Couch", ""],
    ["Summerland Buffalo", ""],
]

# Column headers for pricing section (G-O)
PRICING_HEADERS = [
    "Sell $/SQM", "Cost $/SQM", "Delivery Fee", "Laying Fee",
    "Turf Revenue", "Turf Cost", "Total Revenue", "Gross Profit", "Margin %"
]


def get_credentials():
    """Get Google API credentials from environment or .env file."""
    # Try to load from environment variable first
    creds_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")

    if not creds_json:
        # Try to load from .env file
        env_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), ".env")
        if os.path.exists(env_path):
            with open(env_path, "r") as f:
                for line in f:
                    if line.startswith("GOOGLE_SERVICE_ACCOUNT_JSON="):
                        creds_json = line.split("=", 1)[1].strip()
                        break

    if not creds_json:
        raise ValueError("GOOGLE_SERVICE_ACCOUNT_JSON not found in environment or .env file")

    # Parse JSON credentials
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


def get_existing_sheets(service) -> List[str]:
    """Get list of existing sheet names in the spreadsheet."""
    result = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = result.get("sheets", [])
    return [sheet["properties"]["title"] for sheet in sheets]


def create_sheet_if_not_exists(service, sheet_name: str) -> int:
    """Create a new sheet if it doesn't exist. Returns sheet ID."""
    existing = get_existing_sheets(service)

    if sheet_name in existing:
        logger.info(f"Sheet '{sheet_name}' already exists")
        # Get sheet ID
        result = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        for sheet in result.get("sheets", []):
            if sheet["properties"]["title"] == sheet_name:
                return sheet["properties"]["sheetId"]
        return 0

    # Create new sheet
    request = {
        "requests": [{
            "addSheet": {
                "properties": {
                    "title": sheet_name
                }
            }
        }]
    }
    result = service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body=request
    ).execute()

    sheet_id = result["replies"][0]["addSheet"]["properties"]["sheetId"]
    logger.info(f"Created sheet '{sheet_name}' with ID {sheet_id}")
    return sheet_id


def write_data_to_sheet(service, sheet_name: str, data: List[List[Any]], start_cell: str = "A1"):
    """Write data to a sheet."""
    range_name = f"'{sheet_name}'!{start_cell}"
    body = {"values": data}

    result = service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_name,
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()

    logger.info(f"Wrote {result.get('updatedRows', 0)} rows to '{sheet_name}'")


def setup_pricing_sheet(service):
    """Create and populate the Setup sheet with pricing data."""
    logger.info("Setting up 'Setup' sheet...")
    create_sheet_if_not_exists(service, "Setup")
    write_data_to_sheet(service, "Setup", PRICING_DATA)
    logger.info("Setup sheet created with pricing data")


def setup_lists_sheet(service):
    """Create and populate the Lists sheet with dropdown values."""
    logger.info("Setting up 'Lists' sheet...")
    create_sheet_if_not_exists(service, "Lists")
    write_data_to_sheet(service, "Lists", LISTS_DATA)
    logger.info("Lists sheet created with dropdown values")


def get_weekly_tabs(service) -> List[str]:
    """Get list of weekly tab names (format: Mon-DD like Dec-29, Jan-05)."""
    import re
    existing = get_existing_sheets(service)
    pattern = r'^[A-Z][a-z]{2}-\d{2}$'
    weekly_tabs = [name for name in existing if re.match(pattern, name)]
    logger.info(f"Found {len(weekly_tabs)} weekly tabs: {weekly_tabs}")
    return weekly_tabs


def get_sheet_id_by_name(service, sheet_name: str) -> int:
    """Get sheet ID by name."""
    result = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for sheet in result.get("sheets", []):
        if sheet["properties"]["title"] == sheet_name:
            return sheet["properties"]["sheetId"]
    return None


def add_data_validation_dropdown(service, sheet_id: int, start_row: int, end_row: int, column: int, source_range: str):
    """Add data validation (dropdown) to a range of cells."""
    request = {
        "requests": [{
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_row,
                    "endRowIndex": end_row,
                    "startColumnIndex": column,
                    "endColumnIndex": column + 1
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_RANGE",
                        "values": [{
                            "userEnteredValue": source_range
                        }]
                    },
                    "showCustomUi": True,
                    "strict": False
                }
            }
        }]
    }

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body=request
    ).execute()


def find_slot_rows(service, sheet_name: str) -> List[Dict]:
    """
    Find all slot rows in a weekly sheet.
    Returns list of dicts with row info: {row_index, truck, day, slot}
    """
    # Read the entire sheet
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

        # Check for slot row (1-6)
        try:
            slot_num = int(cell)
            if 1 <= slot_num <= 6 and current_day and current_truck:
                slot_rows.append({
                    "row_index": idx,  # 0-indexed
                    "row_number": idx + 1,  # 1-indexed for formulas
                    "day": current_day,
                    "truck": current_truck,
                    "slot": slot_num
                })
        except (ValueError, TypeError):
            pass

    return slot_rows


def generate_formulas_for_row(row_number: int) -> List[str]:
    """
    Generate formulas for columns F-O for a given row.

    Column layout:
    A: Slot
    B: Variety
    C: Suburb
    D: Service Type
    E: SQM Sold
    F: Pallets (formula)
    G: Sell $/SQM (formula)
    H: Cost $/SQM (formula)
    I: Delivery Fee (manual)
    J: Laying Fee (manual)
    K: Turf Revenue (formula)
    L: Turf Cost (formula)
    M: Total Revenue (formula)
    N: Gross Profit (formula)
    O: Margin % (formula)
    """
    r = row_number

    # F: Pallets formula
    pallets = f'=IF(E{r}="","",IF(OR(B{r}="Empire Zoysia",B{r}="Sir Walter",B{r}="Summerland Buffalo"),E{r}/50,IF(OR(B{r}="Wintergreen Couch",B{r}="AussiBlue Couch"),E{r}/60,"")))'

    # G: Sell $/SQM
    sell_price = f'=IF(B{r}="","",VLOOKUP(B{r},Setup!$A$2:$C$6,2,FALSE))'

    # H: Cost $/SQM
    cost_price = f'=IF(B{r}="","",VLOOKUP(B{r},Setup!$A$2:$C$6,3,FALSE))'

    # I, J: Delivery Fee, Laying Fee - empty (manual entry)
    delivery_fee = ""
    laying_fee = ""

    # K: Turf Revenue
    turf_revenue = f'=IF(OR(E{r}="",G{r}=""),"",E{r}*G{r})'

    # L: Turf Cost
    turf_cost = f'=IF(OR(E{r}="",H{r}=""),"",E{r}*H{r})'

    # M: Total Revenue
    total_revenue = f'=IF(K{r}="","",K{r}+IF(I{r}="",0,I{r})+IF(J{r}="",0,J{r}))'

    # N: Gross Profit
    gross_profit = f'=IF(OR(M{r}="",L{r}=""),"",M{r}-L{r})'

    # O: Margin %
    margin = f'=IF(OR(M{r}="",M{r}=0),"",N{r}/M{r})'

    return [pallets, sell_price, cost_price, delivery_fee, laying_fee,
            turf_revenue, turf_cost, total_revenue, gross_profit, margin]


def update_weekly_tab_formulas(service, sheet_name: str):
    """Update a weekly tab with formulas for all slot rows."""
    logger.info(f"Updating formulas in '{sheet_name}'...")

    # Find all slot rows
    slot_rows = find_slot_rows(service, sheet_name)

    if not slot_rows:
        logger.warning(f"No slot rows found in '{sheet_name}'")
        return

    logger.info(f"Found {len(slot_rows)} slot rows in '{sheet_name}'")

    # Prepare batch update data
    data_updates = []

    for slot_info in slot_rows:
        row_num = slot_info["row_number"]
        formulas = generate_formulas_for_row(row_num)

        # Update columns F-O (indices 5-14)
        data_updates.append({
            "range": f"'{sheet_name}'!F{row_num}:O{row_num}",
            "values": [formulas]
        })

    # Batch update all formulas
    if data_updates:
        body = {
            "valueInputOption": "USER_ENTERED",
            "data": data_updates
        }

        result = service.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=body
        ).execute()

        logger.info(f"Updated {result.get('totalUpdatedRows', 0)} rows with formulas in '{sheet_name}'")


def add_dropdowns_to_weekly_tab(service, sheet_name: str):
    """Add data validation dropdowns to a weekly tab."""
    logger.info(f"Adding dropdowns to '{sheet_name}'...")

    sheet_id = get_sheet_id_by_name(service, sheet_name)
    if sheet_id is None:
        logger.error(f"Could not find sheet ID for '{sheet_name}'")
        return

    # Find all slot rows to determine range
    slot_rows = find_slot_rows(service, sheet_name)

    if not slot_rows:
        logger.warning(f"No slot rows found in '{sheet_name}'")
        return

    # Get min and max row indices
    min_row = min(r["row_index"] for r in slot_rows)
    max_row = max(r["row_index"] for r in slot_rows) + 1  # end_row is exclusive

    try:
        # Add Variety dropdown (column B = index 1)
        add_data_validation_dropdown(
            service, sheet_id,
            start_row=min_row, end_row=max_row,
            column=1,  # Column B
            source_range="=Lists!$A$2:$A$6"
        )
        logger.info(f"Added Variety dropdown to column B, rows {min_row+1}-{max_row}")

        # Add Service Type dropdown (column D = index 3)
        add_data_validation_dropdown(
            service, sheet_id,
            start_row=min_row, end_row=max_row,
            column=3,  # Column D
            source_range="=Lists!$B$2:$B$4"
        )
        logger.info(f"Added Service Type dropdown to column D, rows {min_row+1}-{max_row}")

    except HttpError as e:
        logger.error(f"Error adding dropdowns: {e}")


def update_header_rows(service, sheet_name: str):
    """Add headers for columns G-O in the header rows after each TRUCK header."""
    logger.info(f"Updating headers in '{sheet_name}'...")

    # Read the sheet to find header rows
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A:F"
    ).execute()

    values = result.get("values", [])
    header_rows = []

    for idx, row in enumerate(values):
        if not row:
            continue
        cell = str(row[0]).strip().lower() if row else ""
        if cell == "slot":
            header_rows.append(idx + 1)  # 1-indexed

    if not header_rows:
        logger.warning(f"No header rows found in '{sheet_name}'")
        return

    # Update each header row with G-O headers
    data_updates = []
    for row_num in header_rows:
        data_updates.append({
            "range": f"'{sheet_name}'!G{row_num}:O{row_num}",
            "values": [PRICING_HEADERS]
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

        logger.info(f"Updated {len(header_rows)} header rows in '{sheet_name}'")


def main():
    """Main function to set up the Google Sheet."""
    logger.info("=" * 60)
    logger.info("GLC Dashboard - Google Sheet Setup Script")
    logger.info("=" * 60)

    try:
        service = get_sheets_service()
        logger.info("Connected to Google Sheets API")

        # Step 1: Create Setup sheet with pricing
        setup_pricing_sheet(service)

        # Step 2: Create Lists sheet with dropdowns
        setup_lists_sheet(service)

        # Step 3: Get all weekly tabs
        weekly_tabs = get_weekly_tabs(service)

        if not weekly_tabs:
            logger.warning("No weekly tabs found. The sheet structure may be different.")
            logger.info("Please ensure weekly tabs are named like 'Dec-29', 'Jan-05', etc.")
            return

        # Step 4: Update each weekly tab
        for tab_name in weekly_tabs:
            logger.info(f"\nProcessing '{tab_name}'...")

            # Add headers for new columns
            update_header_rows(service, tab_name)

            # Add formulas to slot rows
            update_weekly_tab_formulas(service, tab_name)

            # Add dropdowns
            add_dropdowns_to_weekly_tab(service, tab_name)

        logger.info("\n" + "=" * 60)
        logger.info("Setup complete!")
        logger.info("=" * 60)
        logger.info("\nNext steps:")
        logger.info("1. Open the Google Sheet and verify the changes")
        logger.info("2. Test the dropdowns in column B (Variety) and D (Service Type)")
        logger.info("3. Enter some SQM values and verify Pallets auto-calculates")
        logger.info("4. Check the TV dashboard to ensure it still displays correctly")

    except Exception as e:
        logger.error(f"Error: {e}")
        raise


if __name__ == "__main__":
    main()
