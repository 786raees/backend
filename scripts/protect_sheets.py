"""
Script to hide calculation columns and protect sheets for staff view.

This script:
1. Hides financial/calculation columns (G-Q) from staff
2. Protects formula cells from editing
3. Allows editing only in data entry columns (A-F, I, J)

Run this script once to set up protections on all weekly sheets.

Usage:
    cd backend
    python scripts/protect_sheets.py
"""
import os
import sys
import logging
from datetime import date

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dotenv import load_dotenv
load_dotenv()

from app.services.google.sheets_client import GoogleSheetsClient
from app.config import settings

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Columns to hide from staff view (0-indexed)
# G=6 (Sell $/SQM) through Q=16 (Margin %)
COLUMNS_TO_HIDE = {
    'start': 6,   # Column G (Sell $/SQM)
    'end': 17     # Through Column Q (Margin %)
}

# Columns staff can edit (0-indexed)
# A=Day, B=Variety, C=Suburb, D=Service Type, E=SQM, I=Delivery Fee, J=Laying Fee
EDITABLE_COLUMNS = [0, 1, 2, 3, 4, 8, 9]  # A, B, C, D, E, I, J

# Manager email who can edit everything
MANAGER_EMAIL = os.getenv('MANAGER_EMAIL', 'manager@greatlawnco.com.au')


def get_sheet_id(client: GoogleSheetsClient, sheet_name: str) -> int:
    """Get the sheet ID for a named sheet."""
    sheets = client.get_available_sheets()

    # Get sheet metadata
    spreadsheet = client._service.spreadsheets().get(
        spreadsheetId=settings.google_spreadsheet_id
    ).execute()

    for sheet in spreadsheet.get('sheets', []):
        props = sheet.get('properties', {})
        if props.get('title') == sheet_name:
            return props.get('sheetId')

    return None


def hide_columns(client: GoogleSheetsClient, sheet_id: int, sheet_name: str):
    """Hide calculation columns from staff view."""
    logger.info(f"Hiding columns G-Q in '{sheet_name}'")

    requests = [
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": COLUMNS_TO_HIDE['start'],
                    "endIndex": COLUMNS_TO_HIDE['end']
                },
                "properties": {"hiddenByUser": True},
                "fields": "hiddenByUser"
            }
        }
    ]

    client._service.spreadsheets().batchUpdate(
        spreadsheetId=settings.google_spreadsheet_id,
        body={"requests": requests}
    ).execute()

    logger.info(f"  Columns G-Q hidden in '{sheet_name}'")


def protect_sheet(client: GoogleSheetsClient, sheet_id: int, sheet_name: str):
    """Protect formula columns, allow editing only in data entry columns."""
    logger.info(f"Protecting formulas in '{sheet_name}'")

    # First, remove any existing protections to avoid duplicates
    spreadsheet = client._service.spreadsheets().get(
        spreadsheetId=settings.google_spreadsheet_id,
        fields="sheets.properties,sheets.protectedRanges"
    ).execute()

    existing_protections = []
    for sheet in spreadsheet.get('sheets', []):
        if sheet.get('properties', {}).get('sheetId') == sheet_id:
            existing_protections = sheet.get('protectedRanges', [])
            break

    # Delete existing protections
    if existing_protections:
        delete_requests = [
            {"deleteProtectedRange": {"protectedRangeId": p['protectedRangeId']}}
            for p in existing_protections
        ]
        client._service.spreadsheets().batchUpdate(
            spreadsheetId=settings.google_spreadsheet_id,
            body={"requests": delete_requests}
        ).execute()
        logger.info(f"  Removed {len(existing_protections)} existing protections")

    # Add protection for formula columns (F=5 onwards except I, J)
    # Protect column F (Pallets - formula)
    # Protect columns K-Q (financial formulas)
    requests = [
        # Protect column F (Pallets - auto-calculated)
        {
            "addProtectedRange": {
                "protectedRange": {
                    "range": {
                        "sheetId": sheet_id,
                        "startColumnIndex": 5,  # Column F
                        "endColumnIndex": 6     # Just column F
                    },
                    "description": "Pallets formula - protected",
                    "warningOnly": False,
                    "editors": {"users": [MANAGER_EMAIL]}
                }
            }
        },
        # Protect columns K-Q (financial formulas)
        {
            "addProtectedRange": {
                "protectedRange": {
                    "range": {
                        "sheetId": sheet_id,
                        "startColumnIndex": 10,  # Column K (Turf Revenue)
                        "endColumnIndex": 17     # Through Column Q
                    },
                    "description": "Financial formulas - protected",
                    "warningOnly": False,
                    "editors": {"users": [MANAGER_EMAIL]}
                }
            }
        }
    ]

    client._service.spreadsheets().batchUpdate(
        spreadsheetId=settings.google_spreadsheet_id,
        body={"requests": requests}
    ).execute()

    logger.info(f"  Protected formula columns in '{sheet_name}'")


def process_all_weekly_sheets(client: GoogleSheetsClient):
    """Process all weekly sheets to hide columns and protect formulas."""
    sheets = client.get_available_sheets()

    # Filter to only weekly sheets (format like "Dec-28", "Jan-05")
    import re
    week_pattern = re.compile(r'^[A-Z][a-z]{2}-\d{2}$')
    weekly_sheets = [s for s in sheets if week_pattern.match(s)]

    logger.info(f"Found {len(weekly_sheets)} weekly sheets to process")

    processed = 0
    errors = []

    for sheet_name in weekly_sheets:
        try:
            sheet_id = get_sheet_id(client, sheet_name)
            if sheet_id is None:
                logger.warning(f"Could not find sheet ID for '{sheet_name}'")
                continue

            hide_columns(client, sheet_id, sheet_name)
            protect_sheet(client, sheet_id, sheet_name)
            processed += 1

        except Exception as e:
            logger.error(f"Error processing '{sheet_name}': {e}")
            errors.append((sheet_name, str(e)))

    logger.info(f"\nCompleted: {processed}/{len(weekly_sheets)} sheets processed")
    if errors:
        logger.warning(f"Errors ({len(errors)}):")
        for name, err in errors:
            logger.warning(f"  - {name}: {err}")


def main():
    """Main entry point."""
    logger.info("=" * 60)
    logger.info("GLC Dashboard - Sheet Protection Script")
    logger.info("=" * 60)

    if not settings.has_google_credentials:
        logger.error("Google credentials not configured. Set GOOGLE_SERVICE_ACCOUNT_JSON in .env")
        sys.exit(1)

    logger.info(f"Spreadsheet ID: {settings.google_spreadsheet_id}")
    logger.info(f"Manager email: {MANAGER_EMAIL}")
    logger.info("")

    # Confirm before proceeding
    print("\nThis will:")
    print("  1. Hide columns G-Q (financial data) on all weekly sheets")
    print("  2. Protect formula columns (F, K-Q) from staff editing")
    print(f"  3. Only '{MANAGER_EMAIL}' can edit protected cells")
    print("")

    confirm = input("Proceed? (yes/no): ").strip().lower()
    if confirm != 'yes':
        logger.info("Aborted by user")
        sys.exit(0)

    logger.info("\nStarting sheet protection...")

    client = GoogleSheetsClient()
    process_all_weekly_sheets(client)

    logger.info("\nDone!")


if __name__ == "__main__":
    main()
