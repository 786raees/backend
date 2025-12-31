"""
Fix Google Sheet Styling
========================

This script fixes the styling and formatting of the weekly tabs:
1. Cleans up duplicate headers
2. Formats the header rows properly
3. Applies consistent styling

Run from the backend directory:
    python scripts/fix_sheet_styling.py
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

# Brand colors
BRAND_GREEN = {"red": 0.34, "green": 0.57, "blue": 0.29}  # #57924b
HEADER_DARK = {"red": 0.2, "green": 0.2, "blue": 0.2}  # Dark gray
WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}
LIGHT_GRAY = {"red": 0.95, "green": 0.95, "blue": 0.95}


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


def get_sheet_id_by_name(service, sheet_name: str) -> int:
    """Get sheet ID by name."""
    result = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for sheet in result.get("sheets", []):
        if sheet["properties"]["title"] == sheet_name:
            return sheet["properties"]["sheetId"]
    return None


def get_weekly_tabs(service) -> List[str]:
    """Get list of weekly tab names."""
    import re
    result = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = result.get("sheets", [])
    names = [sheet["properties"]["title"] for sheet in sheets]
    pattern = r'^[A-Z][a-z]{2}-\d{2}$'
    weekly_tabs = [name for name in names if re.match(pattern, name)]
    return weekly_tabs


def find_header_rows(service, sheet_name: str) -> List[int]:
    """Find all header rows (rows with 'Slot' in column A)."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A:A"
    ).execute()

    values = result.get("values", [])
    header_rows = []

    for idx, row in enumerate(values):
        if row and str(row[0]).strip().lower() == "slot":
            header_rows.append(idx)  # 0-indexed

    return header_rows


def clear_extra_header_content(service, sheet_name: str):
    """Clear any misplaced header content and fix the structure."""
    logger.info(f"Cleaning up '{sheet_name}'...")

    # Read the sheet
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A:O"
    ).execute()

    values = result.get("values", [])
    if not values:
        return

    # Find header rows and clear G-O only in header rows
    header_rows = find_header_rows(service, sheet_name)

    # The correct headers for G-O
    correct_headers = ["Sell $/SQM", "Cost $/SQM", "Delivery Fee", "Laying Fee",
                       "Turf Revenue", "Turf Cost", "Total Revenue", "Gross Profit", "Margin %"]

    updates = []

    for row_idx in header_rows:
        row_num = row_idx + 1  # 1-indexed
        updates.append({
            "range": f"'{sheet_name}'!G{row_num}:O{row_num}",
            "values": [correct_headers]
        })

    if updates:
        body = {
            "valueInputOption": "USER_ENTERED",
            "data": updates
        }
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=body
        ).execute()
        logger.info(f"Fixed {len(header_rows)} header rows in '{sheet_name}'")


def apply_formatting(service, sheet_name: str):
    """Apply consistent formatting to the sheet."""
    sheet_id = get_sheet_id_by_name(service, sheet_name)
    if sheet_id is None:
        logger.error(f"Could not find sheet ID for '{sheet_name}'")
        return

    # Read sheet to find structure
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A:A"
    ).execute()

    values = result.get("values", [])

    requests = []

    for idx, row in enumerate(values):
        if not row:
            continue

        cell = str(row[0]).strip()
        row_idx = idx  # 0-indexed for API

        # Format TRUCK headers (dark background, white text, bold)
        if cell.upper() in ["TRUCK 1", "TRUCK 2"]:
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": row_idx,
                        "endRowIndex": row_idx + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 15  # A-O
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": HEADER_DARK,
                            "textFormat": {
                                "foregroundColor": WHITE,
                                "bold": True,
                                "fontSize": 11
                            }
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat)"
                }
            })

        # Format column header rows (light gray background, bold)
        elif cell.lower() == "slot":
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": row_idx,
                        "endRowIndex": row_idx + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 15  # A-O
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": LIGHT_GRAY,
                            "textFormat": {
                                "bold": True,
                                "fontSize": 10
                            }
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat)"
                }
            })

        # Format Totals rows (bold)
        elif cell.lower() == "totals":
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": row_idx,
                        "endRowIndex": row_idx + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 15  # A-O
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "bold": True
                            },
                            "borders": {
                                "top": {
                                    "style": "SOLID",
                                    "width": 1
                                }
                            }
                        }
                    },
                    "fields": "userEnteredFormat(textFormat,borders)"
                }
            })

        # Format Day headers (green background for days)
        elif any(cell.startswith(day) for day in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]):
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": row_idx,
                        "endRowIndex": row_idx + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 15  # A-O
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": BRAND_GREEN,
                            "textFormat": {
                                "foregroundColor": WHITE,
                                "bold": True,
                                "fontSize": 12
                            }
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat)"
                }
            })

    # Set column widths
    column_widths = [
        (0, 50),   # A - Slot
        (1, 140),  # B - Variety
        (2, 120),  # C - Suburb
        (3, 100),  # D - Service Type
        (4, 80),   # E - SQM Sold
        (5, 70),   # F - Pallets
        (6, 90),   # G - Sell $/SQM
        (7, 90),   # H - Cost $/SQM
        (8, 100),  # I - Delivery Fee
        (9, 90),   # J - Laying Fee
        (10, 100), # K - Turf Revenue
        (11, 90),  # L - Turf Cost
        (12, 110), # M - Total Revenue
        (13, 100), # N - Gross Profit
        (14, 80),  # O - Margin %
    ]

    for col_idx, width in column_widths:
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": col_idx,
                    "endIndex": col_idx + 1
                },
                "properties": {
                    "pixelSize": width
                },
                "fields": "pixelSize"
            }
        })

    # Format number columns
    # Margin % column (O) - format as percentage
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 0,
                "endRowIndex": 200,
                "startColumnIndex": 14,  # Column O
                "endColumnIndex": 15
            },
            "cell": {
                "userEnteredFormat": {
                    "numberFormat": {
                        "type": "PERCENT",
                        "pattern": "0.00%"
                    }
                }
            },
            "fields": "userEnteredFormat.numberFormat"
        }
    })

    # Currency columns (G, H, K, L, M, N) - format as currency
    for col in [6, 7, 10, 11, 12, 13]:  # G, H, K, L, M, N
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 200,
                    "startColumnIndex": col,
                    "endColumnIndex": col + 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": "CURRENCY",
                            "pattern": "$#,##0.00"
                        }
                    }
                },
                "fields": "userEnteredFormat.numberFormat"
            }
        })

    # Execute all formatting requests
    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": requests}
        ).execute()
        logger.info(f"Applied formatting to '{sheet_name}'")


def main():
    """Main function to fix sheet styling."""
    logger.info("=" * 60)
    logger.info("GLC Dashboard - Fix Sheet Styling")
    logger.info("=" * 60)

    try:
        service = get_sheets_service()
        logger.info("Connected to Google Sheets API")

        # Get all weekly tabs
        weekly_tabs = get_weekly_tabs(service)
        logger.info(f"Found {len(weekly_tabs)} weekly tabs")

        # Skip tabs that were already processed (resume from Jun-15)
        start_from = "Jun-15"
        skip = True

        # Fix each weekly tab
        for tab_name in weekly_tabs:
            if skip:
                if tab_name == start_from:
                    skip = False
                else:
                    logger.info(f"Skipping '{tab_name}' (already processed)")
                    continue

            logger.info(f"\nProcessing '{tab_name}'...")

            # Rate limiting - sleep 2 seconds between tabs
            time.sleep(2)

            # Clean up headers
            clear_extra_header_content(service, tab_name)

            # Apply formatting
            apply_formatting(service, tab_name)

        logger.info("\n" + "=" * 60)
        logger.info("Styling fix complete!")
        logger.info("=" * 60)

    except Exception as e:
        logger.error(f"Error: {e}")
        raise


if __name__ == "__main__":
    main()
