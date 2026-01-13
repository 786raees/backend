"""
Add Payment Status dropdown validation to ALL Turf Supply weekly tabs.
Dropdown options: Unpaid, Paid, Partial
"""
import sys
import json
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build
from app.config import settings

# Fix encoding for Windows console
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

def main():
    """Add dropdown validation to Payment Status column (Column P)."""

    # Get credentials
    service_account_info = settings.google_service_account_json
    if hasattr(service_account_info, 'get_secret_value'):
        service_account_info = service_account_info.get_secret_value()
    if isinstance(service_account_info, str):
        service_account_info = json.loads(service_account_info)

    credentials = service_account.Credentials.from_service_account_info(
        service_account_info,
        scopes=['https://www.googleapis.com/auth/spreadsheets']
    )
    service = build('sheets', 'v4', credentials=credentials)

    # Turf Supply spreadsheet
    spreadsheet_id = '1aygUdgHPMuI14uiGfqZdDFIvBHMmarPwbdfdONml_UI'

    # Get all sheets
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = spreadsheet.get('sheets', [])

    # Filter for weekly tabs
    week_pattern = re.compile(r'^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{2}$')
    week_tabs = []
    sheet_ids = {}

    for sheet in sheets:
        title = sheet['properties']['title']
        sheet_id = sheet['properties']['sheetId']
        if week_pattern.match(title):
            week_tabs.append(title)
            sheet_ids[title] = sheet_id

    week_tabs = sorted(week_tabs)

    print("=" * 80)
    print("ADDING PAYMENT STATUS DROPDOWN - TURF SUPPLY")
    print("=" * 80)
    print(f"\nProcessing {len(week_tabs)} weekly tabs\n")
    print("Dropdown Options: Unpaid, Paid, Partial")
    print("=" * 80)

    # Prepare batch update requests
    requests = []

    for week_tab in week_tabs:
        sheet_id = sheet_ids[week_tab]

        # Column P is index 15 (0-indexed)
        # Apply validation to rows 5-300 (covers all data rows)

        validation_rule = {
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 4,  # Row 5 (0-indexed)
                    "endRowIndex": 300,   # Up to row 300
                    "startColumnIndex": 15,  # Column P
                    "endColumnIndex": 16     # Just column P
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [
                            {"userEnteredValue": "Unpaid"},
                            {"userEnteredValue": "Paid"},
                            {"userEnteredValue": "Partial"}
                        ]
                    },
                    "showCustomUi": True,
                    "strict": False
                }
            }
        }

        requests.append(validation_rule)
        print(f"✓ {week_tab} - Added dropdown validation")

    # Execute batch update
    if requests:
        print(f"\nExecuting batch update for {len(requests)} tabs...")

        body = {
            'requests': requests
        }

        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()

        print(f"\n✓ Successfully added dropdown to all {len(week_tabs)} weekly tabs!")
    else:
        print("\nNo tabs to update.")

    print("=" * 80)
    print("COMPLETE")
    print("=" * 80)


if __name__ == "__main__":
    main()
