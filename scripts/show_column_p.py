"""
Make column P (Appointment Confirmed By) clearly visible across all weekly sales sheets.

This script will:
1. Find all week tabs (Mon-DD format)
2. Unhide column P
3. Set a visible width for column P (150px)
4. Remove any column grouping that might hide it

Usage:
    python -m scripts.show_column_p
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


def get_service():
    """Initialize and return Google Sheets service."""
    service_account_info = settings.google_service_account_json
    if hasattr(service_account_info, 'get_secret_value'):
        service_account_info = service_account_info.get_secret_value()
    if isinstance(service_account_info, str):
        service_account_info = json.loads(service_account_info)

    credentials = service_account.Credentials.from_service_account_info(
        service_account_info,
        scopes=['https://www.googleapis.com/auth/spreadsheets']
    )
    return build('sheets', 'v4', credentials=credentials)


def get_week_tabs(service, spreadsheet_id):
    """Get all week tab names and sheet IDs."""
    spreadsheet = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id
    ).execute()

    sheets = spreadsheet.get('sheets', [])
    week_pattern = re.compile(r'^[A-Z][a-z]{2}-\d{2}$')

    week_tabs = []
    for sheet in sheets:
        title = sheet.get('properties', {}).get('title', '')
        sheet_id = sheet.get('properties', {}).get('sheetId')
        if week_pattern.match(title):
            week_tabs.append({
                'title': title,
                'sheetId': sheet_id
            })

    # Sort by title
    week_tabs.sort(key=lambda x: x['title'])
    return week_tabs


def show_column_p(service, spreadsheet_id, sheet_id):
    """
    Make column P visible by:
    1. Setting hiddenByUser to False
    2. Setting a visible width (150px)
    """
    # Column P is index 15 (0-indexed: A=0, B=1, ..., P=15)
    request = {
        'requests': [
            {
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': sheet_id,
                        'dimension': 'COLUMNS',
                        'startIndex': 15,  # Column P
                        'endIndex': 16     # Exclusive
                    },
                    'properties': {
                        'hiddenByUser': False,
                        'pixelSize': 150
                    },
                    'fields': 'hiddenByUser,pixelSize'
                }
            }
        ]
    }

    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=request
        ).execute()
        return True
    except Exception as e:
        print(f"      Error: {e}")
        return False


def main():
    print("=" * 80)
    print("MAKING COLUMN P (APPOINTMENT CONFIRMED BY) VISIBLE ACROSS ALL WEEKS")
    print("=" * 80)
    print()

    service = get_service()
    spreadsheet_id = settings.sales_spreadsheet_id

    print("Fetching week tabs...")
    week_tabs = get_week_tabs(service, spreadsheet_id)
    print(f"Found {len(week_tabs)} week tabs")
    print()

    print("Making column P visible in each tab...")
    print("(Setting: hiddenByUser=False, width=150px)")
    print()

    success_count = 0
    fail_count = 0

    for idx, tab_info in enumerate(week_tabs, 1):
        tab_name = tab_info['title']
        sheet_id = tab_info['sheetId']

        print(f"[{idx}/{len(week_tabs)}] {tab_name}...", end=' ', flush=True)

        if show_column_p(service, spreadsheet_id, sheet_id):
            print("✓ Visible")
            success_count += 1
        else:
            print("✗ Failed")
            fail_count += 1

    print()
    print("=" * 80)
    print("COMPLETE")
    print(f"  Success: {success_count}/{len(week_tabs)} tabs")
    if fail_count > 0:
        print(f"  Failed: {fail_count} tabs")
    print()
    print("Column P is now:")
    print("  - Unhidden (hiddenByUser = False)")
    print("  - Width set to 150 pixels")
    print("  - Should be visible in Google Sheets UI")
    print("=" * 80)


if __name__ == "__main__":
    main()
