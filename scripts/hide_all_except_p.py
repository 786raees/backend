"""
Hide columns K, L, M, N, O, Q but keep P visible.
This will make only column P visible between J and R.

Usage:
    python -m scripts.hide_all_except_p
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

    week_tabs.sort(key=lambda x: x['title'])
    return week_tabs


def configure_columns(service, spreadsheet_id, sheet_id):
    """
    Hide K-O and Q, but show P.
    Column mapping:
    K=10, L=11, M=12, N=13, O=14, P=15, Q=16
    """
    requests = []

    # Hide K through O (indices 10-14)
    for col_idx in [10, 11, 12, 13, 14]:
        requests.append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': col_idx,
                    'endIndex': col_idx + 1
                },
                'properties': {
                    'hiddenByUser': True,
                    'pixelSize': 60
                },
                'fields': 'hiddenByUser,pixelSize'
            }
        })

    # Show P (index 15)
    requests.append({
        'updateDimensionProperties': {
            'range': {
                'sheetId': sheet_id,
                'dimension': 'COLUMNS',
                'startIndex': 15,
                'endIndex': 16
            },
            'properties': {
                'hiddenByUser': False,
                'pixelSize': 150
            },
            'fields': 'hiddenByUser,pixelSize'
        }
    })

    # Hide Q (index 16)
    requests.append({
        'updateDimensionProperties': {
            'range': {
                'sheetId': sheet_id,
                'dimension': 'COLUMNS',
                'startIndex': 16,
                'endIndex': 17
            },
            'properties': {
                'hiddenByUser': True,
                'pixelSize': 60
            },
            'fields': 'hiddenByUser,pixelSize'
        }
    })

    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': requests}
        ).execute()
        return True
    except Exception as e:
        print(f"      Error: {e}")
        return False


def main():
    print("=" * 80)
    print("CONFIGURING COLUMNS: HIDE K-O,Q BUT SHOW P")
    print("=" * 80)
    print()
    print("Target layout: J | P | R")
    print("  - Column J: Visible")
    print("  - Columns K,L,M,N,O: HIDDEN")
    print("  - Column P: VISIBLE (Appointment Confirmed By)")
    print("  - Column Q: HIDDEN")
    print("  - Column R: Visible")
    print()

    service = get_service()
    spreadsheet_id = settings.sales_spreadsheet_id

    print("Fetching week tabs...")
    week_tabs = get_week_tabs(service, spreadsheet_id)
    print(f"Found {len(week_tabs)} week tabs")
    print()

    print("Configuring columns in each tab...")
    print()

    success_count = 0
    fail_count = 0

    for idx, tab_info in enumerate(week_tabs, 1):
        tab_name = tab_info['title']
        sheet_id = tab_info['sheetId']

        print(f"[{idx}/{len(week_tabs)}] {tab_name}...", end=' ', flush=True)

        if configure_columns(service, spreadsheet_id, sheet_id):
            print("✓ Configured")
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
    print("Expected result: Only column P visible between J and R")
    print("Please refresh your Google Sheets to see the changes.")
    print("=" * 80)


if __name__ == "__main__":
    main()
