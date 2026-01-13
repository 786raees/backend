"""
Force column P to be visible by:
1. Unhiding all columns K-Q first
2. Then hiding only K,L,M,N,O,Q
3. Setting different widths to force UI refresh

Usage:
    python -m scripts.force_show_column_p
"""
import sys
import json
import re
import time
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


def force_show_column_p(service, spreadsheet_id, sheet_id):
    """
    Two-step process:
    Step 1: Unhide ALL columns K-Q
    Step 2: Hide K,L,M,N,O,Q but keep P visible with large width
    """

    # Step 1: Unhide all columns K through Q (10-16)
    requests_step1 = []
    for col_idx in range(10, 17):  # K=10 through Q=16
        requests_step1.append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': col_idx,
                    'endIndex': col_idx + 1
                },
                'properties': {
                    'hiddenByUser': False,
                    'pixelSize': 100
                },
                'fields': 'hiddenByUser,pixelSize'
            }
        })

    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': requests_step1}
        ).execute()
    except Exception as e:
        print(f" Error step 1: {e}")
        return False

    # Small delay to ensure changes propagate
    time.sleep(0.5)

    # Step 2: Hide K,L,M,N,O,Q but keep P visible
    requests_step2 = []

    # Hide K through O (10-14)
    for col_idx in [10, 11, 12, 13, 14]:
        requests_step2.append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': col_idx,
                    'endIndex': col_idx + 1
                },
                'properties': {
                    'hiddenByUser': True
                },
                'fields': 'hiddenByUser'
            }
        })

    # Make P visible with a distinctive width (200px)
    requests_step2.append({
        'updateDimensionProperties': {
            'range': {
                'sheetId': sheet_id,
                'dimension': 'COLUMNS',
                'startIndex': 15,  # P
                'endIndex': 16
            },
            'properties': {
                'hiddenByUser': False,
                'pixelSize': 200
            },
            'fields': 'hiddenByUser,pixelSize'
        }
    })

    # Hide Q (16)
    requests_step2.append({
        'updateDimensionProperties': {
            'range': {
                'sheetId': sheet_id,
                'dimension': 'COLUMNS',
                'startIndex': 16,
                'endIndex': 17
            },
            'properties': {
                'hiddenByUser': True
            },
            'fields': 'hiddenByUser'
        }
    })

    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': requests_step2}
        ).execute()
        return True
    except Exception as e:
        print(f" Error step 2: {e}")
        return False


def main():
    print("=" * 80)
    print("FORCE UNHIDING COLUMN P (TWO-STEP PROCESS)")
    print("=" * 80)
    print()
    print("Step 1: Unhide ALL columns K-Q")
    print("Step 2: Hide K,L,M,N,O,Q but keep P visible (200px width)")
    print()

    service = get_service()
    spreadsheet_id = settings.sales_spreadsheet_id

    print("Fetching week tabs...")
    week_tabs = get_week_tabs(service, spreadsheet_id)
    print(f"Found {len(week_tabs)} week tabs")
    print()

    print("Processing each tab (two-step unhide)...")
    print()

    success_count = 0
    fail_count = 0

    for idx, tab_info in enumerate(week_tabs, 1):
        tab_name = tab_info['title']
        sheet_id = tab_info['sheetId']

        print(f"[{idx}/{len(week_tabs)}] {tab_name}...", end=' ', flush=True)

        if force_show_column_p(service, spreadsheet_id, sheet_id):
            print("✓ Done")
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
    print("Column P should now be visible with 200px width.")
    print("IMPORTANT: Hard refresh your browser (Ctrl+Shift+R) or clear cache.")
    print("=" * 80)


if __name__ == "__main__":
    main()
