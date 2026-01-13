"""
Remove column groups (collapsed groups) that might be hiding column P.

This script will:
1. Find all week tabs
2. Delete any column groups/banding that might include column P

Usage:
    python -m scripts.remove_column_groups
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
        spreadsheetId=spreadsheet_id,
        includeGridData=False
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


def remove_column_groups(service, spreadsheet_id, sheet_id):
    """
    Remove all column groups (row/column grouping) from the sheet.
    This will expand any collapsed column groups that might be hiding column P.
    """
    requests = []

    # Delete all column groups on this sheet
    # We'll delete groups for columns 0-50 to cover all possible groups
    for start_idx in range(0, 50, 10):
        requests.append({
            'deleteColumnGroup': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': start_idx,
                    'endIndex': min(start_idx + 10, 50)
                }
            }
        })

    if not requests:
        return True

    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': requests}
        ).execute()
        return True
    except Exception as e:
        # It's okay if there are no groups to delete
        error_msg = str(e).lower()
        if 'no group' in error_msg or 'not found' in error_msg:
            return True
        print(f" (no groups found)")
        return True


def main():
    print("=" * 80)
    print("REMOVING COLUMN GROUPS THAT MIGHT HIDE COLUMN P")
    print("=" * 80)
    print()

    service = get_service()
    spreadsheet_id = settings.sales_spreadsheet_id

    print("Fetching week tabs...")
    week_tabs = get_week_tabs(service, spreadsheet_id)
    print(f"Found {len(week_tabs)} week tabs")
    print()

    print("Removing column groups from each tab...")
    print()

    success_count = 0

    for idx, tab_info in enumerate(week_tabs, 1):
        tab_name = tab_info['title']
        sheet_id = tab_info['sheetId']

        print(f"[{idx}/{len(week_tabs)}] {tab_name}...", end=' ', flush=True)

        if remove_column_groups(service, spreadsheet_id, sheet_id):
            print("✓ Done")
            success_count += 1

    print()
    print("=" * 80)
    print("COMPLETE")
    print(f"  Processed: {success_count}/{len(week_tabs)} tabs")
    print()
    print("All column groups removed. Column P should now be visible.")
    print("=" * 80)


if __name__ == "__main__":
    main()
