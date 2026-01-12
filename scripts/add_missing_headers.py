"""
Add missing column headers (J-Q) to all weekly sales sheets.

This script will:
1. Find all week tabs (Mon-DD format)
2. Locate the header row for each rep section
3. Add missing headers for columns J-Q

Usage:
    python scripts/add_missing_headers.py           # Interactive mode
    python scripts/add_missing_headers.py --yes     # Auto-confirm
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


# Column headers mapping
HEADERS = {
    'J': 'Sell Price Ex GST',
    'K': 'Appointment Time',
    'L': 'Project Type',
    'M': 'Suburb',
    'N': 'Region',
    'O': 'Appointment Set Who',
    'P': 'Appointment Confirmed By',
    'Q': 'Gross Profit Margin %'
}


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
    """Get all week tab names."""
    spreadsheet = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id
    ).execute()

    sheets = spreadsheet.get('sheets', [])
    week_pattern = re.compile(r'^[A-Z][a-z]{2}-\d{2}$')

    week_tabs = []
    for sheet in sheets:
        title = sheet.get('properties', {}).get('title', '')
        if week_pattern.match(title):
            week_tabs.append(title)

    return sorted(week_tabs)


def find_header_rows(service, spreadsheet_id, week_tab):
    """Find all header rows in a week tab (one per rep per day)."""
    # Read first 150 rows to cover all reps across all days
    range_name = f"'{week_tab}'!A1:I150"
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()

    values = result.get('values', [])
    header_rows = []

    for i, row in enumerate(values, 1):
        if len(row) > 0 and row[0] == 'Slot':
            header_rows.append(i)

    return header_rows


def add_headers_to_row(service, spreadsheet_id, week_tab, row_number):
    """Add missing headers J-Q to a specific row."""
    updates = []
    for col_letter, header_text in HEADERS.items():
        cell_ref = f"'{week_tab}'!{col_letter}{row_number}"
        updates.append({
            'range': cell_ref,
            'values': [[header_text]]
        })

    # Batch update all headers at once
    body = {
        'valueInputOption': 'RAW',
        'data': updates
    }

    service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=body
    ).execute()


def main():
    print("=" * 80)
    print("ADDING MISSING COLUMN HEADERS TO SALES SHEETS")
    print("=" * 80)

    service = get_service()
    spreadsheet_id = settings.sales_spreadsheet_id

    # Get all week tabs
    print("\nFetching week tabs...")
    week_tabs = get_week_tabs(service, spreadsheet_id)
    print(f"Found {len(week_tabs)} week tabs")

    # Ask for confirmation
    print(f"\nThis will add headers J-Q to ALL {len(week_tabs)} week tabs.")
    print("Headers to be added:")
    for col_letter, header_text in HEADERS.items():
        print(f"  {col_letter}: {header_text}")

    # Check for --yes flag
    auto_confirm = '--yes' in sys.argv or '-y' in sys.argv

    if auto_confirm:
        print("\nAuto-confirmed with --yes flag")
    else:
        response = input("\nProceed? (yes/no): ").strip().lower()
        if response != 'yes':
            print("Aborted.")
            return

    total_headers_added = 0

    # Process each week tab
    for tab_idx, week_tab in enumerate(week_tabs, 1):
        print(f"\n[{tab_idx}/{len(week_tabs)}] Processing {week_tab}...")

        # Find header rows
        header_rows = find_header_rows(service, spreadsheet_id, week_tab)
        print(f"  Found {len(header_rows)} header rows at: {header_rows}")

        if not header_rows:
            print(f"  ⚠ No header rows found, skipping")
            continue

        # Add headers to each row
        for row_num in header_rows:
            try:
                add_headers_to_row(service, spreadsheet_id, week_tab, row_num)
                print(f"  ✓ Added headers to row {row_num}")
                total_headers_added += 1
            except Exception as e:
                print(f"  ✗ Error on row {row_num}: {e}")

    print("\n" + "=" * 80)
    print(f"COMPLETE - Updated {total_headers_added} header rows across {len(week_tabs)} tabs")
    print("=" * 80)


if __name__ == "__main__":
    main()
