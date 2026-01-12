"""
Add missing column headers (J-R) to all weekly sales sheets with rate limiting.

This script will:
1. Find all week tabs (Mon-DD format)
2. Locate the header row for each rep section
3. Add missing headers for columns J-R (including Paid/Unpaid)
4. Use rate limiting to avoid quota errors

Usage:
    python scripts/add_all_headers_with_rate_limit.py --yes
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


# Column headers mapping (J through R)
HEADERS = {
    'J': 'Sell Price Ex GST',
    'K': 'Appointment Time',
    'L': 'Project Type',
    'M': 'Suburb',
    'N': 'Region',
    'O': 'Appointment Set Who',
    'P': 'Appointment Confirmed By',
    'Q': 'Gross Profit Margin %',
    'R': 'Paid/Unpaid'
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
    """Get all week tab names sorted chronologically."""
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


def check_missing_headers(service, spreadsheet_id, week_tab, row_number):
    """Check which headers are missing in a specific row."""
    range_name = f"'{week_tab}'!J{row_number}:R{row_number}"
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()

    existing_values = result.get('values', [[]])[0] if result.get('values') else []

    missing_headers = {}
    for i, (col_letter, header_text) in enumerate(HEADERS.items()):
        if i >= len(existing_values) or not existing_values[i]:
            missing_headers[col_letter] = header_text

    return missing_headers


def add_headers_to_row(service, spreadsheet_id, week_tab, row_number, headers_to_add):
    """Add specific missing headers to a row."""
    if not headers_to_add:
        return True

    updates = []
    for col_letter, header_text in headers_to_add.items():
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

    try:
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
        return True
    except Exception as e:
        if '429' in str(e):
            # Rate limit error
            return False
        raise


def main():
    print("=" * 80)
    print("ADDING MISSING COLUMN HEADERS (J-R) TO SALES SHEETS")
    print("=" * 80)

    service = get_service()
    spreadsheet_id = settings.sales_spreadsheet_id

    # Get all week tabs
    print("\nFetching week tabs...")
    week_tabs = get_week_tabs(service, spreadsheet_id)
    print(f"Found {len(week_tabs)} week tabs")

    # Ask for confirmation
    print(f"\nThis will add headers J-R to ALL {len(week_tabs)} week tabs.")
    print("Headers to be added (if missing):")
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
    total_rows_updated = 0
    requests_this_minute = 0
    minute_start_time = time.time()

    # Process each week tab
    for tab_idx, week_tab in enumerate(week_tabs, 1):
        print(f"\n[{tab_idx}/{len(week_tabs)}] Processing {week_tab}...")

        # Find header rows
        header_rows = find_header_rows(service, spreadsheet_id, week_tab)
        print(f"  Found {len(header_rows)} header rows")

        if not header_rows:
            print(f"  ⚠ No header rows found, skipping")
            continue

        # Add headers to each row
        for row_num in header_rows:
            # Rate limiting: max 55 requests per minute (leaving buffer)
            if requests_this_minute >= 55:
                elapsed = time.time() - minute_start_time
                if elapsed < 60:
                    wait_time = 60 - elapsed + 1
                    print(f"  ⏸ Rate limit approaching, waiting {wait_time:.0f} seconds...")
                    time.sleep(wait_time)
                requests_this_minute = 0
                minute_start_time = time.time()

            # Check which headers are missing
            missing_headers = check_missing_headers(service, spreadsheet_id, week_tab, row_num)
            requests_this_minute += 1

            if not missing_headers:
                print(f"  ✓ Row {row_num}: All headers present")
                continue

            # Rate limit check before write
            if requests_this_minute >= 55:
                elapsed = time.time() - minute_start_time
                if elapsed < 60:
                    wait_time = 60 - elapsed + 1
                    print(f"  ⏸ Rate limit approaching, waiting {wait_time:.0f} seconds...")
                    time.sleep(wait_time)
                requests_this_minute = 0
                minute_start_time = time.time()

            # Add missing headers
            success = add_headers_to_row(service, spreadsheet_id, week_tab, row_num, missing_headers)
            requests_this_minute += 1

            if success:
                missing_cols = ', '.join(missing_headers.keys())
                print(f"  ✓ Row {row_num}: Added {len(missing_headers)} headers ({missing_cols})")
                total_headers_added += len(missing_headers)
                total_rows_updated += 1
            else:
                print(f"  ✗ Row {row_num}: Rate limit hit, waiting 60s...")
                time.sleep(60)
                requests_this_minute = 0
                minute_start_time = time.time()
                # Retry
                success = add_headers_to_row(service, spreadsheet_id, week_tab, row_num, missing_headers)
                requests_this_minute += 1
                if success:
                    missing_cols = ', '.join(missing_headers.keys())
                    print(f"  ✓ Row {row_num}: Added {len(missing_headers)} headers ({missing_cols}) [retry]")
                    total_headers_added += len(missing_headers)
                    total_rows_updated += 1
                else:
                    print(f"  ✗ Row {row_num}: Failed even after retry")

    print("\n" + "=" * 80)
    print(f"COMPLETE")
    print(f"  Updated {total_rows_updated} header rows")
    print(f"  Added {total_headers_added} individual headers")
    print(f"  Processed {len(week_tabs)} week tabs")
    print("=" * 80)


if __name__ == "__main__":
    main()
