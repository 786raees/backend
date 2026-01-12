"""
Add missing column headers (J-R) efficiently with batched operations.
"""
import sys
import json
import re
import time
from google.oauth2 import service_account
from googleapiclient.discovery import build
from app.config import settings

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

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


def process_week_tab(service, spreadsheet_id, week_tab):
    """Process one week tab - find header rows and add missing headers."""

    # Read entire tab up to column R and row 150
    range_name = f"'{week_tab}'!A1:R150"
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()

    values = result.get('values', [])

    # Find header rows (rows where column A = "Slot")
    header_rows = []
    for i, row in enumerate(values, 1):
        if len(row) > 0 and row[0] == 'Slot':
            header_rows.append((i, row))

    if not header_rows:
        return 0, 0

    # Check which headers are missing and build updates
    updates = []
    rows_to_update = 0

    for row_num, row_values in header_rows:
        # Pad row to check all columns A-R
        padded_row = row_values + [''] * (18 - len(row_values))

        missing = []
        for col_letter, header_text in HEADERS.items():
            col_idx = ord(col_letter) - 65  # J=9, K=10, etc.
            if col_idx < len(padded_row) and not padded_row[col_idx]:
                missing.append(col_letter)
                updates.append({
                    'range': f"'{week_tab}'!{col_letter}{row_num}",
                    'values': [[header_text]]
                })

        if missing:
            rows_to_update += 1

    # Perform batch update if there are missing headers
    if updates:
        body = {
            'valueInputOption': 'RAW',
            'data': updates
        }
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()

    return rows_to_update, len(updates)


def main():
    print("=" * 80)
    print("ADDING MISSING HEADERS (J-R) TO ALL SALES SHEETS")
    print("=" * 80)

    service = get_service()
    spreadsheet_id = settings.sales_spreadsheet_id

    print("\nFetching week tabs...")
    week_tabs = get_week_tabs(service, spreadsheet_id)
    print(f"Found {len(week_tabs)} week tabs\n")

    total_rows_updated = 0
    total_headers_added = 0
    writes_this_minute = 0
    minute_start = time.time()

    for idx, week_tab in enumerate(week_tabs, 1):
        print(f"[{idx}/{len(week_tabs)}] {week_tab}...", end=' ', flush=True)

        try:
            # Check rate limit (55 writes per minute with buffer)
            if writes_this_minute >= 50:
                elapsed = time.time() - minute_start
                if elapsed < 60:
                    wait = 61 - elapsed
                    print(f"⏸ waiting {wait:.0f}s...", end=' ', flush=True)
                    time.sleep(wait)
                writes_this_minute = 0
                minute_start = time.time()

            rows_updated, headers_added = process_week_tab(service, spreadsheet_id, week_tab)
            writes_this_minute += 1  # One write per tab (or zero if no updates)

            if rows_updated > 0:
                print(f"✓ Updated {rows_updated} rows, added {headers_added} headers")
                total_rows_updated += rows_updated
                total_headers_added += headers_added
            else:
                print("✓ Already complete")

        except Exception as e:
            if '429' in str(e):
                print("✗ Rate limit, waiting 60s...")
                time.sleep(60)
                writes_this_minute = 0
                minute_start = time.time()
                # Retry
                try:
                    rows_updated, headers_added = process_week_tab(service, spreadsheet_id, week_tab)
                    writes_this_minute += 1
                    if rows_updated > 0:
                        print(f"  ✓ [Retry] Updated {rows_updated} rows, added {headers_added} headers")
                        total_rows_updated += rows_updated
                        total_headers_added += headers_added
                except Exception as e2:
                    print(f"  ✗ [Retry failed] {e2}")
            else:
                print(f"✗ Error: {e}")

    print("\n" + "=" * 80)
    print("COMPLETE")
    print(f"  Rows updated: {total_rows_updated}")
    print(f"  Headers added: {total_headers_added}")
    print(f"  Tabs processed: {len(week_tabs)}")
    print("=" * 80)


if __name__ == "__main__":
    main()
