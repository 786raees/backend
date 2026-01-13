"""
Check Payment Status column across all Turf Supply weekly tabs.
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
    """Check Payment Status column in all weekly tabs."""

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
    for sheet in sheets:
        title = sheet['properties']['title']
        if week_pattern.match(title):
            week_tabs.append(title)

    week_tabs = sorted(week_tabs)

    print("=" * 80)
    print("CHECKING PAYMENT STATUS COLUMN - TURF SUPPLY")
    print("=" * 80)
    print(f"\nFound {len(week_tabs)} weekly tabs\n")

    results = []

    for week_tab in week_tabs:
        try:
            # Get header row (row 4)
            result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{week_tab}'!A1:P10"
            ).execute()

            values = result.get('values', [])

            # Find header row
            header_row = None
            for row in values[:10]:
                if 'Slot' in row and 'Variety' in row:
                    header_row = row
                    break

            if not header_row:
                results.append((week_tab, False, "Header row not found"))
                continue

            # Check for Payment Status column (should be column P = index 15)
            if len(header_row) >= 16:
                col_p_value = header_row[15]
                if col_p_value and 'payment' in col_p_value.lower():
                    results.append((week_tab, True, col_p_value))
                else:
                    results.append((week_tab, False, f"Column P has: '{col_p_value}'"))
            else:
                results.append((week_tab, False, "Column P does not exist"))

        except Exception as e:
            results.append((week_tab, None, str(e)))

    # Print results
    print("=" * 80)
    print("RESULTS")
    print("=" * 80)
    print(f"{'Week Tab':<15} {'Status':<10} {'Details'}")
    print("-" * 80)

    pass_count = 0
    fail_count = 0

    for week_tab, has_column, detail in results:
        if has_column is True:
            status = "✓ PASS"
            pass_count += 1
        elif has_column is False:
            status = "✗ FAIL"
            fail_count += 1
        else:
            status = "? ERROR"

        print(f"{week_tab:<15} {status:<10} {detail}")

    print("=" * 80)
    print(f"Pass: {pass_count}/{len(results)} | Fail: {fail_count}/{len(results)}")

    if fail_count == 0:
        print("✓ ALL TABS HAVE PAYMENT STATUS COLUMN")
    else:
        print(f"✗ {fail_count} TABS MISSING PAYMENT STATUS COLUMN")
    print("=" * 80)


if __name__ == "__main__":
    main()
