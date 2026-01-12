"""
Check all weekly tabs for Paid/Unpaid column consistency.
"""
import sys
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from app.config import settings

# Fix encoding for Windows console
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

def get_all_sheet_tabs():
    """Get list of all tabs in the spreadsheet."""
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
    spreadsheet_id = settings.sales_spreadsheet_id

    # Get spreadsheet metadata to list all sheets
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = spreadsheet.get('sheets', [])

    # Filter for weekly tabs (format: Mon-DD)
    import re
    week_pattern = re.compile(r'^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{2}$')

    week_tabs = []
    for sheet in sheets:
        title = sheet['properties']['title']
        if week_pattern.match(title):
            week_tabs.append(title)

    return sorted(week_tabs), service, spreadsheet_id


def check_tab_for_column_r(week_tab, service, spreadsheet_id):
    """Check if a specific tab has column R (Paid/Unpaid)."""
    try:
        # Get header row (usually row 4 for GLEN)
        range_name = f"'{week_tab}'!A1:R10"
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()

        values = result.get('values', [])

        if not values:
            return None, "No data found"

        # Find header row
        header_row = None
        for i, row in enumerate(values):
            if 'Slot' in row and 'Lead Name' in row:
                header_row = row
                break

        if not header_row:
            return None, "Header row not found"

        # Check if column R exists
        if len(header_row) >= 18:  # Column R is index 17
            col_r_value = header_row[17]
            if col_r_value and 'paid' in col_r_value.lower():
                return True, col_r_value
            else:
                return False, f"Column R exists but has: '{col_r_value}'"
        else:
            return False, "Column R does not exist (row too short)"

    except Exception as e:
        return None, str(e)


def main():
    print("=" * 80)
    print("CHECKING ALL WEEKLY TABS FOR PAID/UNPAID COLUMN")
    print("=" * 80)
    print()

    week_tabs, service, spreadsheet_id = get_all_sheet_tabs()

    print(f"Found {len(week_tabs)} weekly tabs:\n")

    results = []
    for week_tab in week_tabs:
        has_column, detail = check_tab_for_column_r(week_tab, service, spreadsheet_id)
        results.append((week_tab, has_column, detail))

    # Print results table
    print("=" * 80)
    print("RESULTS")
    print("=" * 80)
    print(f"{'Week Tab':<15} {'Status':<10} {'Details'}")
    print("-" * 80)

    all_pass = True
    for week_tab, has_column, detail in results:
        if has_column is True:
            status = "✓ PASS"
        elif has_column is False:
            status = "✗ FAIL"
            all_pass = False
        else:
            status = "? ERROR"
            all_pass = False

        print(f"{week_tab:<15} {status:<10} {detail}")

    print("=" * 80)
    if all_pass:
        print("✓ ALL TABS HAVE PAID/UNPAID COLUMN")
    else:
        print("✗ SOME TABS ARE MISSING PAID/UNPAID COLUMN")
    print("=" * 80)


if __name__ == "__main__":
    main()
