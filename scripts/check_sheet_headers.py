"""
Check if columns N, O, P, Q have proper headers in the Sales sheet.
"""
import sys
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from app.config import settings

# Fix encoding for Windows console
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

def check_headers():
    """Check the header row for week tabs."""

    # Initialize Google Sheets client
    # Get the plain dict from SecretStr
    service_account_info = settings.google_service_account_json
    if hasattr(service_account_info, 'get_secret_value'):
        service_account_info = service_account_info.get_secret_value()

    # If it's a string, parse it as JSON
    if isinstance(service_account_info, str):
        service_account_info = json.loads(service_account_info)

    credentials = service_account.Credentials.from_service_account_info(
        service_account_info,
        scopes=['https://www.googleapis.com/auth/spreadsheets']
    )
    service = build('sheets', 'v4', credentials=credentials)
    spreadsheet_id = settings.sales_spreadsheet_id

    # Get the first week tab (Jan-05)
    week_tab = "Jan-05"

    print("=" * 80)
    print(f"CHECKING HEADERS IN SHEET: {week_tab}")
    print("=" * 80)

    # Get header rows (first 5 rows to see the structure)
    range_name = f"'{week_tab}'!A1:Q5"
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()

    values = result.get('values', [])

    if not values:
        print("No data found in the sheet!")
        return

    # Find the header row (look for "Slot" or row with column names)
    print("\nFirst 5 rows of the sheet:\n")
    for i, row in enumerate(values, 1):
        # Pad row to show all columns A-Q
        padded_row = row + [''] * (17 - len(row))
        print(f"Row {i}:")
        for j, cell in enumerate(padded_row):
            col_letter = chr(65 + j)  # A, B, C, ...
            print(f"  {col_letter}: {cell if cell else '(empty)'}")
        print()

    # Check if the required columns exist
    print("\n" + "=" * 80)
    print("COLUMN MAPPING CHECK")
    print("=" * 80)

    expected_columns = {
        'A': 'Slot',
        'B': 'Lead Name',
        'C': 'Lead Source',
        'D': 'Appointment Set',
        'E': 'Appointment Confirmed',
        'F': 'Appointment Attended',
        'G': 'Job Sold',
        'H': 'Reason',
        'I': 'Conversion',
        'J': 'Sell Price Ex GST',
        'K': 'Appointment Time',
        'L': 'Project Type',
        'M': 'Suburb',
        'N': 'Region',
        'O': 'Appointment Set Who',
        'P': 'Appointment Confirmed By',
        'Q': 'Gross Profit Margin %'
    }

    # Try to find the header row
    header_row = None
    header_row_idx = None
    for i, row in enumerate(values):
        if 'Slot' in row or 'Lead Name' in row:
            header_row = row
            header_row_idx = i + 1
            break

    if header_row:
        print(f"\nFound header row at Row {header_row_idx}")
        print("\nColumn Status:")
        for col_letter, expected_name in expected_columns.items():
            col_idx = ord(col_letter) - 65
            actual_name = header_row[col_idx] if col_idx < len(header_row) else ''

            if actual_name:
                status = "✓ EXISTS"
                match = "(matches)" if expected_name.lower() in actual_name.lower() else f"(has: '{actual_name}')"
            else:
                status = "✗ MISSING"
                match = ""

            print(f"  {col_letter} - {expected_name:30} : {status} {match}")
    else:
        print("\n✗ Could not find header row with 'Slot' or 'Lead Name'")

    print("\n" + "=" * 80)


if __name__ == "__main__":
    check_headers()
