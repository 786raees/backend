"""
Check Jan-19 sheet structure to identify missing columns.
"""
import sys
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from app.config import settings

# Fix encoding for Windows console
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

def check_sheet_structure(week_tab="Jan-19"):
    """Check the structure of a specific week tab."""

    # Initialize Google Sheets client
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

    print("=" * 80)
    print(f"CHECKING SHEET: {week_tab}")
    print("=" * 80)

    # Get first 10 rows to see structure
    range_name = f"'{week_tab}'!A1:R10"
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()

    values = result.get('values', [])

    if not values:
        print("No data found!")
        return

    print("\nFirst 10 rows:\n")
    for i, row in enumerate(values, 1):
        # Pad row to show all columns A-R
        padded_row = row + [''] * (18 - len(row))
        print(f"Row {i}:")
        for j, cell in enumerate(padded_row[:18]):  # A-R
            col_letter = chr(65 + j)
            if cell:
                print(f"  {col_letter}: {cell}")
        print()

    # Find header row
    header_row = None
    header_row_idx = None
    for i, row in enumerate(values):
        if 'Slot' in row or 'Lead Name' in row:
            header_row = row
            header_row_idx = i + 1
            break

    if header_row:
        print("=" * 80)
        print(f"HEADER ROW FOUND AT ROW {header_row_idx}")
        print("=" * 80)
        print("\nColumn Mapping:")
        for j, header in enumerate(header_row):
            if header:
                col_letter = chr(65 + j)
                print(f"  {col_letter}: {header}")

        # Check for missing columns
        print("\n" + "=" * 80)
        print("EXPECTED COLUMNS CHECK")
        print("=" * 80)

        expected = {
            'A': 'Slot',
            'B': 'Lead Name',
            'C': 'Lead Source',
            'D': 'Appointment Set',
            'E': 'Appointment Confirmed',
            'F': 'Appointment Attended',
            'G': 'Job Sold',
            'H': 'Sold / Not Sold Reason',
            'I': 'Conversion',
            'J': 'Sell Price Ex GST',
            'K': 'Appointment Time',
            'L': 'Project Type',
            'M': 'Suburb',
            'N': 'Region',
            'O': 'Appointment Set Who',
            'P': 'Appointment Confirmed By',
            'Q': 'Gross Profit Margin %',
            'R': 'Paid/Unpaid'  # Check if this exists
        }

        for col_letter, expected_name in expected.items():
            col_idx = ord(col_letter) - 65
            actual_name = header_row[col_idx] if col_idx < len(header_row) else ''

            if actual_name:
                status = "✓ EXISTS"
                if expected_name.lower() in actual_name.lower() or actual_name.lower() in expected_name.lower():
                    match = "(matches)"
                else:
                    match = f"(has: '{actual_name}')"
            else:
                status = "✗ MISSING"
                match = f"(should be: '{expected_name}')"

            print(f"  {col_letter} - {expected_name:30} : {status} {match}")


if __name__ == "__main__":
    # Check Jan-19 first
    check_sheet_structure("Jan-19")
