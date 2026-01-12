"""Analyze Monday data in Google Sheets Jan-12 tab"""
import asyncio
from app.services.google.sheets_client import GoogleSheetsClient
from app.config import settings

async def analyze():
    client = GoogleSheetsClient()
    service = client._ensure_service()

    # Read entire Jan-12 tab
    range_name = 'Jan-12!A1:R200'
    result = service.spreadsheets().values().get(
        spreadsheetId=settings.sales_spreadsheet_id,
        range=range_name
    ).execute()

    values = result.get('values', [])
    print(f'Total rows in Jan-12 tab: {len(values)}')
    print()

    # According to DAY_HEADER_ROWS, Monday starts at row 1 (index 0)
    # Row structure:
    # Row 1: Monday header
    # Row 2-3: Empty/formatting
    # Row 4: GLEN header
    # Row 5-9: GLEN slots 1-5
    # Row 10-12: Empty/formatting
    # Row 13: GREAT REP header
    # Row 14-18: GREAT REP slots 1-5
    # Row 19-21: Empty/formatting
    # Row 22: ILAN header
    # Row 23-27: ILAN slots 1-5

    print('=== MONDAY SECTION (Rows 1-30) ===')
    print()
    for i in range(min(30, len(values))):
        row = values[i]
        # Show first 10 columns (A-J)
        display_row = row[:10] if len(row) > 10 else row
        # Pad to show structure
        display_row = display_row + [''] * (10 - len(display_row))
        print(f'Row {i+1:3d}: {display_row}')

    print()
    print('=== LOOKING FOR REP NAMES ===')
    # GLEN should be at row 4 (index 3)
    # GREAT REP should be at row 13 (index 12)
    # ILAN should be at row 22 (index 21)

    for rep_name, row_offset in [("GLEN", 4), ("GREAT REP", 13), ("ILAN", 22)]:
        row_index = row_offset - 1  # Convert to 0-indexed
        if row_index < len(values):
            row = values[row_index]
            cell_value = row[0] if len(row) > 0 else ""
            print(f'{rep_name} (Row {row_offset}): "{cell_value}"')
        else:
            print(f'{rep_name} (Row {row_offset}): ROW DOES NOT EXIST')

asyncio.run(analyze())
