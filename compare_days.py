"""Compare Monday vs Tuesday structure"""
import asyncio
from app.services.google.sheets_client import GoogleSheetsClient
from app.config import settings

async def compare():
    client = GoogleSheetsClient()
    service = client._ensure_service()

    # Read Jan-12 tab
    range_name = 'Jan-12!A1:R200'
    result = service.spreadsheets().values().get(
        spreadsheetId=settings.sales_spreadsheet_id,
        range=range_name
    ).execute()

    values = result.get('values', [])

    # Monday section: Row 1-30
    # Tuesday section: Row 31-60 (DAY_HEADER_ROWS says Tuesday starts at row 31)

    print('=== MONDAY (Rows 1-10) ===')
    for i in range(0, 10):
        row = values[i] if i < len(values) else []
        print(f'Row {i+1:3d}: {row[:5] if len(row) > 5 else row}')

    print()
    print('=== TUESDAY (Rows 31-40) ===')
    # Tuesday starts at row 31 (index 30)
    for i in range(30, 40):
        row = values[i] if i < len(values) else []
        print(f'Row {i+1:3d}: {row[:5] if len(row) > 5 else row}')

    print()
    print('=== COMPARISON ===')
    print('Monday Row 3:', values[2][:2] if len(values) > 2 else 'N/A')
    print('Monday Row 4:', values[3][:2] if len(values) > 3 else 'N/A')
    print()
    print('Tuesday Row 33:', values[32][:2] if len(values) > 32 else 'N/A')
    print('Tuesday Row 34:', values[33][:2] if len(values) > 33 else 'N/A')

asyncio.run(compare())
