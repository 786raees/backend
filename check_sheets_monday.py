import asyncio
from app.services.google_sheets_service import GoogleSheetsService

async def check_monday():
    gs = GoogleSheetsService()

    # Read the Jan-12 week tab
    data = await gs.read_week_tab('Jan-12')

    print(f'Total rows in Jan-12 tab: {len(data)}')
    print()

    # Find Monday section
    print('Looking for Monday data...')
    for i, row in enumerate(data):
        if len(row) > 0 and 'Monday' in str(row[0]):
            print(f'Found Monday header at row {i+1}: {row[0]}')
            print()
            print('Next 20 rows after Monday header:')
            for j in range(i+1, min(i+21, len(data))):
                print(f'  Row {j+1}: {data[j][:8] if len(data[j]) > 8 else data[j]}')
            break

    print()
    print('Checking for rep names (GLEN, GREAT REP, ILAN)...')
    for i, row in enumerate(data[:50]):
        if len(row) > 0:
            cell = str(row[0])
            if 'GLEN' in cell or 'GREAT' in cell or 'ILAN' in cell:
                print(f'  Row {i+1}: {cell}')

asyncio.run(check_monday())
