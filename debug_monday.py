"""Debug Monday data parsing"""
import asyncio
from datetime import date
from app.services.google.sheets_client import GoogleSheetsClient
from app.config import settings

async def debug_monday():
    client = GoogleSheetsClient()
    service = client._ensure_service()

    # Constants from SalesService
    DAY_HEADER_ROWS = {"Monday": 1, "Tuesday": 31}
    REP_SLOT_OFFSETS = {"GLEN": 4, "GREAT REP": 13, "ILAN": 22}
    SALES_REPS = ["GLEN", "GREAT REP", "ILAN"]

    week_tab = "Jan-12"
    day_name = "Monday"

    # Calculate range (same as in get_daily_schedule)
    day_start_row = DAY_HEADER_ROWS[day_name]
    day_end_row = day_start_row + 28

    range_notation = f"'{week_tab}'!A{day_start_row}:R{day_end_row}"

    print(f"=== DEBUG MONDAY PARSING ===")
    print(f"Week Tab: {week_tab}")
    print(f"Day: {day_name}")
    print(f"Day Start Row: {day_start_row}")
    print(f"Day End Row: {day_end_row}")
    print(f"Range: {range_notation}")
    print()

    # Fetch data
    result = service.spreadsheets().values().get(
        spreadsheetId=settings.sales_spreadsheet_id,
        range=range_notation
    ).execute()

    all_rows = result.get('values', [])
    print(f"Total rows fetched: {len(all_rows)}")
    print()

    # Show first 15 rows
    print("=== FIRST 15 ROWS (0-indexed in array, showing absolute row number) ===")
    for i in range(min(15, len(all_rows))):
        absolute_row = day_start_row + i
        row_data = all_rows[i]
        first_5_cells = row_data[:5] if len(row_data) >= 5 else row_data
        print(f"Index {i:2d} (Row {absolute_row:3d}): {first_5_cells}")
    print()

    # Now show row calculation for each rep
    print("=== ROW CALCULATIONS ===")
    for rep in SALES_REPS:
        print(f"\n{rep}:")
        rep_offset = REP_SLOT_OFFSETS[rep]

        for slot in range(1, 6):  # Slots 1-5
            # This is the calculation from calculate_row_number()
            row_number = day_start_row + rep_offset + (slot - 1)

            # Convert to 0-indexed relative to our fetched range
            relative_row = row_number - day_start_row

            print(f"  Slot {slot}:")
            print(f"    Formula: {day_start_row} (day_start) + {rep_offset} (rep_offset) + {slot-1} (slot-1)")
            print(f"    Absolute Row Number: {row_number}")
            print(f"    Relative Index: {relative_row}")

            if 0 <= relative_row < len(all_rows):
                row_data = all_rows[relative_row]
                lead_name = row_data[1] if len(row_data) > 1 else ""
                lead_source = row_data[2] if len(row_data) > 2 else ""
                print(f"    Data: {row_data[:5] if len(row_data) >= 5 else row_data}")
                print(f"    Lead Name (col B): '{lead_name}'")
                print(f"    Lead Source (col C): '{lead_source}'")
            else:
                print(f"    ERROR: Index {relative_row} out of range (only {len(all_rows)} rows)")

asyncio.run(debug_monday())
