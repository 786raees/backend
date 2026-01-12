"""Debug script to inspect raw week data from Google Sheets."""
import sys
import os

# Add backend to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from app.services.google.sheets_client import GoogleSheetsClient

def main():
    client = GoogleSheetsClient()

    # Fetch Jan-27 week data
    week_tab = "Jan-27"
    print(f"\n{'='*60}")
    print(f"Fetching data from tab: {week_tab}")
    print(f"{'='*60}\n")

    rows = client.get_worksheet_data(week_tab, range_cols="A:J")

    current_truck = 1
    truck_1_laying_fees = 0.0
    truck_2_laying_fees = 0.0
    truck_1_laying_costs = 0.0
    truck_2_laying_costs = 0.0

    print(f"Total rows: {len(rows)}\n")

    for i, row in enumerate(rows):
        # Check for truck markers
        first_cell = str(row[0]).strip().upper() if row else ""

        if "TRUCK 1" in first_cell:
            current_truck = 1
            print(f"Row {i+1}: >>> TRUCK 1 SECTION <<<")
            continue
        elif "TRUCK 2" in first_cell:
            current_truck = 2
            print(f"Row {i+1}: >>> TRUCK 2 SECTION <<<")
            continue

        # Skip if not enough columns or no variety
        if len(row) < 5:
            continue

        variety = row[1] if len(row) > 1 else ""
        if not variety or variety.strip() == "" or variety == "Variety":
            continue

        # Get values
        service_type = str(row[3]).strip().upper() if len(row) > 3 else ""
        sqm = float(row[4]) if len(row) > 4 and row[4] else 0.0
        laying_fee = float(row[9]) if len(row) > 9 and row[9] else 0.0

        # Calculate laying cost (only for SL)
        laying_cost = sqm * 2.20 if service_type == "SL" else 0.0

        # Track by truck
        if current_truck == 1:
            truck_1_laying_fees += laying_fee
            truck_1_laying_costs += laying_cost
        else:
            truck_2_laying_fees += laying_fee
            truck_2_laying_costs += laying_cost

        print(f"Row {i+1} [Truck {current_truck}]: {variety[:15]:15} | {service_type:3} | SQM: {sqm:6.0f} | Laying Fee: ${laying_fee:7.0f} | Laying Cost: ${laying_cost:7.2f}")

    print(f"\n{'='*60}")
    print("SUMMARY")
    print(f"{'='*60}")
    print(f"Truck 1 Laying Fees:  ${truck_1_laying_fees:,.2f}")
    print(f"Truck 1 Laying Costs: ${truck_1_laying_costs:,.2f}")
    print(f"Truck 2 Laying Fees:  ${truck_2_laying_fees:,.2f}")
    print(f"Truck 2 Laying Costs: ${truck_2_laying_costs:,.2f}")
    print(f"\nTotal Laying Fees:  ${truck_1_laying_fees + truck_2_laying_fees:,.2f}")
    print(f"Total Laying Costs: ${truck_1_laying_costs + truck_2_laying_costs:,.2f}")

if __name__ == "__main__":
    main()
