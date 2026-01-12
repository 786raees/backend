"""
Inspect Sales Sheet structure to verify all columns A-Q are present.
"""
import asyncio
import sys
from datetime import date
from app.services.sales_service import sales_service

# Fix encoding for Windows console
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')


async def inspect_sheet():
    """Check the sheet structure and sample data."""

    print("=" * 80)
    print("SALES SHEET INSPECTION")
    print("=" * 80)

    # Get available weeks
    weeks = await sales_service.get_available_weeks()
    print(f"\n✓ Available weeks: {len(weeks)} weeks")
    print(f"  First 5: {weeks[:5]}")
    print(f"  Last 5: {weeks[-5:]}")

    # Get a sample schedule
    target_date = date(2026, 1, 6)  # Monday Jan 6
    print(f"\n\nFetching sample data for {target_date}...")
    result = await sales_service.get_daily_schedule(target_date)

    if not result.get('success'):
        print(f"✗ Error: {result.get('error')}")
        return

    print(f"✓ Success!")
    print(f"  Date: {result['date']}")
    print(f"  Day: {result['day_name']}")
    print(f"  Reps: {len(result['reps'])}")

    # Inspect first appointment from first rep
    if result['reps'] and result['reps'][0]['appointments']:
        first_rep = result['reps'][0]
        first_appt = first_rep['appointments'][0]
        rep_name = first_rep.get('rep', first_rep.get('name', 'Unknown'))

        print(f"\n\nSAMPLE APPOINTMENT DATA (Rep: {rep_name}):")
        print("-" * 80)

        # Core fields
        print("\n📋 CORE FIELDS:")
        print(f"  Slot: {first_appt.get('slot')}")
        print(f"  Row Number: {first_appt.get('row_number')}")
        print(f"  Lead Name: {first_appt.get('lead_name', 'N/A')}")

        # Contact/Location fields
        print("\n📍 LOCATION FIELDS:")
        print(f"  Suburb (M): {first_appt.get('suburb', 'N/A')}")
        print(f"  Region (N): {first_appt.get('region', 'N/A')}")

        # Lead fields
        print("\n📞 LEAD FIELDS:")
        print(f"  Lead Source (C): {first_appt.get('lead_source', 'N/A')}")

        # Appointment status fields
        print("\n📅 APPOINTMENT STATUS:")
        print(f"  Appointment Set (D): {first_appt.get('appointment_set', False)}")
        print(f"  Set By WHO (O): {first_appt.get('appointment_set_who', 'N/A')}")
        print(f"  Appointment Confirmed (E): {first_appt.get('appointment_confirmed', False)}")
        print(f"  Confirmed BY (P): {first_appt.get('appointment_confirmed_by', 'N/A')}")
        print(f"  Attended (F): {first_appt.get('appointment_attended', False)}")

        # Sales fields
        print("\n💰 SALES FIELDS:")
        print(f"  Job Sold (G): {first_appt.get('job_sold', False)}")
        print(f"  Reason (H): {first_appt.get('reason', 'N/A')}")
        print(f"  Sell Price (J): ${first_appt.get('sell_price', 0)}")
        print(f"  Gross Margin % (Q): {first_appt.get('gross_profit_margin_pct', 'N/A')}")

        # Other fields
        print("\n🔧 OTHER FIELDS:")
        print(f"  Appointment Time (K): {first_appt.get('appointment_time', 'N/A')}")
        print(f"  Project Type (L): {first_appt.get('project_type', 'N/A')}")

        # Check for empty new fields
        print("\n\n🔍 NEW FIELDS STATUS CHECK:")
        new_fields = {
            "Region (N)": first_appt.get('region'),
            "Appointment Set Who (O)": first_appt.get('appointment_set_who'),
            "Appointment Confirmed By (P)": first_appt.get('appointment_confirmed_by'),
            "Gross Profit Margin % (Q)": first_appt.get('gross_profit_margin_pct')
        }

        for field_name, value in new_fields.items():
            status = "✓ Has data" if value else "⚠ Empty"
            print(f"  {field_name}: {status} (value: {value})")

    # Check totals
    print("\n\n📊 TOTALS:")
    totals = result.get('totals', {})
    print(f"  Total Set: {totals.get('total_set', 0)}")
    print(f"  Total Confirmed: {totals.get('total_confirmed', 0)}")
    print(f"  Total Attended: {totals.get('total_attended', 0)}")
    print(f"  Total Sold: {totals.get('total_sold', 0)}")
    print(f"  Total Revenue: ${totals.get('total_revenue', 0)}")

    print("\n" + "=" * 80)
    print("INSPECTION COMPLETE")
    print("=" * 80)


if __name__ == "__main__":
    asyncio.run(inspect_sheet())
