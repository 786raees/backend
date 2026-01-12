import requests
import json

r = requests.get('http://localhost:8000/api/v1/sales/schedule?date=2026-01-12')
data = r.json()

print('=== MONDAY API RESPONSE ===')
print(f'Success: {data["success"]}')
print(f'Day: {data.get("day_name")}')
print(f'Week Tab: {data.get("week_tab")}')
print()

print('=== REPS ===')
for rep in data.get('reps', []):
    rep_name = rep['name']
    appointments = rep['appointments']
    filled = [a for a in appointments if a.get('lead_name')]
    print(f"{rep_name}: {len(filled)} filled / {len(appointments)} total")

    for i, appt in enumerate(appointments[:3], 1):
        lead = appt.get('lead_name', '')
        source = appt.get('lead_source', '')
        print(f"  Slot {i}: '{lead}' (source: '{source}')")

print()
print('=== DAY TOTALS ===')
totals = data.get('day_totals', {})
print(f"Set: {totals.get('total_set', 0)}")
print(f"Attended: {totals.get('total_attended', 0)}")
print(f"Sold: {totals.get('total_sold', 0)}")
