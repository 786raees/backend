import requests

# Test today's schedule
r = requests.get('http://localhost:8000/api/v1/sales/schedule?date=2026-01-13')
data = r.json()

print('=== TODAY\'S SCHEDULE (Jan 13, 2026) ===')
print(f'Status: {r.status_code}')
print(f'Day: {data.get("day_name")}')
print(f'Week Tab: {data.get("week_tab")}')
print()

reps = data.get('reps', [])
for rep in reps:
    slots = rep.get('appointments', [])
    filled = [s for s in slots if s.get('client_name')]
    print(f'  {rep.get("rep_name")}: {len(filled)}/{len(slots)} appointments')
    for appt in filled[:3]:  # Show first 3
        print(f'    - {appt.get("time")}: {appt.get("client_name")} ({appt.get("project_type", "N/A")})')

print()
totals = data.get('day_totals', {})
print(f'Total Attended: {totals.get("attended", 0)}')
print(f'Total Sold: {totals.get("sold", 0)}')
print(f'Total Sales: ${totals.get("total_sales", 0):,.0f}')
print(f'Conversion Rate: {totals.get("conversion_rate", 0):.1f}%')
