import requests
import json

# Check Monday's API response
r = requests.get('https://backend-b23y.onrender.com/api/v1/sales/schedule?date=2026-01-12')
print(f'Status Code: {r.status_code}')
print()

data = r.json()
print(f'Date: {data.get("date")}')
print(f'Day: {data.get("day_name")}')
print(f'Week Tab: {data.get("week_tab")}')
print(f'Success: {data.get("success")}')
print()

reps = data.get('reps', [])
print(f'Number of Reps: {len(reps)}')
for rep in reps:
    rep_name = rep.get('rep_name')
    appointments = rep.get('appointments', [])
    filled = [a for a in appointments if a.get('client_name')]
    print(f'  {rep_name}: {len(filled)} filled / {len(appointments)} total slots')

    if filled:
        print(f'    Appointments:')
        for appt in filled[:3]:
            print(f'      - {appt.get("time")}: {appt.get("client_name")}')

print()
print('Day Totals:')
totals = data.get('day_totals', {})
print(f'  Attended: {totals.get("attended", 0)}')
print(f'  Sold: {totals.get("sold", 0)}')
print(f'  Total Sales: ${totals.get("total_sales", 0):,.0f}')
