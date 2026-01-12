import requests
import json

print("=== TESTING JAN-05 WEEK (Monday = Jan 5, 2026) ===")

url = "https://backend-b23y.onrender.com/api/v1/sales/schedule?date=2026-01-05"
print(f"URL: {url}")
print()

try:
    r = requests.get(url, timeout=30)
    print(f"Status Code: {r.status_code}")

    if r.status_code == 200:
        data = r.json()
        print(f"Success: {data.get('success')}")
        print(f"Day: {data.get('day_name')}")
        print(f"Week Tab: {data.get('week_tab')}")
        print()

        if data.get('success'):
            reps = data.get('reps', [])
            total_appointments = sum(len([a for a in rep['appointments'] if a.get('lead_name')]) for rep in reps)
            print(f"Total appointments across all reps: {total_appointments}")

            for rep in reps:
                rep_name = rep['name']
                appointments = rep['appointments']
                filled = [a for a in appointments if a.get('lead_name')]
                print(f"  {rep_name}: {len(filled)} filled")
        else:
            print(f"Error: {data.get('error')}")
    else:
        print(f"HTTP Error {r.status_code}")

except Exception as e:
    print(f"Request failed: {e}")
