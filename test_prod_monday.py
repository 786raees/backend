import requests
import json

print("=== TESTING PRODUCTION API FOR MONDAY ===")

url = "https://backend-b23y.onrender.com/api/v1/sales/schedule?date=2026-01-12"
print(f"URL: {url}")
print()

try:
    r = requests.get(url, timeout=30)
    print(f"Status Code: {r.status_code}")
    print()

    if r.status_code == 200:
        data = r.json()
        print(f"Success: {data.get('success')}")
        print(f"Day: {data.get('day_name')}")
        print(f"Week Tab: {data.get('week_tab')}")
        print()

        if data.get('success'):
            reps = data.get('reps', [])
            print(f"Reps count: {len(reps)}")

            for rep in reps:
                rep_name = rep['name']
                appointments = rep['appointments']
                filled = [a for a in appointments if a.get('lead_name')]
                print(f"  {rep_name}: {len(filled)} filled / {len(appointments)} total")

                for i, appt in enumerate(filled[:2], 1):
                    print(f"    - {appt.get('lead_name')} ({appt.get('lead_source', 'N/A')})")
        else:
            print(f"Error: {data.get('error')}")
    else:
        print(f"HTTP Error: {r.text[:200]}")

except Exception as e:
    print(f"Request failed: {e}")
