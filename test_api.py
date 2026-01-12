import requests

# Get available weeks
weeks_resp = requests.get('http://localhost:8000/api/v1/sales/weeks')
weeks = weeks_resp.json().get('weeks', [])

if not weeks:
    print('❌ No weeks available')
    exit(1)

week = weeks[0]
print(f'Testing with week: {week}')
print()

# Test Sales Stats endpoint
print('=== Sales Manager Stats ===')
r = requests.get(f'http://localhost:8000/api/v1/sales/stats?week={week}')
print(f'Status: {r.status_code}')
data = r.json()
print(f'Success: {data.get("success")}')
print(f'Total Attended: {data.get("attended_count", 0)}')
print(f'Total Sold: {data.get("sold_count", 0)}')
print(f'Total Sales: ${data.get("total_sales_ex_gst", 0):,.0f}')
print(f'Conversion Rate: {data.get("conversion_rate", 0):.1f}%')
print()

# Test CEO Summary endpoint
print('=== CEO Summary ===')
r = requests.get(f'http://localhost:8000/api/v1/sales/ceo-summary?week={week}')
print(f'Status: {r.status_code}')
data = r.json()
print(f'Success: {data.get("success")}')
print(f'Week: {data.get("week_start")}')
print(f'Appointments: {data.get("appointments_count", 0)}')
print(f'Attended: {data.get("attended_count", 0)}')
print(f'Sold: {data.get("sold_count", 0)}')
print(f'Total Revenue: ${data.get("total_revenue", 0):,.0f}')
print()

# Test appointment update endpoint
print('=== Test Appointment Update (Dry Run) ===')
test_payload = {
    "week_tab": week,
    "day": "Monday",
    "rep": "GLEN",
    "slot": 1,
    "column": "F",
    "value": "Yes"
}
r = requests.post('http://localhost:8000/api/v1/sales/appointment/update', json=test_payload)
print(f'Status: {r.status_code}')
result = r.json()
print(f'Success: {result.get("success")}')
print(f'Message: {result.get("message", "N/A")}')
