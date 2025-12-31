"""
Comprehensive Test for Milestone 2 - Turf Supply Dashboard

Tests:
1. API Response Structure
2. Laying Cost Calculations ($2.20/SQM for SL only)
3. Weekly Tab Detection
4. January Data Appearing Correctly
5. Cache Refresh Functionality
6. Frontend Configuration
"""
import os
import sys
import json
import requests
from datetime import date, datetime

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

API_BASE = "http://localhost:8000"
LAYING_COST_PER_SQM = 2.20

class TestResult:
    def __init__(self, name):
        self.name = name
        self.passed = True
        self.messages = []

    def fail(self, msg):
        self.passed = False
        self.messages.append(f"FAIL: {msg}")

    def warn(self, msg):
        self.messages.append(f"WARN: {msg}")

    def info(self, msg):
        self.messages.append(f"INFO: {msg}")

    def __str__(self):
        status = "PASS" if self.passed else "FAIL"
        result = f"[{status}] {self.name}"
        for msg in self.messages:
            result += f"\n       {msg}"
        return result


def test_api_health():
    """Test 1: API Health Check"""
    result = TestResult("API Health Check")
    try:
        resp = requests.get(f"{API_BASE}/health", timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            result.info(f"Status: {data.get('status')}")
            result.info(f"Google Sheets: {data.get('google_sheets_configured')}")
        else:
            result.fail(f"Health endpoint returned {resp.status_code}")
    except Exception as e:
        result.fail(f"Could not connect to API: {e}")
    return result


def test_schedule_endpoint():
    """Test 2: Schedule Endpoint Returns Valid Data"""
    result = TestResult("Schedule Endpoint")
    try:
        resp = requests.get(f"{API_BASE}/api/v1/schedule", timeout=30)
        if resp.status_code != 200:
            result.fail(f"Schedule endpoint returned {resp.status_code}")
            return result

        data = resp.json()

        # Check required fields
        if not data.get("success"):
            result.fail("Response success is not True")

        if data.get("source") != "google_sheets":
            result.warn(f"Source is '{data.get('source')}', expected 'google_sheets'")
        else:
            result.info(f"Source: {data.get('source')}")

        days = data.get("days", [])
        if len(days) != 10:
            result.fail(f"Expected 10 days, got {len(days)}")
        else:
            result.info(f"Days returned: {len(days)}")

        # Check first day structure
        if days:
            day = days[0]
            required_fields = ["date", "day_name", "is_week_two", "truck1", "truck2"]
            for field in required_fields:
                if field not in day:
                    result.fail(f"Missing field '{field}' in day data")

    except Exception as e:
        result.fail(f"Error: {e}")
    return result


def test_laying_cost_calculation():
    """Test 3: Laying Cost Calculations"""
    result = TestResult("Laying Cost Calculation ($2.20/SQM for SL)")
    try:
        resp = requests.get(f"{API_BASE}/api/v1/schedule", timeout=30)
        data = resp.json()

        sl_count = 0
        sd_p_count = 0
        errors = []

        for day in data.get("days", []):
            for truck_name in ["truck1", "truck2"]:
                truck = day.get(truck_name, {})
                calculated_total = 0

                for delivery in truck.get("deliveries", []):
                    sqm = delivery.get("sqm", 0)
                    service_type = delivery.get("service_type", "")
                    laying_cost = delivery.get("laying_cost", 0)

                    if service_type == "SL":
                        sl_count += 1
                        expected_cost = round(sqm * LAYING_COST_PER_SQM, 2)
                        if abs(laying_cost - expected_cost) > 0.01:
                            errors.append(f"SL {sqm} SQM: expected ${expected_cost}, got ${laying_cost}")
                        calculated_total += laying_cost
                    else:
                        sd_p_count += 1
                        if laying_cost != 0:
                            errors.append(f"{service_type} should have $0 laying_cost, got ${laying_cost}")

                # Check total
                reported_total = truck.get("laying_cost_total", 0)
                if abs(calculated_total - reported_total) > 0.01:
                    errors.append(f"{day['day_name']} {truck_name}: total mismatch (calc={calculated_total}, reported={reported_total})")

        result.info(f"SL deliveries found: {sl_count}")
        result.info(f"SD/P deliveries found: {sd_p_count}")

        if errors:
            for err in errors[:5]:  # Show first 5 errors
                result.fail(err)
        else:
            result.info("All laying costs calculated correctly")

    except Exception as e:
        result.fail(f"Error: {e}")
    return result


def test_january_data():
    """Test 4: January Data Appearing"""
    result = TestResult("January Data Detection")
    try:
        resp = requests.get(f"{API_BASE}/api/v1/schedule", timeout=30)
        data = resp.json()

        january_days = []
        december_days = []

        for day in data.get("days", []):
            day_name = day.get("day_name", "")
            date_str = day.get("date", "")

            if "Jan" in day_name:
                january_days.append(day_name)
            elif "Dec" in day_name:
                december_days.append(day_name)

        result.info(f"December days: {len(december_days)} - {december_days}")
        result.info(f"January days: {len(january_days)} - {january_days}")

        if not january_days:
            result.warn("No January days found - may be expected depending on current date")
        else:
            result.info("January data is appearing correctly")

    except Exception as e:
        result.fail(f"Error: {e}")
    return result


def test_truck_data_structure():
    """Test 5: Truck Data Structure"""
    result = TestResult("Truck Data Structure")
    try:
        resp = requests.get(f"{API_BASE}/api/v1/schedule", timeout=30)
        data = resp.json()

        required_truck_fields = ["deliveries", "sqm_total", "pallet_total", "laying_cost_total", "capacity", "available_sqm"]
        required_delivery_fields = ["sqm", "pallets", "variety", "suburb", "service_type", "laying_cost"]

        day = data.get("days", [{}])[0]

        for truck_name in ["truck1", "truck2"]:
            truck = day.get(truck_name, {})

            for field in required_truck_fields:
                if field not in truck:
                    result.fail(f"{truck_name} missing field '{field}'")

            # Check capacity
            expected_cap = 500 if truck_name == "truck1" else 600
            if truck.get("capacity") != expected_cap:
                result.fail(f"{truck_name} capacity should be {expected_cap}, got {truck.get('capacity')}")

            # Check delivery structure
            deliveries = truck.get("deliveries", [])
            if deliveries:
                delivery = deliveries[0]
                for field in required_delivery_fields:
                    if field not in delivery:
                        result.fail(f"Delivery missing field '{field}'")

        result.info("Truck 1 capacity: 500 SQM")
        result.info("Truck 2 capacity: 600 SQM")
        result.info("All required fields present")

    except Exception as e:
        result.fail(f"Error: {e}")
    return result


def test_cache_refresh():
    """Test 6: Cache Refresh Functionality"""
    result = TestResult("Cache Refresh")
    try:
        # Test refresh parameter
        resp = requests.get(f"{API_BASE}/api/v1/schedule?refresh=true", timeout=30)
        if resp.status_code == 200:
            result.info("GET /schedule?refresh=true works")
        else:
            result.fail(f"Refresh parameter returned {resp.status_code}")

        # Test POST refresh endpoint
        resp = requests.post(f"{API_BASE}/api/v1/schedule/refresh", timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            result.info(f"POST /schedule/refresh: {data.get('message', 'OK')}")
        else:
            result.fail(f"POST refresh returned {resp.status_code}")

    except Exception as e:
        result.fail(f"Error: {e}")
    return result


def test_frontend_config():
    """Test 7: Frontend Configuration"""
    result = TestResult("Frontend Configuration (30-second refresh)")
    try:
        frontend_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
            "frontend", "js", "dashboard.js"
        )

        with open(frontend_path, 'r') as f:
            content = f.read()

        # Check refresh interval
        if "REFRESH_INTERVAL: 30000" in content:
            result.info("Refresh interval: 30 seconds (30000ms)")
        elif "REFRESH_INTERVAL: 300000" in content:
            result.fail("Refresh interval still at 5 minutes (300000ms)")
        else:
            result.warn("Could not determine refresh interval")

        # Check API URL
        if "localhost:8000" in content or "API_BASE_URL" in content:
            result.info("API configuration found in dashboard.js")

    except Exception as e:
        result.fail(f"Error reading frontend config: {e}")
    return result


def test_service_types():
    """Test 8: Service Type Distribution"""
    result = TestResult("Service Type Distribution")
    try:
        resp = requests.get(f"{API_BASE}/api/v1/schedule", timeout=30)
        data = resp.json()

        counts = {"SL": 0, "SD": 0, "P": 0}

        for day in data.get("days", []):
            for truck_name in ["truck1", "truck2"]:
                truck = day.get(truck_name, {})
                for delivery in truck.get("deliveries", []):
                    st = delivery.get("service_type", "")
                    if st in counts:
                        counts[st] += 1

        result.info(f"SL (Supply & Lay): {counts['SL']} deliveries")
        result.info(f"SD (Supply & Delivery): {counts['SD']} deliveries")
        result.info(f"P (Project): {counts['P']} deliveries")

        total = sum(counts.values())
        result.info(f"Total deliveries: {total}")

        if total == 0:
            result.warn("No deliveries found in schedule")

    except Exception as e:
        result.fail(f"Error: {e}")
    return result


def test_laying_cost_examples():
    """Test 9: Specific Laying Cost Examples"""
    result = TestResult("Laying Cost Examples")
    try:
        resp = requests.get(f"{API_BASE}/api/v1/schedule", timeout=30)
        data = resp.json()

        examples = []

        for day in data.get("days", []):
            for truck_name in ["truck1", "truck2"]:
                truck = day.get(truck_name, {})
                for delivery in truck.get("deliveries", []):
                    if delivery.get("service_type") == "SL" and len(examples) < 5:
                        sqm = delivery.get("sqm")
                        cost = delivery.get("laying_cost")
                        examples.append(f"{sqm} SQM x $2.20 = ${cost}")

        if examples:
            for ex in examples:
                result.info(ex)
        else:
            result.warn("No SL deliveries found to show examples")

    except Exception as e:
        result.fail(f"Error: {e}")
    return result


def main():
    print("=" * 70)
    print("MILESTONE 2 - COMPREHENSIVE TEST SUITE")
    print("GLC Turf Supply Dashboard")
    print("=" * 70)
    print(f"Test Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"API Base: {API_BASE}")
    print("=" * 70)
    print()

    tests = [
        test_api_health,
        test_schedule_endpoint,
        test_truck_data_structure,
        test_laying_cost_calculation,
        test_laying_cost_examples,
        test_january_data,
        test_service_types,
        test_cache_refresh,
        test_frontend_config,
    ]

    results = []
    passed = 0
    failed = 0

    for test_func in tests:
        print(f"Running: {test_func.__doc__}")
        result = test_func()
        results.append(result)
        if result.passed:
            passed += 1
        else:
            failed += 1
        print(result)
        print()

    print("=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print(f"Total Tests: {len(tests)}")
    print(f"Passed: {passed}")
    print(f"Failed: {failed}")
    print()

    if failed == 0:
        print("ALL TESTS PASSED - Milestone 2 is complete!")
    else:
        print(f"WARNING: {failed} test(s) failed - review above for details")

    print("=" * 70)

    return failed == 0


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
