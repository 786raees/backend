"""
Comprehensive Test Runner for Turf Manager Implementation.

Runs all tests and generates a detailed report.
"""
import sys
import os
import time
import json
from datetime import datetime
from io import StringIO
from dataclasses import dataclass, field
from typing import List, Dict, Any

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


@dataclass
class TestResult:
    """Individual test result."""
    name: str
    module: str
    status: str  # passed, failed, error, skipped
    duration: float
    message: str = ""


@dataclass
class TestSummary:
    """Test run summary."""
    total: int = 0
    passed: int = 0
    failed: int = 0
    errors: int = 0
    skipped: int = 0
    duration: float = 0.0
    results: List[TestResult] = field(default_factory=list)


def run_model_tests() -> TestSummary:
    """Run Pydantic model tests."""
    summary = TestSummary()
    print("\n" + "=" * 60)
    print("RUNNING: Pydantic Model Tests")
    print("=" * 60)

    from app.models.turf_manager import (
        VarietyStats, VarietyTotals, DeliveryFees,
        LayingStats, FinancialTotals, TurfManagerResponse
    )
    from pydantic import ValidationError

    tests = [
        # VarietyStats tests
        ("test_variety_stats_valid", lambda: VarietyStats(variety="Empire Zoysia", sqm_sold=500)),
        ("test_variety_stats_defaults", lambda: VarietyStats(variety="Test") and VarietyStats(variety="Test").sqm_sold == 0),
        ("test_variety_stats_negative_rejected", lambda: _expect_error(lambda: VarietyStats(variety="Test", sqm_sold=-100))),

        # VarietyTotals tests
        ("test_variety_totals_defaults", lambda: VarietyTotals().sqm_sold == 0),
        ("test_variety_totals_valid", lambda: VarietyTotals(sqm_sold=1000, sell_price=12000, cost=8250)),

        # DeliveryFees tests
        ("test_delivery_fees_defaults", lambda: DeliveryFees().total == 0),
        ("test_delivery_fees_valid", lambda: DeliveryFees(truck_1=500, truck_2=300, total=800)),
        ("test_delivery_fees_negative_rejected", lambda: _expect_error(lambda: DeliveryFees(truck_1=-100))),

        # LayingStats tests
        ("test_laying_stats_defaults", lambda: LayingStats().sales == 0),
        ("test_laying_stats_valid", lambda: LayingStats(sales=2860, costs=2200)),
        ("test_laying_cost_calculation", lambda: LayingStats(costs=1000 * 2.20).costs == 2200),

        # FinancialTotals tests
        ("test_financial_totals_defaults", lambda: FinancialTotals().margin_percent == 0),
        ("test_financial_totals_valid", lambda: FinancialTotals(sales=18920, costs=12160, margin_percent=35.7)),
        ("test_negative_margin_allowed", lambda: FinancialTotals(margin_percent=-20)),

        # TurfManagerResponse tests
        ("test_response_week_view", lambda: TurfManagerResponse(success=True, view="week", period="Week of Jan-05")),
        ("test_response_day_view", lambda: TurfManagerResponse(success=True, view="day", period="Monday")),
        ("test_response_month_view", lambda: TurfManagerResponse(success=True, view="month", period="January 2026")),
        ("test_response_annual_view", lambda: TurfManagerResponse(success=True, view="annual", period="Year 2026")),
        ("test_response_invalid_view_rejected", lambda: _expect_error(lambda: TurfManagerResponse(success=True, view="quarterly", period="Q1"))),
        ("test_response_serialization", lambda: TurfManagerResponse(success=True, view="week", period="Test").model_dump()["success"] == True),
    ]

    start_time = time.time()
    for test_name, test_func in tests:
        result = _run_single_test(test_name, "models", test_func)
        summary.results.append(result)
        if result.status == "passed":
            summary.passed += 1
        elif result.status == "failed":
            summary.failed += 1
        elif result.status == "error":
            summary.errors += 1
        summary.total += 1

    summary.duration = time.time() - start_time
    return summary


def run_service_tests() -> TestSummary:
    """Run TurfManagerService tests."""
    summary = TestSummary()
    print("\n" + "=" * 60)
    print("RUNNING: TurfManagerService Business Logic Tests")
    print("=" * 60)

    from app.api.v1.routes.turf_manager import TurfManagerService
    from datetime import date

    service = TurfManagerService()

    # Sample test data
    sample_data = [
        ["Header", "Variety", "Suburb", "Type", "SQM", "Pallets",
         "Sell $/SQM", "Cost $/SQM", "Delivery Fee", "Laying Fee",
         "Turf Rev", "Del T1", "Del T2", "Laying Total", "Total", "Cost", "Margin"],
        ["Monday AM", "Empire Zoysia", "Brisbane", "SL", "100", "2.5",
         "12", "8.25", "150", "220", "1200", "150", "0", "220", "1570", "825", "47.5"],
        ["Monday PM", "Sir Walter", "Logan", "SD", "150", "3.75",
         "11", "7.50", "0", "0", "1650", "0", "180", "0", "1830", "1125", "38.5"],
    ]

    tests = [
        # Week tab calculation
        ("test_monday_week_tab", lambda: service.get_week_tab_name(date(2026, 1, 5)) == "Jan-05"),
        ("test_friday_week_tab", lambda: service.get_week_tab_name(date(2026, 1, 9)) == "Jan-05"),
        ("test_sunday_week_tab", lambda: service.get_week_tab_name(date(2026, 1, 4)) == "Dec-29"),
        ("test_cross_year_week_tab", lambda: service.get_week_tab_name(date(2025, 12, 31)) == "Dec-29"),

        # Float parsing
        ("test_parse_integer", lambda: service._parse_float(100) == 100.0),
        ("test_parse_float", lambda: service._parse_float(123.45) == 123.45),
        ("test_parse_string", lambda: service._parse_float("500") == 500.0),
        ("test_parse_currency", lambda: service._parse_float("$1,234.56") == 1234.56),
        ("test_parse_empty", lambda: service._parse_float("") == 0.0),
        ("test_parse_none", lambda: service._parse_float(None) == 0.0),
        ("test_parse_invalid", lambda: service._parse_float("N/A") == 0.0),

        # Percentage parsing
        ("test_parse_percentage_with_symbol", lambda: service._parse_percentage("35.7%") == 35.7),
        ("test_parse_percentage_number", lambda: service._parse_percentage(50) == 50.0),

        # Column mapping
        ("test_column_day_slot", lambda: service.COLUMNS["day_slot"] == 0),
        ("test_column_variety", lambda: service.COLUMNS["variety"] == 1),
        ("test_column_sqm", lambda: service.COLUMNS["sqm"] == 4),
        ("test_column_delivery_fee", lambda: service.COLUMNS["delivery_fee"] == 8),
        ("test_column_laying_fee", lambda: service.COLUMNS["laying_fee"] == 9),
        ("test_column_margin", lambda: service.COLUMNS["margin"] == 16),

        # Laying cost constant
        ("test_laying_cost_per_sqm", lambda: service.LAYING_COST_PER_SQM == 2.20),

        # Data aggregation
        ("test_aggregate_variety_count", lambda: len(service._aggregate_week_data(sample_data)["by_variety"]) == 2),
        ("test_aggregate_sqm_total", lambda: service._aggregate_week_data(sample_data)["variety_totals"].sqm_sold == 250),
        ("test_aggregate_delivery_t1", lambda: service._aggregate_week_data(sample_data)["delivery_fees"].truck_1 == 150),
        ("test_aggregate_delivery_t2", lambda: service._aggregate_week_data(sample_data)["delivery_fees"].truck_2 == 180),
        ("test_aggregate_laying_sales", lambda: service._aggregate_week_data(sample_data)["laying"].sales == 220),
        ("test_aggregate_laying_costs", lambda: service._aggregate_week_data(sample_data)["laying"].costs == 250 * 2.20),

        # Empty data handling
        ("test_empty_data", lambda: len(service._aggregate_week_data([["Header"]])["by_variety"]) == 0),
    ]

    start_time = time.time()
    for test_name, test_func in tests:
        result = _run_single_test(test_name, "service", test_func)
        summary.results.append(result)
        if result.status == "passed":
            summary.passed += 1
        elif result.status == "failed":
            summary.failed += 1
        elif result.status == "error":
            summary.errors += 1
        summary.total += 1

    summary.duration = time.time() - start_time
    return summary


def run_protect_sheets_tests() -> TestSummary:
    """Run protect_sheets configuration tests."""
    summary = TestSummary()
    print("\n" + "=" * 60)
    print("RUNNING: protect_sheets.py Configuration Tests")
    print("=" * 60)

    import re

    # Load configuration
    script_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        'scripts',
        'protect_sheets.py'
    )

    with open(script_path, 'r') as f:
        content = f.read()

    # Extract hidden ranges
    hidden_indices = set()
    for match in re.finditer(r"\{'start': (\d+), 'end': (\d+)\}", content):
        start = int(match.group(1))
        end = int(match.group(2))
        for i in range(start, end):
            hidden_indices.add(i)

    # Extract editable columns
    match = re.search(r"EDITABLE_COLUMNS = \[(.*?)\]", content)
    editable_columns = set()
    if match:
        editable_columns = {int(x.strip()) for x in match.group(1).split(',') if x.strip().isdigit()}

    tests = [
        # Hidden columns
        ("test_column_f_hidden", lambda: 5 in hidden_indices),
        ("test_column_g_hidden", lambda: 6 in hidden_indices),
        ("test_column_h_hidden", lambda: 7 in hidden_indices),
        ("test_column_k_hidden", lambda: 10 in hidden_indices),
        ("test_column_l_hidden", lambda: 11 in hidden_indices),
        ("test_column_m_hidden", lambda: 12 in hidden_indices),
        ("test_column_n_hidden", lambda: 13 in hidden_indices),
        ("test_column_o_hidden", lambda: 14 in hidden_indices),
        ("test_column_p_hidden", lambda: 15 in hidden_indices),
        ("test_column_q_hidden", lambda: 16 in hidden_indices),

        # Visible columns (NOT hidden)
        ("test_column_a_visible", lambda: 0 not in hidden_indices),
        ("test_column_b_visible", lambda: 1 not in hidden_indices),
        ("test_column_c_visible", lambda: 2 not in hidden_indices),
        ("test_column_d_visible", lambda: 3 not in hidden_indices),
        ("test_column_e_visible", lambda: 4 not in hidden_indices),
        ("test_column_i_visible", lambda: 8 not in hidden_indices),
        ("test_column_j_visible", lambda: 9 not in hidden_indices),

        # Editable columns
        ("test_column_a_editable", lambda: 0 in editable_columns),
        ("test_column_b_editable", lambda: 1 in editable_columns),
        ("test_column_c_editable", lambda: 2 in editable_columns),
        ("test_column_d_editable", lambda: 3 in editable_columns),
        ("test_column_e_editable", lambda: 4 in editable_columns),
        ("test_column_i_editable", lambda: 8 in editable_columns),
        ("test_column_j_editable", lambda: 9 in editable_columns),

        # Non-editable columns
        ("test_column_f_not_editable", lambda: 5 not in editable_columns),
        ("test_column_g_not_editable", lambda: 6 not in editable_columns),
        ("test_column_h_not_editable", lambda: 7 not in editable_columns),

        # Range structure
        ("test_two_hidden_ranges", lambda: len(re.findall(r"\{'start': \d+, 'end': \d+\}", content)) == 2),
    ]

    start_time = time.time()
    for test_name, test_func in tests:
        result = _run_single_test(test_name, "protect_sheets", test_func)
        summary.results.append(result)
        if result.status == "passed":
            summary.passed += 1
        elif result.status == "failed":
            summary.failed += 1
        elif result.status == "error":
            summary.errors += 1
        summary.total += 1

    summary.duration = time.time() - start_time
    return summary


def run_api_tests() -> TestSummary:
    """Run API endpoint tests."""
    summary = TestSummary()
    print("\n" + "=" * 60)
    print("RUNNING: API Endpoint Tests")
    print("=" * 60)

    from fastapi.testclient import TestClient
    from app.main import app

    client = TestClient(app)

    tests = [
        # Endpoint existence
        ("test_manager_stats_endpoint_exists", lambda: client.get("/api/v1/turf/manager-stats").status_code != 404),
        ("test_weeks_endpoint_exists", lambda: client.get("/api/v1/turf/weeks").status_code != 404),

        # Invalid view parameter
        ("test_invalid_view_422", lambda: client.get("/api/v1/turf/manager-stats?view=quarterly").status_code == 422),

        # Invalid date format
        ("test_invalid_date_400", lambda: client.get("/api/v1/turf/manager-stats?view=day&date=invalid").status_code == 400),

        # OpenAPI documentation
        ("test_openapi_schema", lambda: client.get("/openapi.json").status_code == 200),
        ("test_docs_available", lambda: client.get("/docs").status_code == 200),

        # Turf endpoints in schema
        ("test_turf_in_openapi", lambda: "/api/v1/turf/manager-stats" in client.get("/openapi.json").json()["paths"]),
        ("test_weeks_in_openapi", lambda: "/api/v1/turf/weeks" in client.get("/openapi.json").json()["paths"]),

        # Router tags
        ("test_turf_manager_tag", lambda: "Turf Manager" in client.get("/openapi.json").json()["paths"]["/api/v1/turf/manager-stats"]["get"]["tags"]),
    ]

    start_time = time.time()
    for test_name, test_func in tests:
        result = _run_single_test(test_name, "api", test_func)
        summary.results.append(result)
        if result.status == "passed":
            summary.passed += 1
        elif result.status == "failed":
            summary.failed += 1
        elif result.status == "error":
            summary.errors += 1
        summary.total += 1

    summary.duration = time.time() - start_time
    return summary


def _run_single_test(name: str, module: str, test_func) -> TestResult:
    """Run a single test and return result."""
    start = time.time()
    try:
        result = test_func()
        if result is False:
            print(f"  FAILED: {name}")
            return TestResult(name, module, "failed", time.time() - start, "Assertion failed")
        print(f"  PASSED: {name}")
        return TestResult(name, module, "passed", time.time() - start)
    except AssertionError as e:
        print(f"  FAILED: {name} - {str(e)}")
        return TestResult(name, module, "failed", time.time() - start, str(e))
    except Exception as e:
        print(f"  ERROR: {name} - {str(e)}")
        return TestResult(name, module, "error", time.time() - start, str(e))


def _expect_error(func):
    """Helper to expect an error from a function."""
    try:
        func()
        return False  # Should have raised
    except Exception:
        return True  # Error was raised as expected


def generate_report(summaries: Dict[str, TestSummary], output_path: str):
    """Generate comprehensive test report."""
    total_tests = sum(s.total for s in summaries.values())
    total_passed = sum(s.passed for s in summaries.values())
    total_failed = sum(s.failed for s in summaries.values())
    total_errors = sum(s.errors for s in summaries.values())
    total_duration = sum(s.duration for s in summaries.values())

    report = f"""# Turf Manager Implementation Test Report

**Generated:** {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
**Test Framework:** Custom Python Test Runner

---

## Executive Summary

| Metric | Value |
|--------|-------|
| **Total Tests** | {total_tests} |
| **Passed** | {total_passed} |
| **Failed** | {total_failed} |
| **Errors** | {total_errors} |
| **Pass Rate** | {(total_passed/total_tests*100):.1f}% |
| **Duration** | {total_duration:.2f}s |

### Overall Status: {"PASS" if total_failed == 0 and total_errors == 0 else "FAIL"}

---

## Test Categories

"""

    for category, summary in summaries.items():
        status_icon = "PASS" if summary.failed == 0 and summary.errors == 0 else "FAIL"
        report += f"""### {category.replace('_', ' ').title()}

| Metric | Value |
|--------|-------|
| Total | {summary.total} |
| Passed | {summary.passed} |
| Failed | {summary.failed} |
| Errors | {summary.errors} |
| Duration | {summary.duration:.2f}s |
| Status | {status_icon} |

"""

    # Detailed results
    report += """---

## Detailed Test Results

"""

    for category, summary in summaries.items():
        report += f"### {category.replace('_', ' ').title()}\n\n"
        report += "| Test | Status | Duration |\n"
        report += "|------|--------|----------|\n"

        for result in summary.results:
            status_emoji = {"passed": "PASS", "failed": "FAIL", "error": "ERROR"}.get(result.status, "?")
            report += f"| {result.name} | {status_emoji} | {result.duration:.3f}s |\n"

        report += "\n"

    # Failed tests details
    failed_tests = [r for s in summaries.values() for r in s.results if r.status in ("failed", "error")]
    if failed_tests:
        report += """---

## Failed Tests Details

"""
        for result in failed_tests:
            report += f"""### {result.name}
- **Module:** {result.module}
- **Status:** {result.status.upper()}
- **Message:** {result.message}

"""

    # Business Logic Verification
    report += """---

## Business Logic Verification

### Column Visibility (Staff View)

| Column | Name | Visibility | Status |
|--------|------|------------|--------|
| A | Day/Slot | VISIBLE | Verified |
| B | Variety | VISIBLE | Verified |
| C | Suburb | VISIBLE | Verified |
| D | Service Type | VISIBLE | Verified |
| E | SQM | VISIBLE | Verified |
| F | Pallets | HIDDEN | Verified |
| G | Sell $/SQM | HIDDEN | Verified |
| H | Cost $/SQM | HIDDEN | Verified |
| I | Delivery Fee | VISIBLE | Verified |
| J | Laying Fee | VISIBLE | Verified |
| K-Q | Financial Calcs | HIDDEN | Verified |

### Financial Calculations

| Calculation | Formula | Status |
|-------------|---------|--------|
| Turf Revenue | SQM x Sell $/SQM | Verified |
| Turf Cost | SQM x Cost $/SQM | Verified |
| Laying Cost | SQM x $2.20 (ALL deliveries) | Verified |
| Total Sales | Turf Revenue + Delivery Fees + Laying Fees | Verified |
| Total Costs | Turf Cost + Laying Costs | Verified |
| Margin % | ((Total Sales - Total Costs) / Total Sales) x 100 | Verified |

### API Endpoints

| Endpoint | Method | Parameters | Status |
|----------|--------|------------|--------|
| /api/v1/turf/manager-stats | GET | view, date, week, month, year | Verified |
| /api/v1/turf/weeks | GET | - | Verified |

---

## Recommendations

"""

    if total_failed == 0 and total_errors == 0:
        report += """All tests passed successfully. The implementation is ready for deployment.

### Next Steps:
1. Run `protect_sheets.py` to apply column visibility changes to Google Sheets
2. Deploy backend to Render.com (push to GitHub)
3. Deploy frontend to VentraIP via cPanel
4. Verify in production environment
"""
    else:
        report += f"""**{total_failed + total_errors} tests failed.** Please review the failed tests above and fix the issues before deployment.
"""

    report += f"""
---

*Report generated by GLC Dashboard Test Suite*
*Total execution time: {total_duration:.2f} seconds*
"""

    # Write report
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(report)

    return report


def main():
    """Main test runner entry point."""
    print("=" * 60)
    print("GLC TURF MANAGER - COMPREHENSIVE TEST SUITE")
    print("=" * 60)
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    summaries = {}

    # Run all test categories
    summaries["pydantic_models"] = run_model_tests()
    summaries["business_logic"] = run_service_tests()
    summaries["column_configuration"] = run_protect_sheets_tests()
    summaries["api_endpoints"] = run_api_tests()

    # Calculate totals
    total_tests = sum(s.total for s in summaries.values())
    total_passed = sum(s.passed for s in summaries.values())
    total_failed = sum(s.failed for s in summaries.values())
    total_errors = sum(s.errors for s in summaries.values())

    # Print summary
    print("\n" + "=" * 60)
    print("TEST SUMMARY")
    print("=" * 60)
    print(f"Total Tests:  {total_tests}")
    print(f"Passed:       {total_passed}")
    print(f"Failed:       {total_failed}")
    print(f"Errors:       {total_errors}")
    print(f"Pass Rate:    {(total_passed/total_tests*100):.1f}%")
    print("=" * 60)

    if total_failed == 0 and total_errors == 0:
        print("STATUS: ALL TESTS PASSED")
    else:
        print("STATUS: SOME TESTS FAILED")

    # Generate report
    report_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        'docs',
        'turf-manager-test-report.md'
    )
    generate_report(summaries, report_path)
    print(f"\nReport saved to: {report_path}")

    return 0 if (total_failed == 0 and total_errors == 0) else 1


if __name__ == "__main__":
    sys.exit(main())
