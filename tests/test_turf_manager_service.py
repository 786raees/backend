"""
Comprehensive tests for TurfManagerService business logic.

Tests cover:
1. Week tab name calculation
2. Data parsing (floats, currency, percentages)
3. Data aggregation by variety
4. Financial calculations (margin, totals)
5. Laying cost calculation ($2.20/SQM)
6. Edge cases and error handling
"""
import pytest
from datetime import date
from unittest.mock import Mock, patch, MagicMock

from app.api.v1.routes.turf_manager import TurfManagerService, turf_manager_service


class TestWeekTabNameCalculation:
    """Tests for week tab name calculation."""

    def test_monday_returns_same_date(self):
        """Monday should return its own date as tab name."""
        service = TurfManagerService()
        # Jan 5, 2026 is a Monday
        result = service.get_week_tab_name(date(2026, 1, 5))
        assert result == "Jan-05"

    def test_friday_returns_monday_of_week(self):
        """Friday should return Monday of that week."""
        service = TurfManagerService()
        # Jan 9, 2026 is a Friday
        result = service.get_week_tab_name(date(2026, 1, 9))
        assert result == "Jan-05"

    def test_wednesday_returns_monday_of_week(self):
        """Wednesday should return Monday of that week."""
        service = TurfManagerService()
        # Jan 7, 2026 is a Wednesday
        result = service.get_week_tab_name(date(2026, 1, 7))
        assert result == "Jan-05"

    def test_sunday_returns_monday_of_week(self):
        """Sunday should return Monday of that week."""
        service = TurfManagerService()
        # Jan 4, 2026 is a Sunday (belongs to Dec-29 week)
        result = service.get_week_tab_name(date(2026, 1, 4))
        assert result == "Dec-29"

    def test_cross_year_boundary(self):
        """Test week crossing year boundary."""
        service = TurfManagerService()
        # Dec 31, 2025 is a Wednesday (belongs to Dec-29 week)
        result = service.get_week_tab_name(date(2025, 12, 31))
        assert result == "Dec-29"

    def test_february_date(self):
        """Test date in February."""
        service = TurfManagerService()
        # Feb 10, 2026 is a Tuesday
        result = service.get_week_tab_name(date(2026, 2, 10))
        assert result == "Feb-09"


class TestFloatParsing:
    """Tests for float parsing utility."""

    def test_parse_integer(self):
        """Test parsing integer."""
        service = TurfManagerService()
        assert service._parse_float(100) == 100.0

    def test_parse_float(self):
        """Test parsing float."""
        service = TurfManagerService()
        assert service._parse_float(123.45) == 123.45

    def test_parse_string_number(self):
        """Test parsing string number."""
        service = TurfManagerService()
        assert service._parse_float("500") == 500.0

    def test_parse_currency_with_dollar_sign(self):
        """Test parsing currency with $ sign."""
        service = TurfManagerService()
        assert service._parse_float("$1,234.56") == 1234.56

    def test_parse_currency_with_commas(self):
        """Test parsing currency with commas."""
        service = TurfManagerService()
        assert service._parse_float("1,000,000") == 1000000.0

    def test_parse_empty_string(self):
        """Test parsing empty string returns 0."""
        service = TurfManagerService()
        assert service._parse_float("") == 0.0

    def test_parse_none(self):
        """Test parsing None returns 0."""
        service = TurfManagerService()
        assert service._parse_float(None) == 0.0

    def test_parse_invalid_string(self):
        """Test parsing invalid string returns 0."""
        service = TurfManagerService()
        assert service._parse_float("N/A") == 0.0

    def test_parse_whitespace(self):
        """Test parsing whitespace returns 0."""
        service = TurfManagerService()
        assert service._parse_float("  ") == 0.0


class TestPercentageParsing:
    """Tests for percentage parsing utility."""

    def test_parse_percentage_with_symbol(self):
        """Test parsing percentage with % symbol."""
        service = TurfManagerService()
        assert service._parse_percentage("35.7%") == 35.7

    def test_parse_percentage_without_symbol(self):
        """Test parsing percentage without % symbol."""
        service = TurfManagerService()
        assert service._parse_percentage("35.7") == 35.7

    def test_parse_percentage_integer(self):
        """Test parsing integer percentage."""
        service = TurfManagerService()
        assert service._parse_percentage(50) == 50.0

    def test_parse_percentage_empty(self):
        """Test parsing empty percentage."""
        service = TurfManagerService()
        assert service._parse_percentage("") == 0.0


class TestDataAggregation:
    """Tests for data aggregation logic."""

    def test_aggregate_single_row(self, sample_week_data):
        """Test aggregating single data row."""
        service = TurfManagerService()
        # Use only first data row
        data = [sample_week_data[0], sample_week_data[1]]
        result = service._aggregate_week_data(data)

        assert len(result["by_variety"]) == 1
        assert result["by_variety"][0].variety == "Empire Zoysia"
        assert result["by_variety"][0].sqm_sold == 100.0

    def test_aggregate_multiple_same_variety(self, sample_week_data):
        """Test aggregating multiple rows of same variety."""
        service = TurfManagerService()
        result = service._aggregate_week_data(sample_week_data)

        # Empire Zoysia appears twice (100 + 200 = 300 SQM)
        empire = next(v for v in result["by_variety"] if v.variety == "Empire Zoysia")
        assert empire.sqm_sold == 300.0

    def test_aggregate_different_varieties(self, sample_week_data):
        """Test aggregating different varieties."""
        service = TurfManagerService()
        result = service._aggregate_week_data(sample_week_data)

        varieties = {v.variety for v in result["by_variety"]}
        assert "Empire Zoysia" in varieties
        assert "Sir Walter" in varieties

    def test_aggregate_empty_data(self, empty_week_data):
        """Test aggregating empty data (only headers)."""
        service = TurfManagerService()
        result = service._aggregate_week_data(empty_week_data)

        assert len(result["by_variety"]) == 0
        assert result["variety_totals"].sqm_sold == 0

    def test_aggregate_skip_empty_variety(self, malformed_week_data):
        """Test that rows with empty variety are skipped."""
        service = TurfManagerService()
        result = service._aggregate_week_data(malformed_week_data)

        # Only first data row should be included
        assert len(result["by_variety"]) == 1

    def test_aggregate_variety_totals(self, sample_week_data):
        """Test variety totals calculation."""
        service = TurfManagerService()
        result = service._aggregate_week_data(sample_week_data)

        # Total SQM: 100 + 150 + 200 = 450
        assert result["variety_totals"].sqm_sold == 450.0

    def test_aggregate_delivery_fees(self, sample_week_data):
        """Test delivery fees aggregation."""
        service = TurfManagerService()
        result = service._aggregate_week_data(sample_week_data)

        # T1: 150 + 200 = 350, T2: 180
        assert result["delivery_fees"].truck_1 == 350.0
        assert result["delivery_fees"].truck_2 == 180.0
        assert result["delivery_fees"].total == 530.0


class TestLayingCostCalculation:
    """Tests for laying cost calculation."""

    def test_laying_cost_per_sqm(self):
        """Test laying cost constant is $2.20/SQM."""
        service = TurfManagerService()
        assert service.LAYING_COST_PER_SQM == 2.20

    def test_laying_costs_for_all_deliveries(self, sample_week_data):
        """Test laying costs calculated for ALL deliveries, not just SL."""
        service = TurfManagerService()
        result = service._aggregate_week_data(sample_week_data)

        # Total SQM: 450, Laying cost: 450 * 2.20 = 990
        expected_laying_costs = 450 * 2.20
        assert result["laying"].costs == expected_laying_costs

    def test_laying_sales_aggregation(self, sample_week_data):
        """Test laying sales (fees) aggregation."""
        service = TurfManagerService()
        result = service._aggregate_week_data(sample_week_data)

        # Laying fees: 220 + 0 + 440 = 660
        assert result["laying"].sales == 660.0


class TestMarginCalculation:
    """Tests for margin calculation."""

    def test_margin_calculation_formula(self, sample_week_data):
        """Test margin calculation: ((Sales - Costs) / Sales) * 100."""
        service = TurfManagerService()
        result = service._aggregate_week_data(sample_week_data)

        totals = result["totals"]
        if totals.sales > 0:
            expected_margin = ((totals.sales - totals.costs) / totals.sales) * 100
            assert totals.margin_percent == pytest.approx(expected_margin, rel=0.1)

    def test_zero_sales_margin(self, empty_week_data):
        """Test margin is 0 when sales are 0."""
        service = TurfManagerService()
        result = service._aggregate_week_data(empty_week_data)

        assert result["totals"].margin_percent == 0

    def test_total_sales_includes_all_revenue(self, sample_week_data):
        """Test total sales includes turf + delivery + laying."""
        service = TurfManagerService()
        result = service._aggregate_week_data(sample_week_data)

        expected_total = (
            result["variety_totals"].sell_price +
            result["delivery_fees"].total +
            result["laying"].sales
        )
        assert result["totals"].sales == expected_total

    def test_total_costs_calculation(self, sample_week_data):
        """Test total costs includes turf cost + laying costs."""
        service = TurfManagerService()
        result = service._aggregate_week_data(sample_week_data)

        expected_costs = result["variety_totals"].cost + result["laying"].costs
        assert result["totals"].costs == expected_costs


class TestColumnMapping:
    """Tests for column index mapping."""

    def test_column_indices(self):
        """Test column index mapping is correct."""
        service = TurfManagerService()

        assert service.COLUMNS["day_slot"] == 0       # A
        assert service.COLUMNS["variety"] == 1        # B
        assert service.COLUMNS["suburb"] == 2         # C
        assert service.COLUMNS["service_type"] == 3   # D
        assert service.COLUMNS["sqm"] == 4            # E
        assert service.COLUMNS["pallets"] == 5        # F
        assert service.COLUMNS["sell_per_sqm"] == 6   # G
        assert service.COLUMNS["cost_per_sqm"] == 7   # H
        assert service.COLUMNS["delivery_fee"] == 8   # I
        assert service.COLUMNS["laying_fee"] == 9     # J
        assert service.COLUMNS["turf_revenue"] == 10  # K
        assert service.COLUMNS["delivery_t1"] == 11   # L
        assert service.COLUMNS["delivery_t2"] == 12   # M
        assert service.COLUMNS["laying_total"] == 13  # N
        assert service.COLUMNS["total"] == 14         # O
        assert service.COLUMNS["cost"] == 15          # P
        assert service.COLUMNS["margin"] == 16        # Q


class TestTurfRevenueCalculation:
    """Tests for turf revenue calculation."""

    def test_turf_revenue_formula(self):
        """Test turf revenue = SQM × Sell $/SQM."""
        service = TurfManagerService()

        # Single row: 100 SQM × $12 = $1200
        data = [
            ["Header", "Variety", "Suburb", "Type", "SQM", "Pallets",
             "Sell $/SQM", "Cost $/SQM", "Del", "Lay"],
            ["Monday", "Empire Zoysia", "Brisbane", "SL", "100", "2.5",
             "12", "8.25", "150", "220"],
        ]
        result = service._aggregate_week_data(data)

        assert result["by_variety"][0].sell_price == 1200.0

    def test_turf_cost_formula(self):
        """Test turf cost = SQM × Cost $/SQM."""
        service = TurfManagerService()

        # Single row: 100 SQM × $8.25 = $825
        data = [
            ["Header", "Variety", "Suburb", "Type", "SQM", "Pallets",
             "Sell $/SQM", "Cost $/SQM", "Del", "Lay"],
            ["Monday", "Empire Zoysia", "Brisbane", "SL", "100", "2.5",
             "12", "8.25", "150", "220"],
        ]
        result = service._aggregate_week_data(data)

        assert result["by_variety"][0].cost == 825.0
