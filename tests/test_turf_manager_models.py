"""
Comprehensive tests for Turf Manager Pydantic models.

Tests cover:
1. Model instantiation with valid data
2. Field validation and constraints
3. Default values
4. Edge cases (zero values, negative values, etc.)
5. Serialization/deserialization
"""
import pytest
from pydantic import ValidationError

from app.models.turf_manager import (
    VarietyStats,
    VarietyTotals,
    DeliveryFees,
    LayingStats,
    FinancialTotals,
    TurfManagerResponse,
)


class TestVarietyStats:
    """Tests for VarietyStats model."""

    def test_valid_variety_stats(self):
        """Test creating valid VarietyStats."""
        stats = VarietyStats(
            variety="Empire Zoysia",
            sqm_sold=500.0,
            sell_price=6000.0,
            cost=4125.0
        )
        assert stats.variety == "Empire Zoysia"
        assert stats.sqm_sold == 500.0
        assert stats.sell_price == 6000.0
        assert stats.cost == 4125.0

    def test_variety_stats_defaults(self):
        """Test VarietyStats default values."""
        stats = VarietyStats(variety="Sir Walter")
        assert stats.sqm_sold == 0
        assert stats.sell_price == 0
        assert stats.cost == 0

    def test_variety_stats_zero_values(self):
        """Test VarietyStats with zero values."""
        stats = VarietyStats(
            variety="Wintergreen Couch",
            sqm_sold=0,
            sell_price=0,
            cost=0
        )
        assert stats.sqm_sold == 0
        assert stats.sell_price == 0
        assert stats.cost == 0

    def test_variety_stats_requires_variety(self):
        """Test that variety field is required."""
        with pytest.raises(ValidationError) as exc_info:
            VarietyStats()
        assert "variety" in str(exc_info.value)

    def test_variety_stats_negative_sqm_rejected(self):
        """Test that negative SQM values are rejected."""
        with pytest.raises(ValidationError) as exc_info:
            VarietyStats(variety="Test", sqm_sold=-100)
        assert "sqm_sold" in str(exc_info.value)

    def test_variety_stats_serialization(self):
        """Test VarietyStats serialization to dict."""
        stats = VarietyStats(
            variety="Empire Zoysia",
            sqm_sold=100.5,
            sell_price=1206.0,
            cost=828.75
        )
        data = stats.model_dump()
        assert data["variety"] == "Empire Zoysia"
        assert data["sqm_sold"] == 100.5
        assert data["sell_price"] == 1206.0
        assert data["cost"] == 828.75


class TestVarietyTotals:
    """Tests for VarietyTotals model."""

    def test_valid_variety_totals(self):
        """Test creating valid VarietyTotals."""
        totals = VarietyTotals(
            sqm_sold=1300.0,
            sell_price=14490.0,
            cost=9300.0
        )
        assert totals.sqm_sold == 1300.0
        assert totals.sell_price == 14490.0
        assert totals.cost == 9300.0

    def test_variety_totals_defaults(self):
        """Test VarietyTotals default values."""
        totals = VarietyTotals()
        assert totals.sqm_sold == 0
        assert totals.sell_price == 0
        assert totals.cost == 0

    def test_variety_totals_gross_profit_calculation(self):
        """Test calculating gross profit from totals."""
        totals = VarietyTotals(
            sqm_sold=1000,
            sell_price=12000.0,
            cost=8250.0
        )
        gross_profit = totals.sell_price - totals.cost
        assert gross_profit == 3750.0


class TestDeliveryFees:
    """Tests for DeliveryFees model."""

    def test_valid_delivery_fees(self):
        """Test creating valid DeliveryFees."""
        fees = DeliveryFees(
            truck_1=850.0,
            truck_2=720.0,
            total=1570.0
        )
        assert fees.truck_1 == 850.0
        assert fees.truck_2 == 720.0
        assert fees.total == 1570.0

    def test_delivery_fees_defaults(self):
        """Test DeliveryFees default values."""
        fees = DeliveryFees()
        assert fees.truck_1 == 0
        assert fees.truck_2 == 0
        assert fees.total == 0

    def test_delivery_fees_single_truck(self):
        """Test DeliveryFees with only one truck used."""
        fees = DeliveryFees(
            truck_1=500.0,
            truck_2=0,
            total=500.0
        )
        assert fees.truck_1 == 500.0
        assert fees.truck_2 == 0
        assert fees.total == 500.0

    def test_delivery_fees_negative_rejected(self):
        """Test that negative delivery fees are rejected."""
        with pytest.raises(ValidationError):
            DeliveryFees(truck_1=-100)


class TestLayingStats:
    """Tests for LayingStats model."""

    def test_valid_laying_stats(self):
        """Test creating valid LayingStats."""
        laying = LayingStats(
            sales=2860.0,
            costs=2200.0
        )
        assert laying.sales == 2860.0
        assert laying.costs == 2200.0

    def test_laying_stats_defaults(self):
        """Test LayingStats default values."""
        laying = LayingStats()
        assert laying.sales == 0
        assert laying.costs == 0

    def test_laying_stats_net_profit(self):
        """Test calculating laying net profit."""
        laying = LayingStats(
            sales=2860.0,
            costs=2200.0
        )
        net = laying.sales - laying.costs
        assert net == 660.0

    def test_laying_stats_break_even(self):
        """Test laying stats at break-even point."""
        laying = LayingStats(
            sales=2200.0,
            costs=2200.0
        )
        assert laying.sales == laying.costs

    def test_laying_cost_calculation(self):
        """Test laying cost calculation at $2.20/SQM."""
        sqm = 1000
        expected_cost = sqm * 2.20
        laying = LayingStats(costs=expected_cost)
        assert laying.costs == 2200.0


class TestFinancialTotals:
    """Tests for FinancialTotals model."""

    def test_valid_financial_totals(self):
        """Test creating valid FinancialTotals."""
        totals = FinancialTotals(
            sales=18920.0,
            costs=12160.0,
            margin_percent=35.7
        )
        assert totals.sales == 18920.0
        assert totals.costs == 12160.0
        assert totals.margin_percent == 35.7

    def test_financial_totals_defaults(self):
        """Test FinancialTotals default values."""
        totals = FinancialTotals()
        assert totals.sales == 0
        assert totals.costs == 0
        assert totals.margin_percent == 0

    def test_margin_calculation(self):
        """Test margin percentage calculation."""
        sales = 18920.0
        costs = 12160.0
        expected_margin = ((sales - costs) / sales) * 100
        totals = FinancialTotals(
            sales=sales,
            costs=costs,
            margin_percent=round(expected_margin, 1)
        )
        assert totals.margin_percent == pytest.approx(35.7, rel=0.1)

    def test_zero_margin(self):
        """Test zero margin (break-even)."""
        totals = FinancialTotals(
            sales=10000.0,
            costs=10000.0,
            margin_percent=0
        )
        assert totals.margin_percent == 0

    def test_negative_margin_allowed(self):
        """Test that negative margin is allowed (for losses)."""
        totals = FinancialTotals(
            sales=10000.0,
            costs=12000.0,
            margin_percent=-20.0
        )
        assert totals.margin_percent == -20.0


class TestTurfManagerResponse:
    """Tests for TurfManagerResponse model."""

    def test_valid_week_response(self):
        """Test creating valid week view response."""
        response = TurfManagerResponse(
            success=True,
            view="week",
            period="Week of Jan-05",
            by_variety=[
                VarietyStats(variety="Empire Zoysia", sqm_sold=500, sell_price=6000, cost=4125),
                VarietyStats(variety="Sir Walter", sqm_sold=300, sell_price=3300, cost=2250),
            ],
            variety_totals=VarietyTotals(sqm_sold=800, sell_price=9300, cost=6375),
            delivery_fees=DeliveryFees(truck_1=500, truck_2=300, total=800),
            laying=LayingStats(sales=1760, costs=1760),
            totals=FinancialTotals(sales=11860, costs=8135, margin_percent=31.4)
        )
        assert response.success is True
        assert response.view == "week"
        assert len(response.by_variety) == 2

    def test_valid_day_response(self):
        """Test creating valid day view response."""
        response = TurfManagerResponse(
            success=True,
            view="day",
            period="Monday, January 05, 2026"
        )
        assert response.view == "day"

    def test_valid_month_response(self):
        """Test creating valid month view response."""
        response = TurfManagerResponse(
            success=True,
            view="month",
            period="January 2026"
        )
        assert response.view == "month"

    def test_valid_annual_response(self):
        """Test creating valid annual view response."""
        response = TurfManagerResponse(
            success=True,
            view="annual",
            period="Year 2026"
        )
        assert response.view == "annual"

    def test_invalid_view_rejected(self):
        """Test that invalid view values are rejected."""
        with pytest.raises(ValidationError) as exc_info:
            TurfManagerResponse(
                success=True,
                view="quarterly",  # Invalid
                period="Q1 2026"
            )
        assert "view" in str(exc_info.value)

    def test_response_serialization(self):
        """Test full response serialization."""
        response = TurfManagerResponse(
            success=True,
            view="week",
            period="Week of Jan-05",
            by_variety=[VarietyStats(variety="Empire Zoysia", sqm_sold=100)],
            variety_totals=VarietyTotals(sqm_sold=100, sell_price=1200, cost=825),
            delivery_fees=DeliveryFees(total=150),
            laying=LayingStats(sales=220, costs=220),
            totals=FinancialTotals(sales=1570, costs=1045, margin_percent=33.4)
        )
        data = response.model_dump()
        assert data["success"] is True
        assert data["view"] == "week"
        assert len(data["by_variety"]) == 1
        assert data["variety_totals"]["sqm_sold"] == 100

    def test_empty_variety_list(self):
        """Test response with empty variety list."""
        response = TurfManagerResponse(
            success=True,
            view="week",
            period="Week of Jan-05",
            by_variety=[]
        )
        assert response.by_variety == []

    def test_response_with_defaults(self):
        """Test response uses default nested models."""
        response = TurfManagerResponse(
            success=True,
            view="week",
            period="Test Period"
        )
        assert response.variety_totals.sqm_sold == 0
        assert response.delivery_fees.total == 0
        assert response.laying.sales == 0
        assert response.totals.margin_percent == 0
