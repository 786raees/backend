"""
Comprehensive tests for Turf Manager API endpoints.

Tests cover:
1. Endpoint routing and HTTP methods
2. Query parameter validation
3. Response schema validation
4. Error handling
5. Edge cases
"""
import pytest
from unittest.mock import patch, AsyncMock, MagicMock
from fastapi.testclient import TestClient

from app.main import app
from app.models.turf_manager import (
    TurfManagerResponse,
    VarietyStats,
    VarietyTotals,
    DeliveryFees,
    LayingStats,
    FinancialTotals,
)


@pytest.fixture
def client():
    """Create test client."""
    return TestClient(app)


@pytest.fixture
def mock_response():
    """Create mock TurfManagerResponse."""
    return TurfManagerResponse(
        success=True,
        view="week",
        period="Week of Jan-05",
        by_variety=[
            VarietyStats(variety="Empire Zoysia", sqm_sold=500, sell_price=6000, cost=4125),
        ],
        variety_totals=VarietyTotals(sqm_sold=500, sell_price=6000, cost=4125),
        delivery_fees=DeliveryFees(truck_1=500, truck_2=300, total=800),
        laying=LayingStats(sales=1100, costs=1100),
        totals=FinancialTotals(sales=7900, costs=5225, margin_percent=33.9)
    )


class TestManagerStatsEndpoint:
    """Tests for GET /api/v1/turf/manager-stats endpoint."""

    def test_endpoint_exists(self, client):
        """Test that endpoint exists and is accessible."""
        # Will fail without Google Sheets credentials, but should return 500, not 404
        response = client.get("/api/v1/turf/manager-stats")
        assert response.status_code != 404

    def test_default_view_is_week(self, client):
        """Test that default view is 'week'."""
        with patch('app.api.v1.routes.turf_manager.turf_manager_service.get_manager_stats') as mock:
            mock.return_value = TurfManagerResponse(
                success=True,
                view="week",
                period="Week of Jan-05"
            )
            response = client.get("/api/v1/turf/manager-stats")
            # Check that the call was made with view="week"
            call_args = mock.call_args
            assert call_args[1]["view"] == "week"

    def test_day_view_requires_date(self, client):
        """Test day view with date parameter."""
        with patch('app.api.v1.routes.turf_manager.turf_manager_service.get_manager_stats') as mock:
            mock.return_value = TurfManagerResponse(
                success=True,
                view="day",
                period="Monday, January 05, 2026"
            )
            response = client.get("/api/v1/turf/manager-stats?view=day&date=2026-01-05")
            call_args = mock.call_args
            assert call_args[1]["view"] == "day"

    def test_invalid_date_format(self, client):
        """Test that invalid date format returns 400."""
        response = client.get("/api/v1/turf/manager-stats?view=day&date=invalid")
        assert response.status_code == 400
        assert "Invalid date format" in response.json()["detail"]

    def test_invalid_view_returns_422(self, client):
        """Test that invalid view parameter returns validation error."""
        response = client.get("/api/v1/turf/manager-stats?view=quarterly")
        assert response.status_code == 422

    def test_week_view_with_week_param(self, client):
        """Test week view with week parameter."""
        with patch('app.api.v1.routes.turf_manager.turf_manager_service.get_manager_stats') as mock:
            mock.return_value = TurfManagerResponse(
                success=True,
                view="week",
                period="Week of Jan-12"
            )
            response = client.get("/api/v1/turf/manager-stats?view=week&week=Jan-12")
            call_args = mock.call_args
            assert call_args[1]["week"] == "Jan-12"

    def test_month_view(self, client):
        """Test month view parameter."""
        with patch('app.api.v1.routes.turf_manager.turf_manager_service.get_manager_stats') as mock:
            mock.return_value = TurfManagerResponse(
                success=True,
                view="month",
                period="January 2026"
            )
            response = client.get("/api/v1/turf/manager-stats?view=month&month=2026-01")
            call_args = mock.call_args
            assert call_args[1]["view"] == "month"

    def test_annual_view(self, client):
        """Test annual view parameter."""
        with patch('app.api.v1.routes.turf_manager.turf_manager_service.get_manager_stats') as mock:
            mock.return_value = TurfManagerResponse(
                success=True,
                view="annual",
                period="Year 2026"
            )
            response = client.get("/api/v1/turf/manager-stats?view=annual&year=2026")
            call_args = mock.call_args
            assert call_args[1]["view"] == "annual"
            assert call_args[1]["year"] == 2026

    def test_response_schema(self, client, mock_response):
        """Test response matches expected schema."""
        with patch('app.api.v1.routes.turf_manager.turf_manager_service.get_manager_stats') as mock:
            mock.return_value = mock_response
            response = client.get("/api/v1/turf/manager-stats")
            data = response.json()

            assert "success" in data
            assert "view" in data
            assert "period" in data
            assert "by_variety" in data
            assert "variety_totals" in data
            assert "delivery_fees" in data
            assert "laying" in data
            assert "totals" in data

    def test_variety_schema(self, client, mock_response):
        """Test variety data in response."""
        with patch('app.api.v1.routes.turf_manager.turf_manager_service.get_manager_stats') as mock:
            mock.return_value = mock_response
            response = client.get("/api/v1/turf/manager-stats")
            data = response.json()

            variety = data["by_variety"][0]
            assert "variety" in variety
            assert "sqm_sold" in variety
            assert "sell_price" in variety
            assert "cost" in variety

    def test_delivery_fees_schema(self, client, mock_response):
        """Test delivery fees schema in response."""
        with patch('app.api.v1.routes.turf_manager.turf_manager_service.get_manager_stats') as mock:
            mock.return_value = mock_response
            response = client.get("/api/v1/turf/manager-stats")
            data = response.json()

            fees = data["delivery_fees"]
            assert "truck_1" in fees
            assert "truck_2" in fees
            assert "total" in fees

    def test_laying_schema(self, client, mock_response):
        """Test laying schema in response."""
        with patch('app.api.v1.routes.turf_manager.turf_manager_service.get_manager_stats') as mock:
            mock.return_value = mock_response
            response = client.get("/api/v1/turf/manager-stats")
            data = response.json()

            laying = data["laying"]
            assert "sales" in laying
            assert "costs" in laying

    def test_totals_schema(self, client, mock_response):
        """Test totals schema in response."""
        with patch('app.api.v1.routes.turf_manager.turf_manager_service.get_manager_stats') as mock:
            mock.return_value = mock_response
            response = client.get("/api/v1/turf/manager-stats")
            data = response.json()

            totals = data["totals"]
            assert "sales" in totals
            assert "costs" in totals
            assert "margin_percent" in totals


class TestWeeksEndpoint:
    """Tests for GET /api/v1/turf/weeks endpoint."""

    def test_endpoint_exists(self, client):
        """Test that weeks endpoint exists."""
        response = client.get("/api/v1/turf/weeks")
        assert response.status_code != 404

    def test_response_has_weeks(self, client):
        """Test response contains weeks array."""
        with patch('app.api.v1.routes.turf_manager.turf_manager_service.get_available_weeks') as mock:
            mock.return_value = ["Jan-05", "Jan-12", "Dec-29"]
            response = client.get("/api/v1/turf/weeks")
            data = response.json()

            assert "success" in data
            assert "weeks" in data
            assert isinstance(data["weeks"], list)

    def test_weeks_format(self, client):
        """Test weeks are in correct format."""
        with patch('app.api.v1.routes.turf_manager.turf_manager_service.get_available_weeks') as mock:
            mock.return_value = ["Jan-05", "Jan-12", "Dec-29"]
            response = client.get("/api/v1/turf/weeks")
            data = response.json()

            for week in data["weeks"]:
                # Format should be Mon-DD
                assert len(week) == 6
                assert week[3] == "-"


class TestAPIDocumentation:
    """Tests for API documentation."""

    def test_openapi_schema_available(self, client):
        """Test OpenAPI schema is accessible."""
        response = client.get("/openapi.json")
        assert response.status_code == 200
        schema = response.json()
        assert "paths" in schema

    def test_turf_endpoints_in_schema(self, client):
        """Test turf endpoints are documented."""
        response = client.get("/openapi.json")
        schema = response.json()

        assert "/api/v1/turf/manager-stats" in schema["paths"]
        assert "/api/v1/turf/weeks" in schema["paths"]

    def test_docs_endpoint_available(self, client):
        """Test Swagger UI is available."""
        response = client.get("/docs")
        assert response.status_code == 200


class TestRouterTags:
    """Tests for router configuration."""

    def test_turf_manager_tag(self, client):
        """Test endpoints have Turf Manager tag."""
        response = client.get("/openapi.json")
        schema = response.json()

        manager_stats = schema["paths"]["/api/v1/turf/manager-stats"]
        assert "Turf Manager" in manager_stats["get"]["tags"]

    def test_weeks_endpoint_tag(self, client):
        """Test weeks endpoint has correct tag."""
        response = client.get("/openapi.json")
        schema = response.json()

        weeks = schema["paths"]["/api/v1/turf/weeks"]
        assert "Turf Manager" in weeks["get"]["tags"]
