"""
Pytest configuration and fixtures for GLC Dashboard tests.
"""
import pytest
import sys
import os

# Add backend to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from fastapi.testclient import TestClient
from app.main import app


@pytest.fixture
def client():
    """Create a test client for the FastAPI app."""
    return TestClient(app)


@pytest.fixture
def sample_week_data():
    """Sample week data simulating Google Sheets response."""
    return [
        ["Day/Slot", "Variety", "Suburb", "Service Type", "SQM", "Pallets",
         "Sell $/SQM", "Cost $/SQM", "Delivery Fee", "Laying Fee",
         "Turf Revenue", "Delivery T1", "Delivery T2", "Laying Total",
         "Total", "Cost", "Margin %"],
        ["Monday AM", "Empire Zoysia", "Brisbane", "SL", "100", "2.5",
         "12", "8.25", "150", "220",
         "1200", "150", "0", "220", "1570", "825", "47.5"],
        ["Monday PM", "Sir Walter", "Logan", "SD", "150", "3.75",
         "11", "7.50", "0", "0",
         "1650", "0", "180", "0", "1830", "1125", "38.5"],
        ["Tuesday AM", "Empire Zoysia", "Ipswich", "SL", "200", "5",
         "12", "8.25", "200", "440",
         "2400", "200", "0", "440", "3040", "1650", "45.7"],
    ]


@pytest.fixture
def empty_week_data():
    """Empty week data with only headers."""
    return [
        ["Day/Slot", "Variety", "Suburb", "Service Type", "SQM", "Pallets",
         "Sell $/SQM", "Cost $/SQM", "Delivery Fee", "Laying Fee",
         "Turf Revenue", "Delivery T1", "Delivery T2", "Laying Total",
         "Total", "Cost", "Margin %"],
    ]


@pytest.fixture
def malformed_week_data():
    """Malformed data with missing columns."""
    return [
        ["Day/Slot", "Variety", "Suburb"],
        ["Monday AM", "Empire Zoysia", "Brisbane"],
        ["Tuesday AM", "", ""],  # Empty variety
    ]
