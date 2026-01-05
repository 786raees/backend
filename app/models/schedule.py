"""Pydantic models for schedule data."""
from pydantic import BaseModel, Field
from typing import List, Literal, Optional
from datetime import date


class Delivery(BaseModel):
    """Individual delivery record."""
    sqm: float = Field(..., gt=0, description="Square meters of turf")
    pallets: float = Field(..., ge=0, description="Number of pallets")
    variety: str = Field(..., description="Turf variety name")
    suburb: str = Field(..., description="Delivery suburb")
    service_type: Literal["SL", "SD", "P"] = Field(..., description="Service type code")
    laying_cost: float = Field(default=0, ge=0, description="Laying cost ($2.20/SQM for SL only)")
    payment_status: Optional[str] = Field(default=None, description="Payment status: Paid, Payment Pending, Cash")


class TruckData(BaseModel):
    """Per-truck daily data with deliveries and totals."""
    deliveries: List[Delivery] = Field(default_factory=list)
    sqm_total: float = Field(default=0, ge=0, description="Total SQM for the day")
    pallet_total: float = Field(default=0, ge=0, description="Total pallets for the day")
    laying_cost_total: float = Field(default=0, ge=0, description="Total laying cost for the day")
    capacity: int = Field(..., description="Truck capacity in SQM")
    available_sqm: float = Field(..., description="Remaining capacity (can be negative if over)")


class DaySchedule(BaseModel):
    """Complete schedule for a single day."""
    date: date
    day_name: str = Field(..., description="Formatted day name, e.g., 'Mon 22nd Dec'")
    is_week_two: bool = Field(default=False, description="Flag for week 2 styling")
    truck1: TruckData
    truck2: TruckData


class ScheduleResponse(BaseModel):
    """Complete API response for schedule endpoint."""
    success: bool = True
    generated_at: str = Field(..., description="ISO timestamp of generation")
    source: Literal["google_sheets", "onedrive", "mock"] = Field(default="mock", description="Data source indicator")
    days: List[DaySchedule] = Field(..., description="10 business days of schedule data")
