"""Pydantic models for Turf Manager dashboard data."""
from pydantic import BaseModel, Field
from typing import List, Literal, Optional


class VarietyStats(BaseModel):
    """Statistics for a single turf variety."""
    variety: str = Field(..., description="Turf variety name")
    sqm_sold: float = Field(default=0, ge=0, description="Total SQM sold")
    sell_price: float = Field(default=0, ge=0, description="Total sell price ($)")
    cost: float = Field(default=0, ge=0, description="Total cost ($)")


class VarietyTotals(BaseModel):
    """Aggregate totals across all varieties."""
    sqm_sold: float = Field(default=0, ge=0, description="Total SQM sold across all varieties")
    sell_price: float = Field(default=0, ge=0, description="Total sell price ($)")
    cost: float = Field(default=0, ge=0, description="Total cost ($)")


class DeliveryFees(BaseModel):
    """Delivery fee breakdown by truck."""
    truck_1: float = Field(default=0, ge=0, description="Truck 1 delivery fees ($)")
    truck_2: float = Field(default=0, ge=0, description="Truck 2 delivery fees ($)")
    total: float = Field(default=0, ge=0, description="Total delivery fees ($)")


class LayingStats(BaseModel):
    """Laying fees and costs."""
    sales: float = Field(default=0, ge=0, description="Laying fees charged ($)")
    costs: float = Field(default=0, ge=0, description="Laying costs ($2.20/SQM for all deliveries)")


class FinancialTotals(BaseModel):
    """Overall financial totals."""
    sales: float = Field(default=0, ge=0, description="Total sales ($)")
    costs: float = Field(default=0, ge=0, description="Total costs ($)")
    margin_percent: float = Field(default=0, description="Margin percentage")


class TurfManagerResponse(BaseModel):
    """Complete API response for Turf Manager dashboard."""
    success: bool = True
    view: Literal["day", "week", "month", "annual"] = Field(..., description="View type")
    period: str = Field(..., description="Human-readable period description")
    by_variety: List[VarietyStats] = Field(default_factory=list, description="Breakdown by turf variety")
    variety_totals: VarietyTotals = Field(default_factory=VarietyTotals, description="Variety aggregates")
    delivery_fees: DeliveryFees = Field(default_factory=DeliveryFees, description="Delivery fee breakdown")
    laying: LayingStats = Field(default_factory=LayingStats, description="Laying fees and costs")
    totals: FinancialTotals = Field(default_factory=FinancialTotals, description="Overall financial totals")
