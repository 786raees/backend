"""
Sales API Routes - Endpoints for sales appointment tracking.

Endpoints:
- GET /schedule - Daily appointments for staff dashboard
- POST /appointment/update - Update single appointment field
- GET /stats - Weekly statistics for manager dashboard
- GET /ceo-summary - Summary metrics for CEO dashboard
"""

from datetime import date
from typing import Optional

from fastapi import APIRouter, HTTPException, Query
from pydantic import BaseModel, Field

from app.services.sales_service import sales_service

router = APIRouter(prefix="/sales", tags=["Sales"])


# ============ Request/Response Models ============

class AppointmentUpdateRequest(BaseModel):
    """Request body for updating an appointment field."""
    week_tab: str = Field(..., description="Week tab name, e.g., 'Jan-05'")
    row_number: int = Field(..., ge=1, description="Row number in sheet (1-indexed)")
    column: str = Field(..., pattern="^[CFGHJKLMNOPQR]$", description="Column letter (C, F, G, H, J, K, L, M, N, O, P, Q, or R)")
    value: str = Field(..., description="Value to set (e.g., 'Yes', '', reason text, price, time, or project type)")


class AppointmentUpdateResponse(BaseModel):
    """Response for appointment update."""
    success: bool
    updated: Optional[dict] = None
    error: Optional[str] = None


# ============ Endpoints ============

@router.get("/schedule")
async def get_daily_schedule(
    date_param: Optional[str] = Query(None, alias="date", description="Date in YYYY-MM-DD format")
):
    """
    Get all appointments for a specific date (Staff Dashboard).

    Query Parameters:
    - date: YYYY-MM-DD format. Defaults to today.

    Returns appointments grouped by sales rep with status toggles.
    Weekend dates will return an error.
    """
    # Parse date parameter
    if date_param:
        try:
            target_date = date.fromisoformat(date_param)
        except ValueError:
            raise HTTPException(status_code=400, detail="Invalid date format. Use YYYY-MM-DD")
    else:
        target_date = date.today()

    # Check for Sunday only (Saturday has 3 slots per rep)
    if target_date.weekday() == 6:  # Sunday only
        raise HTTPException(status_code=400, detail="No appointments on Sundays")

    result = await sales_service.get_daily_schedule(target_date)

    if not result.get("success"):
        raise HTTPException(status_code=500, detail=result.get("error", "Failed to fetch schedule"))

    return result


@router.post("/appointment/update", response_model=AppointmentUpdateResponse)
async def update_appointment(request: AppointmentUpdateRequest):
    """
    Update a single field in an appointment row (Staff Dashboard writes back).

    Editable columns:
    - C (Lead Source): Lead source name
    - F (Attended): "Yes" or ""
    - G (Sold): "Yes" or ""
    - H (Reason): Reason text
    - J (Sell Price): Price value (e.g., "50000" or "$50,000")
    - K (Appointment Time): Time string (e.g., "10:00 AM")
    - L (Project Type): Project type (e.g., "Turf Install", "Landscaping")
    - M (Suburb): Suburb name
    - N (Region): Region name
    - O (Appointment Set Who): Staff name
    - P (Appointment Confirmed By): Staff name
    - Q (Gross Profit Margin %): Percentage value
    - R (Paid/Unpaid): Payment status
    """
    result = await sales_service.update_appointment_field(
        week_tab=request.week_tab,
        row_number=request.row_number,
        column=request.column,
        value=request.value
    )

    if not result.get("success"):
        raise HTTPException(status_code=400, detail=result.get("error", "Failed to update"))

    return result


@router.get("/stats")
async def get_stats(
    view: str = Query("week", description="View type: day, week, month, annual"),
    week: Optional[str] = Query(None, description="Week tab name for week view, e.g., 'Jan-05'"),
    date_param: Optional[str] = Query(None, alias="date", description="Date for day view, YYYY-MM-DD"),
    month: Optional[str] = Query(None, description="Month for month view, e.g., '2026-01'"),
    year: Optional[str] = Query(None, description="Year for annual view, e.g., '2026'")
):
    """
    Get aggregated statistics (Manager Dashboard).

    Query Parameters:
    - view: View type - 'day', 'week', 'month', or 'annual'
    - week: Week tab name (required for week view)
    - date: YYYY-MM-DD format (for day view)
    - month: YYYY-MM format (for month view)
    - year: YYYY format (for annual view)

    Returns totals, breakdown by rep, lead source, and day.
    """
    if view == "day":
        if date_param:
            try:
                target_date = date.fromisoformat(date_param)
            except ValueError:
                raise HTTPException(status_code=400, detail="Invalid date format. Use YYYY-MM-DD")
        else:
            target_date = date.today()

        result = await sales_service.get_daily_stats(target_date)

    elif view == "week":
        if not week:
            week = sales_service.get_week_tab_name(date.today())
        result = await sales_service.get_weekly_stats(week)

    elif view == "month":
        if not month:
            month = date.today().strftime("%Y-%m")
        result = await sales_service.get_monthly_stats(month)

    elif view == "annual":
        if not year:
            year = str(date.today().year)
        result = await sales_service.get_annual_stats(year)

    else:
        raise HTTPException(status_code=400, detail=f"Invalid view type: {view}")

    if not result.get("success"):
        raise HTTPException(status_code=500, detail=result.get("error", "Failed to fetch stats"))

    return result


@router.get("/ceo-summary")
async def get_ceo_summary(
    week: Optional[str] = Query(None, description="Week tab name. Defaults to current week.")
):
    """
    Get summary metrics for CEO dashboard section.

    Query Parameters:
    - week (optional): Week tab name. Defaults to current week.

    Returns high-level metrics for executive view.
    """
    result = await sales_service.get_ceo_summary(week)

    if not result.get("success"):
        raise HTTPException(status_code=500, detail=result.get("error", "Failed to fetch summary"))

    return result


@router.get("/weeks")
async def get_available_weeks():
    """
    Get list of available week tabs.

    Returns list of week tab names (e.g., ["Jan-05", "Jan-12", "Dec-29"]).
    Useful for populating week selector dropdowns.
    """
    weeks = await sales_service.get_available_weeks()
    return {"success": True, "weeks": weeks}
