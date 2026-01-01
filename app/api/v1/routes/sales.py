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
    column: str = Field(..., pattern="^[FGH]$", description="Column letter (F, G, or H only)")
    value: str = Field(..., description="Value to set (e.g., 'Yes', '', or reason text)")


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

    # Check for weekend
    if target_date.weekday() >= 5:
        raise HTTPException(status_code=400, detail="No appointments on weekends")

    result = await sales_service.get_daily_schedule(target_date)

    if not result.get("success"):
        raise HTTPException(status_code=500, detail=result.get("error", "Failed to fetch schedule"))

    return result


@router.post("/appointment/update", response_model=AppointmentUpdateResponse)
async def update_appointment(request: AppointmentUpdateRequest):
    """
    Update a single field in an appointment row (Staff Dashboard writes back).

    Only columns F (Attended), G (Sold), and H (Reason) can be updated.
    Values should be "Yes", "", or reason text for column H.
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
async def get_weekly_stats(
    week: str = Query(..., description="Week tab name, e.g., 'Jan-05'")
):
    """
    Get aggregated weekly statistics (Manager Dashboard).

    Query Parameters:
    - week (required): Week tab name, e.g., "Jan-05"

    Returns totals, breakdown by rep, lead source, and day.
    """
    result = await sales_service.get_weekly_stats(week)

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
