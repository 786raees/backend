"""Schedule API endpoint."""
import logging
from typing import Optional

from fastapi import APIRouter, HTTPException, Query
from pydantic import BaseModel, Field

from app.config import settings
from app.models.schedule import ScheduleResponse
from app.services.mock_data import generate_mock_schedule
from app.services.onedrive_service import onedrive_service
from app.services.google_sheets_service import google_sheets_service
from app.services.schedule_builder import schedule_builder
from app.services.turf_delivery_service import turf_delivery_service
from app.services.graph.exceptions import GraphAPIError
from app.services.google.exceptions import GoogleSheetsError

logger = logging.getLogger(__name__)
router = APIRouter()


# ============ Request/Response Models for Delivery Management ============

class DeliveryUpdateRequest(BaseModel):
    """Request body for updating a delivery field."""
    week_tab: str = Field(..., description="Week tab name, e.g., 'Jan-05'")
    row_number: int = Field(..., ge=1, description="Row number in sheet (1-indexed)")
    column: str = Field(..., pattern="^[BCDEIJP]$", description="Column letter (B, C, D, E, I, J, or P)")
    value: str = Field(..., description="Value to set")


class DeliveryUpdateResponse(BaseModel):
    """Response for delivery update."""
    success: bool
    updated: Optional[dict] = None
    error: Optional[str] = None


class DeliveryCreateRequest(BaseModel):
    """Request body for creating a new delivery."""
    week_tab: str = Field(..., description="Week tab name, e.g., 'Jan-12'")
    day: str = Field(..., description="Day of week (Monday-Friday)")
    truck: str = Field(..., description="Truck name: 'TRUCK 1' or 'TRUCK 2'")
    slot: int = Field(..., ge=1, le=6, description="Slot number 1-6")
    variety: Optional[str] = Field(default="", description="Turf variety")
    suburb: Optional[str] = Field(default="", description="Suburb/location")
    service_type: Optional[str] = Field(default="", description="Service type (SL/SD/P)")
    sqm_sold: Optional[str] = Field(default="", description="Square meters sold")
    delivery_fee: Optional[str] = Field(default="", description="Delivery fee")
    laying_fee: Optional[str] = Field(default="", description="Laying fee")


class DeliveryCreateResponse(BaseModel):
    """Response for delivery creation."""
    success: bool
    message: Optional[str] = None
    row_number: Optional[int] = None
    delivery: Optional[dict] = None
    error: Optional[str] = None


class DeliveryDeleteRequest(BaseModel):
    """Request body for deleting a delivery."""
    week_tab: str = Field(..., description="Week tab name, e.g., 'Jan-12'")
    row_number: Optional[int] = Field(None, ge=1, description="Row number in sheet (1-indexed)")
    day: Optional[str] = Field(None, description="Day of week (if row_number not provided)")
    truck: Optional[str] = Field(None, description="Truck name (if row_number not provided)")
    slot: Optional[int] = Field(None, ge=1, le=6, description="Slot number (if row_number not provided)")


class DeliveryDeleteResponse(BaseModel):
    """Response for delivery deletion."""
    success: bool
    message: Optional[str] = None
    row_number: Optional[int] = None
    error: Optional[str] = None


class DeliveryMoveRequest(BaseModel):
    """Request body for moving a delivery to a new day/truck/slot."""
    week_tab: str = Field(..., description="Week tab name, e.g., 'Jan-12'")
    from_row: int = Field(..., ge=1, description="Source row number in sheet")
    to_day: str = Field(..., description="Destination day of week (Monday-Friday)")
    to_truck: str = Field(..., description="Destination truck: 'TRUCK 1' or 'TRUCK 2'")
    to_slot: int = Field(..., ge=1, le=6, description="Destination slot number 1-6")


class DeliveryMoveResponse(BaseModel):
    """Response for delivery move."""
    success: bool
    message: Optional[str] = None
    from_row: Optional[int] = None
    to_row: Optional[int] = None
    to_day: Optional[str] = None
    to_truck: Optional[str] = None
    to_slot: Optional[int] = None
    error: Optional[str] = None


@router.post("/schedule/refresh")
async def refresh_schedule():
    """Clear the cache and force a fresh fetch of schedule data.

    Use this endpoint when data has been updated in Google Sheets
    and you want to see the changes immediately without waiting
    for the cache to expire (5 minutes).

    Returns:
        Success message with cache status.
    """
    logger.info("Cache refresh requested")

    # Clear Google Sheets cache
    google_sheets_service.clear_cache()

    # Clear OneDrive cache if available
    if hasattr(onedrive_service, 'clear_cache'):
        onedrive_service.clear_cache()

    return {
        "success": True,
        "message": "Cache cleared. Next request will fetch fresh data.",
    }


@router.get("/schedule", response_model=ScheduleResponse)
async def get_schedule(
    refresh: bool = Query(False, description="Set to true to bypass cache and fetch fresh data")
):
    """Get the delivery schedule for the next 10 business days.

    Returns schedule data for Truck 1 and Truck 2 with:
    - Delivery details (SQM, pallets, variety, suburb, service type)
    - Daily totals per truck
    - Available capacity per truck

    Data source priority:
    1. Google Sheets (primary) - fetches from weekly tabs
    2. OneDrive (fallback) - legacy Excel integration
    3. Mock data (final fallback) - when USE_MOCK_DATA_FALLBACK=true

    Query Parameters:
    - refresh: Set to true to clear cache and fetch fresh data
    """
    # Clear cache if refresh requested
    if refresh:
        logger.info("Refresh requested - clearing cache")
        google_sheets_service.clear_cache()

    # Primary: Try Google Sheets
    if google_sheets_service.is_available():
        try:
            logger.info("Fetching schedule from Google Sheets")

            # Get deliveries from all required weekly tabs for each truck
            truck1_rows, truck2_rows = google_sheets_service.get_all_truck_deliveries()

            # Build the schedule using the schedule builder
            schedule = schedule_builder.build_schedule(
                truck1_deliveries=truck1_rows,
                truck2_deliveries=truck2_rows,
                source="google_sheets",
            )

            logger.info(f"Built schedule with {len(schedule.days)} days from Google Sheets")
            return schedule

        except GoogleSheetsError as e:
            logger.error(f"Google Sheets fetch failed: {e}")
            # Fall through to try OneDrive or mock data

    # Fallback: Try OneDrive
    if onedrive_service.is_available():
        try:
            logger.info("Fetching schedule from OneDrive")

            # Get raw parsed delivery rows from each worksheet
            truck1_rows = await onedrive_service.get_worksheet_deliveries("Truck 1")
            truck2_rows = await onedrive_service.get_worksheet_deliveries("Truck 2")

            # Build the schedule using the schedule builder
            schedule = schedule_builder.build_schedule(
                truck1_deliveries=truck1_rows,
                truck2_deliveries=truck2_rows,
                source="onedrive",
            )

            logger.info(f"Built schedule with {len(schedule.days)} days from OneDrive")
            return schedule

        except GraphAPIError as e:
            logger.error(f"OneDrive fetch failed: {e}")

            if settings.use_mock_data_fallback:
                logger.info("Falling back to mock data")
                return generate_mock_schedule()

            raise HTTPException(
                status_code=e.status_code or 500,
                detail={
                    "error": "Failed to fetch data from OneDrive",
                    "message": str(e),
                    "details": e.details,
                },
            )

        except Exception as e:
            logger.error(f"Unexpected error fetching schedule: {e}")

            if settings.use_mock_data_fallback:
                logger.info("Falling back to mock data due to unexpected error")
                return generate_mock_schedule()

            raise HTTPException(
                status_code=500,
                detail={
                    "error": "Internal server error",
                    "message": str(e),
                },
            )

    # Final fallback: Mock data
    if settings.use_mock_data_fallback:
        logger.info("No data source configured, using mock data")
        return generate_mock_schedule()

    raise HTTPException(
        status_code=503,
        detail={
            "error": "No data source available",
            "message": "Neither Google Sheets nor OneDrive credentials are configured",
        },
    )


# ============ Delivery Management Endpoints ============

@router.post("/delivery/update", response_model=DeliveryUpdateResponse)
async def update_delivery(request: DeliveryUpdateRequest):
    """
    Update a single field in a delivery row.

    Editable columns:
    - B (Variety): Turf variety name
    - C (Suburb): Suburb/location
    - D (Service Type): "SL", "SD", or "P"
    - E (SQM Sold): Square meters value
    - P (Payment Status): "Unpaid", "Paid", or "Partial"
    """
    result = await turf_delivery_service.update_delivery_field(
        week_tab=request.week_tab,
        row_number=request.row_number,
        column=request.column,
        value=request.value
    )

    if not result.get("success"):
        raise HTTPException(status_code=400, detail=result.get("error", "Failed to update"))

    return result


@router.post("/delivery/create", response_model=DeliveryCreateResponse, status_code=201)
async def create_delivery(request: DeliveryCreateRequest):
    """
    Create a new delivery in an empty slot.

    Required fields:
    - week_tab: Week tab name (e.g., "Jan-12")
    - day: Day of week (Monday-Friday)
    - truck: Truck name ("TRUCK 1" or "TRUCK 2")
    - slot: Slot number (1-6)

    Optional fields:
    - variety, suburb, service_type, sqm_sold

    This endpoint will:
    1. Validate all required fields
    2. Calculate the correct row number in Google Sheets
    3. Write the delivery data to the appropriate columns
    4. Return the created delivery with row number
    """
    result = await turf_delivery_service.create_delivery(
        week_tab=request.week_tab,
        day=request.day,
        truck=request.truck,
        slot=request.slot,
        variety=request.variety or "",
        suburb=request.suburb or "",
        service_type=request.service_type or "",
        sqm_sold=request.sqm_sold or "",
        delivery_fee=request.delivery_fee or "",
        laying_fee=request.laying_fee or ""
    )

    if not result.get("success"):
        error_code = result.get("error_code", 400)
        raise HTTPException(status_code=error_code, detail=result.get("error", "Failed to create delivery"))

    return result


@router.post("/delivery/delete", response_model=DeliveryDeleteResponse)
async def delete_delivery(request: DeliveryDeleteRequest):
    """
    Delete a delivery (clear all data in the row).

    Requires either:
    - row_number: Direct row number in the sheet
    - OR day + truck + slot: To calculate the row number

    This endpoint clears all delivery data from the row while preserving
    the row structure. The slot becomes available for new deliveries.
    """
    result = await turf_delivery_service.delete_delivery(
        week_tab=request.week_tab,
        row_number=request.row_number,
        day=request.day,
        truck=request.truck,
        slot=request.slot
    )

    if not result.get("success"):
        error_code = result.get("error_code", 400)
        raise HTTPException(status_code=error_code, detail=result.get("error", "Failed to delete delivery"))

    return result


@router.post("/delivery/move", response_model=DeliveryMoveResponse)
async def move_delivery(request: DeliveryMoveRequest):
    """
    Move a delivery from one day/truck/slot to another.

    This endpoint:
    1. Validates the destination slot is empty
    2. Copies all delivery data to the new location
    3. Clears the old location
    4. Clears cache to reflect changes

    Use this to reschedule deliveries when plans change.
    """
    result = await turf_delivery_service.move_delivery(
        week_tab=request.week_tab,
        from_row=request.from_row,
        to_day=request.to_day,
        to_truck=request.to_truck,
        to_slot=request.to_slot
    )

    if not result.get("success"):
        error_code = result.get("error_code", 400)
        raise HTTPException(status_code=error_code, detail=result.get("error", "Failed to move delivery"))

    return result


@router.get("/delivery/weeks")
async def get_delivery_weeks():
    """
    Get list of available week tabs.

    Returns list of week tab names (e.g., ["Jan-05", "Jan-12", "Dec-29"]).
    Useful for populating week selector dropdowns.
    """
    weeks = await turf_delivery_service.get_available_weeks()
    return {"success": True, "weeks": weeks}
