"""Schedule API endpoint."""
import logging
from typing import Optional

from fastapi import APIRouter, HTTPException, Query

from app.config import settings
from app.models.schedule import ScheduleResponse
from app.services.mock_data import generate_mock_schedule
from app.services.onedrive_service import onedrive_service
from app.services.google_sheets_service import google_sheets_service
from app.services.schedule_builder import schedule_builder
from app.services.graph.exceptions import GraphAPIError
from app.services.google.exceptions import GoogleSheetsError

logger = logging.getLogger(__name__)
router = APIRouter()


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
