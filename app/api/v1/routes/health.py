"""Health check endpoint."""
from fastapi import APIRouter
from datetime import datetime, timezone

router = APIRouter()


@router.get("/health")
async def health_check():
    """Return service health status."""
    return {
        "status": "healthy",
        "service": "glc-dashboard-api",
        "version": "1.0.0",
        "timestamp": datetime.now(timezone.utc).isoformat()
    }
