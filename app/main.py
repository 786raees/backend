"""FastAPI application entry point for GLC Dashboard API."""
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.config import settings
from app.api.v1.routes import health, schedule, sales, turf_manager

# Create FastAPI application
app = FastAPI(
    title="GLC Dashboard API",
    description="API for The Great Lawn Co. TV Delivery Dashboard",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc"
)

# Configure CORS
# Note: For production, restrict origins to glcdashboard.com.au
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins_list,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include routers
app.include_router(health.router, tags=["Health"])
app.include_router(schedule.router, prefix="/api/v1", tags=["Schedule"])
app.include_router(sales.router, prefix="/api/v1", tags=["Sales"])
app.include_router(turf_manager.router, prefix="/api/v1", tags=["Turf Manager"])


@app.get("/")
async def root():
    """Root endpoint with API information."""
    return {
        "name": "GLC Dashboard API",
        "version": "1.0.0",
        "docs": "/docs",
        "health": "/health"
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "app.main:app",
        host=settings.api_host,
        port=settings.api_port,
        reload=settings.debug
    )
