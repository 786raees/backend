"""OneDrive service facade for fetching and caching Excel data."""
import logging
from datetime import datetime, timezone
from typing import Dict, List, Optional, Tuple

from app.config import settings
from app.models.schedule import Delivery, TruckData
from app.services.graph.graph_client import GraphClient
from app.services.excel.excel_parser import ExcelParser, ExcelDeliveryRow
from app.services.graph.exceptions import GraphAPIError

logger = logging.getLogger(__name__)

# Truck capacities in SQM
TRUCK1_CAPACITY = 500
TRUCK2_CAPACITY = 600


class CacheEntry:
    """Simple cache entry with TTL support."""

    def __init__(self, data: any, ttl_seconds: int):
        self.data = data
        self.created_at = datetime.now(timezone.utc)
        self.ttl_seconds = ttl_seconds

    def is_expired(self) -> bool:
        """Check if the cache entry has expired."""
        elapsed = (datetime.now(timezone.utc) - self.created_at).total_seconds()
        return elapsed >= self.ttl_seconds


class OneDriveService:
    """Service for fetching delivery data from OneDrive Excel.

    Combines token management, Graph API client, and Excel parsing
    into a unified interface with caching support.
    """

    _instance: Optional["OneDriveService"] = None
    _cache: Dict[str, CacheEntry] = {}

    def __new__(cls) -> "OneDriveService":
        """Singleton pattern for service reuse."""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._graph_client = None
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        """Initialize the OneDrive service (lazy - Graph client created on first use)."""
        pass

    def _ensure_client(self) -> GraphClient:
        """Lazy initialization of Graph client."""
        if self._graph_client is None:
            if not settings.has_azure_credentials:
                raise GraphAPIError("Azure credentials not configured")
            self._graph_client = GraphClient()
        return self._graph_client

    async def close(self) -> None:
        """Close the underlying HTTP client."""
        if self._graph_client:
            await self._graph_client.close()

    def _get_cache_key(self, worksheet: str) -> str:
        """Generate a cache key for a worksheet."""
        return f"worksheet:{worksheet}"

    def _get_from_cache(self, key: str) -> Optional[any]:
        """Get data from cache if not expired."""
        entry = self._cache.get(key)
        if entry and not entry.is_expired():
            logger.debug(f"Cache hit for {key}")
            return entry.data
        if entry:
            logger.debug(f"Cache expired for {key}")
            del self._cache[key]
        return None

    def _set_cache(self, key: str, data: any) -> None:
        """Store data in cache with TTL."""
        self._cache[key] = CacheEntry(
            data=data,
            ttl_seconds=settings.graph_cache_ttl_seconds,
        )
        logger.debug(f"Cached data for {key} with TTL {settings.graph_cache_ttl_seconds}s")

    def clear_cache(self) -> None:
        """Clear all cached data."""
        self._cache.clear()
        logger.info("OneDrive service cache cleared")

    async def get_worksheet_deliveries(self, worksheet_name: str) -> List[ExcelDeliveryRow]:
        """Fetch and parse deliveries from a worksheet.

        Args:
            worksheet_name: Name of the worksheet (e.g., "Truck 1").

        Returns:
            List of parsed ExcelDeliveryRow objects.

        Raises:
            GraphAPIError: If fetching data fails.
        """
        cache_key = self._get_cache_key(worksheet_name)

        # Check cache first
        cached = self._get_from_cache(cache_key)
        if cached is not None:
            return cached

        # Fetch from Graph API
        logger.info(f"Fetching worksheet data for '{worksheet_name}'")
        client = self._ensure_client()
        values = await client.get_worksheet_data(worksheet_name)

        # Parse the data
        parsed_rows = ExcelParser.parse_rows(values)

        # Cache the result
        self._set_cache(cache_key, parsed_rows)

        return parsed_rows

    async def get_truck_data(self, truck_number: int) -> TruckData:
        """Get truck data with deliveries and totals.

        Args:
            truck_number: 1 or 2 for Truck 1 or Truck 2.

        Returns:
            TruckData with deliveries and calculated totals.

        Raises:
            GraphAPIError: If fetching data fails.
            ValueError: If truck_number is not 1 or 2.
        """
        if truck_number not in (1, 2):
            raise ValueError(f"Invalid truck number: {truck_number}. Must be 1 or 2.")

        worksheet_name = f"Truck {truck_number}"
        capacity = TRUCK1_CAPACITY if truck_number == 1 else TRUCK2_CAPACITY

        try:
            parsed_rows = await self.get_worksheet_deliveries(worksheet_name)

            # Convert to Delivery models
            deliveries = [row.to_delivery() for row in parsed_rows]

            # Calculate totals
            sqm_total = sum(d.sqm for d in deliveries)
            pallet_total = sum(d.pallets for d in deliveries)
            available_sqm = max(0, capacity - sqm_total)

            return TruckData(
                deliveries=deliveries,
                sqm_total=sqm_total,
                pallet_total=pallet_total,
                capacity=capacity,
                available_sqm=available_sqm,
            )

        except GraphAPIError:
            raise
        except Exception as e:
            logger.error(f"Error getting truck data: {e}")
            raise

    async def get_both_trucks(self) -> Tuple[TruckData, TruckData]:
        """Get data for both trucks.

        Returns:
            Tuple of (truck1_data, truck2_data).

        Raises:
            GraphAPIError: If fetching data fails.
        """
        truck1 = await self.get_truck_data(1)
        truck2 = await self.get_truck_data(2)
        return truck1, truck2

    def is_available(self) -> bool:
        """Check if the OneDrive service is available.

        Returns True if Azure credentials are configured.
        """
        return settings.has_azure_credentials


# Global service instance
onedrive_service = OneDriveService()
