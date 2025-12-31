"""Microsoft Graph API client for OneDrive/Excel operations."""
import logging
from typing import Any, Dict, List, Optional

import httpx

from app.config import settings
from app.services.auth.token_manager import TokenManager
from app.services.graph.exceptions import (
    GraphAuthenticationError,
    GraphPermissionError,
    GraphResourceNotFoundError,
    GraphRateLimitError,
    GraphServiceError,
)

logger = logging.getLogger(__name__)

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


class GraphClient:
    """Async HTTP client for Microsoft Graph API operations.

    Handles authentication, request building, and error handling
    for Graph API calls to OneDrive/Excel resources.
    """

    def __init__(self):
        """Initialize the Graph client with token manager."""
        self._token_manager = TokenManager()
        self._http_client: Optional[httpx.AsyncClient] = None

    async def _get_client(self) -> httpx.AsyncClient:
        """Get or create the async HTTP client."""
        if self._http_client is None or self._http_client.is_closed:
            self._http_client = httpx.AsyncClient(
                base_url=GRAPH_BASE_URL,
                timeout=30.0,
            )
        return self._http_client

    async def close(self) -> None:
        """Close the HTTP client connection."""
        if self._http_client and not self._http_client.is_closed:
            await self._http_client.aclose()
            self._http_client = None

    def _get_headers(self) -> Dict[str, str]:
        """Get request headers with authorization token."""
        token = self._token_manager.get_access_token()
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

    def _handle_error_response(self, response: httpx.Response, resource: str = None) -> None:
        """Handle error responses from Graph API.

        Args:
            response: The HTTP response object.
            resource: Optional resource name for error context.

        Raises:
            GraphAuthenticationError: For 401 responses.
            GraphPermissionError: For 403 responses.
            GraphResourceNotFoundError: For 404 responses.
            GraphRateLimitError: For 429 responses.
            GraphServiceError: For 5xx responses.
        """
        status_code = response.status_code

        try:
            error_body = response.json()
            error_info = error_body.get("error", {})
            error_code = error_info.get("code", "unknown")
            error_message = error_info.get("message", response.text)
        except Exception:
            error_code = "parse_error"
            error_message = response.text

        details = {
            "status_code": status_code,
            "error_code": error_code,
            "resource": resource,
        }

        logger.error(f"Graph API error: {status_code} - {error_code}: {error_message}")

        if status_code == 401:
            raise GraphAuthenticationError(
                message=f"Authentication failed: {error_message}",
                details=details,
            )
        elif status_code == 403:
            raise GraphPermissionError(
                message=f"Access denied: {error_message}",
                details=details,
            )
        elif status_code == 404:
            raise GraphResourceNotFoundError(
                resource=resource,
                message=f"Resource not found: {error_message}",
                details=details,
            )
        elif status_code == 429:
            retry_after = int(response.headers.get("Retry-After", 60))
            raise GraphRateLimitError(
                retry_after=retry_after,
                message=f"Rate limited. Retry after {retry_after} seconds.",
                details=details,
            )
        elif status_code >= 500:
            raise GraphServiceError(
                message=f"Graph service error: {error_message}",
                status_code=status_code,
                details=details,
            )

    async def get_worksheet_data(self, worksheet_name: str) -> List[List[Any]]:
        """Fetch worksheet data using the usedRange endpoint.

        Args:
            worksheet_name: Name of the worksheet (e.g., "Truck 1", "Truck 2").

        Returns:
            List of rows, where each row is a list of cell values.

        Raises:
            GraphAPIError: If the API call fails.
        """
        drive_id = settings.onedrive_drive_id
        item_id = settings.onedrive_item_id

        # URL-encode the worksheet name
        encoded_name = httpx.URL(worksheet_name).raw_path.decode() if "/" in worksheet_name else worksheet_name

        # Build the usedRange endpoint URL
        endpoint = f"/drives/{drive_id}/items/{item_id}/workbook/worksheets/{worksheet_name}/usedRange"

        logger.info(f"Fetching worksheet data: {worksheet_name}")

        client = await self._get_client()

        try:
            response = await client.get(
                endpoint,
                headers=self._get_headers(),
            )
        except httpx.RequestError as e:
            logger.error(f"HTTP request failed: {e}")
            raise GraphServiceError(
                message=f"Failed to connect to Microsoft Graph: {e}",
                details={"error": str(e)},
            )

        if response.status_code != 200:
            self._handle_error_response(response, resource=f"worksheet:{worksheet_name}")

        data = response.json()
        values = data.get("values", [])

        logger.info(f"Retrieved {len(values)} rows from worksheet '{worksheet_name}'")

        return values

    async def get_workbook_info(self) -> Dict[str, Any]:
        """Get information about the workbook (for debugging/validation).

        Returns:
            Dictionary with workbook metadata.
        """
        drive_id = settings.onedrive_drive_id
        item_id = settings.onedrive_item_id

        endpoint = f"/drives/{drive_id}/items/{item_id}/workbook/worksheets"

        client = await self._get_client()

        try:
            response = await client.get(
                endpoint,
                headers=self._get_headers(),
            )
        except httpx.RequestError as e:
            raise GraphServiceError(
                message=f"Failed to connect to Microsoft Graph: {e}",
                details={"error": str(e)},
            )

        if response.status_code != 200:
            self._handle_error_response(response, resource="workbook:worksheets")

        return response.json()
