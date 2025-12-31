"""Token manager for Microsoft Graph API authentication using MSAL."""
import logging
from typing import Optional

from msal import ConfidentialClientApplication

from app.config import settings
from app.services.graph.exceptions import GraphAuthenticationError

logger = logging.getLogger(__name__)

# Microsoft Graph API scope for client credentials flow
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]


class TokenManager:
    """Manages OAuth2 access tokens for Microsoft Graph API.

    Uses MSAL ConfidentialClientApplication for client credentials flow.
    Tokens are cached automatically by MSAL and re-acquired when expired.
    """

    _instance: Optional["TokenManager"] = None
    _app: Optional[ConfidentialClientApplication] = None

    def __new__(cls) -> "TokenManager":
        """Singleton pattern to reuse MSAL app instance."""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        """Initialize the MSAL confidential client application."""
        if self._app is not None:
            return

        if not settings.has_azure_credentials:
            logger.warning("Azure AD credentials not configured. Token manager will not function.")
            return

        authority = f"https://login.microsoftonline.com/{settings.azure_tenant_id}"

        try:
            self._app = ConfidentialClientApplication(
                client_id=settings.azure_client_id,
                client_credential=settings.azure_client_secret.get_secret_value(),
                authority=authority,
            )
            logger.info("MSAL ConfidentialClientApplication initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize MSAL application: {e}")
            raise GraphAuthenticationError(
                message=f"Failed to initialize authentication: {e}",
                details={"error": str(e)},
            )

    def get_access_token(self) -> str:
        """Acquire an access token for Microsoft Graph API.

        Uses MSAL's built-in token cache. If a valid cached token exists,
        it will be returned. Otherwise, a new token is acquired.

        Returns:
            str: The access token for Graph API calls.

        Raises:
            GraphAuthenticationError: If token acquisition fails.
        """
        if self._app is None:
            raise GraphAuthenticationError(
                message="Token manager not initialized. Check Azure AD credentials.",
                details={"hint": "Ensure AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET are set."},
            )

        # Try to get token from cache first
        result = self._app.acquire_token_silent(GRAPH_SCOPE, account=None)

        if result and "access_token" in result:
            logger.debug("Using cached access token")
            return result["access_token"]

        # No cached token, acquire a new one
        logger.info("Acquiring new access token from Azure AD")
        result = self._app.acquire_token_for_client(scopes=GRAPH_SCOPE)

        if "access_token" in result:
            logger.info("Successfully acquired new access token")
            return result["access_token"]

        # Token acquisition failed
        error = result.get("error", "unknown_error")
        error_description = result.get("error_description", "No error description provided")

        logger.error(f"Token acquisition failed: {error} - {error_description}")

        raise GraphAuthenticationError(
            message=f"Failed to acquire access token: {error}",
            details={
                "error": error,
                "error_description": error_description,
                "correlation_id": result.get("correlation_id"),
            },
        )

    def clear_cache(self) -> None:
        """Clear the MSAL token cache.

        Useful for testing or when credentials have changed.
        """
        if self._app is not None:
            # MSAL doesn't have a direct clear cache method for in-memory cache
            # Reinitialize the app to clear the cache
            self.__class__._app = None
            self.__class__._instance = None
            logger.info("Token cache cleared by reinitializing MSAL app")
