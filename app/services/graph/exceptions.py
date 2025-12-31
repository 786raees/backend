"""Custom exceptions for Microsoft Graph API operations."""


class GraphAPIError(Exception):
    """Base exception for all Graph API errors."""

    def __init__(self, message: str, status_code: int = None, details: dict = None):
        self.message = message
        self.status_code = status_code
        self.details = details or {}
        super().__init__(self.message)


class GraphAuthenticationError(GraphAPIError):
    """Raised when authentication fails (HTTP 401).

    Troubleshooting:
    - Verify AZURE_TENANT_ID is correct
    - Verify AZURE_CLIENT_ID matches your app registration
    - Verify AZURE_CLIENT_SECRET is valid and not expired
    - Check if the app registration exists in Azure AD
    """

    def __init__(self, message: str = None, details: dict = None):
        super().__init__(
            message or "Authentication failed. Check Azure AD credentials.",
            status_code=401,
            details=details,
        )


class GraphPermissionError(GraphAPIError):
    """Raised when access is forbidden (HTTP 403).

    Troubleshooting:
    - Ensure Files.Read.All permission is granted
    - Verify admin consent has been granted for the permission
    - Check if the app has access to the specific OneDrive/SharePoint site
    - For personal OneDrive, you may need different permissions
    """

    def __init__(self, message: str = None, details: dict = None):
        super().__init__(
            message or "Access forbidden. Ensure Files.Read.All permission is granted with admin consent.",
            status_code=403,
            details=details,
        )


class GraphResourceNotFoundError(GraphAPIError):
    """Raised when a resource is not found (HTTP 404).

    Troubleshooting:
    - Verify ONEDRIVE_DRIVE_ID is correct
    - Verify ONEDRIVE_ITEM_ID matches the Excel file
    - Check if the worksheet name exists in the Excel file
    - Ensure the file hasn't been moved or deleted
    """

    def __init__(self, resource: str = None, message: str = None, details: dict = None):
        self.resource = resource
        super().__init__(
            message or f"Resource not found: {resource or 'Unknown'}. Verify drive ID, item ID, and worksheet name.",
            status_code=404,
            details=details,
        )


class GraphRateLimitError(GraphAPIError):
    """Raised when rate limited by Microsoft Graph (HTTP 429).

    Troubleshooting:
    - Implement exponential backoff retry logic
    - Increase cache TTL to reduce API calls
    - Check the Retry-After header for wait time
    """

    def __init__(self, retry_after: int = None, message: str = None, details: dict = None):
        self.retry_after = retry_after
        super().__init__(
            message or f"Rate limited. Retry after {retry_after or 'unknown'} seconds.",
            status_code=429,
            details=details,
        )


class GraphServiceError(GraphAPIError):
    """Raised for server-side errors (HTTP 5xx).

    Troubleshooting:
    - Microsoft Graph service may be experiencing issues
    - Retry the request after a short delay
    - Check Microsoft 365 Service Health dashboard
    """

    def __init__(self, message: str = None, status_code: int = 500, details: dict = None):
        super().__init__(
            message or "Microsoft Graph service error. Please retry later.",
            status_code=status_code,
            details=details,
        )
