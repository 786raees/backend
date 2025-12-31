"""Custom exceptions for Google Sheets integration."""


class GoogleSheetsError(Exception):
    """Base exception for Google Sheets errors."""

    def __init__(self, message: str, details: dict = None):
        self.message = message
        self.details = details or {}
        super().__init__(self.message)


class GoogleSheetsAuthError(GoogleSheetsError):
    """Raised when authentication fails."""

    pass


class GoogleSheetsNotFoundError(GoogleSheetsError):
    """Raised when a worksheet or spreadsheet is not found."""

    pass


class GoogleSheetsAPIError(GoogleSheetsError):
    """Raised when the Google Sheets API returns an error."""

    def __init__(self, message: str, status_code: int = None, details: dict = None):
        self.status_code = status_code
        super().__init__(message, details)
