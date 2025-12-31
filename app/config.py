"""Configuration management using pydantic-settings."""
import json
from pydantic_settings import BaseSettings
from pydantic import SecretStr
from typing import Any, Dict, List, Optional


class Settings(BaseSettings):
    """Application settings loaded from environment variables."""

    # Server Settings
    api_host: str = "0.0.0.0"
    api_port: int = 8000
    debug: bool = True

    # CORS Settings
    cors_origins: str = "*"

    # Azure AD / Microsoft Graph API (for OneDrive integration - legacy)
    azure_tenant_id: str = ""
    azure_client_id: str = ""
    azure_client_secret: SecretStr = SecretStr("")

    # OneDrive Configuration (legacy)
    onedrive_drive_id: str = ""
    onedrive_item_id: str = ""

    # Google Sheets Configuration (primary data source)
    google_spreadsheet_id: str = ""
    google_service_account_json: SecretStr = SecretStr("")

    # Cache Configuration
    graph_cache_ttl_seconds: int = 300  # 5 minutes

    # Feature Flags
    use_mock_data_fallback: bool = True

    @property
    def cors_origins_list(self) -> List[str]:
        """Parse CORS origins from comma-separated string."""
        if self.cors_origins == "*":
            return ["*"]
        return [origin.strip() for origin in self.cors_origins.split(",")]

    @property
    def has_azure_credentials(self) -> bool:
        """Check if Azure credentials are configured."""
        return bool(
            self.azure_tenant_id
            and self.azure_client_id
            and self.azure_client_secret.get_secret_value()
            and self.onedrive_drive_id
            and self.onedrive_item_id
        )

    @property
    def has_google_credentials(self) -> bool:
        """Check if Google Sheets credentials are configured."""
        return bool(
            self.google_spreadsheet_id
            and self.google_service_account_json.get_secret_value()
        )

    @property
    def google_service_account_info(self) -> Optional[Dict[str, Any]]:
        """Parse Google service account JSON into a dictionary.

        Returns:
            Parsed service account info dict, or None if not configured.
        """
        json_str = self.google_service_account_json.get_secret_value()
        if not json_str:
            return None
        try:
            return json.loads(json_str)
        except json.JSONDecodeError:
            return None

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"


# Global settings instance
settings = Settings()
