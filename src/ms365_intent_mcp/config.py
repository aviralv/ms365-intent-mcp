"""Configuration via pydantic-settings."""

from pathlib import Path

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


class Config(BaseSettings):
    model_config = SettingsConfigDict(
        env_prefix="MS365_INTENT_",
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    client_id: str = Field(
        default="14d82eec-204b-4c2f-b7e8-296a70dab67e",
        description="Microsoft public client ID",
    )
    token_path: Path = Field(
        default=Path.home() / ".config" / "ms365-intent-mcp" / "token.json",
        description="Path to stored OAuth tokens",
    )
    graph_base_url: str = Field(
        default="https://graph.microsoft.com/v1.0",
        description="Microsoft Graph API base URL",
    )
    default_timezone: str = Field(
        default="UTC",
        description="Default timezone for calendar operations. Override with MS365_INTENT_DEFAULT_TIMEZONE.",
    )
    scopes: list[str] = Field(
        default=[
            "Calendars.ReadWrite",
            "Mail.Read",
            "Mail.ReadWrite",
            "Tasks.Read",
            "Files.Read",
            "Contacts.Read",
            "User.Read",
            "Chat.ReadWrite",
            "ChatMessage.Send",
            "Channel.ReadBasic.All",
            "Team.ReadBasic.All",
            "Sites.Read.All",
        ],
        description="Microsoft Graph API scopes",
    )
    cb_failure_threshold: int = Field(default=5)
    cb_recovery_timeout: float = Field(default=60.0)
