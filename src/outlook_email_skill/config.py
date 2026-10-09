"""Validated settings with environment overrides for credentials."""

import json
import math
import os
import re
from dataclasses import dataclass, field
from pathlib import Path

from .errors import ConfigurationError


@dataclass(frozen=True)
class SkillConfig:
    client_id: str = ""
    tenant_id: str = "common"
    client_secret: str = field(default="", repr=False)
    auth_mode: str = "delegated"
    mailbox: str = "me"
    folder: str = "inbox"
    page_size: int = 50
    timeout: float = 30
    max_retries: int = 3
    include_attachments: bool = True
    work_domains: tuple[str, ...] = ()
    work_keywords: tuple[str, ...] = ("project", "meeting", "invoice", "report", "deadline")
    personal_keywords: tuple[str, ...] = ("family", "birthday", "dinner", "holiday")
    urgent_keywords: tuple[str, ...] = ("urgent", "asap", "immediately", "overdue")

    def __post_init__(self):
        if self.auth_mode not in {"delegated", "application"}:
            raise ConfigurationError("auth_mode must be delegated or application.")
        if not isinstance(self.tenant_id, str) or not re.fullmatch(
            r"[A-Za-z0-9.-]+", self.tenant_id
        ):
            raise ConfigurationError("tenant_id must be a tenant ID or domain.")
        for name in ("client_id", "client_secret", "mailbox", "folder"):
            if not isinstance(getattr(self, name), str):
                raise ConfigurationError(f"{name} must be a string.")
        if not self.mailbox or not self.folder:
            raise ConfigurationError("mailbox and folder must not be empty.")
        if type(self.page_size) is not int or not 1 <= self.page_size <= 1000:
            raise ConfigurationError("page_size must be an integer between 1 and 1000.")
        if (
            type(self.timeout) not in (int, float)
            or not math.isfinite(self.timeout)
            or self.timeout <= 0
        ):
            raise ConfigurationError("timeout must be a finite positive number.")
        if type(self.max_retries) is not int or not 0 <= self.max_retries <= 10:
            raise ConfigurationError("max_retries must be between 0 and 10.")
        if type(self.include_attachments) is not bool:
            raise ConfigurationError("include_attachments must be a boolean.")
        for name in ("work_domains", "work_keywords", "personal_keywords", "urgent_keywords"):
            values = getattr(self, name)
            if not isinstance(values, (tuple, list)) or not all(
                isinstance(value, str) and value.strip() for value in values
            ):
                raise ConfigurationError(f"{name} must be a list of nonempty strings.")
            object.__setattr__(self, name, tuple(value.lower().strip() for value in values))
        if self.auth_mode == "application":
            if self.mailbox == "me":
                raise ConfigurationError("Application authentication requires a mailbox.")
            if self.tenant_id in {"common", "consumers", "organizations"}:
                raise ConfigurationError("Application authentication requires a specific tenant.")

    def validate_credentials(self):
        if not self.client_id.strip():
            raise ConfigurationError("Set OUTLOOK_CLIENT_ID or client_id before connecting.")
        if self.auth_mode == "application" and not self.client_secret:
            raise ConfigurationError("Set OUTLOOK_CLIENT_SECRET before connecting.")

    @classmethod
    def from_file(cls, path: str | Path) -> "SkillConfig":
        try:
            with Path(path).open(encoding="utf-8") as stream:
                settings = json.load(stream)
        except (OSError, ValueError):
            raise ConfigurationError("Cannot read the configuration JSON.") from None
        if not isinstance(settings, dict):
            raise ConfigurationError("Configuration must be a JSON object.")
        for key in ("client_id", "tenant_id", "client_secret", "mailbox", "auth_mode"):
            value = os.getenv(f"OUTLOOK_{key.upper()}")
            if value is not None:
                settings[key] = value
        try:
            return cls(**settings)
        except TypeError:
            raise ConfigurationError("Configuration contains unknown or invalid settings.") from None
