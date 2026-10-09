"""Public exceptions; messages never include tokens or email content."""


class OutlookSkillError(Exception):
    """Base class for errors raised by the skill."""


class ConfigurationError(OutlookSkillError):
    """Invalid or unreadable configuration."""


class AuthenticationError(OutlookSkillError):
    """Microsoft authentication failed or interactive sign-in is required."""


class GraphAPIError(OutlookSkillError):
    """Graph request failed."""

    def __init__(self, message: str, status_code: int | None = None):
        super().__init__(message)
        self.status_code = status_code


class EmailParsingError(OutlookSkillError):
    """An email could not be converted to the structured model."""
