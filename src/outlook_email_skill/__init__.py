"""Outlook email reading and local, explainable analysis."""

import logging

from .config import SkillConfig
from .errors import (
    AuthenticationError,
    ConfigurationError,
    EmailParsingError,
    GraphAPIError,
    OutlookSkillError,
)
from .skill import OutlookEmailSkill

logging.getLogger(__name__).addHandler(logging.NullHandler())

__all__ = [
    "OutlookEmailSkill",
    "SkillConfig",
    "OutlookSkillError",
    "AuthenticationError",
    "ConfigurationError",
    "EmailParsingError",
    "GraphAPIError",
]
