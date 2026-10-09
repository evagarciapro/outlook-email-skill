"""Structured, serializable email and analysis results."""

from dataclasses import asdict, dataclass, field
from datetime import datetime


@dataclass(frozen=True)
class Address:
    name: str = ""
    address: str = ""


@dataclass(frozen=True)
class Attachment:
    id: str
    name: str
    content_type: str
    size: int
    is_inline: bool
    kind: str


@dataclass(frozen=True)
class Email:
    id: str
    subject: str
    body: str
    body_html: str | None
    sender: Address
    recipients: list[Address]
    received_at: datetime | None
    importance: str
    is_read: bool
    has_attachments: bool
    attachments: list[Attachment]
    graph_categories: list[str]
    flag_status: str


@dataclass(frozen=True)
class Analysis:
    summary: str
    sentiment: str
    sentiment_score: float
    urgency: str
    intent: str
    keywords: list[str]
    action_items: list[str]
    key_information: dict[str, list[str]] = field(default_factory=dict)


@dataclass(frozen=True)
class EmailResult:
    email: Email
    analysis: Analysis
    categories: list[str]

    def to_dict(self) -> dict:
        result = asdict(self)
        result["email"]["received_at"] = (
            self.email.received_at.isoformat() if self.email.received_at else None
        )
        return result
