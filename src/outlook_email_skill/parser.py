"""Parse Microsoft Graph message objects without rendering HTML."""

from datetime import datetime, timezone
from html.parser import HTMLParser

from .errors import EmailParsingError
from .models import Address, Attachment, Email


class _TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self.hidden = 0

    def handle_starttag(self, tag, attrs):
        if tag in {"script", "style"}:
            self.hidden += 1
        if tag in {"p", "div", "br", "li", "tr", "h1", "h2"} and not self.hidden:
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in {"script", "style"} and self.hidden:
            self.hidden -= 1
        if tag in {"p", "div", "li", "tr", "h1", "h2"} and not self.hidden:
            self.parts.append("\n")

    def handle_data(self, data):
        if not self.hidden:
            self.parts.append(data)


def _text(value) -> str:
    if value is None:
        return ""
    if not isinstance(value, str):
        raise ValueError("Expected text")
    return value


def _address(value) -> Address:
    item = (value or {}).get("emailAddress") or {}
    return Address(_text(item.get("name")), _text(item.get("address")))


def _boolean(value) -> bool:
    if type(value) is not bool:
        raise ValueError("Expected boolean")
    return value


class EmailParser:
    def parse(self, message: dict) -> Email:
        try:
            body = message.get("body") or {}
            content = _text(body.get("content"))
            html = content if _text(body.get("contentType")).lower() == "html" else None
            if html is not None:
                extractor = _TextExtractor()
                extractor.feed(html)
                content = "\n".join(
                    line.strip() for line in "".join(extractor.parts).splitlines() if line.strip()
                )
            timestamp = message.get("receivedDateTime")
            received = datetime.fromisoformat(timestamp.replace("Z", "+00:00")) if timestamp else None
            if received and received.tzinfo is None:
                received = received.replace(tzinfo=timezone.utc)
            attachments = []
            for item in message.get("attachments") or []:
                size = item.get("size", 0)
                if type(size) is not int or size < 0:
                    raise ValueError("Invalid attachment size")
                attachments.append(Attachment(
                    id=_text(item.get("id")),
                    name=_text(item.get("name")),
                    content_type=_text(item.get("contentType")),
                    size=size,
                    is_inline=_boolean(item.get("isInline", False)),
                    kind=_text(item.get("@odata.type")),
                ))
            categories = message.get("categories") or []
            if not isinstance(categories, list) or not all(isinstance(c, str) for c in categories):
                raise ValueError("Invalid categories")
            return Email(
                id=_text(message.get("id")),
                subject=_text(message.get("subject")),
                body=content,
                body_html=html,
                sender=_address(message.get("from") or message.get("sender")),
                recipients=[
                    _address(item)
                    for item in (message.get("toRecipients") or []) + (message.get("ccRecipients") or [])
                ],
                received_at=received,
                importance=_text(message.get("importance") or "normal").lower(),
                is_read=_boolean(message.get("isRead", False)),
                has_attachments=_boolean(message.get("hasAttachments", False)) or bool(attachments),
                attachments=attachments,
                graph_categories=categories,
                flag_status=_text((message.get("flag") or {}).get("flagStatus") or "notFlagged"),
            )
        except (AttributeError, TypeError, ValueError, OverflowError):
            raise EmailParsingError("Invalid Microsoft Graph message fields.") from None
