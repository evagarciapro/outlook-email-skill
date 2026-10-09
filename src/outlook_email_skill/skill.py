"""Public orchestration API for online reading or offline message analysis."""

import logging
from collections.abc import Callable

from .analyzer import EmailAnalyzer
from .categorizer import EmailCategorizer
from .config import SkillConfig
from .email_reader import EmailReader
from .models import EmailResult
from .parser import EmailParser

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class OutlookEmailSkill:
    def __init__(self, config: SkillConfig | None = None, reader: EmailReader | None = None):
        self.config = config if config is not None else SkillConfig()
        self.reader = reader if reader is not None else EmailReader(self.config)
        self.parser = EmailParser()
        self.analyzer = EmailAnalyzer(self.config)
        self.categorizer = EmailCategorizer(self.config)

    def authenticate(self, device_callback: Callable[[dict], None] | None = None):
        self.reader.authenticate(device_callback)

    def process_message(self, message: dict) -> EmailResult:
        """Analyze an existing Graph message without authentication or network access."""
        email = self.parser.parse(message)
        analysis = self.analyzer.analyze(email)
        return EmailResult(email, analysis, self.categorizer.categorize(email, analysis))

    def read_emails(self, limit: int = 25, unread_only: bool = False) -> list[EmailResult]:
        results = [
            self.process_message(message)
            for message in self.reader.list_messages(limit=limit, unread_only=unread_only)
        ]
        logger.info("Analyzed %s emails.", len(results))
        return results

    def read_email(self, message_id: str) -> EmailResult:
        return self.process_message(self.reader.get_message(message_id))

    def close(self):
        self.reader.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()
