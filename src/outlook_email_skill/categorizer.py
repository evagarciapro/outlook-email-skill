"""Multi-label categorization; never modifies the Outlook mailbox."""

from .analyzer import matches
from .config import SkillConfig
from .models import Analysis, Email


class EmailCategorizer:
    def __init__(self, config: SkillConfig):
        self.config = config

    def categorize(self, email: Email, analysis: Analysis) -> list[str]:
        text = f"{email.subject}\n{email.body}"
        domain = email.sender.address.rpartition("@")[2].lower()
        categories = []
        if domain in self.config.work_domains or matches(text, self.config.work_keywords):
            categories.append("work")
        if matches(text, self.config.personal_keywords):
            categories.append("personal")
        if analysis.urgency == "high":
            categories.append("urgent")
        if analysis.intent in {"action_required", "question"}:
            categories.append("follow-up needed")
        return categories or ["other"]
