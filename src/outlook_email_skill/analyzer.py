"""Deterministic English-language heuristics; no external AI service."""

import re
from collections import Counter

from .config import SkillConfig
from .models import Analysis, Email


def matches(text: str, phrases: tuple[str, ...]) -> bool:
    return any(re.search(r"(?<!\w)" + re.escape(phrase) + r"(?!\w)", text, re.I) for phrase in phrases)


class EmailAnalyzer:
    POSITIVE = ("thanks", "thank you", "great", "excellent", "happy", "pleased", "appreciate")
    NEGATIVE = ("disappointed", "unhappy", "angry", "failed", "failure", "problem", "sorry")
    ACTIONS = ("please", "could you", "can you", "action required", "follow up", "follow-up")
    STOP_WORDS = set(
        "the and for you your with this that from have will are was not but can "
        "please thanks hello regards would could should our has".split()
    )

    def __init__(self, config: SkillConfig):
        self.config = config

    def analyze(self, email: Email) -> Analysis:
        text = f"{email.subject}\n{email.body}"
        words = re.findall(r"\b[a-z][a-z'-]*\b", text.lower())
        positive = sum(len(re.findall(r"\b" + re.escape(p) + r"\b", text, re.I)) for p in self.POSITIVE)
        negative = sum(len(re.findall(r"\b" + re.escape(p) + r"\b", text, re.I)) for p in self.NEGATIVE)
        score = (positive - negative) / max(positive + negative, 1)
        sentiment = "positive" if score > 0 else "negative" if score < 0 else "neutral"
        urgency = (
            "high" if email.importance == "high" or matches(text, self.config.urgent_keywords)
            else "low" if email.importance == "low" else "normal"
        )
        sentences = [s.strip() for s in re.split(r"(?<=[.!?])\s+|\n+", text) if s.strip()]
        actions = [s for s in sentences if matches(s, self.ACTIONS)]
        if actions or email.flag_status == "flagged":
            intent = "action_required"
        elif matches(text, ("meeting", "schedule", "appointment", "invitation")):
            intent = "meeting"
        elif "?" in text:
            intent = "question"
        else:
            intent = "information"
        counts = Counter(w for w in words if len(w) > 2 and w not in self.STOP_WORDS)
        information = {
            "email_addresses": sorted(set(re.findall(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}", text))),
            "urls": sorted(set(u.rstrip(".,;!?)") for u in re.findall(r"https?://[^\s<>\"']+", text))),
            "dates": sorted(set(re.findall(
                r"\b(?:\d{4}-\d{2}-\d{2}|\d{1,2}/\d{1,2}/\d{2,4})\b", text
            ))),
            "deadline_mentions": [s for s in sentences if matches(
                s, ("deadline", "due", "by tomorrow", "by today", "end of day")
            )],
        }
        return Analysis(
            summary=" ".join(email.body.split())[:280] or email.subject[:280],
            sentiment=sentiment,
            sentiment_score=score,
            urgency=urgency,
            intent=intent,
            keywords=[word for word, _ in counts.most_common(10)],
            action_items=actions,
            key_information=information,
        )
