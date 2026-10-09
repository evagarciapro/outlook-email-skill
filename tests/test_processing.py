import json
import os
import tempfile
import unittest
from datetime import timezone
from pathlib import Path
from unittest.mock import patch

from outlook_email_skill import ConfigurationError, EmailParsingError, OutlookEmailSkill, SkillConfig
from outlook_email_skill.parser import EmailParser


class ConfigTests(unittest.TestCase):
    def test_defaults_support_offline_usage_but_require_credentials_online(self):
        config = SkillConfig()
        with self.assertRaises(ConfigurationError):
            config.validate_credentials()
        with OutlookEmailSkill(config) as skill:
            self.assertEqual(skill.process_message({}).categories, ["other"])

    def test_environment_overrides_and_secret_not_in_repr(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "settings.json"
            path.write_text(json.dumps({"client_id": "file-id", "work_domains": ["EXAMPLE.COM"]}))
            with patch.dict(os.environ, {"OUTLOOK_CLIENT_ID": "env-id", "OUTLOOK_CLIENT_SECRET": "dummy"}):
                config = SkillConfig.from_file(path)
            self.assertEqual(config.client_id, "env-id")
            self.assertEqual(config.work_domains, ("example.com",))
            self.assertNotIn("dummy", repr(config))

    def test_invalid_settings(self):
        for settings in (
            {"page_size": 0}, {"page_size": True}, {"timeout": float("nan")},
            {"timeout": "30"}, {"max_retries": -1}, {"work_keywords": "project"},
            {"work_domains": [""]}, {"auth_mode": "unknown"}, {"include_attachments": "false"},
            {"tenant_id": "example.com/escape"}, {"auth_mode": "application"},
            {"auth_mode": "application", "mailbox": "person@example.com"},
        ):
            with self.subTest(settings=settings), self.assertRaises(ConfigurationError):
                SkillConfig(**settings)

    def test_invalid_files(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "settings.json"
            for value in ("{", "[]", '{"unknown": 3}'):
                path.write_text(value)
                with self.assertRaises(ConfigurationError):
                    SkillConfig.from_file(path)
            with self.assertRaises(ConfigurationError):
                SkillConfig.from_file(Path(directory) / "missing.json")

    def test_application_credentials(self):
        config = SkillConfig(
            client_id="app", tenant_id="tenant", auth_mode="application",
            mailbox="person@example.com",
        )
        with self.assertRaises(ConfigurationError):
            config.validate_credentials()


class ParserTests(unittest.TestCase):
    def setUp(self):
        self.parser = EmailParser()

    def test_html_sender_date_recipients_and_attachments(self):
        email = self.parser.parse({
            "id": "message", "subject": "Project",
            "body": {"contentType": "HTML", "content": "<p>Hello &amp; welcome</p><script>bad()</script><style>bad</style><p>Next<br>line</p>"},
            "from": {"emailAddress": {"name": "Sender", "address": "sender@example.com"}},
            "toRecipients": [{"emailAddress": {"address": "to@example.com"}}],
            "ccRecipients": [{"emailAddress": {"address": "cc@example.com"}}],
            "receivedDateTime": "2026-10-09T12:00:00Z",
            "hasAttachments": False,
            "attachments": [{
                "id": "attachment", "name": "logo.png", "contentType": "image/png",
                "size": 12, "isInline": True, "@odata.type": "#microsoft.graph.fileAttachment",
            }],
            "categories": ["Existing"], "isRead": True,
        })
        self.assertEqual(email.body, "Hello & welcome\nNext\nline")
        self.assertNotIn("bad", email.body)
        self.assertEqual(email.sender.address, "sender@example.com")
        self.assertEqual(len(email.recipients), 2)
        self.assertEqual(email.received_at.tzinfo, timezone.utc)
        self.assertTrue(email.has_attachments)
        self.assertTrue(email.attachments[0].is_inline)
        self.assertEqual(email.graph_categories, ["Existing"])

    def test_plain_text_and_missing_fields(self):
        email = self.parser.parse({"body": {"contentType": "text", "content": "Hello <world>"}})
        self.assertEqual(email.body, "Hello <world>")
        self.assertIsNone(email.body_html)
        self.assertIsNone(email.received_at)

    def test_malformed_input_has_safe_error(self):
        for message in (
            None, [], {"subject": 42}, {"receivedDateTime": "private-invalid-date"},
            {"from": "invalid"}, {"body": 12}, {"isRead": "false"},
            {"categories": "work"}, {"attachments": [{"size": -1}]},
        ):
            with self.subTest(message=message), self.assertRaises(EmailParsingError) as caught:
                self.parser.parse(message)
            self.assertNotIn("private-invalid-date", str(caught.exception))


class AnalysisTests(unittest.TestCase):
    def setUp(self):
        self.skill = OutlookEmailSkill(SkillConfig(work_domains=("example.com",)))
        self.addCleanup(self.skill.close)

    def test_urgent_work_action_and_key_information(self):
        result = self.skill.process_message({
            "subject": "Urgent project deadline",
            "body": {"contentType": "text", "content": "Please send the report by tomorrow. Thanks! Due 2026-10-10. Contact team@example.com at https://example.com/report."},
            "from": {"emailAddress": {"address": "boss@example.com"}},
            "receivedDateTime": "2026-10-09T12:00:00Z",
        })
        self.assertEqual(result.categories, ["work", "urgent", "follow-up needed"])
        self.assertEqual(result.analysis.intent, "action_required")
        self.assertEqual(result.analysis.sentiment, "positive")
        self.assertTrue(result.analysis.action_items)
        self.assertEqual(result.analysis.key_information["dates"], ["2026-10-10"])
        self.assertEqual(result.analysis.key_information["email_addresses"], ["team@example.com"])
        self.assertEqual(result.analysis.key_information["urls"], ["https://example.com/report"])
        self.assertTrue(result.analysis.key_information["deadline_mentions"])
        json.dumps(result.to_dict())
        self.assertEqual(result.to_dict()["email"]["received_at"], "2026-10-09T12:00:00+00:00")

    def test_personal_and_multiple_categories(self):
        result = self.skill.process_message({"subject": "Family birthday dinner project"})
        self.assertEqual(result.categories, ["work", "personal"])

    def test_importance_flag_and_question(self):
        result = self.skill.process_message({"importance": "high", "flag": {"flagStatus": "flagged"}})
        self.assertEqual(result.categories, ["urgent", "follow-up needed"])
        self.assertEqual(self.skill.process_message({"subject": "How are you?"}).analysis.intent, "question")
        self.assertEqual(self.skill.process_message({"subject": "Meeting invitation"}).analysis.intent, "meeting")
        self.assertEqual(self.skill.process_message({"importance": "low"}).analysis.urgency, "low")

    def test_word_boundaries_sentiment_and_custom_rules(self):
        self.assertEqual(self.skill.process_message({"subject": "Nonurgently"}).analysis.urgency, "normal")
        self.assertEqual(self.skill.process_message({"subject": "Disappointed by failure"}).analysis.sentiment, "negative")
        self.assertEqual(self.skill.process_message({"subject": "Great problem"}).analysis.sentiment, "neutral")
        with OutlookEmailSkill(SkillConfig(urgent_keywords=("priority one",))) as skill:
            self.assertEqual(skill.process_message({"subject": "Priority one"}).analysis.urgency, "high")


class SkillTests(unittest.TestCase):
    def test_reader_injection_orchestration_and_context_cleanup(self):
        from unittest.mock import Mock
        reader = Mock()
        reader.list_messages.return_value = [{"id": "one", "subject": "Project"}]
        reader.get_message.return_value = {"id": "one"}
        with OutlookEmailSkill(reader=reader) as skill:
            self.assertEqual(skill.read_emails(1, True)[0].email.id, "one")
            self.assertEqual(skill.read_email("one").email.id, "one")
            skill.authenticate()
        reader.list_messages.assert_called_once_with(limit=1, unread_only=True)
        reader.get_message.assert_called_once_with("one")
        reader.authenticate.assert_called_once_with(None)
        reader.close.assert_called_once()


if __name__ == "__main__":
    unittest.main()
