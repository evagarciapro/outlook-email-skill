"""Run against Outlook, or use --offline for a credential-free demonstration."""

import argparse
import json
import logging
import sys

from outlook_email_skill import OutlookEmailSkill, OutlookSkillError, SkillConfig


def show_device_code(flow: dict):
    print(
        f"Open {flow['verification_uri']} and enter code {flow['user_code']} to sign in.",
        file=sys.stderr,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", default="config.json", help="Path to configuration JSON")
    parser.add_argument("--limit", type=int, default=10)
    parser.add_argument("--unread-only", action="store_true")
    parser.add_argument("--offline", action="store_true")
    args = parser.parse_args()
    if args.limit < 0:
        parser.error("--limit must be nonnegative")
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    try:
        config = SkillConfig() if args.offline else SkillConfig.from_file(args.config)
        with OutlookEmailSkill(config) as skill:
            if args.offline:
                results = [skill.process_message({
                    "id": "demo-message",
                    "subject": "Urgent project review",
                    "body": {
                        "contentType": "html",
                        "content": "<p>Please review the report by tomorrow.</p><p>Thanks!</p>",
                    },
                    "from": {"emailAddress": {"name": "Example Sender", "address": "sender@example.com"}},
                    "receivedDateTime": "2026-10-09T12:00:00Z",
                    "importance": "high",
                    "attachments": [{
                        "id": "demo-attachment", "name": "report.pdf",
                        "contentType": "application/pdf", "size": 1024, "isInline": False,
                        "@odata.type": "#microsoft.graph.fileAttachment",
                    }],
                })][:args.limit]
            else:
                skill.authenticate(show_device_code)
                results = skill.read_emails(limit=args.limit, unread_only=args.unread_only)
            print(json.dumps([result.to_dict() for result in results], indent=2))
    except OutlookSkillError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
