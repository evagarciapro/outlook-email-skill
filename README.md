# Outlook Email Skill

A Python 3.10+ library for reading Microsoft Outlook emails through Microsoft
Graph, parsing them into structured records, and analyzing and categorizing them
locally. It **does not send, delete, mark as read, or modify emails**.

## Installation and quick demo

From your checkout, use a virtual environment:

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
python -m pip install -r requirements.txt
python -m pip install -e .
python examples/read_emails.py --offline
```

The offline example needs no credentials or network access. It prints structured
JSON with `email`, `analysis`, and `categories`, including
`["work", "urgent", "follow-up needed"]` for the sample message.

## Microsoft setup

### Interactive sign-in (recommended for individual users)

1. Register an application in the [Microsoft Entra admin center](https://entra.microsoft.com/).
   Choose supported account types appropriate to your organization, or include
   personal Microsoft accounts for Outlook.com.
2. Under **Authentication → Advanced settings**, enable **Allow public client flows**.
   Device-code sign-in does not require a redirect URI or client secret.
3. Add the Microsoft Graph **delegated** permission **Mail.Read**. Grant consent
   as required by your tenant; organizational policies may require administrator
   approval and may block device-code sign-in.
4. Copy the **Application (client) ID**. Use your directory tenant ID for
   organization-only applications. `common` supports organizational and personal
   accounts when the registration permits both; `consumers` is personal-only.

```bash
cp config.example.json config.json
export OUTLOOK_CLIENT_ID="your-application-client-id"
export OUTLOOK_TENANT_ID="your-directory-tenant-id"
python examples/read_emails.py --config config.json --limit 10 --unread-only
```

Follow the displayed Microsoft sign-in instructions. Device codes are
short-lived credentials: do not share or record them. Tokens are cached only in
memory by MSAL; each new process may require sign-in again. Tokens are acquired
silently/refreshed by MSAL before requests when possible.

### Unattended service (organization mailboxes only)

Register an application in your organizational tenant, add Microsoft Graph
**application** permission **Mail.Read**, and obtain **administrator consent**.
Create a client secret and store it in your environment or a secret manager.
Set these values in `config.json`:

```json
{
  "auth_mode": "application",
  "tenant_id": "your-directory-tenant-id",
  "mailbox": "mailbox@example.com"
}
```

This is a settings fragment; merge it into the copied template or use it with
the environment variables below. App-only authentication cannot use `/me`, the
`common` tenant, or personal Outlook.com accounts.

```bash
export OUTLOOK_CLIENT_ID="your-application-client-id"
export OUTLOOK_CLIENT_SECRET="your-client-secret"
python examples/read_emails.py --config config.json --limit 10
```

Application `Mail.Read` can grant broad mailbox access. Have your administrator
restrict the app to intended mailboxes using
[Exchange Online application RBAC](https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac).
The configured mailbox is a query target, **not an authorization boundary**.
Delegated mode supports the signed-in user's mailbox; shared-mailbox delegated
permissions are not implemented by this module.

See [Graph authentication](https://learn.microsoft.com/en-us/graph/auth/),
[list messages](https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0),
and [permissions](https://learn.microsoft.com/en-us/graph/permissions-reference).

## Library usage

```python
from outlook_email_skill import OutlookEmailSkill, OutlookSkillError, SkillConfig

config = SkillConfig.from_file("config.json")

def sign_in(flow):
    # Display interactively, never send this code to a log or another service.
    print(f"Open {flow['verification_uri']} and enter {flow['user_code']}")

try:
    with OutlookEmailSkill(config) as skill:
        skill.authenticate(sign_in)  # Callback is unnecessary in application mode.
        results = skill.read_emails(limit=25, unread_only=True)
        for result in results:
            print(result.email.subject, result.categories)
            print(result.analysis.urgency, result.analysis.action_items)
        # Fetch a specific message:
        # result = skill.read_email("microsoft-graph-message-id")
        # json_ready = result.to_dict()
except OutlookSkillError as exc:
    print(f"Unable to process emails: {exc}")
```

For existing Graph payloads or offline analysis, no credentials are needed:

```python
from outlook_email_skill import OutlookEmailSkill, SkillConfig

with OutlookEmailSkill(SkillConfig(work_domains=("company.example",))) as skill:
    result = skill.process_message({
        "subject": "Project deadline",
        "body": {"contentType": "text", "content": "Please send the report by tomorrow."},
        "from": {"emailAddress": {"address": "manager@company.example"}}
    })
    print(result.to_dict())
```

### Results

- **Email:** ID, subject, plain-text body, original HTML (if present), sender,
  To/Cc recipients, timezone-aware received date, read status, importance,
  Outlook flag/categories, and attachment metadata (ID/name/type/size/inline/kind).
  Missing optional text fields become empty strings; a missing date becomes `None`.
- **Analysis:** extractive summary (first 280 body characters), keyword frequencies,
  sentiment and score from -1 to 1, urgency (`high`, `normal`, `low`), intent
  (`action_required`, `meeting`, `question`, `information`), action sentences,
  and key information (email addresses, HTTP(S) URLs, numeric dates, deadline mentions).
- **Categories:** multiple labels may apply: `work`, `personal`, `urgent`,
  `follow-up needed`; unmatched messages are `other`.
  Original Outlook categories are preserved separately. Local labels are never
  written to Outlook.

Analysis uses **explainable English keyword/phrase heuristics**, not a trained
language model. It cannot reliably interpret negation, sarcasm, quoted threads,
non-English text, or implicit deadlines. Detected dates and deadline sentences
are mentions, not validated dates or resolved calendar commitments. Review
results before making decisions; categories are suggestions, not guarantees.
An urgent phrase or Outlook high importance produces high urgency. Action
phrases, an Outlook follow-up flag, or questions suggest follow-up.

## Configuration

Copy `config.example.json` to `config.json`. Environment variables override
values loaded by `SkillConfig.from_file`; direct `SkillConfig(...)` construction
uses only its supplied values.

| Setting | Default | Purpose |
| --- | --- | --- |
| `client_id` | empty | App registration ID; `OUTLOOK_CLIENT_ID` |
| `tenant_id` | `common` | Tenant ID/domain; `OUTLOOK_TENANT_ID` |
| `auth_mode` | `delegated` | `delegated` or `application`; `OUTLOOK_AUTH_MODE` |
| `client_secret` | empty | Application-mode secret; prefer `OUTLOOK_CLIENT_SECRET` |
| `mailbox` | `me` | Signed-in user or explicit mailbox; `OUTLOOK_MAILBOX` |
| `folder` | `inbox` | Well-known folder name or Graph folder ID |
| `page_size` | `50` | Messages per request, 1–1000 |
| `timeout` | `30` | HTTP timeout in seconds (also passed to MSAL) |
| `max_retries` | `3` | Additional transient retries, 0–10 |
| `include_attachments` | `true` | Fetch attachment metadata, including inline items |
| `work_domains` | empty | Exact, case-insensitive sender domains considered work |
| `work_keywords` | see template | Work keyword/phrase rules |
| `personal_keywords` | see template | Personal keyword/phrase rules |
| `urgent_keywords` | see template | Urgency keyword/phrase rules |

Attachment metadata requires additional paginated requests for **each** email:
Graph's `hasAttachments` excludes inline items. Disable `include_attachments`
for faster reads when attachment details are not needed. No attachment bytes
are downloaded or saved, and reference attachment links are not followed.

The reader follows Graph `@odata.nextLink` URLs unchanged, enforces the requested
message limit, detects pagination cycles, and caps each collection at 1000 pages.
It rejects off-host/non-HTTPS pagination and redirects. Network failures and
HTTP 429/500/502/503/504 retry with backoff or `Retry-After`. Retry waits are bounded:
if Graph requests more than 60 seconds, the call fails so you can retry later
without violating that wait. HTTP 401 requires sign-in again; 403 usually means
missing permissions or mailbox authorization. Other errors are not retried.

## Errors, logging, and privacy

Catch `OutlookSkillError` for all module errors, or its subclasses:
`ConfigurationError`, `AuthenticationError`, `GraphAPIError` (with
`status_code` when applicable), and `EmailParsingError`. Invalid method arguments
raise `ValueError`. Reading is fail-fast: a failed message or attachment request
raises rather than silently returning incomplete results.

The library does not configure your application's logging. Enable it yourself:

```python
import logging
logging.basicConfig(level=logging.INFO)
```

Module logs contain only counts, HTTP status codes, and generic failures—not
subjects, senders, message bodies, tokens, or secrets. Avoid enabling verbose
third-party HTTP/authentication logs in production.

Email analysis runs locally; only Microsoft endpoints are contacted for sign-in
and reading. `config.json`, `.env` files, build artifacts, and logs are git-ignored.
Never commit real credentials, token caches, or exported emails. Returned
records and the example's stdout contain private mailbox content: protect them
and define appropriate access/retention policies if you persist them.

HTML is converted to plain text without rendering; scripts/styles are omitted
from that text. **`body_html` preserves untrusted original HTML, not sanitized
HTML. Never render it without sanitization.** Extracted URLs and attachments
remain untrusted and are never automatically opened.

## Project layout and tests

```text
src/outlook_email_skill/
  skill.py          # Public orchestration API
  config.py         # Settings and environment overrides
  email_reader.py   # MSAL and read-only Graph requests
  parser.py         # Structured email/HTML parsing
  models.py         # Dataclasses and JSON serialization
  analyzer.py       # Local content analysis
  categorizer.py    # Multi-label categorization
  errors.py         # Public exceptions
tests/              # Standard-library unittest tests with mocked Graph/MSAL
examples/read_emails.py
config.example.json
requirements.txt
pyproject.toml
```

```bash
python -m unittest discover -s tests -v
python examples/read_emails.py --offline
python -m pip wheel --no-deps . --wheel-dir /tmp/outlook-skill-wheels
```

Tests require no mailbox, credentials, or live Graph calls. A live integration
test still requires your own app registration, consent, and mailbox. This module
targets the public Microsoft cloud; sovereign-cloud endpoints are not supported.
