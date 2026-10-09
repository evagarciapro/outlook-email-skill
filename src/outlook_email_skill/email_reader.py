"""Read-only Microsoft Graph integration with bounded retries."""

import logging
import time
from collections.abc import Callable, Iterator
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from urllib.parse import quote, urlsplit

import msal
import requests

from .config import SkillConfig
from .errors import AuthenticationError, GraphAPIError

logger = logging.getLogger(__name__)
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
SCOPES = ["https://graph.microsoft.com/Mail.Read"]
MESSAGE_FIELDS = (
    "id,subject,body,from,sender,toRecipients,ccRecipients,receivedDateTime,"
    "hasAttachments,importance,isRead,categories,flag"
)


class EmailReader:
    def __init__(self, config: SkillConfig, session: requests.Session | None = None):
        self.config = config
        self.session = session if session is not None else requests.Session()
        self._owns_session = session is None
        self._app = None
        self._account = None

    def _application(self):
        if self._app is None:
            self.config.validate_credentials()
            authority = f"https://login.microsoftonline.com/{self.config.tenant_id}"
            try:
                if self.config.auth_mode == "application":
                    self._app = msal.ConfidentialClientApplication(
                        self.config.client_id,
                        authority=authority,
                        client_credential=self.config.client_secret,
                        timeout=self.config.timeout,
                    )
                else:
                    self._app = msal.PublicClientApplication(
                        self.config.client_id, authority=authority, timeout=self.config.timeout
                    )
            except (ValueError, requests.RequestException):
                raise AuthenticationError("Unable to initialize Microsoft authentication.") from None
        return self._app

    def authenticate(self, device_callback: Callable[[dict], None] | None = None) -> None:
        """Device callback receives verification_uri and user_code; do not log them."""
        app = self._application()
        try:
            if self.config.auth_mode == "application":
                result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
            else:
                accounts = app.get_accounts()
                self._account = self._account or (accounts[0] if accounts else None)
                result = app.acquire_token_silent(SCOPES, account=self._account) if self._account else None
                if not result or "access_token" not in result:
                    if device_callback is None:
                        raise AuthenticationError("Call authenticate with a device sign-in callback.")
                    flow = app.initiate_device_flow(scopes=SCOPES)
                    if "user_code" not in flow:
                        raise AuthenticationError("Unable to start device sign-in.")
                    device_callback({
                        "verification_uri": flow["verification_uri"], "user_code": flow["user_code"]
                    })
                    result = app.acquire_token_by_device_flow(flow)
                    accounts = app.get_accounts()
                    self._account = accounts[0] if accounts else None
            if not result or "access_token" not in result:
                raise AuthenticationError("Microsoft sign-in failed; check consent and credentials.")
        except (ValueError, requests.RequestException):
            raise AuthenticationError("Microsoft authentication request failed.") from None

    def _access_token(self) -> str:
        app = self._application()
        try:
            if self.config.auth_mode == "application":
                result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
            else:
                accounts = app.get_accounts()
                account = self._account or (accounts[0] if accounts else None)
                result = app.acquire_token_silent(SCOPES, account=account) if account else None
            if not result or not isinstance(result.get("access_token"), str) or not result["access_token"]:
                raise AuthenticationError("Sign-in required; call authenticate before reading emails.")
            return result["access_token"]
        except (ValueError, requests.RequestException):
            raise AuthenticationError("Unable to obtain a Microsoft access token.") from None

    @property
    def _mailbox_path(self) -> str:
        return "/me" if self.config.mailbox == "me" else f"/users/{quote(self.config.mailbox, safe='')}"

    @staticmethod
    def _validate_url(url: str):
        try:
            parsed = urlsplit(url)
            valid = (
                parsed.scheme == "https"
                and parsed.netloc == "graph.microsoft.com"
                and parsed.path.startswith("/v1.0/")
                and not parsed.fragment
            )
        except (TypeError, ValueError):
            valid = False
        if not valid:
            raise GraphAPIError("Rejected an unsafe Graph pagination URL.")

    @staticmethod
    def _retry_delay(response, attempt: int) -> float:
        value = response.headers.get("Retry-After", "")
        try:
            delay = float(value)
        except ValueError:
            try:
                delay = (parsedate_to_datetime(value) - datetime.now(timezone.utc)).total_seconds()
            except (TypeError, ValueError, OverflowError):
                delay = 2 ** attempt
        if delay > 60:
            raise GraphAPIError("Graph requested a retry delay over 60 seconds; retry later.")
        return max(delay, 0)

    def _get(self, url: str, params: dict | None = None) -> dict:
        self._validate_url(url)
        for attempt in range(self.config.max_retries + 1):
            token = self._access_token()
            try:
                response = self.session.get(
                    url,
                    params=params,
                    headers={"Authorization": "Bearer " + token, "Accept": "application/json"},
                    timeout=self.config.timeout,
                    allow_redirects=False,
                )
            except requests.RequestException:
                if attempt == self.config.max_retries:
                    raise GraphAPIError("Microsoft Graph network request failed.") from None
                logger.warning("Retrying Graph request after a network failure.")
                time.sleep(min(2 ** attempt, 60))
                continue
            try:
                status = response.status_code
                if status in {429, 500, 502, 503, 504} and attempt < self.config.max_retries:
                    logger.warning("Retrying Graph request after HTTP %s.", status)
                    time.sleep(self._retry_delay(response, attempt))
                    continue
                if status == 401:
                    raise AuthenticationError("Graph rejected authentication; sign in again.")
                if status != 200:
                    raise GraphAPIError(
                        f"Microsoft Graph request failed (HTTP {status}).", status_code=status
                    )
                try:
                    data = response.json()
                except ValueError:
                    raise GraphAPIError("Microsoft Graph returned invalid JSON.") from None
                if not isinstance(data, dict):
                    raise GraphAPIError("Microsoft Graph returned an invalid response.")
                return data
            finally:
                response.close()
        raise GraphAPIError("Microsoft Graph retry limit exceeded.")

    def _pages(self, url: str, params: dict | None = None) -> Iterator[dict]:
        seen = set()
        while url:
            if not isinstance(url, str) or url in seen or len(seen) >= 1000:
                raise GraphAPIError("Invalid or excessive Graph pagination.")
            seen.add(url)
            page = self._get(url, params)
            items = page.get("value")
            if not isinstance(items, list) or not all(isinstance(item, dict) for item in items):
                raise GraphAPIError("Microsoft Graph returned an invalid collection.")
            yield from items
            url = page.get("@odata.nextLink")
            if url is not None and (not isinstance(url, str) or not url):
                raise GraphAPIError("Microsoft Graph returned an invalid pagination link.")
            params = None

    def list_messages(self, limit: int = 25, unread_only: bool = False) -> list[dict]:
        if type(limit) is not int or limit < 0:
            raise ValueError("limit must be a nonnegative integer.")
        if type(unread_only) is not bool:
            raise ValueError("unread_only must be a boolean.")
        if limit == 0:
            return []
        folder = quote(self.config.folder, safe="")
        url = f"{GRAPH_ROOT}{self._mailbox_path}/mailFolders/{folder}/messages"
        params = {
            "$select": MESSAGE_FIELDS,
            "$top": min(limit, self.config.page_size),
            "$orderby": "receivedDateTime desc",
        }
        if unread_only:
            # Graph requires order-by properties to occur first in the filter.
            params["$filter"] = "receivedDateTime ge 0001-01-01T00:00:00Z and isRead eq false"
        messages = []
        for message in self._pages(url, params):
            messages.append(self._with_attachments(message))
            if len(messages) == limit:
                break
        logger.info("Read %s messages from Microsoft Graph.", len(messages))
        return messages

    def _with_attachments(self, message: dict) -> dict:
        if self.config.include_attachments:
            message_id = message.get("id")
            if not isinstance(message_id, str) or not message_id:
                raise GraphAPIError("Microsoft Graph returned a message without an ID.")
            # hasAttachments excludes inline attachments, so do not use it as a gate.
            url = f"{GRAPH_ROOT}{self._mailbox_path}/messages/{quote(message_id, safe='')}/attachments"
            message = dict(message)
            message["attachments"] = list(self._pages(
                url, {"$select": "id,name,contentType,size,isInline"}
            ))
        return message

    def get_message(self, message_id: str) -> dict:
        if not isinstance(message_id, str) or not message_id:
            raise ValueError("message_id must be a nonempty string.")
        url = f"{GRAPH_ROOT}{self._mailbox_path}/messages/{quote(message_id, safe='')}"
        return self._with_attachments(self._get(url, {"$select": MESSAGE_FIELDS}))

    def close(self):
        if self._owns_session:
            self.session.close()
