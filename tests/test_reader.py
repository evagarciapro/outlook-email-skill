import unittest
from unittest.mock import Mock, patch

import requests

from outlook_email_skill import AuthenticationError, ConfigurationError, GraphAPIError, SkillConfig
from outlook_email_skill.email_reader import EmailReader, GRAPH_ROOT, SCOPES


def response(data=None, status=200, headers=None):
    result = Mock()
    result.status_code = status
    result.headers = headers or {}
    result.json.return_value = data
    return result


class ReaderTests(unittest.TestCase):
    def setUp(self):
        self.session = Mock()
        self.reader = EmailReader(
            SkillConfig(client_id="test-app", include_attachments=False, page_size=1),
            session=self.session,
        )
        self.token = patch.object(self.reader, "_access_token", return_value="test-token").start()
        self.addCleanup(patch.stopall)

    def test_pagination_limit_and_unread_filter(self):
        next_url = f"{GRAPH_ROOT}/me/messages?$skiptoken=opaque"
        self.session.get.side_effect = [
            response({"value": [{"id": "a"}], "@odata.nextLink": next_url}),
            response({"value": [{"id": "b"}, {"id": "c"}]}),
        ]
        self.assertEqual(self.reader.list_messages(2, True), [{"id": "a"}, {"id": "b"}])
        first, second = self.session.get.call_args_list
        self.assertEqual(first.kwargs["params"]["$top"], 1)
        self.assertIn("isRead eq false", first.kwargs["params"]["$filter"])
        self.assertEqual(second.args[0], next_url)
        self.assertIsNone(second.kwargs["params"])
        self.assertEqual(first.kwargs["headers"]["Authorization"], "Bearer " + "test-token")
        self.assertFalse(first.kwargs["allow_redirects"])
        self.assertEqual(first.kwargs["timeout"], 30)

    def test_zero_limit_and_input_validation(self):
        self.assertEqual(self.reader.list_messages(0), [])
        self.session.get.assert_not_called()
        for limit in (-1, True, "1"):
            with self.assertRaises(ValueError):
                self.reader.list_messages(limit)
        with self.assertRaises(ValueError):
            self.reader.list_messages(1, unread_only="false")
        with self.assertRaises(ValueError):
            self.reader.get_message("")

    def test_attachment_metadata_including_inline_and_pagination(self):
        self.reader = EmailReader(SkillConfig(client_id="test-app"), session=self.session)
        patch.object(self.reader, "_access_token", return_value="test-token").start()
        attachment_next = f"{GRAPH_ROOT}/me/messages/a/attachments?$skiptoken=next"
        self.session.get.side_effect = [
            response({"id": "a/b", "hasAttachments": False}),
            response({"value": [{"id": "inline", "isInline": True}], "@odata.nextLink": attachment_next}),
            response({"value": [{"id": "file"}]}),
        ]
        message = self.reader.get_message("a/b")
        self.assertEqual(len(message["attachments"]), 2)
        calls = self.session.get.call_args_list
        self.assertIn("/messages/a%2Fb", calls[0].args[0])
        self.assertIn("/messages/a%2Fb/attachments", calls[1].args[0])
        self.assertNotIn("contentBytes", calls[1].kwargs["params"]["$select"])

    def test_application_mailbox_and_folder_are_encoded(self):
        self.reader = EmailReader(SkillConfig(
            client_id="app", tenant_id="tenant", auth_mode="application",
            mailbox="user@example.com", folder="folder/id", include_attachments=False,
        ), session=self.session)
        patch.object(self.reader, "_access_token", return_value="test-token").start()
        self.session.get.return_value = response({"value": []})
        self.reader.list_messages()
        self.assertIn("/users/user%40example.com/mailFolders/folder%2Fid/messages",
                      self.session.get.call_args.args[0])

    def test_unsafe_pagination_never_receives_token(self):
        for url in (
            "https://evil.example/steal", "http://graph.microsoft.com/v1.0/me/messages",
            "https://graph.microsoft.com.evil.example/v1.0/me/messages",
            "https://user@graph.microsoft.com/v1.0/me/messages",
            "https://graph.microsoft.com/v1.0/me/messages#fragment",
            "https://graph.microsoft.com/beta/me/messages",
        ):
            with self.subTest(url=url):
                self.session.get.reset_mock()
                self.session.get.return_value = response({"value": [], "@odata.nextLink": url})
                with self.assertRaises(GraphAPIError):
                    self.reader.list_messages()
                self.assertEqual(self.session.get.call_count, 1)

    def test_pagination_cycles_and_bad_collections(self):
        start = f"{GRAPH_ROOT}/me/mailFolders/inbox/messages"
        for data in (
            {"value": [], "@odata.nextLink": start}, {"value": {}},
            {"value": [None]}, {"value": [], "@odata.nextLink": 1},
            {"value": [], "@odata.nextLink": {}}, [],
        ):
            self.session.get.return_value = response(data)
            with self.subTest(data=data), self.assertRaises(GraphAPIError):
                self.reader.list_messages()

    @patch("outlook_email_skill.email_reader.time.sleep")
    def test_throttle_and_server_retry_close_responses(self, sleep):
        throttle = response(status=429, headers={"Retry-After": "2"})
        unavailable = response(status=503)
        success = response({"value": []})
        self.session.get.side_effect = [throttle, unavailable, success]
        self.assertEqual(self.reader.list_messages(), [])
        self.assertEqual(sleep.call_args_list[0].args, (2.0,))
        self.assertEqual(sleep.call_args_list[1].args, (2,))
        for item in (throttle, unavailable, success):
            item.close.assert_called_once()

    @patch("outlook_email_skill.email_reader.time.sleep")
    def test_network_retry_and_exhaustion(self, sleep):
        self.session.get.side_effect = requests.Timeout("private-token")
        with self.assertRaises(GraphAPIError) as caught:
            self.reader.list_messages()
        self.assertEqual(self.session.get.call_count, 4)
        self.assertNotIn("private-token", str(caught.exception))
        self.assertEqual(sleep.call_count, 3)

    @patch("outlook_email_skill.email_reader.time.sleep")
    def test_status_errors_and_retry_exhaustion(self, sleep):
        for status in (302, 403, 404, 429):
            self.session.get.reset_mock()
            self.session.get.side_effect = None
            self.session.get.return_value = response({"error": "private-content"}, status=status)
            with self.subTest(status=status), self.assertRaises(GraphAPIError) as caught:
                self.reader.list_messages()
            self.assertEqual(caught.exception.status_code, status)
            self.assertNotIn("private-content", str(caught.exception))
            self.assertEqual(self.session.get.call_count, 4 if status == 429 else 1)
        self.session.get.return_value = response(status=401)
        with self.assertRaises(AuthenticationError):
            self.reader.list_messages()

    def test_invalid_json(self):
        self.session.get.return_value = response()
        self.session.get.return_value.json.side_effect = ValueError("sensitive")
        with self.assertRaises(GraphAPIError):
            self.reader.list_messages()

    def test_retry_after_dates_and_long_wait(self):
        self.assertEqual(EmailReader._retry_delay(response(headers={"Retry-After": "bad"}), 2), 4)
        self.assertEqual(EmailReader._retry_delay(response(headers={"Retry-After": "nan"}), 2), 4)
        self.assertEqual(EmailReader._retry_delay(
            response(headers={"Retry-After": "Wed, 21 Oct 2015 07:28:00 GMT"}), 1
        ), 0)
        with self.assertRaises(GraphAPIError):
            EmailReader._retry_delay(response(headers={"Retry-After": "120"}), 1)

    def test_session_ownership(self):
        self.reader.close()
        self.session.close.assert_not_called()
        with patch("outlook_email_skill.email_reader.requests.Session") as factory:
            reader = EmailReader(SkillConfig())
            reader.close()
            factory.return_value.close.assert_called_once()


class AuthenticationTests(unittest.TestCase):
    def test_delegated_device_flow_then_silent_token(self):
        with patch("outlook_email_skill.email_reader.msal.PublicClientApplication") as factory:
            app = factory.return_value
            app.get_accounts.side_effect = [[], [{"id": "account"}], [{"id": "account"}]]
            app.initiate_device_flow.return_value = {
                "user_code": "test-code", "verification_uri": "https://microsoft.com/devicelogin",
                "device_code": "not-for-callback",
            }
            app.acquire_token_by_device_flow.return_value = {"access_token": "test-token"}
            app.acquire_token_silent.return_value = {"access_token": "test-token"}
            callback = Mock()
            reader = EmailReader(SkillConfig(client_id="test-app"))
            self.addCleanup(reader.close)
            reader.authenticate(callback)
            self.assertEqual(reader._access_token(), "test-token")
            self.assertNotIn("device_code", callback.call_args.args[0])
            app.initiate_device_flow.assert_called_once_with(scopes=SCOPES)

    def test_silent_auth_does_not_start_device_flow(self):
        reader = EmailReader(SkillConfig(client_id="test-app"))
        self.addCleanup(reader.close)
        app = reader._app = Mock()
        app.get_accounts.return_value = [{"id": "account"}]
        app.acquire_token_silent.return_value = {"access_token": "test-token"}
        reader.authenticate()
        app.initiate_device_flow.assert_not_called()

    def test_auth_failures_do_not_expose_server_messages(self):
        reader = EmailReader(SkillConfig(client_id="test-app"))
        self.addCleanup(reader.close)
        app = reader._app = Mock()
        app.get_accounts.return_value = []
        with self.assertRaises(AuthenticationError):
            reader._access_token()
        with self.assertRaises(AuthenticationError):
            reader.authenticate()
        app.initiate_device_flow.return_value = {"error_description": "private"}
        with self.assertRaises(AuthenticationError) as caught:
            reader.authenticate(Mock())
        self.assertNotIn("private", str(caught.exception))
        app.initiate_device_flow.return_value = {"user_code": "code", "verification_uri": "uri"}
        app.acquire_token_by_device_flow.return_value = {"error_description": "private"}
        with self.assertRaises(AuthenticationError):
            reader.authenticate(Mock())

    def test_missing_credentials_fail_before_network(self):
        with patch("outlook_email_skill.email_reader.msal.PublicClientApplication") as factory:
            reader = EmailReader(SkillConfig())
            self.addCleanup(reader.close)
            with self.assertRaises(ConfigurationError):
                reader.authenticate()
            factory.assert_not_called()

    def test_application_authentication_and_error(self):
        with patch("outlook_email_skill.email_reader.msal.ConfidentialClientApplication") as factory:
            app = factory.return_value
            app.acquire_token_for_client.return_value = {"access_token": "test-token"}
            reader = EmailReader(SkillConfig(
                client_id="app", client_secret="dummy", tenant_id="tenant",
                auth_mode="application", mailbox="user@example.com",
            ))
            self.addCleanup(reader.close)
            reader.authenticate()
            self.assertEqual(reader._access_token(), "test-token")
            app.acquire_token_for_client.assert_called_with(scopes=["https://graph.microsoft.com/.default"])
            app.acquire_token_for_client.return_value = {"error_description": "private"}
            with self.assertRaises(AuthenticationError):
                reader.authenticate()
            with self.assertRaises(AuthenticationError):
                reader._access_token()


if __name__ == "__main__":
    unittest.main()
