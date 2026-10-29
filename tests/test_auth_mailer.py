"""
Tests for authentication configuration and error handling.
Run with:  python -m pytest tests/test_auth_mailer.py -v
"""
import json
import os
import sys
import types
import unittest
from unittest.mock import MagicMock, patch, PropertyMock

# ── Make sure the project root is on sys.path ──────────────────────────────
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import modules.auth   as auth_mod
import modules.mailer as mailer_mod


# ════════════════════════════════════════════════════════════════════════════
# auth module tests
# ════════════════════════════════════════════════════════════════════════════

class TestIsValidClientId(unittest.TestCase):
    """UUID-format validation used before saving and before signing in."""

    def test_valid_uuid_lowercase(self):
        self.assertTrue(auth_mod.is_valid_client_id(
            "12345678-1234-1234-1234-123456789abc"
        ))

    def test_valid_uuid_uppercase(self):
        self.assertTrue(auth_mod.is_valid_client_id(
            "ABCDEF12-ABCD-ABCD-ABCD-ABCDEF123456"
        ))

    def test_placeholder_is_invalid(self):
        self.assertFalse(auth_mod.is_valid_client_id("YOUR_AZURE_APP_CLIENT_ID"))

    def test_empty_string_is_invalid(self):
        self.assertFalse(auth_mod.is_valid_client_id(""))

    def test_too_short_is_invalid(self):
        self.assertFalse(auth_mod.is_valid_client_id("12345678-1234-1234-1234"))

    def test_spaces_stripped_valid(self):
        # is_valid_client_id strips leading/trailing whitespace
        self.assertTrue(auth_mod.is_valid_client_id(
            "  12345678-1234-1234-1234-123456789abc  "
        ))


class TestGetSetClientId(unittest.TestCase):
    """Persisting and loading the client ID."""

    def setUp(self):
        self._orig = auth_mod.CLIENT_ID

    def tearDown(self):
        auth_mod.CLIENT_ID = self._orig

    def test_set_client_id_updates_module_global(self):
        test_id = "aaaabbbb-cccc-dddd-eeee-ffff00001111"
        with patch.object(auth_mod, "_persist_setting"):
            auth_mod.set_client_id(test_id)
        self.assertEqual(auth_mod.CLIENT_ID, test_id)

    def test_set_client_id_strips_whitespace(self):
        test_id = "  aaaabbbb-cccc-dddd-eeee-ffff00001111  "
        with patch.object(auth_mod, "_persist_setting"):
            auth_mod.set_client_id(test_id)
        self.assertEqual(auth_mod.CLIENT_ID, "aaaabbbb-cccc-dddd-eeee-ffff00001111")

    def test_get_client_id_returns_current(self):
        auth_mod.CLIENT_ID = "test-id"
        self.assertEqual(auth_mod.get_client_id(), "test-id")


class TestBuildApp_ConfigError(unittest.TestCase):
    """Building the MSAL app with a placeholder ID raises ConfigError."""

    def setUp(self):
        self._orig = auth_mod.CLIENT_ID

    def tearDown(self):
        auth_mod.CLIENT_ID = self._orig

    def test_raises_config_error_when_placeholder(self):
        auth_mod.CLIENT_ID = "YOUR_AZURE_APP_CLIENT_ID"
        fake_cache = MagicMock()
        with self.assertRaises(auth_mod.ConfigError):
            auth_mod._build_app(fake_cache)

    def test_raises_config_error_when_empty(self):
        auth_mod.CLIENT_ID = ""
        fake_cache = MagicMock()
        with self.assertRaises(auth_mod.ConfigError):
            auth_mod._build_app(fake_cache)


class TestGetAccessToken(unittest.TestCase):
    """get_access_token() behaviour."""

    def setUp(self):
        self._orig = auth_mod.CLIENT_ID
        auth_mod.CLIENT_ID = "12345678-1234-1234-1234-123456789abc"

    def tearDown(self):
        auth_mod.CLIENT_ID = self._orig

    def test_returns_none_when_no_accounts(self):
        mock_cache = MagicMock()
        mock_app   = MagicMock()
        mock_app.get_accounts.return_value = []
        with patch.object(auth_mod, "_build_token_cache", return_value=mock_cache), \
             patch.object(auth_mod, "_build_app",         return_value=mock_app), \
             patch.object(auth_mod, "_save_token_cache"):
            result = auth_mod.get_access_token()
        self.assertIsNone(result)

    def test_returns_token_when_silent_succeeds(self):
        mock_cache = MagicMock()
        mock_app   = MagicMock()
        mock_app.get_accounts.return_value = [{"username": "user@example.com"}]
        mock_app.acquire_token_silent.return_value = {"access_token": "tok_abc"}
        with patch.object(auth_mod, "_build_token_cache", return_value=mock_cache), \
             patch.object(auth_mod, "_build_app",         return_value=mock_app), \
             patch.object(auth_mod, "_save_token_cache"):
            result = auth_mod.get_access_token()
        self.assertEqual(result, "tok_abc")

    def test_raises_token_expired_on_invalid_grant(self):
        mock_cache = MagicMock()
        mock_app   = MagicMock()
        mock_app.get_accounts.return_value = [{"username": "user@example.com"}]
        mock_app.acquire_token_silent.return_value = {
            "error": "invalid_grant",
            "error_description": "token revoked",
        }
        with patch.object(auth_mod, "_build_token_cache", return_value=mock_cache), \
             patch.object(auth_mod, "_build_app",         return_value=mock_app), \
             patch.object(auth_mod, "_save_token_cache"):
            with self.assertRaises(auth_mod.TokenExpiredError):
                auth_mod.get_access_token()


class TestGetUserInfo(unittest.TestCase):
    """verify_connection() wraps get_user_info()."""

    def test_ok_on_200(self):
        mock_resp = MagicMock()
        mock_resp.status_code = 200
        mock_resp.json.return_value = {
            "displayName": "Ada Lovelace",
            "userPrincipalName": "ada@example.com",
        }
        mock_resp.raise_for_status = MagicMock()
        with patch("requests.get", return_value=mock_resp):
            info = auth_mod.get_user_info("fake_token")
        self.assertEqual(info["displayName"], "Ada Lovelace")

    def test_raises_token_expired_on_401(self):
        mock_resp = MagicMock()
        mock_resp.status_code = 401
        with patch("requests.get", return_value=mock_resp):
            with self.assertRaises(auth_mod.TokenExpiredError):
                auth_mod.get_user_info("fake_token")

    def test_raises_runtime_error_on_403(self):
        mock_resp = MagicMock()
        mock_resp.status_code = 403
        with patch("requests.get", return_value=mock_resp):
            with self.assertRaises(RuntimeError):
                auth_mod.get_user_info("fake_token")


class TestVerifyConnection(unittest.TestCase):
    """verify_connection() returns a structured result dict."""

    def test_returns_ok_true_on_success(self):
        with patch.object(auth_mod, "get_user_info", return_value={
            "displayName": "Test User", "userPrincipalName": "test@example.com"
        }):
            r = auth_mod.verify_connection("fake")
        self.assertTrue(r["ok"])
        self.assertEqual(r["display_name"], "Test User")
        self.assertEqual(r["email"], "test@example.com")

    def test_returns_ok_false_expired(self):
        with patch.object(auth_mod, "get_user_info",
                          side_effect=auth_mod.TokenExpiredError("expired")):
            r = auth_mod.verify_connection("fake")
        self.assertFalse(r["ok"])
        self.assertTrue(r.get("expired"))

    def test_returns_ok_false_generic_error(self):
        with patch.object(auth_mod, "get_user_info",
                          side_effect=RuntimeError("network error")):
            r = auth_mod.verify_connection("fake")
        self.assertFalse(r["ok"])
        self.assertIn("network error", r["error"])


class TestLoginInteractive_ConfigError(unittest.TestCase):
    """login_interactive() raises ConfigError before touching MSAL."""

    def setUp(self):
        self._orig = auth_mod.CLIENT_ID

    def tearDown(self):
        auth_mod.CLIENT_ID = self._orig

    def test_raises_config_error_when_not_configured(self):
        auth_mod.CLIENT_ID = "YOUR_AZURE_APP_CLIENT_ID"
        with self.assertRaises(auth_mod.ConfigError):
            auth_mod.login_interactive()


class TestCompleteDeviceFlow_Errors(unittest.TestCase):
    """complete_device_flow() raises the right exception types."""

    def test_raises_consent_error_on_access_denied(self):
        mock_app = MagicMock()
        mock_app.acquire_token_by_device_flow.return_value = {
            "error": "access_denied",
            "error_description": "User cancelled",
        }
        with patch.object(auth_mod, "_save_token_cache"):
            with self.assertRaises(auth_mod.ConsentError):
                auth_mod.complete_device_flow(mock_app, {}, MagicMock())

    def test_raises_runtime_error_on_code_expired(self):
        mock_app = MagicMock()
        mock_app.acquire_token_by_device_flow.return_value = {
            "error": "code_expired",
            "error_description": "The code has expired",
        }
        with patch.object(auth_mod, "_save_token_cache"):
            with self.assertRaises(RuntimeError) as ctx:
                auth_mod.complete_device_flow(mock_app, {}, MagicMock())
        self.assertIn("expired", str(ctx.exception).lower())

    def test_returns_result_on_success(self):
        mock_app = MagicMock()
        mock_app.acquire_token_by_device_flow.return_value = {
            "access_token": "tok123",
            "account": {"username": "user@example.com"},
        }
        with patch.object(auth_mod, "_save_token_cache"):
            result = auth_mod.complete_device_flow(mock_app, {}, MagicMock())
        self.assertEqual(result["access_token"], "tok123")


# ════════════════════════════════════════════════════════════════════════════
# mailer module tests
# ════════════════════════════════════════════════════════════════════════════

class TestPersonalise(unittest.TestCase):
    def test_replaces_name_placeholder(self):
        self.assertEqual(mailer_mod.personalise("Hello {{Name}}", "Ada"), "Hello Ada")

    def test_replaces_lowercase_name(self):
        self.assertEqual(mailer_mod.personalise("Hi {{name}}!", "Bob"), "Hi Bob!")

    def test_no_placeholder_unchanged(self):
        self.assertEqual(mailer_mod.personalise("No placeholders", "X"), "No placeholders")


class TestBuildMessage(unittest.TestCase):
    def test_basic_structure(self):
        msg = mailer_mod.build_message("a@b.com", "Alice", "Hello", "<p>World</p>")
        self.assertEqual(msg["message"]["subject"], "Hello")
        self.assertEqual(msg["message"]["toRecipients"][0]["emailAddress"]["address"], "a@b.com")
        self.assertTrue(msg["saveToSentItems"])

    def test_personalise_in_subject_and_body(self):
        msg = mailer_mod.build_message("a@b.com", "Alice", "Hi {{Name}}", "<p>Dear {{Name}}</p>")
        self.assertEqual(msg["message"]["subject"], "Hi Alice")
        self.assertIn("Dear Alice", msg["message"]["body"]["content"])

    def test_no_attachments_key_when_empty(self):
        msg = mailer_mod.build_message("a@b.com", "Alice", "S", "B")
        self.assertNotIn("attachments", msg["message"])


class TestSendEmail(unittest.TestCase):
    """send_email() HTTP response handling."""

    def _mock_resp(self, status: int, headers=None, json_body=None):
        resp = MagicMock()
        resp.status_code = status
        resp.headers = headers or {}
        if json_body is not None:
            resp.json.return_value = json_body
        else:
            resp.json.side_effect = ValueError("no json")
        resp.text = str(json_body)
        return resp

    def test_202_returns_success(self):
        with patch("requests.post", return_value=self._mock_resp(202)):
            r = mailer_mod.send_email("tok", "a@b.com", "Alice", "S", "B", delay=0)
        self.assertTrue(r["success"])

    def test_401_returns_code_401(self):
        with patch("requests.post", return_value=self._mock_resp(401)):
            r = mailer_mod.send_email("tok", "a@b.com", "Alice", "S", "B", delay=0)
        self.assertFalse(r["success"])
        self.assertEqual(r["code"], 401)
        self.assertIn("expired", r["error"].lower())

    def test_403_returns_code_403(self):
        body = {"error": {"code": "ErrorAccessDenied", "message": "No permission"}}
        with patch("requests.post", return_value=self._mock_resp(403, json_body=body)):
            r = mailer_mod.send_email("tok", "a@b.com", "Alice", "S", "B", delay=0)
        self.assertFalse(r["success"])
        self.assertEqual(r["code"], 403)
        self.assertIn("403", r["error"])

    def test_429_returns_code_429_after_retries(self):
        resp = self._mock_resp(429, headers={"Retry-After": "1"})
        with patch("requests.post", return_value=resp), \
             patch("time.sleep"):
            r = mailer_mod.send_email("tok", "a@b.com", "Alice", "S", "B", delay=0)
        self.assertFalse(r["success"])
        self.assertEqual(r["code"], 429)

    def test_500_retries_and_fails(self):
        resp = self._mock_resp(500, json_body={"error": {"message": "server err"}})
        with patch("requests.post", return_value=resp), \
             patch("time.sleep"):
            r = mailer_mod.send_email("tok", "a@b.com", "Alice", "S", "B", delay=0)
        self.assertFalse(r["success"])
        self.assertEqual(r["code"], 500)

    def test_timeout_fails_gracefully(self):
        import requests as req_lib
        with patch("requests.post", side_effect=req_lib.exceptions.Timeout()), \
             patch("time.sleep"):
            r = mailer_mod.send_email("tok", "a@b.com", "Alice", "S", "B", delay=0)
        self.assertFalse(r["success"])
        self.assertIn("timed out", r["error"].lower())

    def test_connection_error_fails_gracefully(self):
        import requests as req_lib
        with patch("requests.post",
                   side_effect=req_lib.exceptions.ConnectionError("no route")), \
             patch("time.sleep"):
            r = mailer_mod.send_email("tok", "a@b.com", "Alice", "S", "B", delay=0)
        self.assertFalse(r["success"])
        self.assertIn("network", r["error"].lower())


class TestSendTestEmail(unittest.TestCase):
    def test_delegates_to_send_email(self):
        with patch.object(mailer_mod, "send_email", return_value={"success": True}) as m:
            r = mailer_mod.send_test_email("tok", "a@b.com")
        self.assertTrue(r["success"])
        call_kwargs = m.call_args[1]
        self.assertNotIn("access_token", str(call_kwargs).lower().replace("tok", ""))
        # Check subject contains expected text
        self.assertIn("Connection Test", call_kwargs["subject"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
