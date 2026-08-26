import json
from pathlib import Path
import tempfile
import unittest
from unittest import mock
from urllib.parse import urlparse

from chrome_erp_session import (
    BrowserLoginRequired,
    ChromeErpSession,
    ERP_HOME,
    _close_extra_page_targets,
    _enable_session_restore,
)
from erp_excel_sync import API_URL


class _FakeClient:
    def __init__(self, result):
        self.result = result
        self.calls = []
        self.closed = False

    def call(self, method, params=None):
        self.calls.append((method, params))
        return self.result

    def close(self):
        self.closed = True


class ChromeErpSessionTests(unittest.TestCase):
    def test_dedicated_profile_restores_last_session(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            profile = Path(temp_dir)
            _enable_session_restore(profile)

            preferences = json.loads(
                (profile / "Default" / "Preferences").read_text(encoding="utf-8")
            )

        self.assertEqual(preferences["session"]["restore_on_startup"], 1)

    def test_browser_page_and_api_use_the_same_erp_origin(self):
        self.assertEqual(urlparse(ERP_HOME).netloc, "erp.superboss.cc")
        self.assertEqual(urlparse(API_URL).netloc, urlparse(ERP_HOME).netloc)

    @mock.patch("chrome_erp_session._debug_json")
    def test_only_one_erp_page_is_kept(self, debug_json):
        targets = [
            {"type": "page", "id": "erp"},
            {"type": "page", "id": "old-page"},
            {"type": "service_worker", "id": "worker"},
        ]

        _close_extra_page_targets(targets, "erp")

        debug_json.assert_called_once_with("/json/close/old-page")

    def test_failed_fetch_during_login_navigation_is_retryable(self):
        session = ChromeErpSession()
        session._client = _FakeClient({
            "result": {
                "value": json.dumps({"fetchError": "TypeError: Failed to fetch"})
            }
        })

        with self.assertRaises(BrowserLoginRequired):
            session.post_form_json("https://example.test", {}, "")

    def test_browser_closes_after_successful_sync(self):
        session = ChromeErpSession()
        client = _FakeClient({})
        session._client = client
        session._started_browser = True

        session.__exit__(None, None, None)

        self.assertTrue(any(method == "Browser.close" for method, _ in client.calls))
        self.assertTrue(client.closed)

    def test_browser_stays_open_when_sync_raises(self):
        session = ChromeErpSession()
        client = _FakeClient({})
        session._client = client

        session.__exit__(RuntimeError, RuntimeError("failed"), None)

        self.assertFalse(any(method == "Browser.close" for method, _ in client.calls))
        self.assertTrue(client.closed)

    def test_view_mode_leaves_browser_open(self):
        session = ChromeErpSession()
        client = _FakeClient({})
        session._client = client
        session.keep_open()

        session.__exit__(None, None, None)

        self.assertFalse(any(method == "Browser.close" for method, _ in client.calls))
        self.assertTrue(client.closed)


if __name__ == "__main__":
    unittest.main()
