import json
import unittest
from urllib.parse import urlparse

from chrome_erp_session import BrowserLoginRequired, ChromeErpSession, ERP_HOME
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
    def test_browser_page_and_api_use_the_same_erp_origin(self):
        self.assertEqual(urlparse(ERP_HOME).netloc, "erp.superboss.cc")
        self.assertEqual(urlparse(API_URL).netloc, urlparse(ERP_HOME).netloc)

    def test_failed_fetch_during_login_navigation_is_retryable(self):
        session = ChromeErpSession()
        session._client = _FakeClient({
            "result": {
                "value": json.dumps({"fetchError": "TypeError: Failed to fetch"})
            }
        })

        with self.assertRaises(BrowserLoginRequired):
            session.post_form_json("https://example.test", {}, "")

    def test_browser_stays_open_when_sync_raises(self):
        session = ChromeErpSession()
        client = _FakeClient({})
        session._client = client
        session._started_browser = True

        session.__exit__(RuntimeError, RuntimeError("failed"), None)

        self.assertFalse(any(method == "Browser.close" for method, _ in client.calls))
        self.assertTrue(client.closed)


if __name__ == "__main__":
    unittest.main()
