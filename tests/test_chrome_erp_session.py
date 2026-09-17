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
    ERP_ORIGINS,
    _chrome_executable_candidates,
    _chrome_profile_path,
    _close_extra_page_targets,
    _enable_session_restore,
    _erp_page_target,
    _find_chrome_executable,
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
    def test_macos_keeps_existing_chrome_and_profile_paths(self):
        self.assertEqual(
            _chrome_executable_candidates("darwin", {}),
            (Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),),
        )
        self.assertEqual(
            _chrome_profile_path("darwin", {}, Path("/Users/tester")),
            Path("/Users/tester/Library/Application Support/ERP Excel Sync/ChromeProfile"),
        )

    def test_windows_finds_chrome_in_each_supported_install_location(self):
        variable_names = ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA")
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            for selected_variable in variable_names:
                with self.subTest(selected_variable=selected_variable):
                    environment = {
                        variable: str(root / variable)
                        for variable in variable_names
                    }
                    executable = (
                        Path(environment[selected_variable])
                        / "Google"
                        / "Chrome"
                        / "Application"
                        / "chrome.exe"
                    )
                    executable.parent.mkdir(parents=True, exist_ok=True)
                    executable.touch()
                    try:
                        self.assertEqual(
                            _find_chrome_executable("win32", environment),
                            executable,
                        )
                    finally:
                        executable.unlink()

    def test_windows_profile_uses_local_app_data(self):
        local_app_data = Path("C:/Users/tester/AppData/Local")
        self.assertEqual(
            _chrome_profile_path(
                "win32",
                {"LOCALAPPDATA": str(local_app_data)},
                Path("C:/Users/tester"),
            ),
            local_app_data / "ERP Excel Sync" / "ChromeProfile",
        )

    def test_missing_windows_chrome_has_clear_error(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            environment = {
                "ProgramFiles": str(Path(temp_dir) / "ProgramFiles"),
                "ProgramFiles(x86)": str(Path(temp_dir) / "ProgramFiles-x86"),
                "LOCALAPPDATA": str(Path(temp_dir) / "LocalAppData"),
            }
            with self.assertRaisesRegex(
                RuntimeError,
                "Google Chrome was not found on Windows",
            ):
                _find_chrome_executable("win32", environment)

    def test_dedicated_profile_restores_last_session(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            profile = Path(temp_dir)
            _enable_session_restore(profile)

            preferences = json.loads(
                (profile / "Default" / "Preferences").read_text(encoding="utf-8")
            )

        self.assertEqual(preferences["session"]["restore_on_startup"], 1)

    def test_browser_page_and_api_use_the_same_erp_origin(self):
        self.assertEqual(urlparse(ERP_HOME).netloc, "erpa.superboss.cc")
        self.assertEqual(urlparse(API_URL).netloc, urlparse(ERP_HOME).netloc)

    def test_current_erp_page_is_preferred_but_legacy_login_is_supported(self):
        legacy = {
            "type": "page",
            "id": "legacy-login",
            "url": "https://erp.superboss.cc/login.html",
        }
        current = {
            "type": "page",
            "id": "current-home",
            "url": ERP_HOME,
        }

        self.assertEqual(ERP_ORIGINS[0], "https://erpa.superboss.cc")
        self.assertIs(_erp_page_target([legacy, current]), current)
        self.assertIs(_erp_page_target([legacy]), legacy)

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

    def test_show_login_window_navigates_home_and_brings_it_forward(self):
        session = ChromeErpSession()
        client = _FakeClient({})
        session._client = client

        session.show_login_window()

        self.assertEqual(
            client.calls,
            [
                ("Page.navigate", {"url": ERP_HOME}),
                ("Page.bringToFront", None),
            ],
        )

    def test_debugger_navigation_error_reconnects_and_becomes_retryable(self):
        session = ChromeErpSession()
        client = mock.Mock()
        client.call.side_effect = RuntimeError(
            "Chrome debugger error: Inspected target navigated or closed"
        )
        session._client = client
        session._reconnect_after_navigation = mock.Mock()

        with self.assertRaisesRegex(
            BrowserLoginRequired,
            "navigated while the ERP request was running",
        ):
            session.post_form_json("https://example.test", {}, "")

        session._reconnect_after_navigation.assert_called_once_with()

    def test_unrelated_debugger_error_is_not_hidden_as_login_navigation(self):
        session = ChromeErpSession()
        client = mock.Mock()
        client.call.side_effect = RuntimeError(
            "Chrome debugger error: Method not found"
        )
        session._client = client
        session._reconnect_after_navigation = mock.Mock()

        with self.assertRaisesRegex(RuntimeError, "Method not found"):
            session.post_form_json("https://example.test", {}, "")

        session._reconnect_after_navigation.assert_not_called()

    @mock.patch("chrome_erp_session._CdpWebSocket")
    @mock.patch("chrome_erp_session._debug_json")
    def test_navigation_recovery_replaces_the_debugger_client(
        self, debug_json, websocket_type
    ):
        debug_json.return_value = [
            {
                "type": "page",
                "id": "erp",
                "url": ERP_HOME,
                "webSocketDebuggerUrl": "ws://127.0.0.1:9229/devtools/page/erp",
            }
        ]
        old_client = _FakeClient({})
        replacement = _FakeClient({})
        websocket_type.return_value = replacement
        session = ChromeErpSession()
        session._client = old_client

        session._reconnect_after_navigation()

        debug_json.assert_called_once_with("/json/list")
        websocket_type.assert_called_once_with(
            "ws://127.0.0.1:9229/devtools/page/erp"
        )
        self.assertEqual(replacement.calls, [("Runtime.enable", None)])
        self.assertIs(session._client, replacement)
        self.assertTrue(old_client.closed)

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
