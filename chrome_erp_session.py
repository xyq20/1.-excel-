from __future__ import annotations

import base64
import hashlib
import json
import os
from pathlib import Path
import socket
import struct
import subprocess
import sys
import time
from collections.abc import Mapping
from typing import Any
from urllib.error import URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen


MACOS_CHROME_EXECUTABLE = Path(
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
)
ERP_HOME = "https://erpa.superboss.cc/index.html"
ERP_ORIGINS = (
    "https://erpa.superboss.cc",
    "https://erp.superboss.cc",
    "https://viperp.superboss.cc",
)
DEBUG_PORT = 9229
WEBSOCKET_GUID = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"
RETRYABLE_CDP_NAVIGATION_ERRORS = (
    "inspected target navigated or closed",
    "execution context was destroyed",
    "cannot find context with specified id",
    "no target with given id found",
)


class BrowserLoginRequired(RuntimeError):
    pass


def _chrome_executable_candidates(
    platform_name: str | None = None,
    environ: Mapping[str, str] | None = None,
) -> tuple[Path, ...]:
    platform_name = platform_name or sys.platform
    environment = os.environ if environ is None else environ
    if platform_name == "darwin":
        return (MACOS_CHROME_EXECUTABLE,)
    if platform_name == "win32":
        candidates: list[Path] = []
        for variable in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
            root = environment.get(variable)
            if root:
                candidates.append(
                    Path(root) / "Google" / "Chrome" / "Application" / "chrome.exe"
                )
        return tuple(candidates)
    raise RuntimeError(
        f"Unsupported operating system for Chrome ERP login: {platform_name}"
    )


def _find_chrome_executable(
    platform_name: str | None = None,
    environ: Mapping[str, str] | None = None,
) -> Path:
    platform_name = platform_name or sys.platform
    candidates = _chrome_executable_candidates(platform_name, environ)
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    system_label = "macOS" if platform_name == "darwin" else "Windows"
    searched = ", ".join(str(path) for path in candidates) or "no install paths"
    raise RuntimeError(
        f"Google Chrome was not found on {system_label}. Searched: {searched}"
    )


def _chrome_profile_path(
    platform_name: str | None = None,
    environ: Mapping[str, str] | None = None,
    home: Path | None = None,
) -> Path:
    platform_name = platform_name or sys.platform
    environment = os.environ if environ is None else environ
    home = Path.home() if home is None else Path(home)
    if platform_name == "darwin":
        return (
            home
            / "Library"
            / "Application Support"
            / "ERP Excel Sync"
            / "ChromeProfile"
        )
    if platform_name == "win32":
        local_app_data = environment.get("LOCALAPPDATA")
        base = Path(local_app_data) if local_app_data else home / "AppData" / "Local"
        return base / "ERP Excel Sync" / "ChromeProfile"
    raise RuntimeError(
        f"Unsupported operating system for Chrome ERP login: {platform_name}"
    )


def _read_exact(connection: socket.socket, size: int) -> bytes:
    chunks: list[bytes] = []
    remaining = size
    while remaining:
        chunk = connection.recv(remaining)
        if not chunk:
            raise RuntimeError("Chrome browser connection closed unexpectedly")
        chunks.append(chunk)
        remaining -= len(chunk)
    return b"".join(chunks)


class _CdpWebSocket:
    def __init__(self, websocket_url: str) -> None:
        parsed = urlparse(websocket_url)
        if parsed.scheme != "ws" or not parsed.hostname:
            raise RuntimeError(f"Unsupported Chrome debugger URL: {websocket_url}")
        port = parsed.port or 80
        self._connection = socket.create_connection(
            (parsed.hostname, port), timeout=30
        )
        self._connection.settimeout(180)
        key = base64.b64encode(os.urandom(16)).decode("ascii")
        path = parsed.path or "/"
        if parsed.query:
            path += "?" + parsed.query
        request = (
            f"GET {path} HTTP/1.1\r\n"
            f"Host: {parsed.hostname}:{port}\r\n"
            "Upgrade: websocket\r\n"
            "Connection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {key}\r\n"
            "Sec-WebSocket-Version: 13\r\n"
            f"Origin: http://{parsed.hostname}:{port}\r\n\r\n"
        )
        self._connection.sendall(request.encode("ascii"))
        response = bytearray()
        while b"\r\n\r\n" not in response:
            response.extend(_read_exact(self._connection, 1))
            if len(response) > 65536:
                raise RuntimeError("Invalid Chrome websocket handshake")
        header_text = bytes(response).decode("iso-8859-1")
        expected_accept = base64.b64encode(
            hashlib.sha1((key + WEBSOCKET_GUID).encode("ascii")).digest()
        ).decode("ascii")
        if not header_text.startswith("HTTP/1.1 101"):
            raise RuntimeError("Chrome rejected the debugger connection")
        if (
            f"sec-websocket-accept: {expected_accept}".casefold()
            not in header_text.casefold()
        ):
            raise RuntimeError("Chrome debugger handshake validation failed")
        self._next_id = 1

    def _send_frame(self, opcode: int, payload: bytes) -> None:
        mask = os.urandom(4)
        length = len(payload)
        header = bytearray([0x80 | opcode])
        if length < 126:
            header.append(0x80 | length)
        elif length <= 0xFFFF:
            header.append(0x80 | 126)
            header.extend(struct.pack("!H", length))
        else:
            header.append(0x80 | 127)
            header.extend(struct.pack("!Q", length))
        masked = bytes(value ^ mask[index % 4] for index, value in enumerate(payload))
        self._connection.sendall(bytes(header) + mask + masked)

    def _receive_message(self) -> str:
        fragments: list[bytes] = []
        message_opcode: int | None = None
        while True:
            first, second = _read_exact(self._connection, 2)
            final = bool(first & 0x80)
            opcode = first & 0x0F
            masked = bool(second & 0x80)
            length = second & 0x7F
            if length == 126:
                length = struct.unpack("!H", _read_exact(self._connection, 2))[0]
            elif length == 127:
                length = struct.unpack("!Q", _read_exact(self._connection, 8))[0]
            mask = _read_exact(self._connection, 4) if masked else b""
            payload = _read_exact(self._connection, length)
            if masked:
                payload = bytes(
                    value ^ mask[index % 4]
                    for index, value in enumerate(payload)
                )
            if opcode == 0x8:
                raise RuntimeError("Chrome closed the debugger connection")
            if opcode == 0x9:
                self._send_frame(0xA, payload)
                continue
            if opcode == 0xA:
                continue
            if opcode in (0x1, 0x2):
                message_opcode = opcode
                fragments = [payload]
            elif opcode == 0x0 and message_opcode is not None:
                fragments.append(payload)
            else:
                continue
            if final:
                if message_opcode != 0x1:
                    raise RuntimeError("Chrome returned an unsupported binary message")
                return b"".join(fragments).decode("utf-8")

    def call(self, method: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        request_id = self._next_id
        self._next_id += 1
        document = {"id": request_id, "method": method, "params": params or {}}
        self._send_frame(0x1, json.dumps(document).encode("utf-8"))
        while True:
            response = json.loads(self._receive_message())
            if response.get("id") != request_id:
                continue
            if "error" in response:
                message = response["error"].get("message", "unknown Chrome error")
                raise RuntimeError(f"Chrome debugger error: {message}")
            return response.get("result", {})

    def close(self) -> None:
        try:
            self._send_frame(0x8, b"")
        except OSError:
            pass
        self._connection.close()


def _debug_json(path: str, method: str = "GET") -> Any:
    request = Request(
        f"http://127.0.0.1:{DEBUG_PORT}{path}",
        method=method,
    )
    with urlopen(request, timeout=2) as response:
        return json.loads(response.read().decode("utf-8"))


def _enable_session_restore(profile: Path) -> None:
    preferences_path = profile / "Default" / "Preferences"
    preferences_path.parent.mkdir(parents=True, exist_ok=True)
    document: dict[str, Any] = {}
    if preferences_path.exists():
        try:
            loaded = json.loads(preferences_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            loaded = {}
        if isinstance(loaded, dict):
            document = loaded
    session = document.setdefault("session", {})
    if not isinstance(session, dict):
        session = {}
        document["session"] = session
    if session.get("restore_on_startup") == 1:
        return
    session["restore_on_startup"] = 1
    temporary = preferences_path.with_suffix(".tmp")
    temporary.write_text(
        json.dumps(document, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )
    os.replace(temporary, preferences_path)


def _close_extra_page_targets(
    targets: list[dict[str, Any]],
    keep_target_id: str,
) -> None:
    """Keep one ERP page so restored tabs do not grow on every run."""
    for item in targets:
        target_id = item.get("id")
        if (
            item.get("type") != "page"
            or not isinstance(target_id, str)
            or target_id == keep_target_id
        ):
            continue
        try:
            _debug_json(f"/json/close/{target_id}")
        except (OSError, URLError, ValueError):
            pass


def _erp_page_target(targets: list[dict[str, Any]]) -> dict[str, Any] | None:
    for origin in ERP_ORIGINS:
        target = next(
            (
                item
                for item in targets
                if item.get("type") == "page"
                and item.get("url", "").startswith(origin + "/")
            ),
            None,
        )
        if target is not None:
            return target
    return None


def _is_retryable_navigation_error(error: RuntimeError) -> bool:
    message = str(error).casefold()
    return any(marker in message for marker in RETRYABLE_CDP_NAVIGATION_ERRORS)


class ChromeErpSession:
    def __init__(self) -> None:
        self._client: _CdpWebSocket | None = None
        self._started_browser = False
        self._keep_open = False

    def __enter__(self) -> ChromeErpSession:
        try:
            _debug_json("/json/version")
        except (OSError, URLError, ValueError):
            chrome_executable = _find_chrome_executable()
            profile = _chrome_profile_path()
            _enable_session_restore(profile)
            subprocess.Popen(
                [
                    str(chrome_executable),
                    f"--remote-debugging-port={DEBUG_PORT}",
                    f"--user-data-dir={profile}",
                    "--remote-allow-origins=*",
                    "--no-first-run",
                    "--no-default-browser-check",
                    "--restore-last-session",
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            self._started_browser = True

        deadline = time.monotonic() + 30
        targets: list[dict[str, Any]] = []
        while time.monotonic() < deadline:
            try:
                targets = _debug_json("/json/list")
                if targets:
                    break
            except (OSError, URLError, ValueError):
                pass
            time.sleep(0.25)
        if not targets:
            raise RuntimeError("Chrome did not start its ERP login window")

        target = _erp_page_target(targets)
        if target is None:
            target = _debug_json(
                "/json/new?" + ERP_HOME,
                method="PUT",
            )
        target_id = target.get("id")
        if isinstance(target_id, str):
            _close_extra_page_targets(targets, target_id)
        websocket_url = target.get("webSocketDebuggerUrl")
        if not websocket_url:
            raise RuntimeError("Chrome ERP tab has no debugger connection")
        self._replace_client(websocket_url)
        page_deadline = time.monotonic() + 30
        while time.monotonic() < page_deadline:
            location = self._client.call(
                "Runtime.evaluate",
                {
                    "expression": "location.origin",
                    "returnByValue": True,
                },
            )
            if location.get("result", {}).get("value") in ERP_ORIGINS:
                break
            time.sleep(0.25)
        else:
            raise RuntimeError("ERP login page did not finish opening in Chrome")
        self._client.call("Page.bringToFront")
        return self

    def _replace_client(self, websocket_url: str) -> None:
        replacement = _CdpWebSocket(websocket_url)
        try:
            replacement.call("Runtime.enable")
        except BaseException:
            replacement.close()
            raise
        previous = self._client
        self._client = replacement
        if previous is not None:
            previous.close()

    def _reconnect_after_navigation(self) -> None:
        targets = _debug_json("/json/list")
        target = _erp_page_target(targets)
        if target is None:
            target = _debug_json(
                "/json/new?" + ERP_HOME,
                method="PUT",
            )
        websocket_url = target.get("webSocketDebuggerUrl")
        if not websocket_url:
            raise RuntimeError("Chrome ERP tab has no debugger connection")
        self._replace_client(websocket_url)

    def keep_open(self) -> None:
        """Detach after use without closing the dedicated ERP browser."""
        self._keep_open = True

    def show_login_window(self) -> None:
        if self._client is None:
            raise RuntimeError("Chrome ERP session is not connected")
        try:
            self._client.call("Page.navigate", {"url": ERP_HOME})
            self._client.call("Page.bringToFront")
        except RuntimeError as exc:
            if not _is_retryable_navigation_error(exc):
                raise
            self._reconnect_after_navigation()
            assert self._client is not None
            self._client.call("Page.navigate", {"url": ERP_HOME})
            self._client.call("Page.bringToFront")

    def post_form_json(
        self,
        url: str,
        headers: dict[str, str],
        body: str,
    ) -> dict[str, Any]:
        if self._client is None:
            raise RuntimeError("Chrome ERP session is not connected")
        expression = f"""
        (async () => {{
          try {{
            const requestUrl = new URL({json.dumps(url)}, location.href);
            // Login redirects to viperp.superboss.cc; keep the API request same-origin.
            requestUrl.protocol = location.protocol;
            requestUrl.host = location.host;
            const response = await fetch(requestUrl.href, {{
              method: 'POST',
              credentials: 'include',
              headers: {json.dumps(headers)},
              body: {json.dumps(body)}
            }});
            return JSON.stringify({{
              ok: response.ok,
              status: response.status,
              url: response.url,
              contentType: response.headers.get('content-type') || '',
              text: await response.text()
            }});
          }} catch (error) {{
            return JSON.stringify({{fetchError: String(error)}});
          }}
        }})()
        """
        try:
            result = self._client.call(
                "Runtime.evaluate",
                {
                    "expression": expression,
                    "awaitPromise": True,
                    "returnByValue": True,
                },
            )
        except RuntimeError as exc:
            if not _is_retryable_navigation_error(exc):
                raise
            try:
                self._reconnect_after_navigation()
            except (OSError, URLError, ValueError, RuntimeError):
                # The login redirect may still be replacing its target. The
                # caller's existing login wait loop will retry after a pause.
                pass
            raise BrowserLoginRequired(
                "ERP page navigated while the ERP request was running"
            ) from exc
        if "exceptionDetails" in result:
            raise RuntimeError("Chrome could not execute the ERP request")
        value = result.get("result", {}).get("value")
        if not isinstance(value, str):
            raise RuntimeError("Chrome returned no ERP response")
        wrapper = json.loads(value)
        if wrapper.get("fetchError"):
            fetch_error = str(wrapper["fetchError"])
            if "Failed to fetch" in fetch_error:
                raise BrowserLoginRequired(
                    "ERP page is still completing its login navigation"
                )
            raise RuntimeError(f"Chrome ERP request failed: {fetch_error}")
        content_type = str(wrapper.get("contentType", "")).casefold()
        response_url = str(wrapper.get("url", "")).casefold()
        if (
            int(wrapper.get("status", 0)) in (401, 403)
            or "json" not in content_type
            or "login" in response_url
        ):
            raise BrowserLoginRequired("ERP browser login is required")
        try:
            document = json.loads(wrapper.get("text", ""))
        except json.JSONDecodeError as exc:
            raise BrowserLoginRequired("ERP browser login is required") from exc
        return document

    def __exit__(self, exc_type: Any, exc: Any, traceback: Any) -> None:
        if self._client is None:
            return
        try:
            if exc_type is None and not self._keep_open:
                self._client.call("Browser.close")
        except (OSError, RuntimeError):
            pass
        finally:
            self._client.close()
