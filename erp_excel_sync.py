import argparse
import csv
from datetime import date, datetime, time, timedelta, timezone
from dataclasses import dataclass
from difflib import SequenceMatcher
import getpass
import json
import os
from pathlib import Path
import posixpath
import re
import struct
import sys
import tempfile
import time as time_module
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qsl, urlencode
from urllib.request import Request, urlopen
import zlib
import zipfile
from xml.etree import ElementTree as ET

from chrome_erp_session import BrowserLoginRequired, ChromeErpSession
from monthly_schedule import SyncCycle, resolve_sync_cycle
from xlsx_monthly import (
    apply_monthly_values,
    column_number,
    shift_drawing_anchors,
    shift_qualified_worksheet_formulas,
    update_workbook_xml,
    validate_monthly_sheet,
)


SHANGHAI_TZ = timezone(timedelta(hours=8))
API_URL = "https://erp.superboss.cc/report/sale/dimensions/list"
WORKSHEET_PATH = "xl/worksheets/sheet1.xml"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RETURN_RATE_FIELD = "customA1ED4F3EEFEF30DBB8E9A9A4823B79A3"
COMPARE_FIELDS = ("actualSysConsignCount", RETURN_RATE_FIELD)
DUPLICATE_NAME_MIN_SCORE = 0.60
DUPLICATE_NAME_MIN_MARGIN = 0.15
CELL_RE = re.compile(
    rb'(<c\s+[^>]*?\br="(?P<col2>[A-Z]+)(?P<row2>\d+)"[^>]*/>|'
    rb'<c\s+[^>]*?\br="(?P<col>[A-Z]+)(?P<row>\d+)"[^>]*>(?P<body>.*?)</c>)',
    re.S,
)
ROW_RE = re.compile(rb'<row\b[^>]*\br="(\d+)"[^>]*>.*?</row>', re.S)
BASE_FORM_BODY = (
    "pageNo=1&pageSize=2000&shouldSort=false&sortField=&sortType=&pageId=1302&queryFlag=item&"
    "startTime=&endTime=&sysStatus=created&sellerFlags=&tradeTypes=&excludeTradeTypes=&"
    "containTagIds=&exceptTagIds=&containType=1&exceptType=1&subTagIdsQueryFlag=false&userIds=&"
    "shopUkList=&warehouseIds=&isAccurate=&itemFlag=0&tradeSysStatus=&scalping=&sysSkuIds=&"
    "sysItemIds=&outerIds=&numIids=&platformItemIdQueryFlag=0&platformItemNames=&"
    "platformItemNameQueryFlag=0&platformSkuIdFlag=0&platformSkuIds=&cids=&itemBrandIds=&"
    "skuBrandIds=&containTradeOut=true&onlyTradeOut=false&containNonConsign=true&containCancel=false&"
    "destIds=&sourceIds=&taobaoIds=&supplyIds=&buyerNicks=&buyerNickSelectAll=false&templateIds=&"
    "showProcessItemDetail=0&showGroupItemDetail=0&isOuterIdFuzzy=0&shipper=&queryByCake=&matchFlag=1&"
    "virtualFlag=1&showSuit=0&asTypes=&createdStartTime=&createdEndTime=&buyerNick=&classifyIds=&"
    "classifySkuIds=&itemTagIds=&itemTagQueryType=0&stockLabelIds=&stockLabelQueryType=0&minWeight=&"
    "maxWeight=&authorType=name&authorText=&logisticCompanyIds=&sysConsigned=&definedSearch=&skuCids=&"
    "categoryFilterType=0&expressIds=&provinceNames=&cityNames=&areaNames=&provinceCityAreaFilter=%7B%7D&"
    "street=&asStatus=9%2C2%2C12&itemAttribute=&showSysItem=0&salesmanIds=&showType=normal&"
    "api_name=report_sale_dimensions_list"
)


@dataclass(frozen=True)
class MatchResult:
    status: str
    candidate: dict[str, Any] | None
    confidence: float
    auto_write: bool
    differences: tuple[str, ...] = ()


@dataclass(frozen=True)
class PreparedMonthlyUpdate:
    inserted: bool
    formula_count: int
    rows: tuple[dict[str, Any], ...]
    review: tuple[dict[str, Any], ...]
    critical_results: tuple[dict[str, Any], ...]
    critical_failures: tuple[dict[str, Any], ...]
    replacements: dict[str, bytes]


def date_window_ms(start_date: str, end_date: str) -> tuple[int, int]:
    start = datetime.combine(datetime.strptime(start_date, "%Y-%m-%d").date(), time.min, SHANGHAI_TZ)
    end = datetime.combine(datetime.strptime(end_date, "%Y-%m-%d").date(), time.max, SHANGHAI_TZ)
    return int(start.timestamp() * 1000), int(end.timestamp() * 1000)


def build_payload(start_date: str, end_date: str) -> dict[str, str]:
    start_ms, end_ms = date_window_ms(start_date, end_date)
    payload = dict(parse_qsl(BASE_FORM_BODY, keep_blank_values=True))
    payload["startTime"] = str(start_ms)
    payload["endTime"] = str(end_ms)
    return payload


def normalize_sku(value: Any) -> str:
    text = str(value or "").casefold()
    text = re.sub(r"[‐‑‒–—―]", "-", text)
    return re.sub(r"\s+", "", text)


def normalize_text(value: Any) -> str:
    return re.sub(r"[^0-9a-z\u4e00-\u9fff]+", "", str(value or "").casefold())


def _name_score(name: str, title: str) -> float:
    left = normalize_text(name)
    right = normalize_text(title)
    if not left or not right:
        return 0.0
    if left in right:
        return 1.0
    return SequenceMatcher(None, left, right).ratio()


def _select_by_name(name: str, candidates: list[dict[str, Any]]) -> tuple[dict[str, Any] | None, float]:
    if len(candidates) == 1:
        return candidates[0], _name_score(name, candidates[0].get("title", ""))
    ranked = sorted(
        ((_name_score(name, row.get("title", "")), row) for row in candidates),
        key=lambda item: item[0],
        reverse=True,
    )
    if (
        ranked
        and ranked[0][0] >= DUPLICATE_NAME_MIN_SCORE
        and (
            len(ranked) == 1
            or ranked[0][0] - ranked[1][0] >= DUPLICATE_NAME_MIN_MARGIN
        )
    ):
        return ranked[0][1], ranked[0][0]
    return None, ranked[0][0] if ranked else 0.0


def choose_candidate(
    sheet_sku: Any,
    sheet_name: Any,
    api_rows: list[dict[str, Any]],
    aliases: dict[str, str],
) -> MatchResult:
    sheet_key = normalize_sku(sheet_sku)
    alias_key = aliases.get(sheet_key)
    lookup_key = alias_key or sheet_key
    exact_rows = [row for row in api_rows if normalize_sku(row.get("itemOuterId")) == lookup_key]
    if exact_rows:
        candidate, name_score = _select_by_name(str(sheet_name or ""), exact_rows)
        if candidate is None:
            return MatchResult("ambiguous", None, name_score, False, ("candidate",))
        status = "alias" if alias_key else "exact"
        differences = ("sku",) if alias_key else ()
        return MatchResult(status, candidate, 1.0, True, differences)

    ranked: list[tuple[float, dict[str, Any]]] = []
    for row in api_rows:
        api_key = normalize_sku(row.get("itemOuterId"))
        if not api_key:
            continue
        sku_score = SequenceMatcher(None, sheet_key, api_key).ratio()
        name_score = _name_score(str(sheet_name or ""), row.get("title", ""))
        ranked.append((sku_score * 0.7 + name_score * 0.3, row))
    ranked.sort(key=lambda item: item[0], reverse=True)
    if not ranked or ranked[0][0] < 0.65:
        return MatchResult("unmatched", None, ranked[0][0] if ranked else 0.0, False, ("sku",))
    differences = ["sku"]
    if _name_score(str(sheet_name or ""), ranked[0][1].get("title", "")) < 1.0:
        differences.append("name")
    return MatchResult("review", ranked[0][1], ranked[0][0], False, tuple(differences))


def changed_fields(
    old: dict[str, Any],
    new: dict[str, Any],
    fields: tuple[str, ...] = COMPARE_FIELDS,
) -> dict[str, dict[str, Any]]:
    return {
        field: {"old": old.get(field), "new": new.get(field)}
        for field in fields
        if old.get(field) != new.get(field)
    }


def compare_api_snapshots(
    old_rows: list[dict[str, Any]],
    new_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    def key(row: dict[str, Any]) -> tuple[str, str]:
        return normalize_sku(row.get("itemOuterId")), normalize_text(row.get("title"))

    old_by_key = {key(row): row for row in old_rows}
    new_by_key = {key(row): row for row in new_rows}
    report: list[dict[str, Any]] = []
    for item_key in sorted(set(old_by_key) | set(new_by_key)):
        old = old_by_key.get(item_key)
        new = new_by_key.get(item_key)
        if old is None:
            report.append({
                "status": "added",
                "itemOuterId": new.get("itemOuterId"),
                "title": new.get("title"),
                "changed_fields": {field: {"old": None, "new": new.get(field)} for field in COMPARE_FIELDS},
            })
            continue
        if new is None:
            report.append({
                "status": "removed",
                "itemOuterId": old.get("itemOuterId"),
                "title": old.get("title"),
                "changed_fields": {field: {"old": old.get(field), "new": None} for field in COMPARE_FIELDS},
            })
            continue
        fields = changed_fields(old, new)
        if fields:
            report.append({
                "status": "changed",
                "itemOuterId": new.get("itemOuterId"),
                "title": new.get("title"),
                "changed_fields": fields,
            })
    return report


def set_cell_value(
    row_bytes: bytes,
    col: str,
    row_number: int,
    value: Any,
    default_style: str | None = None,
) -> bytes:
    encoded_value = str(value).encode("ascii")
    for match in CELL_RE.finditer(row_bytes):
        found_col = (match.group("col") or match.group("col2")).decode()
        found_row = int(match.group("row") or match.group("row2"))
        if found_col != col or found_row != row_number:
            continue
        opening = match.group(0).split(b">", 1)[0].rstrip(b"/")
        opening = re.sub(rb'\s+t="[^"]*"', b"", opening)
        opening = re.sub(rb'\s+vm="[^"]*"', b"", opening)
        replacement = opening + b"><v>" + encoded_value + b"</v></c>"
        return row_bytes[: match.start()] + replacement + row_bytes[match.end() :]

    style = b'' if default_style is None else b' s="' + default_style.encode("ascii") + b'"'
    new_cell = (
        b'<c r="'
        + col.encode("ascii")
        + str(row_number).encode("ascii")
        + b'"'
        + style
        + b"><v>"
        + encoded_value
        + b"</v></c>"
    )
    insert_at = row_bytes.rfind(b"</row>")
    target_column = column_number(col)
    for match in CELL_RE.finditer(row_bytes):
        existing_col = (match.group("col") or match.group("col2")).decode()
        if column_number(existing_col) > target_column:
            insert_at = match.start()
            break
    return row_bytes[:insert_at] + new_cell + row_bytes[insert_at:]


def cell_text(row_bytes: bytes, col: str, shared_strings: list[str]) -> str:
    for match in CELL_RE.finditer(row_bytes):
        found_col = (match.group("col") or match.group("col2")).decode()
        if found_col != col:
            continue
        value_match = re.search(rb"<v>(.*?)</v>", match.group("body") or b"", re.S)
        if not value_match:
            return ""
        value = value_match.group(1).decode("utf-8", "ignore")
        attrs = match.group(0).split(b">", 1)[0]
        if re.search(rb'\bt="s"', attrs):
            try:
                return shared_strings[int(value)]
            except (ValueError, IndexError):
                return ""
        return value
    return ""


def _rate_number(value: Any) -> float:
    text = str(value or "0").strip().replace("%", "")
    return float(text or 0) / 100.0


def _number_text(value: float) -> str:
    return ("%.10f" % value).rstrip("0").rstrip(".") or "0"


def _display_sheet_rate(value: str) -> str:
    if value == "":
        return ""
    return f"{float(value) * 100:.2f}%"


def sync_sheet_xml(
    sheet_xml: bytes,
    shared_strings: list[str],
    api_rows: list[dict[str, Any]],
    aliases: dict[str, str],
    target_skus: set[str] | None = None,
) -> tuple[bytes, dict[str, Any]]:
    normalized_aliases = {normalize_sku(key): normalize_sku(value) for key, value in aliases.items()}
    normalized_targets = None if target_skus is None else {normalize_sku(value) for value in target_skus}
    report: dict[str, Any] = {"written": 0, "unchanged": 0, "review": [], "rows": []}

    def replace_row(match: re.Match[bytes]) -> bytes:
        row_number = int(match.group(1))
        row = match.group(0)
        sheet_sku = cell_text(row, "E", shared_strings)
        if not sheet_sku:
            return row
        if normalized_targets is not None and normalize_sku(sheet_sku) not in normalized_targets:
            return row
        sheet_name = cell_text(row, "F", shared_strings)
        result = choose_candidate(sheet_sku, sheet_name, api_rows, normalized_aliases)
        candidate = result.candidate
        if not result.auto_write or candidate is None:
            report["review"].append({
                "row": row_number,
                "sheet_sku": sheet_sku,
                "sheet_name": sheet_name,
                "status": result.status,
                "confidence": round(result.confidence, 4),
                "candidate_sku": candidate.get("itemOuterId") if candidate else "",
                "candidate_title": candidate.get("title") if candidate else "",
                "different_columns": ",".join(result.differences),
            })
            return row

        actual = int(float(candidate.get("actualSysConsignCount") or 0))
        rate = _rate_number(candidate.get(RETURN_RATE_FIELD))
        actual_text = str(actual)
        rate_text = _number_text(rate)
        old_actual = cell_text(row, "Y", shared_strings)
        old_rate = cell_text(row, "AA", shared_strings)
        changed_columns = []
        if old_actual != actual_text:
            changed_columns.append("Y")
        if old_rate != rate_text:
            changed_columns.append("AA")
        patched = set_cell_value(row, "Y", row_number, actual, default_style="141")
        patched = set_cell_value(patched, "AA", row_number, rate_text, default_style="219")
        if changed_columns:
            report["written"] += 1
        else:
            report["unchanged"] += 1
        report["rows"].append({
            "row": row_number,
            "sheet_sku": sheet_sku,
            "sheet_name": sheet_name,
            "match_status": result.status,
            "candidate_sku": candidate.get("itemOuterId"),
            "old_actual": old_actual,
            "new_actual": actual,
            "old_return_rate": _display_sheet_rate(old_rate),
            "new_return_rate": f"{rate * 100:.2f}%",
            "changed_columns": ",".join(changed_columns),
        })
        return patched

    return ROW_RE.sub(replace_row, sheet_xml), report


def fetch_api(
    cookie: str,
    company_id: str,
    start_date: str,
    end_date: str,
    timeout: int = 120,
) -> dict[str, Any]:
    payload = build_payload(start_date, end_date)
    page_size = int(payload["pageSize"])
    rows: list[dict[str, Any]] = []
    first_response: dict[str, Any] | None = None
    page_number = 1
    while True:
        payload["pageNo"] = str(page_number)
        request = Request(
            API_URL,
            data=urlencode(payload).encode("utf-8"),
            method="POST",
            headers={
                "Accept": "application/json, text/plain, */*",
                "Companyid": company_id,
                "Content-Type": "application/x-www-form-urlencoded",
                "Cookie": cookie,
                "Module-Path": "/report/sale_multidimension_next/",
                "Origin": "https://erpa.superboss.cc",
                "Referer": "https://erpa.superboss.cc/index.html",
                "Trackid": f"trackid{int(datetime.now().timestamp() * 1000)}_python",
                "User-Agent": "Mozilla/5.0 ERPExcelSync/1.0",
            },
        )
        try:
            with urlopen(request, timeout=timeout) as response:
                result = json.loads(response.read().decode("utf-8"))
        except HTTPError as exc:
            raise RuntimeError(f"ERP API HTTP {exc.code}: {exc.reason}") from exc
        except URLError as exc:
            raise RuntimeError(f"ERP API connection failed: {exc.reason}") from exc
        except json.JSONDecodeError as exc:
            raise RuntimeError("ERP API did not return JSON; the login session may have expired") from exc
        if result.get("result") != 1 and not result.get("suc"):
            message = result.get("message") or result.get("msg") or "unknown API error"
            raise RuntimeError(f"ERP API rejected the request: {message}")
        if first_response is None:
            first_response = result
        batch = result.get("data", {}).get("list", [])
        rows.extend(batch)
        if len(batch) < page_size:
            break
        page_number += 1
    assert first_response is not None
    first_response.setdefault("data", {})["list"] = rows
    first_response["syncMeta"] = {
        "startDate": start_date,
        "endDate": end_date,
        "fetchedAt": datetime.now(SHANGHAI_TZ).isoformat(timespec="seconds"),
        "pages": page_number,
        "count": len(rows),
    }
    return first_response


def fetch_api_via_chrome(
    session: ChromeErpSession,
    company_id: str,
    start_date: str,
    end_date: str,
    login_timeout: int = 600,
) -> dict[str, Any]:
    payload = build_payload(start_date, end_date)
    page_size = int(payload["pageSize"])
    rows: list[dict[str, Any]] = []
    first_response: dict[str, Any] | None = None
    page_number = 1
    login_deadline = time_module.monotonic() + login_timeout
    login_message_shown = False
    while True:
        payload["pageNo"] = str(page_number)
        headers = {
            "Accept": "application/json, text/plain, */*",
            "Companyid": company_id,
            "Content-Type": "application/x-www-form-urlencoded",
            "Module-Path": "/report/sale_multidimension_next/",
            "Trackid": f"trackid{int(datetime.now().timestamp() * 1000)}_chrome",
        }
        try:
            result = session.post_form_json(
                API_URL,
                headers,
                urlencode(payload),
            )
        except BrowserLoginRequired:
            if time_module.monotonic() >= login_deadline:
                raise RuntimeError("Timed out waiting for ERP browser login")
            if not login_message_shown:
                print(
                    "ERP登录窗口已打开。请在Chrome中完成登录，登录成功后程序会自动继续。",
                    flush=True,
                )
                login_message_shown = True
            time_module.sleep(2)
            continue
        if result.get("result") != 1 and not result.get("suc"):
            message = result.get("message") or result.get("msg") or "unknown API error"
            if any(keyword in str(message) for keyword in ("登录", "登陆", "未认证", "session")):
                if time_module.monotonic() >= login_deadline:
                    raise RuntimeError("Timed out waiting for ERP browser login")
                if not login_message_shown:
                    print(
                        "ERP登录窗口已打开。请在Chrome中完成登录，登录成功后程序会自动继续。",
                        flush=True,
                    )
                    login_message_shown = True
                time_module.sleep(2)
                continue
            raise RuntimeError(f"ERP API rejected the request: {message}")
        if first_response is None:
            first_response = result
        batch = result.get("data", {}).get("list", [])
        if not isinstance(batch, list):
            raise RuntimeError("ERP API response does not contain data.list")
        rows.extend(batch)
        if len(batch) < page_size:
            break
        page_number += 1
    assert first_response is not None
    first_response.setdefault("data", {})["list"] = rows
    first_response["syncMeta"] = {
        "startDate": start_date,
        "endDate": end_date,
        "fetchedAt": datetime.now(SHANGHAI_TZ).isoformat(timespec="seconds"),
        "pages": page_number,
        "count": len(rows),
        "source": "chrome",
    }
    return first_response


def api_rows(document: dict[str, Any]) -> list[dict[str, Any]]:
    rows = document.get("data", {}).get("list")
    if not isinstance(rows, list):
        raise ValueError("API JSON does not contain data.list")
    return rows


def load_json(path: Path) -> dict[str, Any]:
    with path.open(encoding="utf-8-sig") as handle:
        return json.load(handle)


def save_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(value, handle, ensure_ascii=False, indent=2)


def load_aliases(path: Path | None) -> dict[str, str]:
    if path is None or not path.exists():
        return {}
    value = load_json(path)
    if not isinstance(value, dict):
        raise ValueError("Alias file must be a JSON object")
    return {str(key): str(mapped) for key, mapped in value.items()}


def load_target_skus(path: Path | None) -> set[str] | None:
    if path is None:
        return None
    values: set[str] = set()
    with path.open(encoding="utf-8-sig", errors="ignore") as handle:
        for line in handle:
            value = line.strip()
            if value and value != "货号":
                values.add(value)
    return values


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return [
        "".join(node.text or "" for node in item.iter(f"{{{MAIN_NS}}}t"))
        for item in root
    ]


def _zip_extra_has_zip64(extra: bytes) -> bool:
    position = 0
    while position + 4 <= len(extra):
        field_id, size = struct.unpack_from("<HH", extra, position)
        position += 4
        if position + size > len(extra):
            raise RuntimeError("malformed ZIP extra field")
        if field_id == 0x0001:
            return True
        position += size
    return False


def _zip_end_record(
    data: bytes,
) -> tuple[int, tuple[int, int, int, int, int, int], bytes]:
    search_start = max(0, len(data) - (22 + 0xFFFF))
    position = data.rfind(b"PK\x05\x06", search_start)
    while position >= 0:
        if position + 22 <= len(data):
            candidate_comment_size = struct.unpack_from("<H", data, position + 20)[0]
            if position + 22 + candidate_comment_size == len(data):
                break
        position = data.rfind(b"PK\x05\x06", search_start, position)
    if position < 0:
        raise RuntimeError("ZIP end record not found")
    disk, central_disk, disk_count, total_count, central_size, central_start, comment_size = (
        struct.unpack_from("<HHHHIIH", data, position + 4)
    )
    if (
        disk_count == 0xFFFF
        or total_count == 0xFFFF
        or central_size == 0xFFFFFFFF
        or central_start == 0xFFFFFFFF
        or data[max(0, position - 20) : position].startswith(b"PK\x06\x07")
    ):
        raise RuntimeError("ZIP64 archives are unsupported")
    if disk != 0 or central_disk != 0 or disk_count != total_count:
        raise RuntimeError("multi-disk ZIP archives are unsupported")
    return (
        position,
        (disk_count, total_count, central_size, central_start, comment_size, disk),
        data[position : position + 22 + comment_size],
    )


def _central_records(
    data: bytes,
    central_start: int,
    central_size: int,
    count: int,
) -> tuple[list[bytes], bytes]:
    records: list[bytes] = []
    position = central_start
    for _index in range(count):
        if data[position : position + 4] != b"PK\x01\x02":
            raise RuntimeError("ZIP central directory signature not found")
        name_size, extra_size, comment_size = struct.unpack_from(
            "<HHH", data, position + 28
        )
        disk_start = struct.unpack_from("<H", data, position + 34)[0]
        compressed_size, file_size, local_offset = struct.unpack_from(
            "<III", data, position + 20
        )
        length = 46 + name_size + extra_size + comment_size
        if position + length > len(data):
            raise RuntimeError("truncated ZIP central directory")
        extra = data[
            position + 46 + name_size : position + 46 + name_size + extra_size
        ]
        if (
            disk_start != 0
            or compressed_size == 0xFFFFFFFF
            or file_size == 0xFFFFFFFF
            or local_offset == 0xFFFFFFFF
            or _zip_extra_has_zip64(extra)
        ):
            if disk_start != 0:
                raise RuntimeError("multi-disk ZIP archives are unsupported")
            raise RuntimeError("ZIP64 archives are unsupported")
        records.append(data[position : position + length])
        position += length
    central_end = central_start + central_size
    if position > central_end or central_end > len(data):
        raise RuntimeError("invalid ZIP central directory size")
    return records, data[position:central_end]


def _raw_local_records(
    data: bytes,
    infos: list[zipfile.ZipInfo],
    central_start: int,
) -> dict[int, bytes]:
    ordered = sorted(infos, key=lambda item: item.header_offset)
    boundaries = [item.header_offset for item in ordered[1:]] + [central_start]
    records: dict[int, bytes] = {}
    for info, end in zip(ordered, boundaries):
        start = info.header_offset
        if data[start : start + 4] != b"PK\x03\x04" or end < start:
            raise RuntimeError(f"unexpected ZIP local header for {info.filename}")
        name_size, extra_size = struct.unpack_from("<HH", data, start + 26)
        extra = data[start + 30 + name_size : start + 30 + name_size + extra_size]
        local_sizes = struct.unpack_from("<II", data, start + 18)
        if 0xFFFFFFFF in local_sizes or _zip_extra_has_zip64(extra):
            raise RuntimeError("ZIP64 archives are unsupported")
        records[start] = data[start:end]
    return records


def _compress_replacement(value: bytes, method: int) -> bytes:
    if method == zipfile.ZIP_STORED:
        return value
    if method == zipfile.ZIP_DEFLATED:
        compressor = zlib.compressobj(level=6, wbits=-15)
        return compressor.compress(value) + compressor.flush()
    raise RuntimeError(f"unsupported compression method {method}")


def _replacement_local_record(
    original: bytes,
    info: zipfile.ZipInfo,
    value: bytes,
) -> tuple[bytes, int, int]:
    name_size, extra_size = struct.unpack_from("<HH", original, 26)
    header_size = 30 + name_size + extra_size
    header = bytearray(original[:header_size])
    compressed = _compress_replacement(value, info.compress_type)
    crc = zlib.crc32(value) & 0xFFFFFFFF
    if len(value) > 0xFFFFFFFF or len(compressed) > 0xFFFFFFFF:
        raise RuntimeError("ZIP64 archives are unsupported")
    tail_start = header_size + info.compress_size
    old_tail = original[tail_start:]
    if info.flag_bits & 0x08:
        struct.pack_into("<III", header, 14, 0, 0, 0)
        signature = b"PK\x07\x08" if old_tail.startswith(b"PK\x07\x08") else b""
        descriptor_size = 16 if signature else 12
        remainder = old_tail[descriptor_size:]
        descriptor = signature + struct.pack("<III", crc, len(compressed), len(value))
        tail = descriptor + remainder
    else:
        struct.pack_into("<III", header, 14, crc, len(compressed), len(value))
        tail = old_tail
    return bytes(header) + compressed + tail, crc, len(compressed)


def fast_patch_zip(
    source: Path,
    output: Path,
    replacements: dict[str, bytes],
) -> None:
    source = Path(source)
    output = Path(output)
    if source.resolve() == output.resolve():
        raise RuntimeError("source and output must be different paths")
    if not isinstance(replacements, dict) or not replacements:
        raise RuntimeError("at least one ZIP replacement is required")
    original_data = source.read_bytes()
    _end_position, end_values, end_record = _zip_end_record(original_data)
    _disk_count, total_count, central_size, central_start, _comment_size, _disk = end_values
    central_records, central_trailer = _central_records(
        original_data, central_start, central_size, total_count
    )
    with zipfile.ZipFile(source, "r") as archive:
        infos = archive.infolist()
        names = {info.filename for info in infos}
        missing = sorted(set(replacements) - names)
        if missing:
            raise RuntimeError(f"missing replacement target(s): {', '.join(missing)}")
        if len(infos) != total_count or len(central_records) != len(infos):
            raise RuntimeError("ZIP member count does not match central directory")
        encrypted = [info.filename for info in infos if info.flag_bits & 0x01]
        if encrypted:
            raise RuntimeError(f"encrypted ZIP member is unsupported: {encrypted[0]}")
        local_records = _raw_local_records(original_data, infos, central_start)

        output.parent.mkdir(parents=True, exist_ok=True)
        new_offsets: list[int] = []
        replacement_metadata: dict[int, tuple[int, int, int]] = {}
        with output.open("wb") as destination:
            for index, info in enumerate(infos):
                new_offsets.append(destination.tell())
                original_record = local_records[info.header_offset]
                if info.filename in replacements:
                    value = replacements[info.filename]
                    if not isinstance(value, bytes):
                        raise TypeError(f"replacement for {info.filename} must be bytes")
                    record, crc, compressed_size = _replacement_local_record(
                        original_record, info, value
                    )
                    replacement_metadata[index] = (crc, compressed_size, len(value))
                    destination.write(record)
                else:
                    destination.write(original_record)

            new_central_start = destination.tell()
            for index, original_record in enumerate(central_records):
                record = bytearray(original_record)
                if index in replacement_metadata:
                    struct.pack_into("<III", record, 16, *replacement_metadata[index])
                if new_offsets[index] > 0xFFFFFFFF:
                    raise RuntimeError("ZIP64 archives are unsupported")
                struct.pack_into("<I", record, 42, new_offsets[index])
                destination.write(record)
            destination.write(central_trailer)
            new_central_size = destination.tell() - new_central_start
            if new_central_size > 0xFFFFFFFF or new_central_start > 0xFFFFFFFF:
                raise RuntimeError("ZIP64 archives are unsupported")
            updated_end = bytearray(end_record)
            struct.pack_into("<II", updated_end, 12, new_central_size, new_central_start)
            destination.write(updated_end)


def create_same_directory_temp(master: Path) -> Path:
    master = Path(master)
    descriptor, name = tempfile.mkstemp(
        dir=master.parent,
        prefix=f".{master.name}.",
        suffix=".tmp",
    )
    os.close(descriptor)
    return Path(name)


def atomic_replace_master(
    master: Path,
    validated_temp: Path,
    replace=os.replace,
) -> Path:
    master = Path(master)
    validated_temp = Path(validated_temp)
    backup = master.with_suffix(master.suffix + ".bak")
    replace(master, backup)
    try:
        replace(validated_temp, master)
    except BaseException:
        replace(backup, master)
        raise
    return backup


REQUIRED_WORKBOOK_PARTS = (
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    WORKSHEET_PATH,
)


def _validated_archive_parts(
    path: Path,
) -> tuple[dict[str, bytes], dict[str, tuple[int, int]], list[str]]:
    try:
        with zipfile.ZipFile(path, "r") as archive:
            bad_member = archive.testzip()
            if bad_member is not None:
                raise RuntimeError(f"workbook CRC failure in {bad_member}")
            names = set(archive.namelist())
            missing = [name for name in REQUIRED_WORKBOOK_PARTS if name not in names]
            if missing:
                raise RuntimeError(
                    f"missing required workbook part(s): {', '.join(missing)}"
                )
            xml_parts: dict[str, bytes] = {}
            for name in archive.namelist():
                if not (name.endswith(".xml") or name.endswith(".rels")):
                    continue
                value = archive.read(name)
                try:
                    ET.fromstring(value)
                except ET.ParseError as exc:
                    raise RuntimeError(f"readable XML check failed for {name}: {exc}") from exc
                if name in REQUIRED_WORKBOOK_PARTS:
                    xml_parts[name] = value
            media = {
                info.filename: (info.CRC, info.file_size)
                for info in archive.infolist()
                if info.filename.startswith("xl/media/") and not info.is_dir()
            }
    except zipfile.BadZipFile as exc:
        raise RuntimeError(f"invalid ZIP workbook {path}: {exc}") from exc

    workbook_root = ET.fromstring(xml_parts["xl/workbook.xml"])
    sheet_names = [
        str(node.attrib.get("name", ""))
        for node in workbook_root.iter()
        if node.tag.rsplit("}", 1)[-1] == "sheet"
    ]
    return xml_parts, media, sheet_names


def _xml_row_count(sheet_xml: bytes) -> int:
    root = ET.fromstring(sheet_xml)
    return sum(1 for node in root.iter() if node.tag.rsplit("}", 1)[-1] == "row")


def validate_workbook(
    source: Path,
    candidate: Path,
    cycle: SyncCycle,
) -> dict[str, Any]:
    source = Path(source)
    candidate = Path(candidate)
    source_parts, source_media, source_sheet_names = _validated_archive_parts(source)
    candidate_parts, candidate_media, candidate_sheet_names = _validated_archive_parts(
        candidate
    )
    if candidate_sheet_names != source_sheet_names:
        raise RuntimeError(
            f"sheet names changed: source={source_sheet_names!r}, "
            f"candidate={candidate_sheet_names!r}"
        )
    if candidate_media != source_media:
        raise RuntimeError("media manifest differs between source and candidate")

    source_sheet = source_parts[WORKSHEET_PATH]
    candidate_sheet = candidate_parts[WORKSHEET_PATH]
    source_row_count = _xml_row_count(source_sheet)
    candidate_row_count = _xml_row_count(candidate_sheet)
    if candidate_row_count != source_row_count:
        raise RuntimeError(
            f"row count changed: source={source_row_count}, candidate={candidate_row_count}"
        )

    with zipfile.ZipFile(candidate, "r") as archive:
        shared_strings = (
            read_shared_strings(archive)
            if "xl/sharedStrings.xml" in archive.namelist()
            else []
        )
    try:
        sheet_summary = validate_monthly_sheet(candidate_sheet, shared_strings, cycle)
    except (RuntimeError, ValueError) as exc:
        raise RuntimeError(f"monthly worksheet validation failed: {exc}") from exc
    return {
        "source": str(source),
        "candidate": str(candidate),
        "sheet_names": candidate_sheet_names,
        "media_count": len(candidate_media),
        **sheet_summary,
    }


def sync_workbook(
    source: Path,
    output: Path | None,
    rows: list[dict[str, Any]],
    aliases: dict[str, str],
    target_skus: set[str] | None,
) -> dict[str, Any]:
    with zipfile.ZipFile(source, "r") as archive:
        sheet_xml = archive.read(WORKSHEET_PATH)
        strings = read_shared_strings(archive)
    modified_sheet, report = sync_sheet_xml(sheet_xml, strings, rows, aliases, target_skus)
    if output is not None:
        fast_patch_zip(source, output, {WORKSHEET_PATH: modified_sheet})
    return report


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def flatten_api_differences(differences: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for item in differences:
        for field, values in item["changed_fields"].items():
            rows.append({
                "status": item["status"],
                "itemOuterId": item.get("itemOuterId"),
                "title": item.get("title"),
                "field": field,
                "old": values.get("old"),
                "new": values.get("new"),
            })
    return rows


def _drawing_parts_for_main_sheet(
    archive: zipfile.ZipFile,
) -> tuple[str, ...]:
    rels_path = "xl/worksheets/_rels/sheet1.xml.rels"
    if rels_path not in archive.namelist():
        return ()
    root = ET.fromstring(archive.read(rels_path))
    drawing_paths: list[str] = []
    for relationship in root.iter():
        if not relationship.attrib.get("Type", "").endswith("/drawing"):
            continue
        if relationship.attrib.get("TargetMode", "").casefold() == "external":
            continue
        target = relationship.attrib.get("Target", "")
        if not target:
            continue
        if target.startswith("/"):
            resolved = target.lstrip("/")
        else:
            resolved = posixpath.normpath(
                posixpath.join(posixpath.dirname(WORKSHEET_PATH), target)
            )
        if resolved in archive.namelist():
            drawing_paths.append(resolved)
    return tuple(dict.fromkeys(drawing_paths))


def prepare_monthly_update(
    workbook: Path,
    cycle: SyncCycle,
    actual_rows: list[dict[str, Any]],
    return_rows: list[dict[str, Any]],
    aliases: dict[str, str],
    critical_skus: set[str] | None,
) -> PreparedMonthlyUpdate:
    with zipfile.ZipFile(workbook, "r") as archive:
        names = archive.namelist()
        sheet_xml = archive.read(WORKSHEET_PATH)
        shared_strings = (
            read_shared_strings(archive)
            if "xl/sharedStrings.xml" in names
            else []
        )
        result = apply_monthly_values(
            sheet_xml,
            shared_strings,
            cycle,
            actual_rows,
            return_rows,
            aliases,
            critical_skus or set(),
            choose_candidate,
        )

        replacements: dict[str, bytes] = {WORKSHEET_PATH: result.sheet_xml}
        insert_before = result.layout.previous_col if result.inserted else None
        workbook_xml = update_workbook_xml(
            archive.read("xl/workbook.xml"), insert_before
        )
        replacements["xl/workbook.xml"] = workbook_xml

        if result.inserted:
            for name in names:
                if (
                    name.startswith("xl/worksheets/")
                    and name.endswith(".xml")
                    and name != WORKSHEET_PATH
                ):
                    original = archive.read(name)
                    shifted = shift_qualified_worksheet_formulas(
                        original, result.layout.previous_col
                    )
                    if shifted != original:
                        replacements[name] = shifted
            for name in _drawing_parts_for_main_sheet(archive):
                original = archive.read(name)
                shifted = shift_drawing_anchors(
                    original, result.layout.previous_col
                )
                if shifted != original:
                    replacements[name] = shifted

    return PreparedMonthlyUpdate(
        inserted=result.inserted,
        formula_count=result.formula_count,
        rows=result.rows,
        review=result.review,
        critical_results=result.critical_results,
        critical_failures=result.critical_failures,
        replacements=replacements,
    )


def _load_or_fetch_window(
    json_path: Path | None,
    cookie: str | None,
    company_id: str,
    start: str,
    end: str,
) -> dict[str, Any]:
    if json_path is not None:
        return load_json(json_path)
    if cookie is None:
        raise RuntimeError("ERP Cookie is required")
    return fetch_api(cookie, company_id, start, end)


def _previous_snapshot(
    snapshot_dir: Path,
    prefix: str,
    current: Path,
) -> dict[str, Any] | None:
    candidates = sorted(
        (path for path in snapshot_dir.glob(f"{prefix}_*.json") if path != current),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    return load_json(candidates[0]) if candidates else None


def _monthly_report_rows(
    rows: tuple[dict[str, Any], ...],
    field: str,
) -> list[dict[str, Any]]:
    if field == "actual":
        keys = ("row", "sheet_sku", "sheet_name", "actual_status", "old_actual", "new_actual")
    else:
        keys = ("row", "sheet_sku", "sheet_name", "return_status", "old_return", "new_return")
    return [{key: row.get(key) for key in keys} for row in rows]


def run_monthly_sync(args: argparse.Namespace, run_date: date | None = None) -> int:
    run_date = run_date or datetime.now(SHANGHAI_TZ).date()
    cycle = resolve_sync_cycle(run_date)
    workbook = Path(args.workbook) if args.workbook is not None else None
    if workbook is None:
        raise SystemExit("Master workbook is required; set workbook in config.json or pass --workbook")
    if not workbook.exists():
        raise SystemExit(f"Master workbook not found: {workbook}")
    if (args.actual_json is None) != (args.return_json is None):
        raise SystemExit("Offline mode requires both --actual-json and --return-json")
    browser_login = bool(getattr(args, "browser_login", False)) and sys.platform == "darwin"

    actual_start, actual_end = cycle.actual_window.iso()
    return_start, return_end = cycle.return_window.iso()
    cookie: str | None = None
    if args.actual_json is None and not browser_login:
        cookie = os.environ.get(args.cookie_env)
        if not cookie:
            cookie = getpass.getpass(f"{args.cookie_env} is not set; paste ERP Cookie: ")
        if not cookie.strip():
            raise SystemExit("ERP Cookie is required")

    if args.actual_json is not None:
        actual_document = load_json(args.actual_json)
        return_document = load_json(args.return_json)
    elif browser_login:
        print("正在打开ERP专用Chrome登录窗口……", flush=True)
        with ChromeErpSession() as browser_session:
            actual_document = fetch_api_via_chrome(
                browser_session,
                args.company_id,
                actual_start,
                actual_end,
            )
            return_document = fetch_api_via_chrome(
                browser_session,
                args.company_id,
                return_start,
                return_end,
            )
    else:
        actual_document = _load_or_fetch_window(
            None, cookie, args.company_id, actual_start, actual_end
        )
        return_document = _load_or_fetch_window(
            None, cookie, args.company_id, return_start, return_end
        )

    snapshot_dir = Path(args.snapshot_dir)
    snapshot_dir.mkdir(parents=True, exist_ok=True)
    actual_snapshot = snapshot_dir / f"actual_{actual_start}_{actual_end}.json"
    return_snapshot = snapshot_dir / f"return_{return_start}_{return_end}.json"
    previous_actual = _previous_snapshot(snapshot_dir, "actual", actual_snapshot)
    previous_return = _previous_snapshot(snapshot_dir, "return", return_snapshot)
    save_json(actual_snapshot, actual_document)
    save_json(return_snapshot, return_document)

    actual_rows = api_rows(actual_document)
    return_rows = api_rows(return_document)
    aliases = load_aliases(args.aliases)
    critical_skus = load_target_skus(args.skus_file)
    prepared = prepare_monthly_update(
        workbook,
        cycle,
        actual_rows,
        return_rows,
        aliases,
        critical_skus,
    )

    report_dir = Path(args.report_dir) / cycle.node_date.isoformat()
    actual_fields = ["row", "sheet_sku", "sheet_name", "actual_status", "old_actual", "new_actual"]
    return_fields = ["row", "sheet_sku", "sheet_name", "return_status", "old_return", "new_return"]
    write_csv(report_dir / "actual_changes.csv", _monthly_report_rows(prepared.rows, "actual"), actual_fields)
    write_csv(report_dir / "return_changes.csv", _monthly_report_rows(prepared.rows, "return"), return_fields)
    write_csv(
        report_dir / "manual_review.csv",
        list(prepared.review),
        ["field", "row", "sheet_sku", "sheet_name", "status", "confidence", "candidate_sku", "candidate_title", "different_columns"],
    )
    write_csv(
        report_dir / "critical_skus.csv",
        list(prepared.critical_results),
        ["sku", "sheet_exists", "actual_status", "return_status", "actual_candidate_sku", "actual_candidate_title", "return_candidate_sku", "return_candidate_title", "passed"],
    )
    actual_differences = (
        compare_api_snapshots(api_rows(previous_actual), actual_rows)
        if previous_actual is not None else []
    )
    return_differences = (
        compare_api_snapshots(api_rows(previous_return), return_rows)
        if previous_return is not None else []
    )
    difference_fields = ["status", "itemOuterId", "title", "field", "old", "new"]
    write_csv(report_dir / "api_actual_changes.csv", flatten_api_differences(actual_differences), difference_fields)
    write_csv(report_dir / "api_return_changes.csv", flatten_api_differences(return_differences), difference_fields)

    summary: dict[str, Any] = {
        "runDate": run_date.isoformat(),
        "nodeDate": cycle.node_date.isoformat(),
        "nodeKind": cycle.kind,
        "actualWindow": [actual_start, actual_end],
        "returnWindow": [return_start, return_end],
        "actualApiCount": len(actual_rows),
        "returnApiCount": len(return_rows),
        "insertedMonthColumn": prepared.inserted,
        "formulaCount": prepared.formula_count,
        "manualReviewRows": len(prepared.review),
        "criticalFailures": len(prepared.critical_failures),
        "dryRun": bool(args.dry_run),
        "workbook": str(workbook),
        "backup": None,
        "validation": None,
    }
    save_json(report_dir / "summary.json", summary)

    if prepared.critical_failures:
        print(json.dumps(summary, ensure_ascii=False))
        return 2
    if args.dry_run:
        print(json.dumps(summary, ensure_ascii=False))
        return 0

    candidate = create_same_directory_temp(workbook)
    try:
        fast_patch_zip(workbook, candidate, prepared.replacements)
        summary["validation"] = validate_workbook(workbook, candidate, cycle)
        backup = atomic_replace_master(workbook, candidate)
        summary["backup"] = str(backup)
    except BaseException:
        summary["candidate"] = str(candidate)
        save_json(report_dir / "summary.json", summary)
        raise
    save_json(report_dir / "summary.json", summary)
    print(json.dumps(summary, ensure_ascii=False))
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Update the ERP master workbook for the nearest monthly sync node.")
    parser.add_argument("--config", type=Path, default=Path(__file__).with_name("config.json"), help="JSON run configuration")
    parser.add_argument("--workbook", type=Path, help="Master XLSX workbook updated after validation")
    parser.add_argument("--snapshot-dir", type=Path)
    parser.add_argument("--actual-json", type=Path, help="Offline ERP response for the actual-quantity window")
    parser.add_argument("--return-json", type=Path, help="Offline ERP response for the return-rate window")
    parser.add_argument("--company-id")
    parser.add_argument("--cookie-env")
    parser.add_argument(
        "--browser-login",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="Open a persistent Chrome ERP profile and use its signed-in API session",
    )
    parser.add_argument("--aliases", type=Path)
    parser.add_argument("--skus-file", type=Path, help="Critical SKU validation list")
    parser.add_argument("--report-dir", type=Path)
    parser.add_argument(
        "--dry-run",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="Fetch, compare and review without creating XLSX",
    )
    return parser


RUN_DEFAULTS: dict[str, Any] = {
    "company_id": "111873",
    "cookie_env": "ERP_COOKIE",
    "browser_login": True,
    "snapshot_dir": Path("snapshots"),
    "aliases": Path("sku_aliases.json"),
    "report_dir": Path("reports"),
    "dry_run": False,
}
CONFIG_PATH_FIELDS = {
    "workbook", "snapshot_dir", "actual_json", "return_json", "aliases", "skus_file", "report_dir",
}
CONFIG_FIELDS = {
    "workbook", "snapshot_dir", "actual_json", "return_json", "company_id", "cookie_env",
    "browser_login", "aliases", "skus_file", "report_dir", "dry_run",
}
LEGACY_CONFIG_FIELDS = {
    "input", "output", "start", "end", "start_date", "end_date",
    "api_json", "previous_json", "snapshot",
}
LEGACY_CLI_FLAGS = ("--input", "--output", "--start", "--end", "--api-json")


def load_runtime_config(path: Path) -> dict[str, Any]:
    document = load_json(path)
    if not isinstance(document, dict):
        raise SystemExit(f"Config must contain a JSON object: {path}")
    runtime_document = {key: value for key, value in document.items() if not key.startswith("_")}
    legacy = sorted(set(runtime_document) & LEGACY_CONFIG_FIELDS)
    if legacy:
        raise SystemExit(
            "legacy configuration is no longer supported; set workbook and let the script choose both date windows automatically"
        )
    unknown = sorted(set(runtime_document) - CONFIG_FIELDS)
    if unknown:
        raise SystemExit(f"Unknown config fields: {', '.join(unknown)}")

    values: dict[str, Any] = {}
    config_dir = path.resolve().parent
    for source_key, value in runtime_document.items():
        key = source_key
        if key in CONFIG_PATH_FIELDS and value not in (None, ""):
            configured_path = Path(value)
            value = configured_path if configured_path.is_absolute() else config_dir / configured_path
        values[key] = value
    return values


def parse_runtime_args(argv: list[str] | None = None) -> argparse.Namespace:
    argv_list = list(argv) if argv is not None else sys.argv[1:]
    for item in argv_list:
        if any(item == flag or item.startswith(flag + "=") for flag in LEGACY_CLI_FLAGS):
            raise SystemExit(
                "legacy command-line dates/input/output are no longer supported; use --workbook"
            )
    config_parser = argparse.ArgumentParser(add_help=False)
    config_parser.add_argument("--config", type=Path, default=Path(__file__).with_name("config.json"))
    config_args, _ = config_parser.parse_known_args(argv_list)
    explicit_config = any(item == "--config" or item.startswith("--config=") for item in argv_list)

    config_values: dict[str, Any] = {}
    if config_args.config.exists():
        config_values = load_runtime_config(config_args.config)
    elif explicit_config:
        raise SystemExit(f"Config file not found: {config_args.config}")

    parser = build_parser()
    defaults = {**RUN_DEFAULTS, **config_values}
    parser.set_defaults(**defaults)
    return parser.parse_args(argv_list)


def main(argv: list[str] | None = None) -> int:
    args = parse_runtime_args(argv)
    return run_monthly_sync(args)


if __name__ == "__main__":
    raise SystemExit(main())
