import argparse
import csv
from datetime import datetime, time, timedelta, timezone
from dataclasses import dataclass
from difflib import SequenceMatcher
import getpass
import json
import os
from pathlib import Path
import re
import struct
import sys
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qsl, urlencode
from urllib.request import Request, urlopen
import zlib
import zipfile
from xml.etree import ElementTree as ET

from xlsx_monthly import column_number


SHANGHAI_TZ = timezone(timedelta(hours=8))
API_URL = "https://erpa.superboss.cc/report/sale/dimensions/list"
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


def _raw_local_record(data: bytes, info: zipfile.ZipInfo) -> bytes:
    start = info.header_offset
    if data[start : start + 4] != b"PK\x03\x04":
        raise RuntimeError(f"Unexpected ZIP local header for {info.filename}")
    name_length, extra_length = struct.unpack_from("<HH", data, start + 26)
    end = start + 30 + name_length + extra_length + info.compress_size
    return data[start:end]


def fast_patch_zip(source: Path, output: Path, modified_sheet: bytes) -> None:
    original_data = source.read_bytes()
    with zipfile.ZipFile(source, "r") as archive:
        infos = archive.infolist()
        target = next(info for info in infos if info.filename == WORKSHEET_PATH)
        central_directory_start = archive.start_dir
        old_record = _raw_local_record(original_data, target)
        compressor = zlib.compressobj(level=6, wbits=-15)
        compressed_sheet = compressor.compress(modified_sheet) + compressor.flush()
        crc = zlib.crc32(modified_sheet) & 0xFFFFFFFF
        name_length, extra_length = struct.unpack_from("<HH", original_data, target.header_offset + 26)
        header_length = 30 + name_length + extra_length
        local_header = bytearray(original_data[target.header_offset : target.header_offset + header_length])
        struct.pack_into("<III", local_header, 14, crc, len(compressed_sheet), len(modified_sheet))

        output.parent.mkdir(parents=True, exist_ok=True)
        new_offsets: dict[str, int] = {}
        offset = 0
        with output.open("wb") as destination:
            for info in infos:
                new_offsets[info.filename] = offset
                if info.filename == WORKSHEET_PATH:
                    destination.write(local_header)
                    destination.write(compressed_sheet)
                    offset += len(local_header) + len(compressed_sheet)
                else:
                    record = _raw_local_record(original_data, info)
                    destination.write(record)
                    offset += len(record)

            new_central_start = offset
            position = central_directory_start
            for info in infos:
                if original_data[position : position + 4] != b"PK\x01\x02":
                    raise RuntimeError("ZIP central directory signature not found")
                name_length, extra_length, comment_length = struct.unpack_from("<HHH", original_data, position + 28)
                record_length = 46 + name_length + extra_length + comment_length
                record = bytearray(original_data[position : position + record_length])
                struct.pack_into(
                    "<III",
                    record,
                    16,
                    crc if info.filename == WORKSHEET_PATH else info.CRC,
                    len(compressed_sheet) if info.filename == WORKSHEET_PATH else info.compress_size,
                    len(modified_sheet) if info.filename == WORKSHEET_PATH else info.file_size,
                )
                struct.pack_into("<I", record, 42, new_offsets[info.filename])
                destination.write(record)
                position += record_length

            central_size = destination.tell() - new_central_start
            end_record_position = original_data.rfind(b"PK\x05\x06", central_directory_start)
            if end_record_position < 0:
                raise RuntimeError("ZIP end record not found")
            end_record = bytearray(original_data[end_record_position : end_record_position + 22])
            struct.pack_into("<II", end_record, 12, central_size, new_central_start)
            destination.write(end_record)


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
        fast_patch_zip(source, output, modified_sheet)
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


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Fetch ERP sales data and quickly patch an existing XLSX workbook.")
    parser.add_argument("--config", type=Path, default=Path(__file__).with_name("config.json"), help="JSON run configuration")
    parser.add_argument("--input", type=Path, help="Source XLSX workbook")
    parser.add_argument("--output", type=Path, help="Output XLSX workbook; omitted in --dry-run mode")
    parser.add_argument("--start", help="Start date, YYYY-MM-DD")
    parser.add_argument("--end", help="End date, YYYY-MM-DD")
    parser.add_argument("--company-id")
    parser.add_argument("--cookie-env")
    parser.add_argument("--api-json", type=Path, help="Use an existing API JSON instead of making a request")
    parser.add_argument("--previous-json", type=Path)
    parser.add_argument("--snapshot", type=Path)
    parser.add_argument("--aliases", type=Path)
    parser.add_argument("--skus-file", type=Path, help="Optional text file containing one sheet SKU per line")
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
    "previous_json": Path("erp_dimensions.json"),
    "snapshot": Path("erp_dimensions_latest.json"),
    "aliases": Path("sku_aliases.json"),
    "report_dir": Path("reports"),
    "dry_run": False,
}
CONFIG_KEY_ALIASES = {"start_date": "start", "end_date": "end"}
CONFIG_PATH_FIELDS = {
    "input", "output", "api_json", "previous_json", "snapshot", "aliases", "skus_file", "report_dir",
}
CONFIG_FIELDS = {
    "input", "output", "start", "end", "start_date", "end_date", "company_id", "cookie_env",
    "api_json", "previous_json", "snapshot", "aliases", "skus_file", "report_dir", "dry_run",
}


def load_runtime_config(path: Path) -> dict[str, Any]:
    document = load_json(path)
    if not isinstance(document, dict):
        raise SystemExit(f"Config must contain a JSON object: {path}")
    runtime_document = {key: value for key, value in document.items() if not key.startswith("_")}
    unknown = sorted(set(runtime_document) - CONFIG_FIELDS)
    if unknown:
        raise SystemExit(f"Unknown config fields: {', '.join(unknown)}")

    values: dict[str, Any] = {}
    config_dir = path.resolve().parent
    for source_key, value in runtime_document.items():
        key = CONFIG_KEY_ALIASES.get(source_key, source_key)
        if key in CONFIG_PATH_FIELDS and value not in (None, ""):
            configured_path = Path(value)
            value = configured_path if configured_path.is_absolute() else config_dir / configured_path
        values[key] = value
    return values


def parse_runtime_args(argv: list[str] | None = None) -> argparse.Namespace:
    argv_list = list(argv) if argv is not None else sys.argv[1:]
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
    if args.input is None:
        raise SystemExit("Input workbook is required; set input in config.json or pass --input")
    if not args.start or not args.end:
        raise SystemExit("Start and end dates are required; set start_date/end_date in config.json or pass --start/--end")
    if not args.input.exists():
        raise SystemExit(f"Input workbook not found: {args.input}")
    if not args.dry_run and args.output is None:
        raise SystemExit("--output is required unless --dry-run is used")

    previous_document = load_json(args.previous_json) if args.previous_json.exists() else None
    if args.api_json:
        latest_document = load_json(args.api_json)
    else:
        cookie = os.environ.get(args.cookie_env)
        if not cookie:
            cookie = getpass.getpass(f"{args.cookie_env} is not set; paste ERP Cookie: ")
        if not cookie.strip():
            raise SystemExit("ERP Cookie is required")
        latest_document = fetch_api(cookie, args.company_id, args.start, args.end)
        save_json(args.snapshot, latest_document)

    latest_rows = api_rows(latest_document)
    aliases = load_aliases(args.aliases)
    targets = load_target_skus(args.skus_file)
    report = sync_workbook(args.input, None if args.dry_run else args.output, latest_rows, aliases, targets)

    api_differences = (
        compare_api_snapshots(api_rows(previous_document), latest_rows)
        if previous_document is not None
        else []
    )
    write_csv(
        args.report_dir / "manual_review.csv",
        report["review"],
        ["row", "sheet_sku", "sheet_name", "status", "confidence", "candidate_sku", "candidate_title", "different_columns"],
    )
    write_csv(
        args.report_dir / "sync_changes.csv",
        report["rows"],
        ["row", "sheet_sku", "sheet_name", "match_status", "candidate_sku", "old_actual", "new_actual", "old_return_rate", "new_return_rate", "changed_columns"],
    )
    write_csv(
        args.report_dir / "api_changes.csv",
        flatten_api_differences(api_differences),
        ["status", "itemOuterId", "title", "field", "old", "new"],
    )
    summary = {
        "apiCount": len(latest_rows),
        "writtenRows": report["written"],
        "unchangedRows": report["unchanged"],
        "manualReviewRows": len(report["review"]),
        "apiChangedItems": len(api_differences),
        "output": str(args.output) if args.output else None,
        "snapshot": str(args.snapshot) if not args.api_json else str(args.api_json),
        "reportDir": str(args.report_dir),
    }
    save_json(args.report_dir / "summary.json", summary)
    print(json.dumps(summary, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
