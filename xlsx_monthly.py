from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
import html
import math
import re
from typing import Any, Callable

from monthly_schedule import (
    SyncCycle,
    actual_header,
    change_formula,
    peer_formula,
    previous_month_start,
    return_header,
)


MAX_EXCEL_COLUMN = 16384
CELL_RE = re.compile(
    rb'(<c\s+[^>]*?\br="(?P<col2>[A-Z]+)(?P<row2>\d+)"[^>]*/>|'
    rb'<c\s+[^>]*?\br="(?P<col>[A-Z]+)(?P<row>\d+)"[^>]*>(?P<body>.*?)</c>)',
    re.S,
)
ROW_RE = re.compile(rb'<row\b[^>]*\br="(?P<row>\d+)"[^>]*>.*?</row>', re.S)
CELL_REFERENCE_RE = re.compile(
    r"(?<![A-Za-z0-9_.])(?P<absolute>\$?)(?P<column>[A-Za-z]{1,3})(?P<row>\$?\d+)(?![A-Za-z0-9_(])"
)
COLUMN_RANGE_RE = re.compile(
    r"(?<![A-Za-z0-9_.])(?P<absolute1>\$?)(?P<column1>[A-Za-z]{1,3})"
    r":(?P<absolute2>\$?)(?P<column2>[A-Za-z]{1,3})(?![A-Za-z0-9_(])"
)
QUALIFIED_REFERENCE_RE = re.compile(
    r"(?P<sheet>'(?:[^']|'')+'|[^'\"!,()+\-*/=\s]+)!"
    r"(?P<reference>\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|"
    r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})"
)
MONTH_HEADER_RE = re.compile(r"^(?P<month>1[0-2]|[1-9])月实发(?:\([^)]*\))?$")


@dataclass(frozen=True)
class WorkbookLayout:
    header_row: int
    previous_col: str
    peer_col: str
    actual_col: str
    change_col: str
    before_return_col: str
    after_return_col: str
    return_col: str
    month_cols: tuple[str, ...]
    needs_insert: bool
    insert_before_col: str | None


@dataclass(frozen=True)
class MonthInsertionResult:
    sheet_xml: bytes
    layout: WorkbookLayout
    inserted: bool


@dataclass(frozen=True)
class MonthlyWriteResult:
    sheet_xml: bytes
    layout: WorkbookLayout
    inserted: bool
    rows: tuple[dict[str, Any], ...]
    review: tuple[dict[str, Any], ...]
    critical_results: tuple[dict[str, Any], ...]
    critical_failures: tuple[dict[str, Any], ...]
    formula_count: int


def column_number(col: str) -> int:
    if not isinstance(col, str) or not re.fullmatch(r"[A-Z]+", col):
        raise ValueError(f"invalid Excel column label: {col!r}")
    number = 0
    for char in col:
        number = number * 26 + ord(char) - ord("A") + 1
    if number > MAX_EXCEL_COLUMN:
        raise ValueError(f"Excel column label is out of range: {col!r}")
    return number


def column_label(number: int) -> str:
    if isinstance(number, bool) or not isinstance(number, int) or not 1 <= number <= MAX_EXCEL_COLUMN:
        raise ValueError(f"invalid Excel column number: {number!r}")
    chars: list[str] = []
    while number:
        number, remainder = divmod(number - 1, 26)
        chars.append(chr(ord("A") + remainder))
    return "".join(reversed(chars))


def normalize_header(value: str) -> str:
    normalized = re.sub(r"\s+", "", str(value or ""))
    normalized = normalized.translate(str.maketrans({"（": "(", "）": ")"}))
    return normalized.replace("实发销量", "实发").replace("实销", "实发")


def _cell_col(match: re.Match[bytes]) -> str:
    return (match.group("col") or match.group("col2")).decode("ascii")


def _cell_row(match: re.Match[bytes]) -> int:
    return int(match.group("row") or match.group("row2"))


def _cell_text(match: re.Match[bytes], shared_strings: list[str]) -> str:
    attrs = match.group(0).split(b">", 1)[0]
    body = match.group("body") or b""
    if re.search(rb'\bt="s"', attrs):
        value = re.search(rb"<v>(.*?)</v>", body, re.S)
        if value:
            try:
                return shared_strings[int(value.group(1))]
            except (ValueError, IndexError):
                return ""
    if re.search(rb'\bt="inlineStr"', attrs):
        pieces = re.findall(rb"<t(?:\s[^>]*)?>(.*?)</t>", body, re.S)
        return html.unescape(b"".join(pieces).decode("utf-8", "ignore"))
    value = re.search(rb"<v>(.*?)</v>", body, re.S)
    return html.unescape(value.group(1).decode("utf-8", "ignore")) if value else ""


def discover_layout(
    sheet_xml: bytes,
    shared_strings: list[str],
    cycle: SyncCycle,
) -> WorkbookLayout:
    cells: list[tuple[str, int, str]] = []
    for match in CELL_RE.finditer(sheet_xml):
        text = normalize_header(_cell_text(match, shared_strings))
        if (
            text in {"同期销量", "变化情况"}
            or "退货率" in text
            or MONTH_HEADER_RE.fullmatch(text)
        ):
            cells.append((_cell_col(match), _cell_row(match), text))

    def unique_header(label: str, predicate: Callable[[str], bool]) -> tuple[str, int]:
        matches = [(col, row) for col, row, text in cells if predicate(text)]
        if len(matches) != 1:
            raise ValueError(f"expected exactly one {label} header, found {len(matches)}")
        return matches[0]

    peer_col, peer_row = unique_header("同期销量", lambda text: text == "同期销量")
    change_col, change_row = unique_header("变化情况", lambda text: text == "变化情况")
    before_return_col, before_return_row = unique_header(
        "发货前退货率",
        lambda text: text == "发货前退货率" or text.startswith("发货前退货率("),
    )
    after_return_col, after_return_row = unique_header(
        "发货后退货率",
        lambda text: text == "发货后退货率" or text.startswith("发货后退货率("),
    )
    return_col, return_row = unique_header(
        "退货率",
        lambda text: text == "退货率" or text.startswith("退货率("),
    )
    if len({
        peer_row,
        change_row,
        before_return_row,
        after_return_row,
        return_row,
    }) != 1:
        raise ValueError("core headers are not on the same row")
    header_row = peer_row

    month_headers: list[tuple[str, int]] = []
    for col, row, text in cells:
        match = MONTH_HEADER_RE.fullmatch(text)
        if row == header_row and match:
            month_headers.append((col, int(match.group("month"))))
    month_headers.sort(key=lambda item: column_number(item[0]))
    if not month_headers:
        raise ValueError("no monthly actual-sales headers found")

    peer_number = column_number(peer_col)
    change_number = column_number(change_col)
    previous = [item for item in month_headers if column_number(item[0]) < peer_number]
    actual = [
        item
        for item in month_headers
        if peer_number < column_number(item[0]) < change_number
    ]
    if not previous:
        raise ValueError("missing previous-month actual-sales header")
    if len(actual) != 1:
        raise ValueError(f"expected exactly one current actual-sales header, found {len(actual)}")
    previous_col = previous[-1][0]
    actual_col, actual_month = actual[0]
    ordered_columns = (
        column_number(previous_col),
        peer_number,
        column_number(actual_col),
        change_number,
        column_number(before_return_col),
        column_number(after_return_col),
        column_number(return_col),
    )
    if ordered_columns != tuple(sorted(ordered_columns)) or len(set(ordered_columns)) != 7:
        raise ValueError(
            "invalid header column order; expected previous < peer < actual < change "
            "< before return < after return < return"
        )
    if (
        column_number(before_return_col),
        column_number(after_return_col),
        column_number(return_col),
    ) != (change_number + 1, change_number + 2, change_number + 3):
        raise ValueError(
            "expected 发货前退货率 < 发货后退货率 < 退货率 immediately after 变化情况"
        )

    expected_month = cycle.actual_month.month
    if cycle.kind == "first":
        if actual_month != expected_month:
            raise ValueError(
                f"first-node actual header is month {actual_month}, expected {expected_month}"
            )
        needs_insert = False
    elif actual_month == expected_month:
        needs_insert = False
    elif actual_month == cycle.previous_month.month:
        needs_insert = True
    else:
        raise ValueError(
            f"fifteenth-node actual header is month {actual_month}, expected "
            f"{cycle.previous_month.month} or {expected_month}"
        )

    previous_month = previous[-1][1]
    expected_previous_month = (
        previous_month_start(cycle.previous_month).month
        if needs_insert
        else cycle.previous_month.month
    )
    if previous_month != expected_previous_month:
        raise ValueError(
            f"previous-month header is month {previous_month}, "
            f"expected {expected_previous_month}"
        )

    return WorkbookLayout(
        header_row=header_row,
        previous_col=previous_col,
        peer_col=peer_col,
        actual_col=actual_col,
        change_col=change_col,
        before_return_col=before_return_col,
        after_return_col=after_return_col,
        return_col=return_col,
        month_cols=tuple(col for col, _month in month_headers),
        needs_insert=needs_insert,
        insert_before_col=peer_col if needs_insert else None,
    )


def _shift_reference_segment(segment: str, insert_number: int) -> str:
    def shift_column(label: str) -> str:
        number = column_number(label.upper())
        if number < insert_number:
            return label
        shifted = column_label(number + 1)
        return shifted.lower() if label.islower() else shifted

    def replace_column_range(match: re.Match[str]) -> str:
        try:
            left = shift_column(match.group("column1"))
            right = shift_column(match.group("column2"))
        except ValueError:
            return match.group(0)
        return (
            match.group("absolute1")
            + left
            + ":"
            + match.group("absolute2")
            + right
        )

    def replace(match: re.Match[str]) -> str:
        try:
            shifted = shift_column(match.group("column"))
        except ValueError:
            return match.group(0)
        return match.group("absolute") + shifted + match.group("row")

    return CELL_REFERENCE_RE.sub(replace, COLUMN_RANGE_RE.sub(replace_column_range, segment))


def _decoded_sheet_name(encoded_sheet: str) -> str:
    if encoded_sheet.startswith("'"):
        return encoded_sheet[1:-1].replace("''", "'")
    return encoded_sheet


def _shift_formula_segment(
    segment: str,
    insert_number: int,
    target_sheet: str,
    shift_unqualified: bool,
) -> str:
    protected: list[str] = []

    def protect(value: str) -> str:
        token = f"\ue000{len(protected)}\ue001"
        protected.append(value)
        return token

    def replace_qualified(match: re.Match[str]) -> str:
        reference = match.group("reference")
        if _decoded_sheet_name(match.group("sheet")) == target_sheet:
            reference = _shift_reference_segment(reference, insert_number)
        return protect(match.group("sheet") + "!" + reference)

    masked = QUALIFIED_REFERENCE_RE.sub(replace_qualified, segment)
    masked = re.sub(r"'(?:[^']|'')*'", lambda match: protect(match.group(0)), masked)
    if shift_unqualified:
        masked = _shift_reference_segment(masked, insert_number)
    for index, value in enumerate(protected):
        masked = masked.replace(f"\ue000{index}\ue001", value)
    return masked


def shift_formula_references(
    formula: str,
    insert_before_col: str | int,
    target_sheet: str = "分级总表",
    *,
    shift_unqualified: bool = True,
) -> str:
    insert_number = (
        column_number(insert_before_col)
        if isinstance(insert_before_col, str)
        else insert_before_col
    )
    if not isinstance(insert_number, int) or not 1 <= insert_number <= MAX_EXCEL_COLUMN:
        raise ValueError(f"invalid insertion column: {insert_before_col!r}")

    output: list[str] = []
    start = 0
    index = 0
    while index < len(formula):
        quote = formula[index]
        if quote != '"':
            index += 1
            continue
        output.append(
            _shift_formula_segment(
                formula[start:index],
                insert_number,
                target_sheet,
                shift_unqualified,
            )
        )
        end = index + 1
        while end < len(formula):
            if formula[end] != quote:
                end += 1
                continue
            if end + 1 < len(formula) and formula[end + 1] == quote:
                end += 2
                continue
            end += 1
            break
        output.append(formula[index:end])
        index = end
        start = end
    output.append(
        _shift_formula_segment(
            formula[start:],
            insert_number,
            target_sheet,
            shift_unqualified,
        )
    )
    return "".join(output)


def insert_month_column(
    sheet_xml: bytes,
    shared_strings: list[str],
    cycle: SyncCycle,
) -> MonthInsertionResult:
    layout = discover_layout(sheet_xml, shared_strings, cycle)
    if not layout.needs_insert:
        return MonthInsertionResult(sheet_xml=sheet_xml, layout=layout, inserted=False)

    if layout.insert_before_col is None:
        raise ValueError("insertion layout has no insertion column")
    insert_number = column_number(layout.insert_before_col)
    old_actual_number = column_number(layout.actual_col)
    shifted_actual_col = column_label(old_actual_number + 1)
    inserted_col = column_label(insert_number)

    def shift_cell_coordinate(match: re.Match[bytes]) -> bytes:
        number = column_number(match.group(2).decode("ascii"))
        col = column_label(number + 1) if number >= insert_number else column_label(number)
        return match.group(1) + col.encode("ascii") + match.group(3) + match.group(4)

    patched = re.sub(
        rb'(<c\b[^>]*\br=")([A-Z]+)(\d+)(")',
        shift_cell_coordinate,
        sheet_xml,
    )

    def shift_row_span(match: re.Match[bytes]) -> bytes:
        start, end = int(match.group(2)), int(match.group(3))
        if start > insert_number:
            start += 1
        if end >= insert_number:
            end += 1
        return (
            match.group(1)
            + str(start).encode("ascii")
            + b":"
            + str(end).encode("ascii")
            + match.group(4)
        )

    patched = re.sub(rb'(\bspans=")(\d+):(\d+)(")', shift_row_span, patched)

    def shift_formula(match: re.Match[bytes]) -> bytes:
        opening, body, closing = match.group(1), match.group(2), match.group(3)
        formula = html.unescape(body.decode("utf-8", "ignore"))
        shifted = shift_formula_references(formula, insert_number)
        return opening + html.escape(shifted, quote=False).encode("utf-8") + closing

    formula_tag = rb'(?:[A-Za-z_][\w.-]*:)?(?:formula[12]?|f)'
    patched = re.sub(
        rb'(<'+ formula_tag + rb'\b(?![^>]*?/\s*>)[^>]*>)(.*?)(</' + formula_tag + rb'>)',
        shift_formula,
        patched,
        flags=re.S,
    )

    def shift_reference_attribute(match: re.Match[bytes]) -> bytes:
        value = html.unescape(match.group(2).decode("utf-8", "ignore"))
        shifted = shift_formula_references(value, insert_number)
        return match.group(1) + html.escape(shifted, quote=True).encode("utf-8") + match.group(3)

    patched = re.sub(
        rb'((?:\bref|\bsqref|\bactiveCell|\btopLeftCell)=")([^"]*)(")',
        shift_reference_attribute,
        patched,
    )

    def roll_row(match: re.Match[bytes]) -> bytes:
        row = match.group(0)
        row_number = int(match.group("row"))
        actual_match = next(
            (
                cell
                for cell in CELL_RE.finditer(row)
                if _cell_col(cell) == shifted_actual_col
            ),
            None,
        )
        if actual_match is not None:
            copied = actual_match.group(0)
            copied = re.sub(rb'<f\b[^>]*/>', b"", copied)
            copied = re.sub(rb'<f\b[^>]*>.*?</f>', b"", copied, flags=re.S)
            copied = re.sub(
                rb'(\br=")' + re.escape(shifted_actual_col.encode("ascii")) + rb'(\d+")',
                rb'\g<1>' + inserted_col.encode("ascii") + rb'\g<2>',
                copied,
                count=1,
            )
            row = _insert_or_replace_cell(row, inserted_col, row_number, copied)

        if row_number == layout.header_row:
            row = _write_inline_header(
                row,
                inserted_col,
                row_number,
                f"{cycle.previous_month.month}月实发",
            )
            row = _write_inline_header(
                row,
                shifted_actual_col,
                row_number,
                actual_header(cycle),
            )
        return row

    patched = ROW_RE.sub(roll_row, patched)
    patched = _rewrite_column_definitions(patched, layout, insert_number, old_actual_number)
    new_layout = discover_layout(patched, shared_strings, cycle)
    return MonthInsertionResult(sheet_xml=patched, layout=new_layout, inserted=True)


def _insert_or_replace_cell(row: bytes, col: str, row_number: int, cell_xml: bytes) -> bytes:
    target = column_number(col)
    for match in CELL_RE.finditer(row):
        existing = column_number(_cell_col(match))
        if existing == target:
            return row[: match.start()] + cell_xml + row[match.end() :]
        if existing > target:
            return row[: match.start()] + cell_xml + row[match.start() :]
    insertion = row.rfind(b"</row>")
    if insertion < 0:
        raise ValueError(f"row {row_number} has no closing tag")
    return row[:insertion] + cell_xml + row[insertion:]


def _write_inline_header(
    row: bytes,
    col: str,
    row_number: int,
    value: str,
) -> bytes:
    existing = next(
        (
            match
            for match in CELL_RE.finditer(row)
            if _cell_col(match) == col
        ),
        None,
    )
    if existing is None:
        opening = f'<c r="{col}{row_number}"'.encode("ascii")
    else:
        opening = existing.group(0).split(b">", 1)[0].rstrip(b"/")
        opening = re.sub(rb'\s+t="[^"]*"', b"", opening)
    escaped = html.escape(value, quote=False).encode("utf-8")
    cell = opening + b' t="inlineStr"><is><t>' + escaped + b"</t></is></c>"
    return _insert_or_replace_cell(row, col, row_number, cell)


def _normalize_sku(value: Any) -> str:
    text = str(value or "").casefold()
    text = re.sub(r"[‐‑‒–—―]", "-", text)
    return re.sub(r"\s+", "", text)


def _rate_number(value: Any) -> float:
    text = str(value or "0").strip()
    if text.endswith("%"):
        text = text[:-1].strip()
    number = float(text or 0)
    if not math.isfinite(number):
        raise ValueError("return rate must be finite")
    return number / 100.0


def _number_text(value: float) -> str:
    return ("%.10f" % value).rstrip("0").rstrip(".") or "0"


def _cell_style(row: bytes, col: str) -> str | None:
    existing = next(
        (cell for cell in CELL_RE.finditer(row) if _cell_col(cell) == col),
        None,
    )
    if existing is None:
        return None
    style = re.search(rb'\bs="([^"]+)"', existing.group(0).split(b">", 1)[0])
    return style.group(1).decode("ascii") if style else None


def _apply_cell_style(opening: bytes, style: str | None) -> bytes:
    if style is None:
        return opening
    opening = re.sub(rb'\s+s="[^"]*"', b"", opening)
    return opening + b' s="' + style.encode("ascii") + b'"'


def _set_cell_number(
    row: bytes,
    col: str,
    row_number: int,
    value: str,
    *,
    style: str | None = None,
) -> bytes:
    existing = next(
        (cell for cell in CELL_RE.finditer(row) if _cell_col(cell) == col),
        None,
    )
    if existing is None:
        opening = f'<c r="{col}{row_number}"'.encode("ascii")
    else:
        opening = existing.group(0).split(b">", 1)[0].rstrip(b"/")
        opening = re.sub(rb'\s+t="[^"]*"', b"", opening)
        opening = re.sub(rb'\s+vm="[^"]*"', b"", opening)
    opening = _apply_cell_style(opening, style)
    cell = opening + b"><v>" + value.encode("ascii") + b"</v></c>"
    return _insert_or_replace_cell(row, col, row_number, cell)


def _clear_cell_value(
    row: bytes,
    col: str,
    row_number: int,
    *,
    style: str | None = None,
) -> bytes:
    existing = next(
        (cell for cell in CELL_RE.finditer(row) if _cell_col(cell) == col),
        None,
    )
    if existing is None:
        opening = f'<c r="{col}{row_number}"'.encode("ascii")
    else:
        opening = existing.group(0).split(b">", 1)[0].rstrip(b"/")
        opening = re.sub(rb'\s+t="[^"]*"', b"", opening)
        opening = re.sub(rb'\s+vm="[^"]*"', b"", opening)
    opening = _apply_cell_style(opening, style)
    return _insert_or_replace_cell(
        row, col, row_number, opening + b"/>",
    )


def _set_cell_formula(
    row: bytes,
    col: str,
    row_number: int,
    formula: str,
    default_style: str | None = None,
) -> bytes:
    existing = next(
        (cell for cell in CELL_RE.finditer(row) if _cell_col(cell) == col),
        None,
    )
    if existing is None:
        style = "" if default_style is None else f' s="{default_style}"'
        opening = f'<c r="{col}{row_number}"{style}'.encode("ascii")
    else:
        opening = existing.group(0).split(b">", 1)[0].rstrip(b"/")
        opening = re.sub(rb'\s+t="[^"]*"', b"", opening)
        opening = re.sub(rb'\s+vm="[^"]*"', b"", opening)
    encoded = html.escape(formula, quote=False).encode("utf-8")
    cell = opening + b"><f>" + encoded + b"</f></c>"
    return _insert_or_replace_cell(row, col, row_number, cell)


def _formula_column_style(sheet_xml: bytes, col: str) -> str | None:
    styles: list[str] = []
    formula_tag = rb'(?:[A-Za-z_][\w.-]*:)?(?:formula[12]?|f)'
    for cell in CELL_RE.finditer(sheet_xml):
        if _cell_col(cell) != col:
            continue
        if not re.search(rb'<'+ formula_tag + rb'\b', cell.group("body") or b""):
            continue
        style = re.search(rb'\bs="(\d+)"', cell.group(0).split(b">", 1)[0])
        if style:
            styles.append(style.group(1).decode("ascii"))
    return Counter(styles).most_common(1)[0][0] if styles else None


def _column_style(sheet_xml: bytes, col: str) -> str | None:
    styles: list[str] = []
    for cell in CELL_RE.finditer(sheet_xml):
        if _cell_col(cell) != col:
            continue
        if re.search(rb'\bt="(?:inlineStr|s)"', cell.group(0).split(b">", 1)[0]):
            continue
        style = re.search(rb'\bs="(\d+)"', cell.group(0).split(b">", 1)[0])
        if style:
            styles.append(style.group(1).decode("ascii"))
    if not styles:
        return None
    style, count = Counter(styles).most_common(1)[0]
    return style if count > len(styles) / 2 else None


def _extend_change_conditional_formatting(
    sheet_xml: bytes,
    change_col: str,
    formula_rows: set[int],
) -> bytes:
    if not formula_rows:
        return sheet_xml
    wrapper_re = re.compile(
        rb'(?P<open><(?:[A-Za-z_][\w.-]*:)?conditionalFormatting\b[^>]*>)'
        rb'(?P<body>.*?</(?:[A-Za-z_][\w.-]*:)?conditionalFormatting>)',
        re.S,
    )
    contains_text_re = re.compile(
        rb'<(?:[A-Za-z_][\w.-]*:)?cfRule\b(?=[^>]*\btype="containsText")[^>]*>'
    )
    target_pattern = re.compile(
        rf"^\$?{re.escape(change_col)}\$?(\d+)"
        rf"(?::\$?{re.escape(change_col)}\$?(\d+))?$"
    )
    candidates_by_rule: dict[
        tuple[str, ...], list[tuple[int, int, list[tuple[int, int]]]]
    ] = {}
    for candidate in wrapper_re.finditer(sheet_xml):
        rule = contains_text_re.search(candidate.group("body"))
        if rule is None:
            continue
        sqref_match = re.search(rb'\bsqref="([^"]*)"', candidate.group("open"))
        if sqref_match is None:
            continue
        intervals = []
        for token in html.unescape(
            sqref_match.group(1).decode("utf-8", "ignore")
        ).split():
            target = target_pattern.fullmatch(token)
            if target:
                start = int(target.group(1))
                end = int(target.group(2) or start)
                intervals.append((min(start, end), max(start, end)))
        if not intervals:
            continue
        attrs = _tag_attributes(rule.group(0))
        signature = tuple(
            attrs.get(key, "")
            for key in ("type", "dxfId", "operator", "text")
        )
        coverage = sum(end - start + 1 for start, end in intervals)
        candidates_by_rule.setdefault(signature, []).append(
            (coverage, candidate.start(), intervals)
        )
    selected_starts: set[int] = set()
    covered_elsewhere: dict[int, set[int]] = {}
    for candidates in candidates_by_rule.values():
        _coverage, selected_start, _intervals = max(
            candidates, key=lambda item: item[0]
        )
        selected_starts.add(selected_start)
        sibling_intervals = [
            interval
            for _size, start, intervals in candidates
            if start != selected_start
            for interval in intervals
        ]
        covered_elsewhere[selected_start] = {
            row
            for row in formula_rows
            if any(start <= row <= end for start, end in sibling_intervals)
        }

    def extend_wrapper(match: re.Match[bytes]) -> bytes:
        if match.start() not in selected_starts:
            return match.group(0)
        if not contains_text_re.search(match.group("body")):
            return match.group(0)
        opening = match.group("open")
        sqref_match = re.search(rb'\bsqref="([^"]*)"', opening)
        if sqref_match is None:
            return match.group(0)
        sqref = html.unescape(sqref_match.group(1).decode("utf-8", "ignore"))
        tokens = sqref.split()
        other_tokens: list[str] = []
        intervals: list[tuple[int, int]] = []
        for token in tokens:
            target = target_pattern.fullmatch(token)
            if target is None:
                other_tokens.append(token)
                continue
            start = int(target.group(1))
            end = int(target.group(2) or start)
            intervals.append((min(start, end), max(start, end)))
        if not intervals:
            return match.group(0)
        intervals.extend(
            (row, row)
            for row in formula_rows
            if row not in covered_elsewhere.get(match.start(), set())
        )
        merged: list[list[int]] = []
        for start, end in sorted(intervals):
            if merged and start <= merged[-1][1] + 1:
                merged[-1][1] = max(merged[-1][1], end)
            else:
                merged.append([start, end])
        target_tokens = [
            f"{change_col}{start}"
            if start == end
            else f"{change_col}{start}:{change_col}{end}"
            for start, end in merged
        ]
        updated_sqref = " ".join(other_tokens + target_tokens)
        escaped = html.escape(updated_sqref, quote=True).encode("utf-8")
        updated_opening = (
            opening[: sqref_match.start(1)]
            + escaped
            + opening[sqref_match.end(1) :]
        )
        return updated_opening + match.group("body")

    return wrapper_re.sub(extend_wrapper, sheet_xml)


def apply_monthly_values(
    sheet_xml: bytes,
    shared_strings: list[str],
    cycle: SyncCycle,
    actual_rows: list[dict[str, Any]],
    before_return_rows: list[dict[str, Any]],
    overall_return_rows: list[dict[str, Any]],
    aliases: dict[str, str],
    critical_skus: set[str],
    matcher: Callable[[Any, Any, list[dict[str, Any]], dict[str, str]], Any],
    after_return_rows: list[dict[str, Any]] | None = None,
) -> MonthlyWriteResult:
    insertion = insert_month_column(sheet_xml, shared_strings, cycle)
    layout = insertion.layout
    normalized_aliases = {
        _normalize_sku(source): _normalize_sku(target)
        for source, target in aliases.items()
    }
    rows_report: list[dict[str, Any]] = []
    review: list[dict[str, Any]] = []
    row_matches: dict[str, list[dict[str, Any]]] = {}
    formula_count = 0
    formula_rows: set[int] = set()
    peer_formula_style = _formula_column_style(
        insertion.sheet_xml, layout.peer_col
    )
    change_formula_style = _formula_column_style(
        insertion.sheet_xml, layout.change_col
    )
    dominant_return_percentage_style = _column_style(
        insertion.sheet_xml, layout.return_col
    )

    def update_row(match: re.Match[bytes]) -> bytes:
        nonlocal formula_count
        row_number = int(match.group("row"))
        row = match.group(0)
        if row_number == layout.header_row:
            row = _write_inline_header(
                row, layout.actual_col, row_number, actual_header(cycle)
            )
            row = _write_inline_header(
                row,
                layout.before_return_col,
                row_number,
                return_header(cycle, "发货前退货率"),
            )
            row = _write_inline_header(
                row,
                layout.after_return_col,
                row_number,
                return_header(cycle, "发货后退货率"),
            )
            return _write_inline_header(
                row, layout.return_col, row_number, return_header(cycle)
            )

        sheet_sku = _cell_value(row, "E", shared_strings)
        if not sheet_sku:
            return row
        sheet_name = _cell_value(row, "F", shared_strings)
        actual_match = matcher(
            sheet_sku, sheet_name, actual_rows, normalized_aliases
        )
        before_return_match = matcher(
            sheet_sku, sheet_name, before_return_rows, normalized_aliases
        )
        after_return_match = (
            matcher(
                sheet_sku,
                sheet_name,
                after_return_rows,
                normalized_aliases,
            )
            if after_return_rows is not None
            else before_return_match
        )
        overall_return_match = matcher(
            sheet_sku, sheet_name, overall_return_rows, normalized_aliases
        )
        old_actual = _cell_value(row, layout.actual_col, shared_strings)
        old_before_return = _cell_value(
            row, layout.before_return_col, shared_strings
        )
        old_after_return = _cell_value(
            row, layout.after_return_col, shared_strings
        )
        old_return = _cell_value(row, layout.return_col, shared_strings)
        new_actual = old_actual
        new_before_return = ""
        new_after_return = ""
        new_return = ""
        changed: list[str] = []
        actual_status = actual_match.status
        before_return_status = before_return_match.status
        after_return_status = after_return_match.status
        return_status = overall_return_match.status
        actual_auto_write = False
        before_return_match_ok = False
        after_return_match_ok = False
        return_match_ok = False

        if (
            actual_match.auto_write
            and actual_match.status in {"exact", "alias"}
            and actual_match.candidate is not None
        ):
            try:
                raw_actual = actual_match.candidate["actualSysConsignCount"]
                if raw_actual is None or str(raw_actual).strip() == "":
                    raise KeyError("actualSysConsignCount")
                if isinstance(raw_actual, bool):
                    raise ValueError("actual count cannot be boolean")
                parsed_actual = Decimal(str(raw_actual).strip())
                if (
                    not parsed_actual.is_finite()
                    or parsed_actual < 0
                    or parsed_actual != parsed_actual.to_integral_value()
                ):
                    raise ValueError(
                        "actual count must be a finite nonnegative integer"
                    )
                new_actual = str(int(parsed_actual))
            except KeyError:
                actual_status = "missing_value"
            except (InvalidOperation, TypeError, ValueError):
                actual_status = "invalid_actual"
            else:
                actual_auto_write = True
                row = _set_cell_number(row, layout.actual_col, row_number, new_actual)
                if new_actual != old_actual:
                    changed.append("actual")
        if not actual_auto_write:
            new_actual = old_actual
            review.append(_review_record(
                "actual", row_number, sheet_sku, sheet_name, actual_match,
                status=actual_status,
            ))

        return_percentage_style = (
            dominant_return_percentage_style
            or _cell_style(row, layout.return_col)
        )

        def update_return_rate(
            result: Any,
            value_field: str,
            col: str,
            field_name: str,
            old_value: str,
            style: str | None = None,
        ) -> tuple[bytes, str, str, bool]:
            current_row = row
            status = result.status
            match_ok = bool(
                result.auto_write
                and result.status in {"exact", "alias"}
                and result.candidate is not None
            )
            new_value = ""
            if match_ok:
                raw_value = result.candidate.get(value_field)
                if raw_value is None or str(raw_value).strip() == "":
                    try:
                        sales_count = Decimal(str(result.candidate["itemCount"]))
                    except (KeyError, InvalidOperation, TypeError, ValueError):
                        status = "invalid_value"
                    else:
                        status = (
                            "below_threshold"
                            if sales_count.is_finite() and sales_count < 50
                            else "invalid_value"
                        )
                else:
                    try:
                        new_value = _number_text(_rate_number(raw_value))
                    except (TypeError, ValueError):
                        status = "invalid_value"
                    else:
                        current_row = _set_cell_number(
                            current_row,
                            col,
                            row_number,
                            new_value,
                            style=style,
                        )
            if not new_value:
                current_row = _clear_cell_value(
                    current_row, col, row_number, style=style
                )
            if new_value != old_value:
                changed.append(field_name)
            if not match_ok or status == "invalid_value":
                review.append(_review_record(
                    field_name,
                    row_number,
                    sheet_sku,
                    sheet_name,
                    result,
                    status=status,
                ))
            return current_row, new_value, status, match_ok

        row, new_before_return, before_return_status, before_return_match_ok = (
            update_return_rate(
                before_return_match,
                "calculatedBeforeShipmentReturnRate",
                layout.before_return_col,
                "before_return",
                old_before_return,
                return_percentage_style,
            )
        )
        if after_return_rows is None:
            row = _clear_cell_value(
                row,
                layout.after_return_col,
                row_number,
                style=return_percentage_style,
            )
            if old_after_return:
                changed.append("after_return")
        else:
            (
                row,
                new_after_return,
                after_return_status,
                after_return_match_ok,
            ) = update_return_rate(
                after_return_match,
                "calculatedAfterShipmentReturnRate",
                layout.after_return_col,
                "after_return",
                old_after_return,
                return_percentage_style,
            )
        row, new_return, return_status, return_match_ok = update_return_rate(
            overall_return_match,
            "calculatedOverallReturnRate",
            layout.return_col,
            "return",
            old_return,
            return_percentage_style,
        )

        row = _set_cell_formula(
            row,
            layout.peer_col,
            row_number,
            peer_formula(layout.previous_col, row_number, cycle),
            default_style=peer_formula_style,
        )
        row = _set_cell_formula(
            row,
            layout.change_col,
            row_number,
            change_formula(layout.actual_col, layout.peer_col, row_number, cycle),
            default_style=change_formula_style,
        )
        formula_count += 2
        formula_rows.add(row_number)

        report_row = {
            "row": row_number,
            "sheet_sku": sheet_sku,
            "sheet_name": sheet_name,
            "actual_status": actual_status,
            "actual_auto_write": actual_auto_write,
            "actual_candidate_sku": _candidate_value(actual_match, "itemOuterId"),
            "actual_candidate_title": _candidate_value(actual_match, "title"),
            "before_return_status": before_return_status,
            "before_return_match_ok": before_return_match_ok,
            "before_return_candidate_sku": _candidate_value(
                before_return_match, "itemOuterId"
            ),
            "before_return_candidate_title": _candidate_value(
                before_return_match, "title"
            ),
            "after_return_status": after_return_status,
            "after_return_match_ok": after_return_match_ok,
            "after_return_candidate_sku": _candidate_value(
                after_return_match, "itemOuterId"
            ),
            "after_return_candidate_title": _candidate_value(
                after_return_match, "title"
            ),
            "return_status": return_status,
            "return_match_ok": return_match_ok,
            "return_candidate_sku": _candidate_value(
                overall_return_match, "itemOuterId"
            ),
            "return_candidate_title": _candidate_value(
                overall_return_match, "title"
            ),
            "old_actual": old_actual,
            "new_actual": new_actual,
            "old_before_return": old_before_return,
            "new_before_return": new_before_return,
            "old_after_return": old_after_return,
            "new_after_return": new_after_return,
            "old_return": old_return,
            "new_return": new_return,
            "changed_fields": tuple(changed),
        }
        rows_report.append(report_row)
        row_matches.setdefault(_normalize_sku(sheet_sku), []).append(report_row)
        return row

    patched = ROW_RE.sub(update_row, insertion.sheet_xml)
    patched = _extend_change_conditional_formatting(
        patched, layout.change_col, formula_rows
    )
    critical_results: list[dict[str, Any]] = []
    for sku in sorted({_normalize_sku(raw_sku) for raw_sku in critical_skus}):
        matches = row_matches.get(sku, [])
        selected = matches[0] if matches else None
        critical_results.append({
            "sku": sku,
            "sheet_exists": bool(matches),
            "actual_status": selected["actual_status"] if selected else "missing",
            "return_status": selected["return_status"] if selected else "missing",
            "before_return_status": (
                selected["before_return_status"] if selected else "missing"
            ),
            "after_return_status": (
                selected["after_return_status"] if selected else "missing"
            ),
            "actual_candidate_sku": selected["actual_candidate_sku"] if selected else "",
            "actual_candidate_title": selected["actual_candidate_title"] if selected else "",
            "before_return_candidate_sku": (
                selected["before_return_candidate_sku"] if selected else ""
            ),
            "before_return_candidate_title": (
                selected["before_return_candidate_title"] if selected else ""
            ),
            "after_return_candidate_sku": (
                selected["after_return_candidate_sku"] if selected else ""
            ),
            "after_return_candidate_title": (
                selected["after_return_candidate_title"] if selected else ""
            ),
            "return_candidate_sku": selected["return_candidate_sku"] if selected else "",
            "return_candidate_title": selected["return_candidate_title"] if selected else "",
            "passed": bool(
                selected
                and selected["actual_auto_write"]
                and selected["before_return_match_ok"]
                and (
                    after_return_rows is None
                    or selected["after_return_match_ok"]
                )
                and selected["return_match_ok"]
            ),
        })
    critical_failures = tuple(
        item for item in critical_results if not item["passed"]
    )
    return MonthlyWriteResult(
        sheet_xml=patched,
        layout=layout,
        inserted=insertion.inserted,
        rows=tuple(rows_report),
        review=tuple(review),
        critical_results=tuple(critical_results),
        critical_failures=critical_failures,
        formula_count=formula_count,
    )


def _cell_value(row: bytes, col: str, shared_strings: list[str]) -> str:
    cell = next(
        (candidate for candidate in CELL_RE.finditer(row) if _cell_col(candidate) == col),
        None,
    )
    return _cell_text(cell, shared_strings) if cell is not None else ""


def _cell_formula(row: bytes, col: str) -> str | None:
    cell = next(
        (candidate for candidate in CELL_RE.finditer(row) if _cell_col(candidate) == col),
        None,
    )
    if cell is None:
        return None
    formula = re.search(
        rb'<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*>(.*?)</(?:[A-Za-z_][\w.-]*:)?f>',
        cell.group("body") or b"",
        re.S,
    )
    if formula is None:
        return None
    return html.unescape(formula.group(1).decode("utf-8", "ignore")).strip()


def validate_monthly_sheet(
    sheet_xml: bytes,
    shared_strings: list[str],
    cycle: SyncCycle,
) -> dict[str, Any]:
    layout = discover_layout(sheet_xml, shared_strings, cycle)
    rows = list(ROW_RE.finditer(sheet_xml))
    header_match = next(
        (match for match in rows if int(match.group("row")) == layout.header_row),
        None,
    )
    if header_match is None:
        raise RuntimeError(f"header row {layout.header_row} is missing")
    header_row = header_match.group(0)
    found_actual = normalize_header(
        _cell_value(header_row, layout.actual_col, shared_strings)
    )
    expected_actual = normalize_header(actual_header(cycle))
    if found_actual != expected_actual:
        raise RuntimeError(
            f"actual header mismatch: found {found_actual!r}, expected {expected_actual!r}"
        )
    found_before_return = normalize_header(
        _cell_value(header_row, layout.before_return_col, shared_strings)
    )
    expected_before_return = normalize_header(
        return_header(cycle, "发货前退货率")
    )
    if found_before_return != expected_before_return:
        raise RuntimeError(
            "before return header mismatch: "
            f"found {found_before_return!r}, expected {expected_before_return!r}"
        )
    found_after_return = normalize_header(
        _cell_value(header_row, layout.after_return_col, shared_strings)
    )
    expected_after_return = normalize_header(
        return_header(cycle, "发货后退货率")
    )
    if found_after_return != expected_after_return:
        raise RuntimeError(
            "after return header mismatch: "
            f"found {found_after_return!r}, expected {expected_after_return!r}"
        )
    found_return = normalize_header(
        _cell_value(header_row, layout.return_col, shared_strings)
    )
    expected_return = normalize_header(return_header(cycle))
    if found_return != expected_return:
        raise RuntimeError(
            f"return header mismatch: found {found_return!r}, expected {expected_return!r}"
        )

    sku_rows: list[int] = []
    peer_count = 0
    change_count = 0
    for match in rows:
        row_number = int(match.group("row"))
        if row_number == layout.header_row:
            continue
        row = match.group(0)
        if not _cell_value(row, "E", shared_strings):
            continue
        sku_rows.append(row_number)
        found_peer = _cell_formula(row, layout.peer_col)
        expected_peer = peer_formula(layout.previous_col, row_number, cycle)
        if found_peer != expected_peer:
            raise RuntimeError(
                f"peer formula mismatch at row {row_number}: "
                f"found {found_peer!r}, expected {expected_peer!r}"
            )
        peer_count += 1
        found_change = _cell_formula(row, layout.change_col)
        expected_change = change_formula(
            layout.actual_col, layout.peer_col, row_number, cycle
        )
        if found_change != expected_change:
            raise RuntimeError(
                f"change formula mismatch at row {row_number}: "
                f"found {found_change!r}, expected {expected_change!r}"
            )
        change_count += 1

    return {
        "row_count": len(rows),
        "sku_rows": len(sku_rows),
        "formula_sample_rows": sku_rows[:5],
        "peer_formula_count": peer_count,
        "change_formula_count": change_count,
        "actual_header": found_actual,
        "before_return_header": found_before_return,
        "after_return_header": found_after_return,
        "return_header": found_return,
        "layout": {
            "header_row": layout.header_row,
            "previous_col": layout.previous_col,
            "peer_col": layout.peer_col,
            "actual_col": layout.actual_col,
            "change_col": layout.change_col,
            "before_return_col": layout.before_return_col,
            "after_return_col": layout.after_return_col,
            "return_col": layout.return_col,
            "month_cols": list(layout.month_cols),
            "needs_insert": layout.needs_insert,
            "insert_before_col": layout.insert_before_col,
        },
    }


def _candidate_value(result: Any, field: str) -> Any:
    return result.candidate.get(field, "") if result.candidate else ""


def _review_record(
    field: str,
    row_number: int,
    sheet_sku: str,
    sheet_name: str,
    result: Any,
    *,
    status: str | None = None,
) -> dict[str, Any]:
    return {
        "field": field,
        "row": row_number,
        "sheet_sku": sheet_sku,
        "sheet_name": sheet_name,
        "status": status or result.status,
        "confidence": round(result.confidence, 4),
        "candidate_sku": _candidate_value(result, "itemOuterId"),
        "candidate_title": _candidate_value(result, "title"),
        "different_columns": tuple(result.differences),
    }


def _tag_attributes(tag: bytes) -> dict[str, str]:
    return {
        key.decode("ascii"): html.unescape(value.decode("utf-8", "ignore"))
        for key, value in re.findall(rb'([A-Za-z_:][\w:.-]*)="([^"]*)"', tag)
    }


def _column_tag(start: int, end: int, attrs: dict[str, str]) -> bytes:
    pieces = [f'min="{start}"', f'max="{end}"']
    for key, value in attrs.items():
        if key in {"min", "max"}:
            continue
        pieces.append(f'{key}="{html.escape(value, quote=True)}"')
    return ("<col " + " ".join(pieces) + "/>").encode("utf-8")


def _rewrite_column_definitions(
    sheet_xml: bytes,
    layout: WorkbookLayout,
    insert_number: int,
    old_actual_number: int,
) -> bytes:
    cols_match = re.search(rb'<cols\b[^>]*>(?P<body>.*?)</cols>', sheet_xml, re.S)
    effective: list[dict[str, str] | None] = [None] * (MAX_EXCEL_COLUMN + 1)
    max_defined = 0
    if cols_match:
        for tag in re.findall(rb'<col\b[^>]*/>', cols_match.group("body")):
            attrs = _tag_attributes(tag)
            try:
                start, end = int(attrs["min"]), int(attrs["max"])
            except (KeyError, ValueError):
                continue
            if not 1 <= start <= end <= MAX_EXCEL_COLUMN:
                raise ValueError(f"invalid worksheet column span: {start}:{end}")
            clean = {key: value for key, value in attrs.items() if key not in {"min", "max"}}
            for number in range(start, end + 1):
                effective[number] = clean.copy()
            max_defined = max(max_defined, end)

    actual_attrs = (effective[old_actual_number] or {}).copy()
    actual_attrs.pop("hidden", None)
    shifted: list[dict[str, str] | None] = [None] * (MAX_EXCEL_COLUMN + 1)
    upper = min(MAX_EXCEL_COLUMN - 1, max(max_defined, old_actual_number))
    for old_number in range(1, upper + 1):
        new_number = old_number + 1 if old_number >= insert_number else old_number
        shifted[new_number] = None if effective[old_number] is None else effective[old_number].copy()
    shifted[insert_number] = actual_attrs

    shifted_months = [
        column_number(col) + (1 if column_number(col) >= insert_number else 0)
        for col in layout.month_cols
    ]
    shifted_actual = old_actual_number + 1
    for number in shifted_months:
        if number == shifted_actual:
            continue
        attrs = (shifted[number] or {}).copy()
        attrs["hidden"] = "1"
        shifted[number] = attrs
    visible = {
        insert_number,
        column_number(layout.peer_col) + 1,
        shifted_actual,
        column_number(layout.change_col) + 1,
        column_number(layout.before_return_col) + 1,
        column_number(layout.after_return_col) + 1,
        column_number(layout.return_col) + 1,
    }
    for number in visible:
        if shifted[number] is not None:
            shifted[number] = shifted[number].copy()
            shifted[number].pop("hidden", None)

    tags: list[bytes] = []
    number = 1
    limit = max(upper + 1, max(visible), max(shifted_months))
    while number <= limit:
        attrs = shifted[number]
        if attrs is None:
            number += 1
            continue
        end = number
        while end + 1 <= limit and shifted[end + 1] == attrs:
            end += 1
        tags.append(_column_tag(number, end, attrs))
        number = end + 1
    replacement = b"<cols>" + b"".join(tags) + b"</cols>"
    if cols_match:
        return sheet_xml[: cols_match.start()] + replacement + sheet_xml[cols_match.end() :]
    sheet_data = sheet_xml.find(b"<sheetData")
    if sheet_data < 0:
        raise ValueError("worksheet has no sheetData element")
    return sheet_xml[:sheet_data] + replacement + sheet_xml[sheet_data:]


def shift_drawing_anchors(drawing_xml: bytes, insert_before_col: str | int) -> bytes:
    insert_number = (
        column_number(insert_before_col)
        if isinstance(insert_before_col, str)
        else insert_before_col
    )
    if not isinstance(insert_number, int) or not 1 <= insert_number <= MAX_EXCEL_COLUMN:
        raise ValueError(f"invalid insertion column: {insert_before_col!r}")
    zero_based_insertion = insert_number - 1

    def shift(match: re.Match[bytes]) -> bytes:
        value = int(match.group(2))
        if value >= zero_based_insertion:
            value += 1
        return match.group(1) + str(value).encode("ascii") + match.group(3)

    return re.sub(rb'(<(?:[A-Za-z_][\w.-]*:)?col>)(\d+)(</(?:[A-Za-z_][\w.-]*:)?col>)', shift, drawing_xml)


def shift_qualified_worksheet_formulas(
    sheet_xml: bytes,
    insert_before_col: str | int,
    target_sheet: str = "分级总表",
) -> bytes:
    def shift_formula(match: re.Match[bytes]) -> bytes:
        formula = html.unescape(match.group(2).decode("utf-8", "ignore"))
        shifted = shift_formula_references(
            formula,
            insert_before_col,
            target_sheet,
            shift_unqualified=False,
        )
        return (
            match.group(1)
            + html.escape(shifted, quote=False).encode("utf-8")
            + match.group(3)
        )

    formula_tag = rb'(?:[A-Za-z_][\w.-]*:)?(?:formula[12]?|f)'
    return re.sub(
        rb'(<'+ formula_tag + rb'\b(?![^>]*?/\s*>)[^>]*>)(.*?)(</' + formula_tag + rb'>)',
        shift_formula,
        sheet_xml,
        flags=re.S,
    )


def update_workbook_xml(
    workbook_xml: bytes,
    insert_before_col: str | int | None = None,
    sheet_name: str = "分级总表",
) -> bytes:
    updated = workbook_xml
    root_match = re.search(
        rb'<(?P<prefix>[A-Za-z_][\w.-]*:)?workbook\b',
        updated,
    )
    if root_match is None:
        raise ValueError("workbook XML has no workbook root element")
    workbook_prefix = root_match.group("prefix") or b""
    if insert_before_col is not None:
        def shift_name(match: re.Match[bytes]) -> bytes:
            value = html.unescape(match.group(2).decode("utf-8", "ignore"))
            shifted = shift_formula_references(
                value,
                insert_before_col,
                sheet_name,
                shift_unqualified=False,
            )
            return match.group(1) + html.escape(shifted, quote=False).encode("utf-8") + match.group(3)

        updated = re.sub(
            rb'(<(?:[A-Za-z_][\w.-]*:)?definedName\b[^>]*>)(.*?)'
            rb'(</(?:[A-Za-z_][\w.-]*:)?definedName>)',
            shift_name,
            updated,
            flags=re.S,
        )

    calc_match = re.search(
        rb'<(?P<prefix>[A-Za-z_][\w.-]*:)?calcPr\b[^>]*'
        rb'(?:/>|>.*?</(?:[A-Za-z_][\w.-]*:)?calcPr>)',
        updated,
        re.S,
    )
    required = {
        "calcMode": "auto",
        "fullCalcOnLoad": "1",
        "forceFullCalc": "1",
    }
    if calc_match:
        attrs = _tag_attributes(calc_match.group(0))
        attrs.update(required)
        pieces = [
            f'{key}="{html.escape(value, quote=True)}"'
            for key, value in attrs.items()
        ]
        replacement = (
            b"<"
            + workbook_prefix
            + b"calcPr "
            + " ".join(pieces).encode("utf-8")
            + b"/>"
        )
        updated = updated[: calc_match.start()] + replacement + updated[calc_match.end() :]
    else:
        replacement = (
            b"<"
            + workbook_prefix
            + b'calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>'
        )
        later_child = re.search(
            rb'<(?:[A-Za-z_][\w.-]*:)?(?:oleSize|customWorkbookViews|pivotCaches|smartTagPr|'
            rb'smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b',
            updated,
        )
        closing_tag = b"</" + workbook_prefix + b"workbook>"
        insertion = later_child.start() if later_child else updated.rfind(closing_tag)
        if insertion < 0:
            raise ValueError("workbook XML has no closing workbook element")
        updated = updated[:insertion] + replacement + updated[insertion:]
    return updated
