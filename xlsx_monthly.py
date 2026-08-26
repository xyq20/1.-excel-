from __future__ import annotations

from dataclasses import dataclass
import html
import re
from typing import Callable

from monthly_schedule import SyncCycle, actual_header


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
    return_col: str
    month_cols: tuple[str, ...]
    needs_insert: bool
    insert_before_col: str | None


@dataclass(frozen=True)
class MonthInsertionResult:
    sheet_xml: bytes
    layout: WorkbookLayout
    inserted: bool


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
            or text.startswith("退货率")
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
    return_col, return_row = unique_header("退货率", lambda text: text.startswith("退货率"))
    if len({peer_row, change_row, return_row}) != 1:
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
        column_number(return_col),
    )
    if ordered_columns != tuple(sorted(ordered_columns)) or len(set(ordered_columns)) != 5:
        raise ValueError(
            "invalid header column order; expected previous < peer < actual < change < return"
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

    return WorkbookLayout(
        header_row=header_row,
        previous_col=previous_col,
        peer_col=peer_col,
        actual_col=actual_col,
        change_col=change_col,
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


def shift_formula_references(formula: str, insert_before_col: str | int) -> str:
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
        if quote not in {'"', "'"}:
            index += 1
            continue
        output.append(_shift_reference_segment(formula[start:index], insert_number))
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
    output.append(_shift_reference_segment(formula[start:], insert_number))
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


def update_workbook_xml(
    workbook_xml: bytes,
    insert_before_col: str | int | None = None,
    sheet_name: str = "分级总表",
) -> bytes:
    updated = workbook_xml
    if insert_before_col is not None:
        def shift_name(match: re.Match[bytes]) -> bytes:
            value = html.unescape(match.group(2).decode("utf-8", "ignore"))

            def shift_qualified(reference: re.Match[str]) -> str:
                encoded_sheet = reference.group("sheet")
                decoded_sheet = (
                    encoded_sheet[1:-1].replace("''", "'")
                    if encoded_sheet.startswith("'")
                    else encoded_sheet
                )
                if decoded_sheet != sheet_name:
                    return reference.group(0)
                shifted_reference = shift_formula_references(
                    reference.group("reference"),
                    insert_before_col,
                )
                return encoded_sheet + "!" + shifted_reference

            shifted = QUALIFIED_REFERENCE_RE.sub(shift_qualified, value)
            return match.group(1) + html.escape(shifted, quote=False).encode("utf-8") + match.group(3)

        updated = re.sub(
            rb'(<definedName\b[^>]*>)(.*?)(</definedName>)',
            shift_name,
            updated,
            flags=re.S,
        )

    calc_match = re.search(rb'<calcPr\b[^>]*(?:/>|>.*?</calcPr>)', updated, re.S)
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
        replacement = ("<calcPr " + " ".join(pieces) + "/>").encode("utf-8")
        updated = updated[: calc_match.start()] + replacement + updated[calc_match.end() :]
    else:
        replacement = b'<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>'
        later_child = re.search(
            rb'<(?:[A-Za-z_][\w.-]*:)?(?:oleSize|customWorkbookViews|pivotCaches|smartTagPr|'
            rb'smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b',
            updated,
        )
        insertion = later_child.start() if later_child else updated.rfind(b"</workbook>")
        if insertion < 0:
            raise ValueError("workbook XML has no closing workbook element")
        updated = updated[:insertion] + replacement + updated[insertion:]
    return updated
