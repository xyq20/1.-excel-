import re
import unittest
from datetime import date
from xml.etree import ElementTree as ET

from monthly_schedule import resolve_sync_cycle
from xlsx_monthly import (
    apply_monthly_values,
    WorkbookLayout,
    column_label,
    column_number,
    discover_layout,
    insert_month_column,
    normalize_header,
    shift_drawing_anchors,
    shift_formula_references,
    shift_qualified_worksheet_formulas,
    update_workbook_xml,
)
from erp_excel_sync import MatchResult, RETURN_RATE_FIELD, choose_candidate


def workbook_sheet(
    headers: list[tuple[str, str]],
    *,
    rows: bytes = b"",
    cols: bytes = b'<cols><col min="1" max="50" width="12" style="4"/></cols>',
    extras: bytes = b'<autoFilter ref="B1:AY10"/>',
) -> bytes:
    header_cells = b"".join(
        (
            f'<c r="{col}1" s="9" t="inlineStr"><is><t>{text}</t></is></c>'
        ).encode("utf-8")
        for col, text in headers
    )
    return (
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<dimension ref="A1:AY10"/>'
        + cols
        + b'<sheetData><row r="1">'
        + header_cells
        + b'</row>'
        + rows
        + b'</sheetData>'
        + extras
        + b'</worksheet>'
    )


def august_sheet(*, extras: bytes = b'<autoFilter ref="B1:AY10"/>') -> bytes:
    rows = (
        b'<row r="2"><c r="W2" s="140"><v>10</v></c>'
        b'<c r="X2" s="141"><f>W2/31*14</f><v>4</v></c>'
        b'<c r="Y2" s="141"><v>12</v></c>'
        b'<c r="Z2" s="142"><f>TEXT(Y2-X2,&quot;Y2\xe5\x8e\x9f\xe6\x96\x87&quot;)+SUM($X$3:AA3)</f><v>0</v></c>'
        b'<c r="AA2" s="219"><v>0.2</v></c></row>'
    )
    return workbook_sheet(
        [
            ("W", "7月实发"),
            ("X", "同期销量"),
            ("Y", "8月实发（8.15）"),
            ("Z", "变化情况"),
            ("AA", "退货率（7.1-7.31）"),
        ],
        rows=rows,
        extras=extras,
    )


class ColumnAndHeaderTests(unittest.TestCase):
    def test_column_conversions_are_total_for_valid_excel_columns(self):
        self.assertEqual(column_number("A"), 1)
        self.assertEqual(column_number("AA"), 27)
        self.assertEqual(column_label(52), "AZ")
        self.assertEqual(column_label(16384), "XFD")

    def test_column_conversions_reject_invalid_inputs(self):
        for value in ("", "A1", "a", 1, None):
            with self.subTest(value=value), self.assertRaises(ValueError):
                column_number(value)
        for value in (0, -1, 1.5, "1", None):
            with self.subTest(value=value), self.assertRaises(ValueError):
                column_label(value)

    def test_normalizes_historical_month_headers(self):
        self.assertEqual(normalize_header(" 8 月 实发销量（8.15） "), "8月实发(8.15)")
        self.assertEqual(normalize_header("7月实销 (7.15)"), "7月实发(7.15)")


class LayoutDiscoveryTests(unittest.TestCase):
    def test_discovers_august_fifteenth_layout_from_shared_and_inline_headers(self):
        sheet = august_sheet().replace(
            b'<c r="W1" s="9" t="inlineStr"><is><t>7\xe6\x9c\x88\xe5\xae\x9e\xe5\x8f\x91</t></is></c>',
            b'<c r="W1" s="9" t="s"><v>0</v></c>',
        )

        layout = discover_layout(sheet, ["7月实发销量"], resolve_sync_cycle(date(2026, 8, 15)))

        self.assertEqual(
            layout,
            WorkbookLayout(
                header_row=1,
                previous_col="W",
                peer_col="X",
                actual_col="Y",
                change_col="Z",
                return_col="AA",
                month_cols=("W", "Y"),
                needs_insert=False,
                insert_before_col=None,
            ),
        )

    def test_new_month_fifteenth_needs_insertion_before_peer(self):
        layout = discover_layout(august_sheet(), [], resolve_sync_cycle(date(2026, 9, 15)))

        self.assertTrue(layout.needs_insert)
        self.assertEqual(layout.insert_before_col, "X")
        self.assertEqual(layout.actual_col, "Y")

    def test_first_node_accepts_staged_header_without_insertion(self):
        layout = discover_layout(august_sheet(), [], resolve_sync_cycle(date(2026, 9, 1)))

        self.assertFalse(layout.needs_insert)
        self.assertIsNone(layout.insert_before_col)
        self.assertEqual(layout.actual_col, "Y")

    def test_ambiguous_core_headers_raise_value_error(self):
        sheet = august_sheet().replace(
            b'</row>',
            b'<c r="AB1" t="inlineStr"><is><t>\xe5\x90\x8c\xe6\x9c\x9f\xe9\x94\x80\xe9\x87\x8f</t></is></c></row>',
            1,
        )

        with self.assertRaisesRegex(ValueError, "同期销量"):
            discover_layout(sheet, [], resolve_sync_cycle(date(2026, 9, 15)))

    def test_malformed_core_column_order_raises_value_error(self):
        sheet = workbook_sheet(
            [
                ("W", "7月实发"),
                ("X", "同期销量"),
                ("Y", "8月实发（8.15）"),
                ("Z", "退货率（7.1-7.31）"),
                ("AA", "变化情况"),
            ]
        )

        with self.assertRaisesRegex(ValueError, "column order"):
            discover_layout(sheet, [], resolve_sync_cycle(date(2026, 9, 15)))

    def test_partial_already_rolled_layout_rejects_stale_previous_month(self):
        sheet = workbook_sheet(
            [
                ("W", "7月实发"),
                ("X", "同期销量"),
                ("Y", "9月实发（9.15）"),
                ("Z", "变化情况"),
                ("AA", "退货率（8.1-8.31）"),
            ]
        )

        with self.assertRaisesRegex(ValueError, "previous-month"):
            discover_layout(sheet, [], resolve_sync_cycle(date(2026, 9, 15)))


class ReferenceShiftTests(unittest.TestCase):
    def test_formula_translation_leaves_quoted_cell_like_text_unchanged(self):
        formula = 'TEXT(Y2-X2,"Y2原文") + SUM($X$3:AA3)'

        self.assertEqual(
            shift_formula_references(formula, "X"),
            'TEXT(Z2-Y2,"Y2原文") + SUM($Y$3:AB3)',
        )

    def test_formula_translation_does_not_treat_numbered_function_names_as_cells(self):
        self.assertEqual(shift_formula_references("LOG10(X2)", "X"), "LOG10(Y2)")

    def test_formula_translation_leaves_non_target_sheet_references_unchanged(self):
        self.assertEqual(
            shift_formula_references("X2+OtherSheet!X2+'Other Sheet'!$X$2", "X"),
            "Y2+OtherSheet!X2+'Other Sheet'!$X$2",
        )

    def test_formula_translation_shifts_target_sheet_qualified_references(self):
        self.assertEqual(
            shift_formula_references("分级总表!X2+'分级总表'!$X$2", "X"),
            "分级总表!Y2+'分级总表'!$Y$2",
        )

    def test_formula_translation_shifts_whole_column_ranges(self):
        self.assertEqual(
            shift_formula_references('FILTER($AS:$AS,$AR:$AR=F4)+SUM("$X:$X")', "X"),
            'FILTER($AT:$AT,$AS:$AS=F4)+SUM("$X:$X")',
        )

    def test_cross_sheet_helper_only_shifts_qualified_target_references(self):
        sheet = (
            b'<worksheet><sheetData><row r="1"><c r="A1"><f>'
            b'X2+\xe5\x88\x86\xe7\xba\xa7\xe6\x80\xbb\xe8\xa1\xa8!X2+OtherSheet!X2+'
            b'&apos;\xe5\x88\x86\xe7\xba\xa7\xe6\x80\xbb\xe8\xa1\xa8&apos;!$X$2'
            b'</f></c></row></sheetData></worksheet>'
        )

        shifted = shift_qualified_worksheet_formulas(sheet, "X")

        self.assertIn(
            b'X2+\xe5\x88\x86\xe7\xba\xa7\xe6\x80\xbb\xe8\xa1\xa8!Y2+OtherSheet!X2+'
            b"'\xe5\x88\x86\xe7\xba\xa7\xe6\x80\xbb\xe8\xa1\xa8'!$Y$2",
            shifted,
        )


class MonthInsertionTests(unittest.TestCase):
    def test_self_closing_shared_formula_does_not_consume_following_xml(self):
        source = august_sheet().replace(
            b'<f>W2/31*14</f>',
            b'<f t="shared" si="0"/>',
        )

        result = insert_month_column(source, [], resolve_sync_cycle(date(2026, 9, 15)))

        ET.fromstring(result.sheet_xml)
        self.assertIn(b'<f t="shared" si="0"/>', result.sheet_xml)
        self.assertIn(b'<c r="Z2" s="141"><v>12</v></c>', result.sheet_xml)

    def test_copied_previous_value_does_not_retain_an_actual_column_formula(self):
        source = august_sheet().replace(
            b'<c r="Y2" s="141"><v>12</v></c>',
            b'<c r="Y2" s="141"><f t="shared" si="2"/><v>12</v></c>',
        )

        result = insert_month_column(source, [], resolve_sync_cycle(date(2026, 9, 15)))

        self.assertIn(b'<c r="X2" s="141"><v>12</v></c>', result.sheet_xml)
        self.assertIn(b'<c r="Z2" s="141"><f t="shared" si="2"/><v>12</v></c>', result.sheet_xml)

    def test_rediscovery_keeps_using_shared_strings_for_core_headers(self):
        source = august_sheet()
        shared = ["7月实发", "同期销量", "8月实发（8.15）", "变化情况", "退货率（7.1-7.31）"]
        for index, (col, header) in enumerate(
            (("W", shared[0]), ("X", shared[1]), ("Y", shared[2]), ("Z", shared[3]), ("AA", shared[4]))
        ):
            inline = (
                f'<c r="{col}1" s="9" t="inlineStr"><is><t>{header}</t></is></c>'
            ).encode("utf-8")
            source = source.replace(
                inline,
                f'<c r="{col}1" s="9" t="s"><v>{index}</v></c>'.encode("ascii"),
            )

        result = insert_month_column(source, shared, resolve_sync_cycle(date(2026, 9, 15)))

        self.assertTrue(result.inserted)
        self.assertEqual(result.layout.peer_col, "Y")

    def test_september_insertion_rolls_columns_cells_headers_and_formulas(self):
        extras = (
            b'<autoFilter ref="B1:AY10"/>'
            b'<mergeCells><mergeCell ref="W3:X3"/></mergeCells>'
            b'<hyperlinks><hyperlink ref="X2"/></hyperlinks>'
            b'<conditionalFormatting sqref="X2:AA2"><cfRule><formula>X2&gt;0</formula></cfRule></conditionalFormatting>'
            b'<dataValidations><dataValidation sqref="X2 X4:AA4"><formula1>SUM($X$2:AA2)</formula1></dataValidation></dataValidations>'
            b'<pane topLeftCell="X2" activeCell="Y2"/>'
            b'<selection activeCell="X2" sqref="X2:AA2"/>'
        )
        cols = (
            b'<cols><col min="1" max="24" width="12" style="4"/>'
            b'<col min="25" max="25" width="16" style="141" customWidth="1"/>'
            b'<col min="26" max="50" width="12" style="4"/></cols>'
        )
        source = august_sheet(extras=extras).replace(
            b'<cols><col min="1" max="50" width="12" style="4"/></cols>',
            cols,
        )
        source = source.replace(b'<row r="2">', b'<row r="2" spans="23:27">')

        result = insert_month_column(source, [], resolve_sync_cycle(date(2026, 9, 15)))

        self.assertTrue(result.inserted)
        self.assertEqual(
            (
                result.layout.previous_col,
                result.layout.peer_col,
                result.layout.actual_col,
                result.layout.change_col,
                result.layout.return_col,
            ),
            ("X", "Y", "Z", "AA", "AB"),
        )
        self.assertEqual(result.layout.month_cols, ("W", "X", "Z"))
        text = result.sheet_xml.decode("utf-8")
        self.assertRegex(text, r'<c r="X1"[^>]*t="inlineStr"[^>]*><is><t>8月实发</t></is></c>')
        self.assertRegex(text, r'<c r="Z1"[^>]*t="inlineStr"[^>]*><is><t>9月实发（9\.15）</t></is></c>')
        self.assertRegex(text, r'<c r="X2"[^>]*s="141"[^>]*><v>12</v></c>')
        self.assertRegex(text, r'<c r="Z2"[^>]*s="141"[^>]*><v>12</v></c>')
        self.assertIn('TEXT(Z2-Y2,"Y2原文")+SUM($Y$3:AB3)', text)
        self.assertIn('autoFilter ref="B1:AZ10"', text)
        self.assertIn('dimension ref="A1:AZ10"', text)
        self.assertIn('mergeCell ref="W3:Y3"', text)
        self.assertIn('hyperlink ref="Y2"', text)
        self.assertIn('conditionalFormatting sqref="Y2:AB2"', text)
        self.assertIn('<formula>Y2&gt;0</formula>', text)
        self.assertIn('dataValidation sqref="Y2 Y4:AB4"', text)
        self.assertIn('<formula1>SUM($Y$2:AB2)</formula1>', text)
        self.assertIn('pane topLeftCell="Y2" activeCell="Z2"', text)
        self.assertIn('selection activeCell="Y2" sqref="Y2:AB2"', text)
        self.assertIn('<row r="2" spans="23:28">', text)

    def test_insertion_splits_col_spans_copies_actual_style_and_hides_only_older_months(self):
        cols = (
            b'<cols><col min="1" max="24" width="12" style="4"/>'
            b'<col min="25" max="25" width="16" style="141" customWidth="1"/>'
            b'<col min="26" max="50" width="12" style="4"/></cols>'
        )
        source = august_sheet().replace(
            b'<cols><col min="1" max="50" width="12" style="4"/></cols>',
            cols,
        )

        result = insert_month_column(source, [], resolve_sync_cycle(date(2026, 9, 15)))

        cols_xml = re.search(rb'<cols>.*?</cols>', result.sheet_xml, re.S).group(0)
        definitions = []
        for tag in re.findall(rb'<col\b[^>]*/>', cols_xml):
            attrs = dict(re.findall(rb'(\w+)="([^"]*)"', tag))
            definitions.append({key.decode(): value.decode() for key, value in attrs.items()})
        self.assertTrue(any(item["min"] == "24" and item["max"] == "24" and item.get("style") == "141" and item.get("width") == "16" for item in definitions))
        self.assertTrue(any(item["min"] == "23" and item["max"] == "23" and item.get("hidden") == "1" for item in definitions))
        for visible in (24, 25, 26, 27, 28):
            covering = [item for item in definitions if int(item["min"]) <= visible <= int(item["max"])]
            self.assertEqual(len(covering), 1, (visible, definitions))
            self.assertNotEqual(covering[0].get("hidden"), "1")
        for left, right in zip(definitions, definitions[1:]):
            self.assertLess(int(left["max"]), int(right["min"]), definitions)

    def test_repeating_same_fifteenth_cycle_is_byte_identical(self):
        first = insert_month_column(august_sheet(), [], resolve_sync_cycle(date(2026, 9, 15)))

        second = insert_month_column(first.sheet_xml, [], resolve_sync_cycle(date(2026, 9, 30)))

        self.assertFalse(second.inserted)
        self.assertEqual(second.sheet_xml, first.sheet_xml)
        self.assertEqual(second.layout, first.layout)

    def test_first_node_does_not_rewrite_or_insert_staged_workbook(self):
        source = august_sheet()

        result = insert_month_column(source, [], resolve_sync_cycle(date(2026, 9, 1)))

        self.assertFalse(result.inserted)
        self.assertEqual(result.sheet_xml, source)

    def test_span_starting_at_insertion_column_keeps_inserted_column_in_span(self):
        source = august_sheet().replace(
            b'<row r="2">',
            b'<row r="2" spans="24:27">',
        )

        result = insert_month_column(source, [], resolve_sync_cycle(date(2026, 9, 15)))

        self.assertIn(b'<row r="2" spans="24:28">', result.sheet_xml)


class MonthlyValueWriteTests(unittest.TestCase):
    def _sheet(self, *, header_actual="8月实发（8.15）"):
        rows = (
            b'<row r="2"><c r="E2" s="5" t="s"><v>0</v></c><c r="F2" t="s"><v>1</v></c>'
            b'<c r="X2" s="141"><f>stale peer</f><v>999</v></c>'
            b'<c r="Y2" s="177"><v>10</v></c><c r="Z2" s="142"><f>stale change</f><v>8</v></c>'
            b'<c r="AA2" s="219"><v>0.25</v></c></row>'
            b'<row r="3"><c r="E3" t="s"><v>2</v></c><c r="F3" t="s"><v>3</v></c>'
            b'<c r="X3" s="241"><v>4</v></c><c r="Y3" s="178"><v>20</v></c>'
            b'<c r="Z3" s="242"><v>5</v></c><c r="AA3" s="220"><v>0.4</v></c></row>'
            b'<row r="4"><c r="E4" t="s"><v>4</v></c><c r="F4" t="s"><v>5</v></c>'
            b'<c r="X4" s="141"><v>7</v></c><c r="Y4" s="179"><v>30</v></c>'
            b'<c r="Z4" s="142"><v>6</v></c><c r="AA4" s="221"><v>0.6</v></c></row>'
        )
        return workbook_sheet(
            [
                ("W", "7月实发"),
                ("X", "同期销量"),
                ("Y", header_actual),
                ("Z", "变化情况"),
                ("AA", "退货率（7.1-7.31）"),
            ],
            rows=rows,
        )

    def _apply(self, sheet, cycle, actual_rows, return_rows, **kwargs):
        return apply_monthly_values(
            sheet,
            ["A-1", "Alpha", "B-2", "Beta", "Alias SKU", "Gamma"],
            cycle,
            actual_rows,
            return_rows,
            aliases=kwargs.get("aliases", {}),
            critical_skus=kwargs.get("critical_skus", set()),
            matcher=choose_candidate,
        )

    def test_uses_independent_datasets_and_never_leaks_unused_conflicting_fields(self):
        cycle = resolve_sync_cycle(date(2026, 8, 15))
        actual_rows = [{
            "itemOuterId": "A-1", "title": "Alpha", "actualSysConsignCount": 123,
            RETURN_RATE_FIELD: "99.99%",
        }]
        return_rows = [{
            "itemOuterId": "A-1", "title": "Alpha", "actualSysConsignCount": 987,
            RETURN_RATE_FIELD: "51.85%",
        }]

        result = self._apply(self._sheet(), cycle, actual_rows, return_rows)

        self.assertIn(b'<c r="Y2" s="177"><v>123</v></c>', result.sheet_xml)
        self.assertIn(b'<c r="AA2" s="219"><v>0.5185</v></c>', result.sheet_xml)
        self.assertEqual(result.rows[0]["actual_status"], "exact")
        self.assertEqual(result.rows[0]["return_status"], "exact")
        self.assertEqual(result.rows[0]["changed_fields"], ("actual", "return"))

    def test_updates_every_sheet_sku_row_even_when_not_critical(self):
        cycle = resolve_sync_cycle(date(2026, 8, 15))
        actual_rows = [
            {"itemOuterId": "A-1", "title": "Alpha", "actualSysConsignCount": 11},
            {"itemOuterId": "B-2", "title": "Beta", "actualSysConsignCount": 22},
        ]
        return_rows = [
            {"itemOuterId": "A-1", "title": "Alpha", RETURN_RATE_FIELD: "10%"},
            {"itemOuterId": "B-2", "title": "Beta", RETURN_RATE_FIELD: "20%"},
        ]

        result = self._apply(
            self._sheet(), cycle, actual_rows, return_rows, critical_skus={"A-1"}
        )

        self.assertIn(b'<c r="Y3" s="178"><v>22</v></c>', result.sheet_xml)
        self.assertIn(b'<c r="AA3" s="220"><v>0.2</v></c>', result.sheet_xml)
        self.assertEqual(len(result.rows), 3)
        self.assertEqual(len(result.critical_results), 1)

    def test_actual_can_update_while_unmatched_return_retains_old_value_and_reports_field(self):
        cycle = resolve_sync_cycle(date(2026, 8, 15))

        result = self._apply(
            self._sheet(), cycle,
            [{"itemOuterId": "A-1", "title": "Alpha", "actualSysConsignCount": 77}],
            [],
        )

        self.assertIn(b'<c r="Y2" s="177"><v>77</v></c>', result.sheet_xml)
        self.assertIn(b'<c r="AA2" s="219"><v>0.25</v></c>', result.sheet_xml)
        self.assertEqual(result.rows[0]["return_status"], "unmatched")
        self.assertEqual(result.rows[0]["new_return"], "0.25")
        self.assertTrue(any(item["field"] == "return" and item["row"] == 2 for item in result.review))

    def test_exact_candidates_missing_their_own_field_retain_old_values_and_fail_critical(self):
        cycle = resolve_sync_cycle(date(2026, 8, 15))

        result = self._apply(
            self._sheet(), cycle,
            [{"itemOuterId": "A-1", "title": "Alpha", RETURN_RATE_FIELD: "99%"}],
            [{"itemOuterId": "A-1", "title": "Alpha", "actualSysConsignCount": 999}],
            critical_skus={"A-1"},
        )

        self.assertIn(b'<c r="Y2" s="177"><v>10</v></c>', result.sheet_xml)
        self.assertIn(b'<c r="AA2" s="219"><v>0.25</v></c>', result.sheet_xml)
        self.assertEqual(result.rows[0]["actual_status"], "missing_value")
        self.assertEqual(result.rows[0]["return_status"], "missing_value")
        self.assertEqual(
            {(item["field"], item["status"]) for item in result.review if item["row"] == 2},
            {("actual", "missing_value"), ("return", "missing_value")},
        )
        self.assertFalse(result.critical_results[0]["passed"])

    def test_non_finite_target_values_are_reviewed_without_overwriting(self):
        result = self._apply(
            self._sheet(), resolve_sync_cycle(date(2026, 8, 15)),
            [{"itemOuterId": "A-1", "title": "Alpha", "actualSysConsignCount": "inf"}],
            [{"itemOuterId": "A-1", "title": "Alpha", RETURN_RATE_FIELD: "NaN"}],
        )

        self.assertIn(b'<c r="Y2" s="177"><v>10</v></c>', result.sheet_xml)
        self.assertIn(b'<c r="AA2" s="219"><v>0.25</v></c>', result.sheet_xml)
        self.assertEqual(result.rows[0]["actual_status"], "invalid_value")
        self.assertEqual(result.rows[0]["return_status"], "invalid_value")

    def test_alias_matches_both_windows_and_writes(self):
        cycle = resolve_sync_cycle(date(2026, 8, 15))

        result = self._apply(
            self._sheet(), cycle,
            [{"itemOuterId": "ERP-G", "title": "Gamma", "actualSysConsignCount": 44}],
            [{"itemOuterId": "ERP-G", "title": "Gamma", RETURN_RATE_FIELD: "12.5%"}],
            aliases={"alias sku": "erp-g"},
        )

        self.assertIn(b'<c r="Y4" s="179"><v>44</v></c>', result.sheet_xml)
        self.assertIn(b'<c r="AA4" s="221"><v>0.125</v></c>', result.sheet_xml)
        self.assertEqual(result.rows[2]["actual_status"], "alias")
        self.assertEqual(result.rows[2]["return_status"], "alias")

    def test_unsuffixed_erp_rate_is_percentage_points_even_below_one(self):
        result = self._apply(
            self._sheet(), resolve_sync_cycle(date(2026, 8, 15)), [],
            [{"itemOuterId": "A-1", "title": "Alpha", RETURN_RATE_FIELD: "0.5"}],
        )

        self.assertIn(b'<c r="AA2" s="219"><v>0.005</v></c>', result.sheet_xml)

    def test_critical_missing_or_partial_match_returns_failures_without_raising(self):
        cycle = resolve_sync_cycle(date(2026, 8, 15))

        result = self._apply(
            self._sheet(), cycle,
            [{"itemOuterId": "A-1", "title": "Alpha", "actualSysConsignCount": 1}],
            [],
            critical_skus={"A-1", "MISSING"},
        )

        self.assertEqual(len(result.critical_results), 2)
        self.assertEqual(len(result.critical_failures), 2)
        by_sku = {item["sku"]: item for item in result.critical_results}
        self.assertTrue(by_sku["a-1"]["sheet_exists"])
        self.assertEqual(by_sku["a-1"]["actual_status"], "exact")
        self.assertEqual(by_sku["a-1"]["return_status"], "unmatched")
        self.assertFalse(by_sku["a-1"]["passed"])
        self.assertFalse(by_sku["missing"]["sheet_exists"])

    def test_critical_requires_auto_write_even_when_matcher_uses_an_exact_status(self):
        def blocked_matcher(sheet_sku, sheet_name, rows, aliases):
            return MatchResult(
                "exact", rows[0] if rows else None, 1.0, False, ("candidate",)
            )

        result = apply_monthly_values(
            self._sheet(),
            ["A-1", "Alpha", "B-2", "Beta", "Alias SKU", "Gamma"],
            resolve_sync_cycle(date(2026, 8, 15)),
            [{"itemOuterId": "A-1", "title": "Alpha", "actualSysConsignCount": 1}],
            [{"itemOuterId": "A-1", "title": "Alpha", RETURN_RATE_FIELD: "1%"}],
            aliases={},
            critical_skus={"A-1"},
            matcher=blocked_matcher,
        )

        self.assertFalse(result.critical_results[0]["passed"])
        self.assertEqual(len(result.critical_failures), 1)

    def test_critical_reports_are_sorted_after_normalization(self):
        result = self._apply(
            self._sheet(), resolve_sync_cycle(date(2026, 8, 15)), [], [],
            critical_skus=[" Z ", "A"],
        )

        self.assertEqual(
            [item["sku"] for item in result.critical_results],
            ["a", "z"],
        )

    def test_replaces_formulas_for_every_sku_row_preserving_styles_without_caches(self):
        cycle = resolve_sync_cycle(date(2024, 3, 15))
        sheet = self._sheet(header_actual="3月实发（3.15）").replace(
            "7月实发".encode(), "2月实发".encode()
        )

        result = self._apply(sheet, cycle, [], [])

        self.assertEqual(result.formula_count, 6)
        self.assertIn(b'<c r="X2" s="141"><f>W2/29*14</f></c>', result.sheet_xml)
        self.assertIn(b'<c r="X3" s="241"><f>W3/29*14</f></c>', result.sheet_xml)
        self.assertIn(
            'c r="Z2" s="142"><f>TEXT(Y2-X2,"3月增加0件；3月减少0件；持平")</f></c>'.encode(),
            result.sheet_xml,
        )
        self.assertNotRegex(result.sheet_xml, rb'<c r="(?:X|Z)[234]"[^>]*><f>.*?</f><v>')

    def test_first_node_renames_staged_headers_without_insertion(self):
        cycle = resolve_sync_cycle(date(2026, 9, 1))

        result = self._apply(self._sheet(), cycle, [], [])

        self.assertFalse(result.inserted)
        self.assertIn("8月实发".encode(), result.sheet_xml)
        self.assertNotIn("8月实发（8.15）".encode(), result.sheet_xml)
        self.assertIn("退货率（7.15-8.15）".encode(), result.sheet_xml)
        self.assertEqual(result.layout.actual_col, "Y")

    def test_fifteenth_inserts_new_month_once_then_refreshes_idempotently(self):
        cycle = resolve_sync_cycle(date(2026, 9, 15))

        first = self._apply(self._sheet(), cycle, [], [])
        second = self._apply(first.sheet_xml, cycle, [], [])

        self.assertTrue(first.inserted)
        self.assertFalse(second.inserted)
        self.assertEqual(second.sheet_xml, first.sheet_xml)
        self.assertEqual(second.layout.actual_col, "Z")


class WorkbookAndDrawingTests(unittest.TestCase):
    def test_updates_defined_names_and_forces_full_automatic_recalculation(self):
        workbook = (
            b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            b'<definedNames><definedName name="Print_Area">'
            b'\xe5\x88\x86\xe7\xba\xa7\xe6\x80\xbb\xe8\xa1\xa8!$B$1:$AY$561'
            b'</definedName><definedName name="Other">OtherSheet!$X$1:$AY$2</definedName>'
            b'</definedNames>'
            b'<calcPr calcMode="manual" fullCalcOnLoad="0" forceFullCalc="0"/>'
            b'</workbook>'
        )

        updated = update_workbook_xml(workbook, "X")

        self.assertIn(b'!$B$1:$AZ$561', updated)
        self.assertIn(b'OtherSheet!$X$1:$AY$2', updated)
        calc = re.search(rb'<calcPr\b[^>]*/>', updated).group(0)
        self.assertIn(b'calcMode="auto"', calc)
        self.assertIn(b'fullCalcOnLoad="1"', calc)
        self.assertIn(b'forceFullCalc="1"', calc)

    def test_adds_calc_properties_when_the_workbook_has_none(self):
        workbook = b'<workbook><sheets/></workbook>'

        updated = update_workbook_xml(workbook)

        self.assertIn(
            b'<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>',
            updated,
        )

    def test_adds_calc_properties_before_ext_list(self):
        workbook = b'<workbook><sheets/><extLst><ext uri="example"/></extLst></workbook>'

        updated = update_workbook_xml(workbook)

        self.assertLess(updated.index(b"<calcPr "), updated.index(b"<extLst>"))

    def test_updates_and_inserts_calc_properties_with_workbook_prefix(self):
        existing = (
            b'<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            b'<x:sheets/><x:calcPr calcMode="manual"/><x:extLst/></x:workbook>'
        )
        missing = (
            b'<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            b'<x:sheets/><x:extLst/></x:workbook>'
        )

        updated_existing = update_workbook_xml(existing)
        updated_missing = update_workbook_xml(missing)

        self.assertIn(b'<x:calcPr calcMode="auto"', updated_existing)
        self.assertIn(b'<x:calcPr calcMode="auto"', updated_missing)
        self.assertNotIn(b'<calcPr ', updated_existing)
        self.assertNotIn(b'<calcPr ', updated_missing)

    def test_shifts_drawing_anchor_columns_at_or_right_of_insertion(self):
        drawing = (
            b'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'
            b'<xdr:from><xdr:col>22</xdr:col></xdr:from>'
            b'<xdr:from><xdr:col>23</xdr:col></xdr:from>'
            b'<xdr:to><xdr:col>24</xdr:col></xdr:to>'
            b'</xdr:wsDr>'
        )

        shifted = shift_drawing_anchors(drawing, "X")

        self.assertEqual(
            re.findall(rb'<xdr:col>(\d+)</xdr:col>', shifted),
            [b"22", b"24", b"25"],
        )


if __name__ == "__main__":
    unittest.main()
