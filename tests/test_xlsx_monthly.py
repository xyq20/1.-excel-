import re
import unittest
from datetime import date
from xml.etree import ElementTree as ET

from monthly_schedule import resolve_sync_cycle
from xlsx_monthly import (
    WorkbookLayout,
    column_label,
    column_number,
    discover_layout,
    insert_month_column,
    normalize_header,
    shift_drawing_anchors,
    shift_formula_references,
    update_workbook_xml,
)


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


class ReferenceShiftTests(unittest.TestCase):
    def test_formula_translation_leaves_quoted_cell_like_text_unchanged(self):
        formula = 'TEXT(Y2-X2,"Y2原文") + SUM($X$3:AA3)'

        self.assertEqual(
            shift_formula_references(formula, "X"),
            'TEXT(Z2-Y2,"Y2原文") + SUM($Y$3:AB3)',
        )

    def test_formula_translation_does_not_treat_numbered_function_names_as_cells(self):
        self.assertEqual(shift_formula_references("LOG10(X2)", "X"), "LOG10(Y2)")

    def test_formula_translation_preserves_quoted_sheet_names(self):
        self.assertEqual(
            shift_formula_references("SUM('Y2 data'!X2)", "X"),
            "SUM('Y2 data'!Y2)",
        )

    def test_formula_translation_shifts_whole_column_ranges(self):
        self.assertEqual(
            shift_formula_references('FILTER($AS:$AS,$AR:$AR=F4)+SUM("$X:$X")', "X"),
            'FILTER($AT:$AT,$AS:$AS=F4)+SUM("$X:$X")',
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
