import unittest
import json
import os
import shutil
import subprocess
import tempfile
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

from erp_excel_sync import (
    build_payload,
    changed_fields,
    choose_candidate,
    compare_api_snapshots,
    date_window_ms,
    parse_runtime_args,
    populate_calculated_return_rate,
    populate_derived_return_rate,
    run_monthly_sync,
    set_cell_value,
    sync_sheet_xml,
)


@unittest.skipUnless(shutil.which("cmd.exe"), "Windows cmd.exe is required")
class LauncherTests(unittest.TestCase):
    def test_batch_launcher_passes_script_and_config_to_python(self):
        launcher = Path(__file__).resolve().parents[1] / "run_sync.bat"
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            capture_path = temp_path / "args.txt"
            (temp_path / "python.cmd").write_bytes(
                b'@echo off\r\n> "%LAUNCH_CAPTURE%" echo %*\r\nexit /b 0\r\n'
            )
            environment = os.environ.copy()
            environment["PATH"] = f"{temp_path}{os.pathsep}{environment['PATH']}"
            environment["LAUNCH_CAPTURE"] = str(capture_path)

            result = subprocess.run(
                ["cmd.exe", "/d", "/c", str(launcher)],
                stdin=subprocess.DEVNULL,
                capture_output=True,
                env=environment,
                timeout=10,
            )

            self.assertEqual(result.returncode, 0, result.stderr.decode(errors="replace"))
            self.assertTrue(capture_path.exists(), result.stderr.decode(errors="replace"))
            arguments = capture_path.read_text(encoding="utf-8").strip()
            self.assertEqual(
                arguments,
                f'"{launcher.parent / "erp_excel_sync.py"}" --config "{launcher.parent / "config.json"}"',
            )


class RuntimeConfigTests(unittest.TestCase):
    def test_ignores_underscore_prefixed_comment_fields(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = Path(temp_dir) / "config.json"
            config_path.write_text(
                json.dumps({
                    "_说明": {"workbook": "主工作簿"},
                    "workbook": "master.xlsx",
                }, ensure_ascii=False),
                encoding="utf-8",
            )

            args = parse_runtime_args(["--config", str(config_path)])

            self.assertEqual(args.workbook, Path(temp_dir).resolve() / "master.xlsx")

    def test_loads_master_workbook_and_runtime_directories(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_dir = Path(temp_dir)
            config_path = config_dir / "config.json"
            config_path.write_text(
                json.dumps({
                    "workbook": "outputs/master.xlsx",
                    "snapshot_dir": "snapshots",
                    "report_dir": "reports",
                    "skus_file": "target_skus.txt",
                    "dry_run": False,
                }),
                encoding="utf-8",
            )

            args = parse_runtime_args(["--config", str(config_path)])

            resolved_dir = config_dir.resolve()
            self.assertEqual(args.workbook, resolved_dir / "outputs/master.xlsx")
            self.assertEqual(args.snapshot_dir, resolved_dir / "snapshots")
            self.assertEqual(args.report_dir, resolved_dir / "reports")
            self.assertEqual(args.skus_file, resolved_dir / "target_skus.txt")
            self.assertFalse(args.dry_run)

    def test_rejects_legacy_one_window_configuration(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = Path(temp_dir) / "config.json"
            config_path.write_text(
                json.dumps({
                    "input": "source.xlsx",
                    "output": "result.xlsx",
                    "start_date": "2026-08-01",
                    "end_date": "2026-08-24",
                }),
                encoding="utf-8",
            )

            with self.assertRaisesRegex(SystemExit, "legacy configuration"):
                parse_runtime_args(["--config", str(config_path)])


class DateWindowTests(unittest.TestCase):
    def test_converts_shanghai_dates_to_inclusive_erp_milliseconds(self):
        self.assertEqual(
            date_window_ms("2026-08-01", "2026-08-23"),
            (1785513600000, 1787500799999),
        )

    def test_builds_item_query_payload_for_the_requested_window(self):
        payload = build_payload("2026-08-01", "2026-08-23")

        self.assertEqual(payload["startTime"], "1785513600000")
        self.assertEqual(payload["endTime"], "1787500799999")
        self.assertEqual(payload["queryFlag"], "item")
        self.assertEqual(payload["pageSize"], "2000")
        self.assertEqual(payload["sysStatus"], "created")
        self.assertEqual(payload["tradeTypes"], "3")
        self.assertEqual(payload["asStatus"], "9,2,12")

    def test_builds_before_shipment_return_payload(self):
        payload = build_payload(
            "2026-08-01", "2026-08-31", as_types=("5",)
        )

        self.assertEqual(payload["asStatus"], "9,2,12")
        self.assertEqual(payload["asTypes"], "5")

    def test_builds_overall_return_payload_without_exchange(self):
        payload = build_payload(
            "2026-08-01",
            "2026-08-31",
            as_types=("5", "1", "2", "7", "8", "10"),
        )

        self.assertEqual(payload["asTypes"], "5,1,2,7,8,10")
        self.assertNotIn("4", payload["asTypes"].split(","))

    def test_builds_after_shipment_payload_without_unshipped_or_exchange(self):
        payload = build_payload(
            "2026-08-01",
            "2026-08-31",
            as_types=("1", "2", "7", "8", "10"),
        )

        self.assertEqual(payload["asTypes"], "1,2,7,8,10")
        self.assertNotIn("5", payload["asTypes"].split(","))
        self.assertNotIn("4", payload["asTypes"].split(","))


class MonthlyOrchestrationTests(unittest.TestCase):
    def _args(self, root: Path) -> SimpleNamespace:
        workbook = root / "master.xlsx"
        workbook.write_bytes(b"master")
        (root / "aliases.json").write_text("{}", encoding="utf-8")
        (root / "critical.txt").write_text("", encoding="utf-8")
        return SimpleNamespace(
            workbook=workbook,
            snapshot_dir=root / "snapshots",
            report_dir=root / "reports",
            actual_json=None,
            before_return_json=None,
            after_return_json=None,
            return_json=None,
            company_id="111873",
            cookie_env="ERP_COOKIE",
            aliases=root / "aliases.json",
            skus_file=root / "critical.txt",
            dry_run=False,
        )

    @mock.patch("erp_excel_sync.atomic_replace_master")
    @mock.patch("erp_excel_sync.validate_workbook")
    @mock.patch("erp_excel_sync.prepare_monthly_update")
    @mock.patch("erp_excel_sync.fetch_api")
    def test_fetches_independent_windows_and_writes_reports_before_blocking(
        self, fetch_api_mock, prepare_mock, validate_mock, replace_mock
    ):
        with tempfile.TemporaryDirectory() as temp_dir, mock.patch.dict(
            os.environ, {"ERP_COOKIE": "cookie"}
        ):
            root = Path(temp_dir)
            args = self._args(root)
            fetch_api_mock.side_effect = [
                {"data": {"list": [{"itemOuterId": "A"}]}},
                {"data": {"list": [{"itemOuterId": "B"}]}},
                {"data": {"list": [{"itemOuterId": "C"}]}},
                {"data": {"list": [{"itemOuterId": "D"}]}},
            ]
            prepare_mock.return_value = SimpleNamespace(
                inserted=True,
                formula_count=2,
                rows=(),
                review=(),
                critical_results=({"sku": "7057", "passed": False},),
                critical_failures=({"sku": "7057", "passed": False},),
                replacements={},
            )

            result = run_monthly_sync(args, run_date=date(2026, 9, 16))

            self.assertEqual(result, 2)
            self.assertEqual(
                [
                    (call.args[2:], call.kwargs["as_types"])
                    for call in fetch_api_mock.call_args_list
                ],
                [
                    (("2026-09-01", "2026-09-14"), ()),
                    (("2026-08-01", "2026-08-31"), ("5",)),
                    (
                        ("2026-08-01", "2026-08-31"),
                        ("1", "2", "7", "8", "10"),
                    ),
                    (
                        ("2026-08-01", "2026-08-31"),
                        ("5", "1", "2", "7", "8", "10"),
                    ),
                ],
            )
            self.assertTrue((root / "reports" / "2026-09-15" / "summary.json").exists())
            self.assertTrue((root / "reports" / "2026-09-15" / "critical_skus.csv").exists())
            validate_mock.assert_not_called()
            replace_mock.assert_not_called()

    @mock.patch("erp_excel_sync.atomic_replace_master")
    @mock.patch("erp_excel_sync.validate_workbook")
    @mock.patch("erp_excel_sync.fast_patch_zip")
    @mock.patch("erp_excel_sync.prepare_monthly_update")
    def test_dry_run_never_creates_or_replaces_a_candidate(
        self, prepare_mock, patch_mock, validate_mock, replace_mock
    ):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            args = self._args(root)
            args.dry_run = True
            actual = root / "actual.json"
            returns = root / "return.json"
            before_returns = root / "before-return.json"
            after_returns = root / "after-return.json"
            actual.write_text('{"data":{"list":[]}}', encoding="utf-8")
            returns.write_text('{"data":{"list":[]}}', encoding="utf-8")
            before_returns.write_text('{"data":{"list":[]}}', encoding="utf-8")
            after_returns.write_text('{"data":{"list":[]}}', encoding="utf-8")
            args.actual_json = actual
            args.before_return_json = before_returns
            args.after_return_json = after_returns
            args.return_json = returns
            prepare_mock.return_value = SimpleNamespace(
                inserted=False,
                formula_count=0,
                rows=(),
                review=(),
                critical_results=(),
                critical_failures=(),
                replacements={},
            )

            self.assertEqual(run_monthly_sync(args, run_date=date(2026, 9, 8)), 0)
            patch_mock.assert_not_called()
            validate_mock.assert_not_called()
            replace_mock.assert_not_called()


class CandidateMatchingTests(unittest.TestCase):
    def test_duplicate_exact_sku_with_negligible_name_similarity_is_ambiguous(self):
        result = choose_candidate(
            "SKU-1",
            "alpha",
            [
                {"itemOuterId": "SKU-1", "title": "zzzzza"},
                {"itemOuterId": "SKU-1", "title": "qqqq"},
            ],
            aliases={},
        )

        self.assertEqual(result.status, "ambiguous")
        self.assertFalse(result.auto_write)
        self.assertIsNone(result.candidate)

    def test_duplicate_exact_sku_with_close_name_scores_is_ambiguous(self):
        result = choose_candidate(
            "SKU-1",
            "alpha",
            [
                {"itemOuterId": "SKU-1", "title": "alpa"},
                {"itemOuterId": "SKU-1", "title": "alphi"},
            ],
            aliases={},
        )

        self.assertEqual(result.status, "ambiguous")
        self.assertFalse(result.auto_write)
        self.assertIsNone(result.candidate)

    def test_independent_windows_can_choose_different_records_for_the_same_sheet_row(self):
        actual = choose_candidate(
            "A-1", "Alpha",
            [{"itemOuterId": "A-1", "title": "Alpha", "actualSysConsignCount": 12,
              "customA1ED4F3EEFEF30DBB8E9A9A4823B79A3": "99%"}],
            aliases={},
        )
        returns = choose_candidate(
            "A-1", "Alpha",
            [{"itemOuterId": "A-1", "title": "Alpha", "actualSysConsignCount": 999,
              "customA1ED4F3EEFEF30DBB8E9A9A4823B79A3": "20%"}],
            aliases={},
        )

        self.assertEqual(actual.candidate["actualSysConsignCount"], 12)
        self.assertEqual(
            returns.candidate["customA1ED4F3EEFEF30DBB8E9A9A4823B79A3"], "20%"
        )

    def test_uses_product_name_to_resolve_duplicate_exact_skus(self):
        candidates = [
            {"itemOuterId": "7023", "title": "26ss美式复古水洗短裤", "actualSysConsignCount": 0},
            {"itemOuterId": "7023", "title": "【坠落】复古水洗工装短裤", "actualSysConsignCount": 153},
        ]

        result = choose_candidate("7023", "坠落", candidates, aliases={})

        self.assertEqual(result.status, "exact")
        self.assertTrue(result.auto_write)
        self.assertEqual(result.candidate["actualSysConsignCount"], 153)

    def test_unapproved_fuzzy_sku_is_sent_to_manual_review(self):
        candidates = [
            {"itemOuterId": "7060-1", "title": "【冲锋号•珠宝】复古工装裤"},
        ]

        result = choose_candidate("7060--1", "冲锋号•珠宝", candidates, aliases={})

        self.assertEqual(result.status, "review")
        self.assertFalse(result.auto_write)
        self.assertIn("sku", result.differences)

    def test_approved_alias_is_safe_to_write(self):
        candidates = [
            {"itemOuterId": "DG-169鞋子", "title": "【冰沙芒果】复古帆布鞋"},
        ]

        result = choose_candidate(
            "DG -169",
            "冰沙芒果",
            candidates,
            aliases={"dg-169": "dg-169鞋子"},
        )

        self.assertEqual(result.status, "alias")
        self.assertTrue(result.auto_write)
        self.assertEqual(result.candidate["itemOuterId"], "DG-169鞋子")


class ReviewTests(unittest.TestCase):
    def test_calculated_rate_uses_own_sales_threshold_and_money(self):
        rows = [
            {"itemCount": 49, "rawRefundMoney": 25, "saleMoney": 100},
            {"itemCount": 50, "rawRefundMoney": 25, "saleMoney": 100},
        ]

        populated = populate_calculated_return_rate(rows, "calculatedRate")

        self.assertEqual(populated, 1)
        self.assertIsNone(rows[0]["calculatedRate"])
        self.assertEqual(rows[1]["calculatedRate"], "25.00%")

    def test_calculated_rate_blanks_invalid_or_zero_sale_money(self):
        rows = [
            {"itemCount": 50, "rawRefundMoney": 0, "saleMoney": 0},
            {"itemCount": 50, "rawRefundMoney": 1, "saleMoney": "NaN"},
            {"itemCount": "bad", "rawRefundMoney": 1, "saleMoney": 2},
        ]

        populated = populate_calculated_return_rate(rows, "calculatedRate")

        self.assertEqual(populated, 0)
        self.assertEqual(
            [row["calculatedRate"] for row in rows], [None, None, None]
        )

    def test_derives_custom_return_rate_from_refund_and_sale_money(self):
        rows = [{"rawRefundMoney": 69331.11, "saleMoney": 172911.16}]

        populated = populate_derived_return_rate(rows)

        self.assertEqual(populated, 1)
        self.assertEqual(
            rows[0]["customA1ED4F3EEFEF30DBB8E9A9A4823B79A3"],
            "40.10%",
        )

    def test_zero_money_derives_zero_percent_and_existing_value_is_preserved(self):
        rows = [
            {"rawRefundMoney": 0, "saleMoney": 0},
            {
                "rawRefundMoney": 100,
                "saleMoney": 200,
                "customA1ED4F3EEFEF30DBB8E9A9A4823B79A3": "12.34%",
            },
        ]

        populated = populate_derived_return_rate(rows)

        self.assertEqual(populated, 1)
        self.assertEqual(
            rows[0]["customA1ED4F3EEFEF30DBB8E9A9A4823B79A3"],
            "0.00%",
        )
        self.assertEqual(
            rows[1]["customA1ED4F3EEFEF30DBB8E9A9A4823B79A3"],
            "12.34%",
        )

    def test_lists_only_api_columns_that_changed(self):
        old = {
            "itemOuterId": "7057",
            "actualSysConsignCount": 200,
            "customA1ED4F3EEFEF30DBB8E9A9A4823B79A3": "53.24%",
        }
        new = {
            "itemOuterId": "7057",
            "actualSysConsignCount": 263,
            "customA1ED4F3EEFEF30DBB8E9A9A4823B79A3": "43.44%",
        }

        self.assertEqual(
            changed_fields(old, new),
            {
                "actualSysConsignCount": {"old": 200, "new": 263},
                "customA1ED4F3EEFEF30DBB8E9A9A4823B79A3": {"old": "53.24%", "new": "43.44%"},
            },
        )

    def test_compares_duplicate_skus_by_sku_and_title(self):
        old_rows = [
            {"itemOuterId": "7023", "title": "坠落", "actualSysConsignCount": 100},
            {"itemOuterId": "7023", "title": "定金", "actualSysConsignCount": 0},
        ]
        new_rows = [
            {"itemOuterId": "7023", "title": "坠落", "actualSysConsignCount": 153},
            {"itemOuterId": "7023", "title": "定金", "actualSysConsignCount": 0},
        ]

        differences = compare_api_snapshots(old_rows, new_rows)

        self.assertEqual(len(differences), 1)
        self.assertEqual(differences[0]["itemOuterId"], "7023")
        self.assertEqual(differences[0]["title"], "坠落")
        self.assertEqual(differences[0]["changed_fields"]["actualSysConsignCount"], {"old": 100, "new": 153})


class XmlWriteTests(unittest.TestCase):
    def test_inserts_y_before_aa_with_the_quantity_style(self):
        row = b'<row r="5"><c r="E5" t="s"><v>1</v></c><c r="AA5" s="219"><v>0.1</v></c></row>'

        patched = set_cell_value(row, "Y", 5, 153, default_style="141")

        self.assertIn(
            b'<c r="Y5" s="141"><v>153</v></c><c r="AA5" s="219">',
            patched,
        )

    def test_writes_alias_rows_and_leaves_unapproved_fuzzy_rows_for_review(self):
        sheet_xml = (
            b'<worksheet><sheetData>'
            b'<row r="2"><c r="E2" t="s"><v>0</v></c><c r="F2" t="s"><v>1</v></c>'
            b'<c r="Y2" s="141"><v>228</v></c><c r="AA2" s="219"><v>0.5237</v></c></row>'
            b'<row r="3"><c r="E3" t="s"><v>2</v></c><c r="F3" t="s"><v>3</v></c>'
            b'<c r="Y3" s="141"><v>62</v></c><c r="AA3" s="219"><v>0.557</v></c></row>'
            b'</sheetData></worksheet>'
        )
        api_rows = [
            {
                "itemOuterId": "DG-169鞋子",
                "title": "【冰沙芒果】复古帆布鞋",
                "actualSysConsignCount": 366,
                "customA1ED4F3EEFEF30DBB8E9A9A4823B79A3": "49.89%",
            },
            {
                "itemOuterId": "7060-1",
                "title": "【冲锋号•珠宝】复古工装裤",
                "actualSysConsignCount": 145,
                "customA1ED4F3EEFEF30DBB8E9A9A4823B79A3": "47.94%",
            },
        ]

        modified, report = sync_sheet_xml(
            sheet_xml,
            ["DG -169", "冰沙芒果", "7060--1", "冲锋号•珠宝"],
            api_rows,
            aliases={"dg-169": "dg-169鞋子"},
        )

        self.assertIn(b'<c r="Y2" s="141"><v>366</v></c>', modified)
        self.assertIn(b'<c r="AA2" s="219"><v>0.4989</v></c>', modified)
        self.assertIn(b'<c r="Y3" s="141"><v>62</v></c>', modified)
        self.assertEqual(report["written"], 1)
        self.assertEqual(report["rows"][0]["old_return_rate"], "52.37%")
        self.assertEqual(report["rows"][0]["new_return_rate"], "49.89%")
        self.assertEqual(report["review"][0]["sheet_sku"], "7060--1")


if __name__ == "__main__":
    unittest.main()
