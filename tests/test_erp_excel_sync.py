import unittest
import json
import os
import subprocess
import tempfile
from pathlib import Path

from erp_excel_sync import (
    build_payload,
    changed_fields,
    choose_candidate,
    compare_api_snapshots,
    date_window_ms,
    parse_runtime_args,
    set_cell_value,
    sync_sheet_xml,
)


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
                    "_说明": {"start_date": "开始日期"},
                    "start_date": "2026-08-01",
                    "end_date": "2026-08-24",
                    "input": "source.xlsx",
                    "output": "result.xlsx",
                }, ensure_ascii=False),
                encoding="utf-8",
            )

            args = parse_runtime_args(["--config", str(config_path)])

            self.assertEqual(args.start, "2026-08-01")
            self.assertEqual(args.end, "2026-08-24")

    def test_loads_run_parameters_from_config_and_resolves_relative_paths(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_dir = Path(temp_dir)
            config_path = config_dir / "config.json"
            config_path.write_text(
                json.dumps({
                    "input": "source.xlsx",
                    "output": "outputs/result.xlsx",
                    "start_date": "2026-08-01",
                    "end_date": "2026-08-24",
                    "skus_file": "target_skus.txt",
                    "dry_run": False,
                }),
                encoding="utf-8",
            )

            args = parse_runtime_args(["--config", str(config_path)])

            resolved_dir = config_dir.resolve()
            self.assertEqual(args.input, resolved_dir / "source.xlsx")
            self.assertEqual(args.output, resolved_dir / "outputs/result.xlsx")
            self.assertEqual(args.skus_file, resolved_dir / "target_skus.txt")
            self.assertEqual(args.start, "2026-08-01")
            self.assertEqual(args.end, "2026-08-24")
            self.assertFalse(args.dry_run)

    def test_command_line_dates_override_config_dates(self):
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

            args = parse_runtime_args([
                "--config", str(config_path),
                "--start", "2026-08-05",
                "--end", "2026-08-25",
            ])

            self.assertEqual(args.start, "2026-08-05")
            self.assertEqual(args.end, "2026-08-25")


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
        self.assertEqual(payload["tradeTypes"], "")
        self.assertEqual(payload["asStatus"], "9,2,12")


class CandidateMatchingTests(unittest.TestCase):
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
