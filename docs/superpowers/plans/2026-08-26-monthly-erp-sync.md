# Monthly ERP Workbook Sync Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (- [ ]) syntax for tracking.

**Goal:** Replace the fixed-date, fixed-column updater with an idempotent 1st/15th workflow that updates the full master workbook, rolls month columns and formulas, and preserves one recoverable backup.

**Architecture:** Put date-window and label logic in a pure monthly_schedule.py module. Put byte-oriented Open XML layout discovery, column insertion, reference shifting, ZIP validation, and atomic replacement in xlsx_monthly.py; keep ERP HTTP, SKU matching, reporting, and orchestration in erp_excel_sync.py, passing separate actual-quantity and return-rate datasets.

**Tech Stack:** Python 3.10+ standard library, unittest, XLSX/Open XML, Windows batch launcher

---

## File structure

- Create monthly_schedule.py for node selection, two query windows, headers, and formulas.
- Create xlsx_monthly.py for layout discovery, XML mutation, fast multi-part ZIP patching, validation, backup, and rollback.
- Modify erp_excel_sync.py for two ERP requests, full-table matching, critical-SKU validation, snapshots, and reports.
- Modify config.json, README.md, and run_sync.bat for one long-lived master workbook.
- Create tests/test_monthly_schedule.py and tests/test_xlsx_monthly.py.
- Modify tests/test_erp_excel_sync.py for portable launcher, dual-query orchestration, and new config behavior.

## Task 1: Make the current test baseline portable

**Files:**
- Modify: tests/test_erp_excel_sync.py:1-48

- [ ] **Step 1: Guard the Windows-only launcher test**

Add import shutil and decorate the launcher class:

~~~python
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
            self.assertTrue(capture_path.exists())
~~~

- [ ] **Step 2: Run the baseline**

~~~bash
'/Users/linchaoyang/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3' -m unittest tests.test_erp_excel_sync -v
~~~

Expected: 12 tests pass and one launcher test is skipped on macOS; no failures or errors.

- [ ] **Step 3: Commit**

~~~bash
git add tests/test_erp_excel_sync.py
git commit -m "test: skip Windows launcher test off Windows"
~~~

## Task 2: Implement the 1st/15th schedule as pure functions

**Files:**
- Create: monthly_schedule.py
- Create: tests/test_monthly_schedule.py

- [ ] **Step 1: Write failing schedule tests**

Create tests/test_monthly_schedule.py:

~~~python
import unittest
from datetime import date

from monthly_schedule import (
    DateWindow,
    actual_header,
    change_formula,
    peer_formula,
    resolve_sync_cycle,
    return_header,
)


class ScheduleTests(unittest.TestCase):
    def test_first_node_windows(self):
        cycle = resolve_sync_cycle(date(2026, 9, 8))
        self.assertEqual(cycle.node_date, date(2026, 9, 1))
        self.assertEqual(cycle.actual_window, DateWindow(date(2026, 8, 1), date(2026, 8, 31)))
        self.assertEqual(cycle.return_window, DateWindow(date(2026, 7, 15), date(2026, 8, 15)))
        self.assertEqual(actual_header(cycle), "8月实发")
        self.assertEqual(return_header(cycle), "退货率（7.15-8.15）")

    def test_fifteenth_node_windows(self):
        cycle = resolve_sync_cycle(date(2026, 9, 26))
        self.assertEqual(cycle.node_date, date(2026, 9, 15))
        self.assertEqual(cycle.actual_window, DateWindow(date(2026, 9, 1), date(2026, 9, 14)))
        self.assertEqual(cycle.return_window, DateWindow(date(2026, 8, 1), date(2026, 8, 31)))
        self.assertEqual(actual_header(cycle), "9月实发（9.15）")
        self.assertEqual(return_header(cycle), "退货率（8.1-8.31）")

    def test_january_crosses_year_boundary(self):
        cycle = resolve_sync_cycle(date(2027, 1, 1))
        self.assertEqual(cycle.actual_window, DateWindow(date(2026, 12, 1), date(2026, 12, 31)))
        self.assertEqual(cycle.return_window, DateWindow(date(2026, 11, 15), date(2026, 12, 15)))

    def test_peer_formula_uses_real_month_length(self):
        september = resolve_sync_cycle(date(2026, 9, 15))
        march_leap = resolve_sync_cycle(date(2028, 3, 15))
        self.assertEqual(peer_formula("X", 9, september), "X9/31*14")
        self.assertEqual(peer_formula("X", 9, march_leap), "X9/29*14")

    def test_change_formula_uses_current_month(self):
        cycle = resolve_sync_cycle(date(2026, 10, 15))
        self.assertEqual(
            change_formula("AA", "Z", 9, cycle),
            'TEXT(AA9-Z9,"10月增加0件；10月减少0件；持平")',
        )


if __name__ == "__main__":
    unittest.main()
~~~

- [ ] **Step 2: Verify failure before implementation**

~~~bash
'/Users/linchaoyang/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3' -m unittest tests.test_monthly_schedule -v
~~~

Expected: ModuleNotFoundError for monthly_schedule.

- [ ] **Step 3: Implement monthly_schedule.py**

~~~python
from calendar import monthrange
from dataclasses import dataclass
from datetime import date, timedelta


@dataclass(frozen=True)
class DateWindow:
    start: date
    end: date

    def iso(self) -> tuple[str, str]:
        return self.start.isoformat(), self.end.isoformat()


@dataclass(frozen=True)
class SyncCycle:
    run_date: date
    node_date: date
    kind: str
    actual_month: date
    previous_month: date
    actual_window: DateWindow
    return_window: DateWindow


def month_start(value: date) -> date:
    return value.replace(day=1)


def previous_month_start(value: date) -> date:
    return (month_start(value) - timedelta(days=1)).replace(day=1)


def month_end(value: date) -> date:
    return date(value.year, value.month, monthrange(value.year, value.month)[1])


def resolve_sync_cycle(run_date: date) -> SyncCycle:
    current = month_start(run_date)
    previous = previous_month_start(current)
    if run_date.day < 15:
        older = previous_month_start(previous)
        return SyncCycle(
            run_date, current, "first", previous, older,
            DateWindow(previous, month_end(previous)),
            DateWindow(older.replace(day=15), previous.replace(day=15)),
        )
    return SyncCycle(
        run_date, current.replace(day=15), "fifteenth", current, previous,
        DateWindow(current, current.replace(day=14)),
        DateWindow(previous, month_end(previous)),
    )


def actual_header(cycle: SyncCycle) -> str:
    month = cycle.actual_month.month
    return f"{month}月实发" if cycle.kind == "first" else f"{month}月实发（{month}.15）"


def return_header(cycle: SyncCycle) -> str:
    start, end = cycle.return_window.start, cycle.return_window.end
    return f"退货率（{start.month}.{start.day}-{end.month}.{end.day}）"


def peer_formula(previous_col: str, row: int, cycle: SyncCycle) -> str:
    days = monthrange(cycle.previous_month.year, cycle.previous_month.month)[1]
    return f"{previous_col}{row}/{days}*14"


def change_formula(actual_col: str, peer_col: str, row: int, cycle: SyncCycle) -> str:
    month = cycle.actual_month.month
    return f'TEXT({actual_col}{row}-{peer_col}{row},"{month}月增加0件；{month}月减少0件；持平")'
~~~

- [ ] **Step 4: Run tests and commit**

Expected: 5 schedule tests pass.

~~~bash
git add monthly_schedule.py tests/test_monthly_schedule.py
git commit -m "feat: resolve monthly ERP sync windows"
~~~

## Task 3: Discover workbook layout and roll one month column

**Files:**
- Create: xlsx_monthly.py
- Create: tests/test_xlsx_monthly.py
- Modify: erp_excel_sync.py:196-259

- [ ] **Step 1: Write failing layout, translation, and idempotency tests**

Create tests/test_xlsx_monthly.py:

~~~python
import unittest
from datetime import date

from monthly_schedule import resolve_sync_cycle
from xlsx_monthly import discover_layout, insert_month_column, normalize_header, shift_formula_references


SHEET = b'''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="B1:AY10"/><cols><col min="23" max="27" width="12"/></cols>
<sheetData><row r="1">
<c r="W1" t="s"><v>0</v></c><c r="X1" t="s"><v>1</v></c>
<c r="Y1" t="s"><v>2</v></c><c r="Z1" t="s"><v>3</v></c>
<c r="AA1" t="s"><v>4</v></c></row>
<row r="2"><c r="E2" t="s"><v>5</v></c><c r="F2" t="s"><v>6</v></c>
<c r="W2"><v>310</v></c><c r="X2"><f>W2/31*14</f><v>140</v></c>
<c r="Y2"><v>64</v></c><c r="Z2"><f>TEXT(Y2-X2,"8月增加0件；8月减少0件；持平")</f></c>
<c r="AA2"><v>0.51</v></c></row></sheetData><autoFilter ref="B1:AY10"/></worksheet>'''
STRINGS = ["7月实发 ", "同期销量", "8月实发(8.15) ", "变化情况", "退货率（7.1-7.31）", "7057", "黑幕"]


class HeaderTests(unittest.TestCase):
    def test_normalizes_historical_variants(self):
        self.assertEqual(normalize_header(" 8月实发(8.15) "), "8月实发（8.15）")
        self.assertEqual(normalize_header("8月实销"), "8月实发")
        self.assertEqual(normalize_header("8月实发销量"), "8月实发")

    def test_discovers_current_layout(self):
        layout = discover_layout(SHEET, STRINGS, resolve_sync_cycle(date(2026, 8, 26)))
        self.assertEqual(
            (layout.previous_col, layout.peer_col, layout.actual_col, layout.change_col, layout.return_col),
            ("W", "X", "Y", "Z", "AA"),
        )
        self.assertFalse(layout.needs_insert)


class ColumnInsertionTests(unittest.TestCase):
    def test_formula_translation_skips_quoted_text(self):
        self.assertEqual(
            shift_formula_references('TEXT(Y2-X2,"Y2原文")+SUM($X$3:AA3)', "X"),
            'TEXT(Z2-Y2,"Y2原文")+SUM($Y$3:AB3)',
        )

    def test_september_insert_moves_the_rolling_block(self):
        result = insert_month_column(SHEET, STRINGS, resolve_sync_cycle(date(2026, 9, 15)))
        self.assertTrue(result.inserted)
        self.assertEqual(
            (result.layout.previous_col, result.layout.peer_col, result.layout.actual_col,
             result.layout.change_col, result.layout.return_col),
            ("X", "Y", "Z", "AA", "AB"),
        )
        self.assertIn(b'autoFilter ref="B1:AZ10"', result.sheet_xml)

    def test_same_cycle_is_idempotent(self):
        cycle = resolve_sync_cycle(date(2026, 9, 15))
        first = insert_month_column(SHEET, STRINGS, cycle)
        second = insert_month_column(first.sheet_xml, STRINGS, cycle)
        self.assertFalse(second.inserted)
        self.assertEqual(first.sheet_xml, second.sheet_xml)
~~~

- [ ] **Step 2: Verify failure before implementation**

~~~bash
'/Users/linchaoyang/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3' -m unittest tests.test_xlsx_monthly -v
~~~

Expected: ModuleNotFoundError for xlsx_monthly.

- [ ] **Step 3: Implement header and layout contracts**

~~~python
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


def normalize_header(value: str) -> str:
    text = re.sub(r"\s+", "", value).replace("(", "（").replace(")", "）")
    return text.replace("实发销量", "实发").replace("实销", "实发")
~~~

discover_layout must read shared-string and inline-string headers; require exactly one 同期销量, 变化情况, and 退货率 prefix; recognize 实发, 实销, and 实发销量 history; and return needs_insert only when a 15th-node workbook lacks the current staged header. A first-node run locates the matching staged header and never inserts.

- [ ] **Step 4: Implement quote-safe A1 translation**

~~~python
A1_RE = re.compile(r"(?<![A-Z0-9_])(?P<ca>\$?)(?P<col>[A-Z]{1,3})(?P<ra>\$?)(?P<row>\d+)")


def shift_formula_references(formula: str, insert_before_col: str) -> str:
    insert_number = column_number(insert_before_col)

    def shift_plain(segment: str) -> str:
        def replace(match: re.Match[str]) -> str:
            number = column_number(match.group("col"))
            col = column_label(number + 1) if number >= insert_number else match.group("col")
            return f"{match.group('ca')}{col}{match.group('ra')}{match.group('row')}"
        return A1_RE.sub(replace, segment)

    pieces: list[str] = []
    for index, segment in enumerate(re.split(r'("(?:[^"]|"")*")', formula)):
        pieces.append(segment if index % 2 else shift_plain(segment))
    return "".join(pieces)
~~~

Add range translation for ref, sqref, activeCell, topLeftCell, dimensions, auto-filter, merges, hyperlinks, data validations, conditional formatting, panes, selections, and workbook defined names such as 分级总表!$B$1:$AY$561. Detect drawing anchors at or after the insertion point and shift them if present.

- [ ] **Step 5: Implement byte-oriented insertion**

insert_month_column must perform exactly these actions:

1. Return unchanged bytes when the current staged header already exists.
2. Insert before 同期销量 for a new 15th node.
3. Shift every cell and right-side column one place.
4. Split col min/max spans and copy width/style from the old current-actual column.
5. Copy the shifted old actual values/styles into the inserted previous-month column.
6. Write previous and current headers as inline strings.
7. Hide month columns older than the inserted previous month.
8. Keep previous, peer, current, change, and return columns visible.
9. Rediscover and return the new layout.

Keep this transformation byte-oriented because real sheet1.xml is about 82MB uncompressed and contains WPS extensions.

- [ ] **Step 6: Add workbook calculation properties**

~~~python
def update_workbook_xml(workbook_xml: bytes, insert_before_col: str | None) -> bytes:
    updated = shift_defined_names_for_sheet(workbook_xml, "分级总表", insert_before_col)
    return set_calc_properties(
        updated,
        calc_mode="auto",
        full_calc_on_load=True,
        force_full_calc=True,
    )
~~~

- [ ] **Step 7: Run tests and commit**

Expected: header, reference, insertion, hiding, defined-name, and idempotency tests pass.

~~~bash
git add xlsx_monthly.py tests/test_xlsx_monthly.py erp_excel_sync.py
git commit -m "feat: roll monthly workbook columns safely"
~~~

## Task 4: Write formulas and two ERP datasets across the full table

**Files:**
- Modify: xlsx_monthly.py
- Modify: erp_excel_sync.py:76-345
- Modify: tests/test_xlsx_monthly.py
- Modify: tests/test_erp_excel_sync.py

- [ ] **Step 1: Add failing dual-dataset tests**

~~~python
class DualDatasetTests(unittest.TestCase):
    def test_uses_separate_actual_and_return_rows(self):
        actual_rows = [{"itemOuterId": "7057", "title": "黑幕", "actualSysConsignCount": 88}]
        return_rows = [{
            "itemOuterId": "7057", "title": "黑幕",
            "customA1ED4F3EEFEF30DBB8E9A9A4823B79A3": "51.85%",
        }]
        result = apply_monthly_values(
            SHEET, STRINGS, resolve_sync_cycle(date(2026, 8, 26)),
            actual_rows, return_rows, aliases={}, critical_skus={"7057"}, matcher=choose_candidate,
        )
        self.assertIn(b"<v>88</v>", result.sheet_xml)
        self.assertIn(b"<v>0.5185</v>", result.sheet_xml)
        self.assertEqual(result.critical_failures, [])

    def test_missing_critical_sku_blocks_commit(self):
        result = apply_monthly_values(
            SHEET, STRINGS, resolve_sync_cycle(date(2026, 8, 26)),
            [], [], aliases={}, critical_skus={"7057"}, matcher=choose_candidate,
        )
        self.assertEqual(result.critical_failures[0]["sku"], "7057")
        self.assertEqual(result.critical_failures[0]["actual_status"], "unmatched")
        self.assertEqual(result.critical_failures[0]["return_status"], "unmatched")
~~~

- [ ] **Step 2: Verify the new tests fail**

Expected: missing apply_monthly_values.

- [ ] **Step 3: Implement full-table write result**

~~~python
@dataclass(frozen=True)
class MonthlyWriteResult:
    sheet_xml: bytes
    layout: WorkbookLayout
    inserted: bool
    rows: list[dict[str, object]]
    review: list[dict[str, object]]
    critical_results: list[dict[str, object]]
    critical_failures: list[dict[str, object]]
    formula_count: int
~~~

apply_monthly_values iterates every E-column SKU; target_skus.txt never filters ordinary rows. It matches each row independently against actual and return datasets, writes only exact/alias matches, retains old values for review/ambiguous rows, and records both statuses. It then writes peer_formula(previous_col, row, cycle) and change_formula(actual_col, peer_col, row, cycle) for every SKU row.

For critical SKUs, record sheet presence, both match statuses, matched ERP IDs, and pass/fail. Return failures as report data so reports can be written before the orchestrator exits.

- [ ] **Step 4: Handle first-node header completion**

On a first-node run, overwrite staged actual values with the previous full month, rename 8月实发（8.15） or 8月实发(8.15) to 8月实发, update the return header, and keep the column structure unchanged.

- [ ] **Step 5: Run selected tests and commit**

~~~bash
'/Users/linchaoyang/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3' -m unittest tests.test_erp_excel_sync.CandidateMatchingTests tests.test_xlsx_monthly.DualDatasetTests -v
git add erp_excel_sync.py xlsx_monthly.py tests/test_erp_excel_sync.py tests/test_xlsx_monthly.py
git commit -m "feat: sync full workbook from two ERP windows"
~~~

Expected: exact/alias matching, independent values, formulas, and critical validation pass.

## Task 5: Patch ZIP parts, validate, back up, and roll back

**Files:**
- Modify: xlsx_monthly.py
- Modify: erp_excel_sync.py:439-509
- Modify: tests/test_xlsx_monthly.py

- [ ] **Step 1: Add failing workbook safety tests**

~~~python
class WorkbookSafetyTests(unittest.TestCase):
    def test_fast_patch_preserves_unmodified_members(self):
        fast_patch_zip(self.source, self.output, {
            "xl/worksheets/sheet1.xml": b"<worksheet>changed</worksheet>",
            "xl/workbook.xml": b"<workbook>changed</workbook>",
        })
        with zipfile.ZipFile(self.output) as archive:
            self.assertEqual(archive.read("xl/media/image1.png"), b"image-bytes")
            self.assertIsNone(archive.testzip())

    def test_atomic_replace_keeps_previous_master(self):
        atomic_replace_master(self.master, self.validated_temp)
        self.assertEqual(self.master.read_bytes(), b"new")
        self.assertEqual(self.master.with_suffix(".xlsx.bak").read_bytes(), b"old")

    def test_failed_second_replace_restores_master(self):
        with self.assertRaises(OSError):
            atomic_replace_master(self.master, self.validated_temp, replace=FailOnSecondReplace())
        self.assertEqual(self.master.read_bytes(), b"old")
~~~

- [ ] **Step 2: Verify tests fail**

Expected: missing generalized fast_patch_zip and atomic_replace_master.

- [ ] **Step 3: Generalize fast ZIP patching**

~~~python
def fast_patch_zip(source: Path, output: Path, replacements: dict[str, bytes]) -> None:
    """Copy untouched local records byte-for-byte and recompress replacement parts."""
~~~

Rebuild central-directory CRC, compressed size, uncompressed size, and local-header offset for replaced parts. Preserve member order, timestamps, flags, extras, comments, and compression methods. Reject encrypted, multi-disk, missing-target, and ZIP64 inputs explicitly; the current 270MB workbook is standard ZIP.

- [ ] **Step 4: Implement candidate validation**

validate_workbook(source, candidate, cycle) runs ZipFile.testzip(), rediscovers layout, confirms actual/return headers, verifies formula references use discovered columns, and compares media member names/CRCs between source and candidate. It returns a JSON-serializable summary or raises RuntimeError.

- [ ] **Step 5: Implement atomic replacement**

~~~python
def atomic_replace_master(master: Path, validated_temp: Path, replace=os.replace) -> Path:
    backup = master.with_suffix(master.suffix + ".bak")
    replace(master, backup)
    try:
        replace(validated_temp, master)
    except BaseException:
        replace(backup, master)
        raise
    return backup
~~~

Create the temporary XLSX in the master directory. A WPS/Excel lock must fail before replacement. Keep a failed validation candidate for diagnosis; successful replacement consumes the temp file.

- [ ] **Step 6: Run tests and commit**

~~~bash
'/Users/linchaoyang/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3' -m unittest tests.test_xlsx_monthly.WorkbookSafetyTests -v
git add erp_excel_sync.py xlsx_monthly.py tests/test_xlsx_monthly.py
git commit -m "feat: validate and atomically replace master workbook"
~~~

## Task 6: Orchestrate dual queries, snapshots, reports, and configuration

**Files:**
- Modify: erp_excel_sync.py:349-422,512-707
- Modify: tests/test_erp_excel_sync.py
- Modify: config.json

- [ ] **Step 1: Add failing config and orchestration tests**

Test that config resolves workbook, snapshot_dir, report_dir, aliases, and critical SKU paths without manual dates. Mock fetch_api and assert one call uses cycle.actual_window and one uses cycle.return_window. Prove reports are written and atomic_replace_master is not called when a critical SKU fails.

- [ ] **Step 2: Verify failure before orchestration changes**

Expected: config rejects workbook/snapshot_dir and the old main performs one query.

- [ ] **Step 3: Replace manual chaining config**

~~~json
{
  "workbook": "outputs/26年8月分级总表_API同步.xlsx",
  "company_id": "111873",
  "cookie_env": "ERP_COOKIE",
  "skus_file": "target_skus.txt",
  "aliases": "sku_aliases.json",
  "snapshot_dir": "snapshots",
  "report_dir": "reports",
  "dry_run": false
}
~~~

Add --workbook, --snapshot-dir, --actual-json, and --return-json. Reject legacy --input, --output, --start, --end, and --api-json with a migration message so one-window behavior cannot run silently.

- [ ] **Step 4: Implement main orchestration**

Execute exactly:

1. Resolve cycle from datetime.now(SHANGHAI_TZ).date().
2. Validate master path.
3. Read Cookie once.
4. Fetch/load actual and return documents with their own windows.
5. Build `node_dir = cycle.node_date.isoformat()` and save snapshots with `f"actual_{cycle.actual_window.start.isoformat()}_{cycle.actual_window.end.isoformat()}.json"` and `f"return_{cycle.return_window.start.isoformat()}_{cycle.return_window.end.isoformat()}.json"` under `snapshot_dir / node_dir`.
6. Load aliases and critical SKUs.
7. Build modified parts and reports in memory.
8. Under `report_dir / node_dir`, write actual_changes.csv, return_changes.csv, api_actual_changes.csv, api_return_changes.csv, manual_review.csv, critical_skus.csv, and summary.json. Compare each API result only with the previous snapshot of the same kind and date-window rule.
9. On critical failure, exit nonzero after reports without replacing the master.
10. On dry run, exit zero after reports without replacing the master.
11. Otherwise create same-directory temp XLSX, validate it, and atomically replace master.
12. Update summary with node, both windows, API counts, row counts, inserted/reused state, formula count, critical result, master, backup, and validation.

- [ ] **Step 5: Run tests and commit**

~~~bash
'/Users/linchaoyang/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3' -m unittest tests.test_erp_excel_sync -v
git add erp_excel_sync.py config.json tests/test_erp_excel_sync.py
git commit -m "feat: orchestrate automatic first and fifteenth syncs"
~~~

Expected: all tests pass except the Windows launcher skip on macOS.

## Task 7: Update launcher and README

**Files:**
- Modify: README.md
- Modify: run_sync.bat
- Modify: tests/test_erp_excel_sync.py

- [ ] **Step 1: Update launcher banner**

Keep:

~~~bat
python "%~dp0erp_excel_sync.py" --config "%~dp0config.json"
~~~

State that the nearest 1st/15th node is automatic, the master is replaced only after validation, and the previous version remains as .bak. Preserve the exact script/config argument assertion.

- [ ] **Step 2: Rewrite operating instructions**

Document both node windows, nearest-node behavior, dynamic columns, formula rules, idempotent reruns, full-table exact/alias matching, target_skus.txt blocking validation, .bak recovery, dry run, report/snapshot paths, separate --actual-json/--return-json offline inputs, and Python 3.10+ tests.

- [ ] **Step 3: Run tests and commit**

~~~bash
'/Users/linchaoyang/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3' -m unittest tests.test_erp_excel_sync.LauncherTests tests.test_erp_excel_sync.RuntimeConfigTests -v
git add README.md run_sync.bat tests/test_erp_excel_sync.py
git commit -m "docs: explain automatic monthly workbook sync"
~~~

## Task 8: Verify with tests and a real workbook copy

**Files:**
- Modify only when a verification test exposes a defect
- Do not commit generated XLSX, snapshots, reports, or backups

- [ ] **Step 1: Run the full portable suite**

~~~bash
'/Users/linchaoyang/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3' -m unittest discover -s tests -v
~~~

Expected: all non-Windows tests pass; the cmd.exe launcher is the only permitted skip on macOS.

- [ ] **Step 2: Run offline dry-run with two API documents**

Run this exact structural dry-run against the saved local response:

~~~bash
'/Users/linchaoyang/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3' erp_excel_sync.py --config config.json --dry-run --actual-json erp_dimensions_0824.json --return-json erp_dimensions_0824.json
~~~

This command tests the independent input paths and full-table reporting without making a network request; window calculations remain covered by Task 2 tests.

Expected: no master/backup modification; node reports contain both windows and every critical SKU passes.

- [ ] **Step 3: Run against a copy of the 270MB workbook**

Create a unique temporary directory, copy the master there, point a temporary config at the copy, and run with the two saved API documents. Never use the user’s only master for this test.

Expected: master and backup pass ZipFile.testzip(), row count is unchanged, media names/CRCs match, headers follow the approved order, and a second run adds no month column.

- [ ] **Step 4: Inspect the copy in WPS**

Confirm only 上月实发｜同期销量｜本月实发｜变化情况｜退货率 is visible; older months are hidden; formulas recalculate; filters, images, and right-side business columns remain aligned; critical SKU values match offline JSON.

- [ ] **Step 5: Verify backup recovery**

Move the generated master aside, restore .xlsx.bak to .xlsx, rerun CRC validation, and open it in WPS. Restore the generated copy after the recovery test.

- [ ] **Step 6: Review final state**

~~~bash
git diff --check
git status --short
~~~

If verification exposes a defect, add a failing regression test, implement the correction, rerun Steps 1–5, and commit the verified fix. Leave large generated artifacts untracked.
