# Three Independent Return Rates Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Independently query, calculate, report, and write the before-shipment and overall return rates with a per-query sales threshold of 50, while keeping the disabled after-shipment rate blank.

**Architecture:** Add explicit return-rate profiles at the ERP boundary, pass each profile's rows independently through matching and worksheet writing, and extend the workbook layout from one return-rate column to three named columns. Keep the after-shipment profile disabled but model it in the layout and reporting so activation later only requires a status-code configuration change.

**Tech Stack:** Python 3.10+, `unittest`, ERP form API, direct XLSX ZIP/XML patching.

---

## File map

- Modify `erp_excel_sync.py`: ERP status profiles, request payloads, rate calculation, orchestration, snapshots, reports, and offline arguments.
- Modify `monthly_schedule.py`: generate all three dated return-rate headers from the same date window.
- Modify `xlsx_monthly.py`: discover three columns, preserve them through insertion, write or clear each independent value, and validate headers.
- Modify `tests/test_erp_excel_sync.py`: request-code, threshold, orchestration, snapshot, and report tests.
- Modify `tests/test_monthly_schedule.py`: three header-name tests.
- Modify `tests/test_xlsx_monthly.py`: three-column layout, insertion, writing, clearing, critical-SKU, and validation tests.
- Modify `README.md`: describe the new query filters, threshold, offline inputs, and disabled after-shipment column.

### Task 1: Encode ERP query profiles and rate calculation

**Files:**
- Modify: `erp_excel_sync.py:38-103,374-541`
- Test: `tests/test_erp_excel_sync.py:114-132,313-350`

- [ ] **Step 1: Write failing payload and calculation tests**

Add tests that assert `asStatus` remains the work-order status filter `9,2,12`, while `asTypes` is profile-specific:

```python
def test_builds_before_shipment_return_payload(self):
    payload = build_payload("2026-08-01", "2026-08-31", as_types=("5",))
    self.assertEqual(payload["asStatus"], "9,2,12")
    self.assertEqual(payload["asTypes"], "5")

def test_builds_overall_return_payload_without_exchange(self):
    payload = build_payload(
        "2026-08-01", "2026-08-31", as_types=("5", "1", "2", "7", "8", "10")
    )
    self.assertEqual(payload["asTypes"], "5,1,2,7,8,10")
    self.assertNotIn("4", payload["asTypes"].split(","))

def test_rate_uses_own_sales_threshold_and_money(self):
    rows = [
        {"itemCount": 49, "rawRefundMoney": 25, "saleMoney": 100},
        {"itemCount": 50, "rawRefundMoney": 25, "saleMoney": 100},
    ]
    populated = populate_calculated_return_rate(rows, "calculatedRate")
    self.assertEqual(populated, 1)
    self.assertIsNone(rows[0]["calculatedRate"])
    self.assertEqual(rows[1]["calculatedRate"], "25.00%")
```

- [ ] **Step 2: Run focused tests and verify failure**

Run: `python -m unittest tests.test_erp_excel_sync.DateWindowTests tests.test_erp_excel_sync.ReviewTests -v`

Expected: FAIL because `build_payload` lacks `as_types` and `populate_calculated_return_rate` does not exist.

- [ ] **Step 3: Add profiles, exact ERP codes, and calculation**

Implement these constants and functions in `erp_excel_sync.py`:

```python
@dataclass(frozen=True)
class ReturnRateProfile:
    key: str
    header_prefix: str
    as_types: tuple[str, ...]
    value_field: str
    enabled: bool = True


BEFORE_RETURN_FIELD = "calculatedBeforeShipmentReturnRate"
AFTER_RETURN_FIELD = "calculatedAfterShipmentReturnRate"
OVERALL_RETURN_FIELD = "calculatedOverallReturnRate"
BEFORE_RETURN_PROFILE = ReturnRateProfile(
    "before_shipment", "发货前退货率", ("5",), BEFORE_RETURN_FIELD
)
AFTER_RETURN_PROFILE = ReturnRateProfile(
    "after_shipment", "发货后退货率", (), AFTER_RETURN_FIELD, False
)
OVERALL_RETURN_PROFILE = ReturnRateProfile(
    "overall", "退货率", ("5", "1", "2", "7", "8", "10"), OVERALL_RETURN_FIELD
)
RETURN_RATE_PROFILES = (
    BEFORE_RETURN_PROFILE, AFTER_RETURN_PROFILE, OVERALL_RETURN_PROFILE
)


def build_payload(
    start_date: str,
    end_date: str,
    *,
    as_types: tuple[str, ...] = (),
) -> dict[str, str]:
    start_ms, end_ms = date_window_ms(start_date, end_date)
    payload = dict(parse_qsl(BASE_FORM_BODY, keep_blank_values=True))
    payload["startTime"] = str(start_ms)
    payload["endTime"] = str(end_ms)
    payload["asTypes"] = ",".join(as_types)
    return payload


def populate_calculated_return_rate(
    rows: list[dict[str, Any]],
    output_field: str,
    minimum_sales: int = 50,
) -> int:
    populated = 0
    for row in rows:
        row[output_field] = None
        try:
            sales_count = Decimal(str(row["itemCount"]))
            refund_money = Decimal(str(row["rawRefundMoney"]))
            sale_money = Decimal(str(row["saleMoney"]))
        except (KeyError, InvalidOperation, TypeError, ValueError):
            continue
        if not all(value.is_finite() for value in (sales_count, refund_money, sale_money)):
            continue
        if sales_count < minimum_sales or sale_money <= 0:
            continue
        percentage = refund_money / sale_money * 100
        rounded = percentage.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        row[output_field] = f"{rounded:.2f}%"
        populated += 1
    return populated
```

Update `fetch_api` and `fetch_api_via_chrome` to accept keyword-only `as_types: tuple[str, ...] = ()`, pass it to `build_payload`, and store `asTypes` in `syncMeta`.

- [ ] **Step 4: Run focused tests and verify pass**

Run: `python -m unittest tests.test_erp_excel_sync.DateWindowTests tests.test_erp_excel_sync.ReviewTests -v`

Expected: PASS, including 49 blank, 50 calculated, zero/invalid sale money blank, and the exact ERP codes `5,1,2,7,8,10` with exchange code `4` absent.

- [ ] **Step 5: Commit query profiles**

```bash
git add erp_excel_sync.py tests/test_erp_excel_sync.py
git commit -m "feat: add independent ERP return rate profiles"
```

### Task 2: Extend dated headers and workbook layout to three columns

**Files:**
- Modify: `monthly_schedule.py:71-75`
- Modify: `xlsx_monthly.py:44-55,127-230,353-463,929-1009`
- Test: `tests/test_monthly_schedule.py:69-80`
- Test: `tests/test_xlsx_monthly.py:465-759`

- [ ] **Step 1: Write failing header and layout tests**

Use workbook fixtures with `AA=发货前退货率`, `AB=发货后退货率`, and `AC=退货率`. Assert:

```python
self.assertEqual(return_header(cycle, "发货前退货率"), "发货前退货率（8.1-8.31）")
self.assertEqual(return_header(cycle, "发货后退货率"), "发货后退货率（8.1-8.31）")
self.assertEqual(return_header(cycle), "退货率（8.1-8.31）")

layout = discover_layout(sheet, shared, cycle)
self.assertEqual(layout.before_return_col, "AA")
self.assertEqual(layout.after_return_col, "AB")
self.assertEqual(layout.return_col, "AC")
```

Also assert duplicate/missing headers fail and the three columns must be consecutive after `变化情况` in the documented order.

- [ ] **Step 2: Run focused tests and verify failure**

Run: `python -m unittest tests.test_monthly_schedule tests.test_xlsx_monthly.LayoutDiscoveryTests tests.test_xlsx_monthly.MonthInsertionTests -v`

Expected: FAIL because only one return column is represented.

- [ ] **Step 3: Implement header generation and layout fields**

Change the header function and dataclass:

```python
def return_header(cycle: SyncCycle, prefix: str = "退货率") -> str:
    start = cycle.return_window.start
    end = cycle.return_window.end
    return f"{prefix}（{start.month}.{start.day}-{end.month}.{end.day}）"

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
```

In `discover_layout`, identify the exact normalized prefixes `发货前退货率`, `发货后退货率`, and `退货率`, require unique matches on the core header row, and enforce:

```python
expected_tail = (
    change_number + 1,
    change_number + 2,
    change_number + 3,
)
actual_tail = (
    column_number(before_return_col),
    column_number(after_return_col),
    column_number(return_col),
)
if actual_tail != expected_tail:
    raise ValueError("expected 发货前退货率 < 发货后退货率 < 退货率 after 变化情况")
```

Update insertion redistribution and `validate_monthly_sheet` to write and validate all three dated headers.

- [ ] **Step 4: Run focused tests and verify pass**

Run: `python -m unittest tests.test_monthly_schedule tests.test_xlsx_monthly.LayoutDiscoveryTests tests.test_xlsx_monthly.MonthInsertionTests tests.test_xlsx_monthly.WorkbookValidationTests -v`

Expected: PASS with the three columns preserved after monthly insertion.

- [ ] **Step 5: Commit layout support**

```bash
git add monthly_schedule.py xlsx_monthly.py tests/test_monthly_schedule.py tests/test_xlsx_monthly.py
git commit -m "feat: support three dated return rate columns"
```

### Task 3: Write, clear, and validate independent profile values

**Files:**
- Modify: `xlsx_monthly.py:465-565,700-902,1012-1037`
- Modify: `erp_excel_sync.py:991-1055`
- Test: `tests/test_xlsx_monthly.py:760-1010`

- [ ] **Step 1: Write failing independent-write tests**

Create before and overall rows with deliberately different `itemCount` and money. Assert that row 2 writes the before value to `AA`, clears disabled after value in `AB`, and writes overall to `AC`; row 3 with `itemCount=49` clears its old value. Add a critical-SKU case where a matched row below threshold still passes matching validation.

```python
self.assertIn(b'<c r="AA2" s="219"><v>0.25</v></c>', result.sheet_xml)
self.assertIn(b'<c r="AB2" s="219"/>', result.sheet_xml)
self.assertIn(b'<c r="AC2" s="219"><v>0.5</v></c>', result.sheet_xml)
self.assertNotIn(b'<c r="AA3" s="220"><v>', result.sheet_xml)
```

- [ ] **Step 2: Run focused tests and verify failure**

Run: `python -m unittest tests.test_xlsx_monthly.MonthlyValueWriteTests -v`

Expected: FAIL because `apply_monthly_values` accepts one return dataset and never clears disabled/under-threshold cells.

- [ ] **Step 3: Implement clearing and independent matching**

Add a style-preserving cell clearer:

```python
def _clear_cell_value(row: bytes, col: str, row_number: int) -> bytes:
    current = next(
        (cell for cell in CELL_RE.finditer(row) if _cell_col(cell) == col),
        None,
    )
    style = b""
    if current is not None:
        match = re.search(rb'\bs="([^"]+)"', current.group(0).split(b">", 1)[0])
        if match:
            style = b' s="' + match.group(1) + b'"'
    cell = b'<c r="' + col.encode() + str(row_number).encode() + b'"' + style + b'/>'
    return _insert_or_replace_cell(row, col, row_number, cell)
```

Change `apply_monthly_values` and `prepare_monthly_update` to accept `before_return_rows` and `overall_return_rows`. Match each dataset separately, read `calculatedBeforeShipmentReturnRate` and `calculatedOverallReturnRate`, clear values whose calculated field is `None`, and always clear `after_return_col`. Report independent status, sales count, amounts, old/new value, and reason for both enabled profiles.

Critical-SKU pass logic must require an automatic match in actual, before, and overall datasets; a matched row whose calculated rate is `None` because `itemCount < 50` remains a valid profile match.

- [ ] **Step 4: Run focused tests and verify pass**

Run: `python -m unittest tests.test_xlsx_monthly.MonthlyValueWriteTests -v`

Expected: PASS for independent values, stale-value clearing, disabled after column, 49/50 boundary, and critical-SKU behavior.

- [ ] **Step 5: Commit independent workbook writing**

```bash
git add erp_excel_sync.py xlsx_monthly.py tests/test_xlsx_monthly.py
git commit -m "feat: write independent return rates to workbook"
```

### Task 4: Orchestrate live/offline queries, snapshots, and reports

**Files:**
- Modify: `erp_excel_sync.py:1057-1234,1235-1333`
- Test: `tests/test_erp_excel_sync.py:57-222,313-384`

- [ ] **Step 1: Write failing orchestration tests**

Assert live execution calls ERP in this order: actual with empty `asTypes`, before with `("5",)`, overall with `("5","1","2","7","8","10")`. Assert offline mode requires `--actual-json`, `--before-return-json`, and `--return-json` together. Assert snapshot names are `return_before_<window>.json` and `return_overall_<window>.json` and reports include both profiles.

- [ ] **Step 2: Run focused tests and verify failure**

Run: `python -m unittest tests.test_erp_excel_sync.MonthlyOrchestrationTests tests.test_erp_excel_sync.RuntimeConfigTests -v`

Expected: FAIL because orchestration currently loads one return response.

- [ ] **Step 3: Implement three-document orchestration**

Add `--before-return-json` while retaining `--return-json` as the overall input. Require all three offline files together. For live cookie and Chrome paths, fetch:

```python
actual_document = fetch(..., actual_start, actual_end, as_types=())
before_document = fetch(
    ..., return_start, return_end, as_types=BEFORE_RETURN_PROFILE.as_types
)
overall_document = fetch(
    ..., return_start, return_end, as_types=OVERALL_RETURN_PROFILE.as_types
)
```

Save and compare profile-specific snapshots. Calculate each enabled profile into its own output field before calling `prepare_monthly_update`. Write `before_return_changes.csv`, `overall_return_changes.csv`, `manual_review.csv`, `critical_skus.csv`, and `summary.json`. Summary JSON must include each profile's status codes, API count, calculated count, below-threshold count, invalid-money count, and disabled status for after shipment.

- [ ] **Step 4: Run focused tests and verify pass**

Run: `python -m unittest tests.test_erp_excel_sync.MonthlyOrchestrationTests tests.test_erp_excel_sync.RuntimeConfigTests -v`

Expected: PASS with exactly three live requests and independent snapshots/reports.

- [ ] **Step 5: Commit orchestration and reporting**

```bash
git add erp_excel_sync.py tests/test_erp_excel_sync.py
git commit -m "feat: orchestrate independent return rate queries"
```

### Task 5: Documentation and complete verification

**Files:**
- Modify: `README.md:1-48`
- Test: `tests/test_erp_excel_sync.py`
- Test: `tests/test_monthly_schedule.py`
- Test: `tests/test_xlsx_monthly.py`
- Test: `tests/test_chrome_erp_session.py`

- [ ] **Step 1: Update user documentation**

Document the two enabled filters, exact `itemCount >= 50` rule, formula `rawRefundMoney / saleMoney`, disabled after-shipment column, three dated headers, profile-specific snapshots/reports, and the three-file offline command:

```powershell
python .\erp_excel_sync.py --config .\config.json `
  --actual-json .\actual.json `
  --before-return-json .\return-before.json `
  --return-json .\return-overall.json `
  --dry-run
```

- [ ] **Step 2: Run the full automated suite**

Run: `python -m unittest discover -s tests -v`

Expected: all supported-platform tests PASS; Windows-only launcher tests may SKIP on macOS.

- [ ] **Step 3: Run a real-workbook dry-run smoke test**

Use the configured workbook and saved independent snapshots after they have been captured:

```bash
python erp_excel_sync.py --config config.json \
  --actual-json snapshots/actual_2026-08-01_2026-08-14.json \
  --before-return-json snapshots/return_before_2026-07-01_2026-07-31.json \
  --return-json snapshots/return_overall_2026-07-01_2026-07-31.json \
  --dry-run
```

Expected: exit code 0, no master workbook replacement, three headers validated, after-shipment values blank, and report output distinguishes calculated, below-threshold, and invalid rows.

- [ ] **Step 4: Inspect the diff for scope and placeholders**

Run: `git diff --check && rg -n "TBD|TODO|implement later|fill in details" erp_excel_sync.py monthly_schedule.py xlsx_monthly.py README.md tests`

Expected: `git diff --check` exits 0 and the search finds no new placeholder text.

- [ ] **Step 5: Commit documentation and verification updates**

```bash
git add README.md tests erp_excel_sync.py monthly_schedule.py xlsx_monthly.py
git commit -m "docs: explain three return rate sync rules"
```
