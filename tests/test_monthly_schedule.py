import unittest
from datetime import date

from monthly_schedule import (
    DateWindow,
    actual_header,
    change_formula,
    month_end,
    month_start,
    peer_formula,
    previous_month_start,
    resolve_sync_cycle,
    return_header,
)


class CalendarHelperTests(unittest.TestCase):
    def test_date_window_returns_iso_dates(self):
        window = DateWindow(date(2026, 8, 1), date(2026, 8, 31))

        self.assertEqual(window.iso(), ("2026-08-01", "2026-08-31"))

    def test_month_helpers_cross_the_year_boundary(self):
        january = date(2026, 1, 20)

        self.assertEqual(month_start(january), date(2026, 1, 1))
        self.assertEqual(previous_month_start(january), date(2025, 12, 1))
        self.assertEqual(month_end(previous_month_start(january)), date(2025, 12, 31))


class SyncCycleTests(unittest.TestCase):
    def test_day_fourteen_remains_in_the_first_cycle(self):
        cycle = resolve_sync_cycle(date(2026, 9, 14))

        self.assertEqual(cycle.node_date, date(2026, 9, 1))
        self.assertEqual(cycle.kind, "first")

    def test_september_first_node_uses_august_actuals_and_midmonth_returns(self):
        cycle = resolve_sync_cycle(date(2026, 9, 8))

        self.assertEqual(cycle.run_date, date(2026, 9, 8))
        self.assertEqual(cycle.node_date, date(2026, 9, 1))
        self.assertEqual(cycle.kind, "first")
        self.assertEqual(cycle.actual_month, date(2026, 8, 1))
        self.assertEqual(cycle.previous_month, date(2026, 7, 1))
        self.assertEqual(cycle.actual_window, DateWindow(date(2026, 8, 1), date(2026, 8, 31)))
        self.assertEqual(cycle.return_window, DateWindow(date(2026, 7, 15), date(2026, 8, 15)))

    def test_september_fifteenth_node_uses_first_fourteen_days_and_august_returns(self):
        cycle = resolve_sync_cycle(date(2026, 9, 30))

        self.assertEqual(cycle.node_date, date(2026, 9, 15))
        self.assertEqual(cycle.kind, "fifteenth")
        self.assertEqual(cycle.actual_month, date(2026, 9, 1))
        self.assertEqual(cycle.previous_month, date(2026, 8, 1))
        self.assertEqual(cycle.actual_window, DateWindow(date(2026, 9, 1), date(2026, 9, 14)))
        self.assertEqual(cycle.return_window, DateWindow(date(2026, 8, 1), date(2026, 8, 31)))

    def test_january_first_node_crosses_the_year_boundary(self):
        cycle = resolve_sync_cycle(date(2026, 1, 1))

        self.assertEqual(cycle.actual_month, date(2025, 12, 1))
        self.assertEqual(cycle.previous_month, date(2025, 11, 1))
        self.assertEqual(cycle.actual_window, DateWindow(date(2025, 12, 1), date(2025, 12, 31)))
        self.assertEqual(cycle.return_window, DateWindow(date(2025, 11, 15), date(2025, 12, 15)))


class HeaderAndFormulaTests(unittest.TestCase):
    def test_first_node_headers_describe_the_resolved_windows(self):
        cycle = resolve_sync_cycle(date(2026, 9, 8))

        self.assertEqual(actual_header(cycle), "8月实发")
        self.assertEqual(
            return_header(cycle, "发货前退货率"),
            "发货前退货率（7.15-8.15）",
        )
        self.assertEqual(
            return_header(cycle, "发货后退货率"),
            "发货后退货率（7.15-8.15）",
        )
        self.assertEqual(return_header(cycle), "退货率（7.15-8.15）")

    def test_fifteenth_node_headers_describe_the_resolved_windows(self):
        cycle = resolve_sync_cycle(date(2026, 9, 15))

        self.assertEqual(actual_header(cycle), "9月实发（9.15）")
        self.assertEqual(
            return_header(cycle, "发货前退货率"),
            "发货前退货率（8.1-8.31）",
        )
        self.assertEqual(
            return_header(cycle, "发货后退货率"),
            "发货后退货率（8.1-8.31）",
        )
        self.assertEqual(return_header(cycle), "退货率（8.1-8.31）")

    def test_peer_formula_uses_leap_february_day_count(self):
        cycle = resolve_sync_cycle(date(2024, 3, 15))

        self.assertEqual(peer_formula("X", 9, cycle), "ROUND(X9/29*14,0)")

    def test_fifteenth_peer_formula_tracks_28_30_and_31_day_previous_months(self):
        cases = (
            (date(2026, 3, 15), 28),
            (date(2026, 5, 15), 30),
            (date(2026, 6, 15), 31),
        )
        for run_date, days in cases:
            with self.subTest(run_date=run_date):
                cycle = resolve_sync_cycle(run_date)
                self.assertEqual(
                    peer_formula("X", 9, cycle),
                    f"ROUND(X9/{days}*14,0)",
                )

    def test_first_node_peer_formula_keeps_dynamic_month_day_count_unrounded(self):
        cycle = resolve_sync_cycle(date(2026, 9, 1))

        self.assertEqual(peer_formula("W", 9, cycle), "W9/31*14")

    def test_october_change_formula_keeps_month_outside_number_format(self):
        cycle = resolve_sync_cycle(date(2026, 10, 15))

        self.assertEqual(
            change_formula("AA", "Z", 9, cycle),
            'IF(AA9-Z9=0,"持平","10月"&IF(AA9-Z9>0,"增加","减少")'
            '&TEXT(ABS(AA9-Z9),"0")&"件")',
        )

    def test_first_node_change_formula_can_compare_actual_to_previous_actual(self):
        cycle = resolve_sync_cycle(date(2026, 9, 1))

        self.assertEqual(
            change_formula("Y", "W", 4, cycle),
            'IF(Y4-W4=0,"持平","8月"&IF(Y4-W4>0,"增加","减少")'
            '&TEXT(ABS(Y4-W4),"0")&"件")',
        )

    def test_change_formula_keeps_every_month_literal_in_both_cycles(self):
        for month in range(1, 13):
            for day in (1, 15):
                with self.subTest(month=month, day=day):
                    cycle = resolve_sync_cycle(date(2026, month, day))
                    formula = change_formula("AB", "AA", 123, cycle)
                    self.assertIn(f'"{cycle.actual_month.month}月"&', formula)
                    self.assertIn('TEXT(ABS(AB123-AA123),"0")', formula)
                    self.assertIn('IF(AB123-AA123=0,"持平",', formula)
                    self.assertIn('IF(AB123-AA123>0,"增加","减少")', formula)


if __name__ == "__main__":
    unittest.main()
