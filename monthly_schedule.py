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
    return month_start(value.replace(day=28) + timedelta(days=4)) - timedelta(days=1)


def resolve_sync_cycle(run_date: date) -> SyncCycle:
    current_month = month_start(run_date)
    if run_date.day < 15:
        actual_month = previous_month_start(current_month)
        previous_month = previous_month_start(actual_month)
        return SyncCycle(
            run_date=run_date,
            node_date=current_month,
            kind="first",
            actual_month=actual_month,
            previous_month=previous_month,
            actual_window=DateWindow(actual_month, month_end(actual_month)),
            return_window=DateWindow(previous_month.replace(day=15), actual_month.replace(day=15)),
        )

    previous_month = previous_month_start(current_month)
    return SyncCycle(
        run_date=run_date,
        node_date=current_month.replace(day=15),
        kind="fifteenth",
        actual_month=current_month,
        previous_month=previous_month,
        actual_window=DateWindow(current_month, current_month.replace(day=14)),
        return_window=DateWindow(previous_month, month_end(previous_month)),
    )


def actual_header(cycle: SyncCycle) -> str:
    month = cycle.actual_month.month
    if cycle.kind == "fifteenth":
        return f"{month}月实发（{month}.15）"
    return f"{month}月实发"


def return_header(cycle: SyncCycle, prefix: str = "退货率") -> str:
    start = cycle.return_window.start
    end = cycle.return_window.end
    return f"{prefix}（{start.month}.{start.day}-{end.month}.{end.day}）"


def peer_formula(previous_col: str, row: int, cycle: SyncCycle) -> str:
    days = month_end(cycle.previous_month).day
    return f"{previous_col}{row}/{days}*14"


def change_formula(actual_col: str, peer_col: str, row: int, cycle: SyncCycle) -> str:
    month = cycle.actual_month.month
    return (
        f'TEXT({actual_col}{row}-{peer_col}{row},'
        f'"{month}月增加0件；{month}月减少0件；持平")'
    )
