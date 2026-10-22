"""Small, dependency-free period and completeness helpers."""
from __future__ import annotations

import calendar
import re
from datetime import date


def parse_period(value: str) -> str:
    """Accept YYYYQn or Qn_YYYY and return the stable YYYYQn key."""
    value = str(value).strip().upper()
    match = re.fullmatch(r"(20\d{2})Q([1-4])", value)
    if match:
        return value
    match = re.fullmatch(r"Q([1-4])[_ -]?(20\d{2})", value)
    if match:
        return f"{match[2]}Q{match[1]}"
    raise ValueError(f"Invalid period {value!r}; expected YYYYQ1..4 or Q1..4_YYYY")


def quarter_end(period: str) -> date:
    period = parse_period(period)
    year, month = int(period[:4]), int(period[-1]) * 3
    return date(year, month, calendar.monthrange(year, month)[1])


def rolling_months(period: str, count: int = 6) -> list[str]:
    end = quarter_end(period)
    serial = end.year * 12 + end.month - 1
    return [f"{(serial-i)//12:04d}-{(serial-i)%12+1:02d}" for i in range(count)]


def complete_sum(values):
    """Unknown input stays unknown; inapplicable categories are excluded upstream."""
    values = list(values)
    return None if any(value is None for value in values) else sum(values)


def ratio(numerator, denominator):
    return None if numerator is None or denominator in (None, 0) else numerator / denominator
