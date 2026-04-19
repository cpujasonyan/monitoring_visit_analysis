# -*- coding: utf-8 -*-
"""Monitoring Visit Analysis - Date & Working Day Utilities."""

from datetime import datetime, date, timedelta
from chinese_calendar import is_workday


def parse_dates(val):
    """Parse cell value into a list of date objects."""
    if val is None:
        return []
    if isinstance(val, datetime):
        return [val.date()]
    if isinstance(val, date):
        return [val]
    s = str(val).strip()
    if not s:
        return []
    results = []
    for p in s.split(','):
        p = p.strip()
        try:
            results.append(datetime.strptime(p, '%Y-%m-%d').date())
        except (ValueError, TypeError):
            pass
    return results


def last_date(dl):
    """Return the latest date from a list, or None."""
    return max(dl) if dl else None


def first_date(dl):
    """Return the earliest date from a list, or None."""
    return min(dl) if dl else None


def working_days_between(d1, d2):
    """Count working days between d1 (exclusive) and d2 (inclusive).
    Uses chinese_calendar for Chinese holidays/adjusted workdays.
    Returns 0 if d2 <= d1."""
    if not d1 or not d2:
        return None
    if d2 <= d1:
        return 0
    count, cur = 0, d1 + timedelta(days=1)
    while cur <= d2:
        try:
            if is_workday(cur):
                count += 1
        except NotImplementedError:
            if cur.weekday() < 5:
                count += 1
        cur += timedelta(days=1)
    return count


def add_working_days(d, n):
    """Add n working days to date d. Returns the resulting date."""
    if not d:
        return None
    count, cur = 0, d
    while count < n:
        cur += timedelta(days=1)
        try:
            if is_workday(cur):
                count += 1
        except NotImplementedError:
            if cur.weekday() < 5:
                count += 1
    return cur
