"""Update helpers for future modularization.

The current UI still drives updates from app.main. This module exists to
hold update-related utilities as we continue to split responsibilities.
"""

from __future__ import annotations

import datetime


def parse_release_ts(value: str) -> float:
    if not value:
        return 0.0
    for fmt in ("%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.datetime.strptime(value, fmt).timestamp()
        except Exception:
            continue
    return 0.0
