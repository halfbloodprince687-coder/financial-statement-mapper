"""Validation helpers.

Two kinds of checks:
1. Recompute every ``=SUM(...)`` formula that already lives in the
   destination template, using the values the user has just mapped in, and
   compare it against the ORIGINAL total the source file reported for that
   same line (when the user also mapped the total row itself). Large
   mismatches usually mean a component was mapped to the wrong source row,
   or the multiplication factor was applied inconsistently.
2. A basic accounting identity check: Total Equity & Liabilities must equal
   Total Assets, for both the current and previous period.
"""

from __future__ import annotations

import re

SUM_RANGE_RE = re.compile(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?")


def _col_letters_to_idx0(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def formula_rows_same_col(formula: str, col0: int) -> list[int]:
    """Extract the 0-based row numbers a SUM formula references, restricted
    to the given 0-based column (most template formulas are single-column,
    e.g. =SUM(C6:C8,C9,C11:C14))."""
    rows: list[int] = []
    for m in SUM_RANGE_RE.finditer(formula):
        c1, r1, c2, r2 = m.groups()
        c1i = _col_letters_to_idx0(c1)
        if c1i != col0:
            continue
        r1i = int(r1) - 1
        if c2 and r2:
            r2i = int(r2) - 1
            rows.extend(range(min(r1i, r2i), max(r1i, r2i) + 1))
        else:
            rows.append(r1i)
    return rows


def recompute_total(formula: str, col0: int, values_by_row0: dict[int, float]) -> float:
    rows = formula_rows_same_col(formula, col0)
    return sum(values_by_row0.get(r, 0) or 0 for r in rows)


def check_balance_sheet_identity(eql_total: float | None, ca_total: float | None,
                                  tolerance: float = 1.0) -> dict:
    if eql_total is None or ca_total is None:
        return {"checked": False}
    diff = round((eql_total or 0) - (ca_total or 0), 2)
    return {
        "checked": True,
        "equity_and_liabilities": eql_total,
        "total_assets": ca_total,
        "difference": diff,
        "balanced": abs(diff) <= tolerance,
    }
