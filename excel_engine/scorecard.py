"""Before/after data-quality scorecard.

Definitions (all computed over the configured columns only):
  completeness      = non-missing cells / all cells
  native-typed      = number/date cells already stored as real numbers/dates / non-missing number+date cells
  rule pass rate    = rows with no 'error' violation / rows (strict: raw text in a number column fails)
  uniqueness        = 1 - exact-duplicate rows / rows (key = duplicates.key)
  quality score     = mean of the four percentages above
"""

from __future__ import annotations

from datetime import date, datetime

import numpy as np
import pandas as pd

from .config import Config
from .dedupe import norm_value
from .textnorm import is_missing
from .validate import validate


def _native(v, typ: str) -> bool:
    if typ in ("number", "integer"):
        return isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool)
    return isinstance(v, (datetime, date, pd.Timestamp))


def quality_metrics(df: pd.DataFrame, cfg: Config, dup_df: pd.DataFrame | None = None) -> dict:
    """``dup_df`` (optional) is the frame used to count exact duplicates, e.g. the cleaned-but-not-deduped
    data for the 'before' column: formats in raw files differ, so duplicates only become visible after cleaning."""
    cols, n = list(cfg.columns), len(df)
    missing = int(df[cols].map(is_missing).to_numpy().sum()) if n else 0
    cells = n * len(cols)
    typed = [(c, s.type) for c, s in cfg.columns.items() if s.type != "text"]
    nn = native = 0
    for c, typ in typed:
        for v in df[c]:
            if not is_missing(v):
                nn += 1
                native += _native(v, typ)
    dups, dd = 0, (df if dup_df is None else dup_df)
    if cfg.duplicates.key and len(dd):
        keys = ["\x1f".join(norm_value(v) for v in r) for r in dd[cfg.duplicates.key].itertuples(index=False)]
        keys = [k for k in keys if k.strip("\x1f")]
        dups = len(keys) - len(set(keys))
    errs = validate(df, cfg) if cfg.validation else pd.DataFrame(columns=["severity", "clean_row"])
    e = errs[errs["severity"] == "error"]
    completeness = 100 * (1 - missing / cells) if cells else 100.0
    native_pct = 100 * native / nn if nn else 100.0
    rule_pass = 100 * (1 - e["clean_row"].nunique() / n) if n else 100.0
    unique_pct = 100 * (1 - dups / len(dd)) if len(dd) else 100.0
    return {"rows": n, "missing_cells": missing, "completeness_pct": completeness,
            "native_typed_pct": native_pct, "exact_duplicate_rows": dups,
            "rule_violations": len(e), "rule_warnings": int((errs["severity"] == "warning").sum()),
            "rows_with_errors": int(e["clean_row"].nunique()), "rule_pass_pct": rule_pass,
            "uniqueness_pct": unique_pct,
            "quality_score": (completeness + native_pct + rule_pass + unique_pct) / 4}


# (key, label, higher_is_better or None for neutral, is_percent)
_ROWS = [("rows", "Rows", None, False),
         ("missing_cells", "Missing cells", False, False),
         ("completeness_pct", "Completeness (%)", True, True),
         ("native_typed_pct", "Cells stored as native numbers/dates (%)", True, True),
         ("exact_duplicate_rows", "Exact duplicate rows", False, False),
         ("uniqueness_pct", "Uniqueness (%)", True, True),
         ("rule_violations", "Rule violations (errors)", False, False),
         ("rows_with_errors", "Rows with at least one error", False, False),
         ("rule_pass_pct", "Rule pass rate (%)", True, True),
         ("quality_score", "Quality score (0-100)", True, True)]

PCT_KEYS = ["completeness_pct", "native_typed_pct", "uniqueness_pct", "rule_pass_pct", "quality_score"]
PCT_LABELS = {k: lab for k, lab, _, p in _ROWS if p}


def build_scorecard(before: dict, after: dict) -> pd.DataFrame:
    rows = []
    for key, label, better_high, pct in _ROWS:
        b, a = before[key], after[key]
        if pct:
            b, a = round(b, 1), round(a, 1)
        change = round(a - b, 1) if pct else a - b
        if better_high is None or change == 0:
            verdict = "n/a" if better_high is None else "unchanged"
        else:
            verdict = "improved" if (change > 0) == better_high else "worse"
        rows.append({"metric": label, "before": b, "after": a, "change": change, "verdict": verdict})
    return pd.DataFrame(rows)
