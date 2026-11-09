"""Rule-based validation producing an error report with cell coordinates."""

from __future__ import annotations

import re
from datetime import date, datetime

import numpy as np
import pandas as pd
from openpyxl.utils import get_column_letter

from .config import Config
from .dedupe import norm_value
from .textnorm import is_missing, match_key, parse_date, parse_number

ERROR_COLUMNS = ["severity", "rule", "column", "value", "message", "source_file", "source_sheet",
                 "source_row", "source_cell", "clean_row", "clean_cell"]


def _text(v) -> str:
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(v).strftime("%Y-%m-%d")
    if isinstance(v, (float, np.floating)) and float(v).is_integer():
        return str(int(v))
    return str(v)


def _numeric(v):
    return float(v) if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool) else None


def _date(v):
    return pd.Timestamp(v) if isinstance(v, (pd.Timestamp, datetime, date)) else None


def validate(df: pd.DataFrame, cfg: Config, mapping: pd.DataFrame | None = None) -> pd.DataFrame:
    """Apply cfg.validation to df. Works on raw (as-delivered) and cleaned frames.

    ``clean_cell`` is the A1 address in the 'Clean Data' sheet; ``source_cell`` is the address in the
    original file (needs the loader's header mapping and the source_* columns).
    """
    src_col: dict[tuple, int] = {}
    if mapping is not None and len(mapping):
        for m in mapping[mapping["canonical"] != ""].itertuples():
            src_col.setdefault((m.file, m.sheet, m.canonical), int(m.source_col))
    positions = {c: j + 1 for j, c in enumerate(df.columns)}
    has_src = {"source_file", "source_sheet", "source_row"} <= set(df.columns)
    out: list[dict] = []

    def add(rule, i, v, msg):
        row = {"severity": rule.severity, "rule": rule.rule, "column": rule.column,
               "value": "" if is_missing(v) else _text(v)[:80], "message": rule.message or msg,
               "source_file": None, "source_sheet": None, "source_row": None, "source_cell": None,
               "clean_row": i + 2, "clean_cell": f"{get_column_letter(positions[rule.column])}{i + 2}"}
        if has_src:
            f, sh, r = df.at[i, "source_file"], df.at[i, "source_sheet"], df.at[i, "source_row"]
            row.update(source_file=f, source_sheet=sh, source_row=int(r))
            col = src_col.get((f, sh, rule.column))
            if col:
                row["source_cell"] = f"{get_column_letter(col)}{int(r)}"
        out.append(row)

    for rule in cfg.validation:
        typ = cfg.columns[rule.column].type if rule.column in cfg.columns else "number"
        s = df[rule.column]
        if rule.rule == "unique":
            seen: dict[str, int] = {}
            for i, v in s.items():
                if is_missing(v):
                    continue
                k = norm_value(v)
                if k in seen:
                    add(rule, i, v, f"duplicate value; first seen in row {seen[k]} of Clean Data")
                else:
                    seen[k] = i + 2
            continue
        allowed = {match_key(x) for x in rule.values or []}
        rx = re.compile(rule.pattern) if rule.pattern else None
        for i, v in s.items():
            miss = is_missing(v)
            if rule.rule == "required":
                if miss:
                    add(rule, i, v, "value is required")
            elif miss:
                continue
            elif rule.rule == "regex":
                if not rx.fullmatch(_text(v)):
                    add(rule, i, v, f"does not match /{rule.pattern}/")
            elif rule.rule == "allowed":
                if match_key(_text(v)) not in allowed:
                    add(rule, i, v, f"not one of: {', '.join(rule.values)}")
            elif rule.rule == "range":
                if typ == "date":
                    x, lo, hi = _date(v), parse_date(rule.min), parse_date(rule.max)
                else:
                    x = _numeric(v) if typ in ("number", "integer") else parse_number(v)
                    lo = None if rule.min is None else float(rule.min)
                    hi = None if rule.max is None else float(rule.max)
                if x is None:
                    add(rule, i, v, "not a valid " + ("date" if typ == "date" else "number"))
                elif (lo is not None and x < lo) or (hi is not None and x > hi):
                    add(rule, i, v, f"outside range [{rule.min if rule.min is not None else '-inf'}, "
                                    f"{rule.max if rule.max is not None else 'inf'}]")
    res = pd.DataFrame(out, columns=ERROR_COLUMNS)
    if len(res):
        res["_sev"] = res["severity"].map({"error": 0, "warning": 1})
        res = res.sort_values(["_sev", "source_file", "source_row", "column"], kind="stable").drop(columns="_sev")
    return res.reset_index(drop=True)
