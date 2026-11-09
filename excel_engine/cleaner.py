"""Cleaning rules: text normalisation, type coercion and missing-value strategies."""

from __future__ import annotations

import re
from dataclasses import dataclass

import numpy as np
import pandas as pd

from .config import ColumnSpec, Config
from .textnorm import is_missing, match_key, normalize_text, parse_date, parse_integer, parse_number


@dataclass
class CleanResult:
    df: pd.DataFrame
    log: pd.DataFrame        # per column: normalised / unparseable / filled counts
    failures: pd.DataFrame   # cells that could not be coerced (set to missing)
    dropped: pd.DataFrame    # rows removed by a 'drop_row' strategy


def _as_text(v) -> str:
    if isinstance(v, (float, np.floating)) and float(v).is_integer():
        return str(int(v))
    return str(v)


def _coerce(v, spec: ColumnSpec, vmap: dict[str, str], dayfirst: bool):
    """Return (clean_value_or_None, was_changed)."""
    if spec.type == "text":
        s = normalize_text(_as_text(v), case=spec.case, arabic_letters=spec.arabic_letters)
        s = vmap.get(match_key(s), s)
        return (s or None), s != str(v)
    if spec.type == "date":
        new = parse_date(v, dayfirst)
    elif spec.type == "integer":
        new = parse_integer(v)
    else:
        new = parse_number(v)
    return new, isinstance(v, str)


def _build(vals: list, typ: str) -> pd.Series:
    if typ == "number":
        return pd.Series(vals, dtype="float64")
    if typ == "integer":
        return pd.Series(vals, dtype="Int64")
    if typ == "date":
        return pd.to_datetime(pd.Series(vals, dtype=object))
    return pd.Series(vals, dtype=object)


def _fill(s: pd.Series, spec: ColumnSpec) -> pd.Series:
    how = spec.missing
    if how in ("keep", "drop_row") or not s.isna().any():
        return s
    if how == "zero":
        return s.fillna(0)
    if how in ("mean", "median"):
        if s.notna().sum() == 0:
            return s
        m = float(s.mean() if how == "mean" else s.median())
        return s.fillna(int(round(m)) if spec.type == "integer" else m)
    if how == "mode":
        mode = s.dropna().mode()
        return s.fillna(mode.iloc[0]) if len(mode) else s
    if how == "ffill":
        return s.ffill()
    const = how.split(":", 1)[1]  # constant:<value>
    if spec.type == "number":
        return s.fillna(float(const))
    if spec.type == "integer":
        return s.fillna(int(float(const)))
    if spec.type == "date":
        ts = parse_date(const)
        return s.fillna(ts) if ts is not None else s
    return s.fillna(const)


def clean(df: pd.DataFrame, cfg: Config) -> CleanResult:
    out = df.copy()
    log, fails, drop_mask = [], [], pd.Series(False, index=df.index)
    for name, spec in cfg.columns.items():
        vmap = spec.vmap()
        vals, changed, failed = [], 0, 0
        for i, v in df[name].items():
            if is_missing(v):
                vals.append(None)
                continue
            new, was_changed = _coerce(v, spec, vmap, cfg.dayfirst)
            if new is None:
                failed += 1
                fails.append({"column": name, "raw_value": str(v)[:80],
                              "source_file": df.at[i, "source_file"], "source_sheet": df.at[i, "source_sheet"],
                              "source_row": df.at[i, "source_row"]})
            changed += bool(was_changed) and new is not None
            vals.append(new)
        s = _build(vals, spec.type)
        s.index = df.index
        before_missing = int(s.isna().sum())
        s = _fill(s, spec)
        filled = before_missing - int(s.isna().sum())
        if spec.missing == "drop_row":
            drop_mask |= s.isna()
        out[name] = s
        log.append({"column": name, "type": spec.type, "missing_strategy": spec.missing,
                    "cells_normalised": changed, "unparseable_to_missing": failed, "missing_filled": filled})
    for cname, expr in cfg.computed.items():
        a, b = re.fullmatch(r"\s*(\w+)\s*\*\s*(\w+)\s*", expr).groups()
        out[cname] = out[a].astype("float64") * out[b].astype("float64")
    dropped = out[drop_mask].assign(drop_reason="missing value in a 'drop_row' column")
    out = out[~drop_mask].reset_index(drop=True)
    cols = ["column", "raw_value", "source_file", "source_sheet", "source_row"]
    return CleanResult(out, pd.DataFrame(log), pd.DataFrame(fails, columns=cols), dropped.reset_index(drop=True))
