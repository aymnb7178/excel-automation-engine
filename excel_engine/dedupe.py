"""Duplicate detection: exact (normalised key) and fuzzy (difflib, token-aware) with a review table."""

from __future__ import annotations

import itertools
import math
from dataclasses import dataclass, field
from difflib import SequenceMatcher

import numpy as np
import pandas as pd

from .config import Config
from .textnorm import match_key

REVIEW_META = ["group_id", "match_type", "action", "similarity", "source_file", "source_sheet", "source_row"]


@dataclass
class DedupeResult:
    df: pd.DataFrame
    review: pd.DataFrame
    exact_removed: int = 0
    fuzzy_pairs: int = 0
    warnings: list[str] = field(default_factory=list)


def norm_value(v) -> str:
    """Comparison form of one cell: '' for missing, ISO date, rounded number, or match_key text."""
    if v is None or v is pd.NA or v is pd.NaT or (isinstance(v, float) and math.isnan(v)):
        return ""
    if isinstance(v, pd.Timestamp):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
        return repr(round(float(v), 6))
    return match_key(v)


def similarity(a: str, b: str) -> float:
    """max(plain ratio, token-sorted ratio): tolerant of typos and of word order."""
    if not a or not b:
        return 0.0
    plain = SequenceMatcher(None, a, b).ratio()
    ta, tb = " ".join(sorted(a.split())), " ".join(sorted(b.split()))
    return max(plain, SequenceMatcher(None, ta, tb).ratio())


def _review_rows(df: pd.DataFrame, idxs, cols, gid, mtype, actions, sim) -> list[dict]:
    rows = []
    for i, act in zip(idxs, actions):
        r = {"group_id": gid, "match_type": mtype, "action": act, "similarity": round(sim, 3),
             "source_file": df.at[i, "source_file"], "source_sheet": df.at[i, "source_sheet"],
             "source_row": df.at[i, "source_row"]}
        r.update({c: df.at[i, c] for c in cols})
        rows.append(r)
    return rows


def deduplicate(df: pd.DataFrame, cfg: Config) -> DedupeResult:
    d = cfg.duplicates
    cols = list(cfg.columns)
    df = df.reset_index(drop=True)
    review: list[dict] = []
    warnings: list[str] = []
    drop: set[int] = set()
    if d.key:
        keys = pd.Series(["\x1f".join(norm_value(v) for v in row) for row in df[d.key].itertuples(index=False)])
        empty = "\x1f".join([""] * len(d.key))
        n_group = 0
        for _, grp in df.groupby(keys[keys != empty], sort=False):
            if len(grp) < 2:
                continue
            n_group += 1
            idxs = list(grp.index)
            if d.keep == "most_complete":
                score = {i: int(df.loc[i, cols].notna().sum()) for i in idxs}
                keeper = max(idxs, key=lambda i: (score[i], -i))
            else:
                keeper = idxs[0]
            drop.update(i for i in idxs if i != keeper)
            review += _review_rows(df, idxs, cols, f"E{n_group:04d}", "exact",
                                   ["kept" if i == keeper else "removed" for i in idxs], 1.0)
    out = df.drop(index=sorted(drop)).reset_index(drop=True)
    pairs = 0
    if d.fuzzy_columns:
        blocks: dict[str, list[int]] = {}
        if d.fuzzy_block_by:
            for i, row in enumerate(out[d.fuzzy_block_by].itertuples(index=False)):
                blocks.setdefault("\x1f".join(norm_value(v) for v in row), []).append(i)
        elif len(out) <= d.max_unblocked_rows:
            blocks["*"] = list(range(len(out)))
        else:
            warnings.append(f"fuzzy matching skipped: {len(out)} rows exceed max_unblocked_rows "
                            f"({d.max_unblocked_rows}) and no fuzzy_block_by is set")
        texts = {c: [norm_value(v) for v in out[c]] for c in d.fuzzy_columns}
        for members in blocks.values():
            for a, b in itertools.combinations(members, 2):
                sims = [similarity(texts[c][a], texts[c][b]) for c in d.fuzzy_columns]
                if min(sims) >= d.fuzzy_threshold:
                    pairs += 1
                    review += _review_rows(out, [a, b], cols, f"F{pairs:04d}", "fuzzy",
                                           ["review", "review"], min(sims))
    cols_out = REVIEW_META + cols
    return DedupeResult(out, pd.DataFrame(review, columns=cols_out), len(drop), pairs, warnings)
