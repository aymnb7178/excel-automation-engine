"""JSON configuration models (pydantic). Unknown keys are rejected so typos fail loudly."""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, ValidationError, model_validator

from .textnorm import header_key, match_key

META_COLUMNS = ["source_file", "source_sheet", "source_row"]
_STRATEGY = re.compile(r"^(keep|drop_row|zero|mean|median|mode|ffill|constant:.*)$")


class ConfigError(ValueError):
    """Raised for invalid configuration or pipeline files."""


class _Strict(BaseModel):
    model_config = ConfigDict(extra="forbid")


class ColumnSpec(_Strict):
    aliases: list[str] = []
    type: Literal["text", "number", "integer", "date"] = "text"
    case: Literal["title", "upper", "lower"] | None = None
    arabic_letters: bool = False
    value_map: dict[str, str] = {}
    missing: str = "keep"

    def vmap(self) -> dict[str, str]:
        """value_map keyed by match_key so lookups ignore case, spacing and alef variants."""
        return {match_key(k): v for k, v in self.value_map.items()}


class DuplicateConfig(_Strict):
    key: list[str] = []
    keep: Literal["first", "most_complete"] = "most_complete"
    fuzzy_columns: list[str] = []
    fuzzy_block_by: list[str] = []
    fuzzy_threshold: float = Field(0.88, ge=0.0, le=1.0)
    max_unblocked_rows: int = 3000


class Rule(_Strict):
    column: str
    rule: Literal["required", "regex", "range", "allowed", "unique"]
    severity: Literal["error", "warning"] = "error"
    pattern: str | None = None
    min: float | str | None = None
    max: float | str | None = None
    values: list[str] | None = None
    message: str | None = None

    @model_validator(mode="after")
    def _check(self):
        if self.rule == "regex":
            if not self.pattern:
                raise ValueError("regex rule needs 'pattern'")
            re.compile(self.pattern)
        if self.rule == "range" and self.min is None and self.max is None:
            raise ValueError("range rule needs 'min' and/or 'max'")
        if self.rule == "allowed" and not self.values:
            raise ValueError("allowed rule needs 'values'")
        return self


class PivotSpec(_Strict):
    name: str
    index: str
    values: str | None = None
    agg: Literal["sum", "mean", "count"] = "sum"
    columns: str | None = None
    freq: Literal["M"] | None = None  # group a date index by month


class ChartSpec(_Strict):
    pivot: str
    kind: Literal["bar", "line"] = "bar"
    title: str = ""
    y_title: str = ""


class ReportConfig(_Strict):
    title: str = "Data Quality and Merge Report"
    subtitle: str = ""
    pivots: list[PivotSpec] = []
    charts: list[ChartSpec] = []


class Config(_Strict):
    columns: dict[str, ColumnSpec]
    duplicates: DuplicateConfig = DuplicateConfig()
    validation: list[Rule] = []
    report: ReportConfig = ReportConfig()
    header_scan_rows: int = Field(15, ge=1, le=200)
    fuzzy_header_cutoff: float = Field(0.85, ge=0.5, le=1.0)
    dayfirst: bool = True
    keep_unmapped: bool = False
    computed: dict[str, str] = {}  # name -> 'a * b' (product of two numeric columns)

    def alias_index(self) -> dict[str, str]:
        idx: dict[str, str] = {}
        for name, spec in self.columns.items():
            for a in [name, *spec.aliases]:
                idx[header_key(a)] = name
        return idx

    @model_validator(mode="after")
    def _check(self):
        names = set(self.columns)
        computed = set(self.computed)
        for cn, expr in self.computed.items():
            m = re.fullmatch(r"\s*(\w+)\s*\*\s*(\w+)\s*", expr)
            if not m:
                raise ValueError(f"computed '{cn}': only 'a * b' is supported")
            for op in m.groups():
                if op not in names or self.columns[op].type not in ("number", "integer"):
                    raise ValueError(f"computed '{cn}': '{op}' must be a numeric column")
            if cn in names or cn in META_COLUMNS:
                raise ValueError(f"computed name '{cn}' clashes with an existing column")
        if not names:
            raise ValueError("'columns' must define at least one column")
        for n in names & set(META_COLUMNS):
            raise ValueError(f"column name '{n}' is reserved")
        seen: dict[str, str] = {}
        for name, spec in self.columns.items():
            for a in [name, *spec.aliases]:
                k = header_key(a)
                if seen.setdefault(k, name) != name:
                    raise ValueError(f"alias '{a}' is used by both '{seen[k]}' and '{name}'")
            if not _STRATEGY.match(spec.missing):
                raise ValueError(f"column '{name}': unknown missing strategy '{spec.missing}'")
            numeric = spec.type in ("number", "integer")
            if spec.missing in ("zero", "mean", "median") and not numeric:
                raise ValueError(f"column '{name}': '{spec.missing}' needs a numeric column")
            if spec.missing.startswith("constant:"):
                val = spec.missing.split(":", 1)[1]
                if numeric and not re.fullmatch(r"-?\d+(\.\d+)?", val):
                    raise ValueError(f"column '{name}': constant '{val}' is not numeric")
        refs = [*self.duplicates.key, *self.duplicates.fuzzy_columns, *self.duplicates.fuzzy_block_by]
        refs += [r.column for r in self.validation]
        refs += [p.index for p in self.report.pivots] + [p.columns for p in self.report.pivots if p.columns]
        refs += [p.values for p in self.report.pivots if p.values]
        for r in refs:
            if r not in names | computed:
                raise ValueError(f"unknown column '{r}' referenced (defined: {sorted(names | computed)})")
        pivots = {p.name for p in self.report.pivots}
        for c in self.report.charts:
            if c.pivot not in pivots:
                raise ValueError(f"chart references unknown pivot '{c.pivot}'")
        for p in self.report.pivots:
            if p.freq and (p.index not in self.columns or self.columns[p.index].type != "date"):
                raise ValueError(f"pivot '{p.name}': freq needs a date index column")
            if p.agg != "count" and not p.values:
                raise ValueError(f"pivot '{p.name}': agg '{p.agg}' needs 'values'")
        return self


def load_config(source: str | Path | dict) -> Config:
    """Load and validate a config from a JSON file path or a dict."""
    try:
        data = source if isinstance(source, dict) else json.loads(Path(source).read_text(encoding="utf-8"))
        return Config.model_validate(data)
    except ValidationError as exc:
        msgs = "; ".join(f"{'.'.join(map(str, e['loc']))}: {e['msg']}" for e in exc.errors())
        raise ConfigError(f"invalid config: {msgs}") from exc
    except (OSError, json.JSONDecodeError) as exc:
        raise ConfigError(f"cannot read config {source}: {exc}") from exc
