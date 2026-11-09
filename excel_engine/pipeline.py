"""Orchestration: merge -> clean -> de-duplicate -> validate -> scorecard, plus JSON pipeline files."""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from pydantic import BaseModel, ConfigDict, ValidationError

from .cleaner import CleanResult, clean
from .config import Config, ConfigError, load_config
from .dedupe import REVIEW_META, DedupeResult, deduplicate
from .loader import LoadResult, load_inputs
from .scorecard import build_scorecard, quality_metrics
from .validate import ERROR_COLUMNS, validate


@dataclass
class PipelineResult:
    cfg: Config
    input_path: str
    load: LoadResult
    clean: CleanResult
    dedupe: DedupeResult
    errors: pd.DataFrame
    before: dict
    after: dict
    scorecard: pd.DataFrame

    @property
    def df(self) -> pd.DataFrame:
        return self.dedupe.df


def process(input_path: str | Path, cfg: Config, *, dedupe: bool = True, validate_rules: bool = True) -> PipelineResult:
    load = load_inputs(input_path, cfg)
    cl = clean(load.df, cfg)
    if dedupe:
        dd = deduplicate(cl.df, cfg)
    else:
        dd = DedupeResult(cl.df.reset_index(drop=True), pd.DataFrame(columns=REVIEW_META + list(cfg.columns)))
    errors = validate(dd.df, cfg, load.mapping) if validate_rules and cfg.validation else pd.DataFrame(columns=ERROR_COLUMNS)
    before, after = quality_metrics(load.df, cfg, dup_df=cl.df), quality_metrics(dd.df, cfg)
    return PipelineResult(cfg, str(input_path), load, cl, dd, errors, before, after, build_scorecard(before, after))


class PipelineSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")
    input: str
    config: str
    output_dir: str = "output"
    workbook: str = "clean_report.xlsx"
    clean_csv: str | None = None
    errors_csv: str | None = None
    preview_png: str | None = None
    dedupe: bool = True
    validate_rules: bool = True


def load_pipeline(path: str | Path) -> tuple[PipelineSpec, Path]:
    p = Path(path)
    try:
        spec = PipelineSpec.model_validate(json.loads(p.read_text(encoding="utf-8")))
    except ValidationError as exc:
        raise ConfigError("invalid pipeline file: " + "; ".join(f"{'.'.join(map(str, e['loc']))}: {e['msg']}"
                                                                for e in exc.errors())) from exc
    except (OSError, json.JSONDecodeError) as exc:
        raise ConfigError(f"cannot read pipeline file {p}: {exc}") from exc
    return spec, p.parent


def run_pipeline_file(path: str | Path) -> tuple[PipelineResult, dict[str, Path]]:
    """Run a JSON pipeline file; relative paths resolve against the file's own folder."""
    from .report import build_workbook, preview_png  # local import keeps matplotlib optional for other paths

    spec, base = load_pipeline(path)
    rel = lambda s: (base / s).resolve()  # noqa: E731
    cfg = load_config(rel(spec.config))
    res = process(rel(spec.input), cfg, dedupe=spec.dedupe, validate_rules=spec.validate_rules)
    out_dir = rel(spec.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    written = {"workbook": build_workbook(res, out_dir / spec.workbook)}
    if spec.clean_csv:
        written["clean_csv"] = out_dir / spec.clean_csv
        res.df.to_csv(written["clean_csv"], index=False, encoding="utf-8-sig")
    if spec.errors_csv:
        written["errors_csv"] = out_dir / spec.errors_csv
        res.errors.to_csv(written["errors_csv"], index=False, encoding="utf-8-sig")
    if spec.preview_png:
        written["preview_png"] = preview_png(res, rel(spec.preview_png))
    return res, written
