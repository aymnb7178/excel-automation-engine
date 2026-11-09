"""Command line interface: merge, clean, validate, report, run."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

from .cleaner import clean
from .config import ConfigError, load_config
from .loader import load_inputs
from .pipeline import process, run_pipeline_file
from .report import build_workbook, preview_png, write_simple_workbook


def _write_table(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.suffix.lower() == ".csv":
        df.to_csv(path, index=False, encoding="utf-8-sig")
    elif path.suffix.lower() == ".xlsx":
        write_simple_workbook(df, path)
    else:
        raise ConfigError(f"unsupported output extension '{path.suffix}' (use .csv or .xlsx)")


def _headline(res) -> list[str]:
    b, a = res.before, res.after
    return [f"rows read: {b['rows']:,}   rows clean: {a['rows']:,}   "
            f"dropped: {len(res.clean.dropped)}   exact duplicates removed: {res.dedupe.exact_removed}   "
            f"fuzzy pairs to review: {res.dedupe.fuzzy_pairs}",
            f"quality score: {b['quality_score']:.1f} -> {a['quality_score']:.1f}   "
            f"rule errors: {b['rule_violations']} -> {a['rule_violations']}"]


def _cmd_merge(args) -> int:
    cfg = load_config(args.config)
    res = load_inputs(args.input, cfg)
    _write_table(res.df, Path(args.out))
    m = res.mapping["method"].value_counts().to_dict()
    print(f"Merged {len(res.df):,} rows from {int((res.files['status'] == 'ok').sum())} file/sheet table(s) -> {args.out}")
    print(f"Header mapping: {m.get('exact', 0)} exact, {m.get('fuzzy', 0)} fuzzy, {m.get('unmapped', 0)} unmapped")
    for w in res.warnings:
        print("warning:", w, file=sys.stderr)
    return 0


def _cmd_clean(args) -> int:
    cfg = load_config(args.config)
    res = process(args.input, cfg, dedupe=not args.no_dedupe, validate_rules=False)
    _write_table(res.df, Path(args.out))
    print(f"Cleaned dataset -> {args.out}")
    print("\n".join(_headline(res)))
    return 0


def _cmd_validate(args) -> int:
    cfg = load_config(args.config)
    res = process(args.input, cfg)
    _write_table(res.errors, Path(args.out))
    e = res.errors
    n_err = int((e["severity"] == "error").sum()) if len(e) else 0
    print(f"Validation: {n_err} error(s), {len(e) - n_err} warning(s) on {len(res.df):,} clean rows -> {args.out}")
    if len(e):
        print(e.groupby(["column", "rule", "severity"]).size().rename("count").to_string())
    return 1 if (args.fail_on_error and n_err) else 0


def _cmd_report(args) -> int:
    cfg = load_config(args.config)
    res = process(args.input, cfg)
    print(f"Workbook -> {build_workbook(res, args.out)}")
    if args.preview:
        print(f"Preview  -> {preview_png(res, args.preview)}")
    print("\n".join(_headline(res)))
    return 0


def _cmd_run(args) -> int:
    res, written = run_pipeline_file(args.pipeline)
    for k, v in written.items():
        print(f"{k}: {v}")
    print("\n".join(_headline(res)))
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="excel-engine", description="Merge, clean, validate and report on messy spreadsheets.")
    sub = p.add_subparsers(dest="cmd", required=True)

    def common(sp, out_help):
        sp.add_argument("input", help="folder (or single .xlsx/.csv file) of input data")
        sp.add_argument("--config", required=True, help="JSON schema/rules config")
        sp.add_argument("--out", required=True, help=out_help)

    sp = sub.add_parser("merge", help="merge files with header detection and schema mapping only")
    common(sp, "output .csv or .xlsx"); sp.set_defaults(fn=_cmd_merge)
    sp = sub.add_parser("clean", help="merge + clean + de-duplicate")
    common(sp, "output .csv or .xlsx"); sp.add_argument("--no-dedupe", action="store_true"); sp.set_defaults(fn=_cmd_clean)
    sp = sub.add_parser("validate", help="merge + clean + validate; writes the error report")
    common(sp, "error report .csv or .xlsx"); sp.add_argument("--fail-on-error", action="store_true", help="exit 1 if any error remains")
    sp.set_defaults(fn=_cmd_validate)
    sp = sub.add_parser("report", help="full pipeline -> formatted multi-sheet workbook")
    common(sp, "output .xlsx"); sp.add_argument("--preview", help="optional PNG preview path"); sp.set_defaults(fn=_cmd_report)
    sp = sub.add_parser("run", help="run a JSON pipeline file")
    sp.add_argument("pipeline"); sp.set_defaults(fn=_cmd_run)
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return args.fn(args)
    except (ConfigError, FileNotFoundError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
