import json

import openpyxl
import pandas as pd
import pytest

from excel_engine.cli import main
from excel_engine.config import load_config
from excel_engine.datagen import generate_messy_inputs
from excel_engine.pipeline import process
from excel_engine.report import build_pivot, build_workbook, preview_png
from excel_engine.config import PivotSpec


def test_scorecard_improves_and_is_consistent(result):
    b, a = result.before, result.after
    assert a["quality_score"] > b["quality_score"] + 20
    assert a["rule_violations"] < b["rule_violations"] and a["native_typed_pct"] == 100.0
    assert b["exact_duplicate_rows"] == 30 and a["exact_duplicate_rows"] == 0
    sc = result.scorecard.set_index("metric")
    assert sc.loc["Quality score (0-100)", "verdict"] == "improved"
    assert a["rows"] == len(result.df) == b["rows"] - 3 - 30


def test_pivot_sums_match_data(result):
    p = build_pivot(result.df, PivotSpec(name="x", index="region", values="line_total", agg="sum"))
    assert p["line_total"].sum() == pytest.approx(result.df["line_total"].sum(), abs=0.05 * len(p))
    m = build_pivot(result.df, PivotSpec(name="m", index="order_date", values="line_total", agg="sum", freq="M"))
    assert m["order_date"].is_monotonic_increasing and m["order_date"].str.match(r"\d{4}-\d{2}$|\(missing\)$").all()
    c = build_pivot(result.df, PivotSpec(name="c", index="region", columns="status", agg="count"))
    assert c.iloc[:, 1:].to_numpy().sum() == len(result.df)


def test_workbook_structure(result, tmp_path):
    path = build_workbook(result, tmp_path / "r.xlsx")
    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames[:2] == ["Summary", "Clean Data"]
    assert {"Duplicates Review", "Validation Errors", "Header Mapping", "Cleaning Log"} <= set(wb.sheetnames)
    assert sum(len(wb[s]._charts) for s in wb.sheetnames) >= 4  # 3 pivot charts + summary chart
    cd = wb["Clean Data"]
    assert cd.freeze_panes == "A2" and cd.auto_filter.ref and cd.max_row == len(result.df) + 1
    assert cd.column_dimensions["A"].width >= 8 and cd["A1"].style == "eae_header"
    assert wb["Validation Errors"].conditional_formatting and wb["Summary"].conditional_formatting
    assert {"eae_header", "eae_title"} <= set(wb.named_styles)
    date_col = list(result.df.columns).index("order_date") + 1
    assert cd.cell(2, date_col).number_format == "yyyy-mm-dd"


def test_formula_like_text_stays_text(result, tmp_path):
    df = result.df.copy()
    df.loc[0, "customer"] = '=HYPERLINK("http://x","click")'
    res = result.__class__(**{**result.__dict__, "dedupe": result.dedupe.__class__(df, result.dedupe.review)})
    wb = openpyxl.load_workbook(build_workbook(res, tmp_path / "f.xlsx"))
    cell = wb["Clean Data"].cell(2, list(df.columns).index("customer") + 1)
    assert cell.data_type == "s" and cell.value.startswith("=")


def test_preview_png_written(result, tmp_path):
    p = preview_png(result, tmp_path / "p.png")
    assert p.read_bytes()[:8] == b"\x89PNG\r\n\x1a\n" and p.stat().st_size > 10_000


def test_datagen_is_deterministic(tmp_path):
    t1 = generate_messy_inputs(tmp_path / "a", seed=7, n_orders=60)
    t2 = generate_messy_inputs(tmp_path / "b", seed=7, n_orders=60)
    assert t1["rows_per_file"] == t2["rows_per_file"]
    assert (tmp_path / "a" / "sales_east_legacy.csv").read_bytes() == (tmp_path / "b" / "sales_east_legacy.csv").read_bytes()


def test_cli_subcommands_and_pipeline(inputs, tmp_path, capsys):
    d, truth = inputs
    cfgp = "examples/config.json"
    assert main(["merge", str(d), "--config", cfgp, "--out", str(tmp_path / "m.csv")]) == 0
    assert len(pd.read_csv(tmp_path / "m.csv")) == truth["data_rows_written"]
    assert main(["clean", str(d), "--config", cfgp, "--out", str(tmp_path / "c.xlsx")]) == 0
    assert openpyxl.load_workbook(tmp_path / "c.xlsx")["Data"].max_row == 517
    assert main(["validate", str(d), "--config", cfgp, "--out", str(tmp_path / "e.csv"), "--fail-on-error"]) == 1
    assert main(["validate", str(d), "--config", cfgp, "--out", str(tmp_path / "e2.csv")]) == 0
    assert main(["report", str(d), "--config", cfgp, "--out", str(tmp_path / "r.xlsx"), "--preview", str(tmp_path / "p.png")]) == 0
    pipe = tmp_path / "pipe.json"
    pipe.write_text(json.dumps({"input": str(d), "config": str(__import__("pathlib").Path(cfgp).resolve()),
                                "output_dir": "out", "clean_csv": "c.csv", "errors_csv": "e.csv"}))
    assert main(["run", str(pipe)]) == 0
    assert (tmp_path / "out" / "clean_report.xlsx").exists() and len(pd.read_csv(tmp_path / "out" / "c.csv")) == 516
    out = capsys.readouterr().out
    assert "quality score" in out


def test_cli_errors_are_reported_not_raised(tmp_path, capsys):
    assert main(["merge", str(tmp_path / "missing"), "--config", "examples/config.json", "--out", str(tmp_path / "x.csv")]) == 2
    assert main(["merge", str(tmp_path), "--config", str(tmp_path / "nope.json"), "--out", str(tmp_path / "x.csv")]) == 2
    bad = tmp_path / "p.json"
    bad.write_text('{"input": "x"}')
    assert main(["run", str(bad)]) == 2
    assert "error:" in capsys.readouterr().err
