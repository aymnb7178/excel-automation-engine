"""End-to-end demo: generate messy SYNTHETIC files -> run the JSON pipeline -> print real numbers."""

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from excel_engine.datagen import generate_messy_inputs  # noqa: E402
from excel_engine.pipeline import run_pipeline_file  # noqa: E402


def main() -> None:
    truth = generate_messy_inputs(ROOT / "data" / "messy_inputs", seed=42)
    print(f"Generated {len(truth['files'])} messy files, {truth['data_rows_written']:,} data rows (seed 42):")
    for f, n in truth["rows_per_file"].items():
        print(f"  {f:<30} {n:>4} rows")
    print("Injected:", ", ".join(f"{k}={v}" for k, v in truth["injected"].items()))
    res, written = run_pipeline_file(ROOT / "examples" / "pipeline.json")
    shutil.copyfile(written["workbook"], ROOT / "docs" / "sample_report.xlsx")
    print("\nPipeline outputs:")
    for k, p in written.items():
        print(f"  {k:<12} {Path(p).relative_to(ROOT)}")
    inj = truth["injected"]
    rev = res.dedupe.review
    fz = rev[rev.match_type == "fuzzy"]
    print(f"\nRows read {res.before['rows']:,} -> dropped {len(res.clean.dropped)} (no id), "
          f"exact duplicates removed {res.dedupe.exact_removed} (injected {inj['exact_duplicates']}), "
          f"clean rows {res.after['rows']:,}")
    injected_hit = sum(1 for _, g in fz.groupby("group_id") if g["order_id"].str.match(r"ORD-9\d{3}$").any())
    print(f"Fuzzy duplicate pairs flagged for review: {res.dedupe.fuzzy_pairs}; "
          f"{injected_hit} of {inj['fuzzy_duplicates']} injected fuzzy duplicates found, "
          f"{res.dedupe.fuzzy_pairs - injected_hit} other pair(s) to review")
    print(f"Cells unparseable (set to missing): {len(res.clean.failures)}")
    print("\nScorecard:")
    print(res.scorecard.to_string(index=False))
    e = res.errors
    print("\nRemaining validation issues by column/rule/severity:")
    print(e.groupby(["column", "rule", "severity"]).size().rename("count").to_string())
    print("\nDone. Open docs/sample_report.xlsx and docs/preview.png.")


if __name__ == "__main__":
    main()
