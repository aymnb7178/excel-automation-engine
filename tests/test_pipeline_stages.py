import openpyxl
import pandas as pd

from excel_engine.cleaner import clean
from excel_engine.config import load_config
from excel_engine.dedupe import deduplicate, similarity
from excel_engine.loader import load_inputs, match_header
from excel_engine.validate import validate


def small_cfg(**over):
    d = {"columns": {"id": {"type": "text", "case": "upper", "missing": "drop_row"},
                     "name": {"type": "text", "case": "title", "arabic_letters": True, "missing": "constant:Unknown"},
                     "qty": {"type": "integer", "missing": "median"},
                     "price": {"type": "number", "missing": "zero"},
                     "day": {"type": "date"},
                     "tag": {"type": "text", "value_map": {"done": "Delivered"}, "missing": "mode"}}}
    d.update(over)
    return load_config(d)


def frame(rows):
    cols = ["id", "name", "qty", "price", "day", "tag"]
    df = pd.DataFrame(rows, columns=cols, dtype=object)
    df["source_file"], df["source_sheet"], df["source_row"] = "f.csv", "", range(2, 2 + len(df))
    return df


# ---- loader ----
def test_merge_reads_all_rows_and_detects_headers(inputs, cfg):
    d, truth = inputs
    res = load_inputs(d, cfg)
    assert len(res.df) == truth["data_rows_written"]
    north = res.files[res.files.file == "sales_north.xlsx"].iloc[0]
    assert north.header_row == 3  # banner row + blank row above the header
    assert set(res.files.status) == {"ok"} and len(res.files) == 6  # west file has two sheets
    assert list(res.df.columns[-3:]) == ["source_file", "source_sheet", "source_row"]


def test_header_aliases_and_fuzzy_matching(inputs, cfg):
    res = load_inputs(inputs[0], cfg)
    m = res.mapping
    ar = m[(m.file == "sales_arabic_headers.xlsx") & (m.source_header == "رقم الطلب")].iloc[0]
    assert ar.canonical == "order_id" and ar.method == "exact" and ar.source_col == 1
    fz = m[(m.file == "sales_west_two_sheets.xlsx") & (m.source_header == "Custmer")]
    assert set(fz.method) == {"fuzzy"} and set(fz.canonical) == {"customer"}
    assert (res.mapping.method == "unmapped").sum() == 1  # the 'Notes' column
    assert match_header("Totally Different", cfg.alias_index(), 0.85) == (None, "unmapped")


def test_corrupt_file_is_skipped_with_warning(tmp_path, cfg):
    (tmp_path / "broken.xlsx").write_bytes(b"not a zip")
    (tmp_path / "ok.csv").write_text("order_id,customer\nORD-1,Ali\n", encoding="utf-8")
    res = load_inputs(tmp_path, cfg)
    assert len(res.df) == 1 and any("broken.xlsx" in w for w in res.warnings)


# ---- cleaner ----
def test_cleaning_types_maps_and_missing_strategies():
    cfg = small_cfg()
    df = frame([["a-1 ", "  احمد ", "٥", "SAR 1,250.50", "12/03/2025", "done"],
                ["a-2", None, "7", "abc", "TBD", None],
                ["a-3", "x", None, "3", "2025-01-01", "done"],
                [None, "y", "1", "1", "2025-01-01", "done"]])
    r = clean(df, cfg)
    out = r.df
    assert len(out) == 3 and len(r.dropped) == 1  # drop_row on missing id
    assert out.id.tolist() == ["A-1", "A-2", "A-3"] and out.name[0] == "احمد" and out.name[1] == "Unknown"
    assert out.qty.tolist() == [5, 7, 5] and str(out.qty.dtype) == "Int64"  # median of 5, 7 and 1 (the row dropped for a missing id is still part of the median)
    assert out.price.tolist() == [1250.5, 0.0, 3.0]  # unparseable 'abc' -> missing -> zero
    assert out.day[0] == pd.Timestamp("2025-03-12") and pd.isna(out.day[1])
    assert out.tag.tolist() == ["Delivered", "Delivered", "Delivered"]  # value_map + mode fill
    assert set(r.failures.column) == {"price", "day"}
    assert r.log.set_index("column").loc["qty", "missing_filled"] == 1


# ---- dedupe ----
def test_exact_duplicates_all_found_and_most_complete_kept():
    cfg = small_cfg(duplicates={"key": ["id", "name"], "keep": "most_complete"})
    df = clean(frame([["A-1", "Ali", 3, 5, None, "t"],
                      ["a-1", " ali", 3, 5, "2025-01-01", "t"],
                      ["A-2", "Sara", 1, 5, "2025-01-01", "t"]]), cfg).df
    r = deduplicate(df, cfg)
    assert r.exact_removed == 1 and len(r.df) == 2
    assert r.df[r.df.id == "A-1"].iloc[0].source_row == 3  # the row that has a date
    assert r.review.action.tolist() == ["removed", "kept"] or r.review.action.tolist() == ["kept", "removed"]


def test_injected_duplicates_detected_on_synthetic_data(result, inputs):
    truth = inputs[1]["injected"]
    assert result.dedupe.exact_removed == truth["exact_duplicates"]
    rev = result.dedupe.review
    fz = rev[rev.match_type == "fuzzy"]
    hits = sum(g.order_id.str.match(r"ORD-9\d{3}$").any() for _, g in fz.groupby("group_id"))
    assert hits == truth["fuzzy_duplicates"]
    assert (fz.action == "review").all() and len(result.df) == result.after["rows"]


def test_similarity_is_typo_and_order_tolerant():
    assert similarity("ahmed alharbi", "ahmed alharb") > 0.9
    assert similarity("alharbi ahmed", "ahmed alharbi") == 1.0
    assert similarity("ahmed", "sara") < 0.5 and similarity("", "x") == 0.0


def test_fuzzy_skipped_when_unblocked_and_too_large():
    cfg = small_cfg(duplicates={"fuzzy_columns": ["name"], "max_unblocked_rows": 2})
    df = clean(frame([["A-1", "Ali", 1, 1, "2025-01-01", "t"]] * 3), cfg).df
    assert any("skipped" in w for w in deduplicate(df, cfg).warnings)


# ---- validation ----
def test_all_rule_types():
    rules = [{"column": "id", "rule": "unique"}, {"column": "id", "rule": "regex", "pattern": "A-\\d"},
             {"column": "qty", "rule": "range", "min": 1, "max": 9}, {"column": "tag", "rule": "allowed", "values": ["x", "y"]},
             {"column": "day", "rule": "required", "severity": "warning"},
             {"column": "day", "rule": "range", "min": "2025-01-01", "max": "2025-12-31"}]
    cfg = small_cfg(validation=rules)
    df = clean(frame([["A-1", "n", 1, 1, "2025-05-05", "X"], ["A-1", "n", 20, 1, "2025-05-05", "z"],
                      ["B-2", "n", 5, 1, None, "y"], ["A-3", "n", 5, 1, "2030-01-01", "y"]]), cfg).df
    e = validate(df, cfg)
    got = {(r.rule, r.column, r.clean_row) for r in e.itertuples()}
    assert got == {("unique", "id", 3), ("regex", "id", 4), ("range", "qty", 3), ("allowed", "tag", 3),
                   ("required", "day", 4), ("range", "day", 5)}
    assert e[e.rule == "required"].severity.iloc[0] == "warning" and e.severity.iloc[0] == "error"


def test_error_cells_point_at_the_original_files(result, inputs, cfg):
    d = inputs[0]
    checked = 0
    for r in result.errors.itertuples():
        if r.rule != "range" or r.column != "quantity" or not r.source_cell:
            continue
        f = d / r.source_file
        if f.suffix == ".xlsx":
            ws = openpyxl.load_workbook(f)[r.source_sheet]
            raw = ws[r.source_cell].value
            assert str(raw).replace("-", "").replace("−", "").strip("٠١٢٣٤٥٦٧٨٩0123456789") == ""
            assert int(str(raw).translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))) == int(r.value)
            checked += 1
    assert checked >= 1
    sheet_cell = result.errors.clean_cell.iloc[0]
    assert sheet_cell[0].isalpha() and sheet_cell[1:].isdigit()
