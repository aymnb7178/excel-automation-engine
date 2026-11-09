"""Formatted multi-sheet workbook (openpyxl) and a PNG preview (matplotlib, Agg backend)."""

from __future__ import annotations

import math
import re
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from openpyxl.utils import get_column_letter

from .config import Config, PivotSpec
from .scorecard import PCT_KEYS, PCT_LABELS

NAVY, GREEN_FILL, RED_FILL, AMBER_FILL = "1F3864", "C6EFCE", "FFC7CE", "FFEB9C"


def _py(v):
    if v is None or v is pd.NA or v is pd.NaT:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if isinstance(v, np.integer):
        return int(v)
    if isinstance(v, np.floating):
        return float(v)
    if isinstance(v, np.bool_):
        return bool(v)
    return v


def _register_styles(wb: Workbook) -> None:
    thin = Side(style="thin", color="BFBFBF")
    specs = {
        "eae_title": dict(font=Font(name="Calibri", size=18, bold=True, color=NAVY)),
        "eae_subtitle": dict(font=Font(name="Calibri", size=10, italic=True, color="595959")),
        "eae_section": dict(font=Font(name="Calibri", size=12, bold=True, color=NAVY),
                            border=Border(bottom=Side(style="medium", color=NAVY))),
        "eae_header": dict(font=Font(name="Calibri", bold=True, color="FFFFFF"),
                           fill=PatternFill("solid", fgColor=NAVY),
                           alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
                           border=Border(bottom=thin)),
        "eae_cell": dict(font=Font(name="Calibri", size=11), border=Border(bottom=Side(style="hair", color="D9D9D9"))),
        "eae_label": dict(font=Font(name="Calibri", size=11, bold=True)),
    }
    for name, kw in specs.items():
        st = NamedStyle(name=name)
        for k, v in kw.items():
            setattr(st, k, v)
        wb.add_named_style(st)


def _set(ws, row: int, col: int, value, style: str = "eae_cell", fmt: str | None = None):
    c = ws.cell(row=row, column=col)
    v = _py(value)
    if isinstance(v, str):
        v = ILLEGAL_CHARACTERS_RE.sub("", v)
    c.value = v
    if isinstance(v, str):
        c.data_type = "s"  # never let a cell be interpreted as a formula ("=...", "+...", "@...")
    c.style = style
    if fmt:
        c.number_format = fmt
    return c


def write_df(ws, df: pd.DataFrame, row: int = 1, col: int = 1, freeze: bool = True, filt: bool = True) -> int:
    """Write df with a styled header; returns the last row used."""
    for j, name in enumerate(df.columns):
        _set(ws, row, col + j, str(name), "eae_header")
        kind = df[name].dtype
        fmt = None
        if pd.api.types.is_datetime64_any_dtype(kind):
            fmt = "yyyy-mm-dd"
        elif pd.api.types.is_float_dtype(kind):
            fmt = "#,##0.00"
        for i, v in enumerate(df[name].tolist(), start=1):
            _set(ws, row + i, col + j, v, "eae_cell", fmt)
    ws.row_dimensions[row].height = 30
    if freeze:
        ws.freeze_panes = ws.cell(row=row + 1, column=col)
    if filt and len(df.columns):
        ws.auto_filter.ref = f"{get_column_letter(col)}{row}:{get_column_letter(col + len(df.columns) - 1)}{row + len(df)}"
    return row + len(df)


def autofit(ws, min_w: int = 8, max_w: int = 60) -> None:
    widths: dict[int, int] = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is None:
                continue
            v = c.value
            n = 10 if isinstance(v, datetime) else len(f"{v:,.2f}") if isinstance(v, float) else len(str(v))
            if c.style == "eae_title":
                continue  # long titles may overflow to the right
            widths[c.column] = max(widths.get(c.column, 0), n)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = max(min_w, min(max_w, w + 3))


def build_pivot(df: pd.DataFrame, spec: PivotSpec) -> pd.DataFrame:
    d = df
    idx = d[spec.index]
    if spec.freq == "M":
        idx = idx.dt.to_period("M").astype(str).where(idx.notna())
    key = idx.astype(object).where(idx.notna(), "(missing)").astype(str)
    vals = None if spec.values is None else d[spec.values].astype("float64")
    if spec.columns:
        ck = d[spec.columns].astype(object).where(d[spec.columns].notna(), "(missing)").astype(str)
        if spec.agg == "count":
            piv = pd.crosstab(key, ck)
        else:
            piv = pd.crosstab(key, ck, values=vals, aggfunc=spec.agg).fillna(0)
    else:
        grp = pd.Series(1.0, index=d.index) if vals is None or spec.agg == "count" else vals
        g = grp.groupby(key)
        piv = (g.count() if spec.agg == "count" else g.sum() if spec.agg == "sum" else g.mean()).to_frame(
            spec.values or "count")
    piv = piv.astype("float64").round(2)
    piv = piv.sort_index() if spec.freq else piv.loc[piv.sum(axis=1).sort_values(ascending=False, kind="stable").index]
    piv.columns = [str(c) for c in piv.columns]
    piv.index.name = spec.index
    return piv.reset_index()


def build_pivots(df: pd.DataFrame, cfg: Config) -> dict[str, pd.DataFrame]:
    return {p.name: build_pivot(df, p) for p in cfg.report.pivots}


def _sheet_name(base: str, used: set[str]) -> str:
    n = re.sub(r"[\\/*?:\[\]]", " ", base)[:31].strip() or "Sheet"
    k = 2
    while n in used:
        n = f"{n[:28]} {k}"
        k += 1
    used.add(n)
    return n


def _chart(ws, spec, n_rows: int, n_cols: int, hdr: int, anchor: str):
    ch = BarChart() if spec.kind == "bar" else LineChart()
    ch.title, ch.y_axis.title = spec.title or None, spec.y_title or None
    ch.add_data(Reference(ws, min_col=2, max_col=n_cols, min_row=hdr, max_row=hdr + n_rows), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=hdr + 1, max_row=hdr + n_rows))
    ch.x_axis.delete = ch.y_axis.delete = False
    ch.width, ch.height = 20, 10
    if n_cols == 2:
        ch.legend = None
    ws.add_chart(ch, anchor)


def build_workbook(res, path: str | Path) -> Path:
    """Write the multi-sheet report for a PipelineResult."""
    cfg, df, sc = res.cfg, res.df, res.scorecard
    wb = Workbook()
    _register_styles(wb)
    used: set[str] = set()
    ws = wb.active
    ws.title = _sheet_name("Summary", used)
    _set(ws, 1, 1, cfg.report.title, "eae_title")
    _set(ws, 2, 1, cfg.report.subtitle or f"Generated {datetime.now():%Y-%m-%d %H:%M}", "eae_subtitle")
    _set(ws, 4, 1, "Pipeline headline", "eae_section"); _set(ws, 4, 2, None, "eae_section")
    files_ok = int((res.load.files["status"] == "ok").sum())
    headline = [("Input files / sheets read", files_ok), ("Rows read (after header detection)", res.before["rows"]),
                ("Rows dropped (missing required id)", len(res.clean.dropped)),
                ("Exact duplicate rows removed", res.dedupe.exact_removed),
                ("Fuzzy duplicate pairs flagged for review", res.dedupe.fuzzy_pairs),
                ("Rows in clean dataset", res.after["rows"]),
                ("Cells that could not be parsed (set to missing)", len(res.clean.failures)),
                ("Validation errors", int((res.errors["severity"] == "error").sum()) if len(res.errors) else 0),
                ("Validation warnings", int((res.errors["severity"] == "warning").sum()) if len(res.errors) else 0)]
    for i, (k, v) in enumerate(headline, start=5):
        _set(ws, i, 1, k, "eae_label"); _set(ws, i, 2, v, "eae_cell", "#,##0")
    r0 = 5 + len(headline) + 1
    _set(ws, r0, 1, "Data-quality scorecard (before vs after)", "eae_section")
    for c in range(2, 6):
        _set(ws, r0, c, None, "eae_section")
    for j, h in enumerate(["Metric", "Before", "After", "Change", "Verdict"], start=1):
        _set(ws, r0 + 1, j, h, "eae_header")
    for i, row in enumerate(sc.itertuples(index=False), start=r0 + 2):
        for j, v in enumerate(row, start=1):
            _set(ws, i, j, v, "eae_cell", "#,##0.0" if j in (2, 3, 4) and "(" in row[0] else "#,##0" if j in (2, 3, 4) else None)
    vr = f"E{r0 + 2}:E{r0 + 1 + len(sc)}"
    ws.conditional_formatting.add(vr, CellIsRule(operator="equal", formula=['"improved"'], fill=PatternFill("solid", bgColor=GREEN_FILL)))
    ws.conditional_formatting.add(vr, CellIsRule(operator="equal", formula=['"worse"'], fill=PatternFill("solid", bgColor=RED_FILL)))
    ws.conditional_formatting.add(vr, CellIsRule(operator="equal", formula=['"unchanged"'], fill=PatternFill("solid", bgColor=AMBER_FILL)))
    # percentage block feeding the native chart
    pr = r0 + len(sc) + 4
    _set(ws, pr - 1, 1, "Quality percentages (chart data)", "eae_section")
    for c in range(2, 4):
        _set(ws, pr - 1, c, None, "eae_section")
    for j, h in enumerate(["Dimension", "Before", "After"], start=1):
        _set(ws, pr, j, h, "eae_header")
    for i, k in enumerate(PCT_KEYS, start=pr + 1):
        _set(ws, i, 1, PCT_LABELS[k].replace(" (%)", "").replace(" (0-100)", " score"), "eae_cell")
        _set(ws, i, 2, round(res.before[k], 1), "eae_cell", "0.0"); _set(ws, i, 3, round(res.after[k], 1), "eae_cell", "0.0")
    ch = BarChart(); ch.type = "col"; ch.title = "Data quality: before vs after (%)"; ch.y_axis.title = "%"
    ch.add_data(Reference(ws, min_col=2, max_col=3, min_row=pr, max_row=pr + len(PCT_KEYS)), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=pr + 1, max_row=pr + len(PCT_KEYS)))
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = 0, 100
    ch.x_axis.delete = ch.y_axis.delete = False
    ch.width, ch.height = 18, 9
    ws.add_chart(ch, f"G{r0}")
    ws.sheet_view.showGridLines = False
    autofit(ws, max_w=48)
    ws.column_dimensions["A"].width = 52

    ws = wb.create_sheet(_sheet_name("Clean Data", used))
    write_df(ws, df)
    autofit(ws)

    for name, piv in build_pivots(df, cfg).items():
        ws = wb.create_sheet(_sheet_name(f"Pivot {name}", used))
        _set(ws, 1, 1, f"Pivot: {name}", "eae_title")
        end = write_df(ws, piv, row=3, filt=False)
        if len(piv):
            ws.conditional_formatting.add(f"B4:{get_column_letter(len(piv.columns))}{end}",
                                          DataBarRule(start_type="min", end_type="max", color="5B9BD5"))
        for spec in [c for c in cfg.report.charts if c.pivot == name]:
            _chart(ws, spec, len(piv), len(piv.columns), 3, f"{get_column_letter(len(piv.columns) + 2)}3")
        autofit(ws)

    ws = wb.create_sheet(_sheet_name("Duplicates Review", used))
    if len(res.dedupe.review):
        write_df(ws, res.dedupe.review)
        end = len(res.dedupe.review) + 1
        ws.conditional_formatting.add(f"C2:C{end}", CellIsRule(operator="equal", formula=['"removed"'], fill=PatternFill("solid", bgColor=RED_FILL)))
        ws.conditional_formatting.add(f"C2:C{end}", CellIsRule(operator="equal", formula=['"review"'], fill=PatternFill("solid", bgColor=AMBER_FILL)))
    else:
        write_df(ws, res.dedupe.review, filt=False)
    autofit(ws)

    ws = wb.create_sheet(_sheet_name("Validation Errors", used))
    write_df(ws, res.errors, filt=bool(len(res.errors)))
    if len(res.errors):
        end = len(res.errors) + 1
        ws.conditional_formatting.add(f"A2:A{end}", CellIsRule(operator="equal", formula=['"error"'], fill=PatternFill("solid", bgColor=RED_FILL)))
        ws.conditional_formatting.add(f"A2:A{end}", CellIsRule(operator="equal", formula=['"warning"'], fill=PatternFill("solid", bgColor=AMBER_FILL)))
    autofit(ws)

    for title, frame in (("Header Mapping", res.load.mapping), ("Files Read", res.load.files),
                         ("Cleaning Log", res.clean.log), ("Unparsed Cells", res.clean.failures)):
        ws = wb.create_sheet(_sheet_name(title, used))
        write_df(ws, frame, filt=False)
        autofit(ws)
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def preview_png(res, path: str | Path) -> Path:
    """Four-panel PNG preview of the pipeline result (before/after quality, issues, pivot, row flow)."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    BLUE, ORANGE, GREY, INK = "#2F6DB5", "#E08A1E", "#8C8C8C", "#222222"
    plt.rcParams.update({"font.size": 10, "axes.spines.top": False, "axes.spines.right": False, "text.color": INK})
    fig, ax = plt.subplots(2, 2, figsize=(13, 8.2))
    fig.suptitle("Excel Automation Engine: demo run on SYNTHETIC messy files", fontsize=14, fontweight="bold")
    a = ax[0, 0]
    labels = [PCT_LABELS[k].replace(" (%)", "").replace(" (0-100)", "").replace("Cells stored as native numbers/dates", "Native types") for k in PCT_KEYS]
    x = np.arange(len(PCT_KEYS))
    b = [res.before[k] for k in PCT_KEYS]; af = [res.after[k] for k in PCT_KEYS]
    a.bar(x - 0.2, b, 0.4, label="Before", color=ORANGE); a.bar(x + 0.2, af, 0.4, label="After", color=BLUE)
    for xi, v in zip(x - 0.2, b):
        a.text(xi, v + 1, f"{v:.0f}", ha="center", fontsize=8)
    for xi, v in zip(x + 0.2, af):
        a.text(xi, v + 1, f"{v:.0f}", ha="center", fontsize=8)
    a.set_xticks(x, labels, rotation=20, ha="right"); a.set_ylim(0, 125); a.set_ylabel("%")
    a.set_title("Data-quality scorecard"); a.legend(frameon=False, loc="upper left", ncol=2)
    a = ax[0, 1]
    if len(res.errors):
        c = res.errors.groupby(["column", "rule"]).size().sort_values()
        a.barh([f"{i[0]} / {i[1]}" for i in c.index], c.values, color=BLUE)
        for yi, v in enumerate(c.values):
            a.text(v + 0.2, yi, str(v), va="center", fontsize=8)
    a.set_title("Remaining validation issues (after cleaning)"); a.set_xlabel("cells")
    a = ax[1, 0]
    pivs = build_pivots(res.df, res.cfg)
    if pivs:
        name, p = next(iter(pivs.items()))
        a.bar(p.iloc[:, 0].astype(str), p.iloc[:, 1], color=BLUE)
        a.set_title(f"Pivot: {name}"); a.set_ylabel(p.columns[1]); a.tick_params(axis="x", rotation=30)
    a = ax[1, 1]
    steps = [("Rows read", res.before["rows"]), ("Dropped\n(no id)", -len(res.clean.dropped)),
             ("Exact dupes\nremoved", -res.dedupe.exact_removed), ("Clean rows", res.after["rows"])]
    run = 0
    for i, (lab, v) in enumerate(steps):
        if i in (0, 3):
            a.bar(i, v, color=BLUE); a.text(i, v + 5, f"{v:,}", ha="center", fontsize=9); run = v
        else:
            a.bar(i, v, bottom=run, color=ORANGE); a.text(i, run + 5, f"{v:,}", ha="center", fontsize=9); run += v
    a.set_xticks(range(4), [s[0] for s in steps]); a.set_title("Row flow", pad=12)
    a.set_ylim(0, max(res.before["rows"], 1) * 1.15)
    fig.text(0.99, 0.005, "All data synthetic.", ha="right", fontsize=8, color=GREY)
    fig.tight_layout(rect=(0, 0.02, 1, 0.96))
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=110)
    plt.close(fig)
    return path


def write_simple_workbook(df: pd.DataFrame, path: str | Path, sheet: str = "Data") -> Path:
    wb = Workbook()
    _register_styles(wb)
    ws = wb.active
    ws.title = sheet
    write_df(ws, df)
    autofit(ws)
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return Path(path)
