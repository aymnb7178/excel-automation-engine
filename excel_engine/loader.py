"""Read xlsx/csv files, auto-detect the header row, map headers to canonical columns, merge."""

from __future__ import annotations

import csv
import difflib
import io
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from .config import META_COLUMNS, Config
from .textnorm import header_key, is_missing

SUPPORTED = {".xlsx", ".xlsm", ".csv"}
_ENCODINGS = ("utf-8-sig", "utf-8", "cp1256", "latin-1")


@dataclass
class LoadResult:
    df: pd.DataFrame
    mapping: pd.DataFrame
    files: pd.DataFrame
    warnings: list[str] = field(default_factory=list)


def discover_files(path: str | Path) -> list[Path]:
    p = Path(path)
    if p.is_file():
        return [p]
    if not p.is_dir():
        raise FileNotFoundError(f"input path not found: {p}")
    return sorted(f for f in p.iterdir() if f.suffix.lower() in SUPPORTED and not f.name.startswith("~$"))


def _frame(rows: list[tuple]) -> pd.DataFrame:
    width = max((len(r) for r in rows), default=0)
    return pd.DataFrame([list(r) + [None] * (width - len(r)) for r in rows], dtype=object)


def _read_csv(path: Path) -> list[tuple[str, pd.DataFrame]]:
    raw = path.read_bytes()
    text = None
    for enc in _ENCODINGS:
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    try:
        delim = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|").delimiter
    except csv.Error:
        delim = ","
    return [("", _frame(list(csv.reader(io.StringIO(text), delimiter=delim))))]


def _read_xlsx(path: Path) -> list[tuple[str, pd.DataFrame]]:
    wb = load_workbook(path, data_only=True)
    try:
        return [(ws.title, _frame(list(ws.iter_rows(values_only=True)))) for ws in wb.worksheets]
    finally:
        wb.close()


def match_header(h, alias_index: dict[str, str], cutoff: float) -> tuple[str | None, str]:
    """Map a raw header to a canonical column: ('name', 'exact'|'fuzzy') or (None, 'unmapped')."""
    k = header_key(h)
    if not k:
        return None, "unmapped"
    if k in alias_index:
        return alias_index[k], "exact"
    close = difflib.get_close_matches(k, list(alias_index), n=1, cutoff=cutoff)
    if close:
        return alias_index[close[0]], "fuzzy"
    return None, "unmapped"


def detect_header_row(raw: pd.DataFrame, cfg: Config) -> int | None:
    """Row (0-based) among the first N with the most recognisable headers; ties -> earliest."""
    idx, best, best_hits = cfg.alias_index(), None, 0
    for i in range(min(cfg.header_scan_rows, len(raw))):
        cells = [c for c in raw.iloc[i].tolist() if not is_missing(c) and isinstance(c, str)]
        hits = sum(match_header(c, idx, cfg.fuzzy_header_cutoff)[0] is not None for c in cells)
        if hits > best_hits:
            best, best_hits = i, hits
    return best


def load_inputs(path: str | Path, cfg: Config) -> LoadResult:
    alias_index = cfg.alias_index()
    canon = list(cfg.columns)
    frames, maps, files, warnings = [], [], [], []
    for f in discover_files(path):
        try:
            tables = _read_csv(f) if f.suffix.lower() == ".csv" else _read_xlsx(f)
        except Exception as exc:  # unreadable/corrupt file: record and continue
            warnings.append(f"{f.name}: could not be read ({type(exc).__name__}: {exc})")
            files.append({"file": f.name, "sheet": "", "header_row": None, "rows": 0, "status": "unreadable"})
            continue
        for sheet, raw in tables:
            h = detect_header_row(raw, cfg) if len(raw) else None
            if h is None:
                warnings.append(f"{f.name}[{sheet}]: no recognisable header row, skipped")
                files.append({"file": f.name, "sheet": sheet, "header_row": None, "rows": 0, "status": "skipped"})
                continue
            used: dict[int, str] = {}
            for j, cell in enumerate(raw.iloc[h].tolist()):
                if is_missing(cell):
                    continue
                name, method = match_header(cell, alias_index, cfg.fuzzy_header_cutoff)
                if name is not None and name in used.values():
                    name, method = None, "duplicate_mapping"
                if name is not None:
                    used[j] = name
                maps.append({"file": f.name, "sheet": sheet, "source_header": str(cell).strip(),
                             "canonical": name or "", "method": method, "source_col": j + 1})
            body = raw.iloc[h + 1:]
            data = pd.DataFrame({name: body.iloc[:, j].values for j, name in used.items()}, dtype=object)
            if cfg.keep_unmapped:
                for j, cell in enumerate(raw.iloc[h].tolist()):
                    if j not in used and not is_missing(cell):
                        data[f"unmapped: {str(cell).strip()}"] = body.iloc[:, j].values
            data["source_file"], data["source_sheet"] = f.name, sheet
            data["source_row"] = (body.index.values + 1).astype(int)  # 1-based row in the file
            nonblank = ~data[[c for c in used.values()]].map(is_missing).all(axis=1) if used else data.index < 0
            data = data[nonblank.values]
            frames.append(data)
            files.append({"file": f.name, "sheet": sheet, "header_row": h + 1, "rows": len(data), "status": "ok"})
    if frames:
        merged = pd.concat(frames, ignore_index=True)
    else:
        merged = pd.DataFrame(columns=canon + META_COLUMNS)
    extras = [c for c in merged.columns if c not in canon and c not in META_COLUMNS]
    merged = merged.reindex(columns=canon + extras + META_COLUMNS)
    merged[canon] = merged[canon].astype(object)
    return LoadResult(merged, pd.DataFrame(maps, columns=["file", "sheet", "source_header", "canonical", "method", "source_col"]),
                      pd.DataFrame(files, columns=["file", "sheet", "header_row", "rows", "status"]), warnings)
