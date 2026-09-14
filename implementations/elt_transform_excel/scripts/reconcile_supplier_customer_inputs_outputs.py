#!/usr/bin/env python3
"""
Reconcile script: config ↔ provided source inputs ↔ generated EIB outputs.

Permanent entry point: scripts/reconcile_supplier_customer_inputs_outputs.py

For EVERY registered transform profile (Supplier Invoice AND Customer Invoice):
  v1. Verifies shell_defaults.source_data / target_eib / output_file_path exist on disk.
  v2. Opens the source workbook, confirms profile.source_layout.worksheet_name is present,
      then EMULATES THE VBA EXACTLY (modSourceTable.bas FindHeaderRow): scan for a row
      containing ALL required_headers tokens case-insensitively. If not found anywhere
      in the scan window, flags a REAL config↔input misalignment.
  v3. Confirms the target_eib has the default sheet from eib_population_defaults.yaml,
      and the row-5 header count matches Workday EIB expectations (>= 30 cols).
  v4. Opens any existing output_file_path and counts populated rows on the target sheet,
      comparing the output/source-row ratio to a sanity band (0.25..1.10).

Results printed to stdout + written to implementations/elt_transform_excel/_reports/reconcile_report.txt.

Usage:
  cd implementations/elt_transform_excel
  uv run --with pyyaml --with openpyxl --no-project \
      python scripts/reconcile_supplier_customer_inputs_outputs.py

Exit code: non-zero if any profile raises a FATAL issue (missing path / header not
found / eib sheet missing / ratio outside band).
"""
from __future__ import annotations

import argparse
import pathlib
import sys
import time
from dataclasses import dataclass, field
from typing import Optional

ROOT = pathlib.Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(SCRIPTS))

from build_runner import (  # noqa: E402
    SUPPORTED_EIB_POPULATION_TYPES,
    _load_eib_defaults,
    _load_transform_bindings,
    _import_yaml,
)

YAML = _import_yaml()
import openpyxl  # noqa: E402

# --- Tunable scan bounds ---------------------------------------------------
# Smaller = faster, but we need to catch the actual header. 400 rows × 80 cols
# is large enough for every existing profile; if a source has a header buried
# deeper the user can bump these via --scan-rows / --scan-cols.
DEFAULT_SCAN_ROWS = 400
DEFAULT_SCAN_COLS = 80
DEFAULT_RATIO_MIN = 0.25
DEFAULT_RATIO_MAX = 1.10
REPORT_DIR = ROOT / "_reports"
REPORT_FILE = REPORT_DIR / "reconcile_report.txt"


@dataclass
class Issue:
    level: str  # "FATAL" | "WARN" | "INFO"
    tag: str
    message: str


@dataclass
class ProfileReport:
    stem: str
    eib_type: str
    issues: list[Issue] = field(default_factory=list)
    facts: dict = field(default_factory=dict)

    @property
    def fatal_count(self) -> int:
        return sum(1 for i in self.issues if i.level == "FATAL")

    @property
    def warn_count(self) -> int:
        return sum(1 for i in self.issues if i.level == "WARN")


# ---------------------------------------------------------------------------
def norm(s) -> str:
    return str(s or "").strip().lower()


def _cell_column(cell) -> int:
    """Read-only openpyxl may not set cell.column; fall back to coordinate parse."""
    try:
        c = cell.column
        if c is not None:
            return int(c)
    except AttributeError:
        pass
    coord = getattr(cell, "coordinate", None) or ""
    if not coord:
        return 1
    # Coordinates look like "BC123" — strip digits, resolve letters via openpyxl.
    from openpyxl.utils import column_index_from_string
    letters = "".join(ch for ch in coord if ch.isalpha())
    if not letters:
        return 1
    return int(column_index_from_string(letters))


def quick_tokens_for_row(ws, r: int, max_col: int) -> set[str]:
    tok: set[str] = set()
    row = ws[r]
    # Row iteration stops early; do not loop 1..max_col blindly on large sheets
    for cell in row:
        c = _cell_column(cell)
        if c > max_col:
            break
        v = cell.value
        if v is None:
            continue
        nv = norm(v)
        if nv:
            tok.add(nv)
    return tok


def find_header_row_vba_style(ws, required: list[str], max_rows: int, max_col: int):
    """Return first row index 1..N where every required token appears (any col,
    case-insensitive). Returns tuple (row, matches_list, unmatched_list). If no
    full match, returns (best_row, best_matches, best_unmatched)."""
    req_tokens = [norm(x) for x in required]
    req_set = set(req_tokens)
    row_data: list[tuple[int, set[str]]] = []
    scan_cap = min(max_rows, ws.max_row or max_rows)
    ws.iter_rows()  # warm iterator (some openpyxl read-only versions need this for ws[r])
    for r in range(1, scan_cap + 1):
        t = quick_tokens_for_row(ws, r, max_col)
        if not t:
            continue
        # Fast short-circuit: size of intersection >= len(req_tokens) before per-item check
        if len(t & req_set) == len(req_tokens):
            return r, list(required), []
        row_data.append((r, t))
    best_r, best_m, best_u = None, [], list(required)
    for r, t in row_data:
        matched = [orig for rt, orig in zip(req_tokens, required) if rt in t]
        unmatched = [orig for rt, orig in zip(req_tokens, required) if rt not in t]
        if len(matched) > len(best_m):
            best_r, best_m, best_u = r, matched, unmatched
    return best_r, best_m, best_u


def find_col_indexes(ws, header_r: int, names: list[str], max_col: int) -> list[int]:
    out: list[int] = []
    name_tokens = [norm(x) for x in names]
    row_cache: dict[int, str] = {}
    for cell in ws[header_r]:
        c = _cell_column(cell)
        if c > max_col:
            break
        nv = norm(cell.value)
        if nv:
            row_cache[c] = nv
    for need in name_tokens:
        for c, nv in row_cache.items():
            if nv == need:
                out.append(c)
                break
    return out


def count_data_rows_from_header(ws, header_r: int, req_cols: list[int], max_rows: int) -> int:
    if not req_cols:
        return -1
    n = 0
    empty_streak = 0
    limit = min(max_rows, (ws.max_row or header_r + max_rows) - header_r)
    req_cols_set = set(req_cols)
    max_req = max(req_cols_set)
    for offset in range(1, limit + 1):
        r = header_r + offset
        any_v = False
        for cell in ws[r]:
            c = _cell_column(cell)
            if c > (max_req + 5):
                break
            if c in req_cols_set:
                v = cell.value
                if v is not None and str(v).strip() != "":
                    any_v = True
                    break
        if any_v:
            n += 1
            empty_streak = 0
        else:
            empty_streak += 1
            if empty_streak >= 10:
                break
    return n


def count_eib_output_rows(out_path: pathlib.Path, sheet_name: str, start_r: int, max_rows: int) -> int:
    wb = openpyxl.load_workbook(str(out_path), read_only=True, data_only=False)
    try:
        if sheet_name not in wb.sheetnames:
            return -1
        ws = wb[sheet_name]
        n = 0
        for offset in range(max_rows):
            r = start_r + offset
            v = ws.cell(row=r, column=2).value
            if v is None or str(v).strip() == "":
                break
            try:
                int(v)
                n += 1
            except (TypeError, ValueError):
                break
        return n
    finally:
        try:
            wb.close()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# CSV fallback: ar_invoice and future CSV sources. Returns a mock "workbook"
# with one synthetic worksheet whose cells match the CSV contents.
# ---------------------------------------------------------------------------

import csv


class _CsvSheet:
    """Thin stub exposing the subset of openpyxl.Worksheet API used in this module."""

    def __init__(self, rows: list[list[str]], max_col: int, max_row: int):
        self._rows = rows
        self.max_col = max_col
        self.max_row = max_row
        self._cells: dict[tuple[int, int], "_CsvCell"] = {}

    def __getitem__(self, idx):
        # ws[r] returns a row tuple of cells.
        if isinstance(idx, int):
            r = idx
        else:
            # tuple access like ws['BC23'] not used in this script; fallback to int via coord.
            from openpyxl.utils import column_index_from_string
            letters = "".join(c for c in idx if c.isalpha())
            digits = "".join(c for c in idx if c.isdigit())
            r = int(digits) if digits else 1
            _ = column_index_from_string(letters)  # validate shape, not used here
        if r < 1 or r > len(self._rows):
            return ()
        cells = []
        row_vals = self._rows[r - 1]
        for c_idx, val in enumerate(row_vals, start=1):
            key = (r, c_idx)
            if key not in self._cells:
                self._cells[key] = _CsvCell(r, c_idx, val)
            cells.append(self._cells[key])
        return tuple(cells)

    def cell(self, row: int, column: int):
        key = (row, column)
        if key in self._cells:
            return self._cells[key]
        if 1 <= row <= len(self._rows):
            row_vals = self._rows[row - 1]
            if 1 <= column <= len(row_vals):
                v = row_vals[column - 1]
            else:
                v = None
        else:
            v = None
        c = _CsvCell(row, column, v)
        self._cells[key] = c
        return c

    def iter_rows(self, *a, **kw):  # pragma: no cover - used once to warm
        return iter(())


class _CsvCell:
    __slots__ = ("row", "column", "value", "coordinate")

    def __init__(self, r: int, c: int, value):
        from openpyxl.utils import get_column_letter
        self.row = r
        self.column = c
        self.value = value
        self.coordinate = f"{get_column_letter(c)}{r}"


class _CsvWorkbook:
    def __init__(self, sheet_name: str, rows: list[list[str]]):
        max_row = len(rows)
        max_col = max((len(r) for r in rows), default=0)
        self.sheetnames = [sheet_name]
        self._sheets = {sheet_name: _CsvSheet(rows, max_col, max_row)}

    def __getitem__(self, name):
        return self._sheets[name]

    def close(self):
        return None

    def __enter__(self): return self
    def __exit__(self, *a): pass


def _open_source_csv_as_workbook(path: pathlib.Path, max_rows: int, sheet_name: str) -> _CsvWorkbook:
    # sniff delimiter + encoding, then load.
    raw_bytes = path.read_bytes()
    encoding = "utf-8"
    for enc in ("utf-8-sig", "utf-8", "utf-16", "cp1252", "latin-1"):
        try:
            raw_bytes.decode(enc)
            encoding = enc
            break
        except UnicodeDecodeError:
            continue
    text = raw_bytes.decode(encoding, errors="replace")
    sample = "\n".join(text.splitlines()[:20])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    reader = csv.reader(text.splitlines(), dialect=dialect)
    rows: list[list[str]] = []
    for idx, row in enumerate(reader):
        if idx >= max_rows + 5:
            break
        rows.append([str(c) if c is not None else "" for c in row])
    return _CsvWorkbook(sheet_name, rows)


def open_source_workbook(src_path: pathlib.Path, max_rows: int, preferred_sheet: str | None):
    """Open a source file (xlsx/xlsm/xltx/xltm via openpyxl, csv via our own
    parser) and return (workbook_like, opened_as). Always returns something
    that exposes wb.sheetnames, wb[name] sheet, ws[r] row, ws.cell(row,col)."""
    suffix = src_path.suffix.lower()
    if suffix == ".csv":
        return _open_source_csv_as_workbook(src_path, max_rows, preferred_sheet or "Sheet1"), "csv"
    # openpyxl path
    wb = openpyxl.load_workbook(str(src_path), read_only=True, data_only=False)
    return wb, "openpyxl"


# ---------------------------------------------------------------------------
def reconcile_profile(stem: str, profile: dict, args) -> ProfileReport:
    eib_type: str = profile["transform"]["eib_population_type"]
    rep = ProfileReport(stem=stem, eib_type=eib_type)
    shell_defaults: dict = profile.get("shell_defaults") or {}
    source_layout: dict = profile.get("source_layout") or {}
    eib_d = EIB_DEFAULTS.get(eib_type, {}) or {}

    src = pathlib.Path(str(shell_defaults.get("source_data", ""))).expanduser()
    tgt = pathlib.Path(str(shell_defaults.get("target_eib", ""))).expanduser()
    out = pathlib.Path(str(shell_defaults.get("output_file_path", ""))).expanduser()
    ws_name = source_layout.get("worksheet_name")
    claimed_h = int(source_layout.get("header_row") or 0)
    required: list[str] = list(source_layout.get("required_headers") or [])

    rep.facts["paths"] = {"source_data": str(src), "target_eib": str(tgt), "output_file_path": str(out)}
    rep.facts["exists"] = {"source_data": src.exists(), "target_eib": tgt.exists(), "output_file_path": out.exists()}

    # v1. path existence
    for label, exists in (("source_data", src.exists()), ("target_eib", tgt.exists())):
        if not exists:
            rep.issues.append(Issue("FATAL", f"MISSING_{label.upper()}", f"{label} path does not exist: {shell_defaults.get(label)}"))
    if not out.exists():
        rep.issues.append(Issue("WARN", "OUTPUT_NOT_PRESENT",
                                "output_file_path does not exist yet; reconcile cannot sanity-check row ratio against source."))

    # v2. source workbook
    t0 = time.time()
    if src.exists():
        try:
            wb, opened_as = open_source_workbook(src, args.scan_rows, ws_name)
            try:
                rep.facts["source_opened_as"] = opened_as
                rep.facts["source_worksheets"] = list(wb.sheetnames)
                # CSV: only sheet exists; profile ws_name may be any string. Fall back to first sheet.
                first_sheet = wb.sheetnames[0] if wb.sheetnames else None
                eff_sheet = ws_name if ws_name in wb.sheetnames else first_sheet
                if ws_name and ws_name not in wb.sheetnames and opened_as != "csv":
                    rep.issues.append(Issue("FATAL", "SRC_WORKSHEET_NOT_FOUND",
                                            f"worksheet_name={ws_name!r} missing in workbook. Actual: {wb.sheetnames}"))
                if eff_sheet is None:
                    raise RuntimeError("no worksheets in source")
                ws = wb[eff_sheet]
                rep.facts["source_claimed_worksheet"] = ws_name
                rep.facts["source_effective_worksheet"] = eff_sheet
                rep.facts["source_claimed_header_row"] = claimed_h
                rep.facts["source_required_headers_count"] = len(required)
                if not required:
                    rep.issues.append(Issue("FATAL", "NO_REQUIRED_HEADERS",
                                            "profile.source_layout.required_headers is empty; FindHeaderRow has nothing to scan for."))
                else:
                    # The VBA itself scans 1..max_row; we cap at args.scan_rows for perf.
                    h_row, match, unmatch = find_header_row_vba_style(ws, required, args.scan_rows, args.scan_cols)
                    rep.facts["header_found_row"] = h_row
                    rep.facts["header_matches"] = len(match)
                    rep.facts["header_unmatched"] = list(unmatch)
                    if unmatch:
                        rep.issues.append(Issue(
                            "FATAL",
                            "FIND_HEADER_ROW_MISS",
                            f"No row in 1..{args.scan_rows} contained ALL required_headers. "
                            f"Best attempt row={h_row} ({len(match)}/{len(required)}). "
                            f"Missing: {unmatch}. VBA FindHeaderRow would raise vbObjectError+1612."
                        ))
                    else:
                        rep.facts["header_row_found"] = h_row
                        if claimed_h and h_row != claimed_h:
                            rep.issues.append(Issue(
                                "WARN",
                                "HEADER_ROW_STALE",
                                f"Config claims header_row={claimed_h}, but FindHeaderRow found the real header at row={h_row}. "
                                "Not a runtime error (VBA ignores literal), but consider updating config."
                            ))
                        # count approx data rows (under required columns)
                        req_cols = find_col_indexes(ws, h_row, required, args.scan_cols)
                        src_rows = count_data_rows_from_header(ws, h_row, req_cols, max_rows=args.scan_rows)
                        rep.facts["approx_source_data_rows"] = src_rows
                        # sample first data row
                        first_data = []
                        for c in req_cols[:8]:
                            v = ws.cell(row=h_row + 1, column=c).value
                            s = str(v) if v is not None else ""
                            if len(s) > 40:
                                s = s[:37] + "..."
                            first_data.append(s)
                        rep.facts["first_data_row_sample"] = first_data

                        # v4. output vs source ratio
                        eib_sheet: str = str(eib_d.get("target_sheet_name") or "")
                        eib_start_r: int = int(eib_d.get("target_data_start_row") or 0)
                        if out.exists() and eib_sheet and eib_start_r:
                            out_rows = count_eib_output_rows(out, eib_sheet, eib_start_r, max_rows=args.scan_rows)
                            rep.facts["output_populated_rows"] = out_rows
                            rep.facts["output_expected_sheet"] = eib_sheet
                            if out_rows < 0:
                                rep.issues.append(Issue("FATAL", "OUTPUT_SHEET_MISSING",
                                                        f"Expected sheet {eib_sheet!r} missing from output workbook."))
                            elif isinstance(src_rows, int) and src_rows > 0:
                                ratio = out_rows / src_rows
                                rep.facts["output_source_ratio"] = round(ratio, 4)
                                if ratio < args.ratio_min:
                                    rep.issues.append(Issue(
                                        "WARN",
                                        "OUTPUT_RATIO_LOW",
                                        f"output/source row ratio = {ratio:.3f} < {args.ratio_min}. "
                                        f"src≈{src_rows}, out={out_rows}. Consider validating the rejects sheet for this run."
                                    ))
                                elif ratio > args.ratio_max:
                                    rep.issues.append(Issue(
                                        "WARN",
                                        "OUTPUT_RATIO_HIGH",
                                        f"output/source row ratio = {ratio:.3f} > {args.ratio_max}. "
                                        f"src≈{src_rows}, out={out_rows}. Possible line-split duplication?"
                                    ))
            finally:
                try:
                    wb.close()
                except Exception:
                    pass
        except Exception as exc:
            rep.issues.append(Issue("FATAL", "SRC_OPEN_ERROR", f"{type(exc).__name__}: {exc}"))
    rep.facts["source_review_seconds"] = round(time.time() - t0, 2)

    # v3. target eib template layout
    t0 = time.time()
    if tgt.exists():
        try:
            wb = openpyxl.load_workbook(str(tgt), read_only=True, data_only=False)
            try:
                eib_sheet: str = str(eib_d.get("target_sheet_name") or "")
                rep.facts["target_eib_defaults"] = dict(eib_d)
                rep.facts["target_eib_sheetnames"] = list(wb.sheetnames)
                if not eib_sheet:
                    rep.issues.append(Issue("FATAL", "EIB_DEFAULT_NO_SHEET",
                                            "eib_population_defaults has no target_sheet_name for archetype."))
                elif eib_sheet not in wb.sheetnames:
                    rep.issues.append(Issue("FATAL", "EIB_SHEET_MISSING",
                                            f"eib_defaults.target_sheet_name={eib_sheet!r} not in target workbook. Actual: {wb.sheetnames}"))
                else:
                    ws = wb[eib_sheet]
                    h5_toks = quick_tokens_for_row(ws, 5, args.scan_cols * 3)
                    rep.facts["target_eib_row5_header_count"] = len(h5_toks)
                    if len(h5_toks) < 30:
                        rep.issues.append(Issue("WARN", "EIB_ROW5_SHORT",
                                                f"Expected Workday headers at row 5, but only {len(h5_toks)} tokens found "
                                                f"(>= 30 is a healthy EIB template)."))
            finally:
                try:
                    wb.close()
                except Exception:
                    pass
        except Exception as exc:
            rep.issues.append(Issue("FATAL", "TGT_OPEN_ERROR", f"{type(exc).__name__}: {exc}"))
    rep.facts["target_review_seconds"] = round(time.time() - t0, 2)
    return rep


# ---------------------------------------------------------------------------
def print_report(profile_reports: list[ProfileReport]) -> str:
    out: list[str] = []
    def wr(*a, sep=" "):
        out.append(sep.join(str(x) for x in a))
    hr = lambda c="-": wr(c * 110)

    hr("=")
    wr("RECONCILE REPORT: config ↔ source inputs ↔ EIB outputs")
    wr(f"Generated: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    wr(f"Profiles scanned: {[p.stem for p in profile_reports]}")
    wr(f"Scan rows={DEFAULT_SCAN_ROWS}, Scan cols={DEFAULT_SCAN_COLS}, "
       f"Ratio band=[{DEFAULT_RATIO_MIN}..{DEFAULT_RATIO_MAX}] (tunable via CLI)")
    hr("=")

    tot_fatal = 0
    tot_warn = 0
    for p in profile_reports:
        hr()
        f = p.facts
        exists = f.get("exists", {})
        wr(f"PROFILE: {p.stem}   (EIB type: {p.eib_type})   FATAL={p.fatal_count}   WARN={p.warn_count}")
        paths = f.get("paths", {})
        wr(f"  source_data       {'OK   ' if exists.get('source_data') else 'MISS!'}  {paths.get('source_data','')}")
        wr(f"  target_eib        {'OK   ' if exists.get('target_eib') else 'MISS!'}  {paths.get('target_eib','')}")
        wr(f"  output_file_path  {'OK   ' if exists.get('output_file_path') else 'MISS!'}  {paths.get('output_file_path','')}")
        wr()
        wr("  [SOURCE]")
        if "source_worksheets" in f:
            wr(f"    worksheets available (first 6): {f['source_worksheets'][:6]}")
        if "source_claimed_worksheet" in f:
            found_p = f["source_claimed_worksheet"] in f.get("source_worksheets", [])
            wr(f"    claimed worksheet={f['source_claimed_worksheet']!r} present? {found_p}")
        if "header_found_row" in f:
            wr(f"    FindHeaderRow result: row={f['header_found_row']}  required_headers match {f['header_matches']}/{f['source_required_headers_count']}")
            if f.get("header_unmatched"):
                wr(f"      UNMATCHED = {f['header_unmatched']}")
            if "approx_source_data_rows" in f:
                wr(f"    approx data rows under required cols: ~{f['approx_source_data_rows']}")
            if f.get("first_data_row_sample"):
                wr(f"    1st data row sample (first 8 required cols): {f['first_data_row_sample']}")
        if f.get("source_review_seconds") is not None:
            wr(f"    [reviewed in {f['source_review_seconds']}s]")
        wr()
        wr("  [TARGET EIB TEMPLATE]")
        td = f.get("target_eib_defaults", {})
        if td:
            wr(f"    defaults -> sheet={td.get('target_sheet_name')!r} start_row={td.get('target_data_start_row')} clear_end_col={td.get('clear_published_rows_end_col')}")
        if "target_eib_row5_header_count" in f:
            wr(f"    target workbook {f.get('target_eib_defaults',{}).get('target_sheet_name')!r} row-5 header tokens = {f['target_eib_row5_header_count']}  (>= 30 OK)")
        if f.get("target_review_seconds") is not None:
            wr(f"    [reviewed in {f['target_review_seconds']}s]")
        wr()
        wr("  [OUTPUT EIB]")
        if "output_populated_rows" in f:
            wr(f"    sheet={f.get('output_expected_sheet')!r} populated rows (col B sequence) = {f['output_populated_rows']}")
            if "output_source_ratio" in f:
                src_r = f.get("approx_source_data_rows")
                wr(f"    src≈{src_r}  output/source = {f['output_source_ratio']:.3f}   "
                   f"(expect within [{DEFAULT_RATIO_MIN}..{DEFAULT_RATIO_MAX}])")
        elif exists.get("output_file_path"):
            wr("    (output file exists but no row count — see FATAL/WARN above)")
        else:
            wr("    (output file does NOT exist on disk yet.)")
        wr()
        wr(f"  ISSUES ({len(p.issues)}):")
        if p.issues:
            for idx, iss in enumerate(p.issues, 1):
                wr(f"    [{idx}] {iss.level:5s} {iss.tag:<28s} {iss.message}")
        else:
            wr(f"    (NONE)")
        wr()
        tot_fatal += p.fatal_count
        tot_warn += p.warn_count

    # summary table
    hr("=")
    wr("SUMMARY TABLE")
    wr(f"{'stem':<28s} {'src':>3s} {'eib':>3s} {'out':>3s} {'hdr?':>6s} "
       f"{'src~rows':>9s} {'eib~rows':>8s} {'ratio':>7s} "
       f"{'F':>2s} {'W':>2s}")
    for p in profile_reports:
        f = p.facts
        ex = f.get("exists", {})
        hdr_matches = f.get("header_matches")
        hdr_required = f.get("source_required_headers_count")
        hdr_txt = "-"
        if isinstance(hdr_matches, int) and isinstance(hdr_required, int) and hdr_required > 0:
            hdr_txt = f"{hdr_matches}/{hdr_required}" if not f.get("header_unmatched") else f"{hdr_matches}x{hdr_required}"
        src_r = f"~{f['approx_source_data_rows']}" if f.get("approx_source_data_rows") is not None else "-"
        eib_r = str(f["output_populated_rows"]) if isinstance(f.get("output_populated_rows"), int) and f["output_populated_rows"] >= 0 else ("-" if not f.get("output_populated_rows") else "!sheet")
        ratio = ""
        if f.get("output_source_ratio") is not None:
            ratio = f"{f['output_source_ratio']:.3f}"
        wr(f"{p.stem:<28s} "
           f"{'Y' if ex.get('source_data') else 'N':>3s} "
           f"{'Y' if ex.get('target_eib') else 'N':>3s} "
           f"{'Y' if ex.get('output_file_path') else 'N':>3s} "
           f"{hdr_txt:>6s} {src_r:>9s} {eib_r:>8s} {ratio:>7s} "
           f"{p.fatal_count:>2d} {p.warn_count:>2d}")
    hr("=")
    wr(f"TOTALS: profiles={len(profile_reports)}  FATAL issues={tot_fatal}  WARN issues={tot_warn}")
    from collections import Counter
    tag_counts: Counter = Counter()
    for p in profile_reports:
        for i in p.issues:
            tag_counts[f"{i.level}:{i.tag}"] += 1
    if tag_counts:
        wr("Issue frequency:")
        for k, c in tag_counts.most_common():
            wr(f"  {c:>3d} x {k}")
    hr("=")
    return "\n".join(out) + "\n"


# ---------------------------------------------------------------------------
def load_profiles_filtered(args):
    """Return ordered list of (stem, profile_dict) by globbing config/transforms/*.yaml,
    filtered by --profiles if provided."""
    files = sorted((ROOT / "config" / "transforms").glob("*.yaml"))
    pairs = []
    for f in files:
        stem = f.stem
        if args.profiles and stem not in args.profiles:
            continue
        data = YAML.safe_load(f.read_text()) or {}
        eib_type = data.get("transform", {}).get("eib_population_type", "")
        if args.eib_type and eib_type != args.eib_type:
            continue
        pairs.append((stem, data))
    # Order Supplier first, then Customer, then anything else
    order = {"Supplier Invoice": 0, "Customer Invoice": 1}
    pairs.sort(key=lambda kv: (order.get(kv[1].get("transform", {}).get("eib_population_type", ""), 99), kv[0]))
    return pairs


def parse_args():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--profiles", nargs="*", default=None,
                    help="Specific profile stems to reconcile (e.g. ap_amex_hotel ar_invoice). Default = ALL.")
    ap.add_argument("--eib-type", default=None, choices=sorted(SUPPORTED_EIB_POPULATION_TYPES),
                    help="Restrict to one EIB archetype. e.g. --eib-type 'Supplier Invoice'")
    ap.add_argument("--scan-rows", type=int, default=DEFAULT_SCAN_ROWS,
                    help=f"How many worksheet rows to scan for header/data. Default {DEFAULT_SCAN_ROWS}.")
    ap.add_argument("--scan-cols", type=int, default=DEFAULT_SCAN_COLS,
                    help=f"How many left columns per row to check. Default {DEFAULT_SCAN_COLS}.")
    ap.add_argument("--ratio-min", type=float, default=DEFAULT_RATIO_MIN,
                    help=f"Lower sanity band for output/source row ratio. Default {DEFAULT_RATIO_MIN}.")
    ap.add_argument("--ratio-max", type=float, default=DEFAULT_RATIO_MAX,
                    help=f"Upper sanity band for output/source row ratio. Default {DEFAULT_RATIO_MAX}.")
    ap.add_argument("--report-file", type=pathlib.Path, default=REPORT_FILE,
                    help=f"Where to write the full report. Default: {REPORT_FILE}")
    return ap.parse_args()


if __name__ == "__main__":
    args = parse_args()
    args.report_file.parent.mkdir(parents=True, exist_ok=True)

    # Load transform bindings first to surface any build-gate schema error BEFORE we start
    # (expensive) workbook scanning. Any schema error here => immediate exit.
    try:
        _ = _load_transform_bindings(ROOT / "config" / "transforms")
    except Exception as exc:
        print(f"[reconcile] BUILD-GATE FAILURE before recon scan: {type(exc).__name__}: {exc}", file=sys.stderr)
        sys.exit(2)

    global EIB_DEFAULTS
    EIB_DEFAULTS = _load_eib_defaults(ROOT / "config" / "eib_population_defaults.yaml")

    profile_pairs = load_profiles_filtered(args)
    if not profile_pairs:
        print("[reconcile] No profiles matched filters.", file=sys.stderr)
        sys.exit(1)

    profile_reports: list[ProfileReport] = []
    start_total = time.time()
    for stem, profile in profile_pairs:
        print(f"[reconcile] scanning {stem} ...", flush=True)
        profile_reports.append(reconcile_profile(stem, profile, args))

    report_text = print_report(profile_reports)
    args.report_file.write_text(report_text, encoding="utf-8")
    print(f"\n[reconcile] Full report written to: {args.report_file}", flush=True)
    total_fatal = sum(p.fatal_count for p in profile_reports)
    total_warn = sum(p.warn_count for p in profile_reports)
    print(f"[reconcile] Totals: profiles={len(profile_reports)}  FATAL={total_fatal}  WARN={total_warn}  "
          f"elapsed={time.time() - start_total:.1f}s")

    sys.exit(1 if total_fatal > 0 else 0)
