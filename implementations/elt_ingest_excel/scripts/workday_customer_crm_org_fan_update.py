# /// script
# requires-python = ">=3.11"
# dependencies = [
#     "openpyxl>=3.1",
#     "pandas>=2.2",
# ]
# ///
"""Update Workday customer IDs to CRM customer IDs, then build the CRM FAN Number EIB for those customers.

Inputs (never modified):
  (a) mapping  ~/Downloads/workday_crm_customer_mapped.xlsx
      columns found by header name: crm_customer_type, crm_customer_id, crm_fan, wd_customer_id
  (b) master   ~/Downloads/Master Customer EBW CRM ID Update.xlsx
  (c) FAN      ~/Downloads/FAN_Number CRM Update.xlsx

Outputs (overwritten every run, so the process is idempotent):
  ~/Downloads/workday_crm_id_update.xlsx                clone of (b), column H (rows 6+) wd id -> crm id
  ~/Downloads/workday_customer_crm_org_fan_update.xlsx  clone of (c), sheet "FAN Number for Customer"
                                                        rows 6+: A blank | B seq | C "Create" | D crm id | E crm_fan
  ~/Downloads/workday_crm_id_update_stats.xlsx          summary + gap sheets

Part 1 - CRM ID update (all mapping rows, any crm_customer_type):
  1. replaced        - in (a), found in (b) column H
  2. mapping failed  - in (a), not in (b)
  3. not replaced    - in (b), no mapping in (a)

Part 2 - FAN sheet: only customers REPLACED in part 1 whose crm_customer_type = ACCOUNT
(a CRM ID that was not loaded to Workday would not exist there). Replaced ACCOUNT customers
with a blank crm_fan are skipped and listed in the stats file.

Run (from repo root):
  uv run implementations/elt_ingest_excel/scripts/workday_customer_crm_org_fan_update.py [--list]
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.utils import get_column_letter

DOWNLOADS = Path.home() / "Downloads"
DEFAULT_MAPPING = DOWNLOADS / "workday_crm_customer_mapped.xlsx"
DEFAULT_SOURCE = DOWNLOADS / "Master Customer EBW CRM ID Update.xlsx"
DEFAULT_OUTPUT = DOWNLOADS / "workday_crm_id_update.xlsx"
DEFAULT_FAN_SOURCE = DOWNLOADS / "FAN_Number CRM Update.xlsx"
DEFAULT_FAN_OUTPUT = DOWNLOADS / "workday_customer_crm_org_fan_update.xlsx"
DEFAULT_REPORT = DOWNLOADS / "workday_crm_id_update_stats.xlsx"
DEFAULT_FAN_SHEET = "FAN Number for Customer"
TARGET_COLUMN = "H"
DEFAULT_START_ROW = 6
FAN_COLUMNS = 5  # A..E
ACCOUNT = "ACCOUNT"

# Mapping headers (matched case-insensitively, spaces treated as underscores).
COL_TYPE, COL_CRM, COL_FAN, COL_WD = "crm_customer_type", "crm_customer_id", "crm_fan", "wd_customer_id"


@dataclass(frozen=True)
class MapRow:
    crm_customer_id: str
    crm_fan: str | None
    crm_customer_type: str | None
    excel_row: int


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--mapping", type=Path, default=DEFAULT_MAPPING, help="Mapping workbook (a).")
    p.add_argument("--mapping-sheet", default=0, help="Mapping sheet name or index (default: first sheet).")
    p.add_argument("--source", type=Path, default=DEFAULT_SOURCE, help="Master customer workbook (b), left untouched.")
    p.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Clone of (b) with the CRM IDs.")
    p.add_argument(
        "--sheet", default=None,
        help="Sheet of (b) to update (default: the active sheet). Use 'ALL' to update every sheet.",
    )
    p.add_argument(
        "--start-row", type=int, default=DEFAULT_START_ROW,
        help=f"First data row in column {TARGET_COLUMN} of (b) (default: {DEFAULT_START_ROW}).",
    )
    p.add_argument("--fan-source", type=Path, default=DEFAULT_FAN_SOURCE, help="FAN Number workbook (c), left untouched.")
    p.add_argument("--fan-output", type=Path, default=DEFAULT_FAN_OUTPUT, help="Clone of (c) with the FAN rows.")
    p.add_argument("--fan-sheet", default=DEFAULT_FAN_SHEET, help=f"Sheet of (c) to fill (default: {DEFAULT_FAN_SHEET!r}).")
    p.add_argument(
        "--fan-start-row", type=int, default=DEFAULT_START_ROW,
        help=f"First data row of the FAN sheet (default: {DEFAULT_START_ROW}).",
    )
    p.add_argument("--report", type=Path, default=DEFAULT_REPORT, help="Stats workbook to write.")
    p.add_argument("--list", action="store_true", help="Also print the IDs behind the gap stats.")
    return p.parse_args()


def norm(value: object) -> str | None:
    """Normalise an ID for matching: strip whitespace, drop a trailing '.0' from numeric cells."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip()
    return text or None


def header_key(value: object) -> str:
    return re.sub(r"[\s\-]+", "_", str(value).strip().lower())


def load_mapping(path: Path, sheet: str | int) -> tuple[dict[str, MapRow], pd.DataFrame]:
    """Return (wd id -> MapRow, full mapping sheet with an added excel_row column)."""
    sheet = int(sheet) if isinstance(sheet, str) and sheet.isdigit() else sheet
    df = pd.read_excel(path, sheet_name=sheet, dtype=object)
    by_key = {header_key(c): c for c in df.columns}
    missing = [h for h in (COL_TYPE, COL_CRM, COL_FAN, COL_WD) if h not in by_key]
    if missing:
        raise SystemExit(
            f"ERROR: mapping {path.name} is missing header(s) {missing}. "
            f"Headers found: {[str(c) for c in df.columns]}"
        )
    df.insert(0, "excel_row", df.index + 2)  # header is row 1

    mapping: dict[str, MapRow] = {}
    conflicts: list[tuple[str, str, str]] = []
    skipped = 0
    for rec in df.to_dict("records"):
        wd, crm = norm(rec[by_key[COL_WD]]), norm(rec[by_key[COL_CRM]])
        if wd is None or crm is None:
            skipped += 1
            continue
        if wd in mapping and mapping[wd].crm_customer_id != crm:
            conflicts.append((wd, mapping[wd].crm_customer_id, crm))
            continue
        mapping.setdefault(wd, MapRow(
            crm_customer_id=crm,
            crm_fan=norm(rec[by_key[COL_FAN]]),
            crm_customer_type=norm(rec[by_key[COL_TYPE]]),
            excel_row=rec["excel_row"],
        ))
    df["_wd_key"] = df[by_key[COL_WD]].map(norm)

    print(f"(a) mapping rows read: {len(df)} | distinct usable {COL_WD}s: {len(mapping)} "
          f"| skipped (blank {COL_WD} or {COL_CRM}): {skipped}")
    if conflicts:
        print(f"WARNING: {len(conflicts)} {COL_WD}(s) map to more than one {COL_CRM}; first kept:")
        for wd, kept, dropped in conflicts[:20]:
            print(f"  {wd}: kept {kept}, ignored {dropped}")
    return mapping, df


def index_column(sheets: list, start_row: int) -> dict[str, list[Cell]]:
    """Map each non-blank wd_customer_id in the target column to the cells holding it."""
    index: dict[str, list[Cell]] = {}
    for ws in sheets:
        col = ws[f"{TARGET_COLUMN}1"].column
        for (cell,) in ws.iter_rows(min_row=start_row, min_col=col, max_col=col):
            key = norm(cell.value)
            if key is not None:
                index.setdefault(key, []).append(cell)
    return index


def master_rows(cells: list[Cell], header_row: int) -> pd.DataFrame:
    """Full master rows for the given column-H cells, headed by the sheet's header row."""
    records = []
    headers_by_sheet: dict[str, list[str]] = {}
    for cell in cells:
        ws = cell.parent
        if ws.title not in headers_by_sheet:
            # Read populated cells only: ws.max_column can be inflated (e.g. 16384) by sheet formatting.
            last_col = max(
                (c for (r, c), v in ws._cells.items() if v.value is not None and r >= header_row), default=1
            )
            header_vals = {c: v.value for (r, c), v in ws._cells.items() if r == header_row}
            # get_column_letter(c): merged header cells (MergedCell) have no .column_letter.
            headers_by_sheet[ws.title] = [
                norm(header_vals.get(c)) or get_column_letter(c) for c in range(1, last_col + 1)
            ]
        rec = {"sheet": ws.title, "excel_row": cell.row, "wd_customer_id": norm(cell.value)}
        for c, h in enumerate(headers_by_sheet[ws.title], start=1):
            key = h if h not in rec else f"{h} ({get_column_letter(c)})"
            existing = ws._cells.get((cell.row, c))
            rec[key] = existing.value if existing is not None else None
        records.append(rec)
    return pd.DataFrame(records)


def write_fan_sheet(ws, start_row: int, rows: list[tuple[str, str]]) -> int:
    """Clear A..E from start_row down, then write the FAN rows. Returns non-blank cells cleared."""
    # Only visit cells that exist: iterating to ws.max_row can mean ~1M rows when the template
    # is formatted to the bottom of the sheet.
    cleared = 0
    for (r, c), cell in list(ws._cells.items()):
        if r >= start_row and c <= FAN_COLUMNS and cell.value is not None:
            cleared += 1
            cell.value = None
    for seq, (crm, fan) in enumerate(rows, start=1):
        r = start_row + seq - 1
        ws.cell(row=r, column=1).value = None
        ws.cell(row=r, column=2).value = seq
        ws.cell(row=r, column=3).value = "Create"
        ws.cell(row=r, column=4).value = crm
        ws.cell(row=r, column=5).value = fan
    return cleared


def write_report(path: Path, summary_rows: list[dict], sheets: dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        pd.DataFrame(summary_rows).to_excel(xw, sheet_name="Summary", index=False)
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name, index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for idx, col in enumerate(ws.columns, start=1):
                width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
                ws.column_dimensions[get_column_letter(idx)].width = min(max(width + 2, 10), 60)


def main() -> int:
    args = parse_args()
    for path in (args.mapping, args.source, args.fan_source):
        if not path.exists():
            print(f"ERROR: not found: {path}", file=sys.stderr)
            return 1
    for src, out in ((args.source, args.output), (args.fan_source, args.fan_output)):
        if out.resolve() == src.resolve():
            print(f"ERROR: output must differ from source: {src}", file=sys.stderr)
            return 1

    mapping, mapping_df = load_mapping(args.mapping, args.mapping_sheet)

    # ---------- Part 1: CRM ID update in a clone of the master ----------
    shutil.copy2(args.source, args.output)
    print(f"Cloned {args.source.name} -> {args.output}")
    wb = load_workbook(args.output)
    if args.sheet is None:
        sheets = [wb.active]
    elif args.sheet.upper() == "ALL":
        sheets = list(wb.worksheets)
    else:
        sheets = [wb[args.sheet]]
    col_h = index_column(sheets, args.start_row)
    col_h_cells = sum(len(c) for c in col_h.values())
    print(f"(b) column {TARGET_COLUMN} from row {args.start_row}: {col_h_cells} non-blank cell(s), "
          f"{len(col_h)} distinct {COL_WD}(s)")

    # Capture unmatched master rows before replacing anything.
    no_mapping = {wd: cells for wd, cells in col_h.items() if wd not in mapping}
    no_mapping_cells = sum(len(c) for c in no_mapping.values())
    missing_in_mapping = master_rows(
        sorted((c for cells in no_mapping.values() for c in cells), key=lambda c: (c.parent.title, c.row)),
        args.start_row - 1,
    )

    replaced: list[tuple[str, Cell]] = []  # (wd id, first cell) in master order
    replaced_cells = 0
    not_found: list[str] = []
    for wd, m in mapping.items():
        cells = col_h.get(wd)
        if not cells:
            not_found.append(wd)
            continue
        for cell in cells:
            cell.value = m.crm_customer_id
        replaced.append((wd, cells[0]))
        replaced_cells += len(cells)
    replaced.sort(key=lambda t: (t[1].parent.title, t[1].row))

    wb.save(args.output)
    print(f"Saved {args.output}")

    # ---------- Part 2: FAN sheet for replaced ACCOUNT customers ----------
    fan_rows: list[tuple[str, str]] = []
    fan_blank: list[dict] = []
    fan_dupes: list[dict] = []
    non_account = 0
    seen_crm: dict[str, str | None] = {}
    for wd, cell in replaced:
        m = mapping[wd]
        if (m.crm_customer_type or "").upper() != ACCOUNT:
            non_account += 1
            continue
        base = {"wd_customer_id": wd, "crm_customer_id": m.crm_customer_id, "crm_fan": m.crm_fan,
                "mapping_excel_row": m.excel_row, "master_cell": f"'{cell.parent.title}'!{cell.coordinate}"}
        if m.crm_fan is None:
            fan_blank.append(base)
            continue
        if m.crm_customer_id in seen_crm:  # several wd ids -> same crm id: one FAN row only
            fan_dupes.append({**base, "fan_already_written": seen_crm[m.crm_customer_id]})
            continue
        seen_crm[m.crm_customer_id] = m.crm_fan
        fan_rows.append((m.crm_customer_id, m.crm_fan))

    shutil.copy2(args.fan_source, args.fan_output)
    print(f"Cloned {args.fan_source.name} -> {args.fan_output}")
    fan_wb = load_workbook(args.fan_output)
    if args.fan_sheet not in fan_wb.sheetnames:
        print(f"ERROR: sheet {args.fan_sheet!r} not in {args.fan_source.name}; sheets: {fan_wb.sheetnames}",
              file=sys.stderr)
        return 1
    cleared = write_fan_sheet(fan_wb[args.fan_sheet], args.fan_start_row, fan_rows)
    fan_wb.save(args.fan_output)
    print(f"Saved {args.fan_output}" + (f" (cleared {cleared} pre-existing cell(s) in A:E from row "
                                         f"{args.fan_start_row})" if cleared else ""))

    # ---------- Stats ----------
    summary = [
        ("1. Replaced (in mapping, found in master)", len(replaced), replaced_cells),
        ("2. Mapping failed (in mapping, not in master)", len(not_found), None),
        ("3. Not replaced (in master, no mapping)", len(no_mapping), no_mapping_cells),
        (f"4. FAN rows written (replaced + {ACCOUNT})", len(fan_rows), None),
        (f"5. FAN skipped: replaced but not {ACCOUNT}", non_account, None),
        (f"6. FAN skipped: replaced {ACCOUNT}, blank crm_fan", len(fan_blank), None),
        (f"7. FAN skipped: duplicate crm_customer_id", len(fan_dupes), None),
    ]
    missing_in_master = (
        mapping_df[mapping_df["_wd_key"].isin(set(not_found))].drop(columns="_wd_key").reset_index(drop=True)
    )
    write_report(
        args.report,
        [{"stat": s, "distinct_wd_customer_ids": n, "cells": c} for s, n, c in summary]
        + [{"stat": f"mapping: {args.mapping.name}"}, {"stat": f"master: {args.source.name}"},
           {"stat": f"fan: {args.fan_source.name}"}],
        {
            "Master - no CRM mapping": missing_in_mapping,
            "Mapping - not in master": missing_in_master,
            "FAN - blank crm_fan": pd.DataFrame(fan_blank),
            "FAN - duplicate crm id": pd.DataFrame(fan_dupes),
        },
    )
    print(f"Saved {args.report}")

    print()
    print(f"=== Stats (master column {TARGET_COLUMN}, rows {args.start_row}+, "
          f"of: {', '.join(ws.title for ws in sheets)}) ===")
    for s, n, c in summary:
        print(f"{s:<52} {n}" + (f" ({c} cell(s))" if c is not None else ""))
    print(f"Check: 1 + 3 = {len(replaced) + len(no_mapping)} (master ids {len(col_h)}); "
          f"1 + 2 = {len(replaced) + len(not_found)} (mapping ids {len(mapping)}); "
          f"4 + 5 + 6 + 7 = {len(fan_rows) + non_account + len(fan_blank) + len(fan_dupes)} (replaced {len(replaced)})")

    if args.list:
        print()
        print("--- 2. In mapping, not found in master ---")
        for wd in sorted(not_found):
            print(f"  {wd} -> {mapping[wd].crm_customer_id}")
        print("--- 3. In master, no mapping ---")
        for wd, cells in sorted(no_mapping.items()):
            refs = [f"'{c.parent.title}'!{c.coordinate}" for c in cells]
            print(f"  {wd}: " + ", ".join(refs[:5]) + (f", ... (+{len(refs) - 5})" if len(refs) > 5 else ""))
        print(f"--- 6. Replaced {ACCOUNT}, blank crm_fan ---")
        for r in fan_blank:
            print(f"  {r['wd_customer_id']} -> {r['crm_customer_id']}")
    else:
        print("(rerun with --list to see the IDs behind the gap stats; all are in the stats file)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
