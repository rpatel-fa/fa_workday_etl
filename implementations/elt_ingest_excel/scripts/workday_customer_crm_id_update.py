# /// script
# requires-python = ">=3.11"
# dependencies = [
#     "openpyxl>=3.1",
#     "pandas>=2.2",
# ]
# ///
"""Replace Workday customer IDs with CRM customer IDs in a clone of the Master Customer workbook.

Steps:
  1. Copy the source workbook to the output path (the source is never modified).
  2. Load the mapping workbook (a): column A -> wd_customer_id, column D -> crm_customer_id.
  3. Read the clone (b) and build a map of column H: wd_customer_id -> cells it appears in
     (data rows start at row 6; rows above are EIB headers).
  4. Iterate the mapping; for each wd_customer_id found in the column-H map, replace every
     cell with the crm_customer_id.
  5. Save the clone and print stats:
       1. replaced              - in (a) and found in (b)
       2. mapping failed        - in (a), not found in (b)
       3. not replaced          - in (b), no mapping in (a)
  6. Write a gap report workbook (default ~/Downloads/workday_crm_id_update_stats.xlsx):
       Summary                   - the three stats
       Master - no CRM mapping   - full master rows whose column-H wd_customer_id has no mapping (stat 3)
       Mapping - not in master   - full mapping rows whose wd_customer_id is not in the master (stat 2)

Idempotent: every run starts from a fresh copy of the source.

Run (from repo root):
  uv run implementations/elt_ingest_excel/scripts/workday_customer_crm_id_update.py [--list]
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.utils import get_column_letter

DOWNLOADS = Path.home() / "Downloads"
DEFAULT_SOURCE = DOWNLOADS / "Master Customer EBW CRM ID Update.xlsx"
DEFAULT_OUTPUT = DOWNLOADS / "workday_crm_id_update.xlsx"
DEFAULT_MAPPING = DOWNLOADS / "workday_crm_customer_mapped.xlsx"
DEFAULT_REPORT = DOWNLOADS / "workday_crm_id_update_stats.xlsx"
TARGET_COLUMN = "H"
DEFAULT_START_ROW = 6


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--source", type=Path, default=DEFAULT_SOURCE, help="Workbook to clone (left untouched).")
    p.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Clone to write the updated IDs into.")
    p.add_argument("--mapping", type=Path, default=DEFAULT_MAPPING, help="Mapping workbook (A=wd id, D=crm id).")
    p.add_argument("--mapping-sheet", default=0, help="Mapping sheet name or index (default: first sheet).")
    p.add_argument(
        "--sheet",
        default=None,
        help="Sheet in the clone to update (default: the active sheet). Use 'ALL' to update every sheet.",
    )
    p.add_argument(
        "--start-row", type=int, default=DEFAULT_START_ROW,
        help=f"First data row in column {TARGET_COLUMN} of the clone (default: {DEFAULT_START_ROW}).",
    )
    p.add_argument("--report", type=Path, default=DEFAULT_REPORT, help="Gap report workbook to write.")
    p.add_argument("--list", action="store_true", help="Also print the IDs behind stats 2 and 3.")
    return p.parse_args()


def norm(value: object) -> str | None:
    """Normalise an ID for matching: strip whitespace, drop a trailing '.0' from numeric cells."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip()
    return text or None


def load_mapping(path: Path, sheet: str | int) -> tuple[dict[str, str], pd.DataFrame]:
    """Return (wd -> crm mapping, full mapping sheet with an added excel_row column)."""
    sheet = int(sheet) if isinstance(sheet, str) and sheet.isdigit() else sheet
    df = pd.read_excel(path, sheet_name=sheet, dtype=object)
    df.insert(0, "excel_row", df.index + 2)  # header is row 1
    wd_col, crm_col = df.columns[1], df.columns[4]  # columns A and D of the sheet

    mapping: dict[str, str] = {}
    conflicts: list[tuple[str, str, str]] = []
    skipped = 0
    for wd_raw, crm_raw in zip(df[wd_col], df[crm_col]):
        wd, crm = norm(wd_raw), norm(crm_raw)
        if wd is None or crm is None:
            skipped += 1
            continue
        if wd in mapping and mapping[wd] != crm:
            conflicts.append((wd, mapping[wd], crm))
            continue
        mapping[wd] = crm
    df["_wd_key"] = df[wd_col].map(norm)

    print(f"(a) mapping rows read: {len(df)} | distinct usable wd_customer_ids: {len(mapping)} "
          f"| skipped (blank A or D): {skipped}")
    if conflicts:
        print(f"WARNING: {len(conflicts)} wd_customer_id(s) map to more than one crm_customer_id; first kept:")
        for wd, kept, dropped in conflicts[:20]:
            print(f"  {wd}: kept {kept}, ignored {dropped}")
    return mapping, df


def master_rows(cells: list[Cell], header_row: int) -> pd.DataFrame:
    """Full master rows for the given column-H cells, headed by the sheet's header row."""
    records = []
    for cell in cells:
        ws = cell.parent
        # Use get_column_letter(c): merged header cells (MergedCell) have no .column_letter.
        headers = [
            (norm(ws.cell(row=header_row, column=c).value) if header_row >= 1 else None) or get_column_letter(c)
            for c in range(1, ws.max_column + 1)
        ]
        rec = {"sheet": ws.title, "excel_row": cell.row, "wd_customer_id": norm(cell.value)}
        for c, h in enumerate(headers, start=1):
            key = h if h not in rec else f"{h} ({get_column_letter(c)})"
            rec[key] = ws.cell(row=cell.row, column=c).value
        records.append(rec)
    return pd.DataFrame(records)


def write_report(
    path: Path, stats: list[tuple[str, int, int | None]], missing_in_mapping: pd.DataFrame,
    missing_in_master: pd.DataFrame, source: Path, mapping: Path,
) -> None:
    summary = pd.DataFrame(
        [{"stat": s, "distinct_wd_customer_ids": n, "cells": c} for s, n, c in stats]
        + [{"stat": f"master: {source.name}"}, {"stat": f"mapping: {mapping.name}"}]
    )
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Summary", index=False)
        missing_in_mapping.to_excel(xw, sheet_name="Master - no CRM mapping", index=False)
        missing_in_master.to_excel(xw, sheet_name="Mapping - not in master", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 50)


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


def ref(cell: Cell) -> str:
    return f"'{cell.parent.title}'!{cell.coordinate}"


def main() -> int:
    args = parse_args()
    for path in (args.source, args.mapping):
        if not path.exists():
            print(f"ERROR: not found: {path}", file=sys.stderr)
            return 1
    if args.output.resolve() == args.source.resolve():
        print("ERROR: --output must differ from --source", file=sys.stderr)
        return 1

    # 1. clone
    shutil.copy2(args.source, args.output)
    print(f"Cloned {args.source.name} -> {args.output}")

    # 2. mapping (a)
    mapping, mapping_df = load_mapping(args.mapping, args.mapping_sheet)

    # 3. column-H map of the clone (b)
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
          f"{len(col_h)} distinct wd_customer_id(s)")

    # 4. iterate the mapping and replace
    replaced_ids: list[str] = []
    replaced_cells = 0
    not_found: list[str] = []
    for wd, crm in mapping.items():
        cells = col_h.get(wd)
        if not cells:
            not_found.append(wd)
            continue
        for cell in cells:
            cell.value = crm
        replaced_ids.append(wd)
        replaced_cells += len(cells)

    no_mapping = {wd: cells for wd, cells in col_h.items() if wd not in mapping}
    no_mapping_cells = sum(len(c) for c in no_mapping.values())

    # Build the gap report before saving (the unreplaced rows still hold their original values).
    stats = [
        ("1. Replaced (in mapping, found in master)", len(replaced_ids), replaced_cells),
        ("2. Mapping failed (in mapping, not in master)", len(not_found), None),
        ("3. Not replaced (in master, no mapping)", len(no_mapping), no_mapping_cells),
    ]
    missing_in_mapping = master_rows(
        sorted((c for cells in no_mapping.values() for c in cells), key=lambda c: (c.parent.title, c.row)),
        args.start_row - 1,
    )
    not_found_set = set(not_found)
    missing_in_master = (
        mapping_df[mapping_df["_wd_key"].isin(not_found_set)].drop(columns="_wd_key").reset_index(drop=True)
    )

    # 5. save the clone and the report
    wb.save(args.output)
    print(f"Saved {args.output}")
    write_report(args.report, stats, missing_in_mapping, missing_in_master, args.source, args.mapping)
    print(f"Saved {args.report}")

    sheet_names = ", ".join(ws.title for ws in sheets)
    print()
    print(f"=== Stats (column {TARGET_COLUMN}, rows {args.start_row}+, of: {sheet_names}) ===")
    print(f"1. Replaced (in a, found in b):         {len(replaced_ids)} wd_customer_id(s), {replaced_cells} cell(s)")
    print(f"2. Mapping failed (in a, not in b):     {len(not_found)} wd_customer_id(s)")
    print(f"3. Not replaced (in b, no mapping in a): {len(no_mapping)} wd_customer_id(s), {no_mapping_cells} cell(s)")

    if args.list:
        print()
        print("--- 2. In mapping (a), not found in clone (b) ---")
        for wd in sorted(not_found):
            print(f"  {wd} -> {mapping[wd]}")
        print("--- 3. In clone (b), no mapping in (a) ---")
        for wd, cells in sorted(no_mapping.items()):
            refs = [ref(c) for c in cells]
            shown = ", ".join(refs[:5]) + (f", ... (+{len(refs) - 5})" if len(refs) > 5 else "")
            print(f"  {wd}: {shown}")
    elif not_found or no_mapping:
        print("(rerun with --list to see the IDs behind stats 2 and 3)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
