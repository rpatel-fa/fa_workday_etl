#!/usr/bin/env python3
"""Inspect a Runner.xlsm workbook's Transform Registry and sheet visibility.

Usage:
    uv run --with openpyxl --no-project python \\
        elt_transform_excel/scripts/inspect_registry.py [path/to/Runner.xlsm]

Exit codes:
    0   OK.
    1   Usage error / workbook not found / could not parse.
    2   Registry sheet missing or not very-hidden (Trap #5).
    3   Registry rows don't match expected binding count / shape.
"""

from __future__ import annotations

import argparse
import sys
import zipfile
import xml.etree.ElementTree as ElementTree
from dataclasses import dataclass
from pathlib import Path

SCRIPT_PATH = Path(__file__).resolve()
MODULE_ROOT = SCRIPT_PATH.parents[1]
DEFAULT_WORKBOOK = MODULE_ROOT / "dist" / "Runner.xlsm"

REGISTRY_SHEET_NAME = "registry"
RUNNER_SHEET_NAME = "Runner"
CONFIG_SHEET_NAME = "config"
MAPPING_SHEET_NAME = "mapping"

WORKBOOK_XML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

REGISTRY_EXPECTED_COLS = 14
REGISTRY_TITLE_ROW = 1
REGISTRY_HEADER_ROW = 2
REGISTRY_START_ROW = 3
REGISTRY_HEADER = (
    "Transform Label",
    "Profile Key",
    "Source Type",
    "Source Data",
    "Target EIB",
    "Output File Path",
    "Mapping Spreadsheet",
    "Mapping Sheet",
    "Customer Spreadsheet",
    "Supplier Spreadsheet",
    "Workday Reference IDs",
    "EIB Population Type",
    "Step List Macro",
    "Start Row",
)

SHEET_STATE_VERY_HIDDEN = "veryHidden"
SHEET_STATE_HIDDEN = "hidden"
SHEET_STATE_VISIBLE = "visible"


@dataclass
class SheetInfo:
    name: str
    state: str  # one of the SHEET_STATE_* constants
    workbook_order: int


def create_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Dump the Transform Registry rows, sheet order, and sheet_state"
        " (especially 'veryHidden') from any Runner.xlsm workbook."
    )
    parser.add_argument(
        "workbook",
        nargs="?",
        type=Path,
        default=DEFAULT_WORKBOOK,
        help=f"Path to Runner.xlsm (default: {DEFAULT_WORKBOOK})",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit 2 if 'registry' sheet is not veryHidden, exit 3 if the registry "
        "rows don't match the expected build output shape.",
    )
    return parser


def _import_openpyxl():
    try:
        from openpyxl import load_workbook
    except ImportError as exc:  # pragma: no cover
        raise SystemExit(
            "openpyxl is required. Run with:\n"
            "uv run --with openpyxl --no-project python "
            "elt_transform_excel/scripts/inspect_registry.py"
        ) from exc
    return load_workbook


def _list_sheet_info(workbook_path: Path) -> list[SheetInfo]:
    """Return ordered sheet info, reading state directly from workbook.xml so we
    can detect the 'veryHidden' state that openpyxl does not always round-trip
    faithfully for VBA-protected workbooks."""
    with zipfile.ZipFile(workbook_path) as workbook_zip:
        workbook_root = ElementTree.fromstring(workbook_zip.read("xl/workbook.xml"))
    namespace = {"main": WORKBOOK_XML_NS}
    out: list[SheetInfo] = []
    for idx, sheet in enumerate(
        workbook_root.findall("main:sheets/main:sheet", namespace), 1
    ):
        name = sheet.attrib.get("name", f"<sheet{idx}>")
        state = sheet.attrib.get("state", SHEET_STATE_VISIBLE)
        if state not in (SHEET_STATE_VERY_HIDDEN, SHEET_STATE_HIDDEN, SHEET_STATE_VISIBLE):
            state = SHEET_STATE_VISIBLE
        out.append(SheetInfo(name=name, state=state, workbook_order=idx))
    return out


def _cell_text(cell, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(
            node.text or ""
            for node in cell.findall(f".//{{{WORKBOOK_XML_NS}}}t")
        )
    value_node = cell.find(f"{{{WORKBOOK_XML_NS}}}v")
    if value_node is None or value_node.text is None:
        return ""
    if cell_type == "s":
        idx = int(value_node.text)
        return shared_strings[idx] if 0 <= idx < len(shared_strings) else ""
    return value_node.text


def _read_shared_strings(workbook_zip: zipfile.ZipFile) -> list[str]:
    try:
        shared_strings_xml = workbook_zip.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ElementTree.fromstring(shared_strings_xml)
    namespace = {"main": WORKBOOK_XML_NS}
    values: list[str] = []
    for item in root.findall("main:si", namespace):
        text_parts = [node.text or "" for node in item.findall(".//main:t", namespace)]
        values.append("".join(text_parts))
    return values


def _resolve_named_sheet_path(workbook_zip: zipfile.ZipFile, sheet_name: str) -> str:
    workbook_root = ElementTree.fromstring(workbook_zip.read("xl/workbook.xml"))
    rel_root = ElementTree.fromstring(workbook_zip.read("xl/_rels/workbook.xml.rels"))
    namespace = {"main": WORKBOOK_XML_NS, "rel": PACKAGE_REL_NS}
    target_sheet = None
    for sheet in workbook_root.findall("main:sheets/main:sheet", namespace):
        if sheet.attrib.get("name") == sheet_name:
            target_sheet = sheet
            break
    if target_sheet is None:
        sheet_names = [
            s.attrib.get("name")
            for s in workbook_root.findall("main:sheets/main:sheet", namespace)
        ]
        raise RuntimeError(
            f"Workbook does not contain a worksheet named {sheet_name!r}. "
            f"Sheets present: {sheet_names!r}"
        )
    relationship_id = target_sheet.attrib.get(f"{{{REL_NS}}}id")
    if not relationship_id:
        raise RuntimeError(f"Worksheet {sheet_name!r} relationship id not found")
    for relationship in rel_root.findall("rel:Relationship", namespace):
        if relationship.attrib.get("Id") == relationship_id:
            target = relationship.attrib["Target"].lstrip("/")
            if target.startswith("xl/"):
                return target
            return f"xl/{target}"
    raise RuntimeError(
        f"Worksheet target not found for relationship id {relationship_id}"
    )


def _parse_registry_rows(workbook_path: Path) -> list[list[str]]:
    """Return [title_row, header_row, data_row1, data_row2, ...] from registry sheet."""
    with zipfile.ZipFile(workbook_path) as workbook_zip:
        shared_strings = _read_shared_strings(workbook_zip)
        sheet_path = _resolve_named_sheet_path(workbook_zip, REGISTRY_SHEET_NAME)
        worksheet_root = ElementTree.fromstring(workbook_zip.read(sheet_path))
    namespace = {"main": WORKBOOK_XML_NS}
    cells_by_ref: dict[str, str] = {}
    for cell in worksheet_root.findall(".//main:sheetData/main:row/main:c", namespace):
        cell_ref = cell.attrib.get("r")
        if not cell_ref:
            continue
        cells_by_ref[cell_ref] = _cell_text(cell, shared_strings).strip()

    rows: list[list[str]] = []
    for row_idx in (REGISTRY_TITLE_ROW, REGISTRY_HEADER_ROW, *range(REGISTRY_START_ROW, REGISTRY_START_ROW + 40)):
        row_values: list[str] = []
        for col_idx in range(1, REGISTRY_EXPECTED_COLS + 1):
            ref = f"{_col_letters(col_idx)}{row_idx}"
            row_values.append(cells_by_ref.get(ref, ""))
        if row_idx == REGISTRY_TITLE_ROW or row_idx == REGISTRY_HEADER_ROW:
            rows.append(row_values)
            continue
        if not any(v != "" for v in row_values):
            break
        rows.append(row_values)
    return rows


def _col_letters(column_index: int) -> str:
    letters = ""
    while column_index > 0:
        column_index, remainder = divmod(column_index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def _print_table(rows: list[list[str]]) -> None:
    if not rows:
        print("(no rows)")
        return
    widths = [max(len(str(r[i])) for r in rows) for i in range(len(rows[0]))]
    sep = "  ".join("-" * w for w in widths)
    for idx, row in enumerate(rows):
        print(
            "  ".join(f"{str(val):<{w}}" for val, w in zip(row, widths))
        )
        if idx == 1:  # print separator after header
            print(sep)


def main(argv: list[str] | None = None) -> int:
    parser = create_parser()
    args = parser.parse_args(argv)

    workbook_path: Path = args.workbook
    if not workbook_path.exists():
        print(f"ERROR: workbook not found: {workbook_path}", file=sys.stderr)
        return 1

    print(f"=== Workbook: {workbook_path}")
    print()

    sheet_info = _list_sheet_info(workbook_path)
    print(f"=== Sheet order + visibility ({len(sheet_info)} sheets)")
    for s in sheet_info:
        marker = ""
        if s.name == REGISTRY_SHEET_NAME:
            marker = " <-- Transform Registry (expected veryHidden)"
        print(f"  [{s.workbook_order:2d}] {s.name:<25s} state={s.state}{marker}")
    print()

    registry_info = next(
        (s for s in sheet_info if s.name == REGISTRY_SHEET_NAME), None
    )
    if registry_info is None:
        print(
            "ERROR: No sheet named 'registry' in this workbook. "
            "Either this is a pre-registry Runner, or build didn't create it.",
            file=sys.stderr,
        )
        return 2 if args.strict else 0

    if registry_info.state != SHEET_STATE_VERY_HIDDEN:
        msg = (
            f"WARNING: registry sheet state is {registry_info.state!r}, "
            f"expected {SHEET_STATE_VERY_HIDDEN!r}. Finance users can see the "
            f"14-column Transform Registry directly -- rerun build_runner.py to "
            f"hide it."
        )
        if args.strict:
            print("ERROR: " + msg, file=sys.stderr)
            return 2
        print(f"! {msg}")
        print()

    try:
        rows = _parse_registry_rows(workbook_path)
    except RuntimeError as exc:
        print(f"ERROR: could not read registry rows: {exc}", file=sys.stderr)
        return 1

    print(f"=== Transform Registry ({len(rows) - 2} transform rows, "
          f"{REGISTRY_EXPECTED_COLS} columns)")
    _print_table(rows)
    print()

    if args.strict:
        # Verify header row (row 2) matches the canonical REGISTRY_HEADER exactly.
        if len(rows) >= 2:
            header = rows[1]
            if tuple(header) != REGISTRY_HEADER:
                mismatch = []
                for i, (a, b) in enumerate(zip(header, REGISTRY_HEADER), 1):
                    if a != b:
                        mismatch.append(f"col{i}: expected {b!r} got {a!r}")
                print(
                    "ERROR: Registry header row does not match build contract. "
                    + "; ".join(mismatch),
                    file=sys.stderr,
                )
                return 3
        data_rows = rows[2:]
        blanks = [i for i, r in enumerate(data_rows, REGISTRY_START_ROW) if not any(r)]
        if blanks:
            print(
                f"ERROR: Registry contains blank data rows at absolute "
                f"row(s) {blanks}. Rows should be contiguous from {REGISTRY_START_ROW}.",
                file=sys.stderr,
            )
            return 3
        # Step List Macro is non-empty for every populated data row.
        missing_macro_rows: list[int] = []
        step_list_col = REGISTRY_HEADER.index("Step List Macro")
        for i, r in enumerate(data_rows, REGISTRY_START_ROW):
            if any(r) and not r[step_list_col]:
                missing_macro_rows.append(i)
        if missing_macro_rows:
            print(
                f"ERROR: Registry rows {missing_macro_rows} have no 'Step List Macro' "
                f"(col {_col_letters(step_list_col + 1)}).",
                file=sys.stderr,
            )
            return 3
        print(f"=== Strict checks OK: registry veryHidden + {len(data_rows)} transforms shaped correctly.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
