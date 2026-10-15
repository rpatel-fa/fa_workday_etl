"""Render a finance-facing template workbook from an ingest spec.

Reads config/ingest_specs/<key>.yaml and writes the .xlsx that the requesting team populate
by hand. Deterministic: identical spec in, byte-identical workbook out (document timestamps
are pinned, not taken from the clock), so a rebuild produces no diff unless the spec changed.

Domain-free by contract: every header, value and layout decision comes from the spec. This
script must never contain a column name, a tab name, or a business value - it renders any
ingest spec, and a spec change must never require a code change here.

See docs/CONFIGURATION.md ("Ingest Specs") for the schema and docs/ops/CLI.md for the command.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import io
import json
import re
import sys
import zipfile
from pathlib import Path

import yaml
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

SCRIPT_PATH = Path(__file__).resolve()
MODULE_ROOT = SCRIPT_PATH.parents[1]
DEFAULT_SPEC_DIR = MODULE_ROOT / "config" / "ingest_specs"
# Lookups resolve by NAME inside this module — never by a caller-supplied path. That is what keeps
# elt_transform_excel unbound: no spec can point the renderer at another module's config, or at a
# workbook someone happens to have in ~/Downloads.
LOOKUP_DIR = MODULE_ROOT / "config" / "ingest_lookups"
DEFAULT_OUTPUT_DIR = MODULE_ROOT / "dist" / "ingest_templates"

LOOKUP_FILL = PatternFill("solid", fgColor="404040")

# Pinned so output is byte-stable across runs: a rebuild produces no diff unless the spec
# changed, which is what makes "edit spec -> regenerate -> look" a usable review loop.
# Two things fight this and both are handled in save_deterministic():
#   1. openpyxl's save_workbook() assigns properties.modified = utcnow() itself, at save time.
#      Setting it beforehand has no effect - it must be corrected afterwards.
#   2. .xlsx is a zip, and zip entries carry their own mtimes taken from the clock.
FIXED_TIMESTAMP = dt.datetime(2000, 1, 1, 0, 0, 0)
ZIP_TIMESTAMP = FIXED_TIMESTAMP.timetuple()[:6]
FIXED_CREATOR = "elt_transform_excel/scripts/create_ingest_template.py"
EXCEL_MAX_SHEET_TITLE = 31

CORE_PROPERTIES_ENTRY = "docProps/core.xml"
MODIFIED_PATTERN = re.compile(rb"(<dcterms:modified[^>]*>)[^<]*(</dcterms:modified>)")

HEADER_FONT = Font(bold=True, color="FFFFFF")
REQUIRED_FILL = PatternFill("solid", fgColor="1F4E78")
OPTIONAL_FILL = PatternFill("solid", fgColor="7F7F7F")
CONDITIONAL_FILL = PatternFill("solid", fgColor="BF8F00")
FIXED_FILL = PatternFill("solid", fgColor="375623")

MIN_WIDTH = 12
MAX_WIDTH = 42
COMMENT_WIDTH = 320
COMMENT_HEIGHT = 200

# Drop-downs cover the whole column, matching what the EIB does on its own picklist columns
# (e.g. Customer Name!D4:D1048576), so a user pasting 500 rows still gets validation.
PICKLIST_LAST_ROW = 1048576
EMPTY_PICKLIST_RANGE_NAME = "_picklist_empty"
FILTERED_PICKLIST_VALUE_RE = re.compile(r"^[A-Za-z0-9_.]+$")


class SpecError(Exception):
    """The spec cannot be rendered."""


def load_spec(spec_path: Path) -> dict:
    if not spec_path.is_file():
        raise SpecError(f"spec not found: {spec_path}")
    spec = yaml.safe_load(spec_path.read_text(encoding="utf-8"))
    for key in ("meta", "template", "columns"):
        if key not in spec:
            raise SpecError(f"{spec_path.name}: missing top-level '{key}'")
    if not spec["columns"]:
        raise SpecError(f"{spec_path.name}: 'columns' is empty")
    names = [c.get("name") for c in spec["columns"]]
    if None in names:
        raise SpecError(f"{spec_path.name}: every column needs a 'name'")
    duplicates = sorted({n for n in names if names.count(n) > 1})
    if duplicates:
        raise SpecError(f"{spec_path.name}: duplicate column names: {duplicates}")

    # `feeds` must be the documented list of {tab, column, field}. Rejecting the superseded
    # prose-string form here is deliberate: a spec still carrying it predates the EIB-derived
    # direction, so rendering it would produce a template built on the wrong column set - and
    # would look fine. Fail with the reason instead. See docs/CONFIGURATION.md ("Ingest Specs").
    unstructured = [
        c["name"] for c in spec["columns"]
        if c.get("feeds") is not None
        and not (isinstance(c["feeds"], list)
                 and all(isinstance(target, dict) for target in c["feeds"]))
    ]
    if unstructured:
        raise SpecError(
            f"{spec_path.name}: 'feeds' must be a list of {{tab, column, field}}, but "
            f"{len(unstructured)} column(s) use the superseded prose-string form "
            f"(e.g. {unstructured[0]!r}). This spec has not been rebuilt on the EIB-derived "
            f"direction yet - see docs/todo/archive/ar-ingest-specs.md, Item 1b."
        )
    return spec


def fill_for(column: dict) -> PatternFill:
    """Colour encodes how the column must be populated, not what it means."""
    if column.get("default") is not None:
        return FIXED_FILL
    required = column.get("required")
    if required is True:
        return REQUIRED_FILL
    if required == "conditional":
        return CONDITIONAL_FILL
    return OPTIONAL_FILL


def header_note(column: dict) -> str:
    """The spec's knowledge about one column, rendered where the user is typing."""
    required = column.get("required")
    obligation = {True: "REQUIRED", False: "Optional"}.get(required, str(required).upper())
    lines = [f"{obligation}  ({column.get('type', 'text')})"]

    if column.get("picklist"):
        picklist = column["picklist"]
        where = picklist.get("where")
        filter_note = f" where {where}" if where else ""
        lines.append(f"DROP-DOWN: pick from the '{picklist['lookup']}' tab"
                     f" ({picklist['column']}{filter_note})")
        if picklist.get("filtered_by"):
            lines.append(
                f"Filtered by: {picklist['filtered_by']} via {picklist['key']}"
            )
    if column.get("default") is not None:
        lines.append(f"Default: {column['default']} (pre-filled on the first row — change if needed)")
    if column.get("multi_value"):
        lines.append(f"Multiple: {column['multi_value']}")
    if column.get("condition"):
        lines.append(f"Required when: {column['condition']}")
    if column.get("example"):
        lines.append(f"Example: {column['example']}")
    if column.get("allowed_values"):
        lines.append("Allowed: " + ", ".join(str(v) for v in column["allowed_values"]))

    for target in column.get("feeds") or []:
        tab, col, field = target.get("tab"), target.get("column"), target.get("field")
        lines.append(f"-> {tab}!{col} ({field})" if col else f"-> {tab} ({field})")

    description = (column.get("description") or "").strip()
    if description:
        lines.append("")
        lines.append(" ".join(description.split()))
    return "\n".join(lines)


def load_lookup(name: str) -> dict:
    """Load one shared reference dataset by name from config/ingest_lookups/<name>.json.

    The file is self-describing — `description`, `key_column`, `provenance`, `rows` — so a lookup
    shared by several specs documents itself once instead of once per spec.
    """
    if "/" in name or "\\" in name or name.startswith("."):
        raise SpecError(f"lookup name must be a bare name, not a path: {name!r}")
    path = LOOKUP_DIR / f"{name}.json"
    if not path.is_file():
        available = sorted(p.stem for p in LOOKUP_DIR.glob("*.json"))
        raise SpecError(f"unknown lookup {name!r}. Available: {available or 'none'}")

    doc = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(doc, dict) or "rows" not in doc:
        raise SpecError(f"lookup {name!r} must be an object with a 'rows' list")
    rows = doc["rows"]
    if not isinstance(rows, list) or not rows:
        raise SpecError(f"lookup {name!r}: 'rows' must be a non-empty list")
    if not all(isinstance(row, dict) for row in rows):
        raise SpecError(f"lookup {name!r}: 'rows' must be a list of objects")
    return doc


def lookup_sheet_title(name: str) -> str:
    """Deterministic Excel-safe worksheet title for one lookup tab.

    Excel refuses worksheet titles longer than 31 characters. Lookup names remain the stable
    config-facing identifiers; only the rendered tab title is shortened when necessary.
    """
    if len(name) <= EXCEL_MAX_SHEET_TITLE:
        return name

    digest = hashlib.sha1(name.encode()).hexdigest()[:8]
    prefix = name[: EXCEL_MAX_SHEET_TITLE - len(digest) - 1]
    return f"{prefix}_{digest}"


def matched_lookup_rows(rows: list[dict], where: dict) -> list[int]:
    """1-based worksheet row numbers (including header offset) for rows matching `where`."""
    return [
        index for index, row in enumerate(rows, start=2)
        if all(str(row.get(key, "")) == str(value) for key, value in where.items())
    ]


def ensure_contiguous_rows(name: str, column: str, matched: list[int], where: dict) -> None:
    """Excel validations can only point at one contiguous range."""
    if not matched:
        raise SpecError(f"picklist {name}.{column} with where={where} matches no rows")
    if matched != list(range(matched[0], matched[-1] + 1)):
        raise SpecError(
            f"picklist {name}.{column} with where={where} matches non-contiguous rows {matched}. "
            f"An Excel range must be contiguous — reorder config/ingest_lookups/{name}.json so "
            f"the matching rows sit together."
        )


def filtered_picklist_prefix(picklist: dict) -> str:
    """Stable named-range prefix for one filtered picklist definition."""
    signature = json.dumps(
        {
            "lookup": picklist["lookup"],
            "column": picklist["column"],
            "key": picklist["key"],
            "where": picklist.get("where") or {},
        },
        sort_keys=True,
        separators=(",", ":"),
    )
    return f"_pl_{hashlib.sha1(signature.encode()).hexdigest()[:12]}_"


def picklist_formula(picklist: dict) -> str:
    """Resolve a column's `picklist:` to an Excel range on its lookup tab.

    Points at the rendered lookup tab rather than inlining the values: an inline list is capped
    near 255 characters, which 106 countries would blow past, and it would also duplicate values
    the workbook already carries.
    """
    name = picklist["lookup"]
    column = picklist["column"]
    rows = load_lookup(name)["rows"]

    headers = list(rows[0].keys())
    if column not in headers:
        raise SpecError(f"picklist column {column!r} not in lookup {name!r}. Has: {headers}")
    letter = get_column_letter(headers.index(column) + 1)
    sheet_title = lookup_sheet_title(name)

    where = picklist.get("where") or {}
    matched = matched_lookup_rows(rows, where)
    ensure_contiguous_rows(name, column, matched, where)
    return f"'{sheet_title}'!${letter}${matched[0]}:${letter}${matched[-1]}"


def filtered_picklist_formula(picklist: dict, controller_letter: str, data_start_row: int) -> str:
    """Resolve a filtered picklist using named ranges keyed by another template column."""
    controller_ref = f"${controller_letter}{data_start_row}"
    return (
        f'=IF({controller_ref}="",{EMPTY_PICKLIST_RANGE_NAME},'
        f'INDIRECT("{filtered_picklist_prefix(picklist)}"&{controller_ref}))'
    )


def add_picklist(
    workbook: Workbook,
    sheet,
    column: dict,
    letter: str,
    data_start_row: int,
    column_letters: dict[str, str],
) -> None:
    """Attach a list validation covering the whole column."""
    picklist = column["picklist"]
    if picklist.get("filtered_by"):
        controller_name = picklist["filtered_by"]
        controller_letter = column_letters.get(controller_name)
        if controller_letter is None:
            raise SpecError(
                f"picklist {picklist['lookup']}.{picklist['column']} filtered_by "
                f"{controller_name!r}, but no such template column exists"
            )
        formula = filtered_picklist_formula(picklist, controller_letter, data_start_row)
    else:
        formula = picklist_formula(picklist)

    validation = DataValidation(
        type="list",
        formula1=formula,
        allow_blank=True,
        showDropDown=False,  # openpyxl inverts this: False means the arrow IS shown
        showErrorMessage=True,
        errorTitle="Value not allowed",
        error=f"Pick a value from the {picklist['lookup']} tab.",
    )
    sheet.add_data_validation(validation)
    validation.add(f"{letter}{data_start_row}:{letter}{PICKLIST_LAST_ROW}")


def write_lookup_sheet(workbook: Workbook, name: str) -> int:
    """Render one reference dataset as its own worksheet. Column order follows the first row."""
    rows = load_lookup(name)["rows"]
    headers = list(rows[0].keys())

    sheet = workbook.create_sheet(title=lookup_sheet_title(name))
    for index, header in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=index, value=header)
        cell.font = HEADER_FONT
        cell.fill = LOOKUP_FILL
        width = max(MIN_WIDTH, min(MAX_WIDTH, len(str(header)) + 2))
        sheet.column_dimensions[get_column_letter(index)].width = width

    for row_index, row in enumerate(rows, start=2):
        for col_index, header in enumerate(headers, start=1):
            sheet.cell(row=row_index, column=col_index, value=row.get(header))

    sheet.freeze_panes = sheet.cell(row=2, column=1)
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"
    return len(rows)


def add_picklist_helper_sheet(workbook: Workbook) -> None:
    """Hidden sheet used for an always-blank fallback range in cascading picklists."""
    if EMPTY_PICKLIST_RANGE_NAME in workbook.defined_names:
        return
    sheet = workbook.create_sheet(title="_system")
    sheet.sheet_state = "hidden"
    sheet["A1"] = ""
    workbook.defined_names.add(
        DefinedName(EMPTY_PICKLIST_RANGE_NAME, attr_text="'_system'!$A$1")
    )


def add_filtered_picklist_names(workbook: Workbook, columns: list[dict]) -> None:
    """Create named ranges for each filtered-picklist key value."""
    filtered_columns = [column for column in columns if column.get("picklist", {}).get("filtered_by")]
    if not filtered_columns:
        return

    add_picklist_helper_sheet(workbook)
    defined_ranges: dict[str, str] = {}

    for column in filtered_columns:
        picklist = column["picklist"]
        if not picklist.get("key"):
            raise SpecError(
                f"picklist {picklist['lookup']}.{picklist['column']} uses filtered_by but has no key"
            )

        name = picklist["lookup"]
        value_column = picklist["column"]
        key_column = picklist["key"]
        rows = load_lookup(name)["rows"]
        headers = list(rows[0].keys())
        if value_column not in headers:
            raise SpecError(f"picklist column {value_column!r} not in lookup {name!r}. Has: {headers}")
        if key_column not in headers:
            raise SpecError(f"picklist key {key_column!r} not in lookup {name!r}. Has: {headers}")

        value_letter = get_column_letter(headers.index(value_column) + 1)
        where = picklist.get("where") or {}
        matched = matched_lookup_rows(rows, where)
        ensure_contiguous_rows(name, value_column, matched, where)

        groups: dict[str, list[int]] = {}
        for row_index in matched:
            key_value = str(rows[row_index - 2].get(key_column, "")).strip()
            if not key_value:
                raise SpecError(
                    f"picklist {name}.{value_column} filtered by {key_column!r} has a blank key value"
                )
            if not FILTERED_PICKLIST_VALUE_RE.match(key_value):
                raise SpecError(
                    f"picklist {name}.{value_column} filtered by {key_column!r} uses key value "
                    f"{key_value!r}, which cannot become an Excel named range. "
                    "Use values containing only letters, numbers, underscore, and period."
                )
            groups.setdefault(key_value, []).append(row_index)

        for key_value, key_rows in groups.items():
            ensure_contiguous_rows(name, value_column, key_rows, {**where, key_column: key_value})
            defined_name = f"{filtered_picklist_prefix(picklist)}{key_value}"
            sheet_title = lookup_sheet_title(name)
            attr_text = (
                f"'{sheet_title}'!${value_letter}${key_rows[0]}:${value_letter}${key_rows[-1]}"
            )
            existing = defined_ranges.get(defined_name)
            if existing and existing != attr_text:
                raise SpecError(
                    f"filtered picklist name collision for {defined_name!r}: "
                    f"{existing} vs {attr_text}"
                )
            if not existing:
                workbook.defined_names.add(DefinedName(defined_name, attr_text=attr_text))
                defined_ranges[defined_name] = attr_text


def write_instructions_sheet(workbook: Workbook, template: dict, columns: list[dict]) -> None:
    instructions = template.get("instructions") or {}
    if not instructions.get("enabled"):
        return

    sheet_name = (instructions.get("worksheet_name") or "Instructions").strip() or "Instructions"
    sheet_name = sheet_name[:EXCEL_MAX_SHEET_TITLE]
    sheet = workbook.create_sheet(title=sheet_name, index=0)

    sheet["A1"] = "Instructions"
    sheet["A1"].font = Font(bold=True, size=14)
    sheet["A3"] = "Header colours"
    sheet["A3"].font = Font(bold=True)

    legend = [
        (REQUIRED_FILL, "Required"),
        (OPTIONAL_FILL, "Optional"),
        (CONDITIONAL_FILL, "Required (conditional)"),
        (FIXED_FILL, "Has a default (pre-filled in row 2)"),
    ]
    sheet["A4"] = "Colour"
    sheet["B4"] = "Meaning"
    sheet["A4"].font = Font(bold=True, color="FFFFFF")
    sheet["B4"].font = Font(bold=True, color="FFFFFF")
    sheet["A4"].fill = LOOKUP_FILL
    sheet["B4"].fill = LOOKUP_FILL

    for idx, (fill, meaning) in enumerate(legend, start=5):
        sheet.cell(row=idx, column=1, value="").fill = fill
        sheet.cell(row=idx, column=2, value=meaning)

    start_row = 5 + len(legend) + 2
    sheet.cell(row=start_row, column=1, value="Field guide").font = Font(bold=True)
    start_row += 1

    headers = [
        "Column",
        "Required",
        "Default",
        "Drop-down",
        "Notes",
        "Description",
    ]
    for col_index, header in enumerate(headers, start=1):
        cell = sheet.cell(row=start_row, column=col_index, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = LOOKUP_FILL
        sheet.column_dimensions[get_column_letter(col_index)].width = max(MIN_WIDTH, min(MAX_WIDTH, len(header) + 2))

    row_index = start_row + 1
    for column in columns:
        required = column.get("required")
        required_text = (
            "Yes" if required is True else "Conditional" if required == "conditional" else "No"
        )
        default_value = column.get("default")

        picklist = column.get("picklist") or {}
        dropdown = ""
        if picklist:
            dropdown = f"{picklist.get('lookup')}.{picklist.get('column')}"
            if picklist.get("filtered_by"):
                dropdown = (
                    f"{dropdown} (filtered by {picklist.get('filtered_by')} via {picklist.get('key')})"
                )

        notes = []
        if column.get("multi_value"):
            notes.append(f"Multi-value: {column.get('multi_value')}")
        if column.get("condition"):
            notes.append(f"Condition: {column.get('condition')}")
        if picklist:
            notes.append("Use the drop-down (avoid free-typing).")

        values = [
            column.get("name"),
            required_text,
            "" if default_value is None else str(default_value),
            dropdown,
            " ".join(notes),
            " ".join(str(column.get("description") or "").split()),
        ]
        for col_index, value in enumerate(values, start=1):
            sheet.cell(row=row_index, column=col_index, value=value)
        row_index += 1

    sheet.freeze_panes = sheet.cell(row=start_row + 1, column=1)
    if hasattr(sheet, "auto_filter"):
        sheet.auto_filter.ref = f"A{start_row}:{get_column_letter(len(headers))}{start_row}"


def build_workbook(spec: dict, sample_row: bool) -> Workbook:
    template = spec["template"]
    columns = spec["columns"]
    header_row = int(template.get("header_row", 1))
    data_start_row = int(template.get("data_start_row", header_row + 1))

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = template.get("worksheet_name") or spec["meta"]["key"]
    column_letters: dict[str, str] = {}

    for index, column in enumerate(columns, start=1):
        letter = get_column_letter(index)
        column_letters[column["name"]] = letter
        cell = sheet.cell(row=header_row, column=index, value=column["name"])
        cell.font = HEADER_FONT
        cell.fill = fill_for(column)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.comment = Comment(header_note(column), FIXED_CREATOR,
                               width=COMMENT_WIDTH, height=COMMENT_HEIGHT)

        width = max(MIN_WIDTH, min(MAX_WIDTH, len(str(column["name"])) + 2))
        sheet.column_dimensions[letter].width = width

        # A `default` is a real value meant to be kept, so it is always pre-filled on the first
        # data row — that is what makes it a default rather than documentation. An `example` is
        # illustrative and must never ship, so it stays behind --sample-row.
        value = column.get("default")
        if value is None and sample_row:
            value = column.get("example")
        if value is not None:
            sheet.cell(row=data_start_row, column=index, value=value)

    sheet.freeze_panes = sheet.cell(row=data_start_row, column=1)
    if hasattr(sheet, "auto_filter"):
        sheet.auto_filter.ref = (
            f"A{header_row}:{get_column_letter(len(columns))}{header_row}"
        )

    write_instructions_sheet(workbook, template, columns)

    # Reference tabs, after the input sheet so it stays the one the workbook opens on.
    for name in spec.get("lookups") or []:
        write_lookup_sheet(workbook, name)
    add_filtered_picklist_names(workbook, columns)

    for column in columns:
        if column.get("picklist"):
            add_picklist(
                workbook,
                sheet,
                column,
                column_letters[column["name"]],
                data_start_row,
                column_letters,
            )

    workbook.properties.creator = FIXED_CREATOR
    workbook.properties.lastModifiedBy = FIXED_CREATOR
    workbook.properties.created = FIXED_TIMESTAMP
    workbook.properties.modified = FIXED_TIMESTAMP
    workbook.properties.title = spec["meta"].get("label", spec["meta"]["key"])
    workbook.active = workbook.sheetnames.index(sheet.title)
    return workbook


def save_deterministic(workbook: Workbook, output_path: Path) -> None:
    """Save so identical specs yield byte-identical files.

    Saves to memory first, then rewrites the zip with every entry mtime pinned and the
    save-time `modified` property overwritten. See FIXED_TIMESTAMP for why both are needed.
    """
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)

    stamp = FIXED_TIMESTAMP.strftime("%Y-%m-%dT%H:%M:%SZ").encode()
    with zipfile.ZipFile(buffer) as source:
        entries = [(item, source.read(item.filename)) for item in source.infolist()]

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as target:
        for item, data in entries:
            if item.filename == CORE_PROPERTIES_ENTRY:
                data = MODIFIED_PATTERN.sub(rb"\g<1>" + stamp + rb"\g<2>", data)
            info = zipfile.ZipInfo(item.filename, date_time=ZIP_TIMESTAMP)
            info.compress_type = item.compress_type
            info.external_attr = item.external_attr
            target.writestr(info, data)


def create_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Render a finance-facing template workbook from an ingest spec.",
    )
    parser.add_argument(
        "spec",
        help=f"Ingest spec key (resolved under {DEFAULT_SPEC_DIR}) or a path to a spec YAML.",
    )
    parser.add_argument("--output", help="Output .xlsx path. Default: dist/ingest_templates/<key>.xlsx")
    parser.add_argument("--overwrite", action="store_true", help="Replace an existing output file.")
    parser.add_argument(
        "--sample-row",
        action="store_true",
        help="Also fill the first data row from each column's `example`. Off by default: examples "
             "are illustrative and must not ship as submittable data. Columns with a `default` are "
             "pre-filled either way — a default is a real value, not an illustration.",
    )
    return parser


def resolve_spec_path(value: str) -> Path:
    candidate = Path(value).expanduser()
    if candidate.suffix and candidate.is_file():
        return candidate.resolve()
    return (DEFAULT_SPEC_DIR / f"{value}.yaml").resolve()


def main(argv: list[str] | None = None) -> int:
    args = create_parser().parse_args(argv)

    try:
        spec_path = resolve_spec_path(args.spec)
        spec = load_spec(spec_path)
    except SpecError as error:
        print(f"error: {error}", file=sys.stderr)
        return 1

    output_path = (
        Path(args.output).expanduser().resolve()
        if args.output
        else DEFAULT_OUTPUT_DIR / f"{spec['meta']['key']}.xlsx"
    )
    if output_path.exists() and not args.overwrite:
        print(f"error: {output_path} exists (use --overwrite)", file=sys.stderr)
        return 1

    try:
        workbook = build_workbook(spec, sample_row=args.sample_row)
    except SpecError as error:
        print(f"error: {error}", file=sys.stderr)
        return 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    save_deterministic(workbook, output_path)

    columns = spec["columns"]
    required = sum(1 for c in columns if c.get("required") is True)
    defaults = sum(1 for c in columns if c.get("default") is not None)
    picklists = [c["name"] for c in columns if c.get("picklist")]
    print(f"wrote {output_path}")
    print(f"  {len(columns)} columns  ({required} required, {defaults} pre-filled default, "
          f"{len(picklists)} drop-down)")
    print(f"  sheet {spec['template'].get('worksheet_name')!r}, "
          f"header row {spec['template'].get('header_row', 1)}")
    for name in spec.get("lookups") or []:
        print(f"  lookup tab {name!r}: {len(load_lookup(name)['rows'])} rows")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
