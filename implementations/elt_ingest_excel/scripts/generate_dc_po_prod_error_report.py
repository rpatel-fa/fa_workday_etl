"""Build a simple finance-facing Excel report from the DC PO production errors workbook.

The report is designed to help Finance review the current Workday load errors and
identify where mapping corrections or Workday setup actions are needed. It reads
the source error workbook, groups the issues into simple action lists, enriches
them from DuckDB mapping/reference tables, and writes a new workbook with:

- `summary`: high-level issue families and counts
- `camp_setup`: one row per camp/custom-org code missing in Workday
- `mapping_fixes`: non-camp invalid IDs that Finance can review in the mapping table
- `other_issues`: remaining non-ID validation issues grouped into simple buckets

Usage
-----
    uv run --project elt_ingest_excel python \
      elt_ingest_excel/scripts/generate_dc_po_prod_error_report.py
"""

from __future__ import annotations

import argparse
import json
import os
import re
from collections import Counter, defaultdict
from pathlib import Path

import duckdb
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DEFAULT_INPUT = os.path.expanduser("~/Downloads/DC PO Prod Errors TEMP.xlsx")
DEFAULT_OUTPUT = os.path.expanduser("~/Downloads/DC PO Prod Errors Finance Action Report.xlsx")
DEFAULT_RUNTIME_CONFIG = os.path.expanduser(
    "~/Documents/__code/git/emailrak/elt_lake/elt_ingest_excel/config/runtime/default.json"
)
ERROR_PREFIX_RE = re.compile(r"^\s*Validation error occurred\.\s*")
INVALID_ID_RE = re.compile(
    r"Invalid ID value\. '([^']+)' is not a valid ID value for type = '([^']+)'"
)
JPY_PRECISION_RE = re.compile(
    r"Decimal precision of 0 for JPY exceeded for Extended Amount: ([0-9.]+)"
)

REF_COLUMNS = {
    "spend_category_ref_id": ("Spend Category", "Spend_Category_ID"),
    "cost_centre_ref_id": ("Cost Centre", "Cost_Center_Reference_ID"),
    "event_ref_id": ("Events", "Custom_Organization_Reference_ID"),
    "camp_ref_id": ("Camps", "Custom_Organization_Reference_ID"),
    "fa_learning_course_ref_id": ("FA Learning", "Custom_Organization_Reference_ID"),
    "program_ref_id": ("Programmes", "Program_ID"),
    "grant_ref_id": ("Grant", "Grant_ID"),
    "commercial_partner_ref_id": ("Commercial Partner", "Custom_Worktag_3_ID"),
}

HEADER_FILL = PatternFill(fill_type="solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
WRAP = Alignment(wrap_text=True, vertical="top")
TOP = Alignment(vertical="top")


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        default=DEFAULT_INPUT,
        help="Source Workday error workbook.",
    )
    parser.add_argument(
        "--out",
        default=DEFAULT_OUTPUT,
        help="Output workbook path.",
    )
    parser.add_argument(
        "--runtime-config",
        default=DEFAULT_RUNTIME_CONFIG,
        help="Runtime config used to resolve the DuckDB path.",
    )
    parser.add_argument(
        "--duckdb",
        default=None,
        help="DuckDB path override. Defaults to runtime/default.json.",
    )
    return parser.parse_args(argv)


def resolve_duckdb_path(duckdb_path: str | None, runtime_config_path: str) -> str:
    if duckdb_path:
        return os.path.expanduser(duckdb_path)

    with open(os.path.expanduser(runtime_config_path), encoding="utf-8") as handle:
        config = json.load(handle)

    database_path = (
        config.get("databasePath")
        or config.get("duckdb", {}).get("path")
        or config.get("duckdb_path")
    )
    if not database_path:
        raise ValueError(
            f"Could not find a DuckDB path in runtime config: {runtime_config_path}"
        )
    return os.path.expanduser(database_path)


def normalize_message(message: str) -> str:
    return ERROR_PREFIX_RE.sub("", (message or "").strip())


def issue_family(message: str) -> tuple[str, str]:
    normalized = normalize_message(message)
    if INVALID_ID_RE.search(normalized):
        value, type_name = INVALID_ID_RE.search(normalized).groups()
        if type_name == "Custom_Organization_Reference_ID":
            return "Invalid custom org ID", f"{value} ({type_name})"
        return f"Invalid {type_name}", value
    if "Spend category must have a Spend Category Usage type of Procurement" in normalized:
        return "Spend category not procurement-enabled", normalized
    if "Resource_Category_Reference" in normalized:
        return "Missing resource category", normalized
    if 'Spend category "" is inactive' in normalized:
        return "Blank/inactive spend category", normalized
    if "Please specify an Item or a Spend Category" in normalized:
        return "Missing item or spend category", normalized
    if "The following Accounting Worktags are required but have not been assigned: Cost Center" in normalized:
        return "Missing cost centre", normalized
    if "Currency Conversion Rates between currency and company's default currency are not defined" in normalized:
        return "Missing currency conversion rate", normalized
    if "Please specify an Item or an Item Description" in normalized:
        return "Missing item description", normalized
    if JPY_PRECISION_RE.search(normalized):
        return "JPY precision issue", normalized
    return "Other validation issue", normalized


def load_error_workbook(path: str) -> dict[str, list[tuple[str, str]]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheets: dict[str, list[tuple[str, str]]] = {}
    for sheet_name in ("Receipted", "Unreceipted"):
        ws = workbook[sheet_name]
        rows: list[tuple[str, str]] = []
        for row in ws.iter_rows(values_only=True):
            row_id = "" if len(row) < 1 or row[0] is None else str(row[0]).strip()
            message = "" if len(row) < 2 or row[1] is None else str(row[1]).strip()
            if not message:
                continue
            rows.append((row_id, message))
        sheets[sheet_name] = rows

    camp_ws = workbook["Camps Missing"]
    camp_values: list[str] = []
    for row in camp_ws.iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            camp_values.append(str(row[0]).strip())
    sheets["Camps Missing"] = [(camp, "") for camp in camp_values]
    workbook.close()
    return sheets


def summarize_workbook_issues(
    sheets: dict[str, list[tuple[str, str]]]
) -> tuple[
    list[dict[str, str | int]],
    Counter,
    dict[tuple[str, str], Counter],
    dict[tuple[str, str], list[str]],
]:
    family_counts: Counter = Counter()
    invalid_counts: dict[tuple[str, str], Counter] = defaultdict(Counter)
    invalid_row_refs: dict[tuple[str, str], list[str]] = defaultdict(list)

    for scope in ("Receipted", "Unreceipted"):
        for row_id, raw_message in sheets.get(scope, []):
            normalized = normalize_message(raw_message)
            family, detail = issue_family(raw_message)
            family_counts[(family, scope)] += 1

            match = INVALID_ID_RE.search(normalized)
            if match:
                value, type_name = match.groups()
                invalid_counts[(value, type_name)][scope] += 1
                if row_id:
                    invalid_row_refs[(value, type_name)].append(f"{scope}:{row_id}")

    summary_rows: list[dict[str, str | int]] = []
    all_families = sorted({family for family, _scope in family_counts})
    for family in all_families:
        receipted = family_counts.get((family, "Receipted"), 0)
        unreceipted = family_counts.get((family, "Unreceipted"), 0)
        total = receipted + unreceipted
        action = suggested_action_for_family(family)
        summary_rows.append(
            {
                "issue_family": family,
                "receipted_rows": receipted,
                "unreceipted_rows": unreceipted,
                "total_rows": total,
                "owner": action[0],
                "action_needed": action[1],
            }
        )

    summary_rows.sort(key=lambda row: (-int(row["total_rows"]), str(row["issue_family"])))
    return summary_rows, family_counts, invalid_counts, invalid_row_refs


def suggested_action_for_family(family: str) -> tuple[str, str]:
    if family == "Invalid custom org ID":
        return ("Finance / Workday", "Confirm mapping target and set up missing custom org in Workday if required.")
    if family in {
        "Invalid Spend_Category_ID",
        "Spend category not procurement-enabled",
        "Missing resource category",
        "Blank/inactive spend category",
        "Missing item or spend category",
    }:
        return ("Finance", "Review spend-category mapping and correct invalid or unusable values.")
    if family in {"Invalid Cost_Center_Reference_ID", "Missing cost centre"}:
        return ("Finance", "Review cost-centre mapping and populate a valid Workday cost centre.")
    if family == "Invalid Custom_Worktag_3_ID":
        return ("Finance", "Review commercial partner mapping.")
    if family == "Invalid Program_ID":
        return ("Finance / Workday", "Review programme mapping or set up missing programme in Workday.")
    if family == "Invalid Grant_ID":
        return ("Finance / Workday", "Review grant mapping or set up missing grant in Workday.")
    if family == "Missing currency conversion rate":
        return ("Workday", "Create the required currency conversion rate in Workday.")
    if family == "Missing item description":
        return ("Finance", "Populate either Item or Item Description in the source/mapping.")
    if family == "JPY precision issue":
        return ("ETL", "Already fixed in current code. Re-export and retest.")
    return ("Finance", "Review the source row and current mapping value.")


def load_mapping_context(con: duckdb.DuckDBPyConnection, invalid_values: set[str]) -> dict[str, list[dict[str, str]]]:
    if not invalid_values:
        return {}

    placeholders = ", ".join(["?"] * len(invalid_values))
    rows = con.execute(
        f"""
        SELECT
             TRIM(workday_business_object) AS workday_business_object
           , TRIM(old_code)                AS old_code
           , TRIM(workday_ref_id)          AS workday_ref_id
           , TRIM(workday_name)            AS workday_name
           , TRIM(old_segment)             AS old_segment
        FROM fin_ref_wd_fdm_mapping
        WHERE TRIM(workday_ref_id) IN ({placeholders})
           OR UPPER(TRIM(old_code)) IN ({placeholders})
        ORDER BY workday_business_object, old_code
        """,
        list(invalid_values) + [value.upper() for value in invalid_values],
    ).fetchall()

    by_value: dict[str, list[dict[str, str]]] = defaultdict(list)
    for business_object, old_code, workday_ref_id, workday_name, old_segment in rows:
        keys = {workday_ref_id or "", old_code or ""}
        for key in keys:
            if key and key in invalid_values:
                by_value[key].append(
                    {
                        "workday_business_object": business_object or "",
                        "old_code": old_code or "",
                        "workday_ref_id": workday_ref_id or "",
                        "workday_name": workday_name or "",
                        "old_segment": old_segment or "",
                    }
                )
    return by_value


def load_reference_field_usage(
    con: duckdb.DuckDBPyConnection, invalid_values: set[str]
) -> dict[str, dict[str, Counter]]:
    usage: dict[str, dict[str, Counter]] = defaultdict(lambda: defaultdict(Counter))

    rows = con.execute(
        """
        SELECT 'Receipted' AS scope
             , spend_category_ref_id
             , cost_centre_ref_id
             , event_ref_id
             , camp_ref_id
             , fa_learning_course_ref_id
             , program_ref_id
             , grant_ref_id
             , commercial_partner_ref_id
        FROM src_fin_supplier_po_receipt_receipted_ref
        UNION ALL
        SELECT 'Unreceipted' AS scope
             , spend_category_ref_id
             , cost_centre_ref_id
             , event_ref_id
             , camp_ref_id
             , fa_learning_course_ref_id
             , program_ref_id
             , grant_ref_id
             , commercial_partner_ref_id
        FROM src_fin_supplier_po_receipt_undelivered_ref
        """
    ).fetchall()

    column_names = ["scope", *REF_COLUMNS.keys()]
    invalid_lookup = set(invalid_values)
    for row in rows:
        record = dict(zip(column_names, row, strict=True))
        scope = str(record["scope"])
        for column_name, (field_name, _type_name) in REF_COLUMNS.items():
            value = record.get(column_name)
            if value is None:
                continue
            text = str(value).strip()
            if text in invalid_lookup:
                usage[text][field_name][scope] += 1
                usage[text][field_name]["Total"] += 1
    return usage


def join_unique(values: list[str]) -> str:
    seen: set[str] = set()
    ordered: list[str] = []
    for value in values:
        text = str(value).strip()
        if not text or text in seen:
            continue
        seen.add(text)
        ordered.append(text)
    return " | ".join(ordered)


def build_camp_rows(
    camp_values: list[str],
    invalid_counts: dict[tuple[str, str], Counter],
    mapping_context: dict[str, list[dict[str, str]]],
    field_usage: dict[str, dict[str, Counter]],
) -> list[dict[str, str | int]]:
    rows: list[dict[str, str | int]] = []
    for camp_id in sorted(set(camp_values)):
        counts = invalid_counts.get((camp_id, "Custom_Organization_Reference_ID"), Counter())
        mapping_rows = mapping_context.get(camp_id, [])
        usage = field_usage.get(camp_id, {})
        rows.append(
            {
                "camp_id": camp_id,
                "receipted_error_rows": counts.get("Receipted", 0),
                "unreceipted_error_rows": counts.get("Unreceipted", 0),
                "total_error_rows": counts.get("Receipted", 0) + counts.get("Unreceipted", 0),
                "current_field": join_unique(list(usage.keys())) or "Camps",
                "mapping_business_object": join_unique(
                    [row["workday_business_object"] for row in mapping_rows]
                ),
                "source_old_code": join_unique([row["old_code"] for row in mapping_rows]) or camp_id,
                "current_workday_ref_id": join_unique(
                    [row["workday_ref_id"] for row in mapping_rows]
                )
                or camp_id,
                "current_workday_name": join_unique([row["workday_name"] for row in mapping_rows]),
                "old_segment": join_unique([row["old_segment"] for row in mapping_rows]),
                "action_needed": "Set up this camp/custom org in Workday tenant or confirm the target ID.",
            }
        )
    return rows


def build_mapping_fix_rows(
    invalid_counts: dict[tuple[str, str], Counter],
    mapping_context: dict[str, list[dict[str, str]]],
    field_usage: dict[str, dict[str, Counter]],
) -> list[dict[str, str | int]]:
    rows: list[dict[str, str | int]] = []
    for (invalid_value, type_name), counts in sorted(
        invalid_counts.items(), key=lambda item: (-sum(item[1].values()), item[0][0])
    ):
        if type_name == "Custom_Organization_Reference_ID":
            continue

        mapping_rows = mapping_context.get(invalid_value, [])
        usage = field_usage.get(invalid_value, {})
        field_counts = [
            f"{field}: {counter.get('Total', 0)}"
            for field, counter in sorted(usage.items())
        ]
        rows.append(
            {
                "invalid_value": invalid_value,
                "workday_type": type_name,
                "receipted_error_rows": counts.get("Receipted", 0),
                "unreceipted_error_rows": counts.get("Unreceipted", 0),
                "total_error_rows": counts.get("Receipted", 0) + counts.get("Unreceipted", 0),
                "current_field": join_unique(list(usage.keys())),
                "field_counts": join_unique(field_counts),
                "mapping_business_object": join_unique(
                    [row["workday_business_object"] for row in mapping_rows]
                ),
                "source_old_code": join_unique([row["old_code"] for row in mapping_rows]),
                "current_workday_ref_id": join_unique(
                    [row["workday_ref_id"] for row in mapping_rows]
                )
                or invalid_value,
                "current_workday_name": join_unique([row["workday_name"] for row in mapping_rows]),
                "old_segment": join_unique([row["old_segment"] for row in mapping_rows]),
                "action_needed": action_needed_for_mapping_type(type_name, invalid_value),
            }
        )
    return rows


def action_needed_for_mapping_type(type_name: str, invalid_value: str) -> str:
    if invalid_value == "STRING OVERRIDE":
        return "Replace the placeholder target with a valid Workday spend category."
    if type_name == "Spend_Category_ID":
        return "Correct the spend-category mapping target."
    if type_name == "Cost_Center_Reference_ID":
        return "Correct the cost-centre mapping target or confirm the valid Workday ID."
    if type_name == "Custom_Worktag_3_ID":
        return "Correct the commercial partner mapping target."
    if type_name == "Program_ID":
        return "Correct the programme mapping target or set up the programme in Workday."
    if type_name == "Grant_ID":
        return "Correct the grant mapping target or set up the grant in Workday."
    return "Review and correct the current mapping target."


def build_other_issue_rows(
    family_counts: Counter,
) -> list[dict[str, str | int]]:
    rows: list[dict[str, str | int]] = []
    families = sorted({family for family, _scope in family_counts})
    ignored = {
        "Invalid custom org ID",
        "Invalid Spend_Category_ID",
        "Invalid Cost_Center_Reference_ID",
        "Invalid Custom_Worktag_3_ID",
        "Invalid Program_ID",
        "Invalid Grant_ID",
    }
    for family in families:
        if family in ignored:
            continue
        action_owner, action_needed = suggested_action_for_family(family)
        rows.append(
            {
                "issue_family": family,
                "receipted_rows": family_counts.get((family, "Receipted"), 0),
                "unreceipted_rows": family_counts.get((family, "Unreceipted"), 0),
                "total_rows": family_counts.get((family, "Receipted"), 0)
                + family_counts.get((family, "Unreceipted"), 0),
                "owner": action_owner,
                "action_needed": action_needed,
            }
        )
    rows.sort(key=lambda row: (-int(row["total_rows"]), str(row["issue_family"])))
    return rows


def write_sheet(workbook: Workbook, name: str, rows: list[dict[str, str | int]]) -> None:
    ws = workbook.create_sheet(title=name)
    if not rows:
        ws.append(["No rows"])
        return

    headers = list(rows[0].keys())
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP

    for row in rows:
        ws.append([row.get(header, "") for header in headers])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for column in ws.columns:
        letter = get_column_letter(column[0].column)
        max_len = max(len(str(cell.value or "")) for cell in column)
        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 45)
        for cell in column:
            cell.alignment = WRAP if max_len > 24 else TOP


def build_workbook(
    summary_rows: list[dict[str, str | int]],
    camp_rows: list[dict[str, str | int]],
    mapping_rows: list[dict[str, str | int]],
    other_rows: list[dict[str, str | int]],
    output_path: str,
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    write_sheet(workbook, "summary", summary_rows)
    write_sheet(workbook, "camp_setup", camp_rows)
    write_sheet(workbook, "mapping_fixes", mapping_rows)
    write_sheet(workbook, "other_issues", other_rows)
    Path(output_path).expanduser().resolve().parent.mkdir(parents=True, exist_ok=True)
    workbook.save(Path(output_path).expanduser().resolve())


def main(argv=None):
    args = parse_args(argv)
    input_path = os.path.expanduser(args.input)
    output_path = os.path.expanduser(args.out)
    duckdb_path = resolve_duckdb_path(args.duckdb, args.runtime_config)

    sheets = load_error_workbook(input_path)
    summary_rows, family_counts, invalid_counts, _invalid_row_refs = summarize_workbook_issues(
        sheets
    )
    invalid_values = {value for value, _type_name in invalid_counts}

    con = duckdb.connect(duckdb_path, read_only=True)
    try:
        mapping_context = load_mapping_context(con, invalid_values)
        field_usage = load_reference_field_usage(con, invalid_values)
    finally:
        con.close()

    camp_rows = build_camp_rows(
        [camp for camp, _blank in sheets["Camps Missing"]],
        invalid_counts,
        mapping_context,
        field_usage,
    )
    mapping_rows = build_mapping_fix_rows(invalid_counts, mapping_context, field_usage)
    other_rows = build_other_issue_rows(family_counts)

    build_workbook(summary_rows, camp_rows, mapping_rows, other_rows, output_path)
    print(f"Workbook saved: {Path(output_path).expanduser().resolve()}")
    print(f"Summary rows: {len(summary_rows)}")
    print(f"Camp setup rows: {len(camp_rows)}")
    print(f"Mapping fix rows: {len(mapping_rows)}")
    print(f"Other issue rows: {len(other_rows)}")


if __name__ == "__main__":
    main()
