"""Build an operator workbook for unresolved Workday supplier EIB errors.

This is the deterministic follow-up pack for the rows left blank in the strict
`populate_eib_corrections.py` output. It reads the correction audit CSV,
selects all non-AUTO rows, enriches them from the current published supplier
output in DuckDB, and writes a simple Excel workbook for GP remediation.

The workbook is intended for business follow-up, not for patching the EIB:
one row per unresolved error line, with the GP Vendor ID, Workday Supplier ID,
error-tab row, field to correct, and the action needed in GP.

Usage
-----
    uv run --package elt-ingest-excel python \
      elt_ingest_excel/scripts/generate_unresolved_gp_fixes.py
"""

from __future__ import annotations

import argparse
import csv
import os
from collections import Counter, defaultdict

import duckdb
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

DEFAULT_AUDIT = os.path.expanduser(
    "~/Downloads/Master Suppliers V1 24.07.2026 with errors CORRECTED urgent audit.csv"
)
DEFAULT_OUTPUT = os.path.expanduser(
    "~/Downloads/Master Suppliers V1 24.07.2026 unresolved GP fixes.xlsx"
)
DEFAULT_DUCKDB = os.path.expanduser("~/Documents/__data/duckdb/rpatel.duckdb")


def parse_args(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument(
        "--audit",
        default=DEFAULT_AUDIT,
        help="populate_eib_corrections audit CSV (default: latest urgent audit in Downloads)",
    )
    ap.add_argument(
        "--out",
        default=DEFAULT_OUTPUT,
        help="output workbook path (default: unresolved GP fixes workbook in Downloads)",
    )
    ap.add_argument(
        "--duckdb",
        default=DEFAULT_DUCKDB,
        help="published supplier DuckDB for contextual supplier/bank fields",
    )
    return ap.parse_args(argv)


def load_unresolved_rows(audit_path):
    with open(audit_path, newline="", encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))
    unresolved = [r for r in rows if r["status"] != "AUTO"]
    unresolved.sort(key=lambda r: int(r["error_tab_row"]))
    return unresolved


def load_supplier_context(db_path):
    con = duckdb.connect(db_path, read_only=True)
    rows = con.execute(
        """
        SELECT
             n.reference_id
           , n.supplier_id
           , n.supplier_name
           , s.bank_country
           , s.bank_account_number
           , s.bank_code_routing_number
           , s.bank_iban
           , s.bank_swift_code
        FROM workday_supplier_name n
        LEFT JOIN workday_supplier_settlement_account s
          USING (supplier_id)
        WHERE n.reference_id IS NOT NULL
          AND TRIM(n.reference_id) <> ''
        """
    ).fetchall()
    con.close()

    by_ref = defaultdict(list)
    for ref, sid, name, country, acct, routing, iban, bic in rows:
        by_ref[ref].append(
            {
                "reference_id": ref,
                "supplier_id": sid,
                "supplier_name": name,
                "bank_country": country,
                "bank_account_number": acct,
                "bank_code_routing_number": routing,
                "bank_iban": iban,
                "bank_swift_code": bic,
            }
        )
    return by_ref


def field_to_correct(error_text):
    if "Routing Transit or Institution Number is required" in error_text:
        return "Routing Transit or Institution Number"
    if "Branch ID is required" in error_text:
        return "Branch ID"
    if "Bank Identification Code (BIC)" in error_text or "BIC" in error_text:
        return "Bank Identification Code (BIC)"
    if "missing address components" in error_text and "City, Postal Code" in error_text:
        return "City and Postal Code"
    if "Bank Code must be 3 numeric digits" in error_text:
        return "Bank Code"
    if "Bank Sort Code must be 6 numeric digits" in error_text:
        return "Bank Sort Code"
    return "Review error text"


def gp_fix_needed(error_text, status):
    if "Routing Transit or Institution Number is required" in error_text:
        return "Populate a valid domestic routing or institution number in GP."
    if "Branch ID is required" in error_text:
        return "Populate the required branch ID for this bank account in GP."
    if "Bank Identification Code (BIC)" in error_text or "BIC" in error_text:
        if status == "SUGGEST":
            return (
                "Confirm and populate the correct 8 or 11 character BIC in GP. "
                "A likely value existed but was not auto-applied."
            )
        return "Replace the invalid BIC with a valid 8 or 11 character SWIFT/BIC in GP."
    if "missing address components" in error_text and "City, Postal Code" in error_text:
        return "Populate both City and Postal Code in GP for the address sent to Workday."
    if "Bank Code must be 3 numeric digits" in error_text:
        return "Replace the invalid value with a valid 3-digit bank code in GP."
    if "Bank Sort Code must be 6 numeric digits" in error_text:
        return "Replace the invalid value with a valid 6-digit UK bank sort code in GP."
    return "Review and correct the underlying source value in GP."


def current_target_value(field_name, rec):
    if not rec:
        return ""
    if field_name in (
        "Routing Transit or Institution Number",
        "Bank Code",
        "Bank Sort Code",
    ):
        return rec.get("bank_code_routing_number") or ""
    if field_name == "Bank Identification Code (BIC)":
        return rec.get("bank_swift_code") or ""
    return ""


def combine_record_context(records):
    def uniq_join(values):
        ordered = []
        seen = set()
        for v in values:
            if v in (None, ""):
                continue
            text = str(v)
            if text in seen:
                continue
            seen.add(text)
            ordered.append(text)
        return " / ".join(ordered)

    return {
        "supplier_id": uniq_join(r["supplier_id"] for r in records),
        "supplier_name": uniq_join(r["supplier_name"] for r in records),
        "bank_country": uniq_join(r["bank_country"] for r in records),
        "bank_account_number": uniq_join(r["bank_account_number"] for r in records),
        "bank_code_routing_number": uniq_join(
            r["bank_code_routing_number"] for r in records
        ),
        "bank_iban": uniq_join(r["bank_iban"] for r in records),
        "bank_swift_code": uniq_join(r["bank_swift_code"] for r in records),
    }


def assign_combined_reference_rows(unresolved, by_ref):
    rows_by_ref = defaultdict(list)
    for row in unresolved:
        rows_by_ref[row["reference_id"]].append(row)

    assignment = {}
    for combined_ref, rows in rows_by_ref.items():
        if combined_ref in by_ref or "|" not in combined_ref:
            continue
        parts = [p.strip() for p in combined_ref.split("|") if p.strip()]
        part_records = [by_ref[p][0] for p in parts if p in by_ref]
        if not part_records:
            continue
        ordered_rows = sorted(rows, key=lambda r: int(r["error_tab_row"]))
        chunk = max(1, len(ordered_rows) // len(part_records))
        for idx, row in enumerate(ordered_rows):
            part_idx = min(idx // chunk, len(part_records) - 1)
            assignment[(combined_ref, row["error_tab_row"])] = part_records[part_idx]
    return assignment


def resolve_row_context(row, by_ref, combined_assignment):
    ref = row["reference_id"]
    if ref in by_ref:
        return by_ref[ref][0], ref
    if "|" in ref:
        assigned = combined_assignment.get((ref, row["error_tab_row"]))
        if assigned:
            return assigned, assigned["reference_id"]
        parts = [p.strip() for p in ref.split("|") if p.strip()]
        records = [by_ref[p][0] for p in parts if p in by_ref]
        if records:
            return combine_record_context(records), ref
    return None, ref


def build_workbook(unresolved, by_ref, output_path):
    combined_assignment = assign_combined_reference_rows(unresolved, by_ref)

    wb = Workbook()
    ws = wb.active
    ws.title = "gp_fix_rows"

    headers = [
        "GP Vendor ID",
        "Workday Supplier ID",
        "Supplier Name",
        "Error Tab Row",
        "Spreadsheet Key",
        "Field To Correct",
        "Error Message",
        "Bank Country",
        "Current Published Target Value",
        "Current Published Account Number",
        "Current Published Routing/Bank Code",
        "Current Published IBAN",
        "Current Published BIC",
        "GP Fix Needed",
        "Status",
    ]
    ws.append(headers)

    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font

    for row in unresolved:
        rec, gp_vendor_id = resolve_row_context(row, by_ref, combined_assignment)
        field_name = field_to_correct(row["error"])
        status = (
            "Manual intervention required"
            if row["status"] == "MANUAL"
            else "Manual confirmation required"
        )
        ws.append(
            [
                gp_vendor_id,
                rec.get("supplier_id", "") if rec else "",
                rec.get("supplier_name", "") if rec else "",
                int(row["error_tab_row"]),
                int(row["spreadsheet_key"]),
                field_name,
                row["error"],
                rec.get("bank_country", "") if rec else "",
                current_target_value(field_name, rec),
                rec.get("bank_account_number", "") if rec else "",
                rec.get("bank_code_routing_number", "") if rec else "",
                rec.get("bank_iban", "") if rec else "",
                rec.get("bank_swift_code", "") if rec else "",
                gp_fix_needed(row["error"], row["status"]),
                status,
            ]
        )

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for idx, width in {
        1: 16,
        2: 19,
        3: 34,
        4: 12,
        5: 14,
        6: 34,
        7: 95,
        8: 16,
        9: 28,
        10: 24,
        11: 28,
        12: 24,
        13: 20,
        14: 72,
        15: 28,
    }.items():
        ws.column_dimensions[get_column_letter(idx)].width = width

    summary = wb.create_sheet("summary")
    summary.append(["Metric", "Value"])
    summary["A1"].fill = header_fill
    summary["B1"].fill = header_fill
    summary["A1"].font = header_font
    summary["B1"].font = header_font
    summary_rows = [
        ("Total unresolved error rows", len(unresolved)),
        (
            "Rows needing manual intervention",
            sum(1 for r in unresolved if r["status"] == "MANUAL"),
        ),
        (
            "Rows needing manual confirmation",
            sum(1 for r in unresolved if r["status"] == "SUGGEST"),
        ),
    ]
    for metric, value in summary_rows:
        summary.append([metric, value])
    summary.append([])
    summary.append(["Field To Correct", "Count"])
    summary["A5"].fill = header_fill
    summary["B5"].fill = header_fill
    summary["A5"].font = header_font
    summary["B5"].font = header_font
    for field_name, count in Counter(
        field_to_correct(r["error"]) for r in unresolved
    ).most_common():
        summary.append([field_name, count])
    summary.column_dimensions["A"].width = 42
    summary.column_dimensions["B"].width = 12
    summary.freeze_panes = "A2"

    wb.save(output_path)


def verify_workbook(output_path, expected_rows):
    wb = load_workbook(output_path, read_only=True, data_only=True)
    row_count = wb["gp_fix_rows"].max_row - 1
    summary_total = wb["summary"]["B2"].value
    wb.close()
    if row_count != expected_rows or summary_total != expected_rows:
        raise SystemExit(
            f"Verification failed for {output_path}: "
            f"rows={row_count}, summary_total={summary_total}, expected={expected_rows}"
        )


def main(argv=None):
    args = parse_args(argv)
    audit_path = os.path.expanduser(args.audit)
    output_path = os.path.expanduser(args.out)
    duckdb_path = os.path.expanduser(args.duckdb)

    if not os.path.exists(audit_path):
        raise SystemExit(f"Audit CSV not found: {audit_path}")
    if not os.path.exists(duckdb_path):
        raise SystemExit(f"DuckDB not found: {duckdb_path}")

    unresolved = load_unresolved_rows(audit_path)
    by_ref = load_supplier_context(duckdb_path)
    build_workbook(unresolved, by_ref, output_path)
    verify_workbook(output_path, len(unresolved))

    counts = Counter(r["status"] for r in unresolved)
    print(f"Audit : {audit_path}")
    print(f"Output: {output_path}")
    print(
        "Unresolved rows: "
        f"{len(unresolved)}  |  MANUAL={counts['MANUAL']}  SUGGEST={counts['SUGGEST']}"
    )


if __name__ == "__main__":
    main()
