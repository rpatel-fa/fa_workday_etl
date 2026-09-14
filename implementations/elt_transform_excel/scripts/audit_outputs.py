#!/usr/bin/env python3
"""Audit shipped EIB output workbooks (the AP/AR teams' past runs) for the
invoice-line-grouping + tax + amount invariants this project depends on.

This is a READ-ONLY outcome-review / regression tool. It does NOT regenerate
output from the macros (that needs Excel); it inspects the workbooks already
sitting in the configured target folders and flags anomalies:

  * dup-ref risk  : the same Invoice Number (col K, AP) appears on >1 row but
                    those rows carry DIFFERENT Spreadsheet Key (col B) values.
                    Workday reads that as N separate headers sharing one
                    external invoice ref and rejects the load. This is the
                    single-line-assumption bug the grouping generalization fixes.
  * stale-vintage : Row ID (EK) constant 1 with Line Order (EM) blank on a
                    multi-row invoice group -> output predates the multi-line fix.
  * control-amount: Control Amount Total (col AS) populated -> should be blank
                    (Workday computes it on load).
  * ge_ne_gf      : Unit Cost (GE) != Extended Amount (GF) on a row -> pre-VAT
                    -in-GF vintage (both should equal the line net).
  * tax triples   : the distinct FE/FF/FK tax-code triples emitted, so the
                    standard-rated value split across transforms is visible.

Usage (standalone, no repo env needed):
  uv run --with openpyxl --no-project python scripts/audit_outputs.py
  uv run --with openpyxl --no-project python scripts/audit_outputs.py --eib-root /path/to/eib
  uv run --with openpyxl --no-project python scripts/audit_outputs.py --json

Default --eib-root is read from config/reference_paths.yaml's tree if resolvable,
else falls back to the documented machine path. Override with --eib-root.
"""
from __future__ import annotations
import argparse, glob, json, os, sys
from collections import OrderedDict, Counter

try:
    import openpyxl
    from openpyxl.utils import column_index_from_string as ci
except ImportError:
    sys.exit("openpyxl required: run via `uv run --with openpyxl --no-project python ...`")

DEFAULT_EIB_ROOT = "/Users/Rakesh.Patel/Documents/__data/excel/finance_ref/eib"

# AP Supplier Invoice columns of interest (v46.x Submit Supplier Invoice sheet).
AP_COLS = dict(B="B", K="K", BG="BG", AS="AS", EK="EK", EM="EM",
               FE="FE", FF="FF", FK="FK", FL="FL", GE="GE", GF="GF")
# AR Customer Invoice columns of interest (v46.1 Submit Customer Invoice sheet).
AR_COLS = dict(B="B", CM="CM", DV="DV", EM="EM", FB="FB")

BLANK_RUN_STOP = 15   # consecutive all-blank rows that end the data scan
DATA_START_ROW = 6    # EIB data starts at row 6


def _scan(path, sheet_name, cols):
    want = {c: ci(v) for c, v in cols.items()}
    maxc = max(want.values())
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb[wb.sheetnames[0]]
    rows, blanks = [], 0
    for row in ws.iter_rows(min_row=DATA_START_ROW, max_col=maxc, values_only=True):
        rec = {c: row[want[c] - 1] for c in want}
        if all(v in (None, "") for v in rec.values()):
            blanks += 1
            if blanks >= BLANK_RUN_STOP:
                break
            continue
        blanks = 0
        rows.append(rec)
    title = ws.title
    wb.close()
    return rows, title


def audit_ap(path):
    rows, title = _scan(path, "Submit Supplier Invoice", AP_COLS)
    # Group by Invoice Number (K). Blank K means this output uses a non-classic
    # column layout (e.g. AMEX writes the invoice ref elsewhere) -> it can't be
    # audited with the classic letters, so flag rather than false-positive.
    invoice_number_blank = all(r["K"] in (None, "") for r in rows) and bool(rows)
    groups = OrderedDict()
    for r in rows:
        k = str(r["K"])
        if r["K"] in (None, ""):
            continue
        groups.setdefault(k, []).append(r)
    multi = {k: g for k, g in groups.items() if len(g) > 1}
    dup_ref = {k: sorted({str(x["B"]) for x in g})[:8]
               for k, g in multi.items() if len({str(x["B"]) for x in g}) > 1}
    ek = {str(r["EK"]) for r in rows}
    em = {str(r["EM"]) for r in rows}
    stale = bool(multi) and ek == {"1"} and em == {"None"}
    as_populated = any(r["AS"] not in (None, "") for r in rows)
    ge_ne_gf = any(r["GE"] != r["GF"] and not (r["GE"] in (None, "") and r["GF"] in (None, ""))
                   for r in rows)
    tax = sorted({(str(r["FE"]), str(r["FF"]), str(r["FK"])) for r in rows})
    return dict(kind="AP", sheet=title, rows=len(rows), distinct_invoices=len(groups),
                multi_line_groups=len(multi), dup_ref=dup_ref, stale_vintage=stale,
                control_amount_populated=as_populated, ge_ne_gf=ge_ne_gf, tax_triples=tax,
                invoice_number_col_blank=invoice_number_blank)


def audit_ar(path):
    rows, title = _scan(path, "Submit Customer Invoice", AR_COLS)
    bs = [str(r["B"]) for r in rows]
    per_b = Counter(bs)
    multi_b = [b for b, n in per_b.items() if n > 1]
    tax = sorted({str(r["FB"]) for r in rows if r["FB"] not in (None, "")})
    return dict(kind="AR", sheet=title, rows=len(rows), distinct_B=len(set(bs)),
                multi_line_B_groups=len(multi_b), line_order_DV=sorted({str(r["DV"]) for r in rows})[:8],
                tax_codes=tax)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--eib-root", default=os.environ.get("EIB_ROOT", DEFAULT_EIB_ROOT))
    ap.add_argument("--json", action="store_true", help="emit machine-readable JSON")
    args = ap.parse_args()
    root = args.eib_root
    if not os.path.isdir(root):
        sys.exit(f"eib root not found: {root} (pass --eib-root)")

    results = []
    for f in sorted(glob.glob(os.path.join(root, "ap", "**", "*output*.xlsx"), recursive=True)):
        try:
            r = audit_ap(f)
        except Exception as e:
            r = dict(kind="AP", error=str(e))
        r["file"] = os.path.relpath(f, root)
        results.append(r)
    for f in sorted(glob.glob(os.path.join(root, "ar", "**", "*output*.xlsx"), recursive=True)):
        try:
            r = audit_ar(f)
        except Exception as e:
            r = dict(kind="AR", error=str(e))
        r["file"] = os.path.relpath(f, root)
        results.append(r)

    if args.json:
        print(json.dumps(results, indent=2, default=str))
        return

    anomalies = 0
    for r in results:
        print(f"\n=== [{r['kind']}] {r['file']}")
        if r.get("error"):
            print(f"  ERROR: {r['error']}"); anomalies += 1; continue
        if r["kind"] == "AP":
            print(f"  rows={r['rows']} invoices={r['distinct_invoices']} multiLineGroups={r['multi_line_groups']}")
            print(f"  tax triples: {r['tax_triples']}")
            flags = []
            if r["invoice_number_col_blank"]:
                flags.append("INVOICE NUMBER col (K) blank -> non-classic layout (e.g. AMEX); classic-letter audit not applicable")
            if r["dup_ref"]:
                ex = list(r["dup_ref"].items())[:3]
                flags.append(f"DUP-REF (same K, different B): {ex}")
            if r["stale_vintage"]:
                flags.append("STALE VINTAGE (multi-row invoice but EK=1/EM=blank -> pre-grouping-fix)")
            if r["control_amount_populated"]:
                flags.append("CONTROL AMOUNT TOTAL populated (should be blank)")
            if r["ge_ne_gf"]:
                flags.append("GE != GF (Unit Cost != Extended Amount)")
            for fl in flags:
                print(f"  *** {fl}"); anomalies += 1
        else:
            print(f"  rows={r['rows']} distinctB={r['distinct_B']} multiLineBgroups={r['multi_line_B_groups']} DV={r['line_order_DV']}")
            print(f"  tax codes: {r['tax_codes']}")

    print(f"\n---\n{len(results)} outputs audited; {anomalies} anomaly flag(s) raised.")
    sys.exit(1 if anomalies else 0)


if __name__ == "__main__":
    main()
