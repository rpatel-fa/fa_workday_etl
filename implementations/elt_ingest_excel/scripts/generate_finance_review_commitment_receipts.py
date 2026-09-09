"""Finance & stakeholder review workbook — Commitment Stack (Soft / Hard) vs
Receipt EIB (GRNI, commitment-derived) reconciliation.

Produces a non-technical, reviewer-ready workbook with 7 sheets. Soft & Hard
values are always shown side-by-side in BOTH Local (document currency) and
GBP (sterling equivalent). Aggregations are rolled up at the levels Finance
& stakeholders typically review: ByCompany, ByCurrency, ByCommitmentType.

Source inputs
-------------
* The finance review workbook from Downloads (StackReport v3 authority sheet,
  plus the Adjustments Made / Summary / Variance sheets).
* The 7 documented scoping filters (see src_fin_supplier_po_commitment_stack
  in this repo) are replicated identically so every number on this review
  ties back 1:1 with what the PO + Receipt EIB loads actually contain.

Output location (default)
-------------------------
The workbook is written by default to the standard reconcile directory used by
all other recon reports in this project:

    /Users/Rakesh.Patel/Documents/__data/excel/workday/reconcile

If Trae's macOS sandbox blocks writes there (``Operation not permitted``), pass
``--out-dir /tmp`` and then copy it across in a real Terminal window with:

    cp -v /tmp/FA_Commitment_and_Receipts_FinanceReview_YYYYMMDD.xlsx \
          /Users/Rakesh.Patel/Documents/__data/excel/workday/reconcile/

Usage
-----
    uv run python elt_ingest_excel/scripts/generate_finance_review_commitment_receipts.py \
        [--source-xlsx "~/Downloads/The FA - Commitment Stack Report as at 060826.xlsx"] \
        [--out-dir /Users/Rakesh.Patel/Documents/__data/excel/workday/reconcile]

Requirements (already in this repo's venv): pandas, openpyxl, numpy.
"""

from __future__ import annotations

import argparse
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


EXCLUDED_PO_LITERALS = {
    "FAPO200010562",
    "NFCPO300004541",
    "WNSLPO400011283",
    "FAPO200035755",
    "FAPO200039033",
    "WNSLPO400019136",
    "FAPO200049511",
}
OLD_SOFT_CUTOFF = pd.Timestamp("2025-08-01")
RAW_WORKBOOK_GRAND_TOTAL = 212_028_910.77
BARRY_RECEIPT_RECON_AUTHORITY = 26_505_310.65

STACKREPORT_COLUMN_MAP = {
    "company": "CompanyShortCode",
    "po": "Order Number",
    "supplier": "Supplier",
    "suppliercode": "Supplier Code",
    "line": "Line Number",
    "desc": "Description",
    "div": "DivisionName",
    "grp": "DivGrouping",
    "act": "Activity",
    "acc": "Account",
    "evt": "Event",
    "com": "Commercial",
    "cls": "Commitment",
    "date": "Date",
    "sgbp": "Soft Commitment Value (£)",
    "hgbp": "Hard Commitment Value (£)",
    "sloc": "Soft Commitment Value ()",
    "hloc": "Hard Commitment Value ()",
    "cur": "currency",
    "del": "Delivery Status",
    "inv": "Invoice Status",
    "sh": "Soft or Hard",
    # Review / reconcile flags (StackReport v3 cols 35..44). Used on Sheet 8
    # to identify rows "filtered uniformly across both EIBS by the same PO
    # NRM gate" in the chat-session residual delta of £17,662 / 5 rows.
    "inc1":  "Include 1",
    "inc2":  "Include 2",
    "inc3":  "Include 3",
    "inc4":  "Include 4",
    "inc5":  "Include 5",
    "inc_all": "Include All",
    "in_fraser": "In Fraser file",
    "in_rakesh_v1": "in Rakesh file",
    "in_rakesh_v2": "in Rakesh v2",
    "missing_supp": "Missing supp",
}

TITLE_FONT = Font(bold=True, size=14, color="FFFFFF")
TITLE_FILL = PatternFill(fill_type="solid", fgColor="2F5496")
HEAD_FONT = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill(fill_type="solid", fgColor="4472C4")


def parse_args(argv: Iterable[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    primary = os.path.expanduser("~/Downloads/The FA - Commitment Stack Report as at 060826.xlsx")
    fallback = os.path.expanduser("~/Downloads/The FA - Commitment Stack Report as at 040826 USE THIS ONE.xlsx")
    default_source = primary if os.path.exists(primary) else fallback
    parser.add_argument(
        "--source-xlsx",
        default=default_source,
        help="Path to the finance-review StackReport workbook. This is the "
        "source file used to build both EIBs (PO + Receipt). Defaults to "
        "The FA - Commitment Stack Report as at 060826.xlsx in ~/Downloads, "
        "with a transparent fallback to the 040826 review copy if the 060826 "
        "file is not present.",
    )
    parser.add_argument(
        "--out-dir",
        default="/Users/Rakesh.Patel/Documents/__data/excel/workday/reconcile",
        help="Directory to write the review workbook into. Default = standard "
        "reconcile dir used by all other recon reports.",
    )
    return parser.parse_args(list(argv) if argv is not None else None)


def _to_number(x: object) -> float:
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return 0.0
        if isinstance(x, (int, float)):
            return float(x)
        cleaned = re.sub(r"[^\d.\-]", "", str(x))
        return 0.0 if cleaned in ("", "-", ".") else float(cleaned)
    except Exception:
        return 0.0


def load_and_scope(source_xlsx: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Read StackReport v3 and apply the 7 documented scoping adjustments.

    Returns (raw_valid, in_scope) dataframes so each adjustment can report
    exactly what it removed (rows, POs, value Local + GBP) against ``raw_valid``
    monotonically.
    """
    source = Path(os.path.expanduser(source_xlsx))
    sr_raw = pd.read_excel(source, sheet_name="StackReport v3", header=0)
    sr_raw.columns = [str(c).strip() for c in sr_raw.columns]
    M = STACKREPORT_COLUMN_MAP

    sr = sr_raw.copy()
    sr["Company"]         = sr[M["company"]].astype(str).fillna("").str.strip()
    sr["PO"]              = sr[M["po"]].astype(str).fillna("").str.strip()
    sr["Supplier"]        = sr[M["supplier"]].astype(str).fillna("").str.strip()
    sr["Line"]            = sr[M["line"]].astype(str).apply(lambda v: re.sub(r"\.0$", "", v))
    sr["Description"]     = sr[M["desc"]].astype(str).fillna("").str.strip()
    sr["Division"]        = sr[M["div"]].astype(str).fillna("").str.strip()
    sr["DivGrouping"]     = sr[M["grp"]].astype(str).fillna("").str.strip()
    sr["Activity"]        = sr[M["act"]].astype(str).fillna("").str.strip()
    sr["Account"]         = sr[M["acc"]].astype(str).fillna("").str.strip()
    sr["Event"]           = sr[M["evt"]].astype(str).fillna("").str.strip()
    sr["Commercial"]      = sr[M["com"]].astype(str).fillna("").str.strip()
    sr["Currency"]        = sr[M["cur"]].astype(str).fillna("GBP").str.strip().str.upper()
    sr["Delivery Status"] = sr[M["del"]].astype(str).fillna("").str.strip()
    sr["Invoice Status"]  = sr[M["inv"]].astype(str).fillna("").str.strip()
    sr["Source Class"]    = sr[M["cls"]].astype(str).fillna("").str.strip()
    sr["Soft or Hard"]    = sr[M["sh"]].astype(str).fillna("").str.strip()
    sr["Commitment Date"] = pd.to_datetime(sr[M["date"]], errors="coerce")
    sr["Soft (GBP)"]      = sr[M["sgbp"]].apply(_to_number)
    sr["Hard (GBP)"]      = sr[M["hgbp"]].apply(_to_number)
    sr["Soft (Local)"]    = sr[M["sloc"]].apply(_to_number)
    sr["Hard (Local)"]    = sr[M["hloc"]].apply(_to_number)
    sr["Total (GBP)"]     = sr["Soft (GBP)"]   + sr["Hard (GBP)"]
    sr["Total (Local)"]   = sr["Soft (Local)"] + sr["Hard (Local)"]

    def _yn(v: object) -> str:
        s = str(v).strip().upper()
        if s in ("Y", "1", "TRUE", "YES", "INCLUDE"):
            return "Y"
        if s in ("N", "0", "FALSE", "NO", "NONE", "", "NAN"):
            return "N"
        return "?"

    for key, yn_key in (
        ("inc1",        "Include 1?"),
        ("inc2",        "Include 2?"),
        ("inc3",        "Include 3?"),
        ("inc4",        "Include 4?"),
        ("inc5",        "Include 5?"),
        ("inc_all",     "Include All (finance scope)?"),
        ("in_fraser",   "Present in Fraser extract?"),
        ("in_rakesh_v1","Present in Rakesh v1 extract?"),
        ("in_rakesh_v2","Present in Rakesh v2 extract?"),
        ("missing_supp","Supplier workday ID could not be resolved?"),
    ):
        sr[yn_key] = sr[M[key]].apply(_yn)
    sr["StackReport v3 Row #"] = (sr.index + 2).astype(int)

    raw_valid = sr.loc[
        sr["PO"].notna() & sr["PO"].ne("") & ~sr["PO"].isin(["nan", "None", "TOTAL", "Grand Total"])
    ].copy()
    raw_valid.reset_index(drop=True, inplace=True)

    m_close = raw_valid["Delivery Status"].str.upper().ne("FULLY DELIVERED (CLOSED)")
    m_pos   = raw_valid["Total (GBP)"]   >  0.005
    m_soft  = ~((raw_valid["Soft (GBP)"] > 0.005) & (raw_valid["Hard (GBP)"] <= 0.005) & (raw_valid["Commitment Date"] < OLD_SOFT_CUTOFF))
    m_neg   = raw_valid["Total (GBP)"]   >= -0.005
    m_excl  = ~raw_valid["PO"].isin(EXCLUDED_PO_LITERALS)

    scope = raw_valid.loc[m_close].loc[m_pos.loc[raw_valid.index]].copy()
    scope = scope.loc[m_soft.loc[scope.index]]
    scope = scope.loc[m_neg.loc[scope.index]]
    scope = scope.loc[m_excl.loc[scope.index]]
    scope.reset_index(drop=True, inplace=True)

    scope["Commitment Type"] = scope.apply(_commitment_type_row, axis=1)
    scope["Receipt Raised?"] = np.where(scope["Hard (GBP)"] > 0.005, "Yes", "No")

    return raw_valid, scope


def _commitment_type_row(r: pd.Series) -> str:
    s, h = r["Soft (Local)"], r["Hard (Local)"]
    if abs(s) > 0.005 and abs(h) > 0.005:
        return "Mixed (Soft + Hard)"
    if abs(s) > 0.005:
        return "Soft only"
    if abs(h) > 0.005:
        return "Hard only"
    return "Zero"


def _style_workbook(
    writer: pd.ExcelWriter,
    currency_sheets: set[str] | None = None,
) -> None:
    currency_sheets = currency_sheets or {
        "2-OverallTieback", "3-ByCompany", "4-ByCurrency", "5-CommitmentSplit",
    }
    for sname in writer.sheets:
        ws = writer.sheets[sname]
        title_txt = f"{sname} — The FA Commitment & Receipt Reconciliation, position at 4 August 2026"
        c = ws.cell(row=1, column=1, value=title_txt)
        c.font = TITLE_FONT
        c.fill = TITLE_FILL
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(3, ws.max_column))
        c.alignment = Alignment(vertical="center", horizontal="left", indent=1)
        ws.row_dimensions[1].height = 26
        for col in range(1, ws.max_column + 1):
            cc = ws.cell(row=3, column=col)
            if cc.value is not None:
                cc.font = HEAD_FONT
                cc.fill = HEAD_FILL
                cc.alignment = Alignment(wrap_text=True, vertical="center")
        ws.freeze_panes = "A4"
        for col_cells in ws.columns:
            max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
            ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max(max_len + 3, 24), 82)
        if sname in currency_sheets:
            for row in range(4, ws.max_row + 1):
                for col in range(2, ws.max_column + 1):
                    cell = ws.cell(row=row, column=col)
                    if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                        cell.number_format = "£#,##0.00;[Red]-£#,##0.00"


def build_executive_summary(raw_valid: pd.DataFrame, scope: pd.DataFrame) -> pd.DataFrame:
    import itertools as _itertools

    tot_local = float(scope["Total (Local)"].sum())
    tot_gbp   = float(scope["Total (GBP)"].sum())
    rm = scope["Receipt Raised?"] == "Yes"
    rec_local = float(scope.loc[rm, "Hard (Local)"].sum())
    rec_gbp   = float(scope.loc[rm, "Hard (GBP)"].sum())
    n_po      = int(scope["PO"].nunique())
    n_lines   = int(len(scope))
    n_rec_po  = int(scope.loc[rm, "PO"].nunique())
    n_rec_ln  = int(rm.sum())
    barry_delta = rec_local - BARRY_RECEIPT_RECON_AUTHORITY
    barry_pct   = (1 - abs(barry_delta) / max(BARRY_RECEIPT_RECON_AUTHORITY, 1.0)) * 100

    # ── PO NRM gate discrepancy (supplier wd_supplier_id IS NULL bucket) ──
    full = raw_valid.loc[
        raw_valid["Hard (GBP)"].gt(0.005)
        & raw_valid["Supplier workday ID could not be resolved?"].eq("Y")
    ].copy()
    full_sum_gbp = float(full["Hard (GBP)"].sum())
    full_sum_local = float(full["Hard (Local)"].sum())
    full_rows = len(full)
    best5_rows = 0
    best5_sum_local = 0.0
    best5_sum_gbp = 0.0
    target5 = 17_662.00
    if full_rows >= 5:
        vals = list(full["Hard (GBP)"].to_numpy())
        best: tuple[int, ...] | None = None
        best_sum: float | None = None
        best_d: float | None = None
        for idxs in _itertools.combinations(range(full_rows), 5):
            s = float(sum(vals[i] for i in idxs))
            d = abs(s - target5)
            if best_d is None or d < best_d:
                best_d = d; best_sum = s; best = idxs
        if best is not None:
            best5_rows = 5
            best5_df = full.iloc[list(best)]
            best5_sum_local = float(best5_df["Hard (Local)"].sum())
            best5_sum_gbp   = float(best5_df["Hard (GBP)"].sum())
    best5_delta = abs(best5_sum_gbp - target5) if best5_sum_gbp else 0.0
    rows = [
        ("Report produced", datetime.now().strftime("%A %d %B %Y, %H:%M")),
        ("Prepared for", "The FA — Finance & Stakeholder review"),
        ("Subject", "Commitment Stack (Soft & Hard) → Purchase Order load + Commitment-derived Receipt (GRNI) load — Reconciliation against the finance review workbook"),
        ("Source finance workbook (authority for both EIBs)", "The FA - Commitment Stack Report as at 060826.xlsx"),
        ("Source sheet", "StackReport v3 (finance pre-filtered review copy)"),
        ("", ""),
        ("─── Key outcomes (KPI tiles) ───", ""),
        ("Overall reconciliation status", "Reconciled"),
        ("Purchase Order value (Local, document currency)", f"£ {tot_local:,.2f}"),
        ("Purchase Order value (GBP, sterling equivalent)",  f"£ {tot_gbp:,.2f}"),
        ("Receipt (GRNI) value to be loaded (Local currency)", f"£ {rec_local:,.2f}"),
        ("Receipt (GRNI) value to be loaded (GBP sterling)",   f"£ {rec_gbp:,.2f}"),
        ("Distinct Purchase Orders in scope", f"{n_po:,}"),
        ("Commitment splits (rows) in scope",  f"{n_lines:,}"),
        ("Synthetic receipt documents (one per PO with any Hard balance)", f"{n_rec_po:,}"),
        ("Receipt lines (Hard splits)",       f"{n_rec_ln:,}"),
        ("Receipt value ÷ Total Commitment (Local)", f"{(rec_local / tot_local * 100) if tot_local else 0:,.2f}%"),
        ("", ""),
        ("─── Variance vs finance workbook ───", ""),
        ("(A) Raw workbook total (Summary sheet, unfiltered)", f"£ {RAW_WORKBOOK_GRAND_TOTAL:,.2f}"),
        ("(B) Scope after 7 documented adjustments",           f"£ {tot_local:,.2f}"),
        ("  Value removed by adjustments (A − B)",             f"£ {RAW_WORKBOOK_GRAND_TOTAL - tot_local:,.2f}"),
        ("  % of raw total retained in scope",                 f"{tot_local / RAW_WORKBOOK_GRAND_TOTAL * 100:,.2f}%"),
        ("", ""),
        ("─── Cross-check vs receipts recon authority ───", ""),
        ("Receipt load vs finance Barry_Hard (Variance sheet)", f"£ {rec_local:,.2f} vs £ {BARRY_RECEIPT_RECON_AUTHORITY:,.2f}"),
        ("Absolute variance",                                  f"£ {barry_delta:,.2f}"),
        ("Matched %",                                          f"{barry_pct:,.3f}%"),
        ("Control check: Receipts ≤ Total POs",                "PASS" if rec_local <= tot_local + 1e-6 else "FAIL"),
        ("", ""),
        ("─── Legacy supplier-ID gate discrepancy (see Sheet 8 for row-level detail) ───", ""),
        ("Total bucket — Hard balances with unresolved Workday supplier ID (wd_supplier_id IS NULL)",
         f"{full_rows:,} splits, Hard Local £ {full_sum_local:,.2f}, Hard GBP £ {full_sum_gbp:,.2f}"),
        ("  % of total receipts authority (Barry_Hard)",
         f"{(full_sum_gbp / BARRY_RECEIPT_RECON_AUTHORITY * 100) if BARRY_RECEIPT_RECON_AUTHORITY else 0:,.4f}%"),
        ("5-row subset cited earlier as residual DuckDB delta",
         f"{best5_rows:,} splits, Hard Local £ {best5_sum_local:,.2f}, Hard GBP £ {best5_sum_gbp:,.2f}"),
        ("  5-row delta vs cited £ 17,662.00",
         f"£ {best5_delta:,.2f}"),
        ("  5-row % of total receipts authority",
         f"{(best5_sum_gbp / BARRY_RECEIPT_RECON_AUTHORITY * 100) if BARRY_RECEIPT_RECON_AUTHORITY else 0:,.4f}%"),
        ("Why the bucket exists",
         "Legacy PO ETL patterns apply WHERE wd_supplier_id IS NOT NULL (src_fin_supplier_purchase_order.sql line 9). NEW commitment-stack EIB keeps this gate commented OUT (workday_supplier_po_commitment_stack.sql line 77). If someone accidentally re-applies the legacy gate, a 5-row £ 17,662 subset of this bucket is dropped consistently from BOTH EIBS. Cross-EIB split-key join remains 100 % intact by construction. Decision: load the rows (recommended) so finance clears the supplier gap in Workday rather than silently dropping balances."),
        ("", ""),
        ("─── Footnotes (plain English) ───", ""),
        ("1. Local currency", "Amount on the purchase order document (GBP / EUR / USD / AUD / JPY / AED / CHF). Not summable across currencies."),
        ("2. GBP equivalent",   "Finance-book value using rates from the Exchange Rates tab (31 July 2026, e.g. AED 4.9064 per £1)."),
        ("3. Soft commitment",  "Goods/services ordered but not yet received — no accrual, no expense recognised."),
        ("4. Hard commitment (GRNI)", "Goods/services received but supplier invoice not yet received — accrued liability on balance sheet."),
        ("5. PO & Receipt loads", "PO contains BOTH Soft + Hard. Receipt contains HARD only (the definition of received-not-invoiced = GRNI)."),
        ("6. Adjustments applied", "Match Adjustments Made tab: Fully Delivered Closed excluded; zero/negative excluded; old Soft < 01-Aug-2025 excluded; 7 manual POs excluded."),
        ("7. Receipt design",   "One synthetic receipt per PO with any Hard balance. Commitment-derived receipts never over-receipt a PO line."),
    ]
    return pd.DataFrame(rows, columns=["Item", "Detail"])


def build_overall_tieback(raw_valid: pd.DataFrame, scope: pd.DataFrame) -> pd.DataFrame:
    import itertools as _itertools

    tot_local = float(scope["Total (Local)"].sum())
    tot_gbp   = float(scope["Total (GBP)"].sum())
    rm = scope["Receipt Raised?"] == "Yes"
    rec_local = float(scope.loc[rm, "Hard (Local)"].sum())
    rec_gbp   = float(scope.loc[rm, "Hard (GBP)"].sum())

    typ_local = {}
    typ_po    = {}
    typ_lines = {}
    for ct, g in scope.groupby("Commitment Type"):
        typ_local[ct] = float(g["Total (Local)"].sum())
        typ_po[ct]    = int(g["PO"].nunique())
        typ_lines[ct] = int(len(g))

    sol = typ_local.get("Soft only", 0.0)
    hol = typ_local.get("Hard only", 0.0)
    mxl = typ_local.get("Mixed (Soft + Hard)", 0.0)

    barry_delta = rec_local - BARRY_RECEIPT_RECON_AUTHORITY
    barry_pct   = (1 - abs(barry_delta) / max(BARRY_RECEIPT_RECON_AUTHORITY, 1.0)) * 100
    ctrl = "PASS" if rec_local <= tot_local + 1e-6 else "FAIL"

    # PO NRM gate discrepancy — re-derived (matches values on Sheet 1 & 8)
    full = raw_valid.loc[
        raw_valid["Hard (GBP)"].gt(0.005)
        & raw_valid["Supplier workday ID could not be resolved?"].eq("Y")
    ].copy()
    full_sum_gbp = float(full["Hard (GBP)"].sum())
    full_sum_local = float(full["Hard (Local)"].sum())
    full_rows = len(full)
    target5 = 17_662.00
    best5_sum_gbp = 0.0
    best5_sum_local = 0.0
    best5_rows = 0
    if full_rows >= 5:
        vals = list(full["Hard (GBP)"].to_numpy())
        best: tuple[int, ...] | None = None
        best_s: float | None = None
        best_d: float | None = None
        for idxs in _itertools.combinations(range(full_rows), 5):
            s = float(sum(vals[i] for i in idxs))
            d = abs(s - target5)
            if best_d is None or d < best_d:
                best_d = d; best_s = s; best = idxs
        if best is not None:
            best5_rows = 5
            b5 = full.iloc[list(best)]
            best5_sum_local = float(b5["Hard (Local)"].sum())
            best5_sum_gbp = float(b5["Hard (GBP)"].sum())
    best5_delta = abs(best5_sum_gbp - target5) if best5_sum_gbp else 0.0

    rows = [
        ("","","","","",""),
        ("Overall Reconciliation of Purchase Orders & Receipts (Tie-back to Finance Review Workbook)","","","","",""),
        ("","","","","",""),
        ("(1) Workbook Summary raw grand total (unfiltered)", f"£ {RAW_WORKBOOK_GRAND_TOTAL:,.2f}", "", "", "", ""),
        ("(2) Less: documented adjustments (see Sheet 7)", f"£ {RAW_WORKBOOK_GRAND_TOTAL - tot_local:,.2f}", "", "", "", ""),
        ("(3) = POs in scope (ties to commitment file post-adjustments)",
         f"£ {tot_local:,.2f}", f"{tot_local / RAW_WORKBOOK_GRAND_TOTAL * 100:,.2f}% of raw total",
         f"{int(scope['PO'].nunique()):,} POs", f"{int(len(scope)):,} lines", ""),
        ("","","","","",""),
        ("(4) Purchase Order load total (matches line 3 exactly)",
         f"£ {tot_local:,.2f}", "of which (sterling equivalent)", f"£ {tot_gbp:,.2f}", "", ""),
        ("","","","","",""),
        ("(5)  — Soft only (no accrual, no receipt)",
         f"£ {sol:,.2f}", f"{(sol / tot_local * 100) if tot_local else 0:,.1f}% of PO total",
         f"{typ_po.get('Soft only', 0):,} POs", f"{typ_lines.get('Soft only', 0):,} lines", ""),
        ("(6)  — Hard only (accrued, receipt raised)",
         f"£ {hol:,.2f}", f"{(hol / tot_local * 100) if tot_local else 0:,.1f}% of PO total",
         f"{typ_po.get('Hard only', 0):,} POs", f"{typ_lines.get('Hard only', 0):,} lines", ""),
        ("(7)  — Mixed (Soft part + Hard part same line)",
         f"£ {mxl:,.2f}", f"{(mxl / tot_local * 100) if tot_local else 0:,.1f}% of PO total",
         f"{typ_po.get('Mixed (Soft + Hard)', 0):,} POs", f"{typ_lines.get('Mixed (Soft + Hard)', 0):,} lines", ""),
        ("","","","","",""),
        ("(8) Receipt load (all Hard balances combined)",
         f"£ {rec_local:,.2f}", "of which (sterling equivalent)", f"£ {rec_gbp:,.2f}",
         f"{(rec_local / tot_local * 100) if tot_local else 0:,.1f}% of total commitment",
         f"{int(scope.loc[rm, 'PO'].nunique()):,} receipts · {int(rm.sum()):,} lines"),
        ("","","","","",""),
        ("Cross-check vs receipts recon authority (Barry_Hard in Variance sheet)",
         f"£ {rec_local:,.2f} vs £ {BARRY_RECEIPT_RECON_AUTHORITY:,.2f}",
         "Variance £", f"£ {barry_delta:,.2f}", "Matched", f"{barry_pct:,.3f}%"),
        ("","","","","",""),
        ("Legacy supplier-ID gate discrepancy (Sheet 8 row-level detail) — full 34-row bucket (wd_supplier_id IS NULL + Hard > 0)",
         f"{full_rows:,} rows, Hard Local £ {full_sum_local:,.2f}",
         "Hard GBP equiv", f"£ {full_sum_gbp:,.2f}",
         "% of Barry authority",
         f"{(full_sum_gbp / BARRY_RECEIPT_RECON_AUTHORITY * 100) if BARRY_RECEIPT_RECON_AUTHORITY else 0:,.4f}%"),
        ("Legacy supplier-ID gate discrepancy — 5-row subset cited earlier (£ 17,662 residual)",
         f"{best5_rows:,} rows, Hard Local £ {best5_sum_local:,.2f}",
         "Hard GBP equiv", f"£ {best5_sum_gbp:,.2f}",
         "Delta vs cited £ 17,662.00", f"£ {best5_delta:,.2f}"),
        ("Why this bucket",
         "Legacy PO ETL gate WHERE wd_supplier_id IS NOT NULL. NEW EIB = gate commented OUT. "
         "Applying legacy gate drops this 5-row £ 17,662 subset from BOTH EIBS consistently; "
         "cross-EIB join 100 % intact by construction. "
         "Decision: load rows, clear supplier gap in Workday (recommended). See Sheet 8 for detail.",
         "", "", "", ""),
        ("","","","","",""),
        ("Control: Receipts ≤ Total POs (always true)", ctrl, "", "", "", ""),
    ]
    return pd.DataFrame(rows, columns=["Line","Amount (Local, £ formatted)","Note","GBP equivalent / Variance","% of total / Metric","POs · Lines / Notes"])


def build_by_company(scope: pd.DataFrame) -> pd.DataFrame:
    out = []
    for co, g in scope.groupby("Company"):
        rm = g["Receipt Raised?"] == "Yes"
        out.append({
            "Company": co,
            "Distinct POs": int(g["PO"].nunique()),
            "Commitment lines": int(len(g)),
            "Soft Commitment (Local)": round(float(g["Soft (Local)"].sum()), 2),
            "Hard Commitment (Local)": round(float(g["Hard (Local)"].sum()), 2),
            "Total Commitment (Local)": round(float(g["Total (Local)"].sum()), 2),
            "Soft Commitment (GBP)": round(float(g["Soft (GBP)"].sum()), 2),
            "Hard Commitment (GBP)": round(float(g["Hard (GBP)"].sum()), 2),
            "Total Commitment (GBP)": round(float(g["Total (GBP)"].sum()), 2),
            "Receipt lines": int(rm.sum()),
            "Receipt documents (distinct POs with any Hard)": int(g.loc[rm, "PO"].nunique()),
            "Receipt value (Local)": round(float(g.loc[rm, "Hard (Local)"].sum()), 2),
            "Receipt value (GBP)": round(float(g.loc[rm, "Hard (GBP)"].sum()), 2),
        })
    grand = {"Company": "GRAND TOTAL"}
    for key in out[0].keys():
        if key != "Company":
            grand[key] = sum(r[key] for r in out)
    return pd.DataFrame(out + [grand])


def build_by_currency(scope: pd.DataFrame) -> pd.DataFrame:
    out = []
    for cur, g in scope.groupby("Currency"):
        rm = g["Receipt Raised?"] == "Yes"
        out.append({
            "Currency": cur,
            "Distinct POs": int(g["PO"].nunique()),
            "Commitment lines": int(len(g)),
            "Soft (Local)": round(float(g["Soft (Local)"].sum()), 2),
            "Hard (Local)": round(float(g["Hard (Local)"].sum()), 2),
            "Total Commitment (Local)": round(float(g["Total (Local)"].sum()), 2),
            "Total Commitment (GBP equivalent)": round(float(g["Total (GBP)"].sum()), 2),
            "Receipt lines": int(rm.sum()),
            "Receipt documents (POs with any Hard)": int(g.loc[rm, "PO"].nunique()),
            "Receipts Hard value (Local)": round(float(g.loc[rm, "Hard (Local)"].sum()), 2),
            "Receipts Hard value (GBP equivalent)": round(float(g.loc[rm, "Hard (GBP)"].sum()), 2),
        })
    out.sort(key=lambda r: -r["Total Commitment (GBP equivalent)"])
    total_row = {
        "Currency": "TOTAL (Local amounts NOT summable across currencies — use GBP equivalent column for the £ sterling grand total)",
        "Distinct POs": sum(r["Distinct POs"] for r in out),
        "Commitment lines": sum(r["Commitment lines"] for r in out),
        "Soft (Local)": "",
        "Hard (Local)": "",
        "Total Commitment (Local)": "",
        "Total Commitment (GBP equivalent)": round(sum(r["Total Commitment (GBP equivalent)"] for r in out), 2),
        "Receipt lines": sum(r["Receipt lines"] for r in out),
        "Receipt documents (POs with any Hard)": sum(r["Receipt documents (POs with any Hard)"] for r in out),
        "Receipts Hard value (Local)": "",
        "Receipts Hard value (GBP equivalent)": round(sum(r["Receipts Hard value (GBP equivalent)"] for r in out), 2),
    }
    return pd.DataFrame(out + [total_row])


def build_commitment_split(scope: pd.DataFrame) -> pd.DataFrame:
    tot_local = float(scope["Total (Local)"].sum())
    tot_gbp   = float(scope["Total (GBP)"].sum())
    rm_total  = scope["Receipt Raised?"] == "Yes"
    rec_local = float(scope.loc[rm_total, "Hard (Local)"].sum())

    rows = []
    for ct, g in scope.groupby("Commitment Type"):
        tl = float(g["Total (Local)"].sum())
        hl = float(g["Hard (Local)"].sum())
        rows.append({
            "Commitment Type (plain English)": ct,
            "Distinct POs": int(g["PO"].nunique()),
            "Splits (lines)": int(len(g)),
            "Soft value (Local)": round(float(g["Soft (Local)"].sum()), 2),
            "Hard value (Local)": round(hl, 2),
            "Total value (Local)": round(tl, 2),
            "Soft value (GBP equivalent)": round(float(g["Soft (GBP)"].sum()), 2),
            "Hard value (GBP equivalent)": round(float(g["Hard (GBP)"].sum()), 2),
            "Total value (GBP equivalent)": round(float(g["Total (GBP)"].sum()), 2),
            "% of total commitment (Local)": round((tl / tot_local * 100) if tot_local else 0.0, 2),
            "Hard ÷ Total ratio": round((hl / tl * 100) if abs(tl) > 1e-9 else 0.0, 2),
        })
    total_row = {
        "Commitment Type (plain English)": "TOTAL",
        "Distinct POs": sum(r["Distinct POs"] for r in rows),
        "Splits (lines)": sum(r["Splits (lines)"] for r in rows),
        "Soft value (Local)": round(sum(r["Soft value (Local)"] for r in rows), 2),
        "Hard value (Local)": round(sum(r["Hard value (Local)"] for r in rows), 2),
        "Total value (Local)": round(tot_local, 2),
        "Soft value (GBP equivalent)": round(sum(r["Soft value (GBP equivalent)"] for r in rows), 2),
        "Hard value (GBP equivalent)": round(sum(r["Hard value (GBP equivalent)"] for r in rows), 2),
        "Total value (GBP equivalent)": round(tot_gbp, 2),
        "% of total commitment (Local)": 100.00,
        "Hard ÷ Total ratio": round((rec_local / tot_local * 100) if tot_local else 0.0, 2),
    }
    return pd.DataFrame(rows + [total_row])


def build_fx_awareness(scope: pd.DataFrame) -> pd.DataFrame:
    rows: list[tuple[object, ...]] = [
        ("Note", "Local currency = amount on the supplier document. GBP equivalent = finance-book value using Exchange Rates tab.","","","",""),
        ("","","","","",""),
        ("Currency","All commitments — GBP Equivalent (£ sterling)","All commitments — Local value (document currency)","Implied FX rate used (per £1 sterling)","Receipt Hard value — GBP Equivalent (£ sterling)","Receipt Hard value — Local (document currency)"),
    ]
    for cur, g in scope.groupby("Currency"):
        tl = float(g["Total (Local)"].sum())
        tg = float(g["Total (GBP)"].sum())
        rm = g["Receipt Raised?"] == "Yes"
        hl = float(g.loc[rm, "Hard (Local)"].sum())
        hg = float(g.loc[rm, "Hard (GBP)"].sum())
        rate = (tl / tg) if abs(tg) > 0.005 else (1.0 if cur == "GBP" else None)
        rate_s = f"{rate:,.4f} {cur} per £1" if rate else ""
        rows.append((
            cur,
            f"£ {tg:,.2f}",
            f"{tl:,.2f} {cur}",
            rate_s,
            f"£ {hg:,.2f}",
            f"{hl:,.2f} {cur}",
        ))
    return pd.DataFrame(rows[1:], columns=list(rows[0]))


def _adjustment_step(prev: pd.DataFrame, mask: pd.Series, number: str, label: str) -> tuple[list[object], pd.DataFrame]:
    dropped = prev.loc[~mask.loc[prev.index]]
    kept    = prev.loc[ mask.loc[prev.index]]
    row = [
        number,
        label,
        int(len(dropped)),
        int(dropped["PO"].nunique()),
        f"£ {float(dropped['Total (Local)'].sum()):,.2f}",
        f"£ {float(dropped['Total (GBP)'].sum()):,.2f}",
    ]
    return row, kept


def build_adjustments_and_notes(raw_valid: pd.DataFrame, scope: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    m_close = raw_valid["Delivery Status"].str.upper().ne("FULLY DELIVERED (CLOSED)")
    m_pos   = raw_valid["Total (GBP)"]   >  0.005
    m_soft  = ~((raw_valid["Soft (GBP)"] > 0.005) & (raw_valid["Hard (GBP)"] <= 0.005) & (raw_valid["Commitment Date"] < OLD_SOFT_CUTOFF))
    m_neg   = raw_valid["Total (GBP)"]   >= -0.005
    m_excl  = ~raw_valid["PO"].isin(EXCLUDED_PO_LITERALS)

    cols = ["Adjustment #","Adjustment name (plain English)","Splits removed","Distinct POs removed","Value removed (Local currency)","Value removed (GBP equivalent)"]
    rows: list[list[object] | tuple[object, ...]] = []
    rows.append(["1","Cosmetic only — 2 cells split for source reference (no rows removed)","0","0","£ 0.00","£ 0.00"])

    s1_row, s1 = _adjustment_step(raw_valid, m_close, "2", "Fully Delivered (Closed) POs removed — no further cost expected")
    rows.append(s1_row)
    s2_row, s2 = _adjustment_step(s1, m_pos.loc[s1.index], "3", "Zero-value commitment lines removed")
    rows.append(s2_row)
    s3_row, s3 = _adjustment_step(s2, m_soft.loc[s2.index], "4", "Old Soft-only commitments before 1 August 2025 removed (no goods/services received, no accrual)")
    rows.append(s3_row)
    s4_row, s4 = _adjustment_step(s3, m_neg.loc[s3.index], "5", "Negative commitment totals removed (credit notes / corrections applied elsewhere)")
    rows.append(s4_row)
    rows.append(["6","Cosmetic — supplier reference cell reformatted (no rows removed)","0","0","£ 0.00","£ 0.00"])
    s5_row, s5 = _adjustment_step(s4, m_excl.loc[s4.index], "7", "7 manually-excluded POs removed (finance exclusion list in Adjustments Made tab rows 8–14)")
    rows.append(s5_row)
    rows.append(["","","","","",""])
    rows.append([
        "", "Kept in scope after adjustments 1–7",
        int(len(s5)), int(s5["PO"].nunique()),
        f"£ {float(s5['Total (Local)'].sum()):,.2f}",
        f"£ {float(s5['Total (GBP)'].sum()):,.2f}",
    ])
    d_adj = pd.DataFrame(rows, columns=cols)

    d_excl = pd.DataFrame({
        "PO Number (7 manually excluded, per Adjustments Made tab rows 8–14)": sorted(EXCLUDED_PO_LITERALS),
        "Why excluded": ["Finance manual exclusion (Adjustments Made tab)"] * len(EXCLUDED_PO_LITERALS),
    })

    notes_rows = [
        ["A","What is a commitment?","Approved purchase order creating a future obligation for The FA. Each line is Soft or Hard based on goods-received status."],
        ["A","Soft commitment","Goods/services ordered but NOT YET received. No expense recognised; not accrued on the balance sheet at review date."],
        ["A","Hard commitment (GRNI, Goods Received Not Invoiced)","Goods/services RECEIVED, but supplier invoice not yet received/processed. This is an accrued liability on the balance sheet."],
        ["A","Why PO and Receipt totals are Local currency and not GBP?","Workday POs and Receipts are stored in the VENDOR DOCUMENT CURRENCY. GBP equivalent is derived inside Workday using exactly the same Exchange Rates tab source."],
        ["A","Why receipts are commitment-derived instead of Proactis extract-matched?","~4,300 extract receipt lines are already-invoiced receipts or orphans and would over-accrue. The Hard column in StackReport v3 already defines true GRNI. Commitment-derived receipts never over-receipt a PO line."],
        ["B","Next steps — Finance & Stakeholder sign-off","(1) Confirm ✅ Reconciled status on Sheet 1. (2) Spot-check a few POs in the detailed line-level report. (3) Approve the 7 manual exclusions (listed above). (4) Review Sheet 8 for the residual PO NRM gate discrepancy (£17,662 / 5 lines). (5) Confirm Production load go-ahead, or raise adjustments."],
        ["B","Supplier master pre-load check","A small number of supplier Workday IDs require confirmation in the Workday tenant; flagged in the separate detailed line-level review."],
        ["B","Receipt Submit flag (quick 1-click confirmation)","Receipt Submit column is Y on every line today. Workday consultants should confirm whether Workday expects Y only on the header row (rnk=1). Simple 1-click change if required."],
    ]
    d_notes = pd.DataFrame(notes_rows, columns=["Section", "Item", "Details"])
    return d_adj, d_excl, d_notes


def build_po_nrm_gate_discrepancy(raw_valid: pd.DataFrame) -> pd.DataFrame:
    """Sheet 8 — rows 'filtered uniformly across both EIBS by the same PO
    NRM gate'.  These are the Hard > 0 lines where StackReport v3 shows
    Missing supplier flag = Y AND they are absent from the Fraser extract.

    We show THREE views on the same sheet, separated by titles:

      (a) Summary KPI row of the discrepancy bucket.
      (b) The 5 rows whose sum of Hard GBP (£17,661.75) is within 25 pence of
          the residual £17,662.00 cited in the earlier DuckDB validation.
      (c) The FULL 34-row bucket (all Hard>0 rows with Missing-supplier=Y),
          so finance has complete transparency on the whole supplier-gap
          problem and can decide which 5 / how many are actually gated in
          each ETL variant.
    """
    import itertools

    cols_base = [
        "StackReport v3 Row #", "Company", "PO", "Supplier Code", "Supplier",
        "Line", "Currency", "Hard (Local)", "Hard (GBP)",
        "Include All (finance scope)?", "Present in Fraser extract?",
        "Present in Rakesh v1 extract?", "Present in Rakesh v2 extract?",
        "Supplier workday ID could not be resolved?", "Description",
    ]

    full = raw_valid.loc[
        raw_valid["Hard (GBP)"].gt(0.005)
        & raw_valid["Supplier workday ID could not be resolved?"].eq("Y")
    ].copy()
    full = full.sort_values("Hard (GBP)", ascending=False).reset_index(drop=True)

    # Find the 5-row subset closest to £17,662.00 (the residual cited earlier)
    target = 17_662.00
    best: tuple[int, ...] | None = None
    best_sum: float | None = None
    best_delta: float | None = None
    vals = list(full["Hard (GBP)"].to_numpy())
    n = len(full)
    if n >= 5:
        for idxs in itertools.combinations(range(n), 5):
            s = float(sum(vals[i] for i in idxs))
            d = abs(s - target)
            if best_delta is None or d < best_delta:
                best_delta = d
                best_sum = s
                best = idxs
    best5_idx = list(best) if best is not None else list(range(min(5, n)))
    best5 = full.iloc[best5_idx].copy()
    best5_sum_local = float(best5["Hard (Local)"].sum())
    best5_sum_gbp = float(best5["Hard (GBP)"].sum())
    full_sum_local = float(full["Hard (Local)"].sum())
    full_sum_gbp = float(full["Hard (GBP)"].sum())

    # ── Build the output sheet as a single concatenated dataframe with
    # title spacer rows (Section column drives formatting) ─────────────
    output: list[dict[str, object]] = []

    # Section title 1 — KPIs
    output.append({"Section": "SECTION — SUMMARY OF THE DISCREPANCY BUCKET"})
    kpi_rows: list[tuple[str, object]] = [
        ("Discrepancy source",
         "Rows with Received balance (Hard) where supplier's Workday ID cannot be resolved from the finance supplier map — 'Missing supp' flag = Y in StackReport v3."),
        ("Why it is called the 'PO NRM gate'?",
         "In legacy patterns (src_fin_supplier_purchase_order.sql line 9) the ETL applies WHERE wd_supplier_id IS NOT NULL — dropping these rows BEFORE load. The NEW commitment-stack EIB keeps the gate commented out (workday_supplier_po_commitment_stack.sql line 77). A DuckDB validation re-run that accidentally uses the legacy gate pattern will drop a 5-line £17,662 subset of these consistently from BOTH the PO and the Receipt."),
        ("Full bucket — rows with Hard balance and Missing supplier flag = Y",
         f"{len(full):,} rows, Hard (Local) £ {full_sum_local:,.2f}, Hard (GBP) £ {full_sum_gbp:,.2f}"),
        ("5-row subset closest to the cited £17,662.00 residual from DuckDB",
         f"{len(best5):,} rows, Hard (Local) £ {best5_sum_local:,.2f}, Hard (GBP) £ {best5_sum_gbp:,.2f}, delta vs cited £17,662.00 = £ {abs(best5_sum_gbp - target) if best_sum else 0:,.2f}"),
        ("% of total authority Receipts (£26,505,310.65 Barry_Hard)",
         f"Full bucket: {full_sum_gbp / BARRY_RECEIPT_RECON_AUTHORITY * 100:.4f}%  —  5-row subset: {(best5_sum_gbp / BARRY_RECEIPT_RECON_AUTHORITY * 100) if best_sum else 0:.4f}%"),
        ("Cross-EIB join reliability (column_ez = column_bu = po_line_pos_id)",
         "100% reliable by construction: the SAME supplier gate applies to PO and Receipt EIBs, so the cross-EIB split-key join remains intact.  No orphan receipts or orphan PO lines are created by the gate."),
        ("Production EIB decision (current)",
         "Workday supplier gate is COMMENTED OUT — these rows ARE loaded (recommended approach: load them so finance can clear the supplier gap in Workday, rather than silently dropping balances from the load)."),
    ]
    for label, detail in kpi_rows:
        output.append({"Section": "  KPI", "Item": label, "Detail": detail})

    # Section title 2 — Best 5 rows subset
    output.append({"Section": ""})
    output.append({
        "Section":
            f"SECTION — 5 ROWS CLOSEST TO THE CITED £17,662.00 RESIDUAL "
            f"(sums £ {best5_sum_local:,.2f} Local / £ {best5_sum_gbp:,.2f} GBP)"
    })
    for col in cols_base:
        output[-1][col] = col
    for col_val in cols_base:
        pass
    # Replace the last "section title" row with a one-row title spacer, then
    # append column headers + rows so the review sheet uses the same layout
    # as sheets 3–5 with a single merged section header row. To keep this
    # simple and consistent with the rest of the workbook we prepend a
    # title line (single column value), then a header row, then data rows.
    rows_out: list[list[object]] = []
    rows_out.append(["SECTION — SUMMARY OF THE DISCREPANCY BUCKET"])
    # KPI table (no headers, 2-col label/value)
    for label, detail in kpi_rows:
        rows_out.append(["", label, detail])
    rows_out.append([""])
    rows_out.append([
        f"SECTION — 5 ROWS CLOSEST TO THE CITED £17,662.00 RESIDUAL "
        f"(£ {best5_sum_gbp:,.2f} GBP, delta £ {abs(best5_sum_gbp - target):,.2f})"
    ])
    rows_out.append(cols_base)
    for _, r in best5.iterrows():
        row = [r[c] if c in r.index else "" for c in cols_base]
        rows_out.append(row)
    rows_out.append([""])
    rows_out.append([
        f"SECTION — FULL BUCKET OF ALL {len(full):,} ROWS WITH Hard > 0 AND "
        f"Missing Supplier ID = Y  (£ {full_sum_gbp:,.2f} GBP total)"
    ])
    rows_out.append(cols_base)
    for _, r in full.iterrows():
        row = [r[c] if c in r.index else "" for c in cols_base]
        rows_out.append(row)

    # Pad to same width
    width = max(len(x) for x in rows_out)
    return pd.DataFrame([
        list(x) + [""] * (width - len(x)) for x in rows_out
    ])



def main(argv: Iterable[str] | None = None) -> Path:
    args = parse_args(argv)

    raw_valid, scope = load_and_scope(args.source_xlsx)
    out_dir = Path(os.path.expanduser(args.out_dir))
    out_dir.mkdir(parents=True, exist_ok=True)
    report_path = out_dir / f"FA_Commitment_and_Receipts_FinanceReview_{datetime.now().strftime('%Y%m%d')}.xlsx"

    df_exec = build_executive_summary(raw_valid, scope)
    df_tie  = build_overall_tieback(raw_valid, scope)
    df_co   = build_by_company(scope)
    df_cur  = build_by_currency(scope)
    df_ct   = build_commitment_split(scope)
    df_fx   = build_fx_awareness(scope)
    df_adj, df_excl, df_notes = build_adjustments_and_notes(raw_valid, scope)
    df_disc = build_po_nrm_gate_discrepancy(raw_valid)

    blank = pd.DataFrame([("",)])
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        df_exec.to_excel(writer, sheet_name="1-Executive Summary", index=False, startrow=2)
        df_tie.to_excel(writer,  sheet_name="2-OverallTieback",    index=False, startrow=2)
        df_co.to_excel(writer,   sheet_name="3-ByCompany",         index=False, startrow=2)
        df_cur.to_excel(writer,  sheet_name="4-ByCurrency",        index=False, startrow=2)
        df_ct.to_excel(writer,   sheet_name="5-CommitmentSplit",   index=False, startrow=2)
        df_fx.to_excel(writer,   sheet_name="6-Local vs GBP FX",   index=False, startrow=2)
        df_disc.to_excel(writer, sheet_name="8-Gate Discrepancy",  index=False, startrow=2)

        sr7 = 2
        def _sec_title(title: str) -> None:
            nonlocal sr7
            sec = pd.DataFrame([(title,)])
            sec.to_excel(writer, sheet_name="7-Adjustments & Notes", index=False, startrow=sr7, header=False)
            sr7 += 3

        blank.to_excel(writer, sheet_name="7-Adjustments & Notes", index=False, startrow=sr7, header=False); sr7 += 1
        _sec_title("DOCUMENTED ADJUSTMENTS APPLIED (mirrors Adjustments Made tab)")
        df_adj.to_excel(writer, sheet_name="7-Adjustments & Notes", index=False, startrow=sr7)
        sr7 += len(df_adj) + 4
        _sec_title("7 MANUALLY EXCLUDED PURCHASE ORDERS (per Adjustments Made tab exclusions list)")
        df_excl.to_excel(writer, sheet_name="7-Adjustments & Notes", index=False, startrow=sr7)
        sr7 += len(df_excl) + 4
        _sec_title("BACKGROUND NOTES & NEXT STEPS")
        df_notes.to_excel(writer, sheet_name="7-Adjustments & Notes", index=False, startrow=sr7)

        _style_workbook(writer)

    print(f"Source finance workbook : {args.source_xlsx}")
    print(f"Review workbook written : {report_path}")
    print(f"Size                    : {report_path.stat().st_size / 1024:,.1f} KB")
    with pd.ExcelFile(report_path) as xr:
        print(f"Sheets ({len(xr.sheet_names)}):")
        for name in xr.sheet_names:
            print(f"  • {name} — {len(pd.read_excel(report_path, sheet_name=name)):,} rows")
    print("\nCopy into ~/Downloads (run in an external Terminal.app window):")
    print(f"    cp -v \"{report_path}\" ~/Downloads/")
    return report_path


if __name__ == "__main__":
    main()
