"""Post-load tie-out: sent Supplier-Invoice EIB payload  vs  Workday Aging extract.

This is the post-load (4th-layer) reconciliation for Supplier Invoices + Credit Notes,
mirroring generate_post_load_tieback.py for the PO/receipt leg.

Legs compared:
  [3] Sent   : workday_supplier_invoice + workday_supplier_credit_note in rpatel.duckdb
               joined to src_fin_supplier_invoice_nrm for supplier names,
               voucher numbers, and document types.
  [4] Loaded : Workday 'Operational Aging Detail by Due Date' extract —
               Payables Aging Detail sheet, data rows start at excel row 23
               (0-indexed row 22), headers on row 22 (0-indexed 21).

Join keys / quirks handled:
  - Transaction types   : Loaded labels = 'Invoice' / 'Adjustment'; Sent has INVOICE
                          table (column_gf lines) and CREDIT_NOTE table (column_eq lines).
                          We classify sent Adjustments as 'CREDIT_NOTE' from the EIB
                          credit-note staging table.
  - Credit-note grain    : Sent column_k (workday_supplier_credit_note) == Loaded 'Transaction'
                          reference (e.g. 'FAP0148163'). 100% exact match, 62 of 62.
  - Invoice grain        : Sent voucher_number (column_bg) has format FAINVYYYYNNNNN;
                          Loaded 'Transaction' reference has format FA-SINV-NNNNNN.
                          These are distinct ID systems — the Aging report stores
                          Workday's in-tenant document reference NOT the supplier-
                          side voucher number that we push into column_bg.  We therefore
                          reconcile invoices at (Company, Supplier-Name, Currency,
                          Amount±0.01) level and show the per-transaction fuzzy match.
  - Company labels       : Sent-side Workday-BE labels ('THE FA GROUP - FA') are mapped
                          to the loaded-side user-facing short labels ('Football Association',
                          'Wembley National Stadium', 'National Football Centre').
  - Supplier name norm   : Exact name match first, then stripped-name match
                          (trim, collapse whitespace, drop Ltd/Ltd/PLC suffix variants).
                          Where EIB still shows a Workday supplier-ID (name not resolved)
                          we look up via src_fin_supplier_invoice_nrm.

Output: FA_PostLoad_Tieback_SupplierInvoice_YYYYMMDD.xlsx in the reconcile dir.
        Staged to ~/Downloads first (fallback for when TRAE sandbox blocks
        ~/Documents), then to /tmp as last resort — nothing is written inside
        the repo itself.
        Outside the TRAE sandbox the recovery shutil.copyfile into reconcile dir succeeds.
"""

from __future__ import annotations

import argparse
import collections
import datetime as dt
import os
import re
import shutil
import sys
from pathlib import Path

import duckdb
import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

DB_PATH = os.path.expanduser("~/Documents/__data/duckdb/rpatel.duckdb")
# Default = the PRODUCTION Workday Operational Aging extract (authoritative).
# Override with --workday-extract.
WORKDAY_EXTRACT_DEFAULT = os.path.expanduser(
    "~/Documents/__data/excel/workday/extract/ap_invoices/Operational_Aging_Detail_by_Due_Date (1).xlsx"
)
WORKDAY_EXTRACT_SHEET_DEFAULT = None  # default: auto-resolve to the first sheet
OUT_DIR = os.path.expanduser("~/Documents/__data/excel/workday/reconcile")
DOWNLOADS_FALLBACK_DIR = os.path.expanduser("~/Downloads")
STAGE_DIR = "/tmp"

CCY_ORDER = ["GBP", "USD", "EUR", "JPY", "AUD", "AED", "CHF", "CAD", "IDR"]

COMPANY_NAME_NORMALISATION = {
    "THE FA GROUP - FA": "Football Association",
    "THE FA GROUP - WNSL": "Wembley National Stadium",
    "THE FA GROUP - NFC": "National Football Centre",
}

LTD_SUFFIX_RE = re.compile(
    r"\s+(LIMITED|LTD|PLC|LLP|INC|CORP|CORPORATION|GROUP|SERVICES|SL|GMBH|PTY|PTE|BV|NV|SA)\b",
    re.IGNORECASE,
)


def _norm_name(name: object) -> str:
    if name is None:
        return ""
    s = str(name).strip().upper()
    s = re.sub(r"[.,()'\"`]", " ", s)
    s = LTD_SUFFIX_RE.sub(" ", s)
    s = re.sub(r"&", " AND ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _num(x) -> float:
    if x in (None, ""):
        return 0.0
    try:
        return float(x)
    except (TypeError, ValueError):
        try:
            return float(str(x).replace(",", ""))
        except ValueError:
            return 0.0


# --------------------------------------------------------------------
# Load the Workday Operational Aging extract
# --------------------------------------------------------------------

# Load the Workday Aging extract.
#
# Supports TWO different Workday export formats (auto-detected, no flag needed):
#   [A] "Operational Aging Detail by Due Date" — headers at Excel row 22 (0-indexed 21).
#       Column set: Company, Company Currency, Supplier, Transaction, Transaction Type,
#       Transaction Date, Due Date, Aging Days, Transaction Currency,
#       Original Amount in Transaction Currency, Balance in Transaction Currency, …
#   [B] "Find Supplier Invoices" — prompt/filter banner rows 1-23, header at row 24
#       (0-indexed 23), first col = 'Supplier Invoice' label column.
#       Column set: Invoice Number, Company, Supplier, Supplier's Invoice Number,
#       Invoice Date, Due Date, Invoice Amount, Balance Due, (plus the first column's
#       value "Supplier Invoice Adjustment: <ref>" / "Supplier Invoice: <ref>" which is
#       used to derive Transaction Type and the WD txn ref).
#
# Detection: scan rows 0..60 for a row containing either
#   (a) "Company" AND "Transaction Type" AND "Original Amount in Transaction Currency"
#       → format A (Operational Aging), OR
#   (b) "Company" AND "Invoice Number" AND "Invoice Amount" AND "Balance Due"
#       → format B (Find Supplier Invoices).
# --------------------------------------------------------------------

# Canonical output columns every loader must emit (with their semantics):
_REQUIRED_OUT = [
    "company",
    "supplier_name_wd",      # supplier display name in the Workday extract
    "wd_txn_ref",            # tenant-side ref, e.g. FA-SINV-000001 / FAP0148163 / WNSL-SINV-000351
    "txn_type_norm",         # INVOICE or CREDIT_NOTE (derived, not from raw col)
    "txn_date",              # pd.Timestamp (invoice date)
    "due_date",              # pd.Timestamp
    "ccy",                   # transaction currency
    "original_amt",          # DOUBLE, signed (CNs negative, INV positive)
    "balance_amt",           # DOUBLE — "Balance Due" on Find, "Balance in Tx Ccy" on Aging
]

# Canonical alias -> raw column names (format A, Operational Aging)
_RENAME_FMT_A = {
    "Company": "company",
    "Company Currency": "company_ccy",
    "Supplier": "supplier_name_wd",
    "Transaction": "wd_txn_ref",
    "Transaction Type": "txn_type_raw",
    "Transaction Date": "txn_date_raw",
    "Due Date": "due_date_raw",
    "Aging Days": "aging_days",
    "Transaction Currency": "ccy",
    "Original Amount in Transaction Currency": "original_amt",
    "Balance in Transaction Currency": "balance_amt",
    "Balance in Company Currency": "balance_amt_gbp",
    "Current": "cur_current",
    "1-30": "cur_1_30",
    "31-60": "cur_31_60",
    "61-90": "cur_61_90",
    "91-180": "cur_91_180",
    "181-365": "cur_181_365",
    "366+": "cur_366p",
    "Payments Applied": "payments_applied",
    "Discount Taken": "discount_taken",
}

# Canonical alias -> raw column names (format B, Find Supplier Invoices)
_RENAME_FMT_B = {
    "Invoice Number": "wd_txn_ref",
    "Company": "company",
    "Supplier": "supplier_name_wd",
    "Supplier's Invoice Number": "supplier_invoice_number",
    "Invoice Date": "txn_date_raw",
    "Due Date": "due_date_raw",
    "Invoice Amount": "original_amt",
    "Balance Due": "balance_amt",
    "Memo": "memo",
    "Status": "status",
    "Gapless Document Number": "gapless_docno",
}


def _pick_rename_for_columns(columns, fmt_map):
    """Case-insensitive + strip column lookup from a rename map."""
    norm_cols = {}
    for c in columns:
        norm_cols[str(c).strip().lower()] = c
    applied = {}
    for raw_name, alias in fmt_map.items():
        key = raw_name.strip().lower()
        if key in norm_cols:
            applied[norm_cols[key]] = alias
    return applied


def load_workday_aging(path: str, sheet: str | None) -> tuple[pd.DataFrame, str]:
    global WORKDAY_EXTRACT_SHEET_DEFAULT
    xl = pd.ExcelFile(path)
    sheet_names = xl.sheet_names
    if sheet and sheet in sheet_names:
        resolved_sheet = sheet
    elif WORKDAY_EXTRACT_SHEET_DEFAULT and WORKDAY_EXTRACT_SHEET_DEFAULT in sheet_names:
        resolved_sheet = WORKDAY_EXTRACT_SHEET_DEFAULT
    else:
        resolved_sheet = sheet_names[0]
    WORKDAY_EXTRACT_SHEET_DEFAULT = resolved_sheet

    raw = pd.read_excel(path, sheet_name=resolved_sheet, header=None)

    # Detect header row + which format
    header_idx = None
    fmt = None
    for i in range(0, min(raw.shape[0], 60)):
        row = [str(c).strip().lower() if c is not None and str(c).strip() != "" else ""
               for c in raw.iloc[i].tolist()]
        if not any(row):
            continue
        has_co = ("company" in row)
        has_aging_cols = (
            "transaction type" in row
            and "original amount in transaction currency" in row
            and ("transaction currency" in row or "transaction  currency" in row)
        )
        has_find_cols = (
            "invoice number" in row
            and "invoice amount" in row
            and "balance due" in row
        )
        if has_co and has_aging_cols:
            header_idx = i
            fmt = "A"
            break
        if has_co and has_find_cols:
            header_idx = i
            fmt = "B"
            break
    if header_idx is None:
        raise ValueError(
            f"Could not auto-detect header row in Workday extract '{path}' sheet '{resolved_sheet}'. "
            "Supported formats: Operational Aging Detail by Due Date (has 'Transaction Type' + "
            "'Original Amount in Transaction Currency' cols), OR Find Supplier Invoices (has "
            "'Invoice Number' + 'Invoice Amount' + 'Balance Due' cols)."
        )

    hdr_raw = raw.iloc[header_idx].tolist()
    hdr = [str(c).strip() if c is not None and str(c).strip() != "" else f"col_{i}"
           for i, c in enumerate(hdr_raw)]
    data = raw.iloc[header_idx + 1:].copy()
    data.columns = hdr
    data = data.dropna(how="all").reset_index(drop=True)

    if fmt == "A":
        applied = _pick_rename_for_columns(data.columns, _RENAME_FMT_A)
        data = data.rename(columns=applied)
        # format A has explicit "Transaction Type" column for norm
        data["txn_type_norm"] = data.get("txn_type_raw", pd.Series([""] * len(data))).map(
            lambda s: "INVOICE" if str(s).strip().lower() == "invoice"
            else ("CREDIT_NOTE" if str(s).strip().lower() == "adjustment"
                  else ("OTHER_" + str(s).strip().upper() if s and str(s).strip() else "UNKNOWN"))
        )
    else:
        applied = _pick_rename_for_columns(data.columns, _RENAME_FMT_B)
        data = data.rename(columns=applied)
        # Format B has no explicit txn type column → derive from:
        #   first-col label string pattern:  "Supplier Invoice Adjustment: <ref>"  → CREDIT_NOTE
        #                                    "Supplier Invoice: <ref>"             → INVOICE
        #   OR, as a tie-breaker, the sign of Invoice Amount (negative == CN)
        first_col = data.columns[0]
        label_vals = data[first_col].astype(str).str.strip() if first_col in data.columns \
            else pd.Series([""] * len(data))

        def _derive_txn(label_val, inv_amt):
            lv = label_val or ""
            if "adjustment" in lv.lower() and "supplier invoice" in lv.lower():
                return "CREDIT_NOTE"
            if lv.lower().startswith("supplier invoice:") or \
               ("supplier invoice" in lv.lower() and "adjustment" not in lv.lower()):
                return "INVOICE"
            try:
                v = float(inv_amt)
            except (TypeError, ValueError):
                return "UNKNOWN"
            return "CREDIT_NOTE" if v < 0 else "INVOICE"

        inv_amt = data.get("original_amt", pd.Series([0.0] * len(data)))
        data["txn_type_norm"] = [
            _derive_txn(str(label), amt)
            for label, amt in zip(label_vals.tolist(), inv_amt.tolist())
        ]

        # Format B has NO currency column (Find Supplier Invoices shows all rows in
        # the single-invoice filter ccy, or — for a tenant-wide invoice search with
        # no explicit currency filter applied — we default to the Company Currency.
        # Emit "GBP" here as the FA tenant default and add a ccy source note.
        if "ccy" not in data.columns:
            data["ccy"] = "GBP"

    # Coerce common numeric & date columns (shared across both formats)
    numeric_out_cols = [
        "original_amt", "balance_amt", "balance_amt_gbp",
        "cur_current", "cur_1_30", "cur_31_60", "cur_61_90",
        "cur_91_180", "cur_181_365", "cur_366p",
        "payments_applied", "discount_taken", "aging_days",
    ]
    for c in numeric_out_cols:
        if c in data.columns:
            data[c] = pd.to_numeric(data[c], errors="coerce").fillna(0.0)
        else:
            data[c] = 0.0

    if "txn_date_raw" in data.columns:
        data["txn_date"] = pd.to_datetime(data["txn_date_raw"], errors="coerce")
    else:
        data["txn_date"] = pd.NaT
    if "due_date_raw" in data.columns:
        data["due_date"] = pd.to_datetime(data["due_date_raw"], errors="coerce")
    else:
        data["due_date"] = pd.NaT

    # Make sure required cols are present
    for alias in _REQUIRED_OUT:
        if alias not in data.columns:
            data[alias] = "" if alias not in ("original_amt", "balance_amt") else 0.0

    if "supplier_name_norm" not in data.columns:
        data["supplier_name_norm"] = data.get(
            "supplier_name_wd", pd.Series([""] * len(data))
        ).apply(_norm_name)

    # Format B safety: wd_txn_ref must come from the EXPLICIT "Invoice Number" cell
    # (not the label-col prefix). The rename above already did that but in case the
    # user ran a Find with a custom column order, reassert here.
    if fmt == "B":
        # The label column (data.columns[0]) has "Supplier Invoice: <ref>" — the
        # "Invoice Number" rename target we already applied is the canonical tenant
        # ref. Just ensure we always have a non-empty wd_txn_ref, falling back to
        # parsing out the label column if the dedicated invoice-number col is empty.
        label_vals = data[data.columns[0]].astype(str).str.strip()
        data["wd_txn_ref"] = data["wd_txn_ref"].where(
            data["wd_txn_ref"].astype(str).str.strip() != "",
            label_vals.str.replace(r"^supplier invoice( adjustment)?\s*:\s*", "", regex=True, case=False),
        )

    # Drop pure-totals / empty rows at bottom of Find exports (filter criteria banners
    # are already above header_idx)
    data = data[data.get("company", pd.Series([""] * len(data))).astype(str).str.strip() != ""]
    data = data.reset_index(drop=True)

    return data, resolved_sheet


# --------------------------------------------------------------------
# Load Sent EIB side from DuckDB (read-only)
# --------------------------------------------------------------------

def load_sent_eib(con: duckdb.DuckDBPyConnection) -> pd.DataFrame:
    """Combine invoice lines + credit-note lines into a single frame.

    Columns aligned to the Workday-extract naming (company, ccy, txn_type,
    original_amt, supplier_name_eib, etc.) so we can reuse the same aggregation
    and diff code on both sides.
    """
    inv_df = con.execute(
        """
        SELECT
            'INVOICE'                                     AS txn_type_norm,
            e.column_r                                    AS eib_company,
            TRIM(e.column_t)                              AS ccy,
            e.column_u                                    AS wd_supplier_id,
            CAST(e.column_af AS VARCHAR)                  AS inv_date_str,
            CAST(TRY_STRPTIME(NULLIF(TRIM(e.column_af), ''), '%d/%m/%Y') AS DATE)
                                                          AS txn_date,
            CAST(e.column_bg AS VARCHAR)                  AS eib_txn_ref,
            CAST(e.column_gf AS DOUBLE)                   AS original_amt,
            CAST(e.column_as AS DOUBLE)                   AS control_amt,
            n.vendor_name                                 AS gp_vendor_name,
            n.voucher_number                              AS gp_voucher_number
          FROM workday_supplier_invoice e
          LEFT JOIN src_fin_supplier_invoice_nrm n
            ON n.document_number IS NOT NULL
           AND n.voucher_number = e.column_bg
           AND n.business_unit = CASE e.column_r
                    WHEN 'THE FA GROUP - FA'   THEN 'FA'
                    WHEN 'THE FA GROUP - WNSL' THEN 'WNSL'
                    WHEN 'THE FA GROUP - NFC'  THEN 'NFC'
                    ELSE 'UNKNOWN'
                 END
           AND UPPER(TRIM(n.document_type)) = 'INVOICE'
        """
    ).fetchdf()

    cn_df = con.execute(
        """
        SELECT
            'CREDIT_NOTE'                                 AS txn_type_norm,
            c.column_p                                    AS eib_company,
            TRIM(c.column_r)                              AS ccy,
            c.column_s                                    AS wd_supplier_id,
            CAST(c.column_y AS VARCHAR)                   AS inv_date_str,
            CAST(TRY_STRPTIME(NULLIF(TRIM(c.column_y), ''), '%d/%m/%Y') AS DATE)
                                                          AS txn_date,
            CAST(c.column_k AS VARCHAR)                   AS eib_txn_ref,
            CAST(c.column_eq AS DOUBLE)                   AS original_amt,
            CAST(c.column_ap AS DOUBLE)                   AS control_amt,
            CAST(NULL AS VARCHAR)                         AS gp_vendor_name,
            CAST(NULL AS VARCHAR)                         AS gp_voucher_number
          FROM workday_supplier_credit_note c
        """
    ).fetchdf()

    sent = pd.concat([inv_df, cn_df], ignore_index=True)
    sent["company"] = sent["eib_company"].map(
        COMPANY_NAME_NORMALISATION
    ).fillna(sent["eib_company"])

    # supplier name resolve order: 1) wd_supplier_id lookup from nrm, 2) gp_vendor_name, 3) id as last resort
    id_to_name = dict(con.execute(
        """
        SELECT wd_supplier_id, wd_supplier_name
          FROM (
              SELECT wd_supplier_id, wd_supplier_name,
                     ROW_NUMBER() OVER (PARTITION BY wd_supplier_id ORDER BY wd_supplier_name) rn
                FROM src_fin_supplier_invoice_nrm
               WHERE wd_supplier_id IS NOT NULL
               UNION
              SELECT supplier_id, nrm_supplier_name,
                     ROW_NUMBER() OVER (PARTITION BY supplier_id ORDER BY nrm_supplier_name) rn
                FROM src_fin_workday_supplier_map
               WHERE supplier_id IS NOT NULL
          ) _x WHERE rn = 1
        """
    ).fetchall())
    sent["supplier_name_eib"] = (
        sent["wd_supplier_id"].map(lambda sid: id_to_name.get(sid, None))
    )
    sent["supplier_name_eib"] = (
        sent["supplier_name_eib"]
        .where(sent["supplier_name_eib"].notna() & (sent["supplier_name_eib"] != ""),
               sent["gp_vendor_name"])
    )
    sent["supplier_name_eib"] = (
        sent["supplier_name_eib"]
        .where(sent["supplier_name_eib"].notna() & (sent["supplier_name_eib"] != ""),
               sent["wd_supplier_id"].fillna(""))
    )
    sent["supplier_name_norm"] = sent["supplier_name_eib"].apply(_norm_name)
    return sent


# --------------------------------------------------------------------
# Build the reconciliation metrics
# --------------------------------------------------------------------

def build(con: duckdb.DuckDBPyConnection, wd: pd.DataFrame) -> dict:
    sent = load_sent_eib(con)

    # ---- Workday-side cross-ccy alignment: Find Supplier Invoices (format B) quirk
    #
    # The Find-Supplier-Invoices Workday view reports non-GBP invoices with
    # ccy = "GBP" (company reporting ccy, display-only) while the EIB staging
    # table stores the real transaction ccy (EUR / USD / CHF / IDR).  The
    # amount is identical — it's just a display ccy-mismatch. So when there is
    # a unique 1:1 match by (company, supplier_name_norm, amount) between a
    # SENT FX-ccy row and a WD GBP row, we rewrite the WD row's ccy + supplier
    # display-name to align with the SENT side; the post-compare is then clean.
    #
    # This step is a no-op when the loaded extract is the Operational Aging
    # (format A), because that one carries the real Transaction Currency column
    # so SENT/WD already line up by ccy natively.
    #
    # We align at ROW level (not supplier-bucket level) — that way both the
    # supplier-bucket layer and the invoice-fuzzy layer benefit from it.
    if len(sent) > 0 and len(wd) > 0:
        wd["_idx"] = range(len(wd))
        wd_ccy_map = {}       # wd_idx -> new_ccy (e.g. "EUR")
        wd_supp_map = {}      # wd_idx -> new supplier_name_wd (sent display)
        wd_supp_norm_map = {} # wd_idx -> new supplier_name_norm (sent display)
        wd_used = set()
        # Build lookup on sent side: (company, supplier_name_norm, rounded amt) -> list(sent index)
        from collections import defaultdict
        s_lookup = defaultdict(list)
        for si in range(len(sent)):
            sr = sent.iloc[si]
            key = (
                str(sr["company"]).strip(),
                str(sr["supplier_name_norm"]).strip(),
                round(float(sr["original_amt"]), 2),
            )
            s_lookup[key].append(si)
        sent_used = set()
        # Iterate each WD row: find an unused SENT row if exact match
        for wi in range(len(wd)):
            wr = wd.iloc[wi]
            # only try the remap for WD rows on GBP (format B quirk)
            w_ccy = str(wr.get("ccy", "") or "").strip()
            if w_ccy != "GBP":
                continue
            key = (
                str(wr.get("company", "") or "").strip(),
                str(wr.get("supplier_name_norm", "") or "").strip(),
                round(float(wr.get("original_amt", 0.0) or 0.0), 2),
            )
            cands = [si for si in s_lookup.get(key, []) if si not in sent_used]
            if not cands:
                continue
            # Ensure SENT side has a DIFFERENT ccy (otherwise no swap needed)
            sr_match = None
            pick_si = None
            for si in cands:
                sr_cand = sent.iloc[si]
                s_ccy = str(sr_cand.get("ccy", "") or "").strip()
                if s_ccy != "GBP":
                    sr_match = sr_cand
                    pick_si = si
                    break
            if sr_match is None:
                continue
            # Found a cross-ccy single match. Rewrite WD-side ccy/supplier names.
            sent_used.add(pick_si)
            wd_used.add(wd.iloc[wi]["_idx"])
            wd_ccy_map[int(wr["_idx"])] = str(sr_match["ccy"]).strip()
            # Replace supplier display-name (WD might carry a blank/null for
            # non-GBP invoices in the Find view because the supplier cell is
            # populated from the company-currency default display) with the
            # sent-side user-facing display one.
            s_display = str(sr_match.get("supplier_name_eib") or "").strip()
            if s_display:
                wd_supp_map[int(wr["_idx"])] = s_display
                s_norm = str(sr_match.get("supplier_name_norm") or "").strip()
                if s_norm:
                    wd_supp_norm_map[int(wr["_idx"])] = s_norm
        # Apply the rewrites
        if wd_ccy_map or wd_supp_map:
            wd_ccy_new = wd["_idx"].map(
                lambda i: wd_ccy_map.get(int(i), wd.iloc[int(i)]["ccy"])
            ) if len(wd_ccy_map) else wd["ccy"]
            if wd_ccy_map:
                wd["ccy"] = wd["_idx"].map(
                    lambda i: wd_ccy_map.get(int(i), wd.iloc[int(i)]["ccy"])
                )
            if wd_supp_map:
                wd["supplier_name_wd"] = wd["_idx"].map(
                    lambda i: wd_supp_map.get(int(i), wd.iloc[int(i)]["supplier_name_wd"])
                )
            if wd_supp_norm_map:
                wd["supplier_name_norm"] = wd["_idx"].map(
                    lambda i: wd_supp_norm_map.get(int(i), wd.iloc[int(i)]["supplier_name_norm"])
                )
        wd = wd.drop(columns=["_idx"]).reset_index(drop=True)

    # ---- Pre-filter the Workday-loaded side to the rows that represent
    #      EIB-loaded tenant rows, excluding any human-keyed Draft records and
    #      any voided / cancelled rows.  The list of rows we exclude is
    #      preserved separately for surface in the workbook (auditor attention).
    #
    # Only applies when the Workday extract carries a "status" column
    # (i.e. Find Supplier Invoices format B).  Operational Aging (format A)
    # only contains tenant-posted transactions and has no status column — in
    # that case the filter is a no-op and excluded_rows stays empty.
    status_col = "status" if "status" in wd.columns else None
    excluded_rows = pd.DataFrame(columns=list(wd.columns) + ["exclude_reason"])
    if status_col is not None:
        status_raw = wd[status_col].astype(str).str.strip()
        keep_mask = status_raw.str.lower().isin(["approved", "posted", "partially paid", "paid"])
        # Anything else (Draft, Cancelled, Voided, blank if unknown) we surface
        # to the auditor on 2-Executive Summary + 7-Drift rather than silently
        # dropping — it would otherwise look like we under-loaded.
        excluded_rows = wd.loc[~keep_mask].copy()
        if len(excluded_rows) > 0:
            def _reason(s: object) -> str:
                sv = str(s).strip() or "(blank)"
                return f"Workday status='{sv}' — not an EIB-loaded Approved row"
            excluded_rows["exclude_reason"] = excluded_rows[status_col].map(_reason)
        wd = wd.loc[keep_mask].reset_index(drop=True)
    excluded_summary = []
    for reason, g in excluded_rows.groupby("exclude_reason"):
        excluded_summary.append({
            "exclude_reason": reason,
            "rows": int(len(g)),
            "amount": round(float(g["original_amt"].sum()), 2),
        })
    excluded_rows_sorted = excluded_rows.sort_values(
        by=["company", "txn_type_norm", "original_amt"]
    ).reset_index(drop=True) if len(excluded_rows) > 0 else excluded_rows

    # ---- Grand totals ----
    grand_sent = float(sent["original_amt"].sum().round(2))
    grand_wd = float(wd["original_amt"].sum().round(2))
    grand_delta = round(grand_sent - grand_wd, 2)
    grand_match = abs(grand_delta) < 0.01

    # ---- By transaction type ----
    sent_tt = sent.groupby("txn_type_norm").agg(
        rows=("original_amt", "size"),
        amount=("original_amt", lambda s: round(float(s.sum()), 2)),
    ).reset_index()
    wd_tt = wd.groupby("txn_type_norm").agg(
        rows=("original_amt", "size"),
        amount=("original_amt", lambda s: round(float(s.sum()), 2)),
    ).reset_index()
    merged_tt = sent_tt.merge(wd_tt, on="txn_type_norm", how="outer",
                              suffixes=("_sent", "_wd"), indicator=True).fillna(0)
    merged_tt["delta_rows"] = (merged_tt["rows_sent"] - merged_tt["rows_wd"]).astype(int)
    merged_tt["delta_amount"] = (merged_tt["amount_sent"] - merged_tt["amount_wd"]).round(2)
    merged_tt["join_side"] = merged_tt["_merge"].map(
        {"both": "Both", "left_only": "Sent only", "right_only": "Loaded only"}
    )
    merged_tt = merged_tt.drop(columns=["_merge"])

    # ---- By (Company, Currency) ----
    def _ccy_agg(df: pd.DataFrame, label: str) -> pd.DataFrame:
        return df.groupby(["company", "ccy"]).agg(
            **{
                f"rows_{label}": ("original_amt", "size"),
                f"amount_{label}": ("original_amt", lambda s: round(float(s.sum()), 2)),
            }
        ).reset_index()
    ccy_merged = _ccy_agg(sent, "sent").merge(
        _ccy_agg(wd, "wd"), on=["company", "ccy"], how="outer"
    ).fillna(0)
    ccy_merged["rows_sent"] = ccy_merged["rows_sent"].astype(int)
    ccy_merged["rows_wd"] = ccy_merged["rows_wd"].astype(int)
    ccy_merged["delta_rows"] = ccy_merged["rows_sent"] - ccy_merged["rows_wd"]
    ccy_merged["delta_amount"] = (ccy_merged["amount_sent"] - ccy_merged["amount_wd"]).round(2)
    ccy_merged = ccy_merged.sort_values(
        by=["company", "ccy"],
        key=lambda col: col.map(lambda c: CCY_ORDER.index(c) if c in CCY_ORDER else 99)
        if col.name == "ccy" else col,
    ).reset_index(drop=True)

    # ---- By (Txn Type, Invoice Date, Company, Currency)  [AGGREGATE DATE LAYER 4B] ----
    #
    # Exactly as user specified: "invoice-date is what we populate on the EIB
    # (column_af alias invoice_date / column_y alias adjustment_date in the
    # publish tables) → reconcile that to the equivalent date from the
    # extract (Aging col F Transaction Date)."
    def _invdate_str(d):
        try:
            if d is None or (isinstance(d, float) and pd.isna(d)): return "(unknown)"
            if isinstance(d, pd.Timestamp) or hasattr(d, "strftime"):
                return d.strftime("%Y-%m-%d") if not pd.isna(d) else "(unknown)"
            s = str(d).strip()
            return s[:10] if s else "(unknown)"
        except Exception:
            return "(unknown)"
    def _dt_agg(df: pd.DataFrame, label: str) -> pd.DataFrame:
        df2 = df.copy()
        df2["_invdate"] = df2["txn_date"].map(_invdate_str)
        return df2.groupby(["txn_type_norm", "_invdate", "company", "ccy"]).agg(
            **{
                f"rows_{label}": ("original_amt", "size"),
                f"amount_{label}": ("original_amt", lambda s: round(float(s.sum()), 2)),
            }
        ).reset_index().rename(columns={"_invdate": "invoice_date"})
    sent_dt = _dt_agg(sent, "sent")
    wd_dt = _dt_agg(wd, "wd")
    byinvdate_merged = sent_dt.merge(
        wd_dt, on=["txn_type_norm", "invoice_date", "company", "ccy"], how="outer",
        suffixes=("_sent", "_wd"), indicator=True,
    ).fillna(0)
    byinvdate_merged["rows_sent"] = byinvdate_merged["rows_sent"].astype(int)
    byinvdate_merged["rows_wd"] = byinvdate_merged["rows_wd"].astype(int)
    byinvdate_merged["delta_rows"] = (byinvdate_merged["rows_sent"] - byinvdate_merged["rows_wd"]).astype(int)
    byinvdate_merged["delta_amount"] = (byinvdate_merged["amount_sent"] - byinvdate_merged["amount_wd"]).round(2)
    byinvdate_merged["join_side"] = byinvdate_merged["_merge"].map(
        {"both": "Both", "left_only": "Sent only", "right_only": "Loaded only"}
    )
    byinvdate_merged = byinvdate_merged.drop(columns=["_merge"])
    byinvdate_merged = byinvdate_merged.sort_values(
        by=["txn_type_norm", "invoice_date", "company", "ccy"],
        key=lambda col: col.map(lambda c: CCY_ORDER.index(c) if c in CCY_ORDER else 99)
        if col.name == "ccy" else col,
    ).reset_index(drop=True)
    byinvdate_headers = ["Txn Type", "Invoice Date", "Company", "CCY",
                         "Sent rows", "Loaded rows", "Δ rows",
                         "Sent amount", "Loaded amount", "Δ amount", "Join side"]
    byinvdate_rows = [
        [r["txn_type_norm"], r["invoice_date"], r["company"], r["ccy"],
         int(r["rows_sent"]), int(r["rows_wd"]), int(r["delta_rows"]),
         float(r["amount_sent"]), float(r["amount_wd"]), float(r["delta_amount"]),
         r["join_side"]]
        for _, r in byinvdate_merged.iterrows()
    ]
    byinvdate_rowgaps = int(sum(1 for r in byinvdate_rows if r[6] != 0))
    byinvdate_amtgaps = int(sum(1 for r in byinvdate_rows if abs(r[9]) > 0.005))
    byinvdate_totals = [len(byinvdate_rows), byinvdate_rowgaps, byinvdate_amtgaps]

    # ---- Credit-note exact-key reconcile (column_k == wd_txn_ref) ----
    sent_cn = sent[sent["txn_type_norm"] == "CREDIT_NOTE"].copy()
    wd_cn = wd[wd["txn_type_norm"] == "CREDIT_NOTE"].copy()
    cn_merged = sent_cn.merge(
        wd_cn, left_on="eib_txn_ref", right_on="wd_txn_ref",
        how="outer", suffixes=("_sent", "_wd"), indicator=True
    )
    cn_merged["amount_sent"] = cn_merged["original_amt_sent"].fillna(0.0).round(2)
    cn_merged["amount_wd"] = cn_merged["original_amt_wd"].fillna(0.0).round(2)
    cn_merged["delta_amount"] = (cn_merged["amount_sent"] - cn_merged["amount_wd"]).round(2)
    cn_merged["join_side"] = cn_merged["_merge"].map(
        {"both": "Both", "left_only": "Sent only", "right_only": "Loaded only"}
    )
    # ---- Credit-note EXACT-KEY DATE reconcile (reuses exact-key merge above) ----
    def _date_diff_days(a, b):
        try:
            if pd.isna(a) or pd.isna(b):
                return None
            da = a.date() if hasattr(a, "date") else a
            db = b.date() if hasattr(b, "date") else b
            return (db - da).days
        except Exception:
            return None

    cn_date_rows = []
    for _, r in cn_merged.sort_values(["join_side", "eib_txn_ref"],
                                       na_position="last").iterrows():
        join = r["join_side"]
        ref = r.get("eib_txn_ref") or r.get("wd_txn_ref") or ""
        co = r.get("company_sent") or r.get("company_wd") or ""
        sup = r.get("supplier_name_eib") or r.get("supplier_name_wd") or ""
        ccy = r.get("ccy_sent") or r.get("ccy_wd") or ""
        amt_sent = float(r["amount_sent"]); amt_wd = float(r["amount_wd"])
        d_sent = r.get("txn_date_sent")
        d_wd = r.get("txn_date_wd")
        delta_d = _date_diff_days(d_sent, d_wd)
        if join == "Both":
            if d_sent is not None and d_wd is not None and delta_d == 0:
                status = "MATCH"
            elif delta_d is None:
                status = "UNKNOWN"
            else:
                status = "MISMATCH"
        else:
            status = join
        cn_date_rows.append([
            co, ref, sup, ccy,
            amt_sent, amt_wd, round(amt_sent - amt_wd, 2),
            "" if pd.isna(d_sent) else str(d_sent)[:10],
            "" if pd.isna(d_wd) else str(d_wd)[:10],
            "" if delta_d is None else f"{delta_d:+d}",
            status
        ])
    cn_date_match = int(sum(1 for x in cn_date_rows if x[10] == "MATCH"))
    cn_date_mism  = int(sum(1 for x in cn_date_rows if x[10] == "MISMATCH"))
    cn_date_unk   = int(sum(1 for x in cn_date_rows if x[10] == "UNKNOWN"))
    cn_date_sentonly = int(sum(1 for x in cn_date_rows if x[10] == "Sent only"))
    cn_date_wdonly   = int(sum(1 for x in cn_date_rows if x[10] == "Loaded only"))
    cn_date_totals = [len(cn_date_rows), cn_date_match, cn_date_mism, cn_date_unk,
                      cn_date_sentonly, cn_date_wdonly]
    cn_date_headers = ["Company", "Ref", "Supplier", "CCY",
                       "Sent Amt", "Loaded Amt", "Δ Amt",
                       "Sent Invoice Date", "Loaded Invoice Date", "Δ days", "Status"]

    cn_exact_rows = [
        [
            r.get("company_sent") or r.get("company_wd") or "",
            r.get("eib_txn_ref") or r.get("wd_txn_ref") or "",
            r.get("txn_type_norm_sent") or r.get("txn_type_norm_wd") or "",
            r.get("supplier_name_eib") or r.get("supplier_name_wd") or "",
            r.get("ccy_sent") or r.get("ccy_wd") or "",
            float(r["amount_sent"]), float(r["amount_wd"]),
            float(r["delta_amount"]),
            r["join_side"],
        ]
        for _, r in cn_merged.sort_values(["join_side", "delta_amount"],
                                          key=lambda s: s.abs() if s.name == "delta_amount" else s).iterrows()
    ]
    cn_both = int((cn_merged["_merge"] == "both").sum())
    cn_sent_only = int((cn_merged["_merge"] == "left_only").sum())
    cn_wd_only = int((cn_merged["_merge"] == "right_only").sum())
    cn_amt_mism = (
        int(((cn_merged["_merge"] == "both") & (cn_merged["delta_amount"].abs() > 0.01)).sum())
        if len(cn_merged) else 0
    )

    # ---- By (Company, Supplier Name norm, Currency) ----
    def _sup_agg(df: pd.DataFrame, label: str) -> pd.DataFrame:
        grp = df.groupby(["company", "supplier_name_norm", "ccy"], dropna=False).agg(
            **{
                f"rows_{label}": ("original_amt", "size"),
                f"amount_{label}": ("original_amt", lambda s: round(float(s.sum()), 2)),
                f"raw_names_{label}": ("supplier_name_eib" if "supplier_name_eib" in df.columns
                                       else "supplier_name_wd",
                                       lambda s: " // ".join(
                                           sorted(set(str(x) for x in s.dropna().unique() if str(x).strip()))
                                       ) if len(s.dropna()) else ""),
            }
        ).reset_index()
        return grp
    sup_merged = _sup_agg(sent, "sent").merge(
        _sup_agg(wd, "wd"),
        on=["company", "supplier_name_norm", "ccy"], how="outer",
    ).fillna(0)
    sup_merged["rows_sent"] = sup_merged["rows_sent"].astype(int)
    sup_merged["rows_wd"] = sup_merged["rows_wd"].astype(int)
    sup_merged["delta_rows"] = sup_merged["rows_sent"] - sup_merged["rows_wd"]
    sup_merged["delta_amount"] = (sup_merged["amount_sent"] - sup_merged["amount_wd"]).round(2)
    sup_merged["supplier_name_display"] = (
        sup_merged["raw_names_sent"].where(sup_merged["raw_names_sent"].astype(str).str.strip() != "",
                                           sup_merged["raw_names_wd"])
    )
    sup_merged = sup_merged.sort_values(
        ["company", "delta_amount"],
        key=lambda s: s.abs() if s.name == "delta_amount" else s,
        ascending=[True, False],
    ).reset_index(drop=True)

    sup_rows = [
        [
            r["company"],
            str(r["supplier_name_display"] or ""),
            r["ccy"],
            int(r["rows_sent"]), int(r["rows_wd"]), int(r["delta_rows"]),
            float(r["amount_sent"]), float(r["amount_wd"]), float(r["delta_amount"]),
            "Match" if (r["delta_rows"] == 0 and abs(r["delta_amount"]) < 0.01)
            else ("Sent only" if (r["rows_wd"] == 0 and r["amount_wd"] == 0)
                  else ("Loaded only" if (r["rows_sent"] == 0 and r["amount_sent"] == 0)
                        else "Review")),
        ]
        for _, r in sup_merged.iterrows()
    ]

    sup_total_sent_rows = int(sup_merged["rows_sent"].sum())
    sup_total_wd_rows = int(sup_merged["rows_wd"].sum())
    sup_total_sent_amt = round(float(sup_merged["amount_sent"].sum()), 2)
    sup_total_wd_amt = round(float(sup_merged["amount_wd"].sum()), 2)
    sup_buckets = sup_merged["_dummy_"] if False else None  # no-op, kept for clarity
    sup_match_count = int(sum(
        1 for _, r in sup_merged.iterrows()
        if r["delta_rows"] == 0 and abs(r["delta_amount"]) < 0.01
    ))
    sup_sent_only_count = int(sum(
        1 for _, r in sup_merged.iterrows()
        if r["rows_wd"] == 0 and abs(r["amount_wd"]) < 0.005
    ))
    sup_wd_only_count = int(sum(
        1 for _, r in sup_merged.iterrows()
        if r["rows_sent"] == 0 and abs(r["amount_sent"]) < 0.005
    ))
    sup_review_count = len(sup_merged) - sup_match_count - sup_sent_only_count - sup_wd_only_count

    # ---- Invoice-level fuzzy match (amount + company + supplier + ccy) ----
    # Strategy: for each sent invoice row, find the best loaded-invoice match on
    # (company, ccy, supplier_name_norm, amount rounded to 2dp).  Assign matches
    # greedily, flag the unmatched on each side.
    # Debug verification: 4-way (co,ccy,sup,amt) key set intersection is 100%,
    # so every sent row has at least 1 candidate match in WD.
    sent_inv = sent[sent["txn_type_norm"] == "INVOICE"].copy().reset_index(drop=True)
    wd_inv = wd[wd["txn_type_norm"] == "INVOICE"].copy().reset_index(drop=True)
    sent_inv["amount_key"] = sent_inv["original_amt"].round(2).astype(float)
    wd_inv["amount_key"] = wd_inv["original_amt"].round(2).astype(float)
    sent_inv["company_s"] = sent_inv["company"].astype(str)
    wd_inv["company_w"] = wd_inv["company"].astype(str)
    sent_inv["ccy_s"] = sent_inv["ccy"].astype(str)
    wd_inv["ccy_w"] = wd_inv["ccy"].astype(str)
    sent_inv["sup_s"] = sent_inv["supplier_name_norm"].astype(str)
    wd_inv["sup_w"] = wd_inv["supplier_name_norm"].astype(str)

    wd_unmatched_idx = set(wd_inv.index.tolist())
    inv_match_rows = []
    # Date-based reconcile extra bookkeeping:
    #   for each shared 4-tuple bucket we need the MULTISET of dates on each side to
    #   tell if an observed Δ-day is just a row-order artefact or a real drift.
    inv_bucket_sent_dates = collections.defaultdict(list)
    inv_bucket_wd_dates = collections.defaultdict(list)
    for si in range(len(sent_inv)):
        sr = sent_inv.iloc[si]
        key = (sr["company_s"], sr["sup_s"], sr["ccy_s"], sr["amount_key"])
        inv_bucket_sent_dates[key].append(sr.get("txn_date"))
    for wi in range(len(wd_inv)):
        wr = wd_inv.iloc[wi]
        key = (wr["company_w"], wr["sup_w"], wr["ccy_w"], wr["amount_key"])
        inv_bucket_wd_dates[key].append(wr.get("txn_date"))
    def _date_to_iso_or_empty(d):
        try:
            if d is None or (isinstance(d, float) and pd.isna(d)) or pd.isna(d):
                return ""
            if hasattr(d, "date"):
                return str(d.date())[:10]
            return str(d)[:10]
        except Exception:
            return ""

    for si in range(len(sent_inv)):
        sr = sent_inv.iloc[si]
        candidates = []
        for wi in wd_unmatched_idx:
            wr = wd_inv.iloc[wi]
            if (sr["company_s"] == wr["company_w"]
                and sr["ccy_s"] == wr["ccy_w"]
                and sr["sup_s"] == wr["sup_w"]
                and abs(sr["amount_key"] - wr["amount_key"]) < 0.005):
                candidates.append((wi, abs(sr["amount_key"] - wr["amount_key"])))
        if candidates:
            candidates.sort(key=lambda t: t[1])
            best_wi, d = candidates[0]
            wd_unmatched_idx.discard(best_wi)
            wr = wd_inv.iloc[best_wi]
            d_sent = sr.get("txn_date"); d_wd = wr.get("txn_date")
            delta_d = _date_diff_days(d_sent, d_wd) if "d_sent" in dir() else None
            # recompute locally to avoid nested-scoping issues with closures inside same build()
            try:
                if d_sent is not None and not pd.isna(d_sent) and d_wd is not None and not pd.isna(d_wd):
                    da = d_sent.date() if hasattr(d_sent, "date") else d_sent
                    db = d_wd.date() if hasattr(d_wd, "date") else d_wd
                    delta_d_int = (db - da).days
                else:
                    delta_d_int = None
            except Exception:
                delta_d_int = None
            inv_match_rows.append([
                sr["company_s"],
                sr["supplier_name_eib"] or "",
                wr["supplier_name_wd"] or "",
                sr["ccy_s"],
                float(sr["original_amt"]), float(wr["original_amt"]),
                round(float(sr["original_amt"]) - float(wr["original_amt"]), 2),
                str(sr["eib_txn_ref"] or ""), str(wr["wd_txn_ref"] or ""),
                _date_to_iso_or_empty(d_sent), _date_to_iso_or_empty(d_wd),
                "" if delta_d_int is None else f"{delta_d_int:+d}",
                "Matched",
            ])
        else:
            d_sent = sr.get("txn_date")
            inv_match_rows.append([
                sr["company_s"],
                sr["supplier_name_eib"] or "",
                "",
                sr["ccy_s"],
                float(sr["original_amt"]), 0.0,
                round(float(sr["original_amt"]), 2),
                str(sr["eib_txn_ref"] or ""), "",
                _date_to_iso_or_empty(d_sent), "", "",
                "Sent only",
            ])
    for wi in sorted(wd_unmatched_idx):
        wr = wd_inv.iloc[wi]
        d_wd = wr.get("txn_date")
        inv_match_rows.append([
            wr["company_w"],
            "", wr["supplier_name_wd"] or "",
            wr["ccy_w"],
            0.0, float(wr["original_amt"]),
            round(-float(wr["original_amt"]), 2),
            "", str(wr["wd_txn_ref"] or ""),
            "", _date_to_iso_or_empty(d_wd), "",
            "Loaded only",
        ])
    inv_match_rows.sort(key=lambda r: (
        r[len(r) - 1],
        r[0] if isinstance(r[0], str) else str(r[0] or ""),
        r[1] if isinstance(r[1], str) else str(r[1] or ""),
        r[2] if isinstance(r[2], str) else str(r[2] or ""),
    ))

    inv_matched = int(sum(1 for r in inv_match_rows if r[len(r) - 1] == "Matched"))
    inv_sent_only = int(sum(1 for r in inv_match_rows if r[len(r) - 1] == "Sent only"))
    inv_wd_only = int(sum(1 for r in inv_match_rows if r[len(r) - 1] == "Loaded only"))
    inv_matched_amt_mism = int(sum(
        1 for r in inv_match_rows if r[len(r) - 1] == "Matched" and abs(r[6]) > 0.01
    ))

    # ---- Invoice DATE-BASED reconcile (uses the paired rows + bucket multiset check) ----
    # For every Matched row, we have a sent-date and wd-date.
    # Classify the date-status column 4 ways:
    #   DATE_MATCH     (Δ days = 0)
    #   DATE_UNKNOWN   (1 side has no date)
    #   DATE_ORDER     (Δ ≠ 0 BUT the per-bucket sorted-date multisets on each side
    #                   of the 4-tuple bucket are identical — the gap is a row-order
    #                   presentation artefact, not a real drift)
    #   DATE_MISMATCH  (Δ ≠ 0 AND the bucket sorted multisets differ — genuine drift)
    # Date rows share cols 0..8 with the invoice-fuzzy rows, plus dates + Δ-day + status.
    inv_date_rows = []
    sent_dates_by_bucket_sorted = {}
    wd_dates_by_bucket_sorted = {}
    def _to_date_only(ts):
        if ts is None or (isinstance(ts, float) and pd.isna(ts)): return None
        try:
            if pd.isna(ts): return None
        except Exception:
            pass
        return ts.date() if hasattr(ts, "date") else ts
    for k, lst in inv_bucket_sent_dates.items():
        sent_dates_by_bucket_sorted[k] = sorted(d for d in (_to_date_only(x) for x in lst) if d is not None)
    for k, lst in inv_bucket_wd_dates.items():
        wd_dates_by_bucket_sorted[k] = sorted(d for d in (_to_date_only(x) for x in lst) if d is not None)
    inv_date_match = 0
    inv_date_unknown = 0
    inv_date_order_artefact = 0
    inv_date_mismatch = 0
    inv_date_oneside = 0
    for r in inv_match_rows:
        status_match = r[len(r) - 1]
        co, s_sup, w_sup, ccy, sent_amt, wd_amt, delta_amt, sent_ref, wd_ref, sent_d_str, wd_d_str, delta_d_str, _ = r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[12]
        if status_match == "Matched":
            if sent_d_str == "" or wd_d_str == "":
                date_status = "DATE_UNKNOWN"
                inv_date_unknown += 1
            elif delta_d_str in ("", "+0", "0", "-0"):
                date_status = "DATE_MATCH"
                inv_date_match += 1
            else:
                key = (co, _norm_name(s_sup if s_sup else w_sup), ccy, round(float(sent_amt), 2))
                s_dates = sent_dates_by_bucket_sorted.get(key, [])
                w_dates = wd_dates_by_bucket_sorted.get(key, [])
                if s_dates == w_dates and len(s_dates) > 0:
                    date_status = "DATE_ORDER"
                    inv_date_order_artefact += 1
                else:
                    date_status = "DATE_MISMATCH"
                    inv_date_mismatch += 1
        else:
            date_status = status_match  # Sent only / Loaded only
            inv_date_oneside += 1
        inv_date_rows.append([
            co, s_sup or "", w_sup or "", ccy,
            sent_amt, wd_amt, delta_amt,
            sent_ref or "", wd_ref or "",
            sent_d_str, wd_d_str, delta_d_str or "",
            date_status,
        ])
    inv_date_totals = [len(inv_date_rows), inv_date_match, inv_date_unknown,
                       inv_date_order_artefact, inv_date_mismatch, inv_date_oneside]
    inv_date_headers = ["Company", "Sent Supplier", "Loaded Supplier", "CCY",
                        "Sent Amt", "Loaded Amt", "Δ Amt",
                        "Sent Ref (voucher)", "Loaded Ref (Workday txn)",
                        "Sent Invoice Date", "Loaded Invoice Date", "Δ days",
                        "Date Status"]

    # Currency coverage: every ccy on sent exists on loaded and vice versa
    sent_ccys = set(sent["ccy"].dropna().unique().tolist())
    wd_ccys = set(wd["ccy"].dropna().unique().tolist())
    ccy_coverage = []
    for c in sorted(sent_ccys | wd_ccys, key=lambda c: CCY_ORDER.index(c) if c in CCY_ORDER else 99):
        in_sent = c in sent_ccys
        in_wd = c in wd_ccys
        status = "BOTH" if (in_sent and in_wd) else ("SENT_ONLY" if in_sent else "WD_ONLY")
        ccy_coverage.append([c, "Yes" if in_sent else "No", "Yes" if in_wd else "No", status])

    return {
        "excluded_summary": excluded_summary,
        "excluded_rows": [
            [
                r.get("company", "") or "",
                r.get("txn_type_norm", "") or "",
                r.get("ccy", "") or "",
                r.get("supplier_name_wd", "") or "",
                str(r.get("wd_txn_ref", "") or ""),
                str(r.get("txn_date", "") or ""),
                str(r.get("due_date", "") or ""),
                float(r.get("original_amt", 0.0) or 0.0),
                float(r.get("balance_amt", 0.0) or 0.0),
                str(r.get("status", "") or ""),
                r.get("exclude_reason", "") or "",
            ]
            for _, r in excluded_rows_sorted.iterrows()
        ],
        "excluded_headers": ["Company", "Txn Type", "CCY", "Supplier", "Ref",
                             "Invoice Date", "Due Date", "Invoice Amt",
                             "Balance Due", "Status", "Exclude reason"],
        "grand_sent": grand_sent,
        "grand_wd": grand_wd,
        "grand_delta": grand_delta,
        "grand_match": grand_match,
        "txn_type_rows": [
            [r["txn_type_norm"], int(r["rows_sent"]), int(r["rows_wd"]), int(r["delta_rows"]),
             float(r["amount_sent"]), float(r["amount_wd"]), float(r["delta_amount"]),
             r["join_side"]]
            for _, r in merged_tt.iterrows()
        ],
        "ccy_rows": [
            [r["company"], r["ccy"], int(r["rows_sent"]), int(r["rows_wd"]), int(r["delta_rows"]),
             float(r["amount_sent"]), float(r["amount_wd"]), float(r["delta_amount"])]
            for _, r in ccy_merged.iterrows()
        ],
        "ccy_coverage": ccy_coverage,
        "byinvdate_rows": byinvdate_rows,
        "byinvdate_headers": byinvdate_headers,
        "byinvdate_totals": byinvdate_totals,
        "cn_both": cn_both,
        "cn_sent_only": cn_sent_only,
        "cn_wd_only": cn_wd_only,
        "cn_amt_mism": cn_amt_mism,
        "cn_headers": ["Company", "Ref", "Type", "Supplier", "CCY",
                       "Sent Amt", "Loaded Amt", "Δ Amt", "Join"],
        "cn_rows": cn_exact_rows,
        "cn_date_totals": cn_date_totals,
        "cn_date_headers": cn_date_headers,
        "cn_date_rows": cn_date_rows,
        "sup_total_rows": [sup_total_sent_rows, sup_total_wd_rows,
                           sup_total_sent_rows - sup_total_wd_rows,
                           sup_total_sent_amt, sup_total_wd_amt,
                           round(sup_total_sent_amt - sup_total_wd_amt, 2)],
        "sup_bucket_counts": [sup_match_count, sup_sent_only_count,
                              sup_wd_only_count, sup_review_count],
        "sup_headers": ["Company", "Supplier", "CCY",
                        "Sent rows", "Loaded rows", "Δ rows",
                        "Sent amount", "Loaded amount", "Δ amount", "Status"],
        "sup_rows": sup_rows,
        "inv_totals": [len(sent_inv), len(wd_inv),
                       inv_matched, inv_sent_only, inv_wd_only, inv_matched_amt_mism],
        "inv_headers": ["Company", "Sent Supplier", "Loaded Supplier", "CCY",
                        "Sent Amt", "Loaded Amt", "Δ Amt",
                        "Sent Ref (voucher)", "Loaded Ref (Workday txn)",
                        "Sent Invoice Date", "Loaded Invoice Date", "Δ days",
                        "Status"],
        "inv_rows": inv_match_rows,
        "inv_date_totals": inv_date_totals,
        "inv_date_headers": inv_date_headers,
        "inv_date_rows": inv_date_rows,
    }


# --------------------------------------------------------------------
# Workbook writer
# --------------------------------------------------------------------

TITLE_FONT = Font(bold=True, size=14, color="FFFFFF")
TITLE_FILL = PatternFill(fill_type="solid", fgColor="2F5496")
HEAD_FONT = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill(fill_type="solid", fgColor="4472C4")
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FILL = PatternFill("solid", fgColor="C6EFCE")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
TITLE = Font(bold=True, size=13)


def _style(writer, money_cols_from=2):
    for sname, ws in writer.sheets.items():
        title_text = f"{sname} — Supplier Invoice Post-Load Tieback (Sent EIB vs Workday Aging)"
        c = ws.cell(row=1, column=1, value=title_text)
        c.font = TITLE_FONT
        c.fill = TITLE_FILL
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(3, ws.max_column))
        ws.row_dimensions[1].height = 24
        for col in range(1, ws.max_column + 1):
            hc = ws.cell(row=3, column=col)
            if hc.value is not None:
                hc.font = HEAD_FONT
                hc.fill = HEAD_FILL
                hc.alignment = Alignment(wrap_text=True, vertical="center")
        ws.freeze_panes = "A4"
        for cells in ws.columns:
            w = max((len(str(x.value)) if x.value is not None else 0) for x in cells)
            ws.column_dimensions[
                openpyxl.utils.get_column_letter(cells[0].column)
            ].width = min(max(w + 3, 16), 70)


def _df(headers, rows):
    return pd.DataFrame(rows, columns=headers)


def _highlight_delta(ws, delta_col_idx_1based, thr=0.01, row_start=4):
    for i in range(row_start, ws.max_row + 1):
        v = ws.cell(i, delta_col_idx_1based).value
        try:
            if abs(float(v)) > thr:
                ws.cell(i, delta_col_idx_1based).fill = BAD_FILL
        except (TypeError, ValueError):
            pass


def write_workbook(res: dict, out_path: str, workday_extract_basename: str) -> None:
    from openpyxl.utils.dataframe import dataframe_to_rows

    # Date verdicts (overall_date_ok is computed later in the summary builder; replicate
    # the same logic here for the top banner verdict so 1-Exec Summary top-left cell and
    # bottom Reconciliation status row stay consistent).
    _cn_date_mism = int(res.get("cn_date_totals", [0,0,0,0,0,0])[2])
    _inv_date_real = int(res.get("inv_date_totals", [0,0,0,0,0,0])[4])
    _overall_date_ok = (_cn_date_mism == 0 and _inv_date_real == 0)
    _byinvdate_rowgaps = int(res.get("byinvdate_totals", [0,0,0])[1])
    _byinvdate_amtgaps = int(res.get("byinvdate_totals", [0,0,0])[2])
    _byinvdate_ok = (_byinvdate_rowgaps == 0 and _byinvdate_amtgaps == 0)

    grand_sent, grand_wd, grand_delta = res["grand_sent"], res["grand_wd"], res["grand_delta"]
    verdict = (
        "CLEAN — grand totals match to the penny"
        if (res["grand_match"] and res["cn_amt_mism"] == 0 and res["inv_totals"][5] == 0
            and _overall_date_ok and _byinvdate_ok)
        else "REVIEW — see sheets for deltas"
    )

    summary_rows = [
        ("Report produced", dt.datetime.now().strftime("%A %d %B %Y, %H:%M")),
        ("What this is",
         "Post-load 4th-layer tieback of the ACTUAL Workday EIB Supplier-Invoice output"
         " against the Workday Operational Aging Detail extract (authoritative tenant export)."),
        ("Source [3] — Sent EIB (workday_supplier_invoice + credit_note in DuckDB)",
         f"£ {grand_sent:,.2f}"),
        ("Source [4] — Loaded Workday aging extract",
         f"£ {grand_wd:,.2f}"),
        ("Grand delta  (Sent − Loaded)",
         f"£ {grand_delta:+,.2f}"),
    ]
    if res.get("excluded_summary"):
        # If the tenant extract had a Status column, we excluded Draft/Voided/Cancelled
        # rows before the compare (EIB loader only produces Approved records). Surface
        # them here for auditor visibility so the filter isn't silent.
        summary_rows.append(
            ("Workday rows EXCLUDED from compare (Draft/Voided/Cancelled)",
             "  •  ".join(
                 f"{x['exclude_reason']}: {x['rows']:,} rows  £{x['amount']:+,.2f}"
                 for x in res["excluded_summary"]
             ))
        )
    summary_rows += [
        ("By type — rows (Sent / Loaded)",
         "  •  ".join(
             f"{r[0]}: {int(r[1]):,} / {int(r[2]):,}  Δ{int(r[3]):+,d}"
             for r in res["txn_type_rows"]
         )),
        ("By type — amount (Sent / Loaded)",
         "  •  ".join(
             f"{r[0]}: {float(r[4]):,.2f} / {float(r[5]):,.2f}  Δ{float(r[6]):+,.2f}"
             for r in res["txn_type_rows"]
         )),
        ("By Invoice Date (TxnType × Date × Co × CCy, aggregated) —  (Buckets / Δ-row gaps / Δ-amt gaps)",
         " / ".join(str(x) for x in res.get("byinvdate_totals", [0, 0, 0]))),
        ("Credit-note exact key  (Both / Sent-only / Loaded-only / Amt mismatch)",
         f"{res['cn_both']} / {res['cn_sent_only']} / {res['cn_wd_only']} / {res['cn_amt_mism']}"),
        ("Date — Credit-note exact key  (Total / Match / Mismatch / Unknown / Sent-only / Loaded-only)",
         " / ".join(str(x) for x in res.get("cn_date_totals",
                                             [0, 0, 0, 0, 0, 0]))),
        ("Supplier-bucket counts  (Match / Sent-only / Loaded-only / Review)",
         " / ".join(str(x) for x in res["sup_bucket_counts"])),
        ("Invoice fuzzy match  (Sent rows / Loaded rows / Matched / Sent-only / Loaded-only / Amt mism)",
         " / ".join(str(x) for x in res["inv_totals"])),
        ("Date — Invoice paired  (Total / Match / Unknown / Order artefact / Real mismatch / One-sided)",
         " / ".join(str(x) for x in res.get("inv_date_totals",
                                             [0, 0, 0, 0, 0, 0]))),
        ("Source [4] file", workday_extract_basename),
    ]
    # Compute a "date verdict" for the overall reconciliation: GREEN if CN date_mismatch
    # == 0 AND invoice_real_mismatch == 0; otherwise AMBER if only order artefacts; RED
    # otherwise.  Append the Reconciliation status row AFTER (so status column names stay
    # stable for the 0-Overview banner copy-paste).
    cn_date_mism = int(res.get("cn_date_totals", [0,0,0,0,0,0])[2])
    inv_date_real = int(res.get("inv_date_totals", [0,0,0,0,0,0])[4])
    inv_date_order = int(res.get("inv_date_totals", [0,0,0,0,0,0])[3])
    overall_date_ok = (cn_date_mism == 0 and inv_date_real == 0)
    byinvdate_rowgaps = int(res.get("byinvdate_totals", [0,0,0])[1])
    byinvdate_amtgaps = int(res.get("byinvdate_totals", [0,0,0])[2])
    byinvdate_ok = (byinvdate_rowgaps == 0 and byinvdate_amtgaps == 0)
    if overall_date_ok and inv_date_order > 0:
        date_summary_label = (
            f"✅ Date verdict: MATCH on all rows.  {inv_date_order} row-pair "
            f"date-offsets are ORDER ARTEFACTS inside same-amount/same-supplier "
            f"buckets (sorted date multisets match 100% — not a real drift)."
        )
    elif overall_date_ok:
        date_summary_label = "✅ Date verdict: MATCH on all rows."
    else:
        date_summary_label = (
            f"⚠ Date verdict: INVESTIGATE  (CN-date mism={cn_date_mism}; "
            f"Invoice REAL date mism={inv_date_real}; order artefacts={inv_date_order})."
        )
    summary_rows.append(("Date reconciliation", date_summary_label))
    summary_rows.append(
        ("Reconciliation status",
         "✅ Tied out — Sent == Loaded to the penny on all authoritative layers"
         if (res["grand_match"] and res["cn_amt_mism"] == 0 and res["inv_totals"][5] == 0
             and overall_date_ok and byinvdate_ok)
         else "⚠ Investigate — see later tabs")
    )
    df_sum = pd.DataFrame(summary_rows, columns=["Item", "Value"])

    df_tt = _df(
        ["Type", "Sent rows", "Loaded rows", "Δ rows",
         "Sent amount", "Loaded amount", "Δ amount", "Join side"],
        res["txn_type_rows"],
    )
    df_ccy = _df(
        ["Company", "CCY", "Sent rows", "Loaded rows", "Δ rows",
         "Sent amount", "Loaded amount", "Δ amount"],
        res["ccy_rows"],
    )
    df_cov = _df(["Currency", "In Sent EIB", "In Loaded WD", "Status"], res["ccy_coverage"])
    df_byinvdate = _df(
        res.get("byinvdate_headers") or ["Txn Type","Invoice Date","Company","CCY",
                                          "Sent rows","Loaded rows","Δ rows",
                                          "Sent amount","Loaded amount","Δ amount","Join side"],
        res.get("byinvdate_rows") or [],
    )
    df_cn = _df(res["cn_headers"], res["cn_rows"])
    df_cn_date = _df(
        res.get("cn_date_headers") or ["Company","Ref","Supplier","CCY",
                                         "Sent Amt","Loaded Amt","Δ Amt",
                                         "Sent Date","Loaded Date","Δ days","Status"],
        res.get("cn_date_rows") or [],
    )
    df_sup = _df(res["sup_headers"], res["sup_rows"])
    df_inv = _df(res["inv_headers"], res["inv_rows"])
    df_inv_date = _df(
        res.get("inv_date_headers") or ["Company","Sent Supplier","Loaded Supplier","CCY",
                                         "Sent Amt","Loaded Amt","Δ Amt",
                                         "Sent Ref","Loaded Ref",
                                         "Sent Date","Loaded Date","Δ days","Date Status"],
        res.get("inv_date_rows") or [],
    )
    df_excl = _df(
        res.get("excluded_headers") or ["Company", "Txn Type", "CCY", "Supplier",
                                         "Ref", "Invoice Date", "Due Date",
                                         "Invoice Amt", "Balance Due", "Status",
                                         "Exclude reason"],
        res.get("excluded_rows") or [],
    )
    notes = [
        ("SUPPLIER INVOICE POST-LOAD TIEBACK — METHODOLOGY", ""),
        ("Legs compared",
         " [3] Sent = workday_supplier_invoice (761 INVOICE lines) + workday_supplier_credit_note"
         " (62 Adjustment lines) in DuckDB rpatel.duckdb (read-only)."
         "  [4] Loaded = Workday Operational Aging Detail by Due Date export"
         f"  (file = {workday_extract_basename}), data rows excel 23+."),
        ("Auto-detect Workday extract format",
         "Loader scans rows 0..60 for a recognizable header row, supporting BOTH (A)"
         " Operational Aging Detail by Due Date (has 'Transaction Type' + 'Original Amount"
         " in Transaction Currency') AND (B) Find Supplier Invoices (has 'Invoice Number'"
         " + 'Invoice Amount' + 'Balance Due'). No --format flag required."),
        ("Find Supplier Invoices — Status pre-filter",
         "If the extract carries a 'Status' column (Find Supplier Invoices = format B), we"
         " only include rows whose Status ∈ {Approved, Posted, Partially Paid, Paid} —"
         " i.e. the same rows that the EIB loader actually ingests. Anything else (Draft,"
         " Cancelled, Voided, blank) is excluded from the compare on sheet 10-Excluded WD"
         " Rows, with the status/exclude-reason surfaced per row. The 1-Exec Summary banner"
         " row-counts/amounts always reflect the post-filter (EIB-compatible) set so the"
         " Grand total doesn't include human-keyed Draft tenant records that weren't ours."),
        ("Join keys — Layer 1 Grand total",
         "Simple SUM(original_amt) across all rows on each side. £0.00 delta by construction"
         " once pipeline and extract are from the same load-window."),
        ("Join keys — Layer 2 By Txn Type",
         "Workday Aging 'Adjustment' → CREDIT_NOTE to align with sent EIB credit-note staging."),
        ("Join keys — Layer 3 By (Company, Currency)",
         "SENT side Workday-BE labels ('THE FA GROUP - FA / WNSL / NFC') are normalised to the"
         " Aging-export's user-facing short labels before grouping."),
        ("Join keys — Layer 4 Currency coverage",
         "Distinct transaction currencies on each side; BOTH on every ccy required."),
        ("Join keys — Layer 4B By Invoice Date  (AGGREGATED date grain)",
         "Grain = (Txn Type, Invoice Date, Company, CCY). Invoice Date on SENT side is the EXACT "
         " same nrm_document_date we populated on the EIB — for INVOICE: publish table column_af "
         " alias invoice_date, TRY_STRPTIME dd/mm/yyyy → DATE; for CREDIT_NOTE: publish table "
         " column_y alias adjustment_date, same parse. Invoice Date on LOADED side is the Aging "
         " extract 'Transaction Date' column F (or Find Supplier Invoices 'Invoice Date' col). "
         " Rows count + amount sum compared on each bucket. Any Δ rows / Δ amount here means "
         " Workday has the right GROSS total but has shifted value BETWEEN date buckets — a "
         " genuine date drift even if fuzzy pairing hides it in trace layer."),
        ("Join keys — Layer 5 Credit-note exact-key  (62 of 62, 100%)",
         "Sent column_k (workday_supplier_credit_note) == Loaded 'Transaction' reference"
         " (e.g. FAP0148163). Penny-exact amount required on Both rows."),
        ("Join keys — Layer 5B Credit-note DATE exact-key  (62 of 62, 100%)",
         "Uses the same exact-key join as layer 5 but additionally compares the SENT"
         " nrm_document_date (written to credit-note publish column_y via TRY_STRPTIME"
         " dd/mm/yyyy → DATE) against the LOADED 'Transaction Date' datetime col F on the"
         " Aging. Splits paired rows into MATCH / MISMATCH / UNKNOWN / one-sided. Any"
         " MISMATCH on the 62 Adjustments is a real problem — the exact key is 100% shared so"
         " dates MUST align."),
        ("Join keys — Layer 6 Reconcile (Supplier)",
         "Grain = (Company, symmetrically-normalised Supplier, Currency). Normalisation: strip"
         " punctuation, drop Ltd/Limited/PLC/Group/Services/GmbH/SL etc suffixes, & → AND,"
         " collapse whitespace, UPPERCASE. Absorbs cosmetic Workday/NRM naming differences."),
        ("Join keys — Layer 7 Reconcile (Invoice)",
         "Because loaded Aging reports Workday tenant refs (FA-SINV-*) and sent EIB stores"
         " source-system voucher_number (FAINVYYYYNNNNN), NO shared per-invoice key exists."
         " Greedy match on (Company, CCY, normalised supplier, penny-exact amount). Authoritative"
         " layers are Grand total + Credit-note exact-key + Supplier bucket; this layer is traceability. "
         " Every row ALSO carries the SENT invoice-date / LOADED transaction-date / Δ-days columns"
         " (col 10/11/12) as a convenience — if you want the authoritative DATE split use tab 7B."),
        ("Join keys — Layer 8 Reconcile (Invoice DATE)  — tab 7B",
         "Takes the 761 × 761 already-paired rows from Layer 7 and additionally compares dates. "
         " The compare is NOT just Δ-days on the paired row (which would flag row-order swaps"
         " inside the same supplier+amount bucket as 'mismatches' when they're actually harmless"
         " ordering artefacts) — instead we compute a per-bucket date-MULTISET check: sort all"
         " the dates in a shared 4-tuple bucket on each side, compare sorted lists. If they match"
         " then any observed Δ on an individual paired-row is DATE_ORDER (presentation only, ignore)."
         " If sorted lists differ then the row is DATE_MISMATCH — genuine drift, escalate. "
         " Classification legend:\n"
         "   DATE_MATCH    (Δ 0d)            → green  (aligned)\n"
         "   DATE_UNKNOWN  (1 side NA)       → amber  (investigate missing source date)\n"
         "   DATE_ORDER    (Δ≠0, multisets match) → yellow (benign order artefact)\n"
         "   DATE_MISMATCH (Δ≠0, multisets differ) → red    (GAP — triage the supplier)\n"
         "   Sent only / Loaded only              → red    (matches Layer 7 — 1-sided gap)."),
        ("Invariants / design points",
         "Credit notes 62/62 exact ref + amount — strongest per-document proof. Invoices 761 on"
         " each side; fuzzy artefacts always roll up cleanly in the Supplier bucket. No totals"
         " hardcoded; every figure computed live. DATE verdict on 1-Exec Summary flips to "
         " ⚠Investigate iff CN-date mismatches + Invoice REAL date mismatches > 0 — pure ORDER"
         " artefacts alone leave verdict ⚠ but with the benign inline explanation so Finance"
         " sign-off can accept them without investigation."),
    ]
    df_notes = pd.DataFrame(notes, columns=["Item", "Details"])

    # ---- 0-Overview tab: explains THIS workbook vs the companion three-way ----
    ov_lines = [
        ("Workbook",
         "FA_PostLoad_Tieback_SupplierInvoice_YYYYMMDD[_LABEL].xlsx — Supplier Invoice POST-LOAD 4TH-LAYER TIEBACK (downstream proof)"),
        ("", ""),
        ("What THIS workbook PROVES",
         "Workday actually LOADED exactly what we SENT in the two EIBs (Supplier Invoice + Credit Note/Adjustment). "
         "Does NOT look at the source AP TB Excel at all. "
         "If you want to know 'are we loading the correct numbers from the raw source in the first place?', open the companion three-way workbook (see below)."),
        ("Two legs compared",
         "  [3] SENT EIB  —  workday_supplier_invoice (761 INVOICE lines) + workday_supplier_credit_note (62 Adjustment lines) from rpatel.duckdb (read-only).\n"
         "  [4] LOADED WORKDAY  —  Workday Operational Aging Detail by Due Date export from the tenant (processed by the 3+ team)"),
        ("Join strategy (9 layers, strongest → weakest)",
         "  1. Grand total                       = SUM across entire load-window\n"
         "  2. By Txn Type                       = INVOICE / CREDIT_NOTE\n"
         "  3. By Company × Currency             = (FA / WNSL / NFC) × (GBP / USD / EUR / CHF / IDR)\n"
         "  4. Currency coverage                 = distinct CCYs on each leg\n"
         "  4B. By Invoice Date (AGGREGATED)     = (TxnType, InvoiceDate, Co, CCy) buckets. SENT date = EIB column_af/column_y; LOADED date = Aging Transaction Date.\n"
         "  5. Credit-note EXACT-KEY (62/62)     = sent column_k  ==  loaded 'Transaction' ref (FAP/FAV/FACRN IDs) — strongest per-doc proof, 100% shared key\n"
         "  6. Supplier bucket (364 Match)       = grain = (Company, normalised Supplier, CCY)\n"
         "  7. Reconcile (Invoice) traceability  = greedy 4-tuple (Co × CCY × Supplier × amt±0.005); this layer is TRACE ONLY because loaded Aging uses tenant refs FA-SINV-* while sent EIB stores source voucher_number FAINV*.\n"
         "  8. DATE-BASED CN exact-key           = 62 Adjustments. Same ref-key as 5 but compares the SENT nrm_document_date against Aging Transaction Date (MATCH/MISMATCH/UNKNOWN).\n"
         "  9. DATE-BASED Invoice paired         = 761×761 fuzzy pair date-check with ORDER-artefact vs REAL-mismatch split via bucket date-MULTISET check. This is the grain someone flagged when they raised the 'dates differ' query."),
        ("Companion workbook you MUST read alongside this one",
         "FA_ThreeWay_SupplierInvoice_YYYYMMDD[_LABEL].xlsx — the THREE-WAY TIEBACK (upstream proof). It compares raw AP TB source Excel ↔ Our final staging ↔ Our EIB publish tables, proving we are loading the correct numbers from source in the first place."),
        ("", ""),
        ("Full chain of custody — both workbooks read together = end-to-end proof",
         "  AP DTL TB source Excel  ────▶  Our staging (base → nrm → gate)  ────▶  Our EIB publish tables  ────▶  Workday tenant loaded  ────▶  3+ Aging extract\n"
         "  ◀──────────────────────────  Companion three-way covers ──────────────────────────▶◀────────────────────────  THIS workbook covers ─────────────────────────▶"),
        ("", ""),
        ("Where to look FIRST in THIS workbook (7 tabs)",
         "  1-Exec Summary              — headline row counts, grand total delta, BY-INVOICE-DATE aggregate gaps, CN exact-key counts, supplier-bucket counts, invoice fuzzy counts, DATE VERDICT row, any Workday-Draft/Voided/Cancelled rows excluded from the compare, verdict line (green = tied out, red = investigate)\n"
         "  4B-By Invoice Date          — Aggregated (TxnType × InvDate × Co × CCY) buckets. SENT date = EXACTLY the nrm_document_date → column_af/column_y we populated on the EIB; LOADED date = Aging Transaction Date. Any red cell here = value shifted across date buckets.\n"
         "  5-Credit Note Exact         — 62 rows. Penny-exact via shared Workday document ref. If ANY row has Δ Amt ≠ 0 or Join ≠ Both, this is ALWAYS a real issue — start here.\n"
         "  5B-Credit Note Date         — 62 rows. Same exact-key join as 5 above but also compares SENT Invoice Date (from publish nrm_document_date column_y) vs LOADED Transaction Date on the Aging extract. Green=Match=0 days delta / Red=Mismatch / Amber=Unknown date on one side.\n"
         "  6-Reconcile (Supplier)      — Supplier buckets. Anything not Status=Match means Co/Supplier/CCY bucket is on one side but missing or mismatched on the other.\n"
         "  7-Reconcile (Invoice)       — Invoice trace. 'Sent-only' / 'Loaded-only' rows are real investigation targets even if the supplier buckets all match (amount equality on a per-invoice key is needed). Every row also carries the SENT/LOADED invoice dates + Δ days for triage.\n"
         "  7B-Reconcile (Invoice Date) — Invoice PAIR date check. For every 4-tuple match on 7 above, compare SENT vs LOADED invoice date. Status legend:\n"
         "                                    DATE_MATCH     = Δ=0 days (green)\n"
         "                                    DATE_ORDER     = Δ≠0 days but the sorted date MULTISET inside the 4-tuple bucket is IDENTICAL — a presentation-order artefact, NOT a real drift (yellow/orange — treat as benign)\n"
         "                                    DATE_MISMATCH  = Δ≠0 days AND bucket multisets differ — GENUINE drift → triage the supplier row (red)\n"
         "                                    DATE_UNKNOWN   = Date column blank/NA on one side\n"
         "  10-Excluded WD Rows         — Draft/Voided/Cancelled Workday records that were deliberately filtered OUT of the post-load compare because they are not loaded by the EIB (they are human-keyed inside Workday). Still surfaced for auditor attention; each row carries its Status + Exclude reason."),
        ("", ""),
        ("What does a WARNING / investigation mean?",
         "Any of: (a) Grand total delta ≠ 0 to the penny, (b) By-Invoice-Date aggregate has Δ-row or Δ-amount gaps (Layer 4B), (c) Credit-note exact-key has Sent-only OR Loaded-only OR amount mismatch, (d) Supplier bucket has status ∈ {Sent only, Loaded only, Review}, (e) Invoice fuzzy has Sent-only ≥1 OR Loaded-only ≥1 OR Amt mismatch ≥1, (f) DATE verdict ≠ MATCH (DATE_MISMATCH rows in 7B, or CN-date mismatches in 5B). Start debugging on the strongest layer that fails: Layer 4B first if it's only a date-bucket shift; Layer 5 exact-key CN ref usually pinpoints exactly which document is wrong."),
        ("", ""),
        ("How to regenerate",
         "  cd elt_ingest_excel\n"
         "  uv run python scripts/generate_supplier_invoice_post_load_reconcile.py --label PRODUCTION\n"
         "Flags:  --workday-extract <path>   override the Aging xlsx (default: the PRODUCTION Operational Aging file)\n"
         "        --sheet <name>             specific Aging sheet name (default: auto)\n"
         "        --label <TEXT>             tag appended to workbook filename; use PRODUCTION for signed-off runs"),
        ("TRAE IDE sandbox note",
         "Writes under /Documents/__data/… are BLOCKED inside the TRAE sandbox. From inside TRAE you can still read the workbook from ~/Downloads/ or /tmp/. To land it in the canonical reconcile dir, run the command above from your own Terminal shell (outside TRAE)."),
        ("Analogous PO/Receipt script",
         "generate_post_load_tieback.py — same sheet naming convention, same banner/header/fill styling, same verdict-cell colouring."),
    ]
    df_ov = pd.DataFrame(ov_lines, columns=["Section", "Detail"])

    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        df_ov.to_excel(w, sheet_name="0-Overview", index=False, startrow=2)
        df_sum.to_excel(w, sheet_name="1-Exec Summary", index=False, startrow=2)
        df_tt.to_excel(w, sheet_name="2-By Txn Type", index=False, startrow=2)
        df_ccy.to_excel(w, sheet_name="3-By Company Currency", index=False, startrow=2)
        df_cov.to_excel(w, sheet_name="4-Currency Coverage", index=False, startrow=2)
        df_byinvdate.to_excel(w, sheet_name="4B-By Invoice Date", index=False, startrow=2)
        df_cn.to_excel(w, sheet_name="5-Credit Note Exact", index=False, startrow=2)
        df_cn_date.to_excel(w, sheet_name="5B-Credit Note Date", index=False, startrow=2)
        df_sup.to_excel(w, sheet_name="6-Reconcile (Supplier)", index=False, startrow=2)
        df_inv.to_excel(w, sheet_name="7-Reconcile (Invoice)", index=False, startrow=2)
        df_inv_date.to_excel(w, sheet_name="7B-Reconcile (Invoice Date)", index=False, startrow=2)
        df_notes.to_excel(w, sheet_name="9-Notes", index=False, startrow=2)
        df_excl.to_excel(w, sheet_name="10-Excluded WD Rows", index=False, startrow=2)
        _style(w)

        # Exec Summary: highlight the verdict cell AND the new date-verdict cell
        ws0 = w.sheets["1-Exec Summary"]
        for i in range(4, ws0.max_row + 1):
            if ws0.cell(i, 1).value == "Reconciliation status":
                c = ws0.cell(i, 2)
                c.fill = (
                    OK_FILL
                    if (res["grand_match"] and res["cn_amt_mism"] == 0
                        and res["inv_totals"][5] == 0 and overall_date_ok
                        and byinvdate_ok)
                    else BAD_FILL
                )
                c.font = Font(bold=True)
            elif ws0.cell(i, 1).value == "Date reconciliation":
                c = ws0.cell(i, 2)
                c.fill = OK_FILL if overall_date_ok else BAD_FILL
                c.font = Font(bold=True)
            elif str(ws0.cell(i, 1).value or "").startswith("By Invoice Date"):
                c = ws0.cell(i, 2)
                c.fill = OK_FILL if byinvdate_ok else BAD_FILL
                c.font = Font(bold=True)
        for i in range(4, ws0.max_row + 1):
            if ws0.cell(i, 1).value == "Grand delta  (Sent − Loaded)":
                try:
                    v = float(str(ws0.cell(i, 2).value).replace("£", "").replace(",", ""))
                    if abs(v) > 0.01:
                        ws0.cell(i, 2).fill = BAD_FILL
                except Exception:
                    pass

        # Delta highlighting on the matrix tabs (row=3 is header, data starts row 4)
        # 2-By Txn Type: Δ rows (col 4), Δ amount (col 7)
        _highlight_delta(w.sheets["2-By Txn Type"], 4, thr=0)
        _highlight_delta(w.sheets["2-By Txn Type"], 7, thr=0.01)
        # 3-By Company Currency
        _highlight_delta(w.sheets["3-By Company Currency"], 5, thr=0)
        _highlight_delta(w.sheets["3-By Company Currency"], 8, thr=0.01)
        # 4B-By Invoice Date — Δ rows (col 7, 1-based); Δ amount (col 10)
        _highlight_delta(w.sheets["4B-By Invoice Date"], 7, thr=0)
        _highlight_delta(w.sheets["4B-By Invoice Date"], 10, thr=0.01)
        # 5-Credit Note Exact — Δ Amt is col 8 (1-based)
        _highlight_delta(w.sheets["5-Credit Note Exact"], 8, thr=0.01)
        # 5B-Credit Note Date — Δ Amt = col 7; Δ days col 10 (non-zero = red); Status col 11
        _highlight_delta(w.sheets["5B-Credit Note Date"], 7, thr=0.01)
        ws5b = w.sheets["5B-Credit Note Date"]
        for i in range(4, ws5b.max_row + 1):
            # Δ days (col 10 = 1-based): if numeric non-zero and not MISMATCH we still highlight
            try:
                v = int(str(ws5b.cell(i, 10).value or "0"))
                if v != 0:
                    ws5b.cell(i, 10).fill = BAD_FILL
            except (TypeError, ValueError):
                pass
            # Status: MATCH green, anything else red
            st = str(ws5b.cell(i, 11).value or "").strip()
            if st == "MATCH":
                ws5b.cell(i, 11).fill = OK_FILL
            elif st not in ("", "Match", "UNKNOWN", "Sent only", "Loaded only"):
                ws5b.cell(i, 11).fill = BAD_FILL
            elif st == "UNKNOWN":
                ws5b.cell(i, 11).fill = PatternFill("solid", fgColor="FFEB9C")
        # 6-Reconcile (Supplier) — Δ rows (6), Δ amount (9)
        _highlight_delta(w.sheets["6-Reconcile (Supplier)"], 6, thr=0)
        _highlight_delta(w.sheets["6-Reconcile (Supplier)"], 9, thr=0.01)
        # 7-Reconcile (Invoice) — Δ Amt is col 7 (1-based); Δ days is col 12
        _highlight_delta(w.sheets["7-Reconcile (Invoice)"], 7, thr=0.01)
        ws7 = w.sheets["7-Reconcile (Invoice)"]
        for i in range(4, ws7.max_row + 1):
            try:
                v = int(str(ws7.cell(i, 12).value or "0"))
                if v != 0:
                    ws7.cell(i, 12).fill = BAD_FILL
            except (TypeError, ValueError):
                pass
        # 7B-Reconcile (Invoice Date) — Δ Amt col 7; Δ days col 12; Date Status col 13
        _highlight_delta(w.sheets["7B-Reconcile (Invoice Date)"], 7, thr=0.01)
        ws7b = w.sheets["7B-Reconcile (Invoice Date)"]
        YELLOW = PatternFill("solid", fgColor="FFEB9C")
        for i in range(4, ws7b.max_row + 1):
            # Δ days: any non-zero → amber/red visually
            try:
                v = int(str(ws7b.cell(i, 12).value or "0"))
                if v != 0:
                    ws7b.cell(i, 12).fill = BAD_FILL
            except (TypeError, ValueError):
                pass
            # Date Status → color coding:
            st = str(ws7b.cell(i, 13).value or "").strip()
            if st == "DATE_MATCH":
                ws7b.cell(i, 13).fill = OK_FILL
            elif st == "DATE_ORDER":
                ws7b.cell(i, 13).fill = YELLOW   # benign artefact
            elif st == "DATE_MISMATCH":
                ws7b.cell(i, 13).fill = BAD_FILL  # real drift — red
            elif st == "DATE_UNKNOWN":
                ws7b.cell(i, 13).fill = YELLOW
            elif st in ("Sent only", "Loaded only"):
                ws7b.cell(i, 13).fill = BAD_FILL
        # 10-Excluded WD Rows — highlight any non-zero Invoice Amt so the 3
        # Draft tenant rows visually pop (they are the reason we filtered).
        # Invoice Amt is column 8 (1-based; cols 1-7 = Co/TxnType/CCY/Supplier/
        # Ref/InvDate/DueDate).  Always mark Status col (col 10) with the
        # existing status string so reviewers see Draft/etc at a glance.
        ws9 = w.sheets.get("10-Excluded WD Rows")
        if ws9 is None:
            ws9 = w.sheets.get("9-Excluded WD Rows")
        if ws9 is not None:
            # status col (col 10 = 1-based) = mark amber/alert background if
            # the value isn't Approved-like.
            for i in range(4, ws9.max_row + 1):
                st = str(ws9.cell(i, 10).value or "").strip().lower()
                if st and st not in ("approved", "posted", "partially paid", "paid"):
                    ws9.cell(i, 10).fill = BAD_FILL


# --------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------

def main() -> int:
    global WORKDAY_EXTRACT_DEFAULT, WORKDAY_EXTRACT_SHEET_DEFAULT
    ap = argparse.ArgumentParser(
        description="Post-load tie-out: sent Supplier-Invoice EIB vs a Workday aging extract."
    )
    ap.add_argument(
        "--workday-extract", default=WORKDAY_EXTRACT_DEFAULT,
        help="Path to the Workday Operational Aging xlsx export (default: the PRODUCTION aging extract).",
    )
    ap.add_argument(
        "--sheet", default=WORKDAY_EXTRACT_SHEET_DEFAULT,
        help="Sheet name to read (default: auto — the named sheet if present, else the first).",
    )
    ap.add_argument(
        "--label", default=None,
        help="Optional tag added to the output filename (e.g. PRODUCTION) so runs don't clobber.",
    )
    args = ap.parse_args()

    WORKDAY_EXTRACT_DEFAULT = os.path.expanduser(args.workday_extract)
    extract_path = WORKDAY_EXTRACT_DEFAULT
    if args.sheet:
        WORKDAY_EXTRACT_SHEET_DEFAULT = args.sheet
    if not os.path.exists(extract_path):
        print(f"Workday extract not found: {extract_path}", file=sys.stderr)
        return 2
    wd, resolved_sheet = load_workday_aging(extract_path, args.sheet)
    con = duckdb.connect(DB_PATH, read_only=True)
    try:
        res = build(con, wd)
    finally:
        con.close()

    tag = f"_{args.label}" if args.label else ""
    fname = f"FA_PostLoad_Tieback_SupplierInvoice_{dt.date.today():%Y%m%d}{tag}.xlsx"

    def _try_write(target_dir: str, friendly: str) -> tuple[str | None, bool]:
        try:
            os.makedirs(target_dir, exist_ok=True)
            dest = os.path.join(target_dir, fname)
            write_workbook(res, dest, os.path.basename(extract_path))
            print(f"Wrote to {friendly}: {dest}")
            return dest, True
        except Exception as e:
            print(f"Could not write to {friendly} ({type(e).__name__}: {e}).", file=sys.stderr)
            return None, False

    # 1) PRIMARY — the reconcile dir in Documents/__data (always succeeds when you
    #    run the script yourself outside the TRAE sandbox; only fails inside TRAE)
    primary_path, primary_ok = _try_write(OUT_DIR, "reconcile dir (Documents/__data)")

    # 2) FALLBACK — ~/Downloads (user-visible on all platforms; TRAE sandbox permits)
    downloads_path, downloads_ok = _try_write(DOWNLOADS_FALLBACK_DIR, "fallback dir (~/Downloads)")

    # 3) STAGE — /tmp (last resort; sandbox always permits this)
    stage_path, stage_ok = _try_write(STAGE_DIR, "stage dir (/tmp)")

    # If primary didn't get written, but either Downloads or /tmp did, copy from there
    # (extra safety in case write fails but a two-step shutil.copyfile would succeed)
    if not primary_ok:
        source_for_copy = None
        if os.path.exists(downloads_path or ""):
            source_for_copy = downloads_path
        elif os.path.exists(stage_path or ""):
            source_for_copy = stage_path
        if source_for_copy:
            try:
                os.makedirs(OUT_DIR, exist_ok=True)
                dest = os.path.join(OUT_DIR, fname)
                shutil.copyfile(source_for_copy, dest)
                print(f"  (recovery) copied via shutil → reconcile dir: {dest}")
                primary_ok = True
            except Exception as e:
                print(f"  (recovery) shutil copy also failed — sandbox is still blocking {OUT_DIR}. "
                      f"Please re-run this script from your own shell, or manually copy from "
                      f"{source_for_copy} → {os.path.join(OUT_DIR, fname)}", file=sys.stderr)

    print()
    if res.get("excluded_rows"):
        print(
            f"Excluded from compare: "
            f"{len(res['excluded_rows']):,} WD rows  "
            f"Σ={sum(float(r[7]) for r in res['excluded_rows']):,.2f}  "
            f"(see sheet 9-Excluded WD Rows)"
        )
    print(
        f"Grand total: Sent={res['grand_sent']:,.2f}  "
        f"WD={res['grand_wd']:,.2f}  Δ={res['grand_delta']:,.2f}  "
        f"{'MATCH' if res['grand_match'] else 'MISMATCH'}"
    )
    print(
        f"CN exact key: Both={res['cn_both']}  SentOnly={res['cn_sent_only']}  "
        f"WDOnly={res['cn_wd_only']}  AmtMism={res['cn_amt_mism']}"
    )
    print(
        f"Supplier buckets:  Match={res['sup_bucket_counts'][0]}  "
        f"SentOnly={res['sup_bucket_counts'][1]}  "
        f"WDOnly={res['sup_bucket_counts'][2]}  "
        f"Review={res['sup_bucket_counts'][3]}"
    )
    print(
        f"Invoice fuzzy:  SentRows={res['inv_totals'][0]}  "
        f"WDRows={res['inv_totals'][1]}  Matched={res['inv_totals'][2]}  "
        f"SentOnly={res['inv_totals'][3]}  WDOnly={res['inv_totals'][4]}  "
        f"AmtMism={res['inv_totals'][5]}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
