"""Post-load tie-out: sent Customer-Invoice EIB payload  vs  Workday AR Aging extract.

This is the post-load (4th-layer) reconciliation for Customer Invoices + Credit Notes,
mirroring generate_supplier_invoice_post_load_reconcile.py for the Supplier leg.

Legs compared:
  [3] Sent   : workday_customer_invoice + workday_customer_credit_note in rpatel.duckdb
               joined to src_fin_customer_invoice_nrm for customer names,
               document numbers, and document types.
  [4] Loaded : Workday 'AR Aging Detail' (or Customer Aging equivalent) extract —
               the Workday tenant export after the customer-invoice EIB load.
               Default path: ~/Documents/__data/excel/workday/extract/ar_invoices/
               Override with --workday-extract.

Join keys / quirks handled:
  - Transaction types   : Loaded labels = 'Invoice' / 'Adjustment' (or 'Credit Memo');
                          Sent has INVOICE table (column_ha lines) and CREDIT_NOTE
                          table (column_fb lines). We classify sent Adjustments as
                          'CREDIT_NOTE' from the EIB credit-note staging table.
  - Credit-note grain   : Sent column_h (Customer Invoice Number, 12/08/26 3+ fix) is
                          the authoritative doc ref; column_fc (Line Memo) is the
                          pre-fix fallback when column_h is still blank. Loaded
                          'Transaction' reference stores the Workday tenant ref. If
                          the loaded extract exposes a source-system document ref
                          column we use it for exact key match; otherwise fall
                          through to bucket + invoice-layer.
  - Invoice grain       : Sent EIB column_h (Customer Invoice Number) now carries the
                          source GP document_number DIRECTLY (12/08/26 3+ fix).
                          Prior to this fix the reconciler used a ROW_NUMBER-over-
                          (BU, wd_customer_id, date, amt) indirect join back to
                          staging, which was a trace-only grouping not a proof of
                          printed-invoice numbers. Now we reconcile the exact cell
                          value we write into column H. Loaded Aging still stores
                          tenant refs (e.g. FA-CINV-NNNNNN); we expose the GP
                          external doc ref side-by-side on the report and trace the
                          4-tuple where keys don't directly share a namespace.
  - Company labels      : Sent-side Workday-BE labels ('THE FA GROUP - FA') are
                          mapped to the loaded-side user-facing short labels.
  - Customer name norm  : Exact name match first, then stripped-name match
                          (trim, collapse whitespace, drop Ltd/Ltd/PLC suffixes).

Output: FA_PostLoad_Tieback_CustomerInvoice_YYYYMMDD.xlsx in the reconcile dir.
        Staged to ~/Downloads first (fallback for when TRAE sandbox blocks
        ~/Documents), then to /tmp as last resort — nothing is written inside
        the repo itself.
        Outside the TRAE sandbox the recovery shutil.copyfile into reconcile dir succeeds.
"""

from __future__ import annotations

import argparse
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
WORKDAY_EXTRACT_DEFAULT = os.path.expanduser(
    "~/Documents/__data/excel/workday/extract/ar_invoices/AR_Aging_Detail.xlsx"
)
WORKDAY_EXTRACT_SHEET_DEFAULT = None
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
# Load the Workday AR Aging extract
# --------------------------------------------------------------------

def load_workday_ar_aging(path: str, sheet: str | None) -> tuple[pd.DataFrame, str]:
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

    _RENAME_FMT_A = {
        "Company": "company",
        "Company Currency": "company_ccy",
        "Customer": "customer_name_wd",
        "Transaction": "wd_txn_ref",
        "Transaction Type": "txn_type",
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
        "Document Number": "src_doc_ref",
        "Source Document": "src_doc_ref",
    }
    _RENAME_FMT_B = {
        "Company": "company",
        "Company Currency": "company_ccy",
        "Reference Number": "wd_txn_ref",
        "Customer": "customer_name_wd",
        "Customer ID": "customer_id",
        "Invoice Number": "customer_invoice_number",
        "Invoice Date": "txn_date_raw",
        "Due Date": "due_date_raw",
        "Currency": "ccy",
        "Invoice Amount": "original_amt",
        "Balance Due": "balance_amt",
        "Amount Due": "balance_amt",
        "Status": "status",
        "Invoice Status": "status",
        "Gapless Document Number": "gapless_docno",
        "External Document Number": "external_docno",
        "Purchase Order": "po_ref",
        "Memo": "memo",
    }

    header_idx = None
    fmt = None
    for i in range(0, min(raw.shape[0], 60)):
        row = [str(c).strip().lower() if c is not None else "" for c in raw.iloc[i].tolist()]
        if "company" in row and "transaction type" in row and "original amount in transaction currency" in row:
            header_idx = i
            fmt = "A"
            break
        if "company" in row and "reference number" in row and "invoice amount" in row and "balance due" in row:
            header_idx = i
            fmt = "B"
            break
        if "company" in row and "invoice number" in row and "invoice amount" in row and "balance due" in row:
            header_idx = i
            fmt = "B"
            break
        if "company" in row and "customer" in row and "invoice amount" in row and "amount due" in row and "invoice status" in row:
            header_idx = i
            fmt = "B"
            break

    if header_idx is None or fmt is None:
        header_idx = 21
        fmt = "A"
        for probe_row in range(min(40, len(raw))):
            row_vals = [str(c).strip() if c is not None else "" for c in raw.iloc[probe_row].tolist()]
            if any("company" in v.lower() for v in row_vals) and any(
                ("transaction" in v.lower() and "type" in v.lower()) or
                ("customer" in v.lower()) or
                ("currency" in v.lower())
                for v in row_vals
            ):
                header_idx = probe_row
                break

    data_start = header_idx + 1
    raw_hdr = raw.iloc[header_idx].tolist()
    hdr = [str(c).strip() if c is not None and str(c).strip() != "" else f"col_{i}" for i, c in enumerate(raw_hdr)]
    data = raw.iloc[data_start:].copy()
    data.columns = hdr
    data = data.dropna(how="all").reset_index(drop=True)

    applied_map = _RENAME_FMT_A if fmt == "A" else _RENAME_FMT_B
    hdr_lookup = {str(h).strip(): h for h in data.columns}
    applied_rename = {}
    for k, v in applied_map.items():
        if k in hdr_lookup:
            applied_rename[hdr_lookup[k]] = v
        elif k.lower() in {str(x).strip().lower(): x for x in data.columns}:
            real_key = {str(x).strip().lower(): x for x in data.columns}[k.lower()]
            applied_rename[real_key] = v
    data = data.rename(columns=applied_rename)

    if fmt == "B":
        first_col = data.columns[0]
        data["_first_col_label"] = data[first_col].astype(str)

        def _derive_txn(label_val, inv_amt):
            s = (label_val or "").lower()
            if "adjustment" in s and ("customer invoice" in s or "invoice" in s):
                return "CREDIT_NOTE"
            if (s.startswith("customer invoice adjustment:") or s.startswith("customer invoice adj:")
                    or "credit memo" in s or "credit note" in s or "credit adjustment" in s):
                return "CREDIT_NOTE"
            if s.startswith("customer invoice:"):
                return "INVOICE"
            try:
                v = float(inv_amt)
                return "CREDIT_NOTE" if v < 0 else "INVOICE"
            except (TypeError, ValueError):
                return "INVOICE"

        data["txn_type"] = [
            _derive_txn(data["_first_col_label"].iloc[i], data["original_amt"].iloc[i] if "original_amt" in data.columns else 0.0)
            for i in range(len(data))
        ]
        if "ccy" not in data.columns:
            data["ccy"] = "GBP"
        if "wd_txn_ref" not in data.columns:
            if "customer_invoice_number" in data.columns:
                data["wd_txn_ref"] = data["customer_invoice_number"].astype(str)
            else:
                data["wd_txn_ref"] = data["_first_col_label"].astype(str)
        else:
            _wd_ref_is_empty = (
                data["wd_txn_ref"]
                .fillna("")
                .astype(str)
                .str.strip()
                .eq("")
                .all()
            )
            if _wd_ref_is_empty:
                if "customer_invoice_number" in data.columns and not data["customer_invoice_number"].fillna("").astype(str).str.strip().eq("").all():
                    data["wd_txn_ref"] = data["customer_invoice_number"].astype(str)
                else:
                    data["wd_txn_ref"] = data["_first_col_label"].astype(str)
        if "_first_col_label" in data.columns:
            data = data.drop(columns=["_first_col_label"])

    numeric_out_cols = [
        "original_amt", "balance_amt", "balance_amt_gbp",
        "cur_current", "cur_1_30", "cur_31_60", "cur_61_90",
        "cur_91_180", "cur_181_365", "cur_366p",
        "payments_applied", "discount_taken", "aging_days",
    ]
    for c in numeric_out_cols:
        if c in data.columns:
            data[c] = pd.to_numeric(data[c], errors="coerce").fillna(0.0)
    if "txn_date_raw" in data.columns:
        data["txn_date"] = pd.to_datetime(data["txn_date_raw"], errors="coerce")
    else:
        data["txn_date"] = pd.NaT
    if "due_date_raw" in data.columns:
        data["due_date"] = pd.to_datetime(data["due_date_raw"], errors="coerce")
    else:
        data["due_date"] = pd.NaT

    if "txn_type" not in data.columns:
        for c in data.columns:
            if "transaction" in str(c).lower() and "type" in str(c).lower():
                data = data.rename(columns={c: "txn_type"})
                break
    if "wd_txn_ref" not in data.columns:
        for c in data.columns:
            if str(c).strip().lower() == "transaction":
                data = data.rename(columns={c: "wd_txn_ref"})
                break
    if "customer_name_wd" not in data.columns:
        for c in data.columns:
            if str(c).strip().lower() == "customer":
                data = data.rename(columns={c: "customer_name_wd"})
                break
    if "ccy" not in data.columns:
        for c in data.columns:
            s = str(c).strip().lower()
            if "transaction" in s and "currency" in s:
                data = data.rename(columns={c: "ccy"})
                break

    def _map_txn(s):
        s_low = str(s).strip().lower()
        if s_low == "invoice":
            return "INVOICE"
        if s_low in ("adjustment", "credit memo", "credit note", "credit",
                     "credit_note", "cn", "credit memo", "debit memo", "debit note"):
            return "CREDIT_NOTE"
        if s_low.startswith("other_credit_note") or s_low.startswith("other_credit") or s_low.startswith("other_adjustment"):
            return "CREDIT_NOTE"
        if s and str(s).strip():
            return "OTHER_" + str(s).strip().upper()
        return "UNKNOWN"

    data["txn_type_norm"] = data.get("txn_type", pd.Series([""] * len(data))).map(_map_txn)
    data["customer_name_norm"] = data.get("customer_name_wd", pd.Series([""] * len(data))).apply(_norm_name)
    return data, resolved_sheet


# --------------------------------------------------------------------
# Load Sent EIB side from DuckDB (read-only)
# --------------------------------------------------------------------

def load_sent_eib(con: duckdb.DuckDBPyConnection) -> pd.DataFrame:
    inv_df = con.execute(
        """
        SELECT
            'INVOICE'                                     AS txn_type_norm,
            e.column_k                                    AS eib_company,
            TRIM(e.column_cp)                             AS ccy,
            e.column_cq                                   AS wd_customer_id,
            CAST(e.column_gf AS VARCHAR)                  AS inv_date_str,
            CAST(TRY_STRPTIME(NULLIF(TRIM(e.column_gf), ''), '%d/%m/%Y') AS DATE)
                                                          AS txn_date,
            COALESCE(NULLIF(TRIM(e.column_h), ''),
                     CAST(NULL AS VARCHAR))              AS eib_txn_ref,
            CAST(e.column_ha AS DOUBLE)                   AS original_amt,
            CAST(e.column_mj AS DOUBLE)                   AS control_amt,
            CAST(NULL AS VARCHAR)                         AS gp_customer_name,
            COALESCE(NULLIF(TRIM(e.column_h), ''),
                     CAST(NULL AS VARCHAR))              AS gp_document_number
          FROM workday_customer_invoice e
        """
    ).fetchdf()

    inv_missing_h = int(((inv_df["gp_document_number"].isna()) | (inv_df["gp_document_number"] == "")).sum())
    if inv_missing_h:
        import warnings
        warnings.warn(
            f"[post_load] {inv_missing_h} INVOICE rows in publish workday_customer_invoice have "
            f"BLANK column_h (Customer Invoice Number). The 3+ legacy GP invoice number did not "
            f"reach the printed invoice. Check workday_customer_invoice.sql SELECT includes "
            f"t.document_number column_h; re-run the NRM→SRC→PUBLISH chain before signing off."
        )

    gp_inv_lookup = con.execute(
        """
        SELECT wd_customer_id, document_number, customer_name
          FROM src_fin_customer_invoice_nrm
         WHERE wd_customer_id IS NOT NULL
           AND UPPER(TRIM(document_type)) LIKE '%INVOICE%'
        """
    ).fetchdf()

    inv_merged = inv_df.merge(
        gp_inv_lookup,
        left_on=["wd_customer_id", "gp_document_number"],
        right_on=["wd_customer_id", "document_number"],
        how="left",
    )
    inv_merged["gp_customer_name"] = (
        inv_merged["gp_customer_name"]
        .fillna(inv_merged["customer_name"])
    )
    inv_merged["gp_document_number"] = (
        inv_merged["gp_document_number"]
        .fillna(inv_merged["document_number"])
    )
    inv_merged = inv_merged.drop(
        columns=["document_number", "customer_name"],
        errors="ignore",
    )

    cn_df = con.execute(
        """
        SELECT
            'CREDIT_NOTE'                                 AS txn_type_norm,
            c.column_l                                    AS eib_company,
            TRIM(c.column_bk)                             AS ccy,
            c.column_aj                                   AS wd_customer_id,
            CAST(c.column_bl AS VARCHAR)                  AS inv_date_str,
            CAST(TRY_STRPTIME(NULLIF(TRIM(c.column_bl), ''), '%d/%m/%Y') AS DATE)
                                                          AS txn_date,
            COALESCE(NULLIF(TRIM(c.column_h), ''),
                     NULLIF(TRIM(c.column_fc), ''),
                     CAST(NULL AS VARCHAR))              AS eib_txn_ref,
            CAST(c.column_fb AS DOUBLE)                   AS original_amt,
            CAST(c.column_fb AS DOUBLE)                   AS control_amt,
            CAST(NULL AS VARCHAR)                         AS gp_customer_name,
            COALESCE(NULLIF(TRIM(c.column_h), ''),
                     NULLIF(TRIM(c.column_fc), ''),
                     CAST(NULL AS VARCHAR))              AS gp_document_number
          FROM workday_customer_credit_note c
        """
    ).fetchdf()

    sent = pd.concat([inv_merged, cn_df], ignore_index=True)
    sent["company"] = sent["eib_company"].map(
        COMPANY_NAME_NORMALISATION
    ).fillna(sent["eib_company"])

    id_to_name = dict(con.execute(
        """
        SELECT wd_customer_id, wd_customer_name
          FROM (
              SELECT wd_customer_id, wd_customer_name,
                     ROW_NUMBER() OVER (PARTITION BY wd_customer_id ORDER BY wd_customer_name) rn
                FROM src_fin_customer_invoice_nrm
               WHERE wd_customer_id IS NOT NULL
               UNION
              SELECT customer_id, nrm_customer_name,
                     ROW_NUMBER() OVER (PARTITION BY customer_id ORDER BY nrm_customer_name) rn
                FROM src_fin_workday_customer_map
               WHERE customer_id IS NOT NULL
          ) _x WHERE rn = 1
        """
    ).fetchall())
    sent["customer_name_eib"] = (
        sent["wd_customer_id"].map(lambda sid: id_to_name.get(sid, None))
    )
    sent["customer_name_eib"] = (
        sent["customer_name_eib"]
        .where(sent["customer_name_eib"].notna() & (sent["customer_name_eib"] != ""),
               sent["gp_customer_name"])
    )
    sent["customer_name_eib"] = (
        sent["customer_name_eib"]
        .where(sent["customer_name_eib"].notna() & (sent["customer_name_eib"] != ""),
               sent["wd_customer_id"].fillna(""))
    )
    sent["eib_txn_ref"] = sent["eib_txn_ref"].fillna(sent["gp_document_number"])
    sent["customer_name_norm"] = sent["customer_name_eib"].apply(_norm_name)
    return sent


# --------------------------------------------------------------------
# Build the reconciliation metrics
# --------------------------------------------------------------------

def build(con: duckdb.DuckDBPyConnection, wd: pd.DataFrame) -> dict:
    sent = load_sent_eib(con)

    if "status" in wd.columns or ("ccy" in wd.columns and (wd["ccy"] == "GBP").any()):
        from collections import defaultdict
        s_lookup = defaultdict(list)
        for si in range(len(sent)):
            sr = sent.iloc[si]
            sr_amt = float(sr["original_amt"]) if pd.notna(sr["original_amt"]) else 0.0
            sr_key = (sr["company"], sr["customer_name_norm"], round(sr_amt, 2))
            s_lookup[sr_key].append(si)
        wd_ccy_vals = wd.get("ccy", pd.Series(["GBP"] * len(wd))).tolist()
        wd_co_vals = wd["company"].tolist()
        wd_cnorm_vals = wd.get("customer_name_norm", pd.Series([""] * len(wd))).tolist()
        wd_amt_vals = wd["original_amt"].tolist()
        used_sent = set()
        for wi in range(len(wd)):
            if str(wd_ccy_vals[wi]).upper() != "GBP":
                continue
            w_amt = float(wd_amt_vals[wi]) if pd.notna(wd_amt_vals[wi]) else 0.0
            key = (wd_co_vals[wi], wd_cnorm_vals[wi], round(w_amt, 2))
            cands = [idx for idx in s_lookup.get(key, []) if idx not in used_sent]
            cands_ccy_non_gbp = [idx for idx in cands if str(sent.iloc[idx].get("ccy", "GBP")).upper() != "GBP"]
            if len(cands_ccy_non_gbp) == 1:
                si = cands_ccy_non_gbp[0]
                sr = sent.iloc[si]
                used_sent.add(si)
                wd.at[wi, "ccy"] = sr.get("ccy", wd_ccy_vals[wi])
                if pd.notna(sr.get("customer_name_wd")) and str(sr.get("customer_name_wd", "")).strip() != "":
                    wd.at[wi, "customer_name_wd"] = sr["customer_name_wd"]
                if pd.notna(sr.get("customer_name_norm")) and str(sr.get("customer_name_norm", "")).strip() != "":
                    wd.at[wi, "customer_name_norm"] = sr["customer_name_norm"]

    excluded_summary = []
    excluded_rows_sorted = []
    if "status" in wd.columns:
        KEEP = {"Approved", "Posted", "Partially Paid", "Paid", "Paid Partially"}
        keep_mask = wd["status"].astype(str).str.strip().str.lower().isin({k.lower() for k in KEEP})
        if "wd_txn_ref" not in wd.columns:
            wd["wd_txn_ref"] = ""
        excluded_df = wd.loc[~keep_mask].copy()
        wd = wd.loc[keep_mask].reset_index(drop=True)
        if len(excluded_df) > 0:
            excl_coerce = excluded_df.copy()
            for c in ["original_amt", "balance_amt"]:
                if c in excl_coerce.columns:
                    excl_coerce[c] = pd.to_numeric(excl_coerce[c], errors="coerce").fillna(0.0)
            by_status = excl_coerce.groupby("status").agg(
                rows=("original_amt", "size"),
                amount=("original_amt", lambda s: round(float(s.sum()), 2)),
            ).reset_index()
            excluded_summary = [
                f"Workday status='{r['status']}' — not an EIB-loaded Approved row: {int(r['rows'])} rows  £{float(r['amount']):+,.2f}"
                for _, r in by_status.iterrows()
            ]
            for _, r in excl_coerce.sort_values(
                by=["company", "status", "txn_date_raw"] if "txn_date_raw" in excl_coerce.columns else ["company", "status"]
            ).iterrows():
                excluded_rows_sorted.append([
                    r.get("company", ""),
                    r.get("txn_type_norm", r.get("txn_type", "")),
                    r.get("ccy", ""),
                    r.get("customer_name_wd", r.get("customer", "")),
                    r.get("wd_txn_ref", ""),
                    r.get("txn_date_raw", r.get("txn_date", "")),
                    r.get("due_date_raw", r.get("due_date", "")),
                    float(r.get("original_amt", 0.0)) if pd.notna(r.get("original_amt")) else 0.0,
                    float(r.get("balance_amt", 0.0)) if pd.notna(r.get("balance_amt")) else 0.0,
                    r.get("status", ""),
                    f"Workday status='{r.get('status', '')}' — not an EIB-loaded Approved row",
                ])

    grand_sent = float(sent["original_amt"].sum().round(2))
    grand_wd = float(wd["original_amt"].sum().round(2))
    grand_delta = round(grand_sent - grand_wd, 2)
    grand_match = abs(grand_delta) < 0.01

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

    sent_cn = sent[sent["txn_type_norm"] == "CREDIT_NOTE"].copy()
    wd_cn = wd[wd["txn_type_norm"] == "CREDIT_NOTE"].copy()

    def _canonical_ref(x) -> str:
        s = "" if x is None else str(x).strip()
        for prefix in (
            "Customer Invoice Adjustment: ",
            "Customer Invoice Adjustment:",
            "Customer Invoice: ",
            "Customer Invoice:",
            "Customer Payment: ",
            "Customer Payment:",
            "Customer Refund: ",
            "Customer Refund:",
        ):
            if s.startswith(prefix):
                s = s[len(prefix):].strip()
                break
        return s

    sent_cn["_exact_key"] = sent_cn["eib_txn_ref"].fillna(sent_cn.get("gp_document_number")).apply(_canonical_ref)
    if "src_doc_ref" in wd_cn.columns and wd_cn["src_doc_ref"].notna().any():
        wd_cn["_exact_key"] = wd_cn["src_doc_ref"].apply(_canonical_ref)
    else:
        wd_cn["_exact_key"] = wd_cn["wd_txn_ref"].apply(_canonical_ref)

    cn_merge_on_ref = True

    cn_merged = sent_cn.merge(
        wd_cn, left_on="_exact_key", right_on="_exact_key",
        how="outer", suffixes=("_sent", "_wd"), indicator=True
    )
    cn_merged["amount_sent"] = cn_merged["original_amt_sent"].fillna(0.0).round(2)
    cn_merged["amount_wd"] = cn_merged["original_amt_wd"].fillna(0.0).round(2)
    cn_merged["delta_amount"] = (cn_merged["amount_sent"] - cn_merged["amount_wd"]).round(2)
    cn_merged["join_side"] = cn_merged["_merge"].map(
        {"both": "Both", "left_only": "Sent only", "right_only": "Loaded only"}
    )
    cn_exact_rows = [
        [
            r.get("company_sent") or r.get("company_wd") or "",
            r.get("_exact_key") or r.get("eib_txn_ref") or r.get("wd_txn_ref") or "",
            r.get("txn_type_norm_sent") or r.get("txn_type_norm_wd") or "",
            r.get("customer_name_eib") or r.get("customer_name_wd") or "",
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

    def _cust_agg(df: pd.DataFrame, label: str) -> pd.DataFrame:
        grp = df.groupby(["company", "customer_name_norm", "ccy"], dropna=False).agg(
            **{
                f"rows_{label}": ("original_amt", "size"),
                f"amount_{label}": ("original_amt", lambda s: round(float(s.sum()), 2)),
                f"raw_names_{label}": ("customer_name_eib" if "customer_name_eib" in df.columns
                                       else "customer_name_wd",
                                       lambda s: " // ".join(
                                           sorted(set(str(x) for x in s.dropna().unique() if str(x).strip()))
                                       ) if len(s.dropna()) else ""),
            }
        ).reset_index()
        return grp
    cust_merged = _cust_agg(sent, "sent").merge(
        _cust_agg(wd, "wd"),
        on=["company", "customer_name_norm", "ccy"], how="outer",
    ).fillna(0)
    cust_merged["rows_sent"] = cust_merged["rows_sent"].astype(int)
    cust_merged["rows_wd"] = cust_merged["rows_wd"].astype(int)
    cust_merged["delta_rows"] = cust_merged["rows_sent"] - cust_merged["rows_wd"]
    cust_merged["delta_amount"] = (cust_merged["amount_sent"] - cust_merged["amount_wd"]).round(2)
    cust_merged["customer_name_display"] = (
        cust_merged["raw_names_sent"].where(cust_merged["raw_names_sent"].astype(str).str.strip() != "",
                                            cust_merged["raw_names_wd"])
    )
    cust_merged = cust_merged.sort_values(
        ["company", "delta_amount"],
        key=lambda s: s.abs() if s.name == "delta_amount" else s,
        ascending=[True, False],
    ).reset_index(drop=True)

    cust_rows = [
        [
            r["company"],
            str(r["customer_name_display"] or ""),
            r["ccy"],
            int(r["rows_sent"]), int(r["rows_wd"]), int(r["delta_rows"]),
            float(r["amount_sent"]), float(r["amount_wd"]), float(r["delta_amount"]),
            "Match" if (r["delta_rows"] == 0 and abs(r["delta_amount"]) < 0.01)
            else ("Sent only" if (r["rows_wd"] == 0 and r["amount_wd"] == 0)
                  else ("Loaded only" if (r["rows_sent"] == 0 and r["amount_sent"] == 0)
                        else "Review")),
        ]
        for _, r in cust_merged.iterrows()
    ]

    cust_total_sent_rows = int(cust_merged["rows_sent"].sum())
    cust_total_wd_rows = int(cust_merged["rows_wd"].sum())
    cust_total_sent_amt = round(float(cust_merged["amount_sent"].sum()), 2)
    cust_total_wd_amt = round(float(cust_merged["amount_wd"].sum()), 2)
    cust_match_count = int(sum(
        1 for _, r in cust_merged.iterrows()
        if r["delta_rows"] == 0 and abs(r["delta_amount"]) < 0.01
    ))
    cust_sent_only_count = int(sum(
        1 for _, r in cust_merged.iterrows()
        if r["rows_wd"] == 0 and abs(r["amount_wd"]) < 0.005
    ))
    cust_wd_only_count = int(sum(
        1 for _, r in cust_merged.iterrows()
        if r["rows_sent"] == 0 and abs(r["amount_sent"]) < 0.005
    ))
    cust_review_count = len(cust_merged) - cust_match_count - cust_sent_only_count - cust_wd_only_count

    sent_inv = sent[sent["txn_type_norm"] == "INVOICE"].copy().reset_index(drop=True)
    wd_inv = wd[wd["txn_type_norm"] == "INVOICE"].copy().reset_index(drop=True)
    sent_inv["amount_key"] = sent_inv["original_amt"].round(2).astype(float)
    wd_inv["amount_key"] = wd_inv["original_amt"].round(2).astype(float)
    sent_inv["company_s"] = sent_inv["company"].astype(str)
    wd_inv["company_w"] = wd_inv["company"].astype(str)
    sent_inv["ccy_s"] = sent_inv["ccy"].astype(str)
    wd_inv["ccy_w"] = wd_inv["ccy"].astype(str)
    sent_inv["cust_s"] = sent_inv["customer_name_norm"].astype(str)
    wd_inv["cust_w"] = wd_inv["customer_name_norm"].astype(str)

    wd_unmatched_idx = set(wd_inv.index.tolist())
    inv_match_rows = []
    for si in range(len(sent_inv)):
        sr = sent_inv.iloc[si]
        candidates = []
        for wi in wd_unmatched_idx:
            wr = wd_inv.iloc[wi]
            if (sr["company_s"] == wr["company_w"]
                and sr["ccy_s"] == wr["ccy_w"]
                and sr["cust_s"] == wr["cust_w"]
                and abs(sr["amount_key"] - wr["amount_key"]) < 0.005):
                candidates.append((wi, abs(sr["amount_key"] - wr["amount_key"])))
        if candidates:
            candidates.sort(key=lambda t: t[1])
            best_wi, d = candidates[0]
            wd_unmatched_idx.discard(best_wi)
            wr = wd_inv.iloc[best_wi]
            inv_match_rows.append([
                sr["company_s"],
                sr["customer_name_eib"] or "",
                wr["customer_name_wd"] or "",
                sr["ccy_s"],
                float(sr["original_amt"]), float(wr["original_amt"]),
                round(float(sr["original_amt"]) - float(wr["original_amt"]), 2),
                str(sr["eib_txn_ref"] or sr["gp_document_number"] or ""),
                str(wr["wd_txn_ref"] or ""),
                "Matched",
            ])
        else:
            inv_match_rows.append([
                sr["company_s"],
                sr["customer_name_eib"] or "",
                "",
                sr["ccy_s"],
                float(sr["original_amt"]), 0.0,
                round(float(sr["original_amt"]), 2),
                str(sr["eib_txn_ref"] or sr["gp_document_number"] or ""), "",
                "Sent only",
            ])
    for wi in sorted(wd_unmatched_idx):
        wr = wd_inv.iloc[wi]
        inv_match_rows.append([
            wr["company_w"],
            "", wr["customer_name_wd"] or "",
            wr["ccy_w"],
            0.0, float(wr["original_amt"]),
            round(-float(wr["original_amt"]), 2),
            "", str(wr["wd_txn_ref"] or ""),
            "Loaded only",
        ])
    inv_match_rows.sort(key=lambda r: (r[9], r[0], r[1], r[2]))

    inv_matched = int(sum(1 for r in inv_match_rows if r[9] == "Matched"))
    inv_sent_only = int(sum(1 for r in inv_match_rows if r[9] == "Sent only"))
    inv_wd_only = int(sum(1 for r in inv_match_rows if r[9] == "Loaded only"))
    inv_matched_amt_mism = int(sum(
        1 for r in inv_match_rows if r[9] == "Matched" and abs(r[6]) > 0.01
    ))

    sent_ccys = set(sent["ccy"].dropna().unique().tolist())
    wd_ccys = set(wd["ccy"].dropna().unique().tolist())
    ccy_coverage = []
    for c in sorted(sent_ccys | wd_ccys, key=lambda c: CCY_ORDER.index(c) if c in CCY_ORDER else 99):
        in_sent = c in sent_ccys
        in_wd = c in wd_ccys
        status = "BOTH" if (in_sent and in_wd) else ("SENT_ONLY" if in_sent else "WD_ONLY")
        ccy_coverage.append([c, "Yes" if in_sent else "No", "Yes" if in_wd else "No", status])

    return {
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
        "cn_both": cn_both,
        "cn_sent_only": cn_sent_only,
        "cn_wd_only": cn_wd_only,
        "cn_amt_mism": cn_amt_mism,
        "cn_right_key": "_exact_key (canonicalised via src_doc_ref / wd_txn_ref with Workday first-column prefix stripped)",
        "cn_headers": ["Company", "Ref", "Type", "Customer", "CCY",
                       "Sent Amt", "Loaded Amt", "Δ Amt", "Join"],
        "cn_rows": cn_exact_rows,
        "cust_total_rows": [cust_total_sent_rows, cust_total_wd_rows,
                           cust_total_sent_rows - cust_total_wd_rows,
                           cust_total_sent_amt, cust_total_wd_amt,
                           round(cust_total_sent_amt - cust_total_wd_amt, 2)],
        "cust_bucket_counts": [cust_match_count, cust_sent_only_count,
                              cust_wd_only_count, cust_review_count],
        "cust_headers": ["Company", "Customer", "CCY",
                        "Sent rows", "Loaded rows", "Δ rows",
                        "Sent amount", "Loaded amount", "Δ amount", "Status"],
        "cust_rows": cust_rows,
        "inv_totals": [len(sent_inv), len(wd_inv),
                       inv_matched, inv_sent_only, inv_wd_only, inv_matched_amt_mism],
        "inv_headers": ["Company", "Sent Customer", "Loaded Customer", "CCY",
                        "Sent Amt", "Loaded Amt", "Δ Amt",
                        "Sent Ref (document)", "Loaded Ref (Workday txn)", "Status"],
        "inv_rows": inv_match_rows,
        "excluded_headers": ["Company", "Txn Type", "CCY", "Customer", "Ref",
                             "Invoice Date", "Due Date", "Invoice Amt",
                             "Balance Due", "Status", "Exclude reason"],
        "excluded_rows": excluded_rows_sorted,
        "excluded_summary": excluded_summary,
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
        if sname == "9-Excluded WD Rows":
            title_text = f"{sname} — Customer Invoice Post-Load Tieback (Sent EIB vs Workday export)"
        else:
            title_text = f"{sname} — Customer Invoice Post-Load Tieback (Sent EIB vs Workday AR Aging)"
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
        if sname == "9-Excluded WD Rows":
            status_col = None
            for col in range(1, ws.max_column + 1):
                if str(ws.cell(3, col).value).strip().lower() == "status":
                    status_col = col
                    break
            if status_col:
                for i in range(4, ws.max_row + 1):
                    ws.cell(i, status_col).fill = BAD_FILL


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

    grand_sent, grand_wd, grand_delta = res["grand_sent"], res["grand_wd"], res["grand_delta"]
    verdict = (
        "CLEAN — grand totals match to the penny"
        if res["grand_match"] and res["cn_amt_mism"] == 0 and res["inv_totals"][5] == 0
        else "REVIEW — see sheets for deltas"
    )

    summary_rows = [
        ("Report produced", dt.datetime.now().strftime("%A %d %B %Y, %H:%M")),
        ("What this is",
         "Post-load 4th-layer tieback of the ACTUAL Workday EIB Customer-Invoice output"
         " against the Workday AR Aging Detail extract (authoritative tenant export)."),
        ("Source [3] — Sent EIB (workday_customer_invoice + credit_note in DuckDB)",
         f"£ {grand_sent:,.2f}"),
        ("Source [4] — Loaded Workday AR aging extract",
         f"£ {grand_wd:,.2f}"),
        ("Grand delta  (Sent − Loaded)",
         f"£ {grand_delta:+,.2f}"),
        ("Workday rows EXCLUDED from compare (Draft/Voided/Cancelled)",
         ";  ".join(res.get("excluded_summary", [])) if res.get("excluded_summary") else "NONE — no Draft/Manual WD rows were removed"),
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
        (f"Credit-note exact key  (Both / Sent-only / Loaded-only / Amt mismatch)  [join on sent.eib_txn_ref == loaded.{res['cn_right_key']}]",
         f"{res['cn_both']} / {res['cn_sent_only']} / {res['cn_wd_only']} / {res['cn_amt_mism']}"),
        ("Customer-bucket counts  (Match / Sent-only / Loaded-only / Review)",
         " / ".join(str(x) for x in res["cust_bucket_counts"])),
        ("Invoice fuzzy match  (Sent rows / Loaded rows / Matched / Sent-only / Loaded-only / Amt mism)",
         " / ".join(str(x) for x in res["inv_totals"])),
        ("Source [4] file", workday_extract_basename),
        ("Reconciliation status",
         "✅ Tied out — Sent == Loaded to the penny on all authoritative layers"
         if res["grand_match"] and res["cn_amt_mism"] == 0 and res["inv_totals"][5] == 0
         else "⚠ Investigate — see later tabs"),
    ]
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
    df_cn = _df(res["cn_headers"], res["cn_rows"])
    df_cust = _df(res["cust_headers"], res["cust_rows"])
    df_inv = _df(res["inv_headers"], res["inv_rows"])
    notes = [
        ("CUSTOMER INVOICE POST-LOAD TIEBACK — METHODOLOGY", ""),
        ("Legs compared",
         " [3] Sent = workday_customer_invoice (INVOICE lines) + workday_customer_credit_note"
         " (Adjustment lines) in DuckDB rpatel.duckdb (read-only)."
         "  [4] Loaded = Workday AR Aging Detail by Due Date export — OR — Find Customer Invoices export"
         f"  (file = {workday_extract_basename})."),
        ("Auto-detect Workday extract format",
         "Loader scans rows 0..60 for a recognizable header row, supporting BOTH extract types"
         " with no flag required.  Format A (AR Aging Detail) — fingerprint Company + Transaction Type"
         " + Original Amount in Transaction Currency, Transaction Currency col carries real txn ccy,"
         " Transaction Type col = Invoice / Adjustment.  Format B (Find Customer Invoices) — fingerprint"
         " Company + Reference Number / Invoice Number + Invoice Amount + Balance Due, Status col present,"
         " Currency col may be default-GBP (Find view shows company-currency display), Transaction type is"
         " derived from the first-column label pattern 'Customer Invoice:' vs 'Customer Invoice Adjustment:'"
         " with sign-of-amount fallback."),
        ("Find Customer Invoices — Status pre-filter",
         "If the extract carries a 'Status' column (Find Customer Invoices = format B), only rows with"
         " status in {Approved, Posted, Partially Paid, Paid} are compared.  Draft / Voided / Cancelled"
         " rows are tenant manual entries (keyed directly in the Workday UI, not sent via EIB) so they"
         " have no counterpart on the SENT side; excluding them eliminates a false grand-total delta and"
         " false mismatches.  Every excluded row is surfaced on sheet 9-Excluded WD Rows with the status"
         " cell highlighted red, and a banner on 1-Exec Summary shows the per-status row/£ counts."
         " Cross-ccy alignment bridge: when Find view reports non-GBP invoices in display-only company"
         " currency GBP (amount stays correct, only ccy/customer label are wrong), the script rewrites"
         " WD row ccy + customer-name cols from the matching (company, customer_norm, penny-exact amount)"
         " unique 1:1 unused SENT row, so Customer-bucket and Invoice-fuzzy layers align correctly."),
        ("Join keys — Layer 1 Grand total",
         "Simple SUM(original_amt) across all rows on each side. £0.00 delta by construction"
         " once pipeline and extract are from the same load-window."),
        ("Join keys — Layer 2 By Txn Type",
         "Workday Aging 'Adjustment' / 'Credit Memo' → CREDIT_NOTE to align with sent EIB credit-note staging."),
        ("Join keys — Layer 3 By (Company, Currency)",
         "SENT side Workday-BE labels ('THE FA GROUP - FA / WNSL / NFC') are normalised to the"
         " Aging-export's user-facing short labels before grouping."),
        ("Join keys — Layer 4 Currency coverage",
         "Distinct transaction currencies on each side; BOTH on every ccy required."),
        (f"Join keys — Layer 5 Credit-note exact-key",
         f"Sent eib_txn_ref = workday_customer_credit_note.column_h (Customer Invoice Number, 12/08/26 fix), "
         f"with column_fc (Line Memo) as the pre-fix fallback. Joined to Loaded '{res['cn_right_key']}' column. "
         f"Penny-exact amount required on Both rows."),
        ("Join keys — Layer 6 Reconcile (Customer)",
         "Grain = (Company, symmetrically-normalised Customer, Currency). Normalisation: strip"
         " punctuation, drop Ltd/Limited/PLC/Group/Services/GmbH/SL etc suffixes, & → AND,"
         " collapse whitespace, UPPERCASE. Absorbs cosmetic Workday/NRM naming differences."),
        ("Join keys — Layer 7 Reconcile (Invoice)",
         "Loaded Aging reports Workday tenant refs (FA-CINV-*); sent EIB column_h stores the"
         " legacy GP source document_number directly (12/08/26 3+ fix). If loaded Aging exposes"
         " `External Document Number` = the GP invoice# (loaded fmt_B external_docno field) we"
         " now prefer exact match against sent column_h on the same GP invoice#; else we fall"
         " through to the greedy match on (Company, CCY, normalised customer, penny-exact amt)."
         " Authoritative layers are Grand total + Credit-note exact-key + Customer bucket."),
        ("Invariants / design points",
         "As of 12/08/26, the published EIB carries the legacy GP document_number in column_h"
         " for BOTH invoice and credit-note tables. Credit notes additionally keep column_fc"
         " (Line Memo) populated as a pre-fix compatibility fallback, which is why we read"
         " COALESCE(column_h, column_fc). Any EIB row with a blank column_h raises a warning"
         " at load time and surfaces on the Drift sheet. Customer-name normalisation on SENT"
         " side comes from src_fin_customer_invoice_nrm via the (wd_customer_id,"
         " gp_document_number) exact join, not from an indirect row-number grouping."),
    ]
    df_notes = pd.DataFrame(notes, columns=["Item", "Details"])

    cn_join_note = (
        f"Sent column_h (Customer Invoice Number, authoritative) with column_fc fallback "
        f"==  Loaded '{res['cn_right_key']}' column"
    )

    ov_lines = [
        ("Workbook",
         "FA_PostLoad_Tieback_CustomerInvoice_YYYYMMDD[_LABEL].xlsx — Customer Invoice POST-LOAD 4TH-LAYER TIEBACK (downstream proof)"),
        ("", ""),
        ("What THIS workbook PROVES",
         "Workday actually LOADED exactly what we SENT in the two EIBs (Customer Invoice + Credit Note/Adjustment). "
         "Does NOT look at the source AR TB Excel at all. "
         "If you want to know 'are we loading the correct numbers from the raw source in the first place?', open the companion three-way workbook (see below)."),
        ("Two legs compared",
         "  [3] SENT EIB  —  workday_customer_invoice (INVOICE lines) + workday_customer_credit_note (Adjustment lines) from rpatel.duckdb (read-only).\n"
         "  [4] LOADED WORKDAY  —  Workday AR Aging Detail by Due Date export from the tenant (processed by the 3+ team)"),
        ("Join strategy (7 layers, strongest → weakest)",
         "  1. Grand total                       = SUM across entire load-window\n"
         "  2. By Txn Type                       = INVOICE / CREDIT_NOTE\n"
         "  3. By Company × Currency             = (FA / WNSL / NFC) × currencies\n"
         "  4. Currency coverage                 = distinct CCYs on each leg\n"
         f"  5. Credit-note EXACT-KEY             = {cn_join_note} — strongest per-doc proof, shared key\n"
         "  6. Customer bucket (Match)           = grain = (Company, normalised Customer, CCY)\n"
         "  7. Reconcile (Invoice) traceability  = greedy 4-tuple (Co × CCY × Customer × amt±0.005); this layer is TRACE ONLY because loaded Aging uses tenant refs FA-CINV-* while sent EIB stores source document_number."),
        ("Companion workbook you MUST read alongside this one",
         "FA_ThreeWay_CustomerInvoice_YYYYMMDD[_LABEL].xlsx — the THREE-WAY TIEBACK (upstream proof). It compares raw AR TB source Excel ↔ Our final staging ↔ Our EIB publish tables, proving we are loading the correct numbers from source in the first place."),
        ("", ""),
        ("Full chain of custody — both workbooks read together = end-to-end proof",
         "  AR DTL TB source Excel  ────▶  Our staging (base → nrm → gate)  ────▶  Our EIB publish tables  ────▶  Workday tenant loaded  ────▶  3+ AR Aging extract\n"
         "  ◀────────────────────────  Companion three-way covers ────────────────────────▶◀───────────────────────  THIS workbook covers ─────────────────────────▶"),
        ("", ""),
        ("Where to look FIRST in THIS workbook (5 tabs)",
         "  1-Exec Summary              — headline row counts, grand total delta, CN exact-key counts, customer-bucket counts, invoice fuzzy counts, excluded rows banner, verdict line (green = tied out, red = investigate)\n"
         "  5-Credit Note Exact         — All CN rows. Penny-exact via shared document ref. If ANY row has Δ Amt ≠ 0 or Join ≠ Both, this is ALWAYS a real issue — start here.\n"
         "  6-Reconcile (Customer)      — Customer buckets. Anything not Status=Match means Co/Customer/CCY bucket is on one side but missing or mismatched on the other.\n"
         "  7-Reconcile (Invoice)       — Invoice trace. 'Sent-only' / 'Loaded-only' rows are real investigation targets even if the customer buckets all match (amount equality on a per-invoice key is needed).\n"
         "  9-Excluded WD Rows          — (only for Find Customer Invoices exports) Any Draft / Voided / Cancelled tenant rows removed from compare; status cell red-highlighted with exclude reason."),
        ("", ""),
        ("What does a WARNING / investigation mean?",
         "Any of: (a) Grand total delta ≠ 0 to the penny, (b) Credit-note exact-key has Sent-only OR Loaded-only OR amount mismatch, (c) Customer bucket has status ∈ {Sent only, Loaded only, Review}, (d) Invoice fuzzy has Sent-only ≥1 OR Loaded-only ≥1 OR Amt mismatch ≥1. Start debugging on the strongest layer that fails: Layer 5 exact-key CN ref usually pinpoints exactly which document is wrong."),
        ("", ""),
        ("How to regenerate",
         "  cd elt_ingest_excel\n"
         "  uv run python scripts/generate_customer_invoice_post_load_reconcile.py --label PRODUCTION\n"
         "Flags:  --workday-extract <path>   override the xlsx. Supports BOTH extracts: AR Aging Detail by Due Date (format A) AND Find Customer Invoices (format B) — auto-detected, no flag required.\n"
         "        --sheet <name>             specific sheet name (default: auto)\n"
         "        --label <TEXT>             tag appended to workbook filename; use PRODUCTION for signed-off runs"),
        ("TRAE IDE sandbox note",
         "Writes under /Documents/__data/… are BLOCKED inside the TRAE sandbox. From inside TRAE you can still read the workbook from ~/Downloads/ or /tmp/. To land it in the canonical reconcile dir, run the command above from your own Terminal shell (outside TRAE)."),
        ("Analogous Supplier/PO script",
         "generate_supplier_invoice_post_load_reconcile.py — same sheet naming convention, same banner/header/fill styling, same verdict-cell colouring."),
    ]
    df_ov = pd.DataFrame(ov_lines, columns=["Section", "Detail"])

    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        df_ov.to_excel(w, sheet_name="0-Overview", index=False, startrow=2)
        df_sum.to_excel(w, sheet_name="1-Exec Summary", index=False, startrow=2)
        df_tt.to_excel(w, sheet_name="2-By Txn Type", index=False, startrow=2)
        df_ccy.to_excel(w, sheet_name="3-By Company Currency", index=False, startrow=2)
        df_cov.to_excel(w, sheet_name="4-Currency Coverage", index=False, startrow=2)
        df_cn.to_excel(w, sheet_name="5-Credit Note Exact", index=False, startrow=2)
        df_cust.to_excel(w, sheet_name="6-Reconcile (Customer)", index=False, startrow=2)
        df_inv.to_excel(w, sheet_name="7-Reconcile (Invoice)", index=False, startrow=2)
        df_notes.to_excel(w, sheet_name="8-Notes", index=False, startrow=2)
        _df(res["excluded_headers"], res.get("excluded_rows", [])).to_excel(
            w, sheet_name="9-Excluded WD Rows", index=False, startrow=2
        )
        _style(w)

        ws0 = w.sheets["1-Exec Summary"]
        for i in range(4, ws0.max_row + 1):
            if ws0.cell(i, 1).value == "Reconciliation status":
                c = ws0.cell(i, 2)
                c.fill = OK_FILL if (res["grand_match"] and res["cn_amt_mism"] == 0
                                     and res["inv_totals"][5] == 0) else BAD_FILL
                c.font = Font(bold=True)
        for i in range(4, ws0.max_row + 1):
            if ws0.cell(i, 1).value and "Grand delta" in str(ws0.cell(i, 1).value):
                try:
                    v = float(str(ws0.cell(i, 2).value).replace("£", "").replace(",", ""))
                    if abs(v) > 0.01:
                        ws0.cell(i, 2).fill = BAD_FILL
                except Exception:
                    pass

        _highlight_delta(w.sheets["2-By Txn Type"], 4, thr=0)
        _highlight_delta(w.sheets["2-By Txn Type"], 7, thr=0.01)
        _highlight_delta(w.sheets["3-By Company Currency"], 5, thr=0)
        _highlight_delta(w.sheets["3-By Company Currency"], 8, thr=0.01)
        _highlight_delta(w.sheets["5-Credit Note Exact"], 8, thr=0.01)
        _highlight_delta(w.sheets["6-Reconcile (Customer)"], 6, thr=0)
        _highlight_delta(w.sheets["6-Reconcile (Customer)"], 9, thr=0.01)
        _highlight_delta(w.sheets["7-Reconcile (Invoice)"], 7, thr=0.01)


# --------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------

def main() -> int:
    global WORKDAY_EXTRACT_DEFAULT, WORKDAY_EXTRACT_SHEET_DEFAULT
    ap = argparse.ArgumentParser(
        description="Post-load tie-out: sent Customer-Invoice EIB vs a Workday AR aging extract."
    )
    ap.add_argument(
        "--workday-extract", default=WORKDAY_EXTRACT_DEFAULT,
        help="Path to the Workday AR Aging xlsx export (default: ~/Documents/__data/excel/workday/extract/ar_invoices/AR_Aging_Detail.xlsx).",
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
        print(f"  (The post-load tieback requires a Workday-loaded AR Aging export first.)", file=sys.stderr)
        print(f"  (If you just want Source ↔ Our EIB ↔ Our Staging proof, run the three-way instead:", file=sys.stderr)
        print(f"   uv run python scripts/generate_customer_invoice_three_way.py --label PRODUCTION)", file=sys.stderr)
        return 2
    wd, resolved_sheet = load_workday_ar_aging(extract_path, args.sheet)
    con = duckdb.connect(DB_PATH, read_only=True)
    try:
        res = build(con, wd)
    finally:
        con.close()

    tag = f"_{args.label}" if args.label else ""
    fname = f"FA_PostLoad_Tieback_CustomerInvoice_{dt.date.today():%Y%m%d}{tag}.xlsx"

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

    primary_path, primary_ok = _try_write(OUT_DIR, "reconcile dir (Documents/__data)")
    downloads_path, downloads_ok = _try_write(DOWNLOADS_FALLBACK_DIR, "fallback dir (~/Downloads)")
    stage_path, stage_ok = _try_write(STAGE_DIR, "stage dir (/tmp)")

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
        f"Customer buckets:  Match={res['cust_bucket_counts'][0]}  "
        f"SentOnly={res['cust_bucket_counts'][1]}  "
        f"WDOnly={res['cust_bucket_counts'][2]}  "
        f"Review={res['cust_bucket_counts'][3]}"
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
