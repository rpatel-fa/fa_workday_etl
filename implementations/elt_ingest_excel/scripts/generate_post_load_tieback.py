"""Post-load tie-out: sent Workday EIB payload  vs  loaded Workday control report.

This is the 4th-layer (OPEN-8) reconciliation. It compares what we SENT
(the two EIB tables in DuckDB) against what Workday actually LOADED (the control
report exported back from the tenant). Every figure is computed live from:

  [3] Sent  : workday_supplier_po_commitment_stack (+ _receipt) in rpatel.duckdb,
              joined to src_fin_supplier_po_commitment_stack for the transaction
              currency (the EIB stores local-currency amounts as text).
  [4] Loaded: the Workday control-report workbook (single sheet).

No totals are hardcoded (main-doc invariant "No hardcoded magic totals").

Join keys / vendor-side quirks handled:
  - PO grain    : order_number (control "Purchase Order") == EIB column_j.
                  Line-level tie is by per-PO amount MULTISET, because 3+ renumbered
                  the Workday line refs (control "Line" = 101/201/...) so the EIB
                  3-part service_order_line_id is not present in the control report.
  - Receipt doc : control "Receipt" is labelled 'REC-{PO} for {PO}'; strip ' for ...'
                  to recover our EIB column_h 'REC-{PO}'.

Output: FA_PostLoad_Tieback_YYYYMMDD.xlsx in the reconcile dir. Staged to ~/Downloads first
(sandbox-friendly, nothing is written inside the repo itself). The recovery shutil.copyfile
into reconcile dir succeeds when run from your own shell (outside TRAE sandbox).
"""

from __future__ import annotations

import argparse
import collections
import datetime as dt
import os
import re
import shutil
import sys

import duckdb
import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

DB_PATH = os.path.expanduser("~/Documents/__data/duckdb/rpatel.duckdb")
# Default = the PRODUCTION Workday-loaded extract (authoritative). The earlier
# test-tenant control report was `~/Downloads/FA FInd DC Purchase Orders.xlsx`
# (sheet 'Control Reprot', same totals). Override either with --control-report/--sheet.
CONTROL_REPORT = os.path.expanduser(
    "~/Downloads/FA Find DC Purchase Orders PRODUCTION 9 August 2026 1.xlsx"
)
CONTROL_SHEET = "Sheet1"  # production export sheet; auto-resolves if the name differs
OUT_DIR = os.path.expanduser("~/Documents/__data/excel/workday/reconcile")
STAGE_DIR = "/tmp"

CCY_ORDER = ["GBP", "USD", "EUR", "JPY", "AUD", "AED", "CHF", "CAD"]

_STATUS_SUFFIX_RE = re.compile(r"\s*\((Inactive|Hold|Active|Blocked)\)\s*$", re.IGNORECASE)

# Manual supplier re-key confirmed by ops: Workday supplier ID S-008100 was
# retired in the tenant and merged under the retained ID S-004938. Same legal
# entity (NATIONAL FOOTBALL CENTRE) — only the Workday supplier ID changed.
# We apply this mapping on the SENT side so the (Company, Supplier) grain joins
# cleanly on the loaded side where S-004938 is the surviving ID.
SENT_SUPPLIER_ID_NORMALISATION = {"S-008100": "S-004938"}

# Company label mapping: SENT-side column_p carries the Workday-tenant
# "Business Entity" label ("THE FA GROUP - FA") while the LOADED control
# report carries the user-facing short label ("Football Association").
# We normalise the sent side to the loaded-side label BEFORE grouping so
# (Company, Supplier, Currency) join keys match on both sides.
COMPANY_NAME_NORMALISATION = {
    "THE FA GROUP - FA": "Football Association",
    "THE FA GROUP - WNSL": "Wembley National Stadium",
    "THE FA GROUP - NFC": "National Football Centre",
}


def _strip_supplier_status(name: object) -> str:
    """Remove trailing (Inactive) / (Hold) / (Active) / (Blocked) label.

    The control report renames NFCPO300012338's supplier to 'NATIONAL FOOTBALL
    CENTRE (Hold)' as a side-effect of the manual Workday ID merge. The NRM
    name resolution still returns '(Inactive)' for S-008100. Stripping the
    status suffix leaves the same legal-entity name on both sides."""
    if name is None:
        return ""
    return _STATUS_SUFFIX_RE.sub("", str(name)).strip()


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


def load_control(path: str) -> pd.DataFrame:
    global CONTROL_SHEET
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    # Tolerate a differently-named sheet: the test control report uses
    # 'Control Reprot' (sic); the production export uses 'Sheet1'. Fall back to
    # the single/first sheet and record the resolved name for the Notes tab.
    if CONTROL_SHEET not in wb.sheetnames:
        CONTROL_SHEET = wb.sheetnames[0]
    ws = wb[CONTROL_SHEET]
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(h) for h in rows[0]]
    rename_map = {
        "Company": "company",
        "Purchase Order": "po",
        "Supplier": "supplier_name",
        "Line": "wd_line",
        "Currency": "ccy",
        "Distribution Extended Amount": "po_amt",
        "Approved Receipt Received Amount": "rc_amt",
        "Receipt": "receipt",
        "Purchase Order Line Receipt Status": "rc_status",
        "Document Date": "docdate",
        "Validation Errors and Warnings": "val_err",
    }
    df = pd.DataFrame(rows[1:], columns=hdr).rename(columns=rename_map)
    df["po_amt"] = pd.to_numeric(df["po_amt"], errors="coerce").fillna(0.0)
    df["rc_amt"] = pd.to_numeric(df["rc_amt"], errors="coerce").fillna(0.0)
    df["supplier_name_stripped"] = df.get("supplier_name").apply(_strip_supplier_status)
    df["rc_doc"] = df["receipt"].astype("object")
    mask = df["receipt"].notna() & (df["receipt"].astype(str) != "")
    df.loc[mask, "rc_doc"] = (
        df.loc[mask, "receipt"].astype(str).str.split(" for ").str[0]
    )
    df.loc[~mask, "rc_doc"] = None
    return df


def build(con: duckdb.DuckDBPyConnection, ctl: pd.DataFrame) -> dict:
    con.register("ctl", ctl)
    con.execute(
        """CREATE OR REPLACE TEMP VIEW eib_po AS
             SELECT e.column_j AS po, e.column_ez AS line_key,
                    CAST(e.column_gf AS DECIMAL(18,3)) AS po_amt, s.currency AS ccy
               FROM workday_supplier_po_commitment_stack e
               JOIN src_fin_supplier_po_commitment_stack s
                 ON s.po_line_pos_id = e.column_ez"""
    )
    con.execute(
        """CREATE OR REPLACE TEMP VIEW eib_rc AS
             SELECT r.column_h AS rc_doc, r.column_bu AS line_key,
                    CAST(r.column_bz AS DECIMAL(18,3)) AS rc_amt, s.currency AS ccy
               FROM workday_supplier_po_commitment_stack_receipt r
               JOIN src_fin_supplier_po_commitment_stack s
                 ON s.po_line_pos_id = r.column_bu"""
    )

    def ccy_leg(sent_view: str, loaded_pred: str) -> list:
        return con.execute(
            f"""
            WITH s AS (SELECT ccy, COUNT(*) n, SUM({'po_amt' if 'po' in sent_view else 'rc_amt'}) amt
                         FROM {sent_view} GROUP BY ccy),
                 l AS (SELECT ccy, COUNT(*) n,
                              SUM({'po_amt' if 'po' in sent_view else 'rc_amt'}) amt
                         FROM ctl {loaded_pred} GROUP BY ccy)
            SELECT COALESCE(s.ccy, l.ccy) ccy,
                   COALESCE(s.n, 0), COALESCE(l.n, 0), COALESCE(s.n, 0) - COALESCE(l.n, 0),
                   ROUND(COALESCE(s.amt, 0), 2), ROUND(COALESCE(l.amt, 0), 2),
                   ROUND(COALESCE(s.amt, 0) - COALESCE(l.amt, 0), 2)
              FROM s FULL OUTER JOIN l ON s.ccy = l.ccy""",
        ).fetchall()

    po_ccy = ccy_leg("eib_po", "")
    rc_ccy = ccy_leg("eib_rc", "WHERE rc_doc IS NOT NULL")

    po_grain = con.execute(
        """
        WITH s AS (SELECT po, ROUND(SUM(po_amt),2) amt FROM eib_po GROUP BY po),
             l AS (SELECT po, ROUND(SUM(po_amt),2) amt FROM ctl GROUP BY po)
        SELECT SUM((s.po IS NOT NULL AND l.po IS NOT NULL)::INT),
               SUM((s.po IS NOT NULL AND l.po IS NULL)::INT),
               SUM((s.po IS NULL AND l.po IS NOT NULL)::INT),
               SUM((s.amt IS NOT NULL AND l.amt IS NOT NULL AND ABS(s.amt-l.amt)>0.01)::INT)
          FROM s FULL OUTER JOIN l ON s.po = l.po""",
    ).fetchone()

    rc_grain = con.execute(
        """
        WITH s AS (SELECT rc_doc, ROUND(SUM(rc_amt),2) amt FROM eib_rc GROUP BY rc_doc),
             l AS (SELECT rc_doc, ROUND(SUM(rc_amt),2) amt
                     FROM ctl WHERE rc_doc IS NOT NULL GROUP BY rc_doc)
        SELECT SUM((s.rc_doc IS NOT NULL AND l.rc_doc IS NOT NULL)::INT),
               SUM((s.rc_doc IS NOT NULL AND l.rc_doc IS NULL)::INT),
               SUM((s.rc_doc IS NULL AND l.rc_doc IS NOT NULL)::INT),
               SUM((ABS(COALESCE(s.amt,0)-COALESCE(l.amt,0))>0.01)::INT)
          FROM s FULL OUTER JOIN l ON s.rc_doc = l.rc_doc""",
    ).fetchone()

    # per-PO line-amount multiset equality (catches within-PO line reshuffle)
    eib_ms = collections.defaultdict(list)
    for po, amt in con.execute(
        "SELECT column_j, CAST(column_gf AS DECIMAL(18,2)) FROM workday_supplier_po_commitment_stack"
    ).fetchall():
        eib_ms[po].append(round(float(amt), 2))
    ctl_ms = collections.defaultdict(list)
    for po, amt in zip(ctl["po"], ctl["po_amt"]):
        ctl_ms[po].append(round(float(amt), 2))
    for d in (eib_ms, ctl_ms):
        for k in d:
            d[k].sort()
    multiset_mism = [
        po for po in set(eib_ms) | set(ctl_ms) if eib_ms.get(po) != ctl_ms.get(po)
    ]

    val_err = int((ctl["val_err"].notna() & (ctl["val_err"].astype(str) != "")).sum())
    rc_status = ctl["rc_status"].value_counts().to_dict()

    # ================== per-PO recon (Distribution + Approved Receipt) =====
    # Aggregate PO side and receipt side separately first, then merge — avoids
    # n×m row explosion if a PO has multiple lines AND multiple receipt lines.
    sent_po_core = con.execute(
        """
        SELECT e.column_j po, e.column_p company_short, s.currency ccy,
               COUNT(*) po_lines_sent,
               ROUND(SUM(CAST(e.column_gf AS DECIMAL(18,2))),2) po_sent
          FROM workday_supplier_po_commitment_stack e
          JOIN src_fin_supplier_po_commitment_stack s
            ON s.po_line_pos_id = e.column_ez
         GROUP BY e.column_j, e.column_p, s.currency
        """
    ).fetchdf()

    sent_rc_core = con.execute(
        """
        SELECT s.order_number po, s.currency ccy,
               COUNT(*) rc_lines_sent,
               ROUND(SUM(CAST(r.column_bz AS DECIMAL(18,2))),2) rc_sent
          FROM workday_supplier_po_commitment_stack_receipt r
          JOIN src_fin_supplier_po_commitment_stack s
            ON s.po_line_pos_id = r.column_bu
         GROUP BY s.order_number, s.currency
        """
    ).fetchdf()

    sent_po_df = sent_po_core.merge(sent_rc_core, on=["po", "ccy"], how="left")
    sent_po_df["rc_lines_sent"] = sent_po_df["rc_lines_sent"].fillna(0).astype(int)
    sent_po_df["rc_sent"] = sent_po_df["rc_sent"].fillna(0.0).astype(float).round(2)
    sent_po_df["company_short"] = sent_po_df["company_short"].replace(
        COMPANY_NAME_NORMALISATION
    )

    # Loaded company label normalisation: loaded uses short labels ("Football Association")
    # while column_p uses the Workday company_ID label ("THE FA GROUP - FA"). We don't
    # need company on the PO join key (po+ccy is unique), so align names downstream
    # using a (po → loaded_company) lookup for display consistency.
    loaded_po_core = (
        ctl.groupby(["company", "po", "ccy"])
           .agg(
               po_lines_loaded=("po_amt", "size"),
               po_loaded=("po_amt", lambda s: round(float(s.sum()), 2)),
               rc_loaded=("rc_amt", lambda s: round(float(s.sum()), 2)),
           )
           .reset_index()
    )
    rc_lines_loaded_series = (
        ctl.assign(_gt0=((ctl["rc_amt"].fillna(0.0) > 0.001).astype(int)))
           .groupby(["company", "po", "ccy"])["_gt0"]
           .sum()
           .reset_index(name="rc_lines_loaded")
    )
    loaded_po_df = loaded_po_core.merge(
        rc_lines_loaded_series, on=["company", "po", "ccy"], how="left"
    )
    loaded_po_df["rc_lines_loaded"] = (
        loaded_po_df["rc_lines_loaded"].fillna(0).astype(int)
    )

    merged_po = sent_po_df.merge(
        loaded_po_df, on=["po", "ccy"], how="outer", indicator=True
    )
    # Prefer the company label from the loaded side (matches supplier recon's
    # company label since both come from the control-report "Company" column);
    # fall back to the sent company_short for any sent-only row.
    merged_po["company"] = merged_po["company"].fillna(merged_po["company_short"])
    merged_po = merged_po.drop(columns=["company_short"])
    for col in ("po_lines_sent", "rc_lines_sent", "po_lines_loaded", "rc_lines_loaded"):
        merged_po[col] = merged_po[col].fillna(0).astype(int)
    for col in ("po_sent", "rc_sent", "po_loaded", "rc_loaded"):
        merged_po[col] = merged_po[col].fillna(0.0).astype(float).round(2)
    merged_po["po_lines_delta"] = merged_po["po_lines_sent"] - merged_po["po_lines_loaded"]
    merged_po["rc_lines_delta"] = merged_po["rc_lines_sent"] - merged_po["rc_lines_loaded"]
    merged_po["po_amt_delta"] = (merged_po["po_sent"] - merged_po["po_loaded"]).round(2)
    merged_po["rc_amt_delta"] = (merged_po["rc_sent"] - merged_po["rc_loaded"]).round(2)

    po_recon_headers = [
        "Company", "PO", "Currency",
        "PO lines sent", "PO lines loaded", "Δ lines",
        "PO Distrib sent", "PO Distrib loaded", "Δ PO Distrib",
        "Recpt lines sent", "Recpt lines loaded", "Δ Recpt lines",
        "Recpt Apprvd sent", "Recpt Apprvd loaded", "Δ Recpt Apprvd",
        "Join side",
    ]

    def _po_row(r) -> list:
        side = {"both": "Both", "left_only": "Sent only", "right_only": "Loaded only"}[r["_merge"]]
        return [
            r["company"], r["po"], r["ccy"],
            int(r["po_lines_sent"]), int(r["po_lines_loaded"]), int(r["po_lines_delta"]),
            float(r["po_sent"]), float(r["po_loaded"]), float(r["po_amt_delta"]),
            int(r["rc_lines_sent"]), int(r["rc_lines_loaded"]), int(r["rc_lines_delta"]),
            float(r["rc_sent"]), float(r["rc_loaded"]), float(r["rc_amt_delta"]),
            side,
        ]

    po_recon_rows = [_po_row(r) for _, r in merged_po.sort_values(["company", "po"]).iterrows()]
    po_mismatches_n = int(
        sum(
            1
            for r in po_recon_rows
            if r[15] != "Both"
            or abs(r[8]) > 0.01
            or abs(r[14]) > 0.01
            or r[5] != 0
            or r[11] != 0
        )
    )

    # ================== per-Supplier recon (manual ID swap absorbed) =====
    # sent side: (po, supplier_id, ccy) — compute PO + receipt aggregates SEPARATELY
    # first, then LEFT-JOIN — avoids n×m multiplication when a PO has multiple
    # PO lines AND multiple receipt lines (the LEFT JOIN used to re-add sp rows
    # once per matching sr row, double-counting po_lines/po_amt).
    sup_id_name = dict(
        con.execute(
            """SELECT DISTINCT wd_supplier_id, wd_supplier_name
                 FROM src_fin_supplier_po_commitment_stack_nrm
                WHERE wd_supplier_id IS NOT NULL"""
        ).fetchall()
    )

    sent_sup_po = con.execute(
        """
        SELECT e.column_p company_short, s.order_number po,
               e.column_q sup_id, s.currency ccy,
               COUNT(DISTINCT s.order_number) pos,
               COUNT(*) po_lines,
               ROUND(SUM(CAST(e.column_gf AS DECIMAL(18,2))),2) po_amt
          FROM workday_supplier_po_commitment_stack e
          JOIN src_fin_supplier_po_commitment_stack s
            ON s.po_line_pos_id = e.column_ez
         GROUP BY e.column_p, s.order_number, e.column_q, s.currency
        """
    ).fetchdf()

    sent_sup_rc = con.execute(
        """
        SELECT s2.order_number po, s2.currency ccy, s2.wd_supplier_id sup_id,
               COUNT(*) rc_lines,
               ROUND(SUM(CAST(r2.column_bz AS DECIMAL(18,2))),2) rc_amt
          FROM workday_supplier_po_commitment_stack_receipt r2
          JOIN src_fin_supplier_po_commitment_stack s2
            ON s2.po_line_pos_id = r2.column_bu
         GROUP BY s2.order_number, s2.currency, s2.wd_supplier_id
        """
    ).fetchdf()

    sent_sup_core = sent_sup_po.merge(
        sent_sup_rc, on=["po", "ccy", "sup_id"], how="left"
    )
    sent_sup_core["rc_lines"] = sent_sup_core["rc_lines"].fillna(0).astype(int)
    sent_sup_core["rc_amt"] = sent_sup_core["rc_amt"].fillna(0.0).astype(float).round(2)
    sent_sup_core["company_short"] = sent_sup_core["company_short"].replace(
        COMPANY_NAME_NORMALISATION
    )
    sent_sup_core["sup_id_norm"] = sent_sup_core["sup_id"].replace(
        SENT_SUPPLIER_ID_NORMALISATION
    )
    sent_sup_core["sup_name"] = (
        sent_sup_core["sup_id"]
        .map(lambda sid: sup_id_name.get(sid, ""))
        .fillna("")
    )
    sent_sup_core["sup_name_stripped"] = sent_sup_core["sup_name"].apply(
        _strip_supplier_status
    )

    # Aggregate sent by (company_short, sup_name_stripped, ccy) — roll POs.
    # NOTE: multiple distinct Workday supplier IDs may share the same NRM name
    # (e.g. BT, PINSENT MASONS each have 2-3 S-IDs). The loaded side joins on
    # name/company/ccy only, so we must NOT include sup_id_norm in the grouping
    # key — it would split the sent row into multiple rows and not match loaded.
    # For display, collect the normalised S-IDs as a comma-separated list.
    def _cat_ids(ids):
        uniq = []
        for x in ids:
            if x and str(x).strip() and str(x).strip() not in uniq:
                uniq.append(str(x).strip())
        return ", ".join(uniq)

    sent_sup_agg = (
        sent_sup_core.groupby(
            ["company_short", "sup_name_stripped", "ccy"], dropna=False
        )
        .agg(
            sup_id_norm=("sup_id_norm", _cat_ids),
            sent_pos=("po", "nunique"),
            sent_po_lines=("po_lines", "sum"),
            sent_po_amt=("po_amt", lambda s: round(float(pd.Series(s).sum()), 2)),
            sent_rc_lines=("rc_lines", lambda s: int(pd.Series(s).fillna(0).sum())),
            sent_rc_amt=("rc_amt", lambda s: round(float(pd.Series(s).fillna(0).sum()), 2)),
        )
        .reset_index()
    )
    sent_sup_agg["sup_name_stripped"] = sent_sup_agg["sup_name_stripped"].fillna("")
    sent_sup_agg["sup_id_norm"] = sent_sup_agg["sup_id_norm"].fillna("")

    # loaded side: (company, supplier_name_stripped, ccy)
    loaded_sup_df = (
        ctl.assign(rc_lines_positive=((ctl["rc_amt"].fillna(0.0) > 0.001).astype(int)))
           .groupby(["company", "supplier_name_stripped", "ccy"])
           .agg(
               loaded_pos=("po", "nunique"),
               loaded_po_lines=("po_amt", "size"),
               loaded_po_amt=("po_amt", lambda s: round(float(s.sum()), 2)),
               loaded_rc_lines=("rc_lines_positive", "sum"),
               loaded_rc_amt=("rc_amt", lambda s: round(float(s.sum()), 2)),
           )
           .reset_index()
           .rename(
               columns={
                   "company": "company_short",
                   "supplier_name_stripped": "sup_name_stripped",
               }
           )
    )

    merged_sup = sent_sup_agg.merge(
        loaded_sup_df,
        on=["company_short", "sup_name_stripped", "ccy"],
        how="outer",
        indicator=True,
    )
    for col in ("sent_pos", "loaded_pos", "sent_po_lines", "loaded_po_lines",
                "sent_rc_lines", "loaded_rc_lines"):
        merged_sup[col] = merged_sup[col].fillna(0).astype(int)
    for col in ("sent_po_amt", "loaded_po_amt", "sent_rc_amt", "loaded_rc_amt"):
        merged_sup[col] = merged_sup[col].fillna(0.0).astype(float).round(2)
    merged_sup["pos_delta"] = merged_sup["sent_pos"] - merged_sup["loaded_pos"]
    merged_sup["po_lines_delta"] = merged_sup["sent_po_lines"] - merged_sup["loaded_po_lines"]
    merged_sup["rc_lines_delta"] = merged_sup["sent_rc_lines"] - merged_sup["loaded_rc_lines"]
    merged_sup["po_amt_delta"] = (merged_sup["sent_po_amt"] - merged_sup["loaded_po_amt"]).round(2)
    merged_sup["rc_amt_delta"] = (merged_sup["sent_rc_amt"] - merged_sup["loaded_rc_amt"]).round(2)

    sup_recon_headers = [
        "Company", "Supplier ID (sent norm)", "Supplier (legal name, status stripped)",
        "Currency",
        "POs sent", "POs loaded", "Δ POs",
        "PO lines sent", "PO lines loaded", "Δ lines",
        "PO Distrib sent", "PO Distrib loaded", "Δ PO Distrib",
        "Recpt lines sent", "Recpt lines loaded", "Δ Recpt lines",
        "Recpt Apprvd sent", "Recpt Apprvd loaded", "Δ Recpt Apprvd",
        "Join side",
    ]

    def _sup_row(r) -> list:
        side = {"both": "Both", "left_only": "Sent only", "right_only": "Loaded only"}[r["_merge"]]
        return [
            r["company_short"],
            r.get("sup_id_norm") or "",
            r["sup_name_stripped"] or "",
            r["ccy"],
            int(r["sent_pos"]), int(r["loaded_pos"]), int(r["pos_delta"]),
            int(r["sent_po_lines"]), int(r["loaded_po_lines"]), int(r["po_lines_delta"]),
            float(r["sent_po_amt"]), float(r["loaded_po_amt"]), float(r["po_amt_delta"]),
            int(r["sent_rc_lines"]), int(r["loaded_rc_lines"]), int(r["rc_lines_delta"]),
            float(r["sent_rc_amt"]), float(r["loaded_rc_amt"]), float(r["rc_amt_delta"]),
            side,
        ]

    sup_recon_rows = [
        _sup_row(r)
        for _, r in merged_sup.sort_values(
            ["company_short", "sup_name_stripped", "ccy"]
        ).iterrows()
    ]
    sup_mismatches_n = int(
        sum(
            1
            for r in sup_recon_rows
            if r[19] != "Both"
            or abs(r[12]) > 0.01
            or abs(r[18]) > 0.01
            or r[6] != 0
            or r[9] != 0
            or r[15] != 0
        )
    )

    return {
        "po_ccy": po_ccy,
        "rc_ccy": rc_ccy,
        "po_grain": po_grain,
        "rc_grain": rc_grain,
        "multiset_mism": multiset_mism,
        "val_err": val_err,
        "rc_status": rc_status,
        "ctl_rows": len(ctl),
        "ctl_pos": ctl["po"].nunique(),
        "ctl_rc_rows": int((ctl["rc_doc"].notna()).sum()),
        "ctl_rc_docs": ctl.loc[ctl["rc_doc"].notna(), "rc_doc"].nunique(),
        "po_recon_headers": po_recon_headers,
        "po_recon_rows": po_recon_rows,
        "po_mismatches_n": po_mismatches_n,
        "sup_recon_headers": sup_recon_headers,
        "sup_recon_rows": sup_recon_rows,
        "sup_mismatches_n": sup_mismatches_n,
    }


# ------------------------------- workbook -----------------------------------

HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FILL = PatternFill("solid", fgColor="C6EFCE")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
TITLE = Font(bold=True, size=13)


def _sheet(wb, name, headers, rows, ok_col=None, ok_cols=None):
    ws = wb.create_sheet(name)
    for j, h in enumerate(headers, 1):
        c = ws.cell(1, j, h)
        c.fill, c.font = HDR_FILL, HDR_FONT
        c.alignment = Alignment(horizontal="center")
    for i, r in enumerate(rows, 2):
        for j, v in enumerate(r, 1):
            ws.cell(i, j, v)
        # single-ok-col legacy mode (ok_col points to a single delta column)
        if ok_col is not None and abs(_num(r[ok_col])) > 0.01:
            ws.cell(i, ok_col + 1).fill = BAD_FILL
        # multi-delta highlight mode (ok_cols is list of (col_idx, threshold), threshold 0 for ints)
        if ok_cols:
            for (cidx, thr) in ok_cols:
                try:
                    val = _num(r[cidx])
                except Exception:
                    continue
                bad = abs(val) > thr if thr > 0 else (val != 0)
                if bad:
                    ws.cell(i, cidx + 1).fill = BAD_FILL
    for j, h in enumerate(headers, 1):
        ws.column_dimensions[
            openpyxl.utils.get_column_letter(j)
        ].width = max(14, len(str(h)) + 2)
    return ws


def write_workbook(res: dict, out_path: str) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---- Sheet 0 — Overview (THIS workbook vs companion three-way) -----------
    ws_ov = wb.create_sheet("0-Overview")
    ov_rows = [
        ("Workbook", os.path.basename(out_path) + " — PO + Receipt POST-LOAD 4TH-LAYER TIEBACK (downstream proof)"),
        ("", ""),
        ("What THIS workbook PROVES",
         "Workday actually LOADED exactly what we SENT in the two PO + Receipt EIBs. "
         "Does NOT look at the raw Proactis source or Finance target at all. "
         "If you want to know 'are we loading the correct numbers from source in the first place?', open the companion three-way workbook (see below)."),
        ("Two legs compared",
         "  [3] SENT EIB  —  workday_purchase_order (service lines) + workday_goods_receipt (lines/docs) from rpatel.duckdb read-only.\n"
         "  [4] LOADED WORKDAY  —  Workday 'FA Find DC Purchase Orders' control-report xlsx export from the tenant (3+ team)."),
        ("Join strategy (strongest → weakest)",
         "  1. PO & Receipt per currency  — rows and amounts across each ccy\n"
         "  2. Per-PO grain  — Both / Sent-only / Loaded-only / Amt-mismatch\n"
         "  3. Per-PO line-amount multiset  — individual service lines match (same line amounts, same counts)\n"
         "  4. Per-PO + Per-Supplier recon (banner on Exec Summary)\n"
         "  5. Workday validation-errors count  (should be 0 on successful load)"),
        ("Companion workbook you MUST read alongside this one",
         "FA_ThreeWay_Reconciliation_YYYYMMDD.xlsx — the THREE-WAY REVIEW (upstream proof). It compares Finance target ↔ Proactis source ↔ Our EIB publish tables, proving we are loading the correct numbers from source in the first place."),
        ("", ""),
        ("Full chain of custody — both workbooks read together = end-to-end proof",
         "  Finance target  ──▶  Proactis source  ──▶  Our staging  ──▶  Our EIB publish tables  ──▶  Workday tenant loaded  ──▶  3+ PO/Receipt extract\n"
         "  ◀────────────────────────  Companion three-way covers ─────────────────────────▶◀────────────────  THIS workbook covers ─────────────────▶"),
        ("", ""),
        ("Where to look FIRST in THIS workbook (3 tabs)",
         "  Exec Summary           — Result verdict cell (green = CLEAN / red = REVIEW), grand PO service-line sent/loaded counts, receipt sent/loaded, per-PO tie counts, multiset mismatches, validation errors.\n"
         "  PO by Currency         — any per-currency row with rows/≠ 0 or |amount delta| > 0.01 is bad (highlighted red).\n"
         "  Per-PO Tie             — rows marked Sent-only / Loaded-only / Amt-mismatch are always real investigation targets."),
        ("", ""),
        ("What does a REVIEW mean?",
         "Any of: (a) PO per-currency rows or amount delta ≠ 0, (b) Receipt per-currency rows or amount delta ≠ 0, (c) per-PO grain has Sent-only / Loaded-only / Amt-mismatch, (d) ≥1 PO line-amount multiset mismatch, (e) Workday validation errors > 0. Start with the strongest layer that fails first."),
        ("", ""),
        ("How to regenerate",
         "  cd elt_ingest_excel\n"
         "  uv run python scripts/generate_post_load_tieback.py --control-report <path> --sheet <name> [--label PRODUCTION]"),
        ("TRAE IDE sandbox note",
         "Writes under /Documents/__data/… are BLOCKED inside the TRAE sandbox. From inside TRAE you can still read the workbook from ~/Downloads/ or /tmp/. To land it in the canonical reconcile dir, run the command above from your own Terminal shell (outside TRAE)."),
        ("Analogous Invoice script",
         "generate_supplier_invoice_post_load_reconcile.py — same 0-Overview / 1-Exec / matrix / Notes layout, same banner + colouring conventions."),
    ]
    # Header row
    for j, h in enumerate(["Section", "Detail"], 1):
        c = ws_ov.cell(1, j, h)
        c.fill, c.font, c.alignment = HDR_FILL, HDR_FONT, Alignment(horizontal="center")
    for i, r in enumerate(ov_rows, 2):
        for j, v in enumerate(r, 1):
            cell = ws_ov.cell(i, j, v)
            if j == 2:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
    # Column widths
    ws_ov.column_dimensions["A"].width = 46
    ws_ov.column_dimensions["B"].width = 110
    # Auto row heights for wrapped text
    for i in range(2, ws_ov.max_row + 1):
        detail = ws_ov.cell(i, 2).value or ""
        lines = detail.count("\n") + 1 if detail else 1
        ws_ov.row_dimensions[i].height = max(15, lines * 14)
    ws_ov.freeze_panes = "A2"

    # Sheet 1 — Exec Summary
    ws = wb.create_sheet("Exec Summary")
    ws["A1"] = "Post-Load Tie-Out — Sent EIB vs Workday Control Report"
    ws["A1"].font = TITLE
    ws["A2"] = f"Generated {dt.date.today():%Y-%m-%d}  |  source [4] = {os.path.basename(CONTROL_REPORT)}"
    po_ok = all(abs(r[6]) < 0.01 and r[3] == 0 for r in res["po_ccy"])
    rc_ok = all(abs(r[6]) < 0.01 and r[3] == 0 for r in res["rc_ccy"])
    b, so, lo, mm = res["po_grain"]
    verdict = (
        "CLEAN — sent == loaded, to the penny and to the line"
        if po_ok and rc_ok and so == 0 and lo == 0 and mm == 0
        and not res["multiset_mism"] and res["val_err"] == 0
        else "REVIEW — see sheets for deltas"
    )
    summary = [
        ("Result", verdict),
        ("PO service lines  (sent / loaded)", f"{sum(r[1] for r in res['po_ccy'])} / {res['ctl_rows']}"),
        ("Distinct POs  (loaded)", res["ctl_pos"]),
        ("Receipt lines  (sent / loaded)", f"{sum(r[1] for r in res['rc_ccy'])} / {res['ctl_rc_rows']}"),
        ("Receipt docs  (loaded)", res["ctl_rc_docs"]),
        ("PO amount delta  (all currencies)", "£0.00 per currency" if po_ok else "NON-ZERO"),
        ("Receipt amount delta  (all currencies)", "£0.00 per currency" if rc_ok else "NON-ZERO"),
        ("Per-PO tie  (both / sent-only / loaded-only / amt-mismatch)", f"{b} / {so} / {lo} / {mm}"),
        ("Per-PO line-amount multiset mismatches", len(res["multiset_mism"])),
        ("Workday validation errors on load", res["val_err"]),
    ]
    for i, (k, v) in enumerate(summary, 4):
        ws.cell(i, 1, k).font = Font(bold=True)
        cell = ws.cell(i, 2, v)
        if k == "Result":
            cell.fill = OK_FILL if verdict.startswith("CLEAN") else BAD_FILL
            cell.font = Font(bold=True)
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 62

    # per-PO + per-Supplier recon banner
    po_recon_ok = res["po_mismatches_n"] == 0
    sup_recon_ok = res["sup_mismatches_n"] == 0
    banner = [
        ("Per-PO recon  (Distrib Ext + Apprvd Recv)  mismatches", res["po_mismatches_n"]),
        ("Per-Supplier recon  (Distrib Ext + Apprvd Recv)  mismatches", res["sup_mismatches_n"]),
        ("  applied: manual Workday S-ID normalisation",
         ", ".join(f"{k} → {v}" for k, v in SENT_SUPPLIER_ID_NORMALISATION.items())
         or "(none)"),
        ("  applied: supplier-name status-suffix strip",
         "(Inactive) / (Hold) / (Active) / (Blocked)"),
    ]
    for i, (k, v) in enumerate(banner, len(summary) + 4):
        ws.cell(i, 1, k).font = Font(bold=True)
        c = ws.cell(i, 2, v)
        if "recon" in k:
            c.fill = OK_FILL if int(v) == 0 else BAD_FILL
            c.font = Font(bold=True)

    ccy_hdr = ["Currency", "Sent rows", "Loaded rows", "Δ rows",
               "Sent amount", "Loaded amount", "Δ amount"]

    def ccy_sorted(rows):
        return sorted(rows, key=lambda r: (CCY_ORDER.index(r[0]) if r[0] in CCY_ORDER else 99))

    _sheet(wb, "PO by currency", ccy_hdr, ccy_sorted(res["po_ccy"]), ok_col=6)
    _sheet(wb, "Receipt by currency", ccy_hdr, ccy_sorted(res["rc_ccy"]), ok_col=6)

    # Sheet — Grain tie
    grain_rows = [
        ["Purchase Order (order_number)", *res["po_grain"],
         "line tie by per-PO amount multiset"],
        ["Receipt document (REC-*)", *res["rc_grain"],
         "control label 'REC-x for x' normalised to 'REC-x'"],
    ]
    _sheet(wb, "Grain tie", ["Grain", "Both", "Sent only", "Loaded only",
                             "Amt mismatch (>1p)", "Note"], grain_rows)

    # Sheet — per-PO recon (Distribution Ext Amount + Approved Receipt Amount)
    po_deltas = [
        (5, 0),    # Δ PO lines
        (8, 0.01), # Δ PO Distrib
        (11, 0),   # Δ Receipt lines
        (14, 0.01), # Δ Receipt Approved
    ]
    _sheet(wb, "Recon (per PO)", res["po_recon_headers"],
           res["po_recon_rows"], ok_cols=po_deltas)

    # Sheet — per-Supplier recon (manual S-ID normalisation + status-suffix strip)
    sup_deltas = [
        (6, 0),    # Δ PO count
        (9, 0),    # Δ PO lines
        (12, 0.01),# Δ PO Distrib
        (15, 0),   # Δ Receipt lines
        (18, 0.01),# Δ Receipt Approved
    ]
    _sheet(wb, "Recon (per Supplier)", res["sup_recon_headers"],
           res["sup_recon_rows"], ok_cols=sup_deltas)

    # Sheet — Load quality
    lq = [["Workday validation errors", res["val_err"], "must be 0"]]
    for k, v in res["rc_status"].items():
        lq.append([f"Line receipt status: {k}", v, ""])
    lq.append(["Per-PO line-amount multiset mismatches", len(res["multiset_mism"]),
               "0 = every line loaded at the right amount"])
    _sheet(wb, "Load quality", ["Metric", "Value", "Note"], lq)

    # Sheet — Notes / methodology
    wsn = wb.create_sheet("Notes")
    notes = [
        "POST-LOAD TIE-OUT METHODOLOGY (4th layer / OPEN-8)",
        "",
        "Legs compared:",
        "  [3] Sent   = workday_supplier_po_commitment_stack (+ _receipt), DuckDB rpatel,",
        "               joined to src_fin_supplier_po_commitment_stack for transaction currency.",
        f"  [4] Loaded = {CONTROL_REPORT} (sheet '{CONTROL_SHEET}').",
        "",
        "Join keys:",
        "  PO grain     : order_number  (control 'Purchase Order' == EIB column_j).",
        "  Line grain   : per-PO amount MULTISET — 3+ renumbered Workday line refs",
        "                 (control 'Line' = 101/201/...) so the EIB 3-part service_order_line_id",
        "                 is not present in the control report; multiset equality proves every",
        "                 line loaded at the correct amount without a shared line key.",
        "  Receipt doc  : control 'Receipt' labelled 'REC-{PO} for {PO}' -> strip ' for ...'",
        "                 to recover EIB column_h 'REC-{PO}'.",
        "",
        "Vendor (3+) changes confirmed absorbed with zero reconciliation impact:",
        "  - PO due-date blanked (column_ae NULL).",
        "  - Receipt dates set on the loaded receipts.",
        "  - Spreadsheet-key technical changes — did not break line identity.",
        "  - Receipt label rename — cosmetic; prefix + amounts reconcile.",
        "",
        "Per-Supplier recon tab normalisations (applied to produce a clean Both/Only split):",
        "  [0] Company label normalisation on sent side: Workday Business Entity labels",
        "      ('THE FA GROUP - FA / WNSL / NFC') are mapped to the control-report's",
        "      user-facing short labels ('Football Association', 'Wembley National Stadium',",
        "      'National Football Centre') before grouping. Necessary because the EIB",
        "      column_p carries the tenant BE label while the control export carries",
        "      the user-facing label.",
        f"  [A] Supplier ID manual merge on sent side: "
        + (", ".join(f"{k} → {v}" for k, v in SENT_SUPPLIER_ID_NORMALISATION.items()) if SENT_SUPPLIER_ID_NORMALISATION else "(none)")
        + ". Same legal entity, Workday-supplier ID change only.",
        "  [B] Supplier-name status-suffix strip (both sides, symmetric):",
        "      remove trailing (Inactive) / (Hold) / (Active) / (Blocked).",
        "      One row NFCPO300012338 otherwise spans sent-only and loaded-only buckets",
        "      due to (Inactive) → (Hold) label drift introduced by the manual [A] merge.",
        "  [C] When several S-IDs share the same NRM legal name (e.g. BT has 3,",
        "      PINSENT MASONS has 2), all S-ID rows under a shared name are rolled",
        "      into one (Company, Supplier, CCY) bucket on the sent side before",
        "      joining, because the loaded control-report only carries supplier NAME",
        "      — no supplier ID column to match on.",
        "",
        "Amounts are transaction-currency (local). No totals are hardcoded; every figure",
        "is computed live at generation time.",
    ]
    for i, line in enumerate(notes, 1):
        c = wsn.cell(i, 1, line)
        if i == 1:
            c.font = TITLE
    wsn.column_dimensions["A"].width = 96

    wb.save(out_path)


def main() -> int:
    global CONTROL_REPORT, CONTROL_SHEET
    ap = argparse.ArgumentParser(
        description="Post-load tie-out: sent EIB vs a Workday-loaded PO/receipt extract."
    )
    ap.add_argument(
        "--control-report", default=CONTROL_REPORT,
        help="Path to the Workday-loaded extract xlsx (default: the test control report).",
    )
    ap.add_argument(
        "--sheet", default=None,
        help="Sheet name to read (default: auto — the named sheet if present, else the first).",
    )
    ap.add_argument(
        "--label", default=None,
        help="Optional tag added to the output filename (e.g. PRODUCTION) so runs don't clobber.",
    )
    args = ap.parse_args()

    CONTROL_REPORT = os.path.expanduser(args.control_report)
    if args.sheet:
        CONTROL_SHEET = args.sheet
    if not os.path.exists(CONTROL_REPORT):
        print(f"Control report not found: {CONTROL_REPORT}", file=sys.stderr)
        return 2
    ctl = load_control(CONTROL_REPORT)  # resolves CONTROL_SHEET if the name differs
    con = duckdb.connect(DB_PATH, read_only=True)
    try:
        res = build(con, ctl)
    finally:
        con.close()

    tag = f"_{args.label}" if args.label else ""
    fname = f"FA_PostLoad_Tieback_{dt.date.today():%Y%m%d}{tag}.xlsx"

    def _try_write(target_dir: str, friendly: str) -> tuple[str | None, bool]:
        try:
            os.makedirs(target_dir, exist_ok=True)
            dest = os.path.join(target_dir, fname)
            write_workbook(res, dest)
            print(f"Wrote to {friendly}: {dest}")
            return dest, True
        except Exception as e:
            print(f"Could not write to {friendly} ({type(e).__name__}: {e}).", file=sys.stderr)
            return None, False

    DOWNLOADS_FALLBACK_DIR = os.path.expanduser("~/Downloads")

    # 1) PRIMARY — the reconcile dir in Documents/__data
    primary_path, primary_ok = _try_write(OUT_DIR, "reconcile dir (Documents/__data)")

    # 2) FALLBACK — ~/Downloads (user-visible on all platforms; TRAE sandbox permits)
    downloads_path, downloads_ok = _try_write(DOWNLOADS_FALLBACK_DIR, "fallback dir (~/Downloads)")

    # 3) STAGE — /tmp (last resort)
    stage_path, stage_ok = _try_write(STAGE_DIR, "stage dir (/tmp)")

    # Recovery: try shutil.copyfile from Downloads or /tmp to OUT_DIR
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

    po_ok = all(abs(r[6]) < 0.01 and r[3] == 0 for r in res["po_ccy"])
    rc_ok = all(abs(r[6]) < 0.01 and r[3] == 0 for r in res["rc_ccy"])
    b, so, lo, mm = res["po_grain"]
    print(
        f"PO ccy tie: {'£0 all ccy' if po_ok else 'DELTA'} | "
        f"Receipt ccy tie: {'£0 all ccy' if rc_ok else 'DELTA'} | "
        f"per-PO both={b} sent_only={so} loaded_only={lo} amt_mismatch={mm} | "
        f"multiset_mismatch={len(res['multiset_mism'])} | val_err={res['val_err']} | "
        f"per-PO recon mismatches={res['po_mismatches_n']} | "
        f"per-Supplier recon mismatches={res['sup_mismatches_n']}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
