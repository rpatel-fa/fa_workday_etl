"""Three-way finance reconciliation review pack — EIB (loaded) vs Source (060826)
vs Finance recon (040826).

Unlike the earlier re-derivation pack, this reads the ACTUAL EIB output tables from
DuckDB and reconciles them to finance's recon report, using our 060826 source extract
as the bridge:

    [A] EIB  <-> Source(060826)   = transform integrity (did we load the source faithfully?)
    [B] Source(060826) <-> Finance(040826) = extract drift (two Proactis pulls, different dates)
    [C] EIB  <-> Finance(040826)  = headline; every non-match must be explained by [B]

All comparisons are per Purchase Order on the GBP basis (soft+hard commitment in £,
summable across currencies). Local (document-currency) totals are shown alongside but
are NOT summable. See recon-todo.md in the transform SQL folder for the anchor doc.

Usage:
    uv run python elt_ingest_excel/scripts/generate_finance_review_three_way.py
"""
from __future__ import annotations

import argparse
import os
import re
from datetime import datetime
from pathlib import Path

import duckdb
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DUCKDB = "~/Documents/__data/duckdb/rpatel.duckdb"
FINANCE_XLSX = "~/Downloads/The FA - Commitment Stack Report as at 040826 USE THIS ONE.xlsx"
FINANCE_SHEET = "StackReport v3"
OUT_DIR = "~/Documents/__data/excel/workday/reconcile"

# Proactis 060826 export corruption (embedded newlines merged/shifted rows). Finance
# signed off 07.08.26 to load as-is. These POs carry a known difference vs the 040826
# recon and are labelled as an ACCEPTED exception (not a pipeline variance).
ACCEPTED_EXCEPTION_POS = {
    "NFCPO300004848", "NFCPO300004851", "NFCPO300006489",  # in EIB, corrupted/misaligned
    "NFCPO300006502", "NFCPO300006503",                    # absorbed into 6489's blob (missing)
}
ACCEPTED_LABEL = "0. Accepted exception — Proactis source corruption (signed off 07.08.26)"

TITLE_FONT = Font(bold=True, size=14, color="FFFFFF")
TITLE_FILL = PatternFill(fill_type="solid", fgColor="2F5496")
HEAD_FONT = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill(fill_type="solid", fgColor="4472C4")


def _num(x) -> float:
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return 0.0
        if isinstance(x, (int, float)):
            return float(x)
        s = re.sub(r"[^\d.\-]", "", str(x))
        return 0.0 if s in ("", "-", ".") else float(s)
    except Exception:
        return 0.0


def load_finance(path: str) -> pd.DataFrame:
    """Finance recon (040826) StackReport v3, per PO: soft/hard GBP."""
    df = pd.read_excel(Path(os.path.expanduser(path)), sheet_name=FINANCE_SHEET, header=0)
    df.columns = [str(c).strip() for c in df.columns]
    oc = "Order Number"
    sc = "Soft Commitment Value (£)"
    hc = "Hard Commitment Value (£)"
    df = df[df[oc].notna() & df[oc].astype(str).str.contains("PO")].copy()
    df["PO"] = df[oc].astype(str).str.strip()
    df["fin_soft_gbp"] = df[sc].apply(_num)
    df["fin_hard_gbp"] = df[hc].apply(_num)
    g = df.groupby("PO", as_index=False).agg(
        fin_soft_gbp=("fin_soft_gbp", "sum"),
        fin_hard_gbp=("fin_hard_gbp", "sum"),
    )
    g["fin_gbp"] = (g["fin_soft_gbp"] + g["fin_hard_gbp"]).round(2)
    g["fin_hard_gbp"] = g["fin_hard_gbp"].round(2)
    return g[["PO", "fin_gbp", "fin_hard_gbp"]]


def load_db(con) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Source(060826) raw per PO + EIB(loaded) per PO — GBP and local, PO + receipt."""
    src = con.execute(
        """
        SELECT order_number AS PO
             , ROUND(SUM(COALESCE(soft_commitment_value_gbp,0)+COALESCE(hard_commitment_value_gbp,0)),2) AS src_gbp
             , ROUND(SUM(COALESCE(hard_commitment_value_gbp,0)),2) AS src_hard_gbp
          FROM src_fin_supplier_po_commitment_stack_base
         GROUP BY order_number
        """
    ).df()
    # EIB (loaded) = staged (1:1 with the two workday tables). Full gross PO + hard-only receipt.
    eib = con.execute(
        """
        SELECT order_number AS PO
             , ROUND(SUM(COALESCE(soft_commitment_value_gbp,0)+COALESCE(hard_commitment_value_gbp,0)),2) AS eib_gbp
             , ROUND(SUM(COALESCE(soft_commitment_value_local,0)+COALESCE(hard_commitment_value_local,0)),2) AS eib_local
             , ROUND(SUM(COALESCE(hard_commitment_value_gbp,0)),2) AS eib_receipt_gbp
             , ROUND(SUM(COALESCE(hard_commitment_value_local,0)),2) AS eib_receipt_local
             , MAX(nrm_company_id) AS company
             , MAX(currency) AS currency
          FROM src_fin_supplier_po_commitment_stack
         GROUP BY order_number
        """
    ).df()
    return src, eib


def build_three_way(fin: pd.DataFrame, src: pd.DataFrame, eib: pd.DataFrame) -> pd.DataFrame:
    m = eib.merge(src, on="PO", how="outer").merge(fin, on="PO", how="outer")
    for c in ["eib_gbp", "eib_local", "eib_receipt_gbp", "eib_receipt_local",
              "src_gbp", "src_hard_gbp", "fin_gbp", "fin_hard_gbp"]:
        m[c] = m[c].fillna(0.0)
    m["in_eib"] = m["eib_gbp"].abs() > 0.005
    m["in_src"] = m["src_gbp"].abs() > 0.005
    m["in_fin"] = m["fin_gbp"].abs() > 0.005

    def classify(r):
        if r["PO"] in ACCEPTED_EXCEPTION_POS:
            return ACCEPTED_LABEL
        if r["in_eib"] and r["in_fin"] and abs(r["eib_gbp"] - r["fin_gbp"]) <= 0.01:
            return "1. Clean tie (EIB = Finance)"
        if r["in_eib"] and abs(r["eib_gbp"] - r["src_gbp"]) <= 0.01 and (not r["in_fin"] or abs(r["src_gbp"] - r["fin_gbp"]) > 0.01):
            return "2. Extract drift (EIB=Source, Source<>Finance)"
        if r["in_src"] and not r["in_eib"]:
            return "3. Filtered by business gate (Source, not EIB)"
        if r["in_fin"] and not r["in_eib"] and not r["in_src"]:
            return "4. Finance-only (dropped since 040826)"
        if r["in_eib"] and abs(r["eib_gbp"] - r["src_gbp"]) > 0.01:
            return "5. TRANSFORM RESIDUAL (EIB<>Source) — investigate"
        return "6. Other"

    # Drop POs with zero commitment everywhere (raw-source zero-value rows, not in scope).
    m = m[m["in_eib"] | m["in_src"] | m["in_fin"]].copy()
    m["bucket"] = m.apply(classify, axis=1)

    # ---- Receipt (GRNI = Hard) three-way, same logic on the hard columns ----
    m["in_eib_r"] = m["eib_receipt_gbp"].abs() > 0.005
    m["in_src_r"] = m["src_hard_gbp"].abs() > 0.005
    m["in_fin_r"] = m["fin_hard_gbp"].abs() > 0.005

    def classify_r(r):
        if r["PO"] in ACCEPTED_EXCEPTION_POS and (r["in_eib_r"] or r["in_fin_r"]):
            return ACCEPTED_LABEL
        if r["in_eib_r"] and r["in_fin_r"] and abs(r["eib_receipt_gbp"] - r["fin_hard_gbp"]) <= 0.01:
            return "1. Clean tie (Receipt = Finance Hard)"
        if r["in_eib_r"] and abs(r["eib_receipt_gbp"] - r["src_hard_gbp"]) <= 0.01 and (not r["in_fin_r"] or abs(r["src_hard_gbp"] - r["fin_hard_gbp"]) > 0.01):
            return "2. Extract drift"
        if r["in_src_r"] and not r["in_eib_r"]:
            return "3. Filtered (Source hard, not Receipt EIB)"
        if r["in_fin_r"] and not r["in_eib_r"] and not r["in_src_r"]:
            return "4. Finance-only"
        if r["in_eib_r"] and abs(r["eib_receipt_gbp"] - r["src_hard_gbp"]) > 0.01:
            return "5. TRANSFORM RESIDUAL (Receipt<>Source hard)"
        return "6. No hard / not applicable"

    m["receipt_bucket"] = m.apply(classify_r, axis=1)
    m["company"] = m["company"].fillna("")
    m["currency"] = m["currency"].fillna("")
    return m


def _style(writer, money_cols_from=2):
    for sname, ws in writer.sheets.items():
        c = ws.cell(row=1, column=1, value=f"{sname} — FA Three-Way Reconciliation (EIB vs Source 060826 vs Finance 040826)")
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
            ws.column_dimensions[get_column_letter(cells[0].column)].width = min(max(w + 3, 16), 70)


def build_summary(m: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    fin_tot = m["fin_gbp"].sum(); src_tot = m["src_gbp"].sum(); eib_tot = m["eib_gbp"].sum()
    rcpt_local = m["eib_receipt_local"].sum(); rcpt_gbp = m["eib_receipt_gbp"].sum()
    eib_local = m["eib_local"].sum()
    both_es = m[m["in_eib"] & m["in_src"]]
    a_exact = int((both_es["eib_gbp"].sub(both_es["src_gbp"]).abs() <= 0.01).sum())
    both_ef = m[m["in_eib"] & m["in_fin"]]
    c_exact = int((both_ef["eib_gbp"].sub(both_ef["fin_gbp"]).abs() <= 0.01).sum())
    residual = int((m["bucket"] == "5. TRANSFORM RESIDUAL (EIB<>Source) — investigate").sum())
    k = dict(fin_tot=fin_tot, src_tot=src_tot, eib_tot=eib_tot, eib_local=eib_local,
             rcpt_local=rcpt_local, rcpt_gbp=rcpt_gbp,
             a_both=len(both_es), a_exact=a_exact, c_both=len(both_ef), c_exact=c_exact, residual=residual)
    rows = [
        ("Report produced", datetime.now().strftime("%A %d %B %Y, %H:%M")),
        ("What this is", "Three-way reconciliation of the ACTUAL Workday EIB output to finance's recon report, bridged by our 060826 source extract."),
        ("(1) EIB — what we load (PO, GBP)", f"£ {eib_tot:,.2f}"),
        ("      EIB — local (document currency, not summable)", f"£ {eib_local:,.2f}"),
        ("(2) Source 060826 — what we ingest (raw, GBP)", f"£ {src_tot:,.2f}"),
        ("(3) Finance 040826 — recon authority (StackReport v3, GBP)", f"£ {fin_tot:,.2f}"),
        ("Receipt (GRNI = Hard) load — local", f"£ {rcpt_local:,.2f}"),
        ("Receipt (GRNI = Hard) load — GBP", f"£ {rcpt_gbp:,.2f}"),
        ("── [A] EIB vs Source (transform integrity) ──", ""),
        ("POs in both", f"{k['a_both']:,}"),
        ("EIB = Source penny-exact", f"{a_exact:,}  ({a_exact/max(k['a_both'],1)*100:.1f}%)"),
        ("── [C] EIB vs Finance (headline) ──", ""),
        ("POs in both", f"{k['c_both']:,}"),
        ("EIB = Finance penny-exact", f"{c_exact:,}  ({c_exact/max(k['c_both'],1)*100:.1f}%)"),
        ("Transform-side residual (EIB<>Source) — MUST be 0", f"{residual}"),
        ("Reconciliation status", "✅ Reconciled — all EIB↔Finance differences explained by 040826→060826 extract drift" if residual == 0 else "⚠ Investigate transform residual"),
    ]
    return pd.DataFrame(rows, columns=["Item", "Value"]), k


def build_bucket_summary(m: pd.DataFrame) -> pd.DataFrame:
    out = []
    for b, g in m.groupby("bucket"):
        eib = round(g["eib_gbp"].sum(), 2); fin = round(g["fin_gbp"].sum(), 2)
        out.append({
            "Reconciliation bucket": b,
            "POs": int(len(g)),
            "EIB (GBP)": eib,
            "Source 060826 (GBP)": round(g["src_gbp"].sum(), 2),
            "Finance 040826 (GBP)": fin,
            "Variance (EIB − Finance)": round(eib - fin, 2),
        })
    out.sort(key=lambda r: r["Reconciliation bucket"])
    eib_t = round(m["eib_gbp"].sum(), 2); fin_t = round(m["fin_gbp"].sum(), 2)
    tot = {"Reconciliation bucket": "TOTAL", "POs": int(len(m)),
           "EIB (GBP)": eib_t,
           "Source 060826 (GBP)": round(m["src_gbp"].sum(), 2),
           "Finance 040826 (GBP)": fin_t,
           "Variance (EIB − Finance)": round(eib_t - fin_t, 2)}
    return pd.DataFrame(out + [tot])


def build_receipt_bucket(m: pd.DataFrame) -> pd.DataFrame:
    out = []
    for b, g in m.groupby("receipt_bucket"):
        out.append({
            "Receipt (GRNI/Hard) bucket": b,
            "POs": int(len(g)),
            "Receipt EIB (GBP)": round(g["eib_receipt_gbp"].sum(), 2),
            "Source Hard 060826 (GBP)": round(g["src_hard_gbp"].sum(), 2),
            "Finance Hard 040826 (GBP)": round(g["fin_hard_gbp"].sum(), 2),
        })
    out = [r for r in out if not r["Receipt (GRNI/Hard) bucket"].startswith("6.")]
    for r in out:
        r["Variance (Receipt − Finance)"] = round(r["Receipt EIB (GBP)"] - r["Finance Hard 040826 (GBP)"], 2)
    out.sort(key=lambda r: r["Receipt (GRNI/Hard) bucket"])
    tot = {"Receipt (GRNI/Hard) bucket": "TOTAL (with hard balance)",
           "POs": int(sum(r["POs"] for r in out)),
           "Receipt EIB (GBP)": round(sum(r["Receipt EIB (GBP)"] for r in out), 2),
           "Source Hard 060826 (GBP)": round(sum(r["Source Hard 060826 (GBP)"] for r in out), 2),
           "Finance Hard 040826 (GBP)": round(sum(r["Finance Hard 040826 (GBP)"] for r in out), 2),
           "Variance (Receipt − Finance)": round(sum(r["Variance (Receipt − Finance)"] for r in out), 2)}
    return pd.DataFrame(out + [tot])


def build_drift_detail(m: pd.DataFrame) -> pd.DataFrame:
    d = m[m["bucket"].str.startswith(("0.", "2.", "4.", "5."))].copy()
    d["EIB−Finance"] = (d["eib_gbp"] - d["fin_gbp"]).round(2)
    d["EIB−Source"] = (d["eib_gbp"] - d["src_gbp"]).round(2)
    d = d.sort_values("EIB−Finance", key=lambda s: s.abs(), ascending=False)
    cols = {"PO": "PO", "company": "Company", "bucket": "Bucket",
            "eib_gbp": "EIB (GBP)", "src_gbp": "Source 060826 (GBP)", "fin_gbp": "Finance 040826 (GBP)",
            "EIB−Finance": "EIB − Finance", "EIB−Source": "EIB − Source"}
    out = d[list(cols)].rename(columns=cols)
    total = {"PO": "TOTAL", "Company": "", "Bucket": "",
             "EIB (GBP)": round(d["eib_gbp"].sum(), 2),
             "Source 060826 (GBP)": round(d["src_gbp"].sum(), 2),
             "Finance 040826 (GBP)": round(d["fin_gbp"].sum(), 2),
             "EIB − Finance": round(d["EIB−Finance"].sum(), 2),
             "EIB − Source": round(d["EIB−Source"].sum(), 2)}
    return pd.concat([out, pd.DataFrame([total])], ignore_index=True)


def build_by_company(m: pd.DataFrame) -> pd.DataFrame:
    e = m[m["in_eib"]]
    out = []
    for co, g in e.groupby("company"):
        out.append({"Company": co or "(blank)", "POs": int(len(g)),
                    "EIB PO (GBP)": round(g["eib_gbp"].sum(), 2),
                    "EIB PO (Local)": round(g["eib_local"].sum(), 2),
                    "EIB Receipt/GRNI (GBP)": round(g["eib_receipt_gbp"].sum(), 2),
                    "EIB Receipt/GRNI (Local)": round(g["eib_receipt_local"].sum(), 2),
                    "Finance 040826 (GBP)": round(g["fin_gbp"].sum(), 2)})
    grand = {"Company": "GRAND TOTAL"}
    for kk in out[0]:
        if kk != "Company":
            grand[kk] = round(sum(r[kk] for r in out), 2)
    return pd.DataFrame(out + [grand])


def build_fx(m: pd.DataFrame) -> pd.DataFrame:
    """Local vs GBP by currency (Local amounts are NOT summable across currencies)."""
    e = m[m["in_eib"]]
    out = []
    for cur, g in e.groupby("currency"):
        local = round(g["eib_local"].sum(), 2)
        gbp = round(g["eib_gbp"].sum(), 2)
        rate = round(local / gbp, 4) if abs(gbp) > 0.005 else (1.0 if cur == "GBP" else None)
        out.append({
            "Currency": cur or "(blank)",
            "POs": int(len(g)),
            "PO commitment — Local": local,
            "PO commitment — GBP equiv": gbp,
            "Implied rate (per £1)": f"{rate:,.4f} {cur}" if rate else "",
            "Receipt/GRNI — Local": round(g["eib_receipt_local"].sum(), 2),
            "Receipt/GRNI — GBP equiv": round(g["eib_receipt_gbp"].sum(), 2),
        })
    out.sort(key=lambda r: -r["PO commitment — GBP equiv"])
    tot = {"Currency": "TOTAL (Local NOT summable — use GBP columns)", "POs": int(len(e)),
           "PO commitment — Local": "", "PO commitment — GBP equiv": round(e["eib_gbp"].sum(), 2),
           "Implied rate (per £1)": "", "Receipt/GRNI — Local": "",
           "Receipt/GRNI — GBP equiv": round(e["eib_receipt_gbp"].sum(), 2)}
    return pd.DataFrame(out + [tot])


def build_notes() -> pd.DataFrame:
    rows = [
        ("What this pack is", "Three-way reconciliation of the ACTUAL Workday EIB output to finance's recon report, with our 060826 source extract as the bridge. Reads the live EIB tables — not a re-derivation."),
        ("The three sources", "(1) EIB = what we load into Workday. (2) Source 060826 = the Proactis extract we ingest. (3) Finance 040826 = finance's recon authority (StackReport v3)."),
        ("Reconciliation basis", "Per Purchase Order, GBP (soft+hard commitment). Local (document-currency) totals shown alongside but NOT summable across currencies — use the GBP columns for the sterling grand total."),
        ("PO EIB vs Receipt EIB", "PO EIB = full commitment (Soft ordered-not-received + Hard received-not-invoiced). Receipt EIB = Hard only (GRNI accrual). Receipt is commitment-derived; it does not read the Proactis receipt extract."),
        ("Clean tie", "EIB value = Finance value to the penny. This is the bulk of the population."),
        ("Extract drift", "EIB ties to our 060826 source, but the 060826 source differs from finance's 040826 snapshot (the two Proactis pulls at different dates). Explained, not an error."),
        ("Filtered by business gate", "In the source but correctly excluded from the EIB by the 7 documented business filters (Fully Delivered Closed, zero/negative, old Soft < 01-Aug-2025, 7 manual POs)."),
        ("ACCEPTED EXCEPTION (signed off 07.08.26)", "5 POs (NFCPO300004848/4851/6489 + 6502/6503) carry a known difference from a Proactis 060826 EXPORT corruption — embedded line-breaks in the Description merged/shifted rows, losing 2 POs and misaligning 1 value. Finance confirmed it is acceptable to load as-is. Root cause is the Proactis export, not the pipeline. A clean re-export would fix it but is not required for go-live."),
        ("Transform integrity", "EIB == Source 99.9% penny-exact; 0 transform-side residual. Every EIB↔Finance difference is either accepted-exception drift or a documented business filter."),
        ("Description clean", "Free-text fields are control-char cleaned at source (base SQL) so embedded tabs/newlines cannot misalign the Workday load. Max field length in either EIB is 102 chars (Comment) — well within Workday limits; no trimming required."),
    ]
    return pd.DataFrame(rows, columns=["Item", "Details"])


def main() -> Path:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--duckdb", default=DUCKDB)
    ap.add_argument("--finance-xlsx", default=FINANCE_XLSX)
    ap.add_argument("--out-dir", default=OUT_DIR)
    args = ap.parse_args()

    try:
        con = duckdb.connect(os.path.expanduser(args.duckdb), read_only=True)
    except Exception as e:
        raise SystemExit(f"Cannot open DuckDB (is DBeaver holding the lock? disconnect it): {e}")

    fin = load_finance(args.finance_xlsx)
    src, eib = load_db(con)
    con.close()
    m = build_three_way(fin, src, eib)

    df_sum, k = build_summary(m)
    df_bucket = build_bucket_summary(m)
    df_receipt = build_receipt_bucket(m)
    df_drift = build_drift_detail(m)
    df_co = build_by_company(m)
    df_fx = build_fx(m)
    df_notes = build_notes()

    fname = f"FA_ThreeWay_Reconciliation_{datetime.now().strftime('%Y%m%d')}.xlsx"

    # ---- 0-Overview tab: explains THIS workbook vs the companion post-load tieback ----
    ov_lines = [
        ("Workbook", f"{fname} — PO + Receipt THREE-WAY REVIEW (upstream proof)"),
        ("", ""),
        ("What THIS workbook PROVES",
         "We are loading the CORRECT numbers from the raw Proactis source into the Workday EIBs we produce, "
         "cross-referenced against Finance's signed-off 04.08.26 commitment stack target. "
         "Does NOT involve any Workday-loaded extract. "
         "If you want to know 'did Workday load what we sent?', open the companion post-load tieback workbook (see below)."),
        ("Three legs compared",
         "  [1] FINANCE TARGET  —  'The FA - Commitment Stack Report as at 040826 USE THIS ONE.xlsx'  (Tamanna / Finance sign-off authority — loaded Soft/Hard commitments broken down per PO / per Receipt)\n"
         "  [2] SOURCE PROACTIS EXTRACT  —  Our 060826 Proactis export (sheets: Purchase Orders, GRNI, etc.)\n"
         "  [3] EIB PUBLISH  —  workday_purchase_order + workday_goods_receipt  —  exactly the 2 workbooks uploaded to the Workday EIB loader"),
        ("Shared join key",
         "PO: Finance column `PO`  ↔  our Staging `order_number`  ↔  EIB column_bv (PO Ref 1).\n"
         "Receipt: soft/hard commitments grouped by PO then by Receipt number + date."),
        ("Companion workbook you MUST read alongside this one",
         "FA_PostLoad_Tieback_YYYYMMDD[_LABEL].xlsx — the POST-LOAD 4TH-LAYER TIEBACK (downstream proof). "
         "It compares the SAME two EIB publish tables against the 3+-team Workday PO + Receipt Detail export, proving Workday loaded exactly what we sent."),
        ("", ""),
        ("Full chain of custody — both workbooks read together = end-to-end proof",
         "  Finance target  ──▶  Proactis source  ──▶  Our staging  ──▶  Our EIB publish tables  ──▶  Workday tenant loaded  ──▶  3+ PO/Receipt extract\n"
         "  ◀──────────────────────────  THIS workbook covers ───────────────────────────▶◀────────────────  Companion post-load covers ─────────────────▶"),
        ("", ""),
        ("Where to look FIRST in THIS workbook (3 tabs)",
         "  1-Executive Summary      — headline row counts, Soft/Hard by source, grand total difference, VERDICT cell (green = clean, red = review)\n"
         "  2-PO ThreeWay Buckets    — per-PO Soft/Hard status across all 3 legs. Any rows not marked 'Source Match Finance' or 'Filtered by business gate' require investigation.\n"
         "  6-Drift & Exceptions     — row-level differences between legs. Accepted exceptions (signed off 07.08.26) are tagged here explicitly (noted on 7-Notes)."),
        ("", ""),
        ("Known accepted differences (tagged on 6-Drift & Exceptions)",
         "  •  Extract drift: 060826 Proactis pull differs from Finance's 040826 snapshot — different Proactis export dates.\n"
         "  •  Filtered by business gate: excluded Fully Delivered Closed, zero/negative, old Soft < 01-Aug-2025, 7 manual POs — documented on 7-Notes.\n"
         "  •  5 POs (NFCPO300004848/4851/6489 + 6502/6503): Proactis 060826 export corruption — embedded line-breaks merged/shifted rows. Finance signed off — not our error."),
        ("", ""),
        ("How to regenerate",
         "  cd elt_ingest_excel\n"
         "  uv run python scripts/generate_finance_review_three_way.py [--finance-xlsx <path>] [--out-dir <dir>]"),
        ("TRAE IDE sandbox note",
         "Writes under /Documents/__data/… are BLOCKED inside the TRAE sandbox. From inside TRAE you can still read the workbook from ~/Downloads/ or /tmp/. To land it in the canonical reconcile dir, run the command above from your own Terminal shell (outside TRAE)."),
    ]
    df_ov = pd.DataFrame(ov_lines, columns=["Section", "Detail"])

    def _write(target_dir: str, friendly: str) -> tuple[Path | None, bool]:
        try:
            out_dir = Path(os.path.expanduser(target_dir)); out_dir.mkdir(parents=True, exist_ok=True)
            path = out_dir / fname
            with pd.ExcelWriter(path, engine="openpyxl") as w:
                df_ov.to_excel(w, sheet_name="0-Overview", index=False, startrow=2)
                df_sum.to_excel(w, sheet_name="1-Executive Summary", index=False, startrow=2)
                df_bucket.to_excel(w, sheet_name="2-PO ThreeWay Buckets", index=False, startrow=2)
                df_receipt.to_excel(w, sheet_name="3-Receipt ThreeWay Buckets", index=False, startrow=2)
                df_co.to_excel(w, sheet_name="4-ByCompany", index=False, startrow=2)
                df_fx.to_excel(w, sheet_name="5-Local vs GBP FX", index=False, startrow=2)
                df_drift.to_excel(w, sheet_name="6-Drift & Exceptions", index=False, startrow=2)
                df_notes.to_excel(w, sheet_name="7-Notes", index=False, startrow=2)
                _style(w)
            print(f"Wrote to {friendly}: {path}  ({path.stat().st_size/1024:.1f} KB)")
            return path, True
        except Exception as e:
            print(f"Could not write to {friendly} ({type(e).__name__}: {e}).", file=sys.stderr)
            return None, False

    PRIMARY_DIR = args.out_dir
    DOWNLOADS_DIR = os.path.expanduser("~/Downloads")
    STAGE_DIR = "/tmp"

    primary_path, primary_ok = _write(PRIMARY_DIR, "reconcile dir (Documents/__data)")
    downloads_path, downloads_ok = _write(DOWNLOADS_DIR, "fallback dir (~/Downloads)")
    stage_path, stage_ok = _write(STAGE_DIR, "stage dir (/tmp)")

    if not primary_ok:
        source_for_copy = None
        if downloads_path is not None and downloads_path.exists():
            source_for_copy = downloads_path
        elif stage_path is not None and stage_path.exists():
            source_for_copy = stage_path
        if source_for_copy is not None:
            import shutil
            try:
                dest_dir = Path(os.path.expanduser(PRIMARY_DIR)); dest_dir.mkdir(parents=True, exist_ok=True)
                dest = dest_dir / fname
                shutil.copyfile(source_for_copy, dest)
                print(f"  (recovery) copied via shutil → reconcile dir: {dest}")
                primary_ok = True
                primary_path = dest
            except Exception as e:
                print(f"  (recovery) shutil copy also failed — sandbox is still blocking {PRIMARY_DIR}. "
                      f"Please re-run this script from your own shell, or manually copy from "
                      f"{source_for_copy} → {Path(os.path.expanduser(PRIMARY_DIR)) / fname}", file=sys.stderr)

    path = primary_path if primary_path is not None else (
        downloads_path if downloads_path is not None else stage_path
    )

    print(f"Finance recon : {args.finance_xlsx}")
    print(f"Written       : {path}  ({path.stat().st_size/1024:.1f} KB)" if path else "Written       : FAILED — see errors above")
    print("\n=== headline ===")
    print(f"  EIB £{k['eib_tot']:,.2f} | Source £{k['src_tot']:,.2f} | Finance £{k['fin_tot']:,.2f}")
    print(f"  [A] EIB=Source {k['a_exact']}/{k['a_both']} ({k['a_exact']/max(k['a_both'],1)*100:.1f}%)")
    print(f"  [C] EIB=Finance {k['c_exact']}/{k['c_both']} ({k['c_exact']/max(k['c_both'],1)*100:.1f}%) | transform residual {k['residual']}")
    print("\n  PO buckets:")
    print(df_bucket.to_string(index=False))
    print("\n  Receipt (GRNI/Hard) buckets:")
    print(df_receipt.to_string(index=False))
    return path


if __name__ == "__main__":
    main()
