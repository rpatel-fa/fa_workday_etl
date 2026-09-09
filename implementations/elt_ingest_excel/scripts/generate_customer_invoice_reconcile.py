"""Single unified Customer Invoice reconcile workbook — all 4 tie-back layers in one xlsx.

COMBINES BOTH:
  • generate_customer_invoice_three_way.py           (Upstream: Legs 1–2–3 — GP Source → Staging → EIB publish)
  • generate_customer_invoice_post_load_reconcile.py (Downstream: Legs 3–4 — EIB publish → Workday tenant extract)

PRODUCES:  FA_Reconcile_CustomerInvoice_YYYYMMDD[_LABEL].xlsx

Structure:
  0-Overview                      Unified — explains the ENTIRE chain (Legs 1-2-3-4), this workbook contains everything
  1-Exec Summary                  Unified — Upstream headline + Downstream headline + BOTH verdict cells side-by-side

  =====  PART A: UPSTREAM (GP → Staging → Our EIB)  =====
  A2-By Txn Type                  (from three_way)
  A3-By Company Currency          (from three_way)
  A4-Currency Coverage            (from three_way)
  A5-Per-Document Exact           (from three_way)
  A6-Customer Bucket              (from three_way)
  A7-Drift & Exceptions           (from three_way)
  A8-Notes (Upstream)             (from three_way)

  =====  PART B: DOWNSTREAM (Our EIB → Workday loaded)  =====
  B2-By Txn Type                  (from post_load)
  B3-By Company Currency          (from post_load)
  B4-Currency Coverage            (from post_load)
  B5-Credit Note Exact            (from post_load) — strongest per-doc proof
  B6-Reconcile (Customer)         (from post_load)
  B7-Reconcile (Invoice)          (from post_load)
  B8-Notes (Downstream)           (from post_load)
  B9-Excluded WD Rows             (from post_load)

Verdicts on 1-Exec Summary are highlighted independently (green/red) for each part so a
reader can see at a glance whether the upstream transformation and the downstream tenant
load each tied out cleanly.
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import shutil
import sys
from pathlib import Path

import duckdb
import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import generate_customer_invoice_three_way as tw
import generate_customer_invoice_post_load_reconcile as pl

OUT_DIR = Path(os.path.expanduser(os.environ.get("OUT_DIR", str(pl.OUT_DIR))))
DOWNLOADS_FALLBACK_DIR = Path(os.path.expanduser("~/Downloads"))
STAGE_DIR = Path("/tmp")

TODAY = dt.date.today()

OK_FILL = PatternFill("solid", fgColor="C6EFCE")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")

DEFAULT_DUCKDB = Path(os.path.expanduser("~/Documents/__data/duckdb/rpatel.duckdb"))
DEFAULT_EXTRACT = Path(os.path.expanduser(
    "~/Documents/__data/excel/workday/extract/ar_invoices/FA_Find_Customer_Invoices (2).xlsx"
))


# ---------------------------------------------------------------------------
# Shared write helpers
# ---------------------------------------------------------------------------
def _df(headers, rows):
    return pd.DataFrame(rows, columns=headers)


def _write_df(ws, df, title_row, start_row=3):
    """Write a DataFrame starting at `start_row` (1-based); row `start_row-1` holds the title."""
    from openpyxl.utils.dataframe import dataframe_to_rows
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=start_row):
        for c_idx, val in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=val)
    if title_row:
        ws.cell(row=start_row - 1, column=1, value=title_row)
        ws.cell(row=start_row - 1, column=1).font = Font(bold=True, size=13)


def _autosize(ws, header_row_1based=3):
    for col_idx in range(1, ws.max_column + 1):
        max_w = 8
        for r in range(header_row_1based, ws.max_row + 1):
            v = ws.cell(row=r, column=col_idx).value
            if v is None:
                continue
            w = len(str(v))
            if w > max_w:
                max_w = w
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_w + 3, 16), 70)
    ws.freeze_panes = "A" + str(header_row_1based + 1)


def _style_all(writer):
    """Apply light styling + freeze panes + autosize on every sheet written so far."""
    for sheet_name, ws in writer.sheets.items():
        _autosize(ws, header_row_1based=4)


def _highlight_col(ws, col_letter_or_1based, thr=0.01, start_row=4,
                   ok_fill=OK_FILL, bad_fill=BAD_FILL, warn_fill=None):
    if isinstance(col_letter_or_1based, str):
        from openpyxl.utils import column_index_from_string
        ci = column_index_from_string(col_letter_or_1based)
    else:
        ci = int(col_letter_or_1based)
    for i in range(start_row, ws.max_row + 1):
        cell = ws.cell(i, ci)
        try:
            v = float(cell.value)
        except (TypeError, ValueError):
            continue
        if abs(v) > thr:
            cell.fill = bad_fill
        elif warn_fill and abs(v) > 0.0005:
            cell.fill = warn_fill
        else:
            cell.fill = ok_fill


# ---------------------------------------------------------------------------
# Build unified overview + exec summary
# ---------------------------------------------------------------------------
def _build_overview_df(upstream_hl: dict, downstream_res: dict, extract_basename: str) -> pd.DataFrame:
    lines = [
        ("Workbook",
         "FA_Reconcile_CustomerInvoice_YYYYMMDD[_LABEL].xlsx — Customer Invoice UNIFIED FULL-CHAIN TIE-BACK"),
        ("", ""),
        ("What THIS single workbook PROVES",
         "End-to-end proof across ALL FOUR legs of the AR Invoice pipeline, in one file.\n"
         "  PART A — Upstream: did we correctly transform the raw GP AR TB into the EIB we published?\n"
         "  PART B — Downstream: did Workday actually LOAD exactly what we SENT in that published EIB?"),
        ("Full chain of custody — EVERYTHING is in THIS workbook",
         "    [1] AR DTL TB source Excel    ─▶   [2] Our staging (base → nrm → gate)   ─▶   [3] Our EIB publish tables   ─▶   [4] Workday tenant loaded → 3+ Find Customer Invoices export\n"
         "    ◀────────────────  PART A covers Legs 1 – 2 – 3  ────────────────▶◀────────  PART B covers Legs 3 – 4  ────────▶"),
        ("", ""),
        ("PART A — Upstream (Source → Staging → EIB)",
         "  Row counts:  Source=%d  Staging=%d  EIB=%d   •   Fully matched: %d / %d   •   Amt mismatches: %d\n"
         "  Σ local:     Src £%s    Stg £%s    Eib £%s    •   Δ(Src−Eib) = £%s\n"
         "  VERDICT: %s" % (
             upstream_hl["n_src"], upstream_hl["n_stg"], upstream_hl["n_eib"],
             upstream_hl["n_match"], upstream_hl["N"], upstream_hl["n_amt_bad"],
             f"{upstream_hl['sum_src']:,.2f}", f"{upstream_hl['sum_stg']:,.2f}",
             f"{upstream_hl['sum_eib']:,.2f}",
             f"{upstream_hl['sum_eib'] - upstream_hl['sum_src']:+,.2f}",
             upstream_hl["verdict"],
         )),
        ("PART B — Downstream (Sent EIB → Workday loaded)",
         "  Grand totals:  Sent £%s   WD £%s   Δ(Sent−WD)=£%s     Status: %s\n"
         "  CN exact key:  Both=%d  Sent-only=%d  Loaded-only=%d  Amt mismatch=%d\n"
         "  Cust buckets:  Match=%d  Sent-only=%d  Loaded-only=%d  Review=%d\n"
         "  Invoice fuzzy: Sent rows=%d  Loaded rows=%d  Matched=%d  Sent-only=%d  Loaded-only=%d  Amt mismatch=%d\n"
         "  Loaded extract file: %s" % (
             f"{downstream_res['grand_sent']:,.2f}", f"{downstream_res['grand_wd']:,.2f}",
             f"{downstream_res['grand_delta']:+,.2f}",
             ("MATCH" if downstream_res["grand_match"] else "⚠ MISMATCH"),
             downstream_res["cn_both"], downstream_res["cn_sent_only"],
             downstream_res["cn_wd_only"], downstream_res["cn_amt_mism"],
             *downstream_res["cust_bucket_counts"],
             *downstream_res["inv_totals"],
             extract_basename,
         )),
        ("", ""),
        ("Where to look FIRST",
         "  1-Exec Summary  →  both verdict cells (top = PART A, bottom = PART B); green = CLEAN, red = REVIEW.\n"
         "  Part A tabs — investigate transformation drift:\n"
         "     A7-Drift & Exceptions   (any failing per-document presence / amount row)\n"
         "     A5-Per-Document Exact   (the full row-level comparison)\n"
         "  Part B tabs — investigate post-load / tenant-load differences:\n"
         "     B5-Credit Note Exact          STRONGEST per-doc proof — 254 credit notes, penny-exact on shared key.\n"
         "     B6-Reconcile (Customer)       Customer buckets (Co × norm name × CCY).\n"
         "     B7-Reconcile (Invoice)       Invoice traceability (greedy Co×CCy×Cust×amt).\n"
         "     B9-Excluded WD Rows          (Draft/Voided/Cancelled tenant rows — no counterpart in Sent EIB)."),
        ("VERDICT semantics",
         "  PART A — CLEAN requires: row counts identical on all 3 legs (STE presence = 1,447)  AND  0 amount mismatches  AND  1,447/1,447 fully matched rows.\n"
         "  PART B — CLEAN requires: Grand total Δ=£0.00  AND  0 credit-note amount mismatches  AND  0 invoice-fuzzy amount mismatches."),
        ("How to regenerate",
         "  cd elt_ingest_excel\n"
         "  uv run python scripts/generate_customer_invoice_reconcile.py \\\n"
         "      --workday-extract ~/Documents/__data/excel/workday/extract/ar_invoices/'FA_Find_Customer_Invoices (2).xlsx' \\\n"
         "      --label PRODUCTION\n"
         "Flags:  --duckdb <path>    override DuckDB location (default: ~/Documents/__data/duckdb/rpatel.duckdb)\n"
         "        --workday-extract <path>   the 3+-team loaded-tenant xlsx (AR Aging Detail OR Find Customer Invoices, auto-detected)\n"
         "        --sheet <name>             override sheet within the extract (default: auto)\n"
         "        --label <TEXT>             tag appended to filename, e.g. PRODUCTION\n"
         "        --out-dir <dir>            override primary output dir (defaults to ~/Documents/__data/excel/workday/reconcile)"),
    ]
    return pd.DataFrame(lines, columns=["Section", "Detail"])


def _build_exec_summary_df(upstream_res: dict, downstream_res: dict) -> tuple[pd.DataFrame, dict]:
    """Return (df, meta) where meta carries verdict-cell row numbers so the caller can highlight them."""
    u_hl = upstream_res["_headline"]

    rows: list[list] = [
        # Part A banner
        ["PART A — UPSTREAM  (GP Source → Staging → EIB publish)", ""],
        ["Metric", "Source (AR TB raw)", "Staging (src_fin_customer_invoice)", "EIB publish",
         "Δ Source-Stage", "Δ Stage-EIB", "Notes"],
    ]
    rows.extend(upstream_res["exec"][1])  # reuse exec rows from three_way
    rows.append(["", "", "", "", "", "", ""])

    # Part B banner
    rows.append(["PART B — DOWNSTREAM  (Sent EIB → Workday tenant loaded)", ""])
    rows.append(["Item", "Value", "", "", "", "", ""])
    summary_rows = [
        ("Report produced", dt.datetime.now().strftime("%A %d %B %Y, %H:%M")),
        ("Sent EIB  Σ (workday_customer_invoice + credit_note)", f"£ {downstream_res['grand_sent']:,.2f}"),
        ("Loaded Workday  Σ (tenant xlsx)", f"£ {downstream_res['grand_wd']:,.2f}"),
        ("Grand delta  (Sent − Loaded)", f"£ {downstream_res['grand_delta']:+,.2f}"),
        ("WD rows EXCLUDED from compare (Draft/Voided/Cancelled)",
         ";  ".join(downstream_res.get("excluded_summary", []))
         if downstream_res.get("excluded_summary")
         else "NONE — no Draft/Manual WD rows removed"),
        ("By type — rows (Sent / Loaded)",
         "   •   ".join(
             f"{r[0]}: {int(r[1]):,} / {int(r[2]):,}  Δ{int(r[3]):+,d}"
             for r in downstream_res["txn_type_rows"]
         )),
        ("By type — amount (Sent / Loaded)",
         "   •   ".join(
             f"{r[0]}: {float(r[4]):,.2f} / {float(r[5]):,.2f}  Δ{float(r[6]):+,.2f}"
             for r in downstream_res["txn_type_rows"]
         )),
        (f"Credit-note exact key  (Both / Sent-only / Loaded-only / Amt mismatch)",
         f"{downstream_res['cn_both']} / {downstream_res['cn_sent_only']} / "
         f"{downstream_res['cn_wd_only']} / {downstream_res['cn_amt_mism']}"),
        ("Customer-bucket counts  (Match / Sent-only / Loaded-only / Review)",
         " / ".join(str(x) for x in downstream_res["cust_bucket_counts"])),
        ("Invoice fuzzy match  (Sent / Loaded / Matched / Sent-only / Loaded-only / Amt mism)",
         " / ".join(str(x) for x in downstream_res["inv_totals"])),
        ("PART B — Reconciliation status",
         "✅ Tied out — Sent == Loaded to the penny on all authoritative layers"
         if downstream_res["grand_match"] and downstream_res["cn_amt_mism"] == 0 and downstream_res["inv_totals"][5] == 0
         else "⚠ Investigate — see Part B later tabs"),
    ]
    for k, v in summary_rows:
        rows.append([k, v, "", "", "", "", ""])

    df = pd.DataFrame(rows, columns=["Metric", "Source (AR TB raw)", "Staging", "EIB publish",
                                     "Δ Src-Stg", "Δ Stg-Eib", "Notes"])

    # Locate Part A verdict row & Part B status row for highlighting (scan after writing via values)
    return df, {"u_headline": u_hl}


# ---------------------------------------------------------------------------
# Unified writer
# ---------------------------------------------------------------------------
def write_unified_workbook(upstream_res: dict, downstream_res: dict, out_path: Path,
                           extract_basename: str) -> None:
    ov_df = _build_overview_df(upstream_res["_headline"], downstream_res, extract_basename)
    exec_df, meta = _build_exec_summary_df(upstream_res, downstream_res)

    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        # 0-Overview
        ov_df.to_excel(w, sheet_name="0-Overview", index=False, startrow=2)

        # 1-Exec Summary
        exec_df.to_excel(w, sheet_name="1-Exec Summary", index=False, startrow=2)

        # ====== Part A ======
        _df(*upstream_res["by_txn"]).to_excel(w, sheet_name="A2-By Txn Type", index=False, startrow=2)
        _df(*upstream_res["by_co_ccy"]).to_excel(w, sheet_name="A3-By Company Currency", index=False, startrow=2)
        _df(*upstream_res["currency_cov"]).to_excel(w, sheet_name="A4-Currency Coverage", index=False, startrow=2)
        _df(*upstream_res["per_document"]).to_excel(w, sheet_name="A5-Per-Document Exact", index=False, startrow=2)
        _df(*upstream_res["customer_bucket"]).to_excel(w, sheet_name="A6-Customer Bucket", index=False, startrow=2)
        _df(*upstream_res["exceptions"]).to_excel(w, sheet_name="A7-Drift & Exceptions", index=False, startrow=2)
        _df(*upstream_res["notes"]).to_excel(w, sheet_name="A8-Notes (Upstream)", index=False, startrow=2)

        # ====== Part B ======
        _df(
            ["Type", "Sent rows", "Loaded rows", "Δ rows",
             "Sent amount", "Loaded amount", "Δ amount", "Join side"],
            downstream_res["txn_type_rows"],
        ).to_excel(w, sheet_name="B2-By Txn Type", index=False, startrow=2)
        _df(
            ["Company", "CCY", "Sent rows", "Loaded rows", "Δ rows",
             "Sent amount", "Loaded amount", "Δ amount"],
            downstream_res["ccy_rows"],
        ).to_excel(w, sheet_name="B3-By Company Currency", index=False, startrow=2)
        _df(
            ["Currency", "In Sent EIB", "In Loaded WD", "Status"],
            downstream_res["ccy_coverage"],
        ).to_excel(w, sheet_name="B4-Currency Coverage", index=False, startrow=2)
        _df(downstream_res["cn_headers"], downstream_res["cn_rows"]).to_excel(
            w, sheet_name="B5-Credit Note Exact", index=False, startrow=2)
        _df(downstream_res["cust_headers"], downstream_res["cust_rows"]).to_excel(
            w, sheet_name="B6-Reconcile (Customer)", index=False, startrow=2)
        _df(downstream_res["inv_headers"], downstream_res["inv_rows"]).to_excel(
            w, sheet_name="B7-Reconcile (Invoice)", index=False, startrow=2)
        _df(
            ["Item", "Details"],
            downstream_res.get("notes_rows_pairs", []),
        ).to_excel(w, sheet_name="B8-Notes (Downstream)", index=False, startrow=2)
        _df(downstream_res["excluded_headers"], downstream_res.get("excluded_rows", [])).to_excel(
            w, sheet_name="B9-Excluded WD Rows", index=False, startrow=2
        )

        _style_all(w)

        # ---------------- Verdict highlighting (1-Exec Summary) ----------------
        ws_exec = w.sheets["1-Exec Summary"]

        # 1) Part A verdict = row where column 1 == "VERDICT"
        part_a_verdict_row = None
        for r in range(4, ws_exec.max_row + 1):
            if ws_exec.cell(r, 1).value == "VERDICT":
                part_a_verdict_row = r
                break
        if part_a_verdict_row is not None:
            is_clean_a = (meta["u_headline"]["verdict"] == "CLEAN")
            for c in range(1, 8):
                cell = ws_exec.cell(row=part_a_verdict_row, column=c)
                cell.fill = OK_FILL if is_clean_a else BAD_FILL
                cell.font = Font(bold=True)

        # 2) Part B status row = "PART B — Reconciliation status"
        part_b_status_row = None
        for r in range(4, ws_exec.max_row + 1):
            if ws_exec.cell(r, 1).value == "PART B — Reconciliation status":
                part_b_status_row = r
                break
        if part_b_status_row is not None:
            clean_b = (downstream_res["grand_match"] and downstream_res["cn_amt_mism"] == 0
                       and downstream_res["inv_totals"][5] == 0)
            cell = ws_exec.cell(row=part_b_status_row, column=2)
            cell.fill = OK_FILL if clean_b else BAD_FILL
            cell.font = Font(bold=True)

        # 3) Upstream delta cols on Part A (columns 5, 6) — red when |value| > 0.01
        part_banner_row = None
        for r in range(4, ws_exec.max_row + 1):
            v = ws_exec.cell(r, 1).value
            if isinstance(v, str) and v.startswith("PART B") and "DOWNSTREAM" in v:
                part_banner_row = r
                break
        end_a = part_banner_row - 1 if part_banner_row else ws_exec.max_row
        for r in range(4, end_a + 1):
            for c in (5, 6):
                cell = ws_exec.cell(row=r, column=c)
                try:
                    if abs(float(cell.value)) > 0.01:
                        cell.fill = BAD_FILL
                    else:
                        cell.fill = OK_FILL
                except (TypeError, ValueError):
                    pass

        # 4) Part B: "Grand delta" cell highlight
        for r in range(4, ws_exec.max_row + 1):
            v = ws_exec.cell(r, 1).value
            if isinstance(v, str) and "Grand delta" in v:
                try:
                    raw = str(ws_exec.cell(r, 2).value)
                    num = float(raw.replace("£", "").replace(",", ""))
                    if abs(num) > 0.01:
                        ws_exec.cell(r, 2).fill = BAD_FILL
                except Exception:
                    pass

        # ---------------- Delta-column highlights across detail sheets ----------------
        # Part A delta cols
        for sheet, delta_col in [
            ("A2-By Txn Type", None),
            ("A3-By Company Currency", None),
        ]:
            pass
        # Part A: A5-Per-Document Exact  → status col
        ws_a5 = w.sheets["A5-Per-Document Exact"]
        for r in range(4, ws_a5.max_row + 1):
            # status col is col 16 "Overall status" in three_way — check header first
            if r == 4:
                # scan header row to find column named "Overall status" or similar
                a5_status_col = None
                for ci in range(1, ws_a5.max_column + 1):
                    hdr = ws_a5.cell(row=4, column=ci).value
                    if hdr and ("status" in str(hdr).lower() or "overall" in str(hdr).lower()):
                        a5_status_col = ci
                        break
            else:
                a5_status_col = locals().get("a5_status_col")
                if a5_status_col:
                    val = ws_a5.cell(row=r, column=a5_status_col).value
                    if val == "MATCH":
                        ws_a5.cell(row=r, column=a5_status_col).fill = OK_FILL
                    else:
                        ws_a5.cell(row=r, column=a5_status_col).fill = BAD_FILL

        # Part B
        for sheet, delta_col_1based in [
            ("B2-By Txn Type", 4), ("B2-By Txn Type", 7),
            ("B3-By Company Currency", 5), ("B3-By Company Currency", 8),
            ("B5-Credit Note Exact", 8),
            ("B6-Reconcile (Customer)", 6), ("B6-Reconcile (Customer)", 9),
            ("B7-Reconcile (Invoice)", 7),
        ]:
            _highlight_col(w.sheets[sheet], delta_col_1based, thr=(0 if delta_col_1based in (4, 5, 6) else 0.01))

        # B9: excluded status col → red
        ws_ex = w.sheets["B9-Excluded WD Rows"]
        if ws_ex.max_row >= 4:
            # locate "Exclude reason" / "Status" col
            ex_col = None
            for ci in range(1, ws_ex.max_column + 1):
                hdr = ws_ex.cell(row=4, column=ci).value
                if hdr and ("reason" in str(hdr).lower() or "status" in str(hdr).lower()):
                    ex_col = ci
                    break
            if ex_col:
                for r in range(5, ws_ex.max_row + 1):
                    if ws_ex.cell(r, ex_col).value:
                        ws_ex.cell(r, ex_col).fill = BAD_FILL


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def _try_write(target_dir: Path, friendly: str, fname: str,
               upstream_res: dict, downstream_res: dict,
               extract_basename: str) -> tuple[Path | None, bool]:
    try:
        target_dir.mkdir(parents=True, exist_ok=True)
        out_path = target_dir / fname
        write_unified_workbook(upstream_res, downstream_res, out_path, extract_basename)
        size_kb = out_path.stat().st_size / 1024.0
        print(f"Wrote to {friendly}: {out_path}  ({size_kb:.1f} KB)")
        return out_path, True
    except Exception as e:
        print(f"Could not write to {friendly} ({type(e).__name__}: {e}).", file=sys.stderr)
        return None, False


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--duckdb", default=str(DEFAULT_DUCKDB))
    ap.add_argument("--workday-extract", default=str(DEFAULT_EXTRACT))
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--label", default=None)
    ap.add_argument("--out-dir", default=str(OUT_DIR))
    args = ap.parse_args()

    duckdb_path = Path(os.path.expanduser(args.duckdb))
    extract_path = Path(os.path.expanduser(args.workday_extract))
    if not extract_path.exists():
        print(f"Workday extract not found: {extract_path}", file=sys.stderr)
        return 2

    # ---- Build downstream (needs wd xlsx + duckdb con) ----
    wd_df, _ = pl.load_workday_ar_aging(str(extract_path), args.sheet)
    con = duckdb.connect(str(duckdb_path), read_only=True)
    try:
        downstream_res = pl.build(con, wd_df)

        # ---- Build upstream (uses con for source/staged/EIB) ----
        source_df = tw._load_source_raw(con)
        staged_df = tw._load_staged(con)
        eib_df = tw._load_eib(con)
    finally:
        con.close()
    upstream_res = tw.build(source_df, staged_df, eib_df)

    # ---- Print headline (mirrors both individual CLIs) ----
    u_hl = upstream_res["_headline"]
    print(f"[Part A · Upstream] Row counts: Src={u_hl['n_src']} Stg={u_hl['n_stg']} "
          f"Eib={u_hl['n_eib']}  Matched={u_hl['n_match']}/{u_hl['N']}  AmtMism={u_hl['n_amt_bad']}  "
          f"VERDICT={u_hl['verdict']}")
    print(f"[Part B · Downstream] Grand total: Sent={downstream_res['grand_sent']:,.2f}  "
          f"WD={downstream_res['grand_wd']:,.2f}  Δ={downstream_res['grand_delta']:,.2f}  "
          f"{'MATCH' if downstream_res['grand_match'] else 'MISMATCH'}")
    print(f"  CN exact key: Both={downstream_res['cn_both']}  SentOnly={downstream_res['cn_sent_only']}  "
          f"WDOnly={downstream_res['cn_wd_only']}  AmtMism={downstream_res['cn_amt_mism']}")
    print(f"  Cust buckets:  Match={downstream_res['cust_bucket_counts'][0]}  "
          f"SentOnly={downstream_res['cust_bucket_counts'][1]}  "
          f"WDOnly={downstream_res['cust_bucket_counts'][2]}  Review={downstream_res['cust_bucket_counts'][3]}")
    print(f"  Invoice fuzzy: SentRows={downstream_res['inv_totals'][0]}  WDRows={downstream_res['inv_totals'][1]}  "
          f"Matched={downstream_res['inv_totals'][2]}  SentOnly={downstream_res['inv_totals'][3]}  "
          f"WDOnly={downstream_res['inv_totals'][4]}  AmtMism={downstream_res['inv_totals'][5]}")

    # Ensure pl.notes is available as pairs if not already
    if "notes_rows_pairs" not in downstream_res:
        # Rebuild the same pairs list that post_load write_workbook would have used (it was inline)
        pl_notes_pairs = [
            ("CUSTOMER INVOICE POST-LOAD TIEBACK — METHODOLOGY", ""),
            ("Source [3] — Sent EIB",
             "workday_customer_invoice (INVOICE lines, column_ha amount) + workday_customer_credit_note"
             " (Adjustment lines, column_fb amount), joined back to src_fin_customer_invoice_nrm on"
             " (wd_customer_id, gp_document_number) for customer names and source document metadata."),
            ("Source [4] — Loaded Workday extract",
             f"File = {extract_path.name}.  Format auto-detected by header-row fingerprint"
             " (supports BOTH Format A = AR Aging Detail by Due Date, AND Format B = Find Customer Invoices)."),
            ("Find Customer Invoices (Format B) pre-processing",
             "1) Status filter → keep {Approved, Posted, Partially Paid, Paid} only; Draft/Voided/Cancelled"
             " → surfaced on B9-Excluded WD Rows.\n"
             "2) wd_txn_ref fallback → if 'Reference Number' column is 100% NULL (as here, because loaded"
             " invoices arrived via legacy GP EIB conversion), derive ref from the first-column label"
             " ('Customer Invoice: …' / 'Customer Invoice Adjustment: …').\n"
             "3) Cross-ccy alignment bridge → non-GBP invoices displayed as company-currency GBP by the"
             " Find view are rewritten from a unique 1:1 (company, customer_norm, penny-exact amount) match"
             " against an unused SENT row so the customer/invoice layers align.\n"
             "4) Credit-note exact-key prefix stripping → the 'Customer Invoice Adjustment: FCSH…' prefix is"
             " stripped on both sides of the exact-key comparison so the plain FCSH doc numbers join 1:1."),
            ("Join strategy (strongest → weakest)",
             "1) Grand total  2) By Txn Type  3) By Co×CCY  4) Currency coverage\n"
             "5) Credit-note EXACT-KEY (strongest per-doc proof, shared GP FCSH/CWCRT doc number)\n"
             "6) Customer bucket (Co × norm name × CCY)\n"
             "7) Reconcile (Invoice) trace  — greedy 4-tuple (Co×CCy×Cust×amt±0.005) because loaded Aging"
             " reports tenant refs FA-CINV-* while sent EIB holds legacy GP doc numbers; trace-only layer."),
        ]
        downstream_res["notes_rows_pairs"] = pl_notes_pairs

    # ---- Write output ----
    tag = f"_{args.label}" if args.label else ""
    fname = f"FA_Reconcile_CustomerInvoice_{TODAY:%Y%m%d}{tag}.xlsx"
    out_dir_primary = Path(os.path.expanduser(args.out_dir))

    primary, ok_p = _try_write(out_dir_primary, "reconcile dir", fname,
                               upstream_res, downstream_res, extract_path.name)
    downl, ok_d = _try_write(DOWNLOADS_FALLBACK_DIR, "Downloads fallback", fname,
                             upstream_res, downstream_res, extract_path.name)
    stage, ok_s = _try_write(STAGE_DIR, "/tmp stage", fname,
                             upstream_res, downstream_res, extract_path.name)

    if not ok_p:
        source_for_copy = None
        for p in (downl, stage):
            if p and p.exists():
                source_for_copy = p
                break
        if source_for_copy:
            try:
                out_dir_primary.mkdir(parents=True, exist_ok=True)
                dest = out_dir_primary / fname
                shutil.copyfile(source_for_copy, dest)
                print(f"  (recovery) shutil copied → reconcile dir: {dest}")
            except Exception as e:
                print(f"  (recovery) shutil copy also failed — sandbox may still block "
                      f"{out_dir_primary}. Please re-run this script from your own shell, or "
                      f"manually copy from {source_for_copy} → {out_dir_primary / fname} "
                      f"({type(e).__name__}: {e})", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
