# Recon Rebuild — Three-Way Finance Reconciliation (ANCHOR DOC)

**Read this whole file before touching code.** This is the cold-start continuity doc for
rebuilding the finance review pack as a **three-way** reconciliation that reads the
**actual EIB tables** (not a re-derivation from the finance file).

---

## Status (start here)

- **GOAL:** produce a finance-facing review workbook that reconciles **our actual EIB
  output** to **finance's recon report**, using **our source extract** as the bridge.
- **DONE (07.08.26):** `scripts/generate_finance_review_three_way.py` built + validated.
  Reads the real DuckDB tables; reproduces the baseline (PO EIB=Finance 100.0% / 0 transform
  residual; Receipt clean-tie 4,656 penny-exact). Writes 5-sheet workbook
  `FA_ThreeWay_Reconciliation_YYYYMMDD.xlsx` to the reconcile dir.
  Run: `uv run python scripts/generate_finance_review_three_way.py`
- **Review pack COMPLETE (07.08.26):** `generate_finance_review_three_way.py` now has 7 sheets
  (Exec Summary, PO buckets, Receipt buckets, ByCompany, Local-vs-GBP FX, Drift & Exceptions,
  Notes), labels the 5 accepted-exception POs explicitly, no hardcoded totals. Reconciles:
  PO clean-tie 6,968 (100%), Receipt clean-tie 4,656, accepted exception 5, business-filtered
  1,438; 0 transform residual. Output: `FA_ThreeWay_Reconciliation_YYYYMMDD.xlsx` (reconcile dir).
- **NEXT ITEMS (open):**
  1. Re-run TRANSFORM+PUBLISH to persist the description-corruption clean (base file changed;
     the 4 corrupted-memo rows are still in the live tables/xlsx until re-run). Recon VALUES
     unaffected (clean touches free-text only), so the pack numbers hold; re-run for the
     published EIB Memo + load-safety.
  2. Optional: retire the old `generate_finance_review_commitment_receipts.py` (re-derivation
     engine) once finance confirms the three-way pack is the one they want.
- **Output dir:** `~/Documents/__data/excel/workday/reconcile/`
- **Supersedes:** `scripts/generate_finance_review_commitment_receipts.py` (Trae's) —
  that one RE-DERIVES from finance's 040826 StackReport + Python filters and never reads
  the EIB tables, so its "100% matched / £0 variance" is tautological (compares two
  hardcoded 040826 numbers). Do NOT reuse its engine; the layout functions are fine.

## ACCEPTED EXCEPTIONS (finance signed off — do NOT re-flag as a bug)

**Proactis 060826 export corruption — ACCEPTED by finance 07.08.26 to go into Workday as-is.**
The 060826 source export has embedded line-breaks in the Description field that merged/shifted
rows on 5 POs. Finance has confirmed they are happy to live with the resulting difference:

| PO | Our EIB | Finance 040826 | Note |
|---|--:|--:|---|
| NFCPO300004848 | £1,260 | £2,280 | 1 line absorbed into a blob |
| NFCPO300004851 | £500 | £1,500 | 2 lines absorbed |
| NFCPO300006489 | £6,746.94 | £1,242.41 | value misaligned (shifted column) |
| NFCPO300006502 | 0 (missing) | present | whole PO absorbed into 6489's blob |
| NFCPO300006503 | 0 (missing) | present | whole PO absorbed into 6489's blob |

These are exactly the three-way "extract drift" (3) + "finance-only" (2) buckets. **Root cause is
the Proactis export, not the pipeline.** The base text-clean (below) MUST STAY regardless — it strips
the embedded control chars so the Workday EIB load doesn't misalign; that is separate from finance
accepting the value difference. A clean Proactis re-export would fix the values but finance has
decided it is not needed for go-live.

## The three data sources (the "three ways")

| Level | What | Where | Role |
|---|---|---|---|
| **1. EIB (ours)** | `workday_supplier_po_commitment_stack` (PO), `workday_supplier_po_commitment_stack_receipt` (receipt) | DuckDB `~/Documents/__data/duckdb/rpatel.duckdb`; 1:1 with published xlsx in `~/Documents/` | what we LOAD |
| **2. Source (ours, 060826)** | `src_fin_supplier_po_commitment_stack_base` (from `fin_supplier_po_commitment_stack_report`) | DuckDB; raw file `~/Documents/__data/excel/finance_ref/supplier_purchase_order/prod/The FA - Commitment Stack Report as at 060826.xlsx` (Sheet1 only) | what we INGEST from (the bridge) |
| **3. Finance recon (040826)** | sheet **`StackReport v3`** (GBP cols) + `Variance EIB v FM v Receipts` col B (Tamanna target) | `~/Downloads/The FA - Commitment Stack Report as at 040826 USE THIS ONE.xlsx` | finance AUTHORITY/target |

## The three-way logic (what each comparison proves)

- **[A] EIB ↔ Source(060826)** = *transform integrity* — did we load the source faithfully? Target ~100%.
- **[B] Source(060826) ↔ Finance(040826)** = *extract drift* — the two Proactis pulls at different dates. Explainable.
- **[C] EIB ↔ Finance(040826)** = *headline* — every non-match MUST be explained by [B] drift, not transform error.

## VALIDATED BASELINE (GBP basis, per PO) — the rebuild must reproduce these

Aggregates: **Finance £97,868,537 · Source £102,408,426 · EIB £97,850,875**

- **[A]** 6,971 POs both; EIB==Source penny-exact **6,961 (99.9%)**; in-source-not-EIB **2,802 / £4.53m** (business-gate filtered, legitimate); in-EIB-not-source **0**.
- **[B]** identical **6,958 (99.8%)**; drifted **13** (Σ|Δ| £34,924); finance-only 2; source-only 2,802 (same as the gated set).
- **[C]** EIB vs Finance **6,968 / 6,971 (100.0%)**; 3 non-matches, **ALL explained by drift**; **0 transform residual**. (NFCPO300004848 £1,260 vs £2,280 is one of the 3 = drift, proven.)
- **Receipt side** (hard-only): Finance Hard£ ↔ Source Hard£ ↔ Receipt EIB (`column_bz`). Same pattern. Receipt EIB Σ = £26,487,648.24 local; = PO hard commitment to the penny.

## Accumulated Active Constraints

1. **Reconcile on GBP basis** (`soft_commitment_value_gbp + hard_commitment_value_gbp`) for cross-currency summability. The EIB LOADS **local** (`column_gf` / `column_bz`); reconcile the economic value in **GBP**. Present both.
2. **PO EIB = full gross commitment (soft + hard, un-netted).** Receipt EIB = **hard only** (GRNI = received-not-invoiced). Receipt is commitment-derived (does NOT read the receipt extract).
3. **Read the ACTUAL EIB tables** — never re-derive the loaded numbers from the StackReport.
4. **No hardcoded magic totals** (Trae hardcoded 212,028,910.77 and 26,505,310.65). Compute everything from the live tables + files.
5. **The 7 business pre-gates** live in `src_fin_supplier_po_commitment_stack` (zero/old-soft/fully-del-closed/7-PO list). The EIB already reflects them — do not re-implement.

## Per-item verification (the three-way query — run to reproduce the baseline)

```python
# GBP per PO: fin (StackReport v3 Soft£+Hard£) / src (base) / eib (staged) — then [A][B][C]
# fin from 040826 StackReport v3 cols "Soft Commitment Value (£)","Hard Commitment Value (£)"
# src = SUM(soft_commitment_value_gbp+hard_commitment_value_gbp) FROM src_fin_supplier_po_commitment_stack_base GROUP BY order_number
# eib = same SUM FROM src_fin_supplier_po_commitment_stack (staged) GROUP BY order_number
# [A] eib vs src penny-exact; [B] src vs fin; [C] eib vs fin, non-matches explained where eib==src.
```
Expected: [A] 99.9%, [B] 99.8%/13 drifted, [C] 100.0% / 3 drift / 0 residual.

## Gotchas

- **DBeaver holds the DuckDB write lock.** If a query errors with "Conflicting lock … DBeaver", ask the user to disconnect, or read the published xlsx under `~/Documents/` instead.
- **060826 source file has ONLY `Sheet1`** (raw extract). **040826 file has `StackReport v1/v2/v3`** (finance analysis tabs). Trae's script defaults to 060826 but silently falls back to 040826 (060826 not in ~/Downloads + no StackReport v3) — that's why it ran on the wrong extract.
- **StackReport v3 GBP columns:** `Soft Commitment Value (£)`, `Hard Commitment Value (£)`; order col `Order Number`.
- **Description corruption:** 4 source rows (NFCPO300004848 L1/L2, NFCPO300004851 L3, NFCPO300006489 L1) had embedded tabs/CR/multi-row bleed. Fixed in `src_fin_supplier_po_commitment_stack_base.sql` with a 3-stage control-char clean on all free-text fields — needs a TRANSFORM+PUBLISH re-run to persist into the tables/xlsx.
- **Currency:** EIB local totals are mixed-currency (not summable as £). Always show GBP equivalent for the sterling grand total.
- **Finance `col I` (Hard Receipt Left) and `Hard_Barry` are unreliable** (double-count / quantity-inflated). Reconcile receipts against StackReport v3 Hard£ or col B, NOT col I / Hard_Barry.

## Build plan for `generate_finance_review_three_way.py`

1. Engine: load fin (040826 StackReport v3) + src (DuckDB base) + eib (DuckDB workday tables), per PO, GBP + local.
2. Sheets (keep Trae's style): 1-Exec Summary, 2-ThreeWay Tieback ([A][B][C] with £), 3-ByCompany, 4-ByCurrency, 5-DriftDetail (the drifted POs + gated POs), 6-Local vs GBP FX, 7-Notes.
3. Receipt sheet: same three-way for hard-only vs receipt EIB.
4. No hardcoded totals; write to `~/Documents/__data/excel/workday/reconcile/`.
5. Verify the workbook's headline numbers == the baseline above.
