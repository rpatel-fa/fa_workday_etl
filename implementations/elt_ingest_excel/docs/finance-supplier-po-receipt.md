# Finance — Supplier PO / Receipt → Workday

**Single source of truth for the supplier purchase-order + receipt migration.**
Cold-start sessions should be able to continue from this file alone.
**Read §Ownership & Scope, §Architecture Decisions and §Gotchas before touching any SQL.**

Config lives under `config/{ingest,transform/sql,validate,reconcile,publish}/finance/supplier_purchase_order_receipt/`.
Run commands are in [CLI.md](CLI.md) (not duplicated here).

This document absorbs the FORMER general-scope `recon-todo.md` (the old three-way reconciliation anchor doc) — do **not** maintain separate copies of the information below. A small `recon-todo.md` still lives in the transform SQL folder; it was the single-item build plan for OPEN-8 (the post-load 4th-layer tie-out) and is now **RESOLVED** (see §Post-load tie-out and §Open items [OPEN-8], CLOSED). It is retained only as a reference plan for a future in-pipeline 4th leg, not as outstanding work.

---

## Ownership & Scope

- Module: `supplier_purchase_order_receipt` / `supplier_purchase_order_commitment_stack`
- SQL folder: `elt_ingest_excel/config/transform/sql/finance/supplier_purchase_order_receipt/`
- Ingest config: [`config/ingest/finance/supplier_purchase_order_receipt.json`](../config/ingest/finance/supplier_purchase_order_receipt.json)
  - Primary source: `The FA - Commitment Stack Report as at 060826.xlsx` / Sheet1 → `fin_supplier_po_commitment_stack_report`
  - Secondary ingest (position 2, **UNUSED** in output chain — legacy): `The FA - PO Receipts Extract 060826 1627.xlsx` / Sheet1 → `fin_supplier_po_receipt` (candidate for removal post go-live)
- Verification DB: `~/Documents/__data/duckdb/rpatel.duckdb`
  - Read-only inside the TRAE sandbox. TRANSFORM/PUBLISH must be run from a user-owned shell (see §Sandbox limitations).
- Finance reconciliation workbook (target authority): `~/Downloads/The FA - Commitment Stack Report as at 040826 USE THIS ONE.xlsx`
  - Sheet `StackReport v3` is the GBP authority; sheet `Variance EIB v FM v Receipts` col B is Tamanna target.

---

## Status (start here)

**Architecture: Commitment-Stack-Only (as at 07.08.2026). Loaded to Workday and post-load reconciled CLEAN (09.08.2026).**

Both EIBs were loaded into the Workday tenant; the Workday-loaded extract — the test-tenant control report **and** the PRODUCTION extract (09.08.2026) — reconciles **penny-perfect and line-perfect** against the sent EIBs (see §Post-load tie-out and OPEN-8). This closes the source → sent → loaded loop.

Both EIBs are derived from a single source: the Proactis Commitment Stack Report workbook.
The Proactis PO Receipts Extract workbook is ingested (position 2 in ingest config) but **IS NO LONGER USED** in any transform SQL. All receipt values come from the commitment stack's `hard_commitment_value` column (= received-not-yet-invoiced GRNI).

### Three-way finance review pack

Produces the finance-facing Excel workbook that reconciles our actual EIB output to finance's recon report using our source extract as the bridge.

```bash
cd elt_ingest_excel
uv run python scripts/generate_finance_review_three_way.py
```

- **Output directory:** `~/Documents/__data/excel/workday/reconcile/`
- **Output file:** `FA_ThreeWay_Reconciliation_YYYYMMDD.xlsx` (7 sheets: Exec Summary, PO buckets, Receipt buckets, ByCompany, Local-vs-GBP FX, Drift & Exceptions, Notes)
- **Supersedes:** `scripts/generate_finance_review_commitment_receipts.py` (old engine re-derived from finance's 040826 file without ever reading EIB tables; its "100% match" was tautological. Retire once finance confirms the three-way pack.)

### Two EIB outputs produced

| Output SQL table | Publish target file | What it carries |
|---|---|---|
| `workday_supplier_po_commitment_stack` | `workday_fin_supplier_po_commitment_stack.xlsx` | Full gross PO commitment (Soft + Hard, **no netting**) — one service line per PO split. |
| `workday_supplier_po_commitment_stack_receipt` | `workday_fin_supplier_po_commitment_stack_receipt.xlsx` | GRNI accrual (Hard commitment only) — one synthetic receipt document per PO that has hard > 0.005 GBP. |

Key invariants by construction (cannot go wrong as long as the two workday SQL files read from `src_fin_supplier_po_commitment_stack`):

1. **Receipt EIB references always resolve.** Every `column_bu` (purchase_order_line) in the receipt EIB equals `column_ez` (service_order_line_id) of a real row in the PO EIB. No dangling refs. No EXISTS gate. No "fully-invoiced PO rescue" logic required.
2. **No over-receipt possible.** For every receipt row, PO `column_gf` (= soft_local + hard_local) ≥ receipt `column_bz` (= hard_local) — because receipt reads the hard value straight from the same staged row.
3. **Business filters live in one place.** All 4 pre-gates (zero-commit, old-soft cutoff, fully-delivered-closed, 7-PO exclusion list) + the 0.005 GBP gate are applied ONCE in `src_fin_supplier_po_commitment_stack.sql`; both EIBs inherit them automatically.

- **Green through TRANSFORM** — all loads + validators OK.
- **Validation issues** (surfaced via `validation_supplier_purchase_order_receipt_issues`): 34 MANDATORY_SUPPLIER + 34 UNMAPPED_SUPPLIER + 26 TRACKED_SPEND_NO_PROJECT + 3 MISSING_SPEND_CATEGORY + 1 STATUS_ALIGNMENT. These are finance-review items, not pipeline breaks.
- **Filter note:** `wd_supplier_id IS NOT NULL` filter is **commented out** in the PO EIB output SQL (line 81 of `workday_supplier_po_commitment_stack.sql`) and **not present** in the receipt EIB either. This is intentional for current testing until full supplier setup is guaranteed; unmapped suppliers (34 + 34, ~£41.7k exposed) flow through with NULL and appear in the validation issues table for finance sign-off.
- **Open confirmations before go-live:** §Open items C1–C4.

---

## Accepted Exceptions (finance signed off 07.08.2026)

**Proactis 060826 export corruption — ACCEPTED by finance. Do NOT re-flag as a bug.**
The 060826 source export has embedded line-breaks in the Description field that merged/shifted rows on 5 POs. Finance confirmed they are happy to live with the resulting difference:

| PO | Our EIB (GBP) | Finance 040826 (GBP) | Note |
|---|--:|--:|---|
| NFCPO300004848 | £1,260 | £2,280 | 1 line absorbed into a blob |
| NFCPO300004851 | £500 | £1,500 | 2 lines absorbed |
| NFCPO300006489 | £6,746.94 | £1,242.41 | value misaligned (shifted column) |
| NFCPO300006502 | 0 (missing) | present | whole PO absorbed into 6489's blob |
| NFCPO300006503 | 0 (missing) | present | whole PO absorbed into 6489's blob |

These are exactly the three-way "extract drift" (3) + "finance-only" (2) buckets. **Root cause is the Proactis export, not the pipeline.** The base-layer TAB-split + control-char clean (see §Base layer cleaning) MUST STAY regardless — it strips the embedded control chars so the Workday EIB load doesn't misalign; that is separate from finance accepting the value difference. A clean Proactis re-export would fix the values but finance has decided it is not needed for go-live.

---

## What the pipeline does

Migrates Proactis supplier PO commitment data into **two Workday EIB workbooks**:

1. **Purchase Order EIB** — all open PO service lines with soft/hard commitment values (full gross).
2. **Receipt EIB** — GRNI accrual lines (= received-not-invoiced Hard commitment) against those PO service lines (Hard-only, 1 receipt doc per PO).

**Single ingest source.** Both EIBs read from one ingested table (the commitment stack report).
The separate receipt extract table is not referenced by any output SQL. Source tables and target EIB names are defined in:

- Ingest: [`config/ingest/finance/supplier_purchase_order_receipt.json`](../config/ingest/finance/supplier_purchase_order_receipt.json) (authoritative — not duplicated here). The second entry (PO Receipts Extract) is a candidate for removal after go-live.
- Publish: [`config/publish/finance/publish_supplier_purchase_order_receipt.json`](../config/publish/finance/publish_supplier_purchase_order_receipt.json)

### Execution order

**DO NOT re-order SQL files.** [`order.txt`](../config/transform/sql/finance/supplier_purchase_order_receipt/order.txt) in the SQL folder declares execution order. Current (10 files):

1. `src_fin_workday_supplier_map.sql`
2. `src_fin_supplier_po_commitment_stack_base.sql`
3. `src_fin_supplier_po_commitment_stack_nrm.sql`
4. `src_fin_supplier_po_commitment_stack.sql`          — PO staged (gross)
5. `src_fin_supplier_po_commitment_stack_ref.sql`      — FDM mapping
6. `workday_supplier_po_commitment_stack.sql`          — PO EIB output
7. `workday_supplier_po_commitment_stack_receipt.sql`  — Receipt EIB output (GRNI)
8. `reconcile_supplier_po_unresolved.sql`              — unresolved suppliers (PO side)
9. `reconcile_supplier_po_unresolved_workday.sql`
10. `reconcile_supplier_po_unresolved_workday_update.sql`

### Transform structure

| Layer | Active file | Grain / purpose |
|---|---|---|
| Shared reference | [`src_fin_workday_supplier_map.sql`](../config/transform/sql/finance/supplier_purchase_order_receipt/src_fin_workday_supplier_map.sql) | vendor/company → Workday supplier (unnests Workday extract `reference_id` tokens). Single token → one supplier_id; exact `vendor-company` first, bare `vendor` fallback in nrm layer. |
| PO raw base | [`src_fin_supplier_po_commitment_stack_base.sql`](../config/transform/sql/finance/supplier_purchase_order_receipt/src_fin_supplier_po_commitment_stack_base.sql) | semantic aliases + dual text/numeric value cols + `commitment_stack_line_key` + **targeted** Description clean (TAB split + C0/DEL strip only). Footer row filter (END OF FILE + blank order_number). |
| PO normalized | [`src_fin_supplier_po_commitment_stack_nrm.sql`](../config/transform/sql/finance/supplier_purchase_order_receipt/src_fin_supplier_po_commitment_stack_nrm.sql) | supplier map (exact-then-bare) → date parse (Excel serial + timestamp support) → `commitment_class` (SOFT/HARD/MIXED/NONE) + `nrm_company_id` literal. Includes two vendor-code overrides (FAPO200025244 STE096→STE090; NFC-NAT008→NAT015). |
| PO staged | [`src_fin_supplier_po_commitment_stack.sql`](../config/transform/sql/finance/supplier_purchase_order_receipt/src_fin_supplier_po_commitment_stack.sql) | 4 business pre-gates → full-gross remaining cols (GREATEST(x,0) non-negative wrap + 2dp round) → 0.005 GBP final gate → `spreadsheet_key` → `po_line_pos_id`. |
| PO FDM ref | [`src_fin_supplier_po_commitment_stack_ref.sql`](../config/transform/sql/finance/supplier_purchase_order_receipt/src_fin_supplier_po_commitment_stack_ref.sql) | FDM mapping (CORE + PRISM hierarchy joins) + spend category + cost centre from `fin_ref_wd_fdm_mapping` + `reporting_type` derived from account code prefix. |
| PO Workday output | [`workday_supplier_po_commitment_stack.sql`](../config/transform/sql/finance/supplier_purchase_order_receipt/workday_supplier_po_commitment_stack.sql) | unified PO EIB (column_ez = `order_number_lineNum_position`; column_gf = full gross soft_local + hard_local; JPY → BIGINT as VARCHAR; supplier email/print issue_option; `wd_supplier_id IS NOT NULL` filter commented out). |
| **Receipt Workday output** | [`workday_supplier_po_commitment_stack_receipt.sql`](../config/transform/sql/finance/supplier_purchase_order_receipt/workday_supplier_po_commitment_stack_receipt.sql) | **receipt EIB, derived entirely from staged PO hard commitment.** One receipt document per PO (`REC-` + order_number). Gate: raw `hard_commitment_value_gbp > 0.005`. No separate receipt extract used. JPY → BIGINT VARCHAR; document_date = commitment_date proxy clamped to `GREATEST(commitment_date, order_date)` (never before the PO date); submit='Y' on all lines. |
| Unresolved suppliers (layer 1) | [`reconcile_supplier_po_unresolved.sql`](../config/transform/sql/finance/supplier_purchase_order_receipt/reconcile_supplier_po_unresolved.sql) | distinct (BU, supplier_code, supplier_name) rows where `wd_supplier_id IS NULL` → `MATCH_ADD_REFERENCE_ID` vs `GENUINE_NEW` |
| Unresolved suppliers (layer 2) | [`reconcile_supplier_po_unresolved_workday.sql`](../config/transform/sql/finance/supplier_purchase_order_receipt/reconcile_supplier_po_unresolved_workday.sql) | workday-side enrichment of the unresolved set |
| Unresolved suppliers (updates) | [`reconcile_supplier_po_unresolved_workday_update.sql`](../config/transform/sql/finance/supplier_purchase_order_receipt/reconcile_supplier_po_unresolved_workday_update.sql) | update projection for review-pack / manual-update flow |

### Live output chain

```
The FA - Commitment Stack Report ingested → fin_supplier_po_commitment_stack_report
  → src_fin_supplier_po_commitment_stack_base.sql      (footer filter + semantic aliases + TAB-split Description clean)
  → src_fin_supplier_po_commitment_stack_nrm.sql       (supplier map: exact vendor-co first, bare vendor COALESCE; date parse; commitment_class)
  → src_fin_supplier_po_commitment_stack.sql           (4 business pre-gates + GREATEST(0,x) + 0.005 GBP gate)
  → src_fin_supplier_po_commitment_stack_ref.sql       (FDM / spend / cost centre / reporting_type)
    ├─→ workday_supplier_po_commitment_stack.sql         → PO EIB
    │     Publish: workday_fin_supplier_po_commitment_stack.xlsx
    │     (full gross commitment: soft + hard, no netting; JPY integer, else 2dp DECIMAL VARCHAR)
    │
    └─→ workday_supplier_po_commitment_stack_receipt.sql → Receipt EIB
          Publish: workday_fin_supplier_po_commitment_stack_receipt.xlsx
          (hard commitment value only = GRNI; 1 synthetic receipt doc REC-{order_number} per PO)
```

### Legacy files (NOT in order.txt — retained for forensic reference only)

These files implement the old Proactis-receipt-based pipeline and are dead code.
Safe to delete after go-live (see OPEN-3):

- `src_fin_supplier_po_receipt_base.sql`   — raw Proactis receipt rows (UNUSED)
- `src_fin_supplier_po_receipt_nrm.sql`    — receipt allocation/netting (UNUSED)
- `src_fin_supplier_po_receipt.sql`        — old receipt staged (UNUSED)
- `workday_supplier_po_receipt.sql`        — old receipt EIB with EXISTS gate (UNUSED)

Any references to "receipt allocation", "3-tier LEAST cap", "3-signal GREATEST
netting", "EXISTS gate on cte_po_split_eligible", "NO_PO_FULLY_INVOICED bucket",
"receipt_backed_amount rescue", or "FAPO200060260 EXISTS-gate bug" apply ONLY to
this legacy pipeline and have been superseded.

---

## How the two tables get populated (novice walkthrough)

New to this pipeline, or to finance? Read this section first — it uses **real, queryable POs**
so you can verify every number in DuckDB. The dense reference (Decisions, Soft & Hard, Receipt
generation) follows below.

### The one idea: a PO moves through a lifecycle

```
   SOFT commitment            HARD commitment              (invoice)
"ordered, not received"  →  "received, not invoiced"  →   paid / cleared
   a promise on paper          a real liability (GRNI)
```

- **Soft** = ordered but nothing has arrived yet. A budget reservation; nothing hits the accounts.
  **Never** produces a receipt.
- **Hard** = goods/services received but the supplier's invoice has not been matched yet. This is a
  real liability — **GRNI (Goods Received, Not Invoiced)** — and must be accrued. It is the **sole**
  source of the Receipt EIB.

Two tables answer two different questions:

| Table | Question it answers | Value it carries |
|---|---|---|
| `workday_supplier_po_commitment_stack` (PO) | "What did we order in total?" | **Full gross** = soft + hard |
| `workday_supplier_po_commitment_stack_receipt` | "What has arrived and must be accrued?" | **Hard only** (GRNI) |

Only hard can be receipted — you can only receipt goods that were actually received.

### Column population at a glance

**PO EIB** — one row per split (`order_number _ line _ position`):

| Column | Meaning | Source |
|---|---|---|
| `column_j` | document_number | `order_number` (all splits share it → **one PO document**) |
| `column_ez` | service_order_line_id | `order_number_line_position` (unique line key) |
| `column_gf` | extended_amount | `nrm_po_export_local` = **soft_local + hard_local** |
| `column_u` | document_date | `order_date` |
| `column_ae` | due_date | **NULL** (cleared — consultant blanks it in the generated EIB) |

**Receipt EIB** — subset of PO splits where raw `hard_gbp > 0.005`, bundled per PO:

| Column | Meaning | Source |
|---|---|---|
| `column_d` | receipt (document) | `'REC-' \|\| order_number` (**one per PO**) |
| `column_bt` | line rank within the doc | `ROW_NUMBER()` within the PO |
| `column_c` / `column_e` | add_only / auto_complete | `'Y'` on rank 1, `'N'` otherwise |
| `column_bu` | purchase_order_line | `po_line_pos_id` (**= PO `column_ez`, 1:1 link**) |
| `column_bz` | extended_amount | **hard_local only** |
| `column_y` | document_date | `GREATEST(commitment_date, order_date)` (never before the PO date) |

### Example 1 — Hyve `FAPO200064154`: pure soft (the simplest case)

One line, one position, **SOFT £1,730**, hard £0.

- **PO EIB:** 1 row → `column_ez = FAPO200064154_1_1`, `column_gf = 1730.00` (1730 soft + 0 hard).
- **Receipt EIB:** **nothing** — hard = 0 fails the `hard_gbp > 0.005` gate. Nothing received → nothing to accrue.

*A soft-only PO appears in the PO table and never in the receipt table.*

### Example 2 — `WNSLPO400030141`: multi-line, multi-position, hard + soft

Four splits. Line 3 is **one item split into three accounting positions** (same goods, three cost
strings); Line 4 is a separate soft line:

```
ln pos  cls   hard_loc  soft_loc
 3   1 HARD    498.33      0.00     ┐ £1,495 hard, split 33/33/34
 3   2 HARD    498.33      0.00     ┤ across three cost strings
 3   3 HARD    498.34      0.00     ┘
 4   1 SOFT      0.00   1390.00       still only ordered
```

- **PO EIB:** 4 rows → `column_gf` = 498.33 / 498.33 / 498.34 / 1390.00 (whole order £2,885).
- **Receipt EIB:** only the 3 hard lines survive the gate (the £1,390 soft line drops); all bundle
  into **one** document `REC-WNSLPO400030141`:

```
column_bt  column_bu (= PO column_ez)   column_bz  column_c
   1       WNSLPO400030141_3_1           498.33     Y   ← new-document header line
   2       WNSLPO400030141_3_2           498.33     N
   3       WNSLPO400030141_3_3           498.34     N
```

£1,495 of goods arrived → accrued as GRNI across the three cost splits; the £1,390 soft line gets no receipt.

### Example 3 — `WNSLPO400032357`: MIXED splits (proves gross vs hard to the penny)

Six splits; Lines 5–6 are **MIXED** (soft £0.01 + hard £2,832.63):

| | Lines 1–4 (HARD) | Lines 5–6 (MIXED) |
|---|---|---|
| PO EIB `column_gf` (soft + hard) | 2832.63 | **2832.64** |
| Receipt EIB `column_bz` (hard only) | 2832.63 | **2832.63** |

The 1p difference on Lines 5–6 is the soft (ordered-not-received) sliver: carried in the PO, excluded
from the receipt accrual. Because receipt hard ≤ PO gross on every split, **over-receipt is impossible
by construction**.

---

## Architecture decisions (live)

### Decision 1 — PO = Full Gross Commitment (no netting)

Finance rule confirmed: **a PO is a COMMITMENT, not a P&L accrual.** PO must carry
the full gross Soft + Hard commitment value regardless of invoicing status.
Soft = ordered-not-received (never invoiced → never netted); the invoice overlap
is receipt-side only.

High-level implementation summary below; full step-by-step (base → nrm → staged →
PO EIB column transformations, commitment_class derivation, raw-vs-wrapped choices,
and the 0.005 GBP gate on RAW values) lives in §Soft & Hard Commitments.

Implementation in `src_fin_supplier_po_commitment_stack.sql`:
```
nrm_remaining_share    = 1.0                          (kept for column compatibility; no effect)
nrm_soft_remaining_*   = GREATEST(ROUND(raw soft, 2), 0)   (non-negative, 2dp)
nrm_hard_remaining_*   = GREATEST(ROUND(raw hard, 2), 0)   (non-negative, 2dp)
nrm_po_export_local    = soft_local + hard_local       (FULL GROSS)
nrm_po_receipt_rescued_ind = 'N'                      (never set — rescue path removed in this model)
```
The 3-signal netting CTEs, invoice share, and rescue/over-receipt logic from the
legacy pipeline have been REMOVED from the staged layer; only full-gross columns
remain.

**PO EIB key mechanics (`workday_supplier_po_commitment_stack.sql`):**
- **Spreadsheet_key (column_b):** DENSE_RANK() over (nrm_company_id, wd_supplier_id, order_number) — one row per (company, supplier, PO).
- **Header flags (column_c add_only, column_d auto_complete, column_h submit):** `'Y'` on **every** service line (no header-vs-child distinction in the PO EIB — every line carries a full doc header reload; Workday dedupes by document_number).
- **Comment column (column_e):** concatenation, in order: `invoice_status | delivery_status | Source: Commitment Stack Report | Commitment: SOFT/HARD/MIXED (if classified) | Period: {reporting_period} | RECEIPT-BACKED (rescue; always empty today — rescued_ind = 'N' everywhere)`.
- **Document date (column_u):** `strftime(order_date, 'dd/mm/yyyy')` — PO document uses PO order date, not commitment date.
- **Due date (column_ae):** `NULL` — explicitly blanked (Workday consultant blanks this in the generated EIB; no finance-sign-off default).
- **Issue option (column_al):** `CASE WHEN wd_supplier_email_exists_ind = 1 THEN 'EMAIL' ELSE 'PRINT' END` — the `wd_supplier_email_exists_ind` flag comes from the supplier map projection (`src_fin_workday_supplier_map.sql`) which checks `default_po_issue_email` OR `primary_email_address` non-null in the Workday supplier extract; EMAIL when either is populated.
- **Goods line stubs (column_bq, column_bu Goods line_company, column_bz Goods item_description, column_cj Goods spend_category, column_dg Goods unit_of_measure):** All set to `NULL`. The PO EIB uses the Service Line object only (service line rowid column_ex = `po_row_id`); never populates Goods line parallel attributes.
- **Service line numbering (column_ex, column_ez, column_fa, column_fb):** column_ex / column_fa = `po_row_id`; column_ez `service_order_line_id` = `po_line_pos_id` (the 3-part composite key that receipt column_bu will JOIN to); column_fb `line_company` = `nrm_company_id`.
- **Service line item description (column_fc):** plain `line_description` (inherits the base-layer TAB-split + C0/DEL + DEL strip clean).
- **Resource / spend category (column_fk):** `COALESCE(spend_category_override_ref_id, spend_category_ref_id, 'Other_Opex')` — 3-tier precedence:
  1. Explicit override from accounting string override match (`account_string_override_key` join in `src_fin_supplier_po_commitment_stack_ref.sql`)
  2. Standard FDM mapping for the spend category segment
  3. Literal fallback `Other_Opex` if neither FDM row matches (suppliers like the 34+ unmapped have no FDM row yet)
- **Service line FDM coding (columns_gv through _hd):** 9 FDM dimensions exactly as projected in `src_fin_supplier_po_commitment_stack_ref.sql` — cost_centre, reporting_type, event, fa_learning_courses, camps, programmes, grant, commercial_partner, project (capex + opex coalesced). If the ref join is NULL (no FDM mapping found) Workday will reject the row — `validate_supplier_po_commitment_stack_missing_refs.sql` surfaces those rows.
- **WHERE filter:** `wd_supplier_id IS NOT NULL` is intentionally COMMENTED OUT (line 81). Unmapped suppliers flow to the EIB with NULL column_q and are tracked via validation UNMAPPED/MANDATORY buckets instead of being dropped early.

**Negative Hard rows:** 13 rows in the source carry `hard_commitment_value_gbp < 0` (net £-52.8k). These are zeroed by the `GREATEST(x, 0)` wrapper on output and automatically dropped from the receipt EIB (which gates on raw `hard_gbp > 0.005`). They pass the PO pre-gate only where the total `soft_gbp + hard_gbp > 0` after pre-gate, then zero out individually in the output.

### Decision 2 — Receipt EIB = Hard Commitment Value (GRNI), sourced from staged PO

The Proactis receipt extract's `receipted_quantity` is unreliable (~56% of rows
carry a unit count, not a £ value; witness FAPO200023089 loaded £280 vs true
£55,468 when the old pipeline used it). The commitment stack's
`hard_commitment_value` (= received-not-invoiced GRNI per PO split) is the
authority.

High-level implementation summary below; full step-by-step (cte_grni gate on RAW
Hard£ > 0.005, per-PO synthetic document bundling, cross-EIB linkage mechanism,
and the extended_amount formatting branch) lives in §Receipt generation process.

Implementation in `workday_supplier_po_commitment_stack_receipt.sql`:
- One row per (PO, Line, Position) WHERE raw `hard_commitment_value_gbp > 0.005` (note: raw value used for gate, not GREATEST-wrapped value, so negative-hard rows are excluded cleanly)
- `column_bz extended_amount` = `hard_commitment_value_local` via `grni_export_amount_text` (JPY → BIGINT as VARCHAR; else 2dp DECIMAL as VARCHAR — matches EIB template integer rule for JPY)
- One synthetic receipt doc per PO: `column_d / column_h = 'REC-' || order_number`
- `column_bu purchase_order_line = po_line_pos_id` (= PO `column_ez` — 1:1 link)
- `column_y document_date = strftime(GREATEST(commitment_date, order_date), '%d/%m/%Y')` — receipt date uses commitment_date (GRNI-recognised date) as proxy, **clamped to never earlier than the PO order_date** (no true receipt event date exists in the source; Workday rejects any receipt dated before its PO). This partially resolves OPEN-4 (see §Open items — the forward-clamp is locked in; only the question of commitment_date vs a fixed cutover date remains for finance+consultant signoff).
- `column_j submit = 'Y'` on all lines (flag for Workday consultant review — OK to submit all)
- **Receipt EIB NULL stubs (column_g worker, column_i workday_locked):** both explicitly NULL. The receipt EIB only populates: header flags, doc number, comment, company (column_k), document_date (column_y), memo (column_ab), line_id (column_bt), PO ref (column_bu), extended_amount (column_bz).
- **No FDM coding on the receipt EIB.** Workday copies the cost-centre / spend-category / grant / project FDM distribution directly from the referenced PO service line. That's why the cross-EIB `column_bu = po_line_pos_id` link is the single critical invariant; if that link is right, the receipt automatically inherits the PO's coding.

Result: Receipt EIB is by-construction consistent with PO hard commitment and
cannot over-receive (PO `column_gf` = soft+hard ≥ hard = receipt `column_bz`
for every row). No EXISTS gate, no referential-integrity rescue, no 3-cap math.

### Decision 3 — Business pre-gates live in ONE place (single-source)

Four filters, applied ONCE in `src_fin_supplier_po_commitment_stack.sql` and
inherited by BOTH EIB outputs (since receipt EIB reads from the staged PO):

1. Zero commitment: `soft_gbp + hard_gbp > 0`  (drops DROP_ZERO_COMMITMENT)
2. Old SOFT cutoff: `NOT (class = SOFT AND commitment_date < 2025-08-01)`  (drops DROP_OLD_SOFT)
3. Fully Delivered (Closed) status: `delivery_status != 'Fully Delivered (Closed)'`
4. 7 explicit PO numbers from "Adjustments Made" sheet:
   `FAPO200010562, NFCPO300004541, WNSLPO400011283, FAPO200035755, FAPO200039033, WNSLPO400019136, FAPO200049511`

Final 0.005 GBP gate at the bottom of the same file: `soft_gbp + hard_gbp > 0.005`.
Because receipt EIB is `SELECT FROM staged PO WHERE hard_gbp > 0.005`, it automatically respects all 4
business pre-gates without re-implementing them.

### Decision 4 — Three-way reconciliation reads the REAL EIB tables (not re-derived)

The review pack must read from the actual DuckDB EIB tables, not re-derive numbers from finance's file. Three legs:

- **[A] EIB ↔ Source(060826)** = *transform integrity* — did we load the source faithfully? Target ~100%.
- **[B] Source(060826) ↔ Finance(040826)** = *extract drift* — the two Proactis pulls at different dates. Explainable.
- **[C] EIB ↔ Finance(040826)** = *headline* — every non-match MUST be explained by [B] drift, not transform error.

The three data sources:

| Level | What | Where |
|---|---|---|
| 1. EIB (ours) | `workday_supplier_po_commitment_stack` (PO), `workday_supplier_po_commitment_stack_receipt` (receipt) | DuckDB rpatel.duckdb; 1:1 with published xlsx in ~/Documents/ |
| 2. Source (060826) | `src_fin_supplier_po_commitment_stack_base` (from `fin_supplier_po_commitment_stack_report`) | DuckDB; raw file in `~/Documents/__data/excel/finance_ref/supplier_purchase_order/prod/` (Sheet1 only) |
| 3. Finance (040826) | Sheet **`StackReport v3`** (GBP cols: Soft Commitment Value (£), Hard Commitment Value (£)) + `Variance EIB v FM v Receipts` col B | `~/Downloads/The FA - Commitment Stack Report as at 040826 USE THIS ONE.xlsx` |

---

## Soft & Hard Commitments (definition + pipeline semantics)

This section is the detailed implementation complement to Decision 1 (PO = full gross) and
Decision 2 (receipt = hard commitment only). Read those first for the finance rationale,
then this for exactly what the pipeline does at each layer.

### Definitions

The source workbook (Proactis Commitment Stack Report) carries two commitment value
columns in both transaction-currency and GBP:

| Value | Meaning (Proactis semantics) | Pipeline role |
|---|---|---|
| `soft_commitment_value_local / _gbp` | **Ordered-not-yet-received.** Goods/services on the PO that have been placed with the supplier but have no evidence of delivery or invoice. They will (when received) convert to Hard, and (when invoiced) clear the accrual. Never generates a receipt. | **Feeds PO EIB only.** The PO is a commitment — soft is the purest expression of it. Never appears in the receipt EIB. |
| `hard_commitment_value_local / _gbp` | **Received-not-yet-invoiced (= GRNI accrual).** Goods/services have been received/delivered (or at least classified in Proactis as delivered against the PO), but the supplier's invoice has not yet been matched, posted, or paid. Generates the Receipt EIB. | **Feeds PO EIB (as part of the gross) AND is the sole source of the Receipt EIB.** |

Important: this mapping from Proactis columns is the **only** place Soft/Hard is classified.
There is no downstream override, no finance-sign-off flag, no cross-check against
`receipt_status` / `invoice_status` columns in the workbook (we tried that in the legacy
pipeline and it produced the unreliable NO_PO_FULLY_INVOICED rescue path).

### Commitment class derivation (4 buckets)

Derived **once, in `src_fin_supplier_po_commitment_stack_nrm.sql`** from the raw
source numeric Soft£ and Hard£ (NOT from any GREATEST-wrapped column, NOT after rounding):

```
IF  Soft_gbp > 0 AND Hard_gbp > 0  →  MIXED
ELIF Hard_gbp > 0                  →  HARD
ELIF Soft_gbp > 0                  →  SOFT
ELSE (both zero or negative)       →  NONE
```

Note: the check uses `COALESCE(col, 0) > 0` so NULLs behave like zeros.

This class is used for two decisions in the business gates:

1. `SOFT + commitment_date < 2025-08-01` → dropped at the staged layer (DROP_OLD_SOFT).
   Only the oldest pure-soft lines (no delivery evidence) are cut off.
   `HARD` / `MIXED` / pre-cutoff `SOFT` are **all preserved** past this gate.
2. Class is surfaced as a Comment suffix in both EIBs (`| Commitment: SOFT/HARD/MIXED`)
   so Workday consultants can eyeball the load.

### Value columns through the pipeline (Soft & Hard)

Every commitment-value column has two variants:
- `*_local` = transaction currency (used to WRITE the EIB extended_amount — Workday posts in the PO's currency).
- `*_gbp`   = Pound Sterling (used for BUSINESS GATES and FINANCE RECONCILIATION tie-outs — cross-company summability).

Both variants go through the same math; only their numeric source differs.

#### 1. Ingest → base layer (`src_fin_supplier_po_commitment_stack_base.sql`)
No value transformation — values are kept as-is from the workbook; base simply aliases
the workbook columns (`soft_commitment_value_1` and `hard_commitment_value_1` are the
`_local` pair; the `_gbp` columns are directly named `soft_commitment_value_gbp` /
`hard_commitment_value_gbp`) and adds `*_text` TRIM'd copies.

**Gates at this layer:** none (values are not even examined). Only footer rows filtered
and Description cleaned.

#### 2. Base → nrm layer (`src_fin_supplier_po_commitment_stack_nrm.sql`)
No value transformation — still raw values (so commitment_class above, and the 2 pre-gates
below, operate on pre-wrapped numbers).

**Gates at this layer:** none. Zero rows dropped (filters are deferred to staged).

#### 3. Nrm → staged layer (`src_fin_supplier_po_commitment_stack.sql`)
This is where Soft & Hard get wrapped into the staged `nrm_{soft,hard}_remaining_{gbp,local}`
columns that the two EIBs ultimately read from.

Applied **in order**, one line per split (company, order, line, pos):

1. **Business pre-gates (DROP happens first, values untouched):** rows where
   `Soft_gbp + Hard_gbp <= 0` are dropped (zero commitment); rows where
   `class = SOFT AND commitment_date < 2025-08-01` are dropped (old soft); rows
   where delivery_status is Fully Delivered (Closed) are dropped; 7 explicit POs
   are dropped. Values are not examined in these gates except as a row-filter.
2. **Non-negative wrap (per value column, independently):**
   `nrm_soft_remaining_gbp   = ROUND(GREATEST(COALESCE(raw_soft_gbp, 0), 0), 2)`
   `nrm_soft_remaining_local = ROUND(GREATEST(COALESCE(raw_soft_local, 0), 0), 2)`
   `nrm_hard_remaining_gbp   = ROUND(GREATEST(COALESCE(raw_hard_gbp, 0), 0), 2)`
   `nrm_hard_remaining_local = ROUND(GREATEST(COALESCE(raw_hard_local, 0), 0), 2)`
   Soft and Hard are wrapped independently — this is why a MIXED row can have one
   side wrapped to zero and the other preserved.
3. **PO gross export column (PO EIB will read this):**
   `nrm_po_export_local = ROUND(COALESCE(raw_soft_local, 0) + COALESCE(raw_hard_local, 0), 2)`
   **Note: uses RAW values (not GREATEST-wrapped) for sum, then single ROUND.**
   The rationale: the business gates already dropped zero-total rows, so the only
   negative residuals are float dust £0.00 and the GREATEST wrap only catches them.
   Using raw for the sum avoids double-rounding drift between the EIB gross total
   and finance's Soft£+Hard£ target. (The individual remaining columns ARE
   GREATEST-wrapped so they can safely appear as reconciliation columns.)
4. **Final 0.005 GBP gate (on RAW, not wrapped):** the line is dropped if
   `raw_soft_gbp + raw_hard_gbp <= 0.005`. This is on RAW values, not wrapped, so a
   line with (Soft£ -0.004, Hard£ 0.008 — net +0.004) is correctly culled even
   though the GREATEST-wrapped columns would sum to 0.008.

**Receipt EIB does NOT read from here yet — it has its own separate gate below.**

#### 4. Staged → PO EIB (`workday_supplier_po_commitment_stack.sql`)
PO EIB `extended_amount (column_gf)` = the staged `nrm_po_export_local` pre-rounded
total, re-formatted for EIB text output:
- JPY → `CAST(CAST(ROUND(nrm_po_export_local, 0) AS BIGINT) AS VARCHAR)` (whole yen)
- else → `CAST(CAST(ROUND(nrm_po_export_local, 2) AS DECIMAL(18,2)) AS VARCHAR)`

Because staged `nrm_po_export_local = soft_local + hard_local, PO EIB always carries
the full gross commitment. There is no filter on class, no "MIXED-only netting, no
receipt-backend rescue. Receipt EIB comment column has `| Commitment: {SOFT|HARD|MIXED}`
for load review.

**What the PO EIB does NOT have:** any per-column split of Soft vs Hard — the two are
summed in the extended_amount and the supplier/finance only see one combined figure.
The per-split Soft vs Hard is kept in the staged PO table (for reconciliation) and in
the comment column `| Commitment: …` tag for traceability in the EIB workbook.

All other PO EIB mechanics — spreadsheet_key DENSE_RANK pattern, header flags (all Y on every line), comment tag content order including Reporting Period suffix, the RECEIPT-BACKED tag (always empty), document_date vs order_date, due_date NULLed, issue_option email/print, goods line NULL stubs, service line numbering and FDM coding including the 3-tier resource_category precedence (override_ref_id > ref_id > Other_Opex), and the supplier filter commented-out so unmapped flow to Workday — are documented in Decision 1's PO EIB key mechanics bullet list. Don't duplicate those rules here.

#### 5. Staged → Receipt EIB (`workday_supplier_po_commitment_stack_receipt.sql`)
Receipt EIB only uses Hard — Soft never appears. See next section.

---

## Receipt generation process (GRNI from PO Hard commitment)

This is the detailed implementation complement to Decision 2. The short form:
**For every line in the staged PO that has raw `hard_commitment_value_gbp > 0.005`,
write one receipt service line referencing that exact staged PO service line, bundle
all receipt lines for the same PO into one synthetic `REC-{order_number}` receipt
document.**

The legacy pipeline tried to generate receipts from the Proactis PO Receipts Extract
workbook (position 2 in ingest config) — that path is now dead. Receipt generation
reads exclusively from the staged PO commitment table. No separate receipt workbook
is required; nothing is JOIN'd to `fin_supplier_po_receipt`. This guarantees the
cross-EIB link can never drift.

### Stage-by-stage receipt construction

Inputs (both 100% from staged PO — no new data, no new columns):
- Staged PO table: `src_fin_supplier_po_commitment_stack`
- Read columns: (company_shortcode, order_number, line_number, po_position, po_line_pos_id,
  nrm_company_id, wd_supplier_id, line_description, invoice_status, delivery_status,
  commitment_class, commitment_date, currency) AND the two raw values
  (hard_commitment_value_local / _gbp) AND the two wrapped remaining columns
  (nrm_hard_remaining_local / _gbp).

The receipt SQL never reaches back into base/nrm/raw ingest tables. Every value it
uses is already in the staged PO.

#### Step 1 — cte_grni: Gate the raw Hard commitment (receipt = HARD ONLY)

Filter to rows where:
`COALESCE(raw hard_commitment_value_gbp, 0) > 0.005`

**Important gate details:**
1. Gate is on **RAW, not GREATEST-wrapped, Hard GBP value** (same rationale as the 0.005
   staged gate — avoids sub-penny negative-float residuals).
2. Gate uses GBP, not local — this is a finance-control rule ("any material received
   GRNI, measured in the shared recon currency"), not a Workday posting rule. JPY
   receipts with tiny-yen local values that still exceed £0.005 GBP pass through
   correctly.
3. Soft is irrelevant to the gate — a MIXED split passes (its Hard£ is positive);
   a pure HARD split passes; a pure SOFT split never passes.
4. All four business pre-gates from the staged layer are **automatically respected**
   because receipt reads from staged PO (zero-commit/old-soft/fully-delivered/7-POs
   are already gone). There is no need to duplicate them in the receipt SQL.

Outcome of the gate: `cte_grni` contains every split (one row = one Workday receipt
service line) that carries any received-not-invoiced GRNI value.

#### Step 2 — Format receipt extended_amount (`grni_export_amount_text`)

Receipt EIB `extended_amount (column_bz)` is always the transaction-currency Hard
commitment, formatted identically to how the PO EIB formats its gross:
- JPY → whole-BIGINT VARCHAR: `CAST(CAST(ROUND(COALESCE(raw hard_local, 0), 0) AS BIGINT) AS VARCHAR)`
- else → 2dp DECIMAL VARCHAR: `CAST(CAST(ROUND(COALESCE(raw hard_local, 0), 2) AS DECIMAL(18,2)) AS VARCHAR)`

Notice: receipt reads RAW hard_local for the formatting, not GREATEST-wrapped
nrm_hard_remaining_local. This matches the PO EIB's choice to use the raw sum for
its gross — keeps the two formatting branches in lockstep. (GREATEST-wrapped
nrm_hard_remaining_local is still available for reconcile tie-out if needed.)

#### Step 3 — Build synthetic receipt documents (1 per PO)

Within cte_grni, group rows by `order_number`. For each such group:
- Assign a synthetic document_number / receipt_number:
  `synthetic_receipt_number = 'REC-' || order_number`
- Within each document, rank the lines using **numeric-first then lexical** ordering on
  `(line_number, po_position)`:
  ```
  ROW_NUMBER() OVER (
      PARTITION BY order_number
      ORDER BY TRY_CAST(NULLIF(TRIM(line_number), '') AS INTEGER), line_number,
               TRY_CAST(NULLIF(TRIM(po_position), '') AS INTEGER), po_position
  ) receipt_line_rnk
  ```
  The `TRY_CAST … AS INTEGER` primary key ensures pure-integer line/pos numbers sort
  `9, 10, 11` not `10, 11, 9`; the lexical fallback handles any alphanumeric line
  codes (e.g. `LineA`, `SCHED-1`) without error. This rank is the stable display
  order within the receipt document.
- Compute a global `spreadsheet_key` for the EIB using **ROW_NUMBER()** (different
  from the PO EIB, which uses DENSE_RANK):
  ```
  ROW_NUMBER() OVER (
      ORDER BY nrm_company_id, wd_supplier_id, order_number,
               TRY_CAST(NULLIF(TRIM(line_number), '') AS INTEGER),
               TRY_CAST(NULLIF(TRIM(po_position), '') AS INTEGER)
  ) spreadsheet_key
  ```
  Why ROW_NUMBER per line vs DENSE_RANK per PO? Because the receipt template
  spreadsheet_key is per **receipt service line**, not per receipt doc.

**One receipt doc per PO rule:** Even if a PO has 1 GRNI line under Line 1 Position 1
and 3 GRNI lines under Line 4 Positions 1/2/3 (accounting splits), all 4 lines go into
the **same** `REC-{order_number}` document. Receipt document count is always
`COUNT(DISTINCT order_number)` of cte_grni. Lines for a PO are never split across
multiple docs.

**Header-row flags (column_c add_only, column_e auto_complete):**
`CASE receipt_line_rnk WHEN 1 THEN 'Y' ELSE 'N' END`
The top-ranked line in each receipt doc carries Workday "this is a new document"
header flags; the remaining lines are child service lines of the same doc.

#### Step 4 — Cross-EIB reference link (by-construction always valid)

Receipt column_bu (purchase_order_line) = the staged PO's `po_line_pos_id` (which is
also PO EIB column_ez = service_order_line_id). **It is not a JOIN result.** Because
cte_grni is a row-subset of staged PO, `po_line_pos_id` on a cte_grni row maps to
exactly one staged PO row, which (since the PO EIB is 1:1 with staged) maps to exactly
one PO EIB row. No lookup, no EXISTS gate, no key-resolution logic is required; the
value is the same 3-part natural key on both sides.

Two derived invariants from this (both always 0):
- `COUNT(DISTINCT receipt.column_bu)` rows where PO EIB has no matching row (dangling reference check).
- Rows where `TRY_CAST(receipt.column_bz AS DECIMAL) > TRY_CAST(po.column_gf AS DECIMAL)`
  (over-receipt check). Because receipt extended_amount = hard_local ≤ soft_local + hard_local
  = PO extended_amount for the same split.

#### Step 5 — Remaining receipt fields (for completeness, not cold-start logic)

- `column_y document_date = strftime(GREATEST(commitment_date, order_date), 'dd/mm/yyyy')` — proxy
  clamped so the receipt is never dated before the PO's `order_date` (Workday rejects receipt <
  PO date); no receipt event timestamp exists in the commitment stack. (OPEN-4: the forward-clamp
  is locked in; only the choice between commitment_date overall vs a fixed cutover date remains
  with finance + Workday consultant. See Decision 2 for the same wording.)
- `column_f comment`: concatenation, in order: `invoice_status` + (space separator if both non-blank) + `delivery_status` + ` | Source: Commitment Stack (GRNI)` + ` | Commitment: HARD/MIXED` class tag (soft class can never reach receipt). Note the receipt source tag reads **"Commitment Stack (GRNI)"** — slightly different from the PO EIB tag which reads "Commitment Stack Report"; and there is NO Period: suffix and NO RECEIPT-BACKED rescue clause on the receipt side (the rescue path is dead anyway).
- `column_g worker = NULL`, `column_i workday_locked = NULL` — explicit NULL stubs.
- `column_ab memo = line_description` — split description copied verbatim for the
  Workday receipt memo field; Description TAB-split + C0/DEL clean from base layer
  is inherited automatically.
- `column_k company, PO EIB uses column_q (supplier); receipt uses column_k (company). The supplier column on receipt is NOT a separate column letter — it is baked into the header grouping via spreadsheet_key's wd_supplier_id term. Both values read from same staged nrm_company_id / wd_supplier_id so they agree.
- `column_j submit = 'Y'` on ALL lines (not just header); flag for Workday load review.
- `column_bt receipt_service_line_row_id = CAST(receipt_line_rnk AS VARCHAR)` — note: column name in the SQL CTE is `receipt_line_rnk` (shortened from earlier `receipt_service_line_rnk`; they are the same thing, display order within a doc).

### End-to-end example

One PO `FAPO200025244`:
- Line 1 / Position 1: Soft£1000, Hard£0 (pure SOFT)
- Line 2 / Position 1: Soft£0, Hard£5,468.00 (pure HARD)
- Line 2 / Position 2: Soft£50, Hard£250 (MIXED, accounting split — same line,
  two positions with different account strings and net_value shares)
- Line 3 / Position 1: Soft£40, Hard£0.003 (pure SOFT tiny — survives staged 0.005
  gate because Soft£+Hard£ = 40.003, but fails RECEIPT 0.005 Hard gate)

Staged PO has all 4 rows. Commitment classes: SOFT / HARD / MIXED / SOFT.

cte_grni: keeps Line2/Pos1 (Hard£5,468 > 0.005) + Line2/Pos2 (Hard£250 > 0.005).
Line1/Pos1 (Hard£0) and Line3/Pos1 (Hard£0.003) are culled.

Receipt doc `REC-FAPO200025244` has 2 service lines. PO EIB still has all 4 service lines.
Cross-EIB check: receipt Line2/Pos1 col_bu = `FAPO200025244_2_1` → matches PO EIB col_ez
= same value; receipt Line2/Pos2 → same. Receipt extended_amounts (£5,468 / £250) ≤ PO
extended_amounts (Soft+Hard gross = £5,468 / £300) → no over-receipt. Line3/Pos1 which
still appears in PO EIB (£40.003 gross) simply has no matching receipt row — correct,
because there's no received-not-invoiced value on it.

---

## Invariants

- **Service-line identity is the 3-part natural key `order_number + '_' + line_number + '_' + po_position`** — stored as `po_line_pos_id` in staged, `column_ez` in the PO EIB, and `column_bu` in the receipt EIB. This is the cross-EIB join key. **Never** revert `column_ez` to a per-table rowid — it breaks receipt linkage.
- **"Duplicate PO lines" are accounting/Event distribution splits**, not accidental dupes: one `line_number`, several `position`s, each with distinct accounting strings (activity/account/event/commercial) and its own `net_value_gbp` share. Each split = its own Workday service line.
- **Receipt EIB ↔ PO EIB link is 1:1 by position.** Because receipt EIB rows are SELECTed directly FROM staged PO WHERE hard > 0.005, the receipt's `po_line_pos_id` (= column_bu) maps to exactly one row in the PO EIB. No join, no gate, no lookup — it's the same staged row. No "PO Line not found" errors possible.
- **Currency basis is branched (vendor-confirmed).** In the PO commitment stack, `*_gbp` columns are GBP; `*_local` columns (soft_commitment_value_1 / hard_commitment_value_1) are transaction currency. Workday EIB `extended_amount` uses local-currency columns while business gates and audit tie-outs remain in GBP.
- **PO EIB extended amount is transaction currency full gross.** `column_gf` = `soft_commitment_value_local + hard_commitment_value_local` (= `nrm_po_export_local`). No netting, no share multiplication. JPY → BIGINT VARCHAR; all else → 2dp DECIMAL VARCHAR.
- **Receipt EIB extended amount is transaction currency Hard commitment only.** `column_bz` uses `hard_commitment_value_local` formatted via inline `grni_export_amount_text` (JPY → BIGINT as VARCHAR; else 2dp DECIMAL as VARCHAR — matches EIB template integer rule for JPY).
- **Receipt rows are HARD-only by design.** Every receipt EIB row has a Hard commitment from the staged PO because its gate is `WHERE raw hard_gbp > 0.005`. Soft commitment is pure PO-side (ordered − not-yet-received; it never has, and cannot have, a receipt row).
- **Supplier mapping is company-aware exact-then-bare.** Join order in nrm: `(nrm_supplier_code || '-' || nrm_company_name)` exact first, then bare `nrm_supplier_code` via `COALESCE`. The map is built in `src_fin_workday_supplier_map` which unnests pipe-delimited `reference_id` tokens from the Workday supplier extract.
- **Specific vendor overrides:** (1) PO `FAPO200025244` / supplier_code `FA-STE096` uses vendor id `STE090` (not the SPLIT_PART-derived `STE096`). (2) supplier_code `NFC-NAT008` (company NFC) uses vendor id `NAT015` (not SPLIT_PART-derived `NAT008`). Both hardcoded in `src_fin_supplier_po_commitment_stack_nrm.sql`.
- **Source values stay as text in nrm where arithmetic isn't needed.** All value fields have a `*_text` (TRIM'd raw) + numeric (TRY_CAST DOUBLE) pair; the pre-rounded/wrapped `nrm_po_export_local` column (or formatted inline `grni_export_amount_text`) is what actually feeds the EIB to avoid precision loss.
- **Negative value guard:** Both staged PO and receipt use `GREATEST(x, 0)` around remaining commitment value columns (2dp rounded) so no negative amounts reach the EIB. The receipt gate uses raw `hard_commitment_value_gbp` before GREATEST wrapping, so rows with negative hard never appear in receipt even if their parent PO had a positive net.
- **Base layer cleaning (targeted, NOT blanket 3-stage):**
  - **`line_description` only:** SPLIT_PART on first TAB (chr(9)) to fix Proactis CSV Description bleed → REGEXP_REPLACE strips C0 controls (0x00–0x1F) + DEL (0x7F) → NULLIF to empty string. CAST to VARCHAR first because Description column carries mixed cell types (INT, FLOAT, DATE, VARCHAR).
  - **All other free-text fields (`supplier_name`, `division_name`, `originator_name`, etc.):** Plain TRIM only. No punctuation stripping. The earlier design idea of stripping `"'<>|;{}` was **NOT implemented** — those chars appear in ~270 legitimate descriptions and are valid XLSX cell content.
  - **Join keys (`vendor_code`, `company_shortcode`, `order_number`, `line_number`, `po_position`):** TRIM only — never regex strip — to preserve key integrity.
- **Footer row filter (base layer):** Explicitly drops rows where `order_number` is blank AND rows where `companyshortcode` = `'END OF FILE'`. CSR CSV exports carry an `END OF FILE` trailing row plus blank footer rows; this catches both.
- **Description VARCHAR cast before regex.** The Stack Report Description column carries mixed cell types (INT, FLOAT, DATE, VARCHAR). Base layer CASTs to VARCHAR first before TRIM/regex/SPLIT_PART to avoid type-coercion explosions. **DO NOT remove this cast.**
- **JPY = INTEGER in EIB extended amounts.** Both EIBs format JPY values as BIGINT (no decimals) when writing to EIB text columns. This is a Workday EIB template currency rule — do **NOT** flatten to 2dp for JPY.
- **Both EIBs read from the same staged table (`src_fin_supplier_po_commitment_stack`).** If you edit the business gates, 0.005 gate, or remaining columns in staged, **you automatically change both EIB outputs together.** This is by design — do **NOT** add independent filters in either workday output SQL that could desync them.
- **Reconcile on GBP basis** (`soft_commitment_value_gbp + hard_commitment_value_gbp`) for cross-currency summability. The EIB LOADS local (`column_gf` / `column_bz`); reconcile the economic value in GBP. Present both.
- **Staged row-drop gates are 2 (not 3, not 4).** The reconcile layer buckets DROP_ZERO_COMMITMENT and DROP_OLD_SOFT_CUTOFF as actual row drops. Unmapped suppliers (`unmapped_supplier_rows` / `unmapped_supplier_amount`) are an **informational bucket only** — they flow all the way to Workday EIB with NULL supplier_id and are tracked via `validation_supplier_purchase_order_receipt_issues`. Therefore:
  - `expected_to_staged = base_rows - drop_zero_commitment - drop_old_soft` (not minus unmapped — you'd undercount)
  - `expected_to_workday = staged_rows` (1:1, no further subtraction)
  - `unexplained_workday` must be 0.
  The 4 business pre-gates listed in `src_fin_supplier_po_commitment_stack.sql` (zero + old-soft + fully-delivered-closed + 7 explicit POs) collectively deliver the base→staged reduction, but the reconcile buckets only surface the 2 explicit named drop-gates; the 7-PO + fully-delivered-closed reduction falls under generic staged-vs-expected coverage.
- **Validate + Reconcile reference-set rule (enforced in order.txt footers).** Validate/reconcile in-pipeline SQL MUST NOT reference any table outside (a) transform-layer-created, (b) this module's ingest-config-declared source extracts, or (c) post-load Workday extracts that have first passed through ingest-config + transform projection. Cross-module external dependencies (e.g. `fin_supplier_po_audit_commitment`) are only allowed in OFF-PIPELINE, ad-hoc-run SQL files with an explicit header banner stating the dependency. Do not silently add them to order.txt. The 5 `reconcile_cs_audit_*.sql` files live in this off-pipeline bucket.
- **No hardcoded magic totals** in reconciliation or report scripts. Compute everything from live tables + files.
- **Finance `col I` (Hard Receipt Left) and `Hard_Barry` are unreliable** (double-count / quantity-inflated). Reconcile receipts against StackReport v3 Hard£ or Variance sheet col B, NOT col I / Hard_Barry.

---

## Row funnel (commitment-stack-only architecture)

### PO commitment stack

| Stage | Rows | Notes |
|---|---:|---|
| Ingested report (Sheet1) | ~14,285 | Raw before base-layer footer filter |
| Base (`src_fin_supplier_po_commitment_stack_base`) | 14,285 | Footer filter applied (END OF FILE + blanks dropped); TAB-split Description |
| Nrm (`src_fin_supplier_po_commitment_stack_nrm`) | 14,285 | Supplier map + date parse — 0 row drops (filtering happens at staged) |
| After 4 business pre-gates → `cte_po_base_filtered` | 10,833 | Drops: zero-commit + old-soft + fully-delivered-closed + the 7 explicit POs |
| **After 0.005 GBP gate** → Staged (`src_fin_supplier_po_commitment_stack`) | Re-verify live | Full-gross rows where soft_gbp + hard_gbp > 0.005; GREATEST(0,x) applied to each value col |
| Workday EIB (`workday_supplier_po_commitment_stack`) | Re-verify live | 1:1 with staged; FDM ref LEFT JOIN attached; unmapped suppliers NOT filtered (commented gate) |

By company (staged, approximate — see `reconcile_cs_row_counts` for live split): FA largest, then WNSL, then NFC.

### Receipt (GRNI)

Receipt EIB is a strict subset of staged PO rows (those with raw Hard > 0.005 GBP):

| Stage | Rows | Notes |
|---|---:|---|
| Source: Staged PO (`src_fin_supplier_po_commitment_stack`) | Re-verify live | Same table as PO EIB |
| **After raw hard > 0.005 GBP gate** → `cte_grni` | Re-verify live | HARD-or-MIXED-class rows only (those carrying received-not-invoiced value); negative hard rows excluded by this gate |
| Workday EIB (`workday_supplier_po_commitment_stack_receipt`) | Re-verify live | One synthetic receipt document (`REC-` + order_number) per PO; lines ordered by (line_number, po_position) within each doc |

No row is ever dropped from the receipt EIB that wasn't already dropped from the PO EIB.
Because receipt rows are a subset of staged PO rows, the reference coverage (every
`column_bu` in receipt → `column_ez` in PO) is 100% by construction.

### Receipt EIB invariants

- Distinct receipt documents = COUNT(DISTINCT order_number) of the cte_grni set.
- `add_only = 'Y'` header rows (rnk = 1 within each receipt doc) = distinct receipt docs.
- `auto_complete = 'Y'` rows (same rnk = 1 logic) = same count.
- `submit = 'Y'` on **all** receipt lines (not just header). Flag for Workday consultant review.
- **0 dangling receipt references:** `SELECT COUNT(*) FROM workday_supplier_po_commitment_stack_receipt r WHERE NOT EXISTS (SELECT 1 FROM workday_supplier_po_commitment_stack p WHERE p.column_ez = r.column_bu)` is **always 0**.
- Over-receipt check (Σ receipt_extended per split ≤ PO_extended on same split): **always 0 violations** because receipt_extended = PO hard_local and PO_extended = soft_local + hard_local.

---

## Validated baseline (060826 extract / 040826 finance target)

Aggregates (GBP basis, per PO): **Finance £97,868,537 · Source £102,408,426 · EIB £97,850,875**

- **[A] EIB ↔ Source:** 6,971 POs both; EIB==Source penny-exact **6,961 (99.9%)**; in-source-not-EIB **2,802 POs / £4.53m** (business-gate filtered, legitimate); in-EIB-not-source **0**.
- **[B] Source(060826) ↔ Finance(040826):** identical **6,958 (99.8%)**; drifted **13 POs** (Σ|Δ| £34,924); finance-only 2; source-only 2,802 (same as gated set). The 13 drifts include the 5 accepted-exception POs above.
- **[C] EIB ↔ Finance headline:** EIB vs Finance **6,968 / 6,971 (100.0%)**; 3 non-matches, **ALL explained by [B] drift**; **0 transform residual**. (NFCPO300004848 £1,260 vs £2,280 is one of the 3 = drift, proven.)
- **Receipt side (hard-only):** Receipt EIB Σ = £26,487,648.24 local; penny-exact match to PO staged hard commitment total to the penny. Receipt clean-tie: **4,656 rows** penny-exact against Finance Hard£ target.

---

## Post-load tie-out (4th layer — sent EIB vs Workday-loaded control report)

**Status: COMPLETE / CLEAN (09.08.2026) — reconciled against BOTH the test-tenant control report and the PRODUCTION extract.** This is the OPEN-8 final leg — what we SENT vs what Workday actually LOADED.

- **Source [4] (loaded) — PRODUCTION (authoritative):** `~/Downloads/FA Find DC Purchase Orders PRODUCTION 9 August 2026 1.xlsx`, sheet `Sheet1`, 10,833 lines, all `Purchase Order Status = Issued`, **0** rows in `Validation Errors and Warnings`.
- **Source [4] (loaded) — earlier test-tenant control report:** `~/Downloads/FA FInd DC Purchase Orders.xlsx`, sheet `Control Reprot` (sic). Same structure and same totals; reconciled clean first. (May no longer be present locally.)
- **Source [3] (sent):** `workday_supplier_po_commitment_stack` (+ `_receipt`) in `rpatel.duckdb`, joined to `src_fin_supplier_po_commitment_stack` for transaction currency (the EIB stores local amounts as text).
- **Reproduce (production):** `uv run python scripts/generate_post_load_tieback.py --control-report "~/Downloads/FA Find DC Purchase Orders PRODUCTION 9 August 2026 1.xlsx" --label PRODUCTION` → writes `FA_PostLoad_Tieback_YYYYMMDD_PRODUCTION.xlsx` to `~/Documents/__data/excel/workday/reconcile/`. The script auto-resolves the sheet name (named sheet if present, else first). All figures computed live; no hardcoded totals.

Tie-out result — **identical on both extracts** (every delta £0.00 / 0 rows):

| Grain | Sent (EIB) | Loaded (control) | Δ |
|---|--:|--:|--:|
| PO service lines | 10,833 | 10,833 | 0 |
| Distinct POs | 6,971 | 6,971 | 0 |
| Receipt lines | 7,122 | 7,122 | 0 |
| Receipt docs | 4,659 | 4,659 | 0 |
| PO extended amount, per currency (GBP/USD/EUR/JPY/AUD/AED/CHF) | — | — | **£0.00 each** |
| Receipt approved amount, per currency | — | — | **£0.00 each** |
| Per-PO Σ extended (both / sent-only / loaded-only / amt-mismatch) | — | — | 6,971 / 0 / 0 / 0 |
| Per-PO **line-amount multiset** mismatches | — | — | **0** |
| Per-receipt-doc Σ mismatches | — | — | **0** |
| PO document date vs our `order_date` mismatches | — | — | **0** |
| **Per-PO recon** (Distrib Ext + Apprvd Recv, 6,971 rows) mismatches | — | — | **0** |
| **Per-Supplier recon** (Distrib Ext + Apprvd Recv, 1,980 rows) mismatches | — | — | **0** |

The tie-out workbook (8 sheets) adds two per-line-item recon grains beyond the headline: **Recon (per PO)** — every PO's Distribution Extended Amount + Approved Receipt Amount, sent vs loaded — and **Recon (per Supplier)** — rolled to (Company, Supplier legal-name, Currency).

**Join-key notes (why the tie is per-PO/multiset, not per-line-key):** the extract exposes Workday-renumbered `Line` refs (101/201/…), **not** our 3-part `service_order_line_id`, so lines are tied by per-PO amount **multiset** (0 mismatches = every line loaded at the right amount). The receipt is labelled `REC-{PO} for {PO}`; strip ` for …` to recover our `column_h` `REC-{PO}`.

**Per-Supplier recon normalisations** (needed only for the supplier grain — the extract carries supplier NAME, no supplier-ID column): (a) company label — sent-side `column_p` Business-Entity labels (`THE FA GROUP - FA/WNSL/NFC`) mapped to the extract's short labels (`Football Association` etc.); (b) manual Workday supplier-ID merge `S-008100 → S-004938` (NATIONAL FOOTBALL CENTRE — same legal entity, retired ID); (c) symmetric strip of trailing `(Inactive)/(Hold)/(Active)/(Blocked)` (NFCPO300012338's supplier shows `(Hold)` in the extract, `(Inactive)` in NRM). The production extract exercises (b)+(c) — exactly **1** `(Hold)` supplier row — and both resolve to **0** per-supplier mismatches. These are configured in `generate_post_load_tieback.py` (`SENT_SUPPLIER_ID_NORMALISATION`, `COMPANY_NAME_NORMALISATION`, status-suffix strip) and documented in the workbook Notes tab.

**3+ (loader) changes confirmed absorbed with zero reconciliation impact:** PO due-date blanked (`column_ae` NULL), receipt dates set, spreadsheet-key technical changes (did not break line identity), receipt-label rename (cosmetic).

---

## Reconciliation & validation

### Validators (`config/validate/finance/supplier_purchase_order_receipt/`, 12 files)

All currently cover the commitment-stack (PO) side. Receipt-side validators are open (§Outstanding).

**DO NOT re-order SQL files.** The table row order below matches `order.txt` in the validate folder — execute in this exact sequence (init first, then validators, then z-aggregator last):

1. `init.sql`
2. `validate_supplier_po_commitment_stack_required_keys.sql`
3. `validate_workday_extract_supplier_reference_id_conflict.sql`
4. `validate_supplier_po_commitment_stack_unmapped_suppliers.sql`
5. `validate_supplier_po_commitment_stack_unmapped_supplier_names.sql`
6. `validate_supplier_po_commitment_stack_tracked_spend_project.sql`
7. `validate_supplier_po_commitment_stack_missing_refs.sql`
8. `validate_supplier_po_commitment_stack_status_alignment.sql`
9. `validate_supplier_po_commitment_stack_line_id_uniqueness.sql`
10. `validate_workday_supplier_po_commitment_stack_mandatory_cols.sql`
11. `validate_supplier_po_commitment_stack_ref_coverage.sql`
12. `z_validate_supplier_purchase_order_receipt_issues.sql`

| File | What it checks |
|---|---|
| `init.sql` | Initialises `validation_supplier_purchase_order_receipt_issues` sink |
| `validate_supplier_po_commitment_stack_required_keys.sql` | Non-null keys: company, order_number, line_number, position |
| `validate_workday_extract_supplier_reference_id_conflict.sql` | Unnested Workday reference_id tokens map to unique Workday suppliers (MULTI_RECORD_TOKEN / MULTI_SUPPLIER_TOKEN detection) |
| `validate_supplier_po_commitment_stack_unmapped_suppliers.sql` | `wd_supplier_id IS NULL` rows (→ UNMAPPED_SUPPLIER issue bucket) |
| `validate_supplier_po_commitment_stack_unmapped_supplier_names.sql` | Supplier names without matches (→ MANDATORY_SUPPLIER bucket) |
| `validate_supplier_po_commitment_stack_tracked_spend_project.sql` | Tracked-spend lines missing a project mapping |
| `validate_supplier_po_commitment_stack_missing_refs.sql` | FDM mapping misses (→ MISSING_SPEND_CATEGORY etc.) |
| `validate_supplier_po_commitment_stack_status_alignment.sql` | Delivery/status field cross-consistency |
| `validate_supplier_po_commitment_stack_line_id_uniqueness.sql` | Uniqueness of `po_line_pos_id` (column_ez) in staged |
| `validate_workday_supplier_po_commitment_stack_mandatory_cols.sql` | Workday EIB mandatory columns are non-null |
| `validate_supplier_po_commitment_stack_ref_coverage.sql` | Coverage of FDM ref join |
| `z_validate_supplier_purchase_order_receipt_issues.sql` | Issues aggregation sink (currently empty stub) |

### Reconcile reports (`config/reconcile/finance/supplier_purchase_order_receipt/`, 16 files on disk / 11 in pipeline)

All `reconcile_cs_*` = commitment-stack (PO) side. Receipt-side reconcile files are open (§Outstanding).

**Valid reference-set rule for validate + reconcile SQL (MUST be followed):**
A validate or reconcile file is only allowed to SELECT/INSERT/JOIN from tables that fall into ONE of:
> (a) Any table materialised by the transform SQL folder (the 10 entries in `transform/sql/…/supplier_purchase_order_receipt/order.txt`)
> (b) Any of the 5 source extract tables declared in this module's ingest config:
>     `fin_supplier_po_commitment_stack_report`, `fin_supplier_po_receipt` (ingested but unused),
>     `fin_workday_extract_supplier`, `fin_ref_wd_fdm_mapping`, `fin_ref_wd_spend_category`
> (c) Future post-load Workday extracts that become the 4th-layer reconciliation TARGETs
>     — but ONLY AFTER they are added to the ingest config AND projected through a transform-layer
>     base SQL first (per §Open items [OPEN-8]).

**Cross-module external tables (e.g. `fin_supplier_po_audit_commitment`) are NEVER allowed in either folder's in-pipeline SQL.** If you ever need a new cross-module dependency, you MUST first:
1. Add the external extract to `supplier_purchase_order_receipt.json`; then
2. Build a transform-layer base/NRM projection for it; then
3. Add the reconcile/validate file that reads it.

This rule is written as a comment block into each folder's `order.txt`.

**Execution order.** `order.txt` has 11 in-pipeline entries (5 audit reconcile files are deliberately **off-pipeline** — do not add them back without the 3-step process above). Execute:

1. `init.sql`
2. `reconcile_cs_row_counts.sql`
3. `reconcile_cs_amount_by_currency.sql`
4. `reconcile_cs_currency_coverage.sql`
5. `reconcile_cs_unmapped_suppliers.sql`
6. `reconcile_cs_zero_commitments.sql`
7. `reconcile_cs_old_soft_cutoff.sql`
8. `reconcile_cs_commitment_class_summary.sql`
9. `reconcile_cs_finance_actions.sql`
10. `z_reconcile_cs_report.sql`
11. `zz_reconcile_cs_board_summary.sql`

| File | Purpose |
|---|---|
| `init.sql` | Init (stub) |
| `reconcile_cs_row_counts.sql` | Source→nrm→staged→workday row funnel by company. **2 actual row-drop buckets:** DROP_ZERO_COMMITMENT (zero total-commit rows) + DROP_OLD_SOFT (pre-2025-08-01 SOFT class). **1 informational bucket:** unmapped_supplier_rows — NOT dropped; they flow to Workday with NULL supplier_id (tracked via validation issues instead). Expected staged → Workday is therefore 1:1 with staged rows; unexplained_workday must be 0. |
| `reconcile_cs_amount_by_currency.sql` | Amounts by currency with split-aware exported basis. Same 2-drop + 1-info-bucket rule as row_counts. Expected staged → Workday amount is 1:1; unexplained_workday_amount must be £0.00. |
| `reconcile_cs_currency_coverage.sql` | Currency coverage profile |
| `reconcile_cs_unmapped_suppliers.sql` | Unmapped supplier detail (matches validate output structure). **Informational — NOT a drop bucket.** |
| `reconcile_cs_zero_commitments.sql` | Zero-commitment drop detail |
| `reconcile_cs_old_soft_cutoff.sql` | Old-soft cutoff drop detail |
| `reconcile_cs_commitment_class_summary.sql` | SOFT/HARD/MIXED/NONE summary by company |
| `reconcile_cs_audit_coverage.sql` | ⚠ **OFF-PIPELINE.** Fraser audit commitments vs Proactis — PO counts (BOTH / PROACTIS_ONLY / AUDIT_ONLY). Requires external `fin_supplier_po_audit_commitment` table (NOT in this module). Run ad-hoc only. |
| `reconcile_cs_audit_non_overlap.sql` | ⚠ **OFF-PIPELINE.** POs present in only one set with currency/reason. Same external dependency. |
| `reconcile_cs_audit_reconcile.sql` | ⚠ **OFF-PIPELINE.** All scenarios in one view — per-PO present_in, both totals, delta, reconcile_status, commitment_class_agreement. |
| `reconcile_cs_audit_only_triage.sql` | ⚠ **OFF-PIPELINE.** AUDIT_ONLY residual itemised with invoice/delivery status and suggested triage_bucket. |
| `reconcile_cs_audit_commitment_compare.sql` | ⚠ **OFF-PIPELINE.** Stack vs Fraser audit ledger at per-PO level, soft/hard commitment compare. |
| `reconcile_cs_finance_actions.sql` | Finance-actionable drops (unmapped suppliers + PO-level actions) |
| `z_reconcile_cs_report.sql` | Report aggregator |
| `zz_reconcile_cs_board_summary.sql` | Board-level summary (**FINAL** — source→sent legs computed live; post-load sent→loaded tie-out complete/clean, evidenced in §Post-load tie-out). Audit overlap was removed from the board summary to an inline footnote (audit coverage is off-pipeline and does not exist after the pipeline runs — use the 5 ad-hoc audit files directly if the audit table is loaded externally). |

### Audit ↔ Proactis gap

The gap is extract-scope, **not** a filter gap. Applying migration filters to both sides worsens non-overlap. AUDIT_ONLY (~2,224 POs) is ~96% HARD, `Part Invoiced`, `Fully Delivered` — received GRNI accruals the Fraser audit ledger holds but the Proactis *outstanding* extract never pulled. The reconcile is a coverage/exception control, not a zero tie-out. Real gaps are closed by widening extract scope (Proactis pull dates, invoice-status inclusion), not by transform changes.

---

## Open items

### [OPEN-1] Confirmations required before go-live (DO NOT SKIP)

- **C1 (Workday consultant):** confirm PO EIB `column_gf` posts a pure COMMITMENT (balance-sheet, no P&L). If it (unusually) posts an accrual posting, the full-gross vs. net decision needs revisiting. Currently: full-gross, on the basis PO = commitment = no P&L.
- **C2 (Finance):** sign off on Receipt EIB being a GRNI accrual = the commitment stack's `hard_commitment_value` (i.e. received-not-invoiced only). No separate Proactis receipt extract is used for load values.
- **C3 (Finance):** sign off on PO EIB totals at full gross (soft + hard, un-netted).
- **C4 (Ops / Workday):** load order enforced PO-before-receipt in Workday batch sequence. Workday receiving requires an issued PO to exist in the tenant.

### [OPEN-2] Mapping / data-quality sign-off (finance review)

The validation table `validation_supplier_purchase_order_receipt_issues` holds:
- 34 MANDATORY_SUPPLIER + 34 UNMAPPED_SUPPLIER  (~£41.7k exposed)
- 26 TRACKED_SPEND_NO_PROJECT
-  3 MISSING_SPEND_CATEGORY
-  1 STATUS_ALIGNMENT

These are finance-review items, not pipeline breaks. Supplier map resolution logic in `src_fin_workday_supplier_map.sql` (exact `vendor-company` first, bare `vendor` COALESCE second) covers all rows that CAN auto-resolve; the residual MUST be confirmed or extended by finance. Unmapped rows currently flow through both EIBs (the `wd_supplier_id IS NOT NULL` filter is commented in PO output, absent in receipt) so finance can see the impact in the EIB workbooks before sign-off.

### [OPEN-3] Legacy files cleanup (post go-live — low priority, after sign-off)

Once the commitment-stack-only pipeline is signed off and no forensic reference to the old Proactis-receipt-based design is needed:
1. DELETE the following files from the SQL folder (they are not in order.txt; safe to remove):
   - `src_fin_supplier_po_receipt_base.sql`
   - `src_fin_supplier_po_receipt_nrm.sql`
   - `src_fin_supplier_po_receipt.sql`
   - `workday_supplier_po_receipt.sql`
2. Remove the second ingest entry from `config/ingest/finance/supplier_purchase_order_receipt.json` (the PO Receipts Extract sheet → `fin_supplier_po_receipt`, position 2 in the ingest array) so the pipeline no longer ingests the unused receipt workbook.
3. Retire `scripts/generate_finance_review_commitment_receipts.py` (superseded by three-way pack).

### [OPEN-4] Receipt document date proxy

`workday_supplier_po_commitment_stack_receipt.sql` uses `commitment_date` for `column_y` (document_date) because the commitment stack source has no receipt event timestamp. **Current behaviour: clamped to `GREATEST(commitment_date, order_date)`** — a first Workday load rejected 92 POs (142 lines) with "receipt date is before the PO date" because `commitment_date < order_date` on those rows. The clamp keeps `commitment_date` where it is already ≥ the PO date and falls back to `order_date` otherwise. Still confirm with Workday consultant the preferred field:
- (a) `commitment_date`, clamped to ≥ `order_date` (**current**),
- (b) `order_date` for all receipts, or
- (c) a fixed cutover date for all migrated receipts.

The EIB template accepts any valid date, but Workday enforces receipt date ≥ PO date; the clamp satisfies that hard rule. Remaining choice is an accounting-period question, not a technical block.

### [OPEN-5] Free-text cleaning regime (applied 07.08.26, review after finance sign-off)

`src_fin_supplier_po_commitment_stack_base.sql` applies a **targeted** clean (NOT the 3-stage punctuation-strip originally designed):

1. **`line_description` only:** `SPLIT_PART(CAST(description AS VARCHAR), chr(9), 1)` → strips the TAB-bleeded Proactis continuation text, then `REGEXP_REPLACE` to swap C0 controls (0x00–0x1F) + DEL (0x7F) for space, then NULLIF to blank.
2. Other free-text cols: `TRIM` only. **No punctuation stripping was implemented** — `"'<>|;{}` and en/em dashes, smart quotes, etc. appear in ~270 legitimate descriptions and are valid XLSX cell content.
3. Join keys: `TRIM` only.

**CONFIRM on first Workday upload** whether any description/supplier columns are still rejected by the EIB parser; if so, extend the strip list conservatively (not by reintroducing the blanket punctuation strip).

### [OPEN-6] Mixed-type Description cast

The Description column in the Stack Report carries mixed cell types (INT, FLOAT, DATE, VARCHAR). Base layer CASTs to VARCHAR before any regex or TRIM operations to avoid type coercion explosions. **DO NOT remove this cast.**

### [OPEN-7] Receipt side validation & reconcile reports — reconciled 09.08.2026 (in-pipeline SQL deferred)

The receipt side is now **reconciled against the Workday-loaded control report** (§Post-load tie-out): 7,122 receipt lines / 4,659 docs, £0.00 delta in every currency, per-receipt-doc tie 4,659/4,659 with 0 amount mismatches. The finance sign-off obligation for receipt coverage is therefore met by the off-pipeline tie-out artifact.

The three *in-pipeline* receipt-side reconcile SQL files originally scoped here (row-funnel, amount-by-currency, cross-EIB reference) are **deferred, not required for sign-off** — they would duplicate, in-DB, checks the post-load tie-out already proves. Build them only if a repeatable in-pipeline receipt reconcile is later wanted (see recon-todo §D).

### [OPEN-8] Post-load Workday extract leg (three-way tie-out) — CLOSED 09.08.2026 (CLEAN)

**Resolved.** The Workday-loaded extract arrived — first the test-tenant control report, then the **PRODUCTION** extract (`~/Downloads/FA Find DC Purchase Orders PRODUCTION 9 August 2026 1.xlsx`, 10,833 lines) — and the sent→loaded tie-out is **penny-perfect and line-perfect on both**, including the per-PO (6,971) and per-Supplier (1,980) recon grains — see §Post-load tie-out. `zz_reconcile_cs_board_summary` banner flipped INTERIM → FINAL.

Scope taken: the light path (external tie-out artifact + docs FINAL), not the full in-pipeline A–E build in the recon-todo. Reason: the load is provably clean, and the control report does not expose the `service_order_line_id` the plan's §B assumed — it carries Workday-renumbered `Line` refs (101/201/…) and a `REC-{PO} for {PO}` receipt label — so the robust tie is per-PO / per-currency / per-PO line-amount **multiset**, computed live by `scripts/generate_post_load_tieback.py`. The in-pipeline A–E build (ingest the control report → transform base SQLs → extend the 6 `reconcile_cs_*` reports with a wdloaded leg) remains available in the recon-todo if a repeatable in-DB 4th leg is later wanted, but is not required for this one-time migration sign-off.

The validate/reconcile reference-set rule still stands for any future in-pipeline build: a post-load Workday extract MUST first pass through this module's ingest config → a transform-layer base SQL before any in-pipeline reconcile/validate SQL reads it. The tie-out script is **off-pipeline** (reads the control report xlsx directly), so it does not breach that rule.

---

## Gotchas

- **DuckDB lock.** `~/Documents/__data/duckdb/rpatel.duckdb` is locked when DBeaver is open. Close DBeaver or copy to scratch before running transform.
- **Blank values are often the string `'NULL'`**, not SQL NULL, throughout ingest tables. Test with: `UPPER(TRIM(col)) IN ('','NULL') OR col IS NULL`. Source examples: `order_closed_for_invoicing`, `order_closed_for_invoicing_reason`, `delivery_reason`.
- **`po_position` is never null** and `(company_shortcode, order_number, line_number, po_position)` is unique per logical split.
- **Date columns carry fractional timestamps** (e.g. `'2020-08-03 12:14:44.570000'`) AND Excel serial numbers (e.g. `44050.5`). The nrm layer always tries TIMESTAMP cast first, then falls back to DOUBLE→Excel serial math. Always use the nrm-parsed `order_date` / `commitment_date` rather than re-parsing raw text.
- **Round computed money to 2dp before comparing to zero.** Sub-penny float dust flips SOFT↔MIXED non-deterministically between reruns. Use `GREATEST(ROUND(value, 2), 0)` anywhere soft/hard are materialised. A real accrual is ≥ £0.01 so rounding only kills noise.
- **Receipt EIB document_number is synthetic.** `REC-{order_number}` is generated in `workday_supplier_po_commitment_stack_receipt.sql` (no correspondence to any Proactis receipt number). The comment column (`column_f`) carries `Source: Commitment Stack (GRNI)` for traceability. If Workday requires real (not synthetic) receipt doc numbers, source them from the (currently-unused) Proactis receipt extract via an additional LEFT JOIN of the staged row onto `fin_supplier_po_receipt` by (company, order, line, pos) — but note receipt_count per split can be >1 there and you'd need an aggregation policy.
- **060826 source file has ONLY `Sheet1`** (raw extract). **040826 file has `StackReport v1/v2/v3`** (finance analysis tabs). The three-way script defaults to 060826 for Source and 040826 for Finance; do NOT confuse the two extracts or it will silently fall back to the wrong file.
- **Description corruption on 5 accepted-exception POs:** the TAB-split + control-char clean in base.sql fixes the EIB load-safety (no misaligned rows from embedded CRs) but cannot restore the lost/shifted values — those are a source-extract problem accepted by finance. Re-running TRANSFORM is required after any base.sql edit to push the cleaned Description text into the published EIB Memo columns.
- **FX drifts identified between CSR and Exchange Rates tab** for USD, EUR, JPY, AUD, AED, and CHF. The three-way pack has a Local-vs-GBP FX sheet surfacing these. Not a transform bug; it's a Proactis vs authority rate drift.
- **Pipeline RuntimeError on Transform phase** = validation- or transform-abort trigger. The pipeline raises `RuntimeError` before publish if safety checks fail. Read the transform phase output (not just the stacktrace) for the specific abort reason.
- **Stakeholder-facing Excel artifacts must NEVER be written to the repository.** Stage via `/tmp` then `cp` to `/Users/Rakesh.Patel/Documents/__data/excel/workday/reconcile/` (or another user-owned path) to bypass macOS sandbox restrictions on ~/Downloads.

---

## Verification & run-book

Run commands are in [CLI.md](CLI.md). Reproduce from a fresh session:

### Step 1 — Run TRANSFORM + PUBLISH

```bash
# 0. Close DBeaver (releases DuckDB write lock on rpatel.duckdb)

# 1. From a NON-SANDBOXED shell in repo root:
cd elt_ingest_excel
uv run python examples/fin_supplier_purchase_order_receipt.py \
     --run-to-phase TRANSFORM
# Expected: no errors. RuntimeError on Transform phase -> check the transform
# phase output for the specific abort reason (pipeline aborts before publish
# if validation fails).

# 2. After TRANSFORM green, run PUBLISH from the same shell:
uv run python examples/fin_supplier_purchase_order_receipt.py \
     --run-to-phase PUBLISH
# Two Excel files are written to ~/Documents:
#   workday_fin_supplier_po_commitment_stack.xlsx
#   workday_fin_supplier_po_commitment_stack_receipt.xlsx
```

### Step 2 — Quick gate + reconcile snapshot (read-only, safe from TRAE sandbox)

```bash
cd elt_ingest_excel
uv run python - <<'PY'
import duckdb, os
c = duckdb.connect(os.path.expanduser("~/Documents/__data/duckdb/rpatel.duckdb"), read_only=True)
q = lambda s: c.execute(s).fetchone()[0]

# Row counts (live)
for t in ("src_fin_supplier_po_commitment_stack_base","src_fin_supplier_po_commitment_stack_nrm",
          "src_fin_supplier_po_commitment_stack","src_fin_supplier_po_commitment_stack_ref",
          "workday_supplier_po_commitment_stack","workday_supplier_po_commitment_stack_receipt"):
    print(t, q("SELECT COUNT(*) FROM "+t))

# PO line key uniqueness (staged and EIB both must be unique)
print("po_line_pos_id_unique_staged:",
      q("SELECT COUNT(*) FROM src_fin_supplier_po_commitment_stack")
      == q("SELECT COUNT(DISTINCT po_line_pos_id) FROM src_fin_supplier_po_commitment_stack"))
print("po_column_ez_unique_eib:",
      q("SELECT COUNT(*) FROM workday_supplier_po_commitment_stack")
      == q("SELECT COUNT(DISTINCT column_ez) FROM workday_supplier_po_commitment_stack"))

# Cross-EIB reference check (0 dangling = correct)
print("receipt_dangling_refs:", q("""
    SELECT COUNT(*) FROM workday_supplier_po_commitment_stack_receipt r
     WHERE NOT EXISTS (SELECT 1 FROM workday_supplier_po_commitment_stack p
                        WHERE p.column_ez = r.column_bu)"""))

# Over-receipt check (0 = correct)
print("over_receipt_rows:", q("""
    SELECT COUNT(*) FROM (
      SELECT p.column_ez, p.column_gf po_amt, r.column_bz rec_amt
        FROM workday_supplier_po_commitment_stack p
        JOIN workday_supplier_po_commitment_stack_receipt r ON p.column_ez = r.column_bu
       WHERE TRY_CAST(p.column_gf AS DECIMAL(18,2)) < TRY_CAST(r.column_bz AS DECIMAL(18,2))
    ) x"""))

# Receipt EIB invariants
print("receipt_add_only_Y:     ", q("SELECT COUNT(*) FROM workday_supplier_po_commitment_stack_receipt WHERE column_c='Y'"))
print("receipt_auto_complete_Y:", q("SELECT COUNT(*) FROM workday_supplier_po_commitment_stack_receipt WHERE column_e='Y'"))
print("receipt_distinct_docs:  ", q("SELECT COUNT(DISTINCT column_d) FROM workday_supplier_po_commitment_stack_receipt"))

# Totals (GBP columns from staged; EIB is local so compare via staged only)
print("PO_staged_soft+hard_total_gbp:", q("""
    SELECT ROUND(SUM(COALESCE(nrm_soft_remaining_gbp,0)+COALESCE(nrm_hard_remaining_gbp,0)),2)
      FROM src_fin_supplier_po_commitment_stack"""))
print("PO_staged_hard_total_gbp  (= receipt EIB if all rows were GBP):", q("""
    SELECT ROUND(SUM(COALESCE(nrm_hard_remaining_gbp,0)),2)
      FROM src_fin_supplier_po_commitment_stack"""))

# Reconcile summary
print(c.execute("SELECT * FROM reconcile_cs_row_counts ORDER BY company").fetchall())
print(c.execute("SELECT bucket,COUNT(*) FROM reconcile_cs_audit_reconcile GROUP BY 1 ORDER BY 1").fetchall())
c.close()
PY
```

### Step 3 — Generate the finance three-way review pack

```bash
cd elt_ingest_excel
uv run python scripts/generate_finance_review_three_way.py
# Writes FA_ThreeWay_Reconciliation_YYYYMMDD.xlsx to:
#   ~/Documents/__data/excel/workday/reconcile/
```

### Step 4 — Before committing SQL/doc edits

Re-run TRANSFORM + the spot-checks above to confirm counts and invariants are unchanged
(unless the edit INTENDS to change them, in which case record the new numbers in the
Status block above).

---

## Sandbox limitations

- **Do NOT try `duckdb.connect(rpatel, read_only=False)` from the TRAE sandbox.** It always fails with `Operation not permitted` even when DBeaver is closed. Run TRANSFORM or PUBLISH phases from an actual user shell.
- **Read-only SELECT queries and COUNT(*) checks from the sandbox are fine.**
- **Excel artifacts to reconciliation directory:** Stage via `/tmp` then `cp` to `/Users/Rakesh.Patel/Documents/__data/excel/workday/reconcile/` (or another user-owned path) to bypass macOS sandbox restrictions on ~/Downloads.

---

## Where things live

| Artefact | Path |
|---|---|
| Ingest config | `config/ingest/finance/supplier_purchase_order_receipt.json` (entry for PO Receipts Extract is UNUSED in output chain) |
| Transform SQL | `config/transform/sql/finance/supplier_purchase_order_receipt/` |
| Execution order | `config/transform/sql/finance/supplier_purchase_order_receipt/order.txt` (10 files; legacy receipt files NOT included) |
| PO base (footer filter + TAB-split Description) | `src_fin_supplier_po_commitment_stack_base.sql` → drops END OF FILE rows; only Description gets TAB-split + C0/DEL strip |
| PO nrm (supplier map + commitment class + two vendor overrides) | `src_fin_supplier_po_commitment_stack_nrm.sql` → `STE096→STE090` override (FAPO200025244), `NAT008→NAT015` override (NFC), `commitment_class`, `wd_supplier_id`, `order_date`, `commitment_date`, `nrm_company_id` |
| PO business gates + full-gross staged output (GREATEST(0,x) wrap) | `src_fin_supplier_po_commitment_stack.sql` → 4 pre-gates + 0.005 gate → `nrm_{soft,hard}_remaining_{gbp,local}`, `nrm_po_export_local`, `po_line_pos_id` |
| PO FDM / cost-centre ref join + reporting_type | `src_fin_supplier_po_commitment_stack_ref.sql` |
| PO EIB (full gross; JPY integer; submit=all; supplier filter commented) | `workday_supplier_po_commitment_stack.sql` → `column_gf = nrm_po_export_local = soft_local + hard_local`; `column_ez = po_line_pos_id`; line 81 `wd_supplier_id IS NOT NULL` commented |
| **Receipt EIB (GRNI from Hard commitment; submit=all lines)** | **`workday_supplier_po_commitment_stack_receipt.sql`** → reads from staged PO WHERE raw hard_gbp > 0.005; `column_bz = hard_local` (JPY INT, else 2dp); synthetic `REC-{order_number}` receipt doc; `column_bu = po_line_pos_id` (PO link); `column_y = GREATEST(commitment_date, order_date)` proxy (clamped ≥ PO date) |
| PO EIB mandatory / line-id validators | `validate_workday_supplier_po_commitment_stack_mandatory_cols.sql`, `validate_supplier_po_commitment_stack_line_id_uniqueness.sql` |
| Supplier map resolution (pipe-unnested reference_ids) | `src_fin_workday_supplier_map.sql` (exact `vendor-company` first, bare `vendor` COALESCE second) |
| Unresolved suppliers reconcile | `reconcile_supplier_po_unresolved*.sql` (3 files, PO side) |
| PO reconcile reports (16 files on disk / 11 in pipeline) | `config/reconcile/finance/supplier_purchase_order_receipt/reconcile_cs_*.sql` (commitment-stack side; receipt-side TBD post go-live; 5 audit reconcile files are OFF-PIPELINE / ad-hoc only because they depend on `fin_supplier_po_audit_commitment` not in this module) |
| Validate/reconcile reference-set rule (enforced in order.txt footers) | In-pipeline SQL must only reference: transform-layer tables OR ingest-config-declared source extracts OR post-load Workday extracts that have passed through ingest+transform first. Cross-module externals → off-pipeline only. See §Reconciliation & validation. |
| OPEN-8 single-item build-plan todo | `config/transform/sql/finance/supplier_purchase_order_receipt/recon-todo.md` (A→E steps, 7 acceptance criteria; read before adding the 4th reconcile layer) |
| Three-way finance review pack generator | `scripts/generate_finance_review_three_way.py` → 7-sheet Excel, output to reconcile dir |
| Post-load tie-out generator (4th layer, OPEN-8) | `scripts/generate_post_load_tieback.py` → reads sent EIB + a Workday-loaded extract, writes `FA_PostLoad_Tieback_YYYYMMDD[_LABEL].xlsx` (8 sheets incl. per-PO + per-Supplier recon) to reconcile dir. Args: `--control-report`, `--sheet` (auto), `--label`. All live; no hardcoded totals. See §Post-load tie-out |
| Workday-loaded extract — PRODUCTION (input, outside repo) | `~/Downloads/FA Find DC Purchase Orders PRODUCTION 9 August 2026 1.xlsx` (sheet `Sheet1`) — the authoritative post-load [4] source |
| Workday-loaded extract — earlier test control report | `~/Downloads/FA FInd DC Purchase Orders.xlsx` (sheet `Control Reprot`) — same totals; may no longer be present locally |
| Publish config (2 EIBs) | `config/publish/finance/publish_supplier_purchase_order_receipt.json` → PO + Receipt EIBs (both named with `_commitment_stack` infix) |
| EIB templates (outside repo) | `~/Documents/__data/excel/finance_ref/eib/eib_template/Submit_Purchase_Order_v46.0.xlsx`, `Submit_Receipt_v46.0.xlsx` |
| Reconcile output dir (outside repo) | `~/Documents/__data/excel/workday/reconcile/` |
| Finance target (outside repo) | `~/Downloads/The FA - Commitment Stack Report as at 040826 USE THIS ONE.xlsx` (StackReport v3 GBP cols + Variance col B) |
| Source extract (outside repo) | `~/Documents/__data/excel/finance_ref/supplier_purchase_order/prod/The FA - Commitment Stack Report as at 060826.xlsx` (Sheet1 only) |
| Run commands | [CLI.md](CLI.md) |
