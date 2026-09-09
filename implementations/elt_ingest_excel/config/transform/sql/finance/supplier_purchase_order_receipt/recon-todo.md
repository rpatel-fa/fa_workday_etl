# Recon Rebuild — Remaining Work (Post-Load Workday Extract, 4th Layer)

> **✅ RESOLVED 09.08.2026 — post-load tie-out COMPLETE / CLEAN.**
> The Workday control report (`~/Downloads/FA FInd DC Purchase Orders.xlsx`) reconciles
> penny-perfect and line-perfect against the sent EIBs. Evidence + full numbers:
> `docs/finance-supplier-po-receipt.md` §Post-load tie-out (and OPEN-8, now CLOSED).
> Regenerate: `uv run python scripts/generate_post_load_tieback.py`.
>
> **Path taken:** the light option — off-pipeline tie-out script + artifact + docs FINAL,
> **not** the full in-pipeline A–E build below. Two reasons: (1) the load is provably clean,
> so an in-DB 4th leg adds no sign-off value; (2) the control report does not expose the
> `service_order_line_id` that §B assumed — it carries Workday-renumbered `Line` refs
> (101/201/…) and a `REC-{PO} for {PO}` receipt label — so the robust tie is per-PO /
> per-currency / per-PO line-amount **multiset**, which the script does directly.
>
> The A–E plan below is **retained as-is** for reference only, in case a repeatable
> in-pipeline 4th leg is later wanted. Nothing below is outstanding for sign-off.

---

**This file is the focused TODO for the ONE outstanding piece of the reconciliation architecture.**
All historical design decisions, accepted exceptions, validated baseline numbers, constraints,
and the source/sent-leg implementation live in the single authority:

```
elt_ingest_excel/docs/finance-supplier-po-receipt.md
```

Read the main doc first for context. Then come back here for the single remaining item
that must be executed **after the PO and receipt EIBs are loaded into Workday and a
Workday PO/receipt extract is available locally**.

Do not add any other items to this file — everything else from the former recon-todo.md
has been delivered (the three-way pack, the 7-sheet review workbook, the drift and
exception analysis, the accepted-exception sign-off table, and the source-leg reconcile
reports).

---

## Single Outstanding Item: Add the 4th Layer (Workday-Loaded Extract → Sent → Source → Finance)

Today the three-way pack and all `reconcile_cs_*` SQL reports cover only:

- **[1] Source (060826 Proactis Commitment Stack extract)** — what we ingested
- **[2] EIB sent payload (DuckDB workday_* tables / published xlsx)** — what we sent to Workday
- **[3] Finance target (040826 StackReport v3 + Variance col B)** — finance's authority

That is a **pre-load control**. The final sign-off reconciliation needs a **4th layer**:

- **[4] Workday-loaded extract** — the POs and receipts actually sitting in the Workday tenant *after* EIB load

We need to extend every reconcile report and the three-way Excel pack so every comparison
adds leg [4]. This closes the "source → sent → loaded" cold-start loop that OPEN-8 in the
main doc describes.

### When to do this

1. EIB load completed in Workday tenant (PO first, then receipt — load order matters).
2. Workday extract produced for:
   - Purchase Orders (all migrated POs — use the EIB `document_number` range / load-date filter)
   - Receipts (all migrated receipts — use the synthetic `REC-` prefix on `receipt_number` filter)
3. Extract files are placed under `~/Documents/__data/excel/workday/extract/` (or another
   user-owned path; update ingest config accordingly).

### What needs changing — concrete TODO list

#### A. Ingest layer: add the Workday extract files

- Add two new entries to the ingest config
  [`config/ingest/finance/supplier_purchase_order_receipt.json`](../../../ingest/finance/supplier_purchase_order_receipt.json):
  - Workday PO extract → e.g. `wd_extract_purchase_order`
  - Workday Receipt extract → e.g. `wd_extract_receipt`
- Keep the existing 060826 + legacy receipt entries; do **not** remove anything yet (post-go-live
  cleanup is a separate OPEN-3 item).

#### B. Transform SQL folder: add two new base/nrm projection files for the WD extracts

Create two new SQL files (follow the `cte_src_*_base` → `cte_joined` → `SELECT *` pattern used
in every nrm layer in this folder):

- `src_fin_wd_extract_po_base.sql` — semantic aliases, dual text/numeric pairs for value cols,
  TRIM-only on keys, CAST to VARCHAR first for any mixed-type description cell (same pattern as
  `src_fin_supplier_po_commitment_stack_base.sql`). Project the exact fields needed to join on to
  our EIB payload: Workday PO `document_number` = our EIB `column_j order_number`; Workday
  `service_order_line_id` (or the tenant's equivalent) to our EIB `column_ez po_line_pos_id`;
  Workday extended amount (transaction currency + GBP equivalent if provided); Workday supplier
  ref; Workday company ref; load/version timestamp if present.
- `src_fin_wd_extract_receipt_base.sql` — same pattern for receipt extract. Join key is Workday
  `receipt_number` (matches our synthetic `REC-{order_number}`) AND Workday receipt service line
  (matches our `po_line_pos_id`). Project extended amount (transaction currency = our EIB `column_bz`);
  Workday PO line linkage ref (must resolve back to Workday PO service line → our EIB `column_ez`).

Add these two files **after** the existing chain in [`order.txt`](./order.txt):
```
11. src_fin_wd_extract_po_base.sql
12. src_fin_wd_extract_receipt_base.sql
```
No need for a separate nrm layer unless the Workday extract needs supplier remapping or date
parsing beyond base — keep it lean; the comparison SQL does the JOIN work.

#### C. Reconcile SQL reports (16 `reconcile_cs_*` files) — extend each with a [4] leg

For every existing reconcile SQL that today compares [1]↔[2]↔[3] (or subsets), extend the SELECT
so each aggregate bucket also carries:
- sent_amt (already present = our EIB)
- wdloaded_amt (NEW = from the Workday extract projections in §B)
- delta_sent_vs_wdloaded = sent_amt − wdloaded_amt (should be £0 per-bucket if load is perfect)
- load_status flag (BOTH / SENT_ONLY / WDLOADED_ONLY) on the per-PO comparison reports

Specific files that MUST get the 4th leg (finance cares about these):
1. `reconcile_cs_row_counts.sql` — add `wdloaded_rows` column alongside `sent_rows`.
2. `reconcile_cs_amount_by_currency.sql` — add `wdloaded_soft_gbp`, `wdloaded_hard_gbp`,
   `wdloaded_soft_local`, `wdloaded_hard_local`; add `delta_sent_vs_wdloaded_*`.
3. `reconcile_cs_audit_reconcile.sql` — extend the present_in bitmap from 3-way → 4-way
   (add `WDLOADED` bit).
4. `reconcile_cs_audit_commitment_compare.sql` — add wdloaded-side soft/hard to the
   per-PO comparison grid.
5. `reconcile_cs_commitment_class_summary.sql` — extend each SOFT/HARD/MIXED/NONE bucket
   with wdloaded GBP + rowcount.
6. `reconcile_cs_finance_actions.sql` — add `SENT_NOT_IN_WDLOADED` and
   `WDLOADED_NOT_IN_SENT` triage buckets alongside the existing unmapped-supplier and
   PO-level action buckets.

Files that can stay as-is (pure source/validate control, no load-leg needed):
- `reconcile_cs_unmapped_suppliers.sql`, `reconcile_cs_zero_commitments.sql`,
  `reconcile_cs_old_soft_cutoff.sql`, `reconcile_cs_currency_coverage.sql`,
  `reconcile_cs_audit_coverage.sql`, `reconcile_cs_audit_non_overlap.sql`,
  `reconcile_cs_audit_only_triage.sql`.

Finally flip `zz_reconcile_cs_board_summary.sql` from `INTERIM` → `FINAL` (it has a banner
comment calling this out in the main doc OPEN-8). Replace the placeholder text that says
"post-load Workday extract arrives" with the live 4-leg totals.

#### D. Receipt-side reconcile coverage (OPEN-7 in main doc, same delivery batch)

The 16 reconcile_cs_* files are commitment-stack (PO) side only. The post-load batch is a
good moment to add receipt-side reconcile SQLs (OPEN-7 in main doc), since the Workday
receipt extract projection (§B) is already in hand. Add three new files (create them; the
pattern is obvious from the PO-side files — swap PO grain for Receipt grain, hard-only):

- `reconcile_rc_row_counts.sql` — receipt funnel: staged-hard-rows → our receipt EIB →
  WD-receipt-extract, with the 0.005-GBP gate bucketed and `SENT_NOT_IN_WDLOADED` /
  `WDLOADED_NOT_IN_SENT` counts.
- `reconcile_rc_amount_by_currency.sql` — paired with PO amount-by-currency so finance can
  eyeball Σ receipt ≤ Σ hard per currency on both sent and wdloaded legs.
- `reconcile_rc_cross_eib_reference.sql` — every `purchase_order_line` (receipt side) resolves
  to a live Workday PO service line. By construction 0 rows on our side; surface the
  WD-loaded-side equivalent as a zero-row sanity report for the review pack.

Add them after the existing unresolved suppliers in [`order.txt`](./order.txt) if you want them
auto-executed by the pipeline, or keep them as ad-hoc scripts (the 16 reconcile_cs files are
all run during TRANSFORM, so adding new reconcile_* files to order.txt is the standard path).

#### E. Generate-finance-review-three-way.py — add the 4th leg and a new sheet

Extend [`scripts/generate_finance_review_three_way.py`](../../../../../scripts/generate_finance_review_three_way.py)
(currently 7 sheets):
- Add two new data sources: wd_po (Workday PO extract projection) and wd_rc (Workday receipt
  extract projection).
- **Sheet 8 — Post-Load Tieback (new):** 4-way per-PO GBP grid: finance | source | sent | wdloaded.
  Bucket: ALL4 / FINANCE_ONLY / SOURCE_ONLY / SENT_ONLY / WDLOADED_ONLY / DRIFT_3WAY_MATCH_WDLOAD /
  DRIFT_WDLOAD_MISMATCH. Flag every sent↔wdloaded delta (any > £0.01) — these are *load anomalies*
  (not transform drift) and must be zero for sign-off.
- **Sheet 9 — Receipt Post-Load Tieback (new if OPEN-7 done):** 4-way per-receipt-doc hard£ grid.
  Same bucket semantics on the hard-only subset. If OPEN-7 isn't done yet, add a placeholder sheet
  with a banner saying "Receipt post-load leg deferred — see OPEN-7" so the workbook is complete.
- Update **Sheet 1 (Exec Summary)** to carry the 4-leg headline totals and highlight the
  sent↔wdloaded delta (target: £0).
- Do NOT re-introduce hardcoded totals (constraint #4 in the main doc invariants). Read all four
  legs from live tables and files.
- Keep output dir the same: `~/Documents/__data/excel/workday/reconcile/`, filename stays
  `FA_ThreeWay_Reconciliation_YYYYMMDD.xlsx` (now effectively 4-way; the name is fine — rename
  only if finance asks).

### Acceptance criteria for closing this TODO

- Running TRANSFORM produces valid `src_fin_wd_extract_po_base` and `src_fin_wd_extract_receipt_base`
  tables with non-zero, plausible row counts (expected ~ our EIB counts, ± known exclusions).
- `reconcile_cs_row_counts.sql` live output carries wdloaded_rows column; row counts on
  sent ↔ wdloaded match to within known Workday-side auto-suppressed rows (if any — if the counts
  are not penny-exact, investigate before marking this done).
- `reconcile_cs_amount_by_currency.sql` sent vs wdloaded is **£0.00 delta per currency** for the
  top 8 currencies (GBP, USD, EUR, JPY, AUD, AED, CHF, CAD). Any non-zero delta is triaged into
  `reconcile_cs_finance_actions.sql` SENT_NOT_IN_WDLOADED / WDLOADED_NOT_IN_SENT buckets with a
  written explanation.
- The new 4-way Exec Summary shows headline Finance / Source / Sent / WDLoaded aggregates with
  sent↔wdloaded delta = £0.00 on PO total GBP.
- `zz_reconcile_cs_board_summary.sql` banner reads **FINAL** (not INTERIM).
- No new hardcoded totals anywhere.
- Cold-start session can reproduce the 4-way pack end-to-end by running:
  1. ingest (new Workday extract entries),
  2. TRANSFORM (extended order.txt),
  3. `uv run python scripts/generate_finance_review_three_way.py`,
  and get a valid 9-sheet workbook with non-empty post-load legs.

### Cross-references to main doc

- Main authority: [`finance-supplier-po-receipt.md`](../../../../../docs/finance-supplier-po-receipt.md)
- OPEN-7 (Receipt-side validation & reconcile reports) — same batch as §D above.
- OPEN-8 (Post-load Workday extract leg — three-way tie-out) — §C + §E above.
- OPEN-3 (Legacy files cleanup) — separate, later, after this TODO is signed off.
- OPEN-1 C1–C4 (Finance & Workday consultant confirmations) — must already be done *before*
  Workday load; if they are still open, block this TODO until sign-off is in writing.

---

## Nothing else lives here

This file deliberately has no other sections. All historical context lives in the main doc.
If you find yourself wanting to add broader design notes, accepted exceptions, baseline numbers,
or gotchas here — add them to the main doc instead, not this file.
