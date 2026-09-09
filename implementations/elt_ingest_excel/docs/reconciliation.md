# Reconciliation (source → target data integrity)

Reconciliation answers a different question from validation. **Validation** checks whether the
data *in* the final `workday_*` tables meets Workday's data-quality rules (valid country, valid
format, required fields present). **Reconciliation** checks whether the transform *lost, dropped,
or corrupted* anything on the way from the raw GP source extract to those final tables — GP IDs,
addresses, bank accounts, emails, phone numbers, payment terms, tax records.

Both are needed for a business sign-off: validation proves the output is well-formed;
reconciliation proves nothing fell out of the pipeline between source and target.

Reconciliation SQL lives in `config/reconcile/finance/{customer,supplier}/`, the CCS-specific
`customer_ccs`/`supplier_ccs` packs, and the PO/receipt pack in
`config/reconcile/finance/supplier_purchase_order_receipt/`, and is run the same way as validation —
every `.sql` file in the folder executes in alphabetical order via
`run_validation_sql_dir()` (the runner is generic over any SQL directory; nothing
reconciliation-specific lives in the Python layer). `init.sql` runs first and creates a fresh
`reconciliation_{customer,supplier}_issues` table.

## Prerequisite: run validation first

The consolidated report's GP ID section reads `validation_{customer,supplier}_issues`, which is
only populated by `validate_{customer,supplier}_id_coverage.sql`. Run validation before
reconciliation, or the report's GP ID rows will be empty/stale:

```bash
cd elt_ingest_excel
uv run python examples/fin_customer_debtor_validate_sql.py
uv run python examples/fin_customer_debtor_reconcile_sql.py
uv run python examples/fin_supplier_creditor_validate_sql.py
uv run python examples/fin_supplier_creditor_reconcile_sql.py
```

Both accept `--database-path` / `--runtime-config` overrides, same as the validate scripts.

## CCS customer reconciliation status

The separate CCS customer workstream is now complete. It lives under
`config/reconcile/finance/customer_ccs/` and reconciles the loaded Workday customer extract
`src_fin_workday_extract_ccs_customer`, not just the generic customer EIB output tables.

Its contract is deliberately two-layered:

- **Primary lens:** prove whether the loaded extract preserves the sent `workday_customer_*`
  payload.
- **Secondary lens:** explain GP source vs loaded extract differences in the context of approved
  cleansing, normalization, and dedup performed upstream in `config/transform/sql/finance/customer/`.

For CCS customer, the gold expected value is the sent `workday_customer_*` transform output, not
the raw GP row.

Main objects to review:

- board summary: `reconcile_customer_ccs_board_summary`
- summary report: `reconcile_customer_ccs_report`
- closure summary: `reconcile_customer_ccs_closure_summary`
- decision pack: `reconcile_customer_ccs_closure_decision_pack`
- analyst workbook:
  `~/Documents/__data/excel/workday/reconcile/customer_ccs_reconcile_review_pack.xlsx`

The workbook's first tab is `board_summary`, a flat board-level rollup (one row per headline
item: `summary_order, section, item, value, disposition, note`). Every value is pulled live from
the reconcile tables above, so it cannot drift from the detail. `disposition` is the actionable
column: filter `FIX_IN_WORKDAY` for the Finance action list, `OPERATOR_DECISION` for open
decisions; `RECONCILED` / `ACCEPTED` / `INFO` / `NOT_SUPPORTED` need no action. Built by
`zz_reconcile_customer_ccs_board_summary.sql` (the `zz_` prefix makes it run after the `z_`
report it reads from).

Note: the customer pack's rich first `summary` tab (in the delivered
`customer_ccs_reconcile_review_pack.xlsx`) was hand-authored and has no committed generator — so
that pack is not currently regenerable as-is from code. Making it deterministic (like the supplier
`summary` tab) is tracked in `docs/todo/todo_finance_customer_ccs_summary_refactor.md`.

Refresh commands:

```bash
cd elt_ingest_excel

# Refresh CCS extract + reconcile
uv run python examples/fin_customer_ccs_reconcile_sql.py --run-to-phase TRANSFORM

# Rebuild the analyst workbook from the current DuckDB reconciliation state
uv run python examples/fin_customer_ccs_reconcile_sql.py --export-review-workbook-only
```

## CCS supplier reconciliation status

The supplier analogue of the CCS customer workstream lives under
`config/reconcile/finance/supplier_ccs/` and reconciles the loaded Workday supplier extract
(`fin_workday_extract_supplier`, refreshed from the current `Extract Suppliers … .xlsx`) against
the sent `workday_supplier_*` transform output and the GP source. It is run by
`examples/fin_supplier_ccs_reconcile_sql.py`, which also loads the go-live manual-update list
(`Suppliers to be manually updated at Go Live.xlsx`) into `fin_supplier_manual_update`.

Key differences from the customer pack (supplier domain):

- **Join key is the GP Vendor ID in its business-unit-suffixed form** (`nrm_supplier_number`, e.g.
  `LON004-FA`), taken from `src_fin_supplier`. The raw `src_fin_supplier_base.vendor_id` (`LON004`)
  is reused across FA/NFC/WNSL and is NOT the reconcile key. The extract's `supplier_id` (`S-…`) is
  a fresh tenant allocation and is never a cross-run key.
- **Known-exception classification.** The manual-update list is name-keyed only; it is matched to
  the extract on normalised supplier name to recover the reference_id, then address/bank/payment
  gaps are classified `EXPECTED` (on the list) vs `UNEXPLAINED` (genuine residual). "Removed bank"
  also forces payment to Manual, so a payment gap is `EXPECTED` if the supplier is on the bank OR
  payment bucket.
- **Address presence** is tested across all extract address columns (`primary_address`,
  `primary_remit_to_address`, …) — CH/IT suppliers store the address as a Remit-To address, so
  checking `primary_address` alone over-reports gaps.

Main objects to review:

- **Finance worklist (the one tab to action): `reconcile_supplier_ccs_finance_action_list`** — every
  open item to fix in Workday in one place: manual updates still open (extract-verified), residual
  address gaps, and the test row to remove. This is the tab Finance works from.
- board summary: `reconcile_supplier_ccs_board_summary`
- summary report: `reconcile_supplier_ccs_report`
- closure summary: `reconcile_supplier_ccs_closure_summary`
- classified gaps: `reconcile_supplier_ccs_attribute_gap`
- manual-update list (with `still_open_in_workday`): `reconcile_supplier_ccs_manual_update_status`
- analyst workbook:
  `~/Documents/__data/excel/workday/reconcile/supplier_ccs_reconcile_review_pack.xlsx`

The pack's first tab is an 8-section executive `summary` (mirroring the customer pack's format:
population, core attributes, the go-live manual-update list, attribute gaps, actions, accepted
items, status, approvals). It is built by `examples/supplier_ccs_summary_sheet.py` and prepended
to the workbook automatically on every export; every number is read live from the reconcile
tables so the sheet cannot drift from the detail.

```bash
cd elt_ingest_excel

# Refresh supplier extract + manual list + reconcile (add --export-review-workbook to write the pack)
uv run python examples/fin_supplier_ccs_reconcile_sql.py --run-to-phase TRANSFORM --export-review-workbook

# Rebuild the analyst workbook from the current DuckDB reconciliation state
uv run python examples/fin_supplier_ccs_reconcile_sql.py --export-review-workbook-only
```

## PO/receipt reconciliation (source → target EIB)

The supplier PO/receipt reconcile and validation — source→target contract, gates, value reconciles,
the currency basis, the Fraser audit population reconcile, and the review pack — are documented in
the single canonical doc: **[finance-supplier-po-receipt.md](finance-supplier-po-receipt.md)**
(§Reconciliation & validation).
## The one table to review: `reconcile_{customer,supplier}_report`

Everything else feeds into this single consolidated table. Query it directly:

```sql
SELECT * FROM reconcile_customer_report ORDER BY section, metric;
SELECT * FROM reconcile_supplier_report ORDER BY section, metric;
```

It has four columns: `section`, `metric`, `metric_value`, `status`. Every metric is a
numerator/denominator pair (source count vs. target count, or found vs. not-found), never a bare
pass/fail flag — so a reviewer can verify retention by eye rather than trust a label.

Sections (numbered so they sort in reading order):

1. **`1_gp_id_reconciliation`** — distinct GP ID count in source vs. traced to target, the
   `validate_*_id_coverage.sql` issue count (should be 0), and a full row-level found/not-found
   tally from the GP ID proof table.
2. **`2_source_row_funnel`** — row counts at every pipeline stage: raw ingest → union → normalised
   → distinct GP ID → canonical dedup groups → final master table. Includes a defensive
   `UNEXPLAINED_DROP` check (in `reconcile_*_row_funnel.sql`) that fires if a future undocumented
   filter appears anywhere in that chain.
3. **`3_attribute_coverage`** — one count-based line and one row-level proof line per attribute
   (address, email, phone, payment, tax; bank for supplier only). See "Proof tables" below.
4. **`4_workday_eib_output_counts`** — row count per `workday_{customer,supplier}_*` table, i.e.
   what actually landed in each EIB tab. Tables permanently disabled via a hardcoded `1 = 2` (or
   similar always-false) filter in their transform are excluded here — they're expected to be
   empty by design, not a coverage gap, and listing them as `0` would read as a false alarm.
5. **`5_name_traceability`** — count of records with a full source-name audit trail (see below).
   Not a pass/fail metric.

## Side-by-side summary across both domains

There's no single cross-domain table — query both and compare by eye, or build a quick side-by-side:

```sql
SELECT 'customer' AS domain, * FROM reconcile_customer_report
 UNION ALL
SELECT 'supplier' AS domain, * FROM reconcile_supplier_report
 ORDER BY section, metric, domain;
```

For a management-facing summary, group by `section` and eyeball the `1_gp_id_reconciliation` and
`2_source_row_funnel` rows first — if those are clean (0 issues, funnel fully accounted for), the
population itself is trustworthy; the `3_attribute_coverage` rows then tell you which attributes
have known, quantified exceptions.

## Proof tables (row-level detail behind the report)

Two kinds of checks exist per attribute, both feeding the report:

- **Count-based** (`reconcile_*_attribute_coverage.sql`, `reconcile_*_phone_coverage.sql`) —
  aggregate source count vs. target count per entity. Fast, but can miss a *content* mismatch
  where the row count matches but a field within it doesn't (see the address gotcha below).
- **Row-level proof** (`reconcile_*_{address,email,phone,payment,tax}_proof.sql`,
  `reconcile_*_bank_proof.sql` for supplier, `reconcile_*_gp_id_proof.sql`) — one row per
  source value with an explicit `found_in_workday_*` Y/N flag, queryable directly for any specific
  customer/supplier or value. This is what caught the address content-mismatch bug (see below) —
  build the row-level version for anything the count-based check doesn't fully trust.

| Attribute | Match type | Notes |
|---|---|---|
| GP ID | Exact | Backed by `validate_*_id_coverage.sql`; nothing to reconcile beyond it |
| Address | Exact, country-format-aware | See gotcha below |
| Bank/settlement (supplier only) | Exact, on account number/IBAN | Customer's settlement table is disabled (`1=2`), so no customer equivalent |
| Email | Exact, per entity | Shared emails are deduped cross-entity — see gotcha below |
| Phone | **Approximate** | Never treat as an exact proof — see gotcha below |
| Payment / Tax | Coverage only | 1:1 passthrough from the final master table, no aggregation, coverage guaranteed by construction |
| Name | **Traceability, not proof** | See gotcha below |

## Known, permanent gotchas

These aren't bugs to fix — they're structural properties of the transform (which reconciliation
must account for, not "correct"). Read these before re-deriving any of this from scratch.

**Country-specific address field suppression.** `workday_{customer,supplier}_address_base.sql`
deliberately NULLs `address_line_2/3/4`, `city`, `region`, `postal_code` per country, driven by
`ref_workday_validate_{customer,supplier}_address` (e.g. `address_line_3/4 = 'Not Accepted'` for
the US/Canada/China). This is correct Workday-format behaviour, not data loss — the address proof
exempts a field from comparison whenever the country's rule says `Not Accepted`. **The join to
that ref table must be on country NAME, not country code** — the ref table uses 3-letter ISO
codes (`USA`, `CAN`) while the source uses 2-letter codes (`US`, `CA`); joining on code silently
matches nothing and every field-suppression exemption fails to fire, producing false "not found"
rows that look like real bugs but are actually a broken join key in the reconcile check itself.

**Email cross-entity dedup, not loss.** `workday_{customer,supplier}_email.sql` keeps each email
address under only one entity via `RANK() OVER (PARTITION BY email_address ORDER BY
{customer,supplier}_id, suffix)` — when two entities share an email, the lower-ID one wins,
arbitrarily (no business logic decides "who owns it"). Confirmed empirically: every email flagged
"not found" for one entity is present under a *different* entity's ID, never vanished from the
dataset entirely. Query `reconciliation_{customer,supplier}_issues WHERE rule_type =
'CROSS_ID_DEDUP'` for the current list.

**Phone matching is approximate, never exact.** `workday_{customer,supplier}_phone.sql` reformats
numbers heavily (strips leading zeros, guesses/strips country codes, re-pads), so raw source
digits will not literally equal the final `phone_number` string. The row-level phone proof strips
non-digits and tests suffix AND prefix containment in both directions (plus a last-9-digit match).
**Prefix containment matters specifically for supplier**: a lot of supplier source phone values
are a valid number with garbage trailing zeros appended (e.g. `75239943970000`, likely an Excel
numeric-precision export artifact), and the transform correctly recovers the valid embedded number
(`7523994397`) by truncating the tail — the valid number is a *prefix* of the padded source value,
not a suffix. A suffix-only match wrongly flags this as not found even though the transform handled
it correctly. Treat a `Y` from this check as "likely present," never a certainty; treat the
count-based check (`reconcile_*_phone_coverage.sql`) as the more reliable aggregate number, and use
the row-level proof only for drill-down.

**vendor_id business-unit collisions are split, not merged — with a small, surfaced false-positive
rate.** GP `vendor_id` is scoped per company book, not globally unique, so `src_fin_supplier_raw.sql`
splits a `vendor_id` into separate suppliers whenever the fully-normalised name disagrees across
business units (see `finance-supplier.md`). That name-based signal isn't perfect: a handful of
splits are the same real vendor entered with a typo, abbreviation, or spacing/hyphen inconsistency
(confirmed: `ZENITH VEHICLE CONTRACT` vs `...CONTRACTS`, `PCL COPORATEWEAR` vs `PCL CORPORATEWEAR`).
No safe automated rule distinguishes every one of these from a real collision (e.g. `READING CITY
FC` vs `READING FC` looks similar but may be two real clubs). `reconcile_supplier_collision_review`
ranks every split by name similarity (`jaro_winkler_similarity`) so the high-similarity ones can be
checked before publishing as separate suppliers, instead of reviewing the full list blind.

**Names are traceability, not a match target.** The final `nrm_{customer,supplier}_name` is
deliberately conformed (case normalised, `AND`→`&`, legal-suffix stripping, fuzzy merge across
duplicate GP records) — it is *supposed* to differ from the raw source name. There is nothing to
prove by equality here. `reconcile_*_name_traceability.sql` instead reconstructs, for every final
record, the full list of distinct raw source names that were merged into it (by re-joining the
final record's aggregated GP ID list back to the raw base table) — this is an audit trail, not a
pass/fail check.

**Disabled tables are excluded from EIB output counts, not shown as zero.** Both domains have
`workday_*` tables permanently disabled via a hardcoded `WHERE 1 = 2` (or equivalent) filter in
their transform SQL — e.g. `workday_customer_settlement_account`,
`workday_supplier_company_restrictions`. Grep the domain's `config/transform/sql/finance/*/
workday_*.sql` for `1 = 2` before assuming a table belongs in section 4 of the report; a `0` next
to a disabled table would read as a false alarm.

**Payment/tax coverage is guaranteed by construction, not actively enforced by a check.**
`workday_{customer,supplier}_payment.sql` and `workday_{customer,supplier}_tax.sql` are 1-row-per-entity
passthroughs from the final master table (`src_fin_{customer,supplier}`) with no `UNNEST` and no
row-dropping filter that could differ from what's already in the master table — so a shortfall
here is structurally impossible, not merely unlikely. (Supplier's `workday_supplier_payment.sql`
does carry a `WHERE nrm_payment_terms_id IS NOT NULL` clause, but `src_fin_supplier` is *already*
filtered to only contain suppliers with a non-null value there, so the clause is redundant in
practice, not a second filter.) The payment/tax proof tables exist for symmetry and explicit
sign-off evidence, not because a gap is expected.

## Reproducing this from a cold start

1. Confirm the domain's transform and master table exist and are current (re-run
   `fin_{customer_debtor,supplier_creditor}.py --run-to-phase TRANSFORM` if unsure).
2. Run validate, then reconcile, for the domain(s) needed (commands above).
3. Query `reconcile_{customer,supplier}_report ORDER BY section, metric` for the summary.
4. For any non-zero `3_attribute_coverage` row, drill into the matching
   `reconcile_*_{attribute}_proof` table (or `reconciliation_{customer,supplier}_issues` for the
   row-level detail behind the count-based checks) to get the actual customer/supplier IDs and
   values affected.
5. If a *new* discrepancy pattern shows up that isn't one of the gotchas above, do not assume it's
   a transform bug — check whether it's (a) a false positive in the matching logic used by the
   proof query itself (see the phone prefix/suffix gotcha for a real example of this), or (b) a
   genuine source data-quality issue, before concluding the transform is at fault.

## File inventory

Per domain, in `config/reconcile/finance/{customer,supplier}/`:

- `init.sql` — creates `reconciliation_{customer,supplier}_issues`
- `reconcile_*_row_funnel.sql` — section 2 source
- `reconcile_*_attribute_coverage.sql` — address/email count-based checks (+ bank for supplier)
- `reconcile_*_phone_coverage.sql` — phone count-based check
- `reconcile_*_gp_id_proof.sql`, `reconcile_*_address_proof.sql`, `reconcile_*_email_proof.sql`,
  `reconcile_*_phone_proof.sql`, `reconcile_*_payment_proof.sql`, `reconcile_*_tax_proof.sql`,
  `reconcile_*_bank_proof.sql` (supplier only) — row-level proof tables
- `reconcile_*_name_traceability.sql` — audit trail table
- `reconcile_supplier_collision_review.sql` (supplier only) — every vendor_id business-unit
  collision split, ranked by name similarity so likely false positives can be checked before
  publish (see gotcha below)
- `z_reconcile_*_report.sql` — builds `reconcile_{customer,supplier}_report` (prefixed `z_` so it
  runs after every table it reads from)
- `z_reconcile_*_summary.sql` — smaller standalone issue-count summary, mostly superseded by
  section 3 of the consolidated report

CCS customer has a separate pack in `config/reconcile/finance/customer_ccs/`, centered on:

- `reconcile_customer_ccs_report` — completed GP/transform/extract reconciliation summary
- `reconcile_customer_ccs_closure_summary` — close-out bucket summary
- `reconcile_customer_ccs_closure_decision_pack` — preserved operator decision pack
- `zz_reconcile_customer_ccs_board_summary.sql` — board-level rollup (first workbook tab);
  reads the report/provenance/contact/bucket tables, so it is prefixed `zz_` to run after the
  `z_` report. To reuse for suppliers: add an equivalent `zz_reconcile_supplier_*_board_summary.sql`
  and a `board_summary` sheet at the top of that domain's review publish config.
- `examples/fin_customer_ccs_reconcile_sql.py` — end-to-end refresh/reconcile/export runner

Standalone runners: `examples/fin_customer_debtor_reconcile_sql.py`,
`examples/fin_supplier_creditor_reconcile_sql.py` — thin wrappers calling
`run_validation_sql_dir()` against the domain's reconcile folder, same pattern as the validate
runners.
