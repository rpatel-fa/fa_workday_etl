# CCS Customer Reconciliation - Closed

## Resume (start here)

- From `elt_ingest_excel/docs/todo/archive/todo_ccs_customer.md`: CCS customer GP -> Workday reconciliation is closed. Re-run the CCS reconcile/export commands only if a new GP source cut or Workday extract needs reassessment.

## Session start prompt

```text
load the use-context skill, review elt_ingest_excel/docs/reconciliation.md and elt_ingest_excel/docs/todo/archive/todo_ccs_customer.md, then rerun the CCS reconcile/export flow only if a new source or extract refresh needs reassessment
```

## Goal

Create a reconciliation pack that explains, at source `customer_number` grain, how GP customers in
`src_fin_customer_base` map through the merged customer master in `src_fin_customer` and into the
loaded Workday population in `src_fin_workday_extract_ccs_customer`.

The immediate focus is reporting and proof, not reworking the upstream customer transforms under
`config/transform/sql/finance/customer/`. We only go upstream when a reconciliation exception
shows a concrete break in the lineage.

## Current status

The CCS customer reconciliation pack is implemented, verified end to end, and now complete. The
reconciliation logic, closure summaries, and review workbook export are all in place. This doc now
serves as the durable close-out checkpoint and historical audit note for the finished workstream.

- Reconcile runner:
  - `elt_ingest_excel/examples/fin_customer_ccs_reconcile_sql.py`
- Review workbook export config:
  - `elt_ingest_excel/config/publish/finance/publish_customer_ccs_reconcile_review.json`
- Latest workbook output:
  - `~/Documents/__data/excel/workday/reconcile/customer_ccs_reconcile_review_pack.xlsx`
- Workbook tab order now starts with:
  - `ccs_report`

## Cold-start summary

Use this section first in a fresh session.

### Population and contract

- `src_fin_customer_base`: `6,056` rows / `6,031` distinct `customer_number`
- `src_fin_customer`: `4,159` merged customers
- `src_fin_workday_extract_ccs_customer`: `4,160` extract rows
- source customer numbers missing from Workday: `0`
- extra Workday row not in source population: `1`
  - dummy test row `customer_id = 100001`

Primary reconciliation lens:
- direct sent Workday payload (`workday_*`) vs loaded extract
  (`src_fin_workday_extract_ccs_customer`) on shared `customer_id`

### Final contact status at close-out

From `reconcile_customer_ccs_contact_summary`:

- `default_address`
  - `EXACT_MATCH = 5,703`
  - `FORMAT_VARIATION = 260`
  - `VALUE_MISMATCH = 48`
  - `BOTH_BLANK = 20`
- `default_phone_number`
  - `FORMAT_VARIATION = 5,117`
  - `BOTH_BLANK = 914`
  - no remaining phone review rows
- email
  - not supported in the current extract

### Final residual address buckets at close-out

From `reconcile_customer_ccs_contact_review_pattern_summary`:

- `manual_review = 16`
- `blank_extract_default_address = 13`
- `suspected_uk_locality_substitution = 11`
- `postcode_city_layout_compaction = 5`
- `us_city_state_postcode_layout = 3`

### What is already accepted

From `reconcile_customer_ccs_closure_summary`:

- dummy Workday test row -> `ACCEPT_EXCEPTION`
- extract omits company-number tax payload -> `ACCEPT_EXTRACT_BEHAVIOUR`
- extract does not echo transform default tax code -> `ACCEPT_EXTRACT_BEHAVIOUR`
- semantic tax wording differences -> `ACCEPT_EQUIVALENCE`
- Isle of Man -> United Kingdom -> `ACCEPT_WORKDAY_BEHAVIOUR`
- address `FORMAT_VARIATION` rows -> `ACCEPT_EQUIVALENCE`
- phone `FORMAT_VARIATION` rows -> `ACCEPT_EQUIVALENCE`
- email in current extract -> `ACCEPT_NOT_SUPPORTED`

### Final decision pack at close-out

Final `reconcile_customer_ccs_closure_decision_pack` rows at close-out = `8`:

1. `workday_target_state_review`
2. `ireland_sz_rule_gap_review`
3. `default_address_value_mismatch`
4. `address_review_bucket:postcode_city_layout_compaction`
5. `address_review_bucket:us_city_state_postcode_layout`
6. `address_review_bucket:blank_extract_default_address`
7. `address_review_bucket:suspected_uk_locality_substitution`
8. `address_review_bucket:manual_review`

The workbook tab `closure_decisions` remains the quickest operator-facing entry point if this
finished reconciliation ever needs to be revisited.

## Scope and data grain

- **Source grain:** `src_fin_customer_base.customer_number` = one GP customer identity.
- **Merged transform grain:** `src_fin_customer.nrm_agg_customer_number` = one final merged customer,
  possibly containing multiple pipe-delimited GP customer numbers.
- **Target grain for reconciliation:** explode `src_fin_workday_extract_ccs_customer.reference_id`
  back to one row per source `customer_number`, then compare at that grain.
- **Known target exception:** `src_fin_workday_extract_ccs_customer` contains one dummy Workday test
  record that is not present in `src_fin_customer`; that row must be surfaced explicitly and kept
  separate from genuine reconciliation issues.

## Existing repo assets to reuse

- `config/transform/sql/finance/customer/src_fin_customer_rnk.sql`
  - produces `nrm_agg_customer_number` via `STRING_AGG(..., '|')`
- `config/transform/sql/finance/customer/src_fin_customer.sql`
  - final merged customer master; `customer_id` here is run-scoped, not a stable business key
- `config/transform/sql/finance/workday_extract/src_fin_customer_match.sql`
- `config/transform/sql/finance/workday_extract/src_fin_customer_missing.sql`
  - both already show the right join pattern: explode both aggregated ID lists before matching
- `docs/finance-customer.md`
  - documents the two-stage dedup and confirms `reference_id` is the stable business key
- `docs/reconciliation.md`
  - documents the existing proof-table/report style that can be adapted for this CCS-specific work

## Planned report pack

1. **Population and lineage funnel**
   - GP source row count and distinct `customer_number`
   - final merged `src_fin_customer` row count
   - Workday extract row count
   - exploded Workday `reference_id` count
   - source `customer_number` coverage into merged customers and into Workday

2. **Merge bridge**
   - one bridge row per source `customer_number` showing:
     - source attributes from `src_fin_customer_base`
     - merged customer `customer_id`
     - merged `nrm_agg_customer_number`
     - Workday extract row(s) reached via exploded `reference_id`
   - this becomes the audit spine for every later report

3. **Coverage exceptions**
   - GP customer numbers missing from Workday
   - Workday reference IDs not found in GP/transformed customer
   - duplicate mappings:
     - one GP customer number linked to multiple Workday rows
     - one Workday row linked to overlapping GP populations

4. **Merge analytics**
   - how many source customer numbers were merged into one final customer
   - which merged customers appear fully in Workday vs partially represented
   - whether any merge group is split across multiple Workday rows

5. **Attribute reconciliation at source grain**
   - start with identity-safe fields first:
     - customer name
     - business unit
     - country / region
     - payment terms
     - tax identifiers
   - then profile multi-valued/contact fields:
     - address
     - email
     - phone
   - exact comparison rules should be defined from observed extract columns, not assumed up front

## Delivery shape

Prefer a CCS-specific set of SQL outputs rather than overloading the existing generic
`config/reconcile/finance/customer/` pack, because that current pack reconciles GP -> EIB output
tables and the older `fin_workday_extract_customer` pattern, not GP -> `src_fin_workday_extract_ccs_customer`.

Likely implementation shape:

- `config/reconcile/finance/customer_ccs/`
  - `init.sql`
  - source / merged / Workday explode tables
  - bridge and exception tables
  - summary report table
- `docs/reconciliation.md`
  - extend only after the CCS-specific pack is working and the rules are proven

## Backlog Item: CCS customer GP -> Workday reconciliation implementation

### Done

- Created this anchor doc at `elt_ingest_excel/docs/todo/archive/todo_ccs_customer.md`.
- Confirmed the core lineage:
  - `src_fin_customer_base` unions the GP debtor extracts
  - `src_fin_customer_rnk` builds `nrm_agg_customer_number`
  - `src_fin_customer` is the final merged customer master
  - `src_fin_workday_extract_ccs_customer` is the loaded Workday customer extract to reconcile against
- Confirmed the existing repo already uses the correct explode-then-join pattern for aggregated
  IDs in:
  - `config/transform/sql/finance/workday_extract/src_fin_customer_match.sql`
  - `config/transform/sql/finance/workday_extract/src_fin_customer_missing.sql`
- Captured the main planning decision: the new reconciliation must be driven from source
  `customer_number` grain and exploded Workday `reference_id`, not only from merged customer rows.
- Profiled the live DuckDB tables and locked down the core population counts:
  - `src_fin_customer_base`: 6,056 rows / 6,031 distinct `customer_number`
  - `src_fin_customer`: 4,159 merged customers / 6,031 exploded source customer numbers
  - `src_fin_workday_extract_ccs_customer`: 4,160 rows / 6,032 exploded `reference_id` customer numbers
- Confirmed live population coverage:
  - 0 GP source customer numbers missing from Workday
  - 1 extra Workday reference ID not in source: `CUSTOMER-9-1`
  - 1 extra Workday row not in `src_fin_customer`: `customer_id = 100001`, `customer = test1`
- Implemented a CCS-specific reconciliation pack under `elt_ingest_excel/config/reconcile/finance/customer_ccs/`:
  - `init.sql`
  - `01_reconcile_customer_ccs_bridge.sql`
  - `02_reconcile_customer_ccs_expected_and_exceptions.sql`
  - `03_reconcile_customer_ccs_attribute_detail.sql`
  - `z_reconcile_customer_ccs_report.sql`
- Added a dedicated runner:
  - `elt_ingest_excel/examples/fin_customer_ccs_reconcile_sql.py`
- Ran the pack successfully against the live DuckDB database.
- Built the main outputs:
  - `reconcile_customer_ccs_bridge`
  - `reconcile_customer_ccs_missing_in_workday`
  - `reconcile_customer_ccs_extra_in_workday`
  - `reconcile_customer_ccs_workday_dummy_row`
  - `reconcile_customer_ccs_attribute_detail`
  - `reconcile_customer_ccs_attribute_summary`
  - `reconcile_customer_ccs_report`
- First-pass attribute reconciliation at source `customer_number` grain is now quantified:
  - exact matches across all 6,031 source customer numbers for:
    - `customer_name`
    - `reference_id`
    - `customer_category`
    - `default_payment_terms`
    - `default_payment_type`
    - `default_currency`
  - exact mismatches:
    - `tax_id`: 15
    - `tax_id_country`: 15
    - `tax_id_type`: 15
    - `tax_status_country`: 3
    - `transaction_tax_status`: 330
- The dominant mismatch patterns are now visible:
  - 15 source customer numbers where current transformed tax ID fields are populated but the Workday extract is blank
  - 3 tax-status-country outliers:
    - `BOX0624111261` / `BOX0921108112` (`C-003355`) expected `United Kingdom`, actual `Singapore`
    - `L00061` (`C-001773`) expected `Isle of Man`, actual `United Kingdom`
  - 330 exact-string `transaction_tax_status` mismatches, dominated by:
    - expected `Non EU Company Not VAT registered in this country` vs actual `Company is not VAT/GST Registered in this Country` (185)
    - expected `EU Company - Not Registered in this country` vs actual `Company is not VAT/GST Registered in this Country` (143)
- Added a second-pass tax-classification layer in code, pending rerun/verification once the DuckDB
  lock is released:
  - `elt_ingest_excel/config/reconcile/finance/customer_ccs/04_reconcile_customer_ccs_tax_analysis.sql`
  - updated `init.sql` to reset the new tax-analysis tables
  - updated `03_reconcile_customer_ccs_attribute_detail.sql` so raw tax exact-string differences are
    no longer all pushed directly into `reconciliation_customer_ccs_issues`
  - updated `z_reconcile_customer_ccs_report.sql` to surface classified tax outcomes separately
    from the raw exact-string comparison
- Reran the CCS pack successfully after the DuckDB lock was released and verified the new
  tax-classification outputs:
  - `reconcile_customer_ccs_tax_analysis`
  - `reconcile_customer_ccs_tax_analysis_summary`
  - `reconcile_customer_ccs_report`
- Verified second-pass tax breakdown:
  - `tax_id`: 6,016 exact matches / 15 `EXTRACT_OMITS_COMPANY_NUMBER`
  - `tax_id_country`: 6,016 exact matches / 15 `EXTRACT_OMITS_COMPANY_NUMBER`
  - `tax_id_type`: 6,016 exact matches / 15 `EXTRACT_OMITS_COMPANY_NUMBER`
  - `tax_status_country`: 6,028 exact matches / 1 `COUNTRY_NORMALISATION_REVIEW` / 2 `VALUE_MISMATCH`
  - `transaction_tax_status`: 5,701 exact matches / 328 `SEMANTIC_EQUIVALENT` / 2 `VALUE_MISMATCH`
- Verified the issue inventory is now tightly narrowed:
  - `7` total issue rows
  - `2` population exceptions (the extra dummy Workday row reported two ways)
  - `0` non-tax exact attribute mismatches
  - `5` tax review rows
- Proved the former 15 tax-ID missing cases are one systematic extract behavior, not 15 unrelated
  defects:
  - they collapse to 12 merged customers
  - every case is `United Kingdom` with expected tax ID type `Company Number`
  - `src_fin_workday_extract_ccs_customer.company_tax_id_type` is blank for all 4,160 rows
  - these are now classified as `EXTRACT_OMITS_COMPANY_NUMBER` and reported as
    `INFO - extract limitation`
- Verified the remaining review set is now only five tax-review rows across four merged customers:
  - `BOX0624111261` / `BOX0921108112` (`C-003355`): `tax_status_country` expected `United Kingdom`,
    actual `Singapore`
  - `L00061` (`C-001773`): `tax_status_country` expected `Isle of Man`, actual `United Kingdom`
    -> currently classified as `COUNTRY_NORMALISATION_REVIEW`
  - `M01730` (`C-003538`) and `M01976` (`C-001480`): `transaction_tax_status` expected
    `EU Company - Exempt in THIS country`, actual `Customer of Goods`
- Confirmed the two `transaction_tax_status` review rows share the same pattern:
  - country `Ireland`
  - `nrm_tax_schedule_id = 'SZ'`
  - no `tax_registration_number`
- Confirmed the Singapore country mismatch is tied to a merged customer whose source population spans
  both `GB` and `SG` countries inside one `nrm_agg_customer_number` group.
- Added dedicated reporting drilldowns for the remaining tax-review scope:
  - `elt_ingest_excel/config/reconcile/finance/customer_ccs/05_reconcile_customer_ccs_tax_review_detail.sql`
  - `reconcile_customer_ccs_tax_review_customer`
  - `reconcile_customer_ccs_tax_review_source_rows`
  - `reconcile_customer_ccs_tax_review_pattern_summary`
  - `reconcile_customer_ccs_report` now includes section `3_tax_review_focus`
- Verified the remaining review scope is now explicitly reportable as:
  - `4` merged customers
  - `7` contributing source rows
  - `3` distinct review patterns
- Verified the pattern summary:
  - `tax_status_country`: `Isle of Man` -> `United Kingdom` (`COUNTRY_NORMALISATION_REVIEW`) = 1 source row / 1 merged customer
  - `tax_status_country`: `United Kingdom` -> `Singapore` (`VALUE_MISMATCH`) = 2 source rows / 1 merged customer
  - `transaction_tax_status`: `EU Company - Exempt in THIS country` -> `Customer of Goods`
    (`VALUE_MISMATCH`) = 2 source rows / 2 merged customers
- Added an operator-facing review-pack layer:
  - `elt_ingest_excel/config/reconcile/finance/customer_ccs/06_reconcile_customer_ccs_tax_review_pack.sql`
  - `reconcile_customer_ccs_tax_review_case_summary`
  - `reconcile_customer_ccs_tax_review_operator_pack`
  - `reconcile_customer_ccs_report` now includes:
    - `3d_tax_review_case_summaries`
    - `3e_tax_review_operator_pack_rows`
- Verified the operator pack condenses the remaining review scope into 4 prioritised review rows:
  1. `C-003355` (`SCORPIO IT SERVICES PTE`) -> `mixed_source_country` / `High`
  2. `C-001480` (`GOOGLE CLOUD EMEA`) -> `ireland_sz_rule_gap` / `Medium`
  3. `C-003538` (`SSBL`) -> `ireland_sz_rule_gap` / `Medium`
  4. `C-001773` (`ISLE OF MAN FA`) -> `normalisation` / `Low`
- Verified the suggested interpretation text is now stored alongside each review row, so the pack
  is operator-ready without needing ad hoc query notes.
- Added a direct transform-vs-extract comparison layer for the tax payload sent to Workday versus
  the loaded Workday extract:
  - `elt_ingest_excel/config/reconcile/finance/customer_ccs/07_reconcile_customer_ccs_transform_extract_diff.sql`
  - `reconcile_customer_ccs_transform_extract_tax_diff_detail`
  - `reconcile_customer_ccs_transform_extract_tax_diff_summary`
  - `reconcile_customer_ccs_transform_extract_tax_diff_pattern_summary`
  - `reconcile_customer_ccs_report` now includes section `4_transform_vs_extract`
- Verified direct `workday_customer_tax` -> `src_fin_workday_extract_ccs_customer` comparison on
  shared `customer_id` values:
  - `4,159` matched customer IDs
  - `1` extract-only row, confirmed as the dummy Workday record `customer_id = 100001`
  - `0` transform-only rows
  - confirmed sent-vs-extracted diffs on matched customer IDs:
    - `tax_code`: `4,159`
    - `tax_id_country`: `12`
    - `tax_id`: `12`
    - `tax_id_type`: `12`
    - `primary_tax_id`: `12`
    - `transaction_tax_id`: `12`
    - `tax_status_country`: `2`
    - `transaction_tax_status`: `287`
    - `withholding_tax_status`: `0`
- Verified top direct diff patterns:
  - `tax_code`: `GBR 20% VAT (20%)` -> blank = `4,159`
  - `tax_id_country`: `United Kingdom` -> blank = `12`
  - `tax_id_type`: `Company Number` -> blank = `12`
  - `primary_tax_id`: `Yes` -> blank = `12`
  - `transaction_tax_id`: `Yes` -> blank = `12`
  - `tax_status_country`: `Isle of Man` -> `United Kingdom` = `1`
  - `tax_status_country`: `United Kingdom` -> `Singapore` = `1`
  - `transaction_tax_status`: `Non EU Company Not VAT registered in this country`
    -> `Company is not VAT/GST Registered in this Country` = `157`
  - `transaction_tax_status`: `EU Company - Not Registered in this country`
    -> `Company is not VAT/GST Registered in this Country` = `128`
  - `transaction_tax_status`: `EU Company - Exempt in THIS country` -> `Customer of Goods` = `2`
- Added a provenance summary layer that collapses the source, transform, and loaded-extract evidence
  into a small set of operator-facing case types:
  - `elt_ingest_excel/config/reconcile/finance/customer_ccs/08_reconcile_customer_ccs_provenance_summary.sql`
  - `reconcile_customer_ccs_provenance_case_summary`
  - `reconcile_customer_ccs_provenance_bucket_summary`
  - `reconcile_customer_ccs_report` now includes section `5_provenance`
- Verified the provenance summary condenses the current diff landscape into 7 case rows:
  - systematic:
    - `extract_omits_company_number_payload` = 1 case
    - `extract_does_not_echo_transform_default_tax_code` = 1 case
    - `semantic_wording_variation` = 2 cases
  - operator review:
    - `mixed_source_country_review` = 1 case
    - `ireland_sz_rule_gap_review` = 1 case
    - `country_normalisation_review` = 1 case
- Verified the provenance case counts:
  - `extract_omits_company_number_payload` = 15 source customer numbers / 12 merged customers / 12 matched extracted customer IDs
  - `extract_does_not_echo_transform_default_tax_code` = 4,159 matched extracted customer IDs
  - `semantic_wording_variation`:
    - `EU Company - Not Registered in this country` -> `Company is not VAT/GST Registered in this Country`
      = 143 source customer numbers / 128 matched extracted customer IDs
    - `Non EU Company Not VAT registered in this country`
      -> `Company is not VAT/GST Registered in this Country`
      = 185 source customer numbers / 157 matched extracted customer IDs
  - `mixed_source_country_review` = 2 source customer numbers / 1 merged customer / 1 matched extracted customer ID
  - `ireland_sz_rule_gap_review` = 2 source customer numbers / 2 merged customers / 2 matched extracted customer IDs
  - `country_normalisation_review` = 1 source customer number / 1 merged customer / 1 matched extracted customer ID
- Updated the end-to-end runner and CLI so CCS Workday extract refresh plus reconciliation now runs
  as one package of work:
  - `elt_ingest_excel/examples/fin_customer_ccs_reconcile_sql.py` now:
    - runs `ingest/finance/workday_ccs_customer.json`
    - runs `transform/sql/finance/workday_extract_ccs_customer`
    - then runs `config/reconcile/finance/customer_ccs`
  - `elt_ingest_excel/docs/CLI.md` now documents:
    - `uv run python examples/fin_customer_ccs_reconcile_sql.py --run-to-phase TRANSFORM`
- Verified the new end-to-end entry point works on the live DuckDB database:
  - the extract ingest refreshed `fin_workday_extract_customer` with `4,160` rows
  - the transform refreshed `src_fin_workday_extract_ccs_customer`
  - the reconcile pack reran successfully in the same command
  - report verification after the run confirmed:
    - `4_transform_vs_extract.4a_matched_customer_ids = 4159`
    - `5_provenance.5a_provenance_case_rows = 7`
- Improved the end-to-end runner UX so the single CCS command now prints a compact reconcile
  summary after the SQL pack completes:
  - `1b_distinct_gp_customer_number`
  - `1c_src_fin_customer_row_count`
  - `1e_src_fin_workday_extract_ccs_customer_row_count`
  - `2b_source_customer_numbers_missing_in_workday`
  - `2c_extra_workday_reference_ids_not_in_source`
  - `4a_matched_customer_ids`
  - `4a_tax_status_country_diff_on_matched_customer_ids`
  - `4a_transaction_tax_status_diff_on_matched_customer_ids`
  - `5a_provenance_case_rows`
  - plus pointers to:
    - `reconcile_customer_ccs_provenance_case_summary`
    - `reconcile_customer_ccs_tax_review_operator_pack`
    - `reconcile_customer_ccs_report`
- Verified the summary print from the live end-to-end run:
  - GP source customer numbers = `6031`
  - merged customers = `4159`
  - Workday extract rows = `4160`
  - source customer numbers missing in Workday = `0`
  - extra Workday reference IDs not in source = `1`
  - matched customer IDs in direct transform-vs-extract compare = `4159`
  - tax-status-country diffs on matched customer IDs = `2`
  - transaction-tax-status diffs on matched customer IDs = `287`
  - provenance case rows = `7`
- Extended the CCS pack into contact-field reconciliation where the current extract supports it:
  - `elt_ingest_excel/config/reconcile/finance/customer_ccs/09_reconcile_customer_ccs_contact_analysis.sql`
  - `reconcile_customer_ccs_contact_detail`
  - `reconcile_customer_ccs_contact_summary`
  - `reconcile_customer_ccs_contact_extract_support`
  - `reconcile_customer_ccs_report` now includes section `6_contact_reconciliation`
- Verified source-grain contact reconciliation on the live DuckDB database:
  - `default_address`: `5,703` exact matches / `24` `FORMAT_VARIATION` / `284` `VALUE_MISMATCH` / `20` blank in both
  - `default_phone_number`: `5,115` `FORMAT_VARIATION` / `2` `VALUE_MISMATCH` / `914` blank in both
  - merged-customer view of the same contact scope:
    - `default_address`: `3,879` exact / `18` format variation / `249` review / `13` blank in both
    - `default_phone_number`: `3,312` format variation / `1` review / `846` blank in both
- Proved the extract currently does not support clean email reconciliation:
  - `src_fin_workday_extract_ccs_customer.customer_contacts` is blank for all `4,160` extract rows
  - `workday_customer_email` still contains `3,935` distinct sent customer IDs
  - email is now explicitly reported as unsupported in the current extract rather than forced into a false mismatch inventory
- Narrowed the new phone review scope to one merged customer / two source customer numbers:
  - `C-002586` (`MOHAMED MATAR`) / `CWN0523109695|CWN0626113648`
  - transform primary phone = `39099140`
  - loaded extract default phone = `+44  39099140`
  - this is now isolated as the only remaining phone normalization/parsing review row in the current extract
- Updated the end-to-end runner output so the one-command workflow now points operators to:
  - `reconcile_customer_ccs_contact_summary`
  - alongside the existing provenance / tax review / report tables
- Added second-pass contact equivalence rules to reduce false review noise in the current extract:
  - `elt_ingest_excel/config/reconcile/finance/customer_ccs/09_reconcile_customer_ccs_contact_analysis.sql`
    now classifies:
    - address token-set equivalence where the extract reorders or compacts the same address
      components into different lines
    - phone formatting equivalence using `src_fin_customer.nrm_phone_code` as a fallback country
      code when the sent primary Workday phone row loses its explicit international code
- Verified the second-pass contact burn-down on the live DuckDB database:
  - `default_address`: `5,703` exact matches / `159` `FORMAT_VARIATION` / `149` `VALUE_MISMATCH` /
    `20` blank in both
  - `default_phone_number`: `5,117` `FORMAT_VARIATION` / `914` blank in both / `0` remaining
    review rows
  - merged-customer view of the remaining contact scope:
    - `default_address`: `3,879` exact / `141` format variation / `126` review / `13` blank in both
    - `default_phone_number`: `3,313` format variation / `846` blank in both / `0` review
- Proved the former isolated phone review case is now explained as formatting, not a transform defect:
  - `C-002586` (`MOHAMED MATAR`) no longer appears in `VALUE_MISMATCH`
  - the extract prepends `+44` to the same local digits when the sent Workday phone row drops its
    explicit country-code field, and the reconcile layer now classifies that as `FORMAT_VARIATION`
- Verified the remaining contact review queue is now address-only:
  - `149` source-grain address review rows across `126` merged customers
  - the largest remaining buckets are no longer generic line-compaction noise; they are now
    country/state abbreviation or content-shift cases such as:
    - `Minnesota` vs `MN`
    - `Georgia` vs `GA`
    - `California` vs `CA`
    - blank extract default address against a populated sent address
    - a small number of more suspicious locality/content substitutions (for example `Cornwall` vs
      `Coventry`)
- Added a third-pass address-format equivalence rule for observed region/province abbreviations in
  the loaded extract:
  - `elt_ingest_excel/config/reconcile/finance/customer_ccs/09_reconcile_customer_ccs_contact_analysis.sql`
    now treats the following as `FORMAT_VARIATION` when the rest of the address components match:
    - US states such as `California -> CA`, `Florida -> FL`, `New York -> NY`
    - `Australia: New South Wales -> NSW`
    - `Canada: Ontario -> ON`, `Alberta -> AB`
    - `Brazil: Rio De Janeiro -> RJ`
    - `Italy: Perugia -> PG`
    - `Switzerland: Solothurn -> SO`
- Verified the third-pass contact burn-down on the live DuckDB database:
  - `default_address`: `5,703` exact matches / `211` `FORMAT_VARIATION` / `97` `VALUE_MISMATCH` /
    `20` blank in both
  - merged-customer view of the remaining address scope:
    - `default_address`: `3,879` exact / `184` format variation / `83` review / `13` blank in both
- Verified the remaining contact review queue is now narrowed further to:
  - `97` source-grain address review rows across `83` merged customers
  - no remaining phone review rows
- Confirmed the residual address review scope is now concentrated in cases that still merit manual
  decision or a more explicit business rule, including:
  - blank extract default addresses against populated sent addresses
  - suspected location/content substitutions (for example `Cornwall` -> `Coventry`)
  - a smaller set of country-specific representation changes not yet accepted as equivalence
    (for example some Spain/Singapore/France postcode-city layout variants)
- Materialised and verified operator-facing residual address review buckets in:
  - `elt_ingest_excel/config/reconcile/finance/customer_ccs/09_reconcile_customer_ccs_contact_analysis.sql`
  - `reconcile_customer_ccs_contact_review_case_summary`
  - `reconcile_customer_ccs_contact_review_pattern_summary`
- Verified the residual address review queue is now grouped into operator-facing buckets:
  - initial materialised split before regex-fix verification:
    - `manual_review`: `56` source rows / `49` merged customers
    - `country_specific_region_layout`: `17` source rows / `16` merged customers
    - `blank_extract_default_address`: `13` source rows / `9` merged customers
    - `suspected_uk_locality_substitution`: `11` source rows / `9` merged customers
- Fixed the residual-address review bucket regexes in
  `elt_ingest_excel/config/reconcile/finance/customer_ccs/09_reconcile_customer_ccs_contact_analysis.sql`
  so DuckDB matches regex whitespace/newline classes instead of the literal strings `\\s` / `\\n`.
- Reverified the residual address review queue after the regex fix; the review rows are unchanged
  (`97` source rows / `83` merged customers) but the operator buckets are now more informative:
  - `postcode_city_layout_compaction`: `42` source rows / `40` merged customers
  - `manual_review`: `17` source rows / `12` merged customers
  - `blank_extract_default_address`: `13` source rows / `9` merged customers
  - `us_city_state_postcode_layout`: `12` source rows / `11` merged customers
  - `suspected_uk_locality_substitution`: `11` source rows / `9` merged customers
  - `country_specific_region_layout`: `2` source rows / `2` merged customers
- Fixed the postal-code leading-zero normalization in
  `elt_ingest_excel/config/reconcile/finance/customer_ccs/09_reconcile_customer_ccs_contact_analysis.sql`
  so it is more robust to whitespace/newline token boundaries (and uses DuckDB-safe capture group
  replacements). Rerun verification shows the headline address mismatch count remains `97`, so this
  change did not yet burn down the residual review queue.
- Extended the standard operator artifacts so the new contact-review buckets are visible without ad
  hoc queries:
  - `elt_ingest_excel/config/reconcile/finance/customer_ccs/z_reconcile_customer_ccs_report.sql`
    now includes `6_review_case_rows`, `6_review_patterns`, and one metric per review bucket
  - `elt_ingest_excel/examples/fin_customer_ccs_reconcile_sql.py` now prints the two new contact
    review tables in the post-run review list
- Reframed the CCS report contract so the primary reconciliation lens is now explicit:
  - primary: sent Workday payload (`workday_*` transform output tables) vs loaded Workday extract
  - secondary: GP source vs loaded Workday extract as a diagnostic lens to explain safe cleansing
    and upstream transformation effects
  - `elt_ingest_excel/config/reconcile/finance/customer_ccs/z_reconcile_customer_ccs_report.sql`
    now exposes this contract in section `0_reconciliation_contract`
- Added an operator-friendly Excel export of the key review tables so the Step 1 operator review can
  happen outside DuckDB:
  - config: `elt_ingest_excel/config/publish/finance/publish_customer_ccs_reconcile_review.json`
  - runner flag (after running the reconcile pack): `uv run python examples/fin_customer_ccs_reconcile_sql.py --export-review-workbook-only`
  - output: `~/Documents/__data/excel/workday/reconcile/customer_ccs_reconcile_review_pack.xlsx`
- Added an explicit closure layer that answers “does sent Workday payload match loaded extract?” per
  bucket and declares the closure action required:
  - `elt_ingest_excel/config/reconcile/finance/customer_ccs/10_reconcile_customer_ccs_closure_summary.sql`
  - `reconcile_customer_ccs_closure_summary` (exported as workbook tab `closure_summary`)
  - `reconcile_customer_ccs_report` now includes section `7_closure`
- Added a dedicated operator-decision checklist for only the unresolved closure buckets:
  - `elt_ingest_excel/config/reconcile/finance/customer_ccs/11_reconcile_customer_ccs_closure_decision_pack.sql`
  - `reconcile_customer_ccs_closure_decision_pack` (exported as workbook tab `closure_decisions`)
  - verified current operator-decision checklist = `10` rows
- Updated the close-out interpretation after deeper review of the live data:
  - `C-003355` mixed `GB` / `SG` case is now treated as accepted Workday behaviour because the
    loaded tax country follows the Singapore primary address in Workday
  - Spain case `C-000149` is formatting-only and should remain in `FORMAT_VARIATION`
- Added further country-specific contact normalization passes and reran the pack:
  - US `City, ST ZIP` layout normalization for clear formatting-only rows
  - Australia `City ST ZIP` layout normalization
  - stronger postcode/city compaction handling where the extract preserves the same structural
    address components
  - broader region/province abbreviation handling for countries observed in the live extract
- Verified the current contact burn-down on the live DuckDB database:
  - `default_address`: `5,703` exact matches / `260` `FORMAT_VARIATION` / `48` `VALUE_MISMATCH` /
    `20` blank in both
  - no remaining phone review queue
  - email still explicitly `NOT_SUPPORTED` in the current extract
- Verified the current residual address review queue is now:
  - `manual_review`: `16`
  - `blank_extract_default_address`: `13`
  - `suspected_uk_locality_substitution`: `11`
  - `postcode_city_layout_compaction`: `5`
  - `us_city_state_postcode_layout`: `3`
- Verified the current closure summary state:
  - accepted / non-issue classes now include:
    - dummy Workday row
    - company-number extract omission
    - default tax-code extract non-round-trip
    - semantic tax wording equivalence
    - accepted Workday normalisation (`Isle of Man -> United Kingdom`)
    - address formatting equivalence
    - phone formatting equivalence
    - email unsupported in current extract
  - current decision-pack rows reduced from `10` to `8`
- User confirmed the CCS customer reconciliation is now complete.
- Refined the tax close-out classification logic so the generated reconciliation outputs now record:
  - `C-003355` / SCORPIO as `workday_target_state_review`
  - `Isle of Man -> United Kingdom` as `accepted_workday_normalisation`
- Regenerated the DuckDB reconciliation tables and workbook export after the classification fix:
  - `reconcile_customer_ccs_provenance_case_summary`
  - `reconcile_customer_ccs_closure_summary`
  - `reconcile_customer_ccs_closure_decision_pack`
  - `~/Documents/__data/excel/workday/reconcile/customer_ccs_reconcile_review_pack.xlsx`
- Refreshed the analyst review workbook in place at:
  - `~/Documents/__data/excel/workday/reconcile/customer_ccs_reconcile_review_pack.xlsx`
- Reordered the workbook export so management lands on the headline summary first:
  - `elt_ingest_excel/config/publish/finance/publish_customer_ccs_reconcile_review.json`
  - workbook tab order now starts with `ccs_report`
- Reran the workbook export after the sheet-order change:
  - `uv run python examples/fin_customer_ccs_reconcile_sql.py --export-review-workbook-only`
- Updated the repo docs so the completion state and workbook export commands are captured in:
  - `elt_ingest_excel/docs/CLI.md`
  - `elt_ingest_excel/docs/reconciliation.md`

### Still todo (next actions)

1. No open implementation or reconcile-design work remains for this item; this anchor doc is closed.
2. If a future GP source cut or CCS Workday extract arrives, rerun:
   - `uv run python examples/fin_customer_ccs_reconcile_sql.py --run-to-phase TRANSFORM`
   - `uv run python examples/fin_customer_ccs_reconcile_sql.py --export-review-workbook-only`
3. Use the preserved historical tables and workbook tabs (`closure_summary`,
   `closure_decisions`, `contact_pattern`, `provenance_case`) as the baseline for any future
   delta review rather than rebuilding the reconciliation design from scratch.

### Accumulated Active Constraints

- The sent Workday transform output is the gold definition for reconciliation. The purpose of the
  CCS review is to confirm whether the loaded Workday extract preserves exactly what was sent, not
  to redefine the expected value from extract-side formatting.
- GP source vs Workday extract is a secondary diagnostic comparison only. Use it to explain where
  safe cleansing, normalization, or other upstream preparation changed the raw GP value before
  load; do not let that secondary comparison override the sent Workday payload as the expected
  value.
- Reconcile at source `customer_number` grain first. The merged customer row is an intermediate
  lineage step, not the final reporting grain.
- Always explode pipe-delimited ID lists before joining:
  - `src_fin_customer.nrm_agg_customer_number`
  - `src_fin_workday_extract_ccs_customer.reference_id`
- Treat `src_fin_customer.customer_id` as run-scoped only. It is useful inside one run for lineage,
  but not as the business key for GP <-> Workday comparison.
- The Workday dummy test record must be isolated and labelled, not mixed into defect counts.
- Do not start by rewriting customer transforms. Use reconciliation findings to decide whether any
  transform investigation is needed.
- Reuse the repo's existing reconciliation/proof-table style where it helps, but do not force the
  CCS extract into the older `fin_workday_extract_customer` flow.
- The current CCS pack uses exact string comparison for attribute checks. That is appropriate for
  identity/payment/currency fields, but tax-status wording differences must not be treated as
  automatic defects until semantic-equivalence rules are agreed.
- Population retention is currently confirmed end-to-end for GP source customer numbers: 0 missing
  in merged customers, 0 missing in Workday, 1 expected extra dummy Workday row.
- The second-pass tax classification is now verified against DuckDB and should be treated as the
  primary lens for tax review, not the raw exact-string mismatch counts alone.
- `EXTRACT_OMITS_COMPANY_NUMBER` is now a verified extract limitation class, not a review defect:
  do not mix those rows back into the tax-review queue unless a counterexample appears in a future
  extract.
- When discussing whether a difference comes from transform logic or from loaded Workday state,
  validate first against the direct `workday_customer_tax` -> `src_fin_workday_extract_ccs_customer`
  comparison on shared `customer_id` values. Do not infer transform-side divergence without that
  direct proof.
- For operator discussion, use `reconcile_customer_ccs_provenance_case_summary` as the first-stop
  table: it is now the smallest table that still preserves the key source / transform / extract
  evidence and the suggested interpretation.
- For phone reconciliation, compare the loaded extract phone to the sent transform phone after
  normalizing the extract to digits and prepending the transform international phone code. Raw
  string comparison overstates differences because the extract usually renders display formatting
  and country code while the sent Workday phone row stores the local number separately.
- For address reconciliation, separate exact matches from formatting-only differences before
  counting review work. Multi-line address rendering can differ through line compaction,
  punctuation removal, postcode spacing, or regional abbreviation without proving a business-level
  divergence.
- Address token-set equivalence is now an accepted clustering rule in the CCS pack: if the sent and
  extracted default addresses contain the same normalized components in a different order/layout,
  classify that as `FORMAT_VARIATION`, not `VALUE_MISMATCH`, so the row is still visible as
  extract-side drift from the sent gold value without being mixed into the smaller true-content
  mismatch queue.
- Observed extract-side region/province abbreviations are now also accepted clustering rules for
  the proven country pairs above when the rest of the address token set still matches.
- Do not force email mismatches from the current CCS extract. `customer_contacts` is blank across
  the live extract, so email is presently a supportability finding, not a confirmed data defect.
- For phone reconciliation, if the sent primary phone row loses its explicit international code but
  the merged customer still carries `nrm_phone_code`, use that merged customer phone code as the
  fallback when deciding whether the extract only added display formatting.
- Treat `reconcile_customer_ccs_contact_review_pattern_summary` as a triage aid, not proof of
  semantic equivalence. The grouped buckets are there to focus operator review and candidate rule
  selection; they do not on their own justify auto-reclassifying a residual address mismatch.
- In the contact report, `FORMAT_VARIATION` means "Workday extract differs from the sent gold
  value in representation/layout" rather than "ignore this row." It remains a confirmed
  sent-vs-extract difference, just not the same kind of difference as a content mismatch.
- The contact review bucket counts are only as good as the bucket regexes. In DuckDB SQL literals,
  over-escaped regex strings can silently turn a clustering rule into a no-op while leaving the
  underlying `VALUE_MISMATCH` counts unchanged.

### Verification

- `rg --files "elt_ingest_excel/config/transform/sql/finance/customer"`
- `rg -n "src_fin_customer_base|src_fin_customer|src_fin_workday_extract_ccs_customer|mrg_agg_customer_number|reference_id|customer_number" "elt_ingest_excel/config/transform/sql/finance"`
- `sed -n '1,220p' "elt_ingest_excel/config/transform/sql/finance/customer/src_fin_customer_rnk.sql"`
- `sed -n '1,220p' "elt_ingest_excel/config/transform/sql/finance/workday_extract/src_fin_customer_match.sql"`
- `sed -n '1,220p' "elt_ingest_excel/docs/reconciliation.md"`
- `uv run python examples/fin_customer_ccs_reconcile_sql.py`
- `uv run python - <<'PY' ... SELECT section, metric, metric_value, status FROM reconcile_customer_ccs_report ... PY`
- `uv run python - <<'PY' ... SELECT issue_type, COUNT(*) FROM reconciliation_customer_ccs_issues GROUP BY issue_type ... PY`
- `uv run python - <<'PY' ... SELECT attribute_name, comparison_class, row_count FROM reconcile_customer_ccs_tax_analysis_summary ... PY`
- `uv run python - <<'PY' ... SELECT attribute_name, customer_number, expected_value, actual_value, comparison_class FROM reconcile_customer_ccs_tax_analysis WHERE comparison_class IN ('VALUE_MISMATCH', 'COUNTRY_NORMALISATION_REVIEW') ... PY`
- `uv run python - <<'PY' ... SELECT COALESCE(company_tax_id_type,'<NULL>'), COUNT(*) FROM src_fin_workday_extract_ccs_customer GROUP BY 1 ... PY`
- `uv run python - <<'PY' ... SELECT * FROM reconcile_customer_ccs_tax_review_customer ORDER BY expected_customer_id ... PY`
- `uv run python - <<'PY' ... SELECT * FROM reconcile_customer_ccs_tax_review_pattern_summary ORDER BY attribute_name, comparison_class ... PY`
- `uv run python - <<'PY' ... SELECT * FROM reconcile_customer_ccs_tax_review_operator_pack ORDER BY review_order ... PY`
- `uv run python - <<'PY' ... SELECT metric, metric_value, status FROM reconcile_customer_ccs_report WHERE section = '4_transform_vs_extract' ORDER BY metric ... PY`
- `uv run python - <<'PY' ... SELECT attribute_name, expected_value, actual_value, row_count FROM reconcile_customer_ccs_transform_extract_tax_diff_pattern_summary ORDER BY attribute_name, row_count DESC ... PY`
- `uv run python - <<'PY' ... SELECT review_order, provenance_bucket, issue_scope, attribute_name, expected_value, actual_value, source_customer_number_count, merged_customer_count, matched_customer_id_count FROM reconcile_customer_ccs_provenance_case_summary ORDER BY review_order ... PY`
- `uv run python - <<'PY' ... SELECT provenance_bucket, issue_scope, case_count FROM reconcile_customer_ccs_provenance_bucket_summary ORDER BY issue_scope, provenance_bucket ... PY`
- `uv run python examples/fin_customer_ccs_reconcile_sql.py --run-to-phase TRANSFORM`
- `uv run python -m py_compile examples/fin_customer_ccs_reconcile_sql.py`
- `uv run python - <<'PY' ... SELECT attribute_name, comparison_class, row_count FROM reconcile_customer_ccs_contact_summary ORDER BY attribute_name, comparison_class ... PY`
- `uv run python - <<'PY' ... SELECT metric, row_count, status FROM reconcile_customer_ccs_contact_extract_support ORDER BY metric ... PY`
- `uv run python - <<'PY' ... SELECT metric, metric_value, status FROM reconcile_customer_ccs_report WHERE section = '6_contact_reconciliation' ORDER BY metric ... PY`
- `uv run python - <<'PY' ... SELECT issue_type, COUNT(*) FROM reconciliation_customer_ccs_issues GROUP BY issue_type ORDER BY issue_type ... PY`
- `uv run python - <<'PY' ... SELECT customer_number, expected_customer_id, expected_value, actual_value, comparison_reason FROM reconcile_customer_ccs_contact_detail WHERE attribute_name = 'default_phone_number' AND comparison_class = 'VALUE_MISMATCH' ORDER BY expected_customer_id, customer_number ... PY`
- `uv run python - <<'PY' ... SELECT customer_number, expected_customer_id, expected_value, actual_value FROM reconcile_customer_ccs_contact_detail WHERE attribute_name = 'default_address' AND comparison_class = 'VALUE_MISMATCH' ORDER BY expected_customer_id, customer_number ... PY`
- `uv run python - <<'PY' ... SELECT attribute_name, comparison_class, COUNT(DISTINCT expected_customer_id) FROM reconcile_customer_ccs_contact_detail GROUP BY attribute_name, comparison_class ORDER BY attribute_name, comparison_class ... PY`
- `uv run python - <<'PY' ... SELECT review_bucket, row_count, merged_customer_count FROM reconcile_customer_ccs_contact_review_pattern_summary ORDER BY row_count DESC, review_bucket ... PY`
- `uv run python - <<'PY' ... SELECT review_order, review_bucket, customer_number, expected_customer_id, comparison_reason FROM reconcile_customer_ccs_contact_review_case_summary ORDER BY review_order ... PY`

### Gotchas

- The existing `finance/workday_extract/` customer match/missing SQL reads `fin_workday_extract_customer`,
  not `src_fin_workday_extract_ccs_customer`; reuse the explode logic, not the dataset choice.
- `docs/finance-customer.md` explicitly treats `reference_id` / `nrm_agg_customer_number` as the
  stable business key and `customer_id` as run-scoped. That distinction must hold in every new
  report.
- A merged customer can legitimately contain multiple source customer numbers. Any count or join
  done only at merged-row grain will hide partial-coverage problems.
- In this repo state, `uv run --package elt-ingest-excel ...` from the workspace root does not work
  because `elt_ingest_excel` is not declared as a root workspace member. Run `uv run python ...`
  from the `elt_ingest_excel/` package root for this module's commands.
- The Workday extract top-level tax fields are comparable, but not all exact-string mismatches mean
  business divergence. The biggest current bucket is wording drift in `transaction_tax_status`, not
  population loss.
- A full outer join between `workday_customer_tax` and `src_fin_workday_extract_ccs_customer`
  counts the dummy Workday row as a field diff if you do not filter to shared `customer_id` values.
  For confirmed sent-vs-extracted field counts, use only `MATCHED_CUSTOMER_ID` rows.
- The direct transform-vs-extract layer is field-oriented, while the provenance layer is
  case-oriented. Do not expect the provenance case count to equal the raw field diff count; it is a
  deliberate condensation of repeated patterns.
- `examples/fin_customer_ccs_reconcile_sql.py` is no longer reconcile-only. It now refreshes the
  CCS Workday extract pipeline before running reconciliation, so use `--run-to-phase INGEST` only
  when you explicitly want to skip the reconcile SQL step.
- The end-to-end runner now queries `reconcile_customer_ccs_report` after the SQL pack finishes and
  will fail fast if any of the expected headline metrics are missing. That is intentional so the
  one-command workflow does not silently succeed with a partial reconcile output.
- When DBeaver has the DuckDB file open, even read-only inspection from the CLI can fail with an
  exclusive-lock error. Close DBeaver before running any CCS reconciliation verification queries.
- If you query verification tables in parallel with the reconciliation runner, you can read a stale
  intermediate state. For fresh counts, rerun the pack first and query the report tables
  sequentially afterwards.
- When constructing multi-line address comparison strings in DuckDB, use `CHR(10)` for line breaks.
- The residual contact-review buckets are deliberately heuristic. They are meant to shrink the
  operator's review surface and highlight candidate fourth-pass rules, not to replace row-level
  review where the bucket is still `manual_review`.
- In DuckDB regex patterns, `\\s` / `\\n` inside a SQL string literal can accidentally search for a
  literal backslash sequence instead of whitespace/newline. For the contact-review bucket rules,
  use `\s` / `\n` in the SQL regex string so the intended bucket actually fires.
  The literal string `'\n'` is not interpreted as a newline and will create a false address
  mismatch inventory.
