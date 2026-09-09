# Finance Customer Pipeline

**Single source of truth for the customer-debtor + customer-invoice OPENING-BALANCES MIGRATION (Workday go-live, ONE-OFF, not BAU).**

> ⚠ **CONTEXT:** This is a PURE ONE-OFF migration.  Customer master (`workday_customer_*`),
> customer-debtor records, and customer invoices / credit memos are loaded into Workday ONCE at
> go-live.  Source data was statically extracted on one date and will NOT be refreshed; there is
> no on-going BAU pipeline here.  The EIB workbooks produced by this pipeline were
> `AR Conversion` / `Open AR Conversion` / `AP Conversion` line-type only — never recurring
> customer billing runs.  Accordingly, *defensive* code patterns (e.g. casting every numeric
> cell to DOUBLE in the normalisation layer even though the current source data is perfectly
> clean) are unnecessary.  Outputs have been signed off.  Do NOT re-run unless a new migration
> extract has been supplied by Finance/3+.

End-to-end reference for the customer debtor pipeline. Covers source data, table lineage, dedup logic, output tables, reference tables, and validation. Read this before debugging any issue with a `workday_customer_*` table or starting a new validation session.

## DuckDB

- **Database:** `~/Documents/__data/duckdb/rpatel.duckdb`
- **Runtime config:** `config/runtime/default.json`
- DBeaver holds an exclusive lock when open — close it before running any CLI queries.

---

## Source Excel files (INGEST phase)

Ingest config: [config/ingest/finance/active_customer_debtor.json](../config/ingest/finance/active_customer_debtor.json)

### Debtor data (source workbook: `Multicompany Debtors No Activity 1 Year & Activity 2 Years Excluding 55 Debtors & INACTIVE.xlsx`)

Path: `~/Documents/__data/excel/finance_ref/customer/`

| Sheet | DuckDB table |
|-------|--------------|
| Debtors No Activity 1 Year | `fin_customer_debtor_no_activity_1_year` |
| Debtors With Activity 2 Years | `fin_customer_debtor_activity_2_years` |

The workbook's third sheet, `All Parent Debtors`, is deliberately not ingested: it is a
different population cut with a different column set (59 vs 61), not the migration set.

This workbook is replaced in place on each refresh — the filename does not carry a date,
so neither the config nor git records which vintage a given run used.

All source columns are loaded as strings. Type coercion and normalisation happens in SQL.

### Workday extract — previously loaded (source workbook: `Extract_Customers.xlsx`)

Path: `~/Documents/__data/excel/finance_ref/workday_ccs/` — sheet `Sheet1`, headerRow 11.

| Sheet | DuckDB table |
|-------|--------------|
| Sheet1 | `fin_workday_extract_customer` |

Customers already loaded into Workday. Feeds `src_fin_customer_include`, which computes the
customers *not* yet loaded.

**That exclusion is currently switched off** — the `WHERE ... IN (SELECT ... FROM
src_fin_customer_include)` filter in `src_fin_customer.sql` is commented out, so the
`workday_customer_*` outputs carry the full source population. It was on during testing to
avoid re-loading existing records.

**`src_fin_customer_include` still builds regardless**, because it is listed in `order.txt`.
Its result is discarded, but the table must still be creatable: if `fin_workday_extract_customer`
is absent from the database, the transform fails at that step. Keep this workbook ingested, or
drop `src_fin_customer_include.sql` from `order.txt`. An empty extract is fine.

### Reference tables (source workbook: `FIN Master Customer Workbook Instructions.xlsx`)

Path: `~/Documents/__data/excel/threeplus/Supplier and Customer Master Record workbooks & Instructions Guide/`

These are used by validation SQL to cross-check values in the `workday_customer_*` output tables against the Workday-allowed values defined in the instructions workbook.

| Sheet | headerRow | DuckDB table | Used by |
|-------|-----------|--------------|---------|
| Address Requirements | 3 | `ref_workday_validate_customer_address` | `validate_customer_address.sql` — UNPIVOT-based field requirement checks |
| Currency | 2 | `ref_workday_validate_customer_currency` | `validate_customer_currencies.sql` — valid Workday currency codes |
| Bank Account Formatting | **4** | `ref_workday_validate_customer_bank_format` | `validate_customer_settlement_account.sql` — IBAN/BIC/routing/account requirements by country |
| Country States-Regions | 3 | `ref_workday_validate_customer_region` | `validate_customer_address.sql` — valid region values per country |
| Postal Code Formatting | 3 | `ref_workday_validate_customer_postal_code` | `validate_customer_address.sql` — postal code requirement by country |
| Phone Number | 3 | `ref_workday_validate_customer_phone_format` | `validate_customer_phone.sql` — valid country names and expected international phone codes |
| Tax ID Types | 3 | `ref_workday_validate_customer_tax_id_type` | `validate_customer_tax.sql` — valid tax ID types per country |

**Important quirk — Bank Account Formatting:** The source sheet has a merged header row that leaves the country column blank. Because `headerRow=4`, pandas auto-names this column `unnamed_0`. Validation SQL joins on `LOWER(TRIM(r.unnamed_0)) = LOWER(TRIM(a.country))`.

---

## Staging table lineage (TRANSFORM phase)

SQL files execute in the order defined in [config/transform/sql/finance/customer/order.txt](../config/transform/sql/finance/customer/order.txt).

### Stage 1 — `src_fin_customer_base`

**File:** `src_fin_customer_base.sql`

Unions the two source tables into a single flat table. Adds a `business_unit` column derived from the source workbook's `company` column. No filtering, no normalisation.

```
fin_customer_debtor_no_activity_1_year   ─┐
fin_customer_debtor_activity_2_years     ─┴─▶  src_fin_customer_base
```

---

### Stage 2 — `src_fin_customer_raw`

**File:** `src_fin_customer_raw.sql`

Normalises every row. Key derivations:

| Column | Derived from | Notes |
|--------|-------------|-------|
| `nrm_customer_number` | `TRIM(customer_number)` | The cleaned customer number |
| `nrm_customer_name` | `customer_name` / `customer_number` | Uppercased, accents removed, dots removed, FOOTBALL CLUB → FC, parenthesised qualifiers removed e.g. (UK), UNITED → UTD, AND → &, trailing legal suffixes stripped (LIMITED/LTD/COMPANY/PLC/LLP/INC) |
| `nrm_postal_code` | `post_code` | UK postcode format enforcement |
| `nrm_country_code` | Country name/code lookups | Resolves via `ref_source_country_name_mapping` and `ref_source_country_code_mapping` |
| `nrm_business_unit` | `business_unit` | Mapped via `ref_source_business_unit_mapping` (INNER JOIN — unmapped rows are dropped) |
| `nrm_payment_terms_id` | `payment_terms_id` | Mapped via `ref_source_payment_terms` |
| `nrm_customer_category` | `customer_class` | Mapped via `ref_customer_category`; defaults to `Miscellaneous` when unmapped |
| `nrm_tax_registration_number` | `tax_registration_number` | Numeric/alphanumeric only, blanked if non-numeric |
| `nrm_created_date` etc. | `created_date` etc. | Multi-format date parsing via `TRY_STRPTIME` |
| `nrm_address_line_1/2/3` | `address_1/2/3` | Deduplicated list via `LIST_DISTINCT`; special chars removed |
| `nrm_address_line_4` | — | Always `NULL` — customer source has only three address lines |
| `nrm_city_subdivision` | `address_2` (ZA) / `address_3` (FI/SK/MX/AX) | Only computed for the 5 countries Workday requires it for — see note below |
| `nrm_subregion_county` | `county` | Computed unconditionally for every country; only exposed downstream where required — see note below |

`data_rnk` is added (ranked within `nrm_customer_name`, `nrm_customer_number`) for debug inspection; it is not filtered on downstream.

Rows with unmapped `business_unit` are silently dropped here due to the `INNER JOIN` on `ref_source_business_unit_mapping`.

**`city_subdivision` is required for 5 countries, but only derivable for one of them** (ported from
the equivalent supplier-side fix — see `docs/finance-supplier.md`). `ref_workday_validate_customer_address`
requires `city_subdivision` for South Africa, Finland, Slovakia, Mexico, and Åland Islands.
`nrm_city_subdivision` (`src_fin_customer_raw.sql`) sources it per country: South Africa already
has real data in `address_2` (e.g. `Clifton`, a Cape Town suburb — kept as-is); everywhere else
defaults to raw `Address Line 3`, currently empty for the 2 live Finland/Slovakia records. Which
countries actually get the value exposed is decided in `workday_customer_address_base.sql` purely
from `ref_workday_validate_customer_address.city_subdivision = 'Required'` — no hardcoded country
list. Finland and Slovakia stay `NULL` until GP/finance perform the same one-time `Address Line 3`
data migration requested on the supplier side (separate ask, same record shape — not tracked in
this repo).

**`subregion` (Workday's "Region Subdivision 1") was a genuinely missing field, discovered via a
real EIB failure export, not ported from supplier** — both domains had the same gap (`CAST(NULL AS
VARCHAR) subregion` hardcoded in `workday_customer_address_base.sql`/`workday_supplier_address_base.sql`)
and were fixed together. Unlike `city_subdivision`, `nrm_subregion_county` sources from the raw
`county` field **unconditionally for every country** (no country allowlist needed — `county` is
already county/subregion-level data regardless of country) via `src_fin_customer_raw.sql`.
Exposure is still table-driven: only surfaced in `workday_customer_address_base.sql` where
`ref_workday_validate_customer_address.subregion_county = 'Required'` (currently Côte d'Ivoire,
Estonia, Thailand). Stays `NULL` where the raw `county` field is itself empty (e.g. one Thai
customer with no county on file) — a genuine source-data gap, not a pipeline issue.

---

### Stage 3 — `src_fin_customer_id_rnk` *(Level 1 dedup — by customer number)*

**File:** `src_fin_customer_id_rnk.sql`

Groups all raw rows by `nrm_customer_number`. Purpose: the same customer number can appear under different customer names across business units. This stage collapses them.

For each `nrm_customer_number`:
- Scores every raw row using a weighted field-presence score (payment terms 100 → city 10)
- Aggregates all addresses into `nrm_array_customer_address`, phones into `nrm_agg_phone_1/2/3`, emails into `nrm_agg_email_to/cc/bcc_address`, all names into `nrm_agg_customer_name`
- `rnk_score = 1` is the highest-scoring raw record for that customer number — its `nrm_customer_name` becomes the **canonical name** for that customer number

Useful debug query:
```sql
-- See what names were merged under a given customer number
SELECT nrm_customer_number, nrm_agg_customer_name, rnk_score
  FROM src_fin_customer_id_rnk
 WHERE nrm_customer_number = 'CUS001';
```

---

### Stage 4 — `src_fin_customer_rnk` *(Level 2 dedup — by customer name)*

**File:** `src_fin_customer_rnk.sql`

Groups by canonical customer name. Purpose: the same customer can exist under multiple customer numbers across business units. This stage collapses them.

How it works:
1. `cte_canonical_name` — for each `nrm_customer_number`, picks the `nrm_customer_name` from `rnk_score = 1` in `src_fin_customer_id_rnk`
2. All rows from `src_fin_customer_id_rnk` are re-scored and re-grouped by that canonical name
3. Aggregates all customer numbers into `nrm_agg_customer_number`, all addresses, phones, emails, business units
4. `rnk_score = 1` is the highest-scoring record for that canonical name group

The two-level dedup guarantees that `src_fin_customer_rnk` at `rnk_score = 1` is unique by both `nrm_customer_name` and `nrm_customer_number`.

Useful debug query:
```sql
-- See what customer numbers and names were merged under a given canonical name
SELECT nrm_customer_name, nrm_agg_customer_number, cnt_customer_name, rnk_score
  FROM src_fin_customer_rnk
 WHERE nrm_customer_name LIKE '%ACME%';
```

---

### Stage 5 — `src_fin_customer` *(final customer master)*

**File:** `src_fin_customer.sql`

Takes `rnk_score = 1` from `src_fin_customer_rnk`. Assigns a sequential `customer_id`
(`C-000001`, `C-000002`, …) ordered by `nrm_customer_name`.

Unlike the supplier chain, there is **no `nrm_payment_terms_id IS NOT NULL` filter here** —
`src_fin_customer_raw.sql` defaults blank payment terms to `'30 DAYS'` via `COALESCE`, so the
value is never null and nothing can be dropped by it. (`lgcy_src_fin_customer.sql` does carry
that filter, but it is not in `order.txt` and does not run.)

This is the **single source of truth** for the consolidated customer master. All `workday_customer_*` tables are derived from it.

**`customer_id` is run-scoped, not a stable identity.** It is a `ROW_NUMBER()` over the
surviving population ordered by customer name, so any change to that population — a new
extract, an added source sheet, a changed filter — renumbers every customer sorting after
the change. The same GP customer will hold different `C-` values in test and in prod.

Consequences:
- Within a run it is coherent: every `workday_customer_*` table derives from this table, so
  the EIB stitches correctly.
- Across runs or environments it is meaningless. Never compare a `C-` value between runs,
  carry one between environments, or treat it as a Workday identifier.
- `reference_id` (`nrm_agg_customer_number`, the GP customer number) is the stable business
  key. Any join to data originating outside this run must use it — as the EIB readback chain
  does (`workday_eib_submit_customer_address_failure_report.sql` joins on `reference_id`,
  then uses `customer_id` only for same-run lookups).

Uniqueness guarantees:
- One row per `nrm_customer_name`
- One row per `nrm_customer_number` (across all pipe-delimited values in `nrm_agg_customer_number`)

Customers excluded here (present in raw but absent from final):
- No payment terms mapping — `nrm_payment_terms_id IS NULL`
- (Earlier exclusion: unmapped business unit dropped at Stage 2)

---

## Output tables (PUBLISH phase)

Each `workday_customer_*` table maps to one sheet in the Workday EIB workbook. All are derived from `src_fin_customer`.

| Table | Sheet | Key source columns |
|-------|-------|--------------------|
| `workday_customer_name` | Customer Name | `customer_id`, `nrm_customer_name`, `nrm_customer_category`, `nrm_agg_customer_number` |
| `workday_customer_groups` | Customer Groups | `nrm_customer_name` initial letter grouping + `ref_workday_group_additional` |
| `workday_customer_credit` | Customer Credit | Disabled — `WHERE 1=2` |
| `workday_customer_payment` | Customer Payment | `nrm_payment_terms_id`, `nrm_currency_code` |
| `workday_customer_currencies` | Customer Currencies | `nrm_currency_code` |
| `workday_customer_company_restrictions` | Customer Company Restrictions | Disabled — `WHERE 1=2` |
| `workday_customer_document_delivery` | Customer Document Delivery | Derived from `workday_customer_email` (active) |
| `workday_customer_settlement_account` | Customer Settlement Account | Disabled — `WHERE 1=2` |
| `workday_customer_intermediary_account` | Customer Intermediary Account | Disabled — `WHERE 1=2` |
| `workday_customer_tax` | Customer Tax | `nrm_tax_registration_number`, `nrm_tax_schedule_id`, `nrm_country_code` |
| `workday_customer_address` | Customer Address | `nrm_array_customer_address` (unnested) |
| `workday_customer_phone` | Customer Phone | `nrm_agg_phone_1/2/3` (split + parsed); fax excluded — see note below |
| `workday_customer_email` | Customer Email | `nrm_agg_email_to/cc/bcc_address` (split) |

Output workbook template: `FA E2E Customers V1.xlsm`  
Publish config: [config/publish/finance/publish_customer.json](../config/publish/finance/publish_customer.json)  
Output filename prefix: `workday_fin_debtor_customer_active_v1`

### Known data notes

- `workday_customer_credit`, `workday_customer_company_restrictions`, `workday_customer_settlement_account`, `workday_customer_intermediary_account`, `workday_customer_credit_card`, `workday_customer_bank_account_address`, `workday_customer_related_worktags` — transform SQL uses `WHERE 1=2`, always empty. Validation files exist for completeness and will activate once the transform is wired.
- `workday_customer_phone.international_phone_code` — populated via `REGEXP_REPLACE(TRIM(r.phone_code), '[^0-9]', '', 'g')` (digits-only, e.g. `44`). `ref_country.phone_code` has a `+` prefix — this is stripped in the final SELECT of `workday_customer_phone.sql`. **Re-run the full pipeline after any change to this file to propagate the cleaned values.**
- `workday_customer_address.last_modified` — always NULL. Unlike supplier, this field is not populated from `nrm_created_date` even though that column is available in `src_fin_customer`.
- `workday_customer_address` GB filter — the postal_code/region completeness filter for GB addresses is commented out (same approach as supplier). UK addresses missing postal code or region are included and will surface in `validate_customer_address_report.sql`.
- `nrm_address_line_4` — always NULL in `src_fin_customer_raw` (customer source has only three address lines). The field propagates through the pipeline and appears as NULL in `workday_customer_address`.
- Public/primary flag naming — customer output tables use `is_public` (not `public_flag`) for the public visibility column. Validation SQL references `is_public` accordingly.

---

## Validation

### Running validation

Validation runs automatically as part of the main pipeline for any phase beyond INGEST:

```bash
uv run python examples/fin_customer_debtor.py
```

To run validation standalone against an existing database without re-running the pipeline:

```bash
uv run python examples/fin_customer_debtor_validate_sql.py
uv run python examples/fin_customer_debtor_validate_sql.py --database-path /path/to/db.duckdb
```

Both routes execute every `.sql` file in `config/validate/finance/customer/` in alphabetical order. `init.sql` runs first and creates a fresh `validation_customer_issues` table.

To query results after running:

```sql
-- Summary by table and rule type
SELECT table_name, rule_type, col, COUNT(*) n
  FROM validation_customer_issues
 GROUP BY 1, 2, 3
 ORDER BY 1, 2, 3;

-- All issues for a specific customer
SELECT * FROM validation_customer_issues WHERE customer_id = 'C-000001';

-- Count by table
SELECT table_name, COUNT(*) n FROM validation_customer_issues GROUP BY 1 ORDER BY 1;
```

### `validation_customer_issues` table schema

| Column | Description |
|--------|-------------|
| `table_name` | The `workday_customer_*` table the issue was found in |
| `rule_type` | Rule category (see below) |
| `col` | Column(s) the issue relates to |
| `customer_id` | The `customer_id` value from the affected row |
| `detail` | The actual value(s) that failed validation |
| `message` | Human-readable description of the issue |

**Rule types used:**

| Rule type | Meaning |
|-----------|---------|
| `CROSS_TAB` | Value must exist in another table (e.g. Customer ID must be in `workday_customer_name`) |
| `DUPLICATE` | Value must be unique within the table |
| `DUPLICATE_COMBO` | Combination of columns must be unique |
| `MISSING` | Required field is NULL or blank |
| `MATCH` | Value must be in an allowed list |
| `FORMAT` | Value must match a format pattern (regex, length, ISO code) |
| `ONE_PRIMARY` | Exactly one row per group must be marked primary |
| `ENSURE_DATA` | If any field in a group is populated, all must be |
| `BANK_ACCOUNT` | Bank-specific cross-field requirements |
| `BANK_FORMAT` | Bank field format requirements by country (IBAN length, BIC, routing) |
| `ADDRESS_REQ` | Address field requirements by country (from `ref_workday_validate_customer_address`) |
| `COVERAGE` | Customer number from source is missing from the final master table |

### Customer number coverage check

Confirms no customer numbers were lost in the dedup pipeline. Unlike the supplier check, this only covers customers that had a payment terms ID in the source (the same filter applied at Stage 5):

```sql
-- Should return 0 rows
SELECT * FROM validation_customer_issues WHERE rule_type = 'COVERAGE';
```

### Validation SQL files — rule inventory

All files are in [config/validate/finance/customer/](../config/validate/finance/customer/).

**`validate_customer_id_coverage.sql`** — pipeline coverage check (not an EIB sheet)
- COVERAGE: customer numbers from `src_fin_customer_base` (with `payment_terms_id`) must appear in `src_fin_customer.nrm_agg_customer_number`
- COVERAGE: customer numbers from `src_fin_customer_base` must have a matching row in `workday_customer_name` (joined on `reference_id`)
- DUPLICATE: each customer number must match exactly one row in `src_fin_customer`
- DUPLICATE: each customer number must match exactly one row in `workday_customer_name`
- All checks should always return 0 rows

**`validate_customer_name.sql`** — `workday_customer_name`
- DUPLICATE: `customer_id`, `customer_name`, `reference_id` (each must be unique)
- MISSING: `customer_category` required
- MATCH: `customer_category` must be in `ref_customer_category`
- MATCH: `worktag_only` must be Yes/No when populated

**`validate_customer_currencies.sql`** — `workday_customer_currencies`
- CROSS_TAB: `customer_id` must exist in `workday_customer_name`
- MISSING: `accepted_currencies_plus` required
- FORMAT: `accepted_currencies_plus` must match `^[A-Z]{3}$` (ISO 4217)
- FORMAT: `default_currency` must match `^[A-Z]{3}$` when populated
- MISSING: every customer in `workday_customer_name` must have at least one currencies row
- MATCH: `accepted_currencies_plus` must be a valid Workday currency code (ref: `ref_workday_validate_customer_currency`)
- MATCH: `default_currency` must be a valid Workday currency code when populated

**`validate_customer_address.sql`** — `workday_customer_address`
- CROSS_TAB: `customer_id` must exist in `workday_customer_name`
- DUPLICATE: `address_id` must be unique when populated
- MISSING: `country`, `address_line_1`, `is_public`, `is_primary` required
- MATCH: `is_public`, `is_primary` must be Yes/No
- ONE_PRIMARY: exactly one primary per `customer_id` + `use_for` group
- ADDRESS_REQ: country must exist in `ref_workday_validate_customer_address`
- ADDRESS_REQ: required/recommended fields enforced by country via UNPIVOT on `ref_workday_validate_customer_address`
- MATCH: `region` must be a valid Workday region for the country (ref: `ref_workday_validate_customer_region`, joined on `instance` OR `reference_ids`)
- FORMAT: `postal_code` required for countries with a defined validation rule (ref: `ref_workday_validate_customer_postal_code`)

**`validate_customer_address_report.sql`** — diagnostic SELECT (not an INSERT)
- A plain SELECT query (not `INSERT INTO validation_customer_issues`) — run manually for inspection
- Returns addresses flagged in `validation_customer_issues` for `workday_customer_address` where `region IS NULL OR postal_code IS NULL`
- Joined to `src_fin_customer` to show the GP customer number alongside the `customer_id`

**`validate_customer_phone.sql`** — `workday_customer_phone`
- CROSS_TAB: `customer_id` must exist in `workday_customer_name`
- DUPLICATE: `phone_id` must be unique
- MISSING: `phone_device_type`, `is_public`, `primary_flag` required
- MATCH: `is_public`, `primary_flag` must be Yes/No
- ONE_PRIMARY: exactly one primary per `customer_id` + `use_for` group
- FORMAT: UK phones must be 9–10 digits (applied to `area_code || phone_number`, digits only)
- FORMAT: Netherlands phones must be 9 digits
- MATCH: `phone_country` must be a recognised Workday country name (ref: `ref_workday_validate_customer_phone_format`)
- FORMAT: `international_phone_code` must match the expected code for the country (digits-only comparison)

**`validate_customer_tax.sql`** — `workday_customer_tax`
- CROSS_TAB: `customer_id` must exist in `workday_customer_name`
- DUPLICATE: `tax_id` must be unique when populated
- MATCH: `primary_tax_id`, `transaction_tax_id` must be Yes/No when populated
- ENSURE_DATA: if any of `tax_id`, `tax_id_type`, `primary_tax_id`, `tax_id_country` is set, all four must be set
- MATCH: `tax_id_type` must be a valid type for the country (ref: `ref_workday_validate_customer_tax_id_type`, matched on `instance` OR `reference_id`)

**`validate_customer_settlement_account.sql`** — `workday_customer_settlement_account`
- Table is disabled (`WHERE 1=2`) — all checks produce 0 rows until the transform is wired
- Column names differ from supplier: `country` (not `bank_country`), `iban` (not `bank_iban`), `swift_bank_identification_code` (not `bank_swift_code`), `routing_transit_number` (not `bank_code_routing_number`)
- CROSS_TAB, DUPLICATE, MISSING, FORMAT, MATCH, BANK_ACCOUNT, BANK_FORMAT rules all present for future use

**`validate_customer_email.sql`** — `workday_customer_email`
- CROSS_TAB: `customer_id` must exist in `workday_customer_name`
- MISSING: `email_id` required
- DUPLICATE: `email_id`, `email_address` must each be unique
- MISSING: `is_public`, `primary_flag` required
- MATCH: `is_public`, `primary_flag` must be Yes/No
- ONE_PRIMARY: exactly one primary per `customer_id` + `use_for` group
- FORMAT: `email_address` must match a valid email pattern

**`validate_customer_payment.sql`** — `workday_customer_payment`
- CROSS_TAB: `customer_id` must exist in `workday_customer_name`
- MISSING: `default_payment_type` required
- MATCH: `default_payment_type` must be a valid `Payment_Type` named-range value
- MATCH: `payment_terms` must be a valid value when populated
- Note: no `payment_types_accepted` pipe-delimited token check (customer payment table carries only `default_payment_type`)

**`validate_customer_groups.sql`** — `workday_customer_groups`
- CROSS_TAB: `customer_id` must exist in `workday_customer_name`
- MATCH: customer group name must be a valid named-range value (`Customers 0 - 9`, `Customers A - L`, `Customers M - Z`, plus additional from `ref_workday_group_additional` where `source_type = 'customer'`)
- DUPLICATE_COMBO: `customer_id` + `customer_group` combination must be unique

**`validate_customer_company_restrictions.sql`** — `workday_customer_company_restrictions`
- Table is disabled (`WHERE 1=2`) — checks are no-ops until the transform is wired
- CROSS_TAB: `customer_id` must exist in `workday_customer_name`
- MISSING: `restricted_to_companies` required

**`validate_customer_intermediary_account.sql`** — `workday_customer_intermediary_account`
- Table is disabled (`WHERE 1=2`) — checks are no-ops until the transform is wired
- CROSS_TAB, MISSING, FORMAT, MATCH, CROSS_TAB (settlement link) rules all present for future use

---

## Reconciliation

Full mechanism, gotchas, and cold-start reproduction steps: [reconciliation.md](reconciliation.md).
This section covers only what's customer-specific.

Run standalone (validate first — the report's GP ID section reads `validation_customer_issues`):

```bash
uv run python examples/fin_customer_debtor_validate_sql.py
uv run python examples/fin_customer_debtor_reconcile_sql.py
```

```sql
SELECT * FROM reconcile_customer_report ORDER BY section, metric;
```

Files: `config/reconcile/finance/customer/`.

Customer-specific notes:

- **No bank/settlement proof.** `workday_customer_settlement_account.sql` is permanently disabled
  (`WHERE 1=2`), so unlike supplier there is no `reconcile_customer_bank_proof.sql` — nothing to
  reconcile for a table that never populates by design.
- **No payment filter to account for.** `workday_customer_payment.sql` has no `WHERE
  nrm_payment_terms_id IS NOT NULL` clause (supplier's does, redundantly) — coverage is a plain
  1:1 passthrough from `src_fin_customer`, nothing else to note.
- **Phone parses correctly by design** — `workday_customer_phone.sql` reads the aggregated
  `nrm_agg_phone_1/2/3` columns directly (unlike an earlier version of the supplier
  transform, which read the single-row scalar columns instead and lost numbers from merged
  duplicate records — since fixed to match this pattern). The remaining phone shortfall here is
  correctly-excluded garbage input (extensions, placeholder text, malformed fragments), not an
  aggregation gap.
- **Fax excluded from `workday_customer_phone`, ported from the supplier-side fix**: Workday's
  `phone_device_type` field has no valid `"Fax"` value, so `cte_customer_phone`'s fax-extraction
  branch (which previously sourced `nrm_agg_fax`) was removed, along with the now-dead `'fax'`/
  `'Fax'` literals in the ranking `CASE` and `ORDER BY`. `nrm_agg_fax` still exists in
  `src_fin_customer_id_rnk.sql`/`src_fin_customer_rnk.sql` (unused downstream) — same as supplier.
  No live records were affected (0 fax rows before or after).

---

## Scoring weights

Used at both Level 1 (by customer number) and Level 2 (by name) to pick the best representative record:

| Field | Weight |
|-------|--------|
| `nrm_payment_terms_id` | 100 |
| `nrm_tax_schedule_id = 'SS20'` | 90 |
| `nrm_tax_schedule_id` (any other value) | 80 |
| `nrm_tax_registration_number` | 70 |
| `nrm_region` | 60 |
| `nrm_postal_code` | 50 |
| `nrm_address_line_1` | 40 |
| `nrm_address_line_2` | 30 |
| `nrm_address_line_3` | 20 |
| `nrm_city` | 10 |

Note: unlike the supplier pipeline, there is no bank-related scoring weight — customer bank aggregation is not part of the active pipeline.
