# Finance Supplier Pipeline

End-to-end reference for the supplier creditor pipeline. Covers source data, table lineage, dedup logic, output tables, reference tables, and validation. Read this before debugging any issue with a `workday_supplier_*` table or starting a new validation session.

Live supplier reconciliation follow-up is tracked in `elt_ingest_excel/docs/todo/todo_finance_supplier_reconcile.md`.

## DuckDB

- **Database:** `~/Documents/__data/duckdb/rpatel.duckdb`
- **Runtime config:** `config/runtime/default.json`
- DBeaver holds an exclusive lock when open — close it before running any CLI queries.

---

## Source Excel files (INGEST phase)

Ingest config: [config/ingest/finance/active_supplier_creditor.json](../config/ingest/finance/active_supplier_creditor.json)

### Creditor data (source workbook: `Creditor Updates July 16 2026 v2.xlsx`)

Path: `~/Documents/__data/excel/finance_ref/supplier/`

| Sheet | DuckDB table |
|-------|--------------|
| Creditors No Activity 1 Year | `fin_supplier_creditor_no_activity_1_year` |
| Creditors Activity 2 Years | `fin_supplier_creditor_activity_2_years` |
| Ad Hoc Payment Creditors | `fin_supplier_creditor_adhoc_payment` |
| Accrual Suppliers | `fin_supplier_creditor_accrual` |

The sheet name carries a trailing space on `Creditors No Activity 1 Year ` — the ingest
config matches it exactly.

All source columns are loaded as strings. Type coercion and normalisation happens in SQL.

### Workday extract — previously loaded (source workbook: `Extract_Suppliers.xlsx`)

Path: `~/Documents/__data/excel/finance_ref/workday_ccs/` — sheet `Sheet1`, headerRow 13.

| Sheet | DuckDB table |
|-------|--------------|
| Sheet1 | `fin_supplier_creditor_existing` |

Suppliers already loaded into Workday. Feeds `src_fin_supplier_include`, which computes the
suppliers *not* yet loaded.

**That exclusion is currently switched off** — the `WHERE ... IN (SELECT ... FROM
src_fin_supplier_include)` filter in `src_fin_supplier.sql` is commented out, so the
`workday_supplier_*` outputs carry the full source population. It was on during testing to
avoid re-loading existing records.

**`src_fin_supplier_include` still builds regardless**, because it is listed in `order.txt`.
Its result is discarded, but the table must still be creatable: if `fin_supplier_creditor_existing`
is absent from the database, the transform fails at that step. Keep this workbook ingested, or
drop `src_fin_supplier_include.sql` from `order.txt`. An empty extract is fine.

### Reference tables (source workbook: `FIN Master Supplier Workbook Instructions.xlsx`)

Path: `~/Documents/__data/excel/threeplus/Supplier and Customer Master Record workbooks & Instructions Guide/`

These are used by validation SQL to cross-check values in the `workday_supplier_*` output tables against the Workday-allowed values defined in the instructions workbook.

| Sheet | headerRow | DuckDB table | Used by |
|-------|-----------|--------------|---------|
| Address Requirements | 3 | `ref_workday_validate_supplier_address` | `validate_supplier_address.sql` — UNPIVOT-based field requirement checks |
| Currency | 2 | `ref_workday_validate_supplier_currency` | `validate_supplier_currencies.sql` — valid Workday currency codes |
| Bank Account Formatting | **4** | `ref_workday_validate_supplier_bank_format` | `validate_supplier_settlement_account.sql` — IBAN/BIC/routing/account requirements by country |
| Country States-Regions | 3 | `ref_workday_validate_supplier_region` | `validate_supplier_address.sql` — valid region values per country |
| Postal Code Formatting | 3 | `ref_workday_validate_supplier_postal_code` | `validate_supplier_address.sql` — postal code requirement by country |
| Phone Number | 3 | `ref_workday_validate_supplier_phone_format` | `validate_supplier_phone.sql` — valid country names and expected international phone codes |
| Tax ID Types | 3 | `ref_workday_validate_supplier_tax_id_type` | `validate_supplier_tax.sql` — valid tax ID types per country |

**Important quirk — Bank Account Formatting:** The source sheet has a merged header row that leaves the country column blank. Because `headerRow=4`, pandas auto-names this column `unnamed_0`. Validation SQL joins on `LOWER(TRIM(r.unnamed_0)) = LOWER(TRIM(a.bank_country))`.

---

## Staging table lineage (TRANSFORM phase)

SQL files execute in the order defined in [config/transform/sql/finance/supplier/order.txt](../config/transform/sql/finance/supplier/order.txt).

### Stage 1 — `src_fin_supplier_base`

**File:** `src_fin_supplier_base.sql`

Unions the four source tables into a single flat table. Adds a `business_unit` column derived from the source workbook's `company` column. No filtering, no normalisation.

```
fin_supplier_creditor_no_activity_1_year   ─┐
fin_supplier_creditor_activity_2_years     ─┤
fin_supplier_creditor_adhoc_payment        ─┼─▶  src_fin_supplier_base
fin_supplier_creditor_accrual              ─┘
```

All four sheets share an identical 101-column layout, so the union aligns positionally.

---

### Stage 2 — `src_fin_supplier_raw`

**File:** `src_fin_supplier_raw.sql`

Normalises every row. Key derivations:

| Column | Derived from | Notes |
|--------|-------------|-------|
| `nrm_supplier_number` | `TRIM(vendor_id)`, suffixed `-{business_unit}` on detected collision | See collision note below |
| `nrm_supplier_name` | `vendor_name` / `vendor_id` | Uppercased, accents removed, dots removed, FOOTBALL CLUB → FC, parenthesised qualifiers removed e.g. (UK), UNITED → UTD, AND → &, trailing legal suffixes stripped (LIMITED/LTD/COMPANY/PLC/LLP/INC) |
| `nrm_postal_code` | `post_code` | UK postcode format enforcement |
| `nrm_country_code` | Country name/code lookups | Resolves via `ref_source_country_name_mapping` and `ref_source_country_code_mapping` |
| `nrm_business_unit` | `business_unit` | Mapped via `ref_source_business_unit_mapping` (INNER JOIN — unmapped rows are dropped) |
| `nrm_payment_terms_id` | `payment_terms_id` | Mapped via `ref_source_payment_terms` |
| `nrm_bank_*` | Various `eft_*` columns | Sort code, account number, IBAN, SWIFT, routing number |
| `nrm_bank_name` | `ref_bank_sort_code_prefix_mapping` (GB only) → falls back to raw `bank_name` | See bank name note below |
| `nrm_created_date` etc. | `created_date` etc. | Multi-format date parsing via `TRY_STRPTIME` |
| `nrm_subregion_county` | `county` | Computed unconditionally for every country; only exposed downstream where required — see note below |

`data_rnk` is added (ranked within `nrm_supplier_name`, `nrm_supplier_number`) for debug inspection; it is not filtered on downstream.

Rows with unmapped `business_unit` are silently dropped here due to the `INNER JOIN` on `ref_source_business_unit_mapping`.

**`city_subdivision` is required for 5 countries, but only derivable for one of them.**
`ref_workday_validate_supplier_address` requires `city_subdivision` for South Africa, Finland,
Slovakia, Mexico, and Åland Islands — previously this field didn't exist anywhere in the address
transform at all, which was the actual root cause behind Workday's "City Subdivision 1 required"
errors for those countries (not the `city` field itself, which is correctly and separately
suppressed by design for some of them — see the address-suppression note below; don't confuse the
two).

`nrm_city_subdivision` (`src_fin_supplier_raw.sql`) sources it per country: South Africa already
has real data in `address_2` (e.g. `Allen's Nek`, a Johannesburg suburb — kept as-is, no rework
needed, and duplicated rather than moved since Workday marks `address_line_2` `Optional` there,
not `Not Accepted`); everywhere else defaults to raw `Address Line 3`. Which countries actually get
the value exposed is decided in `workday_supplier_address_base.sql` purely from
`ref_workday_validate_supplier_address.city_subdivision = 'Required'` — no hardcoded country list,
so it automatically covers Mexico/Åland Islands too if suppliers from those ever appear.

**Finland/Slovakia — city is the City Subdivision.** For these countries Workday marks
`municipality_city = 'Not Accepted'` **and** `city_subdivision = 'Required'`: the post town/city
(e.g. `Helsinki`) is Workday's own locality field for the address, not a separate municipality.
`Address Line 3` is empty for every affected record and GP is frozen, so
`workday_supplier_address_base.sql` falls back — when `city_subdivision` is required and
`municipality_city = 'Not Accepted'` — to `nrm_city` (`COALESCE` after any real subdivision data).
This is fill-blank-only: it never overrides real subdivision data (South Africa keeps its suburb),
and because `city` is separately blanked for `Not Accepted` municipalities there is no duplication.
Without this the locality would be dropped from the address entirely (`city` and `city_subdivision`
both `NULL`). Table-driven, so it covers any country Workday configures the same way.

**`subregion` (Workday's "Region Subdivision 1") was a genuinely missing field, discovered via a
real customer EIB failure export** (the equivalent supplier failure file isn't ingested in this
repo, but the same schema gap existed here too and was fixed at the same time). Unlike
`city_subdivision`, `nrm_subregion_county` sources from the raw `county` field **unconditionally
for every country** (no country allowlist needed — `county` is already county/subregion-level data
regardless of country). Exposure is still table-driven: only surfaced in
`workday_supplier_address_base.sql` where `ref_workday_validate_supplier_address.subregion_county =
'Required'` (currently Côte d'Ivoire, Estonia, Thailand). Stays `NULL` where the raw `county` field
is itself empty — a genuine source-data gap, not a pipeline issue.

**Bank country/name from a UK sort-code match are trusted regardless of the supplier's
own country.** `ref_bank_sort_code_prefix_mapping` is UK clearing-system reference data
(NatWest/Nationwide/Santander sort-code ranges) with no country column, joined against
every supplier's `eft_bank_code` (first 2 digits) with no country restriction — a
non-UK supplier can genuinely hold a UK bank account, and a match is deliberately
treated as evidence of that (business decision, not a data-quality gate). When it
matches: `nrm_bank_name := rbsc.bank_name_primary` and `nrm_bank_country_name :=
'United Kingdom'` (both computed in `src_fin_supplier_raw.sql`, carried through
`nrm_array_bank`); `workday_supplier_settlement_account.sql` uses
`COALESCE(u.bank.nrm_bank_country_name, t.nrm_country_name)` for `bank_country`, falling
back to the supplier's own resolved country only when there's no match. GP never
populates a bank-country field itself (`bank_country_code`/`central_bank_code` are both
0% populated across all records with a bank account, confirmed) — the sort-code match
is the only source-derived signal available, imperfect as it is: the reference table
covers 88 of 100 possible 2-digit prefixes, so an unrelated non-UK bank code will often
coincidentally match something (confirmed real case: the identical code `600004`
attached to both a US and a Canadian supplier, clearly not the same real account). That
known false-positive risk was assessed and accepted as a deliberate trade-off in favour
of catching genuine non-UK-supplier-with-a-UK-account cases, rather than scoping the
match to the supplier's own country.

A second, independent sort-code source is also tried: a GB-format IBAN embeds the sort
code at characters 9-14 (`GBkk BBBB SSSSSS NNNNNNNN` — 4-letter bank code, 6-digit sort
code, 8-digit account number, 22 characters total). `rbsc_iban` extracts that and joins
it against the same reference table, independently of `eft_bank_code`/`rbsc` — recovers
real accounts (e.g. IBAN `GB21NWBK01026632680481`) that would otherwise have no usable
bank name purely because the sort code only existed inside the IBAN, not in
`eft_bank_code`.

A third tier, `iban_intl`, covers non-UK IBANs: `ref_iban_bank_code_mapping.json`
(loaded via `config/transform/sql/ref/ref_iban_bank_code_mapping.sql`, registered in
`order.txt`) holds bank-code-to-name mappings researched per-country from
official/primary sources (Bundesbank BLZ registry for Germany, Banca d'Italia ABI for
Italy, Banque de France CIB for France, national IBAN/BIC registries for NL/IE/RO/ES/
DK/CH/FI/NO/BE/PL/LU/SK/PT — each row's `source` field cites where it came from and the
retrieval date). The join is via `LEFT JOIN LATERAL`, matching the IBAN's own country
code and prefix-matching its BBAN (from character 5 onward) against each country's own
bank-code length, ordered by longest/most-specific match first — this handles
variable-length codes (e.g. Finland's codes are 1, 2, or 3 characters depending on the
leading digit) without any per-country hardcoded length logic. Two codes could not be
verified from any source and are deliberately left unmapped rather than guessed: Danish
reg.nr. `9227`, Luxembourg bank code `067`.

**IBAN-embedded bank code → `nrm_bank_code_routing_number` (fill-blank-only).** For
several IBAN countries Workday requires a domestic *bank code* (`bank_id_display_option =
'Required'`), yet GP populates no `eft_bank_code`/`eft_transit_routing_no` for those
accounts — only a SWIFT/BIC, from which a domestic code cannot be derived. The national
bank code is, however, embedded in the supplier's own IBAN by ISO 13616. A
`LEFT JOIN LATERAL` (`iban_bankcode`) extracts it using `ref_iban_structure.json` (loaded
via `config/transform/sql/ref/ref_iban_structure.sql`, registered in `order.txt`), which
holds the per-country bank-code offset/length from the SWIFT IBAN Registry (each row cites
its `source`). `nrm_bank_code_routing_number` `COALESCE`s this **after** the
`eft_bank_code`/routing derivation, so a real GP value is **never overwritten** — it fills
only genuinely blank codes. A guard requires the cleaned IBAN to match the country's exact
`iban_length`, so a truncated or placeholder IBAN (e.g. a lone `'0'`) yields nothing rather
than a mangled code, and the account stays flagged for review. Currently covers GE/KZ/UA/CR
(each extracted code cross-checked against its BIC, e.g. Ukraine `305299` = PrivatBank,
BIC `PBANUA2X`); the table extends to further IBAN countries by adding rows only.

When none of the three tiers match, `nrm_bank_name` falls back to the raw `bank_name`
source column, treating a blank value or the literal placeholder string `"Bank Name"`
(GP's own apparent default, ~5,000 rows) as unusable. Since `nrm_bank_name` is a
required field for a bank account to appear in `workday_supplier_settlement_account` at
all (see that file's `WHERE` clause), a supplier with no match from any tier and no
usable raw bank name ends up with no settlement account row — as of this fix, ~71 real
accounts (valid account number or IBAN) remain in this state, needing a source
correction to the raw `bank_name` column at GP.

**A genuine SWIFT/BIC code is never just a sort code.** `nrm_bank_swift_code` nulls out
`swift_address` when it exactly matches (whitespace aside) either `eft_bank_code` or
`eft_transit_routing_no` for the same record — a real SWIFT/BIC is 8-11 alphanumeric
characters (bank + country + location [+ branch]), never a bare numeric sort
code/routing number, so an exact match is a confirmed data-entry error (6 real cases
found), not a genuine BIC.

**vendor_id business-unit collision detection.** GP `vendor_id` is scoped per company book
(`business_unit`), not globally unique — the same code is independently reused for unrelated
vendors in different books (e.g. `HAL003` = three different real companies across FA/NFC/WNSL,
each with its own bank account). If every row sharing a `vendor_id` were collapsed into one
supplier regardless of name, those unrelated companies would be merged into a single composite
Workday supplier record — the wrong address/bank/tax data commingled under one name, and 2 of
every 3 real identities silently disappearing. `cte_supplier_id` detects this (the fully-normalised
`nrm_supplier_name` disagrees for the same `vendor_id`, via `COUNT(DISTINCT nrm_supplier_name)
OVER (PARTITION BY nrm_vendor_id)`) and suffixes `-{business_unit}` onto `nrm_supplier_number`
for those cases only — every other `vendor_id` is untouched. This name-based signal isn't perfect:
see `reconcile_supplier_collision_review` (in `config/reconcile/finance/supplier/`) for every
detected split ranked by name similarity, so likely false positives (same vendor, just a typo or
abbreviation — confirmed examples: `ZENITH VEHICLE CONTRACT` vs `...CONTRACTS`, `PCL
COPORATEWEAR` vs `PCL CORPORATEWEAR`) can be checked before publishing as separate suppliers.

---

### Stage 3 — `src_fin_supplier_id_rnk` *(Level 1 dedup — by supplier ID)*

**File:** `src_fin_supplier_id_rnk.sql`

Groups all raw rows by `nrm_supplier_number` (the vendor-id-collision-aware key from Stage 2 —
already business-unit-suffixed for detected collisions, so this stage only collapses rows that
are genuinely the same vendor). For a vendor_id that appears identically named across business
units, this stage collapses them as before.

For each `nrm_supplier_number`:
- Scores every raw row using a weighted field-presence score (payment terms 100 → city 10)
- Aggregates all addresses into `nrm_array_address`, all bank accounts into `nrm_array_bank`, all names into `nrm_agg_supplier_name`
- `rnk_score = 1` is the highest-scoring raw record for that vendor ID — its `nrm_supplier_name` becomes the **canonical name** for that vendor ID

Useful debug query — `AKE001` is a real vendor_id collision (two unrelated vendors, "AKENHAM" in
FA and "THE AKENHAM PARTNERSHIP" in NFC), so Stage 2 already split it into `AKE001-FA` /
`AKE001-NFC` before this stage ever groups by it:
```sql
-- See what names were merged under a given (possibly business-unit-suffixed) vendor ID
SELECT nrm_supplier_number, nrm_agg_supplier_name, rnk_score
  FROM src_fin_supplier_id_rnk
 WHERE nrm_supplier_number LIKE 'AKE001%';
```

---

### Stage 4 — `src_fin_supplier_rnk` *(Level 2 dedup — by supplier name)*

**File:** `src_fin_supplier_rnk.sql`

Groups by canonical supplier name. Purpose: the same supplier can exist under multiple vendor IDs across business units (e.g. two entities both normalised to "ACME"). This stage collapses them.

How it works:
1. `cte_canonical_name` — for each `nrm_supplier_number`, picks the `nrm_supplier_name` from `rnk_score = 1` in `src_fin_supplier_id_rnk`
2. All rows from `src_fin_supplier_id_rnk` are re-scored and re-grouped by a **bank-aware payee key** (`cte_payee_key`), not by name alone
3. Aggregates all vendor IDs into `nrm_agg_supplier_number`, all addresses, banks, phones, emails, business units
4. `rnk_score = 1` is the highest-scoring record for that payee-key group; `nrm_supplier_name` remains the canonical name (no longer guaranteed unique)

**Bank-aware payee split (`payee_key`).** Grouping by canonical name alone over-collapses:
different GP vendor_ids that merely normalise to the same name (e.g. two unrelated `JOSH SMITH`
payees, or `BT` across `BRI014`/`BTP001`/`BTP002`) were fused into one supplier carrying *both*
parties' bank accounts — a payment-misrouting risk and a source of "multiple active settlement
bank account" EIB rejections. Workday does not require unique supplier names (identity is the
reference ID), so distinct payees may safely stay separate. The key therefore splits a shared
name back out **only when it carries >1 distinct bank account**, keying each row on its GP
vendor_id's representative account (`ARG_MAX(nrm_bank_account_number, score)`):
- A name with ≤1 distinct account is untouched (single group, as before).
- A single `nrm_vendor_id` always stays together (one real supplier may hold several accounts).
- Different vendor_ids that share a representative account still merge (true duplicates, e.g.
  `COMPASS CONTRACT SERVICES` across `COM005`/`COM024`/`COM030`).

Measured effect on the current extract: final supplier records `7581 → 7635`; cross-vendor_id
merges carrying *conflicting* bank accounts `55 → 2` (the 2 residuals genuinely share an account,
so the merge is correct). Row integrity preserved (8089 in = 8089 out).

The two-level dedup guarantees that `src_fin_supplier_rnk` at `rnk_score = 1` is unique by
`payee_key` and by `nrm_supplier_number`. It is **no longer unique by `nrm_supplier_name`** — a
name shared by distinct payees now yields one row per payee (by design).

Useful debug query:
```sql
-- See what vendor IDs and names were merged under a given canonical name
SELECT nrm_supplier_name, nrm_agg_supplier_number, cnt_supplier_name, rnk_score
  FROM src_fin_supplier_rnk
 WHERE nrm_supplier_name LIKE '%AKENHAM%';
```

---

### Stage 5 — `src_fin_supplier` *(final supplier master)*

**File:** `src_fin_supplier.sql`

Takes `rnk_score = 1` from `src_fin_supplier_rnk`, filtered to records with `nrm_payment_terms_id IS NOT NULL`. Assigns a sequential `supplier_id` (`S-000001`, `S-000002`, …) ordered by `nrm_supplier_name`.

This is the **single source of truth** for the consolidated supplier master. All `workday_supplier_*` tables are derived from it.

**`supplier_id` is run-scoped, not a stable identity.** It is a `ROW_NUMBER()` over the
surviving population ordered by supplier name, so any change to that population — a new
extract, an added source sheet, a changed filter — renumbers every supplier sorting after
the change. The same GP supplier will hold different `S-` values in test and in prod.

Consequences:
- Within a run it is coherent: every `workday_supplier_*` table derives from this table, so
  the EIB stitches correctly.
- Across runs or environments it is meaningless. Never compare an `S-` value between runs,
  carry one between environments, or treat it as a Workday identifier.
- `reference_id` (`nrm_agg_supplier_number`, the GP supplier number) is the stable business
  key. Any join to data originating outside this run must use it — as the EIB readback chain
  does (`workday_eib_submit_customer_address_failure_report.sql` joins on `reference_id`,
  then uses the ID only for same-run lookups).

Uniqueness guarantees:
- One row per `payee_key` (canonical name, split by distinct bank account — see Stage 4)
- One row per `nrm_supplier_number` (across all pipe-delimited values in `nrm_agg_supplier_number`)
- **Not** unique by `nrm_supplier_name`: distinct payees sharing a normalised name are kept
  separate on purpose (Workday allows duplicate supplier names — identity is the reference ID)

Vendors excluded here (present in raw but absent from final):
- No payment terms mapping — `nrm_payment_terms_id IS NULL`
- (Earlier exclusion: unmapped business unit dropped at Stage 2)

---

## Output tables (PUBLISH phase)

Each `workday_supplier_*` table maps to one sheet in the Workday EIB workbook. All are derived from `src_fin_supplier` (joined to the arrays/strings it carries).

| Table | Sheet | Key source columns |
|-------|-------|--------------------|
| `workday_supplier_name` | Supplier Name | `supplier_id`, `nrm_supplier_name`, `nrm_supplier_category`, `nrm_payment_terms_id` |
| `workday_supplier_address` | Supplier Address | `nrm_array_address` (unnested) |
| `workday_supplier_settlement_account` | Supplier Settlement Account | `nrm_array_bank` (unnested) |
| `workday_supplier_phone` | Supplier Phone | `nrm_agg_phone_1/2/3` (split) |
| `workday_supplier_email` | Supplier Email | `nrm_agg_email_to_address` (split) |
| `workday_supplier_tax` | Supplier Tax | `nrm_tax_id_number`, `nrm_tax_schedule_id` |
| `workday_supplier_payment` | Supplier Payment | `nrm_payment_terms_id` |
| `workday_supplier_groups` | Supplier Groups | `nrm_array_business_unit` (unnested) |
| `workday_supplier_company_restrictions` | Supplier Company Restrictions | `nrm_agg_business_unit` |
| `workday_supplier_alternate_name` | Supplier Alternate Name | `nrm_agg_supplier_number` (legacy IDs as alternates) |
| `workday_supplier_currencies` | Supplier Currencies | `nrm_currency_code` |
| `workday_supplier_procurement` | Supplier Procurement | PO issue options |
| `workday_supplier_intermediary_account` | Supplier Intermediary Account | Intermediary bank details |

Publish config: [config/publish/finance/publish_supplier.json](../config/publish/finance/publish_supplier.json)
Output workbook template: `FA E2E Suppliers V1.xlsm`

### Known data notes

- `workday_supplier_alternate_name` — the transform SQL uses `WHERE 1=2`, so this table is always empty. The validation file for it is a no-op until the transform is wired.
- `workday_supplier_procurement` — same situation: transform SQL uses `WHERE 1=2`. Always empty.
- `workday_supplier_phone.international_phone_code` — populated via `REGEXP_REPLACE(TRIM(r.phone_code), '[^0-9]', '', 'g')` (digits-only, e.g. `44`). `ref_country.phone_code` has a `+` prefix — this is stripped in the final SELECT of `workday_supplier_phone.sql`. **Re-run the full pipeline after any change to this file to propagate the cleaned values.**
- `workday_supplier_settlement_account.for_supplier_connections_only` — Workday requires exactly
  one unrestricted active settlement account per supplier. For a supplier with genuinely more than
  one distinct bank account (not a vendor_id-collision artifact — those are separated at Stage 2),
  every account beyond the first (ranked by bank name/account number) is set to `'Y'` so it doesn't
  compete with the primary account. This does not set an accepted payment type/currency on the
  secondary account (Workday also allows that as a restriction) — which payment type applies is a
  business decision not derivable from GP data, left for manual assignment.

---

## Validation

### Running validation

From the `elt_ingest_excel/` directory:

```bash
uv run python examples/fin_supplier_creditor_validate_sql.py
```

This executes every `.sql` file in `config/validate/finance/supplier/` in alphabetical order. `init.sql` runs first and creates a fresh `validation_supplier_issues` table.

To query results after running:

```sql
-- Summary by table and rule type
SELECT table_name, rule_type, col, COUNT(*) n
  FROM validation_supplier_issues
 GROUP BY 1, 2, 3
 ORDER BY 1, 2, 3;

-- All issues for a specific supplier
SELECT * FROM validation_supplier_issues WHERE supplier_id = 'S-000001';

-- Count by table
SELECT table_name, COUNT(*) n FROM validation_supplier_issues GROUP BY 1 ORDER BY 1;
```

### `validation_supplier_issues` table schema

| Column | Description |
|--------|-------------|
| `table_name` | The `workday_supplier_*` table the issue was found in |
| `rule_type` | Rule category (see below) |
| `col` | Column(s) the issue relates to |
| `supplier_id` | The `supplier_id` value from the affected row |
| `detail` | The actual value(s) that failed validation |
| `message` | Human-readable description of the issue |

**Rule types used:**

| Rule type | Meaning |
|-----------|---------|
| `CROSS_TAB` | Value must exist in another table (e.g. Supplier ID must be in `workday_supplier_name`) |
| `DUPLICATE` | Value must be unique within the table or a combination |
| `DUPLICATE_COMBO` | Combination of columns must be unique |
| `MISSING` | Required field is NULL or blank |
| `MATCH` | Value must be in an allowed list |
| `FORMAT` | Value must match a format pattern (regex, length, ISO code) |
| `ONE_PRIMARY` | Exactly one row per group must be marked primary |
| `ENSURE_DATA` | If any field in a group is populated, all must be |
| `BANK_ACCOUNT` | Bank-specific cross-field requirements |
| `BANK_FORMAT` | Bank field format requirements by country (IBAN length, BIC, routing) |
| `ADDRESS_REQ` | Address field requirements by country (from `ref_workday_validate_supplier_address`) |

### Vendor ID coverage check

Separate from the EIB sheet validation. Confirms no vendor IDs were lost in the dedup pipeline:

```sql
-- Should return 0 rows
-- File: config/validate/finance/supplier/validate_supplier_id_coverage.sql
SELECT * FROM validation_supplier_issues WHERE rule_type IN ('COVERAGE', 'DUPLICATE');
```

Rule types:
- `COVERAGE` — a `vendor_id` from `src_fin_supplier_base` is missing from `src_fin_supplier` or `workday_supplier_name`
- `DUPLICATE` — a `vendor_id` from `src_fin_supplier_base` maps to more than one row in `src_fin_supplier` or `workday_supplier_name`

### Validation SQL files — rule inventory

All files are in [config/validate/finance/supplier/](../config/validate/finance/supplier/).

**`validate_supplier_name.sql`** — `workday_supplier_name`
- DUPLICATE: `supplier_id`, `supplier_name`, `reference_id` (each must be unique)
- MISSING: `supplier_category` required
- MATCH: `supplier_category` must be in `ref_supplier_category`
- MATCH: `worktag_only` must be Yes/No when populated
- ENSURE_DATA: if `supplier_change_source` set, `supplier_source` must be set
- ENSURE_DATA: if `external_entity_id` set, `supplier_source` must be set

**`validate_supplier_currencies.sql`** — `workday_supplier_currencies`
- CROSS_TAB: `supplier_id` must exist in `workday_supplier_name`
- MISSING: `accepted_currencies_plus` required
- FORMAT: `accepted_currencies_plus` must match `^[A-Z]{3}$` (ISO 4217)
- FORMAT: `default_currency` must match `^[A-Z]{3}$` when populated
- MISSING: every supplier in `workday_supplier_name` must have at least one currencies row
- MATCH: `accepted_currencies_plus` must be a valid Workday currency code (ref: `ref_workday_validate_supplier_currency`)
- MATCH: `default_currency` must be a valid Workday currency code when populated (ref: `ref_workday_validate_supplier_currency`)

**`validate_supplier_address.sql`** — `workday_supplier_address`
- CROSS_TAB: `supplier_id` must exist in `workday_supplier_name`
- DUPLICATE: `address_id` must be unique when populated
- MISSING: `country`, `address_line_1`, `public_flag`, `is_primary` required
- MATCH: `public_flag`, `is_primary` must be Yes/No
- ONE_PRIMARY: exactly one primary per `supplier_id` + `use_for` group
- ADDRESS_REQ: country must exist in `ref_workday_validate_supplier_address` (no unknown countries)
- ADDRESS_REQ: required/recommended fields enforced by country via UNPIVOT on `ref_workday_validate_supplier_address`
- MATCH: `region` must be a valid Workday region for the country (ref: `ref_workday_validate_supplier_region`, joined on `instance` OR `reference_ids`)
- FORMAT: `postal_code` required for countries with a defined validation rule (ref: `ref_workday_validate_supplier_postal_code`)

**`validate_supplier_address_uk.sql`** — diagnostic SELECT (not an INSERT)
- A plain SELECT query (not `INSERT INTO validation_supplier_issues`) — run manually for inspection
- Returns UK (`country_code = 'GB'`) addresses where `region IS NULL OR postal_code IS NULL`
- Joined to `src_fin_supplier` to show the GP supplier number alongside the `supplier_id`

**`validate_supplier_phone.sql`** — `workday_supplier_phone`
- CROSS_TAB: `supplier_id` must exist in `workday_supplier_name`
- DUPLICATE: `phone_id` must be unique
- MISSING: `phone_device_type`, `public_flag`, `primary_flag` required
- MATCH: `public_flag`, `primary_flag` must be Yes/No
- ONE_PRIMARY: exactly one primary per `supplier_id` + `use_for` group
- FORMAT: UK phones must be 9–10 digits (applied to `area_code || phone_number`, digits only)
- FORMAT: Netherlands phones must be 9 digits
- MATCH: `phone_country` must be a recognised Workday country name (ref: `ref_workday_validate_supplier_phone_format`)
- FORMAT: `international_phone_code` must match the expected code for the country (ref: `ref_workday_validate_supplier_phone_format` — digits-only comparison)

**`validate_supplier_tax.sql`** — `workday_supplier_tax`
- CROSS_TAB: `supplier_id` must exist in `workday_supplier_name`
- DUPLICATE: `business_entity_tax_id`, `tax_id` must each be unique when populated
- MATCH: `primary_tax_id`, `transaction_tax_id`, `irs_1099_supplier` must be Yes/No when populated
- ENSURE_DATA: if any of `tax_id`, `tax_id_type`, `primary_tax_id`, `tax_id_country` is set, all four must be set
- MATCH: `tax_id_type` must be a valid type for the country (ref: `ref_workday_validate_supplier_tax_id_type`, matched on `instance` OR `reference_id`)

**`validate_supplier_settlement_account.sql`** — `workday_supplier_settlement_account`
- CROSS_TAB: `supplier_id` must exist in `workday_supplier_name`
- DUPLICATE: `settlement_bank_account_id` must be unique when populated
- MISSING: `bank_country`, `currency`, `bank_account_type` required
- FORMAT: `currency` must match `^[A-Z]{3}$`
- MATCH: `bank_account_type` must be `Checking` or `Savings`
- BANK_ACCOUNT: UK GBP accounts must have `bank_account_number` OR `bank_iban`
- BANK_ACCOUNT: UK GBP accounts with an account number should have a sort code (`bank_code_routing_number`)
- CROSS_TAB: `currency` must match a row in `workday_supplier_currencies` for the same `supplier_id`
- BANK_FORMAT: IBAN required by country (ref: `ref_workday_validate_supplier_bank_format`, joined on `unnamed_0`)
- BANK_FORMAT: IBAN length must match country requirement
- BANK_FORMAT: BIC/SWIFT required by country
- BANK_FORMAT: Bank ID/routing number required by country
- BANK_FORMAT: Account number required by country (waived if IBAN is present)

**`validate_supplier_email.sql`** — `workday_supplier_email`
- CROSS_TAB: `supplier_id` must exist in `workday_supplier_name`
- MISSING: `email_id` required
- DUPLICATE: `email_id`, `email_address` must each be unique
- MISSING: `public_flag`, `primary_flag` required
- MATCH: `public_flag`, `primary_flag` must be Yes/No
- ONE_PRIMARY: exactly one primary per `supplier_id` + `use_for` group
- FORMAT: `email_address` must match a valid email pattern

**`validate_supplier_payment.sql`** — `workday_supplier_payment`
- CROSS_TAB: `supplier_id` must exist in `workday_supplier_name`
- MISSING: `default_payment_type`, `payment_types_accepted` required
- MATCH: `default_payment_type` must be a valid `Payment_Type` named-range value
- MATCH: each pipe-delimited token in `payment_types_accepted` must be a valid `Payment_Type`
- MATCH: `payment_terms` must be a valid value when populated
- CROSS_TAB: if payment type includes EFT/Wire/Direct Deposit, a settlement account must exist

**`validate_supplier_groups.sql`** — `workday_supplier_groups`
- CROSS_TAB: `supplier_id` must exist in `workday_supplier_name`
- MATCH: supplier group name must be a valid named-range value
- DUPLICATE_COMBO: `supplier_id` + `supplier_group` combination must be unique

**`validate_supplier_company_restrictions.sql`** — `workday_supplier_company_restrictions`
- CROSS_TAB: `supplier_id` must exist in `workday_supplier_name`
- MISSING: `restricted_to_companies` required

**`validate_supplier_intermediary_account.sql`** — `workday_supplier_intermediary_account`
- CROSS_TAB: `supplier_id` must exist in `workday_supplier_name`
- MISSING: `intermediary_bank_account_id`, `currency`, `bank_account_type` required
- FORMAT: `currency` must match `^[A-Z]{3}$`
- MATCH: `bank_account_type` must be `Checking` or `Savings`
- CROSS_TAB: `settlement_bank_account_id` must exist in `workday_supplier_settlement_account`

**`validate_supplier_alternate_name.sql`** — `workday_supplier_alternate_name`
- Table is always empty (transform SQL uses `WHERE 1=2`) — all checks are no-ops until the transform is wired

**`validate_supplier_procurement.sql`** — `workday_supplier_procurement`
- Table is always empty (transform SQL uses `WHERE 1=2`) — checks will activate once the transform is wired
- CROSS_TAB: `supplier_id` must exist in `workday_supplier_name`
- CROSS_TAB: if PO issue option = `Email`, supplier must have a row in `workday_supplier_email`

**`validate_supplier_id_coverage.sql`** — pipeline coverage check (not an EIB sheet)
- COVERAGE: vendor IDs from `src_fin_supplier_base` must appear in `src_fin_supplier.nrm_agg_supplier_number`
- COVERAGE: vendor IDs from `src_fin_supplier_base` must have a matching row in `workday_supplier_name` (joined on `reference_id`)
- DUPLICATE: each vendor ID must match exactly one row in `src_fin_supplier`
- DUPLICATE: each vendor ID must match exactly one row in `workday_supplier_name`
- All checks should always return 0 rows

---

## Reconciliation

Full mechanism, gotchas, and cold-start reproduction steps: [reconciliation.md](reconciliation.md).
This section covers only what's supplier-specific.

Run standalone (validate first — the report's GP ID section reads `validation_supplier_issues`):

```bash
uv run python examples/fin_supplier_creditor_validate_sql.py
uv run python examples/fin_supplier_creditor_reconcile_sql.py
```

```sql
SELECT * FROM reconcile_supplier_report ORDER BY section, metric;
```

Files: `config/reconcile/finance/supplier/`.

Supplier-specific notes:

- **Bank/settlement proof exists here only.** `workday_supplier_settlement_account.sql` is active
  (unlike customer's, which is permanently disabled), so `reconcile_supplier_bank_proof.sql` and
  section `3b`/`3f` of the report (bank coverage/row-level proof) have no customer equivalent.
- **`workday_supplier_payment.sql` carries a redundant filter.** It has `WHERE
  nrm_payment_terms_id IS NOT NULL`, but `src_fin_supplier` is already filtered to that condition
  (see `src_fin_supplier.sql`), so it never actually drops a row in practice — coverage is still
  effectively a 1:1 passthrough, same as customer.
- **Phone aggregation was fixed to match customer's pattern.** `workday_supplier_phone.sql`
  previously read the single-row scalar `phone_number_1/2/3`/`fax_number` columns instead of the
  aggregated `nrm_agg_phone_1/2/3` columns already present on `src_fin_supplier` — meaning phone
  numbers captured from merged duplicate GP records never reached the published table. Fixed to
  read the aggregate, mirroring `workday_customer_phone.sql`. The remaining phone shortfall is
  correctly-excluded garbage input, not an aggregation gap.
- **Fax numbers are deliberately excluded**, per business decision — Workday's `phone_device_type`
  field has no valid "Fax" value (the prior EIB load rejected the one supplier that had one,
  Nordomatic UK). `workday_supplier_phone.sql` no longer reads `nrm_agg_fax` at all.
- **Supplier phone source data has a garbage-zero-padding pattern customer's doesn't** — many raw
  values are a valid number with `0000` (or similar) appended, likely an Excel numeric-precision
  export artifact (e.g. `75239943970000` → the transform correctly recovers `7523994397`). See the
  prefix-containment gotcha in [reconciliation.md](reconciliation.md) — this is why the supplier
  phone proof query needs prefix matching, not just suffix matching, to avoid false "not found"
  results for numbers the transform actually handled correctly.

---

## Scoring weights

Used at both Level 1 (by ID) and Level 2 (by name) to pick the best representative record:

| Field | Weight |
|-------|--------|
| `nrm_payment_terms_id` | 100 |
| `nrm_bank_name` | 90 |
| `nrm_tax_schedule_id = 'PS20'` | 80 |
| `nrm_tax_id_number` | 70 |
| `nrm_region` | 60 |
| `nrm_postal_code` | 50 |
| `nrm_address_line_1` | 40 |
| `nrm_address_line_2` | 30 |
| `nrm_address_line_3` | 20 |
| `nrm_address_line_4` | 15 |
| `nrm_city` | 10 |
