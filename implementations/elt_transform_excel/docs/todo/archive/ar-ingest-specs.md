# Anchor Doc — AR Customer Template Operating Contract

## Resume (start here)

- **Status:** `ar_customer` is complete for current scope and ready for Finance AR use.
- **Current rendered state:** `ar_customer.xlsx` currently renders as **43 columns**, **22 required**,
  **18 pre-filled defaults**, **27 drop-downs**, and **16 lookup tabs**. Treat the renderer output as
  the live source of truth for these counts.
- **Scope:** this document now covers the live operating contract for the AR customer input template
  only. Invoice / credit-note template work is out of scope. The implementation trail and old
  milestone log now live in [ar-ingest-specs-history.md](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/docs/todo/archive/ar-ingest-specs-history.md).
- **Current work from here:** hand the template to AR, gather UX feedback, and make any agreed
  refinements in config, then re-render.
- **Current operating rule:** do not change any column's `feeds` / `drives` / `required` / `default`
  unless explicitly agreed with the operator.
- **Read before touching code:** `config/ingest_specs/ar_customer.yaml` and
  `docs/CONFIGURATION.md` § "Ingest Specs".

## Session start prompt

```text
load the use-context skill, and from: elt_transform_excel/docs/todo/archive/ar-ingest-specs.md, continue
```

## Goal

Give the Finance AR team a single standard template workbook to populate by hand — **customers** —
and provide a safe Excel UX (defaults + validated picklists + cascades) to minimise manual-entry errors.

## Current contract

- `ar_customer` is the only in-scope AR input template.
- `ar_customer_invoice.yaml` and `ar_customer_credit_note.yaml` are legacy placeholders and must not be
  progressed unless the operator explicitly reopens that scope.
- `ar_customer` field coverage is complete against the populated Customer EIB outputs currently used by
  the migration flow.
- The workbook is a deterministic render of config: `config/ingest_specs/ar_customer.yaml` plus
  `config/ingest_lookups/*.json` drive `dist/ingest_templates/ar_customer.xlsx`.
- The renderer is generic; business/domain values belong in config/reference files, not in Python.

## Source hierarchy

- Use migration SQL and its populated outputs as the primary authority for what values/defaults must be
  represented in the customer template.
- Use `/Users/Rakesh.Patel/Documents/__data/excel/threeplus/CCS/FA E2E Customers V1.xlsm` only to fill
  genuine gaps or add broader allowed values where the migration does not already define the business-valid set.
- For `Tax ID Type`, the migration remains authoritative for what **not** to derive; the CCS workbook
  supplies the broader country-specific allowed-value lists.

## UX state

- Picklists are complete for the current customer-template scope.
- Cascading picklists are implemented for:
  - `Region` filtered by `Address Country Code`
  - `Tax ID Type` filtered by `Tax ID Country Code`
- Common BAU defaults are pre-filled for the AR use case, while preserving non-GB selections where
  valid.
- The workbook includes an `Instructions` tab generated from config plus the column metadata.

## Verification

- Re-render with:

```bash
uv run --with openpyxl --with pyyaml --no-project python \
  elt_transform_excel/scripts/create_ingest_template.py ar_customer --overwrite
```

- Confirm the renderer summary matches expectations and that the workbook opens cleanly in Excel.
- For schema / `feeds` / `drives` / completeness guidance, see `docs/CONFIGURATION.md` § "Ingest Specs".

## History

- Full implementation history, milestone checkpoints, earlier backlog items, and detailed review trail:
  [ar-ingest-specs-history.md](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/docs/todo/archive/ar-ingest-specs-history.md)

### The method (follow it exactly — it is what caught four silent bugs already)

1. **Find what is actually a drop-down.** Read the EIB's own validations; never infer from the data:
   ```bash
   uv run --with openpyxl --no-project python -c "
   import openpyxl, warnings; warnings.filterwarnings('ignore')
   wb=openpyxl.load_workbook('/Users/rpatel/Documents/__data/excel/threeplus/CCS/FA E2E Customers V1.xlsm', data_only=True)
   for dv in wb['<TAB>'].data_validations.dataValidation:
       if dv.type=='list': print(dv.sqref, '->', dv.formula1)
   "
   ```
2. **Resolve the named range from its OWN `OFFSET(...)` definition** — `wb.defined_names['<Name>'].value`.
   **They start on different rows** (`Customer_Category` A2, `Default_Customer_Group` C3, `Payment_Term`
   E2, `Payment_Type` G3). Assuming a common start row silently drops or invents a value. Read down
   from the start cell until blank (that is what `COUNTA` in the definition means).
3. **Snapshot to `config/ingest_lookups/<name>.json`** — `{description, key_column, provenance, rows}`.
   `provenance` must name the named range, the workbook and the date. Grab the adjacent Ref ID column
   as `reference_id` if there is one (it often differs from the display value — see Gotchas).
4. **Add `picklist: {lookup, column, where?}`** to the column, and the name to `lookups:`.
5. **Verify** (all four must pass — see § Verification for the commands): spec parses; every
   `feeds`/`drives` ref still valid; **0 unaccounted**; render still byte-deterministic. Then check the
   drop-down ranges resolve to the right values by reading them back out of the rendered workbook.

### Remaining — group A: mirror an EIB list validation (12 remaining)

| # | Template column | EIB target | Named range |
|---|---|---|---|
| 1 | Accepted Currencies | `Customer Currencies!C` | `Currency` |
| 2 | Tax Code | `Customer Tax!C` | `Tax_Code` |
| 4 | Primary Tax ID | `Customer Tax!G` | `Yes_no` |
| 5 | Transaction Tax Status | `Customer Tax!J` | `TransactionTaxStatus` |
| 6 | Address Public | `Customer Address!V` | `Yes_no` |
| 10 | Phone Public | `Customer Phone!L` | `Yes_no` |
| 11 | Email Public | `Customer Email!F` | `Yes_no` |
| 12 | Invoice Delivery Method | `Customer Document Delivery!E` | `Invoice_Delivery_Method` |
| 13 | Statement Delivery Method | `Customer Document Delivery!G` | `Invoice_Delivery_Method` |
| 14 | Invoice Notification Email Recipients | `Customer Document Delivery!F` | `Document_Email_Recipient_Invoice` |
| 15 | Statement Notification Email Recipients | `Customer Document Delivery!H` | `Document_Email_Recipient_Statement` |
| 17 | Tax ID Type | `Customer Tax!F` | **`INDIRECT`** — cascading, phase 2 |

Items 12–13 share one lookup; 14–15 use **two different** ranges — check both, do not assume.

### Remaining — group B: the EIB does NOT validate these, but they should be drop-downs anyway

Requested by the operator, and the data agrees. **The EIB's validation coverage is inconsistent** — it
validates `Public` on every tab but `Primary` on none, and leaves `Default Currency` free-text even
though `Accepted Currencies` next to it is a `Currency` list. Our input template may be stricter than
the EIB; that is the point of a form.

| # | Template column | EIB target | Source | Note |
|---|---|---|---|---|
| 18 | Default Currency | `Customer Currencies!D` | `Currency` range | operator: defaults to `GBP` |
| 19 | Address Primary | `Customer Address!W` | Yes/No | always `Yes` — one address per customer |
| 20 | Transaction Tax ID | `Customer Tax!H` | Yes/No | pairs with Primary Tax ID (#4), which IS validated |

`Phone Primary` / `Email Primary` are **not** template columns (transform-supplied from slot order), so
they need nothing here despite also being unvalidated Yes/No columns in the EIB.

Already done, do not redo: Customer Category, Additional Customer Group, Payment Terms, Default
Payment Type, Address/Phone/Tax Status Country Code, Tax ID Country Code, Phone 1/2/3 Device Type,
Region (cascading from Address Country Code).

### Cascading drop-downs — YES, possible, and the EIB already does it

The operator asked whether Region can be filtered by the selected country. **It can, and this workbook
is the proof.** `Customer Address!M` carries:

```
INDIRECT(SUBSTITUTE(SUBSTITUTE(...SUBSTITUTE(TRIM(K4)," ","_"),",","_")...))
```

It reads the country **name** from `K4`, sanitises it into a legal name-range identifier (7 nested
`SUBSTITUTE`s), and `INDIRECT`s to a per-country named range — `United_Kingdom` exists in the workbook,
as does `United_Kingdom_Tax_Type` for the `Customer Tax!F` cascade.

**Our template can do the same, and more cleanly, because AR pick the ISO2 CODE.** A code is already a
legal identifier, so `INDIRECT("region_" & $X2)` → `region_GB` needs **no** `SUBSTITUTE` gymnastics.
That is an unplanned dividend of the `country_by_iso_code` decision.

This is now implemented in `scripts/create_ingest_template.py`: a `picklist` may declare
`filtered_by`/`key`, and the renderer emits one named range per key value plus an `INDIRECT(...)`
validation that uses the controlling country-code cell on the same row. The spec shape is:

```yaml
  picklist: {lookup: ref_country_region, column: region, filtered_by: Address Country Code, key: country_code}
```
Excel quirks to expect: `INDIRECT` validation cannot point at another workbook, every generated range
name must be unique and legal, and a blank key cell makes the list empty rather than erroring.

### Region — must exclude obsolete regions

`ref_workday_country_state_region.csv` (4,349 rows) is the region list; **233 rows are obsolete**,
marked by `(obsolete)` in the `instance` column. `ref_workday_county_obsolete.json` (160 entries) maps
`source_value` (`'Aberdeen City (obsolete)'`) → `target_value` (`'Aberdeenshire'`).

The drop-down must offer **only current** regions — AR must not be able to pick a dead one. The
obsolete→current map is for a *transform* repairing an existing value, not for the input list. Note
both files live in `elt_ingest_excel`: snapshot what you need into `config/ingest_lookups/`, do **not**
read them at render time (see the Constraints on binding).

### Also decide with the operator

- ~~`expected_value` + `picklist`~~ — **RESOLVED 2026-07-15 by the operator: pre-filled and
  changeable is correct.** `expected_value` ("identical on every row") is replaced by `default:`
  ("pre-filled on the first data row; the user may change it"). It composes with `picklist`, so the 12
  columns that carry one — `Accepted Currencies`, `Default Currency` = `GBP`, `Tax Code`, the `Public`
  flags, etc. — keep their pre-fill **and** get a drop-down. A `default` is also distinct from an
  `example`: a default is a real value and always ships pre-filled; an example is illustrative and
  appears only under `--sample-row`. Still open, but narrower: `default_payment_type_default` asks
  whether `Manual` is the right default *value*, not how to express it.
- **RESOLVED 2026-07-21:** `Customer Tax!D` stays a Workday country **name** in the EIB, but the AR
  template now asks for `Tax ID Country Code` and derives the name from `ref_country`, matching the
  other controlled country inputs and avoiding free-text country entry.

## Terminology

AR say **credit note**; Workday calls it a **Customer Invoice Adjustment**. It is a different EIB
workbook (`DC_Customer_Invoice_Adjustment_v46.1.xlsx`, sheet `Submit Customer Invoice Adjustm`) with a
different column layout from the Customer Invoice EIB — the column letters do **not** correspond
between the two. In GP both live in one extract, split on `document_type`; AR get separate templates
(confirmed 2026-07-15), so no template has a Document Type column.

## Done

### Item 1 — Define the ingest specs ✅ (2026-07-15)

- Created `config/ingest_specs/` as a new config layer — the **input** side (a template this repo hands
  *to* finance), as distinct from `config/transform_specs/`, the **output** side (how an existing
  source workbook maps onto an EIB).
- `config/ingest_specs/ar_customer.yaml` — 22 columns (7 required, 1 conditional), 11 excluded, 0 open
  questions. Derived from `elt_ingest_excel/config/transform/sql/finance/customer/`.
- `config/ingest_specs/ar_customer_invoice.yaml` — 9 columns (all required), 7 excluded, 2 open
  questions. Derived from `.../finance/customer_invoice/workday_customer_invoice.sql`.
- `config/ingest_specs/ar_customer_credit_note.yaml` — 12 columns (11 required), 8 excluded, 3 open
  questions. Derived from `.../finance/customer_invoice/workday_customer_credit_note.sql`.
- `docs/CONFIGURATION.md` — four config layers → five; added § "Ingest Specs" (top-level + per-column
  key tables, the provenance rule); qualified "No configuration is shared with `elt_ingest_excel`" to
  say the cited paths are design-time provenance only, followed by hand, never loaded or synced.
- Verification: the command in § Verification passes on all three specs (output recorded there).

### Item 1b — Audit + EIB tab/column mapping: `ar_customer` only ✅ (2026-07-15)

Answered the operator's question "does `ar_customer.yaml` represent the GP extract limited to just the
columns used to populate the `workday_*` tables?" — **yes, and the audit found two defects, both now
fixed.** Verified against the live DuckDB and the real EIB workbook, not by reading SQL alone.

- **Audit result — the GP debtor extract has 61 columns:** 28 referenced by no customer SQL at all; 6
  more (`short_name`, `statement_name`, `inactive`, `checkbook_id`, `credit_limit_type`,
  `credit_limit_amount`) referenced **only** by `WHERE 1=2` transforms; leaving **27 live columns**,
  which is what the template's 22 are built from (plus `Customer ID`, which is AR-supplied and has no
  GP counterpart — the migration generates it). Every omission is now in `excluded:` grouped by reason.
- **Defect 1 (fixed) — `Tax Schedule ID` claimed `feeds: workday_customer_tax.tax_schedule_id`. No such
  output column exists.** It is written nowhere; it *selects* `Transaction Tax Status` (Customer Tax!J):
  `SOS` → "Customer of Services" vs "Customer of Goods" for GB, `SZ` → the EU-exempt branch. A false
  `feeds` is worse than no entry — it reads as verified when it is fiction.
- **Defect 2 (fixed) — `Address Code` was excluded as "never surfaced". It is surfaced**, inside
  `Address ID` (Customer Address!J), built as `customer_id_ADDRESSCODE_rnk`. Still excluded, but now
  for the true reason: at one-address-per-row it is always `MAIN`.
- **Introduced `feeds` vs `drives`.** `feeds` = written to that tab/column; `drives` = never written,
  selects a derived value. `County` and `Tax Schedule ID` have `drives` only — which exposed a bad rule
  in my own validator (it required `feeds` on every column, which would have forced a fictional target
  onto exactly those columns). Rule is now "at least one of the two".
- **Added `template.fan_out`** — one customer row is 1 Customer Name row but *n* Customer Phone rows
  (one per populated phone/fax, each then `;`-split). Same for Email.
- **Resolved `multiple_addresses_per_customer`** (operator, 2026-07-15): **one address per customer,
  full stop.** Recorded under `decisions:` in the spec, not deleted. See § The merge-artifact rule.
- `docs/CONFIGURATION.md` — documented `feeds`/`drives`, `fan_out`, the positional-write hazard, and
  the EIB verification command.
- Verification: 35 tab/column references checked against `FA E2E Customers V1.xlsm` — tab exists,
  column letter exists, header text matches. 0 bad. See § Verification.

### Item 1c — `ar_customer` rebuilt on the reversed derivation direction ✅ (2026-07-15)

Supersedes Item 1b's column set for `ar_customer` (the audit findings and EIB verification method still
stand). Driven by the operator's call that the template should be *the exact list of columns we populate
to the EIB*, so AR provide final Workday values and the transform stops deriving.

- **Measured what the EIB actually receives**, empirically from the live DuckDB rather than by reading
  SQL: of **124 columns across the 9 live tabs, only 64 are ever populated** — 60 are permanently
  empty. Of the 64: ~28 genuinely vary, **9 are a fixed constant on all 4,132 rows**, 3 are generated
  IDs, 1 (`formatted_phone_number`) is an exact duplicate of another column.
- **`ar_customer.yaml` rebuilt: 42 template columns + 12 `transform_supplied`.** Together they account
  for every populated EIB column with **0 unaccounted for** — the spec's central invariant, and now a
  standing check (§ Verification).
- **`drives` went from 2 entries to 0.** Nothing left to drive once Region and Transaction Tax Status
  are typed directly. Recorded in `docs/CONFIGURATION.md` as a design signal: a spec accumulating
  `drives` has probably slipped back toward source-derivation.
- **Dropped `Company`** — its only target (Customer Company Restrictions) is `WHERE 1=2`. It existed
  only because GP had one. Recorded as decision `company_dropped`.
- **Dropped `Fax`** — the reference pipeline's `_FAX` branch produces **zero** rows; the only device
  types present are Mobile (2,613) and Landline (825). We populate no fax data, so there is no column.
- **Dropped `County`** — it was a GP input to the Region lookup, never an EIB field. AR type `Region`.
- **Phones changed to numbered slots + per-slot Device Type; `;` dropped for phones only.** Each EIB
  phone row carries its own Device Type, so a `;` list would force the transform to guess it. Email
  keeps `;` — every address in Email To shares `Type=to`, so there is no per-value attribute. Measured:
  3 slots covers all but 4 of 3,290 customers.
- **`Phone Country` is per-customer, not per-slot** (only 4 of 3,290 customers span two countries) and
  is a **separate column from the address Country** — they differ on 75 of 3,432 rows.
- **Constants are typed by AR** (operator's call), so the template equals the populated EIB set with
  nothing hidden in the transform. Marked with `default:` (originally `expected_value`) so the
  template pre-fills them.
- `docs/CONFIGURATION.md` — documented the derivation direction and its rationale, `transform_supplied`,
  `expected_value` (later renamed `default`), `eib_coverage`, `excluded_tabs`, the completeness invariant + command, and the
  `drives` signal. Also fixed a table broken by an earlier edit that inserted prose between its rows.

### Item 3a — the template generator ✅ (2026-07-15)

Closes the parked "script under `build/` or one-off?" decision: **a script**, because the specs are
now iterated on rather than written once, and the render must stay a pure function of the spec.

- `scripts/create_ingest_template.py` (264 lines) — `config/ingest_specs/<key>.yaml` →
  `dist/ingest_templates/<key>.xlsx`. Command in `docs/ops/CLI.md`; the prototyping loop is
  documented there too.
- **Domain-free by contract** — verified: no column name, tab name or business value appears in the
  script. It renders any ingest spec, and a spec change never requires a code change here.
- **Deterministic** — identical spec in, byte-identical workbook out, proven by hashing two runs 2s
  apart. Needed two fixes, both non-obvious (see Gotchas): openpyxl overwrites `properties.modified`
  at save time, and the `.xlsx` zip entries carry clock mtimes.
- **Header text is the `name` verbatim** — verified: all 42 headers match the spec exactly, in order.
  Obligation is carried by fill colour and a cell comment, never by decorating the header text,
  because the header string is the transform's `source_column` contract.
- **Rejects a spec whose `feeds` is still the prose-string form**, naming the offending columns and
  pointing at Item 1b. Found by running it against `ar_customer_invoice`, where it first crashed with
  a raw `AttributeError` — rendering those specs would otherwise produce a plausible-looking template
  built on the pre-inversion column set.
- Verified: renders 42/42 columns, all with comments; workbook re-opens; error paths (missing spec,
  existing file without `--overwrite`, old-format spec) all exit non-zero with a clear message.
- **Fixed a pre-existing repo bug found on the way, then removed its cause.** `.gitignore` carried the
  stock Python `build/` rule, which excluded `elt_transform_excel/build/` — a **source** directory —
  from version control. All three scripts had never been committed, so a fresh clone got none of the
  tooling `docs/ops/CLI.md` documents. Two changes, both on the operator's call:
  1. Removed the `build/` rule — build is source; only `dist/` is derived output.
  2. **Renamed `build/` → `scripts/`**, because the name was the actual cause: `build/` means
     "generated output" almost everywhere, which is why the stock rule ignored it. Fixing only the
     `.gitignore` would have left the collision live for the next `.dockerignore`, CI default, or
     copied module. `scripts/` matches `elt_ingest_excel/scripts/` and nothing auto-ignores it.
  Cost was zero code change (all three scripts resolve via `SCRIPT_PATH.parents[1]`, and the `build`
  mentions in `vba_src/*.bas` are prose, not paths) and no git history churn (the directory had never
  been tracked, so there was nothing to `git mv` — this was the cheapest possible moment).
  Verified: 0 stale `build/` references repo-wide including memory files; the generator runs and stays
  deterministic from the new path; `build_runner.py --verify` **passes** against the real
  `dist/Runner.xlsm`, which is what proves the existing toolchain survived the move.
  `.gitignore` now explains why no `build/` rule exists, so the stock rule does not come back.

### Item 1d — country by ISO code + reference tabs ✅ (2026-07-15)

- **AR pick the ISO2 country CODE, not the name.** Three separate coded columns —
  `Address Country Code`, `Phone Country Code`, `Tax Status Country Code` — and the transform writes
  the names and the dialling code from `ref_country`. The name is the field that attracts variance
  (UK / United Kingdom / Great Britain — exactly what `ref_source_country_name_mapping` existed to
  repair); the code has none. Verified bijective: all 106 `ref_country` rows have a unique code *and*
  a unique name, so nothing is lost. Confirmed against the accepted upload: `Address!L`/`Phone!E` are
  ISO2, `Address!K`/`Phone!D`/`Tax!I` are names, `Phone!F` is the dialling code with `+` stripped.
- **Three codes, not one — measured, not assumed.** Phone country differs from address country on 75
  of 3,432 reference rows; tax status country differs on 3 of 4,120. Collapsing them would silently
  miscode those rows. `Tax Status Country Code` is `drives`-only: the Customer Tax tab has no
  country-code column.
- **Closed `generated_field_columns`:** the transform populates all three Generated Fields — they fall
  out of the country lookup for free and the accepted upload had them populated.
- **`lookups:` — reference tabs in the generated workbook.** Spec-driven `{tab, source, key_column}`;
  the generator snapshots each JSON source into its own tab (`ref_country`, 106 rows). Generic by
  design: JSON list-of-objects → tab, headers from the first object's keys. Add more entries as other
  coded columns arrive; no code change needed.
- **`drives` was rehabilitated.** The previous note ("should be rare, its absence is a design signal")
  was wrong — coded columns need it. A column may carry both: `Phone Country Code` *feeds* `Phone!E`
  and *drives* `Phone!D`/`!F`. The completeness invariant now counts `feeds` + `drives` +
  `transform_supplied`; `docs/CONFIGURATION.md` records the line a `drives` rule may not cross.
- Verified: 43 columns + 8 transform_supplied; 54 EIB refs valid, 0 bad; **64 populated columns, 0
  unaccounted**; render still byte-deterministic with the lookup tab.

**Decisions taken (2026-07-15, by the operator — do not re-litigate):**

| Decision | Choice | Consequence for later items |
|---|---|---|
| Header naming | Business title case (`Customer ID`) | These strings become `source_column` values verbatim in the Item 4 transform specs. Renaming a header is a transform-breaking change. |
| Credit notes | Separate template, no Document Type column | The GP `document_type` split is not reproduced. Two templates, two transforms, two EIBs. |
| Customer shape | Single flat sheet, GP shape | Address 1/2/3, Phone 1/2/3, Email To/CC/BCC as repeated columns — one row per customer. |
| Customer ID | AR supply the Workday Customer ID; `Reference ID` optional, carries the CRM key or legacy GP number | No ID-generation logic, and no lookup needed between the customer and invoice templates — they join on Customer ID directly. |

## Still todo (next actions)

1. **Item 1b for `ar_customer_invoice.yaml`, then `ar_customer_credit_note.yaml`** — audit each against
   the live GP HATB extract (same method as `ar_customer`: list the source table's real columns, grep
   the transform SQL for genuine references, discount `WHERE 1=2` targets) and convert `feeds:` from a
   prose string to verified `{tab, column, field}` entries against the real EIB workbooks. *Do the
   invoice first and finish it before starting the credit note — one dataset at a time.*
   **Both EIB workbooks live in `~/Documents/__data/excel/finance_ref/customer_invoice/` and use header
   row 5 / data row 6** (not row 3 like the Customer EIB), and both are published with
   `use_column_name_index: Y` — so unlike the Customer EIB they are matched **by header name, not
   position**. The existing `feeds:` strings already name the target column letters as a starting
   hypothesis; they are unverified and must be checked, not trusted.
2. **`BLOCKED:` Item 2 — get AR's answers to the four blocking open questions.** They are recorded in
   the specs' `open_questions:` sections; the § Verification command lists them by ID. The next action
   is to resolve the blocker, not to proceed past it. See § Blocked below for what must not be assumed.
3. **Item 3b/3c — render the invoice and credit note templates.** The generator already handles them
   (it is spec-driven and domain-free); they just need Item 1b first — it rejects their current
   `feeds:` form by design. Then blocked by Item 2 for the credit note: `adjusted_invoice_link`
   changes its *column set*, not just a drop-down.
5. **Item 4 — author the transform specs, build profiles, and macros** to populate the EIBs, per
   `docs/ADDING_A_TRANSFORM.md`. One per template. Each spec's `feeds:` key already records the target
   EIB column letter for every field — that is the starting point, and it needs verifying against the
   real EIB templates (Step 2 of that playbook) before it is trusted.
6. **Item 5 — reconcile with the existing AR transforms.** `ar_invoice` and `ar_discipline_invoice`
   already populate the Customer Invoice EIB from their own bespoke extracts. Decide whether they fold
   into the standard template or stay separate. This is a governance-reviewed decision, not an
   incremental edit — three transforms hitting one EIB from three input shapes is the drift this work
   was started to remove.

## Blocked

**Item 2 — four decisions only Finance (AR) can make.** Each has a conversion-only value hard-coded in
the migration SQL, and that value must **not** become the BAU default:

| ID | Question | Why blocking | Migration's value |
|---|---|---|---|
| `adjusted_invoice_link` | Must a credit note reference the original invoice it adjusts? | Changes the credit note template's **column set** (needs an `Adjusted Invoice` column) *and* needs a target EIB column the conversion mapping never used. The largest conversion-vs-BAU gap. | No link — adjustments stand alone against the customer |
| `revenue_category_values` | Allowed Revenue Category list for AR-raised invoices? | Affects both invoice and credit note templates — resolve once for both | `AR Conversion` |
| `adjustment_reason_values` | Allowed Adjustment Reason list for AR-raised credit notes? | Needed before the template can offer a validated drop-down | `Conversion` |
| `invoice_line_grain` (non-blocking) | Do AR need multi-line invoices? | Template is one row per line with no grouping key, so every row becomes a single-line invoice (Row ID fixed at 1). Multi-line needs a grouping column + per-invoice line sequence | one line per invoice |

**Must not be assumed:** that `AR Conversion` / `Conversion` are acceptable defaults (they are
conversion artifacts, and a BAU invoice coded `AR Conversion` is miscoded but looks valid); that credit
notes do not need an invoice link merely because the migration has none (the migration had no invoices
to link *to* — that is a property of the conversion, not of the business rule); or that single-line
invoices are sufficient.

## Accumulated Active Constraints (active for all remaining items)

- **Derivation direction (supersedes the original provenance rule):** a spec's columns are exactly the
  populated EIB columns; AR supply final Workday values; the transform derives nothing from them. See
  § Derivation direction for the measurements and the line the transform may not cross.
- **The completeness invariant:** `columns` + `transform_supplied` account for every populated EIB
  column, exactly. Check it on every spec edit (§ Verification) — it is the one that catches both a
  field nobody fills and a field AR fill for nothing.
- **Measure, do not read.** Every column claim in `ar_customer` came from counting the live DuckDB
  (populated / distinct / constant), not from reading SQL. Reading the SQL would have kept `Fax` (a
  branch that emits zero rows), kept `Company` (a `WHERE 1=2` target), and preserved a tax derivation
  serving 13 of 4,132 rows as though it were load-bearing.
- **A `WHERE 1=2` target does not count as usage** (established by Item 1b). Six GP columns are
  referenced only by permanently-empty transforms. A grep that finds a column "used" without checking
  whether its target is live will readmit all six as fields AR must populate for no reason.
- **`feeds` vs `drives` (established by Item 1b).** `feeds` = written to that EIB tab/column; `drives` =
  never written, selects a derived value. Every column carries at least one; neither alone is mandatory.
  Do not invent a `feeds` target for a `drives`-only column — that is exactly the Defect 1 failure, and
  a fabricated target reads as verified fact to the next session.
- **Recorded EIB tab/column references must be verified against the real workbook, never reasoned
  about.** The header text at `meta.eib_header_row` is the check. See § Verification for the command.
- **Merge artifacts are not requirements** (established 2026-07-15). The reference pipeline merged three
  legal entities; this work creates one customer at a time. Do not mirror its dedup, aggregation,
  address array, ranking, or name normalisation. See § The merge-artifact rule for the full list and
  the one exception (`';'` multi-value contacts).
- **One address per customer** (decided by AR, 2026-07-15). No child Address tab, no address array, no
  Address Code column. Address ID is always `customer_id_MAIN_1`.
- **A resolved question moves to `decisions:` with its rationale — it is never just deleted.** The
  reference SQL still contains the machinery a closed decision rules out, so without the recorded "why",
  the next session re-derives the question from the SQL's shape and answers it the wrong way.
- **Two directories, two questions:** `config/ingest_specs/` = the input template's shape;
  `config/transform_specs/` = how a source workbook maps onto an EIB. Do not merge them, and do not
  restate one in the other. An ingest spec is written *before* any source data exists.
- **The ingest-template renderer reads the specs; the Runner does not.** `scripts/create_ingest_template.py`
  parses `config/ingest_specs/*.yaml` and renders `dist/ingest_templates/*.xlsx` from them. They are
  still not parsed by `build_runner.py` and not read by the finance user's runtime workbook.
- **No config is shared with `elt_ingest_excel`.** `meta.derived_from` and `allowed_values_ref` cite
  paths there as design-time provenance only — followed by hand, never loaded or synced. Any value
  copied across is a snapshot that can drift; re-check it against the cited path rather than trusting
  the copy.
- **Header text is a contract.** The `name:` of every column is the literal header written into the
  template *and* the `source_column` value the Item 4 transform spec will reference. Renaming one
  silently breaks the transform that reads it.
- **`open_questions` must be empty before the matching transform is authored** — the same rule
  `docs/ADDING_A_TRANSFORM.md` Step 0 applies to transform specs. An open question is a decision the
  source data cannot answer; guessing it produces a silently miscoded Workday-bound file.
- **Customer ID joins the three templates.** AR supply it on all three; there is no lookup and no
  generation. An invoice whose Customer ID does not match a customer has no fallback resolution path.
- **Config purity:** allowed value lists live in config/reference files, not in code or macros.

## Verification

No test suite covers these specs (they are config). This is the actual check — it validates the schema
contract from `docs/CONFIGURATION.md` § "Ingest Specs" and reports each spec's blocking questions:

```bash
cd elt_transform_excel && uv run --with pyyaml --no-project python -c "
import yaml,glob,sys
ok=True
for f in sorted(glob.glob('config/ingest_specs/*.yaml')):
    d=yaml.safe_load(open(f))
    for k in ('meta','template','columns','excluded','open_questions'):
        if k not in d: print(f'{f}: missing {k}'); ok=False
    for c in d.get('columns',[]):
        for k in ('name','required','type'):
            if k not in c: print(f'{f}: {c.get(\"name\")} missing {k}'); ok=False
        if not (c.get('feeds') or c.get('drives')): print(f'{f}: {c[\"name\"]} has neither feeds nor drives'); ok=False
        if c.get('type')=='enum' and 'allowed_values' not in c: print(f'{f}: {c[\"name\"]} enum w/o allowed_values'); ok=False
        if c.get('required')=='conditional' and 'condition' not in c: print(f'{f}: {c[\"name\"]} conditional w/o condition'); ok=False
    n=[c['name'] for c in d.get('columns',[])]
    if len(n)!=len(set(n)): print(f'{f}: duplicate column names'); ok=False
    b=[q['id'] for q in d.get('open_questions',[]) if q.get('blocking')]
    eib='EIB-mapped' if isinstance(d['columns'][0].get('feeds'), list) else 'NOT YET EIB-mapped'
    print(f'{f}: {len(n)} cols, {eib}, blocking={b or \"none\"}')
sys.exit(0 if ok else 1)
"
```

Actual output, re-run 2026-07-15 (exit 0):

```text
config/ingest_specs/ar_customer.yaml: 22 cols, EIB-mapped, blocking=none
config/ingest_specs/ar_customer_credit_note.yaml: 12 cols, NOT YET EIB-mapped, blocking=['adjusted_invoice_link', 'adjustment_reason_values', 'revenue_category_values']
config/ingest_specs/ar_customer_invoice.yaml: 9 cols, NOT YET EIB-mapped, blocking=['revenue_category_values']
```

**Two progress signals live in that output.** `NOT YET EIB-mapped` disappearing is the Item 1b signal;
a shrinking `blocking=` list is the Item 2 signal. Item 4 cannot start for a template until that
template reads `EIB-mapped` and `blocking=none`.

**Item 6 check — do the rendered drop-downs resolve to the right values?** A validation can be
well-formed and point at the wrong cells; only reading it back proves otherwise. The historical counts
above are checkpoint history only; the current `ar_customer` render is the source of truth for the
live drop-down count.

```bash
uv run --with openpyxl --no-project python -c "
import openpyxl, warnings; warnings.filterwarnings('ignore')
wb=openpyxl.load_workbook('elt_transform_excel/dist/ingest_templates/ar_customer.xlsx')
ws=wb['Customer']; hdr={c.column_letter: c.value for c in ws[1]}
def resolve(f):
    sheet, rng = f.split('!'); sheet=sheet.strip(\"'\")
    lo,hi = rng.replace('\$','').split(':')
    col=''.join(c for c in lo if c.isalpha())
    r1=int(''.join(c for c in lo if c.isdigit())); r2=int(''.join(c for c in hi if c.isdigit()))
    return [wb[sheet][f'{col}{r}'].value for r in range(r1,r2+1)]
for dv in sorted(ws.data_validations.dataValidation, key=lambda d: str(d.sqref)):
    for c in sorted({''.join(ch for ch in str(r).split(':')[0] if ch.isalpha()) for r in str(dv.sqref).split()}):
        v = resolve(dv.formula1) if '!' in dv.formula1 else ['(inline/INDIRECT)']
        print(f'{c:<3} {hdr.get(c,\"?\"):<30} {len(v):>3} values  {v[:3]}')
"
```

**Second check — every `feeds`/`drives` tab+column against the real EIB workbook.** The command is in
`docs/CONFIGURATION.md` § "Verifying an ingest spec's feeds/drives". As of 2026-07-15 it checks 35
references for `ar_customer` against `FA E2E Customers V1.xlsm` and reports **0 bad**. Re-run it after
any spec edit and after any change to the `workday_customer_*` SELECT order — the Customer EIB is
written positionally, so a reordered SELECT invalidates every recorded column letter silently.

For Items 3–5 the verification is the existing one — `docs/BUILD.md` plus a manual Excel run; there is
no headless VBA compile in this environment.

## Gotchas

- **The migration's distinct values are NOT the valid-value list.** This has now bitten four times, and
  it is silent every time: Customer Category 10 used vs **12** valid; Customer Group 5 vs **7**; Payment
  Terms 5 vs **9**; Default Payment Type 1 vs **13**. Building a drop-down from the output would have
  quietly denied AR `Football Agents`, `Individual person (Customer)`, both `Experiences by Wembley`
  groups, `Net 21`, `EOM + 30 Days`, `Net 14`, `Immediate`, and 12 payment types. **Always read the
  EIB's named range.**
- **Named ranges start on different rows** — `Customer_Category` A2, `Default_Customer_Group` C3,
  `Payment_Term` E2, `Payment_Type` G3. Resolve each from its own `OFFSET(...)` definition; assuming a
  common start row silently drops or invents a value.
- **`reference_id` usually differs from the display value, and is inconsistently cased**: `Net 21` →
  `NET_21` but `Net 7` → `Net_7`; `SEPA-Payroll` → `EFT`; `SEPA-Urgent` → `EFT_with_Reference`. The
  accepted upload writes the **display** value, so that is what a drop-down offers. Record
  `reference_id` for traceability; do not use it as the picked value without evidence.
- **`picklist` and `multi_value` are mutually exclusive** — a drop-down yields one value per cell and
  cannot produce a delimited list. If a column seems to need both, check the data: that is how
  `Additional Customer Group` turned out to be singular (402 of 402 have exactly one).
- **A `where` filter must match contiguous rows** — an Excel range cannot skip. The renderer fails with
  a clear message; the fix is to reorder the lookup file so the subset sits together.
- **Several template columns share one EIB column.** `Phone 1/2/3 Device Type` all target
  `Customer Phone!K`; a per-EIB-column view collapses them and looks like one. All three need the
  picklist.
- **Two EIB `Type` columns validate against `"Yes,No"`** (`Customer Address!X`, `Customer Email!H`)
  which is nonsense for a Type field — looks like an EIB authoring bug. Neither is a template column, so
  ignore both; just do not copy the pattern.
- **`Customer Children` in the uploaded file has 10 stray rows the pipeline never wrote.** They are
  integers (`2`, `4`) under the headers, not customer IDs (`C-000001`), so they are leftover
  scaffolding shipped inside the EIB template — not data, and not a gap in the spec. The tab is not in
  `publish_customer.json`. The upload was accepted with them present.
- **The generated `.xlsx` is a render, never a source.** Hand-edit it and the next regenerate discards
  the change silently. Edit the spec.
- **openpyxl fights determinism in two places.** `save_workbook()` assigns
  `properties.modified = utcnow()` *itself*, at save time — setting it beforehand has no effect, so it
  must be corrected after saving. And `.xlsx` is a zip whose entries carry clock-derived mtimes.
  `save_deterministic()` handles both by saving to memory and rewriting the zip. If a future edit makes
  the render non-deterministic again, suspect these before anything else.
- **Do not decorate generated header text** (no `*` for required, no type hints). The header string is
  the transform's `source_column` contract — obligation belongs in fill colour and the cell comment.
- **DBeaver holds an exclusive lock on the DuckDB.** Any query fails with a lock IOException while it is
  open. Rather than asking the operator to close it, copy the file and open the copy read-only:
  `cp ~/Documents/__data/duckdb/rpatel.duckdb <scratchpad>/ro.duckdb` (~200MB, instant).
- **The Customer EIB is written POSITIONALLY.** `publish_customer.json` has no `use_column_name_index`,
  so `workday_customer_*` table column *n* lands in EIB column *n*, matched by order alone. Reordering a
  `SELECT` in the transform SQL silently shifts every subsequent EIB column and nothing detects it. The
  Customer **Invoice** and **Adjustment** EIBs are the opposite — `use_column_name_index: Y`, matched by
  header name. Do not carry an assumption from one to the other.
- **A loose grep for GP column usage produces false positives.** `.country_code` matches
  `ref_country.country_code` as well as the source column; `city`/`county` collide with reference-table
  columns too. Item 1b's audit needed a precise word-boundary grep per candidate to be trustworthy —
  the loose pass reported 33 "used" columns where the true figure is 27.
- **The credit note EIB is a different workbook from the invoice EIB.** `DC_Customer_Invoice_Adjustment_v46.1.xlsx`
  vs `Submit_Customer_Invoice_v46.1.xlsx`. Column letters do not correspond between them — the
  adjustment's `bl` is Invoice Date, the invoice's `gf` is. Reading across from one spec to the other
  produces a plausible, entirely wrong mapping.
- **The GP amount is split across two columns, the template has one.** `current_transaction_amount_gbp`
  for GBP; `current_transaction_amount_in_originating_currncy` otherwise (note the source's misspelling
  — it is a real column name, not a typo to fix); JPY rounds to whole units. The template carries a
  single `Amount` because AR state the Currency explicitly next to it.
- **Credit note Amount is positive.** Direction is carried by `Increase Amount Due` (N = credit), not by
  the sign. A negative amount double-negates.
- **`Document Number` does not go where you would expect.** The Customer Invoice EIB has no
  invoice-number field at all — `ar_discipline_invoice` writes it into Memo (`mt`), which is the
  established handling. On the adjustment EIB it goes to the line memo (`fc`), which *does* exist.
- **Three sheets, one GP workbook.** The HATB extract has FA-AR / NFC-AR / WNSL-AR sheets and the
  business unit comes from *which sheet a row is on*, not from a column. The template makes `Company`
  an explicit column instead — do not go looking for a source column that corresponds to it.
- **Several `workday_customer_*` transforms are `WHERE 1=2`** (credit, company restrictions, settlement
  account, intermediary account, credit card, bank account address, related worktags). They are wired
  but always empty. `Company` in `ar_customer.yaml` is optional for exactly this reason — its only
  target is inactive.
- **A blank Payment Terms loses the customer silently.** `src_fin_customer.sql` filters out any record
  with no payment terms, so the row vanishes rather than erroring. Hence required, defaulted `30 DAYS`.
- **Country drives four other fields** (Currency, Language, Tax ID Type, region validation). An
  unrecognised country name degrades all four without failing.
- **YAML flow mappings cannot hold block scalars.** `{column: X, reason: >...}` is a parse error — the
  `excluded:` entries use flow form, so a long reason must switch to block form. This was caught by the
  verification command during Item 1, not by inspection.
