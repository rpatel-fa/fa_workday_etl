# Workday Customer Load — Payload Schema & Valid Values

Canonical cold-start document for `elt_ingest_excel/config/schema/sql/finance/customer/`.
This folder defines the **consolidated payload contract** that other teams use to transform their source data into a Workday Customer EIB-compatible load.

This contract is a **closed, validated deliverable** for the completed GP→Workday customer migration. It covers **only** the 10 Workday FIN Master Customer Workbook sheets that were actually populated, validated, and sent to Workday. Any Workday table outside this set is explicitly out of scope for this payload definition.

---

## 1. What this folder is

Teams ingesting customer data into Workday need an answer to:

> *"What fields do I populate, and what values are valid for each field, so my data actually loads through the FIN Master Customer Workbook EIB?"*

Before this folder existed, the only specification was scattered across:
- The 10 SQL transform files in `elt_ingest_excel/config/transform/sql/finance/customer/` that produce the Workday-bound `workday_customer_*` DuckDB tables
- A handful of JSON/CSV reference tables under `elt_ingest_excel/config/data/` that the SQL joins to
- Address validation rules embedded in the external workbook `FIN Master Customer Workbook Instructions.xlsx`

This folder consolidates **just the outward-facing contract** — no GP-source cleansing logic, no postcode-to-county lookups, no sort-code prefix tables. Two artifacts:

| File | Role |
|---|---|
| `workday_customer_payload.schema.json` | JSON Schema (Draft 2020-12) of the payload. Root = Customer Name sheet fields; nested sections = the 9 other workbook sheets that were actually loaded. Use this to **validate payloads before submission** via any JSON Schema validator (Ajv, jsonschema Python package, .NET DataContract attributes, etc.). |
| `workday_customer_valid_values.json` | Companion dictionary — every constrained attribute has a key here with either the full value list (closed enums), the regex pattern (ISO/open sets), or decision rules (per-country or per-currency Workday-level logic). Use this to **render pickers, build lookups, or debug "invalid value" rejections from Workday**. |

---

## 2. Files in this folder

```
elt_ingest_excel/config/schema/sql/finance/customer/
├── workday_customer_payload.schema.json   # JSON Schema — single consolidated payload per customer
├── workday_customer_valid_values.json     # Valid-value dictionary, cross-referenced from schema via x_valid_values_key
└── README.md                              # This file — canonical cold-start doc for the folder
```

---

## 3. How the two files relate — programmatically

### 3a. Header-level contract in `workday_customer_payload.schema.json`

```
x_valid_values_dict  : "workday_customer_valid_values.json"
x_valid_values_usage : "Every constrained attribute has x_valid_values_key → lookup key in companion dict"
x_workbook_sheets_in_contract : [9 sheet names]   # exactly the sheets covered
x_workbook_sheet_load_order   : [9 step sequence]  # FIN Master Customer Workbook EIB load order
```

### 3b. Inline cross-reference on every constrained attribute

Each attribute in the schema that has an enum, a pattern, or rule-based values carries the extension property:

```
x_valid_values_key: "<dot-notation key>"
```

Examples:
- `customer_category.x_valid_values_key = "root.customer_category"`
- `tax.transaction_tax_status.x_valid_values_key = "tax.transaction_tax_status"`
- `addresses[].country_code.x_valid_values_key = "addresses[].country_code"`
- `document_delivery.invoice_delivery_method.x_valid_values_key = "document_delivery.invoice_delivery_method"`

### 3c. Companion dictionary entry structure

For each `x_valid_values_key`, `workday_customer_valid_values.json` has a matching top-level key, categorised:

| `category` field | Meaning | Content shape |
|---|---|---|
| `enum_closed` | Closed set — only the listed values are accepted by Workday EIB validation for this migration | `values` (or `all_valid_values`) + any per-country/per-currency decision sub-tables +, where applicable, the specific active value used for the GP migration (`default_for_gp_migration` / `value_for_gp_migration`). |
| `pattern_format` | Open set but must match a pattern/regex | `pattern: "^..."` + `examples: [...]` (illustrative only, NOT a closed list) |
| `rule_based` | Value is set by Workday-level logic and cannot be statically listed | Decision rules / tables, example countries where applicable, and a pointer to the authority (e.g. the FIN Master Customer Workbook *Address Requirements* sheet) |

Tail convenience entries in the dictionary:
- `_meta` → description, path convention, category legend
- `_x_yes_no_fields_union.paths` → full list of every Yes/No flag path in the active payload, so teams can auto-render boolean dropdowns without walking the schema

### 3d. Quick validation + lookup snippet (Python)

```python
import json, jsonschema
from pathlib import Path

folder = Path("elt_ingest_excel/config/schema/sql/finance/customer")
schema = json.loads((folder / "workday_customer_payload.schema.json").read_text())
vals   = json.loads((folder / "workday_customer_valid_values.json").read_text())

# 1) Structural + enum + pattern validation
# payload = json.loads(sys.stdin.read())
# jsonschema.validate(instance=payload, schema=schema)

# 2) Given a schema x_valid_values_key, look up its value entry directly:
def lookup_valid_values(x_key):
    return vals[x_key]
```

---

## 4. Payload structure — exactly the 10 loaded Workday sheets

One payload object = **one Workday customer**, loaded through the FIN Master Customer Workbook EIB.

Root (Customer Name workbook sheet, always required):
- `customer_id` (string, minLength 1) — primary key; inherited by all nested sections
- `customer_name` (string)
- `reference_id`, `customer_category` (10 closed Workday values), `business_entity_name`, `external_entity_id`, `preferred_locale`, `preferred_language`, `lockbox`, `customer_security_segment`, `worktag_only`, `submit`, and the four `create_customer_from_*` Yes/No flags

Nested subordinate sections (each maps 1:1 to a workbook sheet):

| Nested property | Type per customer | Workbook sheet represented |
|---|---|---|
| `customer_groups[]` | array of objects (min 1) | Customer Groups |
| `tax` | single object | Customer Tax |
| `payment` | single object | Customer Payment |
| `currencies` | single object | Customer Currencies |
| `document_delivery` | single object — omitted for a customer that has no emails (no email rows → no delivery preference rows sent to Workday) | Customer Document Delivery |
| `addresses[]` | array of objects (min 1) | Customer Address |
| `phones[]` | array of objects | Customer Phone |
| `emails[]` | array of objects | Customer Email |

Child sections never repeat `customer_id` / `customer_name` — they are inherited from the root. The EIB workbook sheet builder (Python that flattens the payload into the xlsx) re-emits them onto each workbook row as needed.

---

## 5. Lineage — source SQL transform per section, and where value lists come from

*Cold-start rule: the SQL transform below is ALWAYS the source of truth if there is any disagreement with the schema or valid-values dictionary.*

### 5.1. Root — Customer Name sheet

- SQL: [workday_customer_name.sql](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_ingest_excel/config/transform/sql/finance/customer/workday_customer_name.sql)
- Upstream tables: `src_fin_customer_base` / `src_fin_customer_rnk` (customer-level aggregation)
- `customer_category` enum: 10 values from `ref_customer_category.target_value` → used via CASE mapping in the transform
- Four `create_customer_from_*` flags + `worktag_only` + `submit`: all Yes/No (driver table: `ref_yes_no_lookup` conceptually; inline in SQL)

### 5.2. `customer_groups[]`

- SQL: [workday_customer_groups.sql](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_ingest_excel/config/transform/sql/finance/customer/workday_customer_groups.sql)
- Always emits exactly one **alphabetical bucket** (A-L / M-Z / 0-9) per customer, derived from the first character of `customer_name` (digit → 0-9 bucket; A-L → A-L bucket; M-Z → M-Z bucket).
- Optionally emits additional class-based groups via a UNION against `ref_workday_group_additional`, mapped through the source customer class/type.

### 5.3. `tax`

- SQL: [workday_customer_tax.sql](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_ingest_excel/config/transform/sql/finance/customer/workday_customer_tax.sql)
- `tax_code`: single value `GBR 20% VAT (20%)` hardcoded in the transform CASE.
- `tax_id_type`: 20-value enum, with per-country default and GB pattern disambiguation (9-digit → VAT Reg No, 8-alnum → Company Number, 10-digit → UTR, else VAT Reg No) sourced from `ref_country_tax_id_type_mapping.target_value` and GB-specific special cases.
- `transaction_tax_status`: 7-value decision table, keyed on country group (GB / EU-27 inline list / RoW) × services_flag × vat_registered × exempt_flag.

### 5.4. `payment`

- SQL: [workday_customer_payment.sql](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_ingest_excel/config/transform/sql/finance/customer/workday_customer_payment.sql)
- `payment_terms`: 9 Workday labels from `ref_source_payment_terms.workday_payment_terms` CASE mapping.
- `default_payment_type`: enum `Manual / BACS / EFT / SEPA / Wire` — fixed to `Manual` for the GP migration.

### 5.5. `currencies`

- SQL: [workday_customer_currencies.sql](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_ingest_excel/config/transform/sql/finance/customer/workday_customer_currencies.sql)
- `accepted_currencies_plus`: 2-value enum `All Currencies / Currency Specific`.
- `default_currency`: ISO 4217 3-letter code (pattern + common examples).

### 5.6. `document_delivery`

- SQL: [workday_customer_document_delivery.sql](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_ingest_excel/config/transform/sql/finance/customer/workday_customer_document_delivery.sql)
- Conditional: per-customer section only emits if any emails exist (the SQL JOIN to `nrm_agg_emails` succeeds).
- Enums: `Email / Print / Electronic Invoicing` for delivery methods; named recipient selector enums for `invoice_notification_email_recipients` (Bill-To Contact Emails on Invoice / Bill-To Primary Contact / Statement Contact Emails / Remit-To Contact Emails) and `statement_notification_email_recipients` (Statement Contact Emails / Bill-To Primary Contact / Bill-To Contact Emails on Invoice).

### 5.7. `addresses[]`

- SQL: [workday_customer_address.sql](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_ingest_excel/config/transform/sql/finance/customer/workday_customer_address.sql) (+ [workday_customer_address_base.sql](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_ingest_excel/config/transform/sql/finance/customer/workday_customer_address_base.sql) base helper)
- Address column order in the schema's `items.properties` matches the EIB sheet column order exactly — do not reorder.
- Country code: ISO 3166-1 alpha-2 pattern.
- Max length of 35 chars enforced on `address_line_1`; HU addresses split at 35 char word boundary.
- Postcode is moved into `address_line_2` for countries whose address rules mark both `postal_code` and `address_line_2` as "not accepted" (unless the postcode matches a PO Box pattern).
- Authority for per-country component gating (required/optional/not accepted) is the `FIN Master Customer Workbook → Address Requirements` sheet (exposed in DuckDB table `ref_workday_validate_customer_address`).

### 5.8. `phones[]`

- SQL: [workday_customer_phone.sql](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_ingest_excel/config/transform/sql/finance/customer/workday_customer_phone.sql)
- 3 GP phone fields aggregated: `nrm_agg_phone_1 → primary`, `nrm_agg_phone_2 → secondary`, `nrm_agg_phone_3 → tertiary`.
- Normalisation pipeline: leading `+` stripped → `+440` → `+44` → embedded `+<code>0<number>` → country lookup via longest prefix match on `ref_country.phone_code` → test values matching `0*12345` discarded → final subscriber number must be digits only, 7–15 long, no leading 0, no country-code prefix.

### 5.9. `emails[]`

- SQL: [workday_customer_email.sql](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_ingest_excel/config/transform/sql/finance/customer/workday_customer_email.sql)
- RFC email regex + lowercasing + semicolon-split into multiple rows + per-customer first-wins global dedupe (no duplicate email across to/cc/bcc).
- Enum slots: `email_type = to / cc / bcc`; `use_for = Business / Personal / Home / Work / Billing / Shipping / Remittance` (Business fixed for the GP migration).

---

## 6. Load order — matches `x_workbook_sheet_load_order` header + SQL `order.txt`

Workday EIB sheets depend on each other: Customer Name must be loaded first, and each nested sheet references the already-loaded parent. Load order for the 9 nested sections is:

1. `root` — Customer Name
2. `customer_groups[]` — Customer Groups
3. `payment{}` — Customer Payment
4. `currencies{}` — Customer Currencies
5. `document_delivery{}` — Customer Document Delivery (skip per customer when no emails)
6. `tax{}` — Customer Tax
7. `addresses[]` — Customer Address
8. `phones[]` — Customer Phone
9. `emails[]` — Customer Email

The SQL run-order file [order.txt](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_ingest_excel/config/transform/sql/finance/customer/order.txt) in the transform folder must stay aligned with this list.

---

## 7. Design decisions written down

These were deliberate choices made while building the folder. Don't re-litigate them cold without a good reason.

### 7.1. Single consolidated payload vs 10 separate per-sheet schemas

A single JSON object per customer is the natural unit other teams want to POST / validate. Separate per-sheet files would re-introduce the exact fragmentation we are trying to eliminate (no way to validate "this whole customer is complete"). Flattening into 10 EIB rows per sheet is handled downstream by the sheet builder, not the payload contract.

### 7.2. ISO code sets (country, currency, phone codes) are patterns, not exhaustive enums

The valid ISO 3166 / ISO 4217 / E.164 lists are public and drift. Enumerating them here would go stale immediately. Each is a pattern plus illustrative examples.

### 7.3. Country/postcode/sort-code cleansing data is deliberately NOT exported

That data exists under `elt_ingest_excel/config/data/` for GP source cleansing only and is not part of the outward Workday EIB load contract. The valid-values dictionary contains only the final Workday-accepted labels, patterns, or Workday-level rules.

### 7.4. Child sections never repeat `customer_id` / `customer_name`

They are inherited from root. This keeps payloads compact and avoids a class of "child's customer_id differs from root" bugs. The EIB sheet builder re-emits them.

### 7.5. `additionalProperties: false` on every object node

Guarantees that typos or extra fields from a future-version drift are rejected at validation time, not silently dropped later by Workday. When adding new legitimate fields, add them to the schema explicitly.

### 7.6. Yes/No flags are string enums ("Yes" / "No"), not JSON booleans

Workday EIB templates expect these strings in the column cells, not `true` / `false`. Using string enums keeps the payload contract one step closer to the actual workbook insert.

### 7.7. Array paths in valid-values keys use `[]` (no index)

Valid values never depend on which position an array item is in (they apply uniformly across all `customer_groups[]`, `addresses[]`, `phones[]`, `emails[]` items). Using a literal `[]` component keeps one dict entry per rule, instead of duplicating the rule for hypothetical indices 0..N.

---

## 8. If something is wrong — cold-start maintenance checklist

Priority when a discrepancy is found:
1. **SQL transform wins.** Check the Section 5 file for the affected section first — it is always the ground truth.
2. If the SQL changed, update **both** the matching schema enum/pattern AND the matching dictionary entry in valid_values together. Never leave one stale.
3. If the value list was driven by a JSON/CSV reference table under `elt_ingest_excel/config/data/` (e.g. customer categories, payment terms mappings, tax ID types), update the reference file first and then refresh the enum here as a second step.
4. For address per-country rules: never hard-code country matrices in this README or the valid-values dict. Query `ref_workday_validate_customer_address` in DuckDB, which is built from the authoritative `FIN Master Customer Workbook → Address Requirements` sheet.

Common maintenance tasks mapped to concrete steps:
- **Add a new payment term mapping** → edit `ref_source_payment_terms.json` → refresh `payment.payment_terms.workday_labels` in valid-values dict + schema `payment.payment_terms` enum.
- **New tax code (e.g. reduced 5% VAT regime)** → edit the CASE in `workday_customer_tax.sql` → append to `tax.tax_code.values` + schema enum.
- **New tax ID type added to tenant** → append label to `tax.tax_id_type.all_valid_values` +, if needed, add a per-country rule row under `per_country_defaults`.
- **New Workday customer category** → append label to `root.customer_category.values` + schema enum.
- **New country address rule** → update the workbook / `ref_workday_validate_customer_address`; nothing needs changing in this folder unless you want illustrative examples refreshed.
