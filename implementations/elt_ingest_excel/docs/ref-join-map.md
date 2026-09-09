# Ref Table Join Map

This module uses DuckDB “reference tables” (`ref_*`) as the mechanism for mapping and normalising source extract values into Workday-approved values.

The fastest way to understand a mapping is to follow this chain:

1. **Data file** under `config/data/` (JSON/CSV)
2. **Ref loader SQL** under `config/transform/sql/ref/` (creates a `ref_*` table)
3. **Pipeline SQL join** under `config/transform/sql/<domain>/<pipeline>/` (typically in a `src_*` step)
4. **Publish table** under the same pipeline folder (`workday_*`) that writes the final worksheet-shaped output

## How ref tables are loaded

Ref tables are loaded by SQL files that read JSON/CSV from `config/data/`, for example:

- [ref_supplier_category.sql](config/transform/sql/ref/ref_supplier_category.sql) reads [ref_supplier_category.json](config/data/ref_supplier_category.json) and creates `ref_supplier_category`.

These ref loader SQL files are included in a pipeline by listing them in that pipeline’s `order.txt`.

## Finance / Supplier: ref joins from source columns

Transform order file: [finance/supplier/order.txt](config/transform/sql/finance/supplier/order.txt)

The primary join hub for ref tables in this pipeline is:

- [src_fin_supplier_raw.sql](config/transform/sql/finance/supplier/src_fin_supplier_raw.sql)

### `ref_source_business_unit_mapping` (source company → business unit)

- Data file: [ref_source_business_unit_mapping.json](config/data/ref_source_business_unit_mapping.json)
- Ref loader: [ref_source_business_unit_mapping.sql](config/transform/sql/ref/ref_source_business_unit_mapping.sql)
- Joined from source column: `fin_supplier_*.<company>` (aliased to `business_unit`)
- Join location: [src_fin_supplier_raw.sql:L85-L88](config/transform/sql/finance/supplier/src_fin_supplier_raw.sql#L85-L88)

### `ref_supplier_category` (source vendor class → Workday supplier category)

- Data file: [ref_supplier_category.json](config/data/ref_supplier_category.json)
- Ref loader: [ref_supplier_category.sql](config/transform/sql/ref/ref_supplier_category.sql)
- Joined from source column: `fin_supplier_*.<vendor_class_id>`
- Join location: [src_fin_supplier_raw.sql:L210-L213](config/transform/sql/finance/supplier/src_fin_supplier_raw.sql#L210-L213)
- Output column: `nrm_supplier_category` (defaults to `Miscellaneous`) at [src_fin_supplier_raw.sql:L109-L110](config/transform/sql/finance/supplier/src_fin_supplier_raw.sql#L109-L110)

### `ref_source_payment_terms` (source payment terms → Workday payment terms id)

- Data file: [ref_source_payment_terms.json](config/data/ref_source_payment_terms.json)
- Ref loader: [ref_source_payment_terms.sql](config/transform/sql/ref/ref_source_payment_terms.sql)
- Joined from source column: `fin_supplier_*.<payment_terms_id>`
- Join location: [src_fin_supplier_raw.sql:L214-L217](config/transform/sql/finance/supplier/src_fin_supplier_raw.sql#L214-L217)
- Output column: `nrm_payment_terms_id` at [src_fin_supplier_raw.sql:L133-L134](config/transform/sql/finance/supplier/src_fin_supplier_raw.sql#L133-L134)

### `ref_source_country_name_mapping` / `ref_source_country_code_mapping` / `ref_country` (source country → country attributes)

- Data files:
  - [ref_source_country_name_mapping.json](config/data/ref_source_country_name_mapping.json)
  - [ref_source_country_code_mapping.json](config/data/ref_source_country_code_mapping.json)
  - [ref_country.json](config/data/ref_country.json)
- Ref loaders:
  - [ref_source_country_name_mapping.sql](config/transform/sql/ref/ref_source_country_name_mapping.sql)
  - [ref_source_country_code_mapping.sql](config/transform/sql/ref/ref_source_country_code_mapping.sql)
  - [ref_country.sql](config/transform/sql/ref/ref_country.sql)
- Joined from source columns:
  - country name: `fin_supplier_*.<country>`
  - country code: `fin_supplier_*.<country_code>`
- Join locations (mapping + final country join): [src_fin_supplier_raw.sql:L179-L209](config/transform/sql/finance/supplier/src_fin_supplier_raw.sql#L179-L209)
- Output columns (examples): `nrm_country_code`, `nrm_country_name`, `nrm_currency_code`, `nrm_phone_code`, `nrm_tax_id_type` at [src_fin_supplier_raw.sql:L102-L110](config/transform/sql/finance/supplier/src_fin_supplier_raw.sql#L102-L110)

### `ref_post_code_district` / `ref_post_code_workday_region` / `ref_workday_country_state_region` (postcode/region normalisation)

- Data files:
  - [ref_post_code_district.csv](config/data/ref_post_code_district.csv)
  - [ref_post_code_workday_region.json](config/data/ref_post_code_workday_region.json)
  - [ref_workday_country_state_region.csv](config/data/ref_workday_country_state_region.csv)
- Ref loaders:
  - [ref_post_code_district.sql](config/transform/sql/ref/ref_post_code_district.sql)
  - [ref_post_code_workday_region.sql](config/transform/sql/ref/ref_post_code_workday_region.sql)
  - [ref_workday_country_state_region.sql](config/transform/sql/ref/ref_workday_country_state_region.sql)
- Joined from source columns:
  - postcode: `fin_supplier_*.<post_code>` (normalised earlier in the same SQL file)
  - county/city: `fin_supplier_*.<county>`, `fin_supplier_*.<city>`
- Join locations:
  - postcode district: [src_fin_supplier_raw.sql:L172-L178](config/transform/sql/finance/supplier/src_fin_supplier_raw.sql#L172-L178)
  - postcode → workday region override: [src_fin_supplier_raw.sql:L223-L226](config/transform/sql/finance/supplier/src_fin_supplier_raw.sql#L223-L226)
  - workday region instance validation/normalisation: [src_fin_supplier_raw.sql:L227-L256](config/transform/sql/finance/supplier/src_fin_supplier_raw.sql#L227-L256)

### `ref_workday_county_obsolete` (obsolete region remap)

- Data file: [ref_workday_county_obsolete.json](config/data/ref_workday_county_obsolete.json)
- Ref loader: [ref_workday_county_obsolete.sql](config/transform/sql/ref/ref_workday_county_obsolete.sql)
- Joined from derived value: `COALESCE(rx.instance, rxo.instance)` (derived from postcode/city/county logic)
- Join location: [src_fin_supplier_raw.sql:L266-L268](config/transform/sql/finance/supplier/src_fin_supplier_raw.sql#L266-L268)
- Output column: `nrm_region` at [src_fin_supplier_raw.sql:L109-L111](config/transform/sql/finance/supplier/src_fin_supplier_raw.sql#L109-L111)

### `ref_country_tax_id_type_mapping` (default tax id type label by country)

- Data file: [ref_country_tax_id_type_mapping.json](config/data/ref_country_tax_id_type_mapping.json)
- Ref loader: [ref_country_tax_id_type_mapping.sql](config/transform/sql/ref/ref_country_tax_id_type_mapping.sql)
- Joined from derived value: `src_fin_supplier.nrm_country_code`
- Join location: [workday_supplier_tax.sql:L63-L66](config/transform/sql/finance/supplier/workday_supplier_tax.sql#L63-L66)

### `ref_workday_group_additional` (extra group assignment by vendor class)

- Data file: [ref_workday_group_additional.json](config/data/ref_workday_group_additional.json)
- Ref loader: [ref_workday_group_additional.sql](config/transform/sql/ref/ref_workday_group_additional.sql)
- Joined from source column: `src_fin_supplier.vendor_class_id` (originating from `fin_supplier_*.<vendor_class_id>`)
- Join location: [workday_supplier_groups.sql:L24-L28](config/transform/sql/finance/supplier/workday_supplier_groups.sql#L24-L28)

