# Publishing workbooks (clone + populate)

Publishing is the final phase: take publish tables in DuckDB and write them into pre-built Excel templates (usually 3+ Workday workbooks).

## What “clone” means

Publishing does not build a workbook from scratch. It copies a template file to create the output workbook, then writes data into specific sheets.

This is why macros, formatting, validation rules, and hidden worksheets still exist in the generated file: they live in the template.

## What “populate” means

For each configured sheet mapping, publishing runs:

```
SELECT * FROM <srcTableName>
```

and writes the rows into the target worksheet starting at `dataRow`.

The publisher does not rewrite headers; headers are expected to exist in the template.

## Finance: Supplier/Creditor

Publish config: [publish_supplier.json](config/publish/finance/publish_supplier.json)

Template workbook to clone:

- `srcWorkbookFileName`: `FA E2E Suppliers V1.xlsm`

Output workbook produced:

- `tgtWorkbookFileName`: `workday_fin_creditor_supplier_active_v1` (extension comes from the template)

Worksheet population mappings:

- DuckDB `workday_supplier_name` → sheet `Supplier Name`
- DuckDB `workday_supplier_groups` → sheet `Supplier Groups`
- DuckDB `workday_supplier_alternate_name` → sheet `Supplier Alternate Name`
- DuckDB `workday_supplier_currencies` → sheet `Supplier Currencies`
- DuckDB `workday_supplier_tax` → sheet `Supplier Tax`
- DuckDB `workday_supplier_payment` → sheet `Supplier Payment`
- DuckDB `workday_supplier_procurement` → sheet `Supplier Procurement`
- DuckDB `workday_supplier_company_restrictions` → sheet `Supplier Company Restrictions`
- DuckDB `workday_supplier_settlement_account` → sheet `Supplier Settlement Account`
- DuckDB `workday_supplier_intermediary_account` → sheet `Supplier Intermediary Account`
- DuckDB `workday_supplier_address` → sheet `Supplier Address`
- DuckDB `workday_supplier_phone` → sheet `Supplier Phone`
- DuckDB `workday_supplier_email` → sheet `Supplier Email`

The tables above are produced by the supplier transform `order.txt` and `workday_supplier_*.sql` files.

## Finance: Customer/Debtor

Publish config: [publish_customer.json](config/publish/finance/publish_customer.json)

Template workbook to clone:

- `srcWorkbookFileName`: `FA E2E Customers V1.xlsm`

Output workbook produced:

- `tgtWorkbookFileName`: `workday_fin_debtor_customer_active_v1`

Worksheet population mappings:

- DuckDB `workday_customer_name` → sheet `Customer Name`
- DuckDB `workday_customer_groups` → sheet `Customer Groups`
- DuckDB `workday_customer_credit` → sheet `Customer Credit`
- DuckDB `workday_customer_payment` → sheet `Customer Payment`
- DuckDB `workday_customer_currencies` → sheet `Customer Currencies`
- DuckDB `workday_customer_company_restrictions` → sheet `Customer Company Restrictions`
- DuckDB `workday_customer_document_delivery` → sheet `Customer Document Delivery`
- DuckDB `workday_customer_settlement_account` → sheet `Customer Settlement Account`
- DuckDB `workday_customer_intermediary_account` → sheet `Customer Intermediary Account`
- DuckDB `workday_customer_tax` → sheet `Customer Tax`
- DuckDB `workday_customer_address` → sheet `Customer Address`
- DuckDB `workday_customer_phone` → sheet `Customer Phone`
- DuckDB `workday_customer_email` → sheet `Customer Email`

Note: the customer transform order currently includes additional publish tables (for example `workday_customer_related_worktags`, `workday_customer_bank_account_address`, `workday_customer_credit_card`) that are not listed in `publish_customer.json`. If those sheets become required, the publish config is the place to wire them into the output workbook.

## HCM: Contingent Worker

Publish config: [publish_contingent_worker.json](config/publish/hcm/publish_contingent_worker.json)

Template workbook to clone:

- `srcWorkbookFileName`: `HCM Workbook (Blank).xlsm`

Output workbook produced:

- `tgtWorkbookFileName`: `workday_hcm_contingent_worker`

Worksheet population mappings:

- DuckDB `workday_contingent_worker_name` → sheet `Name`
- DuckDB `workday_contingent_worker_professional_ee` → sheet `Professional EE`
- DuckDB `workday_contingent_worker_manager` → sheet `Manager`
- DuckDB `workday_contingent_worker_email` → sheet `Email`

## Picking the publisher backend

Publisher backend choice affects how reliable `.xlsm` cloning/preservation is:

- `xlwings`: requires Excel installed, best when the template is `.xlsm` and must preserve macros/shapes
- `openpyxl`: does not require Excel, but may not preserve all `.xlsm` artefacts

