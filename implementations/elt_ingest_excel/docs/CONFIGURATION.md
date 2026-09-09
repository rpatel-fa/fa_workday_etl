# Configs

All runtime behaviour is driven by config files under `elt_ingest_excel/config/`.

Pipelines generally specify three config inputs:

- Ingest config JSON: which source workbooks/sheets to load into DuckDB
- Transform directory: which SQL to execute (via `order.txt`)
- Publish config JSON: which template workbooks to clone and which DuckDB tables to write into which sheets

The canonical examples are:

- Finance supplier/creditor:
  - Ingest: [active_supplier_creditor.json](config/ingest/finance/active_supplier_creditor.json)
  - Transform: [finance/supplier/order.txt](config/transform/sql/finance/supplier/order.txt)
  - Publish: [publish_supplier.json](config/publish/finance/publish_supplier.json)
- Finance customer/debtor:
  - Ingest: [active_customer_debtor.json](config/ingest/finance/active_customer_debtor.json)
  - Transform: [finance/customer/order.txt](config/transform/sql/finance/customer/order.txt)
  - Publish: [publish_customer.json](config/publish/finance/publish_customer.json)
- Finance supplier GP→Workday reconcile (CCS):
  - Ingest: [workday_supplier_extract.json](config/ingest/finance/workday_supplier_extract.json)
  - Reconcile SQL: [finance/supplier_ccs/](config/reconcile/finance/supplier_ccs/) (run in filename order; no `order.txt`)
  - Publish: [publish_supplier_ccs_reconcile_review.json](config/publish/finance/publish_supplier_ccs_reconcile_review.json)

## Ingest config (Excel → DuckDB)

Shape: a JSON list of workbook configs. Each workbook config has one or more sheet configs.

Minimal schema:

```json
[
  {
    "workbookFileName": "source.xlsx",
    "fileType": "EXCEL",
    "sheets": [
      {
        "sheetName": "Sheet1",
        "targetTableName": "src_table_name",
        "headerRow": 1,
        "dataRow": 2
      }
    ]
  }
]
```

Notes:

- `workbookFileName` is interpreted relative to the `data_path` passed by the example runner.
- `headerRow` and `dataRow` are 1-indexed row numbers as seen in Excel.
- All source columns are read as strings; type conversion belongs in SQL.

## Transform config (DuckDB SQL execution)

Transform is defined by a directory containing an `order.txt` file.

- `order.txt` lists SQL filenames to execute, one per line, in order.
- Blank lines and comment lines (starting with `#`) are ignored.
- Entries may include relative paths to shared SQL under `config/transform/sql/ref/` and `config/transform/sql/macro/`.

The transform directory is chosen by the example runner (for example `transform/sql/finance/supplier`).

## Publish config (DuckDB → Excel)

Shape: a JSON list of workbook publish configs. Each workbook publish config has one or more sheet mappings.

Minimal schema:

```json
[
  {
    "srcWorkbookPathName": "/path/to/templates",
    "srcWorkbookFileName": "Template.xlsm",
    "tgtWorkbookPathName": "/path/to/output",
    "tgtWorkbookFileName": "output_file_name_without_extension",
    "sheets": [
      {
        "srcTableName": "workday_table_or_view",
        "sheetName": "Target Sheet Name",
        "headerRow": 3,
        "dataRow": 4
      }
    ]
  }
]
```

Notes:

- The target file extension is taken from the template workbook (for example `.xlsm`).
- `tgtWorkbookFileName` should not include an extension.
- Publishing runs `SELECT * FROM <srcTableName>`; column ordering is whatever the SQL produces.
- Paths in the repo configs are environment-specific and should be edited for the local machine.

