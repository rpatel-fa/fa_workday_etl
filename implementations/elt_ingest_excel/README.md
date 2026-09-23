# elt_ingest_excel

`elt_ingest_excel` is a config-driven ELT pipeline that prepares Workday load workbooks from source Excel extracts:

- Read source Excel workbooks
- Load sheets into DuckDB as raw tables
- Transform via ordered SQL into publish tables shaped like target Workday worksheets
- Clone template Workday workbooks (typically from 3+)
- Populate specific worksheets from DuckDB publish tables

## Start here

- Architecture: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)
- Runbook (setup + CLI): [docs/CLI.md](docs/CLI.md)
- PRD (assistant context): [PRD.md](PRD.md)

## How to navigate the docs

- Pipeline phases and orchestration: [docs/pipeline.md](docs/pipeline.md)
- Config contracts (ingest/transform/publish): [docs/CONFIGURATION.md](docs/CONFIGURATION.md)
- SQL transform model (`order.txt`, shared `ref/` + `macro/`, UDFs): [docs/sql-transform.md](docs/sql-transform.md)
- Ref table join map (finance supplier): [docs/ref-join-map.md](docs/ref-join-map.md)
- Publishing mechanics + workbook mappings: [docs/publish-workbooks.md](docs/publish-workbooks.md)
- Example entry points (what to run): [docs/examples.md](docs/examples.md)
- Reference export outlier (`fin_workday_ref.py`): [docs/reference-export.md](docs/reference-export.md)
- Reconciliation (source→target data integrity, distinct from validation): [docs/reconciliation.md](docs/reconciliation.md)

## Standalone utilities (outside the SQL/ELT pipeline)

These scripts live under `scripts/` but are **not** part of the ingest → transform →
publish pipeline. They solve one-off Finance requests that reuse the repo's Python
tooling (openpyxl, duckdb conventions) but do not touch config or publish workbooks.

| Utility | Docs | When to use |
|---|---|---|
| [scripts/email_to_contact_names.py](scripts/email_to_contact_names.py) | [docs/email-contact-name-fill.md](docs/email-contact-name-fill.md) | Given a 3-col workbook (email / First name / Surname), fill names from the email local-part where confident, blank everything else. For Gayle/3+ style EIB prep: "leave blank if you can't tell unambiguously". |
| [scripts/populate_eib_corrections.py](scripts/populate_eib_corrections.py) | Header docstring in the script | Fills the yellow `CORRECTION` column on a Workday Submit-Supplier EIB error tab, joining the DuckDB transform output keyed on the stable GP Vendor reference so corrections carry leading zeros correctly. |
