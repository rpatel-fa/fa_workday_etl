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
