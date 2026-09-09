# Pipeline

`elt_ingest_excel` runs a config-driven ELT pipeline with three cumulative phases:

```
INGEST → TRANSFORM → PUBLISH
```

Most pipelines are executed via the scripts under `elt_ingest_excel/examples/`. Those scripts pick:

- The source data directory (`data_path`)
- An ingest config JSON (what to read from source Excel)
- A transform directory (which SQL files to execute, and in what order)
- A publish config JSON (which templates to clone and which DuckDB tables to write into which sheets)

## Where the code lives

- Orchestrator: `FileIngestor` in [elt_pipeline.py](src/elt_ingest_excel/elt_pipeline.py)
- Shared runner wiring: [base_runner.py](examples/base_runner.py)

The orchestrator is intentionally thin: it reads configs and delegates to loaders, transform executors, and publishers. Changing behaviour should usually be done by changing config and SQL, not Python.

## Phase: INGEST (Excel → DuckDB)

Goal: load source Excel sheets into DuckDB tables with minimal interpretation.

How it works:

- Ingest config lists one or more source workbooks and their sheets.
- Each sheet is read into a pandas DataFrame (all columns as strings).
- The DataFrame is written to DuckDB as the configured `targetTableName`.

Key properties:

- Source columns are normalised to DB-friendly names (lowercase + underscores).
- Type coercion is done in SQL during TRANSFORM (not during ingest).

## Phase: TRANSFORM (ordered SQL execution)

Goal: turn raw ingested tables into publish-ready tables that match the target workbook’s expected layout and Workday value requirements.

How it works:

- Each pipeline points at a transform directory containing an `order.txt`.
- `order.txt` lists SQL files to execute in order; execution stops on first failure.
- SQL files may contain multiple statements.
- Shared reference loads live under `config/transform/sql/ref/` and are typically executed first.
- Shared macros (SQL functions) live under `config/transform/sql/macro/`.

The naming convention used in the shipped pipelines is:

- `src_*` tables/views: pipeline-specific staging layers
- `workday_*` tables/views: final publish tables, shaped to match workbook sheets

## Phase: PUBLISH (DuckDB → Excel)

Goal: clone a template workbook and populate specific worksheets with the publish tables.

How it works:

- Publish config lists one or more workbooks to produce.
- For each output workbook:
  - A template workbook file is copied to the output location (clone step).
  - Each configured sheet runs `SELECT * FROM <srcTableName>` in DuckDB.
  - Result rows are written starting at `dataRow` on `sheetName`.

Important details:

- “Cloning” is a filesystem copy; it is not re-creating a workbook from scratch.
- The publisher writes data only; headers/formatting/macros come from the template.
- Two backends exist: `xlwings` (requires Excel, preserves macros/shapes) and `openpyxl` (no Excel required, may not preserve `.xlsm` artefacts reliably).
