# Architecture

`elt_ingest_excel` is designed as a config-driven ELT pipeline where:

- Python provides generic plumbing (read Excel, write DuckDB, run SQL in order, write Excel)
- SQL encodes domain logic (coercion, enrichment, mapping to Workday-approved values, shaping to worksheet layouts)
- JSON config defines what workbooks/sheets/tables participate in each pipeline

The intended contract is: adding or changing a pipeline is mostly config + SQL, not new Python.

## Core dataflow

```
Source Excel
   ↓ (INGEST: pandas/openpyxl/xlwings)
DuckDB raw tables
   ↓ (TRANSFORM: ordered SQL + UDFs)
DuckDB publish tables (workday_*)
   ↓ (PUBLISH: clone template + write rows)
Target Workday workbook(s)
```

## Phases and orchestration

Phases are cumulative: `PUBLISH` implies `TRANSFORM` and `INGEST`.

- Orchestrator: `FileIngestor` in [elt_pipeline.py](src/elt_ingest_excel/elt_pipeline.py)
- Entry points: scripts under [examples/](examples)
- Shared runner: [base_runner.py](examples/base_runner.py)

## Config-driven behaviour (contracts)

All operational behaviour is selected at runtime by:

- Ingest JSON (source workbook + sheet → DuckDB table mappings)
- Transform directory + `order.txt` (which SQL files run, and in what order)
- Publish JSON (template workbook → output workbook + DuckDB table → sheet mappings)

See [configs.md](configs.md) for the concrete schema shapes and examples.

## Ingest architecture (Excel → DuckDB)

Responsibilities:

- Read configured worksheets into DataFrames
- Normalise column names to DB-friendly identifiers
- Persist as DuckDB tables using a configured save mode

Key components:

- Excel sheet loading and column cleanup: [excel_reader.py](src/elt_ingest_excel/loaders/excel_reader.py)
- Sheet iteration: [sheet_processor.py](src/elt_ingest_excel/loaders/sheet_processor.py)
- DuckDB persistence: [duckdb_writer.py](src/elt_ingest_excel/writers/duckdb_writer.py)

Design choice:

- Ingest reads everything as strings; typing decisions are pushed into SQL during TRANSFORM. This keeps INGEST resilient to messy source extracts and makes the data contract explicit in SQL.

## Transform architecture (ordered SQL execution)

Responsibilities:

- Execute a deterministic sequence of SQL files to build staging and publish tables
- Fail fast on SQL errors to prevent publishing partial/invalid outputs
- Provide shared reference tables and reusable SQL “macros”
- Register UDFs needed by SQL (for example phone parsing)

Key components:

- Order execution: [executor.py](src/elt_ingest_excel/transform/executor.py)
- `order.txt` parsing: [order_reader.py](src/elt_ingest_excel/transform/order_reader.py)
- Multi-statement execution + path hardening: [sql_file_executor.py](src/elt_ingest_excel/transform/sql_file_executor.py)
- Connection/UDF registration: [db.py](src/elt_ingest_excel/transform/db.py) and [transform/udf](src/elt_ingest_excel/transform/udf)

Structural conventions used by the shipped pipelines:

- `ref_*`: shared reference tables (typically loaded first)
- `src_*`: pipeline-specific staging views/tables derived from ingest tables
- `workday_*`: publish tables intended to map 1:1 onto output worksheets

## Publish architecture (clone workbook + populate sheets)

Responsibilities:

- Clone a template workbook (filesystem copy)
- For each sheet mapping:
  - `SELECT * FROM <srcTableName>` in DuckDB
  - Write rows into `sheetName` starting at `dataRow`

Key components:

- Publish config parsing: [publish_config_parser.py](src/elt_ingest_excel/parsers/publish_config_parser.py)
- Publisher base (clone + query loop): [publish/base.py](src/elt_ingest_excel/publish/base.py)
- Backends:
  - Excel automation (preserve macros/shapes): [excel_publisher_xlwings.py](src/elt_ingest_excel/publish/excel_publisher_xlwings.py)
  - Pure Python writing: [excel_publisher_openpyxl.py](src/elt_ingest_excel/publish/excel_publisher_openpyxl.py)

Design choice:

- Templates carry formatting, validation rules, hidden sheets, and macros. Publishing is “data-only”, aligned to template row offsets.

Gotcha — `xlwings` backend and `Decimal`: on macOS, `xlwings` writes via AppleScript/AppleEvents, which can only pack native `int`/`float`/`str`/`None`. DuckDB `DECIMAL` columns come back from `fetchall()` as Python `decimal.Decimal`, which fails to pack with `Can't pack data into an AEDesc (unsupported type)`. `excel_publisher_xlwings.py` converts `Decimal` to `float` immediately before the write (scoped to that backend only — `excel_publisher_openpyxl.py` writes `Decimal` natively without issue, since it serializes XML directly rather than going through AppleEvents).

## Outliers and optional utilities

Not everything in `examples/` is a standard pipeline run:

- Reference export: [fin_workday_ref.py](examples/fin_workday_ref.py) writes a new `.xlsx` workbook of mapping tables for review/editing (see [reference-export.md](reference-export.md)).
- Legacy/on-hold Excel automation helpers exist for running workbook macros/validations; treat these as utilities rather than core architecture (see [examples.md](examples.md)).

## Design decision — name normalisation and entity matching

**Context:** the pipeline output is loaded directly into Workday as production master data (suppliers, customers). The failure modes are asymmetric:

- **False split** (same entity left as two records): a duplicate lands in Workday. Low blast radius — it is visible and can be corrected.
- **False merge** (two different entities collapsed to one): wrong supplier gets paid, wrong customer gets invoiced, or a genuine entity disappears from the migration. Hard to detect; hard to unwind in a live ERP.

**Decision: deterministic normalisation only in the pipeline.**

Name normalisation uses an explicit REGEXP_REPLACE/REPLACE chain in `src_fin_*_raw.sql` and `lgcy_src_fin_*_raw.sql`. Every normalisation step (e.g. `UNITED→UTD`, `AND→&`, parenthesised qualifier removal) is a named, auditable rule that only fires when validated. No rule is added without a confirmed real-world example.

**What was evaluated and rejected: probabilistic / fuzzy matching as a pipeline step.**

Splink (probabilistic entity resolution) was evaluated against the customer dataset. Result: the unsupervised EM model produced a flat prior — every pair scored 0.311–0.317 regardless of similarity. EM has no training signal on a small, unlabelled list; it outputs the base rate. Jaro-Winkler at threshold 0.88 produced false merges such as `CHRISTOPHER GEMBALA` / `CHRISTOPHER EDDEN` (different people sharing a first name score high due to Jaro-Winkler prefix bias). Neither approach is safe to run unattended against data going into an ERP.

**Role of the Jaro-Winkler comparison tool.**

`fin_customer_splink_compare.py` (which runs a `jaro_winkler_similarity` self-join) is a **diagnostic tool only** — it is not part of the pipeline and its output table (`customer_fuzzy_name_comparison`) is never consumed by any transform SQL. The intended workflow:

1. Run the pipeline to TRANSFORM.
2. Run `fin_customer_splink_compare.py` and inspect the results.
3. If a pattern is found that is safe to generalise (e.g. a new abbreviation variant), add a deterministic rule to the four `*_raw.sql` files and re-run.

The fuzzy output is never applied automatically. A human judges each candidate before it becomes a rule.

**Guidance for future changes:** do not introduce fuzzy or probabilistic matching into the transform SQL or pipeline phases. Extend the deterministic chain instead. Each new rule must have a concrete example that justifies it.

## Invariants and assumptions

- SQL ordering is part of the contract: `order.txt` defines reproducible output.
- Publishing assumes the template already contains headers and correct worksheet layouts.
- Publish tables must be shaped to match the target worksheet columns in the expected order.
- `.xlsm` templates are best handled via `xlwings` if macro/shapes must be preserved.

