# PRD (Product Requirements) — elt_ingest_excel

This document is written to give an AI coding assistant enough context to make safe changes to `elt_ingest_excel` without re-learning the domain every time.

## 1. Purpose

Produce Workday load workbooks (typically macro-enabled templates supplied by 3+) from source Excel extracts by:

- Loading source sheets into DuckDB
- Applying deterministic SQL transformations
- Populating the target worksheets in cloned Workday templates

## 2. Users and stakeholders

- Data/technical analysts: run pipelines, adjust mappings, validate outputs
- Engineers: maintain pipeline plumbing and add new pipelines/domains
- Downstream recipients: Workday/implementation partner receiving the populated workbooks

## 3. Problem statement

Source extracts are messy and inconsistent. Workday load workbooks require:

- Specific worksheet structures and column ordering
- Approved values (countries, regions, categories, terms, etc.)
- Normalised formats (phones, addresses, IDs)

Manual preparation is slow and error-prone; the pipeline must make the transformation reproducible and auditable.

## 4. Scope

In scope:

- Config-driven ingestion from Excel worksheets into DuckDB
- Ordered SQL transformations producing publish tables
- Cloning and populating template workbooks from DuckDB publish tables
- Shared reference data tables and reusable SQL “macros”
- Optional UDFs to support transform logic (for example phone parsing)

Out of scope:

- Direct API integration with Workday
- “Reverse engineering” or generating EIB outputs from hidden workbook logic
- A full UI; pipelines are executed as scripts/CLI entry points

## 5. Core concepts (domain model)

- **Ingest tables**: raw tables created from source sheets (strings, minimal interpretation)
- **Reference tables (`ref_*`)**: mapping/lookup tables loaded from repo-controlled data files
- **Staging tables (`src_*`)**: intermediate transforms derived from ingest tables
- **Publish tables (`workday_*`)**: final tables that map to target workbook worksheets
- **Template workbook**: the macro/formatted workbook used as the basis for output

## 6. Functional requirements

### FR1 — Ingest Excel into DuckDB

- Must load configured workbooks/sheets into DuckDB tables.
- Must support specifying `headerRow` and `dataRow` for each sheet.
- Must normalise column names into safe DuckDB identifiers.
- Must not require schema definitions up front; typing belongs in SQL.

### FR2 — Deterministic SQL transforms

- Must execute SQL files in a deterministic order defined by `order.txt`.
- Must support multi-statement SQL files.
- Must fail fast on SQL errors and halt the pipeline before publishing.
- Must allow shared SQL steps (reference loads and macros) reused across pipelines.

### FR3 — Publish to Excel templates

- Must clone template workbook(s) to create output workbook(s).
- Must populate each configured sheet by reading a DuckDB table/view and writing rows starting at `dataRow`.
- Must preserve template formatting/macros by relying on the template for headers and styling.
- Must support `.xlsm` templates; when preservation is critical, publishing should be able to run through Excel automation.

### FR4 — Pipeline entry points

- Must provide a small number of clear, one-file-per-pipeline entry points under `examples/`.
- Must support running to a phase (INGEST, TRANSFORM, PUBLISH) for iterative development/debugging.

### FR5 — Reference export workbook (outlier)

- Must be able to export the reference mapping tables to a human-friendly `.xlsx` workbook for review/editing.

## 7. Non-functional requirements

- **Reproducibility**: given the same inputs and configs, output should be stable.
- **Traceability**: transformation logic should be discoverable in SQL and config, not hidden in Python.
- **Safety**: prevent executing arbitrary SQL from outside the repo-controlled transform directories.
- **Maintainability**: adding a new pipeline should usually mean adding config + SQL, not new Python.

## 8. Interfaces and contracts

### Ingest config (JSON)

- Contract described in [docs/configs.md](docs/configs.md).
- Interpreted relative to the `data_path` used by the runner script.

### Transform contract

- A transform directory containing `order.txt` and referenced SQL files.
- Shared SQL is under `config/transform/sql/ref/` and `config/transform/sql/macro/`.

### Publish config (JSON)

- Contract described in [docs/configs.md](docs/configs.md).
- Declares template workbook path/name and output path/name, plus sheet mappings.

## 9. Current “primary” pipelines

The key supported pipelines are:

- Finance customer/debtor: [fin_customer_debtor.py](examples/fin_customer_debtor.py)
- Finance supplier/creditor: [fin_supplier_creditor.py](examples/fin_supplier_creditor.py)
- HCM contingent worker: [hcm_contingent_worker.py](examples/hcm_contingent_worker.py)
- Reference export: [fin_workday_ref.py](examples/fin_workday_ref.py)

## 10. Acceptance criteria (assistant checklist)

When changing the module, validate:

- A pipeline can run to INGEST without errors on a representative dataset.
- A pipeline can run to TRANSFORM and produce publish tables (`workday_*`) as expected.
- Publishing writes the expected row counts to the expected worksheets starting at `dataRow`.
- If changing SQL ordering or shared ref tables, downstream pipelines still run.
- Tests pass (`pytest test/`).

