# SQL Transform

The transform phase is where most of the “real” work happens: coercion, mapping, enrichment, and shaping data to match the Workday workbook layout and approved values.

The Python code is a generic SQL runner; domain logic should typically live in SQL.

## Ordering (`order.txt`)

Each pipeline has a transform directory, for example:

- Finance supplier: [finance/supplier/order.txt](config/transform/sql/finance/supplier/order.txt)
- Finance customer: [finance/customer/order.txt](config/transform/sql/finance/customer/order.txt)
- HCM contingent worker: [hcm/contingent_worker/order.txt](config/transform/sql/hcm/contingent_worker/order.txt)

The file lists SQL paths to execute in order, typically:

1. Shared reference data loads from `../../ref/*.sql`
2. Shared macros from `../../macro/*.sql`
3. Pipeline staging (`src_*`)
4. Publish tables (`workday_*`)
5. Optional `validation.sql`

Execution stops on the first failing statement, so ordering is part of the contract.

## Shared reference data (`ref/`)

The directory `config/transform/sql/ref/` contains SQL that loads static mapping/lookup tables into DuckDB (often from files under `config/data/`).

These tables are intentionally shared across multiple pipelines, for example:

- Country and region mappings
- Payment terms mappings
- Source → Workday code/value mappings

## Shared macros (`macro/`)

The directory `config/transform/sql/macro/` defines reusable SQL functions that are referenced by pipeline SQL (for example, `fn_initcap.sql`).

This keeps pipeline SQL files smaller and makes behaviour consistent across domains.

## Naming conventions

The shipped pipelines use these conventions:

- `src_*`: raw/staging tables derived from ingested tables
- `workday_*`: publish tables intended to map 1:1 to output worksheets
- `ref_*`: shared reference tables populated from static data files

## DuckDB UDFs

Some transforms depend on DuckDB UDFs registered by the Python code (for example phone normalisation).

UDFs are registered during transform DB connection setup in [db.py](src/elt_ingest_excel/transform/db.py) and implemented under [transform/udf](src/elt_ingest_excel/transform/udf/).

If SQL fails with “function does not exist”, check:

- The UDF module is imported/registered in `open_connection()`
- The pipeline is running the TRANSFORM phase (UDFs are registered on that connection)

## Fuzzy name similarity

`jaro_winkler_similarity(a, b)` is a built-in DuckDB function — no extension required. It is used by `fin_customer_splink_compare.py` for post-pipeline near-duplicate analysis and is not part of the transform SQL itself.

The `src_fin_*_raw.sql` files use deterministic REGEXP_REPLACE normalisation only (no fuzzy logic at transform time). Fuzzy matching runs as a separate analysis step after the pipeline has been run to at least TRANSFORM.

## Currency / decimal amounts

Source amount columns are ingested as `VARCHAR` and may carry more decimal places than standard 2dp currency (for example General Ledger source data carries 5dp). Two rules apply when casting these to numeric:

- Cast to `DECIMAL(p,s)`, never `DOUBLE`. `DOUBLE` is binary floating point and cannot represent most decimal fractions exactly; rounding a `DOUBLE` per-row and then summing across many rows (e.g. debit/credit totals) produces a non-zero residual even when the source balances exactly. See [workday_fin_gl_line_base.sql](config/transform/sql/finance/general_ledger/pnl_bs_by_month/workday_fin_gl_line_base.sql) for the corrected pattern.
- Do not round to 2dp at the row level before a debit/credit (or any sign-based) split. Carry the source's native precision through (`DECIMAL(20,5)` for GL) and only round at final presentation, if at all — early rounding repeated across many rows is what breaks balance checks, not the final display value.

