# CLI

## Prerequisites

- Python ≥ 3.11
- `uv` package manager: https://docs.astral.sh/uv/
- Excel (macOS/Windows) — required only for the `xlwings` publisher (default)

## Setup

```bash
cd implementations/elt_ingest_excel
uv sync
```

## Running

All pipelines accept `--run-to-phase INGEST|TRANSFORM|PUBLISH`. The default is runner-specific, so check `--help` when in doubt.

Optional flags accepted by all pipeline runners:

- `--database-path /path/to/db.duckdb` — override the DuckDB path from `config/runtime/default.json`
- `--runtime-config filename.json` — use an alternate runtime config under `config/runtime/`

```bash
# Finance – Supplier/Creditor
uv run python examples/fin_supplier_creditor.py [--run-to-phase INGEST|TRANSFORM|PUBLISH]

# Finance – Customer/Debtor
uv run python examples/fin_customer_debtor.py [--run-to-phase INGEST|TRANSFORM|PUBLISH]

# HCM – Contingent Worker
uv run python examples/hcm_contingent_worker.py [--run-to-phase INGEST|TRANSFORM|PUBLISH]

# HCM – Learn (scaffold: ingest workbook, transform mappings and publish template are placeholders)
uv run python examples/hcm_learn.py [--run-to-phase INGEST|TRANSFORM|PUBLISH]

# Finance – Workday CCS Customer Extract
uv run python examples/fin_workday_extract_ccs_customer.py [--run-to-phase INGEST|TRANSFORM|PUBLISH]

# Finance – Workday CCS Customer Extract + Reconcile
uv run python examples/fin_customer_ccs_reconcile_sql.py [--run-to-phase INGEST|TRANSFORM]
uv run python examples/fin_customer_ccs_reconcile_sql.py [--export-review-workbook|--export-review-workbook-only]

# Finance – Workday CCS Supplier Extract + Reconcile (refreshes extract, loads manual-update list)
uv run python examples/fin_supplier_ccs_reconcile_sql.py [--run-to-phase INGEST|TRANSFORM]
uv run python examples/fin_supplier_ccs_reconcile_sql.py [--export-review-workbook|--export-review-workbook-only]
uv run python examples/fin_supplier_ccs_reconcile_sql.py [--manual-update-file <path>]
```

### Override run prevention
### Check offset delta for customer/supplier
```bash
export ALLOW_FIN_SUPPLIER_CREDITOR_RUN=1
export ALLOW_FIN_CUSTOMER_DEBTOR_RUN=1
export ALLOW_HCM_CONTINGENT_WORKER_RUN=1
```

## Examples
```bash
cd implementations/elt_ingest_excel
uv sync

uv run python examples/hcm_learn.py                             --run-to-phase TRANSFORM
uv run python examples/fin_customer_invoice.py                  --run-to-phase PUBLISH
uv run python examples/fin_customer_invoice_adjustment.py       --run-to-phase PUBLISH
uv run python examples/fin_customer_workday_eib_result.py       --run-to-phase INGEST
uv run python examples/fin_general_ledger.py                    --run-to-phase PUBLISH
uv run python examples/fin_workday_extract.py                   --run-to-phase TRANSFORM
uv run python examples/fin_workday_extract_ccs_customer.py      --run-to-phase TRANSFORM
uv run python examples/fin_customer_ccs_reconcile_sql.py        --run-to-phase TRANSFORM
uv run python examples/fin_workday_ref.py

# Needs to be locked
uv run python examples/fin_supplier_invoice.py                  --run-to-phase PUBLISH
uv run python examples/fin_supplier_purchase_order_receipt.py   --run-to-phase PUBLISH

# Locked - needs force run
uv run python examples/fin_supplier_creditor.py                 --run-to-phase INGEST
uv run python examples/fin_customer_debtor.py                   --run-to-phase TRANSFORM
uv run python examples/hcm_contingent_worker.py                 --run-to-phase PUBLISH


# Validation only (no ingest/transform/publish) — re-run against an already-populated database
uv run python examples/fin_customer_debtor_validate_sql.py
uv run python examples/fin_supplier_creditor_validate_sql.py

# Reconciliation only (no ingest/transform/publish) — confirms no data loss in the transform
uv run python examples/fin_customer_debtor_reconcile_sql.py
uv run python examples/fin_supplier_creditor_reconcile_sql.py

# Commitment stack source→target reconciliation
uv run python examples/fin_supplier_po_commitment_stack_reconcile_sql.py [--export-review-workbook]

# PO/receipt post-load tie-out (4th layer): sent EIB vs Workday-loaded extract.
# Reads workday_supplier_po_commitment_stack(_receipt) + the loaded-extract xlsx; all live,
# no hardcoded totals; writes FA_PostLoad_Tieback_YYYYMMDD[_LABEL].xlsx to the reconcile dir.
# Off-pipeline (reads the extract directly) — run after PO/receipt TRANSFORM.
# Default = the earlier test control report; pass --control-report/--label for other extracts.
uv run python scripts/generate_post_load_tieback.py
uv run python scripts/generate_post_load_tieback.py \
  --control-report "~/Downloads/FA Find DC Purchase Orders PRODUCTION 9 August 2026 1.xlsx" \
  --label PRODUCTION            # sheet auto-resolves (named sheet if present, else first)

# CCS extract refresh + reconciliation in one command
uv run python examples/fin_customer_ccs_reconcile_sql.py        --run-to-phase TRANSFORM
uv run python examples/fin_supplier_ccs_reconcile_sql.py        --run-to-phase TRANSFORM --export-review-workbook

# CCS review workbook export from current DuckDB state only
uv run python examples/fin_customer_ccs_reconcile_sql.py        --export-review-workbook-only
uv run python examples/fin_supplier_ccs_reconcile_sql.py        --export-review-workbook-only
```

## Running Excel VBA validations (macOS)

Run the validation macros that 3+ uses to validate the generated `.xlsm` workbook. This allows you to find and fix data quality issues before sending the workbook to 3+.

**Prerequisites:**

- Microsoft Excel installed (macOS)
- Macros enabled for the workbook

**One-time setup: Unhide ValidationManifest sheet**

The ValidationManifest sheet contains the validation rules but is hidden by default. Before running validations, unhide it and add rules for your sheets:

```bash
cd elt_ingest_excel

# Unhide ValidationManifest sheet
uv run python examples/unhide_validation_manifest.py --excel-visible
```

Then in Excel:

1. Open the `ValidationManifest` sheet (now visible at the end of sheet tabs)
2. Add validation rules for your data sheets (e.g., "Supplier Name")
3. Save the workbook

**Run Validation:**

```bash
cd elt_ingest_excel

# Run all validations
uv run python examples/fin_supplier_creditor_validate_vba.py

# Run specific validation macro
uv run python examples/fin_supplier_creditor_validate_vba.py \
  --macro runSpecificValidationsFromSheet

# Run with Excel visible (to watch progress)
uv run python examples/fin_supplier_creditor_validate_vba.py --excel-visible

# Override workbook path
uv run python examples/fin_supplier_creditor_validate_vba.py \
  --workbook /path/to/other_workbook.xlsm
```

**Complete workflow:**

```bash
# 1. Run your ETL pipeline
uv run python examples/fin_supplier_creditor.py --run-to-phase PUBLISH

# 2. Run validation (removes 3+ latency)
uv run python examples/fin_supplier_creditor_validate_vba.py

# 3. Open workbook and review "Validation Results" sheet
#    Fix any data issues if needed, then repeat from step 1
```

For full documentation see [VBA_VALIDATION_RUNNER.md](../VBA_VALIDATION_RUNNER.md).

## Running DuckDB SQL validations

Run the repo SQL validations under `config/validate/...` directly against the DuckDB database (these are separate from the workbook’s VBA validations).

```bash
cd elt_ingest_excel
uv sync
uv run python examples/fin_customer_debtor_validate_sql.py
uv run python examples/fin_supplier_creditor_validate_sql.py
```

These use the DuckDB path from `config/runtime/default.json` unless overridden:

```bash
uv run python examples/fin_customer_debtor_validate_sql.py --database-path /path/to/db.duckdb
uv run python examples/fin_supplier_creditor_validate_sql.py --database-path /path/to/db.duckdb
```

## Running DuckDB SQL reconciliation

Run the repo SQL reconciliation checks under `config/reconcile/...` directly against the DuckDB database. Unlike validation (which checks data quality of the final workday tables), reconciliation confirms that the source-to-workday transform itself didn't lose data: row counts at each pipeline stage, address/bank/email/phone attribute coverage after dedup, and any supplier/customer dropped by a transform filter. Full mechanism, known gotchas, and cold-start reproduction steps: [reconciliation.md](reconciliation.md).

Run validation first, then reconciliation — the consolidated report's GP ID section reads `validation_{customer,supplier}_issues`, which is only populated by `validate_{customer,supplier}_id_coverage.sql`:

```bash
cd elt_ingest_excel
uv sync
uv run python examples/fin_customer_debtor_validate_sql.py
uv run python examples/fin_customer_debtor_reconcile_sql.py
uv run python examples/fin_supplier_creditor_validate_sql.py
uv run python examples/fin_supplier_creditor_reconcile_sql.py
uv run python examples/fin_customer_ccs_reconcile_sql.py --run-to-phase TRANSFORM
```

These accept the same `--database-path` / `--runtime-config` overrides as the validation scripts. The one table to check is `reconcile_{customer,supplier}_report` — it consolidates the GP ID coverage confirmation, the row-count funnel, and attribute coverage issue counts into a single report. The underlying detail lives in `reconcile_{customer,supplier}_row_funnel`, `reconciliation_{customer,supplier}_issues`, and `z_reconcile_{customer,supplier}_summary`.

For supplier **PO commitment stack**, `examples/fin_supplier_po_commitment_stack_reconcile_sql.py` reconciles the Commitment Stack source against the Workday PO EIB tables populated by the active pipeline (`config/reconcile/finance/supplier_purchase_order_receipt`). It covers the active `reconcile_cs_*` suite: **row counts** (`reconcile_cs_row_counts`), **amount by currency** (`reconcile_cs_amount_by_currency`), **currency coverage** (`reconcile_cs_currency_coverage`), **drop buckets** such as unmapped suppliers / zero commitments / old soft cutoff, plus the Fraser audit comparison set (`reconcile_cs_audit_*`). The consolidated review tables are `reconcile_cs_report` and `reconcile_cs_board_summary`; `--export-review-workbook` writes the configured PO commitment-stack review pack. Prerequisite: the supplier PO/receipt pipeline has been run to `TRANSFORM` first.

For CCS customer, `examples/fin_customer_ccs_reconcile_sql.py` is the end-to-end entry point: it first refreshes `src_fin_workday_extract_ccs_customer` via the `workday_ccs_customer.json` ingest plus `transform/sql/finance/workday_extract_ccs_customer` SQL, then immediately runs the reconciliation pack under `config/reconcile/finance/customer_ccs`. Run to `TRANSFORM` for the normal full refresh + reconcile flow; `INGEST` refreshes the raw extract tables only and skips reconciliation. After the reconcile step, the runner prints a compact CCS summary and points you to the main review tables:

- `reconcile_customer_ccs_provenance_case_summary`
- `reconcile_customer_ccs_tax_review_operator_pack`
- `reconcile_customer_ccs_report`

For the finished CCS reconciliation workstream, the same runner also supports review-pack export:

```bash
# Full CCS extract refresh + reconciliation + review workbook export
uv run python examples/fin_customer_ccs_reconcile_sql.py --run-to-phase TRANSFORM --export-review-workbook

# Review workbook export only from the current DuckDB reconciliation state
uv run python examples/fin_customer_ccs_reconcile_sql.py --export-review-workbook-only
```

By default this rewrites:

- `~/Documents/__data/excel/workday/reconcile/customer_ccs_reconcile_review_pack.xlsx`

The first tab is `board_summary`: a flat board-level rollup whose values are pulled live from the reconcile tables. Filter its `disposition` column to `FIX_IN_WORKDAY` for the Finance action list.

## Fuzzy name-match comparison (customer)

Self-joins `src_fin_customer_raw` on Jaro-Winkler name similarity and writes all pairs above the threshold to `customer_fuzzy_name_comparison` in the pipeline database. Only pairs with different `nrm_customer_name` are returned — these are the cases the deterministic normalisation did not group.

Requires the pipeline to have been run to at least the TRANSFORM phase first.

```bash
uv run python examples/fin_customer_splink_compare.py
uv run python examples/fin_customer_splink_compare.py --threshold 0.92
```

Database path is read from `config/runtime/default.json`. Override with `--database-path` if needed. `name_similarity` is a direct Jaro-Winkler score (0–1), not a model probability.

**Why not Splink:** Splink's unsupervised EM training was evaluated against this dataset and produced a flat prior — every pair scored 0.311–0.317 regardless of similarity. EM has no signal on a small, unlabelled customer list and outputs the base rate. A direct `jaro_winkler_similarity` self-join is simpler, transparent, and gives a real score to reason about.

**Threshold guidance:**

| Threshold | Typical content |
|---|---|
| ≥ 0.99 | Typos, spacing artifacts — high confidence near-duplicates |
| 0.95–0.99 | Men's vs women's club variants (FC / WFC / LFC) — legitimately different entities |
| < 0.92 | Personal-name noise — different people sharing a first name score here due to Jaro-Winkler prefix bias |

Default is 0.95. Lowering below 0.92 adds significant noise from personal names.

Suggested queries after running:

```sql
SELECT * FROM customer_fuzzy_name_comparison ORDER BY name_similarity DESC;
SELECT * FROM customer_fuzzy_name_comparison WHERE name_similarity >= 0.99 ORDER BY name_similarity DESC;
```

## Creating the EIB workbook

The Supplier/Creditor `.xlsm` workbook contains an `InitiateSPECAdvanceLoad` macro that generates the EIB output (same as the "Create Specific EIBs" button on the EIB View sheet).

**One-time prerequisite:** The `TempCreator` VBA module has `Option Private Module` which blocks external calls. Remove it once:

1. Open the workbook in Excel
2. Press `Option+F11` to open the VBA Editor
3. In the Project pane: expand the workbook → Modules → TempCreator
4. Delete the line: `Option Private Module`
5. Save (`Cmd+S`) and close the VBA Editor

```bash
cd elt_ingest_excel

# Generate EIB workbook (saves in-place, unhides EIB View sheet)
uv run python examples/fin_supplier_creditor_create_eib.py

# Run with Excel visible (to watch progress)
uv run python examples/fin_supplier_creditor_create_eib.py --excel-visible

# Override workbook path
uv run python examples/fin_supplier_creditor_create_eib.py \
  --workbook /path/to/other_workbook.xlsm
```

After the macro completes, check the EIB View sheet in the workbook for the generated EIB data.

**Full workflow with EIB creation:**

```bash
# 1. Run the ETL pipeline
uv run python examples/fin_supplier_creditor.py --run-to-phase PUBLISH

# 2. Run validations
uv run python examples/fin_supplier_creditor_validate_vba.py

# 3. Create EIB workbook
uv run python examples/fin_supplier_creditor_create_eib.py
```

## Tests

```bash
uv run pytest test/ -v
```
