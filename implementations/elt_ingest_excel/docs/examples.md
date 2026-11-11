# Examples (entry points)

The scripts under `elt_ingest_excel/examples/` are the intended entry points. They:

- Pick the ingest config
- Pick the transform directory
- Pick the publish config
- Provide a consistent `--run-to-phase INGEST|TRANSFORM|PUBLISH` flag

## Primary pipelines

- Finance supplier/creditor: [fin_supplier_creditor.py](examples/fin_supplier_creditor.py)
  - Ingest: `config/ingest/finance/active_supplier_creditor.json`
  - Transform: `config/transform/sql/finance/supplier/`
  - Publish: `config/publish/finance/publish_supplier.json`
- Finance customer/debtor: [fin_customer_debtor.py](examples/fin_customer_debtor.py)
  - Ingest: `config/ingest/finance/active_customer_debtor.json`
  - Transform: `config/transform/sql/finance/customer/`
  - Publish: `config/publish/finance/publish_customer.json`
- HCM contingent worker: [hcm_contingent_worker.py](examples/hcm_contingent_worker.py)
  - Ingest: `config/ingest/hcm/contingent_worker.json`
  - Transform: `config/transform/sql/hcm/contingent_worker/`
  - Publish: `config/publish/hcm/publish_contingent_worker.json`
- HCM learn (scaffold — source workbook, Workday target mappings, and publish template are placeholders): [hcm_learn.py](examples/hcm_learn.py)
  - Ingest: `config/ingest/hcm/learn.json`
  - Transform: `config/transform/sql/hcm/learn/`
  - Publish: `config/publish/hcm/publish_learn.json`
  - Validate: `config/validate/hcm/learn/` (results in `validation_learn_issues`)

## Reference export (outlier)

- Finance Workday reference export: [fin_workday_ref.py](examples/fin_workday_ref.py)

This script is intentionally different: it reads the publish config but creates a brand-new `.xlsx` workbook and writes reference tables into it with formatting. It does not use the standard `FileIngestor` publish step.

Details: [reference-export.md](reference-export.md).

## SQL validation runners

Standalone scripts to run the DuckDB validation SQL against an existing database without re-running the full pipeline:

- [fin_customer_debtor_validate_sql.py](examples/fin_customer_debtor_validate_sql.py): runs `config/validate/finance/customer/` SQL files; results in `validation_customer_issues`
- [fin_supplier_creditor_validate_sql.py](examples/fin_supplier_creditor_validate_sql.py): runs `config/validate/finance/supplier/` SQL files; results in `validation_supplier_issues`

Both accept `--database-path` to override the path from `config/runtime/default.json`.

## Fuzzy name-match comparison (analysis)

- [fin_customer_splink_compare.py](examples/fin_customer_splink_compare.py): identifies customer name pairs that are the same entity under different spellings

Self-joins `src_fin_customer_raw` on Jaro-Winkler name similarity and writes pairs above the threshold to `customer_fuzzy_name_comparison` in the pipeline database. Only pairs with **different** `nrm_customer_name` are returned — these are cases the deterministic normalisation did not group. `name_similarity` is a direct Jaro-Winkler score (0–1), not a model probability.

**Why SQL, not Splink:** Splink's unsupervised EM training was evaluated on the customer dataset and produced a flat prior: every pair scored 0.311–0.317 with no discrimination. EM has no training signal on a small, unlabelled list — it outputs the base rate. A direct `jaro_winkler_similarity` self-join is simpler, transparent, and gives a real score to reason about.

**Threshold guidance:**

| Threshold | Typical content |
|---|---|
| ≥ 0.99 | Typos and formatting artifacts — high confidence genuine near-duplicates |
| 0.95–0.99 | Men's vs women's club variants (FC / WFC / LFC) — legitimately different entities |
| < 0.92 | Personal-name noise — different people sharing a first name score here due to Jaro-Winkler's prefix bias |

Default threshold is 0.95. Override with `--threshold`.

## Legacy / on-hold helpers

Some scripts exist for Excel/VBA automation around the 3+ templates (validations, EIB macro execution, and unhiding hidden sheets). These are useful when they match the current 3+ workbook behaviour, but they are not part of the core ELT pipeline and may not be maintained.

- [fin_supplier_creditor_create_eib.py](examples/fin_supplier_creditor_create_eib.py): triggers the “Create Specific EIBs” macro flow
- [fin_supplier_creditor_validate_vba.py](examples/fin_supplier_creditor_validate_vba.py): runs the workbook’s validation macros
- [unhide_validation_manifest.py](examples/unhide_validation_manifest.py): unhides the `ValidationManifest` sheet

If these remain on hold, a low-friction way to reduce confusion is:

- Move them under `examples/_legacy/` (or similar)
- Update the top-level README to point to them as optional utilities rather than required steps
