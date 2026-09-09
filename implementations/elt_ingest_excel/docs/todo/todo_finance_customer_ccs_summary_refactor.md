# Finance Customer CCS Summary — deterministic-generator refactor — TODO

## Resume (start here)

- Status: NOT STARTED. Goal: make the customer CCS reconcile pack's executive `summary` tab
  script-reproducible, the way the supplier pack already is.
- Problem: the `summary` tab in `~/Downloads/customer_ccs_reconcile_review_pack.xlsx` was
  hand-authored — there is NO committed generator. Re-running
  `examples/fin_customer_ccs_reconcile_sql.py` produces a pack whose first tab is the plain
  `board_summary` table (per `config/publish/finance/publish_customer_ccs_reconcile_review.json`),
  NOT the rich narrative summary that ships in Downloads today. So the customer pack cannot
  currently be regenerated from code.
- Before touching code, read `examples/supplier_ccs_summary_sheet.py` (the reference
  implementation) and open the gold target `~/Downloads/customer_ccs_reconcile_review_pack.xlsx`
  sheet `summary` (A1:E101) to inventory each row as data-derived vs template text.

## Goal

1. Build `examples/customer_ccs_summary_sheet.py` with `build_summary_sheet(pack_path, database_path)`
   that reads live from the `reconcile_customer_ccs_*` tables and reproduces the current summary
   tab, then wire it into `fin_customer_ccs_reconcile_sql.py` so it is prepended on every export
   (mirroring how the supplier runner calls `_export_pack_with_summary`).
2. Extract the shared styling primitives (palette + row helpers `span/kv/section/note`) out of
   `supplier_ccs_summary_sheet.py` into `examples/ccs_summary_common.py` so both summary builders
   use ONE styling source (code-reuse gate). Supplier builder must keep producing identical output
   after the extraction.

## Reference: the gold target

Sheet `summary`, 5 columns, widths A=4.6 B=25.1 C=38.9 D=12.6 E=52.6 F=3.4, gridlines off.
Palette: title `1F3864`/white; section header `2E5496`/white 12pt; table header `D9E1F2`/`1F3864`;
outcome band `E2EFDA`/`548235`; amber action `FCE4D6`/`C55A11`; grey label `F2F2F2`; note `808080`.
Nine sections: 0 How performed · 1 Population · 2 Core attributes · 3 Tax reconciliation ·
4 Address & contact · 5 Fixes required in Workday · 6 Reviewed & accepted · 7 Status · 8 Approvals.
Data sources: `reconcile_customer_ccs_report` (numbers), `..._board_summary`,
`..._tax_review_operator_pack` (§3 dispositions), `..._contact_*` (§4 buckets),
`..._closure_summary`/`..._closure_decision_pack` (§5/§7).

## Accumulated Active Constraints

- Every number in the sheet MUST be read live from a `reconcile_customer_ccs_*` table via a
  `_fetch()`-style query — no baked literals. The supplier refactor already removed the last baked
  annotations (`CHR096`/`TGI003`/counts); do not reintroduce that pattern here.
- Approvers for the customer pack are `Finance Operations` / `Najam Israr` (Responsible Finance
  Team Approver) / `Joel Monaghan` (Group Financial Controller) — same as the supplier pack.
- The shared-helper extraction must not change supplier output. Regenerate the supplier pack and
  confirm the `summary` tab is unchanged before and after.
- Reconcile reads upstream `workday_customer_*` transform outputs, which come from the LOCKED,
  force-gated customer transform (`ALLOW_FIN_CUSTOMER_DEBTOR_RUN=1`). This refactor is
  summary-generation only — do NOT run or modify the customer transform.

## Verification

- Classify every summary row as data vs template FIRST (write the inventory into this file).
- `uv run python examples/fin_customer_ccs_reconcile_sql.py --export-review-workbook-only --export-review-dir ~/Downloads`
  then cell-by-cell text-compare the regenerated `summary` tab against a saved copy of the current
  gold file; every §-number must equal the matching `reconcile_customer_ccs_report` metric.
- `uv run python examples/fin_supplier_ccs_reconcile_sql.py --export-review-workbook-only --export-review-dir /tmp`
  before and after the shared-helper extraction; diff the two `summary` tabs — expect no change.
- `uv run --project . python -m py_compile examples/customer_ccs_summary_sheet.py examples/ccs_summary_common.py examples/supplier_ccs_summary_sheet.py`

## Gotchas

- Save a frozen copy of the current gold file first (e.g. to scratchpad) — the export overwrites
  `~/Downloads/customer_ccs_reconcile_review_pack.xlsx`, destroying the hand-authored reference.
- Some §3/§4 lines are editorial judgement (e.g. Coventry→Cornwall, tax-wording dispositions).
  If a line cannot be derived from a reconcile table, carry it as explicit template text and note
  it here — do not silently invent a derivation that could drift.
- DuckDB is often locked by DBeaver; close it before regenerating.
- `board_summary` remains a separate tab; the new `summary` tab is prepended at index 0, same as
  supplier. Decide whether to drop `board_summary` from the customer publish config for parity.
