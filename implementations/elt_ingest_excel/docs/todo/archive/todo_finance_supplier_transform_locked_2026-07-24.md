# Finance Supplier Transform — LOCKED

## Resume (start here)

This workstream is closed and locked.

Do not make further supplier transform changes from this backlog. If new evidence appears, treat it
as a reopen decision and escalate explicitly before changing SQL, reference data, or publish logic.

## Closed Workstream

Backlog Item: Finance supplier transform completion and production lock

Done
- Published the production supplier workbook via `uv run python examples/fin_supplier_creditor.py --run-to-phase PUBLISH`.
- Confirmed production uses the full supplier population; the historic "already loaded" exclusion is intentionally disabled.
- Resolved the positional publish bug on Supplier Address by replacing `SELECT *` with an explicit workbook-aligned column list.
- Completed the transform-side bank-code recovery from IBAN for the remaining derivable cases.
- Completed the address-field fixes needed for `city_subdivision` / `subregion` handling in the supplier pipeline.
- Completed the production handoff review against the prior finance issue workbook.
- Moved the active supplier continuity doc to reconciliation-only status and archived the transform work in this file.

Still todo (next actions)
1) None. This transform workstream is closed and locked.

Accumulated Active Constraints (active for all remaining items)
- The supplier transform is frozen for normal follow-up work.
- The Excel publisher maps columns positionally, not by column name. Publishing SQL must use explicit workbook-aligned column lists whenever this workstream is ever reopened.
- `supplier_id` is run-scoped. For cross-run comparisons or joins to external review files, use GP Vendor ID / `reference_id`, not the `S-` id.
- Production supplier loads are full-population loads; test-environment exclusion rules are intentionally not applied.
- UK bank-name derivation is not authoritative from the current 2-digit sort-code-prefix mapping. Do not change GP-vs-prefix precedence without authoritative 6-digit reference data and an explicit reopen decision.

Verification
- `uv run python examples/fin_supplier_creditor.py --run-to-phase PUBLISH`
- `uv run python examples/fin_supplier_creditor_validate_sql.py`
- `SELECT rule_type, col, COUNT(*) FROM validation_supplier_issues GROUP BY 1,2 ORDER BY 3 DESC;`

Gotchas
- DuckDB is often locked by DBeaver; close it before rerunning publish or validation.
- Historic Workday Supplier IDs from older finance workbooks are not stable join keys.
- The supplier bank-name question is a data-authority problem, not a simple SQL precedence tweak.
