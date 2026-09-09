# Review — AR Invoice / AR Discipline Invoice transform modules

Read this before touching `modArInvoiceCustomerInvoice.bas` or `modArDisciplineCustInvoice.bas`.
Captured from a code review on 2026-06-24 of the work completed under
`config/transforms/{ar_invoice,ar_discipline_invoice}.yaml`. No code was changed by this review —
these are open items to address later, not confirmed bugs.

## Open items

1. **`NormalizeSourceDate` locale dependency for native Excel dates.**
   Discipline Invoice's source (`Discipline Invoices.xlsx`) stores `Document Date` as a real Excel
   `Date` type, not text (confirmed via openpyxl, e.g. `datetime(2026, 6, 1)`). `CStr()` on a Date
   value renders it using the host machine's regional date format. The DMY split-parser in
   `NormalizeSourceDate` (now a single shared function in `vba_src/modSourceTable.bas`, used by every
   transform) consumes that locale-rendered string before ever reaching the `IsDate`/`CDate` fallback. On a UK-locale machine (`dd/mm/yyyy`) this resolves
   correctly; on a US-locale machine (`mm/dd/yyyy`) dates where day <= 12 would silently swap
   day/month (e.g. `2026-06-01` -> `2026-01-06`). Everything else in this codebase (sample data,
   currency, company) assumes a UK deployment, so this is likely a non-issue in practice today —
   but it is an implicit assumption, not an enforced one. Consider either documenting the UK-locale
   requirement explicitly (e.g. in `docs/RUNBOOK.md` prerequisites) or making `NormalizeSourceDate`
   check `IsDate(sourceValue)` first for genuine Date-typed cells (bypassing the string heuristic
   entirely) and only falling back to DMY string parsing for plain text values such as CSV imports.

2. **`dist/Runner.xlsm` staleness after the eib42 rename.**
   `vba_src/transforms/modArDisciplineCustInvoice.bas` (renamed from
   `modDisciplineInvoiceCustomerInvoice.bas` in eib42, then to the `Ar`-prefixed name later) was last touched after the last
   `dist/Runner.xlsm` build, so the renamed module has not yet been re-imported into the shipped
   workbook. Running `build_runner.py --verify` currently fails with "Missing: Bootstrap Import
   Source, Bootstrap Import Count, Bootstrap Import At" — expected per the documented workflow
   (build -> manual Excel import -> verify), not a code defect, but the manual Excel re-import +
   `--verify` step is still outstanding before this rename is fully closed out.

3. **Minor asymmetry in `RowIsSkippable` (by design, just noting it).**
   AR Invoice skips rows missing Company/Order Date/Customer ID/Gross Price. Discipline Invoice
   skips rows missing Customer Number/Document Date/Amount but does not gate on a blank Document
   Number (used only for the `MT`/Memo column). This appears intentional given the simpler
   discipline mapping, not an oversight — flagged here only so it doesn't get "fixed" by accident
   into false symmetry later.

## Gotchas

(none beyond what's already in `docs/todo/archive/build-elt_transform_excel.md`)
