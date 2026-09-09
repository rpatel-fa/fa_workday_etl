# Anchor Doc — Build elt_transform_excel

## Goal of the current work

Get the Runner workbook's VBA macros executing correctly in real Excel, end to end. This has two
distinct parts that must not be conflated:

1. **Build-time: get every `.bas`/`.cls` module correctly and reliably into the deployed
   `dist/Runner.xlsm`'s VBA project.** This was blocked by, in order: a VBA compile error that turned
   out to be binary corruption from `pyOpenVBA` injecting modules directly into an existing `.xlsm`
   (fixed by never writing the VBA project binary directly outside of a fresh, one-time shell creation);
   then by AppleScript automation of the *import itself* being unreliable on this Mac Excel build across
   multiple distinct failure modes (fixed, per user decision, by making this one step manual). **Done**
   — confirmed working against the live 3-module set (see "Manual import confirmed working..." Done
   entry).
2. **Run-time: get the actual transform macro (`modDispatcher.RunSelectedTransform`) to run correctly
   against real sample data once the modules above are in place.** **Done** — confirmed end to end via
   Tools>Macro against the real (non-test) customer extract, including the result-persistence fix
   (`runnerWorkbook.Save`): reading `dist/Runner.xlsm` directly from disk afterward shows
   `Run Status: Completed`, `Rows Processed: 124`, `Rows Skipped: 3` (see Still todo item 4's Done
   entry). The earlier-session note about the one inconclusive headless `osascript run VB macro` call
   (no error, but `Run Status` never updated, workbook later found corrupt) remains a separate,
   lower-priority open question (headless invocation specifically) — not blocking; the interactive
   Tools>Macro path is fully confirmed.

## Resume (start here)

- **Common-layer consolidation (2026-07-14): shared VBA now has three domain-free modules** —
  `modHelpers.bas` (workbook/control-sheet/run-log plumbing), `modSourceTable.bas` (generic
  header-location + amount/date coercion over the in-memory table), and `modFinanceCommon.bas`
  (reusable finance mappings/derivations). Generic table ops → `modSourceTable`; reusable finance
  rules → `modFinanceCommon`; transform-only column-to-target mapping stays in the profile module.
  See `docs/ARCHITECTURE.md` (Modules) for the authoritative split. Still pending: manual Excel
  re-import + `build_runner.py --verify` to land these in `dist/Runner.xlsm`.
- **Rename (2026-07-14): the Discipline Invoice transform was given the `Ar` prefix to match the
  AR/AP naming convention used by every other transform.** Current names — profile
  `config/transforms/ar_discipline_invoice.yaml`, spec `config/transform_specs/ar_discipline_invoice.yaml`,
  module `vba_src/transforms/modArDisciplineCustInvoice.bas` (`VB_Name = modArDisciplineCustInvoice`),
  `transform.key = ar_discipline_invoice`, `step_list_macro = RunArDisciplineInvoiceCustomerInvoice`.
  Historical entries below use the old `discipline_invoice` / `modDisciplineCustInvoice` names — read
  them as referring to these renamed artefacts.
- **Latest change (2026-07-28): `RunnerShellMapping.xlsx` retired entirely; the runtime `mapping`
  lookup is now generated from config.** The `mapping` sheet inside `RunnerShell.xlsm` is no longer
  hand-authored/preserved — `build_runner.py` regenerates it on every build from two new
  `config/ingest_lookups` ref tables: `ref_company.json` (source→target company names, entity=`company`)
  and `ref_default.json` (per-domain defaults, e.g. currency→GBP, entity=`default`), folded into
  `entity`/`source`/`target` rows via the shared `create_ingest_template.load_lookup` loader — the same
  self-describing reference mechanism as every other lookup. Added `_write_mapping_sheet_block` +
  `_verify_mapping_sheet` to `build_runner.py`. VBA (`modHelpers.LookupMappingValue`, reads
  `entity="company"` only) is unchanged; generated sheet verified byte-identical to the old hand-authored
  one. `RunnerShellMapping.xlsx` deleted. To change a company name/default: edit the ref table + rebuild.
- **Latest change (2026-07-14): transform field mappings moved from `RunnerShellMapping.xlsx` into the
  `mapping:` section of `config/transform_specs/<key>.yaml`.** The per-transform tabs and the entire
  `mapping_content_hash` / `--print-mapping-hashes` build-time drift-gate were removed from
  `build_runner.py` and `create_runner_shell.py`; each `.bas` now mirrors its spec's `mapping:` section,
  so drift is caught in code review (co-located text diff) rather than by a build hash.
  `RunnerShellMapping.xlsx` retains only its `mapping` lookup tab. All hash-gate / mapping-tab Done
  entries below are **history, superseded by this change.**
- The original six "Still todo" items are **Done** (see below). The second vertical slice,
  **Discipline Invoice -> Customer Invoice**, is now also confirmed good in real Excel by the user:
  both the `arinvoice` and `ardisciplineryinvoice` populations produced good/correct output. The next
  unfinished action is now the post-run `build_runner.py --verify` check below.
- Open, lower-priority items not currently blocking anything (do not start without explicit
  redirection): the headless `osascript run VB macro` path for `modDispatcher.RunSelectedTransform`
  remains unverified — see Gotchas.

### Still todo (second vertical slice: Discipline Invoice -> Customer Invoice)

1. ~~Run the manual bootstrap import + transform in real Excel~~ **Done** — user confirmed that both
   the `arinvoice` and `ardisciplineryinvoice` populations are good/correct in the output workbooks,
   which closes the real-Excel validation for the second vertical slice as well as reconfirming the AR
   path after the date-normalization fix.
2. **Re-verify with `build_runner.py --verify`** once the manual run completes, to confirm the
   4-module set and 2-row transform registry persisted to disk.

### Discipline Invoice mapping — now has a live tab, `ardisciplineryinvoice`

**Superseded twice:** this section originally said "no live mapping tab" — the user added a live
`ardisciplineryinvoice` tab, in the same `target_column`/`default`/`source_column`/... declarative
format as `arinvoice`. It then said that tab had no `Mapping Definition Timestamp` cell and therefore
no drift-check — that's also since been resolved: a `P1`/`Q1` pair was added (`2026-06-22 23:30:00`)
and a **build-time** drift-check now covers it, the same mechanism `arinvoice` uses (see the
"Architecture change" Done entry below for the full redesign — the check moved from per-run inside
Excel to build-time in `build_runner.py`, and both mapping tabs moved out of the deployed Runner
workbook into `runner_shell/RunnerShellMapping.xlsx`).

Source: `Discipline Invoices.xlsx` (`Customer Number, Customer Name, Document Number, Document Type,
Document Date, Original Trx Amount`; all 66 rows observed had `Document Type = Payments`, all
`Customer Number` values end in `_D`, e.g. `C00894_D`). Per the user: the customer extract already
carries `_D`-suffixed Reference IDs, so the existing pipe-delimited exact-match lookup (same as AR
Invoice) needs no suffix stripping. Cross-checked the live `ardisciplineryinvoice` tab against
`modDisciplineCustInvoice.bas` row by row — every row matched the code already (`Company`
= fixed `THE FA GROUP - FA`, `Currency` = fixed `GBP`, `Customer` = lookup via Customer Number against
the Customer Spreadsheet's pipe-delimited `Reference ID`, `Invoice Date` = Document Date,
`Amount Due`/`Extended Amount` = Original Trx Amount, `Memo` = Document Number, `Revenue Category`/
`Line Item Description` = `NULL` in the tab, confirmed by the user to mean "leave the target cell
empty" — the code already never writes those two cells, so no change needed) **except one real bug in
the tab itself**: the `ha` (Amount Due) row said `source_column = Gross Price`, a column that does not
exist in `Discipline Invoices.xlsx` (only `Original Trx Amount` does, which `mj` correctly already
used). Confirmed with the user this was a copy-paste leftover from the `arinvoice` tab and fixed it
directly in `runner_shell/RunnerShell.xlsm`'s `ardisciplineryinvoice!C11` (`Gross Price` ->
`Original Trx Amount`), via `openpyxl.load_workbook(..., keep_vba=True)` (cell-only edit, VBA project
binary untouched — same safe pattern `build_runner.py`/`create_runner_shell.py` already use).
Verification: re-loaded the workbook afterward and confirmed `sheetnames` still
`['Runner', 'arinvoice', 'ardisciplineryinvoice', 'mapping']` (no sheet lost) and `ardisciplineryinvoice`
row 11 (`ha`) now reads `Original Trx Amount`. Registered as `config/transforms/discipline_invoice.yaml`
(`discipline_invoice__customer_invoice` selection key).

**Found and fixed a real duplication bug while wiring this up, caught by the user.** The first version
of this change hand-copied the same dropdown metadata into both `ar_invoice.yaml` and
`discipline_invoice.yaml` so the Source Type dropdown would include "Discipline Invoice" — because
`create_runner_shell.py` only ever read one profile (`--profile-config`, default `ar_invoice.yaml`)
for dropdown options, unlike `build_runner.py`, which already scans every profile under
`config/transforms/` for the registry. That is exactly the config-drift risk `governance.md`'s
Config Purity rule exists to prevent: a third transform would require remembering to edit every
existing profile's copy of the list. Fixed by removing that duplicated dropdown metadata from
transform profiles entirely: `create_runner_shell.py` now derives control-sheet options directly from
the registered profiles under `config/transforms/*.yaml`, mirroring `build_runner.py`'s registry scan.
Verification:
`python3 -m py_compile scripts/create_runner_shell.py scripts/build_runner.py` exits 0; a from-scratch
rebuild (`create_runner_shell.py --overwrite` then `build_runner.py`) exits 0, reports `Modules: 4`,
`Bindings: 2`; inspecting the rebuilt shell shows the generated Runner controls are driven from the
registered profiles under `config/transforms/*.yaml`, with no dropdown list duplicated in any profile
YAML.

**Output-collision limitation fixed.** Added a new required control-sheet field, `Output File Path`,
between `Target EIB` and `Mapping Spreadsheet` (per user request — they noticed the collision risk
directly). `modHelpers.CloneTemplateWorkbook` already accepted an optional explicit `outputPath`
argument; both transform modules now read the new control value via
`ReadControlValue(controlSheet, CONTROL_LABEL_OUTPUT_FILE_PATH)`, validate it non-blank with a new
`modHelpers.EnsureNonBlank` helper, and pass it through instead of relying on the old
auto-derived-from-`Target EIB` `"_output"` suffix. `create_runner_shell.py`'s `CONTROL_LABELS` list
and `_build_control_rows` now include `Output File Path`; the previously hardcoded dropdown cell
positions (`B3`/`B10`) are now computed from `CONTROL_LABELS.index(...)` so inserting a new control
row doesn't silently misplace the dropdowns. `ar_invoice.yaml` and `discipline_invoice.yaml` each set
a distinct `shell_defaults.output_file_path`
(`..._ar_invoice_output.xlsx` / `..._discipline_invoice_output.xlsx`) so the two transforms no longer
overwrite each other's output by default. Verification: from-scratch rebuild reports `Control rows: 9`
(was 8), `B11` now holds the `EIB Population Type` dropdown (shifted down from `B10`), and the Runner
sheet's row 6 reads `Output File Path` with the AR default value seeded correctly.

**Near-miss during this fix, worth recording as a Gotcha (see below): `rm`-ing
`runner_shell/RunnerShell.xlsm` directly instead of using `--overwrite` destroyed the `arinvoice`/
`mapping` sheets again** — the exact failure mode from the "Recovered the 'arinvoice'/'mapping'
mapping tabs..." Done entry above, reintroduced by bypassing the preserved-sheets mechanism rather
than the mechanism itself failing. Caught immediately by checking `wb.sheetnames` right after
regenerating, before any further work compounded it. Recovered via
`git show aa20278:elt_transform_excel/runner_shell/RunnerShell.xlsm > /tmp/...` (the file is
git-tracked, not gitignored, and was committed at `aa20278`/eib30 right before this session), copied
back into place, then regenerated with `--overwrite` as normal — which correctly preserved
`arinvoice`/`mapping` this time. No data was actually lost since `aa20278` was only minutes old, but a
session that had gone longer before noticing would have lost more.

**Architecture change: split "build-time mapping shape" from "run-time lookup data" into two
different mechanisms, per direct user design feedback.** The user questioned why `arinvoice`/
`ardisciplineryinvoice` needed to live inside the deployed Runner workbook at all, given the only
thing ever read from them at runtime was a timestamp comparison — the actual column-mapping rows
were already baked into the `.bas` code and never re-read. That's correct, and led to a real
redesign:

- **Moved `arinvoice` and `ardisciplineryinvoice` out of `runner_shell/RunnerShell.xlsm` into a new,
  never-shipped file, `runner_shell/RunnerShellMapping.xlsx`.** These tabs describe each transform's
  column-to-column shape — a one-time code-generation spec, not data that changes independently of
  the code.
- **Removed the per-run drift check from `modArInvoiceCustomerInvoice.bas`** (deleted
  `ValidateMappingDefinitionTimestamp`, `ReadMappingDefinitionTimestamp`, and the
  `MAPPING_SHEET_NAME`/`MAPPING_TIMESTAMP_*`/`CODE_GENERATED_FROM_MAPPING_TIMESTAMP` constants, and
  the call site in `RunArInvoiceCustomerInvoice`) — there is nothing left in the deployed workbook for
  it to read.
- **Added an equivalent check at build time instead**, in `scripts/build_runner.py`:
  `_verify_mapping_freshness` reads each profile's new `transform.mapping_tab` /
  `transform.mapping_generated_at` YAML keys, opens `RunnerShellMapping.xlsx`, and compares the live
  tab's `Q1` against the expected value — failing the **build** (not the run) on mismatch. Verified
  by deliberately corrupting `arinvoice!Q1` to `2099-01-01 00:00:00` and confirming
  `build_runner.py` exits 1 with `Mapping drift detected for 'ar_invoice': ...`; restoring the correct
  value and rebuilding exits 0 again.
- **But the `mapping` tab (entity/source/target lookup rows, e.g. company-entity names) stays inside
  `RunnerShell.xlsm`/`dist/Runner.xlsm`**, on the user's explicit instruction: this is reference *data*
  that can change independently of any transform's code (e.g. a legal entity name changing), not a
  code-generation spec, so it belongs as a live runtime lookup, not a build-time check. Added
  `modHelpers.LookupMappingValue(runnerWorkbook, entityName, sourceValue)` (generic 3-column
  entity/source/target table scan) and rewired both transforms' Company resolution to call it instead
  of hardcoding entity names: AR Invoice's `ResolveCompanyId` `Select Case` became a
  `LookupMappingValue` call with the old Select Case's fallback (pass through the raw code) preserved
  for an unmatched entity; Discipline Invoice's hardcoded `OUTPUT_COMPANY_ID` constant became
  `SOURCE_COMPANY_CODE = "FA"` plus its own `ResolveCompanyId(runnerWorkbook)` wrapper. Both
  `WriteOutputRow` subs gained a `runnerWorkbook` parameter to pass the lookup target through.
  Verified: `mapping` tab content (`company`/`FA`/`THE FA GROUP - FA` etc.) unchanged and present in
  `dist/Runner.xlsm` (confirmed via raw `xl/workbook.xml` sheet-name inspection, not just one
  library's view — see the Gotcha below on why).
- **Real near-miss found while doing this:** an initial attempt to delete `arinvoice` /
  `ardisciplineryinvoice` / `mapping` together from `RunnerShell.xlsm` via openpyxl reported success
  (`Sheets now: ['Runner']`) and `pyopenvba`'s VBA-module listing showed no extra sheet — but the
  `mapping` sheet was still physically present in `xl/workbook.xml` on disk afterward. Root cause not
  fully diagnosed (suspected: an openpyxl `keep_vba=True` save edge case, or the sheet wasn't fully
  detached before save); see the Gotcha below. Recovered by restoring `RunnerShell.xlsm` from the
  last known-good git commit (`25b0eb1`, confirmed via raw XML to have exactly
  `['Runner', 'arinvoice', 'ardisciplineryinvoice', 'mapping']`) and redoing only the
  `arinvoice`/`ardisciplineryinvoice` deletion this time, re-verifying via raw XML (not openpyxl's
  `sheetnames` print or `pyopenvba`'s module list alone) that the result was exactly
  `['Runner', 'mapping']` before proceeding further.
- Added `transform.mapping_tab` / `transform.mapping_generated_at` to `ar_invoice.yaml`
  (`arinvoice` / `2026-06-19 13:55:00`) and `discipline_invoice.yaml` (`ardisciplineryinvoice` /
  `2026-06-22 23:30:00`, a fresh timestamp written into the tab specifically for this change, since
  `ardisciplineryinvoice` never had a timestamp pair before this).
- Updated `docs/ARCHITECTURE.md` (new `RunnerShellMapping.xlsx` artifact row, mapping-lookup-sheet
  bullet under Runner Shell components), `docs/RUNBOOK.md` (replaced "Mapping Timestamp Workflow"
  with a table distinguishing the two mechanisms plus separate workflow sections for each),
  `docs/CONFIGURATION.md` (new YAML keys, corrected profile descriptions), and `README.md`'s
  repository layout. Verification: `python3 -m py_compile scripts/build_runner.py
  scripts/create_runner_shell.py` exits 0; a from-scratch `build_runner.py` run exits 0 with the
  corrected timestamps in place; `dist/Runner.xlsm`'s `xl/workbook.xml` shows exactly
  `['Runner', 'mapping']`.

**Found and fixed a real bug, caught by the user inspecting the mapping/template before running:
`ClearPublishedRows` in both transform modules computed `lastRow` from column A alone
(`Cells(Rows.Count, 1).End(xlUp).Row`), which is too narrow for the real Target EIB template.**
Direct inspection of `Submit_Customer_Invoice_v46.1.xlsx` found column A's last populated row is row 5
(header/instruction rows only), but `KE`/`KF` (Revenue Category / Line Item Description) have sample
placeholder text ("AR Conversion" / "Open AR Conversion") through row 9 — Excel's actual `UsedRange`
extends to row 117. With `targetDataStartRow = 6`, the old calculation produced `lastRow = 6` (only
one row cleared), silently leaving rows 7-9's placeholder text in `KE`/`KF` for any transform run
that doesn't itself write to those cells. This was invisible in AR Invoice's output because that
transform writes fresh values into `KE`/`KF` on every row (`Course Name`/`Course Code`), overwriting
whatever was there regardless of whether the clear was correct. Discipline Invoice's mapping says
`KE`/`KF` should stay empty (`NULL` in `ardisciplineryinvoice`, confirmed earlier to mean "leave the
cell blank") — so it never writes them, which is exactly what exposed the bug: those rows would have
inherited the template's stale placeholder text. Fixed by switching `lastRow` to
`targetSheet.UsedRange.Row + targetSheet.UsedRange.Rows.Count - 1` in both
`modArInvoiceCustomerInvoice.bas` and `modDisciplineCustInvoice.bas` — `UsedRange`
reflects the sheet's true populated extent across every column, not just column A. Verification:
direct inspection of the template confirmed `UsedRange` extends to row 117 (`ws.dimensions` =
`A1:QA117`), comfortably covering the row 6-9 placeholder block; `python3 -m py_compile
scripts/build_runner.py scripts/create_runner_shell.py` exits 0; a from-scratch rebuild exits 0,
`Modules: 4`, sheets `['Runner', 'mapping']`. **Not yet re-verified in real Excel** — the next manual
run should confirm `KE`/`KF` come out blank in the Discipline Invoice output, not just that the build
succeeds.

## Session start prompt

```
load the use-context skill and continue in elt_transform_excel/docs/todo/archive/build-elt_transform_excel.md
```

## Session start checklist

- Read `elt_transform_excel/PRD.md`.
- Read `elt_transform_excel/docs/ARCHITECTURE.md`.
- Read this anchor doc end-to-end before making changes.
- Resume from `## Resume (start here)` and execute the next unfinished "Still todo" step only — that
  is currently the "Still todo (second vertical slice)" section, not the original six-item plan
  (which is fully done and kept only as history further down).
- The scaffold, shared VBA engine, shell generator, and Python-native build are complete; the `AR
  Invoice -> Customer Invoice` vertical slice is confirmed working end to end in real Excel; the
  `Discipline Invoice -> Customer Invoice` vertical slice is code-complete and awaiting its first
  manual-Excel verification run.

## Backlog Item: Build elt_transform_excel

A wholly-Excel, finance-user-run ETL solution that turns a source data extract into a populated
Workday EIB workbook, distinct from and with no dependency on `elt_ingest_excel`. Full intent and
contracts: `elt_transform_excel/PRD.md`. Full design: `elt_transform_excel/docs/ARCHITECTURE.md`.
Read both before continuing — this anchor doc tracks progress and constraints, it does not restate
the design.

### Done

- **Design settled and documented.** `PRD.md` and `docs/ARCHITECTURE.md` fully reflect the final
  architecture: a Runner Shell (one control sheet + bootstrap importer) → Interim
  workbook (mimics DB tables: `raw_*`/`ref_*`/`src_*`/`workday_*`) → Output workbook (cloned, populated
  Target EIB). Verification: both files read end-to-end this session; no stale references to ADO/SQL/
  Power Query remain outside the explicitly-labelled "Rejected alternatives" sections.
- **Architecture decision spike completed and recorded.** Built a throwaway `.xlsm` and tested two
  rejected mechanisms directly on Mac Excel before committing to native VBA — see Gotchas and PRD's
  "Rejected alternatives" section for the evidence. Spike files removed: `find ~/Documents /tmp -iname
  "pq_spike*"` returns nothing.
- **Config-purity split documented.** `docs/ARCHITECTURE.md` § "Config purity: where domain knowledge
  lives" records the decision to keep the shared engine (dispatcher/table-ops/helpers) domain-free and
  push column names/join keys/mapping rules into per-combination `vba_src/transforms/<combo>/` modules
  — the VBA analogue of `elt_ingest_excel`'s SQL files. The stronger alternative (declarative
  worksheet-based mapping tables + generic interpreter) was considered and explicitly rejected for now.
- **Implementation order clarified.** The build plan now makes the intended sequence explicit: scaffold
  the repo, author the shared VBA engine as text, create only a minimal Runner Shell stub, then prove
  the first end-to-end vertical slice for one `(Source Type, EIB Population Type)` before expanding the
  transform catalogue.
- **Reference sample files moved into build-time config.** The first concrete sample input set now
  lives in `config/transforms/ar_invoice.yaml`: source workbook, target EIB workbook, customer ID
  mapping workbook, and master mapping workbook/tab for the AR Invoice -> Customer Invoice scenario.
- **Cold-start prompt standardised.** The session-start handoff phrase is now fixed to: `load the
  use-context skill and continue in elt_transform_excel/docs/todo/archive/build-elt_transform_excel.md`.
- **Step-1 repo scaffold created.** Added the initial implementation skeleton:
  `elt_transform_excel/README.md`, `elt_transform_excel/docs/RUNBOOK.md`,
  `elt_transform_excel/docs/CONFIGURATION.md`, `elt_transform_excel/docs/ops/CLI.md`,
  `elt_transform_excel/vba_src/`, `elt_transform_excel/vba_src/bootstrap/`,
  `elt_transform_excel/vba_src/transforms/`, `elt_transform_excel/scripts/`, and
  `elt_transform_excel/runner_shell/`. Verification: `ls -R elt_transform_excel` shows the new doc
  and directory structure.
- **Step-2 shared VBA engine authored as text.** Added the first generic engine source set under
  `elt_transform_excel/vba_src/`: `modTableOps.bas` (`FilterRows`/`JoinRows`/`GroupByRows`/
  `LookupValue` plus table-array helpers), `modHelpers.bas` (control-sheet reads, workbook open/copy,
  Interim save/clone helpers, results writing), `modDispatcher.bas` (initially a
  selection-key-based dispatcher using an `Application.Run` naming convention, later superseded by the
  YAML-driven transform-registry binding recorded below, with no transform-specific branching),
  `clsKeyedCollection.cls` (native `Collection`-backed keyed lookup wrapper), and
  `bootstrap/BootstrapImporter.bas` (recursive `.bas`/`.cls` importer into the workbook VBA project,
  with visible status write-back). Also aligned the previously-missing `elt_transform_excel/scripts/`
  scaffold directory. Verification: `ls -R elt_transform_excel` shows the new engine files and
  `grep -R -n "ADODB\|Scripting\.Dictionary\|FileSystemObject\|VBScript\.RegExp" elt_transform_excel/vba_src`
  returns no matches.
- **Step-3 Runner build script authored.** Added `elt_transform_excel/scripts/build_runner.py`, a
  standalone macOS build CLI that copies `runner_shell/RunnerShell.xlsm` to `dist/Runner.xlsm`,
  and originally used AppleScript/Excel to run the bootstrap import. That first version was later
  superseded by the Python-native build step recorded below.
- **Step-4 Runner shell generation is now automated.** Added
  `elt_transform_excel/scripts/create_runner_shell.py`, which uses `pyOpenVBA` to create a fresh
  `runner_shell/RunnerShell.xlsm` with a real VBA project, injects
  `vba_src/bootstrap/BootstrapImporter.bas` from text source, reads build-time defaults from
  `config/transforms/ar_invoice.yaml`, and uses `openpyxl` to lay out the single visible `Runner`
  sheet with the 8 control labels, profile-driven default values, profile-driven dropdowns for
  `Source Type` / `EIB Population Type`, a run-action panel, and the initial results area.
  Verification:
  `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/create_runner_shell.py --overwrite`
  exits 0, and direct inspection confirms `runner_shell/RunnerShell.xlsm` contains `xl/vbaProject.bin`
  plus the expected `Runner` sheet labels.
- **Step-5 Python-native Runner build completed.** Reworked
  `elt_transform_excel/scripts/build_runner.py` to remove the live Excel / AppleScript dependency. The
  build now copies the generated shell, uses `pyOpenVBA` to delete previously imported non-document
  modules, re-adds every importable `.bas` / `.cls` file from `vba_src/`, writes the same
  `Bootstrap Import *` status cells directly with `openpyxl`, and verifies the resulting module set
  from the saved workbook itself. Also updated `README.md`, `docs/ARCHITECTURE.md`, `docs/RUNBOOK.md`,
  `docs/ops/CLI.md`, and `docs/TODO.md` to reflect the new build flow. Verification:
  `python3 -m py_compile elt_transform_excel/scripts/build_runner.py`,
  `uv run --with pyOpenVBA --with openpyxl --no-project python elt_transform_excel/scripts/build_runner.py --help`,
  and two consecutive runs of
  `uv run --with pyOpenVBA --with openpyxl --no-project python elt_transform_excel/scripts/build_runner.py`
  all exit 0; direct inspection of `dist/Runner.xlsm` shows the expected module set and status cells.
- **Step-6 Build-time transform config introduced.** Added
  `elt_transform_excel/config/transforms/ar_invoice.yaml` as the first YAML transform profile and
  reworked `elt_transform_excel/scripts/create_runner_shell.py` to read shell defaults and dropdowns from
  that profile instead of hardcoded sample paths. Updated `README.md`, `docs/ARCHITECTURE.md`,
  `docs/CONFIGURATION.md`, and `docs/ops/CLI.md` to document the new split between build-time YAML
  profiles and runtime control-sheet inputs. Verification:
  `python3 -m py_compile elt_transform_excel/scripts/create_runner_shell.py` and
  `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/create_runner_shell.py --overwrite`
  should both exit 0.
- **AR Invoice vertical slice implementation authored.** Added
  `elt_transform_excel/vba_src/transforms/modArInvoiceCustomerInvoice.bas` as the first real
  per-combination transform module. It registers the
  `AR Invoice -> Customer Invoice` step list, expands the Workday customer extract into
  `src_customer_lookup`, builds `src_ar_invoice_customer_invoice`, builds
  `workday_customer_invoice`, and publishes the result into a cloned
  `Submit Customer Invoice` EIB workbook. Verification:
  `uv run --with pyOpenVBA --with openpyxl --no-project python elt_transform_excel/scripts/build_runner.py`
  exits 0 and imports 5 modules into `dist/Runner.xlsm`.
- **Sample-data contract checked against the real workbooks.** Direct workbook inspection on the Mac
  dev machine confirmed the first source workbook has 124 publishable AR rows, 109 rows resolve to a
  Workday customer ID via the customer extract's pipe-delimited `Reference ID` field, and 10 distinct
  source customer IDs remain unresolved and therefore need output flagging rather than silent drop.
  Verification:
  `uv run --with openpyxl --no-project python -c "<sample-inspection script>"` prints
  `ROWS 124`, `RESOLVED 109`, `UNRESOLVED_DISTINCT 10`, and confirms the live workbook/template
  contract used by the AR vertical slice.
- **Transform macro binding is now config-driven.** Extended
  `elt_transform_excel/config/transforms/ar_invoice.yaml` with `runtime.step_list_macro`, reworked
  `elt_transform_excel/scripts/build_runner.py` to read every YAML profile under
  `config/transforms/`, and reworked `elt_transform_excel/vba_src/modDispatcher.bas` to resolve the
  step-list macro from a visible transform-registry block written into the built Runner workbook rather
  than from a hard-coded `GetStepList_<selectionKey>` naming convention alone. Updated
  `README.md`, `docs/ARCHITECTURE.md`, `docs/CONFIGURATION.md`, `docs/RUNBOOK.md`,
  `docs/ops/CLI.md`, and `docs/TODO.md` accordingly. Verification:
  `python3 -m py_compile elt_transform_excel/scripts/build_runner.py` and
  `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/build_runner.py`
  both exit 0; current build output reports `Bindings: 1`.
- **Excel automation blocker reproduced and narrowed.** Rebuilt
  `elt_transform_excel/dist/Runner.xlsm` successfully from the current text sources, confirmed
  `osascript -e 'tell application "Microsoft Excel" to get version"'` still returns `16.110`, and
  confirmed `open -a "Microsoft Excel" ".../dist/Runner.xlsm"` still exits 0. AppleScript `open`
  and `run VB macro` calls against Excel, plus a control `Finder open POSIX file ...` AppleScript,
  all fail from this CLI host with `A privilege violation occurred` (`-10004`). That narrows the
  remaining blocker to Apple Event automation privileges in this host, not to the built workbook
  artifact itself.
- **Priority changed for the next session.** The next work item is now to make the
  `AR Invoice -> Customer Invoice` transform itself correct first, using
  `elt_transform_excel/vba_src/transforms/modArInvoiceCustomerInvoice.bas` and the sample contract in
  `config/transforms/ar_invoice.yaml`, before resuming the blocked manual Excel end-to-end run.
- **AR Invoice live mapping has now been coded into the VBA transform module.**
  Updated `elt_transform_excel/vba_src/transforms/modArInvoiceCustomerInvoice.bas` so the current
  `AR Invoice -> Customer Invoice` transform follows the live mapping authored in
  `runner_shell/RunnerShell.xlsm` tabs `arinvoice` and `mapping`, which are the source of truth for
  this live AR scenario; the module should be treated as code generated from and aligned to that live
  workbook mapping rather than as an independent restatement of the mapping rules. Verification:
  `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/build_runner.py`
  exits 0 and rebuilds `dist/Runner.xlsm`.
- **AR runtime metadata and drift protection are now in place.** The AR build/runtime path now uses
  profile/config metadata for non-mapping runtime settings and keeps an explicit mapping-definition
  timestamp contract so the live `arinvoice` tab and the code derived from it can be checked for drift
  before the transform runs. Verification:
  `python3 -m py_compile elt_transform_excel/scripts/build_runner.py` and
  `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/build_runner.py`
  both exit 0.
- **Keyed VBA state updates no longer fail on existing keys.** Fixed
  `elt_transform_excel/vba_src/clsKeyedCollection.cls` so updates/removals use the tracked numeric
  collection index rather than passing a string key to `Collection.Remove`, which was causing the real
  Excel run to fail with `Invalid procedure call or argument` as soon as the transform tried to update
  existing run-context values such as status/counters/output paths. Verification:
  `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/build_runner.py`
  exits 0 and rebuilds `dist/Runner.xlsm` with the corrected class module.
- **A reusable AppleScript macro runner now exists for user-approved hosts.** Added
  `elt_transform_excel/scripts/run_runner_macro.applescript` plus `docs/ops/CLI.md` usage so the built
  `dist/Runner.xlsm` can be opened and `modDispatcher.RunSelectedTransform` can be invoked from a normal
  macOS Terminal or Script Editor session that has Excel automation permission. This does not change the
  current CLI-host blocker; it provides a reusable script for less-manual local verification. Verification:
  `osacompile -o /tmp/run_runner_macro.scpt elt_transform_excel/scripts/run_runner_macro.applescript`
  should exit 0.
- **Root-caused the original "Compile error: Invalid procedure call or argument" on
  `CollectImportableFiles`.** Reproduced it from a clean Excel session via Debug → Compile VBA Project
  alone (zero code executed), which ruled out a real source-code defect and a stale break-mode artifact.
  Traced it to `scripts/build_runner.py`'s prior approach of injecting/removing VBA modules directly into
  the binary `.xlsm` via `pyOpenVBA`'s `project.add_module`/`delete_module` — confirmed structurally
  corrupt by Excel's own save routine, which looped on "Errors were detected while saving" without ever
  completing a repair. `create_runner_shell.py`'s one-time `add_module` call for `BootstrapImporter` on a
  **freshly created** workbook is unaffected and remains in place — only repeated injection into an
  existing project was unsafe.
- **Reworked `scripts/build_runner.py` to stop writing the VBA project binary directly.** Removed the
  `ModuleSpec`/pyOpenVBA-injection path (`_sync_vba_modules`) entirely. First attempt: drive the
  workbook's own `BootstrapImporter` inside a live Excel host via AppleScript (the same mechanism a
  developer uses manually) instead. Verification: `python3 -m py_compile scripts/build_runner.py` exits 0.
- **AppleScript cannot reliably pass arguments to a Mac Excel macro.** Every syntax tried for
  `run VB macro ... Arg1:value` (`Arg1:`, `arg1:`, `«class Arg1»:`, `«class ARG1»:`, a trailing list
  literal) raised `Parameter error (-50)`. Worked around by adding
  `BootstrapImporter.BootstrapImportFromConfiguredSource()` (new, zero-argument) to
  `vba_src/bootstrap/BootstrapImporter.bas` — it reads the source folder from a labelled
  `Bootstrap Source Folder` cell (written via `openpyxl` beforehand) and calls the existing
  `BootstrapImportFromFolder` internally.
- **Full headless automation of the module-import step proved unreliable even with zero arguments**,
  across several independent failure modes, all on this one Mac Excel build: an Apple Event privilege
  violation (`-10004`, fixed via `tccutil reset AppleEvents com.apple.Terminal` + re-granting Terminal's
  Automation access to Microsoft Excel — see Gotchas), a macOS per-file "Grant Access" sandbox prompt
  with nobody present to click it, an `AppleEvent timed out (-1712)` while a real VBA import error sat on
  a modal dialog, and a repeatable `"The variable targetWorkbook is not defined" (-2753)` AppleScript
  error at the post-macro cleanup step. The common thread: whenever the macro itself doesn't complete
  cleanly inside the automated loop, Excel's Apple Event connection to that document becomes unstable for
  every subsequent command, including our own cleanup code — not just the macro call itself.
- **Confirmed by hand that the import mechanism itself is sound.** In the VBE Immediate window,
  `ThisWorkbook.VBProject.VBComponents.Import "<path>"` succeeded for every module — including
  `clsKeyedCollection.cls` and the large `modArInvoiceCustomerInvoice.bas` — when run interactively by a
  human (after granting the one-time per-file "Grant Access" prompt). Only the *scripted, headless,
  multi-module* loop is unreliable; the underlying VBA mechanism is not at fault.
- **Decision (user-approved): abandoned full CLI automation of the module-import step.**
  `scripts/build_runner.py` (default, no flag) now does only the parts that are reliably scriptable in pure
  Python: copy the shell to `dist/Runner.xlsm`, write the YAML-driven transform registry, write the
  `Bootstrap Source Folder` cell — then prints instructions for the one remaining manual step and exits.
  Added `scripts/build_runner.py --verify` to check an already-built workbook afterwards (reads the
  Bootstrap Import status cells, the transform registry, and the imported module set). Verification:
  `python3 -m py_compile scripts/build_runner.py` exits 0; a prep run exits 0 and prints the manual-step
  instructions; `dist/Runner.xlsm` after prep contains only `BootstrapImporter` plus the written cells
  (no other modules yet — that's expected until the manual step runs).
- **Updated `README.md`, `docs/ARCHITECTURE.md`, `docs/RUNBOOK.md`, and `docs/ops/CLI.md` to match the
  hybrid build.** All four previously described the now-reverted fully-automated build (e.g.
  `ARCHITECTURE.md` said modules are "synchronised...via pyOpenVBA"; `RUNBOOK.md` said the script
  "depends on those Python packages rather than a live Excel host" — both false after this session's
  revert). Rewrote the build/deployment sections in all four to describe: prep via `build_runner.py`,
  the required manual Excel step, and `build_runner.py --verify`; added the Full Disk Access/Automation
  permission prerequisites to `RUNBOOK.md`. Verification: `grep -rn "pyOpenVBA\|live Excel\|AppleScript"
  README.md docs/ARCHITECTURE.md docs/RUNBOOK.md docs/ops/CLI.md` shows only statements consistent with
  the current hybrid build, none claiming full automation.
- **Fixed a real bug found while re-verifying this anchor doc's own claims: `runner_shell/RunnerShell.xlsm`
  was never actually minimal.** `pyOpenVBA`'s `ExcelFile.create_new()` seeds every fresh VBA project with
  a default empty `Module1` placeholder that `create_runner_shell.py` never removed — every shell
  generated this session (and likely before) carried it silently. This also means an earlier diagnosis
  this session ("`Module1` appearing means an import died partway through") was wrong — `Module1` was
  baseline noise present from the very start, unrelated to import success or failure (corrected in
  Gotchas). Fixed `scripts/create_runner_shell.py` to delete `Module1` immediately after `create_new()`,
  before adding `BootstrapImporter`. Verification: regenerated both `runner_shell/RunnerShell.xlsm` and a
  fresh `dist/Runner.xlsm` (prep mode); inspecting both via
  `python -c "import pyopenvba as p; print(p.ExcelFile('<path>').module_names())"` shows exactly
  `['ThisWorkbook', 'Sheet1', 'BootstrapImporter']` in each, no `Module1`.
- **The manual module-import step (Still todo item 1) is now confirmed working end to end, after fixing
  four distinct real bugs found by actually running it in Excel rather than just inspecting code:**
  1. **All `vba_src/*.bas`/`*.cls` source files were LF-only; Mac Excel's `VBComponents.Import` needs
     CRLF to recognise the legacy `VERSION 1.0 CLASS`/`BEGIN`/`END` class-module header.** With LF only,
     `clsKeyedCollection.cls`'s header was mis-parsed as literal code, producing a real-Excel
     "Compile error: Expected: end of statement" on the `VERSION 1.0 CLASS` line. Confirmed by comparing
     against `xlwings`'s known-good `.bas`/`.cls` exports (CRLF throughout) and reproducing the same
     working result after converting all six `vba_src` files to CRLF.
  2. **`BootstrapImporter.bas`'s `FolderExists` called `Dir$` from inside an active `Dir$` enumeration
     loop in `CollectImportableFiles`.** `Dir$` has exactly one global search state and is not
     re-entrant; the nested call corrupted the outer folder-enumeration loop, which silently dropped the
     `transforms/` subfolder, so `modArInvoiceCustomerInvoice.bas` was never discovered or imported (the
     macro still reported success with 4 modules instead of 5, no error). Fixed by switching
     `FolderExists` to `GetAttr`, which does not touch `Dir$`'s state.
  3. **Fixing `BootstrapImporter.bas`'s text source alone did not fix the running macro**, because
     `BootstrapImporter` deliberately skips re-importing itself during a bootstrap run (see existing
     Gotcha), and `scripts/build_runner.py`'s prep step only copies the existing
     `runner_shell/RunnerShell.xlsm` rather than regenerating it — so the stale, pre-fix
     `BootstrapImporter` kept running until `scripts/create_runner_shell.py --overwrite` was re-run to bake
     the fixed source into a fresh shell.
  4. **`scripts/build_runner.py --verify`'s own module-presence check had a false-failure bug.** Its
     "preserve existing module" heuristic matched on the literal string `Attribute VB_Base =` in a
     module's saved source, intending to detect document modules (`ThisWorkbook`/`Sheet1`) — but Excel
     stamps that same attribute onto every class module it re-serializes, document or not, so
     `clsKeyedCollection` (a plain class module) was wrongly flagged as "missing" after a fully correct
     import. Fixed `_preserve_existing_module` to match by explicit module name
     (`bootstrapimporter`/`thisworkbook`/`sheet1`) instead of sniffing saved source text.
  Verification: with all four fixes applied, a from-scratch rebuild (`dist/` deleted, shell regenerated,
  `build_runner.py` prep re-run, manual `BootstrapImporter.BootstrapImportFromConfiguredSource` run by the
  user in Excel) followed by
  `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python scripts/build_runner.py --verify`
  prints `Verification passed.` / `Imported modules: 5`, confirming all five non-bootstrap modules
  (`clsKeyedCollection`, `modDispatcher`, `modHelpers`, `modTableOps`, `modArInvoiceCustomerInvoice`) are
  present and the status/registry cells match expectations.
- **Recovered the "arinvoice"/"mapping" mapping tabs that a prior session silently destroyed, and fixed
  the script defect that destroyed them.** Running `modDispatcher.RunSelectedTransform` against the
  rebuilt `dist/Runner.xlsm` raised "Subscript out of range" on
  `runnerWorkbook.Worksheets(MAPPING_SHEET_NAME)` — `ThisWorkbook` had no `arinvoice` sheet. Root cause
  (confirmed via `git show <rev>:./runner_shell/RunnerShell.xlsm` across history): the `arinvoice` and
  `mapping` worksheets existed in `runner_shell/RunnerShell.xlsm` as of commit `b3a0ba1` (eib15,
  2026-06-19 14:06, `arinvoice!Q1` = `2026-06-19 13:55:00` — exactly matching
  `CODE_GENERATED_FROM_MAPPING_TIMESTAMP` in the VBA code) but were gone by commit `1a3e4c0` (eib19, the
  same day, 11:58 — file size dropped from 18225 to 11795 bytes in that commit). Cause:
  `scripts/create_runner_shell.py` always calls `pyOpenVBA`'s `ExcelFile.create_new()`, which replaces the
  whole file from scratch and only adds the generated `Runner` sheet — any hand-authored sheet present in
  a previous version is silently destroyed by `--overwrite`. This happened in a prior session (running
  `--overwrite` for an unrelated `Module1` cleanup fix), not in this one. Fix, user-approved: (1) restored
  the exact `arinvoice`/`mapping` sheet content from commit `b3a0ba1` into the current
  `runner_shell/RunnerShell.xlsm` (cell values only — no formulas/merges in either sheet, confirmed by
  inspection); (2) added `_read_preserved_sheets`/`_write_preserved_sheets` to
  `scripts/create_runner_shell.py` so any non-`Runner` sheet present in the existing output file is read
  before `create_new()` runs and written back after the generated `Runner` sheet/VBA project are rebuilt —
  `--overwrite` now refreshes the generated parts without destroying hand-authored ones. Verification: ran
  `create_runner_shell.py --overwrite` again immediately after the fix and confirmed
  `runner_shell/RunnerShell.xlsm` still has `sheets: ['Runner', 'arinvoice', 'mapping']` with
  `arinvoice!Q1` unchanged at `2026-06-19 13:55:00`.
- **Fixed a second, unrelated stale-config bug found by the user while testing: the AR Invoice source
  data default pointed at a nonexistent file.** `config/transforms/ar_invoice.yaml`'s
  `shell_defaults.source_data` named `InvoiceDetailFALL.xlsx`, but the real, current source extract on
  disk is `InvoiceDetailFALL.csv` (confirmed: the `.xlsx` does not exist at that path; only the `.csv`
  does). Fixed the YAML to point at the `.csv`. No VBA change was needed: `CopyWorkbookDataToSheet`
  already opens the source via plain `Workbooks.Open` (which Excel handles transparently for `.csv`) and,
  for the `raw_source` sheet specifically, falls back to "first visible worksheet" rather than a
  hardcoded sheet name — so the source file's actual tab name (the user separately confirmed it's `"in"`
  in their copy) does not matter.
- **Fixed a real VBA compile error found on the next real-Excel compile attempt: `Chr$(30)` is not a
  valid `Const` initializer.** `vba_src/modTableOps.bas` declared
  `Private Const KEY_SEPARATOR As String = Chr$(30)` — VBA requires `Const` initializers to be
  literal/compile-time expressions, and a function call like `Chr$(30)` does not qualify, regardless of
  the function being intrinsic. Real Excel raised "Compile error: Constant expression required" on that
  line. Fixed by replacing the `Const` with a one-line `Private Function KeySeparator() As String`
  returning `Chr$(30)`, and updating its one call site (`BuildCompositeKeyFromRow`) to call
  `KeySeparator()`. Verification: `grep -n "KEY_SEPARATOR\|KeySeparator" vba_src/modTableOps.bas` shows no
  remaining `Const` declaration and one call site using the function; rebuilt
  `runner_shell/RunnerShell.xlsm` (`--overwrite`) and `dist/Runner.xlsm` afterward — both still preserve
  `arinvoice`/`mapping` and the corrected `.csv` source path, confirming the preserved-sheets fix holds
  across repeated rebuilds.
- **Fixed a real Mac-VBA-specific runtime error found on the next real-Excel run attempt: `Set x =
  obj.Item(key)` against a parameterized `Property Get` raised error 450.** `modHelpers.ReportIssue`'s
  `Set issues = runContext.Item("Issues")` failed with "Run-time error '450': Wrong number of arguments
  or invalid property assignment," highlighted on the `Set` line. This was the first place in the actual
  run where a parameterized property's return value was captured via `Set` to get an object reference
  back (`clsKeyedCollection.Item` returns a `Collection`); every earlier `.Item(...)` call in the run only
  read primitives via plain assignment (e.g. `CStr(runContext.Item(...))`), and those all worked, which is
  why this surfaced only once `ReportIssue` first ran. Root cause: Mac Excel's VBA implementation is
  unreliable specifically for `Set x = obj.ParameterizedProperty(arg)` returning an object, even though
  the call syntax is otherwise valid VBA. Fixed by converting `clsKeyedCollection.Item` from
  `Public Property Get Item(ByVal key As String) As Variant` to a plain
  `Public Function Item(ByVal key As String) As Variant` — identical call syntax, but functions don't go
  through VBA's property-dispatch path. Confirmed `Item` was the only parameterized property anywhere in
  `vba_src` (`grep -n "Property Get.*(ByVal" vba_src/**/*.bas vba_src/**/*.cls` matches only the
  no-parameter `Count` property, which is unaffected since it returns `Long`, not an object). Verification:
  rebuilt `runner_shell/RunnerShell.xlsm` (`--overwrite`) and `dist/Runner.xlsm`; both still preserve
  `arinvoice`/`mapping` and the corrected `.csv` source path.
- **Fixed the actual error 450 cause (the prior fix attempt — `Item` as `Function` instead of
  `Property Get` — was not it; the real bug needed an explicit `IsObject`/`Set` branch).** Plain
  `Item = mValues.Item(...)` (Let-assignment) does not reliably preserve an object reference into a
  Variant-returning accessor on Mac Excel's VBA — confirmed by the error persisting unchanged after the
  Property→Function change, then resolving once `Item` and `TryGetValue` were rewritten to check
  `IsObject(foundValue)` and use `Set` explicitly for object values, falling back to plain `=` for
  primitives. Both call sites that retrieve `ThisWorkbook`-scoped state (`ReportIssue`'s
  `Set issues = runContext.Item("Issues")` and `TryGetValue`'s internal `Item` call) depend on this.
- **Fixed a real, separate logic bug in `clsKeyedCollection.AddOrUpdate` found via a genuine run-time
  457 ("This key is already associated with an element of this collection") while testing repeated
  unresolved-customer counting (`AddOrIncrement` in the AR transform).** `mValues`, `mDisplayKeys`, and
  `mKeyOrder` are meant to stay positionally in sync — `IndexOfKey` finds a position in `mKeyOrder` and
  uses that same position to index into the other two. The update branch (key already exists) removed
  the old entry from `mValues`/`mDisplayKeys` at that position but re-added the new entry at the *end* of
  those two collections, while leaving `mKeyOrder`'s entry for that key at its original, now-stale
  position — silently breaking the positional-sync invariant after the very first update of any
  non-last key. A later update could then remove the wrong element, eventually leaving a stale duplicate
  entry under the same key, which raised 457 on a subsequent `Add`. `Remove` already handled this
  correctly (removes from all three collections at the same index); fixed `AddOrUpdate` to match —
  remove from all three (if present) then add to all three, always together. Side effect: an updated
  key now moves to the end of iteration order (via `Keys()`); confirmed nothing in `vba_src` depends on
  update-stable key ordering.
- **Fixed a real VBA compile error found on the next real-Excel compile attempt: `parentFolder` (a local
  variable) and `ParentFolder` (a module function) are the same identifier to VBA, since VBA is
  case-insensitive.** `modHelpers.bas`'s `BuildSiblingFilePath` declared `Dim parentFolder As String`
  then wrote `parentFolder = ParentFolder(filePath)` — the local variable shadowed the function of the
  same name (regardless of case), so the right-hand side resolved against the local scalar instead of
  calling the function, raising "Compile error: Expected array" (VBA tried to index the local `String`
  as if it were an array). Fixed by renaming the local variable to `parentFolderPath`, which does not
  collide with any function name. Verification: `grep -n "parentFolder" vba_src/modHelpers.bas` shows
  the local variable as `parentFolderPath` and the function call as `ParentFolder(filePath)`; rebuilt
  `dist/Runner.xlsm` (`build_runner.py`, no flag) — exits 0, reports `Modules: 3`, `Bindings: 1`.
- **Manual import confirmed working against the rewritten 3-module set.** After the `modHelpers.bas`
  fix above, rebuilt `dist/Runner.xlsm` clean from scratch (deleted the file, reran
  `scripts/build_runner.py` with no flag — confirmed only `ThisWorkbook`/`Sheet1`/`BootstrapImporter`
  present beforehand), then ran `BootstrapImporter.BootstrapImportFromConfiguredSource` interactively
  in real Excel. It imported and saved successfully. Direct inspection
  (`pyopenvba.ExcelFile(...).vba_project().modules`) confirmed `modDispatcher`, `modHelpers`,
  `modArInvoiceCustomerInvoice` are present alongside the preserved `arinvoice`/`mapping` sheets.
- **Fixed a false-failure bug in `build_runner.py --verify`'s module-presence check, found by running
  `--verify` against the now-correctly-imported workbook above.** `--verify` reported `unexpected:
  sheet2, sheet3` even though the import was fully correct. Root cause:
  `_preserve_existing_module`'s hardcoded `_ALWAYS_PRESERVED_MODULE_NAMES` set only excluded the
  literal name `sheet1` as a document module — but the preserved-sheets feature (see "Recovered the
  'arinvoice'/'mapping' mapping tabs..." Done entry, added after that hardcoded list was written)
  means the workbook now also has `Sheet2`/`Sheet3` document-module codenames for the `arinvoice`/
  `mapping` sheets, which were never added to the list. (`module.kind` was considered as a more
  principled fix but rejected: pyOpenVBA's `VBAModuleKind` only distinguishes "standard" from "other"
  — "other" lumps document modules together with class modules, which would wrongly hide a future
  real `.cls` module too.) Fixed `_preserve_existing_module` in `scripts/build_runner.py` to match any
  `Sheet<N>` codename via a regex instead of a fixed name list — this generalizes correctly as more
  sheets are preserved, without losing the ability to flag a missing/unexpected real code module.
  Verification: `python3 -m py_compile scripts/build_runner.py` exits 0;
  `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python scripts/build_runner.py --verify`
  now prints `Verification passed.` / `Imported modules: 3` against the same on-disk `dist/Runner.xlsm`
  that previously failed verification.
- **Architecture simplification: removed the Interim workbook, the generic table-operations module, and
  `clsKeyedCollection` entirely, after stepping back to question why a single lookup join needed this
  much machinery.** Prompted by direct user pushback: the actual AR Invoice transform is one lookup
  join (source invoice rows + a customer extract whose `Reference ID` already is the join key), not a
  multi-stage migration needing dedup/rank/normalize the way `elt_ingest_excel`'s DuckDB/SQL pipeline
  does for full Workday CCS migration — and every real bug found and fixed this session (CRLF parsing,
  `Dir$` re-entrancy, the `Set`/`IsObject` Mac VBA quirk, the `AddOrUpdate` position-sync bug) was in
  that generic engine, none in the actual mapping logic. Deleted `vba_src/clsKeyedCollection.cls` and
  `vba_src/modTableOps.bas` entirely. Rewrote `vba_src/transforms/modArInvoiceCustomerInvoice.bas` as a
  single `Public Sub RunArInvoiceCustomerInvoice(runnerWorkbook, controlSheet)` that opens the customer
  extract and source data directly, builds a plain native `Collection` lookup (source customer ID ->
  Workday ID) plus two small Collections for unresolved-customer counting (one for counts looked up by
  key, one an append-only list of distinct keys, since `Collection` has no key-enumeration API — this
  sidesteps the old three-parallel-Collection position-sync design entirely, not just patches it), and
  writes each row straight into the cloned output EIB — no Interim workbook. Simplified
  `vba_src/modDispatcher.bas` to resolve one macro name from the registry and call it directly (dropped
  step-list arrays, the per-selection `Start Row` override, and the `clsKeyedCollection`-based run
  context). Trimmed `vba_src/modHelpers.bas` to only the plumbing still used (control-sheet read/write,
  template cloning, bulk worksheet-to-array read, path helpers); removed everything Interim-specific
  (`CreateRunContext`, `CopyWorkbookDataToSheet`, `WriteWorksheetTable`, `ReportIssue`,
  `WriteResultSummary`, `SaveInterimWorkbook`, `IncrementResultCounter`). Updated
  `config/transforms/ar_invoice.yaml`'s `runtime.step_list_macro` to name `RunArInvoiceCustomerInvoice`
  directly (the YAML key name is unchanged to keep the Python tooling diff minimal, but it now names a
  macro to call, not a step-list-returning function). Updated `scripts/create_runner_shell.py`'s
  `RESULT_ROWS` to drop the "Interim Workbook" result cell. Updated `docs/ARCHITECTURE.md`, `PRD.md`,
  `README.md`, `docs/RUNBOOK.md`, `docs/ops/CLI.md`, and `docs/CONFIGURATION.md` to describe the
  two-artifact (Runner Shell + Output workbook) design; `docs/ARCHITECTURE.md` keeps a "Superseded:
  Interim workbook..." section recording why the original design was rejected, per the project's
  anti-drift/historical-record convention. Verification: `grep -R -n
  "ADODB\|Scripting\.Dictionary\|FileSystemObject\|VBScript\.RegExp" vba_src` returns nothing;
  `python3 -m py_compile scripts/create_runner_shell.py scripts/build_runner.py` exits 0; a from-scratch
  rebuild (`create_runner_shell.py --overwrite` then `build_runner.py`) exits 0, reports `Modules: 3`
  (down from 5) and `Result rows: 7` (down from 8), and the rebuilt `dist/Runner.xlsm` still has
  `arinvoice`/`mapping` restored and the registry row `ar_invoice__customer_invoice -> AR Invoice ->
  Customer Invoice -> RunArInvoiceCustomerInvoice -> 6`. **Not yet verified in real Excel** — the
  manual bootstrap-import + run-the-transform steps below need to be redone against this rewritten
  module set before this can be marked confirmed-working.
- **The `AR Invoice -> Customer Invoice` transform is now confirmed working end to end in real
  Excel, against the real (non-test) customer extract.** First reverted Still todo item 0: the
  trimmed `Extract_Customers_20260611_test100.xlsx` debugging override was removed from
  `config/transforms/ar_invoice.yaml`'s `customer_spreadsheet`, restored to
  `Extract_Customers_20260611.xlsx`, and both `runner_shell/RunnerShell.xlsm` and `dist/Runner.xlsm`
  were regenerated (confirmed the preserved `arinvoice`/`mapping` sheets and mapping timestamp
  `2026-06-19 13:55:00` survived the rebuild). The user then ran the manual bootstrap import and
  `modDispatcher.RunSelectedTransform` via Tools > Macro in real Excel. Verified two ways: (1) the
  output workbook `Submit_Customer_Invoice_v46.1_output.xlsx` has real per-row data across all 124
  data rows in `Company`, `Currency`, `Customer`, `Invoice Date`, `Amount Due` — not placeholder
  rows; (2) a screenshot of the live Runner sheet (before any save) showed `Run Status: Completed`,
  `Rows Processed: 124`, `Rows Skipped: 3`, and a populated `Issues` cell listing the 3 skipped
  source rows plus per-customer UNMAPPED counts in column `CQ` (e.g. `l00030 (2), m01010 (1),
  m00658 (1), l00040 (2), l00050 (2), ...`), `Output Workbook` and `Selection Key
  ar_invoice__customer_invoice` both correct.
- **Found and fixed a real bug while verifying the run above: the Runner workbook's result cells
  were never persisted to disk.** `RunArInvoiceCustomerInvoice` (both the success and `HandleFailure`
  paths) wrote `Run Status`/`Rows Processed`/`Rows Skipped`/`Issues`/`Output Workbook`/`Selection
  Key`/`Last Run At` into `controlSheet` but never called `runnerWorkbook.Save` — so the values were
  only visible live in the open Excel session and `dist/Runner.xlsm` on disk kept showing stale
  "Not run yet" / 0 rows from the last bootstrap-import save. Confirmed by comparing the on-disk
  `Run Status` (still "Not run yet") against the live-session screenshot (showing "Completed") taken
  at the same time. Fixed `vba_src/transforms/modArInvoiceCustomerInvoice.bas` to call
  `runnerWorkbook.Save` after writing the result cells on both the success path and the failure
  path (captured `Err.Number`/`Err.Description` into local variables first on the failure path, since
  the intervening `On Error Resume Next` / save / `On Error GoTo 0` could otherwise disturb `Err`
  before re-raising). Verification: `file vba_src/transforms/modArInvoiceCustomerInvoice.bas`
  confirms CRLF line endings preserved after the edit (required for Mac Excel's
  `VBComponents.Import` to parse class/module headers correctly — see Gotchas);
  `python3 -m py_compile scripts/create_runner_shell.py scripts/build_runner.py` exits 0; both
  `create_runner_shell.py --overwrite` and `build_runner.py` (no flag) were rerun and exit 0,
  reporting `Modules: 3`, `Bindings: 1`. **Not yet re-verified in real Excel** — the next
  manual bootstrap-import + run needs to confirm the result cells now persist to `dist/Runner.xlsm`
  after a normal macro run (no explicit user save needed).

### Original six-item plan (historical — all done; superseded by the second-vertical-slice section above)

0. ~~Revert the temporary `customer_spreadsheet` test override before any real run.~~ **Done** —
   `config/transforms/ar_invoice.yaml` now points at the real `Extract_Customers_20260611.xlsx`;
   both `runner_shell/RunnerShell.xlsm` and `dist/Runner.xlsm` rebuilt and confirmed to still
   preserve `arinvoice`/`mapping`.
1. ~~Run the one remaining manual step in Excel...~~ **Done** — re-confirmed against the rewritten
   3-module set (see "Manual import confirmed working against the rewritten module set" Done entry):
   the manual `BootstrapImporter.BootstrapImportFromConfiguredSource` step imported and saved
   correctly.
2. ~~Verify the build with `build_runner.py --verify`...~~ **Done** — current `dist/Runner.xlsm` passes
   verification with `Imported modules: 3` (after fixing the `--verify` false-failure bug — see the
   matching Done entry above).
3. ~~Run the corrected `AR Invoice -> Customer Invoice` transform macro itself in real Excel.~~
   **Done** — confirmed end to end against the real customer extract: output workbook has real
   per-row data across all 124 rows, and the live Runner sheet showed `Run Status: Completed`,
   `Rows Processed: 124`, `Rows Skipped: 3`, populated `Issues` (see the matching Done entry above).
   A follow-on bug was found and fixed in the same check: the Runner workbook's result cells were
   never saved to disk (see "Found and fixed a real bug..." Done entry) — that fix is **not yet
   re-verified in real Excel**.
4. ~~Re-verify the run with the result-persistence fix applied.~~ **Done** — user ran the manual
   bootstrap import (`Bootstrap Import At` 2026-06-22 20:00:19, 3 modules) followed by
   `modDispatcher.RunSelectedTransform` (`Last Run At` 2026-06-22 20:01:16) in real Excel. Reading
   `dist/Runner.xlsm` directly from disk afterward (not a live-session screenshot) via openpyxl
   confirmed the result cells persisted: `Run Status: Completed`, `Rows Processed: 124`,
   `Rows Skipped: 3`, populated `Issues` text, correct `Output Workbook` and `Selection Key`. The
   `runnerWorkbook.Save` fix is confirmed working.
5. ~~Document the mapping timestamp workflow~~ **Done** — added a "Mapping Timestamp Workflow"
   section to `docs/RUNBOOK.md` describing the `arinvoice` tab's `Mapping Definition Timestamp`
   (`P1`/`Q1`) vs. `modArInvoiceCustomerInvoice.bas`'s `CODE_GENERATED_FROM_MAPPING_TIMESTAMP` drift
   check, and the required refresh/rebuild/re-verify sequence whenever the live mapping changes.
   **Superseded:** that per-run VBA check (and `CODE_GENERATED_FROM_MAPPING_TIMESTAMP` itself) was
   deleted in the "Architecture change" Done entry below — the check is now build-time, in
   `build_runner.py`, against `transform.mapping_generated_at` in YAML. `docs/RUNBOOK.md` has been
   rewritten accordingly; do not expect `CODE_GENERATED_FROM_MAPPING_TIMESTAMP` to exist in
   `vba_src/` any more.
6. ~~Fill out remaining docs with real content.~~ **Done** — `docs/ops/CLI.md` and `docs/RUNBOOK.md`
   already fully documented the `--verify` flag, the manual bootstrap-import step, and the Full Disk
   Access / Automation permission prerequisites from earlier sessions. The one real gap was
   `docs/CONFIGURATION.md`, which only documented `runtime.step_list_macro`; added a full table
   covering every key actually read by `create_runner_shell.py`/`build_runner.py`
   (`transform.key/label/source_type/eib_population_type/start_row/description`,
   `shell_defaults.*`), including a note that `transform.start_row` is vestigial
   (written into the registry but not read by the current AR transform macro since the
   step-list/Interim-workbook design was dropped). Verification: `grep -n "profile\[\|profile\.get"
   scripts/create_runner_shell.py scripts/build_runner.py` confirms the documented keys match every key
   actually read in code. **Superseded:** later cleanup removed the duplicated dropdown metadata from
   transform profiles entirely (see the "Found and fixed a real duplication bug" entry in the
   "Discipline Invoice mapping" section near the top of this doc) — `docs/CONFIGURATION.md` has since
   been updated to match.
- **Fixed a real AR Invoice date-normalization bug that the user found in the live EIB output.**
  The `arinvoice` tab in `runner_shell/RunnerShellMapping.xlsx` correctly declared
  `Order Date` as `dd.mm.yy` with target `YYYY-MM-DD`, but the deployed Runner does **not** read that
  workbook at runtime; it only uses the generated VBA in `vba_src/`. `NormalizeSourceDate` in both
  `vba_src/transforms/modArInvoiceCustomerInvoice.bas` and
  `vba_src/transforms/modDisciplineCustInvoice.bas` previously fell back too early to
  locale-sensitive `IsDate`/`CDate`, so a value shaped like `16.05.26` could still end up rendered in
  the output as Excel serial-0 / `1899-12-30` despite the mapping tab being correct. Fixed by parsing
  delimited DMY text explicitly before the VBA date fallback in both modules. Verification: user
  reran the AR Invoice flow after refreshing the Runner's imported modules and confirmed the fix
  worked ("ok that worked").

### Accumulated Active Constraints (active for all remaining items)

- **Native VBA only.** No `ADODB`, no `Scripting.Dictionary`/`FileSystemObject`/`RegExp`, no Power
  Query automation, anywhere in `vba_src/`. Confirmed by direct testing, not assumption — see Gotchas.
- **Superseded: two artifacts, not three.** The Interim workbook was removed (see "Architecture
  simplification" Done entry) — there is no longer a middle artifact to keep separate. Current
  constraint: **Runner Shell** (orchestration only — one sheet + bootstrap importer only) and **Output
  workbook** (clone of Target EIB, populated directly). The Runner never holds run data.
- **No hidden sheets**, in any workbook, ever.
- **Never delete `runner_shell/RunnerShell.xlsm` with `rm` (or any direct filesystem delete) to force
  a rebuild.** Always regenerate via `create_runner_shell.py --overwrite`, which reads and reattaches
  any hand-authored sheet (e.g. `arinvoice`/`mapping`) before replacing the file. A direct `rm` skips
  that preserved-sheets mechanism and destroys those sheets — this happened twice now (see the
  "Recovered the 'arinvoice'/'mapping' mapping tabs..." Done entry and the "Near-miss during this fix"
  Done entry); the second time was caught immediately and recovered from git history, but do not rely
  on git history being available or recent enough to save the next occurrence.
- **Every required control-sheet field needs an explicit `EnsureNonBlank`/`EnsureFilePath` check in
  the transform module that reads it.** `Output File Path` is required with no fallback (unlike the
  old auto-derived-from-`Target EIB` behavior it replaced) — a transform module that reads it without
  validating non-blank would silently pass an empty string to `CloneTemplateWorkbook`, which falls
  back to the old auto-derived suffix instead of failing loudly.
- **The Runner Shell is generated, not hand-edited.** Regenerate
  `runner_shell/RunnerShell.xlsm` with `scripts/create_runner_shell.py`; do not patch the binary in the
  Excel editor. Every VBA module, including `BootstrapImporter`, remains text-sourced under `vba_src/`.
- **Build-time scenario defaults live in YAML.** Sample paths, default `Source Type`, and default
  `EIB Population Type` belong under `config/transforms/*.yaml`, not as hardcoded literals in Python.
- **Transform-to-macro registration is now YAML-driven at build time.** The shipped Runner reads a
  visible transform-registry block written from `config/transforms/*.yaml`; reassigning a supported
  selection to a different step-list macro should require a YAML edit plus rebuild, not a dispatcher
  code edit.
- **Domain knowledge lives only in the per-combination transform modules under `vba_src/transforms/`.**
  The shared engine (dispatcher, helpers) must never contain a column name, sheet name, or business
  value specific to one transform. (No table-ops module any more — each transform builds whatever
  native `Collection`(s) its own lookup needs; see "Architecture simplification" Done entry.)
- **The first AR Invoice publish shape is now concrete.** The `AR Invoice -> Customer Invoice`
  vertical slice publishes into the target EIB sheet `Submit Customer Invoice` and is now implemented
  from the live `runner_shell/RunnerShell.xlsm` `arinvoice` mapping rather than
  from the historical `elt_ingest_excel` migration SQL.
- **The live `arinvoice` and `ardisciplineryinvoice` tabs in `runner_shell/RunnerShellMapping.xlsx`
  (not `RunnerShell.xlsm` — moved out, see the "Architecture change" Done entry) are the single
  source of truth for each transform's column-mapping shape.** The anchor doc must not restate
  field-level mapping rules; treat `modArInvoiceCustomerInvoice.bas` /
  `modDisciplineCustInvoice.bas` as code developed from those tabs. Both now have their
  content checked at **build time** by `build_runner.py` via an MD5 hash of the tab compared against
  `transform.mapping_content_hash` in the owning YAML profile — neither tab is read by the deployed
  `dist/Runner.xlsm` at runtime any more. **Superseded:** there is no longer a `Mapping Definition
  Timestamp` (`P1`/`Q1`) cell in either tab — a manually-bumped timestamp could silently fail to catch
  drift if someone edited the tab without remembering to bump it; the content hash is computed
  mechanically from cell values, so any edit is caught regardless.
- **`RunnerShellMapping.xlsx` format masks are build-time design inputs, not runtime enforcement.**
  A correct `source_format`/`target_format` row in `arinvoice` or `ardisciplineryinvoice` does not by
  itself make the deployed Runner parse or render dates correctly, because that workbook is never read
  by `dist/Runner.xlsm` at runtime. Any transform that normalizes dates or numbers must implement that
  parsing explicitly in VBA and must not rely on `IsDate`/`CDate` alone for non-locale formats such as
  `dd.mm.yy`.
- **A literal `NULL` string in either mapping tab's `default` column means "leave the target cell
  empty," not "write the text NULL."** Confirmed directly by the user for `ardisciplineryinvoice`'s
  `ke`/`kf` rows — the code correctly already never writes those cells, so no VBA change was needed,
  but don't write the literal string when a future profile's tab uses this convention.
- **Mapping change workflow: update the tab, then the code, then `transform.mapping_content_hash` in
  YAML (via `build_runner.py --print-mapping-hashes`), then rebuild — in that order, every time.** A
  mismatch between the live tab's content hash and the YAML value fails the *build* (not a run in
  Excel) with "Mapping drift detected"; this is the only thing left enforcing that AR/Discipline
  mapping changes don't silently diverge from the code.
- **The separate `mapping` tab (entity/source/target lookup rows, e.g. company-entity names) is a
  different artifact with a different rule: it stays inside `RunnerShell.xlsm`/`dist/Runner.xlsm`
  and is read live at runtime** via `modHelpers.LookupMappingValue`, with no timestamp/drift-check —
  by design, since it holds reference data that can change independently of any transform's code. Do
  not move it into `RunnerShellMapping.xlsx` alongside `arinvoice`/`ardisciplineryinvoice` — that
  would require a rebuild for every reference-data edit, which is exactly what this split was meant
  to avoid.
- **Build/deploy is Mac-only, by the project owner.** No Windows build pipeline, no CI. The Mac-built
  Runner is shipped to finance users as-is.
- **No shared code or config with `elt_ingest_excel`.** The two projects are intentionally independent.
- **Core implementation docs now exist and should be extended in place.** `README.md`,
  `docs/RUNBOOK.md`, `docs/CONFIGURATION.md`, and `docs/ops/CLI.md` are now the canonical locations
  for orientation, operating guidance, configuration rules, and commands; future work should deepen
  these files rather than creating parallel notes.
- **Module import into the deployed Runner workbook must happen via a live Excel host running
  `VBComponents.Import`, never via pyOpenVBA binary injection.** pyOpenVBA's `add_module`/`delete_module`
  directly on an existing `.xlsm` produced a workbook Excel itself could not repair-save — confirmed
  structurally corrupt, not just a stale-cache issue. `create_runner_shell.py`'s single one-time
  `add_module` call on a **freshly created** workbook (for `BootstrapImporter` only) is the one proven-safe
  exception — never extend pyOpenVBA writes beyond that single call.
- **The module-import step is a manual, human-run action in Excel, not a CLI-automated one.**
  `scripts/build_runner.py` (no flag) prepares everything it reliably can — shell copy, transform registry,
  `Bootstrap Source Folder` cell — and stops; the developer runs
  `BootstrapImporter.BootstrapImportFromConfiguredSource` themselves via Tools>Macro, then runs
  `build_runner.py --verify` to confirm. Do not re-attempt full headless automation of this specific step
  without new evidence the underlying Mac Excel Apple-Event instability is resolved.
- **AppleScript cannot reliably pass arguments to a Mac Excel VB macro via `run VB macro`.** Every
  parameter syntax tried (`Arg1:`, lowercase, raw four-char code, positional list) failed with
  "Parameter error (-50)". Any future macro that needs a runtime input from a script must receive it via
  a pre-written cell value, never a macro argument.
- **The confirmed-unreliable automation failures (privilege violation, parameter error, AppleEvent
  timeout, "variable not defined") all occurred specifically around macros that manipulate the VBA
  project itself (`VBComponents.Import`/`Remove`).** That is a real, narrower pattern than "all Excel
  automation is unreliable." But do not over-generalize the other direction either:
  `modDispatcher.RunSelectedTransform` has **not** been confirmed reliable headlessly — the one headless
  call to it this session returned no Apple Event error, but `Run Status` never updated and the workbook
  was later found corrupt, so that result is inconclusive, not a confirmation. Test it fresh (ideally via
  Tools>Macro first, where errors are visible) before relying on headless automation for it.
- **Microsoft Excel needs macOS Full Disk Access** (System Settings → Privacy & Security → Full Disk
  Access) for automated `.bas`/`.cls` file reads to not block on a per-file "Grant Access" sandbox
  prompt. This is separate from the Automation/Apple-Events permission (System Settings → Privacy &
  Security → Automation) needed for Terminal to control Excel at all — both are required, neither
  substitutes for the other.
- **Control-sheet dropdown options must never be hand-copied into more than one transform profile
  YAML.** `Source Type` is derived automatically by `create_runner_shell.py` from every profile's
  `transform.source_type` across `config/transforms/*.yaml` (mirroring how `build_runner.py` already
  scans all profiles for the registry). Adding a new transform profile should never require editing an
  existing profile's YAML to add it to a dropdown list — if a change appears to require that, the
  dropdown logic has regressed back to the per-profile duplication this fix removed (see the "Found
  and fixed a real duplication bug" Done entry above).
- **A mapping-tab row's `mapping_fail` column overrides its `default` column on a failed lookup, not
  the other way round.** `mapping_fail = reject` (the only value used so far, on each tab's `cq`/
  Customer ID row) means an unresolved lookup excludes the source row from output entirely — counted
  in `Rows Skipped`, never written with a placeholder value (no more hardcoded `UNMAPPED`). Any row
  with no `mapping_fail` value falls back to its own `default` column (column D) on a failed lookup —
  baked into the transform module as a `*_DEFAULT_FALLBACK` constant — never to the raw, unmapped
  source value. Any new mapping-tab row that needs lookup-failure handling must declare it explicitly
  via one of these two columns; do not invent a third hardcoded fallback in VBA.
- **`create_runner_shell.py`'s `DEFAULT_PROFILE_CONFIG` tracks whichever transform the business is
  currently testing, not a fixed "first" profile.** Currently `discipline_invoice.yaml` (was
  `ar_invoice.yaml`). Whoever switches it must regenerate `runner_shell/RunnerShell.xlsm`
  (`create_runner_shell.py --overwrite`) and rebuild `dist/Runner.xlsm` (`build_runner.py`) so the
  seeded control-sheet defaults actually change — editing the YAML alone does nothing until both
  build steps re-run.
- **Each profile's `shell_defaults.output_file_path` lives in the same folder as its
  `shell_defaults.source_data`** (both currently `.../finance_ref/eib/ar/`), not next to `target_eib`
  — moved there on user request so a run doesn't require navigating between folders to find the
  source file and check the output. The two profiles still use distinct filenames
  (`*_ar_invoice_output.xlsx` / `*_discipline_invoice_output.xlsx`) even though they now share a
  folder, preserving the no-overwrite-collision invariant from "Output-collision limitation fixed."
- **Any output column written as a literal-format date string (anywhere `target_format` in a mapping
  tab is a string like `YYYY-MM-DD`) must have its cell forced to Text (`NumberFormat = "@"`)
  immediately before the value is assigned.** Without this, Excel auto-detects the date-shaped string
  on write and silently converts it to a real date, which then renders using whatever number format
  the Target EIB template's cell already had (`General` on the `GF`/Invoice Date column, which made
  Excel apply its own default short-date format) — completely overriding the mapping tab's declared
  `target_format`. Fixed for `GF` in both transform modules; apply the same pattern to any future
  literal-format date/text column before assuming `target_format` is honored just because the string
  was built correctly.
- **Locale-sensitive VBA date parsing can still bypass a correct mapping tab.** Even when
  `RunnerShellMapping.xlsx` correctly says `dd.mm.yy -> YYYY-MM-DD`, the deployed Runner will still
  mis-handle values like `16.05.26` if the transform module relies on `IsDate`/`CDate` before doing
  explicit delimiter-based parsing. The fix was to parse dotted/slashed DMY text first in
  `NormalizeSourceDate`; keep that order in any future transform that handles non-US date text.

### Verification

- `cat elt_transform_excel/PRD.md elt_transform_excel/docs/ARCHITECTURE.md | grep -in "ADODB\|ACE OLEDB\|Power Query\|sql_src"` — should only match inside the explicitly-labelled "Rejected alternatives" / historical-decision text, never inside a current-design section.
- `find ~/Documents /tmp -iname "pq_spike*" 2>/dev/null` — should return nothing (spike cleanup confirmed this session).
- `ls -R elt_transform_excel` — should show `README.md`, `docs/RUNBOOK.md`, `docs/CONFIGURATION.md`,
  `docs/ops/CLI.md`, `vba_src/`, `build/`, and `runner_shell/`; after step 2 it should also show
  `vba_src/modHelpers.bas`, `vba_src/modDispatcher.bas`, `vba_src/bootstrap/BootstrapImporter.bas`,
  `vba_src/transforms/modArInvoiceCustomerInvoice.bas`, and (as of the second vertical slice)
  `vba_src/transforms/modDisciplineCustInvoice.bas`,
  `config/transforms/discipline_invoice.yaml`. **Superseded as of
  the architecture simplification:** `vba_src/modTableOps.bas` and `vba_src/clsKeyedCollection.cls` no
  longer exist — do not expect them.
- `grep -R -n "ADODB\|Scripting\.Dictionary\|FileSystemObject\|VBScript\.RegExp" elt_transform_excel/vba_src` — should return nothing.
- `python3 -m py_compile elt_transform_excel/scripts/create_runner_shell.py` — should exit 0.
- `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/create_runner_shell.py --overwrite` — should exit 0 and create `runner_shell/RunnerShell.xlsm`.
- `python3 -m py_compile elt_transform_excel/scripts/build_runner.py` — should exit 0.
- `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/build_runner.py --help` — should exit 0 and
  print CLI usage.
- **Superseded — do not expect this anymore:** `scripts/build_runner.py` (no flag) no longer imports any
  transform modules itself, so it no longer "rebuilds `dist/Runner.xlsm`" with a populated module set or
  produces an "idempotent import" in one command. It now only preps the workbook (shell copy, registry,
  source-folder cell) and reports the YAML-driven binding count — see the `--verify` entries below for how
  the *fully built* workbook (after the manual step) is actually confirmed.
- `uv run --with openpyxl --no-project python -c "<sample-inspection script>"` — should confirm the
  concrete source/customer/target workbook contract used by the first vertical slice (current observed
  output: `ROWS 124`, `RESOLVED 109`, `UNRESOLVED_DISTINCT 10`, and the live workbook/template contract
  used by the AR vertical slice).
- `osascript -e 'tell application "Microsoft Excel" to get version"'` — currently returns `16.110`.
- `open -a "Microsoft Excel" "/Users/rpatel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/dist/Runner.xlsm"` — currently exits 0 and opens the built Runner workbook.
- `osascript -e 'tell application "Microsoft Excel" to open POSIX file "/Users/rpatel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/dist/Runner.xlsm"'` — the original `A privilege violation occurred` (`-10004`) here is now resolved (see Gotchas: `tccutil reset` + re-granting Terminal's Automation access to Excel); this now succeeds.
- `osascript -e 'tell application "Microsoft Excel" to run VB macro "Runner.xlsm!modDispatcher.RunSelectedTransform"'` — same Apple Event fix applies (no privilege-violation error), but headless invocation of this specific macro is **still unverified** — the one time this was tried headlessly (an earlier session) returned no error, but `Run Status` on the Runner sheet never updated from "Not run yet," and that workbook was later found to be structurally corrupt (see Gotchas). The macro itself is now confirmed correct via the **interactive** Tools>Macro path (see "AR Invoice -> Customer Invoice transform is now confirmed working..." Done entry) — this headless-call question is a separate, lower-priority open item, not part of Still todo item 4. "No AppleScript error" is not evidence the macro ran; always check the actual result cells (`Run Status`, `Rows Processed`, etc.) or the Project pane after any headless call, not just the exit code.
- `osascript -e 'tell application "Finder" to open POSIX file "/Users/rpatel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/dist/Runner.xlsm"'` — was the same Apple Event restriction as above; same fix applies.
- `python3 -m py_compile scripts/build_runner.py` — should exit 0 (rewritten this session to drop pyOpenVBA module injection).
- `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python scripts/build_runner.py` — prep-only mode: copies the shell to `dist/Runner.xlsm`, writes the transform registry and `Bootstrap Source Folder` cell, exits 0, and prints the manual bootstrap-import instructions. Does **not** import any transform modules by itself — that is the manual step.
- After completing the manual step in Excel (Tools>Macro, `BootstrapImporter.BootstrapImportFromConfiguredSource`):
  `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python scripts/build_runner.py --verify`
  — should exit 0 and report the Bootstrap Import status timestamp, confirming the status cells, transform
  registry, and module set all match expectations.
- Manual AR Invoice validation after the date-normalization fix: refresh the imported VBA modules in
  Excel, rerun `modDispatcher.RunSelectedTransform`, then confirm a source `Order Date` like
  `16.05.26` lands in output column `GF` as literal `2026-05-16` rather than `1899-12-30`. This was
  confirmed by the user in this session.
- In the VBE Immediate window (manual, interactive only — confirms the import mechanism itself is sound):
  `ThisWorkbook.VBProject.VBComponents.Import "<absolute path to any .bas/.cls under vba_src>"` — should
  succeed (after granting a one-time per-file "Grant Access" prompt if macOS shows one), for every module
  including `modArInvoiceCustomerInvoice.bas`. (`clsKeyedCollection.cls` no longer exists — see the
  "Architecture simplification" Done entry.)
- `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/build_runner.py`
  fails the build (does not reach `shutil.copy2`) if any profile's `transform.mapping_content_hash`
  doesn't match a live MD5 hash of the corresponding tab in `runner_shell/RunnerShellMapping.xlsx` —
  confirm by deliberately editing a value in `arinvoice` without updating
  `ar_invoice.yaml`'s `mapping_content_hash` and observing `Error: Mapping drift detected for
  'ar_invoice': ...`, then reverting the edit and confirming the build exits 0 again. Get the current
  hash for any tab with `build_runner.py --print-mapping-hashes`. (Superseded the old
  `Mapping Definition Timestamp`/`Q1` mechanism — see "Accumulated Active Constraints" above.)
- `python3 -c "import zipfile, re; print(re.findall(r'<sheet [^>]*name=\"([^\"]+)\"', zipfile.ZipFile('dist/Runner.xlsm').read('xl/workbook.xml').decode()))"`
  — should print exactly `['Runner', 'mapping']`. Use this raw-XML check, not `openpyxl.sheetnames`
  or `pyopenvba`'s module list alone, after any sheet add/delete (see Gotchas).

### Gotchas

- **An `openpyxl` `del wb[name]` + `wb.save(..., keep_vba=True)` can report success
  (`wb.sheetnames` shows the sheet gone, in-memory) while the sheet is still physically present in
  the saved file's `xl/workbook.xml` on disk.** Hit this deleting `arinvoice`/`ardisciplineryinvoice`/
  `mapping` together from `RunnerShell.xlsm` — the print statement right after `.save()` showed
  `['Runner']`, but a later raw-XML check found `mapping` still there. Root cause not fully
  diagnosed. `pyopenvba`'s VBA-module listing didn't catch it either, for an unrelated reason: it only
  enumerates worksheets that have a registered VBA document-module component in `vbaProject.bin`, and
  apparently doesn't reliably reflect every worksheet `openpyxl` operations touch — so it under- or
  over-reports compared to ground truth without warning.
- **After any `openpyxl` sheet add/delete on an `.xlsm`, verify via raw `zipfile` + a regex over
  `xl/workbook.xml`'s `<sheet name="...">` entries — not `wb.sheetnames` from a fresh `load_workbook`
  call alone, and not `pyopenvba`'s VBA-module list alone.** Both can disagree with the actual
  on-disk XML; the zip-level check is ground truth. This cost real time this session (re-deriving the
  correct file from git history) and should be the default verification step for any future sheet
  add/delete, not an occasional sanity check.
- **VBA identifiers are case-insensitive — a local variable can silently shadow a module-level function
  of the same name in any case combination.** `parentFolder` (local `String`) and `ParentFolder`
  (a `Function`) collided in `modHelpers.BuildSiblingFilePath`, producing "Compile error: Expected
  array" because the assignment resolved against the local variable, not the function. Watch for this
  whenever a helper function name is a common, lowercase-friendly noun (`parentFolder`, `fileName`,
  etc.) that a caller might also want as a local variable name.
- **A headless `osascript run VB macro` call returning with no error is not proof the macro executed
  correctly — check the actual result cells or Project pane every time.** The one headless call to
  `modDispatcher.RunSelectedTransform` made this session returned exit 0 with no Apple Event error, but
  `Run Status` on the Runner sheet never updated from "Not run yet," and the workbook turned out to be
  structurally corrupt at the time. "No error" can mean "didn't run at all" on this Mac Excel build, not
  just "ran successfully."
- **A `.xlsm`'s VBA project cannot be edited from outside a running Excel/VBA host.** Confirmed:
  `oletools`/`olefile` (already used elsewhere in this repo for VBA patching, see
  `elt_ingest_excel/src/elt_ingest_excel/macro/vba_patcher.py`) can only flip an existing flag in a
  module that's already there — it cannot inject a brand-new module from text. This is *why* the
  bootstrap-importer mechanism exists at all, and why the old "empty `.xlsm` plus later patching"
  approach was replaced with a fresh shell generated via `pyOpenVBA`.
- **`pyOpenVBA` can create a fresh `.xlsm` with a valid VBA project, but it cannot retrofit an
  arbitrary workbook that has no `xl/vbaProject.bin`.** An empty workbook created earlier in this
  effort looked like an `.xlsm` by extension only; zip inspection confirmed it had no VBA project, and
  `pyOpenVBA` refused to open it for module edits. `scripts/create_runner_shell.py` therefore creates a
  fresh shell from scratch instead of trying to upgrade an arbitrary workbook in place.
- **Mac Excel appears sandboxed for ad hoc file I/O.** Direct VBA `Open "..." For Output` to arbitrary
  new paths — even under `~/Documents`, even when an existing file already there was previously opened
  via a save dialog — failed silently with no file ever created. Prefer writing results to workbook
  cells (read back via VBA or relayed by the user) over `Open`/`Print #` to a fresh path; only paths a
  user has explicitly granted via an Open/Save dialog appear reliably writable.
- **`BootstrapImporter` should not try to re-import itself while it is running.** The text-sourced
  `vba_src/bootstrap/BootstrapImporter.bas` is still the source of truth, but the bootstrap macro
  already present in the Runner Shell must stay in sync with that file because removing/importing the
  currently executing module is unsafe. The current implementation therefore skips importing the
  `BootstrapImporter` component during a bootstrap run.
- **Plain `uv run python ...` is currently not usable from this repo root for this script.** In the
  current workspace state it tries to build the repo-level `ai-agent-core` source dependency and fails
  before the script starts (`Multiple top-level packages discovered in a flat-layout`). Use
  `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python ...` for
  `elt_transform_excel/scripts/create_runner_shell.py` and
  `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python ...` for
  `elt_transform_excel/scripts/build_runner.py` until that workspace packaging issue is fixed.
- **`ADODB`/ACE OLEDB does not run on Mac Excel VBA at all** — not "runs differently," fails outright.
  Confirmed by direct test before this constraint was written into the architecture.
- **Power Query's engine works on Mac** (manual UI "Get Data → From File" loaded a CSV correctly) **but
  its VBA automation does not**: `Workbook.Queries.Add` plus `ListObjects.Add` with the standard
  `OLEDB;Provider=Microsoft.Mashup.OleDb.1;...` connection string got stuck at "Getting Data..." with
  Refresh permanently disabled. That connection string is Windows-specific plumbing. Do not re-attempt
  Power Query automation on Mac without new evidence that this specific gap has closed.
- **This CLI environment can launch apps but cannot send Apple Events to them reliably.** Current
  evidence: `osascript -e 'tell application "Microsoft Excel" to get version"'` returns `16.110`, and
  `open -a "Microsoft Excel" ".../dist/Runner.xlsm"` exits 0, but AppleScript `open` and `run VB macro`
  attempts against Excel and even `Finder open POSIX file ...` fail with `A privilege violation
  occurred` (`-10004`). Treat this as an automation-host blocker, not as proof the VBA logic itself is
  wrong.
- **VBA `Collection.Remove` only accepts an index or the original key string, not a re-derived one.**
  Passing a normalized string key to `Collection.Remove` in the now-deleted `clsKeyedCollection` raised
  `Invalid procedure call or argument`. Still a relevant lesson for any future VBA `Collection` use, even
  though that specific class is gone (see "Architecture simplification" Done entry) — the current AR
  transform module avoids the whole problem by keeping plain `Collection`s keyed directly by their own
  normalized string (no separate positional tracking to keep in sync).
- **The new AppleScript runner is only useful from a host Excel is allowed to trust.** It wraps the same
  Excel Apple Event calls in a reusable file, but it will still fail in this constrained CLI host if the
  host itself lacks permission to control Excel.
- **A VBA "Compile error" dialog with a runtime-style message (e.g. "Invalid procedure call or argument")
  appearing from Debug → Compile VBA Project alone, with zero code executed, is a strong signal of
  underlying binary corruption** — not a real source-code defect. Confirmed here by reproducing it from a
  clean Excel session (fresh quit/reopen, nothing run) and root-causing it to pyOpenVBA's binary module
  injection in the old `build_runner.py`.
- **Mac Excel's own "repair" dialog looping without ever completing a save is a stronger corruption
  signal than a compile error alone.** If Excel can't even repair-save a file, stop "fixing" VBA code
  inside it — rebuild the artifact from a clean source instead.
- **Correction to an earlier diagnosis this session: `Module1`'s presence was misread as evidence of a
  failed import — it wasn't.** `pyOpenVBA`'s `ExcelFile.create_new()` seeds every fresh VBA project with
  a default empty `Module1` placeholder; it was present in `runner_shell/RunnerShell.xlsm` straight out of
  `create_runner_shell.py`, before any import loop ever ran. `scripts/create_runner_shell.py` now explicitly
  deletes it right after creation (see the `Done` entry below) — confirmed by inspecting
  `runner_shell/RunnerShell.xlsm`'s module list immediately after generation, both before and after that
  fix. **The actual signal of a failed/partial import is a missing *expected* module** (e.g.
  `modArInvoiceCustomerInvoice` absent from the Project pane), not the presence of `Module1` — always
  check against the expected module list, not just for stray names.
- **macOS's per-file sandbox "Grant Access" prompt for Excel is distinct from both the Automation
  (Apple Events) permission and Full Disk Access, and can still surface even after both of those are
  granted** — observed once when manually importing `modArInvoiceCustomerInvoice.bas` via the Immediate
  window after Full Disk Access was already granted. Granting Full Disk Access to Excel is the most
  reliable way to suppress it across an entire directory tree, but a one-off prompt for a brand-new path
  is still possible.
- **`tccutil reset AppleEvents com.apple.Terminal` plus triggering a fresh permission request (e.g.
  `osascript -e 'tell application "Microsoft Excel" to activate'`) is the reliable way to re-arm a
  stuck/denied Automation permission entry**, used when System Settings showed no toggle at all for the
  Excel↔Terminal pair (compare against a working pair like TRAE↔Excel, which did show toggles).
- **A macOS permission grant (Full Disk Access, Automation) does not always take effect for an
  already-running app instance.** Fully quitting and relaunching Excel (⌘Q, not just closing the
  workbook) was necessary at least once for a new grant to take effect.
