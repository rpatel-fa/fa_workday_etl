# Anchor Doc — Pre-run Validation Gate (elt_transform_excel)

## Resume (start here)

- **Status: NOT STARTED — captured 2026-07-14. No code written yet.** Trae is mid a mapping refactor
  (adding `BuildBusinessObjectReferenceLookup` / `FinalCodeSegment` / `ResolveLookupAcrossFinalCodeSegments`
  to `modFinanceCommon.bas`, plus a new `CONTROL_LABEL_MAPPING_SHEET` and likely new supplier-PO / GL
  transforms). **Do not start this until that refactor lands** — the validator must cover whatever new
  transforms/inputs it introduces, so building against a moving target wastes the work.
- **Next item:** Item 1 — design `ValidateRunPreconditions` (requirements source + where it lives).
- **Read before touching code:** `docs/ARCHITECTURE.md` (Modules — the domain-free split
  `modHelpers` / `modSourceTable` / `modFinanceCommon`) and `docs/CONFIGURATION.md` (control-sheet
  inputs + `source_layout.required_headers`). The per-transform required columns already live in each
  transform's `requiredHeaders` array and in `config/transform_specs/<key>.yaml`'s `source_layout`.

## Goal

Add a single upfront validation gate that runs **before** `BeginRun` and **before** the output EIB is
cloned, aggregates **all** problems into one friendly message, and fails without mutating run state.

**Why this is worth doing:** all the checks below already happen today, but *late, one-at-a-time, and
with poor messages* — after `modHelpers.BeginRun` has flipped Run Status to "Running" and reset the
reject log, and (for column checks) after the `*_output.xlsx` clone already exists. The dispatcher's
`modDispatcher.TryRunTransformMacro` also swallows the specific error into a generic
`1703 "Unable to run configured macro"` and re-runs the whole macro a second time on failure. A gate
that fails first, before side effects, with every problem listed, removes all of that.

## Scope — the four checks

1. **Files exist** — for every file field this transform needs. Two cases: *populated but missing*, and
   *required but blank*. Which fields are required is transform-specific (AR/discipline → customer
   extract; AP → supplier extract; `Output File Path` always; `Target EIB` always). Today only
   `outputFilePath` blankness is checked early (`EnsureNonBlank`); the rest surface only on open via
   `modHelpers.EnsureFilePath`.
2. **Tabs exist** — across **every** input workbook, not just the primary source: the named source
   sheet, the Target EIB target sheet, and (trains) the `Workday Mappings` cost-centre tab. Today a
   renamed tab throws raw Excel `error 9 "Subscript out of range"` mid-run from
   `modHelpers.ResolveSourceWorksheet` / `Worksheets(name)`.
3. **Mapping/lookup columns exist** — across **every** input workbook: the source columns each
   transform maps, plus the extract key columns (`Reference ID` + `Customer ID`/`Supplier ID`) and
   (trains) `OLD CODE` / `WORKDAY REF ID`. Today these raise `modSourceTable.FindColumnIndexInRow`
   `1611 "Header not found"` / `ValidateHeaderRow 1613` — after the clone.
4. **(Optional) Output file not locked/open** — the clone does `Kill` then `FileCopy`; if last month's
   `*_output.xlsx` is still open in Excel, that dies mid-run. Only precondition that can't be fully
   proven without trying; "exists and is locked" is detectable. Lower priority than 1–3.

## Design constraint (config-purity)

The checks are inherently per-transform. Do **not** hard-code file/tab/column lists in a shared
validator. Each transform must *declare* its requirements (reuse `source_layout.required_headers` from
the spec YAML + the transform's own sheet-name consts + which control-sheet file fields it consumes),
and a shared aggregating validator walks them. **Open decision (Item 1):** does the validator live in a
new `modPreflight.bas`, or in `modHelpers` — and how does a transform hand it its requirement set
(a declared function per module vs. reading the spec/registry). Decide this first; it shapes everything
else.

## Done

- (none yet)

## Still todo (next actions)

1. **Design `ValidateRunPreconditions`** — pick where it lives and how each transform declares its
   required files / tabs / columns (see Open decision above). Record the choice here before coding.
2. **Implement the aggregator** — collect every failure across checks 1–3 into one `Collection`, raise
   a single message listing all of them; return without side effects on failure.
3. **Wire it into the dispatcher** — call the gate in `modDispatcher.RunConfiguredTransform` *before*
   dispatching to the transform macro (so it runs before that macro's `BeginRun`). Confirm a validation
   failure does **not** trigger the double-run / generic `1703`.
4. **Per-transform requirement declarations** — add for all current transforms (`ar_invoice`,
   `ar_discipline_invoice`, `ap_amex_hotel`, `ap_amex_trains`) plus any new ones from Trae's refactor.
5. **(Optional) Output-lock check (check 4).**
6. **Docs** — update `ARCHITECTURE.md` (new module/flow) and `CONFIGURATION.md`; mark items here Done.

## Accumulated Active Constraints (active for all remaining items)

- **Shared-layer split (established by the common-layer refactor):** generic table ops →
  `modSourceTable`; workbook/control-sheet/run-log plumbing → `modHelpers`; reusable finance
  mappings/derivations → `modFinanceCommon`; per-profile column-to-target mapping stays in the
  transform module. A preflight validator is generic mechanism → shared layer, never per-transform code.
- **Config-purity:** no file/tab/column domain strings hard-coded in the shared validator; they come
  from per-transform declarations / spec YAML.
- **VBA namespace:** shared helpers are `Public` in exactly one standard module — a second `Public`
  with the same name in another module is an "Ambiguous name" compile error. New shared names must be
  unique across `modHelpers` / `modSourceTable` / `modFinanceCommon` / `modDispatcher`.
- **Fail before side effects:** the gate must run before `BeginRun` (status/reject-log reset) and before
  `CloneTemplateWorkbook`. That is the whole point; a check that runs after either is not this work.
- **Build discovery:** any new `.bas` under `vba_src/` is picked up by `build_runner.py` (`rglob`) and
  the recursive bootstrap import; component-name stems must stay unique.

## Verification

- No headless VBA compile in this environment — verification is manual, in real Excel:
  - `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/build_runner.py --verify`
    (structural: modules present). Run from the workspace root — see `docs/ops/CLI.md`.
  - Manual Excel import of `vba_src/` via BootstrapImporter, then Tools>Macro `RunSelectedTransform`.
- Behavioural acceptance (run each against a real transform):
  - Missing file / blank required file field → **single** aggregated message, Run Status still not
    "Running", and **no** `*_output.xlsx` created.
  - Renamed source/target/cost-centre tab → named in the aggregated message (not raw "error 9").
  - Missing mapped column (in source AND in an extract) → named in the aggregated message, before clone.
  - Happy path (all valid) → runs to completion unchanged; row counts match a pre-gate baseline run.

## Gotchas

- **Trae is editing `modFinanceCommon.bas` and the control-sheet labels concurrently.** Rebase on the
  landed mapping refactor before designing requirement declarations — new transforms and a new
  `Mapping Sheet` control field will exist that this gate must cover.
- **Dispatcher double-run on failure:** `TryRunTransformMacro` tries `wb!macro` then bare `macro`; a
  genuine failure runs the macro body twice. Putting the gate ahead of dispatch avoids this for the
  validation error class — verify the gate raises *before* either `Application.Run`.
- **`ResolveSourceWorksheet` with a blank name returns the first visible sheet** (no error); only a
  *named* missing sheet throws. Check 2 must validate the specific named sheets, not just "a sheet exists".
