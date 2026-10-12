# Architecture

`elt_transform_excel` is a self-contained Excel solution for finance users. There is no Python,
DuckDB, or external process anywhere in the **runtime** the finance user touches — everything below
runs inside Excel via plain VBA, so it runs on a Windows laptop with no install step. Python is used
only in **build tooling that finance users never run** (see Deployment), to assemble the deployed
Runner from text source — that does not affect the runtime.

> **COLD-START PRE-READ NOTICE (read before this document, before opening ANY VBA/config).**
> Canonical read order is enforced by the project contract. Open, IN THIS ORDER:
> 1. `docs/AGENT_COLD_START_PROMPT_CONTRACT.md` §0.4 MANDATORY CANONICAL READ ORDER (what to read first)
> 2. Same file §0.5 MANDATORY VBA PRE-EDIT GATE (G1/G2/G3 — run BEFORE any VBA edit)
> 3. Same file §0.6 SHARED HELPER AUTHORING CONTRACT (build guard table for shared helpers like NormalizeSourceDate)
> 4. **Only then** continue reading this `ARCHITECTURE.md`, then `CONFIGURATION.md`, then registry/specs.
>
> Any order where VBA is grepped/opened before the three `AGENT_COLD_START_PROMPT_CONTRACT.md` sections
> above is a DEFECT. Build hygiene can't detect chronological order — author must self-enforce this.

An earlier version of this project mirrored `elt_ingest_excel`'s INGEST → TRANSFORM → PUBLISH shape and
`ref_*`/`src_*`/`workday_*` staging convention literally, via an Interim workbook. That was superseded
(see "Superseded: Interim workbook..." below) once it became clear the actual task here — a single
lookup join, not a multi-stage migration — didn't need it. The two projects share no code or config.

## Authoring model: judgement vs deterministic

Read this first — it is the mental model a cold-start session most often gets wrong.

Creating one transform has exactly **one step that requires judgement** (today, model-assisted):

1. **Shape the transform spec** — `config/transform_specs/<key>.yaml`: the `mapping:` section plus the
   decision record, arrived at via Q&A from the input file + target EIB.

Everything downstream is **deterministic**: the YAML contract plus the shared-helper registry define the
implementation shape, `vba_src/` must remain traceable 1:1 back to that config, and
`scripts/build_runner.py` enforces config<->code parity before the Runner is prepared. The macro's
runtime behaviour is deterministic for identical inputs, and any VBA change that cannot be justified
from config/specs is a defect.

**The VBA runtime never parses or reads YAML.** YAML files exist in three roles, all consumed
*outside* the finance-user runtime:
- `config/transforms/*.yaml` — build-time profiles. `build_runner.py` reads them to write the
  Runner's registry block and to seed the `config` sheet's defaults (file paths, mapping-sheet
  location, reference-extract paths). Once written, the Runner reads those cell values at runtime
  via `ReadControlValue`. The macros do **not** re-parse the YAML.
- `config/transform_specs/*.yaml` — authoring specs only. Read by the coding agent (and reviewed
  side-by-side with the `.bas`) to drive the `.bas` contents. Not read by any build script and
  certainly not by any VBA.
- `config/reference_paths.yaml` — build-time merge of shared reference-file paths.
  `build_runner.py` folds them into each transform's registry row and config-sheet defaults; the
  macros then read the resulting cells.

A consequence: if you only update a YAML value (e.g. `reference_paths.yaml` → new Mapping MASTER
filename) and the *shape* of the pipeline (column refs, constants, cell writes, business objects)
hasn't changed, the `.bas` source needs no textual edit. Rebuilding the Runner (shell +
`build_runner.py`) is sufficient to propagate the new value into the cells the macros read.

The transform side is **config-authoritative**. `config/transform_specs/*.yaml` plus
`config/shared_helper_registry.yaml` declare the behavior; the checked-in `.bas` must stay traceable
to that contract and `build_runner.py` enforces the parity we have mechanised so far. Do not treat the
VBA as freehand source material or bypass the config-first path when changing behavior. The deliberate
asymmetry with the **ingest** side is only that ingest has a full workbook renderer
(`scripts/create_ingest_template.py`), while transform parity is currently enforced through config +
validation rather than a single monolithic spec->VBA emitter.

### Governance gate for H3 table-driven runner engine

At **transform count < 20**, per-transform `.bas` modules are cheaper (copy/rename/edit, verified via
`scripts/scaffold_transform.py` + `scripts/build_runner.py` hygiene) than maintaining a generic
table-driven engine would be. Switching to a single `modTableDrivenRunner` that interprets spec YAML
is a **governance-reviewed change**, not an incremental edit, and is only started when BOTH
conditions hold:

1. the live transform count reaches **20**, AND
2. tracked per-new-transform authoring cost exceeds **4 hours per transform** on average.

When that happens, the design start-point is the `scripts/inspect_registry.py`-visible 14-column
Transform Registry plus each `config/transform_specs/<key>.yaml`; the engine runs alongside existing
`.bas` modules (migration is not big-bang). Current count as of this document: 9 transforms.

## Native VBA only — no COM-only libraries

This is the project's central engineering constraint, and the reason the design looks the way it
does: **no VBA code in this project may reference a COM-only type library** — no `ADODB`, no
`Scripting.Dictionary`, no `Scripting.FileSystemObject`, no `VBScript.RegExp`. Those are Windows-only;
Mac Excel's VBA host has no COM and cannot run code that references them at all (not "runs
differently" — fails outright).

Plain VBA — loops, arrays, `Collection`, Excel's own `Workbooks`/`Worksheets`/`Range` objects, native
string functions, `Open`/`Print #` file I/O — behaves identically on Mac and Windows. Restricting the
whole codebase to that subset means **the entire solution can be built and tested on a Mac**, with no
gap to how it will behave on the finance users' Windows machines. The cost is giving up SQL as the
transform language (an earlier design used ADO/SQL — rejected specifically because it could only ever
be tested on Windows, an unacceptable risk for something writing into Workday master data). Transform
logic is written as VBA functions instead.

## Config purity: where domain knowledge lives

`elt_ingest_excel` is config-driven, but not because its SQL is free of domain knowledge — column
names, mapping values, and normalisation rules are all domain-specific. What makes it config-driven is
that the domain knowledge lives in **text files interpreted at runtime** (SQL files, `order.txt`,
JSON config), while the Python stays generic plumbing that never changes when a new pipeline is added.
The thing the config-purity rule (`ai_context/pco/ai_harness/rules/governance.md`) protects is the
shared, reusable layer — not every line of logic in the system.

This project applies the same split, with VBA as the only available language:

- **Stays domain-free (the shared engine)**: the dispatcher module, the shared helpers module (file
  I/O, control-sheet reads/writes, template cloning, bulk worksheet-to-array reads), and the shared
  source-table module (`modSourceTable.bas` — header/column location, required-header validation, and
  amount/date coercion over the in-memory array). None of these modules may contain a column name,
  sheet name, or business value specific to any one transform.
- **Build-time profile data**: YAML files under `config/transforms/` may hold transform labels,
  sample file paths, control-sheet defaults, and the step-list macro registration used by the built
  Runner. They exist to keep developer defaults and transform-to-macro bindings out of Python and VBA
  source and to make new transform setups additive.
- **Carries domain knowledge (the config-equivalent layer)**: the per-`(Source Type, EIB Population
  Type)` transform modules under `vba_src/transforms/`, plus any narrowly-scoped reusable
  finance-domain helper modules under `vba_src/` (for example a shared reporting-type derivation used
  by multiple transforms). Transform modules remain the primary home for combination-specific column
  names, join keys, and mapping rules — the moral equivalent of `elt_ingest_excel`'s SQL files.
  Shared domain helpers are allowed only when the same business rule is already reused across
  multiple transforms; keep generic workbook/runtime plumbing separate in the shared engine.

### Per-row-value generalization rule

Any value that varies per source **row** and needs a mapping (company name `FOOTBALL ASSOCIATION →
FA`, country name `UNITED KINGDOM → GB`, supplier human name → credential id, product code → Spend
Category ref id) MUST first be added as a row to the matching `ref_*.json` under
`config/ingest_lookups/` and resolved via the existing mapping lookup vocabulary (the
`Build<BusinessObject>Lookup` / `LookupMappingValue` pattern). Only values that are genuinely fixed
for the transform forever (example: `Submit = Y`, `Currency = GBP`, `Tax Option 1 = CALC_TAX_DUE`)
are allowed as hardcoded VBA constants. Reason: adding a JSON row costs ~one build and composes
across every future transform that needs the same mapping; a per-transform VBA edit costs one
31-char-named module plus ~600 lines of per-transform boilerplate and does not compose. The
`scripts/build_runner.py` hygiene gates do not yet enforce this rule programmatically; it is a
code-review requirement.

### Rejected alternative: declarative worksheet-based mapping tables

A stronger option was considered: push mapping rules out of VBA entirely into visible worksheet
tables (e.g. a "Mapping Rules" sheet listing combination → target sheet/column → source column or
lookup), with one fully generic VBA interpreter reading those rows — true config-as-data, with no
domain-specific VBA at all, closer parity with `elt_ingest_excel`'s JSON config.

Rejected for now: it requires meaningfully more engineering to build a generic rule interpreter; some
logic (phone/address cleanup, conditional business rules) does not fit a flat mapping-table shape and
would still need a VBA escape hatch, undermining the "no domain-specific VBA" goal anyway. Revisit
this only if the number of `(Source Type, EIB Population Type)` combinations grows large enough that
maintaining one VBA module per combination becomes the bottleneck — not before.

**Narrower exception carved out since this was written:** the Runner Shell does now carry one small,
generic, visible lookup sheet (`mapping` — entity/source/target rows) and a matching generic
interpreter (`modHelpers.LookupMappingValue`), used for small reference-data resolution (e.g. company
entity names) that can change independently of any transform's code. This is **not** the rejected
alternative above: it does not interpret the column-to-column transform shape (that's still one
`.bas` module per combination, the config-equivalent layer described above), it only resolves
individual reference values within rows a transform module already decided to write. Keep this
distinction sharp — do not grow `mapping`'s interpreter into a general mapping-table engine without
revisiting this rejection explicitly.

### Superseded: Interim workbook + generic table-ops/keyed-collection engine

An earlier version of this project mirrored `elt_ingest_excel`'s INGEST → TRANSFORM → PUBLISH shape
literally: a fresh Interim workbook with `raw_*`/`ref_*`/`src_*`/`workday_*` staging sheets, a generic
table-operations module (`FilterRows`/`JoinRows`/`GroupByRows`/`LookupValue` over in-memory 2D arrays),
and a hand-rolled `Dictionary`-equivalent class (`clsKeyedCollection`) standing in for
`Scripting.Dictionary`. That generic engine was the source of nearly every real bug found while first
getting this transform to run in actual Excel (Mac-VBA `Dir$` re-entrancy, `Set`-vs-`Let` object
assignment, a position-sync bug across three parallel `Collection`s) — none of them in the AR mapping
logic itself.

The actual transform this project performs is a single lookup join (source invoice rows + a customer
extract whose `Reference ID` column already is the join key), not a multi-stage migration needing
dedup/rank/normalize the way `elt_ingest_excel`'s DuckDB/SQL pipeline does for full Workday CCS
migration. Removed the Interim workbook, the table-operations module, and `clsKeyedCollection`
entirely; each transform module now opens its inputs, does its lookup with one or two native
`Collection`s, and writes straight into the cloned output EIB. See
`docs/todo/archive/build-elt_transform_excel.md` for the full bug history this superseded.

The later `modSourceTable.bas` is not a revival of that engine: it holds only the small,
proven header-location and value-coercion helpers (`FindHeaderRow`, `FindColumnIndexInRow`,
`ResolveAmount`, `NormalizeSourceDate`, …) that were being duplicated verbatim across every
transform module — not a generic relational `FilterRows`/`JoinRows`/`GroupByRows` abstraction.

## Artifacts

| Artifact | Role | Lifetime | Committed to git? |
|---|---|---|---|
| **Runner Shell** (`.xlsm`) | Generated workbook containing the control/form sheet (`Runner`) and a business-editable `config` sheet seeded from YAML `shell_defaults`. The `mapping` runtime lookup sheet (entity/source/target reference data, e.g. company-entity names) is **not** authored here — `scripts/build_runner.py` regenerates it from config on every build. Contains the bootstrap importer and no other logic. Generated by `scripts/create_runner_shell.py` from text source, not hand-edited. | Regenerated whenever the shell layout or bootstrap source changes. | Yes — generated build input. |
| **`config/ingest_lookups/ref_company.json`, `ref_default.json`** | Self-describing reference tables (`description`/`key_column`/`provenance`/`rows`) — the same mechanism as every other `ingest_lookups` dataset. `ref_company` holds source→target company names; `ref_default` holds per-domain defaults (e.g. currency→GBP). `scripts/build_runner.py` folds them into the Deployed Runner's `mapping` sheet (entity/source/target) at build time. Source of truth for the runtime company/default lookup — replaces the retired `RunnerShellMapping.xlsx`. | Edited by hand when the reference data changes; the `mapping` sheet is regenerated from them. | Yes — source of truth for the runtime mapping lookup. |
| **`config/transforms/`** | YAML build-time transform profiles. Defines shell defaults, sample file paths, source layout, transform metadata, and the macro binding. Every `*.y*ml` here is globbed and parsed as a runnable profile by the build. | Extended whenever a new transform type is introduced. | Yes — source of truth for build-time defaults. |
| **`config/reference_paths.yaml`** | Single source of truth for the shared reference files (Workday customer extract, supplier extract, finance Mapping MASTER + tab). The build injects these into each transform's registry row unless the profile's `shell_defaults` overrides the key (profile wins). Loaded once per build; not globbed as a profile. | Edited when a Workday extract or the finance mapping master is reissued. | Yes — source of truth for the shared reference-file paths. |
| **`config/transform_specs/`** | Per-transform canonical authoring files (see `docs/ADDING_A_TRANSFORM.md`). Hold the mapping, derived-field rules, shared-helper selections, and the finance decision record. Deliberately outside `config/transforms/` so the build never parses them as profiles. Parsed by hygiene/parity tooling to ensure the checked-in VBA stays justified by config. | Written/edited whenever the mapping or a coding decision changes. | Yes — source of truth for transform behaviour. |
| **`vba_src/`** | Plain-text `.bas`/`.cls` implementation artifacts for every VBA module, including the bootstrap importer: dispatcher, shared helpers, shared source-table operations, shared finance-domain helpers, and one module per transform step. | Updated only in lock-step with the canonical config/specs and shared-helper registry. | Yes — text implementation artifact imported into the Runner, but never authoritative over config. |
| **Deployed Runner** (`.xlsm`) | Runner Shell + every module from `vba_src/`, synchronised by `scripts/build_runner.py`. This is what finance users actually open. Carries the `mapping` lookup tab, regenerated from `config/ingest_lookups/ref_company.json` + `ref_default.json` at build time, and the hidden `lookups` runtime-lookup tab (see "Runtime lookup sheet"). | Rebuilt whenever `vba_src/` changes. | No — build output. |
| **Output workbook** | A clone of the Target EIB template, populated directly from the source/customer-extract lookup join — no intermediate workbook. | Produced at the end of a successful run. | No — run output. |
| **`RunnerLog.csv` / `RunnerStatus.txt` / `<output>.rejects.csv`** | Run observability files written by every transform via native VBA file I/O (no `FileSystemObject`). `RunnerLog.csv` (append-only run history) and `RunnerStatus.txt` (overwritten with live progress) are runner-scoped singletons written beside `Runner.xlsm`. The rejects file is written **beside the run's output EIB**, named to pair with it (`AP_Travel_output.xlsx` → `AP_Travel_output.rejects.csv`), reset per run, and lists each rejected source row with its reason and the source value the reject turned on (`BuildRejectLogPath` in `modHelpers`; falls back to `RunnerRejects_<transform>.csv` beside `Runner.xlsm` if the output location can't be resolved). Best-effort; see `docs/RUNBOOK.md`'s "Watching a Run". | Written on each run; survive rebuilds; user-deletable. | No — run output. |

Every sheet in every workbook is visible — nothing is hidden.

## Core dataflow (runtime)

```
Source Data, Customer Spreadsheet (user file paths)         Target EIB (user file path)
        |                                                            |
        v                                                            |
  RUNNER: build a lookup Collection from the customer extract        |
  (keyed by each pipe-split "Reference ID" -> Workday Customer ID)   |
        |                                                            |
        v                                                            |
  RUNNER: read the source rows, look up each row's customer,         v
  resolve company/date/amount, and write the row directly  --->  OUTPUT workbook
  into the cloned target EIB sheet                                (clone of Target EIB,
                                                                     populated in place)
```

## Components

### Runner Shell (the generated build input)

- **Control sheet** (`Runner`) — the per-run input/output surface a user touches: a `Transform`
  dropdown driven from YAML `transform.label`, plus auto-populated `Source Type`, `Source Data`,
  `Target EIB`, `Output File Path`, `Mapping Spreadsheet`, `Customer Spreadsheet`, `Supplier Spreadsheet`,
  `Workday Reference IDs`, and `EIB Population Type`, a Run button, and a results area (rows processed,
  rows skipped, unresolved-lookup flags from the last run). `Source Type` and `EIB Population Type`
  resolve from the transform registry; the editable path/file defaults resolve from the workbook's
  `config` sheet. `Output File Path` is required and has no auto-derived fallback — each transform
  profile must set its own distinct path, even when two profiles share the same `Target EIB`.
- **Config sheet** (`config`) — business-editable machine-specific defaults seeded from YAML
  `shell_defaults` at build time. The Runner sheet's path/reference fields look up their values here by
  selected `Transform` label so operators can adjust local file paths without rebuilding the workbook.
- **Mapping lookup sheet** (`mapping`) — entity/source/target reference rows (e.g. `company` /
  `FA` / `THE FA GROUP - FA`) read live, every run, via `modHelpers.LookupMappingValue`. This is
  reference data that can change independently of a transform's field mapping, so it is read
  at runtime rather than baked into the VBA — contrast with the per-transform `mapping:` section in
  `config/transform_specs/<key>.yaml`, which is the canonical declaration of the field-mapping shape
  and shared-helper usage that the checked-in VBA must satisfy.
- **Runtime lookup sheet** (`lookups`, hidden) — long-format `lookup_name | key | field | value` rows
  read via `modHelpers.LookupRuntimeValue(runnerWorkbook, lookupName, key, field)` (cached into a keyed
  Collection on first use). Added 2026-10-12 for the Customer / Supplier master EIBs, which must turn
  the display values chosen in the ingest templates into Workday reference IDs: display -> `reference_id`
  (payment term/type, categories, status, group bands, delivery method, email recipient options, tax
  statuses, PO issue option, bank account type), country code -> `country_name` / `phone_code` / `iso3`, and
  `"<ISO2>|<region>"` -> region `reference_id`. Composed by `scripts/build_runner.py`
  (`RUNTIME_LOOKUP_SOURCES`) from `config/ingest_lookups/*.json`, which stays the single source; rows with
  an empty value (e.g. a region with no resolved ID) are omitted so the reader returns an empty string.
  It is separate from `mapping`, whose rows and reader are unchanged. `LOOKUP_SHEET_NAME` (VBA) and
  `LOOKUP_SHEET_NAME` (`build_runner.py`) are checked for parity at hygiene time.
- **Bootstrap importer macro** — sourced from `vba_src/bootstrap/BootstrapImporter.bas` and injected by
  `scripts/create_runner_shell.py`. Given a folder path, it imports every `.bas`/`.cls` file found into
  the workbook's VBA project and saves. This is what lets every other module be edited as a text file,
  never inside the Excel VBA editor.

### `vba_src/` — all logic, as text

- **Dispatcher module** (`modDispatcher.bas`) — reads the selected `Transform` label from the control
  sheet, looks up the configured macro name from the Runner sheet's visible transform registry, and
  calls that one Sub directly. The transform's underlying runtime key still comes from YAML
  `transform.key`, but the operator-facing selection is the human-readable label. The dispatcher
  contains no transform-specific logic itself — adding or reassigning a supported transform should be a
  config change plus the presence of the target macro, not a dispatcher branching edit.
- **Shared helpers module** (`modHelpers.bas`) — cross-cutting concerns used by every transform:
  opening/validating external files (via `Workbooks.Open`, never `FileSystemObject`), cloning the
  Target EIB template, clearing previously-published output rows (`ClearPublishedRows`), reading a
  worksheet's used range into an in-memory 2D array in one bulk read, reading/writing the control
  sheet's labelled cells, and run status/progress/log emission. This module stays domain-free.
- **Shared entity-block writer** (`modEntityBlockWriter.bas`, registry concern `entity_row_block` /
  variant `item_index_block`) — writes ONE master-data entity (customer, supplier) as a block of
  consecutive EIB rows sharing the Spreadsheet Key in column B: `EntityBlockReset`,
  `EntityBlockAddItem(colLetter, itemIndex, value)`, `EntityBlockAddAncestor(colLetter, value,
  areaColLetters)`, `EntityBlockFlush(ws, startRow, spreadsheetKey)` (returns the next free row). Block height =
  the largest item index; row k gets item k of each list; the parent Row ID columns are repeated on
  continuation rows that carry data for their area. Generic (no customer/supplier strings). Code-complete,
  **pending the manual Excel import** (not yet run in Excel); no transform selects it until the
  `ar_customer_master` / `ap_supplier_master` changes.
- **Shared source-table module** (`modSourceTable.bas`) — the general table-operations layer that
  operates on the in-memory array produced by `modHelpers.ReadWorksheetTable`: `ColumnIndex`,
  `FindHeaderRow`, `FindColumnIndexInRow`/`FindOptionalColumnIndexInRow`, `RowContainsHeader`,
  `ValidateHeaderRow`, `ResolveAmount`, and `NormalizeSourceDate`. Domain-free and reused by every
  transform, so header-location and value-coercion logic is not duplicated per profile.

  `NormalizeSourceDate` has a DECLARED CANONICAL IMPLEMENTATION CONTRACT in
  `config/shared_helper_registry.yaml` under `source_date_normalization.excel_serial_and_dmy_text`.
  A cold-start author MUST implement it by reading that shared-helper variant first — NOT by
  guessing from example VBA. The declared branch order, execution invariants, and anti-drift
  rationale live in the YAML and are the single source of truth; the VBA in `modSourceTable.bas`
  is the checked-in implementation artifact constrained by that contract. Critical invariants stated in the YAML (numeric
  Excel-serial detection by Variant subtype before any delimiter rewriting, `Fix()`-based time
  truncation, locale-aware decimal separator for text serials, conditional dot→slash rewrite
  gated by all-3-parts-numeric) are enforcement of known real-world bugs observed on the Travel
  Places source feed (row-394+ General-formatted date cells leaking their raw serial decimals
  into the output EIB when a naive `IsDate()`/`CDate` fallback was the only numeric branch).
  See `config/transform_specs/ap_travel_places.yaml` fields `invoice_date` and `receipt_date`
  for the profile-level documentation of the exact mixed-formats-in-one-sheet scenario that
  forced the canonical contract.
- **Shared finance-domain helper modules** (for example `modFinanceCommon.bas`) — reusable business
  rules already proven common across multiple transforms: Reference-ID→Workday-ID lookup building
  (`BuildReferenceLookupFromWorkbook`, `BuildDelimitedReferenceLookup`, `LookupCollectionValue`),
  company-code mapping (`ResolveMappedCompany`), reject-summary/count aggregation, and the AMEX
  reporting-type and invoice-period derivations. These hold domain logic that would otherwise be
  duplicated verbatim in several per-profile modules, while still avoiding a general rule-interpreter
  design. For new work, treat `modFinanceCommon.bas` as the first place to look before creating a new
  finance mapping/derivation function in a transform module.
- **Per-profile transform modules** — one module per supported transform profile, each exposing a single
  `Public Sub` (the configured macro name) that does the whole run: open inputs, build any lookup
  `Collection`(s) it needs directly from a source file, iterate the source rows, and write each row
  straight into the cloned output EIB. This is where domain knowledge (column names, join keys, mapping
  rules) lives — see "Config purity" above. Generic table-operations (header/column location, header
  validation, amount/date coercion) live in `modSourceTable.bas`; reusable finance mappings/derivations
  live in `modFinanceCommon.bas`. There is no rule-interpreter layer. New reusable finance
  mappings/derivations should be added to `modFinanceCommon.bas` and new generic table operations to
  `modSourceTable.bas` by default, while transform-only column-to-target mapping and row-writing logic
  stay local to the profile module.

### Output workbook

Produced by cloning the Target EIB template (filesystem copy) and writing each processed source row
directly into the matching EIB worksheet, starting at the expected row, as the source is iterated —
no intermediate workbook or staging sheet in between.

## Deployment (build tooling — not part of the runtime)

The build is a Python-prep stage followed by one required manual step in Excel — it is **not** fully
CLI-automated. An earlier version injected every module directly into the VBA project binary via
`pyOpenVBA`; that was reverted after it produced workbooks Excel itself could not repair-save (see
Gotchas in the archived build anchor, `docs/todo/archive/build-elt_transform_excel.md`). `pyOpenVBA` is now only ever used
to create a **fresh** workbook's VBA project from scratch (step 1 below) — never to modify an existing
one.

1. `scripts/create_runner_shell.py` reads a transform profile YAML from `config/transforms/`, creates
   `runner_shell/RunnerShell.xlsm` from scratch, uses `pyOpenVBA` to create a valid `.xlsm`/VBA
   project and inject `vba_src/bootstrap/BootstrapImporter.bas`, then uses `openpyxl` to lay out the
   single visible control sheet and results area with profile-driven defaults.
2. `scripts/build_runner.py` (no flag) copies the Runner Shell to `dist/Runner.xlsm`, writes the build
   status cells' source-folder input, and writes a visible transform-registry block sourced from
   `config/transforms/*.yaml`. It does **not** import any transform module itself. Paths it writes
   into the workbook pass through the OS-selected path-rewrite profile in
   `config/root_path_substitutions.yaml` — a no-op on the macOS build host, but on a Windows build
   host it (1) rewrites the mac data root baked into `reference_paths.yaml`/`transforms/*.yaml` to
   the local build host's data root, and (2) forward-slashes the `vba_src` source-folder path so
   `BootstrapImporter.DetectSeparator` parses it (see `docs/BUILD_WIN.md`). This is build-host
   tooling config; the VBA runtime never reads it.
3. **Manual step:** a developer opens `dist/Runner.xlsm` in Excel and runs
   `BootstrapImporter.BootstrapImportFromConfiguredSource` (Tools > Macro > Macros). This calls the same
   `VBComponents.Import` mechanism a developer would use by hand, reading the source folder from the cell
   step 2 wrote. Driving this headlessly via AppleScript was attempted and abandoned — it broke
   unpredictably mid-import across several distinct failure modes, while the same import succeeds
   reliably every time it's run interactively. The same step also creates the finance-facing **Run
   Transform** button (see `docs/RUNBOOK.md`'s "The Run button"): it must be created here
   rather than in the shell because its `OnAction` binds to a macro that does not exist until this
   import runs, and because step 2 round-trips the workbook through openpyxl.
4. `scripts/build_runner.py --verify` confirms the resulting module set, registry, and status cells match
   expectations.
5. The build is run solely by the project owner, from their Mac, by repo policy. There is no separate
   Windows build step and no CI pipeline — the deployed Runner produced on Mac is shipped directly to
   finance users as-is.
6. This means an engineer or AI agent only ever edits plain text (`vba_src/`) plus the Python build
   tooling — the Runner Shell binary is generated, not opened to add or change a macro manually; the one
   exception is running the bootstrap-import macro itself (step 3), which is inherently an Excel-side
   action, not a text edit.

Because the runtime code is restricted to native VBA only (no COM-only libraries), the deployed
Runner's behaviour can be exercised end-to-end on the same Mac used to build it — including TRANSFORM
and PUBLISH — closing the testing gap a SQL/ADO-based design would have left open.

## Invariants and assumptions

- **Hard scope boundary with `elt_ingest_excel/`.** `elt_transform_excel/` is the POST Go-Live,
  operations-side interim Excel-based EIB builder. `elt_ingest_excel/` is the migration-phase
  GP / Proactis → Workday legacy ingest pipeline, whose configs are frozen to the migration
  sign-off audit trail. **An agent task whose primary scope is this project (`elt_transform_excel/`)
  MUST NOT produce edits under `elt_ingest_excel/`.** Do not silently "synchronise" Mapping MASTER
  filenames, reference paths, or any other shared-looking values across the boundary. Even if
  `elt_ingest_excel` references a filename that resembles one just updated here, that is NOT a
  signal to touch it. Cross-boundary changes require an explicit user ask that names both trees.
  See also `ai_context/project/governance/PROJECT_RULES.md`.
- No module under `vba_src/` references a COM-only type library (`ADODB`, `Scripting.Dictionary`,
  `Scripting.FileSystemObject`, `VBScript.RegExp`, or equivalent). **This is now enforced by
  `build_runner.py` Guard (vii) `PORTABLE_VBA_NO_ACTIVEX` (ERROR severity) plus `check_vba_hygiene.py`
  Step G1's cross-reference check in `AGENT_COLD_START_PROMPT_CONTRACT.md` §0.5.** Every source line
  of every `.bas` / `.cls` under `vba_src/` (AP, AR, shared, bootstrap) is scanned against a 42-entry
  banned-reference list (class names like `Scripting.Dictionary` / `ADODB.Connection` /
  `VBScript.RegExp`, plus all `CreateObject("Scripting.`) / `CreateObject('ADODB.`) late-bound PROGID
  call variants with both VBA quote styles). Comment-only lines (first non-space char is a VBA
  apostrophe) are excluded so modules can name the banned libraries in rationale comments. Any
  executable-line match raises a hard build ERROR with a remediation block that points to the
  Mac-safe VBA equivalents (Collection for dictionary semantics; `Open` / `Print #` / `Input #` for
  file I/O; Worksheet-array reads + hand-rolled tokenization for SQL / regex). This guard was added
  after the AUG-2026 CFA Claims incident where late-bound
  `CreateObject("Scripting.Dictionary")` compiled cleanly and passed all 6 existing guards but
  raised Run-time error '429' "ActiveX component can't create object" on Finance users' Windows
  laptops (Scripting Runtime ActiveX blocked by Trust Center group policy). Before the guard was
  added this rule was review-only and was missed in a cold-start session.
- **Every `.bas` / `.cls` file under `vba_src/` is source-formatted as ASCII text with CRLF line
  endings.** Not UTF-8, not LF-only. Mac Excel's `VBComponents.Import` parser silently aborts and
  falls back to a default `Module1` placeholder when multi-byte UTF-8 appears in volume (even inside
  comment blocks) or when line endings are LF-only, and the resulting build fails with no useful
  diagnostic. Pure ASCII + CRLF is the only combination that imports reliably on every run. A tiny
  amount of non-ASCII (≤ ~20 bytes total, e.g. 2 em dashes as in `modHelpers.bas`) is tolerated by
  exception, but new modules should aim for 0 non-ASCII bytes. Prefer ASCII equivalents in comments
  (`--` for em dash, `-` for box rules, `->` for arrows, straight quotes for smart quotes). See
  `docs/ADDING_A_TRANSFORM.md` "Source-file hygiene" for the full checklist.
- The Runner Shell is generated from text source and contains the bootstrap importer sourced from
  `vba_src/bootstrap/BootstrapImporter.bas`; every other module is text-sourced and only ever enters
  the deployed workbook via the build step.
- Sample file paths, shell defaults, and transform step-list macro registration used by Python build
  tooling live in YAML under `config/transforms/`, not as hardcoded literals in the build scripts.
- The Runner workbook never writes into the Target EIB template itself — it always clones first.
- The Runner workbook holds no run data — there is no Interim workbook to draw that separation against
  any more (see "Superseded: Interim workbook..." above); a run's only persistent artifact is the
  cloned output workbook.
- **Visible-sheet layering invariant for Finance handover.** The Runner workbook has exactly THREE
  user-visible worksheets: `config` (Finance-user editable shell defaults), `Runner` (dropdown + run
  button + status), `mapping` (readonly if shipped). ONE additional worksheet `registry` lives
  alongside them and is ALWAYS set to `sheet_state = "veryHidden"` at build time (see
  `scripts/build_runner.py` `_set_sheet_very_hidden`; verified by
  `scripts/inspect_registry.py --strict`). No 4th visible sheet. Finance users never edit anything
  outside the `config` tab and the `Runner` dropdown. The very-hidden state is only accessible via
  VBA (not via the Excel Format → Sheet → Unhide UI) so a Finance user cannot accidentally expose
  the 14 routing columns.
- The dispatcher never contains transform-specific logic; that lives only in the per-combination
  transform modules.
- The shared engine (dispatcher, shared helpers, shared source-table ops) never contains a column name, sheet name, or business
  value specific to one transform; reusable finance-domain logic may live in dedicated shared finance
  helper modules, while combination-specific domain knowledge stays in the per-profile transform
  modules (see "Config purity" above).
- A run that cannot resolve every lookup still produces output — with flags — rather than failing the
  whole run or silently dropping rows.
- **No hardcoded file paths in VBA.** Reference-file and template paths reach the macros *only* via
  `ReadControlValue(controlSheet, CONTROL_LABEL_*)` against values the build seeded into the
  `config` sheet or transform registry. Paths live in YAML (`config/reference_paths.yaml` for
  shared paths, `config/transforms/*.yaml` for per-profile overrides). If a path changes, update
  the YAML and rebuild — never edit a `.bas` string literal.
- **`CDec` is forbidden in VBA.** It has no analogue on Mac Excel and silently drops or corrupts
  values when round-tripped through the cross-platform VBComponents import. Use string-based
  formatting instead: `Format$(expression, "0.00")` for 2-decimal currency renders, or explicit
  string-to-long scaling via integer arithmetic, and rely on Excel's cell formatting for display.
- **`ColumnIndex("XX")` literal-string invocations are allowlisted.** Every two-letter column index
  in a `.bas` must be either (a) on the `SAFE_SHARED_COLUMNS` allowlist defined in
  `modSourceTable.bas` (columns whose meaning is truly shared across every transform, e.g.
  reference-ID columns) or (b) declared in the transform profile's `eib_columns` map. Raw
  `ColumnIndex("GB")` / `ColumnIndex("MN")` / `ColumnIndex("KE")` literals that belong only to one
  transform's EIB layout must be represented as `ColumnIndex(EIB_COL_GB)` / constants, never as
  bare strings.
- **VBA constants are agent-driven, not operator-edited.** Constants like `MAPCOL_OLD_SEGMENT`,
  `MAPCOL_WORKDAY_REF_ID`, `MAPVAL_CORE_ONLY`, `WBO_REVENUE_CATEGORY`,
  `WBO_SPEND_CATEGORY`, and per-transform `WBO_*` business-object names live at the top of each
  `.bas` (or `modFinanceCommon.bas` for shared ones) and exist because the coding agent *read the
  corresponding YAML* (`config/transforms/*.yaml`, `config/reference_paths.yaml`'s mapping-sheet
  contract, or the spec's `worktag_business_objects`) and wrote them. Manual edits to these
  constants without a matching YAML/spec update drift the codebase out of alignment with the
  config that is meant to drive it, and the `build_runner.py` hygiene gate rejects them.
- **Archetype constraints — categorically forbidden business objects.** Two WBO combinations are
  always invalid and must be caught at authoring time, not deferred to the runtime:
  - **Supplier Invoice (AP)** profiles MUST NOT contain a `WBO_REVENUE_CATEGORY` lookup or write
    to the Revenue Category EIB cell. The object is categorically absent on AP invoice payloads.
  - **Customer Invoice (AR)** profiles MUST NOT contain a `WBO_SPEND_CATEGORY` lookup or write to
    the Spend Category EIB cell. The object is categorically absent on AR invoice payloads.
  If either is found during a transform review, remove the constant and its call sites rather than
  trying to "map" an EIB field that Workday won't accept for that business process.
- **EIB template identity invariant — same version number does NOT mean same business object.**
  Two EIB xlsx files can ship in the same Workday release (e.g. `v46.1`) and share the five
  scaffolding columns (`A Fields`, `B Spreadsheet Key*`, `C Add Only`, `BQ Row ID*`,
  `CI Row ID*`) but have less than 3% semantic overlap in the remaining column map. Treating
  "same version, looks similar, has sample rows" as a drop-in replacement is a trap that silently
  miswrites every populated cell to the wrong Workday field, producing an output Workday will
  reject row-for-row.
  - **`Submit Supplier Invoice - v46.1`** is the AP supplier-invoice target. Mandatory signatures:
    worksheet name = `Submit Supplier Invoice`, Row 1 banner includes that exact string,
    col `O` (Row 5 header) = `Supplier*` (registered Workday Supplier record),
    worktags block occupies cols `DP..EX` (Spend Category, Reporting Type, Cost Centre, Event,
    Camp, FA Learning Course, Commercial Partner, Grant, Programme, **Project** — 10 worktags
    with Project as col EX), line amounts at `ED/EE`, line tax triple at `DS/DT/DV/DY/DZ`,
    line Row ID at `DD`, Memo at `EO`.
  - **`Submit Ad hoc Payment - v46.1`** is a completely different EIB for one-off Ad Hoc Payees
    who have no Supplier record (e.g. `AD_HOC_PAYEE-6-14` style references). Mandatory
    signatures: worksheet name = `Submit Ad hoc Payment`, Row 1 banner includes that exact string,
    col `K` = `Bank Account*` (mandatory, no Supplier Invoice equivalent), col `L` = `Payee`,
    col `O` = `Single Use`, no `Supplier*` field anywhere in the D..P header area, only 104
    total columns (vs 176), worktags block at `AY..BF` (8 worktags, **no Project column** at all),
    line amounts/quantity at `AP..AU` and line tax at `AK..AO`.
  - **Enforcement at authoring time, not at runtime.** The VBA dispatcher does no header-map
    fingerprinting; it trusts that the profile's `eib_population_type` and `shell_defaults.target_eib`
    describe the same worksheet template. Before setting `target_eib` on ANY profile, verify in
    order: (1) Row 1 banner verbatim text, (2) worksheet title (the tab name Row-1 banner names),
    (3) 3 non-adjacent header semantics from the transform spec — e.g. for a Supplier Invoice
    transform, confirm col `O` = `Supplier*` + col `DP` = `Spend Category` + col `EQ` = `Cost Centre`.
    If those three don't all match, you are looking at the wrong EIB; do not assign it to the
    transform. Consequences for ap_stationery / ap_taxi (and all other AP Supplier Invoice
    transforms): `target_eib` must always point at `Submit_Supplier_Invoice_v46.1.xlsx` (or a
    renamed clone), NEVER at any `Submit Ad hoc Payment*` xlsx.
- **AR Bulk / AR Multi-Line mapping writes 11 specific EIB v46.1 cells.** Both AR transform types
  resolve the account-string coding against the Mapping MASTER and distribute the 8 worktag
  results across these *exact* columns:
  - `KE` — Revenue Category (AR-specific, per archetype constraints above)
  - `GB` — Project Header reference
  - `MN` — Project Line reference
  - `MV … NC` — the 8 consecutive Worktag slots (MV, MW, MX, MY, MZ, NA, NB, NC) for Event / Camp
    / FA Learning / Commercial Partner / Grant / Program / Project / (reserved-8th), populated in
    the order declared by the profile's `worktag_business_objects` list (or left blank when that
    business object has no Core mapping for the given Final Code).
  Any new AR account-string transform mirrors this exact 11-cell write layout; do not invent new
  slot column assignments.
- **AR lookup-unresolved default: silent blanks, not row rejection.** For AR Bulk and AR
  Multi-Line transforms, when a Final Code segment does not resolve against a given business
  object in the Mapping MASTER (and no `unresolved = abort` was explicitly configured), the
  output cell for that slot is written **blank**, not `"N/A"`, not an error flag, and the row is
  NOT rejected. AR invoices are allowed to have sparsely-populated worktag slots; the Workday
  import validates required-ness per integration. Contrast with AP transforms, which do reject
  truly-required misses (e.g. Cost Centre) and count them.
- **AR Invoice GBE resolves the customer by CLUB NAME, not a Reference ID.** `ar_gbe_invoice`
  (Governing Body Endorsement) is the only AR transform whose customer key is the club NAME
  (source col `Club`), because the GBE extract carries no GP Reference ID. `modArGbeInvoiceCustInvoice`
  matches the club against the customer extract's `Customer` (name) column, normalised (upper-case,
  whitespace collapse, `WFC` -> `WOMEN FC`, `WOMENS` -> `WOMEN`). Resolution is precision-first:
  an exact normalised match wins; a miss OR an ambiguous name (same normalised name -> >1 Customer
  ID) is rejected to the reject log, never guessed. This is a known-lossy key (the extract
  abbreviates clubs differently), so a material share of rows reject until Finance supplies a
  Club -> Customer-ID crosswalk. GBE has no COST CODE column, so Revenue Category + the 9 worktag
  slots stay blank (the account-string pipeline is wired inert for archetype parity). Amounts/tax
  mirror AR Bulk (GROSS -> `CM`/`FL`/`FO`; VAT rate derived from `(Gross - Net)/Net` so the Womens
  tab, whose header omits the `VAT` label, is handled identically).
- **`fixed_literal_worktags` over `worktag_business_objects` for truly-fixed values.** When a
  profile needs a slot populated with a constant ref id for every single row (e.g. AMEX category
  always resolving to one ref id regardless of coding string), declare it under
  `fixed_literal_worktags` in the profile YAML — **not** under `worktag_business_objects` with a
  constant source. `worktag_business_objects` triggers a hierarchy lookup (segments vs Core rows),
  and a literal constant there is a category error that surfaces as confusing hierarchy-match
  misses, phantom "unresolved" counts, and config-drift failures when the lookup semantics are
  tightened.
- **Banner-induced row shifts when formula-seeding the shell.**
  `scripts/create_runner_shell.py` writes text rows first (banner section + registry section +
  result section + control section + legend), each of which consumes physical rows in the
  generated `Runner` sheet. When adding or adjusting formulas that reference cells *within* the
  control section (for example auto-populated path cells whose formula is
  `=IFERROR(INDEX(config!$A$2:$G$100,MATCH(Runner!$C$12,config!$A$2:$A$100,0),5), "")`), you
  MUST recalculate the actual row number that cell ends up on after the banner is rendered, not
  use the design-time offset. A 1- or 2-row banner-offset mistake results in every formula
  returning `#N/A` on the built shell. Do a shell rebuild + `inspect_registry.py --strict` after
  touching this logic, and open-check the actual cells in Runner.xlsm before considering it done.

## Non-Obvious Gotchas (Tier-1 / Stage-1 implementation traps)

These were discovered during the AP-Tier-1 helper-hoisting + Stage-1 shared-helper-registry
anti-drift rollout. Each is a "correct-but-silent design decision" that a future cold-start
agent will re-misstep on unless called out explicitly. For the index-pointers version that
cold-start assistants see first, see `ai_context/project/ASSISTANTS.md § Gotchas →
elt_transform_excel repeatable traps`.

1. **YAML unquoted scalars containing `: ` (colon-space) fail to parse as YAML.**
   PyYAML interprets `: ` at end of a plain scalar token as a mapping-key separator, even inside
   prose. This bites inside:
   - `notes:` / `rule:` free-form blocks (e.g. `"rule: V2: Default Cost Code → Final Code"`)
   - flow-mapping rule scalars like `{source: Net (fallback: Gross - Tax), rule: Rule 9 concat: X + ' ' + Y}`
   If a scalar has ANY interior `: `, wrap it as a `|-` literal block scalar (for multi-line
   prose) or as a double-quoted string (for short inline scalars). Pre-existing latent examples
   that had to be fixed when Step 6 started parsing `transform_specs/*.yaml` with safe_load:
     - `ap_amex_hotel.yaml` `label:` (bare `"V2: Default"` scalar), `notes:` + `rule:` blocks
     - `ap_amex_trains_v2.yaml` `reporting_type.rule:` block
     - `ap_taxi.yaml` `note:` line 164 `FIRST CLASS exemplar: …` + lines 364/367 flow-mapping values
     - `ap_travel_places.yaml` lines 344/346 flow-mapping `Net (fallback: …)` and `Rule 9 concat: …`
   *Symptom:* `yaml.scanner.ScannerError: mapping values are not allowed here` pointing at the
   colon-space. *Fix:* Literal block `|-` or wrap problematic scalar in `"…"`.

2. **`shell_defaults.variants:` REPLACES the base binding entirely — the base `profile_key`
   DISAPPEARS from `transform_bindings`.**
   `_expand_shell_variants()` returns only variant-derived bindings when `variants:` is
   non-empty. It NEVER appends `base_binding` alongside them. Example:
   `config/transforms/ap_stationery.yaml` declares 4 variants → `transform_bindings` has 4
   entries (`ap_stationery_addison_lee_fa`, `ap_stationery_ambassador_taxi`,
   `ap_stationery_dhl`, `ap_stationery_e_courier`) and ZERO plain `ap_stationery` key.
   Any Step-6 code that resolves a transform_spec from `binding.profile_key` MUST fall back to:
   (a) other bindings that share the same `profile_to_module[*] → module_path` mapping
   (b) existing specs whose filename STEM is a `<prefix>_` match of the variant profile_key.
   *Symptom:* false-positive "transform spec file not found" errors ONLY for modules whose
   build profile declares variants. *Fix:* dedupe spec validation by module_path (not by
   binding_profile_key) with a `module_to_spec_shared_helpers` cache, and apply the
   exact-match → same-module → stem-prefix-match 3-tier candidate lookup.
   `config/transforms/ar_gbe_invoice.yaml` is a second variants user (`ar_gbe_invoice_mens`,
   `ar_gbe_invoice_womens`) and the first to override `source_worksheet_name` PER VARIANT: both
   variants point at the SAME workbook file but different tabs (`Mens GBE` / `Womens GBE`) — the
   variant field map routes `source_worksheet_name` → the binding's `worksheet_name`, so one file
   with two tabs becomes two dropdown selections sharing one macro.

3. **Step 6 POS-guard hint helpers (concern/variant → bare name tokens) must be DIRECTLY CALLED
   in the transform `.bas` body. You CANNOT list helpers that are only transitive-called INSIDE
   a hoisted public function.**
   The POS guard `_bare_name_is_referenced(raw_bas, name)` performs a simple token match against
   the transform `.bas` source bytes only. It does NOT see the body of `Public Function`
   implementations in `modSourceTable.bas` or `modFinanceCommon.bas`.
   - `NormaliseToWordBoundaries` is an *internal implementation detail* of the hoisted
     `RowIsEmptyOrTrailer()` in `modSourceTable.bas`. Even though declaring
     `row_trailer_detection: text_cols_and_amount_cols_parameterized` semantically implies the
     word-boundary normalizer runs, the token will NOT appear in ANY AP transform `.bas` body.
     Including it in `CONCERN_VARIANT_TO_BARE_HELPER_HINTS` triggers false negatives on 100%
     correct modules.
   - `FinalCodeSegment` is an *internal helper* in `modFinanceCommon.bas` that powers the
     public entry points `ResolveReportingTypeFromFinalCode` and
     `ResolveLowestHierarchyLookupAcrossFinalCodeSegments`. The transform `.bas` calls the
     public names, not `FinalCodeSegment` directly. Same rule.
   *Add a name to the hints table ONLY after* grep confirms the bare-name token is present in
   at least one of `vba_src/transforms/modAp*.bas` outside a `Private Function` declaration
   line. Symptom of a wrongly-added hint: Step 6 fails with "declares runtime use of helper X
   but the .bas body does NOT reference it" on modules whose call sites are correct.

4. **`_validate_vba_source_hygiene()` function signature does NOT carry config paths by default.
   Thread them through explicitly when extending Step 6.**
   Current signature: `_validate_vba_source_hygiene(vba_source, module_names, transform_bindings, profile_dir=None)`.
   The function is also used in modes where `transform_bindings=None` and no config exists.
   When you need to read `config/shared_helper_registry.yaml`, `config/transform_specs/*`,
   `config/reference_paths.yaml`, or any path whose location is derived from the
   `--profile-dir` CLI arg, you must update ALL of:
   (a) the function signature with a new typed, default-None parameter;
   (b) the `--hygiene-only` call site in `main()` + the normal build call site (BOTH in main);
   (c) every `profile_dir.parent / …` derivation inside the validator, gated with
       `if profile_dir is not None:` / `if transform_specs_dir is not None:` checks
       so the partial-hygiene call sites don't NameError.
   Symptom of missing thread-through: `HYGIENE FAILED: name 'profile_dir' is not defined` at
   the very start of the Step-6 validation block.

**5. Correct authoring model — NO exemplar-code-copy "fork + rename".**
   Every time you write a new transform, you follow this EXACT loop, every time, no exceptions
   (also captured in `ai_context/project/governance/PROJECT_RULES.md` under
   *Declarative authoring model (THIS IS THE ONLY AUTHORING MODEL)*):

   1. User provides (natural language):
      - Input source file shape + column names (which column supplies which semantic role:
        supplier ID, invoice number, net, tax, description, etc.)
      - Output EIB template identity: banner string + 3 non-adjacent header-cell semantic
        matches (per the EIB template identity hard rule in PROJECT_RULES.md — *never* trust
        same-version-numbered EIBs; verify Row 1 banner explicitly).
      - **2 explicit shape-decision rules** that FULLY determine the
        `shared_helpers:` variant block:
        (a) *Supplier/Customer mapping column + contains company prefix?*  (yes →
            `supplier_resolution=per_row_split_company_creditor` +
            `company_resolution=per_row_split_from_supplier_id`; no → choose from
            `fixed_reference_id`, `fixed_literal`, `per_row_integration_column_a`,
            `per_row_account_name_mapping` per user statement)
        (b) *Accounting-string → hierarchy mapping column + derivation shape:*  "correct
            as-is for mapping" → `direct_final_code_column`; "needs classic separator
            collapsing before segment match" → `direct_column_separator_normalized`;
            "shaped like AMEX padded pipeline (22-char padding → Seg2 map → overrides)"
            → `amex_padded_22char_pipeline`.
      - Reference to which existing transform's fine-behavior patterns (classic-6 vs
        direct-distribution vs AMEX-3) to align with, for concerns like row-trailer
        detection strategy or uniform-company bootstrap rule.

   2. You (agent) write `config/transform_specs/<base_profile_key>.yaml` DETERMINISTICALLY
      from the above:
      - `fields:` block maps every input source-column name to its semantic role with
        `source_column`, `source_column_fallback`, `post_process`,
        `derivation=per_row_lookup` and its `lookup_table` filter triple, and
        `on_unresolved` ALL explicitly declared.  No implicit "copy from exemplar"
        fields left unspecified.
        - `shared_helpers:` top-level block (after `meta:`, before `parties:`) picks
          exactly one variant per concern, derived from the user's 2 shape-decision rules.
          This now includes `invoice_line_grouping: per_invoice_number` for any transform
          that calls `AssignInvoiceLineGroup` to populate Spreadsheet Key plus Row ID / Line
          Order. Every concern/variant key declared first in `shared_helper_registry.yaml`
          if this transform introduces a new variant shape.

   3. You write the `.bas` module DETERMINISTICALLY from the transform_spec + registry,
      not from any existing transform code file:
      - Call sites reference hoisted helpers by bare name per registry `vba_public_fn`.
        Never add local `Private Function` copies of the 5 hoisted helpers (they fail
        the Step-6 NEG guard).
      - Which concrete source-columns go into `Array(textCheckCols)` /
        `Array(amountCheckCols)` / `Array(labelCheckCols)` for the parameterised
        row-check helpers: DERIVE from `fields:` block type-classification, NOT from
        reading exemplar source Arrays:
        * `textCheckCols` = every `source_column` of a field whose `derivation` is
          non-numeric AND whose `basis` is not amount-based (supplier, invoice,
          cost-code, description are always text).
        * `amountCheckCols` = every source-column referenced in a field with
          `basis: net` / `basis: gross`, or a field that declares
          `source_column_net_alias_primary` / `source_column_gross` /
          `source_column_tax`.
        * `labelCheckCols` = ordered 3-tuple (description/line-description field FIRST,
          supplier/customer mapping column SECOND, invoice-number column THIRD).
          This ordering is MANDATORY — it matches the registry's `requires_columns`
          metadata on the `row_trailer_detection.text_cols_and_amount_cols_parameterized`
          variant and produces the same normalised keyword-match string that the
          original eib19 private implementations used.
      - Orchestration skeleton (bootstrap-2row vs simple-per-row vs fixed-literal) =
        DERIVE directly from the `company_resolution` variant, NOT from reading an
        exemplar's loop code:
        * `per_row_split_from_supplier_id` / `fixed_reference_id` → simple-per-row,
          no bootstrap pre-scan.
        * `per_row_integration_column_a` → bootstrap-2row uniform scan: iterate the
          first up-to-32 non-empty/non-skippable rows, extract
          `sourceCompanyCode` + `batchIdToken`, check they are uniform (abort if
          mixed), cache them for the main loop.
        * `fixed_literal` → write one `Private Const SOURCE_COMPANY_CODE As String =
          "<user-given literal>"`; no split, no bootstrap.
        * `per_row_account_name_mapping` → per-row scan of an `ACCOUNT_NAME →
          Reference_ID` lookup table drawn from the registry / transform_spec
          `parties.suppliers.lookup_table` declaration.
      - `ResolveEffectiveNet` split-rule (mandatory, no exemplar-copy needed):
        * In `RowIsSkippable`'s "net amount zero" predicate: call single bare
          `ResolveEffectiveNet(net, gross, tax)` — RETURNS SIGNED value, the hoisted
          fn deliberately omits absolute-value because skip logic uses sign-agnostic
          `ResolveAmount(...) = 0#` comparison.
        * In `WriteOutputRow` Rule 10 (final line net write): after resolving net via
          ResolveAmount OR inline `gross - tax` fallback, ALWAYS apply the
          absolute-value guard `If netAmt < 0# Then netAmt = -netAmt` before writing
          the ED/EE cells.  Workday EIB loaders reject negative line net; sign
          lives only in Credit/Debit indicator columns, never in the amount cell.

   4. Rebuild via `uv run ... build_runner.py --hygiene-only`.  0 ERROR-class
      violations before committing.  No partial-hygiene acceptance.

   **Forbidden.**  Never copy an existing `modAp*.bas` file and `sed`/rename its
   Private Consts / source-column references.  Write fresh from the spec.  The
   "reference existing example" in step 1 is user-meaning shorthand for "apply the
   same 2 shape-decision rule choices + align on these fine-grained concern
   variants".  It is NEVER a mechanical "copy the exemplar VBA code body"
   instruction.  Copying exemplar code verbatim is the #1 re-drift vector because
   the exemplar silently carries concern-variant assumptions (e.g.
   `labelCheckCols = [description, supplierId, invoiceNo]`) that a naïve rename
   will leave stale for a transform whose mapping column is actually
   `IntegrationColumnA`.  The hygiene POS guard only checks bare-name references
   (not content of the Arrays you pass in) — so a stale Array copied from the wrong
   exemplar would PASS the hygiene gate while producing subtly-wrong trailer
   detection.  Write fresh. Every time.
