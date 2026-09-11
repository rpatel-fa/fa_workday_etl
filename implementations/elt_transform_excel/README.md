# elt_transform_excel

`elt_transform_excel` is a finance-user-run Excel ETL workflow that turns a source workbook into a populated Workday EIB workbook using native VBA only. It is intentionally separate from `elt_ingest_excel`: no Python, DuckDB, shared runtime code, or shared configuration is used in the shipped solution.

## What It Does

- Uses a Runner workbook as the user-facing control surface.
- A single configured native-VBA macro, selected by `Source Type` and `EIB Population Type`, opens the
  source and reference workbooks directly and builds whatever lookup it needs.
- Clones and populates the chosen target EIB template directly as the Output workbook — no
  intermediate workbook.

## Current State

The design contract is complete. Repository scaffolding, core docs, and the first shared VBA engine text
modules exist. `scripts/create_runner_shell.py` creates `runner_shell/RunnerShell.xlsm` from text source
and derives the generated Runner controls from the YAML transform profiles under `config/transforms/`.

The build is a Python-prep step plus one required manual step in Excel, not full CLI automation:
`scripts/build_runner.py` copies the shell to `dist/Runner.xlsm` and writes the YAML-driven
transform-registry block, but a developer must then open that workbook in Excel and run
`BootstrapImporter.BootstrapImportFromConfiguredSource` themselves to actually import the `vba_src/`
modules — an earlier fully-automated version (writing the VBA project binary directly via `pyOpenVBA`,
or driving the import headlessly via AppleScript) was tried and reverted after both proved unreliable.
`scripts/build_runner.py --verify` confirms the manual step succeeded. See `docs/ARCHITECTURE.md`
("Deployment") and the archived build anchor (`docs/todo/archive/build-elt_transform_excel.md`) for the full history.

The first `AR Invoice -> Customer Invoice` vertical slice
(`vba_src/transforms/modArInvoiceCustomerInvoice.bas`) is confirmed working end to end in real Excel,
including on-disk result persistence. A second vertical slice, `AR Discipline Invoice -> Customer
Invoice` (`vba_src/transforms/modArDisciplineCustInvoice.bas`), was confirmed by the user as producing
correct output. A third Customer Invoice slice, `AR Invoice Run` (`modArInvoiceRunCustInvoice.bas`), forks
`AR Discipline Invoice` for a flatter AR extract: fixed company (WNSL) and a customer lookup on the
source `Code` column, with no line cost centre. Output VAT is itemised: `NET` posts as the line
Extended Amount, a per-row line Tax Code (fines → zero-rated, else standard 20%) drives Workday's VAT
calculation, and gross `TOTAL` posts as Amount Due. Code-complete pending manual Excel import + verify. Three
Supplier Invoice transforms follow the same pattern: `AP AMEX Hotel`
(`modApAmexHotelSupplierInvoice.bas`), `AP AMEX Trains` (`modApAmexTrainsSupplierInvoice.bas`), and
`AP CFA Distributions` (`modApCfaDistribSupplierInvoice.bas`) — the account-string exemplar,
where Cost Centre, Spend Category, Reporting Type, and related worktags are all resolved from the
source coding string against the finance Mapping MASTER (`Core vs Prism` = Core, lowest hierarchy).
All three are code-complete pending manual Excel import + verify. Six additional AP transforms (AP AMEX Trains V2, AP CFA Distributions, AP Prize Funds, AP Referees V2 (active go-forward; V1 OG defunct), AP Stationery Companies, AP Taxi) are now also registered and code-complete; the latest batch (Stationery Companies, Taxi, Referees, Prize Funds, CFA Distributions) added the most recent profiles.

`AR Invoice GBE` (`modArGbeInvoiceCustInvoice.bas`, Governing Body Endorsement) is the newest AR
Customer Invoice slice and the first to fan ONE workbook's two tabs into two Runner selections —
`AR Invoice GBE Mens` and `AR Invoice GBE Womens` — via `shell_defaults.variants` with a per-variant
`source_worksheet_name` override (same file, different tab), the ap_stationery one-transform-many-sources
pattern applied to worksheets. It is the only AR transform that resolves the customer by CLUB NAME
(the GBE extract has no Reference ID): precision-first normalised name match against the customer
extract, with misses and ambiguous names rejected to the reject log. Known-lossy until Finance
supplies a Club -> Customer-ID crosswalk. Code-complete pending manual Excel import + verify.

The shared reference files (customer/supplier extracts, Mapping MASTER) are defined once in
`config/reference_paths.yaml` and inherited by every transform. On completion the Runner colour-codes
its result cells (green = all rows wrote, amber = rows rejected, red = failed) so mapping errors are
obvious next to the Run button — see `docs/RUNBOOK.md`.

To add a further transform, follow `docs/ADDING_A_TRANSFORM.md` — the end-to-end authoring playbook
(finance decisions to confirm, source/EIB inspection, mapping-tab shape, macro, build).

## Repository Layout

- `PRD.md` - product contract, scope, invariants, and acceptance criteria.
- `docs/ADDING_A_TRANSFORM.md` - authoring playbook for adding a new source → EIB transform (cold-start).
- `docs/ARCHITECTURE.md` - runtime/build design, artifact boundaries, and deployment model.
- `docs/RUNBOOK.md` - operator setup, current manual workflow, diagnosis, and recovery notes.
- `docs/CONFIGURATION.md` - control-sheet inputs and configuration rules.
- `docs/ops/CLI.md` - command reference for verification and future build operations.
- `docs/TODO.md` - backlog index.
- `docs/todo/` - anchor docs for multi-session backlog items.
- `config/reference_paths.yaml` - single source of truth for the shared reference files (Workday
  customer extract, supplier extract, finance Mapping MASTER + tab). Change a Workday/finance file
  here once; every transform inherits it. See `docs/CONFIGURATION.md`.
- `config/transform_specs/` - per-transform decision record + the `mapping:` section the `.bas`
  mirrors (the field-mapping source of truth). Start here when authoring a transform.
- `config/transforms/` - YAML build profiles used by the build tooling to seed shell defaults.
- `config/root_path_substitutions.yaml` - build-time filesystem-path rewrites `build_runner.py`
  applies to paths it writes into the workbook (auto-selected per OS; the Windows profile swaps
  the mac data root for the local build host's, and forward-slashes the `vba_src` path so
  `BootstrapImporter` can read it). See `docs/BUILD_WIN.md`.
- `vba_src/` - text-source VBA modules and classes.
- `build/` - build tooling that assembles the shipped Runner workbook.
- `runner_shell/` - generated Runner Shell workbook (`RunnerShell.xlsm`, the build input). The runtime
  `mapping` lookup is generated from `config/ingest_lookups/ref_company.json` + `ref_default.json` at build time.

## Key Constraints

- Native VBA only: no `ADODB`, `Scripting.Dictionary`, `FileSystemObject`, `RegExp`, or Power Query automation.
- Two artifacts remain separate: Runner Shell and Output workbook. There is no Interim workbook — each
  transform macro opens its own inputs and writes straight to the cloned output.
- No hidden sheets in any workbook.
- All Runner VBA, including the bootstrap importer, is text-sourced under `vba_src/`; the shell workbook is generated from those sources.
- Build-time defaults for sample files and transform selections live in YAML under `config/transforms/`, not as hardcoded literals in Python.
- Domain-specific transform logic belongs only under `vba_src/transforms/`.

## Entry Points

- Start with `PRD.md` for the behavioural contract.
- Read `docs/ARCHITECTURE.md` for implementation boundaries.
- Use `docs/RUNBOOK.md` and `docs/ops/CLI.md` for operating and verifying the project.
- Use `docs/ADDING_A_TRANSFORM.md` to add a new source → EIB transform from a cold start.
- Use `docs/AGENT_COLD_START_PROMPT_CONTRACT.md` as a cold agent before acting — defines the 3 valid
  prompt forms ("update transform for xyz", "create new transform for source abc → supplier/customer EIB"),
  the ambiguity-rules (when to fire follow-up questions) and the 8 required user prompts + 4-step
  deterministic pipeline with mandatory build-gate + reconcile close-out. If anything in your prompt
  isn't covered there, read that doc first.
- Use `docs/TODO.md` to resume implementation work.
