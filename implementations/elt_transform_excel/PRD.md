# PRD — elt_transform_excel

> A PRD is a contract, not documentation about code. See authoring contract: `ai_context/pco/governance/DOCS_CONTRACT.md`

## Problem

Finance users need to turn a source data extract into a populated Workday EIB (Enterprise Interface
Builder) template, applying a mapping/transform appropriate to the data type (Customer, Customer
Invoices, Customer Adjustments, etc.). `elt_ingest_excel` already solves this for technical users, but
its solution requires Python, DuckDB, and a configured dev environment. Finance users run plain Windows
Excel with no such tooling and no ability to install it. Without a self-contained alternative, finance
users either need a technical analyst to run every transform on their behalf, or they prepare EIB
workbooks by hand — slow and error-prone.

## Solution

A persistent **Runner workbook** that finance users open directly — a single worksheet holding the
control/form inputs (source type, source data file, target EIB template, reference files, and the EIB
population type) and a Run button. Running it: the configured transform macro for the selected
(Source Type, EIB Population Type) combination opens the source data and reference files directly,
builds whatever lookup it needs (e.g. a customer-ID lookup `Collection` keyed by the customer extract's
`Reference ID` field), and writes each processed row straight into a clone of the target EIB template,
the **Output workbook**. No sheet in either workbook is hidden. All runtime logic executes inside Excel
via plain VBA, with no reference to any COM-only library (no ADODB, no
`Scripting.Dictionary`/`FileSystemObject`/`RegExp`, no Power Query automation); no external runtime,
install step, or dev environment is required on the finance user's machine.

An earlier version of this solution routed every run through a fresh Interim workbook that mimicked
database staging tables (`raw_*`/`ref_*`/`src_*`/`workday_*`), with generic VBA table operations
(filter/join/group/lookup) run against it. That was removed once it became clear the actual transform
needed here is a single lookup join, not a multi-stage migration — see Known Limitations.

This is a deliberate choice, not the first design: an earlier version used SQL run via VBA/ADO, and a
later one considered Power Query. Both were rejected — see Known Limitations for why — in favour of
plain VBA, which is the only option that is both fully testable on a Mac dev machine and has no
dependency on a feature (ADO, Power Query) that Windows IT policy could disable independently of VBA
macros.

The Runner's VBA logic is authored as plain text and assembled into the deployed Runner workbook by a
build step — engineers/agents edit text source, never the Excel binary directly. See
`docs/ARCHITECTURE.md` for the deployment mechanism.

## Scope and Boundaries

**Owns:**
- The Runner workbook's control sheet collecting per-run inputs: source type, source data file,
  target EIB file, mapping spreadsheet, customer spreadsheet, supplier spreadsheet, Workday reference
  IDs file, and EIB population type.
- The mapping/transform logic for each (source type, EIB population type) combination, expressed as a
  single native VBA macro that opens its own inputs directly and builds whatever lookup `Collection`(s)
  it needs.
- Cloning the chosen EIB template and populating it directly from the resolved source rows, producing
  the Output workbook.
- Keeping the shared engine (dispatcher, helpers) free of any column name, sheet name, or business
  value specific to one transform, while allowing narrowly-scoped shared finance-domain helper
  modules for reusable finance mappings and derivations, with `modFinanceCommon.bas` as the default
  shared location before creating new local helper functions.

**Delegates:** none — this module has no dependency on `elt_ingest_excel` or any other module in this
repo. It is intentionally self-contained so it can be distributed and run independently of the rest of
the platform's tooling.

**Out of scope:**
- Any transform requiring data volumes or logic beyond what Excel/VBA can handle reasonably
  (no DuckDB, no SQL engine, no Power Query).
- Direct API integration with Workday.
- Migration-grade entity resolution/deduplication logic (that remains `elt_ingest_excel`'s domain,
  for the cases it already covers).
- Producing new (source type, EIB population type) combinations without a corresponding ordered step
  list — the Runner only offers what has been implemented, not arbitrary user-defined transforms.

## Contracts

### Inputs

All entered on the control sheet, as file paths or fixed dropdown selections (not free text):

- **Source Type** — dropdown; identifies the shape/origin of the source data (drives which staging
  logic applies to the raw extract).
- **Source Data** — file path; the data to be transformed.
- **Target EIB** — file path; the EIB template to clone and populate.
- **Mapping Spreadsheet** — file path; Finance-defined mapping reference, saved locally by the user.
- **Customer Spreadsheet** — file path; a Workday customer extract the user runs themselves.
- **Supplier Spreadsheet** — file path; a Workday supplier extract the user runs themselves.
- **Workday Reference IDs** — file path; implementation-partner-defined reference data, saved locally
  by the user.
- **EIB Population Type** — dropdown; one of Customer, Customer Invoice, Customer Invoice Adjustment,
  Supplier, Supplier Invoice, Supplier Invoice Adjustment.

Files that cannot be opened (wrong format, locked, missing expected sheet) are rejected with a message
identifying which input failed and why — not a silent partial run.

### Outputs

- The **Output workbook**: a clone of the Target EIB template, populated with mapped rows in the
  expected worksheet(s) and starting row.
- A run/result summary on the Runner's control sheet stating rows processed, rows skipped, and any
  mapping values that could not be resolved.

### Error Envelope *(required)*

- On a missing or unresolved mapping value (e.g. a customer not found in the lookup), the affected row
  is flagged in the output rather than silently dropped or silently defaulted.
- On a structural error (missing source sheet, wrong template), the macro stops before writing any
  output file and reports which input failed.
- The module never produces a populated EIB workbook that mixes successfully mapped and silently
  defaulted rows without flagging which is which.

## Requirements

### Functional (SHALL statements)

- The Runner workbook SHALL present Source Type and EIB Population Type as fixed dropdown selections.
- The combination of Source Type and EIB Population Type SHALL determine exactly one configured macro
  to run.
- That macro SHALL open its own source/reference inputs directly and build whatever native
  lookup(s) its mapping needs — no Interim workbook or generic table-operations layer.
- No module SHALL reference a COM-only type library — no `ADODB`, no `Scripting.Dictionary`, no
  `Scripting.FileSystemObject`, no `VBScript.RegExp`, and no Power Query automation.
- The Runner SHALL clone the Target EIB template before writing, never overwrite the template itself.
- The Runner SHALL flag unresolved reference lookups in the output rather than dropping rows silently.
- No sheet in the Runner workbook SHALL be hidden.
- The Runner workbook SHALL NOT hold run data — a run's only persistent artifact is the cloned output
  workbook.
- The module SHALL NOT require any software beyond Excel (with VBA enabled) on the user's machine.
- The module SHALL NOT share runtime code or configuration with `elt_ingest_excel`.

### Non-Functional

- Must run on Windows Excel (the finance users' environment) using only native VBA — no COM-only
  library, no macOS-only automation (e.g. AppleScript), in the shipped workbook itself.
- The restriction to native VBA only (no ADODB, no Power Query) means the deployed Runner can be
  built *and tested end-to-end* on a Mac dev machine, with no behavioural gap to Windows. This was
  confirmed empirically: ADODB cannot run at all on Mac Excel, and Power Query's VBA automation
  (creating a query and loading its results into a worksheet table) silently fails on Mac even though
  the Power Query engine itself works there — see Known Limitations.
- The deployed Runner is built solely by the project owner from a Mac dev environment and shipped
  to finance users as-is; there is no Windows build pipeline or CI step.
- VBA source is kept under version control as text files; the binary `.xlsm` is a build output, not
  hand-edited. Text `.bas` / `.cls` source files use ASCII encoding and CRLF line endings exclusively.
  Multi-byte UTF-8 characters (even inside comment blocks) and LF-only line endings are prohibited
  from new source: Mac Excel's `VBComponents.Import` silently aborts and falls back to a default
  `Module1` placeholder when they appear in volume, producing a non-obvious build failure. See
  `docs/ADDING_A_TRANSFORM.md` "Source-file hygiene" for the conversion list and verification commands.

## Acceptance Criteria

- [ ] Given a valid source file and the "Customer" transform selected, the module produces a cloned
      EIB workbook with customer rows written into the expected worksheet starting at the expected row.
- [ ] Given a source row whose customer/supplier cannot be resolved against the reference lookup, the
      output flags that row rather than omitting it or defaulting it silently.
- [ ] Given a source file missing an expected sheet, the module stops and reports the missing sheet
      before producing any output file.
- [ ] The shipped workbook runs end-to-end on Windows Excel with no installed dependency beyond Excel
      itself.

## Known Limitations

- Excel/VBA cannot practically handle very large source extracts (tens of thousands of rows) at the
  same performance as the DuckDB-based pipeline; this module targets finance-user-scale extracts.
- There is no declarative query language for transform steps (no SQL, no Power Query/M) — every
  lookup/mapping is hand-written VBA. This is more verbose than a declarative query and is a
  deliberate trade against the alternatives below, not an oversight.
- Adding a new (Source Type, EIB Population Type) combination requires adding its own VBA transform
  macro — there is no end-user authoring of new transform logic.

### Rejected alternatives (recorded so they are not re-attempted without new evidence)

- **SQL via VBA/ADO** (against an Interim workbook queried like a database, mirroring
  `elt_ingest_excel`'s DuckDB SQL). Rejected: `ADODB`/ACE OLEDB is a Windows-only COM technology with
  no equivalent in Mac Excel's VBA host — it does not run on Mac at all, not just "differently." Given
  this module writes into Workday master data, shipping a design that can never be tested by its
  builder before reaching a finance user's Windows machine was judged an unacceptable risk.
- **Power Query (M language)**. Considered because the engine itself is genuinely cross-platform —
  confirmed by hand: a manual UI-driven "Get Data → From File" load of a CSV on Mac Excel produced
  correct tabular results. Rejected anyway: automating query creation *and* loading results into a
  worksheet table from VBA was tested directly on Mac Excel and failed — the table stuck at "Getting
  Data..." with Refresh permanently disabled, because the required connection mechanism is
  Windows-specific plumbing. Using Power Query would also have meant every transform step needing a
  one-time manual UI setup (the exact manual-Excel-touch this project's deployment model is built to
  avoid), and would add a second, independently IT-policy-disablable feature on top of the VBA macros
  the project already requires.

## Related Modules

| Module | Relationship |
|--------|-------------|
| `elt_ingest_excel` | No dependency. Solves a related but distinct problem (technical-user-run, Python/DuckDB-based data migration) for a different user population. Do not share code or config between the two. |
