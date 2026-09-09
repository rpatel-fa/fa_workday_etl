# Backlog: Cold-Start-Hardened elt_transform_excel + Agent Scale

_This document is the SOLE continuity anchor for all hardening, scaling, and
inference-playbook work in this package. A coding agent dropped into this
package fresh (cold start, no hidden memory, laptop rebuild) should be able to
read ONLY: `README.md` → `docs/ARCHITECTURE.md` → `docs/ADDING_A_TRANSFORM.md`
→ `docs/BUILD.md` → `docs/todo/archive/cold-start-scale.md` (this file) → then execute
work without any interview questions._

## Resume (start here — fresh session)

### STEP 0 (MANDATORY FIRST ACTION ON EVERY SESSION): VBA hygiene pre-flight (G1 + G2).
**Never skip. Always run first, before any VBA edit, Runner shell create, or build_runner.py prep.**

```bash
cd elt_lake
python3 elt_transform_excel/scripts/check_vba_hygiene.py
# If any NON_ASCII WARN/ERROR raised (em-dashes, curly quotes, arrows — copy-paste / smart-quote
# contaminants — auto-fix all 26 common unicode chars in-place across all .bas:
python3 elt_transform_excel/scripts/check_vba_hygiene.py --fix
# Then full-contract hygiene against all 9 profile bindings:
uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project \
    python elt_transform_excel/scripts/build_runner.py --hygiene-only
```

**Rule enforcement (both scripts):** NON_ASCII bytes > 20 ERROR; any lone-LF ERROR; component/stem name > 31 chars ERROR; lines ≥ 300 chars ERROR (251–299 WARN); Worksheets(n) index access ERROR; ResolveSelectedTransformKey arity ≠ 2 ERROR; Public Sub Run<Name>(…) not taking 2 typed Workbook+Worksheet params ERROR; any `runtime.step_list_macro` without matching Public Sub ERROR.
Full rationale + each rule's history in:
`docs/AGENT_COLD_START_PROMPT_CONTRACT.md §0.5`.

If continuing from a prior session, the next unfinished actions are:

1.  **Doc mirrors (low priority, cosmetic — code already works).** The three
    rules below are already written in THIS backlog doc (Constraints § #8, #11,
    #0 self-description rule) and enforced at build time. For full package
    self-description, mirror each paragraph into `docs/ARCHITECTURE.md` as a
    short bullet under `§ Engineering Conventions`:
    - (a) Per-row-value generalization rule (H1.5b)
    - (b) H3 20-transforms / 4-hour governance gate
2.  **(Optional, judgement call) Shell creator cleanup.**
    `scripts/create_runner_shell.py` still seeds a legacy copy of the Transform
    Registry onto Runner columns G–T (REGISTRY_START_COLUMN=7, rows 2/3/4 =
    old title/header/data). This is harmless dead bytes — `build_runner.py`
    ALWAYS wipes Runner columns G+ immediately via
    `_clear_runner_tab_technical_columns(output_path)` before writing the real
    registry onto the dedicated `registry` VeryHidden sheet. Delete the
    legacy block if it confuses a new reader; otherwise leave.
3.  **(H3 future work, NOT to be started now — governance trigger only)**
    Per the H3 threshold rule written in this file, a table-driven runner
    engine does not start until BOTH: transform count ≥ 20 AND per-new
    authoring cost exceeds 4 hours. Current count: 9 transforms.

Resume prompt:
```
From docs/todo/archive/cold-start-scale.md: Continue "Cold start / scale hardening".
Code-level H1 + H2.1 gates are DONE and verified.
Pick one from "Still todo" priority ordered (all are doc mirrors, no code).
After docs, re-run the 4 main close-out gates from § "Verification commands"
(validate_input for touched profiles / verify_gates / gap_sweep / impl_detail_audit).
```

### Close-out gates (VERIFICATION COMMANDS, run before marking any session COMPLETE):

For each profile YAML you edited in this session (replace `<stem>`: `ap_amex_hotel`, `ap_amex_trains_v2`, etc.):

```bash
cd elt_lake
# Touch each touched profile:
uv run --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/validate_input.py -t <stem> | tail -6
# Full invariant cross-check:
uv run --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/verify_gates.py | tail -12
# transform_specs.yaml ↔ transforms.yaml literal fill audit:
uv run --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/gap_sweep.py | tail -8
# VBA literal strings ↔ YAML engine tokens drift check:
uv run --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/impl_detail_audit.py | tail -8
```

Expected all-green: validate_input "Full VERDICT: GREEN / All checks passed", verify_gates
"9/9 cross-checks: VERIFIED" (N+1 as count grows), gap_sweep "XXX FILLED / 0 GAP",
impl_detail_audit "0 STRING DRIFT / 0 NUMERIC DRIFT".

If touched profile has a `final_code_derivation` block (Trains/Hotels/Mapping-MASTER), run
the kept-rows full-file audit `/tmp/_final_auditor_correct_layers.py` before marking COMPLETE —
expect 100.00% exact char match vs Finance "Final Code" legacy source column (99.88% tolerated
only for case-only ERROR/Error cosmetic differences).
See §5 of AGENT_COLD_START_PROMPT_CONTRACT.md for PASS criteria.

---

## Backlog Item: Cold start / scale hardening

### Purpose

Eliminate the 8 traps a cold-start coding session WILL hit when asked
"convert source file A → target EIB B". Also move every "hidden" rule
currently only in assistant memory (`.trae/memory` project_memory.md, prior
chat summaries) INTO canonical docs under `elt_transform_excel/` so nothing
is lost on a laptop rebuild or a new agent instance.

### Package-level self-description contract (non-negotiable)

An agent dropped into `elt_transform_excel/` with NO outside memory must be
able to complete a "new transform" request by reading, in order:

0. `docs/AGENT_COLD_START_PROMPT_CONTRACT.md` → **read first if a user prompt was given.**
   Defines the 3 valid user prompt forms ("update transform for xyz", "create new transform
   for source abc → supplier EIB", "create new transform for source abc → customer EIB"),
   the ambiguity-rules table (when the agent MUST fire a clarifying question vs. can proceed
   silently), the 8 required user prompts when creating a profile from scratch, the
   zero-follow-up short-form syntax, and the mandatory 4-step close-out pipeline
   (build gate → reconcile → hygiene → impl-detail). If the prompt doesn't match a valid
   form → raise a clarifying question first, do not guess.
1. `README.md` → what the package does, 3-tier UX (config / Runner / registry).
2. `docs/ARCHITECTURE.md` → authoring model, judgment vs deterministic,
   native-VBA-only constraint, where domain knowledge lives (shared engine vs
   per-transform `.bas` vs spec YAML vs config YAML).
3. `docs/ADDING_A_TRANSFORM.md` → step-by-step playbook, 5 artifacts, 3
   exemplars (hotel simple / trains lookups / cfa account-string), mapping
   contract, Supplier Invoice EIB field block, account-string coding rules.
4. `docs/BUILD.md` → exact build command, Runner + Bootstrap semantics.
5. `docs/CONFIGURATION.md` → reference_paths, shell_defaults, mapping row
   vocabulary.
6. `docs/todo/archive/cold-start-scale.md` → THIS file: backlog, traps,
   inference-playbook section, gotchas, verification commands.

If an agent cannot do the job after reading those 7 files, a doc gap exists;
fix the docs, do NOT rely on telling the agent verbally.

---

### Done (already completed — verified)

- **Registry moved OFF Runner tab → VeryHidden sheet `registry`** (14 cols A..N,
  title/header/start rows 1/2/3). Runner tab columns A..F user zone only.
  Files: [build_runner.py](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/scripts/build_runner.py),
  `vba_src/modHelpers.bas`, `vba_src/modDispatcher.bas`, all 9
  `vba_src/transforms/*.bas` (ResolveSelectedTransformKey 2-arg signature update).
  Verified via inline hygiene script. 9 bindings parse incl `AP AMEX Trains V2`.
- **Project memory rules forwarded to THIS doc** (see "Accumulated Active
  Constraints" and "Gotchas" below) so `.trae/memory/…/project_memory.md` is
  optional extra context, not required. This is the single canonical anchor.
- **H1.1 — Build-time signature contract validation (trap #1).** Implemented in
  `_validate_vba_source_hygiene()` block #5 in
  [build_runner.py](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/scripts/build_runner.py).
  Three sub-checks: (a) every `ResolveSelectedTransformKey(…)` call-site has
  EXACTLY 2 untyped args `runnerWorkbook, controlSheet` (declaration lines
  excluded, CRLF-safe line counting via cumulative `cr_before[]` byte index map
  to fix off-by-one where `^\s*Public` matches `\r` at the end of the prior
  line); (b) every `Public Sub Run<Name>(…)` has the 2-typed-arg contract
  `ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet` with a
  deliberate carve-out: `RunSelectedTransform` (the button's OnAction macro)
  MUST have zero args because Excel form controls pass nothing; (c) every
  profile `runtime.step_list_macro` resolves to an actual Public Sub in some
  `.bas` file (duplicate / non-ASCII / empty step_list_macro all fail the
  build). VERIFIED: 9 profiles × signatures pass; module hygiene + signature
  gates combined exit clean.
- **H1.2 — Automated hygiene gates (traps #6 + #8).** Pre-existing static
  checks in `_validate_vba_source_hygiene()` blocks #1-#4. Check #4b also
  added: any `.bas` file containing a literal `Worksheets(<integer>)`
  OUTSIDE a VBA single-line comment line FAILS the build with a clear error
  directing the reader at the `RUNNER_SHEET_NAME` / `REGISTRY_SHEET_NAME`
  named-constant pattern. CRLF prefix counter bug-fix: bytearray(N) capped at
  255 entries → switched to `list[int]` for `cr_before[]`. VERIFIED: 14
  modules + 9 bindings; 14× ASCII scan all under ≤20 tolerated legacy bytes
  in modHelpers / modFinanceCommon; zero line-ending defects; 9 transform
  module stems all ≤ 31 chars, VB_Name matches stem.
- **H1.3 — Runner-sheet named lookup (trap #2).**
  (a) `Public Const RUNNER_SHEET_NAME As String = "Runner"` added to
  [modHelpers.bas](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/vba_src/modHelpers.bas)
  next to the `MAPPING_SHEET_NAME` / `REGISTRY_SHEET_NAME` constants;
  (b) Two refs in `vba_src/bootstrap/BootstrapImporter.bas` (FindOrCreateStatusLabel,
  EnsureRunButton) + one ref in `vba_src/modDispatcher.bas` (RunSelectedTransform)
  all migrated from `ThisWorkbook.Worksheets(1)` → `Worksheets(RUNNER_SHEET_NAME)`.
  (c) `create_runner_shell.py` already names sheet #1 at line 19 (`RUNNER_SHEET_NAME =
  "Runner"`) and writes `worksheet.title = RUNNER_SHEET_NAME` at line 163 during
  shell creation — no code change needed.
  (d) build_runner.py: added `RUNNER_SHEET_NAME = "Runner"` const (L25),
  `_read_named_sheet_cells()`/`_resolve_named_sheet_path()` siblings to the
  zip-level sheet-reader functions, `_verify_runner_sheet_invariants()`
  FAILS THE BUILD if "Runner" sheet is missing OR is not the first sheet in
  workbook.xml ordered list (tab drag reordering).
  (e) `_write_label_value_pair_runner()`/`_clear_runner_tab_technical_columns()`
  both upgraded from `workbook.sheetnames[RUNNER_SHEET_INDEX]` →
  `workbook[RUNNER_SHEET_NAME]` explicit named lookup.
- **H1.4 — CLI registry inspector (trap #5).** New file
  [inspect_registry.py](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/scripts/inspect_registry.py).
  Does NOT use openpyxl to read `sheet_state` (openpyxl doesn't reliably
  round-trip `veryHidden` for VBA workbooks) — reads directly from
  `xl/workbook.xml` `<sheets><sheet state="veryHidden">` attribute. Prints
  ordered sheet list with visibility marker `state=<visible|hidden|veryHidden>`
  next to each. Parses the 14-column registry block and pretty-prints the
  title/header/data rows. `--strict` exits non-zero: exit 2 = registry missing
  or not veryHidden; exit 3 = header mismatch, blank-row gap inside data
  rows, or any data row with empty "Step List Macro". VERIFIED: against a
  freshly-built Runner.xlsm → 4 sheets (Runner/config/mapping visible,
  registry veryHidden ✓, 9 transform rows shaped correctly ✓).
- **H1.5 — Config-tab overwrite banner + generalization rules.**
  (a) [build_runner.py](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/scripts/build_runner.py)
  `_write_shell_defaults_config_sheet()` writes a merged cols 1-10 row 1
  banner cell in bold white on `C0504D` red fill: *"AUTO-GENERATED — EDITS BELOW
  ARE OVERWRITTEN ON EVERY BUILD. Edit the source of truth in
  config/transforms/*.yaml, then re-run build_runner.py. Machine-local path
  edits here are intentional (Finance users tune these after each build)
  and only persist until the next build run."* Banner row height = 42px,
  wrapped. All subsequent config rows (title/header/data) shifted +1;
  `_verify_shell_defaults_config` checks banner text + shifted row offsets.
  (b) Per-row-value generalization rule written into Accumulated Active
  Constraints § #8 (and mirror into ARCHITECTURE.md still-todo).
- **H2.1 — Scaffold transform CLI seed.** New file
  [scaffold_transform.py](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/scripts/scaffold_transform.py).
  `--list` prints 9 exemplars with archetype labels (A/B/C MOST-COMMON /
  Trains V2 raw-derivation / Customer Invoice variants) + picker guidance
  matching the inference playbook § I. `--fork-from KEY --new-key KEY
  --new-label LABEL --eib-type "Supplier Invoice|Customer Invoice"
  [--source-type TYPE] [--dry-run] [--force]`. **Before writing ANY file**,
  validates: snake_case new-key (3-63 chars); VBA module stem character
  regex + 31-char HARD LIMIT; Sub name regex. Confirmed to reject a
  53-char stem with a clear "Shorten --new-label" message. Resolves the
  exemplar bas stem by scanning each module's Public Sub declarations for
  the profile's `runtime.step_list_macro` value (works correctly even when
  exemplar bas stems use abbreviations like `modApCfaDistribSupplierInvoice`
  instead of the key-name-inferred `*Distributions*`). Copies 3 artifacts:
  profile YAML, transform_spec YAML (stub if exemplar has none), `.bas`.
  Profile shell_defaults paths → `TO_FILL_*_PATH` sentinels so machine-local
  paths never leak across users; required_headers list replaced with 3
  sentinel comment lines pointing reader at the classification playbook.
  Transform_spec: TO_FILL_MAPPING_SPEC banner directs at § II. .bas: banner
  comment inserted after `Option Explicit` with SCAFFOLDED/Forked-from/4
  TO_FILL_* token names/Archetype guidance block. Non-ASCII tolerated
  chars (legacy emdashes etc.) replaced with ASCII equivalents; output
  always CRLF + pure-ASCII. Post-scaffold 5-step TO_FILL checklist printed
  to stdout ending with exact copy-paste commands for build + registry-inspect.
  VERIFIED end-to-end: real scaffold (key=ap_scaffold_test, module stem
  31 chars = on-boundary, step_macro=RunApScaffoldTstSupplierInvoice)
  generated → run `_validate_vba_source_hygiene` against 15 total modules
  (14 real + 1 scaffold output) → **FULL HYGIENE PASSES**. Generated files
  deleted afterwards so the repo carries no test residue.

### Still todo (next actions, all documentation-only mirrors — code already verified)

Ordered by value per effort; these are mirrors into `docs/ARCHITECTURE.md`
only so the 6-file self-description contract reads cleanly without the
agent having to open THIS backlog file. Code/build behavior is already
correct for every item below.

1.  **ARCHITECTURE.md § Engineering Conventions → add Per-row generalization rule.**
    Copy verbatim the paragraph from § Accumulated Active Constraints #8
    (this file): *"If it varies by source row, it goes into a `ref_*.json`
    row first; no hardcoding into VBA unless it's genuinely a
    transform-forever constant."* + rationale (zero-build-cost JSON row vs
    600-line per-module boilerplate, composes across transforms).
2.  **ARCHITECTURE.md § Authoring model → add H3 governance trigger paragraph.**
    Copy verbatim the paragraph from § Backlog Done item H1.5 (old Still
    todo item #7): *"Switch from per-transform .bas modules to a single
    modTableDrivenRunner IF AND ONLY IF transform count reaches 20 AND
    per-new-transform authoring cost exceeds 4 hours. Below 20, boilerplate
    copy is cheaper than engine maintenance. Engine is governance-reviewed,
    not an incremental edit."* (Note: this text already lives in
    `ADDING_A_TRANSFORM.md § Scaling note`; mirror to ARCHITECTURE.md so
    the reader can see every rule without opening all 6 docs.)
3.  **(Optional) Shell creator dead-byte cleanup.**
    `scripts/create_runner_shell.py` lines 23-27 still carry the legacy
    `REGISTRY_START_COLUMN=7 REGISTRY_END_COLUMN=20 … TITLE_ROW=2
    HEADER_ROW=3 START_ROW=4` constants and `_seed_transform_registry_block(
    worksheet, transform_bindings)` call writes onto Runner.G..T. This is
    never user-visible because `_clear_runner_tab_technical_columns` in
    build_runner.py always wipes Runner columns G+ immediately after
    copying the shell. Delete these two constants + the function call if
    a new reader finds them confusing; otherwise leave.
4.  **(Do NOT start yet) H3 table-driven engine.** Governance gated; no work
    until 20 transforms AND ≥4 hours / new transform cost. Current count: 9.
    If/when the threshold is hit, open THIS backlog as the starting point
    for the engine design doc.

---

### Inference playbook: How a cold-start coding agent should map source X → EIB Y

_This section is the "canonical brain" an agent reads when told "convert this
source file into a target EIB". It is the written-down version of the
inference a senior dev would do when reading the existing 9 configs. Every
step here must be executable from the canonical docs alone._

**Input prompt shape the agent expects:**
```
convert:
  source_file: <path to raw source.xlsx (or csv)>
  target_eib_type: Supplier Invoice | Customer Invoice
  exemplar_hint: (optional) hotel / trains / cfa_distributions — or agent picks
```

**Step-by-step inference:**

#### I. Identify the closest EXISTING exemplar (most important step, 80% of work)

Open `config/transforms/` and list the 9 profiles + spec keys:
- `ap_amex_hotel` → Simple case: spend_cat + cost_centre FIXED constants,
  per-row value lookup only for VAT. START HERE if 90% of EIB values are
  constants for every row.
- `ap_amex_trains` → Rich per-row: cost centre varies by row (from source
  column, live lookup into mapping workbook), VAT varies per-row by source
  column. START HERE if coding rows differ by source-row column values
  (not a coding string).
- `ap_amex_trains_v2` → Raw+derivation: source carries NO "Final Code";
  derive it from a raw string via a multi-step derivation pipeline (see
  spec's `meta.derivations:` block). START HERE if any source column needs
  regex/split/lookup before it can drive a worktag.
- `ap_cfa_distributions` → Account-string case: source carries a coding
  string (`1077-504007-PNGC0001-0000` or similar) with segments that map to
  Spend Category / Cost Centre / Reporting Type / worktags. START HERE for
  the common AP supplier-invoice pattern where a finance GL code is
  provided in a source column.
- `ap_prize_funds` / `ap_referees` → Variants; use only if the
  source exactly matches prize/referee patterns (rare).
- `ar_invoice` / `ar_invoice_run` / `ar_discipline_invoice` → Customer
  Invoice EIB vertical; START HERE for AR.

RULE: Copy the YAML + .bas FROM THE EXEMPLAR THAT HAS THE LARGEST SHARED
STRUCTURE, not the smallest delta of label/names. Account-string feeds
almost always fork from `ap_cfa_distributions`. Derivation feeds fork from
`ap_amex_trains_v2` because the raw→cleaned derivation step adds a whole
new function category to the module skeleton.

#### I.a. Load non-inferred shared conventions (DO THIS BEFORE ANY FORKING — never infer)

Before writing any profile YAML or module code, load these two config files
directly. Every constant on this page is single-source-of-truth from YAML,
never inferred from a `.bas` module. Failure to load these here will fail
the `build_runner.py` schema validation LONG before any workbook write or
Excel import step.

1. **Per-EIB-population-type conventions.** Read
   `config/eib_population_defaults.yaml` keyed by the exact
   `transform.eib_population_type` string. Carries three non-inferred
   constants a cold agent otherwise silently gets wrong:
   - `target_sheet_name` (e.g. "Submit Supplier Invoice")
   - `target_data_start_row` (always 6 in the shipped two EIBs today)
   - `clear_published_rows_end_col` (Supplier Invoice = EX, Customer Invoice = MT)
   (If the new EIB type isn't listed, ADD A NEW TOP-LEVEL KEY to
   `eib_population_defaults.yaml` BEFORE proceeding — don't patch the
   module.)
2. **Per-transform source_layout conventions.** Every profile's
   `source_layout:` block carries four enum/int values that tell the VBA
   module how to open a source file. NEVER derive these from the source
   data by guessing; set them explicitly:
   - `worksheet_resolution: named | first_sheet` — named ⇔ `worksheet_name`
     non-empty string required; first_sheet ⇔ `worksheet_name` MUST be
     omitted (the module uses vbNullString pattern to read sheet#1).
   - `header_detection: literal_row | scan_required_headers` —
     literal_row = header is always row 1 (no scan, use
     `FindColumnIndexInRow(sourceTable, 1, …)`); scan_required_headers =
     source file may have title rows above the real header, call
     `FindHeaderRow(sourceTable, requiredHeaders)` then use the offset
     `dataStartRowIndex = headerRowIndex + (data_start_row - header_row)`.
   - `header_row`, `data_start_row` positive integers. For
     `scan_required_headers` mode, these are offsets relative to whatever
     FindHeaderRow finds; for `literal_row` mode, header_row is always 1.
   - `required_headers: list[str]` — non-empty; both detection modes
     verify.
   Build cross-field validation checks these rules in
   `_validate_source_layout_schema()` (see build_runner.py). Failure
   at profile-parse time with a clear message, not at Excel runtime.

#### II. For EVERY column in the source, classify it (do this BEFORE writing)

Read source header rows (step 1 in ADDING_A_TRANSFORM.md). For each column,
write a one-line classification (this becomes the spec's `fields:` section):

| Class | How it becomes a mapping row |
|---|---|
| **Pass-through** (e.g. `Item Description`) | `target: di`, `source_column: "Description"` direct |
| **Format-translate** (e.g. `Invoice Date` 01/04/2026 → 2026-04-01) | `source_column:`, `source_format: DD/MM/YYYY`, `target_format: YYYY-MM-DD`; in VBA call `NormalizeSourceDate` |
| **Per-row live lookup** (e.g. supplier name → Supplier ID via supplier spreadsheet) | `source_column:`, `lookup.file: Supplier Spreadsheet`, `lookup.tab:`, `lookup.input_column: Name`, `lookup.output_column: Supplier_ID` |
| **Account-string lookup** (segment of coding string maps to worktag ref id) | `source_column:` (the coding col), `function: lookup_worktag_from_final_code`, lookup block filtered by `WORKDAY BUSINESS OBJECT` + `Core vs Prism=Core` (see cfa_distributions spec) |
| **Raw derivation** (e.g. Ref 2 → Final Code — NOT a simple lookup) | Create a spec-level `derivations:` block. In VBA, add one private function `Derive<Name>From<SourceCol>` at the TOP of the module (before WriteOutputRow). Call before writing dependent fields |
| **Fixed constant** (Submit=Y, Currency=GBP) | `default: <value>` only. Confirm with finance before hardcoding, OR use ref_default.json when the value varies by tenant |
| **Cross-sheet within source workbook** (e.g. mapping tab inside source) | `lookup.file: Source Data`, `lookup.tab: <sheet name>` — helper reads same workbook as source |
| **Tenant workday IDs** (Tax_Applicability_ID / Spend_Category_ID — NEVER in source files) | Look up BY NAME (e.g. "Standard", "20%") in ref_tax_code.json / the Mapping spreadsheet. NEVER guess a Workday REF ID. If not in refs, ASK — this is the single silent-miscoding risk |

##### II.a Companion rules — small conventions that silently break correctness if forgotten

These are NOT column-classification items; apply them unconditionally when the
classification pattern matches. Every single one of these was a trap where a
cold agent would write something syntactically valid that silently produces a
wrong EIB file at Workday-import time. Forwarded from the 6-item C-gap review
of ap_amex_hotel + ar_invoice_run trios.

1.  **C1 — TARGET sheet names/clear-end-col come ONLY from eib_population_defaults.yaml.**
    Never embed `TARGET_SHEET_NAME = "Submit Supplier Invoice"` as a const in the
    `.bas` without checking `eib_population_defaults.yaml` first. If a new EIB
    version changes the sheet name or end-col, you change it ONCE in
    eib_population_defaults.yaml, not per-module.
2.  **C2 — YYYY-MM-DD date writes MUST set `.NumberFormat = "@"`.**
    When Class II row says `target_format: YYYY-MM-DD` (Format-translate date
    via `NormalizeSourceDate`), the VBA writer MUST execute two consecutive
    lines per write:
    ```
    targetSheet.Cells(targetRow, ColumnIndex("U")).NumberFormat = "@"
    targetSheet.Cells(targetRow, ColumnIndex("U")).Value = invoiceDate
    ```
    The target template's cell defaults to General format; Excel silently
    reformats a yyyy-mm-dd pure-string write as system short date (e.g.
    `01/04/2026`) which Workday treats as a valid but wrong day/month
    inversion — the worst trap.
3.  **C3 — Shared Engine Helper Surface (call these, don't re-implement).**
    The modHelpers / modSourceTable / modFinanceCommon modules ALREADY export
    these. If you write something that does the same job with a private helper
    in a per-transform module, you're duplicating behaviour (and build gates
    will never catch the mismatch because the private code is self-consistent).
    Grouped:
    - Workday ref-id helpers: BuildReferenceLookupFromWorkbook,
      BuildBusinessObjectHierarchyLookup,
      ResolveLowestHierarchyLookupAcrossFinalCodeSegments, ResolveMappedCompany,
      ResolveReportingTypeFromFinalCode.
    - Date/amount/column helpers: NormalizeSourceDate, ResolveAmount,
      FindHeaderRow, FindColumnIndexInRow, ValidateHeaderRow,
      ResolveSourceWorksheet, ReadWorksheetTable, OpenWorkbookReadOnly,
      CloneTemplateWorkbook, ClearPublishedRows, ColumnIndex, EnsureNonBlank.
    - Runner lifecycle helpers: BeginRun / EndRun, UpdateRunProgress,
      AppendRunLog, AppendRejectRow (plus the 14 RESULT_LABEL_* label-name
      constants the Runner writes to).
    - Dispatcher helpers: ResolveSelectedTransformKey(runnerWorkbook, controlSheet)
      2-arg form; the `Public Sub Run<Name>` entry point contract, the
      `RunSelectedTransform` button macro zero-arg entry point
      (button-dispatcher only — never add args).
4.  **C4 — Company mapping row `lookup.file: this tab: mapping filter_column=entity filter_value=company` pattern → ResolveMappedCompany ONCE before row loop.**
    That specific YAML row means "one-shot per-transform tenant-company
    resolution". Implement it by calling the shared singleton
    `ResolveMappedCompany(runnerWorkbook, sourceCompanyCode, defaultFallbackText)`
    exactly once before the per-row loop, then re-use the same resolved
    company-id for every row. Do NOT implement this as a per-row mapping
    lookup — it's invariant across rows (the shared helper caches it
    internally).
5.  **C5 — Mapping-row special case `default: "NULL"` → write empty string `""`, NOT the literal word "NULL".**
    The Workday loader expects an empty string for certain columns in
    specific EIB layouts (e.g. Customer Invoice KE = Bill From, which is
    left blank when there's no Bill From distinction). The spec writer uses
    the quoted sentinel `"NULL"` (capitalised, quoted) to flag "write empty
    string here". The VBA side must write `targetSheet.Cells(targetRow, ColumnIndex("KE")).Value = ""`
    NOT `.Value = "NULL"`.
6.  **C6 — Header detection pattern follows source_layout.header_detection, not the module copy-paste default.**
    Forking an exemplar always copies `FindHeaderRow(sourceTable, requiredHeaders)`.
    If the new profile has `header_detection: literal_row`, DELETE the
    FindHeaderRow call and replace the FindColumnIndexInRow 2nd arg with
    literal `1` (header is row 1, no title rows above). Keep the
    FindHeaderRow pattern ONLY when header_detection = `scan_required_headers`.
    If you mix them, the module either errors on a perfectly-formed row-1
    source, or silently skips N title rows in a file that actually has them.

#### III. Canonical 6 non-configurable prompts — REQUIRE Finance/user confirmation

These are the only items that are NOT recoverable by reading any config file.
Before considering any new transform done, the agent MUST explicitly list each
of these 6 items to the user/Finance and get written confirmation. Enumerated
from the 2026-08-11 review of ap_amex_hotel + ar_invoice_run trios:

1.  **Source worksheet resolution confirmation.** Inherited `worksheet_resolution`
    enum value from the exemplar + `worksheet_name` string. Confirm: *"Is the
    real source tab name really `<name>` / really the first sheet?"*
2.  **Company code + mapping-tab fallback text.** Hardcoded tenant company code
    (e.g. `FA` or `WNSL`) plus the "THE FA GROUP - FA"-style fallback text that
    ResolveMappedCompany uses on a mapping miss. Never guess; this has silent
    wrong-coding consequences at Workday import time.
3.  **All fixed Workday tenant Ref-IDs that are NOT derived from source.**
    List every constant target-column Ref-ID you plan to write: Tax Applicability
    (e.g. GBR_Outside_of_Scope), Tax Codes (GBR_Zero_Rated / GBR_200_VAT), Tax
    Recoverability (Nonrecoverable_Allocate), Spend Category names, Currency,
    Document Type, Submit Y/N, all fixed Worktag IDs. Rule: Class 8 Tenant
    Workday IDs = NEVER guess; get sign-off.
4.  **Unused source columns explicit confirm.** After the classification table
    is done, list every source header NOT yet assigned to any of the 8 classes.
    Confirm each: *"This column is intentionally NOT mapped — correct?"* A naive
    agent would otherwise assign "Cost Code" → "Cost Centre" and produce a
    silent wrong-coding EIB file that only Finance reconciliation catches.
5.  **Required-headers list confirmation.** Replace the TO_FILL_HEADER_*
    sentinel lines with the real headers from the source file. Confirm: *"Does
    every header below exist verbatim in the source file? FindHeaderRow errors
    hard if any header is missing."* (Applies only if header_detection =
    scan_required_headers.)
6.  **Per-transform tax mode.** Confirm: uniform (one tax code applies for
    every row; e.g. zero-tax Outside of Scope AP) OR per_row with an explicit
    driver column + every when_value → tax-code Ref-ID pair. If per-row,
    confirm the driver column name + all Ref-IDs (Class 8).

All 6 prompts are also emitted by `scaffold_transform.py` after forking an
exemplar (see the "6. Non-configurable tenant / user prompts" section in its
stdout output). If running the scaffolder, use the emitted list directly; do
not re-derive.

#### IV. Populate the 5 artifacts mechanically (after classification + prompts confirmed)

After the classification table is done, EVERYTHING below is deterministic and
can be done by the agent without further human input:

1. Write `config/transform_specs/<new-key>.yaml` — copy closest exemplar's
   structure. Keep `open_questions: []` empty; any missing decision was a
   classification gap. Back-fill from II before writing code.
2. Write `config/transforms/<new-key>.yaml` — copy exemplar profile. Set
   only `key`, `label`, `transform.source_type`,
   `transform.eib_population_type`, `runtime.step_list_macro` (keep pattern
   `Run<CamelKey><Suffix>`), `source_layout` (header rows, required_headers
   list, data start row), `shell_defaults` paths for source/target/output.
   Inherit everything else from `reference_paths.yaml` — do NOT copy global
   paths into every profile (config purity rule).
3. Copy exemplar `.bas` → new module stem. **If 31-char limit would
   breach, short the stem NOW; use the 31-char gate from H1.2 to fail.**
4. The `.bas` deltas after copy are exactly (and only):
   a. Header constant block — replace each source-column NAME with the
      actual names from source file classification (these are the `TO_FILL_*`
      tokens the scaffolder will mark).
   b. `FindColumnIndexInRow` list — keep the exemplar's structure; just
      rename values.
   c. Private derivation helper functions — ADD new for any Class II "Raw
      derivation" cols. Otherwise delete exemplar helpers that don't apply.
   d. Per-row lookup builds — `Build<BusinessObject>Lookup` functions;
      copy cfa_distributions pattern exactly for account-string lookups,
      trains pattern for simple source-col lookups.
   e. `WriteOutputRow` — one cell per target EIB column, following the
      mapping in spec. This is the audit-traced "code mirrors YAML" step.
5. Build per BUILD.md, import via BootstrapImporter, verify via
   `inspect_registry.py`, run via Runner, check Run Status + Rows Processed,
   and confirm a sample row's EIB cells match the exemplar's valid output
   pattern from the target EIB type column block (Supplier Invoice = cols
   B..EX from ADDING_A_TRANSFORM.md).

#### V. Amex Trains V2 special case — raw→cleaned derivation BEFORE EIB mapping

For ANY transform where Finance historically did a manual "Final Code" step
(meaning the raw source file DOESN'T contain a worktag-driving coding
column, but a FUTURE raw file WILL contain the unclean column), the
inference path forks from standard to:

1. FIRST — write a `meta.derivations:` block IN THE SPEC before mapping.
   Each derivation has `from_ref`, `steps: []`, a 6-step pipeline like
   trains v2. Document any lookup tables (Seg2 first char → accounting
   code 4-entry mapping) DIRECTLY in the spec YAML so a future reader can
   verify parity without opening the VBA.
2. Implement derivation as private VBA helpers in the module, ONE helper
   per logical step (even if called in a chain). Name them the same as the
   spec's step names so a code-review diff can verify 1:1 correspondence.
3. Call them in the per-row loop AFTER header resolution but BEFORE any
   worktag lookups (you can't look up Cost Centre on a coding string that
   doesn't exist yet).
4. Map the DERIVED `Final Code` (or whatever name) as the
   `key_source_column` for account-string lookups — NOT the original raw
   `Ref 2`.
5. Verify parity with Finance's historical "final code" column values from
   ONE real file — 100 rows minimum before the transform is marked done.
   This is the ONE place where the human (finance) provides the
   ground-truth answer key; the agent verifies the implementation matches
   the key.

---

### Accumulated Active Constraints (forwarded — EVERY new item inherits these)

0. **Package self-description rule.** Any rule, gotcha, or invariant that
   an agent would need to author a new transform correctly MUST be written
   in one of: `README.md` / `docs/ARCHITECTURE.md` /
   `docs/ADDING_A_TRANSFORM.md` / `docs/BUILD.md` /
   `docs/CONFIGURATION.md` / this file. If it only exists in `.trae/memory`
   or chat summaries → forward it here before considering any new item
   done.
1. **31-char hard limit.** VBA component names (`mod<Name>` stems +
   `Attribute VB_Name`) ≤ 31 characters. `VBComponents.Import` silently
   drops the name → blank `Module1` → import rename fails → runtime macro
   not found (error `-2147219692 / 80040714`). `Run<Name>` Sub names are NOT
   subject to this limit; shorten only the module/stem/VB_Name.
2. **CRLF line endings ONLY for .bas files.** LF-only files break the
   `Attribute VB_Name` parse on Mac Excel VBA importer; pure-LF files
   import with default name and then fail constraint (1).
3. **ASCII-only .bas content.** Non-ASCII codepoints (curly quotes,
   emdashes `—`, arrows `→`, box-drawing, U+00A0 non-breaking spaces) in
   comments OR code cause mid-parse failures even when ASCII-equivalent
   alternatives exist. Tolerance threshold: ≤ ~20 bytes total; ≥ ~hundreds
   fails deterministically.
4. **Registry sheet is VeryHidden and NAMED `registry`.** No technical
   routing data on the user-facing Runner sheet. Runner cols A–F only.
5. **Runner is NAMED `Runner` (not indexed by `Worksheets(1)`).** After
   H1.3 is done, any `Worksheets(1)` remaining in code is a bug.
6. **Native VBA ONLY.** No `ADODB`, `Scripting.Dictionary`,
   `Scripting.FileSystemObject`, `VBScript.RegExp`, no COM-only type libs.
   Code restricted to plain VBA + Excel object model so it builds and
   tests on Mac Excel identically to finance-user Windows laptops.
7. **Domain-free shared engine.** modDispatcher, modHelpers,
   modSourceTable, modFinanceCommon (shared helpers only) must NEVER
   contain a column name, sheet name, or business value specific to ONE
   transform. Per-transform values → the per-transform `.bas`; reusable
   cross-transform-but-finance-domain helpers → modFinanceCommon;
   reference values that change per row → `ref_*.json` under
   `config/ingest_lookups/`.
8. **Per-row-value generalization rule** (see H1.5). If it varies by source
   row, it goes into a `ref_*.json` row first; no hardcoding into VBA
   unless it's genuinely a transform-forever constant.
9. **Ref 2 → Final Code derivation contract** (for trains v2 and future
   raw-pipeline transforms):
   - Strip `-`, space, `.` from Ref 2.
   - Right-pad to 22 chars with `0`.
   - Split 4|6|8|4 segments.
   - Seg2 first char lookup: `3→302101, 4→402201, 5→502201, 6→602201`.
   - Join `Seg1-Seg2Mapped-Seg3-Seg4` = Final Code.
   - (Documented here so future similar pipelines are spotted. This rule
     trains the inference pattern; it is NOT copy-pasted into modules that
     don't need it.)
10. **UX layering invariant for Finance handover** — the 3 visible sheets:
    `config` (user-edits shell defaults here), `Runner` (dropdown + run
    button + status), `mapping` (readonly if shipped), plus VeryHidden
    `registry` + future `transform_specs` companion. No 4th user sheet.
    Finance users never edit anything outside the `config` tab and the
    `Runner` dropdown.
11. **H3 trigger (table-driven runner engine)** — see § "Still todo"
    step 7. Governance-reviewed only after 20+ transforms AND per-new cost
    over 4 hours. No incremental engine builds.
12. **Single-source-of-truth principle for constants — NO INFERENCE FROM .bas.**
    - Per-EIB-population-type constants (target_sheet_name, target_data_start_row,
      clear_published_rows_end_col) live ONLY in `config/eib_population_defaults.yaml`
      keyed by the exact eib_population_type string. Never hardcode these
      values from memory or by reading a module's Private Consts.
    - Per-transform source-layout enum values `worksheet_resolution` (`named`
      / `first_sheet`), `header_detection` (`literal_row` / `scan_required_headers`),
      integer `header_row` / `data_start_row`, and the `required_headers` list
      live ONLY in the profile's `source_layout:` block. Cross-field
      validation is in `build_runner.py` `_validate_source_layout_schema()`.
    - Any constant that varies per transform or per EIB version AND was
      previously inferred from a .bas module MUST be moved to a YAML config
      entry and the .bas reference the config's value — no bidirectional
      inference allowed.
13. **Companion rules C1-C6 from II.a MUST be followed literally.** The six
    items (NumberFormat=@ for date writes; default:"NULL"→empty string;
    ResolveMappedCompany single-call before row-loop; Shared Helper Surface
    use-not-reimplement; eib_population_defaults single-source;
    header_detection pattern match) apply to every new transform regardless
    of exemplar fork. They are not exemplar-specific; they are universal
    package-wide rules. Violation does NOT produce a VBA syntax error; it
    produces a correct-looking EIB file that silently mis-codes in Workday
    — the worst kind of failure.

### Verification commands (exact copy-paste from this doc)

1. **Hygiene-only local preflight (no build, no Excel):** Runs all H1.1/H1.2/H1.3
   gates (signature contract, ASCII/CRLF/31char, Worksheet-index-literal ban,
   step_list_macro existence) against the source tree.
   ```bash
   cd elt_transform_excel
   uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project \
     python -c '
   import sys, types
   sys.path.insert(0, "scripts")
   br_path = "scripts/build_runner.py"
   br = types.ModuleType("build_runner")
   br.__file__ = br_path
   sys.modules["build_runner"] = br
   with open(br_path) as f:
       exec(compile(f.read(), br_path, "exec"), br.__dict__)
   vba_source = br.MODULE_ROOT / "vba_src"
   modules = br._discover_importable_modules(vba_source)
   bindings = br._load_transform_bindings(br.MODULE_ROOT / "config" / "transforms")
   print(f"modules = {len(modules)}, bindings = {len(bindings)}")
   br._validate_vba_source_hygiene(vba_source, modules, bindings)  # raises on fail
   print("OK: full H1 hygiene passes")
   '
   ```

2. **Scaffold exemplar list (confirms 9 archetypes still load):**
   ```bash
   cd elt_transform_excel
   uv run --with pyyaml --no-project \
     python scripts/scaffold_transform.py --list
   ```

3. **Scaffold dry-run for a hypothetical new transform (no file writes):**
   ```bash
   cd elt_transform_excel
   uv run --with pyyaml --no-project \
     python scripts/scaffold_transform.py \
       --fork-from ap_cfa_distributions \
       --new-key ap_bacardi_uk \
       --new-label "AP Bacardi UK" \
       --eib-type "Supplier Invoice" \
       --source-type "Bacardi UK GL Raw" \
       --dry-run
   ```

4. **Registry inspector against the current dist/Runner.xlsm:**
   ```bash
   cd elt_transform_excel
   uv run --with openpyxl --no-project \
     python scripts/inspect_registry.py --strict dist/Runner.xlsm
   ```

5. **Full build (with manual VBA import step):** Follow BUILD.md. After
   bootstrap import completes, open Runner.xlsm manually → confirm:
   - Tab bar shows ONLY `config`, `Runner`, `mapping`. No `registry` tab
     (it's VeryHidden).
   - Runner cols G–T are empty/blank — only A–F populated.
   - Transform dropdown lists all bindings (9 incl AP AMEX Trains V2).
   - Select any, press Run Transform → macro resolved from registry, no
     "macro not found".

### Gotchas (environment/tooling — forwarded from prior sessions)

1. **Python 3.14 dataclass + exec_module edge case.** When importing
   `build_runner.py` programmatically via `importlib.util.spec_from_file_location`,
   the dataclass decorator raises `AttributeError: 'NoneType' object has no
   attribute '__dict__'`. Workaround: register the module in
   `sys.modules['build_runner']` BEFORE running `exec()`. Seen in hygiene
   tests only; the user-facing build command via BUILD.md is unaffected.
2. **Shell-escaping of inline Python heredocs in zsh sandbox.** When
   passing multi-line Python via `python3 -c "…"` with any string
   delimiters or dollar signs, strings get mangled (duplicated text,
   unmatched quotes). Workaround: always write to `/tmp/<name>.py` first
   and execute the file directly.
3. **DuckDB lock contention with DBeaver.** Not transform-package-specific,
   but the dev machine commonly has DBeaver open; any DuckDB write to the
   shared db from a concurrent `elt_ingest_excel` session blocks. Close
   the DBeaver connection first. (Forwarded from ai_context/project/ASSISTANTS.md
   so it's visible here even if that file is not opened by a new reader.)
4. **openpyxl sheet_state spelling.** `"veryHidden"` camelCase matches the
   openpyxl 3.1+ enum; legacy openpyxl raises AttributeError on
   `sheet.sheet_state = "veryHidden"` — the build helper
   `_set_sheet_very_hidden` swallows this gracefully but logs a WARNING.
   If VeryHidden silently fails, upgrade openpyxl via `uv add openpyxl@latest`.
5. **`uv run --package elt-ingest-excel` does NOT work.**
   `elt_ingest_excel` is not a `uv` workspace member today. Run from that
   package dir with `cd elt_ingest_excel; uv run python …`.
6. **Runner build bootstrap writes APPEND status into col A at the bottom
   of the sheet.** FindOrCreateStatusLabel appends label/value pairs in
   A/B below user controls. This is intentional (build audit trail); they
   never trespass on columns G–T (now fully cleared by
   `_clear_runner_tab_technical_columns`). Do NOT mistake these for
   technical routing data — they're build/import metadata, safe for
   Finance users to scroll past.
7. **`reference_paths.yaml` paths are relative to the Runner workbook
   directory on the finance user machine.** Not relative to repo root —
   when setting shell_defaults for a new transform, use paths as they
   appear on the user's laptop (commonly `C:\Workday\References\…` for
   Windows), NOT the dev's local macOS path.
