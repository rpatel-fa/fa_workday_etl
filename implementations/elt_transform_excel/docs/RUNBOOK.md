# Runbook

## Purpose

This runbook describes how the project owner prepares, builds, runs, diagnoses, and recovers `elt_transform_excel`. It reflects the implementation state currently present in the repo: the shared VBA engine, shell generator, and build script exist; the `AR Invoice -> Customer Invoice` vertical slice is confirmed working end to end in real Excel; a second vertical slice, `Discipline Invoice -> Customer Invoice`, is code-complete but not yet verified in real Excel. The earlier Apple Event privilege violation that blocked headless automation from this CLI host is resolved (see Gotchas), but the module-import step remains a deliberate manual step in Excel regardless — see "Current Workflow" below.

## Ownership

- Build and deployment are Mac-only and are performed by the project owner.
- Finance users receive a built Runner workbook and do not run repo tooling.
- The shipped runtime remains Windows Excel with VBA enabled.

## Prerequisites

- Excel for Windows for end-user runtime validation before release.
- Python build dependencies available via `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project ...`.
- Representative source, target EIB, and reference workbooks for the selected vertical slice.
- macOS dev host: Microsoft Excel must have **Full Disk Access** (System Settings → Privacy & Security →
  Full Disk Access) so module imports don't block on a per-file "Grant Access" sandbox prompt, and
  Terminal (or whichever app drives `osascript`) must have **Automation** access to Microsoft Excel
  (System Settings → Privacy & Security → Automation) for any headless macro calls. A new grant may
  require fully quitting and relaunching Excel before it takes effect.

## Current Workflow

1. Keep all implementation work in text files under `vba_src/` and Python build tooling under `build/`.
2. Generate the minimal Runner Shell workbook with `scripts/create_runner_shell.py`.
3. Use `scripts/build_runner.py` (no flag) to copy the Runner Shell to `dist/Runner.xlsm`, write the
   visible transform-registry block from `config/transforms/*.yaml`, and write the source folder into a
   `Bootstrap Source Folder` cell. This step is pure Python and does **not** import any `.bas`/`.cls`
   modules — see Gotchas for why.
4. **Manual step (required):** open `dist/Runner.xlsm` in Excel, Tools > Macro > Macros, run
   `BootstrapImporter.BootstrapImportFromConfiguredSource`. Grant any per-file "Grant Access" prompts
   that appear. The macro imports every module, creates the **Run Transform** button on the
   Runner sheet, and saves the workbook itself.
5. Run `scripts/build_runner.py --verify` to confirm the Bootstrap Import status cells, transform
   registry, and module set all match expectations. Check the `Run Button` status cell on the Runner
   sheet too: it reads `Created -> modDispatcher.RunSelectedTransform`, or explains why it is absent.
6. Open `dist/Runner.xlsm` in Excel, pick a Transform, and click **Run Transform** for the
   first supported `(Source Type, EIB Population Type)` vertical slice against representative sample
   files.

### The Run button

Finance users run transforms by picking a Transform and clicking the **Run Transform**
button on the Runner sheet — they are not expected to open the Macros dialog. The button is a **Form
Control** (never ActiveX: ActiveX controls do not work in Mac Excel, so an ActiveX button would be
invisible on a developer Mac and appear only for Windows users) anchored at `D8`, bound to
`modDispatcher.RunSelectedTransform`.

It is created by `BootstrapImporter`, not by `create_runner_shell.py`, because its `OnAction` binds
to a macro that does not exist until the modules are imported, and because `build_runner.py`
round-trips the workbook through openpyxl after the shell is generated. This means the button does
**not** exist in the committed `runner_shell/RunnerShell.xlsm`, and it is re-created (delete-then-add,
so duplicates cannot stack) on every bootstrap. A rebuild therefore drops the button until step 4 is
run again — which is expected, since a rebuild also drops the modules behind it. If button creation
fails it never aborts the import; the `Run Button` status cell records why, and the Macros dialog
still works as a fallback.

## Reference Workbooks: Mapping vs. Lookups

The per-transform **field mapping** (which source column feeds which target EIB column) is authored as
the `mapping:` section of `config/transform_specs/<key>.yaml` — plain, diff-reviewable text that the
transform's `.bas` module mirrors. It is not a workbook and is not read by the build. (This replaced
the former per-transform tabs in the retired `RunnerShellMapping.xlsx` and the MD5
`mapping_content_hash` build gate; see the change record in `docs/todo/archive/build-elt_transform_excel.md`.)
Because the mapping and the `.bas` change together in the same text diff, drift is caught in code
review rather than by a build-time hash.

The one reference artifact that remains a live lookup is the `mapping` tab (entity/source/target),
which ships inside the deployed workbook:

| Artifact | Lives in | Answers | Checked |
|---|---|---|---|
| `mapping` tab | `runner_shell/RunnerShell.xlsm` (ships with `dist/Runner.xlsm`) | "What does this reference code resolve to right now?" — small lookup tables (e.g. company entity codes) that can change independently of the field mapping. | At **run time**, every run, via `modHelpers.LookupMappingValue`. |

The `mapping` tab is a **build output**, not hand-authored data. `scripts/build_runner.py` regenerates
it on every build from `config/ingest_lookups/ref_company.json` (source→target company names) and
`ref_default.json` (per-domain defaults, e.g. currency→GBP), folding each into `entity`/`source`/`target`
rows — the same self-describing reference-table mechanism as every other `ingest_lookups` dataset. To
change a company name or default, edit the ref table and rebuild; do not edit the workbook tab directly.
(This replaced the retired `RunnerShellMapping.xlsx` documentation mirror.)

### Mapping Lookup Workflow (run-time, `mapping` tab)

The `mapping` tab (entity/source/target columns) stays inside `runner_shell/RunnerShell.xlsm` and is
read live, every run, by `modHelpers.LookupMappingValue(runnerWorkbook, entityName, sourceValue)`.
All transform modules call it for company-entity resolution instead of hardcoding entity names.
(The AMEX transforms' Cost Centre / related-worktag lookups are separate live lookups — built per run
from the finance mapping master workbook named by the `Mapping Spreadsheet` + `Mapping Sheet`
controls, not from this `mapping` tab.)
There is no timestamp/drift check for this tab — editing it takes effect on the very next run with
no rebuild required, which is the point: it holds reference data that can change independently of
any transform's code.

## Watching a Run (progress) and the Run Log

A transform runs on Excel's single UI thread, so the workbook **will** look frozen ("Not
Responding") while it works — this is expected, not a crash. Three signals show it is alive:

1. **Run Status cell** shows `Running` (set at the start, before the heavy work) and then
   `Completed`/`Failed`. The status bar (bottom of the Excel window) shows `... row 3,400 of 8,900
   (12s elapsed)`, updated every 200 rows. When a run finishes, the result cells are **colour-coded**
   so mapping errors are obvious at a glance: **green** = every row wrote; **amber** = rows were
   rejected (`Run Status` + `Rows Skipped` + `Issues` shaded — check `<output>.rejects.csv` beside
   the output EIB for which rows and why); **red** = the run failed. Applied centrally by `HighlightRunResult`
   (`modHelpers`) via the dispatcher, so every transform gets it.
2. **`RunnerStatus.txt`** — written beside `Runner.xlsm`, **overwritten** each progress tick with
   the current line, and a final summary at the end. This is an out-of-band progress channel you can
   watch even while the grid is frozen. Tail it — `Get-Content RunnerStatus.txt -Wait` (PowerShell)
   or `tail -f` — **do not open it in Excel while running**, because Excel would hold a lock and the
   next progress write would be skipped (harmlessly).
3. **`RunnerLog.csv`** — written beside `Runner.xlsm`, **append-only**, one row per completed run:
   `Run Start, Run End, Duration (s), Transform, Status, Rows Processed, Rows Rejected, Output
   Workbook, Issues`. This is the run history; the user may delete it any time.
4. **`<output>.rejects.csv`** — written **beside the output EIB** (named to pair with it, e.g.
   `AP_Travel_output.xlsx` → `AP_Travel_output.rejects.csv`), so the generated EIB and the rows that
   failed to load into it sit in the same folder. **Reset at the start of each run** (so it always
   reflects the latest run), one row per *rejected* source row: `Source Row,
   Reason, Source Value, Details`. `Source Row` is the actual row number in the source sheet, so
   finance can open it directly; `Reason` is why it was dropped (missing field, zero amount, or an
   unresolved lookup — e.g. `Cost centre unresolved: Final Code did not resolve in 'MAPPING MASTER'
   (Cost Center, Core vs Prism = Core, lowest hierarchy match)`); `Source Value` is the source cell
   the reject turned on — the unmapped Final Code or Customer Number — and is blank for rejects with
   no single value to blame (missing-field); `Details` is a free-text handle for recognising the row
   (memo, description, document number). This is how you answer "why is `Rows Rejected` non-zero?" —
   the rejected rows are almost always source-data errors (e.g. a garbled account string that cannot
   be coded) for finance to fix at source.

   **Sending rejects back to the owning team**: `Source Value` is its own column so the file can be
   pivoted on it — a run typically rejects many rows over a handful of distinct bad codes, and the
   team that hand-maintains those codes (AP for Final Code, AR for customer numbers) wants the
   distinct list, not the row list.

All three are best-effort and created on first run (self-healing). A `RunnerLog.csv` write failure
is surfaced in the run's `Issues` (it never aborts the run); a `RunnerStatus.txt` write failure is
skipped silently since it only mirrors progress. `RunnerLog.csv` and `RunnerStatus.txt` are
runner-scoped singletons and stay beside `Runner.xlsm`; only the rejects file is co-located with its
output EIB. If the output location can't be resolved (e.g. an unsaved run), the rejects fall back to
`RunnerRejects_<transform>.csv` beside `Runner.xlsm`. Because none live inside the workbook, they all
survive a rebuild (`build_runner.py` regenerates only `dist/Runner.xlsm`).

## Diagnosis

- If the shell must be recreated, regenerate it with `scripts/create_runner_shell.py`; do not hand-edit the binary as a source of truth.
- **Never let any tool write to the VBA project binary of an existing `.xlsm` directly** (e.g. `pyOpenVBA`'s
  `add_module`/`delete_module` against `dist/Runner.xlsm`) — this was tried and produced a workbook Excel
  itself could not repair-save. The one safe exception is `create_runner_shell.py`'s single `add_module`
  call on a **freshly created** workbook, for `BootstrapImporter` only.
- A VBA "Compile error" with a runtime-style message (e.g. "Invalid procedure call or argument") appearing
  from Debug → Compile VBA Project alone, with no code executed, is a sign of binary corruption, not a
  source defect — rebuild the artifact from a clean shell rather than debugging the VBA further.
- If a VBA routine relies on arbitrary file writes, prefer reporting results into workbook cells instead; ad hoc file output is unreliable on Mac Excel from this environment.
- If a module proposal includes `ADODB`, `Scripting.Dictionary`, `FileSystemObject`, `RegExp`, or Power Query automation, reject it as outside the supported runtime subset.
- If `osascript` can query Excel but `run VB macro` fails with `A privilege violation occurred`, that's
  usually a stuck/denied Automation permission entry — `tccutil reset AppleEvents com.apple.Terminal` and
  re-trigger the prompt (e.g. `osascript -e 'tell application "Microsoft Excel" to activate'`), then
  re-grant access in System Settings.
- Do not attempt to script the module-import macro (`BootstrapImportFromConfiguredSource` /
  `BootstrapImportFromFolder`) headlessly via AppleScript — it is unreliable on this Mac Excel build
  across multiple independent failure modes (parameter-passing failures, AppleEvent timeouts, a broken
  Apple Event connection after a mid-import error). Run it manually in Excel. Zero-argument macros that
  don't touch the VBA project (e.g. `modDispatcher.RunSelectedTransform`) automate fine headlessly — this
  limitation is specific to VBA-project manipulation, not to Excel automation in general.
- If `BootstrapImporter` fails on one specific module with `Method 'Import' of object '_VBComponents' failed`,
  treat that as a source-module parse/import issue first, not a generic Excel failure. If the error
  message additionally says *"Imported 'Module1'"* (the parser's blank fallback) rather than the real
  module name, check these in strict order, because the earlier items are the fastest to verify:
  0. **Module name > 31 characters.** VBA's hard 31-char limit on component names. If
     `Attribute VB_Name` / the `mod<Name>.bas` filename exceeds it, `Import` drops the name, keeps the
     blank `Module1`, and fails the rename back — surfacing as `-2147219692 (80040714)`. This is a
     name-length limit, not a mid-parse abort, and is the cheapest to rule out:
     `python3 -c "import os,glob; print([f for f in glob.glob('vba_src/**/*.bas',recursive=True) if len(os.path.basename(f)[:-4])>31])"`.
     Shorten the module name only; the `Run<Name>` Sub / `step_list_macro` name is not limited.
  1. **Non-ASCII / Unicode content in the `.bas` file.** Mac Excel's `VBComponents.Import` is
     ANSI-first and silently aborts when multi-byte UTF-8 sequences appear in volume, even inside
     comments. Pure-ASCII (0 non-ASCII bytes) always works; a tiny amount (≤ ~20 bytes, e.g. 2 em
     dashes as in `modHelpers.bas`) is tolerated; hundreds of non-ASCII bytes (decorative `─`,
     `—`, `→`, curly quotes in comment headers) reliably fail. Fix: transcode to ASCII
     equivalents and confirm `python3 -c \"d=open(f,'rb').read(); print(max(d))\"` returns < 128.
  2. **CRLF line endings** (not LF-only). With LF-only the parser may fail to recognise the
     `Attribute VB_Name` / class-module headers and again falls back to `Module1`. Fix:
     `sed -i '' 's/$/\r/' file.bas` and verify with `file(1)`.
  3. **Mismatched `End Sub` / `End Function`** — an unterminated procedure aborts the whole import.
  4. **Overly long line-continuation-heavy procedure declarations/calls.** We have hit this exact
     last case before on Mac Excel; refactor shared state into module-level variables or helpers
     to shorten the signature.
- If `BootstrapImporter` fails immediately with `Excel could not access this workbook. VBAProject/VB
  Components. Enable trust access to the VBA Project object model…` (before any module imports), this is
  the Excel trust setting, not the source: Excel → Preferences → Security → under **Developer Macro
  Settings** enable **"Trust access to the VBA project object model"** (and Macro Security → Enable all
  macros), quit Excel (⌘Q), reopen, retry. Separate from macOS Full Disk Access. See `docs/BUILD.md` step 3.

## Recovery

- If the built Runner workbook is incorrect, fix the source text under `vba_src/` or `build/` and rebuild; do not hand-edit imported VBA inside Excel.
- If the Runner Shell workbook is damaged, recreate it with `scripts/create_runner_shell.py`.
- If `dist/Runner.xlsm` won't compile or won't repair-save, delete `dist/Runner.xlsm`, regenerate
  `runner_shell/RunnerShell.xlsm` with `scripts/create_runner_shell.py --overwrite` (never `rm` it
  directly — that bypasses the preserved-sheets mechanism and can destroy the `mapping` tab; see
  Gotchas in the anchor doc), then rerun `scripts/build_runner.py` and redo the manual import step from
  scratch rather than patching the existing binary.
- If a run produces unclear transform results, inspect the cloned Output workbook and the `Issues` /
  `Rows Skipped` result cells on the Runner sheet — there is no separate Interim workbook to inspect;
  each transform macro opens its inputs and writes the output directly.
- If backlog state becomes unclear, resume from `docs/TODO.md` and the active anchor doc under `docs/todo/`.

## Running AR Customer Master (`ar_customer_master`)

Status 2026-10-12: code-complete, **pending the manual Excel import + verify** (the macro, `modEntityBlockWriter`
and `modHelpers.LookupRuntimeValue` have not been run in Excel yet).

1. Build the input template: `uv run --with openpyxl --with pyyaml --no-project python scripts/create_ingest_template.py ar_customer --overwrite`
   (writes `dist/ingest_templates/ar_customer.xlsx`; 36 columns after the 2026-10-12 simplification, was 43).
2. AR fill sheet `Customer`, one row per customer, and save it as the Source Data path in the Runner config.
3. Build/import the Runner as above (steps 1-5 of Current Workflow), pick **AR Customer Master**, click **Run Transform**.
4. Output: `Master_Customers_ar_customer_master_output.xlsx` (clone of the 3+ Master Customers EIB, sheet
   `Submit Customer`, data from row 6) plus a `.rejects.csv` beside it for rejected customers.
5. Rule (operator, 2026-10-12): one address, one email, one phone number, one bank account, one customer group
   per customer. The EIB has one row per customer EXCEPT the address lines: each non-blank Address Line 1/2/3 goes
   on its own EIB row (columns B, BL, BU, BW, BX only on rows 2+). Nothing else repeats. A missing value or a
   missing Workday reference ID rejects the whole customer (no fallbacks).
6. After the Excel run, compare the output with `Master Customers.xlsx` sample rows 6 and 9 (see
   `config/transform_specs/ar_customer_master.yaml`).

## Running AP Supplier Master (`ap_supplier_master`)

Status 2026-10-12: code-complete, **pending the manual Excel import + verify** (the macro, `modEntityBlockWriter`
and `modHelpers.LookupRuntimeValue` have not been run in Excel yet).

1. Build the input template: `uv run --with openpyxl --with pyyaml --no-project python scripts/create_ingest_template.py ap_supplier --overwrite`
   (writes `dist/ingest_templates/ap_supplier.xlsx`; 46 columns after the 2026-10-12 simplification, was 53).
2. AP fill sheet `Supplier`, one row per supplier, and save it as the Source Data path in the Runner config.
   Supplier ID is required and typed by AP (operator, 2026-10-12: "we wont entertain auto generating the workday id
   for the custoer/suplier - they have to present them"); Reference ID is copied as typed.
3. Build/import the Runner as above (steps 1-5 of Current Workflow), pick **AP Supplier Master**, click **Run Transform**.
4. Output: `Master_Suppliers_ap_supplier_master_output.xlsx` (clone of the 3+ Master Suppliers EIB, sheet
   `Submit Supplier`, data from row 6) plus a `.rejects.csv` beside it for rejected suppliers.
5. Rule (operator, 2026-10-12): one address, one email, one phone number, one bank account, one supplier group,
   one payment type per supplier. The EIB has one row per supplier EXCEPT the address lines: each non-blank
   Address Line 1/2/3 goes on its own EIB row (columns B, DD, DM, DO, DP only on rows 2+). Nothing else repeats.
   Use For is BILLING (the 3+ sample's SHIPPING / REMIT / PROCUREMENT rows are not produced). A missing value or a
   missing Workday reference ID rejects the whole supplier (no fallbacks).
6. After the Excel run, compare the output with `Master Suppliers.xlsx` sample rows 6 and 10 (see
   `config/transform_specs/ap_supplier_master.yaml`) and check the Status value (`ACTIVE` vs `Active`, see `docs/TODO.md`).

## Related Docs

- `PRD.md`
- `docs/MASTER_CUSTOMER_SUPPLIER.md` (entry point for AR Customer Master / AP Supplier Master)
- `docs/ARCHITECTURE.md`
- `docs/CONFIGURATION.md`
- `docs/ops/CLI.md`
