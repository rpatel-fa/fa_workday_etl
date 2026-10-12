# Cold-Start Agent Contract: Prompt Interpretation + Ambiguity Handling

> **Canonical home**: `docs/AGENT_COLD_START_PROMPT_CONTRACT.md`
>
> **Who reads this:** An agent dropped into `elt_transform_excel/` fresh (cold start, no hidden memory,
> no prior session context) who is given a user prompt like
> `"in elt_transform_excel, create new transform for …"`.
>
> **Who edits this:** When the project gains a new entry point (e.g. Setup EIB archetype, new source
> format, new valid prompt form) — update THIS file, never rely on assistant memory.

## 1. Purpose

This document is the **sole source of truth** for what a cold-start agent must do when it receives one of
the three standard user prompts. It defines:

1. The **three valid primary prompt forms** a user will typically write.
2. How to **detect ambiguity** in each form (inputs that are unspecified, partial, or under-specified).
3. The **follow-up question templates** the agent MUST fire — and ONLY fire — for each ambiguous input,
   using the project's standard AskUserQuestion mechanic (or equivalent interactive surface).
4. The **deterministic pipeline** the agent follows once all inputs are unambiguous, including the
   mandatory build-gate and reconcile close-out.
5. A **zero-follow-up prompt short form** the user can optionally supply to skip the interview.

If an agent finds a scenario not covered by this contract → **DOCUMENT THE GAP HERE, then ask**.
Never silently assume values — the whole point of the g10/g11 gap sweeps was to eliminate hidden inference.

## 1.1 Communication Rule - Do Not Repeat Canonical Operator Instructions

When reporting status or closing out work, the agent MUST NOT re-type build/manual operator
steps that are already canonically documented in `docs/BUILD.md`, `docs/BUILD_WIN.md`, or
`docs/RUNBOOK.md` unless the user explicitly asks for those steps again.

Required behavior:

1. If the operator procedure is already documented and the user did not ask for it, give the
   result/blocker only and reference the canonical doc path if needed.
2. If the work is pending a documented manual Excel/bootstrap/verify step, say that briefly
   once, without re-listing the click-path / command sequence.
3. Reproduce the actual procedure only when the user explicitly asks for the procedure, or
   when the documented step itself is the subject under discussion.

Anti-pattern:

```text
BAD: "Still required: open Runner.xlsm -> run BootstrapImporter -> save -> close -> run --verify"
     on every completion message, even though BUILD.md already documents it.

GOOD: "Build prep is complete; the remaining manual Excel bootstrap/verify step is the one
      already documented in docs/BUILD.md."
```

---

## 0.3. MANDATORY QUICK-LOOKUP WINS (run FIRST when prompt matches)

> **Purpose.** Cold-start agents often over-read the repo for trivial report-style asks ("how
> many EIBs?", "list all Runner labels", "stakeholder manifest summary"). These shortcuts
> pre-empt that. Match before any other planning step. If the user prompt matches the pattern
> → run the script. If the script exists and exits 0 → its output is the answer. No file
> browsing or manual counting.

| Pattern | Action (run, then paste output verbatim into the answer) |
|---|---|
| Stakeholder-count / manifest / label / list-all-Runner-dropdown / AP-AR-EIB-summary / "how many EIBs?" / "list all labels" | `uv run --with pyyaml --no-project python implementations/elt_transform_excel/scripts/manifest_summary.py` (raw dropdown counts: 15 AP / 5 AR / 4 defunct as of Sep-2026) |
| Same, but stakeholder / C-suite / "Amex Trains recorded as one" / consolidated view | Add `--consolidate-amex-trains` flag (merges `ap_amex_trains` + `ap_amex_trains_v2` → 1 line; 14 AP / 5 AR total) |
| Same, but caller wants JSON / machine-readable counts | Add `--json` flag (includes `_summary` block with totals) |

**Anti-pattern (ALWAYS a defect for these query types):**

```
BAD: manually open each config/transform_specs/*.yaml → count labels by hand →
     possibly drift from build_manifest.yaml, miscount variant fan-outs.
GOOD: run scripts/manifest_summary.py → paste output.
```

The manifest summary script reads ONLY two things: `config/build_manifest.yaml`
(`active_profiles` / `defunct_profiles` lists) and the matching
`config/transforms/<profile_key>.yaml` (`transform.label`, or
`shell_defaults.variants[].profile_label` for fan-out variants like Stationery, GBE, Taxi).
It **never** opens `config/transform_specs/`. If its output disagrees with a manual count, fix
the script or the transform YAML — never patch the report manually.

---

## 0.4. MANDATORY CANONICAL READ ORDER (CONFIG-FIRST, BEFORE ANY VBA OR YAML EDITS)

> **This rule is the single most important rule in this file.** Failure to follow it was
> the root cause of the AUG-2026 shared-helper order-inversion incident where
> `NormalizeSourceDate` was patched in VBA BEFORE `config/shared_helper_registry.yaml`
> and the 5 AR `transform_specs/*.yaml` were updated. The VBA was "correct on the machine"
> but the SSO contract was empty for 5 specs and a cold-start agent reading only
> config/docs would have produced a contradictory implementation.
>
> **There is one correct timeline. ANY other order is a DEFECT and must be reworked before
> the build passes.**

### Correct Timeline (MUST follow exactly)

1. **READ DOCS FIRST.** Open these files in this EXACT order; no VBA file opens, no grep of
   `vba_src/**/*.bas` until step 4.
   1. `docs/AGENT_COLD_START_PROMPT_CONTRACT.md` (THIS FILE) — §0.4, §0.5, §0.6
   2. `docs/ARCHITECTURE.md` — section 4 Shared modules; locate the pointer to the
      appropriate registry concern (shared helpers live here).
   3. `docs/CONFIGURATION.md` — row `source_format` / `target_format`; locate the
      `shared_helper:<concern>.<variant>` token reference.
   4. `config/shared_helper_registry.yaml` — CANONICAL SSO. Locate (or create) the
      `shared_helpers.<concern>.<variant>` dict that describes the behavior. This is the
      authoritative thing you are implementing.
   5. `config/transform_specs/*.yaml` — every profile whose runtime VBA invokes the
      shared helper MUST declare the concern in `shared_helpers:` with variant != `none`.
      Scope is AP + AR. There is no "AR deferred" exemption.
2. **WRITE (or update) THE SSO CONTRACT IN CONFIG.** Only after step 1 is satisfied:
   1. If the concern/variant doesn't exist in `shared_helper_registry.yaml` → create it
      (fields: `description`, `vba_module`, `vba_public_fn`, `execution_order`,
      `invariants[]` with rationale, `examples[]` = list of profile keys).
   2. If any transform spec's `shared_helpers:` block omits the concern → add the entry.
   3. If `ARCHITECTURE.md` / `CONFIGURATION.md` redirect pointers are missing → add them.
3. **BUILD GUARDRAIL PASS #1.** Run
   `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/build_runner.py --hygiene-only`.
   It MUST exit clean with `HYGIENE OK`. If any ERROR from the Step 6 Stage 1
   config↔VBA linkage/coverage guards fires → RESOLVE THE CONFIG GAP. DO NOT start
   writing VBA to "fix" it yet. This pass is config-only.
4. **ONLY NOW** — open the target VBA module (named by `vba_module` in the registry)
   and author the function body as a 1:1 mirror of the registry's `execution_order` and
   `invariants[]`. No "smart" refactors beyond what the invariants allow.
5. **BUILD GUARDRAIL PASS #2.** Re-run `--hygiene-only`. Same `HYGIENE OK` output. If any
   ERROR fires → revert the VBA change and re-check the contract first. If it passed
   before step 4 and fails after → your VBA drifted from the contract. Rewrite.
6. **Bootstrap Import (optional, user-requested deploy step).**
   `BootstrapImporter.BootstrapImportFromConfiguredSource` inside `Runner.xlsm`.

### Order-Inversion Anti-Pattern (ALWAYS A DEFECT — do not ship)

```
BAD ORDER: VBA written → build fails (because registry/specs missing concern) →
config/specs/docs patched post-hoc just to make hygiene pass.
```

This order means the YAML is a post-hoc rationalization of code, not the source of truth.
If your timeline looks like this → **REVERT the VBA change**, re-run from step 1, then
re-apply the VBA after the config pass is clean. Build hygiene can't detect the
chronological order, so this rule is SELF-ENFORCED by the author; the config shape
(naming, invariants, examples) should be authored before the VBA is committed in the
same atomic change.

---

## 0.5. MANDATORY VBA PRE-EDIT / PRE-BUILD GATE (run BEFORE any VBA edit, Runner shell create, or build_runner.py prep)

> **Failure to run this gate first was the #1 recurring cold-start failure in every session through
> AUG-2026 (10+ incidents).** The symptom: you write a valid VBA module in the IDE with em-dashes /
> right-arrows in comments or expanded function calls on 320-char one-liners → Mac Excel's
> `VBComponents.Import()` silently falls back to a blank `Module1` → entire Runner macro contents
> vanish → user sees "sub or function not defined" → 30+ minutes debugging lost modules. **This gate
> prevents it entirely, zero ambiguity about when to run.**

**Run EVERY SESSION the agent touches `implementations/elt_transform_excel/vba_src/` `.bas` files or Runner build files, in this exact order:**

### Step G1 — Ultra-fast stdlib pre-flight (no deps, < 50 ms):

```bash
python3 implementations/elt_transform_excel/scripts/check_vba_hygiene.py
```

If any NON_ASCII WARN/ERROR appears (em-dashes, curly quotes, right-arrows, ellipsis in comments from
copy-paste or IDE smart-quote auto-substitute):

```bash
# Auto-strips the 26 most common non-ASCII chars to ASCII equivalents in-place across all .bas:
python3 implementations/elt_transform_excel/scripts/check_vba_hygiene.py --fix
```

Rules enforced in G1:
- **NON_ASCII bytes count**: 0–20 WARN (legacy tolerated from modHelpers/modFinanceCommon untouched
  source modules), > 20 = HARD ERROR. VBA Import silently falls back to blank Module1 at 40+ bytes on
  Mac Excel; we fail long before then.
- **CRLF**: any lone-LF → HARD ERROR. Line-ending must be `\r\n` everywhere (Excel VBA source spec).
- **COMPONENT_NAME**: filename stem or Attribute VB_Name > 31 chars → HARD ERROR (31 char is the
  VBA hard module limit; overlength modules lose their name at import time and become `Module1`,
  breaking `Application.Run` dispatch).
- **LONG_LINES**: lines 251–299 chars = WARN (safe, consider continuation); **≥ 300 chars = HARD
  ERROR.** Build-observed failure: Mac Excel's VBA parser aborts import mid-file on 258+ char lines
  even when line-continuation syntax is valid. Refactor with ` _` (space + underscore) before
  continuing.
- **WORKSHEETS_INDEX**: `Worksheets(n)` numeric index access (not inside a `'` comment line) → HARD
  ERROR. Finance tabs are re-ordered constantly on Runner sheets; numeric lookups silently point to
  the wrong worksheet.
- **RSTK_CALL_ARITY**: `ResolveSelectedTransformKey()` call sites not taking exactly 2 untyped args
  `(runnerWorkbook, controlSheet)` → HARD ERROR. Stale 1-arg calls from old 2025 patterns fail at
  runtime inside the Run button dispatch.
- **RUN_SUB_SIG**: Every `Public Sub Run<Name>(…)` in a transform module must declare exactly
  `(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)` (no other signature, no
  `Variant`-typed args). Any mismatch → HARD ERROR; dispatcher won't call the right entry point.
- **STEP_LIST_MACRO_MISSING**: If a profile YAML declares `runtime.step_list_macro: RunFancyImport`,
  then `Public Sub RunFancyImport(…)` must exist in some .bas under `vba_src/transforms/`. Missing →
  HARD ERROR (you can't edit YAML and then forget to write the VBA entry point — the most common
  "clean build, runtime explodes" bug).
- **PORTABLE_VBA_NO_ACTIVEX** (THE BIG ONE FOR CROSS-OS DISTRIBUTION — run-time 429 guard): If any
  `.bas`/.cls source file contains `Scripting.Dictionary`, `CreateObject("Scripting.` /
  `CreateObject('Scripting.`, `ADODB.`, `Scripting.FileSystemObject`, `VBScript.RegExp`, `CreateObject("ADODB.`,
  `CreateObject('ADODB.`, `CreateObject("VBScript.`, `CreateObject('VBScript.`, or any of the 42 banned
  COM-only/ActiveX strings from the canonical list in ARCHITECTURE.md § "Native VBA only — no COM-only
  libraries", **outside a comment-only line** (i.e. NOT on a line where the first non-space character is
  `'`) → HARD ERROR.  These COM-only type libraries ship ONLY on Windows VBA hosts.  On Mac Excel
  (the engineering authoring host for this project) and on locked-down Windows laptops with Trust
  Center "Disable all ActiveX" group policy, `CreateObject("Scripting.Dictionary")` raises Run-time
  error '429': "ActiveX component can't create object" the moment Finance users open the Runner
  workbook.  Mac-safe replacements: VBA `Collection` for dictionary semantics (see
  `CollectionKeyExists` / `CollectionSetItemLong` / `CollectionGetItemLong` helpers in
  `modApCfaClaimsSupplierInvoice.bas` as the pattern), `Open`/`Print #`/`Input #` for file I/O instead
  of `Scripting.FileSystemObject`, hand-rolled tokenization instead of `VBScript.RegExp`, Worksheet
  array reads instead of `ADODB` SQL.  The Finance hand-off contract REQUIRES the Runner .xlsm to
  open and execute with ZERO Trust Center prompts, ZERO Tools > References tick boxes, and ZERO
  ActiveX dependency install steps.  Guard (vii) in `build_runner.py` Step 6 Stage 1 enforces this
  mechanically with ERROR severity, not warning.

### Step G2 — Full-contract hygiene run inside build_runner.py (runs `_validate_vba_source_hygiene`
with full profile YAML bindings, checks RSTK/RUN_SUB signatures against all 9 profiles):

```bash
uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project \
    python implementations/elt_transform_excel/scripts/build_runner.py --hygiene-only
```

Output `HYGIENE OK — N modules, M bindings, 0 ERROR-class violations.` → proceed. Anything else →
STOP. Never run `create_runner_shell.py` or full `build_runner.py` prep until G1 and G2 both pass with
0 ERRORs.

### Step G3 (POST-edit, after writing any VBA change) — Re-run G1 + G2 one more time.

This catches any new em-dashes / 310-char one-liners you just wrote in the session (the #1 failure
mode was: newlines added a comment with em-dash mid-stream before build step).

---

## 0.6. SHARED HELPER AUTHORING CONTRACT (build_runner.py Step 6 Stage 1 enforcement)

> Applies to ANY change in `modSourceTable.bas`, `modFinanceCommon.bas`, or any shared
> module whose public exports are declared by a concern/variant pair in
> `config/shared_helper_registry.yaml` and referenced in the `shared_helpers:` block of
> one or more `config/transform_specs/*.yaml` files.

All six sub-guards below are enforced mechanically by
`elt_transform_excel/scripts/build_runner.py` Step 6 Stage 1. If any ERROR fires, the
build passes only when it's resolved. This is the contract `--hygiene-only` validates; if
you understand these six guards you will never produce a shared-helper change that
passes locally but fails the canonical build.

### Guard table

| # | Direction | Guard name | What the build validates | Exact incident it prevented (AUG-2026 witness) |
|---|---|---|---|---|
| (i) | Spec existence | Spec→file | A transform spec YAML exists for every transform .bas module. |  |
| (ii) | Spec→Registry | Concern/variant resolve | Every key in `shared_helpers:` of a transform spec resolves to a valid `shared_helper_registry.yaml` concern + variant. Typo `exel_serial…` → ERROR. |  |
| (iii) NEG | Config→VBA | Hoisted-helper private copies forbidden | Transform .bas (AP + AR) must NOT declare `Private Function` copies of any helper in `HOISTED_HELPER_PRIVATE_FN_NAMES` — these are Public in shared modules. |  |
| (iv) POS forward | Config→VBA | Declared helper must be called | If a spec declares `shared_helpers.<concern> = <non-none variant>`, the corresponding bare helper hint name must be referenced in the .bas at runtime. |  |
| **(v) POS reverse** | **VBA→Spec** | **Called helper must be declared (THE BIG ONE)** | **If a transform .bas references a shared helper name in `BARE_HELPER_TO_REQUIRED_CONCERN` at runtime, the transform spec MUST declare the matching concern under `shared_helpers:` with variant != `'none'`. Scope = AP + AR. No exemptions.** | **AUG-2026 order-inversion incident: all 5 AR specs had NO `shared_helpers:` block at all, yet all 5 called `NormalizeSourceDate`. A cold-start agent reading only config/docs would have correctly concluded no SSO date-normalization contract existed for AR. Guard (v) now fires ERROR with an explicit CONFIG-FIRST remediation message showing exactly which spec/concern to add.** |
| (vi) | Registry→VBA | Public export linkage | For each concern/variant that names `vba_module` + `vba_public_fn` in the registry: (a) the target `.bas`/`.cls` file exists; (b) every function identifier parsed flexibly from `vba_public_fn` (singleton name, `"A + B"` prose, YAML list, or signed parenthesized call) exists as a `Public Function`/`Sub`/`Property` export in that module. Prose `vba_module` strings like `"local transform (...)"` are treated as inline variants and skipped. | AUG-2026 secondary: `company_resolution.per_row_split_from_supplier_id` in the registry named public helper `SplitSupplierIdRow` which was never exported as `Public` (it was a Private internal of `modFinanceCommon`). Guard correctly flagged → registry corrected to `ResolveSupplierWithCompanyPrecedence` (the real runtime call site) with the private split routine listed under `companion_public_helpers:`. |
| **(vii)** | **VBA→Runtime host** | **PORTABLE_VBA_NO_ACTIVEX (run-time 429 guard, cross-OS hand-off contract)** | **Scans every source line of every `.bas` / `.cls` under `vba_src/` (AP, AR, shared, bootstrap) for 42 banned COM-only / ActiveX / Type-Library reference strings (Scripting.Dictionary, ADODB.Connection, VBScript.RegExp, Scripting.FileSystemObject, plus all `CreateObject(\"Scripting.`) / `CreateObject(\"ADODB.`) / `CreateObject(\"VBScript.`) variants with both quote styles). Comment-only lines (first non-space char is `'`) are excluded so docs inside modules still name the banned libraries for historical/rationale context. Any match on an executable (non-comment) line → ERROR.** | **AUG-28-2026 CFA Claims incident: Rule 8 per-group increment used `CreateObject(\"Scripting.Dictionary\")` (late-bound, no Tools > References required, passed Python hygiene on authoring Mac, silently compiled OK in Excel). When Finance ran the built Runner on their Windows laptops the call raised Run-time error '429': 'ActiveX component can't create object' on the `Application.Run` dispatch path. The Registry-to-VBA guard (vi) and all 5 other guards passed because the call was a Private implementation detail, not a shared helper.  Guard (vii) would have fired the ERROR immediately on the banned `CreateObject(\"Scripting.` string during --hygiene-only instead of letting the 429 reach the end user.  Banned libraries are the central engineering constraint recorded verbatim in ARCHITECTURE.md § 'Native VBA only — no COM-only libraries'.**|

### When to add a new concern/variant pair to the registry

Any time you write or modify a shared VBA helper AND that helper is called by 2+ transform
modules (i.e. the helper qualifies for the "shared helper" layer, not a per-transform
private routine), add a new concern or a new variant of an existing concern to
`shared_helper_registry.yaml` BEFORE writing the VBA. Required fields:

- `description` — one-paragraph forcing-function summary of what this helper decides (NOT
  how it's implemented)
- `vba_module` — filename stem (e.g. `modSourceTable.bas`)
- `vba_public_fn` — flexible shape (simple name, `"A + B"`, or YAML list of public
  export identifiers the variant calls at runtime)
- `execution_order` — numbered list of branches with precedence, the thing the VBA
  mirrors 1:1
- `invariants[]` — each entry is an invariant string with a `rationale:` subkey. The
  rationale is mandatory because the invariant is usually the subtle thing that breaks
  (e.g. "`Fix()` not `CLng`/`Round` for time truncation" → rationale is "CLng rounds 12:00+
  to the next day").
- `examples[]` — list of profile keys that consume this variant (helps scope coverage
  checks in guard (v)).

Then add `shared_helpers.<concern>: <variant>` to every transform_spec whose VBA calls
that helper (AP + AR). Then run the §0.4 timeline.

---

## 2. Three Valid Primary Prompt Forms

A user will normally type one of these three. Any deviation is a new form and should raise a clarifying
question before acting.

### Form 1 — Update an existing transform

```
in elt_transform_excel, update transform for <stem> [details]
```

#### Examples

```
in elt_transform_excel, update transform for ap_amex_hotel: change uniform_tax to GBR_Standard_Rated
in elt_transform_excel, update transform for ar_invoice — new source file at /…/InvoiceDetailFALL_v2.csv
in elt_transform_excel, update transform for ap_referees — the batch_id_token is now REFS0427ME
```

#### Ambiguity detection — Form 1

| # | Condition | What agent MUST ask |
|---|---|---|
| 1 | `<stem>` is NOT an exact match for any file in `config/transforms/*.yaml` (after `yaml.safe_load` confirms it's a build profile, i.e. has `transform.key`). | `"I couldn't find a profile named '<stem>'. Available profiles: <alphabetical list>. Which one do you want to update, or is this a new transform?"` |
| 2 | Prompt says "update" but gives **no actual change** (no `details`, no colon, no "change X to Y"). | `"What change would you like to make to transform '<stem>'? I can update: source_layout worksheet/header/required_headers, transform_constants (company, supplier, tax, worktags, final_code), shell_defaults paths (source_data / target_eib / output_file_path), or the transform_spec decision record."` |
| 3 | The requested change touches a **tenanted business value** (e.g. a tax applicability code, a new company code, a supplier Reference ID) and the value does not appear in any `config/ingest_lookups/*.json` map. | `"Value '<value>' you proposed does not appear in the tenant ref-ID maps under config/ingest_lookups/. Confirm this is a new Workday value that Finance has already provisioned, and I'll write it into transform_constants verbatim — or did you mean one of the existing values: <top-5 closest by edit-distance>?"` |
| 4 | The requested change would flip an invariant the build gate validates (e.g. remove `uniform_tax` on a Supplier profile without adding `per_row_tax`, drop a `worktag_business_objects` key). | `"That change would break the per-archetype schema gate (see _validate_transform_constants_schema in scripts/build_runner.py). Specifically: <explain gate>. Do you want me to (a) also add the required sibling field to keep the gate passing, OR (b) keep the old value?"` |

#### Deterministic pipeline after disambiguation

1. Edit `config/transforms/<stem>.yaml` (and, if mapping/decision changes, `config/transform_specs/<stem>.yaml`).
2. Run build gate: `uv run --with pyyaml --no-project python -c "from scripts.build_runner import _load_transform_bindings,_Path; _load_transform_bindings(_Path('config/transforms'))"` — must exit 0.
3. Run reconcile for this profile: `scripts/reconcile_supplier_customer_inputs_outputs.py --profiles <stem> --scan-rows 500 --scan-cols 200` — must exit 0.
4. If a new `.bas` constant was introduced, run impl-detail audit forward to confirm DRIFT stays 0. Report the result.

---

### Form 2 — Create a new transform for Supplier Invoice EIB

```
in elt_transform_excel, create new transform for input source <source-ref> to produce supplier-invoice EIB [details]
```

#### Examples

```
in elt_transform_excel, create new transform for input source /Users/me/AP_amex_airlines.xlsx to produce supplier-invoice EIB
in elt_transform_excel, create new transform for input source abc to produce supplier-invoice EIB
in elt_transform_excel, create new transform for input source abc to produce supplier-invoice EIB. Fork from ap_amex_hotel. Same uniform_tax. Company=FA, supplier ref-id=AME003.
```

#### Ambiguity detection — Form 2

| # | Condition | What agent MUST ask |
|---|---|---|
| 1 | `<source-ref>` does not start with `/` (it's not a full path) AND it does not match a single unambiguous `.xlsx/.xlsm/.csv` file under the standard AP source dir convention `/Users/Rakesh.Patel/Documents/__data/excel/finance_ref/eib/ap/<source-ref>/source/` when we glob for the only source-format file there. | `"To create the transform I need the full disk path to the source input workbook/CSV. I didn't find a single file under the AP convention from the label '<source-ref>'. Please paste the full path (e.g. /Users/…/AMEX_AIRLINES_APR26.xlsx), OR confirm the folder I should scan for the only source file."` |
| 2 | `<source-ref>` resolves to a FOLDER, and that folder contains **2+** files with suffixes `.xlsx / .xlsm / .xltx / .xltm / .csv`. | `"Folder '<source-ref>' contains these source-format files: <list>. Which one should I use?"` |
| 3 | Prompt does **NOT** name an exemplar to fork from, AND the source workbook's header-name similarity (cosine-similarity of header-token set union against each existing 6 Supplier profiles' source-layout required-header sets) doesn't yield a best match ≥ 0.60 cosine. | `"I can't auto-pick a good exemplar fork from the 6 existing Supplier profiles by header similarity. These are the best matches: <top-3 with similarity score>. Which one should I fork from? (ap_amex_hotel = simple/fixed; ap_amex_trains_v2 = per-row lookups; ap_cfa_distributions = Final-Code→Map worktags; ap_prize_funds = no-company-col Integration tab; ap_referees = Referees pattern; blank = start from scratch)."` |
| 4 | The prompt has no `details` block that (together with the forked exemplar defaults) fully determines the 8 required prompts enumerated in §3 below. | Agent MUST then fire the **8 required archetype prompts** verbatim from §3 of this document. It MUST NOT invent values — the whole prior gap audit showed the profile YAML is the only source of truth. |
| 5 | After opening the source workbook, scanning for required headers (see §4) yields no row containing ALL the proposed `required_headers` within `scan_rows=400 × scan_cols=160`. | `"I opened the source and couldn't find a header row containing all standard columns. The best partial match (row X) has Y/Z: <matches vs missing>. Confirm the correct worksheet name / correct header row position, OR confirm I should treat the found row as sufficient and add custom required_headers specific to this source."` |

#### Deterministic pipeline after disambiguation

1. Open the source → scan → propose source_layout (worksheet + header_row + data_start_row + required_headers list).
2. Propose exemplar fork based on header-name similarity (≥ 0.60); otherwise ask (§2.2, #3).
3. Ask (or accept from prompt-details block) the 8 required prompts from §3.
4. Write `config/transforms/<new_key>.yaml` (build profile) and `config/transform_specs/<new_key>.yaml` (decision record, 1:1 field list with column targets).
5. **Mandatory build gate.** Run `_load_transform_bindings(Path('config/transforms'))` → confirm the new profile passes all 4 schema validators: eib_defaults 8-key, source_layout enum check, transform_constants archetype check, SUPPORTED_EIB_POPULATION_TYPES whitelist. Fail FAST and report before generating any code.
6. **Fork/generate the `.bas` module.** Copy the closest exemplar module, rename `Private Const SOURCE_*` values to match the new profile's `transform_constants`. Swap per-archetype column refs if target EIB layout changed. (This is the one manual-ish step; automate with a dedicated `scripts/generate_transform_macro_from_config.py` when project hits 20 transforms per H3 rule.)
7. **Build Runner shell + write registry.** Run `scripts/build_runner.py`.
8. **Reconcile close-out.** Run `scripts/reconcile_supplier_customer_inputs_outputs.py --profiles <new_key> --scan-rows 500 --scan-cols 200`. Result must be 0 FATAL 0 WARN. Report:
   - Paths all present?
   - FindHeaderRow found all N required headers?
   - Approx source row count → output row count ratio within [0.25, 1.10]?
9. **Close.** Report: profile stem, forked-from exemplar, output/source ratio, reconcile exit-code. Offer to run against Excel if user wants to confirm macro compiles.

---

### Form 3 — Create a new transform for Customer Invoice EIB

Identical to Form 2 in structure, archetype swapped.

```
in elt_transform_excel, create new transform for input source <source-ref> to produce customer-invoice EIB [details]
```

The archetype flips:
- Target EIB = `"Submit Customer Invoice_v46.1.xlsx"` (or current version captured in `eib_population_defaults.yaml` under key `"Customer Invoice"`).
- Validator branch = Customer Invoice schema:
  - `fixed_invoice_type` (always required for Customer today; "Invoice" / "Credit Note" / etc.)
  - `fixed_currency` (always required; typically GBP)
  - `per_row_tax` (shape used for FINES subtype; optional for flat Customer Invoice otherwise → check schema for latest)
- Exemplar set = 3 existing Customer profiles:
  - `ar_invoice` = CSV input, fixed company (WNSL / FA fallback), 4 required headers, simple flat → 109 output rows
  - `ar_discipline_invoice` = Excel 1 worksheet, 4 required headers, heavy skip-rule filtering (66→35 rows)
  - `ar_invoice_run` = Excel 1 worksheet, 7 required headers, WNSL company + customer Code col lookup

Ambiguity detection rules are the same as Form 2 (replace "6 Supplier profiles" with "3 Customer profiles"). Deterministic pipeline same as Form 2.

---

## 3. The 8 Required User Prompts (for new Supplier / Customer transforms)

When a user does NOT supply a "zero-follow-up short form" block in their initial prompt, the agent
MUST fire these 8 prompts verbatim (or an equivalent structure-parseable interactive form). Each prompt
provides a default recommendation so a lazy user can click-accept.

### Prompt 1 — Archetype confirm + fork exemplar

> **Ask**:
> `"For this new <Supplier|Customer> Invoice transform — I compared the source header-name token set against all <6|3> existing profiles and the closest match by header-name similarity is: <best_match> (score <Jaccard coeff 0..1>). Fork from this match, OR pick another: <full list grouped by archetype with 1-line rationale>. Which fork?"`

### Prompt 2 — Source layout accept/edit

> **Ask** (all 4 sub-fields are editable):
> `"Proposed source_layout — accept, or edit? Worksheet='<name>' (sheet with most non-empty columns); header_row=<first row with >=4 non-empty cells>; data_start_row=<header_row+1>; validate-presence columns N=<len(required_headers) proposed>: yes|no (show checklist). All proposed values auto-generated from opening the source."`

### Prompt 3 — Company + supplier/customer resolution

> **Ask**:
> `"Company code for every output row: FA / WNSL / <all keys from config/ingest_lookups/ref_company.json> [drop-down]. If source has a Company column with values, I found: <first 5 unique values> — use those instead? Y/N. Default fallback display-name when unresolved: THE FA GROUP - FA | THE FA GROUP - WNSL."`

Supplier Invoice extra in Prompt 3:
> `"Supplier (Reference ID) for every output row: (a) fixed per transform = <lookup from ref_supplier map dropdown>; OR (b) per-row from source column ___; OR (c) supplier spread-sheet lookup using source column ___ → Supplier tab column Reference ID. Which option?"`

Customer Invoice equivalent:
> `"Customer (Reference ID) lookup: same (a)/(b)/(c) using ref_customer map."`

### Prompt 4 — Tax regime (Supplier Invoice: build-gated XOR, never both)

> **Ask**:
> `"Tax — uniform (same on every row) XOR per-row (3 driver source columns). Recommendation: <uniform if no VAT/Tax headers detected in source; else per_row>. Which? If uniform: applicability=____ [dropdown from ref_tax_applicability], code=____ [ref_tax_code dropdown], recoverability=____ [ref_tax_recoverability dropdown]. If per_row: applicability_source_column=____, code_column=____, recoverability_column=____."`

Customer Invoice equivalent (FINES subtype only):
> `"Tax — FINES subtype uses per_row_tax. applicability_col=____, code_col=____, recoverability_col=____. If not FINES confirm fixed values as above."`

### Prompt 5 — Spend category (Supplier Invoice) / Revenue category (Customer)

> **Ask**:
> `"<Spend|Revenue> category: (a) fixed per transform = ____ [dropdown from ref_spend_category / ref_revenue_category]; OR (b) per-row from source column ____; OR (c) Final Code → MAPPING MASTER tab, business object=<Spend Category>, Core vs Prism = Core, lowest hierarchy match, skip+report if unresolved. Recommendation <propose based on exemplar fork>. Which?"`

### Prompt 6 — 9 worktag derivation matrix (Supplier Invoice, 9 repeats; Customer 0 if no worktags yet)

This one prompt carries ~80% of human decision.

> **Ask** (9 times, or one 9-row grid in a structured UI):
> `"Worktag slot <1/9> <worktag_business_objects[i]> — how to derive per output row? (SAME = blank-fixed value) OR (DIRECT = source column ___) OR (FINAL CODE → MAPPING MASTER via business object=<name>, Core vs Prism=Core, unresolved=skip_and_report / abort). Recommendation <wholesale propose pattern from forked exemplar if pattern-match>. Choose SAME|DIRECT|MAP + value/column/map-settings."`

### Prompt 7 — Final Code derivation (conditional)

Fire only if source has a compound accounting string column AND exemplar fork has a `final_code_derivation` block:

> **Ask**:
> `"Final Code sample value = <sample from first 3 rows>. Separators detected = <token>. Proposed segment lengths {seg_lengths: [..]}, seg2_first_char_map {<dict>}. Accept / edit?"`

### Prompt 8 — Target EIB field-mapping close-out (unmatched required fields)

> **Ask**:
> `"Target EIB required fields that fuzzy-name-match missed on source headers: <list>. For each, source column = ___ [dropdown of all source headers]. All can be left blank if that field genuinely isn't populated."`

---

## 4. Zero-Follow-Up Prompt Short Form (power-user)

Users who already know every decision can write a one-shot prompt. The agent MUST accept this form with
ZERO follow-up questions. Template:

```
in elt_transform_excel, create new transform:
  for input source: /full/path/to/source.xlsx
  eib type: supplier-invoice
  profile key: ap_amex_airlines
  fork from: ap_amex_hotel
  source_layout:
    worksheet_name: New CC Updated
    header_row: 9
    data_start_row: 10
    required_headers: [Transaction Account Currency Code, INVOICE NUMBER2, INVOICE DATE, SUPPLIER INVOICE NUMBER, ITEM DESCRIPTION, Transaction Account Amount, Final Code, Journal Name?]
  transform_constants:
    fixed_company_code: FA
    output_company_default_fallback: "THE FA GROUP - FA"
    fixed_supplier_reference_id: AME003
    uniform_tax:
      tax_applicability: GBR_Outside_of_Scope
      tax_code: GBR_Zero_Rated
      tax_recoverability: Nonrecoverable_Allocate
    spend_category_fixed: Staff_Airfares
    mapping_sheet: MAPPING MASTER
    worktag_business_objects:   # must be 9-complete, all supplied here
      - name: Cost Center
        derivation: final_code_map
        core_vs_prism: Core
        on_unresolved: skip_and_report
      # … remaining 8 entries …
    final_code_derivation:
      seg_lengths: [8, 6, 5, 8, 7]
      seg2_first_char_map: {3: H, 4: W, 5: F, 6: M, 7: C, 8: S}
  shell_defaults:
    output_file_path: /full/path/to/output_eib_workbook.xlsx
```

Agent behavior: parse → validate every field against SUPPORTED_EIB_POPULATION_TYPES + the 8-required-prompt
invariant → if any section missing → fall back to the 8 required prompts for that section. If all there →
proceed directly to deterministic pipeline without any questions.

---

## 5. Deterministic Pipeline (summary of close-out steps that run every time)

After disambiguation, before declaring work done, the agent MUST run. For any touch of `.bas` VBA or Runner build, FIRST run the PRE-EDIT GATE from §0.5 (G1+G2 before edit, G3 after edit + before any build step). Then:

| Step | Command / Action | Gate |
|---|---|---|
| Build gate (schema validators 1–4) | `from scripts.build_runner import _load_transform_bindings as L; from pathlib import Path as P; L(P('implementations/elt_transform_excel/config/transforms'))` — no Exception | PASS = 9/9 bindings (now N+1), exit 0. FAIL = report the validator message to user. |
| Reconcile (source ↔ headers ↔ output ratio) | `uv run --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/reconcile_supplier_customer_inputs_outputs.py --profiles <stem> --scan-rows 500 --scan-cols 200` → exit 0; 0 FATAL, 0 WARN | PASS = header found N/N, paths exist, ratio [0.25, 1.10]. FAIL = report reconcile findings. |
| VBA hygiene (MANDATORY if ANY `.bas` module touched this session) | 1) `python3 implementations/elt_transform_excel/scripts/check_vba_hygiene.py --fix` (if NON_ASCII), 2) `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/build_runner.py --hygiene-only`. Then optionally after Excel manual Bootstrap Import, run `uv run … build_runner.py --verify` to confirm macro modules actually loaded. | PASS = 0 ERROR-class violations in both steps, OPTIONAL --verify 9/9 modules FOUND. FAIL = stop everything until hygiene passes (see §0.5 rules list above). |
| Validate input × N touched profiles | `uv run --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/validate_input.py -t <stem>` — once for each profile YAML you touched this session. | PASS = "Full VERDICT: GREEN" / "All checks passed" / 14/14 or equivalent + 0 required-column empty samples. FAIL = report which header could not be resolved on its declared header row. |
| Verify gates (cross-check 9 profiles' 9 config/engine invariants) | `uv run --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/verify_gates.py` → tail -10 lines. | PASS = "9/9 cross-checks: VERIFIED" or equivalent (no FAIL rows). FAIL = report which cross-check ID and which profile. |
| Gap sweep (spec YAML ↔ transforms YAML literal coverage. Any new config tokens in transforms YAML must be reflected in transform_specs or they're orphaned) | `uv run --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/gap_sweep.py` → tail -5 lines. | PASS = "XXX FILLED / 0 GAP". Any GAP means a transform YAML field/section exists but has no decision-record entry in `config/transform_specs/<stem>.yaml`. Fix by appending the missing spec entry. |
| Impl-detail audit (VBA literal strings vs YAML-engine literal tokens) | `uv run --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/impl_detail_audit.py` → tail -8 lines. | PASS = 0 STRING DRIFT / 0 NUMERIC DRIFT. Any DRIFT = a literal string in VBA doesn't match the canonical YAML value (e.g. `"Error"` in VBA YAML `error_token: "Error"`, `"ERROR"` → DRIFT). Fix whichever side is canonical per `config/transforms/<stem>.yaml`. |
| Optional: full-file Final Code derivation audit vs Finance Final Code (Hotel/Trains V2 or any profile with `final_code_derivation` block) | `/tmp/_final_auditor_correct_layers.py` against the real raw source xlsx (scans kept rows only). Run this after any change to DeriveFinalCode macros. | PASS = 100.00% char-for-char equality of engine Final Code vs Finance-precomputed Final Code legacy column on raw source sheet. Trains/Hotels allow 99.88%+ if only case-only ERROR/Error. Any 99.87% or below = 3-tier override misapplication bug, do not ship. |

---

## 6. Reference: 9 Existing Profiles (fork exemplars)

A cold agent should ALWAYS enumerate these 9 first when choosing a fork:

### Supplier Invoice archetype (6 profiles — all target EIB Submit_Supplier_Invoice_v46.1.xlsx)

| Stem | Required headers pattern | When to fork from this |
|---|---|---|
| ap_amex_hotel | 8 cols wide, 2 of them standardized (INVOICE NUMBER2 / Final Code) | Simple case: every coding fixed; uniform tax; 1:1 Final Code→9 worktags map; Supplier ref-id AME001 fixed |
| ap_amex_trains | 9 cols + VAT Percentage | Per-row tax (not uniform); same AMEX-wide headers + VAT-specific |
| ap_amex_trains_v2 | 10 cols + Departure date + Destination airport + per-row VAT | Rich per-row lookup case; if your new source has trip-specific cols this is best match |
| ap_cfa_distributions | 12 cols Finance Master sheet, header row 5, mostly formula | CFA distributions pattern; accounts for Excel formulas still present in raw cells |
| ap_prize_funds | 11 cols INTEGRATION tab, no company column (constant FA enforced) | FA Cup / Prize Fund payment pattern, DocNo = single batch ref for all rows → InvoiceNumber uses BatchID-CreditorID |
| ap_referees | 12 cols Integration tab, company=FA, first cell values FA, REFS 0426 ME, ANT020, INVOICE | Referee expenses: batch_id_token fixed; supplier/coding all per-row |

### Customer Invoice archetype (3 profiles — target EIB Submit_Customer_Invoice_v46.1.xlsx)

| Stem | Required headers pattern | When to fork |
|---|---|---|
| ar_invoice | 4 cols, CSV source, FA/WNSL fallback to FA + WNSL | Generic flat FALL CSV → invoice header/line fields |
| ar_discipline_invoice | 4 cols, 1 worksheet 'Receivables Transactions - Disc', skip-heavy (0.53 ratio) | Discipline fine/charge pattern; many source rows filtered |
| ar_invoice_run | 7 cols, 1 worksheet 'Run 5', WNSL fixed, customer Code column lookup | Invoice-Run pattern (League affiliation, run-number grouping) |

---

## 7. When In Doubt — Escalate (Don't Guess)

If any of these conditions are true, the agent MUST ask (not silently proceed):

1. Source file exists but contains **no worksheet matching** any of the required-headers patterns (after FindHeaderRow 400-row × 200-col scan).
2. The tax-code / company-code / supplier-reference / worktag value user proposed is **NOT present in the tenant ref-ID maps** under `config/ingest_lookups/` — this means either a typo, or Finance hasn't provisioned yet.
3. The new transform would require changing the `SUPPORTED_EIB_POPULATION_TYPES` whitelist in `scripts/build_runner.py` (i.e. new archetype such as "Supplier Setup" / "Customer Setup" / "Journal"). Report to user: this is a 3-step code-change documented inline in the whitelist error message, not config-only — confirm before changing code. *Worked precedent (2026-09-24): the "Supplier Invoice Adjustment" archetype (credit notes) was added exactly this way — (a) name added to `SUPPORTED_EIB_POPULATION_TYPES`, (b) an `eib_populations` block + `required_keys_by_pop` entry in `config/eib_population_defaults.yaml`, (c) an archetype branch in `_validate_transform_constants_schema`. See `ap_cfa_distributions_amendments` (transform + spec + `modApCfaDistribAmendment.bas`) as the template. *Second worked precedent (2026-10-12): the master-data archetypes "Customer" (`Submit Customer`) and "Supplier" (`Submit Supplier`) were added the same way, plus a flat `customer_master_eib_columns` / `supplier_master_eib_columns` validator (`_validate_master_eib_columns`), a Step 6d column-map branch, the hidden `lookups` runtime sheet (`modHelpers.LookupRuntimeValue`) and the registered `entity_row_block` helper (`modEntityBlockWriter`). These types write ONE entity as a block of rows; the `ColumnIndex("XX")` literal gate is unchanged (no table-driven engine).* CFA Distributions splits ONE source (`Finance Master analysis`) into TWO 1-input→1-output transforms filtered by `Invoice type`: `ap_cfa_distributions` (invoices → Supplier Invoice) and `ap_cfa_distributions_amendments` (credit notes → Supplier Invoice Adjustment) — the Runner is one-input-one-output, so multi-output-per-source is done as sibling transforms, not a multi-output macro.*
4. Output/source ratio after reconcile < 0.25 or > 1.10 → flag to user, don't mark reconciled as OK silently.
5. A required_headers token in config exists ONLY in source rows AFTER a macro-internal prep step (i.e. the macro inserts helper columns before calling FindHeaderRow). This case must be called out explicitly: the config's required_headers list then describes POST-PREP headers, not raw source. This is a legitimate pattern but must be tagged as such in the transform_spec to avoid future confusion. Cold agent should: `WARN user + note in spec: 'required_headers are POST macro-prep not raw input; this profile's prep step inserts columns X, Y, Z before FindHeaderRow.'`
