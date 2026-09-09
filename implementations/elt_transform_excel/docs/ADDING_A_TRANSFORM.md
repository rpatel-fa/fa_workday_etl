# Adding a New Transform (Authoring Playbook)

## Purpose

Cold-start guide for adding one source-workbook → EIB transform with minimum human effort.
It captures the repeatable procedure and the non-obvious knowledge from the existing transforms
so a fresh session does not re-derive it. It **points to**, and does not restate, the schema
(`docs/CONFIGURATION.md`), the lookup mechanics (`docs/RUNBOOK.md`), and the build steps
(`docs/BUILD.md`).

**Agent prompt rules:** Before applying this playbook, a cold agent must first read
`docs/AGENT_COLD_START_PROMPT_CONTRACT.md` — it defines the 3 valid user prompt forms, the
ambiguity rules (when the agent MUST fire a clarifying question vs. can proceed silently), the
8 required user prompts (fork exemplar, source layout, company/supplier resolution, tax regime,
spend/worktag derivation, Final Code rules, target-field close-out), the zero-follow-up
short-form syntax, and the mandatory 4-step close-out pipeline (build gate, reconcile,
hygiene, impl-detail). Any prompt that doesn't match a valid form → raise a clarifying question
first, do not guess.

**Three worked exemplars — copy the closest one:**
- `ap_amex_hotel` — the simple case: every domain value (spend category, cost centre, tax) is a
  fixed constant. Start here when one source posts to one coding.
- `ap_amex_trains` — the rich case: per-row cost-centre lookup, per-row VAT selection from a source
  column, cross-sheet reads. Start here when coding varies by row.
- `ap_cfa_distributions` — **the account-string case (the common supplier-invoice pattern): the
  source carries a coding string (e.g. `1077-504007-PNGC0001-0000`) and Cost Centre, Spend Category,
  Reporting Type, and the related worktags are all resolved FROM that string against the finance
  Mapping MASTER** (see "Account-string coding" below). Start here for any AP feed that posts to
  finance codings rather than a fixed cost centre.

## The five artifacts you touch

| # | Artifact | Role |
|---|---|---|
| 0 | `config/transform_specs/<key>.yaml` | The decision record (Step 0) **and** the `mapping:` section — the field-mapping source of truth the `.bas` mirrors. Not read by the build. |
| 1 | `config/transforms/<key>.yaml` | Build profile: label, source/EIB types, macro name, source layout, shell defaults. |
| 2 | `vba_src/transforms/mod<Name>.bas` | Executor. **Hand-authored** to mirror the spec's `mapping:` section (co-located text diff, no hash gate) — there is no `spec → .bas` generator; see `docs/ARCHITECTURE.md` "Authoring model". |
| 3 | build prep (`build_runner.py`) | Writes the registry/config from the profile YAML. |
| 4 | manual Excel import | `BootstrapImporter.BootstrapImportFromConfiguredSource` (see `docs/BUILD.md`). |

## Step 0 — Fill the transform spec (the decisions) BEFORE authoring

This is the biggest time saver and the difference between a session that builds unattended and one
that interviews you. Capture the decisions the source data **cannot** answer into a spec file at
`config/transform_specs/<key>.yaml` (kept out of `config/transforms/`, which the build globs as
runnable profiles). A filled spec is both the session's input and an auditable record of who chose
each coding. Getting these wrong produces a Workday-bound file that is silently miscoded.

**Worked example:** [`config/transform_specs/ap_amex_trains.yaml`](../config/transform_specs/ap_amex_trains.yaml).

**Blank template:**

```yaml
# Identity, EIB population type, template/source/output paths, and source layout (worksheet,
# header row, data-start row, required_headers) are NOT recorded here — they live in the build
# profile config/transforms/<key>.yaml (Step 4). This spec holds decisions + the mapping only.
meta: {key: , label: , requester: , date_confirmed: , exemplar_to_fork: }   # hotel=simple, trains=lookups
parties:
  supplier: {source_code: , resolved_via: , key_column: , value_column: }
  company:  {source_code: , resolved_via: }
fields:                         # one entry per EIB field that is NOT a fixed house default
  invoice_number: {rule: , note: }          # unique per supplier or Workday rejects duplicates
  invoice_date:   {source_column: , format: }
  item_description: {source_column: }
  memo:           {source_column: }
  line_amount:    {source_column: , basis: }   # gross | net
  reporting_type: {value: }     # OR function: reporting_type_from_final_code (account-string feeds)
  # spend_category + cost_centre + related worktags: a FIXED value: for simple feeds, OR a per-row
  # account-string lookup for coding-string feeds. Do NOT assume Spend Category is fixed — on an
  # account-string feed it is resolved from the coding string exactly like Cost Centre (see the
  # "Account-string coding" section). Repeat this block per coding-driven field.
  spend_category:               # fixed value:, OR a per-row account-string lookup:
    derivation:                 # fixed | per_row_lookup
    key_source_column:          # the source coding-string column
    lookup_table: {location: Mapping Spreadsheet, tab: MAPPING MASTER, business_object: Spend Category, key_column: OLD CODE, value_column: WORKDAY REF ID}
    on_unresolved:              # skip_and_report | abort
  cost_centre:                  # fixed value:, OR a per-row account-string lookup:
    derivation:                 # fixed | per_row_lookup
    key_source_column:
    lookup_table: {location: Mapping Spreadsheet, tab: MAPPING MASTER, business_object: Cost Center, key_column: OLD CODE, value_column: WORKDAY REF ID}
    on_unresolved:              # skip_and_report | abort
tax:
  mode:                         # uniform | per_row
  driver_source_column:         # only if per_row
  sets:                         # tenant-specific Workday IDs — appear in NO source file; finance must supply
    - {when_rate: , applicability: , code: , recoverability: }
skip_rule: {drop_when: [ ]}
open_questions: [ ]             # must be empty before authoring begins
```

Field notes — these are the questions finance answers, mapped onto the spec:

- **EIB population type** and template path (e.g. Supplier Invoice vs Customer Invoice) — recorded
  in the profile (`transform.eib_population_type`, `shell_defaults.target_eib`) at Step 4.
- **Source worksheet**, header row, data-start row, and the handful of `required_headers` that
  uniquely identify the header row (the macro finds the header by content, so leading title/blank
  rows are tolerated) — verified in Steps 1–2 and recorded in the profile's `source_layout` (Step 4),
  not in this spec.
- **Supplier / Company** resolution — fixed reference code (e.g. `AME001`) + which extract resolves it.
- **Line amount** — gross vs net, and which source column. (Net when Workday calculates tax; see below.)
- **VAT / tax** — is it uniform or per-row? If per-row, which source column drives it, and the
  **exact Workday tax reference IDs** for each rate (applicability / code / recoverability). These
  are tenant-specific and appear in no source file — you must be given them.
- **Cost Centre** — fixed, or derived per row? If derived, which source column is the key and which
  reference table resolves it to a Workday `Cost_Center_Reference_ID`.
- **Spend Category** — the fixed `Spend_Category_ID`.
- **Invoice Number** — must be unique per supplier in Workday. If two transforms share a supplier
  (hotel and trains both use `AME001`), give them distinct invoice-number tokens or Workday rejects
  the second as a duplicate.
- **Invoice Date** — which source column and its format (`DD/MM/YYYY` text vs a real date).
- **Skip rule** — which rows are dropped (missing key fields, zero amount, unresolved lookup) and
  whether an unresolved lookup should skip-and-report or abort the run.

**With a complete spec (`open_questions` empty), Steps 1–6 are mechanical** — a session executes them
without further finance input. Use Steps 1–2 to *verify* against the real files (sheet/header positions,
which land in the profile's `source_layout` at Step 4; that the spec's named source columns exist; that
the EIB column letters resolve to the field names in the spec's `mapping:`); if a check fails or a
decision is missing, stop and add it to `open_questions` rather than guessing. Everything a session
must not invent (tenant tax/worktag IDs, gross-vs-net, which date column) lives in the spec by design.

## Step 1 — Inspect the source workbook

```bash
uv run --with openpyxl --no-project python -c "
import openpyxl; from openpyxl.utils import get_column_letter
p='<SOURCE.xlsx>'
wb=openpyxl.load_workbook(p, read_only=True, data_only=True)
for ws in wb.worksheets: print(ws.title, ws.sheet_state, ws.max_row, ws.max_column)
ws=wb['<DATA SHEET>']
for i,row in enumerate(ws.iter_rows(min_row=1,max_row=3,values_only=True),1):
    print('row',i,[(get_column_letter(j+1),v) for j,v in enumerate(row) if v is not None])
"
```

Identify: the data sheet, the header row, and any finance-maintained lookup tabs *inside* the source
workbook that the transform needs alongside any external mapping workbook lookups.

## Step 2 — Decode the EIB target columns (reusable across every EIB of that type)

**⚠️ Identity check BEFORE you start decoding.** Two EIB xlsx files can carry the same
Workday release version string (e.g. `v46.1`), share the 5 scaffolding columns A/B/C/BQ/CI,
and still be *completely different business objects* with <3% semantic overlap in the remaining
columns. A classic trap: "this looks fresher and has sample data, let me wire it in instead."
Before assigning any xlsx as `target_eib` for a transform, verify all three:

1. **Row 1 banner verbatim** — it must name the exact business process you are writing for:
   - Supplier Invoice → banner must contain `Submit Supplier Invoice`
   - Ad Hoc Payment → banner must contain `Submit Ad hoc Payment`
   - Customer Invoice → banner must contain `Submit Customer Invoice`
2. **Worksheet tab name** — matches the banner title.
3. **3 non-adjacent header semantics from your spec.** E.g. for any AP Supplier Invoice
   transform: confirm col `O` Row 5 = `Supplier*` AND col `DP` = `Spend Category` AND
   col `EQ` = `Cost Centre`. If any of the three don't match, stop — that is NOT the right
   template.
   - Real-world trap discovered against `Submit Ad hoc Payment - v46.1`: 104 cols vs 176 for
     Supplier Invoice, col `K` = `Bank Account*` (no Supplier equiv), col `L` = `Payee`,
     col `O` = `Single Use`, NO `Supplier*` field anywhere, worktags live at `AY..BF` (not
     `DP..EX`) and there is NO Project column. Wiring ap_stationery / ap_taxi (or any other
     AP Supplier Invoice transform) to that xlsx writes Spend Category into Qty cells, amounts
     into description cells, etc. — silent 100% data-layout corruption, Workday rejects every
     row. Do NOT "fix by rearranging the WriteOutputRow" — use the correct EIB for the business
     process, or if it genuinely *is* Ad Hoc Payment, create a whole new transform key/spec/.bas
     with its own column map.

Mapping `target` values are the EIB's **column letters, lower-cased** (`eq` → column `EQ`).
Decode them from the template's header rows — **row 4 = type / reference-ID format, row 5 = field
name; data starts row 6**:

```bash
uv run --with openpyxl --no-project python -c "
import openpyxl; from openpyxl.utils import get_column_letter
wb=openpyxl.load_workbook('<EIB_TEMPLATE.xlsx>', data_only=True)
ws=wb.worksheets[0]
for ci in range(1, ws.max_column+1):
    name=ws.cell(5,ci).value; typ=ws.cell(4,ci).value
    if name: print(get_column_letter(ci), '|', str(typ).strip(), '|', str(name).strip())
"
```

### Reusable reference — the Supplier Invoice EIB field block

The 21 columns every AMEX-style Supplier Invoice transform populates, in **`Submit_Supplier_Invoice_v46.2`**
column letters. Copy this block from a **v46.2-migrated** exemplar (`ap_taxi` / `ap_cfa_distributions`)
and change only the deltas.

> **LANDMINE (v46.1 → v46.2 migration):** several field names appear **more than once** in v46.2. The
> line-level `Extended Amount` (`GF`) and `Memo` (`GQ`) live in the *Invoice Line Replacement Data+*
> group; identically-named `IK`/`IL` belong to the *Supplier Invoice Split Line Data+* group — writing
> there populates the wrong sub-grid. Likewise the line `Row ID*` is `EK` (not `DD` = Manual Override
> Percent), `Invoice Received Date` is `AG` (not `AR` = Hold Reason), `External PO Number` is `BH` (not
> `BB` = Down Payment). Always resolve an EIB column by its **group + field name**, never by field name
> alone. The not-yet-migrated `ap_amex_*`, `ap_prize_funds`, `ap_travel_places` modules still write the
> old `IK`/`IL` slots — do not copy them.

| Col | EIB field (type) | Typical source |
|---|---|---|
| `B` | Spreadsheet Key (Text) | `function=sequence` |
| `I` | Submit (Y/N) | `default=Y` |
| `K` | Invoice Number (Text) | `function=invoice_period_from_filename` + `source_value` suffix |
| `R` | Company (Company_Reference_ID) | `this`/`mapping` entity=company lookup on `FA` |
| `T` | Currency (Currency_ID) | source currency column / `default=GBP` |
| `U` | Supplier (Supplier_ID) | `source_value` code (e.g. `AME001`) via `Supplier Spreadsheet` lookup |
| `AF` | Invoice Date (YYYY-MM-DD) | source date column + `source_format` |
| `AG` | Invoice Received Date (YYYY-MM-DD) | receipt/received date; fallback = Invoice Date |
| `BG` | Suppliers Invoice Number (Text) | same as `K` |
| `BH` | External PO Number (Text) | source PO column (optional) |
| `EK` | Row ID (line) (Text) | `default=1` |
| `EP` | Item Description (Text) | source description column |
| `EW` | Spend Category (Spend_Category_ID) | fixed `default`, **OR** account-string lookup (business object `Spend Category`) — see below |
| `FE` | Tax Applicability (Tax_Applicability_ID) | fixed or `vat_rate` |
| `FF` | Tax Code (Tax_Code_ID) | fixed or `vat_rate` |
| `FH` | Tax Point Date Type | `default=INVOICE/ADJUSTMENT_DATE` |
| `FK` | Tax Recoverability 1 | fixed or `vat_rate` |
| `FL` | Tax Option 1 | `default=CALC_TAX_DUE` (Workday calculates tax from the code) |
| `GE` | Unit Cost (Number) | source amount column |
| `GF` | Extended Amount (Number) — *line, not split `IK`* | source amount column |
| `GQ` | Memo (Text) — *line, not split `IL`* | source narrative column |
| `GR` | Reporting Type (Custom_Organization_Reference_ID) | fixed `default` or `function=reporting_type_from_final_code` |
| `GT` | Cost Centre (Cost_Center_Reference_ID) | fixed, or account-string lookup (business object `Cost Center`) |

**Account-string worktag family (add these when the source carries a coding string, as
`ap_cfa_distributions` does).** Each is the same account-string lookup as `eq`, differing only by
the `WORKDAY BUSINESS OBJECT` filtered on. A code that doesn't match a given object leaves the cell
blank — that is correct, not an error.

| Col | EIB worktag | WORKDAY BUSINESS OBJECT |
|---|---|---|
| `GV` | Event | `Events - Custom Org` |
| `GX` | Camp | `Camps - Custom Org` |
| `GZ` | FA Learning Course | `FA Learning - Custom Org` |
| `HB` | Commercial Partner | `Commercial Partners - Custom Worktag` |
| `HD` | Grant | `Grant` |
| `HF` | Programme | `Program - Delivered Org` |
| `HH` | Project | `Projects (Opex)` |

### Account-string coding (Cost Centre, Spend Category, worktags, Reporting Type)

Most AP supplier-invoice feeds carry a **coding string** — segments joined by `-`
(`1077-504007-PNGC0001-0000`). Every coding-driven EIB field is resolved from that string against
the finance **Mapping MASTER** (path in `config/reference_paths.yaml`), by the shared VBA helper
`lookup_worktag_from_final_code` / `BuildBusinessObjectHierarchyLookup`. The rule, identical for
Cost Centre, Spend Category, and every worktag:

1. Split the coding string on `-`; try each segment as an `OLD CODE`.
2. Keep only Mapping MASTER rows where **`Core vs Prism` = `Core`** (whitelist Core; a future third
   value never leaks in) and `WORKDAY BUSINESS OBJECT` = the field's object.
3. When more than one segment matches, take the **lowest / most specific hierarchy** (highest
   `OLD SEGMENT`). Output the row's `WORKDAY REF ID`.

**Reporting Type** is derived instead from the first character of accounting element 2
(`function: reporting_type_from_final_code`). **Ledger Code is a GL-journal concern only — never
populate it on supplier/customer invoices.**

The single trap to avoid: do **not** inherit a *fixed* Spend Category from the AMEX exemplars for an
account-string feed — Spend Category is resolved from the coding string exactly like Cost Centre.

## Step 3 — Author the mapping (in the spec)

Add a `mapping:` list to `config/transform_specs/<key>.yaml` — one item per target EIB column, keyed
by `target`. The row vocabulary and every `lookup.file`/`function` token are defined in
`docs/CONFIGURATION.md` ("Mapping Contract"). The five row shapes:

| Intent | How |
|---|---|
| Fixed constant | `default:` only |
| Direct source value | `source_column:` |
| Fixed literal resolved through a reference extract | `source_value:` + `lookup: {file, tab, input_column, output_column}` |
| Per-row live lookup | `source_column:` (the key) + a `lookup:` block; `lookup.file: Source Data` reads a tab inside the source workbook |
| Account-string worktag lookup | `source_column:` (the coding string) + `function: lookup_worktag_from_final_code` + a `lookup:` block filtered to one `WORKDAY BUSINESS OBJECT` and `Core vs Prism` = `Core` (Cost Centre, Spend Category, worktags — see "Account-string coding") |
| Derived value | `function:` (`sequence`, `invoice_period_from_filename`, `vat_rate`, `reporting_type_from_final_code`) |

Copy the whole `mapping:` block from the closest exemplar spec and change only the deltas. Quote any
value containing `?`, `|`, or `:` (e.g. `"Journal Name?"`) so it survives YAML flow-formatting.

## Step 4 — Write the build profile

Create `config/transforms/<key>.yaml`, modelled on `ap_amex_trains.yaml`: unique `key` and `label`,
`transform.source_type` / `eib_population_type`, `runtime.step_list_macro`, `source_layout`, and
`shell_defaults` — set only the transform-specific `source_data`, `target_eib`, and a distinct
`output_file_path`.

**Before you write `shell_defaults.target_eib`, re-run the Step 2 identity check against that
exact file.** Row 1 banner + tab name + 3 non-adjacent column semantics. If the file matches a
different business process but shares the same version number, STOP — do NOT override the
archetype by rearranging WriteOutputRow; assign the correct `eib_population_type` and use the
correct template for that business process, or create a new transform key/spec/.bas for the
other process. Specific gotchas:
- Never use `Submit Ad hoc Payment*` xlsx as `target_eib` for any AP Supplier Invoice transform
  (ap_amex_*, ap_cfa_distributions, ap_referees, ap_prize_funds, ap_stationery, ap_taxi). The
  worktag block and amount block are at completely different column addresses and Workday will
  reject every row.
- Conversely, never use a Supplier Invoice template to populate an Ad Hoc Payment load.

**Do not set `customer_spreadsheet` / `supplier_spreadsheet` /
`mapping_spreadsheet` / `mapping_sheet`** — those are inherited from `config/reference_paths.yaml`
(set an explicit value only to override the shared default). No mapping/hash keys — the mapping lives
in the spec.

## Step 5 — Write the macro

The `.bas` is ~80% boilerplate shared verbatim across transforms and ~20% transform-specific.

- **Copy verbatim** from the closest exemplar: `ColumnIndex`, `FindHeaderRow`, `ValidateHeaderRow`,
  `FindColumnIndexInRow`, `RowContainsHeader`, `NormalizeToken`, `ResolveAmount`,
  `NormalizeSourceDate`, `ClearPublishedRows`, `ResolveCompanyId`, and the run/HandleFailure
  scaffold.
- **Change only**: the constants block (sheet names, header rows, fixed EIB values), the column
  index list, `WriteOutputRow` (which target cell gets what), and any transform-specific lookup
  (e.g. `BuildCostCenterLookup`, `ResolveVatRate`/`Tax*ForRate`).
- Before creating any new finance mapping or derivation helper, check `modFinanceCommon.bas` first.
- If the function is a reusable finance mapping/derivation or common lookup helper, add it to the
  shared finance-domain helper module under `vba_src/` (for example `modFinanceCommon.bas`) by
  default instead of creating it locally first.
- Before copying local helpers such as invoice-period derivation, delimited customer/supplier
  reference lookups, or unresolved-lookup summary formatting, check `modFinanceCommon.bas` first
  and call the shared helper if it already exists there.
- Keep only transform-specific source-layout handling, target-column writing, and one-off business
  exceptions in the transform module itself.
- **Comment placement discipline — single source of truth.** A transform `.bas` file is a pure
  runtime mirror of its corresponding `config/transform_specs/<key>.yaml`. It must carry exactly
  one comment line: a pointer of the form
  `' Source of truth: config/transform_specs/<key>.yaml — mapping SOT the .bas mirrors BY HAND (no auto-regen); build_runner enforces spec<->code column agreement`.
  All prose, rationale, skip-rule documentation, and mapping-table descriptions live in the YAML
  spec. *Only* shared engine modules (`modHelpers`, `modFinanceCommon`, `modSourceTable`,
  `modDispatcher`) retain gotcha comments — those encode institutional knowledge about VBA/Excel/
  Mac sandbox behaviour that is not present in any spec file and cannot be regenerated.
- The public entry point must be `Run<Name>(ByVal runnerWorkbook As Workbook, ByVal controlSheet As
  Worksheet)` and match `runtime.step_list_macro` in the YAML — the dispatcher resolves it from the
  registry by name.
- Fixed domain values live as constants that mirror the spec's `mapping:` section (reviewed together
  in one diff); only genuinely per-row-varying values need a live lookup.
- Write the constants block *from* the corresponding YAML/spec, never by eye against a live workbook.
  Shared mapping constants (`MAPCOL_*` header names, `MAPVAL_CORE_ONLY` filter value) live in
  `modFinanceCommon.bas` and are updated when the Mapping MASTER file contract changes — the agent
  reads the contract description in `config/reference_paths.yaml` and emits the matching constants.
  Per-transform `WBO_*` business-object constants mirror the profile's
  `worktag_business_objects` / `fixed_literal_worktags` lists.
- **Archetype guard before you write:**
  - If the transform archetype is *Supplier Invoice (AP)*, remove any stray `WBO_REVENUE_CATEGORY`
    constant and do not write to Revenue Category cells.
  - If the transform archetype is *Customer Invoice (AR)*, remove any stray `WBO_SPEND_CATEGORY`
    constant and do not write to Spend Category cells.
  Do not treat these as "mapping bugs" to work around — they are categorical payload constraints.
- **For truly-fixed per-slot ref ids (e.g. AMEX always assigns the same category):** declare the
  slot under `fixed_literal_worktags:` in the profile YAML, NOT under
  `worktag_business_objects:`. `worktag_business_objects:` triggers a hierarchy lookup against
  the Mapping MASTER (segments, Core vs Prism filter, lowest-specificity wins); a constant there
  is a category error that looks like an "unresolved lookup" at runtime.
- **For AR account-string transforms (AR Bulk / AR Multi-Line / future ones):** write mapping
  results exactly to the 11-cell EIB layout `KE` (RevCat) · `GB` (Project Header) · `MN`
  (Project Line) · `MV`..`NC` (8 consecutive Worktag slots). Do not invent new column
  assignments. AR transforms default *unresolved* lookups to **blank** — do NOT reject the row,
  do NOT write `"N/A"` or flag text. AP transforms reject required misses (Cost Centre etc.);
  AR does not.
- **Never hardcode file paths in a `.bas` string literal.** Read template/reference paths via
  `ReadControlValue(controlSheet, CONTROL_LABEL_*)`. If a path is wrong or changed, update the
  YAML (`reference_paths.yaml` for shared, `config/transforms/*.yaml` for override) and rebuild
  the Runner; do not edit a `.bas`.
- **Do not use `CDec`.** Mac Excel VBA has no working analogue and values silently degrade. Use
  `Format$(value, "0.00")` (or equivalent string formatting) for decimal renders and do the
  actual math in scaled integer arithmetic where it matters.
- **`ColumnIndex("XX")` allowlist.** Every two-letter column argument must be on
  `modSourceTable.bas`'s `SAFE_SHARED_COLUMNS` list OR declared in the profile's `eib_columns`.
  Transform-specific EIB cells are referenced as `ColumnIndex(EIB_COL_GB)` via constants at the
  top of the `.bas` — never as the bare string.
- **Shell-formula banner offset (when editing `create_runner_shell.py`).** Any formula added to
  the generated `Runner` control sheet (e.g. auto-populated path cells) must use the actual
  post-banner row numbers, not the design-time row number. A 1–2 row off-by-one here leaves
  every such cell as `#N/A` on the built shell; rebuild + open-check to confirm.
- **Stage 1 anti-drift — declare the shared-helper variant selection in the transform spec BEFORE writing the .bas.**
  *This contract is mechanically enforced by `scripts/build_runner.py` Step 6; missing it will fail hygiene.*
  1. Open `config/shared_helper_registry.yaml`. It defines 6 concerns and the valid variant keys for each.
     Read the variant's `vba_module` + `vba_public_fn` + `requires_columns` metadata to decide which one matches the source-export shape.
  2. In `config/transform_specs/<base_profile_key>.yaml`, insert a `shared_helpers:` block AFTER `meta:` and BEFORE `parties:`.
     The block must pick exactly one variant per concern (all 6 are required). Use one of the existing 9 AP transform specs as your exemplar:
     - Classic 6 (taxi/travel/stationery shape): `per_row_split_company_creditor` + `per_row_split_from_supplier_id` + `direct_column_separator_normalized` + `standard_9pack_lowest_specificity` + `effective_net_gross_tax_fallback` + `text_cols_and_amount_cols_parameterized`.
     - Direct-distribution (CFA/PrizeFund/Referees shape): Swap `company_resolution` to `per_row_integration_column_a` or `fixed_literal`; swap `accounting_string_derivation` to `direct_final_code_column`; swap `row_net_resolution` to `direct_amount`.
     - AMEX 3-pack: `fixed_reference_id` + `amex_padded_22char_pipeline` or `direct_final_code_column` + `none` for `row_trailer_detection` (processor exports have no trailer rows — don't check for them).
  3. In the transform `.bas`, call the hoisted shared function by **bare name** — no module qualifier. Public functions in `modSourceTable.bas` / `modFinanceCommon.bas` resolve via standard-module project scope automatically. Do NOT qualify with `modSourceTable.` — it's not needed and future hoisting re-orgs will break the calls.
     Hoisted helpers you should be calling bare:
     - Row checks: `RowIsTrulyEmpty(sourceTable, rowIndex, Array(col1, col2), Array(col3, col4))` / `RowIsEmptyOrTrailer(sourceTable, rowIndex, textCheckCols(), amountCheckCols(), labelCheckCols())` — use the Array() form in ALL new code (the 8-arg positional signatures are dead code from pre-hoist).
     - Net resolution: `ResolveEffectiveNet(netCell, grossCell, taxCell)` instead of the 5-line inline fallback block.
     - Accounting-string → worktag lookup: pass the **RAW** source cell to `ResolveLowestHierarchyLookupAcrossFinalCodeSegments(...)` / `FinalCodeSegment(...)`. Separator normalisation (hyphen / space / dot / `_ | : / \ ; ,`) is done **inside** those shared resolvers, at both the reject gate and the write phase — do NOT pre-normalise the code for lookup, and do NOT re-implement it per module (see `shared_helper_registry.yaml` → `accounting_string_derivation.direct_column_separator_normalized.normalization_owner`). Call `NormalizeCostCodeSeparators(rawAccountingStringCell)` yourself ONLY when you need a normalised string for a **non-lookup** purpose (e.g. the line Memo).
  4. Do NOT declare ANY local `Private Const WBO_*`, `Private Const TAX_APPLICABILITY_ZERO/STANDARD TAX_CODE_*_ZERO/STANDARD TAX_RECOVERABILITY_*_ZERO/STANDARD`, or any `Private Function NormaliseToWordBoundaries / NormalizeCostCodeSeparators / RowIsTrulyEmpty / RowIsEmptyOrTrailer / ResolveEffectiveNet` in a new AP `.bas`. Step 6 rejects them as hoist-regressions (NEG guard).
  5. Run `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/build_runner.py --hygiene-only --runner-shell /tmp/fake.xlsm --vba-source elt_transform_excel/vba_src --profile-dir elt_transform_excel/config/transforms --output /tmp/out.xlsm` after authoring. It must return `HYGIENE OK — 0 ERROR-class violations` before you commit. The 4 sub-checks it runs against your new module:
     - (1) spec EXISTENCE — your module has exactly ONE transform_spec with `shared_helpers:`;
     - (2) registry VALIDITY — every declared concern/variant key resolves in the catalog;
     - (3) NEG guard — zero local Private Const/Function copies of hoisted names;
     - (4) POS guard — for variants != `"none"`, the declared helper bare names are referenced at call sites in your `.bas`.
  *If you are adding a NEW shared concern or a NEW variant (not just selecting an existing one):*
  first add the new key to `config/shared_helper_registry.yaml` with vba_module/vba_public_fn/requires_columns/examples metadata, then add any new POS-guard hint entries to `scripts/build_runner.py → CONCERN_VARIANT_TO_BARE_HELPER_HINTS` inside `_validate_vba_source_hygiene`. New hint entries MUST be confirmed by grepping an existing transform `.bas` body that actually calls the helper — do NOT include transitive-called internal helpers (see ARCHITECTURE.md § Non-Obvious Gotchas #3).
- **For AR Customer Invoice transforms:** Stage 1 AP hoisting rules do NOT apply (AR is out of scope). Write AR modules using the existing legacy Private Const pattern for WBO_* and TAX_*. Step 6 preserves the original negative guard for AR until the AR EIB template parity is reached and Tier 1 hoisting is re-run for the 5 AR files.

### Source-file hygiene (non-negotiable — Mac Excel VBA import is fussy)

Before committing a new `.bas`, run these checks and fix anything that fails. A file that looks fine
in a text editor will still silently fail `VBComponents.Import` on Mac Excel if it violates any of
these, falling back to a blank `Module1` placeholder (see `docs/BUILD.md` "Runner workbook" step 3
and `docs/RUNBOOK.md` for the recovery path):

1. **ASCII-only content.** Multi-byte UTF-8 (em dashes `—`, box-drawing `─`, arrows `→`, curly
   quotes, etc.) inside *even comment blocks* causes the parser to abort when present in volume
   (pure-ASCII always works; ≤ ~20 bytes total like two em-dashes in `modHelpers.bas` is tolerated;
   hundreds of bytes reliably fails). Replace every non-ASCII codepoint with ASCII equivalents:
   `—`→`--`, `─`→`-`, `→`→`->`, curly quotes → straight quotes. Verify:
   ```bash
   python3 -c "d=open('vba_src/transforms/mod<Name>.bas','rb').read(); assert max(d or [0]) < 128, f'{sum(1 for b in d if b>=128)} non-ASCII bytes'"
   ```
2. **CRLF line endings, never LF-only.** `Attribute VB_Name = "mod<Name>"` on line 1 is only parsed
   correctly with CRLF. Convert: `sed -i '' 's/$/\r/' vba_src/transforms/mod<Name>.bas`, then
   confirm with `file vba_src/transforms/mod<Name>.bas` reporting "with CRLF line terminators".
3. **No individual lines > 200 characters after joining line-continuations is not required**, but
   *raw* lines > ~250 chars have been observed to cause mid-parse failures. Use line-continuation
   (` _`) to keep source lines comfortable; a 200-char soft ceiling avoids the case.
4. **Balanced procedure terminators.** Every `Sub` / `Function` must have a matching `End Sub` /
   `End Function`. `Debug -> Compile VBA Project` inside Excel after import catches this too, but
   it's cheaper to check by hand.
5. **Module name <= 31 characters.** VBA has a hard 31-char limit on component (module) names. If
   `Attribute VB_Name` (and therefore the `mod<Name>.bas` filename) exceeds it, `VBComponents.Import`
   silently drops the name, keeps the blank `Module1` placeholder, and then fails the rename back --
   surfacing as run-time error `-2147219692 (80040714)` ("Imported 'Module1' ... kept the VBA
   component name"). This is a name-length limit, not a Mac quirk; it is checkable statically before
   any import. The `Run<Name>` Sub name is *not* subject to this limit (it is invoked by string via
   `Application.Run`), so shorten the module name only -- keep the Sub/`step_list_macro` name intact.
   Verify:
   ```bash
   python3 -c "import os,glob; bad=[f for f in glob.glob('vba_src/transforms/*.bas')+glob.glob('vba_src/*.bas') if len(os.path.basename(f)[:-4])>31]; assert not bad, bad"
   ```

Do not rely on the build to catch these for you. `build_runner.py` writes the registry but does
not parse the `.bas` files; the VBA parser inside Excel does that, and once it silently falls back
to `Module1` you only find out after running the manual import step.

## Step 6 — Build and import

Follow `docs/BUILD.md`. `build_runner.py` writes the registry/config from the profile YAML; the module
import is the one manual Excel step.

## Scaling note (tens of inputs)

The transform-specific surface is small (constants + column list + `WriteOutputRow` + occasional
lookup); the rest is duplicated boilerplate. Below ~5 transforms, copy-and-change is cheaper than
abstraction. As the count grows, a table-driven engine that executes the spec `mapping:` directly
(instead of hand-mirroring it per module) becomes worth a dedicated design decision — treat that as
a governance-reviewed change, not an incremental edit. A measured review of the current duplication
and a tiered simplification plan is recorded in
[`docs/todo/simplify-transform-boilerplate.md`](todo/simplify-transform-boilerplate.md).
