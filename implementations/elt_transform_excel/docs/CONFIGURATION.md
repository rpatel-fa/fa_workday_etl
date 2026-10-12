# Configuration

## Purpose

This document defines the configuration surface for `elt_transform_excel`. There are now five
distinct layers:

- Runtime operator configuration entered through the Runner workbook control sheet (labeled
  input/output rows — see Control-Sheet Inputs below).
- Runtime reference data in the Runner workbook's `mapping` worksheet (entity/source/target lookup
  rows, not labeled control-sheet fields) — read live at run time, no rebuild required to change it.
  See `docs/RUNBOOK.md`'s "Mapping Lookup Workflow".
- Build-time developer profiles stored as YAML under `config/transforms/`, used by Python tooling to
  seed the generated shell with scenario defaults and register which macro a selected transform runs.
- The per-transform field mapping (target EIB column → source) and the human decision record live in
  `config/transform_specs/<key>.yaml`; its `mapping:` section is the authoring source of truth the VBA
  module mirrors. Not read by the build — reviewed as a normal text diff. See "Mapping Contract" below.
  The spec holds decisions + mapping only; it must **not** restate identity, EIB population type,
  file paths, or source layout — those are owned by the build profile above, so a single edit changes
  them in one place and the two files cannot drift.
- The definition of an **input** template workbook that a finance team populates by hand lives in
  `config/ingest_specs/<key>.yaml`. Not read by the build. See "Ingest Specs" below.

## Configuration Model

The shared engine stays domain-free. Runtime choices and file inputs come from the Runner control sheet,
while domain-specific transform rules live in per-combination modules (see `docs/ARCHITECTURE.md`'s
"Config purity" section for where each layer lives). Python build tooling reads YAML profiles both to
pre-populate the shell for a known transform scenario and to register which step-list macro a selected
transform profile should execute. The shipped workbook does not read YAML directly at runtime; instead
the build writes a visible transform-registry block into the `Runner` sheet plus a visible `config`
sheet containing the editable `shell_defaults` values, and the control-sheet formulas read from those
generated workbook tables.

## Build-Time YAML Profiles

Current locations:

- `implementations/elt_transform_excel/config/transforms/ar_invoice.yaml`
- `implementations/elt_transform_excel/config/transforms/ar_invoice_run.yaml`
- `implementations/elt_transform_excel/config/transforms/ar_discipline_invoice.yaml`
- `implementations/elt_transform_excel/config/transforms/ap_amex_hotel.yaml`
- `implementations/elt_transform_excel/config/transforms/ap_amex_trains.yaml`

Current purpose of a profile under `config/transforms/`:

- Define a transform profile key and label.
- Define the default shell values for that transform scenario.
- Define the sample file paths used for that scenario.
- Define the macro that the built Runner should call directly for the selected transform profile.

Every key currently read by the build tooling (see `docs/ops/CLI.md` for the commands):

| Key | Required | Meaning |
|---|---|---|
| `transform.key` | Yes | Unique profile identifier written into the Runner's transform registry and used at runtime as the transform selection key. |
| `transform.label` | Yes | Human-readable profile label shown in the Runner's `Transform` dropdown. Labels must also be unique because the control sheet resolves the selected row in the registry by label. |
| `transform.source_type` | Yes | Auto-populates the Runner's `Source Type` field from the selected transform profile. |
| `transform.eib_population_type` | Yes | Auto-populates the Runner's `EIB Population Type` field from the selected transform profile. Must be one of `SUPPORTED_EIB_POPULATION_TYPES` (`Supplier Invoice`, `Customer Invoice`, `Supplier Invoice Adjustment`, `Customer`, `Supplier`) and have an entry in `config/eib_population_defaults.yaml`. The master-data types `Customer` (-> `Submit Customer`, clear to column `NT`) and `Supplier` (-> `Submit Supplier`, clear to `RH`) are written as row blocks by `modEntityBlockWriter`; their profile must declare `transform_constants.customer_master_eib_columns` / `supplier_master_eib_columns` (flat `field: column letters` map, validated by `_validate_master_eib_columns`) and no worktag/tax pipeline. Added 2026-10-12; `ar_customer_master` uses `Customer` and `ap_supplier_master` uses `Supplier` (both 2026-10-12). |
| `transform_constants.customer_master_eib_columns` / `supplier_master_eib_columns` | Customer / Supplier types only | Flat map of field name -> EIB column letters (1-3 uppercase letters, quoted). Every `ColumnIndex("XX")` literal the master transform writes must appear here or in the type's `safe_shared_columns` (Step 6d). |
| `transform.start_row` | No | Written into the registry's "Start Row" column for reference. **Vestigial:** dates from the dropped step-list/Interim-workbook design; no current transform macro reads it (see the "Architecture simplification" Done entry in the anchor doc). Safe to omit for new profiles. |
| `transform.description` | No | Free-text profile description; not surfaced in the shell or registry. |
| `runtime.step_list_macro` | Yes | The name of the transform macro to call directly for this profile — the macro does the whole run itself, not an ordered step list. |
| `shell_defaults.source_data` | Yes | Default value written into the workbook `config` sheet and auto-populated into the `Source Data` control-sheet cell when this transform is selected. |
| `shell_defaults.target_eib` | Yes | Default value written into the workbook `config` sheet and auto-populated into the `Target EIB` control-sheet cell when this transform is selected. |
| `shell_defaults.output_file_path` | Yes | Default value written into the workbook `config` sheet and auto-populated into the `Output File Path` control-sheet cell when this transform is selected. Must be distinct per profile that shares a `Target EIB` — see the Control-Sheet Inputs table. |
| `shell_defaults.mapping_spreadsheet` | Inherited | **Inherited from `config/reference_paths.yaml`** — omit from the profile unless this transform must use a different mapping file. If set here, the profile value overrides the shared default. |
| `shell_defaults.mapping_sheet` | Inherited | **Inherited from `config/reference_paths.yaml`.** Override here only if the tab name differs. |
| `shell_defaults.customer_spreadsheet` | Inherited | **Inherited from `config/reference_paths.yaml`.** Override here only if this transform must use a different customer extract. |
| `shell_defaults.supplier_spreadsheet` | Inherited | **Inherited from `config/reference_paths.yaml`.** Override here only if this transform must use a different supplier extract. |
| `shell_defaults.workday_reference_ids` | Inherited | Inherited from `config/reference_paths.yaml` (currently unset there — no transform macro reads this field). Override here only if a transform needs it. |

### Shared reference-data paths (`config/reference_paths.yaml`)

The three fixed reference files — the Workday **customer** extract, the Workday **supplier**
extract, and the finance **Mapping MASTER** (plus its `mapping_sheet` tab) — are defined **once**
in `config/reference_paths.yaml`, not per profile. When Workday issues a new extract or finance a
new mapping master, change the path there once; every transform picks it up on the next build.

Resolution precedence at build time: **a profile's `shell_defaults` value wins if present; otherwise
the shared value from `reference_paths.yaml` is injected** into the Runner `config` registry (the
build mechanism is described in `docs/ARCHITECTURE.md`). A transform still only reads the reference
cells it needs (AP → supplier + mapping, AR → customer + mapping).

**Consequence if this drifts:** every transform resolves suppliers/customers/cost-centres against a
stale file and silently rejects valid rows — the exact multi-file-edit drift this file removes.

**Per-host root.** `reference_paths.yaml` — and every `config/transforms/*.yaml` profile's
`source_data` / `target_eib` / `output_file_path` — is authored with one person's local mac data
root (`/Users/Rakesh.Patel/Documents/__data`). Rather than editing any of those committed files
per host, a separate build step (`build_runner.py --substitute-root-paths`, run after the
normal build) rewrites that root for the current build host; see "Build-time path substitutions"
below and `docs/BUILD_WIN.md` Step 3.

### Build-time path substitutions (`config/root_path_substitutions.yaml`)

`scripts/build_runner.py` applies a named path-rewrite profile from `config/root_path_substitutions.yaml`
to filesystem paths it writes into `dist/Runner.xlsm`. Build-time only — the VBA runtime never reads
this file.

The profile is auto-selected from `platform.system()` via the file's `os_default_profile` map
(`--path-profile <name>` overrides; `none` disables). On macOS/Linux it selects `none` (paths
written verbatim). On **Windows** it selects `windows_local`, which does exactly three things,
in order, to whichever field(s) a given call site asks it to rewrite:

1. **`root_path_substitution` → `new_root_path`.** A literal string swap: the mac data root
   baked into `reference_paths.yaml` and `transforms/*.yaml` shell_defaults is replaced with this
   build host's own local root. Edit both values for your machine (see the `EDIT THIS` comment in
   `config/root_path_substitutions.yaml`). No-op for a value that doesn't contain the mac root
   (e.g. `bootstrap_source_folder`, which is the repo's own `vba_src` path, unrelated to the data
   root).
2. **`convert_forward_slash_to_backslash`.** Every remaining `/` is normalized to `\`, so the
   result is a fully native `C:\Users\...\...` path — including any backslashes step 1 just
   introduced, and any forward slashes already present in the original value.

There's no per-field targeting: every field listed in `apply_to` gets the same two-step
treatment via `PathSubstitution.rewrite()`. What differs between fields is **when** each gets
rewritten, not **how**:

- `bootstrap_source_folder` is rewritten automatically as part of the normal build
  (`docs/BUILD_WIN.md` Step 2). It's already a native Windows path before this runs (step 1 is a
  no-op for it), so in practice this step only matters for parity — the result is unchanged from
  what `vba_source.resolve()` already produced.
- `mapping_spreadsheet` / `customer_spreadsheet` / `supplier_spreadsheet` / `source_data` /
  `target_eib` / `output_file_path` are rewritten only by the separate
  `build_runner.py --substitute-root-paths` command (`docs/BUILD_WIN.md` Step 3), which patches
  the already-built workbook in place.

Writing `bootstrap_source_folder` as a native backslash path is safe: `BootstrapImporter.DetectSeparator`
and `BaseName` (`vba_src/bootstrap/BootstrapImporter.bas`) explicitly recognise the Windows
drive-letter shape and handle it correctly. Before that fix, `DetectSeparator` misread a
drive-letter path's `C:` colon as a classic-Mac `HD:Folder:File` separator — this file's
predecessor rewrote that one field to forward-slash as a workaround; that workaround is gone now
that the underlying VBA bug is fixed.

Rules are applied top to bottom (string swap, then separator normalize); `apply_to` lists which
build-written values a profile touches. See `docs/BUILD_WIN.md` and the file's own header comment.

**Consequence if this drifts:** a Windows build silently writes unusable paths pointing at files
that don't exist on this host, if `root_path_substitution`/`new_root_path` don't match this
machine's actual data root.

## Source Layout Metadata

The source workbook's structure is declared in the transform profile under a top-level `source_layout`
mapping so the workbook shape lives in config — not hidden inside the eventual VBA module, and not
restated in the authoring spec. It is essential when the raw data does not start on row 1 of the first
visible worksheet (e.g. the AMEX profiles, header row 9 / row 1 on a named sheet); profiles whose data
starts on row 1 of the first sheet still declare `header_row`/`data_start_row`/`required_headers` here
(omitting `worksheet_name`) so source layout has one consistent home across every profile.

Current `source_layout` keys:

| Key | Required | Meaning |
|---|---|---|
| `source_layout.worksheet_name` | No | Explicit source worksheet to read (for example `New CC Updated`). If omitted, a transform may still choose to use the first visible sheet, but that behaviour should be treated as a transform-level decision rather than an undocumented default. |
| `source_layout.header_row` | No | 1-based row number containing the source headers. Useful when the workbook has title, filter, or staging rows above the actual table. |
| `source_layout.data_start_row` | No | 1-based row number where data records begin. Usually this is `header_row + 1`, but it is declared separately so intentionally blank spacer rows can be represented if needed. |
| `source_layout.required_headers` | No | Ordered list of headers that must exist on `header_row` for the source workbook to be considered structurally valid. |

These keys are currently design-time metadata only: they document the intended source layout and will
be consumed by the future transform macro when that implementation is added. Existing build tooling
does not read them yet, so adding them does not affect shell generation or transform-registry output.

## Ingest Specs

`config/transform_specs/` describes the **output** side: how an existing source workbook maps onto an
EIB. `config/ingest_specs/` describes the **input** side: the shape of a template workbook this repo
hands *to* a finance team for them to populate by hand. The two are separate directories because they
answer different questions and are authored at different times — an ingest spec is written before any
source data exists, and its `open_questions` are the finance decisions that must close before the
matching transform spec can be authored at all.

A spec is the design-time contract for the coding agent and the direct input to the ingest-template
renderer (`scripts/create_ingest_template.py`). It is not read by the Runner build or by the finance
user's runtime workbook; it is reviewed as a normal YAML diff and rendered into `dist/ingest_templates/`.

### Derivation direction: specs are derived from the EIB, not from a source extract

**An ingest spec's column list is exactly the set of EIB columns that actually get populated.** Work
backwards from what you are trying to create. The finance team supply final Workday values, and the
transform performs **no derivation, mapping, or lookup** of what they typed.

This inverts the obvious approach (take a legacy extract, drop the columns nobody uses) and it is worth
understanding why, because the reference pipeline in `elt_ingest_excel` is full of derivations that look
like requirements:

- Those derivations exist because a **migration** had to reconstruct Workday values from legacy GP data
  that did not contain them. A form can simply ask for the value. The cost is only justified by the
  migration.
- Measured on `ar_customer`'s reference output (4,132 rows): `tax_id` is populated on **13 rows**, yet
  drives a format-sniffing rule (9 digits → VAT Reg No, 8 alphanumeric → Company Number, 10 digits →
  UTR) plus `fn_nvl2` gating four more columns. `payment_terms` is `Net 30` on 4,095 of 4,132, behind a
  full source→Workday mapping table. `transaction_tax_status` is `Customer of Goods` on 3,819, behind
  country + tax-schedule + registration-number branching.
- Deriving from the source also produces columns nobody needs: `ar_customer`'s previous version carried
  a `Company` column because GP had one, when its only EIB target (Customer Company Restrictions) is
  `WHERE 1=2` and never populated at all.

**The line to hold:** the transform may generate an ID from another field, restate a value the user
already typed, or infer a value from *which column* something was typed into. It must not make a
business decision. Anything requiring judgement is a template column. See `transform_supplied` below.

**Consequence of getting the direction wrong:** you rebuild the migration. The spec inherits dedup,
aggregation, ranking and mapping machinery that exists only because legacy data was messy, and the
finance team are asked to populate fields that are discarded while the fields Workday actually needs
get guessed at.

Current specs (`ar_customer` and `ap_supplier` are in scope as input templates; invoice/credit note specs are legacy placeholders):

| Spec | Template | EIB population type |
|---|---|---|
| `config/ingest_specs/ar_customer.yaml` | Customer | Customer |
| `config/ingest_specs/ap_supplier.yaml` | Supplier | Supplier |
| `config/ingest_specs/ar_customer_invoice.yaml` | Customer Invoice (out of scope) | Customer Invoice |
| `config/ingest_specs/ar_customer_credit_note.yaml` | Customer Credit Note (out of scope) | Customer Invoice Adjustment |

Top-level keys:

| Key | Required | Meaning |
|---|---|---|
| `meta.key` / `meta.label` | Yes | Spec identity. `key` matches the filename. |
| `meta.requester` | Yes | The team that populates the template and owns its `open_questions`. |
| `meta.eib_population_type` | Yes | The EIB the populated template ultimately feeds. |
| `meta.purpose` | Yes | Why the template exists and what it deliberately leaves out. |
| `meta.derived_from` | Yes | What the column set was derived from — for an EIB-derived spec, the populated columns of the reference `workday_*` outputs, plus when that was measured. The provenance a later session must re-check. |
| `meta.not_derived_from` | No | An explicit non-provenance. Present on `ar_customer` to state that the GP extract is *not* its basis — it was in an earlier version, and the file otherwise reads as though it still is. |
| `meta.eib_workbook` / `eib_header_row` / `eib_data_row` | EIB-mapped specs | The real workbook every `feeds` entry is verified against, and where its headers live. |
| `meta.eib_coverage` | Multi-tab EIBs | Populated-vs-total column count per tab. Makes the "60 of 124 columns are permanently empty" fact visible without opening the workbook. |
| `template.worksheet_name` / `header_row` / `data_start_row` | Yes | The generated workbook's layout. Mirrors `source_layout` in a build profile, which is what this template becomes once populated. |
| `template.grain` | Yes | What one row represents. Names the row-multiplicity assumption a reviewer would otherwise have to infer. |
| `template.fan_out` | Multi-tab EIBs | How many EIB rows one template row produces, per tab. |
| `columns` | Yes | One entry per column the template exposes (see below). |
| `decisions` | Yes | Resolved decisions: `{id, decided, by, date, rationale, consequence}`. A question that moves out of `open_questions` lands here — it is not deleted. |
| `lookups` | No | Shared reference datasets to materialise as tabs — a list of **bare names**, resolved to `config/ingest_lookups/<name>.json`. The renderer uses the lookup name as the logical identifier and workbook formula source; the visible worksheet title may be deterministically shortened to satisfy Excel's 31-character sheet-name limit. See "Ingest Lookups" below. |
| `transform_supplied` | EIB-mapped specs | Populated EIB columns that are **not** template columns: `{tab, column, field, rule}`. Only three rules qualify — generated from an ID, a restatement of a typed value, or implied by which column a value was typed into. Everything else is a template column. |
| `excluded` | Yes | EIB columns that exist but are **never populated** — `{tab, count, columns}`, plus optional `letters` (e.g. `"D-K, M-T"`; `ap_supplier` carries it) so `count` can be checked against the letters and the tab's header width. The transform leaves them empty and the template never mentions them. |
| `excluded_tabs` | Multi-tab EIBs | Whole EIB tabs never populated (the reference transform is `WHERE 1=2`), with the reason. |
| `mapping_phase_findings` | No | Notes for the **later** step that maps the template onto a 3+ Master load file (`Master Customers.xlsx` / `Master Suppliers.xlsx`): Master-only constants (Add Only, Submit, Type BUSINESS, Use For, ...) and where the Master needs Workday reference IDs rather than display values. A list of `{master_column, field, value, note}`. Recorded, not acted on — the renderer ignores it and it changes no `feeds`. |
| `open_questions` | Yes | Unresolved finance decisions. Same role as in a transform spec: must be empty before the transform is authored. |

**Every populated EIB column must be reachable from exactly one of three places — a column's `feeds`, a
column's `drives`, or `transform_supplied`.** That is the spec's central invariant and the thing to
check first on any edit (see the completeness command below). A populated column in none of them is a
field Workday expects that nobody has decided how to fill; a template column that no longer maps to the
EIB is a field the finance team populate for nothing.

**Prefer a coded key over a display name.** Where the EIB carries both a code and a name for the same
thing, have the finance team supply the **code** and derive the name — not the other way round. In
`ar_customer`, AR pick `Address Country Code` = `GB` and the transform writes `United Kingdom` into
`Customer Address!K`. The name is the field that attracts variance (UK / United Kingdom / Great
Britain — precisely what the reference pipeline needed `ref_source_country_name_mapping` to repair);
the code has none. Check the mapping is a bijection first: all 106 `ref_country` rows have a unique
code *and* a unique name, so nothing is lost. This is not a derivation in the prohibited sense — it is
a restatement of a value the user chose, with no business decision in it.

**Take allowed values from the EIB's named ranges, not from what the migration produced.** The EIB
carries the tenant's own valid-value lists (`Customer Specific Named Ranges`, and the `defined_names`
that point at them — e.g. `Default_Customer_Group`). The migration's output is only the subset it
happened to need. For customer groups those differ: the named range has **7** values, the migration
emitted **5**, so building the list from the output would silently deny AR `Experiences by Wembley
Seats` and `Experiences by Wembley Boxes` — legitimate groups with zero migrated rows. Reference data
the EIB or `elt_ingest_excel` does not already own lives in `implementations/elt_transform_excel/config/data/`.

**Derive what the user already told you; ask for the rest.** A value that is mechanically determined by
another column is not a question — record it as a `drives` rule on the column it comes from. The
alphabetical customer-group band is the worked example: it falls out of `Customer Name`'s first
character, so AR type only the *extra* groups. **But the rule's parameters are config, not code.** The
band boundaries live in `ref_customer_group.json` (`first_char_from`/`first_char_to`/`is_fallback`)
because `workday_customer_groups.sql` hard-codes them in a SQL `CASE`, and a hard-coded band cannot be
moved without a code change. The same tab therefore does two jobs: `assignment = selectable` lists what
AR may pick, and `assignment = auto_band` carries the boundaries the transform bands with.

**Do not collapse two coded columns because they usually agree.** `ar_customer` carries three separate
country codes (address, phone, tax status). Phone country differs from address country on 75 of 3,432
reference rows and tax status country on 3 of 4,120 — deriving either from the address would silently
miscode exactly those rows, and silently is the problem.

**Why `decisions` exists rather than just deleting the closed question:** an ingest spec is derived from
a reference pipeline that is usually solving a *different* problem. `ar_customer.yaml` is derived from a
one-off migration that merged the same customer across three legal entities, so that SQL is full of
dedup, aggregation, address arrays and ranking — none of it a business requirement. A session reading
only the SQL will faithfully reproduce the merge machinery. Recording "one address per customer, because
multi-address was a merge artifact" is the only thing that stops that, and a deleted question stops
nothing.

Per-column keys:

| Key | Required | Meaning |
|---|---|---|
| `name` | Yes | The literal header text written into the template. Business title case (`Customer ID`, not `customer_id`) — this is what a person types under, and it becomes the `source_column` value in the matching transform spec. |
| `required` | Yes | `true`, `false`, or `conditional` (which then needs a `condition`). |
| `type` | Yes | `text`, `number`, `date`, or `enum`. |
| `description` | No | Why the column exists and what goes wrong if it is populated badly. |
| `picklist` | No | Makes the column a real Excel drop-down. Static form: `{lookup, column, where}`. Filtered/cascading form: `{lookup, column, filtered_by, key}`. `lookup` is a name from `lookups`, `column` the field in it to offer, `where` an optional fixed equality filter (e.g. `{assignment: selectable}`), `filtered_by` names another template column whose selected value filters this list, and `key` is the lookup field matched to that selected value. The renderer emits a list validation over the whole column pointing at the lookup tab. |
| `multi_value` | No | States the in-cell delimiter when one cell may carry several values (e.g. `';'-delimited list`). Its absence means single-value — in `ar_customer.yaml` a `;` in an address line is just text (since 2026-10-12 `ar_customer.yaml` has one phone and one email per customer, so it no longer has any `;`-split column). |

**`picklist` and `multi_value` are mutually exclusive.** An Excel drop-down yields one value per cell,
so it cannot produce a delimited list. If a column needs both, one of the two assumptions is wrong —
check the data before reaching for free text. That is what happened to `Additional Customer Group` (since removed, 2026-10-12): it
was a `;`-list until the reference output showed every customer has **exactly one** extra group (402 of
402), so single-value plus a drop-down was strictly better.

**`default` and `picklist` compose — deliberately.** A pre-filled cell the user can override is the
common case: `Default Currency` is `GBP` for 3,870 of 4,132 reference customers but genuinely varies
(118 EUR, 40 USD), so it wants both a pre-fill and a drop-down. This is why `expected_value` was
replaced: it asserted "identical on every row", which contradicted offering a choice, and forced a
false either/or on 11 columns.

**`default` vs `example` — the distinction is what ships.** A `default` is a real value meant to be
kept, so the renderer pre-fills it on the first data row every time. An `example` is illustrative
(`Acme Sports Ltd`) and must never reach Workday, so it appears only under `--sample-row`. Getting this
backwards ships either a template missing its defaults, or fake customer data that looks submittable.
| `allowed_values` | enum only | The valid values. |
| `allowed_values_ref` | No | The authoritative reference file the values come from. Present so a drifted `allowed_values` list can be re-derived rather than trusted. |
| `default` | No | Value assumed when the cell is blank. |
| `format` | date only | The format the template expects (e.g. `DD/MM/YYYY`). |
| `example` | No | A sample value. |
| `feeds` | See below | Where the value is **written**. A list of `{tab, column, field}` — the EIB worksheet, its column letter, and the header text at `meta.eib_header_row`. Add `note:` for a reshape (e.g. one row per value). This is the link to the transform spec that does not yet exist. |
| `drives` | See below | Where the value **selects a derived value** without being written. Same shape as `feeds` plus `rule:` stating the derivation. |
| `feeds_master` | See below | For a value the EIB workbook has **no header-row-3 column** for, but the 3+ Master load file does: `{sheet, column, field, status}` naming the Master sheet column. `status: unverified-until-mapping` until the mapping step checks it. Example: `ap_supplier`'s `Supplier Status` → `Submit Supplier - v46.2!MI` (`Status*`). |

| `default` | No | A value **pre-filled on the first data row, which the user may change** (e.g. `GBP`). Not "the value on every row" — that is what the superseded `expected_value` meant, and it contradicted `picklist`. |
| `cross_reference` | No | Another ingest spec column this one must match (e.g. an invoice's Customer ID against the customer template's). |
| `condition` | conditional only | When the column becomes required. |

**Every column must carry at least one of `feeds` or `drives` — a column with neither is the real
error**, because it means nobody knows what the field is for and it should not be in the template.
`feeds_master` is the third, weaker way to satisfy this rule: it is allowed only when the EIB workbook
has no column for the value at all, and it stays **unverified until the mapping step** checks it
against the Master file. The `feeds`/`drives` verification below skips it and should list it as
skipped rather than passed.

**`drives` is how a coded column reaches the columns it expands into.** A column may carry both: in
`ar_customer`, `Phone Country Code` **feeds** `Customer Phone!E` (the code itself) and **drives**
`Customer Phone!D` (the country name) and `!F` (the dialling code). A `drives`-only column is
legitimate — `Tax Status Country Code` writes to no cell at all, because the Customer Tax tab has no
country-code column, yet it determines `Tax Status Country`.

The distinction earns its keep because the two fail differently: a wrong `feeds` shows up as a visibly
wrong cell, while a wrong `drives` yields a valid-looking EIB that is silently miscoded.

**What `drives` must never become is a smuggling route for derivation.** Its `rule` may only restate a
value the user chose (`GB` → `United Kingdom` → `44`, all one bijective lookup away). The moment a
`rule` encodes a business decision — a format-sniffing branch, a threshold, a default that depends on
another field's value — the derivation direction has slipped and the value belongs in its own column.
Cautionary case: an early draft recorded `Tax Schedule ID` as `feeds:
workday_customer_tax.tax_schedule_id`. **No such output column exists** — it was written nowhere and
merely selected `Transaction Tax Status` from a country + schedule + registration-number branch. That
is a business decision, and it is now a typed column. A fabricated `feeds` is worse than no entry: it
reads as verified when it is fiction. Verify against the real workbook (below).

**Multi-tab EIBs fan out.** `template.fan_out` states how many EIB rows one template row produces per
tab. One customer row is one Customer Name row but *n* Customer Phone rows — one per populated phone
field, each then `;`-split. A spec without `fan_out` invites the assumption that one input row is one
output row per tab.

### Verifying an ingest spec's `feeds`/`drives`

Every `{tab, column, field}` is checkable against the real EIB workbook named in `meta.eib_workbook`,
and must be checked rather than reasoned about — column letters are positional and shift.

```bash
uv run --with pyyaml --with openpyxl --no-project python -c "
import yaml, openpyxl, sys
from openpyxl.utils import get_column_letter
d=yaml.safe_load(open('config/ingest_specs/ar_customer.yaml'))
wb=openpyxl.load_workbook('<path to>/FA E2E Customers V1.xlsm', read_only=True, data_only=True)
hdrs={}
for ws in wb.worksheets:
    try: row=next(ws.iter_rows(min_row=3,max_row=3,values_only=True))
    except StopIteration: continue
    hdrs[ws.title]={get_column_letter(i+1): str(v).strip() for i,v in enumerate(row) if v is not None}
bad=0
for c in d['columns']:
    for k in ('feeds','drives'):
        for f in c.get(k,[]) or []:
            if f.get('tab')=='ALL child tabs': continue
            a=hdrs.get(f.get('tab'),{}).get(f.get('column'))
            if a is None or (f.get('field') and a.lower()!=str(f['field']).strip().lower()):
                print(f'BAD: {c[\"name\"]} -> {f.get(\"tab\")}!{f.get(\"column\")} spec={f.get(\"field\")!r} actual={a!r}'); bad+=1
sys.exit(1 if bad else 0)
"
```

**The completeness check — the one that matters most.** The command above proves each reference is
*valid*; this proves the set is *complete*. For each `workday_customer_*` table, every column with a
non-zero populated count must appear in a column's `feeds`, a column's `drives`, or
`transform_supplied` (excluding the `customer_id`/`customer_name` join keys, which are `ALL child
tabs`). As of 2026-07-15 `ar_customer` accounts for **all 64 populated columns — 0 unaccounted for**
(43 template columns + 8 transform-supplied + the 2 join keys, spanning 9 tabs).

```bash
# needs a copy of the DuckDB - DBeaver holds an exclusive lock on the original:
#   cp ~/Documents/__data/duckdb/rpatel.duckdb <scratchpad>/ro.duckdb
uv run --with pyyaml --with openpyxl --with duckdb --no-project python -c "
import yaml, duckdb
from openpyxl.utils import get_column_letter
d=yaml.safe_load(open('config/ingest_specs/ar_customer.yaml'))
covered={(f['tab'], f['column']) for c in d['columns'] for k in ('feeds','drives') for f in (c.get(k) or []) if f.get('tab')!='ALL child tabs'}
covered |= {(f['tab'], f['column']) for f in d['transform_supplied']}
con=duckdb.connect('<scratchpad>/ro.duckdb', read_only=True)
tabs={'workday_customer_name':'Customer Name','workday_customer_groups':'Customer Groups','workday_customer_payment':'Customer Payment','workday_customer_currencies':'Customer Currencies','workday_customer_tax':'Customer Tax','workday_customer_address':'Customer Address','workday_customer_phone':'Customer Phone','workday_customer_email':'Customer Email','workday_customer_document_delivery':'Customer Document Delivery'}
missing=[]
for t,tab in tabs.items():
    cols=[r[0] for r in con.execute(f\"SELECT column_name FROM information_schema.columns WHERE table_name='{t}' ORDER BY ordinal_position\").fetchall()]
    for i,c in enumerate(cols,1):
        if c in ('customer_id','customer_name'): continue
        if con.execute(f'SELECT COUNT(\"{c}\") FROM {t}').fetchone()[0]==0: continue
        if (tab,get_column_letter(i)) not in covered: missing.append(f'{tab}!{get_column_letter(i)} ({c})')
print('unaccounted populated EIB columns:', len(missing)); [print('  ',m) for m in missing]
"
```

**Why the tab/column check matters more than it looks:** `publish_customer.json` has no `use_column_name_index`,
so the publish step writes **positionally** — `workday_customer_*` table column *n* lands in EIB column
*n*, matched by order alone, never by name. Reordering a `SELECT` in the transform SQL silently shifts
every subsequent EIB column, and nothing in the pipeline detects it. The header text is the only
independent check that a recorded column letter still means what the spec says it means.

**Provenance rule:** an ingest spec's column set is derived from the SQL that populates the
corresponding `workday_*` tables in `elt_ingest_excel`, not from the source extract's column list. The
GP extracts carry many columns the pipeline never reads, plus columns it derives rather than reads —
copying the extract's headers reproduces both classes as fields finance would be asked to populate for
no reason. Record each such column under `excluded` with its reason instead.

## Ingest Lookups

`config/ingest_lookups/<name>.json` holds a shared reference dataset. An ingest spec names the ones it
wants; the renderer materialises each as a tab in the generated workbook and can back a column's
drop-down with it. Several specs share one lookup, so each file is **self-describing** — the
documentation lives with the data instead of being restated by every spec that uses it:

| Key | Meaning |
|---|---|
| `description` | What the dataset is and what consumes it. |
| `key_column` | The column that identifies a row. |
| `provenance` | Where the values came from and when — the thing a human re-checks. Always a **snapshot**, never a link. |
| `rows` | A non-empty list of flat objects. Header order follows the first object's keys. |

Current lookups used by ingest specs (24): shared — `ref_country` (106), `ref_country_region` (2522),
`ref_payment_term` (9), `ref_payment_type` (13), `ref_phone_device` (2), `ref_yes_no` (2); customer —
`ref_customer_category` (12), `ref_customer_group` (7), `ref_country_tax_id_type` (534), `ref_currency`
(155), `ref_currency_default` (154), `ref_tax_code` (138), `ref_transaction_tax_status` (5),
`ref_document_delivery_method` (2), `ref_document_email_recipient_invoice` (8),
`ref_document_email_recipient_statement` (7); supplier — `ref_supplier_category` (6),
`ref_supplier_group` (6), `ref_supplier_status` (3), `ref_supplier_currency` (150),
`ref_supplier_transaction_tax_status` (12), `ref_supplier_country_tax_id_type` (541),
`ref_po_issue_option` (9), `ref_bank_account_type` (2). The supplier lookups were built from
`FA E2E Suppliers V1.xlsm`'s named ranges; `ref_supplier_currency` carries a `default_allowed`
column so a `where: {default_allowed: "Yes"}` filter drops `All Currencies` from Default/Bank
Currency.

Two lookups in the same directory use the identical self-describing shape but are consumed by the
**Runner build** rather than by ingest specs: `ref_company` (source→target company names) and
`ref_default` (per-domain defaults, e.g. currency→GBP). `scripts/build_runner.py` folds them into the
Runner's runtime `mapping` sheet — see `docs/RUNBOOK.md`'s "Mapping Lookup Workflow".

Filtered/cascading picklists still use the same lookup files. The difference is that the renderer
must materialise one named range per distinct `key` value (for example `region_GB`, `region_IE`) and
point the validated cells at the range selected by the value in `filtered_by`. The lookup file
therefore needs the display column (`region`) and the filter key (`country_code`), with rows grouped
contiguously by key so each named range resolves to one continuous block.

**Lookups resolve by name inside this module — never by a caller-supplied path.** That is a
structural guarantee, not a convention: the renderer builds the path itself from
`config/ingest_lookups/`, and rejects anything containing a separator. So no spec can point it at
another module's config, or at a workbook that happens to be in someone's `~/Downloads`.
`elt_transform_excel` stays unbound from `elt_ingest_excel` by construction rather than by discipline.

**The cost is real and deliberate: these are copies, and copies drift.** `ref_country` was snapshotted
out of `elt_ingest_excel`; `ref_customer_category` and `ref_customer_group` were transcribed from the
EIB's named ranges. Each file's `provenance` names its origin and date precisely so the check is a
manual diff, not archaeology. The alternative — reading the origin live — would bind this module to a
migration repo whose lifecycle is unrelated, and to a workbook nobody controls.

**Reference data may carry rule parameters, not just valid values.** `ref_customer_group` does both
jobs: `assignment = selectable` lists what a user may pick, and `assignment = auto_band` rows carry the
`first_char_from`/`first_char_to`/`is_fallback` boundaries a transform bands a name with. Those
boundaries are hard-coded in a SQL `CASE` in the reference pipeline, which is exactly why they belong
in config here — a band cannot be moved without a code change otherwise.

**A `where` filter must select contiguous rows**, because an Excel drop-down is a range reference. The
renderer computes the range and fails with a clear message if the matching rows are scattered — the fix
is to reorder the lookup file so the subset sits together, which is why `ref_customer_group` lists its
three `auto_band` rows before its four `selectable` ones.

## Mapping Contract

Each transform's field mapping is a `mapping:` list in `config/transform_specs/<key>.yaml` — one item
per target EIB column, keyed by `target` (the EIB column letter, lower-cased). It is the authoring
source of truth the transform's `.bas` module mirrors; the build does not read it, so a mapping change
is reviewed as a normal YAML diff. (This replaced the former per-transform tabs in the retired
`RunnerShellMapping.xlsx` and the MD5 `mapping_content_hash` build gate.) A row provides
its value in one of these ways:

| Key(s) | Meaning |
|---|---|
| `default` | A fixed literal value. |
| `source_column` | Read the value (or lookup input) from a source-workbook column on the current source row. |
| `source_value` | A fixed literal lookup input instead of a source column — e.g. `ap_amex_hotel` resolves fixed supplier code `AME001` through the supplier extract's `Reference ID` column to the Workday `Supplier ID`. |
| `function` | A derived value (see the function table below). |
| `lookup:` | A reference lookup (see `lookup.file` tokens below); `source_column`/`source_value` supplies the key. |
| `source_format` / `target_format` | Date/text columns only. Declare a date reformat (e.g. `DD/MM/YYYY` → `YYYY-MM-DD`) OR, for feeds that ship multiple source subtypes in the same column, use the token `shared_helper:source_date_normalization.excel_serial_and_dmy_text` to reference the canonical NormalizeSourceDate contract in `config/shared_helper_registry.yaml` (4-branch execution order with mixed-format handling: numeric Excel serials by VarType + text-serial scan → numeric DMY text before IsDate() → IsDate() fallback). When the transform's source column is observed to mix Variant/String human-readable dates (e.g. "14-Aug-26") with General-formatted Variant/Double date serials with time (e.g. 46223.64707) in a single workbook — as on the Travel Places DATA sheet, rows 1-393 vs 394+ — the shared-helper token is the SSO; the per-field `format:` in the transform_spec may still restate the observed subtypes for human readability but MUST reference the same shared helper via `shared_helper_ref`. |

Where a row resolves its output through a reference lookup (`lookup:`), `lookup.file` names which
workbook holds the lookup table. Current tokens:

| `lookup.file` | Meaning |
|---|---|
| `this` | The `mapping` lookup table itself (runtime: the `mapping` tab inside `runner_shell/RunnerShell.xlsm`). Used for the `Company` entity lookup. |
| `Supplier Spreadsheet` | The Workday supplier extract named by the `Supplier Spreadsheet` control-sheet input. Used to resolve `AME001` -> Workday `Supplier ID`. |
| `Mapping Spreadsheet` | The finance mapping workbook named by the `Mapping Spreadsheet` control-sheet input. The tab usually comes from the `Mapping Sheet` control-sheet input (currently `MAPPING MASTER`). Used by the AMEX transforms for Cost Centre and related worktags off `Final Code`. |
| `Source Data` | A tab inside the source workbook named by the `Source Data` control-sheet input. Available for transforms that must read a secondary lookup or staging tab from the same workbook as the primary source data. |

The `function` column may also describe transform-specific derived values that do not come from a
worksheet cell directly. Current examples:

| Function | Meaning |
|---|---|
| `invoice_period_from_filename` | Derive an invoice-period token from the source workbook file name (for example `AMEX HOTELS APR 26_UPDATED.xlsx` -> `APR26`) and append the row's `source_value` as a suffix literal. On `ap_amex_hotel`, rows `g` and `au` therefore capture `APR26 AMEX INV DD` as `function = invoice_period_from_filename` plus `source_value = AMEX INV DD`; on `ap_amex_trains` the suffix is `AMEX TRAINS INV` (kept distinct from the hotel EIB because both post to supplier `AME001`). |
| `vat_rate` | Normalise the row's `source_column` value (`VAT Percentage`) to an integer percent and select the tax set for that rate. On `ap_amex_trains`, rows `ds`/`dt`/`dy` resolve to the zero-rated set (`GBR_Outside_of_Scope` / `GBR_Zero_Rated` / `Nonrecoverable_Allocate`) for `0`% rail fares and the standard-rated set (`GBR_Standard_Rated` / `GBR_Standard_Rated` / `Recoverable_Allocate`) for `20`% booking/admin fees. The two sets are mirrored as constants in `modApAmexTrainsSupplierInvoice` (the `.bas` mirrors the spec `mapping:` section), so the shipped runtime `mapping` tab holds only company/currency, not tax. |
| `reporting_type_from_final_code` | Split the source `Final Code` accounting string on `-`, take the second element, then map its first character to the Reporting Type list: `1/2/3/7/8/9 -> FA Operations`, `4 -> FA Investments`, `5 -> National Game Board (NGB)`, `6 -> Professional Game Board (PGB)`. Unrecognised or malformed values resolve blank. |

## Control-Sheet Selection Model

The Runner now has one controlling dropdown:

| Control | Source | Runtime behaviour |
|---|---|---|
| `Transform` | Derived automatically from every `transform.label` value in `config/transforms/*.yaml`, deduplicated in filename-sorted order. | Selecting a label resolves the matching row in the visible transform registry and the workbook `config` sheet. `Source Type` and `EIB Population Type` come from the transform registry; the editable path/file defaults (`Source Data`, `Target EIB`, `Output File Path`, `Mapping Spreadsheet`, `Customer Spreadsheet`, `Supplier Spreadsheet`, `Workday Reference IDs`) come from `config`. The dispatcher also resolves the macro from the registry row. |

Adding a new transform profile therefore only requires the new YAML file under `config/transforms/`
(its `transform.label` automatically appears in the dropdown) — no separate Runner dropdown list to edit by hand.

> For the full end-to-end procedure of adding a transform (finance decisions to confirm, source/EIB
> inspection, mapping authoring, macro, and build), see `docs/ADDING_A_TRANSFORM.md`.

Current profiles:

- `AR Invoice` -> `Customer Invoice`. Source data and reference files use the repo's recorded
  Customer Invoice sample set. Field mapping lives in the `mapping:` section of
  `config/transform_specs/ar_invoice.yaml`, mirrored by `modArInvoiceCustomerInvoice`; its
  `source_layout` (header row `1`, `worksheet_name` omitted for the single-sheet CSV) sits in the profile.
- `AR Discipline Invoice` -> `Customer Invoice`. Reuses the same target EIB as AR Invoice. Field mapping
  in `config/transform_specs/ar_discipline_invoice.yaml`; Company and Currency are fixed; `source_layout`
  (header row `1`, `worksheet_name` omitted for the single-sheet workbook) sits in the profile.
- `AP AMEX Hotel` -> `Supplier Invoice`. Field mapping in `config/transform_specs/ap_amex_hotel.yaml`;
  declares the source workbook layout explicitly via `source_layout` (`New CC Updated`, header row `9`,
  data start row `10`), and resolves the Supplier field by looking up fixed source supplier code
  `AME001` through the supplier spreadsheet. Spend Category and Tax are fixed constants; Reporting Type,
  Cost Centre, and the optional related-worktag columns `ER:EX` are derived per row from the source
  `Final Code` accounting string via the mapping master (`Core vs Prism = Core`).
- `AP AMEX Trains` -> `Supplier Invoice`. Field mapping in `config/transform_specs/ap_amex_trains.yaml`;
  declares the source workbook layout explicitly via `source_layout` (`CC Updated New`, header row `1`,
  data start row `2`), and resolves the Supplier field the same way as `ap_amex_hotel` (`AME001`). It differs
  from the hotel profile in three domain rules: (1) **Cost Centre** (EIB col `EQ`) and the optional
  related-worktag columns `ER:EX` are resolved per row from the source `Final Code` through the mapping
  master (`WORKDAY BUSINESS OBJECT` filter plus `Core vs Prism = Core`); (2) **Tax** is reflected per row from the source `VAT Percentage`
  column — `0%` rail fares use the zero-rated set, `20%` booking/admin fees use the standard-rated set,
  with `CALC_TAX_DUE` letting Workday calculate the VAT from the tax code on the net amount; (3)
  **Reporting Type** (EIB col `EP`) is derived per row from the source `Final Code` accounting string
  using the supplier-PO mapping above; (4) the line amount (`ED`/`EE`) is the source `Net amount`
  and the Invoice Date (`U`) is the source `Statement date`. Spend Category is the fixed constant
  `Staff_Travel_Trains`.

All four transforms also resolve their `Company` field via a live runtime lookup against the `mapping`
tab inside `runner_shell/RunnerShell.xlsm`. That tab is generated at build time from
`config/ingest_lookups/ref_company.json` (source→target company names) and `ref_default.json`
(per-domain defaults) — the same self-describing reference-table mechanism as every other
`ingest_lookups` dataset. See `docs/RUNBOOK.md`'s "Mapping Lookup Workflow".

Expected expansion:

- Add one YAML file per supported transform profile under `config/transforms/`.
- Keep build-time defaults and sample paths in YAML, not hardcoded in Python.
- Future Python tooling may read the same profile to drive transform-specific build or verification steps.

## Control-Sheet Inputs

| Input | Type | Required | Meaning |
|---|---|---|---|
| `Transform` | Fixed dropdown | Yes | Selects the transform profile by `transform.label`. This is the only controlling dropdown on the Runner sheet. |
| `Source Type` | Derived text | Yes | Auto-populated from the selected transform profile's `transform.source_type`. |
| `Source Data` | File path | Yes | Workbook containing the raw source extract to ingest. Default comes from the `config` worksheet and can be edited there per machine. |
| `Target EIB` | File path | Yes | Template workbook cloned before publish. Default comes from the `config` worksheet and can be edited there per machine. |
| `Output File Path` | File path | Yes | Exact path the cloned, populated output workbook is written to. Default comes from the `config` worksheet and can be edited there per machine. Required for every run — there is no auto-derived fallback. Each transform profile must use a distinct path; two profiles that share a `Target EIB` but reuse the same `Output File Path` will overwrite each other's output on alternating runs. |
| `Mapping Spreadsheet` | File path | Scenario-dependent | Finance-maintained mapping workbook used by transform logic. Default comes from the `config` worksheet and can be edited there per machine. |
| `Mapping Sheet` | Worksheet name | Scenario-dependent | Worksheet inside the `Mapping Spreadsheet` workbook used by transform logic. Default comes from the `config` worksheet and can be edited there per machine. |
| `Customer Spreadsheet` | File path | Scenario-dependent | Workday customer extract used by customer-oriented transforms. Default comes from the `config` worksheet and can be edited there per machine. |
| `Supplier Spreadsheet` | File path | Scenario-dependent | Workday supplier extract used by supplier-oriented transforms. Default comes from the `config` worksheet and can be edited there per machine. |
| `Workday Reference IDs` | File path | Scenario-dependent | Additional reference workbook supplied for transforms that need it. Default comes from the `config` worksheet and can be edited there per machine. |
| `EIB Population Type` | Derived text | Yes | Auto-populated from the selected transform profile's `transform.eib_population_type`. |

## Rules

- `Transform` is selected from fixed values, not free text.
- Missing, unreadable, locked, or structurally invalid files are rejected before publish.
- The selected `Transform` label must resolve to exactly one row in the transform registry and exactly one runtime macro.
- Shared engine modules must not embed scenario-specific column names, sheet names, or business values.
- Scenario-specific rules belong only in per-combination transform modules (see `docs/ARCHITECTURE.md`).
- Python build scripts must read scenario defaults from YAML profiles under `config/transforms/`, not
  from hardcoded sample path constants.
- The built Runner's transform registry and `config` worksheet are generated from YAML; changing the
  macro binding or the seeded shell defaults for a supported transform should require a config edit
  plus rebuild, not a shared-engine code edit.

## Defaults

- Runtime defaults are surfaced through the generated Runner workbook control sheet and the editable
  `config` worksheet.
- Build-time seed values currently come from the `ar_discipline_invoice` profile (the business is
  currently testing AR Discipline Invoice) — see `docs/ops/CLI.md` for how to seed a different profile.
- Operators may change those seeded values in the workbook for a specific run.

## Not Configuration

- The step implementations themselves remain VBA source, not YAML.
- VBA module import order and build mechanics are implementation concerns, not operator configuration.
- Hidden worksheets are not used as a configuration surface.
- No configuration is shared with `elt_ingest_excel`. An ingest spec's `meta.derived_from` and
  `allowed_values_ref` cite paths in that repo as **design-time provenance only** — a reviewer follows
  them by hand to check the values are still right. Nothing here loads, imports, or syncs from them, and
  a value copied across is a snapshot that can drift; re-check it against the cited path rather than
  trusting the copy.
- YAML transform profiles are a developer/build aid; they are **not parsed directly by the
  finance user's runtime workbook**. No `.bas` module reads a `.yaml` file at runtime. Build
  tooling (`create_runner_shell.py`, `build_runner.py`) consumes YAML to write cell values and
  registry rows into the Runner workbook; the VBA macros read those cells via
  `ReadControlValue`. A pure value-swap (e.g. new Mapping MASTER filename with no schema change)
  propagates by rebuilding the Runner with no `.bas` edit required; a structural change
  (new/renamed mapping columns, new business-object constants) requires the coding agent to
  update the `.bas` constants to match the new YAML/spec contract.

## Planned Expansion

This document should be expanded once the first vertical slice exists and the exact profile schema,
dropdown catalogue, required or optional combinations, workbook sheet expectations, and result-reporting
fields are concrete.
