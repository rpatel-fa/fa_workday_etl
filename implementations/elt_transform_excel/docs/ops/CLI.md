# CLI Reference

## Purpose

This document records the command surface for developers and operators working from the repo. Finance users do not run these commands.

## Current Commands

| Purpose | Command |
|---|---|
| List the implementation scaffold | `ls -R implementations/elt_transform_excel` |
| Check the active backlog index (all items archived — shows summary of completed work) | `sed -n '1,120p' implementations/elt_transform_excel/docs/TODO.md` |
| Open the archived cold-start / hardening anchor (H1/H2.1 done) | `sed -n '1,280p' implementations/elt_transform_excel/docs/todo/archive/cold-start-scale.md` |
| Open the archived build anchor (build history) | `sed -n '1,260p' implementations/elt_transform_excel/docs/todo/archive/build-elt_transform_excel.md` |
| Open the archived ar_customer operating contract + history | `sed -n '1,400p' implementations/elt_transform_excel/docs/todo/archive/ar-ingest-specs.md ; sed -n '1,120p' implementations/elt_transform_excel/docs/todo/archive/ar-ingest-specs-history.md` |
| Verify forbidden runtime libraries are only mentioned in rejected-history sections | `grep -in "ADODB\|ACE OLEDB\|Power Query\|sql_src" implementations/elt_transform_excel/PRD.md implementations/elt_transform_excel/docs/ARCHITECTURE.md` |
| Confirm old Power Query spike files remain removed | `find ~/Documents /tmp -iname "pq_spike*" 2>/dev/null` |
| Render a finance-facing input template from an ingest spec (`implementations/elt_transform_excel/config/ingest_specs/<key>.yaml` -> `implementations/elt_transform_excel/dist/ingest_templates/<key>.xlsx`; in-scope keys: `ar_customer`, `ap_supplier`) | `uv run --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/create_ingest_template.py <key> --overwrite` |
| Same, plus one illustrative row from each column's `example` (`default`s are pre-filled either way) | `uv run --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/create_ingest_template.py <key> --overwrite --sample-row` |
| Generate `implementations/elt_transform_excel/runner_shell/RunnerShell.xlsm` with the bootstrap macro and profile-driven control-sheet layout | `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/create_runner_shell.py --overwrite` |
| Prep `implementations/elt_transform_excel/dist/Runner.xlsm` (copy shell, write YAML-driven transform registry, write the source-folder cell) — does **not** import any transform modules itself | `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/build_runner.py` |
| Manual step (required): in Excel, Tools > Macro > Macros, run `BootstrapImporter.BootstrapImportFromConfiguredSource` on the prepped workbook to actually import every `implementations/elt_transform_excel/vba_src/` module | *(no CLI command — must be run inside Excel; see Gotchas below)* |
| Verify the manually-built workbook (status cells, registry, module set) | `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/build_runner.py --verify` |
| Print the current manifest/label summary (raw Runner-dropdown view: all active entries split by AP/AR, plus defunct) | `uv run --with pyyaml --no-project python implementations/elt_transform_excel/scripts/manifest_summary.py` |
| Same, but stakeholder view (Amex Trains V1+V2 consolidated to 1 line item per stakeholder reporting directive) | `uv run --with pyyaml --no-project python implementations/elt_transform_excel/scripts/manifest_summary.py --consolidate-amex-trains` |
| Machine-readable manifest summary (JSON, with `_summary` block for counts) | `uv run --with pyyaml --no-project python implementations/elt_transform_excel/scripts/manifest_summary.py --json` |
| Attempt to open `implementations/elt_transform_excel/dist/Runner.xlsm` in Excel and run `modDispatcher.RunSelectedTransform` from a user-approved Apple Event host | `osascript implementations/elt_transform_excel/scripts/run_runner_macro.applescript` |

### Manifest/label summary — playbook

When a stakeholder asks "how many EIBs do we have?" or "list all the AP/AR labels" — **never**
open the `config/transform_specs/*.yaml` decision records. Read only:

1. `implementations/elt_transform_excel/config/build_manifest.yaml` → the authoritative list of
   `active_profiles` / `defunct_profiles`.
2. `implementations/elt_transform_excel/config/transforms/<profile_key>.yaml` → Runner labels come from
   either `shell_defaults.variants[].profile_label` (per-supplier / per-tab fan-out) or, if
   no variants block, from `transform.label` (one dropdown row per profile key).

`scripts/manifest_summary.py` is the deterministic implementation of those two rules. It
counts each uncommented variant and never reads `transform_specs/`. If the script's output
and a manual count ever disagree, the script is the reference to fix (or the transform
YAML needs correcting).

## Ingest Template Prototyping Loop

The spec is the source of truth; the workbook is a render of it. Never hand-edit the generated
`.xlsx` — the next regenerate silently discards the edit. Change the spec and re-render:

```bash
# 1. edit implementations/elt_transform_excel/config/ingest_specs/<key>.yaml
# 2. re-render (add --sample-row to see example values in row 2):
uv run --with openpyxl --with pyyaml --no-project python \
  implementations/elt_transform_excel/scripts/create_ingest_template.py <key> --overwrite
# 3. open implementations/elt_transform_excel/dist/ingest_templates/<key>.xlsx and review
```

The render is **deterministic** — identical spec in, byte-identical workbook out — so
`git status` staying clean after a regenerate proves the spec did not change, and any diff is
a real change rather than a new timestamp. Two things had to be handled to get that, both in
`save_deterministic()`: openpyxl's `save_workbook()` sets `properties.modified = utcnow()`
itself at save time (so it can only be corrected afterwards), and `.xlsx` is a zip whose
entries carry their own clock-derived mtimes.

What the render shows you, all from the spec — the script contains no domain values:

| Signal | Meaning |
|---|---|
| Header text | The column `name` **verbatim**. It is the contract: the transform's `source_column` will reference this exact string, so renaming a header breaks the transform. |
| Dark blue header | `required: true` |
| Grey header | `required: false` |
| Amber header | `required: conditional` |
| Green header | Has a `default` — pre-filled on the first data row, changeable |
| Cell comment | Obligation, type, expected/example value, multi-value delimiter, allowed values, the target EIB tab/column, and the description |
| Extra tabs after the input sheet | One per `lookups:` name — a reference dataset from `config/ingest_lookups/<name>.json` (e.g. `ref_country`, 106 rows). |
| A drop-down arrow on a cell | The column declares `picklist:`. The list is a range on its lookup tab, so it covers the whole column and survives a paste of hundreds of rows. |

A spec whose `feeds` still uses the superseded prose-string form is **rejected**, naming the
offending columns — rendering it would produce a plausible-looking template built on the
pre-inversion column set. See `implementations/elt_transform_excel/docs/CONFIGURATION.md` ("Ingest Specs").

## Python Invocation Rule

Run Python build tooling from the workspace root.

Preferred house rule:

```bash
uv run --package <package-name> python <script-or-module>
```

Current standalone form for `elt_transform_excel`:

```bash
uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/create_runner_shell.py --overwrite
uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/build_runner.py
```

Run these from the workspace root exactly as shown above.

`create_runner_shell.py` defaults to the transform profile at
`implementations/elt_transform_excel/config/transforms/ar_invoice.yaml`; pass `--profile-config` to seed a different
default selected transform instead, or edit the script's `DEFAULT_PROFILE_CONFIG` constant directly.
Regardless of that starting selection, the generated Runner's `Transform` dropdown is built from every
profile's `transform.label`, and the generated `Source Type` / `EIB Population Type` values also come
from those same profiles via the visible registry block on `Runner`. The path/reference defaults are
formula-populated from the generated `config` sheet so business users can edit those values in the
workbook. `build_runner.py`
(no flag) reads all YAML profiles in `implementations/elt_transform_excel/config/transforms/` to write both the
registry and the `config` sheet — but it does not import the transform modules themselves. Each
transform's field mapping lives as the `mapping:` section of `implementations/elt_transform_excel/config/transform_specs/<key>.yaml` and
is mirrored by the module `.bas`; there is no build-time hash check (a mapping change is reviewed as a
normal text diff).

`build_runner.py`'s module import is a manual Excel step, not a CLI command. Driving
`VBComponents.Import` headlessly via AppleScript proved unreliable (parameter-passing failures, Apple
Event instability mid-import — see `implementations/elt_transform_excel/docs/RUNBOOK.md`); a human must run
`BootstrapImporter.BootstrapImportFromConfiguredSource` from Excel's Tools > Macro > Macros dialog on the
already-prepped workbook. The source folder it reads is pre-filled by `build_runner.py` into a
`Bootstrap Source Folder` cell on the Runner sheet. Run `build_runner.py --verify` afterward to confirm
the import succeeded.

`run_runner_macro.applescript` defaults to `implementations/elt_transform_excel/dist/Runner.xlsm` and
`modDispatcher.RunSelectedTransform`. You can override both:

```bash
osascript implementations/elt_transform_excel/scripts/run_runner_macro.applescript \
  /Users/Rakesh.Patel/Documents/__code/git/rakeshpatelfa/fa_workday_etl/implementations/elt_transform_excel/dist/Runner.xlsm \
  modDispatcher.RunSelectedTransform
```

This still depends on macOS Apple Event permission to control Excel. If the host returns
`A privilege violation occurred`, run the same script from a normal user Terminal or Script Editor
session instead of this constrained CLI host.

A clean exit from this script is not proof the macro actually ran — check the Runner sheet's `Run Status`
cell (or the Output workbook path it reports) after every headless invocation. On this Mac Excel
build, a macro call has returned with no Apple Event error while `Run Status` stayed at "Not run yet,"
so absence of an error code alone does not confirm success.

## Planned Commands

No planned-but-unimplemented commands remain. `build_runner.py --verify` already covers focused
verification for every registered transform profile (currently `AR Invoice -> Customer Invoice`,
`Discipline Invoice -> Customer Invoice`, `AP AMEX Hotel -> Supplier Invoice`, and
`AP AMEX Trains -> Supplier Invoice`).

The build command above is only usable after the shell-generation command creates
`implementations/elt_transform_excel/runner_shell/RunnerShell.xlsm`.
