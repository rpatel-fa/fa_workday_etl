# Building the Runner on Windows

This is the complete, self-contained guide for building `dist/Runner.xlsm` — the finance-user
workbook — on Windows. You don't need `docs/BUILD.md` to follow it; that file is the project
owner's macOS build guide, kept only as the macOS-host equivalent for reference.

> **COLD-START PRE-READ NOTICE (read before editing ANY code/VBA/config).**
> Read `docs/AGENT_COLD_START_PROMPT_CONTRACT.md` §0.4 / §0.5 / §0.6 first — this requirement is
> the same on every host. This file changes *where you run the build*, not *what the build is
> allowed to contain*.

> **Path note (post-extraction).** Commands below are repo-root-relative and use
> `implementations/elt_transform_excel/...`. Historical docs (`BUILD.md`, `RUNBOOK.md`) show
> `elt_transform_excel/...` or `cd elt_lake` from when the module sat at the repo root — treat
> those as the same paths under `implementations/elt_transform_excel/`. See the repo `README.md`
> "Path caveat".

## Why a Windows build works at all

The Python build tooling (`check_vba_hygiene.py`, `create_runner_shell.py`, `build_runner.py`)
is pure cross-platform Python — `openpyxl` + `pyyaml` + `pyOpenVBA`, no COM, no shelling out to
Excel. It runs on Windows exactly as on macOS. `pyOpenVBA` is only used by
`create_runner_shell.py` to stamp a *fresh* `.xlsm` with a valid VBA project holding
`BootstrapImporter` (see `ARCHITECTURE.md` "Deployment"); it never rewrites an existing project
on either platform.

The one manual step — importing every `vba_src/` module via the workbook's own
`BootstrapImporter` — is the same manual step on Windows. On Windows it is actually *more*
reliable: the macOS sandbox failure modes (`com.apple.quarantine` xattrs, container-scope
`Import()` denials, the AppleScript-can't-pass-arguments limitation that forced the manual step
in the first place) simply do not exist. There is one Windows-only snag instead — a path
separator quirk in `BootstrapImporter` — which `build_runner.py` handles automatically via
`config/root_path_substitutions.yaml` (see Step 2 below).

## Prerequisites

- **Excel for Windows** (desktop; Microsoft 365 or Office 2016+). Not Excel on the web.
- **Python 3.11+** and **`uv`** (`winget install --id=astral-sh.uv`, or see <https://docs.astral.sh/uv/>).
  Per `AGENTS.md`, `uv` is the only supported Python tooling — no pip/poetry/conda.
- The repo checked out with **CRLF line endings preserved** on `.bas`/`.cls` files (the default
  on Windows git). Step 0 will fail loudly if any file is LF-only.
- A representative source workbook, target EIB template, and reference workbooks if you also
  intend to smoke-test a transform after building (see `docs/RUNBOOK.md`).

### One-time Excel setting (required)

Excel → **File → Options → Trust Center → Trust Center Settings**:

- **Macro Settings** → *Enable VBA macros* (or at minimum "Disable with notification" and then
  enable on open).
- **Macro Settings** → tick **"Trust access to the VBA project object model."**

Without the trust-access setting, `BootstrapImporter` fails immediately with
*"...cannot access the VBA project..."* / run-time error `1004` before importing anything —
it cannot add modules to the project. Close and reopen Excel after changing this.

## Runner workbook — build steps

Run all commands from the **repo root** in PowerShell. Each is a single line.

```powershell
cd C:\Users\chika\Documents\temp_fa\fa_workday_etl
```

### Step 0 — VBA source hygiene pre-flight (no Excel)

```powershell
uv run --with pyyaml --no-project python implementations/elt_transform_excel/scripts/check_vba_hygiene.py
```

If it reports non-ASCII bytes (em dashes / curly quotes / arrows in comments):

```powershell
uv run --with pyyaml --no-project python implementations/elt_transform_excel/scripts/check_vba_hygiene.py --fix
```

`--fix` only rewrites non-ASCII bytes; it does **not** convert line endings. If it reports
`LF-only line ending(s)`, your checkout stripped the CRLFs — restore them:

```powershell
git config core.autocrlf true
git add --renormalize implementations/elt_transform_excel/vba_src
git checkout -- implementations/elt_transform_excel/vba_src
```

(Or convert the working-tree `.bas`/`.cls` files to CRLF with any editor. The repo blobs are
already CRLF; the goal is only to keep them that way on checkout.)

Rules enforced: non-ASCII byte ceilings (≤ 20 WARN / > 20 ERROR), CRLF everywhere (no lone-LF),
component/stem names ≤ 31 chars, source lines ≥ 300 chars = HARD ERROR, no `Worksheets(n)`
index-based lookups, `ResolveSelectedTransformKey` call sites always take 2 args, every
`Public Sub Run<Name>` has the 2-typed-arg signature, and every `runtime.step_list_macro`
declared in YAML maps to a concrete VBA `Public Sub`. The check script is pure Python and does
not branch on OS, so these rules are identical on every build host.

### Step 1 — Create the shell workbook

```powershell
uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/create_runner_shell.py --overwrite
```

Writes `implementations/elt_transform_excel/runner_shell/RunnerShell.xlsm`.

### Step 2 — Prep the distribution workbook

```powershell
uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/build_runner.py
```

Copies the shell to `implementations/elt_transform_excel/dist/Runner.xlsm`, writes the
very-hidden transform registry, the `config` sheet defaults, the `mapping` lookup sheet, and
fills the **Bootstrap Source Folder** cell on the `Runner` sheet with the absolute path to
`vba_src`. It does **not** import any modules (that is Step 4). The macOS quarantine-xattr
stripping this script does on every build is a no-op on Windows.

This step writes the reference/transform data-root paths (`customer_spreadsheet`,
`supplier_spreadsheet`, `mapping_spreadsheet`, `source_data`, `target_eib`,
`output_file_path`) **exactly as authored** in `reference_paths.yaml`/`transforms/*.yaml` — one
person's mac paths, unchanged. That's deliberate: swapping them for your own data root is a
separate step (Step 3), not part of this build.

**Bootstrap folder path substitution.** This step does rewrite one path:
`build_runner.py` applies a build-time path-rewrite profile from
`config/root_path_substitutions.yaml` to the **Bootstrap Source Folder** value, from
`C:\...\vba_src` to `C:/.../vba_src`. The profile is auto-selected from the OS: on Windows it
selects `windows_forward_slash`; on macOS/Linux, `none` (no rewrite). The console prints the
active profile (`Path profile: windows_forward_slash`) and the rewrite it performed. Override
with `--path-profile <name>` if ever needed (`none` disables it).

> **Why the forward-slash rewrite is required.** `BootstrapImporter.DetectSeparator`
> (`vba_src/bootstrap/BootstrapImporter.bas`) picks the path separator by scanning for `/`
> first, then `:`. A Windows drive-letter path has no `/` but contains the `C:` colon, so it
> wrongly infers `:` as the separator (a leftover from classic Mac `HD:Folder:File` paths) and
> builds `C:\...\vba_src:*.*`, which matches nothing — the import fails with *"No .bas or .cls
> files found."* A forward-slash path is detected correctly, and Windows Excel VBA (`Dir`,
> `GetAttr`, `FileCopy`, `Open`, `Kill`) accepts `/` transparently. UNC paths become
> `//server/share/...`, which Windows also accepts.

### Step 3 — Root-path substitution

A separate command, run after Step 2 and before Step 4 — not merged into the build.

```powershell
uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/build_runner.py --substitute-root-paths
```

Patches the already-built `dist/Runner.xlsm` in place: rewrites the six data-root fields listed
above (both the `registry` and `config` sheets) using the same `root_path_substitutions.yaml`
profile, swapping the mac root baked into `reference_paths.yaml`/`transforms/*.yaml` for this
build host's own data root. Edit that rule's `with:` value for your machine — see its
`EDIT THIS` comment in the YAML. Does **not** re-copy the shell, and does not touch the
**Bootstrap Source Folder** cell (Step 2 already wrote that correctly). The console prints the
active profile and every rewrite it performed
(`Path substitution ... data-root fields: N value(s) rewritten`). If it reports 0 rewrites, the
rule's root already matches — that's fine, not an error. Fails clearly if run before Step 2 (no
built workbook to patch yet).

### Step 4 — Import the modules in Excel (manual)

1. Open `implementations/elt_transform_excel/dist/Runner.xlsm` in Excel. Enable macros if prompted.

2. Confirm the **Bootstrap Source Folder** cell (column A label on the `Runner` sheet, value in
   the cell beside it) shows a **forward-slash** path ending in `.../vba_src`. Step 2's path
   substitution writes it this way; if you see backslashes, Step 2 was run with
   `--path-profile none` or on a non-Windows OS — fix the cell by hand or re-run Step 2.

3. Press **Alt+F8**, select **`BootstrapImporter.BootstrapImportFromConfiguredSource`**, click **Run**.

   It imports every `.bas`/`.cls` under `vba_src` (recursively), creates the **Run Transform**
   form-control button on the `Runner` sheet, records status cells, and **saves the workbook
   itself**. On a locked-down machine you may get one Windows file-access prompt — allow it.

4. Save and close if the macro did not (it normally does).

**Alternative to the configured cell** — call the folder entry point directly from the VBE
Immediate window (**Alt+F11**, then **Ctrl+G**), always with forward slashes:

```vba
BootstrapImporter.BootstrapImportFromFolder "C:/Users/you/.../implementations/elt_transform_excel/vba_src"
```

This bypasses the cell entirely and is a good option for a scripted build.

**If import fails on a specific module** with `Method 'Import' of object '_VBComponents'
failed` — and, if the error additionally says *"Imported 'Module1'"* instead of the real module
name, the parser silently fell back to a blank module — check these causes in order; earliest is
cheapest to rule out:

0. **Module name > 31 characters.** VBA's hard limit on component names. If
   `Attribute VB_Name` / the `mod<Name>.bas` filename exceeds it, `Import` drops the name, keeps
   the blank `Module1`, and fails the rename back — surfacing as `-2147219692 (80040714)`. Not a
   mid-parse abort; check first:
   `python3 -c "import os,glob; bad=[f for f in glob.glob('vba_src/**/*.bas',recursive=True) if len(os.path.basename(f)[:-4])>31]; assert not bad, bad"`.
   Shorten the module name only; the `Run<Name>` Sub / `step_list_macro` name is not limited.
1. **Non-ASCII / Unicode content in the `.bas` file.** Documented as a macOS
   `VBComponents.Import` failure mode (its ANSI-first parser silently aborts on multi-byte UTF-8,
   even inside comments). Step 0 enforces the same byte ceiling on every host, so a file that
   passes Step 0 will not trigger this on Windows either. If you bypassed Step 0, fix by
   transcoding to ASCII equivalents (`—` → `--`, `─` → `-`, `→` → `->`, curly quotes `"`/`'` →
   straight `"`/`'`).
2. **VBA source files must be CRLF, not LF-only.** With LF-only, the parser may fail to
   recognise `Attribute VB_Name` / class-module headers and falls back to `Module1`. Step 0
   catches this — see the CRLF note under Step 0 above for the fix.
3. **Unbalanced procedure terminators** (`Sub` ... `End Sub`, `Function` ... `End Function`).
   An unterminated procedure aborts the whole import.
4. **Long line-continuation-heavy procedure signatures/calls.** Documented as a macOS import
   failure mode; not observed on Windows, but if it recurs here, shorten the signature by moving
   shared state/arguments into module-level variables or helpers.

### Step 5 — Verify

```powershell
uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/build_runner.py --verify
```

Confirms the imported module set, the transform registry, the `config`/`mapping` blocks, and
the Bootstrap Import status cells all match expectations. The `/` vs `\` difference in the
recorded source path does not matter here — `--verify` compares resolved paths, and
`C:/…/vba_src` and `C:\…\vba_src` resolve to the same location.

Then open `dist/Runner.xlsm`, check the **Run Button** status cell reads
`Created -> modDispatcher.RunSelectedTransform`, pick a Transform, and smoke-test **Run
Transform** against sample files (`docs/RUNBOOK.md` → "Watching a Run").

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `BootstrapImporter` fails instantly, before any import, mentioning the VBA project | "Trust access to the VBA project object model" is off | Trust Center → Macro Settings → tick it, restart Excel |
| *"No .bas or .cls files found under: C:\…\vba_src"* | Backslash path + the `DetectSeparator` colon quirk — Step 2's path substitution was bypassed | Re-run Step 2 on Windows without `--path-profile none`; or set the Bootstrap Source Folder cell to a forward-slash path by hand; or use the Immediate-window `BootstrapImportFromFolder` call |
| `--verify`: *"reported N module(s), expected M"* | Import stopped partway | Re-open the workbook, re-run the import macro; check the **Bootstrap Import Current** cell for the file it stalled on, then check that `.bas` against the causes list under Step 4 above (name > 31 chars, non-ASCII, CRLF, unbalanced `End Sub`) |
| Step 0: `LF-only line ending(s)` | git checkout stripped CRLF | `git config core.autocrlf true` and re-checkout `vba_src` (Step 0 above) |
| `pyOpenVBA` import error from Step 1/2/3 | dependency not resolved | Use the exact `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project ...` form; don't drop `--with` flags |
| Built workbook points at reference/transform files that don't exist on this machine | Step 3 (`--substitute-root-paths`) was never run, or `config/root_path_substitutions.yaml`'s data-root rule doesn't match this host's actual data folder | Edit the rule's `with:` value (see its own `EDIT THIS` comment) and run/re-run Step 3; check the console's "Path substitution ... data-root fields: N value(s) rewritten" line to confirm it ran |
| Step 3: *"Built workbook not found"* | Step 3 was run before Step 2 | Run Step 2 first — Step 3 patches an existing build, it does not create one |

## Optional — unattended import (PowerShell + Excel COM)

The manual Step 4 is the canonical, supported path. On Windows only, it *can* be automated,
because Windows Excel exposes a real COM automation surface that reliably passes string
arguments to a macro — the specific limitation that made this unreliable on macOS
(`osascript` "run VB macro" cannot pass arguments; see `BootstrapImporter.bas` header comment)
does not apply here. It still requires the same "Trust access to the VBA project object model"
setting.

```powershell
$src = (Resolve-Path implementations/elt_transform_excel/vba_src).Path -replace '\\','/'
$wbk = (Resolve-Path implementations/elt_transform_excel/dist/Runner.xlsm).Path

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($wbk)
    $xl.Run("BootstrapImporter.BootstrapImportFromFolder", $src)  # macro saves the workbook
    $wb.Close($true)
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
```

Note the `-replace '\\','/'` — the same `DetectSeparator` quirk applies, so the path handed to
the macro must use forward slashes. Follow with Step 5 (`--verify`) as usual. Do **not** point
any tool (`pyOpenVBA` included) at the VBA project of an existing `dist/Runner.xlsm` to add
modules directly — that path is rejected on both platforms (`RUNBOOK.md` → "Diagnosis").

## Ingest template

Render the template a finance team populate by hand, from its ingest spec. One step, no Excel,
no VBA — nothing here touches the Runner build above.

```powershell
uv run --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/create_ingest_template.py ar_customer --overwrite
```

`config/ingest_specs/ar_customer.yaml` → `dist/ingest_templates/ar_customer.xlsx`. The argument
is the spec key (any file under `config/ingest_specs/`); a path to a spec YAML also works. Add
`--sample-row` to also fill row 2 from each column's `example`. Columns with a `default` are
pre-filled either way.

The spec is the source of truth and the workbook is a render of it — **never hand-edit the
generated `.xlsx`**, as the next render silently discards the edit. The render is deterministic
(identical spec in, byte-identical workbook out), so re-rendering an unchanged spec is a no-op
you can run freely.

For the edit → render → review loop and what the header colours mean, see `docs/ops/CLI.md`
("Ingest Template Prototyping Loop"). For the spec schema, see `docs/CONFIGURATION.md` ("Ingest
Specs").
