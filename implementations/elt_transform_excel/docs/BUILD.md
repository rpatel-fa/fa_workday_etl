This module has two independent build outputs. Both are built from the repo root.

> This file assumes the project owner's **macOS** build host. To build `dist/Runner.xlsm` on
> **Windows** instead, see `docs/BUILD_WIN.md` — same rules and same manual import step, with the
> host-specific mechanics (Trust Center setting, a `BootstrapImporter` path-separator quirk)
> spelled out.

> **COLD-START PRE-READ NOTICE (read before editing ANY code/VBA/config).**
> Before this file, read `docs/AGENT_COLD_START_PROMPT_CONTRACT.md` in full, especially:
> - §0.4 MANDATORY CANONICAL READ ORDER (config-first timeline — what to read in what order before writing VBA)
> - §0.5 MANDATORY VBA PRE-EDIT GATE (G1 check_vba_hygiene.py → G2 build_runner --hygiene-only → G3 post-edit re-run)
> - §0.6 SHARED HELPER AUTHORING CONTRACT (6 build-enforced guards for modSourceTable.bas / modFinanceCommon.bas changes).
> If your workflow doesn't start with these three sections, the timeline will be inverted (VBA written before
> config/specs) — this is a DEFECT, regardless of whether the build eventually passes. Re-read, then start.

| Output | Built from | Needs Excel? |
|---|---|---|
| `dist/Runner.xlsm` — the workbook finance users run | `vba_src/` + `config/transforms/` + `config/ingest_lookups/ref_company.json` & `ref_default.json` (the runtime `mapping` sheet) | Yes — one manual import step |
| `dist/ingest_templates/<key>.xlsx` — the input template finance users fill in | `config/ingest_specs/<key>.yaml` | No |

## Runner workbook

Build `elt_transform_excel/dist/Runner.xlsm` from the repo root:

**Step 0 — Run the VBA source hygiene pre-flight (fail fast before Excel):**

```bash
# Fastest (stdlib only, ~25 ms, no pyOpenVBA, no Excel):
python3 elt_transform_excel/scripts/check_vba_hygiene.py

# If NON_ASCII warnings appear (em dashes, curly quotes, arrows in comments):
python3 elt_transform_excel/scripts/check_vba_hygiene.py --fix
```

Rules enforced: non-ASCII byte ceilings (≤ 20 WARN / > 20 ERROR — anything over 20 bytes reliably causes `VBComponents.Import` to silently fall back to a blank `Module1` on Mac Excel), CRLF everywhere (no lone-LF), component/stem names ≤ 31 chars, source lines ≥ 300 chars = HARD ERROR (Mac VBA parser observed to abort mid-import on 258+ char lines), no `Worksheets(n)` index-based lookups, `ResolveSelectedTransformKey` call sites always take 2 args, every `Public Sub Run<Name>` has the 2-typed-arg signature, and every `runtime.step_list_macro` declared in YAML maps to a concrete VBA `Public Sub`.

1. Create the shell workbook:

```bash
uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/create_runner_shell.py --overwrite
```

2. Prep the distribution workbook:

```bash
uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/build_runner.py
```

2.1 Validate (post macro setup in step 3)

```bash
cd ~/Documents/__code/git/emailrak/elt_lake
uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/build_runner.py --verify
```

3. In Excel:

   - **One-time setup:** Excel → Preferences → Security. Set **Macro Security → Enable all macros**, and
     under **Developer Macro Settings** tick **"Trust access to the VBA project object model"**. Then quit
     Excel (⌘Q) and reopen. Without the trust-access setting the macro fails with *"Excel could not access
     this workbook. VBAProject/VB Components. Enable trust access to the VBA Project object model…"* — the
     importer cannot add modules to the VBA project. (These are Excel settings, separate from macOS Full
     Disk Access.)
   - open `elt_transform_excel/dist/Runner.xlsm`
   - run `BootstrapImporter.BootstrapImportFromConfiguredSource`
   - save and close

   If Excel raises `Method 'Import' of object '_VBComponents' failed`, check the source module before
   retrying, in order. (If the error message additionally says *"Imported 'Module1'"* instead of the
   expected module name, either the module name is too long (cause #0 — check this first, it is the
   cheapest and fully deterministic) or the VBA parser silently gave up mid-parse — the latter is
   almost always causes #1 or #2 below.)
   0. **Module name > 31 characters.** VBA has a hard 31-char limit on component (module) names. If
      `Attribute VB_Name` / the `mod<Name>.bas` filename exceeds it, `Import` drops the name, keeps the
      blank `Module1`, and fails the rename back — surfacing as `-2147219692 (80040714)`. Unlike #1/#2
      this is not a mid-parse abort; it is a name-length limit, checkable before import:
      `python3 -c "import os,glob; bad=[f for f in glob.glob('vba_src/**/*.bas',recursive=True) if len(os.path.basename(f)[:-4])>31]; assert not bad, bad"`.
      Shorten the module name only; the `Run<Name>` Sub / `step_list_macro` name is not limited.
   1. **Non-ASCII / Unicode content in the `.bas` file.** Mac Excel's `VBComponents.Import` parser is
      ANSI-first and silently aborts (falling back to a blank `Module1`) when multi-byte UTF-8 sequences
      appear in volume, even if they are only inside comment blocks. Pure-ASCII modules (0 non-ASCII
      bytes) always import. Modules with a tiny amount (≤ ~20 bytes = 2–3 em dashes, e.g.
      `modHelpers.bas`) are tolerated; modules with hundreds of non-ASCII bytes (e.g. decorative
      box-drawing chars `─`, em dashes `—`, arrows `→` in section-header comments) fail. Fix: replace
      any Unicode with ASCII equivalents (`—` → `--`, `─` → `-`, `→` → `->`, curly quotes `"`/`'` →
      straight quotes `"`/`'`) and confirm the file is 100% pure ASCII.
   2. **VBA source files must be CRLF, not LF-only.** With LF-only, the parser may fail to recognise
      `Attribute VB_Name` / class-module headers and again falls back to `Module1`. Convert with
      `sed -i '' 's/$/\r/' file.bas` (or a Python-based writer) and verify with `file(1)`.
   3. **Balanced procedure terminators** (`Sub` ... `End Sub`, `Function` ... `End Function`). An
      unterminated procedure aborts the whole import.
   4. **On Mac Excel, long line-continuation-heavy procedure signatures/calls** can fail import even
      when syntax is formally valid; shorten them by refactoring shared state/arguments into simpler
      helpers or module-level variables.
4. Verify the built workbook:

```bash
uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/build_runner.py --verify
```

## Ingest template

Render the template a finance team populate by hand, from its ingest spec. One step, no Excel, no VBA
— nothing here touches the Runner build above.

```bash
uv run --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/create_ingest_template.py ar_customer --overwrite
```

`config/ingest_specs/ar_customer.yaml` → `dist/ingest_templates/ar_customer.xlsx`. The argument is the
spec key (any file under `config/ingest_specs/`); a path to a spec YAML also works. Add `--sample-row`
to also fill row 2 from each column's `example`. Columns with a `default` are pre-filled either way.

The spec is the source of truth and the workbook is a render of it — **never hand-edit the generated
`.xlsx`**, as the next render silently discards the edit. The render is deterministic (identical spec
in, byte-identical workbook out), so re-rendering an unchanged spec is a no-op you can run freely.

For the edit → render → review loop and what the header colours mean, see `docs/ops/CLI.md`
("Ingest Template Prototyping Loop"). For the spec schema, see `docs/CONFIGURATION.md` ("Ingest Specs").
