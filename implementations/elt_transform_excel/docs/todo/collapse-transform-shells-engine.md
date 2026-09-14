# Collapse forked transform shells → one config-driven engine

> This is the concrete, phased execution of the previously-deferred **Stage 2 orchestration
> codegen** (see `todo/archive/simplify-transform-boilerplate.md`) and the **H3 scaling trigger**
> recorded in `docs/ADDING_A_TRANSFORM.md § Scaling note` (≥20 transforms / ≥4h per transform →
> move from per-transform `.bas` to a table-driven engine). Opened 2026-08-25.

## Resume (start here)

- **Status:** NOT STARTED. Plan only; no code moved yet.
- **First concrete slice split out:** `docs/todo/unify-invoice-line-grouping.md` (2026-09-14) —
  promote the `ap_stationery`/`ap_taxi` multi-line grouping (Spreadsheet Key / Line Order /
  Row ID) into one shared, config-driven concern and retire the per-row single-line
  assumption across the other invoice transforms. That doc also carries the outcome-review
  findings (`scripts/audit_outputs.py`) and the two Workday-tax/AMEX decisions gating it.
- **Next item:** Phase 0 — build the output-diff safety harness. Nothing else starts until it exists.
- **Read before touching code:** `docs/ARCHITECTURE.md` (config-authoritative transform contract and
  current parity model), `docs/ADDING_A_TRANSFORM.md`, and
  `config/shared_helper_registry.yaml` (the shared-helper concerns already extracted).
- **Why:** the ~10 AP supplier-invoice `.bas` modules are near-duplicate forks (`ap_travel_places`
  is literally "a fork of `ap_taxi`"). The row loop, reject handling, and write block are copy-pasted,
  so a single behaviour fix (v46.2 columns, delimiter normalisation, zero-value skip) had to be
  repeated across 6–10 files this session. The `ap_stationery` variants already prove the target
  model (4 suppliers, ONE `.bas`, config-differentiated). Goal: extend that model family-wide so a
  new transform is mostly a config entry.

## Done

- (none)

## Still todo (next actions)

1. **Phase 0 — output-diff safety net (BLOCKING; do first).** Build a script that, for a given
   transform, produces its EIB and asserts byte-identical output vs a committed golden copy
   (compare the populated `Submit Supplier Invoice` sheet cell-by-cell from data row 6). Requires a
   way to capture "golden" outputs from the current shipped macros. Nothing in Phase 1+ proceeds
   without this — it is the only thing that makes the refactor provably safe on shipped code.
2. **Phase 1 — extract ONE family, don't rewrite.** Lift the shared orchestration (row loop,
   `RowIsEmptyOrTrailer`/`RowIsSkippable` handling, the write block, reject/summary counters) out of
   `ap_taxi` + `ap_travel_places` into a shared engine module both call. Leave per-transform only the
   deltas (source layout, which columns, resolution mode). Gate on Phase-0 harness: output
   byte-identical before/after for BOTH, or revert.
3. **Phase 2 — migrate the classic-6 one at a time** (`ap_taxi`, `ap_stationery`, `ap_travel_places`,
   `ap_cfa_claims`, `ap_cfa_distributions`, `ap_referees_v2`), each output-validated by the
   harness. Never more than one transform's output at risk per step.
4. **Phase 3 — config-drive the deltas.** Express the remaining per-transform differences (column map,
   source columns, skip policy) as config the engine reads at build/run time — approaching the
   `ap_stationery` variant model, where a new transform is a config entry, not a new `.bas`.
5. Update `docs/ADDING_A_TRANSFORM.md` + `ARCHITECTURE.md` (Authoring model) as the model changes;
   retire the H3 trigger note once the engine exists.

## Accumulated Active Constraints (active for all remaining items)

- **Byte-identical output is the pass/fail gate.** Every phase step must produce EIB output identical
  to the shipped macros for every affected transform (Phase-0 harness). This is a refactor, not a
  behaviour change — no output may move.
- **One transform's output at risk at a time.** No big-bang. Each step is independently revertible.
- **CRLF-only for `.bas`.** Mac Excel's VBA importer aborts on LF-only lines; any programmatic `.bas`
  write must re-normalise to CRLF (`b.replace(b"\r\n",b"\n").replace(b"\n",b"\r\n")`).
- **≥300-char `.bas` lines are a hard build error**; keep engine lines under ~250 with `_` continuation.
- **Spec↔code column gate is now enforced** (`build_runner.py _parse_spec_mapping_targets`): every
  transform_spec `mapping:` target must be written by its module. The engine must keep this true —
  either per-transform specs still enumerate targets, or the gate is taught the engine's indirection.
- **Shared-helper anti-drift is enforced** (`shared_helper_registry.yaml` + POS/NEG guards, incl.
  comment-stripping): no local `Private` copies of hoisted helpers; declared variants must be really
  called. The engine consolidates callers — keep the declarations honest.
- **`safe_shared_columns` + column contracts live in YAML** (`config/eib_population_defaults.yaml`),
  not code. Keep new shared columns there.
- **1 mandatory manual Excel step remains** (`BootstrapImporter.BootstrapImportFromConfiguredSource`);
  the engine does not remove it.

## Verification

- Build (all gates): `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/build_runner.py` → "Prep complete", 0 hygiene errors, no `SPEC<->CODE DRIFT`.
- Gates cross-check: `uv run --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/verify_gates.py` → "Bindings parsed: 15".
- Per-transform input contract: `uv run --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/validate_input.py -t <profile>` → GREEN.
- **Phase 0 output-diff harness (to be built): `<harness> --transform <profile>` → BYTE-IDENTICAL vs golden.** This is the gate for every Phase 1+ step.

## Gotchas

- **YAML `N`/`Y` → boolean.** PyYAML 1.1 coerces bare `N`,`Y`,`yes`,`no`,`on`,`off` to booleans; a
  bare `target: N` becomes `False`. Quote column letters in YAML, and prefer parsing spec targets
  from raw text (as the spec↔code gate does), not `yaml.safe_load`.
- **transform_specs are only partially build-read today.** Shared-helper declarations, mapping targets,
  the invoice-line grouping contract, and parked invoice-grouping decisions are enforced. The rest of
  the transform-spec surface still needs broader parity coverage before the repo can claim full
  repo-wide config->code proof.
- **`ap_stationery` is the reference implementation** for the target model (one `.bas`, N config
  variants under `shell_defaults.variants`). Study it before designing the engine.
- **Accounting-string normalisation is already centralised** in the shared resolvers
  (`ResolveLowestHierarchyLookupAcrossFinalCodeSegments` + `FinalCodeSegment`) — the engine inherits
  delimiter-robustness for free; do not re-add per-transform normalisation.

---
_Resume prompt: `From docs/todo/collapse-transform-shells-engine.md: Continue → Still todo step 1 (Phase 0 output-diff harness)`_
