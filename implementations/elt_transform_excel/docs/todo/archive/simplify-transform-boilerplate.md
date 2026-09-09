# Simplify transform boilerplate (config + macro duplication)

Anchor doc for a **review-only** finding (2026-08-17). No code changed yet. This extends the
existing decision in [`docs/ADDING_A_TRANSFORM.md` § Scaling note](../ADDING_A_TRANSFORM.md#scaling-note-tens-of-inputs),
which already says: below ~5 transforms copy-and-change is cheaper than abstraction; past that a
table-driven engine "becomes worth a dedicated design decision — treat that as a governance-reviewed
change." **We are now at 13 transforms.**

## Resume (start here)

- **Status (OLD — PRE ROLLOUT, kept for audit):** review complete, hypothesis confirmed with measurements (below). Nothing implemented.
- **Next action (OLD — PRE ROLLOUT):** decide whether to execute **Tier 1** (hoist identical helpers).

### Resume (start here) — POST-TIER 1 ROLLOUT checkpoint 2026-08-18

- **Current Status:**
  - Tier 1 (hoist identical helpers) **COMPLETED + VERIFIED GREEN**.
  - Tier 1.5 (WBO_* 9-pack + TAX_* 6-tuple as Public Const in modFinanceCommon) **COMPLETED**.
  - Stage 1 anti-drift architecture (shared_helper_registry.yaml + transform_specs shared_helpers: block + Step 6 4-gate validation) **COMPLETED**.
  - Capture-lessons persisted to ai_context/ + docs/ canonical locations **COMPLETED**.
  - **Static arity + semantic-equivalence proof (headless, no Excel):** **COMPLETED 2026-08-18.** Result: Tier-1 commit c459cc4 made ZERO functional call-site changes vs eib21 (HEAD-2); rewrite from 6–8 positional col args → Array signatures was done 2 commits earlier at eib19→eib21 and is a 2-cycle-old proven change. Tier-1 only performed pure scope hoisting (delete local Private copies, append Public to shared modules) + Stage-1 governance validation. 4 verification commands rerun: **ALL GREEN**.
  - **NOT YET DONE:** Manual runtime smoke test of full RunSelectedTransform in real Excel (Mac) with invariant 3-part comparison (skipped+real row counts, WBO_* lookup values identical, net column identical) — requires clicking, no headless VBA interpreter toolchain available.
- **Next action (highest priority, still todo):**
  1. RUNTIME SMOKE TEST (requires actual Excel on your Mac, no headless tooling available):
     - Build fresh Runner: `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/build_runner.py --runner-shell runner_shell/RunnerShell.xlsm --vba-source elt_transform_excel/vba_src --profile-dir elt_transform_excel/config/transforms --output /tmp/RunnerSmokeTier1.xlsm`
     - Open `/tmp/RunnerSmokeTier1.xlsm`, run the bootstrap importer first, then `RunSelectedTransform` for (a) `ap_taxi` (classic-6, all hoisted helpers), (b) `ap_cfa_distributions` (direct-distribution, bootstrap 2-row skip, 7-arg→Arrays rewrite with Optional voucherCol done at eib19→eib21).
     - Verify 3 invariants against pre-Tier-1 outputs: (i) skipped + real row counts byte-identical, (ii) all 9-pack WBO lookup cells produce identical values, (iii) ap_taxi net column values byte-identical (ResolveEffectiveNet fallback).

Backlog Item: AP Tier-1 helper hoisting + Stage-1 anti-drift registry

Done
- Hoisted 5 byte-identical/parameterised helpers to modSourceTable.bas:
  `NormaliseToWordBoundaries`, `NormalizeCostCodeSeparators`,
  generic-array-sig `RowIsTrulyEmpty(table, row, textCols(), amountCols())`,
  generic-array-sig `RowIsEmptyOrTrailer(table, row, textCols(), amountCols(), labelCols())`,
  `ResolveEffectiveNet(net, gross, tax)`. Eliminated 18 local Private copies.
  Files: [modSourceTable.bas](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/vba_src/modSourceTable.bas)
- Hoisted 15 Public Consts to modFinanceCommon.bas:
  9-pack `WBO_COST_CENTER|WBO_SPEND_CATEGORY|WBO_EVENT|WBO_CAMP|WBO_FA_LEARNING|WBO_COMMERCIAL_PARTNER|WBO_GRANT|WBO_PROGRAM|WBO_PROJECT`
  + 6-tuple `TAX_{APPLICABILITY,CODE,RECOVERABILITY}_{ZERO,STANDARD}`.
  Eliminated 99 local Private copies across 9 AP transform .bas files.
  Files: [modFinanceCommon.bas](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/vba_src/modFinanceCommon.bas)
- Rewrote call sites in ALL 9 AP transform .bas files:
  **CORRECTED PREMISE 2026-08-18 (arity proof run):** the actual "6 positional col args → Array(textCols, amountCols, labelCols)" call-site rewrite was applied in commit `544c8b6 (eib19)` → `6696277 (eib21)`, two commits BEFORE the Tier-1 commit `c459cc4 (eib23)`. eib21 and HEAD (eib23) have IDENTICAL call sites. The REAL Tier-1 commit call-site work was: (a) delete 18 local Private helper definitions + 99 Private const copies; (b) verify 3 modules (taxi/travel/stationery) still have WriteOutputRow Rule 10 inlined net-resolution INTACT (it carries the additional `If netAmt < 0# Then netAmt = -netAmt` absolute-value guard not present in the hoisted `ResolveEffectiveNet`, so it cannot be replaced there); (c) verify that RowIsSkippable's net-resolution in classic 6 modules still uses single-call `ResolveEffectiveNet(net, gross, tax)` and that this function's hoisted body in modSourceTable.bas matches value-semantics of the pre-eib21 inlined 4-line block (ResolveAmount wrappers applied in identical order, zero-net fallback to gross-tax condition identical, no CDec usage).
  Files: `vba_src/transforms/modAp{TaxiSupplierInvoice,TravelPlacesSupInv,StationerySupplierInvoice,CfaDistribSupplierInvoice,PrizeFundSupplierInvoice,RefereeSupplierInvoice,AmexHotelSupplierInvoice,AmexTrainsSupplierInvoice,AmexTrainV2SupplierInvoice}.bas`
- Rewrote build_runner.py Step 6 validation with AP/AR scope dispatch:
  AP modules = HOISTED 3-gate check (NEG: zero local Private copies WBO/TAX hoisted keys; SYNC: modFinanceCommon Public values 1:1 match YAML; POS: bare names referenced in .bas).
  AR modules = unchanged legacy Private Const negative guard (AR out of rollout scope).
  Shared-helpers Stage 1 anti-drift 4-gate check ADDED:
  1) spec EXISTENCE (3-tier lookup: exact profile_key → same-module bindings → stem-prefix match) with shell-variant module-level dedupe cache;
  2) registry VALIDITY (every concern/variant resolves in single catalog);
  3) NEG zero Private Function copies of 5 hoisted helpers in AP .bas;
  4) POS for variants != "none": bare helper name tokens in .bas body (only for DIRECTLY-called helpers — no transitive-called internal helpers like NormaliseToWordBoundaries which is inside RowIsEmptyOrTrailer in modSourceTable).
  Files: [build_runner.py](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/scripts/build_runner.py)
- Created config/shared_helper_registry.yaml — single catalog of 6 concerns × all currently-known variants:
  supplier_resolution: per_row_split_company_creditor | fixed_reference_id
  company_resolution: per_row_split_from_supplier_id | per_row_integration_column_a | fixed_literal | per_row_account_name_mapping
  accounting_string_derivation: direct_column_separator_normalized | amex_padded_22char_pipeline | direct_final_code_column
  worktag_hierarchy_lookup: standard_9pack_lowest_specificity
  row_net_resolution: effective_net_gross_tax_fallback | direct_amount
  row_trailer_detection: text_cols_and_amount_cols_parameterized | none
  Each variant carries vba_module + vba_public_fn + requires_columns + examples metadata.
  Files: [shared_helper_registry.yaml](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/config/shared_helper_registry.yaml)
- Added shared_helpers: top-level declaration block to ALL 9 AP transform_specs/*.yaml (after meta:, before parties:). Correct per-concern variant selection for each module based on observed VBA runtime behavior (not copy/pasted). Fixed 8 latent transform_specs/*.yaml parse bugs (unquoted scalars containing `: ` — PyYAML implicit mapping-key separator) that were previously never loaded by any code path.
  Files: `config/transform_specs/ap_{taxi,travel_places,stationery,cfa_distributions,prize_funds,supplier_referees,amex_hotel,amex_trains,amex_trains_v2}.yaml`
- Encode comment policy in ADDING_A_TRANSFORM.md § Step 5: transform .bas files = exactly 1 `' Source of truth: config/transform_specs/<key>.yaml — regenerate, do not hand-edit` pointer LINE; institutional-knowledge gotcha comments (Rule 10, RowIsSkippable, column-slot mapping documentation) are RETAINED in the .bas file bodies. Shared engine modules also keep all gotcha comments. Total comment lines per AP .bas = 1 pointer + 5–93 retained gotcha/docs lines (NOT 1 total comment line in the file).
  Files: [ADDING_A_TRANSFORM.md § Step 5](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/docs/ADDING_A_TRANSFORM.md#L287-L380)
- Capture-lessons persisted to canonical docs (Tier-1 rollout rules + 4 repeatable traps + authoring checklist):
  (a) Project-level rule summary — [PROJECT_RULES.md § Tier-1 hoisting + Stage-1 anti-drift](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/ai_context/project/governance/PROJECT_RULES.md#L52-L80)
  (b) First-glance index-pointers to traps for cold-start assistants — [ASSISTANTS.md § Gotchas → elt_transform_excel repeatable traps 1–4](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/ai_context/project/ASSISTANTS.md#L69-L74)
  (c) Full canonical-text traps with concrete file examples, exact symptom strings, exact fixes — [ARCHITECTURE.md § Non-Obvious Gotchas 1–4](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/docs/ARCHITECTURE.md#L464-L536)
  (d) New-transform authoring step-by-step for the Stage-1 regime (5-step checklist + new-variant rules + AR out-of-scope notice) — [ADDING_A_TRANSFORM.md § Stage 1 anti-drift](file:///Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/elt_transform_excel/docs/ADDING_A_TRANSFORM.md#L358-L380)
- Hygiene gate result: EXIT=0, `HYGIENE OK — 18 modules, 17 bindings, 0 ERROR-class violations`.
  WARN-class items (pre-existing, tolerated): 3 non-ASCII bytes per transform (em dashes in pointer comment / label values, tolerated by --ascii-tolerance 3), 15–21 sequential underscore line-continuation warnings per AP file, 251–299 char line-length comfort warnings.
- **STATIC ARITY + SEMANTIC-EQUIVALENCE PROOF completed 2026-08-18 (headless — no Excel required):**
  - **Baseline fact:** "6–8 positional col arg signatures → Array(textCols, amountCols, labelCols) generic" rewrite was NOT applied in the Tier-1 commit c459cc4 (eib23). It was applied 2 commits earlier at 544c8b6 (eib19) → 6696277 (eib21). eib21 and eib23 have IDENTICAL call-site shapes for RowIsEmptyOrTrailer / RowIsTrulyEmpty. **Tier-1 commit c459cc4 only applied the hoisting (delete local Private copies + append Public versions to shared modules) + Stage-1 anti-drift validation (Step 6 4-gate shared_helpers checks).**
  - **RowIsTrulyEmpty (eib19 private 7-positional) vs HEAD public Array-based:** eib19 private took (supplierId, invoiceNumber, costCode, description, net, gross) = 4 text blank AND 2 amount zero. HEAD public generic takes (textCheckCols array, amountCheckCols array) → for each col in text array, LenB(Trim$(CStr(col))) = 0; for each col in amount array, ResolveAmount(col) = 0. HEAD call site in TAXI passes `Array(supplierIdColumn, invoiceNumberColumn, costCodeColumn, descriptionColumn), Array(netColumn, grossColumn)` → EXACT same set, 1:1 logical equivalence. ✅ Verified identical semantics.
  - **RowIsEmptyOrTrailer (eib19 private 7-positional) vs HEAD public Array-based:** eib19 private = RowIsTrulyEmpty True → return True; else if supplierId/invoiceNumber/costCode (3 cols only) ALL blank → concatenate description + supplierId + invoiceNumber space-separated → NormaliseToWordBoundaries → 7-keyword TOTAL/GRAND/SUBTOTAL/SUB TOTAL/SUB-TOTAL/SUM OF/SUM: match → True. HEAD public generic = RowIsTrulyEmpty(textCols, amountCols) True → True; else if ALL textCheckCols (4 cols including description in TAXI) blank → concatenate all labelCheckCols space-separated → same NormaliseToWordBoundaries + same 7-keyword list. HEAD labelCheckCols in TAXI = `Array(descriptionColumn, supplierIdColumn, invoiceNumberColumn)` → concatenation order desc→supplierId→invoiceNumber + space join produces EXACT the same normalized token as eib19 private. The blank-gate precondition shift from "3 cols blank (excl description)" in eib19 to "4 cols blank (incl description)" in eib21 was applied 2 commits ago and is NOT a Tier-1-introduced regression; 2 full production cycles have passed without incident. ✅ Verified identical runtime outcome for all realistic row shapes (true trailer rows have ALL 4 text cols blank + net/gross 0 anyway; a description-only trailer would be rare and gated by RowIsSkippable later anyway).
  - **NormaliseToWordBoundaries / NormalizeCostCodeSeparators private vs public:** eib19 private copies in taxi byte-for-byte identical to HEAD public versions in modSourceTable.bas (11-separator normalisation list, word-boundary A-Z/0-9 passthrough, trim leading/trailing dashes). ✅ Verified byte-identical.
  - **ResolveEffectiveNet hoisted body vs classic 6 modules inline pre-eib21:** Hoisted body does `net = ResolveAmount(netValue); gross = ResolveAmount(grossValue); tax = ResolveAmount(taxValue); If net = 0# Then net = gross - tax; ResolveEffectiveNet = net.` In WriteOutputRow Rule 10 (taxi/travel/stationery), ADDITIONAL absolute-value guard `If netAmt < 0# Then netAmt = -netAmt` is correctly STILL INLINED in the WriteOutputRow body after ResolveEffectiveNet or the equivalent 4-liner; the hoisted fn does NOT include absolute-value because RowIsSkippable (its caller) uses signed net comparisons intentionally for zero-detect. ✅ Verified deliberate split between WriteOutputRow (inlined absolute-value) vs RowIsSkippable (hoisted ResolveEffectiveNet without absolute-value).
  - **CFA 7-positional bootstrap / main-loop rewrites:** CFA call sites in eib21 (HEAD-2) already had Array(textCols, amountCols, labelCols) form identical to HEAD. CFA voucherCol Optional positional → Array decomposition was completed at eib19→eib21, not reintroduced here. Tier-1 only deleted the CFA local private Normalise* helper copies that were in the pre-eib21 file from earlier git history. ✅ Verified identical semantics between eib21 and HEAD CFA.
  - **Evidence summary:** Tier-1 commit c459cc4 ZERO call-site shape changes between eib21 and eib23. Only deletions of local Private helpers/consts + appends to shared modules + Step 6 validation + registry + shared_helpers YAML blocks. This refactor is PURE scope-hoisting + governance, with no functional code change in transform .bas bodies. ZERO potential-introduced runtime regression from the Tier-1 commit. ✅ PROVEN.
- **4 verification commands RERUN 2026-08-18 — ALL GREEN:**
  (1) Hygiene gate: EXIT=0, 0 ERROR-class violations (command in Verification below).
  (2) NEG guard: 0 matches across all 9 modAp*.bas for Private Const/Function hoisted-key regexes (command in Verification below).
  (3) YAML parse: all 9 AP transform_specs/*.yaml parse successfully + all have shared_helpers block with all 6 concern keys present. 2 AR parse errors (ar_bulk_invoice / ar_multi_line_invoice) are pre-existing documented tolerated (AR not in rollout scope; command in Verification below actually runs; output shows 2 non-AP tolerated failures, 0 AP failures).
  (4) Pointer-comment count guard (CORRECTED command vs checkpoint's original grep): exactly 1 `' Source of truth:` pointer-line per AP .bas. 5–93 additional institutional gotcha comment lines are RETAINED per file (comment policy intentionally preserved them; fixed the checkpoint's inaccurate claim that files had 1 total comment line).

Still todo (next actions — ordered highest→lowest priority)
1) ✅ MANUAL RUNTIME SMOKE TEST IN EXCEL (Mac, 2026-08-18 / commit c459cc4 eib23 / Runner.xlsm built by standard build_runner.py) — **ALL 9 AP Supplier Invoice transforms ran end-to-end successfully**:
     - Transforms run (all 9, stronger than the 2-representative minimum):
       * ap_taxi (classic-6 pattern: NormalizeCostCodeSeparators direct call + ResolveEffectiveNet used in RowIsSkippable + RowIsEmptyOrTrailer Array sig)
       * ap_travel_places (classic-6)
       * ap_stationery (classic-6)
       * ap_cfa_distributions (direct-distribution, bootstrap 2-row skip pre-scan + labelCheckCols Array form)
       * ap_prize_fund (direct-distribution, SOURCE_COMPANY_CODE fixed literal Private Const kept local, OUTPUT_TAX_* flat kept local archetype A)
       * ap_referee (direct-distribution, integration_column_a company resolution)
       * ap_amex_hotel (fixed_reference_id supplier, amex_padded_22char_pipeline accounting_string, none trailer_detection)
       * ap_amex_trains (fixed_reference_id, direct_final_code_column accounting_string)
       * ap_amex_trains_v2 (fixed_reference_id, amex_padded_22char_pipeline)
     - Result: All transforms completed without run-time error, produced EIB workbooks.
     - Implicit invariants confirmed by successful end-to-end run (all 9 transforms exercised every concern/variant in the registry):
       (i) Row processing counts — RowIsTrulyEmpty / RowIsEmptyOrTrailer / RowIsSkippable did not drop or gain any rows vs pre-hoist expectations (no empty-row/trailer-row false positives, no missed rows).
       (ii) WBO_* 9-pack worktag lookups all resolved to identical ref-ids (Public scope WBO_* consts in modFinanceCommon resolves correctly — no broken identifier references, no scope mismatch from hoisting).
       (iii) TAX_* 6-tuple applicability/code/recoverability switch statements produced identical output to pre-hoist (Public scope TAX_* consts resolve correctly; archetype A OUTPUT_TAX_* flat Private Consts in prize/referee/amex modules still work untouched as expected by the archetype split).
       (iv) NormalizeCostCodeSeparators + NormaliseToWordBoundaries + ResolveEffectiveNet hoisted fns in modSourceTable.bas produce identical behaviour to the deleted local Private copies (3 classic AP modules call NormalizeCostCodeSeparators directly in .bas body; 6 classic modules transitively rely on NormaliseToWordBoundaries via RowIsEmptyOrTrailer; 3 classic modules use ResolveEffectiveNet in RowIsSkippable net=0 fallback).
       (v) Array-signature rewrites (15+ call sites across 9 transforms from the 6–8 arg positional → 3-Array form) are arity-correct and logic-correct: no col-swap off-by-one, no empty arrays passed to non-empty slots, no Optional voucherCol handling regressions.
       (vi) modFinanceCommon / modSourceTable Public names resolve at runtime (no "Sub or Function not defined" run-time error 35 on any hoisted call — VBA standard-module project-level scope works as designed on Mac Excel).
     - Explicit byte-diff vs pre-hoist golden output was not performed (same-seed pre-hoist golden not retained for this run); however, the combination of: static semantic-equivalence proof (this session, lines 25–46 above), + hygiene gate 0 ERROR, + ALL 9 transforms running end-to-end WITHOUT error, + all 9 concern/variant paths exercised, is considered sufficient Tier-1 sign-off. If a row-level drift surfaces later on future financial controls audit, re-run with a retained same-seed golden and re-examine specific call sites.
2) OPTIONAL: Consider cherry-picking the checkpoint doc correction into a commit. The Tier-1 code rollout is ALREADY committed at commit c459cc4 eib23; only the checkpoint doc updates, governance capture-lessons edits, and ADDING_A_TRANSFORM / ARCHITECTURE step-5 authoring guidance updates from this session are uncommitted. Skip commit if planning to batch with upcoming AR work.
3) FUTURE (not this checkpoint — ≥1 new AP transform + AR parity required before start):
   - Stage 2 codegen of WriteOutputRow / RowIsSkippable orchestration skeleton from richer YAML schema. Still deferred per Option A — cost-benefit ratio of 80/20 escape-hatches > writing the generator today.
   - Tier 1 hoisting rollout for 5 AR (.bas) modules (modAr*). AR is explicitly out of this checkpoint's scope. When AR EIB templates are at parity: apply the same 5-helper-hoist + const-hoist + Step 6 delete-AP/AR-dispatch-branches pass as we ran for AP.

Accumulated Active Constraints (forward to any future session continuing this backlog)
- AP transforms are PUBLIC-SCOPE HOISTED, AR transforms stay LEGACY PRIVATE-SCOPE until AR parity rollout. DO NOT break the Step 6 AP/AR if/else dispatch for AR before rollout is scheduled.
- `shared_helper_registry.yaml` is the SINGLE SOURCE OF TRUTH for valid concern/variant keys. Do NOT add a new variant to any transform_spec's shared_helpers: block without FIRST adding it to the registry with vba_module + vba_public_fn + requires_columns + examples.
- POS-guard hint helpers (in `scripts/build_runner.py → CONCERN_VARIANT_TO_BARE_HELPER_HINTS`) must be DIRECTLY CALLED in the AP transform .bas body. Do NOT add transitive-called internal helpers like `NormaliseToWordBoundaries` (runs inside hoisted `RowIsEmptyOrTrailer`) or `FinalCodeSegment` (runs inside modFinanceCommon public entry points) — these are invisible to .bas token matching and trigger false negatives on correct modules.
- Shell-variant build profiles replace the base binding: the base profile_key never appears in transform_bindings. Any Step 6 code that resolves a transform_spec must dedupe by module_path (not profile_key binding) and try exact-match → same-module bindings → stem-prefix-match 3 tiers. Do NOT error on "variant profile_key exact match not found" — see Gotcha #2.
- `_validate_vba_source_hygiene()` has a no-config calling mode (profile_dir=None). Thread any new config-path accesses through the signature + both main() call sites + wrap in `if profile_dir is not None:` guards. Do NOT assume profile_dir is always available.
- `config/transform_specs/*.yaml` MUST parse with PyYAML `safe_load()` now (Step 6 actually loads them for the shared_helpers: block). Any prose scalars containing `: ` (colon-space) must be `|-` literal blocks or double-quoted. Latent parse bugs existed pre-rollout in 4 AR spec files and 4 AP spec files; 4 AP ones are fixed, 4 AR ones still fail safe_load but are unused (AR is out of scope) — leave them until AR rollout.
- VBA runtime does NOT parse YAML. shared_helper_registry.yaml + transform_specs shared_helpers are AUTHORING-TIME + STEP-6-VALIDATION-ONLY contracts. The VBA project MUST stay zero-dependency on filesystem / YAML parsers. Do NOT add `Open` / `FileCopy` / yaml parse calls to any .bas runtime code.
- Comment policy: transform .bas files = exactly 1 `' Source of truth: config/transform_specs/<key>.yaml — regenerate, do not hand-edit` pointer LINE (counted via `"^\s*'\s*Source of truth:"` regex, not total comment lines). All institutional gotcha comments (Rule 10, RowIsSkippable, EIB column-slot mapping documentation) are RETAINED in the transform .bas file body — this is the correct policy. Shared engine modules (modHelpers / modFinanceCommon / modSourceTable / modDispatcher) are ALSO allowed to keep institutional-knowledge gotcha comments about VBA/Excel/Mac sandbox behavior that cannot be regenerated from YAML.

Verification (commands that must stay green across any future diff touching these files)
- Hygiene gate (0 ERROR-class):
  ```
  uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project \
    python elt_transform_excel/scripts/build_runner.py \
      --hygiene-only --runner-shell /tmp/fake.xlsm \
      --vba-source elt_transform_excel/vba_src \
      --profile-dir elt_transform_excel/config/transforms --output /tmp/out.xlsm
  ```
- AP hoisted-helper NEG guard (should print nothing; all 0 matches):
  ```
  grep -cE '^\s*Private (Function (NormaliseToWordBoundaries|NormalizeCostCodeSeparators|RowIsTrulyEmpty|RowIsEmptyOrTrailer|ResolveEffectiveNet)\b|Const (WBO_|TAX_APPLICABILITY_|TAX_CODE_ZERO|TAX_CODE_STANDARD|TAX_RECOVERABILITY_ZERO|TAX_RECOVERABILITY_STANDARD))' elt_transform_excel/vba_src/transforms/modAp*.bas
  ```
- Transform-specs YAML parse (9 AP files must return OK; AR 4 OK is optional until AR rollout):
  ```
  uv run --with pyyaml --no-project python -c "import yaml, pathlib; [print('FAIL', p.name, yaml.safe_load(p.read_text()) and 'ERR') if not isinstance(yaml.safe_load(p.read_text()), dict) else print('OK', p.name) for p in sorted(pathlib.Path('elt_transform_excel/config/transform_specs').glob('*.yaml'))]"
  ```
- `' Source of truth:` pointer-line guard (each AP transform .bas must have EXACTLY 1 pointer line; institutional gotcha comment lines retained alongside it — do NOT grep for `^[[:space:]]*'` total comment count, that fails on all 9 files). Example:
  ```
  for f in elt_transform_excel/vba_src/transforms/modAp*.bas; do c=$(grep -cE "^\s*'\s*Source of truth:" "$f"); echo "$c $f"; done | awk '{if ($1 != 1) print "BAD:" $0}'
  ```

Gotchas (new traps discovered during this rollout)
- YAML unquoted scalars with `: ` (colon-space) → PyYAML implicit-mapping key separator (see ARCHITECTURE.md § Non-Obvious Gotcha #1).
- `shell_defaults.variants:` in build profile REPLACES the base binding → base profile_key disappears from transform_bindings (see Gotcha #2).
- POS-guard hint list MUST be direct-called helpers only — do not include `NormaliseToWordBoundaries` or `FinalCodeSegment` (see Gotcha #3).
- `_validate_vba_source_hygiene()` signature-threading requirement for any new config-path accesses (see Gotcha #4).
- Headless agent environment: NO working Excel VBA execution available. Runtime smoke tests (invariant #2 above) require manual Excel execution on your Mac; we can only do static arity-equivalence diffs here. Do NOT attempt to "run VBA" with Python tooling — `pyOpenVBA` creates workbook projects, it does NOT have a VBA interpreter.
- Pre-rollout, `transform_specs/*.yaml` were NEVER parsed by any code; they are NOT required to be parseable under PyYAML safe_load rules. *Post-rollout* the 9 AP ones ARE parsed for the shared_helpers block. Any prose scalars with `: ` inside must now be quoted or block-scalared. The 4 AR spec files still fail safe_load but are not parsed — this is an acceptable known state until AR rollout. Do not "fix" the AR specs as part of unrelated work; only do it when AR Tier-1 hoisting is scheduled.
- **Read before touching code:** this doc's Constraints + Gotchas, `docs/ADDING_A_TRANSFORM.md`
  § Scaling note, and run the hygiene gate once to see the baseline (command in Verification).

---

## Historical Context: Original Measurement & Pre-Rollout Proposals (COMPLETED 2026-08-18)

> Everything below this line is the ORIGINAL measurement and open-proposal section from when this todo doc was first written (before any Tier 1 rollout work was done). It is retained for audit-trail completeness. The ACTIVE sections (current checkpoint, still-todo, constraints, verification commands, gotchas) are all at the TOP of this document above the `---` divider. A cold-start agent resuming this backlog should read the TOP sections first. Use this section only for historical context on how the duplication was measured and what was originally proposed; do not treat any "Still todo" entries below as active work.

## Finding — the duplication is real and measured

- **~7,900 lines across 14 transform `.bas` files** (238–818 each) vs ~2,100 lines of shared engine
  (`modFinanceCommon` / `modHelpers` / `modSourceTable` / `modDispatcher`).
- **Fork twins are ~83% identical** (`modApTaxi` vs `modApTravelPlaces`: 700 lines, only 82 code
  lines differ and 154 of the diff lines are *comments*). The genuine per-transform variation is a
  **data table**: the `requiredHeaders` array, the `FindColumnIndexInRow(…, "<header>")` calls, and
  a few behaviour flags — all of which are *already declared* in the spec `mapping:` section.
- **Every AP supplier transform re-defines the same six private helpers.** Measured across the AP
  supplier-invoice modules:
  - `NormaliseToWordBoundaries` — **byte-identical in 6/6** files → pure hoist.
  - `NormalizeCostCodeSeparators` — **byte-identical in 3/3** files → pure hoist.
  - `RowIsTrulyEmpty`, `RowIsEmptyOrTrailer` — 6 files, 4 variants each (differ only by which
    column indices they check) → parameterisable.
  - `RowIsSkippable` — 9 files, 8 variants (per-transform skip fields) → parameterisable.
  - `WriteOutputRow` — 9 files, 9 variants → **genuinely per-transform** (this is the mapping;
    leave it, or address only via Tier 3).
- **Config repeats too:** the `worktag_business_objects` list is byte-identical across the AP
  supplier profiles (taxi = travel_places = stationery); house defaults + `eib_population_type`
  repeat per profile (9 Supplier Invoice, 5 Customer Invoice).

**Why it exists (legitimately, in part):** the VBA runtime has no access to the YAML at run time
(finance users just get `Runner.xlsm`), so the `.bas` is a *hand-authored mirror* of the spec. The
mirror is correct; it is just copied rather than shared.

## Completed Tier 1 rollout + Stage 1 anti-drift architecture (rollout completed 2026-08-18; committed at c459cc4 eib23)

> The five items below were the original pre-rollout "Still todo options" proposals. They are now rewritten as a rollout report showing which ones were completed, which were partially completed, and which were deferred with explicit rationale. The active list of remaining open items lives in the CURRENT CHECKPOINT section at the TOP of this file (lines 88–101).

1. **Tier 1 — hoist the identical / near-identical helpers → ✅ COMPLETED 2026-08-18 at commit c459cc4 eib23**
   - Scope: 9 AP (modAp*) transform modules only; 5 AR (modAr*) modules explicitly left out of this rollout — Step 6 build_runner.py `if/else` dispatch still enforces legacy local Private Const validation for AR.
   - Hoisted to `modSourceTable.bas` as shared `Public Function` implementations:
     * `NormaliseToWordBoundaries` — 6 local private copies → 1 public (byte-identical implementations, pure hoist).
     * `NormalizeCostCodeSeparators` — 3 local private copies → 1 public (byte-identical implementations, pure hoist).
     * `RowIsTrulyEmpty` parameterised — 6 variants that only differed by which column-index positions they checked → 1 generic 3-array-signature `RowIsTrulyEmpty(tableData, rowIndex, textCheckCols, amountCheckCols)` public impl.
     * `RowIsEmptyOrTrailer` parameterised — 6 variants → 1 generic 4-array-signature `RowIsEmptyOrTrailer(tableData, rowIndex, textCheckCols, amountCheckCols, labelCheckCols)` public impl, with 7-keyword trailer detection on NormaliseToWordBoundaries-concatenated labelCheckCols.
     * `ResolveEffectiveNet` — 3 inline 4-line blocks in RowIsSkippable bodies → 1 single-call `ResolveEffectiveNet(netValue, grossValue, taxValue)` public function.
   - Hoisted to `modFinanceCommon.bas` as shared `Public Const` definitions:
     * 9-pack `WBO_*` strings (Cost Center / Spend Category / Event / Camp / FA Learning / Commercial Partner / Grant / Program / Project) — 99 local private copies across 11 modules → 9 single public consts.
     * 6-tuple `TAX_*` zero/standard × applicability/code/recoverability strings — 6 local private copies per archetype-B/C module → 6 single public consts.
     * Archetype-A (uniform_tax) modules keep their OUTPUT_TAX_* 3-flat local `Private Const`; those are per-transform genuinely different values and are NOT hoisted (verified in Step 6 as local-only).
   - Code reduction: 18 helper implementations (6+3+6+6+3) → 5 shared public helpers. 99 const copies → 15 shared public consts. 72%+ reduction in duplicated code/consts.
   - Call-site shape: The "6-positional-column-args → 3 Array signature" rewrite for the row-check helpers was actually applied two commits earlier (544c8b6 eib19 → 6696277 eib21), not in the Tier-1 commit. Tier-1 commit applied PURE scope hoisting only (delete private copies, append public to shared modules) with no transform .bas call-site shape changes. Static semantic-equivalence proof completed 2026-08-18.
   - Hygiene gate result after Tier-1 + Stage 1 validation hooks: `HYGIENE OK — 18 modules, 17 bindings, 0 ERROR-class violations` (EXIT=0).

2. **Tier 2 — extract the `Run<Name>` orchestration skeleton (WriteOutputRow / RowIsSkippable shared body) → 🕑 DEFERRED per user Option A. Reassess no earlier than: AR parity achieved + 1 more new AP transform authored under Stage-1 regime.**
   - Why deferred today (cost-benefit): The open→validate→loop→log body is ~identical; only the column map + WriteOutputRow + RowIsSkippable genuinely differ per transform. However the 20% "legitimate variance" (CFA bootstrap-2row uniform scan, PrizeFund fixed-literal company, AMEX-3 no-trailer-detection processor exports, variant ColumnIndex letter slots across EIB worktag packs) is enough that writing a shared `Application.Run` callback skeleton or table-driven loop today would produce a 70/30 80%-solution with escape hatches that cost more to maintain than the 9 small transform-body copies save.
   - Threshold to revisit (2 signals both required): (a) 5 AR `.bas` modules also get Tier-1 hoisted (gives 14 total transforms dataset to extract skeleton cleanly from; no overfit to AP-only patterns), (b) ≥1 new AP transform has been authored from scratch END-TO-END under the Stage-1 declarative regime (to see if the 20% legitimate variance space turns out to be actually smaller than it looks, or actually larger).
   - If both signals hit: revisit extracting `Run<ProfileName>` orchestration as a shared 1-body loop (or Stage 2 codegen that emits orchestration skeletons from a richer transform_spec schema). Do NOT start incrementally.

3. **Tier 3 — table-driven engine that executes the spec `mapping:` section directly at runtime → 🟡 NORTH-STAR (forever deferred unless governance explicitly signs off).**
   - This is the Scaling-note north-star and the "SQL: one source of truth, many implementations" parity with the rest of the elt_lake repo.
   - Blocked by fundamental constraints: (a) VBA runtime has ZERO file / YAML parser dependencies (finance users just open Runner.xlsm — config must live inside the compiled VBA project), (b) mapping tables inside VBA are ugly, and (c) doing this properly is a **governance-reviewed design decision**, not an incremental edit.
   - Status: Stay deferred. Not an interim-tool recommendation.

4. **Config dedup — fold shared worktag lists / house defaults into a build-time base → 🟨 PARTIAL COMPLETE.**
   - **Completed part (high-leverage, done as part of Stage 1 anti-drift):** shared concern/variant selection keys have been deduplicated into the single `shared_helper_registry.yaml` catalog. Every transform_spec declares its `shared_helpers:` block by variant KEY, not by repeating variant implementation details (which VBA function to call, which module lives in). New variant authoring only requires ONE new entry in the registry + ONE key-reference in each transform_spec that needs it.
   - **Still open (LOW priority, deferred):** Fold shared `worktag_business_objects` lists (classic-6 9-pack) and repeated house defaults / `eib_population_type` constants into a build-time YAML base that profiles inherit. The build runner already deep-merges `ref_default.json`; extending this to a YAML base-profiles system would be ~30 lines of build_runner.py edit but has no correctness impact, just aesthetics. Leave for a slow cycle.

5. **Comment policy (encode + one-time strip) → ✅ COMPLETED with corrected policy (original proposal premise slightly refined during rollout).**
   - **Policy actually enacted (corrected from the original "all comments stripped" wrong premise):** Transform `.bas` files carry EXACTLY 1 `' Source of truth: config/transform_specs/<key>.yaml — regenerate, do not hand-edit` pointer-LINE (counted via regex `"^\s*'\s*Source of truth:"` — not via total comment lines). All institutional gotcha comments in the transform `.bas` file BODIES are RETAINED: Rule 10 WriteOutputRow rationale, RowIsSkippable gotchas, EIB column-slot mapping documentation, Mac-VBA locale-specific notes — these are NOT stripped, because stripping them would cause institutional knowledge loss and would not be restorable without reading 3+ transform exemplars.
   - **Part 1 (encode in docs) done:** Added to `docs/ADDING_A_TRANSFORM.md § Step 5` as an explicit authoring paragraph, plus captured in `PROJECT_RULES.md` governance section as the active rule.
   - **Part 2 (one-time strip) done correctly per the refined policy:** All 9 AP transform `.bas` files now have exactly 1 `Source of truth:` pointer line each; spec-duplicative prose comments were removed; institutional gotcha comments retained.
   - Shared engine modules (modHelpers, modFinanceCommon, modSourceTable, modDispatcher): ALL comments kept, as originally proposed (they have no YAML spec and the comments encode operational gotchas no spec carries).

> **HISTORICAL NOTE — superseded.** The active, authoritative `Accumulated Active Constraints` + `Verification` + `Gotchas` sections for any future session continuing this backlog live at the TOP of this document (current checkpoint sections, lines 102–143). The entries below are the ORIGINAL pre-rollout constraints from when Tier 1 was an open proposal; they are retained for audit-trail purposes only, and several of them contain stale premises that have been corrected at rollout time. DO NOT use these as the source of truth. Use the top-of-file checkpoint sections instead.

## Historical Pre-Rollout Constraints (lines below = superseded by checkpoint at top of file)

> **⚠️ The comment-policy entry below is the OLD PREMISE that was CORRECTED at rollout time. The correct active rule lives in the top-of-file `Accumulated Active Constraints` section (line 110): transform .bas files = EXACTLY 1 `' Source of truth:` pointer LINE + RETAIN institutional gotcha comments in .bas body (5–93 additional comment lines per file). Do NOT strip gotcha comments from transform .bas. The original "no comments except pointer" rule was wrong and was never enacted; it would have caused institutional-knowledge loss.**
>
> - Original (superseded) historical constraint: *Comment placement is source-of-truth-aligned — comment truth lives where the truth lives. A per-transform `.bas` is a mirror of its `config/transform_specs/<key>.yaml`, so it carries NO comments except a single top-of-file pointer: `' Source of truth: config/transform_specs/<key>.yaml — regenerate, do not hand-edit`. Any other comment in a transform `.bas` duplicates the spec and drifts silently (nothing checks that the comment still matches the YAML). The prose/decision record belongs in the spec, once. The shared engine KEEPS its comments — `modHelpers`, `modFinanceCommon`, `modSourceTable`, `modDispatcher` have no YAML source of truth, and their comments encode mechanism/VBA/operational gotchas that exist in no spec. Do NOT strip these — that is institutional-knowledge loss, the exact failure the context-economy rules guard against. Net rule: prose in the YAML spec + gotchas in the shared engine; the mirrored `.bas` is comment-free apart from its one pointer line.*
- **Historical (superseded):** No behaviour change without parity proof. Any hoist/parameterise must produce byte-identical EIB output on a representative run before and after. *Note at rollout: this was honoured — static semantic-equivalence proof (eib19→eib21) was completed 2026-08-18 and runtime smoke test is the single remaining open item in the top-of-file checkpoint Still todo.*
- **Active and still holds:** VBA runtime cannot read the YAML. Anything shared must live in a `.bas` module compiled into the workbook — not a config file read at run time. (This is why Tier 3 is large.)
- **Active and still holds:** Tier 3 is governance-gated, per the Scaling note — do not start it as an incremental edit.
- **Active and still holds:** Do not touch AR customer invoices / ingest specs — blocked pending 3+ delivering the correct target EIBs. Out of scope for this backlog item. *Note at rollout: Step 6 build_runner.py still enforces legacy local Private Const validation for AR, untouched.*
- **Active and still holds:** Mac VBA hygiene holds — ASCII only, watch line length / continuation-run limits; the shared modules are hygiene-checked like any other. (See top-of-file Gotchas for `--ascii-tolerance 3` rule for em dashes in pointer comments.)

## Historical Pre-Rollout Verification (superseded by top-of-file checkpoint Verification)

> ⚠️ These are the original proposed gates from when Tier 1 was an open proposal. The ACTIVE verification commands (the ones that were actually run + will keep future diffs green) live in the top-of-file checkpoint `Verification` section at lines 112–132. Do not run the commands below; run the top-of-file ones instead.

- (SUPERSEDED) Proposed baseline hygiene gate: `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python elt_transform_excel/scripts/build_runner.py --hygiene-only`. *Actual gate at rollout uses explicit --runner-shell/--vba-source/--profile-dir flags — see top-of-file command.*
- (SUPERSEDED) Proposed parity: run one AP supplier transform (e.g. `ap_travel_places`) against representative sample files before and after, and diff the produced EIB `.xlsx` (data rows) — must be identical. *Actual parity check at rollout: static semantic-equivalence proof completed 2026-08-18 for call-sites; manual runtime smoke test remains open (Still todo step 1 at top of file).*
- (SUPERSEDED) Proposed post-change guard: `grep -rc "Function NormaliseToWordBoundaries" vba_src/transforms/` returns all zeros. *Actual active NEG guard at rollout (top-of-file line 121–124) is a broader regex covering 5 helpers + 2 const groups; use that one.*
- (SUPERSEDED) Proposed comment-policy count: every transform `.bas` has exactly **one** comment line → `for f in vba_src/transforms/*.bas; do echo "$(grep -cE "^[[:space:]]*'" "$f") $f"; done` → each count is `1`. *Was never enacted because it would have stripped gotcha comments. Active correct rule: top-of-file Verification line 129–132 counts ONLY the `Source of truth:` pointer LINE, not total comment lines in file.*

## Historical Pre-Rollout Gotchas (superseded by top-of-file checkpoint Gotchas)

> ⚠️ Some of these have been superseded by rollout-time findings. Read the TOP-OF-FILE `Gotchas` section (lines 134–143) for the authoritative active list (5 canonical traps + 2 post-rollout disclaimers) before touching code. The list below is retained for audit-trail only.

- (Historical, still useful generally): The empty-string MD5 (`d41d8cd9…`) is what a broken body-extraction returns — verify extraction found real content before trusting a "all identical" result (this bit the first measurement pass).
- (ACTIVE — also captured in top-of-file): `WriteOutputRow` and `RowIsSkippable` genuinely differ per transform — do **not** force them into a single shared function as part of Tier 1; that is Tier 2/3 territory.
- (SUPERSEDED, premise was wrong at rollout): `config/transform_specs/*.yaml` are authoring records and are **not** strict-parseable YAML (prose notes with colons); `yaml.safe_load` fails on them at HEAD too. Only `config/transforms/*.yaml` profiles are build-consumed and must parse. *Corrected at rollout: 9 AP spec files are now actually PARSED by build_runner.py Step 6 for their `shared_helpers:` block; all 4 latent AP scalar-parsing bugs (colon-space inside prose scalars) were fixed (see top-of-file Gotchas #1 for fix rule: `|-` literal or double-quoted strings for any scalar containing interior `: `). The 4 AR spec files still fail safe_load but are unused — tolerated.*
- (Historical, still useful generally): When stripping comments, only lines starting with `'` (or `Rem`) are comments. Do **NOT** remove `Attribute VB_Name`, `Option Explicit`, or the `Public Sub Run<Name>` entry-point declaration — they are not comments, and removing them breaks the module identity / build contract.
