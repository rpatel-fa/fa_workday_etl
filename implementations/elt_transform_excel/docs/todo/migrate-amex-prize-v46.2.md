# Migrate AMEX ×3 + Prize Funds to v46.2 EIB columns

> **Correctness gap (latent).** Opened 2026-08-25. These 4 AP supplier-invoice transforms still
> write the OLD v46.1 line/split columns while their profiles target `Submit_Supplier_Invoice_v46.2.xlsx`,
> so a run mis-maps Extended Amount / Memo into the *Supplier Invoice Split Line* group — the exact
> defect fixed for the other 7 transforms this session. Benign until these feeds are run; **critical
> the moment finance runs an AMEX or Prize Funds batch.** The spec↔code gate does NOT currently cover
> them (their specs expose no parseable `mapping:` targets), so nothing catches it at build.

## Resume (start here)

- **Status:** PARTIALLY DONE. `ap_prize_funds` complete (2026-08-25); `ap_amex_hotel`, `ap_amex_trains`, `ap_amex_trains_v2` NOT STARTED.
- **Priority:** ABOVE the engine-collapse item ([collapse-transform-shells-engine.md](collapse-transform-shells-engine.md)) — this is correctness, that is velocity. **Do before these feeds next run.**
- **Read before touching code (all on disk — no git archaeology needed):**
  1. `docs/ADDING_A_TRANSFORM.md` → the **LANDMINE** box + the column table (`GF` = line Extended
     Amount *not* split `IK`; `GQ` = line Memo *not* `IL`) — this is the WHY.
  2. An **already-migrated module as the living reference**: `vba_src/transforms/modApTaxiSupplierInvoice.bas`
     — read its `OUTPUT COLUMN MAPPING` doc block (marked `Submit_Supplier_Invoice_v46.2`) and its
     `ColumnIndex()` writes; copy that exact pattern. For `ap_prize_funds` the closest shape is
     `modApTravelPlacesSupInv.bas` (also has an `AR→AG` receipt-date write). For AMEX the closest
     shape is `modApStationerySupplierInvoice.bas` (single fixed coding, per-row supplier split
     absent — company constant).
  3. `config/eib_population_defaults.yaml` → v46.2 column conventions + `safe_shared_columns`.
- **Scope remaining:** `ap_amex_hotel`, `ap_amex_trains`, `ap_amex_trains_v2`.
- **Scope done:** `ap_prize_funds` (all 6 steps below completed 2026-08-25).

## Done

- **ap_prize_funds** — fully migrated 2026-08-25.
  1. Re-audit confirmed offenders: `IK`, `IL`, `AR`, `DD` (no `BB`/PO write).
  2. Applied 4 column swaps in `modApPrizeFundSupplierInvoice.bas`:
     `IK→GF` (Extended Amount), `IL→GQ` (Memo), `AR→AG` (Invoice Received Date), `DD→EK` (line Row ID).
  3. Added 29-entry v46.2 `mapping:` block to `config/transform_specs/ap_prize_funds.yaml` —
     spec↔code gate now covers this module (was previously a 0-target no-op).
  4. Rewrote `OUTPUT COLUMN MAPPING` doc block to v46.2 with correct column letters for every field.
  5. Build + bidirectional audit: 0 `IK/IL/AR/DD` writes, `GF/GQ/AG/EK` all present in
     both code writes and spec mapping (29/29 1:1 correspondence, 0 orphans either side).
     `build_runner.py` → "Prep complete." 0 `SPEC<->CODE DRIFT`.
  6. Real-run reconciliation: pending next finance prize-funds batch; spec↔code gate verified syntactically.

## Still todo (next actions)

1. **Re-audit each AMEX module's exact `ColumnIndex()` writes** (do not assume — confirm per module).
   Current known offenders: amex×3 write `IK`, `IL`, `DD`. Confirm none write `BB` (PO) or `AR` before assuming.
2. **Apply the v46.2 column swaps** in each AMEX `.bas` (same pattern as Prize Funds):
   `IK→GF` (Extended Amount), `IL→GQ` (Memo), `DD→EK` (line Row ID). No `AR→AG` unless re-audit
   finds a receipt-date write. No `BB→BH` unless the re-audit finds a PO write.
   Re-normalise CRLF after any programmatic write.
3. **Bring each AMEX transform_spec `mapping:` to v46.2** so the spec↔code gate covers them.
   NOTE: AMEX specs express mapping via `derived_fields` / a different form — the `mapping:` target
   letters may be absent or v46.1; add/translate them so `_parse_spec_mapping_targets` sees v46.2
   targets and the gate protects these modules going forward (parity with Prize Funds / 7 migrated).
4. **Update each AMEX module's `OUTPUT COLUMN MAPPING` doc block to v46.2 (as done for Prize Funds).
5. **Build + audit**: 0 `IK/IL/AR/DD` writes remain in amex×3; `GF/GQ/AG/EK` present per module; build green.
6. **Confirm with a real run** of at least one AMEX feed — reconcile output like
   the 9 shipped transforms (row counts, v46.2 columns populated, split-line columns empty).

## Accumulated Active Constraints (active for all remaining items)

- **v46.2 line/split contract (the whole point):** line Extended Amount = `GF`, line Memo = `GQ`;
  `IK`/`IL` are the *Split Line* twins and MUST NOT be written by a v46.2 transform. Received Date =
  `AG` (not `AR`=Hold Reason); line Row ID = `EK` (not `DD`=Manual Override Percent). See the landmine
  note in `docs/ADDING_A_TRANSFORM.md`.
- **Spec↔code gate must end up covering these 4** — after step 3, injecting a wrong target must fail
  the build for them too (parity with the migrated 7).
- **CRLF-only for `.bas`**; **≥300-char lines are a hard build error**.
- **AMEX final-code derivation is a different concern** (`amex_padded_22char_pipeline` /
  `direct_final_code_column`) — do NOT touch it; this item is ONLY the line/split column fix.

## Verification

- Per-module audit (expect old=0, v46.2 present):
  `for m in modApAmexHotelSupplierInvoice modApAmexTrainsSupplierInvoice modApAmexTrainV2SupplierInvoice modApPrizeFundSupplierInvoice; do echo $m; grep -oE 'ColumnIndex\("(IK|IL|AR|DD|GF|GQ|AG|EK)"\)' vba_src/transforms/$m.bas | sort | uniq -c; done`
- Build (all gates incl. spec↔code): `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/build_runner.py` → "Prep complete", no `SPEC<->CODE DRIFT`.
- Input contract: `validate_input.py -t ap_amex_hotel` (etc.) → GREEN.

## Gotchas

- **AMEX feeds target the v46.2 template already** (`Submit_Supplier_Invoice_v46.2.xlsx`) — so the bug is
  live the instant they run; there is no "still on v46.1 template" grace. (Prize Funds same — fixed 2026-08-25.)
- **The spec↔code gate was a no-op for all 4** (0 parseable targets) — that is exactly why the
  drift went unnoticed. Prize Funds `mapping:` was added 2026-08-25; AMEX 3 remain to be closed.
- **Prize Funds wrote `AR`** (receipt date → Hold Reason) in addition to IK/IL/DD; AMEX×3 do not
  (known — re-audit to confirm). Re-audit rather than assume symmetry across AMEX modules.

---
_Resume prompt: `From docs/todo/migrate-amex-prize-v46.2.md: Continue → Still todo step 1 for AMEX (re-audit amex×3 ColumnIndex writes)`_
