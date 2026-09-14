# Unify invoice-line grouping (single-line = multi-line degenerate case)

> Opened 2026-09-14 from a design review. Every AP/AR invoice transform should treat the
> source as **multi-line**: group rows by the source invoice number, increment the
> Spreadsheet Key when the invoice number changes, and run a 1..N Line Order / Row ID
> counter within each group. Single-line is just the N=1 case. The remaining holdouts
> were the AP `ap_stationery` / `ap_taxi` bespoke write paths plus the AR
> `ar_invoice` / `ar_invoice_run` / `ar_bulk_invoice` per-row grouping fields.

## Resume (start here)

- **Status (2026-09-14):** The **entire AP side is now unified** on one shared handler
  (`modInvoiceGrouping.AssignInvoiceLineGroup`) — the original 8 migrated transforms plus
  `ap_taxi` and `ap_stationery`. The first AR tranche is also migrated:
  `ar_invoice`, `ar_invoice_run`, and `ar_bulk_invoice` now derive their Spreadsheet Key and
  line-order fields from the same helper. The shared-helper contract is now declared in
  `config/shared_helper_registry.yaml` as `invoice_line_grouping.per_invoice_number`, and
  Step 6 of `build_runner.py` enforces the helper call/spec declaration in both directions.
  **All VBA changes remain pending manual Excel verify**. Still open: unresolved-key AR
  transforms (`ar_gbe`, `ar_discipline`, `ar_multi_line`), AR integration uniformity, and
  AMEX GF (Decision 2).
- **Read first:** `docs/ARCHITECTURE.md` (Authoring model + the AR/AP EIB column invariants),
  `config/shared_helper_registry.yaml` (no `invoice_line_grouping` concern exists yet),
  `config/transform_specs/ap_stationery.yaml` Rules 15/16/18/19 and `ap_taxi.yaml` Rules 15/16/18/19
  (the proven grouping pattern to promote), and the sibling refactor
  `docs/todo/collapse-transform-shells-engine.md` (this is the first concrete slice of it).

## Outcome review — what the shipped past runs actually contain (2026-09-14)

Audited all 23 output EIBs in the AP/AR target folders with the new
`scripts/audit_outputs.py` (read-only). Findings:

- **Grouping is done in only 2 of the classic transforms.** `ap_taxi` (16 invoices / 8
  multi-line groups, EK=EM=1..N) and `ap_stationery` **Ambassador Taxi** run are correct.
- **Live duplicate-ref bugs in the target folders right now:**
  - `ap_wnl_grants` — 46 invoices, **15 genuine multi-line groups mis-keyed** (same
    Invoice Number on 2 rows -> two different Spreadsheet Key B values, e.g.
    `WNLGRNTS DEC 25-KEY002 -> B in {1,39}`). EK=1, EM blank. This is the bug, in
    production-shaped output, on a transform still using the single-line assumption.
  - `ap_stationery` **Addison Lee / DHL / E-Courier** outputs are **stale pre-fix
    vintages**: single invoice split across 13 / 21 / 4 rows with B=1..N (dup-ref),
    EK=1/EM blank, old tax codes, and GE!=GF. Only the Ambassador run was re-run after
    the multi-line + tax + GE/GF fixes. **If AP loaded Addison/DHL/E-Courier, Workday
    would reject them.** These need a re-run once the code is unified (or immediately).
- **Control Amount Total (col AS): blank in every AP output.** The user's rule ("never
  populate it") is already honoured in data. The stale `control_amount_eib_as` doc line
  in `ap_commercial_invoice.yaml` was removed (code never wrote it).
- **Tax value split is latent, not yet emitted wrongly.** Only `ap_taxi` + Ambassador
  emit the AP-Supplier standard-rated triple
  (`GBR_Domestic_Purchase_Of_Goods_And_Services / GBR_Standard / Fully_Recoverable`).
  Every other classic transform's shipped rows are all zero/outside-scope
  (`GBR_Outside_of_Scope / GBR_Zero_Rated / Nonrecoverable_Allocate`), so the shared
  `modFinanceCommon` standard-rated constant (`GBR_Standard_Rated / GBR_Standard_Rated /
  Recoverable_Allocate`) has **not** actually been written yet — but it WILL be the moment
  any of those transforms hits a VAT>0 row. See Decision 1.
- **AMEX outputs use a non-classic column layout** (Invoice Number col K blank; tax/GE/GF
  blank at classic letters). `audit_outputs.py` flags them as not-classic and skips the
  letter-based checks. Their real column map needs confirming before auditing/refactoring.
- **AR:** `ar_multi_line` collapses a whole file into ONE invoice (B=1, DV=1..23); every
  other AR transform keys B per-row. Same single-vs-multi split as AP, but AR has a
  dedicated multi-line MODULE rather than a per-row toggle.

## Done (2026-09-14; all VBA pending manual Excel verify per the contract)

- `scripts/audit_outputs.py` — read-only outcome-audit of the shipped target-folder EIBs;
  flags dup-ref, stale vintage, populated Control Amount Total, GE!=GF, and non-classic
  layout. Run: `uv run --with openpyxl --no-project python scripts/audit_outputs.py`.
- `vba_src/modInvoiceGrouping.bas` — NEW shared module. `AssignInvoiceLineGroup(keyMap,
  lineOrderMap, invoiceNumber, nextKey, lineOrder)`: assigns Spreadsheet Key per NEW invoice
  number and a 1..N Line Order / Row ID within the group. Plain-VBA Collections only. Self-
  contained logic that now sits on top of the shared keyed-Collection primitives hoisted into
  `modHelpers`.
- **All 8 AP invoice transforms migrated onto `AssignInvoiceLineGroup`** (module-level group
  state reset per run; keyed on the `invoiceNumberValue` each already computes): `modApWnlGrnInv`,
  `modApCommInvSupInv`, `modApCfaDistribSupplierInvoice`, `modApCfaClaimsSupplierInvoice`,
  `modApPrizeFundSupplierInvoice`, `modApRefereeSupplierInvoice`, `modApRefereeV2SupplierInvoice`,
  `modApTravelPlacesSupInv`. Each now writes B (per-invoice-group), EK and EM (= 1..N within the
  group). WNL was the actively-buggy one (15 mis-keyed groups); the rest are genuinely single-line
  today (unique K per row) so their output is byte-identical, but they now use the SAME mechanism —
  the single/multi-line distinction is gone, decided only by whether the source Invoice Number
  repeats. Specs updated (B/EK/EM) except referees/referees_v2 which don't enumerate those columns.
- `modApTaxiSupplierInvoice.bas` + `modApStationerySupplierInvoice.bas` — write-time grouping now
  also routes through `AssignInvoiceLineGroup`. Their Phase-1 prepass is retained only for the
  4-field header-consistency guard; the bespoke Spreadsheet Key / Line Order bookkeeping was removed.
- `modArInvoiceCustomerInvoice.bas`, `modArInvoiceRunCustInvoice.bas`, and
  `modArBulkInvoiceCustInvoice.bas` — Spreadsheet Key plus line-order fields now route through
  `AssignInvoiceLineGroup`, keyed by the source invoice number column each profile already reads.
- `config/shared_helper_registry.yaml` + `scripts/build_runner.py` — NEW shared-helper concern
  `invoice_line_grouping.per_invoice_number`; Step-6 POS/NEG/reverse guards now enforce
  `AssignInvoiceLineGroup` declarations and usage. Relevant AP/AR transform specs declare it.
- AP standard-rated tax -> Domestic PURCHASE triple for `ap_cfa_distributions`, `ap_amex_trains`,
  `ap_amex_trains_v2` (build-profile `tax_*` blocks + local Private Const shadows). AR untouched.
- Removed stale `control_amount_eib_as` line from `config/transform_specs/ap_commercial_invoice.yaml`.
- **Consolidated the keyed-Collection primitives** (2026-09-14). The 5 helpers
  `CollectionKeyExists / *SetItemLong / *GetItemLong / *SetItemDouble / *GetItemDouble` were
  copy-pasted (Private) in FOUR modules (`ap_taxi`, `ap_stationery`, `ap_cfa_claims`,
  `ar_integration`) plus a near-twin set in `modInvoiceGrouping`. Hoisted ONE Public copy into the
  domain-free `modHelpers`; deleted all four per-module copies; `AssignInvoiceLineGroup` now calls
  the shared ones too. ~250 lines of duplication removed; the call sites are unchanged so behaviour
  is identical. Hygiene OK.
- `build_runner.py --hygiene-only`: HYGIENE OK, 25 modules, 20 bindings, 0 ERROR-class violations.

## Decisions

1. **Standard-rated AP-Supplier tax triple — RESOLVED 2026-09-14: use Domestic PURCHASE for all AP.**
   `GBR_Domestic_Purchase_Of_Goods_And_Services / GBR_Standard / Fully_Recoverable`. Applied to the
   three AP transforms that were still emitting the shared AR value (`ap_cfa_distributions`,
   `ap_amex_trains`, `ap_amex_trains_v2`) via their build-profile `tax_*` blocks in
   `config/transforms/` + a local Private Const shadow in each `.bas` (the pattern taxi/stationery/
   travel already used). The shared `modFinanceCommon` `TAX_*_STANDARD` constants were LEFT as
   `GBR_Standard_Rated` because AR customer-invoice transforms (bulk, GBE, invoice_run) correctly
   use them (Sale, not Purchase). **Follow-up (not urgent):** teach the build_runner tax gate about
   an AP-specific shared standard constant so the per-module local shadows can be centralised.
2. **AMEX GF — STILL OPEN.** The decision "align GF=GE" was made on the assumption AMEX uses the
   classic GE/GF pair. It does NOT: `modApAmexTrains*` write the line amount to `GE` **and `IK`**
   (a wider EIB column region) and never touch `GF`. Writing `GF=amount` would put a value in a
   column AMEX's layout doesn't use for Extended Amount. HELD pending confirmation of the AMEX EIB
   column map (is `IK` the AMEX Extended Amount, and should `GF` stay blank?). No AMEX amount change
   was made this pass.

## Still todo (next actions, harness-guarded)

1. **Manual Excel verify of the 8 AP migrations** — open `dist/Runner.xlsm`, run each migrated AP
   transform, and run `scripts/audit_outputs.py`. For the genuinely-single-line ones
   (cfa_distributions, cfa_claims, prize_funds, referees, referees_v2, commercial, travel_places)
   the output is identical to before **except `EM` (Line Order) is now `1` where it was previously
   blank** — the correct Line Order for a single line and consistent with the multi-line transforms,
   but confirm Workday accepts it (taxi/stationery already ship EM=1..N, so it should). `ap_wnl_grants`
   must now show its multi-line groups sharing one B with EK=EM=1..N (DUP-REF + STALE flags cleared).
   Non-negotiable gate before trusting the helper.
2. **AR side — finish the unresolved-key transforms only.**
   - `ar_integration`: optional uniformity migration only; its synthetic per-row ref already means
     grouping remains single-line in practice.
   - **NEEDS the grouping key confirmed (do NOT guess):** `ar_gbe` (customer resolved by club NAME,
     no invoice-number column), `ar_discipline` (invoice-number source unclear), and `ar_multi_line`
     (currently treats the WHOLE FILE as one invoice — is its group key a real per-invoice column, or
     is one-file-one-invoice the intended rule?). Ask Finance / the user before wiring these.
4. **AMEX GF (Decision 2)** — confirm the AMEX EIB column map (`IK` vs `GF`), then write or leave blank.
5. **Re-run the stale outputs** the audit flagged (stationery Addison/DHL/E-Courier; AMEX) — they
   are pre-fix vintages carrying the dup-ref bug and must NOT be loaded to Workday as-is.
6. Track whether `ap_referees` / `ap_referees_v2` should also enumerate B/EK/EM explicitly in spec
   prose now that the shared-helper contract is in place; runtime behaviour is already unified.

## Constraints (inherited from collapse-transform-shells-engine.md)

- **Real-Excel verify is the non-negotiable gate.** VBA behaviour cannot be proven in this
  repo's tooling; `build_runner.py --hygiene-only` gates source, not behaviour. Every
  migrated transform is "pending manual Excel import + verify" until confirmed in Excel.
- **One transform's output at risk at a time.** No big-bang. Use `audit_outputs.py` (and
  the collapse TODO's planned byte-diff harness) as the before/after gate on each step.
- **Write fresh from spec — never fork+sed an exemplar `.bas`** (ARCHITECTURE §5).
- **Shared-helper anti-drift + CRLF/ASCII + column-allowlist gates** all still apply.

## Verification

- Outcome audit: `uv run --with openpyxl --no-project python scripts/audit_outputs.py`
  (exit non-zero while anomalies remain; target = only the AMEX not-classic flags remain).
- Source hygiene after any `.bas`/registry change:
  `uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python scripts/build_runner.py --hygiene-only`
  -> 0 ERROR-class violations, no SPEC<->CODE DRIFT.

---
_Resume prompt: `From docs/todo/unify-invoice-line-grouping.md: run the manual Excel verify gate, then resolve the remaining AR grouping-key decisions and AMEX GF.`_
