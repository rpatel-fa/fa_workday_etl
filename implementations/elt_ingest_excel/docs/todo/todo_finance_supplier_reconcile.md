# Finance Supplier Reconciliation — TODO

## Resume (start here)

- From `elt_ingest_excel/docs/todo/todo_finance_supplier_reconcile.md`: Continue
  `## Active -> Finance supplier reconciliation continuity` for any future stakeholder
  follow-up. The `Reopen — Workday EIB load failures (2026-07-24)` engineering item is CLOSED.

Supplier transform and publishing were locked (archived in
`elt_ingest_excel/docs/todo/archive/todo_finance_supplier_transform_locked_2026-07-24.md`), then
DELIBERATELY REOPENED on 2026-07-24 to fix a banking defect surfaced by the Workday EIB load.
See `## Reopen — Workday EIB load failures (2026-07-24)` below for the current state: the
banking transform fix is VERIFIED in DuckDB, and the finance EIB correction path now covers the
7 region failures, the 16 City Subdivision tags, and the residual real-IBAN account slices
(CR/GE/UA/KZ) plus 1 sibling-BIC recovery. Engineering is COMPLETE. Remaining unresolved rows are
now data gaps / operator action only, not open transform or automation defects.

## Reopen — Workday EIB load failures (2026-07-24) [CLOSED]

Continuity: transform DELIBERATELY REOPENED on explicit approval to fix the banking defect
below. Banking fix applied to `config/transform/sql/finance/supplier/src_fin_supplier_raw.sql`
and VERIFIED in DuckDB via a Phase-1 TRANSFORM rerun (not re-loaded to Workday - the finance
EIB was patched directly instead, see Phase 2). The supplier SQL / named-range path is still
UNTOUCHED for regions; instead, the finance EIB correction script now resolves the 7 region
failures in a country-scoped way from reference data. Two-phase decoupled model: Phase 1 = SQL
fix + rerun for lineage; Phase 2 = patch the finance EIB directly, values sourced from DuckDB on
the stable reference_id (never supplier_id, which is run-scoped). Read this section before
touching the transform again.

Evidence source: `~/Downloads/Master Suppliers V1 24.07.2026 with errors.xlsx` (Workday EIB
result: 7,751 txns, 139 failed). Failing keys join to the workbook `Spreadsheet Key` (col 2);
stable business key is `Supplier Reference ID` (col 9) = GP Vendor ID.

Root-cause split (categories overlap; 139 distinct failing transactions):
- ~113 banking — `nrm_bank_account_number` was sourced only from GP `EFT Bank Account`, but GP
  stored the account in its `IBAN` column for non-SEPA/local-clearing countries. Confirmed
  against GP source: US `2TH001` IBAN=`145572120403` (raw US account); GB `AIR003`
  IBAN=`GB84BARC20176865382055` (sort `201768` + account `65382055` embedded); SG `AIR010`,
  PH `ALI024` raw local account in IBAN column. TRANSFORM — data present/derivable.
- 6 region — descriptor emitted by `workday_supplier_address.sql` (`t.region`) is resolved to a
  `Country_Region_ID` by a downstream named-range resolver that is NOT country-scoped:
  LU cities → `BEL-WLX` (Belgium), NL Maastricht → `BEL-VLI`, NL Zierikzee → `DNK-85`. TRANSFORM.
- 1 region — `MAT097` (Genova/IT): resolver emitted its not-found sentinel
  `NOT ON NAMED RANGE TAB!` (string exists in no repo file — it is the resolver's, not ours).
- 16 address — coded City Subdivision (`Country_Subregion_Code_In_Country`, col 122) required by
  Workday for FI/SK/ZA and never populated. SOURCE gap / needs new reference derivation.
- ~7 source-data — genuine GP junk: `QVE001` sort-code field holds a BIC + 10-digit account;
  `PAS003` account `000`/code `0`; BIC junk `0000`/`609561`/`930067`/`CLNOUS66BR`;
  `OAN001` US routing `26002561` = 8 digits (dropped leading zero). FINANCE / source fix.

Done
- Applied fill-blank-only banking recovery in `src_fin_supplier_raw.sql`: recover
  `nrm_bank_account_number` (GB IBAN chars 15-22; raw IBAN-column value for non-IBAN
  countries), recover GB sort code into `nrm_bank_code_routing_number` (IBAN chars 9-14),
  and null `nrm_bank_iban` when the value is not a structural ISO-13616 IBAN so a misfiled
  raw account does not also ship as a bogus IBAN. Extraction expressions validated in
  isolated DuckDB against the real failing IBAN values.
- PHASE 1 (lineage) DONE + VERIFIED: reran `fin_supplier_creditor.py --run-to-phase TRANSFORM`
  (DB backed up first to scratchpad; env gate `ALLOW_FIN_SUPPLIER_CREDITOR_RUN=1`). Did NOT
  publish. Post-fix `workday_supplier_settlement_account`: 7400/7641 rows now carry
  `bank_account_number`; the 241 still blank are all real SEPA/IBAN countries (DE/FR/ES/NL/IE
  ...) that Workday accepts on IBAN alone and that never failed. No US/GB/SG/PH/JP/AU remain
  blank.
- PHASE 2 (EIB delivery) DONE: `scripts/populate_eib_corrections.py` clones the finance EIB and
  fills the yellow CORRECTION column (C) on the error tab. Surgical: only the error-tab sheet
  XML is rewritten inside a byte-for-byte zip clone, so the Submit Supplier sheet + all EIB
  data validations are preserved. Correction VALUES come from DuckDB keyed on reference_id
  (authoritative, leading zeros preserved); the EIB's own block only disambiguates which value.
  Result after the banking-only pass: 217 error lines -> 125 AUTO, 1 SUGGEST, 91 MANUAL.
- PHASE 3 (verify) DONE: 109/109 account + 14/14 sort AUTO corrections lineage-linked to DuckDB
  by reference_id. Verification also exposed that the EIB's numeric account cells are lossy
  (dropped leading zeros; 3 accounts >15 digits already corrupted in the finance file) - which
  is why the authoritative value must come from DuckDB, not re-derived from the EIB.
- PHASE 4 (region correction in EIB) DONE: updated `scripts/populate_eib_corrections.py` to
  convert workbook ISO2 country codes via `config/data/ref_country_iso3_mapping.json`, then do a
  country-scoped `Country_Region_ID` lookup against
  `config/data/ref_workday_country_state_region.csv`, preferring the workbook's human
  `region_desc` before municipality/subregion fallbacks. Verified on the live workbook: all 7
  region errors now resolve automatically (`ADA051`, `BON003`, `GAM002-FA`, `MAI022`, `MAT097`,
  `TAL006`, `YOU016` -> `LUX-L`, `NLD-LI`, `ITA-GE`, `NLD-ZE` as appropriate).
- PHASE 5 (deliver regenerated EIB) DONE: wrote
  `~/Downloads/Master Suppliers V1 24.07.2026 with errors CORRECTED region.xlsx` plus audit CSV
  `~/Downloads/Master Suppliers V1 24.07.2026 with errors CORRECTED region audit.csv`. New
  result: 217 error lines -> 148 AUTO, 0 SUGGEST, 69 MANUAL.
- PHASE 6 (extra urgent EIB recovery) DONE: extended `scripts/populate_eib_corrections.py`
  again so the error-tab correction pass now also:
  (a) extracts separate account numbers from real IBANs for Costa Rica / Georgia / Ukraine /
  Kazakhstan (`FED027`, `GEO060`, `KOB003`, `MYS001`, `NIK009`);
  (b) reuses a valid sibling BIC inside the same EIB key for `ENA002|ENA003` -> `AIBKIE2D`;
  and (c) emits a strong BIC suggestion for `LUC020` -> `CHASUS33` from a unique published
  bank-name/country match. Latest urgent handoff files:
  `~/Downloads/Master Suppliers V1 24.07.2026 with errors CORRECTED urgent.xlsx` and
  `~/Downloads/Master Suppliers V1 24.07.2026 with errors CORRECTED urgent audit.csv`.
  New result: 217 error lines -> 155 AUTO, 2 SUGGEST, 60 MANUAL.
- PHASE 7 (deterministic GP follow-up pack) DONE: added
  `scripts/generate_unresolved_gp_fixes.py` so the unresolved GP remediation workbook can be
  regenerated from a cold-start session using the urgent audit CSV plus DuckDB context.
  Verified output:
  `~/Downloads/Master Suppliers V1 24.07.2026 unresolved GP fixes.xlsx`
  with 62 rows (`60 MANUAL`, `2 SUGGEST`) across `gp_fix_rows` and `summary` tabs.

Still todo (next actions)
1. CLOSED - all deterministic EIB-side fixes have been applied and regenerated into
   `~/Downloads/Master Suppliers V1 24.07.2026 with errors CORRECTED urgent.xlsx`.
2. CLOSED - strict output now writes only confirmed values into error-tab column C; unknown rows
   and previously suggested rows are left blank.
3. OPERATOR / BUSINESS FOLLOW-UP ONLY - the remaining unresolved rows are genuine data gaps:
   52 routing/institution numbers, 3 unresolved BICs, 2 branch IDs, 1 address gap (`HAR050`),
   1 junk bank code (`PAS003`), and 1 junk sort-code field (`QVE001`).
4. OPERATOR / BUSINESS JUDGEMENT ONLY - 2 additional `LUC020` BIC rows were left blank in the
   strict workbook because they were plausible but not confirmed strongly enough to auto-apply.
5. OPTIONAL FUTURE HARDENING - if desired later, the downstream region/named-range resolver can be
   hardened in-product, but it is no longer required for this EIB remediation item.

Accumulated Active Constraints
- SENSOR GAP: internal `validation_supplier_issues` under-reported these — it flagged 67
  banking issues (Workday rejected ~113) and did NOT flag the cross-country region cases. Do
  NOT treat internal-validation silence as proof of a clean load; the Workday EIB result is the
  authoritative evaluator. Add these Workday-confirmed cases as offline regression checks.
- Banking fix is fill-blank-only (COALESCE onto existing logic): it cannot overwrite a real
  GP-supplied account, so it cannot regress currently-passing records. The SQL-side fix is
  lineage-verified; Workday verification still depends on Finance re-loading the patched EIB.
- The Country_Region_ID is produced by a downstream named-range resolver, not by our SQL. A
  region descriptor we emit can be silently mis-resolved across countries; region correctness
  cannot be judged from our SQL output alone. The immediate workaround is now in the EIB patcher,
  not the transform.
- If the delta path (`config/transform/sql/finance/delta_supplier/src_fin_supplier_raw.sql`) is
  a separate copy, it needs the same banking fix — verify before any delta run.
- `scripts/populate_eib_corrections.py` must translate workbook ISO2 country codes to country
  names before region lookup; `ref_workday_country_state_region.csv` is keyed by country name,
  not ISO2.
- The remaining 60 MANUAL lines are now concentrated in genuinely missing values:
  52 routing/institution numbers, 3 unresolved BICs, 2 branch IDs, 1 address gap, 1 junk bank
  code, 1 junk sort code. We have exhausted the deterministic values present in DuckDB/source/EIB.
- The strict deliverable is business-safe: only confirmed values are populated in error-tab
  column C; there are no notes, suggestions, or manual instructions written into the workbook.
- Created a separate operator workbook for the unresolved GP-side follow-up:
  `~/Downloads/Master Suppliers V1 24.07.2026 unresolved GP fixes.xlsx`
  with one row per unresolved error line (62 total), carrying GP Vendor ID, Workday Supplier ID,
  error-tab row, field to correct, and the GP action needed.
- The unresolved GP workbook is now reproducible, not manual: use
  `uv run --project elt_ingest_excel python elt_ingest_excel/scripts/generate_unresolved_gp_fixes.py`
  from the workspace root.

Verification
- `uv run --project elt_ingest_excel python scripts/populate_eib_corrections.py "~/Downloads/Master Suppliers V1 24.07.2026 with errors.xlsx" --audit "/tmp/supplier_eib_audit_after_region_fix.csv" --out "/tmp/Master Suppliers corrected scratch v2.xlsx"`  -- expect `AUTO=148`, `MANUAL=69`, no `SUGGEST`
- `python3 -c "import csv; rows=list(csv.DictReader(open('/tmp/supplier_eib_audit_after_region_fix.csv', newline='', encoding='utf-8'))); region=[(r['reference_id'], r['status'], r['correction']) for r in rows if 'Region Name must be valid' in r['error'] or 'NOT ON NAMED RANGE TAB' in r['error']]; print(region)"`  -- expect `[('ADA051', 'AUTO', 'LUX-L'), ('BON003', 'AUTO', 'LUX-L'), ('GAM002-FA', 'AUTO', 'NLD-LI'), ('MAI022', 'AUTO', 'LUX-L'), ('MAT097', 'AUTO', 'ITA-GE'), ('TAL006', 'AUTO', 'LUX-L'), ('YOU016', 'AUTO', 'NLD-ZE')]`
- `uv run --project elt_ingest_excel python scripts/populate_eib_corrections.py "~/Downloads/Master Suppliers V1 24.07.2026 with errors.xlsx" --audit "/tmp/supplier_eib_audit_after_extra_fixes.csv" --out "/tmp/Master Suppliers corrected scratch v3.xlsx"`  -- expect `AUTO=155`, `SUGGEST=2`, `MANUAL=60`
- `python3 -c "import csv; rows=list(csv.DictReader(open('/tmp/supplier_eib_audit_after_extra_fixes.csv', newline='', encoding='utf-8'))); focus=[(r['reference_id'], r['status'], r['correction']) for r in rows if r['reference_id'] in {'FED027','GEO060','KOB003','MYS001','NIK009','ENA002|ENA003','LUC020'}]; print(focus)"`  -- expect the 5 IBAN-account recoveries, `AIBKIE2D` for `ENA002|ENA003`, and `CHASUS33` suggestions for `LUC020`
- `uv run --project elt_ingest_excel python elt_ingest_excel/scripts/generate_unresolved_gp_fixes.py`  -- expect `Unresolved rows: 62 | MANUAL=60 | SUGGEST=2` and workbook `~/Downloads/Master Suppliers V1 24.07.2026 unresolved GP fixes.xlsx`
- `uv run --project elt_ingest_excel python -m py_compile elt_ingest_excel/scripts/generate_unresolved_gp_fixes.py`
- `SELECT COUNT(*) FROM workday_supplier_settlement_account WHERE NULLIF(TRIM(bank_account_number),'') IS NULL AND NULLIF(TRIM(bank_iban),'') IS NOT NULL;`  -- expect ~0 for GB/non-IBAN
- Re-load `~/Downloads/Master Suppliers V1 24.07.2026 with errors CORRECTED urgent.xlsx` to
  Workday and confirm the banking + region + City Subdivision keys clear, leaving only the
  60 genuine MANUAL/source-data cases plus the 2 intentionally blank non-confirmed `LUC020`
  BIC rows.

Gotchas
- Failing "row" numbers in the errors sheet are the workbook `Spreadsheet Key` (col 2), not
  file rows; one transaction spans multiple sheet rows (address/bank sub-blocks).
- GP source header is offset: col 0 = Business Unit, col 1 = Vendor ID.
- Older `CORRECTED.xlsx` copies in `~/Downloads` predate the country-scoped region fix; use the
  `CORRECTED urgent.xlsx` output for the latest handoff.
- The final strict rerun used a freshly re-downloaded Gayle workbook as input, after clearing
  prior generated copies from `~/Downloads`.
- `elt_ingest_excel` is not currently a root uv workspace member; run these supplier scripts with
  `uv run --project elt_ingest_excel ...` from the repo root, not `uv run --package ...`.

## Active

### Finance supplier GP→Workday CCS reconcile pack (2026-07-29) [COMPLETE — signed off 2026-07-29]

Status: COMPLETE. Supplier reconcile confirmed clean (population 0/0) and the Finance worklist
(`finance_action_list`, 110 extract-verified items) delivered in
`~/Downloads/supplier_ccs_reconcile_review_pack.xlsx`. Remaining work is Finance-side data-gap
cleanup in Workday, not ETL. Do not reopen without fresh evidence of a transform defect.


Built the supplier analogue of the customer CCS reconcile pack. New assets:
`config/ingest/finance/workday_supplier_extract.json`,
`config/reconcile/finance/supplier_ccs/*.sql`,
`config/publish/finance/publish_supplier_ccs_reconcile_review.json`,
`examples/fin_supplier_ccs_reconcile_sql.py`. Deliverable:
`~/Downloads/supplier_ccs_reconcile_review_pack.xlsx` (14 sheets — first tab is the
8-section executive `summary`, built by `examples/supplier_ccs_summary_sheet.py` and prepended on
every export, matching the customer pack's summary format).

Inputs: Workday extract `~/Downloads/Extract Suppliers 2026-07-28 20_40 BST.xlsx` (refreshes
`fin_workday_extract_supplier`, 7767 rows) and the go-live manual-update list
`~/Downloads/Suppliers to be manually updated at Go Live.xlsx` (107 names → `fin_supplier_manual_update`).

Finance work from ONE tab: `finance_action_list` (`reconcile_supplier_ccs_finance_action_list`) —
extract-verified open items only (110 = 106 manual-updates still open + 3 residual address + 1 test
row). The `manual_update` tab carries a `still_open_in_workday` column. Verified vs Gayle's file:
all 107 names match (16 addr / 57 bank / 34 pay); the extract independently confirms 106/107
removals. The 1 exception is HARVEY WEBSDALE (HAR050) — listed for bank removal but its bank loaded
fine and is present, so it is routed to ADDRESS (residual) instead, not reworked as bank.

Result (verified live in DuckDB):
- Population CLEAN: 0 GP suppliers missing in Workday, 0 extra (8118 vendor ids both sides).
- 1 Workday dummy row: `test` / `S`-id `100001`, blank reference_id (tenant test record).
- 107 manual-update suppliers classified (16 address / 57 bank / 34 payment), all matched to the
  extract by normalised name (0 unmatched).
- Attribute gaps: 3 UNEXPLAINED address gaps only (`FCL001`, `ALE055`, `HAR050` — the known
  residuals in this file), 0 unexplained bank, 0 unexplained payment. Everything else explained by
  the manual-update list.
- 7 merge-group diffs (CHR/KIN/TGI) surfaced as INFO — transform merged differently to Workday.

Accumulated Active Constraints (supplier_ccs)
- Reconcile key is `nrm_supplier_number` (BU-suffixed, e.g. `LON004-FA`), NOT
  `src_fin_supplier_base.vendor_id` (raw `LON004`, reused across FA/NFC/WNSL). Building the source
  population from `vendor_id` falsely reports hundreds of missing/extra/dummy rows.
- "Address present" in the extract must check ALL address columns (`primary_address`,
  `primary_remit_to_address`, …); CH/IT suppliers store the address as a Remit-To address, so
  `primary_address` alone over-reports address gaps (was 44 false, 3 true).
- Manual list is name-only; matching is on normalised name against the extract. If a future list
  fails to match a name, it lands on the `manual_unmatched` sheet — investigate rather than ignore.

Verification (supplier_ccs)
- `uv run python examples/fin_supplier_ccs_reconcile_sql.py --run-to-phase TRANSFORM --export-review-workbook --export-review-dir ~/Downloads`
- `SELECT section, metric, metric_value, status FROM reconcile_supplier_ccs_report ORDER BY 1,2;`
  -- expect 2a/2b=0 PASS, 2c=1 REVIEW, 3e=0 PASS, 4b=3 ATTENTION, 4d/4f=0 PASS.

### Finance supplier reconciliation

Backlog Item: Finance supplier reconciliation continuity

Done
- Locked the supplier transform workstream and moved its implementation history into the archive.
- Preserved the production handoff evidence needed for ongoing reconciliation in this file.
- Kept the supplier TODO path active for reconciliation-only follow-up.
- Checked the live published supplier data for duplicate `supplier_name + bank account` cases and found 1 live collision, already surfaced by `validation_supplier_issues`.
- Checked the live published supplier data for duplicate `supplier_name + primary address` cases and found 13 duplicate-address cases, all separated by different bank-account sets except the same live `PINSENT MASONS` case.
- Confirmed the strict combined setup check (`supplier_name + primary address + bank details`) has 1 live exception only: `PINSENT MASONS` (`PIN004` / `PIN005`).

Still todo (next actions)
1. Review the live duplicate `supplier_name + bank account` case for `PINSENT MASONS` (`PIN004` / `PIN005`, supplier_ids `S-005536` / `S-005537`, account `30719986`, sort code `200771`) and decide whether it is a genuine duplicate requiring reopen/escalation.
2. Use the evidence below to handle ongoing reconciliation and stakeholder challenge questions.
3. If a new discrepancy is raised, classify it first against the known residual exceptions before treating it as a new issue.
4. Use stable business keys for every follow-up review: GP Vendor ID / `reference_id`, not `supplier_id`.
5. If someone proposes a new supplier transform fix, treat that as a reopen decision and do not change the transform from this backlog without explicit approval.

Accumulated Active Constraints
- Supplier transform and publishing are locked; this backlog is reconciliation-only.
- `supplier_id` is run-scoped. For cross-run comparisons or joins to external review files, use GP Vendor ID / `reference_id`, not the `S-` id.
- Remaining banking exceptions are mostly non-derivable source-data gaps; do not treat them as open ETL defects without fresh evidence.
- Workday normalisation and validation outcomes must be interpreted from the live validation SQL plus the evidence captured in this file, not from older intermediate issue packs.
- If reconciliation work eventually proves a real transform defect, reopen deliberately rather than editing under the assumption that the lock no longer applies.

Verification
- `uv run python examples/fin_supplier_creditor_validate_sql.py`
- `SELECT rule_type, col, COUNT(*) FROM validation_supplier_issues GROUP BY 1,2 ORDER BY 3 DESC;`
- `SELECT supplier_name, COALESCE(NULLIF(TRIM(bank_account_number), ''), NULLIF(TRIM(bank_iban), '')) bank_account_key, COUNT(DISTINCT supplier_id) supplier_count FROM workday_supplier_settlement_account WHERE COALESCE(NULLIF(TRIM(bank_account_number), ''), NULLIF(TRIM(bank_iban), '')) IS NOT NULL GROUP BY 1,2 HAVING COUNT(DISTINCT supplier_id) > 1;`
- `SELECT supplier_name, country, region, city, address_line_1, address_line_2, postal_code, COUNT(DISTINCT supplier_id) supplier_count FROM workday_supplier_address_base WHERE is_primary = 'Yes' GROUP BY 1,2,3,4,5,6,7 HAVING COUNT(DISTINCT supplier_id) > 1;`
- `SELECT supplier_name, country, region, city, address_line_1, address_line_2, postal_code, bank_account_number, bank_code_routing_number, COUNT(DISTINCT supplier_id) supplier_count FROM workday_supplier_address_base a JOIN workday_supplier_settlement_account s USING (supplier_id) GROUP BY 1,2,3,4,5,6,7,8,9 HAVING COUNT(DISTINCT supplier_id) > 1;`

Gotchas
- DuckDB is often locked by DBeaver; close it before rerunning validation queries.
- Historic Workday Supplier IDs from older finance workbooks are not stable join keys.
- The current supplier workbook is the production baseline; do not casually regenerate it during reconciliation follow-up.

## Reconciliation Evidence

### Validation snapshot at handoff

- 121 total `validation_supplier_issues` rows
- 67 settlement-account / banking issues
- 52 address issues across 26 suppliers
- 2 supplier-name duplicate issues outside the finance action workbook review

Interpretation:

- the outstanding banking issues are mostly missing / invalid routing numbers, bank codes, IBANs,
  or malformed account details
- the outstanding address issues are mostly postal-code quality gaps, plus 2 region-required cases
  and 1 country-reference gap
- these remaining items were known at handoff time and were not considered unresolved ETL defects

### Review basis

The production workbook was reviewed against the prior finance issue workbook using stable business
keys:

- `GP Vendor ID(s)` from the finance workbook
- `reference_id` in `workday_supplier_name`
- supplier name / business unit where needed

Historic `Workday Supplier ID` values from the finance workbook were not used as the primary join
key because they are run-scoped and change between runs.

Reference issue workbook reviewed:

- `~/Downloads/Master Suppliers/Finance Action List - Supplier Fixes.xlsx`

### Direct comparison result

1. Bank Format
- `1 Bank Format`: 33/33 rows matched to the current output; 0 still show current settlement-account validation issues.
- `1b Bank Format - Resolved`: 22/23 rows matched; 0 still show current settlement-account validation issues.
- `1c Bank Format - Uncertain`: 5/5 rows matched; 0 current settlement-account validation issues.
- Interpretation: bank-detail issues that were feasible to resolve in transform logic have been resolved; the 5 rows in `1c Bank Format - Uncertain` remain best-effort UK reclassification cases, not authoritative bank-name confirmations.

2. Phone Issues
- `3 Phone Issues`: 5/5 rows matched; 0 still show current phone validation issues.
- Interpretation: the prior phone issues that were feasible to resolve in the transform are resolved in the current output.

3. Region Invalid
- `4 Region Invalid`: 2/2 rows matched; 0 still show current address-domain validation issues.
- Interpretation: the earlier Luxembourg region concern is not currently surfacing as an open SQL validation issue in the current output.

4. Address Missing
- `2 Address Missing`: 67/67 rows matched.
- 2 of those rows still have current address validation issues:
- `HAR050` — `HARVEY WEBSDALE`: UK postal code still missing
- `JAM055` — `JAMES DURKIN`: UK postal code `DTF 04H` still invalid
- Interpretation: most prior address issues were resolved; not all prior address issues were eliminated.

5. Multi-Bank Payment Type
- secondary accounts are flagged `For Supplier Connections Only = 'Y'`
- payment-type restriction fields remain blank
- Interpretation: this remains a business / Workday-process decision, not something derivable from GP data.

### Banking position

- 0 settlement-account rows have a blank `bank_name`
- 0 current settlement-account validation issues are on `bank_name`
- 67 current settlement-account validation issues remain across 63 suppliers
- remaining banking validation issues are on missing / invalid routing numbers, bank codes, IBANs,
  or malformed account details, not on an unresolved bank-name transform gap

Important known limitation:

- UK bank names for ambiguous sort-code prefixes such as `04` and `23` are best effort, not
  authoritative, because no authoritative free 6-digit sort-code-to-bank dataset was available to
  this repo at handoff time.

### Current residual address exceptions

Region required:

- `ALE055` — `ALESSANDRO GIALLATINI` — region required for Italy
- `FCL001` — `F CLOSS SERVICOS EM TECNOLOGIA DA INFORMACAO EIRELI` — region required for Brazil

Country-reference gap:

- `PAS003` — `PASVORD PRODAKSHN DOO` — no address requirements defined for North Macedonia in the current reference data

Postal code exceptions:

- `ANT029` — `ANTANAS PAULAUSKAS` — `1109` (Lithuania)
- `BES004` — `BESPARKLED` — `6901` (Netherlands)
- `COM012` — `COMPANY FOR ELECTRONIC INDUSTRIAL AUTOMATIONLTD` — `B93 46F` (United Kingdom)
- `CRO005` — `CROFTON INTERIORS` — `BTA 8LA` (United Kingdom)
- `DAR035` — `DARREN WILLIAMS` — `NE7 76Q` (United Kingdom)
- `EMM009` — `EMMA VASE` — `MK43 75R` (United Kingdom)
- `FRE005-FA|FRE006` — `FRENFORD FC` — `IG 3PS` (United Kingdom)
- `FRO004` — `FRONTIERS LADIES FC` — `ESSEX` (United Kingdom)
- `HAR050` — `HARVEY WEBSDALE` — missing postal code (United Kingdom)
- `INT015` — `INTERNATIONAL FEDERATION OF CP FOOTBALL` — `6821` (Netherlands)
- `JAM055` — `JAMES DURKIN` — `DTF 04H` (United Kingdom)
- `JAM126` — `JAMAICA FOOTBALL FEDERATION` — `00000` (Jamaica)
- `JAM135` — `JAMIE LAW` — `PR4 52J` (United Kingdom)
- `JAS002-WNSL` — `JASON MARTIN` — `IP21` (United Kingdom)
- `LAT003` — `LATVIAN FOOTBALL FEDERATION` — `LV 1013` (Latvia)
- `MAL015` — `MALTA FOOTBALL ASSOCIATION` — `ATD 4000` (Malta)
- `MIC059` — `MICROSOFT IRELAND OPERATIONS` — `D18` (Ireland)
- `MIL023|MIL027` — `MILTON KEYNES IRISH FC` — `BERKS & BU` (United Kingdom)
- `MOR015` — `MORGAN EIGHT` — `HR9 6N7` (United Kingdom)
- `NOG002` — `NOGOMETNI FUDBALSKI SAVEZ BIH` — `71` (Bosnia and Herzegovina)
- `PBP001` — `CORSEARCH BV EUROS ONLY` — `1043` (Netherlands)
- `PHI028` — `PHILIP PROSSER` — `WSS 3LW` (United Kingdom)
- `SPA013` — `SPARKLING ORANGE NL` — `1015` (Netherlands)

### Recommended wording when challenged later

Use this position:

> The supplier workbook published on 2026-07-24 was reviewed against the prior finance issue pack and against the live SQL validation outputs. All issues that were feasible to correct in the ETL layer were corrected before handoff. Remaining exceptions were known at the time of handoff and reflect source-data limitations, best-effort derivation limits, reference-data gaps, or business / Workday-side decisions rather than unresolved transform defects.

For the production decision specifically:

> Finance confirmed there would be no further GP updates, and the workbook was therefore issued as the final best-available production output rather than held for further source correction.

For banking specifically:

> Supplier bank details were taken as far as possible using final GP source data, current routing / sort-code mappings, GP fallback logic, and IBAN-derived enrichment. Remaining banking exceptions were known source-data limitations rather than outstanding ETL defects.

### Live duplicate `supplier_name + bank account` finding

Current live check result:

- 1 `supplier_name + bank account` collision exists in the published supplier data
- the same 2 rows are already flagged in `validation_supplier_issues` as duplicate supplier-name cases sharing a bank account

Current known case:

- `PINSENT MASONS`
- `reference_id` values: `PIN004`, `PIN005`
- `supplier_id` values: `S-005536`, `S-005537`
- `bank_account_number`: `30719986`
- `bank_code_routing_number`: `200771`
- `bank_name`: `Barclays Bank UK plc`
- `bank_country`: `United Kingdom`

Interpretation:

- we cannot confirm that no such duplicates exist; the live data shows 1 current case
- source lineage review shows this shared `supplier_name + bank account` combination already exists in `src_fin_supplier_base` / `src_fin_supplier_raw` for `PIN004` and the `FA` row of `PIN005`
- `src_fin_supplier` then preserves that Barclays account on `PIN005` alongside its separate RBS client account, and `workday_supplier_settlement_account` publishes both
- this is reconciliation work first; only reopen transform if review proves it is a true unmerged duplicate rather than an accepted business case

### Current duplicate setup position

Current live published-output position:

- `supplier_name + bank details`: 1 live duplicate case
- `supplier_name + primary address`: 13 live duplicate-address cases
- `supplier_name + primary address + bank details`: 1 live duplicate case

Interpretation:

- the 13 duplicate-address cases were intentionally left as separate suppliers because their bank-account sets differ
- the only remaining case that fails the stricter combined supplier-setup check is `PINSENT MASONS`
- on the current published output, the supplier setup is therefore clean on the combined address-and-bank test except for `PINSENT MASONS`

Duplicate-address cases with different bank-account sets:

- `AFTA STUDIOS` — `AFT002`, `AFT003`
- `ALEX MATTHEWS` — `ALE064`, `ALE065`
- `BARNSLEY FC` — `BAR006-NFC`, `BAR020`
- `CHADDERTON FC` — `CHA001`, `CHA008-NFC`
- `CUSTOMER FIRST VENUE SERVICES` — `CUS001`, `CUS003`
- `DAVID LAMBERT T/A LAUNDRY REPUBLIC` — `DAV009`, `DAV158`
- `JEREMY LAMBERT` — `JER010`, `JER011`
- `MATCHFIT SPORTS PR` — `MAT004`, `MAT009`
- `OFFICE CARE` — `OFF003`, `OFF007`
- `RECREATIONAL SPORTING` — `REC001|REC002|REC004`, `REC008`
- `RHYS JEFFERY` — `RHY005`, `RHY007`
- `TOMMY BREEN` — `TOM015`, `TOM016`
