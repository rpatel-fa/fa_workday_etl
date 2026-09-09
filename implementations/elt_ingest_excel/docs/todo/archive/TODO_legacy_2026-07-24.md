# elt_ingest_excel — TODO

## Resume (start here)

**LIVE ITEM (2026-07-24): UK bank-NAME correctness — blocked on authoritative data.** Latest
finance source re-run to TRANSFORM → **7,751 suppliers, 151 validation issues** (banking issues
72→67). The open decision is bank names: neither GP nor the 2-digit `ref_bank_sort_code_prefix_mapping`
is authoritative (prefix wrong for sub-allocated ranges e.g. Starling on 60-83→"NatWest"; GP wrong for
others e.g. `600846` is NatWest not GP's "Revolut"). Only a valid **6-digit Pay.UK ISCD (sort→bank)**
resolves it (also unlocks UK modulus checking). Do NOT change GP-vs-prefix precedence without it.
See the **"Supplier bank-name derivation & correctness"** item in Active. Memory:
[[project_workday_supplier_bank_name]], [[project_workday_supplier_dedup]], [[feedback_verify_dont_assume]].
Verify state first (DB often locked by DBeaver — ask user to disconnect):
`SELECT rule_type, col, COUNT(*) FROM validation_supplier_issues GROUP BY 1,2 ORDER BY 3 DESC;`

**DONE (2026-07-24): IBAN-embedded bank CODE fix.** Finance confirmed they can make **no further GP
corrections**, so the only lever for the remaining banking gaps is the transform. Fixed the 5 records
where GP supplied no domestic bank code but the supplier's own IBAN embeds it (ISO 13616): GE×2, KZ,
UA, CR. New `ref_iban_structure.json` (per-country bank-code offset/length, SWIFT IBAN Registry) +
`iban_bankcode` LATERAL in `src_fin_supplier_raw.sql`, `COALESCE`d as a **fill-blank-only** fallback
into `nrm_bank_code_routing_number` (never overwrites GP; `iban_length` guard refuses truncated/garbage
IBANs — Georgia's `'0'` record S-004252 correctly stays flagged). Verified: 72→67 banking, grand total
156→151, zero regressions, every code cross-checked vs its BIC. The remaining 67 (US routing ×27,
AU/NZ BSB, Japan, etc.) are **non-IBAN countries with only a BIC** — genuinely underivable, and finance
won't correct them, so those specific accounts will fail Workday validation as a data reality.

**DONE (2026-07-24): Finland/Slovakia city_subdivision — city IS the subdivision.** Earlier design
sourced `city_subdivision` from `Address Line 3` for FI/SK; but address_3 is empty for all 15 affected
records and finance is frozen, and Workday marks `municipality_city = 'Not Accepted'` + `city_subdivision
= 'Required'` for these countries — so the post town (Helsinki etc.) is Workday's locality field. Without
the fix `city` AND `city_subdivision` both came out NULL (locality dropped entirely). Fixed in
`workday_supplier_address_base.sql`: when city_subdivision required AND municipality not accepted,
`COALESCE(nrm_city_subdivision, nrm_city)` — fill-blank-only, table-driven, no GP data needed. Verified:
city_subdivision issues 30→0, grand total 151→**121**, South Africa unaffected (keeps city + real suburb),
zero city/subdivision duplication. Supersedes the old "GP Address Line 3 data ask" for FI/SK.

---

*Prior triage (context — customer EIB export, superseded by the live item above):*
The original three backlog items and every category in the real
customer Workday EIB failure export (438 rows) have now been triaged — see Completed section for
the full list. Summary of that EIB failure export triage:

1. **"Manchester" → Jamaica Country_Region ID** (140 rows) — closed, not a bug in this repo; 3+
   already knows and manually corrects it on the EIB.
2. **Postal-code format** (237 rows) — mostly already fixed by unrelated prior work; of the ~30
   genuinely live issues, fixed 12 (leading-zero padding, UK O/0 typo) in the shared
   `udf_normalize_postal_code`; ~18 customer / ~29 supplier remain as genuine bad source data (data
   ask, not a code fix).
3. **Address Line 3** (24 rows), **Tax `Country_Reference`** (12 rows), **Municipality** (11 rows)
   — all already resolved by prior (untracked) work; zero live issues today, confirmed against
   current `validation_customer_issues`.
4. **`subregion` / "Region Subdivision 1"** (part of the 11-row "missing address components"
   category) — a genuinely missing field in both domains, fixed this session (see "Subregion
   county field" item below). The other rows in that category (Finland/Slovakia/South Africa
   `city_subdivision`) were already handled by the earlier `city_subdivision` item; the remaining
   USA/Monaco "missing City" rows are isolated genuine source-data gaps (a handful of records with
   blank `city` in GP), not a systemic bug — not raised with finance yet.

5. **Supplier settlement account `BANK_FORMAT`/`BANK_ACCOUNT`** (111 issues, 107 distinct
   suppliers) — triaged separately (not part of the customer EIB export; discovered when asked
   whether all transform work was done). All genuine GP data gaps — see "Supplier bank/settlement
   account data quality" item below. No code fix possible or needed.

Outstanding, non-code, to raise with finance together: (a) the one-time GP `Address Line 3` data
ask for the 2 live Finland/Slovakia customer records (`city_subdivision` item), (b) the ~18
customer / ~29 supplier genuinely bad/incomplete postal codes, (c) the handful of USA/Monaco
records with a blank `city` field, (d) 107 suppliers with missing/wrong bank identifiers (routing
number, bank code, BSB, IBAN, IFSC) — see item below for the full list.

**ABANDONED, do not re-attempt without new information**: running 3+'s own workbook
validation/EIB-creation macros ourselves (for internal confidence-checking before handoff) — see
"Running 3+'s workbook macros directly — ABANDONED" item below for why. Transform correctness is
already independently verified via the real customer EIB failure export triage and the full
`validation_*_issues` sweep above; this was always a secondary cross-check, not the primary
verification, and it hit a genuine environment blocker (not a bug in this repo).

## Session start prompt

```
load the use-context skill, and from: elt_ingest_excel/docs/todo/TODO.md, continue
```

## Active

### Supplier bank-name derivation & correctness — LIVE (blocked on ISCD)

**Done (committed):** US routing → `ref_bank_routing_number_mapping` (tenant AT/AU list, 19,221 rows)
built + wired, consumed first in `nrm_bank_name`. Precedence in `src_fin_supplier_raw.sql`:
(1) US routing; (2) UK prefix name if not a `%various%`/`%range owner%` placeholder; (3) GP name
(real, not 'BANK NAME'); (4) IBAN name; (5) last-resort prefix "Various" so a valid sort+account
still ships. Recovered 193 GP names on prefixes 04/23; "Various" 441→248.

**Still todo:**
1. **UK bank-name correctness — need a 6-digit Pay.UK ISCD (sort→bank).** Neither GP nor the 2-digit
   prefix is authoritative (see Resume). Load ISCD into the 32-row stub `ref_bank_sort_code_mapping`
   and make it the PRIMARY UK bank-name source; prefix/GP become fallback. Decide source (official
   licensed vs compiled — flag provenance). Do NOT change GP-vs-prefix precedence without it.
2. **Normalisation map is BUILT but NOT wired:** `config/data/ref_bank_name_normalisation.json`
   (52 rules, keyword→canonical, 90% coverage). Consistency ONLY (`MONZO`/`Monzo Bank Ltd`→`Monzo
   Bank`), not correctness. Wire (add ref SQL + `LEFT JOIN LATERAL`) only if interim consistency is
   wanted; hold if ISCD is imminent.
3. **Regenerate** `~/Downloads/Master Suppliers/Supplier Bank Corrections for GP.xlsx` against the new
   data; confirm which of the prior 189 finance items are now resolved.
4. Optional: UK modulus check (needs ISCD weight tables) + IBAN mod-97 checksum (self-contained).

**Constraints:** `supplier_id` is run-scoped — join finance data on GP Vendor ID / `reference_id`;
bank-number pads are UK-only + exactly-one-short; bank name is free-text for UK (no tenant list),
tenant-listed for US. `validation_supplier_issues` is a partial re-implementation of Workday's
checks — don't equate a low count with "all correct".

### Apply vendor-ID business-unit-collision fix to customer/debtor pipeline — RESOLVED, no port needed

GP `vendor_id` was found to be scoped per `business_unit` (company book), not globally
unique, on the supplier side — the same code gets independently reused for unrelated
vendors in different business units (e.g. `HAL003` = 3 different real companies across
FA/NFC/WNSL). The old canonicalization silently merged all rows sharing a `vendor_id`
into one Workday supplier regardless of name, corrupting ~452 of 7,178 supplier records
(commingled addresses/banks/tax data, 2 of 3 real identities disappearing under one
name). This was the dominant root cause of Workday EIB validation failures on the
supplier load (77% of 564 failed records, via a "must have only 1 active settlement
bank account" error caused by unrelated companies' bank accounts landing on one merged
record).

The fix has been implemented and verified on the **supplier** side only. This item is to
verify whether the same problem exists on the **customer/debtor** side and, if so, apply
the equivalent fix. **Do not assume it applies — verify first** (Still Todo step 1).

#### Done (supplier side, prior session)

- `config/transform/sql/finance/supplier/src_fin_supplier_raw.sql` — added a
  `cte_supplier_id` CTE immediately after the fully-normalised name is computed
  (`cte_supplier_nrm`). Window function
  `COUNT(DISTINCT nrm_supplier_name) OVER (PARTITION BY nrm_vendor_id) > 1` detects the
  collision; when true, every row sharing that vendor_id (not just the differing one)
  gets `nrm_supplier_number := nrm_vendor_id || '-' || business_unit`; otherwise
  unsuffixed. Downstream Level-2 name-based dedup (`src_fin_supplier_rnk.sql`,
  pre-existing, unmodified) automatically re-merges same-named split parts back
  together across business units with no code change needed there — verified on a real
  case: `HAP001` (FA/NFC both "Happy Days Coaches") re-merges into one supplier, WNSL
  ("Happy Ltd") stays separate.
  Verification: `SELECT supplier_id, nrm_supplier_name, nrm_agg_supplier_number FROM
  src_fin_supplier WHERE nrm_agg_supplier_number LIKE '%HAP001%'` → 2 rows, split
  correctly.
- Same file — added a whitespace-collapse step
  (`REGEXP_REPLACE(..., '\s+', ' ', 'g')`) to the name-normalisation regex chain, which
  was missing and caused false-positive collisions from double-spaced source names.
- `config/validate/finance/supplier/validate_supplier_id_coverage.sql` — COVERAGE
  checks: join now tolerates an optional `-{business_unit}` suffix
  (`t.vendor_id = s.vendor_id OR t.vendor_id LIKE s.vendor_id || '-%'`). DUPLICATE
  checks: rewritten to key on `nrm_supplier_number` itself (post-split), not the raw
  `vendor_id` — one vendor_id resolving to multiple supplier_ids is now expected by
  design for collision cases, so a raw-vendor_id 1:1 check would false-positive on
  every expected split.
- `config/reconcile/finance/supplier/reconcile_supplier_gp_id_proof.sql` — same
  tolerant-join fix as above.
- New file `config/reconcile/finance/supplier/reconcile_supplier_collision_review.sql`
  — ranks every collision split by `jaro_winkler_similarity` between the split names,
  flagging `> 0.85` as `REVIEW - likely same vendor` (possible false positive — same
  vendor entered with a typo/abbreviation/spacing difference across business units, not
  a real collision) vs `OK - confidently distinct`.
- `src/elt_ingest_excel/transform/udf/address.py` —
  `udf_normalize_postal_code`, the "postal code is made up of two parts" (UK) branch:
  was returning raw alphanumeric text as a fake postcode when no genuine UK postcode
  pattern matched (e.g. address-line text like `6THIEPVALDRIVE` masquerading as a
  postcode) instead of `None`. Fixed to return `None`. **This UDF is shared with the
  customer domain already** (referenced in
  `config/validate/finance/customer/validate_customer_address.sql` and
  `config/transform/sql/finance/customer/src_fin_customer_raw.sql`) — this specific fix
  is already inherited by customer, no separate action needed there.
- `config/transform/sql/finance/supplier/workday_supplier_settlement_account.sql` —
  added `for_supplier_connections_only = 'Y'` for every bank account beyond the first
  (ranked) per supplier, for suppliers with genuinely multiple real bank accounts (not
  a collision artifact). **Does not apply to customer** — customer's settlement/bank
  table is permanently disabled (`WHERE 1=2`), no customer equivalent exists.
- `docs/finance-supplier.md`, `docs/reconciliation.md` — updated to document the
  supplier-side collision detection, the collision-review table, and the settlement
  disambiguation.

#### Verified — no port needed (this session)

Ran the exact collision-detection methodology against `src_fin_customer_raw` (6,056 rows,
3 business units: FA/NFC/WNSL):

```sql
SELECT TRIM(nrm_customer_number) customer_id, COUNT(*) cnt_rows,
       COUNT(DISTINCT business_unit) bus, COUNT(DISTINCT nrm_customer_name) cnt_names
FROM src_fin_customer_raw
GROUP BY TRIM(nrm_customer_number)
HAVING COUNT(DISTINCT nrm_customer_name) > 1;
-- 0 rows
```

Also checked the weaker precondition directly — does any `customer_number` even appear in
more than one business unit, regardless of name:

```sql
SELECT TRIM(nrm_customer_number) customer_id, COUNT(DISTINCT business_unit) bus
FROM src_fin_customer_raw GROUP BY TRIM(nrm_customer_number) HAVING COUNT(DISTINCT business_unit) > 1;
-- 0 rows
```

No `customer_number` is reused across business units at all in this dataset (unlike
supplier's `HAL003`, reused across FA/NFC/WNSL for 3 different real companies) — the
structural precondition for the bug isn't present, so there is nothing to fix. (A separate,
unrelated finding: 10 `customer_number`s have exact duplicate rows — same business unit,
same name, e.g. `C00037_D` — plain source duplicates, not a collision; out of scope for
this item.)

**Conclusion: do not port.** No code change needed in `src_fin_customer_raw.sql`,
`validate_customer_id_coverage.sql`, or a customer collision-review file — the customer
domain has no cross-business-unit ID reuse for the fix to address.

#### Accumulated Active Constraints (historical — item resolved; retained only in case collision-style work resurfaces on another domain)

- The domain's `*_id` (e.g. `S-XXXXXX` supplier_id, and whatever the customer
  equivalent is) is run-scoped and renumbers on every pipeline run
  (`ROW_NUMBER() OVER (ORDER BY name, ...)`) — never use it to compare across runs or
  against an externally-provided reference file. Always join on the stable business key
  (`nrm_agg_supplier_number` / `reference_id`, or the customer equivalent).
- That stable business key can itself be a pipe-delimited list — pre-existing Level-2
  name merges predate this fix and can combine multiple raw source IDs under one
  canonical record. Split on `|` on BOTH sides of any comparison; never assume either
  side is a single atomic value.
- `jaro_winkler_similarity` is a useful triage heuristic, not authoritative: it
  correctly flags true typos as "likely same vendor" (good), but can also misfire
  "likely same vendor" on two genuinely different entities that merely share a long
  common prefix (e.g. two different universities both starting "University of...").
  Treat any review table built on it as a priority-sort for human attention, never as a
  verdict.
- A business-unit-collision split can get silently re-combined with an unrelated,
  unsuffixed source ID via the pre-existing Level-2 name merge (e.g.
  `'UNI012|UNI021-NFC'`). Any code that parses the suffix back out of
  `nrm_agg_supplier_number` (or its customer equivalent) must UNNEST the pipe-delimited
  parts first, then test/extract per part — never regex against the whole joined
  string (see Gotchas).

#### Verification (superseded — see "Verified — no port needed" above for the query actually run and its result)

#### Gotchas

- Don't regex-extract a vendor/customer ID from the whole `nrm_agg_supplier_number`
  string when building the collision-review table — a real bug this way silently
  dropped a genuine collision (`UNI021`) because it got pipe-joined with an unrelated
  ID (`UNI012`) via the pre-existing Level-2 merge, and a whole-string regex grabbed
  `'UNI012|UNI021'` as one garbled value instead of recognising `UNI021` inside it.
  Unnest the pipe list first, then test each part independently
  (`part LIKE '%-%'` → `SPLIT_PART(part, '-', 1)`).
- `vendor_short_name` (a GP field truncated to 15 characters) is a weaker collision
  signal than the fully-normalised name — two genuinely different long names can share
  the first 15 characters (e.g. two different universities). Don't use it as the
  primary detection signal; it was cross-checked against the actual detection this
  session and diverged on ~2% of cases, all resolved by checking the real
  (untruncated, fully-normalised) name instead.
- `uv run` fails with a `setuptools`/`ai-agent-core` build error if run from the wrong
  working directory — always `cd` into `elt_ingest_excel/` first.
- DuckDB file lock: if another process (e.g. DBeaver) has the database open, queries
  fail with `IOException: Could not set lock`. Ask the user to disconnect it rather than
  working around it.

### Add missing `city_subdivision` address field for customer — DONE, pending finance data ask

Workday requires a `City Subdivision` value for supplier/customer addresses in 5 countries —
South Africa, Finland, Slovakia, Mexico, Åland Islands — per
`ref_workday_validate_supplier_address` (supplier) / `ref_workday_validate_customer_address`
(customer, confirmed to have its own equivalent `city_subdivision` column with the same 5
countries flagged `Required`). This field did not exist anywhere in the schema before the
supplier-side fix — a genuinely missing Workday field, not a data-flow bug.

**This is a live, confirmed-needed port, not a dormant one**: `src_fin_customer_base` currently
has 4 records in the 3 relevant countries (no Mexico/Åland Islands customers) —
South Africa ×2 (`CEN0225111936` / `CEN0422108650`, both "Gregory Bortz" — likely the same person
under two customer numbers), Finland ×1 (`L00807`, Football Association of Finland), Slovakia ×1
(`CSO001`, Karol Csonto). `workday_customer_address_base.sql` currently has **zero**
`city_subdivision` handling (confirmed via grep — no match anywhere in
`config/transform/sql/finance/customer/`).

#### Done (supplier side)

- `config/transform/sql/finance/supplier/src_fin_supplier_raw.sql` — added `nrm_city_subdivision`
  computation: `WHEN r.country_code = 'ZA' THEN t.addr_unique_list[2]` (South Africa already has
  real suburb data in address_2, e.g. "Allen's Nek", a Johannesburg suburb — kept as-is);
  `WHEN r.country_code IN ('FI','SK','MX','AX') THEN` cleaned raw `address_3`. Restricted to
  exactly these 5 country codes via explicit `WHEN` branches (no `ELSE`) so no other country's
  `address_3` is ever touched, even as an unused candidate.
- `config/transform/sql/finance/supplier/src_fin_supplier_id_rnk.sql` and
  `src_fin_supplier_rnk.sql` — added `nrm_city_subdivision := fn_initcap(t.nrm_city_subdivision)`
  to the address `STRUCT_PACK` in both files, propagating the field through both aggregation
  levels.
- `config/transform/sql/finance/supplier/workday_supplier_address_base.sql` — added the output
  column, gated by the Workday requirement flag itself (not a hardcoded country list): `CASE WHEN
  LOWER(TRIM(req.city_subdivision)) = 'required' THEN u.address.nrm_city_subdivision ELSE NULL
  END`. Table-driven so it automatically covers any future country Workday adds this requirement
  for.
- Verified end-to-end: South Africa's `city_subdivision` output ("Allen's Nek") correct;
  Finland/Slovakia still `NULL` pending finance populating GP `Address Line 3` — this is a
  one-time GP data ask (GP freezes after migration — this whole project is a one-time
  migration, not an ongoing integration), communicated to finance separately (not tracked
  in this repo).
- `docs/finance-supplier.md` updated with the full architecture (ZA/FI-SK-MX-AX split, one-time
  migration framing).

**Known, deliberately-deferred side finding (not fixed)**: `ref_workday_validate_supplier_address`
lists North Macedonia as `"Macedonia, the Former Yugoslav Republic of"` (old UN name), but our own
`ref_country.json` correctly uses `"North Macedonia"` — the join in
`workday_supplier_address_base.sql` is by exact name text, so it fails silently for this one
country (1 supplier, S-005220, gets a false "no address requirements defined" warning). Not
finance's problem — a reference-data/join-key issue on our side. Low priority (1 record); a
proper fix would join by ISO code via the currently-unused `ref_country_iso3_mapping` table
instead of by name text. Check whether `ref_workday_validate_customer_address` has the same stale
name for Macedonia if/when this item is worked.

#### Done (customer side, this session)

- Verified `ref_workday_validate_customer_address` matches supplier's column name and country
  list exactly — same `city_subdivision` column, same 5 countries (`ZA`/`FI`/`SK`/`MX`/`AX`)
  flagged `Required`.
- Verified `src_fin_customer_base` has the same raw address column names as supplier
  (`address_1/2/3`, `country_code`) — no naming divergence to account for.
- `config/transform/sql/finance/customer/src_fin_customer_raw.sql` — added the
  `nrm_city_subdivision` `CASE` (immediately after `nrm_address_line_4`), identical
  ZA/FI-SK-MX-AX split to supplier's.
- `config/transform/sql/finance/customer/src_fin_customer_id_rnk.sql` and
  `src_fin_customer_rnk.sql` — added `nrm_city_subdivision := fn_initcap(t.nrm_city_subdivision)`
  to the address `STRUCT_PACK` in both files (Level-1/Level-2 aggregation), same placement as
  supplier.
- `config/transform/sql/finance/customer/workday_customer_address_base.sql` — added the gated
  output column, same table-driven `LOWER(TRIM(req.city_subdivision)) = 'required'` pattern.
  `workday_customer_address.sql` needed no change — it's `SELECT t.*` and inherits the new column
  automatically.
- `docs/finance-customer.md` updated — field mapping table row, Stage 2 narrative note, mirroring
  the supplier doc.
- Verified end-to-end via `uv run python examples/fin_customer_debtor.py --run-to-phase
  TRANSFORM`: South Africa (`C-001523`, Gregory Bortz) resolves to `city_subdivision = 'Clifton'`;
  Finland (`C-001345`) and Slovakia (`C-002018`) correctly stay `NULL` pending the GP data ask.
  `validation_customer_issues` shows 0 rows for `src_fin_customer`/`workday_customer_name` rule
  types — no coverage/name regression from the change.

#### Still Todo (next actions)

1. Confirm with finance whether the South Africa customer data (`Clifton`, already in
   `address_2`) needs the same "already correct, no action" treatment, and whether Finland/
   Slovakia customer records (`L00807` Football Association of Finland, `CSO001` Karol Csonto)
   need the same one-time `Address Line 3` GP data ask as the supplier ones (separate ask,
   different records, same request shape) — communication task, not tracked further in this repo.

#### Gotchas

- The GP data ask for Finland/Slovakia is a **one-time migration request to finance**, not a
  pipeline task — don't try to backfill or guess real district/subdivision values in SQL.

### Port newer supplier-side reference-data/phone fixes to customer domain — DONE

Three more supplier-side fixes landed in a later session than the collision-fix item above,
unrelated to it. One is a global reference-data fix already inherited by customer automatically;
the other is a customer-domain file that still needs the port.

#### Done (supplier side)

- `config/data/ref_country.json` — added Lebanon (`LB`), Andorra (`AD`), Mauritius (`MU`): all
  three were missing entirely, causing `workday_supplier_phone.sql`'s country-code lookup to fail
  silently (phone country/international code came out `NULL` for the 2 affected supplier phone
  records).
- `config/data/ref_source_country_code_mapping.json` — added identity entries for `LB`, `AD`,
  `MU`. Missing entries here caused `src_fin_supplier_raw.sql`'s country resolution
  (`COALESCE(m_name.country_code, m_code.country_code, 'GB')`) to silently default these 3
  suppliers' address `country` to "United Kingdom" instead of their real country — even though
  `nrm_city` correctly resolved (Beirut, Escaldes-Engordany, Floreal), only `country` was wrong.
  Verified via `workday_supplier_address` before/after: `S-000293` (Ali Al Chall) UK→Lebanon,
  `S-000390` (Andorran Football Federation) UK→Andorra, `S-007098` (Type 32 Performance)
  UK→Mauritius.
- `config/transform/sql/finance/supplier/workday_supplier_phone.sql` — removed the fax-number
  extraction branch from `cte_supplier_phone` entirely (business decision from finance: exclude
  fax data — Workday's `phone_device_type` field has no valid `"Fax"` value; the one affected
  supplier, Nordomatic UK `S-004963`, had its fax row rejected by EIB). Cleaned up the now-dead
  `'fax'`/`'Fax'` references in the ranking `CASE` and `ORDER BY` too.

**Already verified — no customer action needed for the reference-data part:** `ref_country.json`
and `ref_source_country_code_mapping.json` are global shared tables — confirmed via grep that
`config/transform/sql/finance/customer/src_fin_customer_raw.sql` joins the same
`ref_country`/`ref_source_country_name_mapping`/`ref_source_country_code_mapping` tables. The
Lebanon/Andorra/Mauritius fix is already inherited by the customer pipeline automatically.

**Already verified — the phone fix still needs porting:** `workday_customer_phone.sql` has the
identical fax-extraction-and-`'Fax'`-device-type-hardcoding pattern the supplier file had before
the fix (same CTE structure, same `'fax'`/`'Fax'` literals — confirmed via grep). A live check
(`SELECT count(*) FROM workday_customer_phone WHERE phone_device_type='Fax'`) currently returns
`0` — no customer record currently has fax data reaching that far, so there is no live validation
failure today. Still worth porting for consistency (trivial, cheap change, and this is a one-time
migration, and GP will be frozen once suppliers/customers are loaded — no future GP data will
ever introduce a fax number that could resurface this). **Low priority**: nothing is broken
right now.

#### Done (customer side, this session)

- `config/transform/sql/finance/customer/workday_customer_phone.sql` — removed the fax-extraction
  `UNION ALL` branch from `cte_customer_phone` (sourced `nrm_agg_fax`), removed the dead `WHEN
  'fax' THEN 4` from the ranking `CASE`, removed the dead `WHEN s.phone_type = 'fax' THEN 'Fax'`
  override (now just `s.parsed_phone.device_type` directly, matching supplier), and removed the
  dead `WHEN 'Fax' THEN 3` from the final `ORDER BY`. `nrm_agg_fax` itself is left in place in
  `src_fin_customer_id_rnk.sql`/`src_fin_customer_rnk.sql` (unused downstream) — same as supplier.
- `docs/finance-customer.md` updated per the Docs Currency rule.
- Verified via `uv run python examples/fin_customer_debtor.py --run-to-phase TRANSFORM`:
  `workday_customer_phone` row count unchanged (3,464 before and after), `phone_device_type =
  'Fax'` count unchanged (0 before and after) — confirmed no-regression cleanup.

#### Gotchas

- Was a consistency/latent-risk cleanup with zero current impact, not an active bug — correctly
  done last, after both higher-priority items above.

### "Manchester" resolves to Jamaica's Country_Region ID during EIB conversion — CLOSED, known Workday bug, already handled by 3+ outside this repo

Discovered while triaging the real customer Workday EIB submission failure export (`workday_eib_submit_customer_address_failure_report`,
built from an actual `Submit E2E Customers With Errors Tab.xlsx` load — 438 failure rows across
358 distinct customers, only 3 of which were the already-fixed `city_subdivision` gap). The
single largest category in that export — **"Region Name must be valid for the specified
Country", 140 of 438 rows** — traces to a genuine root cause, but **not one this repo's SQL/Python
code can fix**. 3+ has confirmed independently that their workbook→EIB conversion sets UK
Manchester addresses up as Jamaica Manchester, which corroborates the finding below.

**Correction to an earlier version of this entry**: this was first mis-diagnosed as "Manchester is
an invalid/unmapped Workday region." That was wrong — `ref_workday_county_obsolete` confirms
Manchester is a valid, deliberately-modelled Workday region (`Bolton (obsolete)`, `Oldham
(obsolete)`, `Bury (obsolete)`, `Stockport (obsolete)`, etc. all roll up to it by design), and the
actual Workday workbook instructions guides (`~/Documents/__data/excel/threeplus/Supplier and
Customer Master Record workbooks & Instructions Guide/`) list it as valid too. Our own SQL
(`src_fin_customer_raw.sql`/`src_fin_supplier_raw.sql`) correctly filters by country when
resolving region text, and correctly outputs the plain text `"Manchester"` for these UK addresses
— confirmed by reading the raw *submitted* EIB data directly (see Verification), not just
re-derived pipeline output.

**Actual root cause**: `Manchester` is a valid `Country_Region` value under **both** Jamaica
(`Parish`, reference id `JAM-12`) **and** United Kingdom (`Region`, reference id `GBR-MAN`) in
`config/data/ref_workday_country_state_region.csv` — a genuine naming collision (Manchester
Parish, Jamaica, was historically named after the English city). The raw submitted EIB data for
these UK customers shows `country_region = 'JAM-12'` with `country_region_descriptor =
'Manchester'` — i.e. the display text sent was correct, but the reference ID sent was Jamaica's,
not the UK's (`GBR-MAN`). Whatever process converts the workbook into the actual EIB (a VBA macro,
`InitiateALLAdvanceLoad`, embedded in the `.xlsm` submission workbook and run by 3+ — not
version-controlled text in this repo, unlike `elt_transform_excel`'s `vba_src/`) is resolving the
region by display text alone, without disambiguating by country.

**Not isolated to Manchester**: a scan of `ref_workday_country_state_region.csv` found **58
other region names that collide across multiple countries** (e.g. `Central` → Botswana/Ghana/
Maldives/Papua New Guinea/Paraguay/Zambia; `Saint James` → Barbados/Jamaica; `Toledo` →
Belize/Spain). Any of these could trigger the same class of failure in the conversion tool if our
supplier/customer data touches the colliding countries — Manchester is just the one that actually
surfaced because our data happens to have Greater Manchester UK addresses and no data in most of
the other colliding countries.

**Blast radius (checked via the stable `reference_id`/`nrm_agg_supplier_number` key, NOT the
run-scoped `supplier_id`/`customer_id` — see Gotchas below, which this investigation initially got
wrong before correcting):**
- Supplier: 355 rows / **342 distinct suppliers** currently have region text `'Manchester'` in
  `workday_supplier_address` (all of which would be at risk of the same JAM-12 misconversion, per
  3+'s confirmation that this is systemic to their conversion step, not row-specific).
- Customer: at least 140 of 438 real EIB failure rows are this exact error.

**Resolution: no action needed from this repo.** 3+ already knows about this independently,
identifies it as a Workday platform bug (not something introduced by our workbook output or their
conversion logic), and manually corrects the region on the EIB itself before submission as an
accepted, if inefficient, workaround on their side. Confirmed by the user 2026-07-22.

#### Still Todo (next actions)

1. None for this repo. Nothing to change in `elt_ingest_excel` SQL/Python — our transform layer
   already outputs the correct country-disambiguated text, 3+ is already aware, and they already
   have a working (manual) mitigation in their process. Do not add a `ref_post_code_workday_region`
   translation entry for Manchester or edit `ref_workday_country_state_region.csv`'s IDs.
2. Optional, not currently prioritized: the 58 other cross-country region-name collisions found in
   `ref_workday_country_state_region.csv` carry the same latent-bug shape (see below) — could be
   worth a proactive heads-up to 3+ if any of those countries ever appear in our supplier/customer
   data, but there's no evidence yet that any of them are live issues today.
3. All remaining categories in the real customer EIB failure export have now been triaged (see
   Completed section for the full breakdown): postal-code format → see "Postal code normalization
   fixes" item below; Address Line 3 / Tax `Country_Reference` / Municipality → all already
   resolved, zero live issues; "missing address components" → `subregion` field fix, see "Subregion
   county field" item below (the rest of that category was either already covered by the
   `city_subdivision` item or is isolated genuine source-data gaps, not a systemic bug).

#### Verification

```sql
-- Raw submitted data (ground truth of what was actually sent to Workday) shows the wrong ID:
SELECT t.customer_reference_id, t.customer_name, t.country, t.country_region,
       t.country_region_descriptor, e.error_description
FROM src_fin_wd_submit_customer_address t
INNER JOIN src_fin_wd_submit_customer_address_error e ON e.spreadsheet_key = t.spreadsheet_key
WHERE t.country_region_descriptor = 'Manchester';
-- country_region = 'JAM-12' (Jamaica), country = 'GB' -- confirms the cross-country ID mismatch

-- Confirm Manchester IS a modelled, valid UK region (not a gap on our side):
SELECT * FROM ref_workday_county_obsolete WHERE target_value = 'Manchester';
-- 9 obsolete Greater Manchester boroughs roll up to it by design

-- Current blast radius (stable-key join, not run-scoped customer_id/supplier_id):
SELECT count(*), count(DISTINCT supplier_id) FROM workday_supplier_address WHERE region = 'Manchester';
SELECT count(*), count(DISTINCT customer_id) FROM workday_customer_address WHERE region = 'Manchester';
```

#### Gotchas

- **Do not join the real EIB failure report to current pipeline output using `customer_id`** —
  `workday_eib_submit_customer_address_failure_report` is a frozen snapshot from whenever it was
  last built; `customer_id` (and `supplier_id`) renumber on every pipeline run. Join via
  `customer_reference_id` (the failure report's stable key) → `workday_customer_name.reference_id`
  → current `customer_id`. Got this wrong once during this investigation — the first pass showed
  completely wrong customer names for the same IDs.
- **Don't conclude a reference-data row is "invalid" just because `workday_id` is blank** — a
  first pass here wrongly treated the UK Manchester row's blank `workday_id` as proof it wasn't a
  real Workday value. 113 of ~272 UK rows in `ref_workday_country_state_region.csv` have a blank
  `workday_id`, including some (Liverpool, Sheffield) alongside others that are populated
  (Birmingham, Leeds) — blank is a common data-completeness gap in this extract, not a validity
  signal. Confirm validity against the actual Workday instructions workbook or `ref_workday_county_obsolete`
  roll-ups instead.
- The failure export (`Submit E2E Customers With Errors Tab.xlsx`) has no known supplier
  equivalent file/ingest config in this repo — the supplier blast-radius number above was derived
  by querying current pipeline output directly (`region = 'Manchester'`), not from a real supplier
  EIB failure file. If one exists, ingest it the same way for corroboration.
- The actual workbook→EIB conversion logic (VBA macro `InitiateALLAdvanceLoad`) is not
  version-controlled text in this repo — `elt_ingest_excel` has no `vba_src/` the way
  `elt_transform_excel` does. It cannot be read or fixed as source here; it lives inside the
  `.xlsm` submission workbook itself and is run/owned by 3+.

### Postal code normalization fixes: leading-zero and O/0 typo — DONE (shared UDF, both domains)

Triaged the "postal-code format" category from the real customer EIB failure export (237
historical rows). Cross-checked each distinct failure against **current** live pipeline output
(not the frozen historical snapshot) via `validation_customer_issues`/`validation_supplier_issues`
— most of the 237 historical rows were already resolved by unrelated prior work; 30 distinct
customer records (and a similar-shaped set on supplier) were still genuinely live. Of those, two
clear, safe-to-fix patterns emerged; the rest are genuine bad/incomplete source data, left alone.

#### Done

- `src/elt_ingest_excel/transform/udf/address.py`, `udf_normalize_postal_code` — shared by both
  supplier and customer pipelines:
  1. **Missing leading zero**: for the plain `"postal code must be N digits"` rule branch and the
     US `zip+4`/`#####-####` branch, if the digit count is short by **exactly one** digit, left-pad
     with a zero (`digits.zfill(n)`) instead of returning `None`. Fixes cases like Spain `8017` →
     `08017`, Romania `31106` → `031106`, US `6901`/`6857` → `06901`/`06857` — all countries/regions
     where a leading zero is a genuine, valid part of the postal code (Barcelona-area Spanish codes,
     Bucharest-area Romanian codes, Connecticut ZIPs), consistent with the raw source being a
     numeric field that dropped a leading zero. Deliberately conservative: only pads when short by
     exactly 1 digit, to avoid masking a genuinely wrong value that's missing more.
  2. **UK single-character OCR repair** (`_gb_ocr_repair`): in the `"postal code is made up of two
     parts"` (UK) branch, if the primary and fallback regex both fail, attempt one self-validating
     character repair. This generalises the original single-direction `O→0` fix to the full set of
     canonical letter↔digit confusions in **both** directions — e.g. `HA9 OWS`→`HA9 0WS`,
     `AL7 IEW`→`AL7 1EW`, `S015 2JU`→`SO15 2JU` (the former reverse case), `WCIH 9LG`→`WC1H 9LG`,
     `BRG 0HX`→`BR6 0HX`. It is safe because the guardrails make fabrication impossible:
     - **at most one** character may be substituted;
     - the two trailing inward letters are never guessed from a digit (unverifiable), only the
       inward leading digit and the structured outward code are coerced;
     - a 2-char outward is never coerced (that just forces an incomplete `IG`→`I6`);
     - the result must be the **unique** candidate that passes the strict UK regex (`_PAT_GB`).
     Genuinely ambiguous or incomplete values therefore still return `None` (e.g. `B93 46F`,
     `NE7 76Q`, `WSS 3LW`, `IG 3PS`, the county word `ESSEX`) rather than a wrong guess.
- Verified via direct UDF unit tests (all target cases now resolve correctly; all genuinely-bad
  cases, including the reverse-typo one, correctly still return `None`, not a wrong guess) and via
  full pipeline re-runs of both `fin_customer_debtor.py` and `fin_supplier_creditor.py`
  (`--run-to-phase TRANSFORM`).
- Customer: distinct genuine postal-code FORMAT issues went from 30 → 18 (12 fixed, including all
  10 "Wembley Stadium" (`HA9 0WS`) records shared by multiple FA-related organisations — a bigger
  real-world win than the sample suggested). Supplier: no regressions; same fix applies where
  relevant (29 distinct genuine issues remain, all different-shaped garbage the fix correctly
  leaves alone — e.g. `B93 46F`, `NE7 76Q`).
- Remaining 18 (customer) / 29 (supplier) genuine issues are bad/incomplete source data with no
  safe inference available — e.g. Dutch codes missing their 2-letter suffix entirely, Malta codes
  missing a leading digit, Irish records with a city name or garbage instead of an Eircode, a
  Thai record with what looks like a house/building number in the postcode field, UK codes that
  are simply incomplete (`NE1`, `W1W`) or don't resemble a UK postcode at all (`190002`). These
  need a GP data-correction ask, not a pipeline fix — not yet raised with finance.

#### Verification

```sql
-- Genuine (non-empty-value) postal-code FORMAT issues, both domains:
SELECT DISTINCT customer_id, detail FROM validation_customer_issues
 WHERE col='postal_code' AND rule_type='FORMAT' AND message NOT LIKE 'Postal code is required%';
SELECT DISTINCT supplier_id, detail FROM validation_supplier_issues
 WHERE col='postal_code' AND rule_type='FORMAT' AND message NOT LIKE 'Postal code is required%';
```

#### Gotchas

- The historical EIB failure export is **not** a reliable live-bug indicator on its own — most of
  its 237 postal-code rows were already fixed by unrelated prior work by the time this session
  checked. Always re-verify against current `validation_*_issues` output before treating a
  historical failure as still open.
- Don't over-fix. The zero-pad is scoped to short-by-exactly-one. The UK OCR repair
  (`_gb_ocr_repair`) now does substitute in both directions, but only because it is bounded to a
  **single** substitution, refuses to guess inward letters or coerce a 2-char outward, and requires
  a unique strict-regex match — so it cannot fabricate a plausible-but-wrong postcode. Before
  loosening any of those guardrails, re-run the unit cases: `ESSEX`, `IG 3PS`, `B93 46F` must still
  return `None`. Loosening trades a few confidently-correct fixes for many confident-looking wrong guesses.

### Subregion county field ("Region Subdivision 1") — DONE, both domains

Discovered while triaging the "missing address components" category of the real customer EIB
failure export. Three of its rows read `"Specify the required missing address components for
Finland/South Africa/Slovakia: City Subdivision 1"` — already covered by the earlier
`city_subdivision` item. The other rows named a **different, previously untouched** field:
`"Specify the required missing address components for Côte d'Ivoire/Estonia/Thailand: Region
Subdivision 1"`.

**Root cause**: `ref_workday_validate_customer_address.subregion_county = 'Required'` for Côte
d'Ivoire, Estonia, and Thailand, but `workday_customer_address_base.sql` (and the identical
`workday_supplier_address_base.sql`) hardcoded `CAST(NULL AS VARCHAR) subregion` — the field never
existed anywhere in the transform, on **either** domain. Not a supplier-fixed-then-port situation
like `city_subdivision` — both domains had the identical gap simultaneously, fixed together.

#### Done (both domains, this session)

- `config/transform/sql/finance/customer/src_fin_customer_raw.sql` and
  `config/transform/sql/finance/supplier/src_fin_supplier_raw.sql` — added `nrm_subregion_county`,
  sourced from the raw `county` field, computed **unconditionally for every country** (unlike
  `city_subdivision`, no country allowlist is needed — `county` is already county/subregion-level
  data regardless of country).
- `src_fin_customer_id_rnk.sql`/`src_fin_customer_rnk.sql` and
  `src_fin_supplier_id_rnk.sql`/`src_fin_supplier_rnk.sql` — added
  `nrm_subregion_county := fn_initcap(t.nrm_subregion_county)` to the address `STRUCT_PACK` in all
  four files.
- `workday_customer_address_base.sql` and `workday_supplier_address_base.sql` — replaced the
  hardcoded `NULL` with a gated output column: `CASE WHEN LOWER(TRIM(req.subregion_county)) =
  'required' THEN u.address.nrm_subregion_county ELSE NULL END`, table-driven off
  `ref_workday_validate_{customer,supplier}_address` — same pattern as `city_subdivision`.
- `docs/finance-customer.md`, `docs/finance-supplier.md` updated (field mapping table row +
  narrative note in both).
- Verified end-to-end via full pipeline re-runs of both `fin_customer_debtor.py` and
  `fin_supplier_creditor.py` (`--run-to-phase TRANSFORM`): Côte d'Ivoire (`Lagunes`/`Lagune`) and
  one Thailand record (`Bangkok`) now resolve correctly on customer; Estonia (`Harjumaa`) and
  Thailand (`Bangkok`) resolve on supplier. One customer record (`Patchara Poosarn`, Thailand)
  correctly stays `NULL` — raw source `county` is genuinely blank for that record, not a pipeline
  gap.

#### Verification

```sql
SELECT customer_id, customer_name, country, region, subregion, city FROM workday_customer_address
 WHERE country ILIKE '%Ivoire%' OR country IN ('Estonia','Thailand');
SELECT supplier_id, supplier_name, country, region, subregion, city FROM workday_supplier_address
 WHERE country ILIKE '%Ivoire%' OR country IN ('Estonia','Thailand');
```

#### Gotchas

- `county` already feeds `region` too (via the `r1`/`r2`/`r3`/`r4`/`rx`/`rxo` join chain in
  `src_fin_customer_raw.sql`) for some countries — `region` and `subregion` can legitimately end up
  holding the same value (e.g. Thailand's `Bangkok` in both) when `county` is our only source of
  subdivision data. This is expected, not a duplication bug: Workday requires both fields
  separately, and duplicating is the same precedent set by `city_subdivision`'s South Africa case.

### Supplier bank/settlement account data quality — TRIAGED, no code fix possible

Discovered when asked whether all transform work was complete — a full, unfiltered
`validation_supplier_issues` summary (not just the narrowly-scoped queries used to verify the
address-related fixes above) surfaced 110 `BANK_FORMAT` + 1 `BANK_ACCOUNT` issues on
`workday_supplier_settlement_account` that hadn't been looked at at all this session. Not part of
the customer EIB export (which is address-only); no supplier or bank/settlement EIB failure file
is ingested in this repo to cross-check against.

**Lesson driving this item's discovery**: repeatedly re-running the full pipeline and checking
"did my change regress the area I just touched" (postal code, address fields) is not the same as
periodically checking "what does the full current validation state look like." `OK
validate_supplier_settlement_account.sql` in the pipeline's console output only means that SQL
*executed without error* — it says nothing about how many issues it found. Always run the
unfiltered summary before declaring a domain regression-free.

#### Done (triage, this session)

Checked every category against raw source data (`src_fin_supplier_base`) for a sample of affected
suppliers in each — 107 distinct suppliers affected out of ~7,463 with a settlement account row
(~1.4%). All trace to genuine GP data gaps; none are pipeline bugs:

- **Routing Transit Number missing, US** (67 suppliers) — `eft_transit_routing_no` and
  `eft_bank_code` both blank at source; these suppliers have a SWIFT/BIC code and account number on
  file instead. No reliable way to derive a US ABA routing number from a SWIFT/BIC code (it
  identifies the bank, not the specific routing number) — genuinely unrecoverable without GP data.
- **Bank Code missing** (23 suppliers, various countries — Philippines, Georgia, Thailand, etc.) —
  `eft_bank_code` blank at source.
- **IBAN wrong length / garbage** (11 suppliers) — `iban` field is populated but either drastically
  short of the required length for that country, or a literal `"0"` placeholder. Not a formatting
  bug on our side (`nrm_bank_iban` just uppercases and strips spaces/hyphens) — the raw value
  itself is wrong/incomplete in GP.
- **BSB Number missing, Australia/NZ** (6 suppliers) — blank at source.
- **Institution Number missing, Canada** (2 suppliers) — blank at source.
- **IFSC Code missing, India** (1 supplier) — blank at source.
- **Sort code missing + SWIFT code entered into the IBAN field** (1 supplier, `S-003683` Kelsey
  Entertainment / Revolut) — GP's own `iban` field holds `REVOGB21` (clearly a SWIFT/BIC-8 format,
  not an IBAN), while the actual `swift_address` field is blank and `eft_bank_code` (sort code) is
  also blank. Confirmed via `src_fin_supplier_base` raw columns directly — our transform just
  passes the raw `iban` value through unmodified; this is a GP data-entry mistake (SWIFT typed into
  the wrong field), not a bug in `src_fin_supplier_raw.sql`. Deliberately not "corrected" by
  detecting the SWIFT-8 shape and moving it to the swift field — that would be inferring intent
  behind a data-entry error, the same category of risk this session already declined to take with
  postal codes (e.g. the `S031`→`SO31` case left alone).

**No code changes made or needed.** This is a data-completeness/correction ask for finance, not a
transform bug — added to the running list of data asks alongside `city_subdivision` and the
postal-code list above.

#### Verification

```sql
SELECT rule_type, count(*) FROM validation_supplier_issues WHERE table_name LIKE '%settlement%' GROUP BY 1;
SELECT DISTINCT supplier_id, message FROM validation_supplier_issues WHERE rule_type IN ('BANK_FORMAT','BANK_ACCOUNT') ORDER BY 2;
```

#### Gotchas

- Don't try to "fix" the Kelsey Entertainment case by moving the SWIFT-shaped value from `iban` to
  `swift_address` — that's guessing at what GP meant, not correcting a pipeline bug. Flag it to
  finance as a data-entry correction instead.
- When re-verifying "no regressions" after a change, run the unfiltered validation summary for the
  whole domain, not just a query scoped to the table/column you touched — see the lesson note
  above.

#### Follow-up finding: `for_supplier_connections_only` alone may not be sufficient (untested assumption)

Verified against the original source material for the "must have only 1 active settlement bank
account" fix (`~/Downloads/Master Suppliers/Master Suppliers.eml` from Gayle Henderson, hr-path.com
— the actual email behind the 564-failure number quoted throughout this file, with the "Master
Supplier Sandbox V1 with errors.xlsx" EIB-errors attachment). Confirmed `S-000509` (ASANA, GP ref
`ASA001|ASA002`) as the real example this repo's fix was modelled on: two accounts, HSBC (primary,
unrestricted) and NatWest (secondary), where Gayle's manual fix note reads exactly "Added Y into
'For Supplier Connections on second account' **and added accepted payment type of Wire**."

`workday_supplier_settlement_account.sql`'s existing logic (see "Done" above,
`account_rnk > 1 THEN 'Y'`) sets **only** the flag — `payment_types_plus` stays blank on every
secondary account. Live check: **73 secondary accounts across 72 suppliers currently have the flag
set and `payment_types_plus` blank.** Because Gayle set both fields together in the one real
example we have, there is no evidence either way whether the flag alone clears Workday's check —
the working assumption that it's sufficient (implicit in the existing code comment) is untested.

**Not resolved — needs a decision, not code changed this session.** Two options on the table:
1. Leave as-is (flag only) and add the payment-type assignment to the running data-ask list for
   finance (same treatment as `city_subdivision`/bad postal codes) — assigning a specific payment
   type (Wire/ACH/etc.) per secondary account is a business decision, not derivable from GP data,
   consistent with the existing code comment's reasoning.
2. Add a validation rule flagging `for_supplier_connections_only = 'Y' AND payment_types_plus IS
   NULL` so these 72 suppliers surface as their own explicit, actionable list rather than being
   invisible to `validate_supplier_settlement_account.sql` (they currently pass our internal
   validation, so this gap wouldn't be caught by re-running the pipeline).

```sql
-- The 73 currently-affected accounts:
SELECT supplier_id, supplier_name, bank_name, bank_country
FROM workday_supplier_settlement_account
WHERE for_supplier_connections_only = 'Y' AND (payment_types_plus IS NULL OR TRIM(payment_types_plus) = '');
```

### Running 3+'s workbook macros directly — ABANDONED, do not re-attempt without new information

Goal was to run 3+'s own workbook validation and EIB-creation macros ourselves (on the published
`.xlsm` output), purely for our own internal confidence that the transforms are correct — **not**
to remove 3+ from the process; they'd still redo everything themselves as normal.

This repo already has partial infrastructure for this (`examples/fin_supplier_creditor_create_eib.py`,
`examples/fin_supplier_creditor_validate_vba.py`, `docs/VBA_VALIDATION_RUNNER.md`,
`src/elt_ingest_excel/macro/`), built in a prior session. A previous session reportedly spent close
to a full day on this without success. This session re-investigated to find the root cause before
deciding whether to keep trying — found it, and it's a dead end on macOS.

#### Investigation findings (this session)

- Confirmed via `oletools` (read-only VBA inspection, no execution) that **both** the customer and
  supplier `.xlsm` templates contain the same rich validation framework (`ModValidation.bas`,
  `ConfigModule.bas`) and EIB-creation macros (`TempCreator.bas`:
  `InitiateSPECAdvanceLoad`/`InitiateALLAdvanceLoad`) — so this was never a "customer doesn't have
  it" gap, both domains have identical macro infrastructure.
- Found two real bugs in the existing automation layer along the way:
  1. `fin_supplier_creditor_validate_vba.py` hardcoded `workday_fin_creditor_supplier_active_v5.xlsm`
     as its default workbook path — the actual publish config produces `..._v1.xlsm`. Stale/wrong
     default, would silently fail with a file-not-found unless `--workbook` was passed explicitly.
  2. `run_excel_vba_macro()` (`excel_macro_runner.py`) invokes macros via AppleScript's `run VB
     macro (contents of m)` with **no way to pass arguments**. The default validation macro
     (`runAllValidationsFromSheet`) requires an `IRibbonControl` argument that AppleScript cannot
     supply at all — confirmed via a live test, which failed with `Cannot run VB macro ... (-2700)`
     (an Excel argument-mismatch error code). The underlying logic sub, `ValidateSheets(domain,
     "All")`, takes plain string arguments and isn't in a private module, so it's a better call
     target — but the runner has no argument-passing support to reach it either.
  3. Found the correct AppleScript syntax for passing arguments (`run VB macro "Name" arg1 "x" arg2
     "y"`, discovered via Excel's own `Excel.sdef` — the intuitive `arguments {...}` record/list
     syntax is wrong and just produces AppleScript syntax errors) — recorded here in case this is
     ever revisited.
- **The real, blocking finding**: even with the corrected argument syntax, calling `ValidateSheets`
  directly failed with an AppleEvent timeout, and manually invoking a macro in the workbook
  (`clearSheetValidations`) produced a native **"Can't find project or library"** VBA compile
  error. Read-only inspection of the workbook's VBA project references (`xl/vbaProject.bin`, `VBA
  dir` stream) found the actual cause: the project references `Microsoft Scripting Runtime`
  pointing at a nonexistent path for `scrrun.dll` (a Windows-only COM library, no macOS equivalent
  at all), plus a duplicate `Microsoft Forms 2.0 Object Library` reference pointing at a literal
  Windows machine path (`C:\Users\GAYLE~1.HEN\AppData\Local\Temp\...`) — both clearly artifacts
  from whoever last edited this template on Windows. A broken/missing VBA reference blocks
  compilation of the **entire** project on load, not just code paths that actually use the missing
  library — which is consistent with every macro in the workbook failing, regardless of which one
  is invoked or how.

#### Why abandoned

- **This is a genuine macOS/Windows compatibility problem in a third-party template file, not a
  bug in this repo.** 3+ almost certainly does all their own work on Windows, where `scrrun.dll`
  exists natively — they likely never encounter this. Fixing it would mean altering 3+'s template's
  VBA project structure (removing/relinking the broken reference via the VBA Editor's References
  dialog), which is their file to maintain, not ours to patch.
- **Poor cost/benefit.** A full day was already spent on this in a prior session with no success;
  this session's re-investigation, despite finding and fixing two real bugs in our own automation
  layer along the way, hit the identical wall from a different angle (compile error → AppleEvent
  timeout). No quick fix is in reach — it would likely require an actual Windows environment.
- **The underlying goal is already met through better means.** "Confidence that our transforms are
  correct" now rests on stronger evidence than re-running 3+'s own checks would add: the real
  Workday EIB rejection file (438 actual failure rows, fully triaged — see the items above) and a
  full, unfiltered sweep of our own `validation_customer_issues`/`validation_supplier_issues`
  across both domains. Running 3+'s macros was always meant as a secondary cross-check, not the
  primary verification mechanism.

#### Gotchas (if this is ever revisited)

- Don't re-attempt without first confirming a Windows environment is available, or that 3+'s
  template's broken VBA references have been cleaned up — the compile error blocks the entire VBA
  project, not just the specific macro being called, so no amount of AppleScript fixes in this
  repo's automation layer can work around it.
- If it is revisited: the `fin_supplier_creditor_validate_vba.py` stale `v5` default path and the
  `run_excel_vba_macro()` missing-argument-support bugs are real and still there, independent of
  the VBA-reference blocker — worth fixing those regardless, and the correct AppleScript
  argument-passing syntax (`arg1`/`arg2`, not `arguments {...}`) is recorded above.
- Excel may be left in a stuck/blocked state (a native error dialog) after a failed macro
  invocation like this — check for and manually dismiss/close before trying anything further.

## Completed

- [x] Finland/Slovakia city_subdivision fix (2026-07-24) — `city_subdivision`
      now falls back to `nrm_city` when Workday won't accept a municipality
      (`workday_supplier_address_base.sql`); 15 records / 30 issues → 0, no GP
      data ask needed. **Customer domain likely has the identical bug (ported
      design) — not yet checked/fixed.** See "DONE (2026-07-24)" in Resume
- [x] IBAN-embedded bank CODE fix (2026-07-24) — `ref_iban_structure.json` +
      fill-blank-only `iban_bankcode` fallback in `src_fin_supplier_raw.sql`
      populates `nrm_bank_code_routing_number` from the supplier's own IBAN for
      GE/KZ/UA/CR (5 records); 72→67 banking issues, zero regressions, source
      data untouched — see the "DONE (2026-07-24)" note in Resume
- [x] Standardised doc names to contract paths (architecture.md → ARCHITECTURE.md, configs.md → CONFIGURATION.md, CLI.md → RUNBOOK.md)
- [x] Supplier/creditor half of vendor-ID business-unit-collision fix — see "Done"
      under the item above
- [x] Vendor-ID business-unit-collision fix, customer/debtor side — verified NOT
      applicable: no `customer_number` is reused across business units in the data
      (0 rows on both the strict and the weaker precondition check against
      `src_fin_customer_raw`); no code port needed
- [x] `city_subdivision` address field, customer side — ported from supplier; South
      Africa resolves, Finland/Slovakia NULL pending the GP data ask — see "Done"
      under the item above
- [x] Fax-extraction removal ported to `workday_customer_phone.sql` — no-regression
      cleanup, 0 fax rows before and after, row count unchanged — see "Done" under
      the item above
- [x] "Manchester" → Jamaica Country_Region ID EIB rejection — root-caused (Workday
      naming collision between Jamaica Parish and UK Region, both called
      "Manchester"); confirmed not a bug in this repo's SQL; confirmed 3+ already
      knows and manually corrects it on the EIB as an accepted workaround — no
      action needed, see item above
- [x] Postal code normalization fixes (leading-zero, UK O/0 typo) in shared
      `udf_normalize_postal_code` — customer genuine issues 30→18, no supplier
      regressions, verified via unit tests and full pipeline re-runs — see "Done"
      under the item above
- [x] Address Line 3 / Tax `Country_Reference` / Municipality EIB failure
      categories — all triaged and confirmed already resolved by prior
      (untracked) work; zero live issues against current
      `validation_customer_issues`
- [x] `subregion` ("Region Subdivision 1") field — genuinely missing in both
      domains, fixed together (not a supplier→customer port); table-driven
      exposure for Côte d'Ivoire/Estonia/Thailand — see "Done" under the item
      above
- [x] Full triage of the real customer Workday EIB failure export (438 rows,
      all 8 error categories) — every category now resolved, fixed, or
      confirmed external to this repo; see Resume section for the summary
- [x] Supplier bank/settlement account data quality triage (111 issues, 107
      distinct suppliers) — all confirmed genuine GP data gaps, no code fix
      possible; see "Supplier bank/settlement account data quality" item above
- [x] Investigated and ABANDONED running 3+'s workbook macros directly —
      root cause found (broken Windows-only VBA references in 3+'s template,
      not fixable from this repo); decided not worth pursuing since transform
      correctness is already independently verified — see "Running 3+'s
      workbook macros directly" item above
