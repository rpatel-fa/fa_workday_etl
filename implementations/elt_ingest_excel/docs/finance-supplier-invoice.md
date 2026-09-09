# Finance — Supplier Invoice / Credit Note → Workday

**Single source of truth for the supplier-invoice + credit-note (adjustment) migration.**
Cold-start sessions should be able to continue from this file alone.
**Read §Ownership & Scope and §Known invariants before touching any SQL or the reconcile script.**

Config lives under `config/{ingest,transform/sql,validate,reconcile,publish}/finance/supplier_invoice/`.
Run commands are in [CLI.md](CLI.md) (not duplicated here).

This document is the supplier-invoice analogue of [finance-supplier-po-receipt.md](finance-supplier-po-receipt.md) (PO/Receipts) and [finance-customer.md](finance-customer.md) (Customer Invoices). The supplier **master/creditor** pipeline is documented separately in [finance-supplier.md](finance-supplier.md) — that doc is about creating `S-` Workday supplier records; **this** doc is about loading the outstanding AP (invoices and credit notes) that sit *against* those supplier records.

The master creditor (`S-` supplier records) **must already be loaded** before this pipeline runs, because invoice rows gate on `wd_supplier_id IS NOT NULL` (exact-then-bare vendor match against the Workday supplier map projection).

---

## Ownership & Scope

- Module: `supplier_invoice`
- SQL folder: `elt_ingest_excel/config/transform/sql/finance/supplier_invoice/`
- Ingest config: [`config/ingest/finance/supplier_invoice.json`](../config/ingest/finance/supplier_invoice.json)
  - Primary source: `AP DTL TB AUG 4 2026.xlsx` / three BU-specific sheets (`FA AP DTL TB`, `NFC AP DTL TB`, `WNSL AP DTL TB`) → `fin_supplier_invoice_fa`, `fin_supplier_invoice_nfc`, `fin_supplier_invoice_wnsl`
  - Secondary ingest (position 2, **NOT published** — used to build the supplier map only): Workday Supplier Extract → `fin_workday_extract_supplier` (shared with PO/receipt pipeline)
- Publish config: [`config/publish/finance/publish_supplier_invoice.json`](../config/publish/finance/publish_supplier_invoice.json) (two-workbook publish: Invoice + Adjustment)
- Verification DB: `~/Documents/__data/duckdb/rpatel.duckdb`
  - Read-only inside the TRAE sandbox. TRANSFORM/PUBLISH must be run from a user-owned shell (see §Sandbox limitations).
- Loaded Workday authority (target for post-load tie-back): `~/Documents/__data/excel/workday/extract/ap_invoices/Operational_Aging_Detail_by_Due_Date (1).xlsx` — the 3+-team Operational Aging Detail by Due Date export from the Workday tenant after load.

---

## Status (start here)

**Amounts CLEAN; invoice-date parity flags 4 rows for REVIEW (re-run 12.08.2026).**

- **Amounts/vouchers — CLEAN.** Grand total + credit-note exact-key + supplier buckets + invoice fuzzy all tie to the penny (Σ = £32,790,897.87 every leg). This is the original 10.08.2026 sign-off and it still holds.
- **GP source ↔ EIB document-date (invoice-date) — 100%.** All 761 published INVOICE rows: `column_af` == GP `invoice_date` exactly, 0 differences. Our population of the document-date is faithful to GP.
- **EIB/source ↔ Workday loaded invoice-date — 4 rows REVIEW.** The date-parity layers (tabs 4B/7B/5B) were added *after* the 10.08 sign-off (commit `cf3b16b`), so the original run never exercised them. Re-run on 12.08.2026: 736 DATE_MATCH + 21 benign DATE_ORDER artefacts + **4 DATE_MISMATCH** — all supplier **D R IMAGES** (vouchers `FAV0114322`–`FAV0114325`). We sent `04/12/2026`; Workday stored `12/04/2026` (day/month transposition on load). The underlying GP invoice date is itself a typo — `04/12/2026` should be `04/12/2025`. Credit-note dates 62/62 exact. **No other invoice is in issue on the date axis.**
- **Out of scope (not a load defect):** the GP consultant's ~90 highlighted *Due Date* differences and the FX-converted GBP-balance differences. We never send a due date — Workday derives it from payment terms + the invoice date — so Due Date divergence is a payment-terms reconcile, and GBP-balance divergence is an FX-rate reconcile. Neither is on the axis this pipeline controls.

**Action:** correct the 4 D R IMAGES invoice dates in GP source (`2026`→`2025`) and confirm Workday holds `04/12/2025`, not `12/04/2026`.

Two EIB workbooks are produced from a single staged source (`src_fin_supplier_invoice` — 823 rows for PRODUCTION):

| Publish table | Workbook template | Output file name prefix | Target sheet | dataRow | Rows (PRODUCTION 04.08.26 extract) |
|---|---|---|---|---|---:|
| `workday_supplier_invoice` | `Submit_Supplier_Invoice_v45.2.xlsx` | `workday_fin_supplier_invoice_v1` | Submit Supplier Invoice | 6 | **761** (INVOICE only) |
| `workday_supplier_credit_note` | `Submit_Supplier_Invoice_Adjustment_v46.0.xlsx` | `workday_fin_supplier_invoice_adjustment_v1` | Submit Supplier Invoice Adjustm | 6 | **62** (non-INVOICE = Adjustment / Credit Note) |

### Run the pipeline

```bash
cd elt_ingest_excel
uv run python examples/fin_supplier_invoice.py --run-to-phase PUBLISH
```

TRAE sandbox blocks DuckDB write-open: inside TRAE run with `--run-to-phase INGEST` at most; for TRANSFORM and above run the command in your own Terminal shell.

### Post-load reconciles — RESULTS (✅ amounts / ⚠ 4 invoice-dates)

Two reconcile workbooks are produced per run:

| Reconcile | Script | What it proves | Result |
|---|---|---|---|
| **Upstream three-way** | [generate_supplier_invoice_three_way.py](../scripts/generate_supplier_invoice_three_way.py) | Raw Source AP TB ↔ Our final Staging ↔ Our EIB publish tables — proves we are sending the correct voucher numbers and amounts in the first place. | ✅ **CLEAN** (10.08.2026) 823/823 matched, 0 amt deltas, Σ=£32,790,897.87 all legs |
| **Post-load 4th-layer tieback (amounts)** | [generate_supplier_invoice_post_load_reconcile.py](../scripts/generate_supplier_invoice_post_load_reconcile.py) | Our EIB publish tables ↔ Workday-loaded Operational Aging extract — proves Workday loaded exactly what we sent. | ✅ **CLEAN** (10.08.2026) Grand Δ=£0.00, CN 62/62 exact, Sup 364 Match, Inv 761/761 |
| **Post-load date parity (4B/7B/5B)** | same script (layers added post-sign-off) | Sent invoice-date (`column_af`/`column_y`) ↔ Workday Aging `Transaction Date` — proves the *dates* loaded as sent. | ⚠ **REVIEW** (12.08.2026) 736 match + 21 order-artefact + **4 DATE_MISMATCH** (D R IMAGES `FAV0114322`–`5`, day/month transposed on load); CN 62/62 date-exact |

Full details on both are in the two dedicated sections below (three-way upstream → post-load downstream).

```bash
cd elt_ingest_excel
# Upstream proof (we are sending correct data)
uv run python scripts/generate_supplier_invoice_three_way.py --label PRODUCTION
# Downstream proof (Workday loaded what we sent)
uv run python scripts/generate_supplier_invoice_post_load_reconcile.py --label PRODUCTION
```

Both write to 3 locations with the same pattern (see §File locations):
- **PRIMARY (always success from your shell):** `~/Documents/__data/excel/workday/reconcile/`
- **STAGE (always):** `/tmp/`
- **FALLBACK (inside TRAE sandbox):** `elt_ingest_excel/scripts/_output/`

TRAE sandbox blocks writes under `__data/` (reading works fine). Re-run both commands from your own Terminal to land the workbooks in the canonical reconcile dir.

---

## Source Excel files (INGEST phase)

### AP Detail Trial Balance (source workbook: `AP DTL TB AUG 4 2026.xlsx`)

Path: `~/Documents/__data/excel/finance_ref/supplier_invoice/`

| Sheet | DuckDB table | Business-unit filter embedded |
|-------|--------------|-------------------------------|
| FA AP DTL TB | `fin_supplier_invoice_fa` | FA (The FA Group - FA) |
| NFC AP DTL TB | `fin_supplier_invoice_nfc` | NFC (National Football Centre) |
| WNSL AP DTL TB | `fin_supplier_invoice_wnsl` | WNSL (Wembley National Stadium) |

All three sheets share an identical column layout; the pipeline UNIONs them in `src_fin_supplier_invoice_base.sql` and re-derives `business_unit` from the originating table name. headerRow=1, dataRow=2.

### Workday supplier extract (position 2 in ingest, NOT published, used to build the map only)

| Sheet | DuckDB table |
|-------|--------------|
| Sheet1 | `fin_workday_extract_supplier` |

headerRow=13, dataRow=14. Feeds `src_fin_workday_supplier_map.sql` (shared pattern with the PO/receipt pipeline). This map is what gates the invoice pipeline — a row missing from the map means `wd_supplier_id IS NULL` and the invoice is dropped at Stage 3 + tracked in `validate_supplier_invoice_supplier.sql`.

---

## Staging table lineage (TRANSFORM phase)

SQL files execute in the order defined in [config/transform/sql/finance/supplier_invoice/order.txt](../config/transform/sql/finance/supplier_invoice/order.txt) (7 files):

1. `src_fin_workday_supplier_map.sql`
2. `src_fin_supplier_invoice_base.sql`
3. `src_fin_supplier_invoice_nrm.sql`
4. `src_fin_supplier_invoice.sql`       — final staged (gate applied)
5. `workday_supplier_invoice.sql`       — INVOICE type → EIB
6. `workday_supplier_credit_note.sql`   — non-INVOICE type → Adjustment EIB
7. `reconcile_supplier_invoice_unresolved.sql`

### Stage 1 — `src_fin_supplier_invoice_base`

**File:** `src_fin_supplier_invoice_base.sql`

UNIONs the three BU-specific raw tables. Adds:
- `business_unit` literal derived from source table (`FA` / `NFC` / `WNSL`)
- Semantic aliases for all column names (raw AP TB columns carry GP field names).

No filtering or normalisation.

### Stage 2 — `src_fin_supplier_invoice_nrm`

**File:** `src_fin_supplier_invoice_nrm.sql`

Adds four key derived columns via the exact-then-bare Workday supplier map:

| Column | Derivation | Notes |
|--------|-----------|-------|
| `wd_supplier_id` | `COALESCE(s_exact.supplier_id, s_bare.supplier_id)` — exact `vendor_id-BU` first, bare `vendor_id` fallback | Same pattern as PO commitment stack; supplier map is built from position-2 Workday supplier extract. |
| `wd_supplier_name` / `wd_supplier_reference_id` | Same `COALESCE(exact, bare)` pattern — used for traceability in reconcile reports | |
| `nrm_document_date` | Triple-format `TRY_STRPTIME` over the raw text `invoice_date` (`%Y-%m-%d %H:%M:%S` / `%Y-%m-%d` / `%d-%m-%Y`), reformatted `dd/mm/yyyy` for EIB | Invoices and credit notes share one format — Workday invoice_date column = adjustment_date column in the credit-note EIB. |
| `nrm_company_id` | Explicit FA → `'THE FA GROUP - FA'`, NFC → `'THE FA GROUP - NFC'`, WNSL → `'THE FA GROUP - WNSL'` mapping (matches what Workday's Business Entity picklist stores). | These exact strings are what the Workday EIB templates require in `column_r`/`column_p`. |
| `nrm_transaction_amt` | Currency-branched: GBP = `gbp_current_trx_amount`; JPY = `local_currency_current_transaction_amount` rounded to 0 dp; else = local_currency_current_transaction_amount as DOUBLE | JPY must reach Workday as whole-integer VARCHAR; same pattern as PO/receipt EIBs. |

**Suppliers that don't resolve are not dropped yet** — they flow through with NULL `wd_supplier_id` so the gate in Stage 3 is auditable (see validation).

### Stage 3 — `src_fin_supplier_invoice` *(final invoice/credit-note staged table)*

**File:** `src_fin_supplier_invoice.sql`

Gate: `wd_supplier_id IS NOT NULL AND nrm_transaction_amt <> 0`.

This is the **single source of truth** for both EIB publish tables. Splits happen downstream — not here.

Row counts for the PRODUCTION extract (AP DTL TB AUG 4 2026): 823 rows survive the gate (761 `INVOICE` + 62 credit/adjustment).

### Stage 4 — Publish tables (split by document type)

**Invoice EIB publish:** [workday_supplier_invoice.sql](../config/transform/sql/finance/supplier_invoice/workday_supplier_invoice.sql)
- `WHERE UPPER(TRIM(document_type)) = 'INVOICE'`
- `column_t` (currency) = raw `currency_id` (no trim in SQL; **trailing spaces live in the EIB text** and resolve correctly when Workday reads it back — verified 761 rows match the loaded extract)
- `column_bg` = `voucher_number` (source system invoice number, not Workday's tenant ref; the post-load tieback's fuzzy match is needed because Aging report stores Workday refs, not voucher numbers)
- `column_gf` = `nrm_transaction_amt` (full gross — 1 invoice line = 1 document line; no multi-line splits)
- `ROW_NUMBER()` over (BU, supplier_id, document_date, document_number) → `column_b` spreadsheet_key. One line per doc = `column_ek = '1'` always.

**Credit Note / Adjustment publish:** [workday_supplier_credit_note.sql](../config/transform/sql/finance/supplier_invoice/workday_supplier_credit_note.sql)
- `WHERE UPPER(TRIM(document_type)) != 'INVOICE'`
- **Column letters differ from the invoice EIB** because credit-note template v46 rearranges: `column_k` = supplier_invoice_number, `column_p/column_r/column_s` = company/currency/supplier (invoice uses column_r/column_t/column_u), `column_y` = adjustment_date, `column_ap/column_eq` = amounts. The raw SQL SELECT in Stage 5 and the post-load tieback script each handle these offsets explicitly — **never swap column letters blindly between the two publish files**.
- `column_k` is where the shared Workday document IDs (`FAP0148163`, `FAV0079667`, `FACRN200014544`, …) live; the post-load tieback's Layer-5 exact-merge matches 62 of 62 on this column.

---

## Validation

Runs automatically as part of the pipeline for phases ≥ TRANSFORM, or standalone:

```bash
uv run python examples/fin_supplier_invoice.py --run-to-phase TRANSFORM
```

Validation files in `config/validate/finance/supplier_invoice/`:

| File | Purpose |
|---|---|
| `init.sql` | Creates / re-seeds `validation_supplier_invoice_issues` (same schema pattern as creditor master: table_name, rule_type, col, supplier_id, detail, message — here `supplier_id` = `wd_supplier_id`). |
| `validate_supplier_invoice_supplier.sql` | Every row in `src_fin_supplier_invoice_nrm` with `wd_supplier_id IS NULL` — dropped at Stage 3 gate. These are supplier-vendor codes that GP/finance haven't mapped in the supplier master yet. |
| `validate_supplier_invoice_missing_payment_term.sql` | Rows where the related Workday supplier has no resolvable payment terms — downstream EIB load rejects invoices linked to suppliers missing default payment terms even if the invoice row itself is clean. |

Query after running:
```sql
SELECT table_name, rule_type, COUNT(*) n
  FROM validation_supplier_invoice_issues
 GROUP BY 1, 2 ORDER BY 1, 2;
```

---

## Upstream 3-way tie-back (Source AP TB ↔ Our Staging ↔ Our EIB publish)

**Script:** [scripts/generate_supplier_invoice_three_way.py](../scripts/generate_supplier_invoice_three_way.py)
**Analogue:** [generate_finance_review_three_way.py](../scripts/generate_finance_review_three_way.py) (PO/Receipt three-way — same 8-sheet layout, banner + header styling, verdict cell colouring, primary/fallback write paths, tolerance 0.005).

Proves we are loading the **correct numbers from the raw source** into the EIB we send to Workday. This is the upstream proof that our base/nrm/staging transforms haven't dropped, added, or mangled any vouchers.

### Join key: voucher_number (clean, 100% shared, 100% unique)

Verified on the PRODUCTION Aug-04 extract:
- Raw 3-BU UNION (vendor_id non-null filter, mirroring base.sql WHERE): **823 distinct voucher_numbers**
- `src_fin_supplier_invoice_base`: **823 distinct**
- `src_fin_supplier_invoice_nrm`: **823 distinct**
- `src_fin_supplier_invoice` (gate applied): **823 distinct**
- EIB INVOICE (`workday_supplier_invoice.column_bg`) + EIB CN (`workday_supplier_credit_note.column_k`): **761 + 62 = 823 distinct**

Voucher numbers are the GP/voucher-system IDs (`FAINVYYYYNNNNN` for source invoice uploads, `FAVNNNNNNN` for payments, `FAPNNNNNNN` / `FACRNNNNNNNNNNNN` for credit notes). They travel unchanged across all 5 layers — no fuzzy matching required in the three-way script.

Amount compare on `local_currency_current_transaction_amount` (raw) = `nrm_transaction_amt` (staging nrm) = `column_gf` invoice OR `column_eq` credit note (EIB publish). Verified per-voucher on all 823 rows: 0 deltas > 0.005.

### Run standalone

```bash
cd elt_ingest_excel
uv run python scripts/generate_supplier_invoice_three_way.py --label PRODUCTION
```

Optional flags:
- `--duckdb <path>` — override default DuckDB at `~/Documents/__data/duckdb/rpatel.duckdb`.
- `--out-dir <dir>` — override primary output dir (defaults to `~/Documents/__data/excel/workday/reconcile`).
- `--label <TEXT>` — tag appended to workbook filename; use `PRODUCTION` for sign-off runs.

### 8 sheets output (mirrors `generate_finance_review_three_way.py`)

1. **1-Executive Summary** — 7-row matrix: Row counts, STE presence, amt OK / MISMATCH counts, Grand totals (Source / Staging / EIB + pairwise Δs), **VERDICT cell** coloured GREEN for CLEAN, RED for REVIEW.
2. **2-By Txn Type** — INVOICE vs CREDIT_NOTE, rows + amounts + deltas across all 3 legs.
3. **3-By Company Currency** — (Company × Currency × Txn Type) rows/amounts, 15 groups total (FA/WNSL/NFC × 5 ccys × 2 types = 30 minus empty combos = 15).
4. **4-Currency Coverage** — CHF / EUR / GBP / IDR / USD all "BOTH" (= present Source, Staging, EIB); Coverage status.
5. **5-Per-Voucher Exact** — All 823 vouchers, Source/Stage/EIB amounts, pairwise Δ columns, `Presence (STE)` key, Amount status, Overall status (MATCH / REVIEW). Overall column cell coloured green/red.
6. **6-Supplier Bucket** — (Company × normalised-supplier × ccy) 364 groups, rows + amounts across 3 legs, pairwise Δs, Status (Match / Source only / Staging only / EIB only / Review) yellow/green/red filled.
7. **7-Drift & Exceptions** — Every voucher failing presence OR amount equality. For a clean run the body is a single row reading `"NONE — all 823 vouchers matched perfectly across all 3 legs."`
8. **8-Notes** — 13-section methodology: Scope, Join key, Amount compare, Company norm, Supplier-name norm, Txn type derivation, Trailing ccys spaces, JPY handling, Gate row-drop audit, Workbook verdict, How to run, TRAE sandbox note, Analogue script.

### PRODUCTION result (10/08/2026 run):

| Metric | Value |
|---|---:|
| Source rows (AP TB raw) | **823** |
| Staging rows (src_fin_supplier_invoice) | **823** |
| EIB rows (publish tables) | **823** |
| Rows present in ALL 3 legs (STE) | **823** |
| Amount ok / mismatch | **823 / 0** |
| Fully matched rows (presence + amt) | **823 / 823** |
| Grand total Σ local: Source / Staging / EIB | **32,790,897.87 / 32,790,897.87 / 32,790,897.87** |
| Δ (Source − EIB) | **0.00** |

Verdict: **✅ CLEAN** — we are sending the exact voucher numbers and exact amounts from the raw AP TB into the Workday EIB.

### File locations (same primary → stage → fallback + recovery pattern)

1. **PRIMARY** — `~/Documents/__data/excel/workday/reconcile/FA_ThreeWay_SupplierInvoice_YYYYMMDD[_LABEL].xlsx`. Succeeds when you run the script directly from your own shell. Fails inside TRAE sandbox.
2. **STAGE** — `/tmp/FA_ThreeWay_SupplierInvoice_YYYYMMDD[_LABEL].xlsx`. Always works.
3. **FALLBACK** — `elt_ingest_excel/scripts/_output/FA_ThreeWay_SupplierInvoice_YYYYMMDD[_LABEL].xlsx`. Always works.

If primary was blocked by TRAE sandbox wrapper: re-run once from your own Terminal (not inside TRAE IDE), or manually copy from `/tmp/` or `scripts/_output/`.

---

## Post-load 4th-layer tieback (Sent EIB ↔ Workday Aging extract) — detail

**Script:** [scripts/generate_supplier_invoice_post_load_reconcile.py](../scripts/generate_supplier_invoice_post_load_reconcile.py)
**Analogue:** [generate_post_load_tieback.py](../scripts/generate_post_load_tieback.py) (PO/receipt post-load tieback — same architecture, conventions, CLI shape).

Compares the 823 rows we **sent** in the two EIBs against what Workday actually **loaded**, as evidenced by the Workday Operational Aging Detail by Due Date export processed by 3+.

### Run standalone

```bash
cd elt_ingest_excel
uv run python scripts/generate_supplier_invoice_post_load_reconcile.py --label PRODUCTION
```

Optional flags:
- `--workday-extract <path>` — override the PRODUCTION Aging extract (defaults to `~/Documents/__data/excel/workday/extract/ap_invoices/Operational_Aging_Detail_by_Due_Date (1).xlsx`).
- `--sheet <name>` — specific sheet name; default = auto-resolve to the first sheet (which is `Payables Aging Detail`).
- `--label <TEXT>` — tag appended to output filename so runs don't clobber; use `PRODUCTION` for sign-off runs.

### Reconcile layers (7 grains, strongest → weakest)

Output is an 8-sheet `.xlsx`, identical sheet-naming and style convention to the PO/receipt post-load tieback (TITLE banner row-1 + HEAD_FILL header row-3 + freeze `A4`, ExcelWriter with `startrow=2` pattern from three_way.py):

1. **1-Exec Summary** — verdict + headline counts + status cell coloured GREEN or RED.
2. **2-By Txn Type** — INVOICE (761 rows, ~£34.34m) + CREDIT_NOTE (62 rows, ~£-1.55m) both sides Δ=0.
3. **3-By Company Currency** — per (FA / WNSL / NFC × ccy) rows and amounts.
4. **4-Currency Coverage** — Distinct transaction currencies SENT vs LOADED (GBP, USD, EUR, CHF, IDR present BOTH sides).
5. **5-Credit Note Exact** — 62 rows, penny-exact merge on `column_k` sent == `Transaction` loaded reference (FAP/FAV/FACRN IDs). This is the **authoritative per-document proof** — Workday document IDs are shared keys.
6. **6-Reconcile (Supplier)** — (Company × normalised-Supplier-Name × CCY) buckets. Status: Match / Sent only / Loaded only / Review. Supplier names are symmetrically normalised on BOTH sides before grouping (strip punctuation, drop Ltd/Limited/PLC/Group/Services/GmbH/SL etc entity suffixes, &↔AND, collapse whitespace, UPPERCASE). The `COMPANY_NAME_NORMALISATION` dict in the script also maps Workday-BE labels ('THE FA GROUP - FA') → user-facing short labels ('Football Association') on the sent side.
7. **7-Reconcile (Invoice)** — Greedy per-invoice match on (Company × normalised Supplier × Currency × Amount±0.005). Used for traceability only: the **loaded Aging report stores Workday in-tenant refs (FA-SINV-NNNNNN)** while the **sent EIB stores source-system voucher numbers (FAINVYYYYNNNNN)**. They are disjoint ID systems, so this layer CANNOT produce a per-document shared-key match — it's a best-effort greedy trace. The authoritative layers are Grand total, Credit-note exact-key (62/62), and Supplier bucket (364 Match).
8. **8-Notes** — full methodology identical in structure to the PO tieback notes tab.

### File locations (stage-copy pattern)

The TRAE sandbox blocks writes under `~/Documents/__data/…` (reading works fine). The script writes to 3 locations (primary **first**, then stage + fallback), then attempts a recovery shutil.copyfile if primary was blocked:

1. **PRIMARY** — `~/Documents/__data/excel/workday/reconcile/FA_PostLoad_Tieback_SupplierInvoice_YYYYMMDD[_LABEL].xlsx`. Succeeds when you run the script directly from a user-owned shell (your own Terminal). Fails inside TRAE.
2. **STAGE** — `/tmp/FA_PostLoad_Tieback_SupplierInvoice_YYYYMMDD[_LABEL].xlsx`. Always works.
3. **FALLBACK** — `elt_ingest_excel/scripts/_output/FA_PostLoad_Tieback_SupplierInvoice_YYYYMMDD[_LABEL].xlsx`. Sandbox permits writing inside the project tree; always works.

All three copies are the identical workbook. To get the file into the canonical reconcile location if the TRAE sandbox blocked the primary write: re-run the script once outside TRAE in your own shell, or manually copy from `scripts/_output/` or `/tmp/`.

---

## Known invariants / quirks / gotchas

- **Date reconcile is invoice-date only, never due-date.** We populate one date — `nrm_document_date` → EIB `column_af` (invoice) / `column_y` (credit note) — from GP `invoice_date`. We do **not** send a due date; Workday derives it from payment terms. So the reconcile compares sent invoice-date vs Workday `Transaction Date` (tabs 4B/7B/5B), and GP↔EIB invoice-date is 100% (761/761). Any GP-vs-Workday **Due Date** difference a consultant raises is a payment-terms reconcile, not our load — do not chase it here. As at 12.08.2026 the only genuine invoice-date defect is the 4 D R IMAGES rows (`FAV0114322`–`5`): GP typo `04/12/2026` (should be `04/12/2025`) that Workday additionally stored day/month-transposed as `12/04/2026`.
- **Date-parity layers post-date the amount sign-off.** The 4B/7B/5B date tabs were added in `cf3b16b`, after the 10.08.2026 CLEAN was recorded (`e97129f`). A CLEAN amount sign-off does **not** imply date parity was checked — re-run the post-load script to exercise the date verdict.
- **Two distinct document-ID systems on invoices.** Never try to merge voucher numbers (source system `FAINVYYYYNNNNN` in sent EIB `column_bg`) to Workday FA-SINV refs (tenant refs in loaded Aging report) — they are disjoint by design. Any per-document equality is coincidence. Use the greedy (company × supplier × ccy × amount) layer as the invoice-level trace; use credit-note exact-key + grand total + supplier buckets as the proof.
- **Credit-note column letters differ from invoice.** `column_k/p/r/s/y/ap/eq` (credit-note template v46) vs `column_r/t/u/af/as/gf` (invoice template v45.2). The two publish SQL files and the reconcile script each encode their column offsets explicitly. **Never copy a column-block from one publish SQL to the other without re-checking every letter against the EIB template CSV definition.**
- **Currency column trailing spaces in the EIB are accepted.** `workday_supplier_invoice.column_t` carries `'GBP            '` (16 chars). The reconcile SQL reads via `TRIM()` for comparisons; Workday's own loader trims the same way when loading. Do NOT add a `TRIM()` in the publish SQL itself — we verified Workday reads it fine as-is, and changing it would break parity with the already-loaded PRODUCTION workbook.
- **JPY is whole-integer VARCHAR on both sides.** Verified in nrm_transaction_amt; same logic as PO/receipt EIBs. No penny-drift possible because JPY has no subunit.
- **No "disabled tables" in this module.** Both publish tables (invoice + credit) are fully wired and both run for every pipeline execution. There is no `WHERE 1=2` disabled-tab analogue to `workday_customer_settlement_account` from the customer module.
- **Unresolved supplier mapping = pipeline silent drop.** The Stage 3 gate (`wd_supplier_id IS NOT NULL`) drops rows without raising an exception. They only surface via `validate_supplier_invoice_supplier.sql`. The post-load tieback won't include them in SENT because they never reach the publish tables. If you ever see Loaded > Sent at the supplier-bucket layer in the reconcile report, check this table first — the mismatch is almost certainly unresolved supplier IDs in the nrm layer (GP/finance haven't added the vendor mapping to the supplier master yet).
- **Company label normalisation in the reconcile script is sent-only.** The loaded Aging report already uses user-facing short labels ('Football Association', 'Wembley National Stadium', 'National Football Centre'). Only the sent side (EIB) carries the Workday-BE 'THE FA GROUP - …' labels that we remap inside `COMPANY_NAME_NORMALISATION` before comparing.
- **Sandbox limitations (TRAE IDE).** Writes under `/Users/Rakesh.Patel/Documents/__data/` are blocked by the TRAE sandbox wrapper even though file permissions are `755` and the directory contains valid output files written from normal shells. Reading from the same path works fine (this is why DuckDB read-only connections + reading the Aging extract work inside TRAE). From TRAE you can run everything up to read-only recon; to run TRANSFORM/PUBLISH or to land the reconcile workbook in the canonical directory, run the commands from your own Terminal shell outside TRAE.
