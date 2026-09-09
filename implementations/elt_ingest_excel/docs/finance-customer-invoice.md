# Finance — Customer Invoice / Credit Note → Workday

**Single source of truth for the customer-invoice + credit-note (adjustment) migration.**
Cold-start sessions should be able to continue from this file alone.
**Read §Ownership & Scope and §Known invariants before touching any SQL or the reconcile scripts.**

Config lives under `config/{ingest,transform/sql,validate,reconcile,publish}/finance/customer_invoice/`.
Run commands are in [CLI.md](CLI.md) (not duplicated here).

This document is the customer-invoice analogue of [finance-supplier-invoice.md](finance-supplier-invoice.md) (Supplier Invoices). The customer **master/debtor** pipeline is documented separately in [finance-customer.md](finance-customer.md) — that doc is about creating `C-` / `workday_customer_*` Workday customer records; **this** doc is about loading the outstanding AR (invoices and credit notes) that sit *against* those customer records.

The master customer (`workday_customer_*` tables) **must already be loaded** before this pipeline runs, because invoice rows gate on `wd_customer_id IS NOT NULL` (exact-then-bare customer match against the Workday customer map projection built from the position-2 extract).

---

## Ownership & Scope

- Module: `customer_invoice`
- SQL folder: `elt_ingest_excel/config/transform/sql/finance/customer_invoice/`
- Ingest config: [`config/ingest/finance/customer_invoice.json`](../config/ingest/finance/customer_invoice.json)
  - Primary source: `AR DTL TB AUG 5 2026.xlsx` / three BU-specific sheets (`FA AR DTL TB`, `NFC AR DTL TB`, `WNSL AR DTL TB`) → `fin_customer_invoice_fa`, `fin_customer_invoice_nfc`, `fin_customer_invoice_wnsl`
  - Secondary ingest (position 2, **NOT published** — used to build the customer map only): Workday Customer Extract → `fin_workday_extract_customer` (shared with customer master pipeline)
- Publish config: [`config/publish/finance/publish_customer_invoice.json`](../config/publish/finance/publish_customer_invoice.json) (two-workbook publish: Invoice + Adjustment)
- Verification DB: `~/Documents/__data/duckdb/rpatel.duckdb`
  - Read-only inside the TRAE sandbox. TRANSFORM/PUBLISH must be run from a user-owned shell (see §Sandbox limitations).
- Loaded Workday authority (target for post-load tie-back): TBC — the Workday tenant export post customer-invoice load (e.g. `AR Aging Detail by Due Date.xlsx`) — place under `~/Documents/__data/excel/workday/extract/ar_invoices/` and pass path via `--workday-extract` to the post-load tieback script.

---

## Status (start here)

> ⚠ **ONE-OFF MIGRATION (not BAU):** Customer invoices / credit memos are loaded ONCE into Workday at go-live as opening AR balances (line types `AR Conversion` / `Open AR Conversion` / `Open AR Adjustment Conversion`). Source data was statically extracted 05.08.26 and will NOT be refreshed; the EIB workbooks produced are signed-off — do NOT re-run unless a new AR extract is supplied.

**EIBs produced:** Two EIB workbooks from a single staged source (`src_fin_customer_invoice`):

| Publish table | Workbook template | Output file name prefix | Target sheet | dataRow |
|---|---|---|---|---|
| `workday_customer_invoice` | `Submit_Customer_Invoice_v46.1.xlsx` | `workday_fin_customer_invoice_v1` | Submit Customer Invoice | 6 |
| `workday_customer_credit_note` | `DC_Customer_Invoice_Adjustment_v46.1.xlsx` | `workday_fin_customer_invoice_adjustment_v1` | Submit Customer Invoice Adjustm | 6 |

### Run the pipeline

```bash
cd elt_ingest_excel
uv run python examples/fin_customer_invoice.py --run-to-phase PUBLISH
```

TRAE sandbox blocks DuckDB write-open: inside TRAE run with `--run-to-phase INGEST` at most; for TRANSFORM and above run the command in your own Terminal shell.

### Reconcile workbooks — two layers, same as supplier pipeline

| Reconcile | Script | What it proves |
|---|---|---|
| **Upstream three-way** | [generate_customer_invoice_three_way.py](../scripts/generate_customer_invoice_three_way.py) | Raw Source AR TB ↔ Our final Staging ↔ Our EIB publish tables — proves we are sending the correct document numbers and amounts in the first place. |
| **Post-load 4th-layer tieback** | [generate_customer_invoice_post_load_reconcile.py](../scripts/generate_customer_invoice_post_load_reconcile.py) | Our EIB publish tables ↔ Workday-loaded AR Aging extract — proves Workday loaded exactly what we sent. |

```bash
cd elt_ingest_excel
# Upstream proof (we are sending correct data — run NOW, no Workday extract needed)
uv run python scripts/generate_customer_invoice_three_way.py --label PRODUCTION
# Downstream proof (Workday loaded what we sent — run AFTER Workday load + extract received)
uv run python scripts/generate_customer_invoice_post_load_reconcile.py --label PRODUCTION --workday-extract /path/to/AR_Aging_Detail.xlsx
```

Both write to 3 locations with the same pattern (see §File locations):
- **PRIMARY (always success from your shell):** `~/Documents/__data/excel/workday/reconcile/`
- **STAGE (always):** `/tmp/`
- **FALLBACK (inside TRAE sandbox):** `elt_ingest_excel/scripts/_output/`

TRAE sandbox blocks writes under `__data/` (reading works fine). Re-run both commands from your own Terminal to land the workbooks in the canonical reconcile dir.

---

## Source Excel files (INGEST phase)

### AR Detail Trial Balance (source workbook: `AR DTL TB AUG 5 2026.xlsx`)

Path: `~/Documents/__data/excel/finance_ref/customer_invoice/`

| Sheet | DuckDB table | Business-unit filter embedded |
|-------|--------------|-------------------------------|
| FA AR DTL TB | `fin_customer_invoice_fa` | FA (The FA Group - FA) |
| NFC AR DTL TB | `fin_customer_invoice_nfc` | NFC (National Football Centre) |
| WNSL AR DTL TB | `fin_customer_invoice_wnsl` | WNSL (Wembley National Stadium) |

All three sheets share an identical column layout; the pipeline UNIONs them in `src_fin_customer_invoice_base.sql` and re-derives `business_unit` from the originating table name. headerRow=1, dataRow=2.

### Workday customer extract (position 2 in ingest, NOT published, used to build the map only)

| Sheet | DuckDB table |
|-------|--------------|
| Sheet1 | `fin_workday_extract_customer` |

headerRow=11, dataRow=12. Feeds `src_fin_workday_customer_map.sql` (shared pattern with the customer master pipeline). This map is what gates the invoice pipeline — a row missing from the map means `wd_customer_id IS NULL` and the invoice is dropped at Stage 3 + tracked in `validate_customer_invoice_customer.sql`.

---

## Staging table lineage (TRANSFORM phase)

SQL files execute in the order defined in [config/transform/sql/finance/customer_invoice/order.txt](../config/transform/sql/finance/customer_invoice/order.txt) (9 files):

1. `src_fin_workday_customer_map.sql`
2. `src_fin_customer_invoice_base.sql`
3. `src_fin_customer_invoice_nrm.sql`
4. `src_fin_customer_invoice.sql`       — final staged (gate applied)
5. `workday_customer_invoice.sql`       — INVOICE type → EIB
6. `workday_customer_credit_note.sql`   — non-INVOICE type → Adjustment EIB
7. `reconcile_customer_invoice_unresolved.sql`
8. `reconcile_customer_invoice_unresolved_workday.sql`
9. `reconcile_customer_invoice_unresolved_workday_update.sql`

### Stage 1 — `src_fin_customer_invoice_base`

**File:** `src_fin_customer_invoice_base.sql`

UNIONs the three BU-specific raw tables. Adds:
- `business_unit` literal derived from source table (`FA` / `NFC` / `WNSL`)
- `customer_id` = `CAST(customer_number AS VARCHAR)` (scalar alias used downstream for map lookups)
- Semantic aliases for amounts: `document_amount` → `current_transaction_amount_gbp` (GBP leg), `local_currency_document_amount` → `current_transaction_amount_in_originating_currncy` (local-leg, same column naming as supplier pipeline for symmetry)
- `exchange_rate` = NULL (customer AR extract does not carry it, unused downstream)

Filter: `NULLIF(UPPER(TRIM(customer_id)), '') IS NOT NULL` — drops blank customer rows before the nrm layer (mirrors supplier `vendor_id` gate in base.sql).

No normalisation or further filtering.

### Stage 2 — `src_fin_customer_invoice_nrm`

**File:** `src_fin_customer_invoice_nrm.sql`

Adds four key derived columns via the exact-then-bare Workday customer map:

| Column | Derivation | Notes |
|--------|-----------|-------|
| `wd_customer_id` | `COALESCE(c_exact.customer_id, c_bare.customer_id)` — exact `customer_id-BU` first, bare `customer_id` fallback | Same pattern as PO commitment stack; customer map built from position-2 Workday customer extract. |
| `wd_customer_name` / `wd_customer_reference_id` | Same `COALESCE(exact, bare)` pattern — used for traceability in reconcile reports | |
| `nrm_document_date` | Triple-format `TRY_STRPTIME` over the raw text `document_date` (`%Y-%m-%d %H:%M:%S` / `%Y-%m-%d` / `%d-%m-%Y`), reformatted `dd/mm/yyyy` for EIB | Invoices and credit notes share one format — Workday invoice_date column = adjustment_date column in the credit-note EIB. |
| `nrm_company_id` | Explicit FA → `'THE FA GROUP - FA'`, NFC → `'THE FA GROUP - NFC'`, WNSL → `'THE FA GROUP - WNSL'` mapping (matches what Workday's Business Entity picklist stores). | These exact strings are what the Workday EIB templates require in `column_k`/`column_l`. |
| `nrm_transaction_amt` | Currency-branched **VARCHAR pass-through**: GBP = raw `document_amount` as VARCHAR; JPY = `ROUND(CAST(src AS DOUBLE), 0)` as whole-integer VARCHAR; else = raw local amount as VARCHAR. | See §Known invariants for the full rationale: the customer data is pure (verified 0 dirty-strings / 1,447 rows); this is a deliberate one-time signed-off choice, NOT a BAU pattern. Align to the supplier DOUBLE-cast pattern if the migration is ever re-extracted (code block with comparison numbers is embedded in the SQL). |

**Customers that don't resolve are not dropped yet** — they flow through with NULL `wd_customer_id` so the gate in Stage 3 is auditable (see validation).

### Stage 3 — `src_fin_customer_invoice` *(final invoice/credit-note staged table)*

**File:** `src_fin_customer_invoice.sql`

Gate: `wd_customer_id IS NOT NULL AND CAST(nrm_transaction_amt AS DOUBLE) <> 0`.

Note the explicit `CAST(... AS DOUBLE)` on the gate — this is required because `nrm_transaction_amt` is stored as VARCHAR in the customer pipeline (compare with supplier pipeline gate on a DOUBLE column directly). The SQL-side cost of this extra per-row cast is negligible on a ~1,400 row one-off dataset.

This is the **single source of truth** for both EIB publish tables. Splits happen downstream — not here.

### Stage 4 — Publish tables (split by document type)

**Invoice EIB publish:** [workday_customer_invoice.sql](../config/transform/sql/finance/customer_invoice/workday_customer_invoice.sql)
- `WHERE UPPER(TRIM(document_type)) LIKE '%INVOICE%'` (note LIKE, not exact match — handles AR-source variants like `INVOICE`, `Sales Invoice`, etc.)
- `column_h` (customer_invoice_number / legacy GP) = `document_number` — the GP Invoice Number (3+ flagged missing before v1.1 fix)
- `column_k` (company) = `nrm_company_id` (Workday-BE labels, 3 values)
- `column_cp` (currency) = raw `currency_id` (no trim in SQL; same trailing-space convention as supplier — Workday loader trims on read)
- `column_cq` (customer) = `wd_customer_id`
- `column_gf` (invoice_date) = `nrm_document_date` (%d/%m/%Y text)
- `column_ha` (amount_due) + `column_mj` (extended_amount) both = `nrm_transaction_amt` (1 invoice line = 1 document line; no multi-line splits)
- `column_ke` = `'AR Conversion'` revenue_category
- `column_kf` = `'Open AR Conversion'` line-item description
- `column_mt` = `'Open AR Conversion'` memo
- `ROW_NUMBER()` over (BU, wd_customer_id, nrm_document_date, document_number) → `column_b` spreadsheet_key. One line per doc = `'1'` always on line row-IDs.

**Credit Note / Adjustment publish:** [workday_customer_credit_note.sql](../config/transform/sql/finance/customer_invoice/workday_customer_credit_note.sql)
- `WHERE UPPER(TRIM(document_type)) NOT LIKE '%INVOICE%'`
- **Column letters differ from the invoice EIB** because credit-note template v46.1 rearranges: `column_h` = document_number (customer_invoice_number / legacy GP — same field as invoice EIB, added 12/08/26 for 3+ parity), `column_l` = company, `column_aj/column_ak` = customer_id + sold_to, `column_bg` = Conversion adjustment_reason, `column_bh` = `'N'` (not an increase = it reduces balance), `column_bk` = currency, `column_bl` = invoice_date, `column_cq` = memo, `column_fb` = extended_amount, `column_fc` = document_number (line_memo). The two publish SQL files and both reconcile scripts encode their offsets explicitly. **Never copy a column-block from one publish SQL to the other without re-checking every letter against the EIB template CSV definition.**
- `column_h` (customer_invoice_number) and `column_fc` (line_memo) both carry the GP reference on credit notes. The post-load tieback's Layer-5 exact-merge uses `column_fc` when the loaded AR extract exposes a source-doc-ref column. If the loaded extract only carries Workday tenant refs (no source doc ref), the CN layer degrades to the same "exact ref missing" pattern as the invoice layer — the authoritative proof then shifts to Grand total + Customer bucket.

---

## Validation

Runs automatically as part of the pipeline for phases ≥ TRANSFORM, or standalone:

```bash
uv run python examples/fin_customer_invoice.py --run-to-phase TRANSFORM
```

Validation files in `config/validate/finance/customer_invoice/`:

| File | Purpose |
|---|---|
| `init.sql` | Creates / re-seeds `validation_customer_invoice_issues` (same schema pattern as supplier: table_name, rule_type, col, invoice_id (=wd_customer_id), detail, message — here `invoice_id` column stores the `wd_customer_id` reference the failing row linked to). |
| `validate_customer_invoice_customer.sql` | Every row in `src_fin_customer_invoice_nrm` with `wd_customer_id IS NULL` — dropped at Stage 3 gate. These are customer codes that GP/finance haven't mapped in the customer master yet. |
| `validate_customer_invoice_missing_payment_term.sql` | Rows where the related Workday customer has no resolvable payment terms — downstream EIB load rejects invoices linked to customers missing default payment terms even if the invoice row itself is clean. |

Query after running:
```sql
SELECT table_name, rule_type, COUNT(*) n
  FROM validation_customer_invoice_issues
 GROUP BY 1, 2 ORDER BY 1, 2;
```

---

## Upstream 3-way tie-back (Source AR TB ↔ Our Staging ↔ Our EIB publish)

**Script:** [scripts/generate_customer_invoice_three_way.py](../scripts/generate_customer_invoice_three_way.py)
**Analogue:** [generate_supplier_invoice_three_way.py](../scripts/generate_supplier_invoice_three_way.py) (Supplier three-way — same 8-sheet layout, banner + header styling, verdict cell colouring, primary/fallback write paths, tolerance 0.005).

Proves we are loading the **correct numbers from the raw source** into the EIB we send to Workday. This is the upstream proof that our base/nrm/staging transforms haven't dropped, added, or mangled any documents.

### Join key: document_number (clean, 100% shared, unique per row)

Verified on the PRODUCTION Aug-11 dataset:
- Raw 3-BU ingest (FA + NFC + WNSL AR Detail TB): 495 + 114 + 840 = 1,449 distinct document_numbers before UNION
- `src_fin_customer_invoice_base` (3-BU UNION DISTINCT + alias + WHERE customer_id NOT NULL): **1,447 distinct** (2 duplicates collapsed by UNION)
- `src_fin_customer_invoice_nrm` (nrm layer + typecasts + exact-then-bare customer map): 1,447 distinct
- `src_fin_customer_invoice` (gate applied: wd_customer_id IS NOT NULL AND nrm_transaction_amt <> 0): **1,447 rows / 1,447 distinct** — no rows dropped at gate
- EIB total rows: `workday_customer_invoice` (1,193 INVOICE rows) + `workday_customer_credit_note` (254 CREDIT_NOTE rows) = **1,447 = gate-survivor count**

Grand totals reconciled on every leg (Σ local amount, tolerance 0.005): **£125,649,286.27** = Source 125,649,286.27 = Staging 125,649,286.27 = EIB 125,649,286.27, Δ(Src-EIB) = **0.00**, VERDICT = **CLEAN**.

#### Column-H presence enforcement (12/08/26 3+-parity fix)

Before the 3+ flag, the reconciler joined INVOICE EIB rows back to staging via an indirect `ROW_NUMBER()` over (BU × wd_customer_id × date × amount) partition — which proved row-for-row *presence* via coincidence of grouping but did NOT prove the printed invoice carried the GP number. This has been replaced:

1. **Invoice EIB read:** three_way.py now reads `workday_customer_invoice.column_h` as the authoritative `document_number`. If any EIB row carries a blank `column_h`, the loader raises a runtime warning.
2. **Credit-note EIB read:** reads `COALESCE(TRIM(column_h), TRIM(column_fc))` — `column_h` is the new authoritative slot; `column_fc` (Line Memo) is retained as the pre-fix fallback so pre-existing reconcile runs still succeed against older EIB batches.
3. **No more indirect joins.** A blank/missing `document_number` on the EIB leg now causes that document to appear in Drift & Exceptions (tab 7) with a `status_presence` that no longer includes `E`, and the overall verdict flips from CLEAN → REVIEW.
4. **Post-load tieback:** same pattern — `workday_customer_invoice.column_h` now drives the `gp_document_number` and `eib_txn_ref` fields; Layer-5 Credit-note exact-key still prefers exact match when the loaded AR Aging export carries a source-ref column (fmt_B `External Document Number`, fmt_A `Document Number` / `Source Document`); Layer-7 invoice traceability now routes any shared GP invoice number exact match before degrading to the 4-tuple.

Document numbers travel unchanged across all 5 layers, and are now explicitly written into `column_h` of the publish tables (confirmed by the reconciler reading the cell value back out). No fuzzy matching required at either the three-way or post-load layer.

Amount compare on `local_currency_document_amount` (raw) = `nrm_transaction_amt` (staging nrm, DOUBLE-cast inside the three-way loader to handle the VARCHAR storage cleanly) = `column_ha/column_mj` invoice OR `column_fb` credit note (EIB publish). Verified per-document on all rows: 0 deltas > 0.005 is the expected signed-off result.

### Run standalone

```bash
cd elt_ingest_excel
uv run python scripts/generate_customer_invoice_three_way.py --label PRODUCTION
```

Optional flags:
- `--duckdb <path>` — override default DuckDB at `~/Documents/__data/duckdb/rpatel.duckdb`.
- `--out-dir <dir>` — override primary output dir (defaults to `~/Documents/__data/excel/workday/reconcile`).
- `--label <TEXT>` — tag appended to workbook filename; use `PRODUCTION` for sign-off runs.

### 8 sheets output (mirrors `generate_supplier_invoice_three_way.py`)

1. **0-Overview** — Workbook purpose, 3 legs explained, companion post-load workbook link, chain-of-custody diagram, where to look first, flags + sandbox note. **MANDATORY first tab.**
2. **1-Executive Summary** — 7-row matrix: Row counts, STE presence, amt OK / MISMATCH counts, Grand totals (Source / Staging / EIB + pairwise Δs), **VERDICT cell** coloured GREEN for CLEAN, RED for REVIEW.
3. **2-By Txn Type** — INVOICE vs CREDIT_NOTE, rows + amounts + deltas across all 3 legs.
4. **3-By Company Currency** — (Company × Currency × Txn Type) rows/amounts.
5. **4-Currency Coverage** — every currency present Source/Staging/EIB on "BOTH" status = good.
6. **5-Per-Document Exact** — all documents, Source/Stage/EIB amounts, pairwise Δ columns, `Presence (STE)` key, Amount status, Overall status (MATCH / REVIEW). Overall column cell coloured green/red.
7. **6-Customer Bucket** — (Company × normalised-customer × ccy) groups, rows + amounts across 3 legs, pairwise Δs, Status (Match / Source only / Staging only / EIB only / Review) yellow/green/red filled.
8. **7-Drift & Exceptions** — Every document failing presence OR amount equality. For a clean run the body is a single placeholder row: `"NONE — all 1,447 documents matched perfectly across all 3 legs."`
9. **8-Notes** — Methodology: Scope, Join key, Amount compare, Company norm, Customer-name norm, Txn type derivation (LIKE '%INVOICE%'), Trailing ccys spaces, JPY handling, Gate row-drop audit, Workbook verdict, How to run, TRAE sandbox note, Analogue script.

### File locations (same primary → stage → fallback + recovery pattern)

1. **PRIMARY** — `~/Documents/__data/excel/workday/reconcile/FA_ThreeWay_CustomerInvoice_YYYYMMDD[_LABEL].xlsx`. Succeeds when you run the script directly from your own shell. Fails inside TRAE sandbox.
2. **STAGE** — `/tmp/FA_ThreeWay_CustomerInvoice_YYYYMMDD[_LABEL].xlsx`. Always works.
3. **FALLBACK** — `elt_ingest_excel/scripts/_output/FA_ThreeWay_CustomerInvoice_YYYYMMDD[_LABEL].xlsx`. Always works.

If primary was blocked by TRAE sandbox wrapper: re-run once from your own Terminal (not inside TRAE IDE), or manually copy from `/tmp/` or `scripts/_output/`.

---

## Post-load 4th-layer tieback (Sent EIB ↔ Workday AR Aging extract) — detail

**Script:** [scripts/generate_customer_invoice_post_load_reconcile.py](../scripts/generate_customer_invoice_post_load_reconcile.py)
**Analogue:** [generate_supplier_invoice_post_load_reconcile.py](../scripts/generate_supplier_invoice_post_load_reconcile.py) (Supplier post-load tieback — same architecture, conventions, CLI shape).

Compares the rows we **sent** in the two EIBs against what Workday actually **loaded**, as evidenced by the Workday AR Aging Detail export (the customer-side equivalent of the supplier Operational Aging report).

### Run standalone

```bash
cd elt_ingest_excel
uv run python scripts/generate_customer_invoice_post_load_reconcile.py --label PRODUCTION --workday-extract /path/to/AR_Aging_Detail.xlsx
```

Optional flags:
- `--workday-extract <path>` — the Workday AR Aging xlsx export (no hardcoded default path ships; set on first run).
- `--sheet <name>` — specific sheet name; default = auto-resolve to the first sheet + auto-detect header row (scans rows 0–40 for a Company + Customer + Transaction Type header).
- `--label <TEXT>` — tag appended to output filename so runs don't clobber; use `PRODUCTION` for sign-off runs.

### Reconcile layers (7 grains, strongest → weakest)

Output is an 9-sheet `.xlsx`, identical sheet-naming and style convention to the supplier post-load tieback (TITLE banner row-1 + HEAD_FILL header row-3 + freeze `A4`, ExcelWriter with `startrow=2` pattern from three_way.py):

1. **0-Overview** — What this workbook proves, 2 legs compared, full join strategy, companion three-way workbook link, chain-of-custody diagram, where to look first, flags + sandbox note. **MANDATORY first tab.**
2. **1-Exec Summary** — verdict + headline counts + status cell coloured GREEN or RED.
3. **2-By Txn Type** — INVOICE + CREDIT_NOTE both sides Δ.
4. **3-By Company Currency** — per (FA / WNSL / NFC × ccy) rows and amounts.
5. **4-Currency Coverage** — Distinct transaction currencies SENT vs LOADED (all present BOTH sides required).
6. **5-Credit Note Exact** — CN rows, merge on sent `eib_txn_ref` (= column_fc = document_number) == loaded `src_doc_ref` if the extract carries it (falls back to `wd_txn_ref` if no source-ref column in the extract). **Strongest per-document proof IF the extract exposes a source-ref column;** if it carries only Workday tenant-refs this layer degrades and you rely on Grand total + Customer bucket + Invoice fuzzy.
7. **6-Reconcile (Customer)** — (Company × normalised-Customer-Name × CCY) buckets. Status: Match / Sent only / Loaded only / Review. Customer names are symmetrically normalised on BOTH sides before grouping (strip punctuation, drop Ltd/Limited/PLC/Group/Services/GmbH/SL etc entity suffixes, &↔AND, collapse whitespace, UPPERCASE). The `COMPANY_NAME_NORMALISATION` dict maps Workday-BE labels ('THE FA GROUP - FA') → user-facing short labels ('Football Association') on the sent side.
8. **7-Reconcile (Invoice)** — Greedy per-invoice match on (Company × normalised Customer × Currency × Amount±0.005). Used for traceability only: loaded AR reports Workday in-tenant refs (FA-CINV-NNNNNN) while sent EIB stores source-system document_number. They are disjoint ID systems, so this layer CANNOT produce a per-document shared-key match — it's a best-effort greedy trace. The authoritative layers are Grand total, Credit-note exact-key (if source-ref column available in extract), and Customer bucket.
9. **8-Notes** — full methodology identical in structure to the supplier tieback notes tab.

### File locations (stage-copy pattern)

1. **PRIMARY** — `~/Documents/__data/excel/workday/reconcile/FA_PostLoad_Tieback_CustomerInvoice_YYYYMMDD[_LABEL].xlsx`. Succeeds from user shell; fails inside TRAE sandbox.
2. **STAGE** — `/tmp/FA_PostLoad_Tieback_CustomerInvoice_YYYYMMDD[_LABEL].xlsx`. Always works.
3. **FALLBACK** — `elt_ingest_excel/scripts/_output/FA_PostLoad_Tieback_CustomerInvoice_YYYYMMDD[_LABEL].xlsx`. Always works.

If TRAE sandbox blocked primary write, re-run outside TRAE or copy from fallback/stage to reconcile dir.

---

## Known invariants / quirks / gotchas

- **nrm_transaction_amt is VARCHAR today (not DOUBLE like the supplier pipeline).** This is a deliberate signed-off choice on a one-off pure dataset (1,447 rows, 0 irregular strings, GBP 1,439 + USD 8 only = no edge branches). The full decision rationale + pairwise comparison numbers + alignment block are embedded as comments in [src_fin_customer_invoice_nrm.sql](../config/transform/sql/finance/customer_invoice/src_fin_customer_invoice_nrm.sql). **Do NOT change this back to DOUBLE without re-running the full customer pipeline from TRANSFORM through PUBLISH + full re-sign-off of both three-way and (if loaded) post-load tieback** — if the AR migration is ever re-extracted, use the verbatim alignment block already embedded in the nrm SQL.
- **Gate at Stage 3 re-casts to DOUBLE:** `AND COALESCE(CAST(t.nrm_transaction_amt AS DOUBLE), 0) <> 0`. This is the cost of keeping nrm as VARCHAR. Negligible on 1,447 rows × every validation run. If you align the nrm column to DOUBLE, simplify this gate to `t.nrm_transaction_amt <> 0` (same as supplier).
- **Two distinct document-ID systems on invoices (same issue as supplier):** Never try to merge source `document_number` (sent EIB) to Workday tenant refs (loaded AR Aging FA-CINV-*). They are disjoint by design. Any per-document equality is coincidence. The authoritative proof layers are: Grand total, Credit-note exact-key (if source ref column in the loaded extract), Customer bucket. The invoice greedy-match is trace-only.
- **Credit-note column letters differ from invoice.** `column_l/aj/ak/bg/bh/bk/bl/cq/fb/fc` (credit-note template v46.1) vs `column_k/cp/cq/gf/gy/gz/ha/hc/ih/ke/kf/mj/mt` (invoice template v46.1). The two publish SQL files and both reconcile scripts encode column letters explicitly. **Never copy-paste a publish column block between the two.**
- **Currency column trailing spaces in the EIB are accepted.** Same convention as supplier pipeline — `workday_customer_invoice.column_cp` carries untrimmed `currency_id` raw. Both reconcile scripts read via `TRIM()` for comparisons; Workday's own loader trims the same way when loading. Do NOT add a `TRIM()` in the publish SQL itself.
- **JPY branch exists in nrm but is dead code today.** 0 JPY rows in the PRODUCTION Aug-05 extract. It still works if ever re-extracted (verified logic: `ROUND(CAST(src AS DOUBLE), 0)` → VARCHAR whole integer; same as supplier JPY handling).
- **No "disabled tables" in this module.** Both publish tables (invoice + credit) are fully wired and both run for every pipeline execution. There is no `WHERE 1=2` disabled-tab analogue to customer master's `workday_customer_settlement_account`.
- **Unresolved customer mapping = pipeline silent drop.** The Stage 3 gate (`wd_customer_id IS NOT NULL`) drops rows without raising an exception. They only surface via `validate_customer_invoice_customer.sql`. The post-load tieback won't include them in SENT because they never reach the publish tables. If you ever see Loaded > Sent at the customer-bucket layer in the reconcile report, check this validation table first — the mismatch is almost certainly unresolved customer IDs in the nrm layer (GP/finance haven't added the mapping to the customer master yet).
- **Company label normalisation in both reconcile scripts is sent-only.** The loaded AR Aging report already uses user-facing short labels ('Football Association', 'Wembley National Stadium', 'National Football Centre'). Only the sent side (EIB) carries the Workday-BE 'THE FA GROUP - …' labels that we remap inside `COMPANY_NAME_NORMALISATION` before comparing.
- **Sandbox limitations (TRAE IDE).** Writes under `/Users/Rakesh.Patel/Documents/__data/…` are blocked by the TRAE sandbox wrapper even though file permissions are `755` and the directory contains valid output files written from normal shells. Reading from the same path works fine (this is why DuckDB read-only connections + reading the Aging extract work inside TRAE). From TRAE you can run everything up to read-only recon; to run TRANSFORM/PUBLISH or to land the reconcile workbook in the canonical directory, run the commands from your own Terminal shell outside TRAE.
