# Claude review — `ar_customer` Item 6 (drop-downs), 2026-07-22

Reviewed against the historical checkpoint notes now preserved in `docs/todo/archive/ar-ingest-specs-history.md`,
which claimed Item 6 was
finished ("Remaining: 0"). Re-ran the render and the doc's own verification commands rather than
trusting the checkpoint notes. This review originally found two issues. Both have now been fixed.
See `ar-ingest-specs-history.md` § Item 6 for the working loop / render command / definition of done —
this doc is retained as a review record.

## Finding 1 (resolved) — two lookup tabs exceeded Excel's 31-character sheet-name limit

`config/ingest_lookups/ref_document_email_recipient_invoice.json` (36 chars) and
`ref_document_email_recipient_statement.json` (38 chars) are used verbatim as the worksheet title in
`write_lookup_sheet()` (`scripts/create_ingest_template.py:296`). Excel's hard limit is 31 characters.

**Reproduce:**

```bash
uv run --with openpyxl --with pyyaml --no-project python \
  elt_transform_excel/scripts/create_ingest_template.py ar_customer --overwrite
```

prints, to stderr:

```
UserWarning: Title is more than 31 characters. Some applications may not be able to read the file
```

Confirmed it's not just an openpyxl warning — the produced `.xlsx`'s `xl/workbook.xml` literally
contains `<sheet name="ref_document_email_recipient_statement" .../>`. Real Excel will hit a repair
prompt on open and typically renames/truncates the offending sheet, which silently breaks the
data-validation formulas that reference it by full name (`'ref_document_email_recipient_invoice'!
$A$2:$A$9` and `'ref_document_email_recipient_statement'!$A$2:$A$8`, backing the Invoice/Statement
Notification Email Recipients drop-downs, columns AO and AQ).

**Fix:** rename both lookups to ≤31 chars, e.g. `ref_email_recipient_invoice` (27 chars) /
`ref_email_recipient_statement` (29 chars) — keeps the `ref_` convention, drops the redundant
`document_`. Rename must touch three places:
1. `config/ingest_lookups/ref_document_email_recipient_invoice.json` →
   `config/ingest_lookups/ref_email_recipient_invoice.json` (and its `_statement` counterpart), and
   update each file's own `key_column` if it echoes the old name.
2. `config/ingest_specs/ar_customer.yaml` — the `lookups:` list entries and the two
   `picklist: {lookup: ...}` references (`Invoice Notification Email Recipients`,
   `Statement Notification Email Recipients`).
3. Re-render and confirm the `UserWarning` is gone and `xl/workbook.xml` has no sheet name over 31
   chars (`unzip -p dist/ingest_templates/ar_customer.xlsx xl/workbook.xml | grep -o 'name="[^"]*"'`
   and check lengths).

**Resolved:** the renderer now uses deterministic Excel-safe lookup sheet titles and updates all
validation / defined-name references to point at those shortened titles.

## Finding 2 (resolved) — `Address Primary` had no picklist

`ar-ingest-specs-history.md`'s own Group B table lists `Address Primary` as item **#19**, a Yes/No
drop-down ("always `Yes` — one address per customer"). It never got one:
`config/ingest_specs/ar_customer.yaml:575-581` is `type: text` with a `default: "Yes"` and no
`picklist:` key, unlike its Group B neighbours `Default Currency` (#18) and `Transaction Tax ID`
(#20), which both have theirs.

Confirmed by the render's own summary line — it reports **26** drop-down columns, not the 27 the
doc's own progress note predicted for "groups A and B... with cascading."

**Resolved:** `Address Primary` now uses `picklist: {lookup: ref_yes_no, column: yes_no}`, matching
the other controlled Yes/No fields. Re-render now reports **27** drop-downs.

## Not blocking — flag, don't fix without operator input

- **`ref_country_tax_id_type.json`'s India (`IN`) row** differs from the live EIB named range by a
  trailing non-breaking space the source workbook cell carries (`'TAN Tax Deduction and Collection
  Account\xa0'` in the workbook vs the clean string in our snapshot). Looks like a deliberate
  whitespace cleanup, not a transcription error — but it means the "match the named range exactly,
  both directions" verification bar in § Item 6's definition of done doesn't literally pass for this
  one row. Leave as is unless the operator wants byte-for-byte fidelity including the workbook's own
  typos.
- **`Invoice Delivery Method` / `Statement Delivery Method` allowed values (`Email`, `Print`)** come
  from the migration SQL per the 2026-07-22 checkpoint, which is a documented, deliberate choice. But
  the live reference workbook's own `Invoice_Delivery_Method` named range currently resolves to
  `{Email, Mail}`, not `{Email, Print}` — unlike every other picklist in this spec, this one wasn't
  cross-checked against the EIB's own live range (the doc's own Gotchas section calls this pattern
  out repeatedly: "Always read the EIB's named range"). The file that would settle it definitively —
  `~/Downloads/workday_fin_debtor_customer_active_v1.xlsm`, the actual accepted E2E upload — is no
  longer present to check against. Recorded here as an open verification gap, not a confirmed defect;
  needs the operator or a fresh copy of that file to close.

## Verification commands used (all run from repo root, 2026-07-22)

Spec + completeness check (`ar-ingest-specs-history.md` § Verification) — passed, exit 0, 43 cols,
EIB-mapped, blocking=none.

feeds/drives check against the real `FA E2E Customers V1.xlsm` (`docs/CONFIGURATION.md` §
"Verifying an ingest spec's feeds/drives") — passed, 0 bad.

Drop-down resolution check (`ar-ingest-specs-history.md` § Item 6 check) — all resolved formulas pointed at
the right cells; this is what surfaced the 26-vs-27 count for Finding 2.

Sheet-name length — not a check the doc lists; found by reading the render's own stderr output,
then confirmed against the raw `xl/workbook.xml`.
