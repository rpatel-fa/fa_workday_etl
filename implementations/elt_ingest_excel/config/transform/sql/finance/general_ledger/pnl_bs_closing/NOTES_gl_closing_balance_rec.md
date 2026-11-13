# GL Closing-Balance EIB — Reconciliation Note

Context: Finance (GL Rec, 2026-09-03) flagged six FA balance-sheet codes as
"variances between the EIB and the source data provided ... investigate within
the EIB process." Investigation showed the EIB process is **not** the cause.

## Verified finding

The closing-balance EIB is a **faithful sum of the extract it is built from** —
`opening balance + the 12 monthly movement columns`, netted per account. No
amount is altered anywhere in `line_base → closing_balance_line`.

Proven against live DuckDB (`~/Documents/__data/duckdb/rpatel.duckdb`):

- Every balance-sheet account line in the EIB equals its ingested-source net,
  to the penny — **2,380 lines, 0 exceptions**.
- The only BS accounts not carried are those netting to ~£0 (suppressed by the
  `ABS(net) > 0.00001` gate in `cte_bs`).
- All P&L is rolled into one Retained Earnings (3300) line per entity; FA / NFC /
  WNSL each **net to £0.00** (balanced journal).

## Why Finance saw variances (root cause)

They compared our EIB against **a different GP extract** (Will's GP TB) than the
one the EIB was built from. The two GP pulls differ (in-flight journals /
reclassifications), so the figures differ — the gap is **extract-vs-extract**,
not source→EIB. A reload from a common GP cut-off collapses it.

- 821010 / 850010: the two **tie to £19,395.58 together**; only which code holds
  the ~£5,492.51 differs (a reclassification). Nets to zero across the pair.
- 860030: +£259.49 gap == a July movement in our extract.
- 830020 (+£104,167.00), 822030 (+£3,180.30), 822050 (−£3,848.26): gaps we can
  size but **cannot decompose without their extract**. Do not assert the specific
  GP journal.

## Gotchas for next time

- **Reconcile per entity.** The same GP account string exists in multiple
  entities (e.g. `9020-822030` in FA and NFC). A sum by account alone combines
  entities — split on `SPLIT_PART(column_q, '_', 1)`.
- **The transform is a pass-through.** Feed a different extract → EIB reproduces
  that extract. To tie to Will's TB exactly, regenerate from his extract (needs
  the opening + 12-movement column layout, not a closing-only TB).
- **We cannot account for their variances journal-by-journal** without their
  extract — only prove our side is faithful and bound each gap.
