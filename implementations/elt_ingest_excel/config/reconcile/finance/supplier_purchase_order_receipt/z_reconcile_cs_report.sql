-- ============================================================================
-- RECONCILE: Commitment Stack PO consolidated sign-off REPORT.
-- Confirms the source (Commitment Stack Report from Proactis) reconciles to the
-- Workday PO EIB payload this pipeline generates. Every metric is a source vs
-- target pair with an explicit residual and CONFIRMED / REVIEW status.
-- Two ACTUAL row-drop gates: zero-commitment + old-soft cutoff. A third
-- informational bucket tracks unmapped supplier exposure (they flow through
-- to the EIB currently, not dropped — see PO EIB SQL for the commented gate).
-- Depends on: row_counts, amount_by_currency, currency_coverage (runs last, z_).
-- ============================================================================

DROP TABLE IF EXISTS reconcile_cs_report
;
CREATE TABLE reconcile_cs_report
    AS
-- 1. ROW COUNTS — staged vs workday EIB rows (1:1 expected; unmapped suppliers NOT dropped).
SELECT '1_row_counts'                                                         section
     , company                                                                metric
     , staged_rows::VARCHAR                                                   source_value
     , workday_rows::VARCHAR                                                  target_value
     , unexplained_workday::VARCHAR                                           difference
     , CASE WHEN unexplained_workday = 0 THEN 'CONFIRMED' ELSE 'REVIEW' END   status
  FROM reconcile_cs_row_counts

 UNION ALL
-- 1a. TOTAL row count across all companies
SELECT '1_row_counts'
     , 'TOTAL'
     , SUM(staged_rows)::VARCHAR
     , SUM(workday_rows)::VARCHAR
     , SUM(unexplained_workday)::VARCHAR
     , CASE WHEN SUM(unexplained_workday) = 0 THEN 'CONFIRMED' ELSE 'REVIEW' END
  FROM reconcile_cs_row_counts

 UNION ALL
-- 1b. Drop bucket breakdown (INFORMATIONAL: 2 ACTUAL drops + unmapped NOT dropped)
SELECT '1b_drop_buckets'
     , 'UNMAPPED_SUPPLIERS (NOT DROPPED — flow to Workday with NULL supplier_id)'
     , SUM(unmapped_supplier_rows)::VARCHAR
     , NULL
     , NULL
     , 'INFORMATIONAL'
  FROM reconcile_cs_row_counts

 UNION ALL
SELECT '1b_drop_buckets'
     , 'DROP_ZERO_COMMITMENT'
     , SUM(drop_zero_commitment)::VARCHAR
     , NULL
     , NULL
     , 'INFORMATIONAL'
  FROM reconcile_cs_row_counts

 UNION ALL
SELECT '1b_drop_buckets'
     , 'DROP_OLD_SOFT_CUTOFF'
     , SUM(drop_old_soft)::VARCHAR
     , NULL
     , NULL
     , 'INFORMATIONAL'
  FROM reconcile_cs_row_counts

 UNION ALL
-- 2. LINE VALUE by (company, currency) — staged vs workday extended_amount
-- (1:1 expected; unmapped suppliers NOT dropped from the EIB, so no subtraction).
SELECT '2_amount_by_currency'
     , company || ' | ' || currency
     , ROUND(staged_amount, 2)::VARCHAR
     , ROUND(workday_amount, 2)::VARCHAR
     , ROUND(unexplained_workday_amount, 2)::VARCHAR
     , CASE WHEN ABS(unexplained_workday_amount) < 0.005 THEN 'CONFIRMED' ELSE 'REVIEW' END
  FROM reconcile_cs_amount_by_currency

 UNION ALL
-- 2a. TOTAL value across all (same GBP row-level £2 tolerance; non-GBP on
-- transaction currency — so sum is multi-currency not directly meaningful,
-- but shown for completeness; the per-currency rows are the real control).
SELECT '2_amount_by_currency'
     , 'ALL_CURRENCIES_TOTAL'
     , ROUND(SUM(staged_amount), 2)::VARCHAR
     , ROUND(SUM(workday_amount), 2)::VARCHAR
     , ROUND(SUM(unexplained_workday_amount), 2)::VARCHAR
     , CASE WHEN ABS(SUM(unexplained_workday_amount)) < 0.005 THEN 'CONFIRMED' ELSE 'REVIEW' END
  FROM reconcile_cs_amount_by_currency

 UNION ALL
-- 2b. Unmapped supplier exposure (INFORMATIONAL — they flow to the EIB,
-- not dropped. Bucket tracks the value committed with NULL supplier_id
-- so finance can size the resolution effort).
SELECT '2_amount_by_currency'
     , 'UNMAPPED_SUPPLIERS (NOT DROPPED — see validation_issues)'
     , ROUND(SUM(unmapped_supplier_amount), 2)::VARCHAR
     , NULL
     , NULL
     , 'INFORMATIONAL'
  FROM reconcile_cs_amount_by_currency

 UNION ALL
-- 3. CURRENCY COVERAGE — every source currency present in the target EIB.
SELECT '3_currency_coverage'
     , currency
     , source_rows::VARCHAR
     , workday_rows::VARCHAR
     , (source_rows - workday_rows)::VARCHAR
     , CASE WHEN present_in = 'BOTH' THEN 'CONFIRMED' ELSE 'REVIEW' END
  FROM reconcile_cs_currency_coverage

 UNION ALL
-- 4. WORKDAY EIB OUTPUT COUNT — rows actually created (informational, target side).
SELECT '4_workday_eib_output_counts'
     , 'workday_supplier_po_commitment_stack'
     , NULL
     , row_count::VARCHAR
     , NULL
     , NULL
  FROM (SELECT COUNT(*) row_count FROM workday_supplier_po_commitment_stack)  t

 ORDER BY 1, 2
;
