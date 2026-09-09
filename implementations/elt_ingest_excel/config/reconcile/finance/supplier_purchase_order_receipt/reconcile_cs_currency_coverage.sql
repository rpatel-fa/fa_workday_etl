-- ============================================================================
-- RECONCILE: Commitment Stack currency coverage
-- Every source currency present in the staged set must exist in the Workday
-- EIB output. SOURCE_ONLY = a whole currency failed to migrate (break).
-- ============================================================================

DROP TABLE IF EXISTS reconcile_cs_currency_coverage
;
CREATE TABLE reconcile_cs_currency_coverage
    AS
  WITH cte_staged
    AS (
SELECT currency, COUNT(*) source_rows
  FROM src_fin_supplier_po_commitment_stack
 WHERE NOT (
           UPPER(TRIM(COALESCE(commitment_class, ''))) = 'SOFT'
       AND order_date < DATE '2025-08-01'
       )
 GROUP BY currency
       )
     , cte_workday
    AS (
SELECT column_ah currency, COUNT(*) workday_rows
  FROM workday_supplier_po_commitment_stack
 GROUP BY column_ah
       )
SELECT COALESCE(s.currency, w.currency)                                   currency
     , COALESCE(s.source_rows, 0)                                         source_rows
     , COALESCE(w.workday_rows, 0)                                        workday_rows
     , CASE
           WHEN s.currency IS NOT NULL AND w.currency IS NOT NULL THEN 'BOTH'
           WHEN s.currency IS NOT NULL                            THEN 'SOURCE_ONLY'
           ELSE                                                        'TARGET_ONLY'
       END                                                                present_in
  FROM cte_staged                                                         s
       FULL OUTER JOIN
       cte_workday                                                        w
          ON w.currency = s.currency
 ORDER BY currency
;

INSERT INTO reconciliation_supplier_po_receipt_issues
SELECT 'reconcile_cs_currency_coverage'                                   table_name
     , 'CURRENCY_NOT_MIGRATED'                                            rule_type
     , 'currency'                                                         col
     , currency                                                          recon_key
     , source_rows::VARCHAR || ' source rows'                            detail
     , 'A Commitment Stack source currency has zero rows in the Workday PO EIB output.' message
  FROM reconcile_cs_currency_coverage
 WHERE present_in = 'SOURCE_ONLY'
;
