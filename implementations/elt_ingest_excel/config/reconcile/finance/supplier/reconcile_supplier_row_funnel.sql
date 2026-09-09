-- ============================================================================
-- RECONCILE: Supplier row-count funnel
-- Tracks row counts at each transform stage so a drop in supplier count can be
-- attributed to a specific, named cause rather than going unnoticed.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_supplier_row_funnel
;
CREATE TABLE reconcile_supplier_row_funnel
    AS
  WITH cte_ingest
    AS (
SELECT COUNT(*) row_count FROM fin_supplier_creditor_no_activity_1_year
 UNION ALL
SELECT COUNT(*) FROM fin_supplier_creditor_activity_2_years
 UNION ALL
SELECT COUNT(*) FROM fin_supplier_creditor_adhoc_payment
 UNION ALL
SELECT COUNT(*) FROM fin_supplier_creditor_accrual
       )
SELECT '1_raw_ingested_total'                                              stage
     , (SELECT SUM(row_count) FROM cte_ingest)                             row_count
 UNION ALL
SELECT '2_src_fin_supplier_base (union of raw tables)'                     stage
     , COUNT(*)                                                            row_count
  FROM src_fin_supplier_base
 UNION ALL
SELECT '3_src_fin_supplier_raw (normalised, distinct)'                     stage
     , COUNT(*)
  FROM src_fin_supplier_raw
 UNION ALL
SELECT '4_distinct_nrm_supplier_number'                                    stage
     , COUNT(DISTINCT nrm_supplier_number)
  FROM src_fin_supplier_raw
 UNION ALL
SELECT '5_canonical_supplier_groups (rnk_score=1, pre payment-terms filter)' stage
     , COUNT(*)
  FROM src_fin_supplier_rnk
 WHERE rnk_score = 1
 UNION ALL
SELECT '6_final_src_fin_supplier (post nrm_payment_terms_id filter)'       stage
     , COUNT(*)
  FROM src_fin_supplier
;

-- ----------------------------------------------------------------------------
-- FILTERED_OUT: supplier groups with no mapped payment terms are dropped by
-- the `WHERE t.nrm_payment_terms_id IS NOT NULL` filter in src_fin_supplier.sql.
-- This is the one known, deliberate row-dropping filter in the pipeline.
-- ----------------------------------------------------------------------------
INSERT INTO reconciliation_supplier_issues
SELECT
       'src_fin_supplier'                                                                       table_name
     , 'FILTERED_OUT'                                                                           rule_type
     , 'nrm_payment_terms_id'                                                                   col
     , CAST(NULL AS VARCHAR)                                                                    supplier_id
     , r.nrm_agg_supplier_number                                                                detail
     , 'Supplier group has no mapped payment terms and is dropped by src_fin_supplier.sql before a supplier_id is assigned.' message
  FROM src_fin_supplier_rnk                                                                     r
 WHERE r.rnk_score          = 1
   AND r.nrm_payment_terms_id IS NULL
;

-- ----------------------------------------------------------------------------
-- UNEXPLAINED_DROP: the only known filter between stage 5 and stage 6 is the
-- payment-terms filter above. If the row-count gap is not fully explained by
-- it, an undocumented filter has been introduced somewhere in the chain.
-- ----------------------------------------------------------------------------
INSERT INTO reconciliation_supplier_issues
SELECT
       'src_fin_supplier'                                                                       table_name
     , 'UNEXPLAINED_DROP'                                                                       rule_type
     , 'row_count'                                                                              col
     , CAST(NULL AS VARCHAR)                                                                    supplier_id
     , (stage5.row_count - stage6.row_count - filtered.filtered_count)::VARCHAR                 detail
     , 'Row-count drop between canonical supplier groups and final src_fin_supplier is not fully explained by the known nrm_payment_terms_id filter.' message
  FROM (SELECT COUNT(*) row_count FROM src_fin_supplier_rnk WHERE rnk_score = 1)                stage5
     , (SELECT COUNT(*) row_count FROM src_fin_supplier)                                        stage6
     , (SELECT COUNT(*) filtered_count FROM src_fin_supplier_rnk WHERE rnk_score = 1 AND nrm_payment_terms_id IS NULL) filtered
 WHERE (stage5.row_count - stage6.row_count - filtered.filtered_count) != 0
;
