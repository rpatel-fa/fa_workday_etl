-- ============================================================================
-- RECONCILE: Customer row-count funnel
-- Tracks row counts at each transform stage so a drop in customer count can be
-- attributed to a specific, named cause rather than going unnoticed.
-- Unlike the supplier pipeline, src_fin_customer.sql has no additional filter
-- beyond rnk_score = 1, so stage 5 and stage 6 are expected to match exactly.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_customer_row_funnel
;
CREATE TABLE reconcile_customer_row_funnel
    AS
  WITH cte_ingest
    AS (
SELECT COUNT(*) row_count FROM fin_customer_debtor_no_activity_1_year
 UNION ALL
SELECT COUNT(*) FROM fin_customer_debtor_activity_2_years
       )
SELECT '1_raw_ingested_total'                                              stage
     , (SELECT SUM(row_count) FROM cte_ingest)                             row_count
 UNION ALL
SELECT '2_src_fin_customer_base (union of raw tables)'                     stage
     , COUNT(*)                                                            row_count
  FROM src_fin_customer_base
 UNION ALL
SELECT '3_src_fin_customer_raw (normalised, distinct)'                     stage
     , COUNT(*)
  FROM src_fin_customer_raw
 UNION ALL
SELECT '4_distinct_nrm_customer_number'                                    stage
     , COUNT(DISTINCT nrm_customer_number)
  FROM src_fin_customer_raw
 UNION ALL
SELECT '5_canonical_customer_groups (rnk_score=1)'                         stage
     , COUNT(*)
  FROM src_fin_customer_rnk
 WHERE rnk_score = 1
 UNION ALL
SELECT '6_final_src_fin_customer'                                          stage
     , COUNT(*)
  FROM src_fin_customer
;

-- ----------------------------------------------------------------------------
-- UNEXPLAINED_DROP: src_fin_customer.sql has no row-dropping filter beyond
-- rnk_score = 1, so stage 5 and stage 6 counts must match exactly. Any gap
-- means an undocumented filter has been introduced somewhere in the chain.
-- ----------------------------------------------------------------------------
INSERT INTO reconciliation_customer_issues
SELECT
       'src_fin_customer'                                                                       table_name
     , 'UNEXPLAINED_DROP'                                                                       rule_type
     , 'row_count'                                                                              col
     , CAST(NULL AS VARCHAR)                                                                    customer_id
     , (stage5.row_count - stage6.row_count)::VARCHAR                                           detail
     , 'Row-count drop between canonical customer groups and final src_fin_customer is unexplained - no known filter exists in src_fin_customer.sql.' message
  FROM (SELECT COUNT(*) row_count FROM src_fin_customer_rnk WHERE rnk_score = 1)                stage5
     , (SELECT COUNT(*) row_count FROM src_fin_customer)                                        stage6
 WHERE (stage5.row_count - stage6.row_count) != 0
;
