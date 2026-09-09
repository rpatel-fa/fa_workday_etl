-- ============================================================================
-- RECONCILE: Commitment class summary
-- Breakdown of every commitment_class (SOFT / HARD / MIXED / NONE) across the nrm,
-- staged and workday layers with row count and GBP value. Shows how each filter
-- shapes the output — e.g. old-soft cutoff removes only SOFT-class rows on the
-- workday layer.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_cs_commitment_class_summary
;
CREATE TABLE reconcile_cs_commitment_class_summary
    AS
  WITH cte_nrm
    AS (
SELECT COALESCE(UPPER(TRIM(commitment_class)), 'NONE')                     cls
     , COUNT(*)                                                           nrm_rows
     , ROUND(SUM(
         COALESCE(soft_commitment_value_gbp, 0)
       + COALESCE(hard_commitment_value_gbp, 0)
       ), 2)                                                               nrm_value_gbp
  FROM src_fin_supplier_po_commitment_stack_nrm
 GROUP BY COALESCE(UPPER(TRIM(commitment_class)), 'NONE')
       )
     , cte_staged
    AS (
SELECT COALESCE(UPPER(TRIM(commitment_class)), 'NONE')                     cls
     , COUNT(*)                                                           staged_rows
     , ROUND(SUM(
         COALESCE(soft_commitment_value_gbp, 0)
       + COALESCE(hard_commitment_value_gbp, 0)
       ), 2)                                                               staged_value_gbp
  FROM src_fin_supplier_po_commitment_stack
 GROUP BY COALESCE(UPPER(TRIM(commitment_class)), 'NONE')
       )
     , cte_workday
    AS (
SELECT COALESCE(UPPER(TRIM(commitment_class)), 'NONE')                     cls
     , COUNT(*)                                                           workday_rows
     , ROUND(SUM(
         COALESCE(soft_commitment_value_gbp, 0)
       + COALESCE(hard_commitment_value_gbp, 0)
       ), 2)                                                               workday_value_gbp
  FROM src_fin_supplier_po_commitment_stack
 WHERE NOT (
           UPPER(TRIM(COALESCE(commitment_class, ''))) = 'SOFT'
       AND order_date < DATE '2025-08-01'
       )
 GROUP BY COALESCE(UPPER(TRIM(commitment_class)), 'NONE')
       )
SELECT COALESCE(n.cls, s.cls, w.cls)                                      commitment_class
     , COALESCE(n.nrm_rows, 0)                                            nrm_rows
     , COALESCE(n.nrm_value_gbp, 0)                                       nrm_value_gbp
     , COALESCE(s.staged_rows, 0)                                          staged_rows
     , COALESCE(s.staged_value_gbp, 0)                                     staged_value_gbp
     , COALESCE(w.workday_rows, 0)                                        workday_rows
     , COALESCE(w.workday_value_gbp, 0)                                     workday_value_gbp
  FROM cte_nrm                                                            n
       FULL OUTER JOIN
       cte_staged                                                          s
          ON s.cls = n.cls
       FULL OUTER JOIN
       cte_workday                                                         w
          ON w.cls = COALESCE(n.cls, s.cls)
 ORDER BY commitment_class
;
