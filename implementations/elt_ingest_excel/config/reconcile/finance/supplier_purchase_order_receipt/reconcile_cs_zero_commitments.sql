-- ============================================================================
-- RECONCILE: Zero-commitment rows (DROP_ZERO_COMMITMENT bucket)
-- Rows in nrm that have a valid supplier but soft + hard commitment <= 0.
-- These are excluded from the staged layer. Shown per PO with counts and value
-- per (company, order_number, line_number) so finance can audit which POs are
-- fully closed out (no remaining commitment).
-- ============================================================================

DROP TABLE IF EXISTS reconcile_cs_zero_commitments
;
CREATE TABLE reconcile_cs_zero_commitments
    AS
SELECT nrm_company_id                                                     company
     , order_number
     , MAX(supplier_name)                                                  supplier_name
     , COUNT(*)                                                            row_count
     , ROUND(
         SUM(
           COALESCE(soft_commitment_value_gbp, 0)
         + COALESCE(hard_commitment_value_gbp, 0)
       ), 2)                                                               total_net_value_gbp
     , ROUND(SUM(COALESCE(soft_commitment_value_gbp, 0)), 2)                total_soft_gbp
     , ROUND(SUM(COALESCE(hard_commitment_value_gbp, 0)), 2)              total_hard_gbp
  FROM src_fin_supplier_po_commitment_stack_nrm
 WHERE wd_supplier_id IS NOT NULL
   AND COALESCE(soft_commitment_value_gbp, 0)
     + COALESCE(hard_commitment_value_gbp, 0) <= 0
 GROUP BY nrm_company_id
        , order_number
 ORDER BY nrm_company_id
        , order_number
;
