-- ============================================================================
-- RECONCILE: Old soft-commitment cutoff (DROP_OLD_SOFT bucket)
-- SOFT-class rows with commitment_date < 2025-08-01 — excluded from the Workday EIB
-- (the staged layer includes them; the workday layer filters them out).
-- Per-PO summary grouped by (company, order_number) with count and value.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_cs_old_soft_cutoff
;
CREATE TABLE reconcile_cs_old_soft_cutoff
    AS
SELECT nrm_company_id                                                     company
     , order_number
     , MAX(supplier_name)                                                  supplier_name
     , commitment_date
     , COUNT(*)                                                            row_count
     , ROUND(
         SUM(
           COALESCE(soft_commitment_value_gbp, 0)
           + COALESCE(hard_commitment_value_gbp, 0)
         ), 2)                                                            total_value_gbp
  FROM src_fin_supplier_po_commitment_stack
 WHERE UPPER(TRIM(COALESCE(commitment_class, ''))) = 'SOFT'
   AND commitment_date < DATE '2025-08-01'
 GROUP BY nrm_company_id
        , order_number
        , commitment_date
 ORDER BY total_value_gbp DESC
;
