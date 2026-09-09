-- ============================================================================
-- RECONCILE: Unmapped suppliers (INFORMATIONAL — they are NOT dropped)
-- Rows in nrm with wd_supplier_id IS NULL.
-- NOTE: In the CURRENT architecture these rows flow through to the Workday
-- EIB (the wd_supplier_id IS NOT NULL filter is commented in the PO EIB SQL,
-- line 81). They are NOT dropped at the staged layer. They appear in the
-- validation issues table (UNMAPPED_SUPPLIER / MANDATORY_SUPPLIER buckets)
-- for finance to decide: add to Workday supplier map, or — if they ever
-- decide to drop them — we re-enable the commented filter.
-- Grouped by (company, supplier_code, supplier_name) with row count and
-- total commitment value so finance-review impact is obvious.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_cs_unmapped_suppliers
;
CREATE TABLE reconcile_cs_unmapped_suppliers
    AS
SELECT nrm_company_id                                                     company
     , nrm_supplier_code                                                 supplier_code
     , MAX(supplier_name)                                                  supplier_name
     , COUNT(*)                                                            row_count
     , ROUND(
         SUM(
           COALESCE(soft_commitment_value_gbp, 0)
         + COALESCE(hard_commitment_value_gbp, 0)
       ), 2)                                                               total_commitment_gbp
  FROM src_fin_supplier_po_commitment_stack_nrm
 WHERE wd_supplier_id IS NULL
 GROUP BY nrm_company_id
        , nrm_supplier_code
 ORDER BY total_commitment_gbp DESC
        , nrm_company_id
        , nrm_supplier_code
;
