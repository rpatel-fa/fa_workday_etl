-- Mechanical supplier-resolution check for supplier PO / receipt migration.
--
-- Rule:
--   1. Try exact reference token: vendor_id || '-' || company
--   2. Fall back to bare vendor_id
--
-- This mirrors:
--   - src_fin_workday_supplier_map.sql
--   - src_fin_supplier_po_receipt_*_nrm.sql

WITH cte_validate
  AS (
SELECT 'validate_supplier_po_undelivered_supplier'                      validation_table
     , t.nrm_source_tab
     , t.nrm_company_name
     , t.nrm_supplier_code
     , t.company
     , t.supplierno
     , t.suppliertitle
     , t.orderno
     , t.pono
  FROM validate_supplier_po_undelivered_supplier                        t
UNION ALL
SELECT 'validate_supplier_po_receipted_supplier'                       validation_table
     , t.nrm_source_tab
     , t.nrm_company_name
     , t.nrm_supplier_code
     , t.company
     , t.supplierno
     , t.suppliertitle
     , t.orderno
     , t.pono
  FROM validate_supplier_po_receipted_supplier                         t
UNION ALL
SELECT 'validate_supplier_receipt_supplier'                            validation_table
     , t.nrm_source_tab
     , t.nrm_company_name
     , t.nrm_supplier_code
     , t.company
     , t.supplierno
     , t.suppliertitle
     , t.orderno
     , t.pono
  FROM validate_supplier_receipt_supplier                              t
       )
SELECT
       t.validation_table
     , t.nrm_source_tab
     , t.nrm_company_name
     , t.nrm_supplier_code
     , t.company
     , t.supplierno
     , t.suppliertitle
     , t.orderno
     , t.pono
     , t.nrm_supplier_code || '-' || UPPER(TRIM(t.nrm_company_name))   candidate_reference_id_exact
     , t.nrm_supplier_code                                              candidate_reference_id_bare
     , s_exact.reference_id_token                                       matched_reference_id_exact
     , s_bare.reference_id_token                                        matched_reference_id_bare
     , COALESCE(s_exact.reference_id_token, s_bare.reference_id_token)  matched_reference_id_token
     , COALESCE(s_exact.supplier_id, s_bare.supplier_id)                wd_supplier_id
     , COALESCE(s_exact.nrm_supplier_name, s_bare.nrm_supplier_name)    wd_supplier_name
     , COALESCE(s_exact.nrm_agg_supplier_number, s_bare.nrm_agg_supplier_number)
                                                                         wd_reference_id
FROM cte_validate                                                       t
     LEFT OUTER JOIN
     src_fin_workday_supplier_map                                       s_exact
       ON s_exact.reference_id_token = t.nrm_supplier_code || '-' || UPPER(TRIM(t.nrm_company_name))
     LEFT OUTER JOIN
     src_fin_workday_supplier_map                                       s_bare
       ON s_bare.reference_id_token  = t.nrm_supplier_code
ORDER BY
       t.validation_table
     , t.nrm_source_tab
     , t.nrm_company_name
     , t.nrm_supplier_code
     , t.pono
;
