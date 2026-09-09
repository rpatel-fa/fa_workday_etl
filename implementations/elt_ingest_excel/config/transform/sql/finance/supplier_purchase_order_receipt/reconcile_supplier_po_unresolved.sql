DROP TABLE IF EXISTS reconcile_supplier_po_unresolved
;
-- Thin layer 1 for the supplier PO / commitment stack unresolved suppliers.
-- Mirrors reconcile_customer_invoice_unresolved exactly: 7 columns, two statuses,
-- no consolidation, no enrichment - kept intentionally thin for clean lineage.
-- The unresolved set is the distinct (business_unit, supplier_code, supplier_name)
-- rows from src_fin_supplier_po_commitment_stack where wd_supplier_id IS NULL
-- (i.e. the nrm layer's two-token reference_id join did not find a WD supplier).
--
-- Status buckets:
--   MATCH_ADD_REFERENCE_ID  -- name matches a supplier ALREADY in Workday extract.
--                            Fix is to add the GP supplier_code (+ BU suffix or bare)
--                            to that WD supplier's reference_id.
--   GENUINE_NEW             -- no name match in WD extract. Must be created.
--
-- Name normalisation on both sides is identical to the customer-invoice and
-- supplier-invoice reconcile layers (strip dots, hyphen->space, FC->FC,
-- drop trailing LIMITED/LTD/COMPANY/PLC/LLP/INC, unaccent, collapse whitespace).
CREATE TABLE reconcile_supplier_po_unresolved
    AS
  WITH cte_unresolved
    AS (
SELECT DISTINCT
       TRIM(s.nrm_company_name)                      gp_business_unit
     , TRIM(s.nrm_supplier_code)                     gp_supplier_code
     , TRIM(s.supplier_name)                         gp_supplier_name
  FROM src_fin_supplier_po_commitment_stack          s
 WHERE s.wd_supplier_id                             IS NULL
       )
SELECT CASE
           WHEN w.supplier_id IS NOT NULL
           THEN 'MATCH_ADD_REFERENCE_ID'
           ELSE 'GENUINE_NEW'
       END                                          match_status
     , u.gp_business_unit
     , u.gp_supplier_code
     , u.gp_supplier_name
     , w.supplier_id                                wd_supplier_id
     , w.supplier                                   wd_supplier_name
     , w.reference_id                               wd_reference_id_current
  FROM cte_unresolved                               u
       LEFT OUTER JOIN
       fin_workday_extract_supplier                 w
          ON TRIM(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(
                REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(unaccent(w.supplier)), '.', ''), '-', ' '), 'WOMEN''S', 'WOMEN'), 'WOMENS', 'WOMEN'), 'FOOTBALL CLUB', 'FC'),
                '\s*\([^)]*\)', '', 'g'), '\bUNITED\b', 'UTD', 'g'), '\bAND\b', '&', 'g'),
                '(\s+|^)((LIMITED|LTD|COMPANY|PLC|LLP|INC)\s*)+$', ''), '\s+', ' ', 'g'))
           =  TRIM(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(
                REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(unaccent(u.gp_supplier_name)), '.', ''), '-', ' '), 'WOMEN''S', 'WOMEN'), 'WOMENS', 'WOMEN'), 'FOOTBALL CLUB', 'FC'),
                '\s*\([^)]*\)', '', 'g'), '\bUNITED\b', 'UTD', 'g'), '\bAND\b', '&', 'g'),
                '(\s+|^)((LIMITED|LTD|COMPANY|PLC|LLP|INC)\s*)+$', ''), '\s+', ' ', 'g'))
 ORDER BY match_status
        , u.gp_business_unit
        , u.gp_supplier_name
;