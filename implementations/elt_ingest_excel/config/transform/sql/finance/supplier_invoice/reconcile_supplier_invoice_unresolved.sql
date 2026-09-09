DROP TABLE IF EXISTS reconcile_supplier_invoice_unresolved
;
-- Triage for invoiced GP suppliers that do NOT resolve on reference_id token
-- (src_fin_supplier_invoice_nrm.wd_supplier_id IS NULL). Classifies each into:
--   MATCH_ADD_REFERENCE_ID = the supplier already exists in Workday (matched by
--                            normalised name); fix = add the GP id to that Workday
--                            supplier's reference_id, then it resolves on the next extract.
--   GENUINE_NEW            = no Workday name match; a real new-supplier create.
--
-- Name normalisation mirrors reconcile_customer_invoice_unresolved (the
-- src_fin_customer_raw.sql chain PLUS hyphen->space and whitespace-collapse),
-- applied to BOTH sides. Review report only: the transform resolves on reference_id
-- token, never on name (a name join can mis-bind to the wrong legal entity). Surfaced
-- for human action (add the token in Workday), never applied automatically.
CREATE TABLE reconcile_supplier_invoice_unresolved
    AS
  WITH cte_unresolved
    AS (
SELECT DISTINCT
       TRIM(n.business_unit)                        gp_business_unit
     , TRIM(n.vendor_id)                            gp_vendor_id
     , TRIM(n.vendor_name)                          gp_vendor_name
  FROM src_fin_supplier_invoice_nrm                 n
 WHERE n.wd_supplier_id                             IS NULL
       )
SELECT CASE
           WHEN w.supplier_id IS NOT NULL
           THEN 'MATCH_ADD_REFERENCE_ID'
           ELSE 'GENUINE_NEW'
       END                                          match_status
     , u.gp_business_unit
     , u.gp_vendor_id
     , u.gp_vendor_name
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
                REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(unaccent(u.gp_vendor_name)), '.', ''), '-', ' '), 'WOMEN''S', 'WOMEN'), 'WOMENS', 'WOMEN'), 'FOOTBALL CLUB', 'FC'),
                '\s*\([^)]*\)', '', 'g'), '\bUNITED\b', 'UTD', 'g'), '\bAND\b', '&', 'g'),
                '(\s+|^)((LIMITED|LTD|COMPANY|PLC|LLP|INC)\s*)+$', ''), '\s+', ' ', 'g'))
 ORDER BY match_status
        , u.gp_business_unit
        , u.gp_vendor_name
;
