DROP TABLE IF EXISTS validate_supplier_po_commitment_stack_unmapped_supplier_names
;
-- Triage for Commitment Stack suppliers that do NOT resolve on reference_id token
-- (wd_supplier_id IS NULL in the staged table). Classifies each into:
--   MATCH_ADD_REFERENCE_ID = the supplier already exists in Workday (matched by
--                            normalised name); fix = add the supplier_code token to
--                            that Workday supplier's reference_id, then it
--                            resolves on the next extract.
--   GENUINE_NEW            = no Workday name match; a real new-supplier create.
--
-- Name normalisation mirrors reconcile_supplier_invoice_unresolved, applied to
-- BOTH sides. Review report only: transform always resolves on reference_id token,
-- never on name (name can mis-bind across legal entities). Surfaced for Finance
-- triage only.
CREATE TABLE validate_supplier_po_commitment_stack_unmapped_supplier_names
    AS
  WITH cte_unresolved
    AS (
SELECT DISTINCT
       n.nrm_company_name                                                        gp_business_unit
     , n.nrm_supplier_code                                                       gp_supplier_code
     , n.company_shortcode
     , n.supplier_code
     , n.supplier_name                                                           gp_supplier_name
     , SUM(n.soft_commitment_value_gbp + n.hard_commitment_value_gbp)           total_commitment_gbp
  FROM src_fin_supplier_po_commitment_stack                                     s
       INNER JOIN
       src_fin_supplier_po_commitment_stack_nrm                                 n
          ON n.commitment_stack_line_key = s.commitment_stack_line_key
 WHERE s.wd_supplier_id                                                          IS NULL
 GROUP BY
       n.nrm_company_name
     , n.nrm_supplier_code
     , n.company_shortcode
     , n.supplier_code
     , n.supplier_name
       )
     , cte_wd_norm
    AS (
SELECT w.supplier_id
     , w.supplier                                                                 wd_supplier_name
     , w.reference_id                                                             wd_reference_id_current
     , TRIM(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(
                REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(unaccent(w.supplier)), '.', ''), '-', ' '), 'WOMEN''S', 'WOMEN'), 'WOMENS', 'WOMEN'), 'FOOTBALL CLUB', 'FC'),
                '\s*\([^)]*\)', '', 'g'), '\bUNITED\b', 'UTD', 'g'), '\bAND\b', '&', 'g'),
                '(\s+|^)((LIMITED|LTD|COMPANY|PLC|LLP|INC)\s*)+$', ''), '\s+', ' ', 'g'))
                                                                                    wd_supplier_name_norm
  FROM fin_workday_extract_supplier                                               w
 WHERE NULLIF(TRIM(w.supplier), '')                                                IS NOT NULL
 GROUP BY
       w.supplier_id, w.supplier, w.reference_id
       )
SELECT CASE
           WHEN w.supplier_id IS NOT NULL
           THEN 'MATCH_ADD_REFERENCE_ID'
           ELSE 'GENUINE_NEW'
        END                                                                        match_status
     , u.gp_business_unit
     , u.gp_supplier_code
     , u.company_shortcode
     , u.supplier_code
     , u.gp_supplier_name
     , TRIM(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(
                REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(unaccent(u.gp_supplier_name)), '.', ''), '-', ' '), 'WOMEN''S', 'WOMEN'), 'WOMENS', 'WOMEN'), 'FOOTBALL CLUB', 'FC'),
                '\s*\([^)]*\)', '', 'g'), '\bUNITED\b', 'UTD', 'g'), '\bAND\b', '&', 'g'),
                '(\s+|^)((LIMITED|LTD|COMPANY|PLC|LLP|INC)\s*)+$', ''), '\s+', ' ', 'g'))
                                                                                    gp_supplier_name_norm
     , ROUND(u.total_commitment_gbp, 2)                                            total_commitment_gbp
     , w.supplier_id                                                                wd_supplier_id
     , w.wd_supplier_name
     , w.wd_reference_id_current
  FROM cte_unresolved                                                               u
       LEFT OUTER JOIN
       cte_wd_norm                                                                  w
          ON w.wd_supplier_name_norm = TRIM(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(
                    REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(unaccent(u.gp_supplier_name)), '.', ''), '-', ' '), 'WOMEN''S', 'WOMEN'), 'WOMENS', 'WOMEN'), 'FOOTBALL CLUB', 'FC'),
                    '\s*\([^)]*\)', '', 'g'), '\bUNITED\b', 'UTD', 'g'), '\bAND\b', '&', 'g'),
                    '(\s+|^)((LIMITED|LTD|COMPANY|PLC|LLP|INC)\s*)+$', ''), '\s+', ' ', 'g'))
 ORDER BY
       match_status
     , gp_business_unit
     , gp_supplier_name
;

INSERT INTO validation_supplier_purchase_order_receipt_issues
SELECT 'validate_supplier_po_commitment_stack_unmapped_supplier_names'           table_name
     , 'NAME_MATCH'                                                               rule_type
     , 'wd_supplier_id,reference_id'                                              col
     , TRIM(gp_supplier_code) || '-' || TRIM(gp_business_unit) || '|' ||
       TRIM(gp_supplier_name)                                                     po_id
     , 'wd=' || COALESCE(TRIM(wd_supplier_name), '') || '|wd_id='
       || COALESCE(wd_supplier_id, '') || '|current_ref='
       || COALESCE(wd_reference_id_current, '') || '|commitment_gbp='
       || COALESCE(CAST(total_commitment_gbp AS VARCHAR), '')                    detail
     , 'Supplier not found by reference_id in the map, but NORMALISED NAME matches an existing Workday supplier. '
       || 'Add the token "' || TRIM(gp_supplier_code) || '-' || UPPER(TRIM(gp_business_unit)) || '" (or bare "'
       || TRIM(gp_supplier_code) || '") to that Workday supplier reference_id and re-run.'
                                                                                   message
  FROM validate_supplier_po_commitment_stack_unmapped_supplier_names
 WHERE match_status = 'MATCH_ADD_REFERENCE_ID'
;
