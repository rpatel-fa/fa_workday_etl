DROP TABLE IF EXISTS validate_supplier_po_commitment_stack_tracked_spend_project
;
CREATE TABLE validate_supplier_po_commitment_stack_tracked_spend_project
    AS
SELECT DISTINCT
       t.accountingelement1
     , t.accountingelement2
     , t.accountingelement3
     , t.accountingelement4
     , COALESCE(t.spend_category_override_ref_id, t.spend_category_ref_id)                    spend_category_effective_ref_id
     , t.project_capex_ref_id
     , t.project_opex_ref_id
  FROM src_fin_supplier_po_commitment_stack_ref                                             t
       INNER JOIN
       fin_ref_wd_spend_category                                                            sc
          ON TRIM(sc.reference_id)                                                           = TRIM(COALESCE(t.spend_category_override_ref_id, t.spend_category_ref_id))
 WHERE UPPER(TRIM(COALESCE(sc.spend_category_is_tracked, '')))                              = 'YES'
   AND NULLIF(TRIM(COALESCE(t.project_capex_ref_id, '')), '')                              IS NULL
   AND NULLIF(TRIM(COALESCE(t.project_opex_ref_id, '')), '')                               IS NULL
 ORDER BY
       t.accountingelement1
     , t.accountingelement2
     , t.accountingelement3
     , t.accountingelement4
     , spend_category_effective_ref_id
;

INSERT INTO validation_supplier_purchase_order_receipt_issues
SELECT 'validate_supplier_po_commitment_stack_tracked_spend_project'                         table_name
     , 'TRACKED_SPEND_NO_PROJECT'                                                            rule_type
     , 'project_capex_ref_id,project_opex_ref_id'                                            col
     , TRIM(accountingelement1) || '-' || TRIM(accountingelement2) || '-' ||
       TRIM(accountingelement3) || '-' || TRIM(accountingelement4) || '|' ||
       TRIM(spend_category_effective_ref_id)                                                 po_id
     , TRIM(spend_category_effective_ref_id)                                                 detail
     , 'Tracked spend category requires a project code but neither capex nor opex project mapping is populated on this accounting string.'
                                                                                             message
  FROM validate_supplier_po_commitment_stack_tracked_spend_project
;
