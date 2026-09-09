DROP TABLE IF EXISTS validate_supplier_po_commitment_stack_missing_refs
;
CREATE TABLE validate_supplier_po_commitment_stack_missing_refs
    AS
SELECT DISTINCT
       t.accountingelement1
     , t.accountingelement2
     , t.accountingelement3
     , t.accountingelement4
     , t.cost_centre_ref_id
     , t.hierarchy_cost_centre
     , t.spend_category_ref_id
     , t.hierarchy_spend_category
     , t.spend_category_override_ref_id
     , t.hierarchy_spend_category_override
     , COALESCE(t.spend_category_override_ref_id, t.spend_category_ref_id)                    spend_category_effective_ref_id
     , CASE WHEN TRIM(COALESCE(t.cost_centre_ref_id, '')) = '' THEN 1 ELSE 0 END
                                                                                                missing_cost_centre
     , CASE WHEN TRIM(COALESCE(
                    t.spend_category_override_ref_id, t.spend_category_ref_id, ''
                  )) = '' THEN 1 ELSE 0 END                                                    missing_spend_category
  FROM src_fin_supplier_po_commitment_stack_ref                                             t
 WHERE TRIM(COALESCE(t.cost_centre_ref_id, '')) = ''
    OR TRIM(COALESCE(
                t.spend_category_override_ref_id, t.spend_category_ref_id, ''
              )) = ''
 ORDER BY
       missing_cost_centre  DESC
     , missing_spend_category DESC
     , t.accountingelement1
     , t.accountingelement2
     , t.accountingelement3
     , t.accountingelement4
;

INSERT INTO validation_supplier_purchase_order_receipt_issues
SELECT 'src_fin_supplier_po_commitment_stack_ref'                                            table_name
     , CASE WHEN missing_cost_centre = 1 AND missing_spend_category = 1
            THEN 'MISSING_COST_CENTRE_AND_SPEND_CATEGORY'
           WHEN missing_cost_centre = 1
            THEN 'MISSING_COST_CENTRE'
           ELSE 'MISSING_SPEND_CATEGORY'
        END                                                                                  rule_type
     , 'cost_centre_ref_id,spend_category_ref_id,spend_category_override_ref_id'             col
     , TRIM(accountingelement1) || '-' || TRIM(accountingelement2) || '-' ||
       TRIM(accountingelement3) || '-' || TRIM(accountingelement4) || '|' ||
       COALESCE(spend_category_effective_ref_id, '')                                         po_id
     , COALESCE(TRIM(cost_centre_ref_id), '') || '|cc_hierarchy='
       || COALESCE(hierarchy_cost_centre, '') || '|sc_effective='
       || COALESCE(spend_category_effective_ref_id, '') || '|sc_hierarchy='
       || COALESCE(hierarchy_spend_category_override, hierarchy_spend_category, '')         detail
     , CASE WHEN missing_cost_centre = 1 AND missing_spend_category = 1
            THEN 'No Workday FDM map found for Cost Center and Spend Category on this accounting string.'
           WHEN missing_cost_centre = 1
            THEN 'No Workday FDM map found for Cost Center on this accounting string.'
           ELSE 'No Workday FDM map found for Spend Category on this accounting string.'
        END
       || ' Check fin_ref_wd_fdm_mapping entries for (Cost Center) and (Spend Category) '
       || 'matching ae1..ae4 on the row.'                                                    message
  FROM validate_supplier_po_commitment_stack_missing_refs
;
