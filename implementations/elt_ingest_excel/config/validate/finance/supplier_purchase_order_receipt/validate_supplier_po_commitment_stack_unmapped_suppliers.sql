DROP TABLE IF EXISTS validate_supplier_po_commitment_stack_unmapped_suppliers
;
CREATE TABLE validate_supplier_po_commitment_stack_unmapped_suppliers
    AS
SELECT nrm_company_id
     , nrm_supplier_code
     , supplier_name
     , order_number
     , line_number
     , po_position
     , COALESCE(soft_commitment_value_gbp, 0)
       + COALESCE(hard_commitment_value_gbp, 0)                         commitment_value_gbp
  FROM src_fin_supplier_po_commitment_stack
 WHERE wd_supplier_id IS NULL
;

INSERT INTO validation_supplier_purchase_order_receipt_issues
SELECT 'src_fin_supplier_po_commitment_stack'                                  table_name
     , 'UNMAPPED_SUPPLIER'                                                     rule_type
     , 'wd_supplier_id'                                                        col
     , COALESCE(order_number, '') || '_' || COALESCE(line_number, '') || '_'
       || COALESCE(po_position, '')                                            po_id
     , nrm_company_id || '|' || nrm_supplier_code || '|'
       || COALESCE(supplier_name, '') || '|PO='
       || COALESCE(order_number, '') || '|line='
       || COALESCE(line_number, '') || '|pos='
       || COALESCE(CAST(po_position AS VARCHAR), '') || '|GBP='
       || COALESCE(CAST(commitment_value_gbp AS VARCHAR), '')                 detail
     , 'No Workday supplier map entry found for code '
       || COALESCE(nrm_supplier_code, '<NULL>')
       || CASE WHEN nrm_supplier_code IS NOT NULL THEN ' (try adding '
              || nrm_supplier_code || '-<company_shortcode> or '
              || nrm_supplier_code || ' to src_fin_workday_supplier_map)' ELSE '' END
                                                                                message
  FROM validate_supplier_po_commitment_stack_unmapped_suppliers
;
