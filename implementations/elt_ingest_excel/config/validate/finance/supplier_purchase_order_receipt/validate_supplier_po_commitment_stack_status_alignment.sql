DROP TABLE IF EXISTS validate_supplier_po_commitment_stack_status_alignment
;
CREATE TABLE validate_supplier_po_commitment_stack_status_alignment
    AS
SELECT order_number
     , line_number
     , po_position
     , delivery_status
     , invoice_status
     , soft_commitment_value_gbp
     , hard_commitment_value_gbp
  FROM src_fin_supplier_po_commitment_stack
 WHERE (
           UPPER(TRIM(COALESCE(delivery_status, ''))) = 'NOT DELIVERED'
       AND COALESCE(hard_commitment_value_gbp, 0) > 0
       )
    OR (
           UPPER(TRIM(COALESCE(delivery_status, ''))) = 'DELIVERED'
       AND COALESCE(soft_commitment_value_gbp, 0) > 0
       )
;

INSERT INTO validation_supplier_purchase_order_receipt_issues
SELECT 'src_fin_supplier_po_commitment_stack'                                   table_name
     , 'STATUS_ALIGNMENT'                                                       rule_type
     , 'delivery_status,soft_commitment_value_gbp,hard_commitment_value_gbp'    col
     , COALESCE(order_number, '') || '_' || COALESCE(line_number, '') || '_'
       || COALESCE(po_position, '')                                             po_id
     , COALESCE(delivery_status, '') || '|' || COALESCE(invoice_status, '') || '|'
       || COALESCE(CAST(soft_commitment_value_gbp AS VARCHAR), '') || '|'
       || COALESCE(CAST(hard_commitment_value_gbp AS VARCHAR), '')              detail
     , 'Delivery status conflicts with the reported soft or hard commitment amount.'
                                                                                message
  FROM validate_supplier_po_commitment_stack_status_alignment
;
