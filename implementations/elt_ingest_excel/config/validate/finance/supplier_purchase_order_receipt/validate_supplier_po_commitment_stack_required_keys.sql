DROP TABLE IF EXISTS validate_supplier_po_commitment_stack_required_keys
;
CREATE TABLE validate_supplier_po_commitment_stack_required_keys
    AS
SELECT company_shortcode
     , order_number
     , line_number
     , po_position
     , currency
  FROM src_fin_supplier_po_commitment_stack
 WHERE TRIM(COALESCE(company_shortcode, '')) = ''
    OR TRIM(COALESCE(order_number, ''))      = ''
    OR TRIM(COALESCE(line_number, ''))       = ''
    OR TRIM(COALESCE(currency, ''))          = ''
;

INSERT INTO validation_supplier_purchase_order_receipt_issues
SELECT 'src_fin_supplier_po_commitment_stack'                                   table_name
     , 'REQUIRED'                                                               rule_type
     , 'company_shortcode,order_number,line_number,currency'                    col
     , COALESCE(order_number, '') || '_' || COALESCE(line_number, '') || '_'
       || COALESCE(po_position, '')                                             po_id
     , COALESCE(company_shortcode, '') || '|' || COALESCE(order_number, '') || '|'
       || COALESCE(line_number, '') || '|' || COALESCE(currency, '')           detail
     , 'Commitment stack row is missing one or more required business keys.'   message
  FROM validate_supplier_po_commitment_stack_required_keys
;
