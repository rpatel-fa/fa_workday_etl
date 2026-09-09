DROP TABLE IF EXISTS validate_supplier_po_commitment_stack_line_id_uniqueness
;
-- Service Order Line ID (column_ez) uniqueness in the Workday PO output.
--
-- column_ez = '{order_number}_{line_number}_{po_position}' and becomes the
-- cross-EIB stable key for receipt linkage. Duplicates here would make any
-- future receipt load ambiguous.
CREATE TABLE validate_supplier_po_commitment_stack_line_id_uniqueness
    AS
SELECT column_j                                                             pono
     , column_ez                                                            service_order_line_id
     , COUNT(*)                                                             rows_in_output
     , STRING_AGG(CAST(column_ex AS VARCHAR), '|' ORDER BY column_ex)       wd_po_rowid_list
  FROM workday_supplier_po_commitment_stack
 GROUP BY
       column_j
     , column_ez
HAVING COUNT(*) > 1
 ORDER BY
       rows_in_output DESC
     , pono
     , service_order_line_id
;

INSERT INTO validation_supplier_purchase_order_receipt_issues
SELECT 'workday_supplier_po_commitment_stack'                                     table_name
     , 'DUPLICATE'                                                                 rule_type
     , 'column_ez'                                                                 col
     , service_order_line_id                                                       po_id
     , pono || '|rows=' || CAST(rows_in_output AS VARCHAR) || '|po_rowids='
       || wd_po_rowid_list                                                         detail
     , 'column_ez (service_order_line_id) must be unique in the Workday PO output; duplicates found above.'
                                                                                   message
  FROM validate_supplier_po_commitment_stack_line_id_uniqueness
;
