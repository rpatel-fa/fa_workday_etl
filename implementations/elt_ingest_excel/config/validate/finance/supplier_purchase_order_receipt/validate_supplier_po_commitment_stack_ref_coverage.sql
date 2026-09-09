DROP TABLE IF EXISTS validate_supplier_po_commitment_stack_ref_coverage
;
-- Cross-tab coverage check between the staged PO table and its reference join.
--
-- Every row in src_fin_supplier_po_commitment_stack must resolve to exactly one
-- row in src_fin_supplier_po_commitment_stack_ref on its natural composite key
-- of the 4 accounting elements. Any miss means the FDM lookup is incomplete
-- even for the Workday-supplier-matched population.
CREATE TABLE validate_supplier_po_commitment_stack_ref_coverage
    AS
SELECT s.commitment_stack_line_key
     , s.po_position
     , s.accountingelement1
     , s.accountingelement2
     , s.accountingelement3
     , s.accountingelement4
     , s.order_number
     , s.line_number
     , s.wd_supplier_id
  FROM src_fin_supplier_po_commitment_stack                                  s
       LEFT OUTER JOIN
       src_fin_supplier_po_commitment_stack_ref                              r
          ON TRIM(r.accountingelement1)                                       = TRIM(s.accountingelement1)
         AND TRIM(r.accountingelement2)                                       = TRIM(s.accountingelement2)
         AND TRIM(r.accountingelement3)                                       = TRIM(s.accountingelement3)
         AND TRIM(r.accountingelement4)                                       = TRIM(s.accountingelement4)
 WHERE r.accountingelement1                                                   IS NULL
 ORDER BY
       s.order_number
     , s.line_number
     , s.po_position
;

INSERT INTO validation_supplier_purchase_order_receipt_issues
SELECT 'src_fin_supplier_po_commitment_stack_ref'                            table_name
     , 'CROSS_TAB'                                                           rule_type
     , 'accountingelement1,accountingelement2,accountingelement3,accountingelement4' col
     , TRIM(accountingelement1) || '-' || TRIM(accountingelement2) || '-' ||
       TRIM(accountingelement3) || '-' || TRIM(accountingelement4) || '|' ||
       COALESCE(order_number, '')                                               po_id
     , CAST(po_position AS VARCHAR) || '|commitment_stack_line_key=' ||
       commitment_stack_line_key || '|wd_supplier=' ||
       COALESCE(wd_supplier_id, '')                                          detail
     , 'Staged PO row has no matching FDM reference row in src_fin_supplier_po_commitment_stack_ref on ae1..ae4. The Workday EIB lateral join will still load it but cost centre and spend category resolution must be reviewed.'
                                                                             message
  FROM validate_supplier_po_commitment_stack_ref_coverage
;
