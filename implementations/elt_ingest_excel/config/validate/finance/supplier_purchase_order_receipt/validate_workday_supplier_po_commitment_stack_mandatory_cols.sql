DROP TABLE IF EXISTS validate_workday_supplier_po_commitment_stack_mandatory_cols
;
-- Mandatory-column check on the FINAL Workday PO EIB output.
--
-- Validates the 3 Workday-side mappings that must be populated for the EIB to
-- actually load into Workday: supplier (column_q), spend category (column_fk)
-- and cost centre (column_gv). This is a stricter check than the staged-layer
-- ref-coverage validators because it operates on the post-filter, post-mapping
-- Workday payload — i.e. exactly the rows Workday will receive.
--
-- The staged-layer validators (ref_coverage, missing_refs) are kept separately
-- because they show the accounting-element source side (the *mapping* input),
-- whereas this check catches rows that somehow slipped through with blank
-- resolved values on the actual EIB columns.
CREATE TABLE validate_workday_supplier_po_commitment_stack_mandatory_cols
    AS
SELECT column_p                                                             company
     , column_j                                                             pono
     , column_ex                                                            wd_po_rowid_per_po
     , column_ez                                                            service_order_line_id
     , column_q                                                             wd_supplier_id
     , column_fk                                                            spend_category_ref_id
     , column_gv                                                            cost_centre_ref_id
     , CASE WHEN TRIM(COALESCE(column_q, ''))  = '' THEN 1 ELSE 0 END       missing_supplier
     , CASE WHEN TRIM(COALESCE(column_fk, '')) = '' THEN 1 ELSE 0 END       missing_spend_category
     , CASE WHEN TRIM(COALESCE(column_gv, '')) = '' THEN 1 ELSE 0 END       missing_cost_centre
     , ROUND(TRY_CAST(column_gf AS DECIMAL(18,2)), 2)                      extended_amount_txn
     , column_ah                                                            currency
  FROM workday_supplier_po_commitment_stack
 WHERE TRIM(COALESCE(column_q, ''))  = ''
    OR TRIM(COALESCE(column_fk, '')) = ''
    OR TRIM(COALESCE(column_gv, '')) = ''
 ORDER BY
       missing_supplier  DESC
     , missing_cost_centre  DESC
     , missing_spend_category DESC
     , company
     , pono
     , wd_po_rowid_per_po
;

INSERT INTO validation_supplier_purchase_order_receipt_issues
SELECT 'workday_supplier_po_commitment_stack'                                   table_name
     , CASE WHEN missing_supplier = 1 AND missing_spend_category = 1 AND missing_cost_centre = 1
            THEN 'MANDATORY_SUPPLIER_SPEND_CC_ALL'
           WHEN missing_supplier = 1 AND missing_spend_category = 1
            THEN 'MANDATORY_SUPPLIER_AND_SPEND_CATEGORY'
           WHEN missing_supplier = 1 AND missing_cost_centre = 1
            THEN 'MANDATORY_SUPPLIER_AND_COST_CENTRE'
           WHEN missing_spend_category = 1 AND missing_cost_centre = 1
            THEN 'MANDATORY_SPEND_CATEGORY_AND_COST_CENTRE'
           WHEN missing_supplier = 1
            THEN 'MANDATORY_SUPPLIER'
           WHEN missing_spend_category = 1
            THEN 'MANDATORY_SPEND_CATEGORY'
           ELSE 'MANDATORY_COST_CENTRE'
        END                                                                       rule_type
     , 'column_q,column_fk,column_gv'                                              col
     , company || '|' || pono || '|' || COALESCE(service_order_line_id, '')      po_id
     , 'supplier=' || COALESCE(TRIM(wd_supplier_id), '')
       || '|spend_category=' || COALESCE(TRIM(spend_category_ref_id), '')
       || '|cost_centre=' || COALESCE(TRIM(cost_centre_ref_id), '')
       || '|extended_amount=' || COALESCE(CAST(extended_amount_txn AS VARCHAR), '')
       || '|currency=' || COALESCE(currency, '')                                 detail
     , CASE
           WHEN missing_supplier = 1 AND missing_spend_category = 1 AND missing_cost_centre = 1
           THEN 'Workday PO EIB row has blank Supplier, Spend Category and Cost Centre. All three mandatory columns must be populated before loading.'
           WHEN missing_supplier = 1 AND missing_spend_category = 1
           THEN 'Workday PO EIB row has blank Supplier and Spend Category. Both mandatory columns must be populated before loading.'
           WHEN missing_supplier = 1 AND missing_cost_centre = 1
           THEN 'Workday PO EIB row has blank Supplier and Cost Centre. Both mandatory columns must be populated before loading.'
           WHEN missing_spend_category = 1 AND missing_cost_centre = 1
           THEN 'Workday PO EIB row has blank Spend Category and Cost Centre. Both mandatory columns must be populated before loading.'
           WHEN missing_supplier = 1
           THEN 'Workday PO EIB row has blank Supplier (column_q). The supplier mapping must resolve before loading.'
           WHEN missing_spend_category = 1
           THEN 'Workday PO EIB row has blank Spend Category (column_fk). The FDM spend-category map must cover this accounting string.'
           ELSE 'Workday PO EIB row has blank Cost Centre (column_gv). The FDM cost-centre map must cover this accounting string.'
        END                                                                       message
  FROM validate_workday_supplier_po_commitment_stack_mandatory_cols
;
