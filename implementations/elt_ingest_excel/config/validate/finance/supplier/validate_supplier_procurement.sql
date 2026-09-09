-- ============================================================================
-- VALIDATE: Supplier Procurement
-- Source table: workday_supplier_procurement
-- Note: This table is currently always empty (WHERE 1=2 in transform SQL —
--       business not ready to populate). Rules recorded for future use.
-- VBA rules: Match (cross-tab), RequiredBetweenTabs (Email → supplier email)
-- ============================================================================

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Supplier ID must exist in Supplier Name tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_procurement'                                  table_name
     , 'CROSS_TAB'                                                   rule_type
     , 'supplier_id'                                                 col
     , p.supplier_id                                                 supplier_id
     , p.supplier_id                                                 detail
     , 'Supplier ID must be listed first on the Supplier Name tab.'  message
  FROM workday_supplier_procurement                                  p
  LEFT OUTER JOIN
       workday_supplier_name                                         n
         ON  n.supplier_id                                           = p.supplier_id
 WHERE n.supplier_id IS NULL
;

-- ----------------------------------------------------------------------------
-- CROSS_TAB: If Purchase Order Issue Option = 'Email', email must be present
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_procurement'                                                   table_name
     , 'CROSS_TAB'                                                                    rule_type
     , 'purchase_order_issue_option'                                                   col
     , p.supplier_id                                                                   supplier_id
     , p.purchase_order_issue_option                                                   detail
     , '"Email" is listed, but Supplier does not have an email on the Supplier Email tab.' message
  FROM workday_supplier_procurement                                                    p
  LEFT OUTER JOIN
       workday_supplier_email                                                          e
         ON  e.supplier_id                                                             = p.supplier_id
 WHERE p.purchase_order_issue_option                                                   = 'Email'
   AND e.supplier_id IS NULL
;
