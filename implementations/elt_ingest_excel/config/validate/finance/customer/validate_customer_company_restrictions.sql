-- ============================================================================
-- VALIDATE: Customer Company Restrictions
-- Source table: workday_customer_company_restrictions
-- VBA rules: Match (cross-tab Customer ID), Missing (restricted_to_companies)
-- ============================================================================

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Customer ID must exist in Customer Name tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_company_restrictions'                          table_name
     , 'CROSS_TAB'                                                     rule_type
     , 'customer_id'                                                   col
     , r.customer_id                                                   customer_id
     , r.customer_id                                                   detail
     , 'Customer ID must be listed first on the Customer Name tab.'    message
  FROM workday_customer_company_restrictions                           r
  LEFT OUTER JOIN
       workday_customer_name                                           n
         ON  n.customer_id                                             = r.customer_id
 WHERE n.customer_id IS NULL
;

-- ----------------------------------------------------------------------------
-- MISSING: Restricted To Companies must be populated
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_company_restrictions'                          table_name
     , 'MISSING'                                                       rule_type
     , 'restricted_to_companies'                                       col
     , r.customer_id                                                   customer_id
     , NULL                                                            detail
     , 'Restricted To Companies must be populated for every row.'      message
  FROM workday_customer_company_restrictions                           r
 WHERE r.restricted_to_companies IS NULL
    OR TRIM(r.restricted_to_companies) = ''
;

-- ----------------------------------------------------------------------------
SELECT 1
;
