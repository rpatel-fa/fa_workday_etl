-- ============================================================================
-- VALIDATE: Customer Currencies
-- Source table: workday_customer_currencies
-- VBA rules: Match (cross-tab Customer ID), Match (currency code), Missing
-- ============================================================================

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Customer ID must exist in Customer Name tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_currencies'                                    table_name
     , 'CROSS_TAB'                                                     rule_type
     , 'customer_id'                                                   col
     , c.customer_id                                                   customer_id
     , c.customer_id                                                   detail
     , 'Customer ID must be listed first on the Customer Name tab.'    message
  FROM workday_customer_currencies                                     c
  LEFT OUTER JOIN
       workday_customer_name                                           n
         ON  n.customer_id                                             = c.customer_id
 WHERE n.customer_id IS NULL
;

-- ----------------------------------------------------------------------------
-- MISSING: Accepted Currencies is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_currencies'                                   table_name
     , 'MISSING'                                                      rule_type
     , 'accepted_currencies_plus'                                     col
     , c.customer_id                                                  customer_id
     , NULL                                                           detail
     , 'Accepted Currencies must be populated.'                       message
  FROM workday_customer_currencies                                    c
 WHERE c.accepted_currencies_plus IS NULL
    OR TRIM(c.accepted_currencies_plus) = ''
;

-- ----------------------------------------------------------------------------
-- FORMAT: Default Currency must be a 3-letter ISO currency code (when populated)
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_currencies'                                           table_name
     , 'FORMAT'                                                                rule_type
     , 'default_currency'                                                      col
     , c.customer_id                                                           customer_id
     , c.default_currency                                                      detail
     , 'Default Currency must match a valid ISO 4217 currency code (3 uppercase letters).' message
  FROM workday_customer_currencies                                             c
 WHERE c.default_currency IS NOT NULL
   AND TRIM(c.default_currency) != ''
   AND NOT regexp_matches(TRIM(c.default_currency), '^[A-Z]{3}$')
;

-- ----------------------------------------------------------------------------
-- MISSING: Every customer in Customer Name must have a Currencies row
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_currencies'                                           table_name
     , 'MISSING'                                                               rule_type
     , 'accepted_currencies_plus'                                              col
     , n.customer_id                                                           customer_id
     , NULL                                                                    detail
     , 'Customer has no row in Customer Currencies tab.'                       message
  FROM workday_customer_name                                                   n
  LEFT OUTER JOIN
       workday_customer_currencies                                             c
         ON  c.customer_id                                                     = n.customer_id
 WHERE c.customer_id IS NULL
;

-- ----------------------------------------------------------------------------
-- MATCH: Default Currency must be a valid Workday currency code (when populated)
-- ref: ref_workday_validate_customer_currency (Currency sheet)
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_currencies'                                                   table_name
     , 'MATCH'                                                                        rule_type
     , 'default_currency'                                                             col
     , c.customer_id                                                                  customer_id
     , c.default_currency                                                             detail
     , 'Default Currency must be a valid Workday currency code.'                      message
  FROM workday_customer_currencies                                                    c
  LEFT JOIN
       ref_workday_validate_customer_currency                                         r
         ON UPPER(TRIM(r.currency_id))                                                = UPPER(TRIM(c.default_currency))
 WHERE c.default_currency IS NOT NULL
   AND TRIM(c.default_currency) != ''
   AND r.currency_id IS NULL
;
