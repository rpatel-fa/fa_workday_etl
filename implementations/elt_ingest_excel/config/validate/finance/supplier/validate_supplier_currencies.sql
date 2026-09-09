-- ============================================================================
-- VALIDATE: Supplier Currencies
-- Source table: workday_supplier_currencies
-- VBA rules: Match (cross-tab Supplier ID), Match (currency code), Missing
-- ============================================================================

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Supplier ID must exist in Supplier Name tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_currencies'                                    table_name
     , 'CROSS_TAB'                                                     rule_type
     , 'supplier_id'                                                   col
     , c.supplier_id                                                   supplier_id
     , c.supplier_id                                                   detail
     , 'Supplier ID must be listed first on the Supplier Name tab.'    message
  FROM workday_supplier_currencies                                     c
  LEFT OUTER JOIN
       workday_supplier_name                                           n
         ON  n.supplier_id                                             = c.supplier_id
 WHERE n.supplier_id IS NULL
;

-- ----------------------------------------------------------------------------
-- MISSING: Accepted Currencies is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_currencies'                                   table_name
     , 'MISSING'                                                      rule_type
     , 'accepted_currencies_plus'                                     col
     , c.supplier_id                                                  supplier_id
     , NULL                                                           detail
     , 'Accepted Currencies must be populated.'                       message
  FROM workday_supplier_currencies                                    c
 WHERE c.accepted_currencies_plus IS NULL
    OR TRIM(c.accepted_currencies_plus) = ''
;

-- ----------------------------------------------------------------------------
-- FORMAT: Default Currency must be a 3-letter ISO currency code (when populated)
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_currencies'                                           table_name
     , 'FORMAT'                                                                rule_type
     , 'default_currency'                                                      col
     , c.supplier_id                                                           supplier_id
     , c.default_currency                                                      detail
     , 'Default Currency must match a valid ISO 4217 currency code (3 uppercase letters).' message
  FROM workday_supplier_currencies                                             c
 WHERE c.default_currency IS NOT NULL
   AND TRIM(c.default_currency) != ''
   AND NOT regexp_matches(TRIM(c.default_currency), '^[A-Z]{3}$')
;

-- ----------------------------------------------------------------------------
-- MISSING: Every supplier in Supplier Name must have a Currencies row
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_currencies'                                           table_name
     , 'MISSING'                                                               rule_type
     , 'accepted_currencies_plus'                                              col
     , n.supplier_id                                                           supplier_id
     , NULL                                                                    detail
     , 'Supplier has no row in Supplier Currencies tab.'                       message
  FROM workday_supplier_name                                                   n
  LEFT OUTER JOIN
       workday_supplier_currencies                                             c
         ON  c.supplier_id                                                     = n.supplier_id
 WHERE c.supplier_id IS NULL
;

-- ----------------------------------------------------------------------------
-- MATCH: Default Currency must be a valid Workday currency code (when populated)
-- ref: ref_workday_validate_supplier_currency (Currency sheet)
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_currencies'                                                   table_name
     , 'MATCH'                                                                        rule_type
     , 'default_currency'                                                             col
     , c.supplier_id                                                                  supplier_id
     , c.default_currency                                                             detail
     , 'Default Currency must be a valid Workday currency code.'                      message
  FROM workday_supplier_currencies                                                    c
  LEFT JOIN
       ref_workday_validate_supplier_currency                                         r
         ON UPPER(TRIM(r.currency_id))                                                = UPPER(TRIM(c.default_currency))
 WHERE c.default_currency IS NOT NULL
   AND TRIM(c.default_currency) != ''
   AND r.currency_id IS NULL
;
