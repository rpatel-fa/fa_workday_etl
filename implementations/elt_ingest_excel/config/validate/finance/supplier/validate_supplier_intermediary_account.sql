-- ============================================================================
-- VALIDATE: Supplier Intermediary Account
-- Source table: workday_supplier_intermediary_account
-- VBA rules: Match (cross-tab), Missing, Match (account type/currency),
--            Match (settlement_bank_account_id exists), bankaccountvalidations
-- ============================================================================

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Supplier ID must exist in Supplier Name tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_intermediary_account'                          table_name
     , 'CROSS_TAB'                                                     rule_type
     , 'supplier_id'                                                   col
     , i.supplier_id                                                   supplier_id
     , i.supplier_id                                                   detail
     , 'Supplier ID must be listed first on the Supplier Name tab.'    message
  FROM workday_supplier_intermediary_account                           i
  LEFT OUTER JOIN
       workday_supplier_name                                           n
         ON  n.supplier_id                                             = i.supplier_id
 WHERE n.supplier_id IS NULL
;

-- ----------------------------------------------------------------------------
-- MISSING: Intermediary Bank Account ID is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_intermediary_account'                          table_name
     , 'MISSING'                                                       rule_type
     , 'intermediary_bank_account_id'                                  col
     , i.supplier_id                                                   supplier_id
     , NULL                                                            detail
     , 'Intermediary Bank Account ID is required.'                     message
  FROM workday_supplier_intermediary_account                           i
 WHERE i.intermediary_bank_account_id IS NULL
    OR TRIM(i.intermediary_bank_account_id) = ''
;

-- ----------------------------------------------------------------------------
-- MISSING: Currency is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_intermediary_account'                                table_name
     , 'MISSING'                                                       rule_type
     , 'currency'                                                      col
     , i.supplier_id                                                   supplier_id
     , NULL                                                            detail
     , 'Currency is a required column.'                                message
  FROM workday_supplier_intermediary_account                           i
 WHERE i.currency IS NULL
    OR TRIM(i.currency) = ''
;

-- ----------------------------------------------------------------------------
-- FORMAT: Currency must be a 3-letter ISO code
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_intermediary_account'                                table_name
     , 'FORMAT'                                                              rule_type
     , 'currency'                                                            col
     , i.supplier_id                                                         supplier_id
     , i.currency                                                            detail
     , 'Currency must match a valid ISO 4217 currency code (3 uppercase letters).' message
  FROM workday_supplier_intermediary_account                                 i
 WHERE i.currency IS NOT NULL
   AND TRIM(i.currency) != ''
   AND NOT regexp_matches(TRIM(i.currency), '^[A-Z]{3}$')
;

-- ----------------------------------------------------------------------------
-- MISSING: Bank Account Type is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_intermediary_account'                           table_name
     , 'MISSING'                                                       rule_type
     , 'bank_account_type'                                             col
     , i.supplier_id                                                   supplier_id
     , NULL                                                            detail
     , 'Bank Account Type is required.'                                message
  FROM workday_supplier_intermediary_account                           i
 WHERE i.bank_account_type IS NULL
    OR TRIM(i.bank_account_type) = ''
;

-- ----------------------------------------------------------------------------
-- MATCH: Bank Account Type must be Checking or Savings
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_intermediary_account'                                    table_name
     , 'MATCH'                                                                        rule_type
     , 'bank_account_type'                                                            col
     , i.supplier_id                                                                  supplier_id
     , i.bank_account_type                                                            detail
     , 'Bank Account Type must match one of the drop-down values: "Checking", "Savings".' message
  FROM workday_supplier_intermediary_account                                          i
 WHERE i.bank_account_type IS NOT NULL
   AND TRIM(i.bank_account_type) != ''
   AND i.bank_account_type NOT IN ('Checking', 'Savings', 'CHECKING', 'SAVINGS')
;

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Settlement Bank Account ID must exist in Settlement Account tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_intermediary_account'                                    table_name
     , 'CROSS_TAB'                                                                    rule_type
     , 'settlement_bank_account_id'                                                   col
     , i.supplier_id                                                                  supplier_id
     , i.settlement_bank_account_id                                                   detail
     , 'The linked Settlement Bank Account ID must first exist on the Settlement Account tab.' message
  FROM workday_supplier_intermediary_account                                          i
  LEFT OUTER JOIN
       workday_supplier_settlement_account                                            a
         ON  a.settlement_bank_account_id                                             = i.settlement_bank_account_id
 WHERE i.settlement_bank_account_id IS NOT NULL
   AND a.settlement_bank_account_id IS NULL
;
