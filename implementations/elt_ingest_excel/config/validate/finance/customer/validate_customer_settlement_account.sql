-- ============================================================================
-- VALIDATE: Customer Settlement Account
-- Source table: workday_customer_settlement_account
-- Note: This table is currently disabled (WHERE 1=2 in transform SQL).
--       Rules are recorded here for completeness and to guard future changes.
-- VBA rules: Match (cross-tab), Missing, Match (account type/currency),
--            Duplicatecolumn, bankaccountvalidations
-- ============================================================================

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Customer ID must exist in Customer Name tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_settlement_account'                            table_name
     , 'CROSS_TAB'                                                     rule_type
     , 'customer_id'                                                   col
     , a.customer_id                                                   customer_id
     , a.customer_id                                                   detail
     , 'Customer ID must be listed first on the Customer Name tab.'    message
  FROM workday_customer_settlement_account                             a
  LEFT OUTER JOIN
       workday_customer_name                                           n
         ON  n.customer_id                                             = a.customer_id
 WHERE n.customer_id IS NULL
;

-- ----------------------------------------------------------------------------
-- DUPLICATE: Settlement Bank Account ID must be unique (when populated)
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_settlement_account'                                            table_name
     , 'DUPLICATE'                                                                     rule_type
     , 'settlement_bank_account_id'                                                    col
     , a.customer_id                                                                   customer_id
     , a.settlement_bank_account_id                                                    detail
     , 'Settlement Bank Account ID must be unique and cannot be duplicated.'           message
  FROM workday_customer_settlement_account                                             a
 WHERE a.settlement_bank_account_id IS NOT NULL
   AND a.settlement_bank_account_id IN
       (
         SELECT d.settlement_bank_account_id
           FROM workday_customer_settlement_account                                    d
          WHERE d.settlement_bank_account_id IS NOT NULL
          GROUP BY d.settlement_bank_account_id
         HAVING COUNT(*) > 1
       )
;

-- ----------------------------------------------------------------------------
-- MISSING: Country is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_settlement_account'                            table_name
     , 'MISSING'                                                       rule_type
     , 'country'                                                       col
     , a.customer_id                                                   customer_id
     , NULL                                                            detail
     , 'Country is required.'                                          message
  FROM workday_customer_settlement_account                             a
 WHERE a.country IS NULL
    OR TRIM(a.country) = ''
;

-- ----------------------------------------------------------------------------
-- MISSING: Currency is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_settlement_account'                            table_name
     , 'MISSING'                                                       rule_type
     , 'currency'                                                      col
     , a.customer_id                                                   customer_id
     , NULL                                                            detail
     , 'Currency is a required column.'                                message
  FROM workday_customer_settlement_account                             a
 WHERE a.currency IS NULL
    OR TRIM(a.currency) = ''
;

-- ----------------------------------------------------------------------------
-- FORMAT: Currency must be a 3-letter ISO code
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_settlement_account'                                  table_name
     , 'FORMAT'                                                              rule_type
     , 'currency'                                                            col
     , a.customer_id                                                         customer_id
     , a.currency                                                            detail
     , 'Currency must match a valid ISO 4217 currency code (3 uppercase letters).' message
  FROM workday_customer_settlement_account                                   a
 WHERE a.currency IS NOT NULL
   AND TRIM(a.currency) != ''
   AND NOT regexp_matches(TRIM(a.currency), '^[A-Z]{3}$')
;

-- ----------------------------------------------------------------------------
-- MISSING: Bank Account Type is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_settlement_account'                            table_name
     , 'MISSING'                                                       rule_type
     , 'bank_account_type'                                             col
     , a.customer_id                                                   customer_id
     , NULL                                                            detail
     , 'Bank Account Type is required.'                                message
  FROM workday_customer_settlement_account                             a
 WHERE a.bank_account_type IS NULL
    OR TRIM(a.bank_account_type) = ''
;

-- ----------------------------------------------------------------------------
-- MATCH: Bank Account Type must be Checking or Savings
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_settlement_account'                                            table_name
     , 'MATCH'                                                                        rule_type
     , 'bank_account_type'                                                            col
     , a.customer_id                                                                  customer_id
     , a.bank_account_type                                                            detail
     , 'Bank Account Type must match one of the drop-down values: "Checking", "Savings".' message
  FROM workday_customer_settlement_account                                            a
 WHERE a.bank_account_type IS NOT NULL
   AND TRIM(a.bank_account_type) != ''
   AND a.bank_account_type NOT IN ('Checking', 'Savings')
;

-- ----------------------------------------------------------------------------
-- BANK_ACCOUNT: UK GBP accounts must have either account number or IBAN
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_settlement_account'                                            table_name
     , 'BANK_ACCOUNT'                                                                 rule_type
     , 'bank_account_number / iban'                                                   col
     , a.customer_id                                                                  customer_id
     , a.country || ' / ' || COALESCE(a.currency, 'NULL')                            detail
     , 'UK GBP accounts must have a Bank Account Number or IBAN populated.'          message
  FROM workday_customer_settlement_account                                            a
 WHERE a.country                                                                      = 'United Kingdom'
   AND a.currency                                                                     = 'GBP'
   AND (a.bank_account_number IS NULL OR TRIM(a.bank_account_number) = '')
   AND (a.iban IS NULL OR TRIM(a.iban) = '')
;

-- ----------------------------------------------------------------------------
-- BANK_ACCOUNT: UK GBP accounts should have a sort code (routing_transit_number)
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_settlement_account'                                           table_name
     , 'BANK_ACCOUNT'                                                                rule_type
     , 'routing_transit_number'                                                      col
     , a.customer_id                                                                 customer_id
     , a.bank_account_number                                                         detail
     , 'UK GBP accounts should have a sort code (Routing Number/Bank Code).'        message
  FROM workday_customer_settlement_account                                           a
 WHERE a.country                                                                     = 'United Kingdom'
   AND a.currency                                                                    = 'GBP'
   AND (a.bank_account_number IS NOT NULL AND TRIM(a.bank_account_number) != '')
   AND (a.routing_transit_number IS NULL OR TRIM(a.routing_transit_number) = '')
;

-- ----------------------------------------------------------------------------
-- BANK_FORMAT: IBAN required by country
-- ref: ref_workday_validate_customer_bank_format (Bank Account Formatting sheet)
-- Note: country column is column due to merged header row in source workbook
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_settlement_account'                                            table_name
     , 'BANK_FORMAT'                                                                  rule_type
     , 'iban'                                                                         col
     , a.customer_id                                                                  customer_id
     , a.country                                                                      detail
     , 'IBAN is required for ' || a.country || ' accounts.'                           message
  FROM workday_customer_settlement_account                                             a
  INNER JOIN
       ref_workday_validate_customer_bank_format                                       r
          ON LOWER(TRIM(r.column))                                                 = LOWER(TRIM(a.country))
 WHERE UPPER(TRIM(r.iban_display_option))                                              = 'REQUIRED'
   AND (a.iban IS NULL OR TRIM(a.iban) = '')
;

-- ----------------------------------------------------------------------------
-- BANK_FORMAT: IBAN length must match country requirement
-- ref: ref_workday_validate_customer_bank_format (Bank Account Formatting sheet)
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_settlement_account'                                              table_name
     , 'BANK_FORMAT'                                                                    rule_type
     , 'iban'                                                                           col
     , a.customer_id                                                                    customer_id
     , a.iban || ' (expected ' || r.iban_length || ' chars)'                            detail
     , 'IBAN for ' || a.country || ' must be exactly ' || r.iban_length || ' characters.' message
  FROM workday_customer_settlement_account                                               a
  INNER JOIN
       ref_workday_validate_customer_bank_format                                         r
          ON LOWER(TRIM(r.column))                                                   = LOWER(TRIM(a.country))
 WHERE TRY_CAST(r.iban_length AS INTEGER)                                                > 0
   AND a.iban IS NOT NULL
   AND TRIM(a.iban) != ''
   AND LENGTH(REPLACE(UPPER(TRIM(a.iban)), ' ', ''))                                    != TRY_CAST(r.iban_length AS INTEGER)
;

-- ----------------------------------------------------------------------------
-- BANK_FORMAT: BIC/SWIFT required by country
-- ref: ref_workday_validate_customer_bank_format (Bank Account Formatting sheet)
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_settlement_account'                                            table_name
     , 'BANK_FORMAT'                                                                  rule_type
     , 'swift_bank_identification_code'                                               col
     , a.customer_id                                                                  customer_id
     , a.country                                                                      detail
     , 'BIC/SWIFT code is required for ' || a.country || ' accounts.'                 message
  FROM workday_customer_settlement_account                                             a
  INNER JOIN
       ref_workday_validate_customer_bank_format                                       r
          ON LOWER(TRIM(r.column))                                                 = LOWER(TRIM(a.country))
 WHERE UPPER(TRIM(r.bic_display_option))                                               = 'REQUIRED'
   AND (a.swift_bank_identification_code IS NULL OR TRIM(a.swift_bank_identification_code) = '')
;

-- ----------------------------------------------------------------------------
-- BANK_FORMAT: Bank ID / routing number required by country
-- ref: ref_workday_validate_customer_bank_format (Bank Account Formatting sheet)
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_settlement_account'                                            table_name
     , 'BANK_FORMAT'                                                                  rule_type
     , 'routing_transit_number'                                                       col
     , a.customer_id                                                                  customer_id
     , a.country || ' (' || r.bank_id_label || ')'                                    detail
     , r.bank_id_label || ' is required for ' || a.country || ' accounts.'            message
  FROM workday_customer_settlement_account                                             a
  INNER JOIN
       ref_workday_validate_customer_bank_format                                       r
          ON LOWER(TRIM(r.column))                                                 = LOWER(TRIM(a.country))
 WHERE UPPER(TRIM(r.bank_id_display_option))                                           = 'REQUIRED'
   AND (a.routing_transit_number IS NULL OR TRIM(a.routing_transit_number) = '')
;

-- ----------------------------------------------------------------------------
-- BANK_FORMAT: Account number required by country
-- ref: ref_workday_validate_customer_bank_format (Bank Account Formatting sheet)
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_settlement_account'                                            table_name
     , 'BANK_FORMAT'                                                                  rule_type
     , 'bank_account_number'                                                          col
     , a.customer_id                                                                  customer_id
     , a.country                                                                      detail
     , 'Bank Account Number is required for ' || a.country || ' accounts.'            message
  FROM workday_customer_settlement_account                                             a
  INNER JOIN
       ref_workday_validate_customer_bank_format                                       r
          ON LOWER(TRIM(r.column))                                                 = LOWER(TRIM(a.country))
 WHERE UPPER(TRIM(r.account_number_display_option))                                    = 'REQUIRED'
   AND (a.bank_account_number IS NULL OR TRIM(a.bank_account_number) = '')
   AND (a.iban IS NULL OR TRIM(a.iban) = '')
;
