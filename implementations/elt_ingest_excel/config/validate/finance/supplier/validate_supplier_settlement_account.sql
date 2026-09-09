-- ============================================================================
-- VALIDATE: Supplier Settlement Account
-- Source table: workday_supplier_settlement_account
-- VBA rules: Match (cross-tab), Missing, Match (account type/currency),
--            Duplicatecolumn, bankaccountvalidations
-- ============================================================================

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Supplier ID must exist in Supplier Name tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                            table_name
     , 'CROSS_TAB'                                                     rule_type
     , 'supplier_id'                                                   col
     , a.supplier_id                                                   supplier_id
     , a.supplier_id                                                   detail
     , 'Supplier ID must be listed first on the Supplier Name tab.'    message
  FROM workday_supplier_settlement_account                             a
  LEFT OUTER JOIN
       workday_supplier_name                                           n
         ON  n.supplier_id                                             = a.supplier_id
 WHERE n.supplier_id IS NULL
;

-- ----------------------------------------------------------------------------
-- DUPLICATE: Settlement Bank Account ID must be unique (when populated)
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                                            table_name
     , 'DUPLICATE'                                                                     rule_type
     , 'settlement_bank_account_id'                                                    col
     , a.supplier_id                                                                   supplier_id
     , a.settlement_bank_account_id                                                    detail
     , 'Settlement Bank Account ID must be unique and cannot be duplicated.'           message
  FROM workday_supplier_settlement_account                                             a
 WHERE a.settlement_bank_account_id IS NOT NULL
   AND a.settlement_bank_account_id IN
       (
         SELECT d.settlement_bank_account_id
           FROM workday_supplier_settlement_account                                    d
          WHERE d.settlement_bank_account_id IS NOT NULL
          GROUP BY d.settlement_bank_account_id
         HAVING COUNT(*) > 1
       )
;

-- ----------------------------------------------------------------------------
-- MISSING: Bank Country is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                            table_name
     , 'MISSING'                                                       rule_type
     , 'bank_country'                                                  col
     , a.supplier_id                                                   supplier_id
     , NULL                                                            detail
     , 'Bank Country is required.'                                     message
  FROM workday_supplier_settlement_account                             a
 WHERE a.bank_country IS NULL
    OR TRIM(a.bank_country) = ''
;

-- ----------------------------------------------------------------------------
-- MISSING: Currency is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                            table_name
     , 'MISSING'                                                       rule_type
     , 'currency'                                                      col
     , a.supplier_id                                                   supplier_id
     , NULL                                                            detail
     , 'Currency is a required column.'                                message
  FROM workday_supplier_settlement_account                             a
 WHERE a.currency IS NULL
    OR TRIM(a.currency) = ''
;

-- ----------------------------------------------------------------------------
-- FORMAT: Currency must be a 3-letter ISO code
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                                  table_name
     , 'FORMAT'                                                              rule_type
     , 'currency'                                                            col
     , a.supplier_id                                                         supplier_id
     , a.currency                                                            detail
     , 'Currency must match a valid ISO 4217 currency code (3 uppercase letters).' message
  FROM workday_supplier_settlement_account                                   a
 WHERE a.currency IS NOT NULL
   AND TRIM(a.currency) != ''
   AND NOT regexp_matches(TRIM(a.currency), '^[A-Z]{3}$')
;

-- ----------------------------------------------------------------------------
-- MISSING: Bank Account Type is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                            table_name
     , 'MISSING'                                                       rule_type
     , 'bank_account_type'                                             col
     , a.supplier_id                                                   supplier_id
     , NULL                                                            detail
     , 'Bank Account Type is required.'                                message
  FROM workday_supplier_settlement_account                             a
 WHERE a.bank_account_type IS NULL
    OR TRIM(a.bank_account_type) = ''
;

-- ----------------------------------------------------------------------------
-- MATCH: Bank Account Type must be Checking or Savings
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                                            table_name
     , 'MATCH'                                                                        rule_type
     , 'bank_account_type'                                                            col
     , a.supplier_id                                                                  supplier_id
     , a.bank_account_type                                                            detail
     , 'Bank Account Type must match one of the drop-down values: "Checking", "Savings".' message
  FROM workday_supplier_settlement_account                                            a
 WHERE a.bank_account_type IS NOT NULL
   AND TRIM(a.bank_account_type) != ''
   AND a.bank_account_type NOT IN ('Checking', 'Savings')
;

-- ----------------------------------------------------------------------------
-- BANK_ACCOUNT: UK GBP accounts must have either account number or IBAN
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                                            table_name
     , 'BANK_ACCOUNT'                                                                 rule_type
     , 'bank_account_number / bank_iban'                                              col
     , a.supplier_id                                                                  supplier_id
     , a.bank_country || ' / ' || COALESCE(a.currency, 'NULL')                       detail
     , 'UK GBP accounts must have a Bank Account Number or IBAN populated.'          message
  FROM workday_supplier_settlement_account                                            a
 WHERE a.bank_country                                                                 = 'United Kingdom'
   AND a.currency                                                                     = 'GBP'
   AND (a.bank_account_number IS NULL OR TRIM(a.bank_account_number) = '')
   AND (a.bank_iban IS NULL OR TRIM(a.bank_iban) = '')
;

-- ----------------------------------------------------------------------------
-- BANK_ACCOUNT: UK GBP accounts should have a sort code (bank_code_routing_number)
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                                           table_name
     , 'BANK_ACCOUNT'                                                                rule_type
     , 'bank_code_routing_number'                                                    col
     , a.supplier_id                                                                 supplier_id
     , a.bank_account_number                                                         detail
     , 'UK GBP accounts should have a sort code (Routing Number/Bank Code).'        message
  FROM workday_supplier_settlement_account                                           a
 WHERE a.bank_country                                                                = 'United Kingdom'
   AND a.currency                                                                    = 'GBP'
   AND (a.bank_account_number IS NOT NULL AND TRIM(a.bank_account_number) != '')
   AND (a.bank_code_routing_number IS NULL OR TRIM(a.bank_code_routing_number) = '')
;

-- ----------------------------------------------------------------------------
-- BANK_FORMAT: UK sort code (Routing Number) must be exactly 6 numeric digits.
-- Catches truncated codes (a dropped leading zero, e.g. '40003') and BIC/SWIFT
-- values wrongly placed in the sort-code field (e.g. 'MIDLGB22'). Mirrors the
-- Workday rule "Bank Sort Code must be 6 numeric digits" that the internal
-- validation previously did not cover.
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                                           table_name
     , 'BANK_FORMAT'                                                                  rule_type
     , 'bank_code_routing_number'                                                     col
     , a.supplier_id                                                                  supplier_id
     , a.bank_code_routing_number                                                     detail
     , 'UK sort code "' || a.bank_code_routing_number || '" must be 6 numeric digits.' message
  FROM workday_supplier_settlement_account                                            a
 WHERE a.bank_country                                                                 = 'United Kingdom'
   AND (a.bank_account_number IS NOT NULL AND TRIM(a.bank_account_number) != '')
   AND a.bank_code_routing_number IS NOT NULL
   AND TRIM(a.bank_code_routing_number)                                               != ''
   AND NOT regexp_matches(REGEXP_REPLACE(a.bank_code_routing_number, '[^0-9]', '', 'g'), '^[0-9]{6}$')
;

-- ----------------------------------------------------------------------------
-- BANK_FORMAT: UK account number must be exactly 8 numeric digits.
-- The 7-digit dropped-leading-zero case is repaired upstream; this surfaces what
-- remains - 6-digit (possible legacy/building-society account) and 9/10-digit
-- (roll-number concatenation) - for review rather than silently submitting them.
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                                           table_name
     , 'BANK_FORMAT'                                                                  rule_type
     , 'bank_account_number'                                                          col
     , a.supplier_id                                                                  supplier_id
     , a.bank_account_number                                                          detail
     , 'UK account number "' || a.bank_account_number || '" must be 8 numeric digits.' message
  FROM workday_supplier_settlement_account                                            a
 WHERE a.bank_country                                                                 = 'United Kingdom'
   AND a.bank_account_number IS NOT NULL
   AND TRIM(a.bank_account_number)                                                    != ''
   AND NOT regexp_matches(REGEXP_REPLACE(a.bank_account_number, '[^0-9]', '', 'g'), '^[0-9]{8}$')
;

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Currency should match Supplier Currencies tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                                           table_name
     , 'CROSS_TAB'                                                                   rule_type
     , 'currency'                                                                    col
     , a.supplier_id                                                                 supplier_id
     , a.currency || ' vs ' || COALESCE(c.accepted_currencies_plus, 'none')         detail
     , 'Settlement Account currency does not match Supplier Currencies tab.'         message
  FROM workday_supplier_settlement_account                                           a
  LEFT OUTER JOIN
       workday_supplier_currencies                                                   c
         ON  c.supplier_id                                                           = a.supplier_id
        AND (c.accepted_currencies_plus                                              = a.currency
         OR  c.accepted_currencies_plus                                              = 'All Currencies')
 WHERE a.currency IS NOT NULL
   AND c.supplier_id IS NULL
;

-- ----------------------------------------------------------------------------
-- BANK_FORMAT: IBAN required by country
-- ref: ref_workday_validate_supplier_bank_format (Bank Account Formatting sheet)
-- Note: country column is column due to merged header row in source workbook
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                                            table_name
     , 'BANK_FORMAT'                                                                  rule_type
     , 'bank_iban'                                                                    col
     , a.supplier_id                                                                  supplier_id
     , a.bank_country                                                                 detail
     , 'IBAN is required for ' || a.bank_country || ' accounts.'                      message
  FROM workday_supplier_settlement_account                                             a
  INNER JOIN
       ref_workday_validate_supplier_bank_format                                       r
          ON LOWER(TRIM(r.column))                                                 = LOWER(TRIM(a.bank_country))
 WHERE UPPER(TRIM(r.iban_display_option))                                              = 'REQUIRED'
   AND (a.bank_iban IS NULL OR TRIM(a.bank_iban) = '')
;

-- ----------------------------------------------------------------------------
-- BANK_FORMAT: IBAN length must match country requirement
-- ref: ref_workday_validate_supplier_bank_format (Bank Account Formatting sheet)
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                                              table_name
     , 'BANK_FORMAT'                                                                    rule_type
     , 'bank_iban'                                                                      col
     , a.supplier_id                                                                    supplier_id
     , a.bank_iban || ' (expected ' || r.iban_length || ' chars)'                       detail
     , 'IBAN for ' || a.bank_country || ' must be exactly ' || r.iban_length || ' characters.' message
  FROM workday_supplier_settlement_account                                               a
  INNER JOIN
       ref_workday_validate_supplier_bank_format                                         r
          ON LOWER(TRIM(r.column))                                                   = LOWER(TRIM(a.bank_country))
 WHERE TRY_CAST(r.iban_length AS INTEGER)                                                > 0
   AND a.bank_iban IS NOT NULL
   AND TRIM(a.bank_iban) != ''
   AND LENGTH(REPLACE(UPPER(TRIM(a.bank_iban)), ' ', ''))                               != TRY_CAST(r.iban_length AS INTEGER)
;

-- ----------------------------------------------------------------------------
-- BANK_FORMAT: BIC/SWIFT required by country
-- ref: ref_workday_validate_supplier_bank_format (Bank Account Formatting sheet)
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                                            table_name
     , 'BANK_FORMAT'                                                                  rule_type
     , 'bank_swift_code'                                                              col
     , a.supplier_id                                                                  supplier_id
     , a.bank_country                                                                 detail
     , 'BIC/SWIFT code is required for ' || a.bank_country || ' accounts.'            message
  FROM workday_supplier_settlement_account                                             a
  INNER JOIN
       ref_workday_validate_supplier_bank_format                                       r
          ON LOWER(TRIM(r.column))                                                 = LOWER(TRIM(a.bank_country))
 WHERE UPPER(TRIM(r.bic_display_option))                                               = 'REQUIRED'
   AND (a.bank_swift_code IS NULL OR TRIM(a.bank_swift_code) = '')
;

-- ----------------------------------------------------------------------------
-- BANK_FORMAT: Bank ID / routing number required by country
-- ref: ref_workday_validate_supplier_bank_format (Bank Account Formatting sheet)
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                                            table_name
     , 'BANK_FORMAT'                                                                  rule_type
     , 'bank_code_routing_number'                                                     col
     , a.supplier_id                                                                  supplier_id
     , a.bank_country || ' (' || r.bank_id_label || ')'                               detail
     , r.bank_id_label || ' is required for ' || a.bank_country || ' accounts.'       message
  FROM workday_supplier_settlement_account                                             a
  INNER JOIN
       ref_workday_validate_supplier_bank_format                                       r
          ON LOWER(TRIM(r.column))                                                 = LOWER(TRIM(a.bank_country))
 WHERE UPPER(TRIM(r.bank_id_display_option))                                           = 'REQUIRED'
   AND (a.bank_code_routing_number IS NULL OR TRIM(a.bank_code_routing_number) = '')
;

-- ----------------------------------------------------------------------------
-- BANK_FORMAT: Account number required by country
-- ref: ref_workday_validate_supplier_bank_format (Bank Account Formatting sheet)
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_settlement_account'                                            table_name
     , 'BANK_FORMAT'                                                                  rule_type
     , 'bank_account_number'                                                          col
     , a.supplier_id                                                                  supplier_id
     , a.bank_country                                                                 detail
     , 'Bank Account Number is required for ' || a.bank_country || ' accounts.'       message
  FROM workday_supplier_settlement_account                                             a
  INNER JOIN
       ref_workday_validate_supplier_bank_format                                       r
          ON LOWER(TRIM(r.column))                                                 = LOWER(TRIM(a.bank_country))
 WHERE UPPER(TRIM(r.account_number_display_option))                                    = 'REQUIRED'
   AND (a.bank_account_number IS NULL OR TRIM(a.bank_account_number) = '')
   AND (a.bank_iban IS NULL OR TRIM(a.bank_iban) = '')
;
