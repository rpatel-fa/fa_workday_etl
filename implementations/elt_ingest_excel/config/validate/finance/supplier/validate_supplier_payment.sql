-- ============================================================================
-- VALIDATE: Supplier Payment
-- Source table: workday_supplier_payment
-- VBA rules: Match (cross-tab), Missing, Match (payment values),
--            RequiredBetweenTabs (EFT/Wire → settlement account)
-- ============================================================================

-- Valid payment types from workbook named range Payment_Type
-- Pipe-delimited combos are split for per-token validation below.

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Supplier ID must exist in Supplier Name tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_payment'                                      table_name
     , 'CROSS_TAB'                                                   rule_type
     , 'supplier_id'                                                 col
     , p.supplier_id                                                 supplier_id
     , p.supplier_id                                                 detail
     , 'Supplier ID must be listed first on the Supplier Name tab.'  message
  FROM workday_supplier_payment                                      p
  LEFT OUTER JOIN
       workday_supplier_name                                         n
         ON  n.supplier_id                                           = p.supplier_id
 WHERE n.supplier_id IS NULL
;

-- ----------------------------------------------------------------------------
-- MISSING: Default Payment Type is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_payment'                                    table_name
     , 'MISSING'                                                    rule_type
     , 'default_payment_type'                                       col
     , p.supplier_id                                                supplier_id
     , NULL                                                         detail
     , 'Default Payment Type is a required value.'                  message
  FROM workday_supplier_payment                                     p
 WHERE p.default_payment_type IS NULL
    OR TRIM(p.default_payment_type) = ''
;

-- ----------------------------------------------------------------------------
-- MISSING: Payment Types Accepted is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_payment'                                    table_name
     , 'MISSING'                                                    rule_type
     , 'payment_types_accepted'                                     col
     , p.supplier_id                                                supplier_id
     , NULL                                                         detail
     , 'Payment Types Accepted is a required field.'                message
  FROM workday_supplier_payment                                     p
 WHERE p.payment_types_accepted IS NULL
    OR TRIM(p.payment_types_accepted) = ''
;

-- ----------------------------------------------------------------------------
-- MATCH: Default Payment Type must be a valid Payment_Type value
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_payment'                                             table_name
     , 'MATCH'                                                               rule_type
     , 'default_payment_type'                                                col
     , p.supplier_id                                                         supplier_id
     , p.default_payment_type                                                detail
     , 'Not a valid payment type. Please select a payment type from the drop-down list.' message
  FROM workday_supplier_payment                                              p
 WHERE p.default_payment_type IS NOT NULL
   AND TRIM(p.default_payment_type) != ''
   AND p.default_payment_type NOT IN
       (
         'Credit Card'
       , 'Debit Card'
       , 'Direct Deposit'
       , 'EFT'
       , 'Wire'
       , 'Manual'
       , 'Direct Debit'
       , 'BACS'
       , 'SEPA'
       , 'SEPA-Payroll'
       , 'BACS-Payroll'
       , 'SEPA-Urgent'
       , 'SWIFT'
       , 'Faster Payments'
       , 'CHAPS'
       , 'EFT with Reference'
       )
;

-- ----------------------------------------------------------------------------
-- MATCH: Each token in Payment Types Accepted must be a valid Payment_Type
-- Payment types are pipe-delimited (e.g. 'BACS|EFT|Manual')
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_payment'                                             table_name
     , 'MATCH'                                                               rule_type
     , 'payment_types_accepted'                                              col
     , p.supplier_id                                                         supplier_id
     , tok.token                                                             detail
     , 'Not a valid payment type in Payment Types Accepted.'                 message
  FROM workday_supplier_payment                                              p
       , UNNEST(STRING_SPLIT(p.payment_types_accepted, '|'))                 tok(token)
 WHERE p.payment_types_accepted IS NOT NULL
   AND TRIM(tok.token) NOT IN
       (
         'Credit Card'
       , 'Debit Card'
       , 'Direct Deposit'
       , 'EFT'
       , 'Wire'
       , 'Manual'
       , 'Direct Debit'
       , 'BACS'
       , 'SEPA'
       , 'SEPA-Payroll'
       , 'BACS-Payroll'
       , 'SEPA-Urgent'
       , 'SWIFT'
       , 'Faster Payments'
       , 'CHAPS'
       , 'EFT with Reference'
       )
;

-- ----------------------------------------------------------------------------
-- MATCH: Payment Terms must be a valid value (when populated)
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_payment'                                             table_name
     , 'MATCH'                                                               rule_type
     , 'payment_terms'                                                       col
     , p.supplier_id                                                         supplier_id
     , p.payment_terms                                                       detail
     , 'Not a valid payment term. Please select a payment term from the drop-down list.' message
  FROM workday_supplier_payment                                              p
  LEFT JOIN
       (
         SELECT DISTINCT
                r.workday_payment_terms                                      workday_payment_terms
           FROM ref_source_payment_terms                                     r
       )                                                                     r
    ON UPPER(TRIM(r.workday_payment_terms))                                  = UPPER(TRIM(p.payment_terms))
 WHERE p.payment_terms IS NOT NULL
   AND TRIM(p.payment_terms) != ''
   AND r.workday_payment_terms IS NULL
;

-- ----------------------------------------------------------------------------
-- CROSS_TAB: If EFT, Wire, or Direct Deposit used, settlement account required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_payment'                                                       table_name
     , 'CROSS_TAB'                                                                    rule_type
     , 'payment_types_accepted'                                                        col
     , p.supplier_id                                                                   supplier_id
     , p.payment_types_accepted                                                        detail
     , 'If Wire, Direct Deposit or EFT is used for payment type, banking information is required on the Settlement Account tab.' message
  FROM workday_supplier_payment                                                        p
  LEFT OUTER JOIN
       workday_supplier_settlement_account                                             a
         ON  a.supplier_id                                                             = p.supplier_id
 WHERE (
           p.payment_types_accepted ILIKE '%EFT%'
        OR p.payment_types_accepted ILIKE '%Wire%'
        OR p.payment_types_accepted ILIKE '%Direct Deposit%'
       )
   AND a.supplier_id IS NULL
;
