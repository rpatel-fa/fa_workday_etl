-- ============================================================================
-- VALIDATE: Customer Payment
-- Source table: workday_customer_payment
-- VBA rules: Match (cross-tab), Missing, Match (payment values)
-- ============================================================================

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Customer ID must exist in Customer Name tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_payment'                                      table_name
     , 'CROSS_TAB'                                                   rule_type
     , 'customer_id'                                                 col
     , p.customer_id                                                 customer_id
     , p.customer_id                                                 detail
     , 'Customer ID must be listed first on the Customer Name tab.'  message
  FROM workday_customer_payment                                      p
  LEFT OUTER JOIN
       workday_customer_name                                         n
         ON  n.customer_id                                           = p.customer_id
 WHERE n.customer_id IS NULL
;

-- ----------------------------------------------------------------------------
-- MISSING: Default Payment Type is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_payment'                                    table_name
     , 'MISSING'                                                    rule_type
     , 'default_payment_type'                                       col
     , p.customer_id                                                customer_id
     , NULL                                                         detail
     , 'Default Payment Type is a required value.'                  message
  FROM workday_customer_payment                                     p
 WHERE p.default_payment_type IS NULL
    OR TRIM(p.default_payment_type) = ''
;

-- ----------------------------------------------------------------------------
-- MATCH: Default Payment Type must be a valid Payment_Type value
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_payment'                                             table_name
     , 'MATCH'                                                               rule_type
     , 'default_payment_type'                                                col
     , p.customer_id                                                         customer_id
     , p.default_payment_type                                                detail
     , 'Not a valid payment type. Please select a payment type from the drop-down list.' message
  FROM workday_customer_payment                                              p
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
-- MATCH: Payment Terms must be a valid value (when populated)
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_payment'                                             table_name
     , 'MATCH'                                                               rule_type
     , 'payment_terms'                                                       col
     , p.customer_id                                                         customer_id
     , p.payment_terms                                                       detail
     , 'Not a valid payment term. Please select a payment term from the drop-down list.' message
  FROM workday_customer_payment                                              p
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
