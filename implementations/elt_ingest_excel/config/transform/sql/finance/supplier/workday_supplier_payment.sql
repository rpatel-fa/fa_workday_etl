DROP TABLE IF EXISTS workday_supplier_payment
;
CREATE TABLE workday_supplier_payment 
    AS
  WITH cte_supplier_paymment
    AS (
SELECT
       t.supplier_id                                     supplier_id
     , t.nrm_supplier_name                               supplier_name
     , t.nrm_payment_terms_id                            payment_terms
     , CASE
         WHEN t.nrm_currency_code         = 'GBP'
          AND t.nrm_bank_name             IS NOT NULL
          AND t.nrm_bank_sort_code        IS NOT NULL
          AND t.nrm_bank_account_number   IS NOT NULL
         THEN 'BACS|Manual'
         WHEN t.nrm_currency_code         = 'GBP'
          AND t.nrm_bank_name             IS NOT NULL
          AND t.nrm_bank_iban             IS NOT NULL          
         THEN 'Wire|Manual'
         WHEN t.nrm_currency_code         = 'EUR'
          AND t.nrm_bank_name             IS NOT NULL
          AND t.nrm_bank_iban             IS NOT NULL
         THEN 'SEPA|Wire|Manual'
         WHEN t.nrm_currency_code         = 'EUR'
          AND t.nrm_bank_name             IS NOT NULL
          AND t.nrm_bank_account_number   IS NOT NULL
          AND t.nrm_bank_swift_code       IS NOT NULL
         THEN 'Wire|Manual'
         WHEN t.nrm_currency_code         = 'EUR'
          AND t.nrm_bank_name             IS NOT NULL
          AND t.nrm_bank_sort_code        IS NOT NULL
          AND t.nrm_bank_account_number   IS NOT NULL
         THEN 'Wire|Manual'
         WHEN t.nrm_currency_code         IN ('USD', 'CHF','AUD','PLN','SEK','AED','QAR','DKK','NOK')
          AND t.nrm_bank_name             IS NOT NULL
          AND t.nrm_bank_iban             IS NOT NULL
         THEN 'Wire|Manual'
         WHEN t.nrm_currency_code         IN ('USD', 'CHF','AUD','PLN','SEK','AED','QAR','DKK','NOK')
          AND t.nrm_bank_name             IS NOT NULL
          AND t.nrm_bank_account_number   IS NOT NULL
          AND t.nrm_bank_swift_code       IS NOT NULL
         THEN 'Wire|Manual'
         WHEN t.nrm_currency_code         IN ('USD', 'CHF','AUD','PLN','SEK','AED','QAR','DKK','NOK')
          AND t.nrm_bank_name             IS NOT NULL
          AND t.nrm_bank_sort_code        IS NOT NULL
          AND t.nrm_bank_account_number   IS NOT NULL
         THEN 'Wire|Manual'
         ELSE 'Manual'
       END                                               payment_types_accepted
     , CAST(NULL AS VARCHAR)                             shipping_terms
     , CAST(NULL AS VARCHAR)                             always_separate_payments
     , CAST(NULL AS VARCHAR)                             do_not_pay_during_bank_account_updates
     , CAST(NULL AS VARCHAR)                             exclude_freight_amount_from_supplier_invoice_discount
     , CAST(NULL AS VARCHAR)                             exclude_other_charge_from_supplier_invoice_discount
     , CAST(NULL AS VARCHAR)                             exclude_tax_amount_from_supplier_invoice_discount
  FROM src_fin_supplier                t
 WHERE t.nrm_payment_terms_id          IS NOT NULL
       )
SELECT t.supplier_id
     , t.supplier_name
     , t.payment_terms
     , t.payment_types_accepted
     , SPLIT_PART(t.payment_types_accepted, '|', 1)      default_payment_type
     , t.shipping_terms
     , t.always_separate_payments
     , t.do_not_pay_during_bank_account_updates
     , t.exclude_freight_amount_from_supplier_invoice_discount
     , t.exclude_other_charge_from_supplier_invoice_discount
     , t.exclude_tax_amount_from_supplier_invoice_discount
  FROM cte_supplier_paymment           t
 ORDER BY
       supplier_id
;
