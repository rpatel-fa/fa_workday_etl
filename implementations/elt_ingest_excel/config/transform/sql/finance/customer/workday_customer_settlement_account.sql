DROP TABLE IF EXISTS workday_customer_settlement_account
;
CREATE TABLE workday_customer_settlement_account
    AS
SELECT
       c.customer_id                                                       customer_id
     , c.nrm_customer_name                                                 customer_name
     , TRIM(c.customer_id) || '_BANK'                                      settlement_bank_account_id
     , c.nrm_country_name                                                  country
     , c.nrm_currency_code                                                 currency
     , COALESCE(NULLIF(TRIM(c.statement_name), ''), c.nrm_customer_name)   bank_account_nickname
     , 'Checking'                                                          bank_account_type
     , CAST(NULL AS VARCHAR)                                               bank_name
     , CAST(NULL AS VARCHAR)                                               routing_transit_number
     , CAST(NULL AS VARCHAR)                                               branch_id
     , CAST(NULL AS VARCHAR)                                               branch_name
     , CAST(NULL AS VARCHAR)                                               bank_account_number
     , CAST(NULL AS VARCHAR)                                               check_digit
     , COALESCE(NULLIF(TRIM(c.statement_name), ''), c.nrm_customer_name)   name_on_account
     , CAST(NULL AS VARCHAR)                                               roll_number
     , CAST(NULL AS VARCHAR)                                               iban
     , CAST(NULL AS VARCHAR)                                               swift_bank_identification_code
     , CASE
         WHEN c.nrm_currency_code = 'GBP'
         THEN 'BACS|EFT|Manual'
         WHEN c.nrm_currency_code = 'EUR'
         THEN 'SEPA|Wire|Manual'
         WHEN c.nrm_currency_code = 'USD'
         THEN 'Wire|EFT|Manual'
         WHEN c.nrm_currency_code IN ('CHF','AUD','PLN','SEK','AED','QAR','DKK','NOK')
         THEN 'Wire|Manual'
         ELSE 'BACS|EFT|Manual'
       END                                                                 payment_types_accepted
     , CASE
         WHEN c.nrm_currency_code = 'GBP'
         THEN 'EFT'
         WHEN c.nrm_currency_code = 'EUR'
         THEN 'SEPA'
         WHEN c.nrm_currency_code IN ('USD','CHF','AUD','SEK','AED','QAR','DKK','NOK')
         THEN 'Wire'
         ELSE 'BACS'
       END                                                                 payment_types
     , CAST(NULL AS VARCHAR)                                               for_supplier_connections_only
     , CAST(NULL AS VARCHAR)                                               requires_prenote
     , CAST(NULL AS VARCHAR)                                               payment_type_prenote
     , CAST(NULL AS VARCHAR)                                               inactive
     , CAST(NULL AS VARCHAR)                                               bank_instructions
  FROM src_fin_customer                c
 WHERE c.checkbook_id IS NOT NULL
   AND TRIM(c.checkbook_id) != ''
   AND 1 = 2
;
