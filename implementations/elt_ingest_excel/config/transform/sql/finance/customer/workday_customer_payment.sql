DROP TABLE IF EXISTS workday_customer_payment
;
CREATE TABLE workday_customer_payment
    AS
SELECT
       c.customer_id                                     customer_id
     , c.nrm_customer_name                               customer_name
     , m.workday_payment_terms                           payment_terms
     , CASE
         WHEN c.nrm_currency_code = 'GBP'
         THEN 'Manual'
         WHEN c.nrm_currency_code = 'EUR'
         THEN 'Manual'
         WHEN c.nrm_currency_code IN ('USD','CHF','AUD','SEK','AED','QAR','DKK','NOK')
         THEN 'Manual'
         ELSE 'Manual'
       END                                               default_payment_type
     , CAST(NULL AS VARCHAR)                             interest_rule
     , CAST(NULL AS VARCHAR)                             late_fee_rule
  FROM src_fin_customer                                  c
       LEFT OUTER JOIN
       ref_source_payment_terms                          m
         ON  UPPER(m.source_payment_terms)               = c.nrm_payment_terms_id
 ORDER BY
       customer_id
;
