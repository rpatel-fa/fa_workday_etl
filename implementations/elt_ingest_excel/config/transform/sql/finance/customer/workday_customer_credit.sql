DROP TABLE IF EXISTS workday_customer_credit
;
CREATE TABLE workday_customer_credit
    AS
SELECT
       c.customer_id                                  customer_id
     , c.nrm_customer_name                            customer_name
     , CAST(NULL AS VARCHAR)                          duns_number
     , CAST(NULL AS VARCHAR)                          exempt_from_dunning
     , TRIM(c.tax_schedule_id)                        tax_code
     , c.nrm_currency_code                            credit_limit_currency
     , CASE
         WHEN UPPER(TRIM(c.credit_limit_type)) = 'UNLIMITED' 
         THEN NULL
         ELSE c.credit_limit_amount
       END                                            credit_limit
     , CAST(NULL AS VARCHAR)                          hierarchy_credit_limit
     , CAST(NULL AS VARCHAR)                          credit_verification_date
     , CAST(NULL AS VARCHAR)                          commercial_credit_score
     , CAST(NULL AS VARCHAR)                          commercial_credit_score_date
     , CAST(NULL AS VARCHAR)                          commercial_credit_score_note
     , CAST(NULL AS VARCHAR)                          composite_risk_score
     , CAST(NULL AS VARCHAR)                          composite_risk_date
     , CAST(NULL AS VARCHAR)                          composite_risk_note
     , CAST(NULL AS VARCHAR)                          customer_satisfaction_score
     , CAST(NULL AS VARCHAR)                          customer_satisfaction_date
     , CAST(NULL AS VARCHAR)                          customer_satisfaction_note
  FROM src_fin_customer                c
 WHERE 1 = 2
;
