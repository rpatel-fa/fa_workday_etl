DROP TABLE IF EXISTS workday_customer_credit_card
;
CREATE TABLE workday_customer_credit_card
    AS
SELECT
       c.customer_id                                  customer_id
     , c.nrm_customer_name                            customer_name
     , CAST(NULL AS VARCHAR)                          merchant_customer_profile_id
     , CAST(NULL AS VARCHAR)                          primary_card
     , CAST(NULL AS VARCHAR)                          merchant_account
     , CAST(NULL AS VARCHAR)                          credit_card_type
     , CAST(NULL AS VARCHAR)                          first_name
     , CAST(NULL AS VARCHAR)                          last_name
     , CAST(NULL AS VARCHAR)                          expiration_date
     , CAST(NULL AS VARCHAR)                          last_4_digits_of_credit_card_number
     , CAST(NULL AS VARCHAR)                          customer_profile_id
  FROM src_fin_customer                c
 WHERE 1 = 2
;
