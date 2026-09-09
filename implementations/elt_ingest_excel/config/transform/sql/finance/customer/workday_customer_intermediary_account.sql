DROP TABLE IF EXISTS workday_customer_intermediary_account
;
CREATE TABLE workday_customer_intermediary_account
    AS
SELECT
       TRIM(c.customer_id) || '_BANK'                 settlement_bank_account_id
     , c.nrm_customer_name                            customer_name
     , c.customer_id                                  customer_id
     , TRIM(c.customer_id) || '_INTBANK'              intermediary_bank_account_id
     , c.nrm_country_name                             bank_country
     , c.nrm_currency_code                            currency
     , 'CHECKING'                                     bank_account_type
     , CAST(NULL AS VARCHAR)                          bank_name
     , CAST(NULL AS VARCHAR)                          name_on_account
     , CAST(NULL AS VARCHAR)                          bank_account_number
     , TRIM(c.short_name) || ' Intermediary'          bank_account_nickname
     , CAST(NULL AS VARCHAR)                          routing_transit_number
     , CAST(NULL AS VARCHAR)                          iban
     , CAST(NULL AS VARCHAR)                          swift_bank_identification_code
     , CAST(NULL AS VARCHAR)                          roll_number
     , CAST(NULL AS VARCHAR)                          check_digit
     , CAST(NULL AS VARCHAR)                          branch_id
     , CAST(NULL AS VARCHAR)                          branch_name
     , CASE
         WHEN UPPER(TRIM(c.inactive)) = 'NO'         
         THEN 'No'
         ELSE 'Yes'
       END                                            inactive
     , CAST(NULL AS VARCHAR)                          bank_instructions
  FROM src_fin_customer                c
 WHERE c.checkbook_id IS NOT NULL
   AND TRIM(c.checkbook_id) != ''
   AND 1 = 2
;
