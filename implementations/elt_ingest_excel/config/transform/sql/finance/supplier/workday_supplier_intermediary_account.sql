DROP TABLE IF EXISTS workday_supplier_intermediary_account
;
CREATE TABLE workday_supplier_intermediary_account 
    AS
SELECT
       s.supplier_id || '_BANK'                                   settlement_bank_account_id
     , s.nrm_supplier_name                                        supplier_name
     , s.supplier_id                                              supplier_id
     , s.supplier_id || '_INTBANK'                                intermediary_bank_account_id
     , CAST(NULL AS VARCHAR)                                      bank_country
     , s.nrm_currency_code                                        currency
     , 'CHECKING'                                                 bank_account_type
     , CAST(NULL AS VARCHAR)                                      bank_name
     , CAST(NULL AS VARCHAR)                                      name_on_account
     , CAST(NULL AS VARCHAR)                                      bank_account_number
     , TRIM(s.vendor_short_name) || ' Intermediary'               bank_account_nickname
     , TRIM(s.central_bank_code)                                  routing_number_bank_code
     , CAST(NULL AS VARCHAR)                                      iban
     , CAST(NULL AS VARCHAR)                                      swift_bank_identification_code
     , CAST(NULL AS VARCHAR)                                      roll_number
     , CAST(NULL AS VARCHAR)                                      check_digit
     , CAST(NULL AS VARCHAR)                                      branch_id
     , CAST(NULL AS VARCHAR)                                      branch_name
     , CASE
         WHEN UPPER(TRIM(s.vendor_status)) = 'ACTIVE'
         THEN 'No'
         ELSE 'Yes'
       END                                                        inactive
     , CAST(NULL AS VARCHAR)                                      bank_instructions
  FROM src_fin_supplier s
 WHERE s.central_bank_code IS NOT NULL
   AND TRIM(s.central_bank_code) != ''
   AND 1 = 2
;
