DROP TABLE IF EXISTS workday_customer_bank_account_address
;
CREATE TABLE workday_customer_bank_account_address
    AS
SELECT
       c.nrm_customer_name                            customer_name
     , c.customer_id                                  customer_id
     , TRIM(c.customer_id) || '_BANK'                 settlement_bank_account_id
     , CAST(NULL AS VARCHAR)                          intermediary_bank_account_id
     , TRIM(c.customer_id) || '_BANK_ADDR'            address_id
     , c.nrm_country_name                             country
     , TRIM(c.county)                                 region
     , CAST(NULL AS VARCHAR)                          subregion
     , TRIM(c.city)                                   city
     , CAST(NULL AS VARCHAR)                          submunicipality
     , TRIM(c.address_1)                              address_line_1
     , NULLIF(TRIM(c.address_2), '')                  address_line_2
     , NULLIF(TRIM(c.address_3), '')                  address_line_3
     , CAST(NULL AS VARCHAR)                          address_line_4
     , TRIM(UPPER(c.post_code))                       postal_code
     , 'Yes'                                          is_public
     , 'Yes'                                          is_primary
     , c.created_date                                 effective_date
     , TRIM(c.address_code)                           address_type
     , TRIM(c.address_code)                           use_for
     , CAST(NULL AS VARCHAR)                          use_for_tenanted
  FROM src_fin_customer                c
 WHERE 1 = 2
;
