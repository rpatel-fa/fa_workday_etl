DROP TABLE IF EXISTS workday_crm_customer_unmapped_account
;
CREATE TABLE workday_crm_customer_unmapped_account
    AS
SELECT ca.* 
  FROM crm_account                     ca
 WHERE NULLIF(NULLIF(TRIM(ca.licenceref_no), ''), 'NULL') IS NOT NULL
   AND ca.account_number_id NOT IN (
     SELECT DISTINCT crm_customer_id
       FROM workday_crm_customer_mapped
      WHERE crm_customer_type = 'ACCOUNT'
     )
;