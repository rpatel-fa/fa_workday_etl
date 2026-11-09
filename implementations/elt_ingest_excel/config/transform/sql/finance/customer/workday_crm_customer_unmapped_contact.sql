DROP TABLE IF EXISTS workday_crm_customer_unmapped_contact
;
CREATE TABLE workday_crm_customer_unmapped_contact
    AS
SELECT cc.*
  FROM crm_contact                     cc
 WHERE NULLIF(NULLIF(TRIM(cc.licenceref_no), ''), 'NULL') IS NOT NULL
   AND cc.fan NOT IN (
     SELECT DISTINCT crm_customer_id
       FROM workday_crm_customer_mapped
      WHERE crm_customer_type = 'INDIVIDUAL'
     )
;