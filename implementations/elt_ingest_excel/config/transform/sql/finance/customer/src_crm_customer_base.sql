DROP TABLE IF EXISTS src_crm_customer_base
;
CREATE TABLE src_crm_customer_base
    AS
WITH cte_ca
  AS (
SELECT 'ACCOUNT'                       crm_customer_type
     , ca.account_number_id            crm_customer_id
     , ca.licenceref_no                crm_licenceref_no 
  FROM crm_account                     ca 
 WHERE NULLIF(NULLIF(TRIM(ca.licenceref_no), ''), 'NULL') IS NOT NULL
     )
     , cte_cc
  AS (
SELECT 'INDIVIDUAL'                    crm_customer_type
     , cc.fan                          crm_customer_id
     , cc.licenceref_no                crm_licenceref_no
  FROM crm_contact                     cc 
 WHERE NULLIF(NULLIF(TRIM(cc.licenceref_no), ''), 'NULL') IS NOT NULL
     )
     , cte_crm
  AS (
SELECT * FROM cte_ca
UNION ALL
SELECT * FROM cte_cc
     )
SELECT *
     , TRIM(UNNEST(STRING_SPLIT(crm_licenceref_no, '|')))  crm_unnest_reference_id
  FROM cte_crm
;
